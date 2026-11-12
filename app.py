# -*- coding: utf-8 -*-
"""保格式翻译 · 本地面板服务（含回译体检）"""
import base64
import subprocess
import difflib
import io
import json
import os
import re
import sys
import tempfile
import threading
import time
import traceback
import uuid
import urllib.parse
import webbrowser
from http.server import ThreadingHTTPServer, BaseHTTPRequestHandler

ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(ROOT, "脚本"))
from translate_xlsx import translate_file, KEY_FILE
from quality_check import check_quality
import update as upd
from suggest_terms import suggest as suggest_terms, add_terms, load_terms_file
import translate_xlsx
from translate_docx import translate_docx, docx_extract, docx_apply

import zipfile
from lxml import etree

NS = {'m': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
M = NS['m']

# 打包成 exe 后的路径处理
FROZEN = getattr(sys, "frozen", False)
EXE_DIR = os.path.dirname(sys.executable) if FROZEN else ROOT
RES_DIR = getattr(sys, "_MEIPASS", ROOT)
CONFIG_PATH = os.path.join(EXE_DIR, "config.json")

PORT = 8770
JOBS = {}
CHECKS = {}
SUGG = {}
CONV = {}
LOCK = threading.Lock()


def current_provider():
    return (load_config().get("provider") or "deepseek").strip()


def load_key(provider=None):
    """按厂商取 key；取不到再退回旧的单 key 写法。"""
    cfg = load_config()
    p = (provider or cfg.get("provider") or "deepseek").strip()
    keys = cfg.get("keys") or {}
    k = (keys.get(p) or "").strip()
    if k:
        return k
    if p == "deepseek":
        k = (cfg.get("api_key") or "").strip()
        if k:
            return k
    cands = [os.path.join(EXE_DIR, "chat-key.txt")]
    if not FROZEN:          # 开发机上才回落到工作目录的 key，打包后绝不读外部路径
        cands.append(KEY_FILE)
    for p in cands:
        if os.path.exists(p):
            try:
                k = open(p, encoding="utf-8").read().strip()
                if k:
                    return k
            except Exception:
                pass
    return ""


def load_config():
    if os.path.exists(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {}


def save_config(patch):
    data = load_config()
    data.update(patch)
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=1)


def save_key(key):
    data = {}
    if os.path.exists(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, encoding="utf-8") as f:
                data = json.load(f)
        except Exception:
            data = {}
    data["api_key"] = key.strip()
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=1)


def read_strings(path):
    with zipfile.ZipFile(path) as z:
        root = etree.fromstring(z.read("xl/sharedStrings.xml"))
    out = []
    for si in root.findall('m:si', NS):
        out.append(''.join((t.text or '') for t in si.findall('.//m:t', NS)))
    return out


def bytes_to_strings(data):
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        root = etree.fromstring(z.read("xl/sharedStrings.xml"))
    return [''.join((t.text or '') for t in si.findall('.//m:t', NS))
            for si in root.findall('m:si', NS)]


def _dump(root):
    raw = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
    return re.sub(rb'&#(\d+);', lambda m: chr(int(m.group(1))).encode('utf-8'), raw)


def rebuild_xlsx(base_bytes, edits):
    """以翻译结果为基底，把用户改过的条目写回去。"""
    with zipfile.ZipFile(io.BytesIO(base_bytes)) as zin:
        names = zin.namelist()
        parts = {n: zin.read(n) for n in names}
    if not edits:
        return base_bytes
    ss = etree.fromstring(parts['xl/sharedStrings.xml'])
    sis = ss.findall('m:si', NS)
    for k, text in edits.items():
        i = int(k)
        if i < 0 or i >= len(sis):
            continue
        si = sis[i]
        for ch in list(si):
            si.remove(ch)
        t = etree.SubElement(si, '{%s}t' % M)
        t.text = text
        t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    parts['xl/sharedStrings.xml'] = _dump(ss)
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zout:
        for n in names:
            zi = zipfile.ZipInfo(n, date_time=(2026, 1, 1, 0, 0, 0))
            zi.compress_type = zipfile.ZIP_DEFLATED
            zout.writestr(zi, parts[n])
    return buf.getvalue()


def worker(job_id, in_path, out_path, direction, bilingual=False, kind="xlsx", mode="cloud", ollama_url="", provider="deepseek", base_url="", cloud_model=""):
    job = JOBS[job_id]
    job["state"] = "running"

    def cb(done, total):
        job["done"] = done
        job["total"] = total

    try:
        if kind == "docx":
            translate_docx(in_path, out_path, load_key(), direction, on_progress=cb)
        else:
            translate_file(in_path, out_path, load_key(provider), direction, on_progress=cb,
                           bilingual=bilingual, mode=mode, ollama_url=ollama_url,
                           provider=provider, base_url=base_url, cloud_model=cloud_model)
        with open(out_path, "rb") as f:
            job["data"] = base64.b64encode(f.read()).decode("ascii")
        job["state"] = "done"
        job["done"] = job["total"] or 1
        # 翻完自动在后台找新词（不阻塞界面，不额外花钱——本地模型干的）
        try:
            threading.Thread(target=auto_suggest, args=(job_id,), daemon=True).start()
        except Exception:
            pass
    except Exception as ex:
        job["state"] = "error"
        job["error"] = str(ex)
        traceback.print_exc()
    finally:
        for p in (in_path, out_path):
            try:
                os.remove(p)
            except OSError:
                pass


def check_worker(cid, src_b64, out_b64, direction, job_mode="cloud", kind="xlsx"):
    chk = CHECKS[cid]
    chk["state"] = "running"
    rev = "zh2en" if direction == "en2zh" else "en2zh"
    tmpdir = tempfile.mkdtemp(prefix="chk_")
    p_src = os.path.join(tmpdir, "src.xlsx")
    p_mid = os.path.join(tmpdir, "mid.xlsx")
    p_back = os.path.join(tmpdir, "back.xlsx")
    try:
        with open(p_src, "wb") as f:
            f.write(base64.b64decode(src_b64))
        with open(p_mid, "wb") as f:
            f.write(base64.b64decode(out_b64))

        def cb(done, total):
            chk["done"] = done
            chk["total"] = total

        if kind == "docx":
            translate_docx(p_mid, p_back, load_key(), rev, on_progress=cb)
        else:
            translate_file(p_mid, p_back, load_key(), rev, on_progress=cb,
                           mode=job_mode)

        a = read_strings(p_src)
        b = read_strings(p_back)
        items = []
        for i in range(min(len(a), len(b))):
            x, y = a[i].strip(), b[i].strip()
            if not x or not y:
                continue
            if x == y:
                continue
            ratio = difflib.SequenceMatcher(None, x.lower(), y.lower()).ratio()
            if ratio < 0.62:
                items.append({"i": i, "orig": x, "back": y, "ratio": round(ratio, 2)})
        items.sort(key=lambda d: d["ratio"])
        chk["items"] = items[:60]
        chk["count"] = len(items)
        chk["state"] = "done"
    except Exception as ex:
        chk["state"] = "error"
        chk["error"] = str(ex)
        traceback.print_exc()
    finally:
        for p in (p_src, p_mid, p_back):
            try:
                os.remove(p)
            except OSError:
                pass


def extract_texts(data, kind):
    if kind == "docx":
        return docx_extract(data)
    return bytes_to_strings(data)


def apply_edits(data, kind, edits):
    if kind == "docx":
        return docx_apply(data, edits)
    return rebuild_xlsx(data, edits)


def suggest_worker(sid, paths):
    st = SUGG[sid]
    st["state"] = "running"
    try:
        items = suggest_terms(paths, load_key())
        st["items"] = items
        st["state"] = "done"
    except Exception as ex:
        st["state"] = "error"
        st["error"] = str(ex)
        traceback.print_exc()


def filee_cli():
    """找 Filee 的命令行；找不到返回空字符串。"""
    cands = [
        os.path.join(EXE_DIR, "工具", "filee", "filee-cli.exe"),
        os.path.join(EXE_DIR, "filee", "filee-cli.exe"),
        os.path.join(os.path.dirname(EXE_DIR), "filee", "filee-cli.exe"),
    ]
    for c in cands:
        if os.path.exists(c):
            return c
    return ""


CONVERT_FORMATS = ["pdf", "docx", "xlsx", "pptx", "png", "jpg", "webp", "txt", "csv", "html", "epub", "md"]


def run_convert(in_path, out_dir, to_fmt):
    cli = filee_cli()
    if not cli:
        raise RuntimeError("没找到文件转换引擎（Filee）。请让管理员安装转换组件。")
    cmd = [cli, "convert", in_path, "--to", to_fmt, "-o", out_dir]
    p = subprocess.run(cmd, capture_output=True, text=True, timeout=300,
                       encoding="utf-8", errors="replace")
    out = (p.stdout or "") + (p.stderr or "")
    files = [f for f in os.listdir(out_dir) if not f.startswith(".")]
    if not files:
        raise RuntimeError("转换没有产出文件：\n" + out[-400:])
    return os.path.join(out_dir, files[0]), out.strip()


def auto_suggest(job_id):
    """翻译完成后自动扫一遍：文件里有多少词是术语表没有的。"""
    job = JOBS.get(job_id)
    if not job:
        return
    try:
        kind = job.get("kind", "xlsx")
        data = base64.b64decode(job["src"])
        tmpdir = tempfile.mkdtemp(prefix="auto_")
        pth = os.path.join(tmpdir, "src" + (".docx" if kind == "docx" else ".xlsx"))
        with open(pth, "wb") as f:
            f.write(data)
        items = suggest_terms([pth], load_key())
        job["suggestions"] = items
        job["suggest_done"] = True
        print(f"[自动找词] {job_id} 发现 {len(items)} 个候选")
    except Exception as ex:
        job["suggest_done"] = True
        job["suggest_error"] = str(ex)[:200]
        print("[自动找词] 失败:", ex)


def _restart_self():
    """更新完重启自己：拉起新进程，退出当前进程。"""
    try:
        if getattr(sys, "frozen", False):
            subprocess.Popen([sys.executable], close_fds=True)
        else:
            subprocess.Popen([sys.executable, os.path.abspath(__file__)], close_fds=True)
    except Exception as ex:
        print("重启失败:", ex)
        return
    time.sleep(0.5)
    os._exit(0)


def retranslate_batch(job, indices):
    from translate_xlsx import call_api, protect, restore, clean_artifacts
    orig = extract_texts(base64.b64decode(job["src"]), job.get("kind", "xlsx"))
    items, spans_map = [], {}
    for i in indices:
        if i >= len(orig):
            continue
        p_, spans = protect(orig[i])
        spans_map[i] = spans
        items.append(("k%d" % i, "", p_))
    results = {}
    B = 6
    for start in range(0, len(items), B):
        got = call_api(load_key(), items[start:start + B], job.get("direction", "en2zh"),
                       hint="这些条目上一轮可能漏译或术语不对，请完整重译，术语严格按表，数字一个不许少。")
        results.update(got)
    fixed = 0
    for i in spans_map:
        key = "k%d" % i
        if key in results and str(results[key]).strip():
            try:
                job["edits"][str(i)] = clean_artifacts(restore(results[key], spans_map[i]))
                fixed += 1
            except Exception:
                pass
    return fixed


def retranslate_one(orig, direction, hint):
    from translate_xlsx import call_api, protect, restore, clean_artifacts
    p, spans = protect(orig)
    got = call_api(load_key(), [("k0", "", p)], direction, hint=hint)
    text = got.get("k0") or (next(iter(got.values())) if got else "")
    if not text:
        raise RuntimeError("模型没有返回结果")
    return clean_artifacts(restore(text, spans))


class Handler(BaseHTTPRequestHandler):
    def log_message(self, fmt, *args):
        pass

    def _json(self, obj, code=200):
        body = json.dumps(obj, ensure_ascii=False).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    MIME = {
        ".pdf": "application/pdf",
        ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        ".pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        ".png": "image/png", ".jpg": "image/jpeg", ".jpeg": "image/jpeg",
        ".webp": "image/webp", ".gif": "image/gif", ".txt": "text/plain; charset=utf-8",
        ".md": "text/plain; charset=utf-8", ".csv": "text/csv; charset=utf-8",
        ".html": "text/html; charset=utf-8", ".epub": "application/epub+zip",
        ".zip": "application/zip",
    }

    def _file(self, name, data):
        self.send_response(200)
        ext = os.path.splitext(name)[1].lower()
        ctype = self.MIME.get(ext, "application/octet-stream")
        self.send_header("Content-Type", ctype)
        q = urllib.parse.quote(name)
        self.send_header("Content-Disposition", f"attachment; filename*=UTF-8''{q}")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def do_GET(self):
        path = urllib.parse.urlparse(self.path).path
        if path in ("/", "/index.html"):
            p = os.path.join(RES_DIR, "panel", "index.html")
            with open(p, "rb") as f:
                body = f.read()
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)
            return
        if path.startswith("/api/job/"):
            job = JOBS.get(path.rsplit("/", 1)[-1])
            if not job:
                self._json({"error": "no such job"}, 404)
                return
            self._json({"state": job["state"], "done": job.get("done", 0),
                        "total": job.get("total", 0), "error": job.get("error", ""),
                        "suggest_count": len(job.get("suggestions") or []),
                        "suggest_done": bool(job.get("suggest_done")),
                        "suggest_error": job.get("suggest_error", "")})
            return
        if path.startswith("/api/file/"):
            job = JOBS.get(path.rsplit("/", 1)[-1])
            if not job or job["state"] != "done":
                self._json({"error": "not ready"}, 404)
                return
            raw = base64.b64decode(job["data"])
            data = apply_edits(raw, job.get("kind", "xlsx"), job.get("edits", {}))
            self._file(job["out_name"], data)
            return
        if path.startswith("/api/convert-file/"):
            cid = path.rsplit("/", 1)[-1]
            it = CONV.get(cid)
            if not it:
                self._json({"error": "文件已失效，请重新转换"}, 404)
                return
            self._file(it["name"], it["data"])
            return
        if path == "/api/version":
            self._json({"version": upd.VERSION})
            return
        if path == "/api/update-check":
            self._json(upd.check_update())
            return
        if path == "/api/convert-info":
            cli = filee_cli()
            self._json({"available": bool(cli), "formats": CONVERT_FORMATS,
                        "path": cli})
            return
        if path == "/api/config":
            k = load_key()
            cfg = load_config()
            from translate_xlsx import PROVIDERS
            self._json({"has_key": bool(k), "tail": (k[-4:] if k else ""),
                        "ollama_url": cfg.get("ollama_url", ""),
                        "provider": current_provider(),
                        "keys": {kk: ("已填" if (vv or "").strip() else "") for kk, vv in (cfg.get("keys") or {}).items()},
                        "custom_url": cfg.get("custom_url", ""),
                        "custom_model": cfg.get("custom_model", ""),
                        "providers": [{"id": pid, "name": pv["name"]} for pid, pv in PROVIDERS.items()]})
            return
        if path.startswith("/api/suggestions/"):
            job = JOBS.get(path.rsplit("/", 1)[-1])
            if not job:
                self._json({"error": "no such job"}, 404)
                return
            self._json({"items": job.get("suggestions") or [],
                        "done": bool(job.get("suggest_done")),
                        "error": job.get("suggest_error", "")})
            return
        if path.startswith("/api/suggest/"):
            st = SUGG.get(path.rsplit("/", 1)[-1])
            if not st:
                self._json({"error": "no such"}, 404)
                return
            self._json({"state": st["state"], "items": st.get("items", []),
                        "error": st.get("error", "")})
            return
        if path.startswith("/api/issues/"):
            job = JOBS.get(path.rsplit("/", 1)[-1])
            if not job or job["state"] != "done":
                self._json({"error": "not ready"}, 404)
                return
            kind = job.get("kind", "xlsx")
            orig = extract_texts(base64.b64decode(job["src"]), kind)
            cur = extract_texts(apply_edits(base64.b64decode(job["data"]), kind,
                                            job.get("edits", {})), kind)
            issues = check_quality(orig, cur, job.get("direction", "en2zh"))
            self._json({"issues": issues, "count": len(issues)})
            return
        if path.startswith("/api/entries/"):
            job = JOBS.get(path.rsplit("/", 1)[-1])
            if not job or job["state"] != "done":
                self._json({"error": "not ready"}, 404)
                return
            kind = job.get("kind", "xlsx")
            orig = extract_texts(base64.b64decode(job["src"]), kind)
            now = extract_texts(base64.b64decode(job["data"]), kind)
            self._json({"orig": orig, "now": now, "edits": job.get("edits", {})})
            return
        if path.startswith("/api/check/"):
            chk = CHECKS.get(path.rsplit("/", 1)[-1])
            if not chk:
                self._json({"error": "no such check"}, 404)
                return
            self._json({"state": chk["state"], "done": chk.get("done", 0),
                        "total": chk.get("total", 0), "error": chk.get("error", ""),
                        "items": chk.get("items", []), "count": chk.get("count", 0)})
            return
        self._json({"error": "not found"}, 404)

    def do_POST(self):
        path = urllib.parse.urlparse(self.path).path
        length = int(self.headers.get("Content-Length", 0))
        raw = self.rfile.read(length)
        if path == "/api/translate":
            try:
                req = json.loads(raw.decode("utf-8"))
                name = req["name"]
                try:
                    data = base64.b64decode(req["data"], validate=True)
                except Exception:
                    raise ValueError("文件内容读不出来，请重新拖一次")
                if not data:
                    raise ValueError("文件是空的")
                direction = req.get("direction", "en2zh")
                bilingual = bool(req.get("bilingual"))
                mode = req.get("mode", "cloud")
                ollama_url = (req.get("ollama_url") or "").strip() or load_config().get("ollama_url", "")
                provider = (req.get("provider") or current_provider()).strip()
                cfg_ = load_config()
                base_url = (cfg_.get("keys" if False else "custom_url") or "").strip()
                cloud_model = (cfg_.get("custom_model") or "").strip()
            except Exception as ex:
                self._json({"error": f"bad request: {ex}"}, 400)
                return
            if mode != "local" and not load_key():
                self._json({"error": "还没有填 API Key，点页面顶上的提示填一下"}, 400)
                return
            jid = uuid.uuid4().hex[:12]
            base = os.path.splitext(os.path.basename(name))[0]
            kind = "docx" if name.lower().endswith(".docx") else "xlsx"
            ext = ".docx" if kind == "docx" else ".xlsx"
            suffix = "中文" if direction == "en2zh" else "English"
            if bilingual:
                suffix = "中英对照" if direction == "en2zh" else "英中对照"
            out_name = f"{base}_{suffix}{ext}"
            tmpdir = tempfile.mkdtemp(prefix="trans_")
            in_path = os.path.join(tmpdir, "in.xlsx")
            out_path = os.path.join(tmpdir, "out.xlsx")
            with open(in_path, "wb") as f:
                f.write(data)
            JOBS[jid] = {"state": "queued", "done": 0, "total": 0, "error": "",
                         "out_name": out_name, "src": req["data"], "edits": {},
                         "direction": direction, "bilingual": bilingual,
                         "kind": kind, "mode": mode}
            threading.Thread(target=worker,
                             args=(jid, in_path, out_path, direction, bilingual, kind, mode,
                                   ollama_url, provider, base_url, cloud_model),
                             daemon=True).start()
            self._json({"id": jid, "out_name": out_name})
            return

        if path == "/api/check":
            try:
                req = json.loads(raw.decode("utf-8"))
                jid = req["id"]
                direction = req.get("direction", "en2zh")
                bilingual = bool(req.get("bilingual"))
                mode = req.get("mode", "cloud")
                ollama_url = (req.get("ollama_url") or "").strip() or load_config().get("ollama_url", "")
                provider = (req.get("provider") or current_provider()).strip()
                cfg_ = load_config()
                base_url = (cfg_.get("keys" if False else "custom_url") or "").strip()
                cloud_model = (cfg_.get("custom_model") or "").strip()
                job = JOBS[jid]
            except Exception:
                self._json({"error": "bad request"}, 400)
                return
            if job["state"] != "done":
                self._json({"error": "翻译还没完成"}, 400)
                return
            cid = uuid.uuid4().hex[:12]
            CHECKS[cid] = {"state": "queued", "done": 0, "total": 0, "error": ""}
            cur = base64.b64encode(
                rebuild_xlsx(base64.b64decode(job["data"]), job.get("edits", {}))).decode("ascii")
            threading.Thread(target=check_worker,
                             args=(cid, job["src"], cur, direction,
                                   job.get("mode", "cloud"), job.get("kind", "xlsx")),
                             daemon=True).start()
            self._json({"id": cid})
            return

        if path == "/api/config":
            try:
                req = json.loads(raw.decode("utf-8"))
                cfg = load_config()
                provider = (req.get("provider") or cfg.get("provider") or "deepseek").strip()
                cfg["provider"] = provider
                if "api_key" in req:
                    keys = cfg.get("keys") or {}
                    keys[provider] = (req.get("api_key") or "").strip()
                    cfg["keys"] = keys
                    if provider == "deepseek":
                        cfg["api_key"] = keys[provider]
                if "ollama_url" in req:
                    cfg["ollama_url"] = (req.get("ollama_url") or "").strip()
                if "custom_url" in req:
                    cfg["custom_url"] = (req.get("custom_url") or "").strip()
                if "custom_model" in req:
                    cfg["custom_model"] = (req.get("custom_model") or "").strip()
                with open(CONFIG_PATH, "w", encoding="utf-8") as f:
                    json.dump(cfg, f, ensure_ascii=False, indent=1)
                self._json({"ok": True, "has_key": bool(load_key(provider)),
                            "provider": provider})
            except Exception as ex:
                self._json({"error": str(ex)}, 400)
            return

        if path == "/api/test-key":
            try:
                req = json.loads(raw.decode("utf-8"))
                provider = (req.get("provider") or current_provider()).strip()
                key = (req.get("api_key") or "").strip() or load_key(provider)
                if not key:
                    self._json({"ok": False, "error": "还没填 Key"})
                    return
                from translate_xlsx import call_api
                got = call_api(key, [("k0", "", "TEST")], "en2zh", provider=provider)
                self._json({"ok": bool(got), "tail": key[-4:]})
            except Exception as ex:
                msg = str(ex)
                if "401" in msg or "invalid" in msg.lower() or "authentication" in msg.lower():
                    msg = "Key 无效（服务方拒绝认证）"
                elif "402" in msg or "balance" in msg.lower() or "insufficient" in msg.lower():
                    msg = "Key 有效，但余额不足"
                self._json({"ok": False, "error": msg[:200]})
            return

        if path == "/api/update-apply":
            try:
                req = json.loads(raw.decode("utf-8"))
                url = (req.get("url") or "").strip()
                if not url:
                    info = upd.check_update()
                    url = info.get("url", "")
                if not url:
                    self._json({"error": "没有可用的更新包地址"}, 400)
                    return
                n = upd.apply_patch(url, EXE_DIR)
                self._json({"ok": True, "files": n,
                            "note": "更新完成，重启软件后生效"})
            except Exception as ex:
                self._json({"error": str(ex)[:300]}, 500)
            return

        if path == "/api/restart":
            # 先回包，再重启，避免前端以为失败
            self._json({"ok": True})
            threading.Timer(0.8, _restart_self).start()
            return

        if path == "/api/convert":
            try:
                req = json.loads(raw.decode("utf-8"))
                name = req["name"]
                data = base64.b64decode(req["data"])
                to_fmt = (req.get("to") or "").strip().lower()
            except Exception:
                self._json({"error": "bad request"}, 400)
                return
            if not to_fmt:
                self._json({"error": "没选目标格式"}, 400)
                return
            tmpdir = tempfile.mkdtemp(prefix="conv_")
            base = os.path.splitext(os.path.basename(name))[0]
            src = os.path.join(tmpdir, os.path.basename(name) or ("in." + to_fmt))
            outdir = os.path.join(tmpdir, "out")
            os.makedirs(outdir, exist_ok=True)
            with open(src, "wb") as f:
                f.write(data)
            try:
                out_path, log = run_convert(src, outdir, to_fmt)
                with open(out_path, "rb") as f:
                    blob = f.read()
                cid = uuid.uuid4().hex[:12]
                CONV[cid] = {"name": os.path.basename(out_path), "data": blob}
                self._json({"ok": True, "id": cid,
                            "out_name": os.path.basename(out_path),
                            "size": len(blob), "log": log[-300:]})
            except Exception as ex:
                self._json({"error": str(ex)}, 500)
            return

        if path == "/api/suggest":
            try:
                req = json.loads(raw.decode("utf-8"))
                job = JOBS[req["id"]]
            except Exception:
                self._json({"error": "bad request"}, 400)
                return
            sid = uuid.uuid4().hex[:12]
            SUGG[sid] = {"state": "queued", "items": []}
            tmpdir = tempfile.mkdtemp(prefix="sugg_")
            pth = os.path.join(tmpdir, "src" + (".docx" if job.get("kind") == "docx" else ".xlsx"))
            with open(pth, "wb") as f:
                f.write(base64.b64decode(job["src"]))
            threading.Thread(target=suggest_worker, args=(sid, [pth]), daemon=True).start()
            self._json({"id": sid})
            return

        if path == "/api/terms-add":
            try:
                req = json.loads(raw.decode("utf-8"))
                items = req.get("items", [])
            except Exception:
                self._json({"error": "bad request"}, 400)
                return
            added = add_terms(items)
            translate_xlsx.GLOSSARY_TEXT = None      # 让新表立刻生效
            self._json({"ok": True, "added": added, "total": len(load_terms_file())})
            return

        if path == "/api/fix-all":
            try:
                req = json.loads(raw.decode("utf-8"))
                job = JOBS[req["id"]]
            except Exception:
                self._json({"error": "bad request"}, 400)
                return
            try:
                k_ = job.get("kind", "xlsx")
                orig = extract_texts(base64.b64decode(job["src"]), k_)
                cur = extract_texts(apply_edits(base64.b64decode(job["data"]), k_,
                                                job.get("edits", {})), k_)
                issues = check_quality(orig, cur, job.get("direction", "en2zh"))
                idx = [x["i"] for x in issues]
                fixed = retranslate_batch(job, idx)
                self._json({"ok": True, "fixed": fixed, "tried": len(idx)})
            except Exception as ex:
                traceback.print_exc()
                self._json({"error": str(ex)}, 500)
            return

        if path == "/api/edit":
            try:
                req = json.loads(raw.decode("utf-8"))
                job = JOBS[req["id"]]
            except Exception:
                self._json({"error": "bad request"}, 400)
                return
            idx = str(req["i"])
            text = req.get("text", "")
            if text.strip() == "":
                job["edits"].pop(idx, None)
            else:
                job["edits"][idx] = text
            self._json({"ok": True, "edits": len(job["edits"])})
            return

        if path == "/api/reset":
            try:
                req = json.loads(raw.decode("utf-8"))
                job = JOBS[req["id"]]
            except Exception:
                self._json({"error": "bad request"}, 400)
                return
            job["edits"] = {}
            self._json({"ok": True})
            return

        if path == "/api/retranslate":
            try:
                req = json.loads(raw.decode("utf-8"))
                job = JOBS[req["id"]]
                i = int(req["i"])
                hint = (req.get("hint") or "").strip()
                direction = job.get("direction", "en2zh")
                orig = extract_texts(base64.b64decode(job["src"]),
                                     job.get("kind", "xlsx"))[i]
                text = retranslate_one(orig, direction, hint)
                self._json({"text": text})
            except Exception as ex:
                traceback.print_exc()
                self._json({"error": str(ex)}, 500)
            return

        self._json({"error": "not found"}, 404)


def _prepare_runtime():
    """打包成无控制台程序时，把输出写进日志文件，别让它因为 print 崩掉。"""
    if not getattr(sys, "frozen", False):
        return
    log = os.path.join(EXE_DIR, "运行日志.txt")
    try:
        f = open(log, "a", encoding="utf-8", buffering=1)
        if sys.stdout is None:
            sys.stdout = f
        if sys.stderr is None:
            sys.stderr = f
    except Exception:
        pass


def _alert(msg):
    """没有控制台时，用系统弹窗告诉用户出了什么事。"""
    try:
        import ctypes
        ctypes.windll.user32.MessageBoxW(0, msg, "保格式翻译", 0x10)
    except Exception:
        print(msg)


def main():
    _prepare_runtime()
    try:
        srv = ThreadingHTTPServer(("127.0.0.1", PORT), Handler)
    except OSError as ex:
        _alert("启动失败：端口 %d 被占用。\n请关掉之前打开的程序（或重启电脑）再试。" % PORT)
        print("端口占用:", ex)
        return
    url = f"http://127.0.0.1:{PORT}/"
    threading.Thread(target=srv.serve_forever, daemon=True).start()

    # 优先用原生窗口；没有 WebView 环境时退回浏览器
    try:
        import webview
        print("=" * 52)
        print("  保格式翻译 · 窗口模式")
        print("=" * 52)
        webview.create_window("保格式翻译", url, width=1200, height=880,
                              min_size=(960, 660))
        webview.start()
        return
    except Exception as ex:
        print("窗口模式不可用，改用浏览器打开：", ex)

    print("=" * 52)
    print("  保格式翻译 · 已启动")
    print("  " + url)
    print("  关掉这个黑窗口就停止服务")
    print("=" * 52)
    threading.Timer(1.0, lambda: None if os.environ.get("NO_BROWSER")
                    else webbrowser.open(url)).start()
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        pass


if __name__ == "__main__":
    main()
