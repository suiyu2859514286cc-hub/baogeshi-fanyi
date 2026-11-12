# -*- coding: utf-8 -*-
"""自动更新：查 GitHub 最新版本，下载补丁包，就地替换。"""
import json
import os
import shutil
import tempfile
import urllib.request
import zipfile

# 更新源（改仓库名就行）
GITHUB_REPO = "suiyu2859514286cc-hub/baogeshi-fanyi"
VERSION = "2026.11.12.1"

UA = {"User-Agent": "baogeshi-translator"}


def _open(url, timeout=30):
    """更新走直连：很多机器挂着没开的代理，走代理只会失败。"""
    proxy_keys = ("HTTP_PROXY", "HTTPS_PROXY", "http_proxy", "https_proxy",
                  "ALL_PROXY", "all_proxy")
    saved = {}
    for k in proxy_keys:
        if k in os.environ:
            saved[k] = os.environ.pop(k)
    try:
        req = urllib.request.Request(url, headers=UA)
        opener = urllib.request.build_opener(urllib.request.ProxyHandler({}))
        return opener.open(req, timeout=timeout)
    finally:
        os.environ.update(saved)


def check_update():
    """返回 {current, latest, has_update, notes, url, error}"""
    out = {"current": VERSION, "latest": "", "has_update": False,
           "notes": "", "url": "", "error": ""}
    try:
        with _open("https://api.github.com/repos/%s/releases/latest" % GITHUB_REPO) as r:
            data = json.loads(r.read().decode("utf-8"))
        latest = (data.get("tag_name") or "").lstrip("vV")
        out["latest"] = latest
        out["notes"] = (data.get("body") or "")[:600]
        for a in data.get("assets") or []:
            if a.get("name", "").lower().endswith(".zip"):
                out["url"] = a["browser_download_url"]
                break
        out["has_update"] = bool(latest) and latest != VERSION
    except Exception as ex:
        msg = str(ex)
        if "404" in msg:
            msg = "更新源还没建好（GitHub 上找不到这个仓库或还没有发布版本）"
        out["error"] = msg[:200]
    return out


def apply_patch(patch_url, target_dir, on_progress=None):
    """下载补丁包并解压覆盖到程序目录。"""
    tmp = tempfile.mkdtemp(prefix="upd_")
    zpath = os.path.join(tmp, "patch.zip")
    try:
        with _open(patch_url, timeout=300) as r, open(zpath, "wb") as f:
            total = int(r.headers.get("Content-Length") or 0)
            got = 0
            while True:
                chunk = r.read(65536)
                if not chunk:
                    break
                f.write(chunk)
                got += len(chunk)
                if on_progress and total:
                    on_progress(got, total)

        with zipfile.ZipFile(zpath) as z:
            names = z.namelist()
            # 统一去掉顶层目录（如果打包时带了）
            tops = {n.split("/")[0] for n in names if "/" in n}
            strip = ""
            if len(tops) == 1:
                only = next(iter(tops))
                if all(n == only or n.startswith(only + "/") for n in names):
                    strip = only + "/"
            count = 0
            for n in names:
                if n.endswith("/"):
                    continue
                rel = n[len(strip):] if strip and n.startswith(strip) else n
                if not rel:
                    continue
                dst = os.path.join(target_dir, rel.replace("/", os.sep))
                if os.path.abspath(dst).lower().startswith(os.path.abspath(target_dir).lower()) is False:
                    continue  # 防止写到外面去
                os.makedirs(os.path.dirname(dst), exist_ok=True)
                with z.open(n) as src, open(dst, "wb") as out:
                    shutil.copyfileobj(src, out)
                count += 1
        return count
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


if __name__ == "__main__":
    info = check_update()
    print(json.dumps(info, ensure_ascii=False, indent=1))
