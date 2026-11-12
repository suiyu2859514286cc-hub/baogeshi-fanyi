# -*- coding: utf-8 -*-
"""保格式 xlsx 翻译 v2：只改 sharedStrings，其他部件字节级原样保留。
改进：术语表注入 / 列上下文 / 富文本保留红字 / key 容错。"""
import sys, os, re, json, zipfile, time, urllib.request
from lxml import etree

NS = {'m': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
M = NS['m']
# 密钥文件：优先环境变量，其次程序目录下的 chat-key.txt（都不放就别用命令行模式）
KEY_FILE = (os.environ.get("BAOGESHI_KEY_FILE")
            or os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "chat-key.txt"))

GLOSSARY = """ITEM=编号
DESCRIPTION=项目描述
UNIT=单位
QTY=数量
RATE=单价
TOTAL=金额
Bills of Quantities/BOQ=工程量清单
SUMMARY=汇总
TOTAL CARRIED FORWARD TO THE SUMMARY PAGE=结转至汇总页
TOTAL CARRIED TO NEXT PAGE=结转下页
TOTAL BROUGHT FORWARD FROM PREVIOUS PAGE=承前页合计
B/f=brought forward=承前
Grand total for the Works=全部工程总计
MCB=微型断路器
MCCB=塑壳断路器
ACB=空气断路器
DB/Distribution Board=配电箱
TPN=三相四线
SPN=单相
TP=三相
SP=单相
isolator=隔离开关
busbar=母排
cable tray=电缆桥架
trunking=线槽
conduit=线管
draw-wire=引线
cable gland=电缆格兰头
shroud=护套
SWA=SWA（保留英文，电缆型号标识）
XLPE=XLPE（保留英文，电缆型号标识）
PVC=PVC（保留英文，电缆型号标识）
downlighter=筒灯
luminaire=灯具
bulkhead=壁装灯
occupancy sensor=人体感应器
PIR=被动红外
emergency light=应急灯
exit light=安全出口灯
fire alarm=火灾报警
addressable=可寻址
commissioning=调试
testing and commissioning=测试与调试
PC Sum=暂定金额
generating set=发电机组
genset=机组
prime rated=主用功率
fuel tank=燃油箱
belly tank=底部油箱
exhaust pipe=排气管
silencer=消声器
earthing=接地
earth electrode=接地极
UPS/Uninterruptible Power Supply=不间断电源
synchronising panel=同步屏
AMF=自动市电切换
ATS=自动转换开关
DOL starter=直接启动器
busbar plug in=母线插接箱
data/voice outlet=数据/语音出线口
CCTV=闭路电视监控
light fitting=灯具
power point=电源插座
socket outlet=插座
blanking cover=盲板
sub-main=干线
riser=竖井
Demolitions=拆除工程
Record drawings=竣工图
No (unit)=个
Lm=延米
Allow for=计列（含）……费用
Item (section marker)=项目
Ditto=同上
"""

BRANDS = (r"(Schneider Electric|Schneider|MK|ABB|Philips|Osram|Ledvance|Thorn|JCC|Protec|"
          r"Crabtree|Vertiv|Perkins|Cummins|Caterpillar|KATKO|Robus|BG|MICROMARK|Legrand|"
          r"Siemens|Hager|Ovano)")
UNITS = r"(kVA|KVA|kV|V|Hz|kW|W|dB|kg|mm²|mm2|mm|cm|m2|m|A|SWG|hours?|hrs?|min|sec|°C|%)"
MODEL = r"[A-Z]{1,6}[- ]?[0-9][A-Z0-9\-/.]{1,}"

MONTH = (r"(?:January|February|March|April|May|June|July|August|"
         r"September|October|November|December|Jan|Feb|Mar|Apr|Jun|Jul|Aug|Sep|Oct|Nov|Dec)")
DATE = (r"\d{1,2}(?:st|nd|rd|th)?\s+" + MONTH + r"[,.]?\s+\d{4}"
        r"|\d{4}-\d{2}-\d{2}"
        r"|" + MONTH + r"\s+\d{1,2}(?:st|nd|rd|th)?[,.]?\s+\d{4}")

PROTECT = re.compile(
    r"(" + DATE +
    r"|NO\.?\s*\d+"
    r"|(?:KSHS?|KShs?|Kshs?|KES)\.?"
    r"|\d+(?:\.\d+)?\s*(?:mm|cm|m)\s*sq(?:uare)?\b"
    r"|" + MODEL +
    r"|\d+(?:\.\d+)?\s*" + UNITS + r"\b"
    r"|\d+(?:[.,]\d+)*"
    r"|" + BRANDS + r")"
)


def protect(text):
    text = re.sub(r'(\d+(?:\.\d+)?)\s*mm\s*sq(?:uare)?\b', r'\1mm²', text)
    spans = []
    def repl(m):
        spans.append(m.group(0))
        return f"⟦{len(spans)-1}⟧"
    return PROTECT.sub(repl, text), spans


def restore(text, spans):
    def repl(m):
        i = int(m.group(1))
        if i >= len(spans):
            raise ValueError(f"占位符越界 {i}/{len(spans)}")
        return spans[i]
    return re.sub(r"⟦(\d+)⟧", repl, text)


QTY_UNIT = r"(?:A|mA|V|kV|W|kW|kVA|KVA|Hz|mm|cm|m|kg|sec|min|h)"
CLEAN1 = re.compile(r"(" + QTY_UNIT + r")\s*个")
CLEAN2 = re.compile(r"(\d)\s*个\s*(?=双联|三联|单联|极|路|位|芯|控)")


def clean_artifacts(t):
    """清掉模型在规格后面自作主张加的量词。"""
    t = CLEAN1.sub(r"\1 ", t)
    t = CLEAN2.sub(r"\1", t)
    t = re.sub(r" {2,}", " ", t)
    return t


def load_glossary_text():
    """术语表以审定的 glossary.json 为唯一权威；没有就退回内置简表。"""
    import sys as _sys
    here = os.path.dirname(os.path.abspath(__file__))
    cands = []
    # 打包后：exe 旁边的「术语表」优先，这样换表不用重装程序
    if getattr(_sys, "frozen", False):
        exe_dir = os.path.dirname(_sys.executable)
        cands.append(os.path.join(exe_dir, "术语表", "glossary.json"))
    cands += [os.path.join(here, "..", "术语表", "glossary.json"),
              os.path.join(here, "术语表", "glossary.json"),
              os.path.join(here, "glossary.json")]
    for p in cands:
        if os.path.exists(p):
            with open(p, encoding="utf-8") as f:
                data = json.load(f)
            lines = [" / ".join(it["en"]) + "=" + it["zh"] for it in data["items"]]
            print(f"[术语表] 已加载 {p}（{len(lines)} 条）")
            return "\n".join(lines)
    print("[术语表] 未找到 glossary.json，使用内置简表")
    return GLOSSARY


USAGE = {"prompt": 0, "completion": 0, "calls": 0}

GLOSSARY_TEXT = None


def glossary_text():
    global GLOSSARY_TEXT
    if GLOSSARY_TEXT is None:
        GLOSSARY_TEXT = load_glossary_text()
    return GLOSSARY_TEXT


PROVIDERS = {
    "deepseek": {"name": "DeepSeek", "url": "https://api.deepseek.com/chat/completions",
                 "model": "deepseek-chat", "style": "openai"},
    "openai": {"name": "OpenAI", "url": "https://api.openai.com/v1/chat/completions",
               "model": "gpt-4o-mini", "style": "openai"},
    "zhipu": {"name": "智谱 GLM", "url": "https://open.bigmodel.cn/api/paas/v4/chat/completions",
              "model": "glm-4-flash", "style": "openai"},
    "moonshot": {"name": "月之暗面 Kimi", "url": "https://api.moonshot.cn/v1/chat/completions",
                 "model": "moonshot-v1-8k", "style": "openai"},
    "qwen": {"name": "阿里通义千问", "url": "https://dashscope.aliyuncs.com/compatible-mode/v1/chat/completions",
             "model": "qwen-plus", "style": "openai"},
    "claude": {"name": "Anthropic Claude", "url": "https://api.anthropic.com/v1/messages",
               "model": "claude-3-5-haiku-latest", "style": "claude"},
    "custom": {"name": "自定义（OpenAI 兼容）", "url": "", "model": "", "style": "openai"},
}


def call_api(api_key, items, direction, hint="", mode="cloud", model="qwen2.5:7b",
             ollama_url="", provider="deepseek", base_url="", cloud_model=""):
    if direction == "en2zh":
        sys_prompt = (
            "你是建筑工程、电气分包与合同领域的英中技术翻译，为工程量清单（BOQ）做译文。\n"
            "规则：\n"
            "1. 术语必须严格按下面这张对照表译，不得自行换词：\n" + glossary_text() + "\n"
            "2. 保留所有 ⟦n⟧ 占位符原样、顺序不变、一个不少；保留原有换行。\n"
            "3. 型号、品牌、规格、数字已在占位符里，不要翻译也不要改动。\n"
            "4. 语域正式书面，不解释、不加注、不合并条目。\n"
            "5. 每条给了一个 col 字段（该单元格所在列的列名），据此判断词义，"
            "例如 UNIT 列里的 No 是数量单位「个」，不是否定词。\n"
            "6. 不要在 ⟦n⟧ 占位符后面自作主张加量词（不要写成「⟦0⟧个」这种）。\n"
            + (("\n附加要求（必须遵守）：" + hint) if hint else "")
            + ambiguous_hint([t for _, _, t in items])
            + "\n返回严格 JSON 对象，键是条目 id，值是译文。"
        )
    else:
        sys_prompt = (
            "你是建筑工程、电气分包与合同领域的中英技术翻译。任务：把每条 text 里**中文**翻译成**英文**，"
            "给国外业主和咨询工程师看的工程量清单（BOQ）正文。译文里不许残留任何汉字（型号、品牌、单位除外）。\n"
            "规则：\n"
            "1. 术语必须严格按下面这张对照表反向取词，不得自行换词：\n" + glossary_text() + "\n"
            "2. 保留所有 ⟦n⟧ 占位符原样、顺序不变、一个不少；保留原有换行。\n"
            "3. 型号、品牌、规格、数字已在占位符里，不要翻译也不要改动。\n"
            "4. 语法必须像英国/肯尼亚咨询工程师写的清单，不是中文语序直译：\n"
            "   - 用祈使/被动句式：Supply, deliver to site, install, test and commission ...；Allow for ...\n"
            "   - 数量用 one / two / 30No. 这类写法，不要用 a/an 硬凑\n"
            "   - 冠词、单复数、时态要正确；分项列表保持 a) b) c) 的编号格式\n"
            "5. 一条译文对应一条原文，不要合并、不要拆分、不要解释。\n"
            + ambiguous_hint([t for _, _, t in items])
            + (("\n附加要求（必须遵守）：" + hint) if hint else "")
            + "\n返回严格 JSON 对象，键是条目 id，是英文译文。"
        )

    user = json.dumps([{"id": i, "col": c, "text": t} for i, c, t in items],
                      ensure_ascii=False, indent=1)

    if mode == "local":
        return call_local(sys_prompt, items, model, ollama_url)

    return call_cloud(provider, api_key, sys_prompt, user, base_url, cloud_model)


def call_cloud(provider, api_key, sys_prompt, user, base_url="", model=""):
    """统一入口：OpenAI 兼容的一大票，Claude 单独一套。"""
    conf = PROVIDERS.get(provider, PROVIDERS["deepseek"])
    url = base_url.strip() or conf["url"]
    mdl = model.strip() or conf["model"]
    if not url:
        raise RuntimeError("这个厂商还没填接口地址（Base URL）。")
    if conf.get("style") == "claude":
        return _call_claude(api_key, url, mdl, sys_prompt, user)
    return _call_openai_like(api_key, url, mdl, sys_prompt, user)


def _call_openai_like(api_key, url, model, sys_prompt, user):
    body = json.dumps({
        "model": model,
        "messages": [{"role": "system", "content": sys_prompt},
                     {"role": "user", "content": "条目：\n" + user}],
        "temperature": 0.15,
    }).encode("utf-8")
    req = urllib.request.Request(
        url, data=body,
        headers={"Content-Type": "application/json",
                 "Authorization": "Bearer " + api_key})
    with urllib.request.urlopen(req, timeout=180) as r:
        data = json.loads(r.read().decode("utf-8"))
    _count_usage(data)
    content = data["choices"][0]["message"]["content"]
    return _parse_json(content)


def _call_claude(api_key, url, model, sys_prompt, user):
    body = json.dumps({
        "model": model,
        "max_tokens": 8192,
        "system": sys_prompt,
        "messages": [{"role": "user", "content": "条目：\n" + user}],
        "temperature": 0.15,
    }).encode("utf-8")
    req = urllib.request.Request(
        url, data=body,
        headers={"Content-Type": "application/json",
                 "x-api-key": api_key,
                 "anthropic-version": "2023-06-01"})
    with urllib.request.urlopen(req, timeout=180) as r:
        data = json.loads(r.read().decode("utf-8"))
    u = data.get("usage") or {}
    USAGE["prompt"] += u.get("input_tokens", 0)
    USAGE["completion"] += u.get("output_tokens", 0)
    USAGE["calls"] += 1
    parts = data.get("content") or []
    content = "".join(p.get("text", "") for p in parts if isinstance(p, dict))
    return _parse_json(content)


def _count_usage(data):
    u = data.get("usage") or {}
    USAGE["prompt"] += u.get("prompt_tokens", 0)
    USAGE["completion"] += u.get("completion_tokens", 0)
    USAGE["calls"] += 1


def _parse_json(content):
    """不同家的模型偶尔会在 JSON 外面裹一层文字，抠出来。"""
    content = (content or "").strip()
    try:
        return json.loads(content)
    except Exception:
        pass
    a, b = content.find("{"), content.rfind("}")
    if a >= 0 and b > a:
        return json.loads(content[a:b + 1])
    raise RuntimeError("模型返回的不是 JSON：" + content[:120])


_OLLAMA_CACHE = {"url": "", "at": 0.0}


def pick_ollama(base_url=""):
    """按顺序试每一个候选地址，返回第一个能连上的。"""
    import time as _t
    import urllib.request
    raw = (base_url or "").replace("，", ",").replace(";", ",")
    urls = [u.strip().rstrip("/") for u in raw.split(",") if u.strip()]
    urls.append("http://127.0.0.1:11434")
    now = _t.time()
    cached = _OLLAMA_CACHE["url"]
    if cached and now - _OLLAMA_CACHE["at"] < 120:
        urls = [cached] + [u for u in urls if u != cached]
    for u in urls:
        try:
            urllib.request.urlopen(u + "/api/tags", timeout=3).read()
            if u != cached:
                print("[本地模型] 使用", u)
            _OLLAMA_CACHE["url"] = u
            _OLLAMA_CACHE["at"] = now
            return u
        except Exception:
            continue
    raise RuntimeError(
        "连不上任何一个本地模型地址（试过：%s）。\n"
        "请确认那台机器开着 Ollama、网络能通；或者把引擎切回「云端」。" % "、".join(urls))


def call_local(sys_prompt, items, model="qwen2.5:7b", base_url=""):
    """本地 Ollama 初翻。质量弱于云端，只作断网应急。"""
    import urllib.request
    base = pick_ollama(base_url)
    user = json.dumps([{"id": i, "col": c, "text": t} for i, c, t in items],
                      ensure_ascii=False)
    body = json.dumps({
        "model": model,
        "messages": [{"role": "system", "content": sys_prompt},
                     {"role": "user", "content": "条目：\n" + user}],
        "stream": False,
        "format": "json",
        "options": {"temperature": 0.2, "num_ctx": 8192},
    }).encode("utf-8")
    req = urllib.request.Request(base + "/api/chat", data=body,
                                 headers={"Content-Type": "application/json"})
    with urllib.request.urlopen(req, timeout=600) as r:
        data = json.loads(r.read().decode("utf-8"))
    USAGE["calls"] += 1
    return json.loads(data["message"]["content"])


AMBIG = None


def load_ambiguous():
    global AMBIG
    if AMBIG is not None:
        return AMBIG
    import sys as _sys
    here = os.path.dirname(os.path.abspath(__file__))
    cands = []
    if getattr(_sys, "frozen", False):
        cands.append(os.path.join(os.path.dirname(_sys.executable), "术语表", "ambiguous.json"))
    cands += [os.path.join(here, "..", "术语表", "ambiguous.json"),
              os.path.join(here, "术语表", "ambiguous.json")]
    AMBIG = {}
    for p_ in cands:
        if os.path.exists(p_):
            try:
                with open(p_, encoding="utf-8") as f:
                    for it in json.load(f):
                        AMBIG[it["en"].lower()] = it["senses"]
            except Exception:
                pass
            break
    return AMBIG


def ambiguous_hint(texts):
    """只挑出这一批文本里真正出现的歧义词，按需注入。"""
    table = load_ambiguous()
    if not table:
        return ""
    blob = " ".join(texts).lower()
    hits = []
    for en, senses in table.items():
        if not re.search(r"(?<![a-z])" + re.escape(en) + r"(?![a-z])", blob):
            continue
        parts = "；".join("「%s」当%s" % (s["zh"], (s.get("when", "") or "")[:26])
                           for s in senses if s.get("zh"))
        if parts:
            hits.append("- %s：%s" % (en, parts))
    if not hits:
        return ""
    return ("\n以下是这批里出现的多义词，必须按语境选对：\n" + "\n".join(hits) + "\n")


def col_of(ref):
    return re.match(r"([A-Z]+)", ref).group(1)


def build_headers(parts, names, sis):
    """返回 si_idx -> 列名（表头行里该列的文本）"""
    header_cols = {}   # (sheet, col) -> text
    cell_header = {}   # si_idx -> text
    for n in names:
        if not re.match(r'xl/worksheets/sheet\d+\.xml$', n):
            continue
        root = etree.fromstring(parts[n])
        rows = root.findall('.//m:row', NS)
        head_row = None
        for row in rows[:15]:
            texts = []
            for c in row.findall('m:c', NS):
                if c.get('t') == 's':
                    v = c.find('m:v', NS)
                    if v is not None and v.text:
                        texts.append(sis[int(v.text)].strip())
            if any(t.upper() in ("ITEM", "DESCRIPTION", "QTY") or t in ("编号", "项目描述", "数量")
                   for t in texts):
                head_row = row
                break
        if head_row is None:
            continue
        for c in head_row.findall('m:c', NS):
            ref = c.get('r')
            val = ''
            if c.get('t') == 's':
                v = c.find('m:v', NS)
                if v is not None and v.text:
                    val = sis[int(v.text)].strip()
            elif c.find('m:v', NS) is not None:
                val = (c.find('m:v', NS).text or '').strip()
            if val:
                header_cols[(n, col_of(ref))] = val.replace('\n', ' ')
        for c in root.iter('{%s}c' % M):
            if c.get('t') != 's':
                continue
            v = c.find('m:v', NS)
            if v is None or not v.text:
                continue
            si = int(v.text)
            h = header_cols.get((n, col_of(c.get('r'))))
            if h and si not in cell_header:
                cell_header[si] = h
    return cell_header


def translate_file(src, dst, api_key, direction="en2zh", max_items=None, on_progress=None,
                   bilingual=False, mode="cloud", ollama_url="",
                   provider="deepseek", base_url="", cloud_model=""):
    zin = zipfile.ZipFile(src)
    names = zin.namelist()
    parts = {n: zin.read(n) for n in names}
    zin.close()

    ss = etree.fromstring(parts['xl/sharedStrings.xml'])
    sis_el = ss.findall('m:si', NS)

    def si_text(si):
        parts_ = []
        for r in si.findall('m:r', NS):
            t = r.find('m:t', NS)
            parts_.append(t.text or '')
        if parts_:
            return ''.join(parts_)
        t = si.find('m:t', NS)
        return (t.text or '') if t is not None else ''

    sis = [si_text(x) for x in sis_el]
    cell_header = build_headers(parts, names, sis)

    entries = []
    for i, si in enumerate(sis_el):
        rich = si.find('m:r', NS) is not None
        entries.append({"idx": i, "text": sis[i], "rich": rich,
                        "col": cell_header.get(i, "")})

    def needs(e):
        s = e["text"].strip()
        if direction == "zh2en":
            return bool(re.search(r"[\u4e00-\u9fff]", s))
        if len(s) < 2:
            return False
        if re.search(r"[\u4e00-\u9fff]", s):
            return False
        return len(re.findall(r"[A-Za-z]{2,}", s)) >= 1

    todo = [e for e in entries if needs(e)]
    if max_items:
        todo = todo[:max_items]
    print(f"总字符串 {len(entries)}，待译 {len(todo)} 条（富文本 {sum(1 for e in todo if e['rich'])} 条）")

    payload, prot = [], {}
    for e in todo:
        p, spans = protect(e["text"])
        prot[e["idx"]] = spans
        payload.append(("k%d" % e["idx"], e["col"], p))

    results = {}
    last_err = ""
    B = 6
    fails = []
    total_batches = (len(payload) + B - 1) // B
    for start in range(0, len(payload), B):
        batch = payload[start:start + B]
        ok = False
        for attempt in range(3):
            try:
                got = call_api(api_key, batch, direction, mode=mode, ollama_url=ollama_url,
                               provider=provider, base_url=base_url, cloud_model=cloud_model)
                results.update(got)
                ok = True
                print(f"  批次 {start//B+1}/{total_batches} ok")
                break
            except Exception as ex:
                last_err = str(ex)
                low = last_err.lower()
                if "本地模型" in last_err:
                    raise                       # 环境问题，重试也没用，直接告诉用户
                if ("401" in last_err or "403" in last_err or "402" in last_err
                        or "unauthorized" in low or "insufficient" in low
                        or "invalid api key" in low):
                    raise RuntimeError(
                        "API 拒绝了请求（Key 无效、已过期或余额不足）。\n原始信息：" + last_err[:160])
                print(f"  批次 {start//B+1} 失败({attempt+1}/3): {ex}")
                time.sleep(3)
        if not ok:
            fails.extend(b[0] for b in batch)
        if on_progress:
            try:
                on_progress(min(start + B, len(payload)), len(payload))
            except Exception:
                pass
        time.sleep(0.25)

    # 写回
    changed, missing, badph, rich_fallback, short = 0, [], [], [], []
    touched = []
    for e in todo:
        key = "k%d" % e["idx"]
        val = results.get(key)
        if val is None:
            nums = str(e["idx"])
            val = results.get(nums) or results.get("item_" + nums)
        if val is None:
            missing.append(e["idx"])
            continue
        try:
            text = restore(val, prot[e["idx"]])
        except Exception as ex:
            badph.append((e["idx"], str(ex)))
            continue
        if not text.strip():
            missing.append(e["idx"])
            continue
        text = clean_artifacts(text)
        # 漏译检测：中文长度远小于英文，多半是整句没译完
        ratio = len(text) / max(1, len(e["text"]))
        if direction == "en2zh" and ratio < 0.15 and len(e["text"]) > 60:
            short.append((e["idx"], round(ratio, 2)))
        if bilingual:
            text = e["text"].strip() + "\n" + text
            touched.append(e["idx"])
        si = sis_el[e["idx"]]
        if si.find('m:r', NS) is not None:
            if bilingual:
                # 双语模式下富文本直接合并成一段（保留内容，放弃局部着色）
                for ch in list(si):
                    si.remove(ch)
                tt = etree.SubElement(si, '{%s}t' % M)
                tt.text = text
                tt.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
            elif not rebuild_rich(si, sis[e["idx"]], text):
                # 富文本：定位每个 run 的原文在译文里的位置，重建 run
                rich_fallback.append(e["idx"])
                continue
        else:
            t = si.find('m:t', NS)
            if t is None:
                t = etree.SubElement(si, '{%s}t' % M)
            t.text = text
            t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        changed += 1

    total_todo = max(1, len(todo))
    if changed == 0 and total_todo > 2:
        raise RuntimeError(
            "一条都没翻出来，整份文件没有被改动。\n原因：" +
            (last_err[:200] if last_err else "所有批次都失败了，请检查网络或 Key。"))
    if len(missing) > total_todo * 0.5:
        raise RuntimeError(
            "%d 条里有 %d 条没翻出来（超过一半），文件不可用。\n原因：%s"
            % (total_todo, len(missing), (last_err[:200] if last_err else "未知")))
    print(f"改写 {changed}，缺 {len(missing)}，占位符错 {len(badph)}，富文本降级 {len(rich_fallback)}")
    print(f"[用量] 调用 {USAGE['calls']} 次，输入 {USAGE['prompt']} tokens，输出 {USAGE['completion']} tokens")
    if short:
        print(f"  疑似漏译 {len(short)} 条（长度比 <0.28）: {short}")
    if missing:
        print("  缺失下标:", missing)
    parts['xl/sharedStrings.xml'] = etree.tostring(
        ss, xml_declaration=True, encoding='UTF-8', standalone=True)
    if bilingual and touched:
        try:
            enable_wrap(parts, build_refs(parts, names), touched)
            print(f"[双语] 已给 {len(touched)} 条开自动换行")
        except Exception as ex:
            print("  [双语] 开换行失败，不影响文本:", ex)

    with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as zout:
        for n in names:
            zi = zipfile.ZipInfo(n, date_time=(2026, 1, 1, 0, 0, 0))
            zi.compress_type = zipfile.ZIP_DEFLATED
            zout.writestr(zi, parts[n])
    print("写出:", dst)


def build_refs(parts, names):
    """sharedString 下标 -> [(sheet 部件名, 单元格坐标)]"""
    refs = {}
    for n in names:
        if not re.match(r"xl/worksheets/sheet\d+\.xml$", n):
            continue
        root = etree.fromstring(parts[n])
        for c in root.iter("{%s}c" % M):
            if c.get("t") != "s":
                continue
            v = c.find("m:v", NS)
            if v is None or not v.text:
                continue
            refs.setdefault(int(v.text), []).append((n, c.get("r")))
    return refs


def enable_wrap(parts, refs_by_si, changed_sis):
    """双语模式下，给被改写的单元格开自动换行（只动换行，不动字体边框颜色）。"""
    from copy import deepcopy
    targets = {}
    for si in changed_sis:
        for sheet, ref in refs_by_si.get(si, []):
            targets.setdefault(sheet, set()).add(ref)
    if not targets:
        return
    st = etree.fromstring(parts["xl/styles.xml"])
    cellXfs = st.find("m:cellXfs", NS)
    if cellXfs is None:
        return
    xfs = list(cellXfs.findall("m:xf", NS))
    cache = {}

    def wrap_xf(idx):
        if idx in cache:
            return cache[idx]
        src = xfs[idx] if 0 <= idx < len(xfs) else xfs[0]
        new = deepcopy(src)
        al = new.find("m:alignment", NS)
        if al is None:
            al = etree.SubElement(new, "{%s}alignment" % M)
        al.set("wrapText", "1")
        al.set("vertical", "top")
        new.set("applyAlignment", "1")
        cellXfs.append(new)
        xfs.append(new)
        ni = len(xfs) - 1
        cache[idx] = ni
        return ni

    for sheet, refs in targets.items():
        root = etree.fromstring(parts[sheet])
        hit = 0
        for c in root.iter("{%s}c" % M):
            if c.get("r") in refs:
                c.set("s", str(wrap_xf(int(c.get("s") or 0))))
                hit += 1
        if hit:
            parts[sheet] = etree.tostring(root, xml_declaration=True,
                                          encoding="UTF-8", standalone=True)
    cellXfs.set("count", str(len(xfs)))
    parts["xl/styles.xml"] = etree.tostring(st, xml_declaration=True,
                                            encoding="UTF-8", standalone=True)


def rebuild_rich(si, old_text, new_text):
    runs = si.findall('m:r', NS)
    texts = []
    for r in runs:
        t = r.find('m:t', NS)
        texts.append(t.text or '')
    rprs = [r.find('m:rPr', NS) for r in runs]
    # 找出原文里"有区别"的 run（有 color 的）
    special = []
    for i, r in enumerate(runs):
        rpr = rprs[i]
        if rpr is not None and rpr.find('m:color', NS) is not None:
            special.append(i)
    if len(special) != 1:
        return False
    idx = special[0]
    seg = texts[idx]
    if not seg or new_text.count(seg) != 1:
        return False
    pos = new_text.index(seg)
    pre, post = new_text[:pos], new_text[pos + len(seg):]
    for r in runs:
        si.remove(r)
    for txt, rpr in ((pre, rprs[0]), (seg, rprs[idx]), (post, rprs[0])):
        if txt == '':
            continue
        r = etree.SubElement(si, '{%s}r' % M)
        if rpr is not None:
            r.append(copy_rpr(rpr))
        t = etree.SubElement(r, '{%s}t' % M)
        t.text = txt
        t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    return True


def copy_rpr(rpr):
    from copy import deepcopy
    return deepcopy(rpr)


if __name__ == "__main__":
    src, dst = sys.argv[1], sys.argv[2]
    direction = sys.argv[3] if len(sys.argv) > 3 else "en2zh"
    max_items = int(sys.argv[4]) if len(sys.argv) > 4 else None
    key = open(KEY_FILE, encoding="utf-8").read().strip()
    translate_file(src, dst, key, direction, max_items)
