# -*- coding: utf-8 -*-
"""从新文件里自动发现术语候选：规则粗筛 → 本地模型判是不是术语 → 云端给建议译法。"""
import json
import os
import re
import sys
import urllib.request
import zipfile
from collections import Counter

from lxml import etree

BASE = r"C:\Users\ASUS\Documents\deepseek-harness\default-workspace"
GLOSS = r"D:\软件制作\工作台\保格式翻译\术语表\glossary.json"
import os as _os
_KEY_PATH = (_os.environ.get("BAOGESHI_KEY_FILE")
             or _os.path.join(_os.path.dirname(_os.path.abspath(__file__)), "..", "chat-key.txt"))
KEY = open(_KEY_PATH, encoding="utf-8").read().strip() if _os.path.exists(_KEY_PATH) else ""

NS = {'m': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

STOP = set("""the a an and or of to in for with on by as at from is are be been shall will must may can
all any other such which that this these those their its it he she they we you i not no
one two three four five six eight nine ten per each including complete provide supply install
install test works work item description unit qty rate total page following above below
but without left inside fully directly within between through during after before under over
shall be may not etc""".split())


def xlsx_text(p):
    z = zipfile.ZipFile(p)
    root = etree.fromstring(z.read('xl/sharedStrings.xml'))
    out = [''.join((t.text or '') for t in si.findall('.//m:t', NS)) for si in root.findall('m:si', NS)]
    z.close()
    return out


def docx_text(p):
    z = zipfile.ZipFile(p)
    out = []
    for n in z.namelist():
        if re.match(r"word/(document|header\d*|footer\d*)\.xml$", n):
            root = etree.fromstring(z.read(n))
            for para in root.iter('{%s}p' % W):
                out.append(''.join(t.text or '' for t in para.iter('{%s}t' % W)))
    z.close()
    return out


def get_texts(path):
    return docx_text(path) if path.lower().endswith(".docx") else xlsx_text(path)


def known_terms():
    s = set()
    if os.path.exists(GLOSS):
        for it in json.load(open(GLOSS, encoding="utf-8"))["items"]:
            for a in it["en"]:
                s.add(a.lower().strip())
    return s


def candidates(texts, top=120):
    big, tri = Counter(), Counter()
    for t in texts:
        ws = [w.lower() for w in re.findall(r"[A-Za-z][A-Za-z0-9/\-]*", t)]
        for i in range(len(ws) - 1):
            a, b = ws[i], ws[i + 1]
            if a in STOP or b in STOP or len(a) < 3 or len(b) < 3:
                continue
            big[a + " " + b] += 1
        for i in range(len(ws) - 2):
            a, b, c = ws[i], ws[i + 1], ws[i + 2]
            if a in STOP or b in STOP or c in STOP:
                continue
            tri[a + " " + b + " " + c] += 1
    known = known_terms()
    out = []
    for phrase, c in (big + tri).most_common(top):
        if c < 2 or phrase in known:
            continue
        if any(phrase in k for k in known if len(k) > 5):
            continue
        if phrase.split()[0] in STOP or phrase.split()[-1] in STOP:
            continue
        out.append((phrase, c))
    return out[:60]


def ask_local(prompt, model="qwen2.5:7b"):
    body = json.dumps({"model": model, "messages": [{"role": "user", "content": prompt}],
                       "stream": False, "format": "json",
                       "options": {"temperature": 0.1, "num_ctx": 4096}}).encode()
    req = urllib.request.Request("http://127.0.0.1:11434/api/chat", data=body,
                                 headers={"Content-Type": "application/json"})
    with urllib.request.urlopen(req, timeout=600) as r:
        return json.loads(json.loads(r.read().decode())["message"]["content"])


def ask_cloud(prompt, api_key=""):
    body = json.dumps({"model": "deepseek-chat", "messages": [{"role": "user", "content": prompt}],
                       "temperature": 0.1, "response_format": {"type": "json_object"}}).encode()
    req = urllib.request.Request("https://api.deepseek.com/chat/completions", data=body,
                                 headers={"Content-Type": "application/json",
                                          "Authorization": "Bearer " + (api_key or KEY)})
    with urllib.request.urlopen(req, timeout=180) as r:
        return json.loads(json.loads(r.read().decode())["choices"][0]["message"]["content"])


def suggest(paths, api_key=""):
    texts = []
    for p in paths:
        texts += get_texts(p)
    cands = candidates(texts)
    print("粗筛候选:", len(cands))

    phrases = [c[0] for c in cands]
    kept = []
    B = 40
    for start in range(0, len(phrases), B):
        batch = phrases[start:start + B]
        prompt = ("下面是工程量清单/合同里反复出现的英文短语。挑出其中**真正的工程技术术语**，"
                  "普通句子片段（but without、left inside 这类）不要。\n"
                  '只输出 JSON：{"terms":["术语"]}\n\n' + "\n".join(batch))
        try:
            kept += ask_local(prompt).get("terms", [])
        except Exception as ex:
            print("本地模型失败:", ex)
            kept += batch
    kept = [k for k in kept if k in dict(cands)]
    print("本地模型筛后:", len(kept))

    if not kept:
        return []
    prompt = ("把下面的工程术语翻成专业中文，只输出 JSON。\n"
              '{"items":[{"en":"英文","zh":"中文"}]}\n\n' + "\n".join(kept))
    try:
        items = ask_cloud(prompt, api_key).get("items", [])
    except Exception as ex:
        print("给译法失败:", ex)
        items = [{"en": k, "zh": ""} for k in kept]
    counts = dict(cands)
    for it in items:
        it["count"] = counts.get(it.get("en", ""), 0)
    return items


if __name__ == "__main__":
    ps = [r"D:\软件制作\工作台\保格式翻译\输入\3-4.专业设备分包工程量清单2.xlsx",
          r"D:\软件制作\工作台\保格式翻译\输入\3.训练场电分包工程量清单2.1.xlsx"]
    for it in suggest(ps):
        print(f"  {it.get('count'):>2}x  {it.get('en'):<32} → {it.get('zh')}")


def load_terms_file(path=GLOSS):
    if not os.path.exists(path):
        return []
    with open(path, encoding="utf-8") as f:
        return json.load(f).get("items", [])


def add_terms(new_items, path=GLOSS):
    """把用户确认过的术语写进 glossary.json（程序立刻生效）。"""
    data = {"source": "用户确认新增", "count": 0, "items": []}
    if os.path.exists(path):
        with open(path, encoding="utf-8") as f:
            data = json.load(f)
    existing = set()
    for it in data["items"]:
        for a in it.get("en", []):
            existing.add(a.lower().strip())
    added = 0
    for it in new_items:
        en = (it.get("en") or "").strip()
        zh = (it.get("zh") or "").strip()
        if not en or not zh or en.lower() in existing:
            continue
        data["items"].append({"en": [en], "zh": zh, "note": "自动发现，人工确认"})
        existing.add(en.lower())
        added += 1
    data["count"] = len(data["items"])
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=1)
    return added
