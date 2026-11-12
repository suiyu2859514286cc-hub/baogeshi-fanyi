# -*- coding: utf-8 -*-
"""保格式 docx 翻译：只改 word/document.xml 里的文字节点，样式链和表格结构不动。"""
import os
import re
import sys
import time
import zipfile

from lxml import etree

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from translate_xlsx import (call_api, protect, restore, clean_artifacts,
                            glossary_text, USAGE)

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
T = "{%s}t" % W
P = "{%s}p" % W


def para_text(p):
    return "".join(t.text or "" for t in p.iter(T))


def set_para_text(p, text):
    ts = list(p.iter(T))
    if not ts:
        return False
    ts[0].text = text
    ts[0].set(XML_SPACE, "preserve")
    for t in ts[1:]:
        t.text = ""
        t.set(XML_SPACE, "preserve")
    return True


def needs(text, direction):
    s = text.strip()
    if len(s) < 2:
        return False
    if direction == "zh2en":
        return bool(re.search(r"[\u4e00-\u9fff]", s))
    if re.search(r"[\u4e00-\u9fff]", s):
        return False
    return len(re.findall(r"[A-Za-z]{2,}", s)) >= 1


def translate_docx(src, dst, api_key, direction="en2zh", on_progress=None, bilingual=False):
    with zipfile.ZipFile(src) as z:
        names = z.namelist()
        parts = {n: z.read(n) for n in names}

    targets = [n for n in names
               if re.match(r"word/(document|header\d*|footer\d*|footnotes|endnotes)\.xml$", n)]
    if not targets:
        raise RuntimeError("这个文件里没找到正文部件")

    roots = {}
    all_paras = []      # [(part, paragraph, text)] 全段落，编号与导出对齐
    for n in targets:
        root = etree.fromstring(parts[n])
        roots[n] = root
        for p in root.iter(P):
            all_paras.append([n, p, para_text(p)])
    units = [u for u in all_paras if needs(u[2], direction)]
    print(f"docx 部件 {len(targets)} 个，段落 {len(all_paras)} 段，待译 {len(units)} 段")

    if not units:
        raise RuntimeError("没有找到需要翻译的文字段落")

    idx_of = {id(u): i for i, u in enumerate(all_paras)}
    payload, prot = [], {}
    for u in units:
        k = idx_of[id(u)]
        p_, spans = protect(u[2])
        prot[k] = spans
        payload.append(("k%d" % k, "", p_))

    results = {}
    last_err = ""
    B = 6
    total = (len(payload) + B - 1) // B
    for start in range(0, len(payload), B):
        batch = payload[start:start + B]
        for attempt in range(3):
            try:
                got = call_api(api_key, batch, direction)
                results.update(got)
                print(f"  批次 {start // B + 1}/{total} ok")
                break
            except Exception as ex:
                last_err = str(ex)
                low = last_err.lower()
                if ("401" in last_err or "403" in last_err or "402" in last_err
                        or "unauthorized" in low or "insufficient" in low
                        or "invalid api key" in low):
                    raise RuntimeError(
                        "API 拒绝了请求（Key 无效、已过期或余额不足）。\n原始信息：" + last_err[:160])
                print(f"  批次 {start // B + 1} 失败({attempt + 1}/3): {ex}")
                time.sleep(3)
        if on_progress:
            try:
                on_progress(min(start + B, len(payload)), len(payload))
            except Exception:
                pass
        time.sleep(0.25)

    changed, touched = 0, set()
    for u in units:
        k = idx_of[id(u)]
        val = results.get("k%d" % k)
        if not val:
            continue
        try:
            text = clean_artifacts(restore(val, prot[k]))
        except Exception:
            continue
        if not text.strip():
            continue
        if set_para_text(u[1], text):
            changed += 1
            touched.add(u[0])

    for n in touched:
        parts[n] = etree.tostring(roots[n], xml_declaration=True,
                                  encoding="UTF-8", standalone=True)

    total_units = max(1, len(units))
    if changed == 0 and total_units > 2:
        raise RuntimeError(
            "一段都没翻出来，整份文件没有被改动。\n原因：" +
            (last_err[:200] if last_err else "所有批次都失败了，请检查网络或 Key。"))
    if changed < total_units * 0.5:
        raise RuntimeError(
            "%d 段里有 %d 段没翻出来（超过一半），文件不可用。\n原因：%s"
            % (total_units, total_units - changed, (last_err[:200] if last_err else "未知")))
    print(f"改写 {changed} 段 / 共 {len(units)} 段")
    print(f"[用量] 调用 {USAGE['calls']} 次，输入 {USAGE['prompt']} tokens，输出 {USAGE['completion']} tokens")

    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as z:
        for n in names:
            zi = zipfile.ZipInfo(n, date_time=(2026, 1, 1, 0, 0, 0))
            zi.compress_type = zipfile.ZIP_DEFLATED
            z.writestr(zi, parts[n])
    print("写出:", dst)


# ---------- 供挑刺/逐条改使用 ----------
DOC_PARTS = r"word/(document|header\d*|footer\d*|footnotes|endnotes)\.xml$"


def _load(data):
    import io as _io
    with zipfile.ZipFile(_io.BytesIO(data)) as z:
        names = z.namelist()
        parts = {n: z.read(n) for n in names}
    return names, parts


def docx_extract(data):
    """按部件顺序返回所有段落的纯文本，编号与翻译时一致。"""
    import re as _re
    names, parts = _load(data)
    out = []
    for n in names:
        if not _re.match(DOC_PARTS, n):
            continue
        root = etree.fromstring(parts[n])
        for p in root.iter(P):
            out.append(para_text(p))
    return out


def docx_apply(data, edits):
    """把 edits（下标 -> 新文本）写回 docx，返回新字节。"""
    import io as _io
    import re as _re
    names, parts = _load(data)
    idx = 0
    touched = set()
    roots = {}
    for n in names:
        if not _re.match(DOC_PARTS, n):
            continue
        root = etree.fromstring(parts[n])
        roots[n] = root
        for p in root.iter(P):
            key = str(idx)
            if key in edits:
                if set_para_text(p, edits[key]):
                    touched.add(n)
            idx += 1
    for n in touched:
        parts[n] = etree.tostring(roots[n], xml_declaration=True,
                                  encoding="UTF-8", standalone=True)
    buf = _io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for n in names:
            zi = zipfile.ZipInfo(n, date_time=(2026, 1, 1, 0, 0, 0))
            zi.compress_type = zipfile.ZIP_DEFLATED
            z.writestr(zi, parts[n])
    return buf.getvalue()
