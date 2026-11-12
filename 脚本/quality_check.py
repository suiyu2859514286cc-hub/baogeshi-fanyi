# -*- coding: utf-8 -*-
"""挑刺：纯规则检查，不需要任何模型。
输入原文列表和译文列表（按 sharedString 下标对齐），输出可疑条目清单。"""
import json
import os
import re


def load_glossary_items():
    here = os.path.dirname(os.path.abspath(__file__))
    for p in (os.path.join(here, "..", "术语表", "glossary.json"),
              os.path.join(here, "glossary.json")):
        if os.path.exists(p):
            with open(p, encoding="utf-8") as f:
                return json.load(f).get("items", [])
    return []


NUM = re.compile(r"\d+(?:[.,]\d+)*")
WORD = re.compile(r"[A-Za-z]{2,}")


def is_label(s):
    """A、B、C、No、IP20 这种标记或规格，本来就不该翻。"""
    if len(s) <= 2:
        return True
    if not WORD.search(s):
        return True
    return False


def digits(s):
    return NUM.findall(s or "")


def check_quality(orig_list, trans_list, direction="en2zh", glossary=None):
    if glossary is None:
        glossary = load_glossary_items()
    issues = []
    n = min(len(orig_list), len(trans_list))
    for i in range(n):
        o = (orig_list[i] or "").strip()
        t = (trans_list[i] or "").strip()
        if not o:
            continue
        # 1) 没翻
        if is_label(o):
            continue
        # i) IP20 这种纯规格条目，不该翻
        if re.match(r"^[a-z]\)\s*[A-Z]{0,4}\d+", o):
            continue
        if o == t and re.search(r"[A-Za-z\u4e00-\u9fff]", o):
            issues.append({"i": i, "kind": "未译", "detail": "译文和原文一模一样"})
            continue
        # 2) 数字缺失
        od, td = set(digits(o)), set(digits(t))
        lost = [d for d in od if d not in td]
        if lost and len(o) > 8:
            issues.append({"i": i, "kind": "数字",
                           "detail": "原文有 " + "、".join(lost[:4]) + "，译文里没有"})
            continue
        # 3) 长度异常
        if direction == "en2zh":
            words = len(WORD.findall(o))
            if words >= 5 and len(t) < words * 0.55:
                issues.append({"i": i, "kind": "长度",
                               "detail": f"原文 {len(o)} 字符，译文只有 {len(t)}，可能漏译"})
                continue
        else:
            if len(o) >= 10 and len(t) > len(o) * 5:
                issues.append({"i": i, "kind": "长度",
                               "detail": f"原文 {len(o)} 字，译文 {len(t)} 字符，可能啰嗦"})
                continue
        # 4) 术语没按表
        hit = check_terms(o, t, glossary, direction)
        if hit:
            issues.append({"i": i, "kind": "术语", "detail": hit})
    return issues


def keywords(zh):
    """把「计列（含）……费用」「单联/双联/三联」拆成可命中的关键词。"""
    out = []
    for part in re.split(r"[/／…]+", zh):
        part = re.sub(r"（[^）]*）", "", part).strip()
        part = re.sub(r"\([^)]*\)", "", part).strip()
        if len(part) >= 2:
            out.append(part)
            if len(part) > 2:
                out.append(part[:2])   # 允许「人孔」对上「人孔井」这种近义短写
    return out


def check_terms(o, t, glossary, direction):
    if direction != "en2zh":
        return ""
    for it in glossary:
        zh = it.get("zh", "")
        if "保留" in zh:
            continue
        for a in it.get("en", []):
            if len(a) < 2:
                continue
            if not re.search(r"(?<![A-Za-z])" + re.escape(a) + r"(?![A-Za-z])", o, re.I):
                continue
            # 全大写缩写（ITEM/DB）只在原文也全大写时才当真
            if a.isupper() and a != a.upper():
                continue
            if a.isupper():
                m = re.search(r"(?<![A-Za-z])" + re.escape(a) + r"(?![A-Za-z])", o)
                if not m or m.group(0) != a:
                    continue
            # 译文里出现了原词本身（保留了英文缩写）也算过关
            if re.search(r"(?<![A-Za-z])" + re.escape(a) + r"(?![A-Za-z])", t, re.I):
                continue
            # TOTAL 在短语里是「合计」，只有作表头时才是「金额」
            if a.upper() == "TOTAL" and any(k in t for k in ("金额", "合计", "总计", "总额")):
                continue
            kws = keywords(zh)
            if any(k in t for k in kws):
                continue
            return f"原文有「{a}」，按表应出现「{zh}」，译文中没有"
    return ""
