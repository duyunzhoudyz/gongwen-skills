#!/usr/bin/env python3
# -*- coding: utf-8 -*-
r"""中立的 Word 公文版式审计脚本（GB/T 9704-2012 可测量条款）。

本脚本只依据 GB/T 9704-2012 的条款独立实现，不调用、不参考任何参赛方
（含本项目）的检查代码，用来对各方交付的 .docx 逐项打分（共 29 个检查项，见 ITEMS）。

命令行：
    python docx_audit.py <docx> [--layout 红头下行|上行|信函|纪要|普通材料|讲话]
                         [--ref 原文.txt] [--json] [--strict-font-names] [--roles]
可导入接口：
    audit_docx(path, layout=None, ref_text=None, strict_font_names=False) -> dict
    打不开的文件返回 {"error": "..."}（命令行退出码 2）；每个检查项独立 try/except，单项异常记 na。

实现要点
--------
1. 解析：zipfile + lxml 直接读 document/styles/settings/numbering/theme 与页眉页脚部件。
   生效属性顺序：docDefaults < 表格样式 < 段落样式链 < 编号级别 < 字符样式链 < 直接格式；
   东亚字体 eastAsiaTheme 优先于 eastAsia，主题字体取 a:ea，为空则取 a:font[script=Hans]。
   mc:AlternateContent 只取 mc:Choice（忽略 Fallback 的 VML 副本）；w:ins/w:moveTo 计入，w:del/w:moveFrom 不计。
2. 文字：自动编号渲染成序号文字拼在段落前（序号字体取编号级别 rPr，其上叠加段落标记 rPr）；
   “东亚字符”指汉字及中文标点（空白不计），按其生效的东亚字体统计；西文/数字不参与字体判定。
3. 段落角色（assign_roles，按文字特征，文档顺序）：
   - 标题：文首窗口（遇到“冒号结尾的主送机关行”或层次标题即停止）内第一个“居中、≥8 字、非红色、
     不是发文字号/份号/密级/紧急程度/日期/“…文件”结尾”的段落；多个候选时优先含“关于”或以文种词结尾者；
     找不到则取窗口内字号最大的居中非红段，再退到按文种词（通知/报告/函…）特征识别的非居中段；
     紧随其后的同格式居中短行并入多行标题。纪要版式的标题为空时，红色“纪要”标志代替标题（06 记 na）。
   - 版头区 = 标题之前的全部段落（含表格、文本框内段落），其中识别发文字号、签发人、份号、密级、紧急程度、
     发文机关标志（红色居中最大字号段；其次居中的“…文件/纪要”段）。发文字号误放在标题之后时兜底识别并在 notes 说明。
   - 标题后紧跟的居中短行 = “署名日期行/副标题”；其后以冒号结尾、≤80 字、非层次标题、非“如下：”引导句的
     连续段落 = 主送机关（讲话稿“同志们：”同理）。
   - 版记区 = 标题之后第一个以“抄送”开头（其次以“印发”结尾）的段落起至文末（含表格形式的版记）。
   - 附件标识 `^附件\s*\d*$`；附件说明 = 以“附件：”开头的段落及其后连续的短行（“2.”“3.”不当作三级标题）；
     成文日期 = 首个附件标识之前最后一个整段日期，署名 = 其上一短段（≤40 字、不以句号结尾）。
   - 层次标题按 一、/（一）/1./（1） 行首匹配；“标题成分”= 行首序号到第一个“。”或“：”（不含该标点），
     其后的同段接排文字计入 03a/03b 的正文字符统计，但 04/05 仍按“整段”统计（正文+层次标题段落）。
   - 其余为“正文”。表格内/文本框内段落不参与角色判定，只计入字体合规与保真文本。
4. 版式判断 detect_layout：发文字号（宽松正则，错误括号也识别）→红头下行，含签发人→上行；红色居中“纪要”首段→纪要；
   含“同志们”且无发文字号→讲话；其余→普通材料（信函无法自动判定，需用 --layout 指定）。
5. 红色分隔线/版记分隔线：识别段落边框（含样式继承）、DrawingML/VML 线条形状（排除 Fallback 副本）、表格边框，
   以及“行高≤2mm 且整行底纹填充”的表格行（用表格行画线的写法）；相邻段落的下/上边框视为同一条线。
"""
from __future__ import annotations

import argparse
import difflib
import json
import os
import posixpath
import re
import sys
import zipfile
from collections import Counter

from lxml import etree

# ====================================================================== 命名空间
NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "wps": "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
    "v": "urn:schemas-microsoft-com:vml",
    "o": "urn:schemas-microsoft-com:office:office",
}
W_NS = NS["w"]
R_NS = NS["r"]


def qn(name):
    """'w:p' -> '{ns}p'"""
    p, t = name.split(":")
    return "{%s}%s" % (NS[p], t)


T_P = qn("w:p")
T_R = qn("w:r")
T_TBL = qn("w:tbl")
T_TR = qn("w:tr")
T_TC = qn("w:tc")
T_SDT = qn("w:sdt")
T_SDTC = qn("w:sdtContent")
T_PPR = qn("w:pPr")
T_RPR = qn("w:rPr")
T_SECT = qn("w:sectPr")
T_TXBX = qn("w:txbxContent")
T_ALT = qn("mc:AlternateContent")
T_CHOICE = qn("mc:Choice")
T_FALLBACK = qn("mc:Fallback")
T_INS = qn("w:ins")
T_DEL = qn("w:del")
T_MOVEFROM = qn("w:moveFrom")
T_MOVETO = qn("w:moveTo")
T_HYPER = qn("w:hyperlink")
T_FLDSIMPLE = qn("w:fldSimple")


def lname(el):
    t = el.tag
    return t.rsplit("}", 1)[-1] if isinstance(t, str) else ""


def wa(el, name, default=None):
    """读取 w: 命名空间属性（容错：无命名空间的同名属性也接受）。"""
    if el is None:
        return default
    v = el.get("{%s}%s" % (W_NS, name))
    if v is None:
        v = el.get(name)
    return default if v is None else v


def is_on(el):
    """开关属性：无元素 -> None；w:val 为 0/false/off 表示关，其余（含缺省）为开。"""
    if el is None:
        return None
    v = wa(el, "val")
    if v is None:
        return True
    return str(v).strip().lower() not in ("0", "false", "off")


def expand(el):
    """遍历子元素；mc:AlternateContent 只取 mc:Choice（无 Choice 时才取 Fallback），
    忽略重复的 mc:Fallback（VML 副本）。"""
    for ch in el:
        if not isinstance(ch.tag, str):
            continue
        if ch.tag == T_ALT:
            pick = None
            for c in ch:
                if c.tag == T_CHOICE:
                    pick = c
                    break
            if pick is None:
                for c in ch:
                    if c.tag == T_FALLBACK:
                        pick = c
                        break
            if pick is not None:
                yield from expand(pick)
        elif ch.tag == T_FALLBACK:
            continue
        else:
            yield ch


def walk_no_fallback(el):
    """深度优先遍历所有后代元素，跳过 mc:Fallback 子树（不进入 w:txbxContent 之外的限制）。"""
    stack = [el]
    while stack:
        cur = stack.pop()
        for ch in reversed(list(cur)):
            if not isinstance(ch.tag, str) or ch.tag == T_FALLBACK:
                continue
            yield ch
            stack.append(ch)


# ====================================================================== 度量与换算
def meas_to_twips(s):
    """解析 ST_TwipsMeasure / ST_SignedTwipsMeasure（含 mm/cm/in/pt/pc/pi 单位）-> twips(float)。"""
    if s is None:
        return None
    s = str(s).strip()
    m = re.match(r"^(-?\d+(?:\.\d+)?)\s*(mm|cm|in|pt|pc|pi)?$", s)
    if not m:
        return None
    v = float(m.group(1))
    u = m.group(2)
    if u is None:
        return v
    return v * {"mm": 56.6929, "cm": 566.929, "in": 1440.0, "pt": 20.0, "pc": 240.0, "pi": 240.0}[u]


def tw2mm(tw):
    return tw / 56.6929


def tw2pt(tw):
    return tw / 20.0


def fmt_pt(x):
    return ("%g" % round(x, 2)) + "pt"


def fmt_mm(x):
    return ("%.1f" % x).rstrip("0").rstrip(".") + "mm"


# ====================================================================== 颜色
NAMED_COLORS = {
    "red": (255, 0, 0), "black": (0, 0, 0), "white": (255, 255, 255), "blue": (0, 0, 255),
    "green": (0, 128, 0), "yellow": (255, 255, 0), "gray": (128, 128, 128), "grey": (128, 128, 128),
    "silver": (192, 192, 192), "maroon": (128, 0, 0), "navy": (0, 0, 128), "purple": (128, 0, 128),
    "orange": (255, 165, 0), "darkred": (139, 0, 0), "crimson": (220, 20, 60),
    "firebrick": (178, 34, 34), "brown": (165, 42, 42),
}


def parse_color(v):
    """'FF0000' / '#FF0000' / 'red' -> (r,g,b)；auto/无法解析 -> None。"""
    if v is None:
        return None
    v = str(v).strip()
    if v:
        v = v.split()[0]  # VML 写法如 "#ff0000 [3204]"、"red [3204]"，只取第一个记号
    if not v or v.lower() == "auto":
        return None
    h = v.lstrip("#")
    if re.fullmatch(r"[0-9a-fA-F]{6}", h):
        return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))
    if re.fullmatch(r"[0-9a-fA-F]{3}", h) and v.startswith("#"):
        return tuple(int(c * 2, 16) for c in h)
    m = re.match(r"^(\w+)", v)
    if m and m.group(1).lower() in NAMED_COLORS:
        return NAMED_COLORS[m.group(1).lower()]
    return None


def is_red(rgb):
    """近似红色：R>=0xB0 且 G<=0x60 且 B<=0x60。"""
    return rgb is not None and rgb[0] >= 0xB0 and rgb[1] <= 0x60 and rgb[2] <= 0x60


def is_black(rgb):
    """黑色：None(自动) 或 R,G,B 均 <=0x30。"""
    return rgb is None or (rgb[0] <= 0x30 and rgb[1] <= 0x30 and rgb[2] <= 0x30)


# ====================================================================== 字符分类
def is_ea_char(ch):
    """汉字及中文标点等“东亚字符”（按东亚字体统计）。空白不计。"""
    o = ord(ch)
    if ch.isspace():
        return False
    return (
        0x4E00 <= o <= 0x9FFF or 0x3400 <= o <= 0x4DBF or 0xF900 <= o <= 0xFAFF
        or 0x3000 <= o <= 0x303F or 0xFF00 <= o <= 0xFFEF or 0x2010 <= o <= 0x2015
        or 0x2018 <= o <= 0x201F or o == 0x2026 or o == 0x00B7 or o == 0x203B
        or 0x2460 <= o <= 0x24FF or 0x3040 <= o <= 0x30FF or 0x20000 <= o <= 0x2FA1F
    )


def compact(s):
    """去掉全部空白字符（含全角空格、制表符、换行、不间断空格）。"""
    return re.sub(r"\s+", "", s or "")


# ====================================================================== 主题
class Theme:
    def __init__(self, root):
        self.fonts = {"major": {}, "minor": {}}
        self.colors = {}
        if root is None:
            return
        for which in ("major", "minor"):
            node = root.find(".//a:%sFont" % which, NS)
            if node is None:
                continue
            d = {"latin": "", "ea": "", "cs": "", "scripts": {}}
            for ch in node:
                n = lname(ch)
                if n in ("latin", "ea", "cs"):
                    d[n] = ch.get("typeface", "") or ""
                elif n == "font":
                    d["scripts"][ch.get("script", "")] = ch.get("typeface", "") or ""
            self.fonts[which] = d
        cs = root.find(".//a:clrScheme", NS)
        if cs is not None:
            for ch in cs:
                nm = lname(ch)
                val = None
                for c in ch:
                    if lname(c) == "srgbClr":
                        val = c.get("val")
                    elif lname(c) == "sysClr":
                        val = c.get("lastClr") or c.get("val")
                if val:
                    self.colors[nm] = parse_color(val) if re.fullmatch(r"[0-9a-fA-F]{6}", val) else None

    def font_for(self, theme_name):
        """theme_name 形如 minorEastAsia / majorHAnsi / minorAscii / minorBidi。"""
        which = "major" if str(theme_name).startswith("major") else "minor"
        d = self.fonts.get(which) or {}
        if "EastAsia" in theme_name:
            v = d.get("ea") or ""
            if not v:
                sc = d.get("scripts", {})
                v = sc.get("Hans") or sc.get("Hant") or ""
            return v or None
        if "Bidi" in theme_name:
            return d.get("cs") or None
        return d.get("latin") or None

    def color_for(self, theme_color):
        m = {"text1": "dk1", "background1": "lt1", "text2": "dk2", "background2": "lt2",
             "dark1": "dk1", "light1": "lt1", "dark2": "dk2", "light2": "lt2",
             "tx1": "dk1", "bg1": "lt1", "tx2": "dk2", "bg2": "lt2",
             "hyperlink": "hlink", "followedHyperlink": "folHlink"}
        key = m.get(theme_color, theme_color)
        return self.colors.get(key)


# ====================================================================== 包读取
class Package:
    """docx 包：读取各部件，解析关系。打不开则抛异常（由调用方转成 {"error": ...}）。"""

    def __init__(self, path):
        self.path = path
        self.zf = zipfile.ZipFile(path)
        try:
            self.names = set(self.zf.namelist())
            self._cache = {}
            self.main = self._find_main()
            self.doc = self.xml(self.main)
            if self.doc is None:
                if self.main in self.names:
                    raise ValueError("%s 不是合法的 XML（文件已损坏）" % self.main)
                raise ValueError("缺少 word/document.xml（不是 Word 文档）")
            self.main_dir = posixpath.dirname(self.main)
            self.rels = self.rels_of(self.main)
            self.styles_root = self.xml(self._by_type("/styles", "word/styles.xml"))
            self.settings_root = self.xml(self._by_type("/settings", "word/settings.xml"))
            self.numbering_root = self.xml(self._by_type("/numbering", "word/numbering.xml"))
            self.theme_root = self.xml(self._by_type("/theme", "word/theme/theme1.xml"))
        except Exception:
            self.zf.close()
            raise

    def _find_main(self):
        rels = self.rels_of("")
        for _, (typ, tgt, _m) in rels.items():
            if typ.endswith("/officeDocument"):
                return tgt
        return "word/document.xml"

    def _by_type(self, suffix, default):
        for _, (typ, tgt, _m) in self.rels.items():
            if typ.endswith(suffix):
                return tgt
        return default

    def rels_of(self, part):
        """返回 {rId: (type, 绝对部件路径, mode)}。part 为空串表示包根关系。"""
        if part:
            d, f = posixpath.split(part)
            rp = posixpath.join(d, "_rels", f + ".rels")
        else:
            d, rp = "", "_rels/.rels"
        out = {}
        root = self.xml(rp)
        if root is None:
            return out
        for ch in root:
            if lname(ch) != "Relationship":
                continue
            tgt = ch.get("Target", "")
            mode = ch.get("TargetMode", "")
            if mode != "External":
                if tgt.startswith("/"):
                    tgt = tgt.lstrip("/")
                else:
                    tgt = posixpath.normpath(posixpath.join(d, tgt))
            out[ch.get("Id")] = (ch.get("Type", ""), tgt, mode)
        return out

    def xml(self, name):
        if name is None:
            return None
        if name in self._cache:
            return self._cache[name]
        root = None
        if name in self.names:
            try:
                parser = etree.XMLParser(resolve_entities=False, no_network=True, huge_tree=True)
                root = etree.fromstring(self.zf.read(name), parser)
            except Exception:
                root = None
        self._cache[name] = root
        return root


# ====================================================================== 属性解析
def parse_rpr(rpr):
    """w:rPr -> dict（只保留审计用到的属性）。"""
    d = {}
    if rpr is None:
        return d
    for ch in rpr:
        n = lname(ch)
        if n == "rFonts":
            d["rFonts"] = {lname_attr(k): v for k, v in ch.attrib.items()}
        elif n == "sz":
            try:
                d["sz"] = int(float(wa(ch, "val")))
            except Exception:
                pass
        elif n == "b":
            d["b"] = is_on(ch)
        elif n == "i":
            d["i"] = is_on(ch)
        elif n == "color":
            d["color"] = {"val": wa(ch, "val"), "themeColor": wa(ch, "themeColor")}
        elif n == "rStyle":
            d["rStyle"] = wa(ch, "val")
        elif n == "vanish":
            d["vanish"] = is_on(ch)
        elif n == "lang":
            d["lang"] = {lname_attr(k): v for k, v in ch.attrib.items()}
    return d


def lname_attr(k):
    return k.rsplit("}", 1)[-1] if isinstance(k, str) else k


def parse_border(el):
    return {
        "val": wa(el, "val"), "sz": wa(el, "sz"), "color": wa(el, "color"),
        "space": wa(el, "space"), "themeColor": wa(el, "themeColor"),
    }


def parse_ppr(ppr):
    """w:pPr -> dict。ind 键统一：start->left, end->right，值保持原字符串。"""
    d = {}
    if ppr is None:
        return d
    for ch in ppr:
        n = lname(ch)
        if n == "pStyle":
            d["pStyle"] = wa(ch, "val")
        elif n == "jc":
            d["jc"] = wa(ch, "val")
        elif n == "ind":
            ind = {lname_attr(k): v for k, v in ch.attrib.items()}
            if "start" in ind and "left" not in ind:
                ind["left"] = ind.pop("start")
            if "end" in ind and "right" not in ind:
                ind["right"] = ind.pop("end")
            if "startChars" in ind and "leftChars" not in ind:
                ind["leftChars"] = ind.pop("startChars")
            if "endChars" in ind and "rightChars" not in ind:
                ind["rightChars"] = ind.pop("endChars")
            d["ind"] = ind
        elif n == "spacing":
            d["spacing"] = {lname_attr(k): v for k, v in ch.attrib.items()}
        elif n == "pageBreakBefore":
            d["pageBreakBefore"] = is_on(ch)
        elif n == "snapToGrid":
            d["snapToGrid"] = is_on(ch)
        elif n == "numPr":
            np_ = {}
            for c in ch:
                if lname(c) == "numId":
                    np_["numId"] = wa(c, "val")
                elif lname(c) == "ilvl":
                    np_["ilvl"] = wa(c, "val")
            d["numPr"] = np_
        elif n == "pBdr":
            d["pBdr"] = {lname(c): parse_border(c) for c in ch if lname(c) in ("top", "bottom", "left", "right", "between")}
        elif n == "tabs":
            d["tabs"] = [(wa(c, "val"), wa(c, "pos")) for c in ch if lname(c) == "tab"]
        elif n == "framePr":
            d["framePr"] = {lname_attr(k): v for k, v in ch.attrib.items()}
        elif n == "shd":
            d["shd"] = {"val": wa(ch, "val"), "fill": wa(ch, "fill"), "color": wa(ch, "color")}
        elif n == "rPr":
            d["markRPr"] = parse_rpr(ch)
        elif n == "sectPr":
            d["sectPr"] = ch
    return d


_IND_FIRST = ("firstLine", "hanging", "firstLineChars", "hangingChars")


def merge_ppr(base, over):
    """段落属性合并：over 覆盖 base；ind/spacing 按属性合并，pBdr 按边合并，tabs 累加。"""
    if not over:
        return base
    out = dict(base)
    for k, v in over.items():
        if k == "ind":
            cur = dict(out.get("ind", {}))
            if any(x in v for x in _IND_FIRST):
                for x in _IND_FIRST:
                    cur.pop(x, None)
            cur.update(v)
            out["ind"] = cur
        elif k == "spacing":
            cur = dict(out.get("spacing", {}))
            cur.update(v)
            if "line" in v and "lineRule" not in v:
                cur["lineRule"] = "auto"
            out["spacing"] = cur
        elif k == "pBdr":
            cur = dict(out.get("pBdr", {}))
            cur.update(v)
            out["pBdr"] = cur
        elif k == "tabs":
            cur = list(out.get("tabs", []))
            for val, pos in v:
                if val == "clear":
                    cur = [t for t in cur if t[1] != pos]
                else:
                    cur = [t for t in cur if t[1] != pos] + [(val, pos)]
            out["tabs"] = cur
        elif k == "numPr":
            cur = dict(out.get("numPr", {}))
            cur.update(v)
            out["numPr"] = cur
        else:
            out[k] = v
    return out


def merge_rpr(base, over):
    """字符属性合并：rFonts 按属性合并（显式字体名覆盖继承的主题字体）。"""
    if not over:
        return base
    out = dict(base)
    for k, v in over.items():
        if k == "rFonts":
            cur = dict(out.get("rFonts", {}))
            for slot, theme in (("ascii", "asciiTheme"), ("hAnsi", "hAnsiTheme"),
                                ("eastAsia", "eastAsiaTheme"), ("cs", "cstheme")):
                if slot in v and theme not in v:
                    cur.pop(theme, None)
                if theme in v and slot not in v:
                    cur.pop(slot, None)
            cur.update(v)
            out["rFonts"] = cur
        elif k == "color":
            out["color"] = v
        else:
            out[k] = v
    return out



# ====================================================================== 样式
class StyleDef:
    __slots__ = ("id", "type", "name", "based", "ppr", "rpr", "default", "tbl_borders")

    def __init__(self):
        self.id = None
        self.type = None
        self.name = ""
        self.based = None
        self.ppr = {}
        self.rpr = {}
        self.default = False
        self.tbl_borders = {}


class Styles:
    """样式表：docDefaults + 样式链 解析。生效顺序：
    docDefaults < 表格样式 < 段落样式链 < (编号级别) < 字符样式链 < 直接格式。"""

    def __init__(self, root, theme):
        self.theme = theme
        self.by_id = {}
        self.doc_rpr = {}
        self.doc_ppr = {}
        self.default_par = None
        self._pc = {}
        self._rc = {}
        self._bp = {}
        self._br = {}
        if root is None:
            return
        dd = root.find(qn("w:docDefaults"))
        if dd is not None:
            r = dd.find(qn("w:rPrDefault"))
            if r is not None:
                self.doc_rpr = parse_rpr(r.find(T_RPR))
            p = dd.find(qn("w:pPrDefault"))
            if p is not None:
                self.doc_ppr = parse_ppr(p.find(T_PPR))
        for st in root.findall(qn("w:style")):
            sd = StyleDef()
            sd.id = wa(st, "styleId")
            sd.type = wa(st, "type")
            nm = st.find(qn("w:name"))
            sd.name = wa(nm, "val", "") if nm is not None else ""
            bo = st.find(qn("w:basedOn"))
            sd.based = wa(bo, "val") if bo is not None else None
            sd.ppr = parse_ppr(st.find(T_PPR))
            sd.rpr = parse_rpr(st.find(T_RPR))
            sd.default = str(wa(st, "default", "0")).lower() in ("1", "true", "on")
            tp = st.find(qn("w:tblPr"))
            if tp is not None:
                tb = tp.find(qn("w:tblBorders"))
                if tb is not None:
                    sd.tbl_borders = {lname(c): parse_border(c) for c in tb}
            if sd.id:
                self.by_id[sd.id] = sd
            if sd.type == "paragraph" and sd.default and self.default_par is None:
                self.default_par = sd.id

    # ---- 样式链
    def chain(self, sid):
        out, seen = [], set()
        while sid and sid in self.by_id and sid not in seen:
            seen.add(sid)
            sd = self.by_id[sid]
            out.append(sd)
            sid = sd.based
        out.reverse()
        return out

    def style_ppr(self, sid):
        if sid not in self._pc:
            d = {}
            for sd in self.chain(sid):
                d = merge_ppr(d, sd.ppr)
            self._pc[sid] = d
        return self._pc[sid]

    def style_rpr(self, sid):
        if sid not in self._rc:
            d = {}
            for sd in self.chain(sid):
                d = merge_rpr(d, sd.rpr)
            self._rc[sid] = d
        return self._rc[sid]

    def table_borders(self, sid):
        """表格样式链上的 tblBorders（派生覆盖基）。"""
        out = {}
        for sd in self.chain(sid):
            out.update(sd.tbl_borders)
        return out

    # ---- 生效属性
    def base_ppr(self, psid, tbl_sid=None):
        psid = psid if psid in self.by_id else self.default_par
        key = (psid, tbl_sid)
        if key not in self._bp:
            d = dict(self.doc_ppr)
            if tbl_sid:
                d = merge_ppr(d, self.style_ppr(tbl_sid))
            d = merge_ppr(d, self.style_ppr(psid))
            self._bp[key] = d
        return self._bp[key]

    def base_rpr(self, psid, tbl_sid=None):
        psid = psid if psid in self.by_id else self.default_par
        key = (psid, tbl_sid)
        if key not in self._br:
            d = dict(self.doc_rpr)
            if tbl_sid:
                d = merge_rpr(d, self.style_rpr(tbl_sid))
            d = merge_rpr(d, self.style_rpr(psid))
            self._br[key] = d
        return self._br[key]

    def eff_rpr(self, psid, direct, tbl_sid=None, extra=None):
        """段落样式 + 字符样式 + 直接格式（extra 为编号级别 rPr 等额外层，位于字符样式之前）。"""
        base = self.base_rpr(psid, tbl_sid)
        if extra:
            base = merge_rpr(base, extra)
        rs = direct.get("rStyle") if direct else None
        if rs:
            base = merge_rpr(base, self.style_rpr(rs))
        return merge_rpr(base, direct or {})

    # ---- 取值
    def fonts(self, eff):
        """返回 (东亚字体名, 西文字体名)；东亚字体解析不到时返回 None。"""
        rf = eff.get("rFonts", {})
        ea = None
        t = rf.get("eastAsiaTheme")
        if t:
            ea = self.theme.font_for(t)
        if not ea:
            ea = rf.get("eastAsia") or None
        asc = None
        t = rf.get("asciiTheme")
        if t:
            asc = self.theme.font_for(t)
        if not asc:
            asc = rf.get("ascii") or rf.get("hAnsi") or None
        return ea, asc

    @staticmethod
    def half_points(eff):
        sz = eff.get("sz")
        return sz if sz is not None else 20  # Word 缺省 10pt

    def color_rgb(self, eff):
        c = eff.get("color")
        if not c:
            return None
        tc = c.get("themeColor")
        if tc:
            rgb = self.theme.color_for(tc)
            if rgb is not None:
                return rgb
        return parse_color(c.get("val"))


# ====================================================================== 编号
_BULLET_MAP = {
    "": "•", "": "▪", "": "➢", "": "❖", "": "✓",
    "": "●", "": "■", "": "◆", "": "•",
}


def cn_count(n, upper=False):
    """1..9999 的中文数字（十、十一、二十、一百零一）。"""
    digits = "零壹贰叁肆伍陆柒捌玖" if upper else "零一二三四五六七八九"
    unit = ("", "拾", "佰", "仟") if upper else ("", "十", "百", "千")
    if n <= 0:
        return digits[0] if n == 0 else str(n)
    if n >= 10000:
        return "".join(digits[int(c)] for c in str(n))
    parts = [(n // 1000) % 10, (n // 100) % 10, (n // 10) % 10, n % 10]
    s, started, zero = "", False, False
    for i, d in enumerate(parts):
        pos = 3 - i
        if d == 0:
            if started:
                zero = True
            continue
        if zero:
            s += digits[0]
            zero = False
        if d == 1 and pos == 1 and not started:
            s += unit[1]
        else:
            s += digits[d] + unit[pos]
        started = True
    return s


def roman(n, lower=False):
    vals = [(1000, "M"), (900, "CM"), (500, "D"), (400, "CD"), (100, "C"), (90, "XC"),
            (50, "L"), (40, "XL"), (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I")]
    s = ""
    for v, r in vals:
        while n >= v:
            s += r
            n -= v
    return s.lower() if lower else s


def fmt_num(n, fmt):
    """按 numFmt 格式化计数值。"""
    fmt = fmt or "decimal"
    if fmt == "decimal":
        return str(n)
    if fmt == "decimalZero":
        return "%02d" % n
    if fmt in ("upperLetter", "lowerLetter"):
        c = chr(65 + (n - 1) % 26) * ((n - 1) // 26 + 1) if n > 0 else str(n)
        return c if fmt == "upperLetter" else c.lower()
    if fmt in ("upperRoman", "lowerRoman"):
        return roman(n, fmt == "lowerRoman") if n > 0 else str(n)
    if fmt in ("chineseCounting", "chineseCountingThousand", "japaneseCounting", "taiwaneseCounting"):
        return cn_count(n)
    if fmt in ("chineseLegalSimplified", "japaneseLegal", "taiwaneseCountingThousand"):
        return cn_count(n, upper=True)
    if fmt == "ideographDigital":
        return "".join("〇一二三四五六七八九"[int(c)] for c in str(n))
    if fmt in ("ideographTraditional", "ideographLegalTraditional"):
        s = "甲乙丙丁戊己庚辛壬癸"
        return s[n - 1] if 1 <= n <= 10 else str(n)
    if fmt == "decimalFullWidth" or fmt == "decimalFullWidth2":
        return "".join(chr(0xFF10 + int(c)) for c in str(n))
    if fmt == "decimalEnclosedCircle":
        return chr(0x2460 + n - 1) if 1 <= n <= 20 else "(%d)" % n
    if fmt == "decimalEnclosedParen":
        return chr(0x2474 + n - 1) if 1 <= n <= 20 else "(%d)" % n
    if fmt == "none":
        return ""
    return str(n)


class Numbering:
    """自动编号渲染：按 abstractNum 分别计数，高级别递增时低级别重置。"""

    def __init__(self, root):
        self.abstract = {}
        self.nums = {}
        self.style_link = {}
        self.counters = {}
        self.started = set()
        if root is None:
            return
        for an in root.findall(qn("w:abstractNum")):
            aid = wa(an, "abstractNumId")
            levels = {}
            for lv in an.findall(qn("w:lvl")):
                il = wa(lv, "ilvl")
                if il is None:
                    continue
                d = self._parse_lvl(lv)
                levels[int(il)] = d
                if d["pstyle"]:
                    self.style_link[d["pstyle"]] = (aid, int(il))
            self.abstract[aid] = levels
        for nm in root.findall(qn("w:num")):
            nid = wa(nm, "numId")
            an = nm.find(qn("w:abstractNumId"))
            n = {"abs": wa(an, "val") if an is not None else None, "lvl_over": {}, "start_over": {}}
            for ov in nm.findall(qn("w:lvlOverride")):
                il = wa(ov, "ilvl")
                if il is None:
                    continue
                il = int(il)
                so = ov.find(qn("w:startOverride"))
                if so is not None:
                    try:
                        n["start_over"][il] = int(wa(so, "val"))
                    except Exception:
                        pass
                lv = ov.find(qn("w:lvl"))
                if lv is not None:
                    n["lvl_over"][il] = self._parse_lvl(lv)
            self.nums[nid] = n

    @staticmethod
    def _parse_lvl(lv):
        d = {"start": 1, "fmt": "decimal", "text": "", "rpr": {}, "ppr": {}, "suff": "tab",
             "pstyle": None, "isLgl": False}
        for ch in lv:
            n = lname(ch)
            v = wa(ch, "val")
            if n == "start":
                try:
                    d["start"] = int(v)
                except Exception:
                    pass
            elif n == "numFmt":
                d["fmt"] = v
            elif n == "lvlText":
                d["text"] = v or ""
            elif n == "pPr":
                d["ppr"] = parse_ppr(ch)
            elif n == "rPr":
                d["rpr"] = parse_rpr(ch)
            elif n == "suff":
                d["suff"] = v
            elif n == "pStyle":
                d["pstyle"] = v
            elif n == "isLgl":
                d["isLgl"] = is_on(ch) is not False
        return d

    def level(self, num_id, ilvl):
        n = self.nums.get(str(num_id))
        if not n:
            return None
        if ilvl in n["lvl_over"]:
            return n["lvl_over"][ilvl]
        return self.abstract.get(n["abs"], {}).get(ilvl)

    def link_for_style(self, style_id):
        """段落样式经 abstractNum 级别 pStyle 关联到的 (abstractNumId, ilvl)。"""
        return self.style_link.get(style_id)

    def num_for_abstract(self, abs_id):
        for nid, n in self.nums.items():
            if n["abs"] == abs_id:
                return nid
        return None

    def next_label(self, num_id, ilvl):
        """推进计数并返回 (label, level_dict)；该编号不存在返回 None。"""
        n = self.nums.get(str(num_id))
        if not n:
            return None
        lvl = self.level(num_id, ilvl)
        if lvl is None:
            return None
        cnt = self.counters.setdefault(n["abs"], {})
        key = str(num_id)
        if key not in self.started:
            self.started.add(key)
            for il, ov in n["start_over"].items():
                cnt[il] = ov - 1
        if cnt.get(ilvl) is None:
            cnt[ilvl] = lvl["start"]
        else:
            cnt[ilvl] += 1
        for deeper in [k for k in cnt if k > ilvl]:
            cnt[deeper] = None
        text = lvl["text"]
        if lvl["fmt"] == "bullet":
            return "".join(_BULLET_MAP.get(c, "•") if "" <= c <= "" else c for c in text), lvl

        def repl(m):
            k = int(m.group(1)) - 1
            sub = self.level(num_id, k)
            if sub is None:
                return ""
            val = cnt.get(k)
            if val is None:
                val = sub["start"]
            f = "decimal" if (lvl["isLgl"] and k != ilvl) else sub["fmt"]
            return fmt_num(val, f)

        return re.sub(r"%([1-9])", repl, text), lvl



# ====================================================================== 文档模型
T_T = qn("w:t")
T_TAB = qn("w:tab")
T_PTAB = qn("w:ptab")
T_BR = qn("w:br")
T_CR = qn("w:cr")
T_NBH = qn("w:noBreakHyphen")
T_FLDCHAR = qn("w:fldChar")
T_INSTR = qn("w:instrText")
T_DRAWING = qn("w:drawing")
T_PICT = qn("w:pict")
T_OBJECT = qn("w:object")
_TRANSPARENT = {qn("w:smartTag"), qn("w:customXml"), qn("w:dir"), qn("w:bdo"), qn("w:ins"), qn("w:moveTo"),
                qn("w:hyperlink")}


class Seg:
    """一段同格式文字。sz 为半磅；color 为 (r,g,b) 或 None(自动)。"""
    __slots__ = ("text", "ea", "asc", "sz", "bold", "color", "kind")

    def __init__(self, text, props, kind="text"):
        self.text = text
        self.ea, self.asc, self.sz, self.bold, self.color = props
        self.kind = kind


class Para:
    def __init__(self):
        self.idx = -1
        self.el = None
        self.in_table = False
        self.in_txbx = False
        self.tbl = None          # (表序号, 行, 列, 本行单元格数)
        self.tbl_sid = None
        self.sect = 0
        self.ppr = {}
        self.psid = None
        self.segs = []
        self.fields = []         # [{'kind','instr','start','end','props'}]
        self.pb = []             # 页分隔符在 text 中的字符偏移
        self.text = ""
        self.sectpr = None
        self.role = ""
        self.area = "body"
        self.flags = set()

    # ---- 派生
    @property
    def stext(self):
        return self.text.strip(" \t\r\n　 ")

    @property
    def ctext(self):
        return compact(self.text)

    @property
    def empty(self):
        return not self.ctext

    @property
    def raw_text(self):
        return "".join(s.text for s in self.segs if s.kind != "num")

    def jc(self):
        """有效对齐：框架段落以 framePr 的 xAlign 为准（视觉位置），否则取 jc；start/end 视同 left/right。"""
        fp = self.ppr.get("framePr")
        if fp and fp.get("xAlign") in ("left", "center", "right"):
            return fp["xAlign"]
        v = self.ppr.get("jc")
        if v in ("end", "right"):
            return "right"
        if v in (None, "", "start", "left"):
            return "left"
        return v

    def chars(self):
        """可见字符 [(ch, seg)]（不含空白）。"""
        out = []
        for s in self.segs:
            for ch in s.text:
                if not ch.isspace():
                    out.append((ch, s))
        return out

    def all_chars(self):
        out = []
        for s in self.segs:
            for ch in s.text:
                out.append((ch, s))
        return out

    def max_pt(self):
        v = [s.sz for _, s in self.chars()]
        return (max(v) / 2.0) if v else 0.0

    def dom_pt(self):
        """字符数最多的字号(pt)。"""
        c = Counter(s.sz for _, s in self.chars())
        return (c.most_common(1)[0][0] / 2.0) if c else 0.0

    def red_ratio(self):
        ch = self.chars()
        if not ch:
            return 0.0
        return sum(1 for _, s in ch if is_red(s.color)) / len(ch)

    def is_red(self):
        return self.red_ratio() >= 0.8

    def ind(self):
        return self.ppr.get("ind", {})

    def short(self, n=24):
        return self.stext[:n]


def iter_txbx(el):
    """依次产出元素内的 w:txbxContent（不进入 mc:Fallback，不重复嵌套）。"""
    for ch in expand(el):
        if ch.tag == T_TXBX:
            yield ch
        else:
            yield from iter_txbx(ch)


class Walker:
    """按文档顺序遍历容器，产出 Para 列表（含表格、文本框、内容控件内段落）。"""

    def __init__(self, doc, area="body", numbering=None):
        self.doc = doc
        self.styles = doc.styles
        self.numbering = numbering if numbering is not None else doc.numbering
        self.area = area
        self.paras = []
        self.tables = []
        self.sections = []
        self.sect_idx = 0
        self.tbl_counter = 0
        self._memo = {}

    # ---- 块级遍历
    def walk(self, container, in_table=False, in_txbx=False, tbl_ctx=None, tbl_sid=None):
        for ch in expand(container):
            t = ch.tag
            if t == T_P:
                self.add_para(ch, in_table, in_txbx, tbl_ctx, tbl_sid)
            elif t == T_TBL:
                self.walk_table(ch, in_table, in_txbx)
            elif t == T_SDT:
                sc = ch.find(T_SDTC)
                if sc is not None:
                    self.walk(sc, in_table, in_txbx, tbl_ctx, tbl_sid)
            elif t in _TRANSPARENT:
                self.walk(ch, in_table, in_txbx, tbl_ctx, tbl_sid)
            elif t == T_SECT and self.area == "body" and container is self.doc.body:
                self.sections.append(ch)

    def walk_table(self, tbl, in_table, in_txbx):
        tid = self.tbl_counter
        self.tbl_counter += 1
        sid = None
        tp = tbl.find(qn("w:tblPr"))
        if tp is not None:
            ts = tp.find(qn("w:tblStyle"))
            sid = wa(ts, "val") if ts is not None else None
        info = {"el": tbl, "id": tid, "first": len(self.paras), "last": None, "sid": sid, "nested": in_table}
        rows = [r for r in expand(tbl) if r.tag == T_TR]
        for ri, tr in enumerate(rows):
            cells = [c for c in expand(tr) if c.tag == T_TC]
            for ci, tc in enumerate(cells):
                self.walk(tc, True, in_txbx, (tid, ri, ci, len(cells)), sid)
        info["last"] = len(self.paras) - 1
        self.tables.append(info)

    # ---- 段落
    def seg_props(self, psid, tbl_sid, rpr_d):
        key = (psid, tbl_sid, repr(rpr_d))
        v = self._memo.get(key)
        if v is None:
            eff = self.styles.eff_rpr(psid, rpr_d, tbl_sid)
            ea, asc = self.styles.fonts(eff)
            v = (ea, asc, Styles.half_points(eff), bool(eff.get("b")), self.styles.color_rgb(eff))
            self._memo[key] = v
        return v

    def add_para(self, p_el, in_table, in_txbx, tbl_ctx, tbl_sid):
        st = self.styles
        para = Para()
        para.el = p_el
        para.in_table = in_table
        para.in_txbx = in_txbx
        para.tbl = tbl_ctx
        para.tbl_sid = tbl_sid
        para.area = self.area
        para.sect = self.sect_idx
        dppr = parse_ppr(p_el.find(T_PPR))
        psid = dppr.get("pStyle")
        if psid not in st.by_id:
            psid = st.default_par
        para.psid = psid
        style_part = st.base_ppr(psid, tbl_sid)
        # ---- 编号
        numpr = dict(style_part.get("numPr", {}))
        numpr.update(dppr.get("numPr", {}))
        label = None
        lvl = None
        nid = numpr.get("numId")
        if nid not in (None, "", "0"):
            il = numpr.get("ilvl")
            if il is None:
                lk = self.numbering.link_for_style(psid)
                il = lk[1] if lk else 0
            try:
                res = self.numbering.next_label(nid, int(il))
            except Exception:
                res = None
            if res:
                label, lvl = res
        eff = style_part
        if lvl is not None:
            eff = merge_ppr(eff, {k: v for k, v in lvl["ppr"].items() if k in ("ind", "tabs")})
        eff = merge_ppr(eff, {k: v for k, v in dppr.items() if k not in ("markRPr", "sectPr", "pStyle")})
        para.ppr = eff
        para.sectpr = dppr.get("sectPr")
        # ---- 文字
        ctx = {"psid": psid, "tbl": tbl_sid, "segs": [], "pos": 0, "fields": [], "fstack": [], "pb": [], "txbx": []}
        if label is not None and label != "":
            mark = dppr.get("markRPr", {})
            base = st.eff_rpr(psid, mark, tbl_sid)
            base = merge_rpr(base, lvl["rpr"])
            ea, asc = st.fonts(base)
            props = (ea, asc, Styles.half_points(base), bool(base.get("b")), st.color_rgb(base))
            ctx["segs"].append(Seg(label, props, "num"))
            ctx["pos"] = len(label)
        self.scan(p_el, ctx)
        para.segs = ctx["segs"]
        para.fields = ctx["fields"]
        para.pb = ctx["pb"]
        para.text = "".join(s.text for s in para.segs)
        para.idx = len(self.paras)
        self.paras.append(para)
        if para.sectpr is not None and self.area == "body" and not in_table and not in_txbx:
            self.sections.append(para.sectpr)
            self.sect_idx += 1
        # 文本框内段落紧随宿主段落之后
        for tb in ctx["txbx"]:
            self.walk(tb, in_table, True, tbl_ctx, tbl_sid)

    # ---- 行内扫描
    def scan(self, el, ctx):
        for ch in expand(el):
            t = ch.tag
            if t == T_R:
                self.scan_run(ch, ctx)
            elif t == T_FLDSIMPLE:
                self.scan_fldsimple(ch, ctx)
            elif t == T_SDT:
                sc = ch.find(T_SDTC)
                if sc is not None:
                    self.scan(sc, ctx)
            elif t in _TRANSPARENT:
                self.scan(ch, ctx)
            # w:del / w:moveFrom 及其余元素：忽略

    def emit(self, ctx, buf, props):
        text = "".join(buf)
        if text:
            ctx["segs"].append(Seg(text, props))
            ctx["pos"] += len(text)
        del buf[:]

    def scan_fldsimple(self, el, ctx):
        instr = wa(el, "instr", "") or ""
        start = ctx["pos"]
        first_run = None
        for r in el.iter(T_R):
            if first_run is None:
                first_run = r
            if r.find(T_RPR) is not None:
                first_run = r
                break
        self.scan(el, ctx)
        tok = instr.strip().split()
        if tok:
            props = None
            if first_run is not None:
                props = self.seg_props(ctx["psid"], ctx["tbl"], parse_rpr(first_run.find(T_RPR)))
            ctx["fields"].append({"kind": tok[0].upper(), "instr": instr.strip(), "start": start,
                                  "end": ctx["pos"], "props": props, "simple": True})

    def scan_run(self, r, ctx):
        rpr_d = parse_rpr(r.find(T_RPR))
        props = self.seg_props(ctx["psid"], ctx["tbl"], rpr_d)
        has_rpr = r.find(T_RPR) is not None
        fst = ctx["fstack"]
        if fst:
            for f in fst:
                if f["props"] is None and (has_rpr or f["first"] is None):
                    if has_rpr:
                        f["props"] = props
                    elif f["first"] is None:
                        f["first"] = props
        buf = []
        for c in expand(r):
            n = c.tag
            if n == T_T:
                buf.append(c.text or "")
            elif n == T_TAB or n == T_PTAB:
                buf.append("\t")
            elif n == T_BR:
                ty = wa(c, "type")
                if ty == "page":
                    self.emit(ctx, buf, props)
                    ctx["pb"].append(ctx["pos"])
                elif ty == "column":
                    pass
                else:
                    buf.append("\n")
            elif n == T_CR:
                buf.append("\n")
            elif n == T_NBH:
                buf.append("-")
            elif n == T_FLDCHAR:
                self.emit(ctx, buf, props)
                ty = wa(c, "fldCharType")
                if ty == "begin":
                    fst.append({"instr": "", "state": "instr", "start": ctx["pos"], "props": props if has_rpr else None,
                                "first": props})
                elif ty == "separate" and fst:
                    fst[-1]["state"] = "result"
                    fst[-1]["start"] = ctx["pos"]
                elif ty == "end" and fst:
                    f = fst.pop()
                    tok = f["instr"].strip().split()
                    if tok:
                        ctx["fields"].append({"kind": tok[0].upper(), "instr": f["instr"].strip(),
                                              "start": f["start"], "end": ctx["pos"],
                                              "props": f["props"] or f["first"], "simple": False})
            elif n == T_INSTR:
                if fst and fst[-1]["state"] == "instr":
                    fst[-1]["instr"] += c.text or ""
            elif n in (T_DRAWING, T_PICT, T_OBJECT):
                for tb in iter_txbx(c):
                    ctx["txbx"].append(tb)
        self.emit(ctx, buf, props)


# ====================================================================== 节与页眉页脚
def parse_sect(sp):
    d = {"el": sp, "w": None, "h": None, "orient": None, "top": None, "bottom": None, "left": None,
         "right": None, "grid": None, "titlePg": False, "type": "nextPage", "refs": {"header": {}, "footer": {}}}
    pg = sp.find(qn("w:pgSz"))
    if pg is not None:
        d["w"] = meas_to_twips(wa(pg, "w"))
        d["h"] = meas_to_twips(wa(pg, "h"))
        d["orient"] = wa(pg, "orient")
    pm = sp.find(qn("w:pgMar"))
    if pm is not None:
        for k in ("top", "bottom", "left", "right", "header", "footer", "gutter"):
            d[k] = meas_to_twips(wa(pm, k))
    dg = sp.find(qn("w:docGrid"))
    if dg is not None:
        lp = wa(dg, "linePitch")
        d["grid"] = {"type": wa(dg, "type") or "default",
                     "linePitch": meas_to_twips(lp) if lp is not None else None}
    tp = sp.find(qn("w:titlePg"))
    d["titlePg"] = bool(tp is not None and is_on(tp))
    ty = sp.find(qn("w:type"))
    if ty is not None:
        d["type"] = wa(ty, "val") or "nextPage"
    for kind, tag in (("header", "w:headerReference"), ("footer", "w:footerReference")):
        for ref in sp.findall(qn(tag)):
            rid = ref.get("{%s}id" % R_NS)
            typ = wa(ref, "type") or "default"
            if rid:
                d["refs"][kind][typ] = rid
    return d


class Doc:
    """已解析的文档：包、样式、编号、段落列表、节、页眉页脚。"""

    def __init__(self, path):
        self.pkg = Package(path)
        pkg = self.pkg
        self.theme = Theme(pkg.theme_root)
        self.styles = Styles(pkg.styles_root, self.theme)
        self.numbering = Numbering(pkg.numbering_root)
        self.body = pkg.doc.find(qn("w:body"))
        if self.body is None:
            raise ValueError("document.xml 缺少 w:body")
        w = Walker(self, "body")
        w.walk(self.body)
        self.paras = w.paras
        self.tables = w.tables
        self.sections = [parse_sect(s) for s in w.sections]
        if not self.sections:
            self.sections = [parse_sect(etree.Element(T_SECT))]
        self.even_odd = False
        if pkg.settings_root is not None:
            eo = pkg.settings_root.find(qn("w:evenAndOddHeaders"))
            self.even_odd = bool(eo is not None and is_on(eo))
        self._hf = {}

    def part_paras(self, part):
        """页眉/页脚部件内段落（缓存）。"""
        if part not in self._hf:
            root = self.pkg.xml(part)
            paras, tables = [], []
            if root is not None:
                w = Walker(self, "footer", numbering=Numbering(self.pkg.numbering_root))
                w.walk(root)
                paras, tables = w.paras, w.tables
            self._hf[part] = (paras, tables)
        return self._hf[part][0]

    def part_tables(self, part):
        self.part_paras(part)
        return self._hf[part][1]

    def section_parts(self, si, kind):
        """第 si 节生效的页眉/页脚部件：{'default':part,'even':part,'first':part}，缺省引用继承前一节。"""
        res = {}
        for i in range(min(si, len(self.sections) - 1) + 1):
            for typ, rid in self.sections[i]["refs"][kind].items():
                tgt = self.pkg.rels.get(rid)
                if tgt:
                    res[typ] = tgt[1]
        return res

    def all_hf_parts(self):
        """所有被引用的页眉页脚部件名（去重，保持顺序）。"""
        seen, out = set(), []
        for i in range(len(self.sections)):
            for kind in ("header", "footer"):
                for typ, part in self.section_parts(i, kind).items():
                    if part not in seen:
                        seen.add(part)
                        out.append(part)
        return out



# ====================================================================== 段落角色识别
RE_FWZH = re.compile(r"\S*[〔\[［（(【〖]\s*\d{4}\s*[〕\]］）)】〗]\s*(?:第\s*)?\d+\s*号")
RE_H = {
    1: re.compile(r"^[一二三四五六七八九十百]+、"),
    2: re.compile(r"^[（(][一二三四五六七八九十百]+[）)]"),
    3: re.compile(r"^\d{1,2}[.．](?!\d)"),
    4: re.compile(r"^[（(]\d{1,2}[）)]"),
}
RE_DATE_AR = re.compile(r"^(\d{4})年(\d{1,2})月(\d{1,2})日$")
RE_DATE_CN = re.compile(r"^[〇○零一二三四五六七八九OＯ]{2,4}年[零〇一二三四五六七八九十]{1,3}月[零〇一二三四五六七八九十]{1,3}日$")
RE_DATE_OTHER = re.compile(r"^\d{4}[-./]\d{1,2}[-./]\d{1,2}日?$")
RE_ATTACH_MARK = re.compile(r"^附件\d*$")
RE_ATTACH_NOTE = re.compile(r"^附件[：:]")
HEAD_ROLE = {1: "一级标题", 2: "二级标题", 3: "三级标题", 4: "四级标题"}
ROLE_LEVEL = {v: k for k, v in HEAD_ROLE.items()}
RE_JIYAO = re.compile(r"纪要(?:[（(][^）)]*[）)])?$")
RE_LEADIN = re.compile(r"(如下|以下|下列|如次|下文|所示)[：:]$")
RE_GENRE_END = re.compile(r"(通知|通报|请示|报告|函|纪要|意见|决定|批复|公告|通告|方案|办法|规定|细则|议案|命令|决议|公报|讲话|总结|条例|规划|计划|答复|说明)$")
LAYOUTS = ("红头下行", "上行", "信函", "纪要", "普通材料", "讲话")
LAYOUT_ALIAS = {
    "红头": "红头下行", "下行": "红头下行", "下行文": "红头下行", "红头下行文": "红头下行",
    "上行文": "上行", "上行": "上行", "请示": "上行", "信函格式": "信函", "函": "信函",
    "纪要格式": "纪要", "材料": "普通材料", "无红头": "普通材料", "讲话稿": "讲话",
}


def norm_layout(v):
    if v is None:
        return None
    v = str(v).strip()
    if v in LAYOUTS:
        return v
    return LAYOUT_ALIAS.get(v)


def ends_colon(t):
    return t.rstrip().endswith(("：", ":"))


def heading_level(stext):
    for lv in (1, 2, 3, 4):
        if RE_H[lv].match(stext):
            return lv
    return 0


def date_kind(c):
    """整段日期识别（c 为去空白文本）：返回 (kind, (y,m,d原文)) 或 None。kind: ar/cn/other。"""
    m = RE_DATE_AR.match(c)
    if m:
        return "ar", m.groups()
    if RE_DATE_CN.match(c):
        return "cn", None
    if RE_DATE_OTHER.match(c):
        return "other", None
    return None


class Roles:
    """角色识别结果。各列表元素均为 Para；para.role 同步写入。"""

    def __init__(self):
        self.layout = None
        self.main = []           # 非表格、非文本框、非空段落（文档顺序）
        self.title = []          # 标题行（多行时多个）
        self.title_how = ""      # 标题识别方式说明
        self.hdr = []            # 版头区全部非空段落（含表格、文本框内）
        self.hdr_end = 0         # 版头区结束位置（第一个标题段落的 idx）
        self.marks = []          # 发文机关标志 / 纪要标志
        self.fwzh = None         # 发文字号所在段落
        self.fwzh_m = None       # 正则匹配对象
        self.issuer = None       # 签发人所在段落
        self.recipients = []
        self.heads = {1: [], 2: [], 3: [], 4: []}
        self.attach_first = None
        self.attach_lines = []   # 附件说明续行
        self.attach_marks = []
        self.date = None
        self.date_note = ""
        self.sign = None
        self.record_start = None  # 版记起始段落 idx
        self.record = []
        self.body = []           # 角色=正文的段落
        self.notes = []


def _special_hdr(p):
    """版头专有行的类型：fwzh/issuer/copynum/secret/urgent/None。"""
    t = p.stext
    c = p.ctext
    m = RE_FWZH.search(t)
    if m and len(c) <= 60 and "。" not in t:
        return "fwzh"
    if "签发人" in c and len(c) <= 40:
        return "issuer"
    if re.fullmatch(r"\d{6}", c):
        return "copynum"
    if "★" in t or re.search(r"(?:^|[^一-鿿])(秘密|机密|绝密)", c) or c in ("秘密", "机密", "绝密"):
        if len(c) <= 20:
            return "secret"
    if c in ("特急", "加急", "特提", "平急"):
        return "urgent"
    return None


def _dom_ea(p):
    c = Counter(s.ea for _, s in p.chars())
    return c.most_common(1)[0][0] if c else None


def _title_like(p, layout, mark_ids):
    c = p.ctext
    if p.jc() != "center" or p.is_red() or len(c) < 8:
        return False
    if id(p) in mark_ids:
        return False
    if _special_hdr(p) or ends_colon(p.stext) or RE_ATTACH_MARK.match(c) or date_kind(c):
        return False
    if re.search(r"文件$", c) and "关于" not in c:
        return False
    if re.fullmatch(r"[（(].*[）)]", c) or re.match(r"^第\d+[期次号]", c):
        return False
    if re.fullmatch(r"[—―─━_\-－=＿·.*]+", c):  # 用字符画的分隔线
        return False
    return True


def assign_roles(doc, layout):
    R = Roles()
    R.layout = layout
    paras = doc.paras
    for p in paras:
        p.role = "空段" if p.empty else ""
    main = [p for p in paras if not p.in_table and not p.in_txbx and not p.empty]
    R.main = main

    # ---- 纪要标志（仅纪要版式）：首个居中且以“纪要”结尾的段落
    mark_ids = set()
    if layout == "纪要":
        for p in main[:8]:
            if p.jc() == "center" and RE_JIYAO.search(p.ctext):
                mark_ids.add(id(p))
                break

    # ---- 标题
    # 候选窗口：标题位于文首，遇到主送机关式的冒号行或层次标题即视为已进入正文，其后的居中长段不再算标题
    win = []
    for p in main[:40]:
        if (ends_colon(p.stext) and len(p.ctext) <= 80) or heading_level(p.stext):
            break
        win.append(p)
    title = None
    cands = [p for p in win if _title_like(p, layout, mark_ids)]
    if cands:
        # 多个候选时优先取带“关于”或以文种词结尾的（避免把未标红的发文机关名误当标题）
        title = next((p for p in cands if "关于" in p.ctext or RE_GENRE_END.search(p.ctext)), cands[0])
        R.title_how = "居中长段"
    if title is None and layout != "纪要":
        cn = [p for p in win if p.jc() == "center" and not p.is_red() and id(p) not in mark_ids
              and not _special_hdr(p) and not RE_ATTACH_MARK.match(p.ctext) and len(p.ctext) >= 6
              and not re.fullmatch(r"[（(].*[）)]", p.ctext) and not re.match(r"^第\d+[期次号]", p.ctext)]
        if cn:
            title = max(cn, key=lambda x: x.max_pt())
            R.title_how = "字号最大的居中段"
        else:
            for p in win:
                c = p.ctext
                if (id(p) not in mark_ids and not p.is_red() and not _special_hdr(p) and len(c) <= 80
                        and RE_GENRE_END.search(c) and not c.endswith(("。", "：", ":")) and heading_level(p.stext) == 0
                        and not (layout in ("红头下行", "上行", "信函") and re.search(r"文件$", c) and "关于" not in c)):
                    title = p
                    R.title_how = "文种词特征（非居中）"
                    break
    tlines = []
    if title is not None:
        i = main.index(title)
        tlines.append(title)
        d_ea, d_pt = _dom_ea(title), title.dom_pt()
        for q in main[i + 1:i + 7]:
            if (q.jc() == "center" and not q.is_red() and len(q.ctext) <= 40 and abs(q.dom_pt() - d_pt) <= 0.5
                    and _dom_ea(q) == d_ea and title.jc() == "center" and not _special_hdr(q)
                    and not ends_colon(q.stext) and not date_kind(q.ctext)):
                tlines.append(q)
            else:
                break
    R.title = tlines
    for p in tlines:
        p.role = "标题"

    # ---- 版头区
    if tlines:
        hdr_end = tlines[0].idx
    else:
        hdr_end = 0
        for p in main:
            if (not p.is_red() and not _special_hdr(p) and p.jc() != "center" and id(p) not in mark_ids):
                hdr_end = p.idx
                break
        else:
            hdr_end = paras[-1].idx + 1 if paras else 0
    R.hdr_end = hdr_end
    R.hdr = [p for p in paras if p.idx < hdr_end and not p.empty]
    names = {"fwzh": "发文字号", "issuer": "签发人", "copynum": "份号", "secret": "密级", "urgent": "紧急程度"}
    rest_hdr = []
    for p in R.hdr:
        sp = _special_hdr(p)
        if sp == "fwzh":
            p.role = "发文字号"
            if R.fwzh is None:
                R.fwzh = p
                R.fwzh_m = RE_FWZH.search(p.stext)
            if "签发人" in p.ctext and R.issuer is None:
                R.issuer = p
        elif sp:
            p.role = names[sp]
            if sp == "issuer" and R.issuer is None:
                R.issuer = p
        else:
            rest_hdr.append(p)
    # 发文机关标志 / 纪要标志
    marks = []
    if layout == "纪要":
        marks = [p for p in rest_hdr if id(p) in mark_ids]
        if not marks:
            marks = [p for p in rest_hdr if p.jc() == "center" and RE_JIYAO.search(p.ctext)][:1]
    if not marks:
        reds = [p for p in rest_hdr if p.is_red()]
        if reds:
            mx = max(p.max_pt() for p in reds)
            marks = [p for p in reds if p.max_pt() >= 0.9 * mx]
    if not marks:
        ends = [p for p in rest_hdr if p.jc() == "center" and re.search(r"(文件|纪要)$", p.ctext)]
        if ends:
            marks = [max(ends, key=lambda x: x.max_pt())]
    if not marks:
        ends = [p for p in rest_hdr if re.search(r"(文件|纪要)$", p.ctext) and "关于" not in p.ctext]
        if ends:
            marks = [max(ends, key=lambda x: x.max_pt())]
    if not marks:
        cen = [p for p in rest_hdr if p.jc() == "center" and not p.in_table]
        if cen and layout in ("红头下行", "上行", "信函"):
            top = max(cen, key=lambda x: x.max_pt())
            if top.max_pt() >= 18:
                marks = [top]
    R.marks = marks
    for p in marks:
        p.role = "发文机关标志"
    for p in rest_hdr:
        if not p.role:
            p.role = "版头其他"

    # ---- 版记起始（标题之后第一个“抄送”段，其次“…印发”段）
    after = [p for p in paras if p.idx >= hdr_end and not p.empty]
    rec = None
    for p in after:
        if re.match(r"^抄\s*送", p.stext) and p.role not in ("标题",):
            rec = p
            break
    if rec is None:
        for p in after:
            c = p.ctext
            if c.endswith("印发") and len(c) <= 60 and p.role not in ("标题",):
                rec = p
                break
    if rec is not None:
        R.record_start = rec.idx
        for p in paras:
            if p.idx >= rec.idx:
                p.role = p.role if p.empty else "版记"
                if not p.empty:
                    R.record.append(p)

    # ---- 标题之后的正文区（仅主文档非表格非文本框段落、版记之前）
    limit = R.record_start if R.record_start is not None else 10 ** 9
    seq = [p for p in main if p.idx >= hdr_end and p.idx < limit and p.role != "标题"]
    # 发文字号误放在标题之后（版头区没有）：在标题后的前几行里兜底找一次，仍按发文字号评判格式
    if R.fwzh is None and R.title:
        for q in seq[:6]:
            if ends_colon(q.stext):
                break
            if _special_hdr(q) == "fwzh":
                q.role = "发文字号"
                R.fwzh = q
                R.fwzh_m = RE_FWZH.search(q.stext)
                R.notes.append("发文字号位于标题之后（版头区内未找到），仍按发文字号评判")
                break
        seq = [q for q in seq if q.role != "发文字号"]
    # 副标题/署名日期行/副标题：标题后紧跟的居中短行
    i = 0
    while (i < len(seq) and R.title and i < 4 and seq[i].jc() == "center" and len(seq[i].ctext) <= 40
           and not ends_colon(seq[i].stext) and heading_level(seq[i].stext) == 0 and not seq[i].is_red()
           and not RE_ATTACH_MARK.match(seq[i].ctext) and not RE_ATTACH_NOTE.match(seq[i].stext)):
        seq[i].role = "署名日期行/副标题"
        i += 1
    # 主送机关（可连续多段；允许前几行以顿号/逗号结尾）
    rec_list = []
    j = i
    while j < len(seq) and j < i + 5:
        p = seq[j]
        t = p.stext
        if (ends_colon(t) and len(p.ctext) <= 80 and heading_level(t) == 0 and not RE_LEADIN.search(t)
                and "。" not in t[:-1]):
            rec_list = seq[i:j + 1]
            k = j + 1
            while (k < len(seq) and ends_colon(seq[k].stext) and len(seq[k].ctext) <= 80
                   and heading_level(seq[k].stext) == 0 and not RE_LEADIN.search(seq[k].stext)
                   and "。" not in seq[k].stext[:-1]):
                rec_list.append(seq[k])
                k += 1
            break
        if len(p.ctext) <= 60 and not t.endswith("。") and heading_level(t) == 0 and not ends_colon(t):
            j += 1
            continue
        break
    for p in rec_list:
        p.role = "主送机关"
    R.recipients = rec_list
    start = (seq.index(rec_list[-1]) + 1) if rec_list else i
    rest = seq[start:]

    # ---- 附件标识（全文范围，标题之后）
    allm = [p for p in main if p.idx >= hdr_end and p.role not in ("标题", "署名日期行/副标题")]
    for p in allm:
        if RE_ATTACH_MARK.match(p.ctext):
            p.role = "附件标识"
            R.attach_marks.append(p)
    # 附件标识之后紧跟的居中短段为附件标题
    for ap in R.attach_marks:
        k = main.index(ap) + 1
        if k < len(main) and main[k].jc() == "center" and len(main[k].ctext) <= 60 and not main[k].role:
            main[k].role = "附件标题"
    first_attach = R.attach_marks[0].idx if R.attach_marks else None

    # ---- 成文日期：附件标识之前最后一个整段日期；其次全局最后一个
    dates = [p for p in rest if not p.role and date_kind(p.ctext)]
    pre_att = [p for p in dates if first_attach is None or p.idx < first_attach]
    if pre_att:
        R.date = pre_att[-1]
    elif dates:
        R.date = dates[-1]
        R.date_note = "日期段落位于附件标识之后"
    if R.date is not None:
        R.date.role = "成文日期"
        # 落款署名：日期上一非空段，短且不以句号结尾
        k = main.index(R.date) - 1 if R.date in main else -1
        if k >= 0:
            q = main[k]
            if (not q.role and len(q.ctext) <= 40 and not q.stext.endswith("。") and heading_level(q.stext) == 0
                    and not RE_ATTACH_NOTE.match(q.stext) and q.idx >= hdr_end):
                q.role = "落款署名"
                R.sign = q

    # ---- 附件说明：正文之后的“附件：”段（位置通常在署名之前；若放在署名/日期之后也识别，位置问题不另扣分）
    stop = (R.sign.idx if R.sign is not None else (R.date.idx if R.date is not None else limit))
    for p in rest:
        if p.role:
            continue
        if RE_ATTACH_NOTE.match(p.stext):
            R.attach_first = p
            p.role = "附件说明"
            k = rest.index(p) + 1
            while (k < len(rest) and not rest[k].role and len(rest[k].ctext) <= 80
                   and (p.idx > stop or rest[k].idx < stop)
                   and heading_level(rest[k].stext) not in (1, 2)
                   and (first_attach is None or rest[k].idx < first_attach)):
                rest[k].role = "附件说明续行"
                R.attach_lines.append(rest[k])
                k += 1
            break

    # ---- 层次标题与正文；日期之后的括注为附注
    date_idx = R.date.idx if R.date is not None else None
    for p in rest:
        if p.role:
            continue
        lv = heading_level(p.stext)
        if lv:
            p.role = HEAD_ROLE[lv]
            p.flags.add("head%d" % lv)
            R.heads[lv].append(p)
        elif (date_idx is not None and p.idx > date_idx and (first_attach is None or p.idx < first_attach)
              and re.fullmatch(r"[（(].*[）)]", p.stext) and len(p.ctext) <= 40):
            p.role = "附注"
        elif first_attach is not None and p.idx > first_attach and date_kind(p.ctext):
            p.role = "附件内日期"  # 附件页中的日期不是成文日期，也不计入正文统计
        else:
            p.role = "正文"
            R.body.append(p)
    # 层次标题的标题成分/正文尾部范围
    for lv in (1, 2, 3, 4):
        for p in R.heads[lv]:
            t = p.text
            m = re.search(r"[。：]", t)
            end = m.start() if m else len(t)
            p.flags.add("hc_end:%d" % end)
    # 表格/文本框/其余
    for p in paras:
        if p.role in ("", "空段"):
            if p.empty:
                p.role = "空段"
            elif p.in_table:
                p.role = "表格内"
            elif p.in_txbx:
                p.role = "文本框内"
            else:
                p.role = "其他"
    return R


def hc_end(p):
    """层次标题的标题成分结束偏移（不含第一个“。”或“：”）。"""
    for f in p.flags:
        if f.startswith("hc_end:"):
            return int(f.split(":")[1])
    return len(p.text)


def body_ranges(p):
    """段落中计入“正文”统计的字符区间 [(start,end)]：正文段落为全段；层次标题同段接排的尾部。"""
    if p.role == "正文":
        return [(0, len(p.text))]
    if p.role in ROLE_LEVEL:
        end = hc_end(p)
        if end < len(p.text):
            tail = p.text[end + 1:]
            if compact(tail):
                return [(end + 1, len(p.text))]
    return []


def head_ranges(p):
    return [(0, hc_end(p))]


def detect_layout(doc):
    """自动判定版式：发文字号 -> 红头下行/上行；红色居中“纪要”首段 -> 纪要；“同志们”无字号 -> 讲话；否则普通材料。"""
    ps = [p for p in doc.paras if not p.empty]
    head = ps[:14]
    fw = None
    for p in head:
        t = p.stext
        if RE_FWZH.search(t) and len(p.ctext) <= 60 and "。" not in t:
            fw = p
            break
    if fw is not None:
        for p in head:
            if "签发人" in p.ctext and len(p.ctext) <= 40:
                return "上行"
        return "红头下行"
    for p in ps[:4]:
        if p.jc() == "center" and re.search(r"纪要$", p.ctext) and (p.is_red() or p.max_pt() >= 18):
            return "纪要"
    for p in ps[:4]:
        if p.jc() == "center" and p.is_red() and re.search(r"文件$", p.ctext):
            return "红头下行"
    if any("同志们" in p.text for p in ps):
        return "讲话"
    return "普通材料"



# ====================================================================== 字体规则与审计上下文
STD_FONTS = ("仿宋_GB2312", "楷体_GB2312", "黑体", "方正小标宋简体")
FONT_ALIASES = {
    "fangsong_gb2312": "仿宋_GB2312",
    "kaiti_gb2312": "楷体_GB2312",
    "simhei": "黑体",
    "simsun": "宋体",
    "fzxiaobiaosong-b05s": "方正小标宋简体",
}
NO_FONT = "（未指定）"
OFFICE_LAYOUTS = ("红头下行", "上行", "信函")


class Ctx:
    def __init__(self, doc, roles, layout, strict):
        self.doc = doc
        self.R = roles
        self.layout = layout
        self.strict = strict
        self.alias_used = {}      # 原名 -> 标准名
        self.notes = []
        self.body_sect = 0
        for p in doc.paras:
            if p.role == "正文":
                self.body_sect = min(p.sect, len(doc.sections) - 1)
                break
        self._foot = None

    def canon(self, name):
        """东亚字体名归一：英文别名（默认）折算为标准中文名；--strict-font-names 时不折算。"""
        n = name if name else NO_FONT
        if not self.strict:
            k = n.lower()
            if k in FONT_ALIASES and n != FONT_ALIASES[k]:
                self.alias_used[n] = FONT_ALIASES[k]
                return FONT_ALIASES[k]
        return n

    def alias_note(self, names):
        used = sorted(n for n in names if n in self.alias_used)
        if not used:
            return ""
        return "；按英文别名视为等同：" + "、".join("%s→%s" % (n, self.alias_used[n]) for n in used)


def chars_in(p, ranges):
    """段落内给定区间的可见字符 [(ch, seg)]。"""
    ac = p.all_chars()
    out = []
    for a, b in ranges:
        for ch, s in ac[a:b]:
            if not ch.isspace():
                out.append((ch, s))
    return out


def share_str(counter, total, limit=5, fmt=lambda k: str(k)):
    items = counter.most_common()
    parts = ["%s(占 %d%%)" % (fmt(k), round(100.0 * v / total)) if total else fmt(k) for k, v in items[:limit]]
    if len(items) > limit:
        parts.append("等 %d 种" % len(items))
    return "；".join(parts)


def pct(x):
    return "%d%%" % round(100.0 * x)


def body_chars(ctx):
    """计入“正文”统计的字符 [(ch, seg)]：正文段落全文 + 层次标题同段接排的句号后部分。"""
    out = []
    for p in ctx.doc.paras:
        rg = body_ranges(p)
        if rg:
            out.extend(chars_in(p, rg))
    return out


def layout_paras(ctx):
    """04/05 统计对象：正文 + 层次标题段落（非表格、非文本框）。"""
    return [p for p in ctx.doc.paras if p.role == "正文" or p.role in ROLE_LEVEL]


# ====================================================================== 01 02
def chk_01(ctx):
    exp = "A4 210×297mm（±1mm）"
    s = ctx.doc.sections[0]
    if s["w"] is None or s["h"] is None:
        return "fail", exp, "首节未设置纸张尺寸 w:pgSz（Word 缺省为 Letter 215.9×279.4mm）"
    wm, hm = tw2mm(s["w"]), tw2mm(s["h"])
    ok = abs(wm - 210) <= 1 and abs(hm - 297) <= 1
    meas = "%.1f×%.1fmm" % (wm, hm)
    if s["orient"] == "landscape":
        meas += "（横向）"
    return ("pass" if ok else "fail"), exp, meas


def chk_02(ctx):
    exp = "上37/左28/右26/下35mm（±1.5mm）"
    s = ctx.doc.sections[0]
    names = (("上", "top", 37), ("左", "left", 28), ("右", "right", 26), ("下", "bottom", 35))
    if any(s[k] is None for _, k, _ in names):
        got = "/".join("%s%s" % (n, ("%.1f" % tw2mm(s[k])) if s[k] is not None else "未设") for n, k, _ in names)
        return "fail", exp, "首节页边距未完整设置：%s（缺项按 Word 缺省值计，视为不符）" % got
    ok = True
    bad = []
    parts = []
    for n, k, std in names:
        v = abs(tw2mm(s[k]))  # 负的上/下边距在 Word 里表示“精确”页眉页脚距离，取绝对值
        parts.append("%s%.1f" % (n, v))
        if abs(v - std) > 1.5:
            ok = False
            bad.append("%s边距偏差 %.1fmm" % (n, v - std))
    meas = "/".join(parts) + "mm"
    if bad:
        meas += "；" + "、".join(bad)
    return ("pass" if ok else "fail"), exp, meas


# ====================================================================== 03a 03b
def chk_03a(ctx):
    exp = "仿宋_GB2312（≥95%）"
    ea = Counter()
    for ch, s in body_chars(ctx):
        if is_ea_char(ch):
            ea[s.ea or NO_FONT] += 1
    total = sum(ea.values())
    if not total:
        return "na", exp, "无正文段落（或正文不含汉字）"
    good = sum(v for k, v in ea.items() if ctx.canon(k) == "仿宋_GB2312")
    meas = share_str(ea, total) + ctx.alias_note(ea.keys())
    return ("pass" if good / total >= 0.95 else "fail"), exp, meas


def chk_03b(ctx):
    exp = "三号 16pt（±0.5pt，≥95%）"
    chs = body_chars(ctx)
    if not chs:
        return "na", exp, "无正文段落"
    c = Counter(s.sz for _, s in chs)
    good = sum(v for k, v in c.items() if 31 <= k <= 33)
    total = len(chs)
    meas = share_str(c, total, fmt=lambda k: fmt_pt(k / 2.0))
    return ("pass" if good / total >= 0.95 else "fail"), exp, meas


# ====================================================================== 04 行距
def line_desc(p, sects):
    """单个段落的行距判定 -> (是否合格, 描述)。"""
    sp = p.ppr.get("spacing", {})
    rule = sp.get("lineRule") or "auto"
    line = meas_to_twips(sp.get("line"))
    sect = sects[min(p.sect, len(sects) - 1)]
    grid = sect.get("grid")
    if rule in ("exact", "atLeast"):
        if line is None:
            return False, "固定值未给出数值"
        ok = 560 <= line <= 600
        return ok, "%s%s" % ("固定值" if rule == "exact" else "最小值", fmt_pt(line / 20.0))
    mult = (line / 240.0) if line else 1.0
    snap = p.ppr.get("snapToGrid") is not False
    if grid and grid.get("type") in ("lines", "linesAndChars") and grid.get("linePitch") and snap:
        rows = 12755.9 / grid["linePitch"]
        ok = 21 <= round(rows) <= 23 and 0.95 <= mult <= 1.05
        return ok, "%s倍行距+文档网格(每页约 %.1f 行)" % (("%g" % round(mult, 2)), rows)
    return False, "%s倍行距(无行网格)" % ("%g" % round(mult, 2))


def chk_04(ctx):
    exp = "每页22行：固定值28～30pt，或网格22行+单倍行距（≥80%段落）"
    ps = layout_paras(ctx)
    if not ps:
        return "na", exp, "无正文/层次标题段落"
    c = Counter()
    good = 0
    for p in ps:
        ok, d = line_desc(p, ctx.doc.sections)
        c[d] += 1
        good += ok
    meas = share_str(c, len(ps), limit=4)
    if any(k.startswith("最小值") for k in c):
        meas += "；含“最小值”行距（按固定值口径容忍）"
    return ("pass" if good / len(ps) >= 0.8 else "fail"), exp, meas


# ====================================================================== 05 首行缩进
def indent_state(p):
    """-> (合格?, 描述)。"""
    ind = p.ppr.get("ind", {})
    fl = meas_to_twips(ind.get("firstLine"))
    flc = meas_to_twips(ind.get("firstLineChars"))
    hg = meas_to_twips(ind.get("hanging"))
    hgc = meas_to_twips(ind.get("hangingChars"))
    if (hg and hg > 0) or (hgc and hgc > 0):
        return False, "悬挂缩进"
    if flc is not None:  # 字符单位优先于绝对单位（firstLineChars=0 也表示无缩进）
        if flc == 0:
            return False, "无首行缩进"
        return (190 <= flc <= 210), "首行缩进 %g 字" % round(flc / 100.0, 2)
    if fl:
        return (560 <= fl <= 720), "首行缩进 %s" % fmt_pt(fl / 20.0)
    raw = p.raw_text
    if raw.startswith("　") or raw.startswith("  ") or raw.startswith("\t"):
        return False, "以空格代替缩进"
    return False, "无首行缩进"


def chk_05(ctx):
    exp = "首行缩进两字（约32pt/200字符，≥90%段落）"
    ps = layout_paras(ctx)
    if not ps:
        return "na", exp, "无正文/层次标题段落"
    c = Counter()
    good = 0
    for p in ps:
        ok, d = indent_state(p)
        c[d] += 1
        good += ok
    return ("pass" if good / len(ps) >= 0.9 else "fail"), exp, share_str(c, len(ps), limit=4)


# ====================================================================== 06 标题
def _prop_check(ctx, p_ranges, want_font, want_sz):
    """对若干(段落,区间)统计字体、字号达标占比 -> (字体占比, 字号占比, 字体计数, 字号计数, 总字符, 东亚字符)。"""
    fonts, sizes = Counter(), Counter()
    nea = 0
    nall = 0
    for p, rg in p_ranges:
        for ch, s in chars_in(p, rg):
            nall += 1
            sizes[s.sz] += 1
            if is_ea_char(ch):
                nea += 1
                fonts[s.ea or NO_FONT] += 1
    fgood = sum(v for k, v in fonts.items() if ctx.canon(k) == want_font) if want_font else nea
    sgood = sum(v for k, v in sizes.items() if want_sz[0] <= k <= want_sz[1])
    return fgood, nea, sgood, nall, fonts, sizes


def chk_06(ctx):
    exp = "二号22pt、方正小标宋简体、居中"
    R = ctx.R
    if not R.title:
        if ctx.layout == "纪要":
            return "na", exp, "纪要以红色纪要标志代替标题，未发现另行设置的标题"
        return "fail", exp, "未识别到标题段落"
    bad = []
    fonts_all, sizes_all = Counter(), Counter()
    aligns = Counter()
    for p in R.title:
        fg, nea, sg, nall, fonts, sizes = _prop_check(ctx, [(p, [(0, len(p.text))])], "方正小标宋简体", (43, 45))
        fonts_all.update(fonts)
        sizes_all.update(sizes)
        aligns[p.jc()] += 1
        if nea and fg / nea < 0.95:
            bad.append("字体")
        if nall and sg / nall < 0.95:
            bad.append("字号")
        if p.jc() != "center":
            bad.append("对齐")
    al = {"center": "居中", "left": "左对齐", "right": "右对齐", "both": "两端对齐", "distribute": "分散对齐"}
    meas = "字体 %s；字号 %s；对齐 %s" % (
        "、".join(sorted(fonts_all, key=lambda k: -fonts_all[k])),
        "、".join(fmt_pt(k / 2.0) for k in sorted(sizes_all, key=lambda k: -sizes_all[k])),
        "、".join(al.get(k, k) for k in aligns))
    meas += ctx.alias_note(fonts_all.keys())
    if R.title_how and R.title_how != "居中长段":
        meas += "；标题按“%s”识别" % R.title_how
    bad = sorted(set(bad))
    if bad:
        meas += "；不符：" + "、".join(bad)
    return ("pass" if not bad else "fail"), exp, meas


# ====================================================================== 07 层次标题
def _chk_heads(ctx, levels, want_font, label):
    exp = "%s、三号16pt" % want_font
    ps = [p for lv in levels for p in ctx.R.heads[lv]]
    if not ps:
        return "na", exp, "无%s" % label
    pr = [(p, head_ranges(p)) for p in ps]
    fg, nea, sg, nall, fonts, sizes = _prop_check(ctx, pr, want_font, (31, 33))
    if not nall:
        return "na", exp, "%s无可见文字" % label
    meas = "%d 个%s；字体 %s；字号 %s" % (
        len(ps), label, share_str(fonts, nea) if nea else "无汉字",
        share_str(sizes, nall, fmt=lambda k: fmt_pt(k / 2.0)))
    meas += ctx.alias_note(fonts.keys())
    ok = (nea == 0 or fg / nea >= 0.95) and sg / nall >= 0.95
    return ("pass" if ok else "fail"), exp, meas


def chk_07a(ctx):
    return _chk_heads(ctx, [1], "黑体", "一级标题")


def chk_07b(ctx):
    return _chk_heads(ctx, [2], "楷体_GB2312", "二级标题")


def chk_07c(ctx):
    return _chk_heads(ctx, [3, 4], "仿宋_GB2312", "三、四级标题")


# ====================================================================== 08 主送机关顶格
def chk_08(ctx):
    exp = "顶格（无首行缩进、无左缩进）"
    rs = ctx.R.recipients
    if not rs:
        return "na", exp, "未识别到主送机关"
    bad = []
    for p in rs:
        ind = p.ppr.get("ind", {})
        probs = []
        for k, nm in (("left", "左缩进"), ("firstLine", "首行缩进"), ("firstLineChars", "首行缩进字符"),
                      ("leftChars", "左缩进字符"), ("hanging", "悬挂缩进"), ("hangingChars", "悬挂缩进字符")):
            v = meas_to_twips(ind.get(k))
            if v:
                probs.append("%s %g" % (nm, round(v, 1)))
        raw = p.raw_text
        if raw.startswith("　") or raw.startswith("  ") or raw.startswith("\t"):
            probs.append("以空格代替顶格")
        if probs:
            bad.append("“%s”%s" % (p.short(10), "、".join(probs)))
    if bad:
        return "fail", exp, "；".join(bad)
    return "pass", exp, "%d 段主送机关均无缩进" % len(rs)



# ====================================================================== 页码（页脚）
DASH_OK = "—―─－"           # — ― ─ －
DASH_BAD = {"一": "汉字“一”", "-": "半角连字符“-”", "–": "en dash“–”"}


def page_fields(doc, part):
    out = []
    for p in doc.part_paras(part):
        for f in p.fields:
            if f["kind"] == "PAGE":
                out.append((p, f))
    return out


def foot_info(ctx):
    """正文所在节的页脚：{'default'|'even'|'first': {'part','pages'} 或 None}。"""
    if ctx._foot is None:
        doc = ctx.doc
        parts = doc.section_parts(ctx.body_sect, "footer")
        info = {"parts": parts}
        for typ in ("default", "even", "first"):
            part = parts.get(typ)
            info[typ] = {"part": part, "pages": page_fields(doc, part)} if part else None
        ctx._foot = info
    return ctx._foot


def _cn_align(a):
    return {"left": "居左", "center": "居中", "right": "居右"}.get(a, a)


def pnum_align(p, f, odd):
    """页码段落的对齐 -> (left|center|right, 说明)。解析 framePr.xAlign / 表格单元格位置 / jc / 制表位。"""
    fp = p.ppr.get("framePr")
    if fp and fp.get("xAlign"):
        xa = fp["xAlign"]
        if xa == "inside":
            xa = "left" if odd else "right"
        elif xa == "outside":
            xa = "right" if odd else "left"
        if xa in ("left", "center", "right"):
            return xa, "框架 xAlign=%s" % fp["xAlign"]
    if p.tbl and p.tbl[3] >= 2:
        ci, n = p.tbl[2], p.tbl[3]
        pos = "left" if ci == 0 else ("right" if ci == n - 1 else "center")
        return pos, "表格第 %d/%d 格" % (ci + 1, n)
    jc = p.jc()
    if jc in ("center", "right"):
        return jc, "段落对齐"
    nt = p.text[:f["start"]].count("\t")
    if nt >= 1:
        stops = sorted([(meas_to_twips(pos) or 0, val) for val, pos in p.ppr.get("tabs", []) if val != "clear"])
        if nt <= len(stops):
            v = stops[nt - 1][1]
            v = {"end": "right", "start": "left"}.get(v, v)
            if v in ("left", "center", "right"):
                return v, "制表位"
    return "left", "段落左对齐(默认)"


def _na_nopage(ctx):
    fi = foot_info(ctx)
    return not any(fi[t] and fi[t]["pages"] for t in ("default", "even"))


def _page_targets(ctx):
    """需要检查格式的 (标签, 段落, 域, 是否奇数页)。"""
    fi = foot_info(ctx)
    out = []
    if fi["default"] and fi["default"]["pages"]:
        p, f = fi["default"]["pages"][0]
        out.append(("奇数页" if ctx.doc.even_odd else "各页", p, f, True))
    if ctx.doc.even_odd and fi["even"] and fi["even"]["pages"]:
        p, f = fi["even"]["pages"][0]
        out.append(("偶数页", p, f, False))
    if not out:
        for t in ("even",):
            if fi[t] and fi[t]["pages"]:
                p, f = fi[t]["pages"][0]
                out.append(("偶数页", p, f, False))
    return out


def chk_09a(ctx):
    exp = "页脚含 PAGE 域"
    fi = foot_info(ctx)
    if not fi["parts"]:
        return "fail", exp, "正文所在节未引用任何页脚"
    got = [t for t in ("default", "even") if fi[t] and fi[t]["pages"]]
    if got:
        return "pass", exp, "含 PAGE 域的页脚：" + "、".join({"default": "奇数页/默认", "even": "偶数页"}[t] for t in got)
    return "fail", exp, "页脚中无 PAGE 域"


def chk_09b(ctx):
    exp = "四号14pt、东亚字体宋体"
    if _na_nopage(ctx):
        return "na", exp, "无页码"
    ok_all, parts = True, []
    for lab, p, f, odd in _page_targets(ctx):
        props = f.get("props")
        if not props:
            ok_all = False
            parts.append("%s：无法取得页码域格式" % lab)
            continue
        ea, asc, sz, _b, _c = props
        okf = ctx.canon(ea) == "宋体"
        oks = 27 <= sz <= 29
        ok_all = ok_all and okf and oks
        parts.append("%s：东亚字体 %s、字号 %s（西文字体 %s，仅作提示）%s" % (
            lab, ea or NO_FONT, fmt_pt(sz / 2.0), asc or NO_FONT, ctx.alias_note([ea]) if ea else ""))
    return ("pass" if ok_all else "fail"), exp, "；".join(parts)


def chk_09c(ctx):
    exp = "页码两侧各有一字线，形如“— 1 —”"
    if _na_nopage(ctx):
        return "na", exp, "无页码"
    ok_all, parts = True, []
    for lab, p, f, odd in _page_targets(ctx):
        before, after = p.text[:f["start"]], p.text[f["end"]:]
        shown = re.sub(r"\s+", " ", (before.strip() + " {页码} " + after.strip())).strip()
        okb = any(c in DASH_OK for c in before)
        oka = any(c in DASH_OK for c in after)
        ok = okb and oka
        ok_all = ok_all and ok
        extra = ""
        if not ok:
            bad = [DASH_BAD[c] for c in DASH_BAD if c in before + after]
            extra = "（使用了%s，不是一字线）" % "、".join(bad) if bad else "（两侧缺少一字线）"
        parts.append("%s：%s%s" % (lab, shown, extra))
    return ("pass" if ok_all else "fail"), exp, "；".join(parts)


def chk_09d(ctx):
    exp = "奇数页页码居右、偶数页居左（需设奇偶页不同）"
    if _na_nopage(ctx):
        return "na", exp, "无页码"
    fi = foot_info(ctx)
    doc = ctx.doc
    if not doc.even_odd:
        p, f = (fi["default"]["pages"] or fi["even"]["pages"])[0] if fi["default"] and fi["default"]["pages"] else fi["even"]["pages"][0]
        a, how = pnum_align(p, f, True)
        return "fail", exp, "未设奇偶页不同，所有页同一页脚，对齐为%s（%s）" % (_cn_align(a), how)
    odd = fi["default"]["pages"][0] if fi["default"] and fi["default"]["pages"] else None
    even = fi["even"]["pages"][0] if fi["even"] and fi["even"]["pages"] else None
    parts = []
    ok = True
    if odd:
        a, how = pnum_align(odd[0], odd[1], True)
        parts.append("奇数页%s（%s）" % (_cn_align(a), how))
        ok = ok and a == "right"
    else:
        parts.append("奇数页无页码")
        ok = False
    if even:
        a, how = pnum_align(even[0], even[1], False)
        parts.append("偶数页%s（%s）" % (_cn_align(a), how))
        ok = ok and a == "left"
    else:
        parts.append("偶数页无页码（未引用 even 页脚或其中无 PAGE 域）")
        ok = False
    return ("pass" if ok else "fail"), exp, "；".join(parts)


def chk_09e(ctx):
    exp = "信函首页不显示页码（titlePg 且首页页脚无 PAGE 域）"
    if ctx.layout != "信函":
        return "na", exp, "仅信函版式检查"
    doc = ctx.doc
    sec0 = doc.sections[0]
    parts0 = doc.section_parts(0, "footer")
    first = parts0.get("first")
    has = bool(first and page_fields(doc, first))
    if not sec0["titlePg"]:
        return "fail", exp, "首节未设“首页不同”(titlePg)，首页与其他页同一页脚"
    if has:
        return "fail", exp, "已设首页不同，但首页页脚仍含页码域"
    return "pass", exp, "已设首页不同(titlePg)；首页页脚%s" % ("无页码域" if first else "未引用（为空）")


# ====================================================================== 10 成文日期
def chk_10(ctx):
    exp = "阿拉伯数字“YYYY年M月D日”，年份四位、月日不编虚位"
    R = ctx.R
    d = R.date
    if d is None:
        if ctx.layout in OFFICE_LAYOUTS:
            return "fail", exp, "未找到成文日期段落"
        return "na", exp, "未找到成文日期（该版式不强制）"
    c = d.ctext
    k = date_kind(c)
    shown = d.stext
    note = ("；" + R.date_note) if R.date_note else ""
    if k is None:
        return "fail", exp, "“%s”不是日期格式%s" % (shown, note)
    kind, g = k
    if kind == "cn":
        return "fail", exp, "“%s”为汉字数字日期%s" % (shown, note)
    if kind == "other":
        return "fail", exp, "“%s”不是“YYYY年M月D日”格式%s" % (shown, note)
    y, m, dd = g
    probs = []
    if not (y + m + dd).isascii():
        probs.append("使用了全角数字")
    if len(y) != 4:
        probs.append("年份不是四位全称")
    if len(m) > 1 and m.startswith("0"):
        probs.append("月份编虚位")
    if len(dd) > 1 and dd.startswith("0"):
        probs.append("日编虚位")
    if probs:
        return "fail", exp, "“%s”：%s%s" % (shown, "、".join(probs), note)
    return "pass", exp, "“%s”%s" % (shown, note)


# ====================================================================== 11 附件
def _first_indent_ok(p):
    """左空二字：首行起点距左边缘约两字。

    首行起点 = left + firstLine - hanging（悬挂缩进时首行比其余各行靠左，附件说明“附件：1. xxx”
    换行后与名称首字对齐就是这种写法：left=2080、hanging=1440，首行起点仍是 640 twips，即两字）。
    """
    ind = p.ppr.get("ind", {})
    fl = meas_to_twips(ind.get("firstLine")) or 0
    flc = meas_to_twips(ind.get("firstLineChars")) or 0
    lf = meas_to_twips(ind.get("left")) or 0
    lfc = meas_to_twips(ind.get("leftChars")) or 0
    hg = meas_to_twips(ind.get("hanging")) or 0
    hgc = meas_to_twips(ind.get("hangingChars")) or 0
    start_c = lfc + flc - hgc       # 字符单位（百分之一字）
    start_t = lf + fl - hg          # 绝对单位（twips）
    if (flc or lfc or hgc) and 190 <= start_c <= 210:
        return True, "首行起点 2 字" if (hg or hgc) else ("首行缩进 2 字" if flc else "左缩进 2 字")
    if (fl or lf or hg) and 560 <= start_t <= 720 and not (flc or lfc or hgc):
        if hg:
            return True, "首行起点 %s（悬挂缩进）" % fmt_pt(start_t / 20.0)
        return True, ("首行缩进 %s" % fmt_pt(fl / 20.0)) if fl else ("左缩进 %s" % fmt_pt(lf / 20.0))
    raw = p.raw_text
    if raw.startswith("　　"):
        return False, "以空格代替缩进"
    return False, "未左空二字"


def chk_11a(ctx):
    exp = "附件说明：左空二字；多个附件用阿拉伯数字编号；名称末尾不加标点"
    R = ctx.R
    p = R.attach_first
    if p is None:
        return "na", exp, "无附件说明"
    ok_ind, ind_d = _first_indent_ok(p)
    lines = [p] + R.attach_lines
    items = []
    for i, q in enumerate(lines):
        t = q.stext
        if i == 0:
            t = re.sub(r"^附件[：:]\s*", "", t)
        items.append(t.strip())
    probs = []
    if len(items) > 1 and not items[0]:  # “附件：”后换行再列名称的写法：名称从下一行起算
        items = items[1:]
    multi = len(items) > 1
    num_desc = "单个附件"
    if multi:
        ok_num = True
        for i, t in enumerate(items, 1):
            m = re.match(r"^(\d+)[.．]\s*", t)
            if not m or int(m.group(1)) != i:
                ok_num = False
        num_desc = "%d 个附件，" % len(items) + ("阿拉伯数字编号 1.～%d." % len(items) if ok_num else "编号不是“1. 2.”形式")
        if not ok_num:
            probs.append("多个附件未按“1. 2.”阿拉伯数字编号")
    names = [re.sub(r"^\d+[.．]\s*", "", t) for t in items]
    badp = [n for n in names if n and n[-1] in "。，、；：！？.,;:!?"]
    if badp:
        probs.append("附件名称末尾带标点（“%s”）" % badp[0][-12:])
    if not ok_ind:
        probs.append(ind_d)
    meas = "%s；%s；%s" % (ind_d, num_desc, "名称末尾无标点" if not badp else "名称末尾有标点")
    if probs:
        meas += "；不符：" + "、".join(probs)
    return ("pass" if not probs else "fail"), exp, meas


def _pb_before(doc, p):
    """附件标识之前是否有分页 -> (bool, 说明)。"""
    if p.ppr.get("pageBreakBefore"):
        return True, "pageBreakBefore"
    t = p.text
    first_vis = next((i for i, ch in enumerate(t) if not ch.isspace()), len(t))
    if any(pos <= first_vis for pos in p.pb):
        return True, "本段首部分页符"
    j = p.idx - 1
    steps = 0
    while j >= 0 and steps < 6:
        q = doc.paras[j]
        if q.in_txbx:
            j -= 1
            continue
        steps += 1
        if q.pb and not compact(q.text[max(q.pb):]):
            return True, "前段分页符"
        if q.sectpr is not None:
            typ = doc.sections[min(p.sect, len(doc.sections) - 1)]["type"]
            if typ in ("nextPage", "oddPage", "evenPage"):
                return True, "分节符(%s)" % typ
            return False, "仅连续分节符"
        if not q.empty:
            break
        j -= 1
    return False, "无分页"


def chk_11b(ctx):
    exp = "每个附件标识之前有分页（另页）"
    ms = ctx.R.attach_marks
    if not ms:
        return "na", exp, "无附件标识段落"
    res = []
    ok = True
    for m in ms:
        b, d = _pb_before(ctx.doc, m)
        ok = ok and b
        res.append("“%s”%s" % (m.stext, "另页(" + d + ")" if b else "未另页"))
    return ("pass" if ok else "fail"), exp, "；".join(res)


def chk_11c(ctx):
    exp = "黑体、三号16pt、顶格"
    ms = ctx.R.attach_marks
    if not ms:
        return "na", exp, "无附件标识段落"
    pr = [(m, [(0, len(m.text))]) for m in ms]
    fg, nea, sg, nall, fonts, sizes = _prop_check(ctx, pr, "黑体", (31, 33))
    probs = []
    if nea and fg / nea < 0.95:
        probs.append("字体")
    if nall and sg / nall < 0.95:
        probs.append("字号")
    for m in ms:
        ind = m.ppr.get("ind", {})
        if any(meas_to_twips(ind.get(k)) for k in ("left", "firstLine", "firstLineChars", "leftChars", "hanging")) \
                or m.raw_text.startswith(("　", "  ")):
            probs.append("未顶格")
            break
    meas = "字体 %s；字号 %s" % (share_str(fonts, nea) if nea else "无汉字", share_str(sizes, nall, fmt=lambda k: fmt_pt(k / 2.0)))
    meas += ctx.alias_note(fonts.keys())
    if probs:
        meas += "；不符：" + "、".join(sorted(set(probs)))
    return ("pass" if not probs else "fail"), exp, meas


# ====================================================================== 线条识别
def border_vis(b):
    """段落/表格边框 -> {'sz','rgb'} 或 None（nil/none/无宽度视为不可见）。"""
    if not b:
        return None
    val = (b.get("val") or "").lower()
    if val in ("", "nil", "none"):
        return None
    try:
        sz = int(float(b["sz"])) if b.get("sz") not in (None, "") else 4
    except Exception:
        sz = 4
    if sz <= 0:
        return None
    col = b.get("color")
    rgb = None if col in (None, "", "auto") else parse_color(col)
    return {"sz": sz, "rgb": rgb}


def _emu_pt(v):
    try:
        return float(v) / 12700.0
    except Exception:
        return None


def shape_lines(p_el):
    """段落内的线条形状（DrawingML 与 VML；不进入 mc:Fallback）-> [{'rgb','w'}]。"""
    out = []
    for e in walk_no_fallback(p_el):
        n = lname(e)
        ns = e.tag.split("}")[0][1:] if isinstance(e.tag, str) and "}" in e.tag else ""
        if n == "wsp" and ns == NS["wps"]:
            sp = e.find("{%s}spPr" % NS["wps"])
            if sp is None or e.find("{%s}txbx" % NS["wps"]) is not None:
                continue
            geom = sp.find("a:prstGeom", NS)
            prst = geom.get("prst") if geom is not None else None
            ext = sp.find("a:xfrm/a:ext", NS)
            cx = cy = None
            if ext is not None:
                try:
                    cx, cy = float(ext.get("cx")), float(ext.get("cy"))
                except Exception:
                    cx = cy = None
            if cx is None:
                anc = e
                while anc is not None and lname(anc) not in ("anchor", "inline"):
                    anc = anc.getparent()
                if anc is not None:
                    ex = anc.find("{%s}extent" % NS["wp"])
                    if ex is not None:
                        try:
                            cx, cy = float(ex.get("cx")), float(ex.get("cy"))
                        except Exception:
                            pass
            thin = cy is not None and cx is not None and cy <= 54000 and cx >= 10 * max(cy, 1)
            is_line = prst in ("line", "straightConnector1") or thin
            if not is_line:
                continue
            ln = sp.find("a:ln", NS)
            stroke = None
            width = None
            if ln is not None and ln.find("a:noFill", NS) is None:
                clr = ln.find("a:solidFill/a:srgbClr", NS)
                if clr is not None:
                    stroke = parse_color(clr.get("val"))
                width = _emu_pt(ln.get("w"))
            fill = None
            sf = sp.find("a:solidFill/a:srgbClr", NS)
            if sf is not None:
                fill = parse_color(sf.get("val"))
            rgb = stroke if (prst in ("line", "straightConnector1") or stroke is not None) else fill
            if width is None:
                width = (cy / 12700.0) if (thin and cy) else 0.75
            out.append({"rgb": rgb, "w": width})
        elif ns == NS["v"] and n in ("line", "rect", "shape", "roundrect"):
            if e.find("{%s}textbox" % NS["v"]) is not None:
                continue
            stroke = parse_color(e.get("strokecolor"))
            fillc = parse_color(e.get("fillcolor"))
            style = e.get("style") or ""
            if n != "line":
                m = re.search(r"height:\s*([\d.]+)\s*(pt|mm|cm|in|px)?", style)
                hpt = None
                if m:
                    f = {"pt": 1.0, "mm": 2.8346, "cm": 28.346, "in": 72.0, "px": 0.75, None: 1.0}[m.group(2)]
                    hpt = float(m.group(1)) * f
                is_hr = e.get("{%s}hr" % NS["o"]) is not None
                if not (is_hr or (hpt is not None and hpt <= 4.0)):
                    continue
                rgb = stroke if (stroke is not None and e.get("stroked") != "f") else (fillc or stroke)
            else:
                rgb = stroke
            wm = re.search(r"([\d.]+)\s*pt", e.get("strokeweight") or "")
            out.append({"rgb": rgb, "w": float(wm.group(1)) if wm else 0.75})
    return out


def _cell_border_sets(tr):
    """一行的单元格 [(top, bottom)] 边框字典（来自 tcBorders）。"""
    res = []
    for tc in tr.findall(T_TC):
        tcp = tc.find(qn("w:tcPr"))
        tb = tcp.find(qn("w:tcBorders")) if tcp is not None else None
        d = {lname(c): parse_border(c) for c in tb} if tb is not None else {}
        res.append(d)
    return res


def table_hlines(doc, tinfo):
    """表格水平线 [(位置键, {'sz','rgb'})]（含上/下边框与行间线）。"""
    tbl = tinfo["el"]
    merged = dict(doc.styles.table_borders(tinfo["sid"])) if tinfo["sid"] else {}
    tp = tbl.find(qn("w:tblPr"))
    if tp is not None:
        tb = tp.find(qn("w:tblBorders"))
        if tb is not None:
            merged.update({lname(c): parse_border(c) for c in tb})
    rows = tbl.findall(T_TR)
    n = len(rows)
    cells = [_cell_border_sets(r) for r in rows]
    out = []
    for i in range(n + 1):
        explicit = []
        total = 0
        if i > 0:
            explicit += [c.get("bottom") for c in cells[i - 1] if c.get("bottom")]
            total += len(cells[i - 1])
        if i < n:
            explicit += [c.get("top") for c in cells[i] if c.get("top")]
            total += len(cells[i])
        vis = [border_vis(b) for b in explicit]
        vis = [v for v in vis if v]
        best = max(vis, key=lambda v: v["sz"]) if vis else None
        if best is None:
            tl = merged.get("top" if i == 0 else ("bottom" if i == n else "insideH"))
            tv = border_vis(tl)
            if tv and (not explicit or len(explicit) < total):
                best = tv
        if best:
            if i == 0:
                key = tinfo["first"] - 0.5
            elif i == n:
                key = tinfo["last"] + 0.5
            else:
                key = ("tbl", tinfo["id"], i)
            out.append((key, dict(best, kind="表格边框")))
    # 细行底纹线：行高 <= 2mm 且整行单元格有底纹填充（用表格行画线的写法，视觉上就是一条线）
    for ri, tr in enumerate(rows):
        trp = tr.find(qn("w:trPr"))
        th = trp.find(qn("w:trHeight")) if trp is not None else None
        h = meas_to_twips(wa(th, "val")) if th is not None else None
        if h is None or h > 113:
            continue
        fills = []
        for tc in tr.findall(T_TC):
            tcp = tc.find(qn("w:tcPr"))
            shd = tcp.find(qn("w:shd")) if tcp is not None else None
            f = wa(shd, "fill") if shd is not None else None
            fills.append(parse_color(f) if f and f.lower() != "auto" else None)
        if fills and all(f is not None for f in fills):
            out.append((("shade", tinfo["id"], ri), {"sz": max(1, int(round(h / 20.0 * 8))), "rgb": fills[0], "kind": "表格底纹线"}))
    return out


def para_hlines(p):
    """段落上/下边框 -> [(位置键, {'sz','rgb'})]；位置键用 idx±0.5 以便相邻段落边框去重。

    另认一种常见的画线写法：没有文字、固定行高不超过 3pt、段落有底纹填充色的“细行”，
    视觉上就是一条水平线（线粗按行高算）。
    """
    out = []
    bd = p.ppr.get("pBdr", {})
    for side, off in (("top", -0.5), ("bottom", 0.5)):
        v = border_vis(bd.get(side))
        if v:
            out.append((p.idx + off, v))
    shd = p.ppr.get("shd")
    sp = p.ppr.get("spacing", {})
    line = meas_to_twips(sp.get("line"))
    if p.empty and shd and line is not None and 0 < line <= 60 and sp.get("lineRule") in ("exact", "atLeast"):
        val = (shd.get("val") or "").lower()
        col = shd.get("color") if val == "solid" else shd.get("fill")
        rgb = parse_color(col) if col and str(col).lower() != "auto" else None
        if val not in ("nil", "none") and rgb is not None and rgb != (255, 255, 255):
            out.append((p.idx, {"sz": max(1, int(round(line / 20.0 * 8))), "rgb": rgb, "kind": "段落底纹线"}))
    return out


def region_lines(doc, paras, tinfos):
    """在给定段落/表格范围内收集所有水平线（去重）-> list of dict(key, sz(磅*8), rgb, kind)。"""
    found = {}
    for p in paras:
        for key, v in para_hlines(p):
            if key not in found or v["sz"] > found[key]["sz"]:
                found[key] = dict(v, kind=v.get("kind", "段落边框"))
        for i, s in enumerate(shape_lines(p.el)):
            found[("shape", p.idx, i)] = {"sz": int((s["w"] or 0.75) * 8), "rgb": s["rgb"], "kind": "线条形状"}
        c = p.ctext
        if len(c) >= 8 and re.fullmatch(r"[—―─━_\-－=＿]+", c) and not p.in_table:
            chs = p.chars()
            found[("char", p.idx)] = {"sz": 8, "rgb": chs[0][1].color if chs else None, "kind": "字符线"}
    for t in tinfos:
        for key, v in table_hlines(doc, t):
            if key not in found or v["sz"] > found[key]["sz"]:
                found[key] = dict(v)
    return list(found.values())



# ====================================================================== 12 版头
def rgb_hex(c):
    return "自动/黑" if c is None else "%02X%02X%02X" % c


def chk_12a(ctx):
    exp = "红色、居中、方正小标宋简体"
    if ctx.layout not in OFFICE_LAYOUTS + ("纪要",):
        return "na", exp, "该版式无版头"
    ms = ctx.R.marks
    if not ms:
        return "fail", exp, "版头区未识别到发文机关标志/纪要标志"
    total = reds = nea = 0
    colors, fonts, sizes = Counter(), Counter(), Counter()
    for p in ms:
        for ch, s in chars_in(p, [(0, len(p.text))]):
            total += 1
            colors[rgb_hex(s.color)] += 1
            sizes[s.sz] += 1
            if is_red(s.color):
                reds += 1
            if is_ea_char(ch):
                nea += 1
                fonts[s.ea or NO_FONT] += 1
    fgood = sum(v for k, v in fonts.items() if ctx.canon(k) == "方正小标宋简体")
    al = Counter(p.jc() for p in ms)
    probs = []
    if total and reds / total < 0.95:
        probs.append("颜色不是红色")
    if any(k != "center" for k in al):
        probs.append("未居中")
    if nea and fgood / nea < 0.95:
        probs.append("字体不是方正小标宋简体")
    names = {"center": "居中", "left": "左对齐", "right": "右对齐", "both": "两端对齐", "distribute": "分散对齐"}
    meas = "“%s”：颜色 %s；对齐 %s；字体 %s；字号 %s" % (
        ms[0].short(14), "、".join(colors), "、".join(names.get(k, k) for k in al),
        "、".join(fonts) if fonts else "无汉字", "、".join(fmt_pt(k / 2.0) for k in sizes))
    meas += ctx.alias_note(fonts.keys())
    if probs:
        meas += "；不符：" + "、".join(probs)
    return ("pass" if not probs else "fail"), exp, meas


RE_FWZH_G = re.compile(r"(\S*?)([〔\[［（(【〖])\s*(\d{4})\s*([〕\]］）)】〗])\s*(第\s*)?(\d+)\s*号")


def chk_12b(ctx):
    exp = "〔YYYY〕N号：六角括号、四位年份、序号不编虚位、不加“第”、后有“号”"
    if ctx.layout not in OFFICE_LAYOUTS + ("纪要",):
        return "na", exp, "该版式无版头"
    R = ctx.R
    if R.fwzh is None:
        if ctx.layout == "纪要":
            return "na", exp, "纪要无发文字号"
        return "fail", exp, "版头区未找到发文字号"
    m = RE_FWZH_G.search(R.fwzh.stext)
    if not m:
        return "fail", exp, "“%s”格式无法解析" % R.fwzh.short(30)
    pre, ob, yr, cb, di, seq = m.groups()
    probs = []
    if ob != "〔" or cb != "〕":
        probs.append("括号为“%s%s”，应为六角括号“〔〕”" % (ob, cb))
    if not yr.isascii():
        probs.append("年份使用全角数字")
    if di:
        probs.append("加了“第”字")
    if seq.startswith("0"):
        probs.append("序号编虚位")
    if not seq.isascii():
        probs.append("序号使用全角数字")
    shown = re.sub(r"\s+", "", m.group(0))
    if probs:
        return "fail", exp, "“%s”：%s" % (shown, "、".join(probs))
    return "pass", exp, "“%s”" % shown


def _lead_spaces(p):
    n = 0
    for ch in p.text:
        if ch == "　":
            n += 1
        elif ch in " \t":
            continue
        else:
            break
    return n


def left_offset_pt(p):
    ind = p.ppr.get("ind", {})
    sz = p.dom_pt() or 16.0
    left, fl, hg = (meas_to_twips(ind.get(k)) or 0 for k in ("left", "firstLine", "hanging"))
    lc, flc, hgc = (meas_to_twips(ind.get(k)) or 0 for k in ("leftChars", "firstLineChars", "hangingChars"))
    off = ((lc / 100.0 * sz) if lc else left / 20.0) + ((flc / 100.0 * sz) if flc else fl / 20.0) \
        - ((hgc / 100.0 * sz) if hgc else hg / 20.0)
    return off + _lead_spaces(p) * sz


def right_offset_pt(ctx, p):
    ind = p.ppr.get("ind", {})
    sz = p.dom_pt() or 16.0
    r = meas_to_twips(ind.get("right")) or 0
    rc = meas_to_twips(ind.get("rightChars")) or 0
    off = (rc / 100.0 * sz) if rc else r / 20.0
    t = p.text.rstrip(" \t\r\n")
    trail = len(t) - len(t.rstrip("　"))
    off += trail * sz
    off_tab = None
    rights = [meas_to_twips(pos) for val, pos in p.ppr.get("tabs", []) if val in ("right", "end") and meas_to_twips(pos)]
    if rights and not p.tbl:
        sect = ctx.doc.sections[min(p.sect, len(ctx.doc.sections) - 1)]
        if sect["w"] and sect["left"] is not None and sect["right"] is not None:
            width = sect["w"] - sect["left"] - sect["right"]
            off_tab = (width - r - max(rights)) / 20.0 + trail * sz
    return off, off_tab


def chk_12c(ctx):
    R = ctx.R
    if ctx.layout in ("红头下行", "信函"):
        exp = "下行文发文字号居中"
        if R.fwzh is None:
            return "na", exp, "未找到发文字号，无法判定位置（已在 12b 扣分）"
        a = R.fwzh.jc()
        names = {"center": "居中", "left": "左对齐", "right": "右对齐", "both": "两端对齐"}
        return ("pass" if a == "center" else "fail"), exp, "发文字号%s" % names.get(a, a)
    if ctx.layout == "上行":
        exp = "上行文：发文字号居左空一字、签发人居右空一字，二者同一行"
        fw, iss = R.fwzh, R.issuer
        if fw is None:
            return "na", exp, "未找到发文字号，无法判定位置（已在 12b 扣分）"
        if iss is None:
            return "fail", exp, "版头区未找到签发人"
        same = fw is iss or (fw.tbl and iss.tbl and fw.tbl[0] == iss.tbl[0] and fw.tbl[1] == iss.tbl[1])
        probs, parts = [], []
        parts.append("发文字号与签发人%s" % ("同一行" if same else "不在同一行"))
        if not same:
            probs.append("未同行")
        lo = left_offset_pt(fw)
        parts.append("发文字号左空 %s" % fmt_pt(lo))
        if not (12 <= lo <= 20):
            probs.append("发文字号未左空一字")
        ro, ro_tab = right_offset_pt(ctx, iss)
        okr = 12 <= ro <= 20 or (ro_tab is not None and 12 <= ro_tab <= 20)
        parts.append("签发人右空 %s" % fmt_pt(ro_tab if (ro_tab is not None and 12 <= ro_tab <= 20) else ro))
        if not okr:
            probs.append("签发人未右空一字")
        meas = "；".join(parts)
        if probs:
            meas += "；不符：" + "、".join(probs)
        return ("pass" if not probs else "fail"), exp, meas
    return "na", "发文字号位置", "纪要/普通材料/讲话不检查"


def chk_12d(ctx):
    exp = "“签发人：”为仿宋_GB2312、姓名为楷体_GB2312（三号）"
    if ctx.layout != "上行":
        return "na", exp, "仅上行文检查"
    p = ctx.R.issuer
    if p is None:
        return "fail", exp, "版头区未找到签发人"
    t = p.text
    i = t.find("签发人")
    m = re.search(r"[：:]", t[i:]) if i >= 0 else None
    if i < 0 or not m:
        return "fail", exp, "签发人段落缺少“签发人：”标签"
    ce = i + m.end()
    name_end = len(t.rstrip(" \t\r\n　"))
    lab = chars_in(p, [(i, ce)])
    nm = chars_in(p, [(ce, name_end)])
    if not nm:
        return "fail", exp, "未找到签发人姓名"

    def stat(chs, want):
        fonts, sizes = Counter(), Counter()
        nea = 0
        for ch, s in chs:
            sizes[s.sz] += 1
            if is_ea_char(ch):
                nea += 1
                fonts[s.ea or NO_FONT] += 1
        fg = sum(v for k, v in fonts.items() if ctx.canon(k) == want)
        sg = sum(v for k, v in sizes.items() if 31 <= k <= 33)
        ok = (nea == 0 or fg / nea >= 0.95) and sg / max(len(chs), 1) >= 0.95
        return ok, fonts, sizes

    ok1, f1, s1 = stat(lab, "仿宋_GB2312")
    ok2, f2, s2 = stat(nm, "楷体_GB2312")
    meas = "标签：%s、%s；姓名：%s、%s" % (
        "/".join(f1) or "无", "/".join(fmt_pt(k / 2.0) for k in s1),
        "/".join(f2) or "无", "/".join(fmt_pt(k / 2.0) for k in s2))
    meas += ctx.alias_note(list(f1) + list(f2))
    if not ok1:
        meas += "；标签不符"
    if not ok2:
        meas += "；姓名不符"
    return ("pass" if ok1 and ok2 else "fail"), exp, meas


def chk_12e(ctx):
    exp = "版头存在红色分隔线（段落边框/线条形状/表格边框）"
    if ctx.layout not in OFFICE_LAYOUTS + ("纪要",):
        return "na", exp, "该版式无版头"
    R, doc = ctx.R, ctx.doc
    paras = [p for p in doc.paras if p.idx < R.hdr_end]
    if R.title:
        paras.append(R.title[0])
    tin = [t for t in doc.tables if t["first"] < R.hdr_end]
    lines = region_lines(doc, paras, tin)
    red = [l for l in lines if is_red(l["rgb"])]
    if red:
        kinds = Counter(l["kind"] for l in red)
        return "pass", exp, "识别到红色线 %d 条（%s）" % (len(red), "、".join("%s×%d" % kv for kv in kinds.items()))
    if lines:
        cols = Counter(rgb_hex(l["rgb"]) for l in lines)
        return "fail", exp, "版头有线但颜色不是红色：%s" % "、".join(cols)
    return "fail", exp, "版头区未发现任何分隔线"


# ====================================================================== 13 版记
def chk_13a(ctx):
    exp = "抄送/印发行：14pt（四号）仿宋_GB2312"
    if ctx.layout not in OFFICE_LAYOUTS:
        return "na", exp, "该版式不检查版记"
    rec = ctx.R.record
    if not rec:
        return "na", exp, "无版记（未发现“抄送：”或“…印发”段落）"
    fonts, sizes = Counter(), Counter()
    nea = nall = 0
    for p in rec:
        for ch, s in chars_in(p, [(0, len(p.text))]):
            nall += 1
            sizes[s.sz] += 1
            if is_ea_char(ch):
                nea += 1
                fonts[s.ea or NO_FONT] += 1
    fg = sum(v for k, v in fonts.items() if ctx.canon(k) == "仿宋_GB2312")
    sg = sum(v for k, v in sizes.items() if 27 <= k <= 29)
    ok = (nea == 0 or fg / nea >= 0.95) and nall and sg / nall >= 0.95
    meas = "字体 %s；字号 %s" % (share_str(fonts, nea) if nea else "无汉字",
                              share_str(sizes, nall, fmt=lambda k: fmt_pt(k / 2.0)))
    meas += ctx.alias_note(fonts.keys())
    return ("pass" if ok else "fail"), exp, meas


def chk_13b(ctx):
    exp = "版记上下各一条粗线、抄送与印发间一条细线（识别≥2条为通过）"
    if ctx.layout not in OFFICE_LAYOUTS:
        return "na", exp, "该版式不检查版记"
    R, doc = ctx.R, ctx.doc
    if not R.record:
        return "na", exp, "无版记（未发现“抄送：”或“…印发”段落）"
    rs = R.record_start
    paras = [p for p in doc.paras if p.idx >= rs - 1]
    tin = [t for t in doc.tables if t["last"] >= rs - 1 and t["first"] >= 0 and t["last"] >= rs]
    lines = region_lines(doc, paras, tin)
    n = len(lines)
    thick = sum(1 for l in lines if l["sz"] >= 8)
    thin = n - thick
    kinds = Counter(l["kind"] for l in lines)
    meas = "识别到 %d 条线（粗 %d、细 %d）" % (n, thick, thin)
    if n:
        meas += "；来源：" + "、".join("%s×%d" % kv for kv in kinds.items())
    return ("pass" if n >= 2 else "fail"), exp, meas


# ====================================================================== 14 正文颜色
def chk_14(ctx):
    exp = "黑色或自动（≥95%）"
    chs = body_chars(ctx)
    if not chs:
        return "na", exp, "无正文段落"
    c = Counter()
    good = 0
    for _, s in chs:
        if is_black(s.color):
            good += 1
        else:
            c[rgb_hex(s.color)] += 1
    ok = good / len(chs) >= 0.95
    meas = "黑色/自动占 %s" % pct(good / len(chs))
    if c:
        meas += "；其他颜色：" + "、".join("%s(%d字)" % kv for kv in c.most_common(4))
    return ("pass" if ok else "fail"), exp, meas


# ====================================================================== 15 字体合规
def font_census(ctx):
    """正文区与页眉页脚区的东亚字体字符数（原名）。"""
    body = Counter()
    for p in ctx.doc.paras:
        for ch, s in p.chars():
            if is_ea_char(ch):
                body[s.ea or NO_FONT] += 1
    hf = Counter()
    for part in ctx.doc.all_hf_parts():
        for p in ctx.doc.part_paras(part):
            for ch, s in p.chars():
                if is_ea_char(ch):
                    hf[s.ea or NO_FONT] += 1
    return body, hf


def chk_15(ctx):
    exp = "东亚字体仅限 仿宋_GB2312/楷体_GB2312/黑体/方正小标宋简体（页脚另允许宋体）"
    body, hf = ctx.census
    badb = {k: v for k, v in body.items() if ctx.canon(k) not in STD_FONTS}
    badf = {k: v for k, v in hf.items() if ctx.canon(k) not in STD_FONTS and ctx.canon(k) != "宋体"}
    if not body and not hf:
        return "na", exp, "文档中没有汉字"
    parts = []
    if badb:
        parts.append("正文区非标准字体：" + "、".join("%s(%d字)" % kv for kv in sorted(badb.items(), key=lambda kv: -kv[1])))
    if badf:
        parts.append("页眉页脚非标准字体：" + "、".join("%s(%d字)" % kv for kv in sorted(badf.items(), key=lambda kv: -kv[1])))
    if parts:
        return "fail", exp, "；".join(parts) + ctx.alias_note(list(body) + list(hf))
    used = "、".join("%s(%d字)" % kv for kv in sorted(body.items(), key=lambda kv: -kv[1]))
    return "pass", exp, "正文区字体均合规：" + (used or "无汉字") + ctx.alias_note(list(body) + list(hf))



# ====================================================================== 保真比对
def doc_text(doc):
    """文档全部文字（正文区含表格、文本框；不含页眉页脚；含自动编号序号；接受修订）。"""
    return "".join(p.text for p in doc.paras)


def _split_sentences(t):
    return re.findall(r"[^。！？；]*[。！？；]|[^。！？；]+$", t)


def _opcodes_large(a, b):
    """长文本（任一方 > 30000 字）的快速比对：先按句对齐，再对不相等的句块做字符级比对。
    返回 (opcodes[(tag,i1,i2,j1,j2)], 匹配字符数)。ratio 为近似值（与整体 SequenceMatcher 可能有微小差异）。"""
    sa, sb = _split_sentences(a), _split_sentences(b)
    pa, pb = [0], [0]
    for x in sa:
        pa.append(pa[-1] + len(x))
    for x in sb:
        pb.append(pb[-1] + len(x))
    ops, matches = [], 0
    for tag, i1, i2, j1, j2 in difflib.SequenceMatcher(None, sa, sb, autojunk=False).get_opcodes():
        a0, a1, b0, b1 = pa[i1], pa[i2], pb[j1], pb[j2]
        if tag == "equal":
            matches += a1 - a0
            ops.append(("equal", a0, a1, b0, b1))
            continue
        A, B = a[a0:a1], b[b0:b1]
        if len(A) * len(B) > 4_000_000:
            ops.append(("replace" if A and B else ("delete" if A else "insert"), a0, a1, b0, b1))
            continue
        for t2, x1, x2, y1, y2 in difflib.SequenceMatcher(None, A, B, autojunk=False).get_opcodes():
            if t2 == "equal":
                matches += x2 - x1
            ops.append((t2, a0 + x1, a0 + x2, b0 + y1, b0 + y2))
    return ops, matches


def fidelity_report(doc, ref_text):
    a = compact(ref_text)
    b = compact(doc_text(doc))
    if a == b:
        return {"ratio": 1.0, "ref_chars": len(a), "doc_chars": len(b), "exact": True, "diffs": []}
    if max(len(a), len(b)) > 30000:
        ops, m = _opcodes_large(a, b)
        ratio = 2.0 * m / (len(a) + len(b)) if (len(a) + len(b)) else 1.0
    else:
        sm = difflib.SequenceMatcher(None, a, b, autojunk=False)
        ratio = sm.ratio()
        ops = sm.get_opcodes()
    diffs = []
    for tag, i1, i2, j1, j2 in ops:
        if tag == "equal":
            continue
        diffs.append({
            "op": tag, "ref": a[i1:i2], "doc": b[j1:j2],
            "context": "%s【%s→%s】%s" % (a[max(0, i1 - 10):i1], a[i1:i2], b[j1:j2], a[i2:i2 + 10]),
        })
        if len(diffs) >= 20:
            break
    return {"ratio": round(ratio, 4), "ref_chars": len(a), "doc_chars": len(b), "exact": False, "diffs": diffs}


# ====================================================================== 检查项清单
ITEMS = [
    ("01", "纸张 A4", "必需", chk_01, "A4 210×297mm（±1mm）"),
    ("02", "页边距", "必需", chk_02, "上37/左28/右26/下35mm（±1.5mm）"),
    ("03a", "正文字体", "必需", chk_03a, "仿宋_GB2312（≥95%）"),
    ("03b", "正文字号", "必需", chk_03b, "三号 16pt（±0.5pt，≥95%）"),
    ("04", "行距", "重要", chk_04, "每页22行：固定值28～30pt，或网格22行+单倍行距（≥80%段落）"),
    ("05", "首行缩进两字", "重要", chk_05, "首行缩进两字（约32pt/200字符，≥90%段落）"),
    ("06", "标题", "必需", chk_06, "二号22pt、方正小标宋简体、居中"),
    ("07a", "一级标题", "重要", chk_07a, "黑体、三号16pt"),
    ("07b", "二级标题", "重要", chk_07b, "楷体_GB2312、三号16pt"),
    ("07c", "三、四级标题", "重要", chk_07c, "仿宋_GB2312、三号16pt"),
    ("08", "主送机关顶格", "重要", chk_08, "顶格（无首行缩进、无左缩进）"),
    ("09a", "页脚含 PAGE 域", "重要", chk_09a, "页脚含 PAGE 域"),
    ("09b", "页码字体字号", "重要", chk_09b, "四号14pt、东亚字体宋体"),
    ("09c", "页码一字线", "重要", chk_09c, "页码两侧各有一字线，形如“— 1 —”"),
    ("09d", "页码单右双左", "重要", chk_09d, "奇数页页码居右、偶数页居左（需设奇偶页不同）"),
    ("09e", "信函首页无页码", "重要", chk_09e, "信函首页不显示页码（titlePg 且首页页脚无 PAGE 域）"),
    ("10", "成文日期", "重要", chk_10, "阿拉伯数字“YYYY年M月D日”，年份四位、月日不编虚位"),
    ("11a", "附件说明", "重要", chk_11a, "附件说明：左空二字；多个附件用阿拉伯数字编号；名称末尾不加标点"),
    ("11b", "附件另页", "重要", chk_11b, "每个附件标识之前有分页（另页）"),
    ("11c", "附件标识", "重要", chk_11c, "黑体、三号16pt、顶格"),
    ("12a", "发文机关标志/纪要标志", "必需", chk_12a, "红色、居中、方正小标宋简体"),
    ("12b", "发文字号格式", "必需", chk_12b, "〔YYYY〕N号：六角括号、四位年份、序号不编虚位、不加“第”、后有“号”"),
    ("12c", "发文字号位置", "重要", chk_12c, "下行文居中；上行文发文字号居左空一字、签发人居右空一字且同行"),
    ("12d", "签发人字体", "重要", chk_12d, "“签发人：”为仿宋_GB2312、姓名为楷体_GB2312（三号）"),
    ("12e", "红色分隔线", "重要", chk_12e, "版头存在红色分隔线（段落边框/线条形状/表格边框）"),
    ("13a", "版记字体字号", "次要", chk_13a, "抄送/印发行：14pt（四号）仿宋_GB2312"),
    ("13b", "版记分隔线", "次要", chk_13b, "版记上下各一条粗线、抄送与印发间一条细线（识别≥2条为通过）"),
    ("14", "正文颜色", "次要", chk_14, "黑色或自动（≥95%）"),
    ("15", "字体合规总表", "必需", chk_15, "东亚字体仅限 仿宋_GB2312/楷体_GB2312/黑体/方正小标宋简体（页脚另允许宋体）"),
]
LEVEL_W = {"必需": 3, "重要": 2, "次要": 1}


def audit_docx(path, layout=None, ref_text=None, strict_font_names=False, with_roles=False):
    """审计一份 docx，返回结果字典。打不开（或遇到无法预料的内部错误）时返回 {"error": "..."}。"""
    holder = {}
    try:
        return _audit_impl(path, layout, ref_text, strict_font_names, with_roles, holder)
    except Exception as e:  # noqa  兜底：任何未预料的异常都不让脚本崩溃
        return {"error": "审计内部错误：%s: %s" % (type(e).__name__, e)}
    finally:
        z = holder.get("zf")
        if z is not None:
            try:
                z.close()  # 及时释放文件句柄，避免调用方随后移动/删除 docx 时被占用
            except Exception:  # noqa
                pass


def _audit_impl(path, layout, ref_text, strict_font_names, with_roles, holder):
    try:
        doc = Doc(path)
        holder["zf"] = doc.pkg.zf
    except zipfile.BadZipFile:
        return {"error": "无法打开：不是有效的 docx（zip 结构损坏或不是 docx 文件）：%s" % os.path.basename(str(path))}
    except FileNotFoundError:
        return {"error": "文件不存在：%s" % path}
    except Exception as e:  # noqa
        return {"error": "无法解析 docx：%s: %s" % (type(e).__name__, e)}
    notes = []
    lay = norm_layout(layout)
    if layout is not None and lay is None:
        return {"error": "不支持的版式：%s（可选：%s）" % (layout, "、".join(LAYOUTS))}
    source = "指定"
    if lay is None:
        source = "自动"
        try:
            lay = detect_layout(doc)
        except Exception as e:  # noqa
            lay = "普通材料"
            notes.append("自动判定版式异常（%s），按普通材料处理" % e)
        notes.append("版式为自动判定（%s）；自动判定无法识别信函版式，必要时请用 --layout 指定" % lay)
    try:
        roles = assign_roles(doc, lay)
    except Exception as e:  # noqa
        roles = Roles()
        roles.layout = lay
        for p in doc.paras:
            p.role = "空段" if p.empty else "正文"
            if p.role == "正文":
                roles.body.append(p)
        notes.append("段落角色识别异常（%s: %s），已退化为全文按正文处理，相关项可能失真" % (type(e).__name__, e))
    ctx = Ctx(doc, roles, lay, bool(strict_font_names))
    try:
        ctx.census = font_census(ctx)
    except Exception:  # noqa
        ctx.census = (Counter(), Counter())
    items = []
    for iid, name, level, fn, default_exp in ITEMS:
        try:
            st, exp, meas = fn(ctx)
        except Exception as e:  # noqa
            st, exp, meas = "na", default_exp, "检查异常：%s: %s" % (type(e).__name__, e)
        items.append({"id": iid, "name": name, "level": level, "status": st, "expected": exp, "measured": meas})
    wp = wt = 0
    cnt = Counter()
    for it in items:
        cnt[it["status"]] += 1
        if it["status"] != "na":
            w = LEVEL_W[it["level"]]
            wt += w
            if it["status"] == "pass":
                wp += w
    score = round(100.0 * wp / wt, 1) if wt else None
    body_f, hf_f = ctx.census
    nonstd = [k for k, _ in sorted(body_f.items(), key=lambda kv: -kv[1]) if ctx.canon(k) not in STD_FONTS]
    for k, _ in sorted(hf_f.items(), key=lambda kv: -kv[1]):
        if ctx.canon(k) not in STD_FONTS and ctx.canon(k) != "宋体" and k not in nonstd:
            nonstd.append(k)
    rs = Counter()
    for p in doc.paras:
        if p.role and p.role != "空段":
            rs[p.role] += 1
    if len(doc.sections) > 1:
        notes.append("文档含 %d 个节：纸张/页边距以首节为准，页码/行距等以正文所在节（第 %d 节）为准" % (
            len(doc.sections), ctx.body_sect + 1))
    if strict_font_names:
        notes.append("--strict-font-names：英文字体别名（如 FangSong_GB2312、SimHei）按不合规处理")
    elif ctx.alias_used:
        notes.append("采用英文别名视为等同：" + "、".join("%s→%s" % kv for kv in sorted(ctx.alias_used.items())))
    notes.extend(roles.notes)
    if roles.title_how and roles.title_how != "居中长段":
        notes.append("标题按“%s”识别" % roles.title_how)
    if not roles.title:
        notes.append("未识别到标题，版头区边界按首个非红色左对齐段落估计")
    res = {
        "file": os.path.basename(str(path)),
        "layout_used": lay,
        "layout_source": source,
        "items": items,
        "score": score,
        "score_detail": {"weighted_pass": wp, "weighted_total": wt, "pass": cnt["pass"], "fail": cnt["fail"], "na": cnt["na"]},
        "fonts_used": dict(sorted(body_f.items(), key=lambda kv: -kv[1])),
        "fonts_used_footer": dict(sorted(hf_f.items(), key=lambda kv: -kv[1])),
        "fonts_nonstandard": nonstd,
        "fidelity": None,
        "roles_summary": dict(rs),
        "notes": notes,
    }
    if ref_text is not None:
        try:
            res["fidelity"] = fidelity_report(doc, ref_text)
        except Exception as e:  # noqa
            res["notes"].append("保真比对异常：%s: %s" % (type(e).__name__, e))
    fid = res.get("fidelity")
    if fid and max(fid["ref_chars"], fid["doc_chars"]) > 30000:
        res["notes"].append("文本超过 3 万字，保真比对改用分句对齐后再逐字符比对，ratio 为近似值")
    if with_roles:
        rl = []
        for p in doc.paras:
            if p.empty:
                continue
            ea = Counter(s.ea for _, s in p.chars() if is_ea_char(_))
            ea_name = ea.most_common(1)[0][0] if ea else None
            rl.append({"idx": p.idx, "role": p.role, "text": p.stext[:24], "in_table": bool(p.in_table),
                       "font": ea_name, "size": p.dom_pt() or None, "align": p.jc()})
        res["roles"] = rl
    return res


# ====================================================================== 文本输出
def dispw(s):
    import unicodedata
    return sum(2 if unicodedata.east_asian_width(c) in ("W", "F") else 1 for c in s)


def padw(s, w):
    return s + " " * max(0, w - dispw(s))


def trunc(s, w):
    s = str(s).replace("\n", " ")
    if dispw(s) <= w:
        return s
    out, n = "", 0
    import unicodedata
    for c in s:
        cw = 2 if unicodedata.east_asian_width(c) in ("W", "F") else 1
        if n + cw > w - 1:
            break
        out += c
        n += cw
    return out + "…"


def render_text(res):
    if "error" in res:
        return "错误：%s" % res["error"]
    lines = []
    lines.append("文件：%s    版式：%s（%s）" % (res["file"], res["layout_used"], res["layout_source"]))
    st = {"pass": "通过", "fail": "未过", "na": "不适用"}
    head = "%s %s %s %s %s %s" % (padw("编号", 4), padw("级别", 4), padw("状态", 6), padw("名称", 22), padw("期望", 40), "实测")
    lines.append(head)
    lines.append("-" * 100)
    for it in res["items"]:
        lines.append("%s %s %s %s %s %s" % (
            padw(it["id"], 4), padw(it["level"], 4), padw(st[it["status"]], 6), padw(trunc(it["name"], 22), 22),
            padw(trunc(it["expected"], 40), 40), trunc(it["measured"], 110)))
    lines.append("-" * 100)
    d = res["score_detail"]
    sc = "无可评项" if res["score"] is None else "%.1f" % res["score"]
    lines.append("总分：%s（加权通过 %d/%d；通过 %d 项、未过 %d 项、不适用 %d 项）" % (
        sc, d["weighted_pass"], d["weighted_total"], d["pass"], d["fail"], d["na"]))
    if res["fonts_used"]:
        lines.append("正文区东亚字体：" + "、".join("%s×%d" % kv for kv in res["fonts_used"].items()))
    if res["fonts_used_footer"]:
        lines.append("页眉页脚东亚字体：" + "、".join("%s×%d" % kv for kv in res["fonts_used_footer"].items()))
    if res["fonts_nonstandard"]:
        lines.append("非标准字体：" + "、".join(res["fonts_nonstandard"]))
    fid = res.get("fidelity")
    if fid:
        lines.append("保真：相似度 %.4f（原文 %d 字，文档 %d 字）%s" % (
            fid["ratio"], fid["ref_chars"], fid["doc_chars"], "，逐字一致" if fid["exact"] else ""))
        for df in fid["diffs"][:20]:
            lines.append("  [%s] %s" % (df["op"], df["context"]))
    if res.get("roles_summary"):
        lines.append("角色统计：" + "、".join("%s×%d" % kv for kv in res["roles_summary"].items()))
    for n in res.get("notes", []):
        lines.append("说明：" + n)
    if "roles" in res:
        lines.append("")
        lines.append("段落角色（调试）：")
        for r in res["roles"]:
            lines.append("  #%d %s%s | %s | %s | %s | %s" % (
                r["idx"], r["role"], "[表]" if r["in_table"] else "", r["text"], r["font"], r["size"], r["align"]))
    return "\n".join(lines)


def read_ref(path):
    raw = open(path, "rb").read()
    for enc in ("utf-8-sig", "gb18030"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", errors="replace")


def main(argv=None):
    for s in (sys.stdout, sys.stderr):
        try:
            s.reconfigure(encoding="utf-8")
        except Exception:  # noqa
            pass
    ap = argparse.ArgumentParser(description="GB/T 9704-2012 公文版式审计（中立实现）")
    ap.add_argument("docx", help="待审计的 .docx 文件")
    ap.add_argument("--layout", help="版式：" + "|".join(LAYOUTS) + "（缺省自动判断）")
    ap.add_argument("--ref", help="原文文本文件（UTF-8），用于保真比对")
    ap.add_argument("--json", action="store_true", help="输出 JSON")
    ap.add_argument("--strict-font-names", action="store_true", help="英文字体别名按不合规处理")
    ap.add_argument("--roles", action="store_true", help="输出段落角色识别结果（调试）")
    a = ap.parse_args(argv)
    ref = None
    if a.ref:
        try:
            ref = read_ref(a.ref)
        except Exception as e:  # noqa
            print(json.dumps({"error": "无法读取原文文件：%s" % e}, ensure_ascii=False))
            return 2
    res = audit_docx(a.docx, layout=a.layout, ref_text=ref, strict_font_names=a.strict_font_names, with_roles=a.roles)
    if "error" in res:
        print(json.dumps({"error": res["error"]}, ensure_ascii=False))
        return 2
    if a.json:
        print(json.dumps(res, ensure_ascii=False, indent=2))
    else:
        print(render_text(res))
    return 0


if __name__ == "__main__":
    sys.exit(main())
