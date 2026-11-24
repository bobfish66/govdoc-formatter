#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""按 GB/T 9704-2012《党政机关公文格式》规范化 Word 文档。

用法示例：
    python format_govdoc.py 报告.docx                     # 输出 报告_国标版.docx
    python format_govdoc.py 报告.doc -o 输出.docx          # .doc 自动转换
    python format_govdoc.py 报告.docx --check              # 只体检不写文件
    python format_govdoc.py 报告.docx --title "关于XX的通知" --date "2026年11月24日"
    python format_govdoc.py 报告.docx --no-page-number     # 不加页码
"""
import argparse
import os
import re
import subprocess
import sys
import tempfile

try:
    from docx import Document
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Mm, Pt
except ImportError:
    sys.exit("缺少 python-docx，请先执行：pip install python-docx")

# ---------------- 国标参数（GB/T 9704-2012 第 4、5、7 章） ----------------
A4_W_MM, A4_H_MM = 210, 297
M_TOP, M_BOTTOM, M_LEFT, M_RIGHT = 37, 35, 28, 26     # 天头/下白边/订口/右白边，版心 156mm×225mm
FOOTER_DIST_MM = 28                                    # 页码位于版心下边缘 7mm 处
BODY_PT = 16.0                                         # 3 号字
BODY_LINE_PT = 28.99                                   # 固定行距，保证每面 22 行
TITLE_PT = 22.0                                        # 2 号字
PAGE_PT = 14.0                                         # 4 号字
TABLE_PT = 12.0                                        # 小四
BODY_EA = "仿宋_GB2312"
H1_EA = "黑体"
H2_EA = "楷体_GB2312"
TITLE_EA = "方正小标宋简体"
PAGE_EA = "宋体"
DEFAULT_LATIN = "Times New Roman"

RE_H1 = re.compile(r"^[一二三四五六七八九十]{1,3}、")
RE_H2 = re.compile(r"^[（(][一二三四五六七八九十]{1,3}[）)]")
RE_H3 = re.compile(r"^\d{1,2}[.．](?!\d)")
RE_H4 = re.compile(r"^[（(]\d{1,2}[）)]")
RE_ATTACH = re.compile(r"^附件\s*[:：]")
RE_DATE = re.compile(r"^(?:\d{4}年\d{1,2}月\d{1,2}日|[〇零○一二三四五六七八九]{2,4}年.{1,4}月.{1,4}日)$")

_RPR_AFTER_SPACING = [
    "w:w", "w:kern", "w:position", "w:sz", "w:szCs", "w:highlight", "w:u",
    "w:effect", "w:bdr", "w:shd", "w:fitText", "w:vertAlign", "w:rtl", "w:cs",
    "w:em", "w:lang", "w:eastAsianLayout", "w:specVanish", "w:oMath",
]


# ---------------- 底层工具 ----------------
def _insert_char_spacing(rPr, val_twips):
    """按 schema 顺序插入 w:spacing（字符间距）。"""
    el = OxmlElement("w:spacing")
    el.set(qn("w:val"), str(val_twips))
    for child in list(rPr):
        for tag in _RPR_AFTER_SPACING:
            if child.tag == qn(tag):
                child.addprevious(el)
                return
    rPr.append(el)


def set_run(run, ea_font, size_pt, bold=False, latin=None, char_spacing_pt=None):
    latin = latin or DEFAULT_LATIN
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    run.font.italic = False
    run.font.name = latin
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    rFonts.set(qn("w:ascii"), latin)
    rFonts.set(qn("w:hAnsi"), latin)
    rFonts.set(qn("w:eastAsia"), ea_font)
    for attr in ("w:asciiTheme", "w:hAnsiTheme", "w:eastAsiaTheme", "w:cstheme"):
        if rFonts.get(qn(attr)) is not None:
            del rFonts.attrib[qn(attr)]
    if char_spacing_pt:
        _insert_char_spacing(rPr, int(round(char_spacing_pt * 20)))


def fmt_para(p, indent_chars=2, align=WD_ALIGN_PARAGRAPH.JUSTIFY, before=0.0,
             after=0.0, line_pt=BODY_LINE_PT, right_chars=0, left_chars=0,
             keep_next=False):
    pf = p.paragraph_format
    pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    pf.line_spacing = Pt(line_pt)
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    pf.alignment = align
    pf.keep_with_next = keep_next
    pf.widow_control = True
    # 用 API 保证 w:ind 位置合法，再补中文度量属性
    pf.first_line_indent = Pt(indent_chars * BODY_PT)
    pf.left_indent = Pt(left_chars * BODY_PT)
    pf.right_indent = Pt(right_chars * BODY_PT)
    ind = p._p.get_or_add_pPr().find(qn("w:ind"))
    if ind is not None:
        for attr in ("w:firstLineChars", "w:leftChars", "w:rightChars"):
            if ind.get(qn(attr)) is not None:
                del ind.attrib[qn(attr)]
        if indent_chars:
            ind.set(qn("w:firstLineChars"), str(int(indent_chars * 100)))
        if left_chars:
            ind.set(qn("w:leftChars"), str(int(left_chars * 100)))
        if right_chars:
            ind.set(qn("w:rightChars"), str(int(right_chars * 100)))


def restyle(p, ea_font, size_pt=BODY_PT, bold=False, char_spacing_pt=None):
    """段落内所有 run 统一字体（含无 run 时补写空 run 属性）。"""
    runs = p.runs
    if not runs:
        return
    for r in runs:
        set_run(r, ea_font, size_pt, bold=bold, char_spacing_pt=char_spacing_pt)


# ---------------- 页面设置 ----------------
def setup_section(sec):
    sec.page_width = Mm(A4_W_MM)
    sec.page_height = Mm(A4_H_MM)
    sec.top_margin = Mm(M_TOP)
    sec.bottom_margin = Mm(M_BOTTOM)
    sec.left_margin = Mm(M_LEFT)
    sec.right_margin = Mm(M_RIGHT)
    sec.header_distance = Mm(15)
    sec.footer_distance = Mm(FOOTER_DIST_MM)


def build_footer(footer, right_align=True):
    footer.is_linked_to_previous = False
    p = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
    for r in list(p.runs):
        r._element.getparent().remove(r._element)
    pf = p.paragraph_format
    pf.line_spacing_rule = WD_LINE_SPACING.SINGLE
    pf.first_line_indent = Pt(0)
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.alignment = WD_ALIGN_PARAGRAPH.RIGHT if right_align else WD_ALIGN_PARAGRAPH.LEFT
    if right_align:
        pf.right_indent = Pt(BODY_PT)
        pf.left_indent = Pt(0)
    else:
        pf.left_indent = Pt(BODY_PT)
        pf.right_indent = Pt(0)
    r1 = p.add_run("— ")
    set_run(r1, PAGE_EA, PAGE_PT, latin=PAGE_EA)
    p._p.append(_mk_page_field())
    r2 = p.add_run(" —")
    set_run(r2, PAGE_EA, PAGE_PT, latin=PAGE_EA)


def _mk_page_field():
    fld = OxmlElement("w:fldSimple")
    fld.set(qn("w:instr"), " PAGE \\* MERGEFORMAT ")
    r = OxmlElement("w:r")
    rPr = OxmlElement("w:rPr")
    rFonts = OxmlElement("w:rFonts")
    for a in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rFonts.set(qn(a), PAGE_EA)
    rPr.append(rFonts)
    sz = OxmlElement("w:sz")
    sz.set(qn("w:val"), str(int(PAGE_PT * 2)))
    rPr.append(sz)
    szcs = OxmlElement("w:szCs")
    szcs.set(qn("w:val"), str(int(PAGE_PT * 2)))
    rPr.append(szcs)
    r.append(rPr)
    t = OxmlElement("w:t")
    t.text = "1"
    t.set(qn("xml:space"), "preserve")
    r.append(t)
    fld.append(r)
    return fld


def enable_odd_even(doc):
    settings = doc.settings.element
    if settings.find(qn("w:evenAndOddHeaders")) is None:
        settings.append(OxmlElement("w:evenAndOddHeaders"))


# ---------------- 表格 ----------------
def format_table(tbl, header_bold=True):
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, row in enumerate(tbl.rows):
        trPr = row._tr.get_or_add_trPr()
        if trPr.find(qn("w:cantSplit")) is None:
            trPr.append(OxmlElement("w:cantSplit"))
        if i == 0 and trPr.find(qn("w:tblHeader")) is None:
            trPr.append(OxmlElement("w:tblHeader"))
        for cell in row.cells:
            for p in cell.paragraphs:
                pf = p.paragraph_format
                pf.line_spacing_rule = WD_LINE_SPACING.SINGLE
                pf.space_before = Pt(0)
                pf.space_after = Pt(0)
                pf.first_line_indent = Pt(0)
                pf.left_indent = Pt(0)
                pf.alignment = WD_ALIGN_PARAGRAPH.CENTER if i == 0 else WD_ALIGN_PARAGRAPH.LEFT
                for r in p.runs:
                    set_run(r, "宋体", TABLE_PT, bold=(header_bold and i == 0), latin="Times New Roman")


# ---------------- 主处理 ----------------
def classify(text):
    if not text:
        return "blank"
    if RE_H1.match(text):
        return "h1"
    if RE_H2.match(text):
        return "h2"
    if RE_H3.match(text):
        return "h3"
    if RE_H4.match(text):
        return "h4"
    if RE_ATTACH.match(text):
        return "attach"
    if RE_DATE.match(text):
        return "date"
    return "body"


def drop_extra_blank(doc):
    removed = 0
    prev_blank = False
    for p in list(doc.paragraphs):
        is_blank = not p.text.strip()
        if is_blank and prev_blank:
            p._p.getparent().remove(p._p)
            removed += 1
            continue
        prev_blank = is_blank
    return removed


def normalize(doc, args, log):
    for sec in doc.sections:
        setup_section(sec)
    log.append("页面：A4，天头 37mm / 下白边 35mm / 订口 28mm / 右白边 26mm，版心 156mm×225mm")

    removed = drop_extra_blank(doc)
    if removed:
        log.append("清理连续空段落 %d 个" % removed)

    paras = doc.paragraphs
    title_idx = None
    for i, p in enumerate(paras):
        if p.text.strip():
            title_idx = i
            break

    counts = {"title": 0, "h1": 0, "h2": 0, "h3": 0, "h4": 0, "body": 0, "attach": 0, "date": 0, "blank": 0}

    for i, p in enumerate(paras):
        text = p.text.strip()
        kind = classify(text)
        if i == title_idx and kind not in ("h1", "h2"):
            kind = "title"

        if kind == "title":
            fmt_para(p, indent_chars=0, align=WD_ALIGN_PARAGRAPH.CENTER, line_pt=BODY_LINE_PT, keep_next=True)
            restyle(p, TITLE_EA, TITLE_PT)
            # 标题回行词意完整：禁止段中强制断行
            p.paragraph_format.keep_together = True
        elif kind == "h1":
            fmt_para(p, indent_chars=2, align=WD_ALIGN_PARAGRAPH.LEFT, before=0, after=0, keep_next=True)
            restyle(p, H1_EA)
        elif kind == "h2":
            fmt_para(p, indent_chars=2, align=WD_ALIGN_PARAGRAPH.LEFT, keep_next=True)
            restyle(p, H2_EA)
        elif kind in ("h3", "h4"):
            fmt_para(p, indent_chars=2, align=WD_ALIGN_PARAGRAPH.LEFT, keep_next=True)
            restyle(p, BODY_EA)
        elif kind == "attach":
            fmt_para(p, indent_chars=0, left_chars=2, align=WD_ALIGN_PARAGRAPH.LEFT)
            restyle(p, BODY_EA)
        elif kind == "date":
            fmt_para(p, indent_chars=0, right_chars=4, align=WD_ALIGN_PARAGRAPH.RIGHT)
            restyle(p, BODY_EA)
        else:
            fmt_para(p, indent_chars=2 if text else 0,
                     align=WD_ALIGN_PARAGRAPH.JUSTIFY if text else WD_ALIGN_PARAGRAPH.LEFT)
            if text:
                restyle(p, BODY_EA, char_spacing_pt=args.char_spacing)
        counts[kind if kind in counts else "body"] += 1

    for tbl in doc.tables:
        format_table(tbl)

    if not args.no_page_number:
        enable_odd_even(doc)
        for sec in doc.sections:
            build_footer(sec.footer, right_align=True)
            build_footer(sec.even_page_footer, right_align=False)
        log.append("页码：4 号半角宋体，奇数页居右空一字、偶数页居左空一字，形式 — 1 —，距版心下边缘 7mm")

    log.append(
        "段落：正文 %d 段（3 号仿宋，固定行距 28.99pt 即每面 22 行，首行缩进 2 字）；"
        "一级标题 %d 处（黑体 3 号）、二级 %d 处（楷体 3 号）、三/四级 %d 处（仿宋 3 号）"
        % (counts["body"], counts["h1"], counts["h2"], counts["h3"] + counts["h4"])
    )
    if counts["attach"]:
        log.append("附件说明 %d 处：左空二字，名称不加标点" % counts["attach"])
    if counts["date"]:
        log.append("成文日期 %d 处：阿拉伯数字，右空四字" % counts["date"])
    if doc.tables:
        log.append("表格 %d 个：宋体小四，表头加粗并跨页重复，行不跨页断开" % len(doc.tables))


# ---------------- 体检 ----------------
def analyze(doc, log):
    sizes = {}
    fonts = {}
    for p in doc.paragraphs:
        for r in p.runs:
            sz = r.font.size.pt if r.font.size else None
            sizes[sz] = sizes.get(sz, 0) + 1
            rPr = r._element.find(qn("w:rPr"))
            ea = None
            if rPr is not None:
                rf = rPr.find(qn("w:rFonts"))
                if rf is not None:
                    ea = rf.get(qn("w:eastAsia"))
            fonts[ea] = fonts.get(ea, 0) + 1
    sec = doc.sections[0]
    log.append(
        "版心：%.0fmm×%.0fmm（标准 156mm×225mm）"
        % ((sec.page_width.mm - sec.left_margin.mm - sec.right_margin.mm),
           (sec.page_height.mm - sec.top_margin.mm - sec.bottom_margin.mm))
    )
    log.append("字号分布：%s" % ", ".join("%spt×%d" % (k, v) for k, v in sorted(sizes.items(), key=lambda x: str(x[0]))))
    log.append("中文字体分布：%s" % ", ".join("%s×%d" % (k, v) for k, v in fonts.items()))
    total_chars = sum(len(p.text) for p in doc.paragraphs)
    log.append("正文段落 %d 段，字符 %d 个" % (len(doc.paragraphs), total_chars))


def count_pages(path):
    try:
        import win32com.client as win32
        app = win32.DispatchEx("Word.Application")
        app.Visible = False
        d = app.Documents.Open(os.path.abspath(path), ReadOnly=True)
        n = int(d.ComputeStatistics(2))
        d.Close(False)
        app.Quit()
        return n, "Word"
    except Exception:
        pass
    try:
        outdir = tempfile.mkdtemp()
        exe = "soffice"
        subprocess.run([exe, "--headless", "--convert-to", "pdf", "--outdir", outdir, path],
                       check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=180)
        pdf = os.path.join(outdir, os.path.splitext(os.path.basename(path))[0] + ".pdf")
        if os.path.exists(pdf):
            with open(pdf, "rb") as f:
                data = f.read()
            return data.count(b"/Type /Page") - data.count(b"/Type /Pages"), "LibreOffice"
    except Exception:
        pass
    return None, None


def ensure_docx(path):
    if path.lower().endswith(".docx"):
        return path
    tmpdir = tempfile.mkdtemp()
    target = os.path.join(tmpdir, os.path.splitext(os.path.basename(path))[0] + ".docx")
    try:
        import win32com.client as win32
        app = win32.DispatchEx("Word.Application")
        app.Visible = False
        d = app.Documents.Open(os.path.abspath(path))
        d.SaveAs2(target, FileFormat=16)
        d.Close(False)
        app.Quit()
        return target
    except Exception:
        pass
    try:
        subprocess.run(["soffice", "--headless", "--convert-to", "docx", "--outdir", tmpdir, path],
                       check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=180)
        if os.path.exists(target):
            return target
    except Exception:
        pass
    sys.exit("无法转换 .doc 文件，请先用 Word 另存为 .docx 后重试")


def main():
    ap = argparse.ArgumentParser(description="按 GB/T 9704-2012 规范化公文格式")
    ap.add_argument("input", help="输入文件（.docx 或 .doc）")
    ap.add_argument("-o", "--output", help="输出文件路径")
    ap.add_argument("--check", action="store_true", help="只体检并输出报告，不写文件")
    ap.add_argument("--no-page-number", action="store_true", help="不生成页码")
    ap.add_argument("--char-spacing", type=float, default=-0.2,
                    help="正文字符间距（磅），默认 -0.2 用于每行 28 字排版")
    args = ap.parse_args()

    src = args.input
    if not os.path.exists(src):
        sys.exit("文件不存在：%s" % src)

    src_docx = ensure_docx(src)
    doc = Document(src_docx)

    if args.check:
        log = []
        analyze(doc, log)
        n, tool = count_pages(src_docx)
        if n:
            log.append("页数：%d 页（由 %s 统计）" % (n, tool))
        print("\n".join(log))
        return

    log = []
    normalize(doc, args, log)

    out = args.output
    if not out:
        base = os.path.splitext(os.path.basename(src))[0]
        out = os.path.join(os.path.dirname(os.path.abspath(src)), base + "_国标版.docx")
    doc.save(out)

    n, tool = count_pages(out)
    log.append("输出：%s" % out)
    if n:
        log.append("页数：%d 页（由 %s 统计）" % (n, tool))
    print("\n".join(log))


if __name__ == "__main__":
    main()
