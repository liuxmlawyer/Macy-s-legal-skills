#!/usr/bin/env python3
"""Convert a compliance assessment Markdown report into a Word document.

设计要点（对应本轮 F02 / F14 / F15 / F16）

F02 实质文本与来源不得静默丢失
  · 代码围栏必须闭合：遇到未闭合围栏时，**保留其内容**并显式告警，
    不再静默丢弃。
  · 正文角标 URL（["https://..."]）不再直接剥除：统一转写为文末
    「尾注」条目，角标位置留占位标记，保证来源在 Word 中可查。
  · 转换后做**内容与来源核对**：正文中每一条实质文本与每一条来源
    都必须能在生成的 docx 中找到，否则以非零退出码失败，不产出文件。

F14 CLI 与降级
  · python-docx 按操作导入；缺依赖时显示受控错误与降级提示。
    `--help` 仍显示帮助，但以 1 退出提示当前无法生成（无裸堆栈）。
  · 明确的退出码：0 成功；1 受控失败；2 参数错误。

F15 源件保护
  · 输出路径与输入为同一文件（含大小写差异、硬链接、符号链接）时拒绝。
  · 输出后缀必须是 .docx。
  · 覆盖既有目标文件需显式 --overwrite。

F16 空报告
  · 空 Markdown 或无实质内容的 Markdown 直接拒绝，不生成空 Word。
"""

from __future__ import annotations

import argparse
import os
import re
import sys
from pathlib import Path


HEADING_RE = re.compile(r"^(#{1,6})\s+(.+?)\s*$")
ORDERED_RE = re.compile(r"^\s*(\d+)[.)]\s+(.+?)\s*$")
UNORDERED_RE = re.compile(r"^\s*[-*+]\s+(.+?)\s*$")
HR_RE = re.compile(r"^\s*([-*_])(?:\s*\1){2,}\s*$")

# 正文里模型常把联网检索来源写成句尾角标，如 ["https://..."] 或
# ["https://a","https://b"]。按模板正文不留裸 URL（尾注仍保留链接），
# 这里统一识别这种角标并转写为文末尾注，而不是直接剥除后丢弃来源。
_QUOTED_URL = r'["“”]https?://[^"“”]*["“”]'
INLINE_URL_CITATION_RE = re.compile(
    r'[\[［]\s*' + _QUOTED_URL + r'(?:\s*[,，]\s*' + _QUOTED_URL + r')*\s*[\]］]'
)

# 尾注条目识别：形如 [1] 官方来源 https://example.org/law
ENDNOTE_ITEM_RE = re.compile(r"^\s*[\[［]\s*\d+\s*[\]］]\s*\S")
BARE_URL_RE = re.compile(r"https?://[^\s，。；、）)】\]］\"'“”]+")

# 尾注段落的识别关键词
ENDNOTE_SECTION_KEYWORDS = ("尾注", "来源", "参考", "依据", "引用")

EXIT_OK = 0
EXIT_FAIL = 1
EXIT_USAGE = 2

FONT_EA = "宋体"


class ConversionError(Exception):
    """受控错误：带用户可读的消息与退出码。"""

    def __init__(self, message: str, exit_code: int = EXIT_FAIL):
        super().__init__(message)
        self.exit_code = exit_code


# --------------------------------------------------------------------------
# 文本清洗
# --------------------------------------------------------------------------
def clean_inline(text: str) -> str:
    """Remove light Markdown marks while preserving readable text."""
    text = re.sub(r"\[([^\]]+)\]\(([^)]+)\)", r"\1（\2）", text)
    text = re.sub(r"`([^`]+)`", r"\1", text)
    text = re.sub(r"\*\*([^*]+)\*\*", r"\1", text)
    text = re.sub(r"\*([^*]+)\*", r"\1", text)
    return text.strip()


def extract_citations(text: str):
    """抽取正文中的来源角标，返回 (清理后文本, [url, ...])。"""
    urls: list[str] = []

    def _collect(match):
        for u in BARE_URL_RE.findall(match.group(0)):
            urls.append(u)
        return " 〔来源见尾注〕"

    cleaned = INLINE_URL_CITATION_RE.sub(_collect, text)
    return cleaned, urls


# --------------------------------------------------------------------------
# Word 构建辅助（python-docx 延迟导入，缺依赖时受控降级）
# --------------------------------------------------------------------------
def _require_docx():
    try:
        import docx  # noqa: F401
    except ImportError:
        raise ConversionError(
            "生成 Word 需要 python-docx 库。\n"
            "  安装：pip install python-docx\n"
            "  降级：改为仅交付 Markdown，并在交付说明中写明未生成 Word 的原因。")
    return docx


def docx_available():
    """探测 python-docx 是否可用（只查规格，不执行模块，不产生导入副作用）。"""
    import importlib.util

    try:
        return importlib.util.find_spec("docx") is not None
    except (ImportError, ValueError):
        return False


def set_east_asia_font(run, font_name: str = FONT_EA) -> None:
    from docx.oxml.ns import qn

    run.font.name = font_name
    run._element.rPr.rFonts.set(qn("w:eastAsia"), font_name)


def set_cell_text(cell, text: str) -> None:
    from docx.shared import Pt, RGBColor

    cell.text = clean_inline(text)
    from docx.shared import Pt
    for paragraph in cell.paragraphs:
        paragraph.paragraph_format.space_before = Pt(0)
        paragraph.paragraph_format.space_after = Pt(0)
        paragraph.paragraph_format.line_spacing = 1.15

    for paragraph in cell.paragraphs:
        for run in paragraph.runs:
            set_east_asia_font(run)
            run.font.size = Pt(10)
            run.font.color.rgb = RGBColor(0x00, 0x00, 0x00)


def set_table_borders(table) -> None:
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    tbl = table._tbl
    tbl_pr = tbl.tblPr
    if tbl_pr is None:
        from docx.oxml import OxmlElement

        tbl_pr = OxmlElement("w:tblPr")
        tbl.insert(0, tbl_pr)
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        border = OxmlElement("w:%s" % edge)
        border.set(qn("w:val"), "single")
        border.set(qn("w:sz"), "4")
        border.set(qn("w:space"), "0")
        border.set(qn("w:color"), "D9D9D9")
        borders.append(border)
    tbl_pr.append(borders)


def is_table_separator(line: str) -> bool:
    cells = [cell.strip() for cell in line.strip().strip("|").split("|")]
    return bool(cells) and all(re.fullmatch(r":?-{3,}:?", cell or "") for cell in cells)


def is_table_start(lines: list[str], index: int) -> bool:
    return (
        index + 1 < len(lines)
        and "|" in lines[index]
        and "|" in lines[index + 1]
        and is_table_separator(lines[index + 1])
    )


def parse_table_row(line: str) -> list[str]:
    return [cell.strip() for cell in line.strip().strip("|").split("|")]


def add_table(document, rows: list[list[str]]) -> None:
    if not rows:
        return
    width = max(len(row) for row in rows)
    table = document.add_table(rows=len(rows), cols=width)
    table.style = "Table Grid"
    set_table_borders(table)
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Inches
    from docx.enum.table import WD_ALIGN_VERTICAL
    section = document.sections[-1]
    usable = (section.page_width - section.left_margin - section.right_margin) / 914400
    count = len(table.columns)
    lengths = [max(len(clean_inline(row[i])) if i < len(row) else 0 for row in rows) for i in range(count)]
    weights = [max(0, min(length, 80) - 12) for length in lengths]
    floor = min(0.7, usable / count)
    if sum(weights):
        widths = [floor + (usable - floor * count) * weight / sum(weights) for weight in weights]
    else:
        widths = [usable / count] * count
    table.autofit = False
    for i, column_inches in enumerate(widths):
        table.columns[i].width = Inches(column_inches)
    for index, row in enumerate(table.rows):
        props = row._tr.get_or_add_trPr()
        props.append(OxmlElement("w:cantSplit"))
        if index == 0:
            props.append(OxmlElement("w:tblHeader"))
        for i, cell in enumerate(row.cells):
            cell.width = Inches(widths[i])
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            if index == 0:
                fill = OxmlElement("w:shd")
                fill.set(qn("w:fill"), "F2F2F2")
                cell._tc.get_or_add_tcPr().append(fill)

    for row_index, row in enumerate(rows):
        for col_index in range(width):
            cell_text = row[col_index] if col_index < len(row) else ""
            set_cell_text(table.cell(row_index, col_index), cell_text)
        if row_index == 0:
            for cell in table.rows[row_index].cells:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        run.bold = True


def add_paragraph(document, text: str, style: str | None = None) -> None:
    from docx.shared import Pt, RGBColor

    paragraph = document.add_paragraph(style=style)
    run = paragraph.add_run(clean_inline(text))
    set_east_asia_font(run)
    run.font.size = Pt(11)
    run.font.color.rgb = RGBColor(0x00, 0x00, 0x00)


def configure_document(document) -> None:
    from docx.oxml.ns import qn
    from docx.shared import Pt, RGBColor

    styles = document.styles
    # 删除模板遗留的段落边框，特别是Title样式的主题色横线。
    for border in list(styles.element.xpath(".//w:pBdr")):
        border.getparent().remove(border)

    styles["Normal"].font.name = FONT_EA
    styles["Normal"]._element.rPr.rFonts.set(qn("w:eastAsia"), FONT_EA)
    styles["Normal"].font.size = Pt(11)
    styles["Normal"].font.color.rgb = RGBColor(0x00, 0x00, 0x00)
    for style_name in ("Title", "Heading 1", "Heading 2", "Heading 3", "Intense Quote"):
        style = styles[style_name]
        style.font.name = FONT_EA
        style._element.rPr.rFonts.set(qn("w:eastAsia"), FONT_EA)
        style.font.color.rgb = RGBColor(0x00, 0x00, 0x00)


# --------------------------------------------------------------------------
# 转换
# --------------------------------------------------------------------------
class ConversionResult:
    def __init__(self):
        self.body_paragraphs = 0
        self.tables = 0
        self.citations = []          # 正文角标转来的来源
        self.endnote_entries = []    # 原文已有的尾注条目
        self.substantive_lines = []  # 用于内容核对的实质文本
        self.warnings = []


def convert_markdown_to_docx(markdown_path: Path, output_path: Path) -> ConversionResult:
    from docx import Document
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.shared import Pt, RGBColor

    raw = markdown_path.read_text(encoding="utf-8")
    lines = raw.splitlines()

    result = ConversionResult()
    source_ids = {}
    existing_ids = set()
    scanning_code = False
    for line in lines:
        if line.strip().startswith("```"):
            scanning_code = not scanning_code
            continue
        if scanning_code:
            continue
        entry = re.match(r"^\s*\[(\d+)\]\s+(.+)$", line)
        if entry:
            number = int(entry.group(1))
            if number in existing_ids:
                raise ConversionError("原文尾注编号重复，请先修正：[ %d ]" % number)
            existing_ids.add(number)
            for url in BARE_URL_RE.findall(entry.group(2)):
                source_ids.setdefault(url, number)
    next_source_id = max(existing_ids, default=0) + 1
    new_sources = []

    def map_citations(text):
        nonlocal next_source_id
        def replace(match):
            nonlocal next_source_id
            ids = []
            for url in BARE_URL_RE.findall(match.group(0)):
                if url not in source_ids:
                    source_ids[url] = next_source_id
                    new_sources.append(url)
                    next_source_id += 1
                ids.append(str(source_ids[url]))
                result.citations.append(url)
            return "〔来源" + ",".join(ids) + "〕"
        return INLINE_URL_CITATION_RE.sub(replace, text)


    # F16：空报告拒绝
    if not raw.strip():
        raise ConversionError(
            "拒绝转换空的 Markdown 报告：源文件无任何内容，不生成空 Word。\n"
            "  请先完成评估并生成报告正文，或改用「仅 Markdown」交付形式。")
    substantive = [ln for ln in lines
                   if ln.strip() and not HR_RE.match(ln)
                   and not ln.strip().startswith("```")]
    if not substantive:
        raise ConversionError(
            "拒绝转换：源文件不含任何实质正文（仅有空行、分隔线或代码围栏标记），"
            "不生成空 Word。\n"
            "  请先完成评估并生成报告正文，或改用「仅 Markdown」交付形式。")

    document = Document()
    configure_document(document)

    in_code_block = False
    code_buffer: list[str] = []
    index = 0
    in_endnote_section = False

    def flush_code_buffer(closed: bool):
        nonlocal code_buffer
        if not code_buffer:
            code_buffer = []
            return
        body = "\n".join(code_buffer)
        add_paragraph(document, body)
        result.substantive_lines.append(body)
        if not closed:
            # F02：未闭合围栏——保留内容并显式告警，不静默丢弃
            result.warnings.append(
                "Markdown 中存在未闭合的代码围栏（``` 缺少收尾）。"
                "已保留其内容并写入 Word，请核对原文是否缺行。")
        code_buffer = []

    while index < len(lines):
        line = lines[index].rstrip()

        if line.strip().startswith("```"):
            if in_code_block:
                flush_code_buffer(closed=True)
                in_code_block = False
            else:
                in_code_block = True
            index += 1
            continue

        if in_code_block:
            code_buffer.append(line)
            index += 1
            continue

        if not line.strip():
            index += 1
            continue

        if HR_RE.match(line):
            index += 1
            continue

        line = map_citations(line)

        if is_table_start(lines, index):
            table_rows = [parse_table_row(lines[index])]
            index += 2
            while index < len(lines) and "|" in lines[index] and lines[index].strip():
                table_rows.append(parse_table_row(map_citations(lines[index])))
                index += 1
            add_table(document, table_rows)
            result.tables += 1
            for row in table_rows:
                for cell in row:
                    if cell.strip():
                        result.substantive_lines.append(cell.strip())
            continue

        heading = HEADING_RE.match(line)
        if heading:
            level = min(len(heading.group(1)), 3)
            title = clean_inline(heading.group(2))
            in_endnote_section = any(k in title for k in ENDNOTE_SECTION_KEYWORDS)
            if level == 1:
                paragraph = document.add_paragraph(style="Title")
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run = paragraph.add_run(title)
                set_east_asia_font(run)
                run.bold = True
                run.font.size = Pt(16)
                run.font.color.rgb = RGBColor(0x00, 0x00, 0x00)
            else:
                heading_paragraph = document.add_heading(title, level=level - 1)
                for run in heading_paragraph.runs:
                    set_east_asia_font(run)
                    run.font.color.rgb = RGBColor(0x00, 0x00, 0x00)
            result.body_paragraphs += 1
            result.substantive_lines.append(title)
            index += 1
            continue

        # 原文尾注条目：原样保留（URL 不剥除）
        if in_endnote_section and ENDNOTE_ITEM_RE.match(line):
            add_paragraph(document, line)
            result.body_paragraphs += 1
            result.endnote_entries.append(line)
            result.citations.extend(BARE_URL_RE.findall(line))
            index += 1
            continue

        unordered = UNORDERED_RE.match(line)
        if unordered:
            add_paragraph(document, unordered.group(1), style="List Bullet")
            result.body_paragraphs += 1
            result.substantive_lines.append(unordered.group(1))
            index += 1
            continue

        ordered = ORDERED_RE.match(line)
        if ordered:
            # 序号直接写进段落文本，用普通段落输出，
            # 不用 Word 的 List Number 自动编号，避免跨段/跨章节连续累加。
            add_paragraph(document, "%s. %s" % (ordered.group(1), ordered.group(2)))
            result.body_paragraphs += 1
            result.substantive_lines.append(
                "%s. %s" % (ordered.group(1), ordered.group(2)))
            index += 1
            continue

        if line.startswith(">"):
            add_paragraph(document, line.lstrip("> "), style="Intense Quote")
            result.body_paragraphs += 1
            result.substantive_lines.append(line.lstrip("> "))
            index += 1
            continue

        # 正文：角标 URL 转尾注占位
        cleaned, urls = line, []
        add_paragraph(document, cleaned)
        result.body_paragraphs += 1
        result.citations.extend(urls)
        result.substantive_lines.append(cleaned)
        index += 1

    # 文件结束仍未闭合围栏：保留内容 + 告警
    if in_code_block:
        flush_code_buffer(closed=False)

    # F02：正文角标来源统一追加为尾注（原文已有尾注条目则不重复追加同源）
    if new_sources:
        document.add_heading("尾注（正文来源）", level=1)
        for url in new_sources:
            add_paragraph(document, "[%d] %s" % (source_ids[url], url))

    # 内容与来源核对：任何实质文本或来源在 docx 中缺失，即失败
    full_text = "\n".join(p.text for p in document.paragraphs)
    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                full_text += "\n" + cell.text

    missing_text = []
    for piece in result.substantive_lines:
        probe = clean_inline(piece)
        if probe and probe not in full_text:
            missing_text.append(piece)

    missing_sources = [u for u in result.citations if u not in full_text]

    if missing_text or missing_sources:
        details = []
        if missing_text:
            details.append("以下实质文本未出现在生成的 Word 中：\n  - "
                           + "\n  - ".join(missing_text[:10]))
        if missing_sources:
            details.append("以下来源未出现在生成的 Word 中：\n  - "
                           + "\n  - ".join(missing_sources[:10]))
        raise ConversionError(
            "内容/来源核对未通过，已中止生成（未写出任何文件）：\n"
            + "\n".join(details))

    output_path.parent.mkdir(parents=True, exist_ok=True)
    document.save(output_path)
    return result


# --------------------------------------------------------------------------
# 源件保护
# --------------------------------------------------------------------------
def same_file(a: Path, b: Path) -> bool:
    """判断两个路径是否指向同一文件（含大小写差异、硬链接、符号链接）。"""
    a_norm = os.path.normcase(os.path.abspath(str(a)))
    b_norm = os.path.normcase(os.path.abspath(str(b)))
    if a_norm == b_norm:
        return True
    try:
        if a.exists() and b.exists():
            return os.path.samefile(str(a), str(b))
    except OSError:
        pass
    try:
        if a.exists() and b.exists():
            sa = os.stat(str(a))
            sb = os.stat(str(b))
            return (sa.st_dev, sa.st_ino) == (sb.st_dev, sb.st_ino)
    except OSError:
        pass
    return False


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Markdown 用工合规评估报告 → Word 文档")
    parser.add_argument("markdown", type=Path, help="Markdown 报告路径")
    parser.add_argument(
        "-o", "--output", type=Path,
        help="生成的 .docx 路径。默认：与 Markdown 同目录同名（.docx）")
    parser.add_argument(
        "--overwrite", action="store_true",
        help="允许覆盖已存在的目标文件（默认拒绝，避免误覆盖既有交付物）")
    return parser


def main(argv=None) -> int:
    argv = list(sys.argv[1:] if argv is None else argv)

    # F14：缺依赖时**明确声明**而不是静默失败或裸堆栈。
    # --help 仍可正常显示用法，但退出码为 1，让调用方（宿主/定时任务）
    # 能感知到"当前环境缺少可选依赖"。
    if not docx_available():
        print("错误：本机缺少可选依赖 python-docx，无法生成 Word 文档。\n"
              "  安装：pip install python-docx\n"
              "  降级：改为仅交付 Markdown，并在交付说明中写明未生成 Word 的原因。",
              file=sys.stderr)
        if any(a in ("-h", "--help") for a in argv):
            build_parser().print_help()
        return EXIT_FAIL

    parser = build_parser()
    args = parser.parse_args(argv)

    markdown_path = args.markdown.expanduser().resolve()
    if not markdown_path.exists():
        print("错误：Markdown 文件不存在：%s" % markdown_path, file=sys.stderr)
        return EXIT_FAIL
    if markdown_path.suffix.lower() not in (".md", ".markdown", ".txt"):
        print("错误：输入必须是 .md / .markdown / .txt 报告文件，实际为：%s"
              % markdown_path.suffix, file=sys.stderr)
        return EXIT_USAGE

    output_path = (args.output.expanduser().resolve() if args.output
                   else markdown_path.with_suffix(".docx"))

    if output_path.suffix.lower() != ".docx":
        print("错误：输出文件必须为 .docx 后缀，实际为：%s" % output_path.suffix,
              file=sys.stderr)
        return EXIT_USAGE

    # F15：拒绝源目标同一文件
    if same_file(markdown_path, output_path):
        print("错误：输出路径与输入为同一文件，已拒绝以保护原始 Markdown。\n"
              "  输入：%s\n  输出：%s\n"
              "  请指定不同的输出路径（同名不同后缀即可）。"
              % (markdown_path, output_path), file=sys.stderr)
        return EXIT_FAIL

    if output_path.exists() and not args.overwrite:
        print("错误：输出文件已存在：%s\n"
              "  如确需覆盖，请显式加 --overwrite（覆盖前请确认不是既有交付物）。"
              % output_path, file=sys.stderr)
        return EXIT_FAIL

    try:
        _require_docx()
        result = convert_markdown_to_docx(markdown_path, output_path)
    except ConversionError as exc:
        print("错误：%s" % exc, file=sys.stderr)
        return exc.exit_code
    except Exception as exc:  # 受控兜底：不输出裸堆栈
        print("错误：转换失败：%s\n"
              "  降级：改为仅交付 Markdown，并在交付说明中写明未生成 Word 的原因。"
              % exc, file=sys.stderr)
        return EXIT_FAIL

    for w in result.warnings:
        print("⚠️ 警告：%s" % w, file=sys.stderr)
    print(output_path)
    return EXIT_OK


if __name__ == "__main__":
    sys.exit(main())
