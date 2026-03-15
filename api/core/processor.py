import re
from typing import Dict, Any
from io import BytesIO
from html import escape

from docx import Document

def convert_docx_to_html(file_bytes: bytes, log_colors: bool = False, log_fonts: bool = False) -> str:
    """
    將 .docx 轉為適合 Email 的 HTML 格式，保留常見的行內樣式（粗體、斜體、底線、文字顏色）
    以及無序/有序清單。
    log_colors / log_fonts 用來決定是否在 stdout 列出偵測到的色彩與字型，預設不列印。
    """
    doc = Document(BytesIO(file_bytes))
    numbering_root = doc.part.numbering_part.element if doc.part.numbering_part else None

    def log_detected_colors():
        colors = []
        for p_idx, paragraph in enumerate(doc.paragraphs):
            for r_idx, run in enumerate(paragraph.runs):
                col = run.font.color.rgb
                if col:
                    snippet = run.text.replace("\n", " ")[:30]
                    colors.append((p_idx, r_idx, str(col), snippet))
        if colors:
            print("[convert_docx_to_html] detected run colors:")
            for p_idx, r_idx, col, snippet in colors:
                print(f"  p{p_idx} r{r_idx}: #{col} text='{snippet}'")
        else:
            print("[convert_docx_to_html] no run-level colors found")

    if log_colors:
        log_detected_colors()

    def get_font_name(run):
        # 優先讀取 rFonts（常見於 Word 預設「新細明體」等）
        rpr = run._r.rPr
        if rpr is not None and getattr(rpr, "rFonts", None) is not None:
            rfonts = rpr.rFonts
            for attr in ("eastAsia", "ascii", "hAnsi", "cs"):
                val = getattr(rfonts, attr, None)
                if val:
                    return val
        # 其次讀取直接套用的 run.font.name
        return run.font.name

    def log_detected_fonts():
        fonts = []
        for p_idx, paragraph in enumerate(doc.paragraphs):
            for r_idx, run in enumerate(paragraph.runs):
                fname = get_font_name(run)
                if fname:
                    snippet = run.text.replace("\n", " ")[:30]
                    fonts.append((p_idx, r_idx, fname, snippet))
        if fonts:
            print("[convert_docx_to_html] detected run fonts:")
            for p_idx, r_idx, fname, snippet in fonts:
                print(f"  p{p_idx} r{r_idx}: font='{fname}' text='{snippet}'")
        else:
            print("[convert_docx_to_html] no run-level font names found")

    if log_fonts:
        log_detected_fonts()

    def get_list_tag(paragraph) -> str:
        """
        根據 numFmt 決定使用 <ul> 或 <ol>。若判斷不到，預設用 <ul>。
        """
        if numbering_root is None:
            return "ul"
        num_pr = paragraph._p.pPr.numPr
        num_id = num_pr.numId.val
        ilvl = num_pr.ilvl.val if num_pr.ilvl is not None else 0
        num_el = numbering_root.xpath(f".//w:num[@w:numId='{num_id}']")
        if not num_el:
            return "ul"
        abstract_id = num_el[0].xpath("w:abstractNumId/@w:val")[0]
        fmt = numbering_root.xpath(
            f".//w:abstractNum[@w:abstractNumId='{abstract_id}']/w:lvl[@w:ilvl='{ilvl}']/w:numFmt/@w:val"
        )
        if fmt and fmt[0] != "bullet":
            return "ol"
        return "ul"

    def run_to_html(run) -> str:
        text = escape(run.text)
        if not text:
            return ""
        # 保留連續空白
        text = text.replace("  ", "&nbsp;&nbsp;")

        if run.bold:
            text = f"<strong>{text}</strong>"
        if run.italic:
            text = f"<em>{text}</em>"
        if run.underline:
            text = f"<u>{text}</u>"

        color = run.font.color.rgb
        if color:
            text = f'<span style="color: #{color}">{text}</span>'
        return text

    html_parts = []
    in_list = False
    current_list_tag = None

    def paragraph_is_list(p) -> bool:
        return p._p.pPr is not None and p._p.pPr.numPr is not None

    for paragraph in doc.paragraphs:
        is_list = paragraph_is_list(paragraph)

        if is_list:
            list_tag = get_list_tag(paragraph)
            if not in_list or list_tag != current_list_tag:
                if in_list:
                    html_parts.append(f"</{current_list_tag}>")
                html_parts.append(f"<{list_tag}>")
                in_list = True
                current_list_tag = list_tag
        else:
            if in_list:
                html_parts.append(f"</{current_list_tag}>")
                in_list = False
                current_list_tag = None

        content = "".join(run_to_html(run) for run in paragraph.runs)
        if not content:
            continue

        if is_list:
            html_parts.append(f"<li>{content}</li>")
        else:
            html_parts.append(f"<p>{content}</p>")

    if in_list and current_list_tag:
        html_parts.append(f"</{current_list_tag}>")

    wrapped_html = (
        "<div style=\"font-family: 'Microsoft JhengHei', sans-serif; "
        "line-height: 1.6; color: #333;\">"
        + "".join(html_parts)
        + "</div>"
    )
    return wrapped_html

def inject_variables(html_template: str, row_data: Dict[str, Any]) -> str:
    """
    將 HTML 模板中的 {{變數}} 替換為 Excel 的資料內容
    支援自定義欄位，只要 Excel 表頭名稱與 {{}} 內一致即可。
    """
    # 使用正則表達式尋找 {{Key}} 並從 row_data 抓取對應的 Value
    def replace_match(match):
        key = match.group(1).strip()
        # 若 Excel 沒這欄位，則保持原樣或回傳空字串
        return str(row_data.get(key, match.group(0)))

    # 匹配 {{ variable_name }} 格式
    pattern = r"\{\{\s*(.*?)\s*\}\}"
    return re.sub(pattern, replace_match, html_template)
