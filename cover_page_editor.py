from __future__ import annotations

from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


def set_paragraph_rtl(paragraph) -> None:
    """
    Makes paragraph RTL for Arabic text.
    """
    p_pr = paragraph._p.get_or_add_pPr()

    bidi = p_pr.find(qn("w:bidi"))
    if bidi is None:
        bidi = OxmlElement("w:bidi")
        p_pr.append(bidi)

    bidi.set(qn("w:val"), "1")


def replace_paragraph_text_keep_style(paragraph, new_text: str) -> None:
    """
    Replaces paragraph text but keeps the first run style.
    """
    new_text = str(new_text or "").strip()

    if not paragraph.runs:
        paragraph.add_run(new_text)
        return

    paragraph.runs[0].text = new_text

    for run in paragraph.runs[1:]:
        run.text = ""


def update_cover_page(doc, form: dict) -> None:
    """
    Updates the first page / cover page only.

    It replaces the first 4 non-empty paragraphs:
    1. Company name
    2. Legal form
    3. Financial statements title
    4. Cover date
    """

    company_name = str(form.get("company_name", "") or "").strip()
    legal_form = str(form.get("legal_form", "") or "").strip()
    statements_title = str(form.get("statements_title", "") or "").strip()
    cover_date = str(form.get("cover_date", "") or "").strip()

    if not legal_form:
        legal_form = "(ذات مسؤولية محدودة)"

    if not statements_title:
        statements_title = "القوائم الماليــة المنفصلة للسنة المالية المنتهية في"

    if not cover_date:
        cover_date = str(form.get("financial_year", "") or "").strip()

    cover_values = [
        company_name,
        legal_form,
        statements_title,
        cover_date,
    ]

    first_non_empty_paragraphs = []

    for paragraph in doc.paragraphs:
        if paragraph.text.strip():
            first_non_empty_paragraphs.append(paragraph)

        if len(first_non_empty_paragraphs) == 4:
            break

    for paragraph, value in zip(first_non_empty_paragraphs, cover_values):
        if value:
            replace_paragraph_text_keep_style(paragraph, value)
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            set_paragraph_rtl(paragraph)
