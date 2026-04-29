from __future__ import annotations

from collections.abc import Iterable
from importlib import resources
from pathlib import Path
import sys

from docx import Document
from docx.document import Document as DocumentObject
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor
from docx.table import _Cell
from docx.text.paragraph import Paragraph

from .model import CompanyFiscalRoutine, ensure_docx_path

ORANGE = "FF6600"
DARK = RGBColor(40, 40, 40)
WHITE = RGBColor(255, 255, 255)
PLACEHOLDER_EMPTY = "A preencher"


def default_template_path() -> Path | None:
    if getattr(sys, "frozen", False):
        bundled_template = (
            Path(getattr(sys, "_MEIPASS")) / "ebook_fiscal" / "templates" / "Ebook_Fiscal_Template.docx"
        )
        if bundled_template.exists():
            return bundled_template

    template = resources.files("ebook_fiscal").joinpath("templates/Ebook_Fiscal_Template.docx")
    return Path(str(template)) if template.is_file() else None


def create_template(output_path: str | Path) -> Path:
    output_path = ensure_docx_path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    document = Document()
    setup_document(document)
    add_cover(document, "{{razao_social}}")
    add_identification(document, placeholder=True)
    add_contacts(document, placeholder=True)
    add_text_section(document, "OBRIGAÇÕES FISCAIS", "{{obrigacoes_fiscais}}")
    add_text_section(document, "ROTINAS OPERACIONAIS", "{{processos}}")
    add_text_section(document, "PARTICULARIDADES DA EMPRESA", "{{particularidades}}")
    add_text_section(document, "PONTOS DE ATENÇÃO", "{{pontos_atencao}}")
    add_text_section(document, "CALENDÁRIO FISCAL", "{{calendario_fiscal}}")
    add_checklist(document)
    add_footer(document)
    document.save(output_path)
    return output_path


def generate_ebook(
    company: CompanyFiscalRoutine,
    output_path: str | Path,
    template_path: str | Path | None = None,
) -> Path:
    output_path = ensure_docx_path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    template = Path(template_path) if template_path else default_template_path()
    if template and template.exists():
        document = Document(template)
        replace_placeholders(document, template_values(company))
    else:
        document = build_document(company)

    document.save(output_path)
    return output_path


def build_document(company: CompanyFiscalRoutine) -> DocumentObject:
    document = Document()
    setup_document(document)
    add_cover(document, value_or_placeholder(company.razao_social))
    add_identification(document, company=company)
    add_contacts(document, company=company)
    add_items_section(document, "OBRIGAÇÕES FISCAIS", company.obrigacoes_fiscais)
    add_items_section(document, "ROTINAS OPERACIONAIS", company.processos)
    add_items_section(document, "PARTICULARIDADES DA EMPRESA", company.particularidades)
    add_items_section(document, "PONTOS DE ATENÇÃO", company.pontos_atencao)
    add_calendar(document, company.calendario_fiscal, company.responsavel)
    add_checklist(document)
    add_footer(document)
    return document


def setup_document(document: DocumentObject) -> None:
    section = document.sections[0]
    section.top_margin = Inches(0.7)
    section.bottom_margin = Inches(0.7)
    section.left_margin = Inches(0.75)
    section.right_margin = Inches(0.75)

    styles = document.styles
    styles["Normal"].font.name = "Aptos"
    styles["Normal"].font.size = Pt(10.5)
    styles["Normal"].font.color.rgb = DARK


def add_cover(document: DocumentObject, company_name: str) -> None:
    title = document.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.add_run("E-BOOK DE ROTINAS FISCAIS")
    run.bold = True
    run.font.size = Pt(24)
    run.font.color.rgb = RGBColor(255, 102, 0)

    subtitle = document.add_paragraph()
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    subtitle_run = subtitle.add_run("Guia padronizado por empresa")
    subtitle_run.font.size = Pt(12)
    subtitle_run.italic = True

    company = document.add_paragraph()
    company.alignment = WD_ALIGN_PARAGRAPH.CENTER
    company_run = company.add_run(company_name)
    company_run.bold = True
    company_run.font.size = Pt(14)

    document.add_paragraph()


def add_identification(
    document: DocumentObject,
    company: CompanyFiscalRoutine | None = None,
    placeholder: bool = False,
) -> None:
    add_section_header(document, "IDENTIFICAÇÃO DA EMPRESA")
    rows = [
        ("Número da empresa", field_value(company, "numero_empresa", placeholder)),
        ("Razão social", field_value(company, "razao_social", placeholder)),
        ("CNPJ", field_value(company, "cnpj", placeholder)),
        ("UF", field_value(company, "uf", placeholder)),
        ("Regime tributário", field_value(company, "regime_tributario", placeholder)),
        ("Tipo Comercial", field_value(company, "tipo_comercial", placeholder)),
        ("Segmento/atividade", field_value(company, "segmento_atividade", placeholder)),
    ]
    add_key_value_table(document, rows)


def add_contacts(
    document: DocumentObject,
    company: CompanyFiscalRoutine | None = None,
    placeholder: bool = False,
) -> None:
    add_section_header(document, "CONTATOS")
    rows = [
        ("Responsável", field_value(company, "responsavel", placeholder)),
        ("E-mail", field_value(company, "email", placeholder)),
        ("Telefone", field_value(company, "telefone", placeholder)),
        ("Cargo", field_value(company, "cargo", placeholder)),
    ]
    add_key_value_table(document, rows)


def add_section_header(document: DocumentObject, title: str) -> None:
    table = document.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = table.cell(0, 0)
    shade_cell(cell, ORANGE)
    paragraph = cell.paragraphs[0]
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = paragraph.add_run(title)
    run.bold = True
    run.font.size = Pt(14)
    run.font.color.rgb = WHITE
    document.add_paragraph()


def add_key_value_table(document: DocumentObject, rows: list[tuple[str, str]]) -> None:
    table = document.add_table(rows=len(rows), cols=2)
    table.style = "Table Grid"
    table.columns[0].width = Inches(1.9)
    table.columns[1].width = Inches(5.9)

    for index, (label, value) in enumerate(rows):
        label_cell = table.cell(index, 0)
        value_cell = table.cell(index, 1)
        shade_cell(label_cell, "F2F2F2")
        label_run = label_cell.paragraphs[0].add_run(label)
        label_run.bold = True
        value_cell.paragraphs[0].add_run(value_or_placeholder(value))
    document.add_paragraph()


def add_text_section(document: DocumentObject, title: str, text: str) -> None:
    add_section_header(document, title)
    paragraph = document.add_paragraph()
    paragraph.add_run(text)
    document.add_paragraph()


def add_items_section(document: DocumentObject, title: str, raw_items: str) -> None:
    add_section_header(document, title)
    items = split_items(raw_items)
    if not items:
        document.add_paragraph(PLACEHOLDER_EMPTY)
    for item in items:
        document.add_paragraph(item, style="List Bullet")
    document.add_paragraph()


def add_calendar(document: DocumentObject, raw_calendar: str, default_responsible: str) -> None:
    add_section_header(document, "CALENDÁRIO FISCAL")
    rows = parse_calendar(raw_calendar)
    if not rows:
        rows = [("A preencher", "A preencher", default_responsible or "A preencher")]

    table = document.add_table(rows=1, cols=3)
    table.style = "Table Grid"
    headers = ("Obrigação", "Data/prazo", "Responsável")
    for index, header in enumerate(headers):
        cell = table.cell(0, index)
        shade_cell(cell, "F2F2F2")
        run = cell.paragraphs[0].add_run(header)
        run.bold = True

    for obligation, due_date, responsible in rows:
        cells = table.add_row().cells
        cells[0].text = obligation
        cells[1].text = due_date
        cells[2].text = responsible
    document.add_paragraph()


def add_checklist(document: DocumentObject) -> None:
    add_section_header(document, "CHECKLIST DE CONFORMIDADE")
    for item in (
        "Conferência de notas fiscais",
        "Apuração de impostos",
        "Entrega de obrigações",
        "Revisão final",
    ):
        document.add_paragraph(f"☐ {item}")


def add_footer(document: DocumentObject) -> None:
    for section in document.sections:
        footer = section.footer.paragraphs[0]
        footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
        footer.text = "Modelo profissional de padronização fiscal"


def replace_placeholders(document: DocumentObject, values: dict[str, str]) -> None:
    replacements = {f"{{{{{key}}}}}": value_or_placeholder(value) for key, value in values.items()}
    for paragraph in iter_all_paragraphs(document):
        text = paragraph.text
        if not any(token in text for token in replacements):
            continue
        for token, value in replacements.items():
            text = text.replace(token, value)
        set_paragraph_text(paragraph, text)


def template_values(company: CompanyFiscalRoutine) -> dict[str, str]:
    values = company.as_placeholders()
    for field in ("obrigacoes_fiscais", "processos", "particularidades", "pontos_atencao"):
        values[field] = format_bullets(values[field])
    values["calendario_fiscal"] = format_calendar_text(values["calendario_fiscal"], company.responsavel)
    return values


def iter_all_paragraphs(document: DocumentObject) -> Iterable[Paragraph]:
    yield from document.paragraphs
    for table in document.tables:
        yield from iter_table_paragraphs(table)
    for section in document.sections:
        yield from section.header.paragraphs
        yield from section.footer.paragraphs
        for table in section.header.tables:
            yield from iter_table_paragraphs(table)
        for table in section.footer.tables:
            yield from iter_table_paragraphs(table)


def iter_table_paragraphs(table) -> Iterable[Paragraph]:
    for row in table.rows:
        for cell in row.cells:
            yield from cell.paragraphs
            for nested_table in cell.tables:
                yield from iter_table_paragraphs(nested_table)


def set_paragraph_text(paragraph: Paragraph, text: str) -> None:
    if paragraph.runs:
        paragraph.runs[0].text = text
        for run in paragraph.runs[1:]:
            run.text = ""
    else:
        paragraph.add_run(text)


def shade_cell(cell: _Cell, fill: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    shading = OxmlElement("w:shd")
    shading.set(qn("w:fill"), fill)
    tc_pr.append(shading)


def field_value(company: CompanyFiscalRoutine | None, field: str, placeholder: bool) -> str:
    if placeholder:
        return f"{{{{{field}}}}}"
    return getattr(company, field, "") if company else ""


def value_or_placeholder(value: str) -> str:
    return value.strip() if value and value.strip() else PLACEHOLDER_EMPTY


def split_items(raw_text: str) -> list[str]:
    normalized = raw_text.replace("•", "\n").replace(";", "\n")
    return [line.strip(" -\t") for line in normalized.splitlines() if line.strip(" -\t")]


def parse_calendar(raw_calendar: str) -> list[tuple[str, str, str]]:
    rows: list[tuple[str, str, str]] = []
    for line in split_items(raw_calendar):
        parts = [part.strip() for part in line.split("|")]
        if len(parts) == 1:
            rows.append((parts[0], "A preencher", "A preencher"))
        elif len(parts) == 2:
            rows.append((parts[0], parts[1], "A preencher"))
        else:
            rows.append((parts[0], parts[1], parts[2]))
    return rows


def format_bullets(raw_text: str) -> str:
    items = split_items(raw_text)
    return "\n".join(f"• {item}" for item in items) if items else PLACEHOLDER_EMPTY


def format_calendar_text(raw_calendar: str, default_responsible: str) -> str:
    rows = parse_calendar(raw_calendar)
    if not rows:
        rows = [("A preencher", "A preencher", default_responsible or "A preencher")]
    lines = ["Obrigação | Data/prazo | Responsável"]
    lines.extend(f"{obligation} | {due_date} | {responsible}" for obligation, due_date, responsible in rows)
    return "\n".join(lines)
