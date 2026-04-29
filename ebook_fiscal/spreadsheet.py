from __future__ import annotations

from collections.abc import Iterable
from datetime import date, datetime
from pathlib import Path
import unicodedata

from openpyxl import load_workbook

from .model import CompanyFiscalRoutine


HEADER_ALIASES = {
    "numero_empresa": {
        "numero empresa",
        "numero da empresa",
        "n empresa",
        "num empresa",
        "codigo empresa",
        "cod empresa",
        "empresa",
    },
    "razao_social": {"razao social", "nome", "nome empresa", "empresa razao social"},
    "cnpj": {"cnpj", "cnpj empresa"},
    "uf": {"uf", "estado", "unidade federativa"},
    "regime_tributario": {"regime tributario", "tributacao", "regime"},
    "tipo_comercial": {
        "tipo comercial",
        "tipo de comercial",
        "comercial",
        "classificacao comercial",
    },
    "segmento_atividade": {
        "segmento atividade",
        "seguimento atividade",
        "segmento",
        "seguimento",
        "atividade",
        "ramo atividade",
    },
    "responsavel": {"responsavel", "contato", "nome responsavel"},
    "email": {"email", "e mail", "e-mail", "mail"},
    "telefone": {"telefone", "fone", "celular", "whatsapp"},
    "cargo": {"cargo", "funcao"},
    "obrigacoes_fiscais": {
        "obrigacoes fiscais",
        "obrigacao fiscal",
        "obrigacoes",
        "rotinas fiscais",
    },
    "processos": {"processos", "processo", "rotinas operacionais", "procedimentos"},
    "particularidades": {
        "particularidades",
        "particulariedades",
        "regras especificas",
        "observacoes fiscais",
    },
    "pontos_atencao": {
        "pontos de atencao",
        "pontos atencao",
        "atencao",
        "alertas",
        "observacoes criticas",
    },
    "calendario_fiscal": {"calendario fiscal", "calendario", "prazos", "vencimentos"},
}

REQUIRED_FIELDS = {
    "numero_empresa",
    "razao_social",
    "cnpj",
    "uf",
    "regime_tributario",
    "tipo_comercial",
    "segmento_atividade",
    "responsavel",
    "email",
    "telefone",
    "cargo",
    "obrigacoes_fiscais",
    "processos",
    "particularidades",
    "pontos_atencao",
}


def read_companies(path: str | Path, sheet_name: str | None = None) -> list[CompanyFiscalRoutine]:
    workbook = load_workbook(path, data_only=True)
    worksheet = workbook[sheet_name] if sheet_name else workbook.active

    rows = worksheet.iter_rows(values_only=True)
    try:
        raw_headers = next(rows)
    except StopIteration:
        return []

    header_map = build_header_map(raw_headers)
    missing = sorted(REQUIRED_FIELDS - set(header_map.values()))
    if missing:
        missing_names = ", ".join(missing)
        raise ValueError(f"Planilha sem colunas obrigatorias: {missing_names}")

    companies: list[CompanyFiscalRoutine] = []
    for row_number, row in enumerate(rows, start=2):
        values = row_to_dict(row, header_map)
        if not any(values.values()):
            continue
        try:
            companies.append(CompanyFiscalRoutine(**values))
        except TypeError as exc:
            raise ValueError(f"Linha {row_number} possui dados invalidos: {exc}") from exc

    return companies


def build_header_map(headers: Iterable[object]) -> dict[int, str]:
    aliases = {
        normalize_header(alias): canonical
        for canonical, alias_group in HEADER_ALIASES.items()
        for alias in alias_group | {canonical}
    }

    header_map: dict[int, str] = {}
    for index, header in enumerate(headers):
        normalized = normalize_header(header)
        if not normalized:
            continue
        canonical = aliases.get(normalized)
        if canonical:
            header_map[index] = canonical

    return header_map


def row_to_dict(row: tuple[object, ...], header_map: dict[int, str]) -> dict[str, str]:
    values = {field: "" for field in HEADER_ALIASES}
    for index, field in header_map.items():
        value = row[index] if index < len(row) else ""
        values[field] = stringify_cell(value)
    return values


def normalize_header(value: object) -> str:
    text = stringify_cell(value).lower()
    for char in ("_", "-", "/", "\\"):
        text = text.replace(char, " ")
    text = "".join(
        char for char in unicodedata.normalize("NFKD", text) if not unicodedata.combining(char)
    )
    return " ".join(text.split())


def stringify_cell(value: object) -> str:
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.strftime("%d/%m/%Y")
    if isinstance(value, date):
        return value.strftime("%d/%m/%Y")
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()
