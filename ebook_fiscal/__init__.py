"""Gerador de e-books de rotinas fiscais por empresa."""

from .document import create_template, generate_ebook
from .model import CompanyFiscalRoutine
from .spreadsheet import read_companies

__all__ = [
    "CompanyFiscalRoutine",
    "create_template",
    "generate_ebook",
    "read_companies",
]
