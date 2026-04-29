from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import re


@dataclass(slots=True)
class CompanyFiscalRoutine:
    numero_empresa: str
    razao_social: str
    cnpj: str
    uf: str
    regime_tributario: str
    tipo_comercial: str
    segmento_atividade: str
    responsavel: str
    email: str
    telefone: str
    cargo: str
    obrigacoes_fiscais: str
    processos: str
    particularidades: str
    pontos_atencao: str
    calendario_fiscal: str = ""

    def as_placeholders(self) -> dict[str, str]:
        return {
            "numero_empresa": self.numero_empresa,
            "razao_social": self.razao_social,
            "cnpj": self.cnpj,
            "uf": self.uf,
            "regime_tributario": self.regime_tributario,
            "tipo_comercial": self.tipo_comercial,
            "segmento_atividade": self.segmento_atividade,
            "responsavel": self.responsavel,
            "email": self.email,
            "telefone": self.telefone,
            "cargo": self.cargo,
            "obrigacoes_fiscais": self.obrigacoes_fiscais,
            "processos": self.processos,
            "particularidades": self.particularidades,
            "pontos_atencao": self.pontos_atencao,
            "calendario_fiscal": self.calendario_fiscal,
        }

    def output_filename(self) -> str:
        base = f"{self.numero_empresa or 'empresa'} - {self.razao_social or self.cnpj or 'ebook fiscal'}"
        return f"{slugify_filename(base)}.docx"


def slugify_filename(value: str) -> str:
    sanitized = re.sub(r'[<>:"/\\|?*\n\r\t]+', " ", value).strip()
    sanitized = re.sub(r"\s+", " ", sanitized)
    return sanitized[:140] or "ebook-fiscal"


def ensure_docx_path(path: str | Path) -> Path:
    path = Path(path)
    return path if path.suffix.lower() == ".docx" else path.with_suffix(".docx")
