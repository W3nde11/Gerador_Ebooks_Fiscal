from __future__ import annotations

import argparse
from pathlib import Path

from openpyxl import Workbook

from .document import create_template, generate_ebook
from .spreadsheet import read_companies


def main() -> None:
    parser = argparse.ArgumentParser(
        prog="gerar-ebooks-fiscais",
        description="Gera e-books de rotinas fiscais em Word a partir de uma planilha Excel.",
    )
    subparsers = parser.add_subparsers(dest="command", required=True)

    generate_parser = subparsers.add_parser("gerar", help="Gerar um Word por linha da planilha")
    generate_parser.add_argument("--planilha", required=True, help="Caminho do arquivo .xlsx")
    generate_parser.add_argument("--saida", default="ebooks_gerados", help="Pasta de saída dos .docx")
    generate_parser.add_argument("--aba", default=None, help="Nome da aba. Se omitido, usa a primeira.")
    generate_parser.add_argument("--template", default=None, help="Template .docx com placeholders")
    generate_parser.set_defaults(func=generate_command)

    template_parser = subparsers.add_parser("criar-template", help="Criar template Word editável")
    template_parser.add_argument(
        "--saida",
        default="ebook_fiscal/templates/Ebook_Fiscal_Template.docx",
        help="Caminho do template .docx",
    )
    template_parser.set_defaults(func=template_command)

    example_parser = subparsers.add_parser("criar-exemplo", help="Criar planilha de exemplo")
    example_parser.add_argument(
        "--saida",
        default="examples/empresas_fiscais_exemplo.xlsx",
        help="Caminho da planilha .xlsx",
    )
    example_parser.set_defaults(func=example_command)

    args = parser.parse_args()
    args.func(args)


def generate_command(args: argparse.Namespace) -> None:
    companies = read_companies(args.planilha, sheet_name=args.aba)
    output_dir = Path(args.saida)
    output_dir.mkdir(parents=True, exist_ok=True)

    for company in companies:
        output_path = output_dir / company.output_filename()
        generate_ebook(company, output_path, template_path=args.template)
        print(f"Gerado: {output_path}")

    print(f"Total de e-books gerados: {len(companies)}")


def template_command(args: argparse.Namespace) -> None:
    path = create_template(args.saida)
    print(f"Template criado: {path}")


def example_command(args: argparse.Namespace) -> None:
    path = Path(args.saida)
    path.parent.mkdir(parents=True, exist_ok=True)

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Empresas"
    sheet.append(
        [
            "Número da empresa",
            "Razão social",
            "CNPJ",
            "UF",
            "Regime tributário",
            "Tipo Comercial",
            "Seguimento/atividade",
            "Responsável",
            "E-mail",
            "Telefone",
            "Cargo",
            "Obrigações fiscais",
            "Processos",
            "Particulariedades",
            "Pontos de atenção",
            "Calendário fiscal",
        ]
    )
    sheet.append(
        [
            "001",
            "Empresa Exemplo LTDA",
            "12.345.678/0001-90",
            "SP",
            "Lucro Presumido",
            "Cliente recorrente",
            "Comércio varejista",
            "Maria Silva",
            "maria@empresaexemplo.com.br",
            "(11) 99999-9999",
            "Gerente fiscal",
            "SPED Fiscal; EFD Contribuições; DCTFWeb; GIA",
            "Receber XMLs até o dia 3; Conferir CFOP/CST; Apurar impostos até o dia 10",
            "Possui benefícios fiscais estaduais; Operações interestaduais recorrentes",
            "Validar crédito de ICMS; Acompanhar vencimentos municipais",
            "SPED Fiscal|20º dia util|Maria Silva; DCTFWeb|Dia 15|Maria Silva",
        ]
    )

    workbook.save(path)
    print(f"Planilha de exemplo criada: {path}")


if __name__ == "__main__":
    main()
