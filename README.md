# Gerador de E-book Fiscal

Pacote Python para gerar automaticamente um e-book em Word para cada empresa/linha de uma planilha Excel. O layout segue o modelo profissional anexado, com capa, seções em destaque laranja, identificação, contatos, obrigações fiscais, rotinas, particularidades, pontos de atenção, calendário fiscal e checklist.

## Instalação

```powershell
python -m pip install -r requirements.txt
```

Opcionalmente, instale o pacote em modo editável:

```powershell
python -m pip install -e .
```

## Abrir o aplicativo

```powershell
python gerador_ebook_fiscal.py
```

O aplicativo permite selecionar a planilha, a pasta de saída, uma aba específica e um template Word personalizado.

## Gerar o executável

```powershell
python scripts\build_exe.py
```

O executável será criado em:

```text
dist\GeradorEbookFiscal\GeradorEbookFiscal.exe
```

## Criar o template Word

```powershell
python -m ebook_fiscal.cli criar-template --saida ebook_fiscal/templates/Ebook_Fiscal_Template.docx
```

O template usa placeholders como `{{razao_social}}`, `{{cnpj}}`, `{{obrigacoes_fiscais}}` e pode ser editado no Word sem alterar o script.

## Criar uma planilha de exemplo

```powershell
python -m ebook_fiscal.cli criar-exemplo --saida examples/empresas_fiscais_exemplo.xlsx
```

## Gerar os e-books

```powershell
python -m ebook_fiscal.cli gerar --planilha examples/empresas_fiscais_exemplo.xlsx --saida ebooks_gerados
```

Também é possível informar uma aba e um template personalizado:

```powershell
python -m ebook_fiscal.cli gerar --planilha minha_planilha.xlsx --aba Empresas --template ebook_fiscal/templates/Ebook_Fiscal_Template.docx --saida ebooks_gerados
```

## Colunas esperadas

A planilha deve ter uma linha de cabeçalho. O leitor aceita variações com ou sem acentos e algumas grafias alternativas.

- `Número da empresa`
- `Razão social`
- `CNPJ`
- `UF`
- `Regime tributário`
- `Tipo Comercial`
- `Segmento/atividade` ou `Seguimento/atividade`
- `Responsável`
- `E-mail`
- `Telefone`
- `Cargo`
- `Obrigações fiscais`
- `Processos`
- `Particularidades` ou `Particulariedades`
- `Pontos de atenção`
- `Calendário fiscal` opcional

Para listas, use quebras de linha ou `;`. No calendário fiscal, use o formato `Obrigação|Data/prazo|Responsável`, separando vários itens por `;`.

link: https://drive.google.com/file/d/1M_UyEr7idrJvPm29heHlBRA8aLmuybnR/view?usp=drive_link
