$ErrorActionPreference = "Stop"

python -m PyInstaller `
  --noconfirm `
  --clean `
  --windowed `
  --name "GeradorEbookFiscal" `
  --add-data "ebook_fiscal\templates\Ebook_Fiscal_Template.docx;ebook_fiscal\templates" `
  --hidden-import "openpyxl.cell._writer" `
  gerador_ebook_fiscal.py

Write-Host ""
Write-Host "Executavel criado em: dist\GeradorEbookFiscal\GeradorEbookFiscal.exe"
