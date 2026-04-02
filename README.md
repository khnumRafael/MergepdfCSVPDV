# MergepdfCSVPDV

Projeto para extrair dados do PDF de pedidos por vendedor, gerar CSV consolidado e gerar PDF final com totalizações.

## Estrutura principal

- `relatorio_unico.py`: script único (extração + geração do PDF).
- `extrair_relatorio.py`: extração e consolidação em CSV.
- `gerar_relatorio_pdf.py`: geração do relatório em PDF a partir do CSV.
- `dist/relatorio_unico.exe`: executável único para Windows.

## Pré-requisito de entrada

Coloque o arquivo PDF de origem em:

- `c:\pdf\porVendedor.pdf`

## Como executar (Windows)

No PowerShell:

```powershell
& "c:\tmp\pedidos pr\dist\relatorio_unico.exe"
```

## Saídas geradas

Após executar, os arquivos são gerados em:

- `c:\pdf\relatorio_agregado.csv`
- `c:\pdf\relatorio_agregado.pdf`

## O que o relatório entrega

- Consolidação por vendedor e produto.
- Colunas: código, barras, descrição, peso, unidade, quantidade e total.
- Subtotal por vendedor.
- Totalizador de peso por vendedor.
- Peso total geral estimado.
- Total financeiro geral.

## Publicação no Git

Comandos básicos:

```powershell
git add README.md
git commit -m "Adicionar README com instruções de uso"
git push
```
