# RotaPro Corretor v1.4

Projeto independente do RotaPro principal.

## Novidades
- Analisa toda a planilha, não apenas um endereço de demonstração.
- Separa os números de apartamento, bloco e unidade antes de identificar o número do imóvel, inclusive quando o complemento aparece antes do número predial.
- Padroniza abreviações e aproveita nomes completos existentes na própria planilha.
- Preserva todas as linhas, colunas e abas ao exportar XLSX.
- Sinaliza endereços com mais de um número predial possível para revisão; não adivinha números.
- Não realiza consultas automáticas ao Google Maps/Waze nem usa IA remota.

## Instalação
Envie index.html, app.js e README.md para a raiz do repositório GitHub Pages do corretor.

## Limites
Endereços sem número explícito ou com múltiplos números não identificáveis exigem conferência. O aplicativo processa os arquivos localmente no navegador. A biblioteca SheetJS é carregada por CDN.
