# RotaPro Corretor v1 — projeto independente

Aplicação estática para GitHub Pages. Faça upload de `index.html` e `app.js` na raiz do novo repositório. Em Settings > Pages, escolha Deploy from a branch > main > /(root).

- Importa XLSX no navegador e exporta XLSX.
- Preserva todas as colunas originais e adiciona colunas de revisão e agrupamento.
- Normaliza abreviações comuns, com regra contextual específica para `Alfredo J da Costa`.
- Agrupa por cidade + rua + número, sem complementos.
- Detecta divergências de CEP e coordenadas dentro de um grupo.
- Abre links de conferência no Google Maps e Waze; **não** faz geocodificação automática nem valida o CEP oficialmente.
- Não publica nem envia a planilha ao GitHub. A biblioteca SheetJS é carregada por CDN.

**Limites:** nomes de ruas abreviados podem ser ambíguos; correções de nomes não são prova de endereço oficial. Coordenadas e CEP originais são preservados. O número de parada corrigida representa grupos, não uma otimização de rota.
