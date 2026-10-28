---
tipo: projeto
criado: 2026-10-28
status: ativo
prazo: 
tags: [projeto, logistica, fretes, reversa, automacao]
---

# Automação Reversa

## Objetivo
Gerar `Orc_vs_Real_Reversa_python.xlsx` ligando cada lançamento do realizado de reversa aos itens transportados (pallet, produto etc.).

Fluxo: **Realizado Reversa** (documento + série + filial) → **NEOGRID_CTE** (XML do CT-e → chave da NF-e de origem) → **RE0501** (.txt com itens e quantidades da NF-e).

Script: `Automacao Reversa/reversa_automacao.py` (rodar com `rodar_reversa.bat`). Pasta raiz: `K:\Logistica\CENTRAL DE FRETES\CUSTOS\Reversa`.

## Abas geradas
- **Realizado_Itens** — cada linha do realizado aberta por item, com chave CT-e, chave NF, quantidade, categoria e valor rateado por quantidade.
- **Resumo_Categoria / Resumo_Item** — totais por categoria e por item.
- **Status_Vinculo / Pendencias** — CT-e não encontrado, NF fora do RE0501 etc.
- **Base_CTe / Base_RE0501 / Parametros** — bases para conferência.

## Próximas ações
- [ ] Instalar dependências: `pip install pandas openpyxl xlrd`
- [ ] Rodar com dados reais e validar os nomes de colunas (`CONFIG["COLS_REALIZADO"]`, `CONFIG["COLS_RE0501"]`)
- [ ] Ajustar regras de categoria (`CONFIG["REGRAS_CATEGORIA"]`)
- [ ] Se o RE0501 for de largura fixa (sem separador), enviar amostra para adaptar o leitor
- [ ] Ligar com a planilha de orçamento (orçado x realizado por categoria)

## Notas e referências
- No XML do CT-e a NF de origem fica em `infCTeNorm/infDoc/infNFe/chave`.
- Leitura dos XMLs é cacheada em `_cache_cte.pkl` (só lê arquivos novos/alterados).

## Log
- 2026-10-28 — projeto criado; primeira versão do script testada com dados fictícios
