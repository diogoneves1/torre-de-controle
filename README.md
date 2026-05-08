# 🏭 Torre de Controle de Estoque e Cobertura

Sistema de análise de itens com falta de estoque para a filial **CDSC**.

## Como usar

1. Acesse o app no Streamlit Cloud
2. Faça upload das 4 planilhas:
   - **Análise de Demanda** — Carteira de Pedidos
   - **DW Quarentena WMS**
   - **Planejado x Vendido**
   - **Planilha de Atendimento**
3. Clique em **Gerar Análise**
4. Filtre e explore os resultados
5. Baixe o relatório em Excel

## Status Operacionais

| Status | Significado |
|---|---|
| ✅ COBERTURA VIA QUARENTENA | Quarentena suficiente para cobrir a falta |
| 🟡 COBERTURA PARCIAL | Quarentena existe mas não cobre tudo |
| 🔵 COBERTURA FUTURA | Cobertura prevista sem quarentena atual |
| 🟠 AGUARDAR PRODUÇÃO | Depende de produção prevista |
| ⚫ SEM PRODUÇÃO NO MÊS | Sem produção prevista no período |
| 🔴 CRÍTICO | Sem quarentena e sem previsão |

## Deploy

Hospedado no [Streamlit Cloud](https://streamlit.io/cloud).  
Não requer instalação local.
