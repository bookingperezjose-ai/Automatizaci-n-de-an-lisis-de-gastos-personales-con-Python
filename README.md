# Automação de Análise de Gastos Pessoais

Pipeline completo em Python que simula, processa e analisa 6 meses de
transações bancárias (320 transações), com classificação automática,
detecção de outliers e dashboard interativo em Excel.

## Stack
Python (Pandas, NumPy, Faker-style synthetic data) · openpyxl · Excel (fórmulas nativas + gráficos)

## Pipeline

```
python generate_expenses_data.py      # 1. Gera transactions.csv (320 transações, 6 meses)
python clean_and_analyze_expenses.py  # 2. Classifica, detecta outliers, calcula KPIs
python build_excel_report.py          # 3. Gera expenses_report.xlsx
```

Abrir `expenses_report.xlsx` → ver dashboard, gráficos e KPIs.

## O que o pipeline faz

1. **Geração de dados** — 320 transações sintéticas mas realistas (comércios
   reais do mercado BR), distribuídas em 5 categorias + renda, com saldo
   bancário coerente transação a transação.
2. **Validação de categorias** contra um conjunto fechado de regras
   (Alimentação, Transporte, Entretenimento, Serviços, Outros). **Importante:**
   neste pipeline a categoria já vem definida na origem dos dados sintéticos —
   o script de limpeza valida consistência, não infere categoria a partir de
   texto livre/sujo. Um extrato bancário real (ex. "PGTO *CARREFOUR 01 CWB")
   exigiria parsing por regex ou matching fuzzy, não implementado aqui.
3. **Detecção de outliers** via IQR (Q3 + 1,5×RIQ) — método robusto para
   distribuições assimétricas (gastos pessoais tendem a cauda longa à
   direita; média + 2 desvios-padrão mascara outliers em séries não-normais).
4. **KPIs calculados**: gasto total, médio mensal, máximo, mínimo, variação
   mês a mês, projeção do próximo mês (média móvel 3 meses).
5. **Dashboard Excel** com 5 abas:
   - `Dashboard`: KPIs destacados + gráfico de linha (evolução mensal) +
     pizza (categorias) + barras (variação % mês a mês)
   - `Top10_Gastos`: as 10 maiores transações + gráfico de barras
   - `Tabla_Dinamica`: pivot categoria × mês com fórmulas nativas do Excel
     (SUM dinâmico, recalculável)
   - `Outliers_Detectados`: transações anômalas isoladas e destacadas
   - `Datos_Completos`: base de dados limpa completa

## Resultados (dados sintéticos desta execução)

| KPI | Valor |
|---|---|
| Gasto total (6 meses) | R$ 75.861,42 |
| Gasto médio mensal | R$ 10.837,35 |
| Maior transação | R$ 1.993,81 |
| Outliers detectados | 8 transações |
| Projeção próximo mês | R$ 12.986,96 |

## Decisões técnicas

- Todas as células calculadas no Excel usam **fórmulas nativas** (SUM,
  variação percentual com proteção contra divisão por zero), não valores
  fixos do Python — a planilha permanece dinâmica se os dados de origem
  forem editados.
- Outliers definidos por **IQR (Q3 + 1,5×RIQ)**, não média + 2σ — método
  robusto ante distribuições assimétricas, padrão em ciência de dados para
  séries financeiras (Bruce & Bruce, *Practical Statistics for Data
  Scientists*, O'Reilly, 2017).
- Dados sintéticos mas com padrões realistas de mercado BR (comércios,
  faixas de valor, sazonalidade de categorias) para demonstrar raciocínio
  de modelagem de dados financeiros.
- **Limitação assumida:** a "classificação" valida categorias pré-definidas
  na origem, não infere a partir de texto sujo. Isso é uma simplificação
  deliberada do escopo, não uma promessa de NLP/regex — declarado aqui para
  honestidade técnica com qualquer avaliador.

## Arquitetura flexível: Excel + Power BI

O pipeline gera `expenses_clean.csv` como saída intermediária limpa, que
alimenta **dois destinos**, cada um resolvendo um problema diferente:

- **`expenses_report.xlsx` (openpyxl):** distribuição imediata, custo zero,
  sem dependência de licença — relevante no mercado de PYMEs do Paraná, onde
  Excel ainda é o padrão operacional (ex.: exportações `.xlsx` em ERPs como
  ContaAzul/Omie).
- **Power BI (via CSV):** importar `expenses_clean.csv` para exploração
  interativa e drill-down — relevante para o mercado corporativo de vagas de
  Analista de Dados/BI, que demanda fluência no ecossistema Python/SQL +
  Power BI.

Não são soluções concorrentes — são saídas para públicos diferentes do
mesmo pipeline de dados.

## Aplicação real

Esta mesma arquitetura (geração → classificação → detecção de anomalias →
dashboard) se aplica diretamente a:
- Extratos bancários reais (Itaú, Nubank, etc.) via export CSV
- Conciliação financeira de pequenas empresas
- Automação de relatórios financeiros recorrentes (mensal/trimestral)
