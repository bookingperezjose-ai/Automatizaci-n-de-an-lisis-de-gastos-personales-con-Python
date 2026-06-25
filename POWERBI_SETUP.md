# Conexão Power BI — expenses_clean.csv

`build_excel_report.py` gera o Artefato A (distribuição imediata, sem
licença). Este arquivo documenta o Artefato B: ingestão do mesmo
`expenses_clean.csv` em Power BI para exploração interativa.

## Passo a passo (Power BI Desktop)

1. **Obter Dados** → **Texto/CSV** → selecionar `expenses_clean.csv`
2. Delimitador: vírgula | Codificação: UTF-8
3. Em **Transformar Dados** (Power Query), aplicar:
   - Coluna `Fecha` → tipo Data
   - Coluna `Monto` → tipo Decimal Fixo
   - Coluna `Mes` já vem pronta como texto (formato AAAA-MM) — usar como
     eixo temporal nas visualizações
4. **Fechar e Aplicar**

## Query M equivalente (copiar direto no Editor Avançado)

```m
let
    Origem = Csv.Document(File.Contents("CAMINHO\expenses_clean.csv"),
        [Delimiter=",", Columns=9, Encoding=65001, QuoteStyle=QuoteStyle.None]),
    CabecalhoPromovido = Table.PromoteHeaders(Origem, [PromoteAllScalars=true]),
    TiposAlterados = Table.TransformColumnTypes(CabecalhoPromovido,
        {{"Fecha", type date}, {"Monto", type number}, {"Categoría", type text},
         {"Mes", type text}, {"Tipo", type text}})
in
    TiposAlterados
```

Substituir `CAMINHO\` pelo diretório real do arquivo.

## Visualizações sugeridas (réplica do Dashboard Excel)

| Visual Power BI | Campos |
|---|---|
| Gráfico de linhas | Eixo: `Mes` · Valor: `Sum(Monto)` |
| Gráfico de pizza | Legenda: `Categoría` · Valor: `Sum(Monto)` |
| Gráfico de barras | Eixo: `Descripción` (Top N por `Monto`) |
| Cartões KPI | `Sum(Monto)`, `Average(Monto)`, `Max(Monto)`, `Min(Monto)` |
| Tabela | Matriz `Categoría` × `Mes`, valores `Sum(Monto)` (equivalente à Tabla_Dinamica) |

## Medida DAX para variação % mês a mês

```dax
Variação % =
VAR MesAnterior =
    CALCULATE(
        SUM(expenses_clean[Monto]),
        DATEADD(expenses_clean[Fecha], -1, MONTH)
    )
VAR MesAtual = SUM(expenses_clean[Monto])
RETURN
    IF(MesAnterior = 0, BLANK(), DIVIDE(MesAtual - MesAnterior, MesAnterior))
```

`DIVIDE()` do DAX já protege contra divisão por zero nativamente — equivalente
à correção `IF(B=0,"",...)` aplicada na versão Excel.

## Por que isso não substitui o Artefato A

Power BI exige licença (Pro) para compartilhamento colaborativo além do
Desktop individual, e o arquivo `.pbix` não é portátil sem o software
instalado. Para envio direto a um cliente/empresa do interior do Paraná sem
infraestrutura de BI, o `.xlsx` standalone resolve um problema real que o
`.pbix` não resolve. Os dois artefatos atacam objetivos diferentes — não é
duplicação de esforço.
