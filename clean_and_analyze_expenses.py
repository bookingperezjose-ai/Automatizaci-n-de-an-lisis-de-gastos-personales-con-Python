"""
Procesa transactions.csv: valida clasificación, detecta outliers (media + 2 std),
calcula KPIs mensuales, proyección simple y genera expenses_clean.csv.
"""
import pandas as pd
import numpy as np

def main():
    df = pd.read_csv("transactions.csv", encoding="utf-8-sig", parse_dates=["Fecha"])
    df["Mes"] = df["Fecha"].dt.strftime("%Y-%m")

    # Solo gastos (Débito) para análisis de consumo
    gastos = df[df["Tipo"] == "Débito"].copy()

    # 1. Clasificación automática — ya viene del generador, se valida consistencia
    categorias_validas = {"Alimentación", "Transporte", "Entretenimiento", "Servicios", "Otros"}
    gastos["Categoría"] = gastos["Categoría"].apply(
        lambda c: c if c in categorias_validas else "Otros"
    )

    # 2. Gasto promedio por mes
    gasto_mensual = gastos.groupby("Mes")["Monto"].sum().reset_index()
    gasto_mensual.columns = ["Mes", "Gasto_Total"]
    promedio_mensual = gasto_mensual["Gasto_Total"].mean()

    # 3. Top 10 gastos más altos
    top10 = gastos.nlargest(10, "Monto")[["Fecha", "Descripción", "Categoría", "Monto"]]

    # 4. Comparativa mes a mes (variación %)
    gasto_mensual["Variación_%"] = gasto_mensual["Gasto_Total"].pct_change().replace([float("inf"), float("-inf")], None)
    gasto_mensual["Variación_%"] = (gasto_mensual["Variación_%"] * 100).round(2)

    # 5. Detección de outliers — IQR (robusto ante asimetría positiva de gastos)
    q1 = gastos["Monto"].quantile(0.25)
    q3 = gastos["Monto"].quantile(0.75)
    iqr = q3 - q1
    umbral = q3 + 1.5 * iqr
    outliers = gastos[gastos["Monto"] > umbral].copy()
    outliers["Umbral_aplicado"] = round(umbral, 2)

    # 6. Proyección simple próximo mes (media móvil de últimos 3 meses)
    ultimos_3 = gasto_mensual["Gasto_Total"].tail(3)
    proyeccion_prox_mes = ultimos_3.mean()

    # 7. Métrica de ahorro: diferencia entre gasto máximo y mínimo (a nivel transacción individual)
    ahorro_potencial = gastos["Monto"].max() - gastos["Monto"].min()

    # --- Resumen KPIs ---
    kpis = {
        "Gasto_Total_6meses": round(gastos["Monto"].sum(), 2),
        "Gasto_Promedio_Mensual": round(promedio_mensual, 2),
        "Gasto_Máximo_Transacción": round(gastos["Monto"].max(), 2),
        "Gasto_Mínimo_Transacción": round(gastos["Monto"].min(), 2),
        "Diferencia_Max_Min (Ahorro potencial)": round(ahorro_potencial, 2),
        "Proyección_Próximo_Mes": round(proyeccion_prox_mes, 2),
        "Umbral_Outlier (Q3+1.5*IQR)": round(umbral, 2),
        "N_Outliers_Detectados": len(outliers),
        "N_Transacciones_Totales": len(gastos),
    }

    # Exportar datos procesados
    gastos.to_csv("expenses_clean.csv", index=False, encoding="utf-8-sig")
    gasto_mensual.to_csv("monthly_summary.csv", index=False, encoding="utf-8-sig")
    outliers.to_csv("outliers_detected.csv", index=False, encoding="utf-8-sig")
    top10.to_csv("top10_expenses.csv", index=False, encoding="utf-8-sig")
    pd.Series(kpis).to_csv("kpis_summary.csv", header=["Valor"])

    print("=== KPIs ===")
    for k, v in kpis.items():
        print(f"{k}: {v}")
    print(f"\nOutliers detectados: {len(outliers)} transacciones > R$ {umbral:.2f}")
    print("\nGenerados: expenses_clean.csv, monthly_summary.csv, outliers_detected.csv, top10_expenses.csv, kpis_summary.csv")

if __name__ == "__main__":
    main()
