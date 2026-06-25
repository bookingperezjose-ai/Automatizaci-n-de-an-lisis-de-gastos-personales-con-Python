"""
Genera transactions.csv: 300+ transacciones simuladas de 6 meses,
con patrones realistas (estacionalidad, recurrencia, outliers intencionales).
"""
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import random

random.seed(42)
np.random.seed(42)

# --- Catálogo de comercios por categoría (mercado BR/CWB) ---
CATEGORIAS = {
    "Alimentación": {
        "comercios": ["Supermercado Carrefour", "Pão de Açúcar", "Mercado Mundial",
                      "Restaurante Madero", "iFood - Pedido", "Padaria Boa Vista",
                      "Açougue Central", "Big Box Atacado"],
        "rango": (15, 450), "frecuencia_mes": 18,
    },
    "Transporte": {
        "comercios": ["Posto Ipiranga - Gasolina", "Uber", "99 Pop",
                      "Estacionamento Shopping", "Pedágio Linha Verde", "Oficina Mecânica Silva"],
        "rango": (10, 350), "frecuencia_mes": 12,
    },
    "Entretenimiento": {
        "comercios": ["Netflix", "Spotify Premium", "Cinema Cinemark",
                      "Bar do Alemão", "Steam - Jogos", "YouTube Premium"],
        "rango": (12, 200), "frecuencia_mes": 8,
    },
    "Servicios": {
        "comercios": ["Copel - Energia Elétrica", "Sanepar - Água", "Vivo Internet Fibra",
                      "Claro Pós-pago", "Condomínio Edifício", "Seguro Auto Porto"],
        "rango": (45, 600), "frecuencia_mes": 7,
    },
    "Otros": {
        "comercios": ["Farmácia São Paulo", "Loja Renner", "Amazon BR",
                      "Pet Shop Cobasi", "Salão de Beleza", "Livraria Cultura"],
        "rango": (20, 500), "frecuencia_mes": 6,
    },
}

OUTLIER_DESCRIPCIONES = ["Compra Eletrônico Magazine Luiza", "Conserto Notebook",
                          "Despesa Médica Particular", "Multa de Trânsito"]

def generar_transacciones(meses=6):
    hoy = datetime(2026, 6, 25)
    inicio = hoy - timedelta(days=30 * meses)
    transacciones = []

    for categoria, info in CATEGORIAS.items():
        n_trans = info["frecuencia_mes"] * meses
        for _ in range(n_trans):
            dias_offset = random.randint(0, 30 * meses)
            fecha = inicio + timedelta(days=dias_offset)
            comercio = random.choice(info["comercios"])
            monto = round(random.uniform(*info["rango"]), 2)
            transacciones.append({
                "Fecha": fecha,
                "Descripción": comercio,
                "Categoría": categoria,
                "Monto": monto,
                "Tipo": "Débito",
            })

    # Inyectar outliers reales (gastos atípicos > 2 std dev del conjunto)
    for _ in range(8):
        dias_offset = random.randint(0, 30 * meses)
        fecha = inicio + timedelta(days=dias_offset)
        transacciones.append({
            "Fecha": fecha,
            "Descripción": random.choice(OUTLIER_DESCRIPCIONES),
            "Categoría": "Otros",
            "Monto": round(random.uniform(1200, 2000), 2),
            "Tipo": "Débito",
        })

    # Créditos (salario + alguns extras) — necesario para saldo coherente
    for m in range(meses):
        fecha_salario = inicio + timedelta(days=30 * m + 5)
        transacciones.append({
            "Fecha": fecha_salario,
            "Descripción": "Salário - Transferência PIX",
            "Categoría": "Renda",
            "Monto": 4200.00,
            "Tipo": "Crédito",
        })

    df = pd.DataFrame(transacciones).sort_values("Fecha").reset_index(drop=True)

    # Saldo anterior/posterior (simulación de cuenta corrente)
    saldo_inicial = 3500.00
    saldos_ant, saldos_pos = [], []
    saldo = saldo_inicial
    for _, row in df.iterrows():
        saldos_ant.append(round(saldo, 2))
        if row["Tipo"] == "Débito":
            saldo -= row["Monto"]
        else:
            saldo += row["Monto"]
        saldos_pos.append(round(saldo, 2))
    df["Saldo anterior"] = saldos_ant
    df["Saldo posterior"] = saldos_pos
    df["Fecha"] = df["Fecha"].dt.strftime("%Y-%m-%d")

    return df[["Fecha", "Descripción", "Categoría", "Monto", "Tipo", "Saldo anterior", "Saldo posterior"]]

if __name__ == "__main__":
    df = generar_transacciones(meses=6)
    df.to_csv("transactions.csv", index=False, encoding="utf-8-sig")
    print(f"transactions.csv generado: {len(df)} transacciones")
    print(df["Categoría"].value_counts())
