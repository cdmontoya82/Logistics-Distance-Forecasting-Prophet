import pandas as pd
import numpy as np
from datetime import timedelta

# 1. Generate Mock Fleet Data (Model Structure)
def load_flota(path: str, group_col: str) -> pd.DataFrame:
    df = pd.read_excel(path)

    # Ensure column names match the expected input format
    required = {"fecha", "km_total", group_col}
    missing = required - set(df.columns)

    if missing:
        raise ValueError(
            f"Missing columns: {missing}. Found: {list(df.columns)}"
        )

    # Data type conversion and cleaning
    df["fecha"] = pd.to_datetime(df["fecha"], errors="coerce")
    df["km_total"] = pd.to_numeric(df["km_total"], errors="coerce")
    df.dropna(subset=["fecha", "km_total"], inplace=True)

    # Trim whitespace in the grouping column (e.g., Bogotá, Medellín, etc.)
    df[group_col] = df[group_col].astype(str).str.strip()

    return df

# 2. Generate Holidays Data (Required format for Facebook Prophet)
def generate_festividades():
    festividades = pd.DataFrame({
      'holiday': 'national_holiday',
      'ds': pd.to_datetime(['2024-01-01', '2024-12-25', '2025-01-01', '2025-12-25']),
      'lower_window': 0,
      'upper_window': 1,
    })
    festividades.to_excel('dummy_festividades.xlsx', index=False)
    print("File generated: dummy_festividades.xlsx")

# 3. Generate Monthly Units (Model Regressor)
def generate_unidades():
    meses = pd.date_range(start='2024-01-01', end='2025-12-31', freq='MS')
    unidades = [np.random.randint(1000, 5000) for _ in range(len(meses))]

    df_unidades = pd.DataFrame({
        'fecha': meses,
        'unidades_vendidas': unidades
    })
    df_unidades.to_excel('dummy_unidades.xlsx', index=False)
    print("File generated: dummy_unidades.xlsx")

if __name__ == "__main__":
    # Ensure generate_flota() is defined or replace with your mock data logic
    # generate_flota() 
    generate_festividades()
    generate_unidades()
    print("Data simulation ready for testing!")
