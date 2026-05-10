import pandas as pd
import numpy as np
from datetime import timedelta

# 1. Generar Datos de Flota Falsos (Estructura para el Modelo)
def generate_flota():
    fechas = pd.date_range(start='2024-01-01', end='2025-12-31', freq='D')
    placas = ['ABC-123', 'DEF-456', 'GHI-789']
    
    datos = []
    for placa in placas:
        for fecha in fechas:
            # Kilometraje base con un poco de ruido aleatorio para simular realidad
            km = np.random.normal(loc=150, scale=30)
            if km > 0:
                datos.append({'fecha': fecha, 'km_total': km, 'placa': placa})
                
    df_flota = pd.DataFrame(datos)
    df_flota.to_excel('dummy_flota.xlsx', index=False)
    print("Archivo generado: dummy_flota.xlsx")

# 2. Generar Festividades (Formato requerido por Prophet)
def generate_festividades():
    festividades = pd.DataFrame({
      'holiday': 'festivo_nacional',
      'ds': pd.to_datetime(['2024-01-01', '2024-12-25', '2025-01-01', '2025-12-25']),
      'lower_window': 0,
      'upper_window': 1,
    })
    festividades.to_excel('dummy_festividades.xlsx', index=False)
    print("Archivo generado: dummy_festividades.xlsx")

# 3. Generar Unidades Mensuales (Regresor del modelo)
def generate_unidades():
    meses = pd.date_range(start='2024-01-01', end='2025-12-31', freq='MS')
    unidades = [np.random.randint(1000, 5000) for _ in range(len(meses))]
    
    df_unidades = pd.DataFrame({
        'fecha': meses,
        'unidades_vendidas': unidades
    })
    df_unidades.to_excel('dummy_unidades.xlsx', index=False)
    print("Archivo generado: dummy_unidades.xlsx")

if __name__ == "__main__":
    generate_flota()
    generate_festividades()
    generate_unidades()
    print("¡Simulación de datos lista para pruebas!")
