import pandas as pd
import numpy as np
import statsmodels.api as sm
import hashlib

# 1. Cargar datos
file_path = r'c:\Users\castudillo\Documents\Proyectos Antigravity\TDS\Exportaciones - Volumetrica.xlsx'
df = pd.read_excel(file_path)

# 2. Asignar 'Día' calculando el día de la semana real a partir de FechaExpedicion
def get_day_from_date(fecha):
    if pd.isna(fecha) or str(fecha).strip().lower() in ['', 'nan', 'nat']:
        return 0
    try:
        fecha_str = str(fecha).strip()
        date_part = fecha_str.split()[0]  # "260112"
        if len(date_part) >= 6:
            yy = int(date_part[0:2])
            mm = int(date_part[2:4])
            dd = int(date_part[4:6])
            dt = pd.to_datetime(f"20{yy:02d}-{mm:02d}-{dd:02d}")
            return dt.dayofweek  # 0: Lunes, ..., 6: Domingo
    except Exception as e:
        pass
    return 0 # Fallback to Monday

df['Dia_Int'] = df['FechaExpedicion'].apply(get_day_from_date)
dias_map = {0: 'Lunes', 1: 'Martes', 2: 'Miércoles', 3: 'Jueves', 4: 'Viernes', 5: 'Sábado', 6: 'Domingo'}
df['Dia'] = df['Dia_Int'].map(dias_map)

# Variable binaria de días: 0 (L-X), 1 (J-V)
df['Dia_Binario'] = df['Dia_Int'].apply(lambda x: 1 if x >= 3 else 0)

# Variable binaria ADR
df['ADR_Binario'] = df['ADR'].apply(lambda x: 1 if str(x).upper() == 'SI' else 0)

with pd.ExcelWriter(r'c:\Users\castudillo\Documents\Proyectos Antigravity\TDS\Resultados_Analisis.xlsx') as writer:
    
    # --- ANÁLISIS DE CUARTILES Y LÍMITES DE TUKEY ---
    # Global
    def calc_stats(group):
        q1 = group['DuracionMinutos'].quantile(0.25)
        q2 = group['DuracionMinutos'].quantile(0.50)
        q3 = group['DuracionMinutos'].quantile(0.75)
        iqr = q3 - q1
        lim_tukey = q3 + 1.5 * iqr
        
        # Datos estándar (limpios)
        clean = group[group['DuracionMinutos'] <= lim_tukey]['DuracionMinutos']
        
        return pd.Series({
            'Q1': q1,
            'Mediana (Q2)': q2,
            'Q3': q3,
            'IQR': iqr,
            'Limite_Tukey': lim_tukey,
            'Mejor_Caso (Q1)': q1,
            'Caso_Comun (Mediana)': q2,
            'Peor_Caso (Max Limpio)': clean.max() if not clean.empty else np.nan,
            'Total_Registros': len(group),
            'Registros_Limpios': len(clean)
        })

    # Por Delegación y Fase
    stats_del_fase = df.groupby(['PlazaOrigen', 'Fase']).apply(calc_stats).reset_index()
    stats_del_fase.to_excel(writer, sheet_name='Stats_Global_Fase', index=False)
    
    # Por Día
    stats_dia = df.groupby(['Dia']).apply(calc_stats).reset_index()
    stats_dia.to_excel(writer, sheet_name='Stats_Por_Dia', index=False)
    
    # Por ADR
    stats_adr = df.groupby(['ADR']).apply(calc_stats).reset_index()
    stats_adr.to_excel(writer, sheet_name='Stats_Por_ADR', index=False)
    
    # Por Partida (agrupando cantidad de partidas)
    stats_part = df.groupby(['NumPartidas']).apply(calc_stats).reset_index()
    stats_part.to_excel(writer, sheet_name='Stats_Por_Partida', index=False)
    
    # --- REGRESIONES ---
    # Solo con datos limpios (<= Limite Tukey) por Fase y Delegación
    resultados_regresion = []
    
    dias_para_dummies = ['Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado', 'Domingo']
    
    for delegacion in df['PlazaOrigen'].unique():
        for fase in ['RECEPCION MERCANCIA ALMACEN ORIGEN', 'CARGA ALMACEN ORIGEN', 'DOCUMENTACION']:
            subset = df[(df['PlazaOrigen'] == delegacion) & (df['Fase'] == fase)].copy()
            if len(subset) < 10:
                continue
                
            q1 = subset['DuracionMinutos'].quantile(0.25)
            q3 = subset['DuracionMinutos'].quantile(0.75)
            iqr = q3 - q1
            lim = q3 + 1.5 * iqr
            clean_data = subset[subset['DuracionMinutos'] <= lim].copy()
            
            if len(clean_data) < 10:
                continue
                
            y = clean_data['DuracionMinutos']
            
            for d in dias_para_dummies:
                clean_data[f'Dia_{d}'] = (clean_data['Dia'] == d).astype(int)
                
            features = ['ADR_Binario'] + [f'Dia_{d}' for d in dias_para_dummies]
            if fase in ['RECEPCION MERCANCIA ALMACEN ORIGEN', 'CARGA ALMACEN ORIGEN']:
                features.insert(0, 'NumPartidas')
                
            X = clean_data[features]
            X = sm.add_constant(X)
            
            # Ajustar modelo
            try:
                model = sm.OLS(y, X).fit()
                res = {
                    'Delegacion': delegacion,
                    'Fase': fase,
                    'R2': model.rsquared,
                    'R (Correlacion)': np.sqrt(model.rsquared),
                    'Intercepto (Lunes)': model.params.get('const', 0),
                    'Coef_ADR': model.params.get('ADR_Binario', 0),
                    'P_Value_ADR': model.pvalues.get('ADR_Binario', np.nan)
                }
                
                if 'NumPartidas' in X.columns:
                    res['Coef_Partidas'] = model.params.get('NumPartidas', 0)
                    res['P_Value_Partidas'] = model.pvalues.get('NumPartidas', np.nan)
                else:
                    res['Coef_Partidas'] = np.nan
                    res['P_Value_Partidas'] = np.nan
                    
                for d in dias_para_dummies:
                    res[f'Coef_Dia_{d}'] = model.params.get(f'Dia_{d}', 0)
                    res[f'P_Value_Dia_{d}'] = model.pvalues.get(f'Dia_{d}', np.nan)
                    
                resultados_regresion.append(res)
            except Exception as e:
                pass
                
    df_reg = pd.DataFrame(resultados_regresion)
    df_reg.to_excel(writer, sheet_name='Regresiones', index=False)
    
print("Análisis completado. Archivo generado: Resultados_Analisis.xlsx")
