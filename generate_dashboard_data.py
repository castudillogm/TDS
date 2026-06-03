import pandas as pd
import numpy as np
import json
import itertools
import statsmodels.api as sm

# 1. Cargar datos
file_path = r'c:\Users\castudillo\Documents\Proyectos Antigravity\TDS\Exportaciones - Volumetrica.xlsx'
df = pd.read_excel(file_path)

# 2. Agregar columnas auxiliares
def get_day_from_date(fecha):
    if pd.isna(fecha) or str(fecha).strip().lower() in ['', 'nan', 'nat']:
        return 0
    try:
        fecha_str = str(fecha).strip()
        date_part = fecha_str.split()[0]
        if len(date_part) >= 6:
            yy = int(date_part[0:2])
            mm = int(date_part[2:4])
            dd = int(date_part[4:6])
            dt = pd.to_datetime(f"20{yy:02d}-{mm:02d}-{dd:02d}")
            return dt.dayofweek
    except Exception as e:
        pass
    return 0

df['Dia_Int'] = df['FechaExpedicion'].apply(get_day_from_date)
dias_map = {0: 'Lunes', 1: 'Martes', 2: 'Miércoles', 3: 'Jueves', 4: 'Viernes', 5: 'Sábado', 6: 'Domingo'}
df['Dia'] = df['Dia_Int'].map(dias_map)
df['ADR_Binario'] = df['ADR'].apply(lambda x: 1 if str(x).upper() == 'SI' else 0)

plazas = ['ALL'] + list(df['PlazaOrigen'].unique())
fases = ['ALL'] + list(df['Fase'].unique())
adrs = ['ALL'] + list(df['ADR'].unique())

dias_base = ['Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado', 'Domingo']
dias_combos = []
for r in range(1, len(dias_base) + 1):
    for combo in itertools.combinations(dias_base, r):
        sorted_combo = sorted(list(combo))
        combo_name = ",".join(sorted_combo)
        dias_combos.append((combo_name, set(combo)))

dias_combos.append(('ALL', set(dias_base)))

combinations_stats = {}
print("Calculando combinaciones...")

for p in plazas:
    for f in fases:
        for a in adrs:
            sub_base = df
            if p != 'ALL': sub_base = sub_base[sub_base['PlazaOrigen'] == p]
            if f != 'ALL': sub_base = sub_base[sub_base['Fase'] == f]
            if a != 'ALL': sub_base = sub_base[sub_base['ADR'] == a]
            
            if len(sub_base) == 0:
                continue
                
            for combo_name, combo_set in dias_combos:
                sub_df = sub_base[sub_base['Dia'].isin(combo_set)]
                key = f"Plaza:{p}|Fase:{f}|ADR:{a}|Dia:{combo_name}"
                
                total = len(sub_df)
                if total == 0:
                    combinations_stats[key] = None
                    continue
                
                q1 = float(sub_df['DuracionMinutos'].quantile(0.25))
                q2 = float(sub_df['DuracionMinutos'].quantile(0.50))
                q3 = float(sub_df['DuracionMinutos'].quantile(0.75))
                iqr = q3 - q1
                lim_tukey = q3 + 1.5 * iqr
                
                clean_df = sub_df[sub_df['DuracionMinutos'] <= lim_tukey]
                limpios = len(clean_df)
                aberrantes = total - limpios
                
                pct_limpios = (limpios / total) * 100 if total > 0 else 0
                pct_aberrantes = 100 - pct_limpios
                
                min_limpio = float(clean_df['DuracionMinutos'].quantile(0.05)) if limpios > 0 else 0.0
                media_limpia = float(clean_df['DuracionMinutos'].mean()) if limpios > 0 else 0.0
                max_limpio = float(clean_df['DuracionMinutos'].quantile(0.95)) if limpios > 0 else 0.0
                
                combinations_stats[key] = {
                    'Total_Registros': total,
                    'Registros_Limpios': limpios,
                    'Registros_Aberrantes': aberrantes,
                    'Percent_Limpios': round(pct_limpios, 2),
                    'Percent_Aberrantes': round(pct_aberrantes, 2),
                    'Q1': round(q1, 2),
                    'Mediana': round(q2, 2),
                    'Q3': round(q3, 2),
                    'IQR': round(iqr, 2),
                    'Limite_Tukey': round(lim_tukey, 2),
                    'Min_Limpio': round(min_limpio, 2),
                    'Media_Limpia': round(media_limpia, 2),
                    'Max_Limpio': round(max_limpio, 2)
                }

reg_data = []
dias_para_dummies = ['Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado', 'Domingo']

for plaza in df['PlazaOrigen'].unique():
    for fase in ['RECEPCION MERCANCIA ALMACEN ORIGEN', 'CARGA ALMACEN ORIGEN', 'DOCUMENTACION']:
        subset = df[(df['PlazaOrigen'] == plaza) & (df['Fase'] == fase)].copy()
        if len(subset) < 10: continue
            
        q1 = subset['DuracionMinutos'].quantile(0.25)
        q3 = subset['DuracionMinutos'].quantile(0.75)
        iqr = q3 - q1
        lim = q3 + 1.5 * iqr
        clean_data = subset[subset['DuracionMinutos'] <= lim].copy()
        
        if len(clean_data) < 10: continue
            
        y = clean_data['DuracionMinutos']
        
        for d in dias_para_dummies:
            clean_data[f'Dia_{d}'] = (clean_data['Dia'] == d).astype(int)
            
        features = ['ADR_Binario'] + [f'Dia_{d}' for d in dias_para_dummies]
        if fase in ['RECEPCION MERCANCIA ALMACEN ORIGEN', 'CARGA ALMACEN ORIGEN']:
            features.insert(0, 'NumPartidas')
            
        X = clean_data[features]
        X = sm.add_constant(X)
        try:
            model = sm.OLS(y, X).fit()
            res = {
                'Delegacion': plaza,
                'Fase': fase,
                'R2': round(float(model.rsquared), 5),
                'Intercepto': round(float(model.params.get('const', 0)), 4),
                'Coef_ADR': round(float(model.params.get('ADR_Binario', 0)), 4),
                'P_Value_ADR': round(float(model.pvalues.get('ADR_Binario', 1.0)), 4),
                'Coef_Partidas': round(float(model.params.get('NumPartidas', 0)), 4) if 'NumPartidas' in X.columns else None,
                'P_Value_Partidas': round(float(model.pvalues.get('NumPartidas', 1.0)), 4) if 'NumPartidas' in X.columns else None
            }
            for d in dias_para_dummies:
                res[f'Coef_Dia_{d}'] = round(float(model.params.get(f'Dia_{d}', 0)), 4)
                res[f'P_Value_Dia_{d}'] = round(float(model.pvalues.get(f'Dia_{d}', 1.0)), 4)
                
            reg_data.append(res)
        except Exception as e:
            print(f"Error en regresión para {plaza} {fase}: {e}")

resultados_file = r'c:\Users\castudillo\Documents\Proyectos Antigravity\TDS\Resultados_Analisis.xlsx'
partidas_df = pd.read_excel(resultados_file, 'Stats_Por_Partida')
partidas_data = partidas_df.to_dict(orient='records')

output_js = f"""// Datos autogenerados por generate_dashboard_data.py
const DASHBOARD_STATS = {json.dumps(combinations_stats, indent=2)};
const REGRESIONES_STATS = {json.dumps(reg_data, indent=2)};
const PARTIDAS_STATS = {json.dumps(partidas_data, indent=2)};
"""

with open(r'c:\Users\castudillo\Documents\Proyectos Antigravity\TDS\dashboard_data_v2.js', 'w', encoding='utf-8') as f:
    f.write(output_js)

print("Datos del dashboard generados con éxito en 'dashboard_data_v2.js'.")
