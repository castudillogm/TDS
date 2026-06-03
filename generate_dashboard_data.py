import pandas as pd
import json

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
    except Exception:
        pass
    return 0

def main():
    print("Loading Excel dataset...")
    df = pd.read_excel('Exportaciones - Volumetrica.xlsx')
    
    print("Cleaning dates...")
    dias_map = {0: 'Lunes', 1: 'Martes', 2: 'Miércoles', 3: 'Jueves', 4: 'Viernes', 5: 'Sábado', 6: 'Domingo'}
    
    # Fast date parsing
    df['Fecha_Limpia'] = df['FechaExpedicion'].astype(str).str.split().str[0]
    mask = df['Fecha_Limpia'].str.len() >= 6
    df.loc[mask, 'Year'] = '20' + df.loc[mask, 'Fecha_Limpia'].str[0:2]
    df.loc[mask, 'Month'] = df.loc[mask, 'Fecha_Limpia'].str[2:4]
    df.loc[mask, 'Day'] = df.loc[mask, 'Fecha_Limpia'].str[4:6]
    df['dt'] = pd.to_datetime(df['Year'] + '-' + df['Month'] + '-' + df['Day'], errors='coerce')
    df['Dia_Int'] = df['dt'].dt.dayofweek.fillna(0).astype(int)
    df['Dia_Grupo'] = df['Dia_Int'].map(dias_map)
    
    # Clean dataset
    clean_data = df[['PlazaOrigen', 'Fase', 'FasePadre', 'DuracionMinutos', 'NumPartidas', 'ADR', 'Dia_Grupo']].copy()
    clean_data = clean_data.dropna(subset=['DuracionMinutos'])
    
    plazas = list(clean_data['PlazaOrigen'].fillna('ALL').astype(str).unique())
    fases = list(clean_data['Fase'].fillna('ALL').astype(str).unique())
    fase_padres = list(clean_data['FasePadre'].fillna('ALL').astype(str).unique())
    adrs = ['SI', 'NO']
    dias = ['Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado', 'Domingo']
    
    plaza_map = {p: i for i, p in enumerate(plazas)}
    fase_map = {f: i for i, f in enumerate(fases)}
    fase_padre_map = {f: i for i, f in enumerate(fase_padres)}
    adr_map = {'SI': 1, 'NO': 0, 'ALL': 0}
    dia_map = {d: i for i, d in enumerate(dias)}
    
    records_raw = []
    for idx, row in clean_data.iterrows():
        try:
            duracion = float(row['DuracionMinutos'])
            partidas = int(row['NumPartidas']) if pd.notna(row['NumPartidas']) else 1
            plaza_idx = plaza_map.get(str(row['PlazaOrigen']), 0)
            fase_idx = fase_map.get(str(row['Fase']), 0)
            fase_padre_idx = fase_padre_map.get(str(row['FasePadre']), 0)
            adr_val = str(row['ADR']).strip().upper()
            adr_idx = 1 if adr_val == 'SI' else 0
            dia_idx = dia_map.get(str(row['Dia_Grupo']), 0)
            
            records_raw.append([
                round(duracion, 2),
                partidas,
                plaza_idx,
                fase_idx,
                fase_padre_idx,
                adr_idx,
                dia_idx
            ])
        except Exception as e:
            continue
            
    print(f"Generated {len(records_raw)} raw records.")
    
    output_js = f"""// Datos crudos comprimidos para el dashboard interactivo
const DASHBOARD_STATS = {{}};
const REGRESIONES_STATS = [];
const PARTIDAS_STATS = [];

const MAP_PLAZA = {json.dumps(plazas)};
const MAP_FASE = {json.dumps(fases)};
const MAP_FASE_PADRE = {json.dumps(fase_padres)};
const MAP_ADR = ["NO", "SI"];
const MAP_DIA = {json.dumps(dias)};

const DEFAULT_RECORDS_RAW = {json.dumps(records_raw)};
"""
    with open('dashboard_data_v2.js', 'w', encoding='utf-8') as f:
        f.write(output_js)
    
    print("Success: Generated dashboard_data_v2.js")

if __name__ == "__main__":
    main()
