import pandas as pd
import json
import requests
import os

def download_file_from_google_drive(id, destination):
    URL = "https://docs.google.com/uc?export=download"
    session = requests.Session()
    response = session.get(URL, params={'id': id, 'confirm': 't'}, stream=True)
    
    token = None
    for key, value in response.cookies.items():
        if key.startswith('download_warning'):
            token = value
            break

    if token:
        params = {'id': id, 'confirm': token}
        response = session.get(URL, params=params, stream=True)

    with open(destination, "wb") as f:
        for chunk in response.iter_content(32768):
            if chunk: 
                f.write(chunk)

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

def clasificar_zona(row):
    destino = str(row.get('PlazaDestino', '')).upper().strip()
    origen = str(row.get('PlazaOrigen', '')).upper().strip()
    
    baleares_codes = ['PMI', 'MAH', 'IBZ']
    canarias_codes = ['LPA', 'ACE', 'TFN', 'TFS', 'TCI', 'FUE', 'SPC', 'SCT']
    
    if destino in baleares_codes or origen in baleares_codes:
        return 'Baleares'
    if destino in canarias_codes or origen in canarias_codes:
        return 'Canarias'
        
    return 'Península'

def main():
    print("Downloading Excel dataset from Google Drive...")
    link = input("Pegue el enlace de Google Drive del archivo Excel (o presione Enter para usar el guardado): ").strip()
    if not link:
        file_id = '1klPa3Xkor3zdT8F5k6k63WR9-gHr4Kx6'
    else:
        import re
        match = re.search(r'/d/([a-zA-Z0-9_-]+)', link)
        if match:
            file_id = match.group(1)
        else:
            match = re.search(r'id=([a-zA-Z0-9_-]+)', link)
            file_id = match.group(1) if match else link

    temp_file = 'temp_dataset.xlsx'
    download_file_from_google_drive(file_id, temp_file)
    print("Loading Excel dataset...")
    df = pd.read_excel(temp_file)
    
    print("Cleaning dates...")
    dias_map = {0: 'Lunes', 1: 'Martes', 2: 'Miércoles', 3: 'Jueves', 4: 'Viernes', 5: 'Sábado', 6: 'Domingo'}
    
    # Fast date parsing
    if 'FechaExpedicion' in df.columns:
        df['Fecha_Limpia'] = df['FechaExpedicion'].astype(str).str.split().str[0]
        mask = df['Fecha_Limpia'].str.len() >= 6
        df.loc[mask, 'Year'] = '20' + df.loc[mask, 'Fecha_Limpia'].str[0:2]
        df.loc[mask, 'Month'] = df.loc[mask, 'Fecha_Limpia'].str[2:4]
        df.loc[mask, 'Day'] = df.loc[mask, 'Fecha_Limpia'].str[4:6]
        df['dt'] = pd.to_datetime(df['Year'] + '-' + df['Month'] + '-' + df['Day'], errors='coerce')
        df['Dia_Int'] = df['dt'].dt.dayofweek.fillna(0).astype(int)
        df['Dia_Grupo'] = df['Dia_Int'].map(dias_map)
    else:
        df['Dia_Grupo'] = 'Lunes' # Fallback
    
    print("Classifying zones...")
    df['Zona'] = df.apply(clasificar_zona, axis=1)
    
    # Clean dataset
    cols_to_keep = ['PlazaOrigen', 'PlazaDestino', 'Zona', 'Fase', 'FasePadre', 'DuracionMinutos', 'NumPartidas', 'ADR', 'Dia_Grupo', 'OrigenDestino']
    cols_to_keep = [c for c in cols_to_keep if c in df.columns]
    
    clean_data = df[cols_to_keep].copy()
    if 'DuracionMinutos' in clean_data.columns:
        clean_data = clean_data.dropna(subset=['DuracionMinutos'])
    
    plazas_origen = list(clean_data['PlazaOrigen'].fillna('ALL').astype(str).unique()) if 'PlazaOrigen' in clean_data else ['ALL']
    plazas_destino = list(clean_data['PlazaDestino'].fillna('ALL').astype(str).unique()) if 'PlazaDestino' in clean_data else ['ALL']
    zonas = list(clean_data['Zona'].fillna('ALL').astype(str).unique()) if 'Zona' in clean_data else ['ALL']
    fases = list(clean_data['Fase'].fillna('ALL').astype(str).unique()) if 'Fase' in clean_data else ['ALL']
    fase_padres = list(clean_data['FasePadre'].fillna('ALL').astype(str).unique()) if 'FasePadre' in clean_data else ['ALL']
    origenes_destinos = list(clean_data['OrigenDestino'].fillna('ALL').astype(str).unique()) if 'OrigenDestino' in clean_data else ['ALL']
    adrs = ['NO', 'SI']
    dias = ['Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado', 'Domingo']
    
    plaza_orig_map = {p: i for i, p in enumerate(plazas_origen)}
    plaza_dest_map = {p: i for i, p in enumerate(plazas_destino)}
    zona_map = {z: i for i, z in enumerate(zonas)}
    fase_map = {f: i for i, f in enumerate(fases)}
    fase_padre_map = {f: i for i, f in enumerate(fase_padres)}
    od_map = {od: i for i, od in enumerate(origenes_destinos)}
    dia_map = {d: i for i, d in enumerate(dias)}
    
    records_raw = []
    for idx, row in clean_data.iterrows():
        try:
            duracion = float(row.get('DuracionMinutos', 0))
            partidas = int(row.get('NumPartidas', 1)) if pd.notna(row.get('NumPartidas')) else 1
            po_idx = plaza_orig_map.get(str(row.get('PlazaOrigen', '')).strip(), 0)
            pd_idx = plaza_dest_map.get(str(row.get('PlazaDestino', '')).strip(), 0)
            z_idx = zona_map.get(str(row.get('Zona', '')).strip(), 0)
            f_idx = fase_map.get(str(row.get('Fase', '')).strip(), 0)
            fp_idx = fase_padre_map.get(str(row.get('FasePadre', '')).strip(), 0)
            od_idx = od_map.get(str(row.get('OrigenDestino', '')).strip(), 0)
            
            adr_val = str(row.get('ADR', '')).strip().upper()
            a_idx = 1 if adr_val == 'SI' else 0
            d_idx = dia_map.get(str(row.get('Dia_Grupo', '')).strip(), 0)
            
            records_raw.append([
                round(duracion, 2),
                partidas,
                po_idx,
                pd_idx,
                z_idx,
                f_idx,
                fp_idx,
                od_idx,
                a_idx,
                d_idx
            ])
        except Exception as e:
            continue
            
    print(f"Generated {len(records_raw)} raw records.")
    
    output_js = f"""// Datos crudos comprimidos para el dashboard interactivo
const DASHBOARD_STATS = {{}};
const REGRESIONES_STATS = [];
const PARTIDAS_STATS = [];

const MAP_PLAZA_ORIGEN = {json.dumps(plazas_origen)};
const MAP_PLAZA_DESTINO = {json.dumps(plazas_destino)};
const MAP_ZONA = {json.dumps(zonas)};
const MAP_FASE = {json.dumps(fases)};
const MAP_FASE_PADRE = {json.dumps(fase_padres)};
const MAP_ORIGEN_DESTINO = {json.dumps(origenes_destinos)};
const MAP_ADR = ["NO", "SI"];
const MAP_DIA = {json.dumps(dias)};

// Formato de registro: [Duracion, Partidas, PlazaOrigen, PlazaDestino, Zona, Fase, FasePadre, OrigenDestino, ADR, Dia]
const DEFAULT_RECORDS_RAW = {json.dumps(records_raw)};
"""
    with open('dashboard_data_v2.js', 'w', encoding='utf-8') as f:
        f.write(output_js)
    
    print("Success: Generated dashboard_data_v2.js")

if __name__ == "__main__":
    main()
