"""
Script de normalización e inyección automatizada de datos de Aulas Virtuales.
Convierte 'course-report-table.xlsx' en JSON estructurado y actualiza el HTML.

Uso:
    python update_dashboard.py
    python update_dashboard.py --excel ruta/al/archivo.xlsx --html dashboard_de_aulas_virtuales_2026.html
"""

import sys
import os
import re
import json
import argparse
import datetime
import zipfile
import xml.etree.ElementTree as ET

ORG_MAP = {
    'DIR. MODERNIZACION': 'Dirección de Modernización',
    'DIR. EPIDEMIOLOGIA': 'Dirección de Epidemiología',
    'DIR. MATERNIDAD': 'Dirección de Maternidad e Infancia',
    'SALUD MENTAL': 'Dirección de Salud Mental',
    'EDUCACION PARA LA SALUD': 'Depto. Educación para la Salud',
    'LAB.NUCLEO': 'Laboratorio Núcleo (LABNU)',
    'CIENCIA Y TECNICA': 'DICYT (Investigación, Ciencia y Técnica)',
    'BIOESTADISTICA': 'Departamento de Bioestadística',
    'DIR. H.ALIMENTOS': 'Depto. Higiene e Inocuidad Alimentaria',
    'CAEP- CURSO ANUAL DE ENFERMERIA PRACTICA': 'CAEP (Enfermería Práctica)',
    'SUBSECRETARIA': 'Subsecretaría de Salud',
    'INMUNIZACIONES': 'Depto. Inmunizaciones',
    'Inmunizaciones': 'Depto. Inmunizaciones',
    'PLANIFICACION': 'Dirección de Planificación',
    'DIR. DISCAPACIDAD': 'Dirección de Discapacidad',
    'EXTERNO': 'Organismos Externos'
}

MESES = {
    'ene': '01', 'jan': '01',
    'feb': '02',
    'mar': '03',
    'abr': '04', 'apr': '04',
    'may': '05',
    'jun': '06',
    'jul': '07',
    'ago': '08', 'aug': '08',
    'sep': '09', 'sept': '09',
    'oct': '10',
    'nov': '11',
    'dic': '12', 'dec': '12'
}

def parse_excel_date_serial(serial_num):
    try:
        base = datetime.date(1899, 12, 30)
        dt = base + datetime.timedelta(days=int(float(serial_num)))
        return dt.strftime('%Y-%m-%d'), dt.strftime('%Y-%m')
    except Exception:
        return None, None

def parse_text_date(txt):
    if not txt or not isinstance(txt, str):
        return None, None
    txt = txt.strip()
    if txt.upper() in ['NO TIENE', 'SIN FECHA INICIO', '']:
        return None, None
    
    m_iso = re.match(r'(\d{4})-(\d{2})-(\d{2})', txt)
    if m_iso:
        return f"{m_iso.group(1)}-{m_iso.group(2)}-{m_iso.group(3)}", f"{m_iso.group(1)}-{m_iso.group(2)}"
    
    m_txt = re.match(r'([a-zA-Z]+)\s+(\d{1,2}),?\s+(\d{4})', txt)
    if m_txt:
        mes_str = m_txt.group(1).lower()
        dia = int(m_txt.group(2))
        anio = m_txt.group(3)
        mes_num = MESES.get(mes_str, MESES.get(mes_str[:3]))
        if mes_num:
            return f"{anio}-{mes_num}-{dia:02d}", f"{anio}-{mes_num}"
    
    return None, None

def normalize_text(val):
    if val is None:
        return ''
    return str(val).strip()

def col_str_to_idx(col_str):
    exp = 0
    idx = 0
    for char in reversed(col_str.upper()):
        idx += (ord(char) - ord('A') + 1) * (26 ** exp)
        exp += 1
    return idx - 1

def extract_rows_from_xlsx(filepath):
    with zipfile.ZipFile(filepath, 'r') as z:
        sst = []
        if 'xl/sharedStrings.xml' in z.namelist():
            tree = ET.fromstring(z.read('xl/sharedStrings.xml'))
            sst = [''.join(node.itertext()) for node in tree.findall('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}si')]
        
        sheet_path = 'xl/worksheets/sheet1.xml'
        tree = ET.fromstring(z.read(sheet_path))
        rows = tree.findall('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheetData/{http://schemas.openxmlformats.org/spreadsheetml/2006/main}row')
        
        extracted = []
        for r in rows:
            cells = r.findall('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c')
            if not cells:
                continue
            row_dict = {}
            for c in cells:
                r_ref = c.get('r', '')
                col_letters = re.sub(r'[0-9]', '', r_ref)
                t = c.get('t')
                v = c.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}v')
                val = v.text if v is not None else ''
                if t == 's' and val.isdigit() and int(val) < len(sst):
                    val = sst[int(val)]
                row_dict[col_letters] = val
                
            cols_sorted = sorted(row_dict.keys(), key=col_str_to_idx)
            if not cols_sorted:
                continue
            max_col_idx = col_str_to_idx(cols_sorted[-1])
            # Construct dense row array
            dense_row = [''] * (max_col_idx + 1)
            for k, val in row_dict.items():
                dense_row[col_str_to_idx(k)] = val
                
            # Filter leading empty columns if table starts at col B
            while dense_row and dense_row[0] == '':
                dense_row.pop(0)
            if any(dense_row):
                extracted.append(dense_row)
        return extracted

def normalize_records(raw_rows):
    if not raw_rows:
        return []
    
    # Header check
    header = [h.strip().upper() for h in raw_rows[0]]
    data_rows = raw_rows[1:]
    
    normalized = []
    record_id = 1
    
    for row in data_rows:
        if not any(row):
            continue
        aula = normalize_text(row[0] if len(row) > 0 else '')
        if not aula or aula.lower() in ['aula', 'nombre']:
            continue
        
        tipo_raw = normalize_text(row[1] if len(row) > 1 else '').upper()
        tipo = 'Capsula' if 'CAPSULA' in tipo_raw else 'Curso'
        
        cat_raw = normalize_text(row[2] if len(row) > 2 else '')
        categoria = ORG_MAP.get(cat_raw, ORG_MAP.get(cat_raw.upper(), cat_raw))
        
        acc_raw = normalize_text(row[3] if len(row) > 3 else '').upper()
        acceso = 'Libre' if 'LIBRE' in acc_raw else 'Solo cursantes'
        
        part_raw = normalize_text(row[4] if len(row) > 4 else '')
        sdate_raw = normalize_text(row[5] if len(row) > 5 else '')
        estado_raw = normalize_text(row[6] if len(row) > 6 else '').upper()
        
        # Anomaly detection (DESEPREC row where date is in part_raw)
        if any(m in part_raw.lower() for m in ['oct', 'sep', 'nov', 'dic', '2026']) and not sdate_raw.isdigit():
            if estado_raw == '' and sdate_raw in ['NO INICIADO', 'INICIADO', 'TERMINADO']:
                estado_raw = sdate_raw
                sdate_raw = part_raw
                part_raw = ''
        
        # Participants parsing
        participantes = None
        cleaned_part = re.sub(r'[^\d.]', '', part_raw)
        if cleaned_part:
            try:
                participantes = float(cleaned_part)
            except ValueError:
                participantes = None
                
        # Estado normalization
        if 'PERMANENTE' in estado_raw:
            estado = 'Permanente'
        elif 'INICIADO' in estado_raw and 'NO' not in estado_raw:
            estado = 'Iniciado'
        elif 'TERMINADO' in estado_raw or 'FINALIZADO' in estado_raw:
            estado = 'Terminado'
        elif 'NO INICIADO' in estado_raw:
            estado = 'No Iniciado'
        elif 'SIN FECHA' in estado_raw:
            estado = 'Sin Fecha Inicio'
        else:
            estado = 'Sin Fecha Inicio' if not sdate_raw or sdate_raw == 'SIN FECHA INICIO' else 'No Iniciado'
            
        # Date parsing
        fecha_iso = None
        mes_iso = None
        if sdate_raw.replace('.', '', 1).isdigit() and len(sdate_raw) >= 5:
            fecha_iso, mes_iso = parse_excel_date_serial(sdate_raw)
        else:
            fecha_iso, mes_iso = parse_text_date(sdate_raw)
            
        if estado == 'Permanente' and not sdate_raw.replace('.', '', 1).isdigit():
            fecha_iso = None
            mes_iso = None
        if estado == 'Sin Fecha Inicio':
            fecha_iso = None
            mes_iso = None
            
        # Raw date representation
        fecha_repr = sdate_raw
        if not fecha_repr:
            fecha_repr = 'NO TIENE' if estado == 'Permanente' else 'SIN FECHA INICIO'
        elif fecha_iso and sdate_raw.isdigit():
            fecha_repr = fecha_iso
            
        record = {
            'id': record_id,
            'aula': aula,
            'tipo': tipo,
            'categoria': categoria,
            'acceso': acceso,
            'participantes': participantes,
            'participantes_raw': part_raw,
            'fecha': fecha_iso,
            'mes': mes_iso,
            'fecha_raw': fecha_repr,
            'estado': estado
        }
        normalized.append(record)
        record_id += 1
        
    return normalized

def update_html(html_path, records):
    with open(html_path, 'r', encoding='utf-8') as f:
        content = f.read()
        
    json_str = json.dumps(records, ensure_ascii=False)
    
    pattern = r'(<script id="data-holder" type="application/json">)(.*?)(</script>)'
    replacement = rf'\g<1>\n{json_str}\n\g<3>'
    
    new_content, count = re.subn(pattern, replacement, content, flags=re.DOTALL)
    if count == 0:
        print("[ERROR] No se encontro el bloque <script id='data-holder'> en el HTML.")
        return False
        
    new_content = re.sub(
        r'<span id="total-label">\d+</span>',
        f'<span id="total-label">{len(records)}</span>',
        new_content
    )
    
    new_content = re.sub(
        r'<span class="status-info" id="status-info">[^<]*</span>',
        f'<span class="status-info" id="status-info">Base consolidada: {len(records)} aulas activas</span>',
        new_content
    )
    
    with open(html_path, 'w', encoding='utf-8') as f:
        f.write(new_content)
        
    return True

def main():
    parser = argparse.ArgumentParser(description="Actualizador del Dashboard de Aulas Virtuales 2026")
    parser.add_argument('--excel', default='course-report-table.xlsx', help="Ruta al archivo Excel de reporte")
    parser.add_argument('--html', default='dashboard_de_aulas_virtuales_2026.html', help="Ruta al archivo HTML del dashboard")
    args = parser.parse_args()
    
    if not os.path.exists(args.excel):
        print(f"[ERROR] Archivo Excel no encontrado: {args.excel}")
        sys.exit(1)
    if not os.path.exists(args.html):
        print(f"[ERROR] Archivo HTML no encontrado: {args.html}")
        sys.exit(1)
        
    print(f"[*] Leyendo y parseando: {args.excel}...")
    raw_rows = extract_rows_from_xlsx(args.excel)
    print(f"[+] Total de filas crudas leidas: {len(raw_rows)}")
    
    records = normalize_records(raw_rows)
    print(f"[+] Registros normalizados: {len(records)}")
    
    total_part = sum(r['participantes'] for r in records if r['participantes'] is not None)
    cursos = sum(1 for r in records if r['tipo'] == 'Curso')
    capsulas = sum(1 for r in records if r['tipo'] == 'Capsula')
    
    print(f"    - Cursos: {cursos}")
    print(f"    - Capsulas: {capsulas}")
    print(f"    - Matricula total acumulada: {int(total_part):,}".replace(',', '.'))
    
    print(f"[*] Actualizando archivo HTML: {args.html}...")
    if update_html(args.html, records):
        print("[OK] Dashboard HTML actualizado exitosamente.")
    else:
        print("[ERROR] Error durante la actualizacion del HTML.")
        sys.exit(1)

if __name__ == '__main__':
    main()
