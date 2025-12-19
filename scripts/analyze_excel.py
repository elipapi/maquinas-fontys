import sys
import os
import json
import re
import pandas as pd


def detect_header_row(excel_path, sheet_name, max_scan=10):
    xls = pd.read_excel(excel_path, sheet_name=sheet_name, header=None)
    for i in range(min(max_scan, len(xls))):
        row = xls.iloc[i].astype(str).str.upper().tolist()
        if any('AREA' == v or v.strip().startswith('AREA') for v in row if v and v != 'nan'):
            return i
    # fallback: try to find a row with 'CÓDIGO' or 'Código' or 'Denominación'
    for i in range(min(max_scan, len(xls))):
        row = xls.iloc[i].astype(str).str.upper().tolist()
        if any('CÓDIGO' in v or 'DENOMINACIÓN' in v or 'DENOMINACION' in v for v in row if v and v != 'nan'):
            return i
    return 0


def parse_criteria(excel_path):
    try:
        crit = pd.read_excel(excel_path, sheet_name='Criterios', header=None)
    except Exception:
        return {}
    mapping = {}
    # heuristic: for each row, find text cells and nearby numeric cells (factors)
    for _, row in crit.iterrows():
        texts = []
        factors = []
        for col_idx, val in enumerate(row):
            if pd.isna(val):
                continue
            if isinstance(val, (int, float)) and not (isinstance(val, bool)):
                factors.append((col_idx, float(val)))
            else:
                s = str(val).strip()
                if s:
                    texts.append((col_idx, s))
        # pair any nearby text with factor: nearest factor to the right or left within 3 cols
        for fcol, fval in factors:
            # find nearest text
            best = None
            best_dist = 999
            for tcol, tval in texts:
                dist = abs(tcol - fcol)
                if dist < best_dist:
                    best = tval
                    best_dist = dist
            if best:
                mapping[best] = fval
    # normalize keys
    mapping = {k.upper(): v for k, v in mapping.items()}
    return mapping


def score_row_by_criteria(row, criteria_map):
    score = 0.0
    matches = []
    joined = ' '.join([str(x) for x in row.values if not pd.isna(x)])
    joined_u = joined.upper()
    for key, factor in criteria_map.items():
        if key and key in joined_u:
            score += factor
            matches.append((key, factor))
    return score, matches


def analyze_excel(path, outdir):
    os.makedirs(outdir, exist_ok=True)

    try:
        # first detect header row for main sheet if possible
        header_row = detect_header_row(path, sheet_name='CM Matrix equipos principales')
        sheets = pd.read_excel(path, sheet_name=None, header=header_row)
    except Exception as e:
        print(f"ERROR leyendo el Excel: {e}")
        return 1

    summary = {"file": os.path.basename(path), "sheets": {}}

    # parse criterios mapping
    criteria_map = parse_criteria(path)

    for name, df in sheets.items():
        info = {}
        # normalize column names
        df.columns = [str(c).strip() for c in df.columns]
        info['rows'], info['cols'] = df.shape
        info['columns'] = list(df.columns)
        info['dtypes'] = {str(c): str(t) for c, t in df.dtypes.items()}
        info['null_counts'] = df.isnull().sum().to_dict()

        # compute score per row if criteria exist
        if criteria_map:
            scores = []
            matches_all = []
            for _, r in df.iterrows():
                sc, matches = score_row_by_criteria(r, criteria_map)
                scores.append(sc)
                matches_all.append(';'.join([f"{m[0]}:{m[1]}" for m in matches]))
            df['_computed_score'] = scores
            df['_computed_matches'] = matches_all

        # numeric summary where applicable
        try:
            info['describe'] = df.select_dtypes(include='number').describe().to_dict()
        except Exception:
            info['describe'] = {}

        # sample and save sheet to CSV
        sample_csv = os.path.join(outdir, f"sheet_{safe_filename(name)}_sample.csv")
        try:
            df.head(20).to_csv(sample_csv, index=False, encoding='utf-8')
            info['sample_csv'] = os.path.relpath(sample_csv)
        except Exception:
            info['sample_csv'] = None

        # save full sheet as CSV
        full_csv = os.path.join(outdir, f"sheet_{safe_filename(name)}.csv")
        try:
            df.to_csv(full_csv, index=False, encoding='utf-8')
            info['full_csv'] = os.path.relpath(full_csv)
        except Exception:
            info['full_csv'] = None

        summary['sheets'][name] = info

    # write summary files
    summary_json = os.path.join(outdir, 'summary.json')
    with open(summary_json, 'w', encoding='utf-8') as f:
        json.dump(summary, f, indent=2, ensure_ascii=False)

    summary_txt = os.path.join(outdir, 'summary.txt')
    with open(summary_txt, 'w', encoding='utf-8') as f:
        f.write(f"Análisis de {summary['file']}\n\n")
        for sname, s in summary['sheets'].items():
            f.write(f"Hoja: {sname}\n")
            f.write(f"  Filas: {s['rows']}, Columnas: {s['cols']}\n")
            f.write(f"  Columnas: {', '.join(s['columns'])}\n")
            f.write(f"  Tipos: {s['dtypes']}\n")
            f.write(f"  Nulos: {s['null_counts']}\n")
            f.write('\n')

    print(f"Análisis completado. Salida en: {os.path.abspath(outdir)}")
    print(f"Resumen JSON: {summary_json}")
    return 0


def safe_filename(name):
    return ''.join(c if c.isalnum() or c in (' ', '-', '_') else '_' for c in name).replace(' ', '_')


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print('Uso: analyze_excel.py <ruta_excel> [outdir]')
        sys.exit(2)
    path = sys.argv[1]
    outdir = sys.argv[2] if len(sys.argv) > 2 else os.path.join(os.path.dirname(path), 'analysis_output')
    rc = analyze_excel(path, outdir)
    sys.exit(rc)
