import os
import sys
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages


def make_report(input_dir, out_pdf):
    # locate main sheet CSV
    main_csv = None
    criterios_csv = None
    for f in os.listdir(input_dir):
        if f.lower().startswith('sheet_cm_matrix_equipos_principales') and f.lower().endswith('.csv'):
            main_csv = os.path.join(input_dir, f)
        if f.lower().startswith('sheet_criterios') and f.lower().endswith('.csv'):
            criterios_csv = os.path.join(input_dir, f)

    if main_csv is None:
        print('No se encontró el CSV principal en', input_dir)
        return 2

    df = pd.read_csv(main_csv, encoding='utf-8')

    # ensure computed score exists
    if '_computed_score' not in df.columns:
        df['_computed_score'] = 0.0

    # Prepare PDF
    with PdfPages(out_pdf) as pdf:
        # page 1: summary
        fig, ax = plt.subplots(figsize=(8.27, 11.69))  # A4
        ax.axis('off')
        total_rows = len(df)
        avg_score = df['_computed_score'].mean()
        top_scores = df['_computed_score'].nlargest(5).tolist()
        summary_text = f"Informe: Matriz de condición\n\nFilas analizadas: {total_rows}\nPuntuación media: {avg_score:.2f}\nTop 5 puntuaciones: {top_scores}"
        ax.text(0.01, 0.95, summary_text, va='top', fontsize=12, wrap=True)
        pdf.savefig(fig)
        plt.close(fig)

        # page 2: histogram of computed scores
        fig, ax = plt.subplots(figsize=(8.27, 11.69))
        ax.hist(df['_computed_score'].dropna(), bins=20, color='C0', edgecolor='black')
        ax.set_title('Distribución de _computed_score')
        ax.set_xlabel('Score')
        ax.set_ylabel('Frecuencia')
        pdf.savefig(fig)
        plt.close(fig)

        # page 3: top 10 rows table
        top10 = df.sort_values('_computed_score', ascending=False).head(10)
        # render table
        fig, ax = plt.subplots(figsize=(8.27, 11.69))
        ax.axis('off')
        tbl = ax.table(cellText=top10.fillna('').values, colLabels=top10.columns, loc='center')
        tbl.auto_set_font_size(False)
        tbl.set_fontsize(8)
        tbl.scale(1, 1.2)
        ax.set_title('Top 10 filas por _computed_score')
        pdf.savefig(fig, bbox_inches='tight')
        plt.close(fig)

    print('Reporte generado en', out_pdf)
    return 0


if __name__ == '__main__':
    input_dir = os.path.join(os.path.dirname(__file__), '..', 'analysis_output')
    input_dir = os.path.abspath(input_dir)
    out_pdf = os.path.join(input_dir, 'report.pdf')
    rc = make_report(input_dir, out_pdf)
    sys.exit(rc)
