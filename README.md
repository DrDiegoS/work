import pandas as pd
from docx import Document
from docx.shared import Inches

# Load filtered data
csv_path = "cronicofiltrado.csv"
df = pd.read_csv(csv_path, low_memory=False)
df_60k = df[df['Custo Total'] >= 60000]

# Prepare statistics
stats = df_60k['Custo Total'].describe()
top10 = df_60k.sort_values(by='Custo Total', ascending=False)[['nome_beneficiario', 'Custo Total']].head(10)

# Cost range distribution
bins = [60000, 80000, 100000, 150000, 200000, df_60k['Custo Total'].max()]
labels = ['60k-80k', '80k-100k', '100k-150k', '150k-200k', '200k+']
df_60k['Cost Range'] = pd.cut(df_60k['Custo Total'], bins=bins, labels=labels, include_lowest=True)
cost_range_dist = df_60k['Cost Range'].value_counts().sort_index()

# Age range distribution
if 'idade' in df_60k.columns:
    age_bins = [0, 30, 45, 60, 75, 100]
    age_labels = ['<30', '30-45', '45-60', '60-75', '75+']
    df_60k['Age Range'] = pd.cut(df_60k['idade'], bins=age_bins, labels=age_labels, include_lowest=True)
    age_range_stats = df_60k.groupby('Age Range')['Custo Total'].describe()
else:
    age_range_stats = None

# Other distributions
uf_dist = df_60k['uf_beneficiario'].value_counts() if 'uf_beneficiario' in df_60k.columns else None
gender_dist = df_60k['genero'].value_counts() if 'genero' in df_60k.columns else None
plan_dist = df_60k['tipo_plano'].value_counts() if 'tipo_plano' in df_60k.columns else None
status_dist = df_60k['situacao_ativo_ou_inativo_plano'].value_counts() if 'situacao_ativo_ou_inativo_plano' in df_60k.columns else None

# Create Word report
doc = Document()
doc.add_heading('High Cost Patients Report (>= 60k)', 0)

doc.add_heading('Descriptive Statistics', level=1)
doc.add_paragraph(str(stats))

doc.add_heading('Top 10 Patients by Total Cost', level=1)
table = doc.add_table(rows=1, cols=2)
hdr_cells = table.rows[0].cells
hdr_cells[0].text = 'Patient Name'
hdr_cells[1].text = 'Total Cost'
for idx, row in top10.iterrows():
    row_cells = table.add_row().cells
    row_cells[0].text = str(row['nome_beneficiario'])
    row_cells[1].text = f"{row['Custo Total']:.2f}"

doc.add_heading('Cost Range Distribution', level=1)
table = doc.add_table(rows=1, cols=2)
hdr_cells = table.rows[0].cells
hdr_cells[0].text = 'Cost Range'
hdr_cells[1].text = 'Count'
for idx, val in cost_range_dist.items():
    row_cells = table.add_row().cells
    row_cells[0].text = str(idx)
    row_cells[1].text = str(val)

if age_range_stats is not None:
    doc.add_heading('Cost by Age Range', level=1)
    doc.add_paragraph(str(age_range_stats))

if uf_dist is not None:
    doc.add_heading('Distribution by State (UF)', level=1)
    table = doc.add_table(rows=1, cols=2)
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = 'State'
    hdr_cells[1].text = 'Count'
    for idx, val in uf_dist.items():
        row_cells = table.add_row().cells
        row_cells[0].text = str(idx)
        row_cells[1].text = str(val)

if gender_dist is not None:
    doc.add_heading('Distribution by Gender', level=1)
    table = doc.add_table(rows=1, cols=2)
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = 'Gender'
    hdr_cells[1].text = 'Count'
    for idx, val in gender_dist.items():
        row_cells = table.add_row().cells
        row_cells[0].text = str(idx)
        row_cells[1].text = str(val)

if plan_dist is not None:
    doc.add_heading('Distribution by Plan Type', level=1)
    table = doc.add_table(rows=1, cols=2)
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = 'Plan Type'
    hdr_cells[1].text = 'Count'
    for idx, val in plan_dist.items():
        row_cells = table.add_row().cells
        row_cells[0].text = str(idx)
        row_cells[1].text = str(val)

if status_dist is not None:
    doc.add_heading('Distribution by Plan Status', level=1)
    table = doc.add_table(rows=1, cols=2)
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = 'Status'
    hdr_cells[1].text = 'Count'
    for idx, val in status_dist.items():
        row_cells = table.add_row().cells
        row_cells[0].text = str(idx)
        row_cells[1].text = str(val)

# Optionally, add images if you want
doc.add_page_break()
doc.add_paragraph('See attached charts for more details.')
# Example: doc.add_picture('grafico_pareto.png', width=Inches(5))
# Example: doc.add_picture('grafico_faixa_custo.png', width=Inches(5))

doc.save('high_cost_patients_report.docx')
print('Word report high_cost_patients_report.docx generated successfully.')
