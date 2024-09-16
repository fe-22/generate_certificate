import os
from flask import Flask, render_template, request, redirect, url_for, flash, send_from_directory
from fpdf import FPDF
import pandas as pd


app = Flask(__name__)

# Constants
EXCEL_FOLDER_PATH = 'data/'
PDF_FOLDER_PATH = 'data/pdfs/'
TEMPLATE_FOLDER_PATH = 'static/'

# Create directories if they don't exist
for path in [EXCEL_FOLDER_PATH, PDF_FOLDER_PATH, TEMPLATE_FOLDER_PATH]:
    os.makedirs(path, exist_ok=True)

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        excel_file = request.files.get('excel_file')
        template_file = request.files.get('template_file')

        # Verify if files were provided
        if not excel_file:
            flash("Nenhum arquivo Excel foi carregado", 'danger')
            return redirect(url_for('index'))

        if not template_file:
            flash("Nenhum template de certificado foi carregado", 'danger')
            return redirect(url_for('index'))

        # Save Excel file if provided
        excel_path = os.path.join(EXCEL_FOLDER_PATH, 'certificados.xlsx')
        excel_file.save(excel_path)
        flash('Excel file carregado com sucesso!', 'success')

        # Save template file if provided
        template_path = os.path.join(TEMPLATE_FOLDER_PATH, 'templateA4.jpeg')
        template_file.save(template_path)
        flash('Template do certificado carregado com sucesso!', 'success')

        # Read Excel file
        try:
            df = pd.read_excel(excel_path)
        except Exception as e:
            flash(f"Erro ao processar o arquivo Excel: {e}", "danger")
            return redirect(url_for('index'))

        # Check if required columns exist
        required_columns = ['Nome Completo', 'Curso', 'Professor', 'Data Início', 'Data Término']
        missing_columns = [col for col in required_columns if col not in df.columns]

        if missing_columns:
            flash(f"As seguintes colunas estão faltando no arquivo Excel: {', '.join(missing_columns)}", "danger")
            return redirect(url_for('index'))

        # Generate certificates for all students in Excel file
        for _, row in df.iterrows():
            try:
                gerar_certificado(row, template_path)
            except Exception as e:
                flash(f"Erro ao gerar certificado para {row['Nome Completo']}: {e}", "danger")
                continue

        flash('Certificados gerados com sucesso!', 'success')
        return redirect(url_for('index'))

    return render_template('index.html')

@app.route('/static/<path:path>')
def send_static(path):
    return send_from_directory('static', path)

def gerar_certificado(row, template_path):
    nome_completo = row['Nome Completo']
    curso = row['Curso']
    descricao_curso = row.get('Descrição Curso', 'Descrição não disponível')
    carga_horaria = row.get('Carga Horária', 'Carga Horária não disponível')
    professor = row['Professor']
    data_inicio = row['Data Início']
    data_termino = row['Data Término']

    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", 'B', size=15)

    if os.path.exists(template_path):
        pdf.image(template_path, x=0, y=0)
    else:
        raise FileNotFoundError("Template de certificado não encontrado.")

    pdf.set_text_color(33, 24, 136)
    
     # Centralizando o nome do aluno
    pdf.set_xy(0, 145)
    pdf.cell(w=210, h=10, txt=nome_completo, border=0, ln=1, align='C')  # Texto centralizado

    # Curso e professor
    pdf.set_xy(20, 165)
    pdf.multi_cell(w=170, h=10, txt=f"concluiu com êxito o curso {curso.upper()} ministrado por {professor.upper()} ENTRE {data_inicio} e {data_termino}, com carga horária de {carga_horaria} horas.", align='L')

    # Descrição do curso (quebrada em várias linhas)
    pdf.set_xy(20, 195)
    pdf.multi_cell(w=170, h=10, txt=descricao_curso, align='L')


 

    # Break long description into multiple lines
    y = 195
    for line in descricao_curso.split("\n"):
        pdf.text(50, y, line)
        y += 10

    nome_arquivo = f"{PDF_FOLDER_PATH}Certificado_{nome_completo.replace(' ', '_')}.pdf"
    pdf.output(nome_arquivo)

if __name__ == '__main__':
    app.run(debug=True)
