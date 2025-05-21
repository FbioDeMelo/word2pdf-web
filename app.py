from flask import Flask, render_template, request, send_file
import os
from docx2pdf import convert
import uuid
import pythoncom

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# Garante que a pasta uploads existe
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

@app.route('/')
def index():
    return render_template('index.html')
@app.route('/sobre')
def sobre():
    return render_template('sobre.html')    
@app.route('/convert', methods=['POST'])
def convert_to_pdf():
    if 'file' not in request.files:
        return 'Nenhum arquivo enviado.', 400

    files = request.files.getlist('file')

    if len(files) == 0:
        return 'Nenhum arquivo selecionado.', 400

    # Inicializa o COM (deve ser feito antes de chamar qualquer função COM)
    pythoncom.CoInitialize()

    converted_files = []

    for file in files:
        if file and file.filename.endswith('.docx'):
            # Salvar arquivo temporariamente
            unique_filename = str(uuid.uuid4())
            docx_path = os.path.join(app.config['UPLOAD_FOLDER'], unique_filename + '.docx')
            pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], unique_filename + '.pdf')

            file.save(docx_path)

            # Converter para PDF
            try:
                convert(docx_path, pdf_path)
                converted_files.append(pdf_path)  # Adiciona o caminho do arquivo convertido
            except Exception as e:
                pythoncom.CoUninitialize()  # Finaliza o COM em caso de erro
                return f"Erro ao converter {file.filename}: {e}", 500

    # Finaliza o COM
    pythoncom.CoUninitialize()

    # Enviar arquivos convertidos como um arquivo ZIP
    if len(converted_files) > 1:
        import zipfile
        zip_filename = str(uuid.uuid4()) + '.zip'
        zip_path = os.path.join(app.config['UPLOAD_FOLDER'], zip_filename)

        with zipfile.ZipFile(zip_path, 'w') as zipf:
            for pdf_file in converted_files:
                zipf.write(pdf_file, os.path.basename(pdf_file))

        # Enviar arquivo ZIP contendo todos os PDFs convertidos
        return send_file(zip_path, as_attachment=True)

    # Caso tenha apenas um arquivo convertido, envia diretamente
    return send_file(converted_files[0], as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)