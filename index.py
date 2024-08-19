from flask import Flask, request, send_file, render_template
import pandas as pd
import vobject
from io import BytesIO, StringIO
import chardet

app = Flask(__name__)

# Função para dividir o número de telefone em DDD e Telefone
def split_phone_number(phone_number):
    digits = ''.join(filter(str.isdigit, phone_number))
    if len(digits) > 11:
        ddd = digits[2:4]
        ddi = digits[:2]
        telefone = digits[4:]
    else:
        ddd = digits[:2]
        ddi = "55"
        telefone = digits[2:]
    return ddd, ddi, telefone

# Função para limpar o conteúdo do arquivo VCF
def clean_vcf_content(content):
    cleaned_lines = []
    for line in content.splitlines():
        if not line.startswith('=') and not line.startswith(';'):
            cleaned_lines.append(line)
    return "\n".join(cleaned_lines)

# Função para ler o arquivo VCF e extrair os contatos
def read_vcf(file):
    contacts = []
    file_content = file.read()
    
    # Detectar a codificação do arquivo
    encoding = chardet.detect(file_content)['encoding']
    file_content = file_content.decode(encoding)
    
    # Limpar o conteúdo do arquivo VCF
    cleaned_content = clean_vcf_content(file_content)
    
    try:
        for vcard in vobject.readComponents(StringIO(cleaned_content)):
            phone_number = ''
            if hasattr(vcard, 'tel'):
                # Extraímos todos os números de telefone e escolhemos o primeiro encontrado
                for tel in vcard.tel_list:
                    tel_type = getattr(tel, 'type', '')
                    if tel_type in ['CELL', 'X-Celular', 'HOME', 'WORK']:
                        phone_number = tel.value
                        break
                if not phone_number:
                    phone_number = vcard.tel_list[0].value if vcard.tel_list else 'Número não disponível'
            
            ddd, ddi, telefone = split_phone_number(phone_number)

            contact = {
                'Nome': getattr(vcard, 'fn', None).value if hasattr(vcard, 'fn') else '',
                'Email': getattr(vcard, 'email', None).value if hasattr(vcard, 'email') else '',
                'Telefone': telefone,
                'DDI': ddi,
                'DDD': ddd,
                'Data Aniversário': '',  # Adicionar valor conforme necessário
                'Empresa': ''  # Adicionar valor conforme necessário
            }
            contacts.append(contact)
    except vobject.base.ParseError as e:
        print(f"Erro ao processar o arquivo VCF: {e}")
    
    return contacts

@app.route('/', methods=['GET'])
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return 'No file part', 400
    file = request.files['file']
    if file.filename == '':
        return 'No selected file', 400
    if file and file.filename.endswith('.vcf'):
        try:
            contatos = read_vcf(file)
            df = pd.DataFrame(contatos)
            df = df[['Nome', 'DDI', 'DDD', 'Telefone', 'Email', 'Data Aniversário', 'Empresa']]

            output = BytesIO()
            df.to_excel(output, index=False)
            output.seek(0)

            return send_file(
                output,
                as_attachment=True,
                download_name='contatos.xlsx',
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
        except Exception as e:
            return f'An error occurred: {e}', 500
    return 'Invalid file type', 400

if __name__ == '_main_':
    app.run(debug=True)