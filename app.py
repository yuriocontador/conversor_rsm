import os
import pandas as pd
from flask import Flask, request, redirect, url_for, send_file, render_template_string
from datetime import datetime

UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'outputs'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

HTML_TEMPLATE = '''
<!doctype html>
<html>
    <head>
        <title>Conversor Excel para OFX</title>
    </head>
    <body>
        <h1>Conversor Excel para OFX - Money 97/2000</h1>
        <form action="/upload" method="post" enctype="multipart/form-data">
            <input type="file" name="file">
            <button type="submit">Converter para OFX</button>
        </form>
        {% if ofx_file %}
            <h2>Arquivo convertido com sucesso!</h2>
            <a href="{{ ofx_file }}">Baixar OFX</a>
        {% endif %}
    </body>
</html>
'''

@app.route('/', methods=['GET'])
def index():
    return render_template_string(HTML_TEMPLATE)

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return redirect(url_for('index'))
    file = request.files['file']
    if file.filename == '':
        return redirect(url_for('index'))
    file_path = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
    file.save(file_path)
    ofx_file = convert_to_ofx(file_path)
    return render_template_string(HTML_TEMPLATE, ofx_file=ofx_file)


def convert_to_ofx(file_path):
    df = pd.read_excel(file_path)
    df.columns = ['Data', 'Historico', 'Documento', 'Valor']
    df['Data'] = pd.to_datetime(df['Data'], errors='coerce', dayfirst=True)
    df = df[df['Data'].notna()]
    df['Valor'] = pd.to_numeric(df['Valor'], errors='coerce').fillna(0)
    ofx_output = [
        "OFXHEADER:100",
        "DATA:OFXSGML",
        "VERSION:102",
        "SECURITY:NONE",
        "ENCODING:USASCII",
        "CHARSET:1252",
        "COMPRESSION:NONE",
        "OLDFILEUID:NONE",
        "NEWFILEUID:NONE",
        "<OFX>",
        "<SIGNONMSGSRSV1>",
        "<SONRS>",
        "<STATUS><CODE>0<SEVERITY>INFO</STATUS>",
        f"<DTSERVER>{datetime.now().strftime('%Y%m%d%H%M%S')}",
        "<LANGUAGE>POR",
        "<FI><ORG>SANTANDER<FID>033</FI>",
        "</SONRS>",
        "</SIGNONMSGSRSV1>",
        "<BANKMSGSRSV1>",
        "<STMTTRNRS>",
        "<TRNUID>1",
        "<STATUS><CODE>0<SEVERITY>INFO</STATUS>",
        "<STMTRS>",
        "<CURDEF>BRL",
        "<BANKACCTFROM><BANKID>033<ACCTID>4307130042633<ACCTTYPE>CHECKING</BANKACCTFROM>",
        "<BANKTRANLIST>"
    ]
    fitid_counter = 1
    for _, row in df.iterrows():
        date_ofx = row['Data'].strftime('%Y%m%d%H%M%S')
        amount = f"{float(row['Valor']):.2f}".replace(',', '.')
        fitid = f"{date_ofx}{fitid_counter}"
        fitid_counter += 1
        memo = str(row['Historico']).replace('&', 'e')
        ofx_output.extend([
            "<STMTTRN>",
            f"<TRNTYPE>{'DEBIT' if float(amount) < 0 else 'CREDIT'}",
            f"<DTPOSTED>{date_ofx}",
            f"<TRNAMT>{amount}",
            f"<FITID>{fitid}",
            f"<CHECKNUM>{fitid}",
            f"<MEMO>{memo}",
            "</STMTTRN>"
        ])
    ofx_output.extend([
        "</BANKTRANLIST>",
        "</STMTRS>",
        "</STMTTRNRS>",
        "</BANKMSGSRSV1>",
        "</OFX>"
    ])
    output_file = os.path.join(OUTPUT_FOLDER, 'extrato_convertido.ofx')
    with open(output_file, 'w', encoding='utf-8') as file:
        file.write("\n".join(ofx_output))
    return url_for('static', filename=f'outputs/extrato_convertido.ofx')

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
