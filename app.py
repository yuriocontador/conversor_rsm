from flask import Flask, request, redirect, render_template, send_file
import pandas as pd
from datetime import datetime
import os

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'outputs'

# Garantindo que as pastas existam
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        # Verifica se o arquivo foi enviado
        if 'file' not in request.files:
            return 'Nenhum arquivo enviado!'
        file = request.files['file']
        if file.filename == '':
            return 'Nenhum arquivo selecionado!'
        
        # Salva o arquivo enviado
        file_path = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(file_path)
        
        # Converte para OFX
        ofx_file = convert_to_ofx(file_path)
        
        # Retorna o arquivo OFX gerado para download
        return send_file(ofx_file, as_attachment=True)
    
    return render_template('index.html')

def convert_to_ofx(file_path):
    # Lê o arquivo Excel
    df = pd.read_excel(file_path)
    
    # Ajuste os nomes das colunas conforme o layout do Excel
    df.columns = [col.strip() for col in df.columns]
    df = df[['Data', 'Historico', 'Valor']]
    df['Data'] = pd.to_datetime(df['Data'], dayfirst=True)
    
    # Estrutura do OFX
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
        date_ofx = row['Data'].strftime('%Y%m%d')
        amount = f"{float(row['Valor']):.2f}"
        fitid = f"{date_ofx}{fitid_counter}"
        fitid_counter += 1
        memo = str(row['Historico']).replace('&', 'e')
        ofx_output.extend([
            "<STMTTRN>",
            f"<TRNTYPE>{'DEBIT' if float(amount) < 0 else 'CREDIT'}",
            f"<DTPOSTED>{date_ofx}000000",
            f"<TRNAMT>{amount}",
            f"<FITID>{fitid}",
            f"<CHECKNUM>{fitid}",
            f"<MEMO>{memo}",
            "</STMTTRN>"
        ])

    ofx_output.extend([
        "</BANKTRANLIST>",
        "<LEDGERBAL>",
        f"<BALAMT>{df['Valor'].sum():.2f}",
        f"<DTASOF>{datetime.now().strftime('%Y%m%d%H%M%S')}",
        "</LEDGERBAL>",
        "</STMTRS>",
        "</STMTTRNRS>",
        "</BANKMSGSRSV1>",
        "</OFX>"
    ])
    
    # Salva o arquivo OFX
    output_file = os.path.join(OUTPUT_FOLDER, 'output.ofx')
    with open(output_file, 'w', encoding='utf-8') as file:
        file.write("\n".join(ofx_output))
    
    return output_file

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
