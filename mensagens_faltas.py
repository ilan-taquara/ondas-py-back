from flask import Flask, request, jsonify
import urllib.parse
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from flask_cors import CORS

app = Flask(__name__)
CORS(app)


@app.route('/mensagens_faltas', methods=['POST'])
def enviar_mensagem():
    df = pd.read_excel(r'C:\Users\gabri\OneDrive\Área de Trabalho\ondas-py-back\CHAMADACADASTRO3.xlsx')
    df = df[df['Nome'] != '']
    colunas_desejadas = [1, 2, 3, 4, 'Nome', 'Contato']
    df = df[colunas_desejadas]
    data = request.get_json()
    aula_faltada = data.get('aulaFaltada')

    if aula_faltada is None:
        return jsonify({"erro": "Nenhuma aula foi especificada."}), 400

    aula_col_name = int(aula_faltada)

    if aula_col_name not in df.columns:
        return jsonify({"erro": f"A coluna '{aula_col_name}' não existe no DataFrame."}), 400

    navegador = webdriver.Chrome()

    for linha in df.index:
        print("Processando linha:", linha)
        if pd.notnull(df.loc[linha,"Nome"]):
            nome_completo = df.loc[linha, "Nome"]
            nome = nome_completo.split()[0] if pd.notnull(nome_completo) else "Aluno"
            print("Nome:", nome)

            numero = df.loc[linha, "Contato"]
            if pd.isna(numero):
                continue

            numero = int(numero)
            print("Número de contato:", numero)

            mensagem = None
            if aula_faltada != 4:
                if pd.isna(df.loc[linha, aula_col_name]):
                    mensagem = f"Olá {nome}, lembrando que tem a aula {aula_faltada} essa semana!"
            else:
                if pd.isna(df.loc[linha, aula_col_name]) and pd.notnull(df.loc[linha,3]):
                    mensagem = f"Olá {nome}, lembrando que tem a aula {aula_faltada} essa semana!"
                elif pd.isna(df.loc[linha,3]):
                    mensagem = f"Olá {nome}, vi que você faltou a aula 3. Para assistir a aula 4, é necessário estar presente na aula 3. Então não precisa vir esse fim de semana. Obrigado!"
            
            if mensagem:
                print("Enviando mensagem:", mensagem)
                texto = urllib.parse.quote(mensagem)
                link = f"https://web.whatsapp.com/send?phone={numero}&text={texto}"

                navegador.get(link)
                while len(navegador.find_elements(By.XPATH, '//*[@id="main"]/div[3]/div/div[2]/div[3]')) < 1:
                    time.sleep(1)
                navegador.find_element(By.XPATH, '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[2]/button/span').click()
                time.sleep(3)
                print(f"Mensagem enviada para {nome}")

    navegador.quit()
    return jsonify({"status": "Mensagens enviadas com sucesso"}), 200

if __name__ == '__main__':
    app.run(debug=True)
