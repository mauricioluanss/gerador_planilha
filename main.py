import requests  
import json
import datetime
import pandas as pd
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill, Color
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image
import os
from pathlib import Path
from dotenv import load_dotenv
import time
import re

print("##############################################")
print("###   GERADOR DE FICHAS DE IMPLANTAÇÃO     ###")
print("###   Autor: Mauricio Luan /2025           ###")
print("##############################################\n")

while True:
    def request():
        load_dotenv()
        API_TOKEN = os.getenv("API_TOKEN")
        API_URL = os.getenv("API_URL")
        API_TICKET_URL = os.getenv("API_TICKET_URL")
 
        id = input("\nID do cliente: ")  
        ticket_id = input("ID do chamado: ")

        api = f"{API_URL}{id}" 
        apiticket = f"{API_TICKET_URL}{ticket_id}"

        headers = {'Authorization': f'Bearer {API_TOKEN}'}

        resposta = requests.get(api, headers=headers) 
        resposta_ticket = requests.get(apiticket, headers=headers)

        return resposta, resposta_ticket
    resposta, resposta_ticket = request()


    arquivo = {}
    while True:
        if resposta.status_code == 200 and resposta_ticket.status_code == 200:
            arquivo["resposta"] = resposta.json()
            arquivo["resposta_ticket"] = resposta_ticket.json()
            break
        else:
            print(f"Retorno da requisição do cliente: {resposta.status_code} - {resposta.reason}")
            print(f"Retorno da requisição do chamado: {resposta_ticket.status_code} - {resposta_ticket.reason}")
            resposta, resposta_ticket = request()


    with open('arquivo.json', 'w', encoding='utf-8') as file:
        json.dump(arquivo, file, ensure_ascii=False, indent=4)

    with open('arquivo.json', 'r', encoding='utf-8') as file:
        dicionario = json.load(file)

    
    dados_filtrados = {}
    try:
        dados_filtrados['Chamado'] = dicionario['resposta_ticket']['data']['protocol']
        dados_filtrados['Razão Social'] = dicionario['resposta']['data'][0]['name']
        nome_fantasia = dicionario['resposta']['data'][0]['custom_fields'][0]['value']

        conta_empresa_loja = dicionario['resposta']['data'][0]['internal_id']
        conta, empresa, loja = conta_empresa_loja.split('-')
        dados_filtrados['Conta'] = f"{conta} - {dados_filtrados['Razão Social']}"
        dados_filtrados['Empresa'] = f"{empresa} - {dados_filtrados['Razão Social']}"
        dados_filtrados['Loja'] = f"{loja} - {nome_fantasia}"
        partes = conta_empresa_loja.split('-')
        conta_editada = partes[1] + partes[2]
        dados_filtrados['pdv'] = dicionario['resposta_ticket']['data']['subject']
        padrao = r"Terminais \[(\d+)\]"
        resultado = re.search(padrao, dados_filtrados['pdv'])

        if resultado:
            numero_terminais = int(resultado.group(1))
            partes = conta_empresa_loja.split('-')
            conta_editada = partes[1] + partes[2]
            sequencias = [f"{conta_editada}{i:02}" for i in range(1, numero_terminais + 1)]

        captura = dicionario['resposta']['data'][0]['custom_fields']
        
        dados_a_capturar = [
            "Nome Fantasia","CNPJ","Endereco","Numero","Bairro","Cidade","COMERCIAL - Contato","COMERCIAL - Telefone","COMERCIAL - E-mail"
        ]

        for campo in dados_a_capturar:
            for item in captura:
                if item['name'] == campo:
                    dados_filtrados[campo] = item['value']

    except KeyError as e:
        print(f"Erro ao acessar uma chave do JSON: {e}")
    except IndexError as e:
        print(f"Erro ao acessar um índice: {e}")


    data = datetime.date.today()
    dados_filtrados_reorganizado = {
        "Chamado": str(dados_filtrados["Chamado"]) + " - " + str(data.strftime("%d/%m/%Y")),
        "Nome Fantasia": dados_filtrados["Nome Fantasia"],
        "Razão Social": dados_filtrados["Razão Social"],
        "CNPJ": dados_filtrados["CNPJ"],
        "Endereco": dados_filtrados["Endereco"] + ", " + dados_filtrados["Numero"],
        "Bairro": dados_filtrados["Bairro"],
        "Cidade": dados_filtrados["Cidade"],
        "Contato": dados_filtrados["COMERCIAL - Contato"],
        "Telefone": dados_filtrados["COMERCIAL - Telefone"],
        "E-mail": dados_filtrados["COMERCIAL - E-mail"],
        "Conta": dados_filtrados["Conta"],
        "Empresa": dados_filtrados["Empresa"],
        "Loja": dados_filtrados["Loja"],
        "Token Payer": " / ".join(sequencias)
    }


    with open('teste', 'w', encoding='utf-8') as file:
        json.dump(dados_filtrados_reorganizado, file, ensure_ascii=False, indent=4) 
        

    def gerar_planilha_estilizada(dados, arquivo_excel, caminho_imagem):
        wb = Workbook()
        ws = wb.active
        ws.title = "Sitef"

        fonte_negrito = Font(bold=True)
        fonte_branca = Font(color="FFFFFF", bold=True)
        fonte_cinza_claro = Font(color="808080", bold=True)
        alinhamento_central = Alignment(horizontal='center', vertical='center', wrap_text=True)
        alinhamento_esquerda = Alignment(horizontal='left', vertical='center', wrap_text=True)
        borda_fina = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        preenchimento_cabecalho = PatternFill(start_color="171717", end_color="171717", fill_type="solid")
        preenchimento_secoes = PatternFill(start_color="CCCCCC", end_color="CCCCCC", fill_type="solid")

        ws.column_dimensions['A'].width = 17
        ws.column_dimensions['B'].width = 50

        ws["A1"].value = "PAYER"
        ws["A1"].font = fonte_branca
        ws["A1"].alignment = alinhamento_central
        ws["A1"].fill = preenchimento_cabecalho
        ws["A1"].border = borda_fina

        img = Image(caminho_imagem)
        img.width = 80  
        img.height = 17  
        ws.add_image(img, 'A1') 

        ws["B1"].value = "FICHA DE IMPLANTAÇÃO"
        ws["B1"].font = fonte_branca
        ws["B1"].alignment = alinhamento_central
        ws["B1"].fill = preenchimento_cabecalho
        ws["B1"].border = borda_fina

        linha_atual = 2
        for chave, valor in dados.items():
            if chave == "Conta":
                ws.merge_cells(f"A{linha_atual}:B{linha_atual}")
                ws[f"A{linha_atual}"].value = "DADOS PAYER"
                ws[f"A{linha_atual}"].font = fonte_negrito
                ws[f"A{linha_atual}"].alignment = alinhamento_central
                ws[f"A{linha_atual}"].fill = preenchimento_secoes
                ws[f"A{linha_atual}"].border = borda_fina
                linha_atual += 1

            ws[f"A{linha_atual}"].value = chave.upper()
            ws[f"A{linha_atual}"].font = fonte_cinza_claro
            ws[f"A{linha_atual}"].border = borda_fina
            ws[f"A{linha_atual}"].alignment = alinhamento_central

            
            ws[f"B{linha_atual}"].value = valor
            ws[f"B{linha_atual}"].border = borda_fina
            ws[f"B{linha_atual}"].alignment = alinhamento_esquerda
            ws.column_dimensions['B'].width = 49
            
            if linha_atual in [13, 14, 15, 16]:
                ws[f"B{linha_atual}"].alignment = alinhamento_central
                ws.column_dimensions['B'].width = 48

            linha_atual += 1

        wb.save(arquivo_excel)
        print(f"\nPlanilha salva em: {arquivo_excel}")


    caminho_base = "G:\\Drives compartilhados\\FICHAS DE IMPLANTACAO"

    primeira_letra = dados_filtrados["Razão Social"][0].upper()
    subpasta_letra = os.path.join(caminho_base, primeira_letra)


    tipo_mensagem_letra = "já existia" if os.path.exists(subpasta_letra) else "foi criada"
    os.makedirs(subpasta_letra, exist_ok=True)
    print(f"\nPasta da letra '{primeira_letra}'{tipo_mensagem_letra}: {subpasta_letra}")


    pasta_razao_social = os.path.join(subpasta_letra, dados_filtrados["Razão Social"])
    tipo_mensagem_razao = "já existia" if os.path.exists(pasta_razao_social) else "foi criada"
    os.makedirs(pasta_razao_social, exist_ok=True)
    print(f"Pasta da Razão Social '{dados_filtrados['Razão Social']}' {tipo_mensagem_razao}: {pasta_razao_social}")

    
    arquivo_excel = os.path.join(pasta_razao_social, f"{dados_filtrados['Loja']}.xlsx")

    
    url_imagem = "https://i.postimg.cc/CMXH4QSk/payer.png"
    resposta = requests.get(url_imagem)
    caminho_imagem = 'payer.png'
    with open(caminho_imagem, 'wb') as f:
        f.write(resposta.content)

    gerar_planilha_estilizada(dados_filtrados_reorganizado, arquivo_excel, caminho_imagem)

    time.sleep(3)
    os.remove('teste') 
    os.remove('arquivo.json')  
    os.remove(caminho_imagem)