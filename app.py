""" 
Preencher os dados mutáveis com os dados da planilha:
Nome do curso, nome do participante, tipo de participação,
data de inicio, data final, carga horaria, data de emissão

Transferir para a imagem do certificado
"""

# Pegar os dados da planilha
import openpyxl
from PIL import Image, ImageDraw, ImageFont

# Abrindo a planilha
workbook_alunos = openpyxl.load_workbook('planilha_alunos.xlsx')
sheet_alunos = workbook_alunos['Sheet1']

#Passar por cada linha a partir de uma linha especifica
for indice,linha in enumerate(sheet_alunos.iter_rows(min_row=2)):
    # cada celula que contém o que precisamos
    nome_curso = linha[0].value 
    nome_participante =  linha[1].value
    tipo_participacao = linha[2].value
    data_inicio = linha[3].value
    data_final = linha[4].value
    carga_horaria = linha[5].value
    data_emissao = linha[6].value

    # Transferir para a imagem do certificado

    #Definindo a fonte a ser usada
    font_nome = ImageFont.truetype('./tahomabd.ttf', 90)
    font_geral = ImageFont.truetype('./tahoma.ttf', 80)
    font_datas = ImageFont.truetype('./tahoma.ttf', 55)

    image = Image.open('certificado_padrao.jpg')
    desenhar = ImageDraw.Draw(image)

    desenhar.text((1020,827), nome_participante, fill='black', font=font_nome)
    desenhar.text((1060,950), nome_curso, fill='black', font=font_geral)
    desenhar.text((1435, 1065), tipo_participacao, fill='black', font=font_geral)
    desenhar.text((1480,1182), str(carga_horaria), fill='black', font=font_geral)
    
    desenhar.text((750, 1770), data_inicio, fill='black', font=font_datas)
    desenhar.text((750, 1930), data_final, fill='black', font=font_datas)
    desenhar.text((2220, 1930), data_emissao, fill='black', font=font_datas)

    image.save(f'./certificados/{indice} {nome_participante} certificado.png')
