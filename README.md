# 🗂️ Projeto de Automação de Certificados 
O Repositório em questão apresenta um projeto de automação de preenchimento de certificados pegando os dados de uma planilha.

Nesse repositório estão contidos:
- Código do projeto.
- Planilha a ser utilizada como exemplo.
- Imagem do certificado que deverá ser utilizado.

Esse projeto tem o intuito de aumentar meu aprendizado na linguagem Python e foi retirado do vídeo [Projeto Automação de Certificados](https://youtu.be/VwYqakOB4ow?si=PnO6CPQEcVJDNkVz)

## ✏️Explicação do Código
Esse projeto tem como objetivo pegar dados de uma planilha do excel e substituir em uma imagem de certificado padrão.

Para fazer isso é possível utilizar algumas bibliotecas do python:
- Pillow: Sobrepor texto na imagem
- openpyxl: Ler dados de uma planilha

Inserindo no terminal o seguinte comando para se instalar o pillow e o openpyxl
```
pip install pillow openpyxl
```
### Passo a passo:
1. Importar as bibliotecas de leitura de planilhas e de imagens:
```
import openpyxl
from PIL import Image, ImageDraw, ImageFont
```
2. Adicionar a planilha à ser manipulada em uma variável e pegar o Sheet dela:
```
workbook_alunos = openpyxl.load_workbook('planilha_alunos.xlsx')
sheet_alunos = workbook_alunos['Sheet1']
```
3. Fazer um laço de repetição que pegue os dados das colunas da planilha e adicionem em uma variável
```
for indice,linha in enumerate(sheet_alunos.iter_rows(min_row=2)):
    # cada celula que contém o que precisamos
    nome_curso = linha[0].value 
    nome_participante =  linha[1].value
    tipo_participacao = linha[2].value
    data_inicio = linha[3].value
    data_final = linha[4].value
    carga_horaria = linha[5].value
    data_emissao = linha[6].value
```
4. Dentro desse mesmo laço, criar uma variável que instancie a imagem do certificado utilizando o Image do Pillow. Além disso, definir as fontes a serem utilizadas e o Draw da imagem (que desenhará o texto vindo da planilha na imagem):
```
font_nome = ImageFont.truetype('./tahomabd.ttf', 90)
font_geral = ImageFont.truetype('./tahoma.ttf', 80)
font_datas = ImageFont.truetype('./tahoma.ttf', 55)

image = Image.open('certificado_padrao.jpg')
desenhar = ImageDraw.Draw(image)
```
5. Após isso, desenhar propriamente os valores da planilha e salvar em um local que você deseje, nesse caso, criei um diretório de certificados que armazena todos os certificados gerados.
```
desenhar.text((1020,827), nome_participante, fill='black', font=font_nome)
desenhar.text((1060,950), nome_curso, fill='black', font=font_geral)
desenhar.text((1435, 1065), tipo_participacao, fill='black', font=font_geral)
desenhar.text((1480,1182), str(carga_horaria), fill='black', font=font_geral)
    
desenhar.text((750, 1770), data_inicio, fill='black', font=font_datas)
desenhar.text((750, 1930), data_final, fill='black', font=font_datas)
desenhar.text((2220, 1930), data_emissao, fill='black', font=font_datas)

image.save(f'./certificados/{indice} {nome_participante} certificado.png')
```

É isso pessoal! Espero que tenham gostado do resumo e qualquer dúvida estou à disposição no meu [LinkedIn](https://www.linkedin.com/in/yasmin-ramos-mello-gon%C3%A7alves-1b033b22a/)
