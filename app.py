import openpyxl
from PIL import Image, ImageDraw, ImageFont

workbook_alunos = openpyxl.load_workbook('planilha_alunos.xlsx')
sheet_alunos = workbook_alunos['Sheet1']

for indice, linha in enumerate(sheet_alunos.iter_rows( max_row= 20)):
    nome_curso = linha[0].value  # Nome do Curso
    nome_participante =linha[1].value # Nome do Participante
    tipo_participacao = linha[2].value # Tipo de Participação
    carga_horaria = linha[5].value # Carga Horária

    data_inicio = linha[3].value # Data de início
    data_final = linha[4].value # Data de Fim

    data_emissao = linha[6].value # Data de Emissão

    fonte_nome = ImageFont.truetype('./tahomabd.ttf', 90)
    fonte_geral = ImageFont.truetype('./tahoma.ttf', 80)
    fonte_data = ImageFont.truetype('./tahoma.ttf', 55)

    image = Image.open('./certificado_padrao.jpg')
    desenhar = ImageDraw.Draw(image)

    desenhar.text((1020,827), str(nome_participante), fill='black', font=fonte_nome)
    desenhar.text((1060,950), str(nome_curso), fill='black', font=fonte_geral)
    desenhar.text((1435,1065), str(tipo_participacao), fill='black', font=fonte_geral)
    desenhar.text((1480,1182), str(carga_horaria), fill='black', font=fonte_geral)

    desenhar.text((750,1770), str(data_inicio), fill='blue', font=fonte_data)
    desenhar.text((750,1930), str(data_final), fill='blue', font=fonte_data)

    desenhar.text((2230,1930), str(nome_participante), fill='blue', font=fonte_data)

    image.save(f'./{indice} {nome_participante}certificado.png' )    