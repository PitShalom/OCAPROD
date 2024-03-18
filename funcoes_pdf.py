import os
from PyPDF2 import PdfReader, PdfWriter,PdfFileReader, PdfFileWriter
from reportlab.pdfgen import canvas
import io
from io import BytesIO
from reportlab.lib.pagesizes import letter
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from datetime import datetime,timedelta
import fitz  # PyMuPDF

def registrar_fontes():
    # Adicione o registro da fonte Arial-Bold
    pdfmetrics.registerFont(TTFont('Arial-Bold', 'arialbd.ttf'))

    # Adicione o registro da fonte Arial
    if 'Arial' not in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFont(TTFont('Arial', 'arial.ttf'))

    # Adicione o registro da fonte IBMPlexSans-Text
    font_path = os.path.join(os.path.dirname(__file__), r'C:\progOca\fonts')
    font_name_text = 'IBMPlexSans-Text'

    if font_name_text not in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFont(TTFont(font_name_text, os.path.join(font_path, 'IBMPlexSans-Text.ttf')))

    # Adicione o registro da fonte IBMPlexSans-Bold
    font_name_bold = 'IBMPlexSans-Bold'

    if font_name_bold not in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFont(TTFont(font_name_bold, os.path.join(font_path, 'IBMPlexSans-Bold.ttf')))


def preencher_nr35(nome, cpf, funcao, dataNR35, Hab_SupInt, nome_TST, n_registroTST, modelo_path, output_path, incluir_funcao=True):
    registrar_fontes()
    pdfmetrics.registerFont(TTFont('Arial-Bold', 'arialbd.ttf'))

    if 'Arial' not in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFont(TTFont('Arial', 'arial.ttf'))

    dataNR35_formatada = formatar_data(dataNR35, formato='curta')

    with open(modelo_path, "rb") as model_file:
        existing_pdf = PdfReader(io.BytesIO(model_file.read()))

    output = PdfWriter()

    for page_num in range(len(existing_pdf.pages)):
        existing_page = existing_pdf.pages[page_num]

        if page_num == 0:  # Somente preencher a primeira página
            packet = io.BytesIO()
            can = canvas.Canvas(packet, pagesize=(800, 1200))

            can.setFont("IBMPlexSans-Bold", 32)
            nome_width = can.stringWidth(nome, "IBMPlexSans-Bold", 32)
            x_position = (800 - nome_width) / 2
            can.drawString(x_position, 380, f'{nome}')

            can.setFont("IBMPlexSans-Bold", 13)
            can.drawString(246.5, 138, f'{cpf}')
#--------------------------------------------------------------
            can.setFont("IBMPlexSans-Bold", 13) 
            can.drawString(455, 123, f'{nome_TST}')

            nome_TST = "Téc. De Seg. Do Trabalho"
            # Use o valor definido na função drawString
            can.setFont("IBMPlexSans-Bold", 13)
            can.drawString(488, 107, f'{nome_TST}')
#--------------------------------------------------------------
            
            can.setFont("IBMPlexSans-Bold", 13) 
            can.drawString(485, 93, f'{n_registroTST}')  # (pros lados, pra cima-baixo,)

            can.setFont("IBMPlexSans-Bold", 18)
            can.drawString(440, 300, f'{dataNR35_formatada}')

            if incluir_funcao:
                can.setFont("IBMPlexSans-Bold", 15)
                can.drawString(326, 450, f'{funcao}')

            can.save()

            packet.seek(0)
            new_pdf_data = packet.getvalue()

            new_page = PdfReader(io.BytesIO(new_pdf_data)).pages[0]
            existing_page.merge_page(new_page)

        output.add_page(existing_page)

    base_name, extension = os.path.splitext(os.path.basename(output_path))
    output_folder = os.path.join(r'C:\pdfBaixados', nome)

    # Criar diretório se não existir
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    unique_output_path = os.path.join(output_folder, f'{base_name}{extension}')

    with open(unique_output_path, "wb") as output_file:
        output.write(output_file)

    print(f'NR35 para {nome} preenchida e salva em {unique_output_path}.')
    return unique_output_path

def preencher_nr18(nome, cpf, funcao, dataNR18, Hab_SupInt, nome_TST, n_registroTST, modelo_path, output_path, incluir_funcao=True):
    registrar_fontes()
    pdfmetrics.registerFont(TTFont('Arial-Bold', 'arialbd.ttf'))

    if 'Arial' not in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFont(TTFont('Arial', 'arial.ttf'))

    dataNR18_formatada = formatar_data(dataNR18, formato='curta')

    with open(modelo_path, "rb") as model_file:
        existing_pdf = PdfReader(io.BytesIO(model_file.read()))

    output = PdfWriter()

    for page_num in range(len(existing_pdf.pages)):
        existing_page = existing_pdf.pages[page_num]

        if page_num == 0:  # Somente preencher a primeira página
            packet = io.BytesIO()
            can = canvas.Canvas(packet, pagesize=(800, 1200))

            can.setFont("IBMPlexSans-Bold", 32)
            nome_width = can.stringWidth(nome, "IBMPlexSans-Bold", 32)
            x_position = (800 - nome_width) / 2
            can.drawString(x_position, 380, f'{nome}')

            can.setFont("IBMPlexSans-Bold", 13)
            can.drawString(246.5, 138, f'{cpf}')
#--------------------------------------------------------------
            can.setFont("IBMPlexSans-Bold", 13) 
            can.drawString(455, 123, f'{Hab_SupInt}')

            nome_TST = "Téc. De Seg. Do Trabalho" 
            # Use o valor definido na função drawString
            can.setFont("IBMPlexSans-Bold", 13)
            can.drawString(488, 107, f'{nome_TST}')
#--------------------------------------------------------------
            
            can.setFont("IBMPlexSans-Bold", 13) 
            can.drawString(485, 93, f'{n_registroTST}')  # (pros lados, pra cima-baixo,)

            can.setFont("IBMPlexSans-Bold", 18)
            can.drawString(440, 300, f'{dataNR18_formatada}')

            if incluir_funcao:
                can.setFont("IBMPlexSans-Bold", 15)
                can.drawString(326, 450, f'{funcao}')

            can.save()

            packet.seek(0)
            new_pdf_data = packet.getvalue()

            new_page = PdfReader(io.BytesIO(new_pdf_data)).pages[0]
            existing_page.merge_page(new_page)

        output.add_page(existing_page)

    base_name, extension = os.path.splitext(os.path.basename(output_path))
    output_folder = os.path.join(r'C:\pdfBaixados', nome)

    # Criar diretório se não existir
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    unique_output_path = os.path.join(output_folder, f'{base_name}{extension}')

    with open(unique_output_path, "wb") as output_file:
        output.write(output_file)

    print(f'NR18 para {nome} preenchida e salva em {unique_output_path}.')
    return unique_output_path

def preencher_nr01(nome, cpf, funcao, dataNR01, Hab_SupInt, nome_TST, n_registroTST, modelo_path, output_path, incluir_funcao=True):
    registrar_fontes()
    pdfmetrics.registerFont(TTFont('Arial-Bold', 'arialbd.ttf'))

    if 'Arial' not in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFont(TTFont('Arial', 'arial.ttf'))

    dataNR01_formatada = formatar_data(dataNR01, formato='curta')

    with open(modelo_path, "rb") as model_file:
        existing_pdf = PdfReader(io.BytesIO(model_file.read()))

    output = PdfWriter()

    for page_num in range(len(existing_pdf.pages)):
        existing_page = existing_pdf.pages[page_num]

        if page_num == 0:  # Somente preencher a primeira página
            packet = io.BytesIO()
            can = canvas.Canvas(packet, pagesize=(800, 1200))

            can.setFont("IBMPlexSans-Bold", 32)
            nome_width = can.stringWidth(nome, "IBMPlexSans-Bold", 32)
            x_position = (800 - nome_width) / 2
            can.drawString(x_position, 380, f'{nome}')

            can.setFont("IBMPlexSans-Bold", 13)
            can.drawString(246.5, 138, f'{cpf}')
#--------------------------------------------------------------
            can.setFont("IBMPlexSans-Bold", 13) 
            can.drawString(455, 123, f'{Hab_SupInt}')

            nome_TST = "Téc. De Seg. Do Trabalho" 
            # Use o valor definido na função drawString
            can.setFont("IBMPlexSans-Bold", 13)
            can.drawString(488, 107, f'{nome_TST}')
#--------------------------------------------------------------
            
            can.setFont("IBMPlexSans-Bold", 13) 
            can.drawString(485, 93, f'{n_registroTST}')  # (pros lados, pra cima-baixo,)

            can.setFont("IBMPlexSans-Bold", 18)
            can.drawString(440, 300, f'{dataNR01_formatada}')

            if incluir_funcao:
                can.setFont("IBMPlexSans-Bold", 15)
                can.drawString(326, 450, f'{funcao}')

            can.save()

            packet.seek(0)
            new_pdf_data = packet.getvalue()

            new_page = PdfReader(io.BytesIO(new_pdf_data)).pages[0]
            existing_page.merge_page(new_page)

        output.add_page(existing_page)

    base_name, extension = os.path.splitext(os.path.basename(output_path))
    output_folder = os.path.join(r'C:\pdfBaixados', nome)

    # Criar diretório se não existir
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    unique_output_path = os.path.join(output_folder, f'{base_name}{extension}')

    with open(unique_output_path, "wb") as output_file:
        output.write(output_file)

    print(f'NR01 para {nome} preenchida e salva em {unique_output_path}.')
    return unique_output_path

def preencher_nr05(nome, cpf, funcao, dataNR05, Hab_SupInt, nome_TST, n_registroTST, modelo_path, output_path, incluir_funcao=True):
    registrar_fontes()
    pdfmetrics.registerFont(TTFont('Arial-Bold', 'arialbd.ttf'))

    if 'Arial' not in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFont(TTFont('Arial', 'arial.ttf'))

    dataNR05_formatada = formatar_data(dataNR05, formato='curta')

    with open(modelo_path, "rb") as model_file:
        existing_pdf = PdfReader(io.BytesIO(model_file.read()))

    output = PdfWriter()

    for page_num in range(len(existing_pdf.pages)):
        existing_page = existing_pdf.pages[page_num]

        if page_num == 0:  # Somente preencher a primeira página
            packet = io.BytesIO()
            can = canvas.Canvas(packet, pagesize=(800, 1200))

            can.setFont("IBMPlexSans-Bold", 32)
            nome_width = can.stringWidth(nome, "IBMPlexSans-Bold", 32)
            x_position = (800 - nome_width) / 2
            can.drawString(x_position, 380, f'{nome}')

            can.setFont("IBMPlexSans-Bold", 13)
            can.drawString(246.5, 138, f'{cpf}')
#--------------------------------------------------------------
            can.setFont("IBMPlexSans-Bold", 13) 
            can.drawString(455, 123, f'{Hab_SupInt}')

            nome_TST = "Téc. De Seg. Do Trabalho"
            # Use o valor definido na função drawString
            can.setFont("IBMPlexSans-Bold", 13)
            can.drawString(488, 107, f'{nome_TST}')
#--------------------------------------------------------------
            
            can.setFont("IBMPlexSans-Bold", 13) 
            can.drawString(485, 93, f'{n_registroTST}')  # (pros lados, pra cima-baixo,)

            can.setFont("IBMPlexSans-Bold", 18)
            can.drawString(440, 300, f'{dataNR05_formatada}')

            if incluir_funcao:
                can.setFont("IBMPlexSans-Text", 15)
                can.drawString(326, 450, f'{funcao}')

            can.save()

            packet.seek(0)
            new_pdf_data = packet.getvalue()

            new_page = PdfReader(io.BytesIO(new_pdf_data)).pages[0]
            existing_page.merge_page(new_page)

        output.add_page(existing_page)

    base_name, extension = os.path.splitext(os.path.basename(output_path))
    output_folder = os.path.join(r'C:\pdfBaixados', nome)

    # Criar diretório se não existir
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    unique_output_path = os.path.join(output_folder, f'{base_name}{extension}')

    with open(unique_output_path, "wb") as output_file:
        output.write(output_file)

    print(f'NR05 para {nome} preenchida e salva em {unique_output_path}.')
    return unique_output_path

def preencher_nr06(nome, cpf, funcao, dataNR06, Hab_SupInt, nome_TST, n_registroTST, modelo_path, output_path, incluir_funcao=True):
    registrar_fontes()
    pdfmetrics.registerFont(TTFont('Arial-Bold', 'arialbd.ttf'))

    if 'Arial' not in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFont(TTFont('Arial', 'arial.ttf'))

    dataNR06_formatada = formatar_data(dataNR06, formato='curta')

    with open(modelo_path, "rb") as model_file:
        existing_pdf = PdfReader(io.BytesIO(model_file.read()))

    output = PdfWriter()

    for page_num in range(len(existing_pdf.pages)):
        existing_page = existing_pdf.pages[page_num]

        if page_num == 0:  
            packet = io.BytesIO()
            can = canvas.Canvas(packet, pagesize=(800, 1200))

            can.setFont("IBMPlexSans-Bold", 32)
            nome_width = can.stringWidth(nome, "IBMPlexSans-Bold", 32)
            x_position = (800 - nome_width) / 2
            can.drawString(x_position, 380, f'{nome}')

            can.setFont("IBMPlexSans-Bold", 13)
            can.drawString(246.5, 138, f'{cpf}')
#--------------------------------------------------------------
            can.setFont("IBMPlexSans-Bold", 13) 
            can.drawString(455, 123, f'{Hab_SupInt}')

            nome_TST = "Téc. De Seg. Do Trabalho"
            # Use o valor definido na função drawString
            can.setFont("IBMPlexSans-Bold", 13)
            can.drawString(488, 107, f'{nome_TST}')
#--------------------------------------------------------------
            
            can.setFont("IBMPlexSans-Bold", 13) 
            can.drawString(485, 93, f'{n_registroTST}')   # (pros lados, pra cima-baixo,)

            can.setFont("IBMPlexSans-Bold", 18)
            can.drawString(440, 300, f'{dataNR06_formatada}')

            if incluir_funcao:
                can.setFont("IBMPlexSans-Text", 15)
                can.drawString(326, 450, f'{funcao}')

            can.save()

            packet.seek(0)
            new_pdf_data = packet.getvalue()

            new_page = PdfReader(io.BytesIO(new_pdf_data)).pages[0]
            existing_page.merge_page(new_page)

        output.add_page(existing_page)

    base_name, extension = os.path.splitext(os.path.basename(output_path))
    output_folder = os.path.join(r'C:\pdfBaixados', nome)

    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    unique_output_path = os.path.join(output_folder, f'{base_name}{extension}')

    with open(unique_output_path, "wb") as output_file:
        output.write(output_file)

    print(f'NR06 para {nome} preenchida e salva em {unique_output_path}.')
    return unique_output_path

def preencher_nr10basic(nome, cpf, funcao, dataNR10_basica, Hab_SupInt, nomeTecRep, n_superInt, modelo_path, output_path, incluir_funcao=True):
    registrar_fontes()
    pdfmetrics.registerFont(TTFont('Arial-Bold', 'arialbd.ttf'))

    if 'Arial' not in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFont(TTFont('Arial', 'arial.ttf'))

    dataNR10basic_formatada = formatar_data(dataNR10_basica, formato='curta')

    with open(modelo_path, "rb") as model_file:
        existing_pdf = PdfReader(io.BytesIO(model_file.read()))

    output = PdfWriter()

    for page_num in range(len(existing_pdf.pages)):
        existing_page = existing_pdf.pages[page_num]

        if page_num == 0: 
            packet = io.BytesIO()
            can = canvas.Canvas(packet, pagesize=(800, 1200))

            can.setFont("IBMPlexSans-Bold", 32)
            nome_width = can.stringWidth(nome, "IBMPlexSans-Bold", 32)
            x_position = (800 - nome_width) / 2
            can.drawString(x_position, 380, f'{nome}')

            can.setFont("IBMPlexSans-Bold", 13)
            can.drawString(246.5, 138, f'{cpf}')
#--------------------------------------------------------------
            can.setFont("IBMPlexSans-Bold", 13) 
            can.drawString(455, 123, f'{Hab_SupInt}')

            can.setFont("IBMPlexSans-Bold", 13) 
            can.drawString(488, 107, f'{nomeTecRep}')
#--------------------------------------------------------------
            
            can.setFont("IBMPlexSans-Bold", 13) 
            can.drawString(485, 93, f'{n_superInt}')  # (pros lados, pra cima-baixo,)

            can.setFont("IBMPlexSans-Bold", 18)
            can.drawString(440, 300, f'{dataNR10basic_formatada}')

            if incluir_funcao:
                can.setFont("IBMPlexSans-Text", 15)
                can.drawString(326, 450, f'{funcao}')

            can.save()

            packet.seek(0)
            new_pdf_data = packet.getvalue()

            new_page = PdfReader(io.BytesIO(new_pdf_data)).pages[0]
            existing_page.merge_page(new_page)

        output.add_page(existing_page)

    base_name, extension = os.path.splitext(os.path.basename(output_path))
    output_folder = os.path.join(r'C:\pdfBaixados', nome)

    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    unique_output_path = os.path.join(output_folder, f'{base_name}{extension}')

    with open(unique_output_path, "wb") as output_file:
        output.write(output_file)

    print(f'NR10 básico para {nome} preenchida e salva em {unique_output_path}.')
    return unique_output_path

def preencher_nr10comp(nome, cpf, funcao, dataNR10_complementar, Hab_SupInt, nomeTecRep, n_superInt, modelo_path, output_path, incluir_funcao=True):
    registrar_fontes()
    pdfmetrics.registerFont(TTFont('Arial-Bold', 'arialbd.ttf'))

    if 'Arial' not in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFont(TTFont('Arial', 'arial.ttf'))

    dataNR10comp_formatada = formatar_data(dataNR10_complementar, formato='curta')

    with open(modelo_path, "rb") as model_file:
        existing_pdf = PdfReader(io.BytesIO(model_file.read()))

    output = PdfWriter()

    for page_num in range(len(existing_pdf.pages)):
        existing_page = existing_pdf.pages[page_num]

        if page_num == 0:  # Somente preencher a primeira página
            packet = io.BytesIO()
            can = canvas.Canvas(packet, pagesize=(800, 1200))

            can.setFont("IBMPlexSans-Bold", 32)
            nome_width = can.stringWidth(nome, "IBMPlexSans-Bold", 32)
            x_position = (800 - nome_width) / 2
            can.drawString(x_position, 380, f'{nome}')

            can.setFont("IBMPlexSans-Bold", 13)
            can.drawString(246.5, 138, f'{cpf}')
#--------------------------------------------------------------
            can.setFont("IBMPlexSans-Bold", 13) 
            can.drawString(455, 123, f'{Hab_SupInt}')

            can.setFont("IBMPlexSans-Bold", 13) 
            can.drawString(488, 107, f'{nomeTecRep}')
#--------------------------------------------------------------
            
            can.setFont("IBMPlexSans-Bold", 13) 
            can.drawString(485, 93, f'{n_superInt}')  # (pros lados, pra cima-baixo,)

            can.setFont("IBMPlexSans-Bold", 18)
            can.drawString(440, 300, f'{dataNR10comp_formatada}')

            if incluir_funcao:
                can.setFont("IBMPlexSans-Text", 15)
                can.drawString(326, 450, f'{funcao}')

            can.save()

            packet.seek(0)
            new_pdf_data = packet.getvalue()

            new_page = PdfReader(io.BytesIO(new_pdf_data)).pages[0]
            existing_page.merge_page(new_page)

        output.add_page(existing_page)

    base_name, extension = os.path.splitext(os.path.basename(output_path))
    output_folder = os.path.join(r'C:\pdfBaixados', nome)

    # Criar diretório se não existir
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    unique_output_path = os.path.join(output_folder, f'{base_name}{extension}')

    with open(unique_output_path, "wb") as output_file:
        output.write(output_file)

    print(f'NR10 complementar para {nome} preenchida e salva em {unique_output_path}.')
    return unique_output_path

def preencher_nr11(nome, cpf, funcao, dataNR11, Hab_SupInt, nome_TST, n_registroTST, modelo_path, output_path, incluir_funcao=True):
    registrar_fontes()
    pdfmetrics.registerFont(TTFont('Arial-Bold', 'arialbd.ttf'))

    if 'Arial' not in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFont(TTFont('Arial', 'arial.ttf'))

    dataNR11_formatada = formatar_data(dataNR11, formato='curta')

    with open(modelo_path, "rb") as model_file:
        existing_pdf = PdfReader(io.BytesIO(model_file.read()))

    output = PdfWriter()

    for page_num in range(len(existing_pdf.pages)):
        existing_page = existing_pdf.pages[page_num]

        if page_num == 0:  # Somente preencher a primeira página
            packet = io.BytesIO()
            can = canvas.Canvas(packet, pagesize=(800, 1200))

            can.setFont("IBMPlexSans-Bold", 32)
            nome_width = can.stringWidth(nome, "IBMPlexSans-Bold", 32)
            x_position = (800 - nome_width) / 2
            can.drawString(x_position, 380, f'{nome}')

            can.setFont("IBMPlexSans-Bold", 13)
            can.drawString(246.5, 138, f'{cpf}')
#--------------------------------------------------------------
            can.setFont("IBMPlexSans-Bold", 13) 
            can.drawString(455, 123, f'{Hab_SupInt}')

            nome_TST = "Téc. De Seg. Do Trabalho"
            # Use o valor definido na função drawString
            can.setFont("IBMPlexSans-Bold", 13)
            can.drawString(488, 107, f'{nome_TST}')
#--------------------------------------------------------------
            
            can.setFont("IBMPlexSans-Bold", 13) 
            can.drawString(485, 93, f'{n_registroTST}')  # (pros lados, pra cima-baixo,)

            can.setFont("IBMPlexSans-Bold", 18)
            can.drawString(440, 300, f'{dataNR11_formatada}')

            if incluir_funcao:
                can.setFont("IBMPlexSans-Bold", 15)
                can.drawString(326, 450, f'{funcao}')

            can.save()

            packet.seek(0)
            new_pdf_data = packet.getvalue()

            new_page = PdfReader(io.BytesIO(new_pdf_data)).pages[0]
            existing_page.merge_page(new_page)

        output.add_page(existing_page)

    base_name, extension = os.path.splitext(os.path.basename(output_path))
    output_folder = os.path.join(r'C:\pdfBaixados', nome)

    # Criar diretório se não existir
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    unique_output_path = os.path.join(output_folder, f'{base_name}{extension}')

    with open(unique_output_path, "wb") as output_file:
        output.write(output_file)

    print(f'NR11 para {nome} preenchida e salva em {unique_output_path}.')
    return unique_output_path

def preencher_nr12(nome, cpf, funcao, dataNR12, Hab_SupInt, nome_TST, n_registroTST, modelo_path, output_path, incluir_funcao=True):
    registrar_fontes()
    pdfmetrics.registerFont(TTFont('Arial-Bold', 'arialbd.ttf'))

    if 'Arial' not in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFont(TTFont('Arial', 'arial.ttf'))

    dataNR12_formatada = formatar_data(dataNR12, formato='curta')

    with open(modelo_path, "rb") as model_file:
        existing_pdf = PdfReader(io.BytesIO(model_file.read()))

    output = PdfWriter()

    for page_num in range(len(existing_pdf.pages)):
        existing_page = existing_pdf.pages[page_num]

        if page_num == 0:  # Somente preencher a primeira página
            packet = io.BytesIO()
            can = canvas.Canvas(packet, pagesize=(800, 1200))

            can.setFont("IBMPlexSans-Bold", 32)
            nome_width = can.stringWidth(nome, "IBMPlexSans-Bold", 32)
            x_position = (800 - nome_width) / 2
            can.drawString(x_position, 380, f'{nome}')

            can.setFont("IBMPlexSans-Bold", 13)
            can.drawString(246.5, 138, f'{cpf}')
#--------------------------------------------------------------
            can.setFont("IBMPlexSans-Bold", 13) 
            can.drawString(455, 123, f'{Hab_SupInt}')

            nome_TST = "Téc. De Seg. Do Trabalho"
            # Use o valor definido na função drawString
            can.setFont("IBMPlexSans-Bold", 13)
            can.drawString(488, 107, f'{nome_TST}')
#--------------------------------------------------------------
            
            can.setFont("IBMPlexSans-Bold", 13) 
            can.drawString(485, 93, f'{n_registroTST}')  # (pros lados, pra cima-baixo,)

            can.setFont("IBMPlexSans-Bold", 18)
            can.drawString(440, 300, f'{dataNR12_formatada}')

            if incluir_funcao:
                can.setFont("IBMPlexSans-Bold", 15)
                can.drawString(326, 450, f'{funcao}')

            can.save()

            packet.seek(0)
            new_pdf_data = packet.getvalue()

            new_page = PdfReader(io.BytesIO(new_pdf_data)).pages[0]
            existing_page.merge_page(new_page)

        output.add_page(existing_page)

    base_name, extension = os.path.splitext(os.path.basename(output_path))
    output_folder = os.path.join(r'C:\pdfBaixados', nome)

    # Criar diretório se não existir
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    unique_output_path = os.path.join(output_folder, f'{base_name}{extension}')

    with open(unique_output_path, "wb") as output_file:
        output.write(output_file)

    print(f'NR12 para {nome} preenchida e salva em {unique_output_path}.')
    return unique_output_path

def preencher_nr17(nome, cpf, funcao, dataNR17, Hab_SupInt, nome_TST, n_registroTST, modelo_path, output_path, incluir_funcao=True):
    registrar_fontes()
    pdfmetrics.registerFont(TTFont('Arial-Bold', 'arialbd.ttf'))

    if 'Arial' not in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFont(TTFont('Arial', 'arial.ttf'))

    dataNR17_formatada = formatar_data(dataNR17, formato='curta')

    with open(modelo_path, "rb") as model_file:
        existing_pdf = PdfReader(io.BytesIO(model_file.read()))

    output = PdfWriter()

    for page_num in range(len(existing_pdf.pages)):
        existing_page = existing_pdf.pages[page_num]

        if page_num == 0:  # Somente preencher a primeira página
            packet = io.BytesIO()
            can = canvas.Canvas(packet, pagesize=(800, 1200))

            can.setFont("IBMPlexSans-Bold", 32)
            nome_width = can.stringWidth(nome, "IBMPlexSans-Bold", 32)
            x_position = (800 - nome_width) / 2
            can.drawString(x_position, 380, f'{nome}')

            can.setFont("IBMPlexSans-Bold", 13)
            can.drawString(246.5, 138, f'{cpf}')
#--------------------------------------------------------------
            can.setFont("IBMPlexSans-Bold", 13) 
            can.drawString(455, 123, f'{Hab_SupInt}')

            nome_TST = "Téc. De Seg. Do Trabalho"
            # Use o valor definido na função drawString
            can.setFont("IBMPlexSans-Bold", 13)
            can.drawString(488, 107, f'{nome_TST}')
#--------------------------------------------------------------
            
            can.setFont("IBMPlexSans-Bold", 13) 
            can.drawString(485, 93, f'{n_registroTST}')  # (pros lados, pra cima-baixo,)

            can.setFont("IBMPlexSans-Bold", 18)
            can.drawString(440, 300, f'{dataNR17_formatada}')

            if incluir_funcao:
                can.setFont("IBMPlexSans-Bold", 15)
                can.drawString(326, 450, f'{funcao}')

            can.save()

            packet.seek(0)
            new_pdf_data = packet.getvalue()

            new_page = PdfReader(io.BytesIO(new_pdf_data)).pages[0]
            existing_page.merge_page(new_page)

        output.add_page(existing_page)

    base_name, extension = os.path.splitext(os.path.basename(output_path))
    output_folder = os.path.join(r'C:\pdfBaixados', nome)

    # Criar diretório se não existir
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    unique_output_path = os.path.join(output_folder, f'{base_name}{extension}')

    with open(unique_output_path, "wb") as output_file:
        output.write(output_file)

    print(f'NR17 para {nome} preenchida e salva em {unique_output_path}.')
    return unique_output_path

def preencher_nr18_pemt(nome, cpf, funcao, dataNR18_pemt, Hab_SupInt, nome_TST, n_registroTST, modelo_path, output_path, incluir_funcao=True):
    registrar_fontes()
    pdfmetrics.registerFont(TTFont('Arial-Bold', 'arialbd.ttf'))

    if 'Arial' not in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFont(TTFont('Arial', 'arial.ttf'))

    dataNR18_pemt_formatada = formatar_data(dataNR18_pemt, formato='curta')

    with open(modelo_path, "rb") as model_file:
        existing_pdf = PdfReader(io.BytesIO(model_file.read()))

    output = PdfWriter()

    for page_num in range(len(existing_pdf.pages)):
        existing_page = existing_pdf.pages[page_num]

        if page_num == 0:  # Somente preencher a primeira página
            packet = io.BytesIO()
            can = canvas.Canvas(packet, pagesize=(800, 1200))

            can.setFont("IBMPlexSans-Bold", 32)
            nome_width = can.stringWidth(nome, "IBMPlexSans-Bold", 32)
            x_position = (800 - nome_width) / 2
            can.drawString(x_position, 380, f'{nome}')

            can.setFont("IBMPlexSans-Bold", 13)
            can.drawString(246.5, 138, f'{cpf}')
#--------------------------------------------------------------
            can.setFont("IBMPlexSans-Bold", 13) 
            can.drawString(455, 123, f'{Hab_SupInt}')

            nome_TST = "Téc. De Seg. Do Trabalho"
            # Use o valor definido na função drawString
            can.setFont("IBMPlexSans-Bold", 13)
            can.drawString(488, 107, f'{nome_TST}')
#--------------------------------------------------------------
            
            can.setFont("IBMPlexSans-Bold", 13) 
            can.drawString(485, 93, f'{n_registroTST}')  # (pros lados, pra cima-baixo,)

            can.setFont("IBMPlexSans-Bold", 18)
            can.drawString(440, 300, f'{dataNR18_pemt_formatada}')

            if incluir_funcao:
                can.setFont("IBMPlexSans-Bold", 15)
                can.drawString(326, 450, f'{funcao}')

            can.save()

            packet.seek(0)
            new_pdf_data = packet.getvalue()

            new_page = PdfReader(io.BytesIO(new_pdf_data)).pages[0]
            existing_page.merge_page(new_page)

        output.add_page(existing_page)

    base_name, extension = os.path.splitext(os.path.basename(output_path))
    output_folder = os.path.join(r'C:\pdfBaixados', nome)

    # Criar diretório se não existir
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    unique_output_path = os.path.join(output_folder, f'{base_name}{extension}')

    with open(unique_output_path, "wb") as output_file:
        output.write(output_file)

    print(f'NR18 pta para {nome} preenchida e salva em {unique_output_path}.')
    return unique_output_path

def preencher_nr20_infla(nome, cpf, funcao, dataNR20_inflamaveis, Hab_SupInt, nome_TST, n_registroTST, modelo_path, output_path, incluir_funcao=True):
    registrar_fontes()
    pdfmetrics.registerFont(TTFont('Arial-Bold', 'arialbd.ttf'))

    if 'Arial' not in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFont(TTFont('Arial', 'arial.ttf'))

    dataNR20_inflamaveis_formatada = formatar_data(dataNR20_inflamaveis, formato='curta')

    with open(modelo_path, "rb") as model_file:
        existing_pdf = PdfReader(io.BytesIO(model_file.read()))

    output = PdfWriter()

    for page_num in range(len(existing_pdf.pages)):
        existing_page = existing_pdf.pages[page_num]

        if page_num == 0:  # Somente preencher a primeira página
            packet = io.BytesIO()
            can = canvas.Canvas(packet, pagesize=(800, 1200))

            can.setFont("IBMPlexSans-Bold", 32)
            nome_width = can.stringWidth(nome, "IBMPlexSans-Bold", 32)
            x_position = (800 - nome_width) / 2
            can.drawString(x_position, 380, f'{nome}')

            can.setFont("IBMPlexSans-Bold", 13)
            can.drawString(246.5, 138, f'{cpf}')
#--------------------------------------------------------------
            can.setFont("IBMPlexSans-Bold", 13) 
            can.drawString(455, 123, f'{Hab_SupInt}')

            nome_TST = "Téc. De Seg. Do Trabalho"
            # Use o valor definido na função drawString
            can.setFont("IBMPlexSans-Bold", 13)
            can.drawString(488, 107, f'{nome_TST}')
#--------------------------------------------------------------
            
            can.setFont("IBMPlexSans-Bold", 13) 
            can.drawString(485, 93, f'{n_registroTST}')  # (pros lados, pra cima-baixo,)

            can.setFont("IBMPlexSans-Bold", 18)
            can.drawString(440, 300, f'{dataNR20_inflamaveis_formatada}')

            if incluir_funcao:
                can.setFont("IBMPlexSans-Bold", 15)
                can.drawString(326, 450, f'{funcao}')

            can.save()

            packet.seek(0)
            new_pdf_data = packet.getvalue()

            new_page = PdfReader(io.BytesIO(new_pdf_data)).pages[0]
            existing_page.merge_page(new_page)

        output.add_page(existing_page)

    base_name, extension = os.path.splitext(os.path.basename(output_path))
    output_folder = os.path.join(r'C:\pdfBaixados', nome)

    # Criar diretório se não existir
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    unique_output_path = os.path.join(output_folder, f'{base_name}{extension}')

    with open(unique_output_path, "wb") as output_file:
        output.write(output_file)

    print(f'NR20 inflamaveis pta para {nome} preenchida e salva em {unique_output_path}.')
    return unique_output_path

def preencher_nr20_brigada(nome, cpf, funcao, dataNR20_brigada, Hab_SupInt, nome_TST, n_registroTST, modelo_path, output_path, incluir_funcao=True):
    registrar_fontes()
    pdfmetrics.registerFont(TTFont('Arial-Bold', 'arialbd.ttf'))

    if 'Arial' not in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFont(TTFont('Arial', 'arial.ttf'))

    dataNR20_brigada_formatada = formatar_data(dataNR20_brigada, formato='curta')

    with open(modelo_path, "rb") as model_file:
        existing_pdf = PdfReader(io.BytesIO(model_file.read()))

    output = PdfWriter()

    for page_num in range(len(existing_pdf.pages)):
        existing_page = existing_pdf.pages[page_num]

        if page_num == 0:  # Somente preencher a primeira página
            packet = io.BytesIO()
            can = canvas.Canvas(packet, pagesize=(800, 1200))

            can.setFont("IBMPlexSans-Bold", 32)
            nome_width = can.stringWidth(nome, "IBMPlexSans-Bold", 32)
            x_position = (800 - nome_width) / 2
            can.drawString(x_position, 380, f'{nome}')

            can.setFont("IBMPlexSans-Bold", 13)
            can.drawString(246.5, 138, f'{cpf}')
#--------------------------------------------------------------
            can.setFont("IBMPlexSans-Bold", 13) 
            can.drawString(455, 123, f'{Hab_SupInt}')

            nome_TST = "Téc. De Seg. Do Trabalho"
            # Use o valor definido na função drawString
            can.setFont("IBMPlexSans-Bold", 13)
            can.drawString(488, 107, f'{nome_TST}')
#--------------------------------------------------------------
            
            can.setFont("IBMPlexSans-Bold", 13) 
            can.drawString(485, 93, f'{n_registroTST}')  # (pros lados, pra cima-baixo,)

            can.setFont("IBMPlexSans-Bold", 18)
            can.drawString(440, 300, f'{dataNR20_brigada_formatada}')

            if incluir_funcao:
                can.setFont("IBMPlexSans-Bold", 15)
                can.drawString(326, 450, f'{funcao}')

            can.save()

            packet.seek(0)
            new_pdf_data = packet.getvalue()

            new_page = PdfReader(io.BytesIO(new_pdf_data)).pages[0]
            existing_page.merge_page(new_page)

        output.add_page(existing_page)

    base_name, extension = os.path.splitext(os.path.basename(output_path))
    output_folder = os.path.join(r'C:\pdfBaixados', nome)

    # Criar diretório se não existir
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    unique_output_path = os.path.join(output_folder, f'{base_name}{extension}')

    with open(unique_output_path, "wb") as output_file:
        output.write(output_file)

    print(f'NR20 Brigada para {nome} preenchida e salva em {unique_output_path}.')
    return unique_output_path

def preencher_nr33(nome, cpf, funcao, dataNR33, Hab_SupInt, nome_TST, n_registroTST, modelo_path, output_path, incluir_funcao=True):
    registrar_fontes()
    pdfmetrics.registerFont(TTFont('Arial-Bold', 'arialbd.ttf'))

    if 'Arial' not in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFont(TTFont('Arial', 'arial.ttf'))

    dataNR33_formatada = formatar_data(dataNR33, formato='curta')

    with open(modelo_path, "rb") as model_file:
        existing_pdf = PdfReader(io.BytesIO(model_file.read()))

    output = PdfWriter()

    for page_num in range(len(existing_pdf.pages)):
        existing_page = existing_pdf.pages[page_num]

        if page_num == 0:  # Somente preencher a primeira página
            packet = io.BytesIO()
            can = canvas.Canvas(packet, pagesize=(800, 1200))

            can.setFont("IBMPlexSans-Bold", 32)
            nome_width = can.stringWidth(nome, "IBMPlexSans-Bold", 32)
            x_position = (800 - nome_width) / 2
            can.drawString(x_position, 380, f'{nome}')

            can.setFont("IBMPlexSans-Bold", 13)
            can.drawString(246.5, 138, f'{cpf}')
#--------------------------------------------------------------
            can.setFont("IBMPlexSans-Bold", 13) 
            can.drawString(455, 123, f'{Hab_SupInt}')

            nome_TST = "Téc. De Seg. Do Trabalho"
            # Use o valor definido na função drawString
            can.setFont("IBMPlexSans-Bold", 13)
            can.drawString(488, 107, f'{nome_TST}')
#--------------------------------------------------------------
            
            can.setFont("IBMPlexSans-Bold", 13) 
            can.drawString(485, 93, f'{n_registroTST}')  # (pros lados, pra cima-baixo,)

            can.setFont("IBMPlexSans-Bold", 18)
            can.drawString(440, 300, f'{dataNR33_formatada}')

            if incluir_funcao:
                can.setFont("IBMPlexSans-Bold", 15)
                can.drawString(326, 450, f'{funcao}')

            can.save()

            packet.seek(0)
            new_pdf_data = packet.getvalue()

            new_page = PdfReader(io.BytesIO(new_pdf_data)).pages[0]
            existing_page.merge_page(new_page)

        output.add_page(existing_page)

    base_name, extension = os.path.splitext(os.path.basename(output_path))
    output_folder = os.path.join(r'C:\pdfBaixados', nome)

    # Criar diretório se não existir
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    unique_output_path = os.path.join(output_folder, f'{base_name}{extension}')

    with open(unique_output_path, "wb") as output_file:
        output.write(output_file)

    print(f'NR33 para {nome} preenchida e salva em {unique_output_path}.')
    return unique_output_path

def preencher_nr34(nome, cpf, funcao, dataNR34, Hab_SupInt, nome_TST, n_registroTST, modelo_path, output_path, incluir_funcao=True):
    registrar_fontes()
    pdfmetrics.registerFont(TTFont('Arial-Bold', 'arialbd.ttf'))

    if 'Arial' not in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFont(TTFont('Arial', 'arial.ttf'))

    dataNR34_formatada = formatar_data(dataNR34, formato='curta')

    with open(modelo_path, "rb") as model_file:
        existing_pdf = PdfReader(io.BytesIO(model_file.read()))

    output = PdfWriter()

    for page_num in range(len(existing_pdf.pages)):
        existing_page = existing_pdf.pages[page_num]

        if page_num == 0:  # Somente preencher a primeira página
            packet = io.BytesIO()
            can = canvas.Canvas(packet, pagesize=(800, 1200))

            can.setFont("IBMPlexSans-Bold", 32)
            nome_width = can.stringWidth(nome, "IBMPlexSans-Bold", 32)
            x_position = (800 - nome_width) / 2
            can.drawString(x_position, 380, f'{nome}')

            can.setFont("IBMPlexSans-Bold", 13)
            can.drawString(246.5, 138, f'{cpf}')
#--------------------------------------------------------------
            can.setFont("IBMPlexSans-Bold", 13) 
            can.drawString(455, 123, f'{Hab_SupInt}')

            nome_TST = "Téc. De Seg. Do Trabalho"
            # Use o valor definido na função drawString
            can.setFont("IBMPlexSans-Bold", 13)
            can.drawString(488, 107, f'{nome_TST}')
#--------------------------------------------------------------
            
            can.setFont("IBMPlexSans-Bold", 13) 
            can.drawString(485, 93, f'{n_registroTST}')  # (pros lados, pra cima-baixo,)

            can.setFont("IBMPlexSans-Bold", 18)
            can.drawString(440, 300, f'{dataNR34_formatada}')

            if incluir_funcao:
                can.setFont("IBMPlexSans-Bold", 15)
                can.drawString(326, 450, f'{funcao}')

            can.save()

            packet.seek(0)
            new_pdf_data = packet.getvalue()

            new_page = PdfReader(io.BytesIO(new_pdf_data)).pages[0]
            existing_page.merge_page(new_page)

        output.add_page(existing_page)

    base_name, extension = os.path.splitext(os.path.basename(output_path))
    output_folder = os.path.join(r'C:\pdfBaixados', nome)

    # Criar diretório se não existir
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    unique_output_path = os.path.join(output_folder, f'{base_name}{extension}')

    with open(unique_output_path, "wb") as output_file:
        output.write(output_file)

    print(f'NR34 Basica para {nome} preenchida e salva em {unique_output_path}.')
    return unique_output_path

def preencher_nr34_adm(nome, cpf, funcao, dataNR34_adm, Hab_SupInt, nome_TST, n_registroTST, modelo_path, output_path, incluir_funcao=True):
    registrar_fontes()
    pdfmetrics.registerFont(TTFont('Arial-Bold', 'arialbd.ttf'))

    if 'Arial' not in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFont(TTFont('Arial', 'arial.ttf'))

    dataNR34_adm_formatada = formatar_data(dataNR34_adm, formato='curta')

    with open(modelo_path, "rb") as model_file:
        existing_pdf = PdfReader(io.BytesIO(model_file.read()))

    output = PdfWriter()

    for page_num in range(len(existing_pdf.pages)):
        existing_page = existing_pdf.pages[page_num]

        if page_num == 0:  # Somente preencher a primeira página
            packet = io.BytesIO()
            can = canvas.Canvas(packet, pagesize=(800, 1200))

            can.setFont("IBMPlexSans-Bold", 32)
            nome_width = can.stringWidth(nome, "IBMPlexSans-Bold", 32)
            x_position = (800 - nome_width) / 2
            can.drawString(x_position, 380, f'{nome}')

            can.setFont("IBMPlexSans-Bold", 13)
            can.drawString(246.5, 138, f'{cpf}')
#--------------------------------------------------------------
            can.setFont("IBMPlexSans-Bold", 13) 
            can.drawString(455, 123, f'{Hab_SupInt}')

            nome_TST = "Téc. De Seg. Do Trabalho"
            # Use o valor definido na função drawString
            can.setFont("IBMPlexSans-Bold", 13)
            can.drawString(488, 107, f'{nome_TST}')
#--------------------------------------------------------------
            
            can.setFont("IBMPlexSans-Bold", 13) 
            can.drawString(485, 93, f'{n_registroTST}')  # (pros lados, pra cima-baixo,)

            can.setFont("IBMPlexSans-Bold", 18)
            can.drawString(440, 300, f'{dataNR34_adm_formatada}')

            if incluir_funcao:
                can.setFont("IBMPlexSans-Bold", 15)
                can.drawString(326, 450, f'{funcao}')

            can.save()

            packet.seek(0)
            new_pdf_data = packet.getvalue()

            new_page = PdfReader(io.BytesIO(new_pdf_data)).pages[0]
            existing_page.merge_page(new_page)

        output.add_page(existing_page)

    base_name, extension = os.path.splitext(os.path.basename(output_path))
    output_folder = os.path.join(r'C:\pdfBaixados', nome)

    # Criar diretório se não existir
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    unique_output_path = os.path.join(output_folder, f'{base_name}{extension}')

    with open(unique_output_path, "wb") as output_file:
        output.write(output_file)

    print(f'NR01 para {nome} preenchida e salva em {unique_output_path}.')
    return unique_output_path

def preencher_nr34_obs_quente(nome, cpf, funcao, dataNR34_obs_quente, Hab_SupInt, nome_TST, n_registroTST, modelo_path, output_path, incluir_funcao=True):
    registrar_fontes()
    pdfmetrics.registerFont(TTFont('Arial-Bold', 'arialbd.ttf'))

    if 'Arial' not in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFont(TTFont('Arial', 'arial.ttf'))

    dataNR34_obs_quente_formatada = formatar_data(dataNR34_obs_quente, formato='curta')

    with open(modelo_path, "rb") as model_file:
        existing_pdf = PdfReader(io.BytesIO(model_file.read()))

    output = PdfWriter()

    for page_num in range(len(existing_pdf.pages)):
        existing_page = existing_pdf.pages[page_num]

        if page_num == 0:  # Somente preencher a primeira página
            packet = io.BytesIO()
            can = canvas.Canvas(packet, pagesize=(800, 1200))

            can.setFont("IBMPlexSans-Bold", 32)
            nome_width = can.stringWidth(nome, "IBMPlexSans-Bold", 32)
            x_position = (800 - nome_width) / 2
            can.drawString(x_position, 380, f'{nome}')

            can.setFont("IBMPlexSans-Bold", 13)
            can.drawString(246.5, 138, f'{cpf}')
#--------------------------------------------------------------
            can.setFont("IBMPlexSans-Bold", 13) 
            can.drawString(455, 123, f'{Hab_SupInt}')

            nome_TST = "Téc. De Seg. Do Trabalho"
            # Use o valor definido na função drawString
            can.setFont("IBMPlexSans-Bold", 13)
            can.drawString(488, 107, f'{nome_TST}')
#--------------------------------------------------------------
            
            can.setFont("IBMPlexSans-Bold", 13) 
            can.drawString(485, 93, f'{n_registroTST}')  # (pros lados, pra cima-baixo,)

            can.setFont("IBMPlexSans-Bold", 18)
            can.drawString(440, 300, f'{dataNR34_obs_quente_formatada}')

            if incluir_funcao:
                can.setFont("IBMPlexSans-Bold", 15)
                can.drawString(326, 450, f'{funcao}')

            can.save()

            packet.seek(0)
            new_pdf_data = packet.getvalue()

            new_page = PdfReader(io.BytesIO(new_pdf_data)).pages[0]
            existing_page.merge_page(new_page)

        output.add_page(existing_page)

    base_name, extension = os.path.splitext(os.path.basename(output_path))
    output_folder = os.path.join(r'C:\pdfBaixados', nome)

    # Criar diretório se não existir
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    unique_output_path = os.path.join(output_folder, f'{base_name}{extension}')

    with open(unique_output_path, "wb") as output_file:
        output.write(output_file)

    print(f'NR34 OBS QUENTE para {nome} preenchida e salva em {unique_output_path}.')
    return unique_output_path

def preencher_fichaEPI(nome, funcao, n_registroTST, modelo_path, output_path, incluir_funcao=True):
    pdfmetrics.registerFont(TTFont('Arial-Bold', 'arialbd.ttf'))

    if 'Arial' not in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFont(TTFont('Arial', 'arial.ttf'))

    packet = io.BytesIO()
    can = canvas.Canvas(packet, pagesize=(800, 1200))

    can.setFont("IBMPlexSans-Bold", 13)
    can.drawString(110, 758, f'{nome}')

    can.setFont("IBMPlexSans-Bold", 13)
    can.drawString(390, 737, f'{n_registroTST}')

    if incluir_funcao:
        can.setFont("IBMPlexSans-Bold", 13)
        can.drawString(110,737 , f'{funcao}')


    can.save()

    packet.seek(0)
    new_pdf_data = packet.getvalue()

    with open(modelo_path, "rb") as model_file:
        existing_pdf = PdfReader(io.BytesIO(model_file.read()))

    output_folder = os.path.join(r'C:\pdfBaixados', nome)
    
    # Criar diretório se não existir
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    base_name, extension = os.path.splitext(os.path.basename(output_path))
    unique_output_path = os.path.join(output_folder, f'{base_name}{extension}')

    output = PdfWriter()
    page = existing_pdf.pages[0]
    page.merge_page(PdfReader(io.BytesIO(new_pdf_data)).pages[0])
    output.add_page(page)

    with open(unique_output_path, "wb") as output_file:
        output.write(output_file)

    print(f'Ficha de EPI para {nome} preenchida e salva em {unique_output_path}.')
    return unique_output_path

def preencher_fichaEPI_adm_geral(nome, funcao, n_registroTST, modelo_path, output_path, incluir_funcao=True):
    pdfmetrics.registerFont(TTFont('Arial-Bold', 'arialbd.ttf'))

    if 'Arial' not in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFont(TTFont('Arial', 'arial.ttf'))

    packet = io.BytesIO()
    can = canvas.Canvas(packet, pagesize=(800, 1200))

    can.setFont("IBMPlexSans-Bold", 13)
    can.drawString(94, 729, f'{nome}')

    can.setFont("IBMPlexSans-Bold", 13)
    can.drawString(355, 710, f'{n_registroTST}')

    if incluir_funcao:
        can.setFont("IBMPlexSans-Bold", 13)
        can.drawString(94,710 , f'{funcao}')


    can.save()

    packet.seek(0)
    new_pdf_data = packet.getvalue()

    with open(modelo_path, "rb") as model_file:
        existing_pdf = PdfReader(io.BytesIO(model_file.read()))

    output_folder = os.path.join(r'C:\pdfBaixados', nome)
    
    # Criar diretório se não existir
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    base_name, extension = os.path.splitext(os.path.basename(output_path))
    unique_output_path = os.path.join(output_folder, f'{base_name}{extension}')

    output = PdfWriter()
    page = existing_pdf.pages[0]
    page.merge_page(PdfReader(io.BytesIO(new_pdf_data)).pages[0])
    output.add_page(page)

    with open(unique_output_path, "wb") as output_file:
        output.write(output_file)

    print(f'Ficha de EPI para {nome} preenchida e salva em {unique_output_path}.')
    return unique_output_path


def preencher_fichaEPI_adm_obra(nome, funcao, n_registroTST, modelo_path, output_path, incluir_funcao=True):
    pdfmetrics.registerFont(TTFont('Arial-Bold', 'arialbd.ttf'))

    if 'Arial' not in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFont(TTFont('Arial', 'arial.ttf'))

    packet = io.BytesIO()
    can = canvas.Canvas(packet, pagesize=(800, 1200))

    can.setFont("IBMPlexSans-Bold", 13)
    can.drawString(94, 729, f'{nome}')

    can.setFont("IBMPlexSans-Bold", 13)
    can.drawString(355, 710, f'{n_registroTST}')

    if incluir_funcao:
        can.setFont("IBMPlexSans-Bold", 13)
        can.drawString(94,710 , f'{funcao}')


    can.save()

    packet.seek(0)
    new_pdf_data = packet.getvalue()

    with open(modelo_path, "rb") as model_file:
        existing_pdf = PdfReader(io.BytesIO(model_file.read()))

    output_folder = os.path.join(r'C:\pdfBaixados', nome)
    
    # Criar diretório se não existir
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    base_name, extension = os.path.splitext(os.path.basename(output_path))
    unique_output_path = os.path.join(output_folder, f'{base_name}{extension}')

    output = PdfWriter()
    page = existing_pdf.pages[0]
    page.merge_page(PdfReader(io.BytesIO(new_pdf_data)).pages[0])
    output.add_page(page)

    with open(unique_output_path, "wb") as output_file:
        output.write(output_file)

    print(f'Ficha de EPI para {nome} preenchida e salva em {unique_output_path}.')
    return unique_output_path

def preencher_fichaEPI_almoxarife(nome, funcao, n_registroTST, modelo_path, output_path, incluir_funcao=True):
    pdfmetrics.registerFont(TTFont('Arial-Bold', 'arialbd.ttf'))

    if 'Arial' not in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFont(TTFont('Arial', 'arial.ttf'))

    packet = io.BytesIO()
    can = canvas.Canvas(packet, pagesize=(800, 1200))

    can.setFont("IBMPlexSans-Bold", 13)
    can.drawString(94, 729, f'{nome}')

    can.setFont("IBMPlexSans-Bold", 13)
    can.drawString(355, 710, f'{n_registroTST}')

    if incluir_funcao:
        can.setFont("IBMPlexSans-Bold", 13)
        can.drawString(94,710 , f'{funcao}')


    can.save()

    packet.seek(0)
    new_pdf_data = packet.getvalue()

    with open(modelo_path, "rb") as model_file:
        existing_pdf = PdfReader(io.BytesIO(model_file.read()))

    output_folder = os.path.join(r'C:\pdfBaixados', nome)
    
    # Criar diretório se não existir
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    base_name, extension = os.path.splitext(os.path.basename(output_path))
    unique_output_path = os.path.join(output_folder, f'{base_name}{extension}')

    output = PdfWriter()
    page = existing_pdf.pages[0]
    page.merge_page(PdfReader(io.BytesIO(new_pdf_data)).pages[0])
    output.add_page(page)

    with open(unique_output_path, "wb") as output_file:
        output.write(output_file)

    print(f'Ficha de EPI para {nome} preenchida e salva em {unique_output_path}.')
    return unique_output_path

def preencher_fichaEPI_civil(nome, funcao, n_registroTST, modelo_path, output_path, incluir_funcao=True):
    pdfmetrics.registerFont(TTFont('Arial-Bold', 'arialbd.ttf'))

    if 'Arial' not in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFont(TTFont('Arial', 'arial.ttf'))

    packet = io.BytesIO()
    can = canvas.Canvas(packet, pagesize=(800, 1200))

    can.setFont("IBMPlexSans-Bold", 13)
    can.drawString(94, 729, f'{nome}')

    can.setFont("IBMPlexSans-Bold", 13)
    can.drawString(355, 710, f'{n_registroTST}')

    if incluir_funcao:
        can.setFont("IBMPlexSans-Bold", 13)
        can.drawString(94,710 , f'{funcao}')


    can.save()

    packet.seek(0)
    new_pdf_data = packet.getvalue()

    with open(modelo_path, "rb") as model_file:
        existing_pdf = PdfReader(io.BytesIO(model_file.read()))

    output_folder = os.path.join(r'C:\pdfBaixados', nome)
    
    # Criar diretório se não existir
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    base_name, extension = os.path.splitext(os.path.basename(output_path))
    unique_output_path = os.path.join(output_folder, f'{base_name}{extension}')

    output = PdfWriter()
    page = existing_pdf.pages[0]
    page.merge_page(PdfReader(io.BytesIO(new_pdf_data)).pages[0])
    output.add_page(page)

    with open(unique_output_path, "wb") as output_file:
        output.write(output_file)

    print(f'Ficha de EPI para {nome} preenchida e salva em {unique_output_path}.')
    return unique_output_path

def preencher_fichaEPI_hidra(nome, funcao, n_registroTST, modelo_path, output_path, incluir_funcao=True):
    pdfmetrics.registerFont(TTFont('Arial-Bold', 'arialbd.ttf'))

    if 'Arial' not in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFont(TTFont('Arial', 'arial.ttf'))

    packet = io.BytesIO()
    can = canvas.Canvas(packet, pagesize=(800, 1200))

    can.setFont("IBMPlexSans-Bold", 13)
    can.drawString(94, 729, f'{nome}')

    can.setFont("IBMPlexSans-Bold", 13)
    can.drawString(355, 710, f'{n_registroTST}')

    if incluir_funcao:
        can.setFont("IBMPlexSans-Bold", 13)
        can.drawString(94,710 , f'{funcao}')


    can.save()

    packet.seek(0)
    new_pdf_data = packet.getvalue()

    with open(modelo_path, "rb") as model_file:
        existing_pdf = PdfReader(io.BytesIO(model_file.read()))

    output_folder = os.path.join(r'C:\pdfBaixados', nome)
    
    # Criar diretório se não existir
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    base_name, extension = os.path.splitext(os.path.basename(output_path))
    unique_output_path = os.path.join(output_folder, f'{base_name}{extension}')

    output = PdfWriter()
    page = existing_pdf.pages[0]
    page.merge_page(PdfReader(io.BytesIO(new_pdf_data)).pages[0])
    output.add_page(page)

    with open(unique_output_path, "wb") as output_file:
        output.write(output_file)

    print(f'Ficha de EPI para {nome} preenchida e salva em {unique_output_path}.')
    return unique_output_path

def preencher_fichaEPI_solda(nome, funcao, n_registroTST, modelo_path, output_path, incluir_funcao=True):
    pdfmetrics.registerFont(TTFont('Arial-Bold', 'arialbd.ttf'))

    if 'Arial' not in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFont(TTFont('Arial', 'arial.ttf'))

    packet = io.BytesIO()
    can = canvas.Canvas(packet, pagesize=(800, 1200))

    can.setFont("IBMPlexSans-Bold", 13)
    can.drawString(94, 729, f'{nome}')

    can.setFont("IBMPlexSans-Bold", 13)
    can.drawString(355, 710, f'{n_registroTST}')

    if incluir_funcao:
        can.setFont("IBMPlexSans-Bold", 13)
        can.drawString(94,710 , f'{funcao}')


    can.save()

    packet.seek(0)
    new_pdf_data = packet.getvalue()

    with open(modelo_path, "rb") as model_file:
        existing_pdf = PdfReader(io.BytesIO(model_file.read()))

    output_folder = os.path.join(r'C:\pdfBaixados', nome)
    
    # Criar diretório se não existir
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    base_name, extension = os.path.splitext(os.path.basename(output_path))
    unique_output_path = os.path.join(output_folder, f'{base_name}{extension}')

    output = PdfWriter()
    page = existing_pdf.pages[0]
    page.merge_page(PdfReader(io.BytesIO(new_pdf_data)).pages[0])
    output.add_page(page)

    with open(unique_output_path, "wb") as output_file:
        output.write(output_file)

    print(f'Ficha de EPI para {nome} preenchida e salva em {unique_output_path}.')
    return unique_output_path






def preencher_CA(nome, cpf, funcao, Hab_SupInt, n_superInt, cpf_superInt, nomeTecRep, modelo_path, output_path, incluir_funcao=True):
    pdfmetrics.registerFont(TTFont('Arial-Bold', 'arialbd.ttf'))

    if 'Arial' not in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFont(TTFont('Arial', 'arial.ttf'))

    packet = io.BytesIO()
    can = canvas.Canvas(packet, pagesize=(800, 1200))

    documento_texto = f'empregado(a)    {nome}    CPF N°:   {cpf}    , empregado (a)  desta  empresa , ocupante do cargo'

    can.setFont("IBMPlexSans-Text", 9)
    can.drawString(20, 721.2, documento_texto)

    documento_texto2 = f'Pelo presente documento,eu, {nomeTecRep} ,{Hab_SupInt} registrado no CREA-SP, sob o n°  {n_superInt}'

    can.setFont("IBMPlexSans-Text", 9)
    can.drawString(20, 769, documento_texto2)



    nome_absoluto_Planem = "Planem Engenharia e Eletricidade Ltda"
    can.setFont("IBMPlexSans-Text", 7)
    can.drawString(20, 126, f'{nome_absoluto_Planem}')

  # documentação header

    can.setFont("IBMPlexSans-Text", 7)
    # Adiciona o separador " - " apenas para o rodapé na coordenada 135
    can.drawString(20, 135, f'{Hab_SupInt} - {n_superInt}')  # documentação rodape\assinatura

    can.setFont("IBMPlexSans-Text", 9)
    can.drawString(56, 753, f'{cpf_superInt}')


    can.setFont("IBMPlexSans-Text", 7)
    can.drawString(20, 145, f'{nomeTecRep}')  # documentação rodape\assinatura

    can.setFont("IBMPlexSans-Text", 7)
    can.drawString(20, 85, f'{nome}')

    if incluir_funcao:
        funcao_texto = f'{funcao} está autorizado(a) formalmente pela empresa a realizar a(s) seguinte(s) atividade(s):'
        can.setFont("IBMPlexSans-Text", 9)
        can.drawString(20, 705.5, funcao_texto)

        can.setFont("IBMPlexSans-Text", 7)
        can.drawString(20, 76.5, f'{funcao}')

    can.save()

    packet.seek(0)
    new_pdf_data = packet.getvalue()

    with open(modelo_path, "rb") as model_file:
        existing_pdf = PdfReader(io.BytesIO(model_file.read()))

    output_folder = os.path.join(r'C:\pdfBaixados', nome)

    # Criar diretório se não existir
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    base_name, extension = os.path.splitext(os.path.basename(output_path))
    unique_output_path = os.path.join(output_folder, f'{base_name}{extension}')

    output = PdfWriter()
    page = existing_pdf.pages[0]
    page.merge_page(PdfReader(io.BytesIO(new_pdf_data)).pages[0])
    output.add_page(page)

    with open(unique_output_path, "wb") as output_file:
        output.write(output_file)

    print(f'Carta de Anuencia para {nome} preenchida e salva em {unique_output_path}.')
    return unique_output_path

def preencher_OS_adm_geral(nome, cpf, funcao, nomeTecRep, n_superInt, modelo_path, output_path, incluir_funcao=True, pagina_para_preencher=0):
    if 'Arial' not in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFont(TTFont('Arial', 'arial.ttf'))

    with open(modelo_path, "rb") as model_file:
        existing_pdf = PdfReader(io.BytesIO(model_file.read()))

    output_folder = os.path.join(r'C:\pdfBaixados', nome)

    # Criar diretório se não existir
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    base_name, extension = os.path.splitext(os.path.basename(output_path))
    unique_output_path = os.path.join(output_folder, f'{base_name}{extension}')

    output = PdfWriter()

    for page_num in range(len(existing_pdf.pages)):
        existing_page = existing_pdf.pages[page_num]

        if page_num == pagina_para_preencher:
            packet = io.BytesIO()
            can = canvas.Canvas(packet, pagesize=letter)
            can.setFont("IBMPlexSans-Bold", 10)
            can.drawString(115, 738, f'{nome}')
            if incluir_funcao:
                can.setFont("IBMPlexSans-Bold", 10)
                can.drawString(85, 721, f'{funcao}')
            can.save()

            packet.seek(0)
            new_pdf_data = packet.getvalue()
            new_page = PdfReader(io.BytesIO(new_pdf_data)).pages[0]
            existing_page.merge_page(new_page)

        elif page_num == 3:  # Página quarta    
            packet = io.BytesIO()
            can = canvas.Canvas(packet, pagesize=letter)
            can.setFont("IBMPlexSans-Bold", 10)
            can.drawString(113, 617, f'{nome}')  # Ajuste as coordenadas conforme necessário
            can.setFont("IBMPlexSans-Bold", 10)
            can.drawString(332, 600, f'{cpf}') 
#--------------------------------------------------
            can.setFont("IBMPlexSans-Bold", 10)
            can.drawString(145, 501.5, f'{n_superInt}') 

            can.setFont("IBMPlexSans-Bold", 10)
            can.drawString(151, 519 , f'{nomeTecRep}') 
            
            if incluir_funcao:
                can.setFont("IBMPlexSans-Bold",10)
                can.drawString(85, 599, f'{funcao}')
            can.save()

            packet.seek(0)
            new_pdf_data = packet.getvalue()
            new_page = PdfReader(io.BytesIO(new_pdf_data)).pages[0]
            existing_page.merge_page(new_page)

        output.add_page(existing_page)

    with open(unique_output_path, "wb") as output_file:
        output.write(output_file)

    print(f'OS adm geral para {nome} preenchida e salva em {unique_output_path}.')
    return unique_output_path

def preencher_OS_adm_obra(nome, cpf, funcao, nomeTecRep, n_superInt, modelo_path, output_path, incluir_funcao=True, pagina_para_preencher=0):
    if 'Arial' not in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFont(TTFont('Arial', 'arial.ttf'))

    with open(modelo_path, "rb") as model_file:
        existing_pdf = PdfReader(io.BytesIO(model_file.read()))

    output_folder = os.path.join(r'C:\pdfBaixados', nome)

    # Criar diretório se não existir
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    base_name, extension = os.path.splitext(os.path.basename(output_path))
    unique_output_path = os.path.join(output_folder, f'{base_name}{extension}')

    output = PdfWriter()

    for page_num in range(len(existing_pdf.pages)):
        existing_page = existing_pdf.pages[page_num]

        if page_num == pagina_para_preencher:
            packet = io.BytesIO()
            can = canvas.Canvas(packet, pagesize=letter)
            can.setFont("IBMPlexSans-Bold", 10)
            can.drawString(115, 738, f'{nome}')
            if incluir_funcao:
                can.setFont("IBMPlexSans-Bold", 10)
                can.drawString(85, 721, f'{funcao}')
            can.save()

            packet.seek(0)
            new_pdf_data = packet.getvalue()
            new_page = PdfReader(io.BytesIO(new_pdf_data)).pages[0]
            existing_page.merge_page(new_page)

        elif page_num == 3:  # Página quarta    
            packet = io.BytesIO()
            can = canvas.Canvas(packet, pagesize=letter)
            can.setFont("IBMPlexSans-Bold", 10)
            can.drawString(113, 617, f'{nome}')  # Ajuste as coordenadas conforme necessário
            can.setFont("IBMPlexSans-Bold", 10)
            can.drawString(332, 600, f'{cpf}') 
            #--------------------------------------------------
            can.setFont("IBMPlexSans-Bold", 10)
            can.drawString(145, 501.5, f'{n_superInt}') 

            can.setFont("IBMPlexSans-Bold", 10)
            can.drawString(151, 519 , f'{nomeTecRep}') 
            if incluir_funcao:
                can.setFont("IBMPlexSans-Bold",10)
                can.drawString(85, 599, f'{funcao}')
            can.save()

            packet.seek(0)
            new_pdf_data = packet.getvalue()
            new_page = PdfReader(io.BytesIO(new_pdf_data)).pages[0]
            existing_page.merge_page(new_page)

        output.add_page(existing_page)

    with open(unique_output_path, "wb") as output_file:
        output.write(output_file)

    print(f'OS adm obra para {nome} preenchida e salva em {unique_output_path}.')
    return unique_output_path

def preencher_OS_aumoxarifado(nome, cpf, funcao, nomeTecRep, n_superInt, modelo_path, output_path, incluir_funcao=True, pagina_para_preencher=0):
    if 'Arial' not in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFont(TTFont('Arial', 'arial.ttf'))

    with open(modelo_path, "rb") as model_file:
        existing_pdf = PdfReader(io.BytesIO(model_file.read()))

    output_folder = os.path.join(r'C:\pdfBaixados', nome)

    # Criar diretório se não existir
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    base_name, extension = os.path.splitext(os.path.basename(output_path))
    unique_output_path = os.path.join(output_folder, f'{base_name}{extension}')

    output = PdfWriter()

    for page_num in range(len(existing_pdf.pages)):
        existing_page = existing_pdf.pages[page_num]

        if page_num == pagina_para_preencher:
            packet = io.BytesIO()
            can = canvas.Canvas(packet, pagesize=letter)
            can.setFont("IBMPlexSans-Bold", 10)
            can.drawString(115, 738, f'{nome}')
            if incluir_funcao:
                can.setFont("IBMPlexSans-Bold", 10)
                can.drawString(85, 721, f'{funcao}')
            can.save()

            packet.seek(0)
            new_pdf_data = packet.getvalue()
            new_page = PdfReader(io.BytesIO(new_pdf_data)).pages[0]
            existing_page.merge_page(new_page)

        elif page_num == 3:  # Página quarta    
            packet = io.BytesIO()
            can = canvas.Canvas(packet, pagesize=letter)
            can.setFont("IBMPlexSans-Bold", 10)
            can.drawString(113, 617, f'{nome}')  # Ajuste as coordenadas conforme necessário
            can.setFont("IBMPlexSans-Bold", 10)
            can.drawString(332, 600, f'{cpf}') 
            #--------------------------------------------------
            can.setFont("IBMPlexSans-Bold", 10)
            can.drawString(145, 501.5, f'{n_superInt}') 

            can.setFont("IBMPlexSans-Bold", 10)
            can.drawString(151, 519 , f'{nomeTecRep}') 
            
            if incluir_funcao:
                can.setFont("IBMPlexSans-Bold",10)
                can.drawString(85, 599, f'{funcao}')
            can.save()

            packet.seek(0)
            new_pdf_data = packet.getvalue()
            new_page = PdfReader(io.BytesIO(new_pdf_data)).pages[0]
            existing_page.merge_page(new_page)

        output.add_page(existing_page)

    with open(unique_output_path, "wb") as output_file:
        output.write(output_file)

    print(f'OS adm geral para {nome} preenchida e salva em {unique_output_path}.')
    return unique_output_path

def preencher_OS_obras_civil(nome, cpf, funcao, nomeTecRep, n_superInt, modelo_path, output_path, incluir_funcao=True, pagina_para_preencher=0):
    if 'Arial' not in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFont(TTFont('Arial', 'arial.ttf'))

    with open(modelo_path, "rb") as model_file:
        existing_pdf = PdfReader(io.BytesIO(model_file.read()))

    output_folder = os.path.join(r'C:\pdfBaixados', nome)

    # Criar diretório se não existir
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    base_name, extension = os.path.splitext(os.path.basename(output_path))
    unique_output_path = os.path.join(output_folder, f'{base_name}{extension}')

    output = PdfWriter()

    for page_num in range(len(existing_pdf.pages)):
        existing_page = existing_pdf.pages[page_num]

        if page_num == pagina_para_preencher:
            packet = io.BytesIO()
            can = canvas.Canvas(packet, pagesize=letter)
            can.setFont("IBMPlexSans-Bold", 10)
            can.drawString(115, 738, f'{nome}')
            if incluir_funcao:
                can.setFont("IBMPlexSans-Bold", 10)
                can.drawString(85, 721, f'{funcao}')
            can.save()

            packet.seek(0)
            new_pdf_data = packet.getvalue()
            new_page = PdfReader(io.BytesIO(new_pdf_data)).pages[0]
            existing_page.merge_page(new_page)

        elif page_num == 3:  # Página quarta    
            packet = io.BytesIO()
            can = canvas.Canvas(packet, pagesize=letter)
            can.setFont("IBMPlexSans-Bold", 10)
            can.drawString(113, 617, f'{nome}')  # Ajuste as coordenadas conforme necessário
            can.setFont("IBMPlexSans-Bold", 10)
            can.drawString(332, 600, f'{cpf}') 
             #--------------------------------------------------
            can.setFont("IBMPlexSans-Bold", 10)
            can.drawString(145, 501.5, f'{n_superInt}') 

            can.setFont("IBMPlexSans-Bold", 10)
            can.drawString(151, 519 , f'{nomeTecRep}')
            if incluir_funcao:
                can.setFont("IBMPlexSans-Bold",10)
                can.drawString(85, 599, f'{funcao}')
            can.save()

            packet.seek(0)
            new_pdf_data = packet.getvalue()
            new_page = PdfReader(io.BytesIO(new_pdf_data)).pages[0]
            existing_page.merge_page(new_page)

        output.add_page(existing_page)

    with open(unique_output_path, "wb") as output_file:
        output.write(output_file)

    print(f'OS Ordem Civil para {nome} preenchida e salva em {unique_output_path}.')
    return unique_output_path

def preencher_OS_obras_eletricas(nome, cpf, funcao, nomeTecRep, n_superInt, modelo_path, output_path, incluir_funcao=True, pagina_para_preencher=0):
    if 'Arial' not in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFont(TTFont('Arial', 'arial.ttf'))

    with open(modelo_path, "rb") as model_file:
        existing_pdf = PdfReader(io.BytesIO(model_file.read()))

    output_folder = os.path.join(r'C:\pdfBaixados', nome)

    # Criar diretório se não existir
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    base_name, extension = os.path.splitext(os.path.basename(output_path))
    unique_output_path = os.path.join(output_folder, f'{base_name}{extension}')

    output = PdfWriter()

    for page_num in range(len(existing_pdf.pages)):
        existing_page = existing_pdf.pages[page_num]

        if page_num == pagina_para_preencher:
            packet = io.BytesIO()
            can = canvas.Canvas(packet, pagesize=letter)
            can.setFont("IBMPlexSans-Bold", 10)
            can.drawString(115, 738, f'{nome}')
            if incluir_funcao:
                can.setFont("IBMPlexSans-Bold", 10)
                can.drawString(85, 721, f'{funcao}')
            can.save()

            packet.seek(0)
            new_pdf_data = packet.getvalue()
            new_page = PdfReader(io.BytesIO(new_pdf_data)).pages[0]
            existing_page.merge_page(new_page)

        elif page_num == 3:  # Página quarta    
            packet = io.BytesIO()
            can = canvas.Canvas(packet, pagesize=letter)
            can.setFont("IBMPlexSans-Bold", 10)
            can.drawString(113, 617, f'{nome}')  # Ajuste as coordenadas conforme necessário
            can.setFont("IBMPlexSans-Bold", 10)
            can.drawString(332, 600, f'{cpf}') 
             #--------------------------------------------------
            can.setFont("IBMPlexSans-Bold", 10)
            can.drawString(145, 501.5, f'{n_superInt}') 

            can.setFont("IBMPlexSans-Bold", 10)
            can.drawString(151, 519 , f'{nomeTecRep}')
            if incluir_funcao:
                can.setFont("IBMPlexSans-Bold",10)
                can.drawString(85, 599, f'{funcao}')
            can.save()

            packet.seek(0)
            new_pdf_data = packet.getvalue()
            new_page = PdfReader(io.BytesIO(new_pdf_data)).pages[0]
            existing_page.merge_page(new_page)

        output.add_page(existing_page)

    with open(unique_output_path, "wb") as output_file:
        output.write(output_file)

    print(f'OS Ordem Civil para {nome} preenchida e salva em {unique_output_path}.')
    return unique_output_path

def preencher_OS_obras_hidraulicas(nome, cpf, funcao, nomeTecRep, n_superInt, modelo_path, output_path, incluir_funcao=True, pagina_para_preencher=0):

    
    if 'Arial' not in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFont(TTFont('Arial', 'arial.ttf'))

    with open(modelo_path, "rb") as model_file:
        existing_pdf = PdfReader(io.BytesIO(model_file.read()))

    output_folder = os.path.join(r'C:\pdfBaixados', nome)

    # Criar diretório se não existir
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    base_name, extension = os.path.splitext(os.path.basename(output_path))
    unique_output_path = os.path.join(output_folder, f'{base_name}{extension}')

    output = PdfWriter()

    for page_num in range(len(existing_pdf.pages)):
        existing_page = existing_pdf.pages[page_num]

        if page_num == pagina_para_preencher:
            packet = io.BytesIO()
            can = canvas.Canvas(packet, pagesize=letter)
            can.setFont("IBMPlexSans-Bold", 10)
            can.drawString(115, 738, f'{nome}')
            if incluir_funcao:
                can.setFont("IBMPlexSans-Bold", 10)
                can.drawString(85, 721, f'{funcao}')
            can.save()

            packet.seek(0)
            new_pdf_data = packet.getvalue()
            new_page = PdfReader(io.BytesIO(new_pdf_data)).pages[0]
            existing_page.merge_page(new_page)

        elif page_num == 3:  # Página quarta    
            packet = io.BytesIO()
            can = canvas.Canvas(packet, pagesize=letter)
            can.setFont("IBMPlexSans-Bold", 10)
            can.drawString(113, 617, f'{nome}')  # Ajuste as coordenadas conforme necessário
            can.setFont("IBMPlexSans-Bold", 10)
            can.drawString(332, 600, f'{cpf}') 
             #--------------------------------------------------
            can.setFont("IBMPlexSans-Bold", 10)
            can.drawString(145, 501.5, f'{n_superInt}') 

            can.setFont("IBMPlexSans-Bold", 10)
            can.drawString(151, 519 , f'{nomeTecRep}')
            if incluir_funcao:
                can.setFont("IBMPlexSans-Bold",10)
                can.drawString(85, 599, f'{funcao}')
            can.save()

            packet.seek(0)
            new_pdf_data = packet.getvalue()
            new_page = PdfReader(io.BytesIO(new_pdf_data)).pages[0]
            existing_page.merge_page(new_page)

        output.add_page(existing_page)

    with open(unique_output_path, "wb") as output_file:
        output.write(output_file)

    print(f'OS Ordem Civil para {nome} preenchida e salva em {unique_output_path}.')
    return unique_output_path

def preencher_OS_soldador(nome, cpf, funcao, nomeTecRep, n_superInt, modelo_path, output_path, incluir_funcao=True, pagina_para_preencher=0):
    if 'Arial' not in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFont(TTFont('Arial', 'arial.ttf'))

    with open(modelo_path, "rb") as model_file:
        existing_pdf = PdfReader(io.BytesIO(model_file.read()))

    output_folder = os.path.join(r'C:\pdfBaixados', nome)

    # Criar diretório se não existir
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    base_name, extension = os.path.splitext(os.path.basename(output_path))
    unique_output_path = os.path.join(output_folder, f'{base_name}{extension}')

    output = PdfWriter()

    for page_num in range(len(existing_pdf.pages)):
        existing_page = existing_pdf.pages[page_num]

        if page_num == pagina_para_preencher:
            packet = io.BytesIO()
            can = canvas.Canvas(packet, pagesize=letter)
            can.setFont("IBMPlexSans-Bold", 10)
            can.drawString(115, 738, f'{nome}')
            if incluir_funcao:
                can.setFont("IBMPlexSans-Bold", 10)
                can.drawString(85, 721, f'{funcao}')
            can.save()

            packet.seek(0)
            new_pdf_data = packet.getvalue()
            new_page = PdfReader(io.BytesIO(new_pdf_data)).pages[0]
            existing_page.merge_page(new_page)

        elif page_num == 3:  # Página quarta    
            packet = io.BytesIO()
            can = canvas.Canvas(packet, pagesize=letter)
            can.setFont("IBMPlexSans-Bold", 10)
            can.drawString(113, 617, f'{nome}')  # Ajuste as coordenadas conforme necessário
            can.setFont("IBMPlexSans-Bold", 10)
            can.drawString(332, 600, f'{cpf}') 
             #--------------------------------------------------
            can.setFont("IBMPlexSans-Bold", 10)
            can.drawString(145, 501.5, f'{n_superInt}') 

            can.setFont("IBMPlexSans-Bold", 10)
            can.drawString(151, 519 , f'{nomeTecRep}')
            if incluir_funcao:
                can.setFont("IBMPlexSans-Bold",10)
                can.drawString(85, 599, f'{funcao}')
            can.save()

            packet.seek(0)
            new_pdf_data = packet.getvalue()
            new_page = PdfReader(io.BytesIO(new_pdf_data)).pages[0]
            existing_page.merge_page(new_page)

        output.add_page(existing_page)

    with open(unique_output_path, "wb") as output_file:
        output.write(output_file)

    print(f'OS Ordem Civil para {nome} preenchida e salva em {unique_output_path}.')
    return unique_output_path

def preencher_prova_NR06(nome, funcao, dataNR06, modelo_path, output_path, incluir_funcao=True):
    pdfmetrics.registerFont(TTFont('Arial-Bold', 'arialbd.ttf'))

    if 'Arial' not in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFont(TTFont('Arial', 'arial.ttf'))

    packet = io.BytesIO()
    can = canvas.Canvas(packet, pagesize=(800, 1200))

    # Tenta formatar a data e lidar com erros
    try:
        data_formatada = formatar_data_prova(dataNR06)
        partes_data = data_formatada.split()
        dia = partes_data[0]
        mes = partes_data[1]
        ano = partes_data[2]
    except Exception as e:
        print(f"Erro ao formatar a data: {e}")
        dia = mes = ano = "n/a"

    can.setFont("IBMPlexSans-Bold", 12)
    can.drawString(165, 710, f'{nome}')

    can.drawString(115, 688, dia)
    can.drawString(148, 688, mes)
    can.drawString(178, 688, ano)

    if incluir_funcao:
        can.setFont("IBMPlexSans-Bold", 8)
        can.drawString(74, 688, f'{funcao}')

    can.save()

    packet.seek(0)
    new_pdf_data = packet.getvalue()

    with open(modelo_path, "rb") as model_file:
        existing_pdf = PdfReader(io.BytesIO(model_file.read()))

    output_folder = os.path.join(r'C:\pdfBaixados', nome)

    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    base_name, extension = os.path.splitext(os.path.basename(output_path))
    unique_output_path = os.path.join(output_folder, f'{base_name}{extension}')

    output = PdfWriter()
    page = existing_pdf.pages[0]
    page.merge_page(PdfReader(io.BytesIO(new_pdf_data)).pages[0])
    output.add_page(page)

    with open(unique_output_path, "wb") as output_file:
        output.write(output_file)

    print(f'Prova da NR06 preenchido para {nome} preenchida e salva em {unique_output_path}.')
    return unique_output_path

def preencher_prova_NR10(nome, funcao, dataNR10, modelo_path, output_path, incluir_funcao=True):
    pdfmetrics.registerFont(TTFont('Arial-Bold', 'arialbd.ttf'))

    if 'Arial' not in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFont(TTFont('Arial', 'arial.ttf'))

    packet = io.BytesIO()
    can = canvas.Canvas(packet, pagesize=(800, 1200))

    # Garante que dataNR10 é uma string
    dataNR10 = str(dataNR10)

    # Chama a função para obter a última data
    data_formatada = obter_ultima_data(dataNR10)

    can.setFont("IBMPlexSans-Bold", 12)
    can.drawString(90, 734, f'{nome}')

    can.setFont("IBMPlexSans-Bold", 12)
    can.drawString(435, 738, f'{data_formatada}')

    if incluir_funcao:
        can.setFont("IBMPlexSans-Bold", 8)
        can.drawString(74, 688, f'{funcao}')

    can.save()

    packet.seek(0)
    new_pdf_data = packet.getvalue()

    with open(modelo_path, "rb") as model_file:
        existing_pdf = PdfReader(io.BytesIO(model_file.read()))

    output_folder = os.path.join(r'C:\pdfBaixados', nome)

    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    base_name, extension = os.path.splitext(os.path.basename(output_path))
    unique_output_path = os.path.join(output_folder, f'{base_name}{extension}')

    output = PdfWriter()
    page = existing_pdf.pages[0]
    page.merge_page(PdfReader(io.BytesIO(new_pdf_data)).pages[0])
    output.add_page(page)

    with open(unique_output_path, "wb") as output_file:
        output.write(output_file)

    print(f'Prova da NR06 preenchido para {nome} preenchida e salva em {unique_output_path}.')
    return unique_output_path

def preencher_prova_NR11(nome,funcao,dataNR06,modelo_path, output_path, incluir_funcao=True):
    pdfmetrics.registerFont(TTFont('Arial-Bold', 'arialbd.ttf'))

    if 'Arial' not in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFont(TTFont('Arial', 'arial.ttf'))

    packet = io.BytesIO()
    can = canvas.Canvas(packet, pagesize=(800, 1200))
    try:
        # Chama a função para formatar a data
        data_formatada = formatar_data_prova(dataNR06)
        partes_data = data_formatada.split()
        dia = partes_data[0]
        mes = partes_data[1]
        ano = partes_data[2]
    except Exception as e:
        print(f"Erro ao formatar a data: {e}")
        dia = mes = ano = "n/a"

    can.setFont("IBMPlexSans-Bold", 12)
    can.drawString(165,740, f'{nome}')

    try:
        can.drawString(115,718, dia)
        can.drawString(148, 718, mes) 
        can.drawString(178, 718, ano)
    except Exception as e:
        print(f"Erro ao formatar a data: {e}")
        dia = mes = ano = "n/a"

    if incluir_funcao:
        can.setFont("IBMPlexSans-Bold",8)
        can.drawString(74,688 , f'{funcao}')

    can.save()

    packet.seek(0)
    new_pdf_data = packet.getvalue()

    with open(modelo_path, "rb") as model_file:
        existing_pdf = PdfReader(io.BytesIO(model_file.read()))

    output_folder = os.path.join(r'C:\pdfBaixados', nome)
    
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    base_name, extension = os.path.splitext(os.path.basename(output_path))
    unique_output_path = os.path.join(output_folder, f'{base_name}{extension}')

    output = PdfWriter()
    page = existing_pdf.pages[0]
    page.merge_page(PdfReader(io.BytesIO(new_pdf_data)).pages[0])
    output.add_page(page)

    with open(unique_output_path, "wb") as output_file:
        output.write(output_file)

    print(f'Prova da NR11 preenchido para {nome} preenchida e salva em {unique_output_path}.')
    return unique_output_path

def preencher_prova_NR12(nome, funcao, dataNR12, modelo_path, output_path, incluir_funcao=True):
    pdfmetrics.registerFont(TTFont('Arial-Bold', 'arialbd.ttf'))

    if 'Arial' not in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFont(TTFont('Arial', 'arial.ttf'))

    packet = io.BytesIO()
    can = canvas.Canvas(packet, pagesize=(800, 1200))

    # Tenta formatar a data e lidar com erros
    try:
        data_formatada = formatar_data_prova(dataNR12)
        partes_data = data_formatada.split()
        dia = partes_data[0]
        mes = partes_data[1]
        ano = partes_data[2]
    except Exception as e:
        print(f"Erro ao formatar a data: {e}")
        dia = mes = ano = "n/a"

    can.setFont("IBMPlexSans-Bold", 12)
    can.drawString(165, 740, f'{nome}')

    can.drawString(115, 718, dia)
    can.drawString(148, 718, mes)
    can.drawString(178, 718, ano)

    if incluir_funcao:
        can.setFont("IBMPlexSans-Bold", 8)
        can.drawString(74, 688, f'{funcao}')

    can.save()

    packet.seek(0)
    new_pdf_data = packet.getvalue()

    with open(modelo_path, "rb") as model_file:
        existing_pdf = PdfReader(io.BytesIO(model_file.read()))

    output_folder = os.path.join(r'C:\pdfBaixados', nome)

    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    base_name, extension = os.path.splitext(os.path.basename(output_path))
    unique_output_path = os.path.join(output_folder, f'{base_name}{extension}')

    output = PdfWriter()
    page = existing_pdf.pages[0]
    page.merge_page(PdfReader(io.BytesIO(new_pdf_data)).pages[0])
    output.add_page(page)

    with open(unique_output_path, "wb") as output_file:
        output.write(output_file)

    print(f'Prova da NR12 preenchido para {nome} preenchida e salva em {unique_output_path}.')
    return unique_output_path


def preencher_prova_NR17(nome,funcao,dataNR17,modelo_path, output_path, incluir_funcao=True):
    pdfmetrics.registerFont(TTFont('Arial-Bold', 'arialbd.ttf'))

    if 'Arial' not in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFont(TTFont('Arial', 'arial.ttf'))

    packet = io.BytesIO()
    can = canvas.Canvas(packet, pagesize=(800, 1200))

    # Chama a função para formatar a data
    try:
        data_formatada = formatar_data_prova(dataNR17)
        partes_data = data_formatada.split()
        dia = partes_data[0]
        mes = partes_data[1]
        ano = partes_data[2]
    except Exception as e:
        print(f"Erro ao formatar a data: {e}")
        dia = mes = ano = "n/a"

    can.setFont("IBMPlexSans-Bold", 12)
    can.drawString(165,740, f'{nome}')

    can.drawString(115,718, dia)
    can.drawString(148, 718, mes) 
    can.drawString(178, 718, ano)


    if incluir_funcao:
        can.setFont("IBMPlexSans-Bold",8)
        can.drawString(74,688 , f'{funcao}')

    can.save()

    packet.seek(0)
    new_pdf_data = packet.getvalue()

    with open(modelo_path, "rb") as model_file:
        existing_pdf = PdfReader(io.BytesIO(model_file.read()))

    output_folder = os.path.join(r'C:\pdfBaixados', nome)
    
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    base_name, extension = os.path.splitext(os.path.basename(output_path))
    unique_output_path = os.path.join(output_folder, f'{base_name}{extension}')

    output = PdfWriter()
    page = existing_pdf.pages[0]
    page.merge_page(PdfReader(io.BytesIO(new_pdf_data)).pages[0])
    output.add_page(page)

    with open(unique_output_path, "wb") as output_file:
        output.write(output_file)

    print(f'Prova da NR11 preenchido para {nome} preenchida e salva em {unique_output_path}.')
    return unique_output_path

def preencher_prova_NR18(nome, funcao, dataNR18, modelo_path, output_path, incluir_funcao=True):
    pdfmetrics.registerFont(TTFont('Arial-Bold', 'arialbd.ttf'))

    if 'Arial' not in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFont(TTFont('Arial', 'arial.ttf'))

    packet = io.BytesIO()
    can = canvas.Canvas(packet, pagesize=(800, 1200))

    # Garante que dataNR10 é uma string
    dataNR18 = str(dataNR18)

    # Chama a função para obter a última data
    data_formatada = obter_ultima_data(dataNR18)

    can.setFont("IBMPlexSans-Bold", 12)
    can.drawString(90, 734, f'{nome}')

    can.setFont("IBMPlexSans-Bold", 12)
    can.drawString(435, 738, f'{data_formatada}')

    if incluir_funcao:
        can.setFont("IBMPlexSans-Bold", 8)
        can.drawString(74, 688, f'{funcao}')

    can.save()

    packet.seek(0)
    new_pdf_data = packet.getvalue()

    with open(modelo_path, "rb") as model_file:
        existing_pdf = PdfReader(io.BytesIO(model_file.read()))

    output_folder = os.path.join(r'C:\pdfBaixados', nome)

    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    base_name, extension = os.path.splitext(os.path.basename(output_path))
    unique_output_path = os.path.join(output_folder, f'{base_name}{extension}')

    output = PdfWriter()
    page = existing_pdf.pages[0]
    page.merge_page(PdfReader(io.BytesIO(new_pdf_data)).pages[0])
    output.add_page(page)

    with open(unique_output_path, "wb") as output_file:
        output.write(output_file)

    print(f'Prova da NR18 preenchido para {nome} preenchida e salva em {unique_output_path}.')
    return unique_output_path

def preencher_prova_NR18pta(nome,funcao,dataNR18_pemt,modelo_path, output_path, incluir_funcao=True):
    pdfmetrics.registerFont(TTFont('Arial-Bold', 'arialbd.ttf'))

    if 'Arial' not in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFont(TTFont('Arial', 'arial.ttf'))

    packet = io.BytesIO()
    can = canvas.Canvas(packet, pagesize=(800, 1200))

    # Chama a função para formatar a data
    try:
        data_formatada = formatar_data_prova(dataNR18_pemt)
        partes_data = data_formatada.split()
        dia = partes_data[0]
        mes = partes_data[1]
        ano = partes_data[2]
    except Exception as e:
        print(f"Erro ao formatar a data: {e}")
        dia = mes = ano = "n/a"

    can.setFont("IBMPlexSans-Bold", 12)
    can.drawString(165,740, f'{nome}')

    can.drawString(115,718, dia)
    can.drawString(148, 718, mes) 
    can.drawString(178, 718, ano)


    if incluir_funcao:
        can.setFont("IBMPlexSans-Bold",8)
        can.drawString(74,688 , f'{funcao}')

    can.save()

    packet.seek(0)
    new_pdf_data = packet.getvalue()

    with open(modelo_path, "rb") as model_file:
        existing_pdf = PdfReader(io.BytesIO(model_file.read()))

    output_folder = os.path.join(r'C:\pdfBaixados', nome)
    
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    base_name, extension = os.path.splitext(os.path.basename(output_path))
    unique_output_path = os.path.join(output_folder, f'{base_name}{extension}')

    output = PdfWriter()
    page = existing_pdf.pages[0]
    page.merge_page(PdfReader(io.BytesIO(new_pdf_data)).pages[0])
    output.add_page(page)

    with open(unique_output_path, "wb") as output_file:
        output.write(output_file)

    print(f'Prova da NR18 preenchido para {nome} preenchida e salva em {unique_output_path}.')
    return unique_output_path

def preencher_prova_NR20_infla(nome,funcao,dataNR20_inflamaveis,modelo_path, output_path, incluir_funcao=True):
    pdfmetrics.registerFont(TTFont('Arial-Bold', 'arialbd.ttf'))

    if 'Arial' not in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFont(TTFont('Arial', 'arial.ttf'))

    packet = io.BytesIO()
    can = canvas.Canvas(packet, pagesize=(800, 1200))

    # Chama a função para formatar a data
    try:
        data_formatada = formatar_data_prova(dataNR20_inflamaveis)
        partes_data = data_formatada.split()
        dia = partes_data[0]
        mes = partes_data[1]
        ano = partes_data[2]
    except Exception as e:
        print(f"Erro ao formatar a data: {e}")
        dia = mes = ano = "n/a"

    can.setFont("IBMPlexSans-Bold", 12)
    can.drawString(165,740, f'{nome}')

    can.drawString(115,718, dia)
    can.drawString(148, 718, mes) 
    can.drawString(178, 718, ano)


    if incluir_funcao:
        can.setFont("IBMPlexSans-Bold",8)
        can.drawString(74,688 , f'{funcao}')

    can.save()

    packet.seek(0)
    new_pdf_data = packet.getvalue()

    with open(modelo_path, "rb") as model_file:
        existing_pdf = PdfReader(io.BytesIO(model_file.read()))

    output_folder = os.path.join(r'C:\pdfBaixados', nome)
    
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    base_name, extension = os.path.splitext(os.path.basename(output_path))
    unique_output_path = os.path.join(output_folder, f'{base_name}{extension}')

    output = PdfWriter()
    page = existing_pdf.pages[0]
    page.merge_page(PdfReader(io.BytesIO(new_pdf_data)).pages[0])
    output.add_page(page)

    with open(unique_output_path, "wb") as output_file:
        output.write(output_file)

    print(f'Prova da NR20 inflamaveis preenchido para {nome} preenchida e salva em {unique_output_path}.')
    return unique_output_path

def preencher_prova_NR33(nome, funcao, dataNR33, modelo_path, output_path, incluir_funcao=True):
    pdfmetrics.registerFont(TTFont('Arial-Bold', 'arialbd.ttf'))

    if 'Arial' not in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFont(TTFont('Arial', 'arial.ttf'))

    packet = io.BytesIO()
    can = canvas.Canvas(packet, pagesize=(800, 1200))

    # Garante que dataNR10 é uma string
    dataNR33 = str(dataNR33)

    # Chama a função para obter a última data
    data_formatada = obter_ultima_data(dataNR33)

    can.setFont("IBMPlexSans-Bold", 12)
    can.drawString(90, 734, f'{nome}')

    can.setFont("IBMPlexSans-Bold", 12)
    can.drawString(435, 736.5, f'{data_formatada}')

    if incluir_funcao:
        can.setFont("IBMPlexSans-Bold", 8)
        can.drawString(74, 688, f'{funcao}')

    can.save()

    packet.seek(0)
    new_pdf_data = packet.getvalue()

    with open(modelo_path, "rb") as model_file:
        existing_pdf = PdfReader(io.BytesIO(model_file.read()))

    output_folder = os.path.join(r'C:\pdfBaixados', nome)

    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    base_name, extension = os.path.splitext(os.path.basename(output_path))
    unique_output_path = os.path.join(output_folder, f'{base_name}{extension}')

    output = PdfWriter()
    page = existing_pdf.pages[0]
    page.merge_page(PdfReader(io.BytesIO(new_pdf_data)).pages[0])
    output.add_page(page)

    with open(unique_output_path, "wb") as output_file:
        output.write(output_file)

    print(f'Prova da NR33 preenchido para {nome} preenchida e salva em {unique_output_path}.')
    return unique_output_path


def preencher_prova_NR34(nome,funcao,dataNR06,modelo_path, output_path, incluir_funcao=True):
    pdfmetrics.registerFont(TTFont('Arial-Bold', 'arialbd.ttf'))

    if 'Arial' not in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFont(TTFont('Arial', 'arial.ttf'))

    packet = io.BytesIO()
    can = canvas.Canvas(packet, pagesize=(800, 1200))
    try:
        # Chama a função para formatar a data
        data_formatada = formatar_data_prova(dataNR06)
        partes_data = data_formatada.split()
        dia = partes_data[0]
        mes = partes_data[1]
        ano = partes_data[2]
    except Exception as e:
        print(f"erro ao formatar{e}")
        dia = mes = ano = "n/a"

    can.setFont("IBMPlexSans-Bold", 12)
    can.drawString(165,740, f'{nome}')

    can.drawString(115,718, dia)
    can.drawString(148, 718, mes) 
    can.drawString(178, 718, ano)


    if incluir_funcao:
        can.setFont("IBMPlexSans-Bold",8)
        can.drawString(74,688 , f'{funcao}')

    can.save()

    packet.seek(0)
    new_pdf_data = packet.getvalue()

    with open(modelo_path, "rb") as model_file:
        existing_pdf = PdfReader(io.BytesIO(model_file.read()))

    output_folder = os.path.join(r'C:\pdfBaixados', nome)
    
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    base_name, extension = os.path.splitext(os.path.basename(output_path))
    unique_output_path = os.path.join(output_folder, f'{base_name}{extension}')

    output = PdfWriter()
    page = existing_pdf.pages[0]
    page.merge_page(PdfReader(io.BytesIO(new_pdf_data)).pages[0])
    output.add_page(page)

    with open(unique_output_path, "wb") as output_file:
        output.write(output_file)

    print(f'Prova da NR34 preenchido para {nome} preenchida e salva em {unique_output_path}.')
    return unique_output_path

def preencher_prova_NR35(nome, funcao, dataNR35, modelo_path, output_path, incluir_funcao=True):
    pdfmetrics.registerFont(TTFont('Arial-Bold', 'arialbd.ttf'))

    if 'Arial' not in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFont(TTFont('Arial', 'arial.ttf'))

    packet = io.BytesIO()
    can = canvas.Canvas(packet, pagesize=(800, 1200))

    # Garante que dataNR10 é uma string
    dataNR35 = str(dataNR35)

    # Chama a função para obter a última data
    data_formatada = obter_ultima_data(dataNR35)

    can.setFont("IBMPlexSans-Bold", 12)
    can.drawString(90, 737, f'{nome}')

    can.setFont("IBMPlexSans-Bold", 12)
    can.drawString(440, 742.5, f'{data_formatada}')

    if incluir_funcao:
        can.setFont("IBMPlexSans-Bold", 8)
        can.drawString(74, 688, f'{funcao}')

    can.save()

    packet.seek(0)
    new_pdf_data = packet.getvalue()

    with open(modelo_path, "rb") as model_file:
        existing_pdf = PdfReader(io.BytesIO(model_file.read()))

    output_folder = os.path.join(r'C:\pdfBaixados', nome)

    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    base_name, extension = os.path.splitext(os.path.basename(output_path))
    unique_output_path = os.path.join(output_folder, f'{base_name}{extension}')

    output = PdfWriter()
    page = existing_pdf.pages[0]
    page.merge_page(PdfReader(io.BytesIO(new_pdf_data)).pages[0])
    output.add_page(page)

    with open(unique_output_path, "wb") as output_file:
        output.write(output_file)

    print(f'Prova da NR35 preenchido para {nome} preenchida e salva em {unique_output_path}.')
    return unique_output_path


def preencher_treino_NR01(nome,nomeTST, funcao, dataNR01, modelo_path, output_path, incluir_funcao=False):
    pdfmetrics.registerFont(TTFont('Arial-Bold', 'arialbd.ttf'))

    if 'Arial' not in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFont(TTFont('Arial', 'arial.ttf'))

    packet = io.BytesIO()
    can = canvas.Canvas(packet, pagesize=(800, 1200))

    can.setFont("IBMPlexSans-Bold", 10)
    can.drawString(129, 618.5, f'{nomeTST}')

    can.setFont("IBMPlexSans-Bold", 9)
    can.drawString(158, 589 , f'{dataNR01}')

    if incluir_funcao:
        can.setFont("IBMPlexSans-Bold", 13)
        can.drawString(110,737 , f'{funcao}')


    can.save()

    packet.seek(0)
    new_pdf_data = packet.getvalue()

    with open(modelo_path, "rb") as model_file:
        existing_pdf = PdfReader(io.BytesIO(model_file.read()))

    output_folder = os.path.join(r'C:\pdfBaixados', nome)
    
    # Criar diretório se não existir
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    base_name, extension = os.path.splitext(os.path.basename(output_path))
    unique_output_path = os.path.join(output_folder, f'{base_name}{extension}')

    output = PdfWriter()
    page = existing_pdf.pages[0]
    page.merge_page(PdfReader(io.BytesIO(new_pdf_data)).pages[0])
    output.add_page(page)

    with open(unique_output_path, "wb") as output_file:
        output.write(output_file)

    print(f'RELATÓRIO DE TREINAMENTO - NR01 para {nome} preenchida e salva em {unique_output_path}.')
    return unique_output_path

def preencher_treino_NR06(nome,nomeTST, funcao, dataNR06, modelo_path, output_path, incluir_funcao=False):
    pdfmetrics.registerFont(TTFont('Arial-Bold', 'arialbd.ttf'))

    if 'Arial' not in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFont(TTFont('Arial', 'arial.ttf'))

    packet = io.BytesIO()
    can = canvas.Canvas(packet, pagesize=(800, 1200))

    can.setFont("IBMPlexSans-Bold", 10)
    can.drawString(132, 643, f'{nomeTST}')

    can.setFont("IBMPlexSans-Bold", 9)
    can.drawString(158, 613 , f'{dataNR06}')

    if incluir_funcao:
        can.setFont("IBMPlexSans-Bold", 13)
        can.drawString(110,737 , f'{funcao}')


    can.save()

    packet.seek(0)
    new_pdf_data = packet.getvalue()

    with open(modelo_path, "rb") as model_file:
        existing_pdf = PdfReader(io.BytesIO(model_file.read()))

    output_folder = os.path.join(r'C:\pdfBaixados', nome)
    
    # Criar diretório se não existir
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    base_name, extension = os.path.splitext(os.path.basename(output_path))
    unique_output_path = os.path.join(output_folder, f'{base_name}{extension}')

    output = PdfWriter()
    page = existing_pdf.pages[0]
    page.merge_page(PdfReader(io.BytesIO(new_pdf_data)).pages[0])
    output.add_page(page)

    with open(unique_output_path, "wb") as output_file:
        output.write(output_file)

    print(f'RELATÓRIO DE TREINAMENTO - NR06 para {nome} preenchida e salva em {unique_output_path}.')
    return unique_output_path

def preencher_treino_NR10(nome,nomeTST, funcao, dataNR10, modelo_path, output_path, incluir_funcao=False):
    pdfmetrics.registerFont(TTFont('Arial-Bold', 'arialbd.ttf'))

    if 'Arial' not in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFont(TTFont('Arial', 'arial.ttf'))

    packet = io.BytesIO()
    can = canvas.Canvas(packet, pagesize=(800, 1200))

    can.setFont("IBMPlexSans-Bold", 10)
    can.drawString(132, 500, f'{nomeTST}')

    can.setFont("IBMPlexSans-Bold", 9)
    can.drawString(158, 470 , f'{dataNR10}')

    if incluir_funcao:
        can.setFont("IBMPlexSans-Bold", 13)
        can.drawString(110,737 , f'{funcao}')


    can.save()

    packet.seek(0)
    new_pdf_data = packet.getvalue()

    with open(modelo_path, "rb") as model_file:
        existing_pdf = PdfReader(io.BytesIO(model_file.read()))

    output_folder = os.path.join(r'C:\pdfBaixados', nome)
    
    # Criar diretório se não existir
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    base_name, extension = os.path.splitext(os.path.basename(output_path))
    unique_output_path = os.path.join(output_folder, f'{base_name}{extension}')

    output = PdfWriter()
    page = existing_pdf.pages[0]
    page.merge_page(PdfReader(io.BytesIO(new_pdf_data)).pages[0])
    output.add_page(page)

    with open(unique_output_path, "wb") as output_file:
        output.write(output_file)

    print(f'RELATÓRIO DE TREINAMENTO - NR10 para {nome} preenchida e salva em {unique_output_path}.')
    return unique_output_path

def preencher_treino_NR10c(nome,nomeTST, funcao, dataNR10_complementar, modelo_path, output_path, incluir_funcao=False):
    pdfmetrics.registerFont(TTFont('Arial-Bold', 'arialbd.ttf'))

    if 'Arial' not in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFont(TTFont('Arial', 'arial.ttf'))

    packet = io.BytesIO()
    can = canvas.Canvas(packet, pagesize=(800, 1200))

    can.setFont("IBMPlexSans-Bold", 10)
    can.drawString(132, 488.5, f'{nomeTST}')

    can.setFont("IBMPlexSans-Bold", 9)
    can.drawString(158,462 , f'{dataNR10_complementar}')

    if incluir_funcao:
        can.setFont("IBMPlexSans-Bold", 13)
        can.drawString(110,737 , f'{funcao}')

    can.save()

    packet.seek(0)
    new_pdf_data = packet.getvalue()

    with open(modelo_path, "rb") as model_file:
        existing_pdf = PdfReader(io.BytesIO(model_file.read()))

    output_folder = os.path.join(r'C:\pdfBaixados', nome)
    
    # Criar diretório se não existir
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    base_name, extension = os.path.splitext(os.path.basename(output_path))
    unique_output_path = os.path.join(output_folder, f'{base_name}{extension}')

    output = PdfWriter()
    page = existing_pdf.pages[0]
    page.merge_page(PdfReader(io.BytesIO(new_pdf_data)).pages[0])
    output.add_page(page)

    with open(unique_output_path, "wb") as output_file:
        output.write(output_file)

    print(f'RELATÓRIO DE TREINAMENTO - NR10 Complementar para {nome} preenchida e salva em {unique_output_path}.')
    return unique_output_path

def preencher_treino_NR11(nome,nomeTST, funcao, dataNR11, modelo_path, output_path, incluir_funcao=False):
    pdfmetrics.registerFont(TTFont('Arial-Bold', 'arialbd.ttf'))

    if 'Arial' not in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFont(TTFont('Arial', 'arial.ttf'))

    packet = io.BytesIO()
    can = canvas.Canvas(packet, pagesize=(800, 1200))

    can.setFont("IBMPlexSans-Bold", 10)
    can.drawString(132, 605, f'{nomeTST}')

    can.setFont("IBMPlexSans-Bold", 9)
    can.drawString(158, 575 , f'{dataNR11}')

    if incluir_funcao:
        can.setFont("IBMPlexSans-Bold", 13)
        can.drawString(110,737 , f'{funcao}')

    can.save()

    packet.seek(0)
    new_pdf_data = packet.getvalue()

    with open(modelo_path, "rb") as model_file:
        existing_pdf = PdfReader(io.BytesIO(model_file.read()))

    output_folder = os.path.join(r'C:\pdfBaixados', nome)
    
    # Criar diretório se não existir
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    base_name, extension = os.path.splitext(os.path.basename(output_path))
    unique_output_path = os.path.join(output_folder, f'{base_name}{extension}')

    output = PdfWriter()
    page = existing_pdf.pages[0]
    page.merge_page(PdfReader(io.BytesIO(new_pdf_data)).pages[0])
    output.add_page(page)

    with open(unique_output_path, "wb") as output_file:
        output.write(output_file)

    print(f'RELATÓRIO DE TREINAMENTO - NR11 Complementar para {nome} preenchida e salva em {unique_output_path}.')
    return unique_output_path

def preencher_treino_NR12(nome,nomeTST, funcao, dataNR12, modelo_path, output_path, incluir_funcao=False):
    pdfmetrics.registerFont(TTFont('Arial-Bold', 'arialbd.ttf'))

    if 'Arial' not in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFont(TTFont('Arial', 'arial.ttf'))

    packet = io.BytesIO()
    can = canvas.Canvas(packet, pagesize=(800, 1200))

    can.setFont("IBMPlexSans-Bold", 10)
    can.drawString(132, 535, f'{nomeTST}')

    can.setFont("IBMPlexSans-Bold", 9)
    can.drawString(158, 508 , f'{dataNR12}')

    if incluir_funcao:
        can.setFont("IBMPlexSans-Bold", 13)
        can.drawString(110,737 , f'{funcao}')

    can.save()

    packet.seek(0)
    new_pdf_data = packet.getvalue()

    with open(modelo_path, "rb") as model_file:
        existing_pdf = PdfReader(io.BytesIO(model_file.read()))

    output_folder = os.path.join(r'C:\pdfBaixados', nome)
    
    # Criar diretório se não existir
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    base_name, extension = os.path.splitext(os.path.basename(output_path))
    unique_output_path = os.path.join(output_folder, f'{base_name}{extension}')

    output = PdfWriter()
    page = existing_pdf.pages[0]
    page.merge_page(PdfReader(io.BytesIO(new_pdf_data)).pages[0])
    output.add_page(page)

    with open(unique_output_path, "wb") as output_file:
        output.write(output_file)

    print(f'RELATÓRIO DE TREINAMENTO - NR12 Complementar para {nome} preenchida e salva em {unique_output_path}.')
    return unique_output_path

def preencher_treino_NR17(nome,nomeTST, funcao, dataNR17, modelo_path, output_path, incluir_funcao=False):
    pdfmetrics.registerFont(TTFont('Arial-Bold', 'arialbd.ttf'))

    if 'Arial' not in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFont(TTFont('Arial', 'arial.ttf'))

    packet = io.BytesIO()
    can = canvas.Canvas(packet, pagesize=(800, 1200))

    can.setFont("IBMPlexSans-Bold", 10)
    can.drawString(129, 662, f'{nomeTST}')

    can.setFont("IBMPlexSans-Bold", 9)
    can.drawString(158, 632 , f'{dataNR17}')

    if incluir_funcao:
        can.setFont("IBMPlexSans-Bold", 13)
        can.drawString(110,737 , f'{funcao}')

    can.save()

    packet.seek(0)
    new_pdf_data = packet.getvalue()

    with open(modelo_path, "rb") as model_file:
        existing_pdf = PdfReader(io.BytesIO(model_file.read()))

    output_folder = os.path.join(r'C:\pdfBaixados', nome)
    
    # Criar diretório se não existir
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    base_name, extension = os.path.splitext(os.path.basename(output_path))
    unique_output_path = os.path.join(output_folder, f'{base_name}{extension}')

    output = PdfWriter()
    page = existing_pdf.pages[0]
    page.merge_page(PdfReader(io.BytesIO(new_pdf_data)).pages[0])
    output.add_page(page)

    with open(unique_output_path, "wb") as output_file:
        output.write(output_file)

    print(f'RELATÓRIO DE TREINAMENTO - NR17 para {nome} preenchida e salva em {unique_output_path}.')
    return unique_output_path

def preencher_treino_NR18civil(nome,nomeTST, funcao, dataNR18, modelo_path, output_path, incluir_funcao=False):
    pdfmetrics.registerFont(TTFont('Arial-Bold', 'arialbd.ttf'))

    if 'Arial' not in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFont(TTFont('Arial', 'arial.ttf'))

    packet = io.BytesIO()
    can = canvas.Canvas(packet, pagesize=(800, 1200))

    can.setFont("IBMPlexSans-Bold", 10)
    can.drawString(129, 662, f'{nomeTST}')

    can.setFont("IBMPlexSans-Bold", 9)
    can.drawString(158, 632 , f'{dataNR18}')

    if incluir_funcao:
        can.setFont("IBMPlexSans-Bold", 13)
        can.drawString(110,737 , f'{funcao}')

    can.save()

    packet.seek(0)
    new_pdf_data = packet.getvalue()

    with open(modelo_path, "rb") as model_file:
        existing_pdf = PdfReader(io.BytesIO(model_file.read()))

    output_folder = os.path.join(r'C:\pdfBaixados', nome)
    
    # Criar diretório se não existir
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    base_name, extension = os.path.splitext(os.path.basename(output_path))
    unique_output_path = os.path.join(output_folder, f'{base_name}{extension}')

    output = PdfWriter()
    page = existing_pdf.pages[0]
    page.merge_page(PdfReader(io.BytesIO(new_pdf_data)).pages[0])
    output.add_page(page)

    with open(unique_output_path, "wb") as output_file:
        output.write(output_file)

    print(f'RELATÓRIO DE TREINAMENTO - NR18 basica para {nome} preenchida e salva em {unique_output_path}.')
    return unique_output_path

def preencher_treino_NR18pta(nome,nomeTST, funcao, dataNR18_pemt, modelo_path, output_path, incluir_funcao=False):
    pdfmetrics.registerFont(TTFont('Arial-Bold', 'arialbd.ttf'))

    if 'Arial' not in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFont(TTFont('Arial', 'arial.ttf'))

    packet = io.BytesIO()
    can = canvas.Canvas(packet, pagesize=(800, 1200))

    can.setFont("IBMPlexSans-Bold", 10)
    can.drawString(130, 593, f'{nomeTST}')

    can.setFont("IBMPlexSans-Bold", 9)
    can.drawString(158, 565 , f'{dataNR18_pemt}')

    if incluir_funcao:
        can.setFont("IBMPlexSans-Bold", 13)
        can.drawString(110,737 , f'{funcao}')

    can.save()

    packet.seek(0)
    new_pdf_data = packet.getvalue()

    with open(modelo_path, "rb") as model_file:
        existing_pdf = PdfReader(io.BytesIO(model_file.read()))

    output_folder = os.path.join(r'C:\pdfBaixados', nome)
    
    # Criar diretório se não existir
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    base_name, extension = os.path.splitext(os.path.basename(output_path))
    unique_output_path = os.path.join(output_folder, f'{base_name}{extension}')

    output = PdfWriter()
    page = existing_pdf.pages[0]
    page.merge_page(PdfReader(io.BytesIO(new_pdf_data)).pages[0])
    output.add_page(page)

    with open(unique_output_path, "wb") as output_file:
        output.write(output_file)

    print(f'RELATÓRIO DE TREINAMENTO - NR18 pta para {nome} preenchida e salva em {unique_output_path}.')
    return unique_output_path

def preencher_treino_NR20(nome,nomeTST, funcao, dataNR20, modelo_path, output_path, incluir_funcao=False):
    pdfmetrics.registerFont(TTFont('Arial-Bold', 'arialbd.ttf'))

    if 'Arial' not in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFont(TTFont('Arial', 'arial.ttf'))

    packet = io.BytesIO()
    can = canvas.Canvas(packet, pagesize=(800, 1200))

    can.setFont("IBMPlexSans-Bold", 10)
    can.drawString(129, 650, f'{nomeTST}')

    can.setFont("IBMPlexSans-Bold", 9)
    can.drawString(158, 621 , f'{dataNR20}')

    if incluir_funcao:
        can.setFont("IBMPlexSans-Bold", 13)
        can.drawString(110,737 , f'{funcao}')

    can.save()

    packet.seek(0)
    new_pdf_data = packet.getvalue()

    with open(modelo_path, "rb") as model_file:
        existing_pdf = PdfReader(io.BytesIO(model_file.read()))

    output_folder = os.path.join(r'C:\pdfBaixados', nome)
    
    # Criar diretório se não existir
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    base_name, extension = os.path.splitext(os.path.basename(output_path))
    unique_output_path = os.path.join(output_folder, f'{base_name}{extension}')

    output = PdfWriter()
    page = existing_pdf.pages[0]
    page.merge_page(PdfReader(io.BytesIO(new_pdf_data)).pages[0])
    output.add_page(page)

    with open(unique_output_path, "wb") as output_file:
        output.write(output_file)

    print(f'RELATÓRIO DE TREINAMENTO - NR20 para {nome} preenchida e salva em {unique_output_path}.')
    return unique_output_path

def preencher_treino_NR33(nome,nomeTST, funcao, dataNR33, modelo_path, output_path, incluir_funcao=False):
    pdfmetrics.registerFont(TTFont('Arial-Bold', 'arialbd.ttf'))

    if 'Arial' not in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFont(TTFont('Arial', 'arial.ttf'))

    packet = io.BytesIO()
    can = canvas.Canvas(packet, pagesize=(800, 1200))

    can.setFont("IBMPlexSans-Bold", 10)
    can.drawString(129, 628, f'{nomeTST}')

    can.setFont("IBMPlexSans-Bold", 9)
    can.drawString(158, 597 , f'{dataNR33}')

    if incluir_funcao:
        can.setFont("IBMPlexSans-Bold", 13)
        can.drawString(110,737 , f'{funcao}')

    can.save()

    packet.seek(0)
    new_pdf_data = packet.getvalue()

    with open(modelo_path, "rb") as model_file:
        existing_pdf = PdfReader(io.BytesIO(model_file.read()))

    output_folder = os.path.join(r'C:\pdfBaixados', nome)
    
    # Criar diretório se não existir
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    base_name, extension = os.path.splitext(os.path.basename(output_path))
    unique_output_path = os.path.join(output_folder, f'{base_name}{extension}')

    output = PdfWriter()
    page = existing_pdf.pages[0]
    page.merge_page(PdfReader(io.BytesIO(new_pdf_data)).pages[0])
    output.add_page(page)

    with open(unique_output_path, "wb") as output_file:
        output.write(output_file)

    print(f'RELATÓRIO DE TREINAMENTO - NR33 para {nome} preenchida e salva em {unique_output_path}.')
    return unique_output_path

def preencher_treino_NR34adm(nome,nomeTST, funcao, dataNR34_adm, modelo_path, output_path, incluir_funcao=False):
    pdfmetrics.registerFont(TTFont('Arial-Bold', 'arialbd.ttf'))

    if 'Arial' not in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFont(TTFont('Arial', 'arial.ttf'))

    packet = io.BytesIO()
    can = canvas.Canvas(packet, pagesize=(800, 1200))

    can.setFont("IBMPlexSans-Bold", 10)
    can.drawString(129, 650, f'{nomeTST}')

    can.setFont("IBMPlexSans-Bold", 9)
    can.drawString(158, 620 , f'{dataNR34_adm}')

    if incluir_funcao:
        can.setFont("IBMPlexSans-Bold", 13)
        can.drawString(110,737 , f'{funcao}')

    can.save()

    packet.seek(0)
    new_pdf_data = packet.getvalue()

    with open(modelo_path, "rb") as model_file:
        existing_pdf = PdfReader(io.BytesIO(model_file.read()))

    output_folder = os.path.join(r'C:\pdfBaixados', nome)
    
    # Criar diretório se não existir
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    base_name, extension = os.path.splitext(os.path.basename(output_path))
    unique_output_path = os.path.join(output_folder, f'{base_name}{extension}')

    output = PdfWriter()
    page = existing_pdf.pages[0]
    page.merge_page(PdfReader(io.BytesIO(new_pdf_data)).pages[0])
    output.add_page(page)

    with open(unique_output_path, "wb") as output_file:
        output.write(output_file)

    print(f'RELATÓRIO DE TREINAMENTO - NR34 ADM para {nome} preenchida e salva em {unique_output_path}.')
    return unique_output_path

def preencher_treino_NR34obs(nome,nomeTST, funcao, dataNR34_obs_quente, modelo_path, output_path, incluir_funcao=False):
    pdfmetrics.registerFont(TTFont('Arial-Bold', 'arialbd.ttf'))

    if 'Arial' not in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFont(TTFont('Arial', 'arial.ttf'))

    packet = io.BytesIO()
    can = canvas.Canvas(packet, pagesize=(800, 1200))

    can.setFont("IBMPlexSans-Bold", 10)
    can.drawString(129, 661, f'{nomeTST}')

    can.setFont("IBMPlexSans-Bold", 9)
    can.drawString(158, 632 , f'{dataNR34_obs_quente}')

    if incluir_funcao:
        can.setFont("IBMPlexSans-Bold", 13)
        can.drawString(110,737 , f'{funcao}')

    can.save()

    packet.seek(0)
    new_pdf_data = packet.getvalue()

    with open(modelo_path, "rb") as model_file:
        existing_pdf = PdfReader(io.BytesIO(model_file.read()))

    output_folder = os.path.join(r'C:\pdfBaixados', nome)
    
    # Criar diretório se não existir
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    base_name, extension = os.path.splitext(os.path.basename(output_path))
    unique_output_path = os.path.join(output_folder, f'{base_name}{extension}')

    output = PdfWriter()
    page = existing_pdf.pages[0]
    page.merge_page(PdfReader(io.BytesIO(new_pdf_data)).pages[0])
    output.add_page(page)

    with open(unique_output_path, "wb") as output_file:
        output.write(output_file)

    print(f'RELATÓRIO DE TREINAMENTO - NR34 OBS  para {nome} preenchida e salva em {unique_output_path}.')
    return unique_output_path

def preencher_treino_NR34bas(nome,nomeTST, funcao, dataNR34, modelo_path, output_path, incluir_funcao=False):
    pdfmetrics.registerFont(TTFont('Arial-Bold', 'arialbd.ttf'))

    if 'Arial' not in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFont(TTFont('Arial', 'arial.ttf'))

    packet = io.BytesIO()
    can = canvas.Canvas(packet, pagesize=(800, 1200))

    can.setFont("IBMPlexSans-Bold", 10)
    can.drawString(129, 440, f'{nomeTST}')

    can.setFont("IBMPlexSans-Bold", 9)
    can.drawString(158, 411 , f'{dataNR34}')

    if incluir_funcao:
        can.setFont("IBMPlexSans-Bold", 13)
        can.drawString(110,737 , f'{funcao}')

    can.save()

    packet.seek(0)
    new_pdf_data = packet.getvalue()

    with open(modelo_path, "rb") as model_file:
        existing_pdf = PdfReader(io.BytesIO(model_file.read()))

    output_folder = os.path.join(r'C:\pdfBaixados', nome)
    
    # Criar diretório se não existir
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    base_name, extension = os.path.splitext(os.path.basename(output_path))
    unique_output_path = os.path.join(output_folder, f'{base_name}{extension}')

    output = PdfWriter()
    page = existing_pdf.pages[0]
    page.merge_page(PdfReader(io.BytesIO(new_pdf_data)).pages[0])
    output.add_page(page)

    with open(unique_output_path, "wb") as output_file:
        output.write(output_file)

    print(f'RELATÓRIO DE TREINAMENTO - NR34 basica  para {nome} preenchida e salva em {unique_output_path}.')
    return unique_output_path

def preencher_treino_NR35(nome,nomeTST, funcao, dataNR35, modelo_path, output_path, incluir_funcao=False):
    pdfmetrics.registerFont(TTFont('Arial-Bold', 'arialbd.ttf'))

    if 'Arial' not in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFont(TTFont('Arial', 'arial.ttf'))

    packet = io.BytesIO()
    can = canvas.Canvas(packet, pagesize=(800, 1200))

    can.setFont("IBMPlexSans-Bold", 10)
    can.drawString(129, 615, f'{nomeTST}')

    can.setFont("IBMPlexSans-Bold", 9)
    can.drawString(158, 586.5 , f'{dataNR35}')

    if incluir_funcao:
        can.setFont("IBMPlexSans-Bold", 13)
        can.drawString(110,737 , f'{funcao}')

    can.save()

    packet.seek(0)
    new_pdf_data = packet.getvalue()

    with open(modelo_path, "rb") as model_file:
        existing_pdf = PdfReader(io.BytesIO(model_file.read()))

    output_folder = os.path.join(r'C:\pdfBaixados', nome)
    
    # Criar diretório se não existir
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    base_name, extension = os.path.splitext(os.path.basename(output_path))
    unique_output_path = os.path.join(output_folder, f'{base_name}{extension}')

    output = PdfWriter()
    page = existing_pdf.pages[0]
    page.merge_page(PdfReader(io.BytesIO(new_pdf_data)).pages[0])
    output.add_page(page)

    with open(unique_output_path, "wb") as output_file:
        output.write(output_file)

    print(f'RELATÓRIO DE TREINAMENTO - NR18 pta para {nome} preenchida e salva em {unique_output_path}.')
    return unique_output_path



def preencher_cracha(nome,nome_obra,funcao,data_aso,dataNR06,dataNR05,dataNR18,dataNR35,dataNR12,dataNR01,dataNR10_basica,dataNR10_complementar,dataNR11,dataNR18_pemt,dataNR20_inflamaveis,dataNR20_brigada,dataNR33,dataNR34,dataNR34_adm,dataNR34_obs_quente,dataNR17,modelo_path, output_path, incluir_funcao=True):
    pdfmetrics.registerFont(TTFont('Arial-Bold', 'arialbd.ttf'))

    if 'Arial' not in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFont(TTFont('Arial', 'arial.ttf'))

    packet = io.BytesIO()
    can = canvas.Canvas(packet, pagesize=(800, 1200))

    data_aso_formatada = formatar_data_cracha(str(data_aso), adicionar_anos=1)
    
    dataNR05_formatada = formatar_data_cracha(str(dataNR05), adicionar_anos=2)
    dataNR06_formatada = formatar_data_cracha(str(dataNR06), adicionar_anos=2)
    dataNR18_formatada = formatar_data_cracha(str(dataNR18), adicionar_anos=2)
    dataNR35_formatada = formatar_data_cracha(str(dataNR35), adicionar_anos=2)
    dataNR12_formatada = formatar_data_cracha(str(dataNR12), adicionar_anos=2)
    dataNR01_formatada = formatar_data_cracha(str(dataNR01), adicionar_anos=2)
    dataNR10_basica_formatada = formatar_data_cracha(str(dataNR10_basica), adicionar_anos=2)
    dataNR10_complementar_formatada = formatar_data_cracha(str(dataNR10_complementar), adicionar_anos=2)
    dataNR11_formatada = formatar_data_cracha(str(dataNR11), adicionar_anos=2)
    dataNR18_pemt_formatada = formatar_data_cracha(str(dataNR18_pemt), adicionar_anos=2)
    dataNR20_inflamaveis_formatada = formatar_data_cracha(str(dataNR20_inflamaveis), adicionar_anos=2)
    dataNR20_brigada_formatada = formatar_data_cracha(str(dataNR20_brigada), adicionar_anos=2)
    dataNR34_formatada = formatar_data_cracha(str(dataNR34), adicionar_anos=2)
    dataNR34_adm_formatada = formatar_data_cracha(str(dataNR34_adm), adicionar_anos=2)
    dataNR34_obs_quente_formatada = formatar_data_cracha(str(dataNR34_obs_quente), adicionar_anos=2)
    dataNR17_formatada = formatar_data_cracha(str(dataNR17), adicionar_anos=2)


    can.setFont("IBMPlexSans-Bold", 8)
    can.drawString(67,698, f'{nome}')
#--------------------------------------------------
    parte1, parte2 = quebrar_nome_obra(nome_obra)
    can.setFont("IBMPlexSans-Bold", 8)
    can.drawString(45, 730, f'{parte1}')
    can.setFont("IBMPlexSans-Bold", 8)
    can.drawString(45, 722, f'{parte2}')
#--------------------------------------------------
    can.setFont("IBMPlexSans-Bold",8)
    can.drawString(120, 574, f'{data_aso_formatada}')

    can.setFont("IBMPlexSans-Bold", 6)
    can.drawString(77, 667, f'{dataNR01_formatada}')

    can.setFont("IBMPlexSans-Bold", 6)
    can.drawString(87, 634, f'{dataNR10_basica_formatada}')

    can.setFont("IBMPlexSans-Bold", 6)
    can.drawString(78, 656.5, f'{dataNR05_formatada}')

    can.setFont("IBMPlexSans-Bold", 6)
    can.drawString(78, 645.5, f'{dataNR06_formatada}')

    can.setFont("IBMPlexSans-Bold", 6)
    can.drawString(88, 623.5, f'{dataNR10_complementar_formatada}')
    
    can.setFont("IBMPlexSans-Bold", 6)
    can.drawString(78, 612.5, f'{dataNR11_formatada}')

    can.setFont("IBMPlexSans-Bold", 6)
    can.drawString(78.5, 602.5, f'{dataNR12_formatada}')

    can.setFont("IBMPlexSans-Bold", 6)
    can.drawString(78, 591.5, f'{dataNR17_formatada}')
#-----------------------------------------------------------------------
    
    can.setFont("IBMPlexSans-Bold", 6)
    can.drawString(165, 668, f'{dataNR18_formatada}')

    can.setFont("IBMPlexSans-Bold", 6)
    can.drawString(175, 657, f'{dataNR18_pemt_formatada}')

    can.setFont("IBMPlexSans-Bold", 6)
    can.drawString(165, 645.5, f'{dataNR20_inflamaveis_formatada}')

    can.setFont("IBMPlexSans-Bold", 6)
    can.drawString(165, 634, f'{dataNR20_brigada_formatada}')

    can.setFont("IBMPlexSans-Bold", 6)
    can.drawString(176, 623.5, f'{dataNR34_formatada}')

    can.setFont("IBMPlexSans-Bold", 6)
    can.drawString(178, 613, f'{dataNR34_adm_formatada}')

    can.setFont("IBMPlexSans-Bold", 6)
    can.drawString(174.5, 601.5, f'{dataNR34_obs_quente_formatada}')

    can.setFont("IBMPlexSans-Bold", 6)
    can.drawString(167,  591.5, f'{dataNR35_formatada}')
    

    if incluir_funcao:
        can.setFont("IBMPlexSans-Bold",8)
        can.drawString(74,688 , f'{funcao}')

    can.save()

    packet.seek(0)
    new_pdf_data = packet.getvalue()

    with open(modelo_path, "rb") as model_file:
        existing_pdf = PdfReader(io.BytesIO(model_file.read()))

    output_folder = os.path.join(r'C:\pdfBaixados', nome)
    
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    base_name, extension = os.path.splitext(os.path.basename(output_path))
    unique_output_path = os.path.join(output_folder, f'{base_name}{extension}')

    output = PdfWriter()
    page = existing_pdf.pages[0]
    page.merge_page(PdfReader(io.BytesIO(new_pdf_data)).pages[0])
    output.add_page(page)

    with open(unique_output_path, "wb") as output_file:
        output.write(output_file)

    print(f'Cracha Planem preenchido para {nome} preenchida e salva em {unique_output_path}.')
    return unique_output_path

def formatar_data_prova(data):
    if isinstance(data, float):
        return "N/A"  # Retorna "N/A" se a entrada for um número float
    elif not isinstance(data, str) or not data:
        return "N/A"  # Retorna "N/A" se a entrada não for uma string ou estiver vazia

    try:
        data = data.replace('a', 'A')

        # Verifica se há um intervalo de datas
        if '-' in data:
            # Divide o intervalo em datas individuais
            datas = data.split('-')
            # Pega a última data do intervalo, se houver
            if datas:
                data = datas[-1].strip()
            else:
                return "N/A"  # Retorna "N/A" se não houver nenhuma data no intervalo
        elif 'A' in data:
            # Se houver um intervalo de datas sem barras, pegue a última parte
            partes = data.split('A')
            if partes:
                data = partes[-1].strip()
            else:
                return "N/A"  # Retorna "N/A" se não houver nenhuma parte da data
        # Substitui as barras por espaços
        data_formatada = data.replace('/', ' ')
        partes_data = data_formatada.split()
        dia = partes_data[0] if len(partes_data) > 0 else "N/A"
        mes = partes_data[1] if len(partes_data) > 1 else "N/A"
        ano = partes_data[2] if len(partes_data) > 2 else "N/A"
        return f"{dia} {mes} {ano}"
    except Exception as e:
        print(f"Erro ao formatar a data: {e}")
        return "N/A"
    
    



def obter_ultima_data(data):
    if isinstance(data, float):
        return "N/A"  # Retorna "N/A" se a entrada for um número float
    elif not isinstance(data, str) or not data:
        return "N/A"  # Retorna "N/A" se a entrada não for uma string ou estiver vazia

    try:
        partes = data.split(' ')
        if len(partes) > 0:
            ultima_data = partes[-1]
            ultima_data = ultima_data.replace("a", "").replace("A", "")
            return ultima_data
        else:
            return "N/A"  # Retorna "N/A" se não houver nenhuma parte da data
    except Exception as e:
        print(f"Erro ao obter a última data: {e}")
        return "N/A"




def quebrar_nome_obra(nome_obra):
    if '/' in nome_obra:
        partes = nome_obra.split('/', 1)
        parte1 = partes[0].strip() + '/'
        parte2 = partes[1].strip()
    else:
        parte1, parte2 = nome_obra.strip(), ''

    return parte1, parte2

def formatar_data_cracha(data, adicionar_anos=0):
    try:
        # Verifica se a data não está vazia
        if not data:
            return "N/A"

        # Verifica se a data está no formato de intervalo
        if ' a ' in data:
            # Divide o intervalo em datas de início e fim
            data_inicio, data_fim = map(lambda x: x.strip(), data.split(' a '))
            
            # Usa apenas a data de fim
            data_formatada = datetime.strptime(data_fim, "%d/%m/%Y")
        else:
            # Trata datas normais
            data_formatada = datetime.strptime(str(data), "%d/%m/%Y")
        
        # Adiciona anos à data formatada
        data_formatada = data_formatada.replace(year=data_formatada.year + adicionar_anos)
        
        return data_formatada.strftime("%d/%m/%Y")
    except ValueError:
        try:
            # Se não for uma data no formato padrão, tenta converter de outra forma (se aplicável)
            data_formatada = datetime.fromordinal(datetime(1900, 1, 1).toordinal() + int(data) - 2)
            
            # Converte para string e, em seguida, para objeto datetime
            data_formatada = datetime.strptime(str(data_formatada), "%Y-%m-%d %H:%M:%S")
            
            # Adiciona anos à data formatada
            data_formatada = data_formatada.replace(year=data_formatada.year + adicionar_anos)
            
            return data_formatada.strftime("%d/%m/%Y")
        except:
            return "N/A"


        
def formatar_data(data, formato='longa'):
    data_formatada = ""  # Inicializa a variável fora dos blocos condicionais

    if isinstance(data, (str, int, float)):
        try:
            # Verifica se a data contém uma barra indicando intervalo
            if ' a ' in str(data):
                datas = data.split(' a ')
                data_inicio = datas[0].strip()
                data_fim = datas[1].strip()
                
                data_inicio_formatada = ""
                if data_inicio:
                    if '/' in data_inicio:
                        data_inicio_formatada = datetime.strptime(data_inicio, "%d/%m/%Y").strftime("%d a %B")
                    else:
                        data_inicio_formatada = data_inicio

                data_fim_formatada = ""
                if data_fim:
                    if '/' in data_fim:
                        data_fim_formatada = datetime.strptime(data_fim, "%d/%m/%Y").strftime("%d de %B %Y")
                    else:
                        data_fim_formatada = data_fim

                data_formatada = f"{data_inicio_formatada}{' a ' if data_inicio_formatada and data_fim_formatada else ''}{'' if data_inicio_formatada and data_fim_formatada else 'de '}{data_fim_formatada}" if data_inicio_formatada or data_fim_formatada else ''
            else:
                # Trata datas normais
                data_formatada = datetime.strptime(str(data), "%d/%m/%Y").strftime("%d de %B de %Y")
        except ValueError:
            try:
                data_formatada = datetime.fromordinal(datetime(1900, 1, 1).toordinal() + int(data) - 2).strftime(
                    "%d de %B de %Y")
            except:
                data_formatada = str(data)
    else:
        if formato == 'curta':
            data_formatada = data.strftime("%d/%m/%Y")
        else:
            data_formatada = data.strftime("%d de %B de %Y")

    # Substituir meses após o tratamento específico
    meses_ingles = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']
    meses_portugues = ['janeiro', 'fevereiro', 'março', 'abril', 'maio', 'junho', 'julho', 'agosto', 'setembro', 'outubro', 'novembro', 'dezembro']

    for mes_ingles, mes_portugues in zip(meses_ingles, meses_portugues):
        data_formatada = data_formatada.replace(mes_ingles, mes_portugues)

    # Adicionar ponto final à data
    data_formatada += '.'
    
    return data_formatada