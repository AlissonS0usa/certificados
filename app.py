from flask import Flask, render_template, request, send_file
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.platypus import Table, TableStyle, Image as ReportLabImage
from reportlab.lib import colors
import io
import os
from reportlab.platypus import Table, TableStyle, Paragraph
from reportlab.lib.styles import getSampleStyleSheet
from datetime import datetime
from werkzeug.utils import secure_filename
from PIL import Image as PilImage
import random




app = Flask(__name__)

UPLOAD_FOLDER = 'static/uploads' # Define a pasta de uploads
os.makedirs(UPLOAD_FOLDER, exist_ok=True) # Cria a pasta de uploads se ela não existir
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER # Define a pasta de uploads como a pasta de uploads do app
app.config['MAXX_CONTENT_LENGTH'] = 16 * 1024 * 1024 # Define o tamanho máximo do arquivo como 16 MB

def gerar_valor_aleatorio(base):
    result = request.form["resultado"]
    if result == "APROVADA":
        min_val = (base - 0.15) * 1000
        max_val = (base + 0.15) * 1000
    else:   
        fator = random.uniform(0.97, 1.07)
        return round(base * fator, 2)
    
 
    return round(random.randint(int(min_val), int(max_val)) / 1000, 2)

# Função para desenhar a tabela em múltiplas páginas
def desenhar_tabela_dinamica(pdf, dados_tabela, y_posicao_p2, resultado):    
    y_posicao = y_posicao_p2  # Posição inicial na página
    max_linhas_por_pagina = 36  # Ajuste conforme necessário
    inicio_linha = 1

    
        
    # Cria a tabela para essas linhas
    t_dinamica = Table(dados_tabela, colWidths=[70, 100, 100, 100, 100, 70])
    t_dinamica.setStyle(TableStyle([
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
        ("FONTSIZE", (0, 0), (-1, -1), 10),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("GRID", (0, 0), (-1, -1), 1, colors.black)
    ]))
    
    while inicio_linha < len(dados_tabela):
        # Seleciona o conjunto de linhas para esta página
        fim_linha = min(inicio_linha + max_linhas_por_pagina, len(dados_tabela))
        dados_pagina = dados_tabela[inicio_linha:fim_linha]  # Ajuste feito aqui
        
        # Cria a tabela para essas linhas
        t_dinamica = Table(dados_pagina, colWidths=[70, 100, 100, 100, 100, 70])
        t_dinamica.setStyle(TableStyle([
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
            ("FONTSIZE", (0, 0), (-1, -1), 10),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
            ("TOPPADDING", (0, 0), (-1, -1), 2),
            ("GRID", (0, 0), (-1, -1), 1, colors.black)
        ]))
        
        # Calcula a altura da tabela para verificar se cabe na página
        largura, altura = t_dinamica.wrap(0, 0)
        
        # Se não couber na página atual, cria uma nova página
        if y_posicao - altura < 50:
            pdf.showPage()
            pdf.drawImage("static/cabecalho.png", 0, 0, width=A4[0], height=A4[1])
            y_posicao = y_posicao_p2 + 20  # Reinicia a posição para a nova página
        
        # Desenha a tabela na posição atual
        t_dinamica.drawOn(pdf, 28, y_posicao - altura)
        y_posicao -= (altura + 10)  # Ajusta a posição para a próxima tabela
        
        # Avança para as próximas linhas
        inicio_linha = fim_linha  # Ajuste feito aqui

    dados_nova_tabela = [["RESULTADO DA CALIBRAÇÃO", f"{resultado}"]]
    
    # Cria a nova tabela
    nova_tabela = Table(dados_nova_tabela, colWidths=[270, 270])
    nova_tabela.setStyle(TableStyle([
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("FONTNAME", (0, 0), (-1, -1), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 12),
        ("GRID", (0, 0), (-1, -1), 1, colors.black)
    ]))
    
    # Calcula a altura da nova tabela
    largura_nova, altura_nova = nova_tabela.wrap(0, 0)
    
    # Verifica se cabe na página atual
    if y_posicao - altura_nova < 50:
        pdf.showPage()
        pdf.drawImage("static/cabecalho.png", 0, 0, width=A4[0], height=A4[1])
        y_posicao = y_posicao_p2  # Reinicia a posição para a nova página
        
    # Desenha a nova tabela
    nova_tabela.drawOn(pdf, 28, (y_posicao - altura_nova)-10)

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/form_valvula")
def form_valvula():
    return render_template("form_valvula.html")

@app.route("/form_manometro")
def form_manometro():
    return render_template("form_manometro.html")  

@app.route("/gerar_pdf_manometro", methods=["GET","POST"])
def gerar_pdf_manometro():
    try:
        # Captura os dados do formulário
        data_inicio = request.form["data_inicio"]
        data_proxima = request.form["data_proxima"] 
        tag_manometro = request.form["tag_manometro"]
        modelo = request.form["modelo"]
        num_linhas = int(request.form["numLinhas"])
        tipo_manometro = request.form["tipo"]
        valor_divisao = request.form["valor_divisao"]
        unidade_pressao = request.form["unidade_pressao"]
        fluido = request.form["fluido_teste"]
        diametro = request.form["diametro_rosca"]
        resultado = request.form["resultado"]   

        if request.method == "POST":
            if 'foto_manometro' not in request.files:
                return 'Nenhum arquivo enviado' 
            foto = request.files["foto_manometro"]
            filepath = None

        if foto.filename == "":
            return 'Nenhum arquivo selecionado'


        if foto and foto.filename != "":
            filename = secure_filename(foto.filename)  # Garante um nome seguro
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)  # Caminho do arquivo
            foto.save(filepath)  # Salva a imagem no servidor

            max_width = 260
            max_height = 195

            # Ajusta o tamanho da imagem mantendo a proporção e melhora a qualidade
            try:
                with PilImage.open(filepath) as img:
                    img = img.convert("RGB")  # Converte para RGB caso seja um formato não suportado
                    img_width, img_height = img.size  # Obtém o tamanho original

                    # Calcula a proporção para manter o aspecto

                    if img_width > max_width or img_height > max_height:
                        ratio = min(max_width / img_width, max_height / img_height)  
                        new_width = int(img_width * ratio)
                        new_height = int(img_height * ratio)

                    img = ReportLabImage(filepath, width=new_width, height=new_height)
                    
            except Exception as e:
                print(f"Erro ao processar a imagem: {e}")
                filepath = None  # Se houver erro, a imagem não será usada no PDF

        else:
            filepath = None # Se não houver imagem, define o caminho como None
        
        
        data_inicio_formatada = datetime.strptime(data_inicio, "%Y-%m-%d").strftime("%d/%m/%Y")
        data_final_formatada = datetime.strptime(data_proxima, "%Y-%m-%d").strftime("%d/%m/%Y")
        
        # Cria um arquivo PDF na memória
        buffer = io.BytesIO()
        pdf = canvas.Canvas(buffer, pagesize=A4)
        pdf.drawImage("static/cabecalho.png", 0, 0, width=A4[0], height=A4[1])

        styles = getSampleStyleSheet()
        estilo_texto = styles["BodyText"]
        estilo_texto.fontSize = 14
        estilo_texto.fontName = "Helvetica"
        estilo_texto.alignment = 1
                
        # Cabeçalho do certificado
        titulo = Paragraph("<b>CERTIFICADO DE CALIBRAÇÃO / AFERIÇÃO</b>", estilo_texto)
        dados_tabela1 = [
            [titulo, Paragraph("<b>CERT.M DCE<br/>2025-68</b>", estilo_texto)],
            ["Data de emissão",f"{data_inicio_formatada}"],
            ["Data de validade", f"{data_final_formatada}"],
        ]
        
        estilo_tabela1 = TableStyle([
            ("ALIGN", (1, 0), (1, 2), "CENTER"),
            ("ALIGN", (0, 1), (0, 2), "LEFT"),
            ("SPAN", (0, 0), (1, 0)),
            ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
            ("FONTSIZE", (0, 0), (-1, -1), 12),
            ("BOTTOMPADDING", (0, 0), (1, 0), 6),
            ('VALIGN', (0, 0), (1, 0), 'MIDDLE'),
            ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
            ("GRID", (0, 0), (-1, -1), 1, colors.black)
        ])
        
        t1 = Table(dados_tabela1, colWidths=[400, 140])
        t1.setStyle(estilo_tabela1)

        dados_tabela2 = [
            ["IDENTIFICAÇÃO DO INSTRUMENTO PADRÃO", ""],
            ["Modelo / Fabricante", "NC"],
            ["Tipo / N° Série", "DIGITAL / 2401170548065"],
            ["Faixa de escala", "0 A 25 MPA"],
            ["Valor de uma divisão", "0,001 MPA"],
            ["Classe de exatidão", "1,0%"],
            ["Origem / Nº certificado", "(CAL ISO 0506 PROXY METROLOGIA) / Nº 8390-24"]
        ]

        estilo_tabela2 = TableStyle([
            ("ALIGN", (0, 0), (-1, -1), "LEFT"),
            ("SPAN", (0, 0), (1, 0)),
            ("ALIGN", (0, 0), (1, 0), "CENTER"),
            ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
            ("FONTNAME", (0, 0), (1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, -1), 10),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 1),
            ("TOPPADDING", (0, 0), (-1, -1), 1),
            ("BOTTOMPADDING", (0, 0), (1, 0), 6),
            ("TOPPADDING", (0, 0), (1, 0), 6),
            ("GRID", (0, 0), (-1, -1), 1, colors.black)
        ])

        t2 = Table(dados_tabela2, colWidths=[270, 270])
        t2.setStyle(estilo_tabela2)

        

        titulo_t3 = Paragraph("<b>IDENTIFICAÇÃO DO INSTRUMENTO EM TESTE</b>", estilo_texto)

        dados_tabela3 = [
            [titulo_t3, ""],
            [img, ""],
            ["Modelo / Fabricante", f"{modelo}"],
            ["Tipo / N° Série", f"{tipo_manometro}"],
            ["Faixa de escala", f"0 a {num_linhas} {unidade_pressao}"],
            ["Valor de uma divisão", f"{valor_divisao} {unidade_pressao}"],
            ["Classe de exatidão", "1,0%"],
            ["Tag Equipamento", f"{tag_manometro}"],
            ["Fluido de Teste", f"{fluido}"],
            ["Diâmetro / Rosca", f"{diametro}"],
        ]

        estilo_tabela3 = TableStyle([
            ("ALIGN", (0, 0), (-1, -1), "LEFT"),
            ("SPAN", (0, 0), (1, 0)),
            ("SPAN", (0, 1), (1, 1)),
            ("ALIGN", (0, 0), (1, 1), "CENTER"),
            ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
            ("FONTNAME", (0, 0), (1, 0), "Helvetica-Bold"),
            ('VALIGN', (0, 0), (1, 1), 'MIDDLE'),
            ("FONTSIZE", (0, 0), (-1, -1), 10),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 1),
            ("TOPPADDING", (0, 0), (-1, -1), 1),
            ("BOTTOMPADDING", (0, 0), (1, 1), 6),
            ("TOPPADDING", (0, 0), (1, 1), 6),
            ("GRID", (0, 0), (-1, -1), 1, colors.black)
        ])

        t3 = Table(dados_tabela3, colWidths=[270, 270])
        t3.setStyle(estilo_tabela3)

        x_texto = 32
        y_texto = 130

        #text = Paragraph("Manômetro calibrado em comparação a um Padrão com auxílio de uma bancada de<br/>calibração. O resultado representa uma média de 04 leituras feitas pelo Padrão de dois ciclos<br/> (crescentes e decrescente). Procedimento de calibração conforme NBR 14105-1:2011")
        pdf.drawString(x_texto, y_texto, "Manômetro calibrado em comparação a um Padrão com auxílio de uma bancada de calibração." )
        pdf.drawString(x_texto, y_texto-12, "O resultado representa uma média de 04 leituras feitas pelo Padrão de dois ciclos (crescentes" )
        pdf.drawString(x_texto, y_texto-25, "e decrescente). Procedimento de calibração conforme NBR 14105-1:2011" )    

        # Posicionar elementos no PDF
        y_posicao = 650
        x_posicao = 28
        t1.wrapOn(pdf, 0, 0)
        t1.drawOn(pdf, x_posicao, y_posicao)

        t2.wrapOn(pdf, 0, 0)
        y_posicao2 = t2._height
        y_posicao2 = (y_posicao - y_posicao2) - 15
        t2.drawOn(pdf, x_posicao, y_posicao2)

        t3.wrapOn(pdf, 0, 0)
        y_posicao3 = t3._height
        y_posicao3 = (y_posicao2 - y_posicao3) - 15
        t3.drawOn(pdf, x_posicao, y_posicao3)


        dados_t1_p2 = [
            ["DADOS OBTIDOS DURANTE A CALIBRAÇÃO DO INSTRUMENTO","" ,"" ,"" ,"" ,""],
            [Paragraph("Instrumento <br/>em teste"), "Pressões indicadas pelo padrão (2 ciclos)","" ,"" ,"" ,"Média" ],
            ["", "1º Crescente", "1º Decrescente", "2º Crescente", "2º Decrescente", ""]    
        ]


        estilo_t1_p2 = TableStyle([
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
            ("FONTSIZE", (0, 0), (-1, -1), 10),
            ("SPAN", (0, 0), (-1, 0)),
            ("SPAN", (0, 1), (0, 2)),
            ("SPAN", (1, 1), (4, 1)),
            ("SPAN", (-1, 1), (-1, 2)),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ("FONTNAME", (0, 0), (1, 0), "Helvetica-Bold"),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
            ("TOPPADDING", (0, 0), (-1, -1), 2),
            ("GRID", (0, 0), (-1, -1), 1, colors.black)
        ])

        dados_t2_p2 = [
            [f"{unidade_pressao}", f"{unidade_pressao}", f"{unidade_pressao}", f"{unidade_pressao}", f"{unidade_pressao}", f"{unidade_pressao}"]
            
        ]

        estilo__t2_p2 = TableStyle([
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
            ("FONTSIZE", (0, 0), (-1, -1), 10),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
            ("TOPPADDING", (0, 0), (-1, -1), 2),
            ("GRID", (0, 0), (-1, -1), 1, colors.black)
        ])

         # Gerar tabela dinâmica
        dados_tabela = [
            ["", "", "", "", "", ""]            
        ]

        
        
        for i in range(1, num_linhas + 1):
            valores = [gerar_valor_aleatorio(i) for _ in range(4)]
            media = round(sum(valores) / 4, 2)
            dados_tabela.append([i] + valores + [media])
                

        pdf.showPage()
        pdf.drawImage("static/cabecalho.png", 0, 0, width=A4[0], height=A4[1])

        t1_p2 = Table(dados_t1_p2, colWidths=[70, 100, 100, 100, 100, 70])
        t1_p2.setStyle(estilo_t1_p2)

        T2_p2 = Table(dados_t2_p2, colWidths=[70, 100, 100, 100, 100, 70])
        T2_p2.setStyle(estilo__t2_p2)

        y_posicao_p2 = 680
        x_posicao_p2 = 28

        t1_p2.wrapOn(pdf, 0, 0)
        t1_p2.drawOn(pdf, x_posicao_p2, y_posicao_p2)   

        T2_p2.wrapOn(pdf, 0, 0)
        y_posicao2_p2 = T2_p2._height
        y_posicao2_p2 = y_posicao_p2 - y_posicao2_p2
        T2_p2.drawOn(pdf, x_posicao_p2, y_posicao2_p2)
        
       
        desenhar_tabela_dinamica(pdf, dados_tabela, y_posicao2_p2, resultado)
        
        pdf.save()
        buffer.seek(0)

        if filepath and os.path.exists(filepath):
            os.remove(filepath)
        
        return send_file(buffer, as_attachment=True, download_name=f"PI - {tag_manometro}.pdf", mimetype="application/pdf")
    
    except Exception as e:
        return f"Ocorreu um erro: {str(e)}"

@app.route("/gerar_psv_pdf", methods=["GET", "POST"])
def gerar_psv_pdf():
    try:
        # Captura os dados do formulário
        data_inicio = request.form["data_inicio"]
        data_proxima = request.form["data_proxima"]
        tag_valvula = request.form["tag_valvula"]
        modelo_valvula = request.form["modelo_valvula"]
        diametro_valvula = request.form["diametro_valvula"]
        fluido_teste = request.form["fluido_teste"]
        pressao_abertura = request.form["pressao_abertura"]
        pressao_fechamento = request.form["pressao_fechamento"]
        unidade_pressao = request.form["unidade_pressao"]
        fabricante = request.form["fabricante"]

        #Salva a imagem temporariamente
        
        if request.method == "POST":
            # Verifica se um arquivo foi enviado
            if 'foto_valvula' in request.files:
                foto = request.files["foto_valvula"]

                # Verifica se o arquivo tem um nome (ou seja, se algo foi selecionado)
                if foto.filename != "":
                    filename = secure_filename(foto.filename)
                    filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
                    foto.save(filepath)

                    # Ajuste o tamanho da imagem mantendo a proporção
                    max_width = 260
                    max_height = 195

                    try:
                        with PilImage.open(filepath) as img:
                            img = img.convert("RGB")  # Converte para RGB caso seja um formato não suportado
                            img_width, img_height = img.size

                            # Calcula a proporção para manter o aspecto
                            if img_width > max_width or img_height > max_height:
                                ratio = min(max_width / img_width, max_height / img_height)
                                new_width = int(img_width * ratio)
                                new_height = int(img_height * ratio)
                            
                            # Ajusta a imagem para o ReportLab
                            img = ReportLabImage(filepath, width=new_width, height=new_height)

                    except Exception as e:
                        print(f"Erro ao processar a imagem: {e}")
                        filepath = None  # Se houver erro, a imagem não será usada no PDF
                else:
                    return 'Nenhum arquivo selecionado'
            else:
                return 'Nenhum arquivo enviado'

        data_incio_formatada = datetime.strptime(data_inicio, "%Y-%m-%d").strftime("%d/%m/%Y")
        data_final_formatada = datetime.strptime(data_proxima, "%Y-%m-%d").strftime("%d/%m/%Y")

        # Caminho da imagem de fundo
        imagem_fundo = "static/cabecalho.png"

        # Cria um arquivo PDF na memória
        buffer = io.BytesIO()
        pdf = canvas.Canvas(buffer, pagesize=A4)

        # Adiciona a imagem de fundo
        pdf.drawImage(imagem_fundo, 0, 0, width=A4[0], height=A4[1])

        styles = getSampleStyleSheet()
        estilo_texto = styles["BodyText"]
        estilo_texto.alignment = 1
        estilo_texto.fontSize = 10
        estilo_texto.fontName = "Helvetica"


        titulo = Paragraph("<b>CERTIFICADO DE CALIBRAÇÃO<br/>VÁLVULA DE SEGURANÇA - (PSV)</b>", estilo_texto)
        titulo2 = Paragraph("<b>CONDIÇÕES FISICAS DA VÁLVULA</b>", estilo_texto)
    
        # Tabela principal com as informações do certificado
        tabela_certificado = [
            [titulo],
            [f"DATA: {data_incio_formatada}", f"CERTIFICADO Nº: PSV-{tag_valvula}"],
            ["CARACTERÍSTICAS TÉCNICAS", "DADOS DO EQUIPAMENTO PROTEGIDO"],
    
        ]

        # Define o estilo da tabela do cabeçalho
        estilo_certificado = TableStyle([
            ("SPAN", (0, 0), (1, 0)),  # Mescla a primeira linha em duas colunas
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),  # Centraliza o texto
            ("FONTNAME", (0, 0), (-1, -1), "Helvetica-Bold"),  # Negrito
            ("FONTNAME", (0, 1), (1, 1), "Helvetica"),
            ("FONTSIZE", (0, 0), (-1, -1), 10),
            ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),  # Cor de fundo
            ("BACKGROUND", (0, 2), (1, 2), colors.lightgrey),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
            ("GRID", (0, 0), (-1, -1), 1, colors.black)  # Adiciona borda
        ])

        t1 = Table(tabela_certificado, colWidths=[270, 270])
        t1.setStyle(estilo_certificado)

        dados_t2 = [
            [Paragraph(f"TAG/LACRE:<br/><b>PSV-{tag_valvula}</b>"), "NORMA: ASME", img, ""],
            [Paragraph(f"FABRICANTE:<br/><b>{fabricante}</b>"), Paragraph(f"FREQUÊNCIA DE<br/>CALIBRAÇÃO:<br/><b>12 MESES</b>"), "", ""],
            [Paragraph("CONSTRUÇÃO:<br/><b>FOFO</b>"), Paragraph(f"TIPO DE INTERVENÇÃO:<br/><b>PREVENTIVA</b>"), "", ""],
            [Paragraph("TIPO DE SEDE<br/><b>PLANA</b>"), Paragraph(f"PRÓXIMA INSPEÇÃO:<br/><b>{data_final_formatada}</b>"), "", ""],
            [Paragraph(f"MODELO: <b>{modelo_valvula}</b>"), Paragraph(f"DIÂMETRO:<br/><b>{diametro_valvula}</b>"), "", ""],
        ] 

        estilo_t2 = TableStyle([
            ("ALIGN", (0, 0), (-1, -1), "LEFT"),  # Centraliza o texto
            ("ALIGN", (2, 0), (3, 0), "CENTER"),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ("SPAN", (2, 0), (3, 0)),
            ("SPAN", (2, 0), (3, 4)),
            ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
            ("FONTSIZE", (0, 0), (-1, -1), 10),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
            ("GRID", (0, 0), (-1, -1), 1, colors.black)  # Adiciona borda
        ])

        t2 = Table(dados_t2, colWidths=[135, 135, 135, 135])
        t2.setStyle(estilo_t2)

        styles_t3 = getSampleStyleSheet()
        estilo_texto4 = styles_t3["BodyText"]
        estilo_texto4.alignment = 1
        estilo_texto4.fontSize = 9
        estilo_texto4.fontName = "Helvetica-Bold"

        dados_t3 = [
            [titulo2, "", "", ""],
            [Paragraph("ITEM", estilo_texto4), Paragraph("COMPONETES", estilo_texto4), Paragraph(f"CONDIÇÕES<br/>ENCONTRADAS", estilo_texto4), Paragraph(f"REPARO<br/>NECESSÁRIO", estilo_texto4)],
            ["01", "LACRE", "OK", ""],
            ["02", "PINTURA", "OK", ""],
            ["03", "ROSCA DE CORPO E BOCAL", "OK", ""],
            ["04", "MOLA", "OK", ""],
            ["05", "CONDIÇÕES PAR. DE TRAVA DO ANEL", "OK", ""]
        ]

        estilo_t3 = TableStyle([
            ("SPAN", (0, 0), (3, 0)),  # Mescla a primeira linha em quatro colunas
            ("FONTSIZE", (0, 0), (3, 0), 9),  # Mescla a primeira linha em quatro colunas
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),  # Centraliza o texto
            ("ALIGN", (1, 2), (1, 6), "LEFT"),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
            ("FONTSIZE", (0, 0), (-1, -1), 8),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 1),
            ("TOPPADDING", (0, 0), (-1, -1), 1),
            ("BOTTOMPADDING", (0, 0), (3, 0), 4),
            ("TOPPADDING", (0, 0), (3, 0), 4),
            ("GRID", (0, 0), (-1, -1), 1, colors.black)
        ])

        t3 = Table(dados_t3, colWidths=[80, 200, 100, 160])
        t3.setStyle(estilo_t3)

        styles_t4 = getSampleStyleSheet()
        estilo_texto4 = styles_t4["BodyText"]
        estilo_texto4.alignment = 1
        estilo_texto4.fontSize = 9
        estilo_texto4.fontName = "Helvetica-Bold"
        estilo_texto4.topPadding = 3
        estilo_texto4.bottomMargin = 3

        styles2_t4 = getSampleStyleSheet()
        styles2_t4 = styles2_t4["BodyText"]
        styles2_t4.alignment = 0
        styles2_t4.fontSize = 8
        styles2_t4.fontName = "Helvetica"

    
        
        dados_t4 = [
            [Paragraph("CALIBRAÇÃO E TESTE FINAL", estilo_texto4)],
            [Paragraph(f"FLUIDO DE TESTE: <b>{fluido_teste}</b>", styles2_t4)],
            [Paragraph(f"PRESSÃO DE ABERTURA: <b>{pressao_abertura} {unidade_pressao}</b>", styles2_t4)],
            [Paragraph(f"PRESSÃO DE FECHAMENTO: <b>{pressao_fechamento} {unidade_pressao}</b>", styles2_t4)],
            [Paragraph("TESTE DE VEDAÇÃO: <b>APROVADO</b>", styles2_t4)],
            [Paragraph("TESTE DE INTEGRIDADE: <b>CORPO APROVADO</b>", styles2_t4)],
            [Paragraph("TEMPERATURA: <b>25 °C</b>", styles2_t4)],
            [Paragraph("CAPACIDADE DE PRESSURIZAÇÃO: <b>60 kgf/cm²</b>", styles2_t4)],
            [Paragraph("ESCALA: <b>0 - 250 kgf/cm²</b>", styles2_t4)],
            [Paragraph("CERT. DE CALIBRAÇÃO:  <b>Nº 8390-24</b>", styles2_t4)],
            [Paragraph("( X )APROVADO&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;( &nbsp;&nbsp; )REPROVADO", estilo_texto4)],
        ]

        estilo_t4 = TableStyle([
            ("ALIGN", (0, 0), (-1, -1), "LEFT"),  
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
            ("FONTSIZE", (0, 0), (-1, -1), 8),
            ("BACKGROUND", (0, 0), (0, 0), colors.lightgrey),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
            ("TOPPADDING", (0, 0), (-1, -1), 0),
            ("BOTTOMPADDING", (0, 0), (1, 0), 3),
            ("TOPPADDING", (0, 0), (1, 0), 3),
            ("GRID", (0, 0), (-1, -1), 1, colors.black)
        ])

        t4 = Table(dados_t4, colWidths=[540])
        t4.setStyle(estilo_t4)

        y_posicao = 640
        x_posicao = 28

        # Posicionando tabelas no PDF
        t1.wrapOn(pdf, 0, 0)
        t1.drawOn(pdf, x_posicao, y_posicao)  # Posição ajustada

        t2.wrapOn(pdf, 0, 0)
        y_posicao2 = t2._height
        y_posicao2= y_posicao - y_posicao2
        t2.drawOn(pdf, x_posicao, y_posicao2)  # Posição ajustada


        t3.wrapOn(pdf, 0, 0)
        y_posicao3 = t3._height
        y_posicao3 = y_posicao2 - y_posicao3
        t3.drawOn(pdf, x_posicao, y_posicao3)

        t4.wrapOn(pdf, 0, 0)
        y_posicao4 = t4._height
        y_posicao4 = (y_posicao3 - y_posicao4) - 15
        t4.drawOn(pdf, x_posicao, y_posicao4)




        pdf.save()
        buffer.seek(0)

        if filepath and os.path.exists(filepath):
            os.remove(filepath)

        return send_file(buffer, as_attachment=True, download_name=f"PSV-{tag_valvula}.pdf", mimetype="application/pdf")
    
    except Exception as e:
        return f"Ocorreu um erro: {str(e)}"


if __name__ == "__main__":
    app.run(debug=True)
