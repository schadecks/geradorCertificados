import PySimpleGUI as sg
import openpyxl
from PIL import Image, ImageDraw, ImageFont
from datetime import datetime


workbookAlunos = openpyxl.load_workbook(filename='C:/Users/eduar/OneDrive/Arquivos/Projetos/geradorCertificados/Planilhas/planilhaAlunos.xlsx',read_only=True)
sheetAlunos = workbookAlunos['Alunos']

fontNormal = ImageFont.truetype('C:/Users/eduar/OneDrive/Arquivos/Projetos/geradorCertificados/Fontes/CONSOLA.TTF',size=20)
fontData = ImageFont.truetype('C:/Users/eduar/OneDrive/Arquivos/Projetos/geradorCertificados/Fontes/CONSOLA.TTF',size=18)
fontBold = ImageFont.truetype('C:/Users/eduar/OneDrive/Arquivos/Projetos/geradorCertificados/Fontes/CONSOLAB.TTF',size=28)

for indice, linha in enumerate(sheetAlunos.iter_rows(min_row=2)):
  aluno = linha[0].value
  cpf = linha[1].value
  curso = linha[2].value
  diretor = linha[3].value
  dataInicio = linha[4].value
  dataFim = linha[5].value
  cargaHoraria = linha[6].value
  
  image = Image.open('C:/Users/eduar/OneDrive/Arquivos/Projetos/geradorCertificados/Certificados/certificado-exemplo.jpg')
  certificado = ImageDraw.Draw(image)
  certificado.text((310,340),aluno, fill='black',font=fontBold)
  certificado.text((190,415),cpf, fill='black',font=fontNormal)
  certificado.text((625,415),curso, fill='black',font=fontNormal)
  certificado.text((330,450),datetime.strftime(dataInicio,'%d/%m/%Y'), fill='black',font=fontData)
  certificado.text((460,450),datetime.strftime(dataFim,'%d/%m/%Y'), fill='black',font=fontData)
  certificado.text((860,450),str(cargaHoraria), fill='black',font=fontNormal)
  certificado.text((800,620),diretor, fill='black',font=fontNormal)
  certificado.text((730,685),'E.E.B. Santos Dumont', fill='black',font=fontNormal)
  image.save(f'C:/Users/eduar/OneDrive/Arquivos/Projetos/geradorCertificados/Certificados/{indice} - {aluno}.jpg')
  
# class  Singleton(object):
#   def __new__( cls ):
#       if not hasattr( cls, 'instance' ):
#           cls.instance = super(Singleton, cls ) .__new__( cls )
#       return cls.instance

# def setTheme(window, new_theme):
#   global CURRENT_THEME
#   CURRENT_THEME = new_theme
#   sg.theme(new_theme)
#   window.TKroot.config(background=sg.theme_background_color())
#   for element in window.element_list():
#       element.Widget.config(background=sg.theme_background_color())
#       element.ParentRowFrame.config(background=sg.theme_background_color())
#       if 'text' in str(type(element)).lower():
#           element.Widget.config(foreground=sg.theme_text_color())
#           element.Widget.config(background=sg.theme_text_element_background_color())
#       if 'input' in str(type(element)).lower():
#           element.Widget.config(foreground=sg.theme_input_text_color())
#           element.Widget.config(background=sg.theme_input_background_color())
#       if 'progress' in str(type(element)).lower():
#           element.Widget.config(foreground=sg.theme_progress_bar_color()[0])
#           element.Widget.config(background=sg.theme_progress_bar_color()[1])
#       if 'slider' in str(type(element)).lower():
#           element.Widget.config(foreground=sg.theme_slider_color())
#       if 'button' in str(type(element)).lower():
#           element.Widget.config(foreground=sg.theme_button_color()[0])
#           element.Widget.config(background=sg.theme_button_color()[1])
#   window.Refresh()

# Dados = Singleton()
  
# def update():
#   print( 'update' )
#   Dados.cotation = random.randrange(0,100)
#   Dados.test2 = random.randrange(0,100)
  
# class TelaPython:    
#   def __init__(self):
#     sg.change_look_and_feel('Dark')
    
#     layout = [
#         [sg.Text('Nome',size=(8,0)),sg.Input(size=(25,0),key='nome')],
#         [sg.Text('Idade',size=(8,0)),sg.Input(size=(25,0),key='idade')],
#         [sg.Text('Quais provadores de e-mail são aceitos?')],
#         [sg.Checkbox('Gmail',key='gmail'),sg.Checkbox('Outlook',key='outlook'),sg.Checkbox('Yahoo',key='yahoo')],
#         [sg.Text('Aceita Cartão?')],
#         [sg.Radio('Sim','cartoes',key='aceitaCartao'), sg.Radio('Não','cartoes',key='naoAceitaCartao')],
#         [sg.Slider(range(0,100,10),default_value=0,orientation='horizontal',size=(10,20),key='slider',tick_interval=10,expand_x=True)],
#         [sg.Text('Cotação: '),sg.Text('',key='cotation')],
#         [sg.Text('Selecione seu Tema')],
#         [sg.Radio('Dark','temas',key='temaDark'), sg.Radio('DarkBlue3','temas',key='temaDarkBlue')],
#         [sg.Button('Enviar Dados',)],
#         [sg.Output(size=(80,20))],
#     ]
    
#     self.janela = sg.Window("Dados do Usuário").layout(layout)

#     while True:
#       self.event, self.values = self.janela.Read(timeout=1000)
      
#       if self.event == sg.WIN_CLOSED or self.event == 'Exit': 
#         break
      
#       nome = self.values['nome']
#       idade = self.values['idade']
#       aceita_gmail = self.values['gmail']
#       aceita_outlook = self.values['outlook']
#       aceita_yahoo = self.values['yahoo']
#       aceita_cartao = self.values['aceitaCartao']
#       nao_aceita_cartao = self.values['naoAceitaCartao']
#       velocidadeScript = self.values['slider']
#       self.janela['cotation'](Dados.cotation)
      
#       # if self.values['temaDark'] != True:
#       #   setTheme(self.janela, 'DarkBlue3')
#       # elif self.values['temaDarkBlue'] != True:
#       #   setTheme(self.janela, 'Dark')
      
      
#       print(f'Nome: {nome}')
#       print(f'Idade: {idade}')
#       print(f'Gmail: {aceita_gmail}')
#       print(f'Outlook: {aceita_outlook}')
#       print(f'Yahoo: {aceita_yahoo}')
#       print(f'Aceita_cartao: {'Sim' if aceita_cartao == True else 'Nao'}')
#       print(f'Slider: {velocidadeScript}')
#       print(f'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - ')
#       self.janela.Refresh()
#       update()

# update()
# TelaPython()