from openpyxl  import  Workbook
from openpyxl.styles  import Font
from openpyxl.chart  import BarChart
from openpyxl.chart import Reference as opyxlReference

wb=Workbook()
ws=wb.active
ws.title="seguidores"
#colocando os dados na nossa tabela
dados=[["Plataforma","nome do usuario ","numero de seguidores"],
       ["Facebook","setprogramação",1000],
       [" Instagram","setprogramação",1800],
       ["youtube ","setprogramação",1200]
       ]

for dado in dados:
    ws.append(dado)
# colocando o nosso cabeçalho de tabela em negrito  
ft=Font(bold=True)
for linha in ws ["A1:C1"]:
    for celula in linha:
        celula.font=ft

      
#criando o nosso grafico de barras

chart=BarChart()
chart.type='col'
chart.title='grafico de controle de seguidores'
chart.y_axis.title='numero de seguidores'
chart.x_axis.title='plataformas'
chart.legend=None


data=opyxlReference(ws,min_col=3,max_col=3,min_row=2,max_row=4)
categoria=opyxlReference(ws,min_col=1,max_col=1,min_row=2,max_row=4)

chart.add_data(data)
chart.set_categories(categoria)


ws.add_chart(chart,'E1')
wb.save(r"seguidores.xlsx")
