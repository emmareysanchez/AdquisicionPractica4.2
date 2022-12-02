import xlsxwriter 
import pandas as pd

def cargar_datos():
    df_prediccion = pd.read_csv('prediccion_final.csv')
    df_popularidad = pd.read_csv('popularidad_pizzas.csv')
    return df_prediccion, df_popularidad

def crear_excel(df_prediccion, df_popularidad):
    df_prediccion=df_prediccion[["ingredientes","cantidad"]]
    cuaderno = xlsxwriter.Workbook('chart_line.xlsx')
    # Añado la hoja 1 al cuaderno
    worksheet = cuaderno.add_worksheet(name = "COMPRA_SEMANAL")
    #Añado la tabla con los datos a analizar
    worksheet.write_column('A1', df_prediccion['ingredientes'].values)
    worksheet.write_column('B1', df_prediccion['cantidad'].values)
    # Creo una gráfica de barras.
    chart = cuaderno.add_chart({'type': 'column'})
    chart.add_series({
    'categories': '=COMPRA_SEMANAL!$A$1:$A$65',
    'values':     '=COMPRA_SEMANAL!$B$1:$B$65',
    })
    # Personalizo la gráfica incluyendo nombres de ejes y títulos
    chart.set_title ({'name': 'COMPRA SEMANAL DE INGREDIENTES'})
    chart.set_y_axis({'name': 'Cantidad'})
    chart.set_x_axis({'name': 'Ingredientes'})
    chart.set_x_axis({'num_font': {'rotation' : 45}})
    chart.set_style(10)
    chart.set_size({'width': 2000, 'height': 576})
    # Inserto la gráfica en la hoja
    worksheet.insert_chart('D2', chart)

    #Creo la hoja 2 del cuaderno
    worksheet2 = cuaderno.add_worksheet(name = "POPULARIDAD_PIZZAS")
    #Creo la tabla con los datos a analizar
    worksheet2.write_column('A1', df_popularidad['pizza_id'].values)
    worksheet2.write_column('B1', df_popularidad['quantity'].values)
    # Creo una gráfica de barras.
    chart2 = cuaderno.add_chart({'type': 'bar'})
    chart2.add_series({
    'categories': '=POPULARIDAD_PIZZAS!$A$1:$A$32',
    'values':     '=POPULARIDAD_PIZZAS!$B$1:$B$32',
    })
    # Añado los nombres de los ejes y el título
    chart2.set_title ({'name': 'POPULARIDAD PIZZAS'})
    chart2.set_y_axis({'name': 'Tipo Pizza'})
    chart2.set_x_axis({'name': 'Ventas'})
    chart2.set_x_axis({'num_font': {'rotation' : 45}})
    chart2.set_style(10)
    chart2.set_size({'width': 2000, 'height': 800})
    # Inserto la gráfica en la hoja
    worksheet2.insert_chart('D2', chart2)
    #Selección de datos
    bottom5  = df_popularidad.head(n= 5)
    top5 = df_popularidad.tail(n= 5)
    #Creo la hoja 3 del cuaderno
    worksheet3 = cuaderno.add_worksheet(name = "TOP_5_PIZZAS")
    worksheet3.write_column('A1', top5['pizza_id'].values)
    worksheet3.write_column('B1', top5['quantity'].values)
    worksheet3.write_column('N1', bottom5['pizza_id'].values)
    worksheet3.write_column('O1', bottom5['quantity'].values)
    # Creo una gráfica de barras.
    chart3 = cuaderno.add_chart({'type': 'pie'})
    chart3.add_series({
    'categories': '=TOP_5_PIZZAS!$A$1:$A$5',
    'values':     '=TOP_5_PIZZAS!$B$1:$B$5',
    })
    #Personalizo la gráfica
    chart3.set_title ({'name': 'TOP 5 PIZZAS MÁS VENDIDAS'})
    chart3.set_style(10)
    chart3.set_size({'width': 500, 'height': 500})
    # Inserto la gráfica en la hoja
    worksheet3.insert_chart('D2', chart3)

    chart4 = cuaderno.add_chart({'type': 'pie'})
    chart4.add_series({
    'categories': '=TOP_5_PIZZAS!$N$1:$N$5',
    'values':     '=TOP_5_PIZZAS!$O$1:$O$5',
    })
    chart4.set_title ({'name': 'LAS 5 PIZZAS MENOS VENDIDAS'})
    chart4.set_style(10)
    chart4.set_size({'width': 500, 'height': 500})
    # Inserto la gráfica en la hoja
    worksheet3.insert_chart('Q2', chart4)
    cuaderno.close()

if __name__ == '__main__':
    df1, df2 = cargar_datos()
    crear_excel(df1, df2)
