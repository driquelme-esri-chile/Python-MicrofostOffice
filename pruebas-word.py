"""Codigo no funcional, falta organizarlo, template de word
este codigo es solo para realizar pruebas"""


from asyncore import write
import datetime
import pandas as pd
import locale



from docxtpl import DocxTemplate
from itertools import groupby
from operator import itemgetter



nombre_salida: str = datetime.datetime.today().strftime('%Y-%m-%d  %H,%M,%S')


Analisis = ["lala","lele","lili","lolo","lulu","loli"]
Nombre = ["rala","lrje","dooli","jilo","lupuuu","lueli"]
Te = ["malm","meme","cicc","fod","ruu","boois"]
Est = ["kjajkla","oper","oisfi","oofo","uuu","ioli"]
Concesionario = ["oieg","yatsd","idtg","fgkj","ysadu","mami"]
Area = ["lala","lele","lili","lolo","lulu","loli"]
Rol = ["lala","lele","lili","lolo","lulu","loli"]

    
df = pd.DataFrame({
    'Analisis': Analisis,
    'Nombre': Nombre,
    'T': Te,
    'Est': Est,
    'Concesionario': Concesionario,
    'Area': Area,
    'Rol': Rol
})

print(df)

now = datetime.datetime.now()
fechaWord = now.strftime("%B %d %Y")

aux = []
itemsWord = []
valores = sorted(items, key = itemgetter('empresa1'))

for key, value in groupby(valores, key = itemgetter('empresa1')):
    aux = []
    for k in value:
        aux.append(k['data'])
    itemsWord.append({
        'empresa1': key,
        'empresa2': 'SQM S.A',
        'data': aux
    })



context = {
    'fecha': fechaWord,
    'items': itemsWord
}


fechaSalida: str = datetime.datetime.today().strftime('%Y-%m-%d  %H,%M,%S')

# Genero el archivo word con el resumen de los datos.
wordName = 'Diario Oficial {}.docx'.format(fechaSalida)

tpl.render(context)
outputFileWord: str = jobsDir + "\\" + wordName
tpl.save(outputFileWord)
