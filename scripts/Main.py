#Carga de funciones creadas en el fichero Glosario_lib.py
from Glosario_lib import *
#Carga de paquete para obtener la fecha actual
from datetime import date
#Carga de paqeute para llamadas al sistema
import os

#Fichero donde se almacena las propiedades de incio sesion
properties="Properties.txt"

#Nombre del Glosario con formato Glosario_YYYYMMDD.xlsx con la fecha que se ejecuto
today = str(date.today())
today=today.replace("-","")
path= r".\Glosario"+"_"+today+".xlsx"

print("eliminado registros temporales del sistema")

#Elimina todos los ficheros que almacena la salida y los scripts de Command Manager
clear_files()

print("registros temporales eliminados")

#Obtiene el nombre del origen y el usuario a partir del fichero properties
origen,usuario=get_properties(properties)


print("Se va a crear el Glosario con los siguientes datos")
print("Nombre de origen de datos : "+origen )
print("Nombre de Usuario : "+usuario )
password=input("Introduzca la contraseña del Usuario : ")

#Peticion de la contraseña
if(password!=""):
    password=" -p "+password

print("obteniendo proyectos con esta configuracion")

#LLamanda al Command Manager para obtener el listado de proyectos y almacenarlos en projects.out
get_inicio_sesion(origen,usuario,password)

#Abre el fichero projects.out y lee todos los projectos y lo almancena en un vector
projects=list_projects('salida/projects.out')

#Salida en pantalla de todos los proyectos que tenemos acceso
print("Elija el numero de proyecto : ")
for i in range(len(projects)):
    print(str(i)+" : "+projects[i])
#Peticion del proyecto
proyecto = int(input("Ponga el numero de proyecto con que quiera obtener el Glosario: "))

#Filtro de inserccion de un proyecto correcto
while(proyecto<0 or proyecto>=len(projects)):
    proyecto = int(input("Ponga correctamente el numero, debe ser entre 0 y "+str(len(projects)-1)+" : "))

print("Creando scp temporales")

#Creacion de todos los scripts de obtencion de las propiedades de todos los objetos
create_scp(projects[proyecto])

print("Scp temporales creados")

#Ejecucion de los Scripts usando el Command Manager y almancenadolos en sus correspondientes salidas
get_properties_command(origen,usuario,password)

print("Todas las caracteristicas obtenidas")

print("Creando Excel")

#Lectura de las propiedades de todos los hechos, filtrados por las datos que queremos y almacenados en un dataframe "hechos"
hechos=file_hechos('salida/hechos.out','./salida/detalles_hechos.out')

#Lectura de las propiedades de todos los atributos, filtrados por las datos que queremos y almacenados en un dataframe "Atributo"
Atributo=file_atributos('salida/atributos.out','./salida/detalles_atributos.out')

#Lectura de las propiedades de todos las metricas, filtrados por las datos que queremos y almacenados en un dataframe "metricas"
metricas=file_metricas('salida/metricas.out','./salida/detalles_metricas.out')

#Lectura de las propiedades de todos los filtros, filtrados por las datos que queremos y almacenados en un dataframe "filtros"
filtros=file_filtros('salida/detalles_filtros.out')

#Filtrado del dataframe Atributo para quedarnos con los Atributos que esten en el directorio "Corporativo"
#Atributo=Atributo[mask_coporativo(Atributo,"Corporativo")]

#Filtrado del dataframe metrica para quedarnos con los metrica que esten en el directorio "Corporativo"
#metricas=metricas[mask_coporativo(metricas,"Corporativo")]

#Formato del dataframe "hechos" para que las columnas 'Nombre', 'Ruta', 'Expresión', 'Tabla de origen' esten en color gris
hechos_style=hechos.style.applymap(highlight_cols_grey,subset=pd.IndexSlice[:, ['Nombre', 'Ruta', 'Expresión', 'Tabla de origen']])

#Formato del dataframe "Atributo" para que las columnas 'Nombre', 'Ruta', 'Representacion', 'Tabla de origen', 'Expresion' esten en color gris
Atributo_style=Atributo.style.applymap(highlight_cols_grey, subset=pd.IndexSlice[:, ['Nombre', 'Ruta', 'Representacion', 'Tabla de origen', 'Expresion']])

#Formato del dataframe "metricas" para que las columnas 'Nombre', 'Ruta', 'Fórmula', 'Condición', 'Transformación', 'Expresión' esten en color gris
metricas_style=metricas.style.applymap(highlight_cols_grey, subset=pd.IndexSlice[:,['Nombre', 'Ruta', 'Fórmula', 'Condición', 'Transformación', 'Expresión']])

#Formato del dataframe "filtros" para que las columnas 'Nombre Filtro', 'Expresión' esten en color gris
filtros_style=filtros.style.applymap(highlight_cols_grey, subset=pd.IndexSlice[:, ['Nombre Filtro', 'Expresión']])

print("Formateando Excel")

#Creacion del excel a partir de los anteriores dataframes
writer = pd.ExcelWriter(path, engine='xlsxwriter')
Atributo_style.to_excel(writer, 'Atributos', index=False)
metricas_style.to_excel(writer, 'Indicadores', index=False)
hechos_style.to_excel(writer, 'Hechos', index=False)
filtros_style.to_excel(writer, 'Filtros', index=False)

#Mofidicacion de la pagina Atributos para poner el ancho de las columnas en funcion de la maxima longitud de contenido en la celda
workbook = writer.book
worksheet = writer.sheets['Atributos']
for i in range(len(Atributo.columns)):
    worksheet.set_column(chr(ord('A')+i)+':'+chr(ord('A')+i+1), Atributo[Atributo.columns[i]].map(len).max()+10)

#Mofidicacion de la pagina Indicadores para poner el ancho de las columnas en funcion de la maxima longitud de contenido en la celda
workbook = writer.book
worksheet = writer.sheets['Indicadores']
for i in range(len(metricas.columns)):
    worksheet.set_column(chr(ord('A')+i)+':'+chr(ord('A')+i+1), metricas[metricas.columns[i]].map(len).max()+10)

#Mofidicacion de la pagina Hechos para poner el ancho de las columnas en funcion de la maxima longitud de contenido en la celda
workbook = writer.book
worksheet = writer.sheets['Hechos']
for i in range(len(hechos.columns)):
    worksheet.set_column(chr(ord('A')+i)+':'+chr(ord('A')+i+1), hechos[hechos.columns[i]].map(len).max()+10)

# Mofidicacion de la pagina Filtros para poner el ancho de las columnas en funcion de la maxima longitud de contenido
# en la celda
workbook = writer.book
worksheet = writer.sheets['Filtros']
for i in range(len(filtros.columns)):
    worksheet.set_column(chr(ord('A')+i)+':'+chr(ord('A')+i+1), filtros[filtros.columns[i]].map(len).max()+10)

#Guardar el Excel y almacenar toda la información en el.
writer.save()
print("Glosario realizado con exito")

eleccion=input("Ponga Y/N si quiere exportar en diferentes CSVs : ")

if(eleccion.lower()=="y"):

    print("Exportando a csv")
    hechos.to_csv("csv/Hechos_"+today+".csv",index = None, header=True,encoding='utf-8')
    filtros.to_csv("csv/Filtros_"+today+".csv",index = None, header=True,encoding='utf-8-sig')
    metricas.to_csv("csv/Metricas_"+today+".csv",index = None, header=True,encoding='utf-8-sig')
    Atributo.to_csv("csv/Atributos_"+today+".csv",index = None, header=True,encoding='utf-8-sig')
    print("csv realizados con exito")
