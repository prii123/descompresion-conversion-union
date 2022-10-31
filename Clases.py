#importacion para manejar archivos comprimidos ZIP
import zipfile
#importacion para manejar archivos XML
from xml.etree import ElementTree
import pandas as pd

import sys
import os
# variable para reconocer la ruta de la carpeta en la que se esta trabajando
rutaRaiz = os.path.dirname(os.path.abspath(__file__))

class Descompresion_ZIP:
    def __init__(self, passwrod=None):
        # se inicializa variables necesarias
        self.ruta_inicio = rutaRaiz
        self.password = passwrod



    def descompresion(self):
         arvhiso_para_descomprimir = self.listar_archivos_para_procesar()
         if(len(arvhiso_para_descomprimir)>0):
             for nombre_archivo_zip in arvhiso_para_descomprimir:

                 # se utiliza la libreria y se le pasa como parametro la ruta
                 archivo_zip = zipfile.ZipFile('{0}\entradas\{1}'.format(self.ruta_inicio,nombre_archivo_zip), "r")
                 try:
                    #print(archivo_zip.namelist())
                    # se utiliza la libreria y se pasa el parametro de password por si esta protegido el archivo y la ruta de salida
                    archivo_zip.extractall(pwd=self.password, path=self.ruta_inicio + '\salidas')
                 except:
                    print('error en la descompresion')
                 archivo_zip.close()

    def listar_archivos_para_procesar(self):
        try:
            self.ruta_entrada = rutaRaiz + '\entradas'
            self.contenido = os.listdir(self.ruta_entrada)
            # matriz inicial en la que se almacenara los datos de los archivos xml leidos
            matriz = []
            # ciclo para listar todos los archivos que se encuentra dentro de la carpeta salidas
            for fichero in self.contenido:
                  # condicional para reconocer todos los archivos de excel con el formato .xlsx
                 if fichero.endswith('.zip'):
                 # si pasa el ocndicional significa que es un archivo de excel y sera leido
                    matriz.append(fichero)
            return matriz
        except Exception as ex:
            print(ex)


#test = Descompresion_ZIP(rutaRaiz)
#test.descompresion()

class Conversion_XML:
    def __init__(self):
        self.ruta_archivo=rutaRaiz


    def conversion(self):
        documentos_xml = self.listar_archivos_para_procesar()
        if(len(documentos_xml)>0):
            for archivo_xml_base_conversion in documentos_xml:
                ruta_de_archivo =  '{0}\salidas\{1}'.format(self.ruta_archivo,archivo_xml_base_conversion)
                print(ruta_de_archivo)
                tree = ElementTree.parse(ruta_de_archivo)
                root = tree.getroot()
                A=[]
                try:
                    for ele in root:
                        B={}
                        for i in list(ele):
                            #if(i.tag.split('}')[1] != '' or i.text != ''):
                                #para iterar cada elemento de la factura
                            for ii in i:
                                if(ii.tag.split('}')[1] == 'TaxAmount' or ii.tag.split('}')[1] == 'PriceAmount'
                                        or ii.tag.split('}')[1] == 'Amount' or ii.tag.split('}')[1] == 'BaseAmount'
                                        or ii.tag.split('}')[1] == 'Description' or ii.tag.split('}')[1] == 'BaseQuantity'):
                                    #para iterar cada elemento de de los productos de la factura
                                    B.update({'cufe': archivo_xml_base_conversion.split('.')[0], ii.tag.split('}')[1]: ii.text})
                                    A.append(B)

                    df = pd.DataFrame(A)
                    df.drop_duplicates(keep='first', inplace=True)

                    df.reset_index(drop=True, inplace=True)
                    #print(df)
                    #writer = pd.ExcelWriter('formatoo.xlsx')
                    df.to_excel('{0}\salidas\{1}.xlsx'.format(self.ruta_archivo,archivo_xml_base_conversion.split('.')[0]), sheet_name='Hoja1')
                except Exception:
                    print('error en la conversion')

    def listar_archivos_para_procesar(self):
        try:
            self.ruta_salida = rutaRaiz + '\salidas'
            self.contenido = os.listdir(self.ruta_salida)
            # matriz inicial en la que se almacenara los datos de los archivos xml leidos
            matriz = []
            # ciclo para listar todos los archivos que se encuentra dentro de la carpeta salidas
            for fichero in self.contenido:
                  # condicional para reconocer todos los archivos de excel con el formato .xlsx
                 if fichero.endswith('.xml'):
                 # si pasa el ocndicional significa que es un archivo de excel y sera leido
                    matriz.append(fichero)
            return matriz
        except Exception as ex:
            print(ex)

#conv = Conversion_XML(rutaRaiz)
#conv.conversion()


class Unir_archivos:
    def __init__(self):
        self.ruta_salida = rutaRaiz + '\salidas'
        self.contenido = os.listdir(self.ruta_salida)

    def unir_archivos(self):
        try:
            # matriz inicial en la que se almacenara los datos de los archivos excel leidos
            matriz = []
            # ciclo para listar todos los archivos que se encuentra dentro de la carpeta salidas
            for fichero in self.contenido:
                # condicional para reconocer todos los archivos de excel con el formato .xlsx
                if fichero.endswith('.xlsx'):
                    # si pasa el ocndicional significa que es un archivo de excel y sera leido
                    df1 = pd.read_excel(self.ruta_salida+'\{0}'.format(fichero))
                    # una vex leido se agregara a la variable matriz
                    matriz.append(df1)
            # se concatenara toda la matriz
            join = pd.concat(matriz)
            # la matriz unida se escribira en un excel y se exportara con el nombre de Union_de_Facturas
            join.to_excel("Union_de_Facturas.xlsx")
        except Exception as ex:
            print(ex)


#test = Unir_archivos()
#test.unir_archivos()









