# -*- coding: utf-8 -*-

from odoo import models, fields, api

from io import BytesIO
from xlrd import open_workbook
import base64
import datetime
import json
import os

class open_academy(models.Model):
    _name = 'open_academy.open_academy'
    _description = 'open_academy.open_academy'

    name = fields.Char()
    #name2 = fields.Char()
    #value = fields.Integer()
    #value2 = fields.Text()
    #description = fields.Text()
    xls = fields.Binary(string='file', attachment=True, help='Upload the excel file', required=True)
    #txt = fields.Binary(string='txt_file')

    """
    @api.depends('value')
    def _value_pc(self):
        for record in self:
            record.value2 = float(record.value) / 100
    """

    # metodo para el boton transformar de la vista listado.
    # obtiene el resultado de procesar el excel con las reglas 

    def trans(self):

        #arreglo total que contendra texto por cada final del excel
        txt_data = []

        # ruta para acceder archivo de reglas
        path = os.getcwd()
        new_path = path.replace(os.sep, '/')
        rule_path_json = new_path + "/../server/odoo/addons/user/open_academy/rules/reglas.json"
                
        
        # extrae informacion del archivo json con las reglas
        with open(rule_path_json) as reglas:
            datos_reglas = json.load(reglas)

        # extraccion en variables de las regla anio
        anio_json_type = datos_reglas[0]['tipo']
        anio_json_size = int(datos_reglas[0]['TAMANO'])

        # extraccion en variables de las regla concepto
        concepto_json_type = datos_reglas[1]['tipo']
        concepto_json_size = int(datos_reglas[1]['TAMANO'])

        # extraccion en variables de las regla valor
        valor_json_type = datos_reglas[2]['tipo']
        valor_json_size = int(datos_reglas[2]['TAMANO'])
        

        for record in self:
            
            # pasos para decodificar archivo de bytes del campo xls        
            inputx = BytesIO()
            inputx.write(base64.decodestring(self.xls))
            book = open_workbook(file_contents=inputx.getvalue())
            
            #acceder a la hoja inicial del archivo de excel
            sh = book.sheet_by_index(0)
            
            #variable cont para evitar la primera fila que son los titulos en bucle for
            cont = True

            #ingresa fila por fila para verificar cada valor de las celdas anio, concepto y valor
            for rx in range(sh.nrows):
                
                if cont == True:
                    cont = False
                    continue

                anio_txt = ''
                concepto_txt = ''
                valor_txt = ''

                anio_txt = validations_anio_valor(sh.row(rx)[0].value, anio_json_type, anio_json_size)
                concepto_txt = validations_concepto(sh.row(rx)[1].value, concepto_json_type, concepto_json_size)
                valor_txt = validations_anio_valor(sh.row(rx)[2].value, valor_json_type, valor_json_size)

                txt_data.append(anio_txt + concepto_txt + valor_txt)

        
        
        # ruta para acceder archivo de reglas
        path = os.getcwd()
        new_path = path.replace(os.sep, '/')
        current_date = datetime.datetime.now()
        cyear = str(current_date.year) + "_"
        cmonth = str(current_date.month) + "_"
        cday = str(current_date.day) + "_"
        chour = str(current_date.hour) + "_"
        cmin = str(current_date.minute) + "_"
        csec = str(current_date.second) 
        filename = "archivo_final_" + cyear + cmonth + cday + chour + cmin + csec + ".txt"  
        txt_path = new_path + "/../server/odoo/addons/user/open_academy/static/" + filename

        # se abre archivo txt para escribir la informacion del array
        with open(txt_path, 'wt') as final_text:
            for n in txt_data:
                final_text.write(n + '\n') 
            
        return {
            "type": "ir.actions.act_url",
            "url": str('/open_academy/static/' + filename + '?download=true'),
            "target": "new",
            "tag": "reload",
        }
    
         


def validations_anio_valor(value, json_type, json_size):
    if not type(value) == str:
        value = int(value)
    if(str(value) == 'NaN' or str(value) == 'None'):
        return ''        
    elif(len(str(value)) < json_size):
        return ('0' * (json_size - len(str(value))) + str(value))
    else:
        return str(value)

def validations_concepto(value, json_type, json_size):
    if not type(value) == str:
        value = int(value)
    if(str(value) == 'NaN' or str(value) == 'None'):
        return ''        
    elif(len(str(value)) < json_size):
        return (str(value) + '$' * (json_size - len(str(value))))
    else:
        return str(value)