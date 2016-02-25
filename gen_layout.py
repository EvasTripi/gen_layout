#!/usr/bin/python
# -*- coding: utf-8 -*-
from __future__ import unicode_literals, print_function

import sys, codecs
from collections import namedtuple
import click
from xlsxwriter import Workbook

Registros = namedtuple('Registros', 'etiqueta nombre inicio longitud tipo_cap tipo_bd respuestas')
Registros2 = namedtuple('Registros', 'etiqueta nombre inicio longitud tipo_cap')


def visualiza(reg_trans):
    for registro in reg_trans:
        print(registro.etiqueta, registro.nombre, registro.inicio,
              registro.longitud, registro.tipo_cap)


def leer_lst(entrada):
    with codecs.open(entrada, "r", encoding='utf-8-sig') as arch:
        lineas = arch.read().split('\r\n')
        registros = []
        for linea in lineas[18:-4]:
            """
            Datos de cada pregunta
            """
            titulo = linea[0:47].strip()
            nombre = linea[47:70].strip()
            inicio = linea[70:75].strip()
            longitud = linea[75:80].strip()
            tipo_entrada = linea[80:85].strip()
            tipo_item = linea[85:90].strip()
            respuestas = linea[90:95].strip()
            registros.append(Registros(titulo, nombre, inicio, longitud, tipo_entrada, tipo_item, respuestas))
    print('Leer archivo lst finalizado')

    reg_trans = []
    print('Transformando')
    res_sub = None
    num_ini = None
    for reg, registro in enumerate(registros):
        try:
            sig_reg = registros[reg + 1]
        except IndexError:
            pass

        # Inicializar variables para los subitems
        if sig_reg.tipo_bd == 'Sub' and registro.tipo_bd == 'I':
            res_sub = registro.respuestas
            num_ini = reg + 1
            nuevo_inicio = int(registro.inicio)


        # Recorrer los registros con un rango para sacar los subitems
        elif registro.tipo_bd == 'Sub' and sig_reg.tipo_bd == 'I':
            num_fin = reg
            # Recorrer ene veces por cada respuesta del subitem
            for resp_sub in range(1, int(res_sub) + 1):
                for reg_sub, registro_sub in enumerate(registros[num_ini:num_fin + 1], 1):
                    nueva_eti = '{} atributo {}'.format(registro_sub.etiqueta, resp_sub)

                    # Almacenar en nueva estructura
                    reg_trans.append(Registros2(nueva_eti, registro_sub.nombre, nuevo_inicio,
                                                registro_sub.longitud, registro_sub.tipo_cap))
                    nuevo_inicio += int(registro_sub.longitud)
        elif registro.tipo_bd == 'Sub':
            pass
        else:
            # if registro.tipo_bd == 'Sub':
            if int(registro.respuestas) > 1:
                nuevo_inicio = int(registro.inicio)
                for num, respuesta in enumerate(range(int(registro.respuestas)), 1):
                    nueva_eti = '{} {}{}'.format(registro.etiqueta, 'mención ', num)
                    # Almacenar en nueva estructura
                    if num == 1:
                        reg_trans.append(Registros2(nueva_eti, registro.nombre, registro.inicio, registro.longitud,
                                                    registro.tipo_cap))
                    else:
                        nuevo_inicio += int(registro.longitud)
                        reg_trans.append(Registros2(nueva_eti, registro.nombre, nuevo_inicio, registro.longitud,
                                                    registro.tipo_cap))

            else:
                reg_trans.append(Registros2(registro.etiqueta, registro.nombre, registro.inicio, registro.longitud,
                                            registro.tipo_cap))


    # visualiza(reg_trans)
    return reg_trans


def generar_layout(datos, salida):
    wb = Workbook(salida)
    ws = wb.add_worksheet('Layout')
    encabezados = '''
Nombre
Inicio
Longitud
Tipo de dato capturado
    '''.split('\n')[1:-1]

    # Escribir encabezados
    col_enc_bg = "#{:02x}{:02x}{:02x}".format(15, 36, 62).upper()
    col_ren2 = "#{:02x}{:02x}{:02x}".format(220, 230, 241).upper()
    format_enc = wb.add_format({'bold': True, 'font_color': 'white', 'bg_color': col_enc_bg})
    format_ren1 = wb.add_format({'border': 1})
    format_ren2 = wb.add_format({'border': 1, 'bg_color': col_ren2})
    for col, encabezado in enumerate(encabezados):
        ws.write(0, col, encabezado, format_enc)

    # Escribir datos del diccionario    
    for renglon, registro in enumerate(datos, 1):
        formato = format_ren1 if renglon % 2 == 0 else format_ren2
        # Registros2 = namedtuple('Registros', 'etiqueta nombre inicio longitud tipo_cap')
        ws.write(renglon, 0, registro.etiqueta, formato)
        ws.write(renglon, 1, int(registro.inicio), formato)
        ws.write(renglon, 2, int(registro.longitud), formato)
        ws.write(renglon, 3, registro.tipo_cap, formato)


    # Aplicando formato a la hoja
    ws.freeze_panes(1, 0)
    ws.autofilter('A1:D1')
    ws.set_column(0, 0, 55)
    ws.set_column(1, 1, 8)
    ws.set_column(2, 2, 10)
    ws.set_column(3, 3, 25)
    ws.hide_gridlines(2)
    wb.close()
    click.launch(salida)

    print("Layout generado")


if __name__ == "__main__":

    if len(sys.argv) != 2:
        print('Especificar el nombre del archivo lst, ejemplo: "001-14 Estudio.lst"')
        sys.exit(1)
    else:
        entrada = sys.argv[1]
        salida = 'Layout ' + sys.argv[1][:-4] + '.xlsx'
        # generar_layout_con_lst(entrada, salida)
        datos = leer_lst(entrada)
        generar_layout(datos, salida)
