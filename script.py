#!/usr/bin/env python
# -*- coding: utf-8 -*-
import xlsxwriter, random

workbook = xlsxwriter.Workbook('Taller2.xlsx')
worksheet = workbook.add_worksheet()

cellHeaderStyling = workbook.add_format({
	'pattern': 1,
	'bg_color': '#FF5504',
	'align': 'center'
})

cellCenteredStyling = workbook.add_format({
	'align': 'center'
})

headerSetup = [
	{
		'width': 6,
		'text': 'Dia',
		'cellStyling': cellHeaderStyling,
		'colStyling': cellCenteredStyling
	},
	{
		'width': 15,
		'text': 'Inventario Inicial',
		'cellStyling': cellHeaderStyling,
		'colStyling': cellCenteredStyling
	},
	{
		'width': 18,
		'text': '#Numero Aleatorio',
		'cellStyling': cellHeaderStyling,
		'colStyling': None
	},
	{
		'width': 15,
		'text': 'Demanda Diaria',
		'cellStyling': cellHeaderStyling,
		'colStyling': cellCenteredStyling
	},
	{
		'width': 15,
		'text': 'Inventario Final',
		'cellStyling': cellHeaderStyling,
		'colStyling': cellCenteredStyling
	},
	{
		'width': 18,
		'text': '#Numero Aleatorio',
		'cellStyling': cellHeaderStyling,
		'colStyling': None
	},
	{
		'width': 17,
		'text': 'Tiempo de Entrega',
		'cellStyling': cellHeaderStyling,
		'colStyling': cellCenteredStyling
	},
	{
		'width': 18,
		'text': '#Numero Aleatorio',
		'cellStyling': cellHeaderStyling,
		'colStyling': None
	},
	{
		'width': 18,
		'text': '#Tiempo de Espera',
		'cellStyling': cellHeaderStyling,
		'colStyling': cellCenteredStyling
	},
	{
		'width': 8,
		'text': 'Faltante',
		'cellStyling': cellHeaderStyling,
		'colStyling': cellCenteredStyling
	},
	{
		'width': 8,
		'text': 'Orden',
		'cellStyling': cellHeaderStyling,
		'colStyling': cellCenteredStyling
	},
	{
		'width': 8,
		'text': 'Espera',
		'cellStyling': cellHeaderStyling,
		'colStyling': cellCenteredStyling
	},
	{
		'width': 23,
		'text': '',
		'cellStyling': None,
		'colStyling': None
	},
]

for key, headerColSetup in enumerate(headerSetup):
	worksheet.set_column(key, key, headerColSetup.get('width'), headerColSetup.get('colStyling'))
	worksheet.write(0, key, headerColSetup.get('text'), headerColSetup.get('cellStyling'))


demanda_diaria = {
	"25"  : [0   , 0.01999],
	"26"  : [0.02, 0.05999],
	"27"  : [0.06, 0.11999],
	"28"  : [0.12, 0.23999],
	"29"  : [0.24, 0.43999],
	"30"  : [0.44, 0.67999],
	"31"  : [0.68, 0.82999],
	"32"  : [0.83, 0.92999],
	"33"  : [0.93, 0.97999],
	"34"  : [0.98, 0.99999],  
}
tiempo_entrega = {
	"1"   : [0   , 0.19999],
	"2"   : [0.20, 0.49999],
	"3"   : [0.50, 0.74999],
	"4"   : [0.75, 0.99999], 
}
tiempo_espera = {
	"0"	  : [0   , 0.39999],
	"1"   : [0.4 , 0.59999],
	"2"   : [0.6 , 0.74999],
	"3"   : [0.75, 0.89999],
	"4"   : [0.90, 0.99999],
}


inventarioInicial = 100
tiempoEntrega = -1
orden = tiempoEspera = dacumulada = cclienteE = cclienteN = 0
dd_espera = []

for i in xrange(1, 261):
	if tiempoEntrega > 0:
		tiempoEntrega = tiempoEntrega -1
	elif tiempoEntrega == 0:
		if(inventarioInicial + 100 - dacumulada >= 0):
			inventarioInicial = inventarioInicial + 100 - dacumulada
			dacumulada = 0
		else:
			inventarioInicial = inventarioInicial + 100
			x = []
			while not dd_espera == []:
				a = dd_espera.pop()
				if inventarioInicial >= a:					
					inventarioInicial = inventarioInicial - a
					dacumulada = dacumulada - a
				else:
					x.append(a)
			
			dd_espera = x				
		tiempoEntrega = -1
		dacumulada = 0
	
	worksheet.write(i, 0, i)
	worksheet.write(i, 1, inventarioInicial)
	n1 = random.random()
	worksheet.write(i, 2, n1)
	dd = 0

	for j in demanda_diaria:
		if n1 >= demanda_diaria[j][0]  and  n1 <= demanda_diaria[j][1]: 
			worksheet.write(i, 3, j)
			dd = int(j)
	if(inventarioInicial >= dd) and (inventarioInicial >= 30):
		inventarioInicial = inventarioInicial - dd
		worksheet.write(i, 4, inventarioInicial)
	else:
		if(tiempoEntrega < 0):
			n2 = random.random()
			worksheet.write(i, 5, n2)
			for j in tiempo_entrega:
				if n2 >= tiempo_entrega[j][0] and n2 <= tiempo_entrega[j][1]:
					worksheet.write(i, 6, j)
					tiempoEntrega = int(j)
					orden = orden +1
					worksheet.write(i, 10, orden)
	if(inventarioInicial < dd) or (tiempoEntrega > 0):
		n3 = random.random()
		worksheet.write(i, 7, n3)
		for j in tiempo_espera:
			if n3 >= tiempo_espera[j][0] and n3 <= tiempo_espera[j][1]:
				tiempoEspera = int(j)
				worksheet.write(i, 8, tiempoEspera)
				if tiempoEspera >= tiempoEntrega:
					worksheet.write(i, 11, 'SI')
					dacumulada = dacumulada + dd
					worksheet.write(i, 9, dacumulada)
					dd_espera.append(dd)
					cclienteE = cclienteE + 1
				else:
					worksheet.write(i, 11, 'NO')
					cclienteN = cclienteN + 1

worksheet.write('M3','Costo de Ordenar')
worksheet.write('M4',orden*100)		
worksheet.write('M5','Costo de Inventario')
worksheet.write('M6',float(52 * 260)/360.0)	
worksheet.write('M7','Costo Cliente Espera')
worksheet.write('M8',20 * cclienteE)	
worksheet.write('M9','Costo Cliente No Espera')
worksheet.write('M10',50 * cclienteN)	

green_format = workbook.add_format({
	'pattern': 1,
	'bg_color': '#00FF00'
})

worksheet.write('M11','COSTO TOTAL', green_format)
costoTotal = orden * 100 + float(52 * 260/360.0) + (20 * cclienteE) + (50 * cclienteN)
worksheet.write('M12', costoTotal, green_format)		
		
		
		
	
	

workbook.close()
