#!/usr/bin/env python
# -*- coding: utf-8 -*-
import xlsxwriter
from random import *

workbook = xlsxwriter.Workbook('Taller2.xlsx')
worksheet = workbook.add_worksheet()

formatc = workbook.add_format()
formatc.set_pattern(1)  # This is optional when using a solid fill.
formatc.set_bg_color('#FF5504')

worksheet.write('A1','dia',formatc)
worksheet.write('B1','Inventario Inicial',formatc)
worksheet.write('C1','#Numero Aleatorio',formatc)
worksheet.write('D1','Demanda Diaria',formatc)
worksheet.write('E1','Inventario Final',formatc)
worksheet.write('F1','#Numero Aleatorio',formatc)
worksheet.write('G1','Tiempo de Entrega',formatc)
worksheet.write('H1','#Numero Aleatorio',formatc)
worksheet.write('I1','#Tiempo de Espera',formatc)
worksheet.write('J1','Faltante',formatc)
worksheet.write('K1','Orden',formatc)
worksheet.write('L1','Espera',formatc)



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
te = -1
orden = 0
ted = 0
dacumulada = 0
dd_espera = list()
cclienteE = 0
cclienteN = 0
for i in range(2,262):
	if te > 0:
		te = te -1
	elif te == 0:
		if(inventarioInicial + 100 - dacumulada >= 0):
			inventarioInicial = inventarioInicial + 100 - dacumulada
			dacumulada = 0
		else:
			inventarioInicial = inventarioInicial + 100
			x = list()
			while not dd_espera == []:
				a = dd_espera.pop()
				if inventarioInicial >= a:					
					inventarioInicial = inventarioInicial - a
					dacumulada = dacumulada - a
				else:
					x.append(a)
			
			dd_espera = x				
		te = -1
		dacumulada = 0
	worksheet.write('A%d' % i, i)
	worksheet.write('B%d' % i, inventarioInicial)
	n1 = random()
	worksheet.write('C%d' % i, n1)
	dd = 0
	for j in demanda_diaria:
		if n1 >= demanda_diaria[j][0]  and  n1 <= demanda_diaria[j][1]: 
			worksheet.write('D%d' % i,j)
			dd = int(j)
	if(inventarioInicial >= dd) and (inventarioInicial >= 30):
		inventarioInicial = inventarioInicial - dd
		worksheet.write('E%d' % i,inventarioInicial)
	else:
		if(te < 0):
			n2 = random()
			worksheet.write('F%d' %i,n2)
			for j in tiempo_entrega:
				if n2 >= tiempo_entrega[j][0] and n2 <= tiempo_entrega[j][1]:
					worksheet.write('G%d'%i,j)
					te = int(j)
					orden = orden +1
					worksheet.write('K%d' % i,orden)
	if(inventarioInicial < dd) or (te > 0):
		n3 = random()
		worksheet.write('H%d' %i,n3)
		for j in tiempo_espera:
			if n3 >= tiempo_espera[j][0] and n3 <= tiempo_espera[j][1]:
				ted = int(j)
				worksheet.write('I%d' %i,ted)
				if ted >= te:
					worksheet.write('L%d' %i,'SI')
					dacumulada = dacumulada + dd
					worksheet.write('J%d' %i,dacumulada)
					dd_espera.append(dd)
					cclienteE = cclienteE + 1
				else:
					worksheet.write('L%d' %i,'NO')
					cclienteN = cclienteN +1

worksheet.write('M3','Costo de Ordenar')
worksheet.write('M4',orden*100)		
worksheet.write('M5','Costo de Inventario')
worksheet.write('M6',float(52*260)/360.0)	
worksheet.write('M7','Costo Cliente Espera')
worksheet.write('M8',20*cclienteE)	
worksheet.write('M9','Costo Cliente No Espera')
worksheet.write('M10',50*cclienteN)	

green_format = workbook.add_format()
green_format.set_pattern(1) 
green_format.set_bg_color('#FF0000')
	
worksheet.write('M11','COSTO TOTAL', green_format)
worksheet.write('M12',orden*100 + float(52*260)/360.0 + 20*cclienteE + 50*cclienteN, green_format)		
		
		
		
	
	

workbook.close()
