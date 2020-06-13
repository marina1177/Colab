#!/usr/bin/env python3
import pandas as pd
import matplotlib.pyplot as plt
from openpyxl import load_workbook, Workbook
from  openpyxl.styles import Alignment, PatternFill, Font
import numpy as np
import math
import cmath
from scipy.optimize import root
from os.path import  join, abspath
#import  matplotlib as plt

# Стандартное импортирование plotly
#import plotly.plotly as py
#import plotly.graph_objs as go
#from plotly.offline import iplot

# Использование cufflinks в офлайн-режиме
import cufflinks
cufflinks.go_offline()
import plotly as pl

cracks_deep = [100, 80, 60, 40, 20]
cracks_files = []

test_deep = [100, 80, 50, 20, 15]
test_files = ['test_deep100.xlsx', 'test_deep80.xlsx', 'test_deep50.xlsx' ]
dir = './rect'
normdir = './norm'

class Calibrate100(object):
	def __init__(self, z=0+0j, freq=0, norm=0+0j):
		self.freq = freq
		self.norm = norm  # на это буду домножать каждый сигнал этой частоты
		self.phase = math.degrees(cmath.phase(z))

	#нормированная разность фаз для каждой частоты
		self.dphase15=0
		self.dphase20=0
		self.dphase50=0
		self.dphase80=0

class Cracks(object):
	def __init__(self,type=None, lengh=0, width=0, deep=0, freq=0):
		self.type=type
		self.l = lengh
		self.w=width
		self.deep =deep
		self.freq = freq

		self.im_cmpl = 1
		self.re_cmpl = 1

def fun(x):
	return [x[0]**2 + x[1]**2 - 100.0,
		x[0]*math.tan(math.radians(40))-x[1]]


def normalize(z):
	A = abs(z)
	fi = math.degrees(cmath.phase(z))
	print(f'normalize:\nA_orig = {A}, fi_orig = {fi}')

	#поиск im/re, соответсвующих амплитуде 10В и фазе 40 градусов
	sol=root(fun, [0, 0])
	z_et = -1*complex(sol.x[0], sol.x[1])
	#norm - нормировочный коэффицент
	norm = z_et/z

	print(f'z_norm = {z*norm}\nAnorm = {abs(z*norm)}, fi = {math.degrees(cmath.phase(z*norm))}')
	#вернуть нормировочный коэффициент
	return(norm)


#def main():

file_path = join(dir, test_files[0])
wb = load_workbook(filename=file_path, data_only=True, read_only=True)
wsn = list(wb.sheetnames)

wsdata=None
for i in wsn:
	if wb[i]['B1'].value == 'freq[kHz]':
		wsdata=i
if wsdata==None:
	print("Error")

ws = wb[wsdata]

print(ws['B1'].value)
header = [cell.value for cell in next(
		ws.iter_rows(min_row=1, min_col=1, max_col=ws.max_column))]
print(header)


wb.close

# обработка сигнала от сквозного отверстия:
# получение нормировочного коэффициента для каждой частоты
# создание листа Calibrate_Curve опорных данных от сквозного отверстия

# составление dict с частотой в качестве ключа и вложенным списком строк данных в качестве значения
# {
# '25' : [[row[0]],[row[1]],..,[row[len(mandata[freq])]]],
# '100': [[row[0]],[row[1]],..,[row[len(mandata[freq])]]],
# '200': [[row[0]],[row[1]],..,[row[len(mandata[freq])]]]
# }
mandata={}
for row in ws.iter_rows(min_row=2, min_col=1, max_row=ws.max_row, max_col=ws.max_column):
	if len(row)>0:
		freq = row[1].value
		if freq is not None:
			data = [cell.value for cell in row]
			#print(data)
			if freq not in mandata:
				mandata[freq] = []
			mandata[freq].append(data)

#lambda-функции для сортировки по ключу
ampl= lambda data:data[4]
def_disp= lambda data:data[0]

Calibrate_Curve = []
for freq in mandata:
	print(f'Key-frequency {freq}, number str: {len(mandata[freq])}')

	#sorting by Amplitude[V](maxZ)
	mandata[freq].sort(key=ampl, reverse=True)

	#complex number formation for turning
	z = complex(mandata[freq][0][3], mandata[freq][0][2])
	print(f'original_z = {z}')
	norm = normalize(z)
	clbrt_data = Calibrate100(complex((z*norm).real, (z*norm).imag),
	                          int(mandata[freq][0][1]), norm)
	Calibrate_Curve.append(clbrt_data)

	for row in mandata[freq]:
		z = complex(row[3], row[2])
		row[3] = (z*norm).real
		row[2] = (z*norm).imag
		row[4] = abs(z*norm)

	# sorting by def_disp[mm]
	mandata[freq].sort(key=def_disp, reverse=False)
	print(mandata[freq])

print('*******\n')

# создание файлов нормированных калибровочных сквозных сигналов
for freq in mandata:
	exname = 'test_deep100_' + str(int(freq)) + 'kHz_norm.xlsx'
	file_path = join(normdir, exname)
	wb = Workbook()
	ws = wb.active

	ws.append(header)
	for row in mandata[freq]:
		ws.append(row)
#выравнивание колонок:
	for i in range(1, 4):
		zagl = ws.cell(row = 1, column=i)
		zagl.alignment = Alignment(horizontal='center')
	wb.save(file_path)

'''movies = pd.read_excel(file)
print(movies.head())
#sorted_tab = movies.sort_values(by=['freq[kHz]', 'deep', 'def_disp[mm]'], ascending=True)
sorted_tab = movies.sort_values(by=['freq[kHz]', 'def_disp[mm]'], ascending=True)
print(sorted_tab.head())
#print(sorted_tab[lambda x: x['freq[kHz]'] == 25])

#print(sorted_tab.index)
#print(sorted_tab.columns)
print(sorted_tab.values)

#for i in range (sorted_tab.index):
	#if (sorted_tab[lambda x: x['freq[kHz]'] == 25]):

#print(sorted_tab['freq[kHz]'].head())

#interpolated =sorted_tab.head(33).interpolate(method='polynomial', order=5)
#interpolated.plot('Re[V]', 'Im[V]')
#sorted_tab.head(62).plot('Re[V]','Im[V]', kind='density')
#plt.show()
print(movies.shape)'''


print('Marina very clever!\n')

