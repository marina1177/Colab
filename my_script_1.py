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
		self.amp = abs(z)
#нормированная разность фаз для каждой частоты
		self.dphase = [] #80/50/20/15
		self.damp = []  #80/50/20/15

		self.dphase15=0
		self.dphase20=0
		self.dphase50=0
		self.dphase80=0


class Cracks(object):
	def __init__(self, zmax=0+0j, crack_type=None,  crack_len=0, crack_w=0, deep=0, freq=0):
		self.type = crack_type
		self.len = crack_len
		self.w = crack_w
		self.deep = deep
		self.freq = freq
		self.phase = math.degrees(cmath.phase(zmax))

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


def fill_dict(ws, start_row):
	mandata = {}
	for row in ws.iter_rows(min_row=start_row, min_col=1, max_row=ws.max_row, max_col=ws.max_column):
		if len(row) > 0:
			freq = row[1].value
			if freq is not None:
				data = [cell.value for cell in row]
				#print(data)
				if freq not in mandata:
					mandata[freq] = []
				mandata[freq].append(data)
	return (mandata)


def fill_refdata(mandata):
	# lambda-функции для сортировки по ключу
	ampl = lambda data: data[4]
	def_disp = lambda data: data[0]

	Calibrate_Curve = []
	for freq in mandata:
		#print(f'Key-frequency {freq}, number str: {len(mandata[freq])}')
		# sorting by Amplitude[V](maxZ)
		mandata[freq].sort(key=ampl, reverse=True)

		# complex number formation for turning
		z = complex(mandata[freq][0][3], mandata[freq][0][2])
		#print(f'original_z = {z}')
		norm = normalize(z)
		print(f'freq = {freq},norm = {norm}')
		clbrt_data = Calibrate100(complex((z * norm).real, (z * norm).imag),
		                          int(mandata[freq][0][1]), norm)
		Calibrate_Curve.append(clbrt_data)
		for row in mandata[freq]:
			z = complex(row[3], row[2])
			row[3] = (z * norm).real
			row[2] = (z * norm).imag
			row[4] = abs(z * norm)

		# sorting by def_disp[mm]
		mandata[freq].sort(key=def_disp, reverse=False)
		#print(mandata[freq])
		print('*******\n')
	return Calibrate_Curve

def main():

	file_path = join(dir, test_files[0])
	wb = load_workbook(filename=file_path, data_only=True, read_only=True)
	wsn = list(wb.sheetnames)
	ws = wb.active
	"""wsdata=None
	for i in wsn:
		if wb[i]['B1'].value == 'freq[kHz]':
			wsdata=i
	if wsdata==None:
		print("Error")"""
	header = [cell.value for cell in next(
			ws.iter_rows(min_row=1, min_col=1, max_col=ws.max_column))]
	print(header)
	wb.close


# обработка сигнала от сквозного отверстия:
# получение нормировочного коэффициента для каждой частоты
# создание листа Calibrate_Curve опорных данных от сквозного отверстия

# составление dict с частотой в качестве ключа и вложенным списком строк данных в качестве значения
# {
# '25' : [[row[0]],[row[1]],..,[row[len(mandata[freq])-1]]],
# '100': [[row[0]],[row[1]],..,[row[len(mandata[freq])-1]]],
# '200': [[row[0]],[row[1]],..,[row[len(mandata[freq])-1]]]
# }
	mandata= fill_dict(ws, start_row=2)
	Calibrate_Curve = fill_refdata(mandata)

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
#*************************************************************

	for i in range(1, len(test_files)):
		name = test_files[i]
		file_path = join(dir, test_files[i])
		print(file_path)
		wb = load_workbook(filename=file_path, data_only=True, read_only=True)
		ws = wb.active
		header = [cell.value for cell in next(
			ws.iter_rows(min_row=1, min_col=1, max_col=ws.max_column))]
		print(header)
		wb.close
		dict=fill_dict(ws=ws, start_row=2)

		ampl = lambda data: data[4]
		def_disp = lambda data: data[0]

		#sort_test(dict, Calibrate_Curve)

		for freq in dict:
			for n in range(0, len(Calibrate_Curve)):
				if Calibrate_Curve[n].freq == freq:
					norm = Calibrate_Curve[n].norm
			print(f'n = {n}, freq = {freq}, norm = {norm}')

			for row in dict[freq]:
				z = complex(row[3], row[2])
				row[3] = (z * norm).real
				row[2] = (z * norm).imag
				row[4] = abs(z * norm)

			# sorting by Amplitude[V](maxZ)
			dict[freq].sort(key=ampl, reverse=True)

			z = complex(dict[freq][0][3], dict[freq][0][2])
			for n in range(0, len(Calibrate_Curve)):
				if Calibrate_Curve[n].freq == freq:
					print(f'n = {n}, i = {i}')
					Calibrate_Curve[n].dphase.append(math.degrees(cmath.phase(z)) - Calibrate_Curve[n].phase)
					Calibrate_Curve[n].damp.append(Calibrate_Curve[n].amp / abs(z))
					print(f'Phase = {math.degrees(cmath.phase(z))}, dphase = {Calibrate_Curve[n].dphase[i-1]}')
					print(f'Amp = {abs(z)}, damp = {Calibrate_Curve[n].damp[i - 1]}')

			# sorting by def_disp[mm]
			dict[freq].sort(key=def_disp, reverse=False)
			print(dict[freq])
			print('*******\n')
			# создание файлов нормированных калибровочных сквозных сигналов
		for freq in dict:
			exname = (name.split('.'))[0] + '_' + str(int(freq)) + 'kHz_norm.xlsx'
			#print(f'new name = {exname} ')
			file_path = join(normdir, exname)
			wb = Workbook()
			ws = wb.active
			ws.append(header)
			for row in dict[freq]:
				ws.append(row)
			# выравнивание колонок:
			for i in range(1, 4):
				zagl = ws.cell(row=1, column=i)
				zagl.alignment = Alignment(horizontal='center')
			wb.save(file_path)
	print('Marina very clever!\n')

if __name__ == "__main__":
	main()

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




