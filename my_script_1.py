import pandas as pd
import matplotlib.pyplot as plt
from openpyxl import load_workbook, Workbook
from  openpyxl.styles import Alignment, PatternFill, Font
#import  matplotlib as plt

# Стандартное импортирование plotly
#import plotly.plotly as py
#import plotly.graph_objs as go
#from plotly.offline import iplot

# Использование cufflinks в офлайн-режиме
import cufflinks
cufflinks.go_offline()
import plotly as pl
import numpy as np
import math

class Calibrate100(object):
	def __init__(self, z=0+0j, freq=0):

		self.deep = 1
		self.freq = freq
		self.Amp = math.sqrt((z.image**2) + (z.real**2))
		self.Phase=math.tan(z.image/z.real)

	#solved Amp*norm*exp(-i*Phase)=etaloneAmp*exp(-i*etalonePhase)
		self.norm = 0+0j #на это буду домножать каждый сигнал этой частоты
		self.z_norm = 0+0j
		self.phase_norm=0
		self.ampl_norm=0

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





#def main():

file = './rect/test_deep100.xlsx'
wb = load_workbook(file, data_only=True, read_only=True)
wsn = list(wb.sheetnames)
print(wsn)

wsdata=None
for i in wsn:
	if wb[i]['B1'].value == 'freq[kHz]':
		wsdata=i
if wsdata==None:
	print("Error")

ws = wb[wsdata]
print(ws['B1'].value)
header = [cell.value for cell in next(
ws.iter_rows(min_row=1,min_col=1,max_col=ws.max_column))]

print(header)
mandata={}

for row in ws.iter_rows(min_row=2,min_col=1,max_row=ws.max_row,max_col=ws.max_column):
	if len(row)>0:
		freq = row[1].value
		if freq is not None:
			data = [cell.value for cell in row]
			#print(data)
			if freq not in mandata:
				mandata[freq] = []
			mandata[freq].append(data)


ampl= lambda data:data[4]

for freq in mandata:
	print(f'Generator{freq}, number str: {len(mandata[freq])}')
	print(mandata[freq])
	mandata[freq].sort(key=ampl, reverse=True)
	print(mandata[freq])

	
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

