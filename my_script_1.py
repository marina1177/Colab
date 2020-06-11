import pandas as pd
import matplotlib.pyplot as plt
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

file = 'rect/rect_5_0.4_40_20deep.xlsx'
movies = pd.read_excel(file)

print(movies.head())
sorted_tab = movies.sort_values(by=['freq[kHz]', 'deep', 'def_disp[mm]'], ascending=True)
print(sorted_tab.head(33))

interpolated =sorted_tab.head(33).interpolate(method='polynomial', order=5)
interpolated.plot('Re[V]', 'Im[V]')
#sorted_tab.head(62).plot('Re[V]','Im[V]', kind='density')

plt.show()
#print(movies.tail())
print(movies.shape)
print('Marina very clever!\n')
