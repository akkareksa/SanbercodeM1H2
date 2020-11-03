# -*- coding: utf-8 -*-
"""
Created on Tue Nov  3 13:02:58 2020

@author: Kevin
"""

from openpyxl import Workbook
from openpyxl.chart import BarChart, Series, Reference 
import csv
import pandas as pd
from io import StringIO

wb = Workbook()
ws = wb.active

#openfile
data_luas = pd.read_csv('luas-wilayah-menurut-kecamatan-di-kota-bandung-2017.csv')
data_penduduk = pd.read_csv('jumlah-penduduk-kota-bandung.csv')

#sebelum digabung lakukan normalisasi data, karena ada beberapa data yang tidak bisa dijoin karena namanya berbeda
data_combined = data_luas.merge(data_penduduk, how='outer', left_on='Nama Kecamatan', right_on='Kecamatan')

#hapus kolom yang tidak digunakan
del data_combined['Kecamatan']
del data_combined['Jumlah_Kelurahan']
data_combined.to_csv('combined.csv')

buffer = StringIO()  #creating an empty buffer
data_combined.to_csv(buffer)  #filling that buffer
buffer.seek(0) #set to the start of the stream

index = 0
for row in csv.reader(buffer):
    data_clean =  []
    for i  in row:
        try:
            i = int(i)
        except:
            pass
        data_clean.append(i)
    try:
        jumlah_penduduk = data_clean[3]
        luas_wilayah_per100 = data_clean[2]/100
        kepadatan = jumlah_penduduk/luas_wilayah_per100
        data_clean.append(float(kepadatan))
    except:
        pass
    ws.append(data_clean)
    index +=1 
len_row = len(data_clean)

#beri judul untuk E1
ws['E1'] = "Kepadatan Penduduk"

chart1 = BarChart()
chart1.type = "col"
chart1.style = 3
chart1.title = "Bar Chart"
chart1.y_axis.title = "Kepadatan per 100m2"
chart1.x_axis.title = "Kecamatan"

data = Reference(ws, min_col=5, min_row=1, max_row=index, max_col=5)
cats = Reference(ws, min_col=2, min_row=2, max_row = index)
chart1.height = 10
chart1.width = 30
chart1.add_data(data, titles_from_data=True)
chart1.set_categories(cats)
ws.add_chart(chart1,"G2")

wb.save("barPenduduk.xlsx")