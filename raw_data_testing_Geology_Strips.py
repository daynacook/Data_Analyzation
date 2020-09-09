import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import datetime as dt

excel_file = 'D:\Data analyzation folder\Excell_data_files\Geology Meter Power Strips.xlsx'

data = pd.read_excel(excel_file, 
    sheet_name='Post Power Strips Holiday Break', header=[20, 21])

dataHeaders = pd.concat(data.values(), axis=0)

# data.set_index([('* Start', '* Date'),
#                 ('Start', 'Time'),
#                 ('VAt', 'Avg')], inplace=True)

#print(data.index.get_level_values(2))

print(dataHeaders)

#data['kWh'] = data[('VAt', 'Avg')] / 1000

# watData = data.loc[:, ['kWh']]

# print(data[:, ('VAt Avg')])

#Multi indexing is probably the solution for easier reading.