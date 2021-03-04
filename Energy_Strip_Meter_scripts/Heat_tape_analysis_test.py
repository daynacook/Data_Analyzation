import pandas as pd
import numpy as np
import matplotlib.pyplot as plot
import xlrd

# this is a test file that aims to understand python's ability to grab specific data in an excell sheet and display it 
# in a neat fashion using graphing as well as understanding how to use python to grab data and manipulate it to do 
# calculations

# path to location of Excel file to be read
excel_file = 'D:\Data analyzation folder\Excell_data_files\SBS W Heat tape 2020_LOG.xlsx'

# df means Data Frame, which represents the frames (or cells) of data in an excell sheet
df_read_file = pd.read_excel(excel_file, sheet_name='Analysis')

# Prints the kWh/5min column in the sheet
#print(df_read_file['kWh/5min'])

# This serves to group the sums of kWh/5min by month and list the total sum of each month 
# and will print the sum total of all the kWh data
watSum = df_read_file.groupby(['* Start']).sum()
watSumTotal = watSum.sum()
watSumPrint = watSum.loc[:,'kWh/5min']

# Groups the data so that it only includes the dates and kWh data,
# used to programatically create the x and y axes of the graph.
watData = df_read_file.loc[:, ['* Start', 'kWh/5min']]

# Optional command that converts watData to datetime to change x axis to months in the year
# instead of number of number of intervals
#watData['* Start'] = pd.to_datetime(watData['* Start'], errors='coerce')
#watData = watData.dropna(subset=['* Start'])

# Configuring/formatting the kWh/5min graph
watData.plot(xlabel = 'Time', ylabel = 'kWh/5min')
# Setting the lower/upper bounds of the y axis
plot.ylim(0, .4)
# getting the lower/upper bounds of the x axis for use in formatting
x_min, x_max = plot.xlim()
# increasing x tick frequency and label rotation for better readability
plot.xticks(np.arange(0, x_max, 250), rotation=45)
plot.savefig("Kwh_5min_heat_tape")

# printing the month sums and overall sum of wat data
print(watSumPrint)
print(watSumTotal.loc['kWh/5min'])

#with pd.ExcelWriter("heat_tape_analysis.xlsx", engine="xlsxwriter") as writer:
#    df_read_file.to_excel(writer, sheet_name="sheet 1")
#    worksheet = writer.sheets['sheet 1']
#    worksheet.insert_image('J17', 'Kwh_5min_heat_tape.png')
