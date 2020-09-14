import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import datetime as dt

excel_file = 'D:\Data analyzation folder\Excell_data_files\Raw_Geology_Data.xlsx'

#Dataframe object
data = pd.read_excel(excel_file,  header=[20, 21])

#cocatenate the header rows into one row
data.columns = (pd.Series(data.columns.get_level_values(0)).apply(str)
                + pd.Series(data.columns.get_level_values(1)).apply(str))


# kWh calculations
data['kWh'] = data['VAtAvg'] / 1000
data['kWh/5min'] = data['kWh'] * 5 * 60 / 3600

# kWh/5min monthly and cumulative sums
watSum = data.groupby(['* Start* Date']).sum()
watSumTotal = watSum.sum()
watSumPrint = watSum.loc[:, ['kWh/5min']]

watData = data.loc[:, ['* Start* Date', 'StartTime', 'kWh/5min']]

#print data and sums
print(watSumPrint)
print(watSumTotal.loc['kWh/5min'])
print(watData)

# Plot data over 5 minute intervals
watData.plot(xlabel = 'Time', ylabel = 'kWh/5min', y='kWh/5min')
x_min, x_max = plt.xlim()
plt.xticks(np.arange(0, x_max, 250), rotation=45)
plt.show()

# Plot sums over monthly intervals
watSumPrint.plot(xlabel = 'Time', ylabel = 'kWh/5min', y='kWh/5min')
plt.show()

#output data into an excel sheet
allData = data.loc[:, ['* Start* Date', 'StartTime', 'VAtAvg', 'kWh', 'kWh/5min']]
allData.to_excel("output.xlsx")
