import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import datetime as dt
import xlrd

csv_file = 'D:\Data analyzation folder\Data_Analyzation\Excell_data_files\SF_Solar_Production.csv'

#Dataframe object
data = pd.read_csv(csv_file, names=["Meter Name", "Date", "Consumption Totalization", "Totalization Difference", "Max of difference"])

data['Date'] = pd.to_datetime(data['Date'])

#Find the difference between the previous interval and the current interval totalization
data['Totalization Difference'] = data['Consumption Totalization'].diff()

selectedData = data.loc[:, ["Meter Name", "Date", "Consumption Totalization", "Totalization Difference", "Max of difference"]]


##############################################Format plot data for difference###############################

plotData = data.loc[:, ["Date", "Totalization Difference"]]
plotData2 = data.loc[:, ["Date", "Consumption Totalization"]]

plotData.plot(x='Date', y='Totalization Difference', kind='line')
plt.savefig('Totalization Spikes')

plotData.plot(x='Date', y='Totalization Difference', kind='line')
plt.ylim(0,200)
plt.savefig('Totalization Averages')

plotData2.plot(x='Date', y='Consumption Totalization', kind='line')
plt.savefig('Consumption Spikes')

plotData2.plot(x='Date', y='Consumption Totalization', kind='line')
plt.ylim(0, 3500000)
plt.savefig('Consumption Averages')


#print(selectedData)


with pd.ExcelWriter("data_analysis.xlsx", engine="xlsxwriter") as writer:
    data.to_excel(writer, sheet_name="sheet 1")
    worksheet = writer.sheets['sheet 1']
    worksheet.insert_image('J17', 'Totalization Spikes.png')
    worksheet.insert_image('J40', 'Totalization Averages.png')
    worksheet.insert_image('J60', 'Consumption Spikes.png')
    worksheet.insert_image('J100', 'Consumption Averages.png')