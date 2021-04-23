import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import datetime as dt
import xlrd

# A csv is a comma sepperated file, where columns of data are separated via commas.
## Dataframe manipulation is more or less the same once the file has been read into a dataframe object.
csv_file = 'D:\Data analyzation folder\Data_Analyzation\Excell_data_files\SF_Solar_Production.csv'

#Dataframe object
## Because a csv file is just a text file with a specific format, it does not have column names in the raw data
#### the names in the list shown are what I have named the columns from left to right for easier manipulation.
data = pd.read_csv(csv_file, names=["Meter Name", "Date", "Consumption Totalization", "Totalization Difference", "Max of difference"])

#Take the dates in the date column and format them into a datetime object for a cleaner date format and easier manipulation.
data['Date'] = pd.to_datetime(data['Date'])

#Find the difference between the previous interval and the current interval totalization
##See pandas diff() function documentation for info on the function's applications
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