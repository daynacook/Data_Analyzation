import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import datetime as dt

excel_file = 'D:\Data analyzation folder\Excell_data_files\Raw_Geology_Data.xlsx'

#https://pandas.pydata.org/docs/reference/api/pandas.DataFrame.loc.html
###the .loc function is used extensively to select key columns of data and manipulate them, 
#####Keeping the docs handy id recomended

#Dataframe object
data = pd.read_excel(excel_file,  header=[20, 21]) #header=[20,21] dictares which row the reader translates as header rows

#concatenate the header rows into one row so as to not mix headers into data rows
data.columns = (pd.Series(data.columns.get_level_values(0)).apply(str)
                + pd.Series(data.columns.get_level_values(1)).apply(str))


# kW calculations
data['kW'] = data['VAtAvg'] / 1000
data['kWh/5min'] = data['kW'] * 5 * 60 / 3600

# create Time column which combines the start date and time into one column
data['Time'] = data['* Start* Date'].apply(str) + " " + data['StartTime'].apply(str)
data['Time'] = pd.to_datetime(data['Time'])
print(data['Time'])

# pricing calculations, found from meeting with John Heitzenger, Contact for new pricing data for future projects
data['kW Price'] = data['kW'] * 18.1143296
data['kWh Price'] = data['kWh/5min'] * 0.038559558

# kWh/5min monthly and cumulative sums
watSum = data.groupby(['* Start* Date']).sum()
watSumTotal = watSum.cumsum()
watSumPrint = watSum.loc[:, ['kW', 'kW Price', 'kWh/5min', 'kWh Price']]
watSumTotalPrint = watSumTotal.loc[:, ['kW', 'kW Price', 'kWh/5min', 'kWh Price']]

#format watData output
watData = data.loc[:, ['* Start* Date', 'StartTime', 'kWh/5min']]

#print data and sums
print(watSumPrint)
print(watSumTotalPrint)
print(watData)




##############################Plotting and formatting###########################




# Plot data over 5 minute intervals
watData.plot(xlabel = 'Time', ylabel = 'kWh/5min', y='kWh/5min')
# Find x min and max values
x_min, x_max = plt.xlim()
# Arange the x-axis ticks to go from 0 to the found max value, showing ticks at intervals of every 250 minutes. Rotates tick labels 45 degrees.
plt.xticks(np.arange(0, x_max, 250), rotation=45)
# Save the figure as a png with the title kWh_5min
plt.savefig('kWh_5min')

# Plot sums over monthly intervals
watSumPrint.plot(xlabel = 'Time', ylabel = 'kWh/5min', y='kWh/5min')
plt.savefig('kWh_5min_Monthly_Intervals.png')

# Format date to print in Excel
data['* Start* Date'] = data['* Start* Date'].dt.strftime('%m/%d/%Y') #strftime is formatting the datetime object to a month/date/year format

# Output data into an excel sheet
allData = data.loc[:, ['* Start* Date', 'StartTime', 'VAtAvg', 'kW', 'kW Price', 'kWh/5min', 'kWh Price']]
#take the data found in the other wat sum data and catagorize them as sumData and cumulative sum data
sumData = watSumPrint
cumSumData = watSumTotalPrint
#Join the sumData and the cumSumData so that the data can be presented next to each other, using suffixes to differentiate the columns
sumData = sumData.join(cumSumData, lsuffix='_sum', rsuffix='_cumsum')

# write the data to an excel sheet, formatting the images so as to not overlap with the rest of the data 
with pd.ExcelWriter("data_analysis.xlsx", engine="xlsxwriter") as writer:
    allData.to_excel(writer, sheet_name="sheet 1")
    sumData.to_excel(writer, sheet_name="sheet 1", startcol=8) #startcol dictates where the start column for sumData starts to be written in the sheet
    worksheet = writer.sheets['sheet 1']
    worksheet.insert_image('J17', 'kWh_5min.png')
    worksheet.insert_image('J40', 'kWh_5min_Monthly_Intervals.png')
