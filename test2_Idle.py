import os
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import seaborn as sns
import pylab
import xlsxwriter

os.chdir('/Users/conordoyle/SCRIPTS/Python/samples')
cwd = os.getcwd()

xlsx = pd.ExcelFile('Timesheets_Test_PY.xlsx')
TSData = pd.read_excel(xlsx, 'TimeSheets Report-Conor')

#df_TSData = sns.load_dataset("TSData")

print(TSData.head(2))
print("Hours data Type is: ", TSData['Hours'].dtype.kind)
print("Outside Hours Data Type is: ",TSData['Outside Hours'].dtype.kind)

#Convert Hours to Float
TSData['Hours'] = TSData['Hours'].astype(float)
TSData['Outside Hours'] = TSData['Outside Hours'].astype(float)

#Sum Hours and Outside Hrs to get Total Hours
TSData['total_hrs'] = TSData['Hours'].astype(float) + TSData['Outside Hours'].astype(float)
print(TSData['total_hrs'].head(5))

sorted_by_project = TSData.sort_values(['Project Name'])

#Add a Country Column, based on Office Name Column - Simple list of hard coded choices
conditions = [
    (TSData['Office Name'] == 'Melbourne') | (TSData['Office Name'] == 'Sydney'),
    (TSData['Office Name'] == 'Tokyo')]
choices = ['Australia', 'Japan']
TSData['Country'] = np.select(conditions, choices, default='Other')

print(TSData[['Country','Office Name']])

#Add Days Column, based on office location. Calculated by dividing total_hrs by a scalar - X2 separate methods used.
TSData['Days1'] = np.where(TSData['Office Name']=='Melbourne', TSData['total_hrs'].div(7.5), TSData['total_hrs'].div(8)) # Method 1 - Only 1 condition & 2 choices

conditions = [
    (TSData['Office Name'] == 'Melbourne') | (TSData['Office Name'] == 'Sydney'),
    (TSData['Office Name'] == 'Tokyo')]
choices = [TSData['total_hrs'].div(7.5), TSData['total_hrs'].div(8)]
TSData['Days'] = np.select(conditions, choices, default=TSData['total_hrs'].div(8)) # Method 2. Multiple conditions & choices

print(TSData[['Country','Office Name','Days','Days1']])

#print(sorted_by_project.head(5))

#print(sorted_by_project[['Employee Name','Project Name']].head(3))

print("marker line1")

# Confirm Data converted to Float type
print("Converted Hours Data Type is: ", TSData['Hours'].dtype.kind)
print("Converted Outside Hours Data Type is: ",TSData['Outside Hours'].dtype.kind)

#Create a Pivot Table - Trying  multiple  options. Adding Plot to graph results
TSData.set_index('Employee Name', inplace=True)
sns.set()
table = pd.pivot_table(TSData, values='Days', index =['Employee Name'], aggfunc=np.sum).plot(kind= 'bar')
plt.ylabel('Days Per Person');
#pylab.show()
plt.show()
#print(table)

print("marker line2")
#print(sorted_by_project[['Employee Name','total_hrs']].head(5))

#Alternative to Pivot table, use Group By

TS_Consol = (TSData.groupby(['Employee Name']).sum())
#print(TS_Consol)
TS_Sorted = TS_Consol.sort_values(['total_hrs'],ascending=False)
print(TS_Sorted)


print("marker line3")
# Graphing Data from Group BY
# TSData.set_index('Employee Name', inplace=True)
# sns.set()
# TS_Plot = TSData.groupby(['Employee Name']).sum().plot(kind='bar')
# x = TS_Plot['Employee Name']
# y = TS_Plot['Days']
# plt.xlabel('Employee Name')
# plt.ylabel('Days')
# plt.title('Days by Employee')
# plt.bar(x, y, label="Employee", color="b")
# pyalab.show()

print("marker line4")
#fig, ax = plt.subplots(figsize=(15,7))
#data.groupby(['Employee Name']).sum()['Days'].plot(ax=ax)

#Create Bar Chart of results
#x = TS_Consol['Employee Name']
#y = TS_Consol['Days']
#plt.xlabel('Employee Name')
#plt.ylabel('Days')
#plt.title('Days by Employee')
#plt.bar(x, y, label="Employee", color="b")
#plt.show()

#for name, data in TS_Sorted:
 #   plt.plot(data['Days'], data['Employee Name'], label=name)
#plt.xlabel('Employee Name')
#plt.ylabel('Days')
#plt.show()

#TS_Sorted.plot.bar()
#Write results to Excel file
# Use simple statement. Only 1 sheet
TS_Sorted.to_excel('TS_Sorted.xlsx',sheet_name='TS_Sorted')
# Use ExcelWriter to allow writing multiple sheets.
#with pd.ExcelWriter('TS_Sorted_2.xlsx') as writer: #//Syntax 2 follows 
writer = pd.ExcelWriter('TS_Sorted_2.xlsx', engine='xlsxwriter')
TS_Sorted.to_excel(writer, sheet_name='TS_Sorted')
TSData.to_excel(writer, sheet_name='TSData_Orig')

workbook = writer.book
worksheet = writer.sheets['TS_Sorted']

# Create a chart object 
chart = workbook.add_chart({'type': 'column'})

# Configure series
chart.add_series({'values': '=TS_Sorted!$K$2:$K$30'})

worksheet.insert_chart('M5', chart)

# Save the Excel file 
writer.save()

#pivoted_TSData =TSData.pivot_table(TSSData, index="Employee Name",values =["total_hrs"],aggfunc=np.sum)
#print(pivoted_TSData)
#print(TSData.groupby('Employee Name').agg('Hours'))
#TSData_subset = TSData[['Employee Name', 'Hours']]
#print(g)
