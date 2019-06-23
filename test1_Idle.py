import os
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import seaborn as sns

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

#Sum Hours and  Outside Hrs to get Total Hours
TSData['total_hrs'] = TSData['Hours'].astype(float) + TSData['Outside Hours'].astype(float)
print(TSData['total_hrs'].head(5))

sorted_by_project = TSData.sort_values(['Project Name'])
plt.show()

print(sorted_by_project.head(5))

print(sorted_by_project[['Employee Name','Project Name']].head(3))

print("spare line")

# Confirm Data converted to Float type
print("Converted Hours Data Type is: ", TSData['Hours'].dtype.kind)
print("Converted Outside Hours Data Type is: ",TSData['Outside Hours'].dtype.kind)

#pd.pivot_table(TSData,index=['Employee Name','Projet Name'],values=['Hours'],aggfunc=np.sum)


#print(sorted_by_project[['Employee Name','total_hrs']].head(5))
TS_Consol = (TSData.groupby(['Employee Name']).sum())
print(TS_Consol)
TS_Sorted = TS_Consol.sort_values(['total_hrs'],ascending=False)
print(TS_Sorted)

#Write results to Excel file
# Use simple statement. Only 1 sheet
TS_Sorted.to_excel('TS_Sorted.xlsx',sheet_name='TS_Sorted')
# Use ExcelWriter to allow writing multiple sheets.
with pd.ExcelWriter('TS_Sorted_2.xlsx') as writer:
    TS_Sorted.to_excel(writer, sheet_name='TS_Sorted')
    TSData.to_excel(writer, sheet_name='TSData_Orig')
#pivoted_TSData =TSData.pivot_table(TSSData, index="Employee Name",values =["total_hrs"],aggfunc=np.sum)
#print(pivoted_TSData)
#print(TSData.groupby('Employee Name').agg('Hours'))
#TSData_subset = TSData[['Employee Name', 'Hours']]
#print(g)
