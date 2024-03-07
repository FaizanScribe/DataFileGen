import xlwings as xw
import pandas as pd

SourceFileFromDump = 'C:\\Users\\USER\\Desktop\\PythonTestings\\report.xlsx'
DestinationFileForDataGen = 'C:\\Users\\USER\\Desktop\\PythonTestings\\DATA.xlsx'
DealersList = 'C:\\Users\\USER\\Desktop\\PythonTestings\\DealerList.xlsx'
app = xw.App(visible= False)
sourceBook = xw.Book(SourceFileFromDump)
destinationBook = xw.Book(DestinationFileForDataGen)
aDealersBook = xw.Book(DealersList)
sourceSheet = sourceBook.sheets['Worksheet']
destinationSheet = destinationBook.sheets['DATA']
DListSheet = aDealersBook.sheets['DealerList']
AcceptedDealers = []
Dealersdf =  pd.read_excel(DealersList, sheet_name='DealerList')
AcceptedDealers = Dealersdf['CreditDealersList'].tolist()

#print(AcceptedDealers)         #works

df2 = pd.read_excel(SourceFileFromDump, sheet_name='Worksheet')
condition = df2['customer_id'].isin(AcceptedDealers)
CreditDealersdf = df2[condition]

# print(CreditDealersdf)                 #Works
destinationSheet.range('A1').options(index=False).value = CreditDealersdf

sourceBook.save()
destinationBook.save()
aDealersBook.save()
sourceBook.close()
destinationBook.close()
aDealersBook.close()
app.quit()

# Works
