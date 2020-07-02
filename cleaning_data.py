import pandas as pd
import numpy as np
from openpyxl.workbook import Workbook


df = pd.read_excel("C:/Users/Mehedi Hassan Galib/Desktop/Python/ggg.xlsx")  #open the excel file


df.drop(columns = "Ok",inplace = True)  #remove the column named "Ok"


df = df.set_index("Year")
print(df.loc[2000])        #loc will return all values of given data
print(df.loc[2000:,"Pop"]) #here Year from 2000 to end will be showed with Pop values


print(df.iloc[0])          #iloc will return all values of given row


df = df.replace(np.nan, "N/A", regex = True)   #replace all the Nan Value with "N/A"
print(df)

to_excel= df.to_excel("modified.xlsx")         #Create a new excel file
