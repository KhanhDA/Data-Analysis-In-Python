import pandas as pd
import matplotlib.pyplot as plt
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog as fd
from tkinter.messagebox import showinfo
import os

def my_func(s):
    s = s[s.find(",")+2:s.rfind(",")]
    return s
def my_func2(s):
    s = s[3:s.rfind("/")]
    return s
def my_func3(s):
    s = s[s.find(" ")+1:s.find(":")]
    return s

df = pd.read_csv('Sales_August_2019.csv')

df = df.dropna()
df = df[df["Quantity Ordered"].str.isdigit()==True]

df["Quantity Ordered"] = df["Quantity Ordered"].astype(float)
df["Price Each"] = df["Price Each"].astype(float)
df["Tong (USD)"] = df["Quantity Ordered"]*df["Price Each"]

df1 = df.drop(["Order Date", "Order ID", "Purchase Address"], axis = 1)
    
df1 = df1.groupby(['Product']).sum()
df1 = df1.merge(df[["Product","Price Each"]], how='inner', on='Product')
df1 = df1.drop(["Price Each_x"],  axis = 1)
df1.drop_duplicates(inplace = True)
df2 = df1.sort_values(by=["Tong (USD)"], axis = 0, ascending = False)
df3 = df1.sort_values(by=["Quantity Ordered"],
                      axis = 0, ascending = False)
df4 = df.drop(["Order Date", "Order ID"], axis = 1)
df4["City"] = df4["Purchase Address"].apply(my_func)
df4 = df4.drop(["Purchase Address", "Product"], axis = 1)
df4 = df4.groupby(['City']).sum()
df4 = df4.sort_values(by=["Tong (USD)"], axis = 0, ascending = False)

df7 = df.drop(["Purchase Address","Order ID","Product"], axis = 1)
df7["Days"] = df7["Order Date"].apply(my_func2)
df7 = df7.drop(["Order Date"], axis = 1)
df7 = df7.groupby(['Days']).sum()
df7 = df7.sort_values(by=["Tong (USD)"], axis = 0, ascending = False)

df5 = df.drop(["Purchase Address","Order ID","Product"], axis = 1)
df5["Hours"] = df5["Order Date"].apply(my_func3)
df5 = df5.drop(["Order Date"], axis = 1)
df5 = df5.groupby(['Hours']).sum()
df5 = df5.sort_values(by=["Tong (USD)"], axis = 0, ascending = False)

df6 = df.drop(["Purchase Address","Order Date","Product","Purchase Address","Price Each"], axis = 1)
df6 = df6.groupby(['Order ID']).sum()
df6 = df6.sort_values(by=["Tong (USD)"], axis = 0, ascending = False)

df2 = df2.head(10)
df3 = df3.head(10)
df7 = df7.head(10)
df6 = df6.head(10)
df5 = df5.head(10)
#code đồ thị cho df7
#df7['Tong (USD)'].plot.bar()
#df2['Tong (USD)'].plot.bar()
df3['Quantity Ordered'].plot.bar()
plt.show()

with pd.ExcelWriter('expored_file4.xlsx') as writer:
    df2.to_excel(writer, sheet_name = 'Sheet1')
    df3.to_excel(writer, sheet_name = 'Sheet2')
    df4.to_excel(writer, sheet_name = 'Sheet3')
    df7.to_excel(writer, sheet_name = 'Sheet4')
    df6.to_excel(writer, sheet_name = 'Sheet5')
    df5.to_excel(writer, sheet_name = 'Sheet6')