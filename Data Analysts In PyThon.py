import pandas as pd
import matplotlib.pyplot as plt
import tkinter as tk
from tkinter import filedialog
def openfile(event=None):
    filename = filedialog.askopenfilename()
    print(filename)
    df = pd.read_csv(filename)
    df = df.dropna()
     # Request 1 
    df=df[df["Quantity Ordered"]!="Quantity Ordered"]
    df["Quantity Ordered"] = df["Quantity Ordered"].astype(float)
    df["Price Each"] = df["Price Each"].astype(float)
    df["Tong (USD)"] = df["Quantity Ordered"]*df["Price Each"]
    df1 = df.drop(["Order Date", "Order ID", "Purchase Address"], axis = 1)
        
    # Request 2
    df1 = df1.groupby(['Product']).sum()
    df1 = df1.merge(df[["Product","Price Each"]], how='inner', on='Product')
    df1 = df1.drop(["Price Each_x"],  axis = 1)
    df1.drop_duplicates(inplace = True)
    df2 = df1.sort_values(by=["Tong (USD)"], axis = 0, ascending = False)
    df2=df2.head(10)           
    fig, b2x=plt.subplots()
    b2x.bar(df2['Product'], df2["Tong (USD)"])
    plt.title("Request 2")
    plt.show()
    #Request 3
    df3=df.groupby(['Product'], as_index=False).sum()
    df3 = df1.sort_values(by=["Quantity Ordered"],
                          axis = 0, ascending = False)
    df3=df3.head(10)
    df3['Quantity Ordered'].plot.bar()
    plt.title("Request 3")
    plt.show()
    #Request 4
    def bai4(s):
        s = s[s.find(",")+2:s.rfind(",")]
        return s
    df4=df.drop(["Order Date","Product", "Order ID"], axis = 1)
    df4['city']=df['Purchase Address'].apply(bai4)
    df4=df4.groupby(['city'], as_index=False).sum()
    df4=df4.sort_values(by="Tong (USD)", ascending=False)
    df4=df4.head(10)
    fig, b4x=plt.subplots()
    b4x.bar(df4['city'], df4["Tong (USD)"])
    plt.title("Request 4")
    plt.show()
    #Request 5
    def bai5(a):
        b=a[a.find(' ')+1:a.find(':')]
        return b
    df5=df.drop(["Purchase Address","Product", "Order ID"], axis = 1)
    df5['Hour']=df['Order Date'].apply(bai5)
    df5=df5.groupby(['Hour'], as_index=False).sum()
    df5=df5.sort_values(by="Tong (USD)", ascending=False)
    df5=df5.head(10)
    fig, b5x=plt.subplots()
    b5x.bar(df5['Hour'], df5["Tong (USD)"])
    plt.title("Request 5")
    plt.show()
    #Request 6
    df6 = df.drop(["Purchase Address","Order Date","Quantity Ordered","Product","Price Each"], axis = 1)
    df6 = df6.groupby(['Order ID'], as_index=False).sum()
    df6 = df6.sort_values(by=["Tong (USD)"], axis = 0, ascending = False)
    df6 = df6.head(10)
    fig, b6x=plt.subplots()
    b6x.bar(df6['Order ID'], df6["Tong (USD)"])
    plt.title("Request 6")
    plt.show()
    #Request 7
    def bai7(c):
        d=c[c.find('/')+1:c.rfind('/')]
        return d
    df7=df.drop(["Purchase Address","Product", "Order ID"], axis = 1)
    df7['Days']=df['Order Date'].apply(bai7)
    df7=df7.groupby(['Days'], as_index=False).sum()
    fig, b7x=plt.subplots()
    b7x.bar(df7['Days'], df7["Tong (USD)"])
    plt.title("Request 7")
    plt.show()       
    with pd.ExcelWriter('Lastfile.xlsx') as writer:
        df2.to_excel(writer, sheet_name = 'Sheet1')
        df3.to_excel(writer, sheet_name = 'Sheet2')
        df4.to_excel(writer, sheet_name = 'Sheet3')
        df5.to_excel(writer, sheet_name = 'Sheet4')
        df6.to_excel(writer, sheet_name = 'Sheet5')
        df7.to_excel(writer, sheet_name = 'Sheet6')
root = tk.Tk()
root.title("Welcome to 369 <3")
root.geometry('150x200')
button = tk.Button(root, text='CHOOSE FILE',bg="yellow", fg="brown", command=openfile)
button.pack()
root.mainloop()


