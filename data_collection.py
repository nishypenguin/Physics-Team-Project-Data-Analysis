import pandas as pd
import matplotlib.pyplot as plt

# initialise date frame
#try:
 #   file_name = 'Data Collection sheets.xlsx'
  #  print("File found")
#except:
 #   print(f"{file_name} not found")   


#try:
 #   df = pd.read_excel(file_name, sheet_name="Control")
  #  df = df[['Number of squares', 'Number of squares']].copy()
   # print("Dataframe initialised")
#except:
 #   print("Dataframe failed to initialise")  

#print(df.head())
#print(df.info())
#print(df.dtypes)





def plot_data(file_name, name_of_sheet, x_data):

    y_data ='Number of microfibres'
    df_original = pd.read_excel(file_name, sheet_name=name_of_sheet)
    df = df_original[[f"{x_data}",f"{y_data}"]].copy()
    print(df.head())
    print(df.info())
    print(df.dtypes)
    plt.plot(df[f"{x_data}"],df[f"{y_data}"])
    plt.xlabel(f"{x_data}")
    plt.ylabel(f"{y_data}")
    plt.title(f"{y_data} vs {x_data}")
    plt.show()



file_name = 'Data Collection sheets.xlsx'
name_of_sheet = input("What is the name of the excel sheet you are using? ")
x_data = input("What variable are you varying, be careful to use the exact name used in the column of the spreadsheet? ")
#y_data = 'Number of microfibres'

plot_data(file_name, name_of_sheet, x_data)



