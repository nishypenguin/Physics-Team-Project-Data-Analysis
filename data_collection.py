import pandas as pd
import matplotlib.pyplot as plt
import numpy as np

class DataAnalyser():

  def __init__(self, file = 'Data Collection sheets'):
     self.file = file
     self.dataframe = pd.read_excel(file)

  def find_appropriate_sheets(self):
    list_of_sheets_for_data_analysis = []
    for sheet_name in self.dataframe.sheet_names:
      if 'DataAnalysis' in sheet_name:
        list_of_sheets_for_data_analysis.append(sheet_name)

      return list_of_sheets_for_data_analysis
  
  def plot_number_of_microfibres_against_variables(self):
      

      list_of_sheets_for_data_analysis = self.find_appropriate_sheets()
      for sheet_name in list_of_sheets_for_data_analysis:
        df = pd.read_excel(self.file, sheet_name)
        independent_variable = df.iat[0,1]
        dependent_variable ='Microfibre Number'
        error_column = 'Microfibre Number Error'

        coefficients = np.polyfit(df[independent_variable], df[dependent_variable], 1)
        line_of_best_fit = np.poly1d(coefficients)
        x_values = np.linspace(df[independent_variable].min(), df[independent_variable].max(), 100)
        y_values = line_of_best_fit(x_values)

        plt.errorbar(df[independent_variable], df[dependent_variable], yerr=df[error_column], fmt='o', capsize=5, label='Data Points')
        plt.plot(x_values, y_values, 'r-', label='Line of Best Fit')
        plt.xlabel(independent_variable)
        plt.ylabel(dependent_variable)
        plt.title(f"{dependent_variable} vs {independent_variable}")
        plt.legend()
        plt.show()

        # Statz - Decide with team how we go forward

        #standard_deviation = np.std(df[dependent_variable])


  def plot_length_of_microfibres_against_variable(self):
     list_of_sheets_for_data_analysis = self.find_appropriate_sheets()
     for sheet_name in list_of_sheets_for_data_analysis:
        df = pd.read_excel(self.file, sheet_name)
        independent_variable = df.iat[0,1]
        dependent_variable ='Microfibre Length'
        error_column = 'Microfibre Length Error'
        # mainpulate dataframe
        # plot histogram
        # 20 bins of 15 units length


analyse_data = DataAnalyser()
plot_numberplot_number_of_microfibres_against_variables_graphs = analyse_data.plot_number_of_microfibres_against_variables()
      



