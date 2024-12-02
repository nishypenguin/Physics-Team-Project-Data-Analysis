import pandas as pd
import numpy as np 
import matplotlib.pyplot as plt
from scipy.stats import linregress
import seaborn as sns

data_sheets = ['DataAnalysis-PH', 'DataAnalysis-VolumeOfWater', 'DataAnalysis-BallBearing', 'DataAnalysis-Time', 'DataAnalysis-SpinRate', 'DataAnalysis-Temperature' ]

def plot_number_of_microfibres_against_variables(file, data_sheets):
    # Load the data from the specified sheet
    for sheet in data_sheets:
        df = pd.read_excel(file, sheet_name= sheet)
        independent_variable = df.columns[1]
        dependent_variable = 'Microfibre Number'
        error_column = 'Microfibre Number Error'
        print(f"Error Column = {df[error_column].head(5)}")

        # Drop rows with missing values for the variables of interest
        df = df.dropna(subset=[independent_variable, dependent_variable])
        print(f"First 5 values of {independent_variable}:")
        print(df[independent_variable].head())
        print(f"First 5 values of {dependent_variable}:")
        print(df[dependent_variable].head())

        # Calculate Q1, Q3, and IQR for the dependent variable
        Q1 = df[dependent_variable].quantile(0.25)
        Q3 = df[dependent_variable].quantile(0.75)
        IQR = Q3 - Q1
        print(f"Q1 (25th percentile): {Q1}")
        print(f"Q3 (75th percentile): {Q3}")
        print(f"IQR: {IQR}")

        # Identify and filter out outliers
        outliers = df[
            (df[dependent_variable] < (Q1 - 1.5 * IQR)) |
            (df[dependent_variable] > (Q3 + 1.5 * IQR))
        ]
        print("Outliers:")
        print(outliers)

        # Perform linear regression using scipy's linregress
        regression = linregress(df[independent_variable], df[dependent_variable])

        # Extract regression metrics
        slope = regression.slope
        intercept = regression.intercept
        r_value = regression.rvalue
        r_squared = r_value**2
        p_value = regression.pvalue
        std_err = regression.stderr

        # Print the statistical analysis
        print("\nStatistical Analysis:")
        print(f"Slope: {slope}")
        print(f"Intercept: {intercept}")
        print(f"R-squared: {r_squared}")
        print(f"P-value: {p_value}")
        print(f"Standard Error: {std_err}")

        # Generate line of best fit
        line_of_best_fit = slope * df[independent_variable] + intercept
        x_values = np.linspace(df[independent_variable].min(), df[independent_variable].max(), 100)
        y_values = slope * x_values + intercept

        # Plot the data with error bars and the line of best fit
        plt.errorbar(df[independent_variable], df[dependent_variable], yerr=df[error_column], fmt='o', capsize=5, label='Data Points')
        plt.plot(x_values, y_values, 'r-', label='Line of Best Fit')
        plt.xlabel(independent_variable)
        plt.ylabel(dependent_variable)
        plt.title(f"{dependent_variable} vs {independent_variable}")
        plt.legend()

        # Annotate the graph with R-squared value
        plt.annotate(f"$R^2$ = {r_squared:.4f}", xy=(0.05, 0.95), xycoords='axes fraction', fontsize=10, ha='left', va='top')
        
        plt.savefig(f"Number of microfibres {sheet}_plot.png")
        plt.clf()
    


def plot_average_length_of_microfibres_against_variables(file, data_sheets):
    # Load the data from the specified sheet
    for sheet in data_sheets:
        df = pd.read_excel(file, sheet_name=sheet)
        independent_variable = df.columns[1]
        dependent_variable = 'Microfibre Length '
        error_column = 'Microfibre Length Error'
        print(f"Microfibre Length = {df[dependent_variable].head(5)}")
        print(f"Error Column = {df[error_column].head(5)}")

        # Drop rows with missing values for the variables of interest
        df = df.dropna(subset=[independent_variable, dependent_variable])
        print(f"First 5 values of {independent_variable}:")
        print(df[independent_variable].head())
        print(f"First 5 values of {dependent_variable}:")
        print(df[dependent_variable].head())

        # Calculate Q1, Q3, and IQR for the dependent variable
        Q1 = df[dependent_variable].quantile(0.25)
        Q3 = df[dependent_variable].quantile(0.75)
        IQR = Q3 - Q1
        print(f"Q1 (25th percentile): {Q1}")
        print(f"Q3 (75th percentile): {Q3}")
        print(f"IQR: {IQR}")

        # Identify and filter out outliers
        outliers = df[
            (df[dependent_variable] < (Q1 - 1.5 * IQR)) |
            (df[dependent_variable] > (Q3 + 1.5 * IQR))
        ]
        print("Outliers:")
        print(outliers)

        # Perform linear regression using scipy's linregress
        regression = linregress(df[independent_variable], df[dependent_variable])

        # Extract regression metrics
        slope = regression.slope
        intercept = regression.intercept
        r_value = regression.rvalue
        r_squared = r_value**2
        p_value = regression.pvalue
        std_err = regression.stderr

        # Print the statistical analysis
        print("\nStatistical Analysis:")
        print(f"Slope: {slope}")
        print(f"Intercept: {intercept}")
        print(f"R-squared: {r_squared}")
        print(f"P-value: {p_value}")
        print(f"Standard Error: {std_err}")

        # Generate line of best fit
        line_of_best_fit = slope * df[independent_variable] + intercept
        x_values = np.linspace(df[independent_variable].min(), df[independent_variable].max(), 100)
        y_values = slope * x_values + intercept

        # Plot the data with error bars and the line of best fit
        plt.errorbar(df[independent_variable], df[dependent_variable], yerr=df[error_column], fmt='o', capsize=5, label='Data Points')
        plt.plot(x_values, y_values, 'r-', label='Line of Best Fit')
        plt.xlabel(independent_variable)
        plt.ylabel(dependent_variable)
        plt.title(f"{dependent_variable} vs {independent_variable}")
        plt.legend()

        # Annotate the graph with R-squared value
        plt.annotate(f"$R^2$ = {r_squared:.4f}", xy=(0.05, 0.95), xycoords='axes fraction', fontsize=10, ha='left', va='top')
        
        plt.savefig(f"Average length of microfibres {sheet}_plot.png")
        plt.clf()

def plot_volume_of_microfibres_against_variables(file, data_sheets):
    # Load the data from the specified sheet
    for sheet in data_sheets:
        df = pd.read_excel(file, sheet_name=sheet)
        independent_variable = df.columns[1]
        dependent_variable = 'Microfibre Volume'
        error_column = 'Microfibre Length Error'
        print(f"Microfibre Length = {df[dependent_variable].head(5)}")
        print(f"Error Column = {df[error_column].head(5)}")

        # Drop rows with missing values for the variables of interest
        df = df.dropna(subset=[independent_variable, dependent_variable])
        print(f"First 5 values of {independent_variable}:")
        print(df[independent_variable].head())
        print(f"First 5 values of {dependent_variable}:")
        print(df[dependent_variable].head())

        # Calculate Q1, Q3, and IQR for the dependent variable
        Q1 = df[dependent_variable].quantile(0.25)
        Q3 = df[dependent_variable].quantile(0.75)
        IQR = Q3 - Q1
        print(f"Q1 (25th percentile): {Q1}")
        print(f"Q3 (75th percentile): {Q3}")
        print(f"IQR: {IQR}")

        # Identify and filter out outliers
        outliers = df[
            (df[dependent_variable] < (Q1 - 1.5 * IQR)) |
            (df[dependent_variable] > (Q3 + 1.5 * IQR))
        ]
        print("Outliers:")
        print(outliers)

        # Perform linear regression using scipy's linregress
        regression = linregress(df[independent_variable], df[dependent_variable])

        # Extract regression metrics
        slope = regression.slope
        intercept = regression.intercept
        r_value = regression.rvalue
        r_squared = r_value**2
        p_value = regression.pvalue
        std_err = regression.stderr

        # Print the statistical analysis
        print("\nStatistical Analysis:")
        print(f"Slope: {slope}")
        print(f"Intercept: {intercept}")
        print(f"R-squared: {r_squared}")
        print(f"P-value: {p_value}")
        print(f"Standard Error: {std_err}")

        # Generate line of best fit
        line_of_best_fit = slope * df[independent_variable] + intercept
        x_values = np.linspace(df[independent_variable].min(), df[independent_variable].max(), 100)
        y_values = slope * x_values + intercept

        # Plot the data with error bars and the line of best fit
        plt.errorbar(df[independent_variable], df[dependent_variable], yerr=df[error_column], fmt='o', capsize=5, label='Data Points')
        plt.plot(x_values, y_values, 'r-', label='Line of Best Fit')
        plt.xlabel(independent_variable)
        plt.ylabel(dependent_variable)
        plt.title(f"{dependent_variable} vs {independent_variable}")
        plt.legend()

        # Annotate the graph with R-squared value
        plt.annotate(f"$R^2$ = {r_squared:.4f}", xy=(0.05, 0.95), xycoords='axes fraction', fontsize=10, ha='left', va='top')
        
        plt.savefig(f"Volume of microfibres {sheet}_plot.png")
        plt.clf()

def plot_correlation_of_variables(file,data_sheets):
    for sheet in data_sheets:
        df = pd.read_excel(file, sheet_name=sheet)
        df = df.drop(['Microfibre Number Error', 'Microfibre Length Error','Unnamed: 0'], axis = 1)
        print(df.head(5))
        correlation_matrix = df.corr()
        plt.figure(figsize=(10, 8))
        sns.heatmap(correlation_matrix, annot=True, cmap='coolwarm', fmt='.2f')
        plt.title('Correlation Matrix')
        plt.savefig(f"Correlation table {sheet}_plot.png")
        plt.clf()

file = "Data Collection sheets Redo 2.xlsx"
plot_number_of_microfibres_against_variables(file, data_sheets)
plot_average_length_of_microfibres_against_variables(file, data_sheets)
plot_volume_of_microfibres_against_variables(file, data_sheets)
plot_correlation_of_variables(file, data_sheets)

    
