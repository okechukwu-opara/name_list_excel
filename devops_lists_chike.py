#import libraries
import pandas as pd
import numpy as np
import openpyxl

#defining the excel file
excel_file = 'Lists.xlsx'

#reading the List 1 from the excel file
df_first_sheet = pd.read_excel(excel_file, sheet_name = 'List 1')

#reading the List 2 from the excel file
df_second_sheet = pd.read_excel(excel_file, sheet_name = 'List 2')

#sorting List 1
sorted_df1 = df_first_sheet.sort_values(by='Surname')

#sorting List 2
sorted_df2 = df_second_sheet.sort_values(by='Surname')

#getting the common names from the List 1 and List 2
common_names = np.intersect1d(sorted_df1, sorted_df2)

new_list = pd.DataFrame(common_names, columns=['Surname'])

#creating and appending a new list, List 3 to an already existing excel file, Lists.xlsx
with pd.ExcelWriter("Lists.xlsx", mode="a", engine="openpyxl", if_sheet_exists="replace") as writer:
    new_list.to_excel(writer, sheet_name="List 3", index= False)