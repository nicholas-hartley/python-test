# Testing Comments

import os
import pandas as pd

path = os.getcwd()
files = os.listdir(path)

excel_files = [file for file in files if '.xlsx' in file]
# excel_files

def create_df_from_excel(file_name):
    print("\n" + file_name + "\n")
    file = pd.ExcelFile(file_name)

    # This gets a list of the sheet names from 'file' and stores it in 'names'
    names = file.sheet_names

    # This is me making a dataframe of the first sheet
    temp = file.parse(names[0])
    # This is the string that I will be throwing on the observations column
    temp_str = temp.iloc[1,1]
    # This is the more confusing way to drop the first 16 rows (this is a linux like thing)
    temp = temp.tail(n=-16)
    # This is me getting the first row
    cols = temp.iloc[0, :]
    # Here I am assigning that first row to be my new column labels
    temp.columns = cols
    # This is me dropping the row I have now made the column labels
    temp = temp.tail(n=-1)
    # This is me changing the row names from numbers to one of the columns
    temp.set_index('Label', inplace=True)
    # This is me dropping the duplicate column data I don't need
    temp = temp.drop(columns=['Period', 'Year'])
    # print(temp_str)

    # This for loop is basically doing the same thing but iteratively through all the sheets
    for name in names[1:]:
        temp2 = file.parse(name)
        temp2_str = temp2.iloc[1,1]
        temp2 = temp2.tail(n=-16)
        cols = temp2.iloc[0, :]
        temp2.columns = cols
        temp2 = temp2.tail(n=-1)
        temp2.set_index('Label', inplace=True)
        temp2 = temp2.drop(columns=['Period', 'Year'])
        temp = temp.join(temp2, lsuffix= temp_str, rsuffix= temp2_str)
        # print(temp)
        # print(temp2)
        temp_str = temp2_str # This is for some loop stuff cause I didn't design it well 

    return temp
    # return pd.concat([file.parse(name) for name in names])
testing = [create_df_from_excel(xl) for xl in excel_files]
df = testing[0].join(testing[1], lsuffix= '_1', rsuffix= '_2')

# save the data frame
writer = pd.ExcelWriter('..\\output.xlsx')
df.to_excel(writer, 'sheet1')
writer.save()