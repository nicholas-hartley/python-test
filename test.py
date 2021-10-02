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

    names = file.sheet_names

    temp = file.parse(names[0])
    temp_str = temp.iloc[1,1]
    temp = temp.tail(n=-16)
    cols = temp.iloc[0, :]
    temp.columns = cols
    temp = temp.tail(n=-1)
    temp.set_index('Label', inplace=True)
    temp = temp.drop(columns=['Period', 'Year'])
    print(temp_str)

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
        print(temp2)
        temp_str = temp2_str

    return temp
    # return pd.concat([file.parse(name) for name in names])
testing = [create_df_from_excel(xl) for xl in excel_files]
df = testing[0].join(testing[1], lsuffix= '_1', rsuffix= '_2')

# save the data frame
writer = pd.ExcelWriter('..\\output.xlsx')
df.to_excel(writer, 'sheet1')
writer.save()