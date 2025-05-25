# Importing packages needed to manipulate data
import numpy as np
import pandas as pd
import os
import glob

# use glob to get all the csv files
# in the folder
path = os.getcwd()
csv_files = glob.glob(os.path.join(path, "*.xlsx"))

# loop over the list of csv files
for f in csv_files:
    df = pd.read_excel(f)  # Makes dataframe using file name
    bottomData = pd.DataFrame(df[-24:])  # Makes dataframe from data at the bottom of the spreadsheet
    bottomData.index = np.arange(0, 24)  # Changes index to 0 through 24

    bottomData.columns = bottomData.iloc[0]  # changes column names to first row
    bottomData.drop(index=0, inplace=True)  # deletes first row

    # DATAFRAME EDITING
    #df.drop(df.index[-24:], inplace=True)  # Deletes last 24 rows
    df = df[['Item', 'ND.Z', 'CentreX [µm]', 'CentreY [µm]']]  # Keeps ND.Z CentreX and CentreY

    # VARIABLES
    df2 = []  # part of dataframe that contains median duplicate cell
    duplicateList = []  # temporarily stores duplicate layers
    winCondition = False  # Checks if x has been added to duplicate list
    breakCondition = False  # Breaks if no more rows left

    df7 = pd.DataFrame(columns=['Row', 'ND.Z', 'Item'])
    for x in df.index:  # Checks every row in df
        if x not in df.index:  # if the row doesn't exist move to next row
            continue
        winCondition = True  # x value will be added to duplicate list
        for b in duplicateList:
            df7.loc[len(df7.index)] = b
        duplicateList = []  # clears duplicateList for next run through y loop

        # Checks every row value less than 4 layers away    ( also edit Layer spaces)
        for y in df[(df['ND.Z'] <= df.loc[x]['ND.Z'] + 5) & (df['ND.Z'] > df.loc[x]['ND.Z'])].index:
            if y > df.index[-1]:  # If there are no rows left leave both loops
                breakCondition = True
                break

            # Checks if x and y rows are duplicates
            if (((pow(abs(df.loc[x]['CentreX [µm]'] - df.loc[y]['CentreX [µm]']), 2) + pow(
                    abs(df.loc[x]['CentreY [µm]'] - df.loc[y]['CentreY [µm]']), 2)) ** 0.5) <= 6.19):

                print(x, y, len(df.index))

                if winCondition:  # if x hasn't been added to duplicatelist yet it is added
                    duplicateList.append([x, df.loc[x]['ND.Z'],df.loc[x]['Item']])

                    winCondition = False
                duplicateList.append([y, df.loc[y]['ND.Z'],df.loc[y]['Item']])  # add y to duplicate list

                df.drop(index=y, inplace=True)  # y row is deleted from dataframe df

        if len(duplicateList) >= 1:  # checks if duplicateList length is greater than or equal to one
            # if length of duplicateList is even then add lower median value to dataframe df2
            if len(duplicateList) % 2 == 0:
                df2.append(duplicateList[int((len(duplicateList) / 2) - 1)])
                duplicateList.remove(duplicateList[int((len(duplicateList) / 2) - 1)])
                df.drop(index=x, inplace=True)
            else:  # if length of duplicateList is odd then add the median value to dataframe df2
                df2.append(duplicateList[int((len(duplicateList) / 2))]) # decimal cells dont exist
                duplicateList.remove(duplicateList[int((len(duplicateList) / 2))])
                df.drop(index=x, inplace=True)

        if breakCondition:  # exits for loop
            break

    print("Final", len(df.index), df.head(14))

    df3 = pd.DataFrame(df2, columns=['Row', 'ND.Z', 'Item'])  # Makes new dataframe with columns row and ND.Z with df2
    #df3.set_index(df3.loc[:, 'Row'], inplace=True)  # Sets df3 index to its original row value
    df3.drop(columns=['Row'], inplace=True)  # Deletes Row column because it is now the index
    df7.drop(columns=['Row'], inplace=True)  # Deletes Row column because it is now the index
    print(df3)
    df4 = pd.concat([df3, df])  # Adds together median values and non duplicates
    #df4 = df4[~df4.index.duplicated(keep='first')]  # Removes same indexes
    print(df4)

    print("Final", len(df.index), df.head(14))

    df6 = pd.DataFrame(columns=['ND.Z', 'Count'])
    for x in range(1, 102, 1):
        df6.loc[len(df6.index)] = [x, len(df4[(df4['ND.Z'] > (x - 1)) & (df4['ND.Z'] < x + 1)])]

    print(df6)

    df4 = pd.concat([df4,df6])

    # CONVERTING TO XLSX FILE
    print(f)
    f = f[0:-5] + " Cleansed.xlsx"
    print(f)
    # IMPORTANT: CHANGE FILE DIRECTORY BUT KEEP /df.xlsx . df is name of the file
    df4.to_excel(f, index=False)

    f = f[0:-14] + " Excised.xlsx"

    print(f)

    df7.to_excel(f, index=False)

