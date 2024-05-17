import pandas as pd
import os
import argparse


"""
Run script: 

    python3 remove_column.py
        --filename <filename.xlsx> 
        --sheet <sheet_name> 
        --from_col <from_column_name> 
        --to_col <to_column_name> 
        --focus <column_name_to_checking>
"""

def remove_columns(df, from_col, to_col, column_name):
    """ Remove the interval Columns. """

    # get Interval columns
    start = df.columns.get_loc(from_col)
    end = df.columns.get_loc(to_col)
    columns_to_remove = [df.columns[i] for i in range(start, end+1)]

    previous_column = current_column = ''
    for index, row in df.iterrows():
        current_column = row[column_name]     # get cell value of row by column name

        if current_column != previous_column:
            previous_column = current_column
        else:
            df.loc[index, columns_to_remove] = [' ' for _ in range(len(columns_to_remove))]

    df.to_excel('Result.xlsx', index=False)

    return True


if __name__ == '__main__':
    # Set up Command Line arguments
    parser = argparse.ArgumentParser()
    parser.add_argument('--filename', dest='filename', type=str, help='Input filename with extension')
    parser.add_argument('--sheet', dest='sheet', type=str, help='Sheet name of Excel file')
    parser.add_argument('--from_col', dest='from_col', type=str, help='From column')
    parser.add_argument('--to_col', dest='to_col', type=str, help='To column')
    parser.add_argument('--focus', dest='focus', type=str, help='Columnn to focus on')
    args = parser.parse_args()

    # filename with extension, e.g. N2401 NPPIA Item Request (1).xlsx
    filename = os.getcwd() + "/" + args.filename
    
    # create Dataframe
    df = pd.read_excel(filename, sheet_name=args.sheet)

    # called function
    remove_columns(df, args.from_col, args.to_col, args.focus)