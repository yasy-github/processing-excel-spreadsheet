import os
import time
import pandas as pd


def find_max_page(files: list) -> tuple:
    """Find max page.
            :param files: list of filenames in str.
            :return: (number of page, file name).

    """

    files_page = []
    for f in files:
        df = pd.read_csv(PATH + f)

        count_page = 0
        for index, row in df.iterrows():
            
            # checking if the substring is present in a given string, e.g. row[20]
            if not pd.isnull(row[20]) and "Page" in row[20]:
                page_name = row[20]
                count_page = int(page_name.replace("of", " ").split()[-1])
                break

        files_page.append((f, count_page))

    max_page = 0
    name = ""
    for filename, page in files_page:
        if page > max_page: 
            max_page = page
            name = filename
        
    return max_page, name

def get_product_indexes(from_index: list, to_index: list) -> list:
    """Get row indices of products from specified boundary.
            :param from_index: list of integer of index.
            :param to_index: list of integer of index.
            :return: list of indices that lie between the boundary.

    """

    product_indexes = []
    for f, t in zip(from_index, to_index):
        product_indexes += [i for i in range(f + 1, t)]

    return product_indexes

def export_to_xlsx(data: list, columns: list, filename: str):
    """Export data as Excel spreadsheet.
            :param data: list of entries.
            :param columns: list of column names.
            :param filename: name in str.
            :return: True if success.

    """
    try:
        df = pd.DataFrame(data)
        # print("New Dataframe", df)
        # print("Shape: ", df.shape)
        df.to_excel(filename, index=False)

        return True
    except Exception as e:
        print("Error: ", e)
        return False

def main():
    # get list of filename
    csv_files = os.listdir(PATH)
    file_in_total = len(csv_files)

    data = []
    for seq, filename in enumerate(csv_files, start=1):
        print(f"\n============= {filename} =============")

        df = pd.read_csv(PATH + filename)
        df.columns = [i for i in range(0, df.shape[1])]     # set column based on shape
        # print(df.head(15))

        # get header information
        col_1 = df.loc[6][35]   # ref
        col_2 = df.loc[7][37]   # delivery date
        col_3 = df.loc[9][1]    # date
        col_4 = df.loc[9][5]    # IR No
        col_5 = df.loc[9][14]    # Project No
        col_6 = df.loc[9][25]    # Project Name
        col_7 = df.loc[10][14]    # Purpose
        col_8 = df.loc[10][7]    # Code 1
        col_9 = df.loc[10][9]    # Code 2

        # col_4 = df.loc[13][4]   # memo
        # print(col_1, col_2, col_3, col_4, col_5, col_6, col_7, col_8, col_9)
        # print(f"PO{po_reference} \tPM{project_manager_name}")

        # find boundary
        # if isinstance(row[1], str) and "CODE" in row[1]:
        from_index = df.index[df[1] == "កូដ\nCODE"].tolist()       # get index of `No` which is one of column header in table, for the starting point
        # if isinstance(row[1], str) and "Total" in row[1]:
        to_index = df.index[df[1] == "Total :"].tolist()        # get index of `Page#of#` to consider it as the ending point
        # print("\nBoundary", from_index, "->", to_index)

        # get product indexes that lie between boundary
        product_indexes = get_product_indexes(from_index, to_index)

        product_lines = []
        for index, row in df.iterrows():
            # append only specified row and its sequence number is not `NaN`
            if index in product_indexes and not pd.isna(row[1]):

            # if isinstance(row[1], str) and "Total" in row[1]:
            #     break

            # if isinstance(row[1], str) and "1" in row[1]:
                
                # get vals in columns
                line = [row[col] for col in LABEL_NUMBER]

                # add 3 more columns for `PO Ref`, `PM Name`, `Source File`
                # if row[1] != "No":

                # line.extend([po_reference[2:], project_manager_name[2:], filename.split(".")[0]])

                line.extend([col_1, col_2, col_3, col_4, col_5, col_6, col_7, col_8, col_9, filename.split(".")[0]])

                # append
                product_lines.append(line)

        # slicing to remove footer such as `Remark`, `Amount`, `Prepared by`, `Name`, `Date`
        # data.extend(product_lines[:-5])
        data.extend(product_lines)

        # break

        # clear console before printing
        # os.system('cls' if os.name == 'nt' else "printf '\033c'")
        print(f"Done processing: {seq} of {file_in_total} files")

    # setting output file, and export
    output_filename = 'result.xlsx'
    res = export_to_xlsx(data, LABEL_NAME, output_filename)
    print(res)

    time.sleep(1)
    
    
if __name__ == "__main__":
    start_time = time.time()

    # Initialize
    # PATH = os.getcwd() + "/dataset/"
    # LABEL_NUMBER = [1, 3, 8, 14, 17, 18, 22]
    # LABEL_NAME = ["No", "Item Code", "Description", "Qty", "UoM", "Unit Price", "Amount", "PO Ref", "PM Name", "Source File"]

    PATH = os.getcwd() + "/convert/result/"
    LABEL_NUMBER = [1, 3, 5, 11, 15, 17, 20, 22, 25, 26, 29, 34, 37]
    LABEL_NAME = ["CODE", "ML ID", "ITEM/Description", "Model", "Specification", "Qty", "Unit", "Cost", "Amount", "Master List", "Warehouse", "PO"]
    
    # run main function
    main()

    end_time = time.time()
    execution_time = end_time - start_time
    print(f"\nExecution time: {execution_time:.2f} seconds")

