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
        df = pd.DataFrame(data, columns=columns)
        df.to_excel(filename, index=False)

        return True
    except:
        return False

def main():
    # get list of filename
    csv_files = os.listdir(PATH)
    file_in_total = len(csv_files)

    data = []
    for seq, filename in enumerate(csv_files, start=1):
        # print(f"============= {filename} =============")

        df = pd.read_csv(PATH + filename)
        df.columns = [i for i in range(0, 23)]
        # print(df.head(23))

        # get header information
        po_reference = df.loc[2][16]
        project_manager_name = df.loc[9][16]
        # print(f"PO{po_reference} \tPM{project_manager_name}")

        # find boundary
        from_index = df.index[df[1] == "No"].tolist()       # get index of `No` which is one of column header in table, for the starting point
        to_index = df.index[df[20].notna()].tolist()        # get index of `Page#of#` to consider it as the ending point
        # print("\nBoundary", from_index, "->", to_index)

        # get product indexes that lie between boundary
        product_indexes = get_product_indexes(from_index, to_index)

        product_lines = []
        for index, row in df.iterrows():
            # append only specified row and its sequence number is not `NaN`
            if index in product_indexes and not pd.isna(row[1]):

                # get vals in columns
                line = [row[col] for col in LABEL_NUMBER]

                # add 3 more columns for `PO Ref`, `PM Name`, `Source File`
                if row[1] != "No":
                    line.extend([po_reference[2:], project_manager_name[2:], filename.split(".")[0]])

                # append
                product_lines.append(line)

        # slicing to remove footer such as `Remark`, `Amount`, `Prepared by`, `Name`, `Date`
        data.extend(product_lines[:-5])

        # clear console before printing
        # os.system('cls' if os.name == 'nt' else "printf '\033c'")
        print(f"Done processing: {seq} of {file_in_total} files")

    # setting output file, and export
    output_filename = 'result.xlsx'
    export_to_xlsx(data, LABEL_NAME, output_filename)

    time.sleep(1)
    
    
if __name__ == "__main__":
    start_time = time.time()

    # Initialize
    PATH = os.getcwd() + "/convert/result/"
    LABEL_NUMBER = [1, 3, 8, 14, 17, 18, 22]
    LABEL_NAME = ["No", "Item Code", "Description", "Qty", "UoM", "Unit Price", "Amount", "PO Ref", "PM Name", "Source File"]
    
    # run main function
    main()

    end_time = time.time()
    execution_time = end_time - start_time
    print(f"\nExecution time: {execution_time:.2f} seconds")

