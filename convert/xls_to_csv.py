import os
import pandas as pd


def convert():
    input_path = os.getcwd() + "/convert/IR/"
    output_path = os.getcwd() + "/convert/result/"
    
    files = os.listdir(input_path)
    sheet_length = 398
    for file in files:
        for i in range(1, sheet_length):
            filename = file.split(".")[0]

            # Read the XLS file
            df = pd.read_excel(input_path + file, sheet_name=i)

            # Save the DataFrame as a CSV file
            df.to_csv(f"{output_path}{filename} ({i}).csv", index=False)

            os.system('cls' if os.name == 'nt' else "printf '\033c'")
            print(f"Done processing: {i} of {sheet_length} files")



if __name__ == "__main__":
    convert()

