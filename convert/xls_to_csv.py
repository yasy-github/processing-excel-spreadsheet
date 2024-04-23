import os
import pandas as pd


def convert(input_path, output_path):
    # input_path = os.getcwd() + "/convert/PO-NBC HQ-N2014/"
    # output_path = os.getcwd() + "/convert/result/"
    
    files = os.listdir(input_path)
    length = len(files)
    for i, file in enumerate(files, start=1):
        filename = file.split(".")[0]

        # Read the XLS file
        df = pd.read_excel(input_path + '/' + file)

        # Save the DataFrame as a CSV file
        df.to_csv(f"{output_path}/{filename} ({i}).csv", index=False)

        os.system('cls' if os.name == 'nt' else "printf '\033c'")
        print(f"Done converting: {i} of {length} files")

    return True



if __name__ == "__main__":
    convert()

