import PySimpleGUI as gui
from pathlib import Path

from tools.process import generate
from convert.xls_to_csv import convert


def is_valid_path(filepath):
    if filepath and Path(filepath).exists():
        return True
    gui.popup_error("Invalid path")
    return False

def progress_window(max_value=100):
    layout= [
        [gui.ProgressBar(max_value=max_value, orientation="h", size=(50,30), key="progress_bar", bar_color=("Blue", "Yellow"))]
    ]

    window = gui.Window("Loading", layout=layout, modal=True)

    while True:
        event, values = window.read()

        if event in (gui.WINDOW_CLOSED, "Exit"):
            break
        elif event == "Display Progress":
            progress_val = 25
            window['progress_bar'].update(progress_val)

    window.close()

def main_window():
    menu = [
        ["Help", ["About"]]
    ]

    layout = [
        [gui.MenubarCustom(menu, tearoff=False)],
        [gui.Text("Input Folder:", s=16, justification="r"), gui.Input(key="input_folder_path"), gui.FolderBrowse()],
        [gui.Text("Output Folder:", s=16, justification="r"), gui.Input(key="output_folder_path"), gui.FolderBrowse()],
        [gui.Text("Features:", s=16, justification="r"), gui.Radio('Convert', group_id=1, key='to_csv'), gui.Radio('Generate', group_id=1, key="is_generate")],
        [gui.Exit(s=16), gui.Button("Confirm", s=16)],
    ]

    window = gui.Window(title="Excel Manipulation Program", layout=layout)

    while True:
        event, values = window.read()
        print(event, values)

        if event in (gui.WINDOW_CLOSED, "Exit"):
            break

        if event == "About":
            window.disappear()
            gui.popup("Excel Manipulation Program", "Version: 1.0.0", "Author: WANG YAXI", "Email: whangyasy@gmail.com")
            window.reappear()

        if event == "Confirm":
            # TODO
            # progress_window()
            result = False

            if is_valid_path(values['input_folder_path']) and is_valid_path(values['output_folder_path']):
                if not values['to_csv'] and not values['is_generate']:
                    gui.popup_error("Please select one of the feature.")

                if values['to_csv']:
                    result = convert(
                        values['input_folder_path'],
                        values['output_folder_path']
                    )
                elif values['is_generate']:
                    label_name=["No", "Item Code", "Description", "Qty", "UoM", "Unit Price", "Amount", "PO Ref", "PM Name", "Source File"]
                    label_number=[1, 3, 8, 14, 17, 18, 22]

                    result = generate(
                        values['input_folder_path'],
                        values['output_folder_path'],
                        label_name,
                        label_number
                    )

                if result:
                    gui.popup_no_titlebar("Done!")
                else:
                    gui.popup_error("Something went wrong!")

    window.close()


if __name__ == '__main__':
    main_window()