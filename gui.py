from PySimpleGUI.PySimpleGUI import Window
import PySimpleGUI as sg
import main
# sg.theme("TealMono")
app_name = "Excel Comparer"

layout = [
    [
        sg.Column(
            [
                [
                    sg.Text(
                        app_name,
                        font=("Helvetica, 24"),
                        justification="c",
                    )
                ],
            ],
            element_justification="c",
            justification="c",
        )
    ],
    [
        sg.Text("Select Workbook 1"),
        sg.Input(),
        sg.FileBrowse("Select", key="--WB1--", file_types=(("Excel Workbook", "*.xlsx"),)),
    ],
    [
        sg.T("Sheet"),
        sg.InputText("Sheet1", key="--ws1--", size=(15, 1)),
        sg.T("Columns (in order)"),
        sg.InputText("A,B", key="--cols1--", size=(20, 1)),
    ],
    [sg.T("")],
    [sg.HorizontalSeparator()],
    [sg.T("")],
    [
        sg.Text("Select Workbook 2"),
        sg.Input(),
        sg.FileBrowse("Select", key="--WB2--", file_types=(("Excel Workbook", "*.xlsx"),)),
    ],
    [
        sg.T("Sheet"),
        sg.InputText("Sheet1", key="--ws2--", size=(15, 1)),
        sg.T("Columns (in order)"),
        sg.InputText("A,B", key="--cols2--", size=(20, 1)),
    ],
    [sg.T("")],
    [sg.HorizontalSeparator()],
    [sg.T("")],
    [
        sg.Button(" Start ", key="--START--"),
        sg.CBox(" Ignore case ", key="--case--")
    ],
]
font = "Arial, 12"
window = sg.Window(
    app_name, layout, resizable=True, font=font, margins=(10, 10), finalize=True
)
window.finalize()
while True:
    event, values = window.read()
    if (
        event == sg.WIN_CLOSED or event == "Cancel"
    ):  # if user closes window or clicks cancel
        break
    elif event == "--START--":
        wb1 = values.get("--WB1--", None)
        wb2 = values.get("--WB2--", None)
        if not wb1 or not wb2:
            print("Please select a valid file")
            continue
        ws1 = values.get("--ws1--", None)
        ws2 = values.get("--ws2--", None)
        cols1 = values.get("--cols1--", None)
        cols2 = values.get("--cols2--", None)
        case = values.get("--case--", None)
        try:
            main.main(wb1, wb2, ws1, ws2, cols1, cols2, case)
        except Exception as e:
            print("Something went wrong:", str(e))
window.close()
