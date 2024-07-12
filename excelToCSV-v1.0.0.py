import tkinter as tk
import tkinter.messagebox
import os
import pandas as pd

# Create an instance of Tkinter frame
font_big = "Calibri 14"
font_small = "Calibri 11"
root = tk.Tk()
root.title('excelToCSV (v1.0.0)  |  © NGUYEN HOANG PHU, 2023')
# Place window center
width = 750  # Width
height = 300  # Height
screen_width = root.winfo_screenwidth()  # Width of the screen
screen_height = root.winfo_screenheight()  # Height of the screen
# Calculate Starting X and Y coordinates for Window
x = (screen_width / 2) - (width / 2)
y = (screen_height / 2) - (height / 2)
root.geometry('%dx%d+%d+%d' % (width, height, x, y))

tk.Label(root, text="Input folder location: ", font=font_small).grid(row=0, column=0)
tk.Label(root, text="Sheet index (1,2,3...): ", font=font_small).grid(row=1, column=0)
tk.Label(root, text="Header starts at (1,2,3...): ", font=font_small).grid(row=2, column=0)
tk.Label(root, text="CSV/TXT: ", anchor="w", font=font_small).grid(row=3, column=0)
tk.Label(root, text="Separator (, ; | ~): ", anchor="w", font=font_small).grid(row=4, column=0)
tk.Label(root, text="Output folder location: ", font=font_small).grid(row=5, column=0)

dir_path = tk.Entry(root, width=80)
s_index = tk.Entry(root, width=80)
header = tk.Entry(root, width=80)
csv_name = tk.Entry(root, width=80)
separator = tk.Entry(root, width=80)
out_path = tk.Entry(root, width=80)
display_name = tk.Text(root, height=4, width=60)

dir_path.grid(row=0, column=1)
s_index.grid(row=1, column=1)
header.grid(row=2, column=1)
csv_name.grid(row=3, column=1)
separator.grid(row=4, column=1)
out_path.grid(row=5, column=1)
display_name.grid(row=6, column=1)


def submit():
    # get input from user
    dir_path_id = dir_path.get()
    s_index_id = s_index.get()
    header_id = header.get()
    csv_name_id = csv_name.get()
    separator_id = separator.get()
    out_path_id = out_path.get()

    # go to input directory
    os.chdir(str(dir_path_id))
    files = os.listdir()

    for file in files:
        try:
            show_up = 'Converting file: ' + file + '\n'
            display_name.insert(tk.END, show_up)

            # check file extension
            if file.endswith('xlsx'):
                engine = 'openpyxl'
                ext = 'xlsx'
            elif file.endswith('xls'):
                engine = 'xlrd'
                ext = 'xls'
            elif file.endswith('xlsb'):
                engine = 'pyxlsb'
                ext = 'xlsb'

            # import txt to pandas
            df = pd.read_excel(file, sheet_name=(int(s_index_id) - 1), header=(int(header_id) - 1), engine=engine)

            # rename txt to csv
            file = file.replace(ext, str(csv_name_id))

            # export to clean csv
            df.to_csv(os.path.join(str(out_path_id), file), sep=separator_id,
                      index=False, encoding='utf-8-sig'
                      )
        except:
            continue

    # Display "Done" message
    tk.messagebox.showinfo("excelToCSV", "Done!")


tk.Button(root, text="Submit", font=font_big, command=submit).grid(column=1)

root.resizable(False, False)
root.mainloop()