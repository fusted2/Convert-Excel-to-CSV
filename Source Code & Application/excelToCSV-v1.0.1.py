import tkinter as tk
import tkinter.messagebox
import os
import pandas as pd

# Create an instance of Tkinter frame
font_big = "Calibri 14"
font_small = "Calibri 11"
root = tk.Tk()
root.title('excelToCSV (v1.0.1)  |  © NGUYEN HOANG PHU, 2023')

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
display_name = tk.Text(root, height=6, width=60)

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
    csv_name_id = csv_name.get()
    separator_id = separator.get()
    out_path_id = out_path.get()

    # user input validation & error handling
    try:
        files = os.listdir(dir_path_id)
    except FileNotFoundError:
        tk.messagebox.showerror("Error", "Input folder not found")
        return

    try:
        s_index_id = int(s_index.get()) - 1
        header_id = int(header.get()) - 1
        if s_index_id < 0 or header_id < 0:
            raise ValueError("Sheet index and header must be positive integers.")
    except ValueError:
        tk.messagebox.showerror("Error", "Sheet index and header must be positive integers.")
        return

    if csv_name_id not in ['csv', 'txt']:
        tk.messagebox.showerror("Error", "CSV/TXT must be either 'csv' or 'txt'.")
        return

    if not separator_id:
        tk.messagebox.showerror("Error", "Separator for output not found.")
        return

    if not os.path.exists(out_path_id):
        tk.messagebox.showerror("Error", "Output folder not found")
        return

    supported_files = [f for f in files if f.endswith(('xlsx', 'xls', 'xlsb'))]
    unsupported_files = [f for f in files if not f.endswith(('xlsx', 'xls', 'xlsb'))]

    if unsupported_files:
        display_name.insert(tk.END, f"Unsupported file types found: {', '.join(unsupported_files)}. " '\n'
                                         "The application only supports .xlsx, .xlsb, .xls files." '\n\n')


    # main function
    for file in supported_files:
        try:
            # update the loop tasks, rather than displaying "Not responding" msg
            show_up = 'Converting file: ' + file + '\n'
            display_name.insert(tk.END, show_up)
            root.update_idletasks()

            # Determine engine and extension
            if file.endswith('xlsx'):
                engine = 'openpyxl'
            elif file.endswith('xls'):
                engine = 'xlrd'
            elif file.endswith('xlsb'):
                engine = 'pyxlsb'
            else:
                continue

            # Read Excel file
            df = pd.read_excel(os.path.join(dir_path_id, file), sheet_name=s_index_id, header=header_id, engine=engine)

            # Rename file to CSV/TXT
            csv_file = file.rsplit('.', 1)[0] + '.' + csv_name_id

            # Export to CSV/TXT
            df.to_csv(os.path.join(out_path_id, csv_file), sep=separator_id, index=False, encoding='utf-8-sig')

        except Exception as e:
            show_up = 'Failed to convert file: ' + file + '\n' + str(e) + '\n'
            display_name.insert(tk.END, show_up)

    # Display "Done" message
    tk.messagebox.showinfo("excelToCSV", "Done!")


tk.Button(root, text="Submit", font=font_big, command=submit).grid(column=1)

root.resizable(False, False)
root.mainloop()