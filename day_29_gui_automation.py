##day-29_GUI_automation
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import os


# -----------------------------
# Global Variables
# -----------------------------
input_file = ""
output_folder = ""


# -----------------------------
# Browse Input Excel File
# -----------------------------
def browse_input():
    global input_file

    file_path = filedialog.askopenfilename(
        title="Select Raw Excel File",
        filetypes=[
            ("Excel Files", "*.xlsx *.xls"),
            ("All Files", "*.*")
        ]
    )

    if file_path:
        input_file = file_path
        input_label.config(
            text=os.path.basename(file_path)
        )


# -----------------------------
# Browse Output Folder
# -----------------------------
def browse_output():
    global output_folder

    folder_path = filedialog.askdirectory(
        title="Select Output Folder"
    )

    if folder_path:
        output_folder = folder_path
        output_label.config(text=folder_path)


# -----------------------------
# Clean Data
# -----------------------------
def clean_data():

    # Check input file
    if not input_file:
        messagebox.showwarning(
            "Missing Input",
            "Please select an Excel input file."
        )
        return

    # Check output folder
    if not output_folder:
        messagebox.showwarning(
            "Missing Output Folder",
            "Please select an output folder."
        )
        return

    # Check if at least one option is selected
    if not any([
        duplicate_var.get(),
        blank_var.get(),
        spaces_var.get(),
        title_var.get()
    ]):
        messagebox.showwarning(
            "No Option Selected",
            "Please select at least one data-cleaning option."
        )
        return

    try:
        # Read Excel file
        df = pd.read_excel(input_file)

        # -----------------------------
        # 1. Remove Duplicate Rows
        # -----------------------------
        if duplicate_var.get():
            df = df.drop_duplicates()

        # -----------------------------
        # 2. Remove Blank Rows
        # -----------------------------
        if blank_var.get():
            df = df.dropna(how="all")

        # -----------------------------
        # 3. Remove Leading/Trailing Spaces
        #    From Text Columns Only
        # -----------------------------
        if spaces_var.get():

            text_columns = df.select_dtypes(
                include=["object", "string"]
            ).columns

            for column in text_columns:
                df[column] = df[column].apply(
                    lambda x: x.strip() if isinstance(x, str) else x
                )

        # -----------------------------
        # 4. Title Text Style
        #    First Letter Capital
        #    For Each Word
        # -----------------------------
        if title_var.get():

            text_columns = df.select_dtypes(
                include=["object", "string"]
            ).columns

            for column in text_columns:
                df[column] = df[column].apply(
                    lambda x: x.title() if isinstance(x, str) else x
                )

        # -----------------------------
        # Create Output File
        # -----------------------------
        output_file = os.path.join(
            output_folder,
            "Cleaned_Data.xlsx"
        )

        # Export cleaned data
        df.to_excel(
            output_file,
            index=False
        )

        # Success message
        messagebox.showinfo(
            "Success",
            f"Data cleaning completed successfully!\n\n"
            f"Cleaned file saved at:\n{output_file}"
        )

    except Exception as e:

        messagebox.showerror(
            "Error",
            f"Something went wrong:\n\n{str(e)}"
        )


# -----------------------------
# Reset Paths
# -----------------------------
def reset_paths():

    global input_file, output_folder

    input_file = ""
    output_folder = ""

    input_label.config(
        text="No input file selected"
    )

    output_label.config(
        text="No output folder selected"
    )


# -----------------------------
# Reset Checkboxes
# -----------------------------
def reset_checkboxes():

    duplicate_var.set(False)
    blank_var.set(False)
    spaces_var.set(False)
    title_var.set(False)


# =====================================================
# GUI WINDOW
# =====================================================

root = tk.Tk()

root.title("Dynamic Excel Data Cleaning Automation")
root.geometry("650x600")
root.resizable(False, False)


# -----------------------------
# Heading
# -----------------------------
heading = tk.Label(
    root,
    text="Dynamic Excel Data Cleaning Automation",
    font=("Arial", 20, "bold")
)

heading.pack(pady=20)


# -----------------------------
# Input File Section
# -----------------------------
input_frame = tk.Frame(root)

input_frame.pack(pady=10)

input_button = tk.Button(
    input_frame,
    text="Browse Excel File",
    width=20,
    command=browse_input
)

input_button.grid(
    row=0,
    column=0,
    padx=10
)

input_label = tk.Label(
    input_frame,
    text="No input file selected",
    width=40,
    anchor="w"
)

input_label.grid(
    row=0,
    column=1,
    padx=10
)


# -----------------------------
# Output Folder Section
# -----------------------------
output_frame = tk.Frame(root)

output_frame.pack(pady=10)

output_button = tk.Button(
    output_frame,
    text="Browse Output Folder",
    width=20,
    command=browse_output
)

output_button.grid(
    row=0,
    column=0,
    padx=10
)

output_label = tk.Label(
    output_frame,
    text="No output folder selected",
    width=40,
    anchor="w"
)

output_label.grid(
    row=0,
    column=1,
    padx=10
)


# -----------------------------
# Cleaning Options
# -----------------------------
options_label = tk.Label(
    root,
    text="Select Data Cleaning Options",
    font=("Arial", 14, "bold")
)

options_label.pack(pady=20)


# Variables
duplicate_var = tk.BooleanVar()
blank_var = tk.BooleanVar()
spaces_var = tk.BooleanVar()
title_var = tk.BooleanVar()


# Checkboxes
duplicate_check = tk.Checkbutton(
    root,
    text="Remove Duplicate Rows",
    variable=duplicate_var,
    font=("Arial", 11)
)

duplicate_check.pack(
    anchor="w",
    padx=180,
    pady=5
)


blank_check = tk.Checkbutton(
    root,
    text="Remove Blank Rows",
    variable=blank_var,
    font=("Arial", 11)
)

blank_check.pack(
    anchor="w",
    padx=180,
    pady=5
)


spaces_check = tk.Checkbutton(
    root,
    text="Remove Leading & Trailing Spaces (Text Columns Only)",
    variable=spaces_var,
    font=("Arial", 11)
)

spaces_check.pack(
    anchor="w",
    padx=180,
    pady=5
)


title_check = tk.Checkbutton(
    root,
    text="Convert Text to Title Case",
    variable=title_var,
    font=("Arial", 11)
)

title_check.pack(
    anchor="w",
    padx=180,
    pady=5
)


# -----------------------------
# Buttons
# -----------------------------
button_frame = tk.Frame(root)

button_frame.pack(pady=30)


clean_button = tk.Button(
    button_frame,
    text="Clean Data",
    width=18,
    height=2,
    command=clean_data
)

clean_button.grid(
    row=0,
    column=0,
    padx=10
)


reset_paths_button = tk.Button(
    button_frame,
    text="Reset Paths",
    width=18,
    height=2,
    command=reset_paths
)

reset_paths_button.grid(
    row=0,
    column=1,
    padx=10
)


reset_checkbox_button = tk.Button(
    button_frame,
    text="Reset Checkboxes",
    width=18,
    height=2,
    command=reset_checkboxes
)

reset_checkbox_button.grid(
    row=0,
    column=2,
    padx=10
)


# -----------------------------
# Start GUI
# -----------------------------
root.mainloop()