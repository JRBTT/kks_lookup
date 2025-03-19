import pandas as pd
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import os


def extractor(file_path, sheet_name):
        # Step 1: Read the file name from config.txt for debugging
    # with open("config.txt", "r") as config_file:
    #     excel_file_name = config_file.readline().strip()
    #     sheet_name = config_file.readline().strip()      # Second line: Sheet name

    # Step 2: Load the Excel file using pandas
    try:
        # Assuming the Excel file is in the same directory as the script
        data = pd.read_excel(file_path, sheet_name=sheet_name)
        print("Excel file loaded successfully!")

        # Step 3: Process each column to find "Adr" and collect values below it
        result = []  # List to store the collected values
        for idy, column in enumerate(data.columns):
            # print(f"Processing column '{column}'...")
            col_data = data[column]  # Get the column data
            next_column = data.columns[idy + 1] if idy + 1 < len(data.columns) else None
            n_next_column = data.columns[idy + 2] if idy + 2 < len(data.columns) else None
            next_col_data = data[next_column] if next_column else None
            n_next_col_data = data[n_next_column] if n_next_column else ""

            for idx, value in col_data.items():
                if value == "Adr.":  # Check if the cell matches "Adr"
                    print(f"Found 'Adr.' in column '{column}' at index {idx}")
                    # Collect all non-NaN values below "Adr"
                    sublist = []
                    for sub_idx in range(idx + 1, len(col_data)):
                        address = col_data[sub_idx]
                        kks = next_col_data[sub_idx]
                        signal = n_next_col_data[sub_idx]
                        if pd.isna(address):  # Stop at the first NaN
                            break
                        elif pd.isna(kks):
                            break
                        else:
                            kks_signal = kks + "|" + signal
                            result.append([kks, signal, kks_signal, address])
            
        # Print the result
        print("Collected values below 'Adr':")
        print(result)

        try:
            # Convert the result list to a DataFrame
            result_df = pd.DataFrame(result, columns=["KKS", "Signal", "KKS_Signal", "Address"])

            # Write the DataFrame to an Excel file
            output_file = "output.xlsx"  # Specify the output file name
            counter = 1
            while os.path.exists(output_file):
                # If the file exists, create a new file name with a counter
                output_file = f"output_{counter}.xlsx"
                counter += 1
            result_df.to_excel(output_file, index=False)  # Save without the index column
            print(f"Result successfully written to {output_file}")
            current_dir = os.getcwd()
            file_path = os.path.join(current_dir, output_file)
            return file_path
        except Exception as e:
            print(f"An error occurred while writing to Excel: {e}")


    except FileNotFoundError:
        print(f"Error:'{file_path}' was not found.")
    except Exception as e:
        print(f"An error occurred: {e}")

def browse_file():
    """Open a file dialog to select an Excel file."""
    file_path = filedialog.askopenfilename(
        filetypes=[("Excel files", "*.xlsx *.xls")],
        title="Select an Excel File"
    )
    if file_path:
        excel_path_var.set(file_path)  # Set the file path in the entry widget
        load_sheet_names(file_path)  # Load sheet names into the dropdown

def load_sheet_names(file_path):
    """Load sheet names from the selected Excel file."""
    try:
        sheet_names = pd.ExcelFile(file_path).sheet_names  # Get sheet names
        sheet_name_dropdown['values'] = sheet_names  # Populate the dropdown
        if sheet_names:
            sheet_name_var.set(sheet_names[0])  # Set the first sheet as default
    except Exception as e:
        messagebox.showerror("Error", f"Failed to load sheet names: {e}")

def on_submit():
    """Handle the submit button click."""
    file_path = excel_path_var.get()
    print(file_path)
    sheet_name = sheet_name_var.get()
    if not file_path or not sheet_name:
        messagebox.showwarning("Warning", "Please select a file and a sheet name.")
        return
    else:
        file_path = extractor(file_path, sheet_name)
        messagebox.showinfo("Success, output file created at: ", file_path)

# Create the main tkinter window
root = tk.Tk()
root.title("KKS Extractor")

# File path input
tk.Label(root, text="Excel File:").grid(row=0, column=0, padx=10, pady=10, sticky="w")
excel_path_var = tk.StringVar()
tk.Entry(root, textvariable=excel_path_var, width=50).grid(row=0, column=1, padx=10, pady=10)
tk.Button(root, text="Browse", command=browse_file).grid(row=0, column=2, padx=10, pady=10)

# Sheet name dropdown
tk.Label(root, text="Sheet Name:").grid(row=1, column=0, padx=10, pady=10, sticky="w")
sheet_name_var = tk.StringVar()
sheet_name_dropdown = ttk.Combobox(root, textvariable=sheet_name_var, state="readonly", width=47)
sheet_name_dropdown.grid(row=1, column=1, padx=10, pady=10)

# Submit button
tk.Button(root, text="Submit", command=on_submit).grid(row=2, column=1, pady=20)

# Run the tkinter main loop
root.mainloop()

