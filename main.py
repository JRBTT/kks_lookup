import pandas as pd
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import os

jobs = []  # List to store job names
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
        Addr_exist = False
        slot_list = [2, 3, 4, 5, 6, 7, 8, 9]
        first_col_data = data[data.columns[0]]  # Gets the first column by name
        diagram_name = ""
        portname = ""
        portType = ""
        aixx_sEU = ""
        aixx_sLL = ""
        kks_signal = ""
        for idy, column in enumerate(data.columns):
            # print(f"Processing column '{column}'...")
            col_data = data[column]  # Get the column data
            next_column = data.columns[idy + 1] if idy + 1 < len(data.columns) else None
            n_next_column = data.columns[idy + 2] if idy + 2 < len(data.columns) else None
            next_col_data = data[next_column] if next_column else None
            n_next_col_data = data[n_next_column] if n_next_column else ""

            for idx, value in col_data.items():
                if value == "Slot":
                    print(f"Found 'Slot' in column '{column}' at index {idx}")
                    slot_temp = next_col_data[idx]
                    print(f"Slot value: {type(slot_temp)}")
                    if slot_temp in slot_list:
                        slot = int(slot_temp)
                        location = first_col_data[idx]
                        diagram_name = f"{location}{slot:02}"
                        print(f"Diagram name: {diagram_name}")
                        card_type = n_next_col_data[idx + 1]
                        match card_type:
                            case "F_DI":
                                portType = "DI"
                            case "F_DO":
                                portType = "DO"
                            case "F_AI":
                                portType = "AI"
                            case "F_AO":
                                portType = "AO"
                            case _:
                                portType = "NA"
                if value == "Adr.":  # Check if the cell matches "Adr"
                    channel = 0
                    Addr_exist = True
                    print(f"Found 'Adr.' in column '{column}' at index {idx}")
                    aixx_sEU_col = data.columns[idy + 9]
                    aixx_sLL_col = data.columns[idy + 10]
                    aixx_sEU_data = data[aixx_sEU_col]
                    aixx_sLL_data = data[aixx_sLL_col]
                    # Collect all non-NaN values below "Adr"
                    sublist = []
                    for sub_idx in range(idx + 1, len(col_data)):
                        address = col_data[sub_idx]
                        kks = next_col_data[sub_idx]
                        signal = n_next_col_data[sub_idx]
                        if pd.isna(address):  # Stop at the first NaN
                            break
                        elif not pd.isna(kks):
                            if pd.isna(signal):
                                print("signal is empty")
                                signal = ""
                            portname = f"{portType}{channel:02}_PV"
                            kks_signal = f"{kks}|{signal}"
                            aixx_sEU = ""
                            aixx_sLL = ""
                            if portType == "AI":
                                aixx_sEU = aixx_sEU_data[sub_idx]
                                aixx_sLL = aixx_sLL_data[sub_idx]
                            result.append([diagram_name, portname, kks, signal, kks_signal, address, aixx_sEU, aixx_sLL])
                        channel += 1
        if not Addr_exist:
            return "1"
        if not result:
            return "2"         
        # Print the result
        print("Collected values below 'Adr':")
        print(result)
    except FileNotFoundError:
        print(f"Error:'{file_path}' was not found.")
    except Exception as e:
        print(f"An error occurred: {e}")
    return result
def write_to_excel(result, destination):
    try:
        # Convert the result list to a DataFrame
        result_df = pd.DataFrame(result, columns=["Diagram- Name", "Port- Name", "KKS", "Signal", "KKS_Signal", "Address", "AIxx_sEU","AIxx_sLL" ])
        
        # Create the initial output file name using the sheet name
        output_file = os.path.join(destination, f"OUTPUT.xlsx")
        counter = 1
        while os.path.exists(output_file):
            # If the file exists, create a new file name with a counter
            output_file = os.path.join(destination, f"OUTPUT_{counter}.xlsx")
            counter += 1
        result_df.to_excel(output_file, index=False)  # Save without the index column
        print(f"Result successfully written to {output_file}")
        current_dir = os.getcwd()
        file_path = os.path.join(current_dir, output_file)
        return file_path
    except Exception as e:
        print(f"An error occurred while writing to Excel: {e}")

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
        messagebox.showerror("Error", f"Failed to load sheet names: {e} Check if the file is saved as strict xlsx format. Resave as normal xlsx format and try again.")


def update_text_field(added_sheet):
    """Update the text field with added sheet names."""
    text_field.config(state="normal")  # Temporarily enable editing
    # text_field.delete("1.0", tk.END)  # Clear the text field
    text_field.insert(tk.END, added_sheet + "\n")  # Insert updated sheet names
    text_field.config(state="disabled")  # Make it read-only again


def on_add():
    """Handle the add button click."""
    file_path = excel_path_var.get()
    print(file_path)
    sheet_name = sheet_name_var.get()
    if not file_path or not sheet_name:
        messagebox.showwarning("Warning", "Please select a file and a sheet name.")
        return
    else:
        text_field_content = text_field.get("1.0", tk.END).strip()  # Get all text from the text field
        if sheet_name in text_field_content.split("\n"):  # Check if the sheet name exists
            messagebox.showinfo("Info", f"'{sheet_name}' is already added.")
            return
        update_text_field(sheet_name)  # Update the text field with the added sheet name
        jobs.append([file_path, sheet_name])

def browse_destination():
    """Open a file dialog to select a destination folder."""
    folder_path = filedialog.askdirectory(title="Select Destination Folder")
    if folder_path:
        destination_path_var.set(folder_path)  # Set the folder path in the entry widget

def on_submit():
    """Handle the submit button click"""
    print(jobs)
    if not jobs:
        messagebox.showwarning("Warning", "No jobs have been added.")
        return
    # Create a new window for displaying messages
    log_window = tk.Toplevel(root)
    log_window.title("Processing Log")
    log_window.geometry("800x500")

    # Add a text field to the new window
    log_text = tk.Text(log_window, height=15, width=60, state="normal", wrap="word")
    log_text.pack(padx=10, pady=10)

    # Function to log messages
    def log_message(message):
        log_text.config(state="normal")
        log_text.insert(tk.END, message + "\n")
        log_text.config(state="disabled")
        log_text.see(tk.END)  # Scroll to the end

    destination = destination_path_var.get()
    big_list = []
    for file_path, sheet_name in jobs:
        log_message(f"Processing file: {file_path}, sheet: {sheet_name}")
        result = extractor(file_path, sheet_name)
        if result == "1":
            log_message(f"Warning: No 'Adr.' column found in sheet '{sheet_name}'. Skipping.")
        elif result == "2":
            log_message(f"Warning: No data returned for sheet '{sheet_name}'. Skipping.")
        else:
            log_message(f"Success: {sheet_name} processed successfully.")
            big_list.extend(result)  # Append the result to the big list
    log_message("writing to excel file")
    if not big_list:
        messagebox.showwarning("Warning", "No data to write to Excel.")
        return
    write_to_excel(big_list, destination)  # Write the big list to Excel
    log_message(f"output file created at {destination}")

def on_clear():
    """Handle the clear button click"""
    jobs.clear()  # Clear the jobs list
    text_field.config(state="normal")  # Temporarily enable editing
    text_field.delete("1.0", tk.END)  # Clear the text field

# Create the main tkinter window
root = tk.Tk()
root.title("KKS Extractor")
root.resizable(False, False)

# File path input
tk.Label(root, text="Excel File:").grid(row=0, column=0, padx=10, pady=10, sticky="w")
excel_path_var = tk.StringVar()
tk.Entry(root, textvariable=excel_path_var, width=50, state="readonly").grid(row=0, column=1, padx=10, pady=10)
tk.Button(root, text="Browse", command=browse_file).grid(row=0, column=2, padx=10, pady=10)

# Destination folder input
tk.Label(root, text="Destination Folder:").grid(row=1, column=0, padx=10, pady=10, sticky="w")
destination_path_var = tk.StringVar()
tk.Entry(root, textvariable=destination_path_var, width=50, state="readonly").grid(row=1, column=1, padx=10, pady=10)
tk.Button(root, text="Browse", command=browse_destination).grid(row=1, column=2, padx=10, pady=10)

# Sheet name dropdown
tk.Label(root, text="Sheet Name:").grid(row=2, column=0, padx=10, pady=10, sticky="w")
sheet_name_var = tk.StringVar()
sheet_name_dropdown = ttk.Combobox(root, textvariable=sheet_name_var, state="readonly", width=47)
sheet_name_dropdown.grid(row=2, column=1, padx=10, pady=10)

# Add button
tk.Button(root, text="Add", command=on_add).grid(row=2, column=2, pady=20)

# Text field to display added sheet names
text_field_var = tk.StringVar()
tk.Label(root, text="Added Sheet Names:").grid(row=3, column=0, padx=10, pady=10, sticky="nw")
text_field = tk.Text(root, height=5, width=50, state="disabled", wrap="word")
text_field.grid(row=3, column=1, padx=10, pady=10)

# Create a frame to hold the buttons
button_frame = tk.Frame(root)
button_frame.grid(row=4, column=1, pady=20)

# Submit button for processing all sheets
tk.Button(button_frame, text="Clear", command=on_clear).grid(row=0, column=0, padx=5)

# Submit button for processing all sheets
tk.Button(button_frame, text="Submit", command=on_submit).grid(row=0, column=1, padx=5)



# Run the tkinter main loop
root.mainloop()

