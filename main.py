import pandas as pd

# Step 1: Read the file name from config.txt
with open("config.txt", "r") as config_file:
    excel_file_name = config_file.readline().strip()
    sheet_name = config_file.readline().strip()      # Second line: Sheet name

# Step 2: Load the Excel file using pandas
try:
    # Assuming the Excel file is in the same directory as the script
    data = pd.read_excel(excel_file_name, sheet_name=sheet_name)
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
        result_df.to_excel(output_file, index=False)  # Save without the index column
        print(f"Result successfully written to {output_file}")
    except Exception as e:
        print(f"An error occurred while writing to Excel: {e}")


except FileNotFoundError:
    print(f"Error: The file '{excel_file_name}' was not found.")
except Exception as e:
    print(f"An error occurred: {e}")