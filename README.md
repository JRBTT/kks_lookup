# Information
kks_lookup tool is designed to look for channel addresses in the excel sheet and extract the corresponding kks and signal names.
It requires a very specific template sheet looking for the Adr. column and iterates through the column until no address is seen.
The result is stored in a pandas dataframe and written to an output excel file.
A simple TKinter gui allows for picking the excel sheet file and relevant sheet.
File format must be workbook xslx, strict xslx format is not compatible with pandas.
