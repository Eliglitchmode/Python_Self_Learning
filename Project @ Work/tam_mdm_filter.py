import pandas as pd
from openpyxl import load_workbook
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog as fd
from tkinter.messagebox import showinfo
import os
import xlsxwriter




# create the root window
root = tk.Tk()
root.title('Select TAM Needed Report for uploading:')
root.resizable(False, False)
root.geometry('250x100')


# select and load csv file
def select_file():
	filetypes = (
		('CSV files', '*.csv'),
		('All files', '*.*')
	)

	global file_name
	file_name = fd.askopenfilename(
		initialdir='/',
		filetypes=filetypes)

	if file_name != "":
		root.after(500,lambda:root.destroy())
		print('File Successfully Loaded: ', file_name)

# open button
open_button = ttk.Button(
	root,
	text='Open a File',
	command=select_file
)

open_button.pack(expand=True)

# run the application
root.mainloop()



# generate a list of MDMs for filtering
lines = []
while True:
	try:
		line = input('Please type or press CTRL+V to paste your MDM#, then enter "q" to proceed: ')
		if line == "q":
			break
		else:
			lines.append(int(line))
	except EOFError:
		break
		print('Please input valid MDM#, and click CTRL+Z and Enter to next step.')
	except ValueError:
		break
		print('Please input valid MDM# and click CTRL+Z and Enter to next step')

# OEM assignment
OEM = input('Which OEM you are assigning to? ')
while True:
	if OEM != "":

		break
	else:
		print('Please input a valid OEM in upper case!')		 

print('Result: ')
print(lines)
print('These MDMs are assigned to: ', OEM.upper())


# filter csv by specified MDM# and insert OEM
base_name = os.path.basename(file_name)
#sheet_name = base_name.replace(".csv", "")
df = pd.read_csv(str(file_name), names=['EG', 'Code', 'Plan_PropertyGroup_Name', 'Plan_ResourceType_Name', 'Workload', 'Plan_SKU_Name',	'PFAM', 'Plan_Geo_Name', 'Plan_Region_Name', 'Plan_DC_Code', 'State_Name', 'Plan_Intent_Name', 'CapEx_Amount', 'Plan_NumberOfRacks', 'CSP_TotalServers', 'Calculated_PDD', 'Fiscal_Year', 'Fiscal_Quarter', 'Fiscal_Month', 'AwardTargetCSP', 'Award_Target_Date', 'Award_Lead_Time_Days', 'SC_OEM', 'ResourceDesignId', 'PlatformFamilyId', 'PlatformFamilyName', 'AlternatePlatformFamilyId', 'AlternatePlatformFamilyName', 'ResourceDesignStatusId', 'ResourceDesignStatusName'
], header=0)

df_filter = df[df["Code"].isin(lines)]
print(len(df_filter))
print(len(lines))
df2 = df_filter.copy()
df2.loc[:, ['SC_OEM']] = OEM.upper()


# save a copy of result and tabular template in excel format
Output_name = "output.xlsx"
with pd.ExcelWriter(Output_name) as writer:  
	df.to_excel(writer, sheet_name='Original', index=False)
	df2.to_excel(writer, sheet_name='Upload', index=False)
	workbook  = writer.book
	worksheet = writer.sheets['Upload']
	column_settings = [{'header': column} for column in df2.columns]
	(max_row, max_col) = df2.shape
	worksheet.add_table(0, 0, max_row, max_col - 1, {'columns': column_settings})
	worksheet.set_column(0, max_col - 1, 12)

# pop up message upon completion and auto-launch output xlsx
print('Work Done!')
print(len(df_filter), ' results are found and filtered of ', len(lines), ' MDMs.')
print(len(lines) - len(df_filter), ' MDMs are missing!')
os.startfile(Output_name)
