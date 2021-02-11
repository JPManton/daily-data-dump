from functions import load_from_mssql
from functions import send_email
import pandas as pd
import numpy as np
from datetime import datetime
import os

#### Import the data from SQL - [AFFILIATE].[affiliate_data_dump] #################################################

# SQL Code to bring in data
str = "select * from [AFFILIATE].[affiliate_data_dump]"
qdata = load_from_mssql(str)
# Change domain and date to text fields
qdata["Date"] = qdata["Date"].astype("str")
qdata["domain"] = qdata["domain"].astype("str")

print("Step 1 Complete: Data loaded in from SQL")

#### Rework data into the desired format, one for each brand ######################################################

# Set the desired column order for final table
column_order = ['Date', 'Avantlink', 'AWIN', 'CJ', 'eBay', 'Impact', 'Bikmo', 'Yellow Jersey', 'Seopa',
				'Tradedoubler', 'Webgains', 'Amazon fees', 'Amazon bounties', 'Amazon (M101)',
				'Amazon fees (US)', 'Amazon bounties (US)', 'Amazon (M101) (US)', 'Amazon fees (AU)',
				'Amazon bounties (AU)', 'Amazon (M101) (AU)', 'M101 (UK)', 'M101 (US)', 'M101 (EU)',
				'M101 (CA)', 'M101 (CHF)', 'M101 (AUD)', 'Skimlinks']

# Pivot the data to get the affiliate partners as columns
qdata_pivoted = pd.pivot_table(qdata,
							   index=["domain", "Date"],
							   columns="AFFILIATE_PARTNER",
							   values="commission_amount",
							   aggfunc=np.sum,
							   fill_value=0,
							   dropna=False
							   )

# Get all domains from base data for the main loop
all_domains = qdata["domain"].unique().tolist()

print("Step 2 Complete: Data formatting")

# Set up the xlwriter to enable saving each dataframe into a sheet
current_timestamp = datetime.now().strftime("%d_%b_%Y_%H_%M_%S")
path = "./todays_excel_file/"
if not os.path.exists(path):
	os.makedirs(path)

file_name = "./todays_excel_file/data_dump_" + current_timestamp + ".xlsx"
# "C:/Users/jonathan.manton/OneDrive - Immediate Media/Projects/
# Affiliates Reporting/data_dump_" + current_timestamp + ".xlsx"
xlwriter = pd.ExcelWriter(file_name)

for dom in all_domains:
	# Filter to only the current domain
	qdata_dom = qdata_pivoted.loc[dom, :]
	# Get a list of the column headers in the current data
	current_columns = list(qdata_dom.columns.values)
	# Loop through the column order list and, if the column header is not already in the data, add it in with 0's
	for col in column_order:
		if col not in current_columns:
			qdata_dom[col] = 0

	# Reorder the columns
	qdata_dom = qdata_dom[column_order]
	# Generate sheet name
	sheet_name = dom.replace("www.", "").replace(".com", "").replace("/", "").replace("*", "").replace("?", "") \
		.replace(":", "").replace("[", "").replace("]", "")

	# Export the current table to the excel file
	qdata_dom.to_excel(xlwriter, sheet_name=sheet_name, index=True)

# Close the excel writer now all tabs are complete
xlwriter.close()

print("Step 3 Complete: Creation of excel file of data")

# Send the email to vickys team with the excel file
send_email("vicky.bruce@immediate.co.uk",
		   "Affiliate Daily Data Dump",
		   "Hi Vicky, this is the affiliate revenue data from yesterday backwards. Any issues please ask.",
		   "./todays_excel_file/data_dump_" + current_timestamp + ".xlsx"
		   )

print("Step 4 Complete: Sending the email to Vicky")

# Send the email to vickys team with the excel file
send_email("Glenn.Caldecott@immediate.co.uk",
		   "Affiliate Daily Data Dump",
		   "Hi Glenn, this is the affiliate revenue data from yesterday backwards. Any issues please ask.",
		   "./todays_excel_file/data_dump_" + current_timestamp + ".xlsx"
		   )

print("Step 5 Complete: Sending the email to Glenn")

# Send the email to vickys team with the excel file
send_email("Aakriti.Wadhwani@immediate.co.uk",
		   "Affiliate Daily Data Dump",
		   "Hi Aakriti, this is the affiliate revenue data from yesterday backwards. Any issues please ask.",
		   "./todays_excel_file/data_dump_" + current_timestamp + ".xlsx"
		   )

print("Step 5 Complete: Sending the email to Aakriti")
