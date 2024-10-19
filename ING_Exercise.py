import pandas as pd

"""
This script will identify the limit of each ID,
iterate through ID and transaction amounts, 
determine if/how much they are over their transaction limits, 
and output the results to a new sheet.
"""


###  1  GET LIMITS  ###

# Location where original Excel file is stored
path = 'C:\\Users\\thoma\\Documents\\ING\\INGTechnicalExercise_orig.xlsx'

# Location where edited Excel file will be stored
path_updated = 'C:\\Users\\thoma\\Documents\\ING\\INGTechnicalExercise_updated.xlsx'

# Create a dataframe to read Excel file sheet "Limits" into
df_limits  = pd.DataFrame()

df_limits = pd.read_excel(io = path, 
                             sheet_name = 'Limits'
                             )

# Print to verify data
# print(f'The limits data sorted by ID:\n\n {df_limits}')

# Create Dictionary to store IDs and Limits for convenience of searching
# Could use Pandas DataFrame.to_dict in future
limits_dict = {}

for index, row in df_limits.iterrows():
    limits_dict.update({row['Trade ID'] : row['Limit']})

# Verify limits dictionary
# print(f'Limits dictionary from iteration: \n\n{limits_dict}')



###  2  GET TRANSACTION DATA  ###

# Create a DataFrame to read Excel file sheet "Data" into
df_data  = pd.DataFrame()

df_data = pd.read_excel(io = path, 
                             sheet_name = 'Data'
                             )

# Print to Verify Data Sheet
# print(f'The Data of actual transactions:\n\n {df_data}')



### 3 COMPARE TRANSACTION DATA TO SPENDING LIMITS ###

# Create a DataFrame to write final Output to
df_output = pd.DataFrame(columns=['ID', 'Amount Over Limit'])
output_idx = 0

for index, row in df_data.iterrows():
    
    # Check the spending of each unordered ID, one at a time:
    trade_ID = row['Trade ID']
    spending = row['Amount']
    limit = limits_dict.get(trade_ID)
    
    # Verify loop results
    # print(f'For the {index}th time, the Trade ID is: {trade_ID}, and the spending is: {spending}, and the limit is {limit}.')
    # print(f'Spending minus limit is: {spending - limit}')
    
    # If the trade is over the limit (result is > 0), then add to Output DataFrame
    if(spending - limit > 0):
        over_limit_amount = spending - limit
        df_output.loc[output_idx] = [trade_ID, over_limit_amount]

        # Increase index for output DataFrame
        output_idx += 1 
   
# Sort the Output DataFrame by ascending ID
df_output.sort_values(by='ID', inplace=True)

# Convert IDs to INTs
df_output_converted = df_output.astype({'ID': 'int32'})

# Verify final sorted output
#print(f'The DataFrame of overspending is: \n\n {df_output_converted}')

# Write Output to Output sheet in xlsx file
df_output_converted.to_excel(excel_writer=path_updated,
                             sheet_name='Output',
                             float_format='%.2f', 
                             index=False
                             )