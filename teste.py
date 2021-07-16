import pandas as pd
df = pd.DataFrame({'Name': ['E','F','G','H'],
                   'Age': [100,70,40,60]})
# Create writer object with an engine xlsxwriter
writer = pd.ExcelWriter('demo.xlsx', engine='xlsxwriter')
# Write data to an excel
df.to_excel(writer,sheet_name="Sheet1",index=False)
# Get sheet for conditional formatting 
worksheet = writer.sheets['Sheet1']

# Close writer
writer.close()