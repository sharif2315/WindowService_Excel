import pandas as pd
import pyodbc

# reading xlsx file
myfile = r'C:\Users\sahmed243\Documents\WebDevelopment\Python\Scripts\excel_Import\sampledata.xlsx'

# creating a dataframe to store xlsx data
df = pd.read_excel(myfile, sheet_name='Sheet2')

# printing the dataframe
print(df)

# connecting to DB
connStr = pyodbc.connect(r'DRIVER={ODBC Driver 13 for SQL Server};SERVER=GBPF0Y89C1\SQLEXPRESS;DATABASE=JobListings;Trusted_Connection=yes')
cursor = connStr.cursor()

# Inserting data from df to DB
for index,row in df.iterrows():
    cursor.execute("INSERT INTO dbo.testTable([fname],[email],[grade],[age]) values (?,?,?,?)",
        row['fname'], row['email'] , row['grade'], row['age'])
    connStr.commit()

# closing DB conn
cursor.close()
connStr.close()

print('Data imported to Database successfully!')
