import openpyxl as openpyxl
import pymysql.cursors
import decimal
import datetime

# Define a converter function to convert pymysql data types to Python built-in types

def convert_pymysql_types(data):
    if isinstance(data, (bytearray, bytes)):
        return data.decode('utf-8')
    elif isinstance(data, (int, float)):
        return data
    elif isinstance(data, decimal.Decimal):
        return float(data)
    elif isinstance(data, (datetime.date, datetime.datetime)):
        return data.isoformat()
    elif data is None:
        return None
    else:
        return str(data)


# Load the Excel sheet and extract the SQL statements
wb = openpyxl.load_workbook('query_excel_file.xlsx')
ws = wb.active
queries = []
for cell in ws['A']:
    if cell.value.startswith('SELECT') or cell.value.startswith('select'):
        queries.append(cell.value)

# Define the database connection parameters
host = '127.0.0.1'
user = 'root'
password = 'your-password'
database = 'sql_invoicing'

# Connect to the database
conn = pymysql.connect(host=host,
                       user=user,
                       password=password,
                       db=database,
                       # cursorclass=pymysql.cursors.DictCursor)
                       )

x = 1
# Execute a SELECT statement
for query in queries:
    with conn.cursor() as cur:
        # query = "SELECT * FROM sql_invoicing.invoices ORDER BY Invoice_Total DESC LIMIT 4;"
        cur.execute(query)
        results = cur.fetchall()

    # Convert the data types in the query results and store in a list of lists
    rows = []
    for result in results:
        row = []
        for col in result:
            row.append(str(convert_pymysql_types(col)))
        rows.append(row)

    # Get the column names from the keys of the first dictionary in the query results
    col_names = list(results)

    # Print the query results
    # print(col_names)
    print(f'---------------Q{x}----------------------')
    master_list = []
    x += 1
    for row in rows:
        # print(row)
        master_list.append(row)

    print(master_list, '\n')

# Close the database connection
conn.close()
