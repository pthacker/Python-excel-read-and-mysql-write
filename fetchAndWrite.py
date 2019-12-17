import pymysql
pymysql.install_as_MySQLdb()
import xlrd
# Open the workbook and define the worksheet
book = xlrd.open_workbook("beginner_assignment01.xlsx")

sheet = book.sheet_by_name("product_listing")

sheet2 = book.sheet_by_name("group_listing")

# Establish a MySql connection
database = pymysql.connect (host="localhost", user = "root", passwd = "root", db = "mysqlPython")

# Get the cursor to traverse the database line by line
cursor = database.cursor()

# Create the INSERT INTO sql query
query = """INSERT INTO product (product_name, model_name, product_serial_no, group_associa, price) VALUES (%s, %s, %s, %s, %s)"""


# Create a For loop to iterate through each row in the XLS file, starting at row 2 to skip the headers
for r in range(1, sheet.nrows):
    product_name= sheet.cell(r,0).value
    model_name= sheet.cell(r,1).value
    product_serial_no= sheet.cell(r,2).value
    group_associa= sheet.cell(r,3).value
    price= sheet.cell(r,4).value

    # Assign values from each row
    values = (product_name, model_name, product_serial_no, group_associa, price)

    # Execute sql Query
    cursor.execute(query,values)
    #print(values)

#for group_listing worksheet

query2 = """INSERT INTO group2 (group_name, group_desc, isActive) VALUES (%s, %s, %s)"""

def str_to_bool(s):
    if s == 'yes' :
         return True
    elif s == 'no':
         return False
    else:
         return False


for r in range(1, sheet2.nrows):
    group_name= sheet2.cell(r,0).value
    group_desc= sheet2.cell(r,1).value
    isActive= sheet2.cell(r,2).value

    # Assign values from each row
    values = (group_name, group_desc, str_to_bool(isActive))

    # Execute sql Query
    cursor.execute(query2,values)
    #print(values)







# Close the cursor
cursor.close()

# Commit the transaction
database.commit()

# Close the database connection
database.close()

# Print results
print("")
print("All Done! .")
print("")



