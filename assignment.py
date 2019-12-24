import xlrd
import pymysql

# Open the workbook and define the worksheet
book = xlrd.open_workbook("assignment01.xlsx")
sheet = book.sheet_by_name("product_listing")
print(sheet)

# Establish a MySQL connection
database = pymysql.connect (host="localhost", user = "root", passwd = "9742723938", db = "mysql")

# Get the cursor, which is used to traverse the database, line by line
cursor = database.cursor()

# Create a For loop to iterate through each row in the XLS file, starting at row 2 to skip the headers
for r in range(1, sheet.nrows):
		product_name	  = sheet.cell(r,0).value
		#print(product_name)
		model_name	  = sheet.cell(r,1).value
		#print(model_name)
		product_serial_no = str(sheet.cell(r,2).value)
		#print(product_serial_no)
		group_associated = sheet.cell(r,3).value
		#print(group_associated)
		product_mrp	  = str(sheet.cell(r,4).value)
		#print(product_mrp)
		# Execute sql Query
		cursor.execute('insert into orders3 values("%s","%s","%s","%s","%s")'%(product_name,model_name,product_serial_no,group_associated,product_mrp))

# Close the cursor
cursor.close()

# Commit the transaction
database.commit()

# Close the database connection
database.close()

# Print results
print("")
print("All Done! Bye, for now.")
print("")
columns = (str(sheet.ncols))
rows = (str(sheet.nrows-1))
print("I just imported ",columns," columns and ",rows," rows to MySQL!")
