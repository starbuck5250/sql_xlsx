# Convert SQL statement to Excel spreadsheet
# Run from PASE using OSS
# Tested on IBM i 7.3
# Buck Calabro
import ibm_db
import sys
import xlsxwriter

# The intended use is to supply a dynamic SQL statement which will return a limited
# number of columns.  Primarily for testing whether the test database is sane

# ================================================
# Close open files, handles, etc
def cleanup():
	wb.close()
	ibm_db.close(connection)
	return;

# ================================================
# Documentation:
# https://github.com/ibmdb/python-ibmdb/wiki/APIs
# https://www.ibm.com/developerworks/community/wikis/home?lang=en#!/wiki/IBM%20i%20Technology%20Updates/page/QSYS2.PARSE_STATEMENT%20UDTF
# https://xlsxwriter.readthedocs.io/ for more info

# Example that should run on all systems:
# python sql_xlsx.py "select option, command from qgpl.qauoopt" "qauoopt.xlsx" "user" "password"
sql_stmt = sys.argv[1]
workbook_file = sys.argv[2]
user = sys.argv[3]
password = sys.argv[4]
print "=============================================================="
print "Processing..."

# ================================================
# Connect to DB2
connection = ibm_db.connect('S104Z8CM', user, password)
if connection == None:
	# error!
	print "Connection error"
	print "SQLSTATE " + ibm_db.conn_error()
	print ibm_db.conn_errormsg()
	exit(-1)

# ================================================
# analyse the SQL statement to extract out the columns
row = 0
column = 0
column_list = []

# need to escape out single quotes for this one
sql_stmt_escaped = sql_stmt.replace("'", "''")

analysis_stmt = "select column_name, sql_statement_type from table(qsys2.parse_statement('" + sql_stmt_escaped + "')) x where name_type = 'COLUMN'"
# debugging
#print sql_stmt
#print analysis_stmt
column_rs = ibm_db.exec_immediate(connection, analysis_stmt)
if column_rs == False:
	# error!
	print "Column list error"
	print "SQLSTATE " + ibm_db.conn_error(connection)
	print ibm_db.conn_errormsg(connection)
	exit(-1)

# ================================================
# Create the empty spreadsheet
wb = xlsxwriter.Workbook(workbook_file)
ws = wb.add_worksheet()
# some style attributes  
normal = wb.add_format({'bold': False})
bold   = wb.add_format({'bold': True})	

# write a textbox with the SQL statement
ws.write(row, column, sql_stmt, normal)
row += 1
	
# ================================================
# Process the list of columns
# These will become the headings	
while ibm_db.fetch_row(column_rs) != False:
	column_name = ibm_db.result(column_rs, "COLUMN_NAME")
	# Debugging
	#print column_name
	sql_statement_type = ibm_db.result(column_rs, "SQL_STATEMENT_TYPE")
	if sql_statement_type != 'QUERY':
		print "SELECT statements only!"
		cleanup()
		exit(-1)
	
	column_list.append(column_name)
	ws.write(row, column, column_name, bold)
	column += 1
	
# If no columns in the list (select * will do this), get the list of columns from the table
if len(column_list) == 0:
	analysis_stmt_table = "select name, schema from table(qsys2.parse_statement('" + sql_stmt_escaped + "')) x where name_type = 'TABLE'"
	# debugging
	#print analysis_stmt
	table_rs = ibm_db.exec_immediate(connection, analysis_stmt_table)
	if table_rs == False:
		# error!
		print "Table list error"
		print "SQLSTATE " + ibm_db.conn_error(connection)
		print ibm_db.conn_errormsg(connection)
		exit(-1)
		
	# process the list of tables
	while ibm_db.fetch_row(table_rs) != False:
		table_name = ibm_db.result(table_rs, "NAME")
		schema_name = ibm_db.result(table_rs, "SCHEMA")
		# Debugging
		#print table_name.strip() + "." + schema_name.strip()

		# now that we have a table (can there be more than one?)
		# go get the columns within it
		table_column_rs = ibm_db.columns(connection, None, schema_name, table_name)
		
		# process the list of columns	
		while ibm_db.fetch_row(table_column_rs) != False:
			column_name = ibm_db.result(table_column_rs, "COLUMN_NAME")
			# Debugging
			#print column_name
			column_list.append(column_name)
			ws.write(row, column, column_name, bold)
			column += 1

# ================================================
# execute the supplied SQL statement
# and write one row for each row in the result set, and one column for each returned column
detail_rs = ibm_db.exec_immediate(connection, sql_stmt)
if detail_rs == False:
	# will apparently not get here; the exec_immediate craps out rather than throw a False
	print "Detail list error"
	print "SQLSTATE " + ibm_db.conn_error(connection)
	print ibm_db.conn_errormsg(connection)
	cleanup()
	exit(-1)

while ibm_db.fetch_row(detail_rs) != False:
	row += 1

	for column in range(len(column_list)):
		value = ibm_db.result(detail_rs, str(column_list[column].strip()))
		# debugging
		#print "Row " + str(row) + " col " + str(column) + " column name " + column_list[column].strip() + " value " + str(value).strip()
		ws.write(row, column, value, normal)
		
# ================================================
# all done
cleanup()
print "Spreadsheet located at " + workbook_file.strip()
exit(0)