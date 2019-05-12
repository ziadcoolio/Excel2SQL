from openpyxl import load_workbook
from pathlib import Path
import datetime
import sys

from openpyxl import Workbook

if len(sys.argv) < 3 :
	print("Not enough arguments")
	exit()
if ".xlsx" not in sys.argv[1] and ".xls" not in sys.argv[2] :
	print("Invalid output file format")
	exit()
if ".sql" not in sys.argv[2]:
	print("Invalid input file format")
	exit()
config = Path(sys.argv[1])
if not config.is_file():
	print("Cannot find input file")
	exit()
	
wb = load_workbook(sys.argv[1])
ws = wb.active

sql_file = open(sys.argv[2], "w+")

sql_file.write("INSERT ALL\n")
is_table = False
is_header = False
is_table_name = False

table_name = ""
table_header_list = []
table_header = ""
values_list = []

column_count = 0
for row in ws.iter_rows(min_row=1, max_col=None, max_row=ws.max_row):
    
    for cell in row:
        column_count += 1
        if cell.value == "~":
            is_table_name = True
            is_table = False
            table_header_list.clear()
            values_list.clear()
            break

        if not is_table:
            if cell.value == "TABLENAME" :
                is_table_name = False
                break
            if cell.value != None :
                table_name = cell.value

        else :
            if is_header:
                if cell.value == "TABLEHEADER" :
                    break
                table_header_list.append(cell.value)
            else :
                if cell.value == None :
                    values_list.append("NULL")
                elif isinstance(cell.value, str):
                    try:
                        value = int(cell.value)
                        values_list.append(cell.value)
                    except ValueError:
                        values_list.append("\'"+cell.value+"\'")
                elif isinstance(cell.value, datetime.datetime):
                    values_list.append("TO_DATE(\'"+str(cell.value)+"\',\'yyyy/mm/dd hh24:mi:ss\')")
                elif isinstance(cell.value, int) or isinstance(cell.value, float):
                    values_list.append(int(cell.value))

                if(column_count == len(table_header_list)):
                    break

    if not is_table_name :
        column_count = 0
        if not is_table :
            is_table = True
            is_header = True

        elif is_table and is_header:
            table_header = '('+','.join(table_header_list) + ')'
            is_header = False

        elif is_table and not is_header:
            sql_file.write("\tINTO "+table_name+" "+table_header+ " VALUES " + "("+", ".join(map(str, values_list))+")\n")
            values_list.clear()

sql_file.write("SELECT * FROM dual;")

sql_file.close()