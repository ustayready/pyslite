'''
Script Name: pyslite_xlsx.py
Version: 1
Revised Date: 10/26/2015
Python Version: 3
Description: A script for performing SQLite database to Excel workbook conversions.
Copyright: 2015 Mike Felch <mike@linux.edu> 
URL: http://www.forensicpy.com/
--
- ChangeLog -
v1 - [10-26-2015]: Wrote original code
'''

import os, sys, sqlite3, xlsxwriter, argparse, ntpath

parser = argparse.ArgumentParser(
        description="pySlite - A SQLite data parsing utility")
parser.add_argument("--db", help="DB file to process")
args = parser.parse_args()

def main(args):
	print("* Reading {}...".format(args.db))
	db_to_excel(args.db)

def path_leaf(path):
	head, tail = ntpath.split(path)
	return tail or ntpath.basename(head)

def db_to_excel(db_file):
	with open(db_file,'rb') as fh:
		magic_number = bytes.fromhex('53514c69746520666f726d6174203300')
		file_magic = fh.read(16)

		print("\t- Checking magic number...")
		if not file_magic == magic_number:
			print('\t- Failed! This is not a valid SQLite database.')
		else:
			print("\t- Success! Verified valid sqlite database.")

			db = sqlite3.connect(db_file)

			tbls_cursor = db.cursor()
			tbls_sql = "SELECT * FROM sqlite_master where type = 'table';"
			tbls_cursor.execute(tbls_sql)

			cwd = os.path.dirname(os.path.abspath(__file__))
			full_path = '{}\output'.format(cwd)

			if not os.path.exists(full_path):
				os.makedirs(full_path)

			print("* Creating XLSX file...")

			db_filename = path_leaf(db_file)
			wb_name = '{}\{}.xlsx'.format(full_path, db_filename)
			workbook  = xlsxwriter.Workbook(wb_name)

			print("\t- Success! {}".format(wb_name))
			print("* Processing tables...")

			for x in enumerate(tbls_cursor):
				tbl_name = x[1][1]
				worksheet = workbook.add_worksheet(tbl_name)

				cur = db.cursor()
				sql = 'SELECT * FROM ' + tbl_name + ';'
				cur.execute(sql)

				rows = [x for x in enumerate(cur)]
				names = [description[0] for description in cur.description]

				col_idx = 0
				bold = workbook.add_format({'bold': True})
				for col_name in cur.description:
					worksheet.write(0, col_idx, col_name[0], bold)
					col_idx += 1

				print("\t- Table: {} | Rows: {} | Columns: {}".format(tbl_name, len(rows), col_idx))

				for x in range(len(rows)):
					row_data = rows[x][1]
					for y in range(len(row_data)):
						worksheet.write(x+1, y, row_data[y])

			workbook.close()

if __name__ == '__main__':
	if args.db:
		main(args)
	else:
		sys.exit("pySlite requires a db file.")
