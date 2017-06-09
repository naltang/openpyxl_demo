""" excel python test """
from __future__ import print_function
import datetime
import openpyxl

def excel_io(filename_excel):
	""" excel test """
	workbook = openpyxl.load_workbook(filename_excel)
	worksheet = workbook.active

	list_numbers = []

	for column in worksheet.iter_cols(min_col=2, max_col=2):
		for cell in column:
			list_numbers.append(cell.value)

	sum_of_them = sum(list_numbers)
	mean_of_them = sum_of_them / len(list_numbers)

	worksheet['D1'] = "Timestamp"
	worksheet['E1'] = datetime.datetime.now()

	worksheet['D2'] = "Sum"
	worksheet['E2'] = sum_of_them

	worksheet['D3'] = "Mean"
	worksheet['E3'] = mean_of_them

	# Save the file
	workbook.save(filename_excel)

def main():
	""" main """
	excel_io("test.xlsx")

if __name__ == '__main__':
	main()
