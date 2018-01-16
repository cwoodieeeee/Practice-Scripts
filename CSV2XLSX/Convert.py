import os
import glob
import csv
import datetime

from xlsxwriter.workbook import Workbook
now = datetime.datetime.now()
workbook = Workbook('{}.xlsx'.format(input('Enter new workbook name:')))        ## Creates the xlsx file with user input name

 for csvfile in glob.glob(os.path.join('.', '*.csv')):
    print('Creating new WS')
    decision = input('Default is to name worksheet {}{}{} Would you like to change WS name? (Y/N)'.format("'",(os.path.splitext(csvfile)[0][2::]),"'."))
    if decision == 'Y' or decision == 'y' or decision == 'yes' or decision == 'Yes' or decision == 'YES':
        worksheet = workbook.add_worksheet(input('Enter new worksheet name:'))
    elif decision == 'N' or decision == 'n' or decision == 'no' or decision == 'NO' or decision == 'No' or decision == '':
        worksheet = workbook.add_worksheet((os.path.splitext(csvfile)[0])[2::])
    else:
        print('Error. Entered value can only be Y or N.')
    print('Writing to new WS')
    with open(csvfile, 'r') as f:
        reader = csv.reader(f)
        for r, row in enumerate(reader):
            for c, col in enumerate(row):
                worksheet.write(r, c, col)
    print('Removing CSV file')
    os.remove(csvfile)
workbook.close()

print('Creating file')
print('Done')
