import xlrd
import xlwt
def main():
  # save the file names of spreadsheets as variables so they may be accessed
  # the files must be in the same folder as this program or there will be problems
  cafo_data = 'CAFOtable.xls'
  playa_RSC = 'playa_RSCs.xlsx'

  # create a 'master' worksheet of the largest spreadsheet
  workbook = xlrd.open_workbook(cafo_data)
  master = workbook.sheet_by_name('Sheet1')

  # create variables to out write relevant data
  writeBook = xlwt.Workbook()
  writeSheet = writeBook.add_sheet('Sheet1', cell_overwrite_ok=True)

  # create variables to store data from smaller spreadsheets
  readbook = xlrd.open_workbook(playa_RSC)
  # access the first sheet (make sure the name of the sheet is 'Sheet1' or else program will not work
  smaller = readbook.sheet_by_name('Sheet1')


  # iterate through each row of the larger spreadsheet
  # we start at the second row because the first row is a title field
  for row in range(1, master.nrows):
    # store the unique number in a variable
    epaNum = master.cell_value(row, 3)
    # iterate through each row of the smaller spreadsheet
    for roW in range(1, smaller.nrows):
      # store unique numbers in variable
      epa = smaller.cell_value(roW, 1)
      # compare the unique numbers, if they are equivalent, out write to the writeSheet
      if epa == epaNum:
        writeSheet.row(row).write(0, playa_RSC)
        
  # out write the rest of the data from the master spreadsheet
  for row in range(0, master.nrows):
    for col in range(0, master.ncols):
      content = master.cell_value(row, col)
      writeSheet.row(row).write(col+1, content)
  # save the new spreadsheet
  writeBook.save('outputTEST.xls')
main()
