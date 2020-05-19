import glob
import os
import xlsxwriter

globalExcelFileName = 'Deneme.xlsx';
globalFilePath = "/Users/kali/Desktop/akin/pm_ornek/*.txt";


def createExcelFile( bookName , data ):   
    # Create a workbook and add a worksheet.
    workbook = xlsxwriter.Workbook(bookName)
    worksheet = workbook.add_worksheet()

    # Start from the first cell. Rows and columns are zero indexed.
    row = 0
    col = 0

    # Iterate over the data and write it out row by row.
    # for col, data in enumerate(data):
    #     worksheet.write_column(row, col, data)

    for date, time, con, zfive, one, twoFive, five, ten, atrh, dpwh  in data:
            worksheet.write(row, col,  date)
            worksheet.write(row, col+1,  time)
            worksheet.write(row, col+2,  con)
            worksheet.write(row, col+3,  zfive)
            worksheet.write(row, col+4,  one)
            worksheet.write(row, col+5,  twoFive)
            worksheet.write(row, col+6,  five)
            worksheet.write(row, col+7,  ten)
            worksheet.write(row, col+8,  atrh)
            worksheet.write(row, col+9,  dpwh)
            row += 1

    workbook.close()



def getRowData( filePath ):
    # parsing date and time values
    baseName = os.path.basename(file)
    dateValue, hourValue = baseName.split(' ');
    hourValue = hourValue.replace(".txt", "");

    # parsing file contents
    with open(filePath, 'r') as f:
        lines = [line.rstrip() for line in f]

    #return all rowData
    return [dateValue, hourValue, lines[0],lines[1],lines[2],lines[3],lines[4],lines[5],lines[6],lines[7]];



allRows=[];

txt_files = glob.glob( globalFilePath );

for file in txt_files:
    rowData= getRowData(file)
    allRows.append(rowData);

createExcelFile( globalExcelFileName , allRows);






