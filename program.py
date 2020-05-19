import glob
import os
import xlsxwriter

globalExcelFileName = 'test.xlsx';
globalFilePath = "datas/*.txt";

def createExcelFile(bookName, data):
    # Create a workbook and add a worksheet.
    workbook = xlsxwriter.Workbook(bookName)
    worksheet = workbook.add_worksheet()

    # Start from the first cell. Rows and columns are zero indexed.
    row = 0
    col = 0

    # Iterate over the data and write it out row by row.
    # for col, data in enumerate(data):
    #     worksheet.write_column(row, col, data)

    for row_num, item in enumerate(data):
        for row_num2, item2 in enumerate(item):
            worksheet.write(col, row_num2, item2)
        col = col +1
    workbook.close()


def getRowData(filePath):
    # parsing date and time values
    baseName = os.path.basename(file)
    dateValue, hourValue = baseName.split(' ');
    hourValue = hourValue.replace(".txt", "");

    lines=[dateValue, hourValue]
    # parsing file contents
    with open(filePath, 'r') as f:
        for line in f:
            # in python 2
            # print line
            # in python 3
            lines.append(line);

        f.close()

    # return all rowData
    return lines;


allRows = [];

txt_files = glob.glob(globalFilePath);

for file in txt_files:
    rowData = getRowData(file)
    allRows.append(rowData);

createExcelFile(globalExcelFileName, allRows);