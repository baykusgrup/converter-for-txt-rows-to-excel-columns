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
        col = col + 1
    workbook.close()


def getRowData(filePath):
    # parsing date and time values
    baseName = os.path.basename(file)
    dateValue, hourValue = baseName.split(' ');
    hourValue = hourValue.replace(".txt", "");

    lines = [dateValue, hourValue]
    exactValue ='';
    # parsing file contents
    with open(filePath, 'r') as f:
        for line in f:
            print(line);
            t = line.find(':');
            print(t);
            #at db
            if t == 2:
                text = line.split(':')[0];
                exactValue = text
                if t == 2:
                    text = line;
                    exactValue = text
                # constration
                elif t == -1:
                    text = line.split(':')[0];
                    exactValue = text
            # constration
            elif t == -1:
                text = line.split(':')[0];
                exactValue = text
            #pm s
            elif t >= 0:
                umText, umValue = line.split(':')
                exactValue = umValue;

            lines.append(exactValue);
            exactValue='';

    f.close()
    print(lines);

    # return all rowData
    return lines;


allRows = [];
allRows.append(['Date', 'Hour', 'Cons.', '05um', '10um', '25um', '50um', '100um', 'atrh', 'dtwb']);
txt_files = glob.glob(globalFilePath);

for file in txt_files:
    rowData = getRowData(file)
    allRows.append(rowData);

createExcelFile(globalExcelFileName, allRows);
