import pandas as pd
import helperFunctions as utility
import sys
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

def generate_difference_sheet(workbook, worksheet1, worksheet2, reshapedRows, reshapedColumns):
    differenceSheet = workbook.create_sheet(title='Differences')
    differenceSheet.append(['Column Name', 'Cell', 'FirstFile Value', 'SecondFile Value'])

    for rows in range(2, reshapedRows + 2):
        for columns in range(1, reshapedColumns + 1):
            worksheetOneValue = worksheet1.cell(row=rows, column=columns).value
            worksheetTwoValue = worksheet2.cell(row=rows, column=columns).value
            if worksheetOneValue != worksheetTwoValue:
                columnHeader = worksheet1.cell(row=1, column=columns).value if worksheet1.cell(row=1, column=columns).value else f"Col_{columns}"
                cell_name = f"{chr(64 + columns)}{rows}" if columns <= 26 else f"{chr(64 + (columns - 1) // 26)}{chr(64 + columns % 26)}{rows}"
                differenceSheet.append([columnHeader, cell_name, worksheetOneValue, worksheetTwoValue])

    return workbook

def compare_excel_files(firstFile, secondFile, outputFile):
    firstFile = input('Please enter the name of the first file:   ') + '.xlsx'
    secondFile = input('Please enter the name of the second file:  ') + '.xlsx'
    
    try:
        dataFrameFirstFile = pd.read_excel(firstFile, sheet_name=0, engine='openpyxl')
        dataFrameSecondFile = pd.read_excel(secondFile, sheet_name=0, engine='openpyxl')
    except FileNotFoundError:
        print("Error! One of the input file is could not be found. Ensure it is spelled correctly and that you have the file extension!")
        sys.exit(1)

    print(f"Comparing {firstFile} and {secondFile}...")
    reshapedRows = max(dataFrameFirstFile.shape[0], dataFrameSecondFile.shape[0])
    reshapedColumns = max(dataFrameFirstFile.shape[1], dataFrameSecondFile.shape[1])
    dataFrameFirstFile = dataFrameFirstFile.reindex(index=range(reshapedRows), columns=dataFrameFirstFile.columns.tolist() + [None]*(reshapedColumns - dataFrameFirstFile.shape[1]))
    dataFrameSecondFile = dataFrameSecondFile.reindex(index=range(reshapedRows), columns=dataFrameSecondFile.columns.tolist() + [None]*(reshapedColumns - dataFrameSecondFile.shape[1]))

    with pd.ExcelWriter(outputFile, engine='openpyxl') as writer:
        dataFrameFirstFile.to_excel(writer, sheet_name='FirstFile', index=False)
        dataFrameSecondFile.to_excel(writer, sheet_name='SecondFile', index=False)

    workbook = load_workbook(outputFile)
    worksheetOne = workbook['FirstFile']
    worksheetTwo = workbook['SecondFile']
    for row in range(2, reshapedRows + 2):
        for col in range(1, reshapedColumns + 1):
            if worksheetOne.cell(row=row, column=col).value != worksheetTwo.cell(row=row, column=col).value:
                worksheetTwo.cell(row=row, column=col).fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')
    
    generate_difference_sheet(workbook, worksheetOne, worksheetTwo, reshapedRows, reshapedColumns)
    worksheet = workbook.active
    for worksheet in workbook.worksheets:
        utility.auto_adjust_columns(worksheet)

    workbook.save(outputFile)
    print(f"Comparison complete! Differences highlighted and saved in: {outputFile}")


if __name__ == "__main__":
    compare_excel_files('File1.xlsx', 'File2.xlsx', 'Compared_Output.xlsx')
