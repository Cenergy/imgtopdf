import openpyxl
from openpyxl.styles import Font, Alignment


def main():
    sSourceFile = "ExcelMod.xlsx"
    sTargetFile = "Target.xlsx"
    wb = openpyxl.load_workbook(sSourceFile)
    wb2 = openpyxl.load_workbook(sTargetFile)

    copy_sheet1 = wb2.copy_worksheet(wb.worksheets[0])

    wb.save('new_'+sTargetFile)

    print("It is over")


if __name__ == "__main__":
    main()
