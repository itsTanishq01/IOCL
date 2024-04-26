from headers import *

def main():
    # Import data into MySQL
    ExcelPath = GetExcel()

    if ExcelPath:
        ImportSQL(ExcelPath)
        print("Data imported to MySQL.")
    else:
        print("No file selected for import. Exiting.")
        return

    # Choose which report to export using the UI
    def ReportSelect(ReportNumber):
        if ReportNumber == 1:
            export1()
        elif ReportNumber == 2:
            export2()
        elif ReportNumber == 3:
            export3()
        elif ReportNumber == 4:  
            export4()
        elif ReportNumber == 5:  
            export5()
        elif ReportNumber == 6:  
            export6()
        elif ReportNumber == 7:  
            export7()    
        else:
            print("Invalid report number. Exiting.")

    ChooseReport(ReportSelect)
    
    DropTable()

if __name__ == "__main__":
    main()
