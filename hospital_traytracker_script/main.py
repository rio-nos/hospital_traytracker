import xlwings as xw
import datetime as dt

tray_tracker = "./TrayTracker.xlsm"

def main():
    wb_test = xw.Book(tray_tracker)
    wb_name1 = wb_test.sheets[0]
    print(wb_name1.name)
    wb_name1.range('D2').value = 'Wednesday'
    wb_name1.range('D3').value = "'January, 31, 2021"
    wb_new = xw.Book(wb_test)
    wb_new.save('1-31-2021.xlsm')
    wb_test.save('1-31-2021.xlsm')

if __name__ == '__main__':
    main()
