from time import sleep
import xlwings as xw
import os
from datetime import timedelta, date


tray_tracker = './tray_tracker.xlsm'

years = {x: str(x) for x in range(2021, 2030+1)}
months = {0: 'January', 1: 'February', 2: 'March', 3: 'April', 4: 'May', 5: 'June',
          6: 'July', 7: 'August', 8: 'September', 9: 'October', 10: 'November', 11: 'December'}


def make_directories():
    for year in years:
        os.mkdir("./"+years[year])
        for month in months:
            os.mkdir("./"+years[year]+"/"+str(month+1)+"-"+months[month])


# Credit to Ber: https://stackoverflow.com/questions/1060279/iterating-through-a-range-of-dates-in-python
def date_range(start_date, end_date):
    for n in range(int((end_date - start_date).days)):
        yield start_date + timedelta(n)


def yield_dates(start, end):
    for single_date in date_range(start, end):
        # Returns the date in the form of '{Name of Day', 'Name of Month', 'Month', 'Day', 'Year'}
        yield single_date.strftime("%A %B %#m %#d %Y")


def main():
    xw.App(visible=False)
    excel_tray_tracker = xw.Book(tray_tracker)
    app = xw.apps.active
    start = date(2022, 1, 1)
    end = date(2031, 1, 1)
    for current_date in yield_dates(start, end):
        s = current_date.split()
        print(s)
        for i in range(20):
            current_sheet = excel_tray_tracker.sheets[i]
            current_sheet.range('D2').value = s[0]
            current_sheet.range('D3').value = "'" + s[1] + ", " + s[3] + ", " + s[4]
        new_directory = './' + s[-1] + '/' + s[2] + '-' + s[1] + '/'
        new_filename = s[2]+'-'+s[3]+'-'+s[4]+'.xlsm'
        excel_tray_tracker.save(new_directory + new_filename)
    sleep(3)
    app.quit()
    app.kill()


if __name__ == '__main__':
    main()

