from datetime import timedelta, date
from time import sleep
import xlwings as xw
import sys
import os


tray_tracker = './TrayTracker.xlsm'

years = {x: str(x) for x in range(2021, 2021+1)}
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
    print(f'Opening {tray_tracker}')
    # make_directories()
    xw.App(visible=False)
    excel_tray_tracker = xw.Book(tray_tracker)
    app = xw.apps.active
    start = date(2021, 4, 9)
    end = date(2021, 5, 1)
    for current_date in yield_dates(start, end):
        s = current_date.split()
        print(s)
        for i in range(20):
            current_sheet = excel_tray_tracker.sheets[i]
            current_sheet.range('E8').value = 'Initials'
            current_sheet.range('D2').value = s[0]
            current_sheet.range('D3').value = "'" + s[1] + ", " + s[3] + ", " + s[4]
        new_directory = './' + s[-1] + '/' + s[2] + '-' + s[1] + '/'
        new_filename = s[2]+'-'+s[3]+'-'+s[4]+'.xlsm'
        excel_tray_tracker.save(new_directory + new_filename)
        # excel_tray_tracker.save('NewTrayTracker.xlsm')
    sleep(1.5)
    app.quit()
    app.kill()


if __name__ == '__main__':
    if len(sys.argv) == 1:
        print("Usage [python3 main.py mm/dd/yyyy mm/dd/yyyy (-d)]: no date range given")
    elif len(sys.argv) == 4:
        print("Creating excel files of dates {} to {}".format(sys.argv[1], sys.argv[2]))

