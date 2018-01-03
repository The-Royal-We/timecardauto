import properties
import winsound, ctypes, shutil, os
from xlwt import easyxf
from xlutils.copy import copy
from xlrd import open_workbook
from datetime import datetime, timedelta, date

def main():
	week_dates = get_week_dates()
	generate_timecard(week_dates)
	week_date_ending = week_dates[-1]
	alertUser(week_date_ending)

	
def generate_timecard(week_dates):
	template_location = os.path.dirname(os.path.realpath(__file__)) + "/timecard_template.xls"
	timecard_folder = os.path.expanduser("~") + '/Timecards/'
	timecard_destination =  timecard_folder + 'BCarrollTimecard_%s' % week_dates[0].replace('/','') + '.xls'

	if not os.path.exists(timecard_folder):
		os.makedirs(timecard_folder)

	rb = open_workbook(template_location)
	r_sheet = rb.sheet_by_index(0)
	wb = copy(rb)
	w_sheet = wb.get_sheet(0)

	start_row = 4
	col_date = 0
	col_project = 1
	col_hours = 2

	for i, row_index in enumerate(range(start_row, start_row + len(week_dates)), 0):
		w_sheet.write(row_index, col_date, week_dates[i])
		w_sheet.write(row_index, col_project, "")
		w_sheet.write(row_index, col_hours, "8")
	wb.save(timecard_destination)

def get_week_dates():
	return [(datetime.now() - timedelta(i)).strftime('%Y/%m/%d') for i in range(5)][::-1]

def alertUser(week_date_ending):
	winsound.MessageBeep()
	msg = "Timecard has been generated for week ending %s, review at your leisure my dude." % week_date_ending
	ctypes.windll.user32.MessageBoxW(0, msg, "ALERT", 0)

if __name__ == '__main__':
	main()