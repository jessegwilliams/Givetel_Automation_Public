from utils import *

DAYS_OF_WEEK = ["MONDAY", "TUESDAY", "WEDNESDAY", "THURSDAY", "FRIDAY", "SATURDAY", "SUNDAY"]

### DATETIME FUNCTIONS
# Return a date for the first of a given month
def first_of(month=None):
    if month:
        if type(month) == float: month = int(month)
        return START_OF_MONTH.replace(month=month)
    else:
        return START_OF_MONTH

# Return a date for the end of a given month
def end_of(month=None):
    next_month = TODAY.replace(day=28) + dt.timedelta(days=4)
    if month:  ## not working
        next_month = next_month - dt.timedelta(days=next_month.day)
        return next_month.replace(month=month)
    else:
        return next_month - dt.timedelta(days=next_month.day)

def date_to_gt_format(date):
    if not isinstance(date, dt.date):
        print('Date must be datetime.date object.')
        return None
    else:
        return date.strftime('%d%m%y')

def date_to_roster_format(date):
    year, month, day = str(date.year), str(date.month), str(date.day)
    if len(day) < 2: day = "0" + day
    if len(month) < 2: month = "0" + month
    if len(year) > 2: year = year[2:]
    return f"WC {day}.{month}.{year}"

def is_day_name(string):
    if string.upper() in DAYS_OF_WEEK: return True
    else: return False

def day_to_date(day_str, week_offset=0):
    ### Convert day of current week (by default) to a datetime object. Use week offset to target past or future weeks.
    today_index = 0
    input_index = 0
    for i, day in enumerate(DAYS_OF_WEEK):
        if TODAY.strftime("%A").upper() == day:
            today_index = i
        if day_str.upper() == day:
            input_index = i
    return TODAY + dt.timedelta(days=int((input_index-today_index)+(week_offset*7)))