import datetime

def get_dates_bytimes(StartDate, EndDate):
    date_list = []
    datestart = datetime.datetime.strptime(StartDate, '%Y-%m-%d')
    dateend = datetime.datetime.strptime(EndDate, '%Y-%m-%d')
    date_list.append(datestart.strftime('%Y-%m-%d'))
    while datestart < dateend:
        datestart += datetime.timedelta(days=1)
        date_list.append(datestart.strftime('%Y-%m-%d'))
    return date_list
