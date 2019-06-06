import win32evtlog
import win32evtlogutil
import win32security
import win32con
import winerror
import time
import re
import string
import sys
import traceback


def date2sec(evt_date):
    """
    This function converts dates with format
    '12/23/99 15:54:09' to seconds since 1970.
    """
    regexp = re.compile('(.*)\\s(.*)')  # store result in site
    reg_result = regexp.search(evt_date)
    date = reg_result.group(1)
    the_time = reg_result.group(2)
    (mon, day, yr) = map(lambda x: string.atoi(x), string.split(date, '/'))
    (hr, min, sec) = map(lambda x: string.atoi(x), string.split(the_time, ':'))
    tup = [yr, mon, day, hr, min, sec, 0, 0, 0]

    sec = time.mktime(tup)

    return sec


# Main program
# initialize variables
flags = win32evtlog.EVENTLOG_BACKWARDS_READ | win32evtlog.EVENTLOG_SEQUENTIAL_READ

# This dict converts the event type into a human readable form
evt_dict = {win32con.EVENTLOG_AUDIT_FAILURE: 'EVENTLOG_AUDIT_FAILURE',
            win32con.EVENTLOG_AUDIT_SUCCESS: 'EVENTLOG_AUDIT_SUCCESS',
            win32con.EVENTLOG_INFORMATION_TYPE: 'EVENTLOG_INFORMATION_TYPE',
            win32con.EVENTLOG_WARNING_TYPE: 'EVENTLOG_WARNING_TYPE',
            win32con.EVENTLOG_ERROR_TYPE: 'EVENTLOG_ERROR_TYPE'}
computer = '192.168.152.4'
logtype = 'System'
begin_sec = time.time()
begin_time = time.strftime('%H:%M:%S  ', time.localtime(begin_sec))

# open event log
hand = win32evtlog.OpenEventLog(computer, logtype)
print(logtype, ' events found in the last 8 hours since:', begin_time)

try:
    events = 1
    while events:
        events = win32evtlog.ReadEventLog(hand, flags, 0)
        for ev_obj in events:
            # check if the event is recent enough
            # only want data from last 8hrs
            the_time = ev_obj.TimeGenerated.Format()
            seconds = date2sec(the_time)
            if seconds < begin_sec - 28800: break

            # data is recent enough, so print it out
            computer = str(ev_obj.ComputerName)
            cat = str(ev_obj.EventCategory)
            src = str(ev_obj.SourceName)
            record = str(ev_obj.RecordNumber)
            evt_id = str(winerror.HRESULT_CODE(ev_obj.EventID))
            evt_type = str(evt_dict[ev_obj.EventType])
            msg = str(win32evtlogutil.SafeFormatMessage(ev_obj, logtype))
            print(string.join((the_time, computer, src, cat, record, evt_id, evt_type, msg[0:15]), ':'))

    if seconds < begin_sec - 28800:
        pass

    win32evtlog.CloseEventLog(hand)
except:
    print(traceback.print_exc(sys.exc_info()))
