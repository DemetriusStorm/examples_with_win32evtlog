import win32evtlog
import win32con
import win32evtlogutil

import traceback
import time
import sys
import re
import string
import datetime
import locale

server = '192.168.152.4'

logtype = 'Microsoft-Windows-PrintService/Operational'
hand = win32evtlog.OpenEventLog(server, logtype)
total = win32evtlog.GetNumberOfEventLogRecords(hand)
flags = win32evtlog.EVENTLOG_BACKWARDS_READ | win32evtlog.EVENTLOG_SEQUENTIAL_READ
events = win32evtlog.ReadEventLog(hand, flags, 0)

for ev_obj in events:
    the_time = ev_obj.TimeGenerated.Format()  # '12/23/99 15:54:09'
    evt_id = str(winerror.HRESULT_CODE(ev_obj.EventID))
    computer = str(ev_obj.ComputerName)
    cat = ev_obj.EventCategory
    seconds = date2sec(the_time)
    record = ev_obj.RecordNumber
    msg = str(win32evtlogutil.SafeFormatMessage(ev_obj, logtype))
    source = str(ev_obj.SourceName)


def date2sec(self, evt_date):
    """
    converts '12/23/99 15:54:09' to seconds
    print '333333',evt_date
    """

    regexp = re.compile('(.*)\\s(.*)')
    reg_result = regexp.search(evt_date)
    date = reg_result.group(1)
    the_time = reg_result.group(2)
    (mon, day, yr) = map(lambda x: string.atoi(x), string.split(date, '/'))
    (hr, min, sec) = map(lambda x: string.atoi(x), string.split(the_time, ':'))
    tup = [yr, mon, day, hr, min, sec, 0, 0, 0]
    sec = time.mktime(tup)
    return sec
