import sys
import win32evtlog
import win32event
import win32api
import win32con
import msvcrt


def main():
    server = None # "localhost" # name of the target computer to get event logs
    source_type = "System" # "Application" # "Security"
    h_log = win32evtlog.OpenEventLog(server, source_type)
    flags = win32evtlog.EVENTLOG_BACKWARDS_READ | win32evtlog.EVENTLOG_SEQUENTIAL_READ
    total = win32evtlog.GetNumberOfEventLogRecords(h_log)
    print(total)
    h_evt = win32event.CreateEvent(None, 1, 0, "evt0")
    win32evtlog.NotifyChangeEventLog(h_log, h_evt)
    print("Waiting for changes in the '{:s}' event log. Press a key to exit...".format(source_type))
    while not msvcrt.kbhit():
        wait_result = win32event.WaitForSingleObject(h_evt, 500)
        if wait_result == win32con.WAIT_OBJECT_0:
            print("The '{:s}' event log has been modified".format(source_type))
            # Any processing goes here
        elif wait_result == win32con.WAIT_ABANDONED:
            print("Abandoned")

    win32api.CloseHandle(h_evt)
    win32evtlog.CloseEventLog(h_log)


if __name__ == "__main__":
    print("Python {:s} on {:s}\n".format(sys.version, sys.platform))
    main()