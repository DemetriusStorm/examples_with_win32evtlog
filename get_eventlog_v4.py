import pywintypes
import win32evtlog


INFINITE = 0xFFFFFFFF
EVTLOG_READ_BUF_LEN_MAX = 0x7FFFF


def get_record_data(eventlog_record):
    ret = dict()
    for key in dir(eventlog_record):
        if 'A' < key[0] < 'Z':
            ret[key] = getattr(eventlog_record, key)
    return ret


def get_eventlogs(source_name="Microsoft-Windows-PrintService/Operational", buf_size=EVTLOG_READ_BUF_LEN_MAX,
                  backwards=True):
    ret = list()
    evt_log = win32evtlog.OpenEventLog(None, source_name)
    read_flags = win32evtlog.EVENTLOG_SEQUENTIAL_READ
    if backwards:
        read_flags |= win32evtlog.EVENTLOG_BACKWARDS_READ
    else:
        read_flags |= win32evtlog.EVENTLOG_FORWARDS_READ
    offset = 0
    eventlog_records = win32evtlog.ReadEventLog(evt_log, read_flags, offset, buf_size)
    while eventlog_records:
        ret.extend(eventlog_records)
        offset += len(eventlog_records)
        eventlog_records = win32evtlog.ReadEventLog(evt_log, read_flags, offset, buf_size)
    win32evtlog.CloseEventLog(evt_log)
    return ret


def get_events_xmls(channel_name="Microsoft-Windows-PrintService/Operational", events_batch_num=100, backwards=True):
    ret = list()
    flags = win32evtlog.EvtQueryChannelPath
    if backwards:
        flags |= win32evtlog.EvtQueryReverseDirection
    try:
        query_results = win32evtlog.EvtQuery(channel_name, flags, None, None)
    except pywintypes.error as e:
        print(e)
        return ret
    events = win32evtlog.EvtNext(query_results, events_batch_num, INFINITE, 0)
    while events:
        for event in events:
            ret.append(win32evtlog.EvtRender(event, win32evtlog.EvtRenderEventXml))
        events = win32evtlog.EvtNext(query_results, events_batch_num, INFINITE, 0)
    return ret


def main():
    import sys, os
    from collections import OrderedDict
    # standard_log_names = ["System", "Security", "Microsoft-Windows-PrintService/Operational"]
    standard_log_names = ["Microsoft-Windows-PrintService/Operational"]
    source_channel_dict = OrderedDict()

    for item in standard_log_names:
        source_channel_dict[item] = item

    # for item in ["Windows Powershell"]:  # !!! This works on my machine (96 events)
    #     source_channel_dict[item] = item

    for source, channel in source_channel_dict.items():
        print(source, channel)
        logs = get_eventlogs(source_name=source)
        xmls = get_events_xmls(channel_name=channel)

        # print("\n", get_record_data(logs[0]))
        print(xmls[0])
        # print("\n", get_record_data(logs[-1]))
        # print(xmls[-1])

        # print(len(logs))
        # print(len(xmls))


if __name__ == "__main__":
    main()
