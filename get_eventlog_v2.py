from winevt import EventLog


query = EventLog.Query("Microsoft-Windows-PrintService/Operational", "Event/System[EventID=307]", handle_event)

for event in query:
    print("Got event: " + str(event))

