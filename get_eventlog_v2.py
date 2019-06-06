from winevt import EventLog

SystemTime = "event.System.TimeCreated['SystemTime']"
PrintServer = "event.System.Computer"
Document = "event.UserData.Param2"
Username = "event.UserData.Param3"
Computer = "event.UserData.Param5"
Count = "event.UserData.Param8"

query = EventLog.Query("Microsoft-Windows-PrintService/Operational", "Event/System[EventID=307]")

for event in query:
    print(event.xml())

    # print(event.System.TimeCreated['SystemTime'],
    #       event.System.Computer)
          # event.UserData.DocumentPrinted.Param3[0],
          # event.UserData.DocumentPrinted.Param5,
          # event.UserData.DocumentPrinted.Param8)

