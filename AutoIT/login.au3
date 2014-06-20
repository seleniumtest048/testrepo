
WinWaitActive("Authentication Required","","20")
 If WinExists("Authentication Required") Then
 Send("Poll Share{TAB}")
Send("{ESCAPE}")
 EndIf