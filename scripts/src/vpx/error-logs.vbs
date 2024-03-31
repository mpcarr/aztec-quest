
'*****************************************************************************************************************************************
'  ERROR LOGS by baldgeek
'*****************************************************************************************************************************************

' Log File Usage:
'   WriteToLog "Label 1", "Message 1 "
'   WriteToLog "Label 2", "Message 2 "

Class DebugLogFile
	Private Filename
	Private TxtFileStream
	
	Private Function LZ(ByVal Number, ByVal Places)
		Dim Zeros
		Zeros = String(CInt(Places), "0")
		LZ = Right(Zeros & CStr(Number), Places)
	End Function
	
	Private Function GetTimeStamp
		Dim CurrTime, Elapsed, MilliSecs
		CurrTime = Now()
		Elapsed = Timer()
		MilliSecs = Int((Elapsed - Int(Elapsed)) * 1000)
		GetTimeStamp = _
		LZ(Year(CurrTime),   4) & "-" _
		 & LZ(Month(CurrTime),  2) & "-" _
		 & LZ(Day(CurrTime),	2) & " " _
		 & LZ(Hour(CurrTime),   2) & ":" _
		 & LZ(Minute(CurrTime), 2) & ":" _
		 & LZ(Second(CurrTime), 2) & ":" _
		 & LZ(MilliSecs, 4)
	End Function
	
	' *** Debug.Print the time with milliseconds, and a message of your choice
	Public Sub WriteToLog(label, message, code)
		Dim FormattedMsg, Timestamp
		'   Filename = UserDirectory + "\" + cGameName + "_debug_log.txt"
		Filename = cGameName + "_debug_log.txt"
		
		Set TxtFileStream = CreateObject("Scripting.FileSystemObject").OpenTextFile(Filename, code, True)
		Timestamp = GetTimeStamp
		FormattedMsg = GetTimeStamp + " : " + label + " : " + message
		TxtFileStream.WriteLine FormattedMsg
		TxtFileStream.Close
		Debug.print label & " : " & message
	End Sub
End Class

Sub WriteToLog(label, message)
	If KeepLogs Then
		Dim LogFileObj
		Set LogFileObj = New DebugLogFile
		LogFileObj.WriteToLog label, message, 8
	End If
End Sub

Sub NewLog()
	If KeepLogs Then
		Dim LogFileObj
		Set LogFileObj = New DebugLogFile
		LogFileObj.WriteToLog "NEW LOG", " ", 2
	End If
End Sub

'*****************************************************************************************************************************************
'  END ERROR LOGS by baldgeek
'*****************************************************************************************************************************************
