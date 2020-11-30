' PAGEFILEMON - WMI Based VBS Script to monitor VIRT_MEM on WIN32 and WIN64 Systems

' If logging is enabled (default), the results are saved in My Documents folder
' as PageFileLog.txt.

'**********************************************************
' Three optional settings are configurable below:
'  WriteToFile - If set to True the information will be added to a log file in
'    your 'My Documents' folder.  Of course, you want this if you are running
'    at logoff, but you might not want it for manually checks.  Changing this
'    to False disables logging.
'   ShowPopup - If set to True then after the script runs a message box is
'     presented with the results.  This might not be desirable when
'     automatically running the script at logoff.  False disables popup.
'   DisplaySeconds - The number of seconds that the results popup will
'     display.  Setting this to 0 (zero) will cause the popup to remain until
'     acknowledged.

WriteToFile = True    'Options: True, False
ShowPopup = True      'Options: True, False
DisplaySeconds = 10    '0 (zero) to force OK

' NOTE: If ANY arguments are used, ALL hardcoded variables are set to
' false or 0, so you must specifically set which options you want.

' To use these options, create a shortcut to the script and add the arguments
' there, or the arguments can be used running the script from command line.
'**********************************************************

' Do not edit below this line
If WScript.Arguments.Count > 0 Then
  WriteToFile = False
  ShowPopup = False
  DisplaySeconds = 0
  For Each arg in WScript.Arguments
    If LCase(arg) = "log" Then
      WriteToFile = True
    End If
    If LCase(arg) = "rpt" Then
      ShowPopup = True
    End If
    If Left(LCase(arg), 2) = "t:" Then
      If IsNumeric(Mid(arg, 3)) Then
        DisplaySeconds = Mid(arg, 3)
      End If
    End If
  Next
End If

For Each obj in GetObject("winmgmts:\\.\root\cimv2").ExecQuery(_
    "Select Name, CurrentUsage, PeakUsage, " & _
    "AllocatedBaseSize from Win32_PageFileUsage",,48)
  s = s & "Pagefile Physical Location: " & vbtab & obj.Name & vbcrlf
  s = s & "Current Pagefile Usage: " & vbtab & obj.CurrentUsage & " MB" & vbcrlf
  s = s & "Session Peak Usage: " & vbtab & vbtab & obj.PeakUsage & " MB" & vbcrlf
  s = s & "Current Pagefile Size: " & vbtab & obj.AllocatedBaseSize & " MB" & vbcrlf
Next

If WriteToFile Then
  Set fso = CreateObject("Scripting.FileSystemObject")
  logfile = CreateObject("WScript.Shell"). _
    SpecialFolders("MyDocuments") & "\PagefileLog.txt"
  If NOT fso.OpenTextFile(logfile, 1, True).AtEndOfStream Then
    With fso.OpenTextFile(logfile, 1)
      s2 = .ReadAll : .Close
    End With
  End If
  With fso.OpenTextFile(logfile, 2)
    .Write Now() & vbcrlf & s & vbcrlf & s2 : .Close
  End With
End If

If ShowPopup Then
  WScript.CreateObject("WScript.Shell").Popup _
    s, DisplaySeconds, "PageFileMon by The Frickster ver. 1.01", 4096
End If
