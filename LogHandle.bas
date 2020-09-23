Attribute VB_Name = "LogHandle"
' *****************************************************
' *****************************************************
' ********    Log Handle module  (Fade)
' ********
' ******** I Originaly created this module for the TVM
' ******** broadcasting system, and stripped it down to
' ******** allow me to implement it to other smaller
' ******** projects which needed some logging capabilities
' ******** It's a bit 'not-so-good' in this stripped down
' ******** form, but it does the trick
' ********
' ******** To use it, user must at least set 'LogPath'
' ******** or current directory will be used, to make
' ******** your project to make a new log file each time
' ******** it runs, use the 'AdvanceLogNum' method at
' ******** project load.


Public Path As String
Public logPath As String
Public DuringSave As Boolean
Public LogNum As Long
Public logInternalSelect As Boolean


    ' Advancing to next log file
Public Sub AdvanceLogNum()
    LogNum = Val(GetSetting("ResourceMonitor", "Local", "LogNum", "0"))
    SaveSetting "ResourceMonitor", "Local", "LogNum", Mid(Str(LogNum + 1), 2)
End Sub

    ' Saves the log to file
    ' if 'Jump' then method will auto set to next log file
    ' Normaly 'Path' is used as "" (Empty String), and the
    ' method will use 'LogPath' as the path, and file name will
    ' be compiled using 'LogNum', otherwise 'Path' must be a full
    ' path & file name (exp: "C:\MyLog.log")
Public Sub SaveLog(Path As String, Jump As Boolean)
    On Error GoTo errHandle
    
    DuringSave = True  'Indicating that the log file is being in use
    If SystemInfo.Log.ListCount > 32000 Then
        Jump = True
        SystemInfo.Log.Clear
        UpdateLog "New log opened.", 0
    End If
    
    If Path = "UpdateSave" Then us = True
    
    If Path = "" Or Path = "UpdateSave" Then ' if Path was retrieved succesfully
            'file = InfPath & LogFilesInf    'Then open log info file
    
            LogNum = Val(GetSetting("ResourceMonitor", "Local", "LogNum", "0"))
            If Jump Then                'If Jump to next file do so
                SaveSetting "ResourceMonitor", "Local", "LogNum", Mid(Str(LogNum + 1), 2)
            End If
        ' inf files path & Log Path & File name    (Systax for below)
            File = logPath & Right(Str(LogNum), Len(Str(LogNum) - 1)) & ".Log"
    Else
            File = Path
    End If
    
    Open File For Output As #10
        If Not (us) Then Call UpdateLog("Saved", 1)
        For i = 0 To SystemInfo.Log.ListCount
            Print #10, SystemInfo.LogTime.List(i), SystemInfo.Log.List(i)
        Next i
        Print #10, ""
        Print #10, Date & " " & Time
    Close #10
        
    DuringSave = False 'Indicating that the file is accessable
    
    Exit Sub
errHandle:
    MsgBox "Cannot Save Log - " & Err.Description
    Exit Sub
End Sub


    ' Adds a new Entry to the log
    ' NOTE: Each entry is auto marked with a time stamp
    ' Method Auto-Saves log to file
    ' 'New-Entry' indicates if to create a new entry or to
    ' append to last entry,
    ' if 0 (zero) the specified entry will be appended to last entry
    ' if 1 (one) a new entry will be created
    ' if >1 a 'Spacing' entry will be added for every incementation.
    ' if <0 no Time-Stamp is set at current method execution
Public Sub UpdateLog(ByVal msg As String, NewEntry As Integer)
    AddTime = True
    If NewEntry < 0 Then
        AddTime = False ' Setting to not add time
        NewEntry = NewEntry * -1
    End If
    If NewEntry = 0 Then ' If only Appending
        i = SystemInfo.Log.ListCount - 1 ' Getting last entry index
            ' Appending to  Entry
        SystemInfo.Log.List(i) = SystemInfo.Log.List(i) & "  " & msg
            ' Appending to  Time-Stamp
        SystemInfo.LogTime.List(i) = SystemInfo.LogTime.List(i) & " To " & Time
    Else                ' If creating new entry
            ' Loop to new entry relative index
        For i = 1 To NewEntry
            If i = NewEntry Then ' If relative index reached
                SystemInfo.Log.AddItem (msg) ' Add entry
                If AddTime Then timeStr = Time ' If add time then set value
                SystemInfo.LogTime.AddItem (timeStr) ' Set Time Stamp
            Else                ' If still 'Spacing"
                SystemInfo.Log.AddItem "" ' Make an Empty entry
                SystemInfo.Log.ItemData(SystemInfo.Log.ListCount - 1) = SystemInfo.Log.ListCount
                SystemInfo.LogTime.AddItem "" ' Make an Empty Time-Stamp
            End If
        Next i
    End If
    logInternalSelect = True ' Set Code-Event flag, so log doesn't think user clicked
    SystemInfo.Log.ListIndex = SystemInfo.Log.ListCount - 1 ' Select new entry (scrolling to entry)
    logInternalSelect = False ' Lower Flag
        ' If not during log save
    If Not (DuringSave) Then SaveLog "UpdateSave", False ' Auto-Save
    
End Sub

    ' Mainly the same as 'Updatelog' only this method will ONLY APPEND
    ' the specified entry to an existing entry indexed 'Index'
Public Sub UpdateLogWithIndex(msg As String, Index)
        ' Validating
    If Index > SystemInfo.Log.ListCount - 1 Then Exit Sub
        ' Setting Entry
    SystemInfo.Log.List(Index) = SystemInfo.Log.List(Index) & "  " & msg
        ' Refreshing Time Stamp
    SystemInfo.LogTime.List(Index) = SystemInfo.LogTime.List(Index) & " To " & Time
End Sub
    
    
    ' This method will Change to a new log file,
    ' Numbered 'NewLog', and will keep current log entries is 'KeepCurrent'
Public Function ChangeLogFile(NewLog As Long, keepCurrent As Boolean) As Boolean
    Dim OldLogNum As Long
    OldLogNum = LogNum
            ' Declare change in current log
    UpdateLog "Changing log file to: " & NewLog & ".Inf", 2
        ' Re-Creating path
    Path = logPath
        ' Incrementing log number
    SaveSetting "ResourceMonitor", "Local", "LogNum", Mid(Str(NewLog), 2)
        ' If not keeping entries
    If Not (keepCurrent) Then
            ' Reset log
        SystemInfo.Log.Clear
            ' Create openning entries
        UpdateLog Date & " " & Time, 1
        UpdateLog "Continuing from log number " & OldLogNum, 2
        ver = App.Major & "." & App.Minor & "." & App.Revision
        UpdateLog "ResourceMonitor ", 1
    Else    ' If keeping entries
                ' Declare change in new log
        UpdateLog "Continuing from log number " & LogNum, 2
    End If
        ' Save in new file
    SaveLog "UpdateSave", False
End Function

    ' Method to be used in an Error-Handle code, to log the error
Public Sub LogError(Optional ByVal ErrDesc As String)
    If ErrDesc = "" Then ErrDesc = Err.Description
        ' log the error
    Call UpdateLog("Error Occured:" & " Error #" & Err.Number & " :" & ErrDesc, 1)
        ' and show it ( with beep )
    LogEntryViewer.updateEntryViewer Time, "The following error occured:" & Chr(13) & "Error #" & _
        Err.Number & " :" & Chr(13) & ErrDesc, True
        
    'Resume Next
End Sub
