VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form SystemInfo 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "System Resources Monitor"
   ClientHeight    =   5475
   ClientLeft      =   45
   ClientTop       =   4785
   ClientWidth     =   9030
   Icon            =   "SystemInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5475
   ScaleWidth      =   9030
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer LogTimer 
      Enabled         =   0   'False
      Interval        =   3500
      Left            =   4080
      Top             =   5040
   End
   Begin VB.Frame Frame2 
      Caption         =   "Log"
      Height          =   4935
      Left            =   5040
      TabIndex        =   13
      Top             =   120
      Width           =   3855
      Begin VB.CommandButton Reset_but 
         Caption         =   "Reset"
         Height          =   255
         Left            =   2880
         TabIndex        =   24
         Top             =   480
         Width           =   855
      End
      Begin VB.ListBox LogTime 
         Height          =   2595
         Left            =   120
         TabIndex        =   16
         Top             =   1200
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.CheckBox Logging_c 
         Caption         =   "Enable Logging"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   2295
      End
      Begin VB.ListBox Log 
         Height          =   3765
         Left            =   120
         TabIndex        =   14
         Top             =   720
         Width           =   3615
      End
      Begin VB.Label Path_l 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   960
         TabIndex        =   21
         Top             =   4560
         Width           =   2775
      End
      Begin VB.Label Label10 
         Caption         =   "Saving to"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   4560
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Monitor"
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4815
      Begin MSComctlLib.ProgressBar ProgressBar3 
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   4440
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin MSComctlLib.ProgressBar ProgressBar2 
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   2880
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1320
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label VirtualMem 
         Caption         =   "Label10"
         Height          =   255
         Left            =   1920
         TabIndex        =   19
         Top             =   4080
         Width           =   2655
      End
      Begin VB.Label PageFile 
         Height          =   255
         Left            =   1680
         TabIndex        =   18
         Top             =   2520
         Width           =   2895
      End
      Begin VB.Label PhysicalMem 
         Height          =   255
         Left            =   2040
         TabIndex        =   17
         Top             =   960
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "Total Physical Memory: "
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   4575
      End
      Begin VB.Label Label2 
         Caption         =   "Free Physical Memory (RAM): "
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   4575
      End
      Begin VB.Label Label4 
         Caption         =   "Total Size of Current Paging File (In KB): "
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1800
         Width           =   4575
      End
      Begin VB.Label Label5 
         Caption         =   "Free Memory in Current Paging File (In KB): "
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   2160
         Width           =   4575
      End
      Begin VB.Label Label7 
         Caption         =   "Total Virtual Memory: "
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   3360
         Width           =   4335
      End
      Begin VB.Label Label8 
         Caption         =   "Free Virtual Memory: "
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   3720
         Width           =   4335
      End
      Begin VB.Label Label3 
         Caption         =   "Physical Memory usage: "
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label6 
         Caption         =   "Paging File Usage:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   2520
         Width           =   1455
      End
      Begin VB.Label Label9 
         Caption         =   "Virtual Memory Usage: "
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   4080
         Width           =   1695
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   4800
      Top             =   4920
   End
   Begin VB.Label Version_l 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Coded by Fade."
      Height          =   255
      Left            =   5160
      TabIndex        =   23
      ToolTipText     =   "Coded by Fade (Amit Ben-Shahar)"
      Top             =   5160
      Width           =   3375
   End
   Begin VB.Label Coder_l 
      BackColor       =   &H00FFFFFF&
      Caption         =   "     Coded by Fade."
      Height          =   255
      Left            =   120
      TabIndex        =   22
      ToolTipText     =   "Coded by Fade (Amit Ben-Shahar)"
      Top             =   5160
      Width           =   8775
   End
End
Attribute VB_Name = "SystemInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ********
' ******** System Resource monitor (Fade)
' ********  Very simple, very easy to understand.
' ********

Dim meminfo As MemoryStatus
Dim PhysUsed
Dim PageUsed
Dim VirtUsed

    ' For loggin
Dim EntryMod As Byte

Private Sub Form_Load()
    meminfo.dwLength = Len(meminfo)
        ' Getting first status
    Call GlobalMemoryStatus(meminfo)
    ResetLog ' Resetting log system
    Reset_info ' Resetting progress bars
    Update_Info ' Loading information to monitor frame
End Sub

Private Sub Reset_info()
    On Error GoTo errHandle

    meminfo.dwLength = Len(meminfo)
    Call GlobalMemoryStatus(meminfo)
    ProgressBar1.Min = 0
    ProgressBar2.Min = 0
    ProgressBar3.Min = 0
    ProgressBar1.Max = meminfo.dwTotalPhys
    ProgressBar2.Max = meminfo.dwTotalPageFile
    ProgressBar3.Max = meminfo.dwTotalVirtual
    
    Exit Sub ' Normal Sub Termination
errHandle:
        ' Logging error
    LogError "UpdateInfo"
End Sub

Private Sub Update_Info()
    On Error GoTo errHandle
    
    meminfo.dwLength = Len(meminfo)
    Call GlobalMemoryStatus(meminfo)
    PhysUsed = meminfo.dwTotalPhys - meminfo.dwAvailPhys
    PageUsed = meminfo.dwTotalPageFile - meminfo.dwAvailPageFile
    VirtUsed = meminfo.dwTotalVirtual - meminfo.dwAvailVirtual
    ProgressBar1.Value = PhysUsed
    ProgressBar2.Value = PageUsed
    ProgressBar3.Value = VirtUsed
    Label1.Caption = "Total Physical Memory: " & meminfo.dwTotalPhys / 1024 & " KB"
    Label2.Caption = "Free Physical Memory (RAM): " & meminfo.dwAvailPhys / 1024 & " KB"
    PhysicalMem.Caption = Format(PhysUsed / meminfo.dwTotalPhys, "0.00%") '  "Physical Memory Usage: " &
    Label4.Caption = "Total Size of Current Paging File (In KB): " & meminfo.dwTotalPageFile / 1024 & " KB"
    Label5.Caption = "Free Memory in Current Paging File (In KB): " & meminfo.dwAvailPageFile / 1024 & " KB"
    PageFile.Caption = Format(PageUsed / meminfo.dwTotalPageFile, "0.00%")  '"Paging Memory Usage: " &
    Label7.Caption = "Total Virtual Memory: " & meminfo.dwTotalVirtual / 1024 & " KB"
    Label8.Caption = "Free Virtual Memory: " & meminfo.dwAvailVirtual / 1024 & " KB"
    VirtualMem.Caption = Format(VirtUsed / meminfo.dwTotalVirtual, "0.00%")  ' "Virtual Memory Usage: " &
    
    Exit Sub ' Normal Sub Termination
errHandle:
        ' Logging error
    LogError "UpdateInfo"
End Sub


' **
' ** Check-Box click
' **

Private Sub Logging_c_Click()
    EntryMod = 1
    LogTimer.Enabled = (Logging_c.Value = 1)
    LogTimer_Timer
End Sub

' **
' ** Refresh Timer
' **

Private Sub Timer1_Timer()
    Update_Info
End Sub

' **
' ** Log Handling
' **

    ' Timer event
Private Sub LogTimer_Timer()
    entry = PhysicalMem.Caption & "/" & PageFile.Caption & "/" & VirtualMem
    If Log.List(Log.ListIndex) = entry Then Exit Sub
    UpdateLog entry, 1 + EntryMod
    EntryMod = 0
End Sub

    ' Resetting the log
Private Sub ResetLog()
        ' Setting log files directory
    logPath = CurDir & "\Logs\"
        ' Making sure path is created
    createPath logPath
        ' refreshing labels
    Path_l.Caption = logPath
    Path_l.ToolTipText = logPath
            ' Getting application version
    ver = App.Major & "." & App.Minor & "." & App.Revision
    Version_l.Caption = "Build " & ver
        ' Setting to next log file (to not erase last one)
    AdvanceLogNum
        ' resetting list boxes
    Log.Clear
    LogTime.Clear
        ' Making first log entries
        ' NOTE: log auto-saves on every entry update
    UpdateLog Date & " " & Time, 1
    UpdateLog "ResourceMonitor " & ver, 1
End Sub

    ' Showing entry details
Private Sub Log_Click()
        ' If not a Code-Generated event
    If logInternalSelect Then Exit Sub
        ' show details
    If Not (Log.Text = "") Then
        MsgBox LogTime.List(Log.ListIndex) & " - " & Log.List(Log.ListIndex)
    Else
        If Not (Log.ListIndex = -1) Then Log.Selected(Log.ListIndex) = False
    End If
End Sub

    ' Resetting log
    ' Save here will auto-Advance to next log file
Private Sub Reset_But_Click()
    Response = MsgBox("Save current log before reset ?", vbYesNoCancel + vbQuestion)
    If Not (Response = vbCancel) Then       'Do the reset
        If Response = vbYes Then Call SaveLog("", True) 'save log
        Log_Form.Log.Clear
    End If
End Sub

