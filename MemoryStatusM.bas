Attribute VB_Name = "MemoryStatusM"
' ********
' ******** Just this little API declaration
' ******** to get things started ...
' ********

Public Type MemoryStatus
        dwLength As Long
        dwMemoryLoad As Long
        dwTotalPhys As Long
        dwAvailPhys As Long
        dwTotalPageFile As Long
        dwAvailPageFile As Long
        dwTotalVirtual As Long
        dwAvailVirtual As Long
End Type

Public Declare Sub GlobalMemoryStatus Lib "kernel32" _
    (lpBuffer As MemoryStatus)

