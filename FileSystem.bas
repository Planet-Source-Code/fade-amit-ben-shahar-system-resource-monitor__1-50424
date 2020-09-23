Attribute VB_Name = "FileSystem"
' *****************************************************
' *****************************************************
' ********    Two file system methods (Fade)
' ********
' ******** These two file system methods are very usefull
' ********
' ********          CreatePath
' ******** this makes sure to crate specified path from scratch-
' ********  this means that in one line u can created a complete path
' ********  (exp: you dont even have 'c:\tmp', but you can still use:
' ********        CratePath ("C:\tmp\another\alsothis\whatthehell")
' ********   and the complete path will be created)
' ********
' ********          FileExists
' ******** Simple enough to check if specified file exists in system
' ******** and returns Boolean Ture/False value. Simple?
' ********
    
    ' creates a complete path- as passed to path
    '  - path has to contain a complete path with a
    '    route directory (exp. "C:\...")
    '  - self managed error trapping
Public Sub createPath(ByVal Path As String)
    On Error Resume Next

    Dim Index As Integer
    Dim Dir As String
        ' and go on:
    
        ' creating file system object
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    Select Case Len(Path)
            ' if invalid
        Case Is <= 3
            Exit Sub
            
            ' if directory set
        Case Is > 3
                ' checking validity
            If Not (Mid(Path, 2, 2) = ":\") Then
                MsgBox "Invalid directory syntax."
                Exit Sub
            End If
                ' getting and verifing route
            route = Mid(Path, 1, 3)
            If Not (fs.folderexists(route)) Then
                MsgBox "Invalid route directory."
            End If
            Path = Mid(Path, 4) ' and removing from string
            Do
                Index = InStr(Path, "\")
                    ' if slash not found
                If Index = 0 Then
                    fs.createfolder (route & Path)
                Else    ' if found
                    Dir = Mid(Path, 1, Index - 1) ' extract directory
                        ' create dir
                    fs.createfolder (route & Dir)
                        ' set route and path for next loop
                    route = route & Dir & "\"
                            ' if not a closing slash (c:\temp'\')
                    If Len(Path) > Index Then
                        Path = Mid(Path, Index + 1) ' update path
                    Else ' if is a closing slash
                        Index = 0 ' exit loop
                    End If
                End If
            Loop Until Index = 0
        
    End Select
        
End Sub

    ' Integrates two path parts together, inforcing a correct
    ' path syntax
Public Function compilePath(ByVal Header As String, ByVal Addition As String, Optional ByVal NoClosingSlash As Boolean) As String
        ' default - simple combine
    compilePath = Header & Addition
        ' validating extended combine
    If Not (Len(Header) > 0) Then Exit Function
    If Not (Len(Addition) > 0) Then Exit Function
        ' removing slash from header (if exists)
    If Right(Header, 1) = "\" Then Header = Left(Header, Len(Header) - 1)
        ' removing slash from Addition (if exists)
    If Left(Addition, 1) = "\" Then Addition = Mid(Header, 2, Len(Header) - 1)
        ' Combining two validated parts
    compilePath = Header & "\" & Addition
        ' Making sure there is a closing Slash
    If Not (NoClosingSlash) And Not (Right(compilePath, 1) = "\") Then compilePath = compilePath & "\"
End Function
