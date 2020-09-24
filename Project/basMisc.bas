Attribute VB_Name = "basMisc"
Option Explicit
''Private Declare Function GetShortPathName Lib "Kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, _
                                                                                        ByVal nSize As Long) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, _
                                                                          ByVal lpBuffer As String) As Long



Public Function GetExt(Path As String) As String

Dim A1 As Long

    On Error Resume Next
    If Not LenB(Path) = 0 Then
        A1 = InStrRev(Path, ".", Len(Path))
        If A1 > 0 Then
            GetExt = LCase$(Mid$(Path, A1 + 1))
        End If
    End If
    err.Clear
    

End Function

Public Function GetFile(Path, _
                        ByVal File As Boolean, _
                        Optional Ends As String) As String


Dim A1 As Long

    On Error Resume Next
    If Not LenB(Path) = 0 Then
        If Not File Then
            A1 = InStrRev(Path, "\", Len(Path))
            If A1 > 0 Then
                GetFile = Mid$(Path, 1, A1 - 1) & Ends
            End If
        Else
            A1 = InStrRev(Path, "\", Len(Path))
            If A1 > 0 Then
                GetFile = Mid$(Path, A1 + 1) & Ends
            Else
                GetFile = Path
            End If
        End If
    End If
    err.Clear
    

End Function

Public Function RemExt(Path As String) As String

Dim A1 As Long
Dim A2 As String

    On Error Resume Next
    If Not LenB(Path) = 0 Then
        A2 = GetFile(Path, True)
        A1 = InStrRev(A2, ".", Len(A2))
        If A1 > 0 Then
            RemExt = LCase$(Mid$(A2, 1, A1 - 1))
        End If
    End If
    err.Clear
    

End Function

Public Function SystemDirectory() As String

Dim WinPath As String

    On Error Resume Next
    WinPath = String$(145, vbNullChar)
    GetSystemDirectory WinPath, 145
    SystemDirectory = Left$(WinPath, InStr(WinPath, vbNullChar) - 1)
    err.Clear
    

End Function

Public Function TempDirectory() As String


Dim WinPath As String

    On Error Resume Next
    WinPath = String$(145, vbNullChar)
    GetTempPath 145, WinPath
    TempDirectory = Left$(WinPath, InStr(WinPath, vbNullChar) - 1)
    err.Clear
    

End Function


