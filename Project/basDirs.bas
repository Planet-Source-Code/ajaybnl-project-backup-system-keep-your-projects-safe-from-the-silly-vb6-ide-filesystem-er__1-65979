Attribute VB_Name = "basDirs"
Option Explicit

Public Sub fnCreateDirs(ByVal Directory As String)

Dim A1 As Long
Dim A3 As String

On Error Resume Next
    A1 = 3
re:
    A1 = InStr(A1 + 1, Directory, "\")
    If A1 > 0 Then
        A3 = Mid$(Directory, 1, A1)
        If LenB(Dir(A3, vbDirectory)) = 0 Then
            MkDir A3
        End If
        GoTo re
    End If
    If LenB(Dir(Directory, vbDirectory)) = 0 Then
        MkDir Directory
    End If
err.Clear

End Sub

