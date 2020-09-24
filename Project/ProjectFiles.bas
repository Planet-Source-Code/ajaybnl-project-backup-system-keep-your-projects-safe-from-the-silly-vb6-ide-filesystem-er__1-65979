Attribute VB_Name = "ProjectFile"
Option Explicit
Private Repair   As Boolean

Public Function ChangeProjectFile(ByVal VBPFile As String, _
                                  FileToChange As String, _
                                  ByVal ChangeWith As String) As String


Dim A1    As String
Dim A2    As String
Dim A4    As String
Dim Log   As String
Dim FileC As String

    On Error GoTo err
    If LenB(VBPFile) Then
        FileC = ReadFile(VBPFile)
        If LenB(FileC) = 0 Then
            Exit Function
        End If
        If Len(FileC) <= 50 Then
'MsgBox "File Not OK " & VBPFile
            Exit Function
        End If
        Open VBPFile For Input As #4
        Do While Not EOF(4)
            Line Input #4, A1
            A2 = LCase$(A1)
            If InStr(1, A1, GetFile(FileToChange, True), vbTextCompare) > 0 Then
                A2 = Split(A1, "=")(1)
                If InStr(1, A2, ";") > 0 Then
                    A2 = Replace$(Trim$(Split(A2, ";")(1)), Chr$(34), vbNullString)
                Else
                    A2 = A2
                End If
                If LenB(A2) Then
                    A4 = Replace$(A1, A2, ChangeWith)
                    FileC = Replace$(FileC, A1, A4)
                    Close #4
                    If Len(FileC) > 50 Then
                        Kill VBPFile
                        WriteFileb VBPFile, FileC
                    End If
                    Log = Log & "Updated File : " & VBPFile & vbNewLine & "Updated Refrence To : " & A4 & vbNewLine
                    GoTo OK
                End If
            End If
        Loop
        Close #4
    End If
OK:
    ChangeProjectFile = Log
    Log = vbNullString
    If FileLen(VBPFile) < 50 Then
        MsgBox "File Bytes Truncated Due to error in function ChangeProjectFile " & VBPFile
    End If

Exit Function

err:
    Close #4
    MsgBox "ChangeProjectFile File " & VBPFile & vbNewLine & err.Description
    err.Clear

End Function

Private Function ReadFile(ByVal File As String) As String

Dim D As String

    On Error GoTo err
    Open File For Binary As #1
    D = String$(LOF(1), 0)
    Get #1, , D
    Close #1
    ReadFile = D

Exit Function

err:
    Close #1
    MsgBox "Read File " & File & vbNewLine & err.Description
    err.Clear

End Function

Public Function SaveProjectFilesToProjectDir(VBPFile As String) As String


Dim Path            As String
Dim A1              As String
Dim ModuleFile      As String
Dim Log             As String
Dim FileToChange    As String
Dim StrProjectTitle As String

' On Error GoTo err
re:
    Close #4
    Open VBPFile For Input As #4
    Do While Not EOF(4)
        Line Input #4, A1
        A1 = LCase$(A1)
        If LCase$(Left$(A1, 6)) = "title=" Then
            StrProjectTitle = Trim$(Replace$(Split(A1, "=")(1), Chr$(34), vbNullString))
'If LenB(StrProjectTitle) And Not StrProjectTitle = RemExt(VBPFile) Then
'If LenB(Dir(GetFile(VBPFile, False, "\") & SafeFileName(StrProjectTitle & ".vbp"))) = 0 Then
'Close #4
'Name VBPFile As GetFile(VBPFile, False, "\") & SafeFileName(StrProjectTitle & ".vbp")
'Restart = True
'Exit Function
'End If
'If LenB(Dir(GetFile(VBPFile, False, "\") & GetFile(VBPFile, True) & ".vbw")) Then
''Close #4
'Name GetFile(VBPFile, False, "\") & GetFile(VBPFile, True) & ".vbw" As GetFile(VBPFile, False, "\") & StrProjectTitle & ".vbw"
'End If
'End If
        ElseIf LCase$(Left$(A1, 5)) = "form=" Or LCase$(Left$(A1, 7)) = "module=" Or LCase$(Left$(A1, 12)) = "usercontrol=" Or LCase$(Left$(A1, 6)) = "class=" Or LCase$(Left$(A1, 10)) = "resfile32=" Then 'NOT LENB(DIR(GETFILE(VBPFILE,...
            A1 = Split(A1, "=")(1)
            If InStr(A1, ";") > 0 Then
                A1 = Replace$(Trim$(Split(A1, ";")(1)), Chr$(34), vbNullString)
            End If
'File To Change
            FileToChange = Replace$(GetFile(A1, True), Chr$(34), "")
'get File Name
            If Right$(Left$(A1, 2), 1) = ":" Then
                If LenB(Dir(A1 & "")) > 0 Then
                    ModuleFile = Dir(A1)
                Else 'NOT DIRE(A1...
                    ModuleFile = GetFile(A1, True)
                End If
                Path = GetFile(A1, False)
            Else
                ModuleFile = GetFile(A1, True)
                Path = GetFile(GetFile(VBPFile, False, "\") & Trim$(A1), False)
            End If
            Path = Replace$(Path, Chr$(34), "")
            ModuleFile = Replace$(ModuleFile, Chr$(34), "")
            If ModuleFile = "" Then
                GoTo nextfile
            End If
'If Fixed path
            If Not LCase$(Path) = LCase$(GetFile(VBPFile, False, "")) And InStr(1, Path, ":") > 0 Then
                If LCase$(GetFile(VBPFile, False, vbNullString)) = LCase$(Path) Then
                    Close #4
                    ChangeProjectFile VBPFile, FileToChange, ModuleFile
'Exit Function
                    GoTo re
'if the file exists
                ElseIf LenB(Dir(Path & "\" & ModuleFile)) > 0 Then
'if project folder has already that file then bakup it
                    If LenB(Dir(GetFile(VBPFile, False, "\") & ModuleFile)) Then
                        Close #4
                        Name GetFile(VBPFile, False, "\") & ModuleFile As GetFile(VBPFile, False, "\") & ModuleFile & ".bkp"
                        GoTo re
'Exit Function
                    End If
                    Close #4
'copy to project path and rename project refrence
                    FileCopy Path & "\" & ModuleFile, GetFile(VBPFile, False, "\") & ModuleFile
                    ChangeProjectFile VBPFile, FileToChange, ModuleFile
                    GoTo re
'Exit Function
                Else 'NOT DIRE(PATH...
'if the file not exists in the fixed path
'if project path has the file
                    If LenB(Dir(GetFile(VBPFile, False, "\") & ModuleFile)) Then
                        Close #4
                        ChangeProjectFile VBPFile, FileToChange, ModuleFile
                        GoTo re
'Exit Function
                    End If
                End If
            End If
            If LenB(Dir(Path & "\" & ModuleFile)) > 0 And LenB(ModuleFile) > 0 Then
FileRepaired:
                If Not LCase$(Path) = LCase$(GetFile(VBPFile, False)) Then
'If File Path is Not Project Path
                    FileCopy Path & "\" & ModuleFile, GetFile(VBPFile, False, "\") & ModuleFile
                    Close #4
                    ChangeProjectFile VBPFile, FileToChange, ModuleFile
Log = Log & vbNewLine & "File : " & ModuleFile & " From : " & Path & " Copied To : " & GetFile(VBPFile, False)
                    GoTo re
                Else
'Close #4
'ChangeProjectFile VBPFile, ModuleFile, GetFile(VBPFile, False, "\") & ModuleFile
'Log = Log & vbNewLine & "File : " & ModuleFile & " From : " & Path & " Already Found in : " & GetFile(VBPFile, False)
'GoTo re
                End If
            Else
                If LenB(Dir(GetFile(VBPFile, False) & "\" & ModuleFile)) Then
                    Path = GetFile(VBPFile, False)
                    Close #4
                    ChangeProjectFile VBPFile, FileToChange, ModuleFile
Log = Log & vbNewLine & "File : " & ModuleFile & " From : " & Path & " Already Found in : " & GetFile(VBPFile, False)
                    GoTo re
                End If
                Log = Log & vbNewLine & "File : " & ModuleFile & " From : " & Path & " Not Found"
            End If
        End If
nextfile:
    Loop
    Close #4
    SaveProjectFilesToProjectDir = Log

Exit Function

err:
    Close #4
    MsgBox "SaveProjectFilesToProjectDir File " & VBPFile & vbNewLine & err.Description
    err.Clear

End Function

Private Sub WriteFileb(ByVal File As String, _
                       ByVal Data As String)

Dim D As String

    On Error GoTo err
    D = Data
    Open File For Binary As #1
    Put #1, , D
    Close #1

Exit Sub

err:
    Close #1
    MsgBox "Write File " & File & vbNewLine & err.Description
    err.Clear

End Sub


