Attribute VB_Name = "basSafe"
Option Explicit
Private Const INVALID_HANDLE_VALUE     As Integer = -1
Private Const MAX_PATH                 As Integer = 260
Private Type FILETIME
    dwLowDateTime                          As Long
    dwHighDateTime                         As Long
End Type
Private Type WIN32_FIND_DATA
    dwFileAttributes                       As Long
    ftCreationTime                         As FILETIME
    ftLastAccessTime                       As FILETIME
    ftLastWriteTime                        As FILETIME
    nFileSizeHigh                          As Long
    nFileSizeLow                           As Long
    dwReserved0                            As Long
    dwReserved1                            As Long
    cFileName                              As String * MAX_PATH
    cAlternate                             As String * 14
End Type
''Private Const WM_SETTEXT               As Long = &HC
Private safesavename                   As String
''Private Declare Sub Sleep Lib "Kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, _
                                                                              lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, _
                                                                                                ByVal lpBuffer As String) As Long

Public Function FileExists(sSource As String) As Boolean

Dim allDrives As String
Dim WFD       As WIN32_FIND_DATA
Dim hFile     As Long

    On Error Resume Next
    If Right$(sSource, 2) = ":\" Then
        allDrives = Space$(64)
        GetLogicalDriveStrings Len(allDrives), allDrives
        FileExists = InStr(1, allDrives, Left$(sSource, 1), 1) > 0
    Else
        If Not LenB(sSource) = 0 Then
            hFile = FindFirstFile(sSource, WFD)
            FileExists = hFile <> INVALID_HANDLE_VALUE
            FindClose hFile
        Else
            FileExists = False
        End If
    End If
    err.Clear
    

End Function

Public Function SafeSave(ByVal Path As String) As String

Dim mPath As String
Dim mname As String
Dim mTemp As String
Dim mfile As String
Dim mExt  As String
Dim m     As Integer

    On Error Resume Next

    mPath = Mid$(Path, 1, InStrRev(Path, "\"))
    mname = Mid$(Path, InStrRev(Path, "\") + 1)
    mfile = Left$(Mid$(mname, 1, InStrRev(mname, ".")), Len(Mid$(mname, 1, InStrRev(mname, "."))) - 1)
    If LenB(mfile) = 0 Then
        mfile = mname
    End If
    mExt = Mid$(mname, InStrRev(mname, "."))
    mTemp = vbNullString
    Do
        If Not FileExists(mPath + mfile + mTemp + mExt) Then
            SafeSave = mPath + mfile + mTemp + mExt
            safesavename = mfile + mTemp + mExt
            Exit Do
        End If
        m = m + 1
        mTemp = Right$(Str$(m), Len(Str$(m)) - 1)
    Loop

    err.Clear


End Function


