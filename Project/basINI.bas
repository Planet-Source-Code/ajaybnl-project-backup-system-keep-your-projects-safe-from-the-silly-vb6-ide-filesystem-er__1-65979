Attribute VB_Name = "basIni"
Option Explicit
Private IniFile   As String
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, _
                                                                                                  ByVal lpKeyName As Any, _
                                                                                                  ByVal lpDefault As String, _
                                                                                                  ByVal lpReturnedString As String, _
                                                                                                  ByVal nSize As Long, _
                                                                                                  ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, _
                                                                                                      ByVal lpKeyName As Any, _
                                                                                                      ByVal lpString As Any, _
                                                                                                      ByVal lpFileName As String) As Long

Public Function CheckSettings() As Boolean
If ReadIniFile("Settings", "BackupsPath", "") = "" Or Dir(ReadIniFile("Settings", "BackupsPath", "TEMP"), vbDirectory) = "" Then
CheckSettings = False
Else
CheckSettings = True
End If



End Function
Private Sub LoadIniPath()

    On Error Resume Next
    IniFile = IIf(Right$(App.Path, 1) = "\", App.Path, App.Path & "\") & App.EXEName & ".ini"
    err.Clear
    

End Sub

Public Function ReadIniFile(ByVal sSection As String, _
                            ByVal sItem As String, _
                            ByVal sDefault As String) As String

Dim iRetAmount As Integer
Dim sTemp      As String

    On Error Resume Next
    LoadIniPath
    sTemp = String$(150, 0)
    iRetAmount = GetPrivateProfileString(sSection, sItem, sDefault, sTemp, 150, IniFile)
    sTemp = Left$(sTemp, iRetAmount)
    ReadIniFile = sTemp
    err.Clear
    

End Function

Public Function WriteIniFile(ByVal sSection As String, _
                             ByVal sItem As String, _
                             ByVal sText As String) As Boolean

    On Error Resume Next
    LoadIniPath
    WritePrivateProfileString sSection, sItem, sText, IniFile
    WriteIniFile = True
    err.Clear
    

End Function


