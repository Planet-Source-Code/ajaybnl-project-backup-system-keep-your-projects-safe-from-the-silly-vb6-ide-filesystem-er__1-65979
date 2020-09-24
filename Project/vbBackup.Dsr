VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} Connect 
   ClientHeight    =   10500
   ClientLeft      =   1740
   ClientTop       =   1545
   ClientWidth     =   13470
   _ExtentX        =   23760
   _ExtentY        =   18521
   _Version        =   393216
   Description     =   "Automatically backup the projects when opening and closing them"
   DisplayName     =   "Project Backup System 1.3"
   AppName         =   "Visual Basic"
   AppVer          =   "Visual Basic 6.0"
   LoadName        =   "Startup"
   LoadBehavior    =   1
   RegLocation     =   "HKEY_CURRENT_USER\Software\Microsoft\Visual Basic\6.0"
   CmdLineSupport  =   -1  'True
End
Attribute VB_Name = "Connect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private FormDisplayed            As Boolean
Public VBInstance                As VBIDE.VBE
Private mcbMenuCommandBar        As Office.CommandBarControl
Private mFrmMain                 As New frmMain
Public WithEvents MenuHandler    As CommandBarEvents    'command bar event handler
Attribute MenuHandler.VB_VarHelpID = -1
Private blnProjectBackuped       As Boolean
Private Const fixMenuName        As String = "Project Backup System 1.3"

Private Sub AddinInstance_OnBeginShutdown(custom() As Variant)

On Error GoTo err

    If Not VBInstance.VBProjects.Count <= 0 Then
        If ReadIniFile("Settings", "AutoBackup", "1") = "1" Then
            If CheckSettings Then
                If ReadIniFile("Backup", "Backup - " & VBInstance.ActiveVBProject.Name, "0") = "1" Then
                    If ReadIniFile("Settings", "Backupon", "Project Close") = "Project Close" Then
                        VBInstance.ActiveVBProject.SaveAs VBInstance.ActiveVBProject.Filename
                        With mFrmMain
                            .Visible = False
                            Set .VBInstance = VBInstance
                            .LoadAll
                        End With 'mFrmMain
                        mFrmMain.Backup

                    End If
                End If
            End If
        End If
        Unload mFrmMain
    End If
Exit Sub
err:
err.Clear
End Sub

Private Sub AddinInstance_OnConnection(ByVal Application As Object, _
                                       ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, _
                                       ByVal AddInInst As Object, _
                                       custom() As Variant)

    On Error GoTo error_handler
    Set VBInstance = Application
    Set mcbMenuCommandBar = AddToAddInCommandBar(fixMenuName)
    Set Me.MenuHandler = VBInstance.Events.CommandBarEvents(mcbMenuCommandBar)
    
    If ReadIniFile("Settings", "AutoBackup", "1") = "1" Then
        If CheckSettings Then
            If ReadIniFile("Backup", "Backup - " & VBInstance.ActiveVBProject.Name, "0") = "1" Then
                If ReadIniFile("Settings", "Backupon", "Project Close") = "Project Open" Then
                    With mFrmMain
                        .Visible = False
                        Set .VBInstance = VBInstance
                        .LoadAll
                        .WaitForProjectOpen
                    End With 'mFrmMain
                End If
            End If
        End If
    End If

Exit Sub

error_handler:
'MsgBox err.Description
    err.Clear

End Sub

Private Sub AddinInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, _
                                          custom() As Variant)

On Error Resume Next

    mcbMenuCommandBar.Delete
    Set mFrmMain = Nothing
err.Clear

End Sub

Private Function AddToAddInCommandBar(ByVal sCaption As String) As Office.CommandBarControl

Dim cbMenuCommandBar As Office.CommandBarControl   'command bar object
Dim cbMenu           As Object

On Error Resume Next
Set cbMenu = VBInstance.CommandBars("Add-Ins")
    If Not cbMenu Is Nothing Then
        Set cbMenuCommandBar = cbMenu.Controls.add(1)
        cbMenuCommandBar.Caption = sCaption
        Set AddToAddInCommandBar = cbMenuCommandBar
    End If
err.Clear
End Function


Public Sub Hide()
On Error Resume Next
    FormDisplayed = False
    Unload mFrmMain
err.Clear

End Sub

Private Sub MenuHandler_Click(ByVal CommandBarControl As Object, _
                              handled As Boolean, _
                              CancelDefault As Boolean)

    If CommandBarControl.Caption = fixMenuName Then
        Showfrm
    End If

End Sub

Public Sub Showfrm()
On Error Resume Next
    Set mFrmMain = New frmMain
    With mFrmMain
        Set .VBInstance = VBInstance
        Set .Connect = Me
        .LoadAll
        .CountFiles
        .Show
        .ZOrder 0
    End With 'mFrmMain
    FormDisplayed = True
err.Clear

End Sub


