VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "vbBackup 1.2"
   ClientHeight    =   5520
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   6945
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "FrmMain"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   6945
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Settings"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   6135
      Begin VB.ComboBox Combo3 
         Height          =   315
         ItemData        =   "frmMain.frx":0000
         Left            =   2160
         List            =   "frmMain.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   37
         Top             =   1680
         Width           =   3375
      End
      Begin VB.CommandButton Command5 
         Caption         =   "OK"
         Height          =   375
         Left            =   4440
         TabIndex        =   33
         Top             =   3360
         Width           =   1335
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Autobackup Enabled (please remember)"
         Height          =   285
         Left            =   2160
         TabIndex        =   12
         Top             =   2760
         Value           =   1  'Checked
         Width           =   3375
      End
      Begin VB.CommandButton Command2 
         Caption         =   ".."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5640
         TabIndex        =   11
         Top             =   720
         Width           =   255
      End
      Begin VB.CommandButton Command1 
         Caption         =   ".."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5640
         TabIndex        =   10
         Top             =   360
         Width           =   255
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   2160
         TabIndex        =   9
         Top             =   720
         Width           =   3375
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   2160
         TabIndex        =   3
         Top             =   360
         Width           =   3375
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frmMain.frx":0004
         Left            =   2160
         List            =   "frmMain.frx":0006
         TabIndex        =   2
         Text            =   "10"
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   2160
         TabIndex        =   1
         Text            =   "bas,cls,ctl,ctx,dca,ddf,dep,dob,dox,dsr,dsx,dws,frm,frx,log,oca,pag,pgx,res,tlb,vbg,vbl,vbp,vbr,vbw,vbz,wct"
         Top             =   2160
         Width           =   3375
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Backup on :"
         Height          =   195
         Left            =   240
         TabIndex        =   38
         Top             =   1680
         Width           =   840
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Backup to Folder:"
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   720
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Projects Path:"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   1020
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Max Project Backups :"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   1320
         Width           =   1590
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Backup File Types:"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Default"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   4920
         MouseIcon       =   "frmMain.frx":0008
         MousePointer    =   99  'Custom
         TabIndex        =   4
         Top             =   2460
         Width           =   615
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Restore"
      Height          =   3975
      Left            =   240
      TabIndex        =   20
      Top             =   360
      Visible         =   0   'False
      Width           =   6135
      Begin VB.CommandButton Command3 
         Caption         =   "OK"
         Height          =   375
         Left            =   4440
         TabIndex        =   31
         Top             =   3360
         Width           =   1335
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   600
         Width           =   5655
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         X1              =   270
         X2              =   5900
         Y1              =   1830
         Y2              =   1830
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Current Project Backups:"
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "File Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   29
         Top             =   1080
         Width           =   810
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date Time"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3840
         TabIndex        =   28
         Top             =   1080
         Width           =   870
      End
      Begin VB.Label lbl_filename 
         BackStyle       =   0  'Transparent
         Caption         =   "None"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   360
         TabIndex        =   27
         Top             =   1440
         Width           =   3180
      End
      Begin VB.Label lbl_datetime 
         BackStyle       =   0  'Transparent
         Caption         =   "None"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   3840
         TabIndex        =   26
         Top             =   1440
         Width           =   1740
      End
      Begin VB.Label btn_openfile 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Open File"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Left            =   360
         MouseIcon       =   "frmMain.frx":0312
         MousePointer    =   99  'Custom
         TabIndex        =   25
         Top             =   1920
         Width           =   825
      End
      Begin VB.Label btn_delfile 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Delete File"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   210
         Left            =   2520
         MouseIcon       =   "frmMain.frx":061C
         MousePointer    =   99  'Custom
         TabIndex        =   24
         Top             =   1920
         Width           =   930
      End
      Begin VB.Label btn_saveas 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Save As"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Left            =   1440
         MouseIcon       =   "frmMain.frx":0926
         MousePointer    =   99  'Custom
         TabIndex        =   23
         Top             =   1920
         Width           =   720
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00000000&
         X1              =   250
         X2              =   5890
         Y1              =   1810
         Y2              =   1810
      End
      Begin VB.Label btn_restore 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Restore"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   4800
         MouseIcon       =   "frmMain.frx":0C30
         MousePointer    =   99  'Custom
         TabIndex        =   22
         Top             =   1920
         Width           =   720
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H00C0FFC0&
         FillStyle       =   0  'Solid
         Height          =   1095
         Left            =   240
         Top             =   1320
         Width           =   5655
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Main"
      Height          =   3975
      Left            =   480
      TabIndex        =   13
      Top             =   720
      Width           =   6135
      Begin VB.ListBox List1 
         Height          =   1230
         Left            =   240
         TabIndex        =   34
         Top             =   1800
         Width           =   5655
      End
      Begin VB.CommandButton Command4 
         Caption         =   "OK"
         Height          =   375
         Left            =   4440
         TabIndex        =   32
         Top             =   3360
         Width           =   1335
      End
      Begin VB.Label btnBuildZip 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Build Zip File"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   3120
         TabIndex        =   40
         ToolTipText     =   "Build Zip File For Submission on PSC.com"
         Top             =   3480
         Width           =   1035
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Author : Ajay ajaybnl@gmail.com"
         Height          =   495
         Left            =   240
         TabIndex        =   39
         Top             =   3360
         Width           =   1575
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Detailed Backups :"
         Height          =   255
         Left            =   240
         TabIndex        =   35
         Top             =   1560
         Width           =   2775
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2760
         TabIndex        =   19
         Top             =   1080
         Width           =   2000
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Backup Files :"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   1080
         Width           =   2415
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2760
         TabIndex        =   17
         Top             =   720
         Width           =   2000
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Backuped Projects :"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   720
         Width           =   2415
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2760
         TabIndex        =   15
         Top             =   360
         Width           =   2000
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Current Backups for this Project :"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   360
         Width           =   2415
      End
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   840
      Top             =   4920
   End
   Begin VB.FileListBox File1 
      Height          =   675
      Left            =   0
      TabIndex        =   36
      Top             =   4800
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Menu menu1 
      Caption         =   "Actions For Current Project"
      Begin VB.Menu autobkp 
         Caption         =   "AutoBackup Current Project"
      End
      Begin VB.Menu l0 
         Caption         =   "-"
      End
      Begin VB.Menu remautobkp 
         Caption         =   "Remove From AutoBackup"
      End
      Begin VB.Menu l98 
         Caption         =   "-"
      End
      Begin VB.Menu backup1 
         Caption         =   "Take A Backup"
      End
      Begin VB.Menu l1 
         Caption         =   "-"
      End
      Begin VB.Menu restore1 
         Caption         =   "Select Backups And Restore"
      End
   End
   Begin VB.Menu mENU2 
      Caption         =   "General"
      Begin VB.Menu settings1 
         Caption         =   "Settings"
      End
      Begin VB.Menu l4 
         Caption         =   "-"
      End
      Begin VB.Menu refresh1 
         Caption         =   "Refresh"
      End
      Begin VB.Menu exnt1 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnu2 
      Caption         =   "Project"
      Visible         =   0   'False
      Begin VB.Menu runproj1 
         Caption         =   "Run Project"
      End
      Begin VB.Menu l5 
         Caption         =   "-"
      End
      Begin VB.Menu seeprojfiles1 
         Caption         =   "Open Backups"
      End
      Begin VB.Menu l6 
         Caption         =   "-"
      End
      Begin VB.Menu delprojbkp1 
         Caption         =   "Delete Project Backups"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public VBInstance           As VBIDE.VBE
Public Connect              As Connect
Option Explicit
Private Type typBackups
    ZipFile                     As String
    ProjectFile                 As String
    nDate                       As String
    nTime                       As String
End Type
Private typBackupsInfo()    As typBackups
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, _
                                                                               ByVal lpOperation As String, _
                                                                               ByVal lpFile As String, _
                                                                               ByVal lpParameters As String, _
                                                                               ByVal lpDirectory As String, _
                                                                               ByVal nShowCmd As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, _
                                                    ByVal hWndInsertAfter As Long, _
                                                    ByVal x As Long, _
                                                    ByVal Y As Long, _
                                                    ByVal cx As Long, _
                                                    ByVal cy As Long, _
                                                    ByVal wFlags As Long) As Long

Private Sub autobkp_Click()
On Error Resume Next
29      WriteIniFile "Backup", "Backup - " & VBInstance.ActiveVBProject.Name, "1"
30      MsgBox "Now Autobackup will always backup this project, If enabled in settings!"
31      Connect.Hide
32  err.Clear
End Sub

Public Function Backup(Optional SaveToZip As String, Optional ShowErrors As Boolean) As Boolean
If CheckSettings = False Then
If ShowErrors Then MsgBox "Please open settings and specify backup and project paths!"
Backup = False
Exit Function
End If

If Dir(SystemDirectory & "\unzip.dll") = "" Or Dir(App.Path & "\unzip.dll") = "" Then
MsgBox "Two dll files: unzip.dll, zip.dll are needed to perform some actions! Please download Info-zip's vbuzip32.dll and rename it to unzip.dll, and download Info-zip's vbzip32.dll and rename it to zip.dll and place it in app or system directory!"
Exit Function
End If



38  Dim cZipper           As New cZip
39  Dim nDescription      As String
40  Dim strLastBackupFile As String
41  Dim TmpComponents     As Long
42  Dim TmpFiles          As Long
43  Dim lngVarFiles       As Long
44  Dim strProjectFile    As String
45  Dim nBkpFilesType     As String
46  Dim nBkpFolder        As String

    On Error GoTo err
49      If LenB(VBInstance.ActiveVBProject.Filename) > 0 Then
50          SaveProjectFilesToProjectDir VBInstance.ActiveVBProject.Filename
51          VBInstance.ActiveVBProject.SaveAs VBInstance.ActiveVBProject.Filename
        
        
54          If LenB(SaveToZip) > 0 Then
55              GoTo skipFolderCreation
        End If
        
58  'Check Project Path
59          If InStr(1, VBInstance.ActiveVBProject.Filename, Text1.Text, vbTextCompare) > 0 Then
        
61          'Check Max Backups
62              If Not Combo1.Text = "Unlimited" Then
63                  If CountFiles > (Int(Combo1.Text) - 1) Then
64                      PopulateRestoreFiles
65                      If Combo2.ListCount > 0 Then
66                          strLastBackupFile = Combo2.List(0)
67                          Kill Text3.Text & "\" & strLastBackupFile
                    End If
                End If
            End If
71 skipFolderCreation:
72              nBkpFolder = Text3.Text
73              nBkpFilesType = Text2.Text
74              If LenB(Dir(nBkpFolder, vbDirectory)) = 0 Then
75                  fnCreateDirs nBkpFolder
            End If
77              If Not Right$(nBkpFolder, 1) = "\" Then
78                  nBkpFolder = nBkpFolder & "\"
            End If
80              With cZipper
81                  If SaveToZip <> "" Then
82                      .ZipFile = SaveToZip
83                  Else
84                      .ZipFile = SafeSave(nBkpFolder & RemExt(VBInstance.ActiveVBProject.Filename) & ".zip")
                End If
86                  .IncludeSystemAndHiddenFiles = True
87                  .MessageLevel = ezpNoMessages
88                  .StoreDirectories = False
89                  .AddComment = True
            End With 'cZipper
91              With VBInstance
92                  For TmpComponents = 1 To .ActiveVBProject.VBComponents.Count
93                      For TmpFiles = 1 To .ActiveVBProject.VBComponents(TmpComponents).FileCount
94                          strProjectFile = .ActiveVBProject.VBComponents(TmpComponents).FileNames(TmpFiles)
95  'Check If The Extension Choosed To Save
96                          If UBound(Split(nBkpFilesType, ",")) > 0 Then
97                              For lngVarFiles = 0 To UBound(Split(nBkpFilesType, ","))
98                                  If LCase$(GetExt(strProjectFile)) = LCase$(Split(nBkpFilesType, ",")(lngVarFiles)) Then
99                                      cZipper.AddFileSpec strProjectFile
100                                      Exit For
                                End If
102                              Next lngVarFiles
                        End If
104                          strProjectFile = vbNullString
105                      Next TmpFiles
106                  Next TmpComponents
107  ' Add Project.vbp,vbg
108                  strProjectFile = .ActiveVBProject.Filename
109  'Check If The Extension  Choosed To Save (vbp,vbg)
110                  If UBound(Split(nBkpFilesType, ",")) > 0 Then
111                      For lngVarFiles = 0 To UBound(Split(nBkpFilesType, ","))
112                          If LCase$(GetExt(strProjectFile)) = LCase$(Split(nBkpFilesType, ",")(lngVarFiles)) Then
113                              cZipper.AddFileSpec strProjectFile
114                              Exit For
                        End If
116                      Next lngVarFiles
                End If
118  ' Add Project.vbw
119                  strProjectFile = GetFile(.ActiveVBProject.Filename, False) & "\" & RemExt(.ActiveVBProject.Filename) & ".vbw"
120  'Check If The Extension Choosed To Save(vbw)
121                  If UBound(Split(nBkpFilesType, ",")) > 0 Then
122                      For lngVarFiles = 0 To UBound(Split(nBkpFilesType, ","))
123                          If LCase$(GetExt(strProjectFile)) = LCase$(Split(nBkpFilesType, ",")(lngVarFiles)) Then
124                              cZipper.AddFileSpec strProjectFile
125                              Exit For
                        End If
127                      Next lngVarFiles
                End If
            End With
130  'Add INFO To Zip
131              If cZipper.FileSpecCount > 0 Then
132                  With cZipper
133                      For lngVarFiles = 1 To .FileSpecCount
134                          nDescription = nDescription & vbNewLine & _
                         "Added File: '" & GetFile(.FileSpec(lngVarFiles), True) & "' Path: '" & .FileSpec(lngVarFiles) & "'"
136                      Next lngVarFiles
                End With 'cZipper
138                  MCComments = "Project Backup Created With *** vbBackup 1.2 - The Ultimate VB6 Project Backups Manager ***" & vbNewLine & _
                             "Date: " & date & vbNewLine & _
                             "Time: " & Time & vbNewLine & _
                             "Project Path: " & GetFile(VBInstance.ActiveVBProject.Filename, False) & vbNewLine & _
                             "Proect File Name: " & GetFile(VBInstance.ActiveVBProject.Filename, True) & vbNewLine & _
                             "Project Name: " & VBInstance.ActiveVBProject.Name & vbNewLine & _
                             vbNewLine & _
                             "*** Project Files Details ***" & vbNewLine & _
                             nDescription & vbNewLine & _
                             vbNewLine & _
                             "*** Project Files Details End ***" & vbNewLine & _
                             vbNewLine & _
                             "Have a Nice Coding!"
151                  cZipper.CommentRequest "Hello", False
                
                

155                  frmShow.Visible = True
156                  DoEvents
157                  cZipper.Zip
158                  If Not cZipper.success Then
159                      GoTo err
160                  Else
161                      frmShow.Shape1.Visible = True
162                      frmShow.Label2.Visible = True
163                      DoEvents
164                      Sleep 500
165                      Unload frmShow
                End If
            End If
        End If
    End If
170  Backup = True
171  Unload frmShow
172  Exit Function

174 err:
175  Backup = False
176  If ShowErrors Then MsgBox "Error Creating Backup : " & vbCrLf & "Line : " & Erl & vbCrLf & err.Description
177  Unload frmShow

End Function

Private Sub backup1_Click()
On Error Resume Next
183      If Backup(, True) = True Then
184          Connect.Hide
185          Else
        
        End If

    
190      CountFiles
191  err.Clear
End Sub

Private Sub btn_delfile_Click()
On Error Resume Next
196      If MsgBox("Do You Really Want to Delete The Backup File: " & lbl_filename.Caption & " ?", vbYesNo, "Delete Backup File") = vbYes Then
197          Kill btn_restore.Tag
198          Combo2.RemoveItem Combo2.ListIndex
199          If Combo2.ListCount > 0 Then
200              Combo2.ListIndex = 0
        End If
202          CheckButtons
    End If
204  err.Clear
End Sub

Private Sub btn_openfile_Click()
On Error Resume Next
209      ShellExecute Me.hWnd, "open", btn_restore.Tag, "", "", 1
210  err.Clear
End Sub

Private Sub btn_restore_Click()
On Error Resume Next

216  Dim Unzip1            As New cUnzip
217  Dim TmpComponents     As Long
218  Dim TmpFiles          As Long
219  Dim strProjectFile

221  Dim A1                As Integer
222  Dim strLastBackupFile As Integer
223  Dim A3                As Integer

225      If LenB(Dir(btn_restore.Tag)) Then
226          With Unzip1
227              .ZipFile = btn_restore.Tag
228              .UseFolderNames = False

230              .OverwriteExisting = True
231              .PromptToOverwrite = False
232              .Directory
        End With 'Unzip1
234          If MsgBox("Do You Really Want to Restore This Backup ?" & vbNewLine & _
          vbNewLine & _
          "This Will OverRight All The Selected Types Of Files. And Would Not Be UnDone!", vbYesNo, "Restore Backup Files") = vbYes Then
237              frmShow.Show
238              frmShow.Label1.Caption = "Restoring Project..."
239              DoEvents
240              With VBInstance
241                  For A1 = 1 To Unzip1.FileCount
242                      strLastBackupFile = 0
243                      For TmpComponents = 1 To .ActiveVBProject.VBComponents.Count
244                          For TmpFiles = 1 To .ActiveVBProject.VBComponents(TmpComponents).FileCount
245                              strProjectFile = .ActiveVBProject.VBComponents(TmpComponents).FileNames(TmpFiles)
246                              If LCase$(GetFile(strProjectFile, True)) = LCase$(Unzip1.Filename(A1)) Then
247                                  Unzip1.UnzipFolder = GetFile(strProjectFile, False)
248                                  strLastBackupFile = A1
249                                  GoTo Ok1
                            End If
251                          Next TmpFiles
252                      Next TmpComponents
253 Ok1:
254                      If strLastBackupFile > 0 Then
255                          For A3 = 1 To Unzip1.FileCount
256                              Unzip1.FileSelected(A3) = False
257                          Next A3
258                          With Unzip1
259                              .FileSelected(strLastBackupFile) = True
260                              .UnzipFolder = GetFile(strProjectFile, False)
261                              .Unzip
262  'Reload The FIle
                        End With 'Unzip1
On Error Resume Next
265                          For TmpComponents = 1 To .ActiveVBProject.VBComponents.Count
266                              For TmpFiles = 1 To .ActiveVBProject.VBComponents(TmpComponents).FileCount
267                                  strProjectFile = .ActiveVBProject.VBComponents(TmpComponents).FileNames(TmpFiles)
268                                  If LCase$(GetFile(strProjectFile, True)) = LCase$(Unzip1.Filename(A1)) Then
269                                      .ActiveVBProject.VBComponents(TmpComponents).Reload
270                                      GoTo ok2
                                End If
272                              Next TmpFiles
273                          Next TmpComponents
274 ok2:
                    End If
276                  Next A1
            End With
        End If
279      Else
280          MsgBox "Backup File Not Found! Check if you Had Changed The Backup Path."
    End If

283  Exit Sub

285      err.Clear
    

End Sub

Private Sub btn_saveas_Click()
On Error Resume Next
292  Dim A1 As String

294      A1 = DialogFile(Me.hWnd, 2, "Save Backup As", lbl_filename.Caption, "Zip Files" & vbNullChar & "*.zip", "", "Zip")
295      If LenB(A1) Then
296          FileCopy btn_restore.Tag, A1
    End If
298  err.Clear
End Sub

Private Sub btnBuildZip_Click()
On Error Resume Next
303  Dim strFile As String

305      If ReadIniFile("Settings", "BuildZipNewWarning", "0") = "0" Then
306          MsgBox "All Project Files Will Be Saved To Project Dir To Ensure Error Free Zip Build!"
307          WriteIniFile "Settings", "BuildZipNewWarning", "1"
    End If
309      strFile = DialogFile(0, 2, "Save Project Zip File", VBInstance.ActiveVBProject.Name, "Zip Files" & vbNullChar & "*.zip" & vbNullChar & "All Files" & vbNullChar & "*.*", Environ$("UserProfile") & "\Desktop", "zip")
310      If strFile <> "" Then
311          If Backup(strFile, True) = True Then
312          Connect.Hide
313           End If
    End If
316  err.Clear
End Sub

Private Sub Check1_MouseUp(Button As Integer, _
                           Shift As Integer, _
                           x As Single, _
                           Y As Single)

324      WriteIniFile "Settings", "AutoBackup", CStr(Check1.Value)

End Sub

Private Sub CheckButtons()

330      If Combo2.ListCount <= 0 Then
331          btn_openfile.Enabled = False
332          btn_delfile.Enabled = False
333          btn_restore.Enabled = False
334          btn_saveas.Enabled = False
335      Else
336          btn_openfile.Enabled = True
337          btn_delfile.Enabled = True
338          btn_restore.Enabled = True
339          btn_saveas.Enabled = True
    End If

End Sub

Private Sub Combo1_Change()

On Error GoTo err

348      WriteIniFile "Settings", "MaxBackups", Combo1.Text
349      Exit Sub
350 err:
351     Combo1.ListIndex = 0

End Sub

Private Sub Combo1_Click()

357      Combo1_Change

End Sub

Private Sub Combo2_Change()

363      Combo2_Click

End Sub

Private Sub Combo2_Click()

369  Dim D1 As String
370  Dim T1 As String
371  Dim P1 As String
372  Dim E1 As String

374      RestoreButtonsEnabled False
375      lbl_filename.Caption = "None"
376      lbl_datetime.Caption = "None"
377      GetSetZipBackupInfo Text3.Text & "\" & Combo2.Text, D1, T1, P1, E1, True
378      lbl_datetime.Caption = D1 & " " & T1
379      lbl_filename.Caption = Combo2.Text
380      btn_restore.Tag = Text3.Text & "\" & Combo2.Text
381      RestoreButtonsEnabled True

End Sub

Private Sub Combo3_Change()

387      WriteIniFile "Settings", "Backupon", Combo3.Text

End Sub

Private Sub Combo3_Click()

393      Combo3_Change

End Sub

Private Sub Command1_Click()

399  Dim A As String

401      A = BrowseFolder(Me.hWnd, "Locate project's path to use with auto backup!", Text1.Text)
402      If A <> "" Then
403          Text1.Text = A
404          Text1_KeyUp 0, 0
    End If

End Sub

Private Sub Command2_Click()

411  Dim A As String

413      A = BrowseFolder(Me.hWnd, "Locate folder to save backups!", Text3.Text)
414      If A <> "" Then
415          Text3.Text = A
416          Text3_KeyUp 0, 0
    End If

End Sub

Private Sub Command3_Click()

423      Frame1.Visible = False
424      Frame3.Visible = False
425      Frame2.Visible = True
426      Frame2.ZOrder 0

End Sub

Private Sub Command4_Click()

432      Connect.Hide

End Sub

Private Sub Command5_Click()
If CheckSettings = False Then
MsgBox "Invalid paths! Please check the paths!"
Else
menu1.Enabled = True
mENU2.Enabled = True

439      Frame1.Visible = False
440      Frame3.Visible = False
441      Frame2.Visible = True
442      Frame2.ZOrder 0

     CountFiles
End If
End Sub

Public Function CountFiles() As Long
If Dir(SystemDirectory & "\unzip.dll") = "" Or Dir(App.Path & "\unzip.dll") = "" Then
MsgBox "Two dll files: unzip.dll, zip.dll are needed to perform some actions! Please download Info-zip's vbuzip32.dll and rename it to unzip.dll, and download Info-zip's vbzip32.dll and rename it to zip.dll and place it in app or system directory!"
Exit Function
End If


449  Dim A1                As Long
450  Dim D1                As String
451  Dim T1                As String
452  Dim P1                As String
453  Dim E1                As String
454  Dim C1                As Long
455  Dim F1                As String
456  Dim strLastBackupFile As Long

    On Error GoTo err
459      If Not VBInstance.ActiveVBProject.Filename = "" Then
460          With File1
461              .Pattern = "*.zip"
462              .Path = Text3.Text
463              .Refresh
464  ' Load Current Project Backups
465              For A1 = 0 To .ListCount - 1
466                  GetSetZipBackupInfo .Path & "\" & .List(A1), D1, T1, P1, E1, True
467                  If LCase$(P1) = LCase$(VBInstance.ActiveVBProject.Filename) Then
468                      C1 = C1 + 1
                End If
470              Next A1
        End With 'File1
472          Label7.Caption = C1
473          Label11.Caption = File1.ListCount
474          CountFiles = C1
475          C1 = 0
476  'Load Total Backuped Projects
477          For A1 = 0 To File1.ListCount - 1
478              GetSetZipBackupInfo File1.Path & "\" & File1.List(A1), D1, T1, P1, E1, True
479              If InStr(1, F1, P1, vbTextCompare) <= 0 Then
480                  F1 = F1 & "_" & P1
481                  C1 = C1 + 1
            End If
483          Next A1
484          Label9.Caption = C1
485          C1 = 0
486          List1.Clear
487  'Load List
488          For A1 = 0 To File1.ListCount - 1
489              GetSetZipBackupInfo File1.Path & "\" & File1.List(A1), D1, T1, P1, E1, True
490              For strLastBackupFile = 0 To List1.ListCount - 1
491                  If List1.List(strLastBackupFile) = Replace$(P1, Text1.Text, "") Then
492                      GoTo Found
                End If
494              Next strLastBackupFile
495              If Not LenB(Trim$(P1)) = 0 Then
496                  List1.AddItem Replace$(P1, Text1.Text, "")
            End If
498 Found:
499          Next A1
500          For strLastBackupFile = 0 To List1.ListCount - 1
501              F1 = vbNullString
502              C1 = 0
503              For A1 = 0 To File1.ListCount - 1
504                  GetSetZipBackupInfo File1.Path & "\" & File1.List(A1), D1, T1, P1, E1, True
505                  If LCase$(Replace$(P1, Text1.Text, "")) = LCase$(List1.List(strLastBackupFile)) Then
506                      If Not LenB(Trim$(P1)) = 0 Then
507                          C1 = C1 + 1
                    End If
                End If
510              Next A1
511              List1.List(strLastBackupFile) = List1.List(strLastBackupFile) & " (" & C1 & ")"
512          Next strLastBackupFile
    End If
514 err:
515      err.Clear

End Function

Private Sub delprojbkp1_Click()

521  Dim D1 As String
522  Dim T1 As String
523  Dim P1 As String
524  Dim E1 As String
525  Dim A1 As Long
526  Dim F1 As String
527  Dim l1 As String

529  Dim F  As String
On Error GoTo err
531      If LenB(List1.Text) Then
532          With File1
533              .Pattern = "*.zip"
534              .Path = Text3.Text
535              .Refresh
536              l1 = List1.Text
537              List1.RemoveItem List1.ListIndex
538              For A1 = 0 To .ListCount - 1
539                  If Left$(Split(l1, " (")(0), 1) = "\" Then
540                      F = Text1.Text & Split(l1, " (")(0)
541                  Else
542                      F = Split(l1, " (")(0)
                End If
544                  GetSetZipBackupInfo .Path & "\" & .List(A1), D1, T1, P1, E1, True
545                  If LCase$(P1) = LCase$(F) And LenB(Trim$(P1)) Then
546                      Kill .Path & "\" & .List(A1)
547                      F1 = F1 & "Deleted : " & .List(A1) & vbNewLine
                End If
549              Next A1
        End With 'File1
551          If LenB(F1) = 0 Then
552              MsgBox "No File Deleted!"
553          Else
554              MsgBox F1, vbInformation, "Deleted Files Summary"
        End If
    End If
557      CountFiles
558      Exit Sub
559 err:
560  MsgBox "Error Deleting Some Files!", vbCritical

End Sub

Private Sub exnt1_Click()

566      Connect.Hide

End Sub

Private Sub Form_Load()
On Error Resume Next
572      Frame1.Top = 0
573      Frame1.Left = 0
574      Frame2.Top = 0
575      Frame2.Left = 0
576      Frame3.Top = 0
577      Frame3.Left = 0
    
579      Me.Height = Frame1.Top + Frame1.Height + 620
580      Me.Width = Frame1.Width + 100
581      Me.Left = (Screen.Width - Me.Width) / 2
582      Me.Top = (Screen.Height - Me.Height) / 2
583      With Me
    
585  'SetWindowPos .hWnd, -1, .Left / 15, .Top / 15, .Width / 15, .Height / 15, 0
    End With 'Me
    
    If CheckSettings = False Then
    menu1.Enabled = False
    mENU2.Enabled = False
    Frame1.Visible = True
    Frame2.Visible = False
    Frame3.Visible = False
    Frame1.ZOrder 0
    Else
    menu1.Enabled = True
    mENU2.Enabled = True
    End If
587  err.Clear
End Sub

Private Sub GetSetZipBackupInfo(ByVal ZipFile As String, _
                                nDate As String, _
                                nTime As String, _
                                ProjectName As String, _
                                Extrainfo As String, _
                                Optional ReadOnly As Boolean = True)

598  Dim lngVar1           As Long
599  Dim FileFound         As Boolean

601  Dim UZ                As cUnzip
602  Dim strLastBackupFile As String
    On Error Resume Next
604      If typBackupsInfo(0).ProjectFile <> "" Then
605          ReDim typBackupsInfo(0)
    End If
607      ProjectName = vbNullString
608      nDate = vbNullString
609      nTime = vbNullString
610      Extrainfo = vbNullString
611      If ReadOnly Then
612          Set UZ = New cUnzip
613  'Load From Zip Comment of  Backup FIle
614          GlobalComment = vbNullString
615          With UZ
616              .ZipFile = ZipFile
617              .ReadComment = True
618              .Directory
        End With
620          If LenB(GlobalComment) Then
621              For lngVar1 = 0 To UBound(Split(GlobalComment, vbLf)) - 1
622                  strLastBackupFile = Split(GlobalComment, vbLf)(lngVar1)
623                  If Left$(strLastBackupFile, 6) = "Date: " Then
624                      nDate = Split(strLastBackupFile, "Date: ")(1)
                End If
626                  If Left$(strLastBackupFile, 6) = "Time: " Then
627                      nTime = Split(strLastBackupFile, "Time: ")(1)
                End If
629                  If Left$(strLastBackupFile, 14) = "Project Path: " Then
630                      ProjectName = Split(strLastBackupFile, "Project Path: ")(1)
                End If
632                  If Left$(strLastBackupFile, 18) = "Proect File Name: " Then
633                      ProjectName = ProjectName & "\" & Split(strLastBackupFile, "Proect File Name: ")(1)
                End If
635              Next lngVar1
        End If
637      Else 'ReadOnly = FALSE/0
638          ReDim Preserve typBackupsInfo(UBound(typBackupsInfo) + 1)
639          With typBackupsInfo(UBound(typBackupsInfo))
640              .nDate = nDate
641              .nTime = nTime
642              .ProjectFile = ProjectName
643              .ZipFile = ZipFile
        End With
    End If
646  'Get If INFO Loaded (CACHE)
647      For lngVar1 = 0 To UBound(typBackupsInfo)
648          If ZipFile = typBackupsInfo(lngVar1).ZipFile Then
649              nDate = typBackupsInfo(lngVar1).nDate
650              nTime = typBackupsInfo(lngVar1).nTime
651              ProjectName = typBackupsInfo(lngVar1).ProjectFile
652              Exit For
        End If
654      Next lngVar1
655  'Load INFO From Zip Files
    

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
661  Timer2.Enabled = False

End Sub

Private Sub Label4_Click()

667      Text2.Text = "bas,cls,ctl,ctx,dca,ddf,dep,dob,dox,dsr,dsx,dws,frm,frx,log,oca,pag,pgx,res,tlb,vbg,vbl,vbp,vbr,vbw,vbz,wct"
668      WriteIniFile "Settings", "Backupfiles", Text2.Text

End Sub

Private Sub List1_MouseUp(Button As Integer, _
                          Shift As Integer, _
                          x As Single, _
                          Y As Single)

On Error Resume Next

679      If Button = 2 Then
680          If Split(List1.Text, "(")(0) <> "" Then
681              PopupMenu mnu2
        End If
    End If
684  err.Clear

End Sub

Public Sub LoadAll()

690  Dim A1 As Long

692      Combo1.Clear
693      Combo3.Clear
694      For A1 = 10 To 100 Step 10
695          Combo1.AddItem A1
696      Next A1
697      Combo1.AddItem "Unlimited"
    
699      Combo3.AddItem "Project Close"
700      Combo3.AddItem "Project Open"
701      Text1.Text = ReadIniFile("Settings", "ProjectsPath", "")
702      Text2.Text = ReadIniFile("Settings", "Backupfiles", "bas,cls,ctl,ctx,dca,ddf,dep,dob,dox,dsr,dsx,dws,frm,frx,log,oca,pag,pgx,res,tlb,vbg,vbl,vbp,vbr,vbw,vbz,wct")
703      Text3.Text = ReadIniFile("Settings", "BackupsPath", "")
704      Combo1.Text = ReadIniFile("Settings", "MaxBackups", "10")
    
706      Combo3.Text = ReadIniFile("Settings", "Backupon", "Project Close")
707      Check1.Value = CInt(ReadIniFile("Settings", "AutoBackup", "0"))

End Sub


Private Sub PopulateRestoreFiles()

714  Dim D1 As String
715  Dim T1 As String
716  Dim P1 As String
717  Dim E1 As String
718  Dim A1 As Long

On Error Resume Next

722      Combo2.Clear
723      With File1
724          .Pattern = "*.zip"
725          .Path = Text3.Text
726          .Refresh
727          For A1 = 0 To .ListCount - 1
728              GetSetZipBackupInfo .Path & "\" & .List(A1), D1, T1, P1, E1, True
729              If LCase$(P1) = LCase$(VBInstance.ActiveVBProject.Filename) Then
730                  If Not LenB(P1) = 0 Then
731                      Combo2.AddItem .List(A1)
                End If
            End If
734          Next A1
    End With 'File1
736      If Combo2.ListCount > 0 Then
737          Combo2.ListIndex = 0
    End If
739  err.Clear
End Sub

Private Sub refresh1_Click()

744      CountFiles

End Sub

Private Sub remautobkp_Click()

750      WriteIniFile "Backup", "Backup - " & VBInstance.ActiveVBProject.Name, "0"
751      MsgBox "Now AutoBackup Will Not backup this project Automatically!"
752      Connect.Hide

End Sub

Private Sub restore1_Click()

758      btn_restore.Visible = True
759      PopulateRestoreFiles
760      CheckButtons
761      Frame3.Visible = True
762      Frame3.ZOrder 0

End Sub

Private Sub RestoreButtonsEnabled(ByVal Var As Boolean)

768      btn_openfile.Enabled = Var
769      btn_saveas.Enabled = Var
770      btn_delfile.Enabled = Var
771      btn_restore.Enabled = Var

End Sub

Private Sub runproj1_Click()

777  Dim F As String

779  'APPROVED(Y ) List1 [Remove the space after the 'Y' in brackets and next run of Code Fixer will create the With Structure for you.
780  'APPROVED(Y ) List1 [Remove the space after the 'Y' in brackets and next run of Code Fixer will create the With Structure for you.
781      If LenB(List1.Text) Then
782          If Left$(Split(List1.Text, "(")(0), 1) = "\" Then
783              F = Text1.Text & Split(List1.Text, "(")(0)
784          Else
785              F = Split(List1.Text, "(")(0)
        End If
787          ShellExecute Me.hWnd, "open", F, "", "", 1
    End If

End Sub

Private Sub seeprojfiles1_Click()

794  Dim D1 As String
795  Dim T1 As String
796  Dim P1 As String
797  Dim E1 As String
798  Dim A1 As Long
799  Dim F  As String

801      btn_restore.Visible = False
802      Combo2.Clear
803      With File1
804          .Pattern = "*.zip"
805          .Path = Text3.Text
806          .Refresh
807          If Left$(Split(List1.Text, " (")(0), 1) = "\" Then
808              F = Text1.Text & Split(List1.Text, " (")(0)
809          Else
810              F = Split(List1.Text, " (")(0)
        End If
812          For A1 = 0 To .ListCount - 1
813              GetSetZipBackupInfo .Path & "\" & .List(A1), D1, T1, P1, E1, True
814              If LCase$(P1) = LCase$(F) Then
815                  If Not LenB(P1) = 0 Then
816                      Combo2.AddItem .List(A1)
                End If
            End If
819          Next A1
    End With 'File1
821      If Combo2.ListCount > 0 Then
822          Combo2.ListIndex = 0
    End If
824      CheckButtons
825      Frame3.Visible = True
826      Frame3.ZOrder 0

End Sub

Private Sub settings1_Click()

832      Frame1.Visible = True
833      Frame1.ZOrder 0

End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, _
                        Shift As Integer)

840      WriteIniFile "Settings", "ProjectsPath", Text1.Text

End Sub

Private Sub Text2_KeyUp(KeyCode As Integer, _
                        Shift As Integer)

847      WriteIniFile "Settings", "Backupfiles", Text2.Text

End Sub

Private Sub Text3_KeyUp(KeyCode As Integer, _
                        Shift As Integer)

860      WriteIniFile "Settings", "BackupsPath", Text3.Text

End Sub


Private Sub Timer2_Timer()

On Error Resume Next

869      If VBInstance.VBProjects.Count > 0 Then
870          If VBInstance.ActiveVBProject.Filename <> "" Then
871              Timer2.Enabled = False
872              If Backup = True Then
            
874              Else
            End If
876              Unload Me
        End If
    End If
879  err.Clear
End Sub

Public Sub WaitForProjectOpen()

884      Timer2.Enabled = True

End Sub


