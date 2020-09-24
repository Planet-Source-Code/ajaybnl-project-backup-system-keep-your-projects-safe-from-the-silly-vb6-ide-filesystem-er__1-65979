VERSION 5.00
Begin VB.Form frmShow 
   BorderStyle     =   5  'Sizable ToolWindow
   ClientHeight    =   1335
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   4095
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1335
   ScaleWidth      =   4095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFC0C0&
      Height          =   1335
      Left            =   0
      ScaleHeight     =   1275
      ScaleWidth      =   4035
      TabIndex        =   0
      Top             =   0
      Width           =   4095
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   3720
         Top             =   360
      End
      Begin VB.Shape Shape1 
         DrawMode        =   7  'Invert
         FillColor       =   &H00FFC0C0&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   375
         Top             =   960
         Visible         =   0   'False
         Width           =   3120
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Completed"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   600
         TabIndex        =   2
         Top             =   960
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Backup in Progress..."
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   360
         TabIndex        =   1
         Top             =   360
         Width           =   3180
      End
   End
End
Attribute VB_Name = "frmShow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, _
                                                    ByVal hWndInsertAfter As Long, _
                                                    ByVal x As Long, _
                                                    ByVal Y As Long, _
                                                    ByVal cx As Long, _
                                                    ByVal cy As Long, _
                                                    ByVal wFlags As Long) As Long

Private Sub Form_Load()

    With Me
        SetWindowPos .hWnd, -1, .Left / 15, .Top / 15, .Width / 15, .Height / 15, 0
        .Top = ((Screen.Height - .Height) / 2) / 2
        .Left = (Screen.Width - .Width) / 2
    End With 'Me
    Timer1.Interval = 5000

End Sub

Private Sub Timer1_Timer()

    Unload Me

End Sub


