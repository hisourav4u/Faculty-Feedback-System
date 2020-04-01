VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "Feed Back System"
   ClientHeight    =   3645
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5730
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIForm1.frx":0442
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   3  'Align Left
      BackColor       =   &H00C0E0FF&
      Height          =   12375
      Left            =   0
      ScaleHeight     =   12315
      ScaleWidth      =   1935
      TabIndex        =   0
      Top             =   0
      Width           =   1995
      Begin VB.CommandButton cmdadminmode 
         BackColor       =   &H00FF8080&
         Caption         =   "Admin Mode"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1440
         Width           =   1815
      End
      Begin VB.CommandButton CMDFEEDBACKFORM 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Feed Back Form"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   600
         Width           =   1815
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdadminmode_Click()
frm_login.Show
End Sub

Private Sub CMDFEEDBACKFORM_Click()
FrmFeedback1.Show
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
End
End Sub
