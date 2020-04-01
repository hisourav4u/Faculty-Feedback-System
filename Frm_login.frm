VERSION 5.00
Begin VB.Form frm_login 
   BackColor       =   &H8000000B&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Login"
   ClientHeight    =   4410
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8220
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   8220
   Begin VB.CommandButton cmdcancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5520
      TabIndex        =   5
      Top             =   3600
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Submit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   4
      Top             =   3600
      Width           =   2055
   End
   Begin VB.TextBox txtpass 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   12
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      IMEMode         =   3  'DISABLE
      Left            =   3960
      PasswordChar    =   "®"
      TabIndex        =   3
      Top             =   2760
      Width           =   3255
   End
   Begin VB.TextBox txtuserid 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   3960
      TabIndex        =   2
      Top             =   2040
      Width           =   3255
   End
   Begin VB.Image Image2 
      Height          =   1935
      Left            =   120
      Picture         =   "Frm_login.frx":0000
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   1695
      Left            =   -120
      Picture         =   "Frm_login.frx":2281
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8415
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2280
      TabIndex        =   1
      Top             =   2880
      Width           =   1170
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User Id"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2280
      TabIndex        =   0
      Top             =   2040
      Width           =   900
   End
End
Attribute VB_Name = "frm_login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdcancel_Click()
End
End Sub

Private Sub Command1_Click()
connect_db
Dim rs As New ADODB.Recordset
Dim sql As String
sql = "SELECT * FROM USER WHERE USERID LIKE '" & Trim(Me.txtuserid) & "' AND USERPASS LIKE '" & Trim(Me.txtpass) & _
    "'"
rs.Open sql, acon, adOpenStatic, adLockOptimistic

If rs.RecordCount > 0 Then
    rs.Close
    acon.Close
    
    FrmOption.Show
    Me.Hide
    Unload Me
Else
    MsgBox "Either UserId or Password is incorrect", vbCritical, "User Check Error"
    Me.txtuserid = ""
    Me.txtpass = ""
    Me.txtuserid.SetFocus
    rs.Close
    acon.Close
End If


End Sub

Private Sub Form_Unload(Cancel As Integer)
Me.Hide
Unload Me
End Sub
