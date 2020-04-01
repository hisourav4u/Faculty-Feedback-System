VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form FrmFeedback1 
   BackColor       =   &H00C0C0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Feed Back Form"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10230
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   10230
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox ComboCourse 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      ItemData        =   "FrmFeedback1.frx":0000
      Left            =   6480
      List            =   "FrmFeedback1.frx":000D
      TabIndex        =   16
      Text            =   "B.Tech"
      Top             =   840
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Need help for Subject??"
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
      Left            =   7920
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2040
      Width           =   2295
   End
   Begin VB.CommandButton cmdsubmit 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Submit && Go Next"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3600
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3720
      Width           =   3015
   End
   Begin VB.ComboBox combosem 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      ItemData        =   "FrmFeedback1.frx":002E
      Left            =   5040
      List            =   "FrmFeedback1.frx":004A
      TabIndex        =   7
      Text            =   "1"
      Top             =   840
      Width           =   615
   End
   Begin VB.ComboBox Combosession 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      ItemData        =   "FrmFeedback1.frx":0066
      Left            =   2400
      List            =   "FrmFeedback1.frx":0091
      TabIndex        =   6
      Text            =   "2018-2019"
      Top             =   2640
      Width           =   1575
   End
   Begin VB.ComboBox Combodept 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      ItemData        =   "FrmFeedback1.frx":0124
      Left            =   2400
      List            =   "FrmFeedback1.frx":014F
      TabIndex        =   5
      Text            =   "CSE"
      Top             =   840
      Width           =   1575
   End
   Begin VB.ComboBox Combosubname 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2400
      TabIndex        =   4
      Text            =   "Choose Subject Name"
      Top             =   2040
      Width           =   5415
   End
   Begin VB.ComboBox Combosubcode 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2400
      TabIndex        =   3
      Text            =   "Subject Code"
      Top             =   1560
      Width           =   2175
   End
   Begin VB.ComboBox ComboFaculty 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2400
      TabIndex        =   1
      Text            =   "Faculty name"
      Top             =   360
      Width           =   5775
   End
   Begin MSMask.MaskEdBox txtdate 
      Height          =   375
      Left            =   2400
      TabIndex        =   8
      Top             =   3120
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Course:"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   5760
      TabIndex        =   17
      Top             =   840
      Width           =   675
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date(DD/MM/YYYY)"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   240
      TabIndex        =   13
      Top             =   3240
      Width           =   1860
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Semester:"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   4080
      TabIndex        =   12
      Top             =   840
      Width           =   915
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Academic Session"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   240
      TabIndex        =   11
      Top             =   2640
      Width           =   1590
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Stream/ Department:"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   240
      TabIndex        =   10
      Top             =   960
      Width           =   1980
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Choose Subject Code:"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   240
      TabIndex        =   9
      Top             =   1560
      Width           =   1890
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Choose Subject Name:"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   240
      TabIndex        =   2
      Top             =   2040
      Width           =   1980
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Choose Faculty name: "
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1995
   End
End
Attribute VB_Name = "FrmFeedback1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdsubmit_Click()
frm_feedback.Combofaculty = Me.Combofaculty
frm_feedback.Combosubcode = Me.Combosubcode
frm_feedback.ComboSubname = Me.ComboSubname
frm_feedback.Combodept = Me.Combodept
frm_feedback.combosem = Me.combosem
frm_feedback.Combosession = Me.Combosession
frm_feedback.ComboCourse = Me.ComboCourse
frm_feedback.txtdate = Me.txtdate
Me.Hide
Unload Me
frm_feedback.Show
End Sub

Private Sub Combosubname_GotFocus()
''''' following code will not work,Do not uncomment this code''''''''''''''
''Dim rs As New ADODB.Recordset
''Dim sql As String, fnd As Boolean
''sql = "SELECT * FROM SUBJECT WHERE SUBCODE LIKE '" & Trim(Me.Combosubcode) & "'"
''If CON = False Then
''    connect_db
''    CON = True
''End If
''rs.Open sql, acon, adOpenStatic, adLockOptimistic
''
''Do While rs.EOF = False
'' Me.Combosubname = rs.Fields("SUBNAME")
'' fnd = True
'' rs.MoveNext
''Loop
''acon.Close
''CON = False
''If fnd = False Then
'' MsgBox "Subject Code is not in database", vbInformation, "Not Found"
'' 'Me.Combosubcode.Text = ""
'' 'Me.Combosubcode.SetFocus
''End If
End Sub

Private Sub Command1_Click()
FrmSubjectSearch.Show
End Sub

Private Sub Form_Activate()


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim rs As New ADODB.Recordset
sql = "SELECT * FROM EMPMASTER"
connect_db
CON = True
rs.Open sql, acon, adOpenStatic, adLockOptimistic

Do While rs.EOF = False
    Me.Combofaculty.AddItem rs.Fields("empname")
    rs.MoveNext
Loop
rs.Close
acon.Close
CON = False

'================================================================================
Dim RS2 As New ADODB.Recordset
SQL2 = "SELECT * FROM SUBJECT"
connect_db
CON = True
RS2.Open SQL2, acon, adOpenStatic, adLockOptimistic

Do While RS2.EOF = False
    Me.Combosubcode.AddItem RS2.Fields("SUBCODE")
    RS2.MoveNext
Loop
RS2.Close
acon.Close
CON = False

'===============================================================================
Dim RS3 As New ADODB.Recordset
SQL3 = "SELECT * FROM SUBJECT"
connect_db
CON = True
RS3.Open SQL3, acon, adOpenStatic, adLockOptimistic

Do While RS3.EOF = False
    Me.ComboSubname.AddItem RS3.Fields("SUBNAME")
    RS3.MoveNext
Loop
RS3.Close
acon.Close
CON = False

'''' calling txtdate_gotfocus procedure
Call txtdate_GotFocus
End Sub

Private Sub txtdate_GotFocus()
'''' fomation of date'''''''''
If Day(Date) <= 9 Then
   d = "0" & Day(Date)
Else
   d = Day(Date)
End If

If Month(Date) <= 9 Then
   m = "0" & Month(Date)
Else
   m = Month(Date)
End If
Me.txtdate = d & "/" & m & "/" & Year(Date)

End Sub
