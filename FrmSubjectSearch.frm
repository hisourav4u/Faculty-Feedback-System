VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmSubjectSearch 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Subject Search"
   ClientHeight    =   8070
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8310
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8070
   ScaleWidth      =   8310
   Begin VB.TextBox Txtsem 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      MaxLength       =   2
      TabIndex        =   7
      Text            =   "%"
      Top             =   240
      Width           =   735
   End
   Begin VB.CommandButton CMDSUBMIT 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Submit Selected Item"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7440
      Width           =   3255
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4575
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   8070
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483624
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "SubCode"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "SubName"
         Object.Width           =   8819
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Stream"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Sem"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Course"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.TextBox TxtSubKeyWord 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   8055
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
      ItemData        =   "FrmSubjectSearch.frx":0000
      Left            =   2160
      List            =   "FrmSubjectSearch.frx":0028
      TabIndex        =   2
      Text            =   "BSH"
      Top             =   240
      Width           =   1575
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   -240
      X2              =   8280
      Y1              =   7320
      Y2              =   7320
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(Type ""%"" for All Subject for a particular semester)"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   225
      Left            =   2760
      TabIndex        =   10
      Top             =   840
      Width           =   4590
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "% means any SEM"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   225
      Left            =   5640
      TabIndex        =   9
      Top             =   240
      Width           =   1650
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   $"FrmSubjectSearch.frx":0065
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   705
      Left            =   240
      TabIndex        =   8
      Top             =   6360
      Width           =   7905
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Semester "
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
      Left            =   3840
      TabIndex        =   6
      Top             =   240
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Keyword For Subject"
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
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   2430
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Choose Stream/Dept"
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
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1875
   End
End
Attribute VB_Name = "FrmSubjectSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdsubmit_Click()
FrmFeedback1.Combosubcode = Me.ListView1.SelectedItem.Text
FrmFeedback1.ComboSubname = Me.ListView1.SelectedItem.SubItems(1)
'''FrmFeedback1.Combodept = Me.ListView1.SelectedItem.SubItems(2)
FrmFeedback1.combosem = Me.ListView1.SelectedItem.SubItems(3)
FrmFeedback1.ComboCourse = Me.ListView1.SelectedItem.SubItems(4)
Me.Hide
Unload Me
End Sub

Private Sub Form_Activate()
Me.Combodept = FrmFeedback1.Combodept
End Sub

Private Sub TxtSubKeyWord_KeyDown(KeyCode As Integer, Shift As Integer)
Dim rs As New ADODB.Recordset
Dim sql As String
sql = "SELECT * FROM SUBJECT WHERE STREAM LIKE '" & Trim(Me.Combodept) & "' AND SUBNAME LIKE '%" & Trim(Me.TxtSubKeyWord) & "%' AND SEM LIKE '" & Trim(Me.Txtsem) & "'"
If CON = False Then
  connect_db
  CON = True
End If
rs.Open sql, acon, adOpenStatic, adLockOptimistic
Me.ListView1.ListItems.Clear
Dim C As Double
Do While rs.EOF = False
    C = C + 1
    Me.ListView1.ListItems.Add
    Me.ListView1.ListItems(C).Text = rs.Fields("SUBCODE")
    Me.ListView1.ListItems(C).SubItems(1) = rs.Fields("SUBNAME")
    Me.ListView1.ListItems(C).SubItems(2) = rs.Fields("STREAM")
    Me.ListView1.ListItems(C).SubItems(3) = rs.Fields("SEM")
    Me.ListView1.ListItems(C).SubItems(4) = rs.Fields("COURSE")
    rs.MoveNext
Loop
acon.Close
CON = False
End Sub
