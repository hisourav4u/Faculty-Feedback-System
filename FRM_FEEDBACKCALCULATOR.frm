VERSION 5.00
Begin VB.Form FRM_FEEDBACKCALCULATOR 
   BackColor       =   &H00FFC0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Feed Back Calculator( By Subject Code)"
   ClientHeight    =   3300
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7620
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   7620
   Begin VB.CommandButton CmdGenerateExcel 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Generate Excel File"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2640
      Width           =   2175
   End
   Begin VB.ComboBox Combosession 
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
      Height          =   345
      ItemData        =   "FRM_FEEDBACKCALCULATOR.frx":0000
      Left            =   2400
      List            =   "FRM_FEEDBACKCALCULATOR.frx":002B
      TabIndex        =   6
      Text            =   "2018-2019"
      Top             =   1800
      Width           =   2295
   End
   Begin VB.CommandButton cmdgeneratefeedback 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Generate Feed Back"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2640
      Width           =   2175
   End
   Begin VB.ComboBox Combosubcode 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2400
      TabIndex        =   3
      Text            =   "Choose Subject Code"
      Top             =   1200
      Width           =   2655
   End
   Begin VB.ComboBox Combofaculty 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2400
      TabIndex        =   1
      Text            =   "Choose Faculty name"
      Top             =   600
      Width           =   5175
   End
   Begin VB.Label Label3 
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
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Choose Subjct Code"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Choose Staff name"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   1695
   End
End
Attribute VB_Name = "FRM_FEEDBACKCALCULATOR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ques(11, 4) As Double
Public Sub array_ques_reset()
For i = 0 To 11
    For X = 0 To 4
      ques(i, X) = 0
    Next X
Next i
End Sub

Private Sub CmdGenerateExcel_Click()
Dim filename As String
filename = Trim(Me.Combofaculty)
filename = InputBox("Enter File Name Please", "Enter File name", filename)
If Trim(filename) = "" Then
    MsgBox "You have not given any filename", vbCritical, "Error"
    Exit Sub
End If

''''''Code for database connectivity and access''''''''''''''

Dim rs4 As New ADODB.Recordset
Dim sql As String
sql = "SELECT * FROM FEEDBACKGRAPH"
If CON = False Then
    connect_db
    CON = True
End If
rs4.Open sql, acon, adOpenStatic, adLockOptimistic

''' CODE FOR EXCEL FILE OPERATION''''''''''''''''''''''''''''''''''''''''''''''
Dim xl As New Excel.Application
Dim xlsheet As Excel.Worksheet
Dim xlwbook As Excel.Workbook


FileCopy App.Path & "\book1.xls", App.Path & "\Outputexcel\" & Trim(filename) & ".xls"
Set xlwbook = xl.Workbooks.Open(App.Path & "\Outputexcel\" & Trim(filename) & ".xls")
Set xlsheet = xlwbook.Sheets.Item(1)

Do While rs4.EOF = False
    xlsheet.Range("B6") = rs4.Fields("EMPNAME")
    xlsheet.Range("B7") = rs4.Fields("SUBCODE")
    xlsheet.Range("B8") = rs4.Fields("SUBNAME")
    xlsheet.Range("B9") = rs4.Fields("DEPT")
    xlsheet.Range("B9") = rs4.Fields("DEPT")
    xlsheet.Range("C9") = "Semester : " & rs4.Fields("SEM")
    xlsheet.Range("C10") = "Course : " & rs4.Fields("COURSE")
    xlsheet.Range("B10") = rs4.Fields("ACADEMICSESSION")
    xlsheet.Range("B11") = rs4.Fields("DATE1")
    
    xlsheet.Range("C19") = Val(rs4.Fields("OUTSTANDING"))
    xlsheet.Range("C20") = Val(rs4.Fields("EXCELLENT"))
    xlsheet.Range("C21") = Val(rs4.Fields("GOOD"))
    xlsheet.Range("C22") = Val(rs4.Fields("AVERAGE"))
    xlsheet.Range("C23") = Val(rs4.Fields("POOR"))
    
    
    ''''''Trnsafer of value from variable ques() to excel sheet[For Outstanding, Excellent,good,avearge,poor Value]'''''''''
    xlsheet.Range("J16") = ques(0, 0)
    xlsheet.Range("J18") = ques(1, 0)
    xlsheet.Range("J20") = ques(2, 0)
    xlsheet.Range("J22") = ques(3, 0)
    xlsheet.Range("J24") = ques(4, 0)
    xlsheet.Range("J26") = ques(5, 0)
    xlsheet.Range("J28") = ques(6, 0)
    xlsheet.Range("J30") = ques(7, 0)
    xlsheet.Range("J32") = ques(8, 0)
    xlsheet.Range("J34") = ques(9, 0)
    xlsheet.Range("J36") = ques(10, 0)
    xlsheet.Range("J38") = ques(11, 0)
    
    xlsheet.Range("K16") = ques(0, 1)
    xlsheet.Range("K18") = ques(1, 1)
    xlsheet.Range("K20") = ques(2, 1)
    xlsheet.Range("K22") = ques(3, 1)
    xlsheet.Range("K24") = ques(4, 1)
    xlsheet.Range("K26") = ques(5, 1)
    xlsheet.Range("K28") = ques(6, 1)
    xlsheet.Range("K30") = ques(7, 1)
    xlsheet.Range("K32") = ques(8, 1)
    xlsheet.Range("K34") = ques(9, 1)
    xlsheet.Range("K36") = ques(10, 1)
    xlsheet.Range("K38") = ques(11, 1)
    
    xlsheet.Range("L16") = ques(0, 2)
    xlsheet.Range("L18") = ques(1, 2)
    xlsheet.Range("L20") = ques(2, 2)
    xlsheet.Range("L22") = ques(3, 2)
    xlsheet.Range("L24") = ques(4, 2)
    xlsheet.Range("L26") = ques(5, 2)
    xlsheet.Range("L28") = ques(6, 2)
    xlsheet.Range("L30") = ques(7, 2)
    xlsheet.Range("L32") = ques(8, 2)
    xlsheet.Range("L34") = ques(9, 2)
    xlsheet.Range("L36") = ques(10, 2)
    xlsheet.Range("L38") = ques(11, 2)
    
    xlsheet.Range("M16") = ques(0, 3)
    xlsheet.Range("M18") = ques(1, 3)
    xlsheet.Range("M20") = ques(2, 3)
    xlsheet.Range("M22") = ques(3, 3)
    xlsheet.Range("M24") = ques(4, 3)
    xlsheet.Range("M26") = ques(5, 3)
    xlsheet.Range("M28") = ques(6, 3)
    xlsheet.Range("M30") = ques(7, 3)
    xlsheet.Range("M32") = ques(8, 3)
    xlsheet.Range("M34") = ques(9, 3)
    xlsheet.Range("M36") = ques(10, 3)
    xlsheet.Range("M38") = ques(11, 3)
    
    xlsheet.Range("N16") = ques(0, 4)
    xlsheet.Range("N18") = ques(1, 4)
    xlsheet.Range("N20") = ques(2, 4)
    xlsheet.Range("N22") = ques(3, 4)
    xlsheet.Range("N24") = ques(4, 4)
    xlsheet.Range("N26") = ques(5, 4)
    xlsheet.Range("N28") = ques(6, 4)
    xlsheet.Range("N30") = ques(7, 4)
    xlsheet.Range("N32") = ques(8, 4)
    xlsheet.Range("N34") = ques(9, 4)
    xlsheet.Range("N36") = ques(10, 4)
    xlsheet.Range("N38") = ques(11, 4)
    
    
    rs4.MoveNext
Loop
xl.ActiveWorkbook.Save
xl.ActiveWorkbook.Close True, App.Path & "\Outputexcel\" & Trim(filename) & ".xls"
xl.Quit
Set xlwbook = Nothing: Set xlsheet = Nothing
Set xl = Nothing

acon.Close
CON = False
MsgBox "Excel File has been generated", vbInformation, "Information"

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''Now disable the Generate Excel File Button''''''''''''''''''''''''
Me.CmdGenerateExcel.Enabled = False
'''''''Reset Arry of ques'''''''''''''''''
Call array_ques_reset
End Sub

Private Sub cmdgeneratefeedback_Click()
Dim TOTOUTSTANDING As Double, TOTEXCELLENT As Double, RCOUNT As Long, TOTGOOD As Double, TOTAVERAGE As Double, TOTPOOR As Double, ACADEMICSESSION As String

Dim EMPNAME As String, SUBCODE As String, course As String, SUBNAME As String, fnd As Boolean
Dim Sem As String, dt As Date, dept As String

CON = False
connect_db
acon.Execute ("DELETE * FROM FEEDBACKGRAPH")
acon.Close
CON = False

Dim rs As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset
Dim sql As String
Dim SQL2 As String
sql = "SELECT * FROM EMP WHERE EMPNAME LIKE '" & Trim(Me.Combofaculty) & "' AND SUBCODE LIKE '" & Trim(Me.Combosubcode) & "' AND ACADEMICSESSION LIKE '" & Trim(Me.Combosession) & "'"
connect_db
rs.Open sql, acon, adOpenStatic, adLockOptimistic
Do While rs.EOF = False
    TOTEXCELLENT = TOTEXCELLENT + rs.Fields("EXCELLENT")
    TOTGOOD = TOTGOOD + rs.Fields("GOOD")
    TOTAVERAGE = TOTAVERAGE + rs.Fields("AVERAGE")
    TOTPOOR = TOTPOOR + rs.Fields("POOR")
    TOTOUTSTANDING = TOTOUTSTANDING + rs.Fields("OUTSTANDING")
    EMPNAME = rs.Fields("EMPNAME")
    SUBCODE = rs.Fields("SUBCODE")
    SUBNAME = rs.Fields("SUBNAME")
    ACADEMICSESSION = rs.Fields("ACADEMICSESSION")
    Sem = rs.Fields("SEM")
    dt = rs.Fields("DATE1")
    dept = rs.Fields("DEPT")
    course = rs.Fields("COURSE")
    
    
    '''''''''''''Checking individual Queries feedback''''''''''''''''''
    If Trim(rs.Fields("ques1")) = "O" Then
        ques(0, 0) = ques(0, 0) + 1
    ElseIf Trim(rs.Fields("ques1")) = "E" Then
        ques(0, 1) = ques(0, 1) + 1
    ElseIf Trim(rs.Fields("ques1")) = "G" Then
        ques(0, 2) = ques(0, 2) + 1
    ElseIf Trim(rs.Fields("ques1")) = "A" Then
        ques(0, 3) = ques(0, 3) + 1
    ElseIf Trim(rs.Fields("ques1")) = "P" Then
        ques(0, 4) = ques(0, 4) + 1
    End If
        
    If Trim(rs.Fields("ques2")) = "O" Then
        ques(1, 0) = ques(1, 0) + 1
    ElseIf Trim(rs.Fields("ques2")) = "E" Then
        ques(1, 1) = ques(1, 1) + 1
    ElseIf Trim(rs.Fields("ques2")) = "G" Then
        ques(1, 2) = ques(1, 2) + 1
    ElseIf Trim(rs.Fields("ques2")) = "A" Then
        ques(1, 3) = ques(1, 3) + 1
    ElseIf Trim(rs.Fields("ques2")) = "P" Then
        ques(1, 4) = ques(1, 4) + 1
    End If
    
    If Trim(rs.Fields("ques3")) = "O" Then
        ques(2, 0) = ques(2, 0) + 1
    ElseIf Trim(rs.Fields("ques3")) = "E" Then
        ques(2, 1) = ques(2, 1) + 1
    ElseIf Trim(rs.Fields("ques3")) = "G" Then
        ques(2, 2) = ques(2, 2) + 1
    ElseIf Trim(rs.Fields("ques3")) = "A" Then
        ques(2, 3) = ques(2, 3) + 1
    ElseIf Trim(rs.Fields("ques3")) = "P" Then
        ques(2, 4) = ques(2, 4) + 1
    End If
    
    If Trim(rs.Fields("ques4")) = "O" Then
        ques(3, 0) = ques(3, 0) + 1
    ElseIf Trim(rs.Fields("ques4")) = "E" Then
        ques(3, 1) = ques(3, 1) + 1
    ElseIf Trim(rs.Fields("ques4")) = "G" Then
        ques(3, 2) = ques(3, 2) + 1
    ElseIf Trim(rs.Fields("ques4")) = "A" Then
        ques(3, 3) = ques(3, 3) + 1
    ElseIf Trim(rs.Fields("ques4")) = "P" Then
        ques(3, 4) = ques(3, 4) + 1
    End If
    
    If Trim(rs.Fields("ques5")) = "O" Then
        ques(4, 0) = ques(4, 0) + 1
    ElseIf Trim(rs.Fields("ques5")) = "E" Then
        ques(4, 1) = ques(4, 1) + 1
    ElseIf Trim(rs.Fields("ques5")) = "G" Then
        ques(4, 2) = ques(4, 2) + 1
    ElseIf Trim(rs.Fields("ques5")) = "A" Then
        ques(4, 3) = ques(4, 3) + 1
    ElseIf Trim(rs.Fields("ques5")) = "P" Then
        ques(4, 4) = ques(4, 4) + 1
    End If
    
    If Trim(rs.Fields("ques6")) = "O" Then
        ques(5, 0) = ques(5, 0) + 1
    ElseIf Trim(rs.Fields("ques6")) = "E" Then
        ques(5, 1) = ques(5, 1) + 1
    ElseIf Trim(rs.Fields("ques6")) = "G" Then
        ques(5, 2) = ques(5, 2) + 1
    ElseIf Trim(rs.Fields("ques6")) = "A" Then
        ques(5, 3) = ques(5, 3) + 1
    ElseIf Trim(rs.Fields("ques6")) = "P" Then
        ques(5, 4) = ques(5, 4) + 1
    End If
    
    If Trim(rs.Fields("ques7")) = "O" Then
        ques(6, 0) = ques(6, 0) + 1
    ElseIf Trim(rs.Fields("ques7")) = "E" Then
        ques(6, 1) = ques(6, 1) + 1
    ElseIf Trim(rs.Fields("ques7")) = "G" Then
        ques(6, 2) = ques(6, 2) + 1
    ElseIf Trim(rs.Fields("ques7")) = "A" Then
        ques(6, 3) = ques(6, 3) + 1
    ElseIf Trim(rs.Fields("ques7")) = "P" Then
        ques(6, 4) = ques(6, 4) + 1
    End If
    
    If Trim(rs.Fields("ques8")) = "O" Then
        ques(7, 0) = ques(7, 0) + 1
    ElseIf Trim(rs.Fields("ques8")) = "E" Then
        ques(7, 1) = ques(7, 1) + 1
    ElseIf Trim(rs.Fields("ques8")) = "G" Then
        ques(7, 2) = ques(7, 2) + 1
    ElseIf Trim(rs.Fields("ques8")) = "A" Then
        ques(7, 3) = ques(7, 3) + 1
    ElseIf Trim(rs.Fields("ques8")) = "P" Then
        ques(7, 4) = ques(7, 4) + 1
    End If
    
    If Trim(rs.Fields("ques9")) = "O" Then
        ques(8, 0) = ques(8, 0) + 1
    ElseIf Trim(rs.Fields("ques9")) = "E" Then
        ques(8, 1) = ques(8, 1) + 1
    ElseIf Trim(rs.Fields("ques9")) = "G" Then
        ques(8, 2) = ques(8, 2) + 1
    ElseIf Trim(rs.Fields("ques9")) = "A" Then
        ques(8, 3) = ques(8, 3) + 1
    ElseIf Trim(rs.Fields("ques9")) = "P" Then
        ques(8, 4) = ques(8, 4) + 1
    End If
    
    If Trim(rs.Fields("ques10")) = "O" Then
        ques(9, 0) = ques(9, 0) + 1
    ElseIf Trim(rs.Fields("ques10")) = "E" Then
        ques(9, 1) = ques(9, 1) + 1
    ElseIf Trim(rs.Fields("ques10")) = "G" Then
        ques(9, 2) = ques(9, 2) + 1
    ElseIf Trim(rs.Fields("ques10")) = "A" Then
        ques(9, 3) = ques(9, 3) + 1
    ElseIf Trim(rs.Fields("ques10")) = "P" Then
        ques(9, 4) = ques(9, 4) + 1
    End If
    
    If Trim(rs.Fields("ques11")) = "O" Then
        ques(10, 0) = ques(10, 0) + 1
    ElseIf Trim(rs.Fields("ques11")) = "E" Then
        ques(10, 1) = ques(10, 1) + 1
    ElseIf Trim(rs.Fields("ques11")) = "G" Then
        ques(10, 2) = ques(10, 2) + 1
    ElseIf Trim(rs.Fields("ques11")) = "A" Then
        ques(10, 3) = ques(10, 3) + 1
    ElseIf Trim(rs.Fields("ques11")) = "P" Then
        ques(10, 4) = ques(10, 4) + 1
    End If
    
    If Trim(rs.Fields("ques12")) = "O" Then
        ques(11, 0) = ques(11, 0) + 1
    ElseIf Trim(rs.Fields("ques12")) = "E" Then
        ques(11, 1) = ques(11, 1) + 1
    ElseIf Trim(rs.Fields("ques12")) = "G" Then
        ques(11, 2) = ques(11, 2) + 1
    ElseIf Trim(rs.Fields("ques12")) = "A" Then
        ques(11, 3) = ques(11, 3) + 1
    ElseIf Trim(rs.Fields("ques12")) = "P" Then
        ques(11, 4) = ques(11, 4) + 1
    End If
    
    
    
    fnd = True
    rs.MoveNext
Loop
''if there is no record''''''''''''''''''''
If fnd = False Then
    MsgBox "No record"
    acon.Close
    CON = False
    Exit Sub
End If
''''''''''''''''''''''''''''''''''''''''''''''

RCOUNT = rs.RecordCount

rs.Close
SQL2 = "SELECT * FROM FEEDBACKGRAPH"
RS2.Open SQL2, acon, adOpenStatic, adLockOptimistic

RS2.AddNew
RS2.Fields("EMPNAME") = EMPNAME
RS2.Fields("SUBCODE") = SUBCODE
RS2.Fields("SUBNAME") = SUBNAME
RS2.Fields("DEPT") = dept
RS2.Fields("SEM") = Sem
RS2.Fields("DATE1") = Format(dt, "DD/MM/YYYY")
RS2.Fields("COURSE") = course

RS2.Fields("ACADEMICSESSION") = ACADEMICSESSION
RS2.Fields("OUTSTANDING") = TOTOUTSTANDING
RS2.Fields("EXCELLENT") = TOTEXCELLENT
RS2.Fields("GOOD") = TOTGOOD
RS2.Fields("AVERAGE") = TOTAVERAGE
RS2.Fields("POOR") = TOTPOOR


RS2.Update
MsgBox "Feedback Report has been updated" & Chr(13) & "Please see report in database", vbInformation, "Updated"

RS2.Close
acon.Close
CON = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''now enable the Generate Excel File Button''''''''''''''''
Me.CmdGenerateExcel.Enabled = True
End Sub

Private Sub Form_Activate()
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
'================================================================
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

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''Generate Excel Fiel Button will be disabled at first in Form Load''''''''''''''
Me.CmdGenerateExcel.Enabled = False
End Sub

