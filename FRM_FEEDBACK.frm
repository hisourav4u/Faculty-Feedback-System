VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frm_feedback 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Student Feed Back"
   ClientHeight    =   11100
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15720
   Icon            =   "FRM_FEEDBACK.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   11100
   ScaleWidth      =   15720
   Begin VB.TextBox txtoutstanding 
      Alignment       =   2  'Center
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
      Left            =   10320
      Locked          =   -1  'True
      TabIndex        =   103
      Text            =   "0"
      Top             =   9240
      Width           =   855
   End
   Begin VB.CheckBox Check60 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10560
      TabIndex        =   102
      Top             =   8640
      Width           =   375
   End
   Begin VB.CheckBox Check59 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10560
      TabIndex        =   101
      Top             =   8160
      Width           =   375
   End
   Begin VB.CheckBox Check58 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10560
      TabIndex        =   100
      Top             =   7680
      Width           =   375
   End
   Begin VB.CheckBox Check57 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10560
      TabIndex        =   99
      Top             =   7200
      Width           =   375
   End
   Begin VB.CheckBox Check56 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10560
      TabIndex        =   98
      Top             =   6720
      Width           =   375
   End
   Begin VB.CheckBox Check55 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10560
      TabIndex        =   97
      Top             =   6240
      Width           =   375
   End
   Begin VB.CheckBox Check54 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10560
      TabIndex        =   96
      Top             =   5760
      Width           =   375
   End
   Begin VB.CheckBox Check53 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10560
      TabIndex        =   95
      Top             =   5280
      Width           =   375
   End
   Begin VB.CheckBox Check52 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10560
      TabIndex        =   94
      Top             =   4800
      Width           =   375
   End
   Begin VB.CheckBox Check51 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10560
      TabIndex        =   93
      Top             =   4200
      Width           =   375
   End
   Begin VB.CheckBox Check50 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10560
      TabIndex        =   92
      Top             =   3720
      Width           =   375
   End
   Begin VB.CheckBox Check49 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10560
      TabIndex        =   91
      Top             =   3240
      Width           =   375
   End
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
      ItemData        =   "FRM_FEEDBACK.frx":0442
      Left            =   10800
      List            =   "FRM_FEEDBACK.frx":044F
      Locked          =   -1  'True
      TabIndex        =   88
      Text            =   "B.Tech"
      Top             =   1320
      Width           =   2175
   End
   Begin VB.ComboBox combodept 
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
      ItemData        =   "FRM_FEEDBACK.frx":0470
      Left            =   8400
      List            =   "FRM_FEEDBACK.frx":0498
      Locked          =   -1  'True
      TabIndex        =   87
      Text            =   "CSE"
      Top             =   1320
      Width           =   1575
   End
   Begin VB.ComboBox Combosem 
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
      ItemData        =   "FRM_FEEDBACK.frx":04D5
      Left            =   12720
      List            =   "FRM_FEEDBACK.frx":04F1
      Locked          =   -1  'True
      TabIndex        =   85
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
      ItemData        =   "FRM_FEEDBACK.frx":050D
      Left            =   10560
      List            =   "FRM_FEEDBACK.frx":0538
      Locked          =   -1  'True
      TabIndex        =   83
      Text            =   "2018-2019"
      Top             =   840
      Width           =   1575
   End
   Begin MSMask.MaskEdBox txtdate 
      Height          =   375
      Left            =   7680
      TabIndex        =   81
      Top             =   840
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   10
      Format          =   "DD/MM/YYYY"
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.CommandButton cmdclose 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10440
      TabIndex        =   79
      Top             =   9960
      Width           =   1695
   End
   Begin VB.CommandButton CMDREST 
      Caption         =   "&Re-Set"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8400
      TabIndex        =   78
      Top             =   9960
      Width           =   1695
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6360
      TabIndex        =   77
      Top             =   9960
      Width           =   1695
   End
   Begin VB.TextBox txtpoor 
      Alignment       =   2  'Center
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
      Left            =   14760
      Locked          =   -1  'True
      TabIndex        =   75
      Text            =   "0"
      Top             =   9240
      Width           =   855
   End
   Begin VB.CheckBox Check48 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   15000
      TabIndex        =   74
      Top             =   8640
      Width           =   375
   End
   Begin VB.CheckBox Check47 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   15000
      TabIndex        =   73
      Top             =   8160
      Width           =   375
   End
   Begin VB.CheckBox Check46 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   15000
      TabIndex        =   72
      Top             =   7680
      Width           =   375
   End
   Begin VB.CheckBox Check45 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   15000
      TabIndex        =   71
      Top             =   7200
      Width           =   375
   End
   Begin VB.CheckBox Check44 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   15000
      TabIndex        =   70
      Top             =   6720
      Width           =   375
   End
   Begin VB.CheckBox Check43 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   15000
      TabIndex        =   69
      Top             =   6240
      Width           =   375
   End
   Begin VB.CheckBox Check42 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   15000
      TabIndex        =   68
      Top             =   5760
      Width           =   375
   End
   Begin VB.CheckBox Check41 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   15000
      TabIndex        =   67
      Top             =   5280
      Width           =   375
   End
   Begin VB.CheckBox Check40 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   15000
      TabIndex        =   66
      Top             =   4800
      Width           =   375
   End
   Begin VB.CheckBox Check39 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   15000
      TabIndex        =   65
      Top             =   4200
      Width           =   375
   End
   Begin VB.CheckBox Check38 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   15000
      TabIndex        =   64
      Top             =   3720
      Width           =   375
   End
   Begin VB.CheckBox Check37 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   15000
      TabIndex        =   63
      Top             =   3240
      Width           =   375
   End
   Begin VB.TextBox txtaverage 
      Alignment       =   2  'Center
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
      Left            =   13680
      Locked          =   -1  'True
      TabIndex        =   62
      Text            =   "0"
      Top             =   9240
      Width           =   855
   End
   Begin VB.CheckBox Check36 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   13920
      TabIndex        =   61
      Top             =   8640
      Width           =   375
   End
   Begin VB.CheckBox Check35 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   13920
      TabIndex        =   60
      Top             =   8160
      Width           =   375
   End
   Begin VB.CheckBox Check34 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   13920
      TabIndex        =   59
      Top             =   7680
      Width           =   375
   End
   Begin VB.CheckBox Check33 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   13920
      TabIndex        =   58
      Top             =   7200
      Width           =   375
   End
   Begin VB.CheckBox Check32 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   13920
      TabIndex        =   57
      Top             =   6720
      Width           =   375
   End
   Begin VB.CheckBox Check31 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   13920
      TabIndex        =   56
      Top             =   6240
      Width           =   375
   End
   Begin VB.CheckBox Check30 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   13920
      TabIndex        =   55
      Top             =   5760
      Width           =   375
   End
   Begin VB.CheckBox Check29 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   13920
      TabIndex        =   54
      Top             =   5280
      Width           =   375
   End
   Begin VB.CheckBox Check28 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   13920
      TabIndex        =   53
      Top             =   4800
      Width           =   375
   End
   Begin VB.CheckBox Check27 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   13920
      TabIndex        =   52
      Top             =   4200
      Width           =   375
   End
   Begin VB.CheckBox Check24 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12840
      TabIndex        =   51
      Top             =   8640
      Width           =   375
   End
   Begin VB.CheckBox Check26 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   13920
      TabIndex        =   50
      Top             =   3720
      Width           =   375
   End
   Begin VB.CheckBox Check25 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   13920
      TabIndex        =   49
      Top             =   3240
      Width           =   375
   End
   Begin VB.TextBox txtgood 
      Alignment       =   2  'Center
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
      Left            =   12600
      Locked          =   -1  'True
      TabIndex        =   48
      Text            =   "0"
      Top             =   9240
      Width           =   855
   End
   Begin VB.CheckBox Check23 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12840
      TabIndex        =   47
      Top             =   8160
      Width           =   375
   End
   Begin VB.CheckBox Check22 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12840
      TabIndex        =   46
      Top             =   7680
      Width           =   375
   End
   Begin VB.CheckBox Check21 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12840
      TabIndex        =   45
      Top             =   7200
      Width           =   375
   End
   Begin VB.CheckBox Check20 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12840
      TabIndex        =   44
      Top             =   6720
      Width           =   375
   End
   Begin VB.CheckBox Check19 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12840
      TabIndex        =   43
      Top             =   6240
      Width           =   375
   End
   Begin VB.CheckBox Check18 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12840
      TabIndex        =   42
      Top             =   5760
      Width           =   375
   End
   Begin VB.CheckBox Check17 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12840
      TabIndex        =   41
      Top             =   5280
      Width           =   375
   End
   Begin VB.CheckBox Check16 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12840
      TabIndex        =   40
      Top             =   4800
      Width           =   375
   End
   Begin VB.CheckBox Check15 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12840
      TabIndex        =   39
      Top             =   4200
      Width           =   375
   End
   Begin VB.CheckBox Check14 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12840
      TabIndex        =   38
      Top             =   3720
      Width           =   375
   End
   Begin VB.CheckBox Check13 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12840
      TabIndex        =   37
      Top             =   3240
      Width           =   375
   End
   Begin VB.TextBox txtexcellent 
      Alignment       =   2  'Center
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
      Left            =   11400
      Locked          =   -1  'True
      TabIndex        =   36
      Text            =   "0"
      Top             =   9240
      Width           =   855
   End
   Begin VB.CheckBox Check12 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11640
      TabIndex        =   35
      Top             =   8640
      Width           =   375
   End
   Begin VB.CheckBox Check11 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11640
      TabIndex        =   34
      Top             =   8160
      Width           =   375
   End
   Begin VB.CheckBox Check10 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11640
      TabIndex        =   33
      Top             =   7680
      Width           =   375
   End
   Begin VB.CheckBox Check9 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11640
      TabIndex        =   32
      Top             =   7200
      Width           =   375
   End
   Begin VB.CheckBox Check8 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11640
      TabIndex        =   31
      Top             =   6720
      Width           =   375
   End
   Begin VB.CheckBox Check7 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11640
      TabIndex        =   30
      Top             =   6240
      Width           =   375
   End
   Begin VB.CheckBox Check6 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11640
      TabIndex        =   29
      Top             =   5760
      Width           =   375
   End
   Begin VB.CheckBox Check5 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11640
      TabIndex        =   28
      Top             =   5280
      Width           =   375
   End
   Begin VB.CheckBox Check4 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   11640
      TabIndex        =   27
      Top             =   4800
      Width           =   375
   End
   Begin VB.CheckBox Check3 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   11640
      TabIndex        =   26
      Top             =   4200
      Width           =   375
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11640
      TabIndex        =   25
      Top             =   3720
      Width           =   375
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11640
      TabIndex        =   24
      Top             =   3240
      Width           =   375
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
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   12
      Text            =   "Choose Subject Code "
      Top             =   1800
      Width           =   4695
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
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   11
      Text            =   "Subject Code"
      Top             =   1320
      Width           =   2655
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
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   8
      Text            =   "Faculty name"
      Top             =   840
      Width           =   4695
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Outstanding"
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
      Left            =   10080
      TabIndex        =   90
      Top             =   2880
      Width           =   1125
   End
   Begin VB.Label Label27 
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
      Left            =   10080
      TabIndex        =   89
      Top             =   1440
      Width           =   675
   End
   Begin VB.Label Label26 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Stream/Dept."
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
      Left            =   7080
      TabIndex        =   86
      Top             =   1440
      Width           =   1245
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sem"
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
      Left            =   12240
      TabIndex        =   84
      Top             =   960
      Width           =   390
   End
   Begin VB.Label Label24 
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
      Left            =   8880
      TabIndex        =   82
      Top             =   960
      Width           =   1710
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date :"
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
      Left            =   7080
      TabIndex        =   80
      Top             =   960
      Width           =   525
   End
   Begin VB.Line Line16 
      X1              =   0
      X2              =   15720
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Votes"
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
      TabIndex        =   76
      Top             =   9240
      Width           =   1005
   End
   Begin VB.Line Line15 
      X1              =   0
      X2              =   15720
      Y1              =   9720
      Y2              =   9720
   End
   Begin VB.Line Line14 
      X1              =   0
      X2              =   15720
      Y1              =   9000
      Y2              =   9000
   End
   Begin VB.Line Line13 
      X1              =   0
      X2              =   15720
      Y1              =   8520
      Y2              =   8520
   End
   Begin VB.Line Line12 
      X1              =   0
      X2              =   15720
      Y1              =   8040
      Y2              =   8040
   End
   Begin VB.Line Line11 
      X1              =   0
      X2              =   15720
      Y1              =   7560
      Y2              =   7560
   End
   Begin VB.Line Line10 
      X1              =   0
      X2              =   15720
      Y1              =   7080
      Y2              =   7080
   End
   Begin VB.Line Line9 
      X1              =   0
      X2              =   15720
      Y1              =   6600
      Y2              =   6600
   End
   Begin VB.Line Line8 
      X1              =   0
      X2              =   15720
      Y1              =   6120
      Y2              =   6120
   End
   Begin VB.Line Line7 
      X1              =   0
      X2              =   15720
      Y1              =   5640
      Y2              =   5640
   End
   Begin VB.Line Line6 
      X1              =   0
      X2              =   15720
      Y1              =   5160
      Y2              =   5160
   End
   Begin VB.Line Line5 
      X1              =   0
      X2              =   15720
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Line Line4 
      X1              =   0
      X2              =   15720
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   15720
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Poor"
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
      Left            =   14880
      TabIndex        =   23
      Top             =   2880
      Width           =   420
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Average"
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
      Left            =   13680
      TabIndex        =   22
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Good"
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
      Left            =   12720
      TabIndex        =   21
      Top             =   2880
      Width           =   450
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Excellent"
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
      Left            =   11400
      TabIndex        =   20
      Top             =   2880
      Width           =   840
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "12. Help from Handouts/e-material given by the teacher Or Help from tutorials "
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
      TabIndex        =   19
      Top             =   8520
      Width           =   10665
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "11. Help From Quiz given by the teacher"
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
      TabIndex        =   18
      Top             =   8040
      Width           =   10665
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "10.Use of teaching aids like Projector/ Other aids etc."
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
      TabIndex        =   17
      Top             =   7560
      Width           =   10665
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "9.Help from assignments given by the teacher"
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
      TabIndex        =   16
      Top             =   7080
      Width           =   10665
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "8. Accessibility of the teacher outside the class room"
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
      TabIndex        =   15
      Top             =   6600
      Width           =   10665
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "7. Encouraging questions by the teacher in class "
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
      TabIndex        =   14
      Top             =   6120
      Width           =   10665
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "6. Support received from class room teaching in facing university exam"
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
      Top             =   5640
      Width           =   10665
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Choose Subject Name :"
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
      Left            =   0
      TabIndex        =   10
      Top             =   1800
      Width           =   2025
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Choose Subject Code :"
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
      Left            =   0
      TabIndex        =   9
      Top             =   1320
      Width           =   1935
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
      Left            =   0
      TabIndex        =   7
      Top             =   840
      Width           =   1995
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   0
      X2              =   15720
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "5.Coverage of syllabus by the teacher"
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
      TabIndex        =   6
      Top             =   5160
      Width           =   10665
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "4.Communication Skill of the teacher"
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
      TabIndex        =   5
      Top             =   4680
      Width           =   10545
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "3. Clarity of presentation of teacher for topics covered in the class"
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
      TabIndex        =   4
      Top             =   4200
      Width           =   10425
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "2. Ability of the teacher to answer all questions and queries"
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
      TabIndex        =   3
      Top             =   3720
      Width           =   10545
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "1. Interest generated by the teacher on the subject"
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
      Top             =   3240
      Width           =   10425
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   0
      X2              =   15720
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Student's Feed Back Form"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6240
      TabIndex        =   1
      Top             =   360
      Width           =   2865
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Techno International-Batanagar"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      TabIndex        =   0
      Top             =   0
      Width           =   4710
   End
End
Attribute VB_Name = "frm_feedback"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ques(11) As String

Private Sub Check1_Click()
If Me.Check1 = vbChecked Then
    Me.txtexcellent = Val(Me.txtexcellent) + 1
    Me.Check13 = 0
    Me.Check25 = 0
    Me.Check37 = 0
    Me.Check49 = 0
    ques(0) = "E"
Else
    Me.txtexcellent = Val(Me.txtexcellent) - 1
End If

If Val(Me.txtexcellent) < 1 Then Me.txtexcellent = 0
End Sub

Private Sub Check10_Click()
If Me.Check10 = vbChecked Then
    Me.txtexcellent = Val(Me.txtexcellent) + 1
    Me.Check22 = 0
    Me.Check34 = 0
    Me.Check46 = 0
    Me.Check58 = 0
    ques(9) = "E"
Else
    Me.txtexcellent = Val(Me.txtexcellent) - 1
End If

If Val(Me.txtexcellent) < 1 Then Me.txtexcellent = 0
End Sub

Private Sub Check11_Click()
If Me.Check11 = vbChecked Then
    Me.txtexcellent = Val(Me.txtexcellent) + 1
    Me.Check23 = 0
    Me.Check35 = 0
    Me.Check47 = 0
    Me.Check59 = 0
    ques(10) = "E"
Else
    Me.txtexcellent = Val(Me.txtexcellent) - 1
End If

If Val(Me.txtexcellent) < 1 Then Me.txtexcellent = 0
End Sub

Private Sub Check12_Click()
If Me.Check12 = vbChecked Then
    Me.txtexcellent = Val(Me.txtexcellent) + 1
    Me.Check24 = 0
    Me.Check36 = 0
    Me.Check48 = 0
    Me.Check60 = 0
    ques(11) = "E"
Else
    Me.txtexcellent = Val(Me.txtexcellent) - 1
End If

If Val(Me.txtexcellent) < 1 Then Me.txtexcellent = 0
End Sub

Private Sub Check13_Click()
If Me.Check13 = 1 Then
  Me.txtgood = Val(Me.txtgood) + 1
  Me.Check1 = 0
  Me.Check25 = 0
  Me.Check37 = 0
  Me.Check49 = 0
  ques(0) = "G"
Else
  Me.txtgood = Val(Me.txtgood) - 1
End If
If Me.txtgood < 1 Then Me.txtgood = 0



End Sub

Private Sub Check14_Click()
If Me.Check14 = 1 Then
  Me.txtgood = Val(Me.txtgood) + 1
  Me.Check2 = 0
  Me.Check26 = 0
  Me.Check38 = 0
  Me.Check50 = 0
  ques(1) = "G"
Else
  Me.txtgood = Val(Me.txtgood) - 1
End If
If Me.txtgood < 1 Then Me.txtgood = 0

End Sub

Private Sub Check15_Click()
If Me.Check15 = 1 Then
  Me.txtgood = Val(Me.txtgood) + 1
    Me.Check3 = 0
    Me.Check27 = 0
    Me.Check39 = 0
    Me.Check51 = 0
    ques(2) = "G"
Else
  Me.txtgood = Val(Me.txtgood) - 1
End If
If Me.txtgood < 1 Then Me.txtgood = 0

End Sub

Private Sub Check16_Click()
If Me.Check16 = 1 Then
  Me.txtgood = Val(Me.txtgood) + 1
  Me.Check4 = 0
  Me.Check28 = 0
  Me.Check40 = 0
  Me.Check52 = 0
  ques(3) = "G"
Else
  Me.txtgood = Val(Me.txtgood) - 1
End If
If Me.txtgood < 1 Then Me.txtgood = 0

End Sub

Private Sub Check17_Click()
If Me.Check17 = 1 Then
  Me.txtgood = Val(Me.txtgood) + 1
  Me.Check5 = 0
  Me.Check29 = 0
  Me.Check41 = 0
  Me.Check53 = 0
  ques(4) = "G"
Else
  Me.txtgood = Val(Me.txtgood) - 1
End If
If Me.txtgood < 1 Then Me.txtgood = 0

End Sub

Private Sub Check18_Click()
If Me.Check18 = 1 Then
  Me.txtgood = Val(Me.txtgood) + 1
  Me.Check6 = 0
  Me.Check30 = 0
  Me.Check42 = 0
  Me.Check54 = 0
  ques(5) = "G"
Else
  Me.txtgood = Val(Me.txtgood) - 1
End If
If Me.txtgood < 1 Then Me.txtgood = 0

End Sub

Private Sub Check19_Click()
If Me.Check19 = 1 Then
  Me.txtgood = Val(Me.txtgood) + 1
    Me.Check7 = 0
    Me.Check31 = 0
    Me.Check43 = 0
    Me.Check55 = 0
    ques(6) = "G"
Else
  Me.txtgood = Val(Me.txtgood) - 1
End If
If Me.txtgood < 1 Then Me.txtgood = 0

End Sub

Private Sub Check2_Click()
If Me.Check2 = vbChecked Then
    Me.txtexcellent = Val(Me.txtexcellent) + 1
    Me.Check14 = 0
    Me.Check26 = 0
    Me.Check38 = 0
    Me.Check50 = 0
    ques(1) = "E"
Else
    Me.txtexcellent = Val(Me.txtexcellent) - 1
End If

If Val(Me.txtexcellent) < 1 Then Me.txtexcellent = 0
End Sub

Private Sub Check20_Click()
If Me.Check20 = 1 Then
  Me.txtgood = Val(Me.txtgood) + 1
  Me.Check8 = 0
  Me.Check32 = 0
  Me.Check44 = 0
  Me.Check56 = 0
  ques(7) = "G"
Else
  Me.txtgood = Val(Me.txtgood) - 1
End If
If Me.txtgood < 1 Then Me.txtgood = 0

End Sub

Private Sub Check21_Click()
If Me.Check21 = 1 Then
  Me.txtgood = Val(Me.txtgood) + 1
  Me.Check9 = 0
  Me.Check33 = 0
  Me.Check45 = 0
  Me.Check57 = 0
  ques(8) = "G"
Else
  Me.txtgood = Val(Me.txtgood) - 1
End If
If Me.txtgood < 1 Then Me.txtgood = 0

End Sub

Private Sub Check22_Click()
If Me.Check22 = 1 Then
  Me.txtgood = Val(Me.txtgood) + 1
  Me.Check10 = 0
  Me.Check34 = 0
  Me.Check46 = 0
  Me.Check58 = 0
  ques(9) = "G"
Else
  Me.txtgood = Val(Me.txtgood) - 1
End If
If Me.txtgood < 1 Then Me.txtgood = 0

End Sub

Private Sub Check23_Click()
If Me.Check23 = 1 Then
  Me.txtgood = Val(Me.txtgood) + 1
  Me.Check11 = 0
  Me.Check35 = 0
  Me.Check47 = 0
  Me.Check59 = 0
  ques(10) = "G"
Else
  Me.txtgood = Val(Me.txtgood) - 1
End If
If Me.txtgood < 1 Then Me.txtgood = 0

End Sub

Private Sub Check24_Click()
If Me.Check24 = 1 Then
  Me.txtgood = Val(Me.txtgood) + 1
  Me.Check12 = 0
  Me.Check36 = 0
  Me.Check48 = 0
  Me.Check60 = 0
  ques(11) = "G"
Else
  Me.txtgood = Val(Me.txtgood) - 1
End If
If Me.txtgood < 1 Then Me.txtgood = 0

End Sub

Private Sub Check25_Click()
If Me.Check25 = vbChecked Then
    Me.txtaverage = Val(Me.txtaverage) + 1
    Me.Check13 = 0
    Me.Check1 = 0
    Me.Check37 = 0
    Me.Check49 = 0
    ques(0) = "A"
Else
    Me.txtaverage = Val(Me.txtaverage) - 1
End If

If Val(Me.txtaverage) < 1 Then Me.txtaverage = 0
End Sub

Private Sub Check26_Click()
If Me.Check26 = vbChecked Then
    Me.txtaverage = Val(Me.txtaverage) + 1
    Me.Check2 = 0
    Me.Check14 = 0
    Me.Check38 = 0
    Me.Check50 = 0
    ques(1) = "A"
Else
    Me.txtaverage = Val(Me.txtaverage) - 1
End If

If Val(Me.txtaverage) < 1 Then Me.txtaverage = 0
End Sub

Private Sub Check27_Click()
If Me.Check27 = vbChecked Then
    Me.txtaverage = Val(Me.txtaverage) + 1
    Me.Check3 = 0
    Me.Check15 = 0
    Me.Check39 = 0
    Me.Check51 = 0
    ques(2) = "A"
Else
    Me.txtaverage = Val(Me.txtaverage) - 1
End If

If Val(Me.txtaverage) < 1 Then Me.txtaverage = 0
End Sub

Private Sub Check28_Click()
If Me.Check28 = vbChecked Then
    Me.txtaverage = Val(Me.txtaverage) + 1
    Me.Check16 = 0
    Me.Check4 = 0
    Me.Check40 = 0
    Me.Check52 = 0
    ques(3) = "A"
Else
    Me.txtaverage = Val(Me.txtaverage) - 1
End If

If Val(Me.txtaverage) < 1 Then Me.txtaverage = 0
End Sub

Private Sub Check29_Click()
If Me.Check29 = vbChecked Then
    Me.txtaverage = Val(Me.txtaverage) + 1
    Me.Check17 = 0
    Me.Check5 = 0
    Me.Check41 = 0
    Me.Check53 = 0
    ques(4) = "A"
Else
    Me.txtaverage = Val(Me.txtaverage) - 1
End If

If Val(Me.txtaverage) < 1 Then Me.txtaverage = 0
End Sub

Private Sub Check3_Click()
If Me.Check3 = vbChecked Then
    Me.txtexcellent = Val(Me.txtexcellent) + 1
    Me.Check15 = 0
    Me.Check27 = 0
    Me.Check39 = 0
    Me.Check51 = 0
    ques(2) = "E"
Else
    Me.txtexcellent = Val(Me.txtexcellent) - 1
End If

If Val(Me.txtexcellent) < 1 Then Me.txtexcellent = 0
End Sub

Private Sub Check30_Click()
If Me.Check30 = vbChecked Then
    Me.txtaverage = Val(Me.txtaverage) + 1
    Me.Check18 = 0
    Me.Check6 = 0
    Me.Check42 = 0
    Me.Check54 = 0
    ques(5) = "A"
Else
    Me.txtaverage = Val(Me.txtaverage) - 1
End If

If Val(Me.txtaverage) < 1 Then Me.txtaverage = 0
End Sub

Private Sub Check31_Click()
If Me.Check31 = vbChecked Then
    Me.txtaverage = Val(Me.txtaverage) + 1
    Me.Check19 = 0
    Me.Check7 = 0
    Me.Check43 = 0
    Me.Check55 = 0
    ques(6) = "A"
Else
    Me.txtaverage = Val(Me.txtaverage) - 1
End If

If Val(Me.txtaverage) < 1 Then Me.txtaverage = 0
End Sub

Private Sub Check32_Click()
If Me.Check32 = vbChecked Then
    Me.txtaverage = Val(Me.txtaverage) + 1
    Me.Check20 = 0
    Me.Check8 = 0
    Me.Check44 = 0
    Me.Check56 = 0
    ques(7) = "A"
Else
    Me.txtaverage = Val(Me.txtaverage) - 1
End If

If Val(Me.txtaverage) < 1 Then Me.txtaverage = 0
End Sub

Private Sub Check33_Click()
If Me.Check33 = vbChecked Then
    Me.txtaverage = Val(Me.txtaverage) + 1
    Me.Check21 = 0
    Me.Check9 = 0
    Me.Check45 = 0
    Me.Check57 = 0
    ques(8) = "A"
Else
    Me.txtaverage = Val(Me.txtaverage) - 1
End If

If Val(Me.txtaverage) < 1 Then Me.txtaverage = 0
End Sub

Private Sub Check34_Click()
If Me.Check34 = vbChecked Then
    Me.txtaverage = Val(Me.txtaverage) + 1
    Me.Check22 = 0
    Me.Check10 = 0
    Me.Check46 = 0
    Me.Check58 = 0
    ques(9) = "A"
Else
    Me.txtaverage = Val(Me.txtaverage) - 1
End If

If Val(Me.txtaverage) < 1 Then Me.txtaverage = 0
End Sub

Private Sub Check35_Click()
If Me.Check35 = vbChecked Then
    Me.txtaverage = Val(Me.txtaverage) + 1
    Me.Check23 = 0
    Me.Check11 = 0
    Me.Check47 = 0
    Me.Check59 = 0
    ques(10) = "A"
Else
    Me.txtaverage = Val(Me.txtaverage) - 1
End If

If Val(Me.txtaverage) < 1 Then Me.txtaverage = 0
End Sub

Private Sub Check36_Click()
If Me.Check36 = vbChecked Then
    Me.txtaverage = Val(Me.txtaverage) + 1
    Me.Check24 = 0
    Me.Check12 = 0
    Me.Check48 = 0
    Me.Check60 = 0
    ques(11) = "A"
Else
    Me.txtaverage = Val(Me.txtaverage) - 1
End If

If Val(Me.txtaverage) < 1 Then Me.txtaverage = 0
End Sub

Private Sub Check37_Click()
If Me.Check37 = 1 Then
    Me.txtpoor = Val(Me.txtpoor) + 1
    Me.Check13 = 0
    Me.Check1 = 0
    Me.Check25 = 0
    Me.Check49 = 0
    ques(0) = "P"
Else
    Me.txtpoor = Val(Me.txtpoor) - 1
End If

If Me.txtpoor < 1 Then Me.txtpoor = 0
End Sub

Private Sub Check38_Click()
If Me.Check38 = 1 Then
    Me.txtpoor = Val(Me.txtpoor) + 1
    Me.Check2 = 0
    Me.Check26 = 0
    Me.Check14 = 0
    Me.Check50 = 0
    ques(1) = "P"
Else
    Me.txtpoor = Val(Me.txtpoor) - 1
End If

If Me.txtpoor < 1 Then Me.txtpoor = 0
End Sub

Private Sub Check39_Click()
If Me.Check39 = 1 Then
    Me.txtpoor = Val(Me.txtpoor) + 1
    Me.Check3 = 0
    Me.Check27 = 0
    Me.Check15 = 0
    Me.Check51 = 0
    ques(2) = "P"
Else
    Me.txtpoor = Val(Me.txtpoor) - 1
End If

If Me.txtpoor < 1 Then Me.txtpoor = 0
End Sub

Private Sub Check4_Click()
If Me.Check4 = vbChecked Then
    Me.txtexcellent = Val(Me.txtexcellent) + 1
    Me.Check16 = 0
    Me.Check28 = 0
    Me.Check40 = 0
    Me.Check52 = 0
    ques(3) = "E"
Else
    Me.txtexcellent = Val(Me.txtexcellent) - 1
End If

If Val(Me.txtexcellent) < 1 Then Me.txtexcellent = 0
End Sub

Private Sub Check40_Click()
If Me.Check40 = 1 Then
    Me.txtpoor = Val(Me.txtpoor) + 1
    Me.Check16 = 0
    Me.Check28 = 0
    Me.Check4 = 0
    Me.Check52 = 0
    ques(3) = "P"
Else
    Me.txtpoor = Val(Me.txtpoor) - 1
End If

If Me.txtpoor < 1 Then Me.txtpoor = 0
End Sub

Private Sub Check41_Click()
If Me.Check41 = 1 Then
    Me.txtpoor = Val(Me.txtpoor) + 1
    Me.Check17 = 0
    Me.Check29 = 0
    Me.Check5 = 0
    Me.Check53 = 0
    ques(4) = "P"
Else
    Me.txtpoor = Val(Me.txtpoor) - 1
End If

If Me.txtpoor < 1 Then Me.txtpoor = 0
End Sub

Private Sub Check42_Click()
If Me.Check42 = 1 Then
    Me.txtpoor = Val(Me.txtpoor) + 1
    Me.Check18 = 0
    Me.Check30 = 0
    Me.Check6 = 0
    Me.Check54 = 0
    ques(5) = "P"
Else
    Me.txtpoor = Val(Me.txtpoor) - 1
End If

If Me.txtpoor < 1 Then Me.txtpoor = 0
End Sub

Private Sub Check43_Click()
If Me.Check43 = 1 Then
    Me.txtpoor = Val(Me.txtpoor) + 1
    Me.Check19 = 0
    Me.Check31 = 0
    Me.Check7 = 0
    Me.Check55 = 0
    ques(6) = "P"
Else
    Me.txtpoor = Val(Me.txtpoor) - 1
End If

If Me.txtpoor < 1 Then Me.txtpoor = 0
End Sub

Private Sub Check44_Click()
If Me.Check44 = 1 Then
    Me.txtpoor = Val(Me.txtpoor) + 1
    Me.Check20 = 0
    Me.Check32 = 0
    Me.Check8 = 0
    Me.Check56 = 0
    ques(7) = "P"
Else
    Me.txtpoor = Val(Me.txtpoor) - 1
End If

If Me.txtpoor < 1 Then Me.txtpoor = 0
End Sub

Private Sub Check45_Click()
If Me.Check45 = 1 Then
    Me.txtpoor = Val(Me.txtpoor) + 1
    Me.Check21 = 0
    Me.Check33 = 0
    Me.Check9 = 0
    Me.Check57 = 0
    ques(8) = "P"
Else
    Me.txtpoor = Val(Me.txtpoor) - 1
End If

If Me.txtpoor < 1 Then Me.txtpoor = 0
End Sub

Private Sub Check46_Click()
If Me.Check46 = 1 Then
    Me.txtpoor = Val(Me.txtpoor) + 1
    Me.Check22 = 0
    Me.Check34 = 0
    Me.Check10 = 0
    Me.Check58 = 0
    ques(9) = "P"
Else
    Me.txtpoor = Val(Me.txtpoor) - 1
End If

If Me.txtpoor < 1 Then Me.txtpoor = 0
End Sub

Private Sub Check47_Click()
If Me.Check47 = 1 Then
    Me.txtpoor = Val(Me.txtpoor) + 1
    Me.Check23 = 0
    Me.Check35 = 0
    Me.Check11 = 0
    Me.Check59 = 0
    ques(10) = "P"
Else
    Me.txtpoor = Val(Me.txtpoor) - 1
End If

If Me.txtpoor < 1 Then Me.txtpoor = 0
End Sub

Private Sub Check48_Click()
If Me.Check48 = 1 Then
    Me.txtpoor = Val(Me.txtpoor) + 1
    Me.Check24 = 0
    Me.Check36 = 0
    Me.Check12 = 0
    Me.Check60 = 0
    ques(11) = "P"
Else
    Me.txtpoor = Val(Me.txtpoor) - 1
End If

If Me.txtpoor < 1 Then Me.txtpoor = 0
End Sub

Private Sub Check49_Click()
If Me.Check49 = vbChecked Then
    Me.txtoutstanding = Val(Me.txtoutstanding) + 1
    Me.Check1 = 0
    Me.Check13 = 0
    Me.Check25 = 0
    Me.Check37 = 0
    ques(0) = "O"
Else
     Me.txtoutstanding = Val(Me.txtoutstanding) - 1
End If

If Val(Me.txtoutstanding) < 1 Then Me.txtoutstanding = 0
End Sub

Private Sub Check5_Click()
If Me.Check5 = vbChecked Then
    Me.txtexcellent = Val(Me.txtexcellent) + 1
    Me.Check17 = 0
    Me.Check29 = 0
    Me.Check41 = 0
    Me.Check53 = 0
    ques(4) = "E"
Else
    Me.txtexcellent = Val(Me.txtexcellent) - 1
End If

If Val(Me.txtexcellent) < 1 Then Me.txtexcellent = 0
End Sub

Private Sub Check50_Click()
If Me.Check50 = vbChecked Then
    Me.txtoutstanding = Val(Me.txtoutstanding) + 1
    Me.Check2 = 0
    Me.Check14 = 0
    Me.Check26 = 0
    Me.Check38 = 0
    ques(1) = "O"
Else
     Me.txtoutstanding = Val(Me.txtoutstanding) - 1
End If

If Val(Me.txtoutstanding) < 1 Then Me.txtoutstanding = 0
End Sub

Private Sub Check51_Click()
If Me.Check51 = vbChecked Then
    Me.txtoutstanding = Val(Me.txtoutstanding) + 1
    Me.Check3 = 0
    Me.Check15 = 0
    Me.Check27 = 0
    Me.Check39 = 0
    ques(2) = "O"
Else
     Me.txtoutstanding = Val(Me.txtoutstanding) - 1
End If

If Val(Me.txtoutstanding) < 1 Then Me.txtoutstanding = 0
End Sub

Private Sub Check52_Click()
If Me.Check52 = vbChecked Then
    Me.txtoutstanding = Val(Me.txtoutstanding) + 1
    Me.Check4 = 0
    Me.Check16 = 0
    Me.Check28 = 0
    Me.Check40 = 0
    ques(3) = "O"
Else
     Me.txtoutstanding = Val(Me.txtoutstanding) - 1
End If

If Val(Me.txtoutstanding) < 1 Then Me.txtoutstanding = 0
End Sub

Private Sub Check53_Click()
If Me.Check53 = vbChecked Then
    Me.txtoutstanding = Val(Me.txtoutstanding) + 1
    Me.Check5 = 0
    Me.Check17 = 0
    Me.Check29 = 0
    Me.Check41 = 0
    ques(4) = "O"
Else
     Me.txtoutstanding = Val(Me.txtoutstanding) - 1
End If

If Val(Me.txtoutstanding) < 1 Then Me.txtoutstanding = 0
End Sub

Private Sub Check54_Click()
If Me.Check54 = vbChecked Then
    Me.txtoutstanding = Val(Me.txtoutstanding) + 1
    Me.Check6 = 0
    Me.Check18 = 0
    Me.Check30 = 0
    Me.Check42 = 0
    ques(5) = "O"
Else
     Me.txtoutstanding = Val(Me.txtoutstanding) - 1
End If

If Val(Me.txtoutstanding) < 1 Then Me.txtoutstanding = 0
End Sub

Private Sub Check55_Click()
If Me.Check55 = vbChecked Then
    Me.txtoutstanding = Val(Me.txtoutstanding) + 1
    Me.Check7 = 0
    Me.Check19 = 0
    Me.Check31 = 0
    Me.Check43 = 0
    ques(6) = "O"
Else
     Me.txtoutstanding = Val(Me.txtoutstanding) - 1
End If

If Val(Me.txtoutstanding) < 1 Then Me.txtoutstanding = 0
End Sub

Private Sub Check56_Click()
If Me.Check56 = vbChecked Then
    Me.txtoutstanding = Val(Me.txtoutstanding) + 1
    Me.Check8 = 0
    Me.Check20 = 0
    Me.Check32 = 0
    Me.Check44 = 0
    ques(7) = "O"
Else
     Me.txtoutstanding = Val(Me.txtoutstanding) - 1
End If

If Val(Me.txtoutstanding) < 1 Then Me.txtoutstanding = 0
End Sub

Private Sub Check57_Click()
If Me.Check57 = vbChecked Then
    Me.txtoutstanding = Val(Me.txtoutstanding) + 1
    Me.Check9 = 0
    Me.Check21 = 0
    Me.Check33 = 0
    Me.Check45 = 0
    ques(8) = "O"
Else
     Me.txtoutstanding = Val(Me.txtoutstanding) - 1
End If

If Val(Me.txtoutstanding) < 1 Then Me.txtoutstanding = 0
End Sub

Private Sub Check58_Click()
If Me.Check58 = vbChecked Then
    Me.txtoutstanding = Val(Me.txtoutstanding) + 1
    Me.Check10 = 0
    Me.Check22 = 0
    Me.Check34 = 0
    Me.Check46 = 0
    ques(9) = "O"
Else
     Me.txtoutstanding = Val(Me.txtoutstanding) - 1
End If

If Val(Me.txtoutstanding) < 1 Then Me.txtoutstanding = 0
End Sub

Private Sub Check59_Click()
If Me.Check59 = vbChecked Then
    Me.txtoutstanding = Val(Me.txtoutstanding) + 1
    Me.Check11 = 0
    Me.Check23 = 0
    Me.Check35 = 0
    Me.Check47 = 0
    ques(10) = "O"
Else
     Me.txtoutstanding = Val(Me.txtoutstanding) - 1
End If

If Val(Me.txtoutstanding) < 1 Then Me.txtoutstanding = 0
End Sub

Private Sub Check6_Click()
If Me.Check6 = vbChecked Then
    Me.txtexcellent = Val(Me.txtexcellent) + 1
    Me.Check18 = 0
    Me.Check30 = 0
    Me.Check42 = 0
    Me.Check54 = 0
    ques(5) = "E"
Else
    Me.txtexcellent = Val(Me.txtexcellent) - 1
End If

If Val(Me.txtexcellent) < 1 Then Me.txtexcellent = 0
End Sub

Private Sub Check60_Click()
If Me.Check60 = vbChecked Then
    Me.txtoutstanding = Val(Me.txtoutstanding) + 1
    Me.Check12 = 0
    Me.Check24 = 0
    Me.Check36 = 0
    Me.Check48 = 0
    ques(11) = "O"
Else
     Me.txtoutstanding = Val(Me.txtoutstanding) - 1
End If

If Val(Me.txtoutstanding) < 1 Then Me.txtoutstanding = 0
End Sub

Private Sub Check7_Click()
If Me.Check7 = vbChecked Then
    Me.txtexcellent = Val(Me.txtexcellent) + 1
    Me.Check19 = 0
    Me.Check31 = 0
    Me.Check43 = 0
    Me.Check55 = 0
    ques(6) = "E"
Else
    Me.txtexcellent = Val(Me.txtexcellent) - 1
End If

If Val(Me.txtexcellent) < 1 Then Me.txtexcellent = 0
End Sub

Private Sub Check8_Click()
If Me.Check8 = vbChecked Then
    Me.txtexcellent = Val(Me.txtexcellent) + 1
    Me.Check20 = 0
    Me.Check32 = 0
    Me.Check44 = 0
    Me.Check56 = 0
    ques(7) = "E"
Else
    Me.txtexcellent = Val(Me.txtexcellent) - 1
End If

If Val(Me.txtexcellent) < 1 Then Me.txtexcellent = 0
End Sub

Private Sub Check9_Click()
If Me.Check9 = vbChecked Then
    Me.txtexcellent = Val(Me.txtexcellent) + 1
    Me.Check21 = 0
    Me.Check33 = 0
    Me.Check45 = 0
    Me.Check57 = 0
    ques(8) = "E"
Else
    Me.txtexcellent = Val(Me.txtexcellent) - 1
End If

If Val(Me.txtexcellent) < 1 Then Me.txtexcellent = 0
End Sub

Private Sub cmdclose_Click()
Call CMDREST_Click
Me.Hide
Unload Me
End Sub

Private Sub CMDREST_Click()
Dim obj As Control
For Each obj In Me
    If TypeOf obj Is CheckBox Then
       obj = 0
       
    End If
Next
Dim obj2 As Control
For Each obj2 In Me
    If TypeOf obj2 Is ComboBox Then
       obj2 = ""
       
    End If
Next
For i = 0 To 11
    ques(i) = ""
Next i
End Sub

Private Sub cmdsave_Click()
''''''''''''''VALIDATION of Different Activities''''''''''''''''''

If (Val(Me.txtoutstanding) + Val(Me.txtexcellent) + Val(Me.txtgood) + Val(Me.txtaverage) + Val(Me.txtpoor)) <> 12 Then
   MsgBox "Total point should be 12", vbCritical, "Error"
   Exit Sub
End If
If Trim(Me.Combofaculty) = "" Then
  MsgBox "Please Input Faculty name", vbCritical, "Validation Error"
  Exit Sub
End If
If Trim(Me.Combosubcode) = "" Then
    MsgBox "Please Input Subject Code", vbCritical, "Validation Error"
    Exit Sub
End If
If Trim(Me.ComboSubname) = "" Then
    MsgBox "Please Input Subject Name", vbCritical, "Validation Error"
    Exit Sub
End If
If Trim(Me.Combosession) = "" Then
  MsgBox "Input Academic Session", vbCritical, "Validation Error"
  Exit Sub
End If
If Trim(Me.Combodept) = "" Then
    MsgBox "Input Department or Stream", vbCritical, "Validation Error": Exit Sub
  
End If
If IsDate(Me.txtdate) = False Then
    MsgBox "Please input date in DD/MM/YYYY format", vbCritical, "Validation Error"
    Exit Sub
End If
If Trim(Me.combosem) = "" Then
    MsgBox "Please Input Sem", vbCritical, "Validation Error"
    Exit Sub
End If
If Trim(Me.ComboCourse) = "" Then
    MsgBox "Please Input Proper Course", vbCritical, "Validation Error": Exit Sub
End If

'''''''Programme for Data saving into database'''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim rs As New ADODB.Recordset
Dim sql As String
sql = "SELECT * FROM EMP"
connect_db
CON = True
rs.Open sql, acon, adOpenStatic, adLockOptimistic
rs.AddNew
rs.Fields("empname") = Trim(Me.Combofaculty)
rs.Fields("subcode") = Trim(Me.Combosubcode)
rs.Fields("subname") = Trim(Me.ComboSubname)
rs.Fields("academicsession") = Trim(Me.Combosession)
rs.Fields("dept") = Trim(Me.Combodept)
rs.Fields("date1") = Format(Me.txtdate, "DD/MM/YYYY")
rs.Fields("sem") = Me.combosem
rs.Fields("course") = Me.ComboCourse

rs.Fields("Outstanding") = Val(Me.txtoutstanding)
rs.Fields("excellent") = Val(Me.txtexcellent)
rs.Fields("good") = Val(Me.txtgood)
rs.Fields("average") = Val(Me.txtaverage)
rs.Fields("poor") = Val(Me.txtpoor)

rs.Fields("ques1") = ques(0)
rs.Fields("ques2") = ques(1)
rs.Fields("ques3") = ques(2)
rs.Fields("ques4") = ques(3)
rs.Fields("ques5") = ques(4)
rs.Fields("ques6") = ques(5)
rs.Fields("ques7") = ques(6)
rs.Fields("ques8") = ques(7)
rs.Fields("ques9") = ques(8)
rs.Fields("ques10") = ques(9)
rs.Fields("ques11") = ques(10)
rs.Fields("ques12") = ques(11)

rs.Update

acon.Close
CON = False
MsgBox "Record Saved", vbInformation, "Saved"
Call CMDREST_Click
Me.Hide
Unload Me
End Sub

Private Sub Combosubname_GotFocus()
'''' Following is not required thats why it is commented '''''''''''''''''

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
'' Me.ComboSubname = rs.Fields("SUBNAME")
'' fnd = True
'' rs.MoveNext
''Loop
''acon.Close
''CON = False
''If fnd = False Then
''  MsgBox "Subject Code is not in database", vbInformation, "Not Found"
'' 'Me.Combosubcode.Text = ""
'' 'Me.Combosubcode.SetFocus
''End If
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
'==============================================================
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
'===============================================================
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

End Sub

