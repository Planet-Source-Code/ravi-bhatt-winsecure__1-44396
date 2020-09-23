VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVBUTTONS.OCX"
Begin VB.Form mw 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00F3CE9C&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "WinSecure 1.0.0 for Windows 9x/Me"
   ClientHeight    =   6540
   ClientLeft      =   105
   ClientTop       =   -180
   ClientWidth     =   9420
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FF0000&
   Icon            =   "mw.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   9420
   ShowInTaskbar   =   0   'False
   Begin LVbuttons.LaVolpeButton settings 
      Height          =   495
      Left            =   4320
      TabIndex        =   10
      Top             =   5850
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "Settings"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   13160664
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "mw.frx":0442
      ALIGN           =   1
      IMGLST          =   "(None)"
      IMGICON         =   "(None)"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   1
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   4635
      Top             =   135
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   4230
      Top             =   135
   End
   Begin LVbuttons.LaVolpeButton mice 
      Height          =   495
      Left            =   5535
      TabIndex        =   12
      Top             =   5850
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "Misc"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   13160664
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "mw.frx":045E
      ALIGN           =   1
      IMGLST          =   "(None)"
      IMGICON         =   "(None)"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   1
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin VB.CheckBox ChkSet 
      BackColor       =   &H00F3CE9C&
      Caption         =   "Select All"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1620
      TabIndex        =   11
      Top             =   5040
      Width           =   1230
   End
   Begin MSComctlLib.TreeView tv 
      Height          =   4425
      Left            =   1575
      TabIndex        =   5
      ToolTipText     =   "A check sign on a check box shows that the setting is enabled."
      Top             =   585
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   7805
      _Version        =   393217
      LineStyle       =   1
      Style           =   7
      Checkboxes      =   -1  'True
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin LVbuttons.LaVolpeButton SM 
      Height          =   495
      Left            =   315
      TabIndex        =   1
      Top             =   1315
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "Start Menu"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   13160664
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "mw.frx":047A
      ALIGN           =   1
      IMGLST          =   "(None)"
      IMGICON         =   "(None)"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   1
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin LVbuttons.LaVolpeButton Exp 
      Height          =   495
      Left            =   315
      TabIndex        =   2
      Top             =   1910
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "Explorer"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   13160664
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "mw.frx":0496
      ALIGN           =   1
      IMGLST          =   "(None)"
      IMGICON         =   "(None)"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   1
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin LVbuttons.LaVolpeButton Desk 
      Height          =   495
      Left            =   315
      TabIndex        =   3
      Top             =   2505
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "Desktop"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   13160664
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "mw.frx":04B2
      ALIGN           =   1
      IMGLST          =   "(None)"
      IMGICON         =   "(None)"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   1
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin LVbuttons.LaVolpeButton Cp 
      Height          =   495
      Left            =   315
      TabIndex        =   4
      Top             =   3100
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "Control Panel"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   13160664
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "mw.frx":04CE
      ALIGN           =   1
      IMGLST          =   "(None)"
      IMGICON         =   "(None)"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   1
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton1 
      Height          =   495
      Left            =   315
      TabIndex        =   6
      Top             =   3695
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "Internet Explorer"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   13160664
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "mw.frx":04EA
      ALIGN           =   1
      IMGLST          =   "(None)"
      IMGICON         =   "(None)"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   1
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton2 
      Height          =   495
      Left            =   315
      TabIndex        =   7
      Top             =   4290
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "Network"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   13160664
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "mw.frx":0506
      ALIGN           =   1
      IMGLST          =   "(None)"
      IMGICON         =   "(None)"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   1
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin LVbuttons.LaVolpeButton Disk 
      Height          =   495
      Left            =   315
      TabIndex        =   0
      Top             =   720
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "Disk"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   13160664
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "mw.frx":0522
      ALIGN           =   1
      IMGLST          =   "(None)"
      IMGICON         =   "(None)"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   1
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin LVbuttons.LaVolpeButton more 
      Height          =   495
      Left            =   6750
      TabIndex        =   13
      Top             =   5850
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "More..."
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   13160664
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "mw.frx":053E
      ALIGN           =   1
      IMGLST          =   "(None)"
      IMGICON         =   "(None)"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   1
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin LVbuttons.LaVolpeButton sd 
      Height          =   495
      Left            =   4320
      TabIndex        =   14
      Top             =   5265
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "Shut Down"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   13160664
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "mw.frx":055A
      ALIGN           =   1
      IMGLST          =   "(None)"
      IMGICON         =   "(None)"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   1
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin LVbuttons.LaVolpeButton rest 
      Height          =   495
      Left            =   5535
      TabIndex        =   15
      Top             =   5265
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "Restart"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   13160664
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "mw.frx":0576
      ALIGN           =   1
      IMGLST          =   "(None)"
      IMGICON         =   "(None)"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   1
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin LVbuttons.LaVolpeButton lo 
      Height          =   495
      Left            =   6750
      TabIndex        =   16
      Top             =   5265
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "LogOff"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   13160664
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "mw.frx":0592
      ALIGN           =   1
      IMGLST          =   "(None)"
      IMGICON         =   "(None)"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   1
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin LVbuttons.LaVolpeButton ApplyChange 
      Height          =   495
      Left            =   3105
      TabIndex        =   17
      Top             =   5265
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "Apply Changes"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   13160664
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "mw.frx":05AE
      ALIGN           =   1
      IMGLST          =   "(None)"
      IMGICON         =   "(None)"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   1
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin VB.TextBox msg 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4425
      Index           =   0
      Left            =   5715
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   22
      Top             =   585
      Visible         =   0   'False
      Width           =   3480
   End
   Begin VB.TextBox msg 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4425
      Index           =   1
      Left            =   5715
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   21
      Top             =   585
      Visible         =   0   'False
      Width           =   3480
   End
   Begin VB.TextBox msg 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4425
      Index           =   2
      Left            =   5715
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   20
      Top             =   585
      Visible         =   0   'False
      Width           =   3480
   End
   Begin VB.TextBox msg 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4425
      Index           =   3
      Left            =   5715
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   19
      Top             =   585
      Visible         =   0   'False
      Width           =   3480
   End
   Begin VB.TextBox msg 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4425
      Index           =   4
      Left            =   5715
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   18
      Top             =   585
      Visible         =   0   'False
      Width           =   3480
   End
   Begin VB.TextBox msg 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4425
      Index           =   6
      Left            =   5715
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   24
      Top             =   585
      Visible         =   0   'False
      Width           =   3480
   End
   Begin VB.TextBox msg 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4425
      Index           =   5
      Left            =   5715
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   23
      Top             =   585
      Visible         =   0   'False
      Width           =   3480
   End
   Begin VB.TextBox msg 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4425
      Index           =   7
      Left            =   5715
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   25
      Top             =   585
      Visible         =   0   'False
      Width           =   3480
   End
   Begin LVbuttons.LaVolpeButton butAbt 
      Height          =   495
      Left            =   7965
      TabIndex        =   29
      Top             =   5850
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "About"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   13160664
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "mw.frx":05CA
      ALIGN           =   1
      IMGLST          =   "(None)"
      IMGICON         =   "(None)"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   1
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ravi Bhatt"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   270
      Left            =   45
      MouseIcon       =   "mw.frx":05E6
      MousePointer    =   99  'Custom
      TabIndex        =   28
      Top             =   5895
      Width           =   1140
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Developed By:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   0
      MouseIcon       =   "mw.frx":0A28
      MousePointer    =   99  'Custom
      TabIndex        =   27
      Top             =   5625
      Width           =   1770
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "rbhatt123@rediffmail.com"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   270
      Left            =   45
      MouseIcon       =   "mw.frx":0E6A
      MousePointer    =   99  'Custom
      TabIndex        =   26
      Top             =   6165
      Width           =   2925
   End
   Begin VB.Shape Shape3 
      Height          =   600
      Left            =   3015
      Top             =   5220
      Width           =   4920
   End
   Begin VB.Shape Shape2 
      Height          =   600
      Left            =   4230
      Top             =   5805
      Width           =   4920
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00F3CE9C&
      BackStyle       =   0  'Transparent
      Caption         =   "Description : "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   5715
      TabIndex        =   9
      Top             =   225
      Width           =   1605
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00F3CE9C&
      BackStyle       =   0  'Transparent
      Caption         =   "Properties : "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   1620
      TabIndex        =   8
      Top             =   225
      Width           =   1515
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00F3CE9C&
      FillColor       =   &H00F3CE9C&
      FillStyle       =   0  'Solid
      Height          =   690
      Index           =   6
      Left            =   0
      Top             =   4230
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00F3CE9C&
      FillColor       =   &H00F3CE9C&
      FillStyle       =   0  'Solid
      Height          =   690
      Index           =   5
      Left            =   0
      Top             =   3600
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00F3CE9C&
      FillColor       =   &H00F3CE9C&
      FillStyle       =   0  'Solid
      Height          =   690
      Index           =   4
      Left            =   0
      Top             =   3015
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00F3CE9C&
      FillColor       =   &H00F3CE9C&
      FillStyle       =   0  'Solid
      Height          =   690
      Index           =   3
      Left            =   0
      Top             =   2430
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00F3CE9C&
      FillColor       =   &H00F3CE9C&
      FillStyle       =   0  'Solid
      Height          =   690
      Index           =   2
      Left            =   0
      Top             =   1845
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00F3CE9C&
      FillColor       =   &H00F3CE9C&
      FillStyle       =   0  'Solid
      Height          =   690
      Index           =   1
      Left            =   0
      Top             =   1215
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00F3CE9C&
      FillColor       =   &H00F3CE9C&
      FillStyle       =   0  'Solid
      Height          =   690
      Index           =   0
      Left            =   0
      Top             =   630
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.Shape Sha 
      BorderColor     =   &H80000004&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   8655
      Left            =   0
      Top             =   -45
      Width           =   960
   End
   Begin VB.Menu mnuDisk 
      Caption         =   "&Disk"
      Visible         =   0   'False
      Begin VB.Menu mnuDiskHideAll 
         Caption         =   "&Hide All Drives"
      End
      Begin VB.Menu mnuDiskRestrictAccess 
         Caption         =   "&Restrict Access"
         Begin VB.Menu mnuDiskRestrictAccessA 
            Caption         =   "A"
            Index           =   3
         End
         Begin VB.Menu mnuDiskRestrictAccessA 
            Caption         =   "B"
            Index           =   4
         End
         Begin VB.Menu mnuDiskRestrictAccessA 
            Caption         =   "C"
            Index           =   5
         End
         Begin VB.Menu mnuDiskRestrictAccessA 
            Caption         =   "D"
            Index           =   6
         End
         Begin VB.Menu mnuDiskRestrictAccessA 
            Caption         =   "E"
            Index           =   7
         End
         Begin VB.Menu mnuDiskRestrictAccessA 
            Caption         =   "F"
            Index           =   8
         End
         Begin VB.Menu mnuDiskRestrictAccessA 
            Caption         =   "G"
            Index           =   9
         End
         Begin VB.Menu mnuDiskRestrictAccessA 
            Caption         =   "H"
            Index           =   10
         End
         Begin VB.Menu mnuDiskRestrictAccessA 
            Caption         =   "I"
            Index           =   11
         End
         Begin VB.Menu mnuDiskRestrictAccessA 
            Caption         =   "J"
            Index           =   12
         End
         Begin VB.Menu mnuDiskRestrictAccessA 
            Caption         =   "K"
            Index           =   13
         End
         Begin VB.Menu mnuDiskRestrictAccessA 
            Caption         =   "L"
            Index           =   14
         End
         Begin VB.Menu mnuDiskRestrictAccessA 
            Caption         =   "M"
            Index           =   15
         End
         Begin VB.Menu mnuDiskRestrictAccessA 
            Caption         =   "N"
            Index           =   16
         End
         Begin VB.Menu mnuDiskRestrictAccessA 
            Caption         =   "O"
            Index           =   17
         End
         Begin VB.Menu mnuDiskRestrictAccessA 
            Caption         =   "P"
            Index           =   18
         End
         Begin VB.Menu mnuDiskRestrictAccessA 
            Caption         =   "Q"
            Index           =   19
         End
         Begin VB.Menu mnuDiskRestrictAccessA 
            Caption         =   "R"
            Index           =   20
         End
         Begin VB.Menu mnuDiskRestrictAccessA 
            Caption         =   "S"
            Index           =   21
         End
         Begin VB.Menu mnuDiskRestrictAccessA 
            Caption         =   "T"
            Index           =   22
         End
         Begin VB.Menu mnuDiskRestrictAccessA 
            Caption         =   "U"
            Index           =   23
         End
         Begin VB.Menu mnuDiskRestrictAccessA 
            Caption         =   "V"
            Index           =   24
         End
         Begin VB.Menu mnuDiskRestrictAccessA 
            Caption         =   "W"
            Index           =   25
         End
         Begin VB.Menu mnuDiskRestrictAccessA 
            Caption         =   "X"
            Index           =   26
         End
         Begin VB.Menu mnuDiskRestrictAccessA 
            Caption         =   "Y"
            Index           =   27
         End
         Begin VB.Menu mnuDiskRestrictAccessA 
            Caption         =   "Z"
            Index           =   28
         End
      End
   End
   Begin VB.Menu mnuStartMenu 
      Caption         =   "&Start Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuStart 
         Caption         =   "Alphabetic Start Menu"
         Index           =   1
      End
      Begin VB.Menu mnuStart 
         Caption         =   "Clear Recent Documents On Exit"
         Index           =   2
      End
      Begin VB.Menu mnuStart 
         Caption         =   "Disable Darg Drop"
         Index           =   3
      End
      Begin VB.Menu mnuStart 
         Caption         =   "No Shut Down"
         Index           =   4
      End
      Begin VB.Menu mnuStart 
         Caption         =   "No Favorites Menu"
         Index           =   5
      End
      Begin VB.Menu mnuStart 
         Caption         =   "No Folder Options"
         Index           =   6
      End
      Begin VB.Menu mnuStart 
         Caption         =   "No Log Off"
         Index           =   7
      End
      Begin VB.Menu mnuStart 
         Caption         =   "No Recent Documents Menu"
         Index           =   8
      End
      Begin VB.Menu mnuStart 
         Caption         =   "No Run"
         Index           =   9
      End
      Begin VB.Menu mnuStart 
         Caption         =   "No Search"
         Index           =   10
      End
      Begin VB.Menu mnuStart 
         Caption         =   "No Taskbar Context Menu"
         Index           =   11
      End
      Begin VB.Menu mnuStart 
         Caption         =   "No Windows Update"
         Index           =   12
      End
      Begin VB.Menu mnuHideSB 
         Caption         =   "Hide Start Button"
      End
   End
   Begin VB.Menu mnuExplorer 
      Caption         =   "&Explorer"
      Visible         =   0   'False
      Begin VB.Menu mnuExplorerDRT 
         Caption         =   "Disable Registry Tools"
      End
      Begin VB.Menu mnuExplorerNoFile 
         Caption         =   "No File Menu"
      End
      Begin VB.Menu mnuExpSet 
         Caption         =   "Show Hidden Files"
         Index           =   26
      End
      Begin VB.Menu mnuExpSet 
         Caption         =   "Show File Extension For Known File Types"
         Index           =   27
      End
      Begin VB.Menu mnuExpSet 
         Caption         =   "Show Operating System Files"
         Index           =   28
      End
      Begin VB.Menu mnuExpSet 
         Caption         =   "Hide Tips On Items"
         Index           =   29
      End
   End
   Begin VB.Menu mnuDesktop 
      Caption         =   "Desk&top"
      Visible         =   0   'False
      Begin VB.Menu mnuDesktopDACD 
         Caption         =   "Disable ALT+CTRL+DEL, ALT+TAB, START BUTTON"
      End
      Begin VB.Menu mnuDesktopDRC 
         Caption         =   "Disable Right Click"
      End
      Begin VB.Menu mnuDesktopERB 
         Caption         =   "Empty Recycle Bin"
      End
      Begin VB.Menu mnuDesktopHide 
         Caption         =   "Hide Desktop"
      End
      Begin VB.Menu mnuDesktopHT 
         Caption         =   "Hide Taskbar"
      End
      Begin VB.Menu mnuDesktopHDT 
         Caption         =   "Hide Date Time And Icon Tray"
      End
      Begin VB.Menu mnuDesktopNNNI 
         Caption         =   "No Network Neighbourhood Icon"
      End
   End
   Begin VB.Menu mnuControlPanel 
      Caption         =   "&Control Panel"
      Visible         =   0   'False
      Begin VB.Menu mnuControlPanelk 
         Caption         =   "Disable Password Control Panel"
         Index           =   13
      End
      Begin VB.Menu mnuControlPanelk 
         Caption         =   "Hide Appearance Page"
         Index           =   14
      End
      Begin VB.Menu mnuControlPanelk 
         Caption         =   "Hide Background Page"
         Index           =   15
      End
      Begin VB.Menu mnuControlPanelk 
         Caption         =   "Hide Device Manager Page"
         Index           =   16
      End
      Begin VB.Menu mnuControlPanelk 
         Caption         =   "Disable Display Control Panel"
         Index           =   17
      End
      Begin VB.Menu mnuControlPanelk 
         Caption         =   "Hide File System Page"
         Index           =   18
      End
      Begin VB.Menu mnuControlPanelk 
         Caption         =   "Hide Hardware Profiles Page"
         Index           =   19
      End
      Begin VB.Menu mnuControlPanelk 
         Caption         =   "Hide Password Change Page"
         Index           =   20
      End
      Begin VB.Menu mnuControlPanelk 
         Caption         =   "Hide Remote Admin Applet"
         Index           =   21
      End
      Begin VB.Menu mnuControlPanelk 
         Caption         =   "Hide Settings Page"
         Index           =   22
      End
      Begin VB.Menu mnuControlPanelk 
         Caption         =   "Hide Screen Saver Page"
         Index           =   23
      End
      Begin VB.Menu mnuControlPanelk 
         Caption         =   "Disable User Profiles Page"
         Index           =   24
      End
      Begin VB.Menu mnuControlPanelk 
         Caption         =   "Hide Virtual Memory Button"
         Index           =   25
      End
   End
   Begin VB.Menu mnuIE 
      Caption         =   "Internet Explorer"
      Visible         =   0   'False
      Begin VB.Menu mnuIESet 
         Caption         =   "Disable Close"
         Index           =   31
      End
      Begin VB.Menu mnuIESet 
         Caption         =   "Disable Right Click"
         Index           =   32
      End
      Begin VB.Menu mnuIESet 
         Caption         =   "Disable Tools/Internet Option Menu"
         Index           =   33
      End
      Begin VB.Menu mnuIESet 
         Caption         =   "Disable Save As"
         Index           =   34
      End
      Begin VB.Menu mnuIESet 
         Caption         =   "Disable Favourites"
         Index           =   35
      End
      Begin VB.Menu mnuIESet 
         Caption         =   "Disable File/New"
         Index           =   36
      End
      Begin VB.Menu mnuIESet 
         Caption         =   "Disable File/Open"
         Index           =   37
      End
      Begin VB.Menu mnuIESet 
         Caption         =   "Disable Find Files"
         Index           =   38
      End
      Begin VB.Menu mnuIESet 
         Caption         =   "Disable Download Directory Option"
         Index           =   39
      End
      Begin VB.Menu mnuIESet 
         Caption         =   "Disable Full Screen View Option"
         Index           =   40
      End
   End
   Begin VB.Menu mnuNetwork 
      Caption         =   "NetWork"
      Visible         =   0   'False
      Begin VB.Menu mnuNetSet 
         Caption         =   "Hide Network Security Page"
         Index           =   41
      End
      Begin VB.Menu mnuNetSet 
         Caption         =   "Hide/Disable Network in Control Panel"
         Index           =   42
      End
      Begin VB.Menu mnuNetSet 
         Caption         =   "Hide Identification Page"
         Index           =   43
      End
      Begin VB.Menu mnuNetSet 
         Caption         =   "Disable File Sharing Controls"
         Index           =   44
      End
      Begin VB.Menu mnuNetSet 
         Caption         =   "Disable Print Sharing Controls"
         Index           =   45
      End
   End
   Begin VB.Menu mnutb 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuMaxi 
         Caption         =   "Open"
      End
      Begin VB.Menu mnurck 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "mw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
DefLng A-Z
Dim n() As Node
Dim i As Integer
Dim lngresult
Dim b As Boolean
Dim pHwnd As Long
Dim cHwnd As Long

Private Sub ApplyChange_Click()
'Where ever  str3 & "\users\" is used, I've stored values in registry
'for future reference. this is used where it is not possible or rather
'difficult to read back original registry value. again all settings
'made by API calls are also stroed here.
'Normally 1 is enabbled and 0 is disabled. except some special cases.
  On Error Resume Next
  If tv.Nodes.Count = 5 Then '5 NETWORK SETTING
    For i = 41 To 45 'LOOP & CHECK WHICH IS ENABLED
      If tv.Nodes.Item(i - 40).Checked = True Then
        RegRW str6, regKeys(i), 1, "REG_DWORD", mnuNetSet(i), True
      Else
        RegRW str6, regKeys(i), 0, "REG_DWORD", mnuNetSet(i), False
      End If
    Next
    Exit Sub
  End If
  If tv.Nodes.Count = 6 Then '6 EXPLORER SETTING
    For i = 26 To 29 'LOOP & CHECK WHICH IS ENABLED
      If tv.Nodes.Item(i - 23).Checked = True Then
        If i = 27 Or i = 29 Then
          RegRW str2, regKeys(i), 0, "REG_DWORD", mnuExpSet(i), True
        Else
          RegRW str2, regKeys(i), 1, "REG_DWORD", mnuExpSet(i), True
        End If
      Else
         If i = 27 Or i = 29 Then
          RegRW str2, regKeys(i), 1, "REG_DWORD", mnuExpSet(i), False
         ElseIf i = 28 Then
          RegRW str2, regKeys(i), 0, "REG_DWORD", mnuExpSet(i), False
         Else
          RegRW str2, regKeys(i), 2, "REG_DWORD", mnuExpSet(i), False
        End If
      End If
    Next
    Exit Sub
  End If
  If tv.Nodes.Count = 10 Then '10 IE SETTINGS
     For i = 31 To 40
       If tv.Nodes.Item(i - 30).Checked = True Then
         RegRW str5, regKeys(i), 1, "REG_DWORD", mnuIESet(i), True
       Else
         RegRW str5, regKeys(i), 0, "REG_DWORD", mnuIESet(i), False
       End If
     Next
     Exit Sub
  End If
  If tv.Nodes.Count = 13 And tv.Nodes.Item(1).Text = "Disable Password Control Panel" Then
    For i = 13 To 25 '13 CONTROL PANEL SETTINGS
        If tv.Nodes.Item(i - 12).Checked = False Then
           RegRW str4, regKeys(i), 0, "REG_DWORD", mnuControlPanelk(i), True
        Else
           RegRW str4, regKeys(i), 1, "REG_DWORD", mnuControlPanelk(i), False
        End If
    Next
    Exit Sub
  End If
  For i = 1 To tv.Nodes.Count 'EVERY THING ELSE
    Select Case tv.Nodes.Item(i).key
      Case "HAD": 'HAD = HIDE ALL DRIVES
                  If tv.Nodes.Item(i).Checked = True Then
                      RegRW str1, "NoDrives", 67108863, "REG_DWORD", mnuDiskHideAll, True
                  Else
                      RegRW str1, "NoDrives", 0, "REG_DWORD", mnuDiskHideAll, False
                  End If
      
      Case "AlphabeticStartMenu":
                                   Call mnuStartAlpha_Click
      Case "ResAc":
      Case "TaskbarContextMenu":
                                   If tv.Nodes.Item(i).Checked = True Then
                                         RegRW str2, "TaskbarContextMenu", 0, "REG_DWORD", mnuStart(i), True
                                   Else
                                         RegRW str2, "TaskbarContextMenu", 1, "REG_DWORD", mnuStart(i), False
                                   End If
      Case regKeys(i): ' START MENU SETTINGS
                             If tv.Nodes.Item(i).Checked = True Then
                                  RegRW str1, regKeys(i), 1, "REG_DWORD", mnuStart(i), True
                                  If i = 7 Then
                                     RegRW str2, "StartMenuLogOff", 0, "REG_DWORD"
                                  End If
                             Else
                                  RegRW str1, regKeys(i), 0, "REG_DWORD", mnuStart(i), False
                                  If i = 7 Then
                                     RegRW str2, "StartMenuLogOff", 1, "REG_DWORD"
                                  End If
                             End If
      Case "DRT":
                         If tv.Nodes.Item(1).Checked = True Then
                              RegRW str4, "DisableRegistryTools", 1, "REG_DWORD", mnuExplorerDRT, True
                         Else
                              RegRW str4, "DisableRegistryTools", 0, "REG_DWORD", mnuExplorerDRT, False
                         End If
      Case "NOFILE":
                         If tv.Nodes.Item(2).Checked = True Then
                              RegRW str1, "NoFileMenu", 1, "REG_DWORD", mnuExplorerNoFile, True
                         Else
                              RegRW str1, "NoFileMenu", 0, "REG_DWORD", mnuExplorerNoFile, False
                         End If
     Case "ALTCTLDEL":
                         If tv.Nodes.Item(i).Checked = True Then
                              Call DisableCtrlAltDelete(True)
                              mnuDesktopDACD.Checked = True
                              wscr.regwrite str3 & "\users\" & "ACD", 1, "REG_DWORD"
                         Else
                              Call DisableCtrlAltDelete(False)
                              mnuDesktopDACD.Checked = False
                              wscr.regwrite str3 & "\users\" & "ACD", 0, "REG_DWORD"
                         End If
     Case "RightClick": 'DESKTOP RIGHT CLICK ENABLE/DISABLE
                         If tv.Nodes.Item(i).Checked = True Then
                              RegRW str1, "NoViewContextMenu", 1, "REG_DWORD", mnuDesktopDRC, True
                         Else
                              RegRW str1, "NoViewContextMenu", 0, "REG_DWORD", mnuDesktopDRC, False
                         End If
     Case "HideDesk": 'HIDE DESKTOP
                         g_cstrShellViewWnd = "progman"
                         If tv.Nodes.Item(i).Checked = True Then
                            b = True
                            Call mnuDesktopHide_Click
                            wscr.regwrite str3 & "\users\" & "HD", 1, "REG_DWORD"
                            mnuDesktopHide.Checked = True
                         Else
                            b = False
                            Call mnuDesktopHide_Click
                            wscr.regwrite str3 & "\users\" & "HD", 0, "REG_DWORD"
                            mnuDesktopHide.Checked = False
                         End If
     Case "HTB": 'HIDE TASKBAR
                           g_cstrShellViewWnd = "Shell_traywnd"
                           
                           Dim hwnd As Long
                           On Error Resume Next
                           If tv.Nodes.Item(i).Checked = True Then
                             hwnd = FindShellWindow()
                             If hwnd <> 0 Then
                                  Call HideShowWindow(hwnd, True)
                                  RegRW str3 & "\users\", "HT", 1, "REG_DWORD", mnuDesktopHT, True
                             End If
                           Else
                             hwnd = FindShellWindow()
                             If hwnd <> 0 Then
                                  Call HideShowWindow(hwnd)
                                  RegRW str3 & "\users\", "HT", 0, "REG_DWORD", mnuDesktopHT, False
                             End If
                           End If
     Case "StartButtonHide":
                               
                               pHwnd = FINDWINDOW("Shell_traywnd", vbNullString)
                               cHwnd = FindWindowEx(pHwnd, 0&, "button", vbNullString)

                               If tv.Nodes.Item(i).Checked = True Then
                                     Call HideShowWindow(cHwnd, True)
                                     RegRW str3 & "\users\", "HSB", 1, "REG_DWORD", mnuHideSB, True
                               Else
                                     Call HideShowWindow(cHwnd)
                                     RegRW str3 & "\users\", "HSB", 0, "REG_DWORD", mnuHideSB, False
                               End If
     
     Case "DateTimeHide":
                              pHwnd = FINDWINDOW("Shell_traywnd", vbNullString)
                              cHwnd = FindWindowEx(pHwnd, 0&, "TrayNotifyWnd", vbNullString)

                               If tv.Nodes.Item(i).Checked = True Then
                                     Call HideShowWindow(cHwnd, True)
                                     RegRW str3 & "\users\", "HDT", 1, "REG_DWORD", mnuDesktopHDT, True
                               Else
                                     Call HideShowWindow(cHwnd)
                                     RegRW str3 & "\users\", "HDT", 0, "REG_DWORD", mnuDesktopHDT, False
                               End If
     
     Case "ERB": 'EMPTY RECYCLE BIN
                         If tv.Nodes.Item(i).Checked = True Then
                              lngresult = SHEmptyRecycleBin(mw.hwnd, "", SHERB_NOPROGRESSUI)
                         End If
     Case "NNNI": 'NO NETWORK ICON ON DESKTOP
                         If tv.Nodes.Item(i).Checked = True Then
                           RegRW str1, "NoNetHood", 1, "REG_DWORD", mnuDesktopNNNI, True
                         Else
                           RegRW str1, "NoNetHood", 0, "REG_DWORD", mnuDesktopNNNI, False
                         End If
     Case Else: 'STILL REMAINING GOES HERE
                'THIS FOR DISABLING ANY COMBINATION OF DRIVES.
                'CREATE NoViewOnDrive KEY INSIDE
                '"HKCU\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\"
                'AND ASSIGN IT VALUE OF 1 FOR A DRIVE, 2 FOR B DRIVE, 4 FOR C DRIVE,
                '8 FOR D DRIVE AND SO ON.....
                'FOR A+C DRIVE VALUE IS 1+4=5 AND SO ON...
                         Dim j As Double
                         j = wscr.regread(str1 & "NoViewOnDrive")
                         If tv.Nodes.Item(i).Checked = False And mnuDiskRestrictAccessA.Item(i).Checked = True Then
                           j = j - mnuDiskRestrictAccessA.Item(i).Tag
                           mnuDiskRestrictAccessA.Item(i).Checked = False
                           wscr.regwrite str3 & "\users\" & i, 0, "REG_DWORD"
                         ElseIf tv.Nodes.Item(i).Checked = True And mnuDiskRestrictAccessA.Item(i).Checked = False Then
                           j = j + mnuDiskRestrictAccessA.Item(i).Tag
                           mnuDiskRestrictAccessA.Item(i).Checked = True
                           wscr.regwrite str3 & "\users\" & i, 1, "REG_DWORD"
                         End If
                         RegRW str1, "NoViewOnDrive", j, "REG_DWORD"
    End Select
  Next
End Sub

Private Sub butsm_Click()

End Sub

Private Sub butAbt_Click()
  frmAbt.Show 1
End Sub

Private Sub ChkSet_Click()
   Dim i As Integer
   Dim j As Boolean
   
   If ChkSet.Value = 1 And tv.Nodes.Count > 0 Then
         j = True
         MsgBox "Applying all the settings at the same time may prove fatal to your system.", vbCritical, "Warning"
   Else
         j = False
   End If
   
   On Error Resume Next
   For i = 0 To tv.Nodes.Count
     tv.Nodes.Item(i).Checked = j
   Next
End Sub

Private Sub Cp_Click()
   Label3.Caption = "Properties : Control Panel"
   tv.Nodes.Clear 'clear previous nodes of treeview.
   ReDim n(13) As Node
   Dim i%
   'as array of menu is created and that array's index is same as the
   'index of string of regkeys.
   For i = 13 To mnuControlPanelk.Count + 12
        Set n(i - 12) = tv.Nodes.Add 'add new node in tree view
        n(i - 12).Text = mnuControlPanelk(i).Caption 'Assign menu's caption to node
        n(i - 12).Checked = mnuControlPanelk(i).Checked 'assign check property
        n(i - 12).key = regKeys(i) 'assigning actual registry string to node's key property
        'this key property is used in half of apply button click's coding.
        'this key is used to identify various settings in that sub.
   Next
   
   Call hidetext ' hide all other visible text boxes.
   msg(4).Visible = True 'text box number 4 is now visible.
   'an array of text boxes are used for displaying messages.
End Sub

Private Sub Cp_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
     Call hideShape
     Shape1(4).Visible = True
End Sub
Private Sub hideShape()
   For i = 0 To Shape1.Count - 1
        Shape1.Item(i).Visible = False
   Next
End Sub

Private Sub Cp_MouseOut()
   Call hideShape
End Sub

Private Sub Desk_Click()
    Label3.Caption = "Properties : Desktop"
    tv.Nodes.Clear
    ReDim n(7) As Node
    'no menu array here.
    Set n(1) = tv.Nodes.Add
    n(1).Text = mnuDesktopDACD.Caption
    n(1).key = "ALTCTLDEL"
    n(1).Checked = mnuDesktopDACD.Checked
    
    Set n(2) = tv.Nodes.Add
    n(2).Text = mnuDesktopDRC.Caption
    n(2).key = "RightClick"
    n(2).Checked = mnuDesktopDRC.Checked
  
    Set n(3) = tv.Nodes.Add
    n(3).Text = mnuDesktopHide.Caption
    n(3).key = "HideDesk"
    n(3).Checked = mnuDesktopHide.Checked
       
    Set n(4) = tv.Nodes.Add
    n(4).Text = "Hide TaskBar"
    n(4).key = "HTB"
    n(4).Checked = mnuDesktopHT.Checked
    
    Set n(5) = tv.Nodes.Add
    n(5).Text = "Empty Recycle Bin"
    n(5).key = "ERB"
    n(5).Checked = mnuDesktopERB.Checked
    
    Set n(6) = tv.Nodes.Add
    n(6).Text = mnuDesktopNNNI.Caption
    n(6).key = "NNNI"
    n(6).Checked = mnuDesktopNNNI.Checked
    
    Set n(7) = tv.Nodes.Add
    n(7).Text = mnuDesktopHDT.Caption
    n(7).key = "DateTimeHide"
    n(7).Checked = mnuDesktopHDT.Checked
    
    Call hidetext
    msg(3).Visible = True
        
End Sub

Private Sub Desk_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
     Call hideShape
     Shape1(3).Visible = True
End Sub

Private Sub Desk_MouseOut()
   Call hideShape
End Sub

Private Sub Disk_Click()
    Label3.Caption = "Properties : Disk"
    tv.Nodes.Clear
    ReDim n(28) As Node
    Set n(1) = tv.Nodes.Add
    n(1).Text = "HideAllDrives"
    n(1).key = "HAD"
    n(1).Checked = mw.mnuDiskHideAll.Checked
    
    Set n(2) = tv.Nodes.Add
    n(2).Text = "RestrictAccess"
    n(2).key = "ResAc"
    n(2).Checked = False
    Dim i%
    For i = 3 To 28
      Set n(i) = tv.Nodes.Add("ResAc", tvwChild, Chr(i + 62) & ":", Chr(i + 62))
      n(i).Checked = mnuDiskRestrictAccessA.Item(i).Checked
    Next
    
    Call hidetext
    msg(0).Visible = True
    
End Sub

Private Sub Disk_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
     Call hideShape
     Shape1(0).Visible = True
End Sub

Private Sub Disk_MouseOut()
    Call hideShape
End Sub

Private Sub Exp_Click()
    Label3.Caption = "Properties : Explorer"
    tv.Nodes.Clear
    ReDim n(6) As Node
    Set n(1) = tv.Nodes.Add
    n(1).Text = mnuExplorerDRT.Caption
    n(1).key = "DRT"
    n(1).Checked = mnuExplorerDRT.Checked
    Set n(2) = tv.Nodes.Add
    n(2).Text = mnuExplorerNoFile.Caption
    n(2).key = "NOFILE"
    n(2).Checked = mnuExplorerNoFile.Checked
    
    Dim i%
    For i = 26 To mnuExpSet.Count + 25
        Set n(i - 23) = tv.Nodes.Add
        n(i - 23).Text = mnuExpSet(i).Caption
        n(i - 23).Checked = mnuExpSet(i).Checked
        n(i - 23).key = regKeys(i)
   Next
    
    Call hidetext
    msg(2).Visible = True
End Sub

Private Sub Exp_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call hideShape
    Shape1.Item(2).Visible = True
End Sub

Private Sub Exp_MouseOut()
   Call hideShape
End Sub

Private Sub Form_Load()
  Me.Top = 950
  Me.Left = 1200
  NID.cbSize = Len(NID)
  NID.hwnd = Me.hwnd
  NID.hIcon = Me.Icon
  NID.uId = vbNull
  NID.szTip = ttip & vbNullChar
  NID.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
  NID.uCallBackMessage = WM_MOUSEMOVE
  'load icon in tray
  Call Shell_NotifyIcon(NIM_ADD, NID)
  'initializing all message text boxes.
  msg(7).Text = "Click on a button to have a list of operations." & vbCrLf & _
  "Clicking on a node on the tree will perform that operation." & vbCrLf & _
  "Unchecking will do the reverse."
  msg(7).Visible = True
  
  msg(0).Text = "HideAllDrives : Clicking on it will hide all the drives on your computer." & _
                vbCrLf & vbCrLf & "RestrictAccess : Allows you to restrict access to " _
                & "drive/drives from A: to Z:.Just click on a combination of drives that you do not want anyone to access and click APPLY CHANGES." _
                & vbCrLf & vbCrLf & "Requirement : LOGOFF"
  
  msg(1).Text = "Alphabetic Start Menu : Arranges the start menu alphabetically." _
   & vbCrLf & "Requirement : LOGOFF" _
   & vbCrLf & vbCrLf _
   & "Clear Recent Docs On Exit : Clears the documents menu when you shut down/restart/logoff." _
   & vbCrLf & vbCrLf _
   & "Disable Drag-Drop : Restrict a user from dragging-dropping inside the start menu." _
   & vbCrLf & "Requirement : LOGOFF" _
   & vbCrLf & vbCrLf _
   & "No Shut Down : Removes the shut down option from the start menu.BE CAREFUL WHILE DOING THIS.YOU MUST USE THIS S/W TO SHUT DOWN OR ENABLE AGAIN THE OPTION." _
   & vbCrLf & "Requirement : LOGOFF" _
   & vbCrLf & vbCrLf _
   & "No Favorite Menu : Removes the favorite option from the start menu." _
   & vbCrLf & "Requirement : LOGOFF" _
   & vbCrLf & vbCrLf _
   & "No Folder Option : Removes the folder options from Explorer->tools in ME. or Explorer->view in 95/98." _
   & "Same as Start->ControlPanel->Folder Options." & vbCrLf & "Requirement : LOGOFF" _
   & vbCrLf & vbCrLf _
   & "No Logoff : Removes log off button from the start menu." _
   & vbCrLf & "Requirement : LOGOFF" _
   & vbCrLf & vbCrLf _
   & "No Recent Docs Menu : Removes documents button from the start menu." _
   & vbCrLf & "Requirement : LOGOFF" _
   & vbCrLf & vbCrLf _
   & "No Run : Removes Run button from the start menu." _
   & vbCrLf & "Requirement : LOGOFF"
   msg(1).Text = msg(1).Text & vbCrLf & vbCrLf & "No Search : Removes Search button from the start menu." _
   & vbCrLf & "Requirement : LOGOFF" _
   & vbCrLf & vbCrLf _
   & "No Taskbar Context Menu : Disables right clicking on the taskbar and start menu." _
   & vbCrLf & "Requirement : REFRESH" _
   & vbCrLf & vbCrLf _
   & "No Windows Update : Removes windows update button from start menu." _
   & vbCrLf & "Requirement : LOGOFF" & vbCrLf & vbCrLf _
   & "Hide Start Button : Hides the start button. Use this with Disable ALT+CTL+DEL,ALT+TAB," _
   & "START BUTTON to completely disable access to start menu. Also use HIDE DESKTOP" _
   & " to prevent user from accessing anything on the computer except the Taskbar, leaving with quick launch bar items to be accessed only." _
   & vbCrLf & "Requirement : REFRESH"
  msg(2).Text = "Disable Registry Tools : Allows you to disable registry opening/manipulating tools like REGEDIT." _
    & vbCrLf & "Requirement : LOGOFF" _
    & vbCrLf & vbCrLf _
    & "No File Menu : Allows you to remove the file menu from Windows/Intenet Explorer." _
    & vbCrLf & "Requirement : LOGOFF" & vbCrLf & vbCrLf _
    & "Show Hidden Files : Shows all the hidden files in My Computer." _
    & vbCrLf & "Requirement : REFRESH" & vbCrLf & vbCrLf _
    & "Show File Extention For Known File Types : Shows a file extention in known file types also." _
    & vbCrLf & "Requirement : REFRESH" & vbCrLf & vbCrLf _
    & "Show Operating System Files : Shows all the hidden system files used by Windows. To view these files first " _
    & "enable Show Hidden Files. DELETING OR MODIFYING THESE FILES ARE HIGHLY UNRECOMMENDED." _
    & vbCrLf & "Requirement : REFRESH" & vbCrLf & vbCrLf _
    & "Hide Tips On Items : Hides tips displayed on various items such as My Computer." _
    & vbCrLf & "Requirement : REFRESH"
    
  msg(3).Text = "Disable ALT+CTL+DEL,ALT+TAB, START BUTTON : Disables above said things." _
    & vbCrLf & "DO NOT USE THIS SETTING WITH HIDING TASKBAR AS IT WILL LEAVE YOU WITH NO OPTION TO SHUT DOWN" _
    & " APART FROM TURNING OFF THE CPU. AS ALT+TAB WILL BE DISABLED YOU CAN NO LONGER SWITCH BETWEEN APPLICATIONS." _
    & vbCrLf & "(Provided you close this software too.)" _
    & vbCrLf & "Requirement : REFRESH" & vbCrLf & vbCrLf _
    & "Disable Right Click : Disables right clicking on desktop. Allows you to prevent users from changing background etc." _
    & " If he/she does not the control panel way!!!!!!" & vbCrLf & "Requirement : LOGOFF" & vbCrLf & vbCrLf _
    & "Hide Desktop : Hide all icons on the desktop. Allows you to prevent user from gaining access to My Computer or deleting a icon from desktop." _
    & "This coupled with DISABLED RIGHT CLICK ON TASKBAR allows stop access to anything." _
    & "(Taskbar context menu can be disabled by START MENU->Disable Taskbar Context Menu(in this s/w))." _
    & vbCrLf & "Requirement : REFRESH" & vbCrLf & vbCrLf _
    & "Hide Taskbar : Hides the taskbar. DO NOT USE THIS WITH Hide Desktop,Disable ALT+CTL+DEL AND CLOSE THIS SOFTWARE AS YOU WILL HAVE NO OPTION TO SHUT DOWN" _
    & " AND TO OPEN ANY PROGRAM AFTERWARDS." & vbCrLf & "Requirement : REFRESH" & vbCrLf & vbCrLf _
    & "Empty Recycle Bin : Empties the recycle bin. It will prompt a message if recycle bin is not empty else not." _
    & vbCrLf & "Requirement : REFRESH" & vbCrLf & vbCrLf _
    & "No Network Neighbourhood Icon : Hides the network neighbourhood icon on desktop." & vbCrLf _
    & "Requirement : LOGOFF/RESTART" & vbCrLf & vbCrLf _
    & "Hide Date Time and Icon Tray : Hides the date time text and other icons in that tray." _
    & vbCrLf & "Requirement : REFRESH"
  
  msg(4).Text = "Disable Password Control panel : Disables the password properties page in Control Panel." _
   & " When disabled prompts a message." & vbCrLf & "Requirement : REFRESH" _
   & vbCrLf & vbCrLf _
   & "Hide Apperence Page : Hides apperence page in Control Panel->Display.(RightClick On Desktop->Properties->Apperence)" _
   & vbCrLf & "Requirement : REFRESH" & vbCrLf & vbCrLf _
   & "Hide Background page : Hides background page in Control Panel->Display.(RightClick On Desktop->Properties->Background)" _
   & vbCrLf & "Requirement : REFRESH" & vbCrLf & vbCrLf _
   & "Hide Device Manager Page : Hides the device manager page in Control Panel->System.(RightClick On MyComputer->Properties->DeviceManager)" _
   & vbCrLf & "Requirement : REFRESH" & vbCrLf & vbCrLf _
   & "Disable Display Control Panel : Disables the display page in Control Panel.(RightClick On Desktop->Properties)" _
   & " When disabled prompts a message." & vbCrLf & "Requirement : REFRESH" & vbCrLf & vbCrLf _
   & "Hide File System Page : Hides File System Page in Control Panel->System->Performance.(RightClick On MyComputer->Properties->Performance->FileSystem)" _
   & vbCrLf & "Once disabled no one can change your COMPUTER'S ROLE, SYSTEM RESTORE setting(in Windows ME only) etc." _
   & vbCrLf & "Requirement : REFRESH" & vbCrLf & vbCrLf _
   & "Hide Hardware Profile Page : Hides Hardware Profile Page in Control Panel->System->Hardware Profile.(RightClick On MyComputer->Properties->Hardware Profile)" _
   & vbCrLf & "Requirement : REFRESH" & vbCrLf & vbCrLf _
   & "Hide Password Change Page : Hide Password Change Page in Control Panel->Password." _
   & vbCrLf & "Requirement : REFRESH" & vbCrLf & vbCrLf _
   & "Hide Remote Admin Page : Hides Remote Admin Page in Control Panel->Password." & vbCrLf & "Requirement : REFRESH" & vbCrLf & vbCrLf _
   & "Hide Setting Page : Hides Settings,Web,Effects Pages in Control Panel->Display.(RightClick On Desktop->Properties)" _
   & " Allows you to prevent the user form changing Screen Resolution,Icons on desktop,Icon size etc." _
   & vbCrLf & "Requirement : REFRESH" & vbCrLf & vbCrLf _
   & "Hide Screen Saver Page : Hides Screen Saver Page in Control Panel->Display.(RightClick On Desktop->Properties)" _
   & vbCrLf & "Requirement : REFRESH" & vbCrLf & vbCrLf _
   & "Disable User Profile Page : Disables User Profile Page in Control Panel. Prevents user from changing user profiles."
   msg(4).Text = msg(4).Text & "Prompts a message when disabled." & vbCrLf & "Requirement : REFRESH" & vbCrLf & vbCrLf _
   & "Hide Virtual Memory Page : Hide Virtual Memory Page in Control Panel->System->Performance.(RightClick On MyComputer->Properties->Performance)" _
   & vbCrLf & "Requirement : REFRESH"
  msg(5).Text = "Disable Close : Disables Close Button in Internet Explorer. When disabled Prompts a message and will not allow user to close IE." _
  & vbCrLf & "While testing use ALT+CTL+DEL/LOGOFF/RESTART to close IE or keep this software running to enable it again." & vbCrLf & "Requirement : REFRESH" _
  & vbCrLf & vbCrLf _
  & "Disable Right Click : Disables right clicking on a page in Internet Explorer. Coupled with Disabled SaveAs button this setting allows you to stop " _
  & "user from copying images etc to hard disk to some degree." & vbCrLf & "Requirement : REFRESH" & vbCrLf & vbCrLf _
  & "Disable Tools/Internet Options menu : Disables the Internet Options button in Tools menu. Prompts a message when disabled." _
  & " Allows you to stop the user form changing various settings like Histrory,Cookie,Default Page,Restrictions etc." _
  & vbCrLf & "Requirement : REFRESH/LOGOFF" & vbCrLf & vbCrLf _
  & "Disable Save As : Hides SaveAs button in File menu. Prevents user from saving displayed page in Internet Explorer." _
  & vbCrLf & "Requirement : REFRESH" & vbCrLf & vbCrLf _
  & "Disable Favourites : Disables favourites menu in Internet Explorer." & vbCrLf & "Requirement : REFRESH" & vbCrLf & vbCrLf _
  & "Disable File/New : Disables New button in File menu. Prompts a message when disabled." _
  & vbCrLf & "Requirement : REFRESH" & vbCrLf & vbCrLf _
  & "Disable File/Open : Disables Open button in File menu. Prompts a message when disabled." _
  & vbCrLf & "Requirement : REFRESH" & vbCrLf & vbCrLf _
  & "Disable Find Files : Disables searching of files in Internet Explorer." _
  & vbCrLf & "Requirement : REFRESH" & vbCrLf & vbCrLf _
  & "Disable Download Directory Option : Disables the download directory option in Internet Explorer. Propmts a message when disabled. " _
  & vbCrLf & "Requirement : REFRESH/LOGOFF" & vbCrLf & vbCrLf _
  & "Disable Full Screen View Option : Disables the Full Screen button in View menu." _
  & vbCrLf & "Requirement : REFRESH"
  msg(6).Text = "Hide Network Security Page : Hides the network security page." _
  & vbCrLf & "Requirement : LOGOFF/RESTART" & vbCrLf & vbCrLf _
  & "Hide/Disable Network in Control Panel : Hides or disables the network option in control panel." _
  & vbCrLf & "Requirement : REFRESH" & vbCrLf & vbCrLf _
  & "Hide Identification Page : Hides network idetification page." _
  & vbCrLf & "Requirement : LOGOFF/RESTART" & vbCrLf & vbCrLf _
  & "Disable File Sharing Controls : Disables all the file sharing controls in network." _
  & vbCrLf & "Requirement : LOGOFF/RESTART" & vbCrLf & vbCrLf _
  & "Disable Print Sharing Controls : Disables all print sharing controls. " _
  & vbCrLf & "Requirement : LOGOFF/RESTART"
  
  'storing values that is to be written in registry for each drive in tag
  'property.
  For i = 3 To 28
      mnuDiskRestrictAccessA.Item(i).Tag = 2 ^ (i - 3)
  Next

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim msg As Long
    
    msg = x / Screen.TwipsPerPixelX
    
    Select Case msg
        Case WM_RBUTTONDOWN:
                               If Me.WindowState = 1 Then
                                  PopupMenu mnutb
                                  mnuMaxi.Enabled = True
                               End If
                            
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Shell_NotifyIcon NIM_DELETE, NID
   End
End Sub

Private Sub LaVolpeButton1_Click()
  Label3.Caption = "Properties : Internet Explorer"
  
  tv.Nodes.Clear
  ReDim n(10) As Node
  Dim i%
  For i = 31 To mnuIESet.Count + 30
       Set n(i - 30) = tv.Nodes.Add
       n(i - 30).Text = mnuIESet(i).Caption
       n(i - 30).Checked = mnuIESet(i).Checked
       n(i - 30).key = regKeys(i)
  Next
  
  Call hidetext
  msg(5).Visible = True
  
End Sub

Private Sub LaVolpeButton1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  Call hideShape
  Shape1(5).Visible = True
End Sub

Private Sub LaVolpeButton1_MouseOut()
  Call hideShape
End Sub

Private Sub LaVolpeButton2_Click()
  Label3.Caption = "Properties : Network"
  
  tv.Nodes.Clear
  ReDim n(5) As Node
  Dim i%
  For i = 41 To mnuNetSet.Count + 40
       Set n(i - 40) = tv.Nodes.Add
       n(i - 40).Text = mnuNetSet(i).Caption
       n(i - 40).Checked = mnuNetSet(i).Checked
       n(i - 40).key = regKeys(i)
  Next
  
  Call hidetext
  msg(6).Visible = True

End Sub

Private Sub LaVolpeButton2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  Call hideShape
  Shape1(6).Visible = True
End Sub

Private Sub LaVolpeButton2_MouseOut()
  Call hideShape
End Sub

Private Sub lo_Click()
  lngresult = ExitWindowsEx(EWX_LOGOFF, 0&)
End Sub

Private Sub mice_Click()
   Me.Hide
   frmMisc.Show
End Sub

Private Sub mnuDesktopDACD_Click()
   If mnuDesktopDACD.Caption = "Disable &ALT+CTRL+DEL, ALT+TAB, START BUTTON" Then
      Call DisableCtrlAltDelete(True)
      mnuDesktopDACD.Caption = "Enable &ALT+CTRL+DEL, ALT+TAB, START BUTTON"
   Else
      Call DisableCtrlAltDelete(False)
      mnuDesktopDACD.Caption = "Disable &ALT+CTRL+DEL, ALT+TAB, START BUTTON"
   End If
End Sub

Private Sub mnuDesktopDRC_Click()
 If mnuDesktopDRC.Checked = False Then
  RegRW str1, "NoViewContextMenu", 1, "REG_DWORD", mnuDesktopDRC, True
 Else
  RegRW str1, "NoViewContextMenu", 0, "REG_DWORD", mnuDesktopDRC, False
 End If
End Sub

Private Sub mnuDesktopHide_Click()
  Dim hwnd As Long
  On Error Resume Next
  
  If mnuDesktopHide.Caption = "Hide Desktop" Then
     hwnd = FindShellWindow()
     If hwnd <> 0 Then
          Call HideShowWindow(hwnd, b)
     End If
  End If
  
End Sub

Private Sub mnuDiskHideAll_Click()
  If mnuDiskHideAll.Checked = False Then
    RegRW str1, "NoDrives", 67108863, "REG_DWORD", mnuDiskHideAll, True
  Else
    RegRW str1, "NoDrives", 0, "REG_DWORD", mnuDiskHideAll, False
  End If
End Sub

Private Sub mnuStartAlpha_Click()
  wscr.regdelete "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\MenuOrder\Start Menu\ORDER"
End Sub

Private Sub mnuMaxi_Click()
   'called when right clicked on tray icon and Open is clicked.
   If bca = True Then
     frmTake.Show
   Else
     mw.Show
     mw.WindowState = 0
     mw.Top = 950
     mw.Left = 1200
   End If
End Sub

Private Sub mnurck_Click()
  'called when right clicked on tray icon and Exit is clicked.
   If bca = True Then
     frmTake.Show
     mca = True
   Else
     End
   End If
End Sub

Private Sub more_Click()
  mw.Hide
  frmMore.Show
End Sub

Private Sub rest_Click()
  lngresult = ExitWindowsEx(EWX_REBOOT, 0&)
End Sub

Private Sub sd_Click()
   lngresult = ExitWindowsEx(EWX_SHUTDOWN, 0&)
End Sub

Private Sub settings_Click()
   Me.Hide
   frmset.Show
End Sub

Private Sub SM_Click()
   Label3.Caption = "Properties : Start Menu"
   tv.Nodes.Clear
   ReDim n(13) As Node
   Dim i%
   For i = 1 To mnuStart.Count
        Set n(i) = tv.Nodes.Add
        n(i).Text = mnuStart(i).Caption
        n(i).Checked = mnuStart(i).Checked
        n(i).key = regKeys(i)
   Next
   
   Set n(13) = tv.Nodes.Add
   n(13).Text = mnuHideSB.Caption
   n(13).Checked = mnuHideSB.Checked
   n(13).key = "StartButtonHide"
   
   Call hidetext
   msg(1).Visible = True
   
End Sub

Private Sub SM_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   Call hideShape
   Shape1(1).Visible = True
End Sub

Private Sub SM_MouseOut()
  Call hideShape
End Sub
Private Sub hidetext()
   For i = 0 To msg.Count - 1
     msg(i).Visible = False
   Next
End Sub

Public Sub Timer1_Timer()
   Dim lngTickCount As Long
   lngTickCount = GetTickCount
   'CALCULATE MINUTES
   ttip = CStr(Round((lngTickCount / 1000 / 60))) & " Minutes in Windows"
   Call frmset.newtip
End Sub

Private Sub Timer2_Timer()
  'this timer is used for hiding the mw form when it is minimized.
  If Me.WindowState = 1 Then Me.Hide
End Sub
