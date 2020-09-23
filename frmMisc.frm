VERSION 5.00
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVBUTTONS.OCX"
Begin VB.Form frmMisc 
   BackColor       =   &H00F3CE9C&
   Caption         =   "WinSecure Miscellaneous Settings"
   ClientHeight    =   6165
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5970
   Icon            =   "frmMisc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6165
   ScaleWidth      =   5970
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00F3CE9C&
      Caption         =   "Description"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1995
      Left            =   1710
      TabIndex        =   7
      Top             =   3960
      Width           =   4020
      Begin VB.TextBox txtdes 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1770
         Left            =   45
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   180
         Width           =   3930
      End
   End
   Begin LVbuttons.LaVolpeButton butMsgLog 
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   495
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "Logon Message"
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
      MICON           =   "frmMisc.frx":0442
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
   Begin LVbuttons.LaVolpeButton AC 
      Height          =   495
      Left            =   3060
      TabIndex        =   6
      Top             =   3465
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
      MICON           =   "frmMisc.frx":045E
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
   Begin LVbuttons.LaVolpeButton butMisc 
      Height          =   495
      Left            =   360
      TabIndex        =   10
      Top             =   1170
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "Mix Settings"
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
      MICON           =   "frmMisc.frx":047A
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
   Begin LVbuttons.LaVolpeButton butReBin 
      Height          =   495
      Left            =   360
      TabIndex        =   12
      Top             =   1845
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "Recycle Bin"
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
      MICON           =   "frmMisc.frx":0496
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
   Begin LVbuttons.LaVolpeButton butSetMIO 
      Height          =   495
      Left            =   360
      TabIndex        =   25
      Top             =   2520
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "Set Owner Name etc"
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
      MICON           =   "frmMisc.frx":04B2
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
   Begin LVbuttons.LaVolpeButton butClear 
      Height          =   495
      Left            =   360
      TabIndex        =   41
      Top             =   3195
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "Clear"
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
      MICON           =   "frmMisc.frx":04CE
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
   Begin VB.Frame containers 
      BackColor       =   &H00F3CE9C&
      Caption         =   "Nothing to show"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3075
      Index           =   0
      Left            =   1710
      TabIndex        =   9
      Top             =   270
      Width           =   4020
   End
   Begin VB.Frame containers 
      BackColor       =   &H00F3CE9C&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3075
      Index           =   1
      Left            =   1710
      TabIndex        =   1
      Top             =   270
      Visible         =   0   'False
      Width           =   4020
      Begin VB.TextBox txtLogMes 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   180
         TabIndex        =   5
         Text            =   "Type in your message here."
         Top             =   1305
         Width           =   3570
      End
      Begin VB.TextBox txtLogTit 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   180
         TabIndex        =   4
         Text            =   "Your logon title goes here."
         Top             =   495
         Width           =   3570
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00F3CE9C&
         Caption         =   "Logon Message:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   180
         TabIndex        =   3
         Top             =   990
         Width           =   1380
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00F3CE9C&
         Caption         =   "Logon Title:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   180
         TabIndex        =   2
         Top             =   225
         Width           =   1005
      End
   End
   Begin VB.Frame containers 
      BackColor       =   &H00F3CE9C&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3075
      Index           =   3
      Left            =   1710
      TabIndex        =   13
      Top             =   270
      Width           =   4020
      Begin VB.CheckBox chkReBin 
         BackColor       =   &H00F3CE9C&
         Caption         =   "Set Recycle Bin to always delete (Same as Shift + Delete)"
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
         Index           =   0
         Left            =   225
         TabIndex        =   18
         Tag             =   "RBAD"
         Top             =   270
         Width           =   3435
      End
      Begin VB.CheckBox chkReBin 
         BackColor       =   &H00F3CE9C&
         Caption         =   "Add rename to right click menu"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   1
         Left            =   225
         TabIndex        =   17
         Tag             =   "RBR"
         Top             =   690
         Width           =   3075
      End
      Begin VB.CheckBox chkReBin 
         BackColor       =   &H00F3CE9C&
         Caption         =   "Add delete to right click menu"
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
         Index           =   2
         Left            =   225
         TabIndex        =   16
         Tag             =   "RBD"
         Top             =   1155
         Width           =   3075
      End
      Begin VB.CheckBox chkReBin 
         BackColor       =   &H00F3CE9C&
         Caption         =   "Add rename and delete to right click menu"
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
         Index           =   3
         Left            =   225
         TabIndex        =   15
         Tag             =   "RBRD"
         Top             =   1590
         Width           =   3750
      End
      Begin VB.CheckBox chkReBin 
         BackColor       =   &H00F3CE9C&
         Caption         =   "Add cut,copy,paste to right click menu"
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
         Index           =   4
         Left            =   225
         TabIndex        =   14
         Tag             =   "RBCCP"
         Top             =   2010
         Width           =   3750
      End
      Begin VB.CheckBox chkReBin 
         BackColor       =   &H00F3CE9C&
         Caption         =   "Add rename,delete,cut,copy,paste to right click menu"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   5
         Left            =   225
         TabIndex        =   23
         Tag             =   "RBWD"
         Top             =   2430
         Width           =   3660
      End
   End
   Begin VB.Frame containers 
      BackColor       =   &H00F3CE9C&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3075
      Index           =   2
      Left            =   1710
      TabIndex        =   11
      Top             =   270
      Width           =   4020
      Begin VB.CheckBox chkMixSet 
         BackColor       =   &H00F3CE9C&
         Caption         =   "Create Taskbar and Start menu properties Shortcut on Desktop"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   6
         Left            =   270
         TabIndex        =   34
         Tag             =   "TBSM"
         Top             =   2565
         Width           =   3345
      End
      Begin VB.CheckBox chkMixSet 
         BackColor       =   &H00F3CE9C&
         Caption         =   "Create Folder Options Shortcut on Desktop"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   5
         Left            =   270
         TabIndex        =   33
         Tag             =   "FOSC"
         Top             =   2105
         Width           =   3345
      End
      Begin VB.CheckBox chkMixSet 
         BackColor       =   &H00F3CE9C&
         Caption         =   "Create Shut Down Shortcut on Desktop"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   4
         Left            =   270
         TabIndex        =   32
         Tag             =   "SCSD"
         Top             =   1648
         Width           =   3660
      End
      Begin VB.CheckBox chkMixSet 
         BackColor       =   &H00F3CE9C&
         Caption         =   "Remove Shortcut Arrow"
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
         Index           =   0
         Left            =   270
         TabIndex        =   22
         Tag             =   "NSAR"
         Top             =   180
         Width           =   2355
      End
      Begin VB.CheckBox chkMixSet 
         BackColor       =   &H00F3CE9C&
         Caption         =   "Remove tips from Min/max/Close buttons on window"
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
         Index           =   1
         Left            =   270
         TabIndex        =   21
         Tag             =   "MMC"
         Top             =   547
         Width           =   3480
      End
      Begin VB.CheckBox chkMixSet 
         BackColor       =   &H00F3CE9C&
         Caption         =   "Disable MS-DOS prompt"
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
         Index           =   2
         Left            =   270
         TabIndex        =   20
         Tag             =   "DDOS"
         Top             =   914
         Width           =   2580
      End
      Begin VB.CheckBox chkMixSet 
         BackColor       =   &H00F3CE9C&
         Caption         =   "Disable Single-Mode MS-DOS"
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
         Index           =   3
         Left            =   270
         TabIndex        =   19
         Tag             =   "DDRM"
         Top             =   1281
         Width           =   2895
      End
   End
   Begin VB.Frame containers 
      BackColor       =   &H00F3CE9C&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3075
      Index           =   4
      Left            =   1710
      TabIndex        =   24
      Top             =   270
      Width           =   4020
      Begin VB.TextBox txtCmp 
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
         Left            =   135
         TabIndex        =   42
         Top             =   1170
         Width           =   3750
      End
      Begin VB.TextBox txtIEDHP 
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
         Left            =   135
         TabIndex        =   31
         Top             =   2610
         Width           =   3750
      End
      Begin VB.TextBox txtWMP 
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
         Left            =   135
         TabIndex        =   29
         Top             =   1890
         Width           =   3750
      End
      Begin VB.TextBox txtOwn 
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
         Left            =   135
         TabIndex        =   26
         Top             =   450
         Width           =   3750
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00F3CE9C&
         Caption         =   "Registered Organization Name:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   135
         TabIndex        =   43
         Top             =   900
         Width           =   2685
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00F3CE9C&
         Caption         =   "Internet Explorer Default Home Page:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   135
         TabIndex        =   30
         Top             =   2340
         Width           =   3255
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00F3CE9C&
         Caption         =   "Windows Media Player Title:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   135
         TabIndex        =   28
         Top             =   1620
         Width           =   2400
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00F3CE9C&
         Caption         =   "Registered Owner Name:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   135
         TabIndex        =   27
         Top             =   180
         Width           =   2160
      End
   End
   Begin VB.Frame containers 
      BackColor       =   &H00F3CE9C&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3075
      Index           =   5
      Left            =   1710
      TabIndex        =   35
      Top             =   270
      Width           =   4020
      Begin VB.CheckBox chkCl 
         BackColor       =   &H00F3CE9C&
         Caption         =   "Clear Text Search"
         Height          =   285
         Index           =   4
         Left            =   225
         TabIndex        =   40
         Tag             =   "RBAD"
         Top             =   2520
         Width           =   1815
      End
      Begin VB.CheckBox chkCl 
         BackColor       =   &H00F3CE9C&
         Caption         =   "Clear Run Commands (Start->Run)"
         Height          =   330
         Index           =   0
         Left            =   225
         TabIndex        =   39
         Tag             =   "RBAD"
         Top             =   270
         Width           =   2850
      End
      Begin VB.CheckBox chkCl 
         BackColor       =   &H00F3CE9C&
         Caption         =   "Clear Found Files (Start->find)"
         Height          =   330
         Index           =   1
         Left            =   225
         TabIndex        =   38
         Tag             =   "RBAD"
         Top             =   810
         Width           =   2445
      End
      Begin VB.CheckBox chkCl 
         BackColor       =   &H00F3CE9C&
         Caption         =   "Clear Found Computers (Find->Computers)"
         Height          =   465
         Index           =   2
         Left            =   225
         TabIndex        =   37
         Tag             =   "RBAD"
         Top             =   1260
         Width           =   3390
      End
      Begin VB.CheckBox chkCl 
         BackColor       =   &H00F3CE9C&
         Caption         =   "Clear Recent Documents"
         Height          =   375
         Index           =   3
         Left            =   225
         TabIndex        =   36
         Tag             =   "RBAD"
         Top             =   1890
         Width           =   2130
      End
   End
   Begin VB.Shape Sha 
      BorderColor     =   &H80000004&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   6405
      Left            =   0
      Top             =   0
      Width           =   960
   End
End
Attribute VB_Name = "frmMisc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Dim vis As Integer
Dim val As String
Dim fifth As String
Dim s As String
Dim written As Boolean
Dim fso As FileSystemObject
Dim shell
Dim dp As String
Dim str


Private Sub AC_Click()
   On Error Resume Next
   Select Case vis
     Case 1: 'write to registry
              wscr.regwrite str3 & "\CurrentVersion\Winlogon\LegalNoticeCaption", txtLogTit, "REG_SZ"
              wscr.regwrite str3 & "\CurrentVersion\Winlogon\LegalNoticeText", txtLogMes, "REG_SZ"
     Case 2:  'mix settings button
              '2nd button has several settings so loop and find.
              'check boxes tag property is used for storing values in reg.
              'that will tell whether a setting is enabled.
              For i = 0 To chkMixSet.UBound
                Select Case i
                   Case 0:
                            If chkMixSet.Item(i).Value = 1 Then
                              wscr.regdelete "HKEY_CLASSES_ROOT\lnkfile\IsShortcut"
                              wscr.regdelete "HKEY_CLASSES_ROOT\piffile\IsShortcut"
                              wscr.regwrite str3 & "\users\" & chkMixSet.Item(i).Tag, 1, "REG_DWORD"
                            Else
                              wscr.regwrite "HKEY_CLASSES_ROOT\lnkfile\IsShortcut", "", "REG_SZ"
                              wscr.regwrite "HKEY_CLASSES_ROOT\piffile\IsShortcut", "", "REG_SZ"
                              wscr.regwrite str3 & "\users\" & chkMixSet.Item(i).Tag, 0, "REG_DWORD"
                            End If
                   Case 1:
                            If chkMixSet.Item(i).Value = 1 Then
                              wscr.regwrite "HKEY_CURRENT_USER\Control Panel\Desktop\MinMaxClose", "1", "REG_SZ"
                              wscr.regwrite str3 & "\users\" & chkMixSet.Item(i).Tag, 1, "REG_DWORD"
                            Else
                              wscr.regwrite "HKEY_CURRENT_USER\Control Panel\Desktop\MinMaxClose", "0", "REG_SZ"
                              wscr.regwrite str3 & "\users\" & chkMixSet.Item(i).Tag, 0, "REG_DWORD"
                            End If
                  
                   Case 2:
                            If chkMixSet.Item(i).Value = 1 Then
                              wscr.regwrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\WinOldApp\Disabled", 1, "REG_DWORD"
                              wscr.regwrite str3 & "\users\" & chkMixSet.Item(i).Tag, 1, "REG_DWORD"
                            Else
                              wscr.regwrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\WinOldApp\Disabled", 0, "REG_DWORD"
                              wscr.regwrite str3 & "\users\" & chkMixSet.Item(i).Tag, 0, "REG_DWORD"
                            End If
                   Case 3:
                            If chkMixSet.Item(i).Value = 1 Then
                              wscr.regwrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\WinOldApp\NoRealMode", 1, "REG_DWORD"
                              wscr.regwrite str3 & "\users\" & chkMixSet.Item(i).Tag, 1, "REG_DWORD"
                            Else
                              wscr.regwrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\WinOldApp\NoRealMode", 0, "REG_DWORD"
                              wscr.regwrite str3 & "\users\" & chkMixSet.Item(i).Tag, 0, "REG_DWORD"
                            End If
                   Case 4:
                            If chkMixSet.Item(i).Value = 1 Then
                                 fso.CopyFile App.path & "\SD.lnk", dp & "\Shut Down.lnk", True
                                 wscr.regwrite str3 & "\users\" & chkMixSet.Item(i).Tag, 1, "REG_DWORD"
                                 SetAttr dp & "\Shut Down.lnk", vbNormal
                            Else
                                 fso.DeleteFile dp & "\Shut Down.lnk"
                                 wscr.regwrite str3 & "\users\" & chkMixSet.Item(i).Tag, 0, "REG_DWORD"
                            End If
                   Case 5:
                            If chkMixSet.Item(i).Value = 1 Then
                                 fso.CopyFile App.path & "\FO.lnk", dp & "\Folder Options.lnk", True
                                 wscr.regwrite str3 & "\users\" & chkMixSet.Item(i).Tag, 1, "REG_DWORD"
                                 SetAttr dp & "\Folder Options.lnk", vbNormal
                            Else
                                 fso.DeleteFile dp & "\Folder Options.lnk"
                                 wscr.regwrite str3 & "\users\" & chkMixSet.Item(i).Tag, 0, "REG_DWORD"
                            End If
                   Case 6:
                            If chkMixSet.Item(i).Value = 1 Then
                                 fso.CopyFile App.path & "\TS.lnk", dp & "\Taskbar and Startmenu.lnk", True
                                 wscr.regwrite str3 & "\users\" & chkMixSet.Item(i).Tag, 1, "REG_DWORD"
                                 SetAttr dp & "\Taskbar and Startmenu.lnk", vbNormal
                            Else
                                 fso.DeleteFile dp & "\Taskbar and Startmenu.lnk"
                                 wscr.regwrite str3 & "\users\" & chkMixSet.Item(i).Tag, 0, "REG_DWORD"
                            End If
                End Select
              Next
      Case 3:  'recycle bin button
               For i = 0 To chkReBin.UBound
                  Select Case i
                     Case 0:
                              If chkReBin.Item(i).Value = 1 Then
                                 wscr.regwrite str3 & "\CurrentVersion\Explorer\BitBucket\NukeOnDelete", 1, "REG_DWORD"
                                 wscr.regwrite str3 & "\users\" & chkReBin.Item(i).Tag, 1, "REG_DWORD"
                              Else
                                 wscr.regwrite str3 & "\CurrentVersion\Explorer\BitBucket\NukeOnDelete", 0, "REG_DWORD"
                                 wscr.regwrite str3 & "\users\" & chkReBin.Item(i).Tag, 0, "REG_DWORD"
                              End If
                     Case 1:
                              If chkReBin.Item(i).Value = 1 Then
                                 wscr.regwrite s, 536871248, "REG_BINARY"
                                 wscr.regwrite str3 & "\users\" & chkReBin.Item(i).Tag, 1, "REG_DWORD"
                                 written = True
                              Else
                                 wscr.regwrite s, 536871232, "REG_BINARY"
                                 wscr.regwrite str3 & "\users\" & chkReBin.Item(i).Tag, 0, "REG_DWORD"
                                 written = False
                              End If
                     Case 2:
                              If chkReBin.Item(i).Value = 1 Then
                                 wscr.regwrite s, 536871264, "REG_BINARY"
                                 wscr.regwrite str3 & "\users\" & chkReBin.Item(i).Tag, 1, "REG_DWORD"
                                 written = True
                              ElseIf written = False Then
                                 wscr.regwrite s, 536871232, "REG_BINARY"
                                 wscr.regwrite str3 & "\users\" & chkReBin.Item(i).Tag, 0, "REG_DWORD"
                                 written = False
                              End If
                     Case 3:
                              If chkReBin.Item(i).Value = 1 Then
                                 wscr.regwrite s, 536871280, "REG_BINARY"
                                 wscr.regwrite str3 & "\users\" & chkReBin.Item(i).Tag, 1, "REG_DWORD"
                                 written = True
                              ElseIf written = False Then
                                 wscr.regwrite s, 536871232, "REG_BINARY"
                                 wscr.regwrite str3 & "\users\" & chkReBin.Item(i).Tag, 0, "REG_DWORD"
                                 written = False
                              End If
                     Case 4:
                              If chkReBin.Item(i).Value = 1 Then
                                 wscr.regwrite s, 536871239, "REG_BINARY"
                                 wscr.regwrite str3 & "\users\" & chkReBin.Item(i).Tag, 1, "REG_DWORD"
                                 written = True
                              ElseIf written = False Then
                                 wscr.regwrite s, 536871232, "REG_BINARY"
                                 wscr.regwrite str3 & "\users\" & chkReBin.Item(i).Tag, 0, "REG_DWORD"
                                 written = False
                              End If
                     Case 5:
                              If chkReBin.Item(i).Value = 1 Then
                                 wscr.regwrite s, 536871223, "REG_BINARY"
                                 wscr.regwrite str3 & "\users\" & chkReBin.Item(i).Tag, 1, "REG_DWORD"
                              ElseIf written = False Then
                                 wscr.regwrite s, 536871232, "REG_BINARY"
                                 wscr.regwrite str3 & "\users\" & chkReBin.Item(i).Tag, 0, "REG_DWORD"
                              End If
                  End Select
               Next
      Case 4:   'set owner button
                wscr.regwrite str3 & "\CurrentVersion\RegisteredOwner", txtOwn, "REG_SZ"
                wscr.regwrite str3 & "\CurrentVersion\RegisteredOrganization", txtCmp, "REG_SZ"
                wscr.regwrite "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Main\Start Page", txtIEDHP, "REG_SZ"
                wscr.regwrite "HKEY_USERS\.DEFAULT\Software\Policies\Microsoft\WindowsMediaPlayer\TitleBar", txtWMP, "REG_SZ"
      Case 5:   'clear button
                Dim j As Integer
                For i = 0 To 4
                  If chkCl.Item(i).Value = 1 Then
                      If i = 4 Then
                        wscr.regdelete fifth & "ContainingTextMRU\"
                        wscr.regdelete fifth & "FilesNamedMRU\"
                      Else
                        wscr.regdelete val & str(i)
                      End If
                  End If
                Next
      
   End Select
End Sub

Private Sub butClear_Click()
   For i = 0 To containers.Count - 1 'hide all frames
     containers.Item(i).Visible = False
   Next
   vis = 5 'is only visible frame or say 5th button on form is clicked
           ' this done for combined apply changes button click.
   containers(5).Visible = True 'make visible
   txtdes.Text = "Clickin a check box above will clear the said item in " _
   & "your computer." & vbCrLf & vbCrLf & "Requirement : LOGOFF/RESTART"
End Sub

Private Sub butMisc_Click()
  For i = 0 To containers.Count - 1
     containers.Item(i).Visible = False
   Next
   vis = 2
   containers(2).Visible = True
   On Error Resume Next
   For i = 0 To chkMixSet.UBound
      chkMixSet.Item(i).Value = wscr.regread(str3 & "\users\" & chkMixSet.Item(i).Tag)
   Next
   txtdes.Text = "Remove Shortcut Arrow : Removes that small ugly arrow from every shortcut." _
   & vbCrLf & "Requirements : LOGOFF" & vbCrLf & vbCrLf _
   & "Remove tips from Min/Max/Close : Removes unneccessary tool tips from above said buttons." _
   & "This may not work in windows me." & vbCrLf & "Requirements : RESTART" & vbCrLf & vbCrLf _
   & "Disable MS-DOS Prompt : Disables MS-DOS prompt. When set user will not be able to goto DOS" _
   & " from Run in Start menu." & vbCrLf & "Requirements : REFRESH" & vbCrLf & vbCrLf _
   & "Disable Single-Mode MS-DOS : Removes single mode MS-DOS." & vbCrLf & "Requirements : REFRESH"

End Sub

Private Sub butMsgLog_Click()
   For i = 0 To containers.Count - 1
     containers.Item(i).Visible = False
   Next
   vis = 1
   containers(1).Visible = True
   On Error Resume Next
   'read values if they exist.
   txtLogTit.Text = wscr.regread(str3 & "\CurrentVersion\Winlogon\LegalNoticeCaption")
   txtLogMes.Text = wscr.regread(str3 & "\CurrentVersion\Winlogon\LegalNoticeText")
   
   txtdes.Text = "You can make windows prompt a message every time" & _
   " it logs on or restart. If you don't have this setting on previously " & _
   "the text boxes will show help text." & vbCrLf & "Type your message and give it a suitable title." & _
   "Make sure that you click Apply Changes."
End Sub

Private Sub butReBin_Click()
   For i = 0 To containers.Count - 1
     containers.Item(i).Visible = False
   Next
   vis = 3
   containers(3).Visible = True
   On Error Resume Next
   For i = 0 To chkReBin.UBound
      chkReBin.Item(i).Value = wscr.regread(str3 & "\users\" & chkReBin.Item(i).Tag)
   Next
   txtdes.Text = "Set Recycle Bin to always delete : Deletes items permanntely from your " _
   & "computer as you delete a particular item. Items are no longer moved to the recycle bin." _
   & " Use this setting with caution as items once deleted can not be recovered afterwards." _
   & vbCrLf & "Requirements : LOGOFF" & vbCrLf & vbCrLf _
   & "All other settings : Clear by their name." & vbCrLf & "Requirements : REFRESH"
End Sub

Private Sub butSetMIO_Click()
   For i = 0 To containers.Count - 1
     containers.Item(i).Visible = False
   Next
   vis = 4
   containers(4).Visible = True
   On Error Resume Next
   txtOwn.Text = wscr.regread(str3 & "\CurrentVersion\RegisteredOwner")
   txtCmp.Text = wscr.regread(str3 & "\CurrentVersion\RegisteredOrganization")
   txtIEDHP.Text = wscr.regread("HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Main\Start Page")
   txtWMP.Text = wscr.regread("HKEY_USERS\.DEFAULT\Software\Policies\Microsoft\WindowsMediaPlayer\TitleBar")
   
   txtdes.Text = "The text shown in the text boxes are current values for that particular setting." _
   & "You can change default home page by typing a URL and clicking on Apply Changes." _
   & "If a text box is showing nothing that means that there is no value for that setting." _
   & vbCrLf & "CHANGING YOUR REGISTERED USER NAME AND/OR ORGANIZATION NAME MAY CAUSE PROBLEMS." _
   & vbCrLf & "Requirements : REFRESH"
End Sub

Private Sub Form_Load()
  vis = 0 'nothing to show.
  written = False
  containers(0).Visible = True
  s = "HKCR\CLSID\{645FF040-5081-101B-9F08-00AA002F954E}\ShellFolder\Attributes"
  Set shell = CreateObject("WScript.Shell")
  Set fso = CreateObject("Scripting.FileSystemObject")
  dp = shell.SpecialFolders("Desktop")
  val = "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\"
  str = Array("RunMRU\", "Doc Find Spec MRU\", "FindComputerMRU\", "RecentDocs\")
  fifth = "HKCU\Software\Microsoft\Internet Explorer\Explorer Bars\{C4EE31F3-4768-11D2-BE5C-00A0C9A83DA1}\"
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set shell = Nothing
  Set fso = Nothing
  mw.Show
End Sub
