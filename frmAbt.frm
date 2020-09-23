VERSION 5.00
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVBUTTONS.OCX"
Begin VB.Form frmAbt 
   BackColor       =   &H00F3CE9C&
   Caption         =   "About WinSecure Ver 1.0.0"
   ClientHeight    =   3675
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5355
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmAbt.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3675
   ScaleWidth      =   5355
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frame 
      BackColor       =   &H00F3CE9C&
      Caption         =   "Developed By"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   1260
      TabIndex        =   6
      Top             =   2160
      Width           =   3120
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "rbhatt123@rediffmail.com"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   210
         Left            =   135
         TabIndex        =   9
         Top             =   1125
         Width           =   2415
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "G-307, srinandnagar-1, Vejalpur , Ahmadabad, Gujarat, India, Pin-380051"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   645
         Left            =   135
         TabIndex        =   8
         Top             =   450
         Width           =   2625
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ravi Bhatt"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   210
         Left            =   135
         TabIndex        =   7
         Top             =   225
         Width           =   960
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   225
      Picture         =   "frmAbt.frx":0442
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   0
      Top             =   405
      Width           =   480
   End
   Begin LVbuttons.LaVolpeButton butok 
      Height          =   495
      Left            =   4500
      TabIndex        =   2
      Top             =   3105
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "Ok"
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
      MICON           =   "frmAbt.frx":0884
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
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAbt.frx":08A0
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   960
      Left            =   1260
      TabIndex        =   5
      Top             =   1035
      Width           =   4065
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Version 1.0.0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   345
      Left            =   3105
      TabIndex        =   4
      Top             =   585
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackColor       =   &H00F3CE9C&
      Caption         =   "WinSecure"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   32.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   690
      Left            =   1260
      TabIndex        =   3
      Top             =   90
      Width           =   4065
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      Caption         =   "WinSecure   ver 1.0.0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   420
      Left            =   90
      TabIndex        =   1
      Top             =   1530
      Width           =   1005
   End
   Begin VB.Shape Sha 
      BorderColor     =   &H80000004&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   4110
      Left            =   0
      Top             =   0
      Width           =   1185
   End
End
Attribute VB_Name = "frmAbt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub butok_Click()
  Unload Me
  mw.Show
End Sub
