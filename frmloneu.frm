VERSION 5.00
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVBUTTONS.OCX"
Begin VB.Form frmloneu 
   BackColor       =   &H00F3CE9C&
   Caption         =   "Change Passowrd"
   ClientHeight    =   2220
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4965
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmloneu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2220
   ScaleWidth      =   4965
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox takingC 
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   11.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   2025
      PasswordChar    =   "l"
      TabIndex        =   3
      Top             =   1080
      Width           =   2715
   End
   Begin VB.TextBox takingN 
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   11.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   2025
      PasswordChar    =   "l"
      TabIndex        =   2
      Top             =   585
      Width           =   2715
   End
   Begin VB.TextBox takingO 
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   11.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   2025
      PasswordChar    =   "l"
      TabIndex        =   0
      Top             =   135
      Width           =   2715
   End
   Begin LVbuttons.LaVolpeButton butCng 
      Height          =   495
      Left            =   1800
      TabIndex        =   4
      Top             =   1620
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "Make Change"
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
      MICON           =   "frmloneu.frx":0442
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
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00F3CE9C&
      BackStyle       =   0  'Transparent
      Caption         =   "Confirm Password"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   210
      Left            =   90
      TabIndex        =   6
      Top             =   1125
      Width           =   1710
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00F3CE9C&
      BackStyle       =   0  'Transparent
      Caption         =   "New Password"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   210
      Left            =   90
      TabIndex        =   5
      Top             =   630
      Width           =   1410
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00F3CE9C&
      BackStyle       =   0  'Transparent
      Caption         =   "Old Password"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   210
      Left            =   90
      TabIndex        =   1
      Top             =   180
      Width           =   1305
   End
End
Attribute VB_Name = "frmloneu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As String
Dim b As Integer

'nothing to explain in this coding.

Private Sub butCng_Click()
  If Not (takingO.Text = "" And takingN.Text = "" And takingC.Text = "") Then
    takingC.Text = tore(takingC.Text) 'encrpt entered text
    wscr.regwrite "HKCR\CLSID\{DBE209UY-093F-R97E-VCTEDER12AWN}\InprocServer32\ThreadingModel", takingC.Text, "REG_SZ"
    MsgBox "Successful...", vbOKOnly, "Done"
  End If
End Sub
Private Sub Form_Load()
  a = wscr.regread("HKCR\CLSID\{DBE209UY-093F-R97E-VCTEDER12AWN}\InprocServer32\ThreadingModel")
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Unload Me
  frmset.Show
End Sub

Private Sub takingC_Validate(Cancel As Boolean)
  If takingC.Text = "" Then
    MsgBox "Type Password....", vbInformation, "Error"
  ElseIf takingC.Text <> takingN.Text Then
    b = MsgBox("Password does not match..." & vbCrLf & "Set New password Again ? ", vbYesNo, "Error")
    If b = vbYes Then
      Cancel = False
    Else
     Cancel = True
    End If
  End If
End Sub
Private Sub takingN_Validate(Cancel As Boolean)
  If takingO.Text = "" Then
   MsgBox "Type Password....", vbInformation, "Error"
  End If
End Sub
Private Sub takingO_Validate(Cancel As Boolean)
  If takingO.Text = "" Then
   MsgBox "Type Password....", vbInformation, "Error"
  ElseIf tore(takingO.Text) <> a Then
    MsgBox "Invalid Old Password ... ", vbCritical, "Wrong password"
    Cancel = True
  End If
End Sub
