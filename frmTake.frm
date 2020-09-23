VERSION 5.00
Begin VB.Form frmTake 
   BackColor       =   &H00F3CE9C&
   ClientHeight    =   7755
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10995
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   7755
   ScaleWidth      =   10995
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox taking 
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   11.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   5040
      PasswordChar    =   "l"
      TabIndex        =   0
      Top             =   4230
      Width           =   2715
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "rbhatt123@rediffmail.com  Ph:+91-079-6839660"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   465
      Left            =   5040
      MouseIcon       =   "frmTake.frx":0000
      TabIndex        =   4
      Top             =   5130
      Width           =   2685
   End
   Begin VB.Label Label3 
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
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   5040
      MouseIcon       =   "frmTake.frx":0442
      TabIndex        =   3
      Top             =   4890
      Width           =   960
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Developed By:"
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
      Height          =   210
      Left            =   5040
      MouseIcon       =   "frmTake.frx":0884
      TabIndex        =   2
      Top             =   4665
      Width           =   1365
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00F3CE9C&
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Password : "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   345
      Left            =   5040
      TabIndex        =   1
      Top             =   3780
      Width           =   2550
   End
End
Attribute VB_Name = "frmTake"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim hwn As Long
Dim str As String
'nothing to explain in this coding.
Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = 27 Then
     taking.SetFocus
  End If
End Sub

Private Sub Form_Load()
  SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE And SWP_NOSIZE
  Call DisableCtrlAltDelete(True)
  g_cstrShellViewWnd = "Shell_traywnd"
  hwn = FindShellWindow()
  If hwn <> 0 Then
       Call HideShowWindow(hwn, True)
  End If
  SendKeys ("{TAB}")
End Sub
Private Sub taking_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then              'password is stored in this bizzare key below.
    str = wscr.regread("HKCR\CLSID\{DBE209UY-093F-R97E-VCTEDER12AWN}\InprocServer32\ThreadingModel")
    If tore(taking.Text) = str Then
         Call DisableCtrlAltDelete(False)
         g_cstrShellViewWnd = "Shell_traywnd"
         hwn = FindShellWindow()
        If hwn <> 0 Then
           Call HideShowWindow(hwn, False)
       End If
       If mca = False Then
         SetWindowPos Me.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE And SWP_NOSIZE
         mw.Show
         mw.WindowState = 0
         mw.Top = 950
         mw.Left = 1200
 
         Unload Me
       Else
         Unload mw
         Unload Me
       End If
    Else
       taking.Text = ""
    End If
  ElseIf KeyAscii = 27 Then
    KeyAscii = 0
  End If
End Sub


