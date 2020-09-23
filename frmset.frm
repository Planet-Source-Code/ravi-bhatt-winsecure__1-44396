VERSION 5.00
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVBUTTONS.OCX"
Begin VB.Form frmset 
   BackColor       =   &H00F3CE9C&
   Caption         =   "Settings for WinSecure"
   ClientHeight    =   2715
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmset.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2715
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox ChkSet 
      BackColor       =   &H00F3CE9C&
      Caption         =   "Show time since computer started in tray icon tool tip"
      Height          =   465
      Index           =   2
      Left            =   1125
      TabIndex        =   3
      Tag             =   "TITI"
      Top             =   1260
      Width           =   3525
   End
   Begin VB.CheckBox ChkSet 
      BackColor       =   &H00F3CE9C&
      Caption         =   "Ask for password to open and close"
      Height          =   375
      Index           =   1
      Left            =   1125
      TabIndex        =   2
      Tag             =   "PTO"
      Top             =   675
      Width           =   3525
   End
   Begin VB.CheckBox ChkSet 
      BackColor       =   &H00F3CE9C&
      Caption         =   "Load WinSecure at startup"
      Height          =   330
      Index           =   0
      Left            =   1125
      TabIndex        =   0
      Tag             =   "LST"
      Top             =   135
      Width           =   2895
   End
   Begin LVbuttons.LaVolpeButton ApplyChange 
      Height          =   495
      Left            =   1440
      TabIndex        =   1
      Top             =   2025
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
      MICON           =   "frmset.frx":0442
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
   Begin LVbuttons.LaVolpeButton SetPass 
      Height          =   495
      Left            =   2700
      TabIndex        =   4
      Top             =   2025
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "Change Password"
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
      MICON           =   "frmset.frx":045E
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
   Begin VB.Shape Sha 
      BorderColor     =   &H80000004&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   3435
      Left            =   -45
      Top             =   -45
      Width           =   960
   End
End
Attribute VB_Name = "frmset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer

Private Sub ApplyChange_Click()
  For i = 0 To ChkSet.UBound
    Select Case ChkSet.Item(i).Tag
      Case "LST": 'load at startup
                    If ChkSet.Item(i).Value = 1 Then
                       wscr.regwrite "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\WinSecure", App.path & "\WinSecure.exe", "REG_SZ"
                       wscr.regwrite str3 & "\users\LST", 1, "REG_DWORD"
                    Else
                       On Error Resume Next
                       wscr.regdelete "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\WinSecure"
                       wscr.regwrite str3 & "\users\LST", 0, "REG_DWORD"
                    End If
      Case "TITI": 'tooltip
                    If ChkSet.Item(i).Value = 1 Then
                       mw.Timer1.Enabled = True
                       'store setting in registry for future references
                       'when computer is turned on again or restarted.
                       'TITI = time in tool tip.
                       wscr.regwrite str3 & "\users\TITI", 1, "REG_DWORD"
                       Call mw.Timer1_Timer
                    Else
                       ttip = "WinSecure"
                       mw.Timer1.Enabled = False
                       wscr.regwrite str3 & "\users\TITI", 0, "REG_DWORD"
                       Call newtip 'show new tip.
                    End If
    Case "PTO": 'password to open
                     If ChkSet.Item(i).Value = 1 Then
                       wscr.regwrite str3 & "\users\PTO", 1, "REG_DWORD"
                       bca = True
                     Else
                       wscr.regwrite str3 & "\users\PTO", 0, "REG_DWORD"
                       bca = False
                     End If
    End Select
  Next
End Sub
Sub newtip()
  Shell_NotifyIcon NIM_DELETE, NID
  NID.cbSize = Len(NID)
  NID.uId = vbNull
  NID.szTip = ttip & vbNullChar
  NID.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
  NID.uCallBackMessage = WM_MOUSEMOVE
  Shell_NotifyIcon NIM_ADD, NID
End Sub


Private Sub Form_Load()
   On Error Resume Next
   For i = 0 To ChkSet.UBound 'read values if they exist.
       ChkSet.Item(i).Value = wscr.regread(str3 & "\users\" & ChkSet.Item(i).Tag)
   Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
   mw.Show
End Sub


Private Sub SetPass_Click()
  Me.Hide
  frmloneu.Show
End Sub
