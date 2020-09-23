VERSION 5.00
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVBUTTONS.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMore 
   BackColor       =   &H00F3CE9C&
   Caption         =   "More Settings"
   ClientHeight    =   7470
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6060
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMore.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7470
   ScaleWidth      =   6060
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00F3CE9C&
      Caption         =   "Description"
      Height          =   2040
      Left            =   1755
      TabIndex        =   28
      Top             =   5310
      Width           =   4020
      Begin VB.TextBox txtdes 
         Height          =   1770
         Left            =   45
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   29
         Top             =   180
         Width           =   3930
      End
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   1350
      Top             =   5310
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin LVbuttons.LaVolpeButton butToolTip 
      Height          =   495
      Left            =   450
      TabIndex        =   0
      Top             =   270
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "Tool Tip"
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
      MICON           =   "frmMore.frx":0442
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
   Begin LVbuttons.LaVolpeButton butFldIcon 
      Height          =   495
      Left            =   450
      TabIndex        =   1
      Top             =   945
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "Folder Icon"
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
      MICON           =   "frmMore.frx":045E
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
      Left            =   3105
      TabIndex        =   18
      Top             =   4815
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
      MICON           =   "frmMore.frx":047A
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
   Begin LVbuttons.LaVolpeButton butBSOD 
      Height          =   495
      Left            =   450
      TabIndex        =   31
      Top             =   1665
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "BSOD"
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
      MICON           =   "frmMore.frx":0496
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
      Height          =   4425
      Index           =   0
      Left            =   1755
      TabIndex        =   2
      Top             =   270
      Width           =   4020
   End
   Begin VB.Frame containers 
      BackColor       =   &H00F3CE9C&
      Caption         =   "Tool Tips"
      Height          =   4425
      Index           =   1
      Left            =   1755
      TabIndex        =   3
      Top             =   270
      Width           =   4020
      Begin VB.TextBox txtMc 
         Height          =   285
         Index           =   6
         Left            =   135
         TabIndex        =   16
         Top             =   4005
         Width           =   3795
      End
      Begin VB.TextBox txtMc 
         Height          =   285
         Index           =   5
         Left            =   135
         TabIndex        =   14
         Top             =   3420
         Width           =   3795
      End
      Begin VB.TextBox txtMc 
         Height          =   285
         Index           =   4
         Left            =   135
         TabIndex        =   12
         Top             =   2835
         Width           =   3795
      End
      Begin VB.TextBox txtMc 
         Height          =   285
         Index           =   3
         Left            =   135
         TabIndex        =   10
         Top             =   2250
         Width           =   3795
      End
      Begin VB.TextBox txtMc 
         Height          =   285
         Index           =   2
         Left            =   135
         TabIndex        =   8
         Top             =   1665
         Width           =   3795
      End
      Begin VB.TextBox txtMc 
         Height          =   285
         Index           =   1
         Left            =   135
         TabIndex        =   6
         Top             =   1080
         Width           =   3795
      End
      Begin VB.TextBox txtMc 
         Height          =   285
         Index           =   0
         Left            =   135
         TabIndex        =   4
         Top             =   495
         Width           =   3795
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00F3CE9C&
         Caption         =   "Outlook Express:"
         Height          =   195
         Left            =   135
         TabIndex        =   17
         Top             =   3780
         Width           =   1470
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00F3CE9C&
         Caption         =   "Internet Explorer 5.5:"
         Height          =   195
         Left            =   135
         TabIndex        =   15
         Top             =   3195
         Width           =   1875
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00F3CE9C&
         Caption         =   "Control Panel:"
         Height          =   195
         Left            =   135
         TabIndex        =   13
         Top             =   2610
         Width           =   1230
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00F3CE9C&
         Caption         =   "Recycle Bin:"
         Height          =   195
         Left            =   135
         TabIndex        =   11
         Top             =   2025
         Width           =   1065
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00F3CE9C&
         Caption         =   "My Network:"
         Height          =   195
         Left            =   135
         TabIndex        =   9
         Top             =   1440
         Width           =   1080
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00F3CE9C&
         Caption         =   "My Documents:"
         Height          =   195
         Left            =   135
         TabIndex        =   7
         Top             =   855
         Width           =   1335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00F3CE9C&
         Caption         =   "My Computer:"
         Height          =   195
         Left            =   135
         TabIndex        =   5
         Top             =   270
         Width           =   1230
      End
   End
   Begin VB.Frame containers 
      BackColor       =   &H00F3CE9C&
      Caption         =   "Change Folder Icon"
      Height          =   4425
      Index           =   2
      Left            =   1755
      TabIndex        =   19
      Top             =   270
      Width           =   4020
      Begin VB.CheckBox chkRD 
         BackColor       =   &H00F3CE9C&
         Caption         =   "Resore Default"
         Height          =   330
         Left            =   1080
         TabIndex        =   30
         Tag             =   "RBD"
         Top             =   3915
         Width           =   1680
      End
      Begin VB.TextBox txtTip 
         Height          =   375
         Left            =   90
         TabIndex        =   26
         Top             =   2115
         Width           =   3300
      End
      Begin VB.CommandButton cmdIco 
         Caption         =   "..."
         Height          =   375
         Left            =   3510
         TabIndex        =   25
         Top             =   1395
         Width           =   420
      End
      Begin VB.TextBox txtIco 
         Height          =   375
         Left            =   90
         TabIndex        =   23
         Top             =   1395
         Width           =   3300
      End
      Begin VB.CommandButton cmdFld 
         Caption         =   "..."
         Height          =   375
         Left            =   3510
         TabIndex        =   22
         Top             =   630
         Width           =   420
      End
      Begin VB.TextBox txtFld 
         Height          =   375
         Left            =   90
         TabIndex        =   20
         Top             =   630
         Width           =   3300
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H00F3CE9C&
         Caption         =   "Enter a Tool Tip ..."
         Height          =   195
         Left            =   90
         TabIndex        =   27
         Top             =   1890
         Width           =   1590
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00F3CE9C&
         Caption         =   "Choose a Icon ..."
         Height          =   195
         Left            =   90
         TabIndex        =   24
         Top             =   1170
         Width           =   1485
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00F3CE9C&
         Caption         =   "Choose a Folder ..."
         Height          =   195
         Left            =   90
         TabIndex        =   21
         Top             =   405
         Width           =   1635
      End
   End
   Begin VB.Frame containers 
      BackColor       =   &H00F3CE9C&
      Caption         =   "Blue Screen Of Death"
      Height          =   4425
      Index           =   3
      Left            =   1755
      TabIndex        =   32
      Top             =   270
      Width           =   4020
      Begin VB.ListBox lstTC 
         Height          =   1425
         ItemData        =   "frmMore.frx":04B2
         Left            =   180
         List            =   "frmMore.frx":04E6
         TabIndex        =   36
         Top             =   2745
         Width           =   3615
      End
      Begin VB.ListBox lstBC 
         Height          =   1425
         ItemData        =   "frmMore.frx":058E
         Left            =   180
         List            =   "frmMore.frx":05C2
         TabIndex        =   35
         Top             =   720
         Width           =   3615
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H00F3CE9C&
         Caption         =   "Text Colour:"
         Height          =   195
         Left            =   135
         TabIndex        =   34
         Top             =   2340
         Width           =   1080
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H00F3CE9C&
         Caption         =   "Background Colour:"
         Height          =   195
         Left            =   135
         TabIndex        =   33
         Top             =   360
         Width           =   1725
      End
   End
   Begin VB.Shape Sha 
      BorderColor     =   &H80000004&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   8340
      Left            =   0
      Top             =   -45
      Width           =   960
   End
End
Attribute VB_Name = "frmMore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Vals(6) As String 'array that will hold class ids.
Dim key As String
Dim final As String
Dim i As Integer
Dim vis As Integer 'which frame is visible
Dim colors


Private Sub AC_Click()
   On Error Resume Next 'done for a situation while looping and writing/reading
                        'values in registry and a particular key is not found.
   Select Case vis
     Case 1:
             For i = 0 To 6
                  wscr.regwrite key & Vals(i) & "\" & final, txtMc(i), "REG_SZ"
             Next
     Case 2:
             'for changing folder's icon a file named "desktop.ini"
             'is required in the desired folder.
             'in that file [.ShellClassInfo] is the section under which
             'a key for desired icon "IconFile" is created and "IconIndex"
             'is the index of a icon in case windows supplied icons are used.
             'or icons from exe's, dll's are used.
             'for other files this index is 0.
             'for tips a key named "Infotip" is used.
             If txtFld.Text = "" Then
               MsgBox "Select a Folder.", vbInformation
               Exit Sub
             ElseIf txtIco.Text = "" And chkRD.Value = 0 Then
               MsgBox "Select a Icon.", vbInformation
               Exit Sub
             End If
             If chkRD.Value = 0 Then
               WritePrivateProfileString ".ShellClassInfo", "IconFile", txtIco.Text, txtFld.Text & "\Desktop.ini"
               WritePrivateProfileString ".ShellClassInfo", "IconIndex", 0, txtFld.Text & "\Desktop.ini"
               WritePrivateProfileString ".ShellClassInfo", "InfoTip", txtTip, txtFld.Text & "\Desktop.ini"
               SetAttr txtFld.Text, vbSystem
               SetAttr txtFld.Text & "\Desktop.ini", vbHidden + vbSystem
             ElseIf chkRD.Value = 1 Then
                WritePrivateProfileString ".ShellClassInfo", "IconFile", "", txtFld.Text & "\Desktop.ini"
                WritePrivateProfileString ".ShellClassInfo", "IconIndex", "", txtFld.Text & "\Desktop.ini"
                WritePrivateProfileString ".ShellClassInfo", "InfoTip", "", txtFld.Text & "\Desktop.ini"
             End If
     Case 3:
             'there are only 16 colours permitted for BSOD.
             'the array colors correspondes to colors displayed in list boxes.
             'system.ini file stores numeric(hexadecimal) values for 16 colors.
             colors = Array("0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "A", "B", "C", "D", "E", "F")
             WritePrivateProfileString "386Enh", "MessageBackColor", colors(lstBC.ListIndex), Environ("windir") & "\SYSTEM.INI"
             WritePrivateProfileString "386Enh", "MessageTextColor", colors(lstTC.ListIndex), Environ("windir") & "\SYSTEM.INI"
   End Select
End Sub

Private Sub butBSOD_Click()
  Call hides
  vis = 3
  containers(3).Visible = True
  txtdes.Text = "BSOD is Blue Screen Of Death. This screen appears whenever " _
  & "Windows crash. That window has a Blue background and White foreground. " _
  & "Now you can change that colour too. Choose from the list above and click " _
  & "Apply Changes. The item currently selected(Highlighted) will be set." & vbCrLf & vbCrLf _
  & "Requirements : LOGOFF"
End Sub

Private Sub butFldIcon_Click()
  Call hides
  vis = 2 'second button clicked or second frame visible(what ever u take)
  containers(2).Visible = True
  txtdes.Text = "This setting allows you to get rid of that old yellow folder icon." _
  & " Now you can select your own icon for your folder." & vbCrLf & vbCrLf _
  & "AFTER CHANGING ICON FOR A FOLDER THAT FOLDER CAN NOT BE SEEN FROM PROGRAMS" _
  & " SUCH AS TURBO C++. TO GO BACK TO YOUR OLD SETTING SELECT A FOLDER NAME," _
  & "CLICK ON RESTORE DEFAULT AND FINALLY CLICK ON APPLY CHANGES." _
  & vbCrLf & vbCrLf & "Requirements : REFRESH"

End Sub

Private Sub butToolTip_Click()
  Call hides
  On Error Resume Next
  For i = 0 To 6
    txtMc(i).Text = wscr.regread(key & Vals(i) & "\" & final)
    txtMc(i).ToolTipText = txtMc(i).Text
  Next
  vis = 1
  containers(1).Visible = True
  txtdes.Text = "The Textboxes displays current tool tips of various items." & vbCrLf _
  & "Type in new tips as you want and click on Apply changes. Be careful as there is no way to restore the default tips again." & vbCrLf _
  & vbCrLf & "Requirements : REFRESH"
End Sub

Private Sub cmdFld_Click()
Dim b As BROWSEINFO
Dim p&, rtn&
Dim path As String
Dim pos As Integer

b.hWndOwner = Me.hwnd
b.lpszTitle = "Browse for Folder"
b.ulFlags = BIF_RETURNONLYFSDIRS
p = SHBrowseForFolder(b)

path$ = Space$(512)
rtn = SHGetPathFromIDList(ByVal p, ByVal path)

If rtn Then
      pos = InStr(path$, Chr$(0))
      txtFld.Text = Left(path$, pos - 1) 'got the path of folder selected by user.
Else
      MsgBox "Dialog was cancelled", vbInformation
End If


End Sub

Private Sub cmdIco_Click()
  With cd
    .DialogTitle = "Choose a icon file"
    .DefaultExt = "*.ico"
    .Filter = "Icon Files(*.ico)|*.ico"
    .ShowOpen
    txtIco.Text = .FileName
  End With
End Sub

Private Sub Form_Load()
  vis = 0
  Call assign
  key = "HKLM\Software\Classes\CLSID\"
  final = "InfoTip"
End Sub

Private Sub Form_Unload(Cancel As Integer)
  mw.Show
End Sub
Private Sub hides()
  For i = 0 To containers.Count - 1
     containers.Item(i).Visible = False
   Next
End Sub
Private Sub assign()
  Dim str
  'clsids for
  '0 = my computer
  '1 = my documents
  '2 = my network
  '3 = recycle bin
  '4 = control panel
  '5 = internet explorer
  '6 = outlook express
  str = Array("{20D04FE0-3AEA-1069-A2D8-08002B30309D}", _
  "{450D8FBA-AD25-11D0-98A8-0800361B1103}", _
  "{208D2C60-3AEA-1069-A2D7-08002B30309D}", _
  "{645FF040-5081-101B-9F08-00AA002F954E}", _
  "{21EC2020-3AEA-1069-A2DD-08002B30309D}", _
  "{871C5380-42A0-1069-A2EA-08002B30309D}", _
  "{00020D75-0000-0000-C000-000000000046}")
  For i = 0 To 6
     Vals(i) = str(i)
  Next
End Sub


