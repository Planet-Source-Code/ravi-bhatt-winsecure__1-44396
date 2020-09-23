Attribute VB_Name = "funAPI"
'***********************************************************************'
'                       Developer Information
'***********************************************************************'
'                   Developed By : Ravi Bhatt
'                        Address : G-307, Srinandnagar-1,Vejalpur,
'                                  Ahmedabad,Gujarat,India.
'                                  pin : 380051
'                         E-mail : rbhatt123@rediffmail.com
'***********************************************************************'

'***********************************************************************'
'                       General Information
'***********************************************************************'
' Initially I was developing this as a software product and was worried
' about Reverse Engineering. That's why some variables have name that
' are out of context. I wanted to make it difficult for that. I have
' changed names of many of the variable but there are quite a few still.
'***********************************************************************'

Option Explicit

''''''''''''''''''''''''''''TYPES'''''''''''''''''''''''''''''''''''''''''
'This type is usd by Shell_NotifyIcon API.
Public Type NOTIFYICONDATA
   cbSize As Long  'size of its variable. use len() function.
   hwnd As Long    'pass the handle of current form.
   uId As Long     'donot know assign vbnull.
   uFlags As Long  'flags for message,tip,icon. constants declared
                   'later in this module.
   uCallBackMessage As Long 'when to display message or tooltip e.g
                            'mouse move.
   hIcon As Long   'icon to displayed in tray. use me.icon.
   szTip As String * 64 'tooltip to be dispalyed.should be null terminated.
End Type

'This type is usd by SHBrowseForFolder API.
Public Type BROWSEINFO
    'Handle of the owner window for the dialog box.
    hWndOwner As Long
    'Pointer to an item identifier list (an ITEMIDLIST structure) specifying the location
    'of the "root" folder to browse from. Only the specified folder and its subfolders
    'appear in the dialog box. This member can be NULL, and in that case, the
    'name space root (the desktop folder) is used.
    'A little info...
    'Objects in the shell’s namespace are assigned item identifiers and item
    'identifier lists. An item identifier uniquely identifies an item within its parent
    'folder. An item identifier list uniquely identifies an item within the shell’s
    'namespace by tracing a path to the item from the desktop.
    pIDLRoot As Long
    'Pointer to a buffer that receives the display name of the folder selected by the
    'user. The size of this buffer is assumed to be MAX_PATH bytes.
    pszDisplayName As String
    'Pointer to a null-terminated string that is displayed above the tree view control
    'in the dialog box. This string can be used to specify instructions to the user.
    lpszTitle As String
    'Value specifying the types of folders to be listed in the dialog box as well as
    'other options. This member can include zero or more of constant values.
    ulFlags As Long
    'Address an application-defined function that the dialog box calls when events
    'occur. For more information, see the description of the BrowseCallbackProc
    'function. This member can be NULL.
    lpfnCallback As Long
    'Application-defined value that the dialog box passes to the callback function
    '(if one is specified).
    lParam As Long
    'Variable that receives the image associated with the selected folder. The image
    'is specified as an index to the system image list.
    iImage As Long
End Type
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''CONSTANTS''''''''''''''''''''''''''''''''''''''''
'The following 2 constants are used by ShowWindow API.
Public Const SW_HIDE = 0 'hide a window.
Public Const SW_SHOW = 5 'show a window.
'This constant is used by SHEmptyRecycleBin API.
Public Const SHERB_NOPROGRESSUI = &H2
'This constant is used by ShellExecute API.
Public Const conSwNormal = 1 'open a window normally.

'The following constants are the messages sent to the
'Shell_NotifyIcon function to add, modify, or delete an icon from the
'taskbar status area.
Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
'The following constants are the flags that indicate the valid
'members of the NOTIFYICONDATA data type.
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4
'The following constant is the message sent when a mouse event occurs
'within the rectangular boundaries of the icon in the taskbar status
'area.
Public Const WM_MOUSEMOVE = &H200

'Left-click constants.(NOT APPLICABLE IN THIS PROJECT.)
'Public Const WM_LBUTTONDBLCLK = &H203   'Double-click
'Public Const WM_LBUTTONDOWN = &H201     'Button down
'Public Const WM_LBUTTONUP = &H202       'Button up

'Right-click constants.
Public Const WM_RBUTTONDOWN = &H204      'Button down
'Public Const WM_RBUTTONDBLCLK = &H206   'Double-click
'Public Const WM_RBUTTONUP = &H205       'Button up

'The folowing constants are used to make a window topmost,not movable etc.
Public Const SWP_NOMOVE = &H2
Public Const HWND_TOPMOST = -1
Public Const SWP_NOSIZE = &H1
Public Const HWND_NOTOPMOST = -2

'This constant is used by SHBrowseForFolder API. There are other constants
'also for this API but they are not declared here. This is used to display
'only folder list without files. Same as when you select something and use
'windows toolbar button to copy.
'Only returns file system directories. If the user selects folders
'that are not part of the file system, the OK button is grayed.
Public Const BIF_RETURNONLYFSDIRS = &H1

'These 3 constants are used by ExitWindowsEx API.
Public Const EWX_LOGOFF As Long = 0
Public Const EWX_SHUTDOWN As Long = 1
Public Const EWX_REBOOT As Long = 2
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''VARIBLES''''''''''''''''''''''''''''''''''''''''''
'This variable is used to access registry. used in all the forms.
Public wscr
'General variable for looping etc.
Dim i As Double
'String for tooltip. used for tray icon tool tip.
Public ttip As String
'These six strings are used for commonly used registry keys.
Public str1 As String
Public str2 As String
Public str3 As String
Public str4 As String
Public str5 As String
Public str6 As String
'Several strings holding registry keys.
Public regKeys(45) As String
'Variable for Password to open setting.
'If TRUE asks for a password to open the s/w.
Public bca As Boolean
'If bca = TRUE then this variable is set to TRUE to make password
'taking form on top of all and disable start button,ALT+CTL+DEL,ALT+TAB
'settings
Public mca As Boolean
'A variable required by Shell_NotifyIcon API. It srores
'Icon to be displayed, tooltip etc.
Public NID As NOTIFYICONDATA
'String that holds the name of the window to be found by FindWindowEx API.
Public g_cstrShellViewWnd As String
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''APIS'''''''''''''''''''''''''''''''''''''''
'API that gives total tickcounts since your computer started/restarted.
'Tickcount is CPU's way to keep track of time.
Public Declare Function GetTickCount Lib "Kernel32.dll" () As Long

'Each Window has a Class. Taskbar,Tray in the taskbar are also windows.
'They belong to a particular class. This API FINDWINDOW takes a window
'class name and the window name as argument and returns a long value.
'That is actully a handle to the window. The second argument is always
'vbNullString. This API can only find parent windows and cannot find
'child windows inside that window.
Declare Function FINDWINDOW Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

'This API finds a child window inside a parent window. Pass to it
'first argument : handle of parent window found by FINDWINDOW
'second argument : 0 as long (????????)
'third argument : class name of child window
'fourth argument : vbnullstring
'returns handle of child window.
Declare Function FindWindowEx Lib "user32" _
  Alias "FindWindowExA" (ByVal hwnd As Long, _
  ByVal hWndChild As Long, ByVal lpszClassName As String, _
  ByVal lpszWindow As String) As Long
  
'This API is used to show/hide a window whose handle is passed as
'first argument to it. Second argument is a constant declared earliar.
'0 for hide and 5 for show.
Declare Function ShowWindow Lib "user32" _
  (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
  
'This API is used to LOGOFF/SHUTDOWN/REBOOT windows.
'First argument is one of the constants declared earliar.
'it is 0 for LOGOFF, 1 for SHUTDOWN, 2 for RESTART. Second argument is
'0 as long.(i don't know what it is.)
Declare Function ExitWindowsEx Lib "user32" (ByVal dwOptions As Long, ByVal dwReserved As Long) As Long

'type help here
Private Declare Function SystemParametersInfo Lib _
"user32" Alias "SystemParametersInfoA" (ByVal uAction _
As Long, ByVal uParam As Long, ByVal lpvParam As Any, _
ByVal fuWinIni As Long) As Long

'This API is used to empty the recycle bin. The first argument is
'handle of current form.(use me.hwnd)
'second argument is "" (again i don't know.)
'third argument is a constant declared earliar.(???HERE)
Public Declare Function SHEmptyRecycleBin Lib "shell32.dll" _
Alias "SHEmptyRecycleBinA" (ByVal hwnd As Long, ByVal _
pszRootPath As String, ByVal dwFlags As Long) As Long

'This API is used to draw a icon in the tray. First argument is a constant
'that tells the API to add,modify or delete icon in tray. Second argument
'is a variable of TYPE NOTIFYICONDATA. This variable contains info like
'tooltip to be displayed,icon to be displayed etc.
Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean


Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

'This API displays a dialog box that enables the user to select a shell folder.
'Returns a pointer to an item identifier list that specifies the location
'of the selected folder relative to the root of the name space. If the user
'chooses the Cancel button in the dialog box, the return value is NULL.
'First argument to this API is a variable of type BROWSEINFO. This variable
'contains details such as from where to start displaying folders etc.
Public Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BROWSEINFO) As Long

'This API converts an item identifier list to a file system path.
'Returns TRUE if successful or FALSE if an error occurs — for example,
'if the location specified by the pidl parameter is not part of the file system.
'First argument is a value returned by SHBrowseForFolder API. Second argument is
'a string that will contain the path.
Public Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long

'This API Updates a complete section contained in a private INI file
'(If the section or any of the values did not previously exist,
'they are created.)
'The first argument is name of section.
'The second argument is the key.
'The third argument is the path of INI file.
'THIS API REWRITES ENTIRE SECTION, SO PREVIOUS VALUES ARE DELETED.
'USE THIS WITH CARE.
Public Declare Function WritePrivateProfileSection Lib "kernel32" _
 Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, _
 ByVal lpString As String, ByVal lpFileName As String) As Long
 
'This API Updates the value of a named key contained in a private INI file
'(If the key did not previously exist, it is created.)
'The first argument is the Section Name inside a INI file.
'The second argument is the Key Name inside a Section.
'The Third argument is a value to be assigned to the Key.
'The Fourth argument is the INI file name.(with path)
Public Declare Function WritePrivateProfileString Lib "kernel32" _
 Alias "WritePrivateProfileStringA" _
 (ByVal lpApplicationName As String, ByVal lpKeyName As String, _
 ByVal lpString As String, ByVal lpFileName_ As String) As Long
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


'''''''''''''''''''''''''''FUNCTIONS''''''''''''''''''''''''''''''''''''''

Public Sub initStr()
' ********************************************************************
' Purpose    : Initialize all the strings used for registry tweaking.
'              and Initialize couple of other variables as well.
' Arguments  : NIL
' Returns    : NIL
' ********************************************************************

   'the following are 12 keys clubbed in Start Menu Button in this s/w.
   'the sequence is same in the list box that contains various settings.
   'with these strings keys are created in registry.
   regKeys(1) = "AlphabeticStartMenu"
   regKeys(2) = "ClearRecentDocsOnExit"
   regKeys(3) = "NoChangeStartMenu"
   regKeys(4) = "NoClose"
   regKeys(5) = "NoFavoritesMenu"
   regKeys(6) = "NoFolderOptions"
   regKeys(7) = "NoLogOff"
   regKeys(8) = "NoRecentDocsMenu"
   regKeys(9) = "NoRun"
   regKeys(10) = "NoFind"
   regKeys(11) = "TaskbarContextMenu"
   regKeys(12) = "NoWindowsUpdate"
   
   'the following are 13 keys clubbed in Control Panel Button in this s/w.
   regKeys(13) = "NoSecCPL"
   regKeys(14) = "NoDispAppearancePage"
   regKeys(15) = "NoDispBackgroundPage"
   regKeys(16) = "NoDevMgrPage"
   regKeys(17) = "NoDispCPL"
   regKeys(18) = "NoFileSysPage"
   regKeys(19) = "NoConfigPage"
   regKeys(20) = "NoPwdPage"
   regKeys(21) = "NoAdminPage"
   regKeys(22) = "NoDispSettingsPage"
   regKeys(23) = "NoDispScrSavPage"
   regKeys(24) = "NoProfilePage"
   regKeys(25) = "NoVirtMemPage"
   
   'the following are 4 keys clubbed in Explorer Button in this s/w.
   regKeys(26) = "Hidden"
   regKeys(27) = "HideFileExt"
   regKeys(28) = "ShowSuperHidden"
   regKeys(29) = "ShowInfoTip"
   regKeys(30) = "SeparateProcess"  'not used.
   
   'the following are 10 keys clubbed in Internet Explorer Button in this s/w.
   regKeys(31) = "NoBrowserClose"
   regKeys(32) = "NoBrowserContextMenu"
   regKeys(33) = "NoBrowserOptions"
   regKeys(34) = "NoBrowserSaveAs"
   regKeys(35) = "NoFavourites"
   regKeys(36) = "NoFileNew"
   regKeys(37) = "NoFileOpen"
   regKeys(38) = "NoFindFiles"
   regKeys(39) = "NoSelectDownloadDir"
   regKeys(40) = "NoTheaterMode"
   
   'the following are 5 keys clubbed in Network Button in this s/w.
   regKeys(41) = "NoNetSetupSecurityPage"
   regKeys(42) = "NoNetSetup"
   regKeys(43) = "NoNetSetupIDPage"
   regKeys(44) = "NoFileSharingControl"
   regKeys(45) = "NoPrintSharing"
   
   'the following are 6 major and most commonly used registry key.
   'inside these major keys above 45 keys will be created.
   str1 = "HKCU\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\"
   str2 = "HKEY_CURRENT_USER\SOFTWARE\MICROSOFT\WINDOWS\CURRENTVERSION\EXPLORER\ADVANCED\"
   str3 = "HKEY_LOCAL_MACHINE\SOFTWARE\MICROSOFT\WINDOWS"
   str4 = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\system\"
   str5 = "HKEY_CURRENT_USER\Software\Policies\Microsoft\Internet Explorer\Restrictions\"
   str6 = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Network\"
   
   '"progman" is the name of the Desktop window. (Explorer window)
   g_cstrShellViewWnd = "progman"
   
   'default tooltip of icon in tray.
   ttip = "WinSecure"
   
   bca = False
   mca = False
End Sub

Sub DisableCtrlAltDelete(bDisabled As Boolean)
    Dim x As Long
    x = SystemParametersInfo(97, bDisabled, CStr(1), 0)
End Sub

Sub Main()
' ************************************************************************
' Purpose    : This function reads values from registry to enable/disable
'              various settings. It also enforce settings made through
'              API calls.
' Arguments  : NIL
' Returns    : NIL
' ************************************************************************

'*************************************************************************
'              -::IMPORTANT READ THIS BEFORE PROCEEDING::-
'I STARTED THIS SOFTWARE WITH ONLY ONE FORM AND A MODULE IN MIND. INITIALLY
'I USED MENUS FOR VARIOUS SETTINGS. THEN I FOUND THIS LAVALOPE BUTTON AND
'DECIDED TO USE IT WITH TREE VIEW. SO I DID NOT WANTED TO CHANGE CODING ALL
'OVER AGAIN. I HAVE USED MENUS' CHECKED PROPERTY TO CHECK/UNCHECK TREEVIEW'S
'ITEMS. THEREFORE IN THIS FUNCTION I FIRST READ VALUES IN MENUS AND THEN
'ASSIGN THOSE VALUES TO TREEVIEW'S ITEMS LATER AT THE FORM LEVEL.
'
'AGAIN MENUS ARE CREATED AS PER THE DECLARATION OF STRINGS INSIDE THE
'INITSTR FUNCTION FOR EASE OF LOOPING. NOW MENUS ARE NOT VISIBLE.
'*************************************************************************
   
   initStr   'Initialize all strings.
   
   'done for a situation while looping and writing/reading
   'values in registry and a particular key is not found.
   On Error Resume Next
  
   Set wscr = CreateObject("Wscript.SHELL") 'create shell's object for accessing registry.
   
   'Almost all keys in registry are DWORD values.
   'value of 1 means setting is enabled and 0 means setting is disabled.
   'read values from reg to menus.
   For i = 2 To 12
     mw.mnuStart.Item(i).Checked = wscr.regread(str1 & regKeys(i))
     If i = 11 Then
       mw.mnuStart.Item(i).Checked = wscr.regread(str2 & regKeys(i))
       mw.mnuStart.Item(i).Checked = Not mw.mnuStart.Item(i).Checked
     End If
   Next
   
   For i = 13 To 25
     mw.mnuControlPanelk.Item(i).Checked = wscr.regread(str4 & regKeys(i))
   Next
   
   For i = 26 To 29
     mw.mnuExpSet(i).Checked = wscr.regread(str2 & regKeys(i))
     If i = 27 Or i = 29 Then
       mw.mnuExpSet(i).Checked = Not mw.mnuExpSet(i).Checked
     ElseIf i = 26 Then
       If wscr.regread(str2 & regKeys(i)) = 1 Then
         mw.mnuExpSet(i).Checked = wscr.regread(str2 & regKeys(i))
       Else
         mw.mnuExpSet(i).Checked = Not mw.mnuExpSet(i).Checked
       End If
     End If
   Next
   
   For i = 31 To 40
     mw.mnuIESet.Item(i).Checked = wscr.regread(str5 & regKeys(i))
   Next
   
   For i = 41 To 45
     mw.mnuNetSet.Item(i).Checked = wscr.regread(str6 & regKeys(i))
   Next
   
   mw.mnuDiskHideAll.Checked = wscr.regread(str1 & "NoDrives")
   mw.mnuDesktopHide.Checked = wscr.regread(str1 & "NoDesktop")
   mw.mnuDesktopDRC.Checked = wscr.regread(str1 & "NoViewContextMenu")
   mw.mnuExplorerDRT.Checked = wscr.regread(str4 & "DisableRegistryTools")
   mw.mnuExplorerNoFile.Checked = wscr.regread(str1 & "NoFileMenu")
   mw.mnuDesktopNNNI.Checked = wscr.regread(str1 & "NoNetHood")
   'these are settings that require APIs. So to enable them at startup
   'again I've created DWORD values(0 or 1) for such setting.reading here.
   mw.mnuDesktopDACD.Checked = wscr.regread(str3 & "\users\" & "ACD") 'ALT+CTL+DEL
   mw.mnuDesktopHide.Checked = wscr.regread(str3 & "\users\" & "HD") 'hide desktop
   mw.mnuDesktopHT.Checked = wscr.regread(str3 & "\users\" & "HT") 'hide taskbar
   mw.mnuHideSB.Checked = wscr.regread(str3 & "\users\" & "HSB") 'hide start button
   mw.mnuDesktopHDT.Checked = wscr.regread(str3 & "\users\" & "HDT") 'hide date time
   

   Call enforceSettings 'enable/disable above read settings.
   
'   default password
'   Dim ss As String
'   ss = tore("ravs") ''function to encrpt password
'   wscr.regwrite "HKCR\CLSID\{DBE209UY-093F-R97E-VCTEDER12AWN}\InprocServer32\ThreadingModel", ss, "REG_SZ"
'   security by obscurity(ha ha.....)
   'check which drive is disabled and check mark menu.
   For i = 0 To mw.mnuDiskRestrictAccessA.Item(mw.mnuDiskRestrictAccessA.UBound).Index
      mw.mnuDiskRestrictAccessA.Item(i).Checked = wscr.regread(str3 & "\users\" & i)
   Next
   
   'reading values for settings form
   For i = 0 To frmset.ChkSet.UBound
       frmset.ChkSet.Item(i).Value = wscr.regread(str3 & "\users\" & frmset.ChkSet.Item(i).Tag)
   Next
  'enable timer that gets Tickcount and convert that to hrs,minutes.
  If frmset.ChkSet.Item(3).Value = 1 Then
    mw.Timer1.Enabled = True
    Call mw.Timer1_Timer
  End If
  
  bca = wscr.regread(str3 & "\users\PTO") 'pto=password to open
  If bca = True Then
    frmTake.Show 'ask for password
  Else
    mw.Show
  End If
End Sub
Public Sub RegRW(s1 As String, s2 As String, Optional b As Double, Optional t As String, Optional c As Menu, Optional chkUNchk As Boolean)
' ************************************************************************
' Purpose    : This function write into registry and checks/unchecks a menu.
' Arguments  : s1 = major registry string
'              s2 = key inside major key.
'               b = value to be written(most of the times 1 or 0).
'               t = type of value to be written in registry.
'                   REG_DWORD , REG_SZ, REG_BINARY etc
'               c = menu that was clicked.
'        chkUNchk = check/uncheck menu. TRUE/FALSE
' Returns    : NIL
' ************************************************************************
   wscr.regwrite s1 & s2, b, t
   c.Checked = chkUNchk
End Sub

Private Sub enforceSettings()
' ************************************************************************
' Purpose    : This subroutine is meant to enforce some settings that are
'              done by a API call. Becasuse once a computer restarts
'              API calls of previous session gets discarded.
' Arguments  : NIL
' Returns    : NIL
' ************************************************************************
   Dim pHwnd As Long 'parent's handle
   Dim cHwnd As Long 'child's handle
   Dim hwnd As Long 'handle
   If mw.mnuDesktopHide.Checked = True Then 'hide desktop
      g_cstrShellViewWnd = "progman"
       hwnd = FindShellWindow() 'finds a handle of a window
       If hwnd <> 0 Then
          Call HideShowWindow(hwnd, True) 'actually hiding
       End If
   End If
   If mw.mnuDesktopHT.Checked = True Then 'hide taskbar
      g_cstrShellViewWnd = "Shell_traywnd" 'name of taskbar class.
      hwnd = FindShellWindow()
      If hwnd <> 0 Then
          Call HideShowWindow(hwnd, True)
      End If
   End If
   If mw.mnuHideSB.Checked = True Then 'hide start button
      pHwnd = FINDWINDOW("Shell_traywnd", vbNullString) 'find parent
      cHwnd = FindWindowEx(pHwnd, 0&, "button", vbNullString) 'find child
      Call HideShowWindow(cHwnd, True) '"button" is start button on taskbar
   End If
   If mw.mnuDesktopHDT.Checked = True Then 'hide date time
      pHwnd = FINDWINDOW("Shell_traywnd", vbNullString)
      cHwnd = FindWindowEx(pHwnd, 0&, "TrayNotifyWnd", vbNullString)
      Call HideShowWindow(cHwnd, True)
   End If
   If mw.mnuDesktopDACD.Checked = True Then 'diable ALT+CTL+DEL
     Call DisableCtrlAltDelete(True)
   End If
End Sub

Public Function FindShellWindow() As Long
' ************************************************************************
' Purpose    : This function finds and returns handle of a window whose
'              class name is contained in globle string g_cstrShellViewWnd.
' Arguments  : NIL
' Returns    : returns handle of a window
' ************************************************************************
Dim hwnd As Long
On Error Resume Next

hwnd = FindWindowEx(0&, 0&, _
  g_cstrShellViewWnd, vbNullString)

If hwnd <> 0 Then
  FindShellWindow = hwnd
End If

End Function

Public Sub HideShowWindow(ByVal hwnd As Long, _
  Optional ByVal Hide As Boolean = False)
' ************************************************************************
' Purpose    : hides/shows window whose handle is passed to it.
' Arguments  : hwnd = handle of window, Hide = TRUE/FALSE
'              default is SHOW
' Returns    : NIL
' ************************************************************************
 Dim lngShowCmd As Long
 On Error Resume Next

 If Hide = True Then
   lngShowCmd = SW_HIDE
 Else
   lngShowCmd = SW_SHOW
 End If

 Call ShowWindow(hwnd, lngShowCmd)

End Sub
Public Function tore(s As String) As String
' ********************************************************************
' Purpose    : This function takes a string that is a password and
'              does some operation on that string to generate a new
'              encrpted password that is stored in registry.
' Arguments  : a string to be encrpted
' Returns    : returns encrpted string
' ********************************************************************
   Dim i As Integer
   Dim j As Integer
   Dim t As Byte
   Dim d As String
   s = StrReverse(s)
   For i = 0 To Len(s) - 1
     t = CByte(Asc(Mid(s, Len(s), 1)))
     t = t + t Mod Len(s)
     d = d & Chr(t)
     s = Mid(s, 1, Len(s) - 1)
   Next
   tore = d
   'does change the string very Slightly.
End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
