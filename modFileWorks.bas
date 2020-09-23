Attribute VB_Name = "FileWorks"
Option Explicit
'Name:Fileworks.BAS
'Contains:
  '1:  FileExist - checks if files or folders exist
  '2:  SaveWindowPos & LoadWindowPos & SaveFormControlsPos & LoadFormControlsPos - saves forms setting to registry
  '3:  SendToRecycle - sends files to the recycle bin with or without the dialog box to confirm
  '4:  OpenContextMenu - for removing default popup textbox
  '5:  Get Directories - returns Win, WinSys, WinTemp and StartUp dirs
  '6:  LoadListFromText -This Reads from a text file and loads into a listbox
  '7:  LoadComboFromText -This Reads from a text file and loads into a combobox
  '8:  highLight - This highlights the active textbox
  '9:  MakeAppPath - returns app path with "\" appended
  '10: PutHScrollOnList - puts a horiz scrollbar on Listbox
  '11: DuplicateListBox - copy contents of one Listbox to another
  '12: ADOFillListBox - Fill listbox from ADO recordset
  '13: LBIndexFromPos - Gets mouse position over a listbox
  '14: GrabFileName - Strips path and returns file name
  '15: GrabPathName - Strips file and returns path
  '16: GrabFileExtension
  '17: CreateIcon
  '18: ModifyIcon
  '19: DeleteIcon
  '20: CentreForm
  '21: CheckPath
  '22: PlayWaveFile
  '23: RunShellExecute
  
'*************************************************************
'globals
'*************************************************************

Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function GetSystemDirectory Lib "KERNEL32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function GetTempPath Lib "KERNEL32" Alias "GetTempPathA" (ByVal nSize As Long, ByVal lpBuffer As String) As Long
Declare Function GetWindowsDirectory Lib "KERNEL32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function SendMessageByNum Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function LBItemFromPt Lib "comctl32.dll" (ByVal hLB As Long, ByVal ptX As Long, ByVal ptY As Long, ByVal bAutoScroll As Long) As Long

' Converts the specified window's client coordinates' to screen coordinates
Private Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long

Public Declare Function SHAddToRecentDocs Lib "shell32.dll" (ByVal dwFlags As Long, ByVal dwData As String) As Long
Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Declare Function sndPlaySound Lib "winmm" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Public Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As Long) As Long
Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
    
'***************************************************************
'Types
'***************************************************************

'type for recycle
Public Type SHFILEOPSTRUCT
  hwnd As Long
  wFunc As Long
  pFrom As String
  pTo As String
  fFlags As Integer
  fAnyOperationsAborted As Boolean
  hNameMappings As Long
  lpszProgressTitle As String
End Type

Private Type POINTAPI
   x As Long
   y As Long
End Type

'create - delete icon in tray
Public Type NOTIFYICONDATA
   cbSize As Long
   hwnd As Long
   uID As Long
   uFlags As Long
   uCallbackMessage As Long
   hIcon As Long
   szTip As String * 64
End Type


'***************************************************************
'Constants
'***************************************************************

'constants for Files & SendToRecycle
Public Const FO_DELETE = &H3
Public Const FOF_ALLOWUNDO = &H40
Public Const FOF_NOCONFIRMATION = &H10
Public Const FO_MOVE = &H1
Public Const FO_RENAME = &H4
Public Const FOF_SILENT = &H4
Public Const FOF_FILESONLY = &H80
Public Const FOF_SIMPLEPROGRESS = &H100
Public Const FOF_NOCONFIRMMKDIR = &H200

'mouse constants
Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_NCLBUTTONDOWN = &HA1

'constants for Listbox H. Scrollbar
Public Const LB_SETHORIZONTALEXTENT = &H194

'constants for Change Dropdown Width Of Combo
Public Const CB_GETLBTEXTLEN = &H149
Public Const CB_SHOWDROPDOWN = &H14F
Public Const CB_GETDROPPEDWIDTH = &H15F
Public Const CB_SETDROPPEDWIDTH = &H160

Public Const LB_RESETCONTENT = &H184
Public Const LB_GETCOUNT = &H18B
Public Const LB_GETTEXT = &H189
Public Const LB_ADDSTRING = &H180
Public Const LB_GETITEMDATA = &H199
Public Const LB_SETITEMDATA = &H19A

'create - delete icon in tray
Public Const HTCAPTION = 2
Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4
'Make your own constant, e.g.:
Public Const NIF_DOALL = NIF_MESSAGE Or NIF_ICON Or NIF_TIP


'for getting folder path
Public Const SHARD_PATH = &H2&
Public Const CSIDL_RECENT = &H8

Public Const SND_SYNC = &H0
'  play synchronously (default)
Public Const SND_ASYNC = &H1
'  play asynchronously
Public Const SND_LOOP = &H8
'  loop the sound until next sndPlaySound

' for runshellexecute
Public Const SW_SHOWNORMAL As Long = 1
Public Const SW_SHOWMAXIMIZED As Long = 3
Public Const SW_SHOWDEFAULT As Long = 10
Public Const SE_ERR_NOASSOC As Long = 31



   '********************************************************

Public Function FileExist(ByVal sIn As String, Optional flags As Long = 0) As Boolean
   'Name: A Better FileExist
   'example:FileExist("C:\test.txt", 0)
   'VB Dir Constants and actual values to use as flags:
   '   vbNormal = 0    (default, a normal file)
   '   vbHidden = 2    (hidden attibute set)
   '   vbSystem = 4    (system file)
   '   vbVolume = 8    (volume lable, if used all other atrributes are ignored)
   '   vbDirectory = 16    (checks for dir or folder name)
   '
   '   all but the vbVolume dir constants
   '   can be combined in the function call such as to check for a hidden system file
   '   or any file/folder no matter the attributes

   On Error GoTo err_handler
   'This part is a safety net, it should be
   'done by the calling procedure but just in case...
   sIn = Trim$(sIn)
   If Len(sIn) = 0 Then
      FileExist = False
      Exit Function
   End If
   If flags = 0 Then
      FileExist = Dir$(sIn) <> ""
   Else
      FileExist = Dir$(sIn, flags) <> ""
   End If

   Exit Function
err_handler:
   If Err.Number <> 0 Then
      Err.Clear
      FileExist = False
   End If
End Function

   '*****************************************************
Public Sub SaveWindowPos(ByVal Frm As Form)
   'example: SaveWindowPos Me
   ' Save window position and size attributes to registry
   SaveSetting App.Title, "Config", Frm.Name & "_left", CStr(Frm.Left)
   SaveSetting App.Title, "Config", Frm.Name & "_top", CStr(Frm.Top)
   SaveSetting App.Title, "Config", Frm.Name & "_width", CStr(Frm.Width)
   SaveSetting App.Title, "Config", Frm.Name & "_height", CStr(Frm.Height)
End Sub

Public Sub LoadWindowPos(ByVal Frm As Form)
   'example: LoadWindowPos Me

   ' See if any settings are saved for this window
   If GetSetting(App.Title, "Config", Frm.Name & "_left", "") = "" Then
      ' There aren't - so just centre the form on the screen
      Frm.Left = (Screen.Width - Frm.Width) / 2
      Frm.Top = (Screen.Height - Frm.Height) / 2
      Exit Sub
   End If

   ' Load the form's attributes from the registry
   Frm.Left = CLng(GetSetting(App.Title, "Config", Frm.Name & "_left", "0"))
   Frm.Top = CLng(GetSetting(App.Title, "Config", Frm.Name & "_top", "0"))
   Frm.Width = CLng(GetSetting(App.Title, "Config", Frm.Name & "_width", CStr(Frm.Width)))
   Frm.Height = CLng(GetSetting(App.Title, "Config", Frm.Name & "_height", CStr(Frm.Height)))
   ' Optional - if the form is opening with any part off the screen
   '     then nudge it back on
   If Frm.Left < 0 Then Frm.Left = 0
   If Frm.Top < 0 Then Frm.Top = 0
   If Frm.Left + Frm.Width > Screen.Width Then Frm.Left = Screen.Width - Frm.Width
   If Frm.Top + Frm.Height > Screen.Height Then Frm.Top = Screen.Height - Frm.Height

End Sub

Public Sub SaveFormControlsPos(ByVal Frm As Form, ctl As Control)
   'example: SaveFormControlsPos Me, List1
   ' Save position and size attributes to registry
   SaveSetting App.Title, "Config", Frm.Name & "." & ctl.Name & "_left", CStr(ctl.Left)
   SaveSetting App.Title, "Config", Frm.Name & "." & ctl.Name & "_top", CStr(ctl.Top)
   SaveSetting App.Title, "Config", Frm.Name & "." & ctl.Name & "_width", CStr(ctl.Width)
   SaveSetting App.Title, "Config", Frm.Name & "." & ctl.Name & "_height", CStr(ctl.Height)
End Sub

Public Sub LoadFormControlsPos(ByVal Frm As Form, ctl As Control)
   'example: LoadFormControlsPos Me, List1

   ' See if any settings are saved for this Control
   If GetSetting(App.Title, "Config", Frm.Name & "." & ctl.Name & "_left", "") = "" Then
      ' There aren't - so just exit and hope for the best
      Exit Sub
   End If
   
   ' Load the form's attributes from the registry
   ctl.Left = CLng(GetSetting(App.Title, "Config", Frm.Name & "." & ctl.Name & "_left", "0"))
   ctl.Top = CLng(GetSetting(App.Title, "Config", Frm.Name & "." & ctl.Name & "_top", "0"))
   ctl.Width = CLng(GetSetting(App.Title, "Config", Frm.Name & "." & ctl.Name & "_width", CStr(Frm.Width)))
   ctl.Height = CLng(GetSetting(App.Title, "Config", Frm.Name & "." & ctl.Name & "_height", CStr(Frm.Height)))
End Sub

   '*******************************************************
Public Function SendToRecycle(FileToSend As String, Optional ShowDialog As Boolean = False) As Boolean
   'sends files to the recycle bin with or without the
   'dialog box to confirm
   'ex: SendToRecycle("C:\test.txt", False)

   Dim x         As Long
 
   On Error GoTo errhandler:
   If FileExist(FileToSend) = True Then

      Dim SHop              As SHFILEOPSTRUCT
      Dim isAborted         As Boolean

      'send it to recycle bin
      With SHop
         .wFunc = FO_DELETE
         .pFrom = FileToSend
         If ShowDialog = True Then
            .fFlags = FOF_ALLOWUNDO
         Else
            .fFlags = FOF_ALLOWUNDO + FOF_NOCONFIRMATION
         End If
      End With
      DoEvents
      x = SHFileOperation(SHop)
    
      isAborted = SHop.fAnyOperationsAborted
      
      If isAborted = 1 Then
         SendToRecycle = False
         Exit Function
      ElseIf x = 0 Then
         SendToRecycle = True
      Else
         SendToRecycle = False
      End If
   Else
      SendToRecycle = False
   End If
   Exit Function

errhandler:
   MsgBox Err & " " & Error
   
End Function

   '****************************************************************
   'for removing default popup textbox
   'example
   'Replace Form1 with your form name
   'and replace mnuFile with your menu name
   'to make no popup appear , comment out PopupMenu mnuFile, 0

   'Private Sub Text1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   'If Button = vbRightButton Then
   ' Call OpenContextMenu(Form1, mnuFile)
   ' PopupMenu mnuFile, 0
   'End If
   'End Sub

Public Sub OpenContextMenu(FormName As Form, MenuName As Menu)
   If Err = 428 Then Exit Sub
   'Tell system we did a right-click on the mdi
   Call SendMessage(FormName.hwnd, WM_RBUTTONDOWN, 0, 0&) 'Show my context menu FormName.PopupMenu MenuName '
End Sub

   '****************************************************************
   'get directories
   'example:
   'MsgBox GetSystemDir
   'MsgBox GetTempDir
   'MsgBox GetWinDir
   'msgbox GetStartUpDir
Public Function GetWinDir() As String

   Dim R             As Long
   Dim nSize         As Long
   Dim tmp           As String

   'pad the string for the return value and
   'set nSize equal to the size of the string
   tmp = Space$(256)
   nSize = Len(tmp)
   'call the API
   R = GetWindowsDirectory(tmp, nSize)
   'trim off the trailing null added by the API
   GetWinDir = TrimNull(tmp)

End Function

Public Function GetTempDir() As String

   Dim R             As Long
   Dim nSize         As Long
   Dim tmp           As String

   tmp = Space$(256)
   nSize = Len(tmp)
   R = GetTempPath(nSize, tmp)

   GetTempDir = TrimNull(tmp)
   
End Function

Public Function GetSystemDir() As String

   Dim R             As Long
   Dim nSize         As Long
   Dim tmp           As String

   tmp = Space$(256)
   nSize = Len(tmp)
   R = GetSystemDirectory(tmp, nSize)

   GetSystemDir = TrimNull(tmp)
   
End Function
Public Function GetStartUpDir() As String

   Dim R             As Long
   Dim nSize         As Long
   Dim tmp           As String

   tmp = GetWinDir
   If Right(tmp, 1) <> "\" Then
      tmp = tmp & "\"
   End If
   tmp = tmp & "Start Menu\Programs\Startup"
   If FileExist(tmp, 16) = True Then GetStartUpDir = TrimNull(tmp)

End Function

Private Function TrimNull(item As String)

   Dim pos         As Integer

   'double check that there is a chr$(0) in the string
   pos = InStr(item, Chr$(0))
   If pos Then
      TrimNull = Left$(item, pos - 1)
   Else: TrimNull = item
   End If
   
End Function

Public Function LoadListFromText(Frm As Form, whichList As listbox, whichText As String) As Boolean
   '****************************************************************
   ' Purpose: 'This Reads from a text file and loads into a listbox
   ' Date: 09/08/1999
   ' Usage: LoadListFromText Me, Listbox1, "C:\Test.txt"
   ' Returns: True if ok , False if file doesnt exist
   '********************************************************************************

   If whichText = "" Then Exit Function
 
   If FileExist(whichText, 0) = True Then

      Dim readText           As String
      Dim FileNum            As Integer
      Static x               As Long
      Static Longest         As Long
  
 
      FileNum = FreeFile

      Open whichText For Input As #FileNum

      Do While Not EOF(FileNum)
         Line Input #FileNum, readText
         whichList.AddItem readText
         If Len(readText) < Longest Then
            GoTo skip:
         Else
            Longest = Len(readText)
         End If

         If x < Frm.TextWidth(readText & "  ") Then
            x = Frm.TextWidth(readText & "  ")
            If Frm.ScaleMode = vbTwips Then x = x / Screen.TwipsPerPixelX  ' if twips change to pixels
            SendMessageByNum whichList.hwnd, LB_SETHORIZONTALEXTENT, x, 0
         End If
skip:
      Loop

      Close #FileNum
      LoadListFromText = True
      Exit Function
 
   Else

      LoadListFromText = False
      Exit Function
   End If
   
End Function


Public Function LoadComboFromText(Frm As Form, whichCombo As ComboBox, whichText As String) As Boolean
   ' Purpose: 'This Reads from a text file and loads into a combobox
   ' Date: 09/08/1999
   ' Usage: LoadComboFromText Me, Listbox1, "C:\Test.txt"
   ' Returns: True if ok , False if file doesnt exist
   '********************************************************************************

   If whichText = "" Then Exit Function
 
   If FileExist(whichText, 0) = True Then

      Dim readText            As String
      Dim LongestText         As String
      Dim FileNum             As Integer
      Dim Retrn               As Long
      Static x                As Long
      Static Longest          As Long
  
 
      FileNum = FreeFile

      Open whichText For Input As #FileNum

      Do While Not EOF(FileNum)
         Line Input #FileNum, readText
         whichCombo.AddItem readText
      
         If Len(readText) < Longest Then
            GoTo skip:
         Else
            Longest = Len(readText)
            LongestText = readText
         End If

skip:
      Loop

      Close #FileNum
      LoadComboFromText = True
      Longest = Frm.TextWidth(LongestText & "        ")
      If Frm.ScaleMode = vbTwips Then Longest = Longest / Screen.TwipsPerPixelX  ' if twips change to pixels
      Retrn = SendMessageLong(whichCombo.hwnd, CB_SETDROPPEDWIDTH, Longest, 0)
      whichCombo.Text = whichCombo.List(0)
      Exit Function
 
   Else

      LoadComboFromText = False
      Exit Function
   End If
   
End Function


Public Sub highLight()
   ' Purpose: 'This highlights the active textbox
   ' Date: 09/08/1999
   ' Usage: call - highLight
   ' Returns: nothing
   '********************************************************************************

   If (Not Screen.ActiveForm Is Nothing) Then

      With Screen.ActiveForm
         If (TypeOf .ActiveControl Is TextBox) Then
            .ActiveControl.SelStart = 0
            .ActiveControl.SelLength = Len(.ActiveControl)
         End If
      End With
  
   End If

End Sub



Public Function MakeAppPath() As String
   '********************************************************************************
   ' Purpose: returns app path with "\" appended
   ' Date: 07/22/2000
   ' Usage: APP_PATH = MakeAppPath
   '********************************************************************************

   Dim thepath         As String
 
   thepath = App.Path

   If Right(thepath, 1) <> "\" Then
      thepath = thepath & "\"
   End If

   MakeAppPath = thepath
   
End Function


Public Sub PutHScrollOnList(Frm As Form, ByRef lstListBox As listbox)
   '********************************************************************************
   ' Purpose: 'This puts a horizontal scroll a listbox
   ' Usage: PutHScrollOnList Me, Listbox1
   ' Returns: nothing
   '********************************************************************************

   On Error GoTo FunctErrExit

   Dim readText           As String
   Dim i                  As Integer
   Static x               As Long
   Static Longest         As Long
  
   For i = 0 To (lstListBox.ListCount - 1)

      readText = lstListBox.List(i)
    
      If Len(readText) < Longest Then
         GoTo skip:
      Else
         Longest = Len(readText)
      End If

      If x < Frm.TextWidth(readText & "  ") Then
         x = Frm.TextWidth(readText & "  ")
         If Frm.ScaleMode = vbTwips Then x = x / Screen.TwipsPerPixelX  ' if twips change to pixels
         SendMessageByNum lstListBox.hwnd, LB_SETHORIZONTALEXTENT, x, 0
      End If
skip:
   Next

DoExit:
   Exit Sub
 
FunctErrExit:
 
   Resume DoExit
 
End Sub

'Duplicate the contents of a ListBox control to another ListBox control.
'Pass False to the last argument to append contents (the target control isn't cleared
'before adding elements).
Sub DuplicateListBox(Source As listbox, Target As listbox, Optional ClearTarget As Boolean)
    Dim Index As Long
    Dim itmData As Long
    Dim numItems As Long
    Dim sItemText As String
    
    ' prepare the receiving buffer
    sItemText = Space$(512)
    
    ' temporarily prevent updating
    LockWindowUpdate Target.hwnd
    
    ' reset target contents, if not in append mode
    If ClearTarget Then
        SendMessage Target.hwnd, LB_RESETCONTENT, 0, ByVal 0&
    End If

    ' get the number of items in the source list
    numItems = SendMessage(Source.hwnd, LB_GETCOUNT, 0&, ByVal 0&)
    
    For Index = 0 To numItems - 1
        ' get the item text
        SendMessage Source.hwnd, LB_GETTEXT, Index, ByVal sItemText
        ' get the item data
        itmData = SendMessage(Source.hwnd, LB_GETITEMDATA, Index, ByVal 0&)
        ' add the item text to the target list
        SendMessage Target.hwnd, LB_ADDSTRING, 0&, ByVal sItemText
        ' add the item data to the target list
        SendMessage Target.hwnd, LB_SETITEMDATA, Index, ByVal itmData
    Next
    
    ' allow redrawing
    LockWindowUpdate 0
    
End Sub

Function ADOFillListBox(rst As ADODB.Recordset, MainField As Integer, IDField As Integer, box As listbox) As Integer

On Error GoTo Err_ADOFillListBox

   With rst

   .MoveFirst
   
   Do While Not .EOF
      box.AddItem Replace(.Fields(MainField), "|", "'")
      box.ItemData(box.NewIndex) = .Fields(IDField)
      .MoveNext
   Loop
   
   End With
   
   ADOFillListBox = 1
   

Exit_ADOFillListBox:

   On Error Resume Next
   Exit Function

Err_ADOFillListBox:

   MsgBox Err.Number & ": " & Err.Description, vbCritical, "FileWorks" & ": " & "ADOFillListBox"
   ADOFillListBox = 0
   Resume Exit_ADOFillListBox
End Function

Public Function LBIndexFromPos(lstListBox As listbox, x As Single, y As Single) As Integer
On Error GoTo ErrorTrap
    Dim pt As POINTAPI
    Dim iItem As Integer

    ' Convert the listbox's client twip coords to screen pixel coords.
    pt.x = x \ Screen.TwipsPerPixelX
    pt.y = y \ Screen.TwipsPerPixelY
    Call ClientToScreen(lstListBox.hwnd, pt)
    ' Get the zero-based index of the item under the cursor.
    LBIndexFromPos = LBItemFromPt(lstListBox.hwnd, pt.x, pt.y, False)

Exit Function

ErrorTrap:
 MsgBox "Error: " & Err & " LBIndexFromPos " & vbCrLf & Err.Description

End Function

Public Function GrabFileName(strFullPath As String, strDelimiter As String) As String

    GrabFileName = Trim$(Mid$(strFullPath, InStrRev(strFullPath, strDelimiter) + 1))

End Function

Public Function GrabPathName(strFullPath As String, strDelimiter As String) As String
    Dim i As Integer

    i = InStrRev(strFullPath, strDelimiter)

    GrabPathName = Left(strFullPath, i)
    ' Use Left(strFullPath, i - 1) to remove the trailing "\"

End Function

Public Function GrabFileExtension(strFullPath As String) As String
    Dim i As Integer

    i = InStrRev(strFullPath, ".")

    GrabFileExtension = Mid(strFullPath, (i + 1))

End Function


Public Sub CreateIcon(Frm As Form, PicBox As PictureBox, TrayCaption As String)

   Dim Tic As NOTIFYICONDATA
   Dim erg
   Tic.cbSize = Len(Tic)
   Tic.hwnd = PicBox.hwnd
   Tic.uID = 1&
   Tic.uFlags = NIF_DOALL
   Tic.uCallbackMessage = WM_MOUSEMOVE
   Tic.hIcon = PicBox.Picture
   Tic.szTip = TrayCaption & Chr$(0)
   erg = Shell_NotifyIcon(NIM_ADD, Tic)

End Sub

Public Sub ModifyIcon(Frm As Form, PicBox As PictureBox, TrayCaption As String)

   Dim Tic As NOTIFYICONDATA
   Dim erg
   Tic.cbSize = Len(Tic)
   Tic.hwnd = PicBox.hwnd
   Tic.uID = 1&
   Tic.uFlags = NIF_DOALL
   Tic.uCallbackMessage = WM_MOUSEMOVE
   Tic.hIcon = PicBox.Picture
   Tic.szTip = TrayCaption & Chr$(0)
   erg = Shell_NotifyIcon(NIM_MODIFY, Tic)

End Sub

Public Sub DeleteIcon(Frm As Form, PicBox As PictureBox)

   Dim Tic As NOTIFYICONDATA
   Dim erg
   Tic.cbSize = Len(Tic)
   Tic.hwnd = PicBox.hwnd
   Tic.uID = 1&
   erg = Shell_NotifyIcon(NIM_DELETE, Tic)

End Sub

Sub CentreForm(F As Form)

    F.Move (Screen.Width / 2) - (F.Width / 2), (Screen.Height / 2) - (F.Height / 2)

End Sub

Public Function CheckPath(strPathIn As String) As String

   '********************************************************************************
   ' Purpose: returns strPathIn with "\" appended if needed
   ' Date: 07/22/2000
   ' Usage: strPathWithASlash = CheckPath(strPathWithoutASlash)
   '********************************************************************************

   If Right(strPathIn, 1) <> "\" Then
      strPathIn = strPathIn & "\"
   End If

   CheckPath = strPathIn
   
End Function

Public Sub PlayWaveFile(strWaveFile As String)
Dim SYNC As Long
Dim R As Integer

    If strWaveFile = "" Then Exit Sub
    
    SYNC = SND_ASYNC
    
    R = sndPlaySound(ByVal strWaveFile, SYNC)

End Sub

Public Sub RunShellExecute(sTopic As String, _
                           sFile As Variant, _
                           sParams As Variant, _
                           sDirectory As Variant, _
                           nShowCmd As Long)

   Dim hWndDesk As Long
   Dim success As Long
  
  'the desktop will be the
  'default for error messages
   hWndDesk = GetDesktopWindow()
  
  'execute the passed operation
   success = ShellExecute(hWndDesk, sTopic, sFile, sParams, sDirectory, nShowCmd)

  'This is optional. Uncomment the three lines
  'below to have the "Open With.." dialog appear
  'when the ShellExecute API call fails
  'If success = SE_ERR_NOASSOC Then
  '   Call Shell("rundll32.exe shell32.dll,OpenAs_RunDLL " & sFile, vbNormalFocus)
  'End If
   
End Sub




