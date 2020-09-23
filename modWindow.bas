Attribute VB_Name = "modWindow"
Option Explicit

'--------------------------------------------------------------
' Copyright ©1996-2003 VBnet, Randy Birch, All Rights Reserved.
' Terms of use http://www.mvps.org/vbnet/terms/pages/terms.htm
'--------------------------------------------------------------

Private Type RECT
    Left    As Long
    Top     As Long
    Right   As Long
    Bottom  As Long
End Type

Private Type POINTAPI
    x       As Long
    y       As Long
End Type

Private Type WINDOWPLACEMENT
    Length            As Long
    flags             As Long
    showCmd           As Long
    ptMinPosition     As POINTAPI
    ptMaxPosition     As POINTAPI
    rcNormalPosition  As RECT
End Type

Private Const SW_SHOWNORMAL = 1
Private Const SW_SHOWMINIMIZED = 2
Private Const SW_SHOWMAXIMIZED = 3
Private Const SW_SHOWNOACTIVATE = 4
Private Const WPF_RESTORETOMAXIMIZED = &H2

Public Declare Function EnumWindows Lib "user32" _
  (ByVal lpEnumFunc As Long, _
   ByVal lParam As Long) As Long

Private Declare Function GetWindowText Lib "user32" _
    Alias "GetWindowTextA" _
   (ByVal hwnd As Long, _
    ByVal lpString As String, _
    ByVal cch As Long) As Long
    
Private Declare Function GetClassName Lib "user32" _
    Alias "GetClassNameA" _
   (ByVal hwnd As Long, _
    ByVal lpClassName As String, _
    ByVal nMaxCount As Long) As Long

Private Declare Function GetWindowTextLength Lib "user32" _
    Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
 
Private Declare Function IsWindowVisible Lib "user32" _
   (ByVal hwnd As Long) As Long
   
Private Declare Function GetParent Lib "user32" _
   (ByVal hwnd As Long) As Long

Private Declare Function IsWindowEnabled Lib "user32" _
   (ByVal hwnd As Long) As Long

Private Declare Function IsZoomed Lib "user32" _
   (ByVal hwnd As Long) As Long

Private Declare Function GetWindowPlacement Lib "user32" _
  (ByVal hwnd As Long, _
   lpwndpl As WINDOWPLACEMENT) As Long
   
Private Declare Function SetWindowPlacement Lib "user32" _
  (ByVal hwnd As Long, _
   lpwndpl As WINDOWPLACEMENT) As Long
   
Private Declare Function BringWindowToTop Lib "user32" _
   (ByVal hwnd As Long) As Long

Private Declare Function SetForegroundWindow Lib "user32" _
   (ByVal hwnd As Long) As Long

Public Function EnumWindowProc(ByVal hwnd As Long, ByVal lParam As Long) As Long
   
  'working vars
   Dim nSize As Long
   Dim sTitle As String
   Dim pos As Integer
   Dim iX As Integer
   
   iX = 1
   
  'eliminate windows that are not top-level.
   If GetParent(hwnd) = 0& And _
      IsWindowVisible(hwnd) And _
      IsWindowEnabled(hwnd) Then

     'get the size of the string required
     'to hold the window title
      nSize = GetWindowTextLength(hwnd)
         
     'if the return is 0, there is no title
      If nSize > 0 Then
         
         sTitle = Space$(nSize + 1)
         Call GetWindowText(hwnd, sTitle, nSize + 1)
         sTitle = TrimNull(sTitle)
            
      Else
         
        'no title, so get the class name instead
         sTitle = Space$(64)
         Call GetClassName(hwnd, sTitle, 64)
         sTitle = TrimNull(sTitle) & "  (class)"
         
      End If
      
        If InStr(1, sTitle, cStatus.MailProgramTitle, vbTextCompare) Then
            cStatus.MailProgramHandle = hwnd
            iX = 0
        End If
    
   End If
   
                       
  'To continue enumeration, return True
  'To stop enumeration return False (0).
  'When 1 is returned, enumeration continues
  'until there are no more windows left.
   EnumWindowProc = iX
   
End Function


Private Function TrimNull(startstr As String) As String

  Dim pos As Integer

  pos = InStr(startstr, Chr$(0))
  
  If pos Then
      TrimNull = Left$(startstr, pos - 1)
      Exit Function
  End If
  
  TrimNull = startstr
  
End Function

Public Sub RestoreWindow(hWndToRestore As Long)

   Dim currWinP As WINDOWPLACEMENT
   
  'if a window handle passed
   If hWndToRestore Then
   
     'prepare the WINDOWPLACEMENT type
     'to receive the window coordinates
     'of the specified handle
      currWinP.Length = Len(currWinP)
   
     'get the info...
      If GetWindowPlacement(hWndToRestore, currWinP) > 0 Then
      
        'based on the returned info,
        'determine the window state
         If currWinP.showCmd = SW_SHOWMINIMIZED Then
      
           'it is minimized, so restore it
            With currWinP
               .Length = Len(currWinP)
               .flags = 0&
               .showCmd = SW_SHOWMAXIMIZED
            End With
            
            Call SetWindowPlacement(hWndToRestore, currWinP)
         
         Else
           
           'it is on-screen, so make it visible
            Call SetForegroundWindow(hWndToRestore)
            Call BringWindowToTop(hWndToRestore)
         
         End If
      End If
   End If
   
End Sub


