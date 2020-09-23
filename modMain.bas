Attribute VB_Name = "modMain"
Option Explicit

Public colEmails As New Collection
Public cStatus As New clsStatus
Public g_bolFirstRun As Boolean

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function SendMessageLong& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function SendMessageByNum& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long

Public Const TH32CS_SNAPPROCESS As Long = 2&
Public Const MAX_PATH As Long = 260

Public Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * MAX_PATH
End Type
    

Public Declare Function CreateToolhelp32Snapshot Lib "KERNEL32" _
   (ByVal lFlags As Long, ByVal lProcessID As Long) As Long

Public Declare Function ProcessFirst Lib "KERNEL32" _
    Alias "Process32First" _
   (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long

Public Declare Function ProcessNext Lib "KERNEL32" _
    Alias "Process32Next" _
   (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long

Public Declare Sub CloseHandle Lib "KERNEL32" _
   (ByVal hPass As Long)

Public Const LB_ADDSTRING& = &H180
Public Const LB_DELETESTRING = &H182
Public Const LB_FINDSTRINGEXACT& = &H1A2
Public Const LB_GETCOUNT& = &H18B
Public Const LB_GETCURSEL& = &H188
Public Const LB_GETITEMDATA = &H199
Public Const LB_GETTEXT = &H189
Public Const LB_GETTEXTLEN& = &H18A
Public Const LB_INSERTSTRING = &H181
Public Const LB_RESETCONTENT& = &H184
Public Const LB_SETHORIZONTALEXTENT = &H194
Public Const LB_SETSEL = &H185

Public g_arrAutoDeleteList() As String
Public g_arrKeyWordList() As String
Public g_arrExclusionList() As String

Public g_AutoDeleteFile As String
Public g_KeyWordFile As String
Public g_ExclusionFile As String
Public g_DeletedEmailFile As String

Public g_intShowExclusionInListview As Integer

Public g_lngListBackColor As Long
Public g_lngUnknownEmailColor As Long
Public g_lngExclusionEmailColor As Long
Public g_lngAutoDeleteEmailColor As Long
Public g_intAutoDeleteOnCheck As Integer
Public g_intMinimizeToTray As Integer
Public g_intSavePassword As Integer
Public g_intPlayFriendWaveForMail As Integer
Public g_strFriendWaveFile As String
Public g_intPlayUnknownWaveForMail As Integer
Public g_strUnknownWaveFile As String
Public g_intCheckMailWhenInTray As Integer
Public g_intCheckMailFrequencyMinutes As Integer
Public g_intOpenEmailFromFriends As Integer
Public g_intOpenEmailAfterProcess As Integer
Public g_strDefaultEmailProgram As String

Public Sub xListKillDupes(listbox As listbox)

    'Kills dublicite items in a listbox

    Dim Search1             As Long
    Dim Search2             As Long
    Dim KillDupe            As Long

    KillDupe = 0

    For Search1& = 0 To listbox.ListCount - 1

        For Search2& = Search1& + 1 To listbox.ListCount - 1
            KillDupe = KillDupe + 1

            If listbox.List(Search1&) = listbox.List(Search2&) Then
                listbox.RemoveItem Search2&
                Search2& = Search2& - 1
            End If

        Next Search2&

    Next Search1&

End Sub

Public Sub GetAutoDeleteList()

    'Loads a file into an array
    On Error Resume Next

    Dim strText            As String
    Dim fFile                  As Integer

    If FileExist(g_AutoDeleteFile) Then
    
        If UBound(g_arrAutoDeleteList) > 0 Then
            Erase g_arrAutoDeleteList
            ReDim Preserve g_arrAutoDeleteList(0)
        End If
    
        fFile = FreeFile
        
        Open g_AutoDeleteFile For Input As fFile
    
        Do While Not EOF(fFile)
            Line Input #fFile, strText
    
            If Len(strText) > 0 Then
                ReDim Preserve g_arrAutoDeleteList(UBound(g_arrAutoDeleteList) + 1)
                g_arrAutoDeleteList(UBound(g_arrAutoDeleteList)) = strText
            End If
    
        Loop
    
        Close fFile
        
    End If

End Sub

Public Sub AddToAutoDeleteList(strAddress As String)

    On Error Resume Next

    Dim Save             As Long
    Dim fFile            As Integer
    Dim i As Integer
    
    For i = 0 To UBound(g_arrAutoDeleteList)
        If strAddress = g_arrAutoDeleteList(i) Then
            Exit Sub
        End If
    Next

    fFile = FreeFile
    Open g_AutoDeleteFile For Append As fFile

    Print #fFile, strAddress

    Close fFile

    ReDim Preserve g_arrAutoDeleteList(UBound(g_arrAutoDeleteList) + 1)
    g_arrAutoDeleteList(UBound(g_arrAutoDeleteList)) = strAddress

End Sub

Public Sub RemoveFromAutoDeleteList(strAddress As String)

    'Loads a file into an array
    On Error Resume Next

    Dim strText            As String
    Dim fFile                  As Integer
    Dim i As Integer
    
    If FileExist(g_AutoDeleteFile) Then
    
        For i = 0 To UBound(g_arrAutoDeleteList)
            If (strAddress = g_arrAutoDeleteList(i)) Then
                g_arrAutoDeleteList(i) = ""
                Exit For
            End If
        Next
        
        fFile = FreeFile
        
        Open g_AutoDeleteFile For Output As fFile
    
        For i = 0 To UBound(g_arrAutoDeleteList)
            If (g_arrAutoDeleteList(i) <> "") Then
                Print #fFile, g_arrAutoDeleteList(i)
            End If
        Next

        Close fFile
        
    End If


End Sub

Public Sub GetExclusionList()

    'Loads a file into an array
    On Error Resume Next

    Dim strText            As String
    Dim fFile                  As Integer

    If FileExist(g_ExclusionFile) Then
    
        If UBound(g_arrExclusionList) > 0 Then
            Erase g_arrExclusionList
            ReDim Preserve g_arrExclusionList(0)
        End If
    
        fFile = FreeFile
        
        Open g_ExclusionFile For Input As fFile
    
        Do While Not EOF(fFile)
            Line Input #fFile, strText
    
            If Len(strText) > 0 Then
                ReDim Preserve g_arrExclusionList(UBound(g_arrExclusionList) + 1)
                g_arrExclusionList(UBound(g_arrExclusionList)) = strText
            End If
    
        Loop
    
        Close fFile
        
    End If

End Sub

Public Sub AddToExclusionList(strAddress As String)

    On Error Resume Next

    Dim Save             As Long
    Dim fFile            As Integer
    Dim i As Integer
    
    For i = 0 To UBound(g_arrExclusionList)
        If strAddress = g_arrExclusionList(i) Then
            Exit Sub
        End If
    Next

    fFile = FreeFile
    Open g_ExclusionFile For Append As fFile

    Print #fFile, strAddress

    Close fFile

    ReDim Preserve g_arrExclusionList(UBound(g_arrExclusionList) + 1)
    g_arrExclusionList(UBound(g_arrExclusionList)) = strAddress

End Sub

Public Sub RemoveFromExclusionList(strAddress As String)

    'Loads a file into an array
    On Error Resume Next

    Dim strText            As String
    Dim fFile                  As Integer
    Dim i As Integer
    
    If FileExist(g_ExclusionFile) Then
    
        For i = 0 To UBound(g_arrExclusionList)
            If (strAddress = g_arrExclusionList(i)) Then
                g_arrExclusionList(i) = ""
                Exit For
            End If
        Next
        
        fFile = FreeFile
        
        Open g_ExclusionFile For Output As fFile
    
        For i = 0 To UBound(g_arrExclusionList)
            If (g_arrExclusionList(i) <> "") Then
                Print #fFile, g_arrExclusionList(i)
            End If
        Next

        Close fFile
        
    End If


End Sub


Public Sub GetKeyWordList()

    'Loads a file into an array
    On Error Resume Next

    Dim strText            As String
    Dim fFile                  As Integer

    If FileExist(g_KeyWordFile) Then
    
        If UBound(g_arrKeyWordList) > 0 Then
            Erase g_arrKeyWordList
            ReDim Preserve g_arrKeyWordList(0)
        End If
    
        fFile = FreeFile
        
        Open g_KeyWordFile For Input As fFile
    
        Do While Not EOF(fFile)
            Line Input #fFile, strText
    
            If Len(strText) > 0 Then
                ReDim Preserve g_arrKeyWordList(UBound(g_arrKeyWordList) + 1)
                g_arrKeyWordList(UBound(g_arrKeyWordList)) = strText
            End If
    
        Loop
    
        Close fFile
        
    End If

End Sub

Public Sub AddToKeyWordList(strKeyWord As String)

    On Error Resume Next

    Dim Save             As Long
    Dim fFile            As Integer
    Dim i As Integer
    
    For i = 0 To UBound(g_arrKeyWordList)
        If strKeyWord = g_arrKeyWordList(i) Then
            Exit Sub
        End If
    Next

    fFile = FreeFile
    Open g_KeyWordFile For Append As fFile

    Print #fFile, strKeyWord

    Close fFile

    ReDim Preserve g_arrKeyWordList(UBound(g_arrKeyWordList) + 1)
    g_arrKeyWordList(UBound(g_arrKeyWordList)) = strKeyWord

End Sub

Public Sub AddToDeletedEmailList(strStringDeleted As String)

    On Error Resume Next

    Dim Save             As Long
    Dim fFile            As Integer
    Dim i As Integer
    
    fFile = FreeFile
    Open g_DeletedEmailFile For Append As fFile

    Print #fFile, strStringDeleted

    Close fFile

End Sub

Public Sub RemoveFromKeyWordList(strKeyWord As String)

    'Loads a file into an array
    On Error Resume Next

    Dim strText            As String
    Dim fFile                  As Integer
    Dim i As Integer
    
    If FileExist(g_KeyWordFile) Then
    
        For i = 0 To UBound(g_arrKeyWordList)
            If (strKeyWord = g_arrKeyWordList(i)) Then
                g_arrKeyWordList(i) = ""
                Exit For
            End If
        Next
        
        fFile = FreeFile
        
        Open g_KeyWordFile For Output As fFile
    
        For i = 0 To UBound(g_arrKeyWordList)
            If (g_arrKeyWordList(i) <> "") Then
                Print #fFile, g_arrKeyWordList(i)
            End If
        Next

        Close fFile
        
    End If


End Sub


Public Sub LoadFileIntoListbox(strFileName As String, lstBox As listbox, Optional strFilterString As String = "")

    'Loads a file into an array
    On Error Resume Next

    Dim strText            As String
    Dim fFile                  As Integer
    Dim bFilter             As Boolean
    
    If Len(strFilterString) > 0 Then
        bFilter = True
    End If

    lstBox.Clear
    
    If FileExist(strFileName) Then
    
        fFile = FreeFile
        
        Open strFileName For Input As fFile
    
        Do While Not EOF(fFile)
            Line Input #fFile, strText
    
            If bFilter Then
                If InStr(1, strText, strFilterString, vbTextCompare) > 0 Then
                    lstBox.AddItem strText
                End If
            Else
                lstBox.AddItem strText
            End If
    
        Loop
    
        Close fFile
        
    End If

End Sub



Public Sub DeleteArrayOfTextLinesFromFile(ByRef arrRemoveArray As Variant, strFileToRemoveFrom As String)


    'Loads a file into an array
    On Error Resume Next

    Dim strText            As String
    Dim fFile                  As Integer
    Dim arrTextKeep()           As String
    Dim i                   As Integer
    

    If FileExist(strFileToRemoveFrom) Then
    
        ReDim Preserve arrTextKeep(0)
        
        fFile = FreeFile
        
        Open strFileToRemoveFrom For Input As fFile
    
        Do While Not EOF(fFile)
            Line Input #fFile, strText
    
            If Len(strText) > 0 Then
                
                For i = 1 To UBound(arrRemoveArray)
                
                    If strText = arrRemoveArray(i) Then
                    
                        GoTo NextLine
                        
                    End If
                    
                Next
                    
                ReDim Preserve arrTextKeep(UBound(arrTextKeep) + 1)
                arrTextKeep(UBound(arrTextKeep)) = strText
                
            End If
    
NextLine:
    
        Loop
    
        Close fFile
        DoEvents
        
        
        If UBound(arrTextKeep) > 0 Then
    
            fFile = FreeFile
            
            Open strFileToRemoveFrom For Output As fFile
        
            For i = 0 To UBound(arrTextKeep)
                If (arrTextKeep(i) <> "") Then
                    Print #fFile, arrTextKeep(i)
                End If
            Next
    
            Close fFile
            
        End If
    
    End If


End Sub


Public Function IsMailProgramRunning(strMailProgram As String) As Boolean
    
    Dim hSnapShot As Long
    Dim uProcess As PROCESSENTRY32
    Dim success As Long
    Dim strExe As String
    Dim i As Integer
    Dim char As String
    
    IsMailProgramRunning = False
    
    hSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0&)

    If hSnapShot = -1 Then Exit Function

    uProcess.dwSize = Len(uProcess)
    success = ProcessFirst(hSnapShot, uProcess)

    If success = 1 Then
    
        Do
        
            strExe = Left(uProcess.szExeFile, InStr(1, uProcess.szExeFile, Chr(0), vbTextCompare) - 1)
            
            If LCase(strExe) = LCase(strMailProgram) Then
                IsMailProgramRunning = True
                Exit Do
            End If
            
        Loop While ProcessNext(hSnapShot, uProcess)
            
    End If

    Call CloseHandle(hSnapShot)


End Function

Public Function GetMailProgramTitle(strMailProgram As String) As String
    Dim strExeName As String
    
    strExeName = GrabFileName(strMailProgram, "\")
    
    If LCase(strExeName) = "outlook.exe" Then
        GetMailProgramTitle = "Inbox - Microsoft Outlook"
    ElseIf LCase(strExeName) = "msimn.exe" Then
        GetMailProgramTitle = "Inbox - Outlook Express"
    Else
        GetMailProgramTitle = "Inbox"
    End If

End Function



