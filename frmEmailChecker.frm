VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmEmailChecker 
   Caption         =   "IMAP Email Checker"
   ClientHeight    =   8145
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10350
   Icon            =   "frmEmailChecker.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8145
   ScaleWidth      =   10350
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrMakeIdle 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   8760
      Top             =   5400
   End
   Begin VB.Timer tmrCheckMail 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   9000
      Top             =   5400
   End
   Begin VB.Frame Frame2 
      Height          =   2175
      Left            =   3600
      TabIndex        =   18
      Top             =   5880
      Width           =   3135
      Begin VB.CommandButton cmdProcessMail 
         Caption         =   "Process Mail"
         Enabled         =   0   'False
         Height          =   375
         Left            =   240
         TabIndex        =   19
         Top             =   720
         Width           =   2535
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "Exit"
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   1680
         Width           =   2535
      End
      Begin VB.CommandButton cmdOptions 
         Caption         =   "Options"
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   1200
         Width           =   2535
      End
      Begin VB.CommandButton cmdCheckMail 
         Caption         =   "Check Mail"
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.Frame Frame3 
      Height          =   2175
      Left            =   6960
      TabIndex        =   14
      Top             =   5880
      Width           =   3255
      Begin VB.Label lblMessages 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2205
         TabIndex        =   16
         Top             =   375
         Width           =   375
      End
      Begin VB.Label Label4 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Messages:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   17
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label lblDetails 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Details"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1200
         Left            =   360
         TabIndex        =   15
         Top             =   735
         Width           =   2535
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2175
      Left            =   120
      TabIndex        =   9
      Top             =   5880
      Width           =   3255
      Begin VB.CheckBox chkSavePassword 
         Caption         =   "Save Password"
         Height          =   255
         Left            =   1440
         TabIndex        =   20
         Top             =   1800
         Width           =   1575
      End
      Begin VB.TextBox txtUser 
         Height          =   285
         Left            =   1440
         TabIndex        =   0
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox txtPassword 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1440
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox txtServer 
         Height          =   285
         Left            =   1440
         TabIndex        =   2
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox txtAddress 
         Height          =   285
         Left            =   1440
         TabIndex        =   3
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Username"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Password"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "IMAP Server"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Email Address"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   10
         Top             =   1440
         Width           =   1215
      End
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   9600
      Picture         =   "frmEmailChecker.frx":0442
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   8
      Top             =   5280
      Visible         =   0   'False
      Width           =   480
   End
   Begin MSWinsockLib.Winsock socket 
      Left            =   9360
      Top             =   5280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   5775
      Left            =   120
      TabIndex        =   7
      ToolTipText     =   "Right Click For Options"
      Top             =   120
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   10186
      View            =   3
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "#"
         Object.Width           =   1324
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "From"
         Object.Width           =   6853
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Subject"
         Object.Width           =   6527
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Date"
         Object.Width           =   4146
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Size"
         Object.Width           =   1590
      EndProperty
   End
   Begin VB.Menu zFileHid 
      Caption         =   "File Hid"
      Visible         =   0   'False
      Begin VB.Menu mnuNextCheck 
         Caption         =   "Next Check"
      End
      Begin VB.Menu zSep00 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSHowChecker 
         Caption         =   "Show Checker"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu zLVHid 
      Caption         =   "LV Hid"
      Visible         =   0   'False
      Begin VB.Menu zAddToAutoDelete 
         Caption         =   "Add Checked Items To Auto-Delete"
         Begin VB.Menu mnuAddToAutoDeleteAddress 
            Caption         =   "Full Email Address  -  spam@virtual.domain.com"
         End
         Begin VB.Menu mnuAddToAutoDeleteFull 
            Caption         =   "Full Domain  -  virtual.domain.com"
         End
         Begin VB.Menu mnuAddToAutoDeletePartial 
            Caption         =   "Partial Domain - domain.com"
         End
         Begin VB.Menu mnuAddToAutoDeleteSuffix 
            Caption         =   "Suffix - .com   BE CAREFUL HERE"
         End
      End
      Begin VB.Menu zSep01 
         Caption         =   "-"
      End
      Begin VB.Menu zAddToExclusion 
         Caption         =   "Add Checked Items To Exclusion"
         Begin VB.Menu mnuAddToExclusionAddress 
            Caption         =   "Full Email Address  -  friend@virtual.domain.com"
         End
         Begin VB.Menu mnuAddToExclusionFull 
            Caption         =   "Full Domain  -  virtual.domain.com"
         End
         Begin VB.Menu mnuAddToExclusionPartial 
            Caption         =   "Partial Domain - domain.com"
         End
         Begin VB.Menu mnuAddToExclusionSuffix 
            Caption         =   "Suffix - .com   BE CAREFUL HERE"
         End
      End
      Begin VB.Menu zSep02 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAddToKeyword 
         Caption         =   "Add Selected Subject To Keyword List"
      End
      Begin VB.Menu zSep02A 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewEmail 
         Caption         =   "View Selected Email"
      End
      Begin VB.Menu zSep03 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpenEmailProgram 
         Caption         =   "Open Email Program"
      End
   End
End
Attribute VB_Name = "frmEmailChecker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const m_intMailPort As Integer = 143
Private Const m_lngBottomSpace As Long = 2340

Dim m_strPassword As String
Dim m_strUser As String
Dim m_strServer As String
Dim m_strEmailAddress As String

Dim m_strSPNS As String
Dim m_strRPNS As String
Dim m_intSendPacketNumber As Integer
Dim m_intReturnPacketNumber As Integer
Dim m_intCountOfEmails As Integer
Dim m_intFetchCounter As Integer
Dim m_intSetFlagCounter As Integer

Dim m_intNumberOnAutoDeleteList As Integer
Dim m_intNumberOnWhiteList As Integer
Dim m_intNumberOnKeyWordList As Integer
Dim m_intNumberUnknown As Integer
Dim m_intManualDelete As Integer

Dim m_strSplitted() As String
Dim m_arrEmailNum() As Integer
Dim m_arrPlainDeletions() As Integer

Private Enum emAddress
    FullAddress = 1
    FullDomain = 2
    PartialDomain = 3
    Suffix = 4
End Enum

Private Enum emType
    Friends = 1
    Unknown = 2
    Bad = 3
End Enum


Private Sub chkSavePassword_Click()
    ' Flip the switch
    g_intSavePassword = chkSavePassword.Value
End Sub

Private Sub cmdCheckMail_Click()

    ' Make sure timer is off, close socket, set the state
    tmrMakeIdle.Enabled = False
    socket.Close
    DoEvents
    cStatus.Checking = True
    
    ' Call process
    CheckNewMail
    
End Sub

Public Sub CheckNewMail()
    
    Dim intCheck As Integer
    
    ' Initiate status checks
    cStatus.SetAllToFalse
    
    ' Set to 1, first packet is sent with
    ' socket connecting
    m_intSendPacketNumber = 1
    
    ' Set counters to 0
    m_intReturnPacketNumber = 0
    m_intFetchCounter = 0
    
    m_intCountOfEmails = 0
    m_intNumberOnAutoDeleteList = 0
    m_intNumberOnWhiteList = 0
    m_intNumberOnKeyWordList = 0
    m_intNumberUnknown = 0
    m_intManualDelete = 0
    
    lblDetails = "Details"
    lblMessages = "0"
    
    ' Clear collection of classes
    Set colEmails = Nothing
    
    ListView1.ListItems.Clear
    
    m_strPassword = txtPassword
    m_strUser = txtUser
    m_strServer = txtServer.Text
    
    ' Connect To Mail Server
    If socket.State = sckConnected Then
        socket.Close
        DoEvents
    End If
    
    socket.LocalPort = 0
    socket.Connect m_strServer, m_intMailPort
    
End Sub

Private Sub cmdExit_Click()
    On Error Resume Next
    Unload Me
End Sub

Private Sub cmdOptions_Click()
    Me.Hide
    DoEvents
    frmOptions.Show
End Sub

Private Sub cmdProcessMail_Click()
    
    ' See if any friend or unknown emails are checked
    ' and then handle appropriately
    If Not IsItSafeToDelete Then
        GoTo DoExit
    End If
    
    tmrMakeIdle.Enabled = False
    socket.Close
    DoEvents
    
    ' Set state
    cStatus.Processing = True
    
    ' Call process
    CheckNewMail
    
DoExit:

    Exit Sub
    
End Sub

Private Sub Form_Load()

    Dim fFile            As Integer
    
    ' Build list paths
    g_AutoDeleteFile = CheckPath(App.Path) & "Autodelete.log"
    g_KeyWordFile = CheckPath(App.Path) & "KeyWords.log"
    g_ExclusionFile = CheckPath(App.Path) & "Exclusions.log"
    g_DeletedEmailFile = CheckPath(App.Path) & "DeletedEmail.log"
    
    g_bolFirstRun = CBool(GetSetting(App.Title, "Settings", "FirstRun", "True"))
    
    ' If its the first run, build files
    If g_bolFirstRun Then
        SaveSetting App.Title, "Settings", "FirstRun", "False"
        Call FirstRunSetup
    End If
    
    ' Load lists into arrays
    Call GetAutoDeleteList
    Call GetExclusionList
    Call GetKeyWordList
    Call GetProgramSettings
    
    ' Set state
    cStatus.Idle = True
    cStatus.TimerOn = False
    
    ' Put words in textboxes
    txtUser = m_strUser
    txtPassword = m_strPassword
    txtServer = m_strServer
    txtAddress = m_strEmailAddress
    
    ' Flip switches
    chkSavePassword.Value = g_intSavePassword
    ListView1.BackColor = g_lngListBackColor
    
    ' Place form size & position
    Call LoadWindowPos(Me)
    
    ' If its the first run, show options form
    If g_bolFirstRun Then
        frmOptions.Show vbModal
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call SaveProgramSettings
    
    If Me.WindowState = vbMaximized Then
        Me.WindowState = vbNormal
    End If
    
    Call SaveWindowPos(Me)
    
End Sub

Private Sub ListView1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = vbRightButton Then
        ' Show listview's popup menu
        PopupMenu zLVHid, 0
    End If

End Sub

Private Sub mnuAddToAutoDeleteAddress_Click()
    Call AddCheckedEmailsToAutoDelete(emAddress.FullAddress)
End Sub

Private Sub mnuAddToAutoDeleteFull_Click()
    Call AddCheckedEmailsToAutoDelete(emAddress.FullDomain)
End Sub

Private Sub mnuAddToAutoDeletePartial_Click()
    Call AddCheckedEmailsToAutoDelete(emAddress.PartialDomain)
End Sub

Private Sub mnuAddToAutoDeleteSuffix_Click()
    Call AddCheckedEmailsToAutoDelete(emAddress.Suffix)
End Sub

Private Sub mnuAddToExclusionAddress_Click()
    Call AddCheckedEmailsToExclusion(emAddress.FullAddress)
End Sub

Private Sub mnuAddToExclusionFull_Click()
    Call AddCheckedEmailsToExclusion(emAddress.FullDomain)
End Sub

Private Sub mnuAddToExclusionPartial_Click()
    Call AddCheckedEmailsToExclusion(emAddress.PartialDomain)
End Sub

Private Sub mnuAddToExclusionSuffix_Click()
    Call AddCheckedEmailsToExclusion(emAddress.Suffix)
End Sub

Private Sub mnuAddToKeyword_Click()

    Dim i As Integer
    Dim sTemp As String
    Dim sSubject As String
    Dim x As Integer

    ' Get subject of selected email
    For i = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(i).Selected = True Then
            sSubject = ListView1.ListItems(i).SubItems(2)
            Exit For
        End If
    Next
    
    If sSubject = "" Then
        MsgBox "Error - Subject Not Added To Keyword List", vbOKOnly, "Error"
        Exit Sub
    End If

    ' Would you like wildcards with that?
    sTemp = InputBox("If You Only Want To Enter Part Of The Subject, Add Your Own Or It Will Be Automatically Be Added With Full Wildcards", "Enter Whole Subject Or Edit It", sSubject)
    
    If sTemp = "" Then
        Exit Sub
    End If

    ' If subject was edited without wildcard, then add full wildcards
    If (sTemp <> sSubject) And (InStr(1, sTemp, "*") = 0) Then
        sTemp = "*" & sTemp & "*"
    End If

    ' See if its already in array
    For i = 0 To UBound(g_arrKeyWordList)
        If sTemp = g_arrKeyWordList(i) Then
            MsgBox "KeyWord Already Added To Key Word List", vbOKOnly, "Already In List"
            Exit Sub
        End If
    Next

    ' Not in array, so add it now
    AddToKeyWordList sTemp

    ' Double check
    If g_arrKeyWordList(UBound(g_arrKeyWordList)) <> sTemp Then
        MsgBox "Error: Key Word Not Added To Key Word List" & vbCrLf & vbCrLf & "Please Try Again", vbOKOnly, "Error"
    End If

End Sub

Private Sub mnuExit_Click()

    ' Exit from tray - remove icon
    DeleteIcon Me, Picture1
    
    ' Set window size to normal so it can save settings
    Me.WindowState = vbNormal
    Me.Show
    
    Unload Me
    
End Sub

Private Sub mnuOpenEmailProgram_Click()

    ' Open Email Program - the one entered in the settings form
    Call OpenEmailProgram

End Sub

Private Sub mnuSHowChecker_Click()

    ' Called from tray - remove icon
    DeleteIcon Me, Picture1
    
    ' Show this form
    Me.WindowState = vbNormal
    Me.Show
    
End Sub

Private Sub mnuViewEmail_Click()
Dim i As Integer
Dim x As Integer

    ' Get selected email's number
    For i = 1 To ListView1.ListItems.Count
    
        If ListView1.ListItems(i).Selected = True Then
            x = ListView1.ListItems(i)
            Exit For
        End If
        
    Next
    
    If x <> 0 Then
    
        ' Pass settings to viewer form & open it
        With frmViewMail
        
            .EmailNumber = x
            .MailServer = m_strServer
            .UserName = m_strUser
            .Password = m_strPassword
            
        End With
        
        Load frmViewMail
        frmViewMail.Show vbModal
        
    End If
            

End Sub

Private Sub Socket_DataArrival(ByVal bytesTotal As Long)

    Dim strPacket As String
    
    socket.GetData strPacket
    
    ' Build short string to check return data against
    m_strRPNS = IncrementRPNS

    ' Check state to see where to send data
    If cStatus.CurrentlyFetching Then
        Call ParseEmailHeaders(strPacket)
    ElseIf (cStatus.SetFlagsStart) And (Not cStatus.SetFlagsFinished) Then
        Call ParseEmailFlags(strPacket)
    Else
        Call ProcessNewData(strPacket)
    End If

End Sub


Private Sub ProcessNewData(strData As String)

    On Error GoTo ErrTrap

    Dim strSendString As String
    Dim strTemp As String
    Dim strEmailsNumString As String
    Dim strACheck As String
    
    
    If cStatus.DoSendLogout Then
        GoTo DoSendLogout
    End If
    
    Select Case m_intSendPacketNumber
    
        Case 1:
        
            'Check if we have connected successfully
            If InStr(strData, "* OK") = 0 Then
                GoTo ErrTrap
            Else
                ' Send login to mail server
                strSendString = " LOGIN " & m_strUser & " " & m_strPassword
                m_strEmailAddress = txtAddress
                SendDataToServer strSendString
            End If
                
        Case Else
                
            ' Build short string to check return data against
            strACheck = m_strRPNS & " OK"
            
            ' If connected, but not logged in yet
            If Not cStatus.LoginOK Then
            
                If InStr(strData, strACheck) = 0 Then
                    ' Disconnect & close socket
                    MsgBox "Login Error"
                    cStatus.DoSendLogout = True
                    GoTo DoSendLogout
                Else
                    ' Set state & send data to open inbox
                    cStatus.LoginOK = True
                    strTemp = " SELECT INBOX"
                    SendDataToServer strTemp
                    GoTo DoExit
                End If
                
            End If
            
            ' Logged in & Select Inbox command sent
            If cStatus.LoginOK And Not cStatus.SelectInbox Then
            
                If InStr(strData, strACheck) = 0 Then
                    ' Disconnect & close socket
                    MsgBox "Select Error"
                    cStatus.DoSendLogout = True
                    GoTo DoSendLogout
                Else
                    ' Set state & send data to select all emails
                    cStatus.SelectInbox = True
                    strTemp = " SEARCH ALL"
                    SendDataToServer strTemp
                    GoTo DoExit
                End If
                
            End If
                
            ' Inbox open & command sent to select all emails
            If cStatus.SelectInbox And Not cStatus.SearchAll Then
            
                If InStr(strData, strACheck) = 0 Then
                    ' Disconnect & close socket
                    MsgBox "SearchAll Error"
                    cStatus.DoSendLogout = True
                    GoTo DoSendLogout
                Else
                
                    ' Get useful data
                    strData = Mid(strData, 10)
                    
                    If Mid(strData, 2, 5) = strACheck Then ' no emails
                        ' Disconnect & close socket
                        m_intCountOfEmails = 0
                        cStatus.DoSendLogout = True
                        GoTo DoSendLogout
                    End If
                    
                    ' Get string of email numbers
                    strData = Mid(strData, 1, InStr(strData, m_strRPNS) - 3)
                    
                    ' Split string of email numbers into individual numbers
                    m_strSplitted = Split(strData, " ")
                    
                    ' Set label & count
                    lblMessages.Caption = IIf(strData <> "", (UBound(m_strSplitted) + 1), 0)
                    m_intCountOfEmails = CInt(lblMessages.Caption)
            
                    ' Set state
                    cStatus.SearchAll = True
                    cStatus.CurrentlyFetching = True
                    
                    ' Preset the fetch counter & call process
                    m_intFetchCounter = 1
                    Call FetchSingleEmailData
                    
                    GoTo DoExit
                            
                End If
                
            End If
            
            ' All emails are selected & we are currently fetching individual email data
            If cStatus.SearchAll And cStatus.FetchStarted And Not cStatus.FetchFinished Then
            
                ' All fetching is done, call procedure to check addresses & subjects
                If InStr(strData, strACheck) = 0 Then
                    If m_intCountOfEmails >= m_intFetchCounter Then
                        MsgBox "Fetch Error"
                        Debug.Print strData
                    End If
                    cStatus.CheckingAddresses = True
                    Call CheckAddressesAndSubjects
                Else
                
                    cStatus.CheckingAddresses = True
                    Call CheckAddressesAndSubjects
                    
                End If
                
            End If
            
            ' Fetching is finished & now we are setting flags on emails for deletion
            If cStatus.SearchAll And cStatus.FetchFinished And Not cStatus.SetFlagsFinished Then
            
                ' Set state
                cStatus.SetFlagsStart = True
                
                ' If processing, mark the email's class as so - add number to array for deletion
                ' and add email to deleted emails list
                If ((cStatus.Checking) And (g_intAutoDeleteOnCheck = 1)) Or (cStatus.Processing) Then
                
                    Dim cEHead As clsEmailHeader
                
                    ReDim Preserve m_arrEmailNum(0)
                    
                    For Each cEHead In colEmails
                    
                        If cEHead.Delete = True Then
                        
                            ReDim Preserve m_arrEmailNum(UBound(m_arrEmailNum) + 1)
                            m_arrEmailNum(UBound(m_arrEmailNum)) = cEHead.Email_Number
                            
                            AddToDeletedEmailList (Format(Now, "mmm d, yyyy") & " --- " & cEHead.From & " --- " & cEHead.Subject)
                    
                        End If
                        
                    Next
                    
                    Set cEHead = Nothing
                    
                    ' Call procedure to set flags for emails you want  deleted
                    Call SetSingleEmailFlag
    
                    GoTo DoExit
                    
                Else
                
                    ' Set state
                    cStatus.SetFlagsFinished = True
                    
                End If
                    
            End If
            
            ' Set state
            If cStatus.SetFlagsFinished And Not cStatus.Closing And Not cStatus.DoSendLogout Then
                cStatus.Closing = True
            End If

            ' Deleting is finished - send Close command to expunge deleted emails
            If cStatus.Closing And Not cStatus.DoSendLogout Then
            
                cStatus.DoSendLogout = True
                strTemp = " CLOSE"
                SendDataToServer strTemp
                GoTo DoExit
               
               If InStr(strData, strACheck) = 0 Then
                    MsgBox "Closing Error"
                    Debug.Print strData
                    cStatus.DoSendLogout = True
                    GoTo DoSendLogout
                Else
                    cStatus.DoSendLogout = True
                    GoTo DoSendLogout
                End If

            End If

DoSendLogout:

            If cStatus.LogoutSent Then
            
                ' give program time to decide how to finish, then set status to idle
                tmrMakeIdle.Enabled = True
                
                socket.Close
                
                If InStr(strData, strACheck) = 0 Then
                    MsgBox "Logout Error"
                Else
                    If m_intCountOfEmails = 0 Then
                        ' If not running in tray
                        If cStatus.TimerOn = False Then
                            MsgBox "Currently No E-Mail", vbOKOnly, "IMAP Email Checker"
                        Else
                            Call DecideProcessEnding
                        End If
                    Else
                        Call DecideProcessEnding
                    End If
                End If
                
                GoTo DoExit
                
            End If

            If cStatus.DoSendLogout And Not cStatus.LogoutSent Then
            
                ' Set state & send Logout command
                cStatus.LogoutSent = True
                strSendString = " LOGOUT"
                SendDataToServer strSendString
                
            End If

    End Select

DoExit:

Exit Sub

ErrTrap:
    
    ' What can ya do?
    socket.Close
    
    Debug.Print Error
    lblDetails.Caption = "Error"
    
    Resume DoExit

End Sub

Private Sub ParseEmailHeaders(strDataIn As String)

    'Debug.Print strDataIn
    
    
    If InStr(1, strDataIn, (m_strSPNS & " OK"), vbTextCompare) > 0 Then
        ' this is where we create new class for each email to later
        ' check address aginst exclusion & autodelete, then keyword list
        
        Dim strUID As String
        Dim strFrom As String
        Dim strSubject As String
        Dim strDate As String
        Dim strSize As String
        
        Dim i As Integer
        Dim x As Integer
        Dim cEHead As clsEmailHeader
        Dim strFromAddress As String
        Dim strAfterAlpha As String
        
        ' A new class for each email
        Set cEHead = New clsEmailHeader
        
        ' Get email number
        cEHead.Email_Number = CInt(Mid$(strDataIn, 3, (InStr(3, strDataIn, " ") - 3)))
        
        ' Get UniqueID
        i = (InStr(1, strDataIn, "UID", vbTextCompare) + 4)
        x = (InStr(i, strDataIn, " ", vbTextCompare))
        strUID = Mid(strDataIn, i, (x - i))
        cEHead.EmailUniqueID = strUID
        
        ' Get from name & address
        i = (InStr(1, strDataIn, "FROM:", vbTextCompare) + 6)
        x = (InStr(i, strDataIn, vbCrLf, vbTextCompare))
        strFrom = Mid(strDataIn, i, (x - i))
        strFrom = Replace(strFrom, Chr(34), "")
        cEHead.From = strFrom
        
        ' Getfrom address
        i = (InStr(1, strFrom, "<", vbTextCompare) + 1)
        
        If i > 1 Then
            x = (InStr(i, strFrom, ">", vbTextCompare))
            strFromAddress = Mid(strFrom, i, (x - i))
        Else
            strFromAddress = strFrom
        End If
        
        cEHead.FromAddress = strFromAddress
        
        ' Get domain info from address
        strAfterAlpha = Mid(strFromAddress, (InStr(1, strFromAddress, "@", vbTextCompare) + 1))
        
        i = (InStrRev(strAfterAlpha, "."))
        cEHead.Suffix = Mid(strAfterAlpha, i)
        
        x = (InStrRev(strAfterAlpha, ".", (i - 1)))
        If x <> 0 Then
            cEHead.Domain = Mid(strAfterAlpha, (x + 1))
            cEHead.HasVirtual = True
            cEHead.Virtual = strAfterAlpha
        Else
            cEHead.Virtual = strAfterAlpha
            cEHead.Domain = strAfterAlpha
            cEHead.HasVirtual = False
        End If
        
        ' Get subject
        i = (InStr(1, strDataIn, "SUBJECT:", vbTextCompare) + 9)
        x = (InStr(i, strDataIn, vbCrLf, vbTextCompare))
        strSubject = Mid(strDataIn, i, (x - i))
        cEHead.Subject = strSubject
        
        ' Get date
        i = (InStr(1, strDataIn, "DATE:", vbTextCompare) + 11)
        x = (InStr((i + 15), strDataIn, " ", vbTextCompare))
        
        If x = 0 Then
            x = (InStr(i, strDataIn, " +", vbTextCompare))
        End If
        
        If x = 0 Then
            x = (InStr(i, strDataIn, vbCrLf, vbTextCompare))
        End If
        
        strDate = Mid(strDataIn, i, (x - i))
        cEHead.SentDate = Format(strDate, "mmm d, yyyy h:nn:ss AMPM")
        
        ' Get size
        i = (InStr(1, strDataIn, "RFC822.SIZE", vbTextCompare) + 12)
        x = (InStr(i, strDataIn, " ", vbTextCompare))
        strSize = Mid(strDataIn, i, (x - i))
        cEHead.Size = strSize
        
        ' Add class to collection
        colEmails.Add cEHead, CStr(cEHead.Email_Number)
        
        Set cEHead = Nothing
        
    End If
        
    ' Increment counter
    m_intFetchCounter = m_intFetchCounter + 1
    
    
    If m_intCountOfEmails < m_intFetchCounter Then
        cStatus.CurrentlyFetching = False
    Else
        lblMessages = m_intFetchCounter
        DoEvents
    End If
        
    ' Fetch next
    FetchSingleEmailData

End Sub

Private Function IncrementSPNS() As String
    ' String to send at start of each packet sent
     m_intSendPacketNumber = m_intSendPacketNumber + 1
    IncrementSPNS = "A" & m_intSendPacketNumber
End Function

Private Function IncrementRPNS() As String
    ' String for checking returned packets against sent packets
    m_intReturnPacketNumber = m_intReturnPacketNumber + 1
    IncrementRPNS = "A" & m_intReturnPacketNumber
End Function


Private Sub FetchSingleEmailData()
    Dim strTemp As String
    
    ' Send command to fetch individual email data
    strTemp = " Fetch " & m_intFetchCounter & " (flags uid RFC822.SIZE rfc822.header.lines (from subject date))"
    SendDataToServer strTemp

End Sub

Private Sub CheckAddressesAndSubjects()
Dim i As Integer
Dim cEHead As clsEmailHeader
Dim strTemp As String
Dim intType As Integer

    ' Loop through collection, make checks on each email
    For Each cEHead In colEmails

        ' if processing emails check if its a plain delete
        If cStatus.Processing Then
        
            For i = 1 To UBound(m_arrPlainDeletions)
                If m_arrPlainDeletions(i) = cEHead.Email_Number Then
                    cEHead.Delete = True
                    m_intManualDelete = m_intManualDelete + 1
                    GoTo NextHead
                End If
            Next
            
        End If
            
        'check for matches in exclusion list
        If cEHead.FromAddress = m_strEmailAddress Then
            cEHead.OK = True
            m_intNumberOnWhiteList = m_intNumberOnWhiteList + 1
            
            If g_intShowExclusionInListview = 1 Then
                intType = emType.Friends
                GoTo AddToList
            Else
                GoTo NextHead
            End If
            
        End If
        
        If IsInExclusionList(cEHead.FromAddress) Then
            cEHead.OK = True
            m_intNumberOnWhiteList = m_intNumberOnWhiteList + 1
            
            If g_intShowExclusionInListview = 1 Then
                intType = emType.Friends
                GoTo AddToList
            Else
                GoTo NextHead
            End If
            
        End If
        
        If cEHead.Virtual <> "" Then
            If IsInExclusionList(cEHead.Virtual) Then
                cEHead.OK = True
                m_intNumberOnWhiteList = m_intNumberOnWhiteList + 1
            
            If g_intShowExclusionInListview = 1 Then
                intType = emType.Friends
                GoTo AddToList
            Else
                GoTo NextHead
            End If
            
            End If
        End If
        
        If IsInExclusionList(cEHead.Domain) Then
            cEHead.OK = True
            m_intNumberOnWhiteList = m_intNumberOnWhiteList + 1
            
            If g_intShowExclusionInListview = 1 Then
                intType = emType.Friends
                GoTo AddToList
            Else
                GoTo NextHead
            End If
            
        End If
        
        'check for matches in autodelete list
        If IsInAutoDeleteList(cEHead.FromAddress) Then
            cEHead.Delete = True
            m_intNumberOnAutoDeleteList = m_intNumberOnAutoDeleteList + 1
            
            If ((cStatus.Checking) And (g_intAutoDeleteOnCheck = 0)) Then
                intType = emType.Bad
                GoTo AddToList
            Else
                GoTo NextHead
            End If
            
        End If
        
        If cEHead.Virtual <> "" Then
            If IsInAutoDeleteList(cEHead.Virtual) Then
                cEHead.Delete = True
                m_intNumberOnAutoDeleteList = m_intNumberOnAutoDeleteList + 1
            
                If ((cStatus.Checking) And (g_intAutoDeleteOnCheck = 0)) Then
                    intType = emType.Bad
                    GoTo AddToList
                Else
                    GoTo NextHead
                End If
            
            End If
        End If
        
        If IsInAutoDeleteList(cEHead.Domain) Then
            cEHead.Delete = True
            m_intNumberOnAutoDeleteList = m_intNumberOnAutoDeleteList + 1
            
            If ((cStatus.Checking) And (g_intAutoDeleteOnCheck = 0)) Then
                intType = emType.Bad
                GoTo AddToList
            Else
                GoTo NextHead
            End If
            
        End If
        
        If cEHead.Suffix <> "" Then
            If IsInAutoDeleteList(cEHead.Suffix) Then
                cEHead.Delete = True
                m_intNumberOnAutoDeleteList = m_intNumberOnAutoDeleteList + 1
            
                If ((cStatus.Checking) And (g_intAutoDeleteOnCheck = 0)) Then
                    intType = emType.Bad
                    GoTo AddToList
                Else
                    GoTo NextHead
                End If
            
            End If
        End If
        
        'check for matches in keyword list
        If cEHead.Subject <> "" Then
            If IsInKeyWordList(cEHead.Subject) Then
                cEHead.Delete = True
                m_intNumberOnKeyWordList = m_intNumberOnKeyWordList + 1
            
                If ((cStatus.Checking) And (g_intAutoDeleteOnCheck = 0)) Then
                    intType = emType.Bad
                    GoTo AddToList
                Else
                    GoTo NextHead
                End If
            
            End If
        End If
        
        ' Not in any list
        m_intNumberUnknown = m_intNumberUnknown + 1
        intType = emType.Unknown
        
AddToList:

        ' Just checking
        If ListView1.BackColor <> g_lngListBackColor Then
            ListView1.BackColor = g_lngListBackColor
        End If
        
        ' Add email to list & set forecolor
        Dim lvItem As ListItem
        Set lvItem = ListView1.ListItems.Add(, , CStr(cEHead.Email_Number))
        lvItem.SubItems(1) = cEHead.From
        lvItem.SubItems(2) = cEHead.Subject
        lvItem.SubItems(3) = cEHead.SentDate
        lvItem.SubItems(4) = cEHead.Size
        
        If intType = emType.Friends Then
            lvItem.ForeColor = g_lngExclusionEmailColor
        ElseIf intType = emType.Bad Then
            lvItem.ForeColor = g_lngAutoDeleteEmailColor
        Else
            lvItem.ForeColor = g_lngUnknownEmailColor
        End If
        
        lvItem.ListSubItems.item(1).ForeColor = lvItem.ForeColor
        lvItem.ListSubItems.item(2).ForeColor = lvItem.ForeColor
        lvItem.ListSubItems.item(3).ForeColor = lvItem.ForeColor
        lvItem.ListSubItems.item(4).ForeColor = lvItem.ForeColor
        
        lvItem.EnsureVisible
        Set lvItem = Nothing
        
NextHead:
        
    Next
        
    ' Make string Details label
    strTemp = "White List: " & Space((9 - Len(CStr(m_intNumberOnWhiteList)))) & m_intNumberOnWhiteList & vbCrLf
    strTemp = strTemp & "Auto-Delete: " & Space((8 - Len(CStr(m_intNumberOnAutoDeleteList)))) & m_intNumberOnAutoDeleteList & vbCrLf
    strTemp = strTemp & "Key Word List: " & Space((6 - Len(CStr(m_intNumberOnKeyWordList)))) & m_intNumberOnKeyWordList & vbCrLf
    strTemp = strTemp & "Unknown: " & Space((12 - Len(CStr(m_intNumberUnknown)))) & m_intNumberUnknown
    If m_intManualDelete > 0 Then
        strTemp = strTemp & vbCrLf & "Manual-Delete: " & Space((6 - Len(CStr(m_intManualDelete)))) & m_intManualDelete
    End If
    
    lblDetails = strTemp
    
    If cStatus.TimerOn Then
        ' Make string for icon tooltip
        strTemp = "White List....." & m_intNumberOnWhiteList & vbCrLf
        strTemp = strTemp & "Auto-Delete.." & m_intNumberOnAutoDeleteList & vbCrLf
        strTemp = strTemp & "Key Word....." & m_intNumberOnKeyWordList & vbCrLf
        strTemp = strTemp & "Unknown......" & m_intNumberUnknown
        ModifyIcon Me, Picture1, strTemp
    End If
    
End Sub

Private Sub AddCheckedEmailsToExclusion(intAddressType As Integer)
Dim i As Integer
Dim x As Integer
Dim cEHead As clsEmailHeader
Dim strAddress As String
Dim intSelected As Integer


    ' See which list item is selected, so we can reselect it later
    For i = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(i).Selected = True Then
            intSelected = i
            Exit For
        End If
    Next
    
    
    ' Loop through & add checked items to exclusion list
    For i = 1 To ListView1.ListItems.Count
    
        If ListView1.ListItems(i).Checked = True Then

            ListView1.ListItems(i).Checked = False
            
            ListView1.ListItems(i).Selected = True
            
            For Each cEHead In colEmails
        
                If cEHead.Email_Number = CInt(ListView1.SelectedItem) Then
                
                    ' Get the full or partial address
                    Select Case intAddressType
                    
                        Case emAddress.FullAddress ' full addy
                            strAddress = cEHead.FromAddress
                            
                        Case emAddress.FullDomain ' virtual
                            If cEHead.Virtual <> "" Then
                                strAddress = cEHead.Virtual
                            Else
                                strAddress = cEHead.Domain
                            End If
                            
                        Case emAddress.PartialDomain ' domain
                            strAddress = cEHead.Domain
                            
                        Case emAddress.Suffix  ' suffix
                            strAddress = cEHead.Suffix
                            
                        Case Else
                            strAddress = cEHead.FromAddress
                            
                    End Select
                
                    ' Add to list
                    Call AddToExclusionList(strAddress)
                    
                    ' If any other emails with same address (per type) then
                    ' make their forecolor the exclusion color
                    Call ColorListEmailsWithSameAddress(strAddress, intAddressType, g_lngExclusionEmailColor)
                
                    Exit For
                    
                End If
        
NextHead:
        
            Next
            
        End If
    
    Next
    
    Set cEHead = Nothing
    
    On Error Resume Next
    
    ' Reselect originally selected email
    ListView1.ListItems(intSelected).Selected = True

    ' Refresh the array
    Call GetExclusionList

End Sub

Private Function IsInExclusionList(strAddress As String) As Boolean

    Dim i            As Integer
    Dim b            As Boolean

    b = False
    
    If UBound(g_arrExclusionList) > 0 Then
    
        ' Loop through exclusion array & see if there is a match
        For i = 0 To UBound(g_arrExclusionList)
        
            If UCase(strAddress) = UCase(g_arrExclusionList(i)) Then
            
                b = True
                Exit For
            
            End If
        
        Next
    
    End If

    IsInExclusionList = b
    
    Exit Function
    
End Function

Private Function IsInAutoDeleteList(strAddress As String) As Boolean

    Dim i            As Integer
    Dim b            As Boolean

    b = False

    If UBound(g_arrAutoDeleteList) > 0 Then
    
        ' Loop through autodelete array & see if there is a match
        For i = 0 To UBound(g_arrAutoDeleteList)
        
            If (UCase(strAddress) = UCase(g_arrAutoDeleteList(i))) Then
                b = True
                Exit For
            End If
                
        Next
    
    End If

    IsInAutoDeleteList = b
    
    Exit Function
    
End Function

Private Function IsInKeyWordList(strSubject As String) As Boolean

    Dim i            As Integer
    Dim b            As Boolean

    b = False

    If UBound(g_arrKeyWordList) > 0 Then
    
        ' Loop through keyword array & see if there is a match
        For i = 0 To UBound(g_arrKeyWordList)
        
            If (g_arrKeyWordList(i) <> "") Then
            
                If (Left(g_arrKeyWordList(i), 1) = "*") Or (Right(g_arrKeyWordList(i), 1) = "*") Then
                
                    ' Keyword has wildcard, see if is match for partial subject
                    If IsPartialInKeyWordList(g_arrKeyWordList(i), strSubject) Then
                        b = True
                        Exit For
                    End If
                    
                Else
                
                    ' Keyword has no wildcards, see if is match for full subject
                    If (UCase(strSubject) = UCase(g_arrKeyWordList(i))) Then
                        b = True
                        Exit For
                    End If
                
                End If
            
            End If
        
        Next
    
    End If

    IsInKeyWordList = b
    
    Exit Function

End Function

Private Function IsPartialInKeyWordList(strPartial As String, strSubject As String) As Boolean

    On Error GoTo ErrorTrap

    Dim b As Boolean
    Dim i As Integer
    Dim s As String
    
    b = False
    
    ' See if keyword is in subject text
    
    If Left(strPartial, 1) = "*" And Right(strPartial, 1) = "*" Then
        If InStr(1, strSubject, Mid(strPartial, 2, (Len(strPartial) - 2)), vbTextCompare) > 0 Then
            b = True
            GoTo DoExit
        End If
        GoTo DoExit
    End If
    
    If Left(strPartial, 1) = "*" Then
        If UCase(Right(strSubject, (Len(strPartial) - 1))) = UCase(Mid(strPartial, 2)) Then
            b = True
            GoTo DoExit
        End If
    End If
    
    If Right(strPartial, 1) = "*" Then
        If UCase(Left(strSubject, (Len(strPartial) - 1))) = UCase(Left(strPartial, (Len(strPartial) - 1))) Then
            b = True
            GoTo DoExit
        End If
    End If

DoExit:
    IsPartialInKeyWordList = b
    Exit Function
    
ErrorTrap:

    b = False
    Resume DoExit
    
End Function

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    ' Callback for icon when running in tray
    
    x = x / Screen.TwipsPerPixelX
    
    Select Case x
        Case WM_LBUTTONDOWN
            'Options
        Case WM_RBUTTONDOWN
            PopupMenu zFileHid
        Case WM_MOUSEMOVE
            'Options
        Case WM_LBUTTONDBLCLK
            DeleteIcon Me, Picture1
            Me.WindowState = vbNormal
            Me.Show
            tmrCheckMail.Enabled = False
            cStatus.TimerOn = False
    End Select

End Sub


Private Sub Form_Resize()

    If Me.WindowState = vbMinimized Then
    
        Me.Hide
        DoEvents
        
        ' Put icon in tray
        CreateIcon Me, Picture1, "IMAP Email Checker"
        
        ' Turn timer on & set status
        tmrCheckMail.Enabled = True
        cStatus.TimerOn = True
        cStatus.Checking = True
        
    Else
    
        ' Turn timer off, set status & resize controls
        tmrCheckMail.Enabled = False
        cStatus.TimerOn = False
        Call SetFormControlSizes
        
    End If
    
End Sub

Private Sub ParseEmailFlags(strDataIn As String)

    If InStr(1, strDataIn, (m_strSPNS & " OK"), vbTextCompare) > 0 Then
        ' will eventually need some kind of error handling here i guess
    End If
        
    ' Increment counter
    m_intSetFlagCounter = m_intSetFlagCounter + 1
    
    If m_intSetFlagCounter >= UBound(m_arrEmailNum) Then
        ' Last email
        cStatus.SetFlagsFinished = True
    End If
    
    ' Call procedure to send Store command
    SetSingleEmailFlag

End Sub

Private Sub SetSingleEmailFlag()
    Dim strTemp As String
    Dim i As Integer
    
    On Error Resume Next
    i = (UBound(m_arrEmailNum) - m_intSetFlagCounter)
    
    If i < 1 Then
        i = 0
    End If
    
    ' Mark email on server to be deleted
    strTemp = " STORE " & m_arrEmailNum(i) & " +FLAGS (\Deleted)"
    SendDataToServer strTemp
    
    'STORE 2:4 +FLAGS (\Deleted)

End Sub

Private Sub SendDataToServer(strDataOut As String)
    
    Dim strTemp As String
    
    ' Increment sent packet number string
    m_strSPNS = IncrementSPNS
    
    ' Send data to email server
    strTemp = m_strSPNS & strDataOut & vbCrLf
    socket.SendData strTemp

End Sub

Private Sub AddCheckedEmailsToAutoDelete(intAddressType As Integer)
Dim i As Integer
Dim x As Integer
Dim cEHead As clsEmailHeader
Dim strAddress As String
Dim intSelected As Integer


    ' See which list item is selected, so we can reselect it later
    For i = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(i).Selected = True Then
            intSelected = i
            Exit For
        End If
    Next
    
    
    ' Loop through & add checked items to autodelete list
    For i = 1 To ListView1.ListItems.Count
    
        If ListView1.ListItems(i).Checked = True Then

            ListView1.ListItems(i).Checked = False
            
            ListView1.ListItems(i).Selected = True
            
            For Each cEHead In colEmails
        
                If cEHead.Email_Number = CInt(ListView1.SelectedItem) Then
                
                    ' Get the full or partial address
                    Select Case intAddressType
                    
                        Case emAddress.FullAddress ' full addy
                            strAddress = cEHead.FromAddress
                            
                        Case emAddress.FullDomain ' virtual
                            If cEHead.Virtual <> "" Then
                                strAddress = cEHead.Virtual
                            Else
                                strAddress = cEHead.Domain
                            End If
                            
                        Case emAddress.PartialDomain ' domain
                            strAddress = cEHead.Domain
                            
                        Case emAddress.Suffix ' suffix
                            strAddress = cEHead.Suffix
                            
                        Case Else
                            strAddress = cEHead.FromAddress
                            
                    End Select
                
                    ' Add to list
                    Call AddToAutoDeleteList(strAddress)
                    
                    ' If any other emails with same address (per type) then
                    ' make their forecolor the autodelete color
                    Call ColorListEmailsWithSameAddress(strAddress, intAddressType, g_lngAutoDeleteEmailColor)
                    
                    Exit For
                    
                End If
        
NextHead:
        
            Next
            
        End If
    
    Next
    
    Set cEHead = Nothing
    
    On Error Resume Next
    
    ' Reselect originally selected email
    ListView1.ListItems(intSelected).Selected = True

    ' Refresh the array
    Call GetAutoDeleteList

End Sub

Private Sub GetProgramSettings()

    ' Use built in VB function to get settings from registry
    
    g_intShowExclusionInListview = CInt(GetSetting(App.Title, "Settings", "ShowExclusionInListview", "1"))
    g_lngListBackColor = CLng(GetSetting(App.Title, "Settings", "ListBackColor", "-2147483643"))
    g_lngUnknownEmailColor = CLng(GetSetting(App.Title, "Settings", "UnknownEmailColor", "0"))
    g_lngExclusionEmailColor = CLng(GetSetting(App.Title, "Settings", "ExclusionEmailColor", "-2147483636"))
    g_lngAutoDeleteEmailColor = CLng(GetSetting(App.Title, "Settings", "AutoDeleteEmailColor", "192"))
    g_intCheckMailWhenInTray = CInt(GetSetting(App.Title, "Settings", "CheckMailWhenInTray", "0"))
    g_intCheckMailFrequencyMinutes = CInt(GetSetting(App.Title, "Settings", "CheckMailFrequencyMinutes", "10"))
    g_intSavePassword = CInt(GetSetting(App.Title, "Settings", "SavePassword", "0"))
    g_intPlayFriendWaveForMail = CInt(GetSetting(App.Title, "Settings", "PlayFriendWaveForMail", "0"))
    g_strFriendWaveFile = GetSetting(App.Title, "Settings", "FriendWaveFile", "")
    g_intPlayUnknownWaveForMail = CInt(GetSetting(App.Title, "Settings", "PlayUnknownWaveForMail", "0"))
    g_strUnknownWaveFile = GetSetting(App.Title, "Settings", "UnknownWaveFile", "")
    g_intAutoDeleteOnCheck = CInt(GetSetting(App.Title, "Settings", "AutoDeleteOnCheck", "0"))
    g_intOpenEmailFromFriends = CInt(GetSetting(App.Title, "Settings", "OpenEmailFromFriends", "0"))
    g_intOpenEmailAfterProcess = CInt(GetSetting(App.Title, "Settings", "OpenEmailAfterProcess", "0"))
    g_strDefaultEmailProgram = GetSetting(App.Title, "Settings", "DefaultEmailProgram", "C:\Program Files\Outlook Express\msimn.exe")

    m_strUser = GetSetting(App.Title, "Settings", "StandardUser", "")
    m_strServer = GetSetting(App.Title, "Settings", "StandardServer", "")
    m_strEmailAddress = GetSetting(App.Title, "Settings", "EmailAddress", "")
    m_strPassword = GetSetting(App.Title, "Settings", "Password", "")

End Sub

Private Sub SaveProgramSettings()

    ' Use built in VB function to save settings to registry
    
    SaveSetting App.Title, "Settings", "ShowExclusionInListview", g_intShowExclusionInListview
    SaveSetting App.Title, "Settings", "ListBackColor", g_lngListBackColor
    SaveSetting App.Title, "Settings", "UnknownEmailColor", g_lngUnknownEmailColor
    SaveSetting App.Title, "Settings", "ExclusionEmailColor", g_lngExclusionEmailColor
    SaveSetting App.Title, "Settings", "AutoDeleteEmailColor", g_lngAutoDeleteEmailColor
    SaveSetting App.Title, "Settings", "CheckMailWhenInTray", g_intCheckMailWhenInTray
    SaveSetting App.Title, "Settings", "SavePassword", g_intSavePassword
    SaveSetting App.Title, "Settings", "PlayFriendWaveForMail", g_intPlayFriendWaveForMail
    SaveSetting App.Title, "Settings", "FriendWaveFile", g_strFriendWaveFile
    SaveSetting App.Title, "Settings", "PlayUnknownWaveForMail", g_intPlayUnknownWaveForMail
    SaveSetting App.Title, "Settings", "UnknownWaveFile", g_strUnknownWaveFile
    SaveSetting App.Title, "Settings", "AutoDeleteOnCheck", g_intAutoDeleteOnCheck
    SaveSetting App.Title, "Settings", "CheckMailFrequencyMinutes", g_intCheckMailFrequencyMinutes
    SaveSetting App.Title, "Settings", "OpenEmailFromFriends", g_intOpenEmailFromFriends
    SaveSetting App.Title, "Settings", "OpenEmailAfterProcess", g_intOpenEmailAfterProcess
    SaveSetting App.Title, "Settings", "DefaultEmailProgram", g_strDefaultEmailProgram
    
    SaveSetting App.Title, "Settings", "StandardUser", m_strUser
    SaveSetting App.Title, "Settings", "StandardServer", m_strServer
    SaveSetting App.Title, "Settings", "EmailAddress", m_strEmailAddress
    
    If g_intSavePassword = 1 Then
        SaveSetting App.Title, "Settings", "Password", m_strPassword
    End If

End Sub

Private Sub DecideProcessEnding()
    Dim intCount As Integer
    
    ' Get total
    intCount = m_intNumberOnWhiteList + m_intNumberUnknown + m_intNumberOnAutoDeleteList
    
    If cStatus.Processing Then
    
        ' open email program after processing emails?
        If (g_intOpenEmailAfterProcess = 1) Then
        
            ' Are there any emails in the list?
            If ListView1.ListItems.Count > 0 Then
                Call OpenEmailProgram
                DoEvents
            End If
        
        End If
        
        cmdProcessMail.Enabled = False
        
    End If
    
    ' If only exclusion emails
    If (m_intNumberOnWhiteList > 0 And m_intNumberUnknown = 0 And m_intNumberOnAutoDeleteList = 0) Then
        
        ' If running in tray
        If cStatus.TimerOn Then
        
            ' New count larger than last count
            If m_intNumberOnWhiteList > cStatus.PreviousCountFriend Then
                cStatus.PreviousCountFriend = m_intNumberOnWhiteList
                
                If (g_strFriendWaveFile <> "") And FileExist(g_strFriendWaveFile) Then
                    Call PlayWaveFile(g_strFriendWaveFile)
                    DoEvents
                End If
                
            End If
            
        Else
        
            If (g_intPlayFriendWaveForMail = 1) And (Not cStatus.Processing) Then
            
                ' New count larger than last count
                If m_intNumberOnWhiteList > cStatus.PreviousCountFriend Then
                    cStatus.PreviousCountFriend = m_intNumberOnWhiteList
                End If
                
                If (g_strFriendWaveFile <> "") And FileExist(g_strFriendWaveFile) Then
                    Call PlayWaveFile(g_strFriendWaveFile)
                End If
                
            End If
            
        End If
        
        cmdProcessMail.Enabled = True
        GoTo CheckTotals
        
    End If
        
    
    ' If exclusion or unknown emails
    If (m_intNumberOnWhiteList > 0 Or m_intNumberUnknown > 0) And (g_intPlayUnknownWaveForMail = 1) And (Not cStatus.Processing) Then
        
        ' If running in tray
        If cStatus.TimerOn Then
        
            ' New count larger than last count
            If intCount > cStatus.PreviousCountAll Then
                cStatus.PreviousCountAll = intCount
                
                ' New count larger than last count
                If m_intNumberOnWhiteList > cStatus.PreviousCountFriend Then
                    cStatus.PreviousCountFriend = m_intNumberOnWhiteList
                End If
                
                If (g_strUnknownWaveFile <> "") And FileExist(g_strUnknownWaveFile) Then
                    Call PlayWaveFile(g_strUnknownWaveFile)
                    DoEvents
                End If
                
            End If
            
        Else
        
            If intCount > cStatus.PreviousCountAll Then
                cStatus.PreviousCountAll = intCount
                
                ' New count larger than last count
                If m_intNumberOnWhiteList > cStatus.PreviousCountFriend Then
                    cStatus.PreviousCountFriend = m_intNumberOnWhiteList
                End If
                
            End If
            
            If (g_strUnknownWaveFile <> "") And FileExist(g_strUnknownWaveFile) Then
                Call PlayWaveFile(g_strUnknownWaveFile)
            End If
            
        End If
        
        cmdProcessMail.Enabled = True
        
    End If
    
CheckTotals:
    
    If cStatus.Processing Then
        intCount = m_intNumberOnWhiteList + m_intNumberUnknown
    End If
    
   ' New count smaller than last count
    If intCount < cStatus.PreviousCountAll Then
        cStatus.PreviousCountAll = intCount
    End If
    
   ' New count smaller than last count
    If m_intNumberOnWhiteList < cStatus.PreviousCountFriend Then
        cStatus.PreviousCountFriend = m_intNumberOnWhiteList
    End If
    
DoExitProcedure:

End Sub

Private Sub OpenEmailProgram()
   Dim sTopic As String
   Dim sFile As String
   Dim sParams As Variant
   Dim sDirectory As Variant
   Dim lngHandle As Long

    ' If email program selected in Options form
    If (g_strDefaultEmailProgram <> "") Then
    
        ' Get it's title - currently just set up for MS Outlook & Outlook Express
        ' Add more email programs to GetMailProgramTitle function if you want more
        If cStatus.MailProgramTitle = "" Then
            cStatus.MailProgramTitle = GetMailProgramTitle(g_strDefaultEmailProgram)
        End If
        
        ' See if program is already running
        If Not IsMailProgramRunning(GrabFileName(g_strDefaultEmailProgram, "\")) Then
        
            If FileExist(g_strDefaultEmailProgram) Then
            
               sTopic = "Open"
               sFile = g_strDefaultEmailProgram
               sParams = ""
               sDirectory = 0&
        
                Call RunShellExecute(sTopic, sFile, sParams, sDirectory, SW_SHOWNORMAL)
            
                DoEvents
                
            Else
            
                MsgBox "Email Program Filepath Is Invalid!", vbOKOnly, "Open Email Program"
                
            End If
            
        Else ' Already running - make it active & bring to front
        
            ' set handle to zero - because it may be changed
            ' enum windows to get handle for mail program
            cStatus.MailProgramHandle = 0
            Call EnumWindows(AddressOf EnumWindowProc, &H0)
            
            lngHandle = cStatus.MailProgramHandle
            
            If lngHandle <> 0 Then
                'activate mail program
                RestoreWindow lngHandle
                DoEvents
            End If
        
        End If
        
    End If

End Sub

Private Sub ColorListEmailsWithSameAddress(strAddy As String, intType As Integer, lngColor As Long)

Dim i As Integer
Dim x As Integer
Dim intCount As Integer
Dim strAddress As String
Dim strFromAddress As String
Dim strAfterAlpha As String
Dim strFrom As String
Dim strDomain As String
Dim strVirtual As String
Dim strSuffix As String
Dim bolMarkIt As Boolean

    ' This takes all emails in the list that have the exact same address qualities
    ' as the address passed to the function - and colors them accordingly

    ' Loop through list
    For intCount = 1 To ListView1.ListItems.Count
    
    bolMarkIt = False
    
    With ListView1.ListItems(intCount)
    
        .Selected = True
        
        ' Get email address in sections
        strFrom = .ListSubItems.item(1).Text
        
        i = (InStr(1, strFrom, "<", vbTextCompare) + 1)
        
        If i > 1 Then
            x = (InStr(i, strFrom, ">", vbTextCompare))
            strFromAddress = Mid(strFrom, i, (x - i))
        Else
            strFromAddress = strFrom
        End If
        
        strAfterAlpha = Mid(strFromAddress, (InStr(1, strFromAddress, "@", vbTextCompare) + 1))
        
        i = (InStrRev(strAfterAlpha, "."))
        strSuffix = Mid(strAfterAlpha, i)
        
        x = (InStrRev(strAfterAlpha, ".", (i - 1)))
        If x <> 0 Then
            strDomain = Mid(strAfterAlpha, (x + 1))
            strVirtual = strAfterAlpha
        Else
            strDomain = strAfterAlpha
            strVirtual = strAfterAlpha
        End If
        
        ' Match it to the type, and address passed to the procedure
        Select Case intType
        
            Case emAddress.FullAddress
                If strAddy = strFromAddress Then
                    bolMarkIt = True
                End If
                
            Case emAddress.FullDomain
                If strAddy = strVirtual Then
                    bolMarkIt = True
                End If
                
            Case emAddress.PartialDomain
                If strAddy = strDomain Then
                    bolMarkIt = True
                End If
                
            Case emAddress.Suffix
                If strAddy = strSuffix Then
                    bolMarkIt = True
                End If
                
            Case Else
                If strAddy = strFromAddress Then
                    bolMarkIt = True
                End If
                
        End Select
        
        ' If full or partial address is a match, color it accordingly
        If bolMarkIt Then
            .ForeColor = lngColor
            .ListSubItems.item(1).ForeColor = lngColor
            .ListSubItems.item(2).ForeColor = lngColor
            .ListSubItems.item(3).ForeColor = lngColor
            .ListSubItems.item(4).ForeColor = lngColor
            DoEvents
        End If
    
    End With
    
    Next

End Sub

Private Sub tmrCheckMail_Timer()
Static intCount As Integer

    intCount = intCount + 1
    
    ' Timer runs only if prog is running in tray
    If g_intCheckMailWhenInTray = 1 Then
        If intCount >= g_intCheckMailFrequencyMinutes Then
            intCount = 0
            CheckNewMail
            mnuNextCheck.Caption = "Next Check: " & g_intCheckMailFrequencyMinutes & " Minutes"
        End If
        mnuNextCheck.Caption = "Next Check: " & (g_intCheckMailFrequencyMinutes - intCount) & " Minutes"
    End If

End Sub

Private Sub SetFormControlSizes()

    On Error Resume Next
    
    Dim lngFrameTop1 As Long
    Dim lngCenter As Long
    
    With ListView1
        .Move 120, 120, (Me.ScaleWidth - 240), (Me.ScaleHeight - m_lngBottomSpace)
        .ColumnHeaders(1).Width = 750
        .ColumnHeaders(2).Width = (.Width / 2.599)
        .ColumnHeaders(3).Width = (.Width / 2.73)
        .ColumnHeaders(4).Width = 2350
        .ColumnHeaders(5).Width = 901
    End With
    
    lngFrameTop1 = ListView1.Height + 120
    lngCenter = Me.ScaleWidth / 2
    
    Frame1.Move 120, lngFrameTop1
    Frame2.Move (lngCenter - (Frame2.Width / 2)), lngFrameTop1
    Frame3.Move (Me.ScaleWidth - (Frame3.Width + 120)), lngFrameTop1

End Sub

Private Sub tmrMakeIdle_Timer()

    tmrMakeIdle.Enabled = False
    cStatus.Idle = True

End Sub

Private Sub zFileHid_Click()
    
    If mnuNextCheck.Caption = "Next Check" Then
        mnuNextCheck.Caption = "Next Check: " & g_intCheckMailFrequencyMinutes & " Minutes"
    End If
    
End Sub

Private Sub zLVHid_Click()

    ' See if an email program is selected
    If g_strDefaultEmailProgram = "" Then
        mnuOpenEmailProgram.Enabled = False
    Else
        mnuOpenEmailProgram.Enabled = True
    End If
    
    'Enable or disable popup functionality
    If ListView1.ListItems.Count = 0 Then
        zAddToAutoDelete.Enabled = False
        zAddToExclusion.Enabled = False
        mnuAddToKeyword.Enabled = False
        mnuViewEmail.Enabled = False
    Else
    
        If IsEmailsChecked Then
            zAddToAutoDelete.Enabled = True
            zAddToExclusion.Enabled = True
        Else
            zAddToAutoDelete.Enabled = False
            zAddToExclusion.Enabled = False
        End If
        
        If IsEmailsSelected Then
            mnuAddToKeyword.Enabled = True
            mnuViewEmail.Enabled = True
        Else
            mnuAddToKeyword.Enabled = False
            mnuViewEmail.Enabled = False
        End If
        
    End If
        
End Sub

Private Function IsItSafeToDelete() As Boolean
Dim i As Integer
Dim intEx As Integer
Dim intUnk As Integer
Dim intAnswer As Integer
Dim strMsg As String

    Erase m_arrPlainDeletions
    ReDim Preserve m_arrPlainDeletions(0)
    
    ' Loop through list & see if any friend or unknown emails are checked
    With ListView1
    
        For i = 1 To .ListItems.Count
            If .ListItems(i).Checked = True Then
                If .ListItems(i).ForeColor <> g_lngAutoDeleteEmailColor Then
                    If .ListItems(i).ForeColor = g_lngExclusionEmailColor Then
                        intEx = intEx + 1
                    ElseIf .ListItems(i).ForeColor = g_lngUnknownEmailColor Then
                        intUnk = intUnk + 1
                    End If
                    
                    ReDim Preserve m_arrPlainDeletions(UBound(m_arrPlainDeletions) + 1)
                    m_arrPlainDeletions(UBound(m_arrPlainDeletions)) = CInt(.ListItems(i).Text)
                    
                End If
            End If
        Next
    
    End With

    ' Build message string
    If (intEx > 0) Then
        strMsg = intEx & " Emails On Your Exclusion List"
    End If
        
    If (intUnk > 0) Then
        If strMsg = "" Then
            strMsg = intUnk & " Unknown Emails"
        Else
            strMsg = strMsg & " And " & intUnk & " Unknown Emails "
        End If
    End If
    
    ' If none, then go ahead & process mail
    If strMsg = "" Then
        IsItSafeToDelete = True
    Else
    
        'Else ask if its ok to delete them
        intAnswer = MsgBox("There Are " & strMsg & "That Are Checked For Deletion" & vbCrLf & vbCrLf & "Do You Want To Delete These Emails?", vbYesNoCancel, "Delete Emails?")
        
        If intAnswer = vbYes Then
            IsItSafeToDelete = True
        ElseIf intAnswer = vbNo Then
            ' No returns a True - because it clears
            ' the arrPlainDeletions array first
            IsItSafeToDelete = True
            Erase m_arrPlainDeletions
        ElseIf intAnswer = vbCancel Then
            ' Cancel returns a False - and it
            ' causes the Process Mail function to halt
            IsItSafeToDelete = False
            Erase m_arrPlainDeletions
        Else
            IsItSafeToDelete = False
            Erase m_arrPlainDeletions
        End If
        
    End If

End Function


Private Function IsEmailsChecked() As Boolean
Dim i As Integer

    IsEmailsChecked = False
    
    ' See if anything in the list is checked
    With ListView1
    
        For i = 1 To .ListItems.Count
            If .ListItems(i).Checked = True Then
                IsEmailsChecked = True
                Exit For
            End If
        Next
        
    End With
        
End Function

Private Function IsEmailsSelected() As Boolean
Dim i As Integer

    IsEmailsSelected = False
    
    ' See if anything in the list is selected
    With ListView1
    
        For i = 1 To .ListItems.Count
            If .ListItems(i).Selected = True Then
                IsEmailsSelected = True
                Exit For
            End If
        Next
        
    End With
        
End Function


Private Sub FirstRunSetup()

    On Error Resume Next
    Dim fFile As Integer
    
    ' create log files & add first line of data
    
    fFile = FreeFile
    Open g_KeyWordFile For Append As fFile
    Print #fFile, "*prescription*"
    Close fFile

    DoEvents
    
    fFile = FreeFile
    Open g_AutoDeleteFile For Append As fFile
    Print #fFile, "allspam.com"
    Close fFile

    DoEvents
    
    fFile = FreeFile
    Open g_ExclusionFile For Append As fFile
    Print #fFile, "hotmail.com"
    Close fFile

    DoEvents
    
End Sub
