VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "IMAP Email Checker - Lists & Settings Form"
   ClientHeight    =   7665
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8340
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7665
   ScaleWidth      =   8340
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CD1 
      Left            =   240
      Top             =   7680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   13150
      _Version        =   393216
      Tabs            =   5
      Tab             =   4
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "EXCLUSION"
      TabPicture(0)   =   "frmOptions.frx":0442
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "cmdExclusionRefresh"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblHeader(1)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "AUTO - DELETE"
      TabPicture(1)   =   "frmOptions.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblHeader(0)"
      Tab(1).Control(1)=   "Frame1"
      Tab(1).Control(2)=   "Frame4"
      Tab(1).Control(3)=   "cmdAutoDeleteRefresh"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "KEYWORD"
      TabPicture(2)   =   "frmOptions.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lblHeader(2)"
      Tab(2).Control(1)=   "Frame5"
      Tab(2).Control(2)=   "Frame6"
      Tab(2).Control(3)=   "cmdKeywordRefresh"
      Tab(2).ControlCount=   4
      TabCaption(3)   =   "DELETED MAIL"
      TabPicture(3)   =   "frmOptions.frx":0496
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "cmdFillDeleted"
      Tab(3).Control(1)=   "DTP1"
      Tab(3).Control(2)=   "txtFilter"
      Tab(3).Control(3)=   "cmdFilter"
      Tab(3).Control(4)=   "lstDeletedEmails"
      Tab(3).Control(5)=   "Label5"
      Tab(3).Control(6)=   "Label4"
      Tab(3).ControlCount=   7
      TabCaption(4)   =   "SETTINGS"
      TabPicture(4)   =   "frmOptions.frx":04B2
      Tab(4).ControlEnabled=   -1  'True
      Tab(4).Control(0)=   "lblHeader(4)"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "Frame7"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "Frame8"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "Frame9"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).ControlCount=   4
      Begin VB.CommandButton cmdKeywordRefresh 
         Caption         =   "Refresh"
         Height          =   255
         Left            =   -68400
         TabIndex        =   59
         Top             =   600
         Width           =   855
      End
      Begin VB.CommandButton cmdAutoDeleteRefresh 
         Caption         =   "Refresh"
         Height          =   255
         Left            =   -68400
         TabIndex        =   58
         Top             =   600
         Width           =   855
      End
      Begin VB.CommandButton cmdExclusionRefresh 
         Caption         =   "Refresh"
         Height          =   255
         Left            =   -68400
         TabIndex        =   57
         Top             =   600
         Width           =   855
      End
      Begin VB.Frame Frame9 
         Height          =   975
         Left            =   360
         TabIndex        =   45
         Top             =   5760
         Width           =   7095
         Begin VB.CheckBox chkCheckMailWhenInTray 
            Caption         =   "Check Email When Running In Tray"
            Height          =   255
            Left            =   240
            TabIndex        =   47
            Top             =   405
            Width           =   3015
         End
         Begin VB.ComboBox cboCheckMailFrequencyMinutes 
            Height          =   315
            Left            =   5280
            TabIndex        =   46
            Text            =   "Combo1"
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label3 
            Caption         =   "Frequency ( Minutes )"
            Height          =   255
            Left            =   3600
            TabIndex        =   48
            Top             =   405
            Width           =   1575
         End
      End
      Begin VB.Frame Frame8 
         Height          =   2415
         Left            =   360
         TabIndex        =   38
         Top             =   840
         Width           =   7095
         Begin VB.CheckBox chkShowExclusionInListview 
            Caption         =   "Show Exclusion List Emails In List"
            Height          =   255
            Left            =   240
            TabIndex        =   56
            Top             =   240
            Value           =   1  'Checked
            Width           =   3855
         End
         Begin VB.CheckBox chkAutoDeleteOnCheck 
            Caption         =   "Auto-Delete Bad Emails When Doing Mail Check"
            Enabled         =   0   'False
            Height          =   255
            Left            =   6480
            TabIndex        =   55
            Top             =   240
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.CheckBox chkOpenEmailFromFriends 
            Caption         =   "Open Email Program If Only Emails From Friends When Doing Mail Check"
            Height          =   255
            Left            =   240
            TabIndex        =   42
            Top             =   600
            Value           =   1  'Checked
            Width           =   5655
         End
         Begin VB.CheckBox chkOpenEmailAfterProcess 
            Caption         =   "Open Email Program After Manually Processing"
            Height          =   255
            Left            =   240
            TabIndex        =   41
            Top             =   960
            Value           =   1  'Checked
            Width           =   3855
         End
         Begin VB.TextBox txtDefaultEmailProgram 
            Height          =   285
            Left            =   240
            Locked          =   -1  'True
            TabIndex        =   40
            Text            =   "C:\Program Files\Outlook Express\msimn.exe"
            Top             =   1800
            Width           =   6135
         End
         Begin VB.CommandButton cmdChangeDefaultEmailProgram 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   6480
            TabIndex        =   39
            Top             =   1800
            Width           =   375
         End
         Begin VB.Label Label2 
            Caption         =   "Email Program"
            Height          =   255
            Left            =   240
            TabIndex        =   44
            Top             =   1560
            Width           =   1335
         End
      End
      Begin VB.Frame Frame7 
         Height          =   2295
         Left            =   360
         TabIndex        =   33
         Top             =   3360
         Width           =   7095
         Begin VB.CheckBox chkPlayUnknownWaveForMail 
            Caption         =   "Play Wave On Mail From Friends Or Unknown"
            Height          =   255
            Left            =   240
            TabIndex        =   53
            Top             =   600
            Width           =   3735
         End
         Begin VB.TextBox txtUnknownWaveFile 
            Height          =   285
            Left            =   240
            Locked          =   -1  'True
            TabIndex        =   52
            Top             =   1800
            Width           =   6135
         End
         Begin VB.CommandButton cmdChangeUnknownWaveFile 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   6480
            TabIndex        =   51
            Top             =   1800
            Width           =   375
         End
         Begin VB.CommandButton cmdChangeFriendWaveFile 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   6480
            TabIndex        =   37
            Top             =   1200
            Width           =   375
         End
         Begin VB.TextBox txtFriendWaveFile 
            Height          =   285
            Left            =   240
            Locked          =   -1  'True
            TabIndex        =   36
            Top             =   1200
            Width           =   6135
         End
         Begin VB.CheckBox chkPlayFriendWaveForMail 
            Caption         =   "Play Wave On Mail From Friends"
            Height          =   255
            Left            =   240
            TabIndex        =   35
            Top             =   240
            Width           =   2655
         End
         Begin VB.CommandButton cmdEmailColors 
            Caption         =   "Email Colors"
            Height          =   375
            Left            =   4800
            TabIndex        =   34
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label6 
            Caption         =   "Friend Or Unknown Wave File"
            Height          =   255
            Left            =   240
            TabIndex        =   54
            Top             =   1560
            Width           =   2295
         End
         Begin VB.Label Label1 
            Caption         =   "Friend Wave File"
            Height          =   255
            Left            =   240
            TabIndex        =   43
            Top             =   960
            Width           =   1335
         End
      End
      Begin VB.CommandButton cmdFillDeleted 
         Caption         =   "Show All"
         Height          =   300
         Left            =   -68400
         TabIndex        =   32
         Top             =   6960
         Width           =   975
      End
      Begin MSComCtl2.DTPicker DTP1 
         Height          =   300
         Left            =   -71040
         TabIndex        =   31
         Top             =   6960
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   529
         _Version        =   393216
         Format          =   24772609
         CurrentDate     =   37958
      End
      Begin VB.TextBox txtFilter 
         Height          =   285
         Left            =   -74520
         TabIndex        =   30
         Top             =   6960
         Width           =   2055
      End
      Begin VB.CommandButton cmdFilter 
         Caption         =   "Find"
         Height          =   300
         Left            =   -72360
         TabIndex        =   29
         Top             =   6960
         Width           =   735
      End
      Begin VB.ListBox lstDeletedEmails 
         Height          =   5910
         ItemData        =   "frmOptions.frx":04CE
         Left            =   -74760
         List            =   "frmOptions.frx":04D0
         MultiSelect     =   2  'Extended
         TabIndex        =   28
         ToolTipText     =   "Right Click For Options"
         Top             =   600
         Width           =   7575
      End
      Begin VB.Frame Frame6 
         Caption         =   "Select To Remove From List..."
         Height          =   6255
         Left            =   -70920
         TabIndex        =   24
         Top             =   840
         Width           =   3375
         Begin VB.CommandButton cmdKeywordRemove 
            Caption         =   "Remove Selected KeyWords"
            Height          =   375
            Left            =   240
            TabIndex        =   26
            Top             =   5760
            Width           =   2895
         End
         Begin VB.ListBox lstKeyword 
            Height          =   5325
            ItemData        =   "frmOptions.frx":04D2
            Left            =   240
            List            =   "frmOptions.frx":04D4
            MultiSelect     =   2  'Extended
            Sorted          =   -1  'True
            TabIndex        =   25
            Top             =   360
            Width           =   2895
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Add Full Or Partial Keywords..."
         Height          =   6255
         Left            =   -74760
         TabIndex        =   19
         Top             =   840
         Width           =   3735
         Begin VB.TextBox txtKeyword 
            Height          =   285
            Left            =   120
            TabIndex        =   21
            Top             =   2520
            Width           =   2415
         End
         Begin VB.CommandButton cmdKeywordAdd 
            Caption         =   "Add"
            Height          =   375
            Left            =   2640
            TabIndex        =   20
            Top             =   2520
            Width           =   855
         End
         Begin VB.Label lblKeyword 
            Alignment       =   2  'Center
            Caption         =   "Enter Full Or Partial Keywords To Be Searched In Subject Line Of Email. Emails Containing A Keyword Will Be Auto-Deleted"
            Height          =   735
            Left            =   120
            TabIndex        =   23
            Top             =   720
            Width           =   3495
         End
         Begin VB.Label lblKeyword2 
            Height          =   2895
            Left            =   120
            TabIndex        =   22
            Top             =   3120
            Width           =   3495
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Add Full Or Partial Email Addresses..."
         Height          =   6255
         Left            =   -74760
         TabIndex        =   13
         Top             =   840
         Width           =   3735
         Begin VB.CommandButton cmdAutoDeleteAdd 
            Caption         =   "Add"
            Height          =   375
            Left            =   2640
            TabIndex        =   15
            Top             =   2520
            Width           =   855
         End
         Begin VB.TextBox txtAutoDelete 
            Height          =   285
            Left            =   120
            TabIndex        =   14
            Top             =   2520
            Width           =   2415
         End
         Begin VB.Label lblAutoDelete2 
            Height          =   1815
            Left            =   120
            TabIndex        =   17
            Top             =   1080
            Width           =   3495
         End
         Begin VB.Label lblAutoDelete 
            Alignment       =   2  'Center
            Caption         =   "Enter Full Or Partial Email Addresses To Be Auto-Deleted"
            Height          =   495
            Left            =   120
            TabIndex        =   16
            Top             =   720
            Width           =   3495
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Select To Remove From List..."
         Height          =   6255
         Left            =   -70920
         TabIndex        =   10
         Top             =   840
         Width           =   3375
         Begin VB.ListBox lstAutoDelete 
            Height          =   5325
            ItemData        =   "frmOptions.frx":04D6
            Left            =   240
            List            =   "frmOptions.frx":04D8
            MultiSelect     =   2  'Extended
            Sorted          =   -1  'True
            TabIndex        =   12
            Top             =   360
            Width           =   2895
         End
         Begin VB.CommandButton cmdAutoDeleteRemove 
            Caption         =   "Remove Selected Addresses"
            Height          =   375
            Left            =   240
            TabIndex        =   11
            Top             =   5760
            Width           =   2895
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Select To Remove From List..."
         Height          =   6255
         Left            =   -70920
         TabIndex        =   6
         Top             =   840
         Width           =   3375
         Begin VB.ListBox lstExclusion 
            Height          =   5325
            ItemData        =   "frmOptions.frx":04DA
            Left            =   240
            List            =   "frmOptions.frx":04DC
            MultiSelect     =   2  'Extended
            Sorted          =   -1  'True
            TabIndex        =   8
            Top             =   360
            Width           =   2895
         End
         Begin VB.CommandButton cmdExclusionRemove 
            Caption         =   "Remove Selected Addresses"
            Height          =   375
            Left            =   240
            TabIndex        =   7
            Top             =   5760
            Width           =   2895
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Add Email Addresses..."
         Height          =   6255
         Left            =   -74760
         TabIndex        =   2
         Top             =   840
         Width           =   3735
         Begin VB.CommandButton cmdExclusionAdd 
            Caption         =   "Add"
            Height          =   375
            Left            =   2640
            TabIndex        =   4
            Top             =   2520
            Width           =   855
         End
         Begin VB.TextBox txtExclusion 
            Height          =   285
            Left            =   120
            TabIndex        =   3
            Top             =   2520
            Width           =   2415
         End
         Begin VB.Label lblExclusion 
            Alignment       =   2  'Center
            Caption         =   "Enter Full Or Partial Email Addresses To Be Excluded From Deletion"
            Height          =   525
            Left            =   120
            TabIndex        =   5
            Top             =   720
            Width           =   3495
         End
      End
      Begin VB.Label Label5 
         Caption         =   "Search By Date Deleted"
         Height          =   255
         Left            =   -71040
         TabIndex        =   50
         Top             =   6720
         Width           =   1935
      End
      Begin VB.Label Label4 
         Caption         =   "Enter Text To Search"
         Height          =   255
         Left            =   -74520
         TabIndex        =   49
         Top             =   6720
         Width           =   2415
      End
      Begin VB.Label lblHeader 
         Alignment       =   2  'Center
         Caption         =   "KEYWORD BLACKLIST"
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   2
         Left            =   -74760
         TabIndex        =   27
         Top             =   600
         Width           =   7215
      End
      Begin VB.Label lblHeader 
         Alignment       =   2  'Center
         Caption         =   "AUTO - DELETE LIST"
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   0
         Left            =   -74760
         TabIndex        =   18
         Top             =   600
         Width           =   7215
      End
      Begin VB.Label lblHeader 
         Alignment       =   2  'Center
         Caption         =   "EXCLUSION LIST"
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   1
         Left            =   -74760
         TabIndex        =   9
         Top             =   600
         Width           =   7095
      End
      Begin VB.Label lblHeader 
         Alignment       =   2  'Center
         Caption         =   "SETTINGS"
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   4
         Left            =   360
         TabIndex        =   1
         Top             =   600
         Width           =   7095
      End
   End
   Begin VB.Menu zHid1 
      Caption         =   "Hid1"
      Visible         =   0   'False
      Begin VB.Menu mnuDeleteFromList 
         Caption         =   "Delete From List"
      End
      Begin VB.Menu zSep01 
         Caption         =   "-"
      End
      Begin VB.Menu zRemoveFromAutoDelete 
         Caption         =   "Remove From Auto-Delete"
         Begin VB.Menu mnuRemoveAutoDeleteAddress 
            Caption         =   "Full Email Address  -  spam@virtual.domain.com"
         End
         Begin VB.Menu mnuRemoveAutoDeleteFull 
            Caption         =   "Full Domain  -  virtual.domain.com"
         End
         Begin VB.Menu mnuRemoveAutoDeletePartial 
            Caption         =   "Partial Domain - domain.com"
         End
         Begin VB.Menu mnuRemoveAutoDeleteSuffix 
            Caption         =   "Suffix - .com   BE CAREFUL HERE"
         End
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum emAddress
    FullAddress = 1
    FullDomain = 2
    PartialDomain = 3
    Suffix = 4
End Enum

Dim strDTPBeforeDropdownValue As String


Private Sub cboCheckMailFrequencyMinutes_Click()
    g_intCheckMailFrequencyMinutes = cboCheckMailFrequencyMinutes.Text
End Sub

Private Sub chkAutoDeleteOnCheck_Click()
    g_intAutoDeleteOnCheck = chkAutoDeleteOnCheck.Value
End Sub

Private Sub chkCheckMailWhenInTray_Click()
    g_intCheckMailWhenInTray = chkCheckMailWhenInTray.Value
End Sub

Private Sub chkOpenEmailAfterProcess_Click()
    g_intOpenEmailAfterProcess = chkOpenEmailAfterProcess.Value
End Sub

Private Sub chkOpenEmailFromFriends_Click()
     g_intOpenEmailFromFriends = chkOpenEmailFromFriends.Value
End Sub

Private Sub chkPlayUnknownWaveForMail_Click()
    g_intPlayUnknownWaveForMail = chkPlayUnknownWaveForMail.Value
End Sub

Private Sub chkShowExclusionInListview_Click()
     g_intShowExclusionInListview = chkShowExclusionInListview.Value
End Sub

Private Sub cmdAutoDeleteRefresh_Click()
    LoadList_AutoDelete
    txtAutoDelete.SetFocus
End Sub

Private Sub cmdChangeDefaultEmailProgram_Click()
    
    Dim strFile As String
    
    With CD1
        .CancelError = False
        .DialogTitle = "Find Email Program"
        .DefaultExt = ".exe"
        .Filter = "Program Files (*.exe)|*.exe"
        .ShowOpen
        strFile = .FileName
    End With
    
    If strFile <> "" Then
        txtDefaultEmailProgram = strFile
        g_strDefaultEmailProgram = strFile
    End If
    
End Sub

Private Sub cmdChangeFriendWaveFile_Click()
    
    Dim strFile As String
    
    With CD1
        .CancelError = False
        .DialogTitle = "Find Wave File"
        .DefaultExt = ".wav"
        .Filter = "Wave Files (*.wav)|*.wav"
        .ShowOpen
        strFile = .FileName
    End With
    
    If strFile <> "" Then
        txtFriendWaveFile = strFile
        g_strFriendWaveFile = strFile
    End If
    
End Sub

Private Sub cmdChangeUnknownWaveFile_Click()
    
    Dim strFile As String
    
    With CD1
        .CancelError = False
        .DialogTitle = "Find Wave File"
        .DefaultExt = ".wav"
        .Filter = "Wave Files (*.wav)|*.wav"
        .ShowOpen
        strFile = .FileName
    End With
    
    If strFile <> "" Then
        txtUnknownWaveFile = strFile
        g_strUnknownWaveFile = strFile
    End If
    
End Sub

Private Sub cmdEmailColors_Click()
    frmColors.Show vbModal
End Sub

Private Sub cmdExclusionRefresh_Click()
    LoadList_Exclusion
    txtExclusion.SetFocus
End Sub

Private Sub cmdFillDeleted_Click()
    LoadList_DeletedEmails
End Sub

Private Sub cmdFilter_Click()
    Call LoadFileIntoListbox(g_DeletedEmailFile, lstDeletedEmails, txtFilter.Text)
End Sub

Private Sub chkPlayFriendWaveForMail_Click()
    g_intPlayFriendWaveForMail = chkPlayFriendWaveForMail.Value
End Sub

Private Sub cmdKeywordRefresh_Click()
    LoadList_KeyWord
    txtKeyword.SetFocus
End Sub

Private Sub DTP1_CloseUp()
Dim strTemp As String

    If DTP1.Value <> strDTPBeforeDropdownValue Then
    
        strTemp = DTP1.Value
        strTemp = Format(strTemp, "mmm d, yyyy")
        strDTPBeforeDropdownValue = DTP1.Value
        
        ' Load deleted email log using date selected
        Call LoadFileIntoListbox(g_DeletedEmailFile, lstDeletedEmails, strTemp)
        
    End If

End Sub

Private Sub Form_Load()
    
    Dim i As Integer

    ' Load lists into listboxes
    LoadList_AutoDelete
    LoadList_Exclusion
    LoadList_KeyWord
    LoadList_DeletedEmails
    
    If g_bolFirstRun Then
        SSTab1.Tab = 4
    Else
        SSTab1.Tab = 0
    End If
    
    DTP1.Value = Now
    
    ' Add minutes
    For i = 1 To 60
        cboCheckMailFrequencyMinutes.AddItem i
    Next
    
    lblKeyword2 = "If Keyword Does Not Have A Wildcard, It Will Be Matched Against Whole Subject." & vbCrLf & vbCrLf & "Keywords With Wildcards:" & vbCrLf & "* At Beginning Will Compare To End Of Subject" & vbCrLf & "* At End Will Compare To Beginning Of Subject" & vbCrLf & "* At Beginning And End Will Searched For Anywhere In Subject" & vbCrLf & vbCrLf & "Partial Key Word Examples:" & vbCrLf & Space(6) & "*viagra*" & vbCrLf & Space(6) & "*Free" & vbCrLf & Space(6) & "RE: Your Free*" & vbCrLf & Space(6) & "*Porn Passwords"
    
    ' Put program settings into their controls
    Call ProcessProgramSettings
    
End Sub

Private Sub cmdAutoDeleteAdd_Click()
    Dim i As Integer
    Dim sTemp As String
    
    sTemp = Trim(txtAutoDelete)
    
    If sTemp = "" Then
        MsgBox "Cant Add Empty Spaces", vbOKOnly, "Enter Email Address"
        txtAutoDelete = ""
        txtAutoDelete.SetFocus
        Exit Sub
    End If
    
    If (InStr(1, sTemp, ".") = 0) Then
        MsgBox "Email Addresses Must Contain A Dot '.'", vbOKOnly, "Check Input"
        txtAutoDelete.SetFocus
        Exit Sub
    End If
    
    ' See if entry is already on list
    For i = 0 To UBound(g_arrAutoDeleteList)
        If sTemp = g_arrAutoDeleteList(i) Then
            MsgBox "Address Already Added To Auto-Delete", vbOKOnly, "Already In List"
            txtAutoDelete.SetFocus
            Exit Sub
        End If
    Next
    
    ' Add entry to the list & array
    AddToAutoDeleteList sTemp
    
    ' Check array to make sure it was added, then reload the list
    If g_arrAutoDeleteList(UBound(g_arrAutoDeleteList)) = sTemp Then
        LoadList_AutoDelete
        txtAutoDelete = ""
        txtAutoDelete.SetFocus
    Else
        MsgBox "Error: Address Not Added To Auto-Delete" & vbCrLf & vbCrLf & "Please Try Again", vbOKOnly, "Error"
        cmdAutoDeleteAdd.SetFocus
    End If

End Sub

Private Sub cmdAutoDeleteRemove_Click()
Dim i As Integer
Dim x As Long
Dim b As Boolean
Dim arrTemp() As String

    ReDim Preserve arrTemp(0)

    ' Loop through & make array of selected items
    For i = 0 To lstAutoDelete.ListCount - 1
        If lstAutoDelete.Selected(i) = True Then
            x = CLng(lstAutoDelete.ItemData(i))
            ReDim Preserve arrTemp(UBound(arrTemp) + 1)
            arrTemp(UBound(arrTemp)) = g_arrAutoDeleteList(x)
        End If
    Next
    
    ' If any, remove from array & list
    If UBound(arrTemp) > 0 Then
        For i = 1 To UBound(arrTemp)
            RemoveFromAutoDeleteList arrTemp(i)
            b = True
        Next
    End If
    
    ' Reload array & list
    If b = True Then
        GetAutoDeleteList
        LoadList_AutoDelete
    End If

End Sub

Private Sub LoadList_AutoDelete()
Dim i As Integer

    lstAutoDelete.Clear
    
    ' Load array into listbox
    For i = 0 To UBound(g_arrAutoDeleteList)
        If g_arrAutoDeleteList(i) <> "" Then
            lstAutoDelete.AddItem g_arrAutoDeleteList(i)
            lstAutoDelete.ItemData(lstAutoDelete.NewIndex) = i
        End If
    Next
    
End Sub


Private Sub cmdExclusionAdd_Click()
    Dim i As Integer
    Dim sTemp As String
    
    sTemp = Trim(txtExclusion)
    
    If sTemp = "" Then
        MsgBox "Cant Add Empty Spaces", vbOKOnly, "Enter Full Email Address"
        txtAutoDelete = ""
        txtAutoDelete.SetFocus
        Exit Sub
    End If
    
    If (InStr(1, sTemp, ".") = 0) Or (InStr(1, sTemp, "@") = 0) Then
        MsgBox "Partial Email Address - Check The Address", vbOKOnly, "Check Input"
        txtAutoDelete.SetFocus
        Exit Sub
    End If
    
    ' See if entry is already on list
    For i = 0 To UBound(g_arrAutoDeleteList)
        If sTemp = g_arrAutoDeleteList(i) Then
            MsgBox "Address Already Added To Exclusion List", vbOKOnly, "Already In List"
            txtAutoDelete.SetFocus
            Exit Sub
        End If
    Next
    
    ' Add entry to the list & array
    AddToExclusionList sTemp
    
    ' Check array to make sure it was added, then reload the list
    If g_arrExclusionList(UBound(g_arrExclusionList)) = sTemp Then
        LoadList_Exclusion
        txtExclusion = ""
        txtExclusion.SetFocus
    Else
        MsgBox "Error: Address Not Added To Exclusion List" & vbCrLf & vbCrLf & "Please Try Again", vbOKOnly, "Error"
        cmdExclusionAdd.SetFocus
    End If

End Sub

Private Sub cmdExclusionRemove_Click()
Dim i As Integer
Dim x As Long
Dim b As Boolean
Dim arrTemp() As String

    ReDim Preserve arrTemp(0)

    ' Loop through & make array of selected items
    For i = 0 To lstExclusion.ListCount - 1
        If lstExclusion.Selected(i) = True Then
            x = CLng(lstExclusion.ItemData(i))
            ReDim Preserve arrTemp(UBound(arrTemp) + 1)
            arrTemp(UBound(arrTemp)) = g_arrExclusionList(x)
        End If
    Next
    
    ' If any, remove from array & list
    If UBound(arrTemp) > 0 Then
        For i = 1 To UBound(arrTemp)
            RemoveFromExclusionList arrTemp(i)
            b = True
        Next
    End If
    
    ' Reload array & list
    If b = True Then
        GetExclusionList
        LoadList_Exclusion
    End If

End Sub

Private Sub LoadList_Exclusion()
Dim i As Integer

    lstExclusion.Clear
    
    ' Load array into listbox
    For i = 0 To UBound(g_arrExclusionList)
        If g_arrExclusionList(i) <> "" Then
            lstExclusion.AddItem g_arrExclusionList(i)
            lstExclusion.ItemData(lstExclusion.NewIndex) = i
        End If
    Next
    
End Sub

Private Sub cmdKeyWordAdd_Click()

    Call AddKeyWord

End Sub

Private Sub cmdKeyWordRemove_Click()

Dim i As Integer
Dim x As Long
Dim b As Boolean
Dim arrTemp() As String

    ReDim Preserve arrTemp(0)

    ' Loop through & make array of selected items
    For i = 0 To lstKeyword.ListCount - 1
        If lstKeyword.Selected(i) = True Then
            x = CLng(lstKeyword.ItemData(i))
            ReDim Preserve arrTemp(UBound(arrTemp) + 1)
            arrTemp(UBound(arrTemp)) = g_arrKeyWordList(x)
        End If
    Next

    ' If any, remove from array & list
    If UBound(arrTemp) > 0 Then
        For i = 1 To UBound(arrTemp)
            RemoveFromKeyWordList arrTemp(i)
            b = True
        Next
    End If

    ' Reload array & list
    If b = True Then
        GetKeyWordList
        LoadList_KeyWord
    End If

End Sub

Private Sub LoadList_KeyWord()
Dim i As Integer

    lstKeyword.Clear

    ' Load array into listbox
    For i = 0 To UBound(g_arrKeyWordList)
        If g_arrKeyWordList(i) <> "" Then
            lstKeyword.AddItem g_arrKeyWordList(i)
            lstKeyword.ItemData(lstKeyword.NewIndex) = i
        End If
    Next

End Sub

Private Sub AddKeyWord()

    Dim i As Integer
    Dim sTemp As String

    sTemp = Trim(txtKeyword)

    If sTemp = "" Then
        MsgBox "Cant Add Empty Spaces", vbOKOnly, "Enter A KeyWord"
        txtKeyword = ""
        txtKeyword.SetFocus
        Exit Sub
    End If

    ' See if entry is already on list
    For i = 0 To UBound(g_arrKeyWordList)
        If sTemp = g_arrKeyWordList(i) Then
            MsgBox "KeyWord Already Added To Key Word List", vbOKOnly, "Already In List"
            txtKeyword.SetFocus
            Exit Sub
        End If
    Next

    ' Add entry to the list & array
    AddToKeyWordList sTemp

    ' Check array to make sure it was added, then reload the list
    If g_arrKeyWordList(UBound(g_arrKeyWordList)) = sTemp Then
        LoadList_KeyWord
        txtKeyword = ""
        txtKeyword.SetFocus
    Else
        MsgBox "Error: Key Word Not Added To Key Word List" & vbCrLf & vbCrLf & "Please Try Again", vbOKOnly, "Error"
        cmdKeywordAdd.SetFocus
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Me.Hide
    frmEmailChecker.Show
End Sub

Private Sub lstDeletedEmails_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = vbRightButton Then
        ' Show popup menu
        PopupMenu zHid1
    End If
    
End Sub

Private Sub mnuDeleteFromList_Click()
    Call DeleteEmailFromFileAndList
End Sub

Private Sub mnuRemoveAutoDeleteAddress_Click()
    Call ParseDeletedListAndRemoveFromAutoDelete(emAddress.FullAddress)
End Sub

Private Sub mnuRemoveAutoDeleteFull_Click()
    Call ParseDeletedListAndRemoveFromAutoDelete(emAddress.FullDomain)
End Sub

Private Sub mnuRemoveAutoDeletePartial_Click()
    Call ParseDeletedListAndRemoveFromAutoDelete(emAddress.PartialDomain)
End Sub

Private Sub mnuRemoveAutoDeleteSuffix_Click()
    Call ParseDeletedListAndRemoveFromAutoDelete(emAddress.Suffix)
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    
    If Me.Visible = True Then
    
        ' Set focus
        If SSTab1.Tab = 0 Then
                txtExclusion.SetFocus
        ElseIf SSTab1.Tab = 1 Then
                txtAutoDelete.SetFocus
        ElseIf SSTab1.Tab = 2 Then
                txtKeyword.SetFocus
        ElseIf SSTab1.Tab = 3 Then
                txtFilter.SetFocus
        End If
    
    End If
    
End Sub

Private Sub txtFilter_GotFocus()
    Call highLight
End Sub

Private Sub txtFilter_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then ' Enter key pressed
        ' Load deleted emails log with emails containing search word
        Call LoadFileIntoListbox(g_DeletedEmailFile, lstDeletedEmails, txtFilter.Text)
        KeyAscii = 0
        Call highLight
    End If
    
End Sub

Private Sub txtKeyWord_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then ' Enter key pressed
        Call AddKeyWord
        KeyAscii = 0
    End If

End Sub

Private Sub LoadList_DeletedEmails()
    ' Load deleted emails log
    Call LoadFileIntoListbox(g_DeletedEmailFile, lstDeletedEmails)
    ' Put horizontal scroll on listbox
    Call PutHScrollOnList(Me, lstDeletedEmails)
End Sub


Private Sub ParseDeletedListAndRemoveFromAutoDelete(intType As Integer)

    Dim i As Integer
    Dim strTemp As String
    Dim intCount As Integer
    
    ' Loop thrugh & get selected emails
    For i = 0 To lstDeletedEmails.ListCount - 1
    
        If lstDeletedEmails.Selected(i) Then
        
            ' Get full or partial address
            strTemp = GetAddressFromString(lstDeletedEmails.List(i), intType)
        
            If strTemp <> "" Then
                ' Remove it from the list & array
                RemoveFromAutoDeleteList strTemp
            End If
            
            intCount = intCount + 1
    
        End If
        
    Next
    
    If intCount > 0 Then
       ' Remove it from the list & listbox
        DeleteEmailFromFileAndList
    End If

End Sub

Private Function GetAddressFromString(strText As String, intEmailType As Integer) As String

Dim i As Integer
Dim x As Integer
Dim strTemp As String
Dim strFullAddress As String
Dim strFullDomain As String
Dim strPartialDomain As String
Dim strSuffix As String

    GetAddressFromString = ""
    
    ' Get full & partial address out of line from from deleted log
    i = InStr(1, strText, "---")
    
    If i > 0 Then
        i = i + 3
        x = InStr(i, strText, "---")
        
        If x < i Then
            GoTo DoExitProcedure
        End If
        
    Else
        GoTo DoExitProcedure
    End If
    
    strTemp = Trim(Mid(strText, i, (x - i)))
    Debug.Print strTemp
    
    
    i = (InStr(1, strTemp, "<", vbTextCompare) + 1)
    
    If i > 1 Then
        x = (InStr(i, strTemp, ">", vbTextCompare))
        strFullAddress = Mid(strTemp, i, (x - i))
    Else
        GoTo DoExitProcedure
    End If
    
    strFullDomain = Mid(strFullAddress, (InStr(1, strFullAddress, "@", vbTextCompare) + 1))
    
    i = (InStrRev(strFullDomain, "."))
    strSuffix = Mid(strFullDomain, i)
    
    x = (InStrRev(strFullDomain, ".", (i - 1)))
    If x <> 0 Then
        strPartialDomain = Mid(strFullDomain, (x + 1))
    Else
        strPartialDomain = strFullDomain
    End If
    
    ' Base on type passed to procedure
    Select Case intEmailType
    
        Case emAddress.FullAddress
            GetAddressFromString = strFullAddress
            
        Case emAddress.FullDomain
            GetAddressFromString = strFullDomain
            
        Case emAddress.PartialDomain
            GetAddressFromString = strPartialDomain
            
        Case emAddress.Suffix
            GetAddressFromString = strSuffix
            
        Case emAddress.FullAddress
            GetAddressFromString = strFullAddress
            
    End Select
    
DoExitProcedure:

    Exit Function

End Function

Private Sub DeleteEmailFromFileAndList()
    Dim i As Integer
    Dim x As Integer
    Dim y As Integer
    Dim strTemp As String
    Dim arrIndexes() As Long
    Dim arrDelete() As String
    
    ReDim Preserve arrIndexes(0)
    ReDim Preserve arrDelete(0)
    
    ' Loop though list & get selected emails
    For i = 0 To lstDeletedEmails.ListCount - 1
    
        If lstDeletedEmails.Selected(i) Then
        
            ReDim Preserve arrDelete(UBound(arrDelete) + 1)
            arrDelete(UBound(arrDelete)) = lstDeletedEmails.List(i)
        
            ReDim Preserve arrIndexes(UBound(arrIndexes) + 1)
            arrIndexes(UBound(arrIndexes)) = i
    
        End If
        
    Next
    
    If UBound(arrDelete) > 0 Then
    
        ' Remove from the log
        Call DeleteArrayOfTextLinesFromFile(arrDelete, g_DeletedEmailFile)
        
    End If
    
    
    If UBound(arrIndexes) > 0 Then
    
        For i = UBound(arrIndexes) To 1 Step -1
    
            ' Remove from the listbox
            lstDeletedEmails.RemoveItem arrIndexes(i)
            
        Next
        
    End If

End Sub

Private Sub ProcessProgramSettings()

    ' Put program settings into their controls

    chkShowExclusionInListview.Value = g_intShowExclusionInListview
    chkPlayFriendWaveForMail.Value = g_intPlayFriendWaveForMail
    txtFriendWaveFile.Text = g_strFriendWaveFile
    chkPlayUnknownWaveForMail.Value = g_intPlayUnknownWaveForMail
    txtUnknownWaveFile.Text = g_strUnknownWaveFile
    chkAutoDeleteOnCheck.Value = g_intAutoDeleteOnCheck
    chkOpenEmailAfterProcess.Value = g_intOpenEmailAfterProcess
    chkOpenEmailFromFriends.Value = g_intOpenEmailFromFriends
    txtDefaultEmailProgram.Text = g_strDefaultEmailProgram
    chkCheckMailWhenInTray.Value = g_intCheckMailWhenInTray
    cboCheckMailFrequencyMinutes.Text = g_intCheckMailFrequencyMinutes

End Sub
