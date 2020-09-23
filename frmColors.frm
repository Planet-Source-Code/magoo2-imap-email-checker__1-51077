VERSION 5.00
Begin VB.Form frmColors 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Change Email Colors"
   ClientHeight    =   2535
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   3885
   Icon            =   "frmColors.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   3885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdChangeAutoDelete 
      Caption         =   "Auto-Delete"
      Height          =   375
      Left            =   2640
      TabIndex        =   12
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton cmdChangeBackGround 
      Caption         =   "BackGround"
      Height          =   375
      Left            =   2640
      TabIndex        =   11
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton cmdChangeFriend 
      Caption         =   "Friend"
      Height          =   375
      Left            =   1440
      TabIndex        =   10
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton cmdChangeUnknown 
      Caption         =   "Unknown"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   2640
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label lblFriend 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Friend Email"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   2415
   End
   Begin VB.Label lblAutoDelete 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Auto-Delete Email"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   1320
      Width           =   2415
   End
   Begin VB.Label lblUnknown 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Unknown Email"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   2415
   End
   Begin VB.Label lblAutoDelete 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Auto-Delete Email"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   2415
   End
   Begin VB.Label lblFriend 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Friend Email"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   2415
   End
   Begin VB.Label lblFriend 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Friend Email"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   2415
   End
   Begin VB.Label lblUnknown 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Unknown Email"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "frmColors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()
    Unload Me
End Sub

Private Sub cmdChangeAutoDelete_Click()

    On Error GoTo errhandler:
    
    Dim i As Integer
    Dim lngColor As Long
    
    With frmOptions.CD1
        .CancelError = True
        .ShowColor
        lngColor = .Color
    End With
    
    ' If this color matches one already in use, can not use again
    If (lngColor = g_lngUnknownEmailColor) Then
        MsgBox "This Color Is Already In Use As The Unknown Email Color"
        GoTo DoExit
    End If
    
    If (lngColor = g_lngExclusionEmailColor) Then
        MsgBox "This Color Is Already In Use As The Exclusion Color"
        GoTo DoExit
    End If
    
    For i = 0 To 1
        lblAutoDelete(i).ForeColor = lngColor
    Next
    
DoExit:
    
    Exit Sub
    
errhandler:
    Exit Sub

End Sub

Private Sub cmdChangeBackGround_Click()

    On Error GoTo errhandler:
    
    Dim i As Integer
    Dim lngColor As Long
    
    With frmOptions.CD1
        .CancelError = True
        .ShowColor
        lngColor = .Color
    End With
    
    For i = 0 To 1
        lblUnknown(i).BackColor = lngColor
    Next
    
    For i = 0 To 1
        lblAutoDelete(i).BackColor = lngColor
    Next
    
    For i = 0 To 2
        lblFriend(i).BackColor = lngColor
    Next
    
    frmEmailChecker.ListView1.BackColor = lngColor
    
DoExit:
    
    Exit Sub
    
errhandler:
    Exit Sub

End Sub

Private Sub cmdChangeFriend_Click()

    On Error GoTo errhandler:
    
    Dim i As Integer
    Dim lngColor As Long
    
    With frmOptions.CD1
        .CancelError = True
        .ShowColor
        lngColor = .Color
    End With
    
    ' If this color matches one already in use, can not use again
    If (lngColor = g_lngAutoDeleteEmailColor) Then
        MsgBox "This Color Is Already In Use As The Auto-Delete Color"
        GoTo DoExit
    End If
    
    If (lngColor = g_lngUnknownEmailColor) Then
        MsgBox "This Color Is Already In Use As The Unknown Email Color"
        GoTo DoExit
    End If
    
    For i = 0 To 2
        lblFriend(i).ForeColor = lngColor
    Next
    
DoExit:
    
    Exit Sub
    
errhandler:
    Exit Sub

End Sub

Private Sub cmdChangeUnknown_Click()

    On Error GoTo errhandler:
    
    Dim i As Integer
    Dim lngColor As Long
    
    With frmOptions.CD1
        .CancelError = True
        .ShowColor
        lngColor = .Color
    End With
    
    ' If this color matches one already in use, can not use again
    If (lngColor = g_lngAutoDeleteEmailColor) Then
        MsgBox "This Color Is Already In Use As The Auto-Delete Color"
        GoTo DoExit
    End If
    
    If (lngColor = g_lngExclusionEmailColor) Then
        MsgBox "This Color Is Already In Use As The Exclusion Color"
        GoTo DoExit
    End If
    
    For i = 0 To 1
        lblUnknown(i).ForeColor = lngColor
    Next
    
DoExit:
    
    Exit Sub
    
errhandler:
    Exit Sub

End Sub

Private Sub Form_Load()
    Dim i As Integer

    ' Color the textboxes
    For i = 0 To 1
        lblUnknown(i).BackColor = g_lngListBackColor
    Next
    
    For i = 0 To 1
        lblAutoDelete(i).BackColor = g_lngListBackColor
    Next
    
    For i = 0 To 2
        lblFriend(i).BackColor = g_lngListBackColor
    Next
    
    For i = 0 To 1
        lblUnknown(i).ForeColor = g_lngUnknownEmailColor
    Next
    
    For i = 0 To 1
        lblAutoDelete(i).ForeColor = g_lngAutoDeleteEmailColor
    Next
    
    For i = 0 To 2
        lblFriend(i).ForeColor = g_lngExclusionEmailColor
    Next
    
End Sub

Private Sub OKButton_Click()
    On Error Resume Next
    
    ' Save changes
    g_lngUnknownEmailColor = lblUnknown(0).ForeColor
    g_lngExclusionEmailColor = lblFriend(0).ForeColor
    g_lngListBackColor = lblAutoDelete(0).BackColor
    g_lngAutoDeleteEmailColor = lblAutoDelete(0).ForeColor
    
    Unload Me
    
End Sub
