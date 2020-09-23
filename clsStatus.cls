VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_LoginOK As Boolean
Private m_SelectInbox As Boolean
Private m_SearchAll As Boolean
Private m_FetchStarted As Boolean
Private m_LogoutSent As Boolean
Private m_DoSendLogout As Boolean
Private m_CheckingAddresses As Boolean
Private m_CurrentlyFetching As Boolean
Private m_FetchFinished As Boolean
Private m_SetFlagsStart As Boolean
Private m_SetFlagsFinished As Boolean
Private m_Closing As Boolean
Private m_TimerOn As Boolean
Private m_PreviousCountAll As Integer
Private m_PreviousCountFriend As Integer
Private m_MailProgramTitle As String
Private m_Idle As Boolean
Private m_Checking As Boolean
Private m_Processing As Boolean
Private m_MailProgramHandle As Long

Public Property Get MailProgramHandle() As Long
    MailProgramHandle = m_MailProgramHandle
End Property

Public Property Let MailProgramHandle(ByVal Value As Long)
    m_MailProgramHandle = Value
End Property

Public Property Get Processing() As Boolean
    Processing = m_Processing
End Property

Public Property Let Processing(ByVal Value As Boolean)
    m_Processing = Value
    
    If m_Processing Then
        m_Idle = False
        m_Checking = False
    End If
    
End Property

Public Property Get Checking() As Boolean
    Checking = m_Checking
End Property

Public Property Let Checking(ByVal Value As Boolean)
    m_Checking = Value
    
    If m_Checking Then
        m_Idle = False
        m_Processing = False
    End If
    
End Property

Public Property Get Idle() As Boolean
    Idle = m_Idle
End Property

Public Property Let Idle(ByVal Value As Boolean)
    m_Idle = Value
    
    If m_Idle Then
        m_Processing = False
        m_Checking = False
    End If
    
End Property

Public Property Get MailProgramTitle() As String
    MailProgramTitle = m_MailProgramTitle
End Property

Public Property Let MailProgramTitle(ByVal Value As String)
    m_MailProgramTitle = Value
End Property

Public Property Get PreviousCountFriend() As Integer
    PreviousCountFriend = m_PreviousCountFriend
End Property

Public Property Let PreviousCountFriend(ByVal Value As Integer)
    m_PreviousCountFriend = Value
End Property

Public Property Get PreviousCountAll() As Integer
    PreviousCountAll = m_PreviousCountAll
End Property

Public Property Let PreviousCountAll(ByVal Value As Integer)
    m_PreviousCountAll = Value
End Property

Public Property Get TimerOn() As Boolean
    TimerOn = m_TimerOn
End Property

Public Property Let TimerOn(ByVal Value As Boolean)
    m_TimerOn = Value
End Property

Public Property Get Closing() As Boolean
    Closing = m_Closing
End Property

Public Property Let Closing(ByVal Value As Boolean)
    m_Closing = Value

    If m_Closing Then
        m_LoginOK = True
        m_SelectInbox = True
        m_CurrentlyFetching = False
        m_FetchFinished = True
        m_SearchAll = True
        m_FetchStarted = True
        m_CheckingAddresses = False
        m_SetFlagsStart = False
        m_SetFlagsFinished = True
        m_LogoutSent = False
        m_DoSendLogout = False
    End If

End Property

Public Property Get SetFlagsFinished() As Boolean
    SetFlagsFinished = m_SetFlagsFinished
End Property

Public Property Let SetFlagsFinished(ByVal Value As Boolean)
    m_SetFlagsFinished = Value

    If m_SetFlagsFinished Then
        m_LoginOK = True
        m_SelectInbox = True
        m_CurrentlyFetching = False
        m_FetchFinished = True
        m_SearchAll = True
        m_FetchStarted = True
        m_CheckingAddresses = False
        m_SetFlagsStart = False
        m_Closing = False
        m_LogoutSent = False
        m_DoSendLogout = False
    End If
    
End Property

Public Property Get SetFlagsStart() As Boolean
    SetFlagsStart = m_SetFlagsStart
End Property

Public Property Let SetFlagsStart(ByVal Value As Boolean)
    m_SetFlagsStart = Value

    If m_SetFlagsStart Then
        m_LoginOK = True
        m_SelectInbox = True
        m_CurrentlyFetching = False
        m_FetchFinished = True
        m_SearchAll = True
        m_FetchStarted = True
        m_CheckingAddresses = False
        m_SetFlagsFinished = False
        m_Closing = False
        m_LogoutSent = False
        m_DoSendLogout = False
    End If
    
End Property

Public Property Get FetchFinished() As Boolean
    FetchFinished = m_FetchFinished
End Property

Public Property Let FetchFinished(ByVal Value As Boolean)
    m_FetchFinished = Value
    
    If m_FetchFinished Then
        m_LoginOK = True
        m_SelectInbox = True
        m_CurrentlyFetching = False
        m_CheckingAddresses = False
        m_SearchAll = True
        m_FetchStarted = True
        m_SetFlagsStart = False
        m_SetFlagsFinished = False
        m_Closing = False
        m_LogoutSent = False
        m_DoSendLogout = False
    End If
    
End Property

Public Property Get CurrentlyFetching() As Boolean
    CurrentlyFetching = m_CurrentlyFetching
End Property

Public Property Let CurrentlyFetching(ByVal Value As Boolean)
    m_CurrentlyFetching = Value
    
    If m_CurrentlyFetching Then
        m_LoginOK = True
        m_SelectInbox = True
        m_FetchStarted = True
        m_CheckingAddresses = False
        m_SearchAll = True
        m_FetchFinished = False
        m_SetFlagsStart = False
        m_SetFlagsFinished = False
        m_Closing = False
        m_LogoutSent = False
        m_DoSendLogout = False
    End If
    
End Property

Public Property Get CheckingAddresses() As Boolean
    CheckingAddresses = m_CheckingAddresses
End Property

Public Property Let CheckingAddresses(ByVal Value As Boolean)
    m_CheckingAddresses = Value

    If m_CheckingAddresses Then
        m_LoginOK = True
        m_SelectInbox = True
        m_CurrentlyFetching = False
        m_FetchFinished = True
        m_SearchAll = True
        m_FetchStarted = True
        m_SetFlagsStart = False
        m_SetFlagsFinished = False
        m_Closing = False
        m_LogoutSent = False
        m_DoSendLogout = False
    End If
    
End Property

Public Property Get DoSendLogout() As Boolean
    DoSendLogout = m_DoSendLogout
End Property

Public Property Let DoSendLogout(ByVal Value As Boolean)
    m_DoSendLogout = Value

    If m_DoSendLogout Then
        m_LoginOK = True
        m_SelectInbox = True
        m_CurrentlyFetching = False
        m_FetchFinished = True
        m_SearchAll = True
        m_FetchStarted = True
        m_CheckingAddresses = False
        m_SetFlagsStart = False
        m_SetFlagsFinished = True
        m_LogoutSent = False
        m_Closing = False
    End If

End Property

Public Property Get LogoutSent() As Boolean
    LogoutSent = m_LogoutSent
End Property

Public Property Let LogoutSent(ByVal Value As Boolean)
    m_LogoutSent = Value

    If m_LogoutSent Then
        m_LoginOK = True
        m_SelectInbox = True
        m_CurrentlyFetching = False
        m_FetchFinished = True
        m_SearchAll = True
        m_FetchStarted = True
        m_CheckingAddresses = False
        m_SetFlagsStart = False
        m_SetFlagsFinished = True
        m_DoSendLogout = True
        m_Closing = False
    End If

End Property

Public Property Get FetchStarted() As Boolean
    FetchStarted = m_FetchStarted
End Property

Public Property Let FetchStarted(ByVal Value As Boolean)
    m_FetchStarted = Value

    If m_FetchStarted Then
        m_LoginOK = True
        m_SelectInbox = True
        m_CurrentlyFetching = True
        m_CheckingAddresses = False
        m_SearchAll = True
        m_FetchFinished = False
        m_SetFlagsStart = False
        m_SetFlagsFinished = False
        m_Closing = False
        m_LogoutSent = False
        m_DoSendLogout = False
    End If
    
End Property

Public Property Get SearchAll() As Boolean
    SearchAll = m_SearchAll
End Property

Public Property Let SearchAll(ByVal Value As Boolean)
    m_SearchAll = Value
    
    If m_SearchAll Then
        m_LoginOK = True
        m_SelectInbox = True
        m_FetchStarted = False
        m_CheckingAddresses = False
        m_CurrentlyFetching = False
        m_FetchFinished = False
        m_SetFlagsStart = False
        m_SetFlagsFinished = False
        m_Closing = False
        m_LogoutSent = False
        m_DoSendLogout = False
    End If
    
End Property

Public Property Get SelectInbox() As Boolean
    SelectInbox = m_SelectInbox
End Property

Public Property Let SelectInbox(ByVal Value As Boolean)
    m_SelectInbox = Value
    
    If m_SelectInbox Then
        m_LoginOK = True
        m_SearchAll = False
        m_FetchStarted = False
        m_CheckingAddresses = False
        m_CurrentlyFetching = False
        m_FetchFinished = False
        m_SetFlagsStart = False
        m_SetFlagsFinished = False
        m_Closing = False
        m_LogoutSent = False
        m_DoSendLogout = False
    End If
    
End Property

Public Property Get LoginOK() As Boolean
    LoginOK = m_LoginOK
End Property

Public Property Let LoginOK(ByVal Value As Boolean)
    m_LoginOK = Value
    
    If m_LoginOK Then
        m_SelectInbox = False
        m_SearchAll = False
        m_FetchStarted = False
        m_CheckingAddresses = False
        m_CurrentlyFetching = False
        m_FetchFinished = False
        m_SetFlagsStart = False
        m_SetFlagsFinished = False
        m_Closing = False
        m_LogoutSent = False
        m_DoSendLogout = False
    End If
    
End Property

Private Sub Class_Initialize()

    Call SetAllToFalse

End Sub

Public Sub SetAllToFalse()

    m_LoginOK = False
    m_SelectInbox = False
    m_SearchAll = False
    m_FetchStarted = False
    m_CheckingAddresses = False
    m_CurrentlyFetching = False
    m_FetchFinished = False
    m_SetFlagsStart = False
    m_SetFlagsFinished = False
    m_Closing = False
    m_LogoutSent = False
    m_DoSendLogout = False

End Sub
