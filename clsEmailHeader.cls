VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEmailHeader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_From As String
Private m_EmailUniqueID As String
Private m_BodyText As String
Private m_Delete As Boolean
Private m_ChangeFlag As Boolean
Private m_Email_Number As Integer
Private m_Subject As String
Private m_FromAddress As String
Private m_Suffix As String
Private m_Domain As String
Private m_Virtual As String
Private m_HasVirtual As Boolean
Private m_SentDate As String
Private m_Size As String
Private m_OK As Boolean

Public Property Get OK() As Boolean
    OK = m_OK
End Property

Public Property Let OK(ByVal Value As Boolean)
    m_OK = Value
End Property

Public Property Get Size() As String
    Size = m_Size
End Property

Public Property Let Size(ByVal Value As String)
    m_Size = Value
End Property

Public Property Get SentDate() As String
    SentDate = m_SentDate
End Property

Public Property Let SentDate(ByVal Value As String)
    m_SentDate = Value
End Property

Public Property Get HasVirtual() As Boolean
    HasVirtual = m_HasVirtual
End Property

Public Property Let HasVirtual(ByVal Value As Boolean)
    m_HasVirtual = Value
End Property

Public Property Get Virtual() As String
    Virtual = m_Virtual
End Property

Public Property Let Virtual(ByVal Value As String)
    m_Virtual = Value
End Property

Public Property Get Domain() As String
    Domain = m_Domain
End Property

Public Property Let Domain(ByVal Value As String)
    m_Domain = Value
End Property

Public Property Get Suffix() As String
    Suffix = m_Suffix
End Property

Public Property Let Suffix(ByVal Value As String)
    m_Suffix = Value
End Property


Public Property Get FromAddress() As String
    FromAddress = m_FromAddress
End Property

Public Property Let FromAddress(ByVal Value As String)
    m_FromAddress = Value
End Property

Public Property Get Subject() As String
    Subject = m_Subject
End Property

Public Property Let Subject(ByVal Value As String)
    m_Subject = Value
End Property

Public Property Get Email_Number() As Integer
    Email_Number = m_Email_Number
End Property

Public Property Let Email_Number(ByVal Value As Integer)
    m_Email_Number = Value
End Property

Public Property Get ChangeFlag() As Boolean
    ChangeFlag = m_ChangeFlag
End Property

Public Property Let ChangeFlag(ByVal Value As Boolean)
    m_ChangeFlag = Value
End Property

Public Property Get Delete() As Boolean
    Delete = m_Delete
End Property

Public Property Let Delete(ByVal Value As Boolean)
    m_Delete = Value
End Property

Public Property Get EmailUniqueID() As String
    EmailUniqueID = m_EmailUniqueID
End Property

Public Property Let EmailUniqueID(ByVal Value As String)
    m_EmailUniqueID = Value
End Property

Public Property Get BodyText() As String
    BodyText = m_BodyText
End Property

Public Property Let BodyText(ByVal Value As String)
    m_BodyText = Value
End Property

Public Property Get From() As String
    From = m_From
End Property

Public Property Let From(ByVal Value As String)
    m_From = Value
End Property
