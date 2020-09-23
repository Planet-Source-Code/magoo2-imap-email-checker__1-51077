VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmViewMail 
   Caption         =   "IMAP Email Checker - ViewMail"
   ClientHeight    =   6495
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11250
   Icon            =   "frmViewMail.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6495
   ScaleWidth      =   11250
   StartUpPosition =   2  'CenterScreen
   Begin RichTextLib.RichTextBox RTB1 
      Height          =   6255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   11033
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmViewMail.frx":0442
   End
   Begin MSWinsockLib.Winsock socket2 
      Left            =   360
      Top             =   6120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   9090
   End
End
Attribute VB_Name = "frmViewMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_EmailNumber As Integer
Private m_MailServer As String
Private m_UserName As String
Private m_Password As String

Const STANDARDMAILPORT As Integer = 143
Dim intSendPacketNumber As Integer
Dim intReturnPacketNumber As Integer
Dim strSPNS As String
Dim strRPNS As String

Dim bolFetchingBody As Boolean
Dim bolFetchingIncremented As Boolean
Dim bolDoneOnce As Boolean

Dim cStatus As New clsStatus


Public Property Let Password(ByVal Value As String)
    m_Password = Value
End Property

Public Property Let UserName(ByVal Value As String)
    m_UserName = Value
End Property

Public Property Let MailServer(ByVal Value As String)
    m_MailServer = Value
End Property

Public Property Let EmailNumber(ByVal Value As Integer)
    m_EmailNumber = Value
End Property

Private Sub Form_Activate()

    If Not bolDoneOnce Then
        
        bolDoneOnce = True
    
    End If
        
End Sub

Private Sub Form_Load()

    
    intSendPacketNumber = 1
    intReturnPacketNumber = 0
    bolDoneOnce = False

    bolFetchingBody = False
    bolFetchingIncremented = False
    DoEvents
    DoEvents
    DoEvents
    
    ' Connect To Mail Server
    socket2.Close
    DoEvents
    DoEvents
    socket2.LocalPort = 0
    
    socket2.Connect m_MailServer, STANDARDMAILPORT
    DoEvents
    DoEvents

End Sub

Private Sub Form_Resize()

    On Error Resume Next
    
    RTB1.Move 120, 120, (Me.ScaleWidth - 240), (Me.ScaleHeight - 240)

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set cStatus = Nothing
End Sub

Private Sub socket2_DataArrival(ByVal bytesTotal As Long)

    Dim strPacket As String
    
    socket2.GetData strPacket
    
    If (Not bolFetchingBody) Or (Not bolFetchingIncremented) Then
        strRPNS = IncrementRPNS
        If bolFetchingBody Then
            bolFetchingIncremented = True
        End If
    End If

    If bolFetchingBody Then
        Call PresentBody(strPacket)
    Else
        Call ProcessNewData(strPacket)
    End If

End Sub


Private Sub ProcessNewData(strData As String)

    On Error GoTo BadSub

    Dim strSendString As String
    Dim strTemp As String
    Dim i As Integer
    Dim strEmailsNumString As String
    Dim count2 As Integer
    Dim pos As Integer
    Dim strFromSubject As String
    
    
    If cStatus.DoSendLogout Then
        GoTo DoSendLogout
    End If
    
    Select Case intSendPacketNumber
    Case 1: 'Check if we have connected successfully
            
            If InStr(strData, "* OK") = 0 Then
                GoTo BadSub
            Else
                strSendString = " LOGIN " & m_UserName & " " & m_Password
                SendDataToServer strSendString
            End If
            
    Case Else
            
            strTemp = strRPNS & " OK"
            
            If Not cStatus.LoginOK Then
            
                If InStr(strData, strTemp) = 0 Then
                    MsgBox "Login Error"
                    cStatus.DoSendLogout = True
                    GoTo DoSendLogout
                Else
                    cStatus.LoginOK = True
                    strTemp = " SELECT INBOX"
                    SendDataToServer strTemp
                    GoTo DOEP
                End If
                
            End If
            
            If cStatus.LoginOK And Not cStatus.SelectInbox Then
            
                If InStr(strData, strTemp) = 0 Then
                    MsgBox "Select Error"
                    cStatus.DoSendLogout = True
                    GoTo DoSendLogout
                Else
                    cStatus.SelectInbox = True
                    Call FetchSingleEmailData
                    GoTo DOEP
                End If
                
            End If
                
            If cStatus.Closing And Not cStatus.DoSendLogout Then
            
                   cStatus.DoSendLogout = True
                     strTemp = " CLOSE"
                    SendDataToServer strTemp
                     GoTo DOEP
               
               If InStr(strData, strTemp) = 0 Then
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
            
                socket2.Close
            
                If InStr(strData, strTemp) = 0 Then
                    MsgBox "LogoutSent Error"
                End If
                
                GoTo DOEP
                
            End If

            If cStatus.DoSendLogout And Not cStatus.LogoutSent Then
            
                cStatus.LogoutSent = True
                strSendString = " LOGOUT"
                SendDataToServer strSendString
                
            End If

    End Select

DOEP:

Exit Sub

BadSub:
    
    Debug.Print Error

End Sub


Private Function IncrementSPNS() As String
     intSendPacketNumber = intSendPacketNumber + 1
    IncrementSPNS = "A" & intSendPacketNumber
End Function

Private Function IncrementRPNS() As String
    intReturnPacketNumber = intReturnPacketNumber + 1
    IncrementRPNS = "A" & intReturnPacketNumber
End Function


Private Sub FetchSingleEmailData()
    Dim strTemp As String
    
    bolFetchingBody = True
    
    strTemp = " Fetch " & m_EmailNumber & " (RFC822.TEXT)"
    SendDataToServer strTemp

End Sub

Private Sub SendDataToServer(strDataOut As String)
    
    Dim strTemp As String
    
    strSPNS = IncrementSPNS
    strTemp = strSPNS & strDataOut & vbCrLf
    socket2.SendData strTemp

End Sub

Private Sub PresentBody(strData As String)
    
    Debug.Print strData
    Dim strTemp As String
    Dim i As Integer
    
    strTemp = strRPNS & " OK"
    i = InStr(1, strData, strTemp, vbTextCompare)
    
    If i > 0 Then
        If i > 4 Then
            RTB1.Text = RTB1.Text & Left(strData, (i - 4))
        End If
        bolFetchingBody = False
        cStatus.LogoutSent = True
        SendDataToServer " LOGOUT"
    Else
        
        If InStr(1, strData, "RFC822.TEXT", vbTextCompare) = 0 Then
            RTB1.Text = RTB1.Text & strData
        End If
    
    End If
    
End Sub
