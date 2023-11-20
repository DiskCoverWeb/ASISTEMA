VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.ocx"
Begin VB.Form PruebaMail 
   BackColor       =   &H8000000A&
   Caption         =   "Form1"
   ClientHeight    =   4950
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5550
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4950
   ScaleWidth      =   5550
   Begin VB.TextBox txtSender 
      Height          =   330
      Left            =   105
      TabIndex        =   7
      Text            =   "informacion@diskcoversystem.com"
      Top             =   2625
      Width           =   4635
   End
   Begin VB.TextBox txtHost 
      Height          =   330
      Left            =   105
      TabIndex        =   6
      Text            =   "relay.dnsexit.com"
      Top             =   2205
      Width           =   4635
   End
   Begin VB.TextBox txtMessage 
      Height          =   330
      Left            =   105
      TabIndex        =   5
      Text            =   "prueba de mail"
      Top             =   1785
      Width           =   4635
   End
   Begin VB.TextBox txtSubject 
      Height          =   330
      Left            =   105
      TabIndex        =   4
      Text            =   "hola"
      Top             =   1365
      Width           =   4635
   End
   Begin VB.TextBox txtRecipient 
      Height          =   330
      Left            =   105
      TabIndex        =   3
      Text            =   "diskcoversystem@msn.com"
      Top             =   945
      Width           =   4635
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   4515
      Top             =   2415
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "cmdSend"
      Height          =   645
      Left            =   3255
      TabIndex        =   2
      Top             =   210
      Width           =   1485
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "cmdNew"
      Height          =   645
      Left            =   1680
      TabIndex        =   1
      Top             =   210
      Width           =   1485
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "cmdClose"
      Height          =   645
      Left            =   105
      TabIndex        =   0
      Top             =   210
      Width           =   1485
   End
End
Attribute VB_Name = "PruebaMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum SMTP_State
    MAIL_CONNECT
    MAIL_HELO
    MAIL_FROM
    MAIL_RCPTTO
    MAIL_DATA
    MAIL_DOT
    MAIL_QUIT
End Enum

Private m_State As SMTP_State
'

Private Sub cmdClose_Click()

    Unload Me

End Sub

Private Sub cmdNew_Click()

    txtRecipient = ""
    txtSubject = ""
    txtMessage = ""

End Sub

Private Sub cmdSend_Click()

    Winsock1.Connect Trim$(txtHost), 25
    m_State = MAIL_CONNECT

End Sub

Private Sub Form_Load()
    '
    'clear all textboxes
    '
    For Each ctl In Me.Controls
        If TypeOf ctl Is TextBox Then
            ctl.Text = ""
        End If
    Next
    '
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)

    Dim strServerResponse   As String
    Dim strResponseCode     As String
    Dim strDataToSend       As String
    '
    'Retrive data from winsock buffer
    '
    Winsock1.GetData strServerResponse
    '
    Debug.Print strServerResponse
    '
    'Get server response code (first three symbols)
    '
    strResponseCode = Left(strServerResponse, 3)
    '
    'Only these three codes tell us that previous
    'command accepted successfully and we can go on
    '
    If strResponseCode = "250" Or _
       strResponseCode = "220" Or _
       strResponseCode = "354" Then

        Select Case m_State
            Case MAIL_CONNECT
                'Change current state of the session
                m_State = MAIL_HELO
                '
                'Remove blank spaces
                strDataToSend = Trim$(txtSender)
                '
                'Retrieve mailbox name from e-mail address
                strDataToSend = Left$(strDataToSend, _
                                InStr(1, strDataToSend, "@") - 1)
                'Send HELO command to the server
                Winsock1.SendData "HELO " & strDataToSend & vbCrLf
                '
                Debug.Print "HELO " & strDataToSend
                '
            Case MAIL_HELO
                '
                'Change current state of the session
                m_State = MAIL_FROM
                '
                'Send MAIL FROM command to the server
                Winsock1.SendData "MAIL FROM:" & Trim$(txtSender) & vbCrLf
                '
                Debug.Print "MAIL FROM:" & Trim$(txtSender)
                '
            Case MAIL_FROM
                '
                'Change current state of the session
                m_State = MAIL_RCPTTO
                '
                'Send RCPT TO command to the server
                Winsock1.SendData "RCPT TO:" & Trim$(txtRecipient) & vbCrLf
                '
                Debug.Print "RCPT TO:" & Trim$(txtRecipient)
                '
            Case MAIL_RCPTTO
                '
                'Change current state of the session
                m_State = MAIL_DATA
                '
                'Send DATA command to the server
                Winsock1.SendData "DATA" & vbCrLf
                '
                Debug.Print "DATA"
                '
            Case MAIL_DATA
                '
                'Change current state of the session
                m_State = MAIL_DOT
                '
                'So now we are sending a message body
                'Each line of text must be completed with
                'linefeed symbol (Chr$(10) or vbLf) not with vbCrLf
                '
                'Send Subject line
                Winsock1.SendData "Subject:" & txtSubject & vbLf
                '
                Debug.Print "Subject:" & txtSubject
                '
                Dim varLines    As Variant
                Dim varLine     As Variant
                '
                'Parse message to get lines (for VB6 only)
                varLines = Split(txtMessage, vbCrLf)
                '
                'Send each line of the message
                For Each varLine In varLines
                    Winsock1.SendData CStr(varLine) & vbLf
                    '
                    Debug.Print CStr(varLine)
                Next
                '
                'Send a dot symbol to inform server
                'that sending of message comleted
                Winsock1.SendData "." & vbCrLf
                '
                Debug.Print "."
                '
            Case MAIL_DOT
                'Change current state of the session
                m_State = MAIL_QUIT
                '
                'Send QUIT command to the server
                Winsock1.SendData "QUIT" & vbCrLf
                '
                Debug.Print "QUIT"
            Case MAIL_QUIT
                '
                'Close connection
                Winsock1.Close
                '
        End Select

    Else
        '
        'If we are here server replied with
        'unacceptable respose code therefore we need
        'close connection and inform user about problem
        '
        Winsock1.Close
        '
        If Not m_State = MAIL_QUIT Then
            MsgBox "SMTP Error: " & strServerResponse, _
                    vbInformation, "SMTP Error"
        Else
            MsgBox "Message sent successfuly.", vbInformation
        End If
        '
    End If

End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

    MsgBox "Winsock Error number " & Number & vbCrLf & _
            Description, vbExclamation, "Winsock Error"

End Sub
