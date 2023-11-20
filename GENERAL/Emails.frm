VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form FEnviarMails 
   ClientHeight    =   6195
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   7590
   LinkTopic       =   "Form1"
   ScaleHeight     =   6195
   ScaleWidth      =   7590
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Txt_Estado 
      Height          =   2745
      Left            =   105
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   14
      Text            =   "Emails.frx":0000
      Top             =   3360
      Width           =   7365
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&SALIR"
      Height          =   435
      Left            =   6090
      TabIndex        =   13
      Top             =   1155
      Width           =   1380
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   540
      Left            =   105
      TabIndex        =   12
      Top             =   2205
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   953
      _Version        =   327682
      Appearance      =   1
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   105
      Top             =   2835
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   435
      Left            =   6090
      TabIndex        =   11
      Top             =   630
      Width           =   1380
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   435
      Left            =   6090
      TabIndex        =   10
      Top             =   105
      Width           =   1380
   End
   Begin VB.TextBox txt_Mensaje 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1485
      Left            =   1575
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Text            =   "Emails.frx":0006
      Top             =   1785
      Width           =   4425
   End
   Begin VB.TextBox txt_Asunto 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1575
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   1365
      Width           =   4425
   End
   Begin VB.TextBox txt_MailTo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1575
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   945
      Width           =   4425
   End
   Begin VB.TextBox txt_emailFrom 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1575
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   525
      Width           =   4425
   End
   Begin VB.TextBox txt_Server_Smtp 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1575
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   105
      Width           =   4425
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mensaje"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   105
      TabIndex        =   8
      Top             =   1785
      Width           =   1485
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Asunto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   105
      TabIndex        =   6
      Top             =   1365
      Width           =   1485
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mail Destino"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   105
      TabIndex        =   4
      Top             =   945
      Width           =   1485
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mail Origen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   105
      TabIndex        =   2
      Top             =   525
      Width           =   1485
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Server SMTP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   1485
   End
End
Attribute VB_Name = "FEnviarMails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private m_Sleep   As Boolean
Private Estado As String

' Botón que envía el e-mail
Private Sub Command1_Click()
   Txt_Estado = ""
   Winsock1.RemoteHost = txt_Server_Smtp
   Winsock1.RemotePort = 25
   Winsock1.Connect
End Sub

' Botón que Finaliza el socket abierto
Private Sub Command2_Click()
    Winsock1.Close
End Sub

Private Sub Command3_Click()
 End
End Sub

  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim LineaMsg As Integer
    '   this is the main processing code for
    '   sending an email message
    '   the iState variable maintains the current
    '   state of the protocol exchange so that we
    '   know what to send next
    Dim strData As String
    Static iState       As Integer
    Dim iMsgNum         As Integer
    Dim szMsg           As String
    Dim I               As Integer
    
    Winsock1.GetData strData, vbString
    
    iMsgNum = Val(Left(strData, InStr(strData, " ")))
    
    Select Case iMsgNum
        Case 220    '   initial message
            Winsock1.SendData "HELO " & txt_Server_Smtp & vbCrLf
            Txt_Estado = Txt_Estado & "Mail Server is ready..." & vbCrLf
            iState = 1
        Case 221
            If iState = 999 Then
                Txt_Estado = Txt_Estado & "Disconnected from mail server after error..." & vbCrLf
            Else
                Txt_Estado = Txt_Estado & "Disconnected from mail server..." & vbCrLf
            End If
            iState = 0
            
        Case 250
            Select Case iState
                Case 1:
                    Winsock1.SendData "MAIL FROM:<" & txt_emailFrom & ">" & vbCrLf
                    'Debug.Print "MAIL FROM:<" & txt_nameFrom & ">" & vbCrLf
                    Txt_Estado = Txt_Estado & "Sending FROM command..." & vbCrLf
                    iState = 2
                Case 2:
                    Winsock1.SendData "RCPT TO:<" & txt_MailTo & ">" & vbCrLf
                    'Debug.Print "RCPT TO:<" & txt_MailTo & ">" & vbCrLf
                    Txt_Estado = Txt_Estado & "Sending RCPT command..." & vbCrLf
                    iState = 3
                    
                Case 3:
                    Winsock1.SendData "DATA" & vbCrLf
                    'Debug.Print "DATA" & vbCrLf
                    Txt_Estado = Txt_Estado & "Sending DATA command..." & vbCrLf
                    iState = 4
                    
                Case 5:
                    Winsock1.SendData "QUIT" & vbCrLf
                    'Debug.Print "QUIT" & vbCrLf
                    Txt_Estado = Txt_Estado & "Sending Quit command to disconnecting from mail server..." & vbCrLf
                    iState = 6
                    Winsock1.Close
                End Select
                
        Case 354
            LineaMsg = 1
            iState = 5
            szMsg = txt_Mensaje & " " & vbCrLf
            Txt_Estado = Txt_Estado & "Sending mail message data..." & vbCrLf
            Winsock1.SendData "Subject: " & txt_Asunto & vbCrLf
            While szMsg <> ""
                Winsock1.SendData Left(szMsg, InStr(szMsg, Chr(10)))
                'Debug.Print "Sending:" & Left(szMsg, InStr(szMsg, Chr(10)))
                Txt_Estado = Txt_Estado & vbTab & LineaMsg & ": " & Left(szMsg, InStr(szMsg, Chr(10)))
                szMsg = Mid(szMsg, InStr(szMsg, Chr(10)) + 1)
                LineaMsg = LineaMsg + 1
            Wend
            Winsock1.SendData "." & vbCrLf
            Txt_Estado = Txt_Estado & "Fin del mensaje " & vbCrLf
        Case 500 To 599
            Winsock1.SendData "QUIT" & vbCrLf
            Txt_Estado = Txt_Estado & "Error sending mail..." & vbCrLf
            'Debug.Print "Error sending mail... quitting..."
            iState = 999
            Winsock1.Close
        End Select
    End Sub

Private Sub Form_Load()
Dim ctl As Control
'''    For Each ctl In Me.Controls
'''        If TypeOf ctl Is TextBox Then
'''            ctl.Text = ""
'''        End If
'''    Next
    Command1.Caption = " Enviar "
    Command2.Caption = " Desconectar "
    txt_Server_Smtp = "mail.diskcoversystem.com"
    txt_emailFrom = "anuncios@diskcoversystem.com"
    txt_MailTo = "diskcover@msn.com"
    txt_Asunto = "Hola"
    txt_Mensaje = "Este es un mensaje de prueba" & vbCrLf _
                & "De varias lineas " & vbCrLf _
                & "en la que se presenta " & vbCrLf _
                & "un resumen de datos" & vbCrLf _
                & "importantes de la contabilidad" & vbCrLf
    Me.Caption = "Envío de email con el control Winsock "
End Sub

