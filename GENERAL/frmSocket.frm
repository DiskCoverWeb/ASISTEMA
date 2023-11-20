VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form FSocket 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Socket"
   ClientHeight    =   7560
   ClientLeft      =   555
   ClientTop       =   855
   ClientWidth     =   9540
   Icon            =   "frmSocket.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7560
   ScaleWidth      =   9540
   Begin VB.CommandButton cmdEscucha 
      Caption         =   "Escucha..."
      Height          =   375
      Left            =   5520
      TabIndex        =   4
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmdEnvia 
      Caption         =   "&Envia Paquete"
      Height          =   375
      Left            =   5520
      TabIndex        =   7
      Top             =   2760
      Width           =   1215
   End
   Begin MSWinsockLib.Winsock TCPSocket 
      Left            =   5565
      Top             =   630
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer tmrInterval 
      Left            =   6090
      Top             =   630
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   5520
      TabIndex        =   8
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdDesconectar 
      Caption         =   "&Desconecta..."
      Height          =   375
      Left            =   5520
      TabIndex        =   6
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton cmdConectar 
      Caption         =   "&Conecta..."
      Height          =   375
      Left            =   5520
      TabIndex        =   5
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Frame fraTraza 
      Caption         =   "Traza"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   105
      TabIndex        =   18
      Top             =   3255
      Width           =   6615
      Begin VB.TextBox txtTraza 
         BackColor       =   &H00E0E0E0&
         Height          =   3135
         Left            =   105
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   210
         Width           =   6375
      End
   End
   Begin VB.Frame fraLocal 
      Caption         =   "Datos de conexión local"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   735
      TabIndex        =   15
      Top             =   105
      Width           =   2295
      Begin VB.TextBox txtLocalIP 
         BackColor       =   &H00E0E0E0&
         Height          =   375
         Left            =   720
         TabIndex        =   20
         TabStop         =   0   'False
         Text            =   "255.255.255.255"
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox txtLocalPort 
         Height          =   375
         Left            =   720
         TabIndex        =   3
         Text            =   "1001"
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label lblInfo 
         Caption         =   "IP"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   495
      End
      Begin VB.Label lblInfo 
         Caption         =   "Puerto"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   16
         Top             =   840
         Width           =   495
      End
   End
   Begin VB.Frame fraConexion 
      Caption         =   "Otros datos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   210
      TabIndex        =   12
      Top             =   1680
      Width           =   4695
      Begin VB.TextBox txtPaquetesGet 
         BackColor       =   &H00E0E0E0&
         Height          =   375
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   22
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox txtInterval 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   1920
         TabIndex        =   9
         Text            =   "1"
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox txtPaquetesSend 
         BackColor       =   &H00E0E0E0&
         Height          =   375
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   360
         Width           =   495
      End
      Begin VB.Label lblInfo 
         Caption         =   "Paquetes recibidos"
         Height          =   375
         Index           =   0
         Left            =   2415
         TabIndex        =   23
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label lblInfo 
         Caption         =   "(Comprobar estado cada             segundos)"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   21
         Top             =   960
         Width           =   3135
      End
      Begin VB.Label lblInfo 
         Caption         =   "Paquetes enviados"
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame fraHost 
      Caption         =   "Datos del Host remoto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   3120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
      Begin VB.TextBox txtRemotePort 
         Height          =   375
         Left            =   720
         TabIndex        =   2
         Text            =   "1001"
         Top             =   840
         Width           =   1455
      End
      Begin VB.TextBox txtRemoteHost 
         Height          =   375
         Left            =   720
         TabIndex        =   1
         Text            =   "255.255.255.255"
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label lblInfo 
         Caption         =   "Puerto"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   11
         Top             =   840
         Width           =   495
      End
      Begin VB.Label lblInfo 
         Caption         =   "IP"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.Image imgConexion 
      Height          =   480
      Left            =   120
      Picture         =   "frmSocket.frx":0442
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "FSocket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Propiedades por defecto del Host remoto
Private Const c_strDefaultRemoteHostIP As String = "192.168.0.2"
Private Const c_strDefaultRemotePort As String = "1002"
'Propiedades por defecto de la máquina local
Private Const c_strDefaultLocalPort As String = "1001"
'Propiedades por defecto
Private Const c_lngDefaultInterval As Long = 1 'Segundos
Private Const c_strMSG As String = "SOCKET_HOLA"

Private Sub cmdConectar_Click()
    
    With Me.TCPSocket
        'Sólo puede conectar si el socket está cerrado o escuchando...
        If .State = 0 Or .State = 2 Then
            'Si está escuchando hay que cerrar el socket primero
            If .State = 2 Then
                .Close
                Do While .State <> 0
                    DoEvents    'Obligamos a que "suelte" el puerto...
                Loop
            End If
            'Muestra traza
            EscribeEnTraza "<Conecta>"
            'Apariencia form
            Me.txtPaquetesGet.Text = "0"
            Me.txtPaquetesSend.Text = "0"
            'Establece las propiedades del socket y conecta
            .RemoteHost = Me.TxtRemoteHost.Text
            .RemotePort = Me.txtRemotePort.Text
            .LocalPort = Me.txtLocalPort.Text
            .Connect
        End If
    End With

End Sub

Private Sub cmdDesconectar_Click()
    'Si el socket no está cerrado, lo cierra
    If Me.TCPSocket.State <> 0 Then
        'Muestra traza
        EscribeEnTraza "<Desconecta>"
        Me.TCPSocket.Close
    End If
End Sub

Private Sub cmdEnvia_Click()
    'Manda el MSG por defecto si la conexión está abierta
    With Me.TCPSocket
        If .State = "7" Then
            EscribeEnTraza "<Mandando paquete...>"
            .SendData c_strMSG
        End If
    End With
End Sub

Private Sub cmdEscucha_Click()
    
    'Si el socket está cerrado, se pone en escucha
    With Me.TCPSocket
        If .State = "0" Then
            EscribeEnTraza "<Escucha>"
            'Apariencia form
            Me.txtPaquetesGet.Text = "0"
            Me.txtPaquetesSend.Text = "0"
            .LocalPort = Me.txtLocalPort.Text
            .Listen
        End If
    End With
    
End Sub

Private Sub cmdSalir_Click()
    'Muestra traza
    EscribeEnTraza "<FIN>"
    'Si el socket no está cerrado, lo cierra
    If Me.TCPSocket.State <> 0 Then Me.TCPSocket.Close
    'Finaliza
    Unload Me
    'End
End Sub

Private Sub Form_Activate()
    RatonNormal
End Sub

Private Sub Form_Load()
    'Muestra traza
    EscribeEnTraza "<INICIO>"
    'Muestra las propiedades por defecto
    Me.TxtRemoteHost.Text = c_strDefaultRemoteHostIP
    Me.txtRemotePort.Text = c_strDefaultRemotePort
    Me.txtLocalIP.Text = Me.TCPSocket.LocalIP   'Lee del socket
    Me.txtLocalPort.Text = c_strDefaultLocalPort
    Me.txtInterval.Text = c_lngDefaultInterval
    'Establece el intervalo en milisegundos
    Me.tmrInterval.Interval = c_lngDefaultInterval * 1000
    'Escucha por defecto
    cmdEscucha_Click
End Sub

Private Sub TCPSocket_Close()
    'Muestra traza
    EscribeEnTraza "<El Host ha cerrado el socket>"
End Sub

Private Sub TCPSocket_Connect()
    'Muestra traza
    EscribeEnTraza "<Se ha conectado al Host>"
End Sub

Private Sub TCPSocket_ConnectionRequest(ByVal requestID As Long)
    'Si el socket está cerrado acepta la petición
    With Me.TCPSocket
        If .State = "2" Then
            EscribeEnTraza "<Acepta la petición " & requestID & ">"
            .Close
            .Accept requestID
            'Actualiza form
            Me.TxtRemoteHost.Text = .RemoteHost
            Me.txtRemotePort.Text = .RemotePort
        End If
    End With
End Sub

Private Sub TCPSocket_DataArrival(ByVal bytesTotal As Long)
    
    Dim strData As String
    'Obtiene el MSG entrante
    TCPSocket.GetData strData
    If strData = c_strMSG Then
        EscribeEnTraza "<Se ha recibido un MSG correctamente>"
        Me.txtPaquetesGet.Text = CLng(Me.txtPaquetesGet.Text) + 1
     Else
        EscribeEnTraza "<Se ha recibido un MSG no válido:" & _
            vbNewLine & "MSG: " & strData & ">"
    End If
    
End Sub

Private Sub TCPSocket_Error(ByVal Number As Integer, _
                            Description As String, _
                            ByVal Scode As Long, _
                            ByVal Source As String, _
                            ByVal HelpFile As String, _
                            ByVal HelpContext As Long, _
                            CancelDisplay As Boolean)
                            
    'Cancela la muestra del msg por defecto
    CancelDisplay = True
    
    'Muestra traza
    EscribeEnTraza "<Error en el socket:" & vbNewLine & _
        "   * Número: " & Number & vbNewLine & _
        "   * Descripción: " & Description & vbNewLine & _
        "   * Código: " & Scode & vbNewLine & _
        "   * Origen: " & Source & ">" & vbNewLine
    
    'Desconecta
    cmdDesconectar_Click
    
End Sub

Private Sub TCPSocket_SendComplete()
    EscribeEnTraza "<Envío correcto>"
    Me.txtPaquetesSend.Text = CLng(Me.txtPaquetesSend.Text) + 1
End Sub

Private Sub TCPSocket_SendProgress(ByVal bytesSent As Long, _
                                    ByVal bytesRemaining As Long)
    EscribeEnTraza "<Enviando: quedan " & bytesRemaining & " bytes>"
End Sub

Private Sub tmrInterval_Timer()
    
    Static intEstadoAnterior As Integer
    
    'Muestra el estado en pantalla, si ha cambiado desde la última vez.
    If intEstadoAnterior <> Me.TCPSocket.State Then
        intEstadoAnterior = Me.TCPSocket.State
        EscribeEnTraza
    End If
    
End Sub

Private Function DimeEstado() As String
    Select Case Me.TCPSocket.State
    Case 0
        DimeEstado = "Cerrado"
    Case 1
        DimeEstado = "Abierto"
    Case 2
        DimeEstado = "Escuchando"
    Case 3
        DimeEstado = "Conexión pendiente..."
    Case 4
        DimeEstado = "Resolviendo host..."
    Case 5
        DimeEstado = "Host resuelto"
    Case 6
        DimeEstado = "Conectando..."
    Case 7
        DimeEstado = "Conectado"
    Case 8
        DimeEstado = "Cerrando..."
    Case 9
        DimeEstado = "ERROR !!"
    Case Else
        DimeEstado = "Desconocido (" & Me.TCPSocket.State & ")"
    End Select
End Function

Private Sub EscribeEnTraza(Optional ByVal strMSG As String)
    Dim strTraza As String
    With Me.txtTraza
        strTraza = Format(Now, "hh:nn:ss") & _
            " --> (Estado del socket: " & DimeEstado & ")" & vbNewLine
        If strMSG <> "" Then strTraza = strTraza & "   " & strMSG & vbNewLine
        .Text = .Text & strTraza
        .SelStart = Len(.Text) - 1
    End With
End Sub

Private Sub txtInterval_Change()
    'Establece el intervalo en milisegundos
    Me.tmrInterval.Interval = CLng(Me.txtInterval.Text) * 1000
End Sub

