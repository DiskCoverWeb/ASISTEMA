VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FCyberPCs 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "RESIDENTE"
   ClientHeight    =   2055
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3810
   ControlBox      =   0   'False
   Icon            =   "CyberPCs.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "CyberPCs.frx":0ECA
   ScaleHeight     =   2055
   ScaleWidth      =   3810
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   105
      Picture         =   "CyberPCs.frx":21DF
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   3
      Top             =   105
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Timer Timer1 
      Left            =   105
      Top             =   630
   End
   Begin MSAdodcLib.Adodc AdoRespaldo 
      Height          =   330
      Left            =   420
      Top             =   2310
      Visible         =   0   'False
      Width           =   2010
      _ExtentX        =   3545
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Respaldo"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Desactiva"
      Height          =   330
      Left            =   2835
      TabIndex        =   1
      Top             =   1050
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000FF&
      Caption         =   "Minimizar"
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   1680
      Width           =   3795
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FFC0FF&
      Height          =   330
      Left            =   0
      TabIndex        =   4
      Top             =   1365
      Width           =   3795
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   1170
      Left            =   1680
      TabIndex        =   2
      Top             =   105
      Width           =   2115
   End
End
Attribute VB_Name = "FCyberPCs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Tic As NOTIFYICONDATA

Public Sub CreateIcon()
    Tic.cbSize = Len(Tic)
    Tic.hwnd = Picture1.hwnd
    Tic.uID = 1&
    Tic.uFlags = NIF_DOALL
    Tic.uCallbackMessage = WM_MOUSEMOVE
    Tic.hIcon = Picture1.Picture
    Tic.szTip = "Control del Cyber Equipo No. " & Format(PC_Numero, "00") & " " & Chr$(0)
    Erg = Shell_NotifyIcon(NIM_ADD, Tic)
End Sub
 
Public Sub DeleteIcon()
Dim Tic As NOTIFYICONDATA
    Tic.cbSize = Len(Tic)
    Tic.hwnd = Picture1.hwnd
    Tic.uID = 1&
    Erg = Shell_NotifyIcon(NIM_DELETE, Tic)
End Sub

Private Sub Command1_Click()
   CreateIcon
   FCyberPCs.Visible = False
End Sub

Private Sub Command2_Click()
  DeleteIcon
End Sub

Private Sub Form_Load()
Dim AdoCon1 As ADODB.Connection
Dim RstSchema As ADODB.Recordset
Dim I As Integer
    TiempoSistema = Time
    Timer1.Enabled = True
    Timer1.Interval = 1000  '1/2 segundo
    
   'obtenemos la ruta del servidor en base a la carpeta de windows del pece actual
    RutaGeneraFile = WindowsDirectory & "\PC_Numero.key"
    NumFile = FreeFile
    Cadena = ""
    Open RutaGeneraFile For Input As #NumFile
      Do While Not EOF(NumFile)
         Cadena = Cadena & Input(1, #NumFile) ' Obtiene un carácter.
      Loop
    Close #NumFile
    Unidad = Trim(Mid(Cadena, 1, 2))
    PC_Numero = Val(Trim(Mid(Cadena, 4, 3)))
    NumEmpresa = Trim(Mid(Cadena, 7, 3))
    RatonReloj
    
    RutaDestino = Unidad & "\SISTEMA"
    RutaSistema = Unidad & "\SISTEMA"
    RutaEmpresa = RutaSistema & "\EMPRESA"
    RutaEmpresaOld = RutaSistema & "\EMPRESA"
    RutaSysBases = Unidad & "\SYSBASES"
    RutaUpdate = RutaDestino
    ChDir RutaDestino
   'Determinar que tipo de bases utilizamos
    Evaluar = False
    SQL_Server = True
    Cadena = Dir(RutaSistema & "\EMPRESA\", vbNormal) 'Recupera la primera entrada.
    Do While Cadena <> ""
       If Cadena <> "." And Cadena <> ".." Then
          If (GetAttr(RutaSistema & "\EMPRESA\" & Cadena) And vbNormal) = vbNormal Then
             If UCase(Cadena) = "DISKCOVE.MDB" Then SQL_Server = False
          End If
       End If
       Cadena = Dir
    Loop
   'Buscamos la cadena de conección a la base
    If SQL_Server Then
       RutaGeneraFile = RutaSistema & "\SERVER.TXT"
    Else
       RutaGeneraFile = RutaSistema & "\ACCESS.TXT"
    End If
    NumFile = FreeFile
    AdoStrCnn = ""
    Open RutaGeneraFile For Input As #NumFile
      Do While Not EOF(NumFile)
         AdoStrCnn = AdoStrCnn & Input(1, #NumFile) ' Obtiene un carácter.
      Loop
    Close #NumFile
   'Verificamos si la base esta en Microsoft Access o en SQL Server 7.0
    If SQL_Server Then
       PathEmpresa = ""
    Else
       PathEmpresa = UCase(RutaEmpresa & "\DISKCOVE.MDB")
       AdoStrCnn = AdoStrCnn & "Data Source=" & PathEmpresa
    End If
    FechaSistema = Format(Date, FormatoFechas)
   'MsgBox Weekday(FechaSistema)
    
    CodigoUsuario = "ACCESO01"
    NombreUsuario = "Supervisor General"
    Empresa = "MODULO DE ACTUALIZACION DE BASES Y DATOS"
    Periodo_Contable = "."
    
    'MsgBox RutaGeneraFile & vbCrLf & vbCrLf & AdoStrCnn
    
   'Sacamos las tablas del sistema
    ConectarAdodc AdoRespaldo
    For I = 0 To 32700: Next I
    MsgBox "COMPUTADOR ACTIVADO PARA" & vbCrLf _
         & "UTILIZAR EL INTERNET"
   'Obtenemos los valores de cobros por minutos
    Cantidad_Cyber_Tiempo = 0
    sSQL = "SELECT * " _
         & "FROM Catalogo_Cyber_Tiempo " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "ORDER BY Desde,Hasta "
    SelectAdodc AdoRespaldo, sSQL
    With AdoRespaldo.Recordset
     If .RecordCount > 0 Then
         Cantidad_Cyber_Tiempo = .RecordCount - 1
         ReDim VCyber_Tiempo(Cantidad_Cyber_Tiempo)
         For I = 0 To Cantidad_Cyber_Tiempo
             VCyber_Tiempo(I).Desde = 0
             VCyber_Tiempo(I).Hasta = 0
             VCyber_Tiempo(I).Valor = 0
         Next I
         I = 0
         Do While Not .EOF
            VCyber_Tiempo(I).Desde = .Fields("Desde")
            VCyber_Tiempo(I).Hasta = .Fields("Hasta")
            VCyber_Tiempo(I).Valor = .Fields("Valor")
            I = I + 1
           .MoveNext
         Loop
     End If
    End With
    Cadena = ""
    For I = 0 To Cantidad_Cyber_Tiempo
        Cadena = Cadena _
               & VCyber_Tiempo(I).Desde & "-" & VCyber_Tiempo(I).Hasta & " (" & VCyber_Tiempo(I).Valor & ")" & vbCrLf
    Next I
'    MsgBox Cadena
    IntervaloTiempo = 2
    RatonNormal
    CreateIcon
    Label2.Caption = WindowsDirectory
    FCyberPCs.Visible = False
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  x = x / Screen.TwipsPerPixelX
  Select Case x
    Case WM_LBUTTONDOWN
         Caption = "Left Click"
    Case WM_RBUTTONDOWN
         Caption = "Right Click"
    Case WM_MOUSEMOVE
         Caption = "Move"
    Case WM_LBUTTONDBLCLK
         Caption = "Double Click"
         FCyberPCs.Visible = True
  End Select
End Sub

Private Sub Timer1_Timer()
Dim MiTiempoFin As Single
Dim MiTiempo As Single
   'MiTiempoFin = CDbl(CDate(Time))
   sSQL = "SELECT * " _
        & "FROM Catalogo_Cyber " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND Periodo = '" & Periodo_Contable & "' " _
        & "AND Codigo = '99.02." & Format(PC_Numero, "000") & "' "
   SelectAdodc AdoRespaldo, sSQL
   With AdoRespaldo.Recordset
    If .RecordCount > 0 Then
        MiTiempo = CSng(CDate(.Fields("Inicio")))
        MiTiempoFin = CSng(CDate(.Fields("Fin")))
        If .Fields("PC_Ocupaga") Then
            Calcular_Total_PC PC_Numero - 1, MiTiempo, MiTiempoFin
            Label1.Caption = "Equipo " & Format(PC_Numero, "00") & vbCrLf _
                           & "[" & Format(TiempoPCs(PC_Numero - 1), FormatoTimes) & "]" & vbCrLf _
                           & "USD " & TotalPCs(PC_Numero - 1)
        Else
            Label1.Caption = "EQUIPO LIBRE" & vbCrLf & "No. " & Format(PC_Numero, "00")
        End If
    End If
   End With
End Sub

