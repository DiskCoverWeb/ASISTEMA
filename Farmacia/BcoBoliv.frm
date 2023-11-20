VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FBancoBolivariano 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "BANCO BOLIVARIANO"
   ClientHeight    =   7035
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11280
   Icon            =   "BcoBoliv.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7035
   ScaleWidth      =   11280
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox CheqPend 
      BackColor       =   &H00FF8080&
      Caption         =   "Sin Deuda Pendiente"
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
      Left            =   6825
      TabIndex        =   4
      Top             =   945
      Width           =   2220
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Generar Codigos de Nomina"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   9450
      Picture         =   "BcoBoliv.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4725
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   9450
      Picture         =   "BcoBoliv.frx":0C7C
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5565
      Width           =   1695
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Recibir Abonos del Banco"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9450
      Picture         =   "BcoBoliv.frx":1672
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1890
      Width           =   1695
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H00C0C0FF&
      Caption         =   "&Enviar Rubros Facturar en Cero"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9450
      Picture         =   "BcoBoliv.frx":1F18
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   3780
      Width           =   1695
   End
   Begin VB.CommandButton Command12 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Enviar Nomina al Banco en cero"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9450
      Picture         =   "BcoBoliv.frx":2722
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   2835
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Enviar Rubros Facturar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9450
      Picture         =   "BcoBoliv.frx":2F2C
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   945
      Width           =   1695
   End
   Begin VB.TextBox TxtFile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   5055
      Left            =   2940
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   15
      Top             =   1365
      Width           =   6420
   End
   Begin VB.FileListBox File1 
      BackColor       =   &H00FFC0C0&
      Height          =   2235
      Left            =   105
      TabIndex        =   13
      Top             =   4200
      Width           =   2850
   End
   Begin ComctlLib.ProgressBar ProgBarra 
      Height          =   330
      Left            =   105
      TabIndex        =   12
      Top             =   6510
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   582
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.DriveListBox Drive1 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   105
      TabIndex        =   9
      Top             =   1575
      Width           =   2850
   End
   Begin VB.DirListBox Dir1 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2115
      Left            =   105
      TabIndex        =   10
      Top             =   1890
      Width           =   2850
   End
   Begin MSAdodcLib.Adodc AdoAux 
      Height          =   330
      Left            =   210
      Top             =   2520
      Visible         =   0   'False
      Width           =   2115
      _ExtentX        =   3731
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
      Caption         =   "Aux"
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
   Begin MSAdodcLib.Adodc AdoQuery 
      Height          =   330
      Left            =   210
      Top             =   2835
      Visible         =   0   'False
      Width           =   2115
      _ExtentX        =   3731
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
      Caption         =   "Query"
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
   Begin MSAdodcLib.Adodc AdoAct 
      Height          =   330
      Left            =   210
      Top             =   3150
      Visible         =   0   'False
      Width           =   2115
      _ExtentX        =   3731
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
      Caption         =   "Act"
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
   Begin MSMask.MaskEdBox MBFechaI 
      Height          =   330
      Left            =   2205
      TabIndex        =   1
      Top             =   945
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   16761024
      AllowPrompt     =   -1  'True
      AutoTab         =   -1  'True
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "0"
   End
   Begin MSAdodcLib.Adodc AdoGrupo 
      Height          =   330
      Left            =   210
      Top             =   3465
      Visible         =   0   'False
      Width           =   2115
      _ExtentX        =   3731
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
      Caption         =   "Grupo"
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
   Begin MSMask.MaskEdBox MBFechaF 
      Height          =   330
      Left            =   5460
      TabIndex        =   3
      Top             =   945
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   16761024
      AllowPrompt     =   -1  'True
      AutoTab         =   -1  'True
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "0"
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Fecha Tope de &Pago"
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
      Left            =   3465
      TabIndex        =   2
      Top             =   945
      Width           =   2010
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " ARCHIVO:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   105
      TabIndex        =   11
      Top             =   3990
      Width           =   2850
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &Fecha de Facturacion"
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
      Top             =   945
      Width           =   2115
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &ORIGEN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   105
      TabIndex        =   8
      Top             =   1365
      Width           =   2850
   End
End
Attribute VB_Name = "FBancoBolivariano"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim AdoStrCnnOld As String
Dim AdoStrCnn1 As String
Dim NumFile As Integer
Dim RutaGeneraFile As String
Dim XAdoStrCnn As String
Dim IJ As Long
Dim ModuloResp As String
Dim RetVal
Dim RutaBackupXX As String

Public Sub TipoProcesos(Opciones As String)
  NombreArchivo = ""
  Select Case Opciones
    Case "ALUMNOS"
         Dir1.Path = RutaBackup & "\" & Opciones & "\"
         File1.FileName = Dir1.Path & "\*.*"
    Case "DESCUENT"
         Dir1.Path = RutaBackup & "\" & Opciones & "\"
         File1.FileName = Dir1.Path & "\*.*"
    Case "FACTURAS"
         Dir1.Path = RutaBackup & "\" & Opciones & "\"
         File1.FileName = Dir1.Path & "\*.*"
    Case "NOMINA"
         Dir1.Path = RutaBackup & "\" & Opciones & "\"
         File1.FileName = Dir1.Path & "\*.*"
    Case Else
         If RutaBackupXX <> "" Then Dir1.Path = RutaBackupXX & "\" Else Dir1.Path = RutaBackup & "\"
         File1.FileName = Dir1.Path & "\*.*"
  End Select
  Dir1.Refresh
End Sub

Private Sub Command10_Click()
Dim Cont As Integer
'MsgBox NombreFile
TipoProcesos "DESCUENT"
TipoDoc = ""
If OpcP.Value Then TipoDoc = "0" Else TipoDoc = "1"
FechaValida MBFechaI
MiMes = Format(FechaMes(MBFechaI.Text), "00")
MiFecha = Format(MBFechaI.Text, "MM/dd/yyyy")
RutaGeneraFile = UCase(Dir1.Path & "\Items" & CodigoDelBanco & ".TXT")
NumFile = FreeFile
Cont = 0
Contador = 0: Total = 0: Abono = 0
ProgBarra.Value = 0
ProgBarra.Min = 0
ProgBarra.Max = 100
'MsgBox RutaGeneraFile
Open RutaGeneraFile For Output As #NumFile ' Abre el archivo.
Close #NumFile
ProgBarra.Value = ProgBarra.Max
RatonNormal
MsgBox "Fin del Proceso " & vbCrLf & "El Archivo se Generara en: " & vbCrLf & RutaGeneraFile
End Sub

Private Sub Command12_Click()
TipoProcesos "NOMINA"
If OpcP.Value Then TipoDoc = "0" Else TipoDoc = "1"
FechaValida MBFechaI
FechaValida MBFechaF
MiMes = Format(FechaMes(MBFechaI), "00")
MiFecha = Format(MBFechaI, "MM/dd/yyyy")
FechaTexto = Format(MBFechaF, "MM/dd/yyyy")
RutaGeneraFile = UCase(Dir1.Path & "\ALUMNOS" & CodigoDelBanco & ".TXT")

sSQL = "SELECT CI_RUC,SUM(Saldo_MN) As Saldo_Pend " _
     & "FROM Facturas As F,Clientes As C " _
     & "WHERE F.Item = '" & NumEmpresa & "' " _
     & "AND F.Periodo = '" & Periodo_Contable & "' " _
     & "AND F.Fecha <= #" & BuscarFecha(MBFechaF) & "# " _
     & "AND F.TC NOT IN ('C','P') " _
     & "AND F.CodigoC = C.Codigo " _
     & "AND F.T <> 'A' " _
     & "GROUP BY CodigoC,CI_RUC " _
     & "HAVING SUM(Saldo_MN) > 0 " _
     & "ORDER BY CI_RUC "
SelectAdodc AdoAct, sSQL

sSQL = "SELECT C.Grupo,C.Codigo,C.Cliente,C.Direccion,C.CI_RUC,C.Casilla,CF.Total_MN " _
     & "FROM Clientes As C,Facturas As CF " _
     & "WHERE C.FA <> " & Val(adFalse) & " " _
     & "AND CF.Fecha = #" & BuscarFecha(MBFechaI) & "# " _
     & "AND CF.Item = '" & NumEmpresa & "' " _
     & "AND CF.Periodo = '" & Periodo_Contable & "' " _
     & "AND CF.TC NOT IN ('C','P') " _
     & "AND CF.T <> 'A' " _
     & "AND C.Codigo = CF.CodigoC " _
     & "ORDER BY C.Grupo,C.CI_RUC,C.Cliente "
SelectAdodc AdoAux, sSQL
NumFile = FreeFile
Contador = 0
ProgBarra.Value = 0
ProgBarra.Min = 0
'MsgBox RutaGeneraFile
Open RutaGeneraFile For Output As #NumFile ' Abre el archivo.
With AdoAux.Recordset
 If .RecordCount > 0 Then
     Print #NumFile, "999";
     Print #NumFile, CodigoDelBanco;
     Print #NumFile, TipoDoc;
     Print #NumFile, "    ";
     Print #NumFile, MiFecha
    .MoveFirst
     ProgBarra.Max = .RecordCount
     Do While Not .EOF
        FBancoBolivariano.Caption = .Fields("Grupo") & " - " & Format(Contador / .RecordCount, "00%")
        SaldoPendiente = 0
        Total_Factura = 0
        Monto_Total = 0
        Total = 0
        ProgBarra.Value = Contador
        CodigoCli = .Fields("CI_RUC")
        Codigo = "0"
        For I = 1 To Len(.Fields("CI_RUC"))
            If IsNumeric(Mid(.Fields("CI_RUC"), I, 1)) Then Codigo = Codigo & Mid(.Fields("CI_RUC"), I, 1)
        Next I
        Codigo = Trim(Str(Val(Codigo)))
        Codigo = Codigo & String(8 - Len(Codigo), " ")
      ' MsgBox "|" & Codigo & "|"
        NombreCliente = SetearBlancos(Mid(.Fields("Cliente"), 1, 30), 30, 0, False)
        Codigo1 = Trim(Mid(SinEspaciosIzq(.Fields("Direccion")), 1, 15))
        Codigo3 = Trim(Mid(SinEspaciosDer(.Fields("Direccion")), 1, 3))
        Codigo2 = Trim(Mid(.Fields("Direccion"), Len(Codigo1) + 1, Len(.Fields("Direccion"))))
        Codigo4 = Mid(.Fields("Casilla"), 1, 10)
        Saldo_ME = 0: Total_Desc = 0: SaldoPendiente = 0
        If AdoAct.Recordset.RecordCount > 0 Then
           AdoAct.Recordset.MoveFirst
           AdoAct.Recordset.Find ("CI_RUC Like '" & CodigoCli & "' ")
           If Not AdoAct.Recordset.EOF Then SaldoPendiente = AdoAct.Recordset.Fields("Saldo_Pend")
        End If
        If CheqPend.Value = 1 Then SaldoPendiente = .Fields("Total_MN")
        Total_Factura = .Fields("Total_MN")
        Monto_Total = Total_Factura
        Total = SaldoPendiente
        If Codigo1 = "" Then Codigo1 = Ninguno
        If Codigo2 = "" Then Codigo2 = Ninguno
        If Codigo3 = "" Then Codigo3 = Ninguno
        Codigo2 = Trim(Mid(Codigo2, 1, Len(Codigo2) - Len(SinEspaciosDer(Codigo2))))
        Codigo1 = SetearBlancos(Codigo1, 15, 0, False)
        Codigo2 = SetearBlancos(Codigo2, 15, 0, False)
        Codigo3 = SetearBlancos(Codigo3, 3, 0, False)
        Codigo4 = SetearBlancos(Codigo4, 10, 0, False)
        If Trim(Codigo4) = Ninguno Then Codigo4 = String(10, " ")
      ' Total = Total - Monto_Total
        If Total < 0 Then Total = 0
      ' Empieza la trama por Alumno
        'MsgBox NombreCliente & vbCrLf & Total
        Print #NumFile, CodigoDelBanco;                        ' Colegio/Institucion
        Print #NumFile, Codigo;                                ' Codigo Alumno
        Print #NumFile, MiFecha;                               ' Fecha Pen: FechaTexto = FechaTexto1
        Print #NumFile, TipoDoc & "  ";                        ' Proceso
        Print #NumFile, Format(Total, "00000000.00"); ' Valor
        Print #NumFile, FechaTexto;                            ' Fecha Cobis
        Print #NumFile, "01/01/1900";                          ' Fecha Pago "01/01/1900";
        Print #NumFile, "N";                                   ' Estado = N
        Print #NumFile, NombreCliente;                         ' Nombre Alumno
        Print #NumFile, Codigo2;                               ' Nombre del Curso
        Print #NumFile, Codigo3;                               ' Nombre del Paralelo
        Print #NumFile, Codigo1;                               ' Nombre de la Seccion
        Print #NumFile, Format(Monto_Total, "00000000.00");    ' Valor Mes
        Print #NumFile, Codigo4;                               ' Pago por Deposito de Cuenta
        Print #NumFile, "1";                                   ' Moneda = 1
        Print #NumFile, Format(Total, "00000000.00"); ' Valor 2
        Print #NumFile, Format(Total, "00000000.00")  ' Valor 1
        Contador = Contador + 1
       .MoveNext
     Loop
 End If
End With
Close #NumFile
ProgBarra.Value = ProgBarra.Max
RatonNormal
MsgBox "Fin del Proceso " & vbCrLf & "El Archivo se Generara en: " & vbCrLf & RutaGeneraFile
End Sub

Private Sub Command2_Click()
  Unload FBancoBolivariano
End Sub

Private Sub Command4_Click()
Dim Cont As Integer
'MsgBox NombreFile
TipoProcesos "DESCUENT"
TipoDoc = ""
If OpcP.Value Then TipoDoc = "0" Else TipoDoc = "1"
FechaValida MBFechaI
MiMes = Format(FechaMes(MBFechaI.Text), "00")
MiFecha = Format(MBFechaI.Text, "MM/dd/yyyy")
RutaGeneraFile = UCase(Dir1.Path & "\Items" & CodigoDelBanco & ".TXT")

sSQL = "SELECT CodigoC,CI_RUC,SUM(Saldo_MN) As Saldo_Pend " _
     & "FROM Facturas As F,Clientes As C " _
     & "WHERE F.Item = '" & NumEmpresa & "' " _
     & "AND NOT F.TC IN ('C','P') " _
     & "AND F.Periodo = '" & Periodo_Contable & "' " _
     & "AND F.CodigoC = C.Codigo " _
     & "AND F.T = 'P' " _
     & "GROUP BY CodigoC,CI_RUC " _
     & "HAVING SUM(Saldo_MN) > 0 " _
     & "ORDER BY CodigoC,CI_RUC "
SelectAdodc AdoAct, sSQL

sSQL = "SELECT CF.Codigo,CF.Valor,CF.Codigo_Inv,C.Grupo,C.CI_RUC,CP.Item_Banco,CP.Desc_Item " _
     & "FROM Clientes_Facturacion As CF, Clientes As C,Catalogo_Productos As CP " _
     & "WHERE CF.Item = '" & NumEmpresa & "' " _
     & "AND CP.Periodo = '" & Periodo_Contable & "' " _
     & "AND CF.Codigo = C.Codigo " _
     & "AND CF.Codigo_Inv = CP.Codigo_Inv " _
     & "AND CF.Item = CP.Item " _
     & "AND CF.Periodo = CP.Periodo " _
     & "ORDER BY C.CI_RUC,C.Grupo,CF.Codigo,CP.Item_Banco,CP.Desc_Item "
SelectAdodc AdoAux, sSQL
NumFile = FreeFile
Cont = 0
Contador = 0: Total = 0: Abono = 0
ProgBarra.Value = 0
ProgBarra.Min = 0
'MsgBox RutaGeneraFile
Open RutaGeneraFile For Output As #NumFile ' Abre el archivo.
With AdoAux.Recordset
 If .RecordCount > 0 Then
     Print #NumFile, "999";
     Print #NumFile, CodigoDelBanco;
     Print #NumFile, TipoDoc
     ProgBarra.Max = .RecordCount
     CodigoCli = .Fields("CI_RUC")
     Codigo = "0"
     For I = 1 To Len(.Fields("CI_RUC"))
         If IsNumeric(Mid(.Fields("CI_RUC"), I, 1)) Then Codigo = Codigo & Mid(.Fields("CI_RUC"), I, 1)
     Next I
     Codigo = Trim(Str(Val(Codigo)))
     Codigo = Codigo & String(8 - Len(Codigo), " ")
     Do While Not .EOF
        If CodigoCli <> .Fields("CI_RUC") Then
           Cont = Cont + 1
           Print #NumFile, CodigoDelBanco;
           Print #NumFile, Codigo;
           Print #NumFile, "930";
           Print #NumFile, SetearBlancos("SUBTOTAL", 20, 0, False);
           Print #NumFile, Format(Total, "00000000.00");
           Print #NumFile, "  " & TipoDoc;
           Print #NumFile, SetearBlancos(Cont, 3, 0, False)
           If Abono > 0 Then
              'MsgBox Total & vbCrLf & Abono
              Cont = Cont + 1
              Print #NumFile, CodigoDelBanco;
              Print #NumFile, Codigo;
              Print #NumFile, "950";
              Print #NumFile, SetearBlancos("DESCUENTO", 20, 0, False);
              Print #NumFile, "-" & Format(Abono, "0000000.00");
              Print #NumFile, "  " & TipoDoc;
              Print #NumFile, SetearBlancos(Cont, 3, 0, False)
              Total = Total - Abono
           End If
           Cont = Cont + 1
           Print #NumFile, CodigoDelBanco;
           Print #NumFile, Codigo;
           Print #NumFile, "960";
           Print #NumFile, SetearBlancos("IVA   0%", 20, 0, False);
           Print #NumFile, Format(0, "00000000.00");
           Print #NumFile, "  " & TipoDoc;
           Print #NumFile, SetearBlancos(Cont, 3, 0, False)
           Cont = Cont + 1
           Print #NumFile, CodigoDelBanco;
           Print #NumFile, Codigo;
           Print #NumFile, "970";
           Print #NumFile, SetearBlancos("IVA 12%", 20, 0, False);
           Print #NumFile, Format(0, "00000000.00");
           Print #NumFile, "  " & TipoDoc;
           Print #NumFile, SetearBlancos(Cont, 3, 0, False)
           Cont = Cont + 1
           Print #NumFile, CodigoDelBanco;
           Print #NumFile, Codigo;
           Print #NumFile, "980";
           Print #NumFile, SetearBlancos("TOTAL", 20, 0, False);
           Print #NumFile, Format(Total, "00000000.00");
           Print #NumFile, "  " & TipoDoc;
           Print #NumFile, SetearBlancos(Cont, 3, 0, False)
           If CheqPend.Value = 1 Then
              Saldo = 0
           Else
           If AdoAct.Recordset.RecordCount > 0 Then
              AdoAct.Recordset.MoveFirst
              AdoAct.Recordset.Find ("CI_RUC Like '" & CodigoCli & "' ")
              If Not AdoAct.Recordset.EOF Then
                 'If Codigo = "1212    " Then MsgBox "______"
                 Saldo = AdoAct.Recordset.Fields("Saldo_Pend")
                 Total = Total + Saldo
                 Cont = Cont + 1
                 Print #NumFile, CodigoDelBanco;
                 Print #NumFile, Codigo;
                 Print #NumFile, "012";
                 Print #NumFile, SetearBlancos("DEUDA PENDIENTE", 20, 0, False);
                 Print #NumFile, Format(Saldo, "00000000.00");
                 Print #NumFile, "  " & TipoDoc;
                 Print #NumFile, SetearBlancos(Cont, 3, 0, False)
              End If
           End If
           End If
           Cont = Cont + 1
           Print #NumFile, CodigoDelBanco;
           Print #NumFile, Codigo;
           Print #NumFile, "999";
           Print #NumFile, SetearBlancos("VALOR A PAGAR", 20, 0, False);
           Print #NumFile, Format(Total, "00000000.00");
           Print #NumFile, "  " & TipoDoc;
           Print #NumFile, SetearBlancos(Cont, 3, 0, False)
           Cont = Cont + 1
           CodigoCli = .Fields("CI_RUC")
           NivelNo = .Fields("Grupo")
           Codigo = "0"
           For I = 1 To Len(.Fields("CI_RUC"))
               If IsNumeric(Mid(.Fields("CI_RUC"), I, 1)) Then Codigo = Codigo & Mid(.Fields("CI_RUC"), I, 1)
           Next I
           Cont = 0: Total = 0: Abono = 0
           Codigo = Trim(Str(Val(Codigo)))
           If (8 - Len(Codigo)) >= 0 Then Codigo = Codigo & String(8 - Len(Codigo), " ")
        End If
        NivelNo = .Fields("Grupo")
        CodigoInv = .Fields("Codigo_Inv")
        If .Fields("Codigo_Inv") = "01.88" Then
            Abono = Abono + .Fields("Valor")
        Else
            Cont = Cont + 1
            Contador = Contador + 1
            ProgBarra.Value = Contador
            Print #NumFile, CodigoDelBanco;
            Print #NumFile, Codigo;
            Print #NumFile, SetearBlancos(.Fields("Item_Banco"), 3, 0, False);
            Print #NumFile, SetearBlancos(.Fields("Desc_Item"), 20, 0, False);
            Print #NumFile, Format(.Fields("Valor"), "00000000.00");
            Print #NumFile, "  " & TipoDoc;
            Print #NumFile, SetearBlancos(Cont, 3, 0, False)
            Total = Total + .Fields("Valor")
        End If
       .MoveNext
     Loop
           Cont = Cont + 1
           Print #NumFile, CodigoDelBanco;
           Print #NumFile, Codigo;
           Print #NumFile, "930";
           Print #NumFile, SetearBlancos("SUBTOTAL", 20, 0, False);
           Print #NumFile, Format(Total, "00000000.00");
           Print #NumFile, "  " & TipoDoc;
           Print #NumFile, SetearBlancos(Cont, 3, 0, False)
           If Abono > 0 Then
              'MsgBox Total & vbCrLf & Abono
              Cont = Cont + 1
              Print #NumFile, CodigoDelBanco;
              Print #NumFile, Codigo;
              Print #NumFile, "950";
              Print #NumFile, SetearBlancos("DESCUENTO", 20, 0, False);
              Print #NumFile, "-" & Format(Abono, "0000000.00");
              Print #NumFile, "  " & TipoDoc;
              Print #NumFile, SetearBlancos(Cont, 3, 0, False)
              Total = Total - Abono
           End If
           Cont = Cont + 1
           Print #NumFile, CodigoDelBanco;
           Print #NumFile, Codigo;
           Print #NumFile, "960";
           Print #NumFile, SetearBlancos("IVA   0%", 20, 0, False);
           Print #NumFile, Format(0, "00000000.00");
           Print #NumFile, "  " & TipoDoc;
           Print #NumFile, SetearBlancos(Cont, 3, 0, False)
           Cont = Cont + 1
           Print #NumFile, CodigoDelBanco;
           Print #NumFile, Codigo;
           Print #NumFile, "970";
           Print #NumFile, SetearBlancos("IVA 12%", 20, 0, False);
           Print #NumFile, Format(0, "00000000.00");
           Print #NumFile, "  " & TipoDoc;
           Print #NumFile, SetearBlancos(Cont, 3, 0, False)
           Cont = Cont + 1
           Print #NumFile, CodigoDelBanco;
           Print #NumFile, Codigo;
           Print #NumFile, "980";
           Print #NumFile, SetearBlancos("TOTAL", 20, 0, False);
           Print #NumFile, Format(Total, "00000000.00");
           Print #NumFile, "  " & TipoDoc;
           Print #NumFile, SetearBlancos(Cont, 3, 0, False)
           If CheqPend.Value = 1 Then
              Saldo = 0
           Else
           If AdoAct.Recordset.RecordCount > 0 Then
              AdoAct.Recordset.MoveFirst
              AdoAct.Recordset.Find ("CI_RUC Like '" & CodigoCli & "' ")
              If Not AdoAct.Recordset.EOF Then
                 Saldo = AdoAct.Recordset.Fields("Saldo_Pend")
                 Total = Total + Saldo
                 Cont = Cont + 1
                 Print #NumFile, CodigoDelBanco;
                 Print #NumFile, Codigo;
                 Print #NumFile, "012";
                 Print #NumFile, SetearBlancos("DEUDA PENDIENTE", 20, 0, False);
                 Print #NumFile, Format(Saldo, "00000000.00");
                 Print #NumFile, "  " & TipoDoc;
                 Print #NumFile, SetearBlancos(Cont, 3, 0, False)
              End If
           End If
           End If
           Cont = Cont + 1
           Print #NumFile, CodigoDelBanco;
           Print #NumFile, Codigo;
           Print #NumFile, "999";
           Print #NumFile, SetearBlancos("VALOR A PAGAR", 20, 0, False);
           Print #NumFile, Format(Total, "00000000.00");
           Print #NumFile, "  " & TipoDoc;
           Print #NumFile, SetearBlancos(Cont, 3, 0, False)
 End If
End With
Close #NumFile
ProgBarra.Value = ProgBarra.Max
MsgBox "Fin del Proceso " & vbCrLf & "El Archivo se Generara en: " & vbCrLf & RutaGeneraFile
End Sub

'Recibir Abonos del Banco
'& "AND FA <> " & Val(adFalse) & " "
Private Sub Command5_Click()
Dim AuxNumEmp As String
Dim AT As Integer
  TotalAbonos1 = 0
  TextoImprimio = ""
  AuxNumEmp = NumEmpresa
  Cta_Cobrar = Ninguno
  RutaGeneraFile = UCase(Dir1.Path & "\" & NombreArchivo)
  RutaBackupXX = Dir1.Path
  sSQL = "SELECT Codigo,CI_RUC,Cliente " _
       & "FROM Clientes " _
       & "WHERE Codigo <> '.' " _
       & "ORDER BY Codigo "
  SelectAdodc AdoQuery, sSQL
  
  sSQL = "UPDATE Facturas " _
       & "SET X = '.' " _
       & "WHERE Periodo = '" & Periodo_Contable & "' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND TC NOT IN ('C','P') "
  ConectarAdoExecute sSQL
  Contador = 0
  FileResp = 0
  FechaValida MBFechaI
  FechaIni = BuscarFecha(MBFechaI.Text)
  FechaFin = BuscarFecha(MBFechaI.Text)
 'Actualizamos las facturas a modificar
  RatonReloj
  NumFile = FreeFile
  Open RutaGeneraFile For Input As #NumFile
       Line Input #NumFile, Cod_Field
       Do While Not EOF(NumFile)
          Line Input #NumFile, Cod_Field
          AT = Len(Cod_Field)
          Codigo = Format(Val(Mid(Cod_Field, AT - 18, 8)), "00000000")
          CodigoCli = Ninguno
          If AdoQuery.Recordset.RecordCount > 0 Then
             AdoQuery.Recordset.MoveFirst
             AdoQuery.Recordset.Find ("CI_RUC = '" & Codigo & "' ")
             If Not AdoQuery.Recordset.EOF Then CodigoCli = AdoQuery.Recordset.Fields("Codigo")
          End If
          sSQL = "UPDATE Facturas " _
               & "SET X = 'F' " _
               & "WHERE CodigoC = '" & CodigoCli & "' " _
               & "AND Periodo = '" & Periodo_Contable & "' " _
               & "AND Item = '" & NumEmpresa & "' " _
               & "AND TC NOT IN ('C','P') " _
               & "AND T <> 'A' "
          ConectarAdoExecute sSQL
          Contador = Contador + 1
          FBancoBolivariano.Caption = "Numero de Alumnos que han Abonado: " & Format(Contador, "0000000")
       Loop
  Close #NumFile
 'CxC Alumnos
  sSQL = "SELECT * " _
       & "FROM Catalogo_Lineas " _
       & "WHERE TL <> " & Val(adFalse) & " " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "ORDER BY Codigo "
  SelectAdodc AdoAux, sSQL
  If AdoAux.Recordset.RecordCount > 0 Then Cta_Cobrar = AdoAux.Recordset.Fields("CxC")
  FileResp = 0
  ProgBarra.Value = 0
  ProgBarra.Min = 0
  ProgBarra.Max = Contador
  Contador = 0
  NumFile = FreeFile
  Open RutaGeneraFile For Input As #NumFile
       Line Input #NumFile, Cod_Field
       If Len(Cod_Field) <> 31 Then
          MsgBox "EL ARCHIVO ESTA INCORRECTO," & vbCrLf _
               & "LA LONGITUD ES DE: " & Len(Cod_Field) & vbCrLf _
               & "CORRIJALO Y VUELVA A SUBIR"
          GoTo Salir_Modulo
       End If
       CodigoCorresp = Mid(Cod_Field, 1, 3)
       MiMes = Mid(Cod_Field, 4, 2)
       MiFecha = Mid(Cod_Field, 12, 2) & "/" & Mid(Cod_Field, 10, 2) & "/" & Mid(Cod_Field, 6, 4)
       Total = Val(CCur(Mid(Cod_Field, 14, 11) & "." & Mid(Cod_Field, 25, 2)))
       MBFechaI = MiFecha
       MBFechaF = MiFecha
       Contador = 0
       sSQL = "DELETE * " _
            & "FROM Trans_Abonos " _
            & "WHERE Fecha = #" & BuscarFecha(MiFecha) & "# " _
            & "AND Periodo = '" & Periodo_Contable & "' " _
            & "AND Item = '" & NumEmpresa & "' "
       ConectarAdoExecute sSQL
       
       Procesar_Saldo_De_Facturas FBancoBolivariano, AdoAux, True
       
       Do While Not EOF(NumFile)
          Line Input #NumFile, Cod_Field
          AT = Len(Cod_Field)
          Codigo = Format(Val(Mid(Cod_Field, AT - 18, 8)), "00000000")
          Total = Val(Mid(Cod_Field, AT - 10, 9) & "." & Mid(Cod_Field, AT - 1, 2))
          TotalAbonos1 = TotalAbonos1 + Total
          Total_ME = Total
          'MsgBox Codigo & vbCrLf & Total & vbCrLf & Mid(Cod_Field, AT - 10, 9) & "." & Mid(Cod_Field, AT - 1, 2)
          Factura_No = Val(Mid(Cod_Field, 26, 7))
          CodigoCli = Ninguno
          If AdoQuery.Recordset.RecordCount > 0 Then
             AdoQuery.Recordset.MoveFirst
             AdoQuery.Recordset.Find ("CI_RUC = '" & Codigo & "' ")
             If Not AdoQuery.Recordset.EOF Then
                CodigoCli = AdoQuery.Recordset.Fields("Codigo")
             End If
          End If
          'MsgBox Codigo & vbCrLf & CodigoCli
          sSQL = "SELECT * " _
               & "FROM Facturas " _
               & "WHERE CodigoC = '" & CodigoCli & "' " _
               & "AND Saldo_MN > 0 " _
               & "AND TC NOT IN ('C','P') " _
               & "AND Periodo = '" & Periodo_Contable & "' " _
               & "AND Item = '" & NumEmpresa & "' " _
               & "AND T <> 'A' " _
               & "ORDER BY Factura "
          SelectAdodc AdoAct, sSQL
          'MsgBox Codigo & vbCrLf & sSQL
          'If Codigo = "30000129" Then MsgBox "YA"
          If AdoAct.Recordset.RecordCount > 0 Then
             Abono = 1
             Do While Not AdoAct.Recordset.EOF And Abono > 0
                Abono = 0
                TipoFactura = AdoAct.Recordset.Fields("TC")
                Saldo = AdoAct.Recordset.Fields("Saldo_MN")
                SaldoTotal = AdoAct.Recordset.Fields("Saldo_MN")
                Factura_No = AdoAct.Recordset.Fields("Factura")
                'MsgBox Saldo & vbCrLf & Factura_No & vbCrLf & SaldoTotal
                If Total <= Saldo Then
                   If Total > 0 Then Abono = Total
                Else
                   Abono = Saldo
                End If
                'MsgBox CodigoCli & " - " & Total_ME & " - Fact = " & Factura_No & " - Abono = " & Abono
                If Abono > 0 Then
                   SetAdoAddNew "Trans_Abonos"
                   SetAdoFields "T", Cancelado
                   SetAdoFields "TP", TipoFactura
                   SetAdoFields "CodigoC", CodigoCli
                   SetAdoFields "Fecha", MiFecha
                   SetAdoFields "Comprobante", Ninguno
                   SetAdoFields "Factura", Factura_No
                   SetAdoFields "Abono", Abono
                   SetAdoFields "Banco", "EFECTIVO"
                   SetAdoFields "Cheque", Format(Total_ME, "##0.00")
                   SetAdoFields "Cta", Cta_CajaG
                   SetAdoFields "Cta_CxP", Cta_Cobrar
                   SetAdoUpdate
                   SaldoTotal = SaldoTotal - Abono
                   If SaldoTotal <= 0 Then
                      sSQL = "UPDATE Facturas " _
                           & "SET Saldo_MN = " & SaldoTotal & ", T = 'C'  " _
                           & "WHERE CodigoC = '" & CodigoCli & "' " _
                           & "AND Factura = " & Factura_No & " " _
                           & "AND TC = '" & TipoFactura & "' " _
                           & "AND Periodo = '" & Periodo_Contable & "' " _
                           & "AND Item = '" & NumEmpresa & "' "
                      ConectarAdoExecute sSQL
                   End If
                End If
                Total = Total - Abono
                AdoAct.Recordset.MoveNext
             Loop
          Else
             TextoImprimio = TextoImprimio & Codigo & vbCrLf
          End If
          Contador = Contador + 1
          If Contador > ProgBarra.Max Then Contador = ProgBarra.Max
          ProgBarra.Value = Contador
       Loop
       Procesar_Saldo_De_Facturas FBancoBolivariano, AdoAux, True
       ProgBarra.Value = ProgBarra.Max
Salir_Modulo:
  Close #NumFile
  FileResp = 0
  RatonNormal
  ProgBarra.Value = ProgBarra.Max
  NumEmpresa = AuxNumEmp
  FBancoBolivariano.Caption = "FACTURACION DE BANCOS"
  MsgBox "Total Abonado de " & ProgBarra.Max & " Facturas USD " & Format(TotalAbonos1, "#,##0.00") & vbCrLf & vbCrLf _
       & "Del dia: " & FechaStrgDias(MiFecha) & vbCrLf & vbCrLf _
       & "Fin del Proceso"
  Contador = 0
  If TextoImprimio <> "" Then FInfoError.Show
  Dir1.Path = RutaBackupX
End Sub

Private Sub Command6_Click()
TipoProcesos "ALUMNOS"
If OpcPension.Value Then TipoDoc = "0" Else TipoDoc = "1"
FechaValida MBFechaI
MiMes = Format(FechaMes(MBFechaI.Text), "00")
MiFecha = Format(MBFechaI.Text, "MM/dd/yyyy")

RutaGeneraFile = UCase(Dir1.Path & "\ALUMNOS" & CodigoDelBanco & ".TXT")

MsgBox "EL ARCHIVO SE GENERARA EN:" & vbCrLf & vbCrLf & RutaGeneraFile
sSQL = "SELECT Codigo,CI_RUC,Direccion,Cliente,Grupo " _
     & "FROM Clientes " _
     & "WHERE FA <> " & Val(adFalse) & " " _
     & "AND Cliente <> 'CONSUMIDOR FINAL' " _
     & "ORDER BY Grupo,Cliente,Direccion,CI_RUC "
SelectAdodc AdoAct, sSQL

NumFile = FreeFile
Contador = 0
ProgBarra.Value = 0
ProgBarra.Min = 0
Open RutaGeneraFile For Output As #NumFile ' Abre el archivo.
With AdoAct.Recordset
 If .RecordCount > 0 Then
    .MoveFirst
     ProgBarra.Max = .RecordCount
     Do While Not .EOF
        FBancoBolivariano.Caption = .Fields("Grupo") & " - " & Format(Contador / .RecordCount, "00%")
        ProgBarra.Value = Contador
        Codigo = .Fields("CI_RUC")
        NombreCliente = Trim(.Fields("Cliente"))
        Codigo1 = Trim(Mid(SinEspaciosIzq(.Fields("Direccion")), 1, 15))
        Codigo3 = Trim(Mid(SinEspaciosDer(.Fields("Direccion")), 1, 3))
        Codigo2 = Trim(Mid(.Fields("Direccion"), Len(Codigo1) + 1, Len(.Fields("Direccion"))))
        If Codigo1 = "" Then Codigo1 = Ninguno
        If Codigo2 = "" Then Codigo2 = Ninguno
        If Codigo3 = "" Then Codigo3 = Ninguno
      ' Empieza la trama por Alumno
        Print #NumFile, Codigo;                               ' Codigo Alumno
        Print #NumFile, ",";
        Print #NumFile, NombreCliente;                        ' Nombre Alumno
        Print #NumFile, ",";
        Print #NumFile, Codigo1;                              ' Nombre del Paralelo
        Print #NumFile, ",";
        Print #NumFile, Codigo3                               ' Nombre de la Seccion
        Contador = Contador + 1
       .MoveNext
     Loop
 End If
End With
Close #NumFile
RutaGeneraFile = UCase(Dir1.Path & "\ALUMNOS" & CodigoDelBanco & ".CSV")
NumFile = FreeFile
Contador = 0
ProgBarra.Value = 0
ProgBarra.Min = 0
Open RutaGeneraFile For Output As #NumFile ' Abre el archivo.
With AdoAct.Recordset
 If .RecordCount > 0 Then
    .MoveFirst
     ProgBarra.Max = .RecordCount
     Do While Not .EOF
        FBancoBolivariano.Caption = .Fields("Grupo") & " - " & Format(Contador / .RecordCount, "00%")
        ProgBarra.Value = Contador
        Codigo = .Fields("CI_RUC")
        NombreCliente = Trim(.Fields("Cliente"))
        Codigo1 = Trim(Mid(SinEspaciosIzq(.Fields("Direccion")), 1, 15))
        Codigo3 = Trim(Mid(SinEspaciosDer(.Fields("Direccion")), 1, 3))
        Codigo2 = Trim(Mid(.Fields("Direccion"), Len(Codigo1) + 1, Len(.Fields("Direccion"))))
        If Codigo1 = "" Then Codigo1 = Ninguno
        If Codigo2 = "" Then Codigo2 = Ninguno
        If Codigo3 = "" Then Codigo3 = Ninguno
      ' Empieza la trama por Alumno
        Print #NumFile, Codigo;                               ' Codigo Alumno
        Print #NumFile, ",";
        Print #NumFile, NombreCliente;                        ' Nombre Alumno
        Print #NumFile, ",";
        Print #NumFile, Codigo1;                              ' Nombre del Paralelo
        Print #NumFile, ",";
        Print #NumFile, Codigo3                               ' Nombre de la Seccion
        Contador = Contador + 1
       .MoveNext
     Loop
 End If
End With
Close #NumFile
ProgBarra.Value = ProgBarra.Max
End Sub

'''Private Sub Command7_Click()
'''TipoProcesos "NOMINA"
'''If OpcPension.Value Then TipoDoc = "0" Else TipoDoc = "1"
'''FechaValida MBFechaI
'''FechaValida MBFechaF
'''MiMes = Format(FechaMes(MBFechaI.Text), "00")
'''NoMeses = Month(MBFechaI)
'''MiFecha = Format(MBFechaI.Text, "MM/dd/yyyy")
'''FechaTexto = Format(MBFechaF.Text, "MM/dd/yyyy")
'''RutaGeneraFile = UCase(Dir1.Path & "\ALUMNOS" & CodigoDelBanco & ".TXT")
'''
'''sSQL = "SELECT CodigoC,CI_RUC,SUM(Saldo_MN) As Saldo_Pend " _
'''     & "FROM Facturas As F,Clientes As C " _
'''     & "WHERE F.Item = '" & NumEmpresa & "' " _
'''     & "AND NOT F.TC IN ('C','P') " _
'''     & "AND F.Periodo = '" & Periodo_Contable & "' " _
'''     & "AND F.CodigoC = C.Codigo " _
'''     & "AND F.T = 'P' " _
'''     & "GROUP BY CodigoC,CI_RUC " _
'''     & "HAVING SUM(Saldo_MN) > 0 " _
'''     & "ORDER BY CodigoC,CI_RUC "
'''SelectAdodc AdoAct, sSQL
'''
'''sSQL = "SELECT C.Grupo,C.Codigo,C.CI_RUC,SUM(CF.Valor) As Valor_Descuento " _
'''     & "FROM Clientes As C,Clientes_Facturacion As CF " _
'''     & "WHERE C.FA <> 0 " _
'''     & "AND C.Codigo = CF.Codigo " _
'''     & "AND CF.Item = '" & NumEmpresa & "' " _
'''     & "AND CF.Periodo = '" & Periodo_Contable & "' " _
'''     & "AND CF.Codigo_Inv = '01.88' " _
'''     & "GROUP BY C.Grupo,C.Codigo,C.CI_RUC " _
'''     & "ORDER BY C.Grupo,C.CI_RUC "
'''SelectAdodc AdoGrupo, sSQL
'''
'''sSQL = "SELECT C.Grupo,C.Codigo,C.Cliente,C.Direccion,C.CI_RUC,C.Casilla,SUM(CF.Valor) As Valor_Pension " _
'''     & "FROM Clientes As C,Clientes_Facturacion As CF " _
'''     & "WHERE C.FA <> 0 " _
'''     & "AND C.Codigo = CF.Codigo " _
'''     & "AND CF.Item = '" & NumEmpresa & "' " _
'''     & "AND CF.Periodo = '" & Periodo_Contable & "' " _
'''     & "AND CF.Codigo_Inv <> '01.88' " _
'''     & "GROUP BY C.Grupo,C.Codigo,C.Cliente,C.Direccion,C.CI_RUC,C.Casilla " _
'''     & "ORDER BY C.CI_RUC,C.Grupo,C.Cliente "
'''SelectAdodc AdoAux, sSQL
'''NumFile = FreeFile
'''Contador = 0
'''ProgBarra.Value = 0
'''ProgBarra.Min = 0
''''MsgBox RutaGeneraFile
'''Open RutaGeneraFile For Output As #NumFile ' Abre el archivo.
'''With AdoAux.Recordset
''' If .RecordCount > 0 Then
'''     Print #NumFile, "999";
'''     Print #NumFile, CodigoDelBanco;
'''     Print #NumFile, TipoDoc;
'''     Print #NumFile, "    ";
'''     Print #NumFile, FechaTexto
'''    .MoveFirst
'''     ProgBarra.Max = .RecordCount
'''     Do While Not .EOF
'''        FBancoBolivariano.Caption = .Fields("Grupo") & " - " & Format(Contador / .RecordCount, "00%")
'''        SaldoPendiente = 0
'''        Total_Factura = 0
'''        Total_Desc = 0
'''        Monto_Total = 0
'''        Total = 0
'''        ProgBarra.Value = Contador
'''        CodigoCli = .Fields("CI_RUC")
'''        Codigo = "0"
'''        For I = 1 To Len(.Fields("CI_RUC"))
'''            If IsNumeric(Mid(.Fields("CI_RUC"), I, 1)) Then Codigo = Codigo & Mid(.Fields("CI_RUC"), I, 1)
'''        Next I
'''        Codigo = Trim(Str(Val(Codigo)))
'''        If (8 - Len(Codigo)) >= 0 Then Codigo = Codigo & String(8 - Len(Codigo), " ")
'''        'MsgBox "|" & Codigo & "|"
'''        NombreCliente = SetearBlancos(Mid(.Fields("Cliente"), 1, 30), 30, 0, False)
'''        Codigo1 = Trim(Mid(SinEspaciosIzq(.Fields("Direccion")), 1, 15))
'''        Codigo3 = Trim(Mid(SinEspaciosDer(.Fields("Direccion")), 1, 3))
'''        Codigo2 = Trim(Mid(.Fields("Direccion"), Len(Codigo1) + 1, Len(.Fields("Direccion"))))
'''        Codigo4 = Mid(.Fields("Casilla"), 1, 10)
'''        Saldo_ME = 0
'''        If AdoGrupo.Recordset.RecordCount > 0 Then
'''           AdoGrupo.Recordset.MoveFirst
'''           AdoGrupo.Recordset.Find ("CI_RUC Like '" & CodigoCli & "' ")
'''           If Not AdoGrupo.Recordset.EOF Then
'''              Total_Desc = AdoGrupo.Recordset.Fields("Valor_Descuento")
'''           End If
'''        End If
'''        If AdoAct.Recordset.RecordCount > 0 Then
'''           AdoAct.Recordset.MoveFirst
'''           AdoAct.Recordset.Find ("CI_RUC Like '" & CodigoCli & "' ")
'''           If Not AdoAct.Recordset.EOF Then
'''              SaldoPendiente = AdoAct.Recordset.Fields("Saldo_Pend")
'''           End If
'''        End If
'''        If OpcMat.Value Then SaldoPendiente = 0
'''        If CheqPend.Value = 1 Then SaldoPendiente = 0
'''        Total_Factura = .Fields("Valor_Pension")
'''        Monto_Total = Total_Factura - Total_Desc
'''        Total = Monto_Total + SaldoPendiente
'''        If Codigo1 = "" Then Codigo1 = Ninguno
'''        If Codigo2 = "" Then Codigo2 = Ninguno
'''        If Codigo3 = "" Then Codigo3 = Ninguno
'''        Codigo2 = Trim(Mid(Codigo2, 1, Len(Codigo2) - Len(SinEspaciosDer(Codigo2))))
'''        Codigo1 = SetearBlancos(Codigo1, 15, 0, False)
'''        Codigo2 = SetearBlancos(Codigo2, 15, 0, False)
'''        Codigo3 = SetearBlancos(Codigo3, 3, 0, False)
'''        Codigo4 = SetearBlancos(Codigo4, 10, 0, False)
'''        If Trim(Codigo4) = Ninguno Then Codigo4 = String(10, " ")
'''      ' Empieza la trama por Alumno
'''        Print #NumFile, CodigoDelBanco;                       ' Colegio/Institucion
'''        Print #NumFile, Codigo;                               ' Codigo Alumno
'''        Print #NumFile, FechaTexto;                              ' Fecha Pen
'''        Print #NumFile, TipoDoc & "  ";                       ' Proceso
'''        Print #NumFile, Format(Total, "00000000.00");         ' Valor
'''        Print #NumFile, FechaTexto;                              ' Fecha Cobis
'''        Print #NumFile, "01/01/1900";                         ' Fecha Pago
'''        Print #NumFile, "N";                                  ' Estado = N
'''        Print #NumFile, NombreCliente;                        ' Nombre Alumno
'''        Print #NumFile, Codigo2;                              ' Nombre del Curso
'''        Print #NumFile, Codigo3;                              ' Nombre del Paralelo
'''        Print #NumFile, Codigo1;                              ' Nombre de la Seccion
'''        Print #NumFile, Format(Monto_Total, "00000000.00");   ' Valor Mes
'''        Print #NumFile, Codigo4;                              ' Pago por Deposito de Cuenta
'''        Print #NumFile, "1";                                  ' Moneda = 1
'''        Print #NumFile, Format(Total, "00000000.00");         ' Valor 2
'''        Print #NumFile, Format(Total, "00000000.00")          ' Valor 1
'''        Contador = Contador + 1
'''       .MoveNext
'''     Loop
''' End If
'''End With
'''Close #NumFile
'''ProgBarra.Value = ProgBarra.Max
'''MsgBox "Fin del Proceso"
'''End Sub

Private Sub Dir1_Change()
  File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
  Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_DblClick()
  SiguienteControl
End Sub

Private Sub File1_KeyDown(KeyCode As Integer, Shift As Integer)
  NombreArchivo = File1.FileName
  If KeyCode = vbKeyDelete Then
     Mensajes = "Esta seguro de Eliminar: " & File1.FileName
     Titulo = "Pregunta de Eliminacion"
     If BoxMensaje = vbYes Then Kill File1.Path & "\" & File1.FileName
     File1.FileName = Dir1.Path & "\*.*"
  End If

End Sub

Private Sub File1_LostFocus()
  NombreArchivo = UCase(File1.FileName)
  RatonReloj
  RutaGeneraFile = UCase(Dir1.Path & "\" & NombreArchivo)
  'MsgBox RutaGeneraFile
  NumFile = FreeFile: TxtFile = ""
  Open RutaGeneraFile For Input As #NumFile
    Do While Not EOF(NumFile)
       Line Input #NumFile, Cod_Field
       TxtFile = TxtFile & Cod_Field & vbCrLf
    Loop
  Close #NumFile
  RatonNormal
End Sub

Private Sub Form_Activate()
  Select Case TextoBanco
    Case "PICHINCHA":
         RutaOrigen = RutaSistema & "\LOGOS\PICHINCHA.GIF"
         FBancoBolivariano.BackColor = &H80FFFF
         FBancoBolivariano.Caption = "BANCO DEL PICHINCHA"
    Case "INTERNACIONAL":
         RutaOrigen = RutaSistema & "\LOGOS\INTERNACIONAL.GIF"
         FBancoBolivariano.BackColor = &HC00000
         FBancoBolivariano.Caption = "BANCO INTERNACIONAL"
    Case "BOLIVARIANO":
         RutaOrigen = RutaSistema & "\LOGOS\BOLIVARIANO.GIF"
         FBancoBolivariano.BackColor = &H808000
         FBancoBolivariano.Caption = "BANCO BOLIVARIANO"
    Case Else
         RutaOrigen = RutaSistema & "\LOGOS\DISKCOVS.GIF"
         FBancoBolivariano.BackColor = &HE0E0E0
         FBancoBolivariano.Caption = "OTROS BANCOS"
  End Select
  FBancoBolivariano.Picture = LoadPicture(RutaOrigen)
  FechaValida MBFechaI
  Drive1.Drive = Mid(RutaSysBases, 1, 2)
  RatonNormal
  RutaBackup = RutaSysBases & "\BANCO"
  TipoProcesos ""
  FBancoBolivariano.Caption = "FACTURACION DE BANCOS (" & CodigoDelBanco & ")"
  sSQL = "SELECT C.Grupo,CF.Codigo_Inv,MAX(CF.Valor) as Maximo " _
     & "FROM Clientes_Facturacion As CF, Clientes As C,Catalogo_Productos As CP " _
     & "WHERE CF.Item = '" & NumEmpresa & "' " _
     & "AND CP.Periodo = '" & Periodo_Contable & "' " _
     & "AND CF.Codigo = C.Codigo " _
     & "AND CF.Codigo_Inv = CP.Codigo_Inv " _
     & "AND CF.Item = CP.Item " _
     & "AND CF.Periodo = CP.Periodo " _
     & "GROUP BY C.Grupo,CF.Codigo_Inv " _
     & "ORDER BY C.Grupo,CF.Codigo_Inv "
 SelectAdodc AdoGrupo, sSQL

  sSQL = "SELECT * " _
       & "FROM Clientes " _
       & "WHERE FA <> 0 " _
       & "ORDER BY CI_RUC "
  SelectAdodc AdoAux, sSQL
  If AdoAux.Recordset.RecordCount > 0 Then
     AdoAux.Recordset.MoveLast
     Codigo = AdoAux.Recordset.Fields("CI_RUC")
  End If
  FBancoBolivariano.Caption = "BANCO BOLIVARIANO (" & CodigoDelBanco & ")" & String(40, " ") & "EL ULTIMO CODIGO: " & Codigo
  Label4.Caption = "ORIGEN" & Space(20) & "COD: " & CodigoDelBanco
  RatonNormal
End Sub

Private Sub Form_Load()

  CentrarForm FBancoBolivariano
  If CodigoUsuario = "ACCESO02" Then
     Command6.Visible = True
  End If
  RutaBackupXX = ""
  ConectarAdodc AdoAux
  ConectarAdodc AdoAct
  ConectarAdodc AdoGrupo
  ConectarAdodc AdoQuery
End Sub

Private Sub MBFechaI_GotFocus()
  MarcarTexto MBFechaI
End Sub

Private Sub MBFechaI_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBFechaI_LostFocus()
  FechaValida MBFechaI
End Sub

Private Sub MBFechaF_GotFocus()
  MarcarTexto MBFechaF
End Sub

Private Sub MBFechaF_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBFechaF_LostFocus()
  FechaValida MBFechaF
End Sub

