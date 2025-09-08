VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FPagosBancos 
   Caption         =   "CONCILIACION DE BANCOS"
   ClientHeight    =   7200
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11490
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7200
   ScaleWidth      =   11490
   WindowState     =   2  'Maximized
   Begin VB.Frame Banco 
      BackColor       =   &H00C0FFFF&
      Caption         =   "SELECCIONE EL BANCO"
      Height          =   3165
      Left            =   3675
      TabIndex        =   10
      Top             =   1260
      Visible         =   0   'False
      Width           =   8205
      Begin MSDataListLib.DataList DLListBanco 
         Bindings        =   "PagoBank.frx":0000
         DataSource      =   "AdoListBanco"
         Height          =   2790
         Left            =   105
         TabIndex        =   11
         Top             =   210
         Width           =   7995
         _ExtentX        =   14102
         _ExtentY        =   4921
         _Version        =   393216
         BackColor       =   16777152
      End
   End
   Begin MSDataGridLib.DataGrid DGBalance 
      Bindings        =   "PagoBank.frx":001B
      Height          =   5790
      Left            =   105
      TabIndex        =   9
      Top             =   1050
      Width           =   11250
      _ExtentX        =   19844
      _ExtentY        =   10213
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      RowDividerStyle =   4
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   12298
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   12298
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSDataListLib.DataCombo DCCtas 
      Bindings        =   "PagoBank.frx":0035
      DataSource      =   "AdoBanco"
      Height          =   315
      Left            =   105
      TabIndex        =   4
      Top             =   525
      Width           =   7365
      _ExtentX        =   12991
      _ExtentY        =   556
      _Version        =   393216
      Text            =   "DataCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSAdodcLib.Adodc AdoBanco 
      Height          =   330
      Left            =   315
      Top             =   1155
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
      Caption         =   "Banco"
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
   Begin VB.CommandButton Command1 
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
      Height          =   855
      Left            =   10710
      Picture         =   "PagoBank.frx":004C
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   105
      Width           =   960
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Grabar"
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
      Left            =   8610
      Picture         =   "PagoBank.frx":0916
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   105
      Width           =   960
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Generar"
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
      Left            =   9660
      Picture         =   "PagoBank.frx":0D58
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   105
      Width           =   960
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Procesar"
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
      Left            =   7560
      Picture         =   "PagoBank.frx":1622
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   105
      Width           =   960
   End
   Begin MSMask.MaskEdBox MBoxFechaF 
      Height          =   330
      Left            =   3150
      TabIndex        =   3
      Top             =   105
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   582
      _Version        =   393216
      AllowPrompt     =   -1  'True
      AutoTab         =   -1  'True
      MaxLength       =   10
      OLEDragMode     =   1
      OLEDropMode     =   2
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
      OLEDragMode     =   1
      OLEDropMode     =   2
   End
   Begin MSMask.MaskEdBox MBoxFechaI 
      Height          =   330
      Left            =   840
      TabIndex        =   1
      Top             =   105
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   582
      _Version        =   393216
      AllowPrompt     =   -1  'True
      AutoTab         =   -1  'True
      MaxLength       =   10
      OLEDragMode     =   1
      OLEDropMode     =   2
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
      OLEDragMode     =   1
      OLEDropMode     =   2
   End
   Begin MSAdodcLib.Adodc AdoCtas 
      Height          =   330
      Left            =   315
      Top             =   1470
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
      Caption         =   "Ctas"
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
   Begin MSAdodcLib.Adodc AdoAsientos 
      Height          =   330
      Left            =   105
      Top             =   6825
      Width           =   11250
      _ExtentX        =   19844
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
      Caption         =   "Asientos"
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
   Begin MSAdodcLib.Adodc AdoListBanco 
      Height          =   330
      Left            =   315
      Top             =   1785
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
      Caption         =   "ListBanco"
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
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&Hasta"
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
      Left            =   2415
      TabIndex        =   2
      Top             =   105
      Width           =   750
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&Desde"
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
      Width           =   750
   End
End
Attribute VB_Name = "FPagosBancos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub Generar_Pagos_Produbanco()
Dim AuxNumEmp As String
Dim DiaV As Integer
Dim MesV As Integer
Dim AñoV As Integer
Dim NumFileFacturas As Long
Dim Fecha_Meses As String
Dim ValorStr As String
Dim GrupoNo As String
Dim CamposFile() As Campos_Tabla
Dim Total_Banco As Currency
Dim EsComa As Boolean
Dim Estab As Boolean

  DGBalance.Visible = False
  Llenar_Pagos True
  
  Cta_Bancaria = SinEspaciosDer(DCCtas)
  Cta_Bancaria = Replace(Cta_Bancaria, "-", "")
  
  sSQL = "SELECT APB.*,C.TD,C.CI_RUC,C.Ciudad,C.Direccion,C.Telefono,C.Email " _
       & "FROM Asiento_PB As APB,Clientes As C " _
       & "WHERE APB.CodigoU = '" & CodigoUsuario & "' " _
       & "AND APB.Codigo_B = C.Codigo "
  If ConSucursal = False Then sSQL = sSQL & "AND APB.Item = '" & NumEmpresa & "' "
  sSQL = sSQL & "AND APB.T_No = " & Trans_No & " "
  Select_Adodc AdoAsientos, sSQL
  
  Total_Banco = 0
  Mifecha = BuscarFecha(MBoxFechaI)
  MiMes = Format$(Month(MBoxFechaI), "00")
  FechaFin = BuscarFecha(UltimoDiaMes(MBoxFechaI))
  TextoImprimio = ""
  AuxNumEmp = NumEmpresa
  Cta_Cobrar = Ninguno
  FechaTexto = FechaSistema
' Comenzamos a generar el archivo: SCRECXX.TXT
  EsComa = True
  Estab = False
  Contador = 0
  Factura_No = 0
  NumFileFacturas = FreeFile
  Fecha_Meses = MBoxFechaI & " al " & MBoxFechaF
  Fecha_Meses = Replace(Fecha_Meses, "/", "-")
  RutaGeneraFile = UCaseStrg(RutaSysBases & "\BANCO\PAGOS_" & Fecha_Meses & ".TXT")
  Open RutaGeneraFile For Output As #NumFileFacturas ' Abre el archivo.
  With AdoAsientos.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       Do While Not .EOF
          Contador = Contador + 1
          CodigoCli = .fields("Codigo_B")
          CodigoDelBanco = Format(Val(.fields("Codigo_Banco")), "0000")
          NombreCliente = TrimStrg(MidStrg(Sin_Signos_Especiales(.fields("BENEFICIARIO")), 1, 40))
         'MsgBox NombreCliente
          Factura_No = Factura_No + 1
          Total = .fields("MONTO")
          Saldo = .fields("MONTO") * 100
          ValorStr = CStr(Saldo)
          ValorStr = String(13 - Len(ValorStr), "0") & ValorStr
         'MsgBox ValorStr
          CodigoP = .fields("CI_RUC")
          CodigoC = CStr(Val(.fields("CI_RUC")))
          CodigoC = CodigoC & String(Abs(4 - Len(CodigoC)), " ")
          
          DireccionCli = Ninguno
          Codigo3 = SinEspaciosDer(DireccionCli)
          DireccionCli = TrimStrg(MidStrg(DireccionCli, 1, Len(DireccionCli) - Len(Codigo3)))
          Codigo3 = TrimStrg(SinEspaciosDer(DireccionCli))
          Codigo1 = Format$(MidStrg(GrupoNo, 1, 1), "00")
          If Len(Cta_Bancaria) < 11 Then Cta_Bancaria = String(11 - Len(Cta_Bancaria), "0") & Cta_Bancaria
          Codigo4 = Ninguno
          Print #NumFileFacturas, "PA" & vbTab;                 '01
          Print #NumFileFacturas, Cta_Bancaria & vbTab;         '02
          Print #NumFileFacturas, Contador & vbTab;             '03
          Print #NumFileFacturas, vbTab;                        '04
          Print #NumFileFacturas, .fields("CI_RUC") & vbTab;    '05
          Print #NumFileFacturas, "USD" & vbTab;                '06
          Print #NumFileFacturas, ValorStr & vbTab;             '07
          Print #NumFileFacturas, "CTA" & vbTab;                '08
          Print #NumFileFacturas, CodigoDelBanco & vbTab;       '09
          Print #NumFileFacturas, .fields("TIPO_CTA") & vbTab;  '10
          Print #NumFileFacturas, .fields("CTA_TRANS") & vbTab; '11
          Print #NumFileFacturas, .fields("TD") & vbTab;        '12
          Print #NumFileFacturas, .fields("CI_RUC") & vbTab;    '13
          Print #NumFileFacturas, NombreCliente & vbTab;        '14
          Print #NumFileFacturas, .fields("Direccion") & vbTab; '15
          Print #NumFileFacturas, .fields("Ciudad") & vbTab;    '16
          Print #NumFileFacturas, .fields("Telefono") & vbTab;  '17
          Print #NumFileFacturas, "SN" & vbTab;                 '18
          Print #NumFileFacturas, "PAGO CE No. " & .fields("NUMERO") & vbTab;          '19
          Print #NumFileFacturas, .fields("Email") & vbTab      '20
         .MoveNext
       Loop
   End If
  End With
  Close #NumFileFacturas
  RatonNormal
  'ProgBarra.value = ProgBarra.Max
 'FBancoPichincha.Caption = "FACTURACION DE BANCO INTERNACIONAL"
  sSQL = "SELECT APB.*,C.TD,C.CI_RUC,C.Ciudad,C.Direccion,C.Telefono,C.Email " _
       & "FROM Asiento_PB As APB,Clientes As C " _
       & "WHERE APB.CodigoU = '" & CodigoUsuario & "' " _
       & "AND APB.Codigo_B = C.Codigo "
  If ConSucursal = False Then sSQL = sSQL & "AND APB.Item = '" & NumEmpresa & "' "
  sSQL = sSQL & "AND APB.T_No = " & Trans_No & " "
  Select_Adodc_Grid DGBalance, AdoAsientos, sSQL
  DGBalance.Visible = True
  MsgBox "SE HA GENERADO EL SIGUIENTE ARCHIVO:" & vbCrLf & vbCrLf _
       & RutaGeneraFile & vbCrLf
End Sub


Private Sub Command1_Click()
  Unload FPagosBancos
End Sub

Private Sub Command2_Click()
  Generar_Pagos_Produbanco
End Sub

Private Sub Command3_Click()
  Llenar_Pagos
  sSQL = "SELECT * " _
       & "FROM Asiento_PB " _
       & "WHERE CodigoU = '" & CodigoUsuario & "' "
  If ConSucursal = False Then sSQL = sSQL & "AND Item = '" & NumEmpresa & "' "
  sSQL = sSQL & "AND T_No = " & Trans_No & " "
  Select_Adodc_Grid DGBalance, AdoAsientos, sSQL
  SumaDebe = 0: SumaHaber = 0
  DGBalance.Visible = True
  Cadena = "Registros: " & Format(AdoCtas.Recordset.RecordCount, "#,##0") & ".   Páginas: " _
         & Format((AdoCtas.Recordset.RecordCount / 45) + 1, "#,##0") & "."
  AdoCtas.Caption = Cadena
End Sub

Private Sub Command4_Click()
  RatonReloj
  DGBalance.Visible = False
  Codigo1 = SinEspaciosIzq(DCCtas.Text)
  With AdoAsientos.Recordset
   If .RecordCount > 0 Then
       If SQL_Server Then
          sSQL = "UPDATE Transacciones " _
               & "SET Pagar = PB.Pagar " _
               & "FROM Transacciones As T,Asiento_PB As PB "
       Else
          sSQL = "UPDATE Transacciones As T,Asiento_PB As PB " _
               & "SET T.C = PB.C "
       End If
       sSQL = sSQL & "WHERE PB.TP = T.TP " _
            & "AND PB.Numero = T.Numero " _
            & "AND PB.MONTO = T.Haber " _
            & "AND PB.FECHA = T.Fecha " _
            & "AND PB.Item = T.Item " _
            & "AND PB.CodigoU = '" & CodigoUsuario & "' " _
            & "AND PB.T_No = " & Trans_No & " " _
            & "AND T.Periodo = '" & Periodo_Contable & "' " _
            & "AND T.Cta = '" & Codigo1 & "' "
       Ejecutar_SQL_SP sSQL
       MsgBox "Proceso Grabado"
   End If
  End With
  SumaDebe = 0: SumaHaber = 0
  DGBalance.Visible = True
  RatonNormal
End Sub

Private Sub DCCtas_LostFocus()
  Codigo1 = SinEspaciosIzq(DCCtas.Text)
  Codigo = Leer_Cta_Catalogo(Codigo1)
End Sub

Private Sub DGBalance_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  If CtrlDown And KeyCode = vbKeyF1 Then GenerarDataTexto FPagosBancos, AdoAsientos
End Sub

Private Sub DGBalance_KeyPress(KeyAscii As Integer)
   Codigo1 = DGBalance.Columns(15)
   Select Case Chr(KeyAscii)
     Case "s", "S", "y", "Y": ' Si
          NivelCta = 1
     Case "n", "N"  ' No
          NivelCta = 0
   End Select
   Select Case Chr(KeyAscii)
     Case "s", "S", "y", "Y", "n", "N":
          If AdoAsientos.Recordset.RecordCount > 0 Then
               sSQL = "UPDATE Asiento_PB " _
                    & "SET Pagar = " & NivelCta & " " _
                    & "WHERE IdTrans = '" & Codigo1 & "' "
               If ConSucursal = False Then sSQL = sSQL & "AND Item = '" & NumEmpresa & "' "
               sSQL = sSQL & "AND CodigoU = '" & CodigoUsuario & "' " _
                    & "AND T_No = " & Trans_No & " "
               Ejecutar_SQL_SP sSQL
               sSQL = "SELECT * " _
                    & "FROM Asiento_PB " _
                    & "WHERE CodigoU = '" & CodigoUsuario & "' "
               If ConSucursal = False Then sSQL = sSQL & "AND Item = '" & NumEmpresa & "' "
               sSQL = sSQL & "AND T_No = " & Trans_No & " " _
                    & "ORDER BY IdTrans "
               Select_Adodc_Grid DGBalance, AdoAsientos, sSQL
               AdoAsientos.Recordset.MoveFirst
               AdoAsientos.Recordset.Find ("IdTrans = '" & Codigo1 & "' ")
               DGBalance.SetFocus
          End If
   End Select
End Sub

Private Sub Form_Activate()
  sSQL = "SELECT (Codigo & Space(20) & Cuenta) As Nombre_Cta " _
       & "FROM Catalogo_Cuentas " _
       & "WHERE TC = 'BA' " _
       & "AND DG = 'D' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "ORDER BY Codigo "
  SelectDB_Combo DCCtas, AdoBanco, sSQL, "Nombre_Cta", False
  
  sSQL = "SELECT Codigo, Descripcion " _
       & "FROM Tabla_Referenciales_SRI " _
       & "WHERE Tipo_Referencia = 'BANCOS Y COOP' " _
       & "ORDER BY Descripcion "
  SelectDB_List DLListBanco, AdoListBanco, sSQL, "Descripcion"
  DGBalance.Visible = False
  RatonNormal
End Sub

Private Sub Form_Load()
  ConectarAdodc AdoCtas
  ConectarAdodc AdoBanco
  ConectarAdodc AdoListBanco
  ConectarAdodc AdoAsientos
  Trans_No = 35
  DGBalance.Height = MDI_Y_Max - DGBalance.Top - 300
  DGBalance.width = MDI_X_Max - DGBalance.Left - 100
  AdoAsientos.Top = DGBalance.Top + DGBalance.Height + 10
  AdoAsientos.width = MDI_X_Max - AdoAsientos.Left - 100
End Sub

Private Sub MBoxFechaF_GotFocus()
  MarcarTexto MBoxFechaF
End Sub

Private Sub MBoxFechaF_LostFocus()
  FechaValida MBoxFechaF
End Sub

Private Sub MBoxFechaI_GotFocus()
  MarcarTexto MBoxFechaI
End Sub

Private Sub MBoxFechaI_LostFocus()
  FechaValida MBoxFechaI
  MBoxFechaF = UltimoDiaMes(MBoxFechaI)
End Sub

Public Sub Llenar_Pagos(Optional Pagar As Boolean)
  RatonReloj
  FechaValida MBoxFechaI, False
  FechaValida MBoxFechaF, False
  Contador = 0
  Codigo1 = SinEspaciosIzq(DCCtas.Text)
  FechaIni = BuscarFecha(MBoxFechaI.Text)
  FechaFin = BuscarFecha(MBoxFechaF.Text)
  
  sSQL = "DELETE * " _
       & "FROM Asiento_PB " _
       & "WHERE CodigoU = '" & CodigoUsuario & "' "
  If ConSucursal = False Then sSQL = sSQL & "AND Item = '" & NumEmpresa & "' "
  sSQL = sSQL & "AND T_No = " & Trans_No & " "
  Ejecutar_SQL_SP sSQL
  sSQL = "SELECT * " _
       & "FROM Asiento_PB " _
       & "WHERE CodigoU = '" & CodigoUsuario & "' "
  If ConSucursal = False Then sSQL = sSQL & "AND Item = '" & NumEmpresa & "' "
  sSQL = sSQL & "AND T_No = " & Trans_No & " "
  Select_Adodc AdoAsientos, sSQL
  sSQL = "SELECT T.*, Cl.Cliente, Cl.Codigo, Cl.Actividad " _
       & "FROM Transacciones As T,Comprobantes As C,Clientes AS Cl " _
       & "WHERE T.Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
       & "AND T.Cta = '" & Codigo1 & "' "
  If ConSucursal = False Then sSQL = sSQL & "AND T.Item = '" & NumEmpresa & "' "
  If Pagar Then sSQL = sSQL & "AND T.Pagar <> " & Val(adFalse) & " "
  sSQL = sSQL _
       & "AND T.Periodo = '" & Periodo_Contable & "' " _
       & "AND C.T <> 'A' " _
       & "AND C.TP = 'CE' " _
       & "AND C.TP = T.TP " _
       & "AND C.Numero = T.Numero " _
       & "AND C.Item = T.Item " _
       & "AND T.Periodo = C.Periodo " _
       & "AND C.Codigo_B = Cl.Codigo " _
       & "ORDER BY T.Fecha,T.TP,T.Numero,Debe DESC,Haber,T.ID "
  Select_Adodc AdoCtas, sSQL
  RatonReloj
  DGBalance.Visible = False
  With AdoCtas.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
         Contador = Contador + 1
         FPagosBancos.Caption = "Conciliando la fecha: " & .fields("Fecha")
         Codigos = .fields("Actividad")
         NombreBanco = Ninguno
         CodigoB = Ninguno
         TipoCta = Ninguno
         Cta = Ninguno
         If Len(Codigo) > 7 Then
            CodigoB = SinEspaciosIzq(Codigos)
            Codigos = TrimStrg(MidStrg(Codigos, Len(CodigoB) + 1, Len(Codigos)))
            TipoCta = SinEspaciosIzq(Codigos)
            Cta = SinEspaciosDer(Codigos)
            CodigoB = CStr(Val(CodigoB))
         End If
         If AdoListBanco.Recordset.RecordCount > 0 Then
            AdoListBanco.Recordset.MoveFirst
            AdoListBanco.Recordset.Find ("Codigo = '" & CodigoB & "' ")
            If Not AdoListBanco.Recordset.EOF Then NombreBanco = AdoListBanco.Recordset.fields("Descripcion")
         End If
         SetAddNew AdoAsientos
         SetFields AdoAsientos, "Pagar", .fields("Pagar")
         SetFields AdoAsientos, "FECHA", .fields("Fecha")
         SetFields AdoAsientos, "BENEFICIARIO", .fields("Cliente")
         SetFields AdoAsientos, "TP", .fields("TP")
        'SetFields AdoAsientos, "FACTURA", .Fields("Factura")
         SetFields AdoAsientos, "NUMERO", .fields("Numero")
         SetFields AdoAsientos, "CTA_TRANS", Cta
         SetFields AdoAsientos, "MONTO", .fields("Haber")
         SetFields AdoAsientos, "Codigo_B", .fields("Codigo")
         SetFields AdoAsientos, "TIPO_CTA", TipoCta
         SetFields AdoAsientos, "BANCO", NombreBanco
         SetFields AdoAsientos, "Codigo_Banco", CodigoB
         SetFields AdoAsientos, "Item", .fields("Item")
         SetFields AdoAsientos, "CodigoU", CodigoUsuario
         SetFields AdoAsientos, "T_No", Trans_No
         SetFields AdoAsientos, "IdTrans", Format(Contador, "0000")
         SetUpdate AdoAsientos
        .MoveNext
       Loop
   End If
  End With
  FPagosBancos.Caption = "CONCILIACION DE BANCOS"
  RatonNormal
End Sub
