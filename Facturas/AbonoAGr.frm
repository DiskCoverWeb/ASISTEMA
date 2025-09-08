VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form AbonoAnticipoGrupo 
   Caption         =   "INGRESO DE CAJA"
   ClientHeight    =   6420
   ClientLeft      =   5100
   ClientTop       =   4455
   ClientWidth     =   13140
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6420
   ScaleWidth      =   13140
   WindowState     =   2  'Maximized
   Begin MSDataGridLib.DataGrid DGFactura 
      Bindings        =   "AbonoAGr.frx":0000
      Height          =   5265
      Left            =   105
      TabIndex        =   11
      Top             =   1050
      Width           =   12930
      _ExtentX        =   22807
      _ExtentY        =   9287
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
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
   Begin VB.CommandButton Command3 
      Caption         =   "&Consultar"
      Height          =   855
      Left            =   9975
      Picture         =   "AbonoAGr.frx":0019
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   105
      Width           =   960
   End
   Begin VB.TextBox TxtRecibo 
      Alignment       =   1  'Right Justify
      Height          =   330
      Left            =   315
      MaxLength       =   14
      TabIndex        =   1
      Text            =   "0"
      Top             =   420
      Width           =   1275
   End
   Begin VB.CheckBox CheqRecibo 
      Caption         =   "&RECIBO No."
      Height          =   330
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Value           =   1  'Checked
      Width           =   1485
   End
   Begin MSAdodcLib.Adodc AdoFactura 
      Height          =   330
      Left            =   420
      Top             =   1470
      Visible         =   0   'False
      Width           =   1905
      _ExtentX        =   3360
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
      Caption         =   "Factura"
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
      Caption         =   "&Salir"
      Height          =   855
      Left            =   12075
      Picture         =   "AbonoAGr.frx":0323
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   105
      Width           =   960
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Grabar"
      Height          =   855
      Left            =   11025
      Picture         =   "AbonoAGr.frx":0BED
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   105
      Width           =   960
   End
   Begin MSMask.MaskEdBox MBFecha 
      Height          =   330
      Left            =   1680
      TabIndex        =   3
      Top             =   420
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   582
      _Version        =   393216
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
   Begin MSDataListLib.DataCombo DCBanco 
      Bindings        =   "AbonoAGr.frx":102F
      DataSource      =   "AdoBanco"
      Height          =   315
      Left            =   3045
      TabIndex        =   5
      Top             =   420
      Width           =   4950
      _ExtentX        =   8731
      _ExtentY        =   556
      _Version        =   393216
      Text            =   "Banco"
   End
   Begin MSAdodcLib.Adodc AdoBanco 
      Height          =   330
      Left            =   420
      Top             =   1785
      Visible         =   0   'False
      Width           =   1905
      _ExtentX        =   3360
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
   Begin MSAdodcLib.Adodc AdoDetAcomp 
      Height          =   330
      Left            =   420
      Top             =   2100
      Visible         =   0   'False
      Width           =   1905
      _ExtentX        =   3360
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
      Caption         =   "DetAcomp"
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
   Begin MSAdodcLib.Adodc AdoSerie 
      Height          =   330
      Left            =   420
      Top             =   2415
      Visible         =   0   'False
      Width           =   1905
      _ExtentX        =   3360
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
      Caption         =   "Serie"
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
   Begin MSAdodcLib.Adodc AdoCliente 
      Height          =   330
      Left            =   420
      Top             =   2730
      Visible         =   0   'False
      Width           =   1905
      _ExtentX        =   3360
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
      Caption         =   "Cliente"
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
   Begin MSAdodcLib.Adodc AdoAutorizacion 
      Height          =   330
      Left            =   420
      Top             =   3045
      Visible         =   0   'False
      Width           =   1905
      _ExtentX        =   3360
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
      Caption         =   "Autorizacion"
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
   Begin VB.Label LabelPend 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   330
      Left            =   8085
      TabIndex        =   7
      Top             =   420
      Width           =   1800
   End
   Begin VB.Label Label10 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Total de Abonos"
      Height          =   330
      Left            =   8085
      TabIndex        =   6
      Top             =   105
      Width           =   1800
   End
   Begin VB.Label Label16 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Cuenta Contable del Abono"
      Height          =   330
      Left            =   3045
      TabIndex        =   4
      Top             =   105
      Width           =   4950
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &Fecha"
      Height          =   330
      Left            =   1680
      TabIndex        =   2
      Top             =   105
      Width           =   1275
   End
End
Attribute VB_Name = "AbonoAnticipoGrupo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
  FechaValida MBFecha
  FechaTexto = MBFecha
  Mensajes = "Esta Seguro que desea grabar los Abonos"
  Titulo = "Formulario de Grabación"
  If BoxMensaje = vbYes Then
     RatonReloj
     FechaTexto = MBFecha ' FechaSistema
     DGFactura.Visible = False
     If CheqRecibo.value = 1 Then
        DiarioCaja = ReadSetDataNum("Recibo_No", True, True)
     Else
        DiarioCaja = Val(TxtRecibo)
     End If
     With AdoFactura.Recordset
      If .RecordCount > 0 Then
         .MoveFirst
          Do While Not .EOF
            'Abono de Factura
             TA.T = Normal
             TA.TP = .fields("TC")
             TA.Fecha = MBFecha
             TA.Factura = .fields("Factura")
             TA.Serie = .fields("Serie")
             TA.Autorizacion = .fields("Autorizacion")
             TA.Cta = .fields("Cta_Abono")
             TA.Cta_CxP = .fields("Cta_Cobrar")
             TA.Banco = "ABONO ANTICIPADO"
             TA.Cheque = .fields("Grupo")
             TA.Abono = .fields("Abono")
             TA.CodigoC = .fields("Codigo_Cliente")
             TA.Recibi_de = .fields("Cliente")
             TA.Recibo_No = Format$(DiarioCaja, "0000000000")
             Grabar_Abonos TA
            'Tipo de Abonos con SubCtas
            ' Grabar_Anticipos TA
             T = "P"
             Saldo = .fields("Saldo") - .fields("Abono")
             If Saldo <= 0 Then T = "C"
             sSQL = "UPDATE Facturas " _
                  & "SET Saldo_MN = " & Saldo & ", T = '" & T & "' " _
                  & "WHERE Item = '" & NumEmpresa & "' " _
                  & "AND Periodo = '" & Periodo_Contable & "' " _
                  & "AND CodigoC = '" & TA.CodigoC & "' " _
                  & "AND TC = '" & TA.TP & "' " _
                  & "AND Serie = '" & TA.Serie & "' " _
                  & "AND Factura = " & TA.Factura & " " _
                  & "AND Autorizacion = '" & TA.Autorizacion & "' "
             Ejecutar_SQL_SP sSQL
             DiarioCaja = DiarioCaja + 1
            .MoveNext
          Loop
      End If
     End With
     DGFactura.Visible = False
     Actualizar_Saldos_Facturas_SP
     RatonNormal
     DGFactura.Visible = True
     MsgBox "Abonos Realizados con éxito"
     TxtRecibo.SetFocus
  End If
 'Unload AbonoEfectivo
End Sub

Private Sub Command2_Click()
   Control_Procesos Normal, "Salir de abonos de facturas por Anticipos en Grupo"
   Unload Me
End Sub

Private Sub Command3_Click()
  FechaValida MBFecha
  Cta_Aux = SinEspaciosIzq(DCBanco)
  TotalSubCta = 0
  
  sSQL = "DELETE * " _
       & "FROM Asiento_F " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' "
  Ejecutar_SQL_SP sSQL
  
  sSQL = "SELECT C.Cliente,SUM(TS.Creditos-TS.Debitos) As Saldo_Pendiente,C.Grupo,C.CI_RUC,TS.Codigo " _
       & "FROM Trans_SubCtas As TS,Comprobantes As CO,Clientes As C " _
       & "WHERE TS.Item = '" & NumEmpresa & "' " _
       & "AND TS.Periodo = '" & Periodo_Contable & "' " _
       & "AND TS.Cta = '" & Cta_Aux & "' " _
       & "AND TS.Fecha <= #" & BuscarFecha(MBFecha) & "# " _
       & "AND CO.T <> 'A' " _
       & "AND TS.Codigo = C.Codigo " _
       & "AND TS.Item = CO.Item " _
       & "AND TS.Periodo = CO.Periodo " _
       & "AND TS.TP = CO.TP " _
       & "AND TS.Numero = CO.Numero " _
       & "GROUP BY C.Cliente,C.Grupo,C.CI_RUC,TS.Codigo " _
       & "HAVING SUM(TS.Creditos-TS.Debitos) > 0 " _
       & "ORDER BY C.Cliente "
  Select_Adodc AdoCliente, sSQL
  DGFactura.Visible = False
  With AdoCliente.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          CodigoCli = .fields("Codigo")
          NombreCliente = .fields("Cliente")
          Total = .fields("Saldo_Pendiente")
          Grupo_No = .fields("Grupo")
          SQL1 = "SELECT TC,Factura,Autorizacion,Serie,CodigoC,Saldo_MN,Cta_CxP " _
               & "FROM Facturas " _
               & "WHERE T = '" & Pendiente & "' " _
               & "AND Item = '" & NumEmpresa & "' " _
               & "AND Periodo = '" & Periodo_Contable & "' " _
               & "AND Fecha <= #" & BuscarFecha(MBFecha) & "# " _
               & "AND Saldo_MN > 0 " _
               & "AND CodigoC = '" & CodigoCli & "' " _
               & "ORDER BY TC,Fecha,Factura "
          Select_Adodc AdoFactura, SQL1
          If AdoFactura.Recordset.RecordCount > 0 Then
             Do While Not AdoFactura.Recordset.EOF
                TipoCta = AdoFactura.Recordset.fields("TC")
                Factura_No = AdoFactura.Recordset.fields("Factura")
                Autorizacion = AdoFactura.Recordset.fields("Autorizacion")
                SerieFactura = AdoFactura.Recordset.fields("Serie")
                Cta = AdoFactura.Recordset.fields("Cta_CxP")
                Saldo = AdoFactura.Recordset.fields("Saldo_MN")
                If Total > 0 Then
                   SetAdoAddNew "Asiento_F"
                   SetAdoFields "CODIGO", Cta_Aux
                   SetAdoFields "PRODUCTO", NombreCliente
                   SetAdoFields "FECHA", MBFecha
                   SetAdoFields "HABIT", TipoCta
                   SetAdoFields "Codigo_Cliente", CodigoCli
                   SetAdoFields "Mes", Grupo_No
                   SetAdoFields "Cta", Cta
                   SetAdoFields "Numero", Factura_No
                   SetAdoFields "Serie", SerieFactura
                   SetAdoFields "Autorizacion", Autorizacion
                   SetAdoFields "PRECIO", Saldo
                   If Total >= Saldo Then
                      SetAdoFields "TOTAL", Saldo
                      Total = Total - Saldo
                   Else
                      SetAdoFields "TOTAL", Total
                      Total = -1
                   End If
                   SetAdoUpdate
                End If
                AdoFactura.Recordset.MoveNext
             Loop
          End If
         .MoveNext
       Loop
   End If
  End With
  Total = 0
  sSQL = "SELECT C.Cliente,F.Numero As Factura,F.PRECIO As Saldo,F.TOTAL As Abono," _
       & "C.Grupo,F.HABIT As TC,F.Serie,F.Autorizacion,F.CODIGO As Cta_Abono,F.Cta As Cta_Cobrar,F.Codigo_Cliente " _
       & "FROM Asiento_F As F,Clientes As C " _
       & "WHERE F.Item = '" & NumEmpresa & "' " _
       & "AND F.CodigoU = '" & CodigoUsuario & "' " _
       & "AND F.Codigo_Cliente = C.Codigo " _
       & "ORDER BY C.Cliente,F.Numero "
  Select_Adodc_Grid DGFactura, AdoFactura, sSQL
  With AdoFactura.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          Total = Total + .fields("Abono")
         .MoveNext
       Loop
      .MoveFirst
   End If
  End With
  DGFactura.Visible = True
  LabelPend.Caption = Format$(Total, "#,##0.00")
End Sub

Private Sub DCBanco_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DGFactura_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  If KeyCode = vbKeyF1 Then
     DGFactura.Visible = False
     GenerarDataTexto AbonoAnticipoGrupo, AdoFactura
     DGFactura.Visible = True
  End If
End Sub

Private Sub Form_Activate()

  sSQL = "SELECT (TS.Cta & '  ' & CC.Cuenta) As NomCuenta,TS.TC " _
       & "FROM Catalogo_Cuentas CC,Trans_SubCtas As TS " _
       & "WHERE CC.Item = '" & NumEmpresa & "' " _
       & "AND CC.Periodo = '" & Periodo_Contable & "' " _
       & "AND TS.TC = 'P' " _
       & "AND CC.Codigo = TS.Cta " _
       & "AND CC.TC = TS.TC " _
       & "AND CC.Item = TS.Item " _
       & "AND CC.Periodo = TS.Periodo " _
       & "GROUP BY TS.Cta,CC.Cuenta,TS.TC " _
       & "ORDER BY TS.Cta "
  SelectDB_Combo DCBanco, AdoBanco, sSQL, "NomCuenta"

  DiarioCaja = ReadSetDataNum("Recibo_No", True, False)
  If CheqRecibo.value = 1 Then TxtRecibo = Format$(DiarioCaja, "0000000") Else TxtRecibo = ""
  Mifecha = BuscarFecha(FechaTexto)
  TxtBanco = Ninguno
  TextCheqNo = Ninguno
  MBFecha = FechaSistema
  RatonNormal
End Sub

Private Sub Form_Load()
   ConectarAdodc AdoSerie
   ConectarAdodc AdoBanco
   ConectarAdodc AdoCliente
   ConectarAdodc AdoFactura
   ConectarAdodc AdoDetAcomp
   ConectarAdodc AdoAutorizacion
End Sub

Private Sub MBFecha_GotFocus()
  MarcarTexto MBFecha
End Sub

Private Sub MBFecha_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBFecha_LostFocus()
  FechaValida MBFecha, True
End Sub

Private Sub TxtRecibo_GotFocus()
  MarcarTexto TxtRecibo
End Sub

Private Sub TxtRecibo_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtRecibo_LostFocus()
  TxtRecibo = Format$(Val(TxtRecibo), "0000000")
End Sub

