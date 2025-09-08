VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form AbonoAnticipado 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "INGRESO DE ABONOS ANTICIPADOS"
   ClientHeight    =   4875
   ClientLeft      =   5025
   ClientTop       =   4380
   ClientWidth     =   7185
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   7185
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Caption         =   "AbonoAnticipado"
      Height          =   3375
      Left            =   105
      TabIndex        =   23
      Top             =   525
      Width           =   5895
      Begin VB.TextBox TxtConcepto 
         Height          =   645
         Left            =   105
         MultiLine       =   -1  'True
         TabIndex        =   27
         Top             =   1155
         Width           =   5685
      End
      Begin MSDataListLib.DataCombo DCClientes 
         Bindings        =   "AbonoAnt.frx":0000
         DataSource      =   "AdoClientes"
         Height          =   315
         Left            =   945
         TabIndex        =   25
         Top             =   210
         Width           =   4845
         _ExtentX        =   8546
         _ExtentY        =   556
         _Version        =   393216
         Text            =   "Cliente"
      End
      Begin MSDataListLib.DataCombo DCBanco 
         Bindings        =   "AbonoAnt.frx":001A
         DataSource      =   "AdoBanco"
         Height          =   315
         Left            =   105
         TabIndex        =   29
         Top             =   2205
         Width           =   5685
         _ExtentX        =   10028
         _ExtentY        =   556
         _Version        =   393216
         Text            =   "Banco"
      End
      Begin MSDataListLib.DataCombo DCCtaAnt 
         Bindings        =   "AbonoAnt.frx":0031
         DataSource      =   "AdoCtaAnt"
         Height          =   315
         Left            =   105
         TabIndex        =   31
         Top             =   2940
         Width           =   5685
         _ExtentX        =   10028
         _ExtentY        =   556
         _Version        =   393216
         Text            =   "Banco"
      End
      Begin VB.Label Label13 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Cuenta Contable de Anticipo"
         Height          =   330
         Left            =   105
         TabIndex        =   30
         Top             =   2625
         Width           =   5685
      End
      Begin VB.Label Label12 
         Caption         =   "USTED ESTA INGRESANDO ABONOS ANTICIPADOS, SE EMITIRA UN COMPROBANTE DE INGRESO DE RESPALDO A SU ABONO"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   540
         Left            =   105
         TabIndex        =   26
         Top             =   630
         Width           =   5685
      End
      Begin VB.Label Label11 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Cliente:"
         Height          =   330
         Left            =   105
         TabIndex        =   24
         Top             =   210
         Width           =   855
      End
      Begin VB.Label Label16 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Cuenta Contable del Ingreso"
         Height          =   330
         Left            =   105
         TabIndex        =   28
         Top             =   1890
         Width           =   5685
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3375
      Left            =   105
      TabIndex        =   4
      Top             =   525
      Width           =   5895
      Begin MSAdodcLib.Adodc AdoIngCaja 
         Height          =   330
         Left            =   105
         Top             =   1470
         Visible         =   0   'False
         Width           =   1800
         _ExtentX        =   3175
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
         Caption         =   "IngCaja"
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
      Begin MSDataListLib.DataCombo DCFactura 
         Bindings        =   "AbonoAnt.frx":0049
         DataSource      =   "AdoFactura"
         Height          =   315
         Left            =   3990
         TabIndex        =   8
         Top             =   210
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   556
         _Version        =   393216
         Text            =   "Factura"
      End
      Begin MSDataListLib.DataCombo DCTipo 
         Bindings        =   "AbonoAnt.frx":0062
         DataSource      =   "AdoDetAcomp"
         Height          =   360
         Left            =   735
         TabIndex        =   6
         Top             =   210
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   635
         _Version        =   393216
         Text            =   "DataCombo1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   105
         TabIndex        =   12
         Top             =   1050
         Width           =   5685
      End
      Begin VB.Label Label6 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Saldo Pendiente"
         Height          =   330
         Left            =   2310
         TabIndex        =   13
         Top             =   1470
         Width           =   1590
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " F&actura No."
         Height          =   330
         Left            =   2415
         TabIndex        =   7
         Top             =   210
         Width           =   1590
      End
      Begin VB.Label LabelSaldo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   3885
         TabIndex        =   14
         Top             =   1470
         Width           =   1905
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000000C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   3570
         TabIndex        =   11
         Top             =   630
         Width           =   2220
      End
      Begin VB.Label LblObs 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Observacion"
         ForeColor       =   &H00C000C0&
         Height          =   645
         Left            =   105
         TabIndex        =   15
         Top             =   1890
         Width           =   5685
      End
      Begin VB.Label LblNota 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nota"
         ForeColor       =   &H00C000C0&
         Height          =   540
         Left            =   105
         TabIndex        =   16
         Top             =   2625
         Width           =   5685
      End
      Begin VB.Label Label4 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " FECHA DE EMISION"
         Height          =   330
         Left            =   105
         TabIndex        =   9
         Top             =   630
         Width           =   2115
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   2205
         TabIndex        =   10
         Top             =   630
         Width           =   1275
      End
      Begin VB.Label Label9 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " TIPO"
         ForeColor       =   &H000000FF&
         Height          =   330
         Left            =   105
         TabIndex        =   5
         Top             =   210
         Width           =   645
      End
   End
   Begin VB.TextBox TxtRecibo 
      Alignment       =   1  'Right Justify
      Height          =   330
      Left            =   2415
      MaxLength       =   14
      TabIndex        =   1
      Text            =   "0"
      Top             =   105
      Width           =   1275
   End
   Begin VB.CheckBox CheqRecibo 
      Caption         =   "&RECIBO DE CAJA No."
      Height          =   330
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Value           =   1  'Checked
      Width           =   2220
   End
   Begin MSAdodcLib.Adodc AdoFactura 
      Height          =   330
      Left            =   105
      Top             =   3990
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
   Begin VB.TextBox TextCajaMN 
      Alignment       =   1  'Right Justify
      Height          =   330
      Left            =   3990
      MaxLength       =   14
      TabIndex        =   18
      Text            =   "0"
      Top             =   3990
      Width           =   1905
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Salir"
      Height          =   855
      Left            =   6090
      Picture         =   "AbonoAnt.frx":007C
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   1155
      Width           =   960
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   855
      Left            =   6090
      Picture         =   "AbonoAnt.frx":0946
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   105
      Width           =   960
   End
   Begin MSMask.MaskEdBox MBFecha 
      Height          =   330
      Left            =   4620
      TabIndex        =   3
      Top             =   105
      Width           =   1380
      _ExtentX        =   2434
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
   Begin MSAdodcLib.Adodc AdoDetAcomp 
      Height          =   330
      Left            =   105
      Top             =   4305
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
   Begin MSAdodcLib.Adodc AdoClientes 
      Height          =   330
      Left            =   105
      Top             =   4620
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
   Begin MSAdodcLib.Adodc AdoBanco 
      Height          =   330
      Left            =   105
      Top             =   4935
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
   Begin MSAdodcLib.Adodc AdoCtaAnt 
      Height          =   330
      Left            =   2415
      Top             =   4935
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
      Caption         =   "CtaAnt"
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
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &FECHA"
      Height          =   330
      Left            =   3780
      TabIndex        =   2
      Top             =   105
      Width           =   855
   End
   Begin VB.Label LabelPend 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   330
      Left            =   3990
      TabIndex        =   20
      Top             =   4410
      Width           =   1905
   End
   Begin VB.Label Label10 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Saldo Actual"
      Height          =   330
      Left            =   2415
      TabIndex        =   19
      Top             =   4410
      Width           =   1590
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Caja MN"
      Height          =   330
      Left            =   2415
      TabIndex        =   17
      Top             =   3990
      Width           =   1590
   End
End
Attribute VB_Name = "AbonoAnticipado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub Listar_Facturas_Pendientes()
  TipoFactura = DCTipo
  SQL1 = "SELECT F.TC,F.Factura,F.CodigoC,F.Fecha,F.Fecha_V,F.Saldo_MN,F.Cta_CxP,F.Nota," _
       & "F.Observacion,C.Cliente,C.Direccion,C.CI_RUC,C.Telefono,C.Grupo " _
       & "FROM Facturas As F,Clientes As C " _
       & "WHERE F.T = '" & Pendiente & "' " _
       & "AND F.Item = '" & NumEmpresa & "' " _
       & "AND F.TC = '" & TipoFactura & "' " _
       & "AND F.Periodo = '" & Periodo_Contable & "' "
  If TipoFactura = "OP" Then SQL1 = SQL1 & "AND Factura = " & FA.Factura & " "
  SQL1 = SQL1 _
       & "AND F.CodigoC = C.Codigo " _
       & "ORDER BY F.TC,F.Factura "
  SelectDB_Combo DCFactura, AdoFactura, SQL1, "Factura"
End Sub

Private Sub Command1_Click()
  Mensajes = "Esta Seguro que desea grabar Abono."
  Titulo = "Formulario de Grabación."
  If BoxMensaje = vbYes Then
     FechaValida MBFecha
     FechaTexto = MBFecha
     FechaComp = FechaTexto
     Total = Val(TotalCajaMN)
     DetalleComp = Ninguno
     CodigoCli = CodigoCliente
     If TipoFactura <> "OP" Then CodigoCli = Buscar_Beneficiario(DCClientes, X_Beneficiario)
     
     Trans_No = 200
     SubCtaGen = SinEspaciosIzq(DCCtaAnt)
     Eliminar_Asientos_SP True
     sSQL = "SELECT * " _
          & "FROM Asiento_SC " _
          & "WHERE TC = 'P' " _
          & "AND Cta = '" & SubCtaGen & "' " _
          & "AND DH = '2' " _
          & "AND TM = '1' " _
          & "AND T_No = " & Trans_No & " " _
          & "AND Item = '" & NumEmpresa & "' " _
          & "AND CodigoU = '" & CodigoUsuario & "' "
     Select_Adodc AdoIngCaja, sSQL
     With AdoIngCaja.Recordset
        .AddNew
        .fields("Prima") = 0
        .fields("Fecha_V") = MBFecha
        .fields("TC") = "P"
        .fields("Prima") = 0
        .fields("Serie") = "001001"
        .fields("Factura") = 0
        .fields("Codigo") = CodigoCliente
        .fields("Beneficiario") = NombreCliente
        .fields("Detalle_SubCta") = "Abono Anticipado"
        .fields("Cta") = SubCtaGen
        .fields("DH") = "2"
        .fields("Valor") = Total
        .fields("Valor_ME") = 0
        .fields("TM") = "1"
        .fields("Item") = NumEmpresa
        .fields("T_No") = Trans_No
        .fields("SC_No") = 1
        .fields("CodigoU") = CodigoUsuario
        .Update
     End With
     sSQL = "SELECT * " _
          & "FROM Asiento " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND CodigoU = '" & CodigoUsuario & "' " _
          & "AND T_No = " & Trans_No & " "
     Select_Adodc AdoIngCaja, sSQL
     If Frame2.Visible Then
        Cta_Aux = SinEspaciosIzq(DCBanco)
        If Len(Cta_Aux) <= 1 Then Cta_Aux = Cta_CajaG
     Else
        Cta_Aux = Cta_CajaG
     End If
     InsertarAsientos AdoIngCaja, Cta_Aux, 0, Total, 0
     InsertarAsientos AdoIngCaja, SubCtaGen, 0, 0, Total
     sSQL = "SELECT * " _
          & "FROM Catalogo_CxCxP " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "AND Codigo = '" & CodigoCliente & "' " _
          & "AND Cta = '" & SubCtaGen & "' " _
          & "AND TC = 'P' "
     Select_Adodc AdoIngCaja, sSQL
     With AdoIngCaja.Recordset
      If .RecordCount <= 0 Then
          SetAddNew AdoIngCaja
          SetFields AdoIngCaja, "Item", NumEmpresa
          SetFields AdoIngCaja, "Periodo", Periodo_Contable
          SetFields AdoIngCaja, "Codigo", CodigoCliente
          SetFields AdoIngCaja, "Cta", SubCtaGen
          SetFields AdoIngCaja, "TC", "P"
          SetUpdate AdoIngCaja
        End If
     End With
     
     RatonReloj
     NumComp = ReadSetDataNum("Ingresos", True, True)
     Co.TP = CompIngreso
     Co.T = Normal
     Co.Fecha = FechaTexto
     Co.Numero = NumComp
     Co.Monto_Total = Total
     If TipoFactura = "OP" Then
        Co.Concepto = "Abono Anticipado de: " & UCaseStrg(NombreCliente) & ", Orden No. " & FA.Factura
     Else
        Co.Concepto = "Abono Anticipado de: " & UCaseStrg(NombreCliente)
        If Len(Grupo) > 1 Then Co.Concepto = Co.Concepto & ", Grupo: " & Grupo
     End If
     If Len(TxtConcepto) > 1 Then Co.Concepto = Co.Concepto & ", " & TxtConcepto
     Co.CodigoB = CodigoCliente
     Co.Efectivo = Total
     Co.Cotizacion = 0
     Co.Item = NumEmpresa
     Co.Usuario = CodigoUsuario
     Co.T_No = Trans_No
     Grabar_Comprobante Co

   ' Seteamos para el siguiente comprobante
     RatonNormal
     Unload AbonoAnticipado
     Imprimir_Recibo_Anticipos Co, True
     
    'Procedemos a enviar por mail el recibo
     If Len(TMail.para) > 3 Then
        Titulo = "Formulario de envio por mail"
        Mensajes = "Enviar por mail el recibo"
        If BoxMensaje = vbYes Then
           TMail.Asunto = "RECIBO ABONO ANTICIPADO No. " & Format$(Year(Co.Fecha), "0000") & "-" & Co.TP & "-" & Format$(Co.Numero, "000000000")
           TMail.Adjunto = RutaDocumentoPDF
           TMail.Mensaje = "Beneficiario: " & Co.Beneficiario & vbCrLf _
                         & "Fecha del Abono: " & Co.Fecha & vbCrLf _
                         & "Abono Anticipado po USD " & Format$(Co.Efectivo, "#,##0.00") & vbCrLf
           FEnviarCorreos.Show 1
        End If
     End If
'    If TipoFactura <> "OP" Then ImprimirComprobantesDe False, Co
  End If
End Sub

Private Sub Command2_Click()
   Unload AbonoAnticipado
End Sub

Private Sub DCBanco_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCClientes_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCClientes_LostFocus()
    TMail.para = ""
    With AdoClientes.Recordset
     If .RecordCount > 0 Then
        .MoveFirst
        .Find ("Cliente = '" & DCClientes & "' ")
         If Not .EOF Then
            Grupo = .fields("Grupo")
            TMail.para = ""
            Insertar_Mail TMail.para, .fields("Email")
            Insertar_Mail TMail.para, .fields("Email2")
         End If
     End If
    End With
End Sub

Private Sub DCCtaAnt_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCCtaAnt_LostFocus()
   TextCajaMN.SetFocus
End Sub

Private Sub DCFactura_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCFactura_LostFocus()
  Codigo1 = Ninguno
  Saldo = 0
  Total_IVA = 0
  Total_Ret = 0
  TotalCajaMN = 0
  TotalCajaME = 0
  Total_Bancos = 0
  Total_Tarjeta = 0
  Cotizacion = 0
  TotalDolar = 0
  Saldo_ME = 0
  Label3.Caption = ""
  Label1.Caption = ""
  LabelPend.Caption = ""
  LabelSaldo.Caption = ""
  With AdoFactura.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Factura = " & Val(DCFactura.Text) & " ")
       If Not .EOF Then
          Grupo_No = .fields("Grupo")
          Label8.Caption = " " & .fields("Fecha")
          LblObs.Caption = " " & .fields("Observacion")
          LblNota.Caption = " " & .fields("Nota")
          CodigoCliente = .fields("CodigoC")
          NombreCliente = .fields("Cliente")
          DireccionCli = .fields("Direccion")
          Factura_No = .fields("Factura")
          Cta_Cobrar = .fields("Cta_CxP")
          TipoFactura = .fields("TC")
          Saldo = .fields("Saldo_MN")
          LabelSaldo.Caption = Format$(Saldo, "#,##0.00")
          Command1.Enabled = True
          Label3.Caption = NombreCliente
          Label1.Caption = " " & Factura_No
          SaldoDisp = Saldo - TotalCajaMN - TotalCajaME - Total_Bancos - Total_Tarjeta - Total_IVA - Total_Ret
          LabelPend.Caption = Format$(SaldoDisp, "#,##0.00")
          TextCajaMN.Text = LabelPend.Caption
          'AbonoEfectivo.Caption = "INGRESO DE CAJA (" & TipoFactura & ")"
          TextCajaMN.SetFocus
       Else
          MsgBox "Esta Factura no esta pendiente"
          Command1.Enabled = False
          DCFactura.SetFocus
       End If
    End If
  End With
End Sub

Private Sub DCTipo_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCTipo_LostFocus()
   Listar_Facturas_Pendientes
End Sub

Private Sub Form_Activate()
  ControlEsNumerico TextCajaMN
  SubCtaGen = Leer_Seteos_Ctas("Cta_Anticipos_Clientes")
  
  sSQL = "SELECT (Codigo & '  ' & Cuenta) As NomCuenta " _
       & "FROM Catalogo_Cuentas " _
       & "WHERE TC = 'P' " _
       & "AND DG = 'D' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "ORDER BY TC DESC,Codigo "
  SelectDB_Combo DCCtaAnt, AdoCtaAnt, sSQL, "NomCuenta"
  With AdoCtaAnt.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("NomCuenta LIKE '%" & SubCtaGen & "%' ")
       If Not .EOF Then DCCtaAnt.Text = .fields("NomCuenta")
   End If
  End With
  sSQL = "SELECT (Codigo & '  ' & Cuenta) As NomCuenta " _
       & "FROM Catalogo_Cuentas " _
       & "WHERE TC IN ('CJ','BA','TJ') " _
       & "AND DG = 'D' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "ORDER BY TC DESC,Codigo "
  SelectDB_Combo DCBanco, AdoBanco, sSQL, "NomCuenta"
  If TipoFactura = "OP" Then
     LabelPend.Visible = True
     Label10.Visible = True
     Frame1.Visible = True
     Frame2.Visible = False
     sSQL = "SELECT TC " _
          & "FROM Facturas " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "AND TC = 'OP' " _
          & "AND Factura = " & FA.Factura & " " _
          & "GROUP BY TC " _
          & "ORDER BY TC DESC "
     SelectDB_Combo DCTipo, AdoDetAcomp, sSQL, "TC"
  Else
     LabelPend.Visible = False
     Label10.Visible = False
     Frame1.Visible = False
     Frame2.Visible = True
     sSQL = "SELECT Grupo,Codigo,Cliente,Email, Email2 " _
          & "FROM Clientes " _
          & "WHERE FA <> '" & adFalse & "' " _
          & "ORDER BY Cliente "
     SelectDB_Combo DCClientes, AdoClientes, sSQL, "Cliente"
  End If
  DiarioCaja = ReadSetDataNum("Recibo_No", True, False)
  If CheqRecibo.value = 1 Then TxtRecibo = Format$(DiarioCaja, "0000000") Else TxtRecibo = ""
  Mifecha = BuscarFecha(FechaTexto)
  MBFecha.Text = FechaSistema
  If Bloquear_Control Then Command1.Enabled = False
  RatonNormal
End Sub

Private Sub Form_Load()
   CentrarForm AbonoAnticipado
   ConectarAdodc AdoBanco
   ConectarAdodc AdoCtaAnt
   ConectarAdodc AdoFactura
   ConectarAdodc AdoIngCaja
   ConectarAdodc AdoDetAcomp
   ConectarAdodc AdoClientes
End Sub

Private Sub TextCajaMN_GotFocus()
  TextCajaMN.Text = Saldo
  MarcarTexto TextCajaMN
End Sub

Private Sub TextCajaMN_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextCajaMN_LostFocus()
  TextoValido TextCajaMN, True
  TotalCajaMN = Redondear(Val(CCur(TextCajaMN.Text)), 2)
  TextCajaMN.Text = Format$(TotalCajaMN, "#,##0.00")
  SaldoDisp = Saldo - TotalCajaMN - TotalCajaME - Total_Bancos - Total_Tarjeta - Total_IVA - Total_Ret
  LabelPend.Caption = Format$(SaldoDisp, "#,##0.00")
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

Private Sub TxtConcepto_GotFocus()
  MarcarTexto TxtConcepto
End Sub

Private Sub TxtConcepto_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtConcepto_LostFocus()
  TextoValido TxtConcepto
End Sub

Private Sub TxtRecibo_GotFocus()
  MarcarTexto TxtRecibo
End Sub

Private Sub TxtRecibo_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

