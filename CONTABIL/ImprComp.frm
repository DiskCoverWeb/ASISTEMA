VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Begin VB.Form ImprimirComprobantes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "IMPRESION DE COMPROBANTES"
   ClientHeight    =   6555
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7995
   Icon            =   "ImprComp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   7995
   ShowInTaskbar   =   0   'False
   Begin MSDataGridLib.DataGrid DGListComp 
      Bindings        =   "ImprComp.frx":08CA
      Height          =   3690
      Left            =   105
      TabIndex        =   10
      Top             =   1260
      Width           =   7785
      _ExtentX        =   13732
      _ExtentY        =   6509
      _Version        =   393216
      AllowUpdate     =   0   'False
      BorderStyle     =   0
      HeadLines       =   1
      RowHeight       =   15
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
      Caption         =   "ListComp"
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
   Begin MSAdodcLib.Adodc AdoDesde 
      Height          =   330
      Left            =   315
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
      Caption         =   "Desde"
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
   Begin MSDataListLib.DataCombo DCDesde 
      Bindings        =   "ImprComp.frx":08E4
      DataSource      =   "AdoDesde"
      Height          =   345
      Left            =   1050
      TabIndex        =   5
      Top             =   840
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   609
      _Version        =   393216
      Text            =   "0"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Consultar"
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
      Left            =   4515
      Picture         =   "ImprComp.frx":08FB
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   105
      Width           =   1065
   End
   Begin MSAdodcLib.Adodc AdoListComp 
      Height          =   330
      Left            =   105
      Top             =   6090
      Width           =   7785
      _ExtentX        =   13732
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
      Caption         =   "ListComp"
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
      Caption         =   "&Imprimir"
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
      Left            =   5670
      Picture         =   "ImprComp.frx":0D3D
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   105
      Width           =   1065
   End
   Begin VB.CommandButton Command3 
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
      Left            =   6825
      Picture         =   "ImprComp.frx":1607
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   105
      Width           =   1065
   End
   Begin VB.ListBox LstComp 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   855
   End
   Begin MSMask.MaskEdBox MBFechaF 
      Height          =   330
      Left            =   2520
      TabIndex        =   4
      Top             =   420
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
   Begin MSMask.MaskEdBox MBFechaI 
      Height          =   330
      Left            =   1050
      TabIndex        =   2
      Top             =   420
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
   Begin MSDataListLib.DataCombo DCHasta 
      Bindings        =   "ImprComp.frx":1FFD
      DataSource      =   "AdoHasta"
      Height          =   345
      Left            =   2520
      TabIndex        =   6
      Top             =   840
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   609
      _Version        =   393216
      Text            =   "0"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSAdodcLib.Adodc AdoHasta 
      Height          =   330
      Left            =   315
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
      Caption         =   "Hasta"
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
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Hasta"
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
      Left            =   2520
      TabIndex        =   3
      Top             =   105
      Width           =   1380
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Desde"
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
      Left            =   1050
      TabIndex        =   1
      Top             =   105
      Width           =   1380
   End
End
Attribute VB_Name = "ImprimirComprobantes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub ImprimirListaDeComprobantes(ImpSoloReten As Boolean)
Dim AdoComp As ADODB.Recordset
Dim AdoTrans As ADODB.Recordset
Dim AdoBanc As ADODB.Recordset
Dim AdoFact As ADODB.Recordset
Dim AdoRet As ADODB.Recordset
Dim AdoSubC1 As ADODB.Recordset
Dim AdoSubC2 As ADODB.Recordset
Dim IdxComp As Long
'
 Select Case TipoComp
   Case CompIngreso: Mensajes = "Imprimir Comprobante de Ingreso" & vbCrLf
   Case CompEgreso: Mensajes = "Imprimir Comprobante de Egreso" & vbCrLf
   Case CompDiario: Mensajes = "Imprimir Comprobante de Diario" & vbCrLf
   Case CompNotaDebito: Mensajes = "Imprimir Comprobante de Nota de Debito" & vbCrLf
   Case CompNotaDebito: Mensajes = "Imprimir Comprobante de Nota de Credito" & vbCrLf
 End Select
 Mensajes = Mensajes & "Desde el " & DCDesde & " Hasta " & DCHasta & " en:" & vbCrLf _
          & Printer.DeviceName & "?"
 Titulo = "IMPRESION"
 Bandera = False
 SetPrinters.Show 1
 If PonImpresoraDefecto(SetNombrePRN) Then
    Escala_Centimetro 1, TipoTimes, 10
   'Listar el Comprobante
    Do While Not AdoListComp.Recordset.EOF
       ConceptoComp = Ninguno
       Co.Item = NumEmpresa
       Co.TP = AdoListComp.Recordset.fields("TP")
       Co.Numero = AdoListComp.Recordset.fields("Numero")
       Co.Fecha = AdoListComp.Recordset.fields("Fecha")
       ImprimirComprobantes.Caption = "Imprimiendo Comprobante de " & TipoComp & " No. " & Co.Numero
       
    'Listar el Comprobante
     sSQL = "SELECT C.*,A.Nombre_Completo,Cl.CI_RUC,Cl.Direccion,Cl.Email," _
          & "Cl.Telefono,Cl.Celular,Cl.FAX,Cl.Cliente,Cl.Codigo,Cl.Ciudad " _
          & "FROM Comprobantes As C,Accesos As A,Clientes As Cl " _
          & "WHERE C.Numero = " & Co.Numero & " " _
          & "AND C.TP = '" & Co.TP & "' " _
          & "AND C.Item = '" & Co.Item & "' " _
          & "AND C.Periodo = '" & Periodo_Contable & "' " _
          & "AND C.CodigoU = A.Codigo " _
          & "AND C.Codigo_B = Cl.Codigo "
     Select_AdoDB AdoComp, sSQL
     If AdoComp.RecordCount > 0 Then
        Co.Fecha = AdoComp.fields("Fecha")
        Co.Concepto = AdoComp.fields("Concepto")
        ConceptoComp = Co.Concepto
     End If
   
     'Listar las Transacciones
      sSQL = "SELECT T.Cta,Ca.Cuenta,Parcial_ME,Debe,Haber,Detalle,Cheq_Dep,Fecha_Efec,Ca.Item " _
           & "FROM Transacciones As T,Catalogo_Cuentas As Ca " _
           & "WHERE T.TP = '" & Co.TP & "' " _
           & "AND T.Numero = " & Co.Numero & " " _
           & "AND T.Item = '" & Co.Item & "' " _
           & "AND T.Periodo = '" & Periodo_Contable & "' " _
           & "AND T.Item = Ca.Item " _
           & "AND T.Cta = Ca.Codigo " _
           & "AND T.Periodo = Ca.Periodo " _
           & "ORDER BY T.ID,Debe DESC,T.Cta "
      Select_AdoDB AdoTrans, sSQL
     'Llenar Bancos
      sSQL = "SELECT T.Cta,C.TC,C.Cuenta,Co.Fecha,Cl.Cliente,T.Cheq_Dep,T.Debe,T.Haber " _
           & "FROM Transacciones As T,Comprobantes As Co,Catalogo_Cuentas As C,Clientes As Cl " _
           & "WHERE T.TP = '" & Co.TP & "' " _
           & "AND T.Numero = -1 " _
           & "AND T.Item = '" & Co.Item & "' " _
           & "AND T.Periodo = '" & Periodo_Contable & "' " _
           & "AND T.Numero = Co.Numero " _
           & "AND T.TP = Co.TP " _
           & "AND T.Cta = C.Codigo " _
           & "AND T.Item = C.Item " _
           & "AND T.Item = Co.Item " _
           & "AND T.Periodo = C.Periodo " _
           & "AND T.Periodo = Co.Periodo " _
           & "AND C.TC = 'BA' " _
           & "AND Co.Codigo_B = Cl.Codigo "
      Select_AdoDB AdoBanc, sSQL
     'Listar las Retenciones del IVA
      sSQL = "SELECT * " _
           & "FROM Trans_Compras " _
           & "WHERE Numero = " & Co.Numero & " " _
           & "AND TP = '" & Co.TP & "' " _
           & "AND Item = '" & Co.Item & "' " _
           & "AND Periodo = '" & Periodo_Contable & "' " _
           & "ORDER BY Cta_Servicio,Cta_Bienes "
     Select_AdoDB AdoFact, sSQL
    'Listar las Retenciones de la Fuente
     sSQL = "SELECT R.*,TIV.Concepto " _
           & "FROM Trans_Air As R,Tipo_Concepto_Retencion As TIV " _
           & "WHERE R.Numero = " & Co.Numero & " " _
           & "AND R.TP = '" & Co.TP & "' " _
           & "AND R.Item = '" & Co.Item & "' " _
           & "AND TIV.Fecha_Inicio <= #" & BuscarFecha(Co.Fecha) & "# " _
           & "AND TIV.Fecha_Final >= #" & BuscarFecha(Co.Fecha) & "# " _
           & "AND R.Periodo = '" & Periodo_Contable & "' " _
           & "AND R.Tipo_Trans IN ('C','I') " _
           & "AND R.CodRet = TIV.Codigo " _
           & "ORDER BY R.Cta_Retencion "
      Select_AdoDB AdoRet, sSQL
     'Llenar SubCtas
      sSQL = "SELECT T.Cta,T.TC,T.Factura,C.Cliente,T.Detalle_SubCta,T.Debitos,T.Creditos,T.Fecha_V,T.Codigo,T.Prima " _
           & "FROM Trans_SubCtas As T,Clientes As C " _
           & "WHERE T.TP = '" & Co.TP & "' " _
           & "AND T.Numero = " & Co.Numero & " " _
           & "AND T.Item = '" & Co.Item & "' " _
           & "AND T.Periodo = '" & Periodo_Contable & "' " _
           & "AND T.TC IN ('C','P') " _
           & "AND T.Codigo = C.Codigo " _
           & "ORDER BY T.Cta,C.Cliente,T.Fecha_V,T.Factura "
      Select_AdoDB AdoSubC1, sSQL
      sSQL = "SELECT T.Cta,T.TC,T.Factura,C.Detalle As Cliente,T.Detalle_SubCta,T.Debitos,T.Creditos,T.Fecha_V,T.Codigo,T.Prima " _
           & "FROM Trans_SubCtas As T,Catalogo_SubCtas As C " _
           & "WHERE T.TP = '" & Co.TP & "' " _
           & "AND T.Numero = " & Co.Numero & " " _
           & "AND T.Item = '" & Co.Item & "' " _
           & "AND T.Periodo = '" & Periodo_Contable & "' " _
           & "AND T.TC = C.TC " _
           & "AND T.Item = C.Item " _
           & "AND T.Periodo = C.Periodo " _
           & "AND T.Codigo = C.Codigo " _
           & "ORDER BY T.Cta,C.Detalle,T.Fecha_V,T.Factura "
      Select_AdoDB AdoSubC2, sSQL
        
       Select Case Co.TP
         Case CompIngreso: ImprimirCompIngreso AdoComp, AdoBanc, AdoTrans, AdoSubC1, AdoSubC2, True
         Case CompEgreso: ImprimirCompEgreso AdoComp, AdoBanc, AdoTrans, AdoFact, AdoRet, AdoSubC1, AdoSubC2, ImpSoloReten, True, True
         Case CompDiario: ImprimirCompDiario AdoComp, AdoTrans, AdoFact, AdoRet, AdoSubC1, AdoSubC2, ImpSoloReten, True, True
         Case CompNotaDebito: ImprimirCompNota_D_C AdoComp, AdoTrans, AdoSubC1, AdoSubC2, "ND", True
         Case CompNotaCredito: ImprimirCompNota_D_C AdoComp, AdoTrans, AdoSubC1, AdoSubC2, "NC", True
       End Select
      'Printer.NewPage
       AdoListComp.Recordset.MoveNext
    Loop
    Printer.EndDoc
    AdoComp.Close
    AdoTrans.Close
    AdoBanc.Close
    AdoRet.Close
    MsgBox "Proceso Terminado"
 End If
End Sub

Private Sub Command1_Click()
  FechaValida MBFechaI
  FechaValida MBFechaF
  FechaIni = BuscarFecha(MBFechaI)
  FechaFin = BuscarFecha(MBFechaF)
  TipoComp = LstComp.Text
  No_Desde = Val(DCDesde)
  No_Hasta = Val(DCHasta)
  sSQL = "SELECT Co.TP,Co.Numero,Co.Fecha,C.Cliente,Co.CodigoU " _
       & "FROM Comprobantes As Co, Clientes As C " _
       & "WHERE Co.Item = '" & NumEmpresa & "' " _
       & "AND Co.Periodo = '" & Periodo_Contable & "' " _
       & "AND Co.TP = '" & LstComp.Text & "' " _
       & "AND Co.Numero BETWEEN " & No_Desde & " and " & No_Hasta & " " _
       & "AND Co.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
       & "AND Co.Codigo_B = C.Codigo " _
       & "ORDER BY Co.TP,Co.Numero "
  Select_Adodc_Grid DGListComp, AdoListComp, sSQL
End Sub

Private Sub Command2_Click()
  If AdoListComp.Recordset.RecordCount > 0 Then
     Control_Procesos "I", "Imprimio Comprobantes de " & LstComp.Text _
                    & ", desde el No. " & DCDesde & " al " & DCHasta
     DGListComp.Visible = False
     ImprimirListaDeComprobantes False
     DGListComp.Visible = True
     Unload ImprimirComprobantes
  End If
End Sub

Private Sub Command3_Click()
  Unload Me
End Sub

Private Sub DCDesde_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCHasta_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub Form_Activate()
  LstComp.Clear
  LstComp.AddItem "CD"
  LstComp.AddItem "CI"
  LstComp.AddItem "CE"
  LstComp.AddItem "ND"
  LstComp.AddItem "NC"
  LstComp.Text = "CD"
  MBFechaI = FechaSistema
  MBFechaF = FechaSistema
  RatonNormal
End Sub

Private Sub Form_Load()
  CentrarForm ImprimirComprobantes
  ConectarAdodc AdoDesde
  ConectarAdodc AdoHasta
  ConectarAdodc AdoListComp
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
   FechaIni = BuscarFecha(MBFechaI)
   FechaFin = BuscarFecha(MBFechaF)
   sSQL = "SELECT Numero " _
        & "FROM Comprobantes " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND Periodo = '" & Periodo_Contable & "' " _
        & "AND TP = '" & LstComp.Text & "' " _
        & "AND Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
        & "ORDER BY Numero "
   SelectDB_Combo DCDesde, AdoDesde, sSQL, "Numero"
   SelectDB_Combo DCHasta, AdoHasta, sSQL, "Numero"
   If AdoHasta.Recordset.RecordCount > 0 Then
      AdoHasta.Recordset.MoveLast
      DCHasta.Text = AdoHasta.Recordset.fields("Numero")
   End If
End Sub

