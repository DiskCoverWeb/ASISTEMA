VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form CierrePuntoVenta 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CIERRE DE PUNTOS DE VENTA"
   ClientHeight    =   5400
   ClientLeft      =   5025
   ClientTop       =   4380
   ClientWidth     =   10620
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
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   10620
   ShowInTaskbar   =   0   'False
   Begin MSDataGridLib.DataGrid DGIngCaja 
      Bindings        =   "PuntoVen.frx":0000
      Height          =   2535
      Left            =   120
      TabIndex        =   15
      Top             =   600
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   4471
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
            LCID            =   2058
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
            LCID            =   2058
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
   Begin MSAdodcLib.Adodc AdoClien 
      Height          =   330
      Left            =   240
      Top             =   2520
      Visible         =   0   'False
      Width           =   2160
      _ExtentX        =   3810
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
      Caption         =   "Clien"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc AdoClientes 
      Height          =   330
      Left            =   240
      Top             =   2280
      Visible         =   0   'False
      Width           =   2160
      _ExtentX        =   3810
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
      Caption         =   "Clientes"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc AdoIngCaja 
      Height          =   330
      Left            =   240
      Top             =   2040
      Visible         =   0   'False
      Width           =   2160
      _ExtentX        =   3810
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc AdoDiarioCaja 
      Height          =   330
      Left            =   240
      Top             =   1800
      Visible         =   0   'False
      Width           =   2160
      _ExtentX        =   3810
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
      Caption         =   "DiarioCaja"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc AdoFactura 
      Height          =   330
      Left            =   240
      Top             =   1560
      Visible         =   0   'False
      Width           =   2160
      _ExtentX        =   3810
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc AdoListFact 
      Height          =   330
      Left            =   240
      Top             =   1320
      Visible         =   0   'False
      Width           =   2160
      _ExtentX        =   3810
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
      Caption         =   "ListFact"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc AdoTrans 
      Height          =   330
      Left            =   240
      Top             =   1080
      Visible         =   0   'False
      Width           =   2160
      _ExtentX        =   3810
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
      Caption         =   "Trans"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Imprimir Reporte"
      Height          =   330
      Left            =   7245
      TabIndex        =   3
      Top             =   105
      Width           =   1800
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Día anterior"
      Height          =   330
      Left            =   5460
      TabIndex        =   14
      Top             =   105
      Width           =   1695
   End
   Begin VB.CheckBox CheqTodos 
      Caption         =   "Facturar de &Todos"
      Height          =   330
      Left            =   3360
      TabIndex        =   2
      Top             =   105
      Width           =   2010
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Salir"
      Height          =   855
      Left            =   1050
      Picture         =   "PuntoVen.frx":0019
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4410
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Grabar Cierre"
      Height          =   855
      Left            =   105
      Picture         =   "PuntoVen.frx":029B
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4410
      Width           =   855
   End
   Begin MSMask.MaskEdBox MBoxFecha 
      Height          =   330
      Left            =   1995
      TabIndex        =   1
      Top             =   105
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
   Begin VB.Label LabelTotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   435
      Left            =   8295
      TabIndex        =   6
      Top             =   4830
      Width           =   2010
   End
   Begin VB.Label Label26 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Total Facturado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   6195
      TabIndex        =   8
      Top             =   4830
      Width           =   2115
   End
   Begin VB.Label LabelIVA 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   435
      Left            =   8295
      TabIndex        =   7
      Top             =   4410
      Width           =   2010
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " I.V.A. 12%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   6195
      TabIndex        =   9
      Top             =   4410
      Width           =   2115
   End
   Begin VB.Label LabelConIVA 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   435
      Left            =   4200
      TabIndex        =   10
      Top             =   4830
      Width           =   2010
   End
   Begin VB.Label LabelSubTotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   435
      Left            =   4200
      TabIndex        =   11
      Top             =   4410
      Width           =   2010
   End
   Begin VB.Label Label23 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Total Tarifa 12%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1995
      TabIndex        =   12
      Top             =   4830
      Width           =   2220
   End
   Begin VB.Label Label22 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Total Tarifa 0%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1995
      TabIndex        =   13
      Top             =   4410
      Width           =   2220
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &FECHA DE CIERRE"
      Height          =   330
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   1905
   End
End
Attribute VB_Name = "CierrePuntoVenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub CalculosTotales()
  Total_Factura = 0: Total_Servicio = 0
  Total_Con_IVA = 0: Total_Sin_IVA = 0
  Total_IVA = 0: Total_Desc = 0
  With AdoIngCaja.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       Do While Not .EOF
          'Total_IVA = Total_IVA + SetFields("IVA")
          'Total = SetFields("Valor_Total")
'          If SetFields("IVA") > 0 Then
'             Total_Con_IVA = Total_Con_IVA + Total
'          Else
'             Total_Sin_IVA = Total_Sin_IVA + Total
'          End If
         .MoveNext
       Loop
   End If
  End With
  Total_Con_IVA = Round(Total_Con_IVA)
  Total_Sin_IVA = Round(Total_Sin_IVA)
  Total_IVA = Round(Total_IVA)
  Total_Factura = Round(Total_Sin_IVA + Total_Con_IVA - Total_Desc + Total_IVA)
  LabelSubTotal.Caption = Format(Total_Sin_IVA, "#,##0.00")
  LabelConIVA.Caption = Format(Total_Con_IVA, "#,##0.00")
  LabelIVA.Caption = Format(Total_IVA, "#,##0.00")
  LabelTotal.Caption = Format(Total_Factura, "#,##0.00")
End Sub

Public Sub ProcGrabar()
 'Seteamos los encabezados para las facturas
  EsNotaVenta = False
  Moneda_US = False
  Total_FacturaME = 0
  CodigoCliente = "99999999"
  FechaTexto = MBoxFecha.Text
  CalculosTotales
  If AdoIngCaja.Recordset.RecordCount > 0 Then
     RatonReloj
     Factura_No = ReadSetDataNum("Cierre Caja", True, True)
     Factura_No = "99" & Format(Factura_No, "000000")
     SelectData AdoFactura, "Facturas"
     SelectData AdoListFact, "Detalle_Factura"
     If Moneda_US Then
        Total_Factura = Round((Total_Sin_IVA + Total_Con_IVA - Total_Desc + Total_IVA + Total_Servicio) * Dolar)
        Total_FacturaME = Round(Total_Sin_IVA + Total_Con_IVA - Total_Desc + Total_IVA + Total_Servicio)
     Else
        Total_Factura = Round(Total_Sin_IVA + Total_Con_IVA - Total_Desc + Total_IVA + Total_Servicio)
        Total_FacturaME = 0
     End If
     sSQL = "DELETE * FROM Detalle_Factura "
     sSQL = sSQL & "WHERE Factura_No = " & Factura_No & " "
     ConectarAdoExecute sSQL
     sSQL = "DELETE * FROM Facturas "
     sSQL = sSQL & "WHERE Factura = " & Factura_No & " "
     ConectarAdoExecute sSQL
     sSQL = "DELETE * FROM Diario_Caja "
     sSQL = sSQL & "WHERE Factura = " & Factura_No & " "
     ConectarAdoExecute sSQL
     SelectData AdoTrans, "Diario_Caja "
     With AdoTrans.Recordset
         SetAddNew AdoTrans
         SetFields AdoTrans, "T", Normal
         SetFields AdoTrans, "TP", ventas
         SetFields AdoTrans, "Fecha", FechaTexto
         SetFields AdoTrans, "Factura", Factura_No
         SetFields AdoTrans, "Monto_ME", Total_FacturaME
         SetFields AdoTrans, "Monto_MN", Total_Factura
         SetFields AdoTrans, "Diario_No", 0
         SetFields AdoTrans, "Caja_No", 0
         SetFields AdoTrans, "Caja_ME", 0
         SetFields AdoTrans, "Caja_MN", 0
         SetFields AdoTrans, "Caja_Vaucher", 0
         SetFields AdoTrans, "Abonos_ME", 0
         SetFields AdoTrans, "Abonos_MN", 0
         SetFields AdoTrans, "Saldo_MN", 0
         SetFields AdoTrans, "Saldo_ME", 0
         SetFields AdoTrans, "Codigo_C", CodigoCliente
         SetFields AdoTrans, "CtaxCob", Cta_Cobrar
         SetFields AdoTrans, "CtaxVent", Cta_Ventas
         SetFields AdoTrans, "Saldo_ME", Total_FacturaME
         SetFields AdoTrans, "Saldo_MN", Total_Factura
         SetFields AdoTrans, "Cotizacion", 0
         SetFields AdoTrans, "Banco", Ninguno
         SetFields AdoTrans, "Cheque", Ninguno
         SetUpdate AdoTrans
     End With
    'Grabamos el numero de factura
     With AdoFactura.Recordset
         SetAddNew AdoFactura
         SetFields AdoFactura, "T", Cancelado
         SetFields AdoFactura, "ME", Moneda_US
         SetFields AdoFactura, "Factura", Factura_No
         SetFields AdoFactura, "Fecha", FechaTexto
         SetFields AdoFactura, "Fecha_C", FechaTexto
         SetFields AdoFactura, "Fecha_V", FechaTexto
         SetFields AdoFactura, "Codigo_C", CodigoCliente
         SetFields AdoFactura, "Vendedor", NombreUsuario
         SetFields AdoFactura, "Pedido_No", 0
         SetFields AdoFactura, "Bultos_No", 0
         SetFields AdoFactura, "Gavetas_No", 0
         SetFields AdoFactura, "Forma_Pago", PagoCont
         SetFields AdoFactura, "Sin_IVA", Total_Sin_IVA
         SetFields AdoFactura, "Con_IVA", Total_Con_IVA
         SetFields AdoFactura, "SubTotal", Total_Sin_IVA + Total_Con_IVA
         SetFields AdoFactura, "Descuento", Total_Desc
         SetFields AdoFactura, "IVA", Total_IVA
         SetFields AdoFactura, "Servicio", Total_Servicio
         SetFields AdoFactura, "Total_MN", Total_Factura
         SetFields AdoFactura, "Total_ME", 0
         SetFields AdoFactura, "Comision", 0
         SetFields AdoFactura, "Saldo_MN", 0
         SetFields AdoFactura, "Saldo_ME", 0
         SetFields AdoFactura, "Cod_Ejec", Ninguno
         SetFields AdoFactura, "Porc_C", 0
         SetFields AdoFactura, "Cotizacion", Dolar
         SetFields AdoFactura, "Observacion", Ninguno
         SetFields AdoFactura, "Nota", Ninguno
         SetFields AdoFactura, "Cta_CxC", Cta_Cobrar
         SetFields AdoFactura, "Cta_Venta", Cta_Ventas
         SetFields AdoFactura, "Contrato_No", Ninguno
          Total = Total_Factura
         SetUpdate AdoFactura
     End With
     AdoIngCaja.Recordset.MoveFirst
     Do While Not AdoIngCaja.Recordset.EOF
        With AdoListFact.Recordset
            SetAddNew AdoListFact
            SetFields AdoListFact, "T", Cancelado
            SetFields AdoListFact, "Factura_No", Factura_No
            SetFields AdoListFact, "Codigo_C", CodigoCliente
            SetFields AdoListFact, "Fecha", FechaTexto
            SetFields AdoListFact, "Codigo", AdoIngCaja.Recordset.Fields("Codigo")
            SetFields AdoListFact, "Producto", AdoIngCaja.Recordset.Fields("Producto")
            SetFields AdoListFact, "Cantidad", AdoIngCaja.Recordset.Fields("Cant")
            SetFields AdoListFact, "Precio", AdoIngCaja.Recordset.Fields("Precio")
            SetFields AdoListFact, "Total", AdoIngCaja.Recordset.Fields("Valor_Total")
            SetFields AdoListFact, "Total_IVA", AdoIngCaja.Recordset.Fields("IVA")
            SetFields AdoListFact, "CodigoL", AdoIngCaja.Recordset.Fields("CodigoL")
            SetFields AdoListFact, "Reposicion", 0
            SetFields AdoListFact, "Total_Desc", 0
            SetFields AdoListFact, "Cod_Ejec", Ninguno
            SetFields AdoListFact, "Porc_C", 0
            SetUpdate AdoListFact
        End With
        AdoIngCaja.Recordset.MoveNext
     Loop
    'Grabamos el numero de factura
     PagoDelContado
  Else
     MsgBox "No se puede grabar la Factura," & Chr(13) & "falta datos."
  End If
End Sub

Private Sub Command1_Click()
  FechaValida MBoxFecha
  ProcGrabar
  CalcularPuntoVenta MBoxFecha, CheqTodos.Value
  'MBoxFecha.SetFocus
End Sub

Private Sub Command2_Click()
  Unload CierrePuntoVenta
End Sub

Private Sub Command3_Click()
  FechaValida MBoxFecha
  CalcularPuntoVenta MBoxFecha, CheqTodos.Value, True
  Command1.Enabled = False
End Sub

Private Sub Command4_Click()
  MensajeEncabData = "VENTAS DIARIAS DEL " & FechaStrg(MBoxFecha.Text)
  'ImprimirData AdoIngCaja, True, 1, 8
End Sub

Private Sub Form_Activate()
  Label23.Caption = " Total Tarifa " & Porc_IVA * 100 & "%"
  Label3.Caption = " Total I.V.A. " & Porc_IVA * 100 & "%"
  MBoxFecha.Text = FechaSistema
  FechaValida MBoxFecha
  CalcularPuntoVenta MBoxFecha, CheqTodos.Value
  'MBoxFecha.SetFocus
End Sub

Private Sub Form_Load()
   CentrarForm CierrePuntoVenta
   ConectarAdodc AdoClien
   ConectarAdodc AdoIngCaja
   ConectarAdodc AdoTrans
   ConectarAdodc AdoFactura
   ConectarAdodc AdoListFact
   ConectarAdodc AdoClientes
   ConectarAdodc AdoDiarioCaja
End Sub

Private Sub MBoxFecha_GotFocus()
 MarcarTexto MBoxFecha
End Sub

Private Sub MBoxFecha_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBoxFecha_LostFocus()
  FechaValida MBoxFecha
  CalcularPuntoVenta MBoxFecha, CheqTodos.Value
  Command1.Enabled = True
  'MBoxFecha.SetFocus
End Sub

Public Sub PagoDelContado()
  If AdoIngCaja.Recordset.RecordCount > 0 Then
     RatonReloj
     SelectData AdoDiarioCaja, "Diario_Caja"
     IngresoCaja = ReadSetDataNum("Ingreso Caja", True, True)
     DiarioCaja = 0
     Cheque = 0
     Efectivo = 0
     Cotizacion = 0
     Total = Total_Factura
     Retencion = Total_Factura
     AdoDiarioCaja.Recordset.AddNew
     AdoDiarioCaja.Recordset.Fields("T") = Normal
     AdoDiarioCaja.Recordset.Fields("TP") = CxC
     AdoDiarioCaja.Recordset.Fields("Fecha") = FechaTexto
     AdoDiarioCaja.Recordset.Fields("Diario_No") = DiarioCaja
     AdoDiarioCaja.Recordset.Fields("Caja_No") = IngresoCaja
     AdoDiarioCaja.Recordset.Fields("Factura") = Factura_No
     AdoDiarioCaja.Recordset.Fields("Monto_ME") = Total_FacturaME
     AdoDiarioCaja.Recordset.Fields("Monto_MN") = Total_Factura
     AdoDiarioCaja.Recordset.Fields("Caja_ME") = 0
     AdoDiarioCaja.Recordset.Fields("Caja_MN") = Total_Factura
     AdoDiarioCaja.Recordset.Fields("Caja_Vaucher") = 0
     AdoDiarioCaja.Recordset.Fields("Abonos_MN") = Total_Factura
     AdoDiarioCaja.Recordset.Fields("Saldo_MN") = 0
     AdoDiarioCaja.Recordset.Fields("Abonos_ME") = 0
     AdoDiarioCaja.Recordset.Fields("Saldo_ME") = 0
     AdoDiarioCaja.Recordset.Fields("Codigo_C") = CodigoCliente
     AdoDiarioCaja.Recordset.Fields("CtaxCob") = Cta_Cobrar
     AdoDiarioCaja.Recordset.Fields("CtaxVent") = Ninguno
     AdoDiarioCaja.Recordset.Fields("Cotizacion") = Cotizacion
     AdoDiarioCaja.Recordset.Fields("Banco") = Ninguno
     AdoDiarioCaja.Recordset.Fields("Cheque") = Ninguno
     AdoDiarioCaja.Recordset.Update
     sSQL = "UPDATE Detalle_Nota_Venta " _
          & "SET T = '" & Procesado & "' " _
          & "WHERE Fecha = #" & BuscarFecha(FechaTexto) & "# " _
          & "AND T = 'N' "
     If CheqTodos.Value <> 1 Then
        sSQL = sSQL & "AND Vendedor = '" & CodigoUsuario & "' "
     End If
     ConectarAdoExecute sSQL
     RatonNormal
     sSQL = "SELECT * FROM Diario_Caja " _
          & "WHERE Fecha = #" & FechaTexto & "# " _
          & "AND Codigo_C = '" & CodigoCliente & "' "
     SelectData AdoDiarioCaja, sSQL
     ImprimirReciboCaja AdoDiarioCaja, NombreUsuario
     Efectivo = 0: Retencion = 0: Cheque = 0
  End If
End Sub

Public Sub CalcularPuntoVenta(MBoxFechaPV As MaskEdBox, Todos As Byte, Optional DiaAnterior As Boolean)
  MiFecha = BuscarFecha(MBoxFechaPV.Text)
  sSQL = "SELECT * FROM Linea_Producto "
  SelectData AdoIngCaja, sSQL
  With AdoIngCaja.Recordset
   If .RecordCount > 0 Then
      'If SetFields("Cta_CxC") <> Ninguno Then Cta_Cobrar = SetFields("Cta_CxC")
   End If
  End With
  sSQL = "SELECT Producto,SUM(Cantidad) As Cant," _
       & "Precio,SUM(Total) As Valor_Total," _
       & "SUM(Total_IVA) As IVA,CodigoL,Codigo " _
       & "FROM Detalle_Nota_Venta " _
       & "WHERE Fecha = #" & MiFecha & "# "
  If DiaAnterior Then
     sSQL = sSQL & "AND T = 'P' "
  Else
     sSQL = sSQL & "AND T = 'N' "
  End If
  If Todos <> 1 Then sSQL = sSQL & "AND Vendedor = '" & CodigoUsuario & "' "
  sSQL = sSQL & "GROUP BY CodigoL,Codigo,Producto,Cantidad,Precio "
  SelectDataGrid DGIngCaja, AdoIngCaja, sSQL
  CalculosTotales
  RatonNormal
End Sub
