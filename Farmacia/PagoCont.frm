VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FPagoContado 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "INGRESO DE CAJA"
   ClientHeight    =   4860
   ClientLeft      =   5025
   ClientTop       =   4380
   ClientWidth     =   7575
   ClipControls    =   0   'False
   ControlBox      =   0   'False
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
   ScaleHeight     =   4860
   ScaleWidth      =   7575
   ShowInTaskbar   =   0   'False
   Begin MSDataGridLib.DataGrid DGIngCaja 
      Bindings        =   "PagoCont.frx":0000
      Height          =   1335
      Left            =   120
      TabIndex        =   30
      Top             =   3480
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   2355
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
   Begin MSAdodcLib.Adodc AdoDiarioCaja 
      Height          =   330
      Left            =   120
      Top             =   3600
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
   Begin MSAdodcLib.Adodc AdoIngCaja 
      Height          =   330
      Left            =   120
      Top             =   3960
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
   Begin MSAdodcLib.Adodc AdoFactura 
      Height          =   330
      Left            =   2280
      Top             =   4320
      Visible         =   0   'False
      Width           =   2280
      _ExtentX        =   4022
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
   Begin MSAdodcLib.Adodc AdoClientes 
      Height          =   330
      Left            =   2280
      Top             =   3960
      Visible         =   0   'False
      Width           =   2280
      _ExtentX        =   4022
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
   Begin MSAdodcLib.Adodc AdoListFact 
      Height          =   330
      Left            =   2280
      Top             =   3600
      Visible         =   0   'False
      Width           =   2280
      _ExtentX        =   4022
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
   Begin MSAdodcLib.Adodc AdoClien 
      Height          =   330
      Left            =   120
      Top             =   4320
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
   Begin VB.TextBox TextCheque 
      Alignment       =   1  'Right Justify
      Height          =   330
      Left            =   4200
      MaxLength       =   14
      MultiLine       =   -1  'True
      TabIndex        =   25
      Text            =   "PagoCont.frx":0019
      Top             =   2835
      Width           =   1800
   End
   Begin VB.TextBox TextRet 
      Alignment       =   1  'Right Justify
      Height          =   330
      Left            =   1155
      MaxLength       =   10
      MultiLine       =   -1  'True
      TabIndex        =   21
      Text            =   "PagoCont.frx":001D
      Top             =   2520
      Width           =   1485
   End
   Begin VB.TextBox TextCheqNo 
      Height          =   330
      Left            =   4725
      MaxLength       =   8
      TabIndex        =   19
      Text            =   "."
      Top             =   2205
      Width           =   1275
   End
   Begin VB.TextBox TextBanco 
      Height          =   330
      Left            =   945
      MaxLength       =   25
      TabIndex        =   17
      Text            =   "."
      Top             =   2205
      Width           =   2535
   End
   Begin VB.TextBox TextEfectivo 
      Alignment       =   1  'Right Justify
      Height          =   330
      Left            =   4200
      MaxLength       =   14
      MultiLine       =   -1  'True
      TabIndex        =   15
      Text            =   "PagoCont.frx":001F
      Top             =   1890
      Width           =   1800
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Factura a &Crédito"
      Height          =   1065
      Left            =   6195
      Picture         =   "PagoCont.frx":0023
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   2205
      Width           =   1170
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Grabar Pago"
      Height          =   1065
      Left            =   6195
      Picture         =   "PagoCont.frx":0ADD
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   1050
      Width           =   1170
   End
   Begin MSMask.MaskEdBox MBoxFecha 
      Height          =   330
      Left            =   1155
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
   Begin VB.Label LabelPend 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   330
      Left            =   4200
      TabIndex        =   27
      Top             =   3150
      Width           =   1800
   End
   Begin VB.Label Label10 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " SALDO"
      Height          =   330
      Left            =   105
      TabIndex        =   26
      Top             =   3150
      Width           =   4110
   End
   Begin VB.Label Label15 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Caja &Vaucher"
      Height          =   330
      Left            =   105
      TabIndex        =   24
      Top             =   2835
      Width           =   4110
   End
   Begin VB.Label LabelRet 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Left            =   4200
      TabIndex        =   23
      Top             =   2520
      Width           =   1800
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " MONEDA MN"
      Height          =   330
      Left            =   2625
      TabIndex        =   22
      Top             =   2520
      Width           =   1590
   End
   Begin VB.Label Label9 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Caja M&E"
      Height          =   330
      Left            =   105
      TabIndex        =   20
      Top             =   2520
      Width           =   1065
   End
   Begin VB.Label Label17 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " CHEQUE No."
      Height          =   330
      Left            =   3465
      TabIndex        =   18
      Top             =   2205
      Width           =   1275
   End
   Begin VB.Label Label16 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " BANCO"
      Height          =   330
      Left            =   105
      TabIndex        =   16
      Top             =   2205
      Width           =   855
   End
   Begin VB.Label LabelSaldo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Left            =   4200
      TabIndex        =   12
      Top             =   1155
      Width           =   1800
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " S/."
      Height          =   330
      Left            =   3465
      TabIndex        =   11
      Top             =   1155
      Width           =   750
   End
   Begin VB.Label LabelSaldoD 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Left            =   1890
      TabIndex        =   10
      Top             =   1155
      Width           =   1590
   End
   Begin VB.Label Label11 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " U.S.$"
      Height          =   330
      Left            =   1155
      TabIndex        =   9
      Top             =   1155
      Width           =   750
   End
   Begin VB.Label LabelDolares 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Left            =   1890
      TabIndex        =   8
      Top             =   840
      Width           =   1590
   End
   Begin VB.Label Label12 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Cotiz."
      Height          =   330
      Left            =   1155
      TabIndex        =   7
      Top             =   840
      Width           =   750
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Left            =   4200
      TabIndex        =   3
      Top             =   105
      Width           =   1800
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cliente"
      Height          =   330
      Left            =   1155
      TabIndex        =   5
      Top             =   525
      Width           =   4845
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Caja M&N"
      Height          =   330
      Left            =   105
      TabIndex        =   14
      Top             =   1890
      Width           =   4110
   End
   Begin VB.Label Label13 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " FORMA DE PAGO:"
      Height          =   330
      Left            =   105
      TabIndex        =   13
      Top             =   1575
      Width           =   5895
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Factura No."
      Height          =   330
      Left            =   2940
      TabIndex        =   2
      Top             =   105
      Width           =   1275
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " SALDO"
      Height          =   330
      Left            =   105
      TabIndex        =   6
      Top             =   840
      Width           =   1065
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &FECHA"
      Height          =   330
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   1065
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " CLIENTE"
      Height          =   330
      Left            =   105
      TabIndex        =   4
      Top             =   525
      Width           =   1065
   End
End
Attribute VB_Name = "FPagoContado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CodigoCliente As String
Dim NombreCliente As String

Private Sub Command1_Click()
  TextoValido TextBanco
  TextoValido TextCheqNo
  Mensajes = "Esta Seguro que desea grabar pago."
  Titulo = "Formulario de Grabación."
  If BoxMensaje = 6 Then
     FechaTexto = FechaSistema
     If AdoIngCaja.Recordset.RecordCount > 0 Then
        RatonReloj
        SelectData AdoDiarioCaja, "Diario_Caja", False
        IngresoCaja = ReadSetDataNum("Ingreso Caja", True, True)
        DiarioCaja = 0
        With AdoIngCaja.Recordset
            .MoveFirst
             Do While Not .EOF
                Cotizacion = .Fields("Cotizacion")
                Factura_No = .Fields("Factura")
                Total = .Fields("Valor_MN")
                Efectivo = .Fields("Caja_ME")
                Retencion = .Fields("Caja_MN")
                Cheque = .Fields("Caja_Vaucher")
                Saldo = Total - Efectivo - Retencion - Cheque
                If .Fields("Caja_ME") = 0 Then Cotizacion = 0
                If Saldo < 0 Then Saldo = 0
                AdoDiarioCaja.Recordset.AddNew
                AdoDiarioCaja.Recordset.Fields("T") = Normal
                AdoDiarioCaja.Recordset.Fields("TP") = CxC
                AdoDiarioCaja.Recordset.Fields("Fecha") = MBoxFecha.Text
                AdoDiarioCaja.Recordset.Fields("Diario_No") = DiarioCaja
                AdoDiarioCaja.Recordset.Fields("Caja_No") = IngresoCaja
                AdoDiarioCaja.Recordset.Fields("Factura") = Factura_No
                AdoDiarioCaja.Recordset.Fields("Monto_ME") = .Fields("Valor_ME")
                AdoDiarioCaja.Recordset.Fields("Monto_MN") = .Fields("Valor_MN")
                AdoDiarioCaja.Recordset.Fields("Caja_ME") = .Fields("Caja_ME")
                AdoDiarioCaja.Recordset.Fields("Caja_MN") = .Fields("Caja_MN")
                AdoDiarioCaja.Recordset.Fields("Caja_Vaucher") = .Fields("Caja_Vaucher")
                AdoDiarioCaja.Recordset.Fields("Abonos_MN") = .Fields("Total_Abono")
                AdoDiarioCaja.Recordset.Fields("Saldo_MN") = .Fields("Saldo")
                AdoDiarioCaja.Recordset.Fields("Abonos_ME") = 0
                AdoDiarioCaja.Recordset.Fields("Saldo_ME") = 0
                If Cotizacion > 0 Then
                   AdoDiarioCaja.Recordset.Fields("Saldo_ME") = .Fields("Saldo") / Cotizacion
                   AdoDiarioCaja.Recordset.Fields("Abonos_ME") = .Fields("Total_Abono") / Cotizacion
                End If
                AdoDiarioCaja.Recordset.Fields("Codigo_C") = Codigo
                AdoDiarioCaja.Recordset.Fields("CtaxCob") = .Fields("CtaxCob")
                AdoDiarioCaja.Recordset.Fields("CtaxVent") = Ninguno
                AdoDiarioCaja.Recordset.Fields("Cotizacion") = Cotizacion
                AdoDiarioCaja.Recordset.Fields("Banco") = .Fields("Banco")
                AdoDiarioCaja.Recordset.Fields("Cheque") = .Fields("Cheque")
                AdoDiarioCaja.Recordset.Update
                If Saldo > 0 Then
                   TextoFormaPago = PagoCred
                   T = Pendiente
                Else
                   If Saldo < 0 Then Saldo = 0
                   TextoFormaPago = PagoCont
                   T = Cancelado
                End If
                sSQL = "UPDATE Facturas SET Saldo_MN = " & Saldo & " "
                If Cotizacion > 0 Then sSQL = sSQL & ", Saldo_ME = " & Saldo / Cotizacion & " "
                sSQL = sSQL & ", T = '" & T & "' "
                sSQL = sSQL & ", Forma_Pago = '" & TextoFormaPago & "' "
                sSQL = sSQL & ", Fecha_C = #" & BuscarFecha(MBoxFecha.Text) & "# "
                sSQL = sSQL & "WHERE Factura = " & Factura_No & " "
                ConectarAdoExecute sSQL
               .MoveNext
             Loop
        End With
        sSQL = "DELETE * FROM Ingreso_Caja "
        AdoIngCaja.Database.Execute sSQL
        sSQL = "SELECT * FROM Ingreso_Caja "
        SelectDataGrid DGIngCaja, AdoIngCaja, sSQL
        Total = Efectivo + Retencion + Cheque
        sSQL = "SELECT Factura,Monto_ME,Monto_MN,Caja_ME,Caja_MN,Caja_Vaucher,Abonos_ME,Abonos_MN,"
        sSQL = sSQL & "Saldo_ME,Saldo_MN,Fecha,Diario_No,Caja_No,CtaxCob,CtaxVent,Cotizacion "
        sSQL = sSQL & "FROM Diario_Caja "
        sSQL = sSQL & "WHERE Fecha = #" & BuscarFecha(MBoxFecha.Text) & "# "
        sSQL = sSQL & "AND T = '" & Normal & "' "
        sSQL = sSQL & "AND Codigo_C = '" & CodigoCliente & "' "
        sSQL = sSQL & "AND TP = '" & CxC & "' "
        sSQL = sSQL & "ORDER BY Factura "
        SelectData AdoDiarioCaja, sSQL, False
        With AdoDiarioCaja.Recordset
        Total = 0
        Do While Not .EOF
           Total = Total + .Fields("Abonos_MN")
          .MoveNext
        Loop
        End With
        RatonNormal
        If Saldo > 0 Then ImprimirReciboCaja AdoDiarioCaja, NombreCliente
     End If
     Unload FPagoContado
  End If
End Sub

Private Sub Command2_Click()
   Unload FPagoContado
End Sub

Private Sub Form_Activate()
   NombreCliente = Cadena
   Codigo = CodigoCli
   CodigoCliente = Codigo
   Efectivo = 0: Retencion = 0: Cheque = 0
   MiFecha = BuscarFecha(FechaSistema)
   Label1.Caption = " " & Codigo & "  " & NombreCliente
   Label14.Caption = Factura_No
   sSQL = "DELETE * FROM Ingreso_Caja "
   ConectarAdoExecute sSQL
   sSQL = "SELECT * FROM Ingreso_Caja "
   SelectDataGrid DGIngCaja, AdoIngCaja, sSQL
   RatonNormal
   MBoxFecha.SetFocus
End Sub

Private Sub Form_Load()
   CentrarForm FPagoContado
   ConectarAdodc AdoFactura
   ConectarAdodc AdoListFact
   ConectarAdodc AdoIngCaja
   ConectarAdodc AdoDiarioCaja
   ConectarAdodc AdoClientes
   ConectarAdodc AdoClien
End Sub

Private Sub MBoxFecha_GotFocus()
  MarcarTexto MBoxFecha
End Sub

Private Sub MBoxFecha_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBoxFecha_LostFocus()
  FechaValida MBoxFecha, False
  TextEfectivo.Enabled = False
  TextCheque.Enabled = False
  TextRet.Enabled = False
  Saldo = 0: Cotizacion = 0
  TotalDolar = 0: Saldo_ME = 0
  sSQL = "SELECT * FROM Facturas "
  sSQL = sSQL & "WHERE Factura = " & Factura_No & " "
  sSQL = sSQL & "AND Codigo_C = '" & Codigo & "' "
  sSQL = sSQL & "AND T = '" & Pendiente & "' "
  SelectData AdoListFact, sSQL, False
  With AdoListFact.Recordset
   If .RecordCount > 0 Then
       Cta_Cobrar = .Fields("Cta_CxC")
       Saldo = Round(.Fields("Saldo_MN"))
       Saldo_ME = Round(.Fields("Saldo_ME"))
       TotalDolar = .Fields("Total_ME")
       Cotizacion = .Fields("Cotizacion")
       Moneda_US = .Fields("ME")
       Efectivo = 0: Cheque = 0: Retencion = 0
       LabelDolares.Caption = Format(Cotizacion, "#,##0.00")
       If TotalDolar <> 0 Then LabelSaldoD.Caption = Format(Saldo_ME, "#,##0.00")
       LabelSaldo.Caption = Format(Saldo, "#,##0.00")
       If Moneda_US Then
          'Saldo_ME = Saldo / Cotizacion
          TextRet.Enabled = True
          TextRet.Text = Format(Saldo_ME, "#,##0.00")
          TextRet.SetFocus
       Else
          'Saldo = Round(Saldo)
          TextEfectivo.Enabled = True
          TextCheque.Enabled = True
          TextEfectivo.Text = Format(Saldo, "#,##0.00")
          TextEfectivo.SetFocus
       End If
   End If
  End With
End Sub

Private Sub TextEfectivo_GotFocus()
  TextEfectivo.Text = ""
  Efectivo = 0
End Sub

Private Sub TextEfectivo_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then TextEfectivo.Text = Saldo
End Sub

Private Sub TextEfectivo_LostFocus()
   TextoValido TextEfectivo, True
   Efectivo = Round(Val(TextEfectivo.Text))
   TextEfectivo.Text = Format(Efectivo, "#,##0.00")
   LabelPend.Caption = Format(Saldo - Efectivo - Retencion - Cheque, "#,##0.00")
End Sub

Private Sub TextCheque_GotFocus()
   TextCheque.Text = ""
   Cheque = 0
End Sub

Private Sub TextCheque_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextCheque_LostFocus()
  'MsgBox Cta_Cobrar
  TextoValido TextBanco
  TextoValido TextCheqNo
  TextoValido TextCheque, True
  Cheque = Round(Val(TextCheque.Text))
  SaldoFinal = Round(Efectivo + Retencion + Cheque)
  LabelPend.Caption = Format(Saldo - SaldoFinal, "#,##0.00")
  If Factura_No > 0 And (SaldoFinal <> 0) Then
     With AdoIngCaja.Recordset
         .AddNew
         .Fields("Factura") = Factura_No
         .Fields("Valor_MN") = Saldo
         .Fields("Valor_ME") = Saldo_ME
         .Fields("Caja_MN") = Efectivo
         .Fields("Caja_ME") = Retencion
         .Fields("Caja_Vaucher") = Cheque
         .Fields("Total_Abono") = SaldoFinal
         .Fields("Saldo") = Round(Saldo - SaldoFinal)
         .Fields("CtaxCob") = Cta_Cobrar 'Cta
         .Fields("Cotizacion") = Cotizacion
         .Fields("Banco") = TextBanco.Text
         .Fields("Cheque") = TextCheqNo.Text
         .Update
     End With
     'DBCFactura.SetFocus
  End If
End Sub

Private Sub TextRet_GotFocus()
  TextRet.Text = ""
End Sub

Private Sub TextRet_LostFocus()
  TextoValido TextRet, True
  Retencion = Round(Val(TextRet.Text) * Cotizacion)
  SaldoFinal = Round(Efectivo + Retencion + Cheque)
  LabelRet.Caption = Format(Retencion, "#,##0.00")
  LabelPend.Caption = Format(Saldo - SaldoFinal, "#,##0.00")
  LabelPend.Caption = Format(Saldo - SaldoFinal, "#,##0.00")
  If Factura_No > 0 And (SaldoFinal <> 0) Then
     With AdoIngCaja.Recordset
         .AddNew
         .Fields("Factura") = Factura_No
         .Fields("Valor_MN") = Saldo
         .Fields("Valor_ME") = Saldo_ME
         .Fields("Caja_MN") = Efectivo
         .Fields("Caja_ME") = Retencion
         .Fields("Caja_Vaucher") = Cheque
         .Fields("Total_Abono") = SaldoFinal
         .Fields("Saldo") = Saldo - SaldoFinal
         .Fields("CtaxCob") = Cta_Cobrar 'Cta
         .Fields("Cotizacion") = Cotizacion
         .Update
     End With
     'DBCFactura.SetFocus
  End If
End Sub

Private Sub TextBanco_GotFocus()
  MarcarTexto TextBanco
End Sub

Private Sub TextBanco_LostFocus()
  TextoValido TextBanco
End Sub

Private Sub TextCheqNo_GotFocus()
  MarcarTexto TextCheqNo
End Sub

Private Sub TextCheqNo_LostFocus()
  TextoValido TextCheqNo
End Sub

