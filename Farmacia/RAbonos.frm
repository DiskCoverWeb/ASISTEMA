VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Begin VB.Form RAbonos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "INGRESO DE CAJA                                                                      Cancelacion Factura"
   ClientHeight    =   6000
   ClientLeft      =   5025
   ClientTop       =   4380
   ClientWidth     =   11175
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
   ScaleHeight     =   6000
   ScaleWidth      =   11175
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "&Recibo"
      Height          =   750
      Left            =   8820
      Picture         =   "RAbonos.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   105
      Width           =   960
   End
   Begin VB.OptionButton OpcCliente 
      Caption         =   "POR CLIENTE"
      Height          =   330
      Left            =   3570
      TabIndex        =   2
      Top             =   105
      Width           =   1590
   End
   Begin VB.OptionButton OpcRecibo 
      Caption         =   "POR RECIBO"
      Height          =   330
      Left            =   1890
      TabIndex        =   1
      Top             =   105
      Width           =   1590
   End
   Begin VB.OptionButton OpcFactura 
      Caption         =   "POR FACTURA"
      Height          =   330
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Value           =   -1  'True
      Width           =   1695
   End
   Begin MSDataGridLib.DataGrid DGIngCaja 
      Bindings        =   "RAbonos.frx":08CA
      Height          =   2535
      Left            =   105
      TabIndex        =   6
      Top             =   1995
      Width           =   9675
      _ExtentX        =   17066
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
   Begin MSAdodcLib.Adodc AdoClientes 
      Height          =   330
      Left            =   315
      Top             =   3045
      Visible         =   0   'False
      Width           =   2430
      _ExtentX        =   4286
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
   Begin MSAdodcLib.Adodc AdoDiarioCaja 
      Height          =   330
      Left            =   315
      Top             =   2625
      Visible         =   0   'False
      Width           =   2430
      _ExtentX        =   4286
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
   Begin VB.CommandButton Command3 
      Caption         =   "&Imprimir"
      Height          =   750
      Left            =   7770
      Picture         =   "RAbonos.frx":08E6
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   105
      Width           =   960
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Salir"
      Height          =   750
      Left            =   9870
      Picture         =   "RAbonos.frx":11B0
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   105
      Width           =   960
   End
   Begin MSMask.MaskEdBox MBFecha 
      Height          =   330
      Left            =   6405
      TabIndex        =   4
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
   Begin MSDataListLib.DataCombo DCRecibo 
      Bindings        =   "RAbonos.frx":1A7A
      DataSource      =   "AdoRecibo"
      Height          =   315
      Left            =   105
      TabIndex        =   5
      Top             =   525
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   556
      _Version        =   393216
      Text            =   "DataCombo1"
      Object.DataMember      =   ""
   End
   Begin MSAdodcLib.Adodc AdoRecibo 
      Height          =   330
      Left            =   315
      Top             =   2310
      Visible         =   0   'False
      Width           =   2430
      _ExtentX        =   4286
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
      Caption         =   "Recibo"
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
   Begin MSAdodcLib.Adodc AdoFacturas 
      Height          =   330
      Left            =   315
      Top             =   3360
      Visible         =   0   'False
      Width           =   2430
      _ExtentX        =   4286
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
      Caption         =   "Facturas"
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
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Left            =   1260
      TabIndex        =   21
      Top             =   945
      Width           =   8520
   End
   Begin VB.Label LblRetencion 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      Height          =   330
      Left            =   8190
      TabIndex        =   20
      Top             =   5145
      Width           =   1590
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " RETENCIONES"
      Height          =   330
      Left            =   5775
      TabIndex        =   19
      Top             =   5145
      Width           =   2430
   End
   Begin VB.Label LblEfectivo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      Height          =   330
      Left            =   8190
      TabIndex        =   16
      Top             =   4830
      Width           =   1590
   End
   Begin VB.Label LblTotBanco 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      Height          =   330
      Left            =   8190
      TabIndex        =   17
      Top             =   4515
      Width           =   1590
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " EFECTIVO"
      Height          =   330
      Left            =   5775
      TabIndex        =   14
      Top             =   4830
      Width           =   2430
   End
   Begin VB.Label LblCheque 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " BANCO"
      Height          =   330
      Left            =   5775
      TabIndex        =   15
      Top             =   4515
      Width           =   2430
   End
   Begin VB.Label LblBanco 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   " BANCO"
      Height          =   1380
      Left            =   105
      TabIndex        =   18
      Top             =   4515
      Width           =   5580
   End
   Begin VB.Label LblConcepto 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Gavetas"
      Height          =   645
      Left            =   2100
      TabIndex        =   12
      Top             =   1365
      Width           =   7680
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " POR CONCEPTO DE:"
      Height          =   645
      Left            =   105
      TabIndex        =   13
      Top             =   1365
      Width           =   2010
   End
   Begin VB.Label LabelTotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      Height          =   330
      Left            =   8190
      TabIndex        =   10
      Top             =   5565
      Width           =   1590
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " T O T A L"
      Height          =   330
      Left            =   5775
      TabIndex        =   9
      Top             =   5565
      Width           =   2430
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " FECHA"
      Height          =   330
      Left            =   5565
      TabIndex        =   3
      Top             =   105
      Width           =   855
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " RECIBI DE:"
      Height          =   330
      Left            =   105
      TabIndex        =   11
      Top             =   945
      Width           =   1170
   End
End
Attribute VB_Name = "RAbonos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DBFactura As ADODB.Recordset

Private Sub Command1_Click()
   Imprimir_Comprobante_Caja_Por_Cliente AdoDiarioCaja, TA
End Sub

Private Sub Command2_Click()
   Unload RAbonos
End Sub

Private Sub Command3_Click()
Dim TipoProc As String
Dim SizeLetra As Single
On Error GoTo Errorhandler
SizeLetra = 10
CEConLineas = ProcesarSeteos("RE")
Mensajes = "Imprimir Recibo de Caja."
Titulo = "Formulario de Impresion."
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
RatonReloj
With AdoDiarioCaja.Recordset
 If .RecordCount > 0 Then
     Pagina = 1: InicioX = 0.5: InicioY = 0.1
     NombreCliente = Label2.Caption
     Escala_Centimetro 1, TipoTimes, SizeLetra
     If (SetD(21).PosX + SetD(21).PosY) > 0 Then
        RutaOrigen = RutaSistema & "\FORMATOS\RECIBO.GIF"
       'PrinterPaint RutaOrigen, InicioX, InicioY, SetD(22).PosX, SetD(22).PosY
        PFil = InicioY
        PCol = InicioX
        PrinterPaint LogoTipo, InicioX + 0.8, InicioY, 3, 1.5
        Printer.FontBold = True
        Printer.ForeColor = QBColor(Negro)
        Printer.FontSize = 20: Printer.FontItalic = True
        Printer.CurrentX = CentrarTextoEncab(Empresa, InicioX, 19.5)
        Printer.CurrentY = InicioY
        Printer.Print Empresa
        Printer.FontSize = 9
        Cadena = NombreCiudad & ", " & Direccion & ". Tel�fono: " & Telefono1 & "."
        Printer.CurrentX = CentrarTextoEncab(Cadena, InicioX, 19.5)
        Printer.CurrentY = InicioY + 0.9
        Printer.Print Cadena
        Printer.FontSize = 10
        Cadena = "R.U.C. " & RUC
        Printer.CurrentX = 5
        Printer.CurrentY = InicioY + 1.5
        Printer.Print Cadena
        Printer.FontItalic = False
        Printer.FontSize = 8
        
       'Grafico de las lineas del Recibo
        Printer.Line (InicioX + 0.8, InicioY + 2.2)-(19.3, 12), QBColor(Negro), B
        
        Printer.Line (InicioX + 0.8, InicioY + 2.85)-(19.3, InicioY + 2.85), QBColor(Negro)
        Printer.Line (InicioX + 0.8, InicioY + 3.5)-(19.3, InicioY + 3.5), QBColor(Negro)
        Printer.Line (InicioX + 0.8, InicioY + 4.6)-(19.3, InicioY + 4.6), QBColor(Negro)
        Printer.Line (InicioX + 0.8, InicioY + 5)-(19.3, InicioY + 5), QBColor(Negro)
       'Totales
        Printer.Line (InicioX + 0.8, InicioY + 10.05)-(19.3, InicioY + 10.05), QBColor(Negro)
        Printer.Line (InicioX + 0.8, InicioY + 10.6)-(19.3, InicioY + 10.6), QBColor(Negro)
        
        Printer.Line (InicioX + 2.7, InicioY + 4.6)-(InicioX + 2.7, InicioY + 10.05), QBColor(Negro)
        Printer.Line (InicioX + 6.4, InicioY + 4.6)-(InicioX + 6.4, InicioY + 10.6), QBColor(Negro)
        Printer.Line (InicioX + 9.1, InicioY + 4.6)-(InicioX + 9.1, InicioY + 10.6), QBColor(Negro)
        Printer.Line (InicioX + 11.7, InicioY + 4.6)-(InicioX + 11.7, InicioY + 10.6), QBColor(Negro)
        Printer.Line (InicioX + 14.2, InicioY + 4.6)-(InicioX + 14.2, InicioY + 10.6), QBColor(Negro)
        Printer.Line (InicioX + 16.5, InicioY + 4.6)-(InicioX + 16.5, InicioY + 10.6), QBColor(Negro)
        
        Printer.Line (InicioX + 4.8, InicioY + 11.5)-(8, InicioY + 11.5), QBColor(Negro)
        Printer.Line (InicioX + 10.8, InicioY + 11.5)-(13.7, InicioY + 11.5), QBColor(Negro)
       'Textos del Detalle del Recibo de Caja
        Printer.FontSize = 7
        PrinterTexto InicioX + 0.8, InicioY + 2.3, "FECHA:"
        PrinterTexto InicioX + 0.8, InicioY + 2.9, "RECIBI DE:"
        PrinterTexto InicioX + 0.8, InicioY + 3.6, "POR CONCEPTO DE:"
        Printer.FontSize = 8
        PrinterTexto InicioX + 5, InicioY + 11.55, "C O N F O R M E"
        PrinterTexto InicioX + 11, InicioY + 11.55, "C L I E N T E"
        
        PrinterTexto InicioX + 0.8, InicioY + 4.65, "Factura No."
        PrinterTexto InicioX + 2.8, InicioY + 4.65, "D E T A L L E"
        PrinterTexto InicioX + 6.5, InicioY + 4.65, "Saldo Pendiente"
        PrinterTexto InicioX + 9.2, InicioY + 4.65, "RET. FTE."
        PrinterTexto InicioX + 11.8, InicioY + 4.65, "RET. I.V.A."
        PrinterTexto InicioX + 14.3, InicioY + 4.65, "ABONO"
        PrinterTexto InicioX + 16.6, InicioY + 4.65, "Saldo Actual"
        PrinterTexto InicioX + 3, InicioY + 10.1, "T O T A L E S"
        Printer.FontSize = 16
        PrinterTexto InicioX + 10.3, InicioY + 1.4, "RECIBO DE CAJA No."
        PrinterTexto InicioX + 16.4, InicioY + 1.4, Format(Val(DCRecibo.Text), "0000000")
        Printer.FontBold = False
     End If
    'Iniciamos la impresion
    .MoveFirst
     Printer.FontBold = True
     Printer.FontName = TipoTimes
     Printer.FontSize = SetD(2).Tama�o
     PrinterTexto SetD(2).PosX, SetD(2).PosY, FechaStrgCiudad(MBFecha)
     Printer.FontSize = SetD(3).Tama�o
     PrinterTexto SetD(3).PosX, SetD(3).PosY, NombreCliente
     Printer.FontSize = SetD(5).Tama�o
     PrinterLineasTexto SetD(5).PosX, SetD(5).PosY, LblConcepto, 15
     Printer.FontSize = SetD(10).Tama�o
     PrinterTexto SetD(10).PosX, SetD(10).PosY, LblBanco     ' Banco
     Printer.FontSize = SetD(12).Tama�o
     PrinterTexto SetD(12).PosX, SetD(12).PosY, LblCheque    ' Cheque
     Printer.FontSize = SetD(8).Tama�o
     PrinterVariables SetD(8).PosX, SetD(8).PosY, Total      ' Efectivo
     Printer.FontSize = SetD(9).Tama�o
     PrinterVariables SetD(9).PosX, SetD(9).PosY, SaldoDisp  ' Valor Banco
     PosLinea = SetD(13).PosY
     Printer.FontBold = False
     Real1 = 0:     Real2 = 0:     Real3 = 0:     Real4 = 0
     Do While Not .EOF
        Printer.FontSize = SetD(13).Tama�o
        PrinterTexto SetD(13).PosX, PosLinea, Format(.Fields("Factura"), "0000000")
        Printer.FontSize = SetD(19).Tama�o
        PrinterFields SetD(19).PosX, PosLinea, .Fields("Comprobante")
        Printer.FontSize = SetD(14).Tama�o
        PrinterFields SetD(14).PosX, PosLinea, .Fields("Saldo_Anterior")
        Printer.FontSize = SetD(15).Tama�o
        PrinterFields SetD(15).PosX, PosLinea, .Fields("RET_FTE")
        Printer.FontSize = SetD(16).Tama�o
        PrinterFields SetD(16).PosX, PosLinea, .Fields("RET_IVA")
        Printer.FontSize = SetD(17).Tama�o
        PrinterFields SetD(17).PosX, PosLinea, .Fields("ABONO")
        Printer.FontSize = SetD(18).Tama�o
        PrinterFields SetD(18).PosX, PosLinea, .Fields("Saldo_Actual")
        Real1 = Real1 + .Fields("RET_FTE")
        Real2 = Real2 + .Fields("RET_IVA")
        Real3 = Real3 + .Fields("ABONO")
        Real4 = Real4 + .Fields("Saldo_Actual")
        PosLinea = PosLinea + 0.4
       .MoveNext
     Loop
     PosLinea = SetD(20).PosY
     Printer.FontSize = SetD(20).Tama�o
     PrinterVariables SetD(15).PosX, PosLinea, Real1
     PrinterVariables SetD(16).PosX, PosLinea, Real2
     PrinterVariables SetD(17).PosX, PosLinea, Real3
     PrinterVariables SetD(18).PosX, PosLinea, Real4
     Printer.FontBold = False
     PosLinea = PosLinea + 0.4
  End If
End With
RatonNormal
MensajeEncabData = ""
Printer.EndDoc
End If
Exit Sub
Errorhandler:
    RatonNormal
    ErrorDeImpresion
    Exit Sub
End Sub

Private Sub DCRecibo_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCRecibo_LostFocus()
  RatonReloj
  DGIngCaja.Visible = False
  Contador = 1
  NombreCliente = DCRecibo.Text
  sSQL = "DELETE * " _
       & "FROM Saldo_Diarios " _
       & "WHERE CodigoU = '" & CodigoUsuario & "' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND TP = 'RECI' "
  ConectarAdoExecute sSQL

  Total = 0: SaldoDisp = 0: SaldoCont = 0
  sSQL = "SELECT TA.*,C.Cliente,C.CI_RUC " _
       & "FROM Trans_Abonos As TA,Clientes AS C " _
       & "WHERE TA.Item = '" & NumEmpresa & "' " _
       & "AND TA.Periodo = '" & Periodo_Contable & "' "
  If OpcFactura.value Then
     sSQL = sSQL & "AND TA.Factura = " & Val(DCRecibo) & " "
  ElseIf OpcRecibo.value Then
     sSQL = sSQL & "AND TA.Recibo_No = '" & DCRecibo & "' "
  Else
     sSQL = sSQL & "AND C.Cliente = '" & NombreCliente & "' " _
          & "AND TA.Fecha = #" & BuscarFecha(MBFecha) & "# "
  End If
  sSQL = sSQL & "AND TA.CodigoC = C.Codigo " _
       & "ORDER BY TA.Factura,TA.Comprobante,TA.Banco,TA.Cheque "
  SelectAdodc AdoDiarioCaja, sSQL
  'MsgBox sSQL
  With AdoDiarioCaja.Recordset
   If .RecordCount > 0 Then
       Label2.Caption = " " & .Fields("Cliente")
       TA.Recibi_de = .Fields("Cliente")
       TA.CodigoC = .Fields("CodigoC")
       TA.TP = .Fields("TP")
       TA.Fecha = MBFecha
       TRecibo.CI_RUC = .Fields("CI_RUC")
       Contrato_No = "Abono de Factura(s): "
       Factura_No = .Fields("Factura")
       
       CodigoCliente = .Fields("CodigoC")
       NumCheque = .Fields("Recibo_No")
       Real1 = 0: Real2 = 0: Real3 = 0: Real4 = 0
       NoCheque = Mid(.Fields("Banco"), 1, 5)
       Saldo = 0
       If DBFactura.RecordCount > 0 Then
          DBFactura.MoveFirst
          DBFactura.Find ("Factura = " & Factura_No & " ")
          If Not DBFactura.EOF Then Saldo = DBFactura.Fields("Saldo_MN")
       End If
       If Len(.Fields("Cheque")) > 1 Then NoCheque = NoCheque & "... No. " & .Fields("Cheque")
       Do While Not .EOF
          If Factura_No <> .Fields("Factura") Then
             'If OpcCliente.Value Then
                Procesar_Abonos
                Contrato_No = Contrato_No & Factura_No & " - "
                Factura_No = .Fields("Factura")
                NumCheque = .Fields("Recibo_No")
                Real1 = 0: Real2 = 0: Real3 = 0: Real4 = 0
                Saldo = 0
                If DBFactura.RecordCount > 0 Then
                   DBFactura.MoveFirst
                   DBFactura.Find ("Factura = " & Factura_No & " ")
                   If Not DBFactura.EOF Then Saldo = DBFactura.Fields("Saldo_MN")
                End If

'''              Else
'''                Procesar_Abonos
'''                Contrato_No = Contrato_No & Factura_No & " - "
'''                Factura_No = .Fields("Factura")
'''                Saldo = .Fields("Saldo_MN")
'''                NumCheque = .Fields("Recibo_No")
'''                Real1 = 0: Real2 = 0: Real3 = 0: Real4 = 0
'''             End If
          End If
          Total = Total + .Fields("Abono")
          Select Case .Fields("Banco")
            Case "RETENCION EN LA FUENTE": Real1 = Real1 + .Fields("Abono")
            Case "RETENCION DEL IVA": Real2 = Real2 + .Fields("Abono")
            Case "EFECTIVO MN"
                 Real4 = Real4 + .Fields("Abono")
                 SaldoCont = SaldoCont + .Fields("Abono")
            Case Else
                 Real3 = Real3 + .Fields("Abono")
                 SaldoDisp = SaldoDisp + .Fields("Abono")
                 If IsNumeric(.Fields("Cheque")) Then
                    NombreBanco = .Fields("Banco")
                    NoCheque = .Fields("Cheque")
                 End If
          End Select
          
          NoCheque = Mid(.Fields("Banco"), 1, 5)
          If Len(.Fields("Cheque")) > 1 Then NoCheque = NoCheque & "... No. " & .Fields("Cheque")
         ' MsgBox Total
         .MoveNext
       Loop
                Procesar_Abonos
                Contrato_No = Contrato_No & Factura_No & " - "
                Real1 = 0: Real2 = 0: Real3 = 0: Real4 = 0
                Saldo = 0
                If DBFactura.RecordCount > 0 Then
                   DBFactura.MoveFirst
                   DBFactura.Find ("Factura = " & Factura_No & " ")
                   If Not DBFactura.EOF Then Saldo = DBFactura.Fields("Saldo_MN")
                End If

'       If OpcCliente.Value Then
          'Procesar_Abonos
'''       Else
'''          Procesar_Abonos
'''       End If
       Contrato_No = Contrato_No & Factura_No & " - "
   End If
  End With
  Contrato_No = Contrato_No '& "del d�a: " & MBFecha
  LblConcepto.Caption = Contrato_No
  Real1 = 0:  Real2 = 0:  Real3 = 0:  Real4 = 0: Total = 0
  sSQL = "SELECT Grupo_No As Recibo_No,Numero As Factura,Comprobante,Saldo_Anterior,Enero As RET_FTE,Febrero As RET_IVA,Marzo As ABONO,Abril As ABONO_E,Saldo_Actual " _
       & "FROM Saldo_Diarios " _
       & "WHERE CodigoU = '" & CodigoUsuario & "' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND TP = 'RECI' " _
       & "ORDER BY Numero "
  SelectDataGrid DGIngCaja, AdoDiarioCaja, sSQL
  With AdoDiarioCaja.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          Real1 = Real1 + .Fields("RET_FTE") + .Fields("RET_IVA")
          Real2 = Real2 + .Fields("ABONO")
          Real4 = Real4 + .Fields("ABONO_E")
          Total = Total + .Fields("Saldo_Actual")
         .MoveNext
       Loop
   End If
  End With
  LblBanco.Caption = NombreBanco
  LblCheque.Caption = "No. " & NoCheque
  LabelTotal.Caption = Format(Total, "#,##0.00")
  LblTotBanco.Caption = Format(Real2, "#,##0.00")
  LblEfectivo.Caption = Format(Real4, "#,##0.00")
  LblRetencion.Caption = Format(Real1, "#,##0.00")
  DGIngCaja.Visible = True
  RatonNormal
End Sub

Private Sub Form_Activate()
  sSQL = "SELECT * " _
       & "FROM Facturas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND T <> '" & Anulado & "' " _
       & "ORDER BY TC,Factura "
  DBFactura.Open sSQL, AdoStrCnn, , , adCmdText
  
  Mifecha = BuscarFecha(FechaSistema)
  Listar_Recibos
  RatonNormal
  OpcFactura.SetFocus
End Sub

Private Sub Form_Load()
   CentrarForm RAbonos
   ConectarAdodc AdoRecibo
   ConectarAdodc AdoFacturas
   ConectarAdodc AdoClientes
   ConectarAdodc AdoDiarioCaja
  'Temporal
   Set DBFactura = New ADODB.Recordset
   DBFactura.CursorType = adOpenStatic
   DBFactura.CursorLocation = adUseClient
End Sub

Private Sub MBFecha_GotFocus()
  MarcarTexto MBFecha
End Sub

Private Sub MBFecha_LostFocus()
  FechaValida MBFecha
End Sub

Public Sub Listar_Recibos()
  FechaValida MBFecha
  If OpcFactura.value Then
     sSQL = "SELECT Factura As Recibo " _
          & "FROM Trans_Abonos " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "GROUP BY Factura " _
          & "ORDER BY Factura "
  ElseIf OpcRecibo.value Then
     sSQL = "SELECT Recibo_No As Recibo " _
          & "FROM Trans_Abonos " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "GROUP BY Recibo_No " _
          & "ORDER BY Recibo_No "
  Else
     sSQL = "SELECT TA.CodigoC,C.Cliente As Recibo " _
          & "FROM Clientes As C,Trans_Abonos As TA " _
          & "WHERE C.Cliente <> '.' " _
          & "AND TA.Item = '" & NumEmpresa & "' " _
          & "AND TA.Periodo = '" & Periodo_Contable & "' " _
          & "AND TA.CodigoC = C.Codigo " _
          & "GROUP BY C.Cliente,TA.CodigoC " _
          & "ORDER BY C.Cliente "
  End If
  SelectDBCombo DCRecibo, AdoRecibo, sSQL, "Recibo", True
End Sub

Public Sub Procesar_Abonos()
   'MsgBox Real1 & vbCrLf & Real2 & vbCrLf & Real3 & vbCrLf & Real4
   Saldo_ME = Saldo + Real1 + Real2 + Real3 + Real4
   If Saldo_ME > 0 Then
   SetAdoAddNew "Saldo_Diarios"
   SetAdoFields "ME", CByte(Contador)
   SetAdoFields "TP", "RECI"
   SetAdoFields "CodigoU", CodigoUsuario
   SetAdoFields "CodigoC", CodigoCliente
   SetAdoFields "Comprobante", NoCheque
   SetAdoFields "Item", NumEmpresa
   SetAdoFields "Numero", Factura_No
   SetAdoFields "Grupo_No", NumCheque
   SetAdoFields "Saldo_Anterior", Saldo_ME
   SetAdoFields "Enero", Real1
   SetAdoFields "Febrero", Real2
   SetAdoFields "Marzo", Real3
   SetAdoFields "Abril", Real4
   SetAdoFields "Saldo_Actual", Saldo
   SetAdoUpdate
   Contador = Contador + 1
   End If
End Sub

Private Sub OpcCliente_Click()
  Listar_Recibos
End Sub

Private Sub OpcFactura_Click()
  Listar_Recibos
End Sub

Private Sub OpcRecibo_Click()
  Listar_Recibos
End Sub
