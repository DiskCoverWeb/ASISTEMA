VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form AbonoAnticipo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "INGRESO DE CAJA"
   ClientHeight    =   4980
   ClientLeft      =   5025
   ClientTop       =   4380
   ClientWidth     =   8100
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
   ScaleHeight     =   4980
   ScaleWidth      =   8100
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox TextCheqNo 
      Height          =   330
      Left            =   5145
      MaxLength       =   10
      TabIndex        =   29
      Top             =   4200
      Width           =   1800
   End
   Begin VB.TextBox TextCajaMN 
      Alignment       =   1  'Right Justify
      Height          =   330
      Left            =   5145
      MaxLength       =   14
      TabIndex        =   27
      Text            =   "0"
      Top             =   3885
      Width           =   1800
   End
   Begin VB.TextBox TxtBanco 
      Height          =   330
      Left            =   3570
      MaxLength       =   25
      TabIndex        =   19
      Top             =   2940
      Width           =   3375
   End
   Begin VB.CheckBox CheqRecibo 
      Caption         =   "&INGRESO DE CAJA No."
      Height          =   330
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Value           =   1  'Checked
      Width           =   2430
   End
   Begin VB.TextBox TxtRecibo 
      Alignment       =   1  'Right Justify
      Height          =   330
      Left            =   2625
      MaxLength       =   14
      TabIndex        =   1
      Text            =   "0"
      Top             =   105
      Width           =   1275
   End
   Begin MSAdodcLib.Adodc AdoFactura 
      Height          =   330
      Left            =   105
      Top             =   4935
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
      Left            =   7035
      Picture         =   "Anticipo.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   1050
      Width           =   960
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   855
      Left            =   7035
      Picture         =   "Anticipo.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   105
      Width           =   960
   End
   Begin MSDataListLib.DataCombo DCFactura 
      Bindings        =   "Anticipo.frx":0D0C
      DataSource      =   "AdoFactura"
      Height          =   315
      Left            =   5145
      TabIndex        =   11
      Top             =   945
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   556
      _Version        =   393216
      Text            =   "Factura"
   End
   Begin MSMask.MaskEdBox MBFecha 
      Height          =   330
      Left            =   5670
      TabIndex        =   3
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
   Begin MSDataListLib.DataCombo DCBanco 
      Bindings        =   "Anticipo.frx":0D25
      DataSource      =   "AdoBanco"
      Height          =   315
      Left            =   1785
      TabIndex        =   15
      Top             =   1785
      Width           =   5160
      _ExtentX        =   9102
      _ExtentY        =   556
      _Version        =   393216
      Text            =   "Banco"
   End
   Begin MSAdodcLib.Adodc AdoBanco 
      Height          =   330
      Left            =   105
      Top             =   5250
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
   Begin MSDataListLib.DataCombo DCTipo 
      Bindings        =   "Anticipo.frx":0D3C
      DataSource      =   "AdoDetAcomp"
      Height          =   315
      Left            =   1785
      TabIndex        =   7
      Top             =   945
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   556
      _Version        =   393216
      Text            =   "FA/NV"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSAdodcLib.Adodc AdoDetAcomp 
      Height          =   330
      Left            =   1995
      Top             =   4935
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
   Begin MSDataListLib.DataCombo DCSerie 
      Bindings        =   "Anticipo.frx":0D56
      DataSource      =   "AdoSerie"
      Height          =   315
      Left            =   3360
      TabIndex        =   9
      Top             =   945
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   556
      _Version        =   393216
      Text            =   "001001"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSAdodcLib.Adodc AdoSerie 
      Height          =   330
      Left            =   1995
      Top             =   5250
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
   Begin MSDataListLib.DataCombo DCCliente 
      Bindings        =   "Anticipo.frx":0D6D
      DataSource      =   "AdoCliente"
      Height          =   345
      Left            =   1050
      TabIndex        =   5
      Top             =   525
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   609
      _Version        =   393216
      Text            =   "Clientes"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSAdodcLib.Adodc AdoCliente 
      Height          =   330
      Left            =   3885
      Top             =   4935
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
   Begin MSDataListLib.DataCombo DCAutorizacion 
      Bindings        =   "Anticipo.frx":0D86
      DataSource      =   "AdoAutorizacion"
      Height          =   315
      Left            =   1365
      TabIndex        =   13
      Top             =   1365
      Width           =   5580
      _ExtentX        =   9843
      _ExtentY        =   556
      _Version        =   393216
      Text            =   "1234567890"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSAdodcLib.Adodc AdoAutorizacion 
      Height          =   330
      Left            =   3885
      Top             =   5250
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
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "FECHA DE EMISION"
      ForeColor       =   &H00C00000&
      Height          =   330
      Left            =   3465
      TabIndex        =   17
      Top             =   2205
      Width           =   3480
   End
   Begin VB.Label LabelPend 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   330
      Left            =   5145
      TabIndex        =   31
      Top             =   4515
      Width           =   1800
   End
   Begin VB.Label LabelAnticipo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   330
      Left            =   5145
      TabIndex        =   25
      Top             =   3570
      Width           =   1800
   End
   Begin VB.Label LabelSaldo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   330
      Left            =   5145
      TabIndex        =   23
      Top             =   3255
      Width           =   1800
   End
   Begin VB.Label Label10 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Saldo Actual"
      Height          =   330
      Left            =   3570
      TabIndex        =   30
      Top             =   4515
      Width           =   1590
   End
   Begin VB.Label Label14 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Cheque No."
      Height          =   330
      Left            =   3570
      TabIndex        =   28
      Top             =   4200
      Width           =   1590
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Valor Abonado"
      Height          =   330
      Left            =   3570
      TabIndex        =   26
      Top             =   3885
      Width           =   1590
   End
   Begin VB.Label Label17 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Detalle del Abono"
      Height          =   330
      Left            =   3570
      TabIndex        =   18
      Top             =   2625
      Width           =   3375
   End
   Begin VB.Label Label16 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Cuenta de Abono"
      Height          =   330
      Left            =   105
      TabIndex        =   14
      Top             =   1785
      Width           =   1695
   End
   Begin VB.Label Label15 
      BackColor       =   &H000040C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Autorización"
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   105
      TabIndex        =   12
      Top             =   1365
      Width           =   1275
   End
   Begin VB.Label Label12 
      BackColor       =   &H000040C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &CLIENTE"
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   105
      TabIndex        =   4
      Top             =   525
      Width           =   960
   End
   Begin VB.Label Label4 
      BackColor       =   &H000040C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Serie"
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   2730
      TabIndex        =   8
      Top             =   945
      Width           =   645
   End
   Begin VB.Label Label13 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Anticipo a Favor"
      Height          =   330
      Left            =   3570
      TabIndex        =   24
      Top             =   3570
      Width           =   1590
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CI/RUC"
      ForeColor       =   &H00C00000&
      Height          =   330
      Left            =   105
      TabIndex        =   16
      Top             =   2205
      Width           =   3375
   End
   Begin VB.Label LblNota 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nota"
      ForeColor       =   &H00C000C0&
      Height          =   1065
      Left            =   105
      TabIndex        =   21
      Top             =   3780
      Width           =   3375
   End
   Begin VB.Label Label9 
      BackColor       =   &H000040C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Tipo Documento"
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   105
      TabIndex        =   6
      Top             =   945
      Width           =   1695
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &Fecha del Abono"
      Height          =   330
      Left            =   3990
      TabIndex        =   2
      Top             =   105
      Width           =   1695
   End
   Begin VB.Label LblObs 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Observacion"
      ForeColor       =   &H00C000C0&
      Height          =   1065
      Left            =   105
      TabIndex        =   20
      Top             =   2625
      Width           =   3375
   End
   Begin VB.Label Label2 
      BackColor       =   &H000040C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " No."
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   4620
      TabIndex        =   10
      Top             =   945
      Width           =   540
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Saldo Pendiente"
      Height          =   330
      Left            =   3570
      TabIndex        =   22
      Top             =   3255
      Width           =   1590
   End
End
Attribute VB_Name = "AbonoAnticipo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Dim Costo_Banco As Double
Dim Total_Costo_Banco As Double
Dim Cta_Gasto_Banco As String
'

Public Sub Listar_Facturas_Pendientes()
  Label2.Caption = " Factura"
  If TipoFactura = "NV" Then Label2.Caption = " Nota de Venta"
  SQL1 = "SELECT F.TC,F.Factura,F.Autorizacion,F.Serie,F.CodigoC,F.Fecha,F.Fecha_V," _
       & "F.Saldo_MN,F.Cta_CxP,F.Nota,F.Observacion," _
       & "C.Cliente,C.Direccion,C.CI_RUC,C.Telefono,C.Grupo " _
       & "FROM Facturas As F,Clientes As C " _
       & "WHERE F.T = '" & Pendiente & "' " _
       & "AND F.Item = '" & NumEmpresa & "' " _
       & "AND F.Periodo = '" & Periodo_Contable & "' " _
       & "AND F.CodigoC = '" & CodigoCli & "' " _
       & "AND F.TC = '" & TipoFactura & "' " _
       & "AND F.Serie = '" & SerieFactura & "' " _
       & "AND F.Autorizacion = '" & Autorizacion & "' " _
       & "AND F.Factura = " & Factura_No & " " _
       & "AND F.Saldo_MN > 0 " _
       & "AND F.CodigoC = C.Codigo " _
       & "ORDER BY F.TC,F.Factura "
  SelectDB_Combo DCFactura, AdoFactura, SQL1, "Factura"
End Sub

Private Sub Command1_Click()
  FechaValida MBFecha
  TextoValido TextCheqNo
  FechaTexto = MBFecha
  If CFechaLong(MBFecha) < CFechaLong(FechaCorte) Then
     MsgBox "No se puede grabar abonos con fecha inferior a la emision de la factura"
     MBFecha.SetFocus
  Else
    Mensajes = "Esta Seguro que desea grabar Abono."
    Titulo = "Formulario de Grabación."
    If BoxMensaje = vbYes Then
       FechaTexto = MBFecha ' FechaSistema
       If CheqRecibo.value = 1 Then
          DiarioCaja = ReadSetDataNum("Recibo_No", True, True)
       Else
          DiarioCaja = Val(TxtRecibo)
       End If
       SaldoDisp = Saldo - TotalCajaMN
       LabelPend.Caption = Format$(SaldoDisp, "#,##0.00")
       Cta = SinEspaciosIzq(DCBanco)
      'Abono de Factura
       TA.Recibi_de = NombreCliente
       TA.T = Normal
       TA.TP = TipoFactura
       TA.Fecha = MBFecha
       TA.Cta = Cta
       TA.Banco = TrimStrg(TxtBanco)
       TA.Cheque = TrimStrg(TextCheqNo)
       TA.Abono = TotalCajaMN
       TA.CodigoC = CodigoCli
       TA.Recibo_No = Format$(DiarioCaja, "0000000000")
       Grabar_Abonos TA
      'Tipo de Abonos con SubCtas
      'Costo Bancario por deposito
       If TipoProc = "TRANSFERENCIA" Then
          If Costo_Banco > 0 Then
             TA.TP = "CB"
             TA.Fecha = MBFecha
             TA.Cta = Cta
             TA.Cta_CxP = Cta_Gasto_Banco
             TA.Banco = "COSTO BANCARIO"
             TA.Cheque = "TRANSF"
             TA.Abono = Costo_Banco
             Grabar_Abonos TA
          End If
       End If
       Actualizar_Saldos_Facturas_SP TA.TP, TA.Serie, TA.Factura
       RatonNormal
       Imprimir_Comprobante_Caja TA
       Listar_Facturas_Pendientes
       MsgBox "Abono Realizado con éxito"
       DCFactura.SetFocus
    End If
  End If
 'Unload AbonoEfectivo
End Sub

Private Sub Command2_Click()
   Control_Procesos Normal, "Salir de abonos de facturas por Anticipos"
   Unload Me
End Sub

Private Sub DCAutorizacion_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCAutorizacion_LostFocus()
  Autorizacion = DCAutorizacion
  If Autorizacion = "" Then Autorizacion = Ninguno
  Codigo1 = Ninguno
  Saldo = 0
  TotalCajaMN = 0
  Cotizacion = 0
  TotalDolar = 0
  Saldo_ME = 0
'  Label3.Caption = ""
'  Label1.Caption = ""
  LabelPend.Caption = ""
  LabelSaldo.Caption = ""
  
  Listar_Facturas_Pendientes
  
  With AdoFactura.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Factura = " & Val(DCFactura) & " ")
       If Not .EOF Then
          Label8.Caption = "FECHA DE EMISION: " & .fields("Fecha")
          Label11.Caption = "C.I./R.U.C.: " & .fields("CI_RUC")
          Grupo_No = " " & .fields("Grupo")
          TextCheqNo = Grupo_No
          LblObs.Caption = " " & .fields("Observacion")
          LblNota.Caption = " " & .fields("Nota")
          CodigoCliente = .fields("CodigoC")
          CodigoCli = .fields("CodigoC")
          NombreCliente = .fields("Cliente")
          DireccionCli = .fields("Direccion")
          Factura_No = .fields("Factura")
          Cta_Cobrar = .fields("Cta_CxP")
          TipoFactura = .fields("TC")
          Saldo = .fields("Saldo_MN")
         'Datos del Abonos
          FechaCorte = .fields("Fecha")
          TA.Serie = .fields("Serie")
          TA.Autorizacion = .fields("Autorizacion")
          TA.TP = TipoFactura
          TA.Fecha = MBFecha
          TA.Cta_CxP = Cta_Cobrar
          TA.Factura = Factura_No
          TA.CodigoC = CodigoCliente
          
          LabelSaldo.Caption = Format$(Saldo, "#,##0.00")
          Command1.Enabled = True
          'Label3.Caption = " " & NombreCliente
          'Label1.Caption = "Autorización: " & TA.Autorizacion & " => Factura No. " & TA.Serie & "-" & Format$(TA.Factura, "0000000")
          SaldoDisp = Saldo
          LabelPend.Caption = Format$(SaldoDisp, "#,##0.00")
          AbonoAnticipo.Caption = "INGRESO DE CAJA (" & TipoFactura & ")"
          TextCajaMN.Text = LabelPend.Caption
         'MsgBox CFechaLong(MBFecha) & vbCrLf & CFechaLong(FechaCorte)
          If CFechaLong(MBFecha) < CFechaLong(FechaCorte) Then
             MsgBox "No se puede grabar abonos con fecha inferior a la emision de la factura"
             MBFecha.SetFocus
          Else
             DCBanco.SetFocus
          End If
       Else
          MsgBox "Esta Factura no esta pendiente"
          Command1.Enabled = False
          DCFactura.SetFocus
       End If
    End If
  End With
End Sub

Private Sub DCBanco_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCBanco_LostFocus()
  FechaValida MBFecha
  Cta_Aux = SinEspaciosIzq(DCBanco)
  TotalSubCta = 0
  Select Case TipoProc
    Case "ANTICIPOS"
         sSQL = "SELECT (CC.Codigo & '  ' & CC.Cuenta) As NomCuenta," _
              & "SUM(TS.Creditos-TS.Debitos) As Saldo_Pendiente " _
              & "FROM Catalogo_Cuentas As CC, Trans_SubCtas As TS " _
              & "WHERE CC.Item = '" & NumEmpresa & "' " _
              & "AND CC.Periodo = '" & Periodo_Contable & "' " _
              & "AND TS.Codigo = '" & CodigoCli & "' " _
              & "AND CC.Codigo = '" & Cta_Aux & "' " _
              & "AND TS.T <> 'A' " _
              & "AND CC.TC = 'P' " _
              & "AND CC.DG = 'D' " _
              & "AND CC.Codigo = TS.Cta " _
              & "AND CC.TC = TS.TC " _
              & "AND CC.Item = TS.Item " _
              & "AND CC.Periodo = TS.Periodo " _
              & "GROUP BY CC.Codigo,CC.Cuenta " _
              & "ORDER BY CC.Codigo "
         SelectDB_Combo DCBanco, AdoBanco, sSQL, "NomCuenta"
         TxtBanco = UCaseStrg(TrimStrg(MidStrg(DCBanco, Len(Cta_Aux) + 1, Len(DCBanco))))
         'MsgBox AdoBanco.Recordset.RecordCount
         If AdoBanco.Recordset.RecordCount > 0 Then
            TotalSubCta = AdoBanco.Recordset.fields("Saldo_Pendiente")
         End If
         If TotalSubCta <= 0 Then
            MsgBox "Este Cliente no tiene alcance para abonar por Anticipos"
            DCCliente.SetFocus
         End If
    Case "ANTICIPOSCXC"
         sSQL = "SELECT (CC.Codigo & '  ' & CC.Cuenta) As NomCuenta," _
              & "SUM(TS.Debitos-TS.Creditos) As Saldo_Pendiente " _
              & "FROM Catalogo_Cuentas As CC, Trans_SubCtas As TS " _
              & "WHERE CC.Item = '" & NumEmpresa & "' " _
              & "AND CC.Periodo = '" & Periodo_Contable & "' " _
              & "AND TS.Codigo = '" & CodigoCli & "' " _
              & "AND CC.Codigo = '" & Cta_Aux & "' " _
              & "AND TS.T = 'N' " _
              & "AND CC.TC = 'C' " _
              & "AND CC.DG = 'D' " _
              & "AND CC.Codigo = TS.Cta " _
              & "AND CC.TC = TS.TC " _
              & "AND CC.Item = TS.Item " _
              & "AND CC.Periodo = TS.Periodo " _
              & "GROUP BY CC.Codigo,CC.Cuenta " _
              & "ORDER BY CC.Codigo "
         SelectDB_Combo DCBanco, AdoBanco, sSQL, "NomCuenta"
         TxtBanco = UCaseStrg(TrimStrg(MidStrg(DCBanco, Len(Cta_Aux) + 1, Len(DCBanco))))
        'MsgBox AdoBanco.Recordset.RecordCount
         If AdoBanco.Recordset.RecordCount > 0 Then
            TotalSubCta = AdoBanco.Recordset.fields("Saldo_Pendiente")
         End If
         If TotalSubCta <= 0 Then
            MsgBox "Este Cliente no tiene alcance para abonar por Anticipos"
            DCCliente.SetFocus
         End If
    Case "EFECTIVO"
         TxtBanco = "EFECTIVO MN"
    Case "BANCOS"
         TxtBanco = "DEPOSITO EN EFECTIVO"
    Case "TARJETA"
         TxtBanco = UCaseStrg(TrimStrg(MidStrg(DCBanco, Len(Cta_Aux) + 1, Len(DCBanco))))
    Case "DIFERENCIAS"
         TxtBanco = UCaseStrg(TrimStrg(MidStrg(DCBanco, Len(Cta_Aux) + 1, Len(DCBanco))))
    Case "TRANSFERENCIA"
         TxtBanco = "TRANSFERENCIA"
  End Select
  LabelAnticipo.Caption = Format$(TotalSubCta, "#,##0.00")
  TxtBanco.SetFocus
End Sub

Private Sub DCCliente_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCCliente_LostFocus()
  CodigoCli = Buscar_Beneficiario(DCCliente, X_Beneficiario)
 'Segun el tipo de pago seleccionamos las cuentas de acreditacion del abono
  Select Case TipoProc
    Case "ANTICIPOS"
         sSQL = "SELECT (CC.Codigo & '  ' & CC.Cuenta) As NomCuenta," _
              & "SUM(TS.Creditos-TS.Debitos) As Saldo_Pendiente " _
              & "FROM Catalogo_Cuentas As CC, Trans_SubCtas As TS " _
              & "WHERE CC.Item = '" & NumEmpresa & "' " _
              & "AND CC.Periodo = '" & Periodo_Contable & "' " _
              & "AND TS.Codigo = '" & CodigoCli & "' " _
              & "AND TS.T = 'N' " _
              & "AND CC.TC = 'P' " _
              & "AND CC.DG = 'D' " _
              & "AND CC.Codigo = TS.Cta " _
              & "AND CC.TC = TS.TC " _
              & "AND CC.Item = TS.Item " _
              & "AND CC.Periodo = TS.Periodo " _
              & "GROUP BY CC.Codigo,CC.Cuenta " _
              & "ORDER BY CC.Codigo "
    Case "ANTICIPOSCXC"
         sSQL = "SELECT (CC.Codigo & '  ' & CC.Cuenta) As NomCuenta," _
              & "SUM(TS.Debitos-TS.Creditos) As Saldo_Pendiente " _
              & "FROM Catalogo_Cuentas As CC, Trans_SubCtas As TS " _
              & "WHERE CC.Item = '" & NumEmpresa & "' " _
              & "AND CC.Periodo = '" & Periodo_Contable & "' " _
              & "AND TS.Codigo = '" & CodigoCli & "' " _
              & "AND TS.T = 'N' " _
              & "AND CC.TC = 'C' " _
              & "AND CC.DG = 'D' " _
              & "AND CC.Codigo = TS.Cta " _
              & "AND CC.TC = TS.TC " _
              & "AND CC.Item = TS.Item " _
              & "AND CC.Periodo = TS.Periodo " _
              & "GROUP BY CC.Codigo,CC.Cuenta " _
              & "ORDER BY CC.Codigo "
    Case "EFECTIVO"
         sSQL = "SELECT (Codigo & '  ' & Cuenta) As NomCuenta " _
              & "FROM Catalogo_Cuentas " _
              & "WHERE TC = 'CJ' " _
              & "AND DG = 'D' " _
              & "AND Item = '" & NumEmpresa & "' " _
              & "AND Periodo = '" & Periodo_Contable & "' " _
              & "ORDER BY Codigo "
    Case "BANCOS", "TRANSFERENCIA"
         sSQL = "SELECT (Codigo & '  ' & Cuenta) As NomCuenta " _
              & "FROM Catalogo_Cuentas " _
              & "WHERE TC = 'BA' " _
              & "AND DG = 'D' " _
              & "AND Item = '" & NumEmpresa & "' " _
              & "AND Periodo = '" & Periodo_Contable & "' " _
              & "ORDER BY Codigo "
    Case "TARJETA"
         sSQL = "SELECT (Codigo & '  ' & Cuenta) As NomCuenta " _
              & "FROM Catalogo_Cuentas " _
              & "WHERE TC = 'TJ' " _
              & "AND DG = 'D' " _
              & "AND Item = '" & NumEmpresa & "' " _
              & "AND Periodo = '" & Periodo_Contable & "' " _
              & "ORDER BY Codigo "
    Case "DIFERENCIAS"
         sSQL = "SELECT (Codigo & '  ' & Cuenta) As NomCuenta " _
              & "FROM Catalogo_Cuentas " _
              & "WHERE Codigo BETWEEN '4' and '5.9.99.99.99.999' " _
              & "AND DG = 'D' " _
              & "AND Item = '" & NumEmpresa & "' " _
              & "AND Periodo = '" & Periodo_Contable & "' " _
              & "ORDER BY Codigo "
  End Select
  SelectDB_Combo DCBanco, AdoBanco, sSQL, "NomCuenta"
  
  sSQL = "SELECT TC " _
       & "FROM Facturas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND TC <> 'OP' " _
       & "AND CodigoC = '" & CodigoCli & "' " _
       & "AND Saldo_MN > 0 " _
       & "AND T <> 'A' " _
       & "GROUP BY TC " _
       & "ORDER BY TC "
  SelectDB_Combo DCTipo, AdoDetAcomp, sSQL, "TC"
End Sub

Private Sub DCFactura_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
  Factura_No = Val(DCFactura)
  sSQL = "SELECT Autorizacion " _
       & "FROM Facturas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND TC <> 'OP' " _
       & "AND CodigoC = '" & CodigoCli & "' " _
       & "AND TC = '" & TipoFactura & "' " _
       & "AND Serie = '" & SerieFactura & "' " _
       & "AND Factura = " & Factura_No & " " _
       & "AND Saldo_MN > 0 " _
       & "AND T <> 'A' " _
       & "GROUP BY Autorizacion " _
       & "ORDER BY Autorizacion "
  SelectDB_Combo DCAutorizacion, AdoAutorizacion, sSQL, "Autorizacion"
 'MsgBox CodigoCli & vbCrLf & AdoAutorizacion.Recordset.RecordCount
End Sub

Private Sub DCSerie_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCSerie_LostFocus()
  SerieFactura = DCSerie
  If SerieFactura = "" Then SerieFactura = Ninguno
  sSQL = "SELECT Factura " _
       & "FROM Facturas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND TC = '" & TipoFactura & "' " _
       & "AND Serie = '" & SerieFactura & "' " _
       & "AND CodigoC = '" & CodigoCli & "' " _
       & "AND Saldo_MN > 0 " _
       & "AND T <> 'A' " _
       & "GROUP BY Factura " _
       & "ORDER BY Factura "
  SelectDB_Combo DCFactura, AdoFactura, sSQL, "Factura"
End Sub

Private Sub DCTipo_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCTipo_LostFocus()
  TipoFactura = DCTipo
  If TipoFactura = "" Then TipoFactura = Ninguno
  
 'Listamos las autorizaciones de facturas pendientes por cliente
  sSQL = "SELECT Serie " _
       & "FROM Facturas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND TC = '" & TipoFactura & "' " _
       & "AND CodigoC = '" & CodigoCli & "' " _
       & "AND Saldo_MN > 0 " _
       & "AND T <> 'A' " _
       & "GROUP BY Serie " _
       & "ORDER BY Serie "
  SelectDB_Combo DCSerie, AdoSerie, sSQL, "Serie"
End Sub

Private Sub Form_Activate()
  ControlEsNumerico TextCajaMN
  Costo_Banco = Leer_Campo_Empresa("Costo_Bancario")
  Cta_Gasto_Banco = Leer_Seteos_Ctas("Cta_Gasto_Bancario")
  
  sSQL = "SELECT C.Cliente,C.Codigo,C.CI_RUC,COUNT(F.Factura) As CantFacturas " _
       & "FROM Clientes As C, Facturas As F " _
       & "WHERE F.Item = '" & NumEmpresa & "' " _
       & "AND F.Periodo = '" & Periodo_Contable & "' " _
       & "AND F.Saldo_MN > 0 " _
       & "AND F.T <> 'A' " _
       & "AND F.TC <> 'OP' " _
       & "AND C.Codigo = F.CodigoC " _
       & "GROUP BY C.Cliente,C.Codigo,C.CI_RUC " _
       & "ORDER BY Cliente "
  SelectDB_Combo DCCliente, AdoCliente, sSQL, "Cliente"
  DiarioCaja = ReadSetDataNum("Recibo_No", True, False)
  If CheqRecibo.value = 1 Then TxtRecibo = Format$(DiarioCaja, "0000000") Else TxtRecibo = ""
  Mifecha = BuscarFecha(FechaTexto)
  TxtBanco = Ninguno
  TextCheqNo = Ninguno
  MBFecha = FechaSistema
  Select Case TipoProc
    Case "ANTICIPOS", "ANTICIPOSCXC"
         Label17.Caption = " DETALLE DEL ANTICIPO"
         Label14.Caption = " GRUPO NO."
    Case "EFECTIVO"
         Label17.Caption = " DETALLE DEL ABONO"
         Label14.Caption = " GRUPO NO."
    Case "BANCOS"
         Label17.Caption = " DETALLE DEL DEPOSITO"
         Label14.Caption = " CHEQ/TRANSF"
    Case "TARJETA"
         Label17.Caption = " DETALLE DE LA TARJETA"
         Label14.Caption = " VAUCHER"
    Case "DIFERENCIAS"
         Label17.Caption = " DETALLE SALDO PENDIENTE"
         Label14.Caption = " GRUPO NO."
    Case "TRANSFERENCIA"
         Label17.Caption = " DETALLE DE LA TRANSFERENCIA"
         Label14.Caption = " Transferencia"
  End Select
  If Bloquear_Control Then Command1.Enabled = False
  RatonNormal
End Sub

Private Sub Form_Load()
   CentrarForm AbonoAnticipo
   ConectarAdodc AdoSerie
   ConectarAdodc AdoBanco
   ConectarAdodc AdoCliente
   ConectarAdodc AdoFactura
   ConectarAdodc AdoDetAcomp
   ConectarAdodc AdoAutorizacion
End Sub

Private Sub TextCajaMN_GotFocus()
  TextCajaMN = Saldo
  MarcarTexto TextCajaMN
End Sub

Private Sub TextCajaMN_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextCajaMN_LostFocus()
  TextoValido TextCajaMN, True
  TotalCajaMN = Redondear(Val(CCur(TextCajaMN)), 2)
  TextCajaMN = Format$(TotalCajaMN, "#,##0.00")
  Select Case TipoProc
    Case "EFECTIVO", "BANCOS", "TARJETA", "DIFERENCIAS", "TRANSFERENCIA": TotalSubCta = TotalCajaMN + 0.1
  End Select
  If TotalCajaMN > TotalSubCta Then
     MsgBox "No hay alcance en Anticipos"
     TextCajaMN = "0.00"
     DCFactura.SetFocus
  Else
     SaldoDisp = Saldo - TotalCajaMN
     If SaldoDisp < 0 Then
        MsgBox "El valor abonado excede el Saldo Pendiente"
        TextCajaMN.SetFocus
     End If
  End If
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

Private Sub TextCheqNo_GotFocus()
  MarcarTexto TextCheqNo
End Sub

Private Sub TextCheqNo_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextCheqNo_LostFocus()
  TextoValido TextCheqNo
End Sub

Private Sub TxtBanco_GotFocus()
  MarcarTexto TxtBanco
End Sub

Private Sub TxtBanco_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
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

