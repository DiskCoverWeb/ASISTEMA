VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Begin VB.Form NotaEgreso 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NOTA DE EGRESO"
   ClientHeight    =   1380
   ClientLeft      =   5025
   ClientTop       =   4380
   ClientWidth     =   7485
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
   ScaleHeight     =   1380
   ScaleWidth      =   7485
   ShowInTaskbar   =   0   'False
   Begin MSAdodcLib.Adodc AdoFactura 
      Height          =   330
      Left            =   105
      Top             =   945
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
      Height          =   750
      Left            =   6510
      Picture         =   "NotaEgre.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   105
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Imprimir"
      Height          =   750
      Left            =   5565
      Picture         =   "NotaEgre.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   105
      Width           =   855
   End
   Begin MSDataListLib.DataCombo DCFactura 
      Bindings        =   "NotaEgre.frx":0BD4
      DataSource      =   "AdoFactura"
      Height          =   315
      Left            =   1260
      TabIndex        =   5
      Top             =   105
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   556
      _Version        =   393216
      Text            =   "Factura"
   End
   Begin MSAdodcLib.Adodc AdoDet 
      Height          =   330
      Left            =   1995
      Top             =   945
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
      Caption         =   "Det"
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
   Begin MSAdodcLib.Adodc AdoFact 
      Height          =   330
      Left            =   3885
      Top             =   945
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
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
      Caption         =   "Fact"
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
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Left            =   3570
      TabIndex        =   4
      Top             =   105
      Width           =   1905
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " F&actura No."
      Height          =   330
      Left            =   105
      TabIndex        =   1
      Top             =   105
      Width           =   1170
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Left            =   105
      TabIndex        =   0
      Top             =   525
      Width           =   5370
   End
End
Attribute VB_Name = "NotaEgreso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
  ImprimirNotaEgreso AdoFact, AdoDet, Factura_No
End Sub

Private Sub Command2_Click()
   Unload NotaEgreso
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
  With AdoFactura.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Factura = " & Val(DCFactura.Text) & " ")
       If Not .EOF Then
          CodigoCliente = .Fields("CodigoC")
          NombreCliente = .Fields("Cliente")
          DireccionCli = .Fields("Direccion")
          Factura_No = .Fields("Factura")
          Cta_Cobrar = .Fields("Cta_CxP")
          Saldo = .Fields("Saldo_MN")
          Command1.Enabled = True
          Label3.Caption = NombreCliente
          Label1.Caption = " " & Factura_No
          SaldoDisp = Saldo - TotalCajaMN - TotalCajaME - Total_Bancos - Total_Tarjeta - Total_IVA - Total_Ret
          Command1.SetFocus
       Else
          MsgBox "Esta Factura no esta pendiente"
          Command1.Enabled = False
          DCFactura.SetFocus
       End If
    End If
  End With
End Sub

Private Sub Form_Activate()
  SQL1 = "SELECT F.TC,F.Factura,F.CodigoC,F.Fecha,F.Fecha_V,F.Saldo_MN,F.Cta_CxP," _
       & "C.Cliente,C.Direccion,C.CI_RUC,C.Telefono " _
       & "FROM Facturas As F,Clientes As C " _
       & "WHERE F.Item = '" & NumEmpresa & "' " _
       & "AND F.TC <> 'C' " _
       & "AND F.TC <> 'P' " _
       & "AND F.CodigoC = C.Codigo " _
       & "ORDER BY F.Factura "
  SelectDBCombo DCFactura, AdoFactura, SQL1, "Factura", True
  Mifecha = BuscarFecha(FechaTexto)
  RatonNormal
End Sub

Private Sub Form_Load()
   CentrarForm NotaEgreso
   ConectarAdodc AdoFactura
   ConectarAdodc AdoFact
   ConectarAdodc AdoDet
End Sub
