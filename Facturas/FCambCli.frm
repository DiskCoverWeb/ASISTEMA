VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Begin VB.Form FCambioCliente 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cambio de Cliente"
   ClientHeight    =   3810
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7170
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   7170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc AdoListCtas 
      Height          =   330
      Left            =   105
      Top             =   3465
      Visible         =   0   'False
      Width           =   6030
      _ExtentX        =   10636
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
      Caption         =   "ListCtas"
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
      Height          =   750
      Left            =   6195
      Picture         =   "FCambCli.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   105
      Width           =   855
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
      Height          =   750
      Left            =   6195
      Picture         =   "FCambCli.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   945
      Width           =   855
   End
   Begin MSDataListLib.DataCombo DCCliente 
      Bindings        =   "FCambCli.frx":0BD4
      DataSource      =   "AdoListCtas"
      Height          =   315
      Left            =   105
      TabIndex        =   0
      Top             =   420
      Width           =   6000
      _ExtentX        =   10583
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      ForeColor       =   8388608
      Text            =   "Cliente"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " NOMBRE DEL CLIENTE"
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
      TabIndex        =   1
      Top             =   105
      Width           =   6000
   End
End
Attribute VB_Name = "FCambioCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
  Unload FCambioCliente
End Sub

Private Sub Command2_Click()
  sSQL = "UPDATE Facturas " _
       & "SET CodigoC = '" & CodigoCliente & "' " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND Factura = " & Factura_No & " " _
       & "AND TC = '" & TipoDoc & "' "
  Ejecutar_SQL_SP sSQL
  
  sSQL = "UPDATE Detalle_Factura " _
       & "SET CodigoC = '" & CodigoCliente & "' " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND Factura = " & Factura_No & " " _
       & "AND TC = '" & TipoDoc & "' "
  Ejecutar_SQL_SP sSQL
  
  sSQL = "UPDATE Trans_Abonos " _
       & "SET CodigoC = '" & CodigoCliente & "' " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND Factura = " & Factura_No & " " _
       & "AND TP = '" & TipoDoc & "' "
  Ejecutar_SQL_SP sSQL
  Unload FCambioCliente
  RatonNormal
End Sub

Private Sub DCCliente_LostFocus()
  CodigoCliente = "9999999999"
  With AdoListCtas.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Cliente = '" & DCCliente & "' ")
       If Not .EOF Then
          CodigoCliente = .fields("Codigo")
          NombreCliente = .fields("Cliente")
       End If
   End If
  End With
  If CodigoCliente = Ninguno Then CodigoCliente = "9999999999"
End Sub

Private Sub Form_Activate()
  sSQL = "SELECT * " _
       & "FROM Clientes " _
       & "WHERE Codigo <> '.' " _
       & "AND FA <> " & Val(adFalse) & " " _
       & "ORDER BY Cliente "
  SelectDB_Combo DCCliente, AdoListCtas, sSQL, "Cliente"
  CodigoCliente = "9999999999"
  Label2.Caption = "FACTURA No. (" & TipoDoc & ") " & Format$(Factura_No, "0000000")
  RatonNormal
End Sub

Private Sub Form_Load()
  CentrarForm FCambioCliente
  ConectarAdodc AdoListCtas
End Sub
