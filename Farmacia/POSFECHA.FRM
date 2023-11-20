VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form PosFecha 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CHEQUES POSFECHADOS"
   ClientHeight    =   5070
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10515
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   10515
   ShowInTaskbar   =   0   'False
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "PosFecha.frx":0000
      Height          =   4740
      Left            =   0
      OleObjectBlob   =   "PosFecha.frx":0015
      TabIndex        =   0
      Top             =   315
      Width           =   10515
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Salir"
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
      Left            =   9135
      TabIndex        =   2
      Top             =   0
      Width           =   1380
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Imprimir"
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
      Left            =   7770
      TabIndex        =   1
      Top             =   0
      Width           =   1380
   End
   Begin VB.Data DataAbonos 
      Caption         =   "Cancelaciones con cheques posfechados"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Width           =   7785
   End
End
Attribute VB_Name = "PosFecha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command2_Click()
  SQLMsg1 = "REPORTE DE CHEQUES POSFECHADOS"
  SQLMsg2 = ""
  ImprimirData DataAbonos, True, 1, 8, 0
End Sub

Private Sub Command3_Click()
  Unload PosFecha
End Sub

Private Sub Form_Activate()
   sSQL = "SELECT Fecha,TP,Cliente,Telefono,Factura,Banco,Cheq_No,Cheque As Valor_Cheq "
   sSQL = sSQL & "FROM Cheques_Posf,Clientes "
   sSQL = sSQL & "WHERE Cheques_Posf.Codigo_C = Clientes.Codigo "
   sSQL = sSQL & "AND T = '" & Procesado & "' "
   sSQL = sSQL & "ORDER BY Fecha,TP,Factura "
   SelectDBGrid DBGrid1, 0, DataAbonos, sSQL
   RatonNormal MDIFact
   If DataAbonos.Recordset.RecordCount <= 0 Then Unload PosFecha
End Sub

Private Sub Form_Load()
   CentrarForm PosFecha
   'Abriendo bases
   DataAbonos.DatabaseName = RutaEmpresa & "\PRODUCC.MDB"
End Sub
