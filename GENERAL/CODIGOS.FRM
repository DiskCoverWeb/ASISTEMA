VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form FCodigo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Codigos"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5370
   Icon            =   "Codigos.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   5370
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "SALIR DE CODIGOS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2835
      TabIndex        =   1
      Top             =   3045
      Width           =   2325
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "Codigos.frx":0442
      Height          =   1275
      Left            =   105
      OleObjectBlob   =   "Codigos.frx":0458
      TabIndex        =   0
      Top             =   105
      Width           =   5160
   End
   Begin VB.Data DataCodigos 
      Caption         =   "Codigos"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   105
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2205
      Visible         =   0   'False
      Width           =   2745
   End
End
Attribute VB_Name = "FCodigo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
  Unload FCodigo
End Sub

Private Sub Form_Activate()
   Select Case TipoSeteos
     Case "COD":
          sSQL = "SELECT * FROM Codigos "
          sSQL = sSQL & "ORDER BY ID "
          FCodigo.Caption = "CODIGOS DE NUMERACION"
          Command1.Caption = "SALIR DE CODIGOS"
     Case "SET":
          sSQL = "SELECT * FROM CtasProceso "
          sSQL = sSQL & "ORDER BY Item "
          FCodigo.Caption = "SETEOS DE CUENTAS GENERALES"
          Command1.Caption = "SALIR DE SETEOS"
     Case Else
          sSQL = ""
   End Select
   SelectData DataCodigos, sSQL, False
End Sub

Private Sub Form_Load()
   CentrarForm FCodigo
   DataCodigos.DatabaseName = RutaEmpresa & "\SETEOS.MDB"
End Sub
