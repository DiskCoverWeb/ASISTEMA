VERSION 5.00
Begin VB.Form RespaldoTotal 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Espere un momento....     Estoy procesando las bases"
   ClientHeight    =   2385
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5250
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   5250
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      Height          =   960
      Left            =   2625
      Picture         =   "RespaldT.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   0
      Width           =   1275
   End
   Begin VB.Data DataOld 
      Caption         =   "Old"
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
      Left            =   2100
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1785
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Data DataAux 
      Caption         =   "Aux"
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
      Left            =   1785
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   945
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Data DataAct 
      Caption         =   "Act"
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
      Top             =   945
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Restaurar Bases"
      Height          =   960
      Left            =   1365
      Picture         =   "RespaldT.frx":0282
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   0
      Width           =   1275
   End
   Begin VB.TextBox TextUnidad 
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
      Left            =   1995
      MaxLength       =   1
      TabIndex        =   1
      Text            =   "A"
      Top             =   1365
      Width           =   330
   End
   Begin VB.DirListBox Dir1 
      Height          =   540
      Left            =   3465
      TabIndex        =   6
      Top             =   945
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.FileListBox File1 
      Height          =   870
      Left            =   3465
      TabIndex        =   5
      Top             =   1470
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.Data DataQuery 
      Caption         =   "Query"
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
      Left            =   105
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1785
      Visible         =   0   'False
      Width           =   2010
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Respaldar Bases"
      Height          =   960
      Left            =   105
      Picture         =   "RespaldT.frx":06C4
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   0
      Width           =   1275
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Unidad Destino:"
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
      TabIndex        =   0
      Top             =   1365
      Width           =   1905
   End
End
Attribute VB_Name = "RespaldoTotal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
  RatonReloj
  RespaldoTotal.Caption = "Copiando bases requeridas"
  Dir1.Path = RutaEmpresa
  File1.filename = Dir1.Path & "\*.MDB"
  For Ind = 0 To File1.ListCount - 1
      RespaldoTotal.Caption = "Copando la base: " & File1.List(Ind)
      RutaOrigen = Dir1.Path & "\" & File1.List(Ind)
      RutaDestino = RutaBackup & "\" & File1.List(Ind)
      FileCopy RutaOrigen, RutaDestino
  Next Ind
  RespaldoTotal.Caption = "Espere un momento, compactando las bases..."
  RutaOrigen = RutaSistema & "\Empresas.mdb"
  RutaDestino = RutaBackup & "\Empresas.mdb"
  FileCopy RutaOrigen, RutaDestino
  RutaOrigen = RutaSistema & "\pkzip.exe"
  RutaDestino = RutaBackup & "\pkzip.exe"
  FileCopy RutaOrigen, RutaDestino
  RutaOrigen = RutaSistema & "\pkunzip.exe"
  RutaDestino = RutaBackup & "\pkunzip.exe"
  FileCopy RutaOrigen, RutaDestino
  Dir1.Path = RutaBackup
  File1.filename = Dir1.Path & "\*.MDB"
  For Ind = 0 To File1.ListCount - 1
      Cadena = Dir1.Path & "\" & File1.List(Ind)
      CompactarDatabase Cadena, False
  Next Ind
  RatonNormal
  Shell RutaBackup & "\respaldo.bat " & TextUnidad.Text & ": ", vbMaximizedFocus
  RespaldoTotal.Caption = "RESPALDO TOTAL DE LAS BASES"
End Sub

Private Sub Command2_Click()
  Unload RespaldoTotal
End Sub

Private Sub Command3_Click()
  RespaldoTotal.Caption = "RESPALDO TOTAL DE LAS BASES"
  ChDrive TextUnidad.Text & ":"
  Shell "restaura.bat " & Mid(RutaSistema, 1, 1) & ": ", vbMaximizedFocus
  ChDrive Mid(RutaSistema, 1, 2)
  
  
  RespaldoTotal.Caption = "Copiando bases requeridas"
  Dir1.Path = RutaEmpresa
  File1.filename = Dir1.Path & "\*.MDB"
  For Ind = 0 To File1.ListCount - 1
      RespaldoTotal.Caption = "Copando la base: " & File1.List(Ind)
      RutaOrigen = RutaBackup & "\" & File1.List(Ind)
      RutaDestino = Dir1.Path & "\" & File1.List(Ind)
      FileCopy RutaOrigen, RutaDestino
  Next Ind
End Sub

Private Sub Form_Activate()
  RespaldoTotal.Caption = "RESPALDO TOTAL DE LAS BASES"
  RutaBackup = Left(RutaEmpresa, 2) & "\SYSBASES"
  Dir1.Path = RutaEmpresa
  File1.filename = Dir1.Path & "\*.MDB"
  RatonNormal
End Sub

Private Sub Form_Load()
  CentrarForm RespaldoTotal
End Sub

Private Sub TextUnidad_GotFocus()
  MarcarTexto TextUnidad
End Sub

Private Sub TextUnidad_LostFocus()
  TextUnidad.Text = UCase(TextUnidad.Text)
End Sub
