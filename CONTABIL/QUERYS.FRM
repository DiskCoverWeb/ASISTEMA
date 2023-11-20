VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form Querys 
   Caption         =   "Presentacion de una base de datos:"
   ClientHeight    =   7410
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11370
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7410
   ScaleWidth      =   11370
   Begin VB.CommandButton Command2 
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
      Left            =   1470
      TabIndex        =   2
      Top             =   0
      Width           =   1590
   End
   Begin VB.CommandButton Command1 
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
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1485
   End
   Begin MSDBGrid.DBGrid DBGridQuery 
      Bindings        =   "Querys.frx":0000
      Height          =   6735
      Left            =   0
      OleObjectBlob   =   "Querys.frx":0014
      TabIndex        =   0
      Top             =   315
      Width           =   11145
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
      Left            =   0
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7035
      Width           =   11145
   End
End
Attribute VB_Name = "Querys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
  MousePointer = vbHourglass
  DBGridQuery.Visible = False
  ImprimirData DataQuery, True, 1, 8, False
  SelectDBGrid DBGridQuery, False, DataQuery, sSQL
  DBGridQuery.Visible = True
  MousePointer = vbDefault
End Sub

Private Sub Command2_Click()
  Unload Querys
End Sub

Private Sub Form_Activate()
  MousePointer = vbHourglass
  SelectDBGrid DBGridQuery, False, DataQuery, sSQL
  Querys.DBGridQuery.Caption = SQLMsg1
  MousePointer = vbDefault
End Sub

Private Sub Form_Load()
  CentrarForm Querys
  DataQuery.DatabaseName = PathDataBase
End Sub

