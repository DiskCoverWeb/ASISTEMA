VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form FTransfCaja 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "DEBITAR / ACREDITAR CAJA"
   ClientHeight    =   2385
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6300
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   6300
   ShowInTaskbar   =   0   'False
   Begin VB.Data DataCaja 
      Caption         =   "Caja"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2100
      Visible         =   0   'False
      Width           =   2430
   End
   Begin VB.Data DataFlujoCajaEfec 
      Caption         =   "FlujoCajaEfec"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2415
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   2430
   End
   Begin VB.Data DataUsuarios 
      Caption         =   "Ususarios"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   2430
   End
   Begin MSDBCtls.DBCombo DBCUsuario 
      Bindings        =   "FTraCaja.frx":0000
      DataSource      =   "DataUsuarios"
      Height          =   315
      Left            =   315
      TabIndex        =   1
      Top             =   735
      Width           =   4425
      _ExtentX        =   7805
      _ExtentY        =   556
      _Version        =   327680
      Text            =   "DBCombo1"
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
   Begin VB.CommandButton Command2 
      Caption         =   "Grabar"
      Height          =   960
      Left            =   4935
      Picture         =   "FTraCaja.frx":0017
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   525
      Width           =   1065
   End
   Begin VB.TextBox TextCant 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   360
      Left            =   2730
      MaxLength       =   14
      MultiLine       =   -1  'True
      TabIndex        =   9
      Text            =   "FTraCaja.frx":0459
      Top             =   1575
      Width           =   2010
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   315
      TabIndex        =   2
      Top             =   1155
      Width           =   1275
      Begin VB.OptionButton OpcI 
         Caption         =   "Ingreso"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   105
         TabIndex        =   3
         Top             =   210
         Value           =   -1  'True
         Width           =   1065
      End
      Begin VB.OptionButton OpcE 
         Caption         =   "Egreso"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   105
         TabIndex        =   4
         Top             =   525
         Width           =   960
      End
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   1680
      TabIndex        =   5
      Top             =   1155
      Width           =   960
      Begin VB.OptionButton OpcMN 
         Caption         =   "MN"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   210
         TabIndex        =   7
         Top             =   525
         Width           =   645
      End
      Begin VB.OptionButton OpcME 
         Caption         =   "ME"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   210
         TabIndex        =   6
         Top             =   210
         Value           =   -1  'True
         Width           =   645
      End
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " USUARIO A DEBITAR/ACREDITAR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   315
      TabIndex        =   0
      Top             =   420
      Width           =   4425
   End
   Begin VB.Label Label16 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Cantidad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2730
      TabIndex        =   8
      Top             =   1260
      Width           =   2010
   End
End
Attribute VB_Name = "FTransfCaja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub OpcE_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub OpcI_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub OpcME_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub OpcMN_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub Command2_Click()
  Mensajes = "Desea Grabar"
  Titulo = "Formulario de Grabacion"
  If BoxMensaje = 6 Then
     Debe = 0: Haber = 0
     If OpcI.Value Then Debe = Round_ME(Val(TextCant.Text)) Else Haber = Round_ME(Val(TextCant.Text))
     sSQL = "SELECT * FROM Transacciones "
     SelectData DataCaja, sSQL, False
     With DataCaja.Recordset
         .AddNew
         .Fields("T") = Normal
         .Fields("Fecha") = FechaSistema
         .Fields("Cuenta_No") = Ninguno
         .Fields("TP") = "BOVE"
         .Fields("Debitos") = Debe
         .Fields("Creditos") = Haber
         .Fields("Usuario") = NombreUsuario
         .Fields("ID") = 0
         .Fields("Item") = NumEmpresa
         .Fields("Cheque") = Ninguno
         .Update
     End With
  End If
End Sub

Private Sub Form_Activate()
  sSQL = "SELECT Usuario FROM Comprobantes " _
       & "GROUP BY Usuario "
  SelectDBCombo DBCUsuario, DataUsuarios, sSQL, "Usuario", False
  RatonNormal
End Sub

Private Sub Form_Load()
  CentrarForm FTransfCaja
  DataCaja.DatabaseName = RutaEmpresa & "\CAJACRED.MDB"
  DataUsuarios.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
  DataFlujoCajaEfec.DatabaseName = RutaEmpresa & "\CAJACRED.MDB"
End Sub

Private Sub TextCant_GotFocus()
  TextCant.Text = ""
End Sub

Private Sub TextCant_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextCant_LostFocus()
  TextoValido TextCant, True
End Sub
