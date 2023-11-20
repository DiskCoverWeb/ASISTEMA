VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Gastos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Gastos de Caja"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5280
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   5280
   ShowInTaskbar   =   0   'False
   Begin VB.Data DataGastos 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2940
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   105
      Visible         =   0   'False
      Width           =   2115
   End
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
      Height          =   435
      Left            =   3990
      TabIndex        =   9
      Top             =   1365
      Width           =   1170
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Grabar"
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
      Left            =   3990
      TabIndex        =   8
      Top             =   840
      Width           =   1170
   End
   Begin VB.TextBox TextMonto 
      Alignment       =   1  'Right Justify
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
      Left            =   2100
      MaxLength       =   12
      MultiLine       =   -1  'True
      TabIndex        =   5
      Text            =   "Gastos.frx":0000
      Top             =   1050
      Width           =   1380
   End
   Begin VB.TextBox TextConcepto 
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
      MaxLength       =   30
      TabIndex        =   7
      Text            =   "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
      Top             =   1785
      Width           =   3795
   End
   Begin MSMask.MaskEdBox MBoxFecha 
      Height          =   330
      Left            =   2100
      TabIndex        =   3
      Top             =   630
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   582
      _Version        =   327680
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
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "000000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   105
      Width           =   1110
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Valor del Gasto S/."
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
      TabIndex        =   4
      Top             =   1050
      Width           =   2010
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Concepto:"
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
      TabIndex        =   6
      Top             =   1470
      Width           =   3795
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Recibo No."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   1485
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Fecha (DD/MM/AA):"
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
      TabIndex        =   2
      Top             =   630
      Width           =   2010
   End
End
Attribute VB_Name = "Gastos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
  sSQL = "SELECT * FROM GastosDia "
  SelectData DataGastos, sSQL, False
  Mensajes = "Seguro Desea grabar el Recibo"
  Titulo = "Pregunta de Impresion"
  TipoDeCaja = 4 + 32: J = MsgBox(Mensajes, TipoDeCaja, Titulo)
  If J = 6 Then
     NumComp = ReadSetDataNum("Recibo", True, True)
     With DataGastos.Recordset
          .AddNew
          .Fields("Recibo") = NumComp
          .Fields("Fecha") = MBoxFecha.Text
          .Fields("Concepto") = TextConcepto.Text
          .Fields("Monto") = Val(TextMonto.Text)
          .Update
     End With
  End If
  Mensajes = "Desea imprimir el Recibo"
  Titulo = "Pregunta de Impresion"
  TipoDeCaja = 4 + 32: J = MsgBox(Mensajes, TipoDeCaja, Titulo)
  If J = 6 Then
     sSQL = "SELECT * FROM GastosDia "
     sSQL = sSQL & "WHERE Recibo =" & NumComp & " "
     SelectData DataGastos, sSQL, False
     SQLMsg1 = "GASTOS DEL DIA"
     ImprimirData DataGastos, True, 1, 10
  End If
  Unload Gastos
End Sub

Private Sub Command2_Click()
  Unload Gastos
End Sub

Private Sub Form_Activate()
   NumComp = ReadSetDataNum("Recibo", True, False)
   Label5.Caption = Format(NumComp, "000000")
   RatonNormal
End Sub

Private Sub Form_Load()
   CentrarForm Gastos
   'Abriendo bases relacionadas
   DataGastos.DatabaseName = RutaEmpresa & "\GASTOS.MDB"
   DataCodigos.DatabaseName = RutaEmpresa & "\SETEOS.MDB"
End Sub

Private Sub MBoxFecha_LostFocus()
  FechaValida MBoxFecha, False
End Sub

Private Sub TextConcepto_GotFocus()
  TextConcepto.Text = ""
End Sub

Private Sub TextMonto_GotFocus()
  TextMonto.Text = ""
End Sub
