VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Begin VB.Form Aprobacion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PRESTAMO"
   ClientHeight    =   5700
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9990
   DrawStyle       =   1  'Dash
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   9990
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8925
      Picture         =   "FPrestam.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   1995
      Width           =   960
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Imprimir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8925
      Picture         =   "FPrestam.frx":09F6
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1050
      Width           =   960
   End
   Begin MSDataGridLib.DataGrid DGTabla 
      Bindings        =   "FPrestam.frx":12C0
      Height          =   3585
      Left            =   105
      TabIndex        =   17
      Top             =   1575
      Width           =   8730
      _ExtentX        =   15399
      _ExtentY        =   6324
      _Version        =   393216
      AllowUpdate     =   0   'False
      ForeColor       =   192
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
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
            LCID            =   12298
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
            LCID            =   12298
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
   Begin VB.TextBox TxtRUCS 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1365
      MaxLength       =   15
      TabIndex        =   16
      Text            =   "0"
      Top             =   420
      Width           =   1275
   End
   Begin MSAdodcLib.Adodc AdoTabla 
      Height          =   330
      Left            =   315
      Top             =   3255
      Visible         =   0   'False
      Width           =   2115
      _ExtentX        =   3731
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
      Caption         =   "Tabla"
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
   Begin VB.TextBox TxtNombresS 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2625
      MaxLength       =   60
      TabIndex        =   3
      Top             =   420
      Width           =   5895
   End
   Begin VB.TextBox TextTP 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   105
      MaxLength       =   30
      TabIndex        =   9
      Top             =   1155
      Width           =   4740
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Aceptar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8925
      Picture         =   "FPrestam.frx":12D7
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   105
      Width           =   960
   End
   Begin VB.TextBox TextInt 
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
      Left            =   4935
      MaxLength       =   49
      MultiLine       =   -1  'True
      TabIndex        =   10
      Top             =   1155
      Width           =   750
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
      Left            =   6615
      MaxLength       =   49
      MultiLine       =   -1  'True
      TabIndex        =   12
      Top             =   1155
      Width           =   1905
   End
   Begin VB.TextBox TextMeses 
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
      Left            =   5775
      MaxLength       =   49
      MultiLine       =   -1  'True
      TabIndex        =   11
      Top             =   1155
      Width           =   750
   End
   Begin MSMask.MaskEdBox MBFechaI 
      Height          =   330
      Left            =   105
      TabIndex        =   1
      ToolTipText     =   "Formato de Fecha: DD/MM/AA"
      Top             =   420
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   582
      _Version        =   393216
      AllowPrompt     =   -1  'True
      AutoTab         =   -1  'True
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
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
   Begin MSAdodcLib.Adodc AdoCta 
      Height          =   330
      Left            =   315
      Top             =   3570
      Visible         =   0   'False
      Width           =   2115
      _ExtentX        =   3731
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
      Caption         =   "Cta"
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
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
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
      Left            =   4830
      TabIndex        =   21
      Top             =   5250
      Width           =   1905
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Total Interés"
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
      Left            =   3465
      TabIndex        =   20
      Top             =   5250
      Width           =   1380
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
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
      TabIndex        =   19
      Top             =   5250
      Width           =   1905
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Total Capital"
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
      TabIndex        =   18
      Top             =   5250
      Width           =   1380
   End
   Begin VB.Label Label15 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Fecha Apro."
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
      Top             =   105
      Width           =   1275
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Nombres"
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
      Left            =   2625
      TabIndex        =   2
      Top             =   105
      Width           =   5895
   End
   Begin VB.Label Label9 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TIPO DE PRESTAMO"
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
      TabIndex        =   5
      Top             =   840
      Width           =   4740
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Interes"
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
      Left            =   4935
      TabIndex        =   6
      Top             =   840
      Width           =   750
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Monto de Prestamo"
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
      Left            =   6615
      TabIndex        =   8
      Top             =   840
      Width           =   1905
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Meses"
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
      Left            =   5775
      TabIndex        =   7
      Top             =   840
      Width           =   750
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " C.I./R.U.C."
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
      Left            =   1365
      TabIndex        =   4
      Top             =   105
      Width           =   1275
   End
End
Attribute VB_Name = "Aprobacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
  SQLMsg1 = UCase$(Nombre_Cta_Ret)
  SQLMsg2 = NombreCliente
  SQLMsg3 = MBFechaI.Text & ", Monto " & TextMonto.Text & ", Plazo: " & TextMeses.Text & ", Interes: " & TextInt.Text
  ImprimirAdodc AdoTabla, 1, 9
End Sub

Private Sub Command2_Click()
 SumatoriaSC = 0
 Unload Aprobacion
End Sub

Private Sub Command3_Click()
 With AdoTabla.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       SumatoriaSC = .Fields("Saldo")
      .MoveNext
       LnSC_No = 1
       Do While Not .EOF
          Debe = Debe + .Fields("Capital")
          Haber = Haber + .Fields("Interes")
          SetAdoAddNew "Asiento_SC"
          SetAdoFields "Prima", 0
          SetAdoFields "FECHA_V", .Fields("Fecha")
          SetAdoFields "Factura", Val(Year(.Fields("Fecha")) & Format(Month(.Fields("Fecha")), "00") & Format(Day(.Fields("Fecha")), "00"))
          SetAdoFields "Valor", .Fields("Pagos")
          SetAdoFields "Codigo", CodigoCliente
          SetAdoFields "Beneficiario", NombreCliente
          SetAdoFields "Cta", SubCtaGen
          SetAdoFields "TM", OpcTM
          SetAdoFields "TC", "C"
          SetAdoFields "Prima", 0
          SetAdoFields "DH", OpcDH
          SetAdoFields "Valor_ME", 0
          SetAdoFields "Item", NumEmpresa
          SetAdoFields "T_No", Trans_No
          SetAdoFields "SC_No", LnSC_No
          SetAdoFields "CodigoU", CodigoUsuario
          SetAdoUpdate
          LnSC_No = LnSC_No + 1
         .MoveNext
       Loop
      .MoveFirst
   End If
 End With
 Unload Aprobacion
End Sub

Private Sub DGTabla_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF1 Then GenerarDataTexto Aprobacion, AdoTabla
End Sub

Private Sub Form_Activate()
   TextTP.Text = Nombre_Cta_Ret
   MBFechaI.Text = FechaTexto
   Mifecha = BuscarFecha(FechaTexto)
   TxtRUCS.Text = CodigoCliente
   TxtNombresS.Text = NombreCliente
   If Bloquear_Control Then Command3.Enabled = False
   RatonNormal
   TextInt.SetFocus
End Sub

Private Sub Form_Load()
   CentrarForm Aprobacion
   ConectarAdodc AdoCta
   ConectarAdodc AdoTabla
End Sub

Private Sub MBFechaI_GotFocus()
  MarcarTexto MBFechaI
End Sub

Private Sub MBFechaI_LostFocus()
  FechaValida MBFechaI
End Sub

Private Sub TextInt_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextMeses_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextMonto_GotFocus()
  SumatoriaSC = 0
End Sub

Private Sub TextMonto_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextMonto_LostFocus()
  TextoValido TextMonto, True
  TextoValido TextMeses, True
  TextoValido TextInt, True
  Total = Round(CCur(TextMonto.Text), 2)
  Numero = Round(CCur(TextMeses.Text), 2)
  Interes = Round(CCur(TextInt.Text) / 100, 4)
  GenerarTablaPrestamoCxCP MBFechaI.Text, AdoTabla, DGTabla, TextInt, TextMeses, TextMonto, SubCtaGen, True
  Debe = 0: Haber = 0
  With AdoTabla.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       Do While Not .EOF
          Debe = Debe + .Fields("Capital")
          Haber = Haber + .Fields("Interes")
         .MoveNext
       Loop
      .MoveFirst
   End If
  End With
  Label2.Caption = Format(Debe, "#,##0.00")
  Label10.Caption = Format(Haber, "#,##0.00")
End Sub
