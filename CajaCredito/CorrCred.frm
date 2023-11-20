VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Begin VB.Form CorreccionCredito 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6975
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11040
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6975
   ScaleWidth      =   11040
   Begin VB.Frame Frame2 
      Height          =   1905
      Left            =   105
      TabIndex        =   0
      Top             =   0
      Width           =   10830
      Begin VB.Frame Frame1 
         Caption         =   "Por:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   4200
         TabIndex        =   13
         Top             =   105
         Width           =   2955
         Begin VB.TextBox TxtCred 
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
            TabIndex        =   17
            Top             =   525
            Width           =   1380
         End
         Begin VB.OptionButton OpcCta 
            Caption         =   "C&uenta No."
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
            Left            =   1470
            TabIndex        =   16
            Top             =   210
            Width           =   1380
         End
         Begin VB.OptionButton OpcCred 
            Caption         =   "C&redito No."
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
            TabIndex        =   15
            Top             =   210
            Value           =   -1  'True
            Width           =   1380
         End
         Begin MSMask.MaskEdBox MBoxCuenta 
            Height          =   330
            Left            =   1575
            TabIndex        =   14
            ToolTipText     =   "Formato de Fecha: DD/MM/AA"
            Top             =   525
            Visible         =   0   'False
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   582
            _Version        =   393216
            ForeColor       =   192
            AllowPrompt     =   -1  'True
            AutoTab         =   -1  'True
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "CCCCCCCC-C"
            Mask            =   "########-#"
            PromptChar      =   "0"
         End
      End
      Begin VB.TextBox TxtNombresS 
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
         Left            =   5460
         MaxLength       =   30
         TabIndex        =   8
         Top             =   1470
         Width           =   5265
      End
      Begin VB.TextBox TxtRazonSocial 
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
         Left            =   5460
         MaxLength       =   49
         TabIndex        =   9
         Top             =   1155
         Width           =   5265
      End
      Begin MSDataListLib.DataList DLTipo 
         Bindings        =   "CorrCred.frx":0000
         DataSource      =   "AdoTipo"
         Height          =   1230
         Left            =   105
         TabIndex        =   10
         Top             =   525
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   2170
         _Version        =   393216
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
      Begin VB.CommandButton Command3 
         Caption         =   "&Consultar"
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
         Left            =   8715
         Picture         =   "CorrCred.frx":0016
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   210
         Width           =   960
      End
      Begin VB.CommandButton Command4 
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
         Height          =   855
         Left            =   9765
         Picture         =   "CorrCred.frx":0458
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   210
         Width           =   960
      End
      Begin VB.OptionButton OpcM 
         Caption         =   "Con &Modificación"
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
         Left            =   7245
         TabIndex        =   3
         Top             =   630
         Width           =   1485
      End
      Begin VB.OptionButton OpcL 
         Caption         =   "Solo &Lectura"
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
         Left            =   7245
         TabIndex        =   2
         Top             =   210
         Value           =   -1  'True
         Width           =   1485
      End
      Begin VB.Label Label3 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Cliente"
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
         Left            =   4200
         TabIndex        =   6
         Top             =   1470
         Width           =   1275
      End
      Begin VB.Label Label25 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Represent."
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
         Left            =   4200
         TabIndex        =   7
         Top             =   1155
         Width           =   1275
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " &TIPO DE PRESTAMO"
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
         Top             =   210
         Width           =   4005
      End
   End
   Begin MSDataGridLib.DataGrid DGSaldos 
      Bindings        =   "CorrCred.frx":0E4E
      Height          =   3375
      Left            =   105
      TabIndex        =   12
      Top             =   3465
      Width           =   10830
      _ExtentX        =   19103
      _ExtentY        =   5953
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
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
   Begin MSDataGridLib.DataGrid DGCredito 
      Bindings        =   "CorrCred.frx":0E66
      Height          =   1485
      Left            =   105
      TabIndex        =   11
      Top             =   1995
      Width           =   10830
      _ExtentX        =   19103
      _ExtentY        =   2619
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
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
   Begin MSAdodcLib.Adodc AdoCredito 
      Height          =   330
      Left            =   525
      Top             =   3885
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
      Caption         =   "Credito"
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
   Begin MSAdodcLib.Adodc AdoSaldos 
      Height          =   330
      Left            =   525
      Top             =   4200
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
      Caption         =   "Saldos"
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
   Begin MSAdodcLib.Adodc AdoTipo 
      Height          =   330
      Left            =   525
      Top             =   4515
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
      Caption         =   "Tipo"
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
   Begin MSAdodcLib.Adodc AdoAux 
      Height          =   330
      Left            =   525
      Top             =   4830
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
      Caption         =   "Aux"
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
End
Attribute VB_Name = "CorreccionCredito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command3_Click()
  TextoValido TxtCred
  RatonReloj
  If OpcL.value Then
     DGCredito.AllowAddNew = False
     DGCredito.AllowDelete = False
     DGCredito.AllowUpdate = False
     DGSaldos.AllowAddNew = False
     DGSaldos.AllowDelete = False
     DGSaldos.AllowUpdate = False
  Else
     DGCredito.AllowAddNew = True
     DGCredito.AllowDelete = True
     DGCredito.AllowUpdate = True
     DGSaldos.AllowAddNew = True
     DGSaldos.AllowDelete = True
     DGSaldos.AllowUpdate = True
  End If
  Codigo = SinEspaciosIzq(DLTipo.Text)
  Codigo1 = Ninguno
  sSQL = "SELECT * " _
       & "FROM Clientes_Datos_Extras " _
       & "WHERE Cuenta_No = '" & MBoxCuenta.Text & "' " _
       & "AND Tipo_Dato = 'LIBRETAS' "
  SelectAdodc AdoAux, sSQL
  With AdoAux.Recordset
   If .RecordCount > 0 Then Codigo1 = .Fields("Codigo")
  End With
' Datos del Socio
  sSQL = "SELECT * " _
       & "FROM Clientes " _
       & "WHERE Codigo = '" & Codigo1 & "' "
  SelectAdodc AdoAux, sSQL
  With AdoAux.Recordset
   If .RecordCount > 0 Then
       TxtRazonSocial.Text = .Fields("Representante")
       TxtNombresS.Text = .Fields("Cliente")
   End If
  End With
' Encabezados de Creditos
  sSQL = "SELECT * FROM Prestamos "
  If OpcCta.value Then
     sSQL = sSQL & "WHERE Cuenta_No = '" & MBoxCuenta.Text & "' "
  Else
     sSQL = sSQL & "WHERE Credito_No = '" & TxtCred.Text & "' "
  End If
  sSQL = sSQL & "AND TP = '" & Codigo & "' " _
       & "ORDER BY Credito_No,Fecha "
  SelectDataGrid DGCredito, AdoCredito, sSQL
' Tabla de Amortizacion
  sSQL = "SELECT * FROM Trans_Prestamos "
  If OpcCta.value Then
     sSQL = sSQL & "WHERE Cuenta_No = '" & MBoxCuenta.Text & "' "
  Else
     sSQL = sSQL & "WHERE Credito_No = '" & TxtCred.Text & "' "
  End If
  sSQL = sSQL & "AND TP = '" & Codigo & "' " _
       & "ORDER BY Credito_No,Fecha,Dia "
  SelectDataGrid DGSaldos, AdoSaldos, sSQL
  RatonNormal
End Sub

Private Sub Command4_Click()
  Unload CorreccionCredito
End Sub

Private Sub Form_Activate()
  'MsgBox NumEmpresa
  sSQL = "SELECT (TP & '   ' & Descripcion) As TipoPrest " _
       & "FROM Catalogo_Prestamo " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "ORDER BY TP,Descripcion "
  'MsgBox sSQL
  SelectDBList DLTipo, AdoTipo, sSQL, "TipoPrest"
  RatonNormal
End Sub

Private Sub Form_Load()
  CentrarForm CorreccionCredito
  ConectarAdodc AdoAux
  ConectarAdodc AdoTipo
  ConectarAdodc AdoSaldos
  ConectarAdodc AdoCredito
End Sub

Private Sub OpcCred_Click()
  MBoxCuenta.Visible = False
  TxtCred.Visible = True
  TxtCred.SetFocus
End Sub

Private Sub OpcCta_Click()
  TxtCred.Visible = False
  MBoxCuenta.Visible = True
  MBoxCuenta.SetFocus
End Sub

Private Sub TxtCred_GotFocus()
  MarcarTexto TxtCred
End Sub
