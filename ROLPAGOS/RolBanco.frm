VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.5#0"; "comctl32.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Begin VB.Form FRolPagoBanco 
   BackColor       =   &H00FF8080&
   Caption         =   "BANCO BOLIVARIANO"
   ClientHeight    =   8670
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11475
   Icon            =   "RolBanco.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "RolBanco.frx":0442
   ScaleHeight     =   8670
   ScaleWidth      =   11475
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Actualizar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   11550
      Picture         =   "RolBanco.frx":4F7A
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   1155
      Width           =   1275
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "TIPO DE ARCHIVO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   4830
      TabIndex        =   15
      Top             =   1155
      Width           =   3900
      Begin VB.OptionButton Opc_10_4to 
         BackColor       =   &H00FFC0C0&
         Caption         =   "10mo. 4to."
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
         Left            =   2520
         TabIndex        =   18
         Top             =   210
         Width           =   1275
      End
      Begin VB.OptionButton Opc_10_3ro 
         BackColor       =   &H00FFC0C0&
         Caption         =   "10mo. 3ro."
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
         Left            =   1155
         TabIndex        =   17
         Top             =   210
         Width           =   1275
      End
      Begin VB.OptionButton OpcSueldo 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Sueldo"
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
         TabIndex        =   16
         Top             =   210
         Value           =   -1  'True
         Width           =   960
      End
   End
   Begin MSDataGridLib.DataGrid DGClientes 
      Bindings        =   "RolBanco.frx":5844
      Height          =   5475
      Left            =   2730
      TabIndex        =   14
      Top             =   1890
      Width           =   7155
      _ExtentX        =   12621
      _ExtentY        =   9657
      _Version        =   393216
      AllowUpdate     =   0   'False
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
         Weight          =   400
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
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFC0C0&
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
      Height          =   645
      Left            =   10185
      Picture         =   "RolBanco.frx":585E
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1155
      Width           =   1275
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Enviar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   8820
      Picture         =   "RolBanco.frx":6254
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1155
      Width           =   1275
   End
   Begin VB.FileListBox File1 
      BackColor       =   &H00FFC0C0&
      Height          =   2820
      Left            =   105
      TabIndex        =   13
      Top             =   4515
      Width           =   2535
   End
   Begin ComctlLib.ProgressBar ProgBarra 
      Height          =   330
      Left            =   105
      TabIndex        =   12
      Top             =   7350
      Width           =   9780
      _ExtentX        =   17251
      _ExtentY        =   582
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.DriveListBox Drive1 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   105
      TabIndex        =   9
      Top             =   2100
      Width           =   2535
   End
   Begin VB.DirListBox Dir1 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1890
      Left            =   105
      TabIndex        =   10
      Top             =   2415
      Width           =   2535
   End
   Begin MSAdodcLib.Adodc AdoAux 
      Height          =   330
      Left            =   3255
      Top             =   4305
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
   Begin MSAdodcLib.Adodc AdoDetalle 
      Height          =   330
      Left            =   3255
      Top             =   3045
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
      Caption         =   "Detalle"
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
   Begin MSAdodcLib.Adodc AdoProducto 
      Height          =   330
      Left            =   3255
      Top             =   3990
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
      Caption         =   "Producto"
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
   Begin MSMask.MaskEdBox MBFechaI 
      Height          =   330
      Left            =   105
      TabIndex        =   1
      Top             =   1470
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   16761024
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
   Begin MSAdodcLib.Adodc AdoAbono 
      Height          =   330
      Left            =   3255
      Top             =   3675
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
      Caption         =   "Abono"
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
   Begin MSMask.MaskEdBox MBFechaF 
      Height          =   330
      Left            =   1680
      TabIndex        =   3
      Top             =   1470
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   16761024
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
   Begin MSAdodcLib.Adodc AdoPendiente 
      Height          =   330
      Left            =   3255
      Top             =   2415
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
      Caption         =   "Pendiente"
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
   Begin MSAdodcLib.Adodc AdoClientes 
      Height          =   330
      Left            =   3255
      Top             =   2730
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
      Caption         =   "Clientes"
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
   Begin MSAdodcLib.Adodc AdoFactura 
      Height          =   330
      Left            =   3255
      Top             =   3360
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
   Begin MSMask.MaskEdBox MBFechaP 
      Height          =   330
      Left            =   3255
      TabIndex        =   5
      Top             =   1470
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   16761024
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
   Begin VB.Label Label2 
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Fecha de &Pago"
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
      Left            =   3255
      TabIndex        =   4
      Top             =   1155
      Width           =   1485
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Fecha Final"
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
      Left            =   1680
      TabIndex        =   2
      Top             =   1155
      Width           =   1485
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " ARCHIVO:"
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
      TabIndex        =   11
      Top             =   4305
      Width           =   2535
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &Fecha Inicio"
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
      Top             =   1155
      Width           =   1485
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &ORIGEN"
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
      TabIndex        =   8
      Top             =   1890
      Width           =   2535
   End
End
Attribute VB_Name = "FRolPagoBanco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim AdoStrCnnOld As String
Dim AdoStrCnn1 As String

Dim NumFile As Integer
Dim NumFileAbonos As Integer
Dim NumFileDetalle As Integer
Dim NumFileAlumnos As Integer
Dim NumFileFacturas As Integer
Dim NumFileProducto As Integer

Dim RutaGeneraFile As String
Dim RutaGeneraFileAbonos As String
Dim RutaGeneraFileDetalle As String
Dim RutaGeneraFileAlumnos As String
Dim RutaGeneraFileFacturas As String
Dim RutaGeneraFileProducto As String

Dim XAdoStrCnn As String
Dim IJ As Long
Dim ModuloResp As String
Dim RetVal
Dim RutaBackupXX As String

Private Sub Command1_Click()
  RatonReloj
  FechaValida MBFechaI
  FechaValida MBFechaF
  FechaIni = BuscarFecha(MBFechaI)
  FechaFin = BuscarFecha(MBFechaF)
  sSQL = "SELECT Codigo, MIN(Fecha_D) As FechaMin, MAX(Fecha_H) As FechaMax, COUNT(Codigo) As NoMeses, SUM(Egresos) As Total_Dec " _
       & "FROM Trans_Rol_de_Pagos " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Fecha_D >= #" & FechaIni & "# " _
       & "AND Fecha_H <= #" & FechaFin & "# " _
       & " "
  If Opc_10_3ro.value Then sSQL = sSQL & "AND Cod_Rol_Pago = 'Decimo_III' "
  If Opc_10_4to.value Then sSQL = sSQL & "AND Cod_Rol_Pago = 'Decimo_IV' "
  sSQL = sSQL _
       & "AND Egresos > 0 " _
       & "GROUP BY Codigo"
  Select_Adodc AdoAux, sSQL
  With AdoAux.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          
         'Dias = CFechaLong(.Fields("FechaMax")) - CFechaLong(.Fields("FechaMin"))
          Dias = .fields("NoMeses") * 30
          sSQL = "UPDATE Catalogo_Rol_Pagos "
          ' Dias_Dec_3ro,
          If Opc_10_3ro.value Then
             sSQL = sSQL _
                  & "SET Valor_Dec_3ro = " & .fields("Total_Dec") & ", " _
                  & "Dias_Dec_3ro = " & Dias & " "
          End If
          If Opc_10_4to.value Then
             sSQL = sSQL _
                  & "SET Valor_Dec_4to = " & .fields("Total_Dec") & ", " _
                  & "Dias_Dec_4to = " & Dias & " "
          End If
          sSQL = sSQL _
               & "WHERE Item = '" & NumEmpresa & "' " _
               & "AND Periodo = '" & Periodo_Contable & "' " _
               & "AND Codigo = '" & .fields("Codigo") & "' "
          Ejecutar_SQL_SP sSQL
          'If .Fields("Codigo") = "1710241991" Then MsgBox sSQL
         .MoveNext
       Loop
   End If
  End With
  RatonNormal
  MsgBox "Proceso Exitoso"
End Sub

Private Sub Command2_Click()
  Unload FRolPagoBanco
End Sub

Private Sub Command4_Click()
Dim Cont As Integer
Dim CaptionTemp As String
  CaptionTemp = FRolPagoBanco.Caption
  FechaValida MBFechaI
  FechaValida MBFechaF
  FechaValida MBFechaP
  TipoDoc = ""
  AuxNumEmp = NumEmpresa
  FechaIni = BuscarFecha(MBFechaI)
  FechaFin = BuscarFecha(MBFechaF)
  FechaTexto = BuscarFecha(MBFechaP)
  FechaTexto1 = Format(MBFechaP, "MM/dd/yyyy")
 'Detalle de las Facturas Emitidas del mes
  If OpcSueldo.value Then
     sSQL = "SELECT C.Cliente,C.TD,C.CI_RUC, R.Egresos As Neto_Recibir,CR.Cta_Transferencia,CR.Acreditar_Cta," _
          & "C.Prov,C.Ciudad,C.Direccion,C.Sexo,C.Telefono,C.TelefonoT,C.Celular," _
          & "C.Email,C.Fecha_N,C.Codigo,R.Grupo_Rol,CR.Codigo_Banco " _
          & "FROM Trans_Rol_de_Pagos As R, Clientes As C, Catalogo_Rol_Pagos As CR " _
          & "WHERE R.Fecha_D >= #" & FechaIni & "# " _
          & "AND R.Fecha_H <= #" & FechaFin & "# " _
          & "AND R.Item = '" & NumEmpresa & "' " _
          & "AND R.Periodo = '" & Periodo_Contable & "' " _
          & "AND CR.T = '" & Normal & "' " _
          & "AND R.Tipo_Rubro = 'PER' " _
          & "AND R.Cod_Rol_Pago = 'Neto_Recibir' " _
          & "AND R.Codigo = C.Codigo " _
          & "AND R.Codigo = CR.Codigo " _
          & "AND R.Item = CR.Item " _
          & "AND R.Periodo = CR.Periodo " _
          & "ORDER BY C.Cliente,R.Codigo "
  ElseIf Opc_10_3ro.value Then
     sSQL = "SELECT C.Cliente,C.CI_RUC,C.TD,CR.Valor_Dec_3ro,CR.Cta_Transferencia,CR.Acreditar_Cta," _
          & "C.Prov,C.Ciudad,C.Direccion,C.Sexo,C.Telefono,C.TelefonoT,C.Celular," _
          & "C.Email,C.Fecha_N,C.Codigo,CR.Grupo_Rol,CR.Codigo_Banco " _
          & "FROM Clientes As C, Catalogo_Rol_Pagos As CR " _
          & "WHERE CR.Item = '" & NumEmpresa & "' " _
          & "AND CR.Periodo = '" & Periodo_Contable & "' " _
          & "AND CR.Valor_Dec_3ro > 0 " _
          & "AND CR.Codigo = C.Codigo " _
          & "ORDER BY C.Cliente,CR.Codigo "
  Else
     sSQL = "SELECT C.Cliente,C.CI_RUC,C.TD,CR.Valor_Dec_4to,CR.Cta_Transferencia,CR.Acreditar_Cta," _
          & "C.Prov,C.Ciudad,C.Direccion,C.Sexo,C.Telefono,C.TelefonoT,C.Celular," _
          & "C.Email,C.Fecha_N,C.Codigo,CR.Grupo_Rol,CR.Codigo_Banco " _
          & "FROM Clientes As C, Catalogo_Rol_Pagos As CR " _
          & "WHERE CR.Item = '" & NumEmpresa & "' " _
          & "AND CR.Periodo = '" & Periodo_Contable & "' " _
          & "AND CR.Valor_Dec_4to > 0 " _
          & "AND CR.Codigo = C.Codigo " _
          & "ORDER BY C.Cliente,CR.Codigo "
  End If
  Select_Adodc_Grid DGClientes, AdoClientes, sSQL
  DGClientes.Visible = False
  Select Case TextoBanco
    Case "PICHINCHA"
         FechaFin = BuscarFecha(UltimoDiaMes(MBFechaF))
         Generar_Pichincha
    Case "INTERNACIONAL"
         FechaFin = BuscarFecha(UltimoDiaMes(MBFechaF))
         Generar_Internacional
    Case "BOLIVARIANO"
         FechaFin = BuscarFecha(UltimoDiaMes(MBFechaF))
         Generar_Bolivariano
    Case "PACIFICO"
         FechaFin = BuscarFecha(UltimoDiaMes(MBFechaF))
         Generar_Pacifico
    Case "GUAYAQUIL"
         Generar_Guayaquil
    Case Else
         FechaFin = BuscarFecha(UltimoDiaMes(MBFechaF))
         MsgBox "No esta definido este Banco"
  End Select
  DGClientes.Visible = True
  FRolPagoBanco.Caption = CaptionTemp
End Sub

Private Sub DGClientes_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  If KeyCode = vbKeyF1 Then GenerarDataTexto FRolPagoBanco, AdoClientes, True
  If CtrlDown And (vbKeyP = KeyCode) Then ImprimirAdodc AdoClientes, 2, 7, True
End Sub

Private Sub Dir1_Change()
  File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
  Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_DblClick()
  SiguienteControl
End Sub

Private Sub File1_KeyDown(KeyCode As Integer, Shift As Integer)
  NombreArchivo = File1.Filename
  If KeyCode = vbKeyDelete Then
     Mensajes = "Esta seguro de Eliminar: " & File1.Filename
     Titulo = "Pregunta de Eliminacion"
     If BoxMensaje = vbYes Then Kill File1.Path & "\" & File1.Filename
     File1.Filename = Dir1.Path & "\*.*"
  End If
End Sub

Private Sub File1_LostFocus()
Dim MaxCar As Integer
  NombreArchivo = UCase(File1.Filename)
  
  RutaGeneraFile = UCase(Dir1.Path & "\" & NombreArchivo)
  If NombreArchivo <> "" Then
     RatonReloj
     MaxCar = 0
    'MsgBox RutaGeneraFile
     NumFile = FreeFile: TxtFile = ""
     Open RutaGeneraFile For Input As #NumFile
       Do While Not EOF(NumFile)
          Line Input #NumFile, Cod_Field
          If Len(Cod_Field) > MaxCar Then MaxCar = Len(Cod_Field)
          TxtFile = TxtFile & Cod_Field & vbCrLf
       Loop
     Close #NumFile
     J = 1: K = 0
     Cadena = ""
     Cadena1 = ""
     For I = 1 To MaxCar
         Cadena = Cadena & CStr(J)
         J = J + 1
         If J > 9 Then
            Cadena = Cadena & "0"
            J = 1
            K = K + 1
            If K <= 10 Then
               Cadena1 = Cadena1 & String(9, " ") & CStr(K)
            Else
               Cadena1 = Cadena1 & String(8, " ") & CStr(K)
            End If
         End If
     Next I
     Cadena = Cadena & vbCrLf
     Cadena1 = Cadena1 & vbCrLf
     
     TxtFile = Cadena1 & Cadena & TxtFile
  Else
     MsgBox "Seleccione un archivo"
  End If
  RatonNormal
End Sub

Private Sub Form_Activate()
  FechaValida MBFechaI
  Label4.Caption = "ORIGEN" & Space(18) & "COD: " & CodigoDelBanco
  FRolPagoBanco.Caption = "FACTURACION DE BANCOS (" & CodigoDelBanco & ")"
 'Alumnos/Clientes que estan activados para Generar las Facturas
  Select Case TextoBanco
    Case "PICHINCHA"
         RutaOrigen = RutaSistema & "\LOGOS\PICHINCHA.GIF"
         FRolPagoBanco.BackColor = &H80FFFF
         FRolPagoBanco.Caption = "BANCO DEL PICHINCHA (" & CodigoDelBanco & ")" & String(40, " ") & "EL ULTIMO CODIGO: " & Codigo
    Case "INTERNACIONAL"
         RutaOrigen = RutaSistema & "\LOGOS\INTERNACIONAL.GIF"
         FRolPagoBanco.BackColor = &HFF8080    '&HFF0000
         FRolPagoBanco.Caption = "BANCO INTERNACIONAL (" & CodigoDelBanco & ")" & String(40, " ") & "EL ULTIMO CODIGO: " & Codigo
    Case "BOLIVARIANO"
         RutaOrigen = RutaSistema & "\LOGOS\BOLIVARIANO.GIF"
         FRolPagoBanco.BackColor = &H808000
         FRolPagoBanco.Caption = "BANCO BOLIVARIANO (" & CodigoDelBanco & ")" & String(40, " ") & "EL ULTIMO CODIGO: " & Codigo
         CheqSAT.Visible = True
    Case "PACIFICO"
         RutaOrigen = RutaSistema & "\LOGOS\PACIFICO.GIF"
         FRolPagoBanco.BackColor = &HC0C000
         FRolPagoBanco.Caption = "BANCO DEL PACIFICO (" & CodigoDelBanco & ")" & String(40, " ") & "EL ULTIMO CODIGO: " & Codigo
    Case "GUAYAQUIL":
         RutaOrigen = RutaSistema & "\LOGOS\GUAYAQUIL.GIF"
         FRolPagoBanco.BackColor = &HFF8080    '&HFF0000
         FRolPagoBanco.Caption = "BANCO DE GUAYQUIL (" & CodigoDelBanco & ")" & String(40, " ") & "EL ULTIMO CODIGO: " & Codigo
    Case Else
         RutaOrigen = RutaSistema & "\LOGOS\DISKCOVS.GIF"
         FRolPagoBanco.BackColor = &HE0E0E0
         FRolPagoBanco.Caption = "OTROS BANCOS (" & CodigoDelBanco & ")" & String(40, " ") & "EL ULTIMO CODIGO: " & Codigo
  End Select
  FRolPagoBanco.Picture = LoadPicture(RutaOrigen)
  Label1.BackColor = FRolPagoBanco.BackColor
  Label2.BackColor = FRolPagoBanco.BackColor
  Label4.BackColor = FRolPagoBanco.BackColor
  Label7.BackColor = FRolPagoBanco.BackColor
  Label8.BackColor = FRolPagoBanco.BackColor
  MBFechaI.BackColor = FRolPagoBanco.BackColor
  MBFechaF.BackColor = FRolPagoBanco.BackColor
  MBFechaP.BackColor = FRolPagoBanco.BackColor
  Frame1.BackColor = FRolPagoBanco.BackColor
    OpcSueldo.BackColor = FRolPagoBanco.BackColor
    Opc_10_3ro.BackColor = FRolPagoBanco.BackColor
    Opc_10_4to.BackColor = FRolPagoBanco.BackColor
  Dir1.BackColor = FRolPagoBanco.BackColor
  File1.BackColor = FRolPagoBanco.BackColor
  Drive1.BackColor = FRolPagoBanco.BackColor
  Dir1.Path = RutaBackup & "\ROLPAGO\"
  File1.Filename = Dir1.Path & "\*.*"
  Dir1.Refresh
  File1.Refresh
  RatonNormal
End Sub

Private Sub Form_Deactivate()
  FRolPagoBanco.WindowState = vbMaximized
End Sub

Private Sub Form_Load()
 'CentrarForm FRolPagoBanco
  RutaBackupXX = ""
  ConectarAdodc AdoAux
  ConectarAdodc AdoAbono
  ConectarAdodc AdoDetalle
  ConectarAdodc AdoFactura
  ConectarAdodc AdoClientes
  ConectarAdodc AdoProducto
  ConectarAdodc AdoPendiente
  DGClientes.width = MDI_X_Max - DGClientes.Left - 100
  DGClientes.Height = MDI_Y_Max - 2200
  ProgBarra.Left = DGClientes.Left
  ProgBarra.Top = MDI_Y_Max - 300
  ProgBarra.width = MDI_X_Max - ProgBarra.Left - 100
  Drive1.Drive = MidStrg(RutaSysBases, 1, 2)
  RutaBackup = RutaSysBases & "\BANCO"
End Sub

Private Sub MBFechaI_GotFocus()
  MarcarTexto MBFechaI
End Sub

Private Sub MBFechaI_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBFechaI_LostFocus()
  MBFechaF.Text = UltimoDiaMes(MBFechaI.Text)
  FechaValida MBFechaI
End Sub

Private Sub MBFechaF_GotFocus()
  MarcarTexto MBFechaF
End Sub

Private Sub MBFechaF_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBFechaF_LostFocus()
  MBFechaP.Text = MBFechaF.Text
  FechaValida MBFechaF
End Sub

Public Sub Generar_Bolivariano()

End Sub

Public Sub Generar_Pacifico()
Dim AuxNumEmp As String

Dim Traza As String
Dim DiaV As Integer
Dim MesV As Integer
Dim A�oV As Integer
Dim CamposFile() As Campos_Tabla
Dim Separador As String
Dim Otra_CI As String
  RatonReloj
  Separador = ","
 'Separador = vbTab
 'Abrimo los archivo que vamos ha necesitar
  Traza = UCase(MidStrg(MesesLetras(Month(MBFechaI)), 1, 3) & "-" & Year(MBFechaI))
  If OpcSueldo.value Then
     RutaGeneraFileAlumnos = UCase(Dir1.Path & "\ROL_DE_" & Traza) & ".txt"
  ElseIf Opc_10_3ro.value Then
     RutaGeneraFileAlumnos = UCase(Dir1.Path & "\DECIMO_TERCERO_" & Traza & ".TXT")
  Else
     RutaGeneraFileAlumnos = UCase(Dir1.Path & "\DECIMO_CUARTO_" & Traza & ".TXT")
  End If
  Traza = ""
  NumFileAlumnos = FreeFile
  Open RutaGeneraFileAlumnos For Output As #NumFileAlumnos
  FechaTexto = FechaSistema
  Mifecha = Format(Year(MBFechaP), "0000") _
          & Format(Month(MBFechaP), "00") _
          & Format(Day(MBFechaP), "00")
  TextoImprimio = ""
  AuxNumEmp = NumEmpresa
  Cta_Cobrar = Ninguno
  Contador = 0
' Comenzamos a generar el archivo: EMPLEADOS.TXT
  With AdoClientes.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       Do While Not .EOF
          If Len(.fields("Cta_Transferencia")) > 4 Then
             Contador = Contador + 1
             CodigoCli = .fields("Codigo")
             NombreCliente = TrimStrg(MidStrg(.fields("Cliente"), 1, 30))
             Codigo4 = TrimStrg(MidStrg(Empresa, 1, 30))
             CodigoP = Format(MidStrg(.fields("CI_RUC"), 1, 10), "0000000000") & String(5, " ")
             Codigo1 = MidStrg(SinEspaciosIzq(.fields("Cta_Transferencia")), 2, 2)
             Otra_CI = Format(MidStrg(.fields("Acreditar_Cta"), 1, 10), "0000000000") & String(5, " ")
            'MsgBox Codigo1
             Select Case Codigo1
               Case "01": Codigo1 = "00"
               Case "02": Codigo1 = "10"
               Case Else: Codigo1 = "10"
             End Select
             'MsgBox "=>>>>>> " & Codigo1
             Codigo2 = Format(Val(SinEspaciosDer(.fields("Cta_Transferencia"))), "00000000")
             
             If OpcSueldo.value Then
                Total = .fields("Neto_Recibir")
                Codigo3 = "ROL DE " & MidStrg(MesesLetras(Month(MBFechaI)), 1, 3) & "-" & CStr(Year(MBFechaI))
             ElseIf Opc_10_3ro.value Then
                 Total = .fields("Valor_Dec_3ro")
                 Codigo3 = "DEC-3ER DE " & MidStrg(MesesLetras(Month(MBFechaI)), 1, 3) & "-" & CStr(Year(MBFechaI))
             Else
                 Total = .fields("Valor_Dec_4to")
                 Codigo3 = "DEC-4TO DE " & MidStrg(MesesLetras(Month(MBFechaI)), 1, 3) & "-" & CStr(Year(MBFechaI))
             End If
            'Total = .Fields("Neto_Recibir")
             I = Int(Total)
             J = (Total - Int(Total)) * 100
           ' Empieza la trama por Alumno
             Print #NumFileAlumnos, "4";                                                  ' Localidad (1)
             Print #NumFileAlumnos, "OCP";                                                ' Transsaccion (2-4)
             Print #NumFileAlumnos, "RP";                                                 ' Codigo de Servicio (5-6)
             Print #NumFileAlumnos, Codigo1;                                              ' Tipo de Cuenta ()
             Print #NumFileAlumnos, Codigo2;                                              ' Numero de Cuenta
             Print #NumFileAlumnos, Format(I, "0000000000000") & Format(J, "00");         ' Valor
             Print #NumFileAlumnos, CodigoP;                                              ' Codigo del Alumno Identificacion Servicios
             Print #NumFileAlumnos, Codigo3 & String(20 - Len(Codigo3), " ");             ' Referencia: Codigo3
             Print #NumFileAlumnos, "CU";                                                 ' Forma de Pago
             Print #NumFileAlumnos, "USD";                                                ' Moneda
             Print #NumFileAlumnos, NombreCliente & String(30 - Len(NombreCliente), " "); ' Nombre del Empleado
             Print #NumFileAlumnos, "  ";                                                 ' Localidad
             Print #NumFileAlumnos, String(2, " ");                                       ' Agencia de Retiro
             Print #NumFileAlumnos, .fields("TD");                                        ' Tipo NUC
             If Val(Otra_CI) > 0 Then
                Print #NumFileAlumnos, Otra_CI;                                           ' Cedula Beneficiario
             Else
                Print #NumFileAlumnos, CodigoP;                                           ' Cedula Beneficiario
             End If
             Print #NumFileAlumnos, String(19, " ")                                       ' Numero de cuenta
          End If
         .MoveNext
       Loop
   End If
  End With
  Close #NumFileAlumnos
 'Finalizamos los Archivos
  ProgBarra.value = ProgBarra.Max
  RatonNormal
  MsgBox "SE GENERO EL SIGUIENTE ARCHIVO: " & vbCrLf & vbCrLf & vbCrLf _
       & RutaGeneraFileAlumnos & vbCrLf & vbCrLf
End Sub

Public Sub Generar_Guayaquil()
Dim AuxNumEmp As String
Dim Traza As String
Dim DiaV As Integer
Dim MesV As Integer
Dim A�oV As Integer
Dim CamposFile() As Campos_Tabla
Dim Separador As String
Dim NumCelular As String
Dim CodBanco As String
Dim ValorStrg As String

  RatonReloj
  Separador = ","
 'Separador = vbTab
 'Abrimo los archivo que vamos ha necesitar
  Traza = UCase(MidStrg(MesesLetras(Month(MBFechaI)), 1, 3) & "-" & Year(MBFechaI))
  If OpcSueldo.value Then
     RutaGeneraFileAlumnos = Dir1.Path & "\NCR" & BuscarFecha(FechaSistema) & "HAL_01.txt"
  ElseIf Opc_10_3ro.value Then
     RutaGeneraFileAlumnos = UCase(Dir1.Path & "\DECIMO_TERCERO_" & Traza & ".TXT")
  Else
     RutaGeneraFileAlumnos = UCase(Dir1.Path & "\DECIMO_CUARTO_" & Traza & ".TXT")
  End If
  Traza = ""
  NumFileAlumnos = FreeFile
  Open RutaGeneraFileAlumnos For Output As #NumFileAlumnos
  FechaTexto = FechaSistema
  Mifecha = Format(Year(MBFechaP), "0000") _
          & Format(Month(MBFechaP), "00") _
          & Format(Day(MBFechaP), "00")
  TextoImprimio = ""
  AuxNumEmp = NumEmpresa
  Cta_Cobrar = Ninguno
  Contador = 0
' Comenzamos a generar el archivo: EMPLEADOS.TXT
  With AdoClientes.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       Do While Not .EOF
         'C.Cliente,C.TD,C.CI_RUC, R.Egresos As Neto_Recibir,CR.Cta_Transferencia,CR.Acreditar_Cta,
         'C.Prov,C.Ciudad,C.Direccion,C.Sexo,C.Telefono,C.TelefonoT,C.Celular,
         'C.Email,C.Fecha_N,C.Codigo,R.Grupo_Rol,CR.Codigo_Banco
          If Len(.fields("Cta_Transferencia")) > 4 Then
             Contador = Contador + 1
             CodigoCli = .fields("Codigo")
             
             Codigo4 = TrimStrg(MidStrg(Empresa, 1, 30))
             CodigoP = Format(MidStrg(.fields("CI_RUC"), 1, 10), "0000000000") & "   "
             Codigo1 = MidStrg(SinEspaciosIzq(.fields("Cta_Transferencia")), 2, 2)
             Otra_CI = Format(MidStrg(.fields("Acreditar_Cta"), 1, 10), "0000000000") & String(5, " ")
            'MsgBox Codigo1
             Select Case Codigo1
               Case "01": Codigo1 = "A"
               Case "02": Codigo1 = "C"
               Case Else: Codigo1 = "M"
             End Select
             'MsgBox "=>>>>>> " & Codigo1}
             If .fields("Codigo_Banco") = 17 Then
                 Codigo2 = SinEspaciosDer(.fields("Cta_Transferencia"))
                 If 10 > Len(Codigo2) Then Codigo2 = String(10 - Len(Codigo2), "0") & Codigo2
                 Codigo3 = String(18, " ")
                 NombreCliente = String(18, " ")
                 Codigo4 = "  "
             Else
                 Codigo2 = String(10, "0")
                 Codigo3 = SinEspaciosDer(.fields("Cta_Transferencia"))
                 Codigo3 = String(18 - Len(Codigo3), "0") & Codigo3
                 NombreCliente = TrimStrg(MidStrg(.fields("Cliente"), 1, 18))
                 If 18 > Len(NombreCliente) Then NombreCliente = NombreCliente & String(18 - Len(NombreCliente), " ")
                 Codigo4 = "XX"
             End If
             Emails = MidStrg(.fields("Email"), 1, 30)
             If 30 > Len(Emails) Then Emails = Emails & String(30 - Len(Emails), " ")
             
             CodBanco = .fields("Codigo_Banco")
             If 3 > Len(CodBanco) Then CodBanco = String(3 - Len(CodBanco), "0") & CodBanco
             
             NumCelular = .fields("Celular")
             If 10 > Len(NumCelular) Then NumCelular = NumCelular & String(10 - Len(NumCelular), " ")
             
             If OpcSueldo.value Then
                ValorStrg = Format(.fields("Neto_Recibir") * 100, String(15, "0"))
             ElseIf Opc_10_3ro.value Then
                 ValorStrg = Format(.fields("Valor_Dec_3ro" * 100), String(15, "0"))
             Else
                 ValorStrg = Format(.fields("Valor_Dec_4to" * 100), String(15, "0"))
             End If
             
           ' Empieza la trama por Alumno
             Traza = Codigo1                            ' Tipo de Cuenta(1)
             Traza = Traza & Codigo2                             ' N�mero de Cuenta Banco Guayaquil
             Traza = Traza & ValorStrg                           ' ValoR
             Traza = Traza & "XX"                                ' Motivo
             Traza = Traza & "Y"                                  ' Tipo de Nota
             Traza = Traza & "01"                                ' Agencia
             Traza = Traza & Codigo4                             ' C�digo de banco para el pago interbancario
             Traza = Traza & Codigo3                             ' N�mero de Cuenta Otros Bancos
             Traza = Traza & NombreCliente                       ' Nombre del Titular de la Cuenta Otros Bancos
             Traza = Traza & "HAL"                               ' NUEVO MOTIVO
             Traza = Traza & Emails                              ' Emails
             Traza = Traza & NumCelular                          ' Celular
             Traza = Traza & CodBanco                            ' Codigo del Banco
             Traza = Traza & .fields("TD")                       ' Tipo ID
             Traza = Traza & CodigoP                             ' N�mero Identificaci�n Beneficiario
             
             Separador = Codigo1 & "=" & Len(Codigo1) & vbCrLf
             Separador = Separador & Codigo2 & "=" & Len(Codigo2) & vbCrLf
             Separador = Separador & ValorStrg & "=" & Len(ValorStrg) & vbCrLf
             Separador = Separador & "XX" & "=" & Len("XX") & vbCrLf
             Separador = Separador & "Y" & "=" & Len("Y") & vbCrLf
             Separador = Separador & "01" & "=" & Len("01") & vbCrLf
             Separador = Separador & Codigo4 & "=" & Len(Codigo4) & vbCrLf
             Separador = Separador & Codigo3 & "=" & Len(Codigo3) & vbCrLf
             Separador = Separador & NombreCliente & "=" & Len(NombreCliente) & vbCrLf
             Separador = Separador & "HAL" & "=" & Len("HAL") & vbCrLf
             Separador = Separador & Emails & "=" & Len(Emails) & vbCrLf
             Separador = Separador & NumCelular & "=" & Len(NumCelular) & vbCrLf
             Separador = Separador & CodBanco & "=" & Len(CodBanco) & vbCrLf
             Separador = Separador & .fields("TD") & "=" & Len(.fields("TD")) & vbCrLf
             Separador = Separador & CodigoP & "=" & Len(CodigoP) & vbCrLf
             
             Print #NumFileAlumnos, Traza
             
             
             
             'MsgBox Traza & vbCrLf & Separador & "(" & Len(Traza) & ")...."
          End If
         .MoveNext
       Loop
   End If
  End With
  Close #NumFileAlumnos
 'Finalizamos los Archivos
  ProgBarra.value = ProgBarra.Max
  RatonNormal
  MsgBox "SE GENERO EL SIGUIENTE ARCHIVO: " & vbCrLf & vbCrLf & vbCrLf _
       & RutaGeneraFileAlumnos & vbCrLf & vbCrLf
End Sub

Public Sub Generar_Internacional()
Dim AuxNumEmp As String

Dim Traza As String
Dim DiaV As Integer
Dim MesV As Integer
Dim A�oV As Integer
Dim CamposFile() As Campos_Tabla
Dim Separador As String
  RatonReloj
  Separador = ","
 'Separador = vbTab
 'Abrimo los archivo que vamos ha necesitar
  Traza = UCase(MidStrg(MesesLetras(Month(MBFechaI)), 1, 3))
  RutaGeneraFileAlumnos = UCase(Dir1.Path & "\EMPLEADOS_DE_" & Traza & ".TXT")
  RutaGeneraFileDetalle = UCase(Dir1.Path & "\ROL_DE_" & Traza & ".TXT")
  Traza = ""
  NumFileAlumnos = FreeFile
  Open RutaGeneraFileAlumnos For Output As #NumFileAlumnos
  FechaTexto = FechaSistema
  Mifecha = Format(Year(MBFechaP), "0000") _
          & Format(Month(MBFechaP), "00") _
          & Format(Day(MBFechaP), "00")
  TextoImprimio = ""
  AuxNumEmp = NumEmpresa
  Cta_Cobrar = Ninguno
' Comenzamos a generar el archivo: EMPLEADOS.TXT
  With AdoClientes.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       Do While Not .EOF
          Contador = Contador + 1
          CodigoCli = .fields("Codigo")
          CodigoP = Format(.fields("CI_RUC"), "0000000000")
          Codigo1 = MidStrg(SinEspaciosIzq(.fields("Cta_Transferencia")), 2, 2)
          Codigo2 = Format(Val(SinEspaciosDer(.fields("Cta_Transferencia"))), "0000000000")
          'MsgBox Codigo1 & vbCrLf & Codigo2
          If OpcSueldo.value Then
             Total = .fields("Neto_Recibir") * 100
          ElseIf Opc_10_3ro.value Then
             Total = .fields("Valor_Dec_3ro") * 100
          Else
             Total = .fields("Valor_Dec_4to") * 100
          End If
         'Total = .Fields("Neto_Recibir") * 100
          If Len(Codigo1) <= 1 Then Codigo1 = " "
         'Empieza la trama
          Traza = SetearBlancos(Codigo1, 2, 0, False) _
                & SetearBlancos(Codigo2, 10, 0, False) _
                & Format(Val(Total), "0000000000000") _
                & Format(Contador, "000000") _
                & "0"
          If Codigo1 <> "." Then Print #NumFileAlumnos, Traza
         .MoveNext
       Loop
   End If
  End With
  Close #NumFileAlumnos
' Comenzamos a generar el archivo: ROL.TXT
  Traza = ""
  Contador = 0
  NumFileAlumnos = FreeFile
  Open RutaGeneraFileDetalle For Output As #NumFileAlumnos
  With AdoClientes.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       Do While Not .EOF
          Contador = Contador + 1
          CodigoCli = .fields("Codigo")
          CodigoP = Format(.fields("CI_RUC"), "0000000000")
          'Codigo1 = MidStrg(SinEspaciosIzq(.Fields("Cta_Transferencia")), 2, 2)
          Codigo1 = SinEspaciosIzq(.fields("Cta_Transferencia"))
          Codigo2 = CStr(Val(SinEspaciosDer(.fields("Cta_Transferencia"))))
          'MsgBox Codigo1 & vbCrLf & Codigo2
             If OpcSueldo.value Then
                Total = .fields("Neto_Recibir") * 100
             ElseIf Opc_10_3ro.value Then
                 Total = .fields("Valor_Dec_3ro") * 100
             Else
                 Total = .fields("Valor_Dec_4to") * 100
             End If
          'Total = .Fields("Neto_Recibir") * 100
          If Len(Codigo1) <= 1 Then Codigo1 = " "
           
''          Select Case Codigo1
''            Case "01": Codigo1 = "CTE"
''            Case "02": Codigo1 = "AHO"
''            Case "03": Codigo1 = "VIR"
''            Case Else: Codigo1 = "XXX"
''          End Select
         'Empieza la trama
          Traza = "PA" & vbTab _
                & TrimStrg(CStr(Contador)) & vbTab _
                & "USD" & vbTab _
                & Format(Val(Total), "0000000000000") & vbTab _
                & "CTA" & vbTab _
                & Codigo1 & vbTab _
                & Codigo2 & vbTab _
                & "ROL DE " & UCase(MidStrg(MesesLetras(Month(MBFechaI)), 1, 3)) & " " & Year(MBFechaI) & vbTab _
                & .fields("TD") & vbTab _
                & .fields("CI_RUC") & vbTab _
                & TrimStrg(MidStrg(.fields("Cliente"), 1, 41)) & vbTab _
                & .fields("Codigo_Banco")
''          Traza = SetearBlancos(Codigo1, 2, 0, False) _
''                & SetearBlancos(Codigo2, 10, 0, False) _
''                & Format(Val(Total), "0000000000000") _
''                & Format(Contador, "000000") _
''                & "0" _
''                & Mifecha _
''                & "00000000" _
''                & String(20, "0") _
''                & SetearBlancos(Codigo1, 2, 0, False) _
''                & SetearBlancos(.Fields("Cliente"), 30, 0, False) _
''                & SetearBlancos(CodigoP, 10, 0, False) & "000" _
''                & "17"
          If Len(Codigo1) >= 3 Then Print #NumFileAlumnos, Traza
         .MoveNext
       Loop
   End If
  End With
  Close #NumFileAlumnos
 'Finalizamos los Archivos
  ProgBarra.value = ProgBarra.Max
  RatonNormal
  MsgBox "SE GENERARON LOS SIGUIENTES ARCHIVOS: " & vbCrLf & vbCrLf & vbCrLf _
       & RutaGeneraFileAlumnos & vbCrLf & vbCrLf _
       & RutaGeneraFileDetalle & vbCrLf & vbCrLf
End Sub

Public Sub Generar_Pichincha()
Dim AuxNumEmp As String
Dim Traza As String
Dim DiaV As Integer
Dim MesV As Integer
Dim A�oV As Integer
Dim CamposFile() As Campos_Tabla
Dim Separador As String
Dim SiImprimir As Boolean
  RatonReloj
  Separador = ","
 'Separador = vbTab
 'Abrimo los archivo que vamos ha necesitar
  Traza = UCase(MidStrg(MesesLetras(Month(MBFechaI)), 1, 3))
  If OpcSueldo.value Then
     If Day(MBFechaF.Text) <= 15 Then
        RutaGeneraFileAlumnos = UCase(Dir1.Path & "\PRIMERA_QUINCENA_DE_" & Traza & ".TXT")
     Else
        RutaGeneraFileAlumnos = UCase(Dir1.Path & "\SEGUNDA_QUINCENA_DE_" & Traza & ".TXT")
     End If
  ElseIf Opc_10_3ro.value Then
     RutaGeneraFileAlumnos = UCase(Dir1.Path & "\DECIMO_TERCERO_" & Traza & ".TXT")
  Else
     RutaGeneraFileAlumnos = UCase(Dir1.Path & "\DECIMO_CUARTO_" & Traza & ".TXT")
  End If
  
  Traza = ""
  NumFileAlumnos = FreeFile
  Open RutaGeneraFileAlumnos For Output As #NumFileAlumnos
  FechaTexto = FechaSistema
  Mifecha = Format(Year(MBFechaP), "0000") _
          & Format(Month(MBFechaP), "00") _
          & Format(Day(MBFechaP), "00")
  TextoImprimio = ""
  AuxNumEmp = NumEmpresa
  Cta_Cobrar = Ninguno
  Contador = 1
' Comenzamos a generar el archivo: EMPLEADOS.TXT
  With AdoClientes.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       Do While Not .EOF
          SiImprimir = True
          CodigoP = Format(.fields("CI_RUC"), "0000000000")
          Codigo1 = MidStrg(SinEspaciosIzq(.fields("Cta_Transferencia")), 2, 2)
          Codigo2 = SinEspaciosDer(.fields("Cta_Transferencia"))
          If Len(.fields("Cta_Transferencia")) <= 1 Then Codigo2 = Ninguno
          If OpcSueldo.value Then
             Total = .fields("Neto_Recibir") * 100
          ElseIf Opc_10_3ro.value Then
             Total = .fields("Valor_Dec_3ro") * 100
          Else
             Total = .fields("Valor_Dec_4to") * 100
          End If
          If Len(Codigo1) <= 1 Then Codigo1 = " "
          If Codigo2 = "" Then Codigo2 = " "
         'Empieza la trama
         '& SetearBlancos(CodigoP, 10, 0, False) & vbTab
          Traza = "PA" & vbTab _
                & .fields("Grupo_Rol") & vbTab _
                & "USD" & vbTab _
                & Format(Val(Total), "##0") & vbTab _
                & "CTA" & vbTab
          If Val(Codigo1) = 1 Then
             Traza = Traza & "CTE" & vbTab
          ElseIf Val(Codigo1) = 2 Then
             Traza = Traza & "AHO" & vbTab
          Else
             Traza = Traza & "VIR" & vbTab
          End If
          Traza = Traza & Codigo2 & vbTab
          If OpcSueldo.value Then
             If Day(MBFechaF.Text) <= 15 Then
                Traza = Traza & "PRIMERA QUINCENA DE " & MesesLetras(Month(MBFechaF.Text)) & " " & Year(MBFechaF.Text) & vbTab
             Else
                Traza = Traza & "SEGUNDA QUINCENA DE " & MesesLetras(Month(MBFechaF.Text)) & " " & Year(MBFechaF.Text) & vbTab
             End If
          ElseIf Opc_10_3ro.value Then
             Traza = Traza & "DECIMO TERCERO " & MesesLetras(Month(MBFechaF.Text)) & " " & Year(MBFechaF.Text) & vbTab
          Else
             Traza = Traza & "DECIMO CUARTO " & MesesLetras(Month(MBFechaF.Text)) & " " & Year(MBFechaF.Text) & vbTab
          End If
          Traza = Traza & "C" & vbTab _
                & SetearBlancos(CodigoP, 10, 0, False) & vbTab _
                & .fields("Cliente") & vbTab _
                & .fields("Codigo_Banco") & vbTab & vbTab & vbTab & vbTab
          'MsgBox Codigo2 & " - " & .Fields("Cta_Transferencia")
          'Grupo_Rol,CR.Codigo_Banco
          If Codigo2 = Ninguno Then SiImprimir = False
          If Total <= 0 Then SiImprimir = False
          If Val(.fields("Codigo_Banco")) <= 0 Then SiImprimir = False
          If Len(.fields("Cta_Transferencia")) <= 1 Then SiImprimir = False
          If SiImprimir Then Print #NumFileAlumnos, Sin_Signos_Especiales(UCase(Traza))
          Contador = Contador + 1
         .MoveNext
       Loop
   End If
  End With
  Close #NumFileAlumnos
 'Finalizamos los Archivos
  ProgBarra.value = ProgBarra.Max
  RatonNormal
  MsgBox "SE GENERARON LOS SIGUIENTES ARCHIVOS: " & vbCrLf & vbCrLf & vbCrLf _
       & RutaGeneraFileAlumnos & vbCrLf & vbCrLf
End Sub

Private Sub MBFechaP_GotFocus()
  MarcarTexto MBFechaP
End Sub

Private Sub MBFechaP_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBFechaP_LostFocus()
  FechaValida MBFechaP
End Sub

