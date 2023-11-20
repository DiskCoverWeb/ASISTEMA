VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FReservaciones 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Asignacion de Clientes"
   ClientHeight    =   6345
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9840
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   9840
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   960
      Left            =   105
      TabIndex        =   14
      Top             =   735
      Width           =   8415
      Begin MSDataListLib.DataCombo DCCliente 
         Bindings        =   "Reservar.frx":0000
         DataSource      =   "AdoDBCliente"
         Height          =   315
         Left            =   105
         TabIndex        =   15
         Top             =   525
         Width           =   6345
         _ExtentX        =   11192
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.Label LabelCodigo 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   6405
         TabIndex        =   11
         Top             =   525
         Width           =   1905
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CODIGO:"
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
         Left            =   6405
         TabIndex        =   12
         Top             =   210
         Width           =   1905
      End
      Begin VB.Label Label4 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " &Cliente:"
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
         TabIndex        =   13
         Top             =   210
         Width           =   6315
      End
   End
   Begin VB.CommandButton Command6 
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
      Left            =   8610
      Picture         =   "Reservar.frx":001B
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   840
      Width           =   1065
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6105
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9825
      _ExtentX        =   17330
      _ExtentY        =   10769
      _Version        =   393216
      TabHeight       =   1041
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&1.- REGISTRO DE HABITACIONES"
      TabPicture(0)   =   "Reservar.frx":0A11
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "LabelPorcDesc"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label25"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label24"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label23"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label22"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "LabelNo_Hab"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "TextCant"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "TextValorUnit"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Command8"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "DCHabitacion"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "DGAcomp"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).ControlCount=   11
      TabCaption(1)   =   "&2.- HABITACIONES OCUPADAS"
      TabPicture(1)   =   "Reservar.frx":0D2B
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DGDetAcomp"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "&3.- LISTADO DE RESERVACIONES"
      TabPicture(2)   =   "Reservar.frx":1045
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "DGHabitacion"
      Tab(2).ControlCount=   1
      Begin MSDataGridLib.DataGrid DGHabitacion 
         Bindings        =   "Reservar.frx":135F
         Height          =   3735
         Left            =   -71640
         TabIndex        =   19
         Top             =   1920
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   6588
         _Version        =   393216
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
               LCID            =   2058
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
               LCID            =   2058
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
      Begin MSDataGridLib.DataGrid DGDetAcomp 
         Bindings        =   "Reservar.frx":137B
         Height          =   3495
         Left            =   -71640
         TabIndex        =   18
         Top             =   1920
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   6165
         _Version        =   393216
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
               LCID            =   2058
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
               LCID            =   2058
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
      Begin MSDataGridLib.DataGrid DGAcomp 
         Bindings        =   "Reservar.frx":1395
         Height          =   3615
         Left            =   3600
         TabIndex        =   17
         Top             =   1800
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   6376
         _Version        =   393216
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
               LCID            =   2058
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
               LCID            =   2058
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
      Begin MSDataListLib.DataCombo DCHabitacion 
         Bindings        =   "Reservar.frx":13AC
         DataSource      =   "AdoDBHabitacion"
         Height          =   315
         Left            =   105
         TabIndex        =   16
         Top             =   2040
         Width           =   3480
         _ExtentX        =   6138
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Grabar Registro de Habitacion"
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
         Left            =   105
         MaskColor       =   &H80000010&
         Picture         =   "Reservar.frx":13CA
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   4200
         Width           =   3480
      End
      Begin VB.TextBox TextValorUnit 
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
         TabIndex        =   3
         Text            =   "0"
         Top             =   2940
         Width           =   1065
      End
      Begin VB.TextBox TextCant 
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
         TabIndex        =   2
         Text            =   "0"
         Top             =   3360
         Width           =   1065
      End
      Begin VB.Label LabelNo_Hab 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "."
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
         TabIndex        =   6
         Top             =   3780
         Width           =   1065
      End
      Begin VB.Label Label22 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " HABITACION LIBRE"
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
         TabIndex        =   10
         Top             =   1785
         Width           =   3480
      End
      Begin VB.Label Label23 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " VALOR DE HABITACION"
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
         TabIndex        =   9
         Top             =   2940
         Width           =   2430
      End
      Begin VB.Label Label24 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " NUMERO DE PERSONAS"
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
         TabIndex        =   8
         Top             =   3360
         Width           =   2430
      End
      Begin VB.Label Label25 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " HABITACION No."
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
         TabIndex        =   7
         Top             =   3780
         Width           =   2430
      End
      Begin VB.Label LabelPorcDesc 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Porcentaje de Descuento"
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
         Top             =   2520
         Width           =   3480
      End
   End
   Begin MSAdodcLib.Adodc AdoAcomp 
      Height          =   330
      Left            =   120
      Top             =   5040
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
      Caption         =   "Acomp"
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
      Left            =   120
      Top             =   4680
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
   Begin MSAdodcLib.Adodc AdoTipoCli 
      Height          =   330
      Left            =   120
      Top             =   4320
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
      Caption         =   "TipoCli"
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
   Begin MSAdodcLib.Adodc AdoDBHabitacion 
      Height          =   330
      Left            =   120
      Top             =   3960
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
      Caption         =   "DBHabitacion"
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
   Begin MSAdodcLib.Adodc AdoFormaPago 
      Height          =   330
      Left            =   120
      Top             =   3600
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
      Caption         =   "FormaPago"
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
   Begin MSAdodcLib.Adodc AdoDBReserva 
      Height          =   330
      Left            =   120
      Top             =   3240
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
      Caption         =   "DBReserva"
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
   Begin MSAdodcLib.Adodc AdoReserva 
      Height          =   330
      Left            =   120
      Top             =   2880
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
      Caption         =   "Reserva"
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
   Begin MSAdodcLib.Adodc AdoCliente 
      Height          =   330
      Left            =   120
      Top             =   2520
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
      Caption         =   "Cliente"
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
   Begin MSAdodcLib.Adodc AdoDBCliente 
      Height          =   330
      Left            =   120
      Top             =   2160
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
      Caption         =   "DBCliente"
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
   Begin MSAdodcLib.Adodc AdoHabitacion 
      Height          =   330
      Left            =   120
      Top             =   5400
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
      Caption         =   "Habitacion"
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
   Begin MSAdodcLib.Adodc AdoDetAcomp 
      Height          =   330
      Left            =   120
      Top             =   1800
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
      Caption         =   "DetAcomp"
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
Attribute VB_Name = "FReservaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command4_Click()
   GrabarCliente True, False
   LlenarCliente DCCliente.Text
End Sub

Private Sub Command6_Click()
   Unload FReservaciones
End Sub

Private Sub Command7_Click()
   GrabarCliente False, True
   LlenarCliente DCCliente.Text
End Sub

Private Sub Command8_Click()
  Contador = 0
  sSQL = "SELECT * FROM Acompañantes "
  SelectData AdoAcomp, sSQL
  
  Mensajes = "Seguro de asignar la habitación"
  Titulo = "Pregunta de Grabación"
  TipoDeCaja = 4 + 32
  J = MsgBox(Mensajes, TipoDeCaja, Titulo)
  If J = 6 Then
     sSQL = "SELECT * FROM DetA_comp "
     SelectData AdoDetAcomp, sSQL
     With AdoDetAcomp.Recordset
      If .RecordCount > 0 Then
          Do While Not .EOF
             'MsgBox .Fields("Acompañante")
             AdoAcomp.Recordset.AddNew
             AdoAcomp.Recordset.Fields("Factura_No") = 0
             AdoAcomp.Recordset.Fields("No_Hab") = LabelNo_Hab.Caption
             AdoAcomp.Recordset.Fields("Acompañante") = .Fields("Acompañante")
             AdoAcomp.Recordset.Update
             Contador = Contador + 1
            .MoveNext
          Loop
      End If
     End With
     sSQL = "DELETE * FROM Det_Acomp "
     ConectarAdoExecute sSQL
     sSQL = "SELECT * FROM Det_Acomp "
     SelectData AdoDetAcomp, sSQL
     sSQL = "UPDATE Habitaciones "
     sSQL = sSQL & "SET Ocupada = True, "
     sSQL = sSQL & "CodigoC = '" & UCase(LabelCodigo.Caption) & "', "
     sSQL = sSQL & "Ingreso = #" & FechaSistema & "#, "
     sSQL = sSQL & "No_Pers = " & Contador + 1 & ", "
     sSQL = sSQL & "Valor = " & CDbl(TextValorUnit.Text) & " "
     sSQL = sSQL & "WHERE No_Hab = '" & LabelNo_Hab.Caption & "' "
     ConectarAdoExecute sSQL
     GrabarCliente False, False
  End If
  LlenarCliente DCCliente.Text
End Sub

Private Sub DCCliente_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCCliente_LostFocus()
  DCCliente.Text = UCase(DCCliente.Text)
  If AdoDBCliente.Recordset.RecordCount > 0 Then
     LlenarCliente DCCliente.Text
  End If
End Sub

Private Sub DCHabitacion_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCHabitacion_LostFocus()
  LabelPorcDesc.Caption = "Porcentaje de descuento"
  sSQL = "SELECT * FROM Habitaciones "
  sSQL = sSQL & "WHERE No_Hab = '" & SinEspaciosIzq(DBCHabitacion.Text) & "' "
  SelectData AdoAux, sSQL
  With AdoAux.Recordset
   If .RecordCount > 0 Then
       TextValorUnit.Text = Round(.Fields("Valor_Hab") - ((.Fields("Valor_Hab") * Total_Desc) / 100))
       LabelNo_Hab.Caption = .Fields("No_Hab")
       LabelPorcDesc.Caption = "Porcentaje de descuento: " & Format(Round(Total_Desc), "#,##0.00")
   End If
  End With
End Sub

Private Sub Form_Activate()
  Grabar = True
  sSQL = "SELECT No_Hab,Acompañante FROM Acompañantes "
  sSQL = sSQL & "WHERE Factura_No = 0 "
  sSQL = sSQL & "ORDER BY No_Hab,Acompañante "
  SelectDataGrid DGAcomp, AdoAcomp, sSQL
  sSQL = "SELECT * FROM Det_Acomp "
  SelectDataGrid DGDetAcomp, AdoDetAcomp, sSQL
  sSQL = "DELETE * FROM Det_Acomp "
  ConectarAdoExecute sSQL
  sSQL = "SELECT No_Hab & '   ' & Descripcion As No_Habit "
  sSQL = sSQL & "FROM Habitaciones "
  sSQL = sSQL & "WHERE Ocupada = False "
  sSQL = sSQL & "ORDER BY No_Hab "
  SelectDBCombo DCHabitacion, AdoDBHabitacion, sSQL, "No_Habit", False
  sSQL = "SELECT Cliente,No_Hab,Descripcion,Ingreso,C.Codigo "
  sSQL = sSQL & "FROM Habitaciones As H,Clientes As C "
  sSQL = sSQL & "WHERE H.CodigoC = C.Codigo "
  sSQL = sSQL & "AND Ocupada = True "
  sSQL = sSQL & "ORDER BY Cliente "
  SelectDataGrid DGHabitacion, AdoHabitacion, sSQL
  sSQL = "SELECT Cliente " _
       & "FROM Clientes ORDER BY Cliente "
  SelectDBCombo DCCliente, AdoDBCliente, sSQL, "Cliente", False
  LlenarCliente DCCliente.Text
  RatonNormal
End Sub

Private Sub Form_Load()
CentrarForm FReservaciones
'Abriendo bases relacionadas
ConectarAdodc AdoAux
ConectarAdodc AdoAcomp
ConectarAdodc AdoDetAcomp
ConectarAdodc AdoCliente
ConectarAdodc AdoReserva
ConectarAdodc AdoDBReserva
ConectarAdodc AdoFormaPago
ConectarAdodc AdoDBCliente
ConectarAdodc AdoTipoCli
ConectarAdodc AdoHabitacion
ConectarAdodc AdoDBHabitacion
End Sub

Private Sub TextCant_LostFocus()
  TextoValido TextCant, True
End Sub

Public Sub LlenarCliente(Clientes As String)
   If Clientes = "" Then Clientes = Ninguno
   sSQL = "SELECT * FROM Clientes " _
        & "WHERE Cliente = '" & Clientes & "' "
   SelectData AdoCliente, sSQL
   With AdoCliente.Recordset
    If .RecordCount > 0 Then
        DCCliente.Text = .Fields("Cliente")
        'TextApellidos.Text = .Fields("Apellidos")
'        TextEmpresa.Text = .Fields("Empresa")
        LabelCodigo.Caption = .Fields("Codigo")
'        TextDireccion.Text = .Fields("Direccion")
  '       MBoxTelef1.Text = .Fields("Telefono")
'        MBoxTelef2.Text = .Fields("Celular")
'        MBoxFAX.Text = .Fields("FAX")
'        MBoxRUC.Text = .Fields("RUC_CI")
'        TextCiudad.Text = .Fields("Ciudad")
'LabelNo_Hab.Caption = .Fields("Cant_Hab")
'        MBoxRUCE.Text = .Fields("RUC_CIE")
'        TextCiudad.Text = .Fields("Ciudad")
'        TextCiudadE.Text = .Fields("Ciudad_E")
'        TextPais.Text = .Fields("Pais")
'        TextPaisE.Text = .Fields("Pais_E")
'        MBoxFecha.Text = .Fields("Fecha_N")
'        TextDias.Text = .Fields("Dias")
'        TextCantHab.Text = .Fields("Cant_Hab")
'        TextTarj.Text = .Fields("Tarjeta")
'        TextTarjCred.Text = .Fields("Tarjeta_No")
        '.Fields ("CC"), TipoDoc
'        TextDirEmp.Text = .Fields("Dir_Emp")
'        TextProfesion.Text = .Fields("Profesion")
    Else
       LabelCodigo.Caption = "Codigo"
       TextEmpresa.Text = ""
       TextDireccion.Text = ""
       MBoxTelef1.Text = "00-000-000"
       MBoxTelef2.Text = "00-000-000"
       MBoxFAX.Text = "00-000-000"
       TextCliente.Text = ""
       TextApellidos.Text = ""
       TextEmpresa.Text = ""
       MBoxRUC.Text = "000000000-0-000"
       LabelNo_Hab.Caption = Ninguno
       MBoxRUCE.Text = "000000000-0-000"
       TextTarj.Text = ""
       TextTarjCred.Text = ""
       '.Fields ("CC"), TipoDoc
       TextDirEmp.Text = ""
       TextProfesion.Text = ""
       TextCiudad.Text = ""
       TextCiudadE.Text = ""
       TextPais.Text = ""
       TextPaisE.Text = ""
       TextCliente.SetFocus
    End If
   End With
   'sSQL = "SELECT * FROM Clientes "
   'sSQL = sSQL & "WHERE R = True "
   'SelectDataGrid DBGReserva, DataReserva, sSQL
   sSQL = "SELECT No_Hab,Acompañante FROM Acompañantes "
   sSQL = sSQL & "WHERE Factura_No = 0 "
   sSQL = sSQL & "ORDER BY No_Hab,Acompañante "
   SelectDataGrid DGAcomp, AdoAcomp, sSQL
End Sub

Private Sub TextValorUnit_LostFocus()
  TextoValido TextValorUnit, True
End Sub

Public Sub GrabarCliente(Reservas As Boolean, Solo_Cli As Boolean)
   TipoDoc = SinEspaciosIzq(DCTipoCli.Text)
   Mensajes = "Esta Seguro que desea grabar el Cliente: " & Chr(13) & "["
   Mensajes = Mensajes & TextCliente.Text & " " & TextApellidos.Text & "]"
   Titulo = "Pregunta de Grabación"
   TipoDeCaja = 4 + 32
   If Reservas = False And Solo_Cli = False Then
      J = 6
   Else
      J = MsgBox(Mensajes, TipoDeCaja, Titulo)
   End If
   If J = 6 Then
   sSQL = "SELECT * FROM Clientes "
   sSQL = sSQL & "WHERE Apellidos & ' ' & Nombres = '" & DCCliente.Text & "' "
   SelectData AdoCliente, sSQL
   With AdoCliente.Recordset
    If .RecordCount > 0 Then
       .Edit
        Nuevo = False
    Else
       .AddNew
        Numero = ReadSetDataNum("Clientes", True, True)
        LabelCodigo.Caption = FormatoCodigo(TextApellidos & " " & TextCliente.Text, Numero)
        SetField AdoCliente, "Codigo", UCase(LabelCodigo.Caption)
        SetField AdoCliente, "R", False
        Nuevo = True
    End If
    SetField AdoCliente, "Nombres", UCase(TextCliente.Text)
    SetField AdoCliente, "Apellidos", UCase(TextApellidos.Text)
    SetField AdoCliente, "Empresa", UCase(TextEmpresa.Text)
    SetField AdoCliente, "Direccion", TextDireccion.Text
    SetField AdoCliente, "Telefono", MBoxTelef1.Text
    SetField AdoCliente, "Celular", MBoxTelef2.Text
    SetField AdoCliente, "FAX", MBoxFAX.Text
    SetField AdoCliente, "RUC_CI", MBoxRUC.Text
    SetField AdoCliente, "RUC_CIE", MBoxRUCE.Text
    SetField AdoCliente, "Ciudad", TextCiudad.Text
    SetField AdoCliente, "Ciudad_E", TextCiudadE.Text
    SetField AdoCliente, "Pais", TextPais.Text
    SetField AdoCliente, "Pais_E", TextPaisE.Text
    SetField AdoCliente, "Fecha_N", MBoxFecha.Text
    SetField AdoCliente, "Dias", TextDias.Text
    SetField AdoCliente, "Cant_Hab", TextCantHab.Text
    SetField AdoCliente, "Tarjeta", TextTarj.Text
    SetField AdoCliente, "Tarjeta_No", TextTarjCred.Text
    SetField AdoCliente, "CC", TipoDoc
    SetField AdoCliente, "Dir_Emp", TextDirEmp.Text
    SetField AdoCliente, "Profesion", TextProfesion.Text
    If Reservas Then
       SetField AdoCliente, "R", True
    Else
       SetField AdoCliente, "R", False
    End If
   .Update
    sSQL = "SELECT No_Hab & '   ' & Descripcion As No_Habit "
    sSQL = sSQL & "FROM Habitaciones "
    sSQL = sSQL & "WHERE Ocupada = False "
    sSQL = sSQL & "ORDER BY No_Hab "
    SelectDBCombo DCHabitacion, AdoDBHabitacion, sSQL, "No_Habit", False
    sSQL = "SELECT Apellidos,Nombres,No_Hab,Descripcion,Ingreso,Salida,Codigo "
    sSQL = sSQL & "FROM Habitaciones As H,Clientes As C "
    sSQL = sSQL & "WHERE H.CodigoC = C.Codigo "
    sSQL = sSQL & "AND Ocupada = True "
    sSQL = sSQL & "ORDER BY Apellidos,Nombres "
    SelectDataGrid DGHabitacion, AdoHabitacion, sSQL
   End With
   End If
   If Nuevo Then
      sSQL = "SELECT Apellidos & ' ' & Nombres As Cliente " _
           & "FROM Clientes ORDER BY Apellidos,Nombres "
      SelectDBCombo DCCliente, AdoDBCliente, sSQL, "Cliente", False
   End If
   sSQL = "SELECT * FROM Clientes "
   sSQL = sSQL & "WHERE R = True "
   SelectDataGrid DGReserva, AdoReserva, sSQL
End Sub

