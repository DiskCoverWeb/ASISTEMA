VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Begin VB.Form HorasEntSal 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "REGISTRO DE HORAS TRABAJADAS"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14730
   Icon            =   "HorasES.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   14730
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame4 
      Caption         =   "Movimientos de: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   105
      TabIndex        =   22
      Top             =   1785
      Width           =   13035
      Begin VB.OptionButton Opc120 
         Caption         =   "Cuatro meses"
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
         Left            =   7665
         TabIndex        =   26
         Top             =   315
         Width           =   1485
      End
      Begin VB.OptionButton Opc90 
         Caption         =   "Tres meses"
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
         Left            =   1995
         TabIndex        =   25
         Top             =   315
         Width           =   1380
      End
      Begin VB.OptionButton Opc31 
         Caption         =   "Mes actual"
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
         TabIndex        =   23
         Top             =   315
         Value           =   -1  'True
         Width           =   1380
      End
      Begin VB.OptionButton Opc60 
         Caption         =   "Dos meses"
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
         Left            =   5880
         TabIndex        =   24
         Top             =   315
         Width           =   1275
      End
      Begin VB.OptionButton OpcTodos 
         Caption         =   "Anual actual"
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
         Left            =   3885
         TabIndex        =   27
         Top             =   315
         Width           =   1485
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3375
      Left            =   105
      TabIndex        =   31
      Top             =   2625
      Width           =   14505
      _ExtentX        =   25585
      _ExtentY        =   5953
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "SUELDO"
      TabPicture(0)   =   "HorasES.frx":0696
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "LabelAbonado"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label5"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "LabelFacturado"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label6"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "DGHorasTrabajadas"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "NOVEDADES"
      TabPicture(1)   =   "HorasES.frx":06B2
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DGNovedades"
      Tab(1).ControlCount=   1
      Begin MSDataGridLib.DataGrid DGHorasTrabajadas 
         Bindings        =   "HorasES.frx":06CE
         Height          =   2430
         Left            =   105
         TabIndex        =   32
         Top             =   420
         Width           =   14295
         _ExtentX        =   25215
         _ExtentY        =   4286
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
               LCID            =   3082
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
               LCID            =   3082
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
      Begin MSDataGridLib.DataGrid DGNovedades 
         Bindings        =   "HorasES.frx":06EF
         Height          =   2850
         Left            =   -74895
         TabIndex        =   37
         ToolTipText     =   "<Insert> Novedades, <Supri> Elimina Novedades"
         Top             =   420
         Width           =   14295
         _ExtentX        =   25215
         _ExtentY        =   5027
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
               LCID            =   3082
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
               LCID            =   3082
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
      Begin VB.Label Label6 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Horas Trabajadas"
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
         TabIndex        =   36
         Top             =   2940
         Width           =   1695
      End
      Begin VB.Label LabelFacturado 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "9999.99"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   330
         Left            =   1785
         TabIndex        =   35
         Top             =   2940
         Width           =   1695
      End
      Begin VB.Label Label5 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Ingreso Liquido"
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
         TabIndex        =   34
         Top             =   2940
         Width           =   1485
      End
      Begin VB.Label LabelAbonado 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   330
         Left            =   4935
         TabIndex        =   33
         Top             =   2940
         Width           =   1485
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Eliminar Días"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   13230
      Picture         =   "HorasES.frx":070A
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   945
      Width           =   1380
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ingreso:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   1470
      TabIndex        =   2
      Top             =   0
      Width           =   5055
      Begin VB.OptionButton OpcQuincena 
         Caption         =   "Quincenal"
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
         Left            =   2415
         TabIndex        =   5
         Top             =   210
         Width           =   1275
      End
      Begin VB.OptionButton OpcSemana 
         Caption         =   "Semanal"
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
         Left            =   1155
         TabIndex        =   4
         Top             =   210
         Width           =   1170
      End
      Begin VB.OptionButton OpcDia 
         Caption         =   "Diario"
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
         Left            =   105
         TabIndex        =   3
         Top             =   210
         Value           =   -1  'True
         Width           =   960
      End
      Begin VB.OptionButton OpcMensual 
         Caption         =   "Mensual"
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
         Left            =   3780
         TabIndex        =   6
         Top             =   210
         Width           =   1170
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Generar Días"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   13230
      Picture         =   "HorasES.frx":0B4C
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   105
      Width           =   1380
   End
   Begin VB.TextBox TxtOrden 
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
      Left            =   8925
      MaxLength       =   10
      TabIndex        =   21
      Top             =   1155
      Width           =   1275
   End
   Begin VB.TextBox TxtDias 
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
      Left            =   3885
      MaxLength       =   10
      TabIndex        =   14
      Top             =   1155
      Width           =   1065
   End
   Begin VB.TextBox TxtPorcHExt 
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
      Left            =   7455
      MaxLength       =   10
      TabIndex        =   19
      Top             =   1155
      Width           =   1380
   End
   Begin VB.TextBox TxtHorasExt 
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
      Left            =   5040
      MaxLength       =   10
      TabIndex        =   16
      Top             =   1155
      Width           =   1590
   End
   Begin VB.CommandButton Command2 
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
      Height          =   750
      Left            =   13230
      Picture         =   "HorasES.frx":0E56
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   1785
      Width           =   1380
   End
   Begin VB.TextBox TxtValorHora 
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
      MaxLength       =   12
      MultiLine       =   -1  'True
      TabIndex        =   10
      Text            =   "HorasES.frx":1720
      Top             =   1155
      Width           =   1485
   End
   Begin VB.TextBox TxtHorasTrab 
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
      MaxLength       =   10
      TabIndex        =   12
      Top             =   1155
      Width           =   2115
   End
   Begin MSDataListLib.DataCombo DCEmpleado 
      Bindings        =   "HorasES.frx":1728
      DataSource      =   "AdoEmpleado"
      Height          =   315
      Left            =   6615
      TabIndex        =   8
      Top             =   420
      Width           =   6525
      _ExtentX        =   11509
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSMask.MaskEdBox MBFechaI 
      Height          =   330
      Left            =   105
      TabIndex        =   1
      Top             =   420
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   582
      _Version        =   393216
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
   Begin MSAdodcLib.Adodc AdoEmpleado 
      Height          =   330
      Left            =   210
      Top             =   2940
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
      Caption         =   "Empleados"
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
   Begin MSAdodcLib.Adodc AdoHorasTrabajadas 
      Height          =   330
      Left            =   210
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
      Caption         =   "HorasTrabajadas"
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
      Left            =   210
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
   Begin VB.ComboBox CTV 
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
      Left            =   6720
      TabIndex        =   18
      Text            =   "%"
      Top             =   1155
      Width           =   750
   End
   Begin MSAdodcLib.Adodc AdoNovedades 
      Height          =   330
      Left            =   210
      Top             =   3885
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
      Caption         =   "Novedades"
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
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " ORDEN"
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
      Left            =   8925
      TabIndex        =   20
      Top             =   840
      Width           =   1275
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " DIAS:"
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
      Left            =   3885
      TabIndex        =   13
      Top             =   840
      Width           =   1065
   End
   Begin VB.Label Label15 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Valor por hora"
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
      Left            =   6720
      TabIndex        =   17
      Top             =   840
      Width           =   2115
   End
   Begin VB.Label Label9 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " HORAS EXTRAS"
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
      Left            =   5040
      TabIndex        =   15
      Top             =   840
      Width           =   1590
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " VALOR HORA:"
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
      Top             =   840
      Width           =   1485
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &Fecha:"
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
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " HORAS TRABAJADAS:"
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
      TabIndex        =   11
      Top             =   840
      Width           =   2115
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " BENEFICIARIO"
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
      TabIndex        =   7
      Top             =   105
      Width           =   6525
   End
End
Attribute VB_Name = "HorasEntSal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim NombreEmpleado As String

Private Sub Command1_Click()
Dim DiasTrabajados As Integer
Dim UltimodDiaMes As Integer
Dim Fecha_Empleado As String
Dim Fecha_Salida As String
Dim Fecha_IESS As String

    RatonReloj
    TextoValido TxtHorasTrab, True
    TextoValido TxtValorHora, True, True, 4
    TextoValido TxtOrden, , True      ' Orden
    FechaValida MBFechaI
    
    If OpcMensual.value Then MBFechaI = UltimoDiaMes(MBFechaI)
    MBFechaI.Refresh
    Datos_IESS MBFechaI
    FechaInicial = PrimerDiaMes(MBFechaI)
    FechaFinal = UltimoDiaMes(MBFechaI)
    FechaIni = BuscarFecha(FechaInicial)
    FechaFin = BuscarFecha(FechaFinal)
    UltimodDiaMes = Day(FechaFinal)
    Fecha_IESS = FechaFinal
    If Val(TxtDias) = "0" Then TxtDias = CStr(Day(FechaFinal))
    
    sSQL = "DELETE * " _
         & "FROM Trans_Rol_Horas " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# "
    Ejecutar_SQL_SP sSQL
    
    sSQL = "INSERT INTO Trans_Rol_Horas (Periodo, Item, T, Dias, Codigo, Fecha, Horas, Horas_Exts, Porc_Hr_Ext, Valor_Hora, Ing_Liquido, Ing_Horas_Ext, Orden, X) "
    
    If OpcSemana.value Then sSQL = sSQL & "SELECT Periodo, Item, T, 30, Codigo, #" & FechaFin & "#, Horas_Sem, 0 ,0, Valor_Hora, ROUND(Salario/4,2,0), 0, '" & TxtOrden.Text & "' ,'.' "
    If OpcQuincena.value Then sSQL = sSQL & "SELECT Periodo, Item, T, 30, Codigo, #" & FechaFin & "#, Horas_Sem*2, 0 ,0, Valor_Hora, ROUND(Salario/2,2,0), 0, '" & TxtOrden.Text & "' ,'.' "
    If OpcMensual.value Then sSQL = sSQL & "SELECT Periodo, Item, T, 30, Codigo, #" & FechaFin & "#, Horas_Sem*4, 0 ,0, Valor_Hora, Salario, 0, '" & TxtOrden.Text & "' ,'.' "
       
    sSQL = sSQL & "FROM Catalogo_Rol_Pagos " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND Fecha <= #" & FechaFin & "# " _
         & "AND Salario > 0 " _
         & "AND T = 'N' " _
         & "ORDER BY Ejecutivo "
    Ejecutar_SQL_SP sSQL
       
    sSQL = "INSERT INTO Trans_Rol_Horas (Periodo, Item, T, Dias, Dias_Enfermedad, Codigo, Fecha, Horas, Horas_Exts, Porc_Hr_Ext, Valor_Hora, Ing_Liquido, Ing_Horas_Ext, Orden, X) "
       
    If OpcSemana.value Then sSQL = sSQL & "SELECT Periodo, Item, T, 30, 0, Codigo, #" & FechaFin & "#, Horas_Sem, 0 ,0, Valor_Hora, ROUND(Salario/4,2,0), 0, '" & TxtOrden.Text & "' ,'.' "
    If OpcQuincena.value Then sSQL = sSQL & "SELECT Periodo, Item, T, 30, 0, Codigo, #" & FechaFin & "#, Horas_Sem*2, 0 ,0, Valor_Hora, ROUND(Salario/2,2,0), 0, '" & TxtOrden.Text & "' ,'.' "
    If OpcMensual.value Then sSQL = sSQL & "SELECT Periodo, Item, T, 30, 0, Codigo, #" & FechaFin & "#, Horas_Sem*4, 0 ,0, Valor_Hora, Salario, 0, '" & TxtOrden.Text & "' ,'.' "
       
    sSQL = sSQL & "FROM Catalogo_Rol_Pagos " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND FechaC BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
         & "AND Salario > 0 " _
         & "AND T = 'R' " _
         & "ORDER BY Ejecutivo "
    Ejecutar_SQL_SP sSQL
    
    sSQL = "UPDATE Trans_Rol_Horas " _
         & "SET Dias = DATEDIFF(DAY,CRP.Fecha,TRH.Fecha)+1 " _
         & "FROM Trans_Rol_Horas As TRH, Catalogo_Rol_Pagos As CRP " _
         & "WHERE TRH.Item = '" & NumEmpresa & "' " _
         & "AND TRH.Periodo = '" & Periodo_Contable & "' " _
         & "AND TRH.Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
         & "AND DATEDIFF(DAY,CRP.Fecha,TRH.Fecha) < 30 " _
         & "AND TRH.Item = CRP.Item " _
         & "AND TRH.Periodo = CRP.Periodo " _
         & "AND TRH.Codigo = CRP.Codigo "
    Ejecutar_SQL_SP sSQL
    
    sSQL = "UPDATE Trans_Rol_Horas " _
         & "SET Dias = DATEDIFF(DAY,'" & FechaIni & "',CRP.FechaC) + 1 " _
         & "FROM Trans_Rol_Horas As TRH, Catalogo_Rol_Pagos As CRP " _
         & "WHERE TRH.Item = '" & NumEmpresa & "' " _
         & "AND TRH.Periodo = '" & Periodo_Contable & "' " _
         & "AND TRH.Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
         & "AND CRP.T = 'R' " _
         & "AND DATEDIFF(DAY,'" & FechaIni & "',CRP.FechaC) < 30 " _
         & "AND TRH.Item = CRP.Item " _
         & "AND TRH.Periodo = CRP.Periodo " _
         & "AND TRH.Codigo = CRP.Codigo "
    Ejecutar_SQL_SP sSQL
    
    sSQL = "UPDATE Trans_Rol_Horas " _
         & "SET Dias = Dias - 1 " _
         & "FROM Trans_Rol_Horas As TRH, Catalogo_Rol_Pagos As CRP " _
         & "WHERE TRH.Item = '" & NumEmpresa & "' " _
         & "AND TRH.Periodo = '" & Periodo_Contable & "' " _
         & "AND TRH.Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
         & "AND DAY(CRP.Fecha) > 1 " _
         & "AND CRP.Fecha >= #" & FechaIni & "# " _
         & "AND " & UltimodDiaMes & " > 30 " _
         & "AND TRH.Item = CRP.Item " _
         & "AND TRH.Periodo = CRP.Periodo " _
         & "AND TRH.Codigo = CRP.Codigo "
    Ejecutar_SQL_SP sSQL
    
    sSQL = "UPDATE Trans_Rol_Horas " _
         & "SET Ing_Liquido = ROUND((Ing_Liquido/30)*Dias,2,0), Horas = ROUND((Horas_Sem*4/30)*Dias,0,0) " _
         & "FROM Trans_Rol_Horas As TRH, Catalogo_Rol_Pagos As CRP " _
         & "WHERE TRH.Item = '" & NumEmpresa & "' " _
         & "AND TRH.Periodo = '" & Periodo_Contable & "' " _
         & "AND TRH.Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
         & "AND Dias < 30 " _
         & "AND TRH.Item = CRP.Item " _
         & "AND TRH.Periodo = CRP.Periodo " _
         & "AND TRH.Codigo = CRP.Codigo "
    Ejecutar_SQL_SP sSQL
    
''  With AdoEmpleado.Recordset
''   If .RecordCount > 0 Then
''       Do While Not .EOF
''          NombreCliente = .fields("Cliente")
''          CodigoCliente = .fields("Codigo")
''          DCEmpleado.Text = .fields("Cliente")
''          Evaluar = .fields("Horas_Ext")
''          Grupo_No = .fields("Grupo")
''          TxtValorHora = .fields("Valor_Hora")
''          Fecha_Empleado = .fields("Fecha")
''          Fecha_Salida = .fields("FechaC")
''
''          If .fields("T") = "R" Then
''              If (CFechaLong(.fields("FechaC")) - CFechaLong(.fields("Fecha"))) > 30 Then Fecha_Empleado = PrimerDiaMes(MBFechaI) Else Fecha_Empleado = .fields("Fecha")
''              DiasTrabajados = CFechaLong(Fecha_Salida) - CFechaLong(Fecha_Empleado) + 1
''          Else
''              DiasTrabajados = CFechaLong(FechaFinal) - CFechaLong(Fecha_Empleado) + 1
''              If Day(UltimoDiaMes(FechaFinal)) = 28 Then DiasTrabajados = DiasTrabajados + 2
''              If Day(UltimoDiaMes(FechaFinal)) = 29 Then DiasTrabajados = DiasTrabajados + 1
''              If Day(UltimoDiaMes(FechaFinal)) = 31 Then DiasTrabajados = DiasTrabajados - 1
''          End If
''
''         'If DiasTrabajados <= UltimodDiaMes Then DiasTrabajados = DiasTrabajados + 1
''
''          'MsgBox NombreCliente & vbCrLf & vbCrLf & .fields("Salario") & vbCrLf & ValorTotal & vbCrLf & "Dias: " & DiasTrabajados
''
''          If OpcSemana.value Then
''             MiTiempo1 = .fields("Horas_Sem")
''             ValorTotal = .fields("Salario") / 4
''          End If
''          If OpcQuincena.value Then
''             MiTiempo1 = .fields("Horas_Sem") * 2
''             ValorTotal = .fields("Salario") / 2
''          End If
''          If OpcMensual.value Then
''             MiTiempo1 = .fields("Horas_Sem") * 4
''             ValorTotal = .fields("Salario")
''
''             If DiasTrabajados < Day(FechaFinal) Then
''                ValorTotal = Redondear((ValorTotal / 30) * DiasTrabajados, 2)
''                Dias_Laborados = DiasTrabajados
''             Else
''                Dias_Laborados = 30 'Day(FechaFinal)
''             End If
''          End If
''
''          MiTiempo = Time
''         'If CodigoCliente = "0020000046" Then MsgBox ".."
''          If ValorTotal > 0 Then
''             SetAdoAddNew "Trans_Rol_Horas"
''             SetAdoFields "T", Val(adFalse)
''             SetAdoFields "Dias", Dias_Laborados 'UltimodDiaMes
''             SetAdoFields "Codigo", CodigoCliente
''             SetAdoFields "Fecha", MBFechaI
''             SetAdoFields "Horas", MiTiempo1
''             SetAdoFields "Horas_Exts", 0
''             SetAdoFields "Porc_Hr_Ext", 0
''             SetAdoFields "Ing_Horas_Ext", 0
''             SetAdoFields "Valor_Hora", .fields("Valor_Hora")
''             If Dias_Laborados = 30 Then
''                SetAdoFields "Ing_Liquido", Redondear(.fields("Salario"), 2)
''             Else
''                SetAdoFields "Ing_Liquido", Redondear(ValorTotal, 2)
''             End If
''             SetAdoFields "Orden", TxtOrden
''             SetAdoFields "CodigoU", CodigoUsuario
''             SetAdoFields "Item", NumEmpresa
''             SetAdoUpdate
''          End If
''         .MoveNext
''       Loop
''   End If
''  End With
  RatonNormal
  MsgBox "PROCESO TERMINADO"
End Sub

Private Sub Command2_Click()
  Unload Me
End Sub

Private Sub Command3_Click()
  FechaValida MBFechaI
  sSQL = "DELETE * " _
       & "FROM Trans_Rol_Horas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND Fecha = #" & BuscarFecha(MBFechaI) & "# "
  Ejecutar_SQL_SP sSQL
  MsgBox "Proceso Terminado, Vuelva a generar las horas laboradas"
End Sub

Private Sub DCEmpleado_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
  If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub DCEmpleado_LostFocus()
  FechaValida MBFechaI
  CodigoCliente = Ninguno
  NombreCliente = DCEmpleado.Text
  Grupo_No = DCEmpleado.Text
  TxtValorHora.Text = "0.00"
  With AdoEmpleado.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Cliente = '" & NombreCliente & "' ")
       If Not .EOF Then
          NombreCliente = .fields("Cliente")
          CodigoCliente = .fields("Codigo")
          DCEmpleado.Text = .fields("Cliente")
          Evaluar = .fields("Horas_Ext")
          Grupo_No = .fields("Grupo")
          TxtValorHora.Text = .fields("Valor_Hora")
          MiTiempo1 = .fields("Horas_Sem") * 4
          TotalIngreso = .fields("Salario")
       Else
          .MoveFirst
          .Find ("Grupo = '" & Grupo_No & "' ")
           If Not .EOF Then
              NombreCliente = .fields("Cliente")
              CodigoCliente = .fields("Codigo")
              DCEmpleado.Text = .fields("Cliente")
              Evaluar = .fields("Horas_Ext")
              Grupo_No = .fields("Grupo")
              TxtValorHora.Text = .fields("Valor_Hora")
              MiTiempo1 = .fields("Horas_Sem") * 4
              TotalIngreso = .fields("Salario")
           Else
              MsgBox "Codigo No asignado"
              MBFechaI.SetFocus
              Exit Sub
           End If
       End If
   End If
  End With
  
  sSQL = "SELECT TOP 1 Codigo, Valor_Hora " _
       & "FROM Trans_Rol_Horas " _
       & "WHERE Fecha <= #" & BuscarFecha(MBFechaI.Text) & "# " _
       & "AND Codigo = '" & CodigoCliente & "' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "ORDER BY Fecha DESC,Orden DESC "
  Select_Adodc AdoAux, sSQL
  If AdoAux.Recordset.RecordCount > 0 Then
     'TxtValorHora.Text = AdoAux.Recordset.Fields("Valor_Hora")
  End If
  
  ListarHorasTrabajadas CodigoCliente
  
  sSQL = "SELECT Fecha,Hora,Proceso,Tarea as Novedades,Codigo " _
       & "FROM Trans_Entrada_Salida " _
       & "WHERE Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
       & "AND Codigo = '" & CodigoCliente & "' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND ES = 'R' " _
       & "ORDER BY Fecha DESC "
  Select_Adodc_Grid DGNovedades, AdoNovedades, sSQL, SQLDec
  NombreEmpleado = DCEmpleado.Text
  TxtValorHora.SetFocus
End Sub
      
Private Sub DGHorasTrabajadas_KeyDown(KeyCode As Integer, Shift As Integer)
    Keys_Especiales Shift
    
    If AdoHorasTrabajadas.Recordset.RecordCount > 0 Then
     ' CodigoCliente = DGHorasTrabajadas.Columns(0)
       Mifecha = DGHorasTrabajadas.Columns(3)
       Double1 = Round(Val(DGHorasTrabajadas.Columns(4)), 4)
       Codigo1 = DGHorasTrabajadas.Columns(7)
    End If
    
'    MsgBox Mifecha
    If KeyCode = vbKeyDelete Then
       Mensajes = "Esta seguro de Eliminar el registro del " & Mifecha & ", Codigo: " & CodigoCliente & vbCrLf
       Titulo = "Pregunta de Eliminacion"
       If BoxMensaje = vbYes Then
          sSQL = "DELETE * " _
               & "FROM Trans_Rol_Horas " _
               & "WHERE Item = '" & NumEmpresa & "' " _
               & "AND Periodo = '" & Periodo_Contable & "' " _
               & "AND Fecha = #" & BuscarFecha(Mifecha) & "# " _
               & "AND Codigo = '" & CodigoCliente & "' " _
               & "AND Horas = " & Double1 & " "
          Ejecutar_SQL_SP sSQL
          ListarHorasTrabajadas CodigoCliente
       End If
    End If
    If CtrlDown And KeyCode = vbKeyE Then
       FechaIni = BuscarFecha(MBFechaI)
       
       NoDias = InputBox("El Sr(a). " & NombreCliente & ", Cuantos dias de Permiso de enfermedad tiene?", "PERMISO DE ENFERMEDAD", "0")
       If NoDias < 0 Then NoDias = 0
       If NoDias > 30 Then NoDias = 30
       If NoDias > 4 Then
          sSQL = "UPDATE Trans_Rol_Horas " _
               & "SET Dias_Enfermedad = " & NoDias & " " _
               & "FROM Trans_Rol_Horas " _
               & "WHERE Item = '" & NumEmpresa & "' " _
               & "AND Periodo = '" & Periodo_Contable & "' " _
               & "AND Codigo = '" & CodigoCliente & "' " _
               & "AND Fecha = #" & FechaFin & "# "
          Ejecutar_SQL_SP sSQL
       Else
          MsgBox "Advertencia: No se pudo hacer el proceso, ingrese mas de 3 dias de enfermedad"
       End If
    End If
End Sub

Private Sub DGNovedades_KeyDown(KeyCode As Integer, Shift As Integer)
Dim Proceso As String
  If KeyCode = vbKeyDelete Then
     Mifecha = DGNovedades.Columns(0)
     MiHora = DGNovedades.Columns(1)
     CodigoCliente = DGNovedades.Columns(4)
     Mensajes = "Esta seguro de Eliminar el registro del " & vbCrLf _
              & Mifecha & ", H: " & MiHora & ", Codigo: " & CodigoCliente & vbCrLf
     Titulo = "Pregunta de Eliminacion"
     If BoxMensaje = vbYes Then
        sSQL = "DELETE * " _
             & "FROM Trans_Entrada_Salida " _
             & "WHERE Item = '" & NumEmpresa & "' " _
             & "AND Periodo = '" & Periodo_Contable & "' " _
             & "AND Fecha = #" & BuscarFecha(Mifecha) & "# " _
             & "AND Codigo = '" & CodigoCliente & "' " _
             & "AND Hora = '" & MiHora & "' "
        Ejecutar_SQL_SP sSQL
     End If
  End If
  If KeyCode = vbKeyInsert Then
     Proceso = UCase(InputBox("INGRESE LA NOVEDAD DEL MES: ", "NOVEDADES", " "))
     If Proceso <> "" Then
        SetAdoAddNew "Trans_Entrada_Salida"
        SetAdoFields "ES", "R"
        SetAdoFields "Codigo", CodigoCliente
        SetAdoFields "Hora", Format(Time, FormatoTimes)
        SetAdoFields "Fecha", MBFechaI
        SetAdoFields "Proceso", "NOVEDADES"
        SetAdoFields "Tarea", TrimStrg(MidStrg(Proceso, 1, 50))
        SetAdoFields "CodigoU", CodigoUsuario
        SetAdoFields "Periodo", Periodo_Contable
        SetAdoFields "Item", NumEmpresa
        SetAdoUpdate
     End If
  End If
  sSQL = "SELECT Fecha,Hora,Proceso,Tarea as Novedades,Codigo " _
       & "FROM Trans_Entrada_Salida " _
       & "WHERE Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
       & "AND Codigo = '" & CodigoCliente & "' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND ES = 'R' " _
       & "ORDER BY Fecha DESC "
  Select_Adodc_Grid DGNovedades, AdoNovedades, sSQL, SQLDec
End Sub

Private Sub Form_Activate()
  CTV.Clear
  CTV.AddItem "V"
  CTV.AddItem "%"
  CTV.Text = CTV.List(0)
  NombreEmpleado = Ninguno
  RatonNormal
End Sub

Private Sub Form_Load()
  CentrarForm HorasEntSal
  ConectarAdodc AdoAux
  ConectarAdodc AdoEmpleado
  ConectarAdodc AdoNovedades
  ConectarAdodc AdoHorasTrabajadas
  HorasEntSal.Caption = "REGISTRO DE HORAS TRABAJADAS"
End Sub

Private Sub MBFechaI_GotFocus()
  MarcarTexto MBFechaI
End Sub

Private Sub MBFechaI_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBFechaI_LostFocus()
Dim PrimerDia As String
Dim UltimoDia As String
  
  FechaValida MBFechaI
  PrimerDia = BuscarFecha(PrimerDiaMes(MBFechaI))
  UltimoDia = BuscarFecha(UltimoDiaMes(MBFechaI))

  sSQL = "SELECT RP.T, C.Cliente, C.CI_RUC, C.Grupo, RP.Codigo, RP.Horas_Ext, RP.Valor_Hora, RP.Fecha, RP.FechaC, RP.Horas_Sem, RP.Salario, RP.Valor_Hora " _
       & "FROM Clientes As C,Catalogo_Rol_Pagos As RP " _
       & "WHERE RP.Item = '" & NumEmpresa & "' " _
       & "AND RP.Periodo = '" & Periodo_Contable & "' " _
       & "AND RP.T = '" & Normal & "' " _
       & "AND RP.Fecha <= #" & UltimoDia & "# " _
       & "AND RP.Salario > 0 " _
       & "AND C.Codigo = RP.Codigo " _
       & "UNION " _
       & "SELECT RP.T, C.Cliente, C.CI_RUC, C.Grupo, RP.Codigo, RP.Horas_Ext, RP.Valor_Hora, RP.Fecha, RP.FechaC, RP.Horas_Sem, RP.Salario, RP.Valor_Hora " _
       & "FROM Clientes As C,Catalogo_Rol_Pagos As RP " _
       & "WHERE RP.Item = '" & NumEmpresa & "' " _
       & "AND RP.Periodo = '" & Periodo_Contable & "' " _
       & "AND RP.T = 'R' " _
       & "AND RP.FechaC BETWEEN #" & PrimerDia & "# and #" & UltimoDia & "# " _
       & "AND RP.Salario > 0 " _
       & "AND C.Codigo = RP.Codigo " _
       & "ORDER BY C.Cliente "
  SelectDB_Combo DCEmpleado, AdoEmpleado, sSQL, "Cliente"
  If NombreEmpleado <> Ninguno Then DCEmpleado.Text = NombreEmpleado
End Sub

Private Sub Opc120_Click()
  ListarHorasTrabajadas CodigoCliente
End Sub

Private Sub Opc31_Click()
   ListarHorasTrabajadas CodigoCliente
End Sub

Private Sub Opc60_Click()
  ListarHorasTrabajadas CodigoCliente
End Sub

Private Sub Opc90_Click()
  ListarHorasTrabajadas CodigoCliente
End Sub

Private Sub OpcTodos_Click()
  ListarHorasTrabajadas CodigoCliente
End Sub

Private Sub TxtDias_GotFocus()
  MarcarTexto TxtDias
End Sub

Private Sub TxtDias_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtDias_LostFocus()
  TextoValido TxtDias, True
  If Val(TxtDias) = "0" Then TxtDias = CStr(Day(MBFechaI))
End Sub

Private Sub TxtHorasExt_GotFocus()
  MarcarTexto TxtHorasExt
End Sub

Private Sub TxtHorasExt_LostFocus()
  TextoValido TxtHorasExt
  If Val(TxtHorasExt.Text) <= 0 Then TxtPorcHExt.Text = "0"
End Sub

Private Sub TxtHorasTrab_GotFocus()
  MarcarTexto TxtHorasTrab
End Sub

Private Sub TxtHorasTrab_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtHorasTrab_LostFocus()
Dim HorasPorDia As Double
Dim HorasTrab As Double
Dim HorasExt As Double
  TextoValido TxtHorasTrab, True
  'MsgBox MiTiempo1
  HorasExt = 0
  HorasTrab = Val(TxtHorasTrab)
  If OpcDia.value Then
     HorasPorDia = MiTiempo1 / 31
  ElseIf OpcSemana.value Then
     HorasPorDia = MiTiempo1 / 4
  Else
     HorasPorDia = MiTiempo1
  End If
  If (Evaluar) And (HorasTrab > HorasPorDia) Then
     HorasExt = HorasTrab - HorasPorDia
     HorasTrab = HorasPorDia
  End If
  NoDias = MiTiempo1 / 20
  If NoDias = 0 Then NoDias = 1
  NoDias = Round(HorasTrab / NoDias)
  TxtHorasExt.Text = Format(HorasExt, "#,##0.00")
  TxtHorasTrab.Text = Format(HorasTrab, "#,##0.00")
End Sub

Private Sub TxtOrden_LostFocus()
  TextoValido TxtHorasTrab, True
  TextoValido TxtValorHora, True, True, 5
  TextoValido TxtOrden, , True      ' Orden
  FechaValida MBFechaI
  Total = Val(TxtHorasTrab.Text) + Val(TxtHorasExt.Text)
  Total_ME = 0
  If CodigoCliente <> Ninguno And Total > 0 Then
     Debe = Val(CCur(TxtHorasTrab.Text))         ' Horas Trabajadas
     VUnitTemp = Val(CSng(TxtValorHora.Text))    ' Valor por Hora
     Debe_ME = Val(CCur(TxtHorasExt.Text))       ' Horas Extras
     If CTV.Text = "%" Then
        Cuota = Val(CSng(TxtPorcHExt.Text))
        Cuota = 1 + (Cuota / 100)                ' Porcentaje de Horas Extras
        Total_ME = Debe_ME * (VUnitTemp * Cuota) ' Calculo horas Extras
     Else
        Cuota = Val(CSng(TxtPorcHExt.Text))
        Total_ME = Debe_ME * Cuota               ' Calculo horas Extras
     End If
     Total = Debe * VUnitTemp                    ' Calculo hornal Normal
     ValorTotal = Total + Total_ME               ' Total de Horas Trabajadas
    'MsgBox ValorTotal
     MiTiempo = Time
     Cadena = "EL SALARIO ASIGNADO" & vbCrLf & vbCrLf _
            & "AL MES ES DE: " & Moneda & " " & Format(TotalIngreso, "#,##0.00") & vbCrLf & vbCrLf _
            & "ESTA CORRECTA LA ASIGNACION:"
     ValorTotal = Val(InputBox(Cadena, "REDONDEO DE ASIGNACION", Format(ValorTotal, "##0.00")))
     If ValorTotal > 0 Then
        SetAdoAddNew "Trans_Rol_Horas"
        SetAdoFields "T", Val(adFalse)
        SetAdoFields "Dias", Val(TxtDias)
        SetAdoFields "Codigo", CodigoCliente
        SetAdoFields "Fecha", MBFechaI
        SetAdoFields "Horas", Debe
        SetAdoFields "Horas_Exts", Debe_ME
        If CTV.Text = "%" Then SetAdoFields "Porc_Hr_Ext", Cuota
        SetAdoFields "Valor_Hora", Haber
        SetAdoFields "Ing_Liquido", ValorTotal  'Total/ValorTotal
        SetAdoFields "Certificado", ValorTotal * Bonif
        SetAdoFields "Aporte_Adm", ValorTotal * Entrada
        SetAdoFields "Ing_Horas_Ext", Total_ME
        SetAdoFields "Orden", TxtOrden
        SetAdoFields "CodigoU", CodigoUsuario
        SetAdoFields "Item", NumEmpresa
        SetAdoUpdate
     End If
  Else
     MsgBox "Datos incompletos"
  End If
  ListarHorasTrabajadas CodigoCliente
  DCEmpleado.SetFocus
End Sub

Private Sub TxtPorcHExt_GotFocus()
  MarcarTexto TxtPorcHExt
End Sub

Private Sub TxtPorcHExt_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Public Sub ListarHorasTrabajadas(TipoCodigo As String)
  FechaValida MBFechaI
  If OpcTodos.value Then
     FechaTexto = "01/01/" & Format(Year(MBFechaI.Text), "0000")
  Else
     If Opc31.value Then NoMeses = Month(MBFechaI.Text)
     If Opc60.value Then NoMeses = Month(MBFechaI.Text) - 1
     If Opc90.value Then NoMeses = Month(MBFechaI.Text) - 2
     If Opc120.value Then NoMeses = Month(MBFechaI.Text) - 3
     If NoMeses <= 0 Then NoMeses = 1
     FechaTexto = "01/" & Format(NoMeses, "00") & "/" & Format(Year(MBFechaI.Text), "0000")
  End If
  FechaIni = BuscarFecha(FechaTexto)
  FechaFin = BuscarFecha(MBFechaI.Text)
  Total = 0: Saldo = 0
  sSQL = "SELECT Codigo, Dias, Dias_Enfermedad, Fecha, Horas, Horas_Exts, Porc_Hr_Ext, Valor_Hora, Ing_Liquido, Ing_Horas_Ext, Orden " _
       & "FROM Trans_Rol_Horas " _
       & "WHERE Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
       & "AND Codigo = '" & TipoCodigo & "' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "ORDER BY Fecha DESC,Orden DESC "
  SQLDec = "Valor_Hora 5|."
  Select_Adodc_Grid DGHorasTrabajadas, AdoHorasTrabajadas, sSQL, SQLDec
  With AdoHorasTrabajadas.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          Total = Total + .fields("Horas")
          Saldo = Saldo + .fields("Ing_Liquido")
         .MoveNext
       Loop
   End If
  End With
  LabelFacturado.Caption = Format(Total, "#,##0.00")
  LabelAbonado.Caption = Format(Saldo, "#,##0.00")
  MBFechaI.SetFocus
End Sub

Private Sub TxtValorHora_GotFocus()
  MarcarTexto TxtValorHora
  DCEmpleado.Text = NombreCliente
End Sub

Private Sub TxtValorHora_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtValorHora_LostFocus()
  TextoValido TxtValorHora, True, True, 5
End Sub
