VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Begin VB.Form ProcesarConciliacion 
   Caption         =   "Transito"
   ClientHeight    =   7365
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11460
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7365
   ScaleWidth      =   11460
   WindowState     =   2  'Maximized
   Begin MSDataListLib.DataCombo DCCtas 
      Bindings        =   "ProcConc.frx":0000
      DataSource      =   "AdoBanco"
      Height          =   345
      Left            =   3150
      TabIndex        =   1
      Top             =   105
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   609
      _Version        =   393216
      Text            =   "Cuentas"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DGTrans 
      Bindings        =   "ProcConc.frx":0017
      Height          =   2325
      Left            =   105
      TabIndex        =   26
      Top             =   3780
      Width           =   13350
      _ExtentX        =   23548
      _ExtentY        =   4101
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
   Begin MSDataGridLib.DataGrid DGNotaDC 
      Bindings        =   "ProcConc.frx":002E
      Height          =   1695
      Left            =   105
      TabIndex        =   25
      Top             =   1155
      Width           =   12300
      _ExtentX        =   21696
      _ExtentY        =   2990
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
   Begin MSAdodcLib.Adodc AdoBanco 
      Height          =   330
      Left            =   315
      Top             =   1260
      Visible         =   0   'False
      Width           =   2010
      _ExtentX        =   3545
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
      Caption         =   "Banco"
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
   Begin MSMask.MaskEdBox MBoxFechaI 
      Height          =   330
      Left            =   1890
      TabIndex        =   3
      Top             =   525
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   582
      _Version        =   393216
      AllowPrompt     =   -1  'True
      AutoTab         =   -1  'True
      MaxLength       =   10
      OLEDragMode     =   1
      OLEDropMode     =   2
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
      OLEDragMode     =   1
      OLEDropMode     =   2
   End
   Begin VB.TextBox TextValor 
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
      Left            =   10815
      MaxLength       =   12
      MultiLine       =   -1  'True
      TabIndex        =   7
      Text            =   "ProcConc.frx":0046
      ToolTipText     =   "Monto del Asiento"
      Top             =   525
      Width           =   1590
   End
   Begin MSMask.MaskEdBox MBoxFechaF 
      Height          =   330
      Left            =   4200
      TabIndex        =   5
      Top             =   525
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   582
      _Version        =   393216
      AllowPrompt     =   -1  'True
      AutoTab         =   -1  'True
      MaxLength       =   10
      OLEDragMode     =   1
      OLEDropMode     =   2
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
      OLEDragMode     =   1
      OLEDropMode     =   2
   End
   Begin VB.CommandButton Command1 
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
      Left            =   12495
      Picture         =   "ProcConc.frx":004D
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1995
      Width           =   960
   End
   Begin VB.CommandButton Command2 
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
      Left            =   12495
      Picture         =   "ProcConc.frx":0917
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1050
      Width           =   960
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Procesar"
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
      Left            =   12495
      Picture         =   "ProcConc.frx":11E1
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   105
      Width           =   960
   End
   Begin MSAdodcLib.Adodc AdoNotaDC 
      Height          =   330
      Left            =   315
      Top             =   1575
      Visible         =   0   'False
      Width           =   2010
      _ExtentX        =   3545
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
      Caption         =   "NotaDC"
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
   Begin MSAdodcLib.Adodc AdoTrans 
      Height          =   330
      Left            =   315
      Top             =   1890
      Visible         =   0   'False
      Width           =   2010
      _ExtentX        =   3545
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
      Caption         =   "Trans"
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
   Begin MSAdodcLib.Adodc AdoCtas 
      Height          =   330
      Left            =   315
      Top             =   2205
      Visible         =   0   'False
      Width           =   2010
      _ExtentX        =   3545
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
      Caption         =   "Ctas"
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
   Begin MSAdodcLib.Adodc AdoND 
      Height          =   330
      Left            =   2310
      Top             =   1260
      Visible         =   0   'False
      Width           =   2010
      _ExtentX        =   3545
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
      Caption         =   "ND"
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
   Begin MSAdodcLib.Adodc AdoNC 
      Height          =   330
      Left            =   2310
      Top             =   1575
      Visible         =   0   'False
      Width           =   2010
      _ExtentX        =   3545
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
      Caption         =   "NC"
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
   Begin MSDataGridLib.DataGrid DGDebHabTrans 
      Bindings        =   "ProcConc.frx":1623
      Height          =   2220
      Left            =   105
      TabIndex        =   27
      Top             =   6510
      Width           =   13350
      _ExtentX        =   23548
      _ExtentY        =   3916
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
   Begin VB.Label Label11 
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
      ForeColor       =   &H00800000&
      Height          =   330
      Left            =   11655
      TabIndex        =   18
      Top             =   9135
      Width           =   1800
   End
   Begin VB.Label Label20 
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
      ForeColor       =   &H00800000&
      Height          =   330
      Left            =   11655
      TabIndex        =   29
      Top             =   8715
      Width           =   1800
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DIFERENCIA TOTAL ENTRE DEBITOS Y CREDITOS EN TRANSITO:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   330
      Left            =   105
      TabIndex        =   28
      Top             =   8715
      Width           =   11565
   End
   Begin VB.Label Label7 
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
      ForeColor       =   &H00800000&
      Height          =   330
      Left            =   10815
      TabIndex        =   15
      Top             =   840
      Width           =   1590
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " SALDO SEGUN LIBRO BANCO:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   330
      Left            =   105
      TabIndex        =   16
      Top             =   840
      Width           =   10725
   End
   Begin VB.Label Label15 
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
      ForeColor       =   &H00800000&
      Height          =   330
      Left            =   11655
      TabIndex        =   22
      Top             =   3465
      Width           =   1800
   End
   Begin VB.Label Label14 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " SALDO SEGUN ESTADO DE CUENTA:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   330
      Left            =   105
      TabIndex        =   21
      Top             =   3465
      Width           =   11565
   End
   Begin VB.Label Label13 
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
      ForeColor       =   &H00800000&
      Height          =   330
      Left            =   11655
      TabIndex        =   20
      Top             =   9450
      Width           =   1800
   End
   Begin VB.Label Label12 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " DIFERENCIA ENTRE SALDOS REALES DE LIBRO BANCOS Y ESTADO DE CUENTA:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   330
      Left            =   105
      TabIndex        =   19
      Top             =   9135
      Width           =   11565
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " HASTA EL"
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
      Left            =   3150
      TabIndex        =   4
      Top             =   525
      Width           =   1065
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " SALDO DEL ESTADO DE CUENTA"
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
      TabIndex        =   6
      Top             =   525
      Width           =   5370
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " FECHA DESDE EL"
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
      Top             =   525
      Width           =   1800
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " BANCO A CONCILIAR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   330
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   3060
   End
   Begin VB.Label Label19 
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
      ForeColor       =   &H00800000&
      Height          =   330
      Left            =   11655
      TabIndex        =   11
      Top             =   3150
      Width           =   1800
   End
   Begin VB.Label Label5 
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
      ForeColor       =   &H00800000&
      Height          =   330
      Left            =   10815
      TabIndex        =   13
      Top             =   2835
      Width           =   1590
   End
   Begin VB.Label Label18 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " SALDO REAL CONCILIADO DE LIBRO BANCOS: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   330
      Left            =   105
      TabIndex        =   12
      Top             =   3150
      Width           =   11565
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DIFERENCIA TOTAL ENTRE NOTAS DE DEBITOS Y NOTAS DE CREDITOS:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   330
      Left            =   105
      TabIndex        =   14
      Top             =   2835
      Width           =   10725
   End
   Begin VB.Label Label16 
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
      ForeColor       =   &H00800000&
      Height          =   330
      Left            =   11655
      TabIndex        =   23
      Top             =   6090
      Width           =   1800
   End
   Begin VB.Label Label10 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " SALDO REAL CONCILIADO DEL ESTADO DE CUENTA:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   330
      Left            =   105
      TabIndex        =   17
      Top             =   9450
      Width           =   11565
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DIFERENCIA TOTAL ENTRE DEBITOS Y CREDITOS:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   330
      Left            =   105
      TabIndex        =   24
      Top             =   6090
      Width           =   11565
   End
End
Attribute VB_Name = "ProcesarConciliacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
  Unload ProcesarConciliacion
End Sub

Private Sub Command2_Click()
  Total_Saldos = Round(Total_Saldos, 2)
  If Total_Saldos = 0 Then
     With Heads
         .MsgTitulo = "CONCILIACION BANCARIA"
         .MsgObjetivo = "Banco: "
         .MsgConcepto = "Fechas: "
         .TextoObjetivo = DCCtas.Text
         .TextoConcepto = "Desde:  " & MBoxFechaI.Text & "   hasta:   " & MBoxFechaF.Text
     End With
     Imprimir_Conciliacion AdoNotaDC, AdoTrans, AdoND, Moneda_US
  Else
     Cadena = "Chequee los Ingresos, Egresos," & vbCrLf _
            & "Notas de Débitos y Notas de Créditos" & vbCrLf _
            & "para poder conciliar el Banco"
     MsgBox Cadena
  End If
End Sub

Private Sub Command3_Click()
  Total_Bancos = Val(CCur(TextValor))  'Saldo del estado de cuenta
  Saldo = 0: Saldo_ME = 0
  Debe = 0: Haber = 0
  Debitos = 0: Creditos = 0
  SumaDebe = 0: SumaHaber = 0
  FechaValida MBoxFechaI, False
  FechaValida MBoxFechaF, False
  Codigo1 = SinEspaciosIzq(DCCtas)
  FechaIni = BuscarFecha(MBoxFechaI)
  FechaFin = BuscarFecha(MBoxFechaF)
  sSQL = "SELECT Fecha,Concepto,TP,Documento,Debitos,Creditos " _
       & "FROM Trans_Conciliacion " _
       & "WHERE Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
       & "AND Cta = '" & Codigo1 & "' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "ORDER BY TP DESC,Fecha "
       
  sSQL = "SELECT T.Fecha,C.Concepto,T.TP,T.Cheq_Dep As Documento,T.Debe As Debitos,T.Haber As Creditos,T.Item " _
       & "FROM Transacciones As T,Comprobantes As C,Clientes As Cl " _
       & "WHERE T.Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
       & "AND T.Periodo = '" & Periodo_Contable & "' " _
       & "AND T.T <> 'A' " _
       & "AND T.TP IN ('ND','NC') " _
       & "AND T.Cta = '" & Codigo1 & "' "
  If ConSucursal = False Then sSQL = sSQL & "AND T.Item = '" & NumEmpresa & "' "
  sSQL = sSQL & "AND C.TP = T.TP " _
       & "AND C.Numero = T.Numero " _
       & "AND C.Item = T.Item " _
       & "AND C.Periodo = T.Periodo " _
       & "AND C.Codigo_B = Cl.Codigo " _
       & "ORDER BY T.Fecha "
  Select_Adodc_Grid DGNotaDC, AdoNotaDC, sSQL
  DGTrans.Visible = False
  sSQL = "SELECT Cta,Fecha,TP,Numero,Debe,Haber,ID,Saldo,Saldo_ME " _
       & "FROM  Transacciones " _
       & "WHERE Fecha <= #" & FechaFin & "# " _
       & "AND T <> '" & Anulado & "' " _
       & "AND Cta = '" & Codigo1 & "' "
  If ConSucursal = False Then sSQL = sSQL & "AND Item = '" & NumEmpresa & "' "
  sSQL = sSQL & "AND Periodo = '" & Periodo_Contable & "' " _
       & "ORDER BY Cta,Fecha,TP,Numero,Debe DESC,Haber,ID "
  Select_Adodc AdoTrans, sSQL
  RatonReloj
  With AdoTrans.Recordset
   If .RecordCount > 0 Then
      .MoveLast
       Saldo = .Fields("Saldo")
       Saldo_ME = .Fields("Saldo_ME")
   End If
  End With
  DGNotaDC.Visible = False
  With AdoNotaDC.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       Do While Not .EOF
          ProcesarConciliacion.Caption = "Procesando Conciliacion de N/D y N/C, fecha: " & .Fields("Fecha")
          SumaDebe = SumaDebe + .Fields("Debitos")
          SumaHaber = SumaHaber + .Fields("Creditos")
         .MoveNext
       Loop
   End If
  End With
  sSQL = "SELECT T.Fecha,Cliente,T.TP,T.Numero,Cheq_Dep,Parcial_ME,Debe,Haber,T.Item " _
       & "FROM Transacciones As T,Comprobantes As C,Clientes As Cl " _
       & "WHERE T.Fecha <= #" & FechaFin & "# " _
       & "AND T.Periodo = '" & Periodo_Contable & "' " _
       & "AND T.T <> 'A' " _
       & "AND T.C = " & Val(adFalse) & " " _
       & "AND T.Cta = '" & Codigo1 & "' "
  If ConSucursal = False Then sSQL = sSQL & "AND T.Item = '" & NumEmpresa & "' "
  sSQL = sSQL & "AND C.TP = T.TP " _
       & "AND C.Numero = T.Numero " _
       & "AND C.Item = T.Item " _
       & "AND C.Periodo = T.Periodo " _
       & "AND C.Codigo_B = Cl.Codigo " _
       & "ORDER BY T.Fecha "
  Select_Adodc_Grid DGTrans, AdoTrans, sSQL
  DGTrans.Visible = False
  With AdoTrans.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       Do While Not .EOF
          ProcesarConciliacion.Caption = "Procesando Conciliacion de Debitos y Creditos, fecha: " & .Fields("Fecha")
          If Moneda_US Then
             'Debe = Debe + .Fields("Debe_ME")
             'Haber = Haber + .Fields("Haber_ME")
          Else
             Debe = Debe + .Fields("Debe")
             Haber = Haber + .Fields("Haber")
          End If
         .MoveNext
       Loop
   End If
  End With
  
  sSQL = "SELECT Fecha, Concepto, TP, Documento, Debe, Haber, Item " _
       & "FROM Trans_Transito " _
       & "WHERE Periodo = '" & Periodo_Contable & "' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
       & "AND Cta = '" & Codigo1 & "' " _
       & "ORDER BY Fecha, TP, Numero "
  Select_Adodc_Grid DGDebHabTrans, AdoND, sSQL
  DGDebHabTrans.Visible = False
  With AdoND.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       Do While Not .EOF
          ProcesarConciliacion.Caption = "Procesando Conciliacion de Debitos y Creditos, fecha: " & .Fields("Fecha")
          Debitos = Debitos + .Fields("Debe")
          Creditos = Creditos + .Fields("Haber")
         .MoveNext
       Loop
   End If
  End With
  
 'MsgBox Total_Bancos
  ProcesarConciliacion.Caption = "PROCESAR CONCILIACION BANCARIA"
  Label7.Caption = Format(Saldo, "#,##0.00")
  Label5.Caption = Format(SumaDebe - SumaHaber, "#,##0.00")
  Label19.Caption = Format(Saldo + SumaDebe - SumaHaber, "#,##0.00")
  Label15.Caption = Format(Total_Bancos, "#,##0.00")
  Label16.Caption = Format(Debe - Haber, "#,##0.00")
  Label19.Caption = Format(Debitos - Credito, "#,##0.00")
  Label11.Caption = Format(Total_Bancos + Debe - Haber + Debitos - Creditos, "#,##0.00")
  '- SumaDebe + SumaHaber
  Total_Saldos = Saldo - Total_Bancos - Debe + Haber - Debitos + Creditos
  Label13.Caption = Format(Total_Saldos, "#,##0.00")
  DGNotaDC.Visible = True
  DGTrans.Visible = True
  DGDebHabTrans.Visible = True
  RatonNormal
End Sub

Private Sub DBCCtas_LostFocus()
  Codigo1 = SinEspaciosIzq(DBCCtas.Text)
  Codigo = Leer_Cta_Catalogo(Codigo1)
End Sub

Private Sub DGNotaDC_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  If CtrlDown And KeyCode = vbKeyF1 Then GenerarDataTexto ProcesarConciliacion, AdoNotaDC
End Sub

Private Sub DGTrans_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  If CtrlDown And KeyCode = vbKeyF1 Then GenerarDataTexto ProcesarConciliacion, AdoTrans
End Sub

Private Sub Form_Activate()
  sSQL = "SELECT Codigo & Space(20) & Cuenta As Nombre_Cta " _
       & "FROM Catalogo_Cuentas " _
       & "WHERE TC = 'BA' " _
       & "AND DG = 'D' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "ORDER BY Codigo "
  SelectDB_Combo DCCtas, AdoBanco, sSQL, "Nombre_Cta", False
  RatonNormal
End Sub

Private Sub Form_Load()
  CentrarForm ProcesarConciliacion
  ConectarAdodc AdoND
  ConectarAdodc AdoNC
  ConectarAdodc AdoCtas
  ConectarAdodc AdoBanco
  ConectarAdodc AdoTrans
  ConectarAdodc AdoNotaDC
End Sub

Private Sub MBoxFechaF_GotFocus()
    MarcarTexto MBoxFechaF
End Sub

Private Sub MBoxFechaF_LostFocus()
  FechaValida MBoxFechaF
End Sub

Private Sub MBoxFechaI_GotFocus()
  MarcarTexto MBoxFechaI
End Sub

Private Sub MBoxFechaI_LostFocus()
  FechaValida MBoxFechaI
  MBoxFechaF = UltimoDiaMes(MBoxFechaI)
End Sub

Private Sub TextValor_GotFocus()
  MarcarTexto TextValor
End Sub

Private Sub TextValor_LostFocus()
  TextoValido TextValor, True
End Sub
