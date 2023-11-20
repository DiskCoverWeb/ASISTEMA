VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Begin VB.Form FPolizas 
   Caption         =   "Catalogo de Rol de Pagos"
   ClientHeight    =   5520
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12210
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5520
   ScaleWidth      =   12210
   WindowState     =   2  'Maximized
   Begin VB.ListBox LstCuentas 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   645
      Left            =   7770
      TabIndex        =   7
      Top             =   945
      Width           =   2850
   End
   Begin VB.TextBox TxtInteres 
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
      Left            =   105
      MaxLength       =   12
      MultiLine       =   -1  'True
      TabIndex        =   9
      Text            =   "FAsientP.frx":0000
      Top             =   1995
      Width           =   750
   End
   Begin MSDataListLib.DataList DLTP 
      Bindings        =   "FAsientP.frx":0004
      DataSource      =   "AdoTP"
      Height          =   780
      Left            =   2625
      TabIndex        =   2
      Top             =   105
      Width           =   4110
      _ExtentX        =   7250
      _ExtentY        =   1376
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Cancelacion de Polizas"
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
      TabIndex        =   1
      Top             =   525
      Width           =   2325
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Inversión a Corto Plazo"
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
      TabIndex        =   0
      Top             =   105
      Value           =   -1  'True
      Width           =   2325
   End
   Begin VB.TextBox TxtCant 
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
      Left            =   945
      MaxLength       =   12
      MultiLine       =   -1  'True
      TabIndex        =   11
      Text            =   "FAsientP.frx":0018
      Top             =   1995
      Width           =   1590
   End
   Begin MSMask.MaskEdBox MBFechaI 
      Height          =   330
      Left            =   105
      TabIndex        =   4
      Top             =   1260
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
   Begin MSDataGridLib.DataGrid DGAsiento 
      Bindings        =   "FAsientP.frx":001C
      Height          =   2535
      Left            =   105
      TabIndex        =   13
      Top             =   2415
      Width           =   11985
      _ExtentX        =   21140
      _ExtentY        =   4471
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
   Begin MSAdodcLib.Adodc AdoAsiento 
      Height          =   330
      Left            =   105
      Top             =   5040
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
      Caption         =   "Asiento"
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
   Begin VB.CommandButton Command7 
      Caption         =   "&Grabar"
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
      Left            =   11235
      Picture         =   "FAsientP.frx":0035
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   105
      Width           =   855
   End
   Begin VB.CommandButton Command3 
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
      Left            =   11235
      Picture         =   "FAsientP.frx":0477
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   945
      Width           =   855
   End
   Begin MSAdodcLib.Adodc AdoSubCta 
      Height          =   330
      Left            =   210
      Top             =   3465
      Visible         =   0   'False
      Width           =   2220
      _ExtentX        =   3916
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
      Caption         =   "SubCta"
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
      Top             =   3780
      Visible         =   0   'False
      Width           =   2220
      _ExtentX        =   3916
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
   Begin MSAdodcLib.Adodc AdoBanco 
      Height          =   330
      Left            =   210
      Top             =   4095
      Visible         =   0   'False
      Width           =   2220
      _ExtentX        =   3916
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
   Begin MSAdodcLib.Adodc AdoCaja 
      Height          =   330
      Left            =   2415
      Top             =   3150
      Visible         =   0   'False
      Width           =   2220
      _ExtentX        =   3916
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
      Caption         =   "Caja"
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
   Begin MSAdodcLib.Adodc AdoAsientoB 
      Height          =   330
      Left            =   2415
      Top             =   3465
      Visible         =   0   'False
      Width           =   2220
      _ExtentX        =   3916
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
      Caption         =   "AsientoB"
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
      Left            =   2415
      Top             =   3780
      Visible         =   0   'False
      Width           =   2220
      _ExtentX        =   3916
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
   Begin MSAdodcLib.Adodc AdoTP 
      Height          =   330
      Left            =   210
      Top             =   3150
      Visible         =   0   'False
      Width           =   2220
      _ExtentX        =   3916
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
      Caption         =   "TP"
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
   Begin MSDataListLib.DataCombo DCCliente 
      Bindings        =   "FAsientP.frx":0D41
      DataSource      =   "AdoCliente"
      Height          =   315
      Left            =   1470
      TabIndex        =   6
      Top             =   1260
      Width           =   6210
      _ExtentX        =   10954
      _ExtentY        =   556
      _Version        =   393216
      Text            =   "Cliente"
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
   Begin VB.Label LblConcepto 
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
      Height          =   645
      Left            =   2625
      TabIndex        =   12
      Top             =   1680
      Width           =   7995
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Interés"
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
      Top             =   1680
      Width           =   750
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Cantidad"
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
      Left            =   945
      TabIndex        =   10
      Top             =   1680
      Width           =   1590
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Beneficiario:"
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
      TabIndex        =   5
      Top             =   945
      Width           =   6210
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Desde"
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
      TabIndex        =   3
      Top             =   945
      Width           =   1275
   End
   Begin VB.Label LabelHaber 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
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
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   9555
      TabIndex        =   16
      Top             =   5040
      Width           =   1800
   End
   Begin VB.Label LabelDebe 
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
      Left            =   7770
      TabIndex        =   17
      Top             =   5040
      Width           =   1800
   End
   Begin VB.Label LabelDiferencia 
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
      TabIndex        =   18
      Top             =   5040
      Width           =   1695
   End
   Begin VB.Label Label19 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Totales"
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
      TabIndex        =   20
      Top             =   5040
      Width           =   1065
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Diferencia"
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
      Left            =   3780
      TabIndex        =   19
      Top             =   5040
      Width           =   1170
   End
End
Attribute VB_Name = "FPolizas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub Insertar_Montos(TipoProc As String, _
                           CuentaNo As String, _
                           TDebe As Currency, _
                           THaber As Currency)
  If CuentaNo <> "00000000-0" Then
     SaldoDisp = 0: SaldoCont = 0: ID_Trans = 0
     TiempoTexto = Format(Time, FormatoTimes)
     If NumeroLineas <= 0 Then NumeroLineas = 1
     SaldoCont = 0
     SaldoDisp = 0
    'Insertar Transacciones de Libreta
     sSQL = "SELECT TOP 1 * " _
          & "FROM Trans_Libretas " _
          & "WHERE Cuenta_No = '" & CuentaNo & "' " _
          & "ORDER BY Fecha DESC,IDT DESC,Hora DESC,ID DESC "
     SelectAdodc AdoAux, sSQL
     NumeroLineas = 0
     With AdoAux.Recordset
      If .RecordCount > 0 Then
          SaldoDisp = .Fields("Saldo_Disp")
          SaldoCont = .Fields("Saldo_Cont")
          Cartilla_No = .Fields("Cartilla_No")
          NumeroLineas = .Fields("ID") + 1
          ID_Trans = .Fields("IDT")
      End If
      If NumeroLineas >= 36 Then NumeroLineas = 1
     .AddNew
     .Fields("Fecha") = MBFechaI
     .Fields("Cuenta_No") = CuentaNo
     .Fields("TP") = TipoProc
     .Fields("Debitos") = TDebe
     .Fields("Creditos") = THaber
     .Fields("Saldo_Cont") = SaldoCont + THaber - TDebe
     .Fields("Saldo_Disp") = SaldoCont + THaber - TDebe
     .Fields("CodigoU") = CodigoUsuario
     .Fields("IDT") = ID_Trans + 1
     .Fields("Hora") = TiempoTexto
     .Fields("Item") = NumEmpresa
     .Fields("Cartilla_No") = Cartilla_No
      SetUpdate AdoAux
     End With
  End If
End Sub

Private Sub Command3_Click()
  Unload Me
End Sub

Private Sub Command7_Click()
  Control_Procesos Normal, "Inversion a Plazo Fijo " & MBFechaI
  FechaIni = BuscarFecha(MBFechaI.Text)
  SumaDebe = 0: SumaHaber = 0
  DGAsiento.Visible = False
  With AdoAsiento.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       Do While Not .EOF
          SumaDebe = SumaDebe + .Fields("DEBE")
          SumaHaber = SumaHaber + .Fields("HABER")
         .MoveNext
       Loop
      .MoveFirst
       If Round(SumaDebe - SumaHaber, 2) = 0 Then
          RatonReloj
          Co.T = Normal
          If Option1.value Then
             Credito_No = NumEmpresa & Format(ReadSetDataNum("Polizas", True, True), "0000000")
             Co.TP = CompIngreso
             Co.Numero = ReadSetDataNum("Ingresos", True, True)
             Co.Efectivo = Total + TotalInteres
             Co.Monto_Total = Total + TotalInteres
          Else
             Co.TP = CompEgreso
             Co.Numero = ReadSetDataNum("Egresos", True, True)
             Co.CodigoB = Ninguno
             Co.Efectivo = Abono
             Co.Monto_Total = Abono
          End If
          Co.CodigoB = CodigoCliente
          Co.Fecha = MBFechaI
          Co.Item = NumEmpresa
          Co.Usuario = CodigoUsuario
          Co.Concepto = LblConcepto.Caption & " Poliza No. " & Credito_No
          Co.T_No = Trans_No
          GrabarComprobante Co
          If Option1.value Then
             SetAdoAddNew "Prestamos"
             SetAdoFields "T", "I"
             SetAdoFields "Tasa", Interes * 100
             SetAdoFields "TP", TipoDoc
             SetAdoFields "ME", False
             SetAdoFields "Credito_No", Credito_No
             SetAdoFields "Cuenta_No", CodigoCliente
             SetAdoFields "Meses", 0
             SetAdoFields "Dia", NoDias
             SetAdoFields "Fecha", MBFechaI
             SetAdoFields "Interes", TotalInteres
             SetAdoFields "Capital", Total
             SetAdoFields "Pagos", TotalInteres
             SetAdoFields "Saldo_Pendiente", Total + TotalInteres
             SetAdoFields "Encaje", 0
             SetAdoFields "Plazo", NoDias
             SetAdoFields "Item", NumEmpresa
             SetAdoUpdate
            'Debitar si es libreta
            If Contra_Cta = Cta_Libretas Then
               Insertar_Montos "N/DP", LstCuentas.Text, Total, 0
            Else
                SetAdoAddNew "Trans_Libretas"
                SetAdoFields "T", Normal
                SetAdoFields "ME", False
                SetAdoFields "Fecha", FechaSistema
                SetAdoFields "Cuenta_No", Credito_No
                SetAdoFields "TP", "BOVE"
                SetAdoFields "Debitos", 0
                SetAdoFields "Creditos", Total
                SetAdoFields "CodigoU", CodigoUsuario
                SetAdoFields "Hora", Format(Time, FormatoTimes)
                SetAdoFields "Item", NumEmpresa
                SetAdoFields "Cheque", ""
                SetAdoFields "Banco", "D.P.F. No. " & Credito_No
                SetAdoFields "ACC", CBool(adFalse)
                SetAdoFields "CHT", CBool(adFalse)
                SetAdoUpdate
            End If
          Else
            'Acreditar si el libreta
             SQL2 = "UPDATE Prestamos " _
                  & "SET T = 'C' " _
                  & "WHERE Credito_No = '" & Credito_No & "' " _
                  & "AND Item = '" & NumEmpresa & "' "
             ConectarAdoExecute SQL2
             If Contra_Cta = Cta_Libretas Then
                Insertar_Montos "N/CP", LstCuentas.Text, 0, Abono
             Else
                SetAdoAddNew "Trans_Libretas"
                SetAdoFields "T", Normal
                SetAdoFields "ME", False
                SetAdoFields "Fecha", FechaSistema
                SetAdoFields "Cuenta_No", Credito_No
                SetAdoFields "TP", "BOVE"
                SetAdoFields "Debitos", Abono
                SetAdoFields "Creditos", 0
                SetAdoFields "CodigoU", CodigoUsuario
                SetAdoFields "Hora", Format(Time, FormatoTimes)
                SetAdoFields "Item", NumEmpresa
                SetAdoFields "Cheque", ""
                SetAdoFields "Banco", "C.D.P.F. No. " & Credito_No
                SetAdoFields "ACC", CBool(adFalse)
                SetAdoFields "CHT", CBool(adFalse)
                SetAdoUpdate
             End If
          End If
          ImprimirComprobantesDe False, Co
          Unload Me
       Else
         MsgBox "Las Transacciones no cuadran"
         DGAsiento.Visible = True
       End If
   Else
      MsgBox "No Existen Datos"
      DGAsiento.Visible = True
   End If
  End With
End Sub

Private Sub DCCliente_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCCliente_LostFocus()
  LstCuentas.Clear
  LstCuentas.AddItem "EFECTIVO"
  LstCuentas.Text = "EFECTIVO"
  Total = 0
  TotalInteres = 0
  Total_Pagar = 0
  With AdoCliente.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Beneficiario = '" & DCCliente & "'")
       If Not .EOF Then
          CodigoBenef = .Fields("Codigo")
          CodigoCliente = .Fields("Codigo")
          NombreCliente = .Fields("Cliente")
          sSQL = "SELECT * " _
               & "FROM Clientes_Datos_Extras " _
               & "WHERE Codigo = '" & CodigoBenef & "' " _
               & "AND Tipo_Dato = 'LIBRETAS' " _
               & "AND T <> 'A' " _
               & "ORDER BY Cuenta_No "
          SelectAdodc AdoAux, sSQL
          If AdoAux.Recordset.RecordCount > 0 Then
             Do While Not AdoAux.Recordset.EOF
                LstCuentas.AddItem AdoAux.Recordset.Fields("Cuenta_No")
                AdoAux.Recordset.MoveNext
             Loop
          End If
       End If
   End If
  End With
End Sub

Private Sub DLTP_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DLTP_LostFocus()
  TipoDoc = SinEspaciosIzq(DLTP)
  NoDias = 0
  With AdoTP.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("CTP = '" & TipoDoc & "' ")
       If Not .EOF Then
          NoDias = .Fields("Dias")
          TxtInteres = Format(.Fields("Interes"), "#,##0.00")
       End If
   End If
  End With
End Sub

Private Sub Form_Activate()
  'MsgBox NumEmpresa & String(5, "0") & "-0"
  Trans_No = 18
  FechaValida MBFechaI
  FechaIni = BuscarFecha(MBFechaI.Text)
  IniciarAsientosDe DGAsiento, AdoAsiento
  sSQL = "SELECT CTP & '  ' & Descripcion As TipoP,* " _
       & "FROM Catalogo_Prestamo " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND TC <> " & Val(adFalse) & " " _
       & "AND MidStrg(CTP,1,2) = 'PO' " _
       & "ORDER BY CTP "
  SelectDBList DLTP, AdoTP, sSQL, "TipoP"
  Listar_Polizas
  RatonNormal
End Sub

Private Sub Form_Load()
  CentrarForm FPolizas
  ConectarAdodc AdoTP
  ConectarAdodc AdoAux
  ConectarAdodc AdoCaja
  ConectarAdodc AdoBanco
  ConectarAdodc AdoSubCta
  ConectarAdodc AdoCliente
  ConectarAdodc AdoAsiento
  ConectarAdodc AdoAsientoB
End Sub

Private Sub LstCuentas_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub LstCuentas_LostFocus()
  With AdoCliente.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Beneficiario = '" & DCCliente & "'")
       If Not .EOF Then
          CodigoBenef = .Fields("Codigo")
          CodigoCliente = .Fields("Codigo")
          NombreCliente = .Fields("Cliente")
          If IsNumeric(SinEspaciosDer(DCCliente)) Then
             Credito_No = SinEspaciosDer(DCCliente)
             TipoDoc = .Fields("TP")
             Fecha_Vence = CLongFecha(CFechaLong(.Fields("Fecha_Pol")) + .Fields("Plazo"))
             FechaCorte = .Fields("Fecha_Pol")
             Total = .Fields("Capital")
             TotalInteres = .Fields("Interes")
             Total_Pagar = .Fields("Pagos")
             TxtInteres = .Fields("Tasa")
             Procesar_Asiento_Poliza
          End If
       End If
   End If
  End With
  
End Sub

Private Sub MBFechaI_GotFocus()
  MarcarTexto MBFechaI
End Sub

Private Sub MBFechaI_LostFocus()
  FechaValida MBFechaI
End Sub

Private Sub Option1_Click()
  TxtCant.Enabled = True
  TxtInteres.Enabled = True
  Listar_Polizas
End Sub

Private Sub Option2_Click()
  TxtCant.Enabled = False
  TxtInteres.Enabled = False
  Listar_Polizas
End Sub

Private Sub TxtCant_GotFocus()
  MarcarTexto TxtCant
End Sub

Private Sub TxtCant_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Public Sub Listar_Polizas()
  If Option1.value Then
     sSQL = "SELECT C.Cliente As Beneficiario,C.* " _
          & "FROM Clientes As C " _
          & "WHERE C.Codigo <> '.' " _
          & "ORDER BY C.Cliente "
  Else
     sSQL = "SELECT (C.Cliente & ' - ' & P.Credito_No) As Beneficiario,C.*,P.Fecha As Fecha_Pol," _
          & "P.Cuenta_No,P.T As T1,P.Credito_No,P.Plazo,P.Capital,P.Interes,P.Pagos,P.TP,P.Tasa " _
          & "FROM Clientes As C,Prestamos As P " _
          & "WHERE C.Codigo = P.Cuenta_No " _
          & "AND P.T = 'I' " _
          & "ORDER BY C.Cliente "
  End If
  SelectDBCombo DCCliente, AdoCliente, sSQL, "Beneficiario"
  RatonNormal
End Sub

Public Sub Procesar_Asiento_Poliza()
  If LstCuentas.Text = "EFECTIVO" Then
     Contra_Cta = Cta_CajaG
  Else
     Contra_Cta = Cta_Libretas
  End If

  SQL2 = "SELECT * " _
       & "FROM Asiento " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND T_No = " & Trans_No & " " _
       & "AND CodigoU = '" & CodigoUsuario & "' "
  SelectDataGrid DGAsiento, AdoAsiento, SQL2
  IniciarAsientosDe DGAsiento, AdoAsiento
  FechaValida MBFechaI
  TextoValido TxtCant, True
  FechaIni = BuscarFecha(MBFechaI)
  RatonReloj
 'Verificamos si existe los prestamos
  With AdoTP.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("CTP = '" & TipoDoc & "' ")
       If Not .EOF Then
          NoDias = .Fields("Dias")
          Interes = Val(TxtInteres) / 100       ' Interes = .Fields("Interes") / 100
          Codigo1 = .Fields("Cta_Interes")      ' Interes a plazo fijo
          Codigo2 = .Fields("Cta_Int_Ganado")   ' Inversiones a Plazo Fijo
          Codigo3 = .Fields("Cta_Int_Efec")     ' Interes por pagar
          Codigo4 = .Fields("Cta_Gas_Oper")     ' Gastos Operativos
          'Cta_Solca = .Fields("Cta_Solca")
          Cta_Ret_Egreso = .Fields("Cta_Impuesto")
          If Option1.value Then
             Credito_No = NumEmpresa & Format(ReadSetDataNum("Polizas", True, False), "0000000")
             Total = Val(CCur(TxtCant))
             TotalInteres = Round(((Total * Interes) / 360) * NoDias, 2)
             InsertarAsientos AdoAsiento, Contra_Cta, 0, Total, 0
             InsertarAsientos AdoAsiento, Codigo2, 0, 0, Total
          Else
             Total_Interes = 0
             If CFechaLong(Fecha_Vence) > CFechaLong(FechaSistema) Then
                NoDias = CFechaLong(FechaSistema) - CFechaLong(FechaCorte)
                Total_Interes = Round(((Total * Interes) / 360) * NoDias, 2)
                'InsertarAsientos AdoAsiento, Codigo1, 0, 0, TotalInteres - Total_Interes
                Total_Interes = TotalInteres - Total_Interes
             End If
            'MsgBox NoDias & vbCrLf & TotalInteres & vbCrLf & Total_Interes
             TxtCant.Enabled = False
             Credito = Round(TotalInteres * 0.006, 2)
             Total_DetRet = Round(TotalInteres * 0.02, 2)
             
            'MsgBox Total & vbCrLf & TotalInteres
'''''             Select Case NoDias
'''''               Case 1 To 30:  Total_Comision = 0.8 / 100      'Total
'''''               Case 31 To 60: Total_Comision = 1 / 100
'''''               Case 61 To 90: Total_Comision = 1.5 / 100
'''''               Case Else:     Total_Comision = 2 / 100
'''''             End Select
'''''             Total_Comision = Round(TotalInteres * Total_Comision, 2)
             Total_Comision = Round(TotalInteres * .Fields("Interes_Gastos_Operativos"), 2)
             'Total_Comision = Total_DetRet
            'MsgBox Total
             InsertarAsientos AdoAsiento, Codigo2, 0, Total, 0
             InsertarAsientos AdoAsiento, Codigo3, 0, TotalInteres - Total_Interes, 0
            'InsertarAsientos AdoAsiento, Cta_Solca, 0, 0, Credito
             InsertarAsientos AdoAsiento, Cta_Ret_Egreso, 0, 0, Total_DetRet
             InsertarAsientos AdoAsiento, Codigo4, 0, 0, Total_Comision         'El 2% sobre el capital
            'MsgBox Total_Comision
             SumaDebe = 0: SumaHaber = 0
             With AdoAsiento.Recordset
              If .RecordCount > 0 Then
                 .MoveFirst
                  Do While Not .EOF
                     SumaDebe = SumaDebe + .Fields("DEBE")
                     SumaHaber = SumaHaber + .Fields("HABER")
                    .MoveNext
                  Loop
              End If
             End With
             Abono = (SumaDebe - SumaHaber)
             If Abono < 0 Then Abono = -Abono
             InsertarAsientos AdoAsiento, Contra_Cta, 0, 0, Abono
          End If
       End If
   End If
  End With
  If Option1.value Then
     LblConcepto.Caption = "Depósito a Plazo Fijo del Sr(a) " & NombreCliente & " Plazo a " _
                         & NoDias & " días, al " & Interes * 100 & "% anual,"
  Else
     LblConcepto.Caption = "Cancelación de Plazo Fijo a " & NoDias & " días del Sr(a) " & NombreCliente _
                         & ", al " & Interes * 100 & "% anual, "
  End If
  SumaDebe = 0: SumaHaber = 0
  With AdoAsiento.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       Do While Not .EOF
          SumaDebe = SumaDebe + .Fields("DEBE")
          SumaHaber = SumaHaber + .Fields("HABER")
         .MoveNext
       Loop
   End If
  End With
  LabelDebe.Caption = Format(SumaDebe, "#,##0.00")
  LabelHaber.Caption = Format(SumaHaber, "#,##0.00")
  LabelDiferencia.Caption = Format(SumaDebe - SumaHaber, "#,##0.00")
  DGAsiento.Visible = True
  RatonNormal
  Command7.SetFocus
End Sub

Private Sub TxtCant_LostFocus()
  If LstCuentas.Text <> "EFECTIVO" Then
     SaldoDisp = 0
     sSQL = "SELECT TOP 1 * " _
          & "FROM Trans_Libretas " _
          & "WHERE Cuenta_No = '" & LstCuentas.Text & "' " _
          & "ORDER BY Fecha DESC,IDT DESC,Hora DESC,ID DESC "
     SelectAdodc AdoAux, sSQL
     If AdoAux.Recordset.RecordCount > 0 Then SaldoDisp = AdoAux.Recordset.Fields("Saldo_Disp")
     If CCur(TxtCant) > SaldoDisp Then
        MsgBox "ESTA CUENTA NO TIENE FONDOS SUFICIENTES PARA REALIZAR LA TRANSACCION"
     End If
  End If
  Procesar_Asiento_Poliza
End Sub

Private Sub TxtInteres_GotFocus()
  MarcarTexto TxtInteres
End Sub

Private Sub TxtInteres_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtInteres_LostFocus()
  TextoValido TxtInteres
End Sub
