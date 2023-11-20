VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Begin VB.Form Ing_Combo 
   Caption         =   "CONTROL DE INVENTARIO PARA INGRESOS/EGRESOS"
   ClientHeight    =   7335
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11460
   Icon            =   "Ing_Comb.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7335
   ScaleWidth      =   11460
   WindowState     =   2  'Maximized
   Begin VB.TextBox TxtLoteNo 
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
      IMEMode         =   3  'DISABLE
      Left            =   11655
      MaxLength       =   10
      MultiLine       =   -1  'True
      TabIndex        =   15
      Text            =   "Ing_Comb.frx":0442
      Top             =   1155
      Width           =   1170
   End
   Begin VB.TextBox TxtRegSanitario 
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
      IMEMode         =   3  'DISABLE
      Left            =   15330
      MaxLength       =   25
      TabIndex        =   21
      Text            =   "."
      Top             =   1155
      Width           =   2115
   End
   Begin VB.TextBox TxtModelo 
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
      IMEMode         =   3  'DISABLE
      Left            =   11655
      MaxLength       =   25
      TabIndex        =   23
      Text            =   "."
      Top             =   1890
      Width           =   3690
   End
   Begin VB.TextBox TxtProcedencia 
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
      IMEMode         =   3  'DISABLE
      Left            =   15330
      MaxLength       =   25
      TabIndex        =   25
      Text            =   "."
      Top             =   1890
      Width           =   2115
   End
   Begin VB.TextBox TxtSerieNo 
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
      IMEMode         =   3  'DISABLE
      Left            =   11655
      MaxLength       =   25
      TabIndex        =   27
      Text            =   "."
      Top             =   2625
      Width           =   3270
   End
   Begin MSDataListLib.DataCombo DCTInv 
      Bindings        =   "Ing_Comb.frx":0444
      DataSource      =   "AdoTInv"
      Height          =   315
      Left            =   6405
      TabIndex        =   12
      Top             =   1155
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16761024
      Text            =   "DC"
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
   Begin VB.TextBox TxtCodBar 
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
      IMEMode         =   3  'DISABLE
      Left            =   14385
      MaxLength       =   25
      TabIndex        =   35
      Text            =   "."
      Top             =   3360
      Width           =   3060
   End
   Begin MSDataListLib.DataCombo DCBodega 
      Bindings        =   "Ing_Comb.frx":045A
      DataSource      =   "AdoBodega"
      Height          =   315
      Left            =   105
      TabIndex        =   8
      Top             =   1155
      Width           =   3165
      _ExtentX        =   5583
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16761024
      Text            =   "DC"
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
   Begin MSDataGridLib.DataGrid DGKardex 
      Bindings        =   "Ing_Comb.frx":0472
      Height          =   4215
      Left            =   210
      TabIndex        =   36
      Top             =   3885
      Width           =   12300
      _ExtentX        =   21696
      _ExtentY        =   7435
      _Version        =   393216
      BackColor       =   16761024
      BorderStyle     =   0
      ForeColor       =   255
      HeadLines       =   1
      RowHeight       =   15
      AllowDelete     =   -1  'True
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
   Begin MSDataListLib.DataCombo DCInv 
      Bindings        =   "Ing_Comb.frx":048A
      DataSource      =   "AdoInv"
      Height          =   2325
      Left            =   105
      TabIndex        =   13
      Top             =   1470
      Width           =   11355
      _ExtentX        =   20029
      _ExtentY        =   4101
      _Version        =   393216
      Style           =   1
      BackColor       =   16744576
      ForeColor       =   16777215
      Text            =   "Productos"
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
   Begin VB.TextBox TextVUnit 
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
      Left            =   12915
      MultiLine       =   -1  'True
      TabIndex        =   33
      Text            =   "Ing_Comb.frx":049F
      Top             =   3360
      Width           =   1485
   End
   Begin VB.TextBox TextEntrada 
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
      IMEMode         =   3  'DISABLE
      Left            =   11655
      MultiLine       =   -1  'True
      TabIndex        =   31
      Text            =   "Ing_Comb.frx":04A1
      Top             =   3360
      Width           =   1275
   End
   Begin MSDataListLib.DataCombo DCBenef 
      Bindings        =   "Ing_Comb.frx":04A3
      DataSource      =   "AdoBenef"
      Height          =   315
      Left            =   3150
      TabIndex        =   3
      Top             =   105
      Width           =   12090
      _ExtentX        =   21325
      _ExtentY        =   556
      _Version        =   393216
      Text            =   "Clientes"
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
   Begin MSAdodcLib.Adodc AdoBenef 
      Height          =   330
      Left            =   2310
      Top             =   6405
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
      Caption         =   "Benef"
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
   Begin MSAdodcLib.Adodc AdoKardex 
      Height          =   330
      Left            =   2310
      Top             =   6720
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
      Caption         =   "Kardex"
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
   Begin MSAdodcLib.Adodc AdoInv 
      Height          =   330
      Left            =   2310
      Top             =   7035
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
      Caption         =   "Inv"
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
   Begin MSAdodcLib.Adodc AdoArt 
      Height          =   330
      Left            =   315
      Top             =   7035
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
      Caption         =   "Art"
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
   Begin MSAdodcLib.Adodc AdoAsientos 
      Height          =   330
      Left            =   315
      Top             =   7350
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
      Caption         =   "Asientos"
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
      Left            =   1155
      Picture         =   "Ing_Comb.frx":04BA
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   8295
      Width           =   960
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
      Left            =   3150
      MaxLength       =   100
      TabIndex        =   6
      Top             =   420
      Width           =   14295
   End
   Begin VB.CommandButton Command1 
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
      Left            =   105
      Picture         =   "Ing_Comb.frx":0D84
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   8295
      Width           =   960
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
   Begin MSAdodcLib.Adodc AdoTInv 
      Height          =   330
      Left            =   315
      Top             =   6720
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
      Caption         =   "TInv"
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
   Begin MSAdodcLib.Adodc AdoBodega 
      Height          =   330
      Left            =   315
      Top             =   6405
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
      Caption         =   "Bodega"
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
   Begin MSAdodcLib.Adodc AdoInvS 
      Height          =   330
      Left            =   315
      Top             =   6720
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
      Caption         =   "InvS"
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
   Begin MSAdodcLib.Adodc AdoTInvS 
      Height          =   330
      Left            =   2310
      Top             =   7350
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
      Caption         =   "TInvS"
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
   Begin MSMask.MaskEdBox MBFechaExp 
      Height          =   330
      Left            =   14070
      TabIndex        =   19
      Top             =   1155
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
   Begin MSMask.MaskEdBox MBFechaFab 
      Height          =   330
      Left            =   12810
      TabIndex        =   17
      Top             =   1155
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
   Begin MSDataListLib.DataCombo DCBodega1 
      Bindings        =   "Ing_Comb.frx":11C6
      DataSource      =   "AdoBodega"
      Height          =   315
      Left            =   3255
      TabIndex        =   10
      Top             =   1155
      Width           =   3165
      _ExtentX        =   5583
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16761024
      Text            =   "DC"
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
   Begin VB.Label Label12 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " PRODUCTO DE ENTRADA Y SU CANTIDAD"
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
      Top             =   840
      Width           =   5055
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " BODEGA DE ENTRADA"
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
      TabIndex        =   9
      Top             =   840
      Width           =   3165
   End
   Begin VB.Label Label24 
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " LOTE No."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   11655
      TabIndex        =   14
      Top             =   840
      Width           =   1170
   End
   Begin VB.Label Label25 
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " FECHA FAB"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   12810
      TabIndex        =   16
      Top             =   840
      Width           =   1275
   End
   Begin VB.Label Label26 
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " FECHA EXP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   14070
      TabIndex        =   18
      Top             =   840
      Width           =   1275
   End
   Begin VB.Label Label27 
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " REG. SANITARIO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   15330
      TabIndex        =   20
      Top             =   840
      Width           =   2115
   End
   Begin VB.Label Label28 
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " MODELO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   11655
      TabIndex        =   22
      Top             =   1575
      Width           =   3690
   End
   Begin VB.Label Label29 
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " PROCEDENCIA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   15330
      TabIndex        =   24
      Top             =   1575
      Width           =   2115
   End
   Begin VB.Label Label30 
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " SERIE No."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   11655
      TabIndex        =   26
      Top             =   2310
      Width           =   3270
   End
   Begin VB.Label Label2 
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
      Left            =   15225
      TabIndex        =   4
      Top             =   105
      Width           =   2220
   End
   Begin VB.Label LabelCodigo 
      BackColor       =   &H80000005&
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
      Left            =   14910
      TabIndex        =   29
      Top             =   2625
      Width           =   2535
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
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
      Height          =   330
      Left            =   7980
      TabIndex        =   38
      Top             =   8295
      Width           =   1695
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "VALOR UNITARIO"
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
      Left            =   6195
      TabIndex        =   37
      Top             =   8295
      Width           =   1800
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " BODEGA DE SALIDA"
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
      Top             =   840
      Width           =   3165
   End
   Begin VB.Label Label17 
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " CODIGO DE BARRA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   14385
      TabIndex        =   34
      Top             =   3045
      Width           =   3060
   End
   Begin VB.Label Label15 
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " CODIGO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   14910
      TabIndex        =   28
      Top             =   2310
      Width           =   2535
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " CONCEPTO"
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
      Top             =   420
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
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
      Height          =   330
      Left            =   11445
      TabIndex        =   40
      Top             =   8295
      Width           =   1695
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " VALOR UNIT."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   12915
      TabIndex        =   32
      Top             =   3045
      Width           =   1485
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " CANTIDAD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   11655
      TabIndex        =   30
      Top             =   3045
      Width           =   1275
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "T O T A L"
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
      Left            =   9765
      TabIndex        =   39
      Top             =   8295
      Width           =   1695
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " FECHA:"
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
      Caption         =   " RESPONSABLE"
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
      Top             =   105
      Width           =   1695
   End
End
Attribute VB_Name = "Ing_Combo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Cta_Inv_S As String
Dim Producto_S As String
Dim A_No As Integer
Dim NoExiste As Boolean

Public Sub ListarProductosES(Optional TipoInvSalida As Boolean)
 CodigoInv = SinEspaciosIzq(DCTInv.Text)
 sSQL = "SELECT CP.Codigo_Inv, CP.Producto, CP.Cta_Inventario, CP.Cta_Costo_Venta, CP.IVA, CP.Unidad, CP.Reg_Sanitario " _
      & "FROM Catalogo_Productos As CP, Catalogo_Recetas AS CR " _
      & "WHERE CP.Item = '" & NumEmpresa & "' " _
      & "AND CP.Periodo = '" & Periodo_Contable & "' " _
      & "AND CP.Codigo_Inv LIKE '%" & CodigoInv & "%' " _
      & "AND LEN(CP.Cta_Inventario) > 1 " _
      & "AND CP.TC = 'P' " _
      & "AND CP.Codigo_Inv = CR.Codigo_PP " _
      & "AND CP.Item = CR.Item " _
      & "AND CP.Periodo = CR.Periodo " _
      & "GROUP BY CP.Codigo_Inv, CP.Producto, CP.Cta_Inventario, CP.Cta_Costo_Venta, CP.IVA,CP.Unidad, CP.Reg_Sanitario " _
      & "ORDER BY CP.Producto "
 SelectDB_Combo DCInv, AdoInv, sSQL, "Producto"
End Sub

Private Sub Command1_Click()
  RatonReloj
  Total_RetCta = 0
  RatonNormal
  Mensajes = "Seguro de Grabar?"
  Titulo = "GRABACION DEL COMPROBANTE"
  If BoxMensaje = vbYes Then
  RatonReloj
  FechaComp = MBFechaI
  FechaTexto = MBFechaI
  sSQL = "DELETE * " _
       & "FROM Asiento " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND T_No = " & Trans_No & " " _
       & "AND CodigoU = '" & CodigoUsuario & "' "
  Ejecutar_SQL_SP sSQL
  sSQL = "DELETE * " _
       & "FROM Asiento_R " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND T_No = " & Trans_No & " " _
       & "AND CodigoU = '" & CodigoUsuario & "' "
  Ejecutar_SQL_SP sSQL
  sSQL = "SELECT * " _
       & "FROM Asiento_K " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND T_No = " & Trans_No & " " _
       & "ORDER BY CTA_INVENTARIO,CONTRA_CTA "
  SQLDec = "VALOR_UNIT " & CStr(Dec_Costo) & "|,VALOR_TOTAL 4|,CANTIDAD 2|,SALDO 4|."
  Select_Adodc_Grid DGKardex, AdoKardex, sSQL, SQLDec
  RatonReloj
  Si_No = True
  CodigoBenef = Ninguno
  With AdoBenef.Recordset
   If .RecordCount > 0 Then
       Cadena = DCBenef.Text
      .MoveFirst
      .Find ("Cliente Like '" & Cadena & "' ")
       If Not .EOF Then CodigoBenef = .Fields("Codigo")
   End If
  End With
  Valor = 0
  TotalInventario
  FechaTexto = MBFechaI.Text
  With AdoKardex.Recordset
   If .RecordCount > 0 Then
       RatonReloj
      'Llenamos los datos ingresados al Kardex
       TipoDoc = Day(FechaSistema) & "-" & Month(FechaSistema) & "-" & Cod_Bodega & "-" & Cod_Bodega1
      .MoveFirst
       Total = 0: ValorDH = 0
       Cta_Inventario = .Fields("CTA_INVENTARIO")
       Do While Not .EOF
          ValorDH = .Fields("VALOR_TOTAL")
          Cta_Inventario = .Fields("CTA_INVENTARIO")
          If .Fields("DH") = 1 Then
              InsertarAsientos AdoAsientos, Cta_Inventario, 0, ValorDH, 0
          Else
              InsertarAsientos AdoAsientos, Cta_Inventario, 0, 0, ValorDH
          End If
          Total = Total + .Fields("VALOR_TOTAL")
         .MoveNext
       Loop
       Total = Round(Total, 2)
       ValorDH = Round(ValorDH, 2)
   End If
  End With
  sSQL = "SELECT * " _
       & "FROM Asiento " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND T_No = " & Trans_No & " "
  Select_Adodc AdoAsientos, sSQL
  With AdoAsientos.Recordset
   If .RecordCount > 0 Then
       Debe = 0: Haber = 0
       Do While Not .EOF
          Debe = Debe + .Fields("DEBE")
          Haber = Haber + .Fields("HABER")
         .MoveNext
       Loop
       If (Debe - Haber) <> 0 Then MsgBox "Verifique el comprobante, no cuadra por: " & Round(Debe - Haber, 2)
      .MoveFirst
       'MsgBox (Debe - Haber)
       'MsgBox "Hola"
       NumComp = ReadSetDataNum("Diario", True, True)
       RatonReloj
          Co.T = Normal
          Co.TP = CompDiario
          Co.Fecha = FechaTexto
          Co.Numero = NumComp
          Co.Concepto = TextConcepto
          Co.CodigoB = CodigoCliente
          Co.Efectivo = 0
          Co.Monto_Total = Total
          Co.Usuario = CodigoUsuario
          Co.T_No = Trans_No
          Co.Item = NumEmpresa
          GrabarComprobante Co
          ImprimirComprobantesDe False, Co
          Imprimir_Nota_Inventario AdoBenef, AdoKardex, NumComp, TipoDoc, "CD", FechaTexto, FechaTexto, Total
          Mensajes = "Imprimir Copia de Nota de Entrada/Salida"
          Titulo = "COPIA DE NOTA"
          If BoxMensaje = vbYes Then Imprimir_Nota_Inventario AdoBenef, AdoKardex, NumComp, TipoDoc, "CD", FechaTexto, FechaTexto, Total
          Unload Ing_Combo
   End If
  End With
  End If
  RatonNormal
End Sub

Private Sub Command2_Click()
  Unload Ing_Combo
End Sub

Private Sub DCBenef_LostFocus()
  InvImp = False
  Label2.Caption = " LOCAL"
  CodigoBenef = Ninguno
  CodigoCliente = Ninguno
  With AdoBenef.Recordset
   If .RecordCount > 0 Then
       Cadena = DCBenef.Text
      .MoveFirst
      .Find ("Cliente Like '" & Cadena & "' ")
       If Not .EOF Then
          CodigoBenef = .Fields("Codigo")
          CodigoCliente = .Fields("Codigo")
          TipoDoc = .Fields("TD")
          Si_No = True
          If TipoDoc = "R" Then Si_No = False
       Else
          Si_No = False
       End If
   End If
  End With
  TextConcepto.Text = DCBenef.Text & " "
End Sub

Private Sub DCBodega_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCBodega_LostFocus()
  Cod_Bodega = Ninguno
  With AdoBodega.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Bodega Like '" & DCBodega & "' ")
       If Not .EOF Then Cod_Bodega = .Fields("CodBod")
   End If
  End With
End Sub

Private Sub DCBodega1_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCBodega1_LostFocus()
  Cod_Bodega1 = Ninguno
  With AdoBodega.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Bodega Like '" & DCBodega1 & "' ")
       If Not .EOF Then Cod_Bodega1 = .Fields("CodBod")
   End If
  End With
End Sub

Private Sub DCInv_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  PresionoEnter KeyCode
  If KeyCode = vbKeyEscape Then Command1.SetFocus
 End Sub

Private Sub DCInv_LostFocus()
  If Len(Cod_Bodega) <= 1 Then Cod_Bodega = "01"
  If Len(Cod_Bodega1) <= 1 Then Cod_Bodega1 = "01"
  
  Contra_Cta = Ninguno
  Contra_Cta1 = Ninguno
  CodigoInv = DCInv.Text
  Si_No = False
  With AdoInv.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Producto Like '" & CodigoInv & "' ")
       If Not .EOF Then
          Si_No = .Fields("IVA")
          Unidad = .Fields("Unidad")
          CodigoInv = .Fields("Codigo_Inv")
          Producto_S = .Fields("Producto")
          Cta_Inv_S = .Fields("Cta_Inventario")
          TxtRegSanitario = .Fields("Reg_Sanitario")
         ' TextEntrada.SetFocus
          TxtLoteNo.SetFocus
       Else
          MsgBox "Este codgo no existe"
          DCInv.SetFocus
       End If
   Else
       MsgBox "Falta asignar proceso para Combos"
       DCTInv.SetFocus
   End If
  End With
End Sub


Private Sub DCTInv_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub DCTInv_LostFocus()
   ListarProductosES True
End Sub

Private Sub DGKardex_BeforeDelete(Cancel As Integer)
  Cancel = DeleteSiNo(AdoKardex)
End Sub

Private Sub DGKardex_DblClick()
Dim ColK(7) As String
   
'''   TxtLoteNo = ""
'''   MBFechaFab = FechaSistema
'''   MBFechaExp = FechaSistema
'''   TxtRegSanitario = ""
'''   TxtModelo = ""
'''   TxtProcedencia = ""
'''   TxtSerieNo = ""
   
   TxtLoteNo = DGKardex.Columns(53).Text
   MBFechaFab = DGKardex.Columns(54).Text
   MBFechaExp = DGKardex.Columns(55).Text
   TxtRegSanitario = DGKardex.Columns(56).Text
   TxtModelo = DGKardex.Columns(57).Text
   TxtProcedencia = DGKardex.Columns(58).Text
   TxtSerieNo = DGKardex.Columns(59).Text
   
   sSQL = "UPDATE Asiento_K " _
        & "SET Lote_No='" & TxtLoteNo & "', Fecha_Fab=#" & BuscarFecha(MBFechaFab) & "#, Fecha_Exp=#" & BuscarFecha(MBFechaExp) & "#, " _
        & "Reg_Sanitario='" & TxtRegSanitario & "', Modelo='" & TxtModelo & "', Procedencia='" & TxtProcedencia & "', " _
        & "Serie_No = '" & TxtSerieNo & "' " _
        & "WHERE CodigoU = '" & CodigoUsuario & "' " _
        & "AND Item = '" & NumEmpresa & "' AND DH = '1' " _
        & "AND Codigo_Inv = '" & CodigoInv & "' "
   Ejecutar_SQL_SP sSQL
   Select_Adodc_Grid_Kardex
End Sub

Private Sub MBFechaExp_GotFocus()
  MarcarTexto MBFechaExp
End Sub

Private Sub MBFechaExp_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBFechaExp_LostFocus()
  FechaValida MBFechaExp, True
End Sub

Private Sub MBFechaFab_GotFocus()
   MarcarTexto MBFechaFab
End Sub

Private Sub MBFechaFab_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBFechaFab_LostFocus()
  FechaValida MBFechaFab
End Sub

Private Sub MBFechaI_GotFocus()
  MarcarTexto MBFechaI
End Sub

Private Sub MBFechaI_LostFocus()
  FechaValida MBFechaI, True
End Sub

Private Sub TextConcepto_GotFocus()
  MarcarTexto TextConcepto
End Sub

Private Sub TextConcepto_LostFocus()
  TextoValido TextConcepto
End Sub

Private Sub TextEntrada_GotFocus()
  Select_Adodc_Grid_Kardex
  SaldoAnterior = 0: ValorUnitAnt = 0: Contador = 0: ValorUnit = 0: Cantidad = 0
  Stock_Actual_Inventario MBFechaI, CodigoInv, Cod_Bodega
  MarcarTexto TextEntrada
End Sub

Private Sub TextEntrada_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextEntrada_LostFocus()
Dim Cod_Inv_Rec As String
Dim Cant_Rec As Currency
    TotalInventario
    If NoExiste Then
         TextoValido TextEntrada, True
         Cantidad = Val(CCur(TextEntrada))
        'Realizamos la salida automatica
         sSQL = "SELECT CP.TC, CR.Codigo_Receta, CP.Cta_Inventario, CP.Cta_Costo_Venta, CP.Reg_Sanitario, " _
              & "CP.Producto, SUM(" & Cantidad & " * CR.Cantidad) As Cant_Salida " _
              & "FROM Catalogo_Recetas As CR, Catalogo_Productos As CP " _
              & "WHERE CP.Periodo = '" & Periodo_Contable & "' " _
              & "AND CP.Item = '" & NumEmpresa & "' " _
              & "AND CR.Codigo_PP = '" & CodigoInv & "' " _
              & "AND CP.Codigo_Inv = CR.Codigo_Receta " _
              & "AND CP.Item = CR.Item " _
              & "AND CP.Periodo = CR.Periodo " _
              & "GROUP BY CP.TC, CR.Codigo_Receta, CP.Cta_Inventario, CP.Cta_Costo_Venta, CP.Producto, CP.Reg_Sanitario " _
              & "ORDER BY CP.TC, CR.Codigo_Receta, CP.Cta_Inventario, CP.Cta_Costo_Venta, CP.Producto, CP.Reg_Sanitario "
         Select_Adodc AdoInvS, sSQL
         With AdoInvS.Recordset
          If .RecordCount > 0 Then
              Do While Not .EOF
                 Salida = Val(CCur(TextEntrada))
                 Cantidad = Salida
                 Cod_Inv_Rec = .Fields("Codigo_Receta")
                 Cant_Rec = .Fields("Cant_Salida")
                 Contra_Cta = .Fields("Cta_Inventario")
                 Producto = .Fields("Producto")
                 Codigo = Leer_Cta_Catalogo(Contra_Cta)  ' Codigo = Cta Contable , Cuenta = Detalle Cuenta
                 
                'MsgBox "Insumos del " & Producto_S & ": " & CodigoInv & " No. " & .RecordCount
                'Resultdo que bota el procedimiento
                'Cantidad = .Fields("Stock")
                'ValorUnit = Redondear(.Fields("TCosto"), Dec_Costo)
                'SaldoAnterior = .Fields("Saldo_Inv")
                 Stock_Actual_Inventario MBFechaI, Cod_Inv_Rec
                 Precio = ValorUnit
                 sSQL = "SELECT Lote_No, Fecha_Fab, Fecha_Exp, Modelo, Procedencia, Serie_No, SUM(Entrada-Salida) As Tot_Stock " _
                      & "FROM Trans_Kardex " _
                      & "WHERE Fecha <= #" & BuscarFecha(MBFechaI) & "# " _
                      & "AND Codigo_Inv = '" & Cod_Inv_Rec & "' " _
                      & "AND Item = '" & NumEmpresa & "' " _
                      & "AND Periodo = '" & Periodo_Contable & "' " _
                      & "AND CodBodega = '" & Cod_Bodega & "' " _
                      & "AND T <> 'A' " _
                      & "GROUP BY Lote_No, Fecha_Fab, Fecha_Exp, Modelo, Procedencia, Serie_No " _
                      & "HAVING SUM(Entrada-Salida) > 0 " _
                      & "ORDER BY Lote_No "
                 Select_Adodc AdoTInvS, sSQL
                'MsgBox Producto & ", " & Cod_Inv_Rec & ": Numeros de Lotes " & .RecordCount
                 If AdoTInvS.Recordset.RecordCount > 0 Then
                    Salida = 1
                    Do While Not AdoTInvS.Recordset.EOF And Salida > 0
                       Salida = AdoTInvS.Recordset.Fields("Tot_Stock")
                      'MsgBox Cod_Inv_Rec & vbCrLf & Cant_Rec & vbCrLf & Salida
                       If Cant_Rec <= Salida Then Salida = Cant_Rec
                       If Salida > 0 And ValorUnit > 0 Then
                          ValorTotal = Redondear(ValorUnit * Salida, 2)
                          SetAddNew AdoKardex
                          SetFields AdoKardex, "CodBod", Cod_Bodega
                          SetFields AdoKardex, "CTA_INVENTARIO", Codigo
                          SetFields AdoKardex, "CODIGO_INV", Cod_Inv_Rec
                          SetFields AdoKardex, "DH", "2"
                          SetFields AdoKardex, "PRODUCTO", Producto
                          SetFields AdoKardex, "CANT_ES", Salida
                          SetFields AdoKardex, "VALOR_UNIT", Round(ValorUnit, Dec_Costo)
                          SetFields AdoKardex, "VALOR_TOTAL", Round(ValorTotal, Dec_Costo)
                          SetFields AdoKardex, "CONTRA_CTA", Ninguno
                          SetFields AdoKardex, "CANTIDAD", Salida
                          SetFields AdoKardex, "VALOR_UNIT", ValorUnit
                          SetFields AdoKardex, "SALDO", Saldo
                          SetFields AdoKardex, "UNIDAD", Unidad
                          SetFields AdoKardex, "Item", NumEmpresa
                          SetFields AdoKardex, "CodigoU", CodigoUsuario
                          SetFields AdoKardex, "T_No", Trans_No
                          SetFields AdoKardex, "SUBCTA", Ninguno 'SubCtaGen
                          SetFields AdoKardex, "TC", Normal 'SubCta
                          SetFields AdoKardex, "COD_BAR", TxtCodBar
                          SetFields AdoKardex, "Codigo_B", CodigoCliente
                          SetFields AdoKardex, "Lote_No", AdoTInvS.Recordset.Fields("Lote_No")
                          SetFields AdoKardex, "Fecha_Fab", AdoTInvS.Recordset.Fields("Fecha_Fab")
                          SetFields AdoKardex, "Fecha_Exp", AdoTInvS.Recordset.Fields("Fecha_Exp")
                          SetFields AdoKardex, "Reg_Sanitario", TxtRegSanitario
                          SetFields AdoKardex, "Modelo", AdoTInvS.Recordset.Fields("Modelo")
                          SetFields AdoKardex, "Procedencia", AdoTInvS.Recordset.Fields("Procedencia")
                          SetFields AdoKardex, "Serie_No", AdoTInvS.Recordset.Fields("Serie_No")
                          SetFields AdoKardex, "A_No", A_No
                          SetUpdate AdoKardex
                          Cant_Rec = Cant_Rec - Salida
                          A_No = A_No + 1
                       End If
                       If Cant_Rec <= 0 Then AdoTInvS.Recordset.MoveLast
                       AdoTInvS.Recordset.MoveNext
                    Loop
                 Else
                    MsgBox "No hay existencia con la Bodega:" & vbCrLf _
                          & vbTab & DCBodega & vbCrLf _
                          & "Del Producto:" & vbCrLf _
                          & vbTab & Producto & vbCrLf _
                          & "Intente con otra Bodega."
                    TextEntr5ada = "0"
                 End If
                .MoveNext
              Loop
          End If
         End With
         TotalInventario
         Entrada = Val(TextEntrada)
         If Entrada > 0 Then
             ValorUnit = Round(Total / Entrada, Dec_Costo)
             ValorTotal = Total
             SetAddNew AdoKardex
             SetFields AdoKardex, "CodBod", Cod_Bodega1
             SetFields AdoKardex, "CTA_INVENTARIO", Cta_Inv_S
             SetFields AdoKardex, "CODIGO_INV", CodigoInv
             SetFields AdoKardex, "DH", "1"
             SetFields AdoKardex, "PRODUCTO", Producto_S
             SetFields AdoKardex, "CANT_ES", Entrada
             SetFields AdoKardex, "VALOR_UNIT", Round(ValorUnit, Dec_Costo)
             SetFields AdoKardex, "VALOR_TOTAL", Round(ValorTotal, Dec_Costo)
             SetFields AdoKardex, "CONTRA_CTA", Ninguno
             SetFields AdoKardex, "CANTIDAD", Entrada
             SetFields AdoKardex, "VALOR_UNIT", ValorUnit
             SetFields AdoKardex, "SALDO", Saldo
             SetFields AdoKardex, "UNIDAD", Unidad
             SetFields AdoKardex, "Item", NumEmpresa
             SetFields AdoKardex, "CodigoU", CodigoUsuario
             SetFields AdoKardex, "T_No", Trans_No
             SetFields AdoKardex, "SUBCTA", Ninguno 'SubCtaGen
             SetFields AdoKardex, "TC", Normal 'SubCta
             SetFields AdoKardex, "COD_BAR", TxtCodBar
             SetFields AdoKardex, "Codigo_B", CodigoCliente
             SetFields AdoKardex, "A_No", 0
             SetUpdate AdoKardex
             Select_Adodc_Grid_Kardex
             TotalInventario
         End If
    Else
       MsgBox "Este producto ya esta procesado, no se puede volver a procesar"
    End If
    LabelCodigo.Caption = CodigoInv
End Sub

Private Sub Form_Activate()
 DGKardex.width = MDI_X_Max - 200
 DGKardex.Height = MDI_Y_Max - DGKardex.Top - 800
 Label1.Top = DGKardex.Top + DGKardex.Height + 30
 Label10.Top = DGKardex.Top + DGKardex.Height + 30
 Label11.Top = DGKardex.Top + DGKardex.Height + 30
 Label13.Top = DGKardex.Top + DGKardex.Height + 30
 Command1.Top = DGKardex.Top + DGKardex.Height + 30
 Command2.Top = DGKardex.Top + DGKardex.Height + 30

' Consultamos las cuentas de la tabla
 Total_IVA = 0
 Label3.Caption = " RESPONSABLE:"
 ListarProveedorUsuario False
 DCBenef.SetFocus
 Trans_No = 80
 A_No = 1
 IniciarAsientosAdo AdoAsientos
 If Inv_Promedio Then
    Ing_Combo.Caption = "CONTROL DE INVENTARIO EN PROMEDIO"
 Else
    Ing_Combo.Caption = "CONTROL DE INVENTARIO EN ULTIMO PRECIO"
 End If
 TipoDoc = CompDiario
'Salida:
 sSQL = "SELECT Codigo_Inv & '  ' & Producto As NomProd " _
      & "FROM Catalogo_Productos " _
      & "WHERE Item = '" & NumEmpresa & "' " _
      & "AND Periodo = '" & Periodo_Contable & "' " _
      & "AND TC = 'I' " _
      & "ORDER BY Codigo_Inv "
 SelectDB_Combo DCTInv, AdoTInv, sSQL, "NomProd"
 
 ListarProductosES
 ListarProductosES True
 ListarProveedorUsuario False
 
 sSQL = "SELECT * " _
      & "FROM Asiento_K " _
      & "WHERE Item = '" & NumEmpresa & "' " _
      & "AND CodigoU = '" & CodigoUsuario & "' " _
      & "AND T_No = " & Trans_No & " "
 SQLDec = "VALOR_UNIT " & CStr(Dec_Costo) & "|,VALOR_TOTAL 4|,CANTIDAD 2|,SALDO 4|."
 Select_Adodc_Grid DGKardex, AdoKardex, sSQL, SQLDec

 sSQL = "SELECT * " _
      & "FROM Catalogo_Bodegas " _
      & "WHERE Item = '" & NumEmpresa & "' " _
      & "AND Periodo = '" & Periodo_Contable & "' " _
      & "ORDER BY CodBod "
 SelectDB_Combo DCBodega, AdoBodega, sSQL, "Bodega"
 SelectDB_Combo DCBodega1, AdoBodega, sSQL, "Bodega"
 Select_Adodc_Grid_Kardex
 FechaValida MBFechaI
 RatonNormal
End Sub

Private Sub Form_Load()
  ConectarAdodc AdoInv
  ConectarAdodc AdoInvS
  ConectarAdodc AdoTInvS
  ConectarAdodc AdoArt
  ConectarAdodc AdoTInv
  ConectarAdodc AdoBenef
  ConectarAdodc AdoKardex
  ConectarAdodc AdoBodega
  ConectarAdodc AdoAsientos
End Sub

Private Sub TextVUnit_GotFocus()
  If Entrada <= 0 Then Entrada = 1
  Precio = Total / Entrada
  TextEntrada = Format(Cantidad, "#,##0.00")
  TextVUnit = Format(Precio, "#,##0." & String(Dec_Costo, "0"))
  MarcarTexto TextVUnit
End Sub

Private Sub TextVUnit_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextVUnit_LostFocus()
  TextoValido TextVUnit, True, , Dec_Costo
  ValorUnit = Val(CCur(TextVUnit.Text))
  ValorTotal = ValorUnit * Entrada
'  TextTotal.Text = Format(ValorTotal, "#,##0.0000")
  TextVUnit.Text = Format(ValorUnit, "#,##0." & String(Dec_Costo, "0"))
End Sub

Public Sub TotalInventario()
Dim TotalInvs As Currency
  Total = 0: Total_IVA = 0: NoExiste = True
  With AdoKardex.Recordset
    If .RecordCount > 0 Then
       .MoveFirst
        Do While Not .EOF
           If .Fields("DH") = "2" Then Total = Total + .Fields("VALOR_TOTAL")
           If .Fields("CODIGO_INV") = CodigoInv Then NoExiste = False
          .MoveNext
        Loop
       .MoveFirst
    End If
  End With
  Total = Round(Total, 4)
  Label1.Caption = Format(Total, "#,##0.00")
End Sub

Public Sub ListarProveedorUsuario(Proveedor As Boolean)
  sSQL = "SELECT C.Cliente,CR.Codigo,C.CI_RUC,C.TD " _
       & "FROM Clientes As C,Catalogo_Rol_Pagos As CR " _
       & "WHERE CR.Item = '" & NumEmpresa & "' " _
       & "AND CR.Periodo = '" & Periodo_Contable & "' " _
       & "AND C.Codigo = CR.Codigo " _
       & "ORDER BY C.Cliente "
  SelectDB_Combo DCBenef, AdoBenef, sSQL, "Cliente"
End Sub

Private Sub TxtCodBar_GotFocus()
  MarcarTexto TxtCodBar
End Sub

Private Sub TxtCodBar_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtCodBar_LostFocus()
  TextoValido TxtCodBar, , True
End Sub

Public Sub Select_Adodc_Grid_Kardex()
  sSQL = "SELECT * " _
       & "FROM Asiento_K " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND T_No = " & Trans_No & " " _
       & "ORDER BY A_No "
  SQLDec = "VALOR_UNIT " & CStr(Dec_Costo) & "|,VALOR_TOTAL 4|,CANTIDAD 2|,SALDO 4|."
  Select_Adodc_Grid DGKardex, AdoKardex, sSQL, SQLDec
End Sub

Private Sub TxtLoteNo_GotFocus()
  MarcarTexto TxtLoteNo
End Sub

Private Sub TxtModelo_GotFocus()
     MarcarTexto TxtModelo
End Sub

Private Sub TxtProcedencia_GotFocus()
  MarcarTexto TxtProcedencia
End Sub

Private Sub TxtRegSanitario_GotFocus()
   MarcarTexto TxtRegSanitario
End Sub

Private Sub TxtRegSanitario_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub TxtSerieNo_GotFocus()
    MarcarTexto TxtSerieNo
End Sub
