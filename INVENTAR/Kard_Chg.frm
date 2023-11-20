VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Begin VB.Form Kard_Cambio 
   Caption         =   "CONTROL DE INVENTARIO PARA INGRESOS/EGRESOS"
   ClientHeight    =   7335
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11460
   Icon            =   "Kard_Chg.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7335
   ScaleWidth      =   11460
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Code39Clt1 
      Height          =   480
      Left            =   2415
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   34
      Top             =   6615
      Width           =   1200
   End
   Begin VB.TextBox TextTotal 
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
      Left            =   9660
      MultiLine       =   -1  'True
      TabIndex        =   24
      Text            =   "Kard_Chg.frx":0442
      Top             =   3225
      Width           =   1695
   End
   Begin VB.OptionButton OpcEntra 
      Caption         =   "Entradas"
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
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   840
      Width           =   855
   End
   Begin VB.OptionButton OpcSale 
      Caption         =   "Salidas"
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
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   840
      Value           =   -1  'True
      Width           =   855
   End
   Begin MSDataListLib.DataCombo DCTInv 
      Bindings        =   "Kard_Chg.frx":0444
      DataSource      =   "AdoTInv"
      Height          =   315
      Left            =   105
      TabIndex        =   9
      Top             =   1470
      Width           =   5580
      _ExtentX        =   9843
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
      Left            =   5775
      MaxLength       =   25
      TabIndex        =   22
      Text            =   "."
      Top             =   3255
      Width           =   3795
   End
   Begin MSDataListLib.DataCombo DCBodega 
      Bindings        =   "Kard_Chg.frx":045A
      DataSource      =   "AdoBodega"
      Height          =   315
      Left            =   105
      TabIndex        =   8
      Top             =   1155
      Width           =   5580
      _ExtentX        =   9843
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
      Bindings        =   "Kard_Chg.frx":0472
      Height          =   2745
      Left            =   105
      TabIndex        =   25
      Top             =   3675
      Width           =   11250
      _ExtentX        =   19844
      _ExtentY        =   4842
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
      Bindings        =   "Kard_Chg.frx":048A
      DataSource      =   "AdoInv"
      Height          =   1155
      Left            =   105
      TabIndex        =   10
      Top             =   1785
      Width           =   5580
      _ExtentX        =   9843
      _ExtentY        =   2037
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
      Left            =   3990
      MultiLine       =   -1  'True
      TabIndex        =   20
      Text            =   "Kard_Chg.frx":049F
      Top             =   3255
      Width           =   1695
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
      Left            =   2625
      MultiLine       =   -1  'True
      TabIndex        =   18
      Text            =   "Kard_Chg.frx":04A1
      Top             =   3255
      Width           =   1275
   End
   Begin MSDataListLib.DataCombo DCBenef 
      Bindings        =   "Kard_Chg.frx":04A3
      DataSource      =   "AdoBenef"
      Height          =   315
      Left            =   3150
      TabIndex        =   3
      Top             =   105
      Width           =   6000
      _ExtentX        =   10583
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
      Top             =   4620
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
      Top             =   4935
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
      Top             =   5250
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
      Top             =   5250
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
      Top             =   5565
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
      Picture         =   "Kard_Chg.frx":04BA
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   6510
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
      Width           =   8205
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
      Picture         =   "Kard_Chg.frx":0D84
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   6510
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
      Top             =   4935
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
      Top             =   4620
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
   Begin MSDataListLib.DataCombo DCTInv1 
      Bindings        =   "Kard_Chg.frx":11C6
      DataSource      =   "AdoTInvS"
      Height          =   315
      Left            =   5775
      TabIndex        =   13
      Top             =   1470
      Width           =   5580
      _ExtentX        =   9843
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
   Begin MSDataListLib.DataCombo DCInv1 
      Bindings        =   "Kard_Chg.frx":11DD
      DataSource      =   "AdoInvS"
      Height          =   1155
      Left            =   5775
      TabIndex        =   14
      Top             =   1785
      Width           =   5580
      _ExtentX        =   9843
      _ExtentY        =   2037
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
   Begin MSAdodcLib.Adodc AdoInvS 
      Height          =   330
      Left            =   315
      Top             =   4935
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
      Top             =   5565
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
   Begin MSDataListLib.DataCombo DCBodega1 
      Bindings        =   "Kard_Chg.frx":11F3
      DataSource      =   "AdoBodega"
      Height          =   315
      Left            =   5775
      TabIndex        =   12
      Top             =   1155
      Width           =   5580
      _ExtentX        =   9843
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
      Left            =   9135
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
      Left            =   105
      TabIndex        =   16
      Top             =   3255
      Width           =   2430
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
      Left            =   7875
      TabIndex        =   32
      Top             =   6825
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
      Left            =   7875
      TabIndex        =   33
      Top             =   6510
      Width           =   1695
   End
   Begin VB.Label Label12 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " EL PRODUCTO ENTRA A:"
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
      TabIndex        =   11
      Top             =   840
      Width           =   4740
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " EL PRODUCTO SALE DE:"
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
      Width           =   4740
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
      Left            =   5775
      TabIndex        =   21
      Top             =   2940
      Width           =   3795
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
      Left            =   105
      TabIndex        =   15
      Top             =   2940
      Width           =   2430
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
      Left            =   9660
      TabIndex        =   27
      Top             =   6825
      Width           =   1695
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " VALOR TOTAL"
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
      Left            =   9660
      TabIndex        =   23
      Top             =   2940
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
      Left            =   3990
      TabIndex        =   19
      Top             =   2940
      Width           =   1695
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
      Left            =   2625
      TabIndex        =   17
      Top             =   2940
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
      Left            =   9660
      TabIndex        =   26
      Top             =   6510
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
      Caption         =   " PROVEEDOR"
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
Attribute VB_Name = "Kard_Cambio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub ListarProductosES(Optional TipoInvSalida As Boolean)
 If TipoInvSalida Then
    CodigoInv = SinEspaciosIzq(DCTInv.Text)
 Else
    CodigoInv1 = SinEspaciosIzq(DCTInv1.Text)
 End If
 sSQL = "SELECT * " _
      & "FROM Catalogo_Productos " _
      & "WHERE Item = '" & NumEmpresa & "' " _
      & "AND Periodo = '" & Periodo_Contable & "' "
 If TipoInvSalida Then
    sSQL = sSQL & "AND MidStrg(Codigo_Inv,1," & CStr(Len(CodigoInv)) & ") = '" & CodigoInv & "' "
 Else
    sSQL = sSQL & "AND MidStrg(Codigo_Inv,1," & CStr(Len(CodigoInv1)) & ") = '" & CodigoInv1 & "' "
 End If
 sSQL = sSQL & "AND LEN(Cta_Inventario) > 1 " _
      & "AND TC = 'P' " _
      & "ORDER BY Producto "
 'MsgBox sSQL
 If TipoInvSalida Then
    SelectDB_Combo DCInv, AdoInv, sSQL, "Producto"
 Else
    SelectDB_Combo DCInv1, AdoInvS, sSQL, "Producto"
 End If
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
          sSQL = "SELECT * " _
               & "FROM Asiento_K " _
               & "WHERE Item = '" & NumEmpresa & "' " _
               & "AND CodigoU = '" & CodigoUsuario & "' " _
               & "AND T_No = " & Trans_No & " "
          Select_Adodc_Grid DGKardex, AdoKardex, sSQL, 4
          Unload Kard_Cambio
   End If
  End With
  End If
  RatonNormal
End Sub

Private Sub Command2_Click()
  Unload Kard_Cambio
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
          InvImp = .Fields("Importaciones")
          If InvImp Then Label2.Caption = " IMPORTACION"
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
      .Find ("Bodega Like '" & DCBodega1.Text & "' ")
       If Not .EOF Then Cod_Bodega1 = .Fields("CodBod")
   End If
  End With
End Sub

Private Sub DCInv_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyEscape: Command1.SetFocus
    Case vbKeyReturn: SiguienteControl
  End Select
End Sub

Private Sub DCInv_LostFocus()
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
          Producto = .Fields("Producto")
          Cta_Inventario = .Fields("Cta_Inventario")
          If OpcSale.value Then TextEntrada.SetFocus Else DCBodega1.SetFocus
       Else
         .MoveFirst
         .Find ("Codigo_Barra Like '" & CodigoInv & "' ")
          If Not .EOF Then
             Si_No = .Fields("IVA")
             Unidad = .Fields("Unidad")
             CodigoInv = .Fields("Codigo_Inv")
             Producto = .Fields("Producto")
             Cta_Inventario = .Fields("Cta_Inventario")
             If OpcSale.value Then TextEntrada.SetFocus Else DCBodega1.SetFocus
          Else
            .MoveFirst
            .Find ("Codigo_Inv Like '" & CodigoInv & "' ")
             If Not .EOF Then
                Si_No = .Fields("IVA")
                Unidad = .Fields("Unidad")
                CodigoInv = .Fields("Codigo_Inv")
                Producto = .Fields("Producto")
                Cta_Inventario = .Fields("Cta_Inventario")
                If OpcSale.value Then TextEntrada.SetFocus Else DCBodega1.SetFocus
             Else
                MsgBox "No existe Productos asignados"
                DCInv.SetFocus
             End If
          End If
       End If
    End If
  End With
End Sub

Private Sub DCInv1_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCInv1_LostFocus()
  CodigoInv1 = DCInv1.Text
  Si_No = False
  'MsgBox CodigoInv1
  With AdoInvS.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Producto Like '" & CodigoInv1 & "' ")
       If Not .EOF Then
          CodigoInv = .Fields("Codigo_Inv")
          Producto = .Fields("Producto")
          Contra_Cta = .Fields("Cta_Inventario")
          TextEntrada.SetFocus
       Else
         .MoveFirst
         .Find ("Codigo_Barra Like '" & CodigoInv1 & "' ")
          If Not .EOF Then
             CodigoInv = .Fields("Codigo_Inv")
             Producto = .Fields("Producto")
             Contra_Cta = .Fields("Cta_Inventario")
             TextEntrada.SetFocus
          Else
            .MoveFirst
            .Find ("Codigo_Inv Like '" & CodigoInv1 & "' ")
             If Not .EOF Then
                CodigoInv = .Fields("Codigo_Inv")
                Producto = .Fields("Producto")
                Contra_Cta = .Fields("Cta_Inventario")
                TextEntrada.SetFocus
             Else
                MsgBox "No existe Productos asignados"
                DCTInv1.SetFocus
             End If
          End If
       End If
    End If
  End With
End Sub

Private Sub DCTInv_KeyDown(KeyCode As Integer, Shift As Integer)
 PresionoEnter KeyCode
End Sub

Private Sub DCTInv_LostFocus()
 ListarProductosES True
End Sub

Private Sub DCTInv1_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCTInv1_LostFocus()
  ListarProductosES
End Sub

Private Sub DGKardex_BeforeDelete(Cancel As Integer)
  Cancel = DeleteSiNo(AdoKardex)
End Sub

Private Sub MBFechaI_GotFocus()
  MarcarTexto MBFechaI
End Sub

Private Sub MBFechaI_LostFocus()
  FechaValida MBFechaI
End Sub

Private Sub TextConcepto_GotFocus()
  MarcarTexto TextConcepto
End Sub

Private Sub TextConcepto_LostFocus()
  TextoValido TextConcepto
End Sub

Private Sub TextEntrada_GotFocus()
  SaldoAnterior = 0: ValorUnitAnt = 0: Contador = 0: ValorUnit = 0: Cantidad = 0
  Stock_Actual_Inventario MBFechaI, CodigoInv
  MarcarTexto TextEntrada
  Cantidad = Val(CCur(TextEntrada))
End Sub

Private Sub TextEntrada_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextEntrada_LostFocus()
  TextoValido TextEntrada, True
  Entrada = Val(CCur(TextEntrada))
  Cantidad = Entrada
  LabelCodigo.Caption = CodigoInv
End Sub

Private Sub Form_Activate()
' Consultamos las cuentas de la tabla
 Total_IVA = 0
 Label3.Caption = " RESPONSABLE:"
 ListarProveedorUsuario False
 DCBenef.SetFocus
 Trans_No = 81
 IniciarAsientosAdo AdoAsientos
 If Inv_Promedio Then
    Kard_Cambio.Caption = "CONTROL DE INVENTARIO EN PROMEDIO"
 Else
    Kard_Cambio.Caption = "CONTROL DE INVENTARIO EN ULTIMO PRECIO"
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
'Entrada:
 sSQL = "SELECT Codigo_Inv & '  ' & Producto As NomProd " _
      & "FROM Catalogo_Productos " _
      & "WHERE Item = '" & NumEmpresa & "' " _
      & "AND Periodo = '" & Periodo_Contable & "' " _
      & "AND TC = 'I' " _
      & "ORDER BY Codigo_Inv "
 SelectDB_Combo DCTInv1, AdoTInvS, sSQL, "NomProd"
 
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

Private Sub TextTotal_GotFocus()
Dim DValorUnit As Double
Dim DValorTotal As Double
   Entrada = Val(CCur(TextEntrada.Text))
   ValorUnit = Val(CDbl(TextVUnit.Text))
   DValorUnit = ValorUnit
   Total_Desc = 0
   DValorUnit = DValorUnit - (DValorUnit * Total_Desc)
   If Total_Desc > 0 Then DValorUnit = DValorUnit - (DValorUnit * Total_Desc)
   DValorTotal = DValorUnit * Entrada
   ValorUnit = CCur(DValorUnit)
   ValorTotal = CCur(DValorTotal)
  'MsgBox ValorTotal
   TextTotal.Text = Format(ValorTotal, "#,##0.0000")
End Sub

Private Sub TextTotal_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextTotal_LostFocus()
   Entrada = Val(CCur(TextEntrada))
  'ValorTotal = ValorUnit * Entrada
   Codigo = Leer_Cta_Catalogo(Contra_Cta)
 ' llenamos el ultimo saldo del kardex
   TextVUnit = Format(ValorUnit, "#,##0." & String(Dec_Costo, "0"))
   If Entrada > 0 And ValorUnit > 0 Then
      SetAddNew AdoKardex
      If OpcSale.value Then
         OpcDH = "2"
         Cantidad = Cantidad - Entrada
         Saldo = SaldoAnterior - ValorTotal
         SetFields AdoKardex, "CodBod", Cod_Bodega
         SetFields AdoKardex, "CTA_INVENTARIO", Cta_Inventario
      Else
         OpcDH = "1"
         Cantidad = Cantidad + Entrada
         Saldo = SaldoAnterior + ValorTotal
         SetFields AdoKardex, "CodBod", Cod_Bodega1
         SetFields AdoKardex, "CTA_INVENTARIO", Contra_Cta
      End If
      
      SetFields AdoKardex, "CODIGO_INV", CodigoInv
      
      SetFields AdoKardex, "DH", OpcDH
      SetFields AdoKardex, "PRODUCTO", Producto
      SetFields AdoKardex, "CANT_ES", Entrada
      SetFields AdoKardex, "VALOR_UNIT", Round(ValorUnit, Dec_Costo)
      SetFields AdoKardex, "VALOR_TOTAL", Round(ValorTotal, Dec_Costo)
      SetFields AdoKardex, "CONTRA_CTA", Ninguno
      SetFields AdoKardex, "CANTIDAD", Cantidad
      SetFields AdoKardex, "VALOR_UNIT", ValorUnit
      SetFields AdoKardex, "SALDO", Saldo
      SetFields AdoKardex, "UNIDAD", Unidad
      SetFields AdoKardex, "Item", NumEmpresa
      SetFields AdoKardex, "CodigoU", CodigoUsuario
      SetFields AdoKardex, "T_No", Trans_No
      SetFields AdoKardex, "SUBCTA", SubCtaGen
      SetFields AdoKardex, "TC", SubCta
      SetFields AdoKardex, "COD_BAR", TxtCodBar
      SetFields AdoKardex, "Codigo_B", CodigoCliente
      SetUpdate AdoKardex
      TextTotal.Text = Format(ValorTotal, "#,##0.00")
      TotalInventario
      If OpcSale.value Then DCInv.SetFocus Else DCInv1.SetFocus
   End If
End Sub

Private Sub TextVUnit_GotFocus()
  If OpcEntra.value Then
     If Entrada <= 0 Then Entrada = 1
     Precio = Total / Entrada
  Else
     Precio = ValorUnit
  End If
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
  TextTotal.Text = Format(ValorTotal, "#,##0.0000")
  TextVUnit.Text = Format(ValorUnit, "#,##0." & String(Dec_Costo, "0"))
End Sub

Public Sub TotalInventario()
Dim TotalInvs As Currency
  Total = 0: Total_IVA = 0
  With AdoKardex.Recordset
    If .RecordCount > 0 Then
       .MoveFirst
        Do While Not .EOF
           If .Fields("DH") = "2" Then Total = Total + .Fields("VALOR_TOTAL")
          .MoveNext
        Loop
       .MoveFirst
    End If
  End With
  Total = Round(Total, 4)
  Label1.Caption = Format(Total, "#,##0.00")
End Sub

Public Sub ListarProveedorUsuario(Proveedor As Boolean)
  sSQL = "SELECT C.Cliente,CR.Codigo,C.CI_RUC,C.Direccion,C.Telefono,C.TD,'.' As Cta,0 As Importaciones " _
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

