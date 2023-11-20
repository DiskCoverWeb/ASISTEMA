VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form RSubCtas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Modulos de Gastos"
   ClientHeight    =   8355
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14100
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8355
   ScaleWidth      =   14100
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   11655
      Top             =   1050
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Rsubctas.frx":0000
            Key             =   "Salir"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Rsubctas.frx":08DA
            Key             =   "Printer"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Rsubctas.frx":11B4
            Key             =   "Uno"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Rsubctas.frx":1A8E
            Key             =   "Dos"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   14100
      _ExtentX        =   24871
      _ExtentY        =   1164
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Salir"
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Imprimir"
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Periodo"
            Object.ToolTipText     =   "Periodo"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Costos"
            Object.ToolTipText     =   "Costos de Pryectos"
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DGSubCtaDet 
      Bindings        =   "Rsubctas.frx":1EE0
      Height          =   2850
      Left            =   105
      TabIndex        =   18
      Top             =   5355
      Width           =   13875
      _ExtentX        =   24474
      _ExtentY        =   5027
      _Version        =   393216
      BorderStyle     =   0
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
   Begin MSDataGridLib.DataGrid DGCostosTours 
      Bindings        =   "Rsubctas.frx":1EFB
      Height          =   1800
      Left            =   105
      TabIndex        =   17
      Top             =   3570
      Width           =   13875
      _ExtentX        =   24474
      _ExtentY        =   3175
      _Version        =   393216
      BorderStyle     =   0
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
   Begin VB.TextBox TextGuia 
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
      MaxLength       =   25
      TabIndex        =   8
      Top             =   3150
      Width           =   3375
   End
   Begin VB.TextBox TextPAX 
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
      MaxLength       =   5
      TabIndex        =   10
      Text            =   "0"
      Top             =   3150
      Width           =   960
   End
   Begin MSDataListLib.DataCombo DCSubCta 
      Bindings        =   "Rsubctas.frx":1F17
      DataSource      =   "AdoSubCta"
      Height          =   345
      Left            =   7035
      TabIndex        =   16
      Top             =   3150
      Visible         =   0   'False
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   609
      _Version        =   393216
      Text            =   "SubCta"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataList DLCtas1 
      Bindings        =   "Rsubctas.frx":1F2F
      DataSource      =   "AdoCtas1"
      Height          =   1635
      Left            =   5670
      TabIndex        =   6
      Top             =   1155
      Width           =   5580
      _ExtentX        =   9843
      _ExtentY        =   2884
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataList DLCtas 
      Bindings        =   "Rsubctas.frx":1F46
      DataSource      =   "AdoCtas"
      Height          =   1635
      Left            =   105
      TabIndex        =   5
      Top             =   1155
      Width           =   5580
      _ExtentX        =   9843
      _ExtentY        =   2884
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSAdodcLib.Adodc AdoAux 
      Height          =   330
      Left            =   105
      Top             =   5040
      Visible         =   0   'False
      Width           =   2325
      _ExtentX        =   4101
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
   Begin VB.CheckBox CheqResumen 
      Caption         =   "Solo Resumen"
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
      Left            =   5565
      TabIndex        =   4
      Top             =   735
      Width           =   2115
   End
   Begin VB.CheckBox CheqSC 
      Caption         =   "Subcuenta"
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
      Left            =   7035
      TabIndex        =   15
      Top             =   2835
      Width           =   1275
   End
   Begin MSMask.MaskEdBox MBFechaF 
      Height          =   330
      Left            =   4200
      TabIndex        =   3
      Top             =   735
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
   Begin MSMask.MaskEdBox MBFechaI 
      Height          =   330
      Left            =   1995
      TabIndex        =   1
      Top             =   735
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
   Begin MSAdodcLib.Adodc AdoCtas1 
      Height          =   330
      Left            =   5460
      Top             =   1260
      Visible         =   0   'False
      Width           =   1800
      _ExtentX        =   3175
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
      Caption         =   "Ctas1"
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
   Begin MSAdodcLib.Adodc AdoSubCtaDet1 
      Height          =   330
      Left            =   105
      Top             =   5355
      Visible         =   0   'False
      Width           =   2325
      _ExtentX        =   4101
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
      Caption         =   "SubCtaDet1"
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
   Begin MSAdodcLib.Adodc AdoCostosTours 
      Height          =   330
      Left            =   105
      Top             =   5670
      Visible         =   0   'False
      Width           =   2325
      _ExtentX        =   4101
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
      Caption         =   "CostosTours"
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
   Begin MSAdodcLib.Adodc AdoSubCtaDet 
      Height          =   330
      Left            =   105
      Top             =   5985
      Visible         =   0   'False
      Width           =   2325
      _ExtentX        =   4101
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
      Caption         =   "SubCtaDet"
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
      Left            =   210
      Top             =   1260
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
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
   Begin MSAdodcLib.Adodc AdoSubCta 
      Height          =   330
      Left            =   7875
      Top             =   735
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
   Begin MSMask.MaskEdBox MBFechaAl 
      Height          =   330
      Left            =   5670
      TabIndex        =   14
      Top             =   3150
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
   Begin MSMask.MaskEdBox MBFechaDe 
      Height          =   330
      Left            =   4410
      TabIndex        =   12
      Top             =   3150
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
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "GUIA"
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
      Top             =   2835
      Width           =   3375
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "No. PAX"
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
      TabIndex        =   9
      Top             =   2835
      Width           =   960
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Desde:"
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
      Left            =   4410
      TabIndex        =   11
      Top             =   2835
      Width           =   1275
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Hasta:"
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
      Left            =   5670
      TabIndex        =   13
      Top             =   2835
      Width           =   1275
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Desde (DD/MM/AA)"
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
      Top             =   735
      Width           =   1905
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Hasta"
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
      Left            =   3360
      TabIndex        =   2
      Top             =   735
      Width           =   855
   End
End
Attribute VB_Name = "RSubCtas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim FilaActual As Integer
Dim SumaSubCta As Currency

Public Sub Consultar()
  TextoValido TextGuia
  TextGuia.Text = UCaseStrg(TextGuia.Text)
  FechaValida MBFechaI
  FechaValida MBFechaF
  FechaValida MBFechaDe
  FechaValida MBFechaAl
  FechaIni = BuscarFecha(MBFechaI)
  FechaFin = BuscarFecha(MBFechaF)
  Codigo = Ninguno
  Codigo2 = Ninguno
  Codigo1 = Ninguno
  If AdoCtas.Recordset.RecordCount > 0 Then Codigo = SinEspaciosIzq(DLCtas)         'Gastos
  If AdoCtas1.Recordset.RecordCount > 0 Then Codigo2 = SinEspaciosIzq(DLCtas1)      'Ingresos
  If AdoSubCta.Recordset.RecordCount > 0 Then Codigo1 = SinEspaciosDer(DCSubCta)    'SunModulos
  
  sSQL = "DELETE * " _
       & "FROM Saldo_Diarios " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND TP = 'COST' "
  Ejecutar_SQL_SP sSQL
  
  sSQL = "SELECT * " _
       & "FROM Saldo_Diarios " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND TP = 'COST' "
  Select_Adodc AdoAux, sSQL
  
''' 'Insertamos Presupuestos de Gastos
'''  RatonReloj
'''  With AdoCtas.Recordset
'''   If .RecordCount > 0 Then
'''       Do While Not .EOF
'''          Cta_Aux = SinEspaciosIzq(.Fields("Nombre_Cta"))
'''          If AdoSubCta.Recordset.RecordCount > 0 Then
'''             AdoSubCta.Recordset.MoveFirst
'''             Do While Not AdoSubCta.Recordset.EOF
'''                Cod_Benef = SinEspaciosDer(AdoSubCta.Recordset.Fields("Nombre_SubCta"))
'''                sSQL = "SELECT Cta,SUM(Presupuesto) As TPresupuesto " _
'''                     & "FROM Trans_Presupuestos " _
'''                     & "WHERE Item = '" & NumEmpresa & "' " _
'''                     & "AND Periodo = '" & Periodo_Contable & "' " _
'''                     & "AND Cta = '" & Cta_Aux & "' " _
'''                     & "AND Codigo = '" & Cod_Benef & "' " _
'''                     & "GROUP BY Cta "
'''                Select_Adodc AdoAux, sSQL
'''                If AdoAux.Recordset.RecordCount <= 0 Then
'''                   For IE = 1 To 12
'''                       Mifecha = UltimoDiaMes("01/" & Format(IE, "00") & "/" & Year(MBFechaI))
'''                       SetAdoAddNew "Trans_Presupuestos"
'''                       SetAdoFields "Cta", Cta_Aux
'''                       SetAdoFields "Codigo", Cod_Benef
'''                       SetAdoFields "Mes_No", Mifecha
'''                       SetAdoFields "Mes", UCaseStrg(MidStrg(MesesLetras(IE), 1, 3))
'''                       SetAdoFields "Item", NumEmpresa
'''                       SetAdoFields "Periodo", Periodo_Contable
'''                       SetAdoUpdate
'''                   Next IE
'''                End If
'''                AdoSubCta.Recordset.MoveNext
'''             Loop
'''          End If
'''         .MoveNext
'''       Loop
'''   End If
'''  End With
'''
''' 'Insertamos Presupuestos de Ingresos
'''  RatonReloj
'''  With AdoCtas1.Recordset
'''   If .RecordCount > 0 Then
'''       Do While Not .EOF
'''          Cta_Aux = SinEspaciosIzq(.Fields("Nombre_Cta"))
'''          sSQL = "SELECT Cta,SUM(Presupuesto) As TPresupuesto " _
'''               & "FROM Trans_Presupuestos " _
'''               & "WHERE Item = '" & NumEmpresa & "' " _
'''               & "AND Periodo = '" & Periodo_Contable & "' " _
'''               & "AND Cta = '" & Cta_Aux & "' " _
'''               & "GROUP BY Cta "
'''          Select_Adodc AdoAux, sSQL
'''          If AdoAux.Recordset.RecordCount <= 0 Then
'''             For IE = 1 To 12
'''                 Mifecha = UltimoDiaMes("01/" & Format(IE, "00") & "/" & Year(MBFechaI))
'''                 SetAdoAddNew "Trans_Presupuestos"
'''                 SetAdoFields "Cta", Cta_Aux
'''                 SetAdoFields "Mes_No", Mifecha
'''                 SetAdoFields "Mes", UCaseStrg(MidStrg(MesesLetras(IE), 1, 3))
'''                 SetAdoFields "Item", NumEmpresa
'''                 SetAdoFields "Periodo", Periodo_Contable
'''                 SetAdoUpdate
'''             Next IE
'''          End If
'''         .MoveNext
'''       Loop
'''   End If
'''  End With
'''  sSQL = "SELECT * FROM Costos " _
'''       & "WHERE Item = '" & NumEmpresa & "' " _
'''       & "AND CodigoU = '" & CodigoUsuario & "' "
'''  Select_Adodc AdoAux, sSQL
'''
'''  sSQL = "DELETE * FROM Costos " _
'''       & "WHERE Item = '" & NumEmpresa & "' " _
'''       & "AND CodigoU = '" & CodigoUsuario & "' "
'''  Ejecutar_SQL_SP sSQL

'''If Cheq_MN_ME.Value Then
'''     sSQL = sSQL & "SUM(SC.Parcial_ME) As Valor_ME, " _
'''          & "(SUM(SC.Parcial_ME) - P.Presupuesto) As Diferencia "
'''  Else
'''   End If
''
''  sSQL = "SELECT Be.Detalle,P.Presupuesto,SUM(SC.Debitos) As Suma_Debitos,SUM(SC.Creditos) As Suma_Creditos," _
''       & "((SUM(SC.Debitos) + SUM(SC.Creditos)) - SUM(P.Presupuesto)) As Diferencia " _
''       & "FROM Catalogo_SubCtas As Be,Trans_SubCtas As SC,Trans_Presupuestos As P " _
''       & "WHERE SC.Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
''       & "AND SC.Item = '" & NumEmpresa & "' " _
''       & "AND SC.Periodo = '" & Periodo_Contable & "' " _
''       & "AND SC.Cta = '" & Codigo & "' " _
''       & "AND SC.Cta = P.Cta " _
''       & "AND Be.Periodo = SC.Periodo " _
''       & "AND Be.Periodo = P.Periodo " _
''       & "AND Be.Item = SC.Item " _
''       & "AND Be.Item = P.Item " _
''       & "AND Be.Codigo = SC.Codigo " _
''       & "AND Be.Codigo = P.Codigo "
''  If CheqSC.value = 1 Then sSQL = sSQL & "AND SC.Codigo = '" & Codigo1 & "' "
''  sSQL = sSQL & "GROUP BY Be.Detalle,P.Presupuesto "
''  Select_Adodc_Grid DGCostosTours, AdoSubCtaDet1, sSQL
  
'''  sSQL = "SELECT * FROM Costos_R " _
'''       & "WHERE Item = '" & NumEmpresa & "' " _
'''       & "AND CodigoU = '" & CodigoUsuario & "' "
'''  Select_Adodc AdoAux, sSQL
'''
  sSQL = "SELECT Co.Codigo_B,SC.Cta,SC.Fecha,CSC.Detalle,SC.Codigo,SC.TP,SC.Numero,SC.Debitos,SC.Creditos " _
       & "FROM Comprobantes As Co,Trans_SubCtas As SC,Catalogo_SubCtas As CSC " _
       & "WHERE SC.Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
       & "AND SC.Item = '" & NumEmpresa & "' " _
       & "AND SC.Periodo = '" & Periodo_Contable & "' " _
       & "AND SC.Cta = '" & Codigo & "' " _
       & "AND Co.TP = SC.TP " _
       & "AND Co.Numero = SC.Numero " _
       & "AND Co.Item = SC.Item " _
       & "AND SC.Item = CSC.Item " _
       & "AND Co.Periodo = SC.Periodo " _
       & "AND SC.Periodo = CSC.Periodo " _
       & "AND CSC.Codigo = SC.Codigo "
  If CheqSC.value = 1 Then sSQL = sSQL & "AND SC.Codigo = '" & Codigo1 & "' "
  sSQL = sSQL _
       & "ORDER BY SC.Fecha,Co.Codigo_B,SC.TP,SC.Numero,SC.Debitos,SC.Creditos "
  Select_Adodc AdoSubCtaDet, sSQL
  With AdoSubCtaDet.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          SetAdoAddNew "Saldo_Diarios"
          SetAdoFields "Fecha", .Fields("Fecha")
          SetAdoFields "CodigoC", .Fields("Codigo_B")
          SetAdoFields "Cta", .Fields("Cta")
          SetAdoFields "Comprobante", .Fields("Detalle")
          SetAdoFields "TC", .Fields("TP")
          SetAdoFields "Numero", .Fields("Numero")
          SetAdoFields "Ingresos", .Fields("Debitos")
          SetAdoFields "Egresos", .Fields("Creditos")
          SetAdoFields "Item", NumEmpresa
          SetAdoFields "CodigoU", CodigoUsuario
          SetAdoFields "TP", "COST"
          'If Cheq_MN_ME.Value Then SetFields AdoAux, "Valor_ME", .Fields("Valor_ME")
          SetAdoUpdate
         .MoveNext
       Loop
   End If
  End With
  sSQL = "SELECT Co.Codigo_B,SC.Cta,SC.Fecha,CSC.Detalle,SC.Codigo,SC.TP,SC.Numero,SC.Debitos,SC.Creditos " _
       & "FROM Comprobantes As Co,Trans_SubCtas As SC,Catalogo_SubCtas As CSC " _
       & "WHERE SC.Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
       & "AND SC.Item = '" & NumEmpresa & "' " _
       & "AND SC.Periodo = '" & Periodo_Contable & "' " _
       & "AND SC.Cta = '" & Codigo2 & "' " _
       & "AND Co.TP = SC.TP " _
       & "AND Co.Numero = SC.Numero " _
       & "AND Co.Item = SC.Item " _
       & "AND SC.Item = CSC.Item " _
       & "AND Co.Periodo = SC.Periodo " _
       & "AND SC.Periodo = CSC.Periodo " _
       & "AND CSC.Codigo = SC.Codigo "
  If CheqSC.value = 1 Then sSQL = sSQL & "AND SC.Codigo = '" & Codigo1 & "' "
  sSQL = sSQL _
       & "ORDER BY SC.Fecha,Co.Codigo_B,SC.TP,SC.Numero,SC.Debitos,SC.Creditos "
  Select_Adodc AdoSubCtaDet, sSQL
  With AdoSubCtaDet.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          SetAdoAddNew "Saldo_Diarios"
          SetAdoFields "Fecha", .Fields("Fecha")
          SetAdoFields "CodigoC", .Fields("Codigo_B")
          SetAdoFields "Cta", .Fields("Cta")
          SetAdoFields "Comprobante", .Fields("Detalle")
          SetAdoFields "TC", .Fields("TP")
          SetAdoFields "Numero", .Fields("Numero")
          SetAdoFields "Ingresos", .Fields("Debitos")
          SetAdoFields "Egresos", .Fields("Creditos")
          SetAdoFields "Item", NumEmpresa
          SetAdoFields "CodigoU", CodigoUsuario
          SetAdoFields "TP", "COST"
          'If Cheq_MN_ME.Value Then SetFields AdoAux, "Valor_ME", .Fields("Valor_ME")
          SetAdoUpdate
         .MoveNext
       Loop
   End If
  End With
  Debe = 0: Haber = 0
'''  sSQL = "SELECT Beneficiario,Presupuesto,SumaDebitos,SumaCreditos,"
'''  If Cheq_MN_ME.Value Then
'''     sSQL = sSQL & "Valor_ME,Diferencia "
'''  Else
'''     sSQL = sSQL & "Diferencia "
'''  End If
'''  sSQL = sSQL & "FROM Costos " _
'''       & "WHERE Item = '" & NumEmpresa & "' " _
'''       & "AND CodigoU = '" & CodigoUsuario & "' " _
'''       & "ORDER BY Beneficiario,Presupuesto "
'''  Select_Adodc AdoSubCtaDet1, sSQL, False
'''

  sSQL = "SELECT C.Cliente As Beneficiario,SD.Comprobante As Cuenta,SUM(SD.Ingresos) As Suma_Debitos,SUM(SD.Egresos) As Suma_Creditos " _
       & "FROM Saldo_Diarios As SD,Clientes As C " _
       & "WHERE SD.Item = '" & NumEmpresa & "' " _
       & "AND SD.CodigoU = '" & CodigoUsuario & "' " _
       & "AND SD.TP = 'COST' " _
       & "AND SD.CodigoC = C.Codigo " _
       & "GROUP BY C.Cliente,SD.Comprobante " _
       & "ORDER BY C.Cliente,SD.Comprobante "
  Select_Adodc_Grid DGCostosTours, AdoSubCtaDet1, sSQL

  sSQL = "SELECT SD.Fecha,C.Cliente As Beneficiario,SD.Comprobante As Cuenta,SD.TC As Tipo,SD.Numero,SD.Ingresos As Debitos,SD.Egresos As Creditos " _
       & "FROM Saldo_Diarios As SD,Clientes As C " _
       & "WHERE SD.Item = '" & NumEmpresa & "' " _
       & "AND SD.CodigoU = '" & CodigoUsuario & "' " _
       & "AND SD.TP = 'COST' " _
       & "AND SD.CodigoC = C.Codigo " _
       & "ORDER BY SD.Fecha,C.Cliente,SD.TC,SD.Numero,SD.Ingresos,SD.Egresos "
  Select_Adodc_Grid DGSubCtaDet, AdoSubCtaDet, sSQL
  With AdoSubCtaDet.Recordset
    If .RecordCount > 0 Then
        Do While Not .EOF
           Debe = Debe + .Fields("Debitos")
           Haber = Haber + .Fields("Creditos")
          .MoveNext
        Loop
       .MoveFirst
    End If
  End With
'''  Total = 0: Saldo = 0
'''  If Dolar <> 0 And Val(TextPAX.Text) <> 0 Then
'''     Total = (Haber - Debe) / Dolar
'''     Saldo = Total / Val(TextPAX.Text)
'''  End If
 'DBLCtas.SetFocus
 RatonNormal
End Sub

Private Sub CheqSC_Click()
  If CheqSC.value = 1 Then DCSubCta.Visible = True Else DCSubCta.Visible = False
End Sub

Private Sub Command1_Click()

End Sub

Public Sub Imprimir_Resultados()
  If AdoSubCtaDet.Recordset.RecordCount > 0 Then
  sSQL = "SELECT * " _
       & "FROM Trans_Costos " _
       & "WHERE FechaI = #" & FechaIni & "# " _
       & "AND FechaF = #" & FechaFin & "# " _
       & "AND Cta = '" & Codigo & "' "
  Select_Adodc AdoCostosTours, sSQL, False
  With AdoCostosTours.Recordset
       If .RecordCount <= 0 Then .AddNew
      .Fields("Cta") = Codigo
      .Fields("Guia") = TextGuia.Text
      .Fields("FechaI") = MBFechaI.Text
      .Fields("FechaF") = MBFechaF.Text
      .Fields("Fecha_De") = MBFechaDe.Text
      .Fields("Fecha_Al") = MBFechaAl.Text
      .Fields("PAX") = Val(CCur(TextPAX.Text))
      .Fields("Cotizacion") = Dolar
      .Update
  End With
  End If
  sSQL = "SELECT * " _
       & "FROM Trans_Costos " _
       & "WHERE FechaI = #" & FechaIni & "# " _
       & "AND FechaF = #" & FechaFin & "# " _
       & "AND Cta = '" & Codigo & "' "
  Select_Adodc AdoCostosTours, sSQL, False
  SQLMsg1 = DLCtas.Text
  'DBGrid1.Visible = False
  ImprimirSubCtas AdoSubCtaDet, AdoSubCtaDet1, AdoCostosTours, True, 1, 7, False, CheqResumen.value
  'DBGrid1.Visible = True
End Sub

Private Sub Command3_Click()

End Sub

Private Sub DGSubCtaDet_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  If CtrlDown And KeyCode = vbKeyF1 Then GenerarDataTexto RSubCtas, AdoSubCtaDet
End Sub

Private Sub Form_Activate()
  sSQL = "SELECT Codigo & Space(10) & Cuenta As Nombre_Cta " _
       & "FROM Catalogo_Cuentas " _
       & "WHERE TC = 'G' " _
       & "AND DG = 'D' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "ORDER BY Codigo "
  SelectDB_List DLCtas, AdoCtas, sSQL, "Nombre_Cta"
  sSQL = "SELECT Codigo & Space(10) & Cuenta As Nombre_Cta " _
       & "FROM Catalogo_Cuentas " _
       & "WHERE TC = 'I' " _
       & "AND DG = 'D' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "ORDER BY Codigo "
  SelectDB_List DLCtas1, AdoCtas1, sSQL, "Nombre_Cta"
  sSQL = "SELECT Detalle & ' - ' & Codigo As Nombre_SubCta " _
       & "FROM Catalogo_SubCtas " _
       & "WHERE TC = 'G' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "ORDER BY Detalle "
  SelectDB_Combo DCSubCta, AdoSubCta, sSQL, "Nombre_SubCta"
  RatonNormal
End Sub

Private Sub Form_Load()
   CentrarForm RSubCtas
   ConectarAdodc AdoAux
   ConectarAdodc AdoCtas
   ConectarAdodc AdoCtas1
   ConectarAdodc AdoSubCta
   ConectarAdodc AdoSubCtaDet
   ConectarAdodc AdoSubCtaDet1
   ConectarAdodc AdoCostosTours
End Sub

Private Sub MBFechaAl_GotFocus()
  MarcarTexto MBFechaAl
End Sub

Private Sub MBFechaAl_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBFechaAl_LostFocus()
  FechaValida MBFechaAl
End Sub

Private Sub MBFechaDe_GotFocus()
  MarcarTexto MBFechaDe
End Sub

Private Sub MBFechaDe_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBFechaDe_LostFocus()
  FechaValida MBFechaDe
End Sub

Private Sub MBFechaF_GotFocus()
  MarcarTexto MBFechaF
End Sub

Private Sub MBFechaF_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBFechaF_LostFocus()
  FechaValida MBFechaF
End Sub

Private Sub MBFechaI_GotFocus()
  MarcarTexto MBFechaI
End Sub

Private Sub MBFechaI_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBFechaI_LostFocus()
  FechaValida MBFechaI
End Sub

Private Sub TextGuia_GotFocus()
  TextoValido TextGuia
End Sub

Private Sub TextGuia_LostFocus()
  TextoValido TextGuia
  TextGuia.Text = UCaseStrg(TextGuia.Text)
End Sub

Private Sub TextPAX_LostFocus()
  If TextPAX.Text = "" Then TextPAX.Text = "1"
  If TextPAX.Text = "0" Then TextPAX.Text = "1"
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
  Select Case Button.key
    Case "Salir"
         Unload RSubCtas
    Case "Imprimir"
         Imprimir_Resultados
    Case "Periodo"
         Consultar
    Case "Costos"
  End Select
End Sub

