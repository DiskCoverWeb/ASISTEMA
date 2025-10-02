VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.5#0"; "comctl32.Ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form FImporta 
   Caption         =   "Importar Datos"
   ClientHeight    =   9180
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15960
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9180
   ScaleWidth      =   15960
   WindowState     =   2  'Maximized
   Begin VB.ListBox LstStatud 
      Appearance      =   0  'Flat
      BackColor       =   &H00C00000&
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   570
      Left            =   11865
      TabIndex        =   24
      Top             =   1575
      Visible         =   0   'False
      Width           =   5685
   End
   Begin VB.ComboBox CTP 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   345
      Left            =   9870
      TabIndex        =   9
      Text            =   "CD"
      Top             =   420
      Width           =   1170
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H8000000B&
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
      Left            =   10290
      Picture         =   "FImporta.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5145
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000B&
      Caption         =   "&Grabar Asiento"
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
      Left            =   10920
      TabIndex        =   22
      Top             =   5145
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton CommandButton1 
      Caption         =   "Command Button1"
      Height          =   645
      Left            =   11130
      TabIndex        =   10
      Top             =   105
      Width           =   1170
   End
   Begin MSDataGridLib.DataGrid DGAsiento 
      Bindings        =   "FImporta.frx":030A
      Height          =   2430
      Left            =   105
      TabIndex        =   15
      Top             =   5880
      Width           =   11670
      _ExtentX        =   20585
      _ExtentY        =   4286
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
   Begin VB.CommandButton Command3 
      Caption         =   "&Salir"
      Height          =   645
      Left            =   13545
      TabIndex        =   12
      Top             =   105
      Width           =   1065
   End
   Begin MSAdodcLib.Adodc AdoAux 
      Height          =   330
      Left            =   15225
      Top             =   2520
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
   Begin VB.CommandButton Command2 
      Caption         =   "Subir al Sistema"
      Height          =   645
      Left            =   12390
      TabIndex        =   11
      Top             =   105
      Width           =   1065
   End
   Begin MSAdodcLib.Adodc AdoAct 
      Height          =   330
      Left            =   15225
      Top             =   2100
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
      Caption         =   "Act"
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
   Begin MSMask.MaskEdBox MBoxCta_Inv 
      Height          =   330
      Left            =   1575
      TabIndex        =   3
      Top             =   420
      Visible         =   0   'False
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   582
      _Version        =   393216
      AllowPrompt     =   -1  'True
      AutoTab         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MBoxCta_Pat 
      Height          =   330
      Left            =   3570
      TabIndex        =   5
      Top             =   420
      Visible         =   0   'False
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   582
      _Version        =   393216
      AllowPrompt     =   -1  'True
      AutoTab         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSDataListLib.DataCombo DCLinea 
      Bindings        =   "FImporta.frx":0323
      DataSource      =   "AdoLinea"
      Height          =   360
      Left            =   5565
      TabIndex        =   7
      Top             =   420
      Visible         =   0   'False
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   635
      _Version        =   393216
      Text            =   "CxC Clientes"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSAdodcLib.Adodc AdoLinea 
      Height          =   330
      Left            =   15225
      Top             =   1680
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
      Caption         =   "Linea"
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
   Begin MSAdodcLib.Adodc AdoAsiento 
      Height          =   330
      Left            =   105
      Top             =   8400
      Width           =   4530
      _ExtentX        =   7990
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
   Begin MSMask.MaskEdBox MBFecha 
      Height          =   330
      Left            =   105
      TabIndex        =   13
      Top             =   8820
      Visible         =   0   'False
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
   Begin MSAdodcLib.Adodc AdoClientes 
      Height          =   330
      Left            =   15225
      Top             =   2940
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
   Begin MSMask.MaskEdBox MBFechaI 
      Height          =   330
      Left            =   105
      TabIndex        =   1
      Top             =   420
      Width           =   1380
      _ExtentX        =   2434
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
   Begin MSDataGridLib.DataGrid DGExcelAdodc 
      Bindings        =   "FImporta.frx":033A
      Height          =   2430
      Left            =   105
      TabIndex        =   23
      Top             =   840
      Width           =   11670
      _ExtentX        =   20585
      _ExtentY        =   4286
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
   Begin MSAdodcLib.Adodc AdoExcelAdodc 
      Height          =   330
      Left            =   105
      Top             =   3255
      Width           =   4530
      _ExtentX        =   7990
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
      Caption         =   "ExcelAdodc"
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
   Begin ComctlLib.ListView LstVwFTP 
      Height          =   645
      Left            =   11865
      TabIndex        =   25
      Top             =   840
      Visible         =   0   'False
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   1138
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      Icons           =   "ImgLstFTP"
      SmallIcons      =   "ImgLstFTP"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Archivos"
         Object.Width           =   3351
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Tamaño"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Modificado"
         Object.Width           =   2646
      EndProperty
   End
   Begin ComctlLib.ImageList ImgLstFTP 
      Left            =   13860
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   12
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FImporta.frx":0356
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FImporta.frx":0670
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FImporta.frx":098A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FImporta.frx":0C90
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FImporta.frx":0FAA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FImporta.frx":12C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FImporta.frx":15B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FImporta.frx":1DD0
            Key             =   "archivo"
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FImporta.frx":20EA
            Key             =   "carpeta"
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FImporta.frx":2404
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FImporta.frx":2642
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FImporta.frx":295C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &Tipo Comp."
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
      Left            =   9870
      TabIndex        =   8
      Top             =   105
      Width           =   1170
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
      Width           =   1380
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TOTALES "
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
      TabIndex        =   21
      Top             =   8400
      Width           =   1065
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Diferencia "
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
      Left            =   4620
      TabIndex        =   20
      Top             =   8400
      Width           =   1065
   End
   Begin VB.Label LblDiferencia 
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
      Height          =   330
      Left            =   5670
      TabIndex        =   19
      Top             =   8400
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
      Height          =   330
      Left            =   8505
      TabIndex        =   18
      Top             =   8400
      Width           =   1800
   End
   Begin VB.Label LabelHaber 
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
      Height          =   330
      Left            =   10290
      TabIndex        =   17
      Top             =   8400
      Width           =   1800
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &Linea de Facturacion:"
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
      TabIndex        =   6
      Top             =   105
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " CTA. &ACTIVO"
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
      Left            =   1575
      TabIndex        =   2
      Top             =   105
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.Label Label12 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " CTA. &PATRIM."
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
      Left            =   3570
      TabIndex        =   4
      Top             =   105
      Visible         =   0   'False
      Width           =   1905
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
      Left            =   105
      TabIndex        =   16
      Top             =   5145
      Width           =   10095
   End
End
Attribute VB_Name = "FImporta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim CtasProc() As CtasAsiento
Dim ContCtas As Integer
Dim NumTrans As Integer
Dim IdField As Integer
Dim SubModuloGasto As String
Dim SubModuloCxCxP As String
Dim SerieF As String
Dim RUC_CII As String
Dim RUC_CIF As String

Public Function Dato_Campo(Campo As Variant, Optional EsFecha As Boolean, Optional Mayuscula As Boolean) As String
Dim Codigo As String
    If IsNull(Campo) Then
       Codigo = Ninguno
    Else
       Codigo = CStr(Campo)
       Codigo = Replace(Codigo, vbCr, "")
       Codigo = Replace(Codigo, vbLf, "")
       Codigo = Replace(Codigo, "'", "")
       Codigo = Replace(Codigo, ",", ".")
       Codigo = Replace(Codigo, "&", "y")
       Codigo = Replace(Codigo, "#", "No.")
    End If
    Codigo = TrimStrg(Codigo)
    If EsFecha Then
       Codigo = Format(Replace(Codigo, "-", "/"), "dd/MM/yyyy")
       If Not IsDate(Codigo) Then Codigo = FechaSistema
    Else
       Codigo = Replace(Codigo, "-", "")
    End If
    If Codigo = "" Then Codigo = Ninguno
    If Mayuscula Then Codigo = UCaseStrg(Codigo)
    Dato_Campo = Codigo
End Function

Public Sub Datos_Default_Beneficiario(Optional EsFacturacion As Boolean)
    TBeneficiario.T = Normal
    TBeneficiario.FA = EsFacturacion
    TBeneficiario.TP = "E"
    TBeneficiario.Codigo = Ninguno
    TBeneficiario.CI_RUC = Ninguno
    TBeneficiario.Fecha = FechaSistema
    TBeneficiario.Fecha_N = FechaSistema
    TBeneficiario.Cliente = Ninguno
    TBeneficiario.Sexo = "M"
    TBeneficiario.Email1 = Ninguno
    TBeneficiario.Email2 = Ninguno
    TBeneficiario.Direccion = "SD"
    TBeneficiario.DirNumero = "SN"
    TBeneficiario.Telefono1 = "022000000"
    TBeneficiario.Celular = "0990000000"
    TBeneficiario.Ciudad = "QUITO"
    TBeneficiario.Prov = "17"
    TBeneficiario.Pais = "593"
    TBeneficiario.Grupo_No = "NUEVO"
    TBeneficiario.Cod_Ejec = Ninguno
    TBeneficiario.Cta_CxP = Ninguno
End Sub

Public Sub Procesar_Tipo_Carga(ArchivoSubido As Adodc)
Dim IdField As Integer
Dim IdName As String
    RatonReloj

'''       .Row = 1
'''       case 4
'''        If IsNumeric(IdName) Then Tipo_Carga = 1 Else Tipo_Carga = 2
'''       case 5
'''        If IdName = "Sexo" Then Tipo_Carga = 4
'''       'Tipo de subida de la informacion al sistema
'''       .Row = 0

    With ArchivoSubido.Recordset
     If .RecordCount > 0 Then
         Tipo_Carga = 1
         For IdField = 0 To .fields.Count - 1
             IdName = .fields(IdField).Name
            'MsgBox IdField & "-" & .Fields(IdField).Name
             Select Case IdField
               Case 0
                    If IdName = "CI_CLIENTE" Then Tipo_Carga = 20
                    If IdName = "OTRO_PROGRAMA" Then Tipo_Carga = 105         ' Importar Otro Plan de Cuentas
               Case 1
                    If IdName = "Codigo_Inv" Then Tipo_Carga = 9
                    If IdName = "CONSUMO (M3)" Then Tipo_Carga = 17
                    If IdName = "CI_RUC_PAS" Then Tipo_Carga = 10             ' Clientes Facturacion
               Case 2
                    If IdName = "Codigo_Nuevo" Then Tipo_Carga = 7
                    If IdName = "CodMateria" Then Tipo_Carga = 22
                    If IdName = "ALUMNOS" Then Tipo_Carga = 8
                    If IdName = "Autorizacion" Then Tipo_Carga = 25
                    If IdName = "FECHA DE NACIMIENTO" Then Tipo_Carga = 50
                    If IdName = "MADRE" Then Tipo_Carga = 51
                    If IdName = "LUGAR Y FECHA" Then Tipo_Carga = 52
               Case 3
                    If IdName = "Razon_Social" Then Tipo_Carga = 24
                    If IdName = "COMPROBANTE" Then Tipo_Carga = 12
                    If IdName = "CC" Then Tipo_Carga = 30
                    If IdName = "SALDO_ACT" Then Tipo_Carga = 32
               Case 4
                    If IdName = "DETALLE_DESCUENTO" Then Tipo_Carga = 19
                    If IdName = "CODIGO_EXT" Then Tipo_Carga = 4
               Case 5
                    If IdName = "Num_Lista" Then Tipo_Carga = 13
                    If IdName = "Correcto" Then Tipo_Carga = 16
                    If IdName = "VALOR DIARIO" Then Tipo_Carga = 26
                    If IdName = "CATEGORIA" Then Tipo_Carga = 11
               Case 6
                    If IdName = "FECHA_DOC" Then Tipo_Carga = 32
                    If IdName = "Desc_Item" Then Tipo_Carga = 101             ' Catalogo_Productos
               Case 7
                    If IdName = "CI_RUC_Codigo" Then Tipo_Carga = 15
                    If IdName = "PROFESION" Then Tipo_Carga = 18
                    If IdName = "Sustento" Then Tipo_Carga = 23
                    If IdName = "ruc_proveedor" Then Tipo_Carga = 38
               Case 8
                    If IdName = "emision" Then Tipo_Carga = 5
                    If IdName = "Telefono_Rep" Then Tipo_Carga = 107
                    If IdName = "CI_RUC_P_SUBMOD" Then Tipo_Carga = 99
               Case 9
                    If IdName = "Tipo_Abonos" Then Tipo_Carga = 6
                    If IdName = "CodMateria" Then Tipo_Carga = 21
                    If IdName = "AUXILIAR" Then Tipo_Carga = 28
               Case 10
                    If IdName = "EDUCATIVO" Then Tipo_Carga = 3
                    If IdName = "Diferencias" Then Tipo_Carga = 14
               Case 11
                    If IdName = "NOTA" Then Tipo_Carga = 106
               Case 12
                    If IdName = "Fecha" Then Tipo_Carga = 12
               Case 13
                    If IdName = "Bonificacion Adicional" Then Tipo_Carga = 29
                    If IdName = "Serie" Then Tipo_Carga = 103                ' Detalle_Factura
               Case 14
                    If IdName = "CUENTA ACTIVO" Then Tipo_Carga = 253
               Case 15
                    If IdName = "Razon_Social" Then Tipo_Carga = 102         ' Facturas
               Case 16
                    If IdName = "Grupo" Then Tipo_Carga = 100                ' Clientes_Facturacion
               Case 32
                    If IdName = "COD_MES" Then Tipo_Carga = 27
               Case 41
                    If IdName = "SUB_MOD_GASTO" Then Tipo_Carga = 254
               Case 42
             End Select
         Next IdField
     End If
    End With
    RatonNormal
End Sub

Private Sub Command1_Click()
    FechaTexto = FechaFin
    FechaComp = FechaTexto
    NumComp = ReadSetDataNum("Diario", True, False)
    Mensajes = "Esta seguro de Grabar el Comprobante No. " & NumComp
    Titulo = "Pregunta de grabación"
    If BoxMensaje = vbYes Then
       NumComp = ReadSetDataNum("Diario", True, True)
       DiarioCaja = NumComp
      'Grabacion del Comprobante
       Co.T = Normal
       Co.TP = CompDiario
       Co.Fecha = FechaTexto
       Co.Numero = NumComp
       If Modulo = "INVENTARIO" Then
          Co.Concepto = LblConcepto.Caption
       Else
          If Tipo_Carga = 12 Then
             Co.Concepto = LblConcepto.Caption
          Else
             If FechaIni = FechaFin Then
                Co.Concepto = "Gastos de Caja del " & FechaIni & ", Diario No. " & NumComp
             Else
                Co.Concepto = "Gastos de Caja del " & FechaIni & " al " & FechaFin & ", Diario No. " & NumComp
             End If
          End If
       End If
       Co.CodigoB = Ninguno
       Co.Efectivo = Debe
       Co.Monto_Total = Debe
       Co.T_No = Trans_No
       Co.Usuario = CodigoUsuario
       Co.Item = NumEmpresa
        
        Grabar_Comprobante Co
        Control_Procesos Normal, Co.Concepto
        ImprimirComprobantesDe False, Co
        
        sSQL = "UPDATE Trans_Compras " _
             & "SET Numero = " & Co.Numero & ",TP = '" & Co.TP & "' " _
             & "WHERE Fecha BETWEEN #" & BuscarFecha(FechaIni) & "# and #" & BuscarFecha(FechaFin) & "# " _
             & "AND Item = '" & NumEmpresa & "' " _
             & "AND Periodo = '" & Periodo_Contable & "' " _
             & "AND Numero = -1 " _
             & "AND TP = 'NN' "
        Ejecutar_SQL_SP sSQL
        sSQL = "UPDATE Trans_Air " _
             & "SET Numero = " & Co.Numero & ",TP = '" & Co.TP & "' " _
             & "WHERE Fecha BETWEEN #" & BuscarFecha(FechaIni) & "# and #" & BuscarFecha(FechaFin) & "# " _
             & "AND Item = '" & NumEmpresa & "' " _
             & "AND Periodo = '" & Periodo_Contable & "' " _
             & "AND Numero = -1 " _
             & "AND TP = 'NN' "
        Ejecutar_SQL_SP sSQL
        FA.Numero = Co.Numero
        FA.TP = Co.TP
        SRI_Crear_Clave_Acceso_Retenciones FA, False
        Unload FImporta
    End If
End Sub

'Procesa importaciones
Private Sub Command2_Click()
    RatonReloj
    Progreso_Barra.Mensaje_Box = "Empezando a copiar archivos"
    Progreso_Iniciar
    Progreso_Barra.Valor_Maximo = 100
    
    TextoImprimio = ""
    sSQL = "SELECT Codigo,CI_RUC,Cliente " _
         & "FROM Clientes " _
         & "WHERE LEN(Codigo) > 0 " _
         & "ORDER BY Cliente "
    Select_Adodc AdoClientes, sSQL

    sSQL = "DELETE * " _
         & "FROM Tabla_Temporal " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Modulo = '" & NumModulo & "' " _
         & "AND CodigoU = '" & CodigoUsuario & "' "
    Ejecutar_SQL_SP sSQL

   'MsgBox Tipo_Carga
   
    RatonReloj
    Select Case Tipo_Carga
      Case 1: Importar_Facturas
      'Case 2: Importar_Facturas_2
      'Case 3: Importar_Facturas_3
      Case 4: Importar_Plan_Cuentas
      Case 5: Importar_Contabilidad
      Case 6: Importar_Abonos
      'Case 7: Transferir_Plan_Cuentas
      'Case 8: Importar_Alumnos_Contabilidad
      Case 9: Importar_Inventarios
      Case 10: Importar_Personas
      'Case 11: Importar_Facturas_Contabilidad
      'Case 12: Importar_SubModulo
      'Case 14: Importar_Sobrantes_Faltantes
      Case 15: Importar_Abonos_Transferencias
      'Case 16: Cambio_Numero_Secuencial
      'Case 17: Importar_Consumos
      Case 18: Importar_Empleados
      Case 19: Importar_Descuento_Empleados
      Case 20: Importar_Facturas_Farmacias
      'Case 21: Importar_Notas_Materias
      'Case 22: Importar_Informes_Materias
      'Case 23: Importar_Codigos_Retenciones
      Case 24: Importar_Retenciones_Farmacia
               Generar_Asiento_Compras True
      'Case 25:Importar_Compras_Inventario_Lotes
      Case 25: Importar_Autorizacion_Electronica
      'Case 26: Diarios_Automaticos
      Case 27: Importar_Compras_Diarias
      'Case 28: Diarios_Automaticos_Ventas
      'Case 29: Importar_Catalogo_RolPagos
      'Case 30: Diarios_Automaticos_Bata
      'Case 32: Diarios_Automaticos_Bata_SC_CxC
      'Case 38: Diarios_Automaticos_Contabilidad_Externa
    '''           Case Else
    '''                Importar_Compras
    '''                Generar_Asiento_Compras
    
    '''    Case "CAJACREDITO"
    '''         Importar_Depositos
      'Case 50, 51, 52: Importar_Parroquias
      Case 99: Importar_Contabilidad_SubModulos
      Case 100: Importar_Estudiantes_Representantes
      'Case 101: Importar_Contabilidad_Catalogo_Productos
      'Case 102: Importar_Contabilidad_Facturas
      'Case 103: Importar_Contabilidad_Detalle_Factura
      Case 104: Importar_Personas
      'Case 105: Importar_Plan_Cuentas_Externas
      Case 106: Importar_Estudiantes_PreFacturas
      Case 107: Importar_Actualizacion_Estudiantes
      Case 253: Importar_Activos
      Case Else: MsgBox "Este archivo no tiene formatos establecidos"
    End Select
    RatonNormal
    Progreso_Barra.Mensaje_Box = "Proceso Terminado"
    Progreso_Final
    If Len(TextoImprimio) > 2 Then FInfoError.Show Else MsgBox "Proceso Terminado"
    MBFechaI.SetFocus
End Sub

Private Sub Command3_Click()
    Unload FImporta
End Sub

Private Sub Command4_Click()
   Mensajes = "Seguro de Encerar Compras y Retenciones"
   Titulo = "Pregunta de Eliminación"
   If BoxMensaje = vbYes Then
      sSQL = "DELETE * " _
           & "FROM Trans_Compras " _
           & "WHERE Item = '" & NumEmpresa & "' " _
           & "AND Periodo = '" & Periodo_Contable & "' " _
           & "AND TP = 'NN' " _
           & "AND Numero = -1 "
      Ejecutar_SQL_SP sSQL
      sSQL = "DELETE * " _
           & "FROM Trans_Air " _
           & "WHERE Item = '" & NumEmpresa & "' " _
           & "AND Periodo = '" & Periodo_Contable & "' " _
           & "AND TP = 'NN' " _
           & "AND Numero = -1 "
      Ejecutar_SQL_SP sSQL
   End If
End Sub

''''Private Sub Command5_Click()
''''Dim NumFile As Long
''''Dim FinComp As Boolean
''''Dim NuevoComp As Boolean
''''Dim I As Integer
''''Dim Codigos(200) As String
''''Dim ValCC(200) As String
''''Dim TotDebe(200) As Currency
''''Dim TotHaber(200) As Currency
''''
''''  CDialogDir.Filter = "Todos los archivos|*.*"
''''  CDialogDir.InitDir = RutaSysBases & "\Datos"
''''  RutaGeneraFile = SelectDialogFile(CDialogDir, SelectAll)
''''  If RutaGeneraFile <> "" Then
''''     DGAsiento.Visible = False
''''     Progreso_Barra.Mensaje_Box = ""
''''     Progreso_Iniciar
''''     Contador = 0
''''     NumFile = FreeFile
''''     Open RutaGeneraFile For Input As #NumFile
''''     Do While Not EOF(NumFile)
''''        Line Input #NumFile, Cod_Field
''''        Contador = Contador + 1
''''     Loop
''''     Close #NumFile
''''     Progreso_Barra.Valor_Maximo = (Contador * 2) + 100
''''
''''    'Iniciamos la Carga
''''     sSQL = "DELETE * " _
''''          & "FROM Asiento_SC " _
''''          & "WHERE Item = '" & NumEmpresa & "' " _
''''          & "AND T_No = " & Trans_No & " " _
''''          & "AND CodigoU = '" & CodigoUsuario & "' "
''''     Ejecutar_SQL_SP sSQL
''''
''''     IniciarAsientosDe DGAsiento, AdoAsiento
''''
''''     Contador = 0
''''     Progreso_Barra.Incremento = 0
''''     For I = 1 To 200
''''        Codigos(I) = ""
''''        ValCC(I) = ""
''''        TotDebe(I) = 0
''''        TotHaber(I) = 0
''''     Next I
''''     I = 1
''''     NuevoComp = False
''''     FinComp = False
''''     Co.Numero = 0
''''     Co.Fecha = Ninguno
''''     Cadena = ""
''''     NumFile = FreeFile
''''     Open RutaGeneraFile For Input As #NumFile
''''     Do While Not EOF(NumFile)
''''        Line Input #NumFile, Cod_Field
''''        Progreso_Barra.Mensaje_Box = "Procesando Comprobante: " & Co.Fecha & " - " & Co.TP & " - " & Co.Numero
''''        Progreso_Esperar
''''        If IsDate(Co.Fecha) And Co.Numero <> Val(MidStrg(Cod_Field, 8, 10)) And FinComp And Co.Numero > 0 Then
''''          'Procesamos el Asiento Contable
''''           For I = 1 To 200
''''               If Codigos(I) <> "" Then
''''                  Cta = Codigos(I)
''''                  CodigoCC = ValCC(I)
''''                  DetalleComp = "."
''''                  NoCheque = "."
''''                  If TotDebe(I) > 0 Then OpcDH = "1"
''''                  If TotHaber(I) > 0 Then OpcDH = "2"
''''
''''                  If OpcDH = "1" Then
''''                     ValorDH = TotDebe(I)
''''                     InsertarAsientos AdoAsiento, Cta, 0, ValorDH, 0
''''                  Else
''''                     ValorDH = TotHaber(I)
''''                     InsertarAsientos AdoAsiento, Cta, 0, 0, ValorDH
''''                  End If
''''                  'valCC(I) = ""
''''               End If
''''           Next I
''''           FechaComp = Co.Fecha
''''           FechaTexto = Co.Fecha
''''           Co.Concepto = Co.Concepto & ", Ref. " & Format(Co.Numero, "00000000")
''''           NumComp = ReadSetDataNum("Diario", True, True)
''''           DiarioCaja = NumComp
''''           Progreso_Barra.Mensaje_Box = "Procesando Comprobante: " & Co.Fecha & " - " & Co.TP & " - " & Co.Numero
''''           Progreso_Esperar True
''''
''''          'Grabacion del Comprobante
''''           Co.T = Normal
''''           Co.TP = CompDiario
''''           Co.Numero = NumComp
''''           Co.CodigoB = Ninguno
''''           Co.Efectivo = Debe
''''           Co.Monto_Total = Debe
''''           Co.T_No = Trans_No
''''           Co.Usuario = CodigoUsuario
''''           Co.Item = NumEmpresa
''''           GrabarComprobante Co
''''
''''          'MsgBox "Comprobante: " & Co.TP & " - " & Co.Numero & vbCrLf & Co.Concepto
''''           DGAsiento.Visible = False
''''           NuevoComp = False
''''           FinComp = False
''''           Co.Numero = 0
''''           Co.Fecha = Ninguno
''''           For I = 1 To 200
''''               Codigos(I) = ""
''''               ValCC(I) = ""
''''               TotDebe(I) = 0
''''               TotHaber(I) = 0
''''           Next I
''''           I = 1
''''           sSQL = "DELETE * " _
''''                & "FROM Asiento_SC " _
''''                & "WHERE Item = '" & NumEmpresa & "' " _
''''                & "AND T_No = " & Trans_No & " " _
''''                & "AND CodigoU = '" & CodigoUsuario & "' "
''''           Ejecutar_SQL_SP sSQL
''''           IniciarAsientosDe DGAsiento, AdoAsiento
''''        End If
''''
''''        If MidStrg(Cod_Field, 64, 20) = "COMPROBANTE CONTABLE" Then
''''           Co.Numero = Val(MidStrg(Cod_Field, 88, 10))
''''           NuevoComp = True
''''        End If
''''        If MidStrg(Cod_Field, 2, 22) = "FECHA DE COMPROBANTE :" Then Co.Fecha = MidStrg(Cod_Field, 25, 11)
''''        If MidStrg(Cod_Field, 2, 22) = "GLOSA GENERAL .......:" Then Co.Concepto = MidStrg(Cod_Field, 25, 100)
''''        If MidStrg(Cod_Field, 2, 11) = "HECHO POR :" Then FinComp = True
''''        If IsDate(Co.Fecha) And Len(Co.Concepto) > 1 And Co.Numero <> 0 And NuevoComp Then
''''           Codigo = MidStrg(Cod_Field, 2, 7)
''''           If IsNumeric(Codigo) Then
''''              Codigos(I) = MidStrg(Cod_Field, 2, 7)
''''              ValCC(I) = TrimStrg(MidStrg(Cod_Field, 11, 7))
''''              TotDebe(I) = Val(Replace(MidStrg(Cod_Field, 126, 17), ",", ""))
''''              TotHaber(I) = Val(Replace(MidStrg(Cod_Field, 144, 17), ",", ""))
''''              I = I + 1
''''           End If
''''        End If
'''''''        MsgBox Cod_Field & vbCrLf _
'''''''             & "-------------------------------------------" & vbCrLf _
'''''''             & MidStrg(Cod_Field, 126, 17) & " <=====> " & MidStrg(Cod_Field, 144, 17) & vbCrLf _
'''''''             & "-------------------------------------------" & vbCrLf _
'''''''             & Co.Numero & vbCrLf _
'''''''             & "-------------------------------------------" & vbCrLf _
'''''''             & Co.Fecha & vbCrLf _
'''''''             & "-------------------------------------------" & vbCrLf _
'''''''             & Co.Concepto & vbCrLf _
'''''''             & "-------------------------------------------" & vbCrLf _
'''''''             & FinComp & vbCrLf _
'''''''             & Cadena
''''     Loop
''''     Close #NumFile
''''     Progreso_Final
''''  End If
''''  DGAsiento.Visible = True
''''  MsgBox "Proceso Terminado"
''''End Sub

Private Sub CommandButton1_Click()
Dim Tipo_Carga1 As Integer
  DGAsiento.Visible = False
  DGExcelAdodc.Visible = False
  RutaOrigen = SelectDialogFile(RutaSysBases)
 'Le pasamos el Path del Libro y una variable de tipo T_Rango para retornar los valores
  If RutaOrigen <> "" Then
     Tipo_Carga = 0
     Set AdoExcelAdodc.Recordset = Importar_Excel_AdoDB(ftp, LstStatud, LstVwFTP, RutaOrigen)
     AdoExcelAdodc.Caption = "Reg. No. " & AdoExcelAdodc.Recordset.RecordCount & ", Archivo: " & RutaOrigen
     Procesar_Tipo_Carga AdoExcelAdodc
  End If
  DGAsiento.Visible = True
  DGExcelAdodc.Visible = True
  Label5.Caption = " FECHA (" & Tipo_Carga & "):"
End Sub

Private Sub CTP_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub DCLinea_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCLinea_LostFocus()
  If Modulo = "FACTURACION" Then
     FA.Cod_CxC = DCLinea.Text
     FA.Fecha = MBFechaI
     Lineas_De_CxC FA
  End If
End Sub

Private Sub Form_Activate()
   RatonReloj
   Set ftp = New cFTP
   
   EsFileCSV = False
   CTP.Clear
   CTP.AddItem "CD"
   CTP.AddItem "CE"
   CTP.AddItem "CI"
   CTP.Text = "CD"
   
   MBFecha = FechaSistema
   DGExcelAdodc.width = MDI_X_Max - 100
   DGAsiento.width = MDI_X_Max - 100
   AdoExcelAdodc.width = MDI_X_Max - 100
   
   DGExcelAdodc.Height = (MDI_Y_Max - DGExcelAdodc.Top - 1200) / 2

   AdoExcelAdodc.Top = DGExcelAdodc.Top + DGExcelAdodc.Height + 10
   LblConcepto.Top = AdoExcelAdodc.Top + AdoExcelAdodc.Height + 10
   
   LblConcepto.width = MDI_X_Max - (Command1.width + Command4.width) - 130
   
   DGAsiento.Height = DGExcelAdodc.Height
   
   Command4.Left = LblConcepto.Left + LblConcepto.width + 20
   Command1.Left = Command4.Left + Command4.width + 20
   Command1.Top = AdoExcelAdodc.Top + AdoExcelAdodc.Height + 10
   Command4.Top = AdoExcelAdodc.Top + AdoExcelAdodc.Height + 10
   DGAsiento.Top = LblConcepto.Top + LblConcepto.Height + 10
   
   AdoAsiento.Top = DGAsiento.Top + DGAsiento.Height + 10
   Label1.Top = DGAsiento.Top + DGAsiento.Height + 10
   Label11.Top = DGAsiento.Top + DGAsiento.Height + 10
   LabelDebe.Top = DGAsiento.Top + DGAsiento.Height + 10
   LabelHaber.Top = DGAsiento.Top + DGAsiento.Height + 10
   LblDiferencia.Top = DGAsiento.Top + DGAsiento.Height + 10
   
'''   Select Case Modulo
'''     Case "FACTURACION", "EDUCATIVO"
'''          MSFlexGrid1.Height = MDI_Y_Max - MSFlexGrid1.Top - 600
'''          DGAsiento.Visible = False
'''   End Select
   
  Trans_No = 199
  FormatoMaskCta MBoxCta_Inv
  FormatoMaskCta MBoxCta_Pat
  Select Case Modulo
    Case "INVENTARIO"
         Label2.Caption = "Seleccione la Cuenta"
         Command1.Visible = True
         DCLinea.Visible = True
         Label2.Visible = True
         sSQL = "SELECT Codigo, Cuenta " _
              & "FROM Catalogo_Cuentas " _
              & "WHERE Item = '" & NumEmpresa & "' " _
              & "AND Periodo = '" & Periodo_Contable & "' " _
              & "AND TC = 'P' " _
              & "ORDER BY Cuenta "
         SelectDB_Combo DCLinea, AdoLinea, sSQL, "Cuenta"
    Case "CONTABILIDAD"
         Label2.Visible = False
         Label7.Visible = False
         Label12.Visible = False
         MBoxCta_Pat.Visible = False
         MBoxCta_Inv.Visible = False
         DCLinea.Visible = False
         Command1.Visible = True
         Command4.Visible = True
    Case "FACTURACION"
         sSQL = "SELECT * " _
              & "FROM Catalogo_Lineas " _
              & "WHERE TL <> " & Val(adFalse) & " " _
              & "AND Item = '" & NumEmpresa & "' " _
              & "AND Fact IN ('FA','NV') " _
              & "AND Periodo = '" & Periodo_Contable & "' " _
              & "ORDER BY Serie, Codigo "
         SelectDB_Combo DCLinea, AdoLinea, sSQL, "Concepto"
         Label1.Visible = True
         Label11.Visible = False
         LblConcepto.Visible = False
         LblDiferencia.Visible = False
         LabelDebe.Visible = False
         LabelHaber.Visible = False
         Command1.Visible = False
         DGAsiento.Visible = False
         AdoAsiento.Visible = False
         Label2.Visible = True
         DCLinea.Visible = True
         'MBoxCta_Inv.SetFocus
         DCLinea.SetFocus
    Case "EDUCATIVO"
         sSQL = "SELECT * " _
              & "FROM Catalogo_Cursos " _
              & "WHERE Item = '" & NumEmpresa & "' " _
              & "AND Periodo = '" & Periodo_Contable & "' " _
              & "AND LEN(Curso)>4 " _
              & "ORDER BY Curso "
         SelectDB_Combo DCLinea, AdoLinea, sSQL, "Curso"
         Label2.Caption = "SELECCIONE EL CURSO"
         Label2.Visible = True
         DCLinea.Visible = True
    Case "ROL PAGOS"
         Label1.Visible = True
         Label11.Visible = False
         LblConcepto.Visible = False
         LblDiferencia.Visible = False
         LabelDebe.Visible = False
         LabelHaber.Visible = False
         Command1.Visible = False
         DGAsiento.Visible = False
         AdoAsiento.Visible = False
         Label2.Visible = True
         DCLinea.Visible = True
    Case "CAJACREDITO"
  End Select
  RatonNormal
End Sub

Private Sub Form_Load()
   Me.Caption = "Importar desde Excel"
   CommandButton1.Caption = "Importar" & vbCrLf & "de" & vbCrLf & "Excell"
   ConectarAdodc AdoAux
   ConectarAdodc AdoAct
   ConectarAdodc AdoLinea
   ConectarAdodc AdoAsiento
   ConectarAdodc AdoClientes
   ConectarAdodc AdoExcelAdodc
End Sub

Public Sub Importar_Activos()
   Progreso_Barra.Mensaje_Box = "Importacion de Activos Fijos"
   Progreso_Iniciar
   sSQL = "DELETE * " _
        & "FROM Catalogo_Productos " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND Periodo = '" & Periodo_Contable & "' " _
        & "AND TDP = 'ACT' "
   Ejecutar_SQL_SP sSQL
   With AdoExcelAdodc.Recordset
    If .RecordCount > 0 Then
        Progreso_Barra.Valor_Maximo = (.RecordCount * 2) + 100
        Do While Not .EOF
           SetAdoAddNew "Catalogo_Productos"
           SetAdoFields "TDP", "ACT"
           SetAdoFields "T", Normal
           SetAdoFields "TC", "P"
           For IdField = 0 To .fields.Count - 1
               Codigo = TrimStrg(.fields(IdField))
               Codigo1 = Codigo
               Select Case IdField + 1
                 Case 1: SetAdoFields "Fecha", Codigo
                 Case 2: If Codigo = "9" Then Codigo = "9999999999"
                         If MidStrg(Codigo, 1, 1) = "9" Then Codigo = "9999999999"
                         SetAdoFields "Codigo_P", MidStrg(Codigo, 1, 10)
                 Case 3: SetAdoFields "Factura", Val(Codigo)
                 Case 4: SetAdoFields "Vida_Util", Val(Codigo)
                 Case 5: SetAdoFields "Tipo", Codigo
                 Case 6: SetAdoFields "Departamento", Codigo
                 Case 7: SetAdoFields "Producto", Codigo
                 Case 8: If Val(Codigo) = 0 Then Codigo = "1"
                         SetAdoFields "Cantidad", Val(Codigo)
                 Case 9: SetAdoFields "Ubicacion", Codigo
                 Case 10: SetAdoFields "Valor_Historico", Val(Codigo)
                 Case 11: SetAdoFields "Valor_Actual", Val(Codigo)
                 Case 13: SetAdoFields "Codigo_R", Format$(Codigo, "0000000000")
                 Case 14: SetAdoFields "Valor_Depreciacion", Val(Codigo)
                 Case 15: SetAdoFields "Cta_Inventario", CambioCodigoCta(Codigo)
                 Case 16: SetAdoFields "Cta_Ventas", CambioCodigoCta(Codigo)
                 Case 22: SetAdoFields "Codigo_Inv", Codigo
                          SetAdoFields "Codigo_Barra", Replace(Codigo, ".", "")
               End Select
           Next
           SetAdoUpdate
           Progreso_Esperar
          .MoveNext
        Loop
    End If
   End With

  sSQL = "SELECT Codigo_Inv,Departamento,Ubicacion,Tipo " _
       & "FROM Catalogo_Productos " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND TDP = 'ACT' " _
       & "ORDER BY Codigo_Inv "
  Select_Adodc AdoAux, sSQL
  With AdoAux.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          CodigoInv = .fields("Codigo_Inv")
          Codigo1 = .fields("Departamento")
          Codigo2 = .fields("Ubicacion")
          Codigo3 = .fields("Tipo")
          CodigoInv = CodigoCuentaSup(CodigoInv)
          sSQL = "SELECT Codigo_Inv " _
               & "FROM Catalogo_Productos " _
               & "WHERE Item = '" & NumEmpresa & "' " _
               & "AND Periodo = '" & Periodo_Contable & "' " _
               & "AND TC = 'I' " _
               & "AND TDP = 'ACT' " _
               & "AND Codigo_Inv = '" & CodigoInv & "' "
          Select_Adodc AdoAct, sSQL
          If AdoAct.Recordset.RecordCount <= 0 Then
             SetAdoAddNew "Catalogo_Productos"
             SetAdoFields "T", Normal
             SetAdoFields "TC", "I"
             SetAdoFields "TDP", "ACT"
             SetAdoFields "Codigo_Inv", CodigoInv
             SetAdoFields "Producto", UCaseStrg(Codigo3)
             SetAdoUpdate
          End If
          
          CodigoInv = CodigoCuentaSup(CodigoInv)
          sSQL = "SELECT Codigo_Inv " _
               & "FROM Catalogo_Productos " _
               & "WHERE Item = '" & NumEmpresa & "' " _
               & "AND Periodo = '" & Periodo_Contable & "' " _
               & "AND TC = 'I' " _
               & "AND TDP = 'ACT' " _
               & "AND Codigo_Inv = '" & CodigoInv & "' "
          Select_Adodc AdoAct, sSQL
          If AdoAct.Recordset.RecordCount <= 0 Then
             SetAdoAddNew "Catalogo_Productos"
             SetAdoFields "T", Normal
             SetAdoFields "TC", "I"
             SetAdoFields "TDP", "ACT"
             SetAdoFields "Codigo_Inv", CodigoInv
             SetAdoFields "Producto", UCaseStrg(Codigo2)
             SetAdoUpdate
          End If
          CodigoInv = CodigoCuentaSup(CodigoInv)
          
          sSQL = "SELECT Codigo_Inv " _
               & "FROM Catalogo_Productos " _
               & "WHERE Item = '" & NumEmpresa & "' " _
               & "AND Periodo = '" & Periodo_Contable & "' " _
               & "AND TC = 'I' " _
               & "AND TDP = 'ACT' " _
               & "AND Codigo_Inv = '" & CodigoInv & "' "
          Select_Adodc AdoAct, sSQL
          If AdoAct.Recordset.RecordCount <= 0 Then
             SetAdoAddNew "Catalogo_Productos"
             SetAdoFields "T", Normal
             SetAdoFields "TC", "I"
             SetAdoFields "TDP", "ACT"
             SetAdoFields "Codigo_Inv", CodigoInv
             SetAdoFields "Producto", UCaseStrg(Codigo1)
             SetAdoUpdate
          End If
          Progreso_Barra.Mensaje_Box = "Generando Agrupaciones de Activos Fijos"
          Progreso_Esperar
         .MoveNext
       Loop
   End If
  End With
  Progreso_Final
End Sub

Public Sub Importar_Autorizacion_Electronica()
Dim I As Long
Dim N As Long
   With AdoExcelAdodc.Recordset
    If .RecordCount > 0 Then
        Progreso_Barra.Valor_Maximo = .RecordCount + 100
        Do While Not .EOF
           For IdField = 0 To .fields.Count - 1
               Codigo = TrimStrg(Replace(.fields(IdField), "-", ""))
               Select Case IdField + 1
                 Case 1
                      If Len(Codigo) > 1 Then Numero = Val(Codigo)
                 Case 2
                     'Factura 001005000000001
                      If Len(Codigo) > 1 Then
                         SerieFactura = MidStrg(Codigo, 9, 6)
                         Factura_No = Val(MidStrg(Codigo, 15, 9))
                      End If
                 Case 3
                     'CA:0109201601179186149300120010050000000010909201615
                     'NA:3009201612333317918614930010962016447
                      If (I Mod 2) = 0 Then Autorizacion = MidStrg(Codigo, 4, 49) Else CodigoC = MidStrg(Codigo, 4, 49)
                 Case 4
                     '30/09/2016  12:33:33
                      If Len(Codigo) > 1 Then
                         Mifecha = MidStrg(Codigo, 1, 10)
                         MiHora = MidStrg(Codigo, 13, 9)
                      End If
               End Select
           Next IdField
           If (I Mod 2) = 0 Then
              sSQL = "UPDATE Facturas " _
                   & "SET Autorizacion = '" & Autorizacion & "' " _
                   & "WHERE Item = '" & NumEmpresa & "' " _
                   & "AND Periodo = '" & Periodo_Contable & "' " _
                   & "AND Serie = '" & SerieFactura & "' " _
                   & "AND Factura = " & Factura_No & " " _
                   & "AND Clave_Acceso = '" & CodigoC & "' " _
                   & "AND LEN(Autorizacion) = 13 "
              Ejecutar_SQL_SP sSQL
               
              sSQL = "UPDATE Detalle_Factura " _
                   & "SET Autorizacion = '" & Autorizacion & "' " _
                   & "WHERE Item = '" & NumEmpresa & "' " _
                   & "AND Periodo = '" & Periodo_Contable & "' " _
                   & "AND Serie = '" & SerieFactura & "' " _
                   & "AND Factura = " & Factura_No & " " _
                   & "AND LEN(Autorizacion) = 13 "
              Ejecutar_SQL_SP sSQL
               
              sSQL = "UPDATE Trans_Abonos " _
                   & "SET Autorizacion = '" & Autorizacion & "' " _
                   & "WHERE Item = '" & NumEmpresa & "' " _
                   & "AND Periodo = '" & Periodo_Contable & "' " _
                   & "AND Serie = '" & SerieFactura & "' " _
                   & "AND Factura = " & Factura_No & " " _
                   & "AND LEN(Autorizacion) = 13 "
              Ejecutar_SQL_SP sSQL
           End If
          .MoveNext
        Loop
        Me.Caption = "Importar de FlexGrid a Sistema " & I & " de " & Rango.NumFila2
    End If
  End With
  MsgBox "Proceso Terminado"
End Sub

Public Sub Importar_Facturas()
Dim SubTotal As Currency
Dim SubTotalIVA As Currency
Dim SubTotalServicio As Currency
Dim SubTotalPorcComision As Currency
Dim SubTotalDescuento As Currency

  DGExcelAdodc.Visible = False
  
  Progreso_Iniciar
  Encerar_Factura FA
  
  Iniciar_Asiento_Beneficiario
       
  FA.Fecha = MBFechaI
  FA.Cod_CxC = DCLinea.Text
  FA.TC = "FA"
  FA.Factura = 0
 'Eliminando facturas del excel
  With AdoExcelAdodc.Recordset
   If .RecordCount > 0 Then
       Progreso_Barra.Valor_Maximo = (.RecordCount * 2) + 100
       Progreso_Barra.Incremento = 0
      .MoveFirst
       RUC_CII = Dato_Campo(.fields(0))
       Beneficiario = Dato_Campo(.fields(12), , True)
       FA.Serie = TrimStrg(MidStrg(Dato_Campo(.fields(10)), 1, 6))
       If Len(FA.Serie) < 6 Then FA.Serie = "001001"
       FA.Desde = Val(.fields(2))
       FA.Hasta = FA.Desde
       Insertar_Asiento_Beneficiario RUC_CII, Beneficiario
       Do While Not .EOF
          RUC_CIF = Dato_Campo(.fields(0))
          SerieF = TrimStrg(MidStrg(Dato_Campo(.fields(10)), 1, 6))
          If Len(SerieF) < 6 Then SerieF = "001001"
          FA.Fecha = Dato_Campo(.fields(1), True)
          Mifecha = FA.Fecha
          If IsNull(.fields(2)) Then FA.Factura = FA.Desde Else FA.Factura = Val(.fields(2))
          
          If RUC_CII <> RUC_CIF Then
            'MsgBox RUC_CII & vbCrLf & Beneficiario
             Insertar_Asiento_Beneficiario RUC_CII, Beneficiario
             RUC_CII = Dato_Campo(.fields(0))
             Beneficiario = Dato_Campo(.fields(12), , True)
          End If

          If FA.Serie <> SerieF Then
             Progreso_Barra.Mensaje_Box = "Eliminando Facturas entre " & Format$(FA.Desde, "000000000") & " al " & Format$(FA.Hasta, "000000000")
             Progreso_Esperar True
             Eliminar_Facturas
             FA.Serie = TrimStrg(MidStrg(Dato_Campo(.fields(10)), 1, 6))
             If Len(FA.Serie) < 6 Then FA.Serie = "001001"
             FA.Desde = Val(.fields(2))
             FA.Hasta = FA.Desde
          End If

          If FA.Factura < FA.Desde Then FA.Desde = FA.Factura
          If FA.Factura > FA.Hasta Then FA.Hasta = FA.Factura
          Progreso_Barra.Mensaje_Box = "Actualizando RUC/CI/Pasaporte: " & RUC_CIF & ", Fecha: " & Mifecha
          Progreso_Esperar
         .MoveNext
       Loop
       Insertar_Asiento_Beneficiario RUC_CII, Beneficiario
       Progreso_Barra.Mensaje_Box = "Eliminando Facturas entre " & Format$(FA.Desde, "000000000") & " al " & Format$(FA.Hasta, "000000000")
       Progreso_Esperar True
       Eliminar_Facturas
   End If
  End With
  Actualizar_Asiento_Beneficiario_Clientes True
  
 'Empezamos la importacion de las facturas
 '-------------------------------------------
  With AdoExcelAdodc.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      'MsgBox .Rows & vbCrLf & .Cols
       FA.Cod_CxC = Ninguno
       FA.TC = "FA"
       Leer_Encabezado_FA
       
       Lineas_De_CxC FA
       FechaTexto = FA.Fecha
       SerieFactura = FA.Serie
       
       Fecha_Vence = FA.Vencimiento
       Cta_Cobrar = FA.Cta_CxP
       Bandera = False
       Evaluar = True
       Ln_No = 1
       TA.Abono = 0
       FA.Servicio = 0
       FA.Propina = 0
       If Len(FA.Cod_CxC) <= 1 Then FA.Cod_CxC = Ninguno
       sSQL = "DELETE * " _
            & "FROM Asiento_F " _
            & "WHERE Item = '" & NumEmpresa & "' " _
            & "AND CodigoU = '" & CodigoUsuario & "' "
       Ejecutar_SQL_SP sSQL
       Do While Not .EOF
          If Not IsNull(.fields(0)) And Not IsNull(.fields(1)) Then
             Factura_No = Val(Dato_Campo(.fields(2)))
             FechaTexto = Dato_Campo(.fields(1), True)
             SerieFactura = MidStrg(Dato_Campo(.fields(10)), 1, 6)
             If Len(SerieFactura) < 6 Then SerieFactura = "001001"
            'MsgBox FA.Factura & vbCrLf & Factura_No
             If FA.Factura <> Factura_No Then
                'MsgBox FA.Cod_CxC & "..."
                'FA.Autorizacion = Autorizacion
                 Calculos_Totales_Factura FA
                'MsgBox FA.Factura & "|-> "
                 If Len(FA.Autorizacion) > 13 Then FA.ClaveAcceso = FA.Autorizacion
                 Grabar_Factura FA, True, True
                 If TA.Abono > 0 Then
                    TA.Fecha = FA.Fecha
                    TA.Abono = FA.Total_MN
                    TA.CodigoC = FA.CodigoC
                    TA.TP = FA.TC
                    TA.Serie = FA.Serie
                    TA.Factura = FA.Factura
                    TA.Autorizacion = FA.Autorizacion
                    TA.Cta_CxP = FA.Cta_CxP
                    Grabar_Abonos TA, True
                 End If
                'MsgBox FA.Factura & " ------->"
                 Leer_Encabezado_FA
                 FechaTexto = FA.Fecha
                 SerieFactura = FA.Serie
                 Fecha_Vence = FA.Vencimiento
                 Cta_Cobrar = FA.Cta_CxP
                 Bandera = False
                 Evaluar = True
                 Ln_No = 1
                 TA.Abono = 0
                 FA.Servicio = 0
                 FA.Propina = 0
                 
                 Progreso_Barra.Mensaje_Box = "Documento " & FA.TC & " No. " & Format$(FA.Factura, "000000000") & ", gabado con exito."
                 Progreso_Esperar True
                               
                 Ln_No = 1
                 FA.Servicio = 0
                 FA.Propina = 0
             End If
             
             For IdField = 0 To .fields.Count - 1
                 Codigo = Dato_Campo(.fields(IdField), , True)
                 Select Case IdField
                   Case 3: Precio = Redondear(Val(Codigo), Dec_PVP)
                   Case 4: Cantidad = Redondear(Val(Codigo), 2)
                   Case 5: SubTotalDescuento = Redondear(Val(Codigo), 2)
                   Case 6: SubTotalServicio = Redondear(Val(Codigo), 2)
                   Case 7: FA.Propina = FA.Propina + Redondear(Val(Codigo), 2)
                   Case 8: SubTotalIVA = Redondear(Val(Codigo), 2)
                   Case 9: TA.Banco = Codigo
                   Case 13: If Len(Codigo) <= 1 Then Producto = "VENTAS DEL DIA" Else Producto = Codigo
                   Case 14: If IsDate(Codigo) Then FA.Fecha_V = Codigo Else FA.Fecha_V = FechaSistema
                   Case 15: If Len(Codigo) <= 1 Then CodigoInv = "99.99" Else CodigoInv = Codigo
                   Case 16: Mes = Codigo
                   Case 17: TA.Cta = Codigo
                            If Len(TA.Banco) > 1 And Len(TA.Cta) > 1 Then
                               FA.T = Cancelado
                               TA.Abono = TA.Abono + FA.Total_MN
                            End If
                   Case 18: If Len(Codigo) > 1 Then TA.Comprobante = Codigo Else TA.Comprobante = Ninguno
                   Case 20: If Len(Codigo) > 1 Then FA.Cod_Ejec = Codigo Else FA.Cod_Ejec = Ninguno
                 End Select
             Next IdField
             SubTotal = Redondear(Cantidad * Precio, 2)
             
             SetAdoAddNew "Asiento_F"
             SetAdoFields "CODIGO", CodigoInv
             SetAdoFields "CANT", Cantidad
             SetAdoFields "PRECIO", Precio
             SetAdoFields "Total_Desc", SubTotalDescuento
             SetAdoFields "Total_IVA", SubTotalIVA
             SetAdoFields "SERVICIO", SubTotalServicio
             SetAdoFields "PRODUCTO", Producto
             SetAdoFields "TOTAL", SubTotal
             SetAdoFields "VALOR_TOTAL", SubTotal - SubTotalDescuento + SubTotalIVA
             SetAdoFields "CODIGO_L", FA.Cod_CxC
             If Len(Mes) >= 3 Then
                SetAdoFields "Mes", Mes
                SetAdoFields "TICKET", Year(FA.Fecha)
             End If
             SetAdoFields "Item", NumEmpresa
             SetAdoFields "CodigoU", CodigoUsuario
             SetAdoFields "Cod_Ejec", FA.Cod_Ejec
             SetAdoFields "A_No", CByte(Ln_No)
             SetAdoUpdate
             Ln_No = Ln_No + 1
          End If
          Progreso_Barra.Mensaje_Box = "Generando El Documento " & FA.TC & " No. " & Format$(FA.Factura, "000000000")
          Progreso_Esperar
         .MoveNext
      Loop
      Calculos_Totales_Factura FA
      If Len(FA.Autorizacion) > 13 Then FA.ClaveAcceso = FA.Autorizacion
      Grabar_Factura FA, True, True
      If TA.Abono > 0 Then
         TA.Fecha = FA.Fecha
         TA.Abono = FA.Total_MN
         TA.CodigoC = FA.CodigoC
         TA.TP = FA.TC
         TA.Serie = FA.Serie
         TA.Factura = FA.Factura
         TA.Autorizacion = FA.Autorizacion
         TA.Cta_CxP = FA.Cta_CxP
         Grabar_Abonos TA, True
      End If
   End If
  End With
  FA.Fecha_Corte = FechaSistema
  Actualizar_Abonos_Facturas_SP FA
  Progreso_Final
  DGExcelAdodc.Visible = True
End Sub

Public Sub Importar_Facturas_Farmacias()
Dim I As Long
Dim N As Long
Dim FacturaTemp As Long
Dim Tot_Propinas As Currency
Dim Precio2 As Currency
  
  Progreso_Iniciar
  DGExcelAdodc.Visible = False
  Encerar_Factura FA
  Iniciar_Asiento_Beneficiario
  Bandera = False
  Evaluar = True
  With AdoExcelAdodc.Recordset
   If .RecordCount > 0 Then
       Progreso_Barra.Valor_Maximo = (.RecordCount * 2) + 10
      .MoveFirst
       FA.Fecha = Dato_Campo(.fields(4), True)
       Mifecha = FA.Fecha
       FA.Serie = Ninguno
       FA.Autorizacion = Ninguno
       FA.Cod_CxC = DCLinea.Text
       Lineas_De_CxC FA
       SerieFactura = FA.Serie
       Fecha_Vence = FA.Vencimiento
       Autorizacion = FA.Autorizacion
       Cta_Cobrar = FA.Cta_CxP
       FA.Desde = Val(Dato_Campo(.fields(10)))
       FA.Hasta = FA.Desde

       RUC_CII = Dato_Campo(.fields(0))
       Beneficiario = Dato_Campo(.fields(1))
       RUC_CIF = Dato_Campo(.fields(2))
       NombreCliente = Dato_Campo(.fields(3))
       Insertar_Asiento_Beneficiario RUC_CII, Beneficiario
       Insertar_Asiento_Beneficiario RUC_CIF, NombreCliente
       Do While Not .EOF
          If RUC_CII <> Dato_Campo(.fields(0)) Then
            'MsgBox Beneficiario & vbCrLf & NombreCliente
             RUC_CII = Dato_Campo(.fields(0))
             Beneficiario = Dato_Campo(.fields(1))
             RUC_CIF = Dato_Campo(.fields(2))
             NombreCliente = Dato_Campo(.fields(3))
             Insertar_Asiento_Beneficiario RUC_CII, Beneficiario
             Insertar_Asiento_Beneficiario RUC_CIF, NombreCliente
          End If
          FA.Hasta = Val(Dato_Campo(.fields(10)))
          Progreso_Barra.Mensaje_Box = "Verificando: " & Beneficiario
          Progreso_Esperar
         .MoveNext
       Loop
       Control_Procesos "E", "Se eliminaron " & FA.TC & " No. " & FA.Serie & ": " & Format$(FA.Desde, "000000000") & " <-> " & Format$(FA.Hasta, "000000000")
       Progreso_Barra.Mensaje_Box = "Eliminando " & FA.TC & " desde la " & FA.Desde & " a la " & FA.Hasta
       Progreso_Esperar
       Actualiza_Procesado_Kardex_Rango_Factura FA
       Eliminar_Facturas
       Actualizar_Asiento_Beneficiario_Clientes True, True
   End If
  End With
  sSQL = "DELETE * " _
       & "FROM Asiento_F " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' "
  Ejecutar_SQL_SP sSQL
 'Empezamos la importacion de las facturas
  With AdoExcelAdodc.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      'MsgBox .Rows & vbCrLf & .Cols
       TA.Abono = 0
       FA.Servicio = 0
       FA.Propina = 0
       FA.Desde = Val(Dato_Campo(.fields(10)))
       FA.Factura = FA.Desde
       Factura_No = FA.Factura
       Codigo = Dato_Campo(.fields(4), True)
       If IsDate(Codigo) Then FA.Fecha = Codigo Else FA.Fecha = FechaSistema
       FA.Fecha_C = FA.Fecha
       FA.Autorizacion = Autorizacion
       If Len(FA.Autorizacion) > 13 Then FA.ClaveAcceso = FA.Autorizacion
       RUC_CII = Dato_Campo(.fields(0))
       RUC_CIF = Dato_Campo(.fields(2))
       CodigoCli = "9999999999"   'RUC/Cedula/Consumidor Final
       sSQL = "SELECT Codigo " _
            & "FROM Clientes " _
            & "WHERE CI_RUC = '" & RUC_CII & "' "
       Select_Adodc AdoAux, sSQL
       If AdoAux.Recordset.RecordCount > 0 Then CodigoCli = AdoAux.Recordset.fields("Codigo")
       FA.CodigoC = CodigoCli
       
       sSQL = "SELECT Codigo " _
            & "FROM Clientes " _
            & "WHERE CI_RUC = '" & RUC_CIF & "' "
       Select_Adodc AdoAux, sSQL
       If AdoAux.Recordset.RecordCount > 0 Then CodigoB = AdoAux.Recordset.fields("Codigo")
       Do While Not .EOF
'          If Not IsNull(.Fields(0)) And Not IsNull(.Fields(1)) Then
             Factura_No = Val(Dato_Campo(.fields(10)))
             If Factura_No <> FA.Factura Then
                Calculos_Totales_Factura FA
                If FA.Total_MN <= 0 Then FA.T = Anulado
                Grabar_Factura FA, True, True
                FA.Factura = Val(.fields(10))
                Codigo = Dato_Campo(.fields(4), True)
                If IsDate(Codigo) Then FA.Fecha = Codigo Else FA.Fecha = FechaSistema
                FA.Fecha_C = FA.Fecha
                RUC_CII = Dato_Campo(.fields(0))
                RUC_CIF = Dato_Campo(.fields(2))
                CodigoCli = "9999999999"   'RUC/Cedula/Consumidor Final
                sSQL = "SELECT Codigo " _
                     & "FROM Clientes " _
                     & "WHERE CI_RUC = '" & RUC_CII & "' "
                Select_Adodc AdoAux, sSQL
                If AdoAux.Recordset.RecordCount > 0 Then CodigoCli = AdoAux.Recordset.fields("Codigo")
                FA.CodigoC = CodigoCli
                
                sSQL = "SELECT Codigo " _
                     & "FROM Clientes " _
                     & "WHERE CI_RUC = '" & RUC_CIF & "' "
                Select_Adodc AdoAux, sSQL
                If AdoAux.Recordset.RecordCount > 0 Then CodigoB = AdoAux.Recordset.fields("Codigo")
             End If
             For IdField = 0 To .fields.Count - 1
                 Codigo = Dato_Campo(.fields(IdField))
                 If Codigo = "" Then Codigo = Ninguno
                 Codigo1 = Codigo
                 Select Case IdField + 1
                   Case 6: CodigoInv = Codigo
                   Case 8: Cantidad = Val(Codigo)
                   Case 9: Precio = Val(Codigo)
                   Case 10: Precio2 = Val(Codigo)
                 End Select
             Next IdField
             FA.SubTotal = Redondear(Cantidad * Precio, 2)
             sSQL = "SELECT Codigo_Inv, Producto, IVA " _
                  & "FROM Catalogo_Productos " _
                  & "WHERE Item = '" & NumEmpresa & "' " _
                  & "AND Periodo = '" & Periodo_Contable & "' " _
                  & "AND Codigo_IESS = '" & CodigoInv & "' "
             Select_Adodc AdoAux, sSQL
             If AdoAux.Recordset.RecordCount > 0 Then
                CodigoInv = AdoAux.Recordset.fields("Codigo_Inv")
                Producto = AdoAux.Recordset.fields("Producto")
                If AdoAux.Recordset.fields("IVA") Then FA.Total_IVA = Redondear(FA.SubTotal * Porc_IVA, 2)
             Else
                CodigoInv = "99.99"
                Producto = "VENTAS DEL DIA"
                FA.Total_IVA = 0
             End If
             SetAdoAddNew "Asiento_F"
             SetAdoFields "CODIGO", CodigoInv
             SetAdoFields "PRODUCTO", Producto
             SetAdoFields "CANT", Cantidad
             SetAdoFields "PRECIO", Precio
             SetAdoFields "TOTAL", FA.SubTotal
             SetAdoFields "Total_IVA", FA.Total_IVA
             SetAdoFields "Codigo_B", CodigoB
             SetAdoFields "PRECIO2", Precio2
             SetAdoFields "CodigoU", CodigoUsuario
             SetAdoFields "Item", NumEmpresa
             SetAdoUpdate
          Progreso_Barra.Mensaje_Box = "Importar de Excel a Sistema de Facturacion El Numero: " & Factura_No
          Progreso_Esperar
         .MoveNext
      Loop
      'FA.Factura = Factura_No
      FA.Hasta = FA.Factura
      Calculos_Totales_Factura FA
      If FA.Total_MN <= 0 Then FA.T = Anulado
      Grabar_Factura FA, True, True
   End If
  End With
  Actualiza_Procesado_Kardex_Rango_Factura FA
  Control_Procesos "G", "Se Grabaron " & FA.TC & " No. " & FA.Serie & ": " & Format$(FA.Desde, "000000000") & " <-> " & Format$(FA.Hasta, "000000000")
  Progreso_Final
  DGExcelAdodc.Visible = True
End Sub

Public Sub Importar_Plan_Cuentas()
Dim Clave As Long
Dim CodigoExt As String
Dim ComentarioDebe As String
Dim ComentarioHaber As String
  RatonReloj
  Parpadear = True
  Clave = 0
  With AdoExcelAdodc.Recordset
   If .RecordCount > 0 Then
       Progreso_Iniciar
       Progreso_Barra.Valor_Maximo = .RecordCount
       
       sSQL = "DELETE * " _
            & "FROM Catalogo_Cuentas " _
            & "WHERE Item = '" & NumEmpresa & "' " _
            & "AND Periodo = '" & Periodo_Contable & "' "
       Ejecutar_SQL_SP sSQL
       
      'Empezamos la importacion de las facturas
       Do While Not .EOF
          For IdField = 0 To .fields.Count - 1
              Codigo = Dato_Campo(.fields(IdField))
             ' MsgBox Codigo
              If Codigo = "" Then Codigo = Ninguno
              Select Case IdField + 1
                Case 1: TipoCta = Codigo     'TC
                Case 2: TipoDoc = Codigo     'DG
                Case 3: Codigo2 = Codigo     'Codigo Nuevo
                        If MidStrg(Codigo2, Len(Codigo2), 1) = "." Then Codigo2 = MidStrg(Codigo2, 1, Len(Codigo2) - 1)
                        If Codigo2 = "" Then Codigo2 = Ninguno
                Case 4: Cuenta = Codigo      'Cuenta
                Case 5: CodigoExt = Codigo      'Cuenta
                Case 6: ComentarioDebe = Codigo   'Comentario Dede
                Case 7: ComentarioHaber = Codigo   'Comentario Haber
                Case 8: TextoTraza = Codigo   'Comentario
              End Select
          Next IdField
         'Insertamos el Codigo nuevo
          'MsgBox Codigo2
          If Codigo2 <> Ninguno And Len(Codigo2) >= 1 Then
             Progreso_Barra.Mensaje_Box = "Migracion en Curso, Cuenta: " & Codigo2
             Progreso_Esperar
             SetAdoAddNew "Catalogo_Cuentas"
             SetAdoFields "TC", TipoCta
             SetAdoFields "DG", TipoDoc
             SetAdoFields "Codigo", Codigo2
             SetAdoFields "Codigo_Ext", CodigoExt
             SetAdoFields "Cuenta", TrimStrg(MidStrg(Cuenta, 1, 90))
             If TipoDoc = "D" Then
                Clave = Clave + 1
                SetAdoFields "Clave", Clave
             End If
             SetAdoFields "Periodo", Periodo_Contable
             SetAdoFields "Item", NumEmpresa
             SetAdoUpdate
          End If
         .MoveNext
       Loop
   End If
  End With
  RatonNormal
  Progreso_Final
End Sub

Public Sub Importar_Compras_Diarias()
Dim EsAlPasivo As Boolean
Dim CodRetBien As Byte
Dim CodRetServ As Byte
Dim I As Long
Dim N As Long
Dim A_No_SB As Long
Dim VInc As Long
Dim VMax As Long
Dim SecuencialF As Long
Dim SecuencialR As Long
Dim PorcRet As Single
Dim PorcIVAB As Single
Dim PorcIVAS As Single
Dim PorcIVABS As Single
Dim Valor As Currency
Dim TotalIVAB As Currency
Dim TotalIVAS As Currency
Dim Tot_Propinas As Currency
Dim SubTotalRet As Currency
Dim TotalRetFuente As Currency
Dim TotalRetIVABien As Currency
Dim TotalRetIVAServ As Currency

Dim DigCta As String
Dim CodMes As String
Dim SerieF1 As String
Dim SerieF2 As String
Dim SerieR1 As String
Dim SerieR2 As String
Dim CodigoTemp As String
Dim Cta_Gasto As String
Dim Cta_IVA_Gasto As String
Dim FormaPago As String
Dim SubModuloPago As String
Dim TipoSustento As String
Dim CtaRetFuente As String
Dim CtaRetIVABien As String
Dim CtaRetIVAServ As String
Dim ConceptoDiario As String
Dim Identificacion As String
Dim MifechaTemp As String
Dim NombreClienteTemp As String
Dim ErrorLinea As String

    Trans_No = 180
    A_No_SB = 1
    CodigoCC = Ninguno
    DGExcelAdodc.Visible = False
    DGAsiento.Visible = False
    
    Select Case CTP.Text
      Case "CE": NumComp = ReadSetDataNum("Egresos", True, True)
      Case "CI": NumComp = ReadSetDataNum("Ingresos", True, True)
      Case Else: NumComp = ReadSetDataNum("Diario", True, True)
                 CTP.Text = "CD"
    End Select
    Importar_Compras_Diarias_SP CTP.Text, NumComp
    DGAsiento.Visible = True
    DGExcelAdodc.Visible = True
    Progreso_Final
    MsgBox "Proceso realizado con exito, se han procesado" & vbCrLf & vbCrLf _
         & "Comprobantes de " & CTP.Text & " desde el numero: " & NumComp & " en adelante"
End Sub

Public Sub Eliminar_Comprobantes_Contabilidad(vTP As String, _
                                              vFechaIni As String, _
                                              vFechaFin As String, _
                                              vNo_Desde As Long, _
                                              vNo_Hasta As Long)
                                              
 If IsDate(vFechaIni) And IsDate(vFechaFin) And vNo_Desde > 0 And vNo_Hasta > 0 And Len(vTP) > 1 Then
    If vFechaIni <= vFechaFin And vNo_Desde <= vNo_Hasta Then
       sSQL = "DELETE * " _
            & "FROM Trans_SubCtas " _
            & "WHERE Item = '" & NumEmpresa & "' " _
            & "AND Periodo = '" & Periodo_Contable & "' " _
            & "AND Numero BETWEEN " & vNo_Desde & " AND " & vNo_Hasta & " " _
            & "AND Fecha BETWEEN #" & BuscarFecha(vFechaIni) & "# AND #" & BuscarFecha(vFechaFin) & "# " _
            & "AND TP = '" & vTP & "' "
       Ejecutar_SQL_SP sSQL
         
       sSQL = "DELETE * " _
            & "FROM Transacciones " _
            & "WHERE Item = '" & NumEmpresa & "' " _
            & "AND Periodo = '" & Periodo_Contable & "' " _
            & "AND Numero BETWEEN " & vNo_Desde & " AND " & vNo_Hasta & " " _
            & "AND Fecha BETWEEN #" & BuscarFecha(vFechaIni) & "# AND #" & BuscarFecha(vFechaFin) & "# " _
            & "AND TP = '" & vTP & "' "
       Ejecutar_SQL_SP sSQL
       
       sSQL = "DELETE * " _
            & "FROM Comprobantes " _
            & "WHERE Item = '" & NumEmpresa & "' " _
            & "AND Periodo = '" & Periodo_Contable & "' " _
            & "AND Numero BETWEEN " & vNo_Desde & " AND " & vNo_Hasta & " " _
            & "AND Fecha BETWEEN #" & BuscarFecha(vFechaIni) & "# AND #" & BuscarFecha(vFechaFin) & "# " _
            & "AND TP = '" & vTP & "' "
       Ejecutar_SQL_SP sSQL
      'MsgBox "Eliminar_Comprobantes_Contabilidad:" & vbCrLf & sSQL
    End If
 End If
End Sub

Public Sub Importar_Contabilidad()
Dim AdoCompDB As ADODB.Recordset
Dim CIRUC As String
Dim CodigoOld As String
Dim CodigoNew As String
    Progreso_Barra.Mensaje_Box = "Subiendo Contabilidad Externa"
    Progreso_Iniciar
    DGExcelAdodc.Visible = False
    TextoImprimio = ""
    
    Importar_Contabilidad_SP CTP
    
    sSQL = "SELECT Codigo, Cliente, CI_RUC, ID " _
         & "FROM Clientes " _
         & "WHERE Codigo LIKE '--%' " _
         & "ORDER BY Cliente "
    Select_AdoDB AdoCompDB, sSQL
    With AdoCompDB
     If .RecordCount > 0 Then
         Progreso_Barra.Valor_Maximo = .RecordCount
         Progreso_Barra.Incremento = 1
         Do While Not .EOF
            Progreso_Barra.Mensaje_Box = "Grabando las Transacciones del " & FechaTexto & ", CD No. " & Comp_No
            Progreso_Esperar
            ID_Trans = .fields("ID")
            CodigoOld = .fields("Codigo")
            CodigoNew = MidStrg(.fields("CI_RUC"), 1, 10)
            If UCase(GetUrlSource(urlEsUnRUC & .fields("CI_RUC"))) = "TRUE" Then
               SQL1 = "UPDATE Clientes " _
                    & "SET Codigo = '" & CodigoNew & "', TD = 'R' " _
                    & "WHERE Codigo = '" & CodigoOld & "' "
               Ejecutar_SQL_SP SQL1
               
               SQL1 = "UPDATE Comprobantes " _
                    & "SET Codigo_B = '" & CodigoNew & "' " _
                    & "WHERE Codigo_B = '" & CodigoOld & "' "
               Ejecutar_SQL_SP SQL1
            
               SQL1 = "UPDATE Transacciones " _
                    & "SET Codigo_C = '" & CodigoNew & "' " _
                    & "WHERE Codigo_C = '" & CodigoOld & "' "
               Ejecutar_SQL_SP SQL1
            Else
               
            End If
           .MoveNext
         Loop
     End If
    End With
    AdoCompDB.Close
    ConectarAdodc AdoExcelAdodc
    Select_Adodc AdoExcelAdodc, "SELECT * FROM Asiento_CSV_" & CodigoUsuario
    DGExcelAdodc.Visible = True
    Progreso_Final
    If Len(TextoImprimio) > 2 Then FInfoError.Show
End Sub

Public Sub Importar_Contabilidad_SubModulos()
Dim AdoCatalogoDB As ADODB.Recordset
Dim AdoSubCtaDB As ADODB.Recordset
Dim I As Long
Dim N As Long
Dim CtaD As String
Dim CtaH As String
Dim SerieF1 As String
Dim SerieF2 As String
Dim SecuencialF As Long
Dim SerieR1 As String
Dim SerieR2 As String
Dim FechaI As String
Dim FechaF As String
Dim SecuencialR As Long
Dim NumTrans As Long
Dim NumTransR As Long
Dim Tot_Propinas As Currency
Dim Cta_Gasto As String
Dim Fecha_Borrar As String
    
    Progreso_Barra.Mensaje_Box = "Subiendo Contabilidad Externa con SubModulos"
    Progreso_Iniciar
    RatonReloj
    DGExcelAdodc.Visible = False
    sSQL = "DELETE * " _
         & "FROM Tabla_Temporal " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Modulo = '" & NumModulo & "' " _
         & "AND CodigoU = '" & CodigoUsuario & "' "
    Ejecutar_SQL_SP sSQL
    
    TextoImprimio = ""
    Importar_Contabilidad_SubModulos_SP
    
    ConectarAdodc AdoExcelAdodc
    Select_Adodc AdoExcelAdodc, "SELECT * FROM Asiento_CSV_" & CodigoUsuario
    
    DGExcelAdodc.Visible = True
    RatonNormal
    Progreso_Final
    If Len(TextoImprimio) > 2 Then FInfoError.Show
End Sub

Public Sub Eliminar_Facturas()
 If FA.Desde <= FA.Hasta Then
   'MsgBox FA.TC & ": " & FA.Serie & " - " & FA.Desde & "<->" & FA.Hasta
    sSQL = "DELETE * " _
         & "FROM Facturas " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND TC = '" & FA.TC & "' " _
         & "AND Serie = '" & FA.Serie & "' " _
         & "AND Factura BETWEEN " & FA.Desde & " and " & FA.Hasta & " "
        ' & "AND Autorizacion = '" & FA.Autorizacion & "' "
    Ejecutar_SQL_SP sSQL

    sSQL = "DELETE * " _
         & "FROM Detalle_Factura " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND TC = '" & FA.TC & "' " _
         & "AND Serie = '" & FA.Serie & "' " _
         & "AND Factura BETWEEN " & FA.Desde & " and " & FA.Hasta & " "
         '& "AND Autorizacion = '" & FA.Autorizacion & "' "
    Ejecutar_SQL_SP sSQL
    
    sSQL = "DELETE * " _
         & "FROM Trans_Abonos " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND TP = '" & FA.TC & "' " _
         & "AND Serie = '" & FA.Serie & "' " _
         & "AND Factura BETWEEN " & FA.Desde & " and " & FA.Hasta & " "
         '& "AND Autorizacion = '" & FA.Autorizacion & "' "
    Ejecutar_SQL_SP sSQL
    
    sSQL = "DELETE * " _
         & "FROM Trans_Kardex " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND TC = '" & FA.TC & "' " _
         & "AND Serie = '" & FA.Serie & "' " _
         & "AND Factura BETWEEN " & FA.Desde & " and " & FA.Hasta & " "
    Ejecutar_SQL_SP sSQL
    Control_Procesos "G", "Grabar " & FA.TC & " No. " & FA.Serie & " Desde " & Format$(FA.Desde, "000000000") & " - " & Format$(FA.Hasta, "000000000") & " [" & FA.Hora & "]"
 End If
End Sub

Public Sub Actualizar_Facturas(Nom_Tabla As String, Mi_Fecha As String)
   sSQL = "UPDATE " & Nom_Tabla & " " _
        & "SET T = 'C' " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND Periodo = '" & Periodo_Contable & "' " _
        & "AND Fecha = #" & BuscarFecha(Mi_Fecha) & "# " _
        & "AND T <> 'C' "
   Ejecutar_SQL_SP sSQL
End Sub

Public Sub InsValorCta(NCta As String, _
                       NValor As Currency)
  For IE = 0 To ContCtas - 1
      If CtasProc(IE).Cta = NCta Then
         CtasProc(IE).Valor = CtasProc(IE).Valor + Redondear(NValor, 2)
      End If
  Next IE
End Sub

Public Sub SetearCtasCierre(CtaFields As String)
  Si_No = True
  For IE = 0 To ContCtas - 1
      If CtaFields = CtasProc(IE).Cta Then Si_No = False
  Next IE
  If CtaFields = Ninguno Then Si_No = False
  If Si_No Then
     IE = 0
     While IE < ContCtas
        If CtasProc(IE).Cta = "0" Then
           CtasProc(IE).Cta = CtaFields
           IE = ContCtas + 1
        End If
        IE = IE + 1
     Wend
  End If
End Sub

Public Sub Generar_Facturas()
    TotalIngreso = 0
    Cta_Banco = Ninguno
    If Len(Cta_Banco) <= 1 Then Cta_Banco = Cta_Del_Banco
    RatonReloj
    DGExcelAdodc.Visible = False
    FA.Cod_CxC = DCLinea.Text
    Lineas_De_CxC FA
    Contador = 0
    sSQL = "SELECT * " _
         & "FROM Asiento_F " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND CodigoU = '" & CodigoUsuario & "' " _
         & "ORDER BY Numero "
    Select_Adodc AdoAct, sSQL
    With AdoAct.Recordset
     If .RecordCount > 0 Then
         Do While Not .EOF
            Factura_No = .fields("Numero")
            FA.Factura = Factura_No
            sSQL = "DELETE * " _
                 & "FROM Facturas " _
                 & "WHERE Periodo = '" & Periodo_Contable & "' " _
                 & "AND Item = '" & NumEmpresa & "' " _
                 & "AND Factura = " & Factura_No & " " _
                 & "AND Autorizacion = '" & FA.Autorizacion & "' " _
                 & "AND Serie = '" & FA.Serie & "' " _
                 & "AND TC = '" & FA.TC & "' "
            Ejecutar_SQL_SP sSQL
            
            sSQL = "DELETE * " _
                 & "FROM Detalle_Factura " _
                 & "WHERE Periodo = '" & Periodo_Contable & "' " _
                 & "AND Item = '" & NumEmpresa & "' " _
                 & "AND Factura = " & Factura_No & " " _
                 & "AND Autorizacion = '" & FA.Autorizacion & "' " _
                 & "AND Serie = '" & FA.Serie & "' " _
                 & "AND TC = '" & FA.TC & "' "
            Ejecutar_SQL_SP sSQL
            
            sSQL = "DELETE * " _
                 & "FROM Trans_Abonos " _
                 & "WHERE Periodo = '" & Periodo_Contable & "' " _
                 & "AND Item = '" & NumEmpresa & "' " _
                 & "AND Factura = " & Factura_No & " " _
                 & "AND Autorizacion = '" & FA.Autorizacion & "' " _
                 & "AND Serie = '" & FA.Serie & "' " _
                 & "AND TP = '" & FA.TC & "' "
            Ejecutar_SQL_SP sSQL
           .MoveNext
         Loop
        .MoveFirst
         Do While Not .EOF
            Factura_No = .fields("Numero")
            FechaTexto = .fields("FECHA")
            CodigoCli = .fields("Codigo_Cliente")
            NombreCliente = .fields("RUTA")
            NoMeses = .fields("A_No")
            Codigo = .fields("CODIGO")
            Codigo2 = .fields("HABIT")
            SubCta = .fields("Cta")
            SubTotal = .fields("PRECIO")
            SubTotal_IVA = .fields("Total_IVA")
            Producto = .fields("PRODUCTO")
            FA.Factura = Factura_No
            SetAdoAddNew "Detalle_Factura"
            SetAdoFields "T", FA.T
            SetAdoFields "TC", FA.TC
            SetAdoFields "Factura", Factura_No
            SetAdoFields "CodigoC", CodigoCli
            SetAdoFields "Fecha", FechaTexto
            SetAdoFields "Codigo", Codigo
            SetAdoFields "Cantidad", 1
            SetAdoFields "CodigoL", FA.Cod_CxC
            SetAdoFields "Precio", SubTotal
            SetAdoFields "Total", SubTotal
            SetAdoFields "Total_IVA", SubTotal_IVA
            SetAdoFields "Producto", Producto
            SetAdoFields "CodigoU", CodigoUsuario
            SetAdoFields "Periodo", Periodo_Contable
            SetAdoFields "Item", NumEmpresa
            SetAdoFields "Autorizacion", FA.Autorizacion
            SetAdoFields "Serie", FA.Serie
            SetAdoUpdate
            Contador = Contador + 1
            Me.Caption = "Generando el Detalle de " & FA.TC & " No. " & FA.Factura & ": " & Format$(Contador / .RecordCount, "00%")
           .MoveNext
         Loop
     End If
    End With
    Contador = 0
    sSQL = "SELECT Codigo_Cliente,FECHA,Numero,SUM(PRECIO) As TSubTotal,SUM(Total_IVA) As TTotal_IVA " _
         & "FROM Asiento_F " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND CodigoU = '" & CodigoUsuario & "' " _
         & "GROUP BY Codigo_Cliente,FECHA,Numero " _
         & "ORDER BY Codigo_Cliente,FECHA,Numero "
    Select_Adodc AdoAct, sSQL
    With AdoAct.Recordset
     If .RecordCount > 0 Then
         Do While Not .EOF
            Factura_No = .fields("Numero")
            Total_IVA = .fields("TTotal_IVA")
            Total_Con_IVA = Redondear(Total_IVA / 0.12, 2)
            Total_Sin_IVA = .fields("TSubTotal") - Total_Con_IVA
            Total_Servicio = 0
            Total_Desc = 0
            FechaTexto = .fields("FECHA")
            CodigoCli = .fields("Codigo_Cliente")
            NombreCliente = Ninguno
            FA.Factura = Factura_No
            Total_Factura = Redondear(Total_Sin_IVA + Total_Con_IVA - Total_Desc + Total_IVA + Total_Servicio, 2)
            If Redondear(Total_Sin_IVA + Total_Con_IVA, 2) > 0 Then
               Saldo = Total_Factura
               TotalIngreso = TotalIngreso + Total_Factura
              'Grabamos las Facturas
               SetAdoAddNew "Facturas"
               SetAdoFields "C", adFalse
               SetAdoFields "T", FA.T
               SetAdoFields "TC", FA.TC
               SetAdoFields "Factura", Factura_No
               SetAdoFields "Fecha", FechaTexto
               SetAdoFields "Fecha_C", FechaTexto
               SetAdoFields "Fecha_V", FechaTexto
               SetAdoFields "CodigoC", CodigoCli
               SetAdoFields "Sin_IVA", Total_Sin_IVA
               SetAdoFields "Con_IVA", Total_Con_IVA
               SetAdoFields "SubTotal", Total_Sin_IVA + Total_Con_IVA
               SetAdoFields "IVA", Total_IVA
               SetAdoFields "Total_MN", Total_Factura
               SetAdoFields "Saldo_MN", Total_Factura
               SetAdoFields "Nota", "Factura importada por execel "
               SetAdoFields "Cod_CxC", FA.Cod_CxC
               SetAdoFields "Cta_CxP", FA.Cta_CxP
               SetAdoFields "Cta_Venta", FA.Cta_Venta
               SetAdoFields "CodigoU", CodigoUsuario
               SetAdoFields "Periodo", Periodo_Contable
               SetAdoFields "Item", NumEmpresa
               SetAdoFields "Hora", Format$(Time, FormatoTimes)
               SetAdoFields "Vencimiento", FA.Vencimiento
               SetAdoFields "Autorizacion", FA.Autorizacion
               SetAdoFields "Serie", FA.Serie
               SetAdoUpdate

'''             SetAdoAddNew "Trans_Abonos"
'''             SetAdoFields "T", Cancelado
'''             SetAdoFields "TP", FA.TC
'''             SetAdoFields "CodigoC", CodigoCli
'''             SetAdoFields "Fecha", FechaTexto
'''             SetAdoFields "Comprobante", Val(SubCta)
'''             SetAdoFields "Recibo_No", TrimStrg(SubCta)
'''             SetAdoFields "Factura", Factura_No
'''             SetAdoFields "Abono", Total_Factura
'''             SetAdoFields "Banco", "DEPOSITO POR BANCO"
'''             SetAdoFields "Cheque", Grupo_No
'''             SetAdoFields "Cta", Cta_Banco
'''             SetAdoFields "Cta_CxP", Cta_Cobrar
'''             SetAdoFields "Serie", FA.Serie
'''             SetAdoFields "Autorizacion", FA.Autorizacion
'''             SetAdoUpdate
           End If
           Factura_No = Factura_No + 1
           Control_Procesos Normal, "Grabar Factura No. " & FA.Serie & "-" & Format$(Factura_No, "0000000")
           Contador = Contador + 1
           Me.Caption = "Generando Encabezado de " & FA.TC & " No. " & FA.Factura & ": " & Format$(Contador / .RecordCount, "00%")
          .MoveNext
         Loop
     End If
    End With
    RatonNormal
   DGExcelAdodc.Visible = True
   RatonNormal
End Sub

Public Sub Migrar_Cta_Nueva(Tabla As String, Campo As String, Cod_Antiguo As String, Cod_Nuevo As String)
    If Tabla <> "" And Campo <> "" Then
       If Cod_Antiguo = "" Then Cod_Antiguo = Ninguno
       If Cod_Nuevo = "" Then Cod_Nuevo = Ninguno
       If Cod_Antiguo <> Ninguno And Cod_Nuevo <> Ninguno Then
          sSQL = "UPDATE " & Tabla & " " _
               & "SET " & Campo & " = '" & Cod_Nuevo & "' " _
               & "WHERE Item = '" & NumEmpresa & "' " _
               & "AND Periodo = '" & Periodo_Contable & "' " _
               & "AND " & Campo & " = '" & Cod_Antiguo & "' "
          Ejecutar_SQL_SP sSQL
       End If
    End If
End Sub

Public Sub Importar_Abonos()
Dim I As Long
Dim N As Long
Dim CodRet As String
Dim CompRet As String
Dim Tot_Propinas As Currency
 
    Progreso_Barra.Mensaje_Box = "Subiendo Abonos diarios"
    Progreso_Iniciar
 
    Encerar_Factura FA
  
    sSQL = "UPDATE Facturas " _
         & "SET X = '.' " _
         & "WHERE Periodo = '" & Periodo_Contable & "' " _
         & "AND Item = '" & NumEmpresa & "' " _
         & "AND TC NOT IN ('C','P') " _
         & "AND X <> '.' " _
         & "AND T <> 'A' "
    Ejecutar_SQL_SP sSQL
  
    FechaTexto = FechaSistema
    Bandera = False
    Evaluar = True
    DiarioCaja = ReadSetDataNum("Recibo_No", True, True)
    
   'Empezamos la importacion de las facturas
    TA.Cta_CxP = FA.Cta_CxP
    TA.Autorizacion = FA.Autorizacion
    TA.Serie = FA.Serie
    TA.TP = FA.TC
    TA.T = Normal
    With AdoExcelAdodc.Recordset
     If .RecordCount > 0 Then
         Progreso_Barra.Valor_Maximo = .RecordCount
        .MoveFirst
         Do While Not .EOF
            For IdField = 0 To .fields.Count - 1
                Select Case IdField + 1
                  Case 1: TA.Fecha = Dato_Campo(.fields(IdField), True)
                  Case 2: Codigo = Dato_Campo(.fields(IdField))
                          TA.Serie = MidStrg(Codigo, 1, 3) & MidStrg(Codigo, 4, 3)
                          TA.Factura = Val(MidStrg(Codigo, 7, Len(Codigo)))
                  Case 3: TA.Autorizacion = Dato_Campo(.fields(IdField))
                  Case 4: TA.AutorizacionR = Dato_Campo(.fields(IdField))
                  Case 5: TA.Abono = Redondear(Val(Dato_Campo(.fields(IdField))), 2)
                  Case 6: Codigo = Dato_Campo(.fields(IdField))
                          TA.Serie_R = MidStrg(Codigo, 1, 3) & MidStrg(Codigo, 4, 3)
                          TA.Secuencial_R = Val(MidStrg(Codigo, 7, Len(Codigo)))
                          
                          TA.Establecimiento = MidStrg(Codigo, 1, 3)
                          TA.Emision = MidStrg(Codigo, 4, 3)
                          CompRet = Val(MidStrg(Codigo, 7, Len(Codigo)))
                  Case 7: CodRet = Dato_Campo(.fields(IdField))
                  Case 8: TA.Cheque = Dato_Campo(.fields(IdField))
                          TA.Banco = Dato_Campo(.fields(IdField))
                  Case 9: TA.Cta = Dato_Campo(.fields(IdField))
                End Select
            Next IdField
            If Len(TA.Fecha) >= 10 And Len(Dato_Campo(.fields(9))) >= 2 Then
               sSQL = "DELETE * " _
                    & "FROM Trans_Abonos " _
                    & "WHERE Item = '" & NumEmpresa & "' " _
                    & "AND Periodo = '" & Periodo_Contable & "' " _
                    & "AND TP = '" & TA.TP & "' " _
                    & "AND Serie = '" & TA.Serie & "' " _
                    & "AND Factura = " & TA.Factura & " " _
                    & "AND Abono = " & TA.Abono & " "
               If Len(TA.Autorizacion) > 1 Then sSQL = sSQL & "AND Autorizacion = '" & TA.Autorizacion & "' "
               Ejecutar_SQL_SP sSQL
               
               TA.Porcentaje = 0
               If Len(CodRet) > 1 And Len(TA.Serie_R) = 6 Then
                  sSQL = "SELECT Porcentaje " _
                       & "FROM Tipo_Concepto_Retencion " _
                       & "WHERE Codigo = '" & CodRet & "' " _
                       & "AND Fecha_Inicio <= #" & BuscarFecha(TA.Fecha) & "# " _
                       & "AND Fecha_Final >= #" & BuscarFecha(TA.Fecha) & "# "
                  Select_Adodc AdoAux, sSQL
                  If AdoAux.Recordset.RecordCount > 0 Then TA.Porcentaje = AdoAux.Recordset.fields("Porcentaje")
               End If
               
               If TA.Banco = "" Then TA.Banco = Ninguno
               TA.CodigoC = Ninguno
               sSQL = "SELECT * " _
                    & "FROM Facturas " _
                    & "WHERE Item = '" & NumEmpresa & "' " _
                    & "AND Periodo = '" & Periodo_Contable & "' " _
                    & "AND TC = '" & TA.TP & "' " _
                    & "AND Serie = '" & TA.Serie & "' " _
                    & "AND Factura = " & TA.Factura & " "
               If Len(TA.Autorizacion) > 1 Then sSQL = sSQL & "AND Autorizacion = '" & TA.Autorizacion & "' "
               sSQL = sSQL & "ORDER BY Autorizacion "
               Select_Adodc AdoAux, sSQL
               
'            MsgBox TA.TP & vbCrLf & TA.Serie & vbCrLf & TA.Factura & vbCrLf & vbCrLf & TA.Serie_R & vbCrLf & TA.AutorizacionR & vbCrLf & TA.Secuencial_R & vbCrLf & CodRet & vbCrLf & TA.Porcentaje
               
               If AdoAux.Recordset.RecordCount > 0 Then
                  TA.CodigoC = AdoAux.Recordset.fields("CodigoC")
                  TA.Autorizacion = AdoAux.Recordset.fields("Autorizacion")
                  TA.Cta_CxP = AdoAux.Recordset.fields("Cta_CxP")
                  Select Case UCaseStrg(Dato_Campo(.fields(9)))
                    Case "RIB"
                         TA.Banco = "RETENCION IVA BIENES"
                         TA.Cheque = CompRet
                         Grabar_Abonos TA
                    Case "RIS"
                         TA.Banco = "RETENCION IVA SERVICIO"
                         TA.Cheque = CompRet
                         Grabar_Abonos TA
                    Case "RF"
                         TA.Banco = "RETENCION FUENTE - " & CodRet
                         TA.Cheque = CompRet
                         Grabar_Abonos TA
                    Case "EFE"
                         TA.Banco = "EFECTIVO MN"
                         Grabar_Abonos TA
                    Case "BCO"
                         TA.Banco = "DEPOSITO EN EFECTIVO"
                         Grabar_Abonos TA
                    Case Else
                         TA.Banco = "OTROS TIPOS ABONOS"
                         Grabar_Abonos TA
                  End Select
               End If
            End If
            Progreso_Barra.Mensaje_Box = "Importando Abonos de " & FA.TC & ": " & FA.Serie & "-" & FA.Factura
            Progreso_Esperar
           .MoveNext
         Loop
     End If
    End With
    FA.Factura = 0
    FA.Fecha_Corte = FechaSistema
    Actualizar_Abonos_Facturas_SP FA
    Progreso_Final
End Sub

Public Sub Importar_Estudiantes_Representantes()
Dim I As Long
Dim N As Long
Dim Cl As Tipo_Beneficiarios
Dim Tot_Propinas As Currency
    Progreso_Barra.Mensaje_Box = "Subiendo Abonos diarios"
    Progreso_Iniciar
    TextoImprimio = ""
    With AdoExcelAdodc.Recordset
     If .RecordCount > 0 Then
         Progreso_Barra.Valor_Maximo = .RecordCount
        .MoveFirst
         Do While Not .EOF
            For IdField = 0 To .fields.Count - 1
                If IdField = 2 Then Codigo = Dato_Campo(.fields(IdField), True) Else Codigo = Dato_Campo(.fields(IdField), , True)
                Codigo = Sin_Signos_Especiales(Codigo)
                Select Case IdField + 1
                  Case 1: Cl.T = Codigo               ' T
                  Case 2: Cl.Codigo = Codigo          ' Codigo
                  Case 3: Cl.Fecha_N = Codigo         ' Fecha_N
                  Case 4: Cl.Cliente = Codigo         ' Cliente
                  Case 5: Cl.Sexo = Codigo            ' Sexo
                  Case 6: Cl.Email1 = LCase$(Codigo)  ' Email
                  Case 7: Cl.Email2 = LCase$(Codigo)  ' Email2
                  Case 8: Cl.Direccion = Codigo       ' Descripcion del Curso
                  Case 9: Cl.Telefono1 = Codigo       ' Telefono
                  Case 10: Cl.Celular = Codigo        ' Celular
                  Case 11: Cl.Ciudad = Codigo         ' Ciudad
                  Case 12: Cl.Prov = Codigo           ' Prov
                  Case 13: Cl.DirNumero = Codigo      ' Dir Numerico
                  Case 14: Cl.Representante = Codigo  ' Representante
                  Case 15: Cl.RUC_CI_Rep = Codigo     ' Cedula del Representante
                  Case 16: Cl.Direccion_Rep = Codigo  ' Direccion del Representante
                  Case 17: Cl.Grupo_No = Codigo       ' Curso avreviado
                End Select
            Next IdField
            Cl.EmailR = Cl.Email1
            Cl.CI_RUC = Cl.Codigo
           
            If Len(Cl.Codigo) <= 6 Then Cl.Codigo = NumEmpresa & Format$(Val(Cl.Codigo), "000000")
            DigVerif = Digito_Verificador(Cl.Codigo)
            Cl.TD = Tipo_RUC_CI.Tipo_Beneficiario
            Cl.Codigo = Tipo_RUC_CI.Codigo_RUC_CI
            Cl.CI_RUC = Tipo_RUC_CI.RUC_CI
         
            DigVerif = Digito_Verificador(Cl.RUC_CI_Rep)
            Cl.TD_Rep = Tipo_RUC_CI.Tipo_Beneficiario
    
            sSQL = "SELECT Codigo " _
                 & "FROM Clientes " _
                 & "WHERE Cliente = '" & Cl.Cliente & "' "
            Select_Adodc AdoAux, sSQL
            
            If AdoAux.Recordset.RecordCount > 0 Then
               Cl.Codigo = AdoAux.Recordset.fields("Codigo")
            Else
               SetAdoAddNew "Clientes"
               SetAdoFields "T", Cl.T
               SetAdoFields "Codigo", Cl.Codigo
               SetAdoFields "CI_RUC", Cl.CI_RUC
               SetAdoFields "Fecha_N", Cl.Fecha_N
               SetAdoFields "Cliente", Cl.Cliente
               SetAdoFields "Sexo", Cl.Sexo
               SetAdoFields "Email", Cl.Email1
               SetAdoFields "Email2", Cl.Email2
               SetAdoFields "Direccion", Cl.Direccion
               SetAdoFields "Telefono", Cl.Telefono1
               SetAdoFields "Celular", Cl.Celular
               SetAdoFields "Ciudad", Cl.Ciudad
               SetAdoFields "Prov", Cl.Prov
               SetAdoFields "DirNumero", Cl.DirNumero
               SetAdoFields "Grupo", Cl.Grupo_No
               SetAdoFields "FA", True
               SetAdoFields "TD", Cl.TD
               SetAdoUpdate
               TextoImprimio = TextoImprimio & Cl.Cliente & vbCrLf
            End If
               
            sSQL = "SELECT Representante " _
                 & "FROM Clientes_Matriculas " _
                 & "WHERE Codigo = '" & Cl.Codigo & "' " _
                 & "AND Item = '" & NumEmpresa & "' " _
                 & "AND Periodo = '" & Periodo_Contable & "' "
            Select_Adodc AdoAux, sSQL
            If AdoAux.Recordset.RecordCount <= 0 Then
               SetAdoAddNew "Clientes_Matriculas"
               SetAdoFields "T", Cl.T
               SetAdoFields "Codigo", Cl.Codigo
               SetAdoFields "TD", Cl.TD_Rep
               SetAdoFields "Representante", Cl.Representante
               SetAdoFields "Representante_Alumno", Cl.Representante
               SetAdoFields "Cedula_R", Cl.RUC_CI_Rep
               SetAdoFields "Grupo_No", Cl.Grupo_No
               SetAdoFields "Lugar_Trabajo_R", Cl.Direccion_Rep
               SetAdoFields "Item", NumEmpresa
               SetAdoFields "Periodo", Periodo_Contable
               SetAdoUpdate
            End If
            Progreso_Barra.Mensaje_Box = "Importando Estudiante " & Cl.Cliente
            Progreso_Esperar
           .MoveNext
         Loop
     End If
    End With
    Progreso_Final
    If Len(TextoImprimio) > 1 Then
       TextoImprimio = "CLIENTES NUEVOS: " & vbCrLf & TextoImprimio & vbCrLf
       FInfoError.Show 1
    End If
End Sub

Public Sub Importar_Estudiantes_PreFacturas()
Dim I As Long
Dim N As Long
Dim Cl As Tipo_Beneficiarios
Dim Tot_Propinas As Currency
    Progreso_Barra.Mensaje_Box = "Subiendo Prefacturas mensuales"
    Progreso_Iniciar
    TextoImprimio = ""
    With AdoExcelAdodc.Recordset
     If .RecordCount > 0 Then
         Progreso_Barra.Valor_Maximo = .RecordCount
        .MoveFirst
         Do While Not .EOF
            For IdField = 0 To .fields.Count - 1
                If IdField = 1 Then Codigo = Dato_Campo(.fields(IdField), True) Else Codigo = Dato_Campo(.fields(IdField))
                Codigo = UCaseStrg(Codigo)
                Codigo = Sin_Signos_Especiales(Codigo)
                Select Case IdField + 1
                  Case 1: CodigoCli = Codigo          ' CI_RUC
                  Case 2: Mifecha = Codigo           ' Fecha
                  Case 3: CodigoInv = Codigo         ' Codigo_Inv
                  Case 4: NoMes = Val(Codigo)              ' Mes
                  Case 5: Anio = Codigo               ' Año
                  Case 6: Total = Val(Codigo)         ' Valor
                  Case 7: Total_IVA = Val(Codigo)     ' IVA
                  Case 8: Total_Desc = Val(Codigo)          ' Descuento1
                  Case 9: Total_Desc2 = Val(Codigo)         ' Descuento2
                  Case 10: ValorTotal = Val(Codigo)             ' Total
                  Case 11: Cl.Cliente = Codigo        ' Cliente
                End Select
            Next IdField
            Mes = MesesLetras(NoMes)
            SetAdoAddNew "Clientes_Facturacion"
            SetAdoFields "T", Normal
            SetAdoFields "Codigo", CodigoCli
            SetAdoFields "Codigo_Inv", CodigoInv
            SetAdoFields "Valor", Total
            SetAdoFields "Periodo", Anio
            SetAdoFields "Num_Mes", NoMes
            SetAdoFields "Mes", Mes
            SetAdoFields "Fecha", Mifecha
            SetAdoFields "Descuento", Total_Desc
            SetAdoFields "Descuento2", Total_Desc2
            SetAdoUpdate
            Progreso_Barra.Mensaje_Box = "Importando Estudiante " & Cl.Cliente
            Progreso_Esperar
           .MoveNext
         Loop
     End If
    End With
    Progreso_Final
    If Len(TextoImprimio) > 1 Then
       TextoImprimio = "CLIENTES NUEVOS: " & vbCrLf & TextoImprimio & vbCrLf
       FInfoError.Show 1
    End If
End Sub

Public Sub Importar_Personas()
Dim I As Long
Dim N As Long
Dim Lista_Clientes_Nuevos As String
Dim Crear_Nuevo As Boolean
    Lista_Clientes_Nuevos = ""
    Progreso_Barra.Mensaje_Box = "Subiendo Beneficiarios"
    Progreso_Iniciar
    TextoImprimio = ""
    With AdoExcelAdodc.Recordset
     If .RecordCount > 0 Then
         Progreso_Barra.Valor_Maximo = .RecordCount
        .MoveFirst
         Do While Not .EOF
            Datos_Default_Beneficiario
            Crear_Nuevo = False
            Codigo = Dato_Campo(.fields(1))
           'MsgBox Codigo
           'RUC/Cedula/Codigo Alumno/Consumidor Final
            If Len(Codigo) > 1 Then
               If IsNumeric(Codigo) Then
                  If Len(Codigo) < 9 Then
                     Codigo = Format$(Val(Codigo), "00000000")
                     TBeneficiario.FA = True
                  ElseIf Len(Codigo) = 9 Then
                     Codigo = "0" & Codigo
                  ElseIf Len(Codigo) = 11 Then
                     Codigo = "00" & Codigo
                  ElseIf Len(Codigo) = 12 Then
                     Codigo = "0" & Codigo
                  End If
               End If
               TBeneficiario.CI_RUC = Codigo
               DigVerif = Digito_Verificador(Codigo)
               Caracter = MidStrg(Codigo, 10, 1)
               TBeneficiario.TD_Rep = Tipo_RUC_CI.Tipo_Beneficiario
               TBeneficiario.TP = Tipo_RUC_CI.Tipo_Beneficiario
               TBeneficiario.Codigo = Tipo_RUC_CI.Codigo_RUC_CI
            End If
            If TBeneficiario.Codigo <> Ninguno Then
               For IdField = 0 To .fields.Count - 1
                   Codigo = Dato_Campo(.fields(IdField))
                   Select Case IdField
                     Case 0: TBeneficiario.T = Codigo       'T
                    'Case 1: YA ESTA ARRIBA DETERMINADO
                     Case 2: TBeneficiario.Fecha_N = Codigo 'Fecha_N
                     Case 3: TBeneficiario.Cliente = UCaseStrg(TrimStrg(MidStrg(Codigo, 1, 60)))  'Cliente
                     Case 4: TBeneficiario.Sexo = MidStrg(Codigo, 1, 1) 'Sexo"
                     Case 5: TBeneficiario.Email1 = TrimStrg(MidStrg(Codigo, 1, 50)) 'Email
                     Case 6: TBeneficiario.Email2 = TrimStrg(MidStrg(Codigo, 1, 50)) 'Email2
                     Case 7: TBeneficiario.Direccion = UCaseStrg(MidStrg(Codigo, 1, 50)) 'Direccion
                     Case 8: TBeneficiario.Telefono1 = MidStrg(Codigo, 1, 10) 'Telefono
                     Case 9: TBeneficiario.Celular = MidStrg(Codigo, 1, 10) 'Celular
                     Case 10: TBeneficiario.Ciudad = UCaseStrg(Codigo) 'Ciudad
                     Case 11: TBeneficiario.Prov = MidStrg(Codigo, 1, 2) 'Prov
                     Case 12: TBeneficiario.DirNumero = MidStrg(Codigo, 1, 8) 'DirNumero
                     Case 13: TBeneficiario.Grupo_No = UCaseStrg(TrimStrg(MidStrg(Codigo, 1, 9)))
                     Case 14: TBeneficiario.Cod_Ejec = TrimStrg(MidStrg(Codigo, 1, 10))
                     Case 15: TBeneficiario.Cta_CxP = TrimStrg(MidStrg(Codigo, 1, 16))
                     Case 16: TBeneficiario.Plan_Afiliado = Format(Val(Codigo), "0000")
                   End Select
               Next IdField
               'MsgBox TBeneficiario.CI_RUC
               If Len(TBeneficiario.CI_RUC) > 1 Then
                  sSQL = "SELECT * " _
                       & "FROM Clientes " _
                       & "WHERE CI_RUC = '" & TBeneficiario.CI_RUC & "' "
                  Select_Adodc AdoClientes, sSQL
                  If AdoClientes.Recordset.RecordCount > 0 Then
                     If Len(TBeneficiario.T) > 0 Then AdoClientes.Recordset.fields("T") = TBeneficiario.T
                     If Len(TBeneficiario.TD_Rep) > 0 Then AdoClientes.Recordset.fields("TD") = TBeneficiario.TD_Rep
                     If IsDate(TBeneficiario.Fecha_N) > 1 Then AdoClientes.Recordset.fields("Fecha_N") = TBeneficiario.Fecha_N
                     If Len(TBeneficiario.Cliente) > 1 Then AdoClientes.Recordset.fields("Cliente") = TBeneficiario.Cliente
                     If Len(TBeneficiario.Sexo) > 1 Then AdoClientes.Recordset.fields("Sexo") = TBeneficiario.Sexo
                     If Len(TBeneficiario.Email1) > 1 Then AdoClientes.Recordset.fields("Email") = TBeneficiario.Email1
                     If Len(TBeneficiario.Email2) > 1 Then AdoClientes.Recordset.fields("Email2") = TBeneficiario.Email2
                     If Len(TBeneficiario.Direccion) > 1 Then AdoClientes.Recordset.fields("Direccion") = TBeneficiario.Direccion
                     If Len(TBeneficiario.DirNumero) > 1 Then AdoClientes.Recordset.fields("DirNumero") = TBeneficiario.DirNumero
                     If Len(TBeneficiario.Telefono1) > 1 Then AdoClientes.Recordset.fields("Telefono") = TBeneficiario.Telefono1
                     If Len(TBeneficiario.Celular) > 1 Then AdoClientes.Recordset.fields("Celular") = TBeneficiario.Celular
                     If Len(TBeneficiario.Ciudad) > 1 Then AdoClientes.Recordset.fields("Ciudad") = MidStrg(TBeneficiario.Ciudad, 1, 35)
                     If Len(TBeneficiario.Prov) > 1 Then AdoClientes.Recordset.fields("Prov") = TBeneficiario.Prov
                     If Len(TBeneficiario.Pais) > 1 Then AdoClientes.Recordset.fields("Pais") = TBeneficiario.Pais
                     If Len(TBeneficiario.Grupo_No) > 1 Then AdoClientes.Recordset.fields("Grupo") = TBeneficiario.Grupo_No
                     If Len(TBeneficiario.Cod_Ejec) > 1 Then AdoClientes.Recordset.fields("Cod_Ejec") = TBeneficiario.Cod_Ejec
                     If Len(TBeneficiario.Cta_CxP) > 1 Then AdoClientes.Recordset.fields("Cta_CxP") = TBeneficiario.Cta_CxP
                     If Len(TBeneficiario.Plan_Afiliado) > 1 Then AdoClientes.Recordset.fields("Plan_Afiliado") = TBeneficiario.Plan_Afiliado
                     AdoClientes.Recordset.fields("FA") = TBeneficiario.FA
                     AdoClientes.Recordset.Update
                  Else
                     Crear_Nuevo = True
                  End If
               End If
              '
               If Crear_Nuevo Then
                  SetAdoAddNew "Clientes"
                  If Len(TBeneficiario.T) > 0 Then SetAdoFields "T", TBeneficiario.T
                  If Len(TBeneficiario.TD_Rep) > 0 Then SetAdoFields "TD", TBeneficiario.TD_Rep
                  If Len(TBeneficiario.Codigo) > 1 Then SetAdoFields "Codigo", TBeneficiario.Codigo
                  If Len(TBeneficiario.CI_RUC) > 1 Then SetAdoFields "CI_RUC", TBeneficiario.CI_RUC
                  If IsDate(TBeneficiario.Fecha) > 1 Then SetAdoFields "Fecha", TBeneficiario.Fecha
                  If IsDate(TBeneficiario.Fecha_N) > 1 Then SetAdoFields "Fecha_N", TBeneficiario.Fecha_N
                  If Len(TBeneficiario.Cliente) > 1 Then SetAdoFields "Cliente", TBeneficiario.Cliente
                  If Len(TBeneficiario.Sexo) > 1 Then SetAdoFields "Sexo", TBeneficiario.Sexo
                  If Len(TBeneficiario.Email1) > 1 Then SetAdoFields "Email", TBeneficiario.Email1
                  If Len(TBeneficiario.Email2) > 1 Then SetAdoFields "Email2", TBeneficiario.Email2
                  If Len(TBeneficiario.Direccion) > 1 Then SetAdoFields "Direccion", TBeneficiario.Direccion
                  If Len(TBeneficiario.DirNumero) > 1 Then SetAdoFields "DirNumero", TBeneficiario.DirNumero
                  If Len(TBeneficiario.Telefono1) > 1 Then SetAdoFields "Telefono", TBeneficiario.Telefono1
                  If Len(TBeneficiario.Celular) > 1 Then SetAdoFields "Celular", TBeneficiario.Celular
                  If Len(TBeneficiario.Ciudad) > 1 Then SetAdoFields "Ciudad", TBeneficiario.Ciudad
                  If Len(TBeneficiario.Prov) > 1 Then SetAdoFields "Prov", TBeneficiario.Prov
                  If Len(TBeneficiario.Grupo_No) > 1 Then SetAdoFields "Grupo", TBeneficiario.Grupo_No
                  If Len(TBeneficiario.Cod_Ejec) > 1 Then SetAdoFields "Cod_Ejec", TBeneficiario.Cod_Ejec
                  If Len(TBeneficiario.Cta_CxP) > 1 Then SetAdoFields "Cta_CxP", TBeneficiario.Cta_CxP
                  If Len(TBeneficiario.Plan_Afiliado) > 1 Then SetAdoFields "Plan_Afiliado", TBeneficiario.Plan_Afiliado
                  SetAdoUpdate
               End If
            End If
            Progreso_Barra.Mensaje_Box = "Importando El Beneficiario: " & TBeneficiario.CI_RUC
            Progreso_Esperar
           .MoveNext
         Loop
     End If
    End With
    Progreso_Final
End Sub

Public Sub Importar_Descuento_Empleados()
Dim I As Long
Dim N As Long
Dim Lista_Clientes_Nuevos As String
Dim CodigoPT As String
Dim Novedad As String
Dim NoMesT As Integer
Dim Crear_Nuevo As Boolean
Dim Aplica_FP As Boolean

    Progreso_Barra.Mensaje_Box = "Subiendo Abonos diarios"
    Progreso_Iniciar
    TextoImprimio = ""
    
    Lista_Clientes_Nuevos = ""
  
    sSQL = "SELECT Codigo, CI_RUC, Cliente " _
         & "FROM Clientes " _
         & "WHERE Codigo <> '.' " _
         & "ORDER BY Codigo "
    Select_Adodc AdoClientes, sSQL
    
    With AdoExcelAdodc.Recordset
     If .RecordCount > 0 Then
         Progreso_Barra.Valor_Maximo = .RecordCount * 2
        .MoveFirst
         'Cuenta = Dato_Campo(.fields(4))
         NoMes = Dato_Campo(.fields(5))
         Codigo = Dato_Campo(.fields(3))
         Cta = Leer_Cta_Catalogo(Codigo)
         CodigoP = CodRolPago ' Rubro_Rol_Pago(Cuenta)
         CodigoPT = Codigo
         Do While Not .EOF
            NoMesT = Dato_Campo(.fields(5))
            'CodigoPT = Rubro_Rol_Pago(Cuenta)
            NombreCliente = Dato_Campo(.fields(1))
            CodigoPT = Dato_Campo(.fields(3))
            If NoMesT <> NoMes Or CodigoPT <> Codigo Then
               sSQL = "DELETE * " _
                    & "FROM Catalogo_Rol_Rubros " _
                    & "WHERE Item = '" & NumEmpresa & "' " _
                    & "AND Periodo = '" & Periodo_Contable & "' " _
                    & "AND Mes = " & NoMes & " " _
                    & "AND Cod_Rol_Pago = '" & CodigoP & "' "
               Ejecutar_SQL_SP sSQL
               NoMes = NoMesT
               Codigo = Dato_Campo(.fields(3))
               Cta = Leer_Cta_Catalogo(Codigo)
               CodigoP = CodRolPago ' Rubro_Rol_Pago(Cuenta)
               'CodigoP = CodigoPT
            End If
            Progreso_Barra.Mensaje_Box = "Verificando rubros del Empleado: " & NombreCliente
            Progreso_Esperar
           .MoveNext
         Loop
         sSQL = "DELETE * " _
              & "FROM Catalogo_Rol_Rubros " _
              & "WHERE Item = '" & NumEmpresa & "' " _
              & "AND Periodo = '" & Periodo_Contable & "' " _
              & "AND Mes = " & NoMes & " " _
              & "AND Cod_Rol_Pago = '" & CodigoP & "' "
         Ejecutar_SQL_SP sSQL
         
        .MoveFirst
         Do While Not .EOF
            NombreCliente = Dato_Campo(.fields(1))
            Codigo = Dato_Campo(.fields(0))
            CICliente = Ninguno
            If AdoClientes.Recordset.RecordCount > 0 Then
               AdoClientes.Recordset.MoveFirst
               AdoClientes.Recordset.Find ("CI_RUC = '" & Codigo & "' ")
               If Not AdoClientes.Recordset.EOF Then CICliente = AdoClientes.Recordset.fields("Codigo")
            End If
            Codigo = Dato_Campo(.fields(2))
            Valor = Val(Codigo)
            If Valor > 0 And CICliente <> Ninguno Then
               Crear_Nuevo = False
               For IdField = 0 To .fields.Count - 1
                   Codigo = Dato_Campo(.fields(IdField))
                   Select Case IdField + 1
                     Case 2: CodigoCliente = Codigo
                     Case 4: Cta = Leer_Cta_Catalogo(Codigo)
                     Case 5: 'Cuenta = Codigo
                             CodigoP = CodRolPago   'Rubro_Rol_Pago(Cuenta)
                     Case 6: NoMes = Val(Codigo)
                     Case 7: TipoDoc = UCaseStrg(Codigo)
                     Case 8: Novedad = Codigo
                   End Select
               Next IdField
              'Creamos los Clientes Rol de Pagos
               sSQL = "SELECT Valor, Calc_IESS " _
                    & "FROM Catalogo_Rol_Rubros " _
                    & "WHERE Item = '" & NumEmpresa & "' " _
                    & "AND Periodo = '" & Periodo_Contable & "' " _
                    & "AND Codigo = '" & CICliente & "' " _
                    & "AND Mes = " & NoMes & " " _
                    & "AND Cod_Rol_Pago = '" & CodigoP & "' " _
                    & "ORDER BY Codigo,Cod_Rol_Pago "
               Select_Adodc AdoAux, sSQL
               If AdoAux.Recordset.RecordCount > 0 Then
                  AdoAux.Recordset.fields("Valor") = Valor
                  Select Case TipoDoc
                    Case "I": AdoAux.Recordset.fields("Calc_IESS") = adTrue
                    Case "IS": AdoAux.Recordset.fields("Calc_IESS") = adFalse
                    Case "EC": AdoAux.Recordset.fields("Calc_IESS") = adTrue
                    Case Else: AdoAux.Recordset.fields("Calc_IESS") = adFalse
                  End Select
                  AdoAux.Recordset.Update
               Else
                  SetAdoAddNew "Catalogo_Rol_Rubros"
                  SetAdoFields "Codigo", CICliente
                  SetAdoFields "I_E", MidStrg(TipoDoc, 1, 1)
                  SetAdoFields "Mes", NoMes
                  SetAdoFields "Detalle", Cuenta
                  SetAdoFields "Cta", Cta
                  SetAdoFields "TV", "V"
                  SetAdoFields "Valor", Redondear(Valor, 2)
                  SetAdoFields "CPais", "593"
                  SetAdoFields "Cod_Rol_Pago", CodigoP
                  Select Case TipoDoc
                    Case "I": SetAdoFields "Calc_IESS", adTrue
                    Case "IS": SetAdoFields "Calc_IESS", adFalse
                    Case "EC": SetAdoFields "Calc_IESS", adTrue
                    Case Else: SetAdoFields "Calc_IESS", adFalse
                  End Select
                  SetAdoUpdate
               End If
               
               Mifecha = UltimoDiaMes("01/" & Format(NoMes, "00") & "/" & Year(FechaSistema))
               sSQL = "DELETE * " _
                    & "FROM Trans_Entrada_Salida " _
                    & "WHERE Item = '" & NumEmpresa & "' " _
                    & "AND Periodo = '" & Periodo_Contable & "' " _
                    & "AND Fecha = #" & BuscarFecha(Mifecha) & "# " _
                    & "AND Codigo = '" & CICliente & "' "
               Ejecutar_SQL_SP sSQL
               If Len(Novedad) > 1 Then
                  SetAdoAddNew "Trans_Entrada_Salida"
                  SetAdoFields "ES", "R"
                  SetAdoFields "Codigo", CICliente
                  SetAdoFields "Hora", Format(Time, FormatoTimes)
                  SetAdoFields "Fecha", Mifecha
                  SetAdoFields "Proceso", "NOVEDADES"
                  SetAdoFields "Tarea", TrimStrg(MidStrg(Novedad, 1, 50))
                  SetAdoFields "CodigoU", CodigoUsuario
                  SetAdoFields "Periodo", Periodo_Contable
                  SetAdoFields "Item", NumEmpresa
                  SetAdoUpdate
               End If
               
            End If
            If CICliente = Ninguno Then Lista_Clientes_Nuevos = Lista_Clientes_Nuevos & CICliente & vbTab & CodigoCliente & vbCrLf
            Progreso_Barra.Mensaje_Box = "Importando El Beneficiario: " & NombreCliente & ", en Rol Pagos"
            Progreso_Esperar
           .MoveNext
         Loop
     End If
    End With
    Progreso_Final
    If Len(Lista_Clientes_Nuevos) > 2 Then
       TextoImprimio = Lista_Clientes_Nuevos
       Unload FImporta
      'FInfoError.Show
    End If
End Sub

Public Sub Insertar_Ventas(CodigoProv As String, _
                           CantidadFacturas As Long, _
                           TotalCero As Currency, _
                           TotalGravada As Currency, _
                           TotalIVA As Currency, _
                           SubTotalFA As Currency, _
                           TipoDocumento As String)
Dim BaseImpIVAB As Currency
Dim BaseImpIVAS As Currency
Dim PorcIVA_B As Single
Dim PorcIVA_S As Single
  'If (TotalCero + TotalGravada + SubTotalFA) > 0 Then
    'MsgBox Total_Con_IVA & vbCrLf & Total_Sin_IVA & vbCrLf & TotalIVA
     
     BaseImpIVAB = 0
     BaseImpIVAS = 0
     PorcIVA_B = 0
     PorcIVA_S = 0
     If Val(TipoDocumento) <> 4 Then
        If PorcIVAB > 0 Then BaseImpIVAB = Redondear((Total_RetIVA * 100) / PorcIVAB, 2)
        If PorcIVAS > 0 Then BaseImpIVAS = Redondear((Total_RetIVA * 100) / PorcIVAS, 2)
        Select Case PorcIVAB
          Case 1 To 30: PorcIVA_B = 1
          Case 31 To 100: PorcIVA_B = 3
        End Select
        Select Case PorcIVAS
          Case 1 To 70: PorcIVA_S = 2
          Case 71 To 100: PorcIVA_S = 3
        End Select
     End If
     SetAdoAddNew "Trans_Ventas"
     SetAdoFields "IdProv", CodigoProv
     SetAdoFields "TipoComprobante", TipoDocumento
     SetAdoFields "FechaRegistro", FechaFinal
     SetAdoFields "FechaEmision", FechaFinal
     SetAdoFields "Establecimiento", Codigo1
     SetAdoFields "PuntoEmision", Codigo2
     SetAdoFields "NumeroComprobantes", CantidadFacturas
     SetAdoFields "Autorizacion", Autorizacion
     SetAdoFields "BaseImponible", TotalCero
     SetAdoFields "BaseImpGrav", TotalGravada
     SetAdoFields "MontoIva", TotalIVA
     SetAdoFields "IvaPresuntivo", "N"
     SetAdoFields "PorcentajeIva", 2
     SetAdoFields "RetPresuntiva", "S"
     SetAdoFields "MontoIvaBienes", BaseImpIVAB
     SetAdoFields "PorRetBienes", PorcIVA_B
     SetAdoFields "ValorRetBienes", Total_RetIVAB
     SetAdoFields "Porc_Bienes", PorcIVAB
     SetAdoFields "MontoIvaServicios", BaseImpIVAS
     SetAdoFields "PorRetServicios", PorcIVA_S
     SetAdoFields "ValorRetServicios", Total_RetIVAS
     SetAdoFields "Porc_Servicios", PorcIVAS
     SetAdoFields "Linea_SRI", 0
     SetAdoFields "T", Normal
     SetAdoFields "TP", "CD"
     SetAdoFields "Numero", Numero
     SetAdoFields "Fecha", FechaFinal
    ''SetAdoFields "ID", NumTrans
     SetAdoUpdate
  'End If
End Sub

Public Sub Insertar_Ventas_Air(CodigoProv As String, _
                               BaseImponible As Currency, _
                               PorcentajeRet As Single, _
                               ValorRet As Currency, _
                               FacturaNo As Long, _
                               RetencionNo As Long, _
                               EstabRet As String, _
                               PuntoRet As String, _
                               AutorRet As String, _
                               CuentRet As String)
  If ValorRet > 0 Then
    'PorcentajeRet = PorcentajeRet / 100
     SetAdoAddNew "Trans_Air"
     SetAdoFields "T", Normal
     SetAdoFields "TP", "CD"
     SetAdoFields "Numero", Numero
     SetAdoFields "Fecha", FechaFinal
     SetAdoFields "IdProv", CodigoProv
     SetAdoFields "CodRet", TipoDoc
     SetAdoFields "Tipo_Trans", "V"
     SetAdoFields "BaseImp", BaseImponible
     SetAdoFields "Porcentaje", PorcentajeRet
     SetAdoFields "ValRet", ValorRet
     SetAdoFields "EstabRetencion", EstabRet
     SetAdoFields "PtoEmiRetencion", PuntoRet
     SetAdoFields "SecRetencion", RetencionNo
     SetAdoFields "AutRetencion", AutorRet
     SetAdoFields "EstabFactura", Codigo1
     SetAdoFields "PuntoEmiFactura", Codigo2
     SetAdoFields "Factura_No", FacturaNo
     SetAdoFields "Cta_Retencion", CuentRet
    ''SetAdoFields "ID", NumTrans
     SetAdoFields "Linea_SRI", 0
     SetAdoUpdate
  End If
End Sub

Public Sub Importar_Retenciones_Farmacia()
Dim NumTrans As Long
Dim NumTransR As Long
Dim SecuencialF As Long
Dim SecuencialR As Long

Dim Tot_Propinas As Currency
Dim Total_Ret_1 As Currency
Dim Total_Ret_30 As Currency

Dim Porcentaje As Single

Dim SerieF1 As String
Dim SerieF2 As String
Dim SerieR1 As String
Dim SerieR2 As String
Dim Cta_Gasto As String
Dim Cta_Ret_1 As String
Dim Cta_Ret_30 As String
Dim FechaCodAir As String
Dim Detalle_Ret As String

    Progreso_Barra.Mensaje_Box = "Subiendo Abonos diarios"
    Progreso_Iniciar
    TextoImprimio = ""
    
    Cta_Ret_30 = Leer_Seteos_Ctas("Cta_Ret_IVA_30")
   'Encerar_Facturas
    sSQL = "DELETE * " _
         & "FROM Trans_Compras " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND Numero = -1 " _
         & "AND TP = 'NN' "
    Ejecutar_SQL_SP sSQL
    
    sSQL = "DELETE * " _
         & "FROM Trans_Air " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND Numero = -1 " _
         & "AND TP = 'NN' "
    Ejecutar_SQL_SP sSQL
    Eliminar_Asientos_SP True
    IniciarAsientosDe DGAsiento, AdoAsiento
    Bandera = False
    Evaluar = True
        
    With AdoExcelAdodc.Recordset
     If .RecordCount > 0 Then
         Progreso_Barra.Valor_Maximo = .RecordCount
        .MoveFirst
        'Nombre del Proveedor
         NombreCliente = Dato_Campo(.fields(3))
         Codigo = Dato_Campo(.fields(4))
        'RUC/Cedula/Consumidor Final
         CodigoCli = "9999999999"
         If Len(Codigo) > 1 Then
            CI_Representante = Codigo
            sSQL = "SELECT Codigo, Cliente, CI_RUC " _
                 & "FROM Clientes " _
                 & "WHERE CI_RUC = '" & Codigo & "' "
            Select_Adodc AdoAux, sSQL
            If AdoAux.Recordset.RecordCount > 0 Then CodigoCli = AdoAux.Recordset.fields("Codigo")
         End If
        .MoveFirst
         Do While Not .EOF
            For IdField = 0 To .fields.Count - 1
                Codigo = Dato_Campo(.fields(IdField))
                Select Case IdField + 1
                  Case 2: If Codigo = "FA" Then TipoDoc = "01" Else TipoDoc = "02"
                  Case 3: Mifecha = Codigo
                          FechaTexto = Mifecha                     'Caducidad de la Factura
                          FechaCodAir = BuscarFecha(Mifecha)
                  Case 6: SerieF1 = MidStrg(Codigo, 1, 3)
                          SerieF2 = MidStrg(Codigo, 4, 3)
                  Case 7: Autorizacion = Codigo
                  Case 8: SecuencialF = Val(Codigo)
                  Case 9: Total_Sin_No_IVA = 0
                          Total_Con_IVA = Redondear(Val(Codigo), 2)
                  Case 10: Total_Sin_IVA = Redondear(Val(Codigo), 2)
                  Case 12: Total_IVA = Redondear(Val(Codigo), 2)
                  Case 13: Total = Redondear(Val(Codigo), 2)
                  Case 14: SerieR1 = MidStrg(Codigo, 1, 3)
                           SerieR2 = MidStrg(Codigo, 4, 3)
                  Case 15: SecuencialR = Val(Codigo)
                  Case 16: AutorizaRet = Codigo
                  Case 17: CodigoP = Codigo
                  Case 19: Cta_Gasto = Codigo
                End Select
            Next IdField
            Detalle_Ret = "Cta_Ret_1"
            Cta_Ret_1 = Leer_Seteos_Ctas(Detalle_Ret)
            Porcentaje = 0
            sSQL = "SELECT Porcentaje " _
                 & "FROM Tipo_Concepto_Retencion " _
                 & "WHERE Codigo = '" & CodigoP & "' " _
                 & "AND Fecha_Inicio <= #" & FechaCodAir & "# " _
                 & "AND Fecha_Final >= #" & FechaCodAir & "# "
            Select_Adodc AdoAux, sSQL
            If AdoAux.Recordset.RecordCount > 0 Then Porcentaje = AdoAux.Recordset.fields("Porcentaje")
            Detalle_Ret = "Cta_Ret_" & Format(Porcentaje, "#0.00")
            Cta_Ret_1 = Leer_Seteos_Ctas(Detalle_Ret)
            Porcentaje = Porcentaje / 100
            SubTotal = Total_Con_IVA + Total_Sin_IVA + Total_Sin_No_IVA
            If SubTotal > 0 Then Total_Ret_1 = Redondear(SubTotal * Porcentaje, 2) Else Total_Ret_1 = 0
            If Total_IVA > 0 Then Total_Ret_30 = Redondear(Total_IVA * 0.3, 2) Else Total_Ret_30 = 0
           'MsgBox Cta_Ret_1 & vbCrLf & Detalle_Ret & vbCrLf & SubTotal & vbCrLf & Total_Ret_1

           'MsgBox NombreCliente
            If IsDate(Mifecha) Then
'''              sSQL = "DELETE * " _
'''                   & "FROM Trans_Compras " _
'''                   & "WHERE Item = '" & NumEmpresa & "' " _
'''                   & "AND Periodo = '" & Periodo_Contable & "' " _
'''                   & "AND TP = 'NN' " _
'''                   & "AND Numero = -1 " _
'''                   & "AND FechaEmision = #" & BuscarFecha(Mifecha) & "# " _
'''                   & "AND IdProv = '" & CodigoCli & "' " _
'''                   & "AND Establecimiento = '" & SerieF1 & "' " _
'''                   & "AND PuntoEmision = '" & SerieF2 & "' " _
'''                   & "AND Secuencial = " & SecuencialF & " " _
'''                   & "AND TipoComprobante = " & Val(TipoDoc) & " " _
'''                   & "AND Autorizacion = '" & Autorizacion & "' "
'''              Ejecutar_SQL_SP sSQL
'''              sSQL = "DELETE * " _
'''                   & "FROM Trans_Air " _
'''                   & "WHERE Item = '" & NumEmpresa & "' " _
'''                   & "AND Periodo = '" & Periodo_Contable & "' " _
'''                   & "AND TP = 'NN' " _
'''                   & "AND Numero = -1 " _
'''                   & "AND IdProv = '" & CodigoCli & "' " _
'''                   & "AND EstabRetencion = '" & SerieR1 & "' " _
'''                   & "AND PtoEmiRetencion = '" & SerieR2 & "' " _
'''                   & "AND SecRetencion = " & SecuencialR & " " _
'''                   & "AND AutRetencion = '" & AutorizaRet & "' " _
'''                   & "AND EstabFactura = '" & SerieF1 & "' " _
'''                   & "AND PuntoEmiFactura = '" & SerieF2 & "' " _
'''                   & "AND Factura_No = " & SecuencialF & " "
'''              Ejecutar_SQL_SP sSQL
             'Empezamos a grabar los datos de la retencion
             'MsgBox Mifecha & vbCrLf & FechaTexto
               SetAdoAddNew "Trans_Compras"
               SetAdoFields "IdProv", CodigoCli
               SetAdoFields "DevIva", "N"
               SetAdoFields "CodSustento", "02"
               SetAdoFields "TipoComprobante", Val(TipoDoc)
               SetAdoFields "Establecimiento", SerieF1
               SetAdoFields "PuntoEmision", SerieF2
               SetAdoFields "Secuencial", SecuencialF
               SetAdoFields "Autorizacion", Autorizacion
               SetAdoFields "FechaEmision", Mifecha
               SetAdoFields "FechaRegistro", Mifecha
               SetAdoFields "FechaCaducidad", FechaTexto
               If Total_IVA > 0 Then
                  SetAdoFields "BaseImpGrav", Total_Con_IVA
                  SetAdoFields "MontoIva", Total_IVA
                  SetAdoFields "PorcentajeIva", 2
               End If
               SetAdoFields "BaseImponible", Total_Sin_IVA
               SetAdoFields "BaseNoObjIVA", Total_Sin_No_IVA
               If Total_Ret_30 > 0 Then
                  SetAdoFields "MontoIvaBienes", Total_IVA
                  SetAdoFields "PorRetBienes", 1
                  SetAdoFields "ValorRetBienes", Total_Ret_30
                  SetAdoFields "Porc_Bienes", "30"
                  SetAdoFields "Cta_Bienes", Cta_Ret_30
               End If
               SetAdoFields "Cta_Pago", Cta_CxP_Retenciones   'Cta_CajaG
               SetAdoFields "Cta_Gasto", Cta_Gasto
               SetAdoFields "PagoLocExt", "01"
               SetAdoFields "PaisEfecPago", "NA"
               SetAdoFields "AplicConvDobTrib", "NA"
               SetAdoFields "PagExtSujRetNorLeg", "NA"
               SetAdoFields "FormaPago", "01"
               SetAdoFields "Serie_Retencion", SerieR1 & SerieR2
               SetAdoFields "SecRetencion", SecuencialR
               SetAdoFields "AutRetencion", AutorizaRet
               SetAdoFields "Linea_SRI", 0
               SetAdoFields "FechaEmiModificado", "000"
               SetAdoFields "EstabModificado", "000"
               SetAdoFields "PtoEmiModificado", "000"
               SetAdoFields "SecModificado", "000"
               SetAdoFields "AutModificado", "000"
               SetAdoFields "T", Normal
               SetAdoFields "TP", "NN"
               SetAdoFields "Numero", -1
               SetAdoFields "Fecha", Mifecha
               SetAdoUpdate
              'MsgBox Total_Sin_IVA & vbCrLf & Total_Con_IVA & vbCrLf & NombreCliente
               NumTrans = NumTrans + 1
               
              'RETENCION EN LA FUENTE
              'MsgBox CodigoP & vbCrLf & Porcentaje
               Total_Ret = Total_Ret_1
               SetAdoAddNew "Trans_Air"
               SetAdoFields "CodRet", CodigoP
               SetAdoFields "BaseImp", SubTotal
               SetAdoFields "ValRet", Total_Ret
               SetAdoFields "EstabRetencion", SerieR1
               SetAdoFields "PtoEmiRetencion", SerieR2
               SetAdoFields "SecRetencion", SecuencialR
               SetAdoFields "AutRetencion", AutorizaRet
               SetAdoFields "Tipo_Trans", "C"
               SetAdoFields "IdProv", CodigoCli
               If Total_Ret_1 > 0 Then
                  SetAdoFields "Cta_Retencion", Cta_Ret_1
                  SetAdoFields "Porcentaje", Porcentaje
               End If
               SetAdoFields "EstabFactura", SerieF1
               SetAdoFields "PuntoEmiFactura", SerieF2
               SetAdoFields "Factura_No", SecuencialF
               SetAdoFields "Linea_SRI", 0
               SetAdoFields "T", Normal
               SetAdoFields "TP", "NN"
               SetAdoFields "Numero", -1
               SetAdoFields "Fecha", Mifecha
               SetAdoUpdate
               NumTransR = NumTransR + 1
            End If
            Progreso_Barra.Mensaje_Box = "Importando: Fecha = " & Mifecha & ", Proveedor: " & NombreCliente
            Progreso_Esperar
           .MoveNext
         Loop
     End If
    End With
    Progreso_Final
End Sub

Public Sub Generar_Asiento_Compras(Optional ParaFarmacia As Boolean)
Dim I As Integer
Dim J As Integer
Dim ContSC As Integer
Dim Fecha_Sem As String
Dim Cta_Gasto As String

   DGAsiento.Visible = False
   sSQL = "SELECT Fecha " _
        & "FROM Trans_Compras " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND Periodo = '" & Periodo_Contable & "' " _
        & "AND Numero = -1 " _
        & "AND TP = 'NN' " _
        & "ORDER BY Fecha "
   Select_Adodc AdoAsiento, sSQL
   With AdoAsiento.Recordset
    If .RecordCount > 0 Then
        FechaIni = .fields("Fecha")
       .MoveLast
        FechaFin = .fields("Fecha")
    End If
   End With
   sSQL = "SELECT Cta_Gasto,Cta_Pago,Cta_Servicio,Cta_Bienes " _
        & "FROM Trans_Compras " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND Periodo = '" & Periodo_Contable & "' " _
        & "AND Numero = -1 " _
        & "AND TP = 'NN' " _
        & "GROUP BY Cta_Gasto,Cta_Pago,Cta_Servicio,Cta_Bienes " _
        & "ORDER BY Cta_Gasto,Cta_Pago,Cta_Servicio,Cta_Bienes "
   Select_Adodc AdoAsiento, sSQL
   ContCtas = AdoAsiento.Recordset.RecordCount * 4

   sSQL = "SELECT Cta_Retencion " _
        & "FROM Trans_Air " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND Periodo = '" & Periodo_Contable & "' " _
        & "AND Numero = -1 " _
        & "AND TP = 'NN' " _
        & "GROUP BY Cta_Retencion " _
        & "ORDER BY Cta_Retencion "
   Select_Adodc AdoAux, sSQL
   ContCtas = ContCtas + AdoAux.Recordset.RecordCount + 20
   ReDim CtasProc(ContCtas) As CtasAsiento
  'Procedemos a organizar las cuentas involucradas
   For IE = 0 To ContCtas - 1
       CtasProc(IE).Cta = "0"
       CtasProc(IE).Valor = 0
   Next IE
   With AdoAsiento.Recordset
    If .RecordCount > 0 Then
        Do While Not .EOF
           SetearCtasCierre .fields("Cta_Gasto")
          .MoveNext
        Loop
    End If
   End With
   SetearCtasCierre Cta_IVA_Inventario
   With AdoAux.Recordset
    If .RecordCount > 0 Then
        Do While Not .EOF
           SetearCtasCierre .fields("Cta_Retencion")
          .MoveNext
        Loop
    End If
   End With
   With AdoAsiento.Recordset
    If .RecordCount > 0 Then
       .MoveFirst
        Do While Not .EOF
           SetearCtasCierre .fields("Cta_Servicio")
           SetearCtasCierre .fields("Cta_Bienes")
          .MoveNext
        Loop
    End If
   End With
   With AdoAsiento.Recordset
    If .RecordCount > 0 Then
       .MoveFirst
        Do While Not .EOF
           SetearCtasCierre .fields("Cta_Pago")
          .MoveNext
        Loop
    End If
   End With

   sSQL = "SELECT CodSustento, Cta_Gasto, Cta_Pago, Cta_Servicio, Cta_Bienes, Establecimiento, PuntoEmision, Secuencial, " _
        & "SUM(BaseImponible+BaseNoObjIVA+BaseImpGrav) As TGasto, SUM(MontoIva) As TMontoIva, " _
        & "SUM(ValorRetBienes) As TValorRetBienes, SUM(ValorRetServicios) As TValorRetServicios " _
        & "FROM Trans_Compras " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND Periodo = '" & Periodo_Contable & "' " _
        & "AND Numero = -1 " _
        & "AND TP = 'NN' " _
        & "GROUP BY CodSustento,Cta_Gasto,Cta_Pago,Cta_Servicio,Cta_Bienes,Establecimiento,PuntoEmision,Secuencial " _
        & "ORDER BY CodSustento,Cta_Gasto,Cta_Pago,Cta_Servicio,Cta_Bienes,Establecimiento,PuntoEmision,Secuencial "
   Select_Adodc AdoAsiento, sSQL
   With AdoAsiento.Recordset
    If .RecordCount > 0 Then
        Do While Not .EOF
           'MsgBox "...."
           InsValorCta .fields("Cta_Pago"), .fields("TValorRetServicios")
           InsValorCta .fields("Cta_Pago"), .fields("TValorRetBienes")
           InsValorCta .fields("Cta_Pago"), -.fields("TGasto")
           InsValorCta .fields("Cta_Pago"), -.fields("TMontoIva")
           If .fields("CodSustento") = "01" Then
               InsValorCta .fields("Cta_Gasto"), .fields("TGasto")
               InsValorCta Cta_IVA_Inventario, .fields("TMontoIva")
           Else
               InsValorCta .fields("Cta_Gasto"), .fields("TGasto") + .fields("TMontoIva")
           End If
           InsValorCta .fields("Cta_Servicio"), -.fields("TValorRetServicios")
           InsValorCta .fields("Cta_Bienes"), -.fields("TValorRetBienes")

           sSQL = "SELECT Cta_Retencion,SUM(ValRet) As TValRet " _
                & "FROM Trans_Air " _
                & "WHERE Item = '" & NumEmpresa & "' " _
                & "AND Periodo = '" & Periodo_Contable & "' " _
                & "AND EstabFactura = '" & .fields("Establecimiento") & "' " _
                & "AND PuntoEmiFactura = '" & .fields("PuntoEmision") & "' " _
                & "AND Factura_No = " & .fields("Secuencial") & " " _
                & "AND Numero = -1 " _
                & "AND TP = 'NN' " _
                & "GROUP BY Cta_Retencion " _
                & "ORDER BY Cta_Retencion "
           Select_Adodc AdoAux, sSQL
           If AdoAux.Recordset.RecordCount > 0 Then
              Do While Not AdoAux.Recordset.EOF
                 'MsgBox "...."
                 InsValorCta AdoAux.Recordset.fields("Cta_Retencion"), -AdoAux.Recordset.fields("TValRet")
                 InsValorCta .fields("Cta_Pago"), AdoAux.Recordset.fields("TValRet")
                 AdoAux.Recordset.MoveNext
              Loop
           End If
          .MoveNext
        Loop
    End If
   End With
  'Procesamos el Asiento Contable
   sSQL = "DELETE * " _
        & "FROM Asiento_SC " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND T_No = " & Trans_No & " " _
        & "AND CodigoU = '" & CodigoUsuario & "' "
   Ejecutar_SQL_SP sSQL
   Factura_No = 0
   SerieFactura = "000000"
   ContSC = 0
   IniciarAsientosDe DGAsiento, AdoAsiento
    With AdoExcelAdodc.Recordset
     If .RecordCount > 0 Then
        .MoveFirst
         Do While Not .EOF
            For IdField = 0 To .fields.Count - 1
                If IdField = 2 Then Codigo = Dato_Campo(.fields(IdField), True) Else Codigo = Dato_Campo(.fields(IdField))
                Select Case IdField + 1
                  Case 3: Mifecha = Codigo
                  Case 4: SubModuloCxCxP = Codigo
                  Case 5: If ParaFarmacia Then SubModuloCxCxP = Codigo
                  Case 6: If ParaFarmacia Then SerieFactura = Codigo
                  Case 7
                          If Not ParaFarmacia Then
                             SerieFactura = MidStrg(Codigo, 1, 6)
                             Factura_No = Val(MidStrg(Codigo, 7, 12))
                          End If
                  Case 8: If ParaFarmacia Then Factura_No = Val(Codigo) Else Cta_Gasto = Codigo
                 'Concepto que no se procesa
                  Case 11: Total_Sin_No_IVA = Redondear(Val(Codigo), 2)
                  Case 12: Total_Sin_IVA = Redondear(Val(Codigo), 2)
                  Case 13: If ParaFarmacia Then Total_Venta = Redondear(Val(Codigo), 2) Else Total_Con_IVA = Redondear(Val(Codigo), 2)
                  Case 14: Total_IVA = Redondear(Val(Codigo), 2)
                  Case 15: Total = Redondear(Val(Codigo), 2)
                           SubTotal = Redondear(Total_Con_IVA + Total_Sin_IVA + Total_Sin_No_IVA, 2)
                  Case 18: If ParaFarmacia Then Total_Ret = Redondear(Val(Codigo), 2)
                  Case 36: Cta_CajaG = Codigo
                  Case 42: SubModuloGasto = Codigo
                End Select
            Next IdField
           'Leemos el tipo de cta de gasto
            Codigo = Leer_Cta_Catalogo(Cta_Gasto)
            If SubCta = "G" Then
               sSQL = "SELECT Detalle " _
                    & "FROM Catalogo_SubCtas " _
                    & "WHERE Periodo = '" & Periodo_Contable & "' " _
                    & "AND Item = '" & NumEmpresa & "' " _
                    & "AND Codigo = '" & SubModuloGasto & "' " _
                    & "AND TC = '" & SubCta & "' "
               Select_Adodc AdoAux, sSQL
               'MsgBox AdoAux.Recordset.RecordCount
               If AdoAux.Recordset.RecordCount > 0 Then
                  SetAdoAddNew "Asiento_SC"
                  SetAdoFields "Codigo", SubModuloGasto
                  SetAdoFields "Beneficiario", AdoAux.Recordset.fields("Detalle")
                  SetAdoFields "DH", "1"
                  SetAdoFields "Valor", SubTotal
                  SetAdoFields "FECHA_V", Mifecha
                  SetAdoFields "TC", SubCta
                  SetAdoFields "Cta", Cta_Gasto
                  SetAdoFields "TM", "1"
                  SetAdoFields "T_No", Trans_No
                  SetAdoFields "SC_No", ContSC
                  SetAdoFields "Item", NumEmpresa
                  SetAdoFields "CodigoU", CodigoUsuario
                  SetAdoUpdate
                  ContSC = ContSC + 1
               End If
            End If
           'Ahora insertamos el abono
            If ParaFarmacia Then Codigo = Leer_Cta_Catalogo(Cta_CxP_Retenciones) Else Codigo = Leer_Cta_Catalogo(Cta_CajaG)
           'MsgBox ParaFarmacia & vbCrLf & SubModuloCxCxP & vbCrLf & SubCta
            Select Case SubCta
               Case "C", "P"
                    sSQL = "SELECT Codigo, Cliente " _
                         & "FROM Clientes " _
                         & "WHERE CI_RUC = '" & SubModuloCxCxP & "' "
                    Select_Adodc AdoAux, sSQL
                    If AdoAux.Recordset.RecordCount > 0 Then
                       CodigoCli = AdoAux.Recordset.fields("Codigo")
                       Beneficiario = AdoAux.Recordset.fields("Cliente")
                    Else
                       CodigoCli = "9999999999"
                       Beneficiario = "CONSUMIDOR FINAL"
                    End If
                    SetAdoAddNew "Asiento_SC"
                    SetAdoFields "Codigo", CodigoCli
                    SetAdoFields "Beneficiario", Beneficiario
                    SetAdoFields "DH", "2"
                    SetAdoFields "FECHA_V", Mifecha
                    SetAdoFields "TC", SubCta
                    SetAdoFields "TM", "1"
                    SetAdoFields "T_No", Trans_No
                    SetAdoFields "SC_No", ContSC
                    SetAdoFields "Item", NumEmpresa
                    SetAdoFields "CodigoU", CodigoUsuario
                    SetAdoFields "Detalle_SubCta", "DOC No. " & SerieFactura & "-" & Format(Factura_No, "000000000")
                    If ParaFarmacia Then
                       SetAdoFields "Cta", Cta_CxP_Retenciones
                       SetAdoFields "Factura", Val(Year(Mifecha) & Month(Mifecha))
                       SetAdoFields "Valor", Total_Venta - Total_Ret
                    Else
                       SetAdoFields "Valor", Total
                       SetAdoFields "Cta", Cta_CajaG
                       'SetAdoFields "Serie", SerieFactura
                       SetAdoFields "Factura", Factura_No
                    End If
                    SetAdoUpdate
                    ContSC = ContSC + 1
                   'MsgBox CodigoCli
            End Select
           .MoveNext
         Loop
     End If
    End With
    DGAsiento.Visible = False
    For IE = 0 To ContCtas - 1
        If CtasProc(IE).Valor >= 0 Then
           InsertarAsientos AdoAsiento, CtasProc(IE).Cta, 0, CtasProc(IE).Valor, 0
        Else
           InsertarAsientos AdoAsiento, CtasProc(IE).Cta, 0, 0, -CtasProc(IE).Valor
        End If
    Next IE
    DGAsiento.Visible = True
    'Verificacion SubTotal
    Debe = 0: Haber = 0
    With AdoAsiento.Recordset
     If .RecordCount > 0 Then
        .MoveFirst
         Do While Not .EOF
            Debe = Debe + .fields("DEBE")
            Haber = Haber + .fields("HABER")
           .MoveNext
         Loop
     End If
    End With
    LabelDebe.Caption = Format$(Debe, "#,##0.00")
    LabelHaber.Caption = Format$(Haber, "#,##0.00")
    LblDiferencia.Caption = Format$(Debe - Haber, "#,##0.00")
    If FechaIni = FechaFin Then
       LblConcepto.Caption = "Gastos de Caja del " & FechaIni & ", Diario No. ?"
    Else
       LblConcepto.Caption = "Gastos de Caja del " & FechaIni & " al " & FechaFin & ", Diario No. ?"
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set ftp = Nothing
   Control_Procesos "Q", "Salir Modulo de Importa Excel"
End Sub

Private Sub MBFechaI_GotFocus()
  MarcarTexto MBFechaI
End Sub

Private Sub MBFechaI_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBFechaI_LostFocus()
  FechaValida MBFechaI
  FechaComp = MBFechaI
  Fecha_Vence = MBFechaI
  FechaTexto = MBFechaI
End Sub

Public Sub Leer_Encabezado_FA()
   FA.CodigoC = Ninguno
   With AdoExcelAdodc.Recordset
    If .RecordCount > 0 And Not .EOF Then
       'CodigoCliente
        Codigo = Dato_Campo(.fields(0), , True)
        CodigoCli = "9999999999"   'RUC/Cedula/Consumidor Final
        If Len(Codigo) > 1 Then
           sSQL = "SELECT Codigo " _
                & "FROM Asiento_Beneficiarios " _
                & "WHERE Item = '" & NumEmpresa & "' " _
                & "AND RUC_CI = '" & Codigo & "' "
           Select_Adodc AdoAux, sSQL
           If AdoAux.Recordset.RecordCount > 0 Then CodigoCli = AdoAux.Recordset.fields("Codigo")
        End If
        CI_Representante = Codigo
        FA.CodigoC = CodigoCli
        FA.CI_RUC = Codigo
       'Fecha
        FA.Fecha = Dato_Campo(.fields(1), True)
       'Fecha Venc
        FA.Fecha_V = Dato_Campo(.fields(14), True)
       'Factura
        FA.Factura = Val(Dato_Campo(.fields(2)))
       'Serie
        FA.Serie = MidStrg(Dato_Campo(.fields(10)), 1, 6)
        If Len(FA.Serie) < 6 Then FA.Serie = "001001"
       'Estado
        Select Case MidStrg(Dato_Campo(.fields(11), , True), 1, 1)
          Case "T": FA.T = Anulado
          Case "C": FA.T = Cancelado
          Case Else: FA.T = Pendiente
        End Select
       'Cliente
        NombreCliente = Dato_Campo(.fields(12), , True)
        FA.Cliente = NombreCliente

       'Autorizacion
        Autorizacion = Dato_Campo(.fields(19), , True)
        If Len(Autorizacion) >= 6 Then FA.Autorizacion = Autorizacion Else FA.Autorizacion = Ninguno
    End If
   End With
End Sub

Public Sub Iniciar_Asiento_Beneficiario()
  sSQL = "DELETE * " _
       & "FROM Asiento_Beneficiarios " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND RUC_CI <> '9999999999999' "
  Ejecutar_SQL_SP sSQL
End Sub

Public Sub Insertar_Asiento_Beneficiario(RUC_CI_NIC As String, Beneficiario As String, Optional TipoTC As String)
Dim CodRUC_CI_NIC As String
Dim NoInserto As Boolean
Dim AdoDBAux As ADODB.Recordset

    NoInserto = True
    If TipoTC = "" Then TipoTC = Ninguno
    CodRUC_CI_NIC = Dato_Campo(RUC_CI_NIC, , True)
    If Len(CodRUC_CI_NIC) >= 1 And Len(Beneficiario) >= 1 Then
       sSQL = "INSERT INTO Asiento_Beneficiarios (Codigo, Beneficiario, TD, RUC_CI, Item) " _
            & "SELECT '.','" & UCaseStrg(Beneficiario) & "','" & TipoTC & "','" & CodRUC_CI_NIC & "','" & NumEmpresa & "' " _
            & "WHERE NOT EXISTS(SELECT 1 FROM Asiento_Beneficiarios " _
            & "                 WHERE Item = '" & NumEmpresa & "' " _
            & "                 AND RUC_CI = '" & CodRUC_CI_NIC & "' " _
            & "                 OR Beneficiario = '" & Beneficiario & "') "
       Ejecutar_SQL_SP sSQL
    End If
End Sub

Public Sub Actualizar_Asiento_Beneficiario_Clientes(ClienteFa As Boolean, Optional InsClienteFA As Boolean)
Dim AdoDBAux As ADODB.Recordset

    sSQL = "UPDATE Asiento_Beneficiarios " _
         & "SET Codigo = CS.Codigo, TD = CS.TC, Beneficiario = CS.Detalle " _
         & "FROM Asiento_Beneficiarios As AB, Catalogo_SubCtas As CS " _
         & "WHERE CS.Item = '" & NumEmpresa & "' " _
         & "AND CS.Periodo = '" & Periodo_Contable & "' " _
         & "AND CS.Codigo <> '.' " _
         & "AND AB.Item = CS.Item " _
         & "AND AB.RUC_CI = CS.Codigo "
    Ejecutar_SQL_SP sSQL

    sSQL = "UPDATE Asiento_Beneficiarios " _
         & "SET Codigo = C.Codigo, TD = C.TD " _
         & "FROM Asiento_Beneficiarios As AB, Clientes As C " _
         & "WHERE AB.Item = '" & NumEmpresa & "' " _
         & "AND C.CI_RUC <> '.' " _
         & "AND AB.RUC_CI = C.CI_RUC "
    Ejecutar_SQL_SP sSQL

    sSQL = "UPDATE Asiento_Beneficiarios " _
         & "SET Codigo = C.Codigo, TD = C.TD " _
         & "FROM Asiento_Beneficiarios As AB, Clientes As C " _
         & "WHERE AB.Item = '" & NumEmpresa & "' " _
         & "AND LEN(AB.Codigo) = 1 " _
         & "AND AB.RUC_CI = C.Codigo "
    Ejecutar_SQL_SP sSQL

    sSQL = "SELECT Codigo, Beneficiario, TD, RUC_CI, ID " _
         & "FROM Asiento_Beneficiarios " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND LEN(Codigo) = 1 " _
         & "AND LEN(RUC_CI) > 1 " _
         & "ORDER BY Beneficiario "
    Select_Adodc AdoAux, sSQL
    With AdoAux.Recordset
     If .RecordCount > 0 Then
         Progreso_Barra.Valor_Maximo = Progreso_Barra.Valor_Maximo + .RecordCount
         Do While Not .EOF
            Progreso_Barra.Mensaje_Box = "CREANDO A: " & .fields("Beneficiario")
            Progreso_Esperar
           'MsgBox "....."
            Select Case .fields("TD")
              Case "G", "I", "CC"
                   If .fields("TD") = "G" Then Si_No = True Else Si_No = False
                   Insertar_SubModulo .fields("RUC_CI"), .fields("Beneficiario"), .fields("TD"), Si_No
                  .fields("Codigo") = .fields("RUC_CI")
              Case Else
                   Insertar_Beneficiario_Nuevo .fields("RUC_CI"), .fields("Beneficiario"), InsClienteFA
                  .fields("Codigo") = Tipo_RUC_CI.Codigo_RUC_CI
                  .fields("TD") = Tipo_RUC_CI.Tipo_Beneficiario
            End Select
           .MoveNext
         Loop
        .UpdateBatch
     End If
    End With

    sSQL = "SELECT Codigo, Beneficiario, TD, RUC_CI, ID " _
         & "FROM Asiento_Beneficiarios " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND LEN(RUC_CI) = 1 " _
         & "ORDER BY Beneficiario "
    Select_Adodc AdoAux, sSQL
    With AdoAux.Recordset
     If .RecordCount > 0 Then
         Progreso_Barra.Valor_Maximo = Progreso_Barra.Valor_Maximo + .RecordCount
         DigVerif = Digito_Verificador(NumEmpresa)
         Codigo = Format(Tipo_RUC_CI.Codigo_RUC_CI, "0000000000")
         Do While Not .EOF
            Progreso_Barra.Mensaje_Box = "ACTUALIZANDO A: " & .fields("Beneficiario")
            Progreso_Esperar True
           .fields("TD") = "P"
           .fields("Codigo") = Codigo
           .fields("RUC_CI") = Codigo
            Codigo = Format(Val(Codigo) + 1, "0000000000")
           .MoveNext
         Loop
        .UpdateBatch
     End If
    End With
End Sub

Public Sub Actualizar_Asiento_Clientes(ClienteFa As Boolean, Optional InsClienteFA As Boolean)
Dim AdoDBAux As ADODB.Recordset
Dim Porc_Proc As String
Dim N As Long
    N = 0
    sSQL = "UPDATE Clientes " _
         & "SET Codigo = CS.Codigo, TD = CS.TC, Cliente = CS.Detalle " _
         & "FROM Clientes As C, Catalogo_SubCtas As CS " _
         & "WHERE CS.Item = '" & NumEmpresa & "' " _
         & "AND CS.Periodo = '" & Periodo_Contable & "' " _
         & "AND LEN(CS.TC) = 1 " _
         & "AND CS.Codigo <> '.' " _
         & "AND C.CI_RUC = CS.Codigo "
    Ejecutar_SQL_SP sSQL
        
    sSQL = "SELECT Codigo, Cliente, TD, CI_RUC, ID " _
         & "FROM Clientes " _
         & "WHERE Codigo = 'NINGUNO' " _
         & "AND LEN(CI_RUC) > 1 " _
         & "ORDER BY Cliente "
    Select_Adodc AdoAux, sSQL
    With AdoAux.Recordset
     If .RecordCount > 0 Then
         Do While Not .EOF
            Porc_Proc = Format(N / .RecordCount, "00.0%")
            Progreso_Barra.Mensaje_Box = "(" & Porc_Proc & ") CREANDO A: " & .fields("Cliente")
            Progreso_Esperar True
            DigVerif = Digito_Verificador(.fields("CI_RUC"))
           'MsgBox "....."
            Select Case Tipo_RUC_CI.Tipo_Beneficiario
              Case "G", "I", "CC"
                   If .fields("TD") = "G" Then Si_No = True Else Si_No = False
                   Insertar_SubModulo .fields("CI_RUC"), .fields("Cliente"), .fields("TD"), Si_No
                  .fields("Codigo") = .fields("CI_RUC")
              Case Else
                  .fields("Codigo") = Tipo_RUC_CI.Codigo_RUC_CI
                  .fields("TD") = Tipo_RUC_CI.Tipo_Beneficiario
            End Select
            N = N + 1
           .MoveNext
         Loop
        .UpdateBatch
     End If
    End With
    
    N = 0
    sSQL = "SELECT Codigo, Cliente, TD, CI_RUC, ID " _
         & "FROM Clientes " _
         & "WHERE LEN(CI_RUC) = 1 " _
         & "ORDER BY Cliente "
    Select_Adodc AdoAux, sSQL
    With AdoAux.Recordset
     If .RecordCount > 0 Then
         Progreso_Barra.Valor_Maximo = Progreso_Barra.Valor_Maximo + .RecordCount
         DigVerif = Digito_Verificador(NumEmpresa)
         Codigo = Format(Tipo_RUC_CI.Codigo_RUC_CI, "0000000000")
         Do While Not .EOF
            Porc_Proc = Format(N / .RecordCount, "00.0%")
            Progreso_Barra.Mensaje_Box = "(" & Porc_Proc & ") ACTUALIZANDO A: " & .fields("Cliente")
            Progreso_Esperar True
           .fields("TD") = "P"
           .fields("Codigo") = Codigo
           .fields("CI_RUC") = Codigo
            Codigo = Format(Val(Codigo) + 1, "0000000000")
            N = N + 1
           .MoveNext
         Loop
        .UpdateBatch
     End If
    End With
End Sub

Public Sub Insertar_Beneficiario_Nuevo(RUC_CI_NIC As String, Beneficiario As String, ClienteFa As Boolean, Optional InsClienteFA As Boolean)
Dim CodRUC_CI_NIC As String
Dim AdoDBAux As ADODB.Recordset

    CodRUC_CI_NIC = Dato_Campo(RUC_CI_NIC, , True)
    sSQL = "SELECT Codigo, TD " _
         & "FROM Clientes " _
         & "WHERE CI_RUC = '" & CodRUC_CI_NIC & "' "
    Select_AdoDB AdoDBAux, sSQL
    If AdoDBAux.RecordCount <= 0 Then
       If Len(CodRUC_CI_NIC) > 1 Then
         'MsgBox "Nuevo: " & Beneficiario
          DigVerif = Digito_Verificador(CodRUC_CI_NIC)
          Select Case Tipo_RUC_CI.Tipo_Beneficiario
            Case "C", "R", "P"
                 SetAdoAddNew "Clientes"
                 SetAdoFields "T", Normal
                 SetAdoFields "Codigo", Tipo_RUC_CI.Codigo_RUC_CI
                 SetAdoFields "TD", Tipo_RUC_CI.Tipo_Beneficiario
                 SetAdoFields "CI_RUC", CodRUC_CI_NIC
                 SetAdoFields "Cliente", UCaseStrg(Beneficiario)
                 SetAdoFields "Fecha", FechaSistema
                 SetAdoFields "Direccion", "SD"
                 SetAdoFields "DirNumero", "SN"
                 SetAdoFields "Ciudad", "QUITO"
                 SetAdoFields "Prov", "17"
                 SetAdoFields "Pais", "593"
                 SetAdoFields "FA", ClienteFa
                 SetAdoFields "CodigoU", CodigoUsuario
                 SetAdoUpdate
          End Select
       End If
    Else
       Tipo_RUC_CI.Codigo_RUC_CI = AdoDBAux.fields("Codigo")
       Tipo_RUC_CI.Tipo_Beneficiario = AdoDBAux.fields("TD")
    End If
    AdoDBAux.Close
    
    If InsClienteFA And RUC = "1792509327001" Then
       sSQL = "SELECT * " _
            & "FROM Clientes_Matriculas " _
            & "WHERE Item = '" & NumEmpresa & "' " _
            & "AND Periodo = '" & Periodo_Contable & "' " _
            & "AND Codigo = '" & Tipo_RUC_CI.Codigo_RUC_CI & "' "
       Select_AdoDB AdoDBAux, sSQL
       If AdoDBAux.RecordCount <= 0 Then
          SetAdoAddNew "Clientes_Matriculas"
          SetAdoFields "T", Normal
          SetAdoFields "Codigo", Tipo_RUC_CI.Codigo_RUC_CI
          SetAdoFields "TD", "R"
          SetAdoFields "Cedula_R", "1790764575001"
          SetAdoFields "Representante", "CENTRO MEDICO MATERNAL PAEZ ALMEIDA Y NARANJO"
          SetAdoFields "Lugar_Trabajo_R", "GARCIA MORENO Y ESMERALDAS"
          SetAdoFields "Telefono_R", "022282950"
          SetAdoFields "Grupo_No", "FHARMA-" & NumEmpresa
          SetAdoFields "Item", NumEmpresa
          SetAdoFields "Periodo", Periodo_Contable
          SetAdoFields "CodigoU", CodigoUsuario
          SetAdoUpdate
       End If
       AdoDBAux.Close
    End If
End Sub

Public Sub Insertar_SubModulo(CodigoSubModulo As String, CuentaSubModulo As String, TipoDeCta As String, EsSubModuloGasto As Boolean)
Dim CodigoSubMod As String
Dim NombreSubMod As String

    CodigoSubMod = UCaseStrg(Dato_Campo(CodigoSubModulo))
    'MsgBox CodigoSubMod
    If Len(CodigoSubMod) > 1 Then
       If EsSubModuloGasto And TipoDeCta <> Ninguno Then
          sSQL = "SELECT Codigo " _
               & "FROM Catalogo_SubCtas " _
               & "WHERE Item = '" & NumEmpresa & "' " _
               & "AND Periodo = '" & Periodo_Contable & "' " _
               & "AND Codigo = '" & CodigoSubMod & "' "
          Select_Adodc AdoAux, sSQL
          If AdoAux.Recordset.RecordCount <= 0 Then
             sSQL = "SELECT Cliente " _
                  & "FROM Clientes " _
                  & "WHERE Codigo = '" & CodigoSubMod & "' "
             Select_Adodc AdoAux, sSQL
             If AdoAux.Recordset.RecordCount > 0 Then
                NombreSubMod = TrimStrg(MidStrg(AdoAux.Recordset.fields("Cliente"), 1, 60))
             Else
                NombreSubMod = "Codigo sin asignar: " & CodigoSubMod
             End If
             SetAdoAddNew "Catalogo_SubCtas"
             SetAdoFields "TC", TipoDeCta
             SetAdoFields "Codigo", CodigoSubMod
             SetAdoFields "Detalle", NombreSubMod
             SetAdoFields "Nivel", "00"
             SetAdoUpdate
          End If
       Else
          sSQL = "SELECT Codigo " _
               & "FROM Catalogo_CxCxP " _
               & "WHERE Item = '" & NumEmpresa & "' " _
               & "AND Periodo = '" & Periodo_Contable & "' " _
               & "AND TC = '" & TipoDeCta & "' " _
               & "AND Cta = '" & CuentaSubModulo & "' " _
               & "AND Codigo = '" & CodigoSubMod & "' "
          Select_Adodc AdoAux, sSQL
          If AdoAux.Recordset.RecordCount <= 0 Then
             SetAdoAddNew "Catalogo_CxCxP"
             SetAdoFields "TC", TipoDeCta
             SetAdoFields "Codigo", CodigoSubMod
             SetAdoFields "Cta", CuentaSubModulo
             SetAdoUpdate
          End If
       End If
    End If
End Sub

Public Sub Sacar_Ciudad_Fecha(Campo As String, vCiudad As String, vFecha As String)
Dim PC As Long
Dim dFecha As String
Dim dDia As Byte
Dim dAnio As Integer
    Campo = TrimStrg(UCaseStrg(Campo))
    If Len(Campo) <= 1 Then
       vCiudad = NombreCiudad
       vFecha = FechaSistema
    Else
       PC = InStr(Campo, ",")
       If PC = 0 Then
          PC = 1
          vCiudad = Ninguno
          dFecha = Ninguno
          While PC <= Len(Campo)
            If IsNumeric(MidStrg(Campo, PC, 1)) Then
               If PC = 1 Then
                  dFecha = Campo
               Else
                  vCiudad = TrimStrg(MidStrg(Campo, 1, PC - 1))
                  dFecha = TrimStrg(MidStrg(Campo, PC, Len(Campo)))
               End If
               PC = Len(Campo)
            End If
            PC = PC + 1
          Wend
          If vCiudad = Ninguno Then vCiudad = NombreCiudad
          If dFecha = Ninguno Then dFecha = UCaseStrg(FechaStrg(FechaSistema))
       Else
          vCiudad = TrimStrg(MidStrg(Campo, 1, PC - 1))
          vCiudad = Replace(vCiudad, ",", "")
          dFecha = TrimStrg(MidStrg(Campo, PC + 1, Len(Campo)))
       End If
       PC = 1
       If InStr(dFecha, "ENE") > 0 Then PC = 1
       If InStr(dFecha, "FEB") > 0 Then PC = 2
       If InStr(dFecha, "MAR") > 0 Then PC = 3
       If InStr(dFecha, "ABR") > 0 Then PC = 4
       If InStr(dFecha, "MAY") > 0 Then PC = 5
       If InStr(dFecha, "JUN") > 0 Then PC = 6
       If InStr(dFecha, "JUL") > 0 Then PC = 7
       If InStr(dFecha, "AGO") > 0 Then PC = 8
       If InStr(dFecha, "SEP") > 0 Then PC = 9
       If InStr(dFecha, "OCT") > 0 Then PC = 10
       If InStr(dFecha, "NOV") > 0 Then PC = 11
       If InStr(dFecha, "DIC") > 0 Then PC = 12
       If Val(TrimStrg(SinEspaciosIzq(dFecha))) <= 31 Then dDia = Val(TrimStrg(SinEspaciosIzq(dFecha))) Else dDia = 0
       If 1900 <= Val(TrimStrg(SinEspaciosDer(dFecha))) And Val(TrimStrg(SinEspaciosDer(dFecha))) <= 2050 Then
          dAnio = Val(TrimStrg(SinEspaciosDer(dFecha)))
       Else
          dAnio = 0
       End If
       If dDia = 0 Then dDia = Day(FechaSistema)
       If dAnio = 0 Then dAnio = Year(FechaSistema)
       vFecha = Format(dDia, "00") & "/" & Format(PC, "00") & "/" & Format(dAnio, "0000")
      'MsgBox vCiudad & vbCrLf & vFecha
    End If
End Sub

Public Sub Quitar_Duplicado_Trans(DebeHaber As String)
Dim AdoDBAsiento As ADODB.Recordset
Dim AdoDBAux As ADODB.Recordset
Dim A_No As Long
Dim IdA As Long
Dim SiInsTrans As Boolean
Dim Suma As Currency
Dim SQL As String
Dim Cta As String
   A_No = 1000
   SQL = "SELECT CODIGO, SUM(" & DebeHaber & ") As SumTrans " _
       & "FROM Asiento " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND T_No = " & Trans_No & " " _
       & "AND " & DebeHaber & " > 0 " _
       & "GROUP BY CODIGO, CUENTA " _
       & "ORDER BY CODIGO "
   Select_AdoDB AdoDBAsiento, SQL
   If AdoDBAsiento.RecordCount > 0 Then
      Do While Not AdoDBAsiento.EOF
         Suma = AdoDBAsiento.fields("SumTrans")
         Cta = AdoDBAsiento.fields("CODIGO")
         SQL = "SELECT TOP 1 * " _
             & "FROM Asiento " _
             & "WHERE Item = '" & NumEmpresa & "' " _
             & "AND CodigoU = '" & CodigoUsuario & "' " _
             & "AND T_No = " & Trans_No & " " _
             & "AND CODIGO = '" & Cta & "' " _
             & "AND " & DebeHaber & " > 0 " _
             & "ORDER BY A_No "
         Select_AdoDB AdoDBAux, SQL
         If AdoDBAux.RecordCount > 0 Then
            SetAdoAddNew "Asiento"
            For IdA = 0 To AdoDBAux.fields.Count - 1
                SiInsTrans = True
                If AdoDBAux.fields(IdA).Name = "A_No" Then SiInsTrans = False
                If AdoDBAux.fields(IdA).Name = DebeHaber Then SiInsTrans = False
                If SiInsTrans Then SetAdoFields AdoDBAux.fields(IdA).Name, AdoDBAux.fields(IdA)
            Next IdA
            SetAdoFields DebeHaber, Suma
            SetAdoFields "A_No", A_No
            SetAdoUpdate
         End If
         AdoDBAux.Close
         A_No = A_No + 1
         AdoDBAsiento.MoveNext
      Loop
   End If
   AdoDBAsiento.Close
   SQL = "DELETE * " _
       & "FROM Asiento " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND T_No = " & Trans_No & " " _
       & "AND " & DebeHaber & " > 0 " _
       & "AND A_No < 1000 "
   Ejecutar_SQL_SP SQL
End Sub
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'''Public Sub Importar_Depositos()
'''Dim I As Long
'''Dim N As Long
'''Dim NumSem As String
'''Dim TDebe As Currency
'''Dim THaber As Currency
'''Dim Saldo_Final As Currency
'''Dim Dias_Fin_Anio As Integer
'''
'''  Cadena = "INGRESE EL NUEMRO DE SEMANA" & vbCrLf & vbCrLf _
'''         & "O PRESIONE:" & vbCrLf & vbCrLf _
'''         & "T = DEPOSITO TOTAL" & vbCrLf & vbCrLf _
'''         & "R = N/C POR ROL DE PAGOS"
'''  NumSem = InputBox(Cadena, "IMPORTACION DE TABLA DE DEPOSITOS", "T")
'''  Mifecha = BuscarFecha(FechaSistema)
'''
'''  Dias_Fin_Anio = CFechaLong("31/12/" & Year(FechaSistema)) - CFechaLong(FechaSistema)
'''  If Dias_Fin_Anio <= 0 Then Dias_Fin_Anio = 1
'''  sSQL = "DELETE * " _
'''       & "FROM Trans_Libretas " _
'''       & "WHERE Item = '" & NumEmpresa & "' " _
'''       & "AND TP = 'DEFR' " _
'''       & "AND Fecha = #" & Mifecha & "# "
'''  Ejecutar_SQL_SP sSQL
'''  With AdoExcelAdodc.Recordset
'''       For i = 1 To .rows - 1
'''          .Row = i
'''          .Col = 1
'''           Cuenta_No = TrimStrg(.Text)
'''           If NumSem = "1" Then .Col = 4
'''           If NumSem = "2" Then .Col = 5
'''           If NumSem = "3" Then .Col = 6
'''           If NumSem = "4" Then .Col = 7
'''           If NumSem = "5" Then .Col = 8
'''           If NumSem = "T" Then .Col = 9
'''           If NumSem = "R" Then .Col = 9
'''           Total = Val(.Text)
'''           TDebe = 0
'''           THaber = Total
'''           If NumSem = "R" Then
'''              TipoProc = "N/CR"
'''           Else
'''              TipoProc = "DEFR"
'''           End If
'''
'''         'Insertar Transacciones de Libreta
'''          If Total > 0 Then
'''            sSQL = "SELECT TOP 1 * " _
'''                 & "FROM Trans_Libretas " _
'''                 & "WHERE Cuenta_No = '" & Cuenta_No & "' " _
'''                 & "ORDER BY Fecha DESC,IDT DESC,Hora DESC,ID DESC "
'''            Select_Adodc AdoAux, sSQL
'''            With AdoAux.Recordset
'''                 If .RecordCount > 0 Then
'''                     SaldoDisp = .Fields("Saldo_Disp")
'''                     SaldoCont = .Fields("Saldo_Cont")
'''                     ID_Trans = .Fields("IDT")
'''                     NumeroLineas = .Fields("ID")
'''                 Else
'''                     SaldoCont = 0
'''                 End If
'''                 TiempoTexto = Format$(Time, FormatoTimes)
'''                .AddNew
'''                .Fields("Fecha") = FechaSistema
'''                .Fields("Cuenta_No") = Cuenta_No
'''                .Fields("TP") = TipoProc
'''                .Fields("Debitos") = TDebe
'''                .Fields("Creditos") = THaber
'''                .Fields("Saldo_Cont") = SaldoCont + THaber - TDebe
'''                 If TipoGrupo Then
'''                    If THaber <> 0 Then
'''                      '.Fields("Saldo_Disp") = SaldoDisp
'''                      .Fields("Saldo_Disp") = SaldoCont + THaber - TDebe
'''                      .Fields("T") = Normal
'''                      Saldo_Final = SaldoCont + THaber - TDebe
'''                    Else
'''                      .Fields("Saldo_Disp") = SaldoDisp + THaber - TDebe
'''                      .Fields("T") = Normal
'''                      Saldo_Final = SaldoDisp + THaber - TDebe
'''                    End If
'''                 Else
'''                   .Fields("Saldo_Disp") = SaldoDisp + THaber - TDebe
'''                   .Fields("T") = Normal
'''                    Saldo_Final = SaldoDisp + THaber - TDebe
'''                 End If
'''                .Fields("CodigoU") = CodigoUsuario
'''                 If NumeroLineas >= 36 Then NumeroLineas = 1
'''                .Fields("IP") = adFalse
'''                .Fields("CHT") = adFalse
'''                 If NumSem = "R" Then
'''                    .Fields("Banco") = "N/C POR ROL"
'''                 Else
'''                    .Fields("Banco") = "DEP TABLA"
'''                 End If
'''                .Fields("ACL") = adFalse
'''                .Fields("AC") = adFalse              ' Quitar
'''                .Fields("ACC") = adFalse
'''                .Fields("IDT") = ID_Trans + 1
'''                .Fields("Hora") = TiempoTexto
'''                .Fields("Item") = NumEmpresa
'''                .Fields("ME") = adFalse
'''                .Fields("Cartilla_No") = Cartilla_No
'''                .Fields("Papeleta_No") = 0
'''                 SetUpdate AdoAux
'''            End With
'''            If TipoProc = "DEFR" Then
'''               sSQL = "SELECT * " _
'''                    & "FROM Trans_Bloqueos "
'''               Select_Adodc AdoAux, sSQL
'''               With AdoAux.Recordset
'''                   .AddNew
'''                   .Fields("T") = Normal
'''                   .Fields("Fecha") = FechaSistema
'''                   .Fields("Cuenta_No") = Cuenta_No
'''                   .Fields("Valor") = Total
'''                   .Fields("Cheque") = TipoProc
'''                   .Fields("Banco") = "DEPOSITO PROGRAMADO"
'''                   .Fields("Dias") = Dias_Fin_Anio
'''                   .Fields("Item") = NumEmpresa
'''                   .Update
'''               End With
'''            End If
'''         End If
'''         Me.Caption = "Importar de FlexGrid a Sistema " & i & " de " & Rango.NumFila2 & " - " & Cuenta_No & " - " & Format$(Total, "#,##0.00")
'''       Next
'''  End With
'''End Sub
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'''Public Sub Importar_Facturas_2()
'''Dim F As Long
'''Dim N As Long
'''Dim Tot_Propinas As Currency
'''
'''  Encerar_Factura FA
'''  FA.Cod_CxC = DCLinea.Text
'''  Lineas_De_CxC FA
'''  SerieFactura = FA.Serie
'''  Fecha_Vence = FA.Vencimiento
'''  Autorizacion = FA.Autorizacion
'''  Cta_Cobrar = FA.Cta_CxP
'''  Bandera = False
'''  Evaluar = True
''' 'Borramos la tabla temporal de facturas emitidas
'''  sSQL = "DELETE * " _
'''       & "FROM Asiento_F " _
'''       & "WHERE Item = '" & NumEmpresa & "' " _
'''       & "AND CodigoU = '" & CodigoUsuario & "' "
'''  Ejecutar_SQL_SP sSQL
'''  With AdoExcelAdodc.Recordset
'''     'Empezamos la importacion de las facturas
'''       For F = 1 To .rows - 1
'''          .Row = F
'''           FA.Total_MN = 0
'''           FA.Total_IVA = 0
'''           FA.SubTotal = 0
'''           FA.Con_IVA = 0
'''           FA.Sin_IVA = 0
'''           For N = 1 To .cols - 1
'''              .Col = N
'''               Codigo = TrimStrg(Replace(.Text, "'", ""))
'''              'MsgBox Codigo & "...."
'''               Select Case N
'''                 Case 1: FA.Serie = Codigo
'''                         FA.Serie = Replace(FA.Serie, "-", "")
'''                         If Val(FA.Serie) < 1001 Then FA.Serie = "001001"
'''                 Case 2: FA.Factura = Val(Codigo)
'''                 Case 3: FA.Fecha = Codigo
'''                 Case 4: FA.TC = MidStrg(Codigo, 1, 2)
'''                 Case 5: NombreCliente = Codigo
'''                 Case 6: Producto = .Text
'''                 Case 7: Producto = Producto & vbTab & Codigo
'''                 Case 8: Producto = Producto & vbTab & Codigo
'''                 Case 9: Producto = Producto & vbTab & Codigo
'''                 Case 12: FA.SubTotal = Redondear(Val(Codigo), 2)
'''                 Case 13: If Codigo = "X" Then FA.Total_IVA = Redondear(FA.SubTotal * Porc_IVA, 2)
'''                 Case 14: CodigoP = Codigo
'''                          Producto = Codigo & vbTab & Producto
'''                 Case 15: Cta = Codigo
'''                          CodigoInv = "99.99"
'''                          sSQL = "SELECT * " _
'''                               & "FROM Catalogo_Productos " _
'''                               & "WHERE Codigo_Inv = '" & Cta & "' " _
'''                               & "AND Item = '" & NumEmpresa & "' " _
'''                               & "AND Periodo = '" & Periodo_Contable & "' "
'''                          Select_Adodc AdoAux, sSQL
'''                          If AdoAux.Recordset.RecordCount > 0 Then
'''                             CodigoA = AdoAux.Recordset.Fields("Producto")
'''                             CodigoInv = AdoAux.Recordset.Fields("Codigo_Inv")
'''                          End If
'''                 Case 16: CodigoCli = "9999999999"
'''                          Beneficiario = "CONSUMIDOR FINAL"
'''                          CI_Representante = CodigoCli
'''                          sSQL = "SELECT * " _
'''                               & "FROM Clientes " _
'''                               & "WHERE CI_RUC = '" & Codigo & "' "
'''                          Select_Adodc AdoAux, sSQL
'''                          If AdoAux.Recordset.RecordCount > 0 Then
'''                             CodigoCli = AdoAux.Recordset.Fields("Codigo")
'''                             CI_Representante = AdoAux.Recordset.Fields("CI_RUC")
'''                             Beneficiario = AdoAux.Recordset.Fields("Cliente")
'''                          End If
'''                          FA.CodigoC = CodigoCli
'''               End Select
'''
'''           Next N
'''           'MsgBox FA.SubTotal
'''            If FA.SubTotal > 0 Then
'''               SetAdoAddNew "Asiento_F"
'''               SetAdoFields "FECHA", FA.Fecha
'''               SetAdoFields "CODIGO", CodigoInv
'''               SetAdoFields "PRODUCTO", Producto
'''               SetAdoFields "CANT", 1
'''               SetAdoFields "PRECIO", FA.SubTotal
'''               SetAdoFields "TOTAL", FA.SubTotal
'''               SetAdoFields "Total_IVA", FA.Total_IVA
'''               SetAdoFields "Serie", FA.Serie
'''               SetAdoFields "Autorizacion", FA.Autorizacion
'''               SetAdoFields "Numero", FA.Factura
'''               SetAdoFields "Codigo_Cliente", CodigoCli
'''               SetAdoFields "CodigoU", CodigoUsuario
'''               SetAdoFields "A_No", F
'''               SetAdoFields "Item", NumEmpresa
'''               SetAdoUpdate
'''            End If
'''           Me.Caption = "Importar de FlexGrid a Sistema de Facturacion El Numero: " & FA.Factura & ": " & Format$(F / Rango.NumFila2, "00%")
'''      Next F
'''  End With
'''  Generar_Facturas
'''End Sub
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'''Public Sub Importar_Facturas_3()
'''Dim I As Long
'''Dim N As Long
'''Dim Tot_Propinas As Currency
'''  Encerar_Factura FA
'''  FA.Cod_CxC = DCLinea.Text
'''  Lineas_De_CxC FA
'''  SerieFactura = FA.Serie
'''  Fecha_Vence = FA.Vencimiento
'''  Autorizacion = FA.Autorizacion
'''  Cta_Cobrar = FA.Cta_CxP
'''  Bandera = False
'''  Evaluar = True
'''  With AdoExcelAdodc.Recordset
'''      .Row = 1
'''      .Col = 2
'''       Mifecha = TrimStrg(.Text)
'''       For i = 1 To .rows - 1
'''          .Row = i
'''          .Col = 3
'''           If Mifecha <> TrimStrg(.Text) Then
'''              If IsDate(Mifecha) Then Eliminar_Facturas
'''             'MsgBox Mifecha
'''              Mifecha = TrimStrg(.Text)
'''           End If
'''           Me.Caption = "Revisando Datos en el excel: " & i & " de " & Rango.NumFila2 & ", Fecha: " & Mifecha
'''      Next i
'''      If IsDate(Mifecha) Then Eliminar_Facturas
'''  End With
''' 'Empezamos la importacion de las facturas
'''  FA.Factura = 1
'''  sSQL = "SELECT * " _
'''       & "FROM Facturas " _
'''       & "WHERE Item = '" & NumEmpresa & "' " _
'''       & "AND Periodo = '" & Periodo_Contable & "' " _
'''       & "AND TC = '" & FA.TC & "' " _
'''       & "ORDER BY Factura DESC "
'''  Select_Adodc AdoAux, sSQL
'''  If AdoAux.Recordset.RecordCount > 0 Then FA.Factura = AdoAux.Recordset.Fields("Factura") + 1
'''  With AdoExcelAdodc.Recordset
'''      'MsgBox .Rows & vbCrLf & .Cols
'''       For i = 1 To .rows - 1
'''          .Row = i
'''           FA.Total_MN = 0
'''           FA.Total_IVA = 0
'''           FA.SubTotal = 0
'''           FA.Con_IVA = 0
'''           FA.Sin_IVA = 0
'''           CodigoP = Ninguno
'''           CI_Representante = Ninguno
'''           For N = 1 To .cols - 1
'''              .Col = N
'''               Codigo = TrimStrg(Replace(.Text, "'", ""))
'''               Codigo1 = TrimStrg(Codigo)
'''               Select Case N
'''                 Case 1: 'RUC/Cedula/Consumidor Final
'''                          CodigoCli = "9999999999"
'''                         'MsgBox "Codigo: " & Codigo
'''                          If Len(Codigo) > 1 Then
'''                             DigVerif = Digito_Verificador( Codigo)
'''                             CodigoCli = Tipo_RUC_CI.Codigo_RUC_CI
'''                          End If
'''                          CI_Representante = Codigo
'''                         'MsgBox "Digito Verif. " & DigVerif & " (" & Tipo_RUC_CI.Tipo_Beneficiario & ")"
'''                 Case 2: If Val(Codigo) > 0 Then FA.Factura = Val(Codigo)
'''                 Case 3: FA.Fecha = Convertir_Fecha(Codigo)
'''                 Case 4: FA.SubTotal = Redondear(Val(Codigo), 2)
'''                 Case 5: FA.Total_IVA = Redondear(Val(Codigo), 2)
'''                 Case 6: FA.Descuento = Redondear(Val(Codigo), 2)
'''                 Case 7: FA.Total_MN = Redondear(Val(Codigo), 2)
'''                 Case 8: CodigoP = Codigo        'Representante
'''                 Case 9: CodigoA = Codigo        'Codigo Alumno
'''                         CodigoB = Codigo
'''                 Case 10: NombreCliente = Codigo  'Alumno
'''               End Select
'''           Next N
'''           Factura_No = FA.Factura
'''           Si_No = True
'''           CodigoCli = Ninguno
'''           Grupo_No = "ATS"
'''           With AdoClientes.Recordset
'''            If .RecordCount > 0 Then
'''                Do While Len(CodigoA) <= 10 And Si_No
'''                  .MoveFirst
'''                  .Find ("CI_RUC = '" & CodigoA & "' ")
'''                   If Not .EOF Then
'''                      CodigoCli = .Fields("Codigo")
'''                      NombreCliente = .Fields("Cliente")
'''                      Grupo_No = .Fields("Grupo")
'''                      Si_No = False
'''                   Else
'''                      CodigoA = "0" & CodigoA
'''                   End If
'''                Loop
'''            End If
'''           End With
'''           If CodigoCli = Ninguno Then CodigoCli = CodigoB
'''          'Grabamos el numero de factura
'''          'MsgBox FA.Factura
'''           If FA.Factura <> 0 Then
'''              'MsgBox FA.Factura
'''              sSQL = "DELETE * " _
'''                   & "FROM Asiento_F " _
'''                   & "WHERE Item = '" & NumEmpresa & "' " _
'''                   & "AND CodigoU = '" & CodigoUsuario & "' "
'''              Ejecutar_SQL_SP sSQL
'''              SetAdoAddNew "Asiento_F"
'''              If FA.TC = "NV" Then
'''                 SetAdoFields "CODIGO", "99.97"
'''                 SetAdoFields "PRODUCTO", "VENTAS TICKET DEL DIA"
'''                 TA.Banco = "Efectivo"
'''              Else
'''                 SetAdoFields "CODIGO", "99.99"
'''                 SetAdoFields "PRODUCTO", "VENTAS DEL DIA"
'''              End If
'''              SetAdoFields "CANT", 1
'''              SetAdoFields "PRECIO", FA.SubTotal
'''              SetAdoFields "TOTAL", FA.SubTotal
'''              SetAdoFields "Total_Desc", FA.Descuento
'''              SetAdoFields "Total_IVA", FA.Total_IVA
'''              SetAdoFields "CodigoU", CodigoUsuario
'''              SetAdoFields "TICKET", CStr(Year(FA.Fecha))
'''              SetAdoFields "Item", NumEmpresa
'''              SetAdoUpdate
'''              FA.T = "P"
'''              FA.CodigoC = CodigoCli
'''              FA.Fecha_C = FA.Fecha
'''              FA.Fecha_V = FA.Fecha
'''              If FA.Total_IVA > 0 Then
'''                 FA.Con_IVA = FA.SubTotal
'''              Else
'''                 FA.Sin_IVA = FA.SubTotal
'''              End If
'''              If FA.TC = "NV" Then
'''                 Tot_Propinas = 0
'''                 FA.Descuento = 0
'''                 FA.Servicio = 0
'''              End If
'''              If FA.Total_MN <> (FA.SubTotal + FA.Total_IVA - FA.Descuento + FA.Servicio + Tot_Propinas) Then
'''                 FA.Total_MN = FA.SubTotal + FA.Total_IVA - FA.Descuento + FA.Servicio + Tot_Propinas
'''              End If
'''              FA.Saldo_MN = FA.Total_MN
'''             'MsgBox FA.Fecha
'''              Grabar_Factura FA, True
'''              FA.Factura = FA.Factura + 1
'''           End If
'''           sSQL = "SELECT * " _
'''                & "FROM Clientes " _
'''                & "WHERE Codigo = '" & CodigoCli & "' "
'''           Select_Adodc AdoAux, sSQL
'''           If AdoAux.Recordset.RecordCount > 0 Then
'''              AdoAux.Recordset.Fields("Cedula") = CI_Representante
'''              AdoAux.Recordset.Fields("CI_RUC_SRI") = CI_Representante
'''              AdoAux.Recordset.Fields("TD_SRI") = Tipo_RUC_CI.Tipo_Beneficiario
'''              AdoAux.Recordset.Fields("Representante") = UCaseStrg(CodigoP)
'''              AdoAux.Recordset.Update
'''           Else
'''             'MsgBox NombreCliente
'''              SetAdoAddNew "Clientes"
'''              SetAdoFields "T", Normal
'''              SetAdoFields "Codigo", CodigoCli
'''              SetAdoFields "TD", "O"
'''              SetAdoFields "CI_RUC", CodigoB
'''              SetAdoFields "Cliente", UCaseStrg(NombreCliente)
'''              SetAdoFields "Representante", UCaseStrg(CodigoP)
'''              SetAdoFields "Cedula", CI_Representante
'''              SetAdoFields "CI_RUC_SRI", CI_Representante
'''              SetAdoFields "TD_SRI", Tipo_RUC_CI.Tipo_Beneficiario
'''              SetAdoFields "Fecha", FechaSistema
'''              SetAdoFields "Direccion", "SD"
'''              SetAdoFields "DirNumero", "SN"
'''              SetAdoFields "Ciudad", "QUITO"
'''              SetAdoFields "Grupo", "ATS"
'''              SetAdoFields "Prov", "17"
'''              SetAdoFields "Pais", "593"
'''              SetAdoFields "FA", True
'''              SetAdoFields "CodigoU", CodigoUsuario
'''              SetAdoUpdate
'''           End If
'''           sSQL = "SELECT * " _
'''                & "FROM Clientes_Matriculas " _
'''                & "WHERE Codigo = '" & CodigoCli & "' "
'''           Select_Adodc AdoAux, sSQL
'''           If AdoAux.Recordset.RecordCount > 0 Then
'''              AdoAux.Recordset.Fields("Cedula_R") = CI_Representante
'''              AdoAux.Recordset.Fields("TD") = Tipo_RUC_CI.Tipo_Beneficiario
'''              AdoAux.Recordset.Fields("Representante") = UCaseStrg(CodigoP)
'''              AdoAux.Recordset.Fields("Representante_Alumno") = UCaseStrg(CodigoP)
'''              AdoAux.Recordset.Update
'''           Else
'''              SetAdoAddNew "Clientes_Matriculas"
'''              SetAdoFields "T", Normal
'''              SetAdoFields "Codigo", CodigoCli
'''              SetAdoFields "TD", Tipo_RUC_CI.Tipo_Beneficiario
'''              SetAdoFields "Cedula_R", CI_Representante
'''              SetAdoFields "Representante", UCaseStrg(CodigoP)
'''              SetAdoFields "Representante_Alumno", UCaseStrg(CodigoP)
'''              SetAdoFields "Fecha", FechaSistema
'''              SetAdoFields "Direccion", "SD"
'''              SetAdoFields "DirNumero", "SN"
'''              SetAdoFields "Grupo_No", "ATS"
'''              SetAdoFields "Lugar_Nac", "QUITO"
'''              SetAdoFields "Nacionalidad", "ECUATORIANA"
'''              SetAdoFields "CodigoU", CodigoUsuario
'''              SetAdoUpdate
'''           End If
'''           Me.Caption = "Importar de FlexGrid a Sistema de Facturacion El Numero: " & FA.Factura & ": " & i & " de " & Rango.NumFila2
'''      Next i
'''  End With
'''End Sub
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'''Public Sub Importar_Consumos()
'''Dim I As Long
'''Dim N As Long
'''Dim Cod1 As String
'''Dim Cod2 As String
'''Dim Cod3 As String
'''Dim Cod4 As String
'''Dim Cod5 As String
'''Dim Cod6 As String
'''Dim Cod7 As String
'''Dim Cod8 As String
'''
'''Dim Tot_Propinas As Currency
'''Dim Encontro_Consumo As Boolean
'''  DGExcelAdodc.Visible = False
'''  Ln_No = 0
'''  sSQL = "SELECT C.Codigo,C.Cliente,C.Grupo,CDE.Cuenta_No " _
'''       & "FROM Clientes As C, Clientes_Datos_Extras As CDE " _
'''       & "WHERE CDE.Item = '" & NumEmpresa & "' " _
'''       & "AND CDE.Tipo_Dato = 'MEDIDOR' " _
'''       & "AND C.Codigo = CDE.Codigo " _
'''       & "ORDER BY C.Cliente,CDE.Cuenta_No "
'''  Select_Adodc AdoClientes, sSQL
''' 'Empezamos la importacion de las facturas
'''  With AdoExcelAdodc.Recordset
'''      'MsgBox .Rows & vbCrLf & .Cols
'''      .Row = 0
'''      .Col = 3: Cod1 = SinEspaciosDer(.Text)
'''      .Col = 9: Cod2 = SinEspaciosDer(.Text)
'''      .Col = 11: Cod3 = SinEspaciosDer(.Text)
'''      .Col = 12: Cod4 = SinEspaciosDer(.Text)
'''      .Col = 13: Cod5 = SinEspaciosDer(.Text)
'''      .Col = 14: Cod6 = SinEspaciosDer(.Text)
'''      .Col = 15: Cod7 = SinEspaciosDer(.Text)
'''      .Col = 16: Cod8 = SinEspaciosDer(.Text)
'''      .Row = 1
'''      .Col = 10
'''       Mifecha = TrimStrg(.Text)
'''       NoMes = CInt(Month(Mifecha))
'''       MiAnio = Year(Mifecha)
'''       sSQL = "DELETE * " _
'''            & "FROM Clientes_Facturacion " _
'''            & "WHERE Item = '" & NumEmpresa & "' " _
'''            & "AND Num_Mes = " & NoMes & " " _
'''            & "AND Periodo = '" & MiAnio & "' " _
'''            & "AND Codigo_Inv IN ('" & Cod1 & "','" & Cod2 & "','" & Cod3 & "','" & Cod4 & "','" & Cod5 & "','" & Cod6 & "','" & Cod7 & "','" & Cod8 & "') "
'''       Ejecutar_SQL_SP sSQL
'''       For i = 1 To .rows - 1
'''          .Row = i
'''           For N = 1 To .cols - 1
'''              .Col = N
'''               Codigo = TrimStrg(Replace(.Text, "'", ""))
'''               Select Case N
'''                 Case 1: Cuenta_No = Format$(Val(Codigo), "000000")   ' Medidor
'''                 Case 2: TipoDoc = Codigo  ' Consumo
'''                 Case 3: Real1 = Redondear(Val(Codigo), 2)  ' Base
'''                 Case 9: Real2 = Redondear(Val(Codigo), 2)  ' Excedente
'''                 Case 10: Mifecha = TrimStrg(Codigo)           ' Fecha
'''                 Case 11: Real3 = Redondear(Val(Codigo), 2) ' Mora
'''                 Case 12: Real4 = Redondear(Val(Codigo), 2) ' Multa Sesiones
'''                 Case 13: Real5 = Redondear(Val(Codigo), 2) ' Reconecciones
'''                 Case 14: Real6 = Redondear(Val(Codigo), 2) ' Otros
'''                 Case 15: Real7 = Redondear(Val(Codigo), 2) ' Otros
'''                 Case 16: Real8 = Redondear(Val(Codigo), 2) ' Otros
'''               End Select
'''           Next N
'''
'''           If Val(TipoDoc) >= 0 And IsDate(Mifecha) Then
'''              NoMes = CInt(Month(Mifecha))
'''              MiMes = MesesLetras(NoMes)
'''              MiAnio = Year(Mifecha)
'''              CodigoCli = Ninguno
'''              If AdoClientes.Recordset.RecordCount > 0 Then
'''                 AdoClientes.Recordset.MoveFirst
'''                 AdoClientes.Recordset.Find ("Cuenta_No = '" & Cuenta_No & "' ")
'''                 If Not AdoClientes.Recordset.EOF Then
'''                    CodigoCli = AdoClientes.Recordset.Fields("Codigo")
'''                    Grupo_No = AdoClientes.Recordset.Fields("Grupo")
'''                 End If
'''              End If
'''              If CodigoCli <> Ninguno Then
'''                 If Real1 > 0 Then
'''                    SetAdoAddNew "Clientes_Facturacion"
'''                    SetAdoFields "T", Normal
'''                    SetAdoFields "Codigo", CodigoCli
'''                    SetAdoFields "Codigo_Inv", Cod1
'''                    SetAdoFields "Valor", Real1
'''                    SetAdoFields "GrupoNo", Grupo_No
'''                    SetAdoFields "Mes", MiMes
'''                    SetAdoFields "Num_Mes", NoMes
'''                    SetAdoFields "Periodo", MiAnio
'''                    SetAdoFields "Fecha", Mifecha
'''                    SetAdoFields "Mensaje", "Med. No. " & Cuenta_No & ", Cons. " & TipoDoc & "M3"
'''                    SetAdoFields "CodigoU", CodigoUsuario
'''                    SetAdoFields "Item", NumEmpresa
'''                    SetAdoFields "Credito_No", Cuenta_No
'''                    SetAdoUpdate
'''                 End If
'''                 If Real2 > 0 Then
'''                    SetAdoAddNew "Clientes_Facturacion"
'''                    SetAdoFields "T", Normal
'''                    SetAdoFields "Codigo", CodigoCli
'''                    SetAdoFields "Codigo_Inv", Cod2
'''                    SetAdoFields "Valor", Real2
'''                    SetAdoFields "GrupoNo", Grupo_No
'''                    SetAdoFields "Mes", MiMes
'''                    SetAdoFields "Num_Mes", NoMes
'''                    SetAdoFields "Periodo", MiAnio
'''                    SetAdoFields "Fecha", Mifecha
'''                    SetAdoFields "Mensaje", "Med. No. " & Cuenta_No
'''                    SetAdoFields "CodigoU", CodigoUsuario
'''                    SetAdoFields "Item", NumEmpresa
'''                    SetAdoFields "Credito_No", Cuenta_No
'''                    SetAdoUpdate
'''                 End If
'''                 If Real3 > 0 Then
'''                    SetAdoAddNew "Clientes_Facturacion"
'''                    SetAdoFields "T", Normal
'''                    SetAdoFields "Codigo", CodigoCli
'''                    SetAdoFields "Codigo_Inv", Cod3
'''                    SetAdoFields "Valor", Real3
'''                    SetAdoFields "GrupoNo", Grupo_No
'''                    SetAdoFields "Mes", MiMes
'''                    SetAdoFields "Num_Mes", NoMes
'''                    SetAdoFields "Periodo", MiAnio
'''                    SetAdoFields "Fecha", Mifecha
'''                    SetAdoFields "Mensaje", "Med. No. " & Cuenta_No
'''                    SetAdoFields "CodigoU", CodigoUsuario
'''                    SetAdoFields "Item", NumEmpresa
'''                    SetAdoFields "Credito_No", Cuenta_No
'''                    SetAdoUpdate
'''                 End If
'''                 If Real4 > 0 Then
'''                    SetAdoAddNew "Clientes_Facturacion"
'''                    SetAdoFields "T", Normal
'''                    SetAdoFields "Codigo", CodigoCli
'''                    SetAdoFields "Codigo_Inv", Cod4
'''                    SetAdoFields "Valor", Real4
'''                    SetAdoFields "GrupoNo", Grupo_No
'''                    SetAdoFields "Mes", MiMes
'''                    SetAdoFields "Num_Mes", NoMes
'''                    SetAdoFields "Periodo", MiAnio
'''                    SetAdoFields "Fecha", Mifecha
'''                    SetAdoFields "Mensaje", "Med. No. " & Cuenta_No
'''                    SetAdoFields "CodigoU", CodigoUsuario
'''                    SetAdoFields "Item", NumEmpresa
'''                    SetAdoFields "Credito_No", Cuenta_No
'''                    SetAdoUpdate
'''                 End If
'''                 If Real5 > 0 Then
'''                    SetAdoAddNew "Clientes_Facturacion"
'''                    SetAdoFields "T", Normal
'''                    SetAdoFields "Codigo", CodigoCli
'''                    SetAdoFields "Codigo_Inv", Cod5
'''                    SetAdoFields "Valor", Real5
'''                    SetAdoFields "GrupoNo", Grupo_No
'''                    SetAdoFields "Mes", MiMes
'''                    SetAdoFields "Num_Mes", NoMes
'''                    SetAdoFields "Periodo", MiAnio
'''                    SetAdoFields "Fecha", Mifecha
'''                    SetAdoFields "Mensaje", "Med. No. " & Cuenta_No
'''                    SetAdoFields "CodigoU", CodigoUsuario
'''                    SetAdoFields "Item", NumEmpresa
'''                    SetAdoFields "Credito_No", Cuenta_No
'''                    SetAdoUpdate
'''                 End If
'''                 If Real6 > 0 Then
'''                    SetAdoAddNew "Clientes_Facturacion"
'''                    SetAdoFields "T", Normal
'''                    SetAdoFields "Codigo", CodigoCli
'''                    SetAdoFields "Codigo_Inv", Cod6
'''                    SetAdoFields "Valor", Real6
'''                    SetAdoFields "GrupoNo", Grupo_No
'''                    SetAdoFields "Mes", MiMes
'''                    SetAdoFields "Num_Mes", NoMes
'''                    SetAdoFields "Periodo", MiAnio
'''                    SetAdoFields "Fecha", Mifecha
'''                    SetAdoFields "Mensaje", "Med. No. " & Cuenta_No
'''                    SetAdoFields "CodigoU", CodigoUsuario
'''                    SetAdoFields "Item", NumEmpresa
'''                    SetAdoFields "Credito_No", Cuenta_No
'''                    SetAdoUpdate
'''                 End If
'''                 If Real7 > 0 Then
'''                    SetAdoAddNew "Clientes_Facturacion"
'''                    SetAdoFields "T", Normal
'''                    SetAdoFields "Codigo", CodigoCli
'''                    SetAdoFields "Codigo_Inv", Cod7
'''                    SetAdoFields "Valor", Real7
'''                    SetAdoFields "GrupoNo", Grupo_No
'''                    SetAdoFields "Mes", MiMes
'''                    SetAdoFields "Num_Mes", NoMes
'''                    SetAdoFields "Periodo", MiAnio
'''                    SetAdoFields "Fecha", Mifecha
'''                    SetAdoFields "Mensaje", "Med. No. " & Cuenta_No
'''                    SetAdoFields "CodigoU", CodigoUsuario
'''                    SetAdoFields "Item", NumEmpresa
'''                    SetAdoFields "Credito_No", Cuenta_No
'''                    SetAdoUpdate
'''                 End If
'''                 If Real8 > 0 Then
'''                    SetAdoAddNew "Clientes_Facturacion"
'''                    SetAdoFields "T", Normal
'''                    SetAdoFields "Codigo", CodigoCli
'''                    SetAdoFields "Codigo_Inv", Cod8
'''                    SetAdoFields "Valor", Real8
'''                    SetAdoFields "GrupoNo", Grupo_No
'''                    SetAdoFields "Mes", MiMes
'''                    SetAdoFields "Num_Mes", NoMes
'''                    SetAdoFields "Periodo", MiAnio
'''                    SetAdoFields "Fecha", Mifecha
'''                    SetAdoFields "Mensaje", "Med. No. " & Cuenta_No
'''                    SetAdoFields "CodigoU", CodigoUsuario
'''                    SetAdoFields "Item", NumEmpresa
'''                    SetAdoFields "Credito_No", Cuenta_No
'''                    SetAdoUpdate
'''                 End If
'''               End If
'''           End If
'''           Me.Caption = "Importar de FlexGrid a Sistema de Facturacion El Numero: " & Cuenta_No & ": " & i & " de " & Rango.NumFila2
'''      Next i
'''  End With
'''  DGExcelAdodc.Visible = True
'''End Sub
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Sub Importar_Inventarios()
Dim I As Long
Dim N As Long
Dim Precio_2 As Currency
Dim Precio_3 As Currency
Dim Cod_IESS As String
Dim Cod_R_RES As String
Dim Marca As String
Dim Tot_Propinas As Currency
Dim Servicio As Boolean
Dim Unidad As String
Dim InvMin As Currency
Dim InvMax As Currency
Dim Detalle As String
Dim Ubicacion As String
Dim Consignacion As Boolean
  DGExcelAdodc.Visible = False
  Bandera = False
  Evaluar = True
  Progreso_Barra.Incremento = 0
 'Empezamos la importacion de las facturas
  With AdoExcelAdodc.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       Do While Not .EOF
          CodigoInv = Dato_Campo(.fields(1))
          If CodigoInv = "" Then CodigoInv = Ninguno
          sSQL = "DELETE * " _
               & "FROM Catalogo_Productos " _
               & "WHERE Item = '" & NumEmpresa & "' " _
               & "AND Periodo = '" & Periodo_Contable & "' " _
               & "AND Codigo_Inv = '" & CodigoInv & "' "
          Ejecutar_SQL_SP sSQL
         .MoveNext
       Loop
   End If
  End With
  
 'Empezamos la importacion de las facturas
   With AdoExcelAdodc.Recordset
    If .RecordCount > 0 Then
        Progreso_Barra.Valor_Maximo = .RecordCount + 100
       .MoveFirst
        Do While Not .EOF
           For IdField = 0 To .fields.Count - 1
               Codigo = Dato_Campo(.fields(IdField))
               Codigo1 = Codigo
               Select Case IdField + 1
                 Case 1: TipoCta = Codigo
                         If TipoCta <> "P" Then TipoCta = "I"
                 Case 2: CodigoInv = Codigo
                 Case 3: Producto = Codigo
                 Case 4: Precio = Redondear(Val(Codigo), Dec_PVP)
                 Case 5: CodigoB = TrimStrg(Replace(Codigo, "/", ""))
                 Case 6: Cta_Inventario = Codigo
                         If Len(Cta_Inventario) <= 1 Then Cta_Inventario = "0"
                 Case 7: Cta_Costo_Ventas = Codigo
                         If Len(Cta_Costo_Ventas) <= 1 Then Cta_Costo_Ventas = "0"
                 Case 8: Cta_Ventas = Codigo
                         If Len(Cta_Ventas) <= 1 Then Cta_Ventas = "0"
                 Case 10: Si_No = CBool(Val(Codigo))
                 Case 11: Nuevo = CBool(Val(Codigo))
                 Case 12: Cantidad = Val(Codigo)
                 Case 13: Valor = Redondear(Val(Codigo), Dec_Costo)
                 Case 14: Mifecha = Codigo
                 Case 15: Numero = Val(Codigo)
                 Case 16: Precio_2 = Redondear(Val(Codigo), Dec_Costo)
                 Case 17: Precio_3 = Redondear(Val(Codigo), Dec_Costo)
                 Case 18: Cod_R_RES = Codigo
                 Case 19: Cod_IESS = Codigo
                 Case 20: Marca = Codigo
                 Case 21: Servicio = Val(Codigo)
                 Case 22: Unidad = Codigo
                 Case 23: InvMax = Val(Codigo)
                 Case 24: InvMin = Val(Codigo)
                 Case 26: Detalle = Codigo
                 Case 27: Ubicacion = Codigo
                 Case 28: Consignacion = CBool(Val(Codigo))
               End Select
           Next IdField
          'MsgBox "Cod. Final: " & Cantidad
           SetAdoAddNew "Catalogo_Productos"
           SetAdoFields "TC", TipoCta
           SetAdoFields "Codigo_Inv", CodigoInv
           SetAdoFields "Codigo_IESS", Cod_IESS
           SetAdoFields "Codigo_RES", Cod_R_RES
           SetAdoFields "Producto", Producto
           If IsNumeric(Unidad) Then
              If Val(Unidad) > 1 Then
                 Precio = Redondear(Precio / Val(Unidad), 4)
                 Precio_2 = Redondear(Precio_2 / Val(Unidad), 4)
                 Precio_3 = Redondear(Precio_3 / Val(Unidad), 4)
                 Valor = Redondear(Valor / Val(Unidad), 4)
              End If
           End If
           SetAdoFields "PVP", Precio
           SetAdoFields "PVP_2", Precio_2
           SetAdoFields "PVP_3", Precio_3
           SetAdoFields "Marca", Marca
           SetAdoFields "Unidad", Unidad
           SetAdoFields "Servicio", Servicio
           SetAdoFields "Minimo", InvMin
           SetAdoFields "Maximo", InvMax
           If TipoCta = "P" Then
              SetAdoFields "Codigo_Barra", CodigoB
              SetAdoFields "Cta_Inventario", Cta_Inventario
              SetAdoFields "Cta_Costo_Venta", Cta_Costo_Ventas
              SetAdoFields "Cta_Ventas", Cta_Ventas
              SetAdoFields "Cta_Ventas_0", Cta_Ventas
              SetAdoFields "Cta_Ventas_Ant", Cta_Ventas
              SetAdoFields "IVA", Si_No
              SetAdoFields "INV", Nuevo
              SetAdoFields "Detalle", Detalle
              SetAdoFields "Ubicacion", Ubicacion
              SetAdoFields "Consignacion", Consignacion
           End If
           SetAdoFields "Item", NumEmpresa
           SetAdoUpdate
          'If Cantidad <> 0 Then MsgBox TipoCta
           If TipoCta = "P" And Numero <> 0 And Cantidad <> 0 Then
              If IsNumeric(Unidad) Then
                 If Val(Unidad) > 1 Then Cantidad = Redondear(Cantidad * Val(Unidad), 2)
              End If
                SetAdoAddNew "Trans_Kardex"
                SetAdoFields "T", Normal
                SetAdoFields "CodBodega", "01"
                SetAdoFields "Codigo_Inv", CodigoInv
                SetAdoFields "Entrada", Cantidad
                SetAdoFields "TP", "CD"
                SetAdoFields "Numero", Numero
                SetAdoFields "PVP", Precio
                SetAdoFields "Valor_Unitario", Valor
                SetAdoFields "Valor_Total", Redondear(Cantidad * Valor, 2)
                SetAdoFields "Costo", Valor
                SetAdoFields "Total", Redondear(Cantidad * Valor, 2)
                SetAdoFields "Codigo_Barra", CodigoB
                SetAdoFields "Cta_Inv", Cta_Inventario
                SetAdoFields "Contra_Cta", Cta_Costo_Ventas
                SetAdoFields "Fecha", Mifecha
                SetAdoFields "Item", NumEmpresa
                SetAdoUpdate
           End If
           Progreso_Barra.Mensaje_Box = "Grabando al Sistema: " & CodigoInv & " - " & Producto
           Progreso_Esperar
           'MsgBox "..."
          .MoveNext
        Loop
    End If
   End With
   Progreso_Barra.Mensaje_Box = "Reindexandos Saldos"
   Progreso_Esperar
   sSQL = "UPDATE Trans_Kardex " _
        & "SET Procesado = 0 " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND Periodo = '" & Periodo_Contable & "' "
   Ejecutar_SQL_SP sSQL

   sSQL = "UPDATE Transacciones " _
        & "SET Procesado = 0 " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND Periodo = '" & Periodo_Contable & "' "
   Ejecutar_SQL_SP sSQL

   Progreso_Final
   DGExcelAdodc.Visible = True
End Sub
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'''Public Sub Importar_Facturas_Contabilidad()
'''Dim I As Long
'''Dim N As Long
'''Dim Tot_Propinas As Currency
''' 'Empezamos la importacion de las facturas
'''  NumTrans = 0
'''  DGExcelAdodc.Visible = False
'''  With AdoExcelAdodc.Recordset
'''      'MsgBox .Rows & vbCrLf & .Cols
'''       For i = 1 To .rows - 1
'''          .Row = i
'''           CodigoCli = "9999999999"
'''           NombreCliente = "CONSUMIDOR FINAL"
'''           For N = 1 To .cols - 1
'''              .Col = N
'''               Codigo = UCaseStrg(Replace(.Text, "'", ""))
'''               Codigo = TrimStrg(Replace(Codigo, "-", ""))
'''               Codigo3 = TrimStrg(Codigo)
'''               Select Case N
'''                 Case 2: 'RUC/Cedula/Consumidor Final
'''                          If Len(Codigo) > 1 Then
'''                             DigVerif = Digito_Verificador( Codigo)
'''                             CodigoCli = Tipo_RUC_CI.Codigo_RUC_CI
'''                          End If
'''                          CI_Representante = Codigo
'''                          NombreCliente = Codigo
'''                         'MsgBox "Digito Verif. " & DigVerif & " (" & Tipo_RUC_CI.Tipo_Beneficiario & ")"
'''                 Case 3: TipoDoc = Codigo
'''                 Case 4: Cantidad = Val(Codigo)
'''                 Case 5: Total_Sin_IVA = Abs(Val(Codigo))
'''                 Case 6: Total_Con_IVA = Abs(Val(Codigo))
'''                 Case 7: Total_Sin_No_IVA = Abs(Val(Codigo))
'''                 Case 8: Total_IVA = Abs(Val(Codigo))
'''                 Case 9: Total_RetIVA = Abs(Val(Codigo))
'''                 Case 10: Total_Ret = Abs(Val(Codigo))
'''                 Case 11: Real1 = Val(Codigo)
'''                          PorcIVAB = 0
'''                          PorcIVAS = 0
'''                          If (0 < Real1) And (Real1 <= 30) Then
'''                             PorcIVAB = Real1
'''                          Else
'''                             PorcIVAS = Real1
'''                          End If
'''                 Case 12: Porc = Val(Codigo)
'''                 Case 13: Mifecha = Codigo
'''               End Select
'''           Next N
'''           sSQL = "SELECT * " _
'''                & "FROM Clientes " _
'''                & "WHERE Codigo = '" & CodigoCli & "' "
'''           Select_Adodc AdoAux, sSQL
'''           If AdoAux.Recordset.RecordCount <= 0 Then
'''              'MsgBox "CLI: " & NombreCliente & vbCrLf & "COD: " & CodigoCli & vbCrLf & "CI/RUC: " & CI_Representante & vbCrLf & "TB: " & Tipo_RUC_CI.Tipo_Beneficiario
'''              SetAdoAddNew "Clientes"
'''              SetAdoFields "T", Normal
'''              SetAdoFields "Codigo", CodigoCli
'''              SetAdoFields "TD", Tipo_RUC_CI.Tipo_Beneficiario
'''              SetAdoFields "CI_RUC", CI_Representante
'''              SetAdoFields "Cliente", UCaseStrg(NombreCliente)
'''              SetAdoFields "Representante", UCaseStrg(NombreCliente)
'''              SetAdoFields "Cedula", CI_Representante
'''              SetAdoFields "CI_RUC_SRI", CI_Representante
'''              SetAdoFields "TD_SRI", Tipo_RUC_CI.Tipo_Beneficiario
'''              SetAdoFields "Fecha", FechaSistema
'''              SetAdoFields "Direccion", "SD"
'''              SetAdoFields "DirNumero", "SN"
'''              SetAdoFields "Ciudad", "QUITO"
'''              SetAdoFields "Grupo", "ATS"
'''              SetAdoFields "Prov", "17"
'''              SetAdoFields "Pais", "593"
'''              SetAdoFields "FA", True
'''              SetAdoFields "CodigoU", CodigoUsuario
'''              SetAdoUpdate
'''           End If
'''           sSQL = "DELETE * " _
'''                & "FROM Trans_Ventas " _
'''                & "WHERE Item = '" & NumEmpresa & "' " _
'''                & "AND Periodo = '" & Periodo_Contable & "' " _
'''                & "AND Fecha = #" & BuscarFecha(Mifecha) & "# " _
'''                & "AND IdProv = '" & CodigoCli & "' " _
'''                & "AND TipoComprobante = " & TipoDoc & " "
'''           Ejecutar_SQL_SP sSQL
'''           sSQL = "DELETE * " _
'''                & "FROM Trans_Air " _
'''                & "WHERE Item = '" & NumEmpresa & "' " _
'''                & "AND Periodo = '" & Periodo_Contable & "' " _
'''                & "AND Fecha = #" & BuscarFecha(Mifecha) & "# " _
'''                & "AND IdProv = '" & CodigoCli & "' " _
'''                & "AND Tipo_Trans = 'V' "
'''           Ejecutar_SQL_SP sSQL
'''           FechaFinal = Mifecha
'''           Numero = 0
'''           Codigo1 = "001"
'''           Codigo2 = "001"
'''           Autorizacion = "9999999999"
'''           SubTotal = Total_Sin_IVA + Total_Con_IVA
'''           Insertar_Ventas CodigoCli, CLng(Cantidad), Total_Sin_IVA, Total_Con_IVA, Total_IVA, SubTotal, TipoDoc
'''           If Total_Ret > 0 Then Insertar_Ventas_Air CodigoCli, SubTotal, Porc, Total_Ret, CLng(Cantidad), 99999999, "001", "001", "9999999999", Ninguno
'''           NumTrans = NumTrans + 1
'''           Me.Caption = "Importar de FlexGrid a Sistema a Anexos Transaccionales: " & Mifecha & ": " & i & " de " & Rango.NumFila2
'''      Next i
'''  End With
'''  Me.Caption = "Importar de FlexGrid a Sistema a Anexos Transaccionales: " & Mifecha & ": " & i & " de " & Rango.NumFila2
'''  DGExcelAdodc.Visible = True
'''End Sub
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'''Public Sub Cambio_Numero_Secuencial()
'''Dim I As Long
'''Dim N As Long
'''Dim Tot_Propinas As Currency
''' 'Empezamos la importacion de las facturas
'''  sSQL = "SELECT Codigo,Cliente,CI_RUC " _
'''       & "FROM Clientes " _
'''       & "WHERE Codigo <> '.' " _
'''       & "ORDER BY Cliente "
'''  Select_Adodc AdoClientes, sSQL
'''
'''  NumTrans = 0
'''  DGExcelAdodc.Visible = False
'''  With AdoExcelAdodc.Recordset
'''      'MsgBox .Rows & vbCrLf & .Cols
'''       For i = 1 To .rows - 1
'''          .Row = i
'''           For N = 1 To .cols - 1
'''              .Col = N
'''               Codigo = UCaseStrg(Replace(.Text, "'", ""))
'''               Codigo = TrimStrg(Replace(Codigo, "-", ""))
'''               Select Case N
'''                 Case 1: NombreCliente = Codigo
'''                 Case 2: SerieFactura = Codigo
'''                 Case 3: Autorizacion = Codigo
'''                 Case 4: Factura_Desde = Val(Codigo)
'''                 Case 5: Codigo1 = Codigo
'''                 Case 6: Factura_Hasta = Val(Codigo)
'''               End Select
'''           Next N
'''           CodigoN = Ninguno
'''           CodigoA = Ninguno
'''           If AdoClientes.Recordset.RecordCount > 0 Then
'''              AdoClientes.Recordset.MoveFirst
'''              AdoClientes.Recordset.Find ("Cliente = '" & NombreCliente & "' ")
'''              If Not AdoClientes.Recordset.EOF Then
'''                 CodigoN = AdoClientes.Recordset.Fields("Codigo")
'''                 Codigo1 = AdoClientes.Recordset.Fields("CI_RUC")
'''              End If
'''           End If
'''           sSQL = "SELECT * " _
'''                & "FROM Facturas " _
'''                & "WHERE Factura = " & Factura_Desde & " " _
'''                & "AND Serie = '" & SerieFactura & "' " _
'''                & "AND Autorizacion = '" & Autorizacion & "' " _
'''                & "AND Item = '" & NumEmpresa & "' " _
'''                & "AND Periodo = '" & Periodo_Contable & "' "
'''           Select_Adodc AdoAux, sSQL
'''           If AdoAux.Recordset.RecordCount > 0 Then
'''              CodigoA = AdoAux.Recordset.Fields("CodigoC")
'''           End If
'''           sSQL = "SELECT * " _
'''                & "FROM Facturas " _
'''                & "WHERE Factura = " & Factura_Hasta & " " _
'''                & "AND Item = '" & NumEmpresa & "' " _
'''                & "AND Periodo = '" & Periodo_Contable & "' "
'''           Select_Adodc AdoAux, sSQL
'''           If AdoAux.Recordset.RecordCount > 0 Then
'''              CodigoCli = AdoAux.Recordset.Fields("CodigoC")
'''           End If
'''           If CodigoN <> Ninguno And CodigoCli <> Ninguno Then
'''              sSQL = "UPDATE Facturas " _
'''                   & "SET CodigoC = '" & CodigoN & "' " _
'''                   & "WHERE Item = '" & NumEmpresa & "' " _
'''                   & "AND Periodo = '" & Periodo_Contable & "' " _
'''                   & "AND Factura = " & Factura_Hasta & " " _
'''                   & "AND CodigoC = '" & CodigoCli & "' "
'''              Ejecutar_SQL_SP sSQL
'''
'''              sSQL = "UPDATE Detalle_Factura " _
'''                   & "SET CodigoC = '" & CodigoN & "' " _
'''                   & "WHERE Item = '" & NumEmpresa & "' " _
'''                   & "AND Periodo = '" & Periodo_Contable & "' " _
'''                   & "AND Factura = " & Factura_Hasta & " " _
'''                   & "AND CodigoC = '" & CodigoCli & "' "
'''              Ejecutar_SQL_SP sSQL
'''
'''              sSQL = "UPDATE Trans_Abonos " _
'''                   & "SET CodigoC = '" & CodigoN & "' " _
'''                   & "WHERE Item = '" & NumEmpresa & "' " _
'''                   & "AND Periodo = '" & Periodo_Contable & "' " _
'''                   & "AND Factura = " & Factura_Hasta & " " _
'''                   & "AND CodigoC = '" & CodigoCli & "' "
'''              Ejecutar_SQL_SP sSQL
'''           End If
'''           Me.Caption = "Actualizando Nombre Alumnos: " & i & " de " & Rango.NumFila2
'''      Next i
'''  End With
'''  Me.Caption = "Actualizando Nombre Alumnos: " & i & " de " & Rango.NumFila2
'''  DGExcelAdodc.Visible = True
'''End Sub
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'''Public Sub Importar_Notas_Materias()
'''Dim F As Long
'''Dim N As Long
'''Dim E As Byte
'''Dim Nota_Alumno As Currency
'''Dim Nota_TAI As Currency
'''Dim Nota_AIC As Currency
'''Dim Nota_AGC As Currency
'''Dim Nota_L As Currency
'''Dim Nota_P As Currency
'''Dim Nota_Suma As Currency
'''Dim Nota_Prom As Currency
'''Dim Nota_Prom1 As Currency
'''Dim Nota_Prom2 As Currency
'''Dim Nota_Prom3 As Currency
'''Dim Nota_ExaP As Currency
'''Dim Porc_Nota As Single
'''Dim Dias_L As Integer
'''Dim ExamenQuimestre As Boolean
'''Dim ExamenSupletorio As Boolean
'''Dim ExamenRemedial As Boolean
'''Dim EncontroNota As Boolean
'''   RatonReloj
'''   DGExcelAdodc.Visible = False
'''   ExamenQuimestre = False
'''   ExamenSupletorio = False
'''   ExamenRemedial = False
'''   CodMatP = Ninguno
'''   SQLNotas = ""
'''   SQLTAI = ""
'''   SQLAIC = ""
'''   SQLAGC = ""
'''   SQLL = ""
'''
'''   sSQL = "SELECT * " _
'''        & "FROM Catalogo_Periodo_Lectivo " _
'''        & "WHERE Item = '" & NumEmpresa & "' " _
'''        & "AND Periodo = '" & Periodo_Contable & "' "
'''   Select_Adodc AdoAux, sSQL
'''   With AdoAux.Recordset
'''    If .RecordCount > 0 Then
'''        Asistencias = .Fields("Asistencias")
'''        If MidStrg(FormatoLibreta, 1, 9) = "QUIMESTRE" Then
'''           If .Fields("NPQP1") Then
'''               SQLTAI = "PQTAI1"
'''               SQLAIC = "PQAIC1"
'''               SQLAGC = "PQAGC1"
'''               SQLL = "PQL1"
'''               SQLProm = "PQBim1"
'''               SQLExaP = "PQExaP1"
'''               SQLNotas = "PQBim1"
'''
'''               Evaluacion = "ConductaPQ1"
'''               SQLDias = "PQDias1"
'''               SQLFJ = "PQBFJ1"
'''               SQLFI = "PQBFI1"
'''               SQLAtrasos = "PQBA1"
'''               Porc_Nota = 0.3
'''           End If
'''           If .Fields("NPQP2") Then
'''               SQLTAI = "PQTAI2"
'''               SQLAIC = "PQAIC2"
'''               SQLAGC = "PQAGC2"
'''               SQLL = "PQL2"
'''               SQLProm = "PQBim2"
'''               SQLExaP = "PQExaP2"
'''               SQLNotas = "PQBim2"
'''
'''               Evaluacion = "ConductaPQ2"
'''               SQLDias = "PQDias2"
'''               SQLFJ = "PQBFJ2"
'''               SQLFI = "PQBFI2"
'''               SQLAtrasos = "PQBA2"
'''
'''               Porc_Nota = 0.25
'''           End If
'''           If .Fields("NPQP3") Then
'''               SQLTAI = "PQTAI3"
'''               SQLAIC = "PQAIC3"
'''               SQLAGC = "PQAGC3"
'''               SQLL = "PQL3"
'''               SQLProm = "PQBim3"
'''               SQLExaP = "PQExaP3"
'''               SQLNotas = "PQBim3"
'''
'''               Evaluacion = "ConductaPQ3"
'''               SQLDias = "PQDias3"
'''               SQLFJ = "PQBFJ3"
'''               SQLFI = "PQBFI3"
'''               SQLAtrasos = "PQBA3"
'''
'''               Porc_Nota = 0.25
'''           End If
'''           If .Fields("NPQEX") Then
'''               SQLTAI = "XXX"
'''               SQLAIC = "XXX"
'''               SQLAGC = "XXX"
'''               SQLL = "XXX"
'''               SQLProm = "PromPQ"
'''               SQLProm1 = "PQBim1"
'''               SQLProm2 = "PQBim2"
'''               SQLProm3 = "PQBim3"
'''               SQLExaP = "ExamenPQ"
'''               SQLNotas = "XXX"
'''
'''               Evaluacion = "XXX"
'''               SQLDias = "XXX"
'''               SQLFJ = "XXX"
'''               SQLFI = "XXX"
'''               SQLAtrasos = "XXX"
'''
'''               Porc_Nota = 0.25
'''               ExamenQuimestre = True
'''           End If
'''           If .Fields("NSQP1") Then
'''               SQLTAI = "SQTAI1"
'''               SQLAIC = "SQAIC1"
'''               SQLAGC = "SQAGC1"
'''               SQLL = "SQL1"
'''               SQLProm = "SQBim1"
'''               SQLExaP = "SQExaP1"
'''               SQLNotas = "SQBim1"
'''
'''               Evaluacion = "ConductaSQ1"
'''               SQLDias = "SQDias1"
'''               SQLFJ = "SQBFJ1"
'''               SQLFI = "SQBFI1"
'''               SQLAtrasos = "SQBA1"
'''
'''               Porc_Nota = 0.3
'''           End If
'''           If .Fields("NSQP2") Then
'''               SQLTAI = "SQTAI2"
'''               SQLAIC = "SQAIC2"
'''               SQLAGC = "SQAGC2"
'''               SQLL = "SQL2"
'''               SQLProm = "SQBim2"
'''               SQLExaP = "SQExaP2"
'''               SQLNotas = "SQBim2"
'''
'''               Evaluacion = "ConductaSQ2"
'''               SQLDias = "SQDias2"
'''               SQLFJ = "SQBFJ2"
'''               SQLFI = "SQBFI2"
'''               SQLAtrasos = "SQBA2"
'''
'''               Porc_Nota = 0.25
'''           End If
'''           If .Fields("NSQP3") Then
'''               SQLTAI = "SQTAI3"
'''               SQLAIC = "SQAIC3"
'''               SQLAGC = "SQAGC3"
'''               SQLL = "SQL3"
'''               SQLProm = "SQBim3"
'''               SQLExaP = "SQExaP3"
'''               SQLNotas = "SQBim3"
'''
'''               Evaluacion = "ConductaSQ3"
'''               SQLDias = "SQDias3"
'''               SQLFJ = "SQBFJ3"
'''               SQLFI = "SQBFI3"
'''               SQLAtrasos = "SQBA3"
'''
'''               Porc_Nota = 0.25
'''           End If
'''           If .Fields("NSQEX") Then
'''               SQLTAI = "XXX"
'''               SQLAIC = "XXX"
'''               SQLAGC = "XXX"
'''               SQLL = "XXX"
'''               SQLProm = "PromSQ"
'''               SQLProm1 = "SQBim1"
'''               SQLProm2 = "SQBim2"
'''               SQLProm3 = "SQBim3"
'''               SQLExaP = "ExamenSQ"
'''               SQLNotas = "XXX"
'''
'''               Evaluacion = "XXX"
'''               SQLDias = "XXX"
'''               SQLFJ = "XXX"
'''               SQLFI = "XXX"
'''               SQLAtrasos = "XXX"
'''
'''               Porc_Nota = 0.25
'''               ExamenQuimestre = True
'''           End If
'''        Else
'''           SQLTAI = "X"
'''           SQLAIC = "X"
'''           SQLAGC = "X"
'''           SQLL = "X"
'''           If .Fields("NPQP1") Then SQLNotas = "PQBim1"
'''           If .Fields("NPQP2") Then SQLNotas = "PQBim2"
'''           If .Fields("NPQP3") Then SQLNotas = "PQBim3"
'''           If .Fields("NPQEX") Then SQLNotas = "ExamenPQ"
'''
'''           If .Fields("NSQP1") Then SQLNotas = "SQBim1"
'''           If .Fields("NSQP2") Then SQLNotas = "SQBim2"
'''           If .Fields("NSQP3") Then SQLNotas = "SQBim3"
'''           If .Fields("NSQEX") Then SQLNotas = "ExamenSQ"
'''
'''           If .Fields("NTQP1") Then SQLNotas = "TQBim1"
'''           If .Fields("NTQP2") Then SQLNotas = "TQBim2"
'''           If .Fields("NTQP3") Then SQLNotas = "TQBim3"
'''           If .Fields("NTQEX") Then SQLNotas = "ExamenTQ"
'''        End If
'''        If .Fields("NSUPL") Then
'''            SQLTAI = "XXX"
'''            SQLAIC = "XXX"
'''            SQLAGC = "XXX"
'''            SQLL = "XXX"
'''            SQLProm = "PromSQ"
'''            SQLProm1 = "SQBim1"
'''            SQLProm2 = "SQBim2"
'''            SQLProm3 = "SQBim3"
'''            SQLExaP = "ExamenSQ"
'''
'''            Evaluacion = "XXX"
'''            SQLDias = "XXX"
'''            SQLFJ = "XXX"
'''            SQLFI = "XXX"
'''            SQLAtrasos = "XXX"
'''
'''            SQLNotas = "Supletorio"
'''            ExamenSupletorio = True
'''        End If
'''        If .Fields("NREME") Then
'''            SQLTAI = "XXX"
'''            SQLAIC = "XXX"
'''            SQLAGC = "XXX"
'''            SQLL = "XXX"
'''            SQLProm = "PromSQ"
'''            SQLProm1 = "SQBim1"
'''            SQLProm2 = "SQBim2"
'''            SQLProm3 = "SQBim3"
'''            SQLExaP = "ExamenSQ"
'''
'''            Evaluacion = "XXX"
'''            SQLDias = "XXX"
'''            SQLFJ = "XXX"
'''            SQLFI = "XXX"
'''            SQLAtrasos = "XXX"
'''
'''            SQLNotas = "Remedial"
'''            ExamenRemedial = True
'''        End If
'''        If .Fields("NGRADO") Then SQLNotas = ""
'''    End If
'''   End With
'''  With AdoExcelAdodc.Recordset
'''     'Empezamos la importacion de las facturas
'''       For F = 1 To .rows - 1
'''          .Row = F
'''           Dias_L = 0
'''            For N = 1 To .cols - 1
'''               .Col = N
'''                Codigo = TrimStrg(Replace(.Text, "'", ""))
'''               'MsgBox Codigo & "...."
'''                EncontroNota = True
'''                For E = 0 To UBound(Equivalencias) - 1
'''                    If EncontroNota And Equivalencias(E).Letras = UCaseStrg(Codigo) Then
'''                       Codigo = CStr(Equivalencias(E).Hasta)
'''                       EncontroNota = False
'''                    End If
'''                    If EncontroNota And Equivalencias(E).Cualitativa = UCaseStrg(Codigo) Then
'''                       Codigo = CStr(Equivalencias(E).Hasta)
'''                       EncontroNota = False
'''                    End If
'''                    If EncontroNota And Equivalencias(E).Cualitativa2 = UCaseStrg(Codigo) Then
'''                       Codigo = CStr(Equivalencias(E).Hasta)
'''                       EncontroNota = False
'''                    End If
'''                Next E
'''
'''                Select Case N
'''                  Case 1: NombreCliente = Codigo
'''                  Case 2: Nota_TAI = Format$(Val(Codigo), "00.00")
'''                          Nota_Alumno = Format$(Val(Codigo), "00.00")
'''                          If Nota_TAI > 10 Then Nota_TAI = 10
'''                          If Nota_Alumno > 20 Then Nota_Alumno = 20
'''                  Case 3: Nota_AIC = Format$(Val(Codigo), "00.00")
'''                          Dias_L = CInt(Nota_AIC)
'''                          If Nota_AIC > 10 Then Nota_AIC = 10
'''                  Case 4: Nota_AGC = Format$(Val(Codigo), "00.00")
'''                          If Nota_AGC > 10 Then Nota_AGC = 10
'''                  Case 5: Nota_L = Format$(Val(Codigo), "00.00")
'''                          If Nota_L > 10 Then Nota_L = 10
'''                  Case 6: Nota_ExaP = Format$(Val(Codigo), "00.00")
'''                          If Nota_ExaP > 10 Then Nota_ExaP = 10
'''                          Dias_L = Val(Codigo)
'''                  Case 7: Faltas_Just = Val(Codigo)
'''                  Case 8: Faltas_Injust = Val(Codigo)
'''                  Case 9: Atrasos = Val(Codigo)
'''                  Case 10: CodMat = Codigo       '10
'''                  Case 11: CodigoCli = Codigo
'''                  Case 12: TipoCta = Codigo
'''                  Case 13: CodigoP = Codigo
'''                End Select
'''            Next N
'''            'If F = 1 Then
'''               sSQL = "SELECT * " _
'''                    & "FROM Catalogo_Estudiantil " _
'''                    & "WHERE Item = '" & NumEmpresa & "' " _
'''                    & "AND Periodo = '" & Periodo_Contable & "' " _
'''                    & "AND CodMat = '" & CodMat & "' " _
'''                    & "AND MidStrg(CodigoE,1," & Len(CodigoP) & ") = '" & CodigoP & "' "
'''               Select_Adodc AdoAux, sSQL
'''               If AdoAux.Recordset.RecordCount > 0 Then CodMatP = AdoAux.Recordset.Fields("CodMatP")
'''            'End If
'''           'MsgBox sSQL
'''           'MsgBox NombreCliente & vbCrLf & Nota_TAI
'''
'''            If SQLTAI <> "" And SQLAIC <> "" And SQLAGC <> "" And SQLL <> "" And SQLNotas <> "" _
'''               And (Nota_TAI + Nota_AIC + Nota_AGC + Nota_L + Nota_ExaP) > 0 Then
'''               If MidStrg(CodigoP, 1, 4) <= "1.01" Then
'''                  Nota_AIC = Nota_TAI
'''                  Nota_AGC = Nota_TAI
'''                  Nota_L = Nota_TAI
'''                  Nota_ExaP = Nota_TAI
'''               End If
'''              'Si las notas estan en cero colocamos la anterior
'''               Nota_Suma = Nota_TAI + Nota_AIC + Nota_AGC + Nota_L + Nota_ExaP
'''               Nota_Prom = Redondear(Nota_Suma / 5, 2)
'''               Nota_P = Redondear(Nota_ExaP * Porc_Nota, 2)
'''               If MidStrg(FormatoLibreta, 1, 9) = "QUIMESTRE" Then
'''                  If ExamenQuimestre Then
'''                     If CodMatP = Ninguno Then
'''                        sSQL = "UPDATE Trans_Notas SET "
'''                     Else
'''                        sSQL = "UPDATE Trans_Notas_Auxiliares SET "
'''                     End If
'''                     sSQL = sSQL _
'''                          & SQLExaP & " = " & Nota_TAI & "," _
'''                          & SQLProm & " = ROUND((" & SQLProm1 & "+" & SQLProm2 & "+" & SQLProm3 & ")/3,2,0) " _
'''                          & "WHERE Item = '" & NumEmpresa & "' " _
'''                          & "AND Periodo = '" & Periodo_Contable & "' " _
'''                          & "AND CodMat = '" & CodMat & "' " _
'''                          & "AND CodE = '" & CodigoP & "' " _
'''                          & "AND Codigo = '" & CodigoCli & "' "
'''                     Ejecutar_SQL_SP sSQL
'''                  ElseIf ExamenSupletorio Then
'''                     If CodMatP = Ninguno Then
'''                        sSQL = "UPDATE Trans_Notas SET "
'''                     Else
'''                        sSQL = "UPDATE Trans_Notas_Auxiliares SET "
'''                     End If
'''                     sSQL = sSQL _
'''                          & SQLNotas & " = " & Nota_TAI & " " _
'''                          & "WHERE Item = '" & NumEmpresa & "' " _
'''                          & "AND Periodo = '" & Periodo_Contable & "' " _
'''                          & "AND CodMat = '" & CodMat & "' " _
'''                          & "AND CodE = '" & CodigoP & "' " _
'''                          & "AND Codigo = '" & CodigoCli & "' "
'''                     Ejecutar_SQL_SP sSQL
'''                  ElseIf ExamenRemedial Then
'''                     If CodMatP = Ninguno Then
'''                        sSQL = "UPDATE Trans_Notas SET "
'''                     Else
'''                        sSQL = "UPDATE Trans_Notas_Auxiliares SET "
'''                     End If
'''                     sSQL = sSQL _
'''                          & SQLNotas & " = " & Nota_TAI & " " _
'''                          & "WHERE Item = '" & NumEmpresa & "' " _
'''                          & "AND Periodo = '" & Periodo_Contable & "' " _
'''                          & "AND CodMat = '" & CodMat & "' " _
'''                          & "AND CodE = '" & CodigoP & "' " _
'''                          & "AND Codigo = '" & CodigoCli & "' "
'''                     Ejecutar_SQL_SP sSQL
'''                  Else
'''                     sSQL = "UPDATE Trans_Asistencia SET " _
'''                          & Evaluacion & " = " & Nota_TAI & ", " _
'''                          & SQLDias & " = " & Dias_L & ", " _
'''                          & SQLFJ & " = " & Faltas_Just & ", " _
'''                          & SQLFI & " = " & Faltas_Injust & ", " _
'''                          & SQLAtrasos & " = " & Atrasos & " " _
'''                          & "WHERE Item = '" & NumEmpresa & "' " _
'''                          & "AND Periodo = '" & Periodo_Contable & "' " _
'''                          & "AND CodMat = '" & CodMat & "' " _
'''                          & "AND CodE = '" & CodigoP & "' " _
'''                          & "AND Codigo = '" & CodigoCli & "' "
'''                     Ejecutar_SQL_SP sSQL
'''
'''                     If CodMatP = Ninguno Then
'''                        sSQL = "UPDATE Trans_Notas SET "
'''                     Else
'''                        sSQL = "UPDATE Trans_Notas_Auxiliares SET "
'''                     End If
'''                     sSQL = sSQL _
'''                          & SQLTAI & " = " & Nota_TAI & ", " _
'''                          & SQLAIC & " = " & Nota_AIC & ", " _
'''                          & SQLAGC & " = " & Nota_AGC & ", " _
'''                          & SQLL & " = " & Nota_L & ", " _
'''                          & SQLProm & " = " & Nota_Prom & ", " _
'''                          & SQLExaP & " = " & Nota_ExaP & " " _
'''                          & "WHERE Item = '" & NumEmpresa & "' " _
'''                          & "AND Periodo = '" & Periodo_Contable & "' " _
'''                          & "AND CodMat = '" & CodMat & "' " _
'''                          & "AND CodE = '" & CodigoP & "' " _
'''                          & "AND Codigo = '" & CodigoCli & "' "
'''                     Ejecutar_SQL_SP sSQL
'''                  End If
'''               Else
'''                  If CodMatP = Ninguno Then
'''                     sSQL = "UPDATE Trans_Notas SET "
'''                  Else
'''                     sSQL = "UPDATE Trans_Notas_Auxiliares SET "
'''                  End If
'''                  sSQL = sSQL _
'''                       & SQLNotas & " = " & Nota_Alumno & " " _
'''                       & "WHERE Item = '" & NumEmpresa & "' " _
'''                       & "AND Periodo = '" & Periodo_Contable & "' " _
'''                       & "AND CodMat = '" & CodMat & "' " _
'''                       & "AND CodE = '" & CodigoP & "' " _
'''                       & "AND Codigo = '" & CodigoCli & "' "
'''                  Ejecutar_SQL_SP sSQL
'''               End If
'''           End If
'''           Me.Caption = "Importar de FlexGrid a Sistema de Notas - " & Format$(F / Rango.NumFila2, "00%")
'''      Next F
'''  End With
'''  DGExcelAdodc.Visible = True
'''  RatonNormal
'''End Sub
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'''Public Sub Importar_Informes_Materias()
'''Dim F As Long
'''Dim N As Long
'''Dim E As Byte
'''Dim Informe_Alumno As String
'''
'''   RatonReloj
'''   DGExcelAdodc.Visible = False
'''   CodMatP = Ninguno
'''   Evaluacion = ""
'''
'''   sSQL = "SELECT * " _
'''        & "FROM Catalogo_Periodo_Lectivo " _
'''        & "WHERE Item = '" & NumEmpresa & "' " _
'''        & "AND Periodo = '" & Periodo_Contable & "' "
'''   Select_Adodc AdoAux, sSQL
'''   With AdoAux.Recordset
'''    If .RecordCount > 0 Then
'''        If MidStrg(FormatoLibreta, 1, 9) = "QUIMESTRE" Then
'''           If .Fields("NPQP1") Then Evaluacion = "Informe_PQ1"
'''           If .Fields("NPQP2") Then Evaluacion = "Informe_PQ2"
'''           If .Fields("NPQP3") Then Evaluacion = "Informe_PQ3"
'''           If .Fields("NPQEX") Then Evaluacion = "Informe_PQ"
'''           If .Fields("NSQP1") Then Evaluacion = "Informe_SQ1"
'''           If .Fields("NSQP2") Then Evaluacion = "Informe_SQ2"
'''           If .Fields("NSQP3") Then Evaluacion = "Informe_SQ3"
'''           If .Fields("NSQEX") Then Evaluacion = "Informe_SQ"
'''           If .Fields("NSUPL") Then Evaluacion = ""
'''           If .Fields("NREME") Then Evaluacion = ""
'''        End If
'''    End If
'''   End With
'''  With AdoExcelAdodc.Recordset
'''     'Empezamos la importacion de las facturas
'''       For F = 1 To .rows - 1
'''          .Row = F
'''            For N = 1 To .cols - 1
'''               .Col = N
'''                Codigo = TrimStrg(Replace(.Text, "'", ""))
'''                Codigo = TrimStrg(Replace(Codigo, "/", " "))
'''                Codigo = TrimStrg(Replace(Codigo, "&", " "))
'''               'MsgBox Codigo & "...."
'''                Select Case N
'''                  Case 1: NombreCliente = Codigo
'''                  Case 2: Informe_Alumno = TrimStrg(MidStrg(Codigo, 1, 100))
'''                  Case 3: CodMat = Codigo
'''                  Case 4: CodigoCli = Codigo
'''                  Case 5: TipoCta = Codigo
'''                  Case 6: CodigoP = Codigo
'''                End Select
'''            Next N
'''            If F = 1 Then
'''               sSQL = "SELECT * " _
'''                    & "FROM Catalogo_Estudiantil " _
'''                    & "WHERE Item = '" & NumEmpresa & "' " _
'''                    & "AND Periodo = '" & Periodo_Contable & "' " _
'''                    & "AND CodMat = '" & CodMat & "' " _
'''                    & "AND MidStrg(CodigoE,1," & Len(CodigoP) & ") = '" & CodigoP & "' "
'''               Select_Adodc AdoAux, sSQL
'''               If AdoAux.Recordset.RecordCount > 0 Then CodMatP = AdoAux.Recordset.Fields("CodMatP")
'''               'MsgBox CodMatP & vbCrLf & sSQL
'''            End If
'''
'''            If Evaluacion <> "" And Len(Informe_Alumno) > 1 Then
'''               If MidStrg(FormatoLibreta, 1, 9) = "QUIMESTRE" Then
'''                  If CodMatP = Ninguno Then
'''                     sSQL = "UPDATE Trans_Notas SET "
'''                  Else
'''                     sSQL = "UPDATE Trans_Notas_Auxiliares SET "
'''                  End If
'''                  sSQL = sSQL _
'''                       & Evaluacion & " = '" & Informe_Alumno & "' " _
'''                       & "WHERE Item = '" & NumEmpresa & "' " _
'''                       & "AND Periodo = '" & Periodo_Contable & "' " _
'''                       & "AND CodMat = '" & CodMat & "' " _
'''                       & "AND CodE = '" & CodigoP & "' " _
'''                       & "AND Codigo = '" & CodigoCli & "' "
'''                  Ejecutar_SQL_SP sSQL
'''               End If
'''            End If
'''            Me.Caption = "Importar de FlexGrid a Sistema de Informes académicos - " & Format$(F / Rango.NumFila2, "00%")
'''      Next F
'''  End With
'''  DGExcelAdodc.Visible = True
'''  RatonNormal
'''End Sub
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'''Public Sub Transferir_Plan_Cuentas()
'''Dim F As Long
'''Dim N As Long
'''Dim Clave As Long
'''Dim Parpadear As Boolean
'''  RatonReloj
'''  Parpadear = True
'''  Clave = 0
'''  With AdoExcelAdodc.Recordset
'''      'Empezamos la importacion de las facturas
'''       For F = 1 To .rows - 1
'''          .Row = F
'''           Codigo1 = Ninguno
'''           Codigo2 = Ninguno
'''           For N = 1 To .cols - 1
'''              .Col = N
'''               Codigo = TrimStrg(Replace(.Text, "'", ""))
'''              'MsgBox Codigo & "...."
'''               Select Case N
'''                 Case 1: Codigo1 = Codigo     'Codigo Antiguo
'''                 Case 3: Codigo2 = Codigo     'Codigo Nuevo
'''               End Select
'''           Next N
'''           If Codigo1 = "" Then Codigo1 = Ninguno
'''           If Codigo2 = "" Then Codigo2 = Ninguno
'''          'Catalogo de Cuentas
'''           Migrar_Cta_Nueva "Catalogo_CxCxP", "Cta", Codigo1, Codigo2
'''           Migrar_Cta_Nueva "Catalogo_Lineas", "CxC_Anterior", Codigo1, Codigo2
'''           Migrar_Cta_Nueva "Catalogo_Lineas", "CxC", Codigo1, Codigo2
'''           Migrar_Cta_Nueva "Catalogo_Productos", "Cta_Inventario", Codigo1, Codigo2
'''           Migrar_Cta_Nueva "Catalogo_Productos", "Cta_Costo_Venta", Codigo1, Codigo2
'''           Migrar_Cta_Nueva "Catalogo_Productos", "Cta_Ventas", Codigo1, Codigo2
'''           Migrar_Cta_Nueva "Catalogo_Productos", "Cta_Ventas_0", Codigo1, Codigo2
'''           Migrar_Cta_Nueva "Catalogo_Productos", "Cta_Ventas_Ant", Codigo1, Codigo2
'''           Migrar_Cta_Nueva "Catalogo_Productos", "Cta_Venta_Anticipada", Codigo1, Codigo2
'''           Migrar_Cta_Nueva "Catalogo_Rol_Pagos", "Cta_Sueldo", Codigo1, Codigo2
'''           Migrar_Cta_Nueva "Catalogo_Rol_Pagos", "Cta_Vacacion", Codigo1, Codigo2
'''           Migrar_Cta_Nueva "Catalogo_Rol_Pagos", "Cta_Horas_Ext", Codigo1, Codigo2
'''           Migrar_Cta_Nueva "Catalogo_Rol_Pagos", "Cta_Antiguedad", Codigo1, Codigo2
'''           Migrar_Cta_Nueva "Catalogo_Rol_Pagos", "Cta_Diferencia", Codigo1, Codigo2
'''           Migrar_Cta_Nueva "Catalogo_Rol_Pagos", "Cta_IESS_Patronal", Codigo1, Codigo2
'''           Migrar_Cta_Nueva "Catalogo_Rol_Pagos", "Cta_IESS_Personal", Codigo1, Codigo2
'''           Migrar_Cta_Nueva "Catalogo_Rol_Pagos", "Cta_Aporte_Patronal_G", Codigo1, Codigo2
'''           Migrar_Cta_Nueva "Catalogo_Rol_Pagos", "Cta_Decimo_Cuarto_G", Codigo1, Codigo2
'''           Migrar_Cta_Nueva "Catalogo_Rol_Pagos", "Cta_Decimo_Tercer_G", Codigo1, Codigo2
'''           Migrar_Cta_Nueva "Catalogo_Rol_Pagos", "Cta_Fondo_Reserva_G", Codigo1, Codigo2
'''           Migrar_Cta_Nueva "Catalogo_Rol_Pagos", "Cta_Decimo_Cuarto_P", Codigo1, Codigo2
'''           Migrar_Cta_Nueva "Catalogo_Rol_Pagos", "Cta_Decimo_Tercer_P", Codigo1, Codigo2
'''           Migrar_Cta_Nueva "Catalogo_Rol_Pagos", "Cta_Fondo_Reserva_P", Codigo1, Codigo2
'''           Migrar_Cta_Nueva "Catalogo_Rol_Pagos", "Cta_Quincena", Codigo1, Codigo2
'''           Migrar_Cta_Nueva "Catalogo_Rol_Pagos", "Cta_Forma_Pago", Codigo1, Codigo2
'''           Migrar_Cta_Nueva "Catalogo_Rol_Rubros", "Cta", Codigo1, Codigo2
'''           Migrar_Cta_Nueva "Detalle_Factura", "Cta_Venta", Codigo1, Codigo2
'''           Migrar_Cta_Nueva "Facturas", "Cta_CxP", Codigo1, Codigo2
'''           Migrar_Cta_Nueva "Facturas", "Cta_Venta", Codigo1, Codigo2
'''
'''          'Transacciones
'''           Migrar_Cta_Nueva "Trans_Abonos", "Cta_CxP", Codigo1, Codigo2
'''           Migrar_Cta_Nueva "Trans_Abonos", "Cta", Codigo1, Codigo2
'''           Migrar_Cta_Nueva "Trans_Air", "Cta_Retencion", Codigo1, Codigo2
'''           Migrar_Cta_Nueva "Trans_Compras", "Cta_Servicio", Codigo1, Codigo2
'''           Migrar_Cta_Nueva "Trans_Compras", "Cta_Bienes", Codigo1, Codigo2
'''           Migrar_Cta_Nueva "Trans_Compras", "Cta_Pago", Codigo1, Codigo2
'''           Migrar_Cta_Nueva "Trans_Compras", "Cta_Gasto", Codigo1, Codigo2
'''           Migrar_Cta_Nueva "Trans_Gastos_Caja", "Cta", Codigo1, Codigo2
'''           Migrar_Cta_Nueva "Trans_Kardex", "Cta_Inv", Codigo1, Codigo2
'''           Migrar_Cta_Nueva "Trans_Kardex", "Contra_Cta", Codigo1, Codigo2
'''''           Migrar_Cta_Nueva "Trans_Prestamos", "Cta", Codigo1, Codigo2
'''''           Migrar_Cta_Nueva "Trans_Prestamos", "Cta_Cartera", Codigo1, Codigo2
'''           Migrar_Cta_Nueva "Trans_Presupuestos", "Cta", Codigo1, Codigo2
'''''           Migrar_Cta_Nueva "Trans_Retenciones", "Cta", Codigo1, Codigo2
'''           Migrar_Cta_Nueva "Trans_Rol_de_Pagos", "Cta", Codigo1, Codigo2
'''           Migrar_Cta_Nueva "Trans_Rol_Pagos", "Cta", Codigo1, Codigo2
'''           Migrar_Cta_Nueva "Trans_Rol_Pagos", "Cta_HExt", Codigo1, Codigo2
'''           Migrar_Cta_Nueva "Trans_Rol_Pagos", "Cta_IESS", Codigo1, Codigo2
'''           Migrar_Cta_Nueva "Trans_SubCtas", "Cta", Codigo1, Codigo2
'''           Migrar_Cta_Nueva "Trans_Ventas", "Cta_Servicio", Codigo1, Codigo2
'''           Migrar_Cta_Nueva "Trans_Ventas", "Cta_Bienes", Codigo1, Codigo2
'''           Migrar_Cta_Nueva "Transacciones", "Cta", Codigo1, Codigo2
'''           If Parpadear Then
'''              Me.Caption = Format$(F / Rango.NumFila2, "00%") & " Migracion en Curso..."
'''           Else
'''              Me.Caption = Format$(F / Rango.NumFila2, "00%")
'''           End If
'''           Parpadear = Not Parpadear
'''      Next F
'''  End With
'''  RatonNormal
'''End Sub
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'''Public Sub Importar_Codigos_Retenciones()
'''Dim F As Long
'''Dim N As Long
'''Dim Parpadear As Boolean
'''  RatonReloj
'''  Parpadear = True
'''  With AdoExcelAdodc.Recordset
'''   If .rows > 0 Then
'''      .Row = 1
'''      .Col = 1
'''       Mifecha = TrimStrg(Replace(.Text, "'", ""))
'''       sSQL = "DELETE * " _
'''            & "FROM Tipo_Concepto_Retencion " _
'''            & "WHERE Fecha_Inicio >= #" & BuscarFecha(Mifecha) & "# "
'''       Ejecutar_SQL_SP sSQL
'''       sSQL = "SELECT * " _
'''            & "FROM Tipo_Concepto_Retencion " _
'''            & "WHERE Fecha_Inicio >= #" & BuscarFecha(Mifecha) & "# "
'''       Select_Adodc AdoAux, sSQL
'''      'Empezamos la importacion de las facturas
'''       For F = 1 To .rows - 1
'''          .Row = F
'''           For N = 1 To .cols - 1
'''              .Col = N
'''               Codigo = TrimStrg(Replace(.Text, "'", ""))
'''               Select Case N
'''                 Case 1: TipoCta = Codigo     'Fecha_Inicio
'''                 Case 2: TipoDoc = Codigo     'Fecha_Final
'''                 Case 3: Codigo2 = Codigo     'Codigo
'''                 Case 4: Cuenta = Codigo      'Concepto Retención en la Fuente de Impuesto a la Renta
'''                 Case 5: Codigo1 = Codigo     'Porcentaje
'''                 Case 6: Codigo3 = Codigo     'Tipo_Pago
'''                 Case 7: Codigo4 = Codigo     'Ingresar_Porcentaje
'''                 Case 8: CodigoB = Codigo     'Sustento
'''               End Select
'''           Next N
'''          'Insertamos el Codigo nuevo
'''           If Codigo2 <> Ninguno Then
'''              AdoAux.Recordset.AddNew
'''              AdoAux.Recordset.Fields("Fecha_Inicio") = TipoCta
'''              AdoAux.Recordset.Fields("Fecha_Final") = TipoDoc
'''              AdoAux.Recordset.Fields("Codigo") = Codigo2
'''              AdoAux.Recordset.Fields("Concepto") = Cuenta
'''              AdoAux.Recordset.Fields("Porcentaje") = Val(Codigo1)
'''              AdoAux.Recordset.Fields("Ingresar_Porcentaje") = Codigo4
'''              AdoAux.Recordset.Fields("T") = CodigoB
'''              AdoAux.Recordset.Fields("Tipo_Pago") = TrimStrg(MidStrg(Codigo3, 1, 1))
'''              AdoAux.Recordset.Update
'''           End If
'''           If Parpadear Then
'''              Me.Caption = Format$(F / Rango.NumFila2, "00%") & " Migracion en Curso..."
'''           Else
'''              Me.Caption = Format$(F / Rango.NumFila2, "00%")
'''           End If
'''           Parpadear = Not Parpadear
'''       Next F
'''   End If
'''  End With
'''  RatonNormal
'''End Sub
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'''Public Sub Importar_Plan_Cuentas_Externas()
'''Dim F As Long
'''Dim N As Long
'''Dim Clave As Long
'''Dim Parpadear As Boolean
'''  RatonReloj
'''  Parpadear = True
'''  Clave = 0
'''  With AdoExcelAdodc.Recordset
'''       If .rows > 0 Then
'''           sSQL = "DELETE * " _
'''                & "FROM Catalogo_Cuentas " _
'''                & "WHERE Item = '" & NumEmpresa & "' " _
'''                & "AND Periodo = '" & Periodo_Contable & "' "
'''           Ejecutar_SQL_SP sSQL
'''       End If
'''      'Empezamos la importacion de las facturas
'''       For F = 1 To .rows - 1
'''          .Row = F
'''           For N = 1 To .cols - 1
'''              .Col = N
'''               Codigo = TrimStrg(Replace(.Text, "'", ""))
'''               Select Case N
'''                 Case 1: Codigo3 = Codigo
'''                         If Codigo3 = "" Then Codigo3 = Ninguno
'''                 Case 2: TipoDoc = Codigo     'DG
'''                 Case 3: Codigo2 = Codigo     'Codigo Nuevo
'''                         If Codigo2 = "" Then Codigo2 = "0"
'''                 Case 4: Cuenta = Codigo      'Cuenta
'''               End Select
'''           Next N
'''          'Insertamos el Codigo nuevo
'''           If Codigo2 <> Ninguno Then
'''              SetAdoAddNew "Catalogo_Cuentas"
'''              SetAdoFields "TC", "N"
'''              SetAdoFields "DG", TipoDoc
'''              SetAdoFields "Codigo", Codigo2
'''              SetAdoFields "Codigo_Ext", Codigo3
'''              SetAdoFields "Cuenta", TrimStrg(MidStrg(Cuenta, 1, 80))
'''              If TipoDoc = "D" Then
'''                 Clave = Clave + 1
'''                 SetAdoFields "Clave", Clave
'''              End If
'''              SetAdoFields "Periodo", Periodo_Contable
'''              SetAdoFields "Item", NumEmpresa
'''              SetAdoUpdate
'''           End If
'''           If Parpadear Then
'''              Me.Caption = Format$(F / Rango.NumFila2, "00%") & " Migracion en Curso..."
'''           Else
'''              Me.Caption = Format$(F / Rango.NumFila2, "00%")
'''           End If
'''           Parpadear = Not Parpadear
'''      Next F
'''  End With
'''  RatonNormal
'''End Sub
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'''Public Sub Importar_SubModulo()
'''Dim I As Long
'''Dim N As Long
'''Dim SecuencialF As Long
'''
'''Dim NumTrans As Long
'''  Progreso_Iniciar
'''  Trans_No = 199
'''  IniciarAsientosDe DGAsiento, AdoAsiento
'''  With AdoExcelAdodc.Recordset
'''      .Row = 1
'''      .Col = 1
'''       Mifecha = TrimStrg(.Text)
'''       FechaFin = Mifecha
'''      .Col = 6
'''       LblConcepto.Caption = TrimStrg(.Text)
'''       Progreso_Barra.Valor_Maximo = .rows
'''       For i = 1 To .rows - 1
'''           Progreso_Barra.Mensaje_Box = "Generando Fecha: " & Mifecha & ", Submodulo: " & NombreCliente
'''           Progreso_Esperar
'''          .Row = i
'''           CodigoCli = Ninguno
'''           Debitos = 0
'''           Creditos = 0
'''          .Col = 2
'''           CI_Representante = TrimStrg(Replace(.Text, "'", ""))
'''           sSQL = "SELECT * " _
'''                & "FROM Clientes " _
'''                & "WHERE CI_RUC = '" & CI_Representante & "' "
'''           Select_Adodc AdoAux, sSQL
'''           If AdoAux.Recordset.RecordCount > 0 Then CodigoCli = AdoAux.Recordset.Fields("Codigo")
'''          .Col = 3
'''           NombreCliente = UCaseStrg(TrimStrg(.Text))
'''          .Col = 4
'''           Factura_No = Val(TrimStrg(.Text))
'''          .Col = 5
'''           Cta = TrimStrg(.Text)
'''           Leer_Cta_Catalogo( Cta
'''          .Col = 6
'''           CodigoA = TrimStrg(.Text)
'''          .Col = 7
'''           Debitos = Val(TrimStrg(.Text))
'''          .Col = 8
'''           Creditos = Val(TrimStrg(.Text))
'''
'''          'Procedemos a ingresar los submodulos
'''           If IsDate(Mifecha) And Len(Codigo) > 1 And Len(CodigoCli) > 1 Then
'''              SetAdoAddNew "Asiento_SC"
'''              SetAdoFields "Codigo", CodigoCli
'''              SetAdoFields "Beneficiario", NombreCliente
'''              SetAdoFields "Factura", Factura_No
'''              SetAdoFields "Detalle_SubCta", CodigoA
'''              SetAdoFields "FECHA_V", Mifecha
'''              SetAdoFields "TC", SubCta
'''              SetAdoFields "Cta", Cta
'''              SetAdoFields "T_No", Trans_No
'''              SetAdoFields "SC_No", i
'''              SetAdoFields "TM", "1"
'''              If Debitos > 0 Then
'''                 OpcDH = "1"
'''                 ValorDH = Debitos
'''              End If
'''              If Creditos > 0 Then
'''                 OpcDH = "2"
'''                 ValorDH = Creditos
'''              End If
'''              SetAdoFields "DH", OpcDH
'''              SetAdoFields "Valor", Redondear(ValorDH, 2)
'''              SetAdoUpdate
'''            End If
'''
'''      Next i
'''  End With
'''  DGAsiento.Visible = False
'''  CodigoCli = Ninguno
'''  SQL2 = "SELECT Cta,DH,SUM(Valor) As TValor " _
'''       & "FROM Asiento_SC " _
'''       & "WHERE Item = '" & NumEmpresa & "' " _
'''       & "AND CodigoU = '" & CodigoUsuario & "' " _
'''       & "AND T_No = " & Trans_No & " " _
'''       & "GROUP BY Cta,DH " _
'''       & "ORDER BY Cta,DH "
'''  Select_Adodc AdoAux, SQL2
'''  With AdoAux.Recordset
'''   If .RecordCount > 0 Then
'''       Do While Not .EOF
'''          Cta = .Fields("Cta")
'''          OpcDH = .Fields("DH")
'''          ValorDH = .Fields("TValor")
'''          If OpcDH = "1" Then
'''             InsertarAsientos AdoAsiento, Cta, 0, ValorDH, 0
'''          Else
'''             InsertarAsientos AdoAsiento, Cta, 0, 0, ValorDH
'''          End If
'''         .MoveNext
'''       Loop
'''   End If
'''  End With
''' 'Verificacion SubTotal
'''  SQL2 = "SELECT * " _
'''       & "FROM Asiento " _
'''       & "WHERE Item = '" & NumEmpresa & "' " _
'''       & "AND CodigoU = '" & CodigoUsuario & "' " _
'''       & "AND T_No = " & Trans_No & " " _
'''       & "ORDER BY CODIGO,DEBE,HABER "
'''  Select_Adodc_Grid DGAsiento, AdoAsiento, SQL2
'''  Debe = 0: Haber = 0
'''  With AdoAsiento.Recordset
'''   If .RecordCount > 0 Then
'''      .MoveFirst
'''       Do While Not .EOF
'''          Debe = Debe + .Fields("DEBE")
'''          Haber = Haber + .Fields("HABER")
'''         .MoveNext
'''       Loop
'''   End If
'''  End With
'''  DGAsiento.Visible = True
'''  LabelDebe.Caption = Format$(Debe, "#,##0.00")
'''  LabelHaber.Caption = Format$(Haber, "#,##0.00")
'''  LblDiferencia.Caption = Format$(Debe - Haber, "#,##0.00")
'''  Progreso_Final
'''End Sub
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'''Public Sub Importar_Compras()
'''Dim I As Long
'''Dim N As Long
'''Dim SerieF1 As String
'''Dim SerieF2 As String
'''Dim SecuencialF As Long
'''Dim SerieR1 As String
'''Dim SerieR2 As String
'''Dim SecuencialR As Long
'''
''''Dim NumTrans As Long
''''Dim NumTransR As Long
'''Dim Tot_Propinas As Currency
'''
'''Dim Cta_Gasto As String
'''
'''Dim PagoLocExt As String
'''Dim PaisEfecPago As String
'''Dim AplicConvDobTrib As String
'''Dim PagExtSujRetNorLeg As String
'''Dim FormaPago As String
'''
'''Dim Cta_Ret_1 As String
'''Dim Cta_Ret_1_75 As String
'''Dim Cta_Ret_2 As String
'''Dim Cta_Ret_2_75 As String
'''Dim Cta_Ret_5 As String
'''Dim Cta_Ret_8 As String
'''Dim Cta_Ret_10 As String
'''Dim Cta_Ret_25 As String
'''Dim Cta_Ret_IVA_10 As String
'''Dim Cta_Ret_IVA_20 As String
'''Dim Cta_Ret_IVA_30 As String
'''Dim Cta_Ret_IVA_50 As String
'''Dim Cta_Ret_IVA_70 As String
'''Dim Cta_Ret_IVAB_100 As String
'''Dim Cta_Ret_IVAS_100 As String
'''
'''Dim Total_Ret_1 As Currency
'''Dim Total_Ret_1_75 As Currency
'''Dim Total_Ret_2 As Currency
'''Dim Total_Ret_2_75 As Currency
'''Dim Total_Ret_5 As Currency
'''Dim Total_Ret_8 As Currency
'''Dim Total_Ret_10 As Currency
'''Dim Total_Ret_25 As Currency
'''Dim Total_Ret_IVA_10 As Currency
'''Dim Total_Ret_IVA_20 As Currency
'''Dim Total_Ret_IVA_30 As Currency
'''Dim Total_Ret_IVA_50 As Currency
'''Dim Total_Ret_IVA_70 As Currency
'''Dim Total_Ret_IVAB_100 As Currency
'''Dim Total_Ret_IVAS_100 As Currency
'''
'''  Cta_Ret_1 = Leer_Seteos_Ctas("Cta_Ret_1")
'''  Cta_Ret_1_75 = Leer_Seteos_Ctas("Cta_Ret_1.75")
'''  Cta_Ret_2 = Leer_Seteos_Ctas("Cta_Ret_2")
'''  Cta_Ret_2_75 = Leer_Seteos_Ctas("Cta_Ret_2.75")
'''  Cta_Ret_5 = Leer_Seteos_Ctas("Cta_Ret_5")
'''  Cta_Ret_8 = Leer_Seteos_Ctas("Cta_Ret_8")
'''  Cta_Ret_10 = Leer_Seteos_Ctas("Cta_Ret_10")
'''  Cta_Ret_25 = Leer_Seteos_Ctas("Cta_Ret_10")
'''
'''  Cta_Ret_IVA_10 = Leer_Seteos_Ctas("Cta_Ret_IVA_30")
'''  Cta_Ret_IVA_20 = Leer_Seteos_Ctas("Cta_Ret_IVA_30")
'''  Cta_Ret_IVA_30 = Leer_Seteos_Ctas("Cta_Ret_IVA_30")
'''  Cta_Ret_IVA_50 = Leer_Seteos_Ctas("Cta_Ret_IVA_30")
'''  Cta_Ret_IVA_70 = Leer_Seteos_Ctas("Cta_Ret_IVA_70")
'''  Cta_Ret_IVAB_100 = Leer_Seteos_Ctas("Cta_Ret_IVAS_100")
'''  Cta_Ret_IVAS_100 = Leer_Seteos_Ctas("Cta_Ret_IVAS_100")
'''
'''  'Encerar_Facturas
''''''  FA.Cod_CxC = DCLinea.Text
''''''  Lineas_De_CxC FA
''''''  SerieFactura = FA.Serie
''''''  Fecha_Vence = FA.Vencimiento
''''''  Autorizacion = FA.Autorizacion
''''''  Cta_Cobrar = FA.Cta_CxP
'''  Eliminar_Asientos_SP True
'''  Bandera = False
'''  Evaluar = True
''''  NumTrans = Maximo_De("Trans_Compras", "ID")
''''  NumTransR = Maximo_De("Trans_Air", "ID")
''' 'MsgBox Tipo_Carga
'''  With AdoExcelAdodc.Recordset
'''       For i = 1 To .rows - 1
'''          .Row = i
'''           CodigoCli = "9999999999"
'''          .Col = 1
'''           TipoDoc = TrimStrg(.Text)
'''          .Col = 2
'''           If Len(TrimStrg(.Text)) = 2 Then TipoCta = TrimStrg(.Text) Else TipoCta = "01"
'''          .Col = 3
'''           Mifecha = TrimStrg(.Text)
'''          .Col = 9
'''           NombreCliente = UCaseStrg(TrimStrg(.Text))
'''          'MsgBox NombreCliente
'''          .Col = 4
'''           Codigo = TrimStrg(Replace(.Text, "'", ""))
'''          'RUC/Cedula/Consumidor Final
'''           If Len(Codigo) > 1 Then
'''              If IsNumeric(Codigo) And Len(Codigo) = 12 Then Codigo = "0" & Codigo
'''              If IsNumeric(Codigo) And Len(Codigo) = 9 Then Codigo = "0" & Codigo
'''              DigVerif = Digito_Verificador( Codigo)
'''              Caracter = MidStrg(Codigo, 10, 1)
'''              CodigoCli = Tipo_RUC_CI.Codigo_RUC_CI
'''           End If
'''           CI_Representante = Codigo
'''          .Col = 5
'''           Autorizacion = TrimStrg(.Text)
'''          .Col = 6
'''           SerieR1 = Format$(Val(TrimStrg(MidStrg(.Text, 1, 3))), "000")
'''           SerieR2 = Format$(Val(TrimStrg(MidStrg(.Text, 5, 3))), "000")
'''           SecuencialR = Val(TrimStrg(MidStrg(.Text, 9, 10)))
'''          .Col = 7
'''           SerieF1 = Format$(Val(TrimStrg(MidStrg(.Text, 1, 3))), "000")
'''           SerieF2 = Format$(Val(TrimStrg(MidStrg(.Text, 5, 3))), "000")
'''           SecuencialF = Val(TrimStrg(MidStrg(.Text, 9, 10)))
'''          .Col = 8
'''           Cta_Gasto = TrimStrg(.Text)
'''          .Col = 10
'''          'Concepto que no se procesa
'''          .Col = 11
'''           Total_Sin_No_IVA = Redondear(Val(TrimStrg(.Text)), 2)
'''          'No Objeto de IVA
'''          .Col = 12
'''           Total_Sin_IVA = Redondear(Val(TrimStrg(.Text)), 2)
'''          .Col = 13
'''           Total_Con_IVA = Redondear(Val(TrimStrg(.Text)), 2)
'''          .Col = 14
'''           Total_IVA = Redondear(Val(TrimStrg(.Text)), 2)
'''          .Col = 15
'''           Total = Redondear(Val(TrimStrg(.Text)), 2)
'''           SubTotal = Total_Con_IVA + Total_Sin_IVA + Total_Sin_No_IVA
'''
'''          .Col = 16   '1%
'''           Total_Ret_1 = Redondear(Val(TrimStrg(.Text)), 2)
'''          .Col = 17   '1.75%
'''           Total_Ret_1_75 = Redondear(Val(TrimStrg(.Text)), 2)
'''          .Col = 18
'''           Total_Ret_2 = Redondear(Val(TrimStrg(.Text)), 2)
'''          .Col = 19
'''           Total_Ret_2_75 = Redondear(Val(TrimStrg(.Text)), 2)
'''          .Col = 20
'''           Total_Ret_5 = Redondear(Val(TrimStrg(.Text)), 2)
'''          .Col = 21
'''           Total_Ret_8 = Redondear(Val(TrimStrg(.Text)), 2)
'''          .Col = 22
'''           Total_Ret_10 = Redondear(Val(TrimStrg(.Text)), 2)
'''          .Col = 23
'''           Total_Ret_25 = Redondear(Val(TrimStrg(.Text)), 2)
'''          .Col = 24
'''           Total_Ret_IVA_10 = Redondear(Val(TrimStrg(.Text)), 2)
'''          .Col = 25
'''           Total_Ret_IVA_20 = Redondear(Val(TrimStrg(.Text)), 2)
'''          .Col = 26
'''           Total_Ret_IVA_30 = Redondear(Val(TrimStrg(.Text)), 2)
'''          .Col = 27
'''           Total_Ret_IVA_50 = Redondear(Val(TrimStrg(.Text)), 2)
'''          .Col = 28
'''           Total_Ret_IVA_70 = Redondear(Val(TrimStrg(.Text)), 2)
'''          .Col = 29
'''           Total_Ret_IVAB_100 = Redondear(Val(TrimStrg(.Text)), 2)
'''          .Col = 30
'''           Total_Ret_IVAS_100 = Redondear(Val(TrimStrg(.Text)), 2)
'''          .Col = 33
'''           CodigoP = Val(TrimStrg(.Text))    'Codigo de Retencion
'''          .Col = 34
'''           AutorizaRet = TrimStrg(.Text)
'''          .Col = 35
'''           FechaTexto = TrimStrg(.Text)       'Caducidad de la Factura
'''          .Col = 36
'''           Cta_CajaG = TrimStrg(.Text)
'''          .Col = 37
'''           If Len(TrimStrg(.Text)) = 2 Then PagoLocExt = TrimStrg(.Text) Else PagoLocExt = "01"
'''          .Col = 38
'''           If Len(TrimStrg(.Text)) = 2 Then PaisEfecPago = TrimStrg(.Text) Else PaisEfecPago = "NA"
'''          .Col = 39
'''           If Len(TrimStrg(.Text)) = 2 Then AplicConvDobTrib = TrimStrg(.Text) Else AplicConvDobTrib = "NA"
'''          .Col = 40
'''           If Len(TrimStrg(.Text)) = 2 Then PagExtSujRetNorLeg = TrimStrg(.Text) Else PagExtSujRetNorLeg = "NA"
'''          .Col = 41
'''           If Len(TrimStrg(.Text)) = 2 Then FormaPago = TrimStrg(.Text) Else FormaPago = "01"
'''          .Col = 42
'''           SubModuloGasto = TrimStrg(.Text)
'''          .Col = 43
'''           SubModuloCxCxP = TrimStrg(.Text)
'''
'''          'MsgBox NombreCliente
'''           If IsDate(Mifecha) Then
'''              If CI_Representante = "9999999999999" Then CI_Representante = Ninguno
'''              If CI_Representante <> Ninguno Then
'''                 sSQL = "SELECT Codigo, Cliente " _
'''                      & "FROM Clientes " _
'''                      & "WHERE Codigo = '" & CodigoCli & "' "
'''                 Select_Adodc AdoAux, sSQL
'''                 If AdoAux.Recordset.RecordCount <= 0 Then
'''                    SetAdoAddNew "Clientes"
'''                    SetAdoFields "T", Normal
'''                    SetAdoFields "Codigo", CodigoCli
'''                    SetAdoFields "TD", Tipo_RUC_CI.Tipo_Beneficiario
'''                    SetAdoFields "CI_RUC", CI_Representante
'''                    SetAdoFields "Cliente", UCaseStrg(NombreCliente)
'''                    SetAdoFields "Fecha", FechaSistema
'''                    SetAdoFields "Direccion", "SD"
'''                    SetAdoFields "DirNumero", "SN"
'''                    SetAdoFields "Ciudad", "QUITO"
'''                    SetAdoFields "Prov", "17"
'''                    SetAdoFields "Pais", "593"
'''                    SetAdoFields "CodigoU", CodigoUsuario
'''                    SetAdoUpdate
'''                 End If
'''              End If
'''              sSQL = "DELETE * " _
'''                   & "FROM Trans_Compras " _
'''                   & "WHERE Item = '" & NumEmpresa & "' " _
'''                   & "AND Periodo = '" & Periodo_Contable & "' " _
'''                   & "AND FechaEmision = #" & BuscarFecha(Mifecha) & "# " _
'''                   & "AND IdProv = '" & CodigoCli & "' " _
'''                   & "AND Establecimiento = '" & SerieF1 & "' " _
'''                   & "AND PuntoEmision = '" & SerieF2 & "' " _
'''                   & "AND Secuencial = " & SecuencialF & " " _
'''                   & "AND TipoComprobante = " & Val(TipoDoc) & " " _
'''                   & "AND Autorizacion = '" & Autorizacion & "' "
'''              Ejecutar_SQL_SP sSQL
'''              sSQL = "DELETE * " _
'''                   & "FROM Trans_Air " _
'''                   & "WHERE Item = '" & NumEmpresa & "' " _
'''                   & "AND Periodo = '" & Periodo_Contable & "' " _
'''                   & "AND IdProv = '" & CodigoCli & "' " _
'''                   & "AND EstabRetencion = '" & SerieR1 & "' " _
'''                   & "AND PtoEmiRetencion = '" & SerieR2 & "' " _
'''                   & "AND SecRetencion = " & SecuencialR & " " _
'''                   & "AND AutRetencion = '" & AutorizaRet & "' " _
'''                   & "AND EstabFactura = '" & SerieF1 & "' " _
'''                   & "AND PuntoEmiFactura = '" & SerieF2 & "' " _
'''                   & "AND Factura_No = " & SecuencialF & " "
'''              Ejecutar_SQL_SP sSQL
'''             'Empezamos a grabar los datos de la retencion
'''             'MsgBox Mifecha & vbCrLf & FechaTexto & vbCrLf & CodigoCli
'''
'''              SetAdoAddNew "Trans_Compras"
'''              SetAdoFields "IdProv", CodigoCli
'''              SetAdoFields "DevIva", "N"
'''              SetAdoFields "CodSustento", TipoCta
'''              SetAdoFields "TipoComprobante", Val(TipoDoc)
'''              SetAdoFields "Establecimiento", SerieF1
'''              SetAdoFields "PuntoEmision", SerieF2
'''              SetAdoFields "Secuencial", SecuencialF
'''              SetAdoFields "Autorizacion", Autorizacion
'''              SetAdoFields "FechaEmision", Mifecha
'''              SetAdoFields "FechaRegistro", Mifecha
'''              SetAdoFields "FechaCaducidad", FechaTexto
'''              If Total_IVA > 0 Then
'''                 SetAdoFields "BaseImpGrav", Total_Con_IVA
'''                 SetAdoFields "MontoIva", Total_IVA
'''                 SetAdoFields "PorcentajeIva", 2
'''              End If
'''              SetAdoFields "BaseImponible", Total_Sin_IVA
'''              SetAdoFields "BaseNoObjIVA", Total_Sin_No_IVA
'''
'''              If Total_Ret_IVA_10 > 0 Then
'''                 SetAdoFields "MontoIvaBienes", Total_IVA
'''                 SetAdoFields "PorRetBienes", 1
'''                 SetAdoFields "ValorRetBienes", Total_Ret_IVA_10
'''                 SetAdoFields "Porc_Bienes", "10"
'''                 SetAdoFields "Cta_Bienes", Cta_Ret_IVA_10
'''              End If
'''              If Total_Ret_IVA_30 > 0 Then
'''                 SetAdoFields "MontoIvaBienes", Total_IVA
'''                 SetAdoFields "PorRetBienes", 1
'''                 SetAdoFields "ValorRetBienes", Total_Ret_IVA_30
'''                 SetAdoFields "Porc_Bienes", "30"
'''                 SetAdoFields "Cta_Bienes", Cta_Ret_IVA_30
'''              End If
'''              If Total_Ret_IVAB_100 > 0 Then
'''                 SetAdoFields "MontoIvaBienes", Total_IVA
'''                 SetAdoFields "PorRetBienes", 1
'''                 SetAdoFields "ValorRetBienes", Total_Ret_IVAB_100
'''                 SetAdoFields "Porc_Bienes", "100"
'''                 SetAdoFields "Cta_Bienes", Cta_Ret_IVAB_100
'''              End If
'''              If Total_Ret_IVA_20 > 0 Then
'''                 SetAdoFields "MontoIvaServicios", Total_IVA
'''                 SetAdoFields "PorRetServicios", 2
'''                 SetAdoFields "ValorRetServicios", Total_Ret_IVA_20
'''                 SetAdoFields "Porc_Servicios", "20"
'''                 SetAdoFields "Cta_Servicio", Cta_Ret_IVA_20
'''              End If
'''              If Total_Ret_IVA_50 > 0 Then
'''                 SetAdoFields "MontoIvaServicios", Total_IVA
'''                 SetAdoFields "PorRetServicios", 2
'''                 SetAdoFields "ValorRetServicios", Total_Ret_IVA_50
'''                 SetAdoFields "Porc_Servicios", "50"
'''                 SetAdoFields "Cta_Servicio", Cta_Ret_IVA_50
'''              End If
'''              If Total_Ret_IVA_70 > 0 Then
'''                 SetAdoFields "MontoIvaServicios", Total_IVA
'''                 SetAdoFields "PorRetServicios", 2
'''                 SetAdoFields "ValorRetServicios", Total_Ret_IVA_70
'''                 SetAdoFields "Porc_Servicios", "70"
'''                 SetAdoFields "Cta_Servicio", Cta_Ret_IVA_70
'''              End If
'''              If Total_Ret_IVAS_100 > 0 Then
'''                 SetAdoFields "MontoIvaServicios", Total_IVA
'''                 SetAdoFields "PorRetServicios", 2
'''                 SetAdoFields "ValorRetServicios", Total_Ret_IVAS_100
'''                 SetAdoFields "Porc_Servicios", "100"
'''                 SetAdoFields "Cta_Servicio", Cta_Ret_IVAS_100
'''              End If
'''              SetAdoFields "Cta_Pago", Cta_CajaG
'''              SetAdoFields "Cta_Gasto", Cta_Gasto
'''              SetAdoFields "PagoLocExt", PagoLocExt
'''              SetAdoFields "PaisEfecPago", PaisEfecPago
'''              SetAdoFields "AplicConvDobTrib", AplicConvDobTrib
'''              SetAdoFields "PagExtSujRetNorLeg", PagExtSujRetNorLeg
'''              SetAdoFields "FormaPago", FormaPago
'''              SetAdoFields "Linea_SRI", 0
'''              SetAdoFields "FechaEmiModificado", "000"
'''              SetAdoFields "EstabModificado", "000"
'''              SetAdoFields "PtoEmiModificado", "000"
'''              SetAdoFields "SecModificado", "000"
'''              SetAdoFields "AutModificado", "000"
'''              SetAdoFields "T", Normal
'''              SetAdoFields "TP", "NN"
'''              SetAdoFields "Numero", -1
'''              SetAdoFields "Fecha", Mifecha
'''              SetAdoUpdate
'''              'MsgBox Total_Sin_IVA & vbCrLf & Total_Con_IVA & vbCrLf & NombreCliente
''''              NumTrans = NumTrans + 1
'''             'RETENCION EN LA FUENTE
'''             'MsgBox CodigoP
'''              Total_Ret = Total_Ret_1 + Total_Ret_1_75 + Total_Ret_2 + Total_Ret_2_75 + Total_Ret_5 + Total_Ret_8 + Total_Ret_10 + Total_Ret_25
'''              SetAdoAddNew "Trans_Air"
'''              SetAdoFields "CodRet", CodigoP
'''              SetAdoFields "BaseImp", SubTotal
'''              SetAdoFields "ValRet", Total_Ret
'''              SetAdoFields "EstabRetencion", SerieR1
'''              SetAdoFields "PtoEmiRetencion", SerieR2
'''              SetAdoFields "SecRetencion", SecuencialR
'''              SetAdoFields "AutRetencion", AutorizaRet
'''              SetAdoFields "Tipo_Trans", "C"
'''              SetAdoFields "IdProv", CodigoCli
'''
'''              If Total_Ret_1 > 0 Then
'''                 SetAdoFields "Cta_Retencion", Cta_Ret_1
'''                 SetAdoFields "Porcentaje", 0.01
'''              End If
'''              If Total_Ret_1_75 > 0 Then
'''                 SetAdoFields "Cta_Retencion", Cta_Ret_1_75
'''                 SetAdoFields "Porcentaje", 0.0175
'''              End If
'''              If Total_Ret_2 > 0 Then
'''                 SetAdoFields "Cta_Retencion", Cta_Ret_2
'''                 SetAdoFields "Porcentaje", 0.02
'''              End If
'''              If Total_Ret_2_75 > 0 Then
'''                 SetAdoFields "Cta_Retencion", Cta_Ret_2_75
'''                 SetAdoFields "Porcentaje", 0.0275
'''              End If
'''              If Total_Ret_5 > 0 Then
'''                 SetAdoFields "Cta_Retencion", Cta_Ret_5
'''                 SetAdoFields "Porcentaje", 0.05
'''              End If
'''              If Total_Ret_8 > 0 Then
'''                 SetAdoFields "Cta_Retencion", Cta_Ret_8
'''                 SetAdoFields "Porcentaje", 0.08
'''              End If
'''              If Total_Ret_10 > 0 Then
'''                 SetAdoFields "Cta_Retencion", Cta_Ret_10
'''                 SetAdoFields "Porcentaje", 0.1
'''              End If
'''              If Total_Ret_25 > 0 Then
'''                 SetAdoFields "Cta_Retencion", Cta_Ret_25
'''                 SetAdoFields "Porcentaje", 0.25
'''              End If
'''              SetAdoFields "EstabFactura", SerieF1
'''              SetAdoFields "PuntoEmiFactura", SerieF2
'''              SetAdoFields "Factura_No", SecuencialF
'''              SetAdoFields "Linea_SRI", 0
'''              SetAdoFields "T", Normal
'''              SetAdoFields "TP", "NN"
'''              SetAdoFields "Numero", -1
'''              SetAdoFields "Fecha", Mifecha
'''              SetAdoUpdate
'''
''''              NumTransR = NumTransR + 1
'''           End If
'''           Me.Caption = "Revisando Datos en el excel: " & i & " de " & .rows - 1 & ", Fecha: " & Mifecha & ", Proveedor: " & NombreCliente
'''      Next i
'''  End With
'''End Sub
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'''Public Sub Importar_Sobrantes_Faltantes()
'''Dim I As Long
'''Dim N As Long
'''Dim CodBodega As String
'''Dim Cta_Sobrantes As String
'''Dim Cta_Faltantes As String
'''Dim Total_Sobrantes As Currency
'''Dim Total_Faltantes As Currency
'''
'''  Cta_Sobrantes = Leer_Seteos_Ctas("Cta_Sobrantes")
'''  Cta_Faltantes = Leer_Seteos_Ctas("Cta_Faltantes")
'''  Trans_No = 100
'''
'''  Eliminar_Asientos_SP True
'''
'''  SQL2 = "SELECT * " _
'''       & "FROM Asiento " _
'''       & "WHERE Item = '" & NumEmpresa & "' " _
'''       & "AND CodigoU = '" & CodigoUsuario & "' " _
'''       & "AND T_No = " & Trans_No & " "
'''  Select_Adodc_Grid DGAsiento, AdoAsiento, SQL2
'''
'''  SQL2 = "SELECT * " _
'''       & "FROM Asiento_K " _
'''       & "WHERE Item = '" & NumEmpresa & "' " _
'''       & "AND CodigoU = '" & CodigoUsuario & "' " _
'''       & "AND T_No = " & Trans_No & " "
'''  Select_Adodc AdoAux, SQL2
'''
'''  With AdoExcelAdodc.Recordset
'''       For I = 1 To .rows - 1
'''          .Row = I
'''           CodigoCli = "9999999999"
'''          .Col = 1
'''           TipoDoc = TrimStrg(.Text)
'''          .Col = 2
'''           CodigoInv = TrimStrg(.Text)
'''          .Col = 3
'''           Producto = TrimStrg(.Text)
'''          .Col = 8
'''           Precio = Val(TrimStrg(.Text))
'''          .Col = 11
'''           Cantidad = Val(TrimStrg(.Text))
'''          .Col = 12
'''           CodBodega = TrimStrg(.Text)
'''          'Empezamos a grabar los datos de la retencion
'''           Cta_Inventario = Ninguno
'''           SQL2 = "SELECT * " _
'''                & "FROM Catalogo_Productos " _
'''                & "WHERE Item = '" & NumEmpresa & "' " _
'''                & "AND Periodo = '" & Periodo_Contable & "' " _
'''                & "AND Codigo_Inv = '" & CodigoInv & "' "
'''           Select_Adodc AdoAct, SQL2
'''           If AdoAct.Recordset.RecordCount > 0 Then
'''              Cta_Inventario = AdoAct.Recordset.Fields("Cta_Inventario")
'''           End If
'''           SetAdoAddNew "Asiento_K"
'''           SetAdoFields "TC", TipoDoc
'''           SetAdoFields "CODIGO_INV", CodigoInv
'''           SetAdoFields "PRODUCTO", Producto
'''           SetAdoFields "VALOR_UNIT", Precio
'''           SetAdoFields "VALOR_TOTAL", Redondear(Precio * Abs(Cantidad), 2)
'''           SetAdoFields "CANTIDAD", Abs(Cantidad)
'''           SetAdoFields "CTA_INVENTARIO", Cta_Inventario
'''           SetAdoFields "CANT_ES", Abs(Cantidad)
'''           SetAdoFields "CodBod", CodBodega
'''           SetAdoFields "A_No", I
'''           SetAdoFields "T_No", Trans_No
'''           If Cantidad > 0 Then
'''              Cta = Cta_Faltantes
'''              SetAdoFields "DH", "1"
'''           Else
'''              Cta = Cta_Sobrantes
'''              SetAdoFields "DH", "2"
'''           End If
'''           SetAdoFields "CONTRA_CTA", Cta
'''           SetAdoUpdate
'''           Me.Caption = "Revisando Datos en el excel: " & I & " de " & .rows - 1 & ", Fecha: " & Mifecha & ", Proveedor: " & NombreCliente
'''      Next I
'''  End With
'''  SQL2 = "SELECT DH,CTA_INVENTARIO,CONTRA_CTA,SUM(VALOR_TOTAL) As TOTAL_INV " _
'''       & "FROM Asiento_K " _
'''       & "WHERE Item = '" & NumEmpresa & "' " _
'''       & "AND CodigoU = '" & CodigoUsuario & "' " _
'''       & "AND T_No = " & Trans_No & " " _
'''       & "GROUP BY DH,CTA_INVENTARIO,CONTRA_CTA " _
'''       & "ORDER BY DH,CTA_INVENTARIO,CONTRA_CTA "
'''  Select_Adodc AdoAux, SQL2
'''  CodigoCli = Ninguno
'''  FechaFin = MBFechaI
'''  Fecha_Vence = MBFechaI
'''  CodigoCli = Ninguno
'''  NoCheque = Ninguno
'''  DetalleComp = Ninguno
'''  LblConcepto.Caption = "(" & NumEmpresa & ") Ingreso de Sobrantes o Faltante de Inventario del " & MBFechaI
'''  With AdoAux.Recordset
'''   If .RecordCount > 0 Then
'''       Do While Not .EOF
'''          Valor = Redondear(.Fields("TOTAL_INV"), 2)
'''          If .Fields("DH") = "1" Then
'''             InsertarAsientos AdoAsiento, .Fields("CTA_INVENTARIO"), 0, Valor, 0
'''             InsertarAsientos AdoAsiento, .Fields("CONTRA_CTA"), 0, 0, Valor
'''          Else
'''             InsertarAsientos AdoAsiento, .Fields("CONTRA_CTA"), 0, Valor, 0
'''             InsertarAsientos AdoAsiento, .Fields("CTA_INVENTARIO"), 0, 0, Valor
'''          End If
'''         .MoveNext
'''       Loop
'''   End If
'''  End With
'''  SQL2 = "SELECT * " _
'''       & "FROM Asiento " _
'''       & "WHERE Item = '" & NumEmpresa & "' " _
'''       & "AND CodigoU = '" & CodigoUsuario & "' " _
'''       & "AND T_No = " & Trans_No & " "
'''  Select_Adodc_Grid DGAsiento, AdoAsiento, SQL2
'''  Debe = 0
'''End Sub
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'''Public Sub Importar_Parroquias()
'''Dim I As Long
'''Dim N As Long
'''Dim p As Long
'''Dim D As Long
'''Dim Dato() As String
'''
'''  FechaTexto = FechaSistema
'''  With AdoExcelAdodc.Recordset
'''       ReDim Dato(.cols) As String
'''      'MsgBox .Rows & vbCrLf & .Cols
'''       For i = 1 To .rows - 1
'''          .Row = i
'''           For N = 1 To .cols - 1
'''              .Col = N
'''               Codigo = Replace(.Text, "'", "")
'''               Codigo = Replace(Codigo, "-", "")
'''               Codigo = Replace(Codigo, "´", "")
'''               Codigo = Replace(Codigo, "`", "")
'''               Codigo = TrimStrg(Codigo)
'''               If Codigo = "" Then Codigo = Ninguno
'''               Dato(.Col) = Codigo
'''           Next N
'''           sSQL = "SELECT * " _
'''                & "FROM Trans_Parroquias " _
'''                & "WHERE Item = '" & NumEmpresa & "' " _
'''                & "AND Periodo = '" & Periodo_Contable & "' " _
'''                & "AND Beneficiario = '" & Dato(1) & "' "
'''           Select_Adodc AdoAux, sSQL
'''           If AdoAux.Recordset.RecordCount <= 0 Then
'''              SetAdoAddNew "Trans_Parroquias"
'''              SetAdoFields "T", Normal
'''              SetAdoFields "Fecha", FechaSistema
'''              SetAdoFields "Cedula", Ninguno
'''              SetAdoFields "Cedula_P", Ninguno
'''              SetAdoFields "Cedula_M", Ninguno
'''              SetAdoFields "Beneficiario", Dato(1)
'''              Select Case Tipo_Carga
'''                Case 50: SetAdoFields "Tipo_Certificado", "BAUTIZO"
'''                         p = InStr(Dato(2), " Y ")
'''                         CodigoA = TrimStrg(MidStrg(Dato(2), 1, p))
'''                         CodigoB = TrimStrg(MidStrg(Dato(2), p + 2, Len(Dato(2))))
'''                         If CodigoA = "" Then CodigoA = Ninguno
'''                         If CodigoB = "" Then Codigo = Ninguno
'''                         Sacar_Ciudad_Fecha Dato(3), CodigoC, Mifecha
'''                         Sacar_Ciudad_Fecha Dato(4), CodigoP, Fecha_Vence
'''                         SetAdoFields "Padrinos", Dato(5)
'''                         SetAdoFields "Ministro", Dato(6)
'''                         SetAdoFields "Nota_Marginal", Dato(7)
'''                         SetAdoFields "Tomo", Val(Dato(8))
'''                         SetAdoFields "Pagina", Val(Dato(9))
'''                         SetAdoFields "Numero", Val(Dato(10))
'''                Case 51: SetAdoFields "Tipo_Certificado", "CONFIRMACION"
'''                         CodigoA = Dato(2)
'''                         CodigoB = Dato(3)
'''                         CodigoC = NombreCiudad
'''                         Mifecha = FechaSistema
'''                         Sacar_Ciudad_Fecha Dato(6), CodigoP, Fecha_Vence
'''                         SetAdoFields "Padrinos", Dato(4)
'''                         SetAdoFields "Ministro", Dato(5)
'''                         SetAdoFields "Tomo", Val(Dato(7))
'''                         SetAdoFields "Pagina", Val(Dato(8))
'''                         SetAdoFields "Numero", Val(Dato(9))
'''                Case 52: SetAdoFields "Tipo_Certificado", "MATRIMONIO"
'''                         CodigoA = Ninguno
'''                         CodigoB = Dato(2)
'''                         CodigoC = NombreCiudad
'''                         Mifecha = FechaSistema
'''                         Sacar_Ciudad_Fecha Dato(3), CodigoP, Fecha_Vence
'''                         SetAdoFields "Padrinos", Dato(4)
'''                         SetAdoFields "Ministro", Dato(5)
'''                         SetAdoFields "Tomo", Val(Dato(6))
'''                         SetAdoFields "Pagina", Val(Dato(7))
'''                         SetAdoFields "Numero", Val(Dato(8))
'''              End Select
'''             'MsgBox CodigoA & vbCrLf & CodigoB & vbCrLf & CodigoC & vbCrLf & Mifecha & vbCrLf & CodigoP & vbCrLf & Fecha_Vence
'''              SetAdoFields "Padre", CodigoA
'''              SetAdoFields "Madre", CodigoB
'''              SetAdoFields "Ciudad_Nacimiento", CodigoC
'''              SetAdoFields "Fecha_Nacimiento", Mifecha
'''              SetAdoFields "Ciudad_B_C_M", CodigoP
'''              SetAdoFields "Fecha_B_C_M", Fecha_Vence
'''              SetAdoUpdate
''           End If
'''           Me.Caption = "Importar de FlexGrid a Sistema de Parroquias: " & i & " de " & Rango.NumFila2
'''      Next i
'''  End With
'''  Me.Caption = "IMPORTACION DE DATOS DE PARROQUIA"
'''  MsgBox "Proceso Terminado con exito," & vbCrLf & "Revise los datos procesados"
'''End Sub
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Sub Importar_Abonos_Transferencias()
Dim AdoCatalogoDB As ADODB.Recordset
Dim AdoSubCtaDB As ADODB.Recordset
    
    Progreso_Barra.Mensaje_Box = "Subiendo Contabilidad Externa con SubModulos"
    Progreso_Iniciar
    RatonReloj
    DGExcelAdodc.Visible = False
    sSQL = "DELETE * " _
         & "FROM Tabla_Temporal " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Modulo = '" & NumModulo & "' " _
         & "AND CodigoU = '" & CodigoUsuario & "' "
    Ejecutar_SQL_SP sSQL
    
    TextoImprimio = ""
    Importar_Abonos_Facturas_SP
    
    ConectarAdodc AdoExcelAdodc
    Select_Adodc AdoExcelAdodc, "SELECT * FROM Asiento_CSV_" & CodigoUsuario
    
    DGExcelAdodc.Visible = True
    RatonNormal
    Progreso_Final
    If Len(TextoImprimio) > 2 Then FInfoError.Show
  
    FA.Factura = 0
    FA.Fecha_Corte = FechaSistema
    Actualizar_Abonos_Facturas_SP FA
    Me.Caption = "IMPORTACION DE ABONOS AUTOMATICOS"
End Sub
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Sub Importar_Empleados()
Dim I As Long
Dim N As Long
Dim Lista_Clientes_Nuevos As String
Dim Cta_Transf As String
Dim Crear_Nuevo As Boolean
Dim Aplica_FP As Boolean
  Lista_Clientes_Nuevos = ""
  sSQL = "SELECT * " _
       & "FROM Clientes " _
       & "WHERE Codigo <> '.' " _
       & "ORDER BY CI_RUC "
  Select_Adodc AdoClientes, sSQL

  sSQL = "SELECT * " _
       & "FROM Catalogo_Rol_Pagos " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "ORDER BY Codigo "
  Select_Adodc AdoAux, sSQL

  With AdoExcelAdodc.Recordset
   If .RecordCount > 0 Then
       Progreso_Barra.Valor_Maximo = .RecordCount
      .MoveFirst
       Do While Not .EOF
           TBeneficiario.T = Normal
           TBeneficiario.FA = False
           TBeneficiario.TP = "E"
           TBeneficiario.Codigo = Ninguno
           TBeneficiario.CI_RUC = "9999999999999"
           TBeneficiario.Fecha = FechaSistema
           TBeneficiario.Fecha_A = FechaSistema
           TBeneficiario.Fecha_N = FechaSistema
           TBeneficiario.Cliente = Ninguno
           TBeneficiario.Sexo = Ninguno
           TBeneficiario.Email1 = Ninguno
           TBeneficiario.Email2 = Ninguno
           TBeneficiario.Direccion = "SD"
           TBeneficiario.DirNumero = "SN"
           TBeneficiario.Telefono1 = "022000000"
           TBeneficiario.Celular = "0990000000"
           TBeneficiario.Ciudad = NombreCiudad
           TBeneficiario.Prov = CodigoProv
           TBeneficiario.Pais = CodigoPais
           TBeneficiario.Grupo_No = "NUEVOS"
           TBeneficiario.Profesion = Ninguno
           Crear_Nuevo = False
           
           Codigo = Dato_Campo(.fields(9))
          'RUC/Cedula/Codigo Alumno/Consumidor Final
           If Len(Codigo) > 1 Then
              If IsNumeric(Codigo) Then
                 If Len(Codigo) < 9 Then
                    Codigo = Format$(Val(Codigo), "00000000")
                    TBeneficiario.FA = True
                 ElseIf Len(Codigo) = 9 Then
                    Codigo = "0" & Codigo
                 ElseIf Len(Codigo) = 11 Then
                    Codigo = "00" & Codigo
                 ElseIf Len(Codigo) = 12 Then
                    Codigo = "0" & Codigo
                 End If
              End If
              TBeneficiario.CI_RUC = Codigo
              DigVerif = Digito_Verificador(Codigo)
              Caracter = MidStrg(Codigo, 10, 1)
              TBeneficiario.TP = Tipo_RUC_CI.Tipo_Beneficiario
              TBeneficiario.Codigo = Tipo_RUC_CI.Codigo_RUC_CI
           End If
           TBeneficiario.Fecha_A = Dato_Campo(.fields(6))
           If TBeneficiario.Codigo <> Ninguno And IsDate(TBeneficiario.Fecha_A) Then
              For IdField = 0 To .fields.Count - 1
                  If IdField = 100 Then Codigo = Dato_Campo(.fields(IdField), True) Else Codigo = Dato_Campo(.fields(IdField))
                  Codigo = UCaseStrg(Codigo)
                  Codigo = Sin_Signos_Especiales(Codigo)
                  If Codigo = "" Then Codigo = Ninguno
                 'MsgBox .fields(IdField) & vbCrLf & IdField
                  Select Case IdField + 1
                    Case 1: TBeneficiario.Cliente = UCaseStrg(TrimStrg(MidStrg(Codigo, 1, 60))) 'Cliente
                    Case 2: TBeneficiario.Grupo_No = UCaseStrg(TrimStrg(MidStrg(Codigo, 1, 9)))
                    Case 3: If Codigo = "AFR" Then Aplica_FP = False Else Aplica_FP = True
                   'Case 4: No se utiliza
                    Case 5: TBeneficiario.Salario = Val(Codigo)
                    Case 6: If Len(Codigo) > 3 Then TBeneficiario.Ciudad = UCaseStrg(Codigo) 'Ciudad
                   'Case 7: ya esta asignado la Fecha de Ingreso
                    Case 8: TBeneficiario.Profesion = UCaseStrg(Codigo)
                   'case 9: ya esta asignado la CI
                    Case 10: TBeneficiario.Direccion = UCaseStrg(MidStrg(Codigo, 1, 50)) 'Direccion
                    Case 11: TBeneficiario.Telefono1 = MidStrg(Replace(Codigo, " ", ""), 1, 10) 'Telefono
                    Case 12: TBeneficiario.Fecha_N = Codigo 'Fecha_N
                    Case 13: TBeneficiario.Sexo = MidStrg(Codigo, 1, 1) 'Sexo"
                    Case 14: TBeneficiario.Cte_Ahr_Otro = Codigo
                    Case 15: TBeneficiario.Cod_Banco = Val(Codigo)
                    Case 16: TBeneficiario.Cta_Transf = Codigo
                    Case 17: Cta_Aux = Codigo
                    Case 18: Cta_Gastos = Codigo
                  End Select
              Next IdField
              Cta_Transf = Ninguno
              If Len(TrimStrg(TBeneficiario.Cte_Ahr_Otro) & TrimStrg(TBeneficiario.Cta_Transf)) > 2 Then
                 Cta_Transf = TBeneficiario.Cte_Ahr_Otro & " " & TBeneficiario.Cta_Transf
              End If
              If AdoClientes.Recordset.RecordCount > 0 Then
                 AdoClientes.Recordset.MoveFirst
                 AdoClientes.Recordset.Find ("CI_RUC = '" & TBeneficiario.CI_RUC & "' ")
                 If Not AdoClientes.Recordset.EOF Then
                    TBeneficiario.Codigo = AdoClientes.Recordset.fields("Codigo")
                    AdoClientes.Recordset.fields("T") = TBeneficiario.T
                    If IsDate(TBeneficiario.Fecha_N) Then AdoClientes.Recordset.fields("Fecha_N") = TBeneficiario.Fecha_N
                    If IsDate(TBeneficiario.Fecha) Then AdoClientes.Recordset.fields("Fecha") = TBeneficiario.Fecha
                    If Len(TBeneficiario.Cliente) > 1 Then AdoClientes.Recordset.fields("Cliente") = TBeneficiario.Cliente
                    If Len(TBeneficiario.Sexo) > 1 Then AdoClientes.Recordset.fields("Sexo") = TBeneficiario.Sexo
                    If Len(TBeneficiario.Email1) > 1 Then AdoClientes.Recordset.fields("Email") = TBeneficiario.Email1
                    If Len(TBeneficiario.Email2) > 1 Then AdoClientes.Recordset.fields("Email2") = TBeneficiario.Email2
                    If Len(TBeneficiario.Direccion) > 1 Then AdoClientes.Recordset.fields("Direccion") = TBeneficiario.Direccion
                    If Len(TBeneficiario.DirNumero) > 1 Then AdoClientes.Recordset.fields("DirNumero") = TBeneficiario.DirNumero
                    If Len(TBeneficiario.Telefono1) > 1 Then AdoClientes.Recordset.fields("Telefono") = TBeneficiario.Telefono1
                    If Len(TBeneficiario.Celular) > 1 Then AdoClientes.Recordset.fields("Celular") = TBeneficiario.Celular
                    If Len(TBeneficiario.Ciudad) > 1 Then AdoClientes.Recordset.fields("Ciudad") = TBeneficiario.Ciudad
                    If Len(TBeneficiario.Prov) > 1 Then AdoClientes.Recordset.fields("Prov") = TBeneficiario.Prov
                    If Len(TBeneficiario.Pais) > 1 Then AdoClientes.Recordset.fields("Pais") = TBeneficiario.Pais
                    If Len(TBeneficiario.Grupo_No) > 1 Then AdoClientes.Recordset.fields("Grupo") = TBeneficiario.Grupo_No
                    'MsgBox Len(TBeneficiario.Profesion)
                    If Len(TBeneficiario.Profesion) > 1 Then AdoClientes.Recordset.fields("Profesion") = TBeneficiario.Profesion
                    AdoClientes.Recordset.Update
                 Else
                    Crear_Nuevo = True
                 End If
              Else
                 Crear_Nuevo = True
              End If
              If Crear_Nuevo Then
                 SetAdoAddNew "Clientes"
                 SetAdoFields "T", TBeneficiario.T
                 SetAdoFields "TD", TBeneficiario.TP
                 If Len(TBeneficiario.Codigo) > 1 Then SetAdoFields "Codigo", TBeneficiario.Codigo
                 If Len(TBeneficiario.CI_RUC) > 1 Then SetAdoFields "CI_RUC", TBeneficiario.CI_RUC
                 If IsDate(TBeneficiario.Fecha) Then SetAdoFields "Fecha", TBeneficiario.Fecha
                 If IsDate(TBeneficiario.Fecha_N) Then SetAdoFields "Fecha_N", TBeneficiario.Fecha_N
                 If Len(TBeneficiario.Cliente) > 1 Then SetAdoFields "Cliente", TBeneficiario.Cliente
                 If Len(TBeneficiario.Sexo) > 1 Then SetAdoFields "Sexo", TBeneficiario.Sexo
                 If Len(TBeneficiario.Email1) > 1 Then SetAdoFields "Email", TBeneficiario.Email1
                 If Len(TBeneficiario.Email2) > 1 Then SetAdoFields "Email2", TBeneficiario.Email2
                 If Len(TBeneficiario.Direccion) > 1 Then SetAdoFields "Direccion", TBeneficiario.Direccion
                 If Len(TBeneficiario.DirNumero) > 1 Then SetAdoFields "DirNumero", TBeneficiario.DirNumero
                 If Len(TBeneficiario.Telefono1) > 1 Then SetAdoFields "Telefono", TBeneficiario.Telefono1
                 If Len(TBeneficiario.Celular) > 1 Then SetAdoFields "Celular", TBeneficiario.Celular
                 If Len(TBeneficiario.Ciudad) > 1 Then SetAdoFields "Ciudad", TBeneficiario.Ciudad
                 If Len(TBeneficiario.Prov) > 1 Then SetAdoFields "Prov", TBeneficiario.Prov
                 If Len(TBeneficiario.Grupo_No) > 1 Then SetAdoFields "Grupo", TBeneficiario.Grupo_No
                 If Len(TBeneficiario.Profesion) > 1 Then SetAdoFields "Profesion", TBeneficiario.Profesion
                 SetAdoUpdate
                 Lista_Clientes_Nuevos = Lista_Clientes_Nuevos _
                                       & TBeneficiario.CI_RUC & vbTab _
                                       & TBeneficiario.Cliente & vbTab _
                                       & TBeneficiario.Grupo_No & vbCrLf
              End If
             'Creamos los Clientes Rol de Pagos
              Crear_Nuevo = False
              If AdoAux.Recordset.RecordCount > 0 Then
                 AdoAux.Recordset.MoveFirst
                 AdoAux.Recordset.Find ("Codigo = '" & TBeneficiario.Codigo & "' ")
                 If Not AdoAux.Recordset.EOF Then
                    If Len(TBeneficiario.Fecha_A) Then AdoAux.Recordset.fields("Fecha") = TBeneficiario.Fecha_A
                    If Len(TBeneficiario.Grupo_No) > 1 Then AdoAux.Recordset.fields("Grupo_Rol") = TBeneficiario.Grupo_No
                    If Len(TBeneficiario.Salario) > 1 Then AdoAux.Recordset.fields("Salario") = TBeneficiario.Salario
                    AdoAux.Recordset.fields("T") = TBeneficiario.T
                    AdoAux.Recordset.fields("SN") = "1"
                    AdoAux.Recordset.fields("Valor_Hora") = Redondear(TBeneficiario.Salario / 240, 2)
                    AdoAux.Recordset.fields("Horas_Sem") = 60
                    AdoAux.Recordset.fields("Porc_IESS_Per") = 0.0945
                    AdoAux.Recordset.fields("Porc_IESS_Pat") = 0.1215
                    AdoAux.Recordset.fields("Cta_Transferencia") = Cta_Transf
                    AdoAux.Recordset.fields("Codigo_Banco") = TBeneficiario.Cod_Banco
                    AdoAux.Recordset.fields("Pagar_Fondo_Reserva") = Aplica_FP
                    AdoAux.Recordset.fields("Cta_Forma_Pago") = Cta_Aux
                    If Len(TBeneficiario.Cte_Ahr_Otro) >= 3 Then
                       AdoAux.Recordset.fields("FP") = "T"
                       AdoAux.Recordset.fields("TC") = "BA"
                    Else
                       AdoAux.Recordset.fields("FP") = "E"
                       AdoAux.Recordset.fields("TC") = "CJ"
                    End If
                    AdoAux.Recordset.Update
                 Else
                    Crear_Nuevo = True
                 End If
              Else
                 Crear_Nuevo = True
              End If
              If Crear_Nuevo Then
                 SetAdoAddNew "Catalogo_Rol_Pagos"
                 If Len(TBeneficiario.Codigo) > 1 Then SetAdoFields "Codigo", TBeneficiario.Codigo
                 If Len(TBeneficiario.Fecha_A) Then SetAdoFields "Fecha", TBeneficiario.Fecha_A
                 If Len(TBeneficiario.Grupo_No) > 1 Then SetAdoFields "Grupo_Rol", TBeneficiario.Grupo_No
                 If Len(TBeneficiario.Salario) > 1 Then SetAdoFields "Salario", TBeneficiario.Salario
                 SetAdoFields "T", TBeneficiario.T
                 SetAdoFields "SN", "1"
                 SetAdoFields "Valor_Hora", Redondear(TBeneficiario.Salario / 240, 2)
                 SetAdoFields "Horas_Sem", 60
                 SetAdoFields "IESS_Per", 0.0945
                 SetAdoFields "IESS_Pat", 0.1215
                 SetAdoFields "Cta_Transferencia", Cta_Transf
                 SetAdoFields "Codigo_Banco", TBeneficiario.Cod_Banco
                 SetAdoFields "Pagar_Fondo_Reserva", Aplica_FP
                 SetAdoFields "Cta_Forma_Pago", Cta_Aux
                 If Len(TBeneficiario.Cte_Ahr_Otro) >= 3 Then
                    SetAdoFields "FP", "T"
                    SetAdoFields "TC", "BA"
                 Else
                    SetAdoFields "FP", "E"
                    SetAdoFields "TC", "CJ"
                 End If
                 SetAdoUpdate
              End If
           End If
           Me.Caption = "Importar de FlexGrid a Sistema El Beneficiario: " & TBeneficiario.CI_RUC & ": " & I & " de " & Rango.NumFila2
        .MoveNext
      Loop
   End If
  End With
  If Len(Lista_Clientes_Nuevos) > 2 Then
     TextoImprimio = Lista_Clientes_Nuevos
     Unload FImporta
     FInfoError.Show
  End If
End Sub
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'''Public Sub Importar_Catalogo_RolPagos()
'''Dim I As Long
'''Dim N As Long
'''Dim Lista_Clientes_Nuevos As String
'''Dim Cta_Transf As String
'''Dim Crear_Nuevo As Boolean
'''Dim Aplica_FP As Boolean
'''  Lista_Clientes_Nuevos = ""
'''  sSQL = "SELECT * " _
'''       & "FROM Clientes " _
'''       & "WHERE Codigo <> '.' " _
'''       & "ORDER BY CI_RUC "
'''  Select_Adodc AdoClientes, sSQL
'''
'''  sSQL = "SELECT * " _
'''       & "FROM Catalogo_Rol_Pagos " _
'''       & "WHERE Item = '" & NumEmpresa & "' " _
'''       & "AND Periodo = '" & Periodo_Contable & "' " _
'''       & "ORDER BY Codigo "
'''  Select_Adodc AdoAux, sSQL
'''
'''  With AdoExcelAdodc.Recordset
'''      'MsgBox .Rows & vbCrLf & .Cols
'''       For i = 1 To .rows - 1
'''           TBeneficiario.T = Normal
'''           TBeneficiario.FA = False
'''           TBeneficiario.TP = "E"
'''           TBeneficiario.Codigo = Ninguno
'''           TBeneficiario.CI_RUC = "9999999999999"
'''           TBeneficiario.Fecha = FechaSistema
'''           TBeneficiario.Fecha_A = FechaSistema
'''           TBeneficiario.Fecha_N = FechaSistema
'''           TBeneficiario.Cliente = Ninguno
'''           TBeneficiario.Sexo = Ninguno
'''           TBeneficiario.Email1 = Ninguno
'''           TBeneficiario.Email2 = Ninguno
'''           TBeneficiario.Direccion = "SD"
'''           TBeneficiario.DirNumero = "SN"
'''           TBeneficiario.Telefono1 = "022000000"
'''           TBeneficiario.Celular = "0990000000"
'''           TBeneficiario.Ciudad = NombreCiudad
'''           TBeneficiario.Prov = CodigoProv
'''           TBeneficiario.Pais = CodigoPais
'''           TBeneficiario.Grupo_No = "NUEVOS"
'''           TBeneficiario.Profesion = Ninguno
'''           TBeneficiario.Cte_Ahr_Otro = Ninguno
'''           TBeneficiario.Cta_Transf = Ninguno
'''
'''           Crear_Nuevo = False
'''          .Row = i
'''          .Col = 2
'''           Codigo = Replace(.Text, "'", "")
'''           Codigo = TrimStrg(Replace(Codigo, "-", ""))
'''           If Codigo = "" Then Codigo = Ninguno
'''          'RUC/Cedula/Codigo Alumno/Consumidor Final
'''           If Len(Codigo) > 1 Then
'''              If IsNumeric(Codigo) Then
'''                 If Len(Codigo) < 9 Then
'''                    Codigo = Format$(Val(Codigo), "00000000")
'''                    TBeneficiario.FA = True
'''                 ElseIf Len(Codigo) = 9 Then
'''                    Codigo = "0" & Codigo
'''                 ElseIf Len(Codigo) = 11 Then
'''                    Codigo = "00" & Codigo
'''                 ElseIf Len(Codigo) = 12 Then
'''                    Codigo = "0" & Codigo
'''                 End If
'''              End If
'''              TBeneficiario.CI_RUC = Codigo
'''              DigVerif = Digito_Verificador( Codigo)
'''              Caracter = MidStrg(Codigo, 10, 1)
'''              TBeneficiario.TP = Tipo_RUC_CI.Tipo_Beneficiario
'''              TBeneficiario.Codigo = Tipo_RUC_CI.Codigo_RUC_CI
'''           End If
'''          .Col = 6
'''           Codigo = Replace(.Text, "'", "")
'''           Codigo = TrimStrg(Replace(Codigo, "-", ""))
'''           If Codigo = "" Then Codigo = Ninguno
'''           TBeneficiario.Fecha_A = Codigo
'''           If TBeneficiario.Codigo <> Ninguno And IsDate(TBeneficiario.Fecha_A) Then
'''              TBeneficiario.Fecha_N = TBeneficiario.Fecha_A
'''              TBeneficiario.Fecha = FechaSistema
'''              For N = 1 To .cols - 1
'''                 .Col = N
'''                  Codigo = Replace(.Text, "'", "")
'''                  Codigo = TrimStrg(Replace(Codigo, "-", ""))
'''                  If Codigo = "" Then Codigo = Ninguno
'''                  Select Case N
'''                   'case 2: ya esta asignado la CI
'''                    Case 3: TBeneficiario.Cliente = UCaseStrg(TrimStrg(MidStrg(Codigo, 1, 60))) 'Cliente
'''                    Case 4: TBeneficiario.Grupo_No = UCaseStrg(TrimStrg(MidStrg(Codigo, 1, 25)))
'''                    Case 5: SubCta = UCaseStrg(TrimStrg(MidStrg(Codigo, 1, 10)))
'''                    Case 6: TBeneficiario.Fecha_N = Codigo
'''                    Case 7: If Codigo = "SI" Then Aplica_FP = False Else Aplica_FP = True
'''                    Case 8: TBeneficiario.Salario = Val(Codigo)
'''                    Case 9: TBeneficiario.Profesion = UCaseStrg(Codigo)
'''                    Case 26: If Len(Codigo) > 3 Then TBeneficiario.Ciudad = UCaseStrg(Codigo) 'Ciudad
'''
'''                   'Case 10: TBeneficiario.Direccion = UCaseStrg(MidStrg(Codigo, 1, 50)) 'Direccion
'''                   'Case 11: TBeneficiario.Telefono1 = MidStrg(Codigo, 1, 10) 'Telefono
'''                   'Case 13: TBeneficiario.Sexo = MidStrg(Codigo, 1, 1) 'Sexo"
'''''                    Case 14: TBeneficiario.Cte_Ahr_Otro = Codigo
'''''                    Case 15: TBeneficiario.Cod_Banco = Val(Codigo)
'''''                    Case 16: TBeneficiario.Cta_Transf = Codigo
'''''                    Case 17: Cta_Aux = Codigo
'''''                    Case 18: Cta_Gastos = Codigo
'''                  End Select
'''              Next N
'''              Cta_Transf = Ninguno
'''              If Len(TrimStrg(TBeneficiario.Cte_Ahr_Otro) & TrimStrg(TBeneficiario.Cta_Transf)) > 2 Then
'''                 Cta_Transf = TBeneficiario.Cte_Ahr_Otro & " " & TBeneficiario.Cta_Transf
'''              End If
'''              If AdoClientes.Recordset.RecordCount > 0 Then
'''                 AdoClientes.Recordset.MoveFirst
'''                 AdoClientes.Recordset.Find ("CI_RUC = '" & TBeneficiario.CI_RUC & "' ")
'''                 If Not AdoClientes.Recordset.EOF Then
'''                    TBeneficiario.Codigo = AdoClientes.Recordset.Fields("Codigo")
'''                    AdoClientes.Recordset.Fields("T") = TBeneficiario.T
'''                    If IsDate(TBeneficiario.Fecha_N) Then AdoClientes.Recordset.Fields("Fecha_N") = TBeneficiario.Fecha_N
'''                    If IsDate(TBeneficiario.Fecha) Then AdoClientes.Recordset.Fields("Fecha") = TBeneficiario.Fecha
'''                    If Len(TBeneficiario.Cliente) > 1 Then AdoClientes.Recordset.Fields("Cliente") = TBeneficiario.Cliente
'''                    If Len(TBeneficiario.Sexo) > 1 Then AdoClientes.Recordset.Fields("Sexo") = TBeneficiario.Sexo
'''                    If Len(TBeneficiario.Email1) > 1 Then AdoClientes.Recordset.Fields("Email") = TBeneficiario.Email1
'''                    If Len(TBeneficiario.Email2) > 1 Then AdoClientes.Recordset.Fields("Email2") = TBeneficiario.Email2
'''                    If Len(TBeneficiario.Direccion) > 1 Then AdoClientes.Recordset.Fields("Direccion") = TBeneficiario.Direccion
'''                    If Len(TBeneficiario.DirNumero) > 1 Then AdoClientes.Recordset.Fields("DirNumero") = TBeneficiario.DirNumero
'''                    If Len(TBeneficiario.Telefono1) > 1 Then AdoClientes.Recordset.Fields("Telefono") = TBeneficiario.Telefono1
'''                    If Len(TBeneficiario.Celular) > 1 Then AdoClientes.Recordset.Fields("Celular") = TBeneficiario.Celular
'''                    If Len(TBeneficiario.Ciudad) > 1 Then AdoClientes.Recordset.Fields("Ciudad") = TBeneficiario.Ciudad
'''                    If Len(TBeneficiario.Prov) > 1 Then AdoClientes.Recordset.Fields("Prov") = TBeneficiario.Prov
'''                    If Len(TBeneficiario.Pais) > 1 Then AdoClientes.Recordset.Fields("Pais") = TBeneficiario.Pais
'''                    If Len(TBeneficiario.Grupo_No) > 1 Then AdoClientes.Recordset.Fields("Grupo") = MidStrg(TBeneficiario.Grupo_No, 1, 10)
'''                    If Len(TBeneficiario.Profesion) > 1 Then AdoClientes.Recordset.Fields("Profesion") = TBeneficiario.Profesion
'''                    AdoClientes.Recordset.Update
'''                 Else
'''                    Crear_Nuevo = True
'''                 End If
'''              Else
'''                 Crear_Nuevo = True
'''              End If
'''              If Crear_Nuevo Then
'''                 SetAdoAddNew "Clientes"
'''                 SetAdoFields "T", TBeneficiario.T
'''                 SetAdoFields "TD", TBeneficiario.TP
'''                 If Len(TBeneficiario.Codigo) > 1 Then SetAdoFields "Codigo", TBeneficiario.Codigo
'''                 If Len(TBeneficiario.CI_RUC) > 1 Then SetAdoFields "CI_RUC", TBeneficiario.CI_RUC
'''                 If IsDate(TBeneficiario.Fecha) Then SetAdoFields "Fecha", TBeneficiario.Fecha
'''                 If IsDate(TBeneficiario.Fecha_N) Then SetAdoFields "Fecha_N", TBeneficiario.Fecha_N
'''                 If Len(TBeneficiario.Cliente) > 1 Then SetAdoFields "Cliente", TBeneficiario.Cliente
'''                 If Len(TBeneficiario.Sexo) > 1 Then SetAdoFields "Sexo", TBeneficiario.Sexo
'''                 If Len(TBeneficiario.Email1) > 1 Then SetAdoFields "Email", TBeneficiario.Email1
'''                 If Len(TBeneficiario.Email2) > 1 Then SetAdoFields "Email2", TBeneficiario.Email2
'''                 If Len(TBeneficiario.Direccion) > 1 Then SetAdoFields "Direccion", TBeneficiario.Direccion
'''                 If Len(TBeneficiario.DirNumero) > 1 Then SetAdoFields "DirNumero", TBeneficiario.DirNumero
'''                 If Len(TBeneficiario.Telefono1) > 1 Then SetAdoFields "Telefono", TBeneficiario.Telefono1
'''                 If Len(TBeneficiario.Celular) > 1 Then SetAdoFields "Celular", TBeneficiario.Celular
'''                 If Len(TBeneficiario.Ciudad) > 1 Then SetAdoFields "Ciudad", TBeneficiario.Ciudad
'''                 If Len(TBeneficiario.Prov) > 1 Then SetAdoFields "Prov", TBeneficiario.Prov
'''                 If Len(TBeneficiario.Grupo_No) > 1 Then SetAdoFields "Grupo", TBeneficiario.Grupo_No
'''                 If Len(TBeneficiario.Profesion) > 1 Then SetAdoFields "Profesion", TBeneficiario.Profesion
'''                 SetAdoUpdate
'''                 Lista_Clientes_Nuevos = Lista_Clientes_Nuevos _
'''                                       & TBeneficiario.CI_RUC & vbTab _
'''                                       & TBeneficiario.Cliente & vbTab _
'''                                       & TBeneficiario.Grupo_No & vbCrLf
'''              End If
'''             'Creamos los Clientes Rol de Pagos
'''              Crear_Nuevo = False
'''              If AdoAux.Recordset.RecordCount > 0 Then
'''                 AdoAux.Recordset.MoveFirst
'''                 AdoAux.Recordset.Find ("Codigo = '" & TBeneficiario.Codigo & "' ")
'''                 If Not AdoAux.Recordset.EOF Then
'''                    If Len(TBeneficiario.Fecha_A) = 10 Then AdoAux.Recordset.Fields("Fecha") = TBeneficiario.Fecha_A
'''                    If Len(TBeneficiario.Grupo_No) > 1 Then AdoAux.Recordset.Fields("Grupo_Rol") = TBeneficiario.Grupo_No
'''                    If Len(TBeneficiario.Salario) > 1 Then AdoAux.Recordset.Fields("Salario") = TBeneficiario.Salario
'''                    AdoAux.Recordset.Fields("T") = TBeneficiario.T
'''                    AdoAux.Recordset.Fields("SN") = "1"
'''                    AdoAux.Recordset.Fields("Valor_Hora") = Redondear(TBeneficiario.Salario / 240, 2)
'''                    AdoAux.Recordset.Fields("Horas_Sem") = 60
'''                    AdoAux.Recordset.Fields("IESS_Per") = 0.0945
'''                    AdoAux.Recordset.Fields("IESS_Pat") = 0.1215
'''                    AdoAux.Recordset.Fields("Cta_Transferencia") = Cta_Transf
'''                    AdoAux.Recordset.Fields("Codigo_Banco") = TBeneficiario.Cod_Banco
'''                    AdoAux.Recordset.Fields("Pagar_Fondo_Reserva") = Aplica_FP
'''                    AdoAux.Recordset.Fields("Cta_Forma_Pago") = Cta_Aux
'''                    If Len(TBeneficiario.Cte_Ahr_Otro) >= 3 Then
'''                       AdoAux.Recordset.Fields("FP") = "T"
'''                       AdoAux.Recordset.Fields("TC") = "BA"
'''                    Else
'''                       AdoAux.Recordset.Fields("FP") = "E"
'''                       AdoAux.Recordset.Fields("TC") = "CJ"
'''                    End If
'''                    AdoAux.Recordset.Update
'''                 Else
'''                    Crear_Nuevo = True
'''                 End If
'''              Else
'''                 Crear_Nuevo = True
'''              End If
'''              If Crear_Nuevo Then
'''                 SetAdoAddNew "Catalogo_Rol_Pagos"
'''                 If Len(TBeneficiario.Grupo_No) > 1 Then SetAdoFields "Grupo_Rol", TBeneficiario.Grupo_No
'''                 If Len(TBeneficiario.Salario) > 1 Then SetAdoFields "Salario", TBeneficiario.Salario
'''                 SetAdoFields "T", TBeneficiario.T
'''                 SetAdoFields "Codigo", TBeneficiario.Codigo
'''                 SetAdoFields "Fecha", TBeneficiario.Fecha_A
'''                 SetAdoFields "SN", "1"
'''                 SetAdoFields "Valor_Hora", Redondear(TBeneficiario.Salario / 240, 2)
'''                 SetAdoFields "Horas_Sem", 60
'''                 SetAdoFields "IESS_Per", 0.0945
'''                 SetAdoFields "IESS_Pat", 0.1215
'''                 SetAdoFields "Cta_Transferencia", Cta_Transf
'''                 SetAdoFields "Codigo_Banco", TBeneficiario.Cod_Banco
'''                 SetAdoFields "Pagar_Fondo_Reserva", Aplica_FP
'''                 SetAdoFields "Cta_Forma_Pago", Cta_Aux
'''                 If Len(TBeneficiario.Cte_Ahr_Otro) >= 3 Then
'''                    SetAdoFields "FP", "T"
'''                    SetAdoFields "TC", "BA"
'''                 Else
'''                    SetAdoFields "FP", "E"
'''                    SetAdoFields "TC", "CJ"
'''                 End If
'''                 SetAdoUpdate
'''              End If
'''           End If
'''           Me.Caption = "Importar de FlexGrid a Sistema El Beneficiario: " & TBeneficiario.CI_RUC & ": " & i & " de " & Rango.NumFila2
'''      Next i
'''  End With
'''  If Len(Lista_Clientes_Nuevos) > 2 Then
'''     TextoImprimio = Lista_Clientes_Nuevos
'''     Unload FImporta
'''     FInfoError.Show
'''  End If
'''End Sub
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'''Public Sub Importar_Alumnos_Contabilidad()
'''Dim I As Long
'''Dim N As Long
'''Dim NumEmp As Integer
'''Dim Tot_Propinas As Currency
''' 'Empezamos la importacion de las facturas
'''  sSQL = "UPDATE Empresas " _
'''       & "SET Seguro = 0 "
'''  Ejecutar_SQL_SP sSQL
'''  NumTrans = 0
'''  DGExcelAdodc.Visible = False
'''  With AdoExcelAdodc.Recordset
'''      'Barremos toda la Base de Alumnos para saber quien se retiro
'''       For i = 1 To .rows - 1
'''          .Row = i
'''          .Col = 2
'''           Codigo = UCaseStrg(Replace(.Text, "'", ""))
'''           Codigo = TrimStrg(Replace(Codigo, "-", ""))
'''           CI_Representante = Codigo
'''
'''           sSQL = "UPDATE Empresas " _
'''                & "SET Seguro = 1 " _
'''                & "WHERE RUC = '" & CI_Representante & "' "
'''           Ejecutar_SQL_SP sSQL
'''       Next i
'''
'''       sSQL = "UPDATE Empresas " _
'''            & "SET Seguro = 1 " _
'''            & "WHERE Item = '001' "
'''       Ejecutar_SQL_SP sSQL
'''
''''''       sSQL = "UPDATE Empresas " _
''''''            & "SET Seguro = 1 " _
''''''            & "WHERE RUC = '" & CI_Representante & "' "
''''''       Ejecutar_SQL_SP sSQL
'''
'''      'Borramos los alumnos retirados
'''       sSQL = "DELETE * " _
'''            & "FROM Empresas " _
'''            & "WHERE Seguro = 0 "
'''       Ejecutar_SQL_SP sSQL
'''
'''       sSQL = "DELETE * " _
'''            & "FROM Acceso_Empresa " _
'''            & "WHERE Codigo <> '.' "
'''       Ejecutar_SQL_SP sSQL
'''
'''       For i = 1 To .rows - 1
'''          .Row = i
'''           For N = 1 To .cols - 1
'''              .Col = N
'''               Codigo = UCaseStrg(Replace(.Text, "'", ""))
'''               Codigo = TrimStrg(Replace(Codigo, "-", ""))
'''               Select Case N
'''                 Case 2: CI_Representante = Codigo
'''                 Case 3: NombreCliente = Codigo
'''                 Case 4: Grupo_No = TrimStrg(Replace(Codigo, ".", ""))
'''                 Case 5: NivelNo = Codigo
'''                 Case 9: Codigo2 = Codigo
'''                 Case 13: Codigo1 = Codigo
'''                 Case 14: CodigoCli = Codigo
'''               End Select
'''           Next N
'''
'''           sSQL = "SELECT Empresa,Item " _
'''                & "FROM Empresas " _
'''                & "WHERE Item <> '.' " _
'''                & "ORDER BY Item "
'''           Select_Adodc AdoAux, sSQL
'''           For J = 2 To 999
'''               If AdoAux.Recordset.RecordCount > 0 Then
'''                  AdoAux.Recordset.MoveFirst
'''                  AdoAux.Recordset.Find ("Item = '" & Format(J, "000") & "' ")
'''                  If AdoAux.Recordset.EOF Then
'''                     NumEmp = J
'''                     J = 999
'''                  End If
'''               Else
'''                  NumEmp = 2
'''               End If
'''           Next J
'''
'''            sSQL = "SELECT * " _
'''                 & "FROM Accesos " _
'''                 & "WHERE Codigo = '" & CodigoCli & "' "
'''            Select_Adodc AdoAux, sSQL
'''            If AdoAux.Recordset.RecordCount <= 0 Then
'''               SetAdoAddNew "Accesos", True
'''               SetAdoFields "Codigo", CodigoCli
'''               SetAdoFields "Usuario", CI_Representante
'''               SetAdoFields "Clave", CI_Representante
'''               SetAdoFields "Nombre_Completo", UCaseStrg(NombreCliente)
'''               SetAdoFields "Supervisor", True
'''               SetAdoUpdate
'''
'''               SetAdoAddNew "Empresas"
'''               SetAdoFields "Grupo", Format$(NumEmp, "000")
'''               SetAdoFields "Item", Format$(NumEmp, "000")
'''               SetAdoFields "Pais", "593"
'''               SetAdoFields "Empresa", "ESTUDIANTE " & UCaseStrg(NombreCliente)
'''               SetAdoFields "Nombre_Comercial", "SISTEMA PARA ESTUDIANTES: " & CI_Representante
'''               SetAdoFields "Gerente", UCaseStrg(NombreCliente)
'''               SetAdoFields "RUC", CI_Representante
'''               SetAdoFields "Direccion", NivelNo
'''               SetAdoFields "SubDir", Grupo_No
'''               SetAdoFields "Logo_Tipo", Codigo1
'''               SetAdoFields "Contador", UCaseStrg(NombreCliente)
'''               SetAdoFields "Formato_Cuentas", "C.C.CC.CC.CC.CCC"
'''               SetAdoFields "Formato_Inventario", "CC.CC.CCC.CCCCCC"
'''               SetAdoFields "Formato_Activo", "CC.CC.CCC.CCCCCC"
'''               SetAdoUpdate
'''            End If
'''            sSQL = "SELECT TOP 1 * " _
'''                 & "FROM Modulos " _
'''                 & "WHERE Modulo <> '.' " _
'''                 & "ORDER BY Modulo DESC "
'''            Select_Adodc AdoAux, sSQL
'''            If AdoAux.Recordset.RecordCount > 0 Then K = Val(AdoAux.Recordset.Fields("Modulo"))
'''            If K <= 0 Then K = 99
'''            For J = 1 To K
'''                SetAdoAddNew "Acceso_Empresa"
'''                SetAdoFields "Modulo", Format(J, "00")
'''                SetAdoFields "Item", Format$(NumEmp, "000")
'''                SetAdoFields "Codigo", CodigoCli
'''                SetAdoUpdate
'''            Next J
'''           Me.Caption = Format$(i / Rango.NumFila2, "00%") & " Importando: " & NombreCliente & " Alumnos a Contabilidad "
'''      Next i
'''  End With
'''  Me.Caption = Format$(i / Rango.NumFila2, "00%") & " Importando: " & NombreCliente & " Alumnos a Contabilidad "
'''  DGExcelAdodc.Visible = True
'''End Sub
''-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'''Public Sub Diarios_Automaticos_Ventas()
'''Dim CodRetBien As Byte
'''Dim CodRetServ As Byte
'''Dim AsientoIni As Boolean
'''Dim AsientoFin As Boolean
'''Dim I As Long
'''Dim N As Long
'''Dim SecuencialF As Long
'''Dim SecuencialR As Long
'''Dim PorcRet As Single
'''Dim PorcIVAB As Single
'''Dim PorcIVAS As Single
'''Dim Valor As Currency
'''Dim TotalIVAB As Currency
'''Dim TotalIVAS As Currency
'''Dim Tot_Propinas As Currency
'''Dim SubTotalRet As Currency
'''Dim TotalRetFuente As Currency
'''Dim TotalRetIVABien As Currency
'''Dim TotalRetIVAServ As Currency
'''
'''Dim DigCta As String
'''Dim SerieF1 As String
'''Dim SerieF2 As String
'''Dim SerieR1 As String
'''Dim SerieR2 As String
'''Dim Cta_Gasto As String
'''Dim FormaPago As String
'''Dim TipoSustento As String
'''Dim CtaRetFuente As String
'''Dim CtaRetIVABien As String
'''Dim CtaRetIVAServ As String
'''Dim ConceptoDiario As String
'''
''' 'Encerar_Facturas
'''  Trans_No = 181
''' 'Insertamos asiento Gastos+IVA contra CxP Proveedores
'''  Eliminar_Asientos_SP True
'''  IniciarAsientosDe DGAsiento, AdoAsiento
'''  Bandera = False
'''  Evaluar = True
'''  TextoImprimio = ""
'''  ConceptoDiario = ""
'''  Co.Fecha = FechaSistema
'''  Co.Beneficiario = Ninguno
'''  Co.CodigoB = Ninguno
'''  CodigoCliente = Ninguno
'''
'''  With AdoExcelAdodc.Recordset
'''       For i = 1 To .rows - 1
'''          'Fila actual
'''          .Row = i
'''          'Iniciamos la recoleccion de datos
'''          .Col = 1
'''           Cta = TrimStrg(.Text)
'''          .Col = 2
'''           Mifecha = TrimStrg(.Text)
'''          .Col = 3
'''           CodigoA = TrimStrg(.Text)
'''          .Col = 4
'''           Ln_No = Val(TrimStrg(.Text))
'''          .Col = 5
'''           Codigo1 = TrimStrg(.Text)
'''          .Col = 6
'''           Codigo2 = TrimStrg(Replace(.Text, "'", ""))
'''          .Col = 7
'''           Codigo3 = TrimStrg(Replace(.Text, "'", ""))
'''          .Col = 8
'''           FechaTexto = TrimStrg(.Text)
'''          .Col = 9
'''           Codigo4 = TrimStrg(.Text)
'''          .Col = 10
'''           CodigoN = TrimStrg(.Text)
'''          .Col = 11
'''           Debe = Redondear(Val(TrimStrg(.Text)), 2)
'''          .Col = 12
'''           Haber = Redondear(Val(TrimStrg(.Text)), 2)
'''
'''          'Codigo Proveedor por default
'''           CodigoCliente = "9999999999"
'''          'Fila actual
'''          .Row = i
'''           PorcRet = 0
'''           PorcIVAB = 0
'''           PorcIVAS = 0
'''           DigCta = ""
'''          'Iniciamos la recoleccion de datos
'''          .Col = 1
'''           TipoDoc = TrimStrg(.Text)
'''          .Col = 2
'''           TipoSustento = "01"
'''           If Len(TrimStrg(.Text)) = 2 Then TipoSustento = TrimStrg(.Text)
'''          .Col = 3
'''           Mifecha = TrimStrg(.Text)
'''          .Col = 4
'''          'RUC/Cedula/Consumidor Final
'''           Codigo = TrimStrg(Replace(.Text, "'", ""))
'''           Co.RUC_CI = Codigo
'''           If IsNumeric(Codigo) And Len(Codigo) = 12 Then Codigo = "0" & Codigo
'''           If IsNumeric(Codigo) And Len(Codigo) = 9 Then Codigo = "0" & Codigo
'''           CI_Representante = Codigo
'''           If Len(Codigo) > 1 Then
'''              DigVerif = Digito_Verificador( Codigo)
'''              Caracter = MidStrg(Codigo, 10, 1)
'''              CodigoCliente = Tipo_RUC_CI.Codigo_RUC_CI
'''           End If
'''          .Col = 5
'''           Autorizacion = TrimStrg(.Text)
'''          .Col = 6
'''           Codigo = TrimStrg(Replace(.Text, "'", ""))
'''           SerieR1 = Format$(Val(TrimStrg(MidStrg(Codigo, 1, 3))), "000")
'''           SerieR2 = Format$(Val(TrimStrg(MidStrg(Codigo, 5, 3))), "000")
'''           SecuencialR = Val(TrimStrg(MidStrg(Codigo, 9, 10)))
'''          .Col = 7
'''           Codigo = TrimStrg(Replace(.Text, "'", ""))
'''           SerieF1 = Format$(Val(TrimStrg(MidStrg(Codigo, 1, 3))), "000")
'''           SerieF2 = Format$(Val(TrimStrg(MidStrg(Codigo, 5, 3))), "000")
'''           SecuencialF = Val(TrimStrg(MidStrg(Codigo, 9, 10)))
'''          .Col = 8
'''           NombreCliente = UCaseStrg(TrimStrg(.Text))
'''          'MsgBox NombreCliente
'''          .Col = 9
'''           ConceptoDiario = TrimStrg(.Text)
'''          .Col = 10
'''           Total_Sin_No_IVA = Redondear(Val(TrimStrg(.Text)), 2)
'''          'No Objeto de IVA
'''          .Col = 11
'''           Total_Sin_IVA = Redondear(Val(TrimStrg(.Text)), 2)
'''          .Col = 12
'''           Total_Con_IVA = Redondear(Val(TrimStrg(.Text)), 2)
'''          .Col = 13
'''           Total_IVA = Redondear(Val(TrimStrg(.Text)), 2)
'''          .Col = 14
'''           Total = Redondear(Val(TrimStrg(.Text)), 2)
'''           SubTotal = Total_Con_IVA + Total_Sin_IVA + Total_Sin_No_IVA
'''          .Col = 15
'''           TotalIVAB = Redondear(Val(TrimStrg(.Text)), 2)
'''          .Col = 16
'''           TotalIVAS = Redondear(Val(TrimStrg(.Text)), 2)
'''          .Col = 17
'''           Valor = Redondear(Val(TrimStrg(.Text)), 2)
'''           If Valor > 0 Then
'''              CtaRetFuente = Leer_Seteos_Ctas("Cta_Ret_1")
'''              TotalRetFuente = Valor
'''              PorcRet = 1
'''           End If
'''          .Col = 18
'''           Valor = Redondear(Val(TrimStrg(.Text)), 2)
'''           If Valor > 0 Then
'''              CtaRetFuente = Leer_Seteos_Ctas("Cta_Ret_2")
'''              TotalRetFuente = Valor
'''              PorcRet = 2
'''           End If
'''          .Col = 19
'''           Valor = Redondear(Val(TrimStrg(.Text)), 2)
'''           If Valor > 0 Then
'''              CtaRetFuente = Leer_Seteos_Ctas("Cta_Ret_5")
'''              TotalRetFuente = Valor
'''              PorcRet = 5
'''           End If
'''          .Col = 20
'''           Valor = Redondear(Val(TrimStrg(.Text)), 2)
'''           If Valor > 0 Then
'''              CtaRetFuente = Leer_Seteos_Ctas("Cta_Ret_8")
'''              TotalRetFuente = Valor
'''              PorcRet = 8
'''           End If
'''          .Col = 21
'''           Valor = Redondear(Val(TrimStrg(.Text)), 2)
'''           If Valor > 0 Then
'''              CtaRetFuente = Leer_Seteos_Ctas("Cta_Ret_10")
'''              TotalRetFuente = Valor
'''              PorcRet = 10
'''           End If
'''          .Col = 22
'''           Valor = Redondear(Val(TrimStrg(.Text)), 2)
'''           If Valor > 0 Then
'''              CtaRetFuente = Leer_Seteos_Ctas("Cta_Ret_25")
'''              TotalRetFuente = Valor
'''              PorcRet = 25
'''           End If
'''          .Col = 23
'''           Valor = Redondear(Val(TrimStrg(.Text)), 2)
'''           If Valor > 0 Then
'''              CtaRetIVABien = Leer_Seteos_Ctas("Cta_Ret_IVA_10")
'''              TotalRetIVABien = Valor
'''              PorcIVAB = 10
'''           End If
'''          .Col = 24
'''           Valor = Redondear(Val(TrimStrg(.Text)), 2)
'''           If Valor > 0 Then
'''              CtaRetIVAServ = Leer_Seteos_Ctas("Cta_Ret_IVA_20")
'''              TotalRetIVAServ = Valor
'''              PorcIVAS = 20
'''           End If
'''          .Col = 25
'''           Valor = Redondear(Val(TrimStrg(.Text)), 2)
'''           If Valor > 0 Then
'''              CtaRetIVABien = Leer_Seteos_Ctas("Cta_Ret_IVA_30")
'''              TotalRetIVABien = Valor
'''              PorcIVAB = 30
'''           End If
'''          .Col = 26
'''           Valor = Redondear(Val(TrimStrg(.Text)), 2)
'''           If Valor > 0 Then
'''              CtaRetIVABien = Leer_Seteos_Ctas("Cta_Ret_IVA_50")
'''              TotalRetIVABien = Valor
'''              PorcIVAB = 50
'''           End If
'''          .Col = 27
'''           Valor = Redondear(Val(TrimStrg(.Text)), 2)
'''           If Valor > 0 Then
'''              CtaRetIVAServ = Leer_Seteos_Ctas("Cta_Ret_IVA_70")
'''              TotalRetIVAServ = Valor
'''              PorcIVAS = 70
'''           End If
'''          .Col = 28
'''           Valor = Redondear(Val(TrimStrg(.Text)), 2)
'''           If Valor > 0 Then
'''              CtaRetIVABien = Leer_Seteos_Ctas("Cta_Ret_IVAB_100")
'''              TotalRetIVABien = Valor
'''              PorcIVAB = 100
'''           End If
'''          .Col = 29
'''           Valor = Redondear(Val(TrimStrg(.Text)), 2)
'''           If Valor > 0 Then
'''              CtaRetIVAServ = Leer_Seteos_Ctas("Cta_Ret_IVAS_100")
'''              TotalRetIVAServ = Valor
'''              PorcIVAS = 100
'''           End If
'''         '.Col = 30
'''         ' Total Retenciones
'''         '.Col = 31
'''         ' Total Abonos
'''          .Col = 32
'''           CodigoP = Val(TrimStrg(.Text))     'Codigo de Retencion
'''          .Col = 33
'''           AutorizaRet = TrimStrg(.Text)
'''          .Col = 34
'''           FechaTexto = TrimStrg(.Text)       'Caducidad de la Factura
'''          .Col = 35
'''           Cta_Gasto = TrimStrg(.Text)
'''          .Col = 36
'''           Cta_CajaG = TrimStrg(.Text)
'''          .Col = 37
'''           DigCta = TrimStrg(.Text)
'''           If Len(DigCta) > 1 Then CtaRetFuente = CtaRetFuente & "." & DigCta
'''          .Col = 38
'''           FormaPago = "01"
'''           If Len(TrimStrg(.Text)) = 2 Then FormaPago = TrimStrg(.Text)
'''          .Col = 39
'''           SubModuloGasto = TrimStrg(.Text)
'''         '.Col = 40
'''         ' SubModuloCxCxP = TrimStrg(.Text)
''''          Cuenta = .Fields("Cuenta")
''''          SubCta = .Fields("TC")
''''          Moneda_US = .Fields("ME")
''''          TipoCta = .Fields("DG")
''''          TipoPago = .Fields("Tipo_Pago")
'''
'''           Codigo = Leer_Cta_Catalogo(Cta_Gasto)
'''           Codigo1 = Leer_Cta_Catalogo(Cta_CajaG)
'''           Codigo2 = Leer_SubCta_Modulo(SubModuloGasto)
'''           If IsDate(Mifecha) And Codigo <> Ninguno And Codigo1 <> Ninguno And Codigo2 <> Ninguno Then
'''              Co.Concepto = NombreCliente
'''              If Len(SerieF1 & SerieF2) = 6 Then
'''                 Co.Concepto = Co.Concepto & ", Documento No. " & SerieF1 & SerieF2 & "-" & Format(SecuencialF, "000000000")
'''              End If
'''              If Len(SerieR1 & SerieR2) = 6 Then
'''                 Co.Concepto = Co.Concepto & ", Retencion No. " & SerieR1 & SerieR2 & "-" & Format(SecuencialR, "000000000") _
'''                             & " Codigo Ret: " & CodigoP
'''              End If
'''              Co.Concepto = Co.Concepto & ", Por: " & ConceptoDiario
'''              If CI_Representante = "9999999999999" Then CI_Representante = Ninguno
'''              If CI_Representante <> Ninguno Then
'''                 sSQL = "SELECT Codigo,Cliente " _
'''                      & "FROM Clientes " _
'''                      & "WHERE Codigo = '" & CodigoCliente & "' "
'''                 Select_Adodc AdoAux, sSQL
'''                 If AdoAux.Recordset.RecordCount <= 0 Then
'''                    SetAdoAddNew "Clientes"
'''                    SetAdoFields "T", Normal
'''                    SetAdoFields "Codigo", CodigoCliente
'''                    SetAdoFields "TD", Tipo_RUC_CI.Tipo_Beneficiario
'''                    SetAdoFields "CI_RUC", CI_Representante
'''                    SetAdoFields "Cliente", UCaseStrg(NombreCliente)
'''                    SetAdoFields "Fecha", FechaSistema
'''                    SetAdoFields "Direccion", "SD"
'''                    SetAdoFields "DirNumero", "SN"
'''                    SetAdoFields "Ciudad", "QUITO"
'''                    SetAdoFields "Prov", "17"
'''                    SetAdoFields "Pais", "593"
'''                    SetAdoFields "CodigoU", CodigoUsuario
'''                    SetAdoUpdate
'''                 End If
'''              End If
'''             'Insertamos asiento Gastos+IVA contra CxP Proveedores
'''              Eliminar_Asientos_SP True
'''              IniciarAsientosDe DGAsiento, AdoAsiento
'''              Codigo2 = Leer_SubCta_Modulo(SubModuloGasto)
'''              Codigo = Leer_Cta_Catalogo(Cta_Gasto)
'''              If SubCta = "G" Then
'''                 SetAdoAddNew "Asiento_SC"
'''                 SetAdoFields "Codigo", Codigo2
'''                 SetAdoFields "Beneficiario", UCaseStrg(NombreCliente)
'''                 SetAdoFields "DH", "1"
'''                 SetAdoFields "Valor", SubTotal
'''                 SetAdoFields "FECHA_V", Mifecha
'''                 SetAdoFields "TC", SubCta
'''                 SetAdoFields "Cta", Cta_Gasto
'''                 SetAdoFields "TM", "1"
'''                 SetAdoFields "T_No", Trans_No
'''                 SetAdoFields "SC_No", 1
'''                 SetAdoFields "Item", NumEmpresa
'''                 SetAdoFields "CodigoU", CodigoUsuario
'''                 SetAdoUpdate
'''              End If
'''
'''              Codigo1 = Leer_Cta_Catalogo(Cta_CajaG)
'''              If SubCta = "C" Or SubCta = "P" Then
'''                 SetAdoAddNew "Asiento_SC"
'''                 SetAdoFields "Codigo", CodigoCliente
'''                 SetAdoFields "Serie", SerieF1 & SerieF2
'''                 SetAdoFields "Factura", SecuencialF
'''                 SetAdoFields "Beneficiario", UCaseStrg(NombreCliente)
'''                 SetAdoFields "DH", "2"
'''                 SetAdoFields "Valor", SubTotal + Total_IVA
'''                 SetAdoFields "FECHA_V", Mifecha
'''                 SetAdoFields "TC", SubCta
'''                 SetAdoFields "Cta", Cta_CajaG
'''                 SetAdoFields "TM", "1"
'''                 SetAdoFields "T_No", Trans_No
'''                 SetAdoFields "SC_No", 1
'''                 SetAdoFields "Item", NumEmpresa
'''                 SetAdoFields "CodigoU", CodigoUsuario
'''                 SetAdoUpdate
'''              End If
'''              CodRetBien = 0
'''              CodRetServ = 0
'''              sSQL = "SELECT Codigo " _
'''                   & "FROM Tabla_Por_IVA " _
'''                   & "WHERE Porc = '" & CStr(PorcIVAB) & "' "
'''              Select_Adodc AdoAux, sSQL
'''              If AdoAux.Recordset.RecordCount > 0 Then CodRetBien = AdoAux.Recordset.Fields("Codigo")
'''              sSQL = "SELECT Codigo " _
'''                   & "FROM Tabla_Por_IVA " _
'''                   & "WHERE Porc = '" & CStr(PorcIVAS) & "' "
'''              Select_Adodc AdoAux, sSQL
'''              If AdoAux.Recordset.RecordCount > 0 Then CodRetServ = AdoAux.Recordset.Fields("Codigo")
'''             'Grabo en el Asiento_Compras
'''              SetAdoAddNew "Asiento_Compras"
'''              SetAdoFields "IdProv", CodigoCliente
'''              SetAdoFields "DevIva", "N"
'''              SetAdoFields "CodSustento", TipoSustento
'''              SetAdoFields "TipoComprobante", TipoDoc
'''              SetAdoFields "Establecimiento", SerieF1
'''              SetAdoFields "PuntoEmision", SerieF2
'''              SetAdoFields "Secuencial", SecuencialF
'''              SetAdoFields "Autorizacion", Autorizacion
'''              SetAdoFields "FechaEmision", Mifecha
'''              SetAdoFields "FechaRegistro", Mifecha
'''              SetAdoFields "FechaCaducidad", FechaTexto
'''              SetAdoFields "BaseNoObjIVA", Total_Sin_No_IVA
'''              SetAdoFields "BaseImponible", Total_Sin_IVA
'''              SetAdoFields "BaseImpGrav", Total_Con_IVA
'''              SetAdoFields "PorcentajeIva", 2
'''              SetAdoFields "MontoIva", Total_IVA
'''
'''              SetAdoFields "Porc_Bienes", PorcIVAB
'''              SetAdoFields "MontoIvaBienes", TotalIVAB
'''              SetAdoFields "PorRetBienes", CodRetBien
'''              SetAdoFields "ValorRetBienes", TotalRetIVABien
'''
'''              SetAdoFields "Porc_Servicios", PorcIVAS
'''              SetAdoFields "MontoIvaServicios", TotalIVAB
'''              SetAdoFields "PorRetServicios", CodRetServ
'''              SetAdoFields "ValorRetServicios", TotalRetIVAServ
'''
'''              SetAdoFields "DocModificado", "0"
'''              SetAdoFields "EstabModificado", "000"
'''              SetAdoFields "PtoEmiModificado", "000"
'''              SetAdoFields "SecModificado", "0000000"
'''              SetAdoFields "AutModificado", "0000000000"
'''              SetAdoFields "ContratoPartidoPolitico", "0000000000"
'''
'''              SetAdoFields "PagoLocExt", "01"
'''              SetAdoFields "PaisEfecPago", "NA"
'''              SetAdoFields "AplicConvDobTrib", "NA"
'''              SetAdoFields "PagExtSujRetNorLeg", "NA"
'''              SetAdoFields "FormaPago", FormaPago
'''              SetAdoFields "A_No", 1
'''              SetAdoFields "T_No", Trans_No
'''              SetAdoFields "CodigoU", CodigoUsuario
'''              SetAdoFields "Cta_Bienes", CtaRetIVABien
'''              SetAdoFields "Cta_Servicio", CtaRetIVAServ
'''              SetAdoUpdate
'''
'''             'Grabo Asiento_Air
'''              SetAdoAddNew "Asiento_Air"
'''              SetAdoFields "CodRet", CodigoP
'''              SetAdoFields "Detalle", Ninguno
'''              SetAdoFields "BaseImp", SubTotal
'''              SetAdoFields "Porcentaje", Redondear(PorcRet / 100, 2)
'''              SetAdoFields "ValRet", TotalRetFuente
'''              SetAdoFields "EstabRetencion", SerieR1
'''              SetAdoFields "PtoEmiRetencion", SerieR2
'''              SetAdoFields "SecRetencion", SecuencialR
'''              SetAdoFields "AutRetencion", AutorizaRet
'''              SetAdoFields "FechaEmiRet", Mifecha
'''              SetAdoFields "Cta_Retencion", CtaRetFuente '?????
'''              SetAdoFields "EstabFactura", SerieF1
'''              SetAdoFields "PuntoEmiFactura", SerieF1
'''              SetAdoFields "Factura_No", SecuencialF
'''              SetAdoFields "IdProv", CodigoCliente
'''              SetAdoFields "A_No", 1
'''              SetAdoFields "T_No", Trans_No
'''              SetAdoFields "Tipo_Trans", "C"
'''              SetAdoUpdate
'''
'''             ' SubTotalRet = TotalIVAB + TotalIVAS + TotalRetFuente
'''              SubTotalRet = TotalRetIVABien + TotalRetIVAServ + TotalRetFuente
'''              Codigo1 = Leer_Cta_Catalogo(Cta_CajaG)
'''              If SubCta = "C" Or SubCta = "P" Then
'''                 SetAdoAddNew "Asiento_SC"
'''                 SetAdoFields "Codigo", CodigoCliente
'''                 SetAdoFields "Serie", SerieF1 & SerieF2
'''                 SetAdoFields "Factura", SecuencialF
'''                 SetAdoFields "Beneficiario", UCaseStrg(NombreCliente)
'''                 SetAdoFields "DH", "1"
'''                 SetAdoFields "Valor", SubTotalRet
'''                 SetAdoFields "FECHA_V", Mifecha
'''                 SetAdoFields "TC", SubCta
'''                 SetAdoFields "Cta", Cta_CajaG
'''                 SetAdoFields "TM", "1"
'''                 SetAdoFields "T_No", Trans_No
'''                 SetAdoFields "SC_No", 1
'''                 SetAdoFields "Item", NumEmpresa
'''                 SetAdoFields "CodigoU", CodigoUsuario
'''                 SetAdoUpdate
'''              End If
'''             'Grabacion del Comprobante del Gasto contra CxP o Caja
'''              DetalleComp = Ninguno
'''              Debe = SubTotal + Total_IVA
'''              InsertarAsientos AdoAsiento, Cta_Gasto, 0, SubTotal, 0
'''              InsertarAsientos AdoAsiento, Cta_IVA_Inventario, 0, Total_IVA, 0
'''              InsertarAsientos AdoAsiento, Cta_CajaG, 0, 0, SubTotal + Total_IVA
'''
'''             'Insertamos el asiento contable de Retenciones
'''              InsertarAsientos AdoAsiento, Cta_CajaG, 0, SubTotalRet, 0
'''              DetalleComp = "Retencion IVA del " & PorcIVAB & "%, Factura No. " & SerieF1 & SerieF2 & "-" & Format(SecuencialF, "000000000")
'''              InsertarAsientos AdoAsiento, CtaRetIVABien, 0, 0, TotalRetIVABien
'''              DetalleComp = "Retencion IVA del " & PorcIVAS & "%, Factura No. " & SerieF1 & SerieF2 & "-" & Format(SecuencialF, "000000000")
'''              InsertarAsientos AdoAsiento, CtaRetIVAServ, 0, 0, TotalRetIVAServ
'''              DetalleComp = "Retencion del " & PorcRet & "% (" & CodigoP & ") No. " & SerieR1 & SerieR2 & "-" & Format(SecuencialR, "000000000")
'''              InsertarAsientos AdoAsiento, CtaRetFuente, 0, 0, TotalRetFuente
'''              FechaComp = Mifecha
'''              NumComp = ReadSetDataNum("Diario", True, True)
'''              Co.Cotizacion = 0
'''              Co.Beneficiario = UCaseStrg(NombreCliente)
'''              Co.Total_Banco = 0
'''              Co.RetNueva = True
'''              If (Len(SerieR1 & SerieR1) = 6) And (SecuencialR > 0) Then
'''                 Co.Serie_R = SerieR1 & SerieR1
'''                 Co.Retencion = SecuencialR
'''              Else
'''                 Co.Serie_R = Ninguno
'''                 Co.Retencion = 0
'''              End If
'''              Co.T = Normal
'''              Co.TP = CompDiario
'''              Co.Fecha = Mifecha
'''              Co.Numero = NumComp
'''              Co.CodigoB = CodigoCliente
'''              Co.Efectivo = Debe
'''              Co.Monto_Total = Debe
'''              Co.T_No = Trans_No
'''              Co.Usuario = CodigoUsuario
'''              Co.Item = NumEmpresa
'''              GrabarComprobante Co
'''           Else
'''             'No se puede ingresar
'''              TextoImprimio = TextoImprimio _
'''                            & "Fecha: " & Mifecha & ", Beneficiario: " & NombreCliente & ", Cta Gato: " & Cta_Gasto _
'''                            & ", Contra Cta: " & Cta_CajaG & ",SubModulo: " & SubModuloGasto & vbCrLf
'''           End If
'''           Me.Caption = "Revisando Datos en el excel: " & i & " de " & .rows - 1 & ", Fecha: " & Mifecha & ", Proveedor: " & NombreCliente
'''      Next i
'''      Eliminar_Asientos_SP True
'''  End With
'''  If TextoImprimio <> "" Then FInfoError.Show
'''End Sub
''-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'''Public Sub Diarios_Automaticos()
'''Dim I As Long
'''Dim N As Long
'''Dim DiarioInicial As Long
'''
'''  Trans_No = 190
'''  DiarioInicial = 0
'''  With AdoExcelAdodc.Recordset
'''      'MsgBox .Rows & vbCrLf & .Cols
'''       For i = 1 To .rows - 1
'''           TBeneficiario.T = Normal
'''           TBeneficiario.FA = True
'''           TBeneficiario.TP = "E"
'''           TBeneficiario.Codigo = Ninguno
'''           TBeneficiario.CI_RUC = "9999999999999"
'''           TBeneficiario.Fecha = FechaSistema
'''           TBeneficiario.Fecha_N = FechaSistema
'''           TBeneficiario.Cliente = Ninguno
'''           TBeneficiario.Sexo = Ninguno
'''           TBeneficiario.Email1 = Ninguno
'''           TBeneficiario.Email2 = Ninguno
'''           TBeneficiario.Direccion = "SD"
'''           TBeneficiario.DirNumero = "SN"
'''           TBeneficiario.Telefono1 = "022000000"
'''           TBeneficiario.Celular = "0990000000"
'''           TBeneficiario.Ciudad = "QUITO"
'''           TBeneficiario.Prov = "17"
'''           TBeneficiario.Pais = "593"
'''           TBeneficiario.Grupo_No = "NUEVOS"
'''           TBeneficiario.Cod_Ejec = Ninguno
'''           TBeneficiario.Cta_CxP = Ninguno
'''          .Row = i
'''          .Col = 2
'''           Codigo = Replace(.Text, "'", "")
'''           Codigo = Replace(Codigo, "-", "")
'''           Codigo = TrimStrg(Codigo)
'''           If Codigo = "" Then Codigo = Ninguno
'''          'RUC/Cedula/Codigo Alumno/Consumidor Final
'''           If Len(Codigo) > 1 Then
'''              If IsNumeric(Codigo) Then
'''                 If Len(Codigo) < 9 Then
'''                    Codigo = Format$(Val(Codigo), "00000000")
'''                 ElseIf Len(Codigo) = 9 Then
'''                    Codigo = "0" & Codigo
'''                 ElseIf Len(Codigo) = 11 Then
'''                    Codigo = "00" & Codigo
'''                 ElseIf Len(Codigo) = 12 Then
'''                    Codigo = "0" & Codigo
'''                 End If
'''              End If
'''              TBeneficiario.CI_RUC = Codigo
'''              DigVerif = Digito_Verificador( Codigo)
'''              Caracter = MidStrg(Codigo, 10, 1)
'''              TBeneficiario.TD_Rep = Tipo_RUC_CI.Tipo_Beneficiario
'''              TBeneficiario.TP = Tipo_RUC_CI.Tipo_Beneficiario
'''              TBeneficiario.Codigo = Tipo_RUC_CI.Codigo_RUC_CI
'''           End If
'''           If TBeneficiario.Codigo <> Ninguno Then
'''              For N = 1 To .cols - 1
'''                 .Col = N
'''                  Codigo = Replace(.Text, "'", "")
'''                  Codigo = Replace(Codigo, "-", "")
'''                  Codigo = TrimStrg(Codigo)
'''                  If Codigo = "" Then Codigo = Ninguno
'''                  Select Case N
'''                    Case 1: Co.Fecha = Codigo                'Fecha
'''                    Case 2: TBeneficiario.CI_RUC = Codigo
'''                    Case 3: Co.Cta_Banco = Codigo  'Fecha_N
'''                    Case 4: TBeneficiario.Cliente = UCaseStrg(TrimStrg(MidStrg(Codigo, 1, 80)))  'Cliente
'''                    Case 5: Co.Concepto = TrimStrg(Codigo)   'Concepto
'''                    Case 6: Total = Redondear(Val(Codigo), 2)
'''                    Case 7: Co.Ctas_Modificar = Codigo
'''                  End Select
'''              Next N
'''              sSQL = "SELECT * " _
'''                   & "FROM Clientes " _
'''                   & "WHERE CI_RUC = '" & TBeneficiario.CI_RUC & "' "
'''              Select_Adodc AdoClientes, sSQL
'''              If AdoClientes.Recordset.RecordCount > 0 Then
'''                 TBeneficiario.Codigo = AdoClientes.Recordset.Fields("Codigo")
'''              Else
'''                 SetAdoAddNew "Clientes"
'''                 If Len(TBeneficiario.T) > 0 Then SetAdoFields "T", TBeneficiario.T
'''                 If Len(TBeneficiario.TD_Rep) > 0 Then SetAdoFields "TD", TBeneficiario.TD_Rep
'''                 If Len(TBeneficiario.Codigo) > 1 Then SetAdoFields "Codigo", TBeneficiario.Codigo
'''                 If Len(TBeneficiario.CI_RUC) > 1 Then SetAdoFields "CI_RUC", TBeneficiario.CI_RUC
'''                 If IsDate(TBeneficiario.Fecha) > 1 Then SetAdoFields "Fecha", TBeneficiario.Fecha
'''                 If IsDate(TBeneficiario.Fecha_N) > 1 Then SetAdoFields "Fecha_N", TBeneficiario.Fecha_N
'''                 If Len(TBeneficiario.Cliente) > 1 Then SetAdoFields "Cliente", TBeneficiario.Cliente
'''                 If Len(TBeneficiario.Sexo) > 1 Then SetAdoFields "Sexo", TBeneficiario.Sexo
'''                 If Len(TBeneficiario.Email1) > 1 Then SetAdoFields "Email", TBeneficiario.Email1
'''                 If Len(TBeneficiario.Email2) > 1 Then SetAdoFields "Email2", TBeneficiario.Email2
'''                 If Len(TBeneficiario.Direccion) > 1 Then SetAdoFields "Direccion", TBeneficiario.Direccion
'''                 If Len(TBeneficiario.DirNumero) > 1 Then SetAdoFields "DirNumero", TBeneficiario.DirNumero
'''                 If Len(TBeneficiario.Telefono1) > 1 Then SetAdoFields "Telefono", TBeneficiario.Telefono1
'''                 If Len(TBeneficiario.Celular) > 1 Then SetAdoFields "Celular", TBeneficiario.Celular
'''                 If Len(TBeneficiario.Ciudad) > 1 Then SetAdoFields "Ciudad", TBeneficiario.Ciudad
'''                 If Len(TBeneficiario.Prov) > 1 Then SetAdoFields "Prov", TBeneficiario.Prov
'''                 If Len(TBeneficiario.Grupo_No) > 1 Then SetAdoFields "Grupo", TBeneficiario.Grupo_No
'''                 If Len(TBeneficiario.Cod_Ejec) > 1 Then SetAdoFields "Cod_Ejec", TBeneficiario.Cod_Ejec
'''                 If Len(TBeneficiario.Cta_CxP) > 1 Then SetAdoFields "Cta_CxP", TBeneficiario.Cta_CxP
'''                 SetAdoUpdate
'''              End If
'''           End If
'''          'Generamos el Asiento
'''           FechaComp = Co.Fecha
'''           IniciarAsientosDe DGAsiento, AdoAsiento
'''          'Insertamos las transacciones
'''           InsertarAsientos AdoAsiento, Co.Cta_Banco, 0, Total, 0
'''           InsertarAsientos AdoAsiento, Co.Ctas_Modificar, 0, 0, Total
'''           NumComp = ReadSetDataNum("Diario", True, True)
'''           If DiarioInicial = 0 Then DiarioInicial = NumComp
'''           DiarioCaja = NumComp
'''          'Grabacion del Comprobante
'''           Co.T = Normal
'''           Co.TP = CompDiario
'''           Co.Numero = NumComp
'''           Co.CodigoB = TBeneficiario.Codigo
'''           Co.Efectivo = Total
'''           Co.Monto_Total = Total
'''           Co.T_No = Trans_No
'''           Co.Usuario = CodigoUsuario
'''           Co.Item = NumEmpresa
'''           GrabarComprobante Co
'''           Me.Caption = "Importar de FlexGrid a Sistema El Beneficiario: " & TBeneficiario.CI_RUC & ": " & i & " de " & Rango.NumFila2
'''      Next i
'''      If .rows >= 1 Then
'''          Cadena = "Generacion de CD desde el " & DiarioInicial & " al " & Co.Numero
'''          Control_Procesos Normal, Cadena
'''          MsgBox Cadena
'''      End If
'''  End With
'''
'''End Sub
''-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'''Public Sub Diarios_Automaticos_Contabilidad_Externa()
'''Dim I As Long
'''Dim N As Long
'''
'''Dim DiarioInicial As Long
'''Dim NumCompBata As String
'''
'''  Trans_No = 185
'''  Eliminar_Asientos_SP True
'''  DiarioInicial = 0
'''  NoCheque = Ninguno
'''  SumaDebe = 0
'''  SumaHaber = 0
'''  Co.Fecha = MBFechaI
'''  FechaComp = Co.Fecha
'''  DetalleComp = Ninguno
'''  With AdoExcelAdodc.Recordset
'''      'MsgBox .Rows & vbCrLf & .Cols
'''       Datos_Default_Beneficiario
'''      'Generamos el Asiento
'''       IniciarAsientosDe DGAsiento, AdoAsiento
'''       DGAsiento.Visible = False
'''      'Recolectamos los datos iniciales del concepto del comprobante
'''      .Row = 1
'''      .Col = 1
'''       Co.Fecha = Dato_Campo(.Text)
'''      .Col = 2
'''       Co.TP = Dato_Campo(.Text)
'''      .Col = 3
'''       Co.Numero = Val(Replace(Replace(.Text, "-", ""), " ", ""))
'''      .Col = 5
'''       Co.Concepto = TrimStrg(.Text)
'''       Co.CodigoB = TBeneficiario.Codigo
'''       Progreso_Barra.Valor_Maximo = .rows
'''       For i = 1 To .rows - 1
'''          .Row = i
'''          .Col = 1
'''           Mifecha = Dato_Campo(.Text)
'''          .Col = 2
'''           TP = Dato_Campo(.Text)
'''          .Col = 3
'''           NumComp = Val(Replace(Replace(.Text, "-", ""), " ", ""))
'''          .Col = 4
'''           CodigoT = TrimStrg(.Text)
'''           If InStr(Co.Concepto, CodigoT) <= 0 Then Co.Concepto = Co.Concepto & ", " & CodigoT
'''          'Grabamos el Comprobante Contable
'''           If Mifecha <> Co.Fecha Or TP <> Co.TP Or NumComp <> Co.Numero Then
'''              sSQL = "SELECT CODIGO, TC, COUNT(CODIGO) As Cant_Cta, MIN(A_No) As A_No_Ok, SUM(DEBE) As TDEBE " _
'''                   & "FROM Asiento " _
'''                   & "WHERE Item = '" & NumEmpresa & "' " _
'''                   & "AND T_No = " & Trans_No & " " _
'''                   & "AND CodigoU = '" & CodigoUsuario & "' " _
'''                   & "GROUP BY CODIGO, TC " _
'''                   & "HAVING COUNT(CODIGO) > 1 " _
'''                   & "ORDER BY CODIGO, TC "
'''              Select_Adodc AdoAux, sSQL
'''              If AdoAux.Recordset.RecordCount > 0 Then
'''                ' MsgBox "Debe -------> " & Co.Numero & " <=> " & AdoAux.Recordset.RecordCount
'''                 Do While Not AdoAux.Recordset.EOF
'''                    CantFils = AdoAux.Recordset.Fields("Cant_Cta")
'''                    ID_Reg = AdoAux.Recordset.Fields("A_No_Ok")
'''                    Cta = AdoAux.Recordset.Fields("CODIGO")
'''                    Debe = Redondear(AdoAux.Recordset.Fields("TDEBE"), 2)
'''                    sSQL = "UPDATE Asiento " _
'''                         & "SET Debe = " & Debe & " " _
'''                         & "WHERE Item = '" & NumEmpresa & "' " _
'''                         & "AND T_No = " & Trans_No & " " _
'''                         & "AND A_No = " & ID_Reg & " " _
'''                         & "AND CodigoU = '" & CodigoUsuario & "' " _
'''                         & "AND CODIGO = '" & Cta & "' "
'''                    Ejecutar_SQL_SP sSQL
'''
'''                    sSQL = "DELETE " _
'''                         & "FROM Asiento " _
'''                         & "WHERE Item = '" & NumEmpresa & "' " _
'''                         & "AND T_No = " & Trans_No & " " _
'''                         & "AND A_No <> " & ID_Reg & " " _
'''                         & "AND Debe > 0 " _
'''                         & "AND CodigoU = '" & CodigoUsuario & "' " _
'''                         & "AND CODIGO = '" & Cta & "' "
'''                    Ejecutar_SQL_SP sSQL
'''                    AdoAux.Recordset.MoveNext
'''                 Loop
'''              End If
'''              sSQL = "SELECT CODIGO, TC, COUNT(CODIGO) As Cant_Cta, MIN(A_No) As A_No_Ok, SUM(HABER) As THABER " _
'''                   & "FROM Asiento " _
'''                   & "WHERE Item = '" & NumEmpresa & "' " _
'''                   & "AND T_No = " & Trans_No & " " _
'''                   & "AND CodigoU = '" & CodigoUsuario & "' " _
'''                   & "GROUP BY CODIGO, TC " _
'''                   & "HAVING COUNT(CODIGO) > 1 " _
'''                   & "ORDER BY CODIGO, TC "
'''              Select_Adodc AdoAux, sSQL
'''              If AdoAux.Recordset.RecordCount > 0 Then
'''                ' MsgBox "Haber -------> " & Co.Numero & " <=> " & AdoAux.Recordset.RecordCount
'''                 Do While Not AdoAux.Recordset.EOF
'''                    CantFils = AdoAux.Recordset.Fields("Cant_Cta")
'''                    ID_Reg = AdoAux.Recordset.Fields("A_No_Ok")
'''                    Cta = AdoAux.Recordset.Fields("CODIGO")
'''                    Haber = Redondear(AdoAux.Recordset.Fields("THABER"), 2)
'''                    sSQL = "UPDATE Asiento " _
'''                         & "SET Haber = " & Haber & " " _
'''                         & "WHERE Item = '" & NumEmpresa & "' " _
'''                         & "AND T_No = " & Trans_No & " " _
'''                         & "AND A_No = " & ID_Reg & " " _
'''                         & "AND CodigoU = '" & CodigoUsuario & "' " _
'''                         & "AND CODIGO = '" & Cta & "' "
'''                    Ejecutar_SQL_SP sSQL
'''
'''                    sSQL = "DELETE " _
'''                         & "FROM Asiento " _
'''                         & "WHERE Item = '" & NumEmpresa & "' " _
'''                         & "AND T_No = " & Trans_No & " " _
'''                         & "AND A_No <> " & ID_Reg & " " _
'''                         & "AND Haber > 0 " _
'''                         & "AND CodigoU = '" & CodigoUsuario & "' " _
'''                         & "AND CODIGO = '" & Cta & "' "
'''                    Ejecutar_SQL_SP sSQL
'''                    AdoAux.Recordset.MoveNext
'''                 Loop
'''              End If
'''             'MsgBox "Comp No. -> " & Co.Numero
'''              NumComp = Co.Numero
'''              Select Case Co.TP
'''                Case "C/I": Co.TP = "CI"
'''                Case "C/E": Co.TP = "CE"
'''                Case Else: Co.TP = "CD"
'''              End Select
'''
'''             'Grabacion del Comprobante
'''              Co.T = Normal
'''              Co.Efectivo = Total
'''              Co.Monto_Total = Total
'''              Co.T_No = Trans_No
'''              Co.Usuario = CodigoUsuario
'''              Co.Item = NumEmpresa
'''              GrabarComprobante Co
'''
'''              Ln_No = 1
'''              Eliminar_Asientos_SP True
'''              IniciarAsientosDe DGAsiento, AdoAsiento
'''              DetalleComp = Ninguno
'''              Datos_Default_Beneficiario
'''              Co.CodigoB = TBeneficiario.Codigo
'''              SumaDebe = 0
'''              SumaHaber = 0
'''             .Col = 1
'''              Co.Fecha = Dato_Campo(.Text)
'''             .Col = 2
'''              Co.TP = Dato_Campo(.Text)
'''             .Col = 3
'''              Co.Numero = Val(Replace(Replace(.Text, "-", ""), " ", ""))
'''             .Col = 5
'''              Co.Concepto = TrimStrg(.Text)
'''           End If
'''          'Referencia
'''          .Col = 4
'''           Codigo1 = TrimStrg(.Text)
'''          .Col = 5
'''           DetalleComp = TrimStrg(MidStrg(.Text, 1, 60))
'''          'Debe
'''          .Col = 6
'''           Debe = Redondear(Val(Dato_Campo(.Text)), 2)
'''          'Haber
'''          .Col = 7
'''           Haber = Redondear(Val(Dato_Campo(.Text)), 2)
'''          'RUC Proveedor
'''          .Col = 8
'''           Co.RUC_CI = Dato_Campo(.Text)
'''          'Nombre Proveedor
'''          .Col = 9
'''           Co.Beneficiario = UCaseStrg(Dato_Campo(.Text))
'''          'Serie
'''          .Col = 10
'''           SerieFactura = Dato_Campo(.Text)
'''          'Factura
'''          .Col = 11
'''           Factura_No = Val(Dato_Campo(.Text))
'''          'Cuenta Contable
'''          .Col = 12
'''           Cta = Dato_Campo(.Text)
'''          'Codigo_Catalogo
'''          'Cuenta
'''          'SubCta
'''          'Moneda_US
'''          'TipoCta
'''          'TipoPago
'''           Cta1 = Leer_Cta_Catalogo(Cta)
'''
'''          'RUC/Cedula/Codigo Alumno/Consumidor Final
'''           If Len(Co.RUC_CI) > 1 Then
'''              TBeneficiario.CI_RUC = Co.RUC_CI
'''              DigVerif = Digito_Verificador( TBeneficiario.CI_RUC)
'''              Co.CodigoB = Tipo_RUC_CI.Codigo_RUC_CI
'''              sSQL = "SELECT Codigo, Cliente, CI_RUC " _
'''                   & "FROM Clientes " _
'''                   & "WHERE Codigo = '" & Co.CodigoB & "' "
'''              Select_Adodc AdoClientes, sSQL
'''              If AdoClientes.Recordset.RecordCount <= 0 Then
'''                 SetAdoAddNew "Clientes"
'''                 SetAdoFields "T", Normal
'''                 SetAdoFields "TD", Tipo_RUC_CI.Tipo_Beneficiario
'''                 SetAdoFields "Codigo", Co.CodigoB
'''                 SetAdoFields "CI_RUC", Co.RUC_CI
'''                 SetAdoFields "Cliente", Co.Beneficiario
'''                 SetAdoFields "Direccion", "S/D"
'''                 SetAdoFields "DirNumero", "S/N"
'''                 SetAdoFields "Ciudad", "QUITO"
'''                 SetAdoFields "Prov", "17"
'''                 SetAdoFields "Cod_Ejec", Abreviatura_Texto(Co.Beneficiario)
'''                 SetAdoFields "Cta_CxP", Cta
'''                 SetAdoUpdate
'''              End If
'''
'''              sSQL = "SELECT Codigo, TC, ID " _
'''                   & "FROM Catalogo_CxCxP " _
'''                   & "WHERE Item = '" & NumEmpresa & "' " _
'''                   & "AND Periodo = '" & Periodo_Contable & "' " _
'''                   & "AND Codigo = '" & Co.CodigoB & "' " _
'''                   & "AND Cta = '" & Cta & "' "
'''              Select_Adodc AdoClientes, sSQL
'''              If AdoClientes.Recordset.RecordCount <= 0 Then
'''                 SetAdoAddNew "Catalogo_CxCxP"
'''                 SetAdoFields "TC", SubCta
'''                 SetAdoFields "Codigo", Co.CodigoB
'''                 SetAdoFields "Cta", Cta
'''                 SetAdoUpdate
'''              Else
'''                 sSQL = "UPDATE Catalogo_CxCxP " _
'''                      & "SET TC = '" & SubCta & "' " _
'''                      & "WHERE Item = '" & NumEmpresa & "' " _
'''                      & "AND Periodo = '" & Periodo_Contable & "' " _
'''                      & "AND Codigo = '" & Co.CodigoB & "' " _
'''                      & "AND Cta = '" & Cta & "' "
'''                 Ejecutar_SQL_SP sSQL
'''              End If
'''
'''             'Insertamos los submodulos
'''              SetAdoAddNew "Asiento_SC"
'''              SetAdoFields "Codigo", Co.CodigoB
'''              SetAdoFields "Detalle_SubCta", Ninguno
'''              SetAdoFields "FECHA_V", Co.Fecha
'''              SetAdoFields "Beneficiario", Co.Beneficiario
'''              SetAdoFields "Serie", SerieFactura
'''              SetAdoFields "Factura", Factura_No
'''              If Debe > 0 Then
'''                 OpcDH = "1"
'''                 SetAdoFields "Valor", Debe
'''              End If
'''              If Haber > 0 Then
'''                 OpcDH = "2"
'''                 SetAdoFields "Valor", Haber
'''              End If
'''              SetAdoFields "DH", OpcDH
'''              SetAdoFields "TC", SubCta
'''              SetAdoFields "Cta", Cta
'''              SetAdoFields "TM", "1"
'''              SetAdoFields "T_No", Trans_No
'''              SetAdoFields "SC_No", Ln_No
'''              SetAdoUpdate
'''           End If
'''
'''          'Insertamos las transacciones
'''           SumaDebe = SumaDebe + Debe
'''           SumaHaber = SumaHaber + Haber
'''           If Debe > 0 Then InsertarAsientos AdoAsiento, Cta, 0, Debe, 0
'''           If Haber > 0 Then InsertarAsientos AdoAsiento, Cta, 0, 0, Haber
'''           Progreso_Barra.Mensaje_Box = "Importando a Contabilidad el " & Co.TP & ": " & Co.Numero & ": " & i & " de " & Rango.NumFila2
'''           Progreso_Esperar
'''           DoEvents
'''      Next i
'''     sSQL = "SELECT CODIGO, TC, COUNT(CODIGO) As Cant_Cta, MIN(A_No) As A_No_Ok, SUM(DEBE) As TDEBE " _
'''          & "FROM Asiento " _
'''          & "WHERE Item = '" & NumEmpresa & "' " _
'''          & "AND T_No = " & Trans_No & " " _
'''          & "AND CodigoU = '" & CodigoUsuario & "' " _
'''          & "GROUP BY CODIGO, TC " _
'''          & "HAVING COUNT(CODIGO) > 1 " _
'''          & "ORDER BY CODIGO, TC "
'''     Select_Adodc AdoAux, sSQL
'''     If AdoAux.Recordset.RecordCount > 0 Then
'''       ' MsgBox "Debe -------> " & Co.Numero & " <=> " & AdoAux.Recordset.RecordCount
'''        Do While Not AdoAux.Recordset.EOF
'''           CantFils = AdoAux.Recordset.Fields("Cant_Cta")
'''           ID_Reg = AdoAux.Recordset.Fields("A_No_Ok")
'''           Cta = AdoAux.Recordset.Fields("CODIGO")
'''           Debe = Redondear(AdoAux.Recordset.Fields("TDEBE"), 2)
'''           sSQL = "UPDATE Asiento " _
'''                & "SET Debe = " & Debe & " " _
'''                & "WHERE Item = '" & NumEmpresa & "' " _
'''                & "AND T_No = " & Trans_No & " " _
'''                & "AND A_No = " & ID_Reg & " " _
'''                & "AND CodigoU = '" & CodigoUsuario & "' " _
'''                & "AND CODIGO = '" & Cta & "' "
'''           Ejecutar_SQL_SP sSQL
'''
'''           sSQL = "DELETE " _
'''                & "FROM Asiento " _
'''                & "WHERE Item = '" & NumEmpresa & "' " _
'''                & "AND T_No = " & Trans_No & " " _
'''                & "AND A_No <> " & ID_Reg & " " _
'''                & "AND Debe > 0 " _
'''                & "AND CodigoU = '" & CodigoUsuario & "' " _
'''                & "AND CODIGO = '" & Cta & "' "
'''           Ejecutar_SQL_SP sSQL
'''           AdoAux.Recordset.MoveNext
'''        Loop
'''     End If
'''     sSQL = "SELECT CODIGO, TC, COUNT(CODIGO) As Cant_Cta, MIN(A_No) As A_No_Ok, SUM(HABER) As THABER " _
'''          & "FROM Asiento " _
'''          & "WHERE Item = '" & NumEmpresa & "' " _
'''          & "AND T_No = " & Trans_No & " " _
'''          & "AND CodigoU = '" & CodigoUsuario & "' " _
'''          & "GROUP BY CODIGO, TC " _
'''          & "HAVING COUNT(CODIGO) > 1 " _
'''          & "ORDER BY CODIGO, TC "
'''     Select_Adodc AdoAux, sSQL
'''     If AdoAux.Recordset.RecordCount > 0 Then
'''        'MsgBox "Haber -------> " & Co.Numero & " <=> " & AdoAux.Recordset.RecordCount
'''        Do While Not AdoAux.Recordset.EOF
'''           CantFils = AdoAux.Recordset.Fields("Cant_Cta")
'''           ID_Reg = AdoAux.Recordset.Fields("A_No_Ok")
'''           Cta = AdoAux.Recordset.Fields("CODIGO")
'''           Haber = Redondear(AdoAux.Recordset.Fields("THABER"), 2)
'''           sSQL = "UPDATE Asiento " _
'''                & "SET Haber = " & Haber & " " _
'''                & "WHERE Item = '" & NumEmpresa & "' " _
'''                & "AND T_No = " & Trans_No & " " _
'''                & "AND A_No = " & ID_Reg & " " _
'''                & "AND CodigoU = '" & CodigoUsuario & "' " _
'''                & "AND CODIGO = '" & Cta & "' "
'''           Ejecutar_SQL_SP sSQL
'''
'''           sSQL = "DELETE " _
'''                & "FROM Asiento " _
'''                & "WHERE Item = '" & NumEmpresa & "' " _
'''                & "AND T_No = " & Trans_No & " " _
'''                & "AND A_No <> " & ID_Reg & " " _
'''                & "AND Haber > 0 " _
'''                & "AND CodigoU = '" & CodigoUsuario & "' " _
'''                & "AND CODIGO = '" & Cta & "' "
'''           Ejecutar_SQL_SP sSQL
'''           AdoAux.Recordset.MoveNext
'''        Loop
'''     End If
'''     'MsgBox "Comp No. -> " & Co.Numero
'''     NumComp = Co.Numero
'''     Select Case Co.TP
'''       Case "C/I": Co.TP = "CI"
'''       Case "C/E": Co.TP = "CE"
'''       Case Else: Co.TP = "CD"
'''     End Select
'''
'''    'Grabacion del Comprobante
'''     Co.T = Normal
'''     Co.Efectivo = Total
'''     Co.Monto_Total = Total
'''     Co.T_No = Trans_No
'''     Co.Usuario = CodigoUsuario
'''     Co.Item = NumEmpresa
'''     GrabarComprobante Co
'''
'''     Eliminar_Asientos_SP True
'''     IniciarAsientosDe DGAsiento, AdoAsiento
'''     DetalleComp = Ninguno
'''     DGAsiento.Visible = True
'''  End With
'''End Sub
''-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
''-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
''-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=

Public Sub Importar_Actualizacion_Estudiantes()
Dim I As Long
Dim N As Long
Dim Cl As Tipo_Beneficiarios
Dim Tot_Propinas As Currency

    Progreso_Barra.Mensaje_Box = "Actualizando Datos de Estudiantes"
    Progreso_Iniciar
    TextoImprimio = ""
    With AdoExcelAdodc.Recordset
     If .RecordCount > 0 Then
         Progreso_Barra.Valor_Maximo = .RecordCount + 10
        .MoveFirst
         Do While Not .EOF
            For IdField = 0 To .fields.Count - 1
                If IdField = 20 Then Codigo = Dato_Campo(.fields(IdField), True) Else Codigo = Dato_Campo(.fields(IdField), , True)
               'Codigo = Sin_Signos_Especiales(Codigo)
               'MsgBox IdField & " - " & Codigo
                Select Case IdField
                  Case 0: Cl.Grupo_No = Codigo        ' Curso avreviado
                  Case 1: Cl.CI_RUC = Codigo          ' CI_RUC Estudiante
                  Case 2: Cl.Cliente = Codigo         ' Cliente
                  Case 3: Cl.Representante = Codigo   ' Representante
                  Case 4: Cl.RUC_CI_Rep = Codigo      ' Cedula del Representante
                  Case 5: Cl.Direccion = Codigo       ' Descripcion del Curso
                  Case 6: Cl.Email1 = LCase$(Codigo)  ' Email
                  Case 7: Cl.Celular = Codigo         ' Celular
                  Case 8: Cl.Telefono1 = Codigo       ' Telefono
                End Select
            Next IdField
            CodigoCli = Ninguno
            sSQL = "SELECT Codigo " _
                 & "FROM Clientes " _
                 & "WHERE CI_RUC = '" & Cl.CI_RUC & "' "
            Select_Adodc AdoAux, sSQL
            If AdoAux.Recordset.RecordCount > 0 Then CodigoCli = AdoAux.Recordset.fields("Codigo")
           'MsgBox CodigoCli
            If CodigoCli <> Ninguno Then
              'Grupo
               sSQL = "UPDATE Clientes " _
                    & "SET Grupo = '" & Cl.Grupo_No & "' " _
                    & "WHERE Codigo = '" & CodigoCli & "' " _
                    & "AND Grupo <> '" & Cl.Grupo_No & "' "
               Ejecutar_SQL_SP sSQL
            
               sSQL = "UPDATE Clientes_Matriculas " _
                    & "SET Grupo_No = '" & Cl.Grupo_No & "' " _
                    & "WHERE Codigo = '" & CodigoCli & "' " _
                    & "AND Periodo = '" & Periodo_Contable & "' " _
                    & "AND Grupo_No <> '" & Cl.Grupo_No & "' "
               Ejecutar_SQL_SP sSQL
            
              'Cliente / Estudiante
               sSQL = "UPDATE Clientes " _
                    & "SET Cliente = '" & Cl.Cliente & "' " _
                    & "WHERE Codigo = '" & CodigoCli & "' " _
                    & "AND Cliente <> '" & Cl.Cliente & "' "
               Ejecutar_SQL_SP sSQL
            
              'Email Rep
               sSQL = "UPDATE Clientes " _
                    & "SET EmailR = '" & Cl.Email1 & "' " _
                    & "WHERE Codigo = '" & CodigoCli & "' " _
                    & "AND EmailR <> '" & Cl.Email1 & "' "
               Ejecutar_SQL_SP sSQL
            
               sSQL = "UPDATE Clientes_Matriculas " _
                    & "SET Email_R = '" & Cl.Email1 & "' " _
                    & "WHERE Codigo = '" & CodigoCli & "' " _
                    & "AND Periodo = '" & Periodo_Contable & "' " _
                    & "AND Email_R <> '" & Cl.Email1 & "' "
               Ejecutar_SQL_SP sSQL
            
              'Representante
               sSQL = "UPDATE Clientes " _
                    & "SET Representante = '" & Cl.Representante & "' " _
                    & "WHERE Codigo = '" & CodigoCli & "' " _
                    & "AND Representante <> '" & Cl.Representante & "' "
               Ejecutar_SQL_SP sSQL
            
               sSQL = "UPDATE Clientes_Matriculas " _
                    & "SET Representante = '" & Cl.Representante & "' " _
                    & "WHERE Codigo = '" & CodigoCli & "' " _
                    & "AND Periodo = '" & Periodo_Contable & "' " _
                    & "AND Representante <> '" & Cl.Representante & "' "
               Ejecutar_SQL_SP sSQL
            
              'CI_RUC Representante
               sSQL = "UPDATE Clientes " _
                    & "SET CI_RUC_R = '" & Cl.RUC_CI_Rep & "' " _
                    & "WHERE Codigo = '" & CodigoCli & "' " _
                    & "AND CI_RUC_R <> '" & Cl.RUC_CI_Rep & "' "
               Ejecutar_SQL_SP sSQL
            
               sSQL = "UPDATE Clientes_Matriculas " _
                    & "SET Cedula_R = '" & Cl.RUC_CI_Rep & "' " _
                    & "WHERE Codigo = '" & CodigoCli & "' " _
                    & "AND Periodo = '" & Periodo_Contable & "' " _
                    & "AND Cedula_R <> '" & Cl.RUC_CI_Rep & "' "
               Ejecutar_SQL_SP sSQL
            
              'Direccion Representante
               sSQL = "UPDATE Clientes " _
                    & "SET DireccionT = '" & Cl.Direccion & "' " _
                    & "WHERE Codigo = '" & CodigoCli & "' " _
                    & "AND DireccionT <> '" & Cl.Direccion & "' "
               Ejecutar_SQL_SP sSQL
            
               sSQL = "UPDATE Clientes_Matriculas " _
                    & "SET Lugar_Trabajo_R = '" & Cl.Direccion & "' " _
                    & "WHERE Codigo = '" & CodigoCli & "' " _
                    & "AND Periodo = '" & Periodo_Contable & "' " _
                    & "AND Lugar_Trabajo_R <> '" & Cl.Direccion & "' "
               Ejecutar_SQL_SP sSQL
            
              'Telefono Representante
               sSQL = "UPDATE Clientes " _
                    & "SET TelefonoT = SUBSTRING('" & Cl.Telefono1 & "',1,10) " _
                    & "WHERE Codigo = '" & CodigoCli & "' " _
                    & "AND TelefonoT <> '" & Cl.Telefono1 & "' "
               Ejecutar_SQL_SP sSQL
            
               sSQL = "UPDATE Clientes_Matriculas " _
                    & "SET Telefono_RS = SUBSTRING('" & Cl.Telefono1 & "',1,10) " _
                    & "WHERE Codigo = '" & CodigoCli & "' " _
                    & "AND Periodo = '" & Periodo_Contable & "' " _
                    & "AND Telefono_RS <> '" & Cl.Telefono1 & "' "
               Ejecutar_SQL_SP sSQL
            Else
            
               Insertar_Texto_Temporal_SP "Verifique: " & Cl.Cliente & " - " & Cl.RUC_CI_Rep
            End If
            Progreso_Barra.Mensaje_Box = "[" & Progreso_Barra.Incremento & "/" & Progreso_Barra.Valor_Maximo & "] Importando Estudiante " & Cl.Cliente
            Progreso_Esperar
           .MoveNext
         Loop
     End If
    End With
    Progreso_Final
    FInfoError.Show 1
End Sub

