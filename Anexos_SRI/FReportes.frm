VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.5#0"; "comctl32.Ocx"
Begin VB.Form FReportes 
   Caption         =   "Reportes Generales"
   ClientHeight    =   8100
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11190
   Icon            =   "FReportes.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8100
   ScaleWidth      =   11190
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   795
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   11190
      _ExtentX        =   19738
      _ExtentY        =   1402
      ButtonWidth     =   1455
      ButtonHeight    =   1244
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   8
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Salir"
            Object.ToolTipText     =   "Salir del Módulo"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Print_Retenciones"
            Object.ToolTipText     =   "Imprimir Resumen de Retenciones"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "BCompras"
            Object.ToolTipText     =   "Resumen de Compras"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "BVentas"
            Object.ToolTipText     =   "Resumen de Ventas"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "BImportaciones"
            Object.ToolTipText     =   "Resumen de Importaciones"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "BExportaciones"
            Object.ToolTipText     =   "Resumen de Exportaciones"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "BAnulados"
            Object.ToolTipText     =   "Resumen de Transacciones Anuladas"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "MATSAnual"
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.Frame Frame1 
         Height          =   855
         Left            =   6615
         TabIndex        =   8
         Top             =   -105
         Width           =   5055
         Begin VB.ComboBox CMes 
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
            Left            =   3360
            TabIndex        =   11
            Text            =   "Enero"
            Top             =   315
            Width           =   1515
         End
         Begin VB.ComboBox CAño 
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
            ItemData        =   "FReportes.frx":030A
            Left            =   1050
            List            =   "FReportes.frx":030C
            TabIndex        =   9
            Text            =   "2000"
            Top             =   315
            Width           =   1410
         End
         Begin VB.Label Label4 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " &Mes:"
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
            TabIndex        =   12
            Top             =   315
            Width           =   855
         End
         Begin VB.Label Label1 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " &Año:"
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
            Left            =   210
            TabIndex        =   10
            Top             =   315
            Width           =   855
         End
      End
   End
   Begin MSDataGridLib.DataGrid DGAir 
      Bindings        =   "FReportes.frx":030E
      Height          =   2490
      Left            =   105
      TabIndex        =   14
      Top             =   4725
      Width           =   11670
      _ExtentX        =   20585
      _ExtentY        =   4392
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   18
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
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Retenciones de las Compras"
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
   Begin MSDataGridLib.DataGrid DGATS 
      Bindings        =   "FReportes.frx":0323
      Height          =   2385
      Left            =   105
      TabIndex        =   13
      Top             =   2205
      Width           =   11670
      _ExtentX        =   20585
      _ExtentY        =   4207
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   18
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
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Compras Emitidas"
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
   Begin MSDataListLib.DataCombo DCCIRUC 
      Bindings        =   "FReportes.frx":0338
      DataSource      =   "AdoClien"
      Height          =   315
      Left            =   1995
      TabIndex        =   5
      Top             =   1785
      Visible         =   0   'False
      Width           =   9000
      _ExtentX        =   15875
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo DCBeneficiario 
      Bindings        =   "FReportes.frx":034F
      DataSource      =   "AdoClien"
      Height          =   315
      Left            =   1995
      TabIndex        =   3
      Top             =   1365
      Visible         =   0   'False
      Width           =   9000
      _ExtentX        =   15875
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo DCPorCodigo 
      Bindings        =   "FReportes.frx":0366
      DataSource      =   "AdoRetencion"
      Height          =   315
      Left            =   1995
      TabIndex        =   1
      Top             =   945
      Visible         =   0   'False
      Width           =   9000
      _ExtentX        =   15875
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
   End
   Begin VB.CheckBox ChkPorCIRUC 
      Caption         =   "Por CI/RUC"
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
      TabIndex        =   4
      Top             =   1785
      Width           =   1800
   End
   Begin VB.CheckBox ChkPorBenef 
      Caption         =   "Por Beneficiario"
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
      Top             =   1365
      Width           =   1800
   End
   Begin VB.CheckBox ChkPorCodigo 
      Caption         =   "Por Código"
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
      Top             =   945
      Width           =   1800
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&S"
      Height          =   225
      Left            =   11130
      Picture         =   "FReportes.frx":0381
      TabIndex        =   6
      ToolTipText     =   "Salir"
      Top             =   945
      Width           =   225
   End
   Begin MSAdodcLib.Adodc AdoMes 
      Height          =   330
      Left            =   525
      Top             =   4620
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
      Caption         =   "Meses"
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
   Begin MSAdodcLib.Adodc AdoATS 
      Height          =   330
      Left            =   525
      Top             =   3990
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
      Caption         =   "ATS"
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
   Begin MSAdodcLib.Adodc AdoAir 
      Height          =   330
      Left            =   525
      Top             =   4305
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
      Caption         =   "Air"
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
      Left            =   525
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
   Begin MSAdodcLib.Adodc AdoQuery 
      Height          =   330
      Left            =   525
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
      Caption         =   "Query"
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
   Begin MSAdodcLib.Adodc AdoRetencion 
      Height          =   330
      Left            =   525
      Top             =   2985
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
      Caption         =   "Retencion"
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
   Begin MSAdodcLib.Adodc AdoClien 
      Height          =   330
      Left            =   525
      Top             =   4935
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
      Caption         =   "Clien"
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
      Top             =   5250
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
   Begin ComctlLib.ImageList ImageList1 
      Left            =   11130
      Top             =   1365
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   8
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FReportes.frx":0A17
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FReportes.frx":0D31
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FReportes.frx":104B
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FReportes.frx":1365
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FReportes.frx":167F
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FReportes.frx":1999
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FReportes.frx":1CB3
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FReportes.frx":1FCD
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FReportes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public K As Byte
Public Desde, Hasta As Date

Private Sub ChkPorBenef_Click()
  If ChkPorBenef.Value <> 0 Then DCBeneficiario.Visible = True Else DCBeneficiario.Visible = False
End Sub

Private Sub ChkPorCIRUC_Click()
If ChkPorCIRUC.Value <> 0 Then DCCIRUC.Visible = True Else DCCIRUC.Visible = False
End Sub

Private Sub ChkPorCodigo_click()
  If ChkPorCodigo.Value <> 0 Then DCPorCodigo.Visible = True Else DCPorCodigo.Visible = False
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub CMes_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub CMes_LostFocus()
  Fecha_Del_AT CMes, CAño
  Carga_Benef
  Carga_CodigosRet
End Sub

Private Sub DCPorCodigo_LostFocus()
  'Aqui capturo el código de la retención que seleccione el us
   Clave = SinEspaciosIzq(DCPorCodigo)
End Sub

Private Sub DGAir_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  If CtrlDown And KeyCode = vbKeyF1 Then GenerarDataTexto FReportes, AdoAir
  If CtrlDown And KeyCode = vbKeyDelete Then
     Codigo = DGAir.Columns(3)
     Ln_No = DGAir.Columns(0)
    ' Mifecha = BuscarFecha(DGAirCompras.Columns(2))
     If Codigo = "" Then Codigo = Ninguno
     If Codigo1 = "" Then Codigo1 = Ninguno
     Eliminar_Trans_AT NivelNo, Ln_No, True
  End If
  If CtrlDown And KeyCode = vbKeyM Then
     Codigo = DGAir.Columns(3)
     Ln_No = DGAir.Columns(0)
     Cadena = "Desea cambiar el código " & Codigo & vbCrLf _
            & "Por el siguiente código:"
     Codigo1 = InputBox(Cadena, "CAMBIO DE CODIGOS", Codigo)
     If Codigo = "" Then Codigo = Ninguno
     If Codigo1 = "" Then Codigo1 = Ninguno
     If AdoRetencion.Recordset.RecordCount > 0 Then
        AdoRetencion.Recordset.MoveFirst
        AdoRetencion.Recordset.Find ("Codigo = '" & Codigo1 & "' ")
        If Not AdoRetencion.Recordset.EOF Then
          'Actualizo el código anterior por el seleccionado
           sSQL = "UPDATE Trans_Air " _
                & "SET CodRet = '" & Codigo1 & "' " _
                & "WHERE Fecha Between #" & FechaIni & "# AND #" & FechaFin & "# " _
                & "AND Item = '" & NumEmpresa & "' " _
                & "AND Periodo = '" & Periodo_Contable & "' " _
                & "AND CodRet = '" & Codigo & "' " _
                & "AND Tipo_Trans = '" & NivelNo & "' "
           Ejecutar_SQL_SP sSQL
           MsgBox "Proceso Terminado, vuelva a consultar"
        End If
     End If
  End If
  If CtrlDown And KeyCode = vbKeyR Then
     Codigo = DGAir.Columns(3)
     Ln_No = DGAir.Columns(0)
     TipoCta = DGAir.Columns(4)
     Numero = DGAir.Columns(5)
     Cadena = "Desea cambiar el código " & Codigo & vbCrLf _
            & "Del Comprobante: " & TipoCta & ", Numero: " & Numero & vbCrLf _
            & "Por el siguiente código: "
     Codigo1 = InputBox(Cadena, "CAMBIO DE CODIGOS", Codigo)
     If Codigo = "" Then Codigo = Ninguno
     If Codigo1 = "" Then Codigo1 = Ninguno
     If AdoRetencion.Recordset.RecordCount > 0 Then
        AdoRetencion.Recordset.MoveFirst
        AdoRetencion.Recordset.Find ("Codigo = '" & Codigo1 & "' ")
        If Not AdoRetencion.Recordset.EOF Then
          'Actualizo el código anterior por el seleccionado
           sSQL = "UPDATE Trans_Air " _
                & "SET CodRet = '" & Codigo1 & "' " _
                & "WHERE Fecha Between #" & FechaIni & "# AND #" & FechaFin & "# " _
                & "AND Item = '" & NumEmpresa & "' " _
                & "AND Periodo = '" & Periodo_Contable & "' " _
                & "AND CodRet = '" & Codigo & "' " _
                & "AND TP = '" & TipoCta & "' " _
                & "AND Numero = " & Numero & " " _
                & "AND Tipo_Trans = '" & NivelNo & "' "
           Ejecutar_SQL_SP sSQL
           'MsgBox sSQL
           MsgBox "Proceso Terminado, vuelva a consultar"
        End If
     End If
  End If
End Sub

Private Sub DGATS_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  If CtrlDown And KeyCode = vbKeyF1 Then GenerarDataTexto FReportes, AdoATS
  If CtrlDown And KeyCode = vbKeyDelete Then
     Numero = DGATS.Columns(7) 'No. de la factura
     Ln_No = DGATS.Columns(0) 'Linea_SRI
     Eliminar_Trans_AT NivelNo, Ln_No
  End If
End Sub

Private Sub Form_Activate()
Dim AltoTab As Single
Dim AnchoTab As Single
Dim InicioTab As Single
Dim MitadTab As Single
   Carga_Benef
   NivelNo = Ninguno
   Fecha_Del_AT CMes, CAño
   AnchoTab = MDI_X_Max
   AltoTab = MDI_Y_Max
   MitadTab = (MDI_Y_Max / 2) - 800
       
   InicioTab = DGATS.Top
   DGATS.Height = MitadTab
   DGATS.width = AnchoTab - 250
   DGAir.width = DGATS.width
   DGAir.Top = DGATS.Top + DGATS.Height + 50
   DGAir.Height = AltoTab - DGAir.Top
   
   Carga_CodigosRet
   Carga_Benef
   CMes.Clear
   CAño.Clear
   sSQL = Listar_Meses
   Select_Adodc AdoAux, sSQL
   With AdoAux.Recordset
    If .RecordCount > 0 Then
        Do While Not .EOF
           CMes.AddItem .fields("Dia_Mes")
           CMes.Tag = .fields("No_D_M")
          .MoveNext
        Loop
    End If
   End With
   For I = Year(FechaSistema) To 2000 Step -1
       CAño.AddItem Format(I, "0000")
   Next I
   CAño.Text = CAño.List(0)
   CMes.Text = MesesLetras(Month(FechaSistema))
   CAño.SetFocus
   RatonNormal
   sSQL = "SELECT Linea_SRI,Fecha,'Compras' As Tipo,COUNT(Linea_SRI) " _
       & "FROM Trans_Compras " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND Linea_SRI = 0 " _
       & "GROUP BY Linea_SRI, Fecha " _
       & "UNION " _
       & "SELECT Linea_SRI,Fecha,'Ventas' As Tipo,COUNT(Linea_SRI) " _
       & "FROM Trans_Ventas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND Linea_SRI = 0 " _
       & "GROUP BY Linea_SRI, Fecha "
   sSQL = sSQL & "UNION " _
       & "SELECT Linea_SRI,Fecha,'Importaciones' As Tipo,COUNT(Linea_SRI)" _
       & "FROM Trans_Importaciones " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND Linea_SRI = 0 " _
       & "GROUP BY Linea_SRI, Fecha " _
       & "UNION " _
       & "SELECT Linea_SRI,Fecha,'Exportaciones' As Tipo,COUNT(Linea_SRI) " _
       & "FROM Trans_Exportaciones " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND Linea_SRI = 0 " _
       & "GROUP BY Linea_SRI, Fecha " _
       & "ORDER BY Linea_SRI, Fecha "
   Select_Adodc AdoQuery, sSQL
   With AdoQuery.Recordset
    If .RecordCount > 0 Then
       .MoveFirst
        Cadena = "Por:" & vbCrLf & vbCrLf
        Mifecha = .fields("Fecha")
        TipoDoc = .fields("Tipo")
        Do While Not .EOF
           If Month(Mifecha) <> Month(.fields("Fecha")) Or _
              TipoDoc <> .fields("Tipo") Then
              Cadena = Cadena _
                     & MesesLetras(Month(Mifecha)) & "/" _
                     & Year(Mifecha) & " - " & TipoDoc & vbCrLf
              Mifecha = .fields("Fecha")
              TipoDoc = .fields("Tipo")
           End If
          .MoveNext
        Loop
        Cadena = Cadena _
               & MesesLetras(Month(Mifecha)) & "/" _
               & Year(Mifecha) & " - " & TipoDoc & vbCrLf
               
        MsgBox UCaseStrg("Advertencia: " & vbCrLf & vbCrLf _
             & "Debe generar el Anexo Transaccional antes de ingresar" & vbCrLf & vbCrLf _
             & "a este Proceso; tiene que generar " _
             & Cadena & " ")
'        Unload FReportes
'        FAnexoTransaccional.Show
    End If
   End With
End Sub

Private Sub Form_Load()
    ConectarAdodc AdoATS
    ConectarAdodc AdoMes
    ConectarAdodc AdoAux
    ConectarAdodc AdoRetencion
    ConectarAdodc AdoClientes
    ConectarAdodc AdoClien
    ConectarAdodc AdoAir
    ConectarAdodc AdoQuery
End Sub

Public Sub Carga_CodigosRet()
 'Cargo los Códigos de los Conceptos de retención
  sSQL = "SELECT (Codigo & ' - ' & STR(Porcentaje) & ' - ' &  Concepto) As Detalle_Conceptos,* " _
       & "FROM Tipo_Concepto_Retencion " _
       & "WHERE Fecha_Inicio <= #" & FechaMid & "# " _
       & "AND Fecha_Final >= #" & FechaMid & "# " _
       & "ORDER BY Codigo,Concepto,Fecha_Inicio "
  SelectDB_Combo DCPorCodigo, AdoRetencion, sSQL, "Detalle_Conceptos"
End Sub

Public Sub Carga_Benef()
   'Carga en el Data Combo los Clientes con su RUC
    Fecha_Del_AT CMes, CAño
    sSQL = "SELECT C.Codigo,C.Cliente,C.CI_RUC " _
         & "FROM Clientes As C,Trans_Compras As TC " _
         & "WHERE TC.Fecha Between # " & FechaIni & "# AND #" & FechaFin & "# " _
         & "AND C.Codigo = TC.IdProv " _
         & "UNION " _
         & "SELECT C.Codigo,C.Cliente,C.CI_RUC " _
         & "FROM Clientes As C,Trans_Ventas As TV " _
         & "WHERE TV.Fecha Between # " & FechaIni & "# AND #" & FechaFin & "# " _
         & "AND C.Codigo = TV.IdProv " _
         & "UNION " _
         & "SELECT C.Codigo,C.Cliente,C.CI_RUC " _
         & "FROM Clientes As C, Trans_Exportaciones As TE " _
         & "WHERE TE.Fecha Between # " & FechaIni & "# AND #" & FechaFin & "# " _
         & "AND C.Codigo = TE.IdFiscalProv " _
         & "UNION " _
         & "SELECT C.Codigo,C.Cliente,C.CI_RUC " _
         & "FROM Clientes As C, Trans_Importaciones As TI " _
         & "WHERE TI.Fecha Between # " & FechaIni & "# AND #" & FechaFin & "# " _
         & "AND C.Codigo = TI.IdFiscalProv " _
         & "GROUP BY C.Codigo,C.Cliente,C.CI_RUC " _
         & "ORDER BY C.Cliente "
    SelectDB_Combo DCBeneficiario, AdoClien, sSQL, "Cliente"
    SelectDB_Combo DCCIRUC, AdoClien, sSQL, "CI_RUC"
    ' MsgBox sSQL & vbCrLf & AdoClien.Recordset.RecordCount
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
  Carga_Benef
End Sub

Public Sub Eliminar_Trans_AT(Tipo_AT As String, Del_Ln_No As Integer, Optional Solo_Air As Boolean)
   Fecha_Del_AT CMes, CAño
   Select Case Tipo_AT
     Case "C": Titulo = "ELIMINACION DE COMPRAS"
     Case "V": Titulo = "ELIMINACION DE VENTAS"
     Case "E": Titulo = "ELIMINACION DE EXPORTACIONES"
     Case "I": Titulo = "ELIMINACION DE IMPORTACIONES"
     Case "A": Titulo = "ELIMINACION DE ANULADOS"
     Case Else: Titulo = "ELIMINACION DE DATOS"
   End Select
   Mensajes = "Usted al eliminar esta transacción tendra que " & vbCrLf _
            & "volverla a ingresar, no podrá recuperar este dato" & vbCrLf _
            & "Seguro de Eliminar?"
   If BoxMensaje = vbYes Then
     'Eliminar la transaccion del mes
      sSQL = "DELETE * "
      Select Case Tipo_AT
        Case "C": sSQL = sSQL & "FROM Trans_Compras " _
                       & "WHERE Fecha Between #" & FechaIni & "# AND #" & FechaFin & "# "
        Case "V": sSQL = sSQL & "FROM Trans_Ventas " _
                       & "WHERE Fecha Between #" & FechaIni & "# AND #" & FechaFin & "# "
        Case "E": sSQL = sSQL & "FROM Trans_Exportaciones " _
                       & "WHERE Fecha Between #" & FechaIni & "# AND #" & FechaFin & "# "
        Case "I": sSQL = sSQL & "FROM Trans_Importaciones " _
                       & "WHERE Fecha Between #" & FechaIni & "# AND #" & FechaFin & "# "
        Case "A": sSQL = sSQL & "FROM Trans_Anulados " _
                       & "WHERE FechaAnulacion Between #" & FechaIni & "# AND #" & FechaFin & "# "
      End Select
      sSQL = sSQL & "AND Linea_SRI = " & Del_Ln_No & " "
      If ConSucursal = False Then sSQL = sSQL & "AND Item = '" & NumEmpresa & "' "
      sSQL = sSQL & "AND Periodo = '" & Periodo_Contable & "' "
      If Not Solo_Air Then Ejecutar_SQL_SP sSQL
     'Elimino el Air de la Compra
      sSQL = "DELETE * " _
           & "FROM Trans_Air " _
           & "WHERE Linea_SRI = " & Del_Ln_No & " " _
           & "AND Fecha Between #" & FechaIni & "# AND #" & FechaFin & "# "
      If ConSucursal = False Then sSQL = sSQL & "AND Item = '" & NumEmpresa & "' "
      sSQL = sSQL & "AND Periodo = '" & Periodo_Contable & "' " _
           & "AND Tipo_Trans = '" & Tipo_AT & "' "
      Ejecutar_SQL_SP sSQL
   End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
Dim No_Mes As Byte
    Carga_Benef
    Fecha_Del_AT CMes, CAño
   'Consultamos el codigo que pertenece al clietne seleccionado
    NombreCliente = TrimStrg(DCBeneficiario.Text)
    CodigoCli = Ninguno
    With AdoClien.Recordset
     If .RecordCount > 0 Then
        .MoveFirst
        .Find ("Cliente = '" & NombreCliente & "' ")
         If Not .EOF Then CodigoCli = .fields("Codigo")
     End If
    End With
    CICliente = TrimStrg(DCCIRUC.Text)
   '
    Select Case Button.key
      Case "Salir"
           Unload Me
      Case "Print_Retenciones"
      
      Case "BCompras"
            NivelNo = "C"
            sSQL = "SELECT TC.Linea_SRI,C.Cliente,C.TD,C.CI_RUC,TC.TipoComprobante,TC.CodSustento,TC.PagoLocExt,TC.FormaPago,TC.PaisEfecPago,TC.AplicConvDobTrib,TC.PagExtSujRetNorLeg,TC.TP,TC.Numero,TC.Establecimiento,TC.PuntoEmision,TC.Secuencial,TC.Autorizacion,TC.FechaEmision,TC.FechaRegistro,TC.FechaCaducidad,TC.BaseNoObjIVA,TC.BaseImponible,TC.BaseImpGrav,TC.PorcentajeIva,TC.MontoIva,TC.BaseImpIce,TC.PorcentajeIce,TC.MontoIce,TC.MontoIvaBienes,TC.PorRetBienes,TC.ValorRetBienes,TC.MontoIvaServicios,TC.PorRetServicios,TC.ValorRetServicios,TC.ContratoPartidoPolitico,TC.MontoTituloOneroso,TC.MontoTituloGratuito,TC.DevIva " _
                 & "FROM Trans_Compras As TC, Clientes As C " _
                 & "WHERE TC.Fecha Between #" & FechaIni & "# AND #" & FechaFin & "#  "
            If ConSucursal = False Then sSQL = sSQL & "AND TC.Item = '" & NumEmpresa & "' "
            sSQL = sSQL & "AND TC.Periodo = '" & Periodo_Contable & "' "
            If ChkPorBenef.Value <> 0 Then sSQL = sSQL & "AND C.Codigo = '" & CodigoCli & "' "
            If ChkPorCIRUC.Value <> 0 Then sSQL = sSQL & "AND C.CI_RUC = '" & CICliente & "' "
            sSQL = sSQL & "AND TC.IdProv = C.Codigo " _
                 & "ORDER BY TC.Linea_SRI, C.Cliente, C.CI_RUC, C.TD "
            Select_Adodc_Grid DGATS, AdoATS, sSQL
            
            sSQL = "SELECT TA.Linea_SRI,C.Cliente,TA.Fecha,TA.CodRet,TA.TP,TA.Numero,TA.BaseImp,TA.Porcentaje,TA.ValRet,TA.EstabRetencion,TA.PtoEmiRetencion,TA.SecRetencion,TA.EstabFactura,TA.PuntoEmiFactura,TA.AutRetencion,TA.Factura_No " _
                 & "FROM Trans_Air As TA,Clientes As C " _
                 & "WHERE TA.Fecha Between #" & FechaIni & "# AND #" & FechaFin & "# "
            If ConSucursal = False Then sSQL = sSQL & "AND TA.Item = '" & NumEmpresa & "' "
            sSQL = sSQL & "AND TA.Periodo = '" & Periodo_Contable & "' " _
                 & "AND TA.Tipo_Trans = 'C' "
            If ChkPorCodigo.Value <> 0 Then sSQL = sSQL & "AND TA.CodRet = '" & Clave & "' "
            If ChkPorBenef.Value <> 0 Then sSQL = sSQL & "AND C.Codigo = '" & CodigoCli & "' "
            If ChkPorCIRUC.Value <> 0 Then sSQL = sSQL & "AND C.CI_RUC = '" & CICliente & "' "
            sSQL = sSQL & "AND TA.IdProv = C.Codigo " _
                 & "ORDER BY TA.Linea_SRI, C.Cliente, TA.Fecha "
            Select_Adodc_Grid DGAir, AdoAir, sSQL
            DGATS.Caption = "COMPRAS EMITIDAS"
            DGAir.Caption = "RETENCIONES REALIZADAS"
      Case "BVentas"
            NivelNo = "V"
            sSQL = "SELECT TV.Linea_SRI,C.Cliente,TV.TipoComprobante,C.CI_RUC,C.TD,TV.TP,TV.Numero,TV.Establecimiento,TV.PuntoEmision,TV.NumeroComprobantes,TV.Secuencial,TV.FechaEmision,TV.FechaRegistro,TV.BaseImponible,TV.BaseImpGrav,TV.PorcentajeIva,TV.MontoIva,TV.BaseImpIce,TV.PorcentajeIce,TV.MontoIce,TV.MontoIvaBienes,TV.PorRetBienes,TV.ValorRetBienes,TV.MontoIvaServicios,TV.PorRetServicios,TV.ValorRetServicios,TV.RetPresuntiva " _
                 & "FROM Trans_Ventas As TV, Clientes As C " _
                 & "WHERE TV.Fecha Between #" & FechaIni & "# AND #" & FechaFin & "#  "
            If ConSucursal = False Then sSQL = sSQL & "AND TV.Item = '" & NumEmpresa & "' "
            sSQL = sSQL & "AND TV.Periodo = '" & Periodo_Contable & "' "
            If ChkPorBenef.Value <> 0 Then sSQL = sSQL & "AND C.Codigo = '" & CodigoCli & "' "
            If ChkPorCIRUC.Value <> 0 Then sSQL = sSQL & "AND C.CI_RUC = '" & CICliente & "' "
            sSQL = sSQL & "AND TV.IdProv = C.Codigo  " _
                 & "ORDER BY TV.Linea_SRI, C.Cliente, C.CI_RUC, C.TD "
            Select_Adodc_Grid DGATS, AdoATS, sSQL, "Datos"
            
            sSQL = "SELECT TA.Linea_SRI,C.Cliente,TA.Fecha,TA.CodRet,TA.TP,TA.Numero,TA.BaseImp,TA.Porcentaje,TA.ValRet,TA.EstabRetencion,TA.PtoEmiRetencion,TA.SecRetencion,TA.EstabFactura,TA.PuntoEmiFactura,TA.AutRetencion,TA.Factura_No " _
                 & "FROM Trans_Air As TA,Clientes As C " _
                 & "WHERE TA.Fecha Between #" & FechaIni & "# AND #" & FechaFin & "# "
            If ConSucursal = False Then sSQL = sSQL & "AND TA.Item = '" & NumEmpresa & "' "
            sSQL = sSQL & "AND TA.Periodo = '" & Periodo_Contable & "' " _
                 & "AND TA.Tipo_Trans = 'V' "
            If ChkPorCodigo.Value <> 0 Then sSQL = sSQL & "AND TA.CodRet = '" & Clave & "' "
            If ChkPorBenef.Value <> 0 Then sSQL = sSQL & "AND C.Codigo = '" & CodigoCli & "' "
            If ChkPorCIRUC.Value <> 0 Then sSQL = sSQL & "AND C.CI_RUC = '" & CICliente & "' "
            sSQL = sSQL & "AND TA.IdProv = C.Codigo " _
                 & "ORDER BY TA.Linea_SRI, C.Cliente,TA.Fecha "
            Select_Adodc_Grid DGAir, AdoAir, sSQL
            DGATS.Caption = "VENTAS EMITIDAS"
            DGAir.Caption = "RETENCIONES REALIZADAS"
      Case "BImportaciones"
            NivelNo = "I"
            sSQL = "SELECT TI.Linea_SRI,C.Cliente,C.CI_RUC,C.TD,TI.TP,TI.Numero,TI.FechaLiquidacion,TI.DistAduanero,TI.Anio,TI.Regimen,TI.Correlativo,TI.Verificador,TI.ValorCIF,TI.BaseImponible,TI.BaseImpGrav,TI.PorcentajeIva,TI.MontoIva,TI.BaseImpIce,TI.PorcentajeIce,TI.MontoIce " _
                 & "FROM Trans_Importaciones As TI, Clientes As C " _
                 & "WHERE TI.Fecha Between #" & FechaIni & "# AND #" & FechaFin & "#  "
            If ConSucursal = False Then sSQL = sSQL & "AND TI.Item = '" & NumEmpresa & "' "
            sSQL = sSQL & "AND TI.Periodo = '" & Periodo_Contable & "' "
            If ChkPorBenef.Value <> 0 Then sSQL = sSQL & "AND C.Codigo = '" & CodigoCli & "' "
            If ChkPorCIRUC.Value <> 0 Then sSQL = sSQL & "AND C.CI_RUC = '" & CICliente & "' "
            sSQL = sSQL & "AND TI.IdFiscalProv = C.Codigo  " _
                 & "ORDER BY TI.Linea_SRI, C.Cliente, C.CI_RUC, C.TD "
            Select_Adodc_Grid DGATS, AdoATS, sSQL
            
            sSQL = "SELECT TA.Linea_SRI,C.Cliente,TA.Fecha,TA.CodRet,TA.TP,TA.Numero,TA.BaseImp,TA.Porcentaje,TA.ValRet,TA.EstabRetencion,TA.PtoEmiRetencion,TA.SecRetencion,TA.EstabFactura,TA.PuntoEmiFactura,TA.AutRetencion,TA.Factura_No " _
                 & "FROM Trans_Air As TA,Clientes As C " _
                 & "WHERE TA.Fecha Between #" & FechaIni & "# AND #" & FechaFin & "# "
            If ConSucursal = False Then sSQL = sSQL & "AND TA.Item = '" & NumEmpresa & "' "
            sSQL = sSQL & "AND TA.Periodo = '" & Periodo_Contable & "' " _
                 & "AND TA.Tipo_Trans = 'I' "
            If ChkPorCodigo.Value <> 0 Then sSQL = sSQL & "AND TA.CodRet = '" & Clave & "' "
            If ChkPorBenef.Value <> 0 Then sSQL = sSQL & "AND C.Codigo = '" & CodigoCli & "' "
            If ChkPorCIRUC.Value <> 0 Then sSQL = sSQL & "AND C.CI_RUC = '" & CICliente & "' "
            sSQL = sSQL & "AND C.Codigo = TA.IdProv    " _
                 & "ORDER BY TA.Linea_SRI,C.Cliente, TA.Fecha "
            Select_Adodc_Grid DGAir, AdoAir, sSQL
            DGATS.Caption = "IMPORTACIONES EMITIDAS"
            DGAir.Caption = "RETENCIONES REALIZADAS"
      Case "BExportaciones"
            NivelNo = "E"
            sSQL = "SELECT TE.Linea_SRI,C.Cliente,C.CI_RUC,C.TD,TE.TP,TE.Numero,TE.FechaEmbarque,TE.DistAduanero,TE.Anio,TE.Regimen,TE.Correlativo,TE.Verificador,TE.NumeroDctoTransporte,TE.ValorFOB,TE.ValorFOBComprobante,TE.Establecimiento,TE.PuntoEmision,TE.Secuencial,TE.FechaRegistro,TE.Autorizacion,TE.FechaEmision " _
                 & "FROM Trans_Exportaciones As TE, Clientes As C " _
                 & "WHERE TE.Fecha Between #" & FechaIni & "# AND #" & FechaFin & "# "
            If ConSucursal = False Then sSQL = sSQL & "AND TE.Item = '" & NumEmpresa & "' "
            sSQL = sSQL & "AND TE.Periodo = '" & Periodo_Contable & "' "
            If ChkPorBenef.Value <> 0 Then sSQL = sSQL & "AND C.Codigo = '" & CodigoCli & "' "
            If ChkPorCIRUC.Value <> 0 Then sSQL = sSQL & "AND C.CI_RUC = '" & CICliente & "' "
            sSQL = sSQL & "AND TE.IdFiscalProv = C.Codigo " _
                 & "ORDER BY TE.Linea_SRI, C.Cliente, C.CI_RUC, C.TD "
            Select_Adodc_Grid DGATS, AdoATS, sSQL
      Case "BAnuñados"
            NivelNo = "A"
            sSQL = "SELECT Linea_SRI,FechaAnulacion,TP,Numero,Establecimiento,PuntoEmision,Secuencial1,Secuencial2,Autorizacion " _
                 & "FROM Trans_Anulados " _
                 & "WHERE FechaAnulacion Between #" & FechaIni & "# AND #" & FechaFin & "#  "
            If ConSucursal = False Then sSQL = sSQL & "AND Item = '" & NumEmpresa & "' "
            sSQL = sSQL & "AND Periodo = '" & Periodo_Contable & "' " _
                 & "ORDER BY Linea_SRI "
            Select_Adodc_Grid DGATS, AdoATS, sSQL
      Case "MATSAnual"
            NivelNo = "C"
            FechaIni = BuscarFecha("01/01/" & CAño)
            FechaFin = BuscarFecha("31/12/" & CAño)

            sSQL = "UPDATE Trans_Compras " _
                 & "SET ValorRetFuente = (SELECT ROUND(SUM(TA.ValRet),2,0) " _
                 & "                     FROM Trans_Air As TA " _
                 & "                     WHERE TA.ValRet <> 0 " _
                 & "                     AND Trans_Compras.Item = TA.Item " _
                 & "                     AND Trans_Compras.Periodo = TA.Periodo " _
                 & "                     AND Trans_Compras.TP = TA.TP " _
                 & "                     AND Trans_Compras.Numero = TA.Numero) " _
                 & "WHERE Item = '" & NumEmpresa & "' " _
                 & "AND Periodo = '" & Periodo_Contable & "' " _
                 & "AND Fecha Between #" & FechaIni & "# AND #" & FechaFin & "#  "
            Ejecutar_SQL_SP sSQL
            
            sSQL = "UPDATE Trans_Compras " _
                 & "SET ValorRetFuente = 0 " _
                 & "WHERE Periodo = '" & Periodo_Contable & "' " _
                 & "AND Item = '" & NumEmpresa & "' " _
                 & "AND ValorRetFuente IS NULL "
            Ejecutar_SQL_SP sSQL
            
            sSQL = "SELECT TC.FechaEmision As [Fecha de emisión], (TC.Establecimiento + TC.PuntoEmision + RIGHT('0000000' + LTRIM(RTRIM(TC.Secuencial)), 7)) As [Número secuencial del comprobante de venta], " _
                 & "TC.Autorizacion As [Número de autorización del comprobante de venta], C.CI_RUC As [Ruc proveedor], C.Cliente As [Razón social proveedor], " _
                 & "Co.Concepto As [Detalle del Costo o Gasto], TC.BaseImpGrav As [Base imponible 12%], (TC.BaseNoObjIVA+TC.BaseImponible) As [Base imponible 0%], " _
                 & "TC.MontoIva As [iva], (TC.BaseImpGrav+TC.BaseNoObjIVA+TC.BaseImponible+TC.MontoIva) As Total, TC.FechaRegistro As [Fecha Comprobante de retención], " _
                 & "TC.FechaRegistro As [Fecha de emisión del comprobante de retención], TC.ValorRetFuente As [Valor de Impuesto a la Renta retenido], " _
                 & "(TC.ValorRetBienes + TC.ValorRetServicios) As [Valor de Impuesto al Valor Agregado retenido], TC.FormaPago As [Forma de pago] " _
                 & "FROM Trans_Compras As TC, Clientes As C, Comprobantes As Co " _
                 & "WHERE TC.Fecha Between #" & FechaIni & "# AND #" & FechaFin & "#  " _
                 & "AND TC.Periodo = '" & Periodo_Contable & "' "
            If ConSucursal = False Then sSQL = sSQL & "AND TC.Item = '" & NumEmpresa & "' "
            sSQL = sSQL _
                 & "AND TC.IdProv = C.Codigo " _
                 & "AND TC.Item = Co.Item " _
                 & "AND TC.Periodo = Co.Periodo " _
                 & "AND TC.TP = Co.TP " _
                 & "AND TC.Numero = Co.Numero " _
                 & "ORDER BY TC.Fecha, TC.TP, TC.Numero, C.TD "
            Select_Adodc_Grid DGATS, AdoATS, sSQL
            DGATS.Caption = "COMPRAS EMITIDAS"
    End Select
End Sub
