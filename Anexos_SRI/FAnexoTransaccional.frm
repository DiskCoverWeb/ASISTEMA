VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FAnexoTransaccional 
   Caption         =   "Talón Resumen"
   ClientHeight    =   10470
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16515
   Icon            =   "FAnexoTransaccional.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10470
   ScaleWidth      =   16515
   WindowState     =   2  'Maximized
   Begin MSDataGridLib.DataGrid DGRolPagos 
      Bindings        =   "FAnexoTransaccional.frx":0946
      Height          =   2430
      Left            =   105
      TabIndex        =   7
      Top             =   1680
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   4286
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
      Caption         =   "&Formularios 107"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   1680
      Picture         =   "FAnexoTransaccional.frx":0960
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   945
      Width           =   1485
   End
   Begin VB.TextBox TxtCodigo 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1410
      TabIndex        =   4
      Text            =   "Codigo"
      Top             =   6930
      Width           =   1170
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   20265
      Top             =   1260
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FAnexoTransaccional.frx":122A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FAnexoTransaccional.frx":1B04
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FAnexoTransaccional.frx":24B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FAnexoTransaccional.frx":27CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FAnexoTransaccional.frx":2AE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FAnexoTransaccional.frx":2E00
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FAnexoTransaccional.frx":311A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FAnexoTransaccional.frx":3434
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FAnexoTransaccional.frx":3ADA
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FAnexoTransaccional.frx":4180
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&S"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3255
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1050
      Width           =   330
   End
   Begin MSAdodcLib.Adodc AdoMes 
      Height          =   330
      Left            =   2895
      Top             =   1920
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
   Begin MSAdodcLib.Adodc AdoCompras 
      Height          =   330
      Left            =   480
      Top             =   1920
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
      Caption         =   "Compras"
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
   Begin MSAdodcLib.Adodc AdoVentas 
      Height          =   330
      Left            =   480
      Top             =   2235
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
      Caption         =   "Ventas"
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
      Left            =   480
      Top             =   3495
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
   Begin MSAdodcLib.Adodc AdoImportaciones 
      Height          =   330
      Left            =   480
      Top             =   2550
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
      Caption         =   "Importaciones"
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
   Begin MSAdodcLib.Adodc AdoExportaciones 
      Height          =   330
      Left            =   480
      Top             =   2865
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
      Caption         =   "Exportaciones"
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
      Left            =   2895
      Top             =   2865
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
   Begin MSAdodcLib.Adodc AdoAux 
      Height          =   330
      Left            =   2895
      Top             =   2235
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
      Caption         =   "AdoAux"
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
   Begin MSAdodcLib.Adodc AdoSaldos 
      Height          =   330
      Left            =   2895
      Top             =   2550
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
      Caption         =   "AdoSaldos"
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
   Begin MSAdodcLib.Adodc AdoAnulados 
      Height          =   330
      Left            =   480
      Top             =   3180
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
      Caption         =   "AdoAnulados"
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
   Begin MSAdodcLib.Adodc AdoCiudades 
      Height          =   330
      Left            =   2895
      Top             =   3180
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
      Caption         =   "Ciudades"
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
   Begin MSAdodcLib.Adodc AdoCodRet 
      Height          =   330
      Left            =   2895
      Top             =   3495
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
      Caption         =   "CodRet"
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
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   900
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   16515
      _ExtentX        =   29131
      _ExtentY        =   1588
      ButtonWidth     =   1799
      ButtonHeight    =   1429
      Appearance      =   1
      ImageList       =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Años"
            Key             =   "Anos"
            ImageIndex      =   2
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "A2000"
                  Text            =   "2000"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Meses"
            Key             =   "Meses"
            ImageIndex      =   3
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   14
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Enero"
                  Text            =   "Enero"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Febrero"
                  Text            =   "Febrero"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Marzo"
                  Text            =   "Marzo"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Abril"
                  Text            =   "Abril"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Mayo"
                  Text            =   "Mayo"
               EndProperty
               BeginProperty ButtonMenu6 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Junio"
                  Text            =   "Junio"
               EndProperty
               BeginProperty ButtonMenu7 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Julio"
                  Text            =   "Julio"
               EndProperty
               BeginProperty ButtonMenu8 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Agosto"
                  Text            =   "Agosto"
               EndProperty
               BeginProperty ButtonMenu9 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Septiembre"
                  Text            =   "Septiembre"
               EndProperty
               BeginProperty ButtonMenu10 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Octubre"
                  Text            =   "Octubre"
               EndProperty
               BeginProperty ButtonMenu11 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Noviembre"
                  Text            =   "Noviembre"
               EndProperty
               BeginProperty ButtonMenu12 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Diciembre"
                  Text            =   "Diciembre"
               EndProperty
               BeginProperty ButtonMenu13 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "1er_Semestre"
                  Text            =   "1er_Semestre"
               EndProperty
               BeginProperty ButtonMenu14 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "2do_Semestre"
                  Text            =   "2do_Semestre"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   4
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "ATS"
            Key             =   "ATS"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "ATS Financ."
            Key             =   "ATSFinanc"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "REOC"
            Key             =   "REOC"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "RDEP"
            Key             =   "RDEP"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Salir"
            Key             =   "Salir"
            ImageIndex      =   10
         EndProperty
      EndProperty
      MousePointer    =   1
      Begin VB.Frame Frame1 
         Height          =   855
         Left            =   8820
         TabIndex        =   2
         Top             =   0
         Width           =   2955
         Begin VB.CheckBox CheqATSConElect 
            Caption         =   "Generar ATS con Comprobante Electronicos"
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
            Left            =   210
            TabIndex        =   3
            Top             =   210
            Width           =   2640
         End
      End
   End
   Begin MSAdodcLib.Adodc AdoPuntosVentas 
      Height          =   330
      Left            =   420
      Top             =   3885
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
      Caption         =   "PuntosVentas"
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
   Begin MSAdodcLib.Adodc AdoRolPagos 
      Height          =   330
      Left            =   2625
      Top             =   6930
      Width           =   3690
      _ExtentX        =   6509
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
      Caption         =   "RolPagos"
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
   Begin MSMask.MaskEdBox MBFechaE 
      Height          =   330
      Left            =   105
      TabIndex        =   8
      ToolTipText     =   "Formato de Fecha: DD/MM/AA"
      Top             =   1260
      Width           =   1485
      _ExtentX        =   2619
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
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &Fecha Entrega"
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
      Top             =   945
      Width           =   1485
   End
   Begin VB.Label LblA4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   105
      TabIndex        =   5
      Top             =   6930
      Width           =   1275
   End
End
Attribute VB_Name = "FAnexoTransaccional"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Totales_SRI(15) As Currency
Dim cod, Ct, TipoComp As Byte
Dim Valor, SecRet, Acum As Integer
Dim BaseCero, BaseDoce, ValorIva As Double
Dim Porcentaje As Single
Dim FechaRet, Estab, PtoRet, AutRet As String
Dim Archivo_XML As String
Dim Anio_Anexo As String
Dim Mes_Anexo As String
Dim sItem As String
Dim tItem As String

Dim NumTrans As Long
Dim Segunda_Pag As Boolean
Dim EsSemestral As Boolean

Private Sub CmdSalir_Click()
  Unload Me
End Sub

Private Sub Command2_Click()
  FechaValida MBFechaE
  Imprimir_Formulario_107 Format(MBFechaE, "yyyymmdd")
End Sub

Private Sub DGRolPagos_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  If Periodo_Contable = "" Then Periodo_Contable = Ninguno
  If CtrlDown And KeyCode = vbKeyF1 Then
     DGRolPagos.Visible = False
     GenerarDataTexto FAnexoTransaccional, AdoRolPagos
     DGRolPagos.Visible = True
  End If
  If CtrlDown And KeyCode = vbKeyP Then
     MensajeEncabData = "RESUMEN DE ANEXOS TRANSACCIONALES"
     SQLMsg1 = "IMPUESTO A LA RENTA BAJO RELACIÓN DE DEPENDENCIA"
     SQLMsg2 = "Enero a Diciembre de " & Anio
     Cuadricula = True
     DGRolPagos.Visible = False
     ImprimirAdo AdoRolPagos, True, 2, 6, True
     DGRolPagos.Visible = True
  End If
  If CtrlDown And KeyCode = vbKeyF Then
     DGRolPagos.Visible = False
     FechaValida MBFechaE
     Imprimir_Formulario_107 Format(MBFechaE, "yyyymmdd"), True
     DGRolPagos.Visible = True
  End If
End Sub

Private Sub Form_Activate()
Dim ExisteAnioActual As Boolean
  DGRolPagos.width = MDI_X_Max - 100
  DGRolPagos.Height = MDI_Y_Max - 1800
  AdoRolPagos.Top = DGRolPagos.Height - 500
  
  LblA4.Top = DGRolPagos.Height - 500
  TxtCodigo.Top = LblA4.Top
  TxtCodigo = ""
  LblA4.Caption = ""
  
 'Empresas si es con sucursales
  tItem = ""
  sSQL = "SELECT Sucursal " _
       & "FROM Acceso_Sucursales " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND No_ATS = 0 " _
       & "ORDER BY Sucursal "
  Select_Adodc AdoAux, sSQL
  With AdoAux.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          tItem = tItem & "'" & .fields("Sucursal") & "',"
         .MoveNext
       Loop
   End If
  End With
  If Len(tItem) > 1 Then sItem = MidStrg(tItem, 1, Len(tItem) - 1) Else sItem = "'" & NumEmpresa & "','XXX'"
  
  ExisteAnioActual = False
  sSQL = "SELECT YEAR(Fecha) As Anio " _
       & "FROM Trans_Compras " _
       & "WHERE Periodo = '" & Periodo_Contable & "' "
       If ConSucursal Then
          If Len(No_ATS) > 3 Then sSQL = sSQL & "AND Item NOT IN (" & No_ATS & ") "
       Else
          sSQL = sSQL & "AND Item = '" & NumEmpresa & "' "
       End If
  sSQL = sSQL _
       & "UNION " _
       & "SELECT YEAR(Fecha) As Anio " _
       & "FROM Trans_Ventas " _
       & "WHERE Periodo = '" & Periodo_Contable & "' "
       If ConSucursal Then
          If Len(No_ATS) > 3 Then sSQL = sSQL & "AND Item NOT IN (" & No_ATS & ") "
       Else
          sSQL = sSQL & "AND Item = '" & NumEmpresa & "' "
       End If
  sSQL = sSQL _
       & "GROUP BY YEAR(Fecha) " _
       & "ORDER BY YEAR(Fecha) DESC "
  Select_Adodc AdoAux, sSQL, , , "Periodos_ATS"
  With AdoAux.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          Toolbar1.buttons.Item(1).ButtonMenus.Add , "A" & .fields("Anio"), .fields("Anio")
          If Year(FechaSistema) = .fields("Anio") Then ExisteAnioActual = True
         .MoveNext
       Loop
   End If
  End With
 'MsgBox Toolbar1.Buttons.Item(1).ButtonMenus.Count
  If Not ExisteAnioActual And Periodo_Contable = Ninguno Then
     Toolbar1.buttons.Item(1).ButtonMenus.Add , "A" & Year(FechaSistema), Year(FechaSistema)
  End If
  Toolbar1.buttons.Item(1).ButtonMenus.Remove ("A2000")
  
 'verificamos si la carpeta AT y la subcarpeta de la empresa existe sino la creamos
 
  Anio = Year(FechaSistema)
  Mes = MesesLetras(Month(FechaSistema))
  Anio_Anexo = Year(FechaSistema)
  Mes_Anexo = Month(FechaSistema)
  Toolbar1.buttons(2).Caption = Anio
  Toolbar1.buttons(3).Caption = Mes
  Toolbar1.Refresh

  Codigo = RutaSysBases & "\AT"
  Cadena = Dir(Codigo, vbDirectory)
  If Cadena = "" Then MkDir (Codigo)
  
  Codigo = RutaSysBases & "\PRINTER"
  Cadena = Dir(Codigo, vbDirectory)
  If Cadena = "" Then MkDir (Codigo)
  
  Codigo = RutaSysBases & "\AT\AT" & NumEmpresa
  Cadena = Dir(Codigo, vbDirectory)
  If Cadena = "" Then MkDir (Codigo)
  
  If Bloquear_Control Then
     Toolbar1.buttons("ATS").Enabled = False
     Toolbar1.buttons("ATSFinanc").Enabled = False
     Toolbar1.buttons("REOC").Enabled = False
     Toolbar1.buttons("RDEP").Enabled = False
  End If
  
  RatonNormal
End Sub

Private Sub Form_Load()
    ConectarAdodc AdoAux
    ConectarAdodc AdoMes
    ConectarAdodc AdoAir
    ConectarAdodc AdoCodRet
    ConectarAdodc AdoSaldos
    ConectarAdodc AdoVentas
    ConectarAdodc AdoPuntosVentas
    ConectarAdodc AdoCompras
    ConectarAdodc AdoAnulados
    ConectarAdodc AdoRolPagos
    ConectarAdodc AdoCiudades
    ConectarAdodc AdoClientes
    ConectarAdodc AdoImportaciones
    ConectarAdodc AdoExportaciones
    'WBPDF.Navigate "file:///C:/SISTEMA/FONDOS/index_pdf.html"
End Sub

'''Public Sub Numerar_Lineas_SRI(Tabla As String, TipoTrans As String)
'''Dim Contador_AT As Long
'''  Contador_AT = 0
'''  If Tabla <> "" Then
'''     If Tabla = "Trans_Anulados" Then
'''        sSQL = "SELECT * " _
'''             & "FROM Trans_Anulados " _
'''             & "WHERE FechaAnulacion Between #" & FechaIni & "# AND #" & FechaFin & "#  "
'''             If ConSucursal Then
'''                If Len(No_ATS) > 3 Then sSQL = sSQL & "AND Item NOT IN (" & No_ATS & ") "
'''             Else
'''                sSQL = sSQL & "AND Item = '" & NumEmpresa & "' "
'''             End If
'''        sSQL = sSQL & "AND Periodo = '" & Periodo_Contable & "' " _
'''             & "ORDER BY FechaAnulacion "
'''        Select_Adodc AdoAux, sSQL
'''        With AdoAux.Recordset
'''         If .RecordCount > 0 Then
'''             Do While Not .EOF
'''                Contador_AT = Contador_AT + 1
'''               .Fields("Linea_SRI") = Contador_AT
'''               .Update
'''               .MoveNext
'''             Loop
'''         End If
'''        End With
'''     Else
'''        sSQL = "SELECT * " _
'''             & "FROM " & Tabla & " " _
'''             & "WHERE Fecha Between #" & FechaIni & "# AND #" & FechaFin & "# "
'''             If ConSucursal Then
'''                If Len(No_ATS) > 3 Then sSQL = sSQL & "AND Item NOT IN (" & No_ATS & ") "
'''             Else
'''                sSQL = sSQL & "AND Item = '" & NumEmpresa & "' "
'''             End If
'''        sSQL = sSQL & "AND Periodo = '" & Periodo_Contable & "' "
'''        Select Case TipoTrans
'''          Case "C", "V": sSQL = sSQL & "ORDER BY TipoComprobante,IdProv,Fecha "
'''          Case "I", "E": sSQL = sSQL & "ORDER BY IdFiscalProv,Fecha "
'''        End Select
'''        Select_Adodc AdoAir, sSQL
'''        With AdoAir.Recordset
'''         If .RecordCount > 0 Then
'''             Do While Not .EOF
'''                Mifecha = BuscarFecha(.Fields("Fecha"))
'''                Numero = .Fields("Numero")
'''                TipoDoc = .Fields("TP")
'''                Select Case TipoTrans
'''                  Case "C", "V": CodigoCliente = .Fields("IdProv")
'''                                 Factura_No = .Fields("Secuencial")
'''                                 Codigo1 = .Fields("Establecimiento")
'''                                 Codigo2 = .Fields("PuntoEmision")
'''                  Case "I", "E": CodigoCliente = .Fields("IdFiscalProv")
'''                                 Factura_No = .Fields("Correlativo")
'''                End Select
'''                Contador_AT = Contador_AT + 1
'''               .Fields("Linea_SRI") = Contador_AT
'''               'If Tabla = "Trans_Ventas" Then MsgBox Contador_AT & vbCrLf & sSQL
'''               .Update
'''               'Actualizo en el Trans_Air
'''                Select Case TipoTrans
'''                  Case "C": sSQL = "UPDATE Trans_Air " _
'''                                 & "SET Linea_SRI = " & Contador_AT & " " _
'''                                 & "WHERE Periodo = '" & Periodo_Contable & "' "
'''                            If ConSucursal Then
'''                               If Len(No_ATS) > 3 Then sSQL = sSQL & "AND Item NOT IN (" & No_ATS & ") "
'''                            Else
'''                               sSQL = sSQL & "AND Item = '" & NumEmpresa & "' "
'''                            End If
'''                            sSQL = sSQL _
'''                                 & "AND IdProv = '" & CodigoCliente & "' " _
'''                                 & "AND Numero = " & Numero & " " _
'''                                 & "AND TP = '" & TipoDoc & "' " _
'''                                 & "AND Tipo_Trans = 'C' "
'''                            Ejecutar_SQL_SP sSQL
''''''                  Case "V": sSQL = "UPDATE Trans_Air " _
''''''                                 & "SET Linea_SRI = " & Contador_AT & " " _
''''''                                 & "WHERE Periodo = '" & Periodo_Contable & "' "
''''''                                 If ConSucursal Then
''''''                                    If Len(No_ATS) > 3 Then sSQL = sSQL & "AND Item NOT IN (" & No_ATS & ") "
''''''                                 Else
''''''                                    sSQL = sSQL & "AND Item = '" & NumEmpresa & "' "
''''''                                 End If
''''''
''''''                            sSQL = sSQL & "AND IdProv = '" & CodigoCliente & "' " _
''''''                                 & "AND Tipo_Trans = 'V' "
''''''                            Ejecutar_SQL_SP sSQL
'''                  Case "I": sSQL = "UPDATE Trans_Air " _
'''                                 & "SET Linea_SRI = " & Contador_AT & " " _
'''                                 & "WHERE Fecha = #" & Mifecha & "# "
'''                                 If ConSucursal Then
'''                                    If Len(No_ATS) > 3 Then sSQL = sSQL & "AND Item NOT IN (" & No_ATS & ") "
'''                                 Else
'''                                    sSQL = sSQL & "AND Item = '" & NumEmpresa & "' "
'''                                 End If
'''
'''                            sSQL = sSQL & "AND Periodo = '" & Periodo_Contable & "' " _
'''                                 & "AND IdProv = '" & CodigoCliente & "' " _
'''                                 & "AND Numero = " & Numero & " " _
'''                                 & "AND TP = '" & TipoDoc & "' " _
'''                                 & "AND Tipo_Trans = 'I' "
'''                            Ejecutar_SQL_SP sSQL
'''                End Select
'''                Progreso_Barra.Mensaje_Box = "ATS: Renumerando lineas"
'''                Progreso_Esperar True
'''               .MoveNext
'''             Loop
'''         End If
'''        End With
'''     End If
'''  End If
'''End Sub


'''Public Sub Numerar_Lineas_SRI(Tabla As String, TipoTrans As String)
'''Dim Contador_AT As Long
'''  Contador_AT = 0
'''  If Tabla <> "" Then
'''     If Tabla = "Trans_Anulados" Then
'''        sSQL = "SELECT * " _
'''             & "FROM Trans_Anulados " _
'''             & "WHERE FechaAnulacion Between #" & FechaIni & "# AND #" & FechaFin & "#  "
'''             If ConSucursal Then
'''                If Len(No_ATS) > 3 Then sSQL = sSQL & "AND Item NOT IN (" & No_ATS & ") "
'''             Else
'''                sSQL = sSQL & "AND Item = '" & NumEmpresa & "' "
'''             End If
'''        sSQL = sSQL & "AND Periodo = '" & Periodo_Contable & "' " _
'''             & "ORDER BY FechaAnulacion "
'''        Select_Adodc AdoAux, sSQL
'''        With AdoAux.Recordset
'''         If .RecordCount > 0 Then
'''             Do While Not .EOF
'''                Contador_AT = Contador_AT + 1
'''               .Fields("Linea_SRI") = Contador_AT
'''               .Update
'''               .MoveNext
'''             Loop
'''         End If
'''        End With
'''     Else
'''        sSQL = "SELECT * " _
'''             & "FROM " & Tabla & " " _
'''             & "WHERE Fecha Between #" & FechaIni & "# AND #" & FechaFin & "# "
'''             If ConSucursal Then
'''                If Len(No_ATS) > 3 Then sSQL = sSQL & "AND Item NOT IN (" & No_ATS & ") "
'''             Else
'''                sSQL = sSQL & "AND Item = '" & NumEmpresa & "' "
'''             End If
'''        sSQL = sSQL & "AND Periodo = '" & Periodo_Contable & "' "
'''        Select Case TipoTrans
'''          Case "C", "V": sSQL = sSQL & "ORDER BY TipoComprobante,IdProv,Fecha "
'''          Case "I", "E": sSQL = sSQL & "ORDER BY IdFiscalProv,Fecha "
'''        End Select
'''        Select_Adodc AdoAir, sSQL
'''        With AdoAir.Recordset
'''         If .RecordCount > 0 Then
'''             Do While Not .EOF
'''                Select Case TipoTrans
'''                  Case "C", "V": CodigoCliente = .Fields("IdProv")
'''                                 Factura_No = .Fields("Secuencial")
'''                                 Codigo1 = .Fields("Establecimiento")
'''                                 Codigo2 = .Fields("PuntoEmision")
'''                  Case "I", "E": CodigoCliente = .Fields("IdFiscalProv")
'''                                 Factura_No = .Fields("Correlativo")
'''                End Select
'''                Mifecha = BuscarFecha(.Fields("Fecha"))
'''                Contador_AT = Contador_AT + 1
'''               .Fields("Linea_SRI") = Contador_AT
'''               'If Tabla = "Trans_Ventas" Then MsgBox Contador_AT & vbCrLf & sSQL
'''               .Update
'''               'Actualizo en el Trans_Air
'''                Select Case TipoTrans
'''                  Case "C": sSQL = "UPDATE Trans_Air " _
'''                                 & "SET Linea_SRI = " & Contador_AT & " " _
'''                                 & "WHERE Periodo = '" & Periodo_Contable & "' "
'''                                 If ConSucursal Then
'''                                    If Len(No_ATS) > 3 Then sSQL = sSQL & "AND Item NOT IN (" & No_ATS & ") "
'''                                 Else
'''                                    sSQL = sSQL & "AND Item = '" & NumEmpresa & "' "
'''                                 End If
'''                            sSQL = sSQL & "AND IdProv = '" & CodigoCliente & "' " _
'''                                 & "AND Factura_No = " & Factura_No & " " _
'''                                 & "AND EstabFactura = '" & Codigo1 & "' " _
'''                                 & "AND PuntoEmiFactura = '" & Codigo2 & "' " _
'''                                 & "AND Tipo_Trans = 'C' "
'''                            Ejecutar_SQL_SP sSQL
'''                  Case "V": sSQL = "UPDATE Trans_Air " _
'''                                 & "SET Linea_SRI = " & Contador_AT & " " _
'''                                 & "WHERE Periodo = '" & Periodo_Contable & "' "
'''                                 If ConSucursal Then
'''                                    If Len(No_ATS) > 3 Then sSQL = sSQL & "AND Item NOT IN (" & No_ATS & ") "
'''                                 Else
'''                                    sSQL = sSQL & "AND Item = '" & NumEmpresa & "' "
'''                                 End If
'''
'''                            sSQL = sSQL & "AND IdProv = '" & CodigoCliente & "' " _
'''                                 & "AND Tipo_Trans = 'V' "
'''                            Ejecutar_SQL_SP sSQL
'''                  Case "I": sSQL = "UPDATE Trans_Air " _
'''                                 & "SET Linea_SRI = " & Contador_AT & " " _
'''                                 & "WHERE Fecha = #" & Mifecha & "# "
'''                                 If ConSucursal Then
'''                                    If Len(No_ATS) > 3 Then sSQL = sSQL & "AND Item NOT IN (" & No_ATS & ") "
'''                                 Else
'''                                    sSQL = sSQL & "AND Item = '" & NumEmpresa & "' "
'''                                 End If
'''
'''                            sSQL = sSQL & "AND Periodo = '" & Periodo_Contable & "' " _
'''                                 & "AND IdProv = '" & CodigoCliente & "' " _
'''                                 & "AND Factura_No = " & Factura_No & " " _
'''                                 & "AND Tipo_Trans = 'I' "
'''                            Ejecutar_SQL_SP sSQL
'''                End Select
'''                If ProgBar.value < ProgBar.Max Then Progreso_Esperar
'''               .MoveNext
'''             Loop
'''         End If
'''        End With
'''     End If
'''  End If
'''End Sub
'''

'''Public Sub Crear_Anexos_Tipo_01(NombreArchivo As String)
'''Dim RazonSocial As String
'''  Dim NumFile As Long
'''  Dim Cargar_Air_AT As Boolean
'''  RazonSocial = NombreComercial
'''  RazonSocial = UCaseStrg(TrimStrg(Replace(RazonSocial, " ", "")))
'''  NumFile = FreeFile
'''  Open NombreArchivo For Output As #NumFile   ' Abre el archivo
''' 'ENCABEZADO
'''  Print #NumFile, "<?xml version=";
'''  Print #NumFile, Chr(34);
'''  Print #NumFile, "1.0";
'''  Print #NumFile, Chr(34);
'''  Print #NumFile, " encoding=";
'''  Print #NumFile, Chr(34);
'''  Print #NumFile, "ISO-8859-1";
'''  Print #NumFile, Chr(34);
'''  Print #NumFile, " standalone=";
'''  Print #NumFile, Chr(34);
'''  Print #NumFile, "yes";
'''  Print #NumFile, Chr(34);
'''  Print #NumFile, "?>"
'''  Print #NumFile, AbrirXML("iva")
'''  Print #NumFile, vbTab & CampoXML("numeroRuc", RUC)
'''  Print #NumFile, CampoXML("razonSocial", RazonSocial)
'''  Print #NumFile, vbTab & CampoXML("anio", Format(Anio, "0000"))
'''  Print #NumFile, vbTab & CampoXML("mes", Format(NumMeses, "00"))
''' 'COMPRAS
'''  With AdoCompras.Recordset
'''   If .RecordCount > 0 Then
'''       Print #NumFile, AbrirXML("compras")
'''       Do While Not .EOF
'''          TipoDoc = "03"
'''          Select Case .Fields("TD")
'''             Case "R": TipoDoc = "01"
'''             Case "C": TipoDoc = "02"
'''             Case "P": TipoDoc = "03"
'''          End Select
'''         'Aqui genero el Detalle de Compras
'''          Print #NumFile, AbrirXML("detalleCompras")
'''          Print #NumFile, CampoXML("codSustento", .Fields("CodSustento"))
'''          Print #NumFile, CampoXML("tpIdProv", TipoDoc)
'''          Print #NumFile, CampoXML("idProv", .Fields("CI_RUC"))
'''          Print #NumFile, CampoXML("tipoComprobante", .Fields("TipoComprobante"))
'''          Print #NumFile, CampoXML("fechaRegistro", .Fields("FechaRegistro"))
'''          Print #NumFile, CampoXML("establecimiento", .Fields("Establecimiento"))
'''          Print #NumFile, CampoXML("puntoEmision", .Fields("PuntoEmision"))
'''          Print #NumFile, CampoXML("secuencial", .Fields("Secuencial"))
'''          Print #NumFile, CampoXML("fechaEmision", .Fields("FechaEmision"))
'''          Print #NumFile, CampoXML("autorizacion", .Fields("Autorizacion"))
'''          Print #NumFile, CampoXML("baseNoGraIva", Format(0, "#0.00"))         ' hay que aumentar
'''          Print #NumFile, CampoXML("baseImponible", Format(.Fields("BaseImponible"), "#0.00"))
'''          Print #NumFile, CampoXML("baseImpGrav", Format(.Fields("BaseImpGrav"), "#0.00"))
'''          Print #NumFile, CampoXML("montoIce", Format(.Fields("MontoIce"), "#0.00"))
'''          Print #NumFile, CampoXML("montoIva", Format(.Fields("MontoIva"), "#0.00"))
'''          Print #NumFile, CampoXML("valorRetBienes", Format(.Fields("valorRetBienes"), "#0.00"))
'''          If .Fields("ValorRetServicios") = .Fields("MontoIva") Then
'''              Print #NumFile, CampoXML("valorRetServicios", Format(0, "#0.00"))
'''              Print #NumFile, CampoXML("valRetServ100", Format(.Fields("ValorRetServicios"), "#0.00"))    'hay que aumentar
'''          Else
'''              Print #NumFile, CampoXML("valorRetServicios", Format(.Fields("ValorRetServicios"), "#0.00"))
'''              Print #NumFile, CampoXML("valRetServ100", Format(0, "#0.00"))    'hay que aumentar
'''          End If
'''         'Asigno a variables los campos para el concepto Air
'''          CodigoCliente = .Fields("IdProv")
'''          Factura_No = .Fields("Secuencial")
'''          Codigo1 = .Fields("Establecimiento")
'''          Codigo2 = .Fields("PuntoEmision")
'''          Estab = "000"
'''          PtoRet = "000"
'''          FechaRet = .Fields("FechaRegistro")
'''          AutRet = "0000000000"
'''          Number = False
'''         'Genero el AIR
'''          sSQL = "SELECT * " _
'''               & "FROM Trans_Air " _
'''               & "WHERE Fecha Between #" & FechaIni & "# AND #" & FechaFin & "# " _
'''               & "AND Periodo = '" & Periodo_Contable & "' "
'''  If ConSucursal Then
'''     If Len(No_ATS) > 3 Then sSQL = sSQL & "AND Item NOT IN (" & No_ATS & ") "
'''  Else
'''     sSQL = sSQL & "AND Item = '" & NumEmpresa & "' "
'''  End If
'''
'''          sSQL = sSQL & "AND T = '" & Normal & "' " _
'''               & "AND Tipo_Trans = 'C' " _
'''               & "AND IdProv = '" & CodigoCliente & "' " _
'''               & "AND Factura_No = " & Factura_No & " " _
'''               & "AND EstabFactura = '" & Codigo1 & "' " _
'''               & "AND PuntoEmiFactura = '" & Codigo2 & "' " _
'''               & "ORDER BY SecRetencion,EstabFactura,PuntoEmiFactura,Factura_No "
'''          Select_Adodc AdoAir, sSQL
'''          If AdoAir.Recordset.RecordCount > 0 Then
'''             Print #NumFile, AbrirXML("air")
'''             Do While Not AdoAir.Recordset.EOF
'''               'Asigno en variables los campos
'''                Codigo = AdoAir.Recordset.Fields("CodRet")
'''                Total_SubTotal = AdoAir.Recordset.Fields("BaseImp")
'''                Porcentaje = AdoAir.Recordset.Fields("Porcentaje")
'''                Valor = AdoAir.Recordset.Fields("ValRet")
'''                Estab = AdoAir.Recordset.Fields("EstabRetencion")
'''                PtoRet = AdoAir.Recordset.Fields("PtoEmiRetencion")
'''                SecRet = AdoAir.Recordset.Fields("SecRetencion")
'''                AutRet = AdoAir.Recordset.Fields("AutRetencion")
'''                If Len(AutRet) < 3 Then AutRet = "0000000000"
'''                Number = True
'''                Print #NumFile, AbrirXML("detalleAir")
'''                Print #NumFile, CampoXML("codRetAir", Codigo)
'''                Print #NumFile, CampoXML("baseImpAir", Format(Total_SubTotal, "##0.00"))
'''                Print #NumFile, CampoXML("porcentajeAir", Format(Porcentaje * 100, "#0.00"))
'''                Print #NumFile, CampoXML("valRetAir", Format(Valor, "##0.00"))
'''                Print #NumFile, CerrarXML("detalleAir")
'''                AdoAir.Recordset.MoveNext
'''             Loop
'''             Print #NumFile, CerrarXML("air")
'''          Else
'''             Print #NumFile, VacioXML("air")
'''          End If
'''         'Aqui se genera la segunda parte del detalle Compras
'''          Print #NumFile, CampoXML("estabRetencion1", Estab)
'''          Print #NumFile, CampoXML("ptoEmiRetencion1", PtoRet)
'''          If Number = False Then
'''             Print #NumFile, CampoXML("secRetencion1", 0)
'''          Else
'''             Print #NumFile, CampoXML("secRetencion1", Format(SecRet, "#0"))
'''          End If
'''          Print #NumFile, CampoXML("autRetencion1", AutRet)
'''          Print #NumFile, CampoXML("fechaEmiRet1", FechaRet)
'''          Print #NumFile, CampoXML("docModificado", Val(.Fields("DocModificado")))
'''          Print #NumFile, CampoXML("estabModificado", .Fields("EstabModificado"))
'''          Print #NumFile, CampoXML("ptoEmiModificado", .Fields("PtoEmiModificado"))
'''          Print #NumFile, CampoXML("secModificado", .Fields("SecModificado"))
'''          Print #NumFile, CampoXML("autModificado", .Fields("AutModificado"))
'''          Print #NumFile, CerrarXML("detalleCompras")
'''         .MoveNext
'''       Loop
'''       Print #NumFile, CerrarXML("compras")
'''   Else
'''       Print #NumFile, VacioXML("compras")
'''   End If
'''  End With
''''' 'VENTAS
'''''  FechaRet = UltimoDiaMes("01/" & NumMeses & "/" & Anio)
'''''  With AdoVentas.Recordset
'''''   If .RecordCount > 0 Then
'''''       Print #NumFile, AbrirXML("ventas")
'''''       Do While Not .EOF
'''''          Contador = .Fields("NumeroComprobantesV")
'''''          If Contador <= 0 Then Contador = 1
'''''          Factura_No = Contador
'''''          CodigoCliente = .Fields("IdProv")
'''''          Codigo1 = .Fields("Establecimiento")
'''''          Codigo2 = .Fields("PuntoEmision")
'''''          Estab = "000"
'''''          PtoRet = "000"
'''''          'FechaRet = .Fields("FechaRegistro")
'''''          AutRet = "0000000000"
'''''          TipoDoc = "06"
'''''          Select Case .Fields("TD_SRI")
'''''            Case "R": TipoDoc = "04"
'''''            Case "C": TipoDoc = "05"
'''''            Case "P": TipoDoc = "06"
'''''            Case "F": TipoDoc = "07"
'''''          End Select
'''''          If .Fields("CI_RUC_SRI") = "9999999999999" Then TipoDoc = "07"
'''''         'Genero el talón de Ventas
'''''          Print #NumFile, AbrirXML("detalleVentas")
'''''          Print #NumFile, CampoXML("tpIdCliente", TipoDoc)
'''''          Print #NumFile, CampoXML("idCliente", .Fields("CI_RUC_SRI"))
'''''          Print #NumFile, CampoXML("tipoComprobante", .Fields("TipoComprobante"))
'''''          Print #NumFile, CampoXML("numeroComprobantes", Contador)
'''''          Print #NumFile, CampoXML("baseNoGraIva", Format(0, "#0.00"))         ' hay que aumentar
'''''          Print #NumFile, CampoXML("baseImponible", Format(.Fields("BaseImponibleV"), "#0.00"))
'''''          Print #NumFile, CampoXML("baseImpGrav", Format(.Fields("BaseImpGravV"), "#0.00"))
'''''          Print #NumFile, CampoXML("montoIva", Format(.Fields("MontoIvaV"), "#0.00"))
'''''          Print #NumFile, CampoXML("valorRetIva", Format(.Fields("ValorRetBienesV") + .Fields("ValorRetServiciosV"), "#0.00"))
'''''         'Asigno a variables los campos
'''''         'Genero el AIR
'''''          Cargar_Air_AT = True
'''''          Select Case Val(.Fields("TipoComprobante"))
'''''            Case 4, 41: Cargar_Air_AT = False
'''''          End Select
'''''          SubTotal = 0
'''''          If Cargar_Air_AT Then
'''''             sSQL = "SELECT * " _
'''''                  & "FROM Trans_Air " _
'''''                  & "WHERE Fecha Between #" & FechaIni & "# AND #" & FechaFin & "#  " _
'''''                  & "AND Periodo = '" & Periodo_Contable & "' "
'''''             If ConSucursal = False Then sSQL = sSQL & "AND Item = '" & NumEmpresa & "' "
'''''             sSQL = sSQL & "AND T = '" & Normal & "' " _
'''''                  & "AND Tipo_Trans = 'V' " _
'''''                  & "AND IdProv = '" & CodigoCliente & "' " _
'''''                  & "AND EstabFactura = '" & Codigo1 & "' " _
'''''                  & "AND PuntoEmiFactura = '" & Codigo2 & "' " _
'''''                  & "ORDER BY SecRetencion,EstabFactura,PuntoEmiFactura,Factura_No "
'''''             Select_Adodc AdoAir, sSQL
'''''             If AdoAir.Recordset.RecordCount > 0 Then
'''''                Do While Not AdoAir.Recordset.EOF
'''''                   SubTotal = SubTotal + AdoAir.Recordset.Fields("ValRet")
'''''                   AdoAir.Recordset.MoveNext
'''''                Loop
'''''             End If
'''''          End If
'''''          Print #NumFile, CampoXML("valorRetRenta", Format(SubTotal, "#0.00"))
'''''          Print #NumFile, CerrarXML("detalleVentas")
'''''         .MoveNext
'''''       Loop
'''''       Print #NumFile, CerrarXML("ventas")
'''''   Else
'''''       Print #NumFile, VacioXML("ventas")
'''''   End If
'''''  End With
''' 'EXPORTACIONES
'''  With AdoExportaciones.Recordset
'''   If .RecordCount > 0 Then
'''       Print #NumFile, AbrirXML("exportaciones")
'''       Do While Not .EOF
'''          TipoDoc = "1"
'''          Select Case .Fields("TD")
'''            Case "R", "C": TipoDoc = "2"
'''          End Select
'''         'Genero el talón de Exportaciones
'''          Print #NumFile, AbrirXML("detalleExportaciones")
'''          Print #NumFile, CampoXML("exportacionDe", .Fields("ExportacionDe"))
'''          Print #NumFile, CampoXML("tipoComprobante", .Fields("TipoComprobante"))
'''          If .Fields("ExportacionDe") = 1 Then
'''              Print #NumFile, CampoXML("distAduanero", .Fields("DistAduanero"))
'''              Print #NumFile, CampoXML("anio", .Fields("Anio"))
'''              Print #NumFile, CampoXML("regimen", .Fields("Regimen"))
'''              Print #NumFile, CampoXML("correlativo", .Fields("Correlativo"))
'''              Print #NumFile, CampoXML("verificador", .Fields("Verificador"))
'''              Print #NumFile, CampoXML("docTransp", .Fields("NumeroDctoTransporte"))
'''          End If
'''          Print #NumFile, CampoXML("fechaEmbarque", .Fields("FechaEmbarque"))
'''          Print #NumFile, CampoXML("fue", "000001")                                       '*
'''          Print #NumFile, CampoXML("valorFOB", Format(.Fields("ValorFOB"), "#0.00"))
'''          Print #NumFile, CampoXML("valorFOBComprobante", Format(.Fields("ValorFOBComprobante")))
'''          Print #NumFile, CampoXML("establecimiento", .Fields("Establecimiento"))
'''          Print #NumFile, CampoXML("puntoEmision", .Fields("PuntoEmision"))
'''          Print #NumFile, CampoXML("secuencial", .Fields("Secuencial"))
'''          Print #NumFile, CampoXML("autorizacion", Format(Val(.Fields("Autorizacion")), "#0"))
'''          Print #NumFile, CampoXML("fechaEmision", .Fields("FechaEmision"))
'''          Print #NumFile, CerrarXML("detalleExportaciones")
'''         .MoveNext
'''       Loop
'''       Print #NumFile, CerrarXML("exportaciones")
'''   Else
'''       Print #NumFile, VacioXML("exportaciones")
'''   End If
'''  End With
''' 'ANULADOS
'''   With AdoAnulados.Recordset
'''    If .RecordCount > 0 Then
'''        Print #NumFile, AbrirXML("anulados")
'''        Do While Not .EOF
'''           Print #NumFile, AbrirXML("detalleAnulados")
'''           Print #NumFile, CampoXML("tipoComprobante", .Fields("TipoComprobante"))
'''           Print #NumFile, CampoXML("establecimiento", .Fields("Establecimiento"))
'''           Print #NumFile, CampoXML("puntoEmision", .Fields("PuntoEmision"))
'''           Print #NumFile, CampoXML("secuencialInicio", .Fields("Secuencial1"))
'''           Print #NumFile, CampoXML("secuencialFin", .Fields("Secuencial2"))
'''           Print #NumFile, CampoXML("autorizacion", .Fields("Autorizacion"))
'''           Print #NumFile, CerrarXML("detalleAnulados")
'''          .MoveNext
'''        Loop
'''        Print #NumFile, CerrarXML("anulados")
'''    Else
'''       Print #NumFile, VacioXML("anulados")
'''    End If
'''   End With
'''  Print #NumFile, CerrarXML("iva")
'''  Close NumFile   ' Cierra los 3 archivos abiertos.
'''  Archivo_XML = NombreArchivo
'''End Sub

Public Sub Crear_Anexos_Tipo_01_Marzo_2015(NombreArchivo As String, ATFisico As Boolean)
Dim NumFile As Long
Dim Cargar_Air_AT As Boolean
Dim Poner_Porc_Air As Boolean
Dim Generar_Ventas As Boolean
Dim RazonSocial As String
Dim TipoDeFactura As String
Dim Numero_Establecimientos As Integer
Dim Tipo_Pago As String
Dim Tipo_Pagos(8) As String
Dim Total_Ventas(1 To 999) As Currency
Dim Total_NC(1 To 999) As Currency
Dim Total_Reembolso(1 To 999) As Currency
Dim Total_Ventas_Estab As Currency
Dim Total_Reembolso_Estab As Currency
Dim Total_NC_Estab As Currency
Dim Suma_Ventas As Currency
Dim ValRetBien10 As Currency
Dim ValRetServ20 As Currency
Dim ValRetServ50 As Currency
  RatonReloj
  Total_Reembolso_Estab = 0
  Total_Ventas_Estab = 0
  Total_NC_Estab = 0
  Numero_Establecimientos = 0
  For IE = 1 To 999
      Total_Reembolso(IE) = 0
      Total_Ventas(IE) = 0
      Total_NC(IE) = 0
  Next IE
 'MsgBox "NC: " & NombreComercial
  RazonSocial = Replace(NombreComercial, ".", " ")
  RazonSocial = UCaseStrg(TrimStrg(Replace(RazonSocial, " ", "")))
 'Totalizamos de Ventas
  With AdoPuntosVentas.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          KE = CInt(.fields("Establecimiento"))
          Suma_Ventas = .fields("BaseImponibleV")
          If Not IsNull(.fields("BaseImpGravV")) Then Suma_Ventas = Suma_Ventas + .fields("BaseImpGravV")  ' + .Fields("MontoIvaV")
          If .fields("TipoComprobante") = 4 Then
              Total_NC_Estab = Total_NC_Estab + Suma_Ventas
              Total_NC(KE) = Total_NC(KE) + Suma_Ventas
          ElseIf .fields("TipoComprobante") = 41 Then
              Total_Reembolso(KE) = Total_Reembolso(KE) + Suma_Ventas
              Total_Reembolso_Estab = Total_Reembolso_Estab + Suma_Ventas
          Else
              Total_Ventas(KE) = Total_Ventas(KE) + Suma_Ventas
              Total_Ventas_Estab = Total_Ventas_Estab + Suma_Ventas
          End If
         .MoveNext
       Loop
      .MoveFirst
   End If
  End With
  For IE = 1 To 999
      If (Total_Ventas(IE) - Total_NC(IE) - Total_Reembolso(IE)) <> 0 Then Numero_Establecimientos = Numero_Establecimientos + 1
  Next IE
'' MsgBox Numero_Establecimientos
 'Numero libre en el disco para grabar
  NumFile = FreeFile
 'Abrimos el archivo para grabar
  Open NombreArchivo For Output As #NumFile   ' Abre el archivo
 
 'ENCABEZADO
 'MsgBox RazonSocial & vbCrLf & Anio & vbCrLf & NumMeses
  Print #NumFile, "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""no""?>"
  Print #NumFile, AbrirXML("iva")
  Print #NumFile, vbTab & CampoXML("TipoIDInformante", "R")
  Print #NumFile, vbTab & CampoXML("IdInformante", RUC)
  Print #NumFile, vbTab & CampoXML("razonSocial", RazonSocial)
  Print #NumFile, vbTab & CampoXML("Anio", Format(Anio, "0000"))
  Print #NumFile, vbTab & CampoXML("Mes", Format(NumMeses, "00"))
  
  Select Case Mes_Anexo
    Case "1er_Semestre", "2do_Semestre"
         Print #NumFile, vbTab & CampoXML("regimenMicroempresa", "SI")
  End Select

  If (Total_Ventas_Estab - Total_NC_Estab) <> 0 Then
     Print #NumFile, vbTab & CampoXML("numEstabRuc", Format(Numero_Establecimientos, "000"))
  End If
  Print #NumFile, vbTab & CampoXML("totalVentas", Format(Total_Ventas_Estab - Total_NC_Estab, "#0.00"))
  Print #NumFile, vbTab & CampoXML("codigoOperativo", "IVA")
  
 'COMPRAS
  With AdoCompras.Recordset
   If .RecordCount > 0 Then
       Print #NumFile, AbrirXML("compras")
       Do While Not .EOF
          'If .Fields("TP") = "CE" And .Fields("Numero") = 3000338 Then MsgBox "..."
          TipoDoc = "03"
          Select Case .fields("TD")
             Case "R": TipoDoc = "01"
             Case "C": TipoDoc = "02"
             Case "P": TipoDoc = "03"
          End Select
         'Aqui genero el Detalle de Compras
          Print #NumFile, AbrirXML("detalleCompras")
          Print #NumFile, CampoXML("codSustento", .fields("CodSustento"))
          Print #NumFile, CampoXML("tpIdProv", TipoDoc)
          Print #NumFile, CampoXML("idProv", .fields("CI_RUC"))
          Print #NumFile, CampoXML("tipoComprobante", Format(.fields("TipoComprobante"), "00"))
         'Si es Pasaporte genera la parte relacionada
          If .fields("TD") = "P" Then
             If CFechaLong(FechaInicial) >= CFechaLong("01/05/2016") Then
                Print #NumFile, CampoXML("tipoProv", "01")
                Print #NumFile, CampoXML("denoProv", .fields("Cliente"))
                Print #NumFile, CampoXML("parteRel", .fields("Parte_Relacionada"))
             Else
                Print #NumFile, CampoXML("tipoProv", .fields("Tipo_Pasaporte"))
                Print #NumFile, CampoXML("parteRel", .fields("Parte_Relacionada"))
             End If
          Else
             Print #NumFile, CampoXML("tipoProv", "01")
             Print #NumFile, CampoXML("parteRel", "NO")
          End If
          Print #NumFile, CampoXML("fechaRegistro", .fields("FechaRegistro"))
          Print #NumFile, CampoXML("establecimiento", .fields("Establecimiento"))
          Print #NumFile, CampoXML("puntoEmision", .fields("PuntoEmision"))
          Print #NumFile, CampoXML("secuencial", .fields("Secuencial"))
          Print #NumFile, CampoXML("fechaEmision", .fields("FechaEmision"))
          Print #NumFile, CampoXML("autorizacion", .fields("Autorizacion"))
          Print #NumFile, CampoXML("baseNoGraIva", Format(.fields("BaseNoObjIVA"), "#0.00"))         ' hay que aumentar
          Print #NumFile, CampoXML("baseImponible", Format(.fields("BaseImponible"), "#0.00"))
          Print #NumFile, CampoXML("baseImpGrav", Format(.fields("BaseImpGrav"), "#0.00"))
          Print #NumFile, CampoXML("baseImpExe", "0.00")      'Por Adaptar en la pantalla
          Print #NumFile, CampoXML("montoIce", Format(.fields("MontoIce"), "#0.00"))
          Print #NumFile, CampoXML("montoIva", Format(.fields("MontoIva"), "#0.00"))
          ValRetBien10 = 0
          ValRetServ20 = 0
          ValRetServ50 = 0
          If .fields("Porc_Bienes") = "10" Then ValRetBien10 = .fields("valorRetBienes")
          If .fields("Porc_Servicios") = "20" Then ValRetServ20 = .fields("ValorRetServicios")
          If .fields("Porc_Servicios") = "50" Then ValRetServ50 = .fields("ValorRetServicios")
           
          Print #NumFile, CampoXML("valRetBien10", Format(ValRetBien10, "#0.00"))
          Print #NumFile, CampoXML("valRetServ20", Format(ValRetServ20, "#0.00"))
          If ValRetBien10 = 0 Then
             Print #NumFile, CampoXML("valorRetBienes", Format(.fields("valorRetBienes"), "#0.00"))
          Else
             Print #NumFile, CampoXML("valorRetBienes", "0.00")
          End If
          Print #NumFile, CampoXML("valRetServ50", Format(ValRetServ50, "#0.00"))
          If ValRetServ20 = 0 Then
             If .fields("ValorRetServicios") = .fields("MontoIva") Then
                 Print #NumFile, CampoXML("valorRetServicios", Format(0, "#0.00"))
                 Print #NumFile, CampoXML("valRetServ100", Format(.fields("ValorRetServicios"), "#0.00"))    'hay que aumentar
             Else
                 Print #NumFile, CampoXML("valorRetServicios", Format(.fields("ValorRetServicios"), "#0.00"))
                 Print #NumFile, CampoXML("valRetServ100", Format(0, "#0.00"))    'hay que aumentar
             End If
          Else
             Print #NumFile, CampoXML("valorRetServicios", Format(0, "#0.00"))
             Print #NumFile, CampoXML("valRetServ100", Format(0, "#0.00"))    'hay que aumentar
          End If
                    
          Print #NumFile, CampoXML("totbasesImpReemb", "0.00")      'Por Adaptar en la pantalla
          Print #NumFile, AbrirXML("pagoExterior")
                Print #NumFile, CampoXML("pagoLocExt", .fields("PagoLocExt"))
                If IsNumeric(.fields("PaisEfecPago")) Then
                   Print #NumFile, CampoXML("paisEfecPago", Val(.fields("PaisEfecPago")))
                Else
                   Print #NumFile, CampoXML("paisEfecPago", .fields("PaisEfecPago"))
                End If
                Print #NumFile, CampoXML("aplicConvDobTrib", .fields("AplicConvDobTrib"))
                Print #NumFile, CampoXML("pagExtSujRetNorLeg", .fields("PagExtSujRetNorLeg"))
               'Print #NumFile, CampoXML("pagoRegFis", "NA")  'Por programar en la pantalla
          Print #NumFile, CerrarXML("pagoExterior")
          
          If (.fields("BaseNoObjIVA") + .fields("BaseImponible") + .fields("BaseImpGrav") + .fields("MontoIva")) >= 1000 Then
             Print #NumFile, AbrirXML("formasDePago")
                   Print #NumFile, CampoXML("formaPago", .fields("FormaPago"))
             Print #NumFile, CerrarXML("formasDePago")
          End If
         'Asigno a variables los campos para el concepto Air
          CodigoCliente = .fields("IdProv")
          Factura_No = .fields("Secuencial")
          Codigo1 = .fields("Establecimiento")
          Codigo2 = .fields("PuntoEmision")
          Estab = "000"
          PtoRet = "000"
          FechaRet = .fields("FechaRegistro")
          AutRet = "0000000000"
          SecRet = 0
          Number = False
          Poner_Porc_Air = False
         'Genero el AIR
          sSQL = "SELECT * " _
               & "FROM Trans_Air " _
               & "WHERE Fecha Between #" & FechaIni & "# AND #" & FechaFin & "# " _
               & "AND Periodo = '" & Periodo_Contable & "' " _
               & "AND Item IN (" & sItem & ") " _
               & "AND T = '" & Normal & "' " _
               & "AND Tipo_Trans = 'C' " _
               & "AND IdProv = '" & CodigoCliente & "' " _
               & "AND Factura_No = " & Factura_No & " " _
               & "AND EstabFactura = '" & Codigo1 & "' " _
               & "AND PuntoEmiFactura = '" & Codigo2 & "' " _
               & "ORDER BY SecRetencion,EstabFactura,PuntoEmiFactura,Factura_No "
          Select_Adodc AdoAir, sSQL
          If AdoAir.Recordset.RecordCount > 0 Then
             Print #NumFile, AbrirXML("air")
             Do While Not AdoAir.Recordset.EOF
               'Asigno en variables los campos
                If Porcentaje <> 0 Then Poner_Porc_Air = True
                Codigo = AdoAir.Recordset.fields("CodRet")
                Total_SubTotal = AdoAir.Recordset.fields("BaseImp")
                Porcentaje = AdoAir.Recordset.fields("Porcentaje")
                Valor = AdoAir.Recordset.fields("ValRet")
                Estab = AdoAir.Recordset.fields("EstabRetencion")
                PtoRet = AdoAir.Recordset.fields("PtoEmiRetencion")
                SecRet = AdoAir.Recordset.fields("SecRetencion")
                AutRet = AdoAir.Recordset.fields("AutRetencion")
                If Len(AutRet) < 3 Then AutRet = "0000000000"
                Number = True
                Print #NumFile, AbrirXML("detalleAir")
                Print #NumFile, CampoXML("codRetAir", Codigo)
                Print #NumFile, CampoXML("baseImpAir", Format(Total_SubTotal, "##0.00"))
                Print #NumFile, CampoXML("porcentajeAir", Format(Porcentaje * 100, "#0.00"))
                Print #NumFile, CampoXML("valRetAir", Format(Valor, "##0.00"))
                Print #NumFile, CerrarXML("detalleAir")
                'If .Fields("TP") = "CE" And .Fields("Numero") = 3000338 Then MsgBox "..."
                AdoAir.Recordset.MoveNext
             Loop
             Print #NumFile, CerrarXML("air")
            'Aqui se genera la segunda parte del detalle Compras
             If Poner_Porc_Air And SecRet > 0 Then
                Print #NumFile, CampoXML("estabRetencion1", Estab)
                Print #NumFile, CampoXML("ptoEmiRetencion1", PtoRet)
                If Number = False Then
                   Print #NumFile, CampoXML("secRetencion1", 0)
                Else
                   Print #NumFile, CampoXML("secRetencion1", Format(SecRet, "#0"))
                End If
                Print #NumFile, CampoXML("autRetencion1", AutRet)
                Print #NumFile, CampoXML("fechaEmiRet1", FechaRet)
             End If
          End If
         'Aqui colocamos si existe una nota de credito o debito
          Select Case Val(.fields("TipoComprobante"))
            Case 4, 5, 41
                 Print #NumFile, CampoXML("docModificado", Format(Val(.fields("DocModificado")), "00"))
                 Print #NumFile, CampoXML("estabModificado", .fields("EstabModificado"))
                 Print #NumFile, CampoXML("ptoEmiModificado", .fields("PtoEmiModificado"))
                 Print #NumFile, CampoXML("secModificado", .fields("SecModificado"))
                 Print #NumFile, CampoXML("autModificado", .fields("AutModificado"))
          End Select
          Print #NumFile, CerrarXML("detalleCompras")
         .MoveNext
       Loop
       Print #NumFile, CerrarXML("compras")
   Else
       Print #NumFile, VacioXML("compras")
   End If
  End With
  
 'VENTAS
  Generar_Ventas = False
  If CFechaLong(FechaInicial) < CFechaLong("01/01/2016") Then Generar_Ventas = True
  sSQL = "SELECT Autorizacion " _
       & "FROM Facturas " _
       & "WHERE Periodo = '" & Periodo_Contable & "' " _
       & "AND Item IN (" & sItem & ") " _
       & "AND Fecha Between #" & FechaIni & "# AND #" & FechaFin & "# "
  If Not ATFisico Then sSQL = sSQL & "AND LEN(Autorizacion) < 13 "
  sSQL = sSQL & "ORDER BY Autorizacion "
  Select_Adodc AdoAir, sSQL
  If AdoAir.Recordset.RecordCount > 0 Then
     Generar_Ventas = True
     TipoDeFactura = "F"
  Else
     TipoDeFactura = "E"
  End If
 'MsgBox TipoDeFactura & vbCrLf & Generar_Ventas
 'Generar_Ventas = True
  
  If Generar_Ventas Then
  With AdoVentas.Recordset
   If .RecordCount > 0 Then
       Print #NumFile, AbrirXML("ventas")
       Do While Not .EOF
          Contador = .fields("NumeroComprobantesV")
          If Contador <= 0 Then Contador = 1
          Factura_No = Contador
          CodigoCliente = .fields("RUC_CI")
           
          'Codigo1 = .Fields("Establecimiento") '"001" '
          'Codigo2 = .Fields("PuntoEmision") ' "001" '
          Estab = "000"
          PtoRet = "000"
          'FechaRet = .Fields("FechaRegistro")
          FechaRet = FechaFinal
          AutRet = "0000000000"
          TipoDoc = "06"
          Select Case .fields("TB")
            Case "R": TipoDoc = "04"
            Case "C": TipoDoc = "05"
            Case "P": TipoDoc = "06"
            Case "F": TipoDoc = "07"
          End Select
          If .fields("RUC_CI") = "9999999999999" Then TipoDoc = "07"
         'Genero el talón de Ventas
          'If .Fields("RUC_CI") = Ninguno Then MsgBox .Fields("RUC_CI")
          Print #NumFile, AbrirXML("detalleVentas")
          Print #NumFile, CampoXML("tpIdCliente", TipoDoc)
          Print #NumFile, CampoXML("idCliente", .fields("RUC_CI"))
          If TipoDoc <> "07" Then Print #NumFile, CampoXML("parteRelVtas", "NO")
          If TipoDoc = "06" And CFechaLong(FechaInicial) >= CFechaLong("01/05/2016") Then
             Print #NumFile, CampoXML("tipoCliente", "01")
             Print #NumFile, CampoXML("denoCli", .fields("Razon_Social"))
          End If
          Print #NumFile, CampoXML("tipoComprobante", Format(.fields("TipoComprobante"), "00"))
          Print #NumFile, CampoXML("tipoEmision", TipoDeFactura)
          Print #NumFile, CampoXML("numeroComprobantes", Contador)
          Print #NumFile, CampoXML("baseNoGraIva", Format(0, "#0.00"))         ' hay que aumentar
          Print #NumFile, CampoXML("baseImponible", Format(.fields("BaseImponibleV"), "#0.00"))
          Print #NumFile, CampoXML("baseImpGrav", Format(.fields("BaseImpGravV"), "#0.00"))
          Print #NumFile, CampoXML("montoIva", Format(.fields("MontoIvaV"), "#0.00"))
          Print #NumFile, CampoXML("montoIce", "0.00")
          Print #NumFile, CampoXML("valorRetIva", Format(.fields("ValorRetBienesV") + .fields("ValorRetServiciosV"), "#0.00"))
         'Asigno a variables los campos
         'Genero el AIR
          Cargar_Air_AT = True
          Select Case Val(.fields("TipoComprobante"))
            Case 4, 41: Cargar_Air_AT = False   ' NC o ND
          End Select
          SubTotal = 0
          If Cargar_Air_AT Then
             sSQL = "SELECT IdProv,SUM(ValRet) As TValRet " _
                  & "FROM Trans_Air " _
                  & "WHERE Fecha Between #" & FechaIni & "# AND #" & FechaFin & "#  " _
                  & "AND Periodo = '" & Periodo_Contable & "' " _
                  & "AND Item IN (" & sItem & ") " _
                  & "AND T = '" & Normal & "' " _
                  & "AND Tipo_Trans = 'V' " _
                  & "AND RUC_CI = '" & CodigoCliente & "' " _
                  & "GROUP BY IdProv "
             Select_Adodc AdoAir, sSQL
             'MsgBox sSQL & vbCrLf & AdoAir.Recordset.RecordCount
             If AdoAir.Recordset.RecordCount > 0 Then SubTotal = SubTotal + AdoAir.Recordset.fields("TValRet")
            'MsgBox sSQL & vbCrLf & AdoAir.Recordset.RecordCount
          End If
         'MsgBox SubTotal
          Print #NumFile, CampoXML("valorRetRenta", Format(SubTotal, "#0.00"))
          
          If CFechaLong(FechaInicial) >= CFechaLong("01/06/2016") Then
             For I = 0 To 7
                 Tipo_Pagos(I) = ""
             Next I
            'Ventas Normales
             sSQL = "SELECT Tipo_Pago " _
                  & "FROM Facturas " _
                  & "WHERE Fecha Between #" & FechaIni & "# AND #" & FechaFin & "#  " _
                  & "AND Periodo = '" & Periodo_Contable & "' " _
                  & "AND LEN(Tipo_Pago) > 1 " _
                  & "AND Item IN (" & sItem & ") " _
                  & "AND RUC_CI = '" & CodigoCliente & "' " _
                  & "GROUP BY Tipo_Pago " _
                  & "ORDER BY Tipo_Pago "
             Select_Adodc AdoAir, sSQL
             If AdoAir.Recordset.RecordCount > 0 Then
                Do While Not AdoAir.Recordset.EOF
                   Tipo_Pago = AdoAir.Recordset.fields("Tipo_Pago")
                   For I = 0 To 7
                       If Tipo_Pago <> Tipo_Pagos(I) Then
                          Tipo_Pagos(I) = Tipo_Pago
                          Exit For
                       End If
                   Next I
                   AdoAir.Recordset.MoveNext
                Loop
             End If
             
            'Ventas sin modulos de Facturacion
             sSQL = "SELECT Tipo_Pago " _
                  & "FROM Trans_Ventas " _
                  & "WHERE Fecha Between #" & FechaIni & "# AND #" & FechaFin & "#  " _
                  & "AND Periodo = '" & Periodo_Contable & "' " _
                  & "AND LEN(Tipo_Pago) > 1 " _
                  & "AND Item IN (" & sItem & ") " _
                  & "AND RUC_CI = '" & CodigoCliente & "' " _
                  & "GROUP BY Tipo_Pago " _
                  & "ORDER BY Tipo_Pago "
             Select_Adodc AdoAir, sSQL
             If AdoAir.Recordset.RecordCount > 0 Then
                Do While Not AdoAir.Recordset.EOF
                   Tipo_Pago = AdoAir.Recordset.fields("Tipo_Pago")
                   For I = 0 To 7
                       If Tipo_Pago <> Tipo_Pagos(I) Then
                          Tipo_Pagos(I) = Tipo_Pago
                          Exit For
                       End If
                   Next I
                   AdoAir.Recordset.MoveNext
                Loop
             End If
             
            'Verificamos si ponemo o no el tipo de pago
             Select Case Val(.fields("TipoComprobante"))
               Case 4
                    Tipo_Pago = ""
               Case Else
                    Tipo_Pago = ""
                    For I = 0 To 7
                        If Tipo_Pagos(I) <> "" Then Tipo_Pago = Tipo_Pago & Tipo_Pagos(I) & ","
                    Next I
                    If Tipo_Pago = "" Then
                       Tipo_Pago = "01,"
                       Tipo_Pagos(0) = "01"
                    End If
             End Select
            'If Len(Tipo_Pago) > 3 Then MsgBox Tipo_Pago
             If Tipo_Pago <> "" Then
                Print #NumFile, AbrirXML("formasDePago")
                For I = 0 To 7
                    If Tipo_Pagos(I) <> "" Then Print #NumFile, CampoXML("formaPago", Tipo_Pagos(I))
                Next I
                Print #NumFile, CerrarXML("formasDePago")
             End If
          End If
          Print #NumFile, CerrarXML("detalleVentas")
         .MoveNext
       Loop
       Print #NumFile, CerrarXML("ventas")
   End If
  End With
 'Resumen Total de Ventas por Establecimientos
  If (Total_Ventas_Estab + Total_NC_Estab) > 0 Then
     Print #NumFile, AbrirXML("ventasEstablecimiento")
     For KE = 1 To 999
      If Total_Ventas(KE) + Total_NC(KE) + Total_Reembolso(KE) > 0 Then
         Print #NumFile, AbrirXML("ventaEst")
         Print #NumFile, CampoXML("codEstab", Format(KE, "000"))
         Print #NumFile, CampoXML("ventasEstab", Format(Total_Ventas(KE) - Total_NC(KE), "#0.00"))
         Print #NumFile, CerrarXML("ventaEst")
         '+ Total_Reembolso(KE)
      End If
     Next KE
     Print #NumFile, CerrarXML("ventasEstablecimiento")
  End If
  End If
 'EXPORTACIONES
  With AdoExportaciones.Recordset
   If .RecordCount > 0 Then
       Print #NumFile, AbrirXML("exportaciones")
       Do While Not .EOF
          TipoDoc = "1"
          Select Case .fields("TD")
            Case "R", "C": TipoDoc = "2"
          End Select
         'Genero el talón de Exportaciones
          Print #NumFile, AbrirXML("detalleExportaciones")
          Print #NumFile, CampoXML("exportacionDe", .fields("ExportacionDe"))
          Print #NumFile, CampoXML("tipoComprobante", Format(.fields("TipoComprobante"), "00"))
          If .fields("ExportacionDe") = 1 Then
              Print #NumFile, CampoXML("distAduanero", .fields("DistAduanero"))
              Print #NumFile, CampoXML("anio", .fields("Anio"))
              Print #NumFile, CampoXML("regimen", .fields("Regimen"))
              Print #NumFile, CampoXML("correlativo", .fields("Correlativo"))
              Print #NumFile, CampoXML("verificador", .fields("Verificador"))
              Print #NumFile, CampoXML("docTransp", .fields("NumeroDctoTransporte"))
          End If
          Print #NumFile, CampoXML("fechaEmbarque", .fields("FechaEmbarque"))
          Print #NumFile, CampoXML("fue", "000001")                                       '*
          Print #NumFile, CampoXML("valorFOB", Format(.fields("ValorFOB"), "#0.00"))
          Print #NumFile, CampoXML("valorFOBComprobante", Format(.fields("ValorFOBComprobante")))
          Print #NumFile, CampoXML("establecimiento", .fields("Establecimiento"))
          Print #NumFile, CampoXML("puntoEmision", .fields("PuntoEmision"))
          Print #NumFile, CampoXML("secuencial", .fields("Secuencial"))
          Print #NumFile, CampoXML("autorizacion", Format(Val(.fields("Autorizacion")), "#0"))
          Print #NumFile, CampoXML("fechaEmision", .fields("FechaEmision"))
          Print #NumFile, CerrarXML("detalleExportaciones")
         .MoveNext
       Loop
       Print #NumFile, CerrarXML("exportaciones")
   End If
  End With
 'ANULADOS
   With AdoAnulados.Recordset
    If .RecordCount > 0 Then
        Print #NumFile, AbrirXML("anulados")
        Do While Not .EOF
           Print #NumFile, AbrirXML("detalleAnulados")
           Print #NumFile, CampoXML("tipoComprobante", Format(.fields("TipoComprobante"), "00"))
           Print #NumFile, CampoXML("establecimiento", .fields("Establecimiento"))
           Print #NumFile, CampoXML("puntoEmision", .fields("PuntoEmision"))
           Print #NumFile, CampoXML("secuencialInicio", .fields("Secuencial1"))
           Print #NumFile, CampoXML("secuencialFin", .fields("Secuencial2"))
           Print #NumFile, CampoXML("autorizacion", .fields("Autorizacion"))
           Print #NumFile, CerrarXML("detalleAnulados")
          .MoveNext
        Loop
        Print #NumFile, CerrarXML("anulados")
    End If
   End With
  Print #NumFile, CerrarXML("iva")
  Close NumFile   ' Cierra los 3 archivos abiertos.
  RatonNormal
  Archivo_XML = NombreArchivo
End Sub

'''Public Sub Crear_Anexos_Tipo_01_Nuevo(NombreArchivo As String)
'''Dim NumFile As Long
'''Dim Cargar_Air_AT As Boolean
'''Dim RazonSocial As String
'''Dim Numero_Establecimientos As Integer
'''Dim Total_Ventas(1 To 999) As Currency
'''Dim Total_NC(1 To 999) As Currency
'''Dim Total_Ventas_Estab As Currency
'''Dim Total_NC_Estab As Currency
'''Dim Suma_Ventas As Currency
'''
'''  Total_Ventas_Estab = 0
'''  Total_NC_Estab = 0
'''  Numero_Establecimientos = 0
'''  For IE = 1 To 999
'''      Total_Ventas(IE) = 0
'''      Total_NC(IE) = 0
'''  Next IE
'''  RazonSocial = Replace(NombreComercial, ".", " ")
'''  RazonSocial = UCaseStrg(TrimStrg(Replace(RazonSocial, " ", "")))
''' 'Totalizamos de Ventas
'''  With AdoPuntosVentas.Recordset
'''   If .RecordCount > 0 Then
'''       Do While Not .EOF
'''          KE = CInt(.Fields("Establecimiento"))
'''          Suma_Ventas = .Fields("BaseImponibleV") + .Fields("BaseImpGravV") ' + .Fields("MontoIvaV")
'''          If .Fields("TipoComprobante") = 4 Then
'''              Total_NC_Estab = Total_NC_Estab + Suma_Ventas
'''              Total_NC(KE) = Total_NC(KE) + Suma_Ventas
'''          Else
'''              Total_Ventas(KE) = Total_Ventas(KE) + Suma_Ventas
'''              Total_Ventas_Estab = Total_Ventas_Estab + Suma_Ventas
'''          End If
'''         .MoveNext
'''       Loop
'''      .MoveFirst
'''   End If
'''  End With
'''  For IE = 1 To 999
'''      If (Total_Ventas(IE) - Total_NC(IE)) <> 0 Then Numero_Establecimientos = Numero_Establecimientos + 1
'''  Next IE
'''''  MsgBox Numero_Establecimientos
''' 'Numero libre en el disco para grabar
'''  NumFile = FreeFile
''' 'Abrimos el archivo para grabar
'''  Open NombreArchivo For Output As #NumFile   ' Abre el archivo
'''
''' 'ENCABEZADO
'''  Print #NumFile, "<?xml version=""1.0"" encoding=""UTF-8""?>"
'''  Print #NumFile, AbrirXML("iva")
'''  Print #NumFile, vbTab & CampoXML("TipoIDInformante", "R")
'''  Print #NumFile, vbTab & CampoXML("IdInformante", RUC)
'''  Print #NumFile, vbTab & CampoXML("razonSocial", RazonSocial)
'''  Print #NumFile, vbTab & CampoXML("Anio", Format(Anio, "0000"))
'''  Print #NumFile, vbTab & CampoXML("Mes", Format(NumMeses, "00"))
'''  If (Total_Ventas_Estab - Total_NC_Estab) <> 0 Then
'''     Print #NumFile, vbTab & CampoXML("numEstabRuc", Format(Numero_Establecimientos, "000"))
'''  End If
'''  Print #NumFile, vbTab & CampoXML("totalVentas", Format(Total_Ventas_Estab - Total_NC_Estab, "#0.00"))
'''  Print #NumFile, vbTab & CampoXML("codigoOperativo", "IVA")
'''
''' 'COMPRAS
'''  With AdoCompras.Recordset
'''   If .RecordCount > 0 Then
'''       Print #NumFile, AbrirXML("compras")
'''       Do While Not .EOF
'''          TipoDoc = "03"
'''          Select Case .Fields("TD")
'''             Case "R": TipoDoc = "01"
'''             Case "C": TipoDoc = "02"
'''             Case "P": TipoDoc = "03"
'''          End Select
'''         'Aqui genero el Detalle de Compras
'''          Print #NumFile, AbrirXML("detalleCompras")
'''          Print #NumFile, CampoXML("codSustento", .Fields("CodSustento"))
'''          Print #NumFile, CampoXML("tpIdProv", TipoDoc)
'''          Print #NumFile, CampoXML("idProv", .Fields("CI_RUC"))
'''          Print #NumFile, CampoXML("tipoComprobante", Format(.Fields("TipoComprobante"), "00"))
'''         'Si es Pasaporte genera la parte relacionada
'''          If .Fields("TD") = "P" Then
'''             Print #NumFile, CampoXML("tipoProv", .Fields("Tipo_Pasaporte"))
'''             Print #NumFile, CampoXML("parteRel", .Fields("Parte_Relacionada"))
'''          End If
'''          Print #NumFile, CampoXML("fechaRegistro", .Fields("FechaRegistro"))
'''          Print #NumFile, CampoXML("establecimiento", .Fields("Establecimiento"))
'''          Print #NumFile, CampoXML("puntoEmision", .Fields("PuntoEmision"))
'''          Print #NumFile, CampoXML("secuencial", .Fields("Secuencial"))
'''          Print #NumFile, CampoXML("fechaEmision", .Fields("FechaEmision"))
'''          Print #NumFile, CampoXML("autorizacion", .Fields("Autorizacion"))
'''          Print #NumFile, CampoXML("baseNoGraIva", Format(.Fields("BaseNoObjIVA"), "#0.00"))         ' hay que aumentar
'''          Print #NumFile, CampoXML("baseImponible", Format(.Fields("BaseImponible"), "#0.00"))
'''          Print #NumFile, CampoXML("baseImpGrav", Format(.Fields("BaseImpGrav"), "#0.00"))
'''          Print #NumFile, CampoXML("montoIce", Format(.Fields("MontoIce"), "#0.00"))
'''          Print #NumFile, CampoXML("montoIva", Format(.Fields("MontoIva"), "#0.00"))
'''          Print #NumFile, CampoXML("valorRetBienes", Format(.Fields("valorRetBienes"), "#0.00"))
'''          If .Fields("ValorRetServicios") = .Fields("MontoIva") Then
'''              Print #NumFile, CampoXML("valorRetServicios", Format(0, "#0.00"))
'''              Print #NumFile, CampoXML("valRetServ100", Format(.Fields("ValorRetServicios"), "#0.00"))    'hay que aumentar
'''          Else
'''              Print #NumFile, CampoXML("valorRetServicios", Format(.Fields("ValorRetServicios"), "#0.00"))
'''              Print #NumFile, CampoXML("valRetServ100", Format(0, "#0.00"))    'hay que aumentar
'''          End If
'''          Print #NumFile, AbrirXML("pagoExterior")
'''                Print #NumFile, CampoXML("pagoLocExt", .Fields("PagoLocExt"))
'''                If IsNumeric(.Fields("PaisEfecPago")) Then
'''                   Print #NumFile, CampoXML("paisEfecPago", Val(.Fields("PaisEfecPago")))
'''                Else
'''                   Print #NumFile, CampoXML("paisEfecPago", .Fields("PaisEfecPago"))
'''                End If
'''                Print #NumFile, CampoXML("aplicConvDobTrib", .Fields("AplicConvDobTrib"))
'''                Print #NumFile, CampoXML("pagExtSujRetNorLeg", .Fields("PagExtSujRetNorLeg"))
'''          Print #NumFile, CerrarXML("pagoExterior")
'''
'''          If (.Fields("BaseNoObjIVA") + .Fields("BaseImponible") + .Fields("BaseImpGrav") + .Fields("MontoIva")) >= 1000 Then
'''             Print #NumFile, AbrirXML("formasDePago")
'''                   Print #NumFile, CampoXML("formaPago", .Fields("FormaPago"))
'''             Print #NumFile, CerrarXML("formasDePago")
'''          End If
'''         'Asigno a variables los campos para el concepto Air
'''          CodigoCliente = .Fields("IdProv")
'''          Factura_No = .Fields("Secuencial")
'''          Codigo1 = .Fields("Establecimiento")
'''          Codigo2 = .Fields("PuntoEmision")
'''          Estab = "000"
'''          PtoRet = "000"
'''          FechaRet = .Fields("FechaRegistro")
'''          AutRet = "0000000000"
'''          Number = False
'''         'Genero el AIR
'''          sSQL = "SELECT * " _
'''               & "FROM Trans_Air " _
'''               & "WHERE Fecha Between #" & FechaIni & "# AND #" & FechaFin & "# " _
'''               & "AND Periodo = '" & Periodo_Contable & "' "
'''          If ConSucursal Then
'''             If Len(No_ATS) > 3 Then sSQL = sSQL & "AND Item NOT IN (" & No_ATS & ") "
'''          Else
'''             sSQL = sSQL & "AND Item = '" & NumEmpresa & "' "
'''          End If
'''          sSQL = sSQL & "AND T = '" & Normal & "' " _
'''               & "AND Tipo_Trans = 'C' " _
'''               & "AND IdProv = '" & CodigoCliente & "' " _
'''               & "AND Factura_No = " & Factura_No & " " _
'''               & "AND EstabFactura = '" & Codigo1 & "' " _
'''               & "AND PuntoEmiFactura = '" & Codigo2 & "' " _
'''               & "ORDER BY SecRetencion,EstabFactura,PuntoEmiFactura,Factura_No "
'''          Select_Adodc AdoAir, sSQL
'''          If AdoAir.Recordset.RecordCount > 0 Then
'''             Print #NumFile, AbrirXML("air")
'''             Do While Not AdoAir.Recordset.EOF
'''               'Asigno en variables los campos
'''                Codigo = AdoAir.Recordset.Fields("CodRet")
'''                Total_SubTotal = AdoAir.Recordset.Fields("BaseImp")
'''                Porcentaje = AdoAir.Recordset.Fields("Porcentaje")
'''                Valor = AdoAir.Recordset.Fields("ValRet")
'''                Estab = AdoAir.Recordset.Fields("EstabRetencion")
'''                PtoRet = AdoAir.Recordset.Fields("PtoEmiRetencion")
'''                SecRet = AdoAir.Recordset.Fields("SecRetencion")
'''                AutRet = AdoAir.Recordset.Fields("AutRetencion")
'''                If Len(AutRet) < 3 Then AutRet = "0000000000"
'''                Number = True
'''                Print #NumFile, AbrirXML("detalleAir")
'''                Print #NumFile, CampoXML("codRetAir", Codigo)
'''                Print #NumFile, CampoXML("baseImpAir", Format(Total_SubTotal, "##0.00"))
'''                Print #NumFile, CampoXML("porcentajeAir", Format(Porcentaje * 100, "#0.00"))
'''                Print #NumFile, CampoXML("valRetAir", Format(Valor, "##0.00"))
'''                Print #NumFile, CerrarXML("detalleAir")
'''                AdoAir.Recordset.MoveNext
'''             Loop
'''             Print #NumFile, CerrarXML("air")
'''            'Aqui se genera la segunda parte del detalle Compras
'''             Print #NumFile, CampoXML("estabRetencion1", Estab)
'''             Print #NumFile, CampoXML("ptoEmiRetencion1", PtoRet)
'''             If Number = False Then
'''                Print #NumFile, CampoXML("secRetencion1", 0)
'''             Else
'''                Print #NumFile, CampoXML("secRetencion1", Format(SecRet, "#0"))
'''             End If
'''             Print #NumFile, CampoXML("autRetencion1", AutRet)
'''             Print #NumFile, CampoXML("fechaEmiRet1", FechaRet)
'''          End If
'''         'Aqui colocamos si existe una nota de credito o debito
'''          Select Case Val(.Fields("TipoComprobante"))
'''            Case 4, 41
'''                 Print #NumFile, CampoXML("docModificado", Format(Val(.Fields("DocModificado")), "00"))
'''                 Print #NumFile, CampoXML("estabModificado", .Fields("EstabModificado"))
'''                 Print #NumFile, CampoXML("ptoEmiModificado", .Fields("PtoEmiModificado"))
'''                 Print #NumFile, CampoXML("secModificado", .Fields("SecModificado"))
'''                 Print #NumFile, CampoXML("autModificado", .Fields("AutModificado"))
'''          End Select
'''          Print #NumFile, CerrarXML("detalleCompras")
'''         .MoveNext
'''       Loop
'''       Print #NumFile, CerrarXML("compras")
'''   Else
'''       Print #NumFile, VacioXML("compras")
'''   End If
'''  End With
'''
''' 'VENTAS
'''  With AdoVentas.Recordset
'''   If .RecordCount > 0 Then
'''       Print #NumFile, AbrirXML("ventas")
'''       Do While Not .EOF
'''          Contador = .Fields("NumeroComprobantesV")
'''          If Contador <= 0 Then Contador = 1
'''          Factura_No = Contador
'''          CodigoCliente = .Fields("RUC_CI")
'''
'''          'Codigo1 = .Fields("Establecimiento") '"001" '
'''          'Codigo2 = .Fields("PuntoEmision") ' "001" '
'''          Estab = "000"
'''          PtoRet = "000"
'''          'FechaRet = .Fields("FechaRegistro")
'''          FechaRet = FechaFinal
'''          AutRet = "0000000000"
'''          TipoDoc = "06"
'''          Select Case .Fields("TB")
'''            Case "R": TipoDoc = "04"
'''            Case "C": TipoDoc = "05"
'''            Case "P": TipoDoc = "06"
'''            Case "F": TipoDoc = "07"
'''          End Select
'''          If .Fields("RUC_CI") = "9999999999999" Then TipoDoc = "07"
'''         'Genero el talón de Ventas
'''          'If .Fields("RUC_CI") = Ninguno Then MsgBox .Fields("RUC_CI")
'''          Print #NumFile, AbrirXML("detalleVentas")
'''          Print #NumFile, CampoXML("tpIdCliente", TipoDoc)
'''          Print #NumFile, CampoXML("idCliente", .Fields("RUC_CI"))
'''          Print #NumFile, CampoXML("tipoComprobante", Format(.Fields("TipoComprobante"), "00"))
'''          Print #NumFile, CampoXML("numeroComprobantes", Contador)
'''          Print #NumFile, CampoXML("baseNoGraIva", Format(0, "#0.00"))         ' hay que aumentar
'''          Print #NumFile, CampoXML("baseImponible", Format(.Fields("BaseImponibleV"), "#0.00"))
'''          Print #NumFile, CampoXML("baseImpGrav", Format(.Fields("BaseImpGravV"), "#0.00"))
'''          Print #NumFile, CampoXML("montoIva", Format(.Fields("MontoIvaV"), "#0.00"))
'''          Print #NumFile, CampoXML("valorRetIva", Format(.Fields("ValorRetBienesV") + .Fields("ValorRetServiciosV"), "#0.00"))
'''         'Asigno a variables los campos
'''         'Genero el AIR
'''          Cargar_Air_AT = True
'''          Select Case Val(.Fields("TipoComprobante"))
'''            Case 4, 41: Cargar_Air_AT = False   ' NC o ND
'''          End Select
'''          SubTotal = 0
'''          If Cargar_Air_AT Then
'''             sSQL = "SELECT IdProv,SUM(ValRet) As TValRet " _
'''                  & "FROM Trans_Air " _
'''                  & "WHERE Fecha Between #" & FechaIni & "# AND #" & FechaFin & "#  " _
'''                  & "AND Periodo = '" & Periodo_Contable & "' "
'''             If ConSucursal Then
'''                If Len(No_ATS) > 3 Then sSQL = sSQL & "AND Item NOT IN (" & No_ATS & ") "
'''             Else
'''                sSQL = sSQL & "AND Item = '" & NumEmpresa & "' "
'''             End If
'''             sSQL = sSQL _
'''                  & "AND T = '" & Normal & "' " _
'''                  & "AND Tipo_Trans = 'V' " _
'''                  & "AND RUC_CI = '" & CodigoCliente & "' " _
'''                  & "AND EstabFactura = '" & Codigo1 & "' " _
'''                  & "AND PuntoEmiFactura = '" & Codigo2 & "' " _
'''                  & "GROUP BY IdProv "
'''                  '& "ORDER BY SecRetencion,EstabFactura,PuntoEmiFactura,Factura_No "
'''            'If CodigoCliente = "1790892875" Then MsgBox sSQL
'''             Select_Adodc AdoAir, sSQL
'''             If AdoAir.Recordset.RecordCount > 0 Then
''''                Do While Not AdoAir.Recordset.EOF
'''                   SubTotal = SubTotal + AdoAir.Recordset.Fields("TValRet")
'''                   'If CodigoCliente = "1790892875" Then MsgBox SubTotal
''''                   AdoAir.Recordset.MoveNext
''''                Loop
'''             End If
'''          End If
'''         'MsgBox SubTotal
'''          Print #NumFile, CampoXML("valorRetRenta", Format(SubTotal, "#0.00"))
'''          Print #NumFile, CerrarXML("detalleVentas")
'''         .MoveNext
'''       Loop
'''       Print #NumFile, CerrarXML("ventas")
'''   End If
'''  End With
''' 'Resumen Total de Ventas por Establecimientos
'''  If (Total_Ventas_Estab + Total_NC_Estab) > 0 Then
'''     Print #NumFile, AbrirXML("ventasEstablecimiento")
'''     For KE = 1 To 999
'''      If Total_Ventas(KE) + Total_NC(KE) > 0 Then
'''         Print #NumFile, AbrirXML("ventaEst")
'''         Print #NumFile, CampoXML("codEstab", Format(KE, "000"))
'''         Print #NumFile, CampoXML("ventasEstab", Format(Total_Ventas(KE) - Total_NC(KE), "#0.00"))
'''         Print #NumFile, CerrarXML("ventaEst")
'''      End If
'''     Next KE
'''     Print #NumFile, CerrarXML("ventasEstablecimiento")
'''  End If
''' 'EXPORTACIONES
'''  With AdoExportaciones.Recordset
'''   If .RecordCount > 0 Then
'''       Print #NumFile, AbrirXML("exportaciones")
'''       Do While Not .EOF
'''          TipoDoc = "1"
'''          Select Case .Fields("TD")
'''            Case "R", "C": TipoDoc = "2"
'''          End Select
'''         'Genero el talón de Exportaciones
'''          Print #NumFile, AbrirXML("detalleExportaciones")
'''          Print #NumFile, CampoXML("exportacionDe", .Fields("ExportacionDe"))
'''          Print #NumFile, CampoXML("tipoComprobante", Format(.Fields("TipoComprobante"), "00"))
'''          If .Fields("ExportacionDe") = 1 Then
'''              Print #NumFile, CampoXML("distAduanero", .Fields("DistAduanero"))
'''              Print #NumFile, CampoXML("anio", .Fields("Anio"))
'''              Print #NumFile, CampoXML("regimen", .Fields("Regimen"))
'''              Print #NumFile, CampoXML("correlativo", .Fields("Correlativo"))
'''              Print #NumFile, CampoXML("verificador", .Fields("Verificador"))
'''              Print #NumFile, CampoXML("docTransp", .Fields("NumeroDctoTransporte"))
'''          End If
'''          Print #NumFile, CampoXML("fechaEmbarque", .Fields("FechaEmbarque"))
'''          Print #NumFile, CampoXML("fue", "000001")                                       '*
'''          Print #NumFile, CampoXML("valorFOB", Format(.Fields("ValorFOB"), "#0.00"))
'''          Print #NumFile, CampoXML("valorFOBComprobante", Format(.Fields("ValorFOBComprobante")))
'''          Print #NumFile, CampoXML("establecimiento", .Fields("Establecimiento"))
'''          Print #NumFile, CampoXML("puntoEmision", .Fields("PuntoEmision"))
'''          Print #NumFile, CampoXML("secuencial", .Fields("Secuencial"))
'''          Print #NumFile, CampoXML("autorizacion", Format(Val(.Fields("Autorizacion")), "#0"))
'''          Print #NumFile, CampoXML("fechaEmision", .Fields("FechaEmision"))
'''          Print #NumFile, CerrarXML("detalleExportaciones")
'''         .MoveNext
'''       Loop
'''       Print #NumFile, CerrarXML("exportaciones")
'''   End If
'''  End With
''' 'ANULADOS
'''   With AdoAnulados.Recordset
'''    If .RecordCount > 0 Then
'''        Print #NumFile, AbrirXML("anulados")
'''        Do While Not .EOF
'''           Print #NumFile, AbrirXML("detalleAnulados")
'''           Print #NumFile, CampoXML("tipoComprobante", Format(.Fields("TipoComprobante"), "00"))
'''           Print #NumFile, CampoXML("establecimiento", .Fields("Establecimiento"))
'''           Print #NumFile, CampoXML("puntoEmision", .Fields("PuntoEmision"))
'''           Print #NumFile, CampoXML("secuencialInicio", .Fields("Secuencial1"))
'''           Print #NumFile, CampoXML("secuencialFin", .Fields("Secuencial2"))
'''           Print #NumFile, CampoXML("autorizacion", .Fields("Autorizacion"))
'''           Print #NumFile, CerrarXML("detalleAnulados")
'''          .MoveNext
'''        Loop
'''        Print #NumFile, CerrarXML("anulados")
'''    End If
'''   End With
'''  Print #NumFile, CerrarXML("iva")
'''  Close NumFile   ' Cierra los 3 archivos abiertos.
'''  Archivo_XML = NombreArchivo
'''End Sub

Public Sub Crear_Anexos_Tipo_02(NombreArchivo As String)
  Dim NumFile As Long
  NumFile = FreeFile
  Open NombreArchivo For Output As #NumFile   ' Abre el archivo
 'ENCABEZADO
  Print #NumFile, "<?xml version='1.0' encoding='UTF-8'?>"
  Print #NumFile, "<iva xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance'>"
  Print #NumFile, CampoXML("numeroRuc", RUC)
  Print #NumFile, CampoXML("razonSocial", UCaseStrg(NombreComercial))
  Print #NumFile, CampoXML("direccionMatriz", Direccion)
  Print #NumFile, CampoXML("telefono", Telefono1)
  Print #NumFile, CampoXML("email", EmailEmpresa)
  Print #NumFile, CampoXML("tpIdRepre", TID_Repres)
  Print #NumFile, CampoXML("idRepre", CI_Representante)
  Print #NumFile, CampoXML("rucContador", RUC_Contador)
  Print #NumFile, CampoXML("anio", Format(Anio, "0000"))
  Print #NumFile, CampoXML("mes", Format(NumMeses, "00"))
 'COMPRAS
  With AdoCompras.Recordset
   If .RecordCount > 0 Then
       Print #NumFile, AbrirXML("compras")
       Do While Not .EOF
          TipoDoc = "03"
          Select Case .fields("TD")
             Case "R": TipoDoc = "01"
             Case "C": TipoDoc = "02"
             Case "P": TipoDoc = "03"
          End Select
         'Aqui genero el Detalle de Compras
          Print #NumFile, AbrirXML("detalleCompras")
          Print #NumFile, CampoXML("codSustento", .fields("CodSustento"))
          Print #NumFile, CampoXML("devIva", .fields("DevIva"))
          Print #NumFile, CampoXML("tpIdProv", TipoDoc)
          Print #NumFile, CampoXML("idProv", .fields("CI_RUC"))
          Print #NumFile, CampoXML("tipoComprobante", .fields("TipoComprobante"))
          Print #NumFile, CampoXML("fechaRegistro", .fields("FechaRegistro"))
          Print #NumFile, CampoXML("establecimiento", .fields("Establecimiento"))
          Print #NumFile, CampoXML("puntoEmision", .fields("PuntoEmision"))
          Print #NumFile, CampoXML("secuencial", .fields("Secuencial"))
          Print #NumFile, CampoXML("fechaEmision", .fields("FechaEmision"))
          Print #NumFile, CampoXML("autorizacion", Format(Val(.fields("Autorizacion")), "0000000000"))
          Print #NumFile, CampoXML("fechaCaducidad", Mes_Año(.fields("FechaCaducidad")))
          Print #NumFile, CampoXML("baseImponible", Format(.fields("BaseImponible")))
          Print #NumFile, CampoXML("baseImpGrav", Format(.fields("BaseImpGrav"), "#0.00"))
          Print #NumFile, CampoXML("porcentajeIva", .fields("PorcentajeIva"))
          Print #NumFile, CampoXML("montoIva", Format(.fields("MontoIva"), "#0.00"))
          Print #NumFile, CampoXML("baseImpIce", Format(.fields("BaseImpIce"), "#0.00"))
          Print #NumFile, CampoXML("porcentajeIce", Format(.fields("PorcentajeIce"), "#0"))
          Print #NumFile, CampoXML("montoIce", Format(.fields("MontoIce"), "#0.00"))
          Print #NumFile, CampoXML("montoIvaBienes", Format(.fields("MontoIvaBienes"), "#0.00"))
          Print #NumFile, CampoXML("porRetBienes", .fields("PorRetBienes"))
          Print #NumFile, CampoXML("valorRetBienes", Format(.fields("valorRetBienes"), "#0.00"))
          Print #NumFile, CampoXML("montoIvaServicios", Format(.fields("MontoIvaServicios"), "#0.00"))
          Print #NumFile, CampoXML("porRetServicios", .fields("PorRetServicios"))
          Print #NumFile, CampoXML("valorRetServicios", Format(.fields("ValorRetServicios"), "#0.00"))
         'Asigno a variables los campos
          CodigoCliente = .fields("IdProv")
          Factura_No = .fields("Secuencial")
          Codigo1 = .fields("Establecimiento")
          Codigo2 = .fields("PuntoEmision")
          Estab = "000"
          PtoRet = "000"
          FechaRet = .fields("FechaRegistro")
          AutRet = "0000000000"
          Number = False
         'Genero el AIR
          sSQL = "SELECT * " _
               & "FROM Trans_Air " _
               & "WHERE Fecha Between #" & FechaIni & "# AND #" & FechaFin & "# " _
               & "AND Periodo = '" & Periodo_Contable & "' "
  If ConSucursal Then
     If Len(No_ATS) > 3 Then sSQL = sSQL & "AND Item NOT IN (" & No_ATS & ") "
  Else
     sSQL = sSQL & "AND Item = '" & NumEmpresa & "' "
  End If

          sSQL = sSQL & "AND T = '" & Normal & "' " _
               & "AND Tipo_Trans = 'C' " _
               & "AND IdProv = '" & CodigoCliente & "' " _
               & "AND Factura_No = " & Factura_No & " " _
               & "AND EstabFactura = '" & Codigo1 & "' " _
               & "AND PuntoEmiFactura = '" & Codigo2 & "' " _
               & "ORDER BY SecRetencion,EstabFactura,PuntoEmiFactura,Factura_No "
          Select_Adodc AdoAir, sSQL
          If AdoAir.Recordset.RecordCount > 0 Then
             Print #NumFile, AbrirXML("air")
             Do While Not AdoAir.Recordset.EOF
               'Asigno en variables los campos
                Codigo = AdoAir.Recordset.fields("CodRet")
                Total_SubTotal = AdoAir.Recordset.fields("BaseImp")
                Porcentaje = AdoAir.Recordset.fields("Porcentaje")
                Valor = AdoAir.Recordset.fields("ValRet")
                Estab = AdoAir.Recordset.fields("EstabRetencion")
                PtoRet = AdoAir.Recordset.fields("PtoEmiRetencion")
                SecRet = AdoAir.Recordset.fields("SecRetencion")
                AutRet = AdoAir.Recordset.fields("AutRetencion")
                Number = True
                Print #NumFile, AbrirXML("detalleAir")
                Print #NumFile, CampoXML("codRetAir", Codigo)
                Print #NumFile, CampoXML("baseImpAir", Format(Total_SubTotal, "##0.00"))
                Print #NumFile, CampoXML("porcentajeAir", Format(Porcentaje * 100, "#0.00"))
                Print #NumFile, CampoXML("valRetAir", Format(Valor, "#0.00"))
                Print #NumFile, CerrarXML("detalleAir")
                AdoAir.Recordset.MoveNext
             Loop
             Print #NumFile, CerrarXML("air")
          Else
             Print #NumFile, VacioXML("air")
          End If
         'Aqui se genera la segunda parte del detalle Compras
          Print #NumFile, CampoXML("estabRetencion1", Estab)
          Print #NumFile, CampoXML("ptoEmiRetencion1", PtoRet)
          If Number = False Then
             Print #NumFile, CampoXML("secRetencion1", 0)
          Else
             Print #NumFile, CampoXML("secRetencion1", Format(SecRet, "#0"))
          End If
          Print #NumFile, CampoXML("autRetencion1", Format(Val(AutRet), "#0"))
          Print #NumFile, CampoXML("fechaEmiRet1", FechaRet)
          Print #NumFile, CampoXML("docModificado", .fields("DocModificado"))
          Print #NumFile, CampoXML("fechaEmiModificado", .fields("FechaEmiModificado"))
          Print #NumFile, CampoXML("estabModificado", .fields("EstabModificado"))
          Print #NumFile, CampoXML("ptoEmiModificado", .fields("PtoEmiModificado"))
          Print #NumFile, CampoXML("secModificado", .fields("SecModificado"))
          Print #NumFile, CampoXML("autModificado", .fields("AutModificado"))
          If .fields("ContratoPartidoPolitico") <> "0000000000" Then
              Print #NumFile, CampoXML("contratoPartidoPolitico", Format(.fields("ContratoPartidoPolitico"), "#0"))
          End If
          Print #NumFile, CampoXML("montoTituloOneroso", Format(.fields("MontoTituloOneroso"), "#0.00"))
          Print #NumFile, CampoXML("montoTituloGratuito", Format(.fields("MontoTituloGratuito"), "#0.00"))
          Print #NumFile, CerrarXML("detalleCompras")
         .MoveNext
       Loop
       Print #NumFile, CerrarXML("compras")
   Else
       Print #NumFile, VacioXML("compras")
   End If
  End With
 'VENTAS
  With AdoVentas.Recordset
   If .RecordCount > 0 Then
       Print #NumFile, AbrirXML("ventas")
       Do While Not .EOF
          TipoDoc = "06"
          Select Case .fields("TD")
            Case "R": TipoDoc = "04"
            Case "C": TipoDoc = "05"
            Case "P": TipoDoc = "06"
            Case "F": TipoDoc = "07"
          End Select
         'Genero el talón de Ventas
         
          Print #NumFile, AbrirXML("detalleVentas")
          If .fields("CI_RUC") = "9999999999999" Then
              Print #NumFile, CampoXML("tpIdCliente", "07")
              Print #NumFile, CampoXML("idCliente", .fields("CI_RUC"))
              Print #NumFile, CampoXML("tipoComprobante", .fields("TipoComprobante"))
          Else
              Print #NumFile, CampoXML("tpIdCliente", TipoDoc)
              Print #NumFile, CampoXML("idCliente", .fields("CI_RUC"))
              Print #NumFile, CampoXML("tipoComprobante", .fields("TipoComprobante"))
          End If
          
          Print #NumFile, CampoXML("fechaRegistro", UltimoDiaMes(.fields("FechaRegistro")))
          Contador = .fields("NumeroComprobantesV")
          If Contador <= 0 Then Contador = 1
          Print #NumFile, CampoXML("numeroComprobantes", Contador)
          Print #NumFile, CampoXML("fechaEmision", UltimoDiaMes(.fields("FechaEmision")))
          Print #NumFile, CampoXML("baseImponible", Format(.fields("BaseImponibleV"), "#0.00"))
          Print #NumFile, CampoXML("ivaPresuntivo", .fields("IvaPresuntivo"))
          Print #NumFile, CampoXML("baseImpGrav", Format(.fields("BaseImpGravV"), "#0.00"))
          Print #NumFile, CampoXML("porcentajeIva", .fields("PorcentajeIva"))
          Print #NumFile, CampoXML("montoIva", Format(.fields("MontoIvaV"), "#0.00"))
          Print #NumFile, CampoXML("baseImpIce", Format(.fields("BaseImpIceV"), "#0.00"))
          Print #NumFile, CampoXML("porcentajeIce", .fields("PorcentajeIce"))
          Print #NumFile, CampoXML("montoIce", Format(.fields("MontoIceV"), "#0.00"))
          Print #NumFile, CampoXML("montoIvaBienes", Format(.fields("MontoIvaBienesV"), "#0.00"))
          Print #NumFile, CampoXML("porRetBienes", .fields("PorRetBienes"))
          Print #NumFile, CampoXML("valorRetBienes", Format(.fields("ValorRetBienesV"), "#0.00"))
          Print #NumFile, CampoXML("montoIvaServicios", Format(.fields("MontoIvaServiciosV"), "#0.00"))
          Print #NumFile, CampoXML("porRetServicios", .fields("PorRetServicios"))
          Print #NumFile, CampoXML("valorRetServicios", Format(.fields("ValorRetServiciosV"), "#0.00"))
          Print #NumFile, CampoXML("retPresuntiva", .fields("RetPresuntiva"))
          
          'Asigno a variables los campos
          CodigoCliente = .fields("IdProv")
          Factura_No = .fields("NumeroComprobantes")
          Codigo1 = .fields("Establecimiento")
          Codigo2 = .fields("PuntoEmision")
          Estab = "000"
          PtoRet = "000"
          FechaRet = .fields("FechaRegistro")
          AutRet = "0000000000"
          
          'Genero el AIR
          sSQL = "SELECT * " _
               & "FROM Trans_Air " _
               & "WHERE Fecha Between #" & FechaIni & "# AND #" & FechaFin & "#  " _
               & "AND Periodo = '" & Periodo_Contable & "' "
  If ConSucursal Then
     If Len(No_ATS) > 3 Then sSQL = sSQL & "AND Item NOT IN (" & No_ATS & ") "
  Else
     sSQL = sSQL & "AND Item = '" & NumEmpresa & "' "
  End If

          sSQL = sSQL & "AND T = '" & Normal & "' " _
               & "AND Tipo_Trans = 'V' " _
               & "AND IdProv = '" & CodigoCliente & "' " _
               & "AND Factura_No = " & Factura_No & " " _
               & "AND EstabFactura = '" & Codigo1 & "' " _
               & "AND PuntoEmiFactura = '" & Codigo2 & "' " _
               & "ORDER BY SecRetencion,EstabFactura,PuntoEmiFactura,Factura_No "
          Select_Adodc AdoAir, sSQL
          If AdoAir.Recordset.RecordCount > 0 Then
             Print #NumFile, AbrirXML("air")
             Do While Not AdoAir.Recordset.EOF
                Codigo = AdoAir.Recordset.fields("CodRet")
                Total_SubTotal = AdoAir.Recordset.fields("BaseImp")
                Porcentaje = AdoAir.Recordset.fields("Porcentaje")
                Valor = AdoAir.Recordset.fields("ValRet")
                Estab = AdoAir.Recordset.fields("EstabRetencion")
                PtoRet = AdoAir.Recordset.fields("PtoEmiRetencion")
                SecRet = AdoAir.Recordset.fields("SecRetencion")
                AutRet = AdoAir.Recordset.fields("AutRetencion")
                
                Print #NumFile, AbrirXML("detalleAir")
                Print #NumFile, CampoXML("codRetAir", Codigo)
                Print #NumFile, CampoXML("baseImpAir", Format(Total_SubTotal, "##0.00"))
                Print #NumFile, CampoXML("porcentajeAir", Format(Porcentaje * 100, "#0.00"))
                Print #NumFile, CampoXML("valRetAir", Format(Valor, "#0.00"))
                Print #NumFile, CerrarXML("detalleAir")
                AdoAir.Recordset.MoveNext
             Loop
             Print #NumFile, CerrarXML("air")
          Else
             Print #NumFile, VacioXML("air")
          End If
        
          Print #NumFile, CerrarXML("detalleVentas")
         .MoveNext
       Loop
       Print #NumFile, CerrarXML("ventas")
   Else
       Print #NumFile, VacioXML("ventas")
   End If
  End With
 'IMPORTACIONES
  With AdoImportaciones.Recordset
   If .RecordCount > 0 Then
       Print #NumFile, AbrirXML("importaciones")
       Do While Not .EOF
          TipoDoc = "1"
          If .fields("TD") = "R" Then TipoDoc = "2"
          'Genero el talón de Importaciones
          Print #NumFile, AbrirXML("detalleImportaciones")
          Print #NumFile, CampoXML("codSustento", .fields("CodSustento"))
          Print #NumFile, CampoXML("importacionDe", .fields("ImportacionDe"))
          Print #NumFile, CampoXML("fechaLiquidacion", .fields("FechaLiquidacion"))
          Print #NumFile, CampoXML("tipoComprobante", .fields("TipoComprobante"))
          Print #NumFile, CampoXML("distAduanero", .fields("DistAduanero"))
          Print #NumFile, CampoXML("anio", .fields("Anio"))
          Print #NumFile, CampoXML("regimen", .fields("Regimen"))
          Print #NumFile, CampoXML("correlativo", .fields("Correlativo"))
          Print #NumFile, CampoXML("verificador", .fields("Verificador"))
          Print #NumFile, CampoXML("idFiscalProv", .fields("CI_RUC"))
          Print #NumFile, CampoXML("valorCIF", Format(.fields("ValorCIF"), "#0.00"))
          Print #NumFile, CampoXML("razonSocialProv", UCaseStrg(NombreComercial))
          Print #NumFile, CampoXML("tipoSujeto", TipoDoc)
          Print #NumFile, CampoXML("baseImponible", Format(.fields("BaseImponible"), "#0.00"))
          Print #NumFile, CampoXML("baseImpGrav", Format(.fields("BaseImpGrav"), "#0.00"))
          Print #NumFile, CampoXML("porcentajeIva", .fields("PorcentajeIva"))
          Print #NumFile, CampoXML("montoIva", Format(.fields("MontoIva"), "#0.00"))
          Print #NumFile, CampoXML("baseImpIce", Format(.fields("BaseImpIce"), "#0.00"))
          Print #NumFile, CampoXML("porcentajeIce", .fields("PorcentajeIce"))
          Print #NumFile, CampoXML("montoIce", Format(.fields("MontoIce"), "#0.00"))
                 
          'Asigno a variables los campos
          CodigoCliente = .fields("IdFiscalProv")
          Factura_No = .fields("Correlativo")
          Estab = "000"
          PtoRet = "000"
          FechaRet = "00/00/0000"
          AutRet = "0000000000"
          
          'Genero el AIR
          sSQL = "SELECT * " _
               & "FROM Trans_Air " _
               & "WHERE Fecha Between #" & FechaIni & "# AND #" & FechaFin & "#  " _
               & "AND Periodo = '" & Periodo_Contable & "' "
  If ConSucursal Then
     If Len(No_ATS) > 3 Then sSQL = sSQL & "AND Item NOT IN (" & No_ATS & ") "
  Else
     sSQL = sSQL & "AND Item = '" & NumEmpresa & "' "
  End If

          sSQL = sSQL & "AND T = '" & Normal & "' " _
               & "AND Tipo_Trans = 'I' " _
               & "AND IdProv = '" & CodigoCliente & "' " _
               & "AND Factura_No = " & Factura_No & " " _
               & "ORDER BY SecRetencion,EstabFactura,PuntoEmiFactura,Factura_No "
          Select_Adodc AdoAir, sSQL
          If AdoAir.Recordset.RecordCount > 0 Then
             Print #NumFile, AbrirXML("air")
             Do While Not AdoAir.Recordset.EOF
                'Asigno en variables los campos
                Codigo = AdoAir.Recordset.fields("CodRet")
                Total_SubTotal = AdoAir.Recordset.fields("BaseImp")
                Porcentaje = AdoAir.Recordset.fields("Porcentaje")
                Valor = AdoAir.Recordset.fields("ValRet")
                Estab = AdoAir.Recordset.fields("EstabRetencion")
                PtoRet = AdoAir.Recordset.fields("PtoEmiRetencion")
                SecRet = AdoAir.Recordset.fields("SecRetencion")
                AutRet = AdoAir.Recordset.fields("AutRetencion")
                FechaRet = AdoAir.Recordset.fields("Fecha")
                
                Print #NumFile, AbrirXML("detalleAir")
                Print #NumFile, CampoXML("codRetAir", Codigo)
                Print #NumFile, CampoXML("baseImpAir", Format(Total_SubTotal, "##0.00"))
                Print #NumFile, CampoXML("porcentajeAir", Format(Porcentaje * 100, "#0.00"))
                Print #NumFile, CampoXML("valRetAir", Format(Valor, "#0.00"))
                Print #NumFile, CerrarXML("detalleAir")
                AdoAir.Recordset.MoveNext
             Loop
             Print #NumFile, CerrarXML("air")
          Else
             Print #NumFile, VacioXML("air")
          End If
          
          Print #NumFile, CerrarXML("detalleImportaciones")
         .MoveNext
       Loop
       Print #NumFile, CerrarXML("importaciones")
   Else
       Print #NumFile, VacioXML("importaciones")
   End If
  End With
 'EXPORTACIONES
  With AdoExportaciones.Recordset
   If .RecordCount > 0 Then
       Print #NumFile, AbrirXML("exportaciones")
       Do While Not .EOF
          TipoDoc = "1"
          Select Case .fields("TD")
            Case "R", "C": TipoDoc = "2"
          End Select
          'Genero el talón de Exportaciones
          Print #NumFile, AbrirXML("detalleExportaciones")
          Print #NumFile, CampoXML("exportacionDe", .fields("ExportacionDe"))
          Print #NumFile, CampoXML("tipoComprobante", .fields("TipoComprobante"))
          If .fields("ExportacionDe") = 1 Then
              Print #NumFile, CampoXML("distAduanero", .fields("DistAduanero"))
              Print #NumFile, CampoXML("anio", .fields("Anio"))
              Print #NumFile, CampoXML("regimen", .fields("Regimen"))
              Print #NumFile, CampoXML("correlativo", .fields("Correlativo"))
              Print #NumFile, CampoXML("verificador", .fields("Verificador"))
              Print #NumFile, CampoXML("docTransp", .fields("NumeroDctoTransporte"))
          End If
          Print #NumFile, CampoXML("fechaEmbarque", .fields("FechaEmbarque"))
          Print #NumFile, CampoXML("idFiscalCliente", .fields("CI_RUC"))
          Print #NumFile, CampoXML("tipoSujeto", TipoDoc)
          Print #NumFile, CampoXML("valorFOB", Format(.fields("ValorFOB"), "#0.00"))
          Print #NumFile, CampoXML("razonSocialCliente", UCaseStrg(NombreComercial))
          Print #NumFile, CampoXML("devIva", .fields("DevIva"))
          Print #NumFile, CampoXML("facturaExportacion", .fields("FacturaExportacion"))
          Print #NumFile, CampoXML("valorFOBComprobante", Format(.fields("ValorFOBComprobante")))
          Print #NumFile, CampoXML("establecimiento", .fields("Establecimiento"))
          Print #NumFile, CampoXML("puntoEmision", .fields("PuntoEmision"))
          Print #NumFile, CampoXML("secuencial", .fields("Secuencial"))
          Print #NumFile, CampoXML("fechaRegistro", .fields("FechaRegistro"))
          Print #NumFile, CampoXML("autorizacion", Format(Val(.fields("Autorizacion")), "#0"))
          Print #NumFile, CampoXML("fechaEmision", .fields("FechaEmision"))
          Print #NumFile, CerrarXML("detalleExportaciones")
         .MoveNext
       Loop
       Print #NumFile, CerrarXML("exportaciones")
   Else
       Print #NumFile, VacioXML("exportaciones")
   End If
  End With
'  'RECAP
'  Print #NumFile, AbrirXML("recap")
'  Print #NumFile, AbrirXML("detalleRecap")
'  Print #NumFile, CampoXML("establecimientoRecap")
'  Print #NumFile, CampoXML("identificacionRecap")
'  Print #NumFile, CampoXML("tipoComprobante")
'  Print #NumFile, CampoXML("numero Recap")
'  Print #NumFile, CampoXML("fechaPago")
'  Print #NumFile, CampoXML("tarjetaCredito")
'  Print #NumFile, CampoXML("fechaEmisionRecap")
'  Print #NumFile, CampoXML("consumoCero")
'  Print #NumFile, CampoXML("consumoGravado")
'  Print #NumFile, CampoXML("totalConsumo")
'  Print #NumFile, CampoXML("montoIva")
'  Print #NumFile, CampoXML("comision")
'  Print #NumFile, CampoXML("numeroVouchers")
'  Print #NumFile, CampoXML("montoIvaBienes")
'  Print #NumFile, CampoXML("porRetBienes")
'  Print #NumFile, CampoXML("valorRetBienes")
'  Print #NumFile, CampoXML("montoIvaServicios")
'  Print #NumFile, CampoXML("porRetServicios")
'  Print #NumFile, CampoXML("valorRetServicios")
'  Print #NumFile, AbrirXML("air")
'  Print #NumFile, AbrirXML("detalleAir")
'  Print #NumFile, CampoXML("codRetAir")
'  Print #NumFile, CampoXML("baseImpAir")
'  Print #NumFile, CampoXML("porcentajeAir")
'  Print #NumFile, CampoXML("valRetAir")
'  Print #NumFile, CerrarXML("detalleAir")
'  Print #NumFile, CerrarXML("air")
'  Print #NumFile, CampoXML("establecimiento")
'  Print #NumFile, CampoXML("puntoEmision")
'  Print #NumFile, CampoXML("secuencial")
'  Print #NumFile, CampoXML("fechaRegistro")
'  Print #NumFile, CerrarXML("autorizacion")
'  Print #NumFile, CampoXML("fechaEmision")
'  Print #NumFile, CerrarXML("detalleRecap")
'  Print #NumFile, CerrarXML("recap")
'
'  'FIDEICOMISOS
'  Print #NumFile, AbrirXML("fideicomisos")
'  Print #NumFile, AbrirXML("detalleFideicomisos")
'  Print #NumFile, CampoXML("tipoBeneficiaro")
'  Print #NumFile, CampoXML("idBeneficiario")
'  Print #NumFile, CampoXML("rucFideicomiso")
'  Print #NumFile, AbrirXML("fvalor")
'  Print #NumFile, AbrirXML("detallefValor")
'  Print #NumFile, CampoXML("tipoFideicomiso")
'  Print #NumFile, CampoXML("totalF")
'  Print #NumFile, CampoXML("individual")
'  Print #NumFile, CampoXML("porRetF")
'  Print #NumFile, CampoXML("valorRetF")
'  Print #NumFile, CerrarXML("detallefValor")
'  Print #NumFile, CerrarXML("fValor")
'
'  'ANULADOS
   With AdoAnulados.Recordset
    If .RecordCount > 0 Then
        Print #NumFile, AbrirXML("anulados")
        Do While Not .EOF
           Print #NumFile, AbrirXML("detalleAnulados")
           Print #NumFile, CampoXML("tipoComprobante", .fields("TipoComprobante"))
           Print #NumFile, CampoXML("establecimiento", .fields("Establecimiento"))
           Print #NumFile, CampoXML("puntoEmision", .fields("PuntoEmision"))
           Print #NumFile, CampoXML("secuencialInicio", .fields("Secuencial1"))
           Print #NumFile, CampoXML("secuencialFin", .fields("Secuencial2"))
           Print #NumFile, CampoXML("autorizacion", .fields("Autorizacion"))
           Print #NumFile, CampoXML("fechaAnulacion", .fields("FechaAnulacion"))
           Print #NumFile, CerrarXML("detalleAnulados")
          .MoveNext
        Loop
        Print #NumFile, CerrarXML("anulados")
    Else
       Print #NumFile, VacioXML("anulados")
    End If
   End With
'
'  'RENDIMIENTOS FINANCIEROS
'  Print #NumFile, AbrirXML("rendFinancieros")
'  Print #NumFile, AbrirXML("detalleRendFinancieros")
'  Print #NumFile, CampoXML("retenido")
'  Print #NumFile, CampoXML("idRetenido")
'  Print #NumFile, CampoXML("tipoCompR")
'  Print #NumFile, AbrirXML("airRend")
'  Print #NumFile, AbrirXML("detalleAirRen")
'  Print #NumFile, CampoXML("codRetAir")
'  Print #NumFile, CampoXML("deposito")
'  Print #NumFile, CampoXML("baseImpAir")
'  Print #NumFile, CampoXML("porcentajeAir")
'  Print #NumFile, CampoXML("valRetAir")
'  Print #NumFile, CerrarXML("detalleAirRen")
'  Print #NumFile, CerrarXML("detalleAirRen")
'  Print #NumFile, CampoXML("estabRetencion")
'  Print #NumFile, CampoXML("ptoEmiRetencion")
'  Print #NumFile, CampoXML("secRetencion")
'  Print #NumFile, CampoXML("autRetencion")
'  Print #NumFile, CampoXML("fechaEmiRet")
'  Print #NumFile, CerrarXML("detalleRendFinancieros")
'  Print #NumFile, CerrarXML("rendFinancieros")
  Print #NumFile, CerrarXML("iva")
  Close NumFile   ' Cierra los 3 archivos abiertos.
  Archivo_XML = NombreArchivo
End Sub

Public Sub Crear_Anexos_REOC(NombreArchivo As String)
  Dim NumFile As Long
  If Periodo_Contable = "" Then Periodo_Contable = Ninguno
  For I = 0 To 14
     Totales_SRI(I) = 0
  Next I
  NumFile = FreeFile
  Open NombreArchivo For Output As #NumFile   ' Abre el archivo
 'ENCABEZADO
  Print #NumFile, "<?xml version=";
  Print #NumFile, Chr(34);
  Print #NumFile, "1.0";
  Print #NumFile, Chr(34);
  Print #NumFile, " encoding=";
  Print #NumFile, Chr(34);
  Print #NumFile, "ISO-8859-1";
  Print #NumFile, Chr(34);
  Print #NumFile, " standalone=";
  Print #NumFile, Chr(34);
  Print #NumFile, "yes";
  Print #NumFile, Chr(34);
  Print #NumFile, "?>"
  Print #NumFile, AbrirXML("reoc")
  Print #NumFile, vbTab & CampoXML("numeroRuc", RUC)
  Print #NumFile, vbTab & CampoXML("anio", Format(Anio, "0000"))
  Print #NumFile, vbTab & CampoXML("mes", Format(NumMeses, "00"))
 'COMPRAS
  With AdoCompras.Recordset
   If .RecordCount > 0 Then
       Print #NumFile, vbTab & AbrirXML("compras")
       Do While Not .EOF
         'Asigno a variables los campos
          CodigoCliente = .fields("IdProv")
          Factura_No = .fields("Secuencial")
          Codigo1 = .fields("Establecimiento")
          Codigo2 = .fields("PuntoEmision")
          Saldo = .fields("BaseImpGrav")
          FechaRet = .fields("FechaRegistro")
          Estab = "000"
          PtoRet = "000"
          
          AutRet = "0000000000"
          Number = False
          TipoDoc = "03"
          Select Case .fields("TD")
             Case "R": TipoDoc = "01"
             Case "C": TipoDoc = "02"
             Case "P": TipoDoc = "03"
          End Select
         'Aqui genero el Detalle de Compras
          Print #NumFile, vbTab & vbTab & AbrirXML("detalleCompras")
          Print #NumFile, vbTab & vbTab & vbTab & CampoXML("tpIdProv", TipoDoc)
          Print #NumFile, vbTab & vbTab & vbTab & CampoXML("idProv", .fields("CI_RUC"))
          Print #NumFile, vbTab & vbTab & vbTab & CampoXML("tipoComp", .fields("TipoComprobante"))
          Print #NumFile, vbTab & vbTab & vbTab & CampoXML("aut", Val(.fields("Autorizacion")))
          Print #NumFile, vbTab & vbTab & vbTab & CampoXML("estab", .fields("Establecimiento"))
          Print #NumFile, vbTab & vbTab & vbTab & CampoXML("ptoEmi", .fields("PuntoEmision"))
          Print #NumFile, vbTab & vbTab & vbTab & CampoXML("sec", .fields("Secuencial"))
          Print #NumFile, vbTab & vbTab & vbTab & CampoXML("fechaEmiCom", .fields("FechaEmision"))
          AutRet = Ninguno
          Total = 0
         'Genero el AIR
          sSQL = "SELECT * " _
               & "FROM Trans_Air " _
               & "WHERE Fecha Between #" & FechaIni & "# AND #" & FechaFin & "# " _
               & "AND Periodo = '" & Periodo_Contable & "' "
  If ConSucursal Then
     If Len(No_ATS) > 3 Then sSQL = sSQL & "AND Item NOT IN (" & No_ATS & ") "
  Else
     sSQL = sSQL & "AND Item = '" & NumEmpresa & "' "
  End If

          sSQL = sSQL & "AND Tipo_Trans = 'C' " _
               & "AND IdProv = '" & CodigoCliente & "' " _
               & "AND Factura_No = " & Factura_No & " " _
               & "AND EstabFactura = '" & Codigo1 & "' " _
               & "AND PuntoEmiFactura = '" & Codigo2 & "' " _
               & "ORDER BY SecRetencion,EstabFactura,PuntoEmiFactura,Factura_No "
          Select_Adodc AdoAir, sSQL
          Print #NumFile, vbTab & vbTab & vbTab & AbrirXML("air")
          'MsgBox "No. Retenciones: " & AdoAir.Recordset.RecordCount & vbCrLf & sSQL
          If AdoAir.Recordset.RecordCount > 0 Then
             Do While Not AdoAir.Recordset.EOF
               'Asigno en variables los campos
                Estab = AdoAir.Recordset.fields("EstabRetencion")
                PtoRet = AdoAir.Recordset.fields("PtoEmiRetencion")
                SecRet = AdoAir.Recordset.fields("SecRetencion")
                AutRet = AdoAir.Recordset.fields("AutRetencion")
               'MsgBox AdoAir.Recordset.Fields("Porcentaje") & vbCrLf & AdoAir.Recordset.Fields("BaseImp") & vbCrLf & SubTotal
                Print #NumFile, vbTab & vbTab & vbTab & vbTab & AbrirXML("detalleAir")
                Print #NumFile, vbTab & vbTab & vbTab & vbTab & vbTab & CampoXML("codRetAir", AdoAir.Recordset.fields("CodRet"))
                Print #NumFile, vbTab & vbTab & vbTab & vbTab & vbTab & CampoXML("porcentaje", (AdoAir.Recordset.fields("Porcentaje") * 100), 1)
                If Saldo > 0 Then
                   Print #NumFile, vbTab & vbTab & vbTab & vbTab & vbTab & CampoXML("base0", "0.00")
                   Print #NumFile, vbTab & vbTab & vbTab & vbTab & vbTab & CampoXML("baseGrav", AdoAir.Recordset.fields("BaseImp"), 2)
                Else
                   Print #NumFile, vbTab & vbTab & vbTab & vbTab & vbTab & CampoXML("base0", AdoAir.Recordset.fields("BaseImp"))
                   Print #NumFile, vbTab & vbTab & vbTab & vbTab & vbTab & CampoXML("baseGrav", "0.00", 2)
                End If
                Print #NumFile, vbTab & vbTab & vbTab & vbTab & vbTab & CampoXML("baseNoGrav", "0.00")
                Print #NumFile, vbTab & vbTab & vbTab & vbTab & vbTab & CampoXML("valRetAir", AdoAir.Recordset.fields("ValRet"), 2)
                'MsgBox AdoAir.Recordset.Fields("BaseImp") & vbTab & AdoAir.Recordset.Fields("ValRet")
                Print #NumFile, vbTab & vbTab & vbTab & vbTab & CerrarXML("detalleAir")
                Total = Total + AdoAir.Recordset.fields("ValRet")
                AdoAir.Recordset.MoveNext
             Loop
          End If
          Print #NumFile, vbTab & vbTab & vbTab & CerrarXML("air")
         'Aqui se genera la segunda parte del detalle Compras
          If Val(AutRet) > 0 And Total > 0 Then
             Print #NumFile, vbTab & vbTab & vbTab & CampoXML("autRet", Format(Val(AutRet), "#0"))
             Print #NumFile, vbTab & vbTab & vbTab & CampoXML("estabRet", Estab)
             Print #NumFile, vbTab & vbTab & vbTab & CampoXML("ptoEmiRet", PtoRet)
             Print #NumFile, vbTab & vbTab & vbTab & CampoXML("secRet", Format(SecRet, "#0"))
             Print #NumFile, vbTab & vbTab & vbTab & CampoXML("fechaEmiRet", FechaRet)
          End If
          Print #NumFile, vbTab & vbTab & CerrarXML("detalleCompras")
          'MsgBox Factura_No
         .MoveNext
       Loop
       Print #NumFile, vbTab & CerrarXML("compras")
   Else
       Print #NumFile, vbTab & VacioXML("compras")
   End If
  End With
  Print #NumFile, CerrarXML("reoc")
  Close NumFile   ' Cierra los 3 archivos abiertos.
  Archivo_XML = NombreArchivo
End Sub

Public Sub Crear_Anexos_RDEP(NombreArchivo As String)
Dim Valor_Retenido As Currency
Dim Ruc_Empleado As String
Dim Empleado As String
Dim ApellidoP As String
Dim ApellidoM As String
Dim Nombre1 As String
Dim Nombre2 As String
  Dim NumFile As Long
  Dim CodCiudad As String
  For I = 0 To 14
     Totales_SRI(I) = 0
  Next I
  NumFile = FreeFile
  Open NombreArchivo For Output As #NumFile   ' Abre el archivo
 'ENCABEZADO
  Print #NumFile, "<?xml version='1.0' encoding='ISO-8859-1' standalone='yes'?>"
  Print #NumFile, AbrirXML("rdep")
  Print #NumFile, vbTab & CampoXML("numRuc", RUC)
  Print #NumFile, vbTab & CampoXML("anio", Format(Anio, "0000"))
  Print #NumFile, vbTab & CampoXML("tipoEmpleador", "PRIVADO_MIXTO")
  Print #NumFile, vbTab & CampoXML("enteSegSocial", "IESS")
   With AdoRolPagos.Recordset
    If .RecordCount > 0 Then
        Print #NumFile, AbrirXML("retRelDep")
        Do While Not .EOF
           ApellidoP = ""
           ApellidoM = ""
           Nombre1 = ""
           Nombre2 = ""
           If Len(.fields("Cliente")) > 1 Then
              Empleado = .fields("Cliente")
              ApellidoP = SinEspaciosIzq(Empleado)
              Empleado = TrimStrg(MidStrg(Empleado, Len(ApellidoP) + 1, Len(Empleado)))
              ApellidoM = SinEspaciosIzq(Empleado)
              Empleado = TrimStrg(MidStrg(Empleado, Len(ApellidoM) + 1, Len(Empleado)))
              Nombre1 = SinEspaciosIzq(Empleado)
              Empleado = TrimStrg(MidStrg(Empleado, Len(Nombre1) + 1, Len(Empleado)))
              Nombre2 = Empleado
              If TrimStrg(Nombre1) = "" And TrimStrg(Nombre2) = "" Then
                 Nombre1 = ApellidoM
                 ApellidoM = ""
              End If
           End If
'''           If Anio < 2012 Then
'''              CodCiudad = "99999"
'''              If AdoCiudades.Recordset.RecordCount > 0 Then
'''                 AdoCiudades.Recordset.MoveFirst
'''                 AdoCiudades.Recordset.Find ("CProvincia = '" & .Fields("Prov") & "' ")
'''                 If Not AdoCiudades.Recordset.EOF Then
'''                    CodCiudad = AdoCiudades.Recordset.Fields("CCiudad")
'''                    AdoCiudades.Recordset.Find ("Descripcion_Rubro = '" & .Fields("Ciudad") & "' ")
'''                    If Not AdoCiudades.Recordset.EOF Then
'''                       CodCiudad = AdoCiudades.Recordset.Fields("Codigo")
'''                    End If
'''                 End If
'''              End If
'''           End If
           Ruc_Empleado = .fields("CI_RUC")
           
           Print #NumFile, vbTab & AbrirXML("datRetRelDep")
           
           Print #NumFile, vbTab & vbTab & AbrirXML("empleado")
           Print #NumFile, vbTab & vbTab & vbTab & CampoXML("benGalpg", "NO")
           Print #NumFile, vbTab & vbTab & vbTab & CampoXML("enfcatastro", "NO")
           Print #NumFile, vbTab & vbTab & vbTab & CampoXML("numCargRebGastPers", "0")
           If .fields("TD") = "R" Then
               Print #NumFile, vbTab & vbTab & vbTab & CampoXML("tipIdRet", "C")
               Print #NumFile, vbTab & vbTab & vbTab & CampoXML("idRet", MidStrg(Ruc_Empleado, 1, 10))
           Else
               Print #NumFile, vbTab & vbTab & vbTab & CampoXML("tipIdRet", .fields("TD"))
               Print #NumFile, vbTab & vbTab & vbTab & CampoXML("idRet", Ruc_Empleado)
           End If
           Print #NumFile, vbTab & vbTab & vbTab & CampoXML("apellidoTrab", TrimStrg(ApellidoP & " " & ApellidoM))
           Print #NumFile, vbTab & vbTab & vbTab & CampoXML("nombreTrab", TrimStrg(Nombre1 & " " & Nombre2))
           Print #NumFile, vbTab & vbTab & vbTab & CampoXML("estab", "001")
           If Val(.fields("Pais")) = 593 Then
               Print #NumFile, vbTab & vbTab & vbTab & CampoXML("residenciaTrab", "01")
           Else
               Print #NumFile, vbTab & vbTab & vbTab & CampoXML("residenciaTrab", "02")
           End If
           Print #NumFile, vbTab & vbTab & vbTab & CampoXML("paisResidencia", .fields("Pais"))
           Print #NumFile, vbTab & vbTab & vbTab & CampoXML("aplicaConvenio", .fields("Aplica"))
           Print #NumFile, vbTab & vbTab & vbTab & CampoXML("tipoTrabajDiscap", .fields("Condicion"))
           Print #NumFile, vbTab & vbTab & vbTab & CampoXML("porcentajeDiscap", .fields("Porcentaje"))
           Print #NumFile, vbTab & vbTab & vbTab & CampoXML("tipIdDiscap", .fields("TIdentificacion"))
           Print #NumFile, vbTab & vbTab & vbTab & CampoXML("idDiscap", .fields("Identificacion"))
           Print #NumFile, vbTab & vbTab & CerrarXML("empleado")
           
           
'''           Print #NumFile, vbTab & vbTab & CampoXML("anioRet", Format(Anio, "0000"))
'''           If Anio < 2012 Then
'''              Print #NumFile, vbTab & vbTab & CampoXML("dirCal", TrimStrg(MidStrg(.Fields("Direccion"), 1, 20)))
'''              Print #NumFile, vbTab & vbTab & CampoXML("dirNum", .Fields("DirNumero"))
'''              Print #NumFile, vbTab & vbTab & CampoXML("dirCiu", CodCiudad)     ' 21701
'''              Print #NumFile, vbTab & vbTab & CampoXML("dirProv", MidStrg(CodCiudad, 1, 3))    '217
'''              Print #NumFile, vbTab & vbTab & CampoXML("tel", .Fields("Telefono"))
'''           End If
           Totales_SRI(1) = Totales_SRI(1) + 1
           Print #NumFile, vbTab & vbTab & CampoXML("suelSal", .fields("SuelSal"))
           Totales_SRI(2) = Totales_SRI(2) + .fields("SuelSal")
           Print #NumFile, vbTab & vbTab & CampoXML("sobSuelComRemu", .fields("SobSuelComRemu"))
           Totales_SRI(3) = Totales_SRI(3) + .fields("SobSuelComRemu")
           Print #NumFile, vbTab & vbTab & CampoXML("partUtil", .fields("PartUtil"))
           Totales_SRI(6) = Totales_SRI(6) + .fields("PartUtil")
           Print #NumFile, vbTab & vbTab & CampoXML("intGrabGen", .fields("IngOtrosEmp"))
           Print #NumFile, vbTab & vbTab & CampoXML("impRentEmpl", .fields("ImpRentEmpl"))
           Print #NumFile, vbTab & vbTab & CampoXML("decimTer", .fields("Valor_Dec_3ro"))
           Totales_SRI(4) = Totales_SRI(4) + .fields("Valor_Dec_3ro")
           Print #NumFile, vbTab & vbTab & CampoXML("decimCuar", .fields("Valor_Dec_4to"))
           Totales_SRI(5) = Totales_SRI(5) + .fields("Valor_Dec_4to")
           Print #NumFile, vbTab & vbTab & CampoXML("fondoReserva", .fields("FondoReserva"))
           Print #NumFile, vbTab & vbTab & CampoXML("salarioDigno", "0.00")
           Print #NumFile, vbTab & vbTab & CampoXML("otrosIngRenGrav", "0.00")
           Print #NumFile, vbTab & vbTab & CampoXML("ingGravConEsteEmpl", .fields("SuelSal"))
           Print #NumFile, vbTab & vbTab & CampoXML("sisSalNet", .fields("SN"))
           Print #NumFile, vbTab & vbTab & CampoXML("apoPerIess", .fields("ApoPerIess"))
           Totales_SRI(7) = Totales_SRI(7) + .fields("ApoPerIess")
           Print #NumFile, vbTab & vbTab & CampoXML("aporPerIessConOtrosEmpls", .fields("Rebajas"))
           Print #NumFile, vbTab & vbTab & CampoXML("deducVivienda", .fields("Vivienda"))
           Print #NumFile, vbTab & vbTab & CampoXML("deducSalud", .fields("Salud"))
           'Print #NumFile, vbTab & vbTab & CampoXML("deducEduca", .Fields("Educacion"))
           Print #NumFile, vbTab & vbTab & CampoXML("deducEducartcult", .fields("Educacion"))
           Print #NumFile, vbTab & vbTab & CampoXML("deducAliement", .fields("Alimentacion"))
           Print #NumFile, vbTab & vbTab & CampoXML("deducVestim", .fields("Vestimenta"))
           Print #NumFile, vbTab & vbTab & CampoXML("deduccionTurismo", 0)
           Print #NumFile, vbTab & vbTab & CampoXML("exoDiscap", .fields("RebEspDiscap"))
           Totales_SRI(8) = Totales_SRI(8) + .fields("RebEspDiscap")
           Print #NumFile, vbTab & vbTab & CampoXML("exoTerEd", .fields("RebEspTerEd"))
           Totales_SRI(9) = Totales_SRI(9) + .fields("RebEspTerEd")
           Print #NumFile, vbTab & vbTab & CampoXML("basImp", .fields("BasImp"))
           Totales_SRI(11) = Totales_SRI(11) + .fields("BasImp")
           Valor_Retenido = .fields("ImpRentCaus") + .fields("Deduccion") + .fields("ImpAnterior")
           Print #NumFile, vbTab & vbTab & CampoXML("impRentCaus", .fields("ImpRentCaus"))
           
           Print #NumFile, vbTab & vbTab & CampoXML("rebajaGastosPersonales", "0.00")
           Print #NumFile, vbTab & vbTab & CampoXML("impuestoRentaRebajaGastosPersonales", "0.00")
           
           Print #NumFile, vbTab & vbTab & CampoXML("valRetAsuOtrosEmpls", .fields("Deduccion"))
           Print #NumFile, vbTab & vbTab & CampoXML("valImpAsuEsteEmpl", .fields("ImpAnterior"))
           Print #NumFile, vbTab & vbTab & CampoXML("valRet", Valor_Retenido)
           
           Totales_SRI(12) = Totales_SRI(12) + Valor_Retenido
           Total = .fields("SuelSal") + .fields("SobSuelComRemu") + .fields("PartUtil") + .fields("ImpRentEmpl")

           Totales_SRI(10) = Totales_SRI(10) + .fields("SubTotal")
           Totales_SRI(13) = Totales_SRI(13) + .fields("NumRet")
           
'''           Print #NumFile, vbTab & vbTab & CampoXML("desauOtras", .Fields("Desahucio"))
'''           Print #NumFile, vbTab & vbTab & CampoXML("subTotal", .Fields("SubTotal"))
'''           Print #NumFile, vbTab & vbTab & CampoXML("numRet", .Fields("NumRet"))
'''           Print #NumFile, vbTab & vbTab & CampoXML("numMesEmplead", .Fields("Numero"))
'''           Print #NumFile, vbTab & vbTab & CampoXML("valorImpempAnter", .Fields("ImpAnterior"))
           Print #NumFile, vbTab & CerrarXML("datRetRelDep")
          .MoveNext
        Loop
        Print #NumFile, CerrarXML("retRelDep")
    End If
   End With
 'Retencion sin Dependencia en blanco/vacio
  Print #NumFile, CerrarXML("rdep")
  Close NumFile '
  Archivo_XML = NombreArchivo
End Sub

Private Sub MBFechaE_GotFocus()
  MarcarTexto MBFechaE
End Sub

Private Sub MBFechaE_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBFechaE_LostFocus()
  FechaValida MBFechaE
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim Id_Mes As Byte
  
  Anio_Anexo = Toolbar1.buttons(2).Caption
  Mes_Anexo = Toolbar1.buttons(3).Caption
  Progreso_Barra.Mensaje_Box = "Consultando el ATS de " & Anio_Anexo & "-" & Mes_Anexo
  Progreso_Iniciar
  Archivo_XML = ""
  If Button.key <> "Salir" Then
    'Determinamos la fecha del Anexo
     Fecha_Del_AT Mes_Anexo, Anio_Anexo
     Id_Mes = Month(FechaMitad)
     Anio = Anio_Anexo
     NumMeses = LetrasMeses(Mes_Anexo)
     FA.Factura = 0
     FA.Fecha_Corte = FechaSistema
     Progreso_Esperar
     
     If InStr(1, CStr(FechaFin), "-") > 0 Then Numero = CLng(Replace(FechaFin, "-", "")) Else Numero = CLng(FechaFin)
     
     Actualizar_Datos_ATS_SP NumEmpresa, FechaInicial, FechaFinal, Numero, CheqATSConElect.value
               
    'LISTA DE CODIGO DE ANEXOS
     Progreso_Esperar
     RatonReloj
     sSQL = "SELECT Codigo, Concepto " _
          & "FROM Tipo_Concepto_Retencion " _
          & "WHERE Codigo <> '.' " _
          & "AND Fecha_Inicio <= #" & FechaMid & "# " _
          & "AND Fecha_Final >= #" & FechaMid & "# " _
          & "ORDER BY Codigo "
     Select_Adodc AdoCodRet, sSQL
  End If
  RatonReloj
 'Limpiamos la pagina si no imprimimos
 'MsgBox Button.key
     '<<--== Hasta aqui
  Select Case Button.key
    Case "ATS"
        'NumMeses = 1
        'Aqui se genera el Archivo XML
        ' Insertar_Ventas_ATS
         'Insertar_Liquidacion_Compras_ATS
         Consultar_Anexos
         Archivo_XML = RutaSysBases & "\AT\AT" & NumEmpresa & "\AT"
         Select Case Mes_Anexo
           Case "1er_Semestre", "2do_Semestre"
                Archivo_XML = Archivo_XML & "_" & Mes_Anexo
           Case Else
                Archivo_XML = Archivo_XML & Format(LetrasMeses(Mes_Anexo), "00")
         End Select
         Archivo_XML = Archivo_XML & "_" & Anio_Anexo & ".xml"
         RutaSubDirTemp = Archivo_XML
         Progreso_Barra.Mensaje_Box = "ATS-Generando en: " & Archivo_XML
         Progreso_Esperar
         Crear_Anexos_Tipo_01_Marzo_2015 Archivo_XML, CheqATSConElect.value
'''         Select Case Val(Anio)
'''           Case 2008 To 2012: Crear_Anexos_Tipo_01 Archivo_XML
'''           Case 2013 To 2050:
'''         End Select
        'Presentamos el anexo en la pantalla
         Carga_Paginas_AT
    Case "ATSFinanc"
        'NumMeses = 1
        'Aqui se genera el Archivo XML
'         Insertar_Ventas_ATS
         Consultar_Anexos
         Archivo_XML = RutaSysBases & "\AT\AT" & NumEmpresa & "\AT" & Format(LetrasMeses(Mes_Anexo), "00") & Anio_Anexo & ".xml"
         RutaSubDirTemp = Archivo_XML
         Crear_Anexos_Tipo_02 Archivo_XML
        'Presentamos el anexo en la pantalla
         Carga_Paginas_AT
    Case "REOC"
        'NumMeses = 1
        'Aqui se genera el Archivo XML
         Consultar_Anexos
         Archivo_XML = RutaSysBases & "\AT\AT" & NumEmpresa & "\REOC" & Format(LetrasMeses(Mes_Anexo), "00") & Anio_Anexo & ".xml"
         RutaSubDirTemp = Archivo_XML
         Crear_Anexos_REOC Archivo_XML
        'Presentamos el anexo en la pantalla
         Carga_Paginas_REOC
    Case "RDEP"
        'NumMeses = 1
         Consultar_Anexos
         Archivo_XML = RutaSysBases & "\AT\AT" & NumEmpresa & "\RDEP" & Format(Anio, "0000") & ".xml"
         Crear_Anexos_RDEP Archivo_XML
         Carga_Paginas_RDEP
    Case "Salir"
         Progreso_Final
         Unload FAnexoTransaccional
  End Select
  If Button.key <> "Salir" Then
     Progreso_Final
     MsgBox "EL ARCHIVO SE GENERO EN: " & vbCrLf & vbCrLf & Archivo_XML
  End If
  
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
  If IsNumeric(ButtonMenu) Then
     Anio_Anexo = ButtonMenu
     Toolbar1.buttons(2).Caption = Anio_Anexo
  Else
     Mes_Anexo = ButtonMenu
     Toolbar1.buttons(3).Caption = Mes_Anexo
  End If
'  SSTab1.Caption = "ANEXO TRANSACIONAL DE " & Anio_Anexo & "-" & UCaseStrg(Mes_Anexo)
  Toolbar1.Refresh
End Sub

Public Sub Carga_Paginas_AT()
Dim NombFilePict As String

  Segunda_Pag = False
  NombFilePict = "ATS " & Format(NumMeses, "00") & "-" & Anio_Anexo & " " & RUC
  SetNombrePRN = Impresota_PDF
  
 'Geneeramos el documento
  tPrint.TipoImpresion = Es_PDF
  tPrint.NombreArchivo = NombFilePict
  tPrint.TituloArchivo = "Anexo Transaccional del " & Format(NumMeses, "00") & "-" & Anio_Anexo & " " & RUC
  tPrint.TipoLetra = TipoHelvetica
  tPrint.OrientacionPagina = 1
  tPrint.PaginaA4 = True
  tPrint.EsCampoCorto = False
  tPrint.VerDocumento = True
  Set cPrint = New cImpresion
  cPrint.iniciaImpresion

  PosLinea = 1
  Pagina = 1
 'Pagina No. 1
  cPrint.printImagen RutaSistema & "\LOGOS\SRI.jpg", 1, 0.7, 3.5, 2
  cPrint.fondoDeLetra = Negro
  cPrint.tipoNegrilla = True
  cPrint.PorteDeLetra = 14
  cPrint.colorDeLetra = Azul
  cPrint.printTexto 5, PosLinea, "TALON RESUMEN DE ANEXO TRANSACCIONAL", 14, "C", 19.5
  PosLinea = PosLinea + 0.5
  cPrint.PorteDeLetra = 12
  cPrint.printTexto 5, PosLinea, "SERVICIO DE RENTAS INTERNAS", 12, "C", 19.5
  PosLinea = PosLinea + 0.5
  cPrint.PorteDeLetra = 14
  cPrint.printTexto 5, PosLinea, NombreComercial, 14, "C", 19.5
  PosLinea = PosLinea + 0.6
  cPrint.PorteDeLetra = 12
  cPrint.printTexto 5, PosLinea, "RUC:" & RUC, 12, "C", 19.5
  PosLinea = PosLinea + 0.6
  cPrint.PorteDeLetra = 9
  cPrint.tipoNegrilla = False
  cPrint.colorDeLetra = Negro
  Cadena = "Certifico que la información contenida en el medio magnético adjunto al presente " _
         & "Anexo Transaccional para el período " & Format(NumMeses, "00") & "-" & Anio_Anexo & " " _
         & "es fiel reflejo del siguiente reporte:"
  PosLinea = cPrint.printTextoMultiple(1.5, PosLinea, Cadena, 19)
  PosLinea = PosLinea + 0.6
 'COMPRAS
  Vista_Compras
 'VENTAS
  Vista_Ventas
 'ANULADOS
  Vista_Anulados
 'RETENCION IMPUESTO RENTA
  Vista_Retencion_Impuesto_Renta
 'RETENCION FUENTE IVA
  Vista_Retencion_Fuente_Iva
  
  sSQL = "SELECT TI.TipoComprobante,TCC.Descripcion,COUNT(TipoComprobante) As Cant,SUM(ValorCIF) As VC,SUM(MontoIva) As MI " _
      & "FROM Trans_Importaciones As TI,Clientes As C,Tipo_Comprobante As TCC " _
      & "WHERE TI.Fecha Between #" & FechaIni & "# AND #" & FechaFin & "#  "
  If ConSucursal Then
     If Len(No_ATS) > 3 Then sSQL = sSQL & "AND TI.Item NOT IN (" & No_ATS & ") "
  Else
     sSQL = sSQL & "AND TI.Item = '" & NumEmpresa & "' "
  End If
  sSQL = sSQL _
       & "AND TI.Periodo = '" & Periodo_Contable & "' " _
       & "AND TCC.TC = 'TDC' " _
       & "AND TI.TipoComprobante = TCC.Tipo_Comprobante_Codigo  " _
       & "AND TI.IdFiscalProv = C.Codigo " _
       & "GROUP BY TI.TipoComprobante,TCC.Descripcion " _
       & "ORDER BY TI.TipoComprobante "
  Select_Adodc AdoAux, sSQL
  If AdoAux.Recordset.RecordCount > 0 Then Segunda_Pag = True
  
  sSQL = "SELECT TE.TipoComprobante,TCC.Descripcion,COUNT(TE.TipoComprobante) As Cant,SUM(ValorFOB) As VF " _
       & "FROM Trans_Exportaciones As TE,Clientes As C,Tipo_Comprobante As TCC " _
       & "WHERE TE.Fecha Between #" & FechaIni & "# AND #" & FechaFin & "#  "
  If ConSucursal Then
     If Len(No_ATS) > 3 Then sSQL = sSQL & "AND TE.Item NOT IN (" & No_ATS & ") "
  Else
     sSQL = sSQL & "AND TE.Item = '" & NumEmpresa & "' "
  End If
  sSQL = sSQL _
       & "AND TE.Periodo = '" & Periodo_Contable & "' " _
       & "AND TCC.TC = 'TDC' " _
       & "AND TE.TipoComprobante = TCC.Tipo_Comprobante_Codigo  " _
       & "AND TE.IdFiscalProv = C.Codigo " _
       & "GROUP BY TE.TipoComprobante,TCC.Descripcion " _
       & "ORDER BY TE.TipoComprobante "
  Select_Adodc AdoAux, sSQL
  If AdoAux.Recordset.RecordCount > 0 Then Segunda_Pag = True
  
  If Segunda_Pag Then
    'Pagina No. 2
      
     cPrint.paginaNueva
     PosLinea = 2.5
     Pagina = 2
    'IMPORTACIONES
     Vista_Importaciones
    'EXPORTACIONES
     Vista_Exportaciones
    'EMPRESA EMISORA
    '  Vista_EmpresaEmisora
    'FIDEICOMISOS
    '  Vista_Fondos_Fideicomisos
  End If
 'VISTA FINAL
  If PosLinea < (cPrint.dAltoPapel - 2) Then
     Vista_Final
  Else
     cPrint.paginaNueva
     PosLinea = 2.5
     Vista_Final
  End If
 'fin del documento
  cPrint.finalizaImpresion
  'Presentar_PDF ATSPDF, RutaDocumentoPDF
End Sub

Public Sub Carga_Paginas_RDEP()
Dim NombFilePict As String
  PosLinea = 1
  Pagina = 1
 'Pagina No. 1
 'Geneeramos el documento
  NombFilePict = "RDEP " & Format(NumMeses, "00") & "-" & Anio_Anexo & " " & RUC
  SetNombrePRN = Impresota_PDF
  
  tPrint.TipoImpresion = Es_PDF
  tPrint.NombreArchivo = NombFilePict
  tPrint.TituloArchivo = "Anexo RDEP del " & Format(NumMeses, "00") & "-" & Anio_Anexo & " " & RUC
  tPrint.TipoLetra = TipoHelvetica
  tPrint.OrientacionPagina = 1
  tPrint.PaginaA4 = True
  tPrint.EsCampoCorto = False
  tPrint.VerDocumento = True
  Set cPrint = New cImpresion
  cPrint.iniciaImpresion
 
  cPrint.printImagen RutaSistema & "\LOGOS\SRI.JPG", 1.5, 0.7, 4, 2.5
  cPrint.fondoDeLetra = Negro
  cPrint.tipoDeLetra = TipoHelvetica
  cPrint.tipoNegrilla = True
  cPrint.PorteDeLetra = 11
  cPrint.printTexto 7.5, PosLinea, "TALON RESUMEN DE RETENCIONES EN LA FUENTE DE"
  PosLinea = PosLinea + 0.6
  cPrint.printTexto 7, PosLinea, "IMPUESTO A LA RENTA BAJO RELACIÓN DE DEPENDENCIA"
  PosLinea = PosLinea + 0.6
  cPrint.printTexto 9, PosLinea, "SERVICIO DE RENTAS INTERNAS - RIG"
  PosLinea = PosLinea + 0.6
  cPrint.printTexto 6, PosLinea, "RAZÓN SOCIAL:"
  cPrint.tipoNegrilla = False
  cPrint.printTexto 9.5, PosLinea, UCaseStrg(NombreComercial)
  PosLinea = PosLinea + 0.6
  cPrint.tipoNegrilla = True
  cPrint.printTexto 6, PosLinea, "RUC:"
  cPrint.tipoNegrilla = False
  cPrint.printTexto 9.5, PosLinea, RUC
  PosLinea = PosLinea + 0.6
  cPrint.tipoNegrilla = True
  cPrint.printTexto 6, PosLinea, "PERÍODO:"
  cPrint.tipoNegrilla = False
  cPrint.printTexto 9.5, PosLinea, "Enero a Diciembre de " & Format(Anio, "0000")
  PosLinea = PosLinea + 0.6
  cPrint.tipoNegrilla = True
  cPrint.printTexto 6, PosLinea, "FECHA:"
  cPrint.tipoNegrilla = False
  cPrint.printTexto 9.5, PosLinea, FechaSistema & " " & Format(Time, "hh:mm")
  PosLinea = PosLinea + 2
  cPrint.PorteDeLetra = 11
  cPrint.tipoNegrilla = False
  Cadena = "Certifico que la información contenida en el medio magnético adjunto sobre la " _
         & "Retención en la fuente de impuesto a la Renta bajo Relación de Dependencia Realizadas " _
         & "durante el año indicado, es el fiel reflejo de los registrado en este formulario:"
  PosLinea = cPrint.printTextoMultiple(1.2, PosLinea, Cadena, 19)
  PosLinea = PosLinea + 2
  cPrint.tipoNegrilla = True
  cPrint.tipoDeLetra = TipoArial
  
  cPrint.printTexto 3.8, PosLinea, "INFORMACIÓN ORIGINAL", 11, "C", 17, Azul_Claro
  PosLinea = PosLinea + 0.7
  cPrint.printTexto 6, PosLinea, "Descripción"
  cPrint.printTexto 15, PosLinea, "Valor"
  PosLinea = PosLinea + 0.6
  cPrint.PorteDeLetra = 9
  cPrint.tipoNegrilla = False
  cPrint.printTexto 4, PosLinea, "Número de Registros"
  cPrint.printVariable 14.2, PosLinea, Totales_SRI(1)
  PosLinea = PosLinea + 0.6
  cPrint.printTexto 4, PosLinea, "Sueldos y Salarios"
  cPrint.printVariable 14.5, PosLinea, Totales_SRI(2)
  PosLinea = PosLinea + 0.6
  cPrint.printTexto 4, PosLinea, "Sobresueldos, Comisiones y Otras Remuneraciones"
  cPrint.printVariable 14.5, PosLinea, Totales_SRI(3)
  PosLinea = PosLinea + 0.6
  cPrint.printTexto 4, PosLinea, "Décimo Tercer Sueldo"
  cPrint.printVariable 14.5, PosLinea, Totales_SRI(4)
  PosLinea = PosLinea + 0.6
  cPrint.printTexto 4, PosLinea, "Décimo Cuarto Sueldo"
  cPrint.printVariable 14.5, PosLinea, Totales_SRI(5)
  PosLinea = PosLinea + 0.6
  cPrint.printTexto 4, PosLinea, "Participación Utilidades"
  cPrint.printVariable 14.5, PosLinea, Totales_SRI(6)
  PosLinea = PosLinea + 0.6
  cPrint.printTexto 4, PosLinea, "Aporte IESS"
  cPrint.printVariable 14.5, PosLinea, Totales_SRI(7)
  PosLinea = PosLinea + 0.6
  cPrint.printTexto 4, PosLinea, "Rebajas Especiales Discapacitados"
  cPrint.printVariable 14.5, PosLinea, Totales_SRI(8)
  PosLinea = PosLinea + 0.6
  cPrint.printTexto 4, PosLinea, "Rebajas Especiales Tercera Edad"
  cPrint.printVariable 14.5, PosLinea, Totales_SRI(9)
  PosLinea = PosLinea + 0.6
  cPrint.printTexto 4, PosLinea, "Sub Total"
  cPrint.printVariable 14.5, PosLinea, Totales_SRI(10)
  PosLinea = PosLinea + 0.6
  cPrint.printTexto 4, PosLinea, "Base Imponible"
  cPrint.printVariable 14.5, PosLinea, Totales_SRI(11)
  PosLinea = PosLinea + 0.6
  cPrint.printTexto 4, PosLinea, "Valor Retenido"
  cPrint.printVariable 14.5, PosLinea, Totales_SRI(12)
  PosLinea = PosLinea + 0.6
  cPrint.printTexto 4, PosLinea, "Número de Retenciones"
  cPrint.printVariable 14.2, PosLinea, Totales_SRI(13)
  PosLinea = PosLinea + 2
  cPrint.tipoDeLetra = TipoHelvetica
  cPrint.PorteDeLetra = 9
  Cadena = "Declaro que los datos contenidos en este anexo son verdaderos, por lo que asumo la responsabilidad correspondiente, " _
         & "de acuerdo a lo establecido en el Art. 101 de la Codificación de la Ley de Régimen Tributario Interno."
  PosLinea = cPrint.printTextoMultiple(1.2, PosLinea, Cadena, 19)
  PosLinea = PosLinea + 2
  cPrint.PorteDeLetra = 10
  cPrint.printTexto 1.2, PosLinea, "__________________________"
  cPrint.printTexto 10.5, PosLinea, "______________________________"
  PosLinea = PosLinea + 0.6
  cPrint.printTexto 2, PosLinea, "Firma del Contador"
  cPrint.printTexto 11, PosLinea, "Firma del Representante Legal"
 'RDEP
  cPrint.finalizaImpresion
  'Presentar_PDF ATSPDF, RutaDocumentoPDF
End Sub

Public Sub Carga_Paginas_REOC()
Dim NombFilePict As String
Dim PosXo As Single
Dim PosYo As Single
Dim PosXf As Single
Dim PosYf As Single
  For I = 0 To 10
      Totales_SRI(I) = 0
  Next I
 'AIR
  sSQL = "SELECT TA.CodRet, TC.TipoComprobante, " _
       & "COUNT(TA.IdProv) As Cantidad, SUM(TA.BaseImp) As BaseImpo, SUM(TA.ValRet) As Retenido " _
       & "FROM Trans_Compras As TC,Trans_Air As TA, Clientes As C " _
       & "WHERE TC.Fecha Between #" & FechaIni & "# AND #" & FechaFin & "# " _
       & "AND TC.Periodo = '" & Periodo_Contable & "' "
  If ConSucursal Then
     If Len(No_ATS) > 3 Then sSQL = sSQL & "AND Item NOT IN (" & No_ATS & ") "
  Else
     sSQL = sSQL & "AND Item = '" & NumEmpresa & "' "
  End If

        sSQL = sSQL & "AND TA.Tipo_Trans = 'C' " _
       & "AND TA.Factura_No = TC.Secuencial " _
       & "AND TA.EstabFactura = TC.Establecimiento " _
       & "AND TA.PuntoEmiFactura = TC.PuntoEmision " _
       & "AND TC.Item = TA.Item " _
       & "AND TC.Periodo = TA.Periodo " _
       & "AND TC.IdProv = C.Codigo " _
       & "AND TA.IdProv = C.Codigo " _
       & "GROUP BY TA.CodRet, TC.TipoComprobante " _
       & "ORDER BY TA.CodRet, TC.TipoComprobante "
  Select_Adodc AdoAir, sSQL
 'Empezamos la presentacion del REOC
  PosLinea = 1
  Pagina = 1
  cPrint.printImagen RutaSistema & "\LOGOS\SRI.JPG", 1, 0.7, 4, 2.5
  cPrint.fondoDeLetra = Negro
  cPrint.tipoDeLetra = TipoVerdana
  cPrint.tipoNegrilla = True
  cPrint.PorteDeLetra = 10
  cPrint.printTexto 6.5, PosLinea, "TALON RESUMEN DE RETENCIONES EN LA FUENTE DE"
  PosLinea = PosLinea + 0.6
  cPrint.printTexto 6, PosLinea, "IMPUESTO A LA RENTA BAJO RELACIÓN DE DEPENDENCIA"
  PosLinea = PosLinea + 0.6
  cPrint.printTexto 8, PosLinea, "SERVICIO DE RENTAS INTERNAS - RIG"
  cPrint.PorteDeLetra = 9
  PosLinea = PosLinea + 0.6
  cPrint.printTexto 5, PosLinea, "RAZÓN SOCIAL:"
  cPrint.tipoNegrilla = False
  cPrint.printTexto 8, PosLinea, UCaseStrg(NombreComercial)
  PosLinea = PosLinea + 0.6
  cPrint.tipoNegrilla = True
  cPrint.printTexto 5, PosLinea, "RUC:"
  cPrint.tipoNegrilla = False
  cPrint.printTexto 8, PosLinea, RUC
  PosLinea = PosLinea + 1
  cPrint.PorteDeLetra = 10
  cPrint.tipoNegrilla = False
  Cadena = "Certifico que la información contenida en el medio magnético adjunto al presente " _
         & "Anexo de Compras y Retenciones en la fuente de Impuesto a la Renta por " _
         & "Otros Conceptos para el período " & Format(NumMeses, "00") & "-" & Format(Anio, "0000") _
         & ", es fiel reflejo del siguiente reporte:"
  PosLinea = cPrint.printTextoMultiple(1.2, 19, PosLinea, Cadena)
  PosLinea = PosLinea + 1
  PosXo = 0.7
  PosXf = 19.5
  PosYo = PosLinea - 0.05
  cPrint.tipoNegrilla = True
  cPrint.printTexto 0.7, PosLinea, "RESUMEN DE RETENCIONES -- RETENCION EN LA FUENTE DE IMPUESTO A LA RENTA", "C", 19.5, Azul_Claro
  PosLinea = PosLinea + 0.6
  cPrint.PorteDeLetra = 9
  cPrint.tipoDeLetra = TipoArial
  cPrint.printTexto 0.9, PosLinea, "Cód."
  cPrint.printTexto 4, PosLinea, "Concepto de Retención"
  cPrint.printTexto 11.6, PosLinea, "No. Registros"
  cPrint.printTexto 14.3, PosLinea, "Base Imponible"
  cPrint.printTexto 17, PosLinea, "Valor Retenido"
  cPrint.PorteDeLetra = 9
  PosLinea = PosLinea + 0.5
  cPrint.tipoNegrilla = False
  Total = 0
  Saldo = 0
  Cantidad = 0
  With AdoAir.Recordset
   If .RecordCount > 0 Then
       Codigo = .fields("CodRet")
       Do While Not .EOF
          If Codigo <> .fields("CodRet") Then
             DetalleComp = Ninguno
             If AdoCodRet.Recordset.RecordCount > 0 Then
                AdoCodRet.Recordset.MoveFirst
                AdoCodRet.Recordset.Find ("Codigo = '" & Codigo & "' ")
                If Not AdoCodRet.Recordset.EOF Then DetalleComp = UCaseStrg(AdoCodRet.Recordset.fields("Concepto"))
             End If
             cPrint.tipoDeLetra = TipoArialNarrow
             cPrint.printTexto 0.9, PosLinea, Codigo
             cPrint.printTexto 1.8, PosLinea, DetalleComp
             cPrint.tipoDeLetra = TipoArial
             cPrint.printTexto 11.4, PosLinea, Format(Cantidad, "##0"), True, 2
             cPrint.printTexto 14.6, PosLinea, Format(Total, "#,##0.00"), True, 2
             cPrint.printTexto 17.2, PosLinea, Format(Saldo, "#,##0.00"), True, 2
             PosLinea = PosLinea + 0.5
             Total = 0
             Saldo = 0
             Cantidad = 0
             Codigo = .fields("CodRet")
          End If
          Select Case .fields("TipoComprobante")
            Case 4: Totales_SRI(1) = Totales_SRI(1) + .fields("Cantidad")  '04
            Case 5: Totales_SRI(2) = Totales_SRI(2) + .fields("Cantidad")  '05
            Case 41: Totales_SRI(3) = Totales_SRI(3) + .fields("Cantidad") '41
            Case 47: Totales_SRI(4) = Totales_SRI(4) + .fields("Cantidad") '47
            Case 48: Totales_SRI(5) = Totales_SRI(5) + .fields("Cantidad") '48
            Case Else
                 Totales_SRI(0) = Totales_SRI(0) + .fields("Cantidad")     '04-05-47-48-07-41
                 Totales_SRI(10) = Totales_SRI(10) + .fields("BaseImpo")
                 Totales_SRI(11) = Totales_SRI(11) + .fields("Retenido")
                 Cantidad = Cantidad + .fields("Cantidad")
                 Total = Total + .fields("BaseImpo")
                 Saldo = Saldo + .fields("Retenido")
          End Select
         .MoveNext
       Loop
       DetalleComp = Ninguno
       If AdoCodRet.Recordset.RecordCount > 0 Then
          AdoCodRet.Recordset.MoveFirst
          AdoCodRet.Recordset.Find ("Codigo = '" & Codigo & "' ")
          If Not AdoCodRet.Recordset.EOF Then DetalleComp = UCaseStrg(AdoCodRet.Recordset.fields("Concepto"))
       End If
       cPrint.tipoDeLetra = TipoArialNarrow
       cPrint.printTexto 0.9, PosLinea, Codigo
       cPrint.printTexto 1.8, PosLinea, DetalleComp
       cPrint.tipoDeLetra = TipoArial
       cPrint.printTexto 11.4, PosLinea, Format(Cantidad, "##0"), True, 2
       cPrint.printTexto 14.6, PosLinea, Format(Total, "#,##0.00"), True, 2
       cPrint.printTexto 17.2, PosLinea, Format(Saldo, "#,##0.00"), True, 2
       PosLinea = PosLinea + 0.5
   End If
  End With
  cPrint.tipoDeLetra = TipoArial
  cPrint.tipoNegrilla = True
  cPrint.printTexto 8.5, PosLinea, "TOTAL", , 11.5, Azul_Claro
  cPrint.tipoNegrilla = False
  cPrint.printTexto 11.4, PosLinea, Format(Totales_SRI(0), "##0"), True, 2
  cPrint.printTexto 14.6, PosLinea, Format(Totales_SRI(10), "#,##0.00"), True, 2
  cPrint.printTexto 17.2, PosLinea, Format(Totales_SRI(11), "#,##0.00"), True, 2
  PosLinea = PosLinea + 0.45
  PosYf = PosLinea
  cPrint.printCuadro PosXo, PosYo, PosXf, PosYf, Negro, "B"
  cPrint.printCuadro PosXo, PosYo + 0.55, PosXf, PosYo + 0.55, Negro, "B"
  cPrint.printCuadro PosXo, PosYo + 1.05, PosXf, PosYo + 1.05, Negro, "B"
  PosLinea = PosLinea + 0.3
  PosYo = PosLinea - 0.05
  cPrint.printTexto 0.9, PosLinea, "Total de Comprobante de Venta execpto N/C y N/D"
  cPrint.printTexto 11.4, PosLinea, Format(Totales_SRI(0), "##0"), True, 2
  PosLinea = PosLinea + 0.5
  cPrint.printTexto 0.9, PosLinea, "Total de Notas de Crédito"
  cPrint.printTexto 11.4, PosLinea, Format(Totales_SRI(1), "#0"), True, 2
  PosLinea = PosLinea + 0.5
  cPrint.printTexto 0.9, PosLinea, "Total de Notas de Débito"
  cPrint.printTexto 11.4, PosLinea, Format(Totales_SRI(2), "#0"), True, 2
  PosLinea = PosLinea + 0.5
  cPrint.printTexto 0.9, PosLinea, "Total de Comprobantes de Reembolsos de Gastos execpto N/C y N/D"
  PosLinea = PosLinea + 0.5
  cPrint.printTexto 0.9, PosLinea, "de Reembolsos de Gastos"
  cPrint.printTexto 11.4, PosLinea, Format(Totales_SRI(3), "##0"), True, 2
  PosLinea = PosLinea + 0.5
  cPrint.printTexto 0.9, PosLinea, "Total de Notas de Crédito de Reembolsos de Gastos"
  cPrint.printTexto 11.4, PosLinea, Format(Totales_SRI(4), "##0"), True, 2
  PosLinea = PosLinea + 0.5
  cPrint.printTexto 0.9, PosLinea, "Total de Notas de Dédito de Reembolsos de Gastos"
  cPrint.printTexto 11.4, PosLinea, Format(Totales_SRI(5), "##0"), True, 2
  PosLinea = PosLinea + 0.5
  PosYf = PosLinea
  cPrint.printCuadro PosXo, PosYo, PosXf, PosYf, Negro, "B"
  PosLinea = PosLinea + 1.5
  cPrint.PorteDeLetra = 11
  cPrint.tipoDeLetra = TipoArialNarrow
  cPrint.printTexto 0.7, PosLinea, "Declaro que los datos contenidos en este anexo son verdaderos, por lo que asumo la responsabilidad, de acuerdo a lo establecido"
  PosLinea = PosLinea + 0.45
  cPrint.printTexto 0.7, PosLinea, "en el Art. 101 de la Codificación de la Ley de Régimen Tributario Interno."
  PosLinea = PosLinea + 2
  cPrint.PorteDeLetra = 11
  cPrint.printTexto 2.2, PosLinea, "__________________________"
  cPrint.printTexto 10.5, PosLinea, "______________________________"
  PosLinea = PosLinea + 0.6
  cPrint.printTexto 3, PosLinea, "Firma del Contador"
  cPrint.printTexto 11, PosLinea, "Firma del Representante Legal"
 'REOC
   
  cPrint.finalizaImpresion
  DetalleComp = Ninguno
  MsgBox "EL ARCHIVO SE GENERÓ EN: " & vbCrLf & vbCrLf _
       & Archivo_XML & vbCrLf
End Sub

Public Sub Vista_Compras()
 sSQL = "SELECT TC.TipoComprobante,TCC.Descripcion,COUNT(TipoComprobante) As Cant,SUM(BaseImponible) As BI,SUM(BaseImpGrav) As BIG,SUM(MontoIva) As MI " _
      & "FROM Trans_Compras As TC,Clientes As C,Tipo_Comprobante As TCC " _
      & "WHERE TC.Fecha Between #" & FechaIni & "# AND #" & FechaFin & "#  "
 If ConSucursal Then
    If Len(No_ATS) > 3 Then sSQL = sSQL & "AND TC.Item NOT IN (" & No_ATS & ") "
 Else
    sSQL = sSQL & "AND TC.Item = '" & NumEmpresa & "' "
 End If
 sSQL = sSQL _
      & "AND TC.Periodo = '" & Periodo_Contable & "' " _
      & "AND TCC.TC = 'TDC' " _
      & "AND TC.TipoComprobante = TCC.Tipo_Comprobante_Codigo  " _
      & "AND TC.IdProv = C.Codigo " _
      & "GROUP BY TC.TipoComprobante,TCC.Descripcion " _
      & "ORDER BY TC.TipoComprobante "
 Select_Adodc AdoAux, sSQL
 With AdoAux.Recordset
  If .RecordCount > 0 Then
     'Sumo para cada tipo de comprobante
      Real1 = 0
      Real2 = 0
      Real3 = 0

      cPrint.tipoDeLetra = TipoArial
      cPrint.tipoNegrilla = True
      cPrint.PorteDeLetra = 11
      cPrint.colorDeLetra = Blanco
      cPrint.printTexto 1.5, PosLinea, "C O M P R A S", 11, "C", 20, Plata
      cPrint.colorDeLetra = Negro
      PosLinea = PosLinea + 0.5
      cPrint.PorteDeLetra = 8
      cPrint.printTexto 1.5, PosLinea, "Cód"
      cPrint.printTexto 6, PosLinea, "Transacción"
      cPrint.printTexto 12, PosLinea, "No. Reg."
      cPrint.printTexto 13.6, PosLinea, "BI tarifa 0%"
      cPrint.printTexto 15.6, PosLinea, "BI tarifa 12%"
      cPrint.printTexto 18.1, PosLinea, "Valor IVA"
      cPrint.tipoNegrilla = False
      cPrint.PorteDeLetra = 8
      PosLinea = PosLinea + 0.4
      cPrint.printLinea 1.5, PosLinea, 19.5, PosLinea, Negro
      PosLinea = PosLinea + 0.1
      Do While Not .EOF
         cPrint.printTexto 1.5, PosLinea, .fields("TipoComprobante")
         cPrint.printTexto 2, PosLinea, UCaseStrg(.fields("Descripcion"))
         cPrint.printVariable 11.2, PosLinea, .fields("Cant")
         cPrint.printVariable 13, PosLinea, .fields("BI")
         cPrint.printVariable 15.1, PosLinea, .fields("BIG")
         cPrint.printVariable 17.3, PosLinea, .fields("MI")
         PosLinea = PosLinea + 0.35
         
         'MsgBox .Fields("TipoComprobante")
         Real1 = Real1 + .fields("BI")
         Real2 = Real2 + .fields("BIG")
         Real3 = Real3 + .fields("MI")
        .MoveNext
      Loop
      PosLinea = PosLinea + 0.1
      cPrint.printLinea 1.5, PosLinea, 19.5, PosLinea, Azul
      cPrint.printTexto 1.5, PosLinea, "TOTAL"
      cPrint.printVariable 13, PosLinea, Real1
      cPrint.printVariable 15.1, PosLinea, Real2
      cPrint.printVariable 17.3, PosLinea, Real3
      PosLinea = PosLinea + 0.5
  End If
 End With
 sSQL = "SELECT TC.CodSustento,TCC.Descripcion,COUNT(TipoComprobante) As Cant,SUM(BaseImponible) As BI,SUM(BaseImpGrav) As BIG,SUM(MontoIva) As MI " _
      & "FROM Trans_Compras As TC,Clientes As C,Tipo_Tributario As TCC " _
      & "WHERE TC.Fecha Between #" & FechaIni & "# AND #" & FechaFin & "#  "
 If ConSucursal Then
    If Len(No_ATS) > 3 Then sSQL = sSQL & "AND TC.Item NOT IN (" & No_ATS & ") "
 Else
    sSQL = sSQL & "AND TC.Item = '" & NumEmpresa & "' "
 End If
 sSQL = sSQL & "AND TC.Periodo = '" & Periodo_Contable & "' " _
      & "AND TC.CodSustento IN ('01','03','06','08','09') " _
      & "AND TC.CodSustento = TCC.Credito_Tributario  " _
      & "AND TC.IdProv = C.Codigo " _
      & "GROUP BY TC.CodSustento,TCC.Descripcion " _
      & "ORDER BY TC.CodSustento "
 Select_Adodc AdoAux, sSQL
 With AdoAux.Recordset
  If .RecordCount > 0 Then
      cPrint.PorteDeLetra = 8
      cPrint.tipoNegrilla = True
      cPrint.printTexto 1.5, PosLinea, "Se verificará con los casilleros asignados en la declaración de IVA (form.104) de acuerdo al siguiente esquema:"
      PosLinea = PosLinea + 0.4
      cPrint.tipoNegrilla = False
      cPrint.tipoSubrayado = True
      cPrint.printTexto 1.5, PosLinea, "Sustento Crédito Tributario"
      cPrint.printTexto 13, PosLinea, "Casilleros"
      cPrint.printTexto 15.6, PosLinea, "Base"
      cPrint.printTexto 18, PosLinea, "Impuesto"
      cPrint.tipoSubrayado = False
      PosLinea = PosLinea + 0.5
      Real1 = 0 ' Base Imp 12% 1,3,6
      Real2 = 0 ' IVA 12% 1,3,6
      Real3 = 0 ' 0% 1,3,6
      Real4 = 0 ' Base Imp 12% 8 y 9
      Real5 = 0 ' IVA 12% 8 y 9
      Real6 = 0 ' 0% 8 y 9
      Do While Not .EOF
         Select Case .fields("CodSustento")
           Case "01", "03", "06"
                Real1 = Real1 + .fields("BIG")
                Real2 = Real2 + .fields("MI")
                Real3 = Real3 + .fields("BI")
           Case "08", "09"
                Real4 = Real4 + .fields("BIG")
                Real5 = Real5 + .fields("MI")
                Real6 = Real6 + .fields("BI")
         End Select
        .MoveNext
      Loop
      cPrint.printTexto 1.5, PosLinea, "Compras Netas (12/14)% -Sustento de crédito tributario corresponde a los códigos 1,3 y 6"
      cPrint.printTexto 13, PosLinea, "631+633+635"
      cPrint.printVariable 15.1, PosLinea, Real1
      cPrint.printVariable 17.3, PosLinea, Real2
      PosLinea = PosLinea + 0.35
      cPrint.printTexto 1.5, PosLinea, "Compras Netas 0% -Sustento de crédito tributario corresponde a los códigos 1,3 y 6"
      cPrint.printTexto 13, PosLinea, "601+603+605"
      cPrint.printVariable 15.1, PosLinea, Real3
      PosLinea = PosLinea + 0.35
      cPrint.printTexto 1.5, PosLinea, "Pago por Reembolso de gastos (12/14)% -Corresponde a los códigos 8 y 9"
      cPrint.printTexto 13, PosLinea, "637"
      cPrint.printVariable 15.1, PosLinea, Real4
      cPrint.printVariable 17.3, PosLinea, Real5
      PosLinea = PosLinea + 0.35
      cPrint.printTexto 1.5, PosLinea, "Pago por Reembolso de gastos 0% -Corresponde a los códigos 8 y 9"
      cPrint.printTexto 13, PosLinea, "607"
      cPrint.printVariable 15.1, PosLinea, Real6
      PosLinea = PosLinea + 0.35
  End If
 End With
 PosLinea = PosLinea + 0.35
End Sub

Public Sub Vista_Ventas()
'Sumo para cada tipo de comprobante
 Real1 = 0
 Real2 = 0
 Real3 = 0
 
 sSQL = "SELECT TV.TipoComprobante,TCC.Descripcion,COUNT(TipoComprobante) As Cant,SUM(BaseImponible) As BI,SUM(BaseImpGrav) As BIG,SUM(MontoIva) As MI " _
      & "FROM Trans_Ventas As TV,Clientes As C,Tipo_Comprobante As TCC " _
      & "WHERE TV.Fecha Between #" & FechaIni & "# AND #" & FechaFin & "#  "
 If ConSucursal Then
    If Len(No_ATS) > 3 Then sSQL = sSQL & "AND TV.Item NOT IN (" & No_ATS & ") "
 Else
    sSQL = sSQL & "AND TV.Item = '" & NumEmpresa & "' "
 End If
 sSQL = sSQL _
      & "AND TV.Periodo = '" & Periodo_Contable & "' " _
      & "AND TCC.TC = 'TDC' " _
      & "AND TV.TipoComprobante = TCC.Tipo_Comprobante_Codigo  " _
      & "AND TV.IdProv = C.Codigo " _
      & "GROUP BY TV.TipoComprobante,TCC.Descripcion " _
      & "ORDER BY TV.TipoComprobante DESC "
 Select_Adodc AdoAux, sSQL
 With AdoAux.Recordset
  If .RecordCount > 0 Then
      cPrint.tipoDeLetra = TipoArial
      cPrint.tipoNegrilla = True
      cPrint.PorteDeLetra = 11
      cPrint.colorDeLetra = Blanco
      cPrint.printTexto 1.5, PosLinea, "V E N T A S", 11, "C", 20, Plata
      cPrint.colorDeLetra = Negro
      PosLinea = PosLinea + 0.5
      cPrint.PorteDeLetra = 8
      cPrint.printTexto 1.5, PosLinea, "Cód"
      cPrint.printTexto 6, PosLinea, "Transacción"
      cPrint.printTexto 12, PosLinea, "No. Reg."
      cPrint.printTexto 13.6, PosLinea, "BI tarifa 0%"
      cPrint.printTexto 15.6, PosLinea, "BI tarifa 12%"
      cPrint.printTexto 18.1, PosLinea, "Valor IVA"
      cPrint.tipoSubrayado = False
      cPrint.tipoNegrilla = False
      PosLinea = PosLinea + 0.4
      cPrint.printLinea 1.5, PosLinea, 19.5, PosLinea, Negro
      PosLinea = PosLinea + 0.1
      Do While Not .EOF
         cPrint.printTexto 1.5, PosLinea, .fields("TipoComprobante")
         cPrint.printTexto 2, PosLinea, UCaseStrg(.fields("Descripcion"))
         cPrint.printTexto 12, PosLinea, .fields("Cant"), True, 1
         cPrint.printVariable 13.4, PosLinea, .fields("BI")
         cPrint.printVariable 15.5, PosLinea, .fields("BIG")
         cPrint.printVariable 17.7, PosLinea, .fields("MI")
         PosLinea = PosLinea + 0.35
         If .fields("TipoComprobante") = 4 Then
             'MsgBox .Fields("TipoComprobante")
             Real1 = Real1 - .fields("BI")
             Real2 = Real2 - .fields("BIG")
             Real3 = Real3 - .fields("MI")
         Else
             Real1 = Real1 + .fields("BI")
             Real2 = Real2 + .fields("BIG")
             Real3 = Real3 + .fields("MI")
         End If
        .MoveNext
      Loop
      PosLinea = PosLinea + 0.1
      cPrint.printLinea 1.5, PosLinea, 19.5, PosLinea, Negro
      PosLinea = PosLinea + 0.1
      cPrint.printTexto 1.5, PosLinea, "TOTAL"
      cPrint.printVariable 13.4, PosLinea, Real1
      cPrint.printVariable 15.5, PosLinea, Real2
      cPrint.printVariable 17.7, PosLinea, Real3
      PosLinea = PosLinea + 0.5
      cPrint.tipoNegrilla = True
      cPrint.printTexto 1.5, PosLinea, "Se verificará con los casilleros asignados en la declaración de IVA (form.104) de acuerdo al siguiente esquema:"
      PosLinea = PosLinea + 0.35
      cPrint.tipoSubrayado = True
      cPrint.printTexto 1.5, PosLinea, "Sustento Crédito Tributario"
      cPrint.printTexto 12.5, PosLinea, "Casilleros"
      cPrint.printTexto 15.6, PosLinea, "Base"
      cPrint.printTexto 18, PosLinea, "Impuesto"
      PosLinea = PosLinea + 0.35
      cPrint.tipoNegrilla = False
      cPrint.tipoSubrayado = False
  End If
 End With
 
 sSQL = "SELECT SUM(BaseImponible) As BI,SUM(BaseImpGrav) As BIG,SUM(MontoIva) As MI " _
      & "FROM Trans_Ventas As TV,Clientes As C " _
      & "WHERE TV.Fecha Between #" & FechaIni & "# AND #" & FechaFin & "#  "
  If ConSucursal Then
     If Len(No_ATS) > 3 Then sSQL = sSQL & "AND TV.Item NOT IN (" & No_ATS & ") "
  Else
     sSQL = sSQL & "AND TV.Item = '" & NumEmpresa & "' "
  End If

 sSQL = sSQL & "AND TV.Periodo = '" & Periodo_Contable & "' " _
      & "AND TV.IdProv = C.Codigo " _
      & "GROUP BY BaseImponible,BaseImpGrav,MontoIva "
 Select_Adodc AdoAux, sSQL
 Real1 = 0 ' Base Imp 12%
 Real2 = 0 ' IVA 12%
 Real3 = 0 ' 0%
 Real4 = 0 ' Base Imp 12%
 Real5 = 0 ' IVA 12%
 Real6 = 0  ' 0%
 With AdoAux.Recordset
  If .RecordCount > 0 Then
      Do While Not .EOF
         Real1 = Real1 + .fields("BIG")
         Real2 = Real2 + .fields("MI")
         Real3 = Real3 + .fields("BI")
         Real4 = Real4 + .fields("BIG")
         Real5 = Real5 + .fields("MI")
         Real6 = Real6 + .fields("BI")
        .MoveNext
      Loop
      cPrint.printTexto 1.5, PosLinea, "Ventas Netas Base Imponible 12%"
      cPrint.printTexto 12.5, PosLinea, "531+533+535+537"
      cPrint.printVariable 15.5, PosLinea, Real1
      cPrint.printVariable 17.7, PosLinea, Real2
      PosLinea = PosLinea + 0.35
      cPrint.printTexto 1.5, PosLinea, "Ventas Netas Base Imponible 0%"
      cPrint.printTexto 12.5, PosLinea, "501+503+505+507"
      cPrint.printVariable 15.5, PosLinea, Real3
      PosLinea = PosLinea + 0.35
      cPrint.printTexto 1.5, PosLinea, "Ingresos por Reembolso de Gastos Tarifa 12%"
      cPrint.printTexto 12.5, PosLinea, "539"
      cPrint.printVariable 15.5, PosLinea, Real4
      cPrint.printVariable 17.7, PosLinea, Real5
      PosLinea = PosLinea + 0.35
      cPrint.printTexto 1.5, PosLinea, "Ingresos por Reembolso de Gastos Tarifa0%"
      cPrint.printTexto 12.5, PosLinea, "509"
      cPrint.printVariable 15.5, PosLinea, Real6
  End If
 End With
 PosLinea = PosLinea + 0.8
End Sub

Public Sub Vista_Anulados()
Dim Total_A As Integer
 cPrint.tipoDeLetra = TipoArial
 cPrint.tipoNegrilla = True
 cPrint.PorteDeLetra = 11
 cPrint.colorDeLetra = Blanco
 cPrint.printTexto 1.5, PosLinea, "A N U L A D O S", 11, "C", 20, Plata
 cPrint.colorDeLetra = Negro
 cPrint.PorteDeLetra = 8
 PosLinea = PosLinea + 0.5
 cPrint.tipoNegrilla = False
 Total_A = 0
 sSQL = "SELECT COUNT(T) As Cantidad " _
      & "FROM Trans_Anulados " _
      & "WHERE Fecha Between #" & FechaIni & "# AND #" & FechaFin & "#  "
  If ConSucursal Then
     If Len(No_ATS) > 3 Then sSQL = sSQL & "AND Item NOT IN (" & No_ATS & ") "
  Else
     sSQL = sSQL & "AND Item = '" & NumEmpresa & "' "
  End If

 sSQL = sSQL & "AND Periodo = '" & Periodo_Contable & "' " _
      & "GROUP BY T "
 Select_Adodc AdoAux, sSQL
 If AdoAux.Recordset.RecordCount > 0 Then Total_A = AdoAux.Recordset.fields("Cantidad")
 cPrint.printTexto 1.5, PosLinea, "Total de Comprobantes Anulados en el período informado(no incluye los dados de baja)"
 cPrint.printTexto 15.8, PosLinea, "SUMATORIA"
 cPrint.printVariable 19, PosLinea, Total_A
 PosLinea = PosLinea + 0.8
End Sub

Public Sub Vista_Importaciones()
 Real1 = 0
 Real2 = 0
 Real3 = 0
 
 sSQL = "SELECT TI.TipoComprobante,TCC.Descripcion,COUNT(TipoComprobante) As Cant,SUM(ValorCIF) As VC,SUM(MontoIva) As MI " _
      & "FROM Trans_Importaciones As TI,Clientes As C,Tipo_Comprobante As TCC " _
      & "WHERE TI.Fecha Between #" & FechaIni & "# AND #" & FechaFin & "#  "
 If ConSucursal Then
    If Len(No_ATS) > 3 Then sSQL = sSQL & "AND TI.Item NOT IN (" & No_ATS & ") "
 Else
    sSQL = sSQL & "AND TI.Item = '" & NumEmpresa & "' "
 End If
 sSQL = sSQL _
      & "AND TI.Periodo = '" & Periodo_Contable & "' " _
      & "AND TCC.TC = 'TDC' " _
      & "AND TI.TipoComprobante = TCC.Tipo_Comprobante_Codigo  " _
      & "AND TI.IdFiscalProv = C.Codigo " _
      & "GROUP BY TI.TipoComprobante,TCC.Descripcion " _
      & "ORDER BY TI.TipoComprobante "
 Select_Adodc AdoAux, sSQL
 With AdoAux.Recordset
  If .RecordCount > 0 Then
      Segunda_Pag = True
      cPrint.tipoDeLetra = TipoArial
      cPrint.tipoNegrilla = True
      cPrint.PorteDeLetra = 11
      cPrint.colorDeLetra = Blanco
      cPrint.printTexto 1.5, PosLinea, "I M P O R T A C I O N E S", "C", 20, Plata
      cPrint.colorDeLetra = Negro
      PosLinea = PosLinea + 0.5
      cPrint.PorteDeLetra = 8
      cPrint.printTexto 1.2, PosLinea, "Cód"
      cPrint.printTexto 6, PosLinea, "Transacción"
      cPrint.printTexto 12, PosLinea, "No. Reg."
      cPrint.printTexto 15.6, PosLinea, "Valor CIF%"
      cPrint.printTexto 18.1, PosLinea, "Valor IVA"
      PosLinea = PosLinea + 0.4
      cPrint.tipoNegrilla = False
      Do While Not .EOF
         cPrint.tipoNegrilla = False
         cPrint.printTexto 1.5, PosLinea, .fields("TipoComprobante")
         cPrint.printTexto 2.5, PosLinea, UCaseStrg(.fields("Descripcion"))
         cPrint.printVariable 12, PosLinea, .fields("Cant")
         cPrint.printVariable 15.5, PosLinea, .fields("VC")
         cPrint.printVariable 17.7, PosLinea, .fields("MI")
         PosLinea = PosLinea + 0.4
         Real1 = Real1 + .fields("VC")
         Real2 = Real2 + .fields("MI")
        .MoveNext
      Loop
      PosLinea = PosLinea + 0.4
      cPrint.printLinea 1.5, PosLinea, 20, PosLinea, Negro
      PosLinea = PosLinea + 0.05
      cPrint.printTexto 11.5, PosLinea, "TOTAL"
      cPrint.printVariable 15.5, PosLinea, Real1
      cPrint.printVariable 17.7, PosLinea, Real2
      PosLinea = PosLinea + 0.8
      cPrint.printTexto 1.5, PosLinea, "Se verificará con los casilleros asignados en la declaración de IVA (form.104) de acuerdo al siguiente esquema:"
      PosLinea = PosLinea + 0.4
  End If
 End With
 sSQL = "SELECT TI.CodSustento,TCC.Descripcion,COUNT(TipoComprobante) As Cant,SUM(BaseImponible) As BI,SUM(BaseImpGrav) As BIG,SUM(MontoIva) As MI " _
      & "FROM Trans_Importaciones As TI,Clientes As C,Tipo_Tributario As TCC " _
      & "WHERE TI.Fecha Between #" & FechaIni & "# AND #" & FechaFin & "#  "
  If ConSucursal Then
     If Len(No_ATS) > 3 Then sSQL = sSQL & "AND TI.Item NOT IN (" & No_ATS & ") "
  Else
     sSQL = sSQL & "AND TI.Item = '" & NumEmpresa & "' "
  End If

 sSQL = sSQL & "AND TI.Periodo = '" & Periodo_Contable & "' " _
      & "AND TI.CodSustento IN ('01','03','06') " _
      & "AND TI.CodSustento = TCC.Credito_Tributario  " _
      & "AND TI.IdFiscalProv = C.Codigo " _
      & "GROUP BY TI.CodSustento,TCC.Descripcion " _
      & "ORDER BY TI.CodSustento "
 Select_Adodc AdoAux, sSQL
 Real1 = 0 ' Base Imp 12% 1,3,6
 Real2 = 0 ' IVA 12% 1,3,6
 Real3 = 0 ' 0% 1,3,6
 With AdoAux.Recordset
  If .RecordCount > 0 Then
      cPrint.printTexto 1.5, PosLinea, "Sustento Crédito Tributario"
      cPrint.printTexto 13, PosLinea, "Casilleros"
      cPrint.printTexto 15.6, PosLinea, "Base"
      cPrint.printTexto 18, PosLinea, "Impuesto"
      PosLinea = PosLinea + 0.4
      Do While Not .EOF
         Select Case .fields("CodSustento")
           Case "01", "03", "06"
                Real1 = Real1 + .fields("BIG")
                Real2 = Real2 + .fields("MI")
                Real3 = Real3 + .fields("BI")
         End Select
        .MoveNext
      Loop
      cPrint.printTexto 1.5, PosLinea, "Importaciones Netas 12% -Sustento de crédito tributario corresponde a los códigos 1,3 y 6"
      cPrint.printTexto 13, PosLinea, "639+641+643+645"
      cPrint.printTexto 15.5, PosLinea, Format(Real1, "##0.00"), True, 1.9
      cPrint.printTexto 17.7, PosLinea, Format(Real2, "##0.00"), True, 1.7
      PosLinea = PosLinea + 0.4
      cPrint.printTexto 1.5, PosLinea, "Importaciones Netas 0% -Sustento de crédito tributario corresponde a los códigos 1,3 y 6"
      cPrint.printTexto 13, PosLinea, "609+611+613"
      cPrint.printTexto 15.5, PosLinea, Format(Real3, "##0.00"), True, 1.9
      PosLinea = PosLinea + 0.4
  End If
 End With
End Sub

Public Sub Vista_Exportaciones()
'Sumo para cada tipo de comprobante
 Real1 = 0
 
 sSQL = "SELECT TE.TipoComprobante,TCC.Descripcion,COUNT(TE.TipoComprobante) As Cant,SUM(ValorFOB) As VF " _
      & "FROM Trans_Exportaciones As TE,Clientes As C,Tipo_Comprobante As TCC " _
      & "WHERE TE.Fecha Between #" & FechaIni & "# AND #" & FechaFin & "#  "
 If ConSucursal Then
    If Len(No_ATS) > 3 Then sSQL = sSQL & "AND TE.Item NOT IN (" & No_ATS & ") "
 Else
    sSQL = sSQL & "AND TE.Item = '" & NumEmpresa & "' "
 End If
 sSQL = sSQL _
      & "AND TE.Periodo = '" & Periodo_Contable & "' " _
      & "AND TCC.TC = 'TDC' " _
      & "AND TE.TipoComprobante = TCC.Tipo_Comprobante_Codigo  " _
      & "AND TE.IdFiscalProv = C.Codigo " _
      & "GROUP BY TE.TipoComprobante,TCC.Descripcion " _
      & "ORDER BY TE.TipoComprobante "
 Select_Adodc AdoAux, sSQL
 With AdoAux.Recordset
  If .RecordCount > 0 Then
      Segunda_Pag = True
      cPrint.tipoDeLetra = TipoArial
      cPrint.tipoNegrilla = True
      cPrint.PorteDeLetra = 11
      cPrint.colorDeLetra = Blanco
      cPrint.printTexto 1.5, PosLinea, "E X P O R T A C I O N E S", "C", 20, Plata
      cPrint.colorDeLetra = Negro
      PosLinea = PosLinea + 0.5
      cPrint.PorteDeLetra = 8
      cPrint.printTexto 1.5, PosLinea, "Cód"
      cPrint.printTexto 6, PosLinea, "Transacción"
      cPrint.printTexto 12, PosLinea, "No. Reg."
      cPrint.printTexto 15.6, PosLinea, "Valor FOB"
      PosLinea = PosLinea + 0.4
      cPrint.tipoNegrilla = False
      Do While Not .EOF
         cPrint.tipoNegrilla = False
         cPrint.printTexto 1.2, PosLinea, .fields("TipoComprobante")
         cPrint.printTexto 2.5, PosLinea, UCaseStrg(.fields("Descripcion"))
         cPrint.printVariable 12, PosLinea, .fields("Cant")
         cPrint.printVariable 15.5, PosLinea, .fields("VF")
         PosLinea = PosLinea + 0.4
         Real1 = Real1 + .fields("VF")
        .MoveNext
      Loop
      cPrint.printLinea 1.5, PosLinea, 20, PosLinea, Negro
      PosLinea = PosLinea + 0.05
      cPrint.printTexto 1.5, PosLinea, "TOTAL"
      cPrint.printVariable 15.5, PosLinea, Real1
      PosLinea = PosLinea + 0.8
      cPrint.tipoNegrilla = True
      cPrint.printTexto 1.5, PosLinea, "Se verificará con los casilleros asignados en la declaración de IVA (form.104) de acuerdo al siguiente esquema:"
      PosLinea = PosLinea + 0.4
  End If
 End With
 
 sSQL = "SELECT TE.IdFiscalProv,SUM(ValorFOB) As VF " _
      & "FROM Trans_Exportaciones As TE,Clientes As C " _
      & "WHERE TE.Fecha Between #" & FechaIni & "# AND #" & FechaFin & "#  "
 If ConSucursal Then
    If Len(No_ATS) > 3 Then sSQL = sSQL & "AND TE.Item NOT IN (" & No_ATS & ") "
 Else
    sSQL = sSQL & "AND TE.Item = '" & NumEmpresa & "' "
 End If
 sSQL = sSQL & "AND TE.Periodo = '" & Periodo_Contable & "' " _
      & "AND TE.IdFiscalProv = C.Codigo  GROUP BY TE.IdFiscalProv "
 Select_Adodc AdoAux, sSQL
 Real1 = 0 ' Base Imp 12%
 With AdoAux.Recordset
  If .RecordCount > 0 Then
      cPrint.printTexto 1.5, PosLinea, "Sustento Crédito Tributario"
      cPrint.printTexto 13, PosLinea, "Casilleros"
      cPrint.printTexto 15.6, PosLinea, "Base"
      PosLinea = PosLinea + 0.4
      Do While Not .EOF
         Real1 = Real1 + .fields("VF")
        .MoveNext
      Loop
      cPrint.printTexto 1.5, PosLinea, "Exportaciones Netas"
      cPrint.printTexto 13, PosLinea, "511+513"
      cPrint.printTexto 15.5, PosLinea, Format(Real1, "##0.00"), True, 1.9
      PosLinea = PosLinea + 0.4
  End If
 End With
End Sub

Public Sub Vista_EmpresaEmisora()
'Sumo para cada tipo de comprobante
 PosLinea = 1.5
 cPrint.tipoNegrilla = True
 cPrint.PorteDeLetra = 10
 cPrint.printTexto 1.5, PosLinea, "E M P R E S A   E M I S O R A   T A R J E T A   D E   C R E D I T O", "C", 20, Plata
 cPrint.PorteDeLetra = 9
 cPrint.tipoSubrayado = True
 PosLinea = PosLinea + 0.5
 cPrint.printTexto 1.2, PosLinea, "Cód"
 cPrint.printTexto 6, PosLinea, "Transacción"
 cPrint.printTexto 12, PosLinea, "No. Reg."
 cPrint.printTexto 15.6, PosLinea, "Total"
 cPrint.printTexto 17.7, PosLinea, "ValorIVA"
 PosLinea = PosLinea + 0.5
 cPrint.tipoSubrayado = False
 cPrint.printTexto 1.2, PosLinea, "TOTAL"
 PosLinea = PosLinea + 0.4
End Sub

Public Sub Vista_Fondos_Fideicomisos()
'Sumo para cada tipo de comprobante
 cPrint.PorteDeLetra = 10
 cPrint.tipoNegrilla = True
 cPrint.printTexto 1.5, PosLinea, "F O N D O S   Y   F I D E I C O M I S O S", "C", 20, Plata
 cPrint.PorteDeLetra = 9
 cPrint.tipoSubrayado = True
 PosLinea = PosLinea + 0.5
 cPrint.printTexto 1.2, PosLinea, "Cód"
 cPrint.printTexto 6, PosLinea, "Transacción"
 cPrint.printTexto 12, PosLinea, "No. Reg."
 cPrint.printTexto 15.6, PosLinea, "Total"
 cPrint.printTexto 17.7, PosLinea, "ValorIVA"
 cPrint.tipoSubrayado = False
 PosLinea = PosLinea + 0.5
 cPrint.printTexto 1.2, PosLinea, "TOTAL"
 PosLinea = PosLinea + 0.4
End Sub

Public Sub Vista_Retencion_Impuesto_Renta()
 Real1 = 0
 Real2 = 0
 'MsgBox FechaIni & vbCrLf & FechaFin
 sSQL = "SELECT TA.CodRet,CCR.Concepto,COUNT(Concepto) As Cant,SUM(BaseImp) As BI,SUM(ValRet) As VR " _
      & "FROM Trans_Air As TA,Tipo_Concepto_Retencion As CCR " _
      & "WHERE TA.Fecha Between #" & FechaIni & "# AND #" & FechaFin & "#  "
 If ConSucursal Then
    If Len(No_ATS) > 3 Then sSQL = sSQL & "AND TA.Item NOT IN (" & No_ATS & ") "
 Else
    sSQL = sSQL & "AND TA.Item = '" & NumEmpresa & "' "
 End If
 sSQL = sSQL & "AND TA.Periodo = '" & Periodo_Contable & "' " _
      & "AND TA.Tipo_Trans = 'C' " _
      & "AND CCR.Fecha_Inicio <= #" & FechaMid & "# " _
      & "AND CCR.Fecha_Final >= #" & FechaMid & "# " _
      & "AND TA.CodRet = CCR.Codigo  " _
      & "GROUP BY TA.CodRet,CCR.Concepto " _
      & "ORDER BY TA.CodRet "
 Select_Adodc AdoAux, sSQL
 With AdoAux.Recordset
  If .RecordCount > 0 Then
      cPrint.tipoDeLetra = TipoArial
      cPrint.tipoNegrilla = True
      cPrint.PorteDeLetra = 11
      cPrint.colorDeLetra = Blanco
      cPrint.printTexto 1.5, PosLinea, "RESUMEN DE RETENCIONES - AGENTE DE RETENCION", 11, "C", 20, Plata
      cPrint.colorDeLetra = Negro
      PosLinea = PosLinea + 0.5
      cPrint.tipoNegrilla = False
      cPrint.PorteDeLetra = 8
      cPrint.printTexto 1.5, PosLinea, "Cód"
      cPrint.printTexto 2.5, PosLinea, "Concepto de Retención en la Fuente de Impuesto a la Renta"
      cPrint.printTexto 14.3, PosLinea, "No. Reg."
      cPrint.printTexto 16.6, PosLinea, "Base"
      cPrint.printTexto 18.7, PosLinea, "Valor"
      PosLinea = PosLinea + 0.4
      cPrint.printLinea 1.5, PosLinea, 20, PosLinea, Negro
      PosLinea = PosLinea + 0.1
      Do While Not .EOF
         cPrint.printTexto 1.5, PosLinea, .fields("CodRet")
         PosLinea = cPrint.printTextoMultiple(2.5, PosLinea, UCaseStrg(.fields("Concepto")), 11.5)
         cPrint.printFields 13.5, PosLinea, .fields("Cant")
         cPrint.printFields 15.5, PosLinea, .fields("BI")
         cPrint.printFields 17.7, PosLinea, .fields("VR")
         PosLinea = PosLinea + 0.35
         Real1 = Real1 + .fields("BI")
         Real2 = Real2 + .fields("VR")
        .MoveNext
      Loop
      PosLinea = PosLinea + 0.1
      cPrint.printLinea 1.5, PosLinea, 20, PosLinea, Negro
      cPrint.printTexto 11.5, PosLinea, "TOTAL"
      cPrint.printVariable 15.5, PosLinea, Real1
      cPrint.printVariable 17.7, PosLinea, Real2
  End If
 End With
 PosLinea = PosLinea + 0.4
 End Sub

Public Sub Vista_Retencion_Fuente_Iva()
 Real1 = 0
 Real2 = 0
'COMPRAS
 sSQL = "SELECT PorRetBienes,PorRetServicios,COUNT(PorRetBienes) As Cant, " _
      & "SUM(MontoIvaBienes) As BIB,SUM(MontoIvaServicios) As BIS, " _
      & "SUM(ValorRetBienes)As VRB,SUM(ValorRetServicios)As VRS " _
      & "FROM Trans_Compras " _
      & "WHERE Fecha Between #" & FechaIni & "# AND #" & FechaFin & "#  "
 If ConSucursal Then
    If Len(No_ATS) > 3 Then sSQL = sSQL & "AND Item NOT IN (" & No_ATS & ") "
 Else
    sSQL = sSQL & "AND Item = '" & NumEmpresa & "' "
 End If
 sSQL = sSQL _
      & "AND Periodo = '" & Periodo_Contable & "' " _
      & "AND (PorRetBienes + PorRetServicios) > 0 " _
      & "GROUP BY PorRetBienes,PorRetServicios " _
      & "ORDER BY PorRetBienes,PorRetServicios "
 Select_Adodc AdoAux, sSQL
 With AdoAux.Recordset
  If .RecordCount > 0 Then
      cPrint.PorteDeLetra = 8
      cPrint.printTexto 1.5, PosLinea, "Operación: COMPRAS"
      cPrint.printTexto 4.7, PosLinea, "Concepto de Retención en la Fuente del IVA"
      cPrint.printTexto 13, PosLinea, "No. Reg."
      cPrint.printTexto 16, PosLinea, "BaseI"
      cPrint.printTexto 17.2, PosLinea, "% Ret."
      cPrint.printTexto 18.7, PosLinea, "Valor"
      cPrint.tipoSubrayado = False
      PosLinea = PosLinea + 0.4
      cPrint.printLinea 1.5, PosLinea, 20, PosLinea, Negro
      PosLinea = PosLinea + 0.1
      Do While Not .EOF
         Select Case .fields("PorRetBienes")
           Case 1
                cPrint.printTexto 3, PosLinea, "Retención en Bienes"
                cPrint.printVariable 12, PosLinea, .fields("Cant")
                cPrint.printVariable 14, PosLinea, .fields("BIB")
                cPrint.printVariable 17.7, PosLinea, .fields("VRB")
                cPrint.printTexto 17.3, PosLinea, "30%"
                Real1 = Real1 + .fields("VRB")
                PosLinea = PosLinea + 0.35
           Case 3
                cPrint.printTexto 3, PosLinea, "Retención en Bienes"
                cPrint.printVariable 12, PosLinea, .fields("Cant")
                cPrint.printVariable 14, PosLinea, .fields("BIB")
                cPrint.printVariable 17.7, PosLinea, .fields("VRB")
                cPrint.printTexto 17.3, PosLinea, "100%"
                Real2 = Real2 + .fields("VRB")
                PosLinea = PosLinea + 0.35
         End Select
         Select Case .fields("PorRetServicios")
           Case 2
                cPrint.printTexto 3, PosLinea, "Retención en Servicios"
                cPrint.printVariable 12, PosLinea, .fields("Cant")
                cPrint.printVariable 14, PosLinea, .fields("BIS")
                cPrint.printVariable 17.7, PosLinea, .fields("VRS")
                cPrint.printTexto 17.3, PosLinea, "70%"
                Real3 = Real3 + .fields("VRS")
                PosLinea = PosLinea + 0.35
           Case 4
                cPrint.printTexto 3, PosLinea, "Retención en Servicios"
                cPrint.printVariable 12, PosLinea, .fields("Cant")
                cPrint.printVariable 14, PosLinea, .fields("BIS")
                cPrint.printVariable 17.7, PosLinea, .fields("VRS")
                cPrint.printTexto 17.3, PosLinea, "70%"
                PosLinea = PosLinea + 0.35
                Real4 = Real4 + .fields("VRS")
                PosLinea = PosLinea + 0.35
           Case 3
                cPrint.printTexto 3, PosLinea, "Retención en Servicios"
                cPrint.printVariable 12, PosLinea, .fields("Cant")
                cPrint.printVariable 14, PosLinea, .fields("BIS")
                cPrint.printVariable 17.7, PosLinea, .fields("VRS")
                cPrint.printTexto 17.3, PosLinea, "100%"
                Real2 = Real2 + .fields("VRS")
                PosLinea = PosLinea + 0.35
         End Select
         Real5 = Real5 + .fields("BIB") + .fields("BIS")
         Total = Real1 + Real2 + Real3 + Real4
        .MoveNext
      Loop
      PosLinea = PosLinea + 0.1
      cPrint.printLinea 1.5, PosLinea, 20, PosLinea, Negro
      cPrint.printTexto 10.5, PosLinea, "TOTAL"
      cPrint.printVariable 14, PosLinea, Real5
      cPrint.printVariable 17.7, PosLinea, Total
      PosLinea = PosLinea + 0.4
  End If
 End With
 'VENTAS
 Real1 = 0: Real2 = 0: Real3 = 0: Real4 = 0: Real5 = 0: Total = 0
 
 sSQL = "SELECT PorRetBienes,PorRetServicios, COUNT(PorRetBienes) As Cant, SUM(PorRetServicios) As Cant1, " _
      & "SUM(MontoIvaBienes) As BIB, SUM(MontoIvaServicios) As BIS, " _
      & "SUM(ValorRetBienes) As VRB, SUM(ValorRetServicios) As VRS " _
      & "FROM Trans_Ventas " _
      & "WHERE Fecha Between #" & FechaIni & "# AND #" & FechaFin & "#  "
  If ConSucursal Then
     If Len(No_ATS) > 3 Then sSQL = sSQL & "AND Item NOT IN (" & No_ATS & ") "
  Else
     sSQL = sSQL & "AND Item = '" & NumEmpresa & "' "
  End If

 sSQL = sSQL _
      & "AND Periodo = '" & Periodo_Contable & "' " _
      & "AND (PorRetBienes + PorRetServicios) > 0 " _
      & "GROUP BY PorRetBienes,PorRetServicios " _
      & "ORDER BY PorRetBienes,PorRetServicios "
 Select_Adodc AdoAux, sSQL
 With AdoAux.Recordset
  If .RecordCount > 0 Then
      PosLinea = PosLinea + 0.5
      cPrint.PorteDeLetra = 8
      cPrint.printTexto 1.5, PosLinea, "Operación: VENTAS"
      cPrint.printTexto 4.7, PosLinea, "Resumen de Retenciones que le efectuaron en el Periodo"
      cPrint.printTexto 14, PosLinea, "No. Reg."
      cPrint.printTexto 16, PosLinea, "BaseI"
      cPrint.printTexto 17.2, PosLinea, "% Ret."
      cPrint.printTexto 18.7, PosLinea, "Valor"
      PosLinea = PosLinea + 0.4
      cPrint.printLinea 1.5, PosLinea, 20, PosLinea, Negro
      PosLinea = PosLinea + 0.1
      Do While Not .EOF
        'MsgBox .Fields("PorRetBienes") & vbCrLf & .Fields("PorRetServicios")
         cPrint.printTexto 1.5, PosLinea, "VENTAS"
         cPrint.printVariable 13, PosLinea, .fields("Cant")
         Saldo = .fields("BIB") + .fields("BIS")
         cPrint.printVariable 14.5, PosLinea, Saldo
        'MsgBox .Fields("PorRetBienes") & vbCrLf & .Fields("PorRetServicios")
         Select Case .fields("PorRetBienes")
           Case 1
                cPrint.printTexto 3, PosLinea, "Retención en Bienes"
                cPrint.printTexto 17.3, PosLinea, "30%"
                Real1 = Real1 + .fields("VRB")
           Case 3
                cPrint.printTexto 3, PosLinea, "Retención en Bienes"
                cPrint.printTexto 17.3, PosLinea, "100%"
                Real2 = Real2 + .fields("VRB")
         End Select
         Select Case .fields("PorRetServicios")
           Case 2
                cPrint.printTexto 3, PosLinea, "Retención en Servicios"
                cPrint.printTexto 17.3, PosLinea, "70%"
                Real3 = Real3 + .fields("VRS")
           Case 3
                cPrint.printTexto 3, PosLinea, "Retención en Servicios"
                cPrint.printTexto 17.3, PosLinea, "100%"
                Real4 = Real4 + .fields("VRS")
         End Select
         cPrint.printVariable 17.7, PosLinea, .fields("VRB")
         Real5 = Real5 + .fields("BIB") + .fields("BIS")
         Total = Real1 + Real2 + Real3 + Real4
         PosLinea = PosLinea + 0.35
        .MoveNext
      Loop
      PosLinea = PosLinea + 0.1
      cPrint.printLinea 1.5, PosLinea, 20, PosLinea, Negro
      cPrint.printTexto 10.5, PosLinea, "TOTAL"
      cPrint.printTexto 15, PosLinea, Format(Real5, "##0.00"), True, 1.9
      cPrint.printTexto 17.7, PosLinea, Format(Total, "##0.00"), True, 1.7
      PosLinea = PosLinea + 0.4
  End If
 End With
 
'Retenciones en la Fuente que le efectuaron
 sSQL = "SELECT TA.CodRet, TR.Concepto, TR.Porcentaje, COUNT(TA.CodRet) As Cant, SUM(TA.BaseImp) As BI, SUM(TA.ValRet) As VR " _
      & "FROM Trans_Air As TA,Tipo_Concepto_Retencion As TR " _
      & "WHERE TA.Fecha Between #" & FechaIni & "# AND #" & FechaFin & "#  "
 If ConSucursal Then
    If Len(No_ATS) > 3 Then sSQL = sSQL & "AND TA.Item NOT IN (" & No_ATS & ") "
 Else
    sSQL = sSQL & "AND TA.Item = '" & NumEmpresa & "' "
 End If
 sSQL = sSQL & "AND TA.Periodo = '" & Periodo_Contable & "' " _
      & "AND TA.Tipo_Trans = 'V' " _
      & "AND TR.Fecha_Inicio <= #" & FechaIni & "# " _
      & "AND TR.Fecha_Final >= #" & FechaIni & "# " _
      & "AND TA.CodRet = TR.Codigo " _
      & "GROUP BY TA.CodRet,TR.Concepto,TR.Porcentaje " _
      & "ORDER BY TA.CodRet,TR.Concepto,TR.Porcentaje "
 Select_Adodc AdoAux, sSQL
 With AdoAux.Recordset
  If .RecordCount > 0 Then
      PosLinea = PosLinea + 0.4
      cPrint.tipoDeLetra = TipoArial
      cPrint.tipoNegrilla = True
      cPrint.PorteDeLetra = 11
      cPrint.colorDeLetra = Blanco
      cPrint.printTexto 1.5, PosLinea, "RESUMEN DE RETENCIONES EN LA FUENTE QUE LE EFECTUARON", 11, "C", 20, Plata
      cPrint.colorDeLetra = Negro
      PosLinea = PosLinea + 0.5
      cPrint.tipoNegrilla = False
      cPrint.PorteDeLetra = 8
      PosLinea = PosLinea + 0.1
      cPrint.printTexto 1.5, PosLinea, "Codigo"
      cPrint.printTexto 2.7, PosLinea, "Concepto de Retención efectuadas en Ventas"
      cPrint.printTexto 14, PosLinea, "No. Reg."
      cPrint.printTexto 16, PosLinea, "BaseI"
      cPrint.printTexto 17.3, PosLinea, "% Ret."
      cPrint.printTexto 18.7, PosLinea, "Valor"
      PosLinea = PosLinea + 0.4
      cPrint.printLinea 1.5, PosLinea, 19.5, PosLinea, Negro
      PosLinea = PosLinea + 0.1
      Real1 = 0
      Real2 = 0
      Do While Not .EOF
         cPrint.printTexto 1.5, PosLinea, .fields("CodRet")
         PosLinea = cPrint.printTextoMultiple(2.7, PosLinea, .fields("Concepto"), 11.5)
         cPrint.printVariable 13.5, PosLinea, .fields("Cant")
         cPrint.printVariable 14.7, PosLinea, .fields("BI")
         cPrint.printTexto 17.5, PosLinea, CStr(.fields("Porcentaje")) & "%"
         cPrint.printVariable 17.7, PosLinea, .fields("VR")
         PosLinea = PosLinea + 0.35
         Real1 = Real1 + .fields("BI")
         Real2 = Real2 + .fields("VR")
        .MoveNext
      Loop
      PosLinea = PosLinea + 0.1
      cPrint.printLinea 1.5, PosLinea, 19.5, PosLinea, Negro
      PosLinea = PosLinea + 0.05
      cPrint.printTexto 11.5, PosLinea, "TOTAL"
      cPrint.printVariable 15, PosLinea, Real1
      cPrint.printVariable 17.9, PosLinea, Real2
      PosLinea = PosLinea + 0.4
  End If
 End With
 End Sub

Public Sub Vista_Final()
 cPrint.PorteDeLetra = 8
 PosLinea = PosLinea + 0.4
 Cadena = "Declaro que los datos contenidos en este anexo son verdaderos, " _
        & "por lo que asumo la responsabilidad correspondiente, de acuerdo a " _
        & "lo establecido en el Art.101 de la Codificación de la Ley de Régimen Tributario Interno"
 PosLinea = cPrint.printTextoMultiple(1.2, PosLinea, Cadena, 19)
 PosLinea = PosLinea + 1.5
 cPrint.printTexto 1.3, PosLinea, "_________________________"
 cPrint.printTexto 15.6, PosLinea, "_________________________"
 PosLinea = PosLinea + 0.4
 cPrint.printTexto 1.8, PosLinea, "Firma del Contador"
 cPrint.printTexto 15.9, PosLinea, "Firma del representante"
 PosLinea = PosLinea + 0.5
 cPrint.printTexto 1.8, PosLinea, RUC_Contador
 cPrint.printTexto 15.9, PosLinea, CI_Representante
 PosLinea = PosLinea + 0.4
End Sub

Public Sub Consultar_Anexos()
  Dim FechaIniDep As String
  Dim FechaFinDep As String
  Dim CodCiudad As String
  Dim Cont_RDEP As Integer
  
  Anio_Anexo = Toolbar1.buttons(2).Caption
  Mes_Anexo = Toolbar1.buttons(3).Caption
  Fecha_Del_AT Mes_Anexo, Anio_Anexo
  
 'Asignar el contador de las Transacciones
'  Numerar_Lineas_SRI "Trans_Anulados", ""
'  Numerar_Lineas_SRI "Trans_Compras", "C"
  'Numerar_Lineas_SRI "Trans_Ventas", "V"
'  Numerar_Lineas_SRI "Trans_Importaciones", "I"
'  Numerar_Lineas_SRI "Trans_Exportaciones", "E"
      
  If Periodo_Contable = "" Then Periodo_Contable = Ninguno
 'CLIENTES
  sSQL = "SELECT * " _
       & "FROM Clientes " _
       & "WHERE Codigo <> '.' " _
       & "ORDER BY Cliente "
  Select_Adodc AdoClientes, sSQL
  
 'ANULADOS
  sSQL = "SELECT * " _
       & "FROM Trans_Anulados " _
       & "WHERE FechaAnulacion Between #" & FechaIni & "# AND #" & FechaFin & "# " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND Item IN (" & sItem & ") " _
       & "ORDER BY Linea_SRI,FechaAnulacion "
  Select_Adodc AdoAnulados, sSQL
  
 'COMPRAS
  sSQL = "SELECT C.Cliente, C.Codigo, C.CI_RUC, C.TD, C.Tipo_Pasaporte, C.Parte_Relacionada, TC.* " _
       & "FROM Trans_Compras As TC, Clientes As C " _
       & "WHERE TC.Fecha Between #" & FechaIni & "# AND #" & FechaFin & "# " _
       & "AND TC.Periodo = '" & Periodo_Contable & "' " _
       & "AND TC.Item IN (" & sItem & ") " _
       & "AND TC.IdProv = C.Codigo " _
       & "ORDER BY TC.Linea_SRI,C.Cliente, C.CI_RUC, C.TD "
  Select_Adodc AdoCompras, sSQL
  
 'IMPORTACIONES
  sSQL = "SELECT C.Cliente, C.Codigo, C.CI_RUC, C.TD,TI.* " _
       & "FROM Trans_Importaciones As TI, Clientes As C " _
       & "WHERE TI.Fecha Between #" & FechaIni & "# AND #" & FechaFin & "#  " _
       & "AND TI.Periodo = '" & Periodo_Contable & "' " _
       & "AND TI.Item IN (" & sItem & ") " _
       & "AND TI.IdFiscalProv = C.Codigo  " _
       & "ORDER BY TI.Linea_SRI,C.Cliente, C.CI_RUC, C.TD "
  Select_Adodc AdoImportaciones, sSQL
 
 'EXPORTACIONES
  sSQL = "SELECT C.Cliente, C.Codigo, C.CI_RUC, C.TD,TE.* " _
       & "FROM Trans_Exportaciones As TE, Clientes As C " _
       & "WHERE TE.Fecha Between #" & FechaIni & "# AND #" & FechaFin & "#  " _
       & "AND TE.Periodo = '" & Periodo_Contable & "' " _
       & "AND TE.Item IN (" & sItem & ") " _
       & "AND TE.IdFiscalProv = C.Codigo  " _
       & "ORDER BY TE.Linea_SRI,C.Cliente, C.CI_RUC, C.TD "
  Select_Adodc AdoExportaciones, sSQL
  
  FechaIniDep = BuscarFecha("01/01/" & Year(FechaFinal))
  FechaFinDep = BuscarFecha("31/12/" & Year(FechaFinal))
    
 'RELACION POR DEPENDENCIA
  SQL2 = "SELECT C.Codigo,C.Cliente,TD.Subtotal,TD.Item,TD.Periodo " _
       & "FROM Catalogo_Rol_Pagos As TD, Clientes As C " _
       & "WHERE TD.Fecha_107 Between #" & FechaIniDep & "# AND #" & FechaFinDep & "# " _
       & "AND TD.Periodo = '" & Periodo_Contable & "' " _
       & "AND TD.Item IN (" & sItem & ") " _
       & "AND (TD.SuelSal+TD.SobSuelComRemu) > 0 " _
       & "AND TD.Codigo = C.Codigo " _
       & "ORDER BY C.Cliente "
  Select_Adodc_Grid DGRolPagos, AdoRolPagos, SQL2
  With AdoRolPagos.Recordset
   If .RecordCount > 0 Then
       Cont_RDEP = 0
       Do While Not .EOF
          CodigoCli = .fields("Codigo")
          Cont_RDEP = Cont_RDEP + 1
         'Redondeamos los campos
          sSQL = "UPDATE Catalogo_Rol_Pagos " _
               & "SET Linea_SRI = " & Cont_RDEP & " " _
               & "WHERE Fecha Between #" & FechaIniDep & "# AND #" & FechaFinDep & "# " _
               & "AND Codigo = '" & CodigoCli & "' " _
               & "AND Periodo = '" & Periodo_Contable & "' " _
               & "AND Item IN (" & sItem & ") "
          Ejecutar_SQL_SP sSQL
         'Calculamos el SubTotal
          sSQL = "UPDATE Catalogo_Rol_Pagos " _
               & "SET SubTotal=SuelSal+SobSuelComRemu+DecimTer+DecimCuar-ApoPerIess-RebEspDiscap-RebEspTerEd " _
               & "WHERE Fecha Between #" & FechaIniDep & "# AND #" & FechaFinDep & "# " _
               & "AND Codigo = '" & CodigoCli & "' " _
               & "AND Periodo = '" & Periodo_Contable & "' " _
               & "AND Item IN (" & sItem & ") "
         'Ejecutar_SQL_SP sSQL
         'Calculamos la base imponible
          sSQL = "UPDATE Catalogo_Rol_Pagos " _
               & "SET BasImp = SubTotal-ImpRentEmpl " _
               & "WHERE Fecha Between #" & FechaIniDep & "# AND #" & FechaFinDep & "# " _
               & "AND Codigo = '" & CodigoCli & "' " _
               & "AND Periodo = '" & Periodo_Contable & "' " _
               & "AND Item IN (" & sItem & ") "
          'Ejecutar_SQL_SP sSQL
         .MoveNext
       Loop
   End If
  End With
  SQL2 = "SELECT E.RUC,C.Cliente,C.CI_RUC,C.TD,C.Direccion,C.DirNumero,C.Ciudad,C.Prov,C.Telefono,C.Pais,TD.* " _
       & "FROM Catalogo_Rol_Pagos As TD, Clientes As C, Empresas As E " _
       & "WHERE TD.Fecha_107 Between #" & FechaIniDep & "# AND #" & FechaFinDep & "# " _
       & "AND TD.Periodo = '" & Periodo_Contable & "' " _
       & "AND TD.Item IN (" & sItem & ") " _
       & "AND (TD.SuelSal+TD.SobSuelComRemu) > 0 " _
       & "AND TD.Item = E.Item " _
       & "AND TD.Codigo = C.Codigo " _
       & "ORDER BY TD.Linea_SRI,C.Cliente "
  Select_Adodc_Grid DGRolPagos, AdoRolPagos, SQL2
 'MsgBox SQL2 & vbCrLf & AdoRolPagos.Recordset.RecordCount
 
 'VENTAS
  sSQL = "SELECT X, RUC_CI,TB,Razon_Social,TipoComprobante,RetPresuntiva," _
       & "PorcentajeIva,PorcentajeIce,IvaPresuntivo,PorRetBienes,PorRetServicios," _
       & "SUM(NumeroComprobantes) As NumeroComprobantesV," _
       & "SUM(BaseImponible) As BaseImponibleV," _
       & "SUM(BaseImpGrav) As BaseImpGravV," _
       & "SUM(MontoIva) As MontoIvaV," _
       & "SUM(BaseImpIce) As BaseImpIceV," _
       & "SUM(MontoIce) As MontoIceV," _
       & "SUM(MontoIvaBienes) As MontoIvaBienesV," _
       & "SUM(ValorRetBienes) As ValorRetBienesV," _
       & "SUM(MontoIvaServicios) As MontoIvaServiciosV," _
       & "SUM(ValorRetServicios) As ValorRetServiciosV " _
       & "FROM Trans_Ventas " _
       & "WHERE Fecha Between #" & FechaIni & "# AND #" & FechaFin & "# " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND Item IN (" & sItem & ") " _
       & "GROUP BY X, RUC_CI,TB,Razon_Social,TipoComprobante,RetPresuntiva,PorcentajeIva,PorcentajeIce,PorRetBienes,PorRetServicios,IvaPresuntivo " _
       & "ORDER BY X, RUC_CI,TB,Razon_Social "
  Select_Adodc AdoVentas, sSQL
 
 'PUNTO DE VENTAS
 'TV.PuntoEmision,
  sSQL = "SELECT TipoComprobante,Establecimiento, " _
       & "RetPresuntiva,PorcentajeIva,PorcentajeIce," _
       & "IvaPresuntivo,PorRetBienes,PorRetServicios," _
       & "SUM(NumeroComprobantes) As NumeroComprobantesV," _
       & "SUM(BaseImponible) As BaseImponibleV," _
       & "SUM(BaseImpGrav) As BaseImpGravV," _
       & "SUM(MontoIva) As MontoIvaV," _
       & "SUM(BaseImpIce) As BaseImpIceV," _
       & "SUM(MontoIce) As MontoIceV," _
       & "SUM(MontoIvaBienes) As MontoIvaBienesV," _
       & "SUM(ValorRetBienes) As ValorRetBienesV," _
       & "SUM(MontoIvaServicios) As MontoIvaServiciosV," _
       & "SUM(ValorRetServicios) As ValorRetServiciosV " _
       & "FROM Trans_Ventas " _
       & "WHERE Fecha Between #" & FechaIni & "# AND #" & FechaFin & "# " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND Item IN (" & sItem & ") " _
       & "GROUP BY Establecimiento,PuntoEmision,TipoComprobante,RetPresuntiva,PorcentajeIva,PorcentajeIce,PorRetBienes,PorRetServicios,IvaPresuntivo " _
       & "ORDER BY Establecimiento "
  Select_Adodc AdoPuntosVentas, sSQL
 'MsgBox AdoVentas.Recordset.RecordCount & vbCrLf & AdoPuntosVentas.Recordset.RecordCount
 'LISTA LA BASES DE CIUDADES
  sSQL = "SELECT * " _
       & "FROM Tabla_Naciones " _
       & "WHERE TR = 'C' " _
       & "ORDER BY CProvincia, Descripcion_Rubro "
  Select_Adodc AdoCiudades, sSQL
End Sub

Public Sub Imprimir_Formulario_107(FechaEntrega As String, Optional UnFormulario As Boolean)
On Error GoTo Errorhandler
Mensajes = "Imprimir formularios 107?"
Titulo = "IMPRESION DE FORMULARIOS 107"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
   RatonReloj
   DGRolPagos.Visible = False
   InicioX = 0.5: InicioY = 0
   Escala_Centimetro 1, TipoCourier, 9
   Pagina = 1
   With AdoRolPagos.Recordset
    If .RecordCount Then
        Progreso_Barra.Valor_Maximo = .RecordCount + 1
        Progreso_Esperar
        If UnFormulario Then
           Formulario_107 FechaEntrega
        Else
          .MoveFirst
           Do While Not .EOF
              Formulario_107 FechaEntrega
              Progreso_Esperar
             .MoveNext
           Loop
        End If
    End If
   End With
   DGRolPagos.Visible = True
   RatonNormal
   MensajeEncabData = ""
   Printer.EndDoc
   Exit Sub
Errorhandler:
             RatonNormal
             ErrorDeImpresion
             Exit Sub
Else
   RatonNormal
End If
End Sub

Public Sub Formulario_107(FechaEntrega As String)
Dim AltoLetra As Single
Dim AnchoLetra As Single
On Error GoTo Errorhandler
   RatonReloj
   Imprimir_Formato_Propio "RD", 2, 1.5
   AltoLetra = 0.4
   CRConLineas = ProcesarSeteos("RD")
   With AdoRolPagos.Recordset
    If .RecordCount Then
       'ENCABEZADO DEL EMPLEADO
        Printer.FontSize = SetD(2).Tamaño
        Printer.FontBold = True
        Codigo = CStr(.fields("AñoRet"))
        AnchoLetra = SetD(2).PosX
        For I = 1 To 4
            PrinterTexto AnchoLetra, SetD(2).PosY, MidStrg(Codigo, I, 1)
            AnchoLetra = AnchoLetra + 0.5
        Next I
        Codigo = CStr(FechaEntrega)
        AnchoLetra = SetD(3).PosX
        For I = 1 To Len(Codigo)
            If MidStrg(Codigo, I, 1) <> "/" Then
               PrinterTexto AnchoLetra, SetD(3).PosY, MidStrg(Codigo, I, 1)
               AnchoLetra = AnchoLetra + 0.5
            End If
        Next I
        AnchoLetra = SetD(4).PosX
        For I = 1 To 13
            PrinterTexto AnchoLetra, SetD(4).PosY, MidStrg(RUC, I, 1)
            AnchoLetra = AnchoLetra + 0.5
        Next I
        Printer.FontSize = SetD(5).Tamaño
        PrinterTexto SetD(5).PosX, SetD(5).PosY, UCaseStrg(NombreComercial)
        AnchoLetra = SetD(6).PosX
        For I = 1 To 10
            Printer.FontSize = SetD(6).Tamaño
            PrinterTexto AnchoLetra, SetD(6).PosY, MidStrg(.fields("CI_RUC"), I, 1)
            AnchoLetra = AnchoLetra + 0.5
        Next I
        Printer.FontSize = SetD(7).Tamaño
        PrinterTexto SetD(7).PosX, SetD(7).PosY, .fields("Cliente")
       'LIQUIDACION DE IMPUESTOS
        Printer.FontSize = SetD(8).Tamaño
        PrinterFields SetD(8).PosX, SetD(8).PosY, .fields("SuelSal"), , True
        PrinterFields SetD(9).PosX, SetD(9).PosY, .fields("SobSuelComRemu"), , True
        PrinterFields SetD(10).PosX, SetD(10).PosY, .fields("PartUtil"), , True
        PrinterFields SetD(11).PosX, SetD(11).PosY, .fields("IngOtrosEmp"), , True
        PrinterFields SetD(12).PosX, SetD(12).PosY, .fields("DecimTer"), , True
        PrinterFields SetD(13).PosX, SetD(13).PosY, .fields("DecimCuar"), , True
        PrinterFields SetD(14).PosX, SetD(14).PosY, .fields("FondoReserva"), , True
        PrinterFields SetD(15).PosX, SetD(15).PosY, .fields("Desahucio"), , True
        PrinterFields SetD(16).PosX, SetD(16).PosY, .fields("ApoPerIess"), , True
        PrinterFields SetD(17).PosX, SetD(17).PosY, .fields("Rebajas"), , True
        PrinterFields SetD(18).PosX, SetD(18).PosY, .fields("Vivienda"), , True
        PrinterFields SetD(19).PosX, SetD(19).PosY, .fields("Salud"), , True
        PrinterFields SetD(20).PosX, SetD(20).PosY, .fields("Educacion"), , True
        PrinterFields SetD(21).PosX, SetD(21).PosY, .fields("Alimentacion"), , True
        PrinterFields SetD(22).PosX, SetD(22).PosY, .fields("Vestimenta"), , True
        PrinterFields SetD(23).PosX, SetD(23).PosY, .fields("RebEspDiscap"), , True
        PrinterFields SetD(24).PosX, SetD(24).PosY, .fields("RebEspTerEd"), , True
        PrinterFields SetD(25).PosX, SetD(25).PosY, .fields("ImpRentEmpl"), , True
        PrinterFields SetD(26).PosX, SetD(26).PosY, .fields("BasImp"), , True     ' .Fields("SubTotal")
        PrinterFields SetD(27).PosX, SetD(27).PosY, .fields("ImpRentCaus"), , True
        PrinterFields SetD(28).PosX, SetD(28).PosY, .fields("ImpAnterior"), , True
        PrinterFields SetD(29).PosX, SetD(29).PosY, .fields("ValRet"), , True
        Saldo = 0
        PrinterVariables SetD(30).PosX, SetD(30).PosY, Saldo, True
        Total = .fields("SuelSal") + .fields("SobSuelComRemu") + .fields("PartUtil") + .fields("ImpRentEmpl")
        PrinterVariables SetD(31).PosX, SetD(31).PosY, Total, True
'''        PrinterFields SetD(25).PosX, SetD(36).PosY, .Fields("Numero")
       'CONSOLIDACION DE INGRESOS

'''        PrinterFields SetD(28).PosX, SetD(28).PosY, .Fields("Deduccion"), , True

        AnchoLetra = SetD(34).PosX
        For I = 1 To 13
            PrinterTexto AnchoLetra, SetD(34).PosY, MidStrg(RUC_Contador, I, 1)
            AnchoLetra = AnchoLetra + 0.5
        Next I
        Printer.NewPage
    End If
   End With
   RatonNormal
   Exit Sub
Errorhandler:
             RatonNormal
             ErrorDeImpresion
             Exit Sub
End Sub

''''Private Sub Insertar_Liquidacion_Compras_ATS()
''''    RatonReloj
''''    Trans_No = 96
''''    Ln_SRI = 1
''''    If InStr(1, CStr(FechaFinal), "/") > 0 Then
''''       Numero = CLng(Val(Replace(Format(FechaFinal, "yyyy/MM/dd"), "/", "")))
''''    Else
''''       Numero = CLng(Val(FechaFinal))
''''    End If
''''
'''''''    SQL1 = "DELETE * " _
'''''''         & "FROM Trans_Compras " _
'''''''         & "WHERE TP = 'CD' " _
'''''''         & "AND Numero = " & Numero & " " _
'''''''         & "AND TipoComprobante = 3 " _
'''''''         & "AND Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# "
'''''''    If Not ConSucursal Then sSQL = sSQL & "AND Item = '" & NumEmpresa & "' "
'''''''    SQL1 = SQL1 & "AND Periodo = '" & Periodo_Contable & "' "
'''''''    Ejecutar_SQL_SP SQL1
'''''''
'''''''    SQL1 = "DELETE * " _
'''''''         & "FROM Trans_Air " _
'''''''         & "WHERE TP = 'CD' " _
'''''''         & "AND Numero = " & Numero & " " _
'''''''         & "AND Tipo_Trans = 'C' " _
'''''''         & "AND Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# "
'''''''    If Not ConSucursal Then sSQL = sSQL & "AND Item = '" & NumEmpresa & "' "
'''''''    SQL1 = SQL1 & "AND Periodo = '" & Periodo_Contable & "' "
'''''''    Ejecutar_SQL_SP SQL1
''''
'''''''   'COMPRAS RETENCION CERO COD 332 DEL MES
'''''''    sSQL = "INSERT INTO Trans_Air (T, CodRet, BaseImp, Porcentaje, ValRet, EstabRetencion, PtoEmiRetencion, " _
'''''''         & "SecRetencion, AutRetencion, EstabFactura, PuntoEmiFactura, Factura_No, Item, CodigoU, " _
'''''''         & "Periodo, Cta_Retencion, IdProv, TP, Numero, Fecha, Tipo_Trans, RUC_CI, TB, Razon_Social, IDT) " _
'''''''         & "SELECT 'N' As TT, '332F' As CR, Total_MN, 0 As Porc,0 As VRet, '000' As SE, '000' As SPE, 0 As SRet, " _
'''''''         & "'0' Aut, '000' As EF, '000' As PEF, Factura, Item, '" & CodigoUsuario & "' As CodigoUx, Periodo, '0' As CtaRet, CodigoC, " _
'''''''         & "'CD' As TPx, " & Numero & " As Num, #" & FechaFin & "# As FechaAT, 'C' As TTs, " _
'''''''         & "RUC_CI, TB, Razon_Social, Factura As IDTF " _
'''''''         & "FROM Facturas " _
'''''''         & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
'''''''         & "AND TC = 'LC' "
'''''''    If ConSucursal Then
'''''''       If Len(No_ATS) > 3 Then sSQL = sSQL & "AND Item NOT IN (" & No_ATS & ") "
'''''''    Else
'''''''       sSQL = sSQL & "AND Item = '" & NumEmpresa & "' "
'''''''    End If
'''''''    sSQL = sSQL _
'''''''         & "AND Periodo = '" & Periodo_Contable & "' " _
'''''''         & "AND T <> '" & Anulado & "' " _
'''''''         & "ORDER BY Item, Periodo, RUC_CI, TB, Razon_Social "
'''''''    Ejecutar_SQL_SP sSQL
'''''''    Progreso_Esperar
'''''''
'''''''   'LIQUIDACIONES DE COMPRAS DEL MES
'''''''    sSQL = "INSERT INTO Trans_Compras (T, IdProv, DevIva, CodSustento, TipoComprobante, Establecimiento," _
'''''''         & "PuntoEmision, Secuencial, Autorizacion, FechaEmision, FechaRegistro, FechaCaducidad," _
'''''''         & "TP, Numero, Fecha, BaseImponible, BaseImpGrav, PorcentajeIva, MontoIva, Item, CodigoU, Periodo, " _
'''''''         & "PagoLocExt, PaisEfecPago, AplicConvDobTrib, PagExtSujRetNorLeg, FormaPago, " _
'''''''         & "PorRetBienes, PorRetServicios, MontoIvaBienes, MontoIvaServicios, ValorRetBienes, ValorRetServicios, " _
'''''''         & "BaseImpIce, PorcentajeIce, MontoIce, BaseNoObjIVA) " _
'''''''         & "SELECT 'N' As TT, CodigoC, 'N' As DIVA, '01' As CodS, 3 AS TComp, MidStrg(Serie,1,3) As Estab," _
'''''''         & "MidStrg(Serie,4,3) As PuntoE, Factura, Autorizacion, Fecha As FechaE, Fecha As FechaR, Fecha As FechaC," _
'''''''         & "'CD' As TPx, " & Numero & " As Num, #" & FechaFin & "# As FechaAT, Total_MN, 0 As BIG,2 As Porc, 0 As MontIVA," _
'''''''         & "Item, '" & CodigoUsuario & "' As CodigoUx, Periodo," _
'''''''         & "'01' As PLE, 'NA' PE1, 'NA' PE2, 'NA' PE3,'01' FP, " _
'''''''         & "0 As T1, 0 As T2, 0 As T3, 0 As T4, 0 As T5, 0 As T6,0 As T7, 0 As T8, 0 As T9, 0 As T10 " _
'''''''         & "FROM Facturas " _
'''''''         & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
'''''''         & "AND TC = 'LC' "
'''''''    If ConSucursal Then
'''''''       If Len(No_ATS) > 3 Then sSQL = sSQL & "AND Item NOT IN (" & No_ATS & ") "
'''''''    Else
'''''''       sSQL = sSQL & "AND Item = '" & NumEmpresa & "' "
'''''''    End If
'''''''    sSQL = sSQL _
'''''''         & "AND Periodo = '" & Periodo_Contable & "' " _
'''''''         & "AND T <> '" & Anulado & "' " _
'''''''         & "ORDER BY Item, Periodo, RUC_CI, TB, Razon_Social "
'''''''    Ejecutar_SQL_SP sSQL
''''End Sub

'''Private Sub Insertar_Ventas_ATS()
'''Dim Fecha_a_Numero
'''Dim TTotal As Currency
'''Dim TTotal_IVA As Currency
'''Dim TBaseCero As Currency
'''Dim TBaseDescuento As Currency
'''Dim TBaseGravada As Currency
'''Dim TBaseSubTotal As Currency
'''Dim Inserta_Anulado As Boolean
'''Dim Serie1 As String
'''Dim Serie2 As String
'''
'''    Trans_No = 96
'''    Ln_SRI = 1
'''    RatonReloj
''''''    Progreso_Esperar
''''''    sSQL = "SELECT F.RUC_CI, F.TB, F.Razon_Social, COUNT(TA.Factura) As Cantidad " _
''''''         & "FROM Trans_Abonos As TA, Facturas As F " _
''''''         & "WHERE TA.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
''''''         & "AND TA.T <> 'A' " _
''''''         & "AND TA.Banco = 'NOTA DE CREDITO' "
''''''    If ConSucursal Then
''''''       If Len(No_ATS) > 3 Then sSQL = sSQL & "AND TA.Item NOT IN (" & No_ATS & ") "
''''''    Else
''''''       sSQL = sSQL & "AND TA.Item = '" & NumEmpresa & "' "
''''''    End If
''''''    sSQL = sSQL _
''''''         & "AND TA.Periodo = '" & Periodo_Contable & "' " _
''''''         & "AND F.TC IN ('FA','NV') " _
''''''         & "AND LEN(F.Autorizacion) < 13 " _
''''''         & "AND TA.TP = F.TC " _
''''''         & "AND TA.Serie = F.Serie " _
''''''         & "AND TA.Factura = F.Factura " _
''''''         & "AND TA.CodigoC = F.CodigoC " _
''''''         & "GROUP BY F.RUC_CI, F.TB, F.Razon_Social " _
''''''         & "ORDER BY F.RUC_CI, F.TB, F.Razon_Social "
''''''    Select_Adodc AdoVentas, sSQL
''''''    Progreso_Barra.Valor_Maximo = Progreso_Barra.Valor_Maximo + AdoVentas.Recordset.RecordCount + 1
''''''    Progreso_Esperar
''''''
''''''   'NOTA DE CREDITO DEL MES
''''''    If AdoVentas.Recordset.RecordCount > 0 Then
''''''       Do While Not AdoVentas.Recordset.EOF
''''''          CodigoCliente = AdoVentas.Recordset.Fields("RUC_CI")
''''''          Codigo3 = AdoVentas.Recordset.Fields("TB")
''''''          NombreCliente = AdoVentas.Recordset.Fields("Razon_Social")
''''''          Cantidad = Round(AdoVentas.Recordset.Fields("Cantidad") / 2)
''''''          sSQL = "SELECT F.RUC_CI,TA.Cheque,TA.Serie,SUM(TA.Abono) As Total_NC " _
''''''               & "FROM Trans_Abonos AS TA, Facturas As F " _
''''''               & "WHERE TA.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
''''''               & "AND TA.Banco = 'NOTA DE CREDITO' " _
''''''               & "AND TA.T <> 'A' "
''''''          If ConSucursal Then
''''''             If Len(No_ATS) > 3 Then sSQL = sSQL & "AND TA.Item NOT IN (" & No_ATS & ") "
''''''          Else
''''''             sSQL = sSQL & "AND TA.Item = '" & NumEmpresa & "' "
''''''          End If
''''''          sSQL = sSQL _
''''''               & "AND TA.Periodo = '" & Periodo_Contable & "' " _
''''''               & "AND F.RUC_CI = '" & CodigoCliente & "' " _
''''''               & "AND F.TC IN ('FA','NV') " _
''''''               & "AND LEN(F.Autorizacion)<13 " _
''''''               & "AND TA.TP = F.TC " _
''''''               & "AND TA.Serie = F.Serie " _
''''''               & "AND TA.Factura = F.Factura " _
''''''               & "AND TA.CodigoC = F.CodigoC " _
''''''               & "GROUP BY F.RUC_CI,TA.Cheque,TA.Serie " _
''''''               & "ORDER BY F.RUC_CI,TA.Cheque DESC "
''''''          Select_Adodc AdoAux, sSQL
''''''          With AdoAux.Recordset
''''''           If .RecordCount > 0 Then
''''''              'MsgBox .Fields("RUC_CI") & vbCrLf & .RecordCount
''''''               Total = 0
''''''               TBaseCero = 0
''''''               TTotal_IVA = 0
''''''               TBaseGravada = 0
''''''               TBaseSubTotal = 0
''''''               PorcIVAB = 0
''''''               PorcIVAS = 0
''''''               Total_RetIVAS = 0
''''''               Total_RetIVAB = 0
''''''               Codigo1 = MidStrg(.Fields("Serie"), 1, 3)
''''''               Codigo2 = MidStrg(.Fields("Serie"), 4, 3)
''''''               Do While Not .EOF
''''''                  Select Case .Fields("Cheque")
''''''                    Case "VENTAS"
''''''                         Total = Total + .Fields("Total_NC")
''''''                         TBaseSubTotal = TBaseSubTotal + .Fields("Total_NC")
''''''                    Case "I.V.A."
''''''                         TTotal_IVA = TTotal_IVA + .Fields("Total_NC")
''''''                  End Select
''''''                 .MoveNext
''''''               Loop
''''''               If TTotal_IVA > 0 Then TBaseGravada = Total Else TBaseCero = Total
''''''           End If
''''''          End With
''''''          Insertar_Ventas_NC CodigoCliente, Codigo3, NombreCliente, CLng(Cantidad), TBaseCero, TBaseGravada, TTotal_IVA, TBaseSubTotal, "4"
''''''          Progreso_Esperar
''''''          AdoVentas.Recordset.MoveNext
''''''       Loop
''''''    End If
'''
''''''   'RETENCIONES VENTAS DEL MES
''''''    sSQL = "INSERT INTO Trans_Air (Item, Periodo, Tipo_Trans, Fecha, EstabFactura, PuntoEmiFactura, " _
''''''         & "EstabRetencion, PtoEmiRetencion, SecRetencion, AutRetencion, BaseImp, ValRet, Porcentaje, " _
''''''         & "Factura_No, T, TP, Numero, Linea_SRI, CodigoU, RUC_CI, TB, Razon_Social,CodRet) " _
''''''         & "SELECT TA.Item, TA.Periodo, 'V' As TT, #" & FechaFin & "# As FechaE, " _
''''''         & "MidStrg(TA.Serie,1,3) As Establecimento,MidStrg(TA.Serie,4,3) As Emision,'001' As Est_R, '001' As Punt_R, " _
''''''         & "'0' As SecRet,'.' Aut_Ret,SUM(TA.Base_Imponible) As BaseImp, SUM(TA.Abono) As Abonos, " _
''''''         & "((SUM(TA.Abono)/SUM(TA.Base_Imponible)) * 100) AS Porc,COUNT(F.RUC_CI) As CantF, 'N' As T_T, " _
''''''         & "'CD' As T_P, " & Numero & " As Numero_TP, 0 As L_SRI, '" & CodigoUsuario & "' As Cod_U, " _
''''''         & "F.RUC_CI, F.TB, F.Razon_Social, MidStrg(TA.Banco,20,3) " _
''''''         & "FROM Trans_Abonos As TA, Facturas As F " _
''''''         & "WHERE TA.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
''''''         & "AND TA.TP IN ('FA','NV') " _
''''''         & "AND MidStrg(TA.Banco,1,16) = 'RETENCION FUENTE' "
''''''    If ConSucursal Then
''''''       If Len(No_ATS) > 3 Then sSQL = sSQL & "AND TA.Item NOT IN (" & No_ATS & ") "
''''''    Else
''''''       sSQL = sSQL & "AND TA.Item = '" & NumEmpresa & "' "
''''''    End If
''''''    sSQL = sSQL _
''''''         & "AND TA.Periodo = '" & Periodo_Contable & "' " _
''''''         & "AND F.TC IN ('FA','NV') " _
''''''         & "AND LEN(TA.Autorizacion_R) < 13 " _
''''''         & "AND F.T <> '" & Anulado & "' " _
''''''         & "AND TA.CodigoC = F.CodigoC " _
''''''         & "AND TA.TP = F.TC " _
''''''         & "AND TA.Serie = F.Serie " _
''''''         & "AND TA.Factura = F.Factura " _
''''''         & "AND TA.Item = F.Item " _
''''''         & "AND TA.Periodo = F.Periodo " _
''''''         & "GROUP BY TA.Item, TA.Periodo, F.RUC_CI, F.TB, F.Razon_Social, MidStrg(TA.Serie,1,3), MidStrg(TA.Serie,4,3), MidStrg(TA.Banco,20,3) "
''''''    Ejecutar_SQL_SP sSQL
''''''    Progreso_Esperar
'''
''''''   'RETENCIONES DE TARJETAS VENTAS DEL MES
''''''    sSQL = "INSERT INTO Trans_Air (Item, Periodo, Tipo_Trans, Fecha, EstabFactura, PuntoEmiFactura, " _
''''''         & "EstabRetencion, PtoEmiRetencion, SecRetencion, AutRetencion, BaseImp, ValRet, Porcentaje, " _
''''''         & "Factura_No, T, TP, Numero, Linea_SRI, CodigoU, RUC_CI, TB, Razon_Social, CodRet) " _
''''''         & "SELECT TA.Item, TA.Periodo, 'V' As TT, #" & FechaFin & "# As FechaE, " _
''''''         & "'001' As Est_R, '001' As Punt_R, MidStrg(TA.Serie_R,1,3) As Establecimento,MidStrg(TA.Serie_R,4,3) As Emision, " _
''''''         & "'0' As SecRet,'.' Aut_Ret,SUM(TA.Abono/1.12) As BaseImp, SUM(TA.IRF_Ret) As Abonos, " _
''''''         & "Porc_Ret,COUNT(CI_RUC) As CantF, 'N' As T_T, " _
''''''         & "'CD' As T_P, " & Numero & " As Numero_TP, 0 As L_SRI, '" & CodigoUsuario & "' As Cod_U, " _
''''''         & "C.CI_RUC, C.TD, C.Cliente, '344' As RetFuente " _
''''''         & "FROM Trans_Abonos As TA, Clientes As C " _
''''''         & "WHERE TA.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
''''''         & "AND TA.TP IN ('FA','NV') " _
''''''         & "AND LEN(TA.Autorizacion) < 13 " _
''''''         & "AND Tipo_Cta = 'TJ' " _
''''''         & "AND Porc_Ret < 10 "
''''''    If ConSucursal Then
''''''       If Len(No_ATS) > 3 Then sSQL = sSQL & "AND TA.Item NOT IN (" & No_ATS & ") "
''''''    Else
''''''       sSQL = sSQL & "AND TA.Item = '" & NumEmpresa & "' "
''''''    End If
''''''    sSQL = sSQL _
''''''         & "AND TA.Periodo = '" & Periodo_Contable & "' " _
''''''         & "AND TA.TP IN ('FA','NV') " _
''''''         & "AND TA.T <> '" & Anulado & "' " _
''''''         & "AND TA.Codigo_Prov = C.Codigo " _
''''''         & "GROUP BY TA.Item, TA.Periodo, TA.Porc_Ret, C.CI_RUC, C.TD, C.Cliente, MidStrg(TA.Serie_R,1,3), MidStrg(TA.Serie_R,4,3) "
''''''    Ejecutar_SQL_SP sSQL
''''''    Progreso_Esperar
'''
'''''     SetAdoFields "Factura_No", FacturaNo
'''''     SetAdoFields "Cta_Retencion", CuentRet
'''
''''ROUND((Total_Ret_IVA_B/IVA)*100,0,0) As Por_B,ROUND((Total_Ret_IVA_S/IVA)*100,0,0) As Por_S
'''
''''''   'VENTAS DEL MES
''''''    sSQL = "INSERT INTO Trans_Ventas (Item, Periodo, RUC_CI, TB, Razon_Social, TipoComprobante, " _
''''''         & "FechaRegistro, FechaEmision, Fecha, Establecimiento, PuntoEmision, NumeroComprobantes, " _
''''''         & "BaseImponible, BaseImpGrav, MontoIva, ValorRetBienes, ValorRetServicios, IvaPresuntivo, " _
''''''         & "PorcentajeIva, RetPresuntiva, T, TP, Numero, Linea_SRI, CodigoU, Secuencial) " _
''''''         & "SELECT Item, Periodo, RUC_CI, TB, Razon_Social,'18' As TComp, " _
''''''         & "#" & FechaFin & "# As FechaR, #" & FechaFin & "# As FechaE, #" & FechaFin & "# As FechaC, " _
''''''         & "MidStrg(Serie,1,3) As Establecimento,MidStrg(Serie,4,3) As Emision,COUNT(Factura) As NumComp, " _
''''''         & "SUM(Sin_IVA-Desc_0) As Base_0, SUM(Con_IVA-Desc_X) As Base_X, SUM(IVA) As T_IVA, " _
''''''         & "SUM(Total_Ret_IVA_B) As Total_IB,SUM(Total_Ret_IVA_S) As Total_IS, 'N' As IVAPres, " _
''''''         & "2 As PorcIVA, 'S' RetPres, 'N' As T_T, 'CD' As TP, " & Numero & " As Numero_TP, " _
''''''         & "(ROW_NUMBER() OVER(ORDER BY Razon_Social)) As L_SRI, " _
''''''         & "'" & CodigoUsuario & "' As Cod_U, 0 As Secuen " _
''''''         & "FROM Facturas " _
''''''         & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
''''''         & "AND TC IN ('FA','NV') "
''''''    If ConSucursal Then
''''''       If Len(No_ATS) > 3 Then sSQL = sSQL & "AND Item NOT IN (" & No_ATS & ") "
''''''    Else
''''''       sSQL = sSQL & "AND Item = '" & NumEmpresa & "' "
''''''    End If
''''''    sSQL = sSQL _
''''''         & "AND Periodo = '" & Periodo_Contable & "' " _
''''''         & "AND LEN(Autorizacion) < 13 " _
''''''         & "AND T <> '" & Anulado & "' " _
''''''         & "GROUP BY Item, Periodo, RUC_CI, TB, Razon_Social, MidStrg(Serie,1,3),MidStrg(Serie,4,3) "
''''''    Ejecutar_SQL_SP sSQL
''''''
''''''    Progreso_Esperar
'''
''''''   'Actualizamos la linea del ATS en Retenciones
''''''    If SQL_Server Then
''''''       sSQL = "UPDATE Trans_Air " _
''''''            & "SET Linea_SRI = TV.Linea_SRI " _
''''''            & "FROM Trans_Air As TA, Trans_Ventas As TV "
''''''    Else
''''''       sSQL = "UPDATE Trans_Air As TV, Trans_Ventas As TV " _
''''''            & "SET TA.Linea_SRI = TV.Linea_SRI "
''''''    End If
''''''    sSQL = sSQL _
''''''         & "WHERE TA.Periodo = '" & Periodo_Contable & "' " _
''''''         & "AND TA.Item = '" & NumEmpresa & "' " _
''''''         & "AND TA.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
''''''         & "AND TA.Item = TV.Item " _
''''''         & "AND TA.Periodo = TV.Periodo " _
''''''         & "AND TA.RUC_CI = TV.RUC_CI " _
''''''         & "AND TA.Fecha = TV.Fecha " _
''''''         & "AND TA.TP = TV.TP " _
''''''         & "AND TA.Numero = TV.Numero "
''''''    Ejecutar_SQL_SP sSQL
''''''    Progreso_Esperar
'''
'''''    Progreso_Esperar
''''PorRetBienes, PorRetServicios
'''
''''''    sSQL = "UPDATE Trans_Ventas " _
''''''         & "SET BaseImpIce = 0, PorcentajeIce = 0, MontoIce = 0, MontoIvaBienes = 0, MontoIvaServicios = 0, " _
''''''         & "Autorizacion = '.', PorRetBienes = 0, PorRetServicios = 0 " _
''''''         & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
''''''         & "AND Periodo = '" & Periodo_Contable & "' " _
''''''         & "AND Item = '" & NumEmpresa & "' "
''''''    Ejecutar_SQL_SP sSQL
''''''    Progreso_Esperar
''''''
''''''    sSQL = "UPDATE Trans_Ventas " _
''''''         & "SET PorRetBienes = ROUND((ValorRetBienes/MontoIva)*100,0,0), " _
''''''         & "PorRetServicios = ROUND((ValorRetServicios/MontoIva)*100,0,0) " _
''''''         & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
''''''         & "AND Periodo = '" & Periodo_Contable & "' " _
''''''         & "AND Item = '" & NumEmpresa & "' " _
''''''         & "AND MontoIva <> 0 "
''''''    Ejecutar_SQL_SP sSQL
''''''    Progreso_Esperar
''''''
''''''    sSQL = "UPDATE Trans_Ventas " _
''''''         & "SET MontoIvaBienes = ROUND((ValorRetBienes*100)/PorRetBienes,2,0) " _
''''''         & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
''''''         & "AND Periodo = '" & Periodo_Contable & "' " _
''''''         & "AND Item = '" & NumEmpresa & "' " _
''''''         & "AND PorRetBienes <> 0 "
''''''    Ejecutar_SQL_SP sSQL
''''''    Progreso_Esperar
''''''
''''''    sSQL = "UPDATE Trans_Ventas " _
''''''         & "SET MontoIvaServicios = ROUND((ValorRetServicios*100)/PorRetServicios,2,0) " _
''''''         & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
''''''         & "AND Periodo = '" & Periodo_Contable & "' " _
''''''         & "AND Item = '" & NumEmpresa & "' " _
''''''         & "AND PorRetServicios <> 0 "
''''''    Ejecutar_SQL_SP sSQL
''''''    Progreso_Esperar
'''
''''''   'Actualizamos el codigo del Beneficiario
''''''    If SQL_Server Then
''''''       sSQL = "UPDATE Trans_Ventas " _
''''''            & "SET IdProv = F.CodigoC " _
''''''            & "FROM Trans_Ventas As TV, Facturas As F "
''''''    Else
''''''       sSQL = "UPDATE Trans_Ventas As TV, Facturas As F " _
''''''            & "SET TV.IdProv = F.CodigoC "
''''''    End If
''''''    sSQL = sSQL _
''''''         & "WHERE TV.Periodo = '" & Periodo_Contable & "' " _
''''''         & "AND TV.Item = '" & NumEmpresa & "' " _
''''''         & "AND TV.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
''''''         & "AND TV.Item = F.Item " _
''''''         & "AND TV.Periodo = F.Periodo " _
''''''         & "AND TV.RUC_CI = F.RUC_CI "
''''''    Ejecutar_SQL_SP sSQL
''''''    Progreso_Esperar
''''''
''''''    If SQL_Server Then
''''''       sSQL = "UPDATE Trans_Air " _
''''''            & "SET IdProv = F.Codig 1oC " _
''''''            & "FROM Trans_Air As TA, Facturas As F "
''''''    Else
''''''       sSQL = "UPDATE Trans_Air As TA, Facturas As F " _
''''''            & "SET TA.IdProv = F.CodigoC "
''''''    End If
''''''    sSQL = sSQL _
''''''         & "WHERE TA.Periodo = '" & Periodo_Contable & "' " _
''''''         & "AND TA.Item = '" & NumEmpresa & "' " _
''''''         & "AND TA.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
''''''         & "AND TA.Tipo_Trans = 'V' " _
''''''         & "AND TA.Item = F.Item " _
''''''         & "AND TA.Periodo = F.Periodo " _
''''''         & "AND TA.RUC_CI = F.RUC_CI "
''''''    Ejecutar_SQL_SP sSQL
''''''    Progreso_Esperar
'''
'''   'FACTURAS ANULADAS
''''''    Numero = CLng(Val(Replace(Format(FechaFinal, "yyyy/MM/dd"), "/", "")))
''''''    SQL1 = "DELETE * " _
''''''         & "FROM Trans_Anulados " _
''''''         & "WHERE TP = 'CD' " _
''''''         & "AND Numero = " & Numero & " " _
''''''         & "AND Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
''''''         & "AND Periodo = '" & Periodo_Contable & "' " _
''''''         & "AND LEN(Autorizacion)<13 "
''''''    If ConSucursal Then
''''''       If Len(No_ATS) > 3 Then sSQL = sSQL & "AND Item NOT IN (" & No_ATS & ") "
''''''    Else
''''''       sSQL = sSQL & "AND Item = '" & NumEmpresa & "' "
''''''    End If
''''''    Ejecutar_SQL_SP SQL1
'''
'''    Progreso_Esperar
'''
'''    sSQL = "SELECT * " _
'''         & "FROM Facturas " _
'''         & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# "
'''   If ConSucursal Then
'''      If Len(No_ATS) > 3 Then sSQL = sSQL & "AND Item NOT IN (" & No_ATS & ") "
'''   Else
'''      sSQL = sSQL & "AND Item = '" & NumEmpresa & "' "
'''   End If
'''    sSQL = sSQL & "AND Periodo = '" & Periodo_Contable & "' " _
'''         & "AND T = 'A' " _
'''         & "AND TC IN ('FA','NV') " _
'''         & "AND LEN(Autorizacion)<13 " _
'''         & "ORDER BY TC,Autorizacion,Serie,Factura "
'''    Select_Adodc AdoAux, sSQL
'''    With AdoAux.Recordset
'''     If .RecordCount > 0 Then
'''         Do While Not .EOF
'''            Factura_Desde = .Fields("Factura")
'''            Factura_Hasta = .Fields("Factura")
'''            If .Fields("TC") = "FA" Then TipoDoc = "1"
'''            If .Fields("TC") = "NV" Then TipoDoc = "2"
'''            SerieFactura = .Fields("Serie")
'''            Serie1 = MidStrg(SerieFactura, 1, 3)
'''            Serie2 = MidStrg(SerieFactura, 4, 3)
'''            If Val(Serie1) <= 0 Then Serie1 = "001"
'''            If Val(Serie2) <= 0 Then Serie2 = "001"
'''            SetAdoAddNew "Trans_Anulados"
'''            SetAdoFields "T", Normal
'''            SetAdoFields "TipoComprobante", TipoDoc
'''            SetAdoFields "Establecimiento", Serie1
'''            SetAdoFields "PuntoEmision", Serie2
'''            SetAdoFields "Secuencial1", Factura_Desde
'''            SetAdoFields "Secuencial2", Factura_Hasta
'''            SetAdoFields "Autorizacion", .Fields("Autorizacion")
'''            SetAdoFields "FechaAnulacion", FechaFinal
'''            SetAdoFields "TP", "CD"
'''            SetAdoFields "Fecha", FechaFinal
'''            SetAdoFields "Numero", Numero
'''            SetAdoUpdate
'''            Progreso_Esperar
'''           .MoveNext
'''         Loop
'''     End If
'''    End With
'''
'''''' sSQL = "UPDATE Trans_Ventas " _
''''''      & "SET MontoIvaBienes = 0 " _
''''''      & "WHERE MontoIvaBienes IS NULL "
'''''' If ConSucursal Then
''''''    If Len(No_ATS) > 3 Then sSQL = sSQL & "AND Item NOT IN (" & No_ATS & ") "
'''''' Else
''''''    sSQL = sSQL & "AND Item = '" & NumEmpresa & "' "
'''''' End If
'''''' sSQL = sSQL _
''''''      & "AND Periodo = '" & Periodo_Contable & "' "
'''''' Ejecutar_SQL_SP sSQL
''''''
'''''' sSQL = "UPDATE Trans_Ventas " _
''''''      & "SET MontoIvaServicios = 0 " _
''''''      & "WHERE MontoIvaServicios IS NULL "
'''''' If ConSucursal Then
''''''    If Len(No_ATS) > 3 Then sSQL = sSQL & "AND Item NOT IN (" & No_ATS & ") "
'''''' Else
''''''    sSQL = sSQL & "AND Item = '" & NumEmpresa & "' "
'''''' End If
'''''' sSQL = sSQL _
''''''      & "AND Periodo = '" & Periodo_Contable & "' "
'''''' Ejecutar_SQL_SP sSQL
''' 'MsgBox "Insertar Ventas ATS"
'''  RatonNormal
'''End Sub

Public Sub Insertar_Ventas(CodigoProv As String, _
                           CantidadFacturas As Long, _
                           TotalCero As Currency, _
                           TotalGravada As Currency, _
                           TotalIVA As Currency, _
                           SubTotalFA As Currency, _
                           TipoDocumento As String)
Dim BaseImpIVAB As Currency
Dim BaseImpIVAS As Currency
Dim TotalImpIVA As Currency
Dim PorcIVA_B As Single
Dim PorcIVA_S As Single
  'If (TotalCero + TotalGravada + SubTotalFA) > 0 Then
    'MsgBox Total_Con_IVA & vbCrLf & Total_Sin_IVA & vbCrLf & TotalIVA
     BaseImpIVAB = 0
     BaseImpIVAS = 0
     PorcIVA_B = 0
     PorcIVA_S = 0
     If Val(TipoDocumento) <> 4 Then
        If PorcIVAB > 0 Then BaseImpIVAB = Redondear((Total_RetIVAB * 100) / PorcIVAB, 2)
        If PorcIVAS > 0 Then BaseImpIVAS = Redondear((Total_RetIVAS * 100) / PorcIVAS, 2)
        Select Case PorcIVAB
          Case 1 To 30: PorcIVA_B = 1
          Case 31 To 100: PorcIVA_B = 3
        End Select
        Select Case PorcIVAS
          Case 1 To 70: PorcIVA_S = 2
          Case 71 To 100: PorcIVA_S = 3
        End Select
     End If
     If TotalGravada <> 0 Then
        TotalImpIVA = Redondear(TotalGravada * 0.12, 2)
       'If TotalGravada = 1113.37 Then MsgBox TotalImpIVA
     Else
        TotalImpIVA = 0
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
     SetAdoFields "MontoIva", TotalImpIVA
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
     SetAdoFields "T", Normal
     SetAdoFields "TP", "CD"
     SetAdoFields "Numero", Numero
     SetAdoFields "Fecha", FechaFinal
     SetAdoUpdate
  'End If
End Sub

Public Sub Insertar_Ventas_NC(RUC_CI As String, _
                              TB As String, _
                              Beneficiario As String, _
                              CantidadFacturas As Long, _
                              TotalCero As Currency, _
                              TotalGravada As Currency, _
                              TotalIVA As Currency, _
                              SubTotalFA As Currency, _
                              TipoDocumento As String)
Dim BaseImpIVAB As Currency
Dim BaseImpIVAS As Currency
Dim TotalImpIVA As Currency
Dim PorcIVA_B As Single
Dim PorcIVA_S As Single
'CodigoCliente, Codigo3, NombreCliente, CLng(Cantidad),  TBaseCero, TBaseGravada, TTotal_IVA, TBaseSubTotal,  "4"
  'If (TotalCero + TotalGravada + SubTotalFA) > 0 Then
    'MsgBox Total_Con_IVA & vbCrLf & Total_Sin_IVA & vbCrLf & TotalIVA
     BaseImpIVAB = 0
     BaseImpIVAS = 0
     PorcIVA_B = 0
     PorcIVA_S = 0
     SetAdoAddNew "Trans_Ventas"
     SetAdoFields "RUC_CI", RUC_CI
     SetAdoFields "TB", TB
     SetAdoFields "Razon_Social", Beneficiario
     SetAdoFields "TipoComprobante", TipoDocumento
     SetAdoFields "FechaRegistro", FechaFinal
     SetAdoFields "FechaEmision", FechaFinal
     SetAdoFields "Establecimiento", Codigo1
     SetAdoFields "PuntoEmision", Codigo2
     SetAdoFields "NumeroComprobantes", CantidadFacturas
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
     SetAdoFields "Linea_SRI", Ln_SRI
     SetAdoFields "T", Normal
     SetAdoFields "TP", "CD"
     SetAdoFields "Numero", Numero
     SetAdoFields "Fecha", FechaFinal
     SetAdoUpdate
     Ln_SRI = Ln_SRI + 1
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
     SetAdoFields "ID", NumTrans
     SetAdoFields "Linea_SRI", 0
     SetAdoUpdate
  End If
End Sub

Public Sub Insertar_Ret_Fuente_Ventas(CodigoCli As String)
Dim RegAdodc As ADODB.Recordset
Dim DatosSelect As String
  RatonReloj
  'If CodigoCli = "1790892875" Then MsgBox CodigoCliente & vbCrLf & FechaInicial
  NumTrans = Maximo_De("Trans_Air", "ID") + 1
  DatosSelect = "SELECT CodigoC,Banco,SUM(Base_Imponible) As BaseImp,SUM(Abono) As Abonos " _
              & "FROM Trans_Abonos " _
              & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
              & "AND MidStrg(Banco,1,16) = 'RETENCION FUENTE' "
   If ConSucursal Then
      If Len(No_ATS) > 3 Then DatosSelect = DatosSelect & "AND Item NOT IN (" & No_ATS & ") "
   Else
      DatosSelect = DatosSelect & "AND Item = '" & NumEmpresa & "' "
   End If
   DatosSelect = DatosSelect & "AND Periodo = '" & Periodo_Contable & "' " _
               & "AND CodigoC = '" & CodigoCli & "' " _
               & "AND X = '.' " _
               & "GROUP BY CodigoC,Banco " _
               & "ORDER BY CodigoC,Banco "
   Select_AdoDB RegAdodc, DatosSelect
 'Insertamos Retenciones en la Fuente en ventas
  If RegAdodc.RecordCount > 0 Then
     Do While Not RegAdodc.EOF
        TipoDoc = SinEspaciosDer(RegAdodc.fields("Banco"))
        Total = RegAdodc.fields("BaseImp")
        CR_ATS = Leer_Concepto_Retencion(FechaInicial, TipoDoc)
        Porc = CR_ATS.Porcentaje
        'MsgBox TipoDoc
        Insertar_Ventas_Air CodigoCli, Total, Porc, RegAdodc.fields("Abonos"), Factura_No, 0, "001", "001", String(10, "9"), Ninguno
        RegAdodc.MoveNext
     Loop
     DatosSelect = "UPDATE Trans_Abonos " _
                 & "SET X = 'V' " _
                 & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
                 & "AND MidStrg(Banco,1,16) = 'RETENCION FUENTE' "
     If ConSucursal Then
        If Len(No_ATS) > 3 Then DatosSelect = DatosSelect & "AND Item NOT IN (" & No_ATS & ") "
     Else
        DatosSelect = DatosSelect & "AND Item = '" & NumEmpresa & "' "
     End If
     DatosSelect = DatosSelect _
                 & "AND Periodo = '" & Periodo_Contable & "' " _
                 & "AND CodigoC = '" & CodigoCli & "' "
     Ejecutar_SQL_SP DatosSelect
  End If
  RegAdodc.Close
End Sub

