VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Begin VB.Form BalanceComp 
   Caption         =   "BALANCE DE COMPROBACION"
   ClientHeight    =   7575
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11280
   Icon            =   "Balacomp.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7575
   ScaleWidth      =   11280
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   210
      Top             =   6195
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Balacomp.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Balacomp.frx":0D1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Balacomp.frx":1036
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Balacomp.frx":1350
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Balacomp.frx":166A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Balacomp.frx":1984
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Balacomp.frx":2216
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Balacomp.frx":2530
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Balacomp.frx":2EDE
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Balacomp.frx":37B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Balacomp.frx":3AD2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   11280
      _ExtentX        =   19897
      _ExtentY        =   1164
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Salir"
            Object.ToolTipText     =   "Salir del Modulo"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Procesar_Balance"
            Object.ToolTipText     =   "Procesar Balance de Comprobación"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Procesar_Balance_Mensual"
            Object.ToolTipText     =   "Procesa Balance Mensual"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Procesar_Balance_Consolidado"
            Object.ToolTipText     =   "Procesa Balance Consolidado de Varias Sucursales/Agencias"
            ImageIndex      =   4
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "BC_BC"
                  Text            =   "Balance de Comprobacion"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "BC_ES"
                  Text            =   "Estado de Situacion"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "BC_ER"
                  Text            =   "Estado de Resultado"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Presentar_Balance_Comprobacion"
            Object.ToolTipText     =   "Presenta Balance de Comprobación"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Presenta_Estado_Situacion"
            Object.ToolTipText     =   "Presenta Estado de Situación (General)"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Presenta_Estado_Resultado"
            Object.ToolTipText     =   "Presenta Estado de Resultado"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Presenta_Balance_Semanal"
            Object.ToolTipText     =   "Presenta Balance Mensual por semanas"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Imprimir"
            Object.ToolTipText     =   "Imprimir Resultados"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SBSB11"
            Object.ToolTipText     =   "SBS B11"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Excel"
            Object.ToolTipText     =   "Exportar a Excel"
            ImageIndex      =   11
         EndProperty
      EndProperty
      Begin VB.Frame Frame1 
         Caption         =   "Tipo de Presentacion de Cuentas"
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
         Left            =   13545
         TabIndex        =   18
         Top             =   0
         Width           =   3270
         Begin VB.OptionButton OpcD 
            Caption         =   "Detalle"
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
            Left            =   2205
            TabIndex        =   21
            Top             =   315
            Width           =   960
         End
         Begin VB.OptionButton OpcG 
            Caption         =   "Grupo"
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
            Left            =   1155
            TabIndex        =   20
            Top             =   315
            Width           =   855
         End
         Begin VB.OptionButton OpcDG 
            Caption         =   "Todas"
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
            TabIndex        =   19
            Top             =   315
            Value           =   -1  'True
            Width           =   855
         End
      End
      Begin VB.Frame Frame2 
         Height          =   645
         Left            =   6720
         TabIndex        =   11
         Top             =   0
         Width           =   6735
         Begin VB.TextBox TextCotiza 
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
            Left            =   5460
            MaxLength       =   11
            MultiLine       =   -1  'True
            TabIndex        =   12
            Top             =   210
            Width           =   1170
         End
         Begin MSMask.MaskEdBox MBFechaI 
            Height          =   330
            Left            =   840
            TabIndex        =   13
            Top             =   210
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
         Begin MSMask.MaskEdBox MBFechaF 
            Height          =   330
            Left            =   2940
            TabIndex        =   14
            Top             =   210
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
         Begin VB.Label Label1 
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
            Index           =   0
            Left            =   105
            TabIndex        =   17
            Top             =   210
            Width           =   750
         End
         Begin VB.Label Label2 
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
            Index           =   0
            Left            =   2205
            TabIndex        =   16
            Top             =   210
            Width           =   750
         End
         Begin VB.Label Label5 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Cotizacion:"
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
            Index           =   0
            Left            =   4305
            TabIndex        =   15
            Top             =   210
            Width           =   1170
         End
      End
   End
   Begin MSDataGridLib.DataGrid DGBalance 
      Bindings        =   "Balacomp.frx":4724
      Height          =   4635
      Left            =   105
      TabIndex        =   8
      Top             =   1470
      Width           =   14190
      _ExtentX        =   25030
      _ExtentY        =   8176
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      WrapCellPointer =   -1  'True
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
   Begin VB.CommandButton Command1 
      Caption         =   "&S"
      Height          =   330
      Left            =   210
      TabIndex        =   7
      Top             =   7665
      Width           =   225
   End
   Begin MSAdodcLib.Adodc AdoTrans 
      Height          =   330
      Left            =   315
      Top             =   2940
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
   Begin MSAdodcLib.Adodc AdoCtas1 
      Height          =   330
      Left            =   315
      Top             =   3255
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
   Begin MSAdodcLib.Adodc AdoCtas 
      Height          =   330
      Left            =   315
      Top             =   3570
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
   Begin MSAdodcLib.Adodc AdoBalance 
      Height          =   330
      Left            =   0
      Top             =   7140
      Width           =   2640
      _ExtentX        =   4657
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
      Caption         =   "Balance"
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
   Begin MSAdodcLib.Adodc AdoCC 
      Height          =   330
      Left            =   315
      Top             =   3885
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
      Caption         =   "CC"
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
   Begin MSAdodcLib.Adodc AdoTotales 
      Height          =   330
      Left            =   315
      Top             =   4200
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
      Caption         =   "Totales"
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
   Begin VB.Label LblTipoBalance 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   645
      Left            =   105
      TabIndex        =   9
      Top             =   735
      Width           =   14085
   End
   Begin VB.Label LabelTotHaber 
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
      Left            =   9555
      TabIndex        =   1
      Top             =   7140
      Width           =   1905
   End
   Begin VB.Label Label9 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Haber:"
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
      Left            =   8820
      TabIndex        =   4
      Top             =   7140
      Width           =   750
   End
   Begin VB.Label LabelTotDebe 
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
      Left            =   6930
      TabIndex        =   2
      Top             =   7140
      Width           =   1905
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Debe:"
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
      TabIndex        =   3
      Top             =   7140
      Width           =   750
   End
   Begin VB.Label Label3 
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
      Left            =   5355
      TabIndex        =   6
      Top             =   7140
      Width           =   855
   End
   Begin VB.Label LabelTotSaldo 
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
      Left            =   3675
      TabIndex        =   0
      Top             =   7140
      Width           =   1695
   End
   Begin VB.Label Label11 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Diferencia:"
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
      Left            =   2625
      TabIndex        =   5
      Top             =   7140
      Width           =   1065
   End
End
Attribute VB_Name = "BalanceComp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim OpcionBalance As Byte
Dim OpcEsMensual  As Boolean
Dim Total_MN      As Currency
Dim BalanceCC     As String
Dim sSQL_Ext      As String

Public Sub ListarTipoDeBalance(EsBalanceMes As Boolean)
  RatonReloj
  FEsperar.Show
  Imagen_Esperar "Procesar Balance de Comprobacion"
  DGBalance.Visible = False
  TextCotiza = Format(Dolar, "#,##0.00")
  sSQL = "SELECT * " _
       & "FROM Fechas_Balance " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' "
  If EsBalanceMes Then
     sSQL = sSQL & "AND Detalle = 'Balance Mes' "
  Else
     sSQL = sSQL & "AND Detalle = 'Balance' "
  End If
  Select_Adodc AdoCtas, sSQL
  If AdoCtas.Recordset.RecordCount > 0 Then
     MBFechaI = AdoCtas.Recordset.Fields("Fecha_Inicial")
     MBFechaF = AdoCtas.Recordset.Fields("Fecha_Final")
  Else
     MBFechaI = FechaSistema
     MBFechaF = FechaSistema
  End If
  Imagen_Esperar "Procesar Balance de Comprobacion"
 'MsgBox "Listar Balance: " & MBFechaI
  FechaValida MBFechaI
  FechaValida MBFechaF
  Select Case OpcionBalance
    Case 1, 2, 4: SQLMsg1 = "BALANCE DE COMPROBACION "
         sSQL = "SELECT DG,Codigo,Cuenta,Saldo_Anterior,Debitos,Creditos,Saldo_Total,TC " _
              & "FROM Catalogo_Cuentas " _
              & "WHERE (Debitos<>0 OR Creditos<>0 OR Saldo_Total<>0) "
    Case 5: SQLMsg1 = "BALANCE GENERAL "
         sSQL = "SELECT Codigo,Cuenta,Total_N6,Total_N5,Total_N4,Total_N3,Total_N2,Total_N1,DG,TC " _
              & "FROM Catalogo_Cuentas " _
              & "WHERE (Total_N6+Total_N5+Total_N4+Total_N3+Total_N2+Total_N1)<>0 " _
              & "AND TB = 'ES' "
    Case 6: SQLMsg1 = "ESTADO DE RESULTADOS "
         sSQL = "SELECT Codigo,Cuenta,Total_N6,Total_N5,Total_N4,Total_N3,Total_N2,Total_N1,DG,TC " _
              & "FROM Catalogo_Cuentas " _
              & "WHERE (Total_N6+Total_N5+Total_N4+Total_N3+Total_N2+Total_N1)<>0 " _
              & "AND TB = 'ER' "
    Case 11: SQLMsg1 = "BALANCE DE PROMEDIOS "
         If Opcion = 1 Then SQLMsg1 = "BALANCE CONSOLIDADO "
         TextoValido TextCotiza, True
         Dolar = Round(CSng(TextCotiza.Text), 2)
         If Dolar > 0 Then
            sSQL = "UPDATE Catalogo_Cuentas " _
                 & "SET Saldo_Total_ME = Saldo_Total / " & Dolar & " " _
                 & "WHERE Item = '" & NumEmpresa & "' " _
                 & "AND Periodo = '" & Periodo_Contable & "' "
            Ejecutar_SQL_SP sSQL
         Else
            sSQL = "UPDATE Catalogo_Cuentas " _
                 & "SET Saldo_Total_ME = 0 " _
                 & "WHERE Item = '" & NumEmpresa & "' " _
                 & "AND Periodo = '" & Periodo_Contable & "' "
            Ejecutar_SQL_SP sSQL
         End If
         sSQL = "SELECT DG,Codigo,Cuenta,Saldo_Total_ME,Saldo_Total,TC " _
              & "FROM Catalogo_Cuentas " _
              & "WHERE DG = 'G' " _
              & "AND Saldo_Total <> 0 "
    Case 12: SQLMsg1 = "BALANCE DE COMPROBACION SBS B11 "
         sSQL = "SELECT DG,Codigo,Cuenta,Saldo_Anterior,Debitos,Creditos,Saldo_Total,TC " _
              & "FROM Catalogo_Cuentas " _
              & "WHERE LEN(Codigo) <= 9 " _
              & "AND ISNUMERIC(MidStrg(Codigo,1,1)) <> " & Val(adFalse) & " "
  End Select
  
  If OpcG.value Then sSQL = sSQL & "AND DG = 'G' "
  If OpcD.value Then sSQL = sSQL & "AND DG = 'D' "
  sSQL = sSQL & "AND Codigo <> '" & Ninguno & "' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "ORDER BY Codigo "
  Select_Adodc_Grid DGBalance, AdoBalance, sSQL
  Imagen_Esperar "Procesar Balance de Comprobacion"
  RatonReloj
  If EsBalanceMes Then SQLMsg1 = SQLMsg1 & "MENSUAL "
  
  SumaDebe = 0: SumaHaber = 0
  sSQLTotales = "SELECT DG, SUM(Debitos) As TDebitos, SUM(Creditos) As TCreditos " _
              & "FROM Catalogo_Cuentas " _
              & "WHERE Item = '" & NumEmpresa & "' " _
              & "AND Periodo = '" & Periodo_Contable & "' " _
              & "AND (Debitos + Creditos) <> 0 " _
              & "AND DG = 'D' " _
              & "GROUP BY DG "
  Select_Adodc AdoTotales, sSQLTotales
  With AdoTotales.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          SumaDebe = SumaDebe + .Fields("TDebitos")
          SumaHaber = SumaHaber + .Fields("TCreditos")
         .MoveNext
       Loop
   End If
  End With
  LabelTotDebe.Caption = Format(SumaDebe, "#,##0.00")
  LabelTotHaber.Caption = Format(SumaHaber, "#,##0.00")
  LabelTotSaldo.Caption = Format(SumaDebe - SumaHaber, "#,##0.00")
  BalanceComp.Caption = "ESTADOS FINANCIEROS"
  LblTipoBalance.Caption = SQLMsg1 & vbCrLf & "DEL  " & FechaStrgCorta(MBFechaI) & "  AL  " & FechaStrgCorta(MBFechaF)
  DGBalance.Visible = True
  RatonNormal
  Unload FEsperar
End Sub

Public Sub ListarTipoDeBalance_Ext(EsBalanceMes As Boolean, _
                                   TipoBalance As String, _
                                   TipoPyGCC As String)
    RatonReloj
    FEsperar.Show
    Imagen_Esperar "Consultando el Balance"
    DGBalance.Visible = False
    TextCotiza = Format(Dolar, "#,##0.00")
    sSQL = "SELECT * " _
         & "FROM Fechas_Balance " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' "
    If EsBalanceMes Then
       sSQL = sSQL & "AND Detalle = 'Balance Mes' "
    Else
       sSQL = sSQL & "AND Detalle = 'Balance' "
    End If
    Select_Adodc AdoCtas, sSQL
    If AdoCtas.Recordset.RecordCount > 0 Then
       MBFechaI = AdoCtas.Recordset.Fields("Fecha_Inicial")
       MBFechaF = AdoCtas.Recordset.Fields("Fecha_Final")
    Else
       MBFechaI = FechaSistema
       MBFechaF = FechaSistema
    End If
    FechaValida MBFechaI
    FechaValida MBFechaF

    RatonReloj
    sSQL = SQL_Tipo_Balance(TipoBalance, TipoPyGCC)
   'MsgBox sSQL
    Select_Adodc_Grid DGBalance, AdoBalance, sSQL
    sSQL_Ext = sSQL

    Select Case TipoBalance
      Case "BC"
           MensajeEncabData = "BALANCE DE COMPROBACION"
      Case "BS"
           MensajeEncabData = "BALANCE GENERAL"
      Case "BR"
           MensajeEncabData = "ESTADO DE RESULTADO"
''           Codigo = SinEspaciosIzq(DCCC)
''           Codigo = MidStrg(DCCC, Len(Codigo) + 1, Len(DCCC))
''           Codigo = TrimStrg(Replace(Codigo, "-", ""))
''           MensajeEncabData = Codigo
      Case Else
           MensajeEncabData = "BALANCE NO DEFINIDO"
    End Select
    SQLMsg1 = "D E L  " & MBFechaI & "  A L  " & MBFechaI

  RatonReloj
  If EsBalanceMes Then SQLMsg1 = SQLMsg1 & "MENSUAL "
  Imagen_Esperar "Consultando el Balance"
  SumaDebe = 0: SumaHaber = 0
  Select_Adodc AdoTotales, sSQLTotales
  With AdoTotales.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          SumaDebe = SumaDebe + .Fields("TDebitos")
          SumaHaber = SumaHaber + .Fields("TCreditos")
         .MoveNext
       Loop
   End If
  End With
  LabelTotDebe.Caption = Format(SumaDebe, "#,##0.00")
  LabelTotHaber.Caption = Format(SumaHaber, "#,##0.00")
  LabelTotSaldo.Caption = Format(SumaDebe - SumaHaber, "#,##0.00")
  BalanceComp.Caption = "ESTADOS FINANCIEROS"
  LblTipoBalance.Caption = MensajeEncabData & vbCrLf & SQLMsg1
  DGBalance.Caption = ""
  DGBalance.Visible = True
  RatonNormal
  Unload FEsperar
End Sub

'''Private Sub Imprimir_parcial()
'''  RatonReloj
'''  DGBalance.Visible = False
'''  FechaIni = MBFechaI.Text
'''  FechaFin = MBFechaF.Text
'''  SQLMsg2 = "AL " & FechaStrg(FechaFin)
'''  Select Case SSTbBalances.Tab
'''    Case 2:
'''         If OpcCoop Then
'''            Imprimir_General_Con AdoBalance, 2, True
'''         Else
'''            Imprimir_General AdoBalance, 2
'''         End If
'''    Case 3
'''         If OpcCoop Then
'''            Imprimir_General_Con AdoBalance, 2, False
'''         Else
'''            Imprimir_General AdoBalance, 2
'''         End If
'''  End Select
'''  DGBalance.Visible = True
'''  RatonNormal
'''End Sub

Private Sub Saldo_Promedio()
Dim NumDia As Integer
Dim Saldos_Prom_MN(31) As Currency
Dim Saldos_Prom_ME(31) As Currency
  RatonReloj
  DGBalance.Visible = False
  FechaValida MBFechaI
  FechaValida MBFechaF
  NumDia = Day(MBFechaF)
  FechaIni = BuscarFecha(MBFechaI)
  FechaFin = BuscarFecha(MBFechaF)
  sSQL = "DELETE * " _
       & "FROM Saldo_Promedios " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' "
  Ejecutar_SQL_SP sSQL
  sSQL = "SELECT * " _
       & "FROM Saldo_Promedios " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' "
  Select_Adodc AdoCtas1, sSQL
  sSQL = "SELECT Cta, Fecha, Saldo, Saldo_ME " _
       & "FROM Transacciones " _
       & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
       & "AND T <> 'A' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "ORDER BY Cta, Fecha "
  Select_Adodc AdoCtas, sSQL
  With AdoCtas.Recordset
   If .RecordCount > 0 Then
       For I = 1 To 31
           Saldos_Prom_MN(I) = 0
           Saldos_Prom_ME(I) = 0
       Next
       Codigo = .Fields("Cta")
       Do While Not .EOF
          If Codigo <> .Fields("Cta") Then
             Cadena = "Codigo: " & Codigo & vbCrLf
             Saldo = Saldos_Prom_MN(1)
             Saldo_ME = Saldos_Prom_ME(1)
             For I = 1 To NumDia
                 If Saldos_Prom_MN(I) <> 0 Then
                    Saldo = Saldos_Prom_MN(I)
                 Else
                    Saldos_Prom_MN(I) = Saldo
                 End If
                 If Saldos_Prom_ME(I) <> 0 Then
                    Saldo_ME = Saldos_Prom_ME(I)
                 Else
                    Saldos_Prom_ME(I) = Saldo_ME
                 End If
             Next
             Saldo = 0
             Saldo_ME = 0
             For I = 1 To 31
                 Saldo = Saldo + Saldos_Prom_MN(I)
                 Saldo_ME = Saldo_ME + Saldos_Prom_ME(I)
             Next
             SetAddNew AdoCtas1
             SetFields AdoCtas1, "Codigo", Codigo
             SetFields AdoCtas1, "Saldo_MN", Round(Saldo / NumDia, 2)
             SetFields AdoCtas1, "Saldo_ME", Round(Saldo_ME / NumDia, 2)
             SetFields AdoCtas1, "Item", NumEmpresa
             SetFields AdoCtas1, "CodigoU", CodigoUsuario
             SetUpdate AdoCtas1
             For I = 1 To 31
                 Saldos_Prom_MN(I) = 0
                 Saldos_Prom_ME(I) = 0
             Next
             Codigo = .Fields("Cta")
          End If
          Dia = Day(.Fields("Fecha"))
          Saldos_Prom_MN(Dia) = .Fields("Saldo")
          Saldos_Prom_ME(Dia) = .Fields("Saldo_ME")
         .MoveNext
       Loop
   End If
  End With
  sSQL = "SELECT SP.Codigo,C.Cuenta,SP.Saldo_MN,SP.Saldo_ME " _
       & "FROM Saldo_Promedios As SP,Catalogo_Cuentas As C " _
       & "WHERE C.Codigo = SP.Codigo " _
       & "AND SP.Item = '" & NumEmpresa & "' " _
       & "AND C.Periodo = '" & Periodo_Contable & "' " _
       & "AND SP.CodigoU = '" & CodigoUsuario & "' " _
       & "ORDER BY SP.Codigo "
  Select_Adodc_Grid DGBalance, AdoBalance, sSQL
  Opcion = 2
  DGBalance.Visible = True
  RatonNormal
End Sub

Private Sub Command1_Click()
  Unload BalanceComp
End Sub

Private Sub Form_Activate()
  MBFechaI = FechaSistema
  MBFechaF = FechaSistema
  Eliminar_Nulos_SP "Catalogo_Cuentas"
  BalanceCC = Ninguno
'''  sSQL = "SELECT (Codigo & ' - ' & Detalle) CentroDeCosto " _
'''       & "FROM Catalogo_SubCtas " _
'''       & "WHERE Item = '" & NumEmpresa & "' " _
'''       & "AND Periodo = '" & Periodo_Contable & "' " _
'''       & "AND TC = 'CCX' " _
'''       & "ORDER BY Agrupacion DESC,Nivel,Codigo, Detalle "
'''  SelectDB_Combo DCCC, AdoCC, sSQL, "CentroDeCosto"
  
  sSQL = "UPDATE Catalogo_Cuentas " _
       & "SET CC = SUBSTRING(Codigo,1,1) " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND CC = '.' "
  Ejecutar_SQL_SP sSQL
  
  DGBalance.Height = MDI_Y_Max - DGBalance.Top - 300
  DGBalance.width = MDI_X_Max - DGBalance.Left
  LblTipoBalance.width = MDI_X_Max - LblTipoBalance.Left
  AdoBalance.Top = DGBalance.Top + DGBalance.Height + 30
   
  Label3.Top = DGBalance.Top + DGBalance.Height + 30
  Label6.Top = DGBalance.Top + DGBalance.Height + 30
  Label9.Top = DGBalance.Top + DGBalance.Height + 30
  Label11.Top = DGBalance.Top + DGBalance.Height + 30
  LabelTotSaldo.Top = DGBalance.Top + DGBalance.Height + 30
  LabelTotDebe.Top = DGBalance.Top + DGBalance.Height + 30
  LabelTotHaber.Top = DGBalance.Top + DGBalance.Height + 30
  OpcionBalance = 4
  OpcEsMensual = False
  
  Toolbar1.buttons("Procesar_Balance_Consolidado").Enabled = ConSucursal
  If Bloquear_Control Then
     Toolbar1.buttons("Procesar_Balance").Enabled = False
     Toolbar1.buttons("Procesar_Balance_Mensual").Enabled = False
  End If
  If CNivel(7) Then
     RatonNormal
     MsgBox "Usted no esta autorizado para ingrersar a este modulo"
     Unload BalanceComp
  Else
     ListarTipoDeBalance OpcEsMensual
     RatonNormal
  End If
End Sub

Private Sub Form_Load()
  ConectarAdodc AdoCC
  ConectarAdodc AdoCtas
  ConectarAdodc AdoCtas1
  ConectarAdodc AdoTrans
 'ConectarAdodc AdoBalGen
 'ConectarAdodc AdoEstRes
 'ConectarAdodc AdoFechaBal
  ConectarAdodc AdoTotales
 'ConectarAdodc AdoBalGenCon
  ConectarAdodc AdoBalance
End Sub

Private Sub MBFechaF_LostFocus()
  FechaValida MBFechaF
End Sub

Private Sub MBFechaI_GotFocus()
  MarcarTexto MBFechaI
End Sub

Private Sub MBFechaF_GotFocus()
  MarcarTexto MBFechaF
End Sub

Private Sub MBFechaI_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBFechaF_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBFechaI_LostFocus()
  FechaValida MBFechaI
End Sub

Private Sub TextCotiza_GotFocus()
  MarcarTexto TextCotiza
End Sub

Private Sub TextCotiza_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF12 Then
     DGBalance.Visible = False
     sSQL = "SELECT * " _
          & "FROM Transacciones " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "ORDER BY Fecha "
     Select_Adodc AdoBalance, sSQL
     RatonReloj
     With AdoBalance.Recordset
      If .RecordCount > 0 Then
          Do While Not .EOF
             BalanceComp.Caption = .Fields("Fecha")
            .Fields("Debe") = Redondear(.Fields("Debe"), 2)
            .Fields("Haber") = Redondear(.Fields("Haber"), 2)
            .Fields("Parcial_ME") = Redondear(.Fields("Parcial_ME"), 2)
            .Update
            .MoveNext
          Loop
      End If
     End With
     DGBalance.Visible = True
     RatonNormal
  End If
End Sub

Public Sub Generar_SBS_B11()
'cas408902
'milcambios4089
Dim NumFile As Long
Dim NombreArchivo As String
Dim TextoLinea As String
  RatonReloj
  DGBalance.Visible = False
  With AdoBalance.Recordset
   If .RecordCount > 0 Then
       CodigoDelBanco = "4089"
       Total = 0
       TotalActivo = 0
       TotalPasivo = 0
       TotalCapital = 0
       TotalIngreso = 0
       TotalEgreso = 0
       Do While Not .EOF
          Select Case MidStrg(.Fields("Codigo"), 1, 1)
            Case "1": TotalActivo = TotalActivo + .Fields("Saldo_Total")
            Case "2": TotalPasivo = TotalPasivo + .Fields("Saldo_Total")
            Case "3": TotalCapital = TotalCapital + .Fields("Saldo_Total")
            Case "4": TotalIngreso = TotalIngreso + .Fields("Saldo_Total")
            Case "5", "6": TotalEgreso = TotalEgreso + .Fields("Saldo_Total")
            Case Else: TotalAbonos = TotalAbonos + .Fields("Saldo_Total")
          End Select
         .MoveNext
       Loop
       Total = TotalActivo + TotalPasivo + TotalCapital + TotalIngreso + TotalEgreso + TotalAbonos
      .MoveFirst
       NombreArchivo = RutaSysBases & "\SBS\B11M" & CodigoDelBanco & Replace(MBFechaF, "/", "") & ".txt"
       NumFile = FreeFile
       Open NombreArchivo For Output As #NumFile
       TextoLinea = "B11" & vbTab & CodigoDelBanco & vbTab & MBFechaF & vbTab & .RecordCount + 1 & vbTab & Format(Total, "#0.00")
       Print #NumFile, TextoLinea
       Do While Not .EOF
          TextoLinea = Replace(.Fields("Codigo"), ".", "") & vbTab & Format(.Fields("Saldo_Total"), "#0.00")
          Print #NumFile, TextoLinea
         .MoveNext
       Loop
       Close #NumFile
   End If
  End With
  RatonNormal
  DGBalance.Visible = True
  MsgBox "El Archivo: " & NombreArchivo & "," & vbCrLf & "Fue Generado exitosamente"
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    BalanceCC = ""
''    If CheqBalExt.Value <> 0 Then
''       BalanceCC = SinEspaciosIzq(DCCC)
''    Else
''       BalanceCC = "00"
''    End If
''    If BalanceCC = "" Then BalanceCC = "00"
    FechaValida MBFechaI
    FechaValida MBFechaF
   'MsgBox Button.key & " - " & BalanceCC
    Select Case Button.key
      Case "Procesar_Balance"
            OpcionBalance = 1
            OpcEsMensual = False
           'Procesar_Balance OpcEsMensual
           'MsgBox BalanceCC
            Procesar_Balance_SP OpcEsMensual, MBFechaI, MBFechaF, BalanceCC
''            If CheqBalExt.Value <> 0 Then
''              'ListarTipoDeBalance_Exterior OpcEsMensual, "BC"
''               Procesar_Balance_Ext_SP
''               ListarTipoDeBalance_Ext OpcEsMensual, "BC", BalanceCC
''               Tipo_Balance_PDF sSQL_Ext, TipoTimes, MBFechaI, MBFechaF, "BC", DCCC, BalanceCC
''            Else
               ListarTipoDeBalance OpcEsMensual
''            End If
      Case "Procesar_Balance_Mensual"
            OpcionBalance = 2
            OpcEsMensual = True
            MBFechaI = "01/" & Format(Month(MBFechaI), "00") & "/" & Format(Year(MBFechaI), "0000")
            MBFechaF = UltimoDiaMes(MBFechaI)
           'Procesar_Balance OpcEsMensual
            Procesar_Balance_SP OpcEsMensual, MBFechaI, MBFechaF, BalanceCC
            ListarTipoDeBalance OpcEsMensual
      Case "Presentar_Balance_Comprobacion"
            OpcionBalance = 4
''            If CheqBalExt.Value <> 0 Then
''               'ListarTipoDeBalance_Exterior OpcEsMensual, "BC"
''               ListarTipoDeBalance_Ext OpcEsMensual, "BC", BalanceCC
''               Tipo_Balance_PDF sSQL_Ext, TipoTimes, MBFechaI, MBFechaF, "BC", DCCC, BalanceCC
''            Else
               ListarTipoDeBalance OpcEsMensual
''            End If
      Case "Presenta_Estado_Situacion"
           'TipoCourierNew
            OpcionBalance = 5
''            If CheqBalExt.Value <> 0 Then
''               'ListarTipoDeBalance_Exterior OpcEsMensual, "BS"
''               ListarTipoDeBalance_Ext OpcEsMensual, "BS", BalanceCC
''               Tipo_Balance_PDF sSQL_Ext, TipoTimes, MBFechaI, MBFechaF, "BS", DCCC, BalanceCC
''            Else
               ListarTipoDeBalance OpcEsMensual
''            End If
      Case "Presenta_Estado_Resultado"
            OpcionBalance = 6
''            If CheqBalExt.Value <> 0 Then
''               'ListarTipoDeBalance_Exterior OpcEsMensual, "BR"
''               ListarTipoDeBalance_Ext OpcEsMensual, "BR", BalanceCC
''               Tipo_Balance_PDF sSQL_Ext, TipoTimes, MBFechaI, MBFechaF, "BR", DCCC, BalanceCC
''            Else
               ListarTipoDeBalance OpcEsMensual
''            End If
           ' MsgBox "ER..."
      Case "Presenta_Balance_Mensual"
            OpcionBalance = 7
      Case "Presenta_Balance_Semanal"
            OpcionBalance = 8
      Case "Imprimir"
            RatonReloj
            DGBalance.Visible = False
            FechaIni = MBFechaI
            FechaFin = MBFechaF
            If OpcEsMensual Then
               SQLMsg2 = "DEL " & FechaStrg(FechaIni) & " AL " & FechaStrg(FechaFin)
            Else
               SQLMsg2 = "AL " & FechaStrg(FechaFin)
            End If
            Select Case OpcionBalance
              Case 1, 2, 4: Imprimir_Balance AdoBalance
              Case 1, 2, 3: Imprimir_General_Con AdoBalance, Opcion, True
              Case 1, 2, 5: Imprimir_General AdoBalance, 1
              Case 1, 2, 6: Imprimir_General AdoBalance, 1
            End Select
            RatonNormal
            DGBalance.Visible = True
      Case "SBSB11"
           OpcionBalance = 12
           OpcEsMensual = False
           ListarTipoDeBalance OpcEsMensual
           Generar_SBS_B11
      Case "Salir"
           Unload BalanceComp
      Case "Excel"
           DGBalance.Visible = False
           GenerarDataTexto BalanceComp, AdoBalance
           DGBalance.Visible = True
    End Select
''    DGBalance.Visible = True
''    DGBalance.Visible = False
    RatonNormal
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Dim ListaSucursales As String
    OpcionBalance = 3
    sSQL = "SELECT TC, DG, Codigo, Cuenta, "
    FechaValida MBFechaF  ' 31/01/2025
    FechaFin = BuscarFecha(MBFechaF) ' 20250131
    LblTipoBalance.Caption = SQLMsg1 & vbCrLf & "AL  " & FechaStrgCorta(MBFechaF)
    Select Case ButtonMenu.key
      Case "BC_BC": SQLMsg1 = "BALANCE DE COMPROBACION CONSOLIDADO"
                    Procesar_Balance_Consolidado_SP FechaFin, "BC", ListaSucursales
      Case "BC_ES": SQLMsg1 = "BALANCE DE SITUACION CONSOLIDADO"
                    Procesar_Balance_Consolidado_SP FechaFin, "ES", ListaSucursales
      Case "BC_ER": SQLMsg1 = "BALANCE DE RESULTADO CONSOLIDADO"
                    Procesar_Balance_Consolidado_SP FechaFin, "ER", ListaSucursales
    End Select
    sSQL = sSQL & ListaSucursales & " " _
         & "FROM Balance_Consolidado " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "ORDER BY Codigo "
    Select_Adodc_Grid DGBalance, AdoBalance, sSQL
    MsgBox "Proceso Terminado con exito"
End Sub
