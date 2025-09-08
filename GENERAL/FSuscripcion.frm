VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form FSuscripcion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "FORMULARIO DE SUSCRIPCION"
   ClientHeight    =   6765
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8625
   Icon            =   "FSuscripcion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6765
   ScaleWidth      =   8625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton OpcAnual 
      Caption         =   "Anual"
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
      Left            =   6195
      TabIndex        =   34
      Top             =   945
      Width           =   1065
   End
   Begin VB.OptionButton OpcTrimestral 
      Caption         =   "Trimestral"
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
      Left            =   6195
      TabIndex        =   33
      Top             =   1365
      Width           =   1170
   End
   Begin VB.OptionButton OpcSemestral 
      Caption         =   "Semestral"
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
      Left            =   6195
      TabIndex        =   32
      Top             =   1785
      Width           =   1170
   End
   Begin VB.TextBox TextComision 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   330
      Left            =   4830
      MaxLength       =   5
      TabIndex        =   27
      Top             =   3465
      Width           =   1170
   End
   Begin VB.TextBox TxtAtencion 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   330
      Left            =   105
      MaxLength       =   40
      TabIndex        =   21
      Top             =   2730
      Width           =   4740
   End
   Begin VB.Frame FrmTipo 
      BorderStyle     =   0  'None
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
      Left            =   4935
      TabIndex        =   30
      Top             =   2415
      Width           =   1590
      Begin VB.OptionButton OpcR 
         Caption         =   "Renovación"
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
         TabIndex        =   23
         Top             =   420
         Width           =   1485
      End
      Begin VB.OptionButton OpcN 
         Caption         =   "Nueva"
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
         TabIndex        =   22
         Top             =   105
         Value           =   -1  'True
         Width           =   1065
      End
   End
   Begin VB.TextBox TextValor 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   330
      Left            =   3360
      MaxLength       =   16
      TabIndex        =   19
      Top             =   1995
      Width           =   1485
   End
   Begin VB.OptionButton OpcSemanal 
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
      Height          =   225
      Left            =   4935
      TabIndex        =   11
      Top             =   1785
      Width           =   1170
   End
   Begin VB.TextBox TextFact 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   330
      Left            =   1890
      MaxLength       =   8
      TabIndex        =   17
      Top             =   1995
      Width           =   1485
   End
   Begin VB.TextBox TextTipo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   330
      Left            =   1155
      MaxLength       =   5
      TabIndex        =   15
      Top             =   1995
      Width           =   750
   End
   Begin MSDataGridLib.DataGrid DGSuscripcion 
      Bindings        =   "FSuscripcion.frx":059A
      Height          =   2430
      Left            =   105
      TabIndex        =   31
      Top             =   3885
      Width           =   6630
      _ExtentX        =   11695
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
   Begin VB.TextBox TxtHasta 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   330
      Left            =   105
      MaxLength       =   5
      TabIndex        =   13
      Top             =   1995
      Width           =   1065
   End
   Begin VB.OptionButton OpcQuince 
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
      Height          =   225
      Left            =   4935
      TabIndex        =   10
      Top             =   1365
      Width           =   1275
   End
   Begin VB.OptionButton OpcMes 
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
      Height          =   225
      Left            =   4935
      TabIndex        =   9
      Top             =   945
      Value           =   -1  'True
      Width           =   1065
   End
   Begin MSDataListLib.DataCombo DCEjecutivo 
      Bindings        =   "FSuscripcion.frx":05B7
      DataSource      =   "AdoAux"
      Height          =   345
      Left            =   105
      TabIndex        =   25
      Top             =   3465
      Width           =   4740
      _ExtentX        =   8361
      _ExtentY        =   609
      _Version        =   393216
      Text            =   "DataCombo1"
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
   Begin MSAdodcLib.Adodc AdoSuscripcion 
      Height          =   330
      Left            =   105
      Top             =   6300
      Width           =   6630
      _ExtentX        =   11695
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
      Caption         =   "Suscripcion"
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
      Caption         =   "&Cancelar"
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
      Left            =   7560
      Picture         =   "FSuscripcion.frx":05CC
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   1050
      Width           =   960
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
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
      Left            =   7560
      Picture         =   "FSuscripcion.frx":0E96
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   105
      Width           =   960
   End
   Begin VB.TextBox TextSector 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   330
      Left            =   3885
      MaxLength       =   8
      TabIndex        =   8
      Top             =   1260
      Width           =   960
   End
   Begin VB.TextBox TextContrato 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   330
      Left            =   2625
      MaxLength       =   8
      TabIndex        =   6
      Top             =   1260
      Width           =   1275
   End
   Begin MSMask.MaskEdBox MBHasta 
      Height          =   330
      Left            =   1365
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
   Begin MSMask.MaskEdBox MBDesde 
      Height          =   330
      Left            =   105
      TabIndex        =   3
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
   Begin MSAdodcLib.Adodc AdoAux 
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
   Begin MSAdodcLib.Adodc AdoAux1 
      Height          =   330
      Left            =   210
      Top             =   4410
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
      Caption         =   "Aux1"
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
   Begin MSAdodcLib.Adodc AdoAux2 
      Height          =   330
      Left            =   210
      Top             =   4725
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
      Caption         =   "Aux2"
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
   Begin MSDataListLib.DataCombo DCCtaVenta 
      Bindings        =   "FSuscripcion.frx":1760
      DataSource      =   "AdoCtaVenta"
      Height          =   345
      Left            =   105
      TabIndex        =   1
      Top             =   420
      Width           =   6630
      _ExtentX        =   11695
      _ExtentY        =   609
      _Version        =   393216
      Text            =   "DataCombo1"
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
   Begin MSAdodcLib.Adodc AdoCtaVenta 
      Height          =   330
      Left            =   210
      Top             =   5040
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
   Begin VB.Label Label11 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Comision %"
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
      TabIndex        =   26
      Top             =   3150
      Width           =   1170
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Atención /Entregar a:"
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
      TabIndex        =   20
      Top             =   2415
      Width           =   4740
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Valor Suscrip"
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
      TabIndex        =   18
      Top             =   1680
      Width           =   1485
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Comp. Venta"
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
      Left            =   1890
      TabIndex        =   16
      Top             =   1680
      Width           =   1485
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Tipo"
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
      Left            =   1155
      TabIndex        =   14
      Top             =   1680
      Width           =   750
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Ent. hasta"
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
      TabIndex        =   12
      Top             =   1680
      Width           =   1065
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Ejecutivo de Venta"
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
      TabIndex        =   24
      Top             =   3150
      Width           =   4740
   End
   Begin VB.Label LblCliente 
      Alignment       =   2  'Center
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
      Height          =   330
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   6630
   End
   Begin VB.Label Label33 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Sector"
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
      TabIndex        =   7
      Top             =   945
      Width           =   960
   End
   Begin VB.Label Label13 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Contrato No."
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
      Top             =   945
      Width           =   1275
   End
   Begin VB.Label Label12 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Periodo"
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
      Top             =   945
      Width           =   2535
   End
End
Attribute VB_Name = "FSuscripcion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
  DGSuscripcion.Visible = False
  TextoValido TextTipo, , True
  If OpcN.value Then TipoProc = "N" Else TipoProc = "R"
  Opcion = CInt(TxtHasta.Text)
  If OpcMes.value Then TipoDoc = "MENS"
  If OpcQuince.value Then TipoDoc = "QUNC"
  If OpcSemanal.value Then TipoDoc = "SEMA"
  If OpcAnual.value Then TipoDoc = "ANUA"
  If OpcTrimestral.value Then TipoDoc = "TRIM"
  If OpcSemestral.value Then TipoDoc = "SEME"
  Credito_No = TextTipo.Text & "-" & Format$(Val(TextContrato.Text), "0000000")
  sSQL = "DELETE * " _
       & "FROM Prestamos " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND TP = '" & TipoDoc & "' " _
       & "AND Credito_No = '" & Credito_No & "' "
  Ejecutar_SQL_SP sSQL
  sSQL = "DELETE * " _
       & "FROM Trans_Suscripciones " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND TP = '" & TipoDoc & "' " _
       & "AND Contrato_No = '" & Credito_No & "' "
  Ejecutar_SQL_SP sSQL
  With AdoSuscripcion.Recordset
   If .RecordCount > 0 Then
       RatonReloj
      'Encabezado
       SetAddNew AdoAux2
       SetFields AdoAux2, "T", TipoProc
       SetFields AdoAux2, "Sector", TextSector.Text
       SetFields AdoAux2, "TP", TipoDoc
       SetFields AdoAux2, "Credito_No", Credito_No
       SetFields AdoAux2, "No_Venc", Opcion
       SetFields AdoAux2, "Cuenta_No", CodigoCliente
       SetFields AdoAux2, "Meses", NumMeses
       SetFields AdoAux2, "Fecha", MBDesde.Text
       SetFields AdoAux2, "Fecha_C", MBHasta.Text
       SetFields AdoAux2, "Capital", Total
       SetFields AdoAux2, "Encaje", Redondear(Total * (Val(TextComision.Text) / 100), 4)
       SetFields AdoAux2, "Numero", CLng(TextFact.Text)
       SetFields AdoAux2, "Atencion", TxtAtencion.Text
       SetFields AdoAux2, "Cta", Cta_Aux
       SetFields AdoAux2, "Pagos", Redondear(.fields("Capital"), 2)
       SetFields AdoAux2, "Item", NumEmpresa
       SetFields AdoAux2, "CodigoE", CodigoEjecutivo
       SetFields AdoAux2, "CodigoU", CodigoUsuario
       SetUpdate AdoAux2
       RatonNormal
   End If
  End With
 'Detalle
  With AdoSuscripcion.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       RatonReloj
       Do While Not .EOF
          FSuscripcion.Caption = .fields("Ejemplar")
          SetAddNew AdoAux1
          SetFields AdoAux1, "T", TipoProc
          SetFields AdoAux1, "AC", adFalse
          SetFields AdoAux1, "TP", TipoDoc
          SetFields AdoAux1, "Contrato_No", Credito_No
          SetFields AdoAux1, "Ent_No", .fields("Ejemplar")
          SetFields AdoAux1, "Fecha", .fields("Fecha")
          SetFields AdoAux1, "E", .fields("Entregado")
          SetFields AdoAux1, "Item", NumEmpresa
          SetUpdate AdoAux1
         .MoveNext
       Loop
       RatonNormal
   End If
  End With
  DGSuscripcion.Visible = True
  Unload FSuscripcion
End Sub

Private Sub Command2_Click()
   Unload Me
End Sub

Private Sub DCCtaVenta_LostFocus()
  Cta_Aux = Ninguno
  With AdoCtaVenta.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       Cta_Aux = .fields("Cta_Ventas")
      .Find ("Cuenta = '" & DCCtaVenta.Text & "' ")
       If Not .EOF Then Cta_Aux = .fields("Cta_Ventas") Else MsgBox "Cliente no Asignado"
   Else
       MsgBox "No existen datos"
   End If
  End With
End Sub

Private Sub DCEjecutivo_LostFocus()
  RatonReloj
  CodigoEjecutivo = Ninguno
  With AdoAux.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Cliente = '" & DCEjecutivo.Text & "' ")
       If Not .EOF Then CodigoEjecutivo = .fields("Codigo")
   End If
  End With
  RatonNormal
End Sub

Private Sub Form_Activate()
  Trans_No = 250
  sSQL = "SELECT * " _
       & "FROM Prestamos " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Fecha = #" & BuscarFecha(FechaSistema) & "# "
  Select_Adodc AdoAux2, sSQL
  
  sSQL = "SELECT * " _
       & "FROM Trans_Suscripciones " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Fecha = #" & BuscarFecha(FechaSistema) & "# "
  Select_Adodc AdoAux1, sSQL
  
  sSQL = "DELETE * " _
       & "FROM Asiento_P " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND T_No = " & Trans_No & " " _
       & "AND CodigoU = '" & CodigoUsuario & "' "
  Ejecutar_SQL_SP sSQL
  
  sSQL = "SELECT Cuotas As Ejemplar,Fecha,Pagos As Entregado,Cta As Sector,Comision,Capital,T_No,Item,CodigoU " _
       & "FROM Asiento_P " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND T_No = " & Trans_No & " " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "ORDER BY Cuotas,Fecha "
  Select_Adodc_Grid DGSuscripcion, AdoSuscripcion, sSQL
  
  sSQL = "SELECT DISTINCT CP.Cta_Ventas,CC.Cuenta " _
       & "FROM Catalogo_Productos As CP,Catalogo_Cuentas As CC " _
       & "WHERE CP.Item = '" & NumEmpresa & "' " _
       & "AND CP.Item = CC.Item " _
       & "AND CP.Cta_Ventas = CC.Codigo " _
       & "ORDER BY CP.Cta_Ventas "
  SelectDB_Combo DCCtaVenta, AdoCtaVenta, sSQL, "Cuenta"
  
  sSQL = "SELECT CR.Codigo,C.Cliente,C.CI_RUC,C.Porc_C " _
        & "FROM Clientes As C,Catalogo_Rol_Pagos As CR " _
        & "WHERE CR.Item = '" & NumEmpresa & "' " _
        & "AND CR.Periodo = '" & Periodo_Contable & "' " _
        & "AND C.Codigo = CR.Codigo " _
        & "ORDER BY C.Cliente "
  SelectDB_Combo DCEjecutivo, AdoAux, sSQL, "Cliente"
  LblCliente.Caption = NombreCliente
  TextFact.Text = CStr(Factura_No)
  RatonNormal
End Sub

Private Sub Form_Load()
  CentrarForm FSuscripcion
  ConectarAdodc AdoAux
  ConectarAdodc AdoAux1
  ConectarAdodc AdoAux2
  ConectarAdodc AdoCtaVenta
  ConectarAdodc AdoSuscripcion
End Sub

Private Sub MBDesde_GotFocus()
  MarcarTexto MBDesde
End Sub

Private Sub MBDesde_LostFocus()
  FechaValida MBDesde
End Sub

Private Sub MBHasta_GotFocus()
  MarcarTexto MBHasta
End Sub

Private Sub MBHasta_LostFocus()
  FechaValida MBHasta
End Sub

Private Sub Option2_Click()
End Sub

Private Sub TextComision_GotFocus()
  MarcarTexto TextComision
End Sub

Private Sub TextComision_LostFocus()
  DGSuscripcion.Visible = False
  FechaValida MBDesde
  FechaValida MBHasta
  TextoValido TextSector, , True
  TextoValido TxtHasta, True
  TextoValido TextFact, True
  TextoValido TextValor, True
  
  sSQL = "DELETE * " _
       & "FROM Asiento_P " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND T_No = " & Trans_No & " " _
       & "AND CodigoU = '" & CodigoUsuario & "' "
  Ejecutar_SQL_SP sSQL
  
  sSQL = "SELECT Cuotas As Ejemplar,Fecha,Pagos As Entregado,Cta As Sector,Comision,Capital,T_No,Item,CodigoU " _
       & "FROM Asiento_P " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND T_No = " & Trans_No & " " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "ORDER BY Cuotas,Fecha "
  Select_Adodc_Grid DGSuscripcion, AdoSuscripcion, sSQL
  Saldo = 0: Diferencia = 0: Cuota_No = 1
  Opcion = CInt(TxtHasta.Text) 'Numero de entrega
  Total = Val(TextValor.Text)  'Valor de la Suscripcion
  Mifecha = MBDesde.Text
  I = CFechaLong(MBDesde.Text)
  J = CFechaLong(MBHasta.Text)
  If OpcMes.value Then NumMeses = (J - I) / 31
  If OpcQuince.value Then NumMeses = (J - I) / 15
  If OpcSemanal.value Then NumMeses = (J - I) / 7
  If OpcAnual.value Then NumMeses = 1
  If OpcTrimestral.value Then NumMeses = 4
  If OpcSemestral.value Then NumMeses = 2
  If NumMeses <= 0 Then NumMeses = 1
  Saldo = Redondear(Total / NumMeses, 2)
  For Cuota_No = 1 To NumMeses
    If OpcMes.value Then Mifecha = SiguienteMes(Mifecha)
    If OpcQuince.value Then Mifecha = CLongFecha(CFechaLong(Mifecha) + 15)
    If OpcSemanal.value Then Mifecha = CLongFecha(CFechaLong(Mifecha) + 7)
    If OpcAnual.value Then Mifecha = CLongFecha(CFechaLong(Mifecha) + 366)
    If OpcTrimestral.value Then Mifecha = CLongFecha(CFechaLong(Mifecha) + 91)
    If OpcSemestral.value Then Mifecha = CLongFecha(CFechaLong(Mifecha) + 182)
    If Opcion < Cuota_No Then Si_No = adFalse Else Si_No = adTrue
    SetAddNew AdoSuscripcion
    SetFields AdoSuscripcion, "Sector", TextSector.Text
    SetFields AdoSuscripcion, "Ejemplar", Cuota_No
    SetFields AdoSuscripcion, "Fecha", Mifecha
    SetFields AdoSuscripcion, "Entregado", Si_No
    SetFields AdoSuscripcion, "Comision", 0
    SetFields AdoSuscripcion, "Capital", Saldo
    SetFields AdoSuscripcion, "T_No", Trans_No
    SetFields AdoSuscripcion, "Item", NumEmpresa
    SetFields AdoSuscripcion, "CodigoU", CodigoUsuario
    SetUpdate AdoSuscripcion
  Next Cuota_No
  AdoSuscripcion.Caption = "Periodos: " & NumMeses
  DGSuscripcion.Visible = True
End Sub

Private Sub TextContrato_GotFocus()
  MarcarTexto TextContrato
End Sub

Private Sub TextContrato_LostFocus()
  TextoValido TextContrato, , True
  TextContrato.Text = Format$(TextContrato.Text, "0000000")
  Credito_No = TextContrato.Text
End Sub

Private Sub TextFact_GotFocus()
  MarcarTexto TextFact
End Sub

Private Sub TextFact_LostFocus()
  TextoValido TextFact, True
End Sub

Private Sub TextSector_GotFocus()
  MarcarTexto TextSector
End Sub

Private Sub TextSector_LostFocus()
  TextoValido TextSector, , True
End Sub

Private Sub TextTipo_GotFocus()
  MarcarTexto TextTipo
End Sub

Private Sub TextTipo_LostFocus()
  TextoValido TextTipo, , True
End Sub

Private Sub TextValor_GotFocus()
   MarcarTexto TextValor
End Sub

Private Sub TextValor_LostFocus()
   TextoValido TextValor, True
End Sub

Private Sub TxtAtencion_GotFocus()
  TxtAtencion.Text = ""
End Sub

Private Sub TxtAtencion_LostFocus()
  TextoValido TxtAtencion, , True
End Sub

Private Sub TxtHasta_GotFocus()
  MarcarTexto TxtHasta
End Sub

Private Sub TxtHasta_LostFocus()
  TextoValido TxtHasta, True
End Sub
