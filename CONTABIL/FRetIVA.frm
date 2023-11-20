VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FRetIVA 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Transaccion"
   ClientHeight    =   6660
   ClientLeft      =   45
   ClientTop       =   450
   ClientWidth     =   9360
   Icon            =   "FRetIVA.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   9360
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox TextAdu 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
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
      Left            =   6300
      MaxLength       =   16
      TabIndex        =   16
      Text            =   "0000000000000000"
      Top             =   1050
      Visible         =   0   'False
      Width           =   2010
   End
   Begin VB.TextBox TextAuto 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
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
      Left            =   5040
      MaxLength       =   10
      TabIndex        =   14
      Text            =   "0000000000"
      Top             =   1050
      Width           =   1275
   End
   Begin VB.TextBox TextSec 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
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
      Left            =   3990
      MaxLength       =   10
      TabIndex        =   12
      Text            =   "0000000"
      Top             =   1050
      Width           =   1065
   End
   Begin VB.TextBox TextSerie 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
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
      Left            =   3045
      MaxLength       =   10
      TabIndex        =   10
      Text            =   "000000"
      Top             =   1050
      Width           =   960
   End
   Begin MSMask.MaskEdBox MBoxFechaR 
      Height          =   330
      Left            =   1575
      TabIndex        =   8
      Top             =   1050
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
   Begin MSDataGridLib.DataGrid DGRet 
      Bindings        =   "FRetIVA.frx":030A
      Height          =   1170
      Left            =   105
      TabIndex        =   54
      Top             =   4305
      Width           =   9045
      _ExtentX        =   15954
      _ExtentY        =   2064
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   -1  'True
      BackColor       =   16777215
      BorderStyle     =   0
      HeadLines       =   1
      RowHeight       =   15
      RowDividerStyle =   6
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
            Format          =   "0%"
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
   Begin VB.ComboBox CPorRetTransf 
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
      Height          =   315
      Left            =   3255
      TabIndex        =   46
      Top             =   3885
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0C0&
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
      Height          =   855
      Left            =   6510
      Picture         =   "FRetIVA.frx":0322
      Style           =   1  'Graphical
      TabIndex        =   55
      Top             =   5670
      Width           =   1275
   End
   Begin VB.CommandButton CmdSalir 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Salir"
      DisabledPicture =   "FRetIVA.frx":0764
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
      Left            =   7875
      MouseIcon       =   "FRetIVA.frx":11AE
      Picture         =   "FRetIVA.frx":1BF8
      Style           =   1  'Graphical
      TabIndex        =   56
      Top             =   5670
      Width           =   1275
   End
   Begin VB.ComboBox CPorcIVA 
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
      Height          =   315
      Left            =   1785
      TabIndex        =   31
      Top             =   3150
      Width           =   1490
   End
   Begin VB.Frame Frame1 
      Caption         =   "Convenio Internacional"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   648
      Left            =   6405
      TabIndex        =   51
      Top             =   3570
      Visible         =   0   'False
      Width           =   2850
      Begin VB.OptionButton Option2 
         Caption         =   "No"
         Height          =   255
         Left            =   1575
         TabIndex        =   53
         Top             =   240
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Si"
         Height          =   255
         Left            =   315
         TabIndex        =   52
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.TextBox TxtValorIVA 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   330
      Left            =   6405
      MaxLength       =   10
      TabIndex        =   37
      Text            =   "0.00"
      Top             =   3150
      Width           =   1485
   End
   Begin VB.TextBox TValRetServ 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
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
      MaxLength       =   10
      TabIndex        =   41
      Text            =   "0.00"
      Top             =   3885
      Width           =   1695
   End
   Begin VB.TextBox TValRetTransf 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
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
      Left            =   4935
      MaxLength       =   10
      TabIndex        =   48
      Text            =   "0.00"
      Top             =   3885
      Width           =   1485
   End
   Begin VB.TextBox TBaseTransf 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   330
      Left            =   1785
      MaxLength       =   10
      TabIndex        =   43
      Text            =   "0.00"
      Top             =   3885
      Width           =   1485
   End
   Begin VB.Frame FrmSN 
      Caption         =   "Devolución"
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
      Left            =   6510
      TabIndex        =   25
      Top             =   2205
      Width           =   2745
      Begin VB.OptionButton OpcS 
         Caption         =   "Si"
         Height          =   255
         Left            =   1200
         TabIndex        =   26
         Top             =   120
         Width           =   645
      End
      Begin VB.OptionButton OpcN 
         Caption         =   "No"
         Height          =   255
         Left            =   1995
         TabIndex        =   27
         Top             =   120
         Value           =   -1  'True
         Width           =   630
      End
   End
   Begin VB.TextBox TxtICE 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
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
      Left            =   4935
      MaxLength       =   10
      TabIndex        =   35
      Text            =   "0.00"
      Top             =   3150
      Width           =   1485
   End
   Begin VB.TextBox Txtbasecero 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
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
      MaxLength       =   10
      TabIndex        =   33
      Text            =   "0.00"
      Top             =   3150
      Width           =   1695
   End
   Begin VB.TextBox TxBaseImpo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
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
      MaxLength       =   10
      TabIndex        =   29
      Text            =   "0.00"
      Top             =   3150
      Width           =   1695
   End
   Begin VB.ComboBox CPorRetServ 
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
      Height          =   315
      Left            =   7875
      TabIndex        =   39
      Top             =   3150
      Width           =   1380
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   330
      Left            =   3255
      MaxLength       =   10
      TabIndex        =   45
      Text            =   "0.00"
      Top             =   3885
      Width           =   1695
   End
   Begin VB.ComboBox ComboCT 
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
      Left            =   1575
      TabIndex        =   24
      Top             =   2310
      Width           =   4845
   End
   Begin VB.Frame Frame3 
      Caption         =   "Cta Ret IVA"
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
      Height          =   540
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9360
      Begin VB.OptionButton Opctrans 
         Caption         =   "Compras"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   210
         TabIndex        =   1
         Top             =   210
         Value           =   -1  'True
         Width           =   1380
      End
      Begin VB.OptionButton Opctrans 
         Caption         =   "Ventas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   2205
         TabIndex        =   2
         Top             =   210
         Width           =   1065
      End
      Begin VB.OptionButton Opctrans 
         Caption         =   "Exportaciones"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   3
         Left            =   6720
         TabIndex        =   4
         Top             =   210
         Width           =   2010
      End
      Begin VB.OptionButton Opctrans 
         Caption         =   "Importaciones"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   4095
         TabIndex        =   3
         Top             =   210
         Width           =   1905
      End
   End
   Begin MSAdodcLib.Adodc AdoRetenido 
      Height          =   330
      Left            =   315
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
      Caption         =   "Retenido"
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
   Begin MSAdodcLib.Adodc AdoNumEg 
      Height          =   330
      Left            =   315
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
      Caption         =   "NumEg"
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
   Begin MSAdodcLib.Adodc AdoConcepto 
      Height          =   330
      Left            =   2520
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
      Caption         =   "Concepto"
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
      Left            =   2520
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
   Begin MSAdodcLib.Adodc AdoDetRet 
      Height          =   330
      Left            =   2520
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
      Caption         =   "DetRet"
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
   Begin MSAdodcLib.Adodc AdoQuery1 
      Height          =   330
      Left            =   315
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
      Caption         =   "Query1"
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
   Begin MSDataListLib.DataCombo DCPovCli 
      Bindings        =   "FRetIVA.frx":203A
      DataSource      =   "AdoRetenido"
      Height          =   315
      Left            =   1575
      TabIndex        =   18
      Top             =   1470
      Width           =   5475
      _ExtentX        =   9657
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
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
   Begin MSDataListLib.DataCombo DCTComp 
      Bindings        =   "FRetIVA.frx":2054
      DataSource      =   "AdoConcepto"
      Height          =   315
      Left            =   1575
      TabIndex        =   21
      Top             =   1890
      Width           =   5475
      _ExtentX        =   9657
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
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
   Begin MSDataListLib.DataCombo DCCompMod 
      Bindings        =   "FRetIVA.frx":206E
      DataSource      =   "AdoConcepto"
      Height          =   315
      Left            =   7035
      TabIndex        =   19
      Top             =   1470
      Visible         =   0   'False
      Width           =   2220
      _ExtentX        =   3916
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
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
   Begin MSDataListLib.DataCombo DCNumMod 
      Bindings        =   "FRetIVA.frx":2088
      DataSource      =   "AdoConcepto"
      Height          =   315
      Left            =   7035
      TabIndex        =   22
      Top             =   1890
      Visible         =   0   'False
      Width           =   2220
      _ExtentX        =   3916
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
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
   Begin MSMask.MaskEdBox MBoxFechaE 
      Height          =   330
      Left            =   105
      TabIndex        =   6
      Top             =   1050
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
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "# Registrado Aduana"
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
      Left            =   6300
      TabIndex        =   15
      Top             =   630
      Visible         =   0   'False
      Width           =   2010
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Autorizacion Factura"
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
      Left            =   5040
      TabIndex        =   13
      Top             =   630
      Width           =   1275
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Factura No"
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
      Left            =   3990
      TabIndex        =   11
      Top             =   630
      Width           =   1065
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "# Serie Factura"
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
      Left            =   3045
      TabIndex        =   9
      Top             =   630
      Width           =   960
   End
   Begin VB.Label LApagar 
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
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   6405
      TabIndex        =   50
      Top             =   3885
      Width           =   1485
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   6405
      TabIndex        =   49
      Top             =   3465
      Width           =   1485
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4935
      TabIndex        =   47
      Top             =   3465
      Width           =   1485
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Porcentaje IVA Retenido Servicios"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   7875
      TabIndex        =   38
      Top             =   2730
      Width           =   1380
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1785
      TabIndex        =   42
      Top             =   3465
      Width           =   1485
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3255
      TabIndex        =   44
      Top             =   3465
      Width           =   1695
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   105
      TabIndex        =   40
      Top             =   3465
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Base Imponible Gravada"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   105
      TabIndex        =   28
      Top             =   2730
      Width           =   1695
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Valor ICE"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4935
      TabIndex        =   34
      Top             =   2730
      Width           =   1485
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Valor IVA:"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   6405
      TabIndex        =   36
      Top             =   2730
      Width           =   1485
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Base Imp. Tarifa 0%"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3255
      TabIndex        =   32
      Top             =   2730
      Width           =   1695
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Porcentaje de IVA"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1785
      TabIndex        =   30
      Top             =   2730
      Width           =   1485
   End
   Begin VB.Label LProCli 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
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
      TabIndex        =   17
      Top             =   1470
      Width           =   1485
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Comprobante:"
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
      Top             =   1890
      Width           =   1485
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Créd.Tributario:"
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
      TabIndex        =   23
      Top             =   2310
      Width           =   1485
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fecha Registro Comprobante:"
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
      Left            =   1575
      TabIndex        =   7
      Top             =   630
      Width           =   1485
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fecha Emision Comprobante:"
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
      Left            =   105
      TabIndex        =   5
      Top             =   630
      Width           =   1485
   End
End
Attribute VB_Name = "FRetIVA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub VerTipoTrans(vtopc As Integer)
  Topc = vtopc
  Select Case vtopc
  Case 1
     FRetIVA.Caption = "COMPRAS LOCALES"
     Label1.Visible = True
     Label1.Caption = "Fecha Emision Comprobante"
     Label12.Visible = True
     Label12.Caption = "Fecha Registro Comprobante"""
     Label8.Visible = True
     Label8.Caption = "# Serie"
     Label21.Visible = True
     Label21.Caption = "Secuencial"
     Label18.Visible = True
     Label18.Caption = "Autorizacion"
     Label3.Visible = False
     MBoxFechaE.Visible = True
     MBoxFechaR.Visible = True
     TextSerie.Visible = True
     TextSec.Visible = True
     TextAuto.Visible = True
     TextAdu.Visible = False
     LProCli.Caption = "Proveedor:"
     Label2.Visible = True
     Label2.Caption = "Base Imponible Gravada con IVA"
     Label9.Width = 1490
     Label9.Visible = True
     Label9.Caption = "Porcentaje de IVA"
     Label10.Visible = True
     Label10.Caption = "Base Imp. Tarifa 0%"
     Label13.Visible = True
     Label13.Caption = "Valor ICE"
     Label6.Visible = True
     Label6.Caption = "IVA en Transf. Bienes"
     Label15.Visible = True
     Label15.Caption = "Porcentaje IVA Transf. Bienes"
     TxBaseImpo.Visible = True
     CPorcIVA.Visible = True
     CPorcIVA.Width = 1490
     Txtbasecero.Visible = True
     TxtICE.Visible = True
     TxtValorIVA.Visible = True
     CPorRetServ.Visible = True
     Label14.Visible = True
     Label14.Caption = "IVA Retenido Transf. Bienes"
     Label17.Visible = True
     Label17.Caption = "IVA en Prestacion de Servicios"
     Label16.Visible = True
     Label16.Caption = "% IVA Retenido Prest.Servicios"
     Label19.Visible = True
     Label19.Caption = "Valor IVA Retenido Prest.Servicios"
     Label20.Visible = True
     Label20.Caption = "Valor Total IVA Pagado"
     TValRetServ.Visible = True
     TValRetServ.Text = "0.00"
     TBaseTransf.Visible = True
     TBaseTransf.Text = "0.00"
     Text3.Visible = False
     CPorRetTransf.Visible = True
     TValRetTransf.Visible = True
     TValRetTransf.Text = "0.00"
     Frame1.Visible = False
     LApagar.Visible = True
     MBoxFechaE.Text = FechaTexto
     MBoxFechaE.SetFocus
  Case 2
     FRetIVA.Caption = "VENTAS LOCALES"
     Label1.Visible = True
     Label1.Caption = "Fecha Emision Comprobante"
     Label12.Visible = True
     Label12.Caption = "Fecha Registro Comprobante"
     Label8.Visible = False
     Label21.Visible = True
     Label21.Caption = "Secuencial"
     Label18.Visible = False
     Label3.Visible = False
     MBoxFechaE.Visible = True
     MBoxFechaR.Visible = True
     TextSerie.Visible = False
     TextSec.Visible = True
     TextAuto.Visible = False
     TextAdu.Visible = False
     LProCli.Visible = True
     LProCli.Caption = "Comp. a Cliente:"
     Label2.Visible = True
     Label2.Caption = "Base Imponible Gravada con IVA"
     Label9.Visible = True
     Label9.Width = 1490
     Label9.Caption = "Porcentaje de IVA"
     Label10.Visible = True
     Label13.Visible = True
     Label6.Visible = True
     Label6.Caption = "Valor IVA"
     Label15.Visible = False
     TxBaseImpo.Visible = True
     CPorcIVA.Visible = True
     CPorcIVA.Width = 1490
     Txtbasecero.Visible = True
     TxtICE.Visible = True
     TxtValorIVA.Visible = True
     CPorRetServ.Visible = False
     Label14.Visible = False
     Label17.Visible = False
     Label16.Visible = False
     Label19.Visible = False
     Label20.Visible = False
     TValRetServ.Visible = False
     TBaseTransf.Visible = False
     CPorRetTransf.Visible = False
     Text3.Visible = False
     TValRetTransf.Visible = False
     LApagar.Visible = False
     FrmSN.Visible = False
     Frame1.Visible = False
     MBoxFechaE.Text = FechaTexto
     MBoxFechaE.SetFocus
  Case 3
     FRetIVA.Caption = "IMPORTACIONES"
     Label1.Visible = False
     Label12.Visible = True
     Label12.Caption = "Fecha Pago DUI o DAS"
     Label8.Visible = True
     Label8.Caption = "Comprobante"
     Label21.Visible = True
     Label21.Caption = "# Registro Aduana"
     Label18.Visible = False
     Label3.Visible = True
     MBoxFechaE.Visible = False
     MBoxFechaE = "01/01/1995"
     MBoxFechaR.Visible = True
     TextSerie.Visible = True
     TextSec.Visible = True
     TextAuto.Visible = False
     TextAdu.Visible = True
     LProCli.Visible = True
     LProCli.Caption = "Banco:"
     Label2.Visible = True
     Label2.Caption = "Valor CIF importados"
     Label9.Visible = True
     Label9.Width = 3100
     Label9.Caption = "Código Tarifa ICE"
     Label10.Visible = False
     Label13.Visible = True
     Label13.Caption = "Valor ICE"
     Label6.Visible = False
     Label15.Visible = False
     TxBaseImpo.Visible = True
     CPorcIVA.Visible = True
     CPorcIVA.Width = 3100
     Txtbasecero.Visible = False
     TxtICE.Visible = True
     TxtValorIVA.Visible = False
     CPorRetServ.Visible = False
     Label14.Visible = True
     Label14.Caption = "Valor IVA importación"
     Label17.Visible = False
     Label16.Visible = False
     Label19.Visible = False
     Label20.Visible = False
     TValRetServ.Visible = True
     TBaseTransf.Visible = False
     Text3.Visible = False
     CPorRetTransf.Visible = False
     TValRetTransf.Visible = False
     LApagar.Visible = False
     Frame1.Visible = True
     'CmdSalir.SetFocus
     MBoxFechaR.Text = FechaTexto
     MBoxFechaR.SetFocus
  Case 4
     FRetIVA.Caption = "EXPORTACIONES"
     Label1.Visible = True
     Label12.Visible = True
     Label12.Caption = "Fecha de Embarque"
     Label8.Visible = True
     Label8.Caption = "# Comp. FUE"
     Label21.Visible = True
     Label21.Caption = "# Secuencial"
     Label18.Visible = True
     Label3.Visible = True
     MBoxFechaE.Visible = True
     MBoxFechaR.Visible = True
     TextSerie.Visible = True
     TextSec.Visible = True
     TextAuto.Visible = True
     TextAdu.Visible = True
     LProCli.Caption = "Documento:"
     Frame1.Visible = True
     Label2.Visible = True
     Label2.Caption = "# Secuencial Factura Exportación"
     Label9.Visible = False
     Label10.Visible = False
     Label13.Visible = False
     Label6.Visible = False
     Label15.Visible = False
     TxBaseImpo.Visible = True
     TxBaseImpo.ForeColor = &H80000001
     TxBaseImpo.Text = "0000000"
     CPorcIVA.Visible = False
     Txtbasecero.Visible = False
     TxtICE.Visible = False
     TxtValorIVA.Visible = False
     CPorRetServ.Visible = False
     Label14.Visible = True
     Label14.Caption = "Valor FOB exportaciones"
     Label17.Visible = False
     Label16.Visible = False
     Label19.Visible = False
     Label20.Visible = False
     TValRetServ.Visible = True
     TBaseTransf.Visible = False
     Text3.ForeColor = &H80000001
     Text3.Text = "000000"
     Text3.Visible = False
     CPorRetTransf.Visible = False
     TValRetTransf.Visible = False
     LApagar.Visible = False
     Frame1.Visible = True
     MBoxFechaE.Text = FechaTexto
     MBoxFechaE.SetFocus
 End Select
 Frame3.Caption = Cta_Ret_Egreso & " - " & Nombre_Cta_Ret
End Sub

Private Sub CmdSalir_Click()
  Total_DetRet = 0
  Unload FRetIVA
End Sub

Private Sub Command3_Click()
  With AdoDetRet.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       Total = 0: Total_DetRet = 0
       Do While Not .EOF
          Total = Total + .Fields("VALOR_FACT")
          Total_DetRet = Total_DetRet + .Fields("VALOR_RET")
         .MoveNext
       Loop
   End If
  End With
  Unload FRetIVA
End Sub

'''Public Sub GrabarTransaccion()
'''  Valor = MsgBox("Seguro de Grabar Transacción S/N?", vbYesNo, "Grabar") = vbYes
'''  If Valor Then
'''  SQL2 = "SELECT Fecha, Valor_Fact, Porc ,Valor_Ret, Item, Codigo, CodigoU, Retencion_No, " _
'''       & "TD, FechaE, Autorizacion, Serie, Secuencial, IdenCT, BImpotcero, MontoICE, MontoIVA1, " _
'''       & "PorRetIVA1, MontoRetIVA1, MontoIVA2, PorRetIVA2, MontoRetIVA2, CodigoTR, TT, Aduana,Dev, SerieEx, CodST, Cambio, ConvInt, CodPorc " _
'''       & "FROM Asiento_R " _
'''       & "WHERE Item = '" & NumEmpresa & "' " _
'''       & "AND CodigoU = '" & CodigoUsuario & "' " _
'''       & "AND T_No = " & Trans_No & " "
'''  SelectDataGrid DGRet, AdoDetRet, SQL2
'''  'SelectAdodc AdoDetRet, SQL2
'''  GrabarRetencion AdoRetencion, AdoDetRet
'''  td = AdoDetRet.Recordset.Fields("TD")
'''  Select Case td
'''         Case "04", "05"
'''         Secuencia = Mid(DCNumMod.Text, 1, 7)
'''         tdm = Mid(DCCompMod.Text, 1, 2)
'''         SQL1 = "SELECT Item, TT, TD, Serie, Secuencial, FechaE, Fecha,Autorizacion, CodigoU " _
'''              & "FROM Trans_Retenciones " _
'''              & "WHERE Item = '" & NumEmpresa & "'" _
'''              & "AND TT = '" & tt & "' " _
'''              & "AND Secuencial = '" & Secuencia & "' " _
'''              & "AND TD = '" & tdm & "' "
'''         SelectData AdoRetencion, SQL1, False
'''         sSQL = "SELECT * FROM Trans_NDNC " _
'''              & "WHERE Item = '" & NumEmpresa & "' "
'''         SelectAdodc AdoNumEg, sSQL
'''         With AdoRetencion.Recordset
'''              If .RecordCount <> 0 Then
'''                 .MoveFirst
'''                 Do While Not .EOF
'''                    SetAddNew AdoNumEg
'''                    SetFields AdoNumEg, "Serie", .Fields("Serie")
'''                    SetFields AdoNumEg, "TD", td
'''                    SetFields AdoNumEg, "TDMod", .Fields("TD")
'''                    SetFields AdoNumEg, "Secuencial", .Fields("Secuencial")
'''                    SetFields AdoNumEg, "Autorizacion", .Fields("Autorizacion")
'''                    SetFields AdoNumEg, "Item", .Fields("Item")
'''                    SetFields AdoNumEg, "TT", .Fields("TT")
'''                    SetFields AdoNumEg, "Fecha", .Fields("Fecha")
'''                    SetFields AdoNumEg, "FechaE", .Fields("FechaE")
'''                    SetFields AdoNumEg, "CodigoU", .Fields("CodigoU")
'''                    SetUpdate AdoNumEg
'''                    .MoveNext
'''                  Loop
'''              End If
'''         End With
'''  End Select
'''  End If
'''  SQL2 = "DELETE * FROM Asiento_R " _
'''       & "WHERE Item ='" & NumEmpresa & "' " _
'''       & "AND CodigoU ='" & CodigoUsuario & "' " _
'''       & "AND CodigoTR = 'IV' " _
'''       & "AND T_No = " & Trans_No & " "
'''       ConectarAdoExecute SQL2
'''  SQL1 = "SELECT * FROM Trans_Retenciones " _
'''       & "WHERE Item ='" & NumEmpresa & "' " _
'''       & "AND CodigoTR = 'IV' " _
'''       & "AND Secuencial = '0000000' "
'''  SelectData AdoQuery1, SQL1
'''  If AdoQuery1.Recordset.RecordCount > 0 Then
'''     SQL1 = "DELETE * FROM Trans_Retenciones " _
'''          & "WHERE Item ='" & NumEmpresa & "' " _
'''          & "AND CodigoTR = 'IV' " _
'''          & "AND Secuencial = '0000000' "
'''     ConectarAdoExecute SQL1
'''  End If
'''End Sub

Private Sub Command3_GotFocus()
  Command3.SetFocus
End Sub

Private Sub ComboCT_GotFocus()
  ComboCT.Clear
  ComboCT.AddItem "00   No Aplica"
  ComboCT.AddItem "01   Crédito tributario para la declacación del IVA"
  ComboCT.AddItem "02   Costo o Gasto para la declaración de IR"
  ComboCT.AddItem "03   Activo Fijo - Crédito tributario para declaración de IVA"
  ComboCT.AddItem "04   Activo Fijo - Costo o Gasto para la declaración de IR"
  ComboCT.AddItem "05   Liquidación Gastos Viaje, hospedaje y alimentación Gastos IR"
  ComboCT.AddItem "06   Inventario - Crédito Tributario para la declaraciónn de IVA"
  ComboCT.AddItem "07   Inventario - Costo o Gasto para declaración de IR"
  ComboCT.Text = ComboCT.List(0)
  MarcarTexto ComboCT
End Sub

Private Sub CPorcIVA_GotFocus()
Select Case Topc
Case "1", "2", "4"
  CPorcIVA.Clear
  CPorcIVA.AddItem "0%"
  CPorcIVA.AddItem "10%"
  CPorcIVA.AddItem "12%"
  CPorcIVA.AddItem "14%"
Case "3"
  CPorcIVA.AddItem "0.00%    "
  CPorcIVA.AddItem "Cigarrillos rubios 77.25%"
  CPorcIVA.AddItem "Cigarrillos negros 18.54%"
  CPorcIVA.AddItem "Cerveza 30.90%"
  CPorcIVA.AddItem "Gaseosas 10.30%"
  CPorcIVA.AddItem "Alcohol 26.78% "
  CPorcIVA.AddItem "Vehículos 5.15%"
  CPorcIVA.AddItem "Suntuarios 10.30%"
  CPorcIVA.AddItem "Telecomunicaciones 15.00%"
End Select
CPorcIVA.Text = CPorcIVA.List(0)
MarcarTexto CPorcIVA
End Sub

Private Sub CPorcIVA_LostFocus()
  If CPorcIVA.ListIndex < 0 Then
     CodPorcIva = 0
  Else
     CodPorcIva = CPorcIVA.ListIndex
  End If
PorcIva = Val(Mid(SinEspaciosDer(CStr(CPorcIVA.Text)), 1, 5))
IvaTotal = 0#
IvaTotal = BaseImpo * PorcIva / 100
End Sub

Private Sub CPorRetServ_LostFocus()
  If CPorRetServ.ListIndex < 0 Then
     CodPorRetSer = 0
  Else
     CodPorRetSer = CPorRetServ.ListIndex
  End If
  PorRetSer = Val(Mid(CPorRetServ.Text, 1, 5))
  IvaServ = 0
  IvaServ = BaseServ * PorRetSer / 100
End Sub

Private Sub CPorRetTransf_GotFocus()
  CPorRetTransf.Clear
  CPorRetTransf.AddItem "0%"
  CPorRetTransf.AddItem "30%"
  CPorRetTransf.AddItem "70%"
  CPorRetTransf.AddItem "100%"
  CPorRetTransf.Text = CPorRetTransf.List(0)
  MarcarTexto CPorRetTransf
End Sub

Private Sub CPorRetServ_GotFocus()
  CPorRetServ.Clear
  CPorRetServ.AddItem "0%"
  CPorRetServ.AddItem "30%"
  CPorRetServ.AddItem "70%"
  CPorRetServ.AddItem "100%"
  CPorRetServ.Text = CPorRetServ.List(0)
  MarcarTexto CPorRetServ
End Sub

Private Sub CPorRetTransf_LostFocus()
  If CPorRetTransf.ListIndex < 0 Then
     CodPorRetTrans = 0
  Else
     CodPorRetTrans = CPorRetTransf.ListIndex
  End If
     PorRetTran = Val(Mid(CPorRetTransf.Text, 1, 5))
End Sub

Private Sub DCCompMod_GotFocus()
  sSQL = "SELECT (T.TD  & '    ' &  D.Detalle) As Detalle " _
       & "FROM Tipo_IVA As D, Trans_Retenciones As T"
  Select Case Topc
    Case "1": sSQL = sSQL & " WHERE T.TT = 'C' "
    Case "2": sSQL = sSQL & " WHERE T.TT = 'V' "
    Case "3": sSQL = sSQL & " WHERE T.TT = 'TI' "
    Case "4": sSQL = sSQL & " WHERE T.TT = 'TE' "
  End Select
  sSQL = sSQL & "AND T.Item = '000' " _
       & "AND TD <> '04' " _
       & "OR TD <> '05' " _
       & "AND T.TD = D.Codigo " _
       & "GROUP BY T.TD, D.Detalle "
  SelectDBCombo DCCompMod, AdoConcepto, sSQL, "Detalle"
End Sub

Private Sub DCNumMod_GotFocus()
  sSQL = "SELECT Secuencial As Secuencial " _
       & "FROM Trans_Retenciones "
  Select Case Topc
         Case "1":    sSQL = sSQL & " WHERE TT = 'C' "
         Case "2":  sSQL = sSQL & " WHERE TT = 'V' "
         Case "3":  sSQL = sSQL & " WHERE TT = 'TI' "
         Case "4":  sSQL = sSQL & " WHERE TT = 'TE' "
  End Select
  sSQL = sSQL & "AND Item = '" & NumEmpresa & "' " _
              & "AND TD <> '04' " _
              & "AND TD <>'05' " _
              & "GROUP BY Secuencial "
  SelectDBCombo DCNumMod, AdoConcepto, sSQL, "Secuencial"
End Sub

Private Sub DCPovCli_GotFocus()
  Select Case Topc
  Case "1", "2"
       sSQL = "SELECT (Cliente & '     ' & TD & '' & Codigo) As CodCli " _
            & "FROM Clientes " _
            & "WHERE Codigo = '" & CodigoCliente & "' "
       If Topc = "1" Then sSQL = sSQL & "AND TD <> 'O' "
       sSQL = sSQL & "ORDER BY Cliente "
  Case "3"
       sSQL = "SELECT (Banco & '     ' & Cod) As CodCli " _
            & "FROM Tipo_Banco " _
            & "WHERE Cod <> '.' " _
            & "ORDER BY Banco"
  Case "4"
       sSQL = "SELECT (Detalle & '     ' & Codigo) As CodCli " _
            & "FROM Tipo_IVA " _
            & "WHERE Item = '000' " _
            & "ORDER BY Codigo "
       DCPovCli.Enabled = False
       ComboCT.SetFocus
  End Select
  SelectDBCombo DCPovCli, AdoRetenido, sSQL, "CodCli"
  MarcarTexto DCPovCli
  If AdoRetenido.Recordset.RecordCount <= 0 Then
     Total_DetRet = 0
     MsgBox "No esta permitido realizar este proceso"
     Unload Me
  End If
End Sub

Private Sub DCPovCli_LostFocus()
  td = Mid(SinEspaciosDer(DCPovCli), 1, 1)
End Sub

Private Sub DCTComp_GotFocus()
  DCCompMod.Visible = False
  DCNumMod.Visible = False
  sSQL = "SELECT (Codigo  & '    ' &  Detalle) As Detalle " _
       & "FROM Tipo_IVA "
  Select Case Topc
    Case "1":  sSQL = sSQL & "WHERE Cod = 'C' "
    Case "2":  sSQL = sSQL & "WHERE Cod = 'V' "
    Case "3":  sSQL = sSQL & "WHERE Cod = 'TI' "
    Case "4":  sSQL = sSQL & "WHERE Cod = 'TE' "
  End Select
  sSQL = sSQL & "AND Item = '000' " _
       & "ORDER BY Codigo "
  SelectDBCombo DCTComp, AdoConcepto, sSQL, "Detalle"
  MarcarTexto DCTComp
End Sub

Private Sub DCTComp_LostFocus()
  td = SinEspaciosIzq(DCTComp.Text)
  Autorizacion = SinEspaciosDer(TextAuto.Text)
  Cadena = Mid(SinEspaciosDer(DCPovCli.Text), 1, 1)
  Select Case td
    Case "04", "05"
         DCCompMod.Visible = True
         DCNumMod.Visible = True
         DCCompMod.SetFocus
  End Select
  Cadena1 = td & Cadena
  Select Case Cadena1
    Case "01C": MsgBox "No debe hacer una Factura a un Proveedor con Cédula debe ser un Proveedor con RUC"
    Case "01O": MsgBox "No debe hacer una Factura a un Proveedor con un Documento que no sea RUC"
    Case "01P": MsgBox "No debe hacer una Factura a un Proveedor con Pasaporte debe ser un Proveedor con RUC"
    Case "02C": MsgBox "No debe hacer una Nota de Venta a un Proveedor con Cédula debe ser un Proveedor con RUC"
    Case "02O": MsgBox "No debe hacer una Nota de Venta a un Proveedor con Cédula debe ser un Proveedor con RUC"
    Case "02P": MsgBox "No debe hacer una Nota de Venta a un Proveedor con Cédula debe ser un Proveedor con RUC"
    Case "03R": MsgBox "No debe hacer una Liquidacion a un Proveedor con RUC"
    Case "03O": MsgBox "No debe hacer una Factura a un Proveedor con Otro Tipo de Documento"
    Case "08C": MsgBox "No debe hacer una Boleta a un Proveedor con Cédula"
    Case "08P": MsgBox "No debe hacer una Boleta a un Proveedor con Pasaporte"
    Case "08O": MsgBox "No debe hacer una Boleta a un Proveedor con Otro tipo de Documento"
    Case "09C": MsgBox "No debe hacer un Tiquete o Vale a un Proveedor con Cédula"
    Case "09P": MsgBox "No debe hacer un Tiquete o Vale a un Proveedor con Pasaporte"
    Case "09O": MsgBox "No debe hacer un Tiquete o Vale a un Proveedor con Otro tipo de Documento"
    Case "10C": MsgBox "No debe hacer un Comprobante de Venta Art. 10 a un Proveedor con Cédula"
    Case "10P": MsgBox "No debe hacer un Comprobante de Venta Art. 10 con Pasaporte"
    Case "10O": MsgBox "No debe hacer un Comprobante de Venta Art. 10 con Otro tipo de Documento"
    Case "11C": MsgBox "No debe hacer un Pasaje a un Proveedor con Cédula"
    Case "11P": MsgBox "No debe hacer un Pasaje a un Proveedor con Pasaporte"
    Case "11O": MsgBox "No debe hacer un Pasaje a un Proveedor con Otro tipo de Documento"
    Case "12C": MsgBox "No debe hacer un Documento emitido por Inst. Financieros a un Proveedor con Cédula"
    Case "12P": MsgBox "No debe hacer un Documento emitido por Inst. Financieros a un Proveedor con Pasaporte"
    Case "12O": MsgBox "No debe hacer un Documento emitido por Inst. Financieros a un Proveedor con Otro tipo de Documento"
    Case "13C": MsgBox "No debe hacer un Documento emitido por Inst. Aseguradoras a un Proveedor con Cédula"
    Case "13P": MsgBox "No debe hacer un Documento emitido por Inst. Aseguradoras a un Proveedor con Pasaporte"
    Case "13O": MsgBox "No debe hacer un Documento emitido por Inst. Aseguradoras a un Proveedor con Otro tipo de Documento"
  End Select
  sSQL = "SELECT * FROM Trans_Retenciones " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Secuencial = '" & Secuencia & "' " _
       & "AND TD = '" & td & "' " _
       & "AND TT = '" & tt & "' " _
       & "AND Autorizacion = '" & Autorizacion & "' " _
       & "AND CodigoTR = 'IV' "
  SelectAdodc AdoRetencion, sSQL
  If AdoRetencion.Recordset.RecordCount > 0 Then
     MsgBox "Este comprobante ya existe"
     TextSec.Text = "0000000"
     TextSec.SetFocus
  End If
End Sub

Private Sub DGRet_GotFocus()
  'MsgBox "Hola"
  sSQL = "SELECT * " _
       & "FROM Asiento_R " _
       & "WHERE CodigoTR = 'IV' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND T_No = " & Trans_No & " " _
       & "AND Cta = '" & Cta_Ret_Egreso & "' "
  SelectDataGrid DGRet, AdoDetRet, sSQL
  'SelectData AdoDetRet, sSQL, False
  SetAddNew AdoDetRet
  Select Case Topc
  Case "1"
     Select Case td
       Case "R": SetFields AdoDetRet, "CodST", "01"
       Case "C": SetFields AdoDetRet, "CodST", "02"
       Case "P": SetFields AdoDetRet, "CodST", "03"
     End Select
     SetFields AdoDetRet, "MontoIVA1", CCur(TxtValorIVA.Text)
     SetFields AdoDetRet, "MontoIVA2", CCur(TBaseTransf.Text)
     SetFields AdoDetRet, "MontoRetIVA1", CCur(TValRetServ.Text)
     SetFields AdoDetRet, "TT", "C"
     SetFields AdoDetRet, "Valor_Fact", CCur(TxBaseImpo.Text)
     SetFields AdoDetRet, "BImpotcero", CCur(Txtbasecero.Text)
     SetFields AdoDetRet, "Codigo", CodigoCliente
     SetFields AdoDetRet, "FechaE", MBoxFechaE.Text
     SetFields AdoDetRet, "Fecha", MBoxFechaR.Text
     SetFields AdoDetRet, "MontoICE", CCur(TxtICE.Text)
     SetFields AdoDetRet, "PorRetIVA1", CodPorRetSer
     SetFields AdoDetRet, "Valor_Ret", CCur(TValRetTransf.Text) + CCur(TValRetServ.Text)
Case "2"
     SetFields AdoDetRet, "Codigo", CodigoCliente
     Select Case td
       Case "R": SetFields AdoDetRet, "CodST", "04"
       Case "C": SetFields AdoDetRet, "CodST", "05"
       Case "P": SetFields AdoDetRet, "CodST", "06"
       Case "O": SetFields AdoDetRet, "CodST", "07"
     End Select
     SetFields AdoDetRet, "MontoIVA1", CCur(TxtValorIVA.Text)
     SetFields AdoDetRet, "Fecha", BuscarFecha(UltimoDiaMes(MBoxFechaR.Text))
     SetFields AdoDetRet, "FechaE", BuscarFecha(UltimoDiaMes(MBoxFechaR.Text))
     SetFields AdoDetRet, "TT", "V"
     SetFields AdoDetRet, "Valor_Fact", CCur(TxBaseImpo.Text)
     SetFields AdoDetRet, "BImpotcero", CCur(Txtbasecero.Text)
     SetFields AdoDetRet, "MontoICE", CCur(TxtICE.Text)
     SetFields AdoDetRet, "MontoRetIVA1", CCur(TValRetServ.Text)
     'SetFields AdoDetRet, "Valor_Ret", CCur(TValRetTransf.Text) + CCur(TValRetServ.Text)
     SetFields AdoDetRet, "Valor_Ret", CCur(TxtValorIVA.Text) + CCur(TBaseTransf.Text)
     
Case "3"
     SetFields AdoDetRet, "FechaE", MBoxFechaE.Text
     SetFields AdoDetRet, "Fecha", MBoxFechaR.Text
     SetFields AdoDetRet, "Codigo", SinEspaciosDer(DCPovCli.Text)
     SetFields AdoDetRet, "CodST", "09"
     SetFields AdoDetRet, "Valor_Fact", CCur(TxBaseImpo.Text)
     SetFields AdoDetRet, "TT", "I"
     SetFields AdoDetRet, "MontoICE", CCur(TxtICE.Text)
     SetFields AdoDetRet, "MontoIVA1", CCur(TValRetServ.Text)
     If opcion1 Then
      SetFields AdoDetRet, "ConvInt", "S"
     Else
      SetFields AdoDetRet, "ConvInt", "N"
    End If
Case "4"
    SetFields AdoDetRet, "FechaE", MBoxFechaE.Text
    SetFields AdoDetRet, "Fecha", MBoxFechaR.Text
    SetFields AdoDetRet, "CodST", "08"
    SetFields AdoDetRet, "SerieEx", TxBaseImpo.Text
    SetFields AdoDetRet, "TT", "E"
    SetFields AdoDetRet, "Codigo", "01"
    SetFields AdoDetRet, "Valor_Fact", CCur(TValRetServ.Text)
    If opcion1 Then
       SetFields AdoDetRet, "ConvInt", "S"
     Else
       SetFields AdoDetRet, "ConvInt", "N"
     End If
End Select
     If OpcS Then
       SetFields AdoDetRet, "Dev", "S"
     Else
       SetFields AdoDetRet, "Dev", "N"
     End If
     Select Case SinEspaciosIzq(DCTComp.Text)
        Case "04", "05"
        SetFields AdoDetRet, "Cambio", SinEspaciosIzq(DCCompMod.Text) & SinEspaciosIzq(DCNumMod.Text)
     End Select
     If Val(TxtValorIVA.Text) > 0 Then
        PorcIva = Val(CPorRetServ.Text)
     Else
        PorcIva = Val(CPorRetTransf.Text)
     End If
     'MsgBox PorcIva
     Cadena = CStr(Val(PorcIva) / 100)
     Cadena1 = "IV"
     SetFields AdoDetRet, "Porc", Cadena
     SetFields AdoDetRet, "CodPorc", CodPorcIva
     SetFields AdoDetRet, "CodigoTR", Cadena1
     SetFields AdoDetRet, "TD", SinEspaciosIzq(DCTComp.Text)
     SetFields AdoDetRet, "Serie", TextSerie.Text
     SetFields AdoDetRet, "Secuencial", TextSec.Text
     SetFields AdoDetRet, "Autorizacion", TextAuto.Text
     SetFields AdoDetRet, "Aduana", TextAdu.Text
     SetFields AdoDetRet, "MontoIVA2", CCur(TBaseTransf.Text)
     SetFields AdoDetRet, "PorRetIVA2", CodPorRetTrans
     SetFields AdoDetRet, "MontoRetIVA2", CCur(TValRetTransf.Text)
     SetFields AdoDetRet, "Item", NumEmpresa
     SetFields AdoDetRet, "CodigoU", CodigoUsuario
     SetFields AdoDetRet, "IdenCT", SinEspaciosIzq(ComboCT.Text)
     SetFields AdoDetRet, "Cta", Cta_Ret_Egreso
     SetFields AdoDetRet, "T_No", Trans_No
     SetFields AdoDetRet, "R_No", Ret_No
     Ret_No = Ret_No + 1
     SetUpdate AdoDetRet
End Sub

Private Sub Form_Activate()
RatonNormal
SQL1 = "SELECT * " _
    & "FROM Asiento_R " _
    & "WHERE CodigoTR = 'IV' " _
    & "AND Item = '" & NumEmpresa & "' " _
    & "AND T_No = " & Trans_No & " " _
    & "AND Cta = '" & Cta_Ret_Egreso & "' "
SelectDataGrid DGRet, AdoDetRet, SQL1
DGRet.Visible = True
Topc = 1
RatonNormal
VerTipoTrans 1
End Sub

Private Sub Form_Deactivate()
ConectarAdoExecute SQL2
End Sub

Private Sub Form_Load()
 CentrarForm FRetIVA
 ConectarAdodc AdoConcepto
 ConectarAdodc AdoRetenido
 ConectarAdodc AdoDetRet
 ConectarAdodc AdoRetencion
 ConectarAdodc AdoNumEg
 ConectarAdodc AdoQuery1
End Sub

Private Sub MBoxFechaE_GotFocus()
  MarcarTexto MBoxFechaE
End Sub

Private Sub MBoxFechaE_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyEscape: CmdSalir.SetFocus
  End Select
End Sub

Private Sub MBoxFechaE_LostFocus()
  FechaValida MBoxFechaE
  VAnio = FechaAnio(MBoxFechaE)
  If VAnio <= 1994 Then MsgBox "La fecha ingresada no es válida, debe ingresar un año superior a 1994"
'    MBoxFechaE.SetFocus
'  Else
'    MBoxFechaR.SetFocus
'  End If
End Sub
Private Sub MBoxFechaR_GotFocus()
   MarcarTexto MBoxFechaR
End Sub

Private Sub MBoxFechaR_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyEscape: CmdSalir.SetFocus
  End Select
End Sub

Private Sub MBoxFechaR_LostFocus()
 FechaValida MBoxFechaR
 Periodo = CalculoPeriodo(MBoxFechaR.Text)
 I = CFechaLong(MBoxFechaE.Text)
 J = CFechaLong(MBoxFechaR.Text)
  If I > J Then MsgBox "La fecha ingresada no es válida, debe ser superior a la fecha de emisión del comprobante"
'    MBoxFechaR.SetFocus
'  Else
'    Select Case Topc
'    Case "1"
'        TextSerie.Visible = True
'        TextSerie.SetFocus
'    Case "2"
'        TextSerie.Visible = False
'        TextSec.SetFocus
''DCTComp.SetFocus
'    Case "3"
'        TextSec.SetFocus
'    End Select
'  End If
End Sub

Private Sub Opctrans_Click(Index As Integer)
   VerTipoTrans (Index + 1)
End Sub

Private Sub TValRetServ_LostFocus()
    TextoValido TValRetServ
    IvaServ = TValRetServ.Text
End Sub

Private Sub TValRetTransf_LostFocus()
TextoValido TValRetTransf
IvaTrans = TValRetTransf.Text
Apagar = IvaTotal - IvaServ - IvaTrans
LApagar.Caption = Format(Apagar, "#,##0.00")
End Sub

Private Sub TxBaseImpo_LostFocus()
Select Case Topc
Case "1", "2", "3"
    BaseImpo = Format(TxBaseImpo.Text, "#,##0.00")
Case "4"
   If Len(TxBaseImpo.Text) <> 7 Then
     MsgBox ("La longitud es incorrecta, debe ser 7 caracteres")
     TxBaseImpo.Text = "0000000"
     TxBaseImpo.SetFocus
  End If
End Select
End Sub

Private Sub TxtValorIVA_GotFocus()
 IvaTotal = Round(IvaTotal, 2)
 ValorIva = IvaTotal
 TxtValorIVA.Text = ValorIva
 MarcarTexto TxtValorIVA
End Sub

Private Sub TxtValorIVA_LostFocus()
BaseServ = Round(Val(TxtValorIVA.Text), 2)
If BaseServ > IvaTotal Then
  MsgBox "El valor ingresado no es correcto"
  TxtValorIVA.SetFocus
Else
  BaseTrans = Round(ValorIva - BaseServ, 2)
End If
TxtValorIVA.Text = Format(TxtValorIVA.Text, "#,##0.00")
End Sub

Private Sub TxBaseImpo_GotFocus()
Select Case Topc
  Case "4"
    TxBaseImpo.ForeColor = &H80000008
End Select
MarcarTexto TxBaseImpo
End Sub

Private Sub Text3_GotFocus()
 MarcarTexto Text3
End Sub

Private Sub TValRetServ_GotFocus()
Select Case Topc
 Case "1", "2"
    TValRetServ.Text = Round(IvaServ, 2)
    TValRetServ.Text = Format(TValRetServ.Text, "#,##0.00")
Case "3", "4"
     TValRetServ.Text = "0.00"
End Select
MarcarTexto TValRetServ
End Sub

Private Sub TValRetTransf_GotFocus()
MarcarTexto TValRetTransf
IvaTrans = Round(BaseTrans * PorRetTran / 100, 2)
TValRetTransf.Text = Format(Round(IvaTrans, 2), "#,##0.00")
IvaTrans = Round(Val(TValRetTransf.Text), 2)
TBaseTransf.Text = Format(TBaseTransf.Text, "#,##0.00")
MarcarTexto TValRetTransf
End Sub

Private Sub TBaseTransf_GotFocus()
BaseTrans = Round(IvaTotal - BaseServ, 2)
Select Case Topc
Case "1"
If BaseTrans > 0 Then
  TBaseTransf.Text = BaseTrans
  BaseTrans = Round(Val(TBaseTransf.Text), 2)
Else
  TBaseTransf.Text = 0
End If
End Select
  MarcarTexto TBaseTransf
End Sub

Private Sub TBaseTransf_LostFocus()
TextoValido TBaseTransf
BaseTransf = TBaseTransf.Text
Parcial = BaseTransf + BaseServ
If BaseTransf <> BaseTrans And Parcial <> IvaTotal Then
    MsgBox "El valor introducido no es correcto"
    Parcial = Parcial - BaseTransf
    TBaseTransf.SetFocus
Else
    BaseTrans = Format(Val(TBaseTransf.Text), "#,##0.00")
End If
End Sub

Private Sub TextSec_GotFocus()
  MarcarTexto TextSec
End Sub

Private Sub TextSec_LostFocus()
  If Val(TextSec.Text) < 1 Then
    MsgBox ("El valor introducido es incorrecto debe ser igual o mayor a 1")
    TextSec.Text = "0000001"
    TextSec.SetFocus
  End If
  Cadena = Format(CStr(TextSec.Text), "0000000")
  TextSec.Text = Cadena
'  MsgBox Cadena
  Select Case Topc
    Case "1": tt = "C"
    Case "3": tt = "I"
    Case "4": tt = "E"
    Case "2": tt = "V"
  End Select
Secuencia = TextSec.Text
End Sub

Private Sub TextAdu_GotFocus()
  MarcarTexto TextAdu
End Sub

Private Sub TextAdu_LostFocus()
  If Val(TextAdu.Text) < 1 Then
    MsgBox ("El valor introducido es incorrecto")
    TextAdu.SetFocus
End If
TextAdu.Text = Format(TextAdu.Text, "0000000000000000")
End Sub

Private Sub TextSerie_GotFocus()
  MarcarTexto TextSerie
End Sub

Private Sub TextSerie_LostFocus()
  If Val(TextSerie.Text) < 1001 Then
     MsgBox ("El valor introducido es incorrecto, el numero de serie del comprobante debe ser mayor o igual que 001001")
     TextSerie.Text = "001001"
     TextSerie.SetFocus
  End If
  Cadena = Format(TextSerie.Text, "000000")
  TextSerie = Cadena
End Sub

Private Sub TextAuto_GotFocus()
  MarcarTexto TextAuto
End Sub

Private Sub TextAuto_LostFocus()
  If Val(TextAuto.Text) < 1 Then
     MsgBox ("El valor introducido es incorrecto")
     'TextAuto.SetFocus
  End If
  TextAuto.Text = Format(TextAuto.Text, "0000000000")
End Sub

Private Sub TxtICE_GotFocus()
Select Case Topc
 Case "3"
    TxtICE.Text = Format(IvaTotal, "#,##0.00")
End Select
  MarcarTexto TxtICE
End Sub

Private Sub Txtbasecero_GotFocus()
  MarcarTexto Txtbasecero
End Sub
