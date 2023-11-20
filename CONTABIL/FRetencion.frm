VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FRetencion 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Retencion en la Fuente"
   ClientHeight    =   4305
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8625
   FillColor       =   &H00FFC0C0&
   Icon            =   "FRetencion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "FRetencion.frx":030A
   ScaleHeight     =   4638.385
   ScaleMode       =   0  'User
   ScaleWidth      =   8683.873
   Begin VB.TextBox TextValorRet 
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
      Left            =   5880
      MaxLength       =   10
      TabIndex        =   40
      Text            =   "0.00"
      Top             =   3570
      Width           =   1485
   End
   Begin VB.TextBox TxPorc 
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
      TabIndex        =   38
      Top             =   3570
      Width           =   540
   End
   Begin VB.TextBox TBaseImpo 
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
      Left            =   3990
      MaxLength       =   10
      TabIndex        =   36
      Text            =   "0.00"
      Top             =   3570
      Width           =   1380
   End
   Begin VB.TextBox TxtAutoriza 
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
      ForeColor       =   &H00C00000&
      Height          =   330
      Left            =   2625
      MaxLength       =   10
      TabIndex        =   34
      Text            =   "0"
      Top             =   3570
      Width           =   1380
   End
   Begin VB.TextBox TextNumComp 
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
      ForeColor       =   &H00C00000&
      Height          =   330
      Left            =   1470
      MaxLength       =   10
      TabIndex        =   32
      Text            =   "0"
      Top             =   3570
      Width           =   1170
   End
   Begin MSDataListLib.DataCombo DCConcepto 
      Bindings        =   "FRetencion.frx":0614
      DataSource      =   "AdoConcepto"
      Height          =   315
      Left            =   105
      TabIndex        =   12
      Top             =   1680
      Width           =   8415
      _ExtentX        =   14843
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
   Begin VB.ComboBox CCtaRet 
      BackColor       =   &H00C0E0FF&
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
      Height          =   315
      Left            =   1575
      TabIndex        =   10
      Text            =   "Combo1"
      Top             =   1050
      Width           =   6945
   End
   Begin MSMask.MaskEdBox MBoxFechaI 
      Height          =   330
      Left            =   1575
      TabIndex        =   4
      Top             =   735
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   582
      _Version        =   393216
      ForeColor       =   255
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Height          =   645
      Left            =   105
      TabIndex        =   0
      Top             =   0
      Width           =   8415
      Begin VB.Label LblProv 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Proveedor:"
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
         Top             =   210
         Width           =   6840
      End
      Begin VB.Label Label30 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Proveedor:"
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
         TabIndex        =   1
         Top             =   210
         Width           =   1380
      End
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H008080FF&
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
      Left            =   7455
      Picture         =   "FRetencion.frx":062E
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   2205
      Width           =   1065
   End
   Begin VB.OptionButton Opctrans 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Compras"
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
      Left            =   2940
      TabIndex        =   5
      Top             =   735
      Value           =   -1  'True
      Width           =   1065
   End
   Begin VB.OptionButton Opctrans 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Ventas"
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
      Index           =   1
      Left            =   4095
      TabIndex        =   6
      Top             =   735
      Width           =   960
   End
   Begin VB.OptionButton Opctrans 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Exportaciones"
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
      Index           =   3
      Left            =   6825
      TabIndex        =   8
      Top             =   735
      Width           =   1695
   End
   Begin VB.OptionButton Opctrans 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Importaciones"
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
      Index           =   2
      Left            =   5145
      TabIndex        =   7
      Top             =   735
      Width           =   1590
   End
   Begin VB.Frame FConcep 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Retenido por"
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
      Left            =   1470
      TabIndex        =   16
      ToolTipText     =   "Tipo de Documento para Retencion"
      Top             =   2100
      Width           =   5895
      Begin VB.OptionButton OpNotVen 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Nota de Venta"
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
         Left            =   1050
         TabIndex        =   18
         Top             =   525
         Width           =   1695
      End
      Begin VB.OptionButton OpLiqui 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Liquidacion de Compra"
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
         Left            =   3255
         TabIndex        =   19
         Top             =   210
         Width           =   2430
      End
      Begin VB.OptionButton OpOtros 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Otros"
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
         Left            =   3255
         TabIndex        =   20
         Top             =   525
         Width           =   2010
      End
      Begin VB.OptionButton OpFac 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Factura"
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
         Left            =   1050
         TabIndex        =   17
         Top             =   210
         Width           =   2220
      End
   End
   Begin VB.TextBox TIngLiqui 
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
      Left            =   1470
      MaxLength       =   10
      TabIndex        =   22
      Text            =   "0.00"
      Top             =   2625
      Width           =   1590
   End
   Begin VB.TextBox TApoIESS 
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
      Left            =   3990
      MaxLength       =   10
      TabIndex        =   26
      Text            =   "0.00"
      Top             =   2625
      Width           =   1695
   End
   Begin VB.TextBox Txtporcap 
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
      Left            =   3045
      MaxLength       =   10
      TabIndex        =   24
      Text            =   "0.00"
      Top             =   2625
      Width           =   960
   End
   Begin VB.TextBox TxtSerieCR 
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
      ForeColor       =   &H00C00000&
      Height          =   330
      Left            =   105
      MaxLength       =   10
      TabIndex        =   30
      Text            =   "0"
      Top             =   3570
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H008080FF&
      Caption         =   "&Salir"
      DisabledPicture =   "FRetencion.frx":0A70
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
      Left            =   7455
      MouseIcon       =   "FRetencion.frx":14BA
      Picture         =   "FRetencion.frx":1F04
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   3045
      Width           =   1065
   End
   Begin MSAdodcLib.Adodc AdoRetenido 
      Height          =   330
      Left            =   210
      Top             =   0
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
   Begin VB.Frame FrmSN 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Salario Neto"
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
      Height          =   855
      Left            =   105
      TabIndex        =   13
      Top             =   2100
      Width           =   1275
      Begin VB.OptionButton OpcN 
         BackColor       =   &H00C0E0FF&
         Caption         =   "No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   15
         ToolTipText     =   "El Empleado paga el aporte"
         Top             =   480
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton OpcS 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Si"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   14
         ToolTipText     =   "Si es salario Neto"
         Top             =   240
         Width           =   855
      End
   End
   Begin MSAdodcLib.Adodc AdoRetencion 
      Height          =   330
      Left            =   210
      Top             =   315
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
      Left            =   2310
      Top             =   0
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
   Begin VB.TextBox TxtOtrosIng 
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
      Left            =   5670
      MaxLength       =   10
      TabIndex        =   28
      Text            =   "0.00"
      Top             =   2625
      Width           =   1695
   End
   Begin MSAdodcLib.Adodc AdoConcepto 
      Height          =   330
      Left            =   4305
      Top             =   0
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
   Begin MSAdodcLib.Adodc AdoQuery1 
      Height          =   330
      Left            =   2415
      Top             =   315
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
      Caption         =   "PorcDep"
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
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Valor Retenido"
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
      Height          =   540
      Left            =   5880
      TabIndex        =   39
      Top             =   3045
      Width           =   1485
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "%"
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
      Height          =   540
      Left            =   5355
      TabIndex        =   37
      Top             =   3045
      Width           =   540
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Base Imponible Rentas"
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
      Height          =   540
      Left            =   3990
      TabIndex        =   35
      ToolTipText     =   "Base para Retencíon"
      Top             =   3045
      Width           =   1380
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "No Autorizacion"
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
      Height          =   540
      Left            =   2625
      TabIndex        =   33
      ToolTipText     =   "Autorizacion del Comprobante de Retención"
      Top             =   3045
      Width           =   1380
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "No Comp. Retencion"
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
      Height          =   540
      Left            =   1470
      TabIndex        =   31
      ToolTipText     =   "Número de comprobante de Retención"
      Top             =   3045
      Width           =   1170
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Concepto:"
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
      Left            =   105
      TabIndex        =   11
      ToolTipText     =   "Por el concepto de:"
      Top             =   1365
      Width           =   8415
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Cta. Retencion:"
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
      Top             =   1050
      Width           =   1485
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fecha:"
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
      Top             =   735
      Width           =   1485
   End
   Begin VB.Label Label16 
      BackColor       =   &H00C0E0FF&
      Caption         =   "* Son Items Obligatorios a partir del 2006"
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
      Height          =   225
      Left            =   105
      TabIndex        =   41
      Top             =   3990
      Width           =   4740
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Salario Básico"
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
      Height          =   435
      Left            =   1470
      TabIndex        =   21
      Top             =   2205
      Width           =   1590
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Aporte Personal"
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
      Height          =   435
      Left            =   3990
      TabIndex        =   25
      Top             =   2205
      Width           =   1695
   End
   Begin VB.Label Label10 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "% Aporte Personal"
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
      TabIndex        =   23
      Top             =   2205
      Width           =   960
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Serie  No. C.Retencion"
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
      Height          =   540
      Left            =   105
      TabIndex        =   29
      ToolTipText     =   "Serie Comprobante de Retención"
      Top             =   3045
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Otros Ingresos"
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
      Left            =   5670
      TabIndex        =   27
      Top             =   2205
      Width           =   1695
   End
End
Attribute VB_Name = "FRetencion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Porcen As Single
Dim Valor As Integer

Private Sub Command1_Click()
  Unload FRetencion
End Sub

Private Sub Command3_Click()
  Cta = SinEspaciosIzq(CCtaRet)
  LeerCta Cta
  sSQL = "DELETE * " _
       & "FROM Asiento_R " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND Cta = '" & Cta & "' " _
       & "AND TD = '" & SinEspaciosIzq(DCConcepto) & "' " _
       & "AND T_No = " & Trans_No & " "
  ConectarAdoExecute sSQL
  sSQL = "DELETE * " _
       & "FROM Asiento_RP " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND Cta = '" & Cta & "' " _
       & "AND T_No = " & Trans_No & " "
  ConectarAdoExecute sSQL
  If IsDep Then
     SetAdoAddNew "Asiento_RP"
     SetAdoFields "Fecha", MBoxFechaI.Text
     SetAdoFields "FechaE", MBoxFechaI.Text
     SetAdoFields "AporteP", TApoIESS.Text
     SetAdoFields "PorAp", Val(Txtporcap.Text) / 100
     SetAdoFields "InLiqui", CCur(TBaseImpo) + CCur(TApoIESS)
     SetAdoFields "Salario", TBaseImpo.Text
     SetAdoFields "Codigo", CodigoCliente
     SetAdoFields "CodigoU", CodigoUsuario
     SetAdoFields "Item", NumEmpresa
     SetAdoFields "Retencion_No", TextNumComp.Text
     SetAdoFields "Porc_Ret", Val(TxPorc.Text) / 100
     SetAdoFields "Retenido", TextValorRet.Text
     SetAdoFields "ID", NumRet
     If OpcS Then SetAdoFields "SN", "2" Else SetAdoFields "SN", "1"
     SetAdoFields "Cta", Cta
     SetAdoFields "T_No", Trans_No
     SetAdoUpdate
  End If   ' retenciones en trans retenciones
  SetAdoAddNew "Asiento_R"
  SetAdoFields "Codigo", CodigoCliente
  SetAdoFields "Fecha", MBoxFechaI.Text
  SetAdoFields "TD", SinEspaciosIzq(DCConcepto.Text)
  SetAdoFields "Valor_Fact", TBaseImpo.Text
  SetAdoFields "Porc", Porcen / 100
  Select Case Porcen
    Case 1: SetAdoFields "CodPorc", "1"
    Case 5: SetAdoFields "CodPorc", "2"
    Case 8: SetAdoFields "CodPorc", "3"
    Case 25: SetAdoFields "CodPorc", "4"
    Case Else: SetAdoFields "CodPorc", "0"
  End Select
  SetAdoFields "Valor_Ret", TextValorRet.Text
  SetAdoFields "Item", NumEmpresa
  SetAdoFields "Secuencial", TextNumComp.Text
  SetAdoFields "Autorizacion", TxtAutoriza.Text
  SetAdoFields "CodigoU", CodigoUsuario
  SetAdoFields "Retencion_No", TextNumComp.Text
  SetAdoFields "ID", NumRet
  NumRet = NumRet + 1
  If OpFac Then SetAdoFields "TT", "F"
  If OpNotVen Then SetAdoFields "TT", "N"
  If OpLiqui Then SetAdoFields "TT", "L"
  If OpOtros Then SetAdoFields "TT", "T"
  SetAdoFields "Porc", Val(TxPorc.Text) / 100
  Select Case Topc
    Case 1: SetAdoFields "CodigoTR", "RF"
    Case 2: SetAdoFields "CodigoTR", "RV"
    Case 3: SetAdoFields "CodigoTR", "RI"
    Case 4: SetAdoFields "CodigoTR", "RE"
    Case Else
         SetAdoFields "CodigoTR", "RF"
         SetAdoFields "TT", "R"
  End Select
  SetAdoFields "IdenCT", "00"
  SetAdoFields "SCT", "00"
  SetAdoFields "Serie", TxtSerieCR.Text
  SetAdoFields "PorRetIVA1", "0"
  SetAdoFields "PorRetIVA2", "0"
  SetAdoFields "Aduana", "0"
  SetAdoFields "Dev", "N"
  SetAdoFields "Cambio", "0"
  SetAdoFields "ConvInt", "N"
  SetAdoFields "CPorcIce", "0"
  SetAdoFields "IdenCT", "00"
  SetAdoFields "Cta", Cta
  SetAdoFields "T_No", Trans_No
  SetAdoUpdate
  Unload FRetencion
End Sub

Private Sub DCConcepto_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCConcepto_LostFocus()
  If KeyCode = vbKeyEscape Then: Command1.SetFocus
End Sub

Private Sub Form_Activate()
  LblProv.Caption = " " & NombreCliente & Space(60 - Len(NombreCliente)) & TipoDoc & CodigoCliente
  If TipoDoc <> "R" Then
     Mensajes = "Por Relación de Dependencia"
     Titulo = "PREGUNTA DE VERIFICACION"
     If BoxMensaje = vbYes Then IsDep = True Else IsDep = False
  End If
  sSQL = "SELECT * " _
       & "FROM Catalogo_Cuentas " _
       & "WHERE TC = 'RF' " _
       & "AND DG = 'D' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "ORDER BY TC,Codigo "
  SelectAdodc AdoRetenido, sSQL
  CCtaRet.Clear
  With AdoRetenido.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          CCtaRet.AddItem .Fields("Codigo") & " - " & .Fields("Cuenta")
         .MoveNext
       Loop
   End If
  End With
  CCtaRet.Text = CCtaRet.List(0)
'  MsgBox SerieRet & vbCrLf & AutorizaRet & vbCrLf & NumRet
  If IsDep Then
     FRetencion.Caption = "Formulario de Retención por Relación de Dependencia"
     Opctrans(1).Visible = False
     Opctrans(2).Visible = False
     Opctrans(3).Visible = False
     Opctrans(0).Visible = False
     DCConcepto.Enabled = False
     FConcep.Visible = False
     Label12.Caption = "Valor Total Pagado"
     Label13.Caption = "Valor Retenido"
'     Label3.Caption = "Total Pagos"
     Label8.Visible = False
     'Label21.Caption = "Comprob. No"
     Label21.Visible = True
     IngLiqui = 0
     BaseImpor = 0
     BaseImpon = 0
     Porcen = 0
     ValRet = 0
     AporIess = 0
  Else
     Opctrans(1).Visible = True
     Opctrans(2).Visible = True
     Opctrans(3).Visible = True
     Opctrans(0).Visible = True
     IngLiqui = 0
     BaseImpor = 0
     BaseImpon = 0
     Porcen = 0
     ValRet = 0
     AporIess = 0
     Label9.Visible = False
     Label18.Visible = False
     Label10.Visible = False
     Label11.Visible = False
     TApoIESS.Visible = False
     TIngLiqui.Visible = False
     TxtOtrosIng.Visible = False
     Txtporcap.Visible = False
     DCConcepto.Enabled = True
     FRetencion.Caption = "Formulario de Retención por Otros Conceptos"
     FrmSN.Visible = False
     Label12.Caption = "Valor de la Factura"
     Label13.Caption = "Valor Retenido"
  '   Label3.Caption = "Total Facturado"
     Label8.Visible = True
     Label8.Caption = "No Serie C. Retencion"
     TxtSerieCR.Visible = True
  End If
  'MsgBox Vanio
  SQL1 = "SELECT (Codigo & '     ' & Tipo_Ret) As CodRet " _
       & "FROM Tipo_Reten " _
       & "WHERE Item = '000' " _
       & "AND Año = '" & Vanio & "' "
  'MsgBox IsDep
  If IsDep Then
     SQL1 = SQL1 & "AND Codigo = 'DEP' "
  Else
     SQL1 = SQL1 & "AND Codigo <> 'DEP' "
  End If
  SQL1 = SQL1 & "ORDER BY Codigo"
  SelectDBCombo DCConcepto, AdoConcepto, SQL1, "CodRet"
  MBoxFechaI.SetFocus
End Sub

Private Sub Form_Load()
  CentrarForm FRetencion
  ConectarAdodc AdoConcepto
  ConectarAdodc AdoRetenido
  ConectarAdodc AdoDetRet
  ConectarAdodc AdoRetencion
  ConectarAdodc AdoQuery1
   If IngMensual Then
      Cadena1 = Mid(Periodo, 4, 2)
      Cadena = MesesLetras(Val(Cadena1))
      Cadena = Cadena & " de " & Mid(Periodo, 7, 4)
   Else
      Cadena = Mid(Periodo, 7, 4)
   End If
   Label1.Caption = ""
   Label1.Caption = "Período: " & Cadena
   Vanio = Mid(Cadena, Len(Cadena) - 3, 4)
  If EsIVA Then
    Lblcuenta.Visible = False
    TxtCuenta.Visible = False
End If
  sSQL = "SELECT MAX(Id) As ID " _
       & "FROM Trans_Retenciones " _
       & "WHERE Item = '" & NumEmpresa & "' "
  SelectAdodc AdoRetencion, sSQL
  If IsNull(AdoRetencion.Recordset.Fields("ID")) Then
     NumRet = 1
  Else
     NumRet = AdoRetencion.Recordset.Fields("ID") + 1
  End If
  If Val(SerieRet) < 1000 Then
   SQL2 = "SELECT Serie,Autorizacion " _
        & "FROM Trans_Retenciones " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND Codigo = '" & CodigoCliente & "' " _
        & "AND CodigoTR = 'RF' " _
        & "ORDER BY Fecha "
   SelectAdodc AdoQuery1, SQL2
   With AdoQuery1.Recordset
    If .RecordCount > 0 Then
       .MoveLast
        SerieRet = .Fields("Serie")
        AutorizaRet = .Fields("Autorizacion")
    End If
   End With
End If
End Sub

Private Sub MBoxFechaI_GotFocus()
  MarcarTexto MBoxFechaI
End Sub

Private Sub MBoxFechaI_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyEscape Then: Command1.SetFocus
  PresionoEnter KeyCode
End Sub

Private Sub MBoxFechaI_LostFocus()
  FechaValida MBoxFechaI
  Vanio = Format(Year(MBoxFechaI), "0000")
  'MsgBox Vanio
  SQL1 = "SELECT (Codigo & '     ' & Tipo_Ret) As CodRet " _
       & "FROM Tipo_Reten " _
       & "WHERE Item = '000' " _
       & "AND Año = '" & Vanio & "' "
  'MsgBox IsDep
  If IsDep Then
     SQL1 = SQL1 & "AND Codigo = 'DEP' "
  Else
     SQL1 = SQL1 & "AND Codigo <> 'DEP' "
  End If
  SQL1 = SQL1 & "ORDER BY Codigo"
  SelectDBCombo DCConcepto, AdoConcepto, SQL1, "CodRet"
End Sub

Private Sub OpcN_Click()
 Txtporcap.Enabled = True
 TApoIESS.Enabled = True
End Sub

Private Sub OpcN_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub OpcS_Click()
 Txtporcap.Enabled = False
 TApoIESS.Enabled = False
End Sub

Private Sub OpcS_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub Opctrans_Click(Index As Integer)
   Topc = (Index + 1)
End Sub

Private Sub Opctrans_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub OpFac_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub OpLiqui_KeyDown(KeyCode As Integer, Shift As Integer)
    PresionoEnter KeyCode
End Sub

Private Sub OpNotVen_KeyDown(KeyCode As Integer, Shift As Integer)
     PresionoEnter KeyCode
End Sub

Private Sub OpOtros_KeyDown(KeyCode As Integer, Shift As Integer)
     PresionoEnter KeyCode
End Sub

Private Sub TApoIESS_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TApoIESS_LostFocus()
       TApoIESS.Text = Format(TApoIESS, "#,##0.00")
       AporIess = TApoIESS.Text
End Sub

Private Sub TBaseImpo_KeyDown(KeyCode As Integer, Shift As Integer)
     PresionoEnter KeyCode
End Sub

Private Sub TBaseImpo_GotFocus()
  If IsDep Then
     If OpcS Then AporIess = 0
  End If
  BaseImpon = TIngLiqui.Text - TApoIESS.Text + TxtOtrosIng.Text
  TBaseImpo.Text = BaseImpon
  MarcarTexto TBaseImpo
End Sub

Private Sub TBaseImpo_LostFocus()
  If IsDep Then
     BaseImpon = Format(BaseImpon, "#,##0.00") ' DEPENDENCIA
     BaseImpor = Format((IngLiqui) - (AporIess) + (TxtOtrosIng.Text), "#,##0.00")
     TBaseImpo.Text = BaseImpon
     If BaseImpon > BaseImpor Then
        MsgBox "Valor Incorrecto"
        TBaseImpo.Text = BaseImpor
        TBaseImpo.SetFocus
     End If
     Anio = Mid(Periodo, 7, 4)
     If IngMensual Then
       SQL2 = "SELECT Desde/12 As Desde, Hasta/12 As Hasta, Basico/12 As Basico, Excede "
     Else
       SQL2 = "SELECT Desde, Hasta, Basico, Excede "
     End If
     SQL2 = SQL2 & "FROM Tabla_Renta " _
                 & "WHERE Año = '" & Anio & "' " _
                 & "ORDER BY Hasta DESC "
     SelectAdodc AdoQuery1, SQL2
        With AdoQuery1.Recordset
            .MoveFirst
            Encontro = False
            Do While Not Encontro
               If (.Fields("Desde") < BaseImpor) And (.Fields("Hasta") > BaseImpor) Then
                  ValRet = ((BaseImpor - .Fields("Desde")) * .Fields("Excede") / 100) + .Fields("Basico")
                  TxPorc.Text = .Fields("Excede")
                  Encontro = True
'                  MsgBox BaseImpor
                Else
                  .MoveNext
                End If
            Loop
        End With
  Else
     BaseImpor = TBaseImpo.Text
  End If
  TBaseImpo.Text = Format(TBaseImpo.Text, "#,##0.00")
End Sub

Private Sub TextNumComp_GotFocus()
   MarcarTexto TextNumComp
End Sub

Private Sub TextNumComp_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub TextNumComp_LostFocus()
   TextoValido TextNumComp, True
   If TextNumComp < 1 Then
      MsgBox "Numero de Comprobante Inválido"
      TextNumComp.SetFocus
   End If
   TextNumComp.Text = Format(TextNumComp, "0000000")
End Sub

Private Sub TIngLiqui_GotFocus()
    MarcarTexto TIngLiqui
End Sub

Private Sub TApoIESS_GotFocus()
    AporIess = IngLiqui * Porcen / 100
    TApoIESS.Text = Val(AporIess)
    MarcarTexto TApoIESS
End Sub

Private Sub TIngLiqui_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TIngLiqui_LostFocus()
  IngLiqui = 0
  TIngLiqui.Text = Format(TIngLiqui.Text, "#,##0.00")
  IngLiqui = TIngLiqui.Text
End Sub

Private Sub TxPorc_GotFocus()
  MarcarTexto TxPorc
End Sub
Private Sub TxPorc_KeyDown(KeyCode As Integer, Shift As Integer)
     PresionoEnter KeyCode
End Sub

Private Sub TxPorc_LostFocus()
   TextoValido TxPorc, True
   Porcen = Format(Val(TxPorc.Text), "#,##0.00")
   If IsDep Then
   Else
     ValRet = BaseImpor * Porcen / 100
   End If
End Sub

Private Sub TxtAutoriza_GotFocus()
   TxtAutoriza = AutorizaRet
   MarcarTexto TxtAutoriza
End Sub

Private Sub TxtAutoriza_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub TxtAutoriza_LostFocus()
'   TextoValido TxtAutoriza, True
   TxtAutoriza.Text = Format(TxtAutoriza, "0000000000")
End Sub

Private Sub TxtOtrosIng_GotFocus()
   MarcarTexto TxtOtrosIng
End Sub

Private Sub TxtOtrosIng_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtOtrosIng_LostFocus()
     TxtOtrosIng.Text = Format(TxtOtrosIng.Text, "#,##0.00")
End Sub

Private Sub Txtporcap_GotFocus()
  If OpcS.Value Then Porcen = 0 Else Porcen = 9.35
  Txtporcap.Text = Porcen
  MarcarTexto Txtporcap
End Sub

Private Sub TextValorRet_GotFocus()
  TextValorRet.Text = Val(ValRet)
  MarcarTexto TextValorRet
End Sub

Private Sub TextValorRet_LostFocus()
  TextValorRet.Text = Format(TextValorRet.Text, "#,##0.00")
End Sub

Private Sub TextValorRet_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyEscape Then: Command1.SetFocus
  PresionoEnter KeyCode
End Sub

Private Sub TxtPorcAp_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub TxtPorcAp_LostFocus()
  TextoValido Txtporcap, True
  If OpcS.Value Then Porcen = 0 Else Porcen = Format(Val(Txtporcap.Text), "#,##0.00")
  PresionoEnter KeyCode
End Sub

Private Sub TxtSerieCR_GotFocus()
    TxtSerieCR = SerieRet
    MarcarTexto TxtSerieCR
End Sub

Private Sub TxtSerieCR_KeyDown(KeyCode As Integer, Shift As Integer)
    PresionoEnter KeyCode
End Sub

Private Sub TxtSerieCR_LostFocus()
   TextoValido TxtSerieCR, True
   If TxtSerieCR < 1000 Then
     MsgBox "Numero de Serie Incorrecto"
     TxtSerieCR.Text = "001001"
     TxtSerieCR.SetFocus
   Else
     TxtSerieCR.Text = Format(TxtSerieCR, "000000")
   End If
End Sub
