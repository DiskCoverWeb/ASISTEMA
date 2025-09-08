VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form FInsPreFacturas 
   BackColor       =   &H00808000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "C:\SISTEMA\BASES\UPDATE_DB"
   ClientHeight    =   3870
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8415
   Icon            =   "FInsPreFacturas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   8415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TxtDescuento2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   330
      Index           =   2
      Left            =   5880
      MaxLength       =   12
      MultiLine       =   -1  'True
      TabIndex        =   35
      Text            =   "FInsPreFacturas.frx":5C12
      Top             =   3360
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.TextBox TxtDescuento2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   330
      Index           =   1
      Left            =   5880
      MaxLength       =   12
      MultiLine       =   -1  'True
      TabIndex        =   23
      Text            =   "FInsPreFacturas.frx":5C19
      Top             =   2100
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.TextBox TxtDescuento2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   330
      Index           =   0
      Left            =   5880
      MaxLength       =   12
      MultiLine       =   -1  'True
      TabIndex        =   11
      Text            =   "FInsPreFacturas.frx":5C20
      Top             =   840
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFC0&
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
      Left            =   7560
      MaskColor       =   &H8000000B&
      Picture         =   "FInsPreFacturas.frx":5C27
      Style           =   1  'Graphical
      TabIndex        =   37
      ToolTipText     =   "Elimina Rubros"
      Top             =   945
      Width           =   750
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFC0&
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
      Left            =   7560
      MaskColor       =   &H8000000B&
      Picture         =   "FInsPreFacturas.frx":5F31
      Style           =   1  'Graphical
      TabIndex        =   38
      ToolTipText     =   "Salir del Modulo"
      Top             =   1785
      Width           =   750
   End
   Begin VB.TextBox TxtCantidad 
      BackColor       =   &H00FFFFC0&
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
      Left            =   1890
      MaxLength       =   2
      TabIndex        =   29
      Text            =   "0"
      Top             =   3360
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.TextBox TxtValor 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   330
      Index           =   2
      Left            =   2730
      MaxLength       =   12
      MultiLine       =   -1  'True
      TabIndex        =   31
      Text            =   "FInsPreFacturas.frx":67FB
      Top             =   3360
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.TextBox TxtDescuento 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   330
      Index           =   2
      Left            =   4305
      MaxLength       =   12
      MultiLine       =   -1  'True
      TabIndex        =   33
      Text            =   "FInsPreFacturas.frx":6802
      Top             =   3360
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.CheckBox CheqProducto 
      BackColor       =   &H00C0C000&
      Caption         =   "PRODUCTO 3:"
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
      Left            =   105
      TabIndex        =   24
      Top             =   2625
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFC0&
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
      Left            =   7560
      MaskColor       =   &H8000000B&
      Picture         =   "FInsPreFacturas.frx":6809
      Style           =   1  'Graphical
      TabIndex        =   36
      ToolTipText     =   "Inserta Rubros a Facturar"
      Top             =   105
      Width           =   750
   End
   Begin VB.TextBox TxtCantidad 
      BackColor       =   &H00FFFFC0&
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
      Left            =   1890
      MaxLength       =   2
      TabIndex        =   17
      Text            =   "0"
      Top             =   2100
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.TextBox TxtValor 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   330
      Index           =   1
      Left            =   2730
      MaxLength       =   12
      MultiLine       =   -1  'True
      TabIndex        =   19
      Text            =   "FInsPreFacturas.frx":70D3
      Top             =   2100
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.TextBox TxtDescuento 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   330
      Index           =   1
      Left            =   4305
      MaxLength       =   12
      MultiLine       =   -1  'True
      TabIndex        =   21
      Text            =   "FInsPreFacturas.frx":70DA
      Top             =   2100
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.CheckBox CheqProducto 
      BackColor       =   &H00FFFF00&
      Caption         =   "PRODUCTO 2:"
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
      Left            =   105
      TabIndex        =   12
      Top             =   1365
      Width           =   1695
   End
   Begin VB.CheckBox CheqProducto 
      BackColor       =   &H00FFFF80&
      Caption         =   "PRODUCTO 1:"
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
      TabIndex        =   0
      Top             =   105
      Width           =   1695
   End
   Begin VB.TextBox TxtDescuento 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   330
      Index           =   0
      Left            =   4305
      MaxLength       =   12
      MultiLine       =   -1  'True
      TabIndex        =   9
      Text            =   "FInsPreFacturas.frx":70E1
      Top             =   840
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.TextBox TxtValor 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   330
      Index           =   0
      Left            =   2730
      MaxLength       =   12
      MultiLine       =   -1  'True
      TabIndex        =   7
      Text            =   "FInsPreFacturas.frx":70E8
      Top             =   840
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.TextBox TxtCantidad 
      BackColor       =   &H00FFFFC0&
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
      Left            =   1890
      MaxLength       =   2
      TabIndex        =   5
      Text            =   "0"
      Top             =   840
      Visible         =   0   'False
      Width           =   750
   End
   Begin MSAdodcLib.Adodc AdoProducto 
      Height          =   330
      Left            =   105
      Top             =   3780
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
      Caption         =   "Prodcuto"
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
   Begin MSDataListLib.DataCombo DCProducto 
      Bindings        =   "FInsPreFacturas.frx":70EF
      DataSource      =   "AdoProducto"
      Height          =   315
      Index           =   1
      Left            =   1785
      TabIndex        =   13
      Top             =   1365
      Visible         =   0   'False
      Width           =   5580
      _ExtentX        =   9843
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16777152
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
   Begin MSDataListLib.DataCombo DCProducto 
      Bindings        =   "FInsPreFacturas.frx":7109
      DataSource      =   "AdoProducto"
      Height          =   315
      Index           =   0
      Left            =   1785
      TabIndex        =   1
      Top             =   105
      Visible         =   0   'False
      Width           =   5580
      _ExtentX        =   9843
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16777152
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
   Begin MSMask.MaskEdBox MBFechaP 
      Height          =   330
      Index           =   0
      Left            =   525
      TabIndex        =   3
      Top             =   840
      Visible         =   0   'False
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   16777152
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
   Begin MSMask.MaskEdBox MBFechaP 
      Height          =   330
      Index           =   1
      Left            =   525
      TabIndex        =   15
      Top             =   2100
      Visible         =   0   'False
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   16777152
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
   Begin MSDataListLib.DataCombo DCProducto 
      Bindings        =   "FInsPreFacturas.frx":7123
      DataSource      =   "AdoProducto"
      Height          =   315
      Index           =   2
      Left            =   1785
      TabIndex        =   25
      Top             =   2625
      Visible         =   0   'False
      Width           =   5580
      _ExtentX        =   9843
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16777152
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
   Begin MSMask.MaskEdBox MBFechaP 
      Height          =   330
      Index           =   2
      Left            =   525
      TabIndex        =   27
      Top             =   3360
      Visible         =   0   'False
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   16777152
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
   Begin VB.Label LblDescuento2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DESCUENTO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Index           =   2
      Left            =   5880
      TabIndex        =   34
      Top             =   3045
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label LblDescuento2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DESCUENTO 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Index           =   1
      Left            =   5880
      TabIndex        =   22
      Top             =   1785
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label LblDescuento2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DESCUENTO 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Index           =   0
      Left            =   5880
      TabIndex        =   10
      Top             =   525
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label LblFecha 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " FECHA INIC."
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
      Left            =   525
      TabIndex        =   26
      Top             =   3045
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label LblCantidad 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CANT."
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
      Left            =   1890
      TabIndex        =   28
      Top             =   3045
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label LblValor 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "VALOR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Index           =   2
      Left            =   2730
      TabIndex        =   30
      Top             =   3045
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label LblDescuento 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DESCUENTO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Index           =   2
      Left            =   4305
      TabIndex        =   32
      Top             =   3045
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label LblFecha 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "FECHA INIC."
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
      Left            =   525
      TabIndex        =   14
      Top             =   1785
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label LblCantidad 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CANT."
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
      Left            =   1890
      TabIndex        =   16
      Top             =   1785
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label LblValor 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "VALOR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Index           =   1
      Left            =   2730
      TabIndex        =   18
      Top             =   1785
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label LblDescuento 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DESCUENTO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Index           =   1
      Left            =   4305
      TabIndex        =   20
      Top             =   1785
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label LblDescuento 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DESCUENTO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Index           =   0
      Left            =   4305
      TabIndex        =   8
      Top             =   525
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label LblValor 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "VALOR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Index           =   0
      Left            =   2730
      TabIndex        =   6
      Top             =   525
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label LblCantidad 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CANT."
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
      Left            =   1890
      TabIndex        =   4
      Top             =   525
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label LblFecha 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "FECHA INIC."
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
      Left            =   525
      TabIndex        =   2
      Top             =   525
      Visible         =   0   'False
      Width           =   1275
   End
End
Attribute VB_Name = "FInsPreFacturas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
  Unload FInsPreFacturas
End Sub

Private Sub Command3_Click()
    If AdoProducto.Recordset.RecordCount > 0 Then
       Titulo = "FORMULARIO DE ELIMINACION"
       Mensajes = "ESTA SEGURO DE ELIMINAR LA PREFACTURACION DE ESTE CLIENTE, " _
                & "NO PODRA REVERSAR ESTE PROCESO."
       If BoxMensaje = vbYes Then
          RatonReloj
          For I = 0 To 2
              If CheqProducto(I).value Then
                 Cantidad = CInt(TxtCantidad(I))
                 If Cantidad > 0 Then
                    CodigoInv = Ninguno
                    Mifecha = PrimerDiaMes(MBFechaP(I))
                    Producto = DCProducto(I).Text
                    AdoProducto.Recordset.MoveFirst
                    AdoProducto.Recordset.Find ("Producto = '" & Producto & "' ")
                    If Not AdoProducto.Recordset.EOF Then CodigoInv = AdoProducto.Recordset.fields("Codigo_Inv")
                    For J = 1 To Cantidad
                        'MsgBox Mifecha
                        NoMes = Month(Mifecha)
                        NoAnio = Year(Mifecha)
                        Anio = CStr(NoAnio)
                        sSQL = "DELETE * " _
                             & "FROM Clientes_Facturacion " _
                             & "WHERE Item = '" & NumEmpresa & "' " _
                             & "AND Codigo = '" & TBeneficiario.Codigo & "' " _
                             & "AND Codigo_Inv = '" & CodigoInv & "' " _
                             & "AND Periodo = '" & Anio & "' " _
                             & "AND Num_Mes = '" & NoMes & "' " _
                             & "AND Item = '" & NumEmpresa & "' "
                        Ejecutar_SQL_SP sSQL
                        Mifecha = PrimerDiaMes(CLongFecha(CFechaLong(Mifecha) + 31))
                    Next J
                 End If
              End If
          Next I
       End If
    End If
    RatonNormal
    MsgBox "Proceso terminado"
End Sub

Private Sub Form_Activate()
    FInsPreFacturas.Caption = TBeneficiario.Cliente
    sSQL = "SELECT Codigo_Inv, Producto " _
         & "FROM Catalogo_Productos " _
         & "WHERE Periodo = '" & Periodo_Contable & "' " _
         & "AND Item = '" & NumEmpresa & "' " _
         & "AND LEN(Cta_Inventario) <= 1 " _
         & "AND LEN(Cta_Ventas) > 1 " _
         & "AND LEN(Cta_Ventas_0) > 1 " _
         & "AND TC = 'P' " _
         & "ORDER BY Producto "
    SelectDB_Combo DCProducto(0), AdoProducto, sSQL, "Producto"
    SelectDB_Combo DCProducto(1), AdoProducto, sSQL, "Producto"
    SelectDB_Combo DCProducto(2), AdoProducto, sSQL, "Producto"
    RatonNormal
End Sub

Private Sub Form_Load()
   'FClientesFlash.Height = TBarCliente.Height + FrmPensiones.Top
    CentrarForm FInsPreFacturas
    Redondear_Formulario FInsPreFacturas, 35
    ConectarAdodc AdoProducto
End Sub

Private Sub Command2_Click()
'0702079658: VACA PRIETO EDY ANGELICA
    If AdoProducto.Recordset.RecordCount > 0 Then
       RatonReloj
       Cadena = ""
      'Eliminamos pensiones si los tuviera
       For I = 0 To 2
           If CheqProducto(I).value Then
              Cantidad = CInt(TxtCantidad(I))
              Valor = CCur(TxtValor(I))
              Total_Desc = CCur(TxtDescuento(I))
              If Cantidad > 0 And Valor > 0 Then
                 CodigoInv = Ninguno
                 Mifecha = PrimerDiaMes(MBFechaP(I))
                 Producto = DCProducto(I).Text
                 AdoProducto.Recordset.MoveFirst
                 AdoProducto.Recordset.Find ("Producto = '" & Producto & "' ")
                 If Not AdoProducto.Recordset.EOF Then CodigoInv = AdoProducto.Recordset.fields("Codigo_Inv")
                 For J = 1 To Cantidad
                     'MsgBox Mifecha
                     NoMes = Month(Mifecha)
                     NoAnio = Year(Mifecha)
                     Anio = CStr(NoAnio)
                     sSQL = "DELETE * " _
                          & "FROM Clientes_Facturacion " _
                          & "WHERE Item = '" & NumEmpresa & "' " _
                          & "AND Codigo = '" & TBeneficiario.Codigo & "' " _
                          & "AND Codigo_Inv = '" & CodigoInv & "' " _
                          & "AND Periodo = '" & Anio & "' " _
                          & "AND Num_Mes = " & NoMes & " "
                     Ejecutar_SQL_SP sSQL
                     Mifecha = PrimerDiaMes(CLongFecha(CFechaLong(Mifecha) + 31))
                 Next J
              End If
           End If
       Next I
      'Procedemos a insertar las pensiones
       For I = 0 To 2
           If CheqProducto(I).value Then
              Cantidad = CInt(TxtCantidad(I))
              Valor = CCur(TxtValor(I))
              Total_Desc = CCur(TxtDescuento(I))
              Total_Desc2 = CCur(TxtDescuento2(I))
              If Cantidad > 0 And Valor > 0 Then
                 CodigoInv = Ninguno
                 Mifecha = PrimerDiaMes(MBFechaP(I))
                 Producto = DCProducto(I).Text
                 AdoProducto.Recordset.MoveFirst
                 AdoProducto.Recordset.Find ("Producto = '" & Producto & "' ")
                 If Not AdoProducto.Recordset.EOF Then CodigoInv = AdoProducto.Recordset.fields("Codigo_Inv")
                 For J = 1 To Cantidad
                     'MsgBox Mifecha
                     NoMes = Month(Mifecha)
                     NoAnio = Year(Mifecha)
                     Anio = CStr(NoAnio)
                    'Cadena = Cadena & CodigoInv & ": " & NoMes & " -> " & MesesLetras(NoMes) & "-" & Anio & "Por = " & Valor & vbCrLf
                     SetAdoAddNew "Clientes_Facturacion"
                     SetAdoFields "T", Normal
                     SetAdoFields "Codigo", TBeneficiario.Codigo
                     SetAdoFields "Codigo_Inv", CodigoInv
                     SetAdoFields "Valor", Valor
                     SetAdoFields "GrupoNo", TBeneficiario.Grupo_No
                     SetAdoFields "Num_Mes", NoMes
                     SetAdoFields "Mes", MesesLetras(NoMes)
                     SetAdoFields "Periodo", Anio
                     SetAdoFields "Fecha", Mifecha
                     SetAdoFields "Descuento", Total_Desc
                     SetAdoFields "Descuento2", Total_Desc2
                     SetAdoUpdate
                     'MsgBox "|| " & CLongFecha(CFechaLong(Mifecha) + 31)
                     Mifecha = PrimerDiaMes(CLongFecha(CFechaLong(Mifecha) + 31))
                     'MsgBox "-> " & Mifecha
                 Next J
              End If
           End If
       Next I
    End If
    RatonNormal
    MsgBox "PROCESO EXITOSO" & vbCrLf & "Vuelva a listar el Cliente y verifique los datos procesados"
    Unload FInsPreFacturas
End Sub

Private Sub CheqProducto_Click(index As Integer)
  If CheqProducto(index).value Then
     DCProducto(index).Visible = True
     LblFecha(index).Visible = True
     MBFechaP(index).Visible = True
     LblCantidad(index).Visible = True
     TxtCantidad(index).Visible = True
     LblValor(index).Visible = True
     TxtValor(index).Visible = True
     LblDescuento(index).Visible = True
     TxtDescuento(index).Visible = True
     LblDescuento2(index).Visible = True
     TxtDescuento2(index).Visible = True
  Else
     DCProducto(index).Visible = False
     LblFecha(index).Visible = False
     MBFechaP(index).Visible = False
     LblCantidad(index).Visible = False
     TxtCantidad(index).Visible = False
     LblValor(index).Visible = False
     TxtValor(index).Visible = False
     LblDescuento(index).Visible = False
     TxtDescuento(index).Visible = False
     LblDescuento2(index).Visible = False
     TxtDescuento2(index).Visible = False
  End If
End Sub

Private Sub MBFechaP_GotFocus(index As Integer)
   MarcarTexto MBFechaP(index)
End Sub

Private Sub MBFechaP_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub MBFechaP_LostFocus(index As Integer)
   FechaValida MBFechaP(index)
End Sub

Private Sub TxtCantidad_GotFocus(index As Integer)
   MarcarTexto TxtCantidad(index)
End Sub

Private Sub TxtCantidad_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub TxtCantidad_LostFocus(index As Integer)
   TextoValido TxtCantidad(index), True, , 0
End Sub

Private Sub TxtValor_GotFocus(index As Integer)
   MarcarTexto TxtValor(index)
End Sub

Private Sub TxtValor_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub TxtValor_LostFocus(index As Integer)
   TextoValido TxtValor(index), True
End Sub

Private Sub TxtDescuento_GotFocus(index As Integer)
   MarcarTexto TxtDescuento(index)
End Sub

Private Sub TxtDescuento_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub TxtDescuento_LostFocus(index As Integer)
   TextoValido TxtDescuento(index), True
End Sub

Private Sub TxtDescuento2_GotFocus(index As Integer)
   MarcarTexto TxtDescuento2(index)
End Sub

Private Sub TxtDescuento2_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub TxtDescuento2_LostFocus(index As Integer)
   TextoValido TxtDescuento2(index), True
End Sub

