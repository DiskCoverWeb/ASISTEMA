VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Begin VB.Form FRecaudacionBancosPreFa 
   BackColor       =   &H00C0FFC0&
   Caption         =   "RECAUDACIONES POR BANCO"
   ClientHeight    =   9450
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12915
   Icon            =   "AutRecau.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9450
   ScaleWidth      =   12915
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&S"
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
      Left            =   21840
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   1050
      Width           =   330
   End
   Begin VB.PictureBox PctBanco 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1275
      Left            =   105
      ScaleHeight     =   1275
      ScaleWidth      =   8730
      TabIndex        =   27
      Top             =   945
      Width           =   8730
   End
   Begin VB.CheckBox CheqAlDia 
      Caption         =   "Generar quienes esten al dia "
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
      Left            =   17010
      TabIndex        =   22
      Top             =   1890
      Width           =   3480
   End
   Begin VB.TextBox TextFacturaNo 
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
      Left            =   15750
      TabIndex        =   19
      Text            =   "99999999"
      Top             =   1890
      Width           =   1170
   End
   Begin VB.CheckBox CheqMatricula 
      BackColor       =   &H00FF8080&
      Caption         =   "Generar &Matricula y Pension"
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
      Left            =   17010
      TabIndex        =   20
      Top             =   1050
      Width           =   3480
   End
   Begin VB.CheckBox Cheq_NumCodigos 
      Caption         =   "Procesar Con Codigo de la Empresa"
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
      Left            =   17010
      TabIndex        =   21
      Top             =   1470
      Value           =   1  'Checked
      Width           =   3480
   End
   Begin MSDataGridLib.DataGrid DGFactura 
      Bindings        =   "AutRecau.frx":0442
      Height          =   2430
      Left            =   105
      TabIndex        =   26
      Top             =   6195
      Width           =   10515
      _ExtentX        =   18547
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
   Begin VB.TextBox TxtFile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3375
      Left            =   105
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   25
      Top             =   2310
      Width           =   10515
   End
   Begin MSAdodcLib.Adodc AdoAux 
      Height          =   330
      Left            =   4095
      Top             =   4725
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
   Begin MSAdodcLib.Adodc AdoDetalle 
      Height          =   330
      Left            =   1995
      Top             =   5040
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
      Caption         =   "Detalle"
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
   Begin MSAdodcLib.Adodc AdoProducto 
      Height          =   330
      Left            =   4095
      Top             =   4410
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
      Caption         =   "Producto"
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
      Left            =   10395
      TabIndex        =   11
      Top             =   1050
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   16761024
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
   Begin MSAdodcLib.Adodc AdoAbono 
      Height          =   330
      Left            =   4095
      Top             =   4095
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
      Caption         =   "Abono"
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
   Begin MSMask.MaskEdBox MBFechaF 
      Height          =   330
      Left            =   10395
      TabIndex        =   13
      Top             =   1470
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   16761024
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
   Begin MSAdodcLib.Adodc AdoPendiente 
      Height          =   330
      Left            =   1995
      Top             =   4410
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
      Caption         =   "Pendiente"
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
      Left            =   1995
      Top             =   4725
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
   Begin MSAdodcLib.Adodc AdoFactura 
      Height          =   330
      Left            =   1575
      Top             =   9450
      Width           =   2955
      _ExtentX        =   5212
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
      Caption         =   "Factura"
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
   Begin MSAdodcLib.Adodc AdoBanco 
      Height          =   330
      Left            =   4095
      Top             =   5040
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
      Caption         =   "Banco"
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
   Begin MSComDlg.CommonDialog CDialogDir 
      Left            =   12285
      Top             =   2940
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin MSAdodcLib.Adodc AdoGrupo 
      Height          =   330
      Left            =   1995
      Top             =   4095
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
      Caption         =   "Grupo"
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
   Begin MSDataListLib.DataCombo DCLinea 
      Bindings        =   "AutRecau.frx":045B
      DataSource      =   "AdoLinea"
      Height          =   315
      Left            =   11760
      TabIndex        =   17
      Top             =   1365
      Width           =   5160
      _ExtentX        =   9102
      _ExtentY        =   556
      _Version        =   393216
      ForeColor       =   255
      Text            =   "CxC Clientes"
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
   Begin MSAdodcLib.Adodc AdoLinea 
      Height          =   330
      Left            =   1995
      Top             =   3780
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
   Begin MSMask.MaskEdBox MBFechaV 
      Height          =   330
      Left            =   10395
      TabIndex        =   15
      Top             =   1890
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   16761024
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
   Begin MSAdodcLib.Adodc AdoClientesMatriculas 
      Height          =   330
      Left            =   1995
      Top             =   5355
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
      Caption         =   "ClientesMatriculas"
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   11655
      Top             =   2940
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AutRecau.frx":0472
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AutRecau.frx":0D4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AutRecau.frx":1626
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AutRecau.frx":1940
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AutRecau.frx":21F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AutRecau.frx":2510
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AutRecau.frx":2DEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AutRecau.frx":3104
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AutRecau.frx":39DE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc AdoTablaSRI 
      Height          =   330
      Left            =   4095
      Top             =   5355
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
      Caption         =   "TablaSRI"
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
      TabIndex        =   0
      Top             =   0
      Width           =   12915
      _ExtentX        =   22781
      _ExtentY        =   1588
      ButtonWidth     =   1720
      ButtonHeight    =   1429
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Salir"
            Key             =   "Salir"
            Object.ToolTipText     =   "Salir del Módulo"
            ImageIndex      =   1
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Boton_Menu"
                  Text            =   "Boton Menu"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Enviar"
            Key             =   "Enviar"
            Object.ToolTipText     =   "Enviar Deuda al Banco"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Recibir"
            Key             =   "SubirAbonos"
            Object.ToolTipText     =   "Recibir Cobros del Banco"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Facturar"
            Key             =   "GenerarFacturas"
            Object.ToolTipText     =   "Generar Facturas"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Alumnos"
            Key             =   "AlumnosContabilidad"
            Object.ToolTipText     =   "Alumnos de Contabilidad"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Renumerar"
            Key             =   "RenumerarCodigos"
            Object.ToolTipText     =   "Renumerar Estudiantes"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Imprimir"
            Key             =   "ImprimirCodigos"
            Object.ToolTipText     =   "Imprimir Codigos"
            ImageIndex      =   8
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSDataListLib.DataCombo DCGrupoF 
         Bindings        =   "AutRecau.frx":42B8
         DataSource      =   "AdoGrupo"
         Height          =   360
         Left            =   12810
         TabIndex        =   5
         Top             =   420
         Visible         =   0   'False
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   635
         _Version        =   393216
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DCGrupoI 
         Bindings        =   "AutRecau.frx":42CF
         DataSource      =   "AdoGrupo"
         Height          =   360
         Left            =   11025
         TabIndex        =   4
         Top             =   420
         Visible         =   0   'False
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   635
         _Version        =   393216
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.CheckBox CheqRangos 
         Caption         =   "Procesar &Por Rangos Grupos:"
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
         Left            =   11025
         TabIndex        =   3
         Top             =   105
         Width           =   3690
      End
      Begin VB.Frame Frame3 
         Caption         =   "ENTIDAD FINANCIERA"
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
         Left            =   6930
         TabIndex        =   1
         Top             =   105
         Width           =   4005
         Begin MSDataListLib.DataCombo DCTablaSRI 
            Bindings        =   "AutRecau.frx":42E6
            DataSource      =   "AdoTablaSRI"
            Height          =   315
            Left            =   105
            TabIndex        =   2
            Top             =   210
            Width           =   3795
            _ExtentX        =   6694
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   4194304
            Text            =   "Banco"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Cod. Banco"
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
         Left            =   14700
         TabIndex        =   6
         Top             =   105
         Width           =   1275
         Begin VB.TextBox TxtCodBanco 
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
            Text            =   "0"
            Top             =   210
            Width           =   1065
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Cuenta de Acreditacion de los Abonos"
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
         Left            =   16065
         TabIndex        =   8
         Top             =   105
         Width           =   5055
         Begin MSDataListLib.DataCombo DCBanco 
            Bindings        =   "AutRecau.frx":4300
            DataSource      =   "AdoBanco"
            Height          =   315
            Left            =   105
            TabIndex        =   9
            Top             =   210
            Width           =   4845
            _ExtentX        =   8546
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   192
            Text            =   "Banco"
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
      End
   End
   Begin VB.Label LabelAbonos 
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
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   6930
      TabIndex        =   24
      Top             =   9450
      Width           =   2115
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Nota de Venta No. 1234567890123-001001"
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
      Left            =   11760
      TabIndex        =   18
      Top             =   1890
      Width           =   4005
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &Vencimiento"
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
      Left            =   8925
      TabIndex        =   14
      Top             =   1890
      Width           =   1485
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " TOTAL RECAUDADO"
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
      TabIndex        =   23
      Top             =   9450
      Width           =   2325
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &LINEA DE CUENTAS POR COBRAR PENSIONES:"
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
      Left            =   11760
      TabIndex        =   16
      Top             =   1050
      Width           =   5160
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Tope de &Pago"
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
      Left            =   8925
      TabIndex        =   12
      Top             =   1470
      Width           =   1485
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &Facturacion"
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
      Left            =   8925
      TabIndex        =   10
      Top             =   1050
      Width           =   1485
   End
End
Attribute VB_Name = "FRecaudacionBancosPreFa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RetVal
Dim TBenef As Tipo_Beneficiarios

Dim AdoStrCnn1 As String
Dim XAdoStrCnn As String
Dim AdoStrCnnOld As String

Dim NumFile As Integer
Dim NumFileAbonos As Integer
Dim NumFileDetalle As Integer
Dim NumFileAlumnos As Integer
Dim NumFileFacturas As Integer
Dim NumFileFacturas2 As Integer
Dim NumFileProducto As Integer

Dim RutaGeneraFile As String
Dim RutaGeneraFile2 As String
Dim RutaGeneraFileAbonos As String
Dim RutaGeneraFileDetalle As String
Dim RutaGeneraFileAlumnos As String
Dim RutaGeneraFileFacturas As String
Dim RutaGeneraFileProducto As String

Dim IJ As Long
Dim ModuloResp As String
Dim RutaBackupXX As String
Dim Cta_Gasto_Banco As String

Dim Costo_Banco As Double
Dim Total_Costo_Banco As Double

Dim Tipo_Carga As Byte

Dim AdoCliDB As ADODB.Recordset

Public Sub Imprimir_Codigos_Banco(Datas As Adodc, _
                                  Optional EsCampoCorto As Boolean)
Dim TipoLetra As String
On Error GoTo Errorhandler

Mensajes = "Seguro de Imprimir en:" & vbCrLf & Printer.DeviceName & "?"
Titulo = "IMPRESION"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
RatonReloj
InicioX = 1.5: InicioY = 0.1
TipoLetra = Tipoarialnew
Escala_Centimetro Orientacion_Pagina, TipoLetra, 10
Pagina = 1
PosLinea = InicioY
'Iniciamos la impresion
Printer.FontBold = False
Printer.FontName = TipoLetra
With Datas.Recordset
If .RecordCount > 0 Then
    .MoveFirst
     'MsgBox .RecordCount
     Do While Not .EOF
        PrinterPaint LogoTipo, InicioX - 1, PosLinea, 3, 1.5
        Printer.FontBold = True
        If UCaseStrg(Empresa) <> UCaseStrg(NombreComercial) Then
           Printer.FontSize = 13
           PrinterTexto CentrarTexto(UCaseStrg(Empresa)), PosLinea, UCaseStrg(Empresa) & " - " & NombreCiudad
           PosLinea = PosLinea + 0.7
           Printer.FontSize = 12
           PrinterTexto CentrarTexto(NombreComercial), PosLinea, NombreComercial
           PosLinea = PosLinea + 0.6
           Printer.FontSize = 8
           Printer.FontItalic = True
           Printer.FontBold = False
           Cadena = NombreCiudad & ", " & Direccion & ". Teléfono: " & Telefono1 & "."
           PrinterTexto CentrarTexto(Cadena), PosLinea, Cadena
           Printer.FontItalic = False
           PosLinea = PosLinea + 0.5
        Else
           Printer.FontSize = 14
           PrinterTexto CentrarTexto(UCaseStrg(Empresa)), PosLinea, UCaseStrg(Empresa)
           PosLinea = PosLinea + 0.6
           Printer.FontSize = 8
           Printer.FontItalic = True
           Printer.FontBold = False
           Cadena = NombreCiudad & ", " & Direccion & ". Teléfono: " & Telefono1 & "."
           PrinterTexto CentrarTexto(Cadena), PosLinea, Cadena
           Printer.FontItalic = False
           PosLinea = PosLinea + 1
        End If
        Printer.FontName = TipoLetra
        Printer.FontSize = 10
        Cadena = "Señores Padres de Familia de la " & Empresa & " - " & ULCase(NombreCiudad) & "."
        Cadena = TrimStrg(Cadena)
        NumeroLineas = PrinterLineasMayor(InicioX, PosLinea, Cadena, 18)
        PosLinea = PosLinea_Aux
        PosLinea = PosLinea + 0.7
        CuentaBanco = SinEspaciosIzq(DCBanco)
        CuentaBanco = TrimStrg(MidStrg(DCBanco, Len(CuentaBanco) + 1, Len(DCBanco)))
        
        Cadena = "Reciban un cordial saludo de quienes conformamos la Unidad Educativa, " _
               & "por medio de la presente le informamos a ustedes que el sistema de pagos " _
               & "de pensiones es en el " & NombreBanco & ", con el código siguiente:"
        Cadena = TrimStrg(Cadena)
        PosLinea = Printer_Texto_Justifica(InicioX, 19.5, PosLinea, Cadena) - 0.2
        Printer.FontBold = True
        PrinterTexto InicioX, PosLinea, "ESTUDIANTE:"
        PrinterTexto InicioX + 14, PosLinea - 0.2, "CODIGO DEL BANCO:"
        Printer.FontBold = False
        PrinterTexto InicioX + 3, PosLinea, .fields("Cliente")
        
        PosLinea = PosLinea + 0.4
        Printer.FontBold = True
        PrinterTexto InicioX, PosLinea, "CURSO/GRADO:"
        Printer.FontSize = 16
        PrinterTexto InicioX + 14, PosLinea - 0.2, "No."
        Printer.FontSize = 10
        Printer.FontBold = False
        PrinterTexto InicioX + 3, PosLinea, .fields("Direccion")
        Printer.FontSize = 16
        PrinterTexto InicioX + 15, PosLinea - 0.2, .fields("CI_RUC")
        Printer.FontSize = 10
        PosLinea = PosLinea + 0.5
        PrinterTexto InicioX, PosLinea, "Les recordamos que la fecha de pago son los diéz primeros días laborables de cada mes."
        '"Además, pasado los diez días del plazo de pago, el sistema realizará automáticamente un recargo el siguiente mes."
        PosLinea = PosLinea + 0.4
        PrinterTexto InicioX, PosLinea, "Agradecemos su comprensión y atención a la presente."
        PosLinea = PosLinea + 0.6
        PrinterTexto InicioX, PosLinea, "Atentamente,"
        PosLinea = PosLinea + 0.4
        PrinterTexto InicioX, PosLinea, NombreGerente
        PosLinea = PosLinea + 0.7
        Imprimir_Linea_H PosLinea, InicioX, 20
        PosLinea = PosLinea + 0.2
        'MsgBox PosLinea
        If PosLinea >= (LimiteAlto - 3) Then
           Printer.NewPage
           PosLinea = InicioY
        End If
       .MoveNext
     Loop
End If
End With
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

Private Sub CheqRangos_Click()
 If CheqRangos.value = 0 Then
    DCGrupoI.Visible = False
    DCGrupoF.Visible = False
 Else
    DCGrupoI.Visible = True
    DCGrupoF.Visible = True
 End If
End Sub

'Sube los abonos
Private Sub Subir_Abonos()
Dim DiaV As Integer
Dim MesV As Integer
Dim AñoV As Integer
Dim Fecha_Tope As String
Dim Tipo_Carga As Byte
Dim IdNumMes As Boolean
Dim ContTAB As Integer
Dim Total_Alumnos As Long
Dim Estab As Boolean
Dim AuxNumEmp As String
Dim Separador As String
Dim CaptionTemp As String
Dim AbonosPar As String
Dim Total_Letras As String
Dim CodigoEncontrado As String
Dim NumeroTarjeta As String
Dim CamposFile() As Campos_Tabla
Dim CostoTarjeta As Currency

   Progreso_Barra.Mensaje_Box = "SUBIENDO ABONOS DEL BANCO " & TextoBanco
   Progreso_Iniciar
   DGFactura.Visible = False
   FA.Cod_CxC = DCLinea.Text
   Lineas_De_CxC FA
   FA.Nuevo_Doc = True
   TextoImprimio = ""
   NumeroTarjeta = ""
   CodigoEncontrado = "'.'"
   CostoTarjeta = 0
  'FA.Factura = Numero_Factura(FA)
   If FA.Nuevo_Doc Then FA.Factura = ReadSetDataNum(FA.TC & "_SERIE_" & FA.Serie, True, False)
   Factura_No = FA.Factura
   
   sSQL = "DELETE * " _
        & "FROM Asiento_F " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND CodigoU = '" & CodigoUsuario & "' "
   Ejecutar_SQL_SP sSQL
   
   sSQL = "DELETE * " _
        & "FROM Tabla_Temporal " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND Modulo = '" & NumModulo & "' " _
        & "AND CodigoU = '" & CodigoUsuario & "' "
   Ejecutar_SQL_SP sSQL
     
   sSQL = "UPDATE Clientes_Facturacion " _
        & "SET X = '.' " _
        & "WHERE Item = '" & NumEmpresa & "' "
   Ejecutar_SQL_SP sSQL
      
  Tipo_Carga = Leer_Campo_Empresa("Tipo_Carga_Banco")
 
  CaptionTemp = FRecaudacionBancosPreFa.Caption
  TextoImprimio = ""
  Separador = ","
  
  FechaValida MBFechaI
  Mifecha = BuscarFecha(MBFechaI)
  FechaTexto = MBFechaI ' FechaSistema
  DiarioCaja = ReadSetDataNum("Recibo_No", True, True)
  
  CDialogDir.Filename = RutaSysBases & "\BANCO\ABONOS\*.*"
  CDialogDir.InitDir = RutaSysBases & "\BANCO\ABONOS\"
  CDialogDir.Flags = cdlOFNFileMustExist + cdlOFNNoChangeDir + cdlOFNHideReadOnly
  CDialogDir.Filter = "Archivos PDF|*.txt"
  CDialogDir.DialogTitle = "Abrir Archivo"
  CDialogDir.Action = 1
  NombreArchivo = CDialogDir.Filename
  If NombreArchivo <> "" Then
     Actualizar_Datos_Representantes_SP
'''     sSQL = "SELECT * " _
'''          & "FROM Clientes " _
'''          & "WHERE Codigo <> '.' " _
'''          & "ORDER BY CI_RUC "
'''     Select_Adodc AdoClientes, sSQL
  RutaGeneraFile = UCaseStrg(NombreArchivo)
  TotalIngreso = 0
  Contador = 0
  FileResp = 0
 'Establecemos los campos del archivo plano del Banco
  NumFile = FreeFile
  Total_Alumnos = 0
  FechaTexto = FechaSistema
  TxtFile = ""
  
 'Determinamos cuantos registro vamos a actualizar y
 'cuantos campos tiene el archivo del Banco
 '--------------------------------------------------
  Cadena = ""
  NumFile = FreeFile
  Open RutaGeneraFile For Input As #NumFile
  Do While Not EOF(NumFile)
     Line Input #NumFile, LineFile
     Cadena = Cadena & LineFile & vbCrLf
  Loop
  Close #NumFile
 
  NumFile = FreeFile
  Open RutaGeneraFile For Output As #NumFile
  Print #NumFile, MidStrg(Cadena, 1, Len(Cadena) - 2)
  Close #NumFile
  Cadena = ""
 '---------------------------------------------------
  ContTAB = 0
  Open RutaGeneraFile For Input As #NumFile
    Do While Not EOF(NumFile)
       Line Input #NumFile, Cod_Field
       Cod_Field = Replace(Cod_Field, Chr(34), "")
       If Contador = 0 Then
          For I = 1 To Len(Cod_Field)
              If MidStrg(Cod_Field, I, 1) = vbTab Then
                 ContTAB = ContTAB + 1
                 Separador = vbTab
              End If
          Next I
       End If
       Contador = Contador + 1
    Loop
  Close #NumFile
  
  'MsgBox ContTAB & vbCrLf & RutaGeneraFile
  'GoTo SalirSubida
  
  ReDim CamposFile(ContTAB) As Campos_Tabla
  FechaTexto = FechaSistema
  For I = 0 To ContTAB
      CamposFile(I).Campo = "C" & Format$(I, "00")
  Next I
 'MsgBox "CantCampos = " & CantCampos
 'Empieza la subida
  Progreso_Barra.Valor_Maximo = Progreso_Barra.Valor_Maximo + (Contador * 2)
  Contador = 0
  Open RutaGeneraFile For Input As #NumFile
    Do While Not EOF(NumFile)
      'Leemos la trama donde esta toda la informacion del archivo de recaudaciones
       Line Input #NumFile, Cod_Field
       TxtFile = TxtFile & Cod_Field & vbCrLf
      'Comenzamos la subida de los Abonos
       For I = 0 To ContTAB
           CamposFile(I).Valor = ""
       Next I
       Cadena = Cod_Field
       I = 0: No_Desde = 1: No_Hasta = 0
       TotalReg = Len(Cadena)
       Do While Len(Cadena) > 0
          No_Hasta = No_Hasta + 1
          'MsgBox "[" & MidStrg(Cadena, No_Hasta, 1) & "] ID: " & No_Hasta & "/" & TotalReg
          If MidStrg(Cadena, No_Hasta, 1) = Separador Then
            'MsgBox "|" & MidStrg(Cadena, No_Desde, No_Hasta - 1) & "| " & No_Hasta & "/" & TotalReg
             CamposFile(I).Valor = MidStrg(Cadena, No_Desde, No_Hasta - 1)
             Cadena = MidStrg(Cadena, No_Hasta + 1, Len(Cadena))
             No_Desde = 1: No_Hasta = 0
             I = I + 1
          End If
          If No_Hasta >= TotalReg Then
            'MsgBox "<>|" & Cadena & "|"
             CamposFile(I).Valor = Cadena
             Cadena = ""
          End If
       Loop
      'Actualizamos de que alumnos vamos a ingresar el abono
       NoMeses = 0
       NoAnio = 0
       Mes = Ninguno
       CodigoInv = Ninguno
       CodigoCli = Ninguno
       Codigo1 = Ninguno
       Codigo3 = Ninguno  'Tipo de pago
       Codigo4 = Ninguno
       CodigoP = "0"
       Select Case TextoBanco
         Case "PICHINCHA"
'             CodigoP = CStr(Val(CamposFile(7).Valor))
'             FechaTexto = CamposFile(12).Valor
             'MsgBox Tipo_Carga
              If Tipo_Carga >= 1 Then
                 CodigoCli = TrimStrg(CStr(Val(MidStrg(Cod_Field, 25, 19))))
                 FechaTexto = MidStrg(Cod_Field, 205, 2) & "/" & MidStrg(Cod_Field, 207, 2) & "/" & MidStrg(Cod_Field, 209, 4)
                 Total = Val(MidStrg(Cod_Field, 48, 13)) / 100
                 NoMeses = Val(MidStrg(Cod_Field, 198, 2))
                 NoAnio = Val(MidStrg(Cod_Field, 201, 4))
                 If Val(NoMeses) <= 0 Then
                    NoMeses = Val(MidStrg(Cod_Field, 207, 2))
                    NoAnio = Val(MidStrg(Cod_Field, 209, 4))
                 End If
                 Mes = MesesLetras(NoMeses)
                 NombreCliente = TrimStrg(MidStrg(Cod_Field, 125, 40))
                 CodigoInv = SinEspaciosIzq(TrimStrg(MidStrg(Cod_Field, 165, 33)))
                 Producto = TrimStrg(MidStrg(Cod_Field, 165, 33))
                 Producto = TrimStrg(MidStrg(Producto, Len(CodigoInv) + 1, Len(Producto)))
                 Cantidad = 1
                 Codigo3 = MidStrg(Cod_Field, 74, 3)
                 If Codigo3 = "EFE" Then
                    NombreBanco = "EFECTIVO POR VENTANILLA"
                 Else
                    Codigo3 = MidStrg(Cod_Field, 74, 3) & ". " & MidStrg(Cod_Field, 81, 3) & ". "
                    Codigo4 = MidStrg(Cod_Field, 92, 12)
                    NombreBanco = Codigo3 & "No. " & MidStrg(Cod_Field, 92, 12)
                 End If
              Else
                 If ContTAB > 40 Then
                    FechaTexto = Replace(CamposFile(25).Valor, " ", "/")
                    If IsDate(FechaTexto) Then
                       CodigoCli = Trim(CamposFile(4).Valor)
                       NombreCliente = CamposFile(5).Valor
                       Codigo3 = CamposFile(16).Valor 'Tipo de pago
                       Codigo4 = CamposFile(26).Valor
                       SubTotal = CamposFile(8).Valor
                       Total = CamposFile(8).Valor
                       
                       Total_Letras = CamposFile(7).Valor
                       CodigoInv = SinEspaciosIzq(Total_Letras)
                       If IsNumeric(CodigoInv) Then
                          NoMeses = Val(MidStrg(Total_Letras, Len(Total_Letras) - 6, 2))
                          NoAnio = SinEspaciosDer(Total_Letras)
                          Mes = MesesLetras(NoMeses)
                          Producto = Trim(MidStrg(Total_Letras, Len(CodigoInv) + 1, 28))
                       Else
                          CodigoInv = Ninguno
                          NoAnio = Year(FechaTexto)
                          If IsNumeric(MidStrg(NombreCliente, 1, 2)) Then
                             NoAnio = 0
                             NoMeses = Val(MidStrg(NombreCliente, 1, 2))
                             Mes = MesesLetras(NoMeses)
                             NombreCliente = MidStrg(NombreCliente, 3, Len(NombreCliente))
                          Else
                             NoMeses = Month(CamposFile(25).Valor)
                             Mes = MesesLetras(NoMeses)
                             NombreCliente = MidStrg(CamposFile(5).Valor, Len(Mes), Len(CamposFile(5).Valor) - 1)
                          End If
                       End If
                       Cantidad = 1
                      'MsgBox "(" & ContTAB & ")" & Producto & vbCrLf & NoAnio & vbCrLf & NoMeses
                    End If
                 Else
                     FechaTexto = CStr(CamposFile(12).Valor)
                     NoAnio = Year(FechaTexto)
                     Codigo3 = CamposFile(16).Valor
                     Codigo4 = CamposFile(17).Valor
                     CodigoCli = Trim(CamposFile(7).Valor)
                     CodigoB = CamposFile(7).Valor
                     
                     If IsNumeric(MidStrg(CamposFile(8).Valor, 1, 2)) Then
                        NoAnio = 0
                        NoMeses = Val(MidStrg(CamposFile(8).Valor, 1, 2))
                        Mes = MesesLetras(NoMeses)
                        NombreCliente = MidStrg(CamposFile(8).Valor, 3, Len(CamposFile(8).Valor))
                     Else
                        NoMeses = Val(CamposFile(14).Valor)
                        Mes = MesesLetras(NoMeses)
                        NombreCliente = MidStrg(CamposFile(8).Valor, Len(Mes), Len(CamposFile(8).Valor) - 1)
                     End If
                     
                    'MsgBox CamposFile(14).Valor
                     
                     Producto = "PENSION DE: " & Mes
                     FechaTexto = CamposFile(12).Valor
                     
                     SubTotal = CamposFile(9).Valor / 100
                     Total = CamposFile(9).Valor / 100
                     If Len(Codigo4) <= 1 Then Codigo4 = Ninguno
                     Cantidad = 1
                    'MsgBox Producto & vbCrLf & NoAnio & vbCrLf & NoMeses
                 End If
              End If
              'MsgBox NombreCliente & vbCrLf & CodigoP & vbCrLf & Total
         Case "INTERNACIONAL", "OTROSBANCOS"
              If Separador = vbTab Then
                 Select Case ContTAB
                  'Otros Bancos y Debitos
                   Case 18, 21
                        If IsNumeric(MidStrg(CamposFile(16).Valor, 1, 4)) Then
                           CodigoB = CamposFile(6).Valor
                           CodigoCli = CamposFile(6).Valor
                           FechaTexto = MidStrg(CamposFile(7).Valor, 1, 10)
                           HoraTexto = TrimStrg(MidStrg(CamposFile(7).Valor, 11, 5))
                           CodigoP = CamposFile(10).Valor
                           Codigo3 = CamposFile(10).Valor
                           Codigo4 = CamposFile(8).Valor
                           ListaFacturas = Ninguno
                           NoAnio = Val(MidStrg(CamposFile(16).Valor, 1, 4))
                           NoMeses = Val(MidStrg(CamposFile(16).Valor, 6, 2))
                           CodigoInv = TrimStrg(MidStrg(CamposFile(16).Valor, 9, 18))
                           Mes = MesesLetras(NoMeses)
                           SubTotal = Val(CamposFile(13).Valor)
                           Total = SubTotal
                           If Len(Codigo4) <= 1 Then Codigo4 = Ninguno
                        End If
                  'Recaudaciones
                   Case Else
                        'MsgBox CamposFile(12).Valor
                        If IsNumeric(MidStrg(CamposFile(12).Valor, 1, 4)) Then
                           CodigoB = CamposFile(3).Valor
                           CodigoCli = CamposFile(3).Valor
                           FechaTexto = MidStrg(CamposFile(16).Valor, 1, 10)
                           HoraTexto = TrimStrg(MidStrg(CamposFile(16).Valor, 11, 5))
                           CodigoP = CamposFile(18).Valor
                           Codigo3 = CamposFile(18).Valor
                           Codigo4 = TrimStrg(MidStrg(MidStrg(CamposFile(7).Valor, 1, 4) & ". " & Replace(CamposFile(20).Valor, "AGENCIA", "AGE."), 1, 30))
                           ListaFacturas = Ninguno
                           NoAnio = Val(MidStrg(CamposFile(12).Valor, 1, 4))
                           NoMeses = Val(MidStrg(CamposFile(12).Valor, 6, 2))
                           CodigoInv = TrimStrg(MidStrg(CamposFile(12).Valor, 9, 18))
                           Mes = MesesLetras(NoMeses)
                           SubTotal = Val(CamposFile(10).Valor)
                           Total = SubTotal
                           If Len(Codigo4) <= 1 Then Codigo4 = Ninguno
                        End If
                 End Select
                 If CodigoCli = Ninguno Then
                    For I = 0 To ContTAB
                        CamposFile(I).Campo = CamposFile(I).Valor
                    Next I
                    FechaTexto = FechaSistema
                    NoAnio = "2000"
                    NoMeses = 12
                    Mes = MesesLetras(NoMeses)
                    CodigoP = Ninguno
                 End If
                 Cantidad = 1
              Else
                 Cod_Field = Replace(Cod_Field, vbTab, " ")
                 CodigoCli = TrimStrg(CStr(Val(MidStrg(Cod_Field, 25, 19))))
                 Mifecha = MidStrg(Cod_Field, 205, 2) & "/" & MidStrg(Cod_Field, 207, 2) & "/" & MidStrg(Cod_Field, 209, 4)
                 Total = Val(MidStrg(Cod_Field, 48, 13)) / 100
                 Cadena = TrimStrg(MidStrg(Cod_Field, 165, 40))
                 I = InStr(Cadena, "-")
                 If I > 1 Then
                    Cadena = TrimStrg(MidStrg(Cadena, I + 1, Len(Cadena)))
                    J = InStr(Cadena, "-")
                    If J > 1 Then Cadena = TrimStrg(MidStrg(Cadena, J + 1, Len(Cadena)))
                    If I > 1 And J > 1 Then Producto = TrimStrg(MidStrg(Producto, I + 1, J - 1)) & ":"
                    NoMeses = LetrasMeses(Cadena)
                    Mes = MesesLetras(NoMeses)
                    NoAnio = Year(Mifecha)
                 Else
                    Producto = "PENSION:"
                    NoMeses = Val(Cadena)
                    Mes = MesesLetras(NoMeses)
                    NoAnio = Year(Mifecha)
                 End If
                 FechaTexto = Mifecha
              End If
         Case "PRODUBANCO"
              CodigoP = CamposFile(7).Valor
              CodigoCli = CamposFile(7).Valor
              NombreCliente = CamposFile(8).Valor
              SubTotal = Val(CamposFile(9).Valor)
              Total = Val(CamposFile(9).Valor)
              FechaTexto = CamposFile(12).Valor
              Producto = CamposFile(14).Valor
              CodigoB = CamposFile(14).Valor
             'MsgBox CodigoB
             
              NoAnio = Val(MidStrg(TrimStrg(SinEspaciosDer(CodigoB)), 1, 4))
              CodigoB = LCase(TrimStrg(MidStrg(TrimStrg(SinEspaciosDer(CodigoB)), 6, 3)))
              NoMeses = LetrasMeses(CodigoB)
              Codigo3 = CamposFile(27).Valor
              Codigo4 = CamposFile(31).Valor
              Codigo2 = CamposFile(31).Valor
              If NoAnio <= "1900" And IsDate(FechaTexto) Then
                 NoMeses = Month(FechaTexto)
                 NoAnio = CStr(Year(FechaTexto))
                 Mes = MesesLetras(NoMeses)
              End If
         Case "BOLIVARIANO"
             'MsgBox "1234567890123456789012345678901234567890" & vbCrLf & Cod_Field
              If MidStrg(Cod_Field, 12, 2) = "00" Then
                 CodigoCli = MidStrg(Cod_Field, 14, 8)
                 CodigoP = MidStrg(Cod_Field, 14, 8)
              Else
                 CodigoCli = MidStrg(Cod_Field, 12, 10)
                 CodigoP = MidStrg(Cod_Field, 12, 10)
              End If
              If Total_Alumnos = 0 Then
                 FechaTexto = MidStrg(Cod_Field, 12, 2) & "/" _
                            & MidStrg(Cod_Field, 10, 2) & "/" _
                            & MidStrg(Cod_Field, 6, 4)
                 CodigoP = Ninguno
              End If
              Producto = "PENSION DE: " & MesesLetras(Val(MidStrg(Cod_Field, 5, 2)))
              Mes = MesesLetras(Val(MidStrg(Cod_Field, 5, 2)))
              NoMeses = MidStrg(Cod_Field, 5, 2)
              NoAnio = MidStrg(Cod_Field, 1, 4)
              Cantidad = 1
              Total = Val(MidStrg(Cod_Field, 22, 11)) / 100
              'CodigoP = MidStrg(Cod_Field, 1, 8)
              'Total = Val(MidStrg(Cod_Field, 9, 11)) / 100
              SubTotal = Total
         Case "GUAYAQUIL"
             'MsgBox "1234567890123456789012345678901234567890" & vbCrLf & Cod_Field
              If MidStrg(Cod_Field, 3, 3) = "RER" Then
                  CodigoCli = Ninguno
                  CodigoP = Ninguno
                  Producto = Ninguno
                  Mes = Ninguno
                  NoMeses = 0
                  NoAnio = 0
                  Cantidad = 1
                  CodigoP = Ninguno
                  Total = 0
                  SubTotal = Total
              Else
                  Mes = Ninguno
                  NoMeses = 0
                  NoAnio = "2000"
                  CodigoCli = MidStrg(Cod_Field, 5, 8)
                  CodigoP = MidStrg(Cod_Field, 5, 8)
                  FechaTexto = MidStrg(Cod_Field, 91, 2) & "/" _
                             & MidStrg(Cod_Field, 89, 2) & "/" _
                             & MidStrg(Cod_Field, 85, 4)
                  Producto = "PENSION DE: " & MidStrg(Cod_Field, 60, 3)
                  Mes = MidStrg(Cod_Field, 60, 3)
                  NoMeses = LetrasMeses(MidStrg(Cod_Field, 60, 3))
                  NoAnio = MidStrg(Cod_Field, 85, 4)
                  Cantidad = 1
                  Total = Val(MidStrg(Cod_Field, 75, 10)) / 100
                  SubTotal = Total
              End If
         Case "PACIFICO"
               CodigoCli = TrimStrg(MidStrg(Cod_Field, 302, 14))
               CodigoP = TrimStrg(Val(MidStrg(Cod_Field, 302, 14)))
               Producto = TrimStrg(MidStrg(Cod_Field, 252, 40))
               Grupo_No = TrimStrg(SinEspaciosIzq(Producto))
               Producto = TrimStrg(MidStrg(Producto, Len(Grupo_No) + 1, Len(Producto)))
               Mes = TrimStrg(SinEspaciosIzq(Producto))
               NoMeses = LetrasMeses(Mes)
               'MsgBox TrimStrg(SinEspaciosDer(Producto))
               If IsNumeric(TrimStrg(SinEspaciosDer(Producto))) Then
                  NoAnio = Val(TrimStrg(SinEspaciosDer(Producto)))
                  If NoAnio < 1900 Then NoAnio = Val(TrimStrg(MidStrg(Producto, 4, 5)))
                  
               Else
                  NoAnio = Year(FechaSistema)
               End If
               If NoAnio < 1900 Then NoAnio = Year(FechaSistema)
               
               Cantidad = 1
               SubTotal = Val(MidStrg(Cod_Field, 32, 15)) / 100
               Total = Val(MidStrg(Cod_Field, 32, 15)) / 100
               NombreCliente = TrimStrg(MidStrg(Cod_Field, 316, 40))
               FechaTexto = MidStrg(Cod_Field, 27, 2) & "/" & MidStrg(Cod_Field, 25, 2) & "/" & MidStrg(Cod_Field, 21, 4)
               SubCta = TrimStrg(MidStrg(Cod_Field, 292, 10))
         Case "POREXCEL"
               CodigoCli = CamposFile(0).Valor
               FechaTexto = CStr(CamposFile(1).Valor)
               NoMeses = Val(CamposFile(2).Valor)
               NoAnio = Val(CamposFile(3).Valor)
               SubTotal = Val(CamposFile(4).Valor)
               Total = SubTotal
              'MsgBox CamposFile(14).Valor
               Mes = MesesLetras(NoMeses)
               Producto = "PENSION DE: " & Mes
               Cantidad = 1
               'If Not IsNumeric(NoMeses) Then MsgBox CodigoP
         Case "TARJETAS"
              Cantidad = 1
              CodigoCli = CamposFile(3).Valor
              NombreCliente = Ninguno
              NumeroTarjeta = CamposFile(4).Valor
              CostoTarjeta = Val(CamposFile(10).Valor)
              SubTotal = Val(CamposFile(15).Valor)
              Total = Val(CamposFile(15).Valor)
              FechaTexto = Format(MidStrg(CamposFile(0).Valor, 1, 10), FormatoFechas)
              CodigoB = FechaTexto
                                   
              CodigoP = Leer_RUC_CI_TARJETA(NumeroTarjeta, Total, CodigoCli)
              CodigoCli = CodigoP
              
              Producto = "Tarjeta: " & NumeroTarjeta & " - " & CamposFile(5).Valor
             ' MsgBox CodigoB & vbCrLf & CodigoP & vbCrLf & NombreCliente
              
              CodigoB = LCase(TrimStrg(MidStrg(TrimStrg(SinEspaciosDer(CodigoB)), 6, 3)))
               'LetrasMeses(CodigoB)
              
              Codigo3 = CamposFile(1).Valor
              Codigo4 = CamposFile(2).Valor
              Codigo2 = CamposFile(3).Valor
'''              MsgBox CodigoCli & vbCrLf & CodigoP & vbCrLf & NombreCliente & vbCrLf & FechaTexto & vbCrLf & _
'''                     Total & vbCrLf & Saldo & vbCrLf & NoAnio & vbCrLf & NoMeses
         Case ".."
               CodigoCli = CamposFile(0).Valor
               FechaTexto = CStr(CamposFile(1).Valor)
               NoMeses = Val(CamposFile(2).Valor)
               NoAnio = Val(CamposFile(3).Valor)
               SubTotal = Val(CamposFile(4).Valor)
               Total = SubTotal
              'MsgBox CamposFile(14).Valor
               Mes = MesesLetras(NoMeses)
               Producto = "PENSION DE: " & Mes
               Cantidad = 1
        End Select
       Si_No = True
       'MsgBox "-> " & NombreCliente & vbCrLf & Len(CodigoCli) & vbCrLf & CodigoCli
       
       If Len(CodigoCli) > 5 Then
          'With AdoClientes.Recordset
           'If AdoClientes.Recordset.RecordCount > 0 Then
               Do While Len(CodigoCli) <= 10 And Si_No
                  sSQL = "SELECT Codigo,Cliente,Grupo, CI_RUC " _
                       & "FROM Clientes " _
                       & "WHERE CI_RUC = '" & CodigoCli & "' "
                  Select_AdoDB AdoCliDB, sSQL, "Clientes"
                  If AdoCliDB.RecordCount > 0 Then
                     CodigoCli = AdoCliDB.fields("Codigo")
                     NombreCliente = AdoCliDB.fields("Cliente")
                     Grupo_No = AdoCliDB.fields("Grupo")
                     Si_No = False
                  Else
                     CodigoCli = "0" & CodigoCli
                  End If
                  AdoCliDB.Close
                  
'''                 AdoClientes.Recordset.MoveFirst
'''                 AdoClientes.Recordset.Find ("CI_RUC = '" & CodigoCli & "' ")
'''                  MsgBox "Ok"
'''                  If Not AdoClientes.Recordset.EOF Then
'''                     CodigoCli = AdoClientes.Recordset.Fields("Codigo")
'''                     NombreCliente = AdoClientes.Recordset.Fields("Cliente")
'''                     Grupo_No = AdoClientes.Recordset.Fields("Grupo")
'''                     Si_No = False
'''                     MsgBox "........"
'''                  Else
'''                     CodigoCli = "0" & CodigoCli
'''                  End If
                     

               Loop
           'Else
           '    CodigoCli = Ninguno
           'End If
          'End With
          Progreso_Barra.Mensaje_Box = NombreCliente & ": " & FechaTexto & " -> " & NoAnio & " -> " & NoMeses
          Progreso_Esperar
       End If
      ' MsgBox NombreCliente & vbCrLf & Len(CodigoCli) & vbCrLf & CodigoCli
       
       If Len(CodigoCli) > 10 Then CodigoCli = Ninguno
       
       Producto = "PENSION DE: "
       If CodigoInv <> Ninguno Then
          sSQL = "SELECT Producto " _
               & "FROM Catalogo_Productos " _
               & "WHERE Item = '" & NumEmpresa & "' " _
               & "AND Periodo = '" & Periodo_Contable & "' " _
               & "AND Codigo_Inv = '" & CodigoInv & "' " _
               & "AND TC = 'P' "
          Select_Adodc AdoAux, sSQL
          If AdoAux.Recordset.RecordCount > 0 Then Producto = AdoAux.Recordset.fields("Producto")
       End If

      'MsgBox CodigoP & vbCrLf & CodigoCli & vbCrLf & NombreCliente & vbCrLf & FechaTexto & vbCrLf & Total & vbCrLf & NoAnio & vbCrLf & NoMeses & vbCrLf & Producto
         'Determinamos las pensiones pendientes
''       Fecha_Tope = "01/" & CStr(NoMeses) & "/" & NoAnio
''       Fecha_Tope = UltimoDiaMes(Fecha_Tope)
       Fecha_Tope = CLongFecha(CFechaLong(FechaSistema) + 300)
       Fecha_Tope = UltimoDiaMes(Fecha_Tope)
       
'''       Cadena = "TRAMA DEL TEXTO:" & vbCrLf
'''       For i = 0 To ContTAB
'''           Cadena = Cadena & Format(i, "00") & " = " & CamposFile(i).Valor & vbCrLf
'''       Next i
'''       Cadena = Cadena _
'''              & String(50, "_") & vbCrLf & vbCrLf _
'''              & "Codigo(C): " & CodigoB & vbCrLf _
'''              & "Codigo: " & CodigoP & vbCrLf _
'''              & "Mes: " & Mes & vbCrLf _
'''              & "No Mes: " & NoMeses & vbCrLf _
'''              & "Año: " & NoAnio & vbCrLf _
'''              & "Producto: " & Producto & vbCrLf _
'''              & "Grupo: " & Grupo_No & vbCrLf _
'''              & "Total: " & Total & vbCrLf _
'''              & "Cliente: " & NombreCliente & vbCrLf _
'''              & "Codigo Int: " & CodigoCli & vbCrLf _
'''              & "Fecha: " & FechaTexto & vbCrLf _
'''              & "Fecha Tope: " & Fecha_Tope & vbCrLf _
'''              & "Tipo Banco: " & TextoBanco
    '''   MsgBox Cadena
      'Verificamos la primera deuda antgua que tenga el Cliente de ese mes
       'MsgBox FechaTexto
       If IsDate(FechaTexto) Then
          sSQL = "SELECT * " _
               & "FROM Clientes_Facturacion " _
               & "WHERE Item = '" & NumEmpresa & "' " _
               & "AND Codigo = '" & CodigoCli & "' "
          If NoMeses > 0 Then sSQL = sSQL & "AND Num_Mes = " & NoMeses & " "
          If NoAnio > 0 Then sSQL = sSQL & "AND Periodo = '" & CStr(NoAnio) & "' "
          If CodigoInv <> Ninguno Then sSQL = sSQL & "AND Codigo_Inv = '" & CodigoInv & "' "
          If Len(Dato_DBF.Carpeta) > 1 Then
             sSQL = sSQL _
                  & "AND Fecha <= #" & BuscarFecha(Fecha_Tope) & "# " _
                  & "AND Fecha >= #" & BuscarFecha(Dato_DBF.FechaI) & "# "
          Else
             sSQL = sSQL _
                  & "AND Fecha <= #" & BuscarFecha(Fecha_Tope) & "# "
          End If
          sSQL = sSQL & "ORDER BY Periodo, Num_Mes "
          Select_Adodc AdoAux, sSQL
          'If CodigoCli = "003001155" Then
          'MsgBox Dato_DBF.Carpeta & vbCrLf & "Tipo de Carga: " & Tipo_Carga & vbCrLf & sSQL & vbCrLf & "Cantidad Regs: " & AdoAux.Recordset.RecordCount
          'MsgBox AdoAux.Recordset.RecordCount
          If AdoAux.Recordset.RecordCount > 0 And (CodigoCli <> Ninguno) Then
             Sumatoria = Total
             Do While Not AdoAux.Recordset.EOF
                NoAnio = AdoAux.Recordset.fields("Periodo")
                NoMeses = AdoAux.Recordset.fields("Num_Mes")
                Mes = MesesLetras(NoMeses)
                CodigoInv = AdoAux.Recordset.fields("Codigo_Inv")
                Total = AdoAux.Recordset.fields("Valor") - (AdoAux.Recordset.fields("Descuento") + AdoAux.Recordset.fields("Descuento2"))
                Progreso_Barra.Mensaje_Box = "Determinando valores a facturar"
                Progreso_Esperar True
               'MsgBox "Codigo: " & CodigoCli & ", De: " & NombreCliente & " Por = " & Format(Total, "#,##0.00") & vbCrLf
                If Total <= Sumatoria Then
                   SetAdoAddNew "Asiento_F"
                   SetAdoFields "CODIGO", CodigoInv
                   SetAdoFields "CANT", Cantidad
                   SetAdoFields "PRODUCTO", Producto
                   SetAdoFields "PRECIO", Total
                   SetAdoFields "TOTAL", Total
                   SetAdoFields "Mes", Mes
                   SetAdoFields "A_No", NoMeses
                   SetAdoFields "HABIT", NoAnio
                   SetAdoFields "FECHA", FechaTexto
                   SetAdoFields "RUTA", NombreCliente
                   SetAdoFields "Codigo_Cliente", CodigoCli
                   SetAdoFields "Cod_Ejec", CodigoP
                   SetAdoFields "Cta", Grupo_No
                   SetAdoFields "Numero", Factura_No
                   SetAdoFields "Serie", FA.Serie
                   SetAdoFields "Autorizacion", FA.Autorizacion
                   SetAdoFields "CODIGO_L", Codigo3
                   SetAdoFields "TICKET", Codigo4
                   If CostoTarjeta > 0 Then SetAdoFields "COSTO", CostoTarjeta
                   SetAdoUpdate
                   Factura_No = Factura_No + 1
                   Sumatoria = Sumatoria - Total
                   AdoAux.Recordset.fields("X") = "X"
                   AdoAux.Recordset.Update
                End If
                AdoAux.Recordset.MoveNext
             Loop
          Else
             If Total_Alumnos <> 0 Then
                Total_Letras = Format$(Total, "#,##0.00")
                Total_Letras = String(12 - Len(Total_Letras), " ") & Total_Letras
                Cadena = "Codigo: " & CodigoCli & " " & vbTab & "Cedula: " & CodigoP & " " & vbTab & FechaTexto & " " & vbTab & Mes & "/" & NoAnio & ", USD " _
                       & Total_Letras & " " & vbTab & NombreCliente
                If Len(NumeroTarjeta) > 1 Then Cadena = Cadena & " " & vbTab & NumeroTarjeta
                'MsgBox Cadena
                Insertar_Texto_Temporal_SP Cadena
                TextoImprimio = TextoImprimio & Cadena & vbCrLf
             End If
          End If
        'MsgBox CodigoCli
         Total_Alumnos = Total_Alumnos + 1
       End If
    Loop
  Close #NumFile
  End If
  
  Factura_No = Val(TextFacturaNo)
  Numero = 0
  DGFactura.Visible = False
  sSQL = "SELECT * " _
       & "FROM Asiento_F " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "ORDER BY FECHA,Codigo_Cliente,Numero "
  Select_Adodc AdoFactura, sSQL
  With AdoFactura.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          Progreso_Barra.Mensaje_Box = "Verificando totales"
          Progreso_Esperar True
          sSQL = "SELECT Codigo,Num_Mes,Periodo,SUM(Valor) AS TValor " _
               & "FROM Clientes_Facturacion " _
               & "WHERE Item = '" & NumEmpresa & "' " _
               & "AND Codigo = '" & .fields("Codigo_Cliente") & "' " _
               & "AND Codigo_Inv = '" & .fields("CODIGO") & "' " _
               & "AND Num_Mes = " & .fields("A_No") & " " _
               & "AND Periodo <= '" & .fields("HABIT") & "' " _
               & "GROUP BY Codigo,Num_Mes,Periodo "
          Select_Adodc AdoAux, sSQL
          If AdoAux.Recordset.RecordCount <= 0 Then .fields("CodBod") = "X"
         'MsgBox AdoAux.Recordset.RecordCount
         .Update
         .MoveNext
       Loop
   End If
  End With
  
  sSQL = "DELETE * " _
       & "FROM Asiento_F " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodBod = 'X' "
  Ejecutar_SQL_SP sSQL
  
  sSQL = "SELECT * " _
       & "FROM Asiento_F " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "ORDER BY RUTA,Codigo_Cliente,FECHA,Numero "
  Select_Adodc AdoFactura, sSQL
  With AdoFactura.Recordset
   If .RecordCount > 0 Then
       Mifecha = .fields("FECHA")
       Numero = Factura_No
       CodigoC = .fields("Codigo_Cliente")
       Do While Not .EOF
          Progreso_Barra.Mensaje_Box = "Actualizando Facturas a Procesar: " & Mifecha & " - " & Numero & " - " & CodigoC
          Progreso_Esperar

          If CodigoC <> .fields("Codigo_Cliente") Or Mifecha <> .fields("FECHA") Then
             CodigoC = .fields("Codigo_Cliente")
             Mifecha = .fields("FECHA")
             Factura_No = Factura_No + 1
          End If
         .fields("Numero") = Factura_No
         .Update
         .MoveNext
       Loop
   End If
  End With
  DGFactura.Visible = True
  DGFactura.Caption = "PREFACTURACION DEL DIA"
  sSQL = "SELECT FECHA,Numero As FACTURA,RUTA As BENEFICIARIO,CANT,CODIGO,PRODUCTO,TOTAL,COSTO As COMISION," _
       & "HABIT As PERIODO,Mes,A_No As NO_MES,Codigo_Cliente,Serie,Autorizacion,Cta As Trans_No," _
       & "CODIGO_L As Forma_P,TICKET As No_Cta,Cod_Ejec As Cod_Banco,CodigoU,Item " _
       & "FROM Asiento_F " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "ORDER BY Numero "
  SQLDec = "PRECIO " & CStr(Dec_PVP) & "|TOTAL 2|."
  Select_Adodc_Grid DGFactura, AdoFactura, sSQL, SQLDec
  TotalIngreso = 0
  With AdoFactura.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          TotalIngreso = TotalIngreso + .fields("TOTAL")
         .MoveNext
       Loop
   End If
  End With
  LabelAbonos.Caption = Format$(TotalIngreso, "#,##0.00")
  RatonNormal
  Progreso_Final
  FRecaudacionBancosPreFa.Caption = CaptionTemp
  'MsgBox "(" & TextoImprimio & ")"
  If Len(TextoImprimio) > 2 Then
     FRecaudacionBancosPreFa.WindowState = vbMaximized
     FInfoError.Show
  End If
SalirSubida:
End Sub

Private Sub Renumerar_Codigos()
Dim CaptionTemp As String
  CaptionTemp = FRecaudacionBancosPreFa.Caption
  Contador = 0
  sSQL = "SELECT Codigo,CI_RUC,Grupo,TD " _
       & "FROM Clientes " _
       & "WHERE TD NOT IN ('C','R') " _
       & "AND FA = " & Val(adFalse) & " " _
       & "AND Cliente <> 'CONSUMIDOR FINAL' " _
       & "ORDER BY Grupo,Cliente,CI_RUC "
  Select_Adodc AdoClientes, sSQL
  With AdoClientes.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          FRecaudacionBancosPreFa.Caption = .fields("TD") & "-" & .fields("Grupo") & " - " & Format$(Contador / .RecordCount, "00%")
          Contador = Contador + 1
         .fields("CI_RUC") = Format$(Contador, "00000000")
         .Update
         .MoveNext
       Loop
   End If
  End With
  Contador = 0
  sSQL = "SELECT Codigo,CI_RUC,Grupo,TD " _
       & "FROM Clientes " _
       & "WHERE FA <> " & Val(adFalse) & " " _
       & "AND Cliente <> 'CONSUMIDOR FINAL' " _
       & "ORDER BY Grupo,Cliente,CI_RUC "
  Select_Adodc AdoClientes, sSQL
  With AdoClientes.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          FRecaudacionBancosPreFa.Caption = .fields("Grupo") & " - " & Format$(Contador / .RecordCount, "00%")
          Contador = Contador + 1
         .fields("CI_RUC") = Format$(Contador, "00000000")
         .Update
         .MoveNext
       Loop
   End If
  End With
  Contador = 0
  sSQL = "SELECT Codigo,CI_RUC,Direccion,Cliente,Grupo,TD " _
       & "FROM Clientes " _
       & "WHERE FA <> " & Val(adFalse) & " "
  If CheqRangos.value <> 0 Then sSQL = sSQL & "AND Grupo BETWEEN '" & DCGrupoI & "' and '" & DCGrupoF & "' "
  sSQL = sSQL & "AND Cliente <> 'CONSUMIDOR FINAL' " _
       & "ORDER BY Grupo,Cliente,Direccion,CI_RUC "
  Select_Adodc AdoClientes, sSQL
  With AdoClientes.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       Do While Not .EOF
          FRecaudacionBancosPreFa.Caption = .fields("Grupo") & " - " & Format$(Contador / .RecordCount, "00%")
          Contador = Contador + 1
          If Cheq_NumCodigos.value <> 0 Then
            .fields("CI_RUC") = NumEmpresa & Format$(Contador, "00000")
          Else
            .fields("CI_RUC") = Format$(Contador, "00000000")
          End If
         .MoveNext
       Loop
      .UpdateBatch
   End If
  End With
  FRecaudacionBancosPreFa.Caption = CaptionTemp
  RatonNormal
  MsgBox "PROCESO DE RENUMERACION TERMINADO"
End Sub

Private Sub Enviar_Cobros()
Dim Cont As Integer
Dim CaptionTemp As String
  DGFactura.Visible = False
  TxtFile.Height = MDI_Y_Max - TxtFile.Top - 100
  Actualizar_Datos_Representantes_SP
  CaptionTemp = FRecaudacionBancosPreFa.Caption
  FechaValida MBFechaI
  FechaValida MBFechaF
  FechaValida MBFechaV
  Cta_Banco = TrimStrg(UCaseStrg(SinEspaciosDer(DCBanco)))
  If Len(Cta_Banco) <= 1 Then Cta_Banco = "0000000000"
  TipoDoc = ""
  AuxNumEmp = NumEmpresa
  FechaFinal = MBFechaF
  FechaTexto = FechaSistema
  FechaTexto1 = Format$(MBFechaI, "MM/dd/yyyy")
  Mifecha = BuscarFecha(MBFechaI)
  MiMes = Format$(Month(MBFechaI), "00")
  Contador = 0
  
 'Consultamos la deuda Pendiente
 sSQL = "UPDATE Clientes_Facturacion " _
      & "SET AlDia = 1 " _
      & "WHERE Item = '" & NumEmpresa & "' "
 Ejecutar_SQL_SP sSQL
 
 sSQL = "UPDATE Clientes_Facturacion " _
      & "SET AlDia = 0 " _
      & "WHERE Item = '" & NumEmpresa & "' " _
      & "AND Fecha < #" & BuscarFecha(MBFechaI) & "# "
 Ejecutar_SQL_SP sSQL
 
  If SQL_Server Then
     sSQL = "UPDATE Clientes_Facturacion " _
          & "SET Fecha = CONVERT(datetime, '01/' & STR(Num_Mes) & '/' & Periodo, 103) " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Fecha <> CONVERT(datetime, '01/' & STR(Num_Mes) & '/' & Periodo, 103) "
     Ejecutar_SQL_SP sSQL
  Else
     sSQL = "UPDATE Clientes_Facturacion " _
          & "SET Fecha = #Num_Mes/01/Periodo# " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Fecha <> #Num_Mes/01/Periodo# "
     'Ejecutar_SQL_SP sSQL
  End If
  
'''  sSQL = "SELECT Codigo, CI, Representante, Cedula_R, Grupo_No, Telefono_R, Lugar_Trabajo_R, TD, Email_R, CI_R, " _
'''       & "Telefono_RS, Tipo_Cta, Cta_Numero, Cod_Banco, Descuento, Caducidad " _
'''       & "FROM Clientes_Matriculas " _
'''       & "WHERE Item = '" & NumEmpresa & "' " _
'''       & "AND Periodo = '" & Periodo_Contable & "' "
'''  If CheqRangos.value <> 0 Then sSQL = sSQL & "AND Grupo_No BETWEEN '" & DCGrupoI & "' and '" & DCGrupoF & "' "
'''  sSQL = sSQL & "ORDER BY Codigo "
'''  Select_Adodc AdoClientesMatriculas, sSQL
  'MsgBox Tipo_Carga
  Select Case Tipo_Carga
    Case 2
         sSQL = "SELECT C.Codigo As CodigoC,F.Codigo,C.Grupo,C.Cliente,C.CI_RUC,C.Direccion,C.Casilla,C.Actividad,F.Periodo,F.Num_Mes,F.Fecha,CC.Codigo_Inv," _
              & "C.Representante,C.CI_RUC_R,C.TD_R,C.Telefono_R,C.Tipo_Cta,C.Cod_Banco,C.Cta_Numero,C.DireccionT,C.Fecha_Cad,C.Email2," _
              & "C.Saldo_Pendiente,CC.Producto,(F.Valor-(F.Descuento+F.Descuento2)) As Valor_Cobro " _
              & "FROM Clientes_Facturacion As F,Clientes As C,Catalogo_Productos As CC " _
              & "WHERE F.Item = '" & NumEmpresa & "' " _
              & "AND CC.Periodo = '" & Periodo_Contable & "' " _
              & "AND (F.Valor-(F.Descuento+F.Descuento2)) > 0 "
    Case 3
         sSQL = "SELECT C.Codigo As CodigoC,F.Codigo,C.Grupo,C.Cliente,C.CI_RUC,C.Direccion,C.Casilla,C.Actividad,C.Plan_Afiliado,C.Saldo_Pendiente," _
              & "C.Representante,C.CI_RUC_R,C.TD_R,C.Telefono_R,C.Tipo_Cta,C.Cod_Banco,C.Cta_Numero,C.DireccionT,C.Fecha_Cad,C.Email2,'PENSION ' As Producto," _
              & "'01.01' As Codigo_Inv,SUM(F.Valor-(F.Descuento+F.Descuento2)) As Valor_Cobro " _
              & "FROM Clientes_Facturacion As F,Clientes As C " _
              & "WHERE F.Item = '" & NumEmpresa & "' "
    Case Else
         sSQL = "SELECT C.Codigo As CodigoC,F.Codigo,C.Grupo,C.Cliente,C.CI_RUC,C.Direccion,C.Casilla,C.Actividad,C.Plan_Afiliado,F.Periodo,F.Num_Mes,F.Fecha," _
              & "C.Representante,C.CI_RUC_R,C.TD_R,C.Telefono_R,C.Tipo_Cta,C.Cod_Banco,C.Cta_Numero,C.DireccionT,C.Fecha_Cad,C.Email2,C.EmailR," _
              & "F.Codigo_Inv,C.Saldo_Pendiente,SUM(F.Valor-(F.Descuento+F.Descuento2)) As Valor_Cobro,"
         If CheqMatricula.value = 1 Then sSQL = sSQL & "'MATRICULAS Y PENSION DE ' As Producto " Else sSQL = sSQL & "'PENSION DE ' As Producto "
         sSQL = sSQL _
              & "FROM Clientes_Facturacion As F, Clientes As C " _
              & "WHERE F.Item = '" & NumEmpresa & "' "
              
  End Select
  sSQL = sSQL & "AND F.Fecha BETWEEN #" & BuscarFecha(MBFechaI) & "# and #" & BuscarFecha(MBFechaF) & "# "
  If CheqRangos.value <> 0 Then sSQL = sSQL & "AND C.Grupo BETWEEN '" & DCGrupoI & "' and '" & DCGrupoF & "' "
  If CheqAlDia.value <> 0 Then sSQL = sSQL & "AND F.AlDia <> 0 "
  Select Case TextoBanco
    Case "BIZBANCKPACIFICO"
    Case "PICHINCHA"
    Case "PACIFICO"
    Case "BOLIVARIANO"
    Case "GUAYAQUIL"
    Case "PRODUBANCO"
    Case "TARJETAS"
         sSQL = sSQL _
              & "AND C.Tipo_Cta = 'TARJETA' " _
              & "AND LEN(C.Cta_Numero) >= 14 "
    Case "INTERNACIONAL"
    Case Else
         sSQL = sSQL _
              & "AND C.Cod_Banco > 0 " _
              & "AND LEN(C.Cta_Numero) > 1 " _
              & "AND C.Tipo_Cta <> '.' "
  End Select
   
  Select Case Tipo_Carga
    Case 2
         sSQL = sSQL _
              & "AND F.Codigo = C.Codigo " _
              & "AND F.Codigo_Inv = CC.Codigo_Inv " _
              & "AND F.Item = CC.Item " _
              & "ORDER BY C.Grupo,C.Cliente,F.Fecha,F.Codigo_Inv "
    Case 3
         sSQL = sSQL _
              & "AND F.Codigo = C.Codigo " _
              & "GROUP BY C.Codigo,F.Codigo,C.Grupo,C.Cliente,C.CI_RUC,C.Direccion,C.Casilla,C.Actividad,C.Plan_Afiliado,C.Saldo_Pendiente," _
              & "C.Representante,C.CI_RUC_R,C.TD_R,C.Telefono_R,C.Tipo_Cta,C.Cod_Banco,C.Cta_Numero,C.DireccionT,C.Fecha_Cad,C.Email2 " _
              & "HAVING SUM(F.Valor-(F.Descuento+F.Descuento2)) > 0 " _
              & "ORDER BY C.Grupo,C.Cliente "
    Case Else
         sSQL = sSQL _
              & "AND F.Codigo = C.Codigo " _
              & "GROUP BY C.Codigo,F.Codigo,C.Grupo,C.Cliente,C.CI_RUC,C.Direccion,C.Casilla,C.Actividad,C.Plan_Afiliado,F.Periodo,F.Codigo_Inv,F.Num_Mes,F.Fecha," _
              & "C.Representante,C.CI_RUC_R,C.TD_R,C.Telefono_R,C.Tipo_Cta,C.Cod_Banco,C.Cta_Numero,C.DireccionT,C.Fecha_Cad,C.Email2,C.EmailR,C.Saldo_Pendiente " _
              & "HAVING SUM(F.Valor-(F.Descuento+F.Descuento2)) > 0 " _
              & "ORDER BY C.Grupo,C.Cliente,F.Fecha "
  End Select
  Select_Adodc AdoFactura, sSQL  ', , , "Envio_Banco_GYE"
  
'''  sSQL = "SELECT F.CodigoC,C.Actividad,C.Cliente,CI_RUC,C.Direccion,C.Grupo,SUM(Saldo_MN) As Saldo_Pend " _
'''       & "FROM Facturas As F, Clientes As C " _
'''       & "WHERE F.Item = '" & NumEmpresa & "' " _
'''       & "AND F.Periodo = '" & Periodo_Contable & "' " _
'''       & "AND F.Fecha < #" & BuscarFecha(MBFechaI) & "# " _
'''       & "AND F.T = 'P' " _
'''       & "AND NOT F.TC IN ('C','P') " _
'''       & "AND F.CodigoC = C.Codigo " _
'''       & "GROUP BY F.CodigoC,C.Actividad,C.Cliente,CI_RUC,C.Direccion,C.Grupo " _
'''       & "HAVING SUM(Saldo_MN) > 0 " _
'''       & "ORDER BY F.CodigoC,C.Actividad,C.Cliente,CI_RUC,C.Direccion,C.Grupo "
'''  Select_Adodc AdoPendiente, sSQL
  
 'Ubicamos donde se grabara el archivo
  'MsgBox Tipo_Carga & vbCrLf & TextoBanco
  
  RutaDestino = RutaSysBases & "\BANCO\FACTURAS\DEUDASESTUDIANTES.csv"
  '''''Generar_Archivo_CSV RutaDestino
  Codigo = UCaseStrg(MesesLetras(Month(MBFechaI)) & "-" & Format$(Day(MBFechaI), "00") & "-" & Year(MBFechaI))
  RutaDestino = RutaSysBases & "\BANCO\FACTURAS\" & Codigo & ".txt"
  NombreArchivoZip = RutaDestino
  FechaFin = BuscarFecha(MBFechaF)
  TextoImprimio = ""
  TxtFile = vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbTab & vbTab & vbTab & "ESPERE UN MOMENTO GENERANDO LOS ARCHIVOS NECESARIOS PARA EL BANCO..."
  Select Case TextoBanco
    Case "PACIFICO"
         FechaFin = BuscarFecha(UltimoDiaMes(MBFechaF))
         Generar_Pacifico RutaDestino
    Case "PICHINCHA"
         FechaFin = BuscarFecha(UltimoDiaMes(MBFechaF))
         Generar_Pichincha
    Case "BOLIVARIANO"
         FechaFin = BuscarFecha(UltimoDiaMes(MBFechaF))
         Generar_Bolivariano RutaDestino
    Case "GUAYAQUIL"
         FechaFin = BuscarFecha(UltimoDiaMes(MBFechaF))
         Generar_Guayaquil
    Case "PRODUBANCO"
         FechaFin = BuscarFecha(UltimoDiaMes(MBFechaF))
         Generar_Produbanco
    Case "TARJETAS"
         FechaFin = BuscarFecha(UltimoDiaMes(MBFechaF))
         Generar_Tarjetas
    Case "INTERNACIONAL"
         FechaFin = BuscarFecha(UltimoDiaMes(MBFechaF))
         Generar_Internacional
    Case Else
         FechaFin = BuscarFecha(UltimoDiaMes(MBFechaF))
         Generar_Otros_Bancos
  End Select
  TxtFile = ""
  FRecaudacionBancosPreFa.Caption = CaptionTemp
End Sub

Private Sub Imprimir_Codigos()
    sSQL = "SELECT Codigo,CI_RUC,Direccion,Cliente,Grupo,TD " _
         & "FROM Clientes " _
         & "WHERE FA <> " & Val(adFalse) & " "
    If CheqRangos.value <> 0 Then sSQL = sSQL & "AND Grupo BETWEEN '" & DCGrupoI & "' and '" & DCGrupoF & "' "
    sSQL = sSQL & "AND Cliente <> 'CONSUMIDOR FINAL' " _
         & "ORDER BY Grupo,Cliente,Direccion,CI_RUC "
    Select_Adodc AdoClientes, sSQL
    Imprimir_Codigos_Banco AdoClientes
End Sub

'Genera las facturas recibidas
Private Sub Generar_Facturas()
Dim ContAbonos As Byte
Dim TextoCarga As String
Dim Fecha_Tope As String
Dim Tipo_Pago As String
Dim TipoDeAbono As String
Dim Cod_Ok As Boolean

   Tipo_Pago = "01"
   Total_Alumnos = 0
   TotalIngreso = 0
   Cta_Banco = TrimStrg(UCaseStrg(SinEspaciosIzq(DCBanco)))
   If Len(Cta_Banco) <= 1 Then Cta_Banco = Cta_Del_Banco
   Mensajes = "Esta Seguro que desea grabar: " & vbCrLf _
            & "La Facturas Preprocesadas " & TextFacturaNo
   Titulo = "Formulario de Grabacion"
   If BoxMensaje = vbYes Then
      RatonReloj
      DGFactura.Visible = False
      With AdoBanco.Recordset
       If .RecordCount > 0 Then
          .MoveFirst
          .Find ("Codigo = '" & Cta_Banco & "' ")
           If Not .EOF Then Tipo_Pago = .fields("Tipo_Pago")
       End If
      End With
       
      Progreso_Iniciar
     'DGFactura.Visible = False
      FA.Cod_CxC = DCLinea.Text
      FA.Tipo_PRN = "FM"
      Lineas_De_CxC FA
      Contador = 0
      
      sSQL = "SELECT FECHA, Numero, Codigo_Cliente, SUM(TOTAL) As Total_Abonos " _
           & "FROM Asiento_F " _
           & "WHERE Item = '" & NumEmpresa & "' " _
           & "AND CodigoU = '" & CodigoUsuario & "' " _
           & "GROUP BY FECHA, Numero, Codigo_Cliente " _
           & "ORDER BY FECHA, Numero, Codigo_Cliente "
      Select_Adodc AdoFactura, sSQL
      With AdoFactura.Recordset
       If .RecordCount > 0 Then
           Total_Alumnos = .RecordCount
           FA.Imp_Mes = True
           Factura_Desde = ReadSetDataNum(FA.TC & "_SERIE_" & FA.Serie, True, False)
           Factura_Hasta = Factura_Desde
           Do While Not .EOF
              Progreso_Barra.Mensaje_Box = "Generando Factura"
              Progreso_Esperar
              Contador = Contador + 1
              Total_Sin_IVA = 0
              Total_Con_IVA = 0
              Total_Desc = 0
              Total_IVA = 0
              Total_Servicio = 0
             'Forma de pago
              Codigo3 = Ninguno
              Codigo4 = Ninguno
              TotalAbonos = .fields("Total_Abonos")
              FechaTexto = .fields("FECHA")
              
              FA.Factura = ReadSetDataNum(FA.TC & "_SERIE_" & FA.Serie, True, True)
              Factura_No = FA.Factura  ' .Fields("Numero")
              Factura_Hasta = FA.Factura
             'MsgBox FA.Factura
              
              CodigoCli = .fields("Codigo_Cliente")
              TBenef = Leer_Datos_Clientes(CodigoCli)
              'MsgBox TBenef.Representante & vbCrLf & TBenef.RUC_CI_Rep & vbCrLf & TBenef.TD_Rep & vbCrLf & TBenef.Cliente
              
              NombreCliente = TBenef.Cliente
              Validar_Porc_IVA FechaTexto
              FA.Porc_IVA = Porc_IVA
             'Fecha_Tope = .Fields("FECHA")
                     
              sSQL = "UPDATE Clientes_Facturacion " _
                   & "SET D = " & Val(adFalse) & " " _
                   & "WHERE Codigo = '" & CodigoCli & "' " _
                   & "AND Item = '" & NumEmpresa & "' "
              Ejecutar_SQL_SP sSQL
              
'''              sSQL = "SELECT Codigo, Representante, Cedula_R, TD, Telefono_R, Lugar_Trabajo_R, Email_R " _
'''                   & "FROM Clientes_Matriculas " _
'''                   & "WHERE Item = '" & NumEmpresa & "' " _
'''                   & "AND Periodo = '" & Periodo_Contable & "' " _
'''                   & "AND Codigo = '" & CodigoCli & "' " _
'''                   & "ORDER BY Codigo "
'''              Select_Adodc AdoAux, sSQL
              
              sSQL = "SELECT CF.*,AF.Codigo_Cliente,AF.CODIGO_L,AF.TICKET,AF.PRODUCTO,AF.HABIT,AF.A_No  " _
                   & "FROM Clientes_Facturacion As CF, Asiento_F As AF " _
                   & "WHERE CF.Codigo = '" & CodigoCli & "' " _
                   & "AND CF.Item = '" & NumEmpresa & "' " _
                   & "AND CF.Codigo_Inv = AF.CODIGO " _
                   & "AND CF.Num_Mes = AF.A_No " _
                   & "AND CF.Periodo = AF.HABIT " _
                   & "AND CF.Item = AF.Item " _
                   & "AND CF.Codigo = AF.Codigo_Cliente " _
                   & "ORDER BY CF.Fecha, CF.Codigo_Inv, Periodo, Num_Mes, CF.ID, CF.Valor DESC "
             'Generar_Consulta_SQL "Facturacion Bancos " & CodigoCli, sSQL
              Select_Adodc AdoProducto, sSQL
             'MsgBox sSQL
             'Detalle de la Factura
              If AdoProducto.Recordset.RecordCount > 0 Then
                 Codigo2 = AdoProducto.Recordset.fields("Periodo")
                 SubCta = AdoProducto.Recordset.fields("TICKET")
                 Codigo3 = AdoProducto.Recordset.fields("CODIGO_L")
                 Codigo4 = AdoProducto.Recordset.fields("PRODUCTO")
                 Do While Not AdoProducto.Recordset.EOF
                    CodigoInv = AdoProducto.Recordset.fields("Codigo_Inv")
                    Cod_Ok = Leer_Codigo_Inv(CodigoInv, FechaSistema)
                    Abono = AdoProducto.Recordset.fields("Valor")
                    Total_Desc_ME = AdoProducto.Recordset.fields("Descuento") + AdoProducto.Recordset.fields("Descuento2")
                    Producto = DatInv.Producto
                    If TotalAbonos >= (Abono - Total_Desc_ME) Then
                       SetAdoAddNew "Detalle_Factura"
                       SetAdoFields "T", Cancelado
                       SetAdoFields "TC", FA.TC
                       SetAdoFields "Porc_IVA", FA.Porc_IVA
                       SetAdoFields "CodigoL", FA.Cod_CxC
                       SetAdoFields "CodigoC", CodigoCli
                       SetAdoFields "Factura", Factura_No
                       SetAdoFields "Cantidad", 1
                       SetAdoFields "Fecha", FechaTexto
                       SetAdoFields "Codigo", CodigoInv
                       SetAdoFields "Producto", Producto
                       SetAdoFields "Precio", Abono
                       SetAdoFields "Total_Desc", Total_Desc_ME
                       SetAdoFields "Total", Abono
                       SetAdoFields "CodigoU", CodigoUsuario
                       SetAdoFields "Mes", AdoProducto.Recordset.fields("Mes")
                       SetAdoFields "Mes_No", AdoProducto.Recordset.fields("Num_Mes")
                       SetAdoFields "Ticket", AdoProducto.Recordset.fields("Periodo")
                       'SetAdoFields "Ruta", "(" & AdoProducto.Recordset.Fields("GrupoNo") & ") " & NombreCliente
                       SetAdoFields "Ruta", AdoProducto.Recordset.fields("GrupoNo")
                       SetAdoFields "Periodo", Periodo_Contable
                       SetAdoFields "Item", NumEmpresa
                       SubTotal_IVA = 0
                       If FA.TC = "NV" Then
                          Total_Sin_IVA = Total_Sin_IVA + Abono
                          SetAdoFields "Cta_Venta", DatInv.Cta_Ventas_0
                       Else
                          If DatInv.IVA Then
                             SubTotal_IVA = Redondear(Abono * Porc_IVA, 2)
                             Total_Con_IVA = Total_Con_IVA + Abono
                          Else
                             Total_Sin_IVA = Total_Sin_IVA + Abono
                          End If
                          SetAdoFields "Cta_Venta", DatInv.Cta_Ventas
                       End If
                       Total_IVA = Total_IVA + SubTotal_IVA
                       Total_Desc = Total_Desc + Total_Desc_ME
                       SetAdoFields "Total_IVA", SubTotal_IVA
                       SetAdoFields "Autorizacion", FA.Autorizacion
                       SetAdoFields "Serie", FA.Serie
                       SetAdoUpdate
                      'MsgBox Factura_No & vbCrLf & CodigoCli & vbCrLf & AdoProducto.Recordset.Fields("Num_Mes") & vbCrLf & AdoProducto.Recordset.Fields("Periodo") & vbCrLf & AdoProducto.Recordset.Fields("Mes")
                      'Activamos el item a borrar
                       sSQL = "UPDATE Clientes_Facturacion " _
                            & "SET D = " & Val(adTrue) & " " _
                            & "WHERE Codigo = '" & CodigoCli & "' " _
                            & "AND Codigo_Inv = '" & AdoProducto.Recordset.fields("Codigo_Inv") & "' " _
                            & "AND Num_Mes = " & AdoProducto.Recordset.fields("Num_Mes") & " " _
                            & "AND Periodo = '" & AdoProducto.Recordset.fields("Periodo") & "' " _
                            & "AND Item = '" & NumEmpresa & "' "
                       Ejecutar_SQL_SP sSQL
                        
                       TotalAbonos = TotalAbonos - (Abono - Total_Desc_ME)
                    End If
                    AdoProducto.Recordset.MoveNext
                 Loop
                 Total_Factura = Total_Sin_IVA + Total_Con_IVA + Total_IVA - Total_Desc
                 TotalIngreso = TotalIngreso + Total_Factura
                'Grabamos las Facturas
                 SetAdoAddNew "Facturas"
                 SetAdoFields "T", Cancelado
                 SetAdoFields "TC", FA.TC
                 SetAdoFields "Porc_IVA", FA.Porc_IVA
                 SetAdoFields "Factura", Factura_No
                 SetAdoFields "Fecha", FechaTexto
                 SetAdoFields "Fecha_C", FechaTexto
                 SetAdoFields "Fecha_V", FechaTexto
                 SetAdoFields "CodigoC", CodigoCli
                 SetAdoFields "Forma_Pago", "DEPOSITO BANCO"
                 SetAdoFields "Sin_IVA", Total_Sin_IVA
                 SetAdoFields "Con_IVA", Total_Con_IVA
                 SetAdoFields "SubTotal", Total_Sin_IVA + Total_Con_IVA
                 SetAdoFields "IVA", Total_IVA
                 SetAdoFields "Descuento", Total_Desc
                 SetAdoFields "Total_MN", Total_Factura
                 SetAdoFields "Saldo_MN", 0
                 SetAdoFields "Nota", "Facturas de " & MesesLetras(NoMeses) & ", Doc. " & SubCta
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
                 SetAdoFields "Imp_Mes", FA.Imp_Mes
                 SetAdoFields "Tipo_Pago", Tipo_Pago
                 
                 If Len(TBenef.Representante) > 1 And Len(TBenef.RUC_CI_Rep) > 1 Then
                    SetAdoFields "Razon_Social", TBenef.Representante
                    SetAdoFields "RUC_CI", TBenef.RUC_CI_Rep
                    SetAdoFields "TB", TBenef.TD_Rep
                 Else
                    Select Case TBenef.TD
                      Case "C", "R", "P"
                           SetAdoFields "Razon_Social", TBenef.Cliente
                           SetAdoFields "RUC_CI", TBenef.CI_RUC
                           SetAdoFields "TB", TBenef.TD
                      Case Else
                           SetAdoFields "Razon_Social", "CONSUMIDOR FINAL"
                           SetAdoFields "RUC_CI", "9999999999999"
                           SetAdoFields "TB", "R"
                    End Select
                 End If
                 
                 If IsNumeric(Codigo2) And Val(Codigo2) <> Year(FechaTexto) Then
                    FA.Cta_CxP = FA.Cta_CxP_Anterior
                    Cta_Cobrar = FA.Cta_CxP_Anterior
                 End If
                 
'''                 If AdoAux.Recordset.RecordCount > 0 Then
'''                    SetAdoFields "Razon_Social", AdoAux.Recordset.Fields("Representante")
'''                    SetAdoFields "RUC_CI", AdoAux.Recordset.Fields("Cedula_R")
'''                    SetAdoFields "TB", AdoAux.Recordset.Fields("TD")
'''                 Else
'''                    SetAdoFields "Razon_Social", "CONSUMIDOR FINAL"
'''                    SetAdoFields "RUC_CI", "9999999999999"
'''                    SetAdoFields "TB", "R"
'''                 End If
                 SetAdoUpdate
                 
                'Abono de la Factura
                 SetAdoAddNew "Trans_Abonos"
                 SetAdoFields "T", Cancelado
                 SetAdoFields "TP", FA.TC
                 SetAdoFields "CodigoC", CodigoCli
                 SetAdoFields "Fecha", FechaTexto
                 SetAdoFields "Factura", Factura_No
                 SetAdoFields "Abono", Total_Factura
                 SetAdoFields "Recibo_No", TrimStrg(SubCta)
                 SetAdoFields "Cheque", TBenef.Grupo_No
                 SetAdoFields "Cta", Cta_Banco
                 SetAdoFields "Cta_CxP", FA.Cta_CxP
                 SetAdoFields "Serie", FA.Serie
                 SetAdoFields "Autorizacion", FA.Autorizacion
''                 SetAdoFields "Mes", AdoProducto.Recordset.Fields("Mes")
''                 SetAdoFields "Mes_No", AdoProducto.Recordset.Fields("Num_Mes")
                 Select Case TextoBanco
                   Case "BIZBANCKPACIFICO"
                        TipoDeAbono = "DEPOSITO EN EL BANCO "
                   Case "PICHINCHA"
                        If Tipo_Carga >= 1 Then
                           TipoDeAbono = "DEPOSITO EN EL BANCO "
                        Else
                           TipoDeAbono = "DEP. " & Codigo3
                           If SubCta <> Ninguno Then TipoDeAbono = TipoDeAbono & ". " & SubCta
                        End If
                   Case "INTERNACIONAL", "OTROSBANCOS"
                        TipoDeAbono = "DEBITO BANCARIO "
                   Case "BOLIVARIANO"
                        TipoDeAbono = "DEPOSITO EN EL BANCO "
                   Case "GUAYAQUIL"
                        TipoDeAbono = "DEPOSITO EN EL BANCO "
                   Case "COOPJEP"
                        TipoDeAbono = "DEPOSITO EN EL BANCO "
                   Case "PRODUBANCO"
                        TipoDeAbono = "DEPOSITO EN EL BANCO "
                   Case "TARJETAS"
                        TipoDeAbono = "CANCELACION TARJETA "
                   Case Else
                        TipoDeAbono = "DEPOSITO EN EL BANCO "
                 End Select
                 If Len(Codigo3) > 1 Then TipoDeAbono = TipoDeAbono & Codigo3
                 SetAdoFields "Banco", TrimStrg(TipoDeAbono)
                 
                 If Len(Codigo4) > 1 Then
                    SetAdoFields "Comprobante", Codigo4
                 Else
                    SetAdoFields "Comprobante", SubCta
                 End If
                 SetAdoUpdate
                 sSQL = "DELETE * " _
                      & "FROM Clientes_Facturacion " _
                      & "WHERE Codigo = '" & CodigoCli & "' " _
                      & "AND Item = '" & NumEmpresa & "' " _
                      & "AND D <> " & Val(adFalse) & " "
                 'Generar_Consulta_SQL "Clientes_Facturacion " & CodigoCli, sSQL
                 Ejecutar_SQL_SP sSQL
              End If
             .MoveNext
           Loop
       End If
      End With
   End If
   FA.Fecha_Corte = FechaSistema
   Procesar_Saldo_De_Facturas FRecaudacionBancosPreFa, True
  'Actualizar_Abonos_Facturas FA, True
   sSQL = "UPDATE Detalle_Factura " _
        & "SET Producto = CP.Producto " _
        & "FROM Detalle_Factura As DF, Catalogo_Productos As CP " _
        & "WHERE DF.Item = '" & NumEmpresa & "' " _
        & "AND DF.Periodo = '" & Periodo_Contable & "' " _
        & "AND DF.TC = '" & FA.TC & "' " _
        & "AND DF.Serie = '" & FA.Serie & "' " _
        & "AND DF.Factura BETWEEN " & Factura_Desde & " AND " & Factura_Hasta & " " _
        & "AND DF.Item = CP.Item " _
        & "AND DF.Periodo = CP.Periodo " _
        & "AND DF.Codigo = CP.Codigo_Inv "
   Ejecutar_SQL_SP sSQL
  
   RatonNormal
   Progreso_Final
   MsgBox "GENERACION DE FACTURAS DEL: " & FechaTexto & vbCrLf _
        & "SE GRABARON: " & Total_Alumnos & " ALUMNOS." & vbCrLf _
        & "DESDE LA " & FA.TC & ": " & Factura_Desde & " HASTA LA " & Factura_Hasta & vbCrLf _
        & "EL CIERRE DIARIO DE CAJA ES POR " & Moneda & " " & Format$(TotalIngreso, "#,##0.00") & vbCrLf _
        & "OBTENIDO DEL ARCHIVO: " & vbCrLf & RutaGeneraFile
   sSQL = "DELETE * " _
        & "FROM Asiento_F " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND CodigoU = '" & CodigoUsuario & "' "
   Ejecutar_SQL_SP sSQL
   DGFactura.Visible = True
   sSQL = "SELECT FECHA,Numero As FACTURA,RUTA As BENEFICIARIO,CODIGO,CANT,PRODUCTO,TOTAL," _
        & "HABIT As PERIODO,Mes,A_No As NO_MES,Codigo_Cliente,CodigoU,Item " _
        & "FROM Asiento_F " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND CodigoU = '" & CodigoUsuario & "' "
   SQLDec = "PRECIO " & CStr(Dec_PVP) & "|TOTAL 2|."
   Select_Adodc_Grid DGFactura, AdoFactura, sSQL, SQLDec
   RatonNormal
   DCLinea.SetFocus
   'Unload Me
End Sub

Private Sub Alumnos_Contabilidad()
Dim CaptionTemp As String
  CaptionTemp = FRecaudacionBancosPreFa.Caption
  Contador = 1
  sSQL = "SELECT * " _
       & "FROM Clientes " _
       & "WHERE FA <> " & Val(adFalse) & " "
  If CheqRangos.value <> 0 Then sSQL = sSQL & "AND Grupo BETWEEN '" & DCGrupoI & "' and '" & DCGrupoF & "' "
  sSQL = sSQL & "ORDER BY CI_RUC,Cliente,Grupo "
  Select_Adodc AdoClientes, sSQL
  With AdoClientes.Recordset
   If .RecordCount > 0 Then
       ProgBarra.Max = .RecordCount
       Do While Not .EOF
          FRecaudacionBancosPreFa.Caption = .fields("Grupo") & " - " & Format$(Contador / ProgBarra.Max, "00%")
         .fields("Num_Lista") = Format$(Contador, "000")
         .Update
         .MoveNext
          Contador = Contador + 1
       Loop
   End If
  End With
  sSQL = "SELECT CI_RUC,Cliente,Codigo,Grupo,TD,Num_Lista " _
       & "FROM Clientes " _
       & "WHERE FA <> " & Val(adFalse) & " "
  If CheqRangos.value <> 0 Then sSQL = sSQL & "AND Grupo BETWEEN '" & DCGrupoI & "' and '" & DCGrupoF & "' "
  sSQL = sSQL & "ORDER BY CI_RUC,Cliente,Grupo "
  Select_Adodc AdoClientes, sSQL
  RatonNormal
  GenerarDataTexto FRecaudacionBancosPreFa, AdoClientes
  FRecaudacionBancosPreFa.Caption = CaptionTemp
End Sub

Private Sub Command2_Click()
  Unload FRecaudacionBancosPreFa
End Sub

Private Sub DCLinea_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCLinea_LostFocus()
  FA.Cod_CxC = DCLinea
  FA.Fecha = FechaSistema
  Lineas_De_CxC FA
  Label6.Caption = " Aut. " & FA.Autorizacion & " " & FA.TC & " No. " & FA.Serie & "-"
  FA.Factura = ReadSetDataNum(FA.TC & "_SERIE_" & FA.Serie, True, False) 'Numero_Factura(FA)
  TextFacturaNo = FA.Factura
End Sub

Private Sub DCTablaSRI_Change()
Dim ColorFondo As Long
  
  TextoBanco = Ninguno
  With AdoTablaSRI.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Descripcion = '" & DCTablaSRI.Text & "' ")
       If Not .EOF Then TextoBanco = .fields("Abreviado")
   End If
  End With
    
  Select Case TextoBanco
    Case "PACIFICO":
         ColorFondo = &HFFC0C0
         RutaOrigen = RutaSistema & "\LOGOS\BIZBANCK.JPG"
         FRecaudacionBancosPreFa.Caption = "BANCO DEL PACIFICO (" & CodigoDelBanco & ")" & String(40, " ") & "EL ULTIMO CODIGO: " & Codigo
    Case "PICHINCHA":
         ColorFondo = &H80FFFF
         RutaOrigen = RutaSistema & "\LOGOS\PICHINCHA.GIF"
         FRecaudacionBancosPreFa.Caption = "BANCO DEL PICHINCHA (" & CodigoDelBanco & ")" & String(40, " ") & "EL ULTIMO CODIGO: " & Codigo
    Case "INTERNACIONAL"
         ColorFondo = &HFF8080
         RutaOrigen = RutaSistema & "\LOGOS\INTERNACIONAL.GIF"
         FRecaudacionBancosPreFa.Caption = "BANCO INTERNACIONAL (" & CodigoDelBanco & ")" & String(40, " ") & "EL ULTIMO CODIGO: " & Codigo
    Case "BOLIVARIANO":
         ColorFondo = &H808000
         RutaOrigen = RutaSistema & "\LOGOS\BOLIVARIANO.GIF"
         FRecaudacionBancosPreFa.Caption = "BANCO BOLIVARIANO (" & CodigoDelBanco & ")" & String(40, " ") & "EL ULTIMO CODIGO: " & Codigo
    Case "GUAYAQUIL":
         ColorFondo = &HFF8080
         RutaOrigen = RutaSistema & "\LOGOS\GUAYAQUIL.GIF"
         FRecaudacionBancosPreFa.Caption = "BANCO DE GUAYQUIL (" & CodigoDelBanco & ")" & String(40, " ") & "EL ULTIMO CODIGO: " & Codigo
    Case "COOPJEP"
         ColorFondo = &H80FF80
         RutaOrigen = RutaSistema & "\LOGOS\COOP_JEP.GIF"
         FRecaudacionBancosPreFa.Caption = "COOPERATIVA JEP (" & CodigoDelBanco & ")" & String(40, " ") & "EL ULTIMO CODIGO: " & Codigo
    Case "PRODUBANCO"
         ColorFondo = &HFFFFFF
         RutaOrigen = RutaSistema & "\LOGOS\PRODUBAN.JPG"
         FRecaudacionBancosPreFa.Caption = "PRODUBANCO (" & CodigoDelBanco & ")" & String(40, " ") & "EL ULTIMO CODIGO: " & Codigo
    Case "TARJETAS"
         ColorFondo = &HC0FFC0
         RutaOrigen = RutaSistema & "\LOGOS\TARJETAS.JPG"
         FRecaudacionBancosPreFa.Caption = "TARJETAS (" & CodigoDelBanco & ")" & String(40, " ") & "EL ULTIMO CODIGO: " & Codigo
    Case "POREXCEL"
         ColorFondo = Blanco
         RutaOrigen = RutaSistema & "\LOGOS\POREXCEL.JPG"
         FRecaudacionBancosPreFa.Caption = "TARJETAS (" & CodigoDelBanco & ")" & String(40, " ") & "EL ULTIMO CODIGO: " & Codigo
    Case Else
         ColorFondo = Blanco
         RutaOrigen = RutaSistema & "\LOGOS\OTROSBAN.JPG"
         FRecaudacionBancosPreFa.Caption = "OTROS BANCOS (" & CodigoDelBanco & ")" & String(40, " ") & "EL ULTIMO CODIGO: " & Codigo
  End Select
  PctBanco.BackColor = ColorFondo
  PctBanco.Picture = LoadPicture(RutaOrigen)   ', 0, 0, 5000, 1100
  FRecaudacionBancosPreFa.BackColor = ColorFondo
  Label1.BackColor = FRecaudacionBancosPreFa.BackColor
  Label3.BackColor = FRecaudacionBancosPreFa.BackColor
  Label4.BackColor = FRecaudacionBancosPreFa.BackColor
  Label5.BackColor = FRecaudacionBancosPreFa.BackColor
  Label8.BackColor = FRecaudacionBancosPreFa.BackColor
  MBFechaI.BackColor = FRecaudacionBancosPreFa.BackColor
  MBFechaF.BackColor = FRecaudacionBancosPreFa.BackColor
  CheqMatricula.BackColor = FRecaudacionBancosPreFa.BackColor
  Cheq_NumCodigos.BackColor = FRecaudacionBancosPreFa.BackColor
  CheqAlDia.BackColor = FRecaudacionBancosPreFa.BackColor
  DCLinea.BackColor = FRecaudacionBancosPreFa.BackColor
End Sub

Private Sub DGFactura_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  If CtrlDown And KeyCode = vbKeyF1 Then
     DGFactura.Visible = False
     GenerarDataTexto FRecaudacionBancosPreFa, AdoFactura
     DGFactura.Visible = True
  End If
End Sub

Private Sub Form_Activate()
  FechaValida MBFechaI
  FechaValida MBFechaF
  FechaValida MBFechaV
  
  NuevoComp = True
  ModificarComp = False
  CopiarComp = False
  Co.CodigoB = ""
  Co.Numero = 0
  Tipo_Carga = Leer_Campo_Empresa("Tipo_Carga_Banco")
 'Label4.Caption = "ORIGEN" & Space(18) & "COD: " & CodigoDelBanco
  FRecaudacionBancosPreFa.Caption = "FACTURACION DE BANCOS (" & CodigoDelBanco & ")"
  
  sSQL = "SELECT Descripcion, Abreviado, ID " _
       & "From Tabla_Referenciales_SRI " _
       & "WHERE Tipo_Referencia = 'BANCOS Y COOP' " _
       & "AND Abreviado <> '.' " _
       & "AND TPFA <> " & Val(adFalse) & " " _
       & "ORDER BY Descripcion "
  SelectDB_Combo DCTablaSRI, AdoTablaSRI, sSQL, "Descripcion"
  
  sSQL = "SELECT * " _
       & "FROM Catalogo_Lineas " _
       & "WHERE TL <> " & Val(adFalse) & " " _
       & "AND Fact NOT IN ('CP','NC','LC') " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "ORDER BY Codigo "
  SelectDB_Combo DCLinea, AdoLinea, sSQL, "Concepto"
  With AdoLinea.Recordset
   If .RecordCount > 0 Then
       Cta_Cobrar = .fields("CxC")
       CxC_Clientes = .fields("Concepto")
       LogoFactura = .fields("Logo_Factura")
       AltoFactura = .fields("Largo")
       AnchoFactura = .fields("Ancho")
       EspacioFactura = .fields("Espacios")
       Pos_Factura = .fields("Pos_Factura")
       Individual = .fields("Individual")
       TipoFactura = .fields("Fact")
    End If
  End With
 'Catalogo de los Rubros ha facturar
  
  sSQL = "SELECT * " _
       & "FROM Catalogo_Productos " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND TC = 'P' " _
       & "ORDER BY Codigo_Inv "
  Select_Adodc AdoProducto, sSQL
 'Alumnos/Clientes que estan activados para Generar las Facturas
  
  Codigo = Ninguno
  
  sSQL = "SELECT Codigo & Space(5) & Cuenta As NomCuenta,* " _
       & "FROM Catalogo_Cuentas " _
       & "WHERE TC = 'BA' " _
       & "AND DG = 'D' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "ORDER BY Codigo "
  SelectDB_Combo DCBanco, AdoBanco, sSQL, "NomCuenta"
  
  sSQL = "SELECT Grupo, Count(Grupo) As Cantidad " _
       & "FROM Clientes " _
       & "WHERE FA <> " & Val(adFalse) & " " _
       & "AND LEN(Grupo) > 1 "
  If Mas_Grupos Then sSQL = sSQL & "AND DirNumero = '" & NumEmpresa & "' "
  sSQL = sSQL _
       & "GROUP BY Grupo " _
       & "ORDER BY Grupo "
  SelectDB_Combo DCGrupoI, AdoGrupo, sSQL, "Grupo"
  SelectDB_Combo DCGrupoF, AdoGrupo, sSQL, "Grupo"
  If AdoGrupo.Recordset.RecordCount > 0 Then
     DCGrupoI.Text = AdoGrupo.Recordset.fields("Grupo")
     AdoGrupo.Recordset.MoveLast
     DCGrupoF.Text = AdoGrupo.Recordset.fields("Grupo")
  End If
 'Datos del Banco a recaudar
  Tipo_Carga = Leer_Campo_Empresa("Tipo_Carga_Banco")
  Costo_Banco = Leer_Campo_Empresa("Costo_Bancario")
  Cta_Bancaria = Leer_Campo_Empresa("Cta_Banco")
  Cta_Gasto_Banco = Leer_Seteos_Ctas("Cta_Gasto_Bancario")
  RatonNormal
End Sub

Private Sub Form_Deactivate()
  FRecaudacionBancosPreFa.WindowState = vbMaximized
End Sub

Private Sub Form_Load()
 'CentrarForm FRecaudacionBancosPreFa
  If CodigoUsuario = "ACCESO02" Then
     Toolbar1.buttons("AlumnosContabilidad").Enabled = True
     Toolbar1.buttons("RenumerarCodigos").Enabled = True
  End If
  RutaBackupXX = ""
  ConectarAdodc AdoAux
  ConectarAdodc AdoAbono
  ConectarAdodc AdoBanco
  ConectarAdodc AdoGrupo
  ConectarAdodc AdoLinea
  ConectarAdodc AdoDetalle
  ConectarAdodc AdoFactura
  ConectarAdodc AdoProducto
  ConectarAdodc AdoTablaSRI
  ConectarAdodc AdoPendiente
  ConectarAdodc AdoClientes
  ConectarAdodc AdoClientesMatriculas
  
  TxtFile.width = MDI_X_Max - TxtFile.Left - 100
  AdoFactura.width = MDI_X_Max - AdoFactura.Left - LabelAbonos.width - Label4.width - 100
  DGFactura.width = MDI_X_Max - DGFactura.Left - 100
  
  TxtFile.Height = ((MDI_Y_Max - TxtFile.Top) / 2) - 200
  DGFactura.Top = TxtFile.Top + TxtFile.Height + 10
  DGFactura.Height = DGFactura.Top - 3300
  AdoFactura.Top = DGFactura.Top + DGFactura.Height + 10
  Frame1.width = MDI_X_Max - Frame1.Left - 50
  DCBanco.width = Frame1.width - 200
  LabelAbonos.Top = DGFactura.Top + DGFactura.Height + 10
  Label4.Top = DGFactura.Top + DGFactura.Height + 10
  LabelAbonos.Left = AdoFactura.Left + AdoFactura.width
  'Label4.width
  
End Sub

Private Sub MBFechaI_GotFocus()
  MarcarTexto MBFechaI
End Sub

Private Sub MBFechaI_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBFechaI_LostFocus()
  MBFechaF.Text = UltimoDiaMes(MBFechaI.Text)
  FechaValida MBFechaI
  sSQL = "SELECT * " _
       & "FROM Catalogo_Lineas " _
       & "WHERE TL <> " & Val(adFalse) & " " _
       & "AND #" & BuscarFecha(MBFechaI) & "# <= Vencimiento  " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND Fact IN ('NV','FA') " _
       & "ORDER BY Codigo "
  SelectDB_Combo DCLinea, AdoLinea, sSQL, "Concepto"
  FA.Fecha = MBFechaI
  FA.TC = "FA"
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

Private Sub MBFechaV_GotFocus()
  MarcarTexto MBFechaV
End Sub

Private Sub MBFechaV_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBFechaV_LostFocus()
  FechaValida MBFechaV
End Sub

Public Sub Generar_Pacifico(RutaDelArchivo As String)
'022750270
'MsgBox RutaSysBases
RutaGeneraFile = RutaDelArchivo
NumFileFacturas = FreeFile
TipoDoc = "0"
Contador = 0
Progreso_Barra.Mensaje_Box = "Generando los Archivos al banco"
Progreso_Iniciar
FechaTexto = BuscarFecha(MBFechaI)
'MsgBox RutaGeneraFile
TxtFile = ""
Open RutaGeneraFile For Output As #NumFileFacturas ' Abre el archivo.
With AdoFactura.Recordset
 If .RecordCount > 0 Then
    .MoveFirst
     TxtFile = "TOTAL NOMINA DE RECAUDACION:" & vbCrLf
     Progreso_Barra.Valor_Maximo = .RecordCount
     Codigo3 = TrimStrg(MidStrg(NombreEmpresa, 1, 30))
     Total = 0
     TotalIngreso = 0
     Total_Factura = 0
     IE = 1
     JE = 1
     KE = 1
     Grupo_No = .fields("Grupo")
     Codigo = .fields("Codigo")
     Do While Not .EOF
        'MsgBox .Fields("Grupo")
        If Grupo_No <> .fields("Grupo") Then
           Codigo4 = Format$(Total, "#,##0.00")
           Codigo4 = String(13 - Len(Codigo4), " ") & Format$(Total, "#,##0.00")
           TxtFile = TxtFile _
                   & "Grupo: " & Grupo_No & vbTab & "Resumen de Registros por Grupo: " & JE & vbTab & "Total a Recaudar USD" & vbTab & Codigo4 & vbCrLf
           JE = 0
           Total = 0
           IE = IE + 1
           Grupo_No = .fields("Grupo")
           Codigo = .fields("Codigo")
        End If
        If Codigo <> .fields("Codigo") Then
           JE = JE + 1
           KE = KE + 1
           Codigo = .fields("Codigo")
        End If
        FRecaudacionBancosPreFa.Caption = .fields("Grupo") & " - " & Format$(Contador / .RecordCount, "00%")
        Contador = Contador + 1
        Progreso_Esperar
        CodigoCli = .fields("CI_RUC")
        NombreCliente = TrimStrg(MidStrg(.fields("Cliente"), 1, 30))
        Codigo1 = TrimStrg(MidStrg(.fields("Direccion"), 1, 30))
        FechaTexto = " " & UCaseStrg(MidStrg(MesesLetras(.fields("Num_Mes")), 1, 3)) & " " & MidStrg(.fields("Periodo"), 3, 2)
        Codigo2 = UCaseStrg(.fields("Grupo")) & " " & UCaseStrg(MidStrg(MesesLetras(.fields("Num_Mes")), 1, 3)) & " " & .fields("Periodo")
        Total_Factura = .fields("Valor_Cobro")
        Total = Total + Total_Factura
        TotalIngreso = TotalIngreso + Total_Factura
        I = Int(Total_Factura)
        J = (Total_Factura - Int(Total_Factura)) * 100
        'MsgBox Grupo_No & "(" & JE & ")" & vbCrLf & Total_Factura & vbCrLf & Total & vbCrLf & TotalIngreso
      ' Empieza la trama por Alumno
        Print #NumFileFacturas, "1";                                                  ' Localidad
        Print #NumFileFacturas, "OCP";                                                ' Transsaccion
        Print #NumFileFacturas, "OC";                                                 ' Codigo de Servicio
        Print #NumFileFacturas, "  ";                                                 ' Tipo de Cuenta
        Print #NumFileFacturas, String(8, " ");                                       ' Numero de Cuenta
        Print #NumFileFacturas, Format$(I, "0000000000000") & Format$(J, "00");       ' Valor
       'Print #NumFileFacturas, CodigoCli & FechaTexto;                               ' Codigo del Alumno : FechaTexto
        Print #NumFileFacturas, CodigoCli & String(15 - Len(CodigoCli), " ");         ' Codigo del Alumno : FechaTexto
        Print #NumFileFacturas, Codigo2 & String(20 - Len(Codigo2), " ");             ' Referencia
        Print #NumFileFacturas, "RE";                                                 ' Forma de Pago
        Print #NumFileFacturas, "USD";                                                ' Moneda
        Print #NumFileFacturas, NombreCliente & String(30 - Len(NombreCliente), " "); ' Nombre del Alumno
        Print #NumFileFacturas, "01";                                                 ' Localidad
        Print #NumFileFacturas, String(2, " ");                                       ' Agencia de Retiro
        Print #NumFileFacturas, " ";                                                  ' Tipo NUC
        Print #NumFileFacturas, CodigoCli & String(14 - Len(CodigoCli), " ");         ' Numero Unico del Beneficiario
        Print #NumFileFacturas, String(10, " ");                                      ' Telefono
        Print #NumFileFacturas, " ";                                                  ' NUC Ordenante
        Print #NumFileFacturas, String(14, " ");                                      ' Numero Unico del Ordenante
        Print #NumFileFacturas, Codigo3 & String(30 - Len(Codigo3), " ");             ' Nombre Colegio
        Print #NumFileFacturas, Format$(Contador, "000000");                          ' Secuencial
        Print #NumFileFacturas, "30";                                                 ' Banco del Pacifico
        Print #NumFileFacturas, String(20, " ")                                       ' Numero de cuenta
       .MoveNext
     Loop
     Codigo4 = Format$(Total, "#,##0.00")
     Codigo4 = String(13 - Len(Codigo4), " ") & Format$(Total, "#,##0.00")
     TxtFile = TxtFile _
             & "Grupo: " & Grupo_No & vbTab & "Resumen de Registros por Grupo: " & JE & vbTab & "Total a Recaudar USD" & vbTab & Codigo4 & vbCrLf
     
     Codigo4 = Format$(TotalIngreso, "#,##0.00")
     Codigo4 = String(13 - Len(Codigo4), " ") & Format$(TotalIngreso, "#,##0.00")
     TxtFile = TxtFile _
             & String(90, "-") & vbCrLf _
             & "Total Grupos: " & IE & vbTab & "Total Alumnos: " & KE & vbTab & vbTab & "Total a Recaudar USD" & vbTab & Codigo4 & vbCrLf
     
 End If
End With
Close #NumFileFacturas
Progreso_Final
RatonNormal
MsgBox UCaseStrg("Fin del Proceso " & vbCrLf & vbCrLf & "El Archivo se Generara en: " & vbCrLf & vbCrLf & RutaGeneraFile)
End Sub

Public Sub Generar_Archivo_CSV(RutaDelArchivo As String)
'022750270
'MsgBox RutaSysBases
RutaGeneraFile = RutaDelArchivo
NumFileFacturas = FreeFile
Contador = 0
Progreso_Barra.Mensaje_Box = "Generando el Archivo CSV"
Progreso_Iniciar
FechaTexto = BuscarFecha(MBFechaI)
'MsgBox RutaGeneraFile
Open RutaGeneraFile For Output As #NumFileFacturas ' Abre el archivo.
With AdoFactura.Recordset
 If .RecordCount > 0 Then
    .MoveFirst
     Progreso_Barra.Valor_Maximo = .RecordCount
     Print #NumFileFacturas, "CODIGO;";         ' Codigo del Alumno : FechaTexto
     Print #NumFileFacturas, "REFERENCIA;";             ' Referencia
     Print #NumFileFacturas, "NOMBRE BENEFICIARIO;";  ' Nombre del Alumno
     Print #NumFileFacturas, "CODIGO INT;";          ' Numero Unico del Beneficiario
     Print #NumFileFacturas, "VALOR A COBRAR;"                    ' Valor
     
     Do While Not .EOF
        FRecaudacionBancosPreFa.Caption = .fields("Grupo") & " - " & Format$(Contador / .RecordCount, "00%")
        Contador = Contador + 1
        Progreso_Esperar
        CodigoCli = .fields("CI_RUC")
        NombreCliente = TrimStrg(MidStrg(.fields("Cliente"), 1, 30))
        Codigo1 = TrimStrg(MidStrg(.fields("Direccion"), 1, 30))
        FechaTexto = " " & UCaseStrg(MidStrg(MesesLetras(.fields("Num_Mes")), 1, 3)) & " " & MidStrg(.fields("Periodo"), 3, 2)
        Codigo2 = "01/" & .fields("Num_Mes") & "/" & .fields("Periodo")
        Total_Factura = .fields("Valor_Cobro")
       'MsgBox Grupo_No & "(" & JE & ")" & vbCrLf & Total_Factura & vbCrLf & Total & vbCrLf & TotalIngreso
      ' Empieza la trama por Alumno
        Print #NumFileFacturas, CodigoCli & ";";          ' Codigo del Alumno : FechaTexto
        Print #NumFileFacturas, Codigo2 & ";";              ' Referencia
        Print #NumFileFacturas, NombreCliente & String(30 - Len(NombreCliente), " ") & ";"; ' Nombre del Alumno
        Print #NumFileFacturas, CodigoCli & String(14 - Len(CodigoCli), " ") & ";";         ' Numero Unico del Beneficiario
        Print #NumFileFacturas, Format$(Total_Factura, "#,##0.00") & ";"                    ' Valor
       .MoveNext
     Loop
 End If
End With
Close #NumFileFacturas
Progreso_Final
RatonNormal
End Sub

Public Sub Generar_Bolivariano(RutaDelArchivo As String)
'022750270
'MsgBox RutaSysBases
RutaGeneraFile = RutaDelArchivo
NumFileFacturas = FreeFile
TipoDoc = "0"
'If CheqMatricula.Value = 1 Then TipoDoc = "1"
Contador = 0
'ProgBarra.value = 0
'ProgBarra.Min = 0
FechaTexto = BuscarFecha(MBFechaF.Text)
FechaTexto1 = Format$(MBFechaV, "MM/dd/yyyy")
'MsgBox RutaGeneraFile
Open RutaGeneraFile For Output As #NumFileFacturas ' Abre el archivo.
With AdoFactura.Recordset
 If .RecordCount > 0 Then
    'Cabecera
     Print #NumFileFacturas, "999";
     Print #NumFileFacturas, CodigoDelBanco;
     Print #NumFileFacturas, TipoDoc;
     Print #NumFileFacturas, Space(11);
     Print #NumFileFacturas, FechaTexto1
    .MoveFirst
     'ProgBarra.Max = .RecordCount
     'Trama / Detalle
     Do While Not .EOF
        FRecaudacionBancosPreFa.Caption = .fields("Grupo") & " - " & Format$(Contador / .RecordCount, "00%")
        SaldoPendiente = 0
        Total_Factura = 0
        Monto_Total = 0
        Total = 0
        'ProgBarra.value = Contador
        CodigoCli = .fields("CI_RUC")
        Codigo = "0"
        For I = 1 To Len(.fields("CI_RUC"))
            If IsNumeric(MidStrg(.fields("CI_RUC"), I, 1)) Then Codigo = Codigo & MidStrg(.fields("CI_RUC"), I, 1)
        Next I
        Codigo = TrimStrg(Str(Val(Codigo)))
        Codigo = Codigo & String(15 - Len(Codigo), " ")
      ' MsgBox "|" & Codigo & "|"
        NombreCliente = SetearBlancos(MidStrg(.fields("Cliente"), 1, 30), 30, 0, False)
        Codigo1 = TrimStrg(MidStrg(SinEspaciosIzq(.fields("Direccion")), 1, 15))
        Codigo3 = TrimStrg(MidStrg(SinEspaciosDer(.fields("Direccion")), 1, 3))
        Codigo2 = TrimStrg(MidStrg(.fields("Direccion"), Len(Codigo1) + 1, Len(.fields("Direccion"))))
        Codigo4 = MidStrg(.fields("Casilla"), 1, 10)
        Saldo_ME = 0: Total_Desc = 0: SaldoPendiente = 0
'        If AdoPendiente.Recordset.RecordCount > 0 Then
'           AdoPendiente.Recordset.MoveFirst
'           AdoPendiente.Recordset.Find ("CI_RUC Like '" & CodigoCli & "' ")
'           If Not AdoPendiente.Recordset.EOF Then SaldoPendiente = AdoPendiente.Recordset.Fields("Saldo_Pend")
'        End If
'        If CheqPend.Value = 1 Then SaldoPendiente = .Fields("Total_MN")
        Total_Factura = .fields("Valor_Cobro")
        Monto_Total = Total_Factura
        SaldoPendiente = Total_Factura
        Total = SaldoPendiente
        If Codigo1 = "" Then Codigo1 = Ninguno
        If Codigo2 = "" Then Codigo2 = Ninguno
        If Codigo3 = "" Then Codigo3 = Ninguno
        Codigo2 = TrimStrg(MidStrg(Codigo2, 1, Len(Codigo2) - Len(SinEspaciosDer(Codigo2))))
        Codigo1 = SetearBlancos(Codigo1, 15, 0, False)
        Codigo2 = SetearBlancos(Codigo2, 15, 0, False)
        Codigo3 = SetearBlancos(Codigo3, 3, 0, False)
        Codigo4 = SetearBlancos(Codigo4, 10, 0, False)
        If TrimStrg(Codigo4) = Ninguno Then Codigo4 = String(10, " ")
      ' Total = Total - Monto_Total
        If Total < 0 Then Total = 0
      ' Empieza la trama por Alumno
        'MsgBox NombreCliente & vbCrLf & Total
        Print #NumFileFacturas, CodigoDelBanco;                        ' Colegio/Institucion
        Print #NumFileFacturas, Codigo;                                ' Codigo Alumno
        Print #NumFileFacturas, Format$(MBFechaI, "MM/dd/yyyy");        ' Fecha Pen: Mifecha/FechaTexto = FechaTexto1
        Print #NumFileFacturas, TipoDoc & "  ";                        ' Proceso
        Print #NumFileFacturas, Format$(Total, "00000000.00");          ' Valor
        Print #NumFileFacturas, Format$(MBFechaV, "MM/dd/yyyy");        ' FechaTexto/Fecha Cobis
        Print #NumFileFacturas, "01/01/1900";                          ' Fecha Pago "01/01/1900";
        Print #NumFileFacturas, "N";                                   ' Estado = N
        Print #NumFileFacturas, Sin_Signos_Especiales(NombreCliente); ' Nombre Alumno
        Print #NumFileFacturas, Codigo2;                               ' Nombre del Curso
        Print #NumFileFacturas, Codigo3;                               ' Nombre del Paralelo
        Print #NumFileFacturas, Codigo1;                               ' Nombre de la Seccion
        Print #NumFileFacturas, Format$(Monto_Total, "00000000.00");    ' Valor Mes
        Print #NumFileFacturas, Codigo4;                               ' Pago por Deposito de Cuenta
        Print #NumFileFacturas, "1";                                   ' Moneda = 1
        Print #NumFileFacturas, Format$(Total, "00000000.00");          ' Valor 2
        Print #NumFileFacturas, Format$(Total, "00000000.00")           ' Valor 1
        Contador = Contador + 1
       .MoveNext
     Loop
 End If
End With
Close #NumFileFacturas
'ProgBarra.value = ProgBarra.Max
RatonNormal
MsgBox "Fin del Proceso " & vbCrLf & vbCrLf & "El Archivo se Generara en: " & vbCrLf & vbCrLf & RutaGeneraFile
End Sub

Public Sub Generar_Pichincha()
Dim AuxNumEmp As String
Dim DiaV As Integer
Dim MesV As Integer
Dim AñoV As Integer
Dim Fecha_Meses As String
Dim CamposFile() As Campos_Tabla
Dim Total_Banco As Currency
  Progreso_Barra.Mensaje_Box = "GENERANDO ARCHIVO PARA EL BANCO"
  Progreso_Iniciar
  Total_Banco = 0
  Mifecha = BuscarFecha(MBFechaI)
  MiMes = Format$(Month(MBFechaI), "00")
  FechaFin = BuscarFecha(UltimoDiaMes(MBFechaI))
  TextoImprimio = ""
  AuxNumEmp = NumEmpresa
  Cta_Cobrar = Ninguno
  FechaTexto = FechaSistema
' Comenzamos a generar el archivo: SCRECXX.TXT
  EsComa = True
  Estab = False
  Contador = 0
  Factura_No = 0
  Fecha_Meses = MBFechaI & " al " & MBFechaF
  Fecha_Meses = Replace(Fecha_Meses, "/", "-")
  
  Mes = Month(MBFechaI)
  Anio = Val(MidStrg(Format$(Year(MBFechaI), "0000"), 2, 3))
  Dia = "15"
  
  NumFileFacturas = FreeFile
  RutaGeneraFile = UCaseStrg(RutaSysBases & "\BANCO\FACTURAS\SCREC " & Fecha_Meses & ".TXT")
  Open RutaGeneraFile For Output As #NumFileFacturas ' Abre el archivo.
  
  NumFileFacturas2 = FreeFile
  RutaGeneraFile2 = UCaseStrg(RutaSysBases & "\BANCO\FACTURAS\SCCOB " & Fecha_Meses & ".TXT")
  Open RutaGeneraFile2 For Output As #NumFileFacturas2 ' Abre el archivo.
  
  With AdoFactura.Recordset
   If .RecordCount > 0 Then
       Progreso_Barra.Valor_Maximo = .RecordCount
      .MoveFirst
       Do While Not .EOF
          Contador = Contador + 1
          CodigoCli = .fields("Codigo")
          NombreCliente = Sin_Signos_Especiales(.fields("Cliente"))
          'MsgBox NombreCliente
          Factura_No = Factura_No + 1
          Total = .fields("Valor_Cobro")
          Saldo = .fields("Valor_Cobro") * 100
          CodigoP = .fields("CI_RUC")
          CodigoC = CStr(Val(.fields("CI_RUC")))
          CodigoC = CodigoC & String(Abs(4 - Len(CodigoC)), " ")
          DireccionCli = Sin_Signos_Especiales(.fields("Direccion"))
          Codigo3 = SinEspaciosDer(DireccionCli)
          DireccionCli = TrimStrg(MidStrg(DireccionCli, 1, Len(DireccionCli) - Len(Codigo3)))
          Codigo3 = TrimStrg(SinEspaciosDer(DireccionCli))
          Codigo1 = Format$(MidStrg(GrupoNo, 1, 1), "00")
          Codigo4 = MidStrg(.fields("Codigo_Inv") & " " & .fields("Producto"), 1, 33)
          Codigo4 = Codigo4 & String(33 - Len(Codigo4), " ") & Format$(.fields("Num_Mes"), "00") & " " & .fields("Periodo")
                    
'''          If Tipo_Carga = 2 Then
'''
'''          Else
'''             If CheqMatricula.value = 1 Then
'''                Codigo4 = "MATRICULAS Y PENSION DE " & MidStrg(MesesLetras(Month(MBFechaI)), 1, 3) & "-" & Year(MBFechaI)
'''             Else
'''                Codigo4 = "PENSION ACUMULADA DE " & MidStrg(MesesLetras(Month(.Fields("Fecha"))), 1, 3) & "-" & Year(.Fields("Fecha"))
'''             End If
'''          End If
         'MsgBox Tipo_Carga & vbCrLf & Cta_Bancaria
          If Len(.fields("Actividad")) >= 3 Then
             Progreso_Barra.Mensaje_Box = "GENERANDO ARCHIVO: " & RutaGeneraFile2
             Print #NumFileFacturas2, "CO" & vbTab;
             Print #NumFileFacturas2, Cta_Bancaria & vbTab;
             Print #NumFileFacturas2, Contador & vbTab;
             Print #NumFileFacturas2, Format$(Factura_No, "0000000000") & vbTab;
             Print #NumFileFacturas2, CodigoP & vbTab;
             Print #NumFileFacturas2, "USD" & vbTab;
             Print #NumFileFacturas2, Format$(Saldo, "0000000000000") & vbTab;
             Print #NumFileFacturas2, "CTA" & vbTab;
             Print #NumFileFacturas2, "10" & vbTab;
             NumStrg = SinEspaciosIzq(.fields("Actividad"))
             If Len(NumStrg) = 3 Then
                Print #NumFileFacturas2, SinEspaciosIzq(.fields("Actividad")) & vbTab;      'CTE/AHO
                Print #NumFileFacturas2, SinEspaciosDer(.fields("Actividad")) & vbTab;     'No. Cta Cte/Aho
             Else
                Print #NumFileFacturas2, vbTab;      'CTE/AHO
                Print #NumFileFacturas2, vbTab;      'No. Cta Cte/Aho
             End If
             Print #NumFileFacturas2, "R" & vbTab;
             Print #NumFileFacturas2, RUC & vbTab;
             Print #NumFileFacturas2, Format$(.fields("Num_Mes"), "00") & " " & MidStrg(NombreCliente, 1, 37) & vbTab;
             Print #NumFileFacturas2, vbTab;
             Print #NumFileFacturas2, vbTab;
             Print #NumFileFacturas2, vbTab;
             Print #NumFileFacturas2, vbTab;
             Print #NumFileFacturas2, Month(.fields("Fecha")) & vbTab;
             Print #NumFileFacturas2, UCaseStrg(Codigo4) & vbTab;
             Print #NumFileFacturas2, Format$(Saldo, "0000000000000") & vbTab
          Else
             Progreso_Barra.Mensaje_Box = "GENERANDO ARCHIVO: " & RutaGeneraFile
             If Tipo_Carga >= 1 Then
               'Tipo Gualaceo
                Print #NumFileFacturas, "CO" & vbTab;
                Print #NumFileFacturas, Format$(Val(CodigoC), "0000000000") & vbTab;
                Print #NumFileFacturas, "USD" & vbTab;
                Print #NumFileFacturas, Saldo & vbTab;
                Print #NumFileFacturas, "REC" & vbTab;
                Print #NumFileFacturas, vbTab;
                Print #NumFileFacturas, vbTab;
                Print #NumFileFacturas, UCaseStrg(Codigo4) & vbTab;
                Print #NumFileFacturas, "N" & vbTab;
                Print #NumFileFacturas, Format$(Val(CodigoC), "0000000000") & vbTab;
                Print #NumFileFacturas, Format$(.fields("Num_Mes"), "00") & " " & MidStrg(NombreCliente, 1, 37) & vbTab
             Else
               'Tipo General
                Print #NumFileFacturas, "CO" & vbTab;
                Print #NumFileFacturas, Cta_Bancaria & vbTab;
                Print #NumFileFacturas, Contador & vbTab;
                Print #NumFileFacturas, Format$(Factura_No, "0000000000") & vbTab;
                Print #NumFileFacturas, CodigoP & vbTab;
                Print #NumFileFacturas, "USD" & vbTab;
                Print #NumFileFacturas, Format$(Saldo, "0000000000000") & vbTab;
                Print #NumFileFacturas, "REC" & vbTab;
                Print #NumFileFacturas, "10" & vbTab;
                Print #NumFileFacturas, vbTab;           'CTE/AHO
                Print #NumFileFacturas, "0" & vbTab;     'No. Cta Cte/Aho
                Print #NumFileFacturas, "R" & vbTab;
                Print #NumFileFacturas, RUC & vbTab;
                Print #NumFileFacturas, Format$(.fields("Num_Mes"), "00") & " " & MidStrg(NombreCliente, 1, 37) & vbTab;
                Print #NumFileFacturas, vbTab;
                Print #NumFileFacturas, vbTab;
                Print #NumFileFacturas, vbTab;
                Print #NumFileFacturas, vbTab;
                Print #NumFileFacturas, Month(.fields("Fecha")) & vbTab;
                Print #NumFileFacturas, UCaseStrg(Codigo4) & vbTab;
                'Print #NumFileFacturas, "Pensión Acumulada" & vbTab;
                Print #NumFileFacturas, Format$(Saldo, "0000000000000") & vbTab
             End If
          End If
          Progreso_Esperar
         .MoveNext
       Loop
   End If
  End With
  Close #NumFileFacturas
  Close #NumFileFacturas2
  
  RatonNormal
  Progreso_Final
 'FBancoPichincha.Caption = "FACTURACION DE BANCO INTERNACIONAL"
  LabelAbonos.Caption = Format$(Total_Banco, "#,##0.00")
  MsgBox "SE GENERARON LOS SIGUIENTES ARCHIVOS:" & vbCrLf & vbCrLf _
       & UCaseStrg(RutaSysBases & "\BANCO\FACTURAS\SCREC " & Fecha_Meses & ".TXT") & vbCrLf & vbCrLf _
       & UCaseStrg(RutaSysBases & "\BANCO\FACTURAS\SCCOB " & Fecha_Meses & ".TXT")
End Sub

Public Sub Generar_Internacional()
Dim obj_Excel As Object
Dim Obj_Libro As Object
Dim Obj_Hoja As Object
Dim ArchivoTexto As String
Dim ArchivoExcel As String
Dim ArchivoExcel2 As String
Dim NombreArchivos As String

Dim AuxNumEmp As String
Dim DiaV As Integer
Dim MesV As Integer
Dim AñoV As Integer
Dim Fecha_Meses As String
Dim CamposFile() As Campos_Tabla
Dim Total_Banco As Currency
Dim TipoCta As String

  Progreso_Barra.Mensaje_Box = "Generando los Archivos al banco"
  Progreso_Iniciar
  Total_Banco = 0
  Mifecha = BuscarFecha(MBFechaI)
  MiMes = Format$(Month(MBFechaI), "00")
  FechaFin = BuscarFecha(UltimoDiaMes(MBFechaI))
  TextoImprimio = ""
  AuxNumEmp = NumEmpresa
  Cta_Cobrar = Ninguno
  FechaTexto = FechaSistema
' Comenzamos a generar el archivo: Internacional Recaudaciones.TXT
  EsComa = True
  Estab = False
  Contador = 0
  Factura_No = 0
  Fecha_Meses = MBFechaI & " al " & MBFechaF
  Fecha_Meses = Replace(Fecha_Meses, "/", "-")
  
  ArchivoTexto = UCaseStrg(RutaSysBases & "\BANCO\FACTURAS\Internacional Recaudaciones " & Fecha_Meses & ".TXT")
  ArchivoExcel = UCaseStrg(RutaSysBases & "\BANCO\FACTURAS\Internacional Recaudaciones " & Fecha_Meses & ".XLSX")   ' xls
  NombreArchivos = ArchivoTexto & vbCrLf & vbCrLf
  
 'Genera un archvo de texto
  NumFileFacturas = FreeFile
  Open ArchivoTexto For Output As #NumFileFacturas ' Abre el archivo.
  
 'Genera un nuevo libro de Excel
  Set obj_Excel = CreateObject("Excel.Application")
  Set Obj_Libro = obj_Excel.Workbooks.Add
  Set Obj_Hoja = Obj_Libro.Worksheets(1)
  obj_Excel.ActiveSheet.Select ' La selecciona
  obj_Excel.ActiveSheet.Name = "Internacional Recaudaciones"
  
  With AdoFactura.Recordset
   If .RecordCount > 0 Then
       Progreso_Barra.Valor_Maximo = Progreso_Barra.Valor_Maximo + (.RecordCount * 2)
      .MoveFirst
       Do While Not .EOF
          CodigoCli = .fields("Codigo")
          GrupoNo = .fields("Grupo")
          NombreCliente = Sin_Signos_Especiales(.fields("Cliente"))
          Producto = Sin_Signos_Especiales(.fields("Producto"))
          CodigoInv = .fields("Codigo_Inv")
          NoMes = .fields("Num_Mes")
          Mes = Format$(NoMes, "00")
          Periodo = .fields("Periodo")
          Codigo2 = Periodo & " " & Mes & " " & CodigoInv & String(20 - Len(CodigoInv), " ") & " " & GrupoNo & String(10 - Len(GrupoNo), " ") & " " & Producto
          
          Progreso_Barra.Mensaje_Box = NombreCliente
          Progreso_Esperar
         'MsgBox NombreCliente
          Factura_No = Val(.fields("Periodo") & Format$(.fields("Num_Mes"), "00"))
          Total = .fields("Valor_Cobro")
          Saldo = .fields("Valor_Cobro") * 100
          CodigoP = .fields("CI_RUC")
          CodigoC = CStr(Val(.fields("CI_RUC")))
          CodigoC = CodigoC & String(Abs(4 - Len(CodigoC)), " ")
          DireccionCli = Sin_Signos_Especiales(.fields("Direccion"))
          Codigo3 = SinEspaciosDer(DireccionCli)
          DireccionCli = TrimStrg(MidStrg(DireccionCli, 1, Len(DireccionCli) - Len(Codigo3)))
          Codigo3 = TrimStrg(SinEspaciosDer(DireccionCli))
          Codigo1 = Format$(MidStrg(GrupoNo, 1, 1), "00")
          
          CadAux = "Pensión Acumulada"
          If Len(.fields("Email2")) > 3 Then CadAux = CadAux & ";" & .fields("Email2")
          
          If Tipo_Carga = 2 Then
             Codigo4 = "RECAUDACIONES"
          Else
             If CheqMatricula.value = 1 Then
                Codigo4 = "MATRICULAS Y PENSION DE " & MidStrg(MesesLetras(Month(MBFechaI)), 1, 3) & "-" & Year(MBFechaI)
             Else
                If Tipo_Carga = 3 Then
                   Codigo4 = "PENSION ACUMULADA"
                Else
                   Codigo4 = "PENSION ACUMULADA DE " & MidStrg(MesesLetras(Month(.fields("Fecha"))), 1, 3) & "-" & Year(.fields("Fecha"))
                End If
             End If
          End If
          If Len(.fields("Email2")) > 3 Then Codigo4 = Codigo4 & "|" & .fields("Email2") & ";"
                   
         'MsgBox Tipo_Carga & vbCrLf & Cta_Bancaria
          If Len(.fields("Actividad")) < 3 Then
             Contador = Contador + 1
            'Linea de texto separados por TAB para archivo plano
             Print #NumFileFacturas, "CO" & vbTab;                                                           '1
             Print #NumFileFacturas, Cta_Banco & vbTab;                                                      '2
             Print #NumFileFacturas, Contador & vbTab;                                                       '3
             Print #NumFileFacturas, Format$(Factura_No, "0000000000") & vbTab;                              '4
             Print #NumFileFacturas, CodigoP & vbTab;                                                        '5
             Print #NumFileFacturas, "USD" & vbTab;                                                          '6
             Print #NumFileFacturas, Format$(Saldo, "0000000000000") & vbTab;                                '7
             Print #NumFileFacturas, "REC" & vbTab;                                                          '8
             Print #NumFileFacturas, "32" & vbTab;                                                           '9
             Print #NumFileFacturas, vbTab;                                                                  '10 -> CTE/AHO
             Print #NumFileFacturas, "0" & vbTab;                                                            '11 -> No. Cta Cte/Aho
             Print #NumFileFacturas, .fields("TD_R") & vbTab;                                                    '12
             Print #NumFileFacturas, .fields("CI_RUC_R") & vbTab;                         '13
             Print #NumFileFacturas, MidStrg(Month(.fields("Fecha")) & " " & NombreCliente, 1, 40) & vbTab;  '14
             Print #NumFileFacturas, vbTab;                                                                  '15
             Print #NumFileFacturas, vbTab;                                                                  '16
             Print #NumFileFacturas, vbTab;                                                                  '17
             Print #NumFileFacturas, vbTab;                                                                  '18
             Print #NumFileFacturas, Codigo2 & vbTab;                                                        '19
             Print #NumFileFacturas, Codigo4 & vbTab                                                         '20 "Pensión Acumulada"
                
            'Linea en excel
             Obj_Hoja.cells(Contador, 1).formula = "CO"
             Obj_Hoja.cells(Contador, 2).formula = Cta_Banco
             Obj_Hoja.cells(Contador, 3).formula = Contador
             Obj_Hoja.cells(Contador, 4).formula = "'" & Format$(Factura_No, "0000000000")
             Obj_Hoja.cells(Contador, 5).formula = "'" & CodigoP
             Obj_Hoja.cells(Contador, 6).formula = "USD"
             Obj_Hoja.cells(Contador, 7).formula = "'" & Format$(Saldo, "0000000000000")
             Obj_Hoja.cells(Contador, 8).formula = "REC"
             Obj_Hoja.cells(Contador, 9).formula = "32"
             Obj_Hoja.cells(Contador, 10).formula = "0"  'CTE/AHO
             Obj_Hoja.cells(Contador, 11).formula = "0"  'No. Cta Cte/Aho
             Obj_Hoja.cells(Contador, 12).formula = .fields("TD_R")
             Obj_Hoja.cells(Contador, 13).formula = "'" & .fields("CI_RUC_R")
             Obj_Hoja.cells(Contador, 14).formula = "'" & MidStrg(Month(.fields("Fecha")) & " " & NombreCliente, 1, 40)
             Obj_Hoja.cells(Contador, 15).formula = "0"
             Obj_Hoja.cells(Contador, 16).formula = "0"
             Obj_Hoja.cells(Contador, 17).formula = "0"
             Obj_Hoja.cells(Contador, 18).formula = "0"
             Obj_Hoja.cells(Contador, 19).formula = Codigo2
             Obj_Hoja.cells(Contador, 20).formula = Codigo4 'CadAux
          End If
         .MoveNext
       Loop
   End If
  End With
  Close #NumFileFacturas
  
  Obj_Hoja.Columns("A:Z").AutoFit
  Obj_Libro.SaveAs ArchivoExcel
  obj_Excel.Quit
  
  Set Obj_Hoja = Nothing
  Set Obj_Libro = Nothing
  Set obj_Excel = Nothing
  
  ArchivoTexto = UCaseStrg(RutaSysBases & "\BANCO\FACTURAS\Internacional Debitos " & Fecha_Meses & ".TXT")
  ArchivoExcel2 = UCaseStrg(RutaSysBases & "\BANCO\FACTURAS\Internacional Debitos " & Fecha_Meses & ".XLSX")   ' xls
  NombreArchivos = NombreArchivos & ArchivoTexto & vbCrLf & vbCrLf
  
 'Genera un archvo de texto
  NumFileFacturas = FreeFile
  Open ArchivoTexto For Output As #NumFileFacturas ' Abre el archivo.
  
 'Genera un nuevo libro de Excel
  Set obj_Excel = CreateObject("Excel.Application")
  Set Obj_Libro = obj_Excel.Workbooks.Add
  Set Obj_Hoja = Obj_Libro.Worksheets(1)
  obj_Excel.ActiveSheet.Select ' La selecciona
  obj_Excel.ActiveSheet.Name = "Internacional Debitos"
   
' Comenzamos a generar el archivo: Internacional Debitos.TXT
  Contador = 0
  Mes = Month(MBFechaI)
  Anio = Val(MidStrg(Format$(Year(MBFechaI), "0000"), 2, 3))
  Dia = "15"
  
  With AdoFactura.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       Do While Not .EOF
          CodigoCli = .fields("Codigo")
          GrupoNo = .fields("Grupo")
          NombreCliente = Sin_Signos_Especiales(.fields("Cliente"))
          Producto = Sin_Signos_Especiales(.fields("Producto"))
          CodigoInv = .fields("Codigo_Inv")
          NoMes = .fields("Num_Mes")
          Mes = Format$(NoMes, "00")
          Periodo = .fields("Periodo")
          Codigo2 = Periodo & " " & Mes & " " & CodigoInv & String(20 - Len(CodigoInv), " ") & " " & GrupoNo & String(10 - Len(GrupoNo), " ") & " " & Producto

          Progreso_Barra.Mensaje_Box = NombreCliente
          Progreso_Esperar
          'TRep = Leer_Datos_Clientes(CodigoCli)
          TipoCta = Ninguno
          Select Case .fields("Tipo_Cta")
            Case "AHORROS":   TipoCta = "AHO"
            Case "CORRIENTE": TipoCta = "CTE"
          End Select
          If .fields("Cod_Banco") > 0 And Len(.fields("Cta_Numero")) > 1 And .fields("Cod_Banco") = Val(TxtCodBanco) And TipoCta <> Ninguno Then
             Contador = Contador + 1
             Factura_No = Val(.fields("Periodo") & Format$(.fields("Num_Mes"), "00"))
             Total = .fields("Valor_Cobro")
             Saldo = .fields("Valor_Cobro") * 100
             CodigoP = .fields("CI_RUC")
             DireccionCli = Sin_Signos_Especiales(.fields("Direccion"))
             Codigo3 = SinEspaciosDer(DireccionCli)
             DireccionCli = TrimStrg(MidStrg(DireccionCli, 1, Len(DireccionCli) - Len(Codigo3)))
             Codigo3 = TrimStrg(SinEspaciosDer(DireccionCli))
             Codigo1 = Format$(MidStrg(GrupoNo, 1, 1), "00")
             
             If Tipo_Carga = 2 Then
                Codigo4 = "DEBITOS AUTOMATICOS"
             Else
                If CheqMatricula.value = 1 Then
                   Codigo4 = "MATRICULAS Y PENSION DE " & MidStrg(MesesLetras(Month(MBFechaI)), 1, 3) & "-" & Year(MBFechaI)
                Else
                   If Tipo_Carga = 3 Then
                      Codigo4 = "PENSION ACUMULADA"
                   Else
                      Codigo4 = "PENSION ACUMULADA DE " & MidStrg(MesesLetras(Month(.fields("Fecha"))), 1, 3) & "-" & Year(.fields("Fecha"))
                   End If
                End If
             End If
             If Len(.fields("Email2")) > 3 Then Codigo4 = Codigo4 & "|" & .fields("Email2") & ";"
             
             Print #NumFileFacturas, "CO" & vbTab;                                                          '1
             Print #NumFileFacturas, Cta_Banco & vbTab;                                                     '2
             Print #NumFileFacturas, Contador & vbTab;                                                      '3
             Print #NumFileFacturas, Format$(Factura_No, "0000000000") & vbTab;                             '4
             Print #NumFileFacturas, CodigoP & vbTab;                                                       '5
             Print #NumFileFacturas, "USD" & vbTab;                                                         '6
             Print #NumFileFacturas, Format$(Saldo, "0000000000000") & vbTab;                               '7
             Print #NumFileFacturas, "CTA" & vbTab;                                                         '8
             Print #NumFileFacturas, .fields("Cod_Banco") & vbTab;                                          '9
             Print #NumFileFacturas, TipoCta & vbTab;                                                       '10
             Print #NumFileFacturas, .fields("Cta_Numero") & vbTab;                                         '11 -> No. Cta Cte/Aho
             Print #NumFileFacturas, .fields("TD_R") & vbTab;                                               '12
             Print #NumFileFacturas, .fields("CI_RUC_R") & vbTab;                                           '13
             Print #NumFileFacturas, MidStrg(Month(.fields("Fecha")) & " " & NombreCliente, 1, 40) & vbTab; '14
             Print #NumFileFacturas, vbTab;                                                                 '15
             Print #NumFileFacturas, vbTab;                                                                 '16
             Print #NumFileFacturas, vbTab;                                                                 '17
             Print #NumFileFacturas, vbTab;                                                                 '18
             Print #NumFileFacturas, Codigo2 ' Month(.Fields("Fecha")) & vbTab;                             '19
             Print #NumFileFacturas, Codigo4 ' "Pensión Acumulada" & vbTab;                                 '20
             Print #NumFileFacturas, Format$(Saldo, "0000000000000") & vbTab;                               '21
             Print #NumFileFacturas, "0" & vbTab;                                                           '22
             Print #NumFileFacturas, "0" & vbTab;                                                           '23
             Print #NumFileFacturas, "0" & vbTab;                                                           '24
             Print #NumFileFacturas, "0" & vbTab;                                                           '25
             Print #NumFileFacturas, "0" & vbTab;                                                           '26
             Print #NumFileFacturas, "0" & vbTab                                                            '27
             
            'Linea en excel
             Obj_Hoja.cells(Contador, 1).formula = "CO"
             Obj_Hoja.cells(Contador, 2).formula = Cta_Banco
             Obj_Hoja.cells(Contador, 3).formula = Contador
             Obj_Hoja.cells(Contador, 4).formula = "'" & Format$(Factura_No, "0000000000")
             Obj_Hoja.cells(Contador, 5).formula = "'" & CodigoP
             Obj_Hoja.cells(Contador, 6).formula = "USD"
             Obj_Hoja.cells(Contador, 7).formula = "'" & Format$(Saldo, "0000000000000")
             Obj_Hoja.cells(Contador, 8).formula = "CTA"
             Obj_Hoja.cells(Contador, 9).formula = .fields("Cod_Banco")
             Obj_Hoja.cells(Contador, 10).formula = TipoCta  'CTE/AHO
             Obj_Hoja.cells(Contador, 11).formula = .fields("Cta_Numero")  'No. Cta Cte/Aho
             Obj_Hoja.cells(Contador, 12).formula = .fields("TD_R")
             Obj_Hoja.cells(Contador, 13).formula = "'" & .fields("CI_RUC_R")
             Obj_Hoja.cells(Contador, 14).formula = "'" & MidStrg(Month(.fields("Fecha")) & " " & NombreCliente, 1, 40)
             Obj_Hoja.cells(Contador, 15).formula = "0"
             Obj_Hoja.cells(Contador, 16).formula = "0"
             Obj_Hoja.cells(Contador, 17).formula = "0"
             Obj_Hoja.cells(Contador, 18).formula = "0"
             Obj_Hoja.cells(Contador, 19).formula = Codigo2
             Obj_Hoja.cells(Contador, 20).formula = Codigo4 'CadAux
             Obj_Hoja.cells(Contador, 21).formula = "'" & Format$(Saldo, "0000000000000")
             Obj_Hoja.cells(Contador, 22).formula = "0"
             Obj_Hoja.cells(Contador, 23).formula = "0"
             Obj_Hoja.cells(Contador, 24).formula = "0"
             Obj_Hoja.cells(Contador, 25).formula = "0"
             Obj_Hoja.cells(Contador, 26).formula = "0"
             Obj_Hoja.cells(Contador, 27).formula = "0"
             Total_Banco = Total_Banco + Total
          End If
         .MoveNext
       Loop
   End If
  End With
  Close #NumFileFacturas
  
  Obj_Hoja.Columns("A:AZ").AutoFit
  Obj_Libro.SaveAs ArchivoExcel2
  obj_Excel.Quit
   
  Set Obj_Hoja = Nothing
  Set Obj_Libro = Nothing
  Set obj_Excel = Nothing
  
  NombreArchivos = NombreArchivos & ArchivoExcel & vbCrLf & vbCrLf
  NombreArchivos = NombreArchivos & ArchivoExcel2 & vbCrLf & vbCrLf
  RatonNormal
  
  Progreso_Final
 'FBancoPichincha.Caption = "FACTURACION DE BANCO INTERNACIONAL"
  LabelAbonos.Caption = Format$(Total_Banco, "#,##0.00")
  MsgBox "SE GENERARON LOS SIGUIENTES ARCHIVOS:" & vbCrLf & vbCrLf & NombreArchivos
  Abrir_Excel ArchivoExcel
  Abrir_Excel ArchivoExcel2
End Sub

Public Sub Generar_Otros_Bancos()
Dim obj_Excel As Object
Dim Obj_Libro As Object
Dim Obj_Hoja As Object
Dim ArchivoTexto As String
Dim ArchivoExcel As String
Dim NombreArchivos As String

Dim AuxNumEmp As String
Dim DiaV As Integer
Dim MesV As Integer
Dim AñoV As Integer
Dim Fecha_Meses As String
Dim CamposFile() As Campos_Tabla
Dim Total_Banco As Currency

'Dim TRep As Tipo_Beneficiarios
  Progreso_Barra.Mensaje_Box = "Generando los Archivos al banco"
  Progreso_Iniciar
  Total_Banco = 0
  Mifecha = BuscarFecha(MBFechaI)
  MiMes = Format$(Month(MBFechaI), "00")
  FechaFin = BuscarFecha(UltimoDiaMes(MBFechaI))
  TextoImprimio = ""
  AuxNumEmp = NumEmpresa
  Cta_Cobrar = Ninguno
  FechaTexto = FechaSistema
' Comenzamos a generar el archivo: Internacional Recaudaciones.TXT
  EsComa = True
  Estab = False
  Contador = 0
  Factura_No = 0
  NumFileFacturas = FreeFile
  Fecha_Meses = MBFechaI & " al " & MBFechaF
  Fecha_Meses = Replace(Fecha_Meses, "/", "-")
' Comenzamos a generar el archivo: Internacional Debitos.TXT
  Contador = 0
  Mes = Month(MBFechaI)
  Anio = Val(MidStrg(Format$(Year(MBFechaI), "0000"), 2, 3))
  Dia = "15"
  
  ArchivoTexto = UCaseStrg(RutaSysBases & "\BANCO\FACTURAS\Otros Bancos Debitos " & Fecha_Meses & ".TXT")
  ArchivoExcel = UCaseStrg(RutaSysBases & "\BANCO\FACTURAS\Otros Bancos Debitos " & Fecha_Meses & ".XLSX")   ' xls
  NombreArchivos = ArchivoTexto & vbCrLf & vbCrLf & ArchivoExcel
  
 'Genera un archvo de texto
  NumFileFacturas = FreeFile
  Open ArchivoTexto For Output As #NumFileFacturas ' Abre el archivo.
  
 'Genera un nuevo libro de Excel
  Set obj_Excel = CreateObject("Excel.Application")
  Set Obj_Libro = obj_Excel.Workbooks.Add
  Set Obj_Hoja = Obj_Libro.Worksheets(1)
  obj_Excel.ActiveSheet.Select ' La selecciona
  obj_Excel.ActiveSheet.Name = "Otros Bancos Debitos"
  
  With AdoFactura.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       Do While Not .EOF
          CodigoCli = .fields("Codigo")
          GrupoNo = .fields("Grupo")
          NombreCliente = Sin_Signos_Especiales(.fields("Cliente"))
          Producto = Sin_Signos_Especiales(.fields("Producto"))
          CodigoInv = .fields("Codigo_Inv")
          NoMes = .fields("Num_Mes")
          Mes = Format$(NoMes, "00")
          Periodo = .fields("Periodo")
          Codigo2 = Periodo & " " & Mes & " " & CodigoInv & String(20 - Len(CodigoInv), " ") & " " & GrupoNo & String(10 - Len(GrupoNo), " ") & " " & Producto
          
          Progreso_Barra.Mensaje_Box = NombreCliente
          Progreso_Esperar
          'TRep = Leer_Datos_Clientes(CodigoCli)
          TipoCta = Ninguno
          Select Case .fields("Tipo_Cta")
            Case "AHORROS":   TipoCta = "AHO"
            Case "CORRIENTE": TipoCta = "CTE"
          End Select
          If .fields("Cod_Banco") > 0 And Len(.fields("Cta_Numero")) > 1 And TipoCta <> Ninguno Then
             If .fields("Cod_Banco") <> Val(TxtCodBanco) Then
                Contador = Contador + 1
                Factura_No = Val(.fields("Periodo") & Format$(.fields("Num_Mes"), "00"))
                Total = .fields("Valor_Cobro")
                Saldo = .fields("Valor_Cobro") * 100
                CodigoP = .fields("CI_RUC")
                DireccionCli = Sin_Signos_Especiales(.fields("Direccion"))
                Codigo3 = SinEspaciosDer(DireccionCli)
                DireccionCli = TrimStrg(MidStrg(DireccionCli, 1, Len(DireccionCli) - Len(Codigo3)))
                Codigo3 = TrimStrg(SinEspaciosDer(DireccionCli))
                Codigo1 = Format$(MidStrg(GrupoNo, 1, 1), "00")
                
                If Tipo_Carga = 2 Then
                   Codigo4 = "DEBITOS AUTOMATICOS"
                Else
                   If CheqMatricula.value = 1 Then
                      Codigo4 = "MATRICULAS Y PENSION DE " & MidStrg(MesesLetras(Month(MBFechaI)), 1, 3) & "-" & Year(MBFechaI)
                   Else
                      If Tipo_Carga = 3 Then
                         Codigo4 = "PENSION ACUMULADA"
                      Else
                         Codigo4 = "PENSION ACUMULADA DE " & MidStrg(MesesLetras(Month(.fields("Fecha"))), 1, 3) & "-" & Year(.fields("Fecha"))
                      End If
                   End If
                End If
                If Len(.fields("Email2")) > 3 Then Codigo4 = Codigo4 & "|" & .fields("Email2") & ";"
          
                Print #NumFileFacturas, "CO" & vbTab;                                                           '1
                Print #NumFileFacturas, Cta_Banco & vbTab;                                                      '2
                Print #NumFileFacturas, Contador & vbTab;                                                       '3
                Print #NumFileFacturas, Format$(Factura_No, "0000000000") & vbTab;                              '4
                Print #NumFileFacturas, CodigoP & vbTab;                                                        '5
                Print #NumFileFacturas, "USD" & vbTab;                                                          '6
                Print #NumFileFacturas, Format$(Saldo, "0000000000000") & vbTab;                                '7
                Print #NumFileFacturas, "CTA" & vbTab;                                                          '8
                Print #NumFileFacturas, .fields("Cod_Banco") & vbTab;                                           '9
                Print #NumFileFacturas, TipoCta & vbTab;                                                        '10
                Print #NumFileFacturas, .fields("Cta_Numero") & vbTab;                                          '11 -> No. Cta Cte/Aho
                Print #NumFileFacturas, .fields("TD_R") & vbTab;                                                '12
                Print #NumFileFacturas, Format$(.fields("CI_RUC_R"), "0000000000") & vbTab;                     '13
                Print #NumFileFacturas, MidStrg(Month(.fields("Fecha")) & " " & NombreCliente, 1, 40) & vbTab;  '14
                Print #NumFileFacturas, vbTab;                                                                  '15
                Print #NumFileFacturas, vbTab;                                                                  '16
                Print #NumFileFacturas, vbTab;                                                                  '17
                Print #NumFileFacturas, vbTab;                                                                  '18
                Print #NumFileFacturas, Codigo2 & vbTab;                                                        '19
                Print #NumFileFacturas, Codigo4 & vbTab;                                                        '20
                Print #NumFileFacturas, Format$(Saldo, "0000000000000") & vbTab                                 '21
                
                Obj_Hoja.cells(Contador, 1).formula = "CO"
                Obj_Hoja.cells(Contador, 2).formula = Cta_Banco
                Obj_Hoja.cells(Contador, 3).formula = Contador
                Obj_Hoja.cells(Contador, 4).formula = "'" & Format$(Factura_No, "0000000000")
                Obj_Hoja.cells(Contador, 5).formula = "'" & CodigoP
                Obj_Hoja.cells(Contador, 6).formula = "USD"
                Obj_Hoja.cells(Contador, 7).formula = "'" & Format$(Saldo, "0000000000000")
                Obj_Hoja.cells(Contador, 8).formula = "CTA"
                Obj_Hoja.cells(Contador, 9).formula = .fields("Cod_Banco")
                Obj_Hoja.cells(Contador, 10).formula = TipoCta
                Obj_Hoja.cells(Contador, 11).formula = "'" & .fields("Cta_Numero")
                Obj_Hoja.cells(Contador, 12).formula = .fields("TD_R")
                Obj_Hoja.cells(Contador, 13).formula = "'" & Format$(.fields("CI_RUC_R"), "0000000000")
                Obj_Hoja.cells(Contador, 14).formula = "'" & MidStrg(Month(.fields("Fecha")) & " " & NombreCliente, 1, 40)
                Obj_Hoja.cells(Contador, 15).formula = "0"
                Obj_Hoja.cells(Contador, 16).formula = "0"
                Obj_Hoja.cells(Contador, 17).formula = "0"
                Obj_Hoja.cells(Contador, 18).formula = "0"
                Obj_Hoja.cells(Contador, 19).formula = Codigo2 ' "'" & Month(.Fields("Fecha"))
                Obj_Hoja.cells(Contador, 20).formula = Codigo4 ' CadAux
                Obj_Hoja.cells(Contador, 21).formula = "'" & Format$(Saldo, "0000000000000")
                
                Total_Banco = Total_Banco + Total
             End If
          End If
         .MoveNext
       Loop
   End If
  End With
  Close #NumFileFacturas
  RatonNormal
  
  Obj_Hoja.Columns("A:Z").AutoFit
  Obj_Libro.SaveAs ArchivoExcel
  obj_Excel.Quit
   
  Set Obj_Hoja = Nothing
  Set Obj_Libro = Nothing
  Set obj_Excel = Nothing
  
  Progreso_Final
 'FBancoPichincha.Caption = "FACTURACION DE BANCO INTERNACIONAL"
  LabelAbonos.Caption = Format$(Total_Banco, "#,##0.00")
  MsgBox "SE GENERARON LOS SIGUIENTES ARCHIVOS:" & vbCrLf & vbCrLf & UCaseStrg(NombreArchivos)
  Abrir_Excel ArchivoExcel
End Sub

Public Sub Generar_Tarjetas()
Dim AuxNumEmp As String
Dim Fecha_Meses As String
Dim Total_Banco As Currency
Dim NombreArchivos As String
Dim ContadorTarjeta As Integer
Dim TRep As Tipo_Beneficiarios
  Progreso_Barra.Mensaje_Box = "Generando los Archivos al banco"
  Progreso_Iniciar
  Total_Banco = 0
  TextoImprimio = ""
  Fecha_Meses = MBFechaI & " al " & MBFechaF
  Fecha_Meses = Replace(Fecha_Meses, "/", "-")
' Comenzamos a generar el archivo: Internacional Debitos.TXT
  ContadorTarjeta = 0
  NumFileFacturas = FreeFile
  RutaGeneraFile = UCaseStrg(RutaSysBases & "\BANCO\FACTURAS\Tarjetas " & Fecha_Meses & ".TXT")
  NombreArchivos = RutaGeneraFile
  Open RutaGeneraFile For Output As #NumFileFacturas ' Abre el archivo.
  With AdoFactura.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       Progreso_Barra.Valor_Maximo = Progreso_Barra.Valor_Maximo + .RecordCount
       Do While Not .EOF
          Progreso_Barra.Mensaje_Box = "Buscando Datos"
          CodigoCli = .fields("Codigo")
          CodigoP = Format$(.fields("CI_RUC"), "00000000000000000")
          TRep = Leer_Datos_Clientes(CodigoCli)
          If .fields("Tipo_Cta") = "TARJETA" And Len(TRep.Cta_Numero) >= 14 Then
             'MsgBox TRep.Cta_Numero & " - " & Len(TRep.Cta_Numero)
             ContadorTarjeta = ContadorTarjeta + 1
             Progreso_Barra.Mensaje_Box = Format(ContadorTarjeta, "000") & " " & .fields("Cliente")
             Total = .fields("Valor_Cobro")
             Saldo = .fields("Valor_Cobro") * 100
             TRep.Cta_Numero = MidStrg(TRep.Cta_Numero, 1, 16)
             TRep.Cta_Numero = String(16 - Len(TRep.Cta_Numero), "0") & TRep.Cta_Numero
             
             Print #NumFileFacturas, TRep.Cta_Numero;
             Print #NumFileFacturas, Format(Val(TxtCodBanco), "00000000");
             Print #NumFileFacturas, Format(FechaSistema, "YYYYMMDD");
             Print #NumFileFacturas, Format$(Saldo, "00000000000000000");
             Print #NumFileFacturas, "00000000000000000";  ' CodigoP
             Print #NumFileFacturas, "202";
             Print #NumFileFacturas, "000000";
             Print #NumFileFacturas, "00";
            'Transaccion/Secuencial
            'MsgBox TRep.CI_RUC & " - " & MidStrg(TRep.CI_RUC, 5, 6)
             Print #NumFileFacturas, MidStrg(TRep.CI_RUC, 7, 4);              ' Codigo Alumno
'''             Print #NumFileFacturas, MidStrg(.Fields("Periodo"), 4, 1);    ' Año
'''             Print #NumFileFacturas, Format$(.Fields("Num_Mes"), "00");    ' Mes
             Print #NumFileFacturas, Format$(ContadorTarjeta, "00");      ' Secuencial
             Print #NumFileFacturas, "439473";
             Print #NumFileFacturas, Format(TRep.Fecha_Cad, "YYYYMM");
             Print #NumFileFacturas, "00000000000000000";   ' IVA
             Print #NumFileFacturas, "00D00000";
             Print #NumFileFacturas, "000000000000000";
             Print #NumFileFacturas, Format$(Saldo, "000000000000000")
             Total_Banco = Total_Banco + Total
          End If
          Progreso_Esperar
         .MoveNext
       Loop
   End If
  End With
  Close #NumFileFacturas
  RatonNormal
  Progreso_Final
 'FBancoPichincha.Caption = "FACTURACION DE BANCO INTERNACIONAL"
  LabelAbonos.Caption = Format$(Total_Banco, "#,##0.00")
  MsgBox "SE GENERO EL SIGUIENTE ARCHIVO:" & vbCrLf _
        & UCaseStrg(NombreArchivos)
End Sub

Public Sub Generar_Produbanco()
Dim AuxNumEmp As String
Dim DiaV As Integer
Dim MesV As Integer
Dim AñoV As Integer
Dim Fecha_Meses As String
Dim ValorStr As String
Dim CamposFile() As Campos_Tabla
Dim Total_Banco As Currency
    
    sSQL = "SELECT Codigo,GrupoNo,Periodo,Mes,Num_Mes,Fecha " _
         & "FROM Clientes_Facturacion " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Fecha BETWEEN #" & BuscarFecha(MBFechaI) & "# and #" & BuscarFecha(MBFechaF) & "# "
    If CheqRangos.value <> 0 Then sSQL = sSQL & "AND GrupoNo BETWEEN '" & DCGrupoI & "' and '" & DCGrupoF & "' "
    sSQL = sSQL _
         & "ORDER BY Codigo,Fecha "
    Select_Adodc AdoAux, sSQL

  Total_Banco = 0
  Mifecha = BuscarFecha(MBFechaI)
  MiMes = Format$(Month(MBFechaI), "00")
  FechaFin = BuscarFecha(UltimoDiaMes(MBFechaI))
  TextoImprimio = ""
  AuxNumEmp = NumEmpresa
  Cta_Cobrar = Ninguno
  FechaTexto = FechaSistema
' Comenzamos a generar el archivo: SCRECXX.TXT
  EsComa = True
  Estab = False
  Contador = 0
  Factura_No = 0
  NumFileFacturas = FreeFile
  Fecha_Meses = MBFechaI & " al " & MBFechaF
  Fecha_Meses = Replace(Fecha_Meses, "/", "-")
  RutaGeneraFile = UCaseStrg(RutaSysBases & "\BANCO\FACTURAS\RECAUDACION " & Fecha_Meses & ".TXT")
  Open RutaGeneraFile For Output As #NumFileFacturas ' Abre el archivo.
  With AdoFactura.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       Do While Not .EOF
          Contador = Contador + 1
          CodigoCli = .fields("CodigoC")
          NombreCliente = TrimStrg(MidStrg(Sin_Signos_Especiales(.fields("Cliente")), 1, 40))
         'MsgBox NombreCliente
          Factura_No = Factura_No + 1
          Total = .fields("Valor_Cobro")
          Saldo = .fields("Valor_Cobro") * 100
          ValorStr = CStr(Saldo)
          ValorStr = String(13 - Len(ValorStr), "0") & ValorStr
         'MsgBox ValorStr
          CodigoP = .fields("CI_RUC")
          CodigoC = CStr(Val(.fields("CI_RUC")))
          CodigoC = CodigoC & String(Abs(4 - Len(CodigoC)), " ")
          DireccionCli = Sin_Signos_Especiales(.fields("Direccion"))
          Codigo3 = SinEspaciosDer(DireccionCli)
          DireccionCli = TrimStrg(MidStrg(DireccionCli, 1, Len(DireccionCli) - Len(Codigo3)))
          Codigo3 = TrimStrg(SinEspaciosDer(DireccionCli))
          Codigo1 = Format$(MidStrg(GrupoNo, 1, 1), "00")
          If Len(Cta_Bancaria) < 11 Then Cta_Bancaria = String(11 - Len(Cta_Bancaria), "0") & Cta_Bancaria
          If Tipo_Carga = 3 Then
             Codigo4 = "Transporte " & .fields("Grupo")
          Else
             Codigo4 = "Transporte " & .fields("Grupo") & " De: " & .fields("Periodo") & "-" & MesesLetras(.fields("Num_Mes"))
          End If
''          If AdoAux.Recordset.RecordCount > 0 Then
''             AdoAux.Recordset.MoveFirst
''             AdoAux.Recordset.Find ("Codigo = '" & CodigoCli & "' ")
''             Do While Not AdoAux.Recordset.EOF
''                If AdoAux.Recordset.Fields("Codigo") = CodigoCli Then
''                   Codigo4 = Codigo4 _
''                           & AdoAux.Recordset.Fields("Periodo") & "-" & AdoAux.Recordset.Fields("Mes") & ","
''                Else
''                   AdoAux.Recordset.MoveLast
''                End If
''                AdoAux.Recordset.MoveNext
''             Loop
''             Codigo4 = MidStrg(Codigo4, 1, Len(Codigo4) - 1)
''          End If
          Print #NumFileFacturas, "CO" & vbTab;            '01
          Print #NumFileFacturas, Cta_Bancaria & vbTab;    '02
          Print #NumFileFacturas, Contador & vbTab;        '03
          Print #NumFileFacturas, vbTab;                   '04
          Print #NumFileFacturas, CodigoP & vbTab;         '05
          Print #NumFileFacturas, "USD" & vbTab;           '06
          Print #NumFileFacturas, ValorStr & vbTab;        '07
          Print #NumFileFacturas, "REC" & vbTab;           '08
          Print #NumFileFacturas, CodigoDelBanco & vbTab;  '09
          Print #NumFileFacturas, vbTab;                   '10
          Print #NumFileFacturas, vbTab;                   '11
          Print #NumFileFacturas, "C" & vbTab;             '12
          Print #NumFileFacturas, CodigoP & vbTab;         '13
          Print #NumFileFacturas, NombreCliente & vbTab;   '14
          Print #NumFileFacturas, vbTab;                   '15
          Print #NumFileFacturas, vbTab;                   '16
          Print #NumFileFacturas, vbTab;                   '17
          Print #NumFileFacturas, vbTab;                   '18
          Print #NumFileFacturas, Codigo4 & vbTab;         '19
          Print #NumFileFacturas, vbTab;                   '20
          Print #NumFileFacturas, vbTab                    '21
         .MoveNext
       Loop
   End If
  End With
  Close #NumFileFacturas
  RatonNormal
  'ProgBarra.value = ProgBarra.Max
 'FBancoPichincha.Caption = "FACTURACION DE BANCO INTERNACIONAL"
  LabelAbonos.Caption = Format$(Total_Banco, "#,##0.00")
  MsgBox "SE HA GENERADO EL SIGUIENTE ARCHIVO:" & vbCrLf & vbCrLf _
       & RutaGeneraFile & vbCrLf
End Sub

Private Sub TextFacturaNo_GotFocus()
  MarcarTexto TextFacturaNo
End Sub

Private Sub TextFacturaNo_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextFacturaNo_LostFocus()
 Factura_No = Val(TextFacturaNo)
 FA.Factura = Val(TextFacturaNo)
 If Existe_Factura(FA) Then
    FA.Factura = TextFacturaNo
 Else
    FA.Nuevo_Doc = True
    FA.Factura = ReadSetDataNum(FA.TC & "_SERIE_" & FA.Serie, True, False)
    If Val(TextFacturaNo) < FA.Factura Then
       Mensajes = "La Factura/Nota de Venta No. " & TextFacturaNo & vbCrLf _
                & "No esta Procesada, Desea Procesarla"
       Titulo = "Formulario de Confirmación"
       If BoxMensaje = vbYes Then FA.Factura = TextFacturaNo
    Else
       FA.Factura = ReadSetDataNum(FA.TC & "_SERIE_" & FA.Serie, True, False)  'Numero_Factura(FA)
       TextFacturaNo = FA.Factura
    End If
 End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
  FA.Tipo_PRN = "FM"
  Cta_Bancaria = SinEspaciosDer(DCBanco)
  Select Case Button.key
    Case "Enviar"
         Enviar_Cobros
    Case "SubirAbonos"
         Subir_Abonos
    Case "GenerarFacturas"
         Generar_Facturas
    Case "AlumnosContabilidad"
         Alumnos_Contabilidad
    Case "RenumerarCodigos"
         Renumerar_Codigos
    Case "ImprimirCodigos"
         Imprimir_Codigos
    Case "Salir"
         Unload FRecaudacionBancosPreFa
  End Select
End Sub

Private Sub TxtCodBanco_GotFocus()
  MarcarTexto TxtCodBanco
End Sub

Private Sub TxtCodBanco_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtCodBanco_LostFocus()
  TextoValido TxtCodBanco
End Sub

Private Sub TxtFile_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  If CtrlDown And KeyCode = vbKeyP Then If TxtFile <> "" Then Imprimir_Texto_Lineas TxtFile
End Sub

Public Sub Generar_Guayaquil()
Dim FechaVenc As String
Dim MesesDif As Byte
RutaGeneraFile = UCaseStrg(RutaSysBases & "\BANCO\FACTURAS\REM_" & Format$(FechaSistema, "YYYYMMDD") & "_" & UCase(CodigoDelBanco) & ".TXT")
NumFileFacturas = FreeFile
TipoDoc = "0"
Contador = 0
FechaTexto = BuscarFecha(MBFechaI)
'MsgBox RutaGeneraFile
TxtFile = ""
Open RutaGeneraFile For Output As #NumFileFacturas ' Abre el archivo.
With AdoFactura.Recordset
 If .RecordCount > 0 Then
    .MoveFirst
     Contador = 0
     TotalIngreso = 0
     Do While Not .EOF
        Contador = Contador + 1
        TotalIngreso = TotalIngreso + .fields("Valor_Cobro")
       .MoveNext
     Loop
    .MoveFirst
     TxtFile = "TOTAL NOMINA DE RECAUDACION:" & vbCrLf
     Codigo3 = TrimStrg(MidStrg(NombreEmpresa, 1, 30))
     Total = 0
     
     Total_Factura = 0
     IE = 1
     JE = 1
     KE = 1
     
     I = Int(TotalIngreso)
     J = (TotalIngreso - I) * 100
     Grupo_No = .fields("Grupo")
     Codigo = .fields("CodigoC")
     
'''     MesesDif = Month(MBFechaV) - Month(MBFechaI)
'''     If MesesDif < 0 Then MesesDif = 0
'''     FechaVenc = CLongFecha(CFechaLong(MBFechaV) - MesesDif)
       'Registro de Cabecera
        Print #NumFileFacturas, "01";                                                   ' Localidad
        Print #NumFileFacturas, "REC";                                                  ' Codigo de Servicio
        Print #NumFileFacturas, "00017";                                                ' Codigo del Banco
        Print #NumFileFacturas, CodigoDelBanco & String(5 - Len(CodigoDelBanco), " ");  ' Codigo de la Empresa
        Print #NumFileFacturas, "01";                                                   ' Contenido del archivo
        Print #NumFileFacturas, Format$(MBFechaI, "YYYYMMDD");                          ' Fecha Inicio de Cobro
        Print #NumFileFacturas, Format$(MBFechaV, "YYYYMMDD");                          ' Fecha Tope de Cobro
        Print #NumFileFacturas, Format(Contador, "00000000");                           ' Numero de Registros
        Print #NumFileFacturas, Format$(I, "0000000000000") & Format$(J, "00");         ' Valor Total a Cobrar
        Print #NumFileFacturas, String(68, " ")                                         ' Fin de la Trama
        
     TotalIngreso = 0
     Do While Not .EOF
        'MsgBox .Fields("Grupo")
        If Grupo_No <> .fields("Grupo") Then
           Codigo4 = Format$(Total, "#,##0.00")
           Codigo4 = String(13 - Len(Codigo4), " ") & Format$(Total, "#,##0.00")
           TxtFile = TxtFile & "Grupo: " & Grupo_No & vbTab & "Resumen de Registros por Grupo: " & JE & vbTab & "Total a Recaudar USD" & vbTab & Codigo4 & vbCrLf
           JE = 0
           Total = 0
           IE = IE + 1
           Grupo_No = .fields("Grupo")
           Codigo = .fields("CodigoC")
           'FechaVenc = CLongFecha(CFechaLong(MBFechaV) - MesesDif)
        End If
        If Codigo <> .fields("CodigoC") Then
           JE = JE + 1
           KE = KE + 1
           Codigo = .fields("CodigoC")
           'FechaVenc = CLongFecha(CFechaLong(MBFechaV) - MesesDif)
        End If
        FRecaudacionBancosPreFa.Caption = .fields("Grupo") & " - " & Format$(Contador / .RecordCount, "00%")
        Contador = Contador + 1
        CodigoCli = CStr(Val(.fields("CI_RUC")))
        NombreCliente = Sin_Signos_Especiales(TrimStrg(MidStrg(.fields("Cliente"), 1, 40)))
        Codigo1 = TrimStrg(MidStrg(.fields("Direccion"), 1, 30))
        FechaTexto = " " & MidStrg(MesesLetras(Month(MBFechaI)), 1, 3) & " " & Year(MBFechaI)
       'Codigo2 = MidStrg(UCaseStrg(MidStrg(MesesLetras(Month(.fields("Fecha"))), 1, 3)) & " " & .fields("Grupo"), 1, 15)
        Codigo2 = TrimStrg(MidStrg(UCaseStrg(.fields("Codigo_Inv") & " " & .fields("Grupo")), 1, 15))
        Total_Factura = .fields("Valor_Cobro")
        If Costo_Banco > 0 Then Total_Factura = Total_Factura + Costo_Banco
        Total = Total + Total_Factura
        TotalIngreso = TotalIngreso + Total_Factura
        I = Int(Total_Factura)
        J = (Total_Factura - I) * 100
       'MsgBox Grupo_No & "(" & JE & ")" & vbCrLf & Total_Factura & vbCrLf & Total & vbCrLf & TotalIngreso
       'Empieza la trama por Alumno
        
       'Registro del Detalle del Cobro
        Print #NumFileFacturas, "02";                                                   ' Tipo Registro
        Print #NumFileFacturas, "01";                                                   ' Novedad
        Print #NumFileFacturas, CodigoCli & String(15 - Len(CodigoCli), " ");           ' Codigo del Alumno
        Print #NumFileFacturas, NombreCliente & String(40 - Len(NombreCliente), " ");   ' Nombre del Alumno
        Print #NumFileFacturas, Format$(I, "00000000") & Format$(J, "00");              ' Valor a Cobrar
        Print #NumFileFacturas, Format$(MBFechaV, "YYYYMMDD");                          ' Fecha Tope de Cobro
        Print #NumFileFacturas, Format$(I, "00000000") & Format$(J, "00");              ' Valor a Cobrar
        Print #NumFileFacturas, String(10, "0");                                        ' Valor Retenido
        Print #NumFileFacturas, Codigo2 & String(15 - Len(Codigo2), " ");               ' Referencia
        Print #NumFileFacturas, Format$(.fields("Fecha"), "YYYYMM");                    ' Mes de Generacion
        Print #NumFileFacturas, Format$(Contador, "00");                                ' Secuencial
        Print #NumFileFacturas, "    "                                                  ' Fin del a Trama
        
       .MoveNext
     Loop
     Codigo4 = Format$(Total, "#,##0.00")
     Codigo4 = String(13 - Len(Codigo4), " ") & Format$(Total, "#,##0.00")
     TxtFile = TxtFile _
             & "Grupo: " & Grupo_No & vbTab & "Resumen de Registros por Grupo: " & JE & vbTab & "Total a Recaudar USD" & vbTab & Codigo4 & vbCrLf
     Codigo4 = Format$(TotalIngreso, "#,##0.00")
     Codigo4 = String(13 - Len(Codigo4), " ") & Format$(TotalIngreso, "#,##0.00")
     TxtFile = TxtFile _
             & String(90, "-") & vbCrLf _
             & "Total Grupos: " & IE & vbTab & "Total Alumnos: " & KE & vbTab & vbTab & "Total a Recaudar USD" & vbTab & Codigo4 & vbCrLf
 End If
End With
Close #NumFileFacturas
RatonNormal
MsgBox UCaseStrg("Fin del Proceso " & vbCrLf & vbCrLf & "El Archivo se Generara en: " & vbCrLf & vbCrLf & RutaGeneraFile)
End Sub

