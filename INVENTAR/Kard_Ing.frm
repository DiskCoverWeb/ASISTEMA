VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.Ocx"
Begin VB.Form Kard_Ing 
   Caption         =   "CONTROL DE INVENTARIO PARA INGRESOS/EGRESOS"
   ClientHeight    =   8340
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13635
   Icon            =   "Kard_Ing.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Kard_Ing.frx":0442
   ScaleHeight     =   15615
   ScaleWidth      =   15960
   WindowState     =   2  'Maximized
   Begin VB.TextBox TxtSerie 
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
      IMEMode         =   3  'DISABLE
      Left            =   11340
      MaxLength       =   10
      MultiLine       =   -1  'True
      TabIndex        =   23
      Text            =   "Kard_Ing.frx":0784
      Top             =   1155
      Width           =   750
   End
   Begin VB.TextBox TxtSerieNo 
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
      IMEMode         =   3  'DISABLE
      Left            =   7665
      MaxLength       =   25
      TabIndex        =   59
      Text            =   "."
      Top             =   4305
      Width           =   3270
   End
   Begin VB.TextBox TxtProcedencia 
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
      IMEMode         =   3  'DISABLE
      Left            =   5040
      MaxLength       =   25
      TabIndex        =   57
      Text            =   "."
      Top             =   4305
      Width           =   2640
   End
   Begin VB.TextBox TxtModelo 
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
      IMEMode         =   3  'DISABLE
      Left            =   10920
      MaxLength       =   25
      TabIndex        =   55
      Text            =   "."
      Top             =   3675
      Width           =   4215
   End
   Begin VB.TextBox TxtRegSanitario 
      Enabled         =   0   'False
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
      IMEMode         =   3  'DISABLE
      Left            =   8925
      MaxLength       =   25
      TabIndex        =   53
      Text            =   "."
      Top             =   3675
      Width           =   2010
   End
   Begin VB.TextBox TxtLoteNo 
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
      IMEMode         =   3  'DISABLE
      Left            =   5040
      MaxLength       =   10
      MultiLine       =   -1  'True
      TabIndex        =   47
      Text            =   "Kard_Ing.frx":078D
      Top             =   3675
      Width           =   1380
   End
   Begin MSMask.MaskEdBox MBFechaExp 
      Height          =   330
      Left            =   7665
      TabIndex        =   51
      Top             =   3675
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
   Begin MSDataListLib.DataCombo DCMarca 
      Bindings        =   "Kard_Ing.frx":078F
      DataSource      =   "AdoMarca"
      Height          =   315
      Left            =   8190
      TabIndex        =   33
      Top             =   2205
      Width           =   3900
      _ExtentX        =   6879
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16761024
      Text            =   "DC"
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
   Begin MSMask.MaskEdBox MBVence 
      Height          =   330
      Left            =   1365
      TabIndex        =   10
      Top             =   735
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
   Begin MSMask.MaskEdBox MBFechaI 
      Height          =   330
      Left            =   1365
      TabIndex        =   8
      Top             =   420
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
   Begin VB.CheckBox CheqContraCta 
      Caption         =   "CONTRA CUENTA"
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
      Left            =   3780
      TabIndex        =   5
      Top             =   105
      Value           =   1  'Checked
      Width           =   2010
   End
   Begin VB.ComboBox CLTP 
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
      Left            =   630
      TabIndex        =   1
      Text            =   "CD"
      Top             =   105
      Width           =   750
   End
   Begin VB.TextBox TextTotal 
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
      Left            =   13440
      MultiLine       =   -1  'True
      TabIndex        =   65
      Text            =   "Kard_Ing.frx":07A6
      Top             =   4305
      Width           =   1695
   End
   Begin VB.TextBox TextDesc1 
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
      Left            =   12180
      MultiLine       =   -1  'True
      TabIndex        =   63
      Text            =   "Kard_Ing.frx":07A8
      Top             =   4305
      Width           =   1275
   End
   Begin VB.TextBox TextDesc 
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
      Left            =   10920
      MultiLine       =   -1  'True
      TabIndex        =   61
      Text            =   "Kard_Ing.frx":07AA
      Top             =   4305
      Width           =   1275
   End
   Begin MSDataListLib.DataCombo DCInv 
      Bindings        =   "Kard_Ing.frx":07AC
      DataSource      =   "AdoInv"
      Height          =   2715
      Left            =   105
      TabIndex        =   27
      Top             =   1890
      Width           =   4845
      _ExtentX        =   8546
      _ExtentY        =   4789
      _Version        =   393216
      Style           =   1
      BackColor       =   16744576
      ForeColor       =   16777215
      Text            =   "Productos"
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
   Begin VB.TextBox TxtCodBar 
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
      IMEMode         =   3  'DISABLE
      Left            =   10815
      MaxLength       =   25
      TabIndex        =   45
      Text            =   "."
      Top             =   2940
      Width           =   4320
   End
   Begin VB.CommandButton Command4 
      Height          =   645
      Left            =   15225
      Picture         =   "Kard_Ing.frx":07C1
      Style           =   1  'Graphical
      TabIndex        =   79
      Top             =   105
      Width           =   645
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Seleccionar Comprobante"
      Height          =   645
      Left            =   2730
      TabIndex        =   78
      Top             =   7875
      Width           =   1170
   End
   Begin MSDataListLib.DataCombo DCCtaObra 
      Bindings        =   "Kard_Ing.frx":1011
      DataSource      =   "AdoCtaObra"
      Height          =   330
      Left            =   5880
      TabIndex        =   6
      Top             =   105
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "Contra cuenta"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo DCDiario 
      Bindings        =   "Kard_Ing.frx":102A
      DataSource      =   "AdoDiario"
      Height          =   315
      Left            =   13335
      TabIndex        =   13
      Top             =   420
      Visible         =   0   'False
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   556
      _Version        =   393216
      Text            =   "Diario"
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
   Begin MSDataListLib.DataCombo DCBenef 
      Bindings        =   "Kard_Ing.frx":1042
      DataSource      =   "AdoBenef"
      Height          =   330
      Left            =   5880
      TabIndex        =   12
      Top             =   420
      Width           =   7470
      _ExtentX        =   13176
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "Clientes"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox TextConcepto 
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
      Left            =   4725
      MaxLength       =   100
      TabIndex        =   15
      Top             =   735
      Width           =   10410
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   1470
      TabIndex        =   2
      Top             =   0
      Width           =   2325
      Begin VB.OptionButton OpcE 
         Caption         =   "Egreso"
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
         Left            =   1155
         TabIndex        =   4
         Top             =   0
         Width           =   960
      End
      Begin VB.OptionButton OpcI 
         Caption         =   "Ingreso"
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
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Value           =   -1  'True
         Width           =   1065
      End
   End
   Begin VB.CheckBox CheqRF 
      Caption         =   "Retención en la Fuente"
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
      TabIndex        =   16
      Top             =   1155
      Width           =   2430
   End
   Begin VB.TextBox TxtFactNo 
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
      IMEMode         =   3  'DISABLE
      Left            =   13335
      MaxLength       =   10
      MultiLine       =   -1  'True
      TabIndex        =   25
      Text            =   "Kard_Ing.frx":1059
      Top             =   1155
      Width           =   2010
   End
   Begin VB.TextBox TxtDifxDec 
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
      Left            =   3990
      MultiLine       =   -1  'True
      TabIndex        =   68
      Text            =   "Kard_Ing.frx":105B
      Top             =   8190
      Width           =   1590
   End
   Begin VB.TextBox TxtSubTotal 
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
      Left            =   5670
      MultiLine       =   -1  'True
      TabIndex        =   75
      Text            =   "Kard_Ing.frx":105D
      Top             =   8190
      Width           =   1590
   End
   Begin MSDataListLib.DataCombo DCTInv 
      Bindings        =   "Kard_Ing.frx":105F
      DataSource      =   "AdoTInv"
      Height          =   315
      Left            =   105
      TabIndex        =   26
      Top             =   1575
      Width           =   4845
      _ExtentX        =   8546
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16761024
      Text            =   "DC"
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
   Begin VB.OptionButton OpcIVA 
      BackColor       =   &H00FF8080&
      Caption         =   "Con IVA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   5040
      TabIndex        =   29
      Top             =   1890
      Width           =   1485
   End
   Begin VB.OptionButton OpcX 
      BackColor       =   &H00FF8080&
      Caption         =   "Sin IVA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   6510
      TabIndex        =   30
      Top             =   1890
      Width           =   1695
   End
   Begin MSDataListLib.DataCombo DCBodega 
      Bindings        =   "Kard_Ing.frx":1075
      DataSource      =   "AdoBodega"
      Height          =   315
      Left            =   5040
      TabIndex        =   31
      Top             =   2205
      Width           =   3165
      _ExtentX        =   5583
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16761024
      Text            =   "DC"
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
   Begin MSDataGridLib.DataGrid DGKardex 
      Bindings        =   "Kard_Ing.frx":108D
      Height          =   3060
      Left            =   0
      TabIndex        =   66
      Top             =   4725
      Width           =   15030
      _ExtentX        =   26511
      _ExtentY        =   5398
      _Version        =   393216
      BackColor       =   16761024
      BorderStyle     =   0
      ForeColor       =   192
      HeadLines       =   1
      RowHeight       =   15
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
   Begin VB.TextBox TextVUnit 
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
      Left            =   9240
      MultiLine       =   -1  'True
      TabIndex        =   43
      Text            =   "Kard_Ing.frx":10A5
      Top             =   2940
      Width           =   1590
   End
   Begin VB.TextBox TextEntrada 
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
      IMEMode         =   3  'DISABLE
      Left            =   8085
      MultiLine       =   -1  'True
      TabIndex        =   41
      Text            =   "Kard_Ing.frx":10A9
      Top             =   2940
      Width           =   1170
   End
   Begin MSAdodcLib.Adodc AdoCtaObra 
      Height          =   330
      Left            =   2310
      Top             =   5565
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
      Caption         =   "CtaObra"
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
   Begin MSAdodcLib.Adodc AdoBenef 
      Height          =   330
      Left            =   2310
      Top             =   5250
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
      Caption         =   "Benef"
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
   Begin MSAdodcLib.Adodc AdoKardex 
      Height          =   330
      Left            =   4320
      Top             =   5520
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
      Caption         =   "Kardex"
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
   Begin MSAdodcLib.Adodc AdoInv 
      Height          =   330
      Left            =   4305
      Top             =   4935
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
      Caption         =   "Inv"
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
   Begin MSAdodcLib.Adodc AdoArt 
      Height          =   330
      Left            =   315
      Top             =   5565
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
      Caption         =   "Art"
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
   Begin MSAdodcLib.Adodc AdoAsientos 
      Height          =   330
      Left            =   4320
      Top             =   5205
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
      Caption         =   "Asientos"
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
   Begin VB.TextBox TextOrden 
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
      IMEMode         =   3  'DISABLE
      Left            =   6615
      MaxLength       =   10
      MultiLine       =   -1  'True
      TabIndex        =   39
      Text            =   "Kard_Ing.frx":10AB
      Top             =   2940
      Width           =   1485
   End
   Begin VB.TextBox TextIVA 
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
      Left            =   7350
      MultiLine       =   -1  'True
      TabIndex        =   70
      Text            =   "Kard_Ing.frx":10AD
      Top             =   8190
      Width           =   1590
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Salir"
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
      Left            =   14070
      Picture         =   "Kard_Ing.frx":10AF
      Style           =   1  'Graphical
      TabIndex        =   74
      Top             =   7875
      Width           =   960
   End
   Begin VB.CommandButton Command1 
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
      Left            =   13020
      Picture         =   "Kard_Ing.frx":1979
      Style           =   1  'Graphical
      TabIndex        =   73
      Top             =   7875
      Width           =   960
   End
   Begin MSAdodcLib.Adodc AdoTInv 
      Height          =   330
      Left            =   315
      Top             =   5250
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
      Caption         =   "TInv"
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
   Begin MSAdodcLib.Adodc AdoBodega 
      Height          =   330
      Left            =   315
      Top             =   4935
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
      Caption         =   "Bodega"
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
   Begin MSAdodcLib.Adodc AdoRet 
      Height          =   330
      Left            =   2310
      Top             =   4935
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
      Caption         =   "Ret"
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
   Begin MSAdodcLib.Adodc AdoIVA 
      Height          =   330
      Left            =   315
      Top             =   5880
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
      Caption         =   "IVA"
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
   Begin MSAdodcLib.Adodc AdoDiario 
      Height          =   330
      Left            =   2310
      Top             =   5880
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
      Caption         =   "Kardex"
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
   Begin MSAdodcLib.Adodc AdoMarca 
      Height          =   330
      Left            =   4305
      Top             =   5880
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
      Caption         =   "Marca"
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
      Left            =   6405
      Top             =   4935
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
   Begin MSMask.MaskEdBox MBFechaFab 
      Height          =   330
      Left            =   6405
      TabIndex        =   49
      Top             =   3675
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
   Begin InetCtlsObjects.Inet URLInet 
      Left            =   15330
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      Protocol        =   5
      RemotePort      =   443
      URL             =   "https://"
   End
   Begin MSDataListLib.DataCombo DCPorcIVA 
      Bindings        =   "Kard_Ing.frx":1DBB
      DataSource      =   "AdoPorcIVA"
      Height          =   360
      Left            =   9450
      TabIndex        =   21
      ToolTipText     =   "Seleccione el Porc del IVA"
      Top             =   1155
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   635
      _Version        =   393216
      Text            =   "12%"
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
   Begin MSAdodcLib.Adodc AdoPorcIVA 
      Height          =   330
      Left            =   315
      Top             =   6195
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
      Caption         =   "PorcIVA"
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
   Begin VB.Label Label32 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Porc IVA"
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
      TabIndex        =   20
      Top             =   1155
      Width           =   960
   End
   Begin VB.Label Label31 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Serie No."
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
      Left            =   10395
      TabIndex        =   22
      Top             =   1155
      Width           =   960
   End
   Begin VB.Label Label30 
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " SERIE No."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   7665
      TabIndex        =   58
      Top             =   3990
      Width           =   3270
   End
   Begin VB.Label Label29 
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " PROCEDENCIA/UBICACION"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   5040
      TabIndex        =   56
      Top             =   3990
      Width           =   2640
   End
   Begin VB.Label Label28 
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " MODELO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   10920
      TabIndex        =   54
      Top             =   3360
      Width           =   4215
   End
   Begin VB.Label Label27 
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " REG. SANITARIO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   8925
      TabIndex        =   52
      Top             =   3360
      Width           =   2010
   End
   Begin VB.Label Label26 
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " FECHA EXP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   7665
      TabIndex        =   50
      Top             =   3360
      Width           =   1275
   End
   Begin VB.Label Label25 
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " FECHA FAB"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   6405
      TabIndex        =   48
      Top             =   3360
      Width           =   1275
   End
   Begin VB.Label Label24 
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " LOTE No."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   5040
      TabIndex        =   46
      Top             =   3360
      Width           =   1380
   End
   Begin VB.Label Label21 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Vencimiento:"
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
      Top             =   735
      Width           =   1275
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Fecha:"
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
      Top             =   420
      Width           =   1275
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " POR CONCEPTO DE:"
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
      Left            =   2730
      TabIndex        =   14
      Top             =   735
      Width           =   2010
   End
   Begin VB.Label Label23 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " TD:"
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
      Width           =   540
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " VALOR TOTAL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   13440
      TabIndex        =   64
      Top             =   3990
      Width           =   1695
   End
   Begin VB.Label Label14 
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " DESC. 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   12180
      TabIndex        =   62
      Top             =   3990
      Width           =   1275
   End
   Begin VB.Label Label12 
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " DESC. 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   10920
      TabIndex        =   60
      Top             =   3990
      Width           =   1275
   End
   Begin VB.Label Label22 
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " MARCA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   8190
      TabIndex        =   32
      Top             =   1890
      Width           =   3900
   End
   Begin VB.Label Label2 
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
      Left            =   2730
      TabIndex        =   11
      Top             =   420
      Width           =   3165
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
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
      Height          =   645
      Left            =   105
      TabIndex        =   77
      Top             =   7875
      Width           =   2535
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " UNIDAD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   5040
      TabIndex        =   36
      Top             =   2625
      Width           =   1590
   End
   Begin VB.Label Label15 
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " CODIGO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   12075
      TabIndex        =   34
      Top             =   1890
      Width           =   3060
   End
   Begin VB.Label LabelProducto 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Producto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   5040
      TabIndex        =   28
      Top             =   1575
      Width           =   10095
   End
   Begin VB.Label Label20 
      Caption         =   " Retención del I.V.A."
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
      Left            =   4725
      TabIndex        =   18
      Top             =   1155
      Width           =   1905
   End
   Begin VB.Label LblRF 
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
      Height          =   330
      Left            =   2730
      TabIndex        =   17
      Top             =   1155
      Width           =   1695
   End
   Begin VB.Label LblRIVA 
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
      Height          =   330
      Left            =   6720
      TabIndex        =   19
      Top             =   1155
      Width           =   1695
   End
   Begin VB.Label Label19 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Factura No."
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
      Left            =   12180
      TabIndex        =   24
      Top             =   1155
      Width           =   1170
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DIFxDECIMALES"
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
      TabIndex        =   67
      Top             =   7875
      Width           =   1590
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "S U B T O T A L"
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
      TabIndex        =   76
      Top             =   7875
      Width           =   1590
   End
   Begin VB.Label Label17 
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " CODIGO DE BARRA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   10815
      TabIndex        =   44
      Top             =   2625
      Width           =   4320
   End
   Begin VB.Label Label13 
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " GUIA No."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   6615
      TabIndex        =   38
      Top             =   2625
      Width           =   1485
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
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
      Left            =   9030
      TabIndex        =   72
      Top             =   8190
      Width           =   1590
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " VALOR UNIT."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   9240
      TabIndex        =   42
      Top             =   2625
      Width           =   1590
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " CANTIDAD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   8085
      TabIndex        =   40
      Top             =   2625
      Width           =   1170
   End
   Begin VB.Label LabelUnidad 
      BackColor       =   &H80000005&
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
      Left            =   5040
      TabIndex        =   37
      Top             =   2940
      Width           =   1590
   End
   Begin VB.Label LabelCodigo 
      BackColor       =   &H80000005&
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
      Left            =   12075
      TabIndex        =   35
      Top             =   2205
      Width           =   3060
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "T O T A L"
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
      Left            =   9030
      TabIndex        =   71
      Top             =   7875
      Width           =   1590
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "I.V.A."
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
      Left            =   7350
      TabIndex        =   69
      Top             =   7875
      Width           =   1590
   End
End
Attribute VB_Name = "Kard_Ing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Variables para acceder a la hoja excel
Dim obj_Excel As Object, obj_Workbook As Object, obj_Worksheet As Object

'Type para el Rango
Private Type T_Rango
        NumFila1 As Long
        NumFila2 As Long
        NumCol1 As Long
        NumCol2 As Long
End Type

'Variable para el UDT que almacena cuatro variables par el Rango de datos de la hoha Excel
Dim Rango As T_Rango

'función que recibe el Apth del ArchivoExcel y una variable de tipo T_Rango para almacenar los valores y retornarlos
Private Function Obtener_Rango_Excel(path_excel As String, Rango As T_Rango) As Boolean
On Error GoTo ErrSub
'Variables de objeto Excel
Dim objExcel As Object, obj_Hoja As Object
Dim Primera_Fila As Integer, Primera_Column As Integer
Dim num_Filas As Integer, num_Col As Integer

   'Crear la instancia de la aplicación Excel
    Set objExcel = CreateObject("Excel.Application")
    objExcel.Workbooks.open Filename:=path_excel
    If Val(objExcel.Application.version) >= 8 Then
        Set obj_Hoja = objExcel.ActiveSheet
    Else
        Set obj_Hoja = objExcel
    End If
   'Almacenamos en la variable Type el rango, es decir la primera fila la ultima fila, La primer Columna y la ultima
    With Rango
        .NumFila1 = Format$(obj_Hoja.UsedRange.Row)
        .NumFila2 = Format$(obj_Hoja.UsedRange.Row + obj_Hoja.UsedRange.rows.Count - 1)
        .NumCol1 = Format$(obj_Hoja.UsedRange.Column)
        .NumCol2 = Format$(obj_Hoja.UsedRange.Column + obj_Hoja.UsedRange.Columns.Count - 1)
    End With
    objExcel.ActiveWorkbook.Close False
   'cierra el Excel
    objExcel.Quit
   'Elimina las referencias y objetos
    Set obj_Hoja = Nothing
    Set objExcel = Nothing
Exit Function
ErrSub:
MsgBox " Número de Error: " & Err.Number & vbNewLine & "Descripción: " & Err.Description
End Function

Private Sub Descargar()
    On Local Error Resume Next
    Set obj_Workbook = Nothing
    Set obj_Excel = Nothing
    Set obj_Worksheet = Nothing
    Me.MousePointer = vbDefault
End Sub

Private Sub Excel_Compras(sPath As String, Rango As T_Rango)
Dim I As Long
Dim N As Long
Dim No_Bancos As Long
Dim TextoCel As String
Dim ContItem As Integer
Dim DepItem As String
Dim PVP As Currency
Dim PVP2 As Currency
Dim PVP3 As Currency

    On Error GoTo error_sub
    If Len(Dir(sPath)) = 0 Then
       MsgBox "El archivo no existe", vbCritical
       Exit Sub
    End If
    RatonReloj
    Set obj_Excel = CreateObject("Excel.Application")
   'obj_Excel.Visible = True
    Set obj_Workbook = obj_Excel.Workbooks.open(sPath)
    Set obj_Worksheet = obj_Workbook.ActiveSheet
    ContItem = 1
    
    DGKardex.Visible = False
    IniciarAsientosAdo AdoAsientos
   
    Cod_Bodega = Ninguno
    Cod_Marca = Ninguno
    With AdoBodega.Recordset
     If .RecordCount > 0 Then
        .MoveFirst
        .Find ("Bodega Like '" & DCBodega & "' ")
         If Not .EOF Then Cod_Bodega = .fields("CodBod")
     End If
    End With
    
    With AdoMarca.Recordset
     If .RecordCount > 0 Then
        .MoveFirst
        .Find ("Marca Like '" & DCMarca & "' ")
         If Not .EOF Then Cod_Marca = .fields("CodMar")
     End If
    End With
    
    sSQL = "SELECT * " _
         & "FROM Asiento_K " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND CodigoU = '" & CodigoUsuario & "' " _
         & "AND T_No = " & Trans_No & " "
    SQLDec = "VALOR_UNIT " & CStr(Dec_Costo) & "|,VALOR_TOTAL 4|,CANTIDAD 2|,SALDO 4|."
    Select_Adodc_Grid DGKardex, AdoKardex, sSQL, SQLDec
    FechaTexto = MBFechaI
    Factura_No = Val(TxtFactNo)
    If Factura_No <= 0 Then Factura_No = 0
    TxtFactNo = Factura_No
    SaldoAnterior = 0
    Cantidad = 0
    'OpcDH = "1"
    If OpcI.value Then OpcDH = 1 Else OpcDH = 2
    Contra_Cta = SinEspaciosDer(DCCtaObra)
    Codigo = Leer_Cta_Catalogo(Contra_Cta)
    With Rango
       ' Especificar acá la cantidad de filas y columnas
       ' MsgBox rango.NumCol2 & vbCrLf & rango.NumFila2
       ' Recorremos las filas del excel
         For I = 2 To .NumFila2
            'Recorremos las columnas y Fila del FlexGrid
             Entrada = Redondear(CCur(obj_Worksheet.cells(I, 1).value), 2)
             CodigoBarra = TrimStrg(obj_Worksheet.cells(I, 2).value)
             CodigoInv = TrimStrg(obj_Worksheet.cells(I, 3).value)
             Costo = TrimStrg(obj_Worksheet.cells(I, 4).value)
             PVP = obj_Worksheet.cells(I, 5).value
             PVP2 = obj_Worksheet.cells(I, 6).value
             PVP3 = obj_Worksheet.cells(I, 7).value
             CadenaParcial = obj_Worksheet.cells(I, 9).value
             Tasa = Val(obj_Worksheet.cells(I, 10).value)
             'MsgBox "Desktop Test"
             If AdoInv.Recordset.RecordCount > 0 Then
                AdoInv.Recordset.MoveFirst
                AdoInv.Recordset.Find ("Codigo_Inv Like '" & CodigoInv & "' ")
                If Not AdoInv.Recordset.EOF Then
                   Unidad = AdoInv.Recordset.fields("Unidad")
                   CodigoInv = AdoInv.Recordset.fields("Codigo_Inv")
                   Producto = AdoInv.Recordset.fields("Producto")
                   Cta_Inventario = AdoInv.Recordset.fields("Cta_Inventario")
                Else
                   CodigoInv = Ninguno
                   Producto = Ninguno
                End If
             End If
             If OpcDH = 1 Then
                ValorUnit = Redondear(CCur(obj_Worksheet.cells(I, 4).value), 6)
             Else
                Stock_Actual_Inventario FechaSistema, CodigoInv
             End If
             If ValorUnit <= 0 Then Stock_Actual_Inventario FechaSistema, CodigoInv
             SubTotal_IVA = 0
             ValorTotal = Redondear(Entrada * ValorUnit, 2)
             If OpcIVA.value Then SubTotal_IVA = Redondear(ValorTotal * Porc_IVA, 2)
             Cantidad = Entrada
             Saldo = SaldoAnterior + ValorTotal
             If Entrada > 0 And ValorUnit > 0 And Len(CodigoInv) > 1 Then
                SetAddNew AdoKardex
                SetFields AdoKardex, "DH", OpcDH
                SetFields AdoKardex, "CODIGO_INV", CodigoInv
                SetFields AdoKardex, "PRODUCTO", Producto
                SetFields AdoKardex, "CANT_ES", Entrada
                SetFields AdoKardex, "VALOR_UNIT", ValorUnit
                SetFields AdoKardex, "VALOR_TOTAL", ValorTotal
                SetFields AdoKardex, "IVA", SubTotal_IVA
                SetFields AdoKardex, "CTA_INVENTARIO", Cta_Inventario
                SetFields AdoKardex, "CONTRA_CTA", Contra_Cta
                SetFields AdoKardex, "CANTIDAD", Cantidad
                SetFields AdoKardex, "VALOR_UNIT", ValorUnit
                SetFields AdoKardex, "SALDO", Saldo
                SetFields AdoKardex, "UNIDAD", Unidad
                SetFields AdoKardex, "CodBod", Cod_Bodega
                SetFields AdoKardex, "CodMar", Cod_Marca
                SetFields AdoKardex, "Item", NumEmpresa
                SetFields AdoKardex, "CodigoU", CodigoUsuario
                SetFields AdoKardex, "T_No", Trans_No
                SetFields AdoKardex, "SUBCTA", SubCtaGen
                SetFields AdoKardex, "TC", SubCta
                SetFields AdoKardex, "Codigo_B", CodigoCliente
                SetFields AdoKardex, "COD_BAR", CodigoBarra
                SetFields AdoKardex, "ORDEN", TextOrden
                If Tasa > 0 Then SetFields AdoKardex, "Peso", Tasa
                If Len(CadenaParcial) > 1 Then SetFields AdoKardex, "Detalle", UCase(MidStrg(CadenaParcial, 1, 60))
                SetUpdate AdoKardex
                
                sSQL = "UPDATE Catalogo_Productos " _
                     & "SET PVP = " & PVP & " " _
                     & "WHERE Item = '" & NumEmpresa & "' " _
                     & "AND Periodo = '" & Periodo_Contable & "' " _
                     & "AND Codigo_Inv = '" & CodigoInv & "' " _
                     & "AND PVP <> " & PVP & " "
                Ejecutar_SQL_SP sSQL
                sSQL = "UPDATE Catalogo_Productos " _
                     & "SET PVP_2 = " & PVP2 & " " _
                     & "WHERE Item = '" & NumEmpresa & "' " _
                     & "AND Periodo = '" & Periodo_Contable & "' " _
                     & "AND Codigo_Inv = '" & CodigoInv & "' " _
                     & "AND PVP_2 <> " & PVP2 & " "
                Ejecutar_SQL_SP sSQL
                sSQL = "UPDATE Catalogo_Productos " _
                     & "SET PVP_3 = " & PVP3 & " " _
                     & "WHERE Item = '" & NumEmpresa & "' " _
                     & "AND Periodo = '" & Periodo_Contable & "' " _
                     & "AND Codigo_Inv = '" & CodigoInv & "' " _
                     & "AND PVP_3 <> " & PVP3 & " "
                Ejecutar_SQL_SP sSQL
              End If
         Next I
    End With
    obj_Workbook.Close
    obj_Excel.Quit
    Descargar
    TotalInventario
    'TextTotal = Format(ValorTotal, "#,##0.00")
    DGKardex.Visible = True
    Command1.SetFocus
    RatonNormal
Exit Sub
error_sub:
MsgBox Err.Description
Descargar
End Sub
 
 
Private Sub CheqContraCta_Click()
 If CheqContraCta.value = 1 Then
    DCCtaObra.Visible = True
    DCBenef.Visible = True
 Else
    If OpcE.value Then
       DCCtaObra.Visible = False
       DCBenef.Visible = False
    End If
 End If
End Sub

Private Sub CheqRF_Click()
 If CheqRF.value = 1 Then
 
    FComprasAT.Show 1
 End If
 TxtFactNo = Format(Factura_CxP, "000000000")
 LblRF.Caption = Format(Total_Ret, "#,##0.00")
 LblRIVA.Caption = Format(Total_RetIVA, "#,##0.00")
End Sub

Private Sub CLTP_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

'Grabar Comprobante de Ingreso/Egreso de Inventario
Private Sub Command1_Click()
  RatonNormal
  Trans_No = 97
  DetalleComp = Ninguno
  Ln_No = 1
  Asiento = 1
  If CodigoCliente = "" Then CodigoCliente = Ninguno
  Total_Factura = 0
  Si_No = False
  FechaValida MBFechaI
  FechaValida MBVence
  FechaTexto = MBFechaI
  TextoValido TextOrden, , True
  If Len(TxtSerie) <> 6 Then TxtSerie = "001001"
 'TotalInventario
 'MsgBox CodigoUsuario & vbCrLf & CodigoCliente
  Factura_No = Val(TxtFactNo)
  If Factura_No <= 0 Then Factura_No = 0 'Val(Format(Year(FechaTexto), "0000") & Format(Month(FechaTexto), "00") & Format(Day(FechaTexto), "00"))
  sSQL = "DELETE * " _
       & "FROM Asiento_SC " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND T_No = " & Trans_No & " "
  Ejecutar_SQL_SP sSQL
  
  sSQL = "DELETE * " _
       & "FROM Asiento " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND T_No = " & Trans_No & " "
  Ejecutar_SQL_SP sSQL
' Cta de Inventario
  sSQL = "SELECT * " _
       & "FROM Asiento_K " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND T_No = " & Trans_No & " " _
       & "ORDER BY CTA_INVENTARIO,CONTRA_CTA "
  SQLDec = "VALOR_UNIT " & CStr(Dec_Costo) & "|,VALOR_TOTAL 4|,CANTIDAD 2|,SALDO 4|."
  Select_Adodc_Grid DGKardex, AdoKardex, sSQL, SQLDec, , , "Asiento K"
  With AdoKardex.Recordset
   If .RecordCount > 0 Then
       Mensajes = "Seguro de Grabar?"
       Titulo = "GRABACION DEL COMPROBANTE"
       If BoxMensaje = vbYes Then
          RatonReloj
          Si_No = True
          FechaComp = MBFechaI
          FechaTexto = MBFechaI
          
          If CLTP.Text = "" Then CLTP.Text = "CD"
          Select Case CLTP.Text
            Case "CD": NumComp = ReadSetDataNum("Diario", True, True)
            Case "ND": NumComp = ReadSetDataNum("NotaDebito", True, True)
            Case "NC": NumComp = ReadSetDataNum("NotaCredito", True, True)
          End Select
             
          CodigoInv = .fields("Codigo_Inv")
          Cta_Inventario = .fields("CTA_INVENTARIO")
          Total = 0: ValorDH = 0
         'Llenamos los datos ingresados al Kardex
          Do While Not .EOF
             If Cta_Inventario <> .fields("CTA_INVENTARIO") Then
                If OpcI.value Then
                   InsertarAsientos AdoAsientos, Cta_Inventario, 0, ValorDH, 0
                Else
                   InsertarAsientos AdoAsientos, Cta_Inventario, 0, 0, ValorDH
                End If
                CodigoInv = .fields("Codigo_Inv")
                Cta_Inventario = .fields("CTA_INVENTARIO")
                ValorDH = 0
             End If
             ValorDH = ValorDH + .fields("VALOR_TOTAL")
            .MoveNext
          Loop
          If OpcI.value Then
             InsertarAsientos AdoAsientos, Cta_Inventario, 0, ValorDH, 0
          Else
             InsertarAsientos AdoAsientos, Cta_Inventario, 0, 0, ValorDH
          End If
       End If
   End If
  End With
  
  OpcTM = 1
  OpcDH = 2
  NoCheque = Ninguno
 'Grabamos el Asiento de la Compra
  sSQL = "SELECT * " _
       & "FROM Asiento_Compras " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND T_No = " & Trans_No & " "
  Select_Adodc AdoRet, sSQL
    With AdoRet.Recordset
     If .RecordCount > 0 Then
        'Porcentaje por Servicio: 0,30,100
         Cta = .fields("Cta_Servicio")
         DetalleComp = "Retencion del " & .fields("Porc_Bienes") & "%, Factura No. " & .fields("Secuencial") & ", de " & NombreCliente
         Codigo = Leer_Cta_Catalogo(Cta)
         ValorDH = .fields("ValorRetServicios")
         Total_RetIVA = Total_RetIVA + .fields("ValorRetServicios")
         If ValorDH > 0 Then InsertarAsiento AdoAsientos
        'Porcentaje por Bienes: 0,70,100
         Cta = .fields("Cta_Bienes")
         DetalleComp = "Retencion del " & .fields("Porc_Servicios") & "%, Factura No. " & .fields("Secuencial") & ", de " & NombreCliente
         Codigo = Leer_Cta_Catalogo(Cta)
         ValorDH = .fields("ValorRetBienes")
         Total_RetIVA = Total_RetIVA + .fields("ValorRetBienes")
         If ValorDH > 0 Then InsertarAsiento AdoAsientos
     End If
    End With
   'Grabamos el Asiento de las Retenciones
    sSQL = "SELECT * " _
         & "FROM Asiento_Air " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND CodigoU = '" & CodigoUsuario & "' " _
         & "AND T_No = " & Trans_No & " " _
         & "AND Tipo_Trans = 'C' " _
         & "ORDER BY Cta_Retencion,A_No,ValRet "
    Select_Adodc AdoRet, sSQL
    With AdoRet.Recordset
     If .RecordCount > 0 Then
         Do While Not .EOF
            Cta = .fields("Cta_Retencion")
            DetalleComp = "Retencion (" & .fields("CodRet") & ") No. " & .fields("SecRetencion") & " del " & (.fields("Porcentaje") * 100) & "%, de " & NombreCliente
            Codigo = Leer_Cta_Catalogo(Cta)
            ValorDH = .fields("ValRet")
            Total_Ret = Total_Ret + .fields("ValRet")
            If ValorDH > 0 Then InsertarAsiento AdoAsientos
           .MoveNext
         Loop
     End If
    End With
    DetalleComp = Ninguno
' Contra Cuenta del Kardex
  If Si_No Then
     sSQL = "SELECT * " _
          & "FROM Asiento_K " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND CodigoU = '" & CodigoUsuario & "' " _
          & "AND T_No = " & Trans_No & " " _
          & "ORDER BY CONTRA_CTA "
     SQLDec = "VALOR_UNIT " & CStr(Dec_Costo) & "|,VALOR_TOTAL 4|,CANTIDAD 2|,SALDO 4|."
     Select_Adodc_Grid DGKardex, AdoKardex, sSQL, SQLDec
     With AdoKardex.Recordset
      If .RecordCount > 0 Then
          RatonReloj
          SubCta = .fields("TC")
          Contra_Cta = .fields("CONTRA_CTA")
          Total = 0: ValorDH = 0
         'Llenamos los datos ingresados al Kardex
          If OpcE.value And CheqContraCta.value = 0 Then
             Do While Not .EOF
                If Contra_Cta <> .fields("CONTRA_CTA") Then
                   If OpcI.value Then
                      InsertarAsientos AdoAsientos, Contra_Cta, 0, 0, ValorDH
                   Else
                      InsertarAsientos AdoAsientos, Contra_Cta, 0, ValorDH, 0
                   End If
                   SubCta = .fields("TC")
                   Contra_Cta = .fields("CONTRA_CTA")
                   ValorDH = 0
                End If
                ValorDH = ValorDH + .fields("VALOR_TOTAL")
               .MoveNext
             Loop
             If OpcI.value Then
                InsertarAsientos AdoAsientos, Contra_Cta, 0, 0, ValorDH
             Else
                InsertarAsientos AdoAsientos, Contra_Cta, 0, ValorDH, 0
             End If
          End If
      End If
     End With
     TotalInventario
   ' Insertamos El IVA de la compra
     If OpcI.value Then
        InsertarAsientos AdoAsientos, Cta_IVA_Inventario, 0, Total_IVA, 0
     Else
        InsertarAsientos AdoAsientos, Cta_IVA_Inventario, 0, 0, Total_IVA
     End If
     If CCur(Val(TxtDifxDec)) > 0 Then
        InsertarAsientos AdoAsientos, Cta_Faltantes, 0, CCur(Val(TxtDifxDec)), 0
     ElseIf CCur(Val(TxtDifxDec)) < 0 Then
        InsertarAsientos AdoAsientos, Cta_Faltantes, 0, 0, -CCur(Val(TxtDifxDec))
     End If
     sSQL = "SELECT * " _
          & "FROM Asiento " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND CodigoU = '" & CodigoUsuario & "' " _
          & "AND T_No = " & Trans_No & " "
     Select_Adodc AdoAsientos, sSQL
     With AdoAsientos.Recordset
      If .RecordCount > 0 Then
          Debe = 0: Haber = 0
          Do While Not .EOF
             Debe = Debe + .fields("DEBE")
             Haber = Haber + .fields("HABER")
            .MoveNext
          Loop
      End If
     End With
     If CheqContraCta.value = 1 Then
        Contra_Cta = SinEspaciosDer(DCCtaObra)
        Codigo = Leer_Cta_Catalogo(Contra_Cta)
        ValorDH = Debe - Haber
        If ValorDH < 0 Then ValorDH = -ValorDH
        Select Case SubCta
          Case "C", "P", "G", "I"
               Total_Factura = ValorDH
               SetAdoAddNew "Asiento_SC"
               SetAdoFields "TM", "1"
               SetAdoFields "Serie", TxtSerie
               SetAdoFields "Factura", Val(TxtFactNo)
               SetAdoFields "Codigo", CodigoCliente
               SetAdoFields "FECHA_V", MBVence
               SetAdoFields "Cta", Contra_Cta
               SetAdoFields "TC", SubCta
               SetAdoFields "T_No", Trans_No
               SetAdoFields "SC_No", Ln_No
               If OpcI.value Then SetAdoFields "DH", "2" Else SetAdoFields "DH", "1"
               SetAdoFields "Valor", ValorDH
               SetAdoFields "Periodo", Periodo_Contable
               SetAdoUpdate
        End Select
        If OpcI.value Then
           InsertarAsientos AdoAsientos, Contra_Cta, 0, 0, ValorDH
        Else
           InsertarAsientos AdoAsientos, Contra_Cta, 0, ValorDH, 0
        End If
     End If
     
     'MsgBox CodigoUsuario & vbCrLf & CodigoCliente
     
     'MsgBox CodigoUsuario
     If Cod_Benef = "X" Then CodigoCliente = Ninguno
  End If
  'MsgBox CodigoUsuario
  'MsgBox CodigoUsuario & vbCrLf & CodigoCliente
  sSQL = "SELECT * " _
       & "FROM Asiento " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND T_No = " & Trans_No & " "
  Select_Adodc AdoAsientos, sSQL
  With AdoAsientos.Recordset
   If .RecordCount > 0 Then
       Debe = 0: Haber = 0
       Do While Not .EOF
          Debe = Debe + .fields("DEBE")
          Haber = Haber + .fields("HABER")
         .MoveNext
       Loop
       If (Debe - Haber) <> 0 Then MsgBox "Verifique el comprobante, no cuadra por: " & Redondear(Debe - Haber, 2)
      .MoveFirst
     ' MsgBox (Debe - Haber)
          RatonReloj
          Co.T = Normal
          Co.TP = CLTP.Text
          Co.Fecha = FechaTexto
          Co.Numero = NumComp
          Co.Concepto = TextConcepto
          Co.CodigoB = CodigoCliente
          Co.Efectivo = 0
          Co.Monto_Total = Total - Total_RetCta
          Co.Usuario = CodigoUsuario
          Co.T_No = Trans_No
          Co.Item = NumEmpresa
          Factura_No = Val(TxtFactNo)
          If Factura_No <= 0 Then Factura_No = 0 'Val(Format(Year(FechaTexto), "0000") & Format(Month(FechaTexto), "00") & Format(Day(FechaTexto), "00"))
          TxtFactNo = Factura_No
          If Len(TextOrden) > 1 Then Co.Concepto = Co.Concepto & ", Orden No. " & TextOrden
          If Val(TxtFactNo) > 0 Then Co.Concepto = Co.Concepto & ", Factura No. " & TxtFactNo
          Grabar_Comprobante Co
         'MsgBox "Hola"
          ImprimirComprobantesDe False, Co
          Imprimir_Nota_Inventario AdoBenef, AdoKardex, NumComp, TextOrden.Text, "CD", FechaTexto, FechaTexto, Total
          Mensajes = "Imprimir Copia de Nota de Entrada/Salida"
          Titulo = "COPIA DE NOTA"
          If BoxMensaje = vbYes Then Imprimir_Nota_Inventario AdoBenef, AdoKardex, NumComp, TextOrden.Text, "CD", FechaTexto, FechaTexto, Total
          Unload Kard_Ing
          Mayorizar_Inventario_SP
   Else
       MsgBox "No existen Datos para procesar"
   End If
  End With
  RatonNormal
End Sub

Private Sub Command2_Click()
  Unload Kard_Ing
End Sub

Private Sub Command3_Click()
' Grabamos Inventarios
  If CLTP.Text = "" Then CLTP.Text = "CD"
  Numero = Val(InputBox("INGRESE EL NUMERO DE COMPROBANTE", "GRABACION DE COMPROBANTE: " & CLTP.Text, "0"))
  If Numero > 1 Then
     SQL1 = "DELETE * " _
          & "FROM Trans_Kardex " _
          & "WHERE TP = 'CD' " _
          & "AND Numero = " & Numero & " " _
          & "AND Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' "
     Ejecutar_SQL_SP SQL1
     RatonReloj
     Trans_No = 97
     sSQL = "SELECT * " _
          & "FROM Asiento_K " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND CodigoU = '" & CodigoUsuario & "' " _
          & "AND T_No = " & Trans_No & " "
     Select_Adodc_Grid DGKardex, AdoKardex, sSQL, 4
     With AdoKardex.Recordset
      If .RecordCount > 0 Then
          Do While Not .EOF
           ' Asiento de Inventario
             SetAdoAddNew "Trans_Kardex"
             SetAdoFields "T", Normal
             SetAdoFields "TP", CLTP.Text
             SetAdoFields "Numero", Numero
             SetAdoFields "Fecha", MBFechaI
             SetAdoFields "Codigo_Inv", .fields("CODIGO_INV")
             SetAdoFields "Codigo_P", .fields("Codigo_B")
             SetAdoFields "Descuento", .fields("P_DESC")
             SetAdoFields "Descuento1", .fields("P_DESC1")
             SetAdoFields "Valor_Total", .fields("VALOR_TOTAL")
             SetAdoFields "Existencia", .fields("CANTIDAD")
             SetAdoFields "Valor_Unitario", .fields("VALOR_UNIT")
             SetAdoFields "Total", .fields("SALDO")
             SetAdoFields "Cta_Inv", .fields("CTA_INVENTARIO")
             SetAdoFields "Contra_Cta", .fields("CONTRA_CTA")
             SetAdoFields "Orden_No", .fields("ORDEN")
             SetAdoFields "CodBodega", .fields("CodBod")
             SetAdoFields "CodMar", .fields("CodMar")
             SetAdoFields "Codigo_Barra", .fields("COD_BAR")
             SetAdoFields "Costo", .fields("VALOR_UNIT")
             SetAdoFields "PVP", .fields("PVP")
             If Inv_Promedio Then
                Cantidad = .fields("CANTIDAD")
                Saldo = .fields("SALDO")
                If Cantidad <= 0 Then Cantidad = 1
                SetAdoFields "Costo", Saldo / Cantidad
             End If
             If .fields("DH") = 1 Then
                SetAdoFields "Entrada", .fields("CANT_ES")
             Else
                SetAdoFields "Salida", .fields("CANT_ES")
                Si_No = False
             End If
             SetAdoFields "CodigoU", CodigoUsuario
             SetAdoFields "Item", NumEmpresa
             SetAdoUpdate
             NumTrans = NumTrans + 1
            .MoveNext
          Loop
          Eliminar_Asientos_SP True
      End If
     End With
     RatonNormal
     MsgBox "Proceso Exitoso"
  End If
End Sub

Private Sub Command4_Click()
  RutaOrigen = SelectDialogFile(RutaSysBases)
  If RutaOrigen <> "" Then
    'Le pasamos el Path del Libro y una variable de tipo T_Rango para retornar los valores
     Call Obtener_Rango_Excel(RutaOrigen, Rango)
     Call Excel_Compras(RutaOrigen, Rango)
  End If
End Sub

Private Sub DCBenef_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCBenef_LostFocus()
  InvImp = False
  Label2.Caption = " LOCAL"
  CodigoCliente = Ninguno
  NombreCliente = DCBenef
  With AdoBenef.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Cliente Like '" & NombreCliente & "' ")
       If Not .EOF Then
          CICliente = .fields("CI_RUC")
          CodigoBenef = .fields("Codigo")
          CodigoCliente = .fields("Codigo")
          NombreCliente = .fields("Cliente")
          Grupo_No = .fields("Grupo")
          TipoDoc = .fields("TD")
          TipoBenef = .fields("TD")
          Cod_Benef = .fields("TipoBenef")
          InvImp = .fields("Importaciones")
          If InvImp Then Label2.Caption = " IMPORTACION"
          Si_No = True
          If TipoDoc = "R" Then Si_No = False
          'MsgBox "...."
          Co.CodigoB = CodigoBenef
          Datos_del_Cliente Co
       Else
          Si_No = False
       End If
   End If
  End With
  Label3.Caption = CodigoCliente
  TextConcepto = DCBenef & " "
End Sub

Private Sub DCBodega_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCCtaObra_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCCtaObra_LostFocus()
' SubCta = .Fields("TC")
  Contra_Cta = SinEspaciosDer(DCCtaObra)
  Codigo = Leer_Cta_Catalogo(Contra_Cta)
  Contra_Cta = Codigo
  ListarProveedorUsuario SubCta
End Sub

Private Sub DCDiario_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
  If KeyCode = vbKeyEscape Then SiguienteControl
End Sub

Private Sub DCDiario_LostFocus()
  If Val(DCDiario.Text) > 0 Then MsgBox "Modificar"
  DCDiario.Visible = False
End Sub

Private Sub DCInv_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyEscape: Command1.SetFocus
    Case vbKeyReturn: SiguienteControl
  End Select
End Sub

Private Sub DCInv_LostFocus()

  Si_No = False: Evaluar = False
  If Opciones = 1 Then
     CodigoInv = DCInv.Text
  Else
     CodigoInv = SinEspaciosIzq(DCInv.Text)
  End If
  With AdoInv.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Producto Like '" & CodigoInv & "' ")
       If Not .EOF Then
          CodigoInv = .fields("Codigo_Inv")
          Evaluar = True
       Else
         .MoveFirst
         .Find ("Codigo_Barra Like '" & CodigoInv & "' ")
          If Not .EOF Then
             CodigoInv = .fields("Codigo_Inv")
             Evaluar = True
          Else
            .MoveFirst
            .Find ("Codigo_Inv Like '" & CodigoInv & "' ")
             If Not .EOF Then
                CodigoInv = .fields("Codigo_Inv")
                Evaluar = True
             Else
                Evaluar = False
             End If
          End If
       End If
    End If
  End With
  
  With AdoInv.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Codigo_Inv = '" & CodigoInv & "' ")
       If Not .EOF Then
          Si_No = .fields("IVA")
          Unidad = .fields("Unidad")
          Producto = .fields("Producto")
          Cta_Inventario = .fields("Cta_Inventario")
          Contra_Cta1 = .fields("Cta_Costo_Venta")
          TxtRegSanitario = .fields("Reg_Sanitario")
       End If
    End If
  End With

'''  sSQL = "SELECT Porc " _
'''       & "FROM Tabla_Por_ICE_IVA " _
'''       & "WHERE IVA <> " & Val(adFalse) & " " _
'''       & "AND Fecha_Inicio <= #" & BuscarFecha(MBFechaI) & "# " _
'''       & "AND Fecha_Final >= #" & BuscarFecha(MBFechaI) & "# " _
'''       & "ORDER BY Porc "
'''  Select_Adodc AdoAux, sSQL
'''  If AdoAux.Recordset.RecordCount > 0 Then Porc_IVA = Redondear(AdoAux.Recordset.fields("Porc") / 100, 2)
  
  LabelCodigo.Caption = CodigoInv
  LabelUnidad.Caption = Unidad
  LabelProducto.Caption = Producto
  If Si_No Then OpcIVA.value = True Else OpcX.value = True
  If Empleados Then OpcX.value = True
  If InvImp Then OpcX.value = True
  If Evaluar Then
     DCBodega.SetFocus
  Else
     MsgBox "No existe Productos asignados"
     DCTInv.SetFocus
  End If
End Sub

Private Sub DCMarca_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCPorcIVA_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCPorcIVA_LostFocus()
  If IsNumeric(DCPorcIVA.Text) Then
     If AdoPorcIVA.Recordset.RecordCount > 0 Then Porc_IVA = Redondear(DCPorcIVA / 100, 2)
  Else
     Porc_IVA = 0
  End If
End Sub

Private Sub DCTInv_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCTInv_LostFocus()
  ListarProductos
End Sub

Private Sub DGKardex_BeforeDelete(Cancel As Integer)
  Cancel = DeleteSiNo(AdoKardex)
End Sub

Private Sub MBFechaExp_GotFocus()
  MarcarTexto MBFechaExp
End Sub

Private Sub MBFechaExp_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBFechaExp_LostFocus()
  FechaValida MBFechaExp
End Sub

Private Sub MBFechaFab_GotFocus()
  MarcarTexto MBFechaFab
End Sub

Private Sub MBFechaFab_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBFechaFab_LostFocus()
    FechaValida MBFechaFab
End Sub

Private Sub MBFechaI_GotFocus()
  MarcarTexto MBFechaI
End Sub

Private Sub MBFechaI_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  PresionoEnter KeyCode
  If CtrlDown And KeyCode = vbKeyK Then
     DCDiario.Visible = True
     DCDiario.SetFocus
  End If
End Sub

Private Sub MBFechaI_LostFocus()
  FechaValida MBFechaI
  FechaComp = MBFechaI
  Fecha_Vence = MBFechaI
    sSQL = "SELECT Codigo, Porc " _
         & "FROM Tabla_Por_ICE_IVA " _
         & "WHERE IVA <> " & Val(adFalse) & " " _
         & "AND Fecha_Inicio <= #" & BuscarFecha(FechaComp) & "# " _
         & "AND Fecha_Final >= #" & BuscarFecha(FechaComp) & "# " _
         & "ORDER BY Porc DESC "
    SelectDB_Combo DCPorcIVA, AdoPorcIVA, sSQL, "Porc"
End Sub

Private Sub MBVence_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBVence_LostFocus()
  FechaValida MBVence
End Sub

Private Sub OpcE_Click()
   Total_IVA = 0
   Label11.Visible = False
   TextIVA.Visible = False
   CheqContraCta.value = 0
   DCCtaObra.Visible = False
   DCBenef.Visible = False
   CodigoCliente = Ninguno
   CheqRF.Enabled = False
   Select Case CLTP
     Case "ND", "NC"
          CheqRF.Enabled = True
          Label11.Visible = True
          TextIVA.Visible = True
   End Select
End Sub

Private Sub OpcI_Click()
   Total_IVA = 0
   Label11.Visible = True
   TextIVA.Visible = True
   CheqContraCta.value = 1
   DCCtaObra.Visible = True
   DCBenef.Visible = True
   CodigoCliente = Ninguno
   CheqRF.Enabled = True
   DCCtaObra.SetFocus
End Sub

Private Sub TextConcepto_GotFocus()
  MarcarTexto TextConcepto
End Sub

Private Sub TextConcepto_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextConcepto_LostFocus()
  TextoValido TextConcepto
End Sub

Private Sub TextDesc_GotFocus()
  MarcarTexto TextDesc
End Sub

Private Sub TextDesc_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextDesc1_GotFocus()
  MarcarTexto TextDesc1
End Sub

Private Sub TextDesc1_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextEntrada_GotFocus()
  MarcarTexto TextEntrada
  If OpcI.value Then OpcDH = 1 Else OpcDH = 2
  If OpcE.value Then
     If Precio = 0 Then
        MsgBox "Warning: " & vbCrLf _
             & "        Falta de Ingresar en este codigo" & vbCrLf _
             & "        " & Codigo & ": La entrada inicial"
     End If
  End If
  TextVUnit.Text = Format(ValorUnit, "#,##0.00000")
End Sub

Private Sub TextEntrada_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextEntrada_LostFocus()
  TextoValido TextEntrada, True, , Dec_Cant
  Entrada = Val(TextEntrada)
  ValorTotal = Redondear(ValorUnit * Entrada, 2)
  TextTotal.Text = Format(ValorTotal, "#,##0.0000")
End Sub

Private Sub Form_Activate()
 Opciones = ReadSetDataNum("PorCodigo", True, False)
'Consultamos las cuentas de la tabla
 Trans_No = 97
 IniciarAsientosAdo AdoAsientos
 
 DGKardex.Height = MDI_Y_Max - DGKardex.Top - 950
 DGKardex.width = MDI_X_Max - DGKardex.Left
 Label3.Top = DGKardex.Top + DGKardex.Height + 50
 Label10.Top = DGKardex.Top + DGKardex.Height + 50
 Label11.Top = DGKardex.Top + DGKardex.Height + 50
 Label16.Top = DGKardex.Top + DGKardex.Height + 50
 Label18.Top = DGKardex.Top + DGKardex.Height + 50
 Label1.Top = Label11.Top + Label11.Height
 TxtDifxDec.Top = Label11.Top + Label11.Height
 TxtSubTotal.Top = Label11.Top + Label11.Height
 TextIVA.Top = Label11.Top + Label11.Height
 Command1.Top = DGKardex.Top + DGKardex.Height + 50
 Command2.Top = DGKardex.Top + DGKardex.Height + 50
 Command3.Top = DGKardex.Top + DGKardex.Height + 50
 
' AdoDetKardex.Top = DGQuery.Top + DGQuery.Height + 50

 
 If Inv_Promedio Then
    Kard_Ing.Caption = "CONTROL DE INVENTARIO EN PROMEDIO"
 Else
    Kard_Ing.Caption = "CONTROL DE INVENTARIO EN ULTIMO PRECIO"
 End If
 TipoDoc = CompDiario
 sSQL = "SELECT Codigo_Inv & '  ' & Producto As NomProd " _
      & "FROM Catalogo_Productos " _
      & "WHERE Item = '" & NumEmpresa & "' " _
      & "AND Periodo = '" & Periodo_Contable & "' " _
      & "AND TC = 'I' " _
      & "ORDER BY Codigo_Inv "
 SelectDB_Combo DCTInv, AdoTInv, sSQL, "NomProd"
 ListarProductos
 sSQL = "SELECT * " _
      & "FROM Asiento_K " _
      & "WHERE Item = '" & NumEmpresa & "' " _
      & "AND CodigoU = '" & CodigoUsuario & "' " _
      & "AND T_No = " & Trans_No & " "
 SQLDec = "VALOR_UNIT " & CStr(Dec_Costo) & "|,VALOR_TOTAL 4|,CANTIDAD 2|,SALDO 4|."
 Select_Adodc_Grid DGKardex, AdoKardex, sSQL, SQLDec
  
 sSQL = "SELECT Cuenta & SPACE(67-LEN(Cuenta)) & Codigo As Nomb_Cta " _
      & "FROM Catalogo_Cuentas " _
      & "WHERE Item = '" & NumEmpresa & "' " _
      & "AND Periodo = '" & Periodo_Contable & "' " _
      & "AND TC IN ('RP','C','P','HC','I','G') " _
      & "AND DG = 'D' " _
      & "ORDER BY Codigo,Cuenta "
 SelectDB_Combo DCCtaObra, AdoCtaObra, sSQL, "Nomb_Cta"
 
 Contra_Cta = SinEspaciosDer(DCCtaObra)
 Codigo = Leer_Cta_Catalogo(Contra_Cta)
 ListarProveedorUsuario SubCta
 
 sSQL = "SELECT Numero " _
      & "FROM Trans_Kardex " _
      & "WHERE Item = '" & NumEmpresa & "' " _
      & "AND Periodo = '" & Periodo_Contable & "' " _
      & "AND TP = 'CD' " _
      & "AND Entrada > 0 " _
      & "GROUP BY Numero " _
      & "ORDER BY Numero DESC "
 SelectDB_Combo DCDiario, AdoDiario, sSQL, "Numero"
 
 sSQL = "SELECT * " _
      & "FROM Catalogo_Bodegas " _
      & "WHERE Item = '" & NumEmpresa & "' " _
      & "AND Periodo = '" & Periodo_Contable & "' " _
      & "ORDER BY Bodega "
 SelectDB_Combo DCBodega, AdoBodega, sSQL, "Bodega"
 
 sSQL = "SELECT * " _
      & "FROM Catalogo_Marcas " _
      & "WHERE Item = '" & NumEmpresa & "' " _
      & "AND Periodo = '" & Periodo_Contable & "' " _
      & "ORDER BY Marca "
 SelectDB_Combo DCMarca, AdoMarca, sSQL, "Marca"
 
 FechaValida MBFechaI
 FechaComp = MBFechaI
 RatonNormal
 CLTP.Clear
 CLTP.AddItem "CD"
 CLTP.AddItem "NC"
 CLTP.AddItem "ND"
 CLTP.Text = "CD"
 CLTP.SetFocus
End Sub

Private Sub Form_Load()
  CodigoBenef = Ninguno
  CodigoCliente = Ninguno
  ConectarAdodc AdoAux
  ConectarAdodc AdoInv
  ConectarAdodc AdoRet
  ConectarAdodc AdoIVA
  ConectarAdodc AdoArt
  ConectarAdodc AdoTInv
  ConectarAdodc AdoBenef
  ConectarAdodc AdoMarca
  ConectarAdodc AdoDiario
  ConectarAdodc AdoKardex
  ConectarAdodc AdoBodega
  ConectarAdodc AdoPorcIVA
  ConectarAdodc AdoCtaObra
  ConectarAdodc AdoAsientos
End Sub

Private Sub TextIVA_GotFocus()
  TextIVA.Text = ""
  TotalInventario
End Sub

Private Sub TextIVA_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextIVA_LostFocus()
  TextoValido TextIVA, True
  TotalInventario
End Sub

Private Sub TextOrden_GotFocus()
  SaldoAnterior = 0: ValorUnitAnt = 0: Contador = 0: ValorUnit = 0: Cantidad = 0
  Stock_Actual_Inventario MBFechaI, CodigoInv
  Precio = ValorUnit
  TextVUnit = Format(Precio, "#,##0." & String(Dec_Costo, "0"))
  MarcarTexto TextOrden
End Sub

Private Sub TextOrden_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextOrden_LostFocus()
  TextoValido TextOrden, , True
End Sub

Private Sub TextTotal_GotFocus()
Dim DValorUnit As Double
Dim DValorTotal As Double
   TextoValido TextOrden
   Cod_Bodega = Ninguno
   Cod_Marca = Ninguno
   With AdoBodega.Recordset
    If .RecordCount > 0 Then
       .MoveFirst
       .Find ("Bodega Like '" & DCBodega & "' ")
        If Not .EOF Then Cod_Bodega = .fields("CodBod")
    End If
   End With
   With AdoMarca.Recordset
    If .RecordCount > 0 Then
       .MoveFirst
       .Find ("Marca Like '" & DCMarca & "' ")
        If Not .EOF Then Cod_Marca = .fields("CodMar")
    End If
   End With
   
   Entrada = Val(CCur(TextEntrada))
   ValorUnit = Val(CDbl(TextVUnit))
   DValorUnit = ValorUnit
   Total_Desc = Val(CCur(TextDesc) / 100)
   DValorUnit = DValorUnit - (DValorUnit * Total_Desc)
   Total_Desc = Val(CCur(TextDesc1) / 100)
   If Total_Desc > 0 Then DValorUnit = DValorUnit - (DValorUnit * Total_Desc)
   DValorTotal = DValorUnit * Entrada
   ValorUnit = Val(CDbl(DValorUnit))
   ValorTotal = Redondear(Val(CCur(DValorTotal)), 2)
   TextTotal = Format(ValorTotal, "#,##0.0000")
End Sub

Private Sub TextTotal_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextTotal_LostFocus()
   FechaTexto = MBFechaI
   Entrada = Val(CCur(TextEntrada))
   Factura_No = Val(TxtFactNo)
   If Factura_No <= 0 Then Factura_No = 0 'Val(Format(Year(FechaTexto), "0000") & Format(Month(FechaTexto), "00") & Format(Day(FechaTexto), "00"))
   TxtFactNo = Factura_No
 ' Llenamos el ultimo saldo del kardex
   TextVUnit = Format(ValorUnit, "#,##0." & String(Dec_Costo, "0"))
   SaldoAnterior = Redondear(SaldoAnterior, 4)
   SubTotal_IVA = 0
   If OpcIVA.value Then SubTotal_IVA = Redondear(ValorTotal * Porc_IVA, 4)
   If OpcI.value Then
      Cantidad = Cantidad + Entrada
      Saldo = SaldoAnterior + ValorTotal
      OpcDH = "1"
      Contra_Cta = SinEspaciosDer(DCCtaObra)
   Else
      Cantidad = Cantidad - Entrada
      Saldo = SaldoAnterior - ValorTotal
      OpcDH = "2"
      If CheqContraCta.value = 1 Then
         Contra_Cta = SinEspaciosDer(DCCtaObra)
      Else
         Contra_Cta = Contra_Cta1
      End If
   End If
   Codigo = Leer_Cta_Catalogo(Contra_Cta)
   If Entrada > 0 And ValorUnit > 0 Then
      SetAddNew AdoKardex
      SetFields AdoKardex, "DH", OpcDH
      SetFields AdoKardex, "CODIGO_INV", CodigoInv
      SetFields AdoKardex, "P_DESC", Val(CCur(TextDesc))
      SetFields AdoKardex, "P_DESC1", Val(CCur(TextDesc1))
      SetFields AdoKardex, "PRODUCTO", Producto
      SetFields AdoKardex, "CANT_ES", Entrada
      SetFields AdoKardex, "VALOR_UNIT", ValorUnit
      SetFields AdoKardex, "VALOR_TOTAL", ValorTotal
      SetFields AdoKardex, "IVA", SubTotal_IVA
      SetFields AdoKardex, "CTA_INVENTARIO", Cta_Inventario
      SetFields AdoKardex, "CONTRA_CTA", Contra_Cta
      SetFields AdoKardex, "CANTIDAD", Cantidad
      SetFields AdoKardex, "VALOR_UNIT", ValorUnit
      SetFields AdoKardex, "SALDO", Saldo
      SetFields AdoKardex, "UNIDAD", LabelUnidad.Caption
      SetFields AdoKardex, "CodBod", Cod_Bodega
      SetFields AdoKardex, "CodMar", Cod_Marca
      SetFields AdoKardex, "Item", NumEmpresa
      SetFields AdoKardex, "CodigoU", CodigoUsuario
      SetFields AdoKardex, "T_No", Trans_No
      SetFields AdoKardex, "SUBCTA", SubCtaGen
      SetFields AdoKardex, "TC", SubCta
      SetFields AdoKardex, "Codigo_B", CodigoCliente
      SetFields AdoKardex, "COD_BAR", TxtCodBar
      SetFields AdoKardex, "ORDEN", TextOrden
      SetFields AdoKardex, "Lote_No", TxtLoteNo
      SetFields AdoKardex, "Fecha_Fab", MBFechaFab
      SetFields AdoKardex, "Fecha_Exp", MBFechaExp
      SetFields AdoKardex, "Reg_Sanitario", TxtRegSanitario
      SetFields AdoKardex, "Modelo", TxtModelo
      SetFields AdoKardex, "Procedencia", TxtProcedencia
      SetFields AdoKardex, "Serie_No", TxtSerieNo
      SetUpdate AdoKardex
      TextTotal = Format(ValorTotal, "#,##0.00")
      TotalInventario
      DCInv.SetFocus
   End If
End Sub

Private Sub TextVUnit_GotFocus()
  MarcarTexto TextVUnit
End Sub

Private Sub TextVUnit_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextVUnit_LostFocus()
  'MsgBox Dec_Costo
  TextoValido TextVUnit, True, , Dec_Costo
  ValorUnit = Val(CDbl(TextVUnit))
  
  ValorTotal = Redondear(ValorUnit * Entrada, 2)
  TextTotal = Format(ValorTotal, "#,##0.0000")
  TextVUnit = FormatDec(ValorUnit, Dec_Costo)
End Sub

Public Sub TotalInventario()
Dim TotalInvs As Currency
  Total = 0: Total_IVA = 0
  With AdoKardex.Recordset
    If .RecordCount > 0 Then
       .MoveFirst
        Do While Not .EOF
           Total = Total + .fields("VALOR_TOTAL")
           Total_IVA = Total_IVA + .fields("IVA")
          .MoveNext
        Loop
       .MoveFirst
    End If
  End With
  Total = Redondear(Total, 2)
  Total_IVA = Redondear(Total_IVA, 2)
  TxtSubTotal.Text = Format(Total, "#,##0.00")
  TextIVA.Text = Format(Total_IVA, "#,##0.00")
  Label1.Caption = Format(Total + Total_IVA, "#,##0.00")
End Sub

Public Sub ListarProductos()
 Opciones = ReadSetDataNum("PorCodigo", True, False)
 CodigoInv = SinEspaciosIzq(DCTInv.Text)
 sSQL = "SELECT *,(Codigo_Inv & '  ' & Producto) As NomProd " _
      & "FROM Catalogo_Productos " _
      & "WHERE Item = '" & NumEmpresa & "' " _
      & "AND Periodo = '" & Periodo_Contable & "' " _
      & "AND MidStrg(Codigo_Inv,1," & CStr(Len(CodigoInv)) & ") = '" & CodigoInv & "' " _
      & "AND LEN(Cta_Inventario) > 1 " _
      & "AND TC = 'P' "
 If Opciones = 1 Then
    sSQL = sSQL & "ORDER BY Producto "
    SelectDB_Combo DCInv, AdoInv, sSQL, "Producto"
 Else
    sSQL = sSQL & "ORDER BY Codigo_Inv "
    SelectDB_Combo DCInv, AdoInv, sSQL, "NomProd"
 End If
End Sub

Public Sub ListarProveedorUsuario(TipoSubCta As String)
  Select Case TipoSubCta
    Case "RP"
         'OpcIVA.Visible = False
         'OpcX.Visible = False
         OpcX.value = 1
         sSQL = "SELECT C.Cliente,CR.Codigo,C.CI_RUC,C.Direccion,C.Telefono,C.TD,'.' As Cta,0 As Importaciones,C.CI_RUC,C.Grupo,'P' As TipoBenef " _
              & "FROM Clientes As C,Catalogo_Rol_Pagos As CR " _
              & "WHERE CR.Item = '" & NumEmpresa & "' " _
              & "AND CR.Periodo = '" & Periodo_Contable & "' " _
              & "AND C.Codigo = CR.Codigo " _
              & "ORDER BY C.Cliente "
    Case "P", "C"
         'OpcIVA.Visible = True
         'OpcX.Visible = True
         sSQL = "SELECT C.Cliente,C.Codigo,C.CI_RUC,C.Direccion,C.Telefono,C.TD,CP.Cta,CP.Importaciones,C.CI_RUC,C.Grupo,'P' As TipoBenef " _
              & "FROM Clientes As C,Catalogo_CxCxP As CP " _
              & "WHERE CP.TC = '" & TipoSubCta & "' " _
              & "AND CP.Item = '" & NumEmpresa & "' " _
              & "AND CP.Periodo = '" & Periodo_Contable & "' " _
              & "AND CP.Cta = '" & Contra_Cta & "' " _
              & "AND C.Codigo = CP.Codigo " _
              & "GROUP BY C.Cliente,C.Codigo,C.CI_RUC,C.Direccion,C.Telefono,C.TD,CP.Cta,CP.Importaciones,C.CI_RUC,C.Grupo " _
              & "ORDER BY C.Cliente "
    Case "G", "I"
         'OpcIVA.Visible = False
         'OpcX.Visible = False
         OpcX.value = 1
         sSQL = "SELECT CS.Detalle As Cliente,CS.*,'O' As TD,'.' As Cta,0 As Importaciones,'9999999999999' As CI_RUC,'999999' As Grupo,'X' As TipoBenef " _
              & "FROM Catalogo_SubCtas As CS " _
              & "WHERE CS.Item = '" & NumEmpresa & "' " _
              & "AND CS.Periodo = '" & Periodo_Contable & "' " _
              & "AND CS.TC = '" & TipoSubCta & "' " _
              & "ORDER BY Detalle "
    Case Else
         sSQL = "SELECT Cliente,Codigo,CI_RUC,Grupo,Direccion,Telefono,TD,'.' As Cta,0 As Importaciones,'X' As TipoBenef " _
              & "FROM Clientes " _
              & "WHERE Grupo = '.' " _
              & "ORDER BY Cliente "
  End Select
  SelectDB_Combo DCBenef, AdoBenef, sSQL, "Cliente"
  If AdoBenef.Recordset.RecordCount <= 0 Then
     MsgBox "No existen Proveedores Asignados a Esta Cuenta"
     MBFechaI.SetFocus
  End If
End Sub

Private Sub TxtCodBar_GotFocus()
  MarcarTexto TxtCodBar
End Sub

Private Sub TxtCodBar_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtCodBar_LostFocus()
  TextoValido TxtCodBar, , True
End Sub

Private Sub TxtDifxDec_GotFocus()
  MarcarTexto TxtDifxDec
End Sub

Private Sub TxtDifxDec_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtDifxDec_LostFocus()
  TextoValido TxtDifxDec, True
End Sub

Private Sub TxtFactNo_GotFocus()
  MarcarTexto TxtFactNo
End Sub

Private Sub TxtFactNo_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtFactNo_LostFocus()
  TextoValido TxtFactNo, , True
  Factura_No = Val(TxtFactNo)
  FechaTexto = MBFechaI
  If Factura_No < 0 Then Factura_No = 0
  TxtFactNo = Factura_No
End Sub

Private Sub TxtLoteNo_GotFocus()
  MarcarTexto TxtLoteNo
End Sub

Private Sub TxtLoteNo_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtLoteNo_LostFocus()
  TextoValido TxtLoteNo, , True
End Sub

Private Sub TxtModelo_GotFocus()
  MarcarTexto TxtModelo
End Sub

Private Sub TxtModelo_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtModelo_LostFocus()
  TextoValido TxtModelo, , True
End Sub

Private Sub TxtProcedencia_GotFocus()
  MarcarTexto TxtProcedencia
End Sub

Private Sub TxtProcedencia_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtProcedencia_LostFocus()
  TextoValido TxtProcedencia, , True
End Sub

Private Sub TxtRegSanitario_GotFocus()
  MarcarTexto TxtRegSanitario
End Sub

Private Sub TxtRegSanitario_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtRegSanitario_LostFocus()
   TextoValido TxtRegSanitario, , True
End Sub

Private Sub TxtSerie_GotFocus()
  MarcarTexto TxtSerie
End Sub

Private Sub TxtSerie_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtSerie_LostFocus()
   TextoValido TxtSerie
End Sub

Private Sub TxtSerieNo_GotFocus()
  MarcarTexto TxtSerieNo
End Sub

Private Sub TxtSerieNo_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtSerieNo_LostFocus()
  TextoValido TxtSerieNo, , True
End Sub
