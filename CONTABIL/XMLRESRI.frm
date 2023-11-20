VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.5#0"; "comctl32.Ocx"
Begin VB.Form FXMLRecibidosSRI 
   BackColor       =   &H80000002&
   Caption         =   "ftpLinode"
   ClientHeight    =   9450
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   17460
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9450
   ScaleWidth      =   17460
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   795
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   17460
      _ExtentX        =   30798
      _ExtentY        =   1402
      ButtonWidth     =   1455
      ButtonHeight    =   1244
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   4
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Salir"
            Object.ToolTipText     =   "Salir del Módulo"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Leer_XML"
            Object.ToolTipText     =   "Importar documentos electronicos"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Excel"
            Object.ToolTipText     =   "Resumen por Código de Retencion"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Grabar"
            Object.ToolTipText     =   "Resumen Detalle por Código de Retención"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.PictureBox PctSRI 
         BorderStyle     =   0  'None
         Height          =   540
         Left            =   3360
         ScaleHeight     =   540
         ScaleWidth      =   10725
         TabIndex        =   16
         Top             =   105
         Width           =   10725
      End
   End
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
      Left            =   12285
      TabIndex        =   14
      Top             =   2835
      Visible         =   0   'False
      Width           =   5685
   End
   Begin MSDataGridLib.DataGrid DGDocSRI 
      Bindings        =   "XMLRESRI.frx":0000
      Height          =   2430
      Left            =   105
      TabIndex        =   11
      Top             =   2100
      Width           =   12090
      _ExtentX        =   21325
      _ExtentY        =   4286
      _Version        =   393216
      AllowUpdate     =   -1  'True
      BackColor       =   16761024
      BorderStyle     =   0
      Enabled         =   0   'False
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
   Begin MSAdodcLib.Adodc AdoAux 
      Height          =   330
      Left            =   525
      Top             =   5565
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
   Begin MSAdodcLib.Adodc AdoAbono 
      Height          =   330
      Left            =   525
      Top             =   5145
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
   Begin MSAdodcLib.Adodc AdoClientes 
      Height          =   330
      Left            =   525
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
   Begin MSAdodcLib.Adodc AdoDocSRI 
      Height          =   330
      Left            =   5460
      Top             =   1680
      Width           =   2745
      _ExtentX        =   4842
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
      Caption         =   "DocSRI"
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
   Begin MSAdodcLib.Adodc AdoCxP 
      Height          =   330
      Left            =   525
      Top             =   5985
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
      Caption         =   "CxP"
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
      Left            =   1050
      Top             =   3990
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      PrinterDefault  =   0   'False
   End
   Begin MSAdodcLib.Adodc AdoAsiento 
      Height          =   330
      Left            =   525
      Top             =   6405
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
   Begin MSDataListLib.DataCombo DCCxP 
      Bindings        =   "XMLRESRI.frx":0018
      DataSource      =   "AdoCxP"
      Height          =   345
      Left            =   10290
      TabIndex        =   3
      Top             =   840
      Width           =   6105
      _ExtentX        =   10769
      _ExtentY        =   609
      _Version        =   393216
      Text            =   ""
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
   Begin MSDataListLib.DataCombo DCSerieRetencion 
      Bindings        =   "XMLRESRI.frx":002D
      DataSource      =   "AdoSerieRetencion"
      Height          =   345
      Left            =   2100
      TabIndex        =   9
      Top             =   1680
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   609
      _Version        =   393216
      Text            =   ""
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
   Begin MSAdodcLib.Adodc AdoSerieRetencion 
      Height          =   330
      Left            =   2730
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
      Caption         =   "SerieRetencion"
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
   Begin MSDataListLib.DataCombo DCSustento 
      Bindings        =   "XMLRESRI.frx":004D
      DataSource      =   "AdoSustento"
      Height          =   360
      Left            =   2100
      TabIndex        =   5
      ToolTipText     =   "En este campo de selección se despliega un lista de tipos de sustentos tributarios correspondientes a la transacción escogida"
      Top             =   1260
      Width           =   6105
      _ExtentX        =   10769
      _ExtentY        =   635
      _Version        =   393216
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSAdodcLib.Adodc AdoSustento 
      Height          =   330
      Left            =   2730
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
      Caption         =   "Sustento"
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
      Height          =   330
      Left            =   16380
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   840
      Width           =   435
   End
   Begin MSDataListLib.DataCombo DCCtaGasto 
      Bindings        =   "XMLRESRI.frx":0067
      DataSource      =   "AdoCtaGasto"
      Height          =   345
      Left            =   2100
      TabIndex        =   1
      Top             =   840
      Width           =   6105
      _ExtentX        =   10769
      _ExtentY        =   609
      _Version        =   393216
      Text            =   ""
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
   Begin MSDataListLib.DataCombo DCTipoPago 
      Bindings        =   "XMLRESRI.frx":0081
      DataSource      =   "AdoTipoPago"
      Height          =   360
      Left            =   10290
      TabIndex        =   7
      ToolTipText     =   "En este campo de selección se despliega un lista de tipos de sustentos tributarios correspondientes a la transacción escogida"
      Top             =   1260
      Width           =   6105
      _ExtentX        =   10769
      _ExtentY        =   635
      _Version        =   393216
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ComctlLib.ListView LstVwFTP 
      Height          =   645
      Left            =   12285
      TabIndex        =   15
      Top             =   2100
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
   Begin MSAdodcLib.Adodc AdoTxt 
      Height          =   330
      Left            =   2730
      Top             =   5460
      Visible         =   0   'False
      Width           =   2145
      _ExtentX        =   3784
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
      Caption         =   "Txt"
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
   Begin MSAdodcLib.Adodc AdoTipoPago 
      Height          =   330
      Left            =   2730
      Top             =   6300
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
      Caption         =   "TipoPago"
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
   Begin MSAdodcLib.Adodc AdoCtaGasto 
      Height          =   330
      Left            =   2730
      Top             =   5880
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
      Caption         =   "CtaGasto"
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
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Tipo de Pago"
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
      Left            =   8295
      TabIndex        =   6
      Top             =   1260
      Width           =   2010
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Cuenta de Gasto"
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
      Top             =   840
      Width           =   2010
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Tipo de Sustento"
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
      Top             =   1260
      Width           =   2010
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "9999999999"
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
      Left            =   4200
      TabIndex        =   10
      Top             =   1680
      Width           =   1170
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Retencion No."
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
      TabIndex        =   8
      Top             =   1680
      Width           =   2010
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Cuenta de Proveedor"
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
      Left            =   8295
      TabIndex        =   2
      Top             =   840
      Width           =   2010
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   315
      Top             =   3885
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   4
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "XMLRESRI.frx":009B
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "XMLRESRI.frx":03B5
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "XMLRESRI.frx":0BCF
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "XMLRESRI.frx":1821
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FXMLRecibidosSRI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Cta_Prov As String
Dim Cta_Prov_Aut As String
Dim CodRet As String

''Dim Cta_D As String
''Dim Cta_C As String

Dim IColIndex As Integer

Dim CodRetBien As Byte
Dim CodRetServ As Byte

Dim SecRetencion As Long

Dim Porcentaje As Single

Private Sub Leer_Porc_Retenciones(TipoRetencion As String)
Dim AdoDBTemp As ADODB.Recordset
    
    If Len(AXML.Cod_Ret) > 1 Then
       sSQL = "SELECT TOP 1 Porcentaje " _
            & "FROM Tipo_Concepto_Retencion " _
            & "WHERE Codigo = '" & AXML.Cod_Ret & "' " _
            & "AND Fecha_Inicio <= #" & BuscarFecha(AXML.Fecha_Emision) & "# " _
            & "AND Fecha_Final >= #" & BuscarFecha(AXML.Fecha_Emision) & "# "
       Select_AdoDB AdoDBTemp, sSQL
       If AdoDBTemp.RecordCount > 0 Then AXML.Porc_Ret = AdoDBTemp.fields("Porcentaje")
       AdoDBTemp.Close
    End If
    
    sSQL = "SELECT TC, Codigo, Cuenta " _
         & "FROM Catalogo_Cuentas " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND DG = 'D' "
    If TipoRetencion = "1" Then
       sSQL = sSQL & "AND TC IN ('CF','CB','CI') "
    Else
       sSQL = sSQL & "AND TC IN ('RF','RB','RI') "
    End If
    sSQL = sSQL & "ORDER BY TC, Cuenta "
    Select_AdoDB AdoDBTemp, sSQL
    With AdoDBTemp
     If .RecordCount > 0 Then
         Do While Not .EOF
            Cuenta = .fields("Cuenta")
            If InStr(Cuenta, CStr(AXML.Porc_Ret)) And AXML.Porc_Ret > 0 Then AXML.Cta_Ret_Fuente = .fields("Codigo")
            If InStr(Cuenta, CStr(AXML.Porc_Ret_IVA_B)) And AXML.Porc_Ret_IVA_B > 0 Then AXML.Cta_Ret_IVA_B = .fields("Codigo")
            If InStr(Cuenta, CStr(AXML.Porc_Ret_IVA_S)) And AXML.Porc_Ret_IVA_S > 0 Then AXML.Cta_Ret_IVA_S = .fields("Codigo")
           .MoveNext
         Loop
     End If
    End With
    AdoDBTemp.Close
    
    If TipoRetencion = "1" Then
        If AXML.Porc_Ret > 0 And AXML.Cta_Ret_Fuente = "" And Len(AXML.Cod_Ret) > 1 Then AXML.Cta_Ret_Fuente = Cta_Ret
        If AXML.Porc_Ret_IVA_B > 0 And AXML.Cta_Ret_IVA_B = "" Then AXML.Cta_Ret_IVA_B = Cta_Ret_IVA
        If AXML.Porc_Ret_IVA_S > 0 And AXML.Cta_Ret_IVA_S = "" Then AXML.Cta_Ret_IVA_S = Cta_Ret_IVA
    Else
        If AXML.Porc_Ret > 0 And AXML.Cta_Ret_Fuente = "" And Len(AXML.Cod_Ret) > 1 Then AXML.Cta_Ret_Fuente = Cta_Ret_Egreso
        If AXML.Porc_Ret_IVA_B > 0 And AXML.Cta_Ret_IVA_B = "" Then AXML.Cta_Ret_IVA_B = Cta_Ret_IVA_Egreso
        If AXML.Porc_Ret_IVA_S > 0 And AXML.Cta_Ret_IVA_S = "" Then AXML.Cta_Ret_IVA_S = Cta_Ret_IVA_Egreso
    End If
    
    If AXML.Cod_Ret = "" Then AXML.Cod_Ret = Ninguno
    If AXML.Cta_Debito = "" Then AXML.Cta_Debito = "0"
    If AXML.Cta_Credito = "" Then AXML.Cta_Credito = "0"
    If AXML.Cta_IVA_Gasto = "" Then AXML.Cta_IVA_Gasto = "0"
    If AXML.Cta_Ret_Fuente = "" Then AXML.Cta_Ret_Fuente = "0"
    If AXML.Cta_Ret_IVA_B = "" Then AXML.Cta_Ret_IVA_B = "0"
    If AXML.Cta_Ret_IVA_S = "" Then AXML.Cta_Ret_IVA_S = "0"
    
End Sub
'Sube los abonos
Private Sub Leer_XML()
On Error GoTo Errorhandler

Dim ArchivoValido As Boolean
Dim ReceptorValido As Boolean

Dim ClaveAcceso As String
Dim RutaXMLAutorizado As String
Dim RutaXMLRechazado As String
   
   FechaTexto = FechaSistema
   Cta_Prov_Aut = SinEspaciosIzq(DCCxP.Text)
   PrimeraLinea = True
   ReceptorValido = False
   ArchivoValido = False
   
   TextoImprimio = ""
   TextoFile = "Asiento_TXT_" & CodigoUsuario
   
   sSQL = "DELETE * " _
        & "FROM Tabla_Temporal " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND Modulo = '" & NumModulo & "' " _
        & "AND CodigoU = '" & CodigoUsuario & "' "
   Ejecutar_SQL_SP sSQL
   
  TextoImprimio = ""
  RutaSubDirTemp = RutaSysBases & "\SRI\Comprobantes Recibidos\*.xml"
  CodigoP = RutaSysBases & "\SRI\*.txt"
  CDialogDir.DialogTitle = "Abrir Archivo"
  CDialogDir.InitDir = RutaSysBases & "\SRI\"
  CDialogDir.Filename = CodigoP
  CDialogDir.Flags = cdlOFNFileMustExist + cdlOFNNoChangeDir + cdlOFNHideReadOnly
  CDialogDir.Filter = "Archivos TXT|*.txt"
  CDialogDir.FilterIndex = 2
  CDialogDir.DefaultExt = "txt"
  CDialogDir.CancelError = False
  CDialogDir.ShowOpen
  If CodigoP <> CDialogDir.Filename Then NombreArchivo = CDialogDir.Filename Else NombreArchivo = ""
  If NombreArchivo <> "" Then
     Toolbar1.buttons("Salir").Enabled = False
     Progreso_Barra.Mensaje_Box = "SUBIENDO ARCHIVOS XML DEL SRI "
     Progreso_Iniciar
     
     DGDocSRI.Visible = False
     DGDocSRI.Caption = NombreArchivo
     
     Subir_Archivo_FTP_Linode ftp, LstStatud, LstVwFTP, NombreArchivo
     Subir_Archivo_TXT_SP NombreArchivo
     Eliminar_Archivo_FTP_Linode ftp, LstStatud, LstVwFTP, NombreArchivo

    'Determinamos cantidad de archivos a subir
     Contador = 1
     Progreso_Barra.Valor_Maximo = 0
     
     sSQL = "SELECT IDENTIFICACION_RECEPTOR " _
          & "FROM " & TextoFile & " " _
          & "WHERE COMPROBANTE <> '.' " _
          & "GROUP BY IDENTIFICACION_RECEPTOR "
     Select_Adodc AdoTxt, sSQL
     If AdoTxt.Recordset.RecordCount > 0 Then
        Do While Not AdoTxt.Recordset.EOF
           If RUC = AdoTxt.Recordset.fields("IDENTIFICACION_RECEPTOR") Then ReceptorValido = True
           AdoTxt.Recordset.MoveNext
        Loop
     End If
     
     If ReceptorValido Then
        ArchivoValido = True
        sSQL = "SELECT " & Full_Fields(TextoFile) & " " _
             & "FROM " & TextoFile & " " _
             & "WHERE COMPROBANTE <> '.' " _
             & "ORDER BY COMPROBANTE, IDENTIFICACION_RECEPTOR, FECHA_EMISION "
        Select_Adodc AdoTxt, sSQL
        With AdoTxt.Recordset
         If .RecordCount > 0 Then
             Progreso_Barra.Valor_Maximo = .RecordCount
             TotalIngreso = 0
             FileResp = 0
            'Establecemos los campos del archivo plano del Banco
             FechaTexto = FechaSistema
             Do While Not .EOF
                ClaveAcceso = .fields("CLAVE_ACCESO")
                ID_Reg = .fields("ID")
                If Len(ClaveAcceso) = 49 Then
                   RutaXMLAutorizado = RutaSysBases & "\SRI\Comprobantes Recibidos\" & ClaveAcceso & ".xml"
                   RutaXMLRechazado = RutaSysBases & "\SRI\Comprobantes no Autorizados\" & ClaveAcceso & ".xml"
                  'MsgBox RutaXMLAutorizado
                   If Not Existe_File(RutaXMLAutorizado) Then
                      SRI_Autorizacion = SRI_Leer_XML_Autorizado(RutaXMLAutorizado, RutaXMLRechazado)
                      TextoFileEmp = SRI_Autorizacion.Documento_XML
                      I = InStr(TextoFileEmp, "<![CDATA[")
                      F = InStr(TextoFileEmp, "]]></comprobante>")
                      If I > 0 And F > 0 Then
                         I = I + 9
                         Escribir_Archivo RutaXMLAutorizado, TrimStrg(MidStrg(TextoFileEmp, I, F - I))
                      End If
                   End If
                  'Procedemos a leer la informacion del Documento
                   AXML.Documento = Ninguno
                   AXML.Razon_Social_Emisor = Ninguno
                   AXML.RUC_Emisor = Ninguno
                   AXML.Direccion_Emisor = Ninguno
                   AXML.Fecha_Emision = Ninguno
                   AXML.Serie = Ninguno
                   AXML.Autorizacion = Ninguno
                   AXML.Cod_Ret = Ninguno
                   AXML.Serie_Receptor = Ninguno
                   AXML.Cta_Debito = Ninguno
                   AXML.Cta_Credito = Ninguno
                   AXML.Cta_IVA_Gasto = Ninguno
                   AXML.Cta_Ret_Fuente = Ninguno
                   AXML.Cta_Ret_IVA_B = Ninguno
                   AXML.Cta_Ret_IVA_S = Ninguno
                   AXML.SubModulo = Ninguno
                   AXML.Codigo_B = Ninguno
                   AXML.Ambiente = 0
                   AXML.CodPorIva = 0
                   AXML.Comprobante = 0
                   AXML.SubTotal = 0
                   AXML.Total_IVA = 0
                   AXML.Total = 0
                   AXML.Ret_IVA_B = 0
                   AXML.Ret_IVA_S = 0
                   AXML.Ret_Fuente = 0
                   AXML.Porc_Ret = 0
                   AXML.Porc_Ret_IVA_B = 0
                   AXML.Porc_Ret_IVA_S = 0
                   AXML.Cod_Ret_IVA_B = 0
                   AXML.Cod_Ret_IVA_S = 0
        
                  'Recolectamos la informacion del documento electronico recibido y lo insertamos en la tabla
                   Progreso_Barra.Mensaje_Box = "Procesando: " & RutaXMLAutorizado
                   Progreso_Esperar
                    
                   Leer_Archivo_XML RutaXMLAutorizado
                   
                   DigVerif = Digito_Verificador(AXML.RUC_Emisor)
                   CodigoCli = Tipo_RUC_CI.Codigo_RUC_CI
                   If AXML.RUC_Emisor <> Ninguno And AXML.Ambiente = 2 And MidStrg(AXML.RUC_Receptor, 1, 10) = MidStrg(RUC, 1, 10) Then
                      If AXML.Documento = "Retencion" Then
                        '-----------------------------------------
                        'Retenciones emitidas por el Cliente
                        '-----------------------------------------
                         Leer_Porc_Retenciones "1"
                         FA.T = Normal
                         FA.TC = "FA"
                         FA.Serie = AXML.Serie_Receptor
                         FA.Factura = AXML.Comprobante
                         FA.Fecha = AXML.Fecha_Emision
                         AXML.Cta_Credito = Ninguno
                         sSQL = "SELECT Cta_CxP, Autorizacion " _
                              & "FROM Facturas " _
                              & "WHERE Item = '" & NumEmpresa & "' " _
                              & "AND Periodo = '" & Periodo_Contable & "' " _
                              & "AND T <> 'A' " _
                              & "AND TC = '" & FA.TC & "' " _
                              & "AND Serie = '" & FA.Serie & "' " _
                              & "AND Factura = " & FA.Factura & " "
                         Select_Adodc AdoAux, sSQL
                         If AdoAux.Recordset.RecordCount > 0 Then AXML.Cta_Credito = AdoAux.Recordset.fields("Cta_CxP")
                      Else
                        '-----------------------------------------
                        'Facturas de Proveedores/Facturas al Gasto
                        '-----------------------------------------
                         Eliminar_Nulos_SP "Catalogo_CxCxP"
                         AXML.Cta_Credito = Cta_Prov_Aut
                         sSQL = "SELECT TOP 1 Cta, Cta_Gasto, Cta_IVA_Gasto, SubModulo, Cod_Ret, Porc_IVAB, Porc_IVAS " _
                              & "FROM Catalogo_CxCxP " _
                              & "WHERE Item = '" & NumEmpresa & "' " _
                              & "AND Periodo = '" & Periodo_Contable & "' " _
                              & "AND Codigo = '" & CodigoCli & "' " _
                              & "AND Cta = '" & AXML.Cta_Credito & "' " _
                              & "AND TC = 'P' " _
                              & "ORDER BY Cta "
                         Select_Adodc AdoAux, sSQL
                         If AdoAux.Recordset.RecordCount > 0 Then
                            AXML.Cta_Debito = AdoAux.Recordset.fields("Cta_Gasto")
                            AXML.SubModulo = AdoAux.Recordset.fields("SubModulo")
                            AXML.Cod_Ret = AdoAux.Recordset.fields("Cod_Ret")
                            AXML.Porc_Ret_IVA_B = AdoAux.Recordset.fields("Porc_IVAB")
                            AXML.Porc_Ret_IVA_S = AdoAux.Recordset.fields("Porc_IVAS")
                            AXML.Cta_IVA_Gasto = AdoAux.Recordset.fields("Cta_IVA_Gasto")
                         Else
                            If Len(AXML.RUC_Receptor) = 13 Then
                               AXML.Cta_Debito = Cta_Gastos
                            Else
                               AXML.Cta_Debito = Cta_Gastos_Personales
                               AXML.SubTotal = AXML.SubTotal + AXML.Total_IVA
                               AXML.Total_IVA = 0
                            End If
                            SetAdoAddNew "Catalogo_CxCxP"
                            SetAdoFields "TC", "P"
                            SetAdoFields "Codigo", CodigoCli
                            SetAdoFields "Cta", AXML.Cta_Credito
                            SetAdoFields "Cta_Gasto", AXML.Cta_Debito
                            SetAdoUpdate
                         End If
                         Leer_Porc_Retenciones "2"
                      End If
        ''            If Cta_D = Ninguno Then Cta_D = Cta_Gastos
        ''            If Cta_C = Ninguno Then Cta_C = Cta_Prov_Aut
                     'If AXML.Autorizacion = "0510202107139172193500120010030000004531234567818" Then MsgBox AXML.Porc_Ret_IVA_S
                     
                     'Actualizamos los datos del cliente/proveedor
                      If AXML.Documento = "Factura" Then .fields("CodSustento") = SinEspaciosIzq(DCSustento)
                      .fields("PROCESAR") = 1
                      .fields("IMPORTE_TOTAL") = AXML.Total
                      .fields("Cod_Ret") = AXML.Cod_Ret
                      .fields("Serie_Receptor") = AXML.Serie_Receptor
                      .fields("Documento") = AXML.Comprobante
                      .fields("Subtotal") = AXML.SubTotal
                      .fields("Total_IVA") = AXML.Total_IVA
                      .fields("Ret_IVA_B") = AXML.Ret_IVA_B
                      .fields("Ret_IVA_S") = AXML.Ret_IVA_S
                      .fields("Ret_Fuente") = AXML.Ret_Fuente
                      .fields("Porc_Ret") = AXML.Porc_Ret
                      .fields("Porc_Ret_IVA_B") = AXML.Porc_Ret_IVA_B
                      .fields("Porc_Ret_IVA_S") = AXML.Porc_Ret_IVA_S
                      .fields("Cod_Ret_Bien") = AXML.Cod_Ret_IVA_B
                      .fields("Cod_Ret_Servicio") = AXML.Cod_Ret_IVA_S
                      .fields("Cta_Debito") = AXML.Cta_Debito
                      .fields("Cta_Credito") = AXML.Cta_Credito
                      .fields("Cta_IVA_Gasto") = AXML.Cta_IVA_Gasto
                      .fields("Cta_Ret_Fuente") = AXML.Cta_Ret_Fuente
                      .fields("Cta_Ret_IVA_B") = AXML.Cta_Ret_IVA_B
                      .fields("Cta_Ret_IVA_S") = AXML.Cta_Ret_IVA_S
                      .fields("CodPorIva") = AXML.CodPorIva
                      .fields("CodSustento") = AXML.Cod_Sustento
                      .Update
                       Contador = Contador + 1
                   Else
                      TextoImprimio = AXML.RUC_Receptor & " - " & TextoImprimio & RutaXMLAutorizado & vbCrLf
                      Insertar_Texto_Temporal_SP RutaXMLAutorizado
                   End If
                End If
               .MoveNext
             Loop
         End If
        End With
     Else
         TextoFile = "Asiento_TXT_" & CodigoUsuario
         sSQL = "DELETE " _
              & "FROM " & TextoFile & " " _
              & "WHERE COMPROBANTE <> '.' "
         Ejecutar_SQL_SP sSQL
         MsgBox "La informacion de este archivo no es para la Empresa. No se procede a procesar"
     End If
    'Si es el archivo correcto procedemosa subir los comprobantes
     If Not ArchivoValido Then MsgBox "ESTE ARCHIVO NO ES VALIDO, VUELVA A SUBIR"
     
     DGDocSRI.Visible = True
     TextoFile = "Asiento_TXT_" & CodigoUsuario
     
     sSQL = "SELECT " & Full_Fields(TextoFile) & " " _
          & "FROM " & TextoFile & " " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "ORDER BY COMPROBANTE, FECHA_EMISION, ID "
     Select_Adodc_Grid DGDocSRI, AdoDocSRI, sSQL
     DGDocSRI.Visible = True
     Progreso_Final
     RatonNormal
     MsgBox "Proceso terminado"
  Else
     Progreso_Final
     RatonNormal
     MsgBox "No se procesara ningun archivo"
  End If
  
  If Len(TextoImprimio) > 2 Then
     FXMLRecibidosSRI.WindowState = vbMaximized
     FInfoError.Show
  End If
 
 'Activamos lo botones una vez subida la informacion
  Toolbar1.buttons("Salir").Enabled = True
  Toolbar1.buttons("Excel").Enabled = True
  Toolbar1.buttons("Grabar").Enabled = True
  Exit Sub
Errorhandler:
    RatonNormal
    ErrorDeImpresion
    Exit Sub
End Sub

Public Sub Grabar_Comprobantes_XML()
    Progreso_Barra.Mensaje_Box = "Generando Abonos o Comprobantes"
    Progreso_Iniciar
    
    TextoFile = "Asiento_TXT_" & CodigoUsuario

    sSQL = "SELECT " & Full_Fields(TextoFile) & " " _
         & "FROM " & TextoFile & " " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Procesar <> 0 " _
         & "ORDER BY COMPROBANTE, FECHA_EMISION, ID "
    Select_Adodc AdoAux, sSQL
    
    With AdoAux.Recordset
     If .RecordCount > 0 Then
         Progreso_Barra.Valor_Maximo = .RecordCount
         Do While Not .EOF
            Progreso_Barra.Mensaje_Box = .fields("COMPROBANTE") & ": " & .fields("SERIE_COMPROBANTE") & ", " & .fields("RAZON_SOCIAL_EMISOR")
            Progreso_Esperar
            
            SecRetencion = Val(Label3.Caption)
'''            CodRetBien = 0
'''            CodRetServ = 0
'''            If .fields("Porc_Ret_IVA_B") > 0 Then
'''                sSQL = "SELECT Codigo " _
'''                     & "FROM Tabla_Por_IVA " _
'''                     & "WHERE Porc = '" & CStr(.fields("Porc_Ret_IVA_B")) & "' " _
'''                     & "AND Bienes <> 0 "
'''                Select_Adodc AdoAux, sSQL
'''                If AdoAux.Recordset.RecordCount > 0 Then CodRetBien = AdoAux.Recordset.fields("Codigo")
'''            End If
'''            If .fields("Porc_Ret_IVA_S") > 0 Then
'''                sSQL = "SELECT Codigo " _
'''                     & "FROM Tabla_Por_IVA " _
'''                     & "WHERE Porc = '" & CStr(.fields("Porc_Ret_IVA_S")) & "' " _
'''                     & "AND Servicios <> 0 "
'''                Select_Adodc AdoAux, sSQL
'''                If AdoAux.Recordset.RecordCount > 0 Then CodRetServ = AdoAux.Recordset.fields("Codigo")
'''            End If

            CodRetBien = AdoAux.Recordset.fields("Cod_Ret_Bien")
            CodRetServ = AdoAux.Recordset.fields("Cod_Ret_Servicio")

            Select Case .fields("COMPROBANTE")
              Case "Retencion"
                   TA.T = Normal
                   TA.CodigoC = Ninguno
                   TA.TP = "FA"
                   TA.Serie = .fields("Serie_Receptor")
                   TA.Factura = .fields("Documento")
                   TA.Fecha = .fields("FECHA_EMISION")
                   TA.Cta_CxP = .fields("Cta_Credito")
            
                   sSQL = "SELECT CodigoC, Autorizacion " _
                        & "FROM Facturas " _
                        & "WHERE Item = '" & NumEmpresa & "' " _
                        & "AND Periodo = '" & Periodo_Contable & "' " _
                        & "AND T = '" & Pendiente & "' " _
                        & "AND TC = '" & TA.TP & "' " _
                        & "AND Serie = '" & TA.Serie & "' " _
                        & "AND Factura = " & TA.Factura & " "
                   Select_Adodc AdoAbono, sSQL
                   If AdoAbono.Recordset.RecordCount > 0 Then
                      TA.Autorizacion = AdoAbono.Recordset.fields("Autorizacion")
                      TA.CodigoC = AdoAbono.Recordset.fields("CodigoC")
                   End If
                   If TA.CodigoC <> Ninguno Then
                       TA.Serie_R = Replace(MidStrg(.fields("SERIE_COMPROBANTE"), 1, 7), "-", "")
                       TA.Secuencial_R = Val(MidStrg(.fields("SERIE_COMPROBANTE"), 9, 9))
                       TA.AutorizacionR = .fields("NUMERO_AUTORIZACION")
                       TA.Cheque = TA.Serie_R & "-" & MidStrg(.fields("SERIE_COMPROBANTE"), 9, 9)
                       
                       sSQL = "DELETE " _
                            & "FROM Trans_Abonos " _
                            & "WHERE Item = '" & NumEmpresa & "' " _
                            & "AND Periodo = '" & Periodo_Contable & "' " _
                            & "AND TP = '" & TA.TP & "' " _
                            & "AND Serie = '" & TA.Serie & "' " _
                            & "AND Factura = " & TA.Factura & " " _
                            & "AND Serie_R = '" & TA.Serie_R & "' " _
                            & "AND Autorizacion_R = '" & TA.AutorizacionR & "' " _
                            & "AND Secuencial_R = " & TA.Secuencial_R & " "
                       Ejecutar_SQL_SP sSQL
                       
                      'Abono de Factura Retencion fuente
                       If .fields("Ret_Fuente") >= 0 And Len(.fields("Cod_Ret")) > 1 Then
                           TA.Abono = .fields("Ret_Fuente")
                           TA.Banco = "RETENCION FUENTE - " & .fields("Cod_Ret")
                           TA.Cta = TA.Cta_CxP
                           TA.Porcentaje = .fields("Porc_Ret")
                           Grabar_Abonos TA
                       End If
                    
                      'Abono de Factura Retencion IVA Bientes
                       If .fields("Ret_IVA_B") > 0 Then
                           TA.Abono = .fields("Ret_IVA_B")
                           TA.Banco = "RETENCION IVA BIENES"
                           TA.Cta = .fields("Cta_Ret_IVA_B")
                           TA.Porcentaje = .fields("Porc_Ret_IVA_B")
                           Grabar_Abonos TA
                       End If
                    
                      'Abono de Factura Retencion IVA Bientes
                       If .fields("Ret_IVA_S") > 0 Then
                           TA.Abono = .fields("Ret_IVA_S")
                           TA.Banco = "RETENCION IVA SERVICIO"
                           TA.Cta = .fields("Cta_Ret_IVA_S")
                           TA.Porcentaje = .fields("Porc_Ret_IVA_S")
                           Grabar_Abonos TA
                       End If
                   End If
              Case "Factura", "Factura Al Gasto"
                  'Generamos el Asiento
                   Trans_No = 79
                   FechaComp = Co.Fecha
                   
                   Co.Fecha = .fields("Fecha_Emision")
                   Co.CodigoB = AdoAbono.Recordset.fields("Codigo_B")
                   NombreCliente = AdoAbono.Recordset.fields("Razon_Social_Emisor")

                  'Insertamos las transacciones
                   BorrarAsientos True
                    
                   sSQL = "SELECT " & Full_Fields("Asiento") & " " _
                        & "FROM Asiento " _
                        & "WHERE Item = '" & NumEmpresa & "' " _
                        & "AND CodigoU = '" & CodigoUsuario & "' " _
                        & "AND T_No = " & Trans_No & " "
                   Select_Adodc AdoAsiento, sSQL
                   InsertarAsientos AdoAsiento, .fields("Cta_Debito"), 0, .fields("Subtotal"), 0
                   If .fields("Total_IVA") > 0 Then
                       If Len(.fields("Cta_IVA_Gasto")) > 1 Then
                          InsertarAsientos AdoAsiento, .fields("Cta_IVA_Gasto"), 0, .fields("Total_IVA"), 0
                       Else
                          InsertarAsientos AdoAsiento, Cta_IVA_Inventario, 0, .fields("Total_IVA"), 0
                       End If
                   End If
                   InsertarAsientos AdoAsiento, .fields("Cta_Credito"), 0, 0, .fields("Total")
                    
                  'Insertamos el submodulo
                   SetAdoAddNew "Asiento_SC"
                   SetAdoFields "FECHA_V", Co.Fecha
                   SetAdoFields "Codigo", Co.CodigoB
                   SetAdoFields "TC", "P"
                   SetAdoFields "Cta", .fields("Cta_Credito")
                   SetAdoFields "Beneficiario", NombreCliente
                   SetAdoFields "TM", "1"
                   SetAdoFields "DH", "2"
                   SetAdoFields "Valor", .fields("Total")
                   SetAdoFields "Serie", .fields("Serie")
                   SetAdoFields "Factura", .fields("Comprobante")
                   SetAdoFields "Detalle_SubCta", "Aut. No. " & .fields("Autorizacion")
                   SetAdoFields "T_No", Trans_No
                   SetAdoFields "SC_No", 1
                   SetAdoFields "Item", NumEmpresa
                   SetAdoFields "CodigoU", CodigoUsuario
                   SetAdoUpdate
                   
                  'Grabo en el Asiento_Compras e implicito Asiento_Air
                   Total = 0
                   If .fields("Documento") = "Factura" And Len(.fields("Cod_Ret")) > 1 Then
                      'If ChRetB = 1 Then SetAdoFields "Cta_Bienes", SinEspaciosIzq(DCRetIBienes)
                      'If ChRetS = 1 Then SetAdoFields "Cta_Servicio", SinEspaciosIzq(DCRetISer)
                      SetAdoAddNew "Asiento_Compras"
                      SetAdoFields "IdProv", Co.CodigoB
                      SetAdoFields "DevIva", "N"
                      SetAdoFields "CodSustento", .fields("CodSustento")
                      SetAdoFields "TipoComprobante", 1
                      SetAdoFields "Establecimiento", MidStrg(.fields("Serie"), 1, 3)
                      SetAdoFields "PuntoEmision", MidStrg(.fields("Serie"), 4, 3)
                      SetAdoFields "Secuencial", .fields("Comprobante")
                      SetAdoFields "Autorizacion", .fields("Autorizacion")
                      SetAdoFields "FechaEmision", .fields("Fecha_Emision")
                      SetAdoFields "FechaRegistro", .fields("Fecha_Emision")
                      SetAdoFields "FechaCaducidad", .fields("Fecha_Emision")
                      SetAdoFields "BaseNoObjIVA", "0"
                      SetAdoFields "MontoIva", .fields("Total_IVA")
                      'Subtotal, Total_IVA, Total
                      If .fields("Total_IVA") = 0 Then
                          SetAdoFields "BaseImponible", .fields("Subtotal")
                      Else
                          SetAdoFields "BaseImpGrav", .fields("Subtotal")
                      End If
                      SetAdoFields "PorcentajeIva", AXML.CodPorIva
                      If .fields("Total_IVA") > 0 Then
                          SetAdoFields "Porc_Bienes", .fields("Porc_Ret_IVA_B")
                          SetAdoFields "MontoIvaBienes", .fields("Total_IVA")
                          SetAdoFields "PorRetBienes", CodRetBien
                          SetAdoFields "ValorRetBienes", .fields("Ret_IVA_B")
                          SetAdoFields "Porc_Servicios", .fields("Porc_Ret_IVA_S")
                          SetAdoFields "MontoIvaServicios", .fields("Total_IVA")
                          SetAdoFields "PorRetServicios", CodRetServ
                          SetAdoFields "ValorRetServicios", .fields("Ret_IVA_S")
                      End If
                      SetAdoFields "PagoLocExt", "01"
                      SetAdoFields "PaisEfecPago", "NA"
                      SetAdoFields "AplicConvDobTrib", "NA"
                      SetAdoFields "PagExtSujRetNorLeg", "NA"
                      SetAdoFields "BaseImpIce", 0
                      SetAdoFields "PorcentajeIce", 0
                      SetAdoFields "MontoIce", 0
                      SetAdoFields "DocModificado", "0"
                      SetAdoFields "FechaEmiModificado", FechaSistema
                      SetAdoFields "EstabModificado", "000"
                      SetAdoFields "PtoEmiModificado", "000"
                      SetAdoFields "SecModificado", "0000000"
                      SetAdoFields "AutModificado", "0000000000"
                      SetAdoFields "ContratoPartidoPolitico", "0000000000"
                      SetAdoFields "MontoTituloOneroso", 0
                      SetAdoFields "MontoTituloGratuito", 0
                     'Verifico si activaron los checks de retenciones: Forma de Pago
                      SetAdoFields "FormaPago", "20"
                      SetAdoFields "A_No", 1
                      SetAdoFields "T_No", Trans_No
                      SetAdoFields "CodigoU", CodigoUsuario
                      SetAdoUpdate

                      If Len(.fields("Cod_Ret")) > 1 Then
                         RatonReloj
'''                         Espizq = SinEspaciosIzq(DCConceptoRet)
'''                         Espder = TrimStrg(MidStrg(DCConceptoRet, Len(Espizq) + 3, Len(DCConceptoRet)))
                         SetAdoAddNew "Asiento_Air"
                         SetAdoFields "CodRet", .fields("Cod_Ret")
                         SetAdoFields "Detalle", "Retencion Fuente"
                         SetAdoFields "BaseImp", .fields("Subtotal")
                         SetAdoFields "Porcentaje", .fields("Porc_Ret") / 100
                         SetAdoFields "ValRet", .fields("Ret_Fuente")
                         SetAdoFields "EstabRetencion", MidStrg(DCSerieRetencion, 1, 3)
                         SetAdoFields "PtoEmiRetencion", MidStrg(DCSerieRetencion, 4, 3)
                         SetAdoFields "SecRetencion", SecRetencion
                         SetAdoFields "AutRetencion", RUC
                         SetAdoFields "FechaEmiRet", .fields("Fecha_Emision")
                         SetAdoFields "EstabFactura", "001"
                         SetAdoFields "PuntoEmiFactura", "001"
                         SetAdoFields "Factura_No", .fields("Comprobante")
                         SetAdoFields "Cta_Retencion", .fields("Cta_Ret_Fuente")
                         SetAdoFields "IdProv", Co.CodigoB
                         SetAdoFields "A_No", 1
                         SetAdoFields "T_No", Trans_No
                         SetAdoFields "Tipo_Trans", "C"
                         SetAdoUpdate
                      End If

                      OpcDH = 2
                      sSQL = "SELECT " & Full_Fields("Asiento_Compras") & " " _
                           & "FROM Asiento_Compras " _
                           & "WHERE Item = '" & NumEmpresa & "' " _
                           & "AND CodigoU = '" & CodigoUsuario & "' " _
                           & "AND T_No = " & Trans_No & " "
                      Select_Adodc AdoAux, sSQL
                      With AdoAux.Recordset
                       If .RecordCount > 0 Then
                        'Porcentaje por Servicio: 0,30,100
                         Cta = .fields("Cta_Servicio")
                         DetalleComp = "Retencion del " & .fields("Porc_Servicios") & "%, Factura No. " & .fields("Secuencial") & ", de " & NombreCliente
                         Codigo = Leer_Cta_Catalogo(Cta)
                         ValorDH = .fields("ValorRetServicios")
                         Total_RetIVA = Total_RetIVA + .fields("ValorRetServicios")
                '''         If ValorDH > 0 Then InsertarAsiento AdoAsientos
                        'Porcentaje por Bienes: 0,70,100
                         Cta = .fields("Cta_Bienes")
                         DetalleComp = "Retencion del " & .fields("Porc_Bienes") & "%, Factura No. " & .fields("Secuencial") & ", de " & NombreCliente
                         Codigo = Leer_Cta_Catalogo(Cta)
                         ValorDH = .fields("ValorRetBienes")
                         Total_RetIVA = Total_RetIVA + .fields("ValorRetBienes")
                '''         If ValorDH > 0 Then InsertarAsiento AdoAsientos
                       End If
                      End With
                     'Grabamos el Asiento de las Retenciones
                      sSQL = "SELECT " & Full_Fields("Asiento_Air") & " " _
                           & "FROM Asiento_Air " _
                           & "WHERE Item = '" & NumEmpresa & "' " _
                           & "AND CodigoU = '" & CodigoUsuario & "' " _
                           & "AND T_No = " & Trans_No & " " _
                           & "AND Tipo_Trans = 'C' " _
                           & "ORDER BY Cta_Retencion,A_No,ValRet "
                      Select_Adodc AdoAux, sSQL
                      With AdoAux.Recordset
                       If .RecordCount > 0 Then
                           Do While Not .EOF
                              Cta = .fields("Cta_Retencion")
                              DetalleComp = "Retencion (" & .fields("CodRet") & ") No. " & .fields("SecRetencion") & " del " & (.fields("Porcentaje") * 100) & "%, de " & NombreCliente
                              Codigo = Leer_Cta_Catalogo(Cta)
                              ValorDH = .fields("ValRet")
                              Total_Ret = Total_Ret + .fields("ValRet")
                '''            If ValorDH > 0 Then InsertarAsiento AdoAsientos
                             .MoveNext
                           Loop
                       End If
                      End With
                   End If
                  'Procedemos a Grabar el Comprobante
                   NumComp = ReadSetDataNum("Diario", True, True)
                    
                   DiarioCaja = NumComp
                  'Grabacion del Comprobante
                   Co.Concepto = "Doc. No. " & .fields("Serie") & "-" & Format(.fields("Comprobante"), "000000000") & ", Aut. " & .fields("Autorizacion") _
                               & "; R.U.C. " & .fields("RUC_Emisor") & ", " & NombreCliente
                   If .fields("Documento") = "Factura" Then Co.Concepto = "Compra, " & Co.Concepto Else Co.Concepto = "Gastos Personales, " & Co.Concepto
                   Co.T = Normal
                   Co.TP = CompDiario
                   Co.Numero = NumComp
                   Co.CodigoB = CodigoCli
                   Co.Efectivo = Total
                   Co.Monto_Total = Total
                   Co.T_No = Trans_No
                   Co.Usuario = CodigoUsuario
                   Co.Item = NumEmpresa
                   
                   Co.RetNueva = True
                   Co.RetSecuencial = True
                   Co.Serie_R = DCSerieRetencion.Text
                   GrabarComprobante Co
                   Control_Procesos Normal, Co.Concepto
            End Select
           .MoveNext
         Loop
         Progreso_Final
         MsgBox "Proceso Terminado, proceda a revisar la informacion subida"
         Unload Me
     End If
    End With
End Sub

Public Sub Leer_Archivo_XML(RutaArchivoXML As String)
Dim doc As New MSXML2.DOMDocument
Dim nodeList As MSXML2.IXMLDOMNodeList
Dim nodeList1 As MSXML2.IXMLDOMNodeList
Dim node As MSXML2.IXMLDOMNode
Dim node1 As MSXML2.IXMLDOMNode
Dim success As Boolean
Dim IdVer As Long
Dim IdXML As Long
Dim IdXML1 As Long
Dim nodeName As String
Dim ExisteInformacion As Boolean
Dim TipoDoc As String
Dim TipoCodRet As String
Dim VersionXML As String
Dim VerXMLTemp As String
   
   RatonReloj
   VerXMLTemp = Leer_Archivo_Texto(RutaArchivoXML)
   IdVer = InStr(VerXMLTemp, "id=""comprobante"" version=""2.0.0""")
   VerXMLTemp = MidStrg(VerXMLTemp, IdVer, 50)
   IdVer = InStr(VerXMLTemp, "version=")
   VersionXML = MidStrg(VerXMLTemp, IdVer + 9, 5)
   'MsgBox VersionXML
   success = doc.Load(RutaArchivoXML)
   If success = False Then
      TextoImprimio = TextoImprimio & doc.parseError.reason & vbCrLf
      Insertar_Texto_Temporal_SP doc.parseError.reason
      'MsgBox doc.parseError.reason
   Else
      RatonReloj
      TipoDoc = ""
      ExisteInformacion = False
      Set nodeList = doc.selectNodes("/factura/infoTributaria")
      If Not nodeList Is Nothing Then
         For Each node In nodeList
             TipoDoc = node.selectSingleNode("codDoc").Text
         Next node
      End If
      Set nodeList = doc.selectNodes("/comprobanteRetencion/infoTributaria")
      If Not nodeList Is Nothing Then
         For Each node In nodeList
             TipoDoc = node.selectSingleNode("codDoc").Text
         Next node
      End If
      Select Case TipoDoc
        Case "01"
             Set nodeList = doc.selectNodes("/factura/infoTributaria")
             AXML.Documento = "Factura"
        Case "07"
             Set nodeList = doc.selectNodes("/comprobanteRetencion/infoTributaria")
             AXML.Documento = "Retencion"
      End Select
     'MsgBox "ARCHIVO: (" & tipoDoc & ")" & AXML.Documento & ": " & vbCrLf & RutaArchivoXML
      If Not nodeList Is Nothing Then
         For Each node In nodeList
             AXML.Ambiente = Val(node.selectSingleNode("ambiente").Text)
             AXML.Razon_Social_Emisor = UCaseStrg(node.selectSingleNode("razonSocial").Text)
             AXML.RUC_Emisor = node.selectSingleNode("ruc").Text
             AXML.Direccion_Emisor = UCaseStrg(node.selectSingleNode("dirMatriz").Text)
             AXML.Serie = node.selectSingleNode("estab").Text & node.selectSingleNode("ptoEmi").Text
             AXML.Comprobante = node.selectSingleNode("secuencial").Text
             AXML.Autorizacion = node.selectSingleNode("claveAcceso").Text
             TipoDoc = node.selectSingleNode("codDoc").Text
         Next node
      End If
      RatonReloj
      Select Case TipoDoc
        Case "01" 'Facturas
             Set nodeList = doc.selectNodes("/factura/infoFactura")
             If Not nodeList Is Nothing Then
                For Each node In nodeList
                    AXML.RUC_Receptor = node.selectSingleNode("identificacionComprador").Text
                    AXML.Fecha_Emision = node.selectSingleNode("fechaEmision").Text
                    AXML.SubTotal = Val(node.selectSingleNode("totalSinImpuestos").Text)
                    AXML.Total = Val(node.selectSingleNode("importeTotal").Text)
                    AXML.Total_IVA = 0
                    Set nodeList1 = doc.selectNodes("/factura/infoFactura/totalConImpuestos/totalImpuesto")
                    If Not nodeList1 Is Nothing Then
                       For Each node1 In nodeList1
                           If Val(node1.selectSingleNode("valor").Text) > 0 And Val(node1.selectSingleNode("codigoPorcentaje").Text) > 0 Then
                              AXML.Total_IVA = AXML.Total_IVA + Val(node1.selectSingleNode("valor").Text)
                              AXML.CodPorIva = Val(node1.selectSingleNode("codigoPorcentaje").Text)
                           End If
                       Next node1
                    End If
                    Set nodeList1 = doc.selectNodes("/factura/infoFactura/pagos/pago")
                    If Not nodeList1 Is Nothing Then
                       For Each node1 In nodeList1
                           AXML.FormaPago = node1.selectSingleNode("formaPago").Text
                       Next node1
                    End If
                Next node
             End If
        Case "07" 'Retenciones
             AXML.SubTotal = 0
             AXML.Total_IVA = 0
             Set nodeList = doc.selectNodes("/comprobanteRetencion/infoCompRetencion")
             If Not nodeList Is Nothing Then
                For Each node In nodeList
                    AXML.Fecha_Emision = node.selectSingleNode("fechaEmision").Text
                    AXML.RUC_Receptor = node.selectSingleNode("identificacionSujetoRetenido").Text
                Next node
             End If
             If VersionXML = "2.0.0" Then
                Set nodeList = doc.selectNodes("/comprobanteRetencion/docsSustento/docSustento")
             Else
               'Version: 1.0.0
                Set nodeList = doc.selectNodes("/comprobanteRetencion/impuestos/impuesto")
             End If
             If Not nodeList Is Nothing Then
                For Each node In nodeList
                    AXML.Cod_Sustento = node.selectSingleNode("codDocSustento").Text
                    AXML.Serie_Receptor = MidStrg(node.selectSingleNode("numDocSustento").Text, 1, 6)
                    AXML.Comprobante = CLng(MidStrg(node.selectSingleNode("numDocSustento").Text, 7, 9))
                Next node
             End If
             If VersionXML = "2.0.0" Then
                Set nodeList = doc.selectNodes("/comprobanteRetencion/docsSustento/docSustento/retenciones/retencion")
             Else
               'Version: 1.0.0
                Set nodeList = doc.selectNodes("/comprobanteRetencion/impuestos/impuesto")
             End If
             If Not nodeList Is Nothing Then
                For Each node In nodeList
                    If node.selectSingleNode("codigo").Text = "1" Then
                       AXML.Ret_Fuente = Val(node.selectSingleNode("valorRetenido").Text)
                       AXML.Cod_Ret = node.selectSingleNode("codigoRetencion").Text
                       AXML.SubTotal = AXML.SubTotal + Val(node.selectSingleNode("baseImponible").Text)
                       AXML.Porc_Ret = Val(node.selectSingleNode("porcentajeRetener").Text)
                       AXML.Ret_Fuente = Val(node.selectSingleNode("valorRetenido").Text)
                    End If
                Next node
             End If
             If VersionXML = "2.0.0" Then
                Set nodeList = doc.selectNodes("/comprobanteRetencion/docsSustento/docSustento/retenciones/retencion")
             Else
               'Version: 1.0.0
                Set nodeList = doc.selectNodes("/comprobanteRetencion/impuestos/impuesto")
             End If
             If Not nodeList Is Nothing Then
                For Each node In nodeList
                    If node.selectSingleNode("codigo").Text = "2" Then
                       AXML.Total_IVA = AXML.Total_IVA + Val(node.selectSingleNode("baseImponible").Text)
                       AXML.CodPorIva = node.selectSingleNode("codigo").Text
                       Select Case Val(node.selectSingleNode("porcentajeRetener").Text)
                         Case 10, 30: AXML.Ret_IVA_B = Val(node.selectSingleNode("valorRetenido").Text)
                                      AXML.Porc_Ret_IVA_B = Val(node.selectSingleNode("porcentajeRetener").Text)
                                      AXML.Cod_Ret_IVA_B = Val(node.selectSingleNode("codigoRetencion").Text)
                         Case 20, 70: AXML.Ret_IVA_S = Val(node.selectSingleNode("valorRetenido").Text)
                                      AXML.Porc_Ret_IVA_S = Val(node.selectSingleNode("porcentajeRetener").Text)
                                      AXML.Cod_Ret_IVA_S = Val(node.selectSingleNode("codigoRetencion").Text)
                         Case 50:     AXML.Ret_IVA_S = Val(node.selectSingleNode("valorRetenido").Text)
                                      AXML.Porc_Ret_IVA_S = Val(node.selectSingleNode("porcentajeRetener").Text)
                                      AXML.Cod_Ret_IVA_S = Val(node.selectSingleNode("codigoRetencion").Text)
                         Case 100:    AXML.Ret_IVA_S = Val(node.selectSingleNode("valorRetenido").Text)
                                      AXML.Porc_Ret_IVA_S = Val(node.selectSingleNode("porcentajeRetener").Text)
                                      AXML.Cod_Ret_IVA_S = Val(node.selectSingleNode("codigoRetencion").Text)
                       End Select
                    End If
                Next node
             End If
             AXML.Total = AXML.SubTotal + AXML.Total_IVA
             'MsgBox "...."
        Case Else
             TextoImprimio = TextoImprimio & RutaArchivoXML & vbCrLf
             Insertar_Texto_Temporal_SP RutaArchivoXML
             'MsgBox "No exite este documentos en la base"
      End Select
      RatonReloj
      DigVerif = Digito_Verificador(AXML.RUC_Emisor)
      AXML.Codigo_B = Tipo_RUC_CI.Codigo_RUC_CI
      sSQL = "SELECT Codigo, Cliente, TD, CI_RUC " _
           & "FROM Clientes " _
           & "WHERE Codigo = '" & AXML.Codigo_B & "' "
      Select_Adodc AdoAux, sSQL
      If AdoAux.Recordset.RecordCount <= 0 Then
         SetAdoAddNew "Clientes"
         SetAdoFields "T", Normal
         SetAdoFields "Codigo", AXML.Codigo_B
         SetAdoFields "TD", Tipo_RUC_CI.Tipo_Beneficiario
         SetAdoFields "CI_RUC", AXML.RUC_Emisor
         SetAdoFields "Cliente", AXML.Razon_Social_Emisor
         SetAdoFields "Direccion", AXML.Direccion_Emisor
         SetAdoFields "Fecha", FechaSistema
         SetAdoFields "DirNumero", "SN"
         SetAdoFields "Ciudad", NombreCiudad
         SetAdoFields "Prov", "17"
         SetAdoFields "Pais", "593"
         SetAdoFields "CodigoU", CodigoUsuario
         SetAdoUpdate
      End If
      'MsgBox Cadena
   End If
   RatonNormal
End Sub

Private Sub Command2_Click()
  Unload FXMLRecibidosSRI
End Sub

Private Sub DCSerieRetencion_Change()
  If AdoSerieRetencion.Recordset.RecordCount > 0 Then
     AdoSerieRetencion.Recordset.MoveFirst
     AdoSerieRetencion.Recordset.Find ("Concepto = '" & DCSerieRetencion & "' ")
     If Not AdoSerieRetencion.Recordset.EOF Then Label3.Caption = Format$(AdoSerieRetencion.Recordset.fields("Numero"), "000000000")
  End If
End Sub

Private Sub DGDocSRI_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
  Select Case DGDocSRI.Columns(ColIndex).Caption
    Case "Procesar"
         If DGDocSRI.Columns("Procesar").Value = 0 Or DGDocSRI.Columns("Procesar").Value = 1 Then Cancel = False Else Cancel = True
    Case "Cod_Ret"
         Mifecha = BuscarFecha(DGDocSRI.Columns("Fecha_Emision").Value)
         CodRet = DGDocSRI.Columns("Cod_Ret").Value
         If Len(CodRet) > 1 Then
            AXML.Cta_Ret_Fuente = Ninguno
            sSQL = "SELECT TOP 1 Codigo, Porcentaje " _
                 & "FROM Tipo_Concepto_Retencion " _
                 & "WHERE Codigo = '" & CodRet & "' " _
                 & "AND Fecha_Inicio <= #" & Mifecha & "# " _
                 & "AND Fecha_Final >= #" & Mifecha & "# "
            Select_Adodc AdoAux, sSQL
            If AdoAux.Recordset.RecordCount > 0 Then
               AXML.Porc_Ret = AdoAux.Recordset.fields("Porcentaje")
               Porcentaje = AXML.Porc_Ret / 100
               DGDocSRI.Columns("Porc_Ret").Value = AdoAux.Recordset.fields("Porcentaje")
               DGDocSRI.Columns("Ret_Fuente").Value = Redondear(DGDocSRI.Columns("Subtotal").Value * Porcentaje, 2)
               If AXML.Porc_Ret > 0 Then
                  sSQL = "SELECT Codigo " _
                       & "FROM Catalogo_Cuentas " _
                       & "WHERE Item = '" & NumEmpresa & "' " _
                       & "AND Periodo = '" & Periodo_Contable & "' " _
                       & "AND TC = 'RF' " _
                       & "AND DG = 'D' " _
                       & "AND Cuenta LIKE '%" & CStr(AXML.Porc_Ret) & "%%' "
                  Select_Adodc AdoAux, sSQL
                  If AdoAux.Recordset.RecordCount > 0 Then DGDocSRI.Columns("Cta_Ret_Fuente").Value = AdoAux.Recordset.fields("Codigo")
               End If
               Cancel = False
            Else
               MsgBox "Este Codigo no esta permitido en esta fecha"
               Cancel = True
            End If
         Else
            If CodRet = Ninguno Then
               DGDocSRI.Columns("Cod_Ret").Value = CodRet
               DGDocSRI.Columns("Porc_Ret").Value = 0
               DGDocSRI.Columns("Ret_Fuente").Value = 0
               Cancel = False
            Else
               Cancel = True
            End If
         End If
         'Cancel = False
    Case "Porc_Ret_IVA_B"
         Porcentaje = DGDocSRI.Columns("Porc_Ret_IVA_B").Value
         If 0 <= Porcentaje And Porcentaje <= 100 Then
            sSQL = "SELECT Codigo " _
                 & "FROM Tabla_Por_IVA " _
                 & "WHERE Bienes <> " & Val(adFalse) & " " _
                 & "AND Porc = '" & CStr(Porcentaje) & "' "
            Select_Adodc AdoAux, sSQL
            If AdoAux.Recordset.RecordCount > 0 Then
               Porcentaje = Porcentaje / 100
               DGDocSRI.Columns("Ret_IVA_B").Value = Redondear(DGDocSRI.Columns("Total_IVA").Value * Porcentaje, 2)
               Cancel = False
            Else
               Cancel = True
            End If
         Else
            Cancel = True
         End If
    Case "Porc_Ret_IVA_S"
         Porcentaje = DGDocSRI.Columns("Porc_Ret_IVA_S").Value
         If 0 <= Porcentaje And Porcentaje <= 100 Then
            sSQL = "SELECT Codigo " _
                 & "FROM Tabla_Por_IVA " _
                 & "WHERE Servicios <> " & Val(adFalse) & " " _
                 & "AND Porc = '" & CStr(Porcentaje) & "' "
            Select_Adodc AdoAux, sSQL
            If AdoAux.Recordset.RecordCount > 0 Then
               Porcentaje = Porcentaje / 100
               DGDocSRI.Columns("Ret_IVA_S").Value = Redondear(DGDocSRI.Columns("Total_IVA").Value * Porcentaje, 2)
               Cancel = False
            Else
               Cancel = True
            End If
         Else
            Cancel = True
         End If
    Case Else
         Cancel = True
  End Select
  If DGDocSRI.Columns("Documento").Value = "Retencion" Then Cancel = True
End Sub

Private Sub DGDocSRI_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  If KeyCode = vbKeyEscape Then MsgBox "Proceda a grabar"
  If KeyCode = vbKeyReturn Then
     If AdoDocSRI.Recordset.RecordCount > 0 Then
        AdoDocSRI.Recordset.MoveNext
        If AdoDocSRI.Recordset.EOF Then AdoDocSRI.Recordset.MoveFirst
     End If
  End If
  If CtrlDown And KeyCode = vbKeyInsert Then
     Comp_No = DGDocSRI.Columns(0).Text
     Mensajes = "Asignar CxP las cuentas de procesos a:" & vbCrLf & DGDocSRI.Columns(3).Text
     Titulo = "Pregunta de CxP"
     If BoxMensaje = vbYes Then
        CodigoCliente = DGDocSRI.Columns(31).Text
        NombreCliente = DGDocSRI.Columns(3).Text
        SubCta = "P"
        AXML.Codigo_B = CodigoCliente
        FCxCxP.Show 1
        
       'Viene del formulario anteror: AXML.Cta_Credito
        sSQL = "SELECT TOP 1 Cta, Cta_Gasto, SubModulo, Cod_Ret, Porc_IVAB, Porc_IVAS " _
             & "FROM Catalogo_CxCxP " _
             & "WHERE Item = '" & NumEmpresa & "' " _
             & "AND Periodo = '" & Periodo_Contable & "' " _
             & "AND Codigo = '" & AXML.Codigo_B & "' " _
             & "AND Cta = '" & AXML.Cta_Credito & "' " _
             & "AND TC = 'P' " _
             & "ORDER BY Cta "
        Select_Adodc AdoAux, sSQL
        If AdoAux.Recordset.RecordCount > 0 Then
           AXML.Cta_Debito = AdoAux.Recordset.fields("Cta_Gasto")
           AXML.SubModulo = AdoAux.Recordset.fields("SubModulo")
           AXML.Cod_Ret = AdoAux.Recordset.fields("Cod_Ret")
           AXML.Porc_Ret_IVA_B = AdoAux.Recordset.fields("Porc_IVAB")
           AXML.Porc_Ret_IVA_S = AdoAux.Recordset.fields("Porc_IVAS")
           
           Leer_Porc_Retenciones "2"
           
           'If AXML.CodPorIva = "" Then AXML.CodPorIva = Ninguno

           sSQL = "UPDATE Asiento_SRI " _
                & "SET Cod_Ret = '" & AXML.Cod_Ret & "', " _
                & "Cta_Debito = '" & AXML.Cta_Debito & "', " _
                & "Cta_Credito = '" & AXML.Cta_Credito & "', " _
                & "Cta_IVA_Gasto = '" & AXML.Cta_IVA_Gasto & "', " _
                & "Cta_Ret_Fuente = '" & AXML.Cta_Ret_Fuente & "', " _
                & "Cta_Ret_IVA_B = '" & AXML.Cta_Ret_IVA_B & "', " _
                & "Cta_Ret_IVA_S = '" & AXML.Cta_Ret_IVA_S & "', " _
                & "CodPorIva = '" & AXML.CodPorIva & "', " _
                & "Porc_Ret = " & AXML.Porc_Ret & ", " _
                & "Porc_Ret_IVA_B = " & AXML.Porc_Ret_IVA_B & ", " _
                & "Porc_Ret_IVA_S = " & AXML.Porc_Ret_IVA_S & " " _
                & "WHERE Item = '" & NumEmpresa & "' " _
                & "AND CodigoU = '" & CodigoUsuario & "' " _
                & "AND Codigo_B = '" & CodigoCliente & "' "
           Ejecutar_SQL_SP sSQL
        End If
        
        sSQL = "SELECT " & Full_Fields("Asiento_SRI") & " " _
             & "FROM Asiento_SRI " _
             & "WHERE Item = '" & NumEmpresa & "' " _
             & "AND CodigoU = '" & CodigoUsuario & "' " _
             & "ORDER BY No "
        Select_Adodc_Grid DGDocSRI, AdoDocSRI, sSQL
     End If
  End If
End Sub

Private Sub Form_Activate()
    
    Set ftp = New cFTP

    Toolbar1.buttons("Excel").Enabled = False
    Toolbar1.buttons("Grabar").Enabled = True
        
    FXMLRecibidosSRI.Caption = "LECTURA DE COMPROBANTES ELECTRONICOS DEL SRI"
    PctSRI.Picture = LoadPicture(RutaSistema & "\LOGOS\srilinea.jpg")
    
    Codigo = Ninguno
    NuevoComp = True
    ModificarComp = False
    CopiarComp = False
    Co.CodigoB = ""
    Co.Numero = 0
  
    Cta_Gastos = Leer_Seteos_Ctas("Cta_Gastos")
    Cta_Gastos_Personales = Leer_Seteos_Ctas("Cta_Gastos_Personales")
    
    sSQL = "SELECT " & Full_Fields("Asiento_SRI") & " " _
         & "FROM Asiento_SRI " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND CodigoU = '" & CodigoUsuario & "' " _
         & "ORDER BY No "
    Select_Adodc_Grid DGDocSRI, AdoDocSRI, sSQL
    If AdoDocSRI.Recordset.RecordCount > 0 Then
       Titulo = "PREGUNTA DE ELIMINACION"
       Mensajes = "Existen procesos pendientes, Desea Eliminar las Transacciones?"
       If BoxMensaje = vbYes Then
          sSQL = "DELETE * " _
               & "FROM Asiento_SRI " _
               & "WHERE Item = '" & NumEmpresa & "' " _
               & "AND CodigoU = '" & CodigoUsuario & "' "
          Ejecutar_SQL_SP sSQL
          
          sSQL = "SELECT " & Full_Fields("Asiento_SRI") & " " _
               & "FROM Asiento_SRI " _
               & "WHERE Item = '" & NumEmpresa & "' " _
               & "AND CodigoU = '" & CodigoUsuario & "' " _
               & "ORDER BY No "
          Select_Adodc_Grid DGDocSRI, AdoDocSRI, sSQL
       Else
          DGDocSRI.Enabled = True
       End If
    End If
    
    sSQL = "SELECT Codigo & ' - ' & Cuenta As Nombre_Cta " _
         & "FROM Catalogo_Cuentas " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND DG = 'D' " _
         & "AND TC = 'P' " _
         & "ORDER BY Cuenta "
    SelectDB_Combo DCCxP, AdoCxP, sSQL, "Nombre_Cta", False
    
    sSQL = "SELECT Concepto, Numero " _
         & "FROM Codigos " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND Concepto LIKE 'RE_SERIE%' " _
         & "AND LEN(Concepto) = 15 " _
         & "ORDER BY Concepto "
    SelectDB_Combo DCSerieRetencion, AdoSerieRetencion, sSQL, "Concepto"
    
    sSQL = "SELECT (Credito_Tributario & ' - ' & Descripcion) As Sustento,Codigo_Tipo_Comprobante " _
         & "FROM Tipo_Tributario " _
         & "WHERE Credito_Tributario <> '.' " _
         & "AND Fecha_Inicio <= #" & BuscarFecha(FechaSistema) & "# " _
         & "AND Fecha_Final >= #" & BuscarFecha(FechaSistema) & "# " _
         & "ORDER BY Credito_Tributario "
    SelectDB_Combo DCSustento, AdoSustento, sSQL, "Sustento"
   
    sSQL = "SELECT (Codigo & ' - ' & Cuenta) As CtaGasto " _
         & "FROM Catalogo_Cuentas " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND TC IN ('G','CC') " _
         & "AND DG = 'D' " _
         & "ORDER BY Codigo "
    SelectDB_Combo DCCtaGasto, AdoCtaGasto, sSQL, "CtaGasto"
    
    sSQL = "SELECT (Codigo & ' ' & Descripcion) As CTipoPago " _
         & "FROM Tabla_Referenciales_SRI " _
         & "WHERE Tipo_Referencia = 'FORMA DE PAGO' " _
         & "ORDER BY Codigo "
    SelectDB_Combo DCTipoPago, AdoTipoPago, sSQL, "CTipoPago"
   
   'MsgBox DCCxP & vbCrLf & DCSerieRetencion & vbCrLf & DCSustento
    RatonNormal
    DCCxP.SetFocus
End Sub

Private Sub Form_Deactivate()
  FXMLRecibidosSRI.WindowState = vbMaximized
End Sub

Private Sub Form_Load()
  ConectarAdodc AdoAux
  ConectarAdodc AdoTxt
  ConectarAdodc AdoCxP
  ConectarAdodc AdoAbono
  ConectarAdodc AdoCtaGasto
  ConectarAdodc AdoTipoPago
  ConectarAdodc AdoDocSRI
  ConectarAdodc AdoAsiento
  ConectarAdodc AdoClientes
  ConectarAdodc AdoSustento
  ConectarAdodc AdoSerieRetencion
  
  DGDocSRI.Height = MDI_Y_Max - 2150
  DGDocSRI.width = MDI_X_Max - DGDocSRI.Left - 50
  AdoDocSRI.width = MDI_X_Max - AdoDocSRI.Left - 50
  'DCSustento.width = MDI_X_Max - DCSustento.Left - 50
  Label4.width = MDI_X_Max - Label4.Left - 50
  AdoDocSRI.Top = DGDocSRI.Top + DGDocSRI.Height + 10
  PctSRI.width = MDI_X_Max - PctSRI - 50
  
 'Verificamos y creamos carpetas de firma electronica
  RutaDocumentos = RutaSysBases & "\SRI"
  If Not Existe_Carpeta(RutaDocumentos) Then MkDir RutaDocumentos

  RutaDocumentos = RutaSysBases & "\SRI\Comprobantes no Autorizados"
  If Not Existe_Carpeta(RutaDocumentos) Then MkDir RutaDocumentos
  
  RutaDocumentos = RutaSysBases & "\SRI\Comprobantes Recibidos"
  If Not Existe_Carpeta(RutaDocumentos) Then MkDir RutaDocumentos
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
    DGDocSRI.Enabled = True
   'MsgBox Button.key
    Select Case Button.key
      Case "Salir"
           Unload Me
      Case "Leer_XML"
           Leer_XML
      Case "Excel"
           DGDocSRI.Visible = False
           GenerarDataTexto FXMLRecibidosSRI, AdoDocSRI
           DGDocSRI.Visible = True
      Case "Grabar"
           Grabar_Comprobantes_XML
    End Select
End Sub
