VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.4#0"; "comctl32.Ocx"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "dblist32.Ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FGeneral 
   Caption         =   "Form1"
   ClientHeight    =   5760
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9555
   LinkTopic       =   "Form1"
   ScaleHeight     =   5760
   ScaleWidth      =   9555
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   9555
      _ExtentX        =   16854
      _ExtentY        =   741
      Appearance      =   1
      _Version        =   327682
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   855
      Left            =   120
      TabIndex        =   23
      Top             =   3885
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1508
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
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   405
      Left            =   6360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3720
      Width           =   2430
   End
   Begin MSComctlLib.ImageCombo ImageCombo1 
      Height          =   330
      Left            =   3780
      TabIndex        =   20
      Top             =   2835
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   582
      _Version        =   393216
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Text            =   "ImageCombo1"
   End
   Begin MSComctlLib.TabStrip TabStrip3 
      Height          =   750
      Left            =   6930
      TabIndex        =   21
      Top             =   420
      Width           =   2010
      _ExtentX        =   3545
      _ExtentY        =   1323
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox UpDown1 
      Height          =   750
      Left            =   5775
      ScaleHeight     =   690
      ScaleWidth      =   195
      TabIndex        =   22
      Top             =   945
      Width           =   255
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   5775
      Top             =   2625
      Width           =   1800
      _ExtentX        =   3175
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
      Caption         =   "Adodc1"
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
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   315
      Left            =   7560
      TabIndex        =   19
      Top             =   2625
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   556
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataList DataList1 
      Height          =   840
      Left            =   105
      TabIndex        =   18
      Top             =   2835
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1482
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4305
      Top             =   2205
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ComctlLib.Slider Slider1 
      Height          =   435
      Left            =   2040
      TabIndex        =   17
      Top             =   3840
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   767
      _Version        =   327682
   End
   Begin ComctlLib.TabStrip TabStrip1 
      Height          =   750
      Left            =   6930
      TabIndex        =   15
      Top             =   1260
      Width           =   2010
      _ExtentX        =   3545
      _ExtentY        =   1323
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   1
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   330
      Left            =   1890
      TabIndex        =   14
      Top             =   2205
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   582
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   1065
      Left            =   3780
      TabIndex        =   13
      Top             =   1050
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   1879
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "FGeneral.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "Tab 1"
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Tab 2"
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   315
      Left            =   1890
      TabIndex        =   12
      Top             =   2835
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   556
      _Version        =   393216
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBList DBList1 
      Height          =   450
      Left            =   1890
      TabIndex        =   11
      Top             =   3255
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   794
      _Version        =   393216
   End
   Begin VB.FileListBox File1 
      Height          =   1065
      Left            =   1890
      TabIndex        =   9
      Top             =   945
      Width           =   1065
   End
   Begin VB.DirListBox Dir1 
      Height          =   765
      Left            =   525
      TabIndex        =   8
      Top             =   1365
      Width           =   1275
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   525
      TabIndex        =   7
      Top             =   945
      Width           =   1275
   End
   Begin VB.Timer Timer1 
      Left            =   6405
      Top             =   420
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   1170
      Left            =   105
      TabIndex        =   6
      Top             =   945
      Width           =   330
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   330
      Left            =   5355
      TabIndex        =   5
      Top             =   420
      Width           =   960
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   330
      Left            =   1155
      TabIndex        =   4
      Top             =   525
      Width           =   960
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   225
      Left            =   210
      TabIndex        =   3
      Top             =   525
      Width           =   960
   End
   Begin VB.ListBox List1 
      Height          =   255
      Left            =   3255
      TabIndex        =   2
      Top             =   525
      Width           =   1065
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2205
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   525
      Width           =   1020
   End
   Begin VB.PictureBox Picture1 
      Height          =   435
      Left            =   3255
      ScaleHeight     =   375
      ScaleWidth      =   900
      TabIndex        =   0
      Top             =   3255
      Width           =   960
   End
   Begin VB.OLE OLE1 
      Height          =   330
      Left            =   525
      TabIndex        =   10
      Top             =   2205
      Width           =   1275
   End
   Begin VB.Image Image1 
      Height          =   540
      Left            =   3045
      Top             =   1680
      Width           =   645
   End
   Begin VB.Line Line1 
      X1              =   4305
      X2              =   4830
      Y1              =   525
      Y2              =   945
   End
   Begin VB.Shape Shape1 
      Height          =   645
      Left            =   3045
      Top             =   945
      Width           =   645
   End
End
Attribute VB_Name = "FGeneral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
'InStr (comienzo, Cadena donde buscar , La Cadena a buscar)
'Mid (cadena, inicio, longitud)
'LeftStrg(Cadena, Longitud)
'Rigth$(Cadena, Longitud)
End Sub
