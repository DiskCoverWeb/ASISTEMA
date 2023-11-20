VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.4#0"; "comctl32.Ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Begin VB.Form IngProdInv 
   Caption         =   "Ingreso/Modificacion de Productos de Inventario"
   ClientHeight    =   7605
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11385
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7605
   ScaleWidth      =   11385
   WindowState     =   2  'Maximized
   Begin ComctlLib.TreeView TVCatalogo 
      Height          =   3270
      Left            =   105
      TabIndex        =   62
      Top             =   105
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   5768
      _Version        =   327682
      Style           =   7
      ImageList       =   "ImageList1"
      Appearance      =   1
   End
   Begin VB.Frame FrmInv1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "LISTA DE PRODUCTOS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3360
      Left            =   3885
      TabIndex        =   53
      Top             =   3780
      Visible         =   0   'False
      Width           =   7335
      Begin VB.CommandButton Command8 
         Caption         =   "Cancelar"
         Height          =   750
         Left            =   6090
         Picture         =   "IngInven.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   61
         Top             =   2520
         Width           =   1065
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Insertar"
         Height          =   750
         Left            =   4935
         Picture         =   "IngInven.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   60
         Top             =   2520
         Width           =   1065
      End
      Begin VB.TextBox TxtCantReceta 
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
         Left            =   105
         MaxLength       =   12
         MultiLine       =   -1  'True
         TabIndex        =   56
         Top             =   2835
         Width           =   1170
      End
      Begin MSDataListLib.DataList DLInv1 
         Bindings        =   "IngInven.frx":1194
         DataSource      =   "AdoInv1"
         Height          =   2205
         Left            =   120
         TabIndex        =   54
         Top             =   240
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   3889
         _Version        =   393216
         BackColor       =   16777088
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label20 
         BackColor       =   &H0080FF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Cantidad"
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
         TabIndex        =   55
         Top             =   2520
         Width           =   1170
      End
   End
   Begin VB.CommandButton Command6 
      Caption         =   "&Imprimir"
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
      Left            =   10290
      Picture         =   "IngInven.frx":11AA
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   945
      Width           =   960
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4005
      Left            =   120
      TabIndex        =   0
      Top             =   3480
      Width           =   11145
      _ExtentX        =   19659
      _ExtentY        =   7064
      _Version        =   393216
      TabHeight       =   1058
      TabCaption(0)   =   "Datos Principales"
      TabPicture(0)   =   "IngInven.frx":15EC
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label7"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label10"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label12"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label11"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label13"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label9"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label4"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label17"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label14"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label15"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label8"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label19"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label16"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label18"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label6"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Option1"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "CheqInv"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Option2"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "CheqIVA"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Picture1"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "MBoxCodigo"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "TextSubCta"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "MBoxCta_Inv"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "MBoxCta1"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "TextUnidad"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "TextPVP"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "MBoxCta_Ing"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "MBoxCta_Ing0"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "TextMinimo"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "TextMaximo"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "TxtItem_Banco"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "TxtPX"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "TxtPY"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "TxtGramaje"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "CheqAgrupacion"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "MBoxCta_IngAnt"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "TxtDesc_Item"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "TxtCorte"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "TxtAyuda"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "TextBarra"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "Frame1"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).ControlCount=   44
      TabCaption(1)   =   "Validación de Texto"
      TabPicture(1)   =   "IngInven.frx":1906
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label5"
      Tab(1).Control(1)=   "TxtDetalle"
      Tab(1).Control(2)=   "Command5"
      Tab(1).Control(3)=   "Command4"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Salidas Automáticas"
      TabPicture(2)   =   "IngInven.frx":21E0
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "DGReceta"
      Tab(2).Control(1)=   "AdoReceta"
      Tab(2).ControlCount=   2
      Begin VB.Frame Frame1 
         Caption         =   "Calcular Divisas"
         Height          =   540
         Left            =   3885
         TabIndex        =   57
         Top             =   3255
         Width           =   3375
         Begin VB.OptionButton OpcMul 
            Caption         =   "Multiplicar"
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
            Left            =   1785
            TabIndex        =   59
            Top             =   210
            Width           =   1380
         End
         Begin VB.OptionButton OpcDiv 
            Caption         =   "Dividir"
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
            TabIndex        =   58
            Top             =   210
            Width           =   1170
         End
      End
      Begin MSAdodcLib.Adodc AdoReceta 
         Height          =   330
         Left            =   -74895
         Top             =   3360
         Width           =   10935
         _ExtentX        =   19288
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
         Caption         =   "Receta"
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
      Begin MSDataGridLib.DataGrid DGReceta 
         Bindings        =   "IngInven.frx":2ABA
         Height          =   2640
         Left            =   -74895
         TabIndex        =   52
         ToolTipText     =   "<Ctrl+Insert> Insertar Rubro, <Ctrl+Suprimir> Elimina Rubro"
         Top             =   735
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   4657
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
      Begin VB.TextBox TextBarra 
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
         MaxLength       =   20
         TabIndex        =   40
         ToolTipText     =   "<Alt+F2> Codigo Automático"
         Top             =   2940
         Width           =   3375
      End
      Begin VB.TextBox TxtAyuda 
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
         MaxLength       =   60
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   48
         Top             =   2940
         Width           =   3795
      End
      Begin VB.TextBox TxtCorte 
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
         Left            =   3045
         MaxLength       =   10
         MultiLine       =   -1  'True
         TabIndex        =   38
         Top             =   2625
         Width           =   855
      End
      Begin VB.TextBox TxtDesc_Item 
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
         MaxLength       =   20
         TabIndex        =   32
         Top             =   2310
         Width           =   3795
      End
      Begin MSMask.MaskEdBox MBoxCta_IngAnt 
         Height          =   330
         Left            =   5670
         TabIndex        =   18
         Top             =   1995
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         _Version        =   393216
         AllowPrompt     =   -1  'True
         AutoTab         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin VB.CheckBox CheqAgrupacion 
         Caption         =   "&Agrupación"
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
         Left            =   9660
         TabIndex        =   51
         Top             =   1050
         Width           =   1380
      End
      Begin VB.TextBox TxtGramaje 
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
         Left            =   8190
         MaxLength       =   4
         TabIndex        =   36
         Text            =   "0"
         Top             =   2310
         Width           =   960
      End
      Begin VB.TextBox TxtPY 
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
         Left            =   10080
         MaxLength       =   12
         TabIndex        =   30
         Top             =   1995
         Width           =   960
      End
      Begin VB.TextBox TxtPX 
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
         Left            =   8190
         MaxLength       =   12
         TabIndex        =   28
         Top             =   1995
         Width           =   960
      End
      Begin VB.TextBox TxtItem_Banco 
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
         Left            =   6825
         MaxLength       =   3
         TabIndex        =   34
         Text            =   "000"
         Top             =   2310
         Width           =   435
      End
      Begin VB.TextBox TextMaximo 
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
         Left            =   10080
         MaxLength       =   10
         MultiLine       =   -1  'True
         TabIndex        =   26
         Top             =   1680
         Width           =   960
      End
      Begin VB.TextBox TextMinimo 
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
         Left            =   8190
         MaxLength       =   10
         MultiLine       =   -1  'True
         TabIndex        =   24
         Top             =   1680
         Width           =   960
      End
      Begin MSMask.MaskEdBox MBoxCta_Ing0 
         Height          =   330
         Left            =   5670
         TabIndex        =   16
         Top             =   1680
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         _Version        =   393216
         AllowPrompt     =   -1  'True
         AutoTab         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MBoxCta_Ing 
         Height          =   330
         Left            =   1890
         TabIndex        =   14
         Top             =   1680
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         _Version        =   393216
         AllowPrompt     =   -1  'True
         AutoTab         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin VB.TextBox TextPVP 
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
         Left            =   9870
         MaxLength       =   12
         MultiLine       =   -1  'True
         TabIndex        =   22
         Top             =   1365
         Width           =   1170
      End
      Begin VB.TextBox TextUnidad 
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
         Left            =   8190
         MaxLength       =   6
         TabIndex        =   20
         Top             =   1365
         Width           =   960
      End
      Begin MSMask.MaskEdBox MBoxCta1 
         Height          =   330
         Left            =   5670
         TabIndex        =   12
         Top             =   1365
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         _Version        =   393216
         AllowPrompt     =   -1  'True
         AutoTab         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MBoxCta_Inv 
         Height          =   330
         Left            =   1890
         TabIndex        =   10
         Top             =   1365
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         _Version        =   393216
         AllowPrompt     =   -1  'True
         AutoTab         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin VB.TextBox TextSubCta 
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
         MaxLength       =   80
         TabIndex        =   4
         Top             =   735
         Width           =   5370
      End
      Begin MSMask.MaskEdBox MBoxCodigo 
         Height          =   330
         Left            =   1155
         TabIndex        =   2
         Top             =   735
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   582
         _Version        =   393216
         MaxLength       =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "CC.CC.CCC.CCCCC"
         Mask            =   "CC.CC.CCC.CCCCC"
         PromptChar      =   " "
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Cambios"
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
         Left            =   -64815
         TabIndex        =   50
         Top             =   3150
         Width           =   855
      End
      Begin VB.CommandButton Command5 
         Caption         =   "&Generar P.V.P."
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
         Height          =   435
         Left            =   -66495
         Picture         =   "IngInven.frx":2AD2
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   3150
         Width           =   1590
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   7350
         ScaleHeight     =   855
         ScaleWidth      =   3690
         TabIndex        =   41
         Top             =   2730
         Width           =   3690
      End
      Begin VB.CheckBox CheqIVA 
         Caption         =   "&Facturar con IVA"
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
         TabIndex        =   7
         Top             =   1050
         Width           =   1800
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Producto final"
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
         Left            =   2940
         TabIndex        =   6
         Top             =   1050
         Width           =   2745
      End
      Begin VB.CheckBox CheqInv 
         Caption         =   "Producto para Facturar"
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
         Left            =   7245
         TabIndex        =   8
         Top             =   1050
         Width           =   2325
      End
      Begin VB.TextBox TxtDetalle 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1905
         Left            =   -74895
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   46
         Text            =   "IngInven.frx":2F14
         Top             =   1155
         Width           =   10935
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Tipo d&e Inventario"
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
         TabIndex        =   5
         Top             =   1050
         Value           =   -1  'True
         Width           =   2850
      End
      Begin VB.Label Label6 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " CODIGO DE B&ARRA"
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
         TabIndex        =   39
         Top             =   2625
         Width           =   3375
      End
      Begin VB.Label Label18 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " UTILIDAD DE VENTA PROD. %"
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
         TabIndex        =   37
         Top             =   2625
         Width           =   2955
      End
      Begin VB.Label Label16 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Descripción Item"
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
         TabIndex        =   31
         Top             =   2310
         Width           =   1800
      End
      Begin VB.Label Label19 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " CTA. DE &VENTA PERIODO ANTERIOR"
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
         Top             =   1995
         Width           =   5580
      End
      Begin VB.Label Label8 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Gramaje"
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
         Left            =   7245
         TabIndex        =   35
         Top             =   2310
         Width           =   960
      End
      Begin VB.Label Label15 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " POS. Y"
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
         Left            =   9135
         TabIndex        =   29
         Top             =   1995
         Width           =   960
      End
      Begin VB.Label Label14 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " POS. X"
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
         Left            =   7245
         TabIndex        =   27
         Top             =   1995
         Width           =   960
      End
      Begin VB.Label Label17 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Item Banco"
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
         TabIndex        =   33
         Top             =   2310
         Width           =   1170
      End
      Begin VB.Label Label4 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " MA&XIMO:"
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
         Left            =   9135
         TabIndex        =   25
         Top             =   1680
         Width           =   960
      End
      Begin VB.Label Label9 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " &MINIMO:"
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
         Left            =   7245
         TabIndex        =   23
         Top             =   1680
         Width           =   960
      End
      Begin VB.Label Label13 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " CTA. VENTA TARIF. 0"
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
         Left            =   3465
         TabIndex        =   15
         Top             =   1680
         Width           =   2220
      End
      Begin VB.Label Label11 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " CTA. DE &VENTA"
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
         TabIndex        =   13
         Top             =   1680
         Width           =   1800
      End
      Begin VB.Label Label12 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " &P.V.P."
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
         Left            =   9135
         TabIndex        =   21
         Top             =   1365
         Width           =   750
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " &UNIDAD:"
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
         Left            =   7245
         TabIndex        =   19
         Top             =   1365
         Width           =   960
      End
      Begin VB.Label Label10 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " CTA. COS&TO DE VENT."
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
         Left            =   3465
         TabIndex        =   11
         Top             =   1365
         Width           =   2220
      End
      Begin VB.Label Label7 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " CTA. &INVENTARIO"
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
         Top             =   1365
         Width           =   1800
      End
      Begin VB.Label Label3 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " C&oncepto"
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
         Left            =   2940
         TabIndex        =   3
         Top             =   735
         Width           =   2745
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " &Codigo:"
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
         Top             =   735
         Width           =   1065
      End
      Begin VB.Label Label5 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " &DETALLE COMPLETO DEL PRODUCTO"
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
         Left            =   -74895
         TabIndex        =   47
         Top             =   735
         Width           =   10935
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Imprimir"
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
      Left            =   10290
      Picture         =   "IngInven.frx":2F18
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   1785
      Width           =   960
   End
   Begin VB.PictureBox Code39Clt1 
      Height          =   480
      Left            =   9030
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   63
      Top             =   210
      Width           =   1200
   End
   Begin MSAdodcLib.Adodc AdoInv 
      Height          =   330
      Left            =   420
      Top             =   1260
      Visible         =   0   'False
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   2
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
      Left            =   10290
      Picture         =   "IngInven.frx":335A
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   2625
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
      Left            =   10290
      Picture         =   "IngInven.frx":3C24
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   105
      Width           =   960
   End
   Begin MSAdodcLib.Adodc AdoAux 
      Height          =   330
      Left            =   420
      Top             =   945
      Visible         =   0   'False
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
   Begin MSAdodcLib.Adodc AdoTInv 
      Height          =   330
      Left            =   420
      Top             =   1575
      Visible         =   0   'False
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   2
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
   Begin MSAdodcLib.Adodc AdoCodInv 
      Height          =   330
      Left            =   420
      Top             =   1890
      Visible         =   0   'False
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
      Caption         =   "CodInv"
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
   Begin MSAdodcLib.Adodc AdoInv1 
      Height          =   330
      Left            =   420
      Top             =   2205
      Visible         =   0   'False
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   2
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
      Caption         =   "Inv1"
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
   Begin ComctlLib.ImageList ImageList1 
      Left            =   5250
      Top             =   1260
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   3
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "IngInven.frx":4066
            Key             =   "UNO"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "IngInven.frx":4380
            Key             =   "DOS"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "IngInven.frx":469A
            Key             =   "TRES"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "IngProdInv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Cta_Ini As String
Dim Cta_Fin As String
Dim nodX As Node

Public Sub AddNewCtaInv(TipoTC As String)
Dim ICod As Byte
  ICod = 1
  Do While Mid$(MascaraCodigoK, ICod, 1) <> "."
     ICod = ICod + 1
  Loop
  If Len(Codigo) = ICod Then
     Set nodX = TVCatalogo.Nodes.Add(, , Codigo, Cuenta)
     nodX.Image = ImageList1.ListImages(1).key
     nodX.SelectedImage = ImageList1.ListImages(1).key
  Else
     Set nodX = TVCatalogo.Nodes.Add(Cta_Sup, tvwChild, Codigo, Cuenta)
     Select Case TipoTC
       Case "I": IE = 1
       Case "P": IE = 2
       Case Else: IE = 3
     End Select
     nodX.Image = ImageList1.ListImages(IE).key
     nodX.SelectedImage = ImageList1.ListImages(IE).key
     TVCatalogo.Tag = Codigo
  End If
End Sub

Public Sub ColocarCodigoBarra(CodigoDeBarra As String, PVP As Currency)
''Dim PosSup, PosIzq As Single
''  Code39Clt1.AlturaBarra = 45
''  Code39Clt1.TamBarra = 1
''  Code39Clt1.ColorCodigo = "N"
''  Code39Clt1.ValorCodigo = CodigoDeBarra
''  Code39Clt1.RealizarCodigo
''
''  Picture1 = Clipboard.GetData
''  Picture1.FontBold = True
''  PosIzq = ((Picture1.width - Picture1.TextWidth(Code39Clt1.ValorCodigo)) / 2)
''  PosSup = Picture1.Height - 360
''  If PosIzq < 0 Then PosIzq = 0.1
''  If PosSup < 0 Then PosSup = 0.1
''  Picture1.Line (150, PosSup)-(Picture1.width - 150, Picture1.Height), QBColor(Blanco_Brillante), BF
''  Picture1.CurrentX = PosIzq
''  Picture1.CurrentY = PosSup
''  Picture1.Print Code39Clt1.ValorCodigo
''  PosSup = Picture1.Height - 160
''  Cadena = "USD$ " & Format(PVP, "#,##0.0000")
''  PosIzq = ((Picture1.width - Picture1.TextWidth(Cadena)) / 2)
''  Picture1.CurrentX = PosIzq
''  Picture1.CurrentY = PosSup
''  Picture1.Print Cadena
''  Picture1.FontBold = False
''  If TextBarra.Text = Ninguno Then TextBarra.Text = Code39Clt1.ValorCodigo
End Sub

Private Sub Command1_Click()
  GrabarInv
End Sub

Private Sub Command2_Click()
  sSQL = "UPDATE Catalogo_Productos " _
       & "SET Cta_Inventario = '0' " _
       & "WHERE Cta_Inventario = '.' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' "
  ConectarAdoExecute sSQL
  sSQL = "UPDATE Catalogo_Productos " _
       & "SET Cta_Costo_Venta = '0' " _
       & "WHERE Cta_Costo_Venta = '.' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' "
  ConectarAdoExecute sSQL
  sSQL = "UPDATE Catalogo_Productos " _
       & "SET Cta_Ventas = '0' " _
       & "WHERE Cta_Ventas = '.' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' "
  ConectarAdoExecute sSQL
  sSQL = "UPDATE Catalogo_Productos " _
       & "SET Cta_Ventas_0 = '0' " _
       & "WHERE Cta_Ventas_0 = '.' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' "
  ConectarAdoExecute sSQL
  sSQL = "UPDATE Catalogo_Productos " _
       & "SET Cta_Venta_Anticipada = '0' " _
       & "WHERE Cta_Venta_Anticipada = '.' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' "
  ConectarAdoExecute sSQL
  Unload IngProdInv
End Sub

Private Sub Command3_Click()
  IE = Val(InputBox("Cantidad de Etiquetas", "IMPRESION CODIGO DE BARRAS", 0))
  If IE > 0 Then Imprimir_Codigos_De_Barras IE, TextBarra
End Sub

Private Sub Command4_Click()
  If ClaveSupervisor Then
     RatonReloj
     If Option2.value Then
        Codigo1 = CambioCodigoCta(MBoxCodigo.Text)
        FChangeCtaInv.Show 1
     Else
        RatonNormal
        MsgBox "Solo puede cambiar el Codigo de Inventario"
     End If
  End If
End Sub

Private Sub Command5_Click()
 RatonReloj
 sSQL = "UPDATE Catalogo_Productos "
 Mensajes = "Calcular Con" & vbCrLf _
          & "SI = 2 Decimales" & vbCrLf _
          & "NO = 4 Decimales"
 Titulo = "Formulario de Grabacion"
 If BoxMensaje = vbYes Then
    sSQL = sSQL & "SET PVP = ROUND((Promedio * (1 + Utilidad)),2,0) "
 Else
    sSQL = sSQL & "SET PVP = ROUND((Promedio * (1 + Utilidad)),4,0) "
 End If
 sSQL = sSQL & "WHERE Item = '" & NumEmpresa & "' " _
      & "AND Periodo = '" & Periodo_Contable & "' " _
      & "AND LEN(Cta_Inventario) > 1 " _
      & "AND Utilidad > 0 " _
      & "AND TC = 'P' "
 ConectarAdoExecute sSQL
 RatonNormal
 MsgBox "Precios Calculados con exito"
End Sub

Private Sub Command6_Click()
  Imprimir_Codigos_Estanteria SinEspaciosIzq(TVCatalogo.SelectedItem)
End Sub

Private Sub Command7_Click()
     Codigo = UCase$(CambioCodigoCta(MBoxCodigo))
     With AdoInv1.Recordset
      If .RecordCount > 0 Then
         .MoveFirst
         .Find ("Producto = '" & DLInv1.Text & "' ")
          If Not .EOF Then
             sSQL = "SELECT * " _
                  & "FROM Catalogo_Recetas " _
                  & "WHERE Codigo_PP = '" & Codigo & "' " _
                  & "AND Codigo_Receta = '" & .Fields("Codigo_Inv") & "' " _
                  & "AND Item = '" & NumEmpresa & "' " _
                  & "AND Periodo = '" & Periodo_Contable & "' "
             SelectAdodc AdoAux, sSQL
             If AdoAux.Recordset.RecordCount <= 0 Then
                SetAdoAddNew "Catalogo_Recetas"
                SetAdoFields "TC", "P"
                SetAdoFields "Codigo_PP", Codigo
                SetAdoFields "Producto", .Fields("Producto")
                SetAdoFields "Codigo_Receta", .Fields("Codigo_Inv")
                SetAdoFields "Cantidad", CCur(TxtCantReceta)
                SetAdoUpdate
             End If
          End If
      End If
     End With
     sSQL = "SELECT * " _
          & "FROM Catalogo_Recetas " _
          & "WHERE Codigo_PP = '" & Codigo & "' " _
          & "AND Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "ORDER BY Codigo_Receta "
     SelectDataGrid DGReceta, AdoReceta, sSQL
     FrmInv1.Visible = False
     DGReceta.SetFocus
End Sub

Private Sub Command8_Click()
    FrmInv1.Visible = False
End Sub

Private Sub DGReceta_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  If CtrlDown And KeyCode = vbKeyInsert Then
     sSQL = "SELECT * " _
          & "FROM Catalogo_Productos " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "AND LEN(Cta_Inventario ) > 1 " _
          & "AND LEN(Cta_Costo_Venta) > 1 " _
          & "AND TC = 'P' " _
          & "AND Codigo_Inv <> '" & Codigo & "' " _
          & "ORDER BY Codigo_Inv "
     SelectDBList DLInv1, AdoInv1, sSQL, "Producto"
     FrmInv1.Visible = True
     DLInv1.SetFocus
  End If
  If CtrlDown And KeyCode = vbKeyDelete Then
     Codigo = UCase$(CambioCodigoCta(MBoxCodigo))
     With AdoReceta.Recordset
      If .RecordCount >= 0 Then
          sSQL = "DELETE * " _
               & "FROM Catalogo_Recetas " _
               & "WHERE Codigo_PP = '" & Codigo & "' " _
               & "AND Codigo_Receta = '" & .Fields("Codigo_Receta") & "' " _
               & "AND Item = '" & NumEmpresa & "' " _
               & "AND Periodo = '" & Periodo_Contable & "' "
          ConectarAdoExecute sSQL
      End If
     End With
     sSQL = "SELECT * " _
          & "FROM Catalogo_Recetas " _
          & "WHERE Codigo_PP = '" & Codigo & "' " _
          & "AND Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "ORDER BY Codigo_Receta "
     SelectDataGrid DGReceta, AdoReceta, sSQL
     MsgBox "Producto de Subproceso Eliminado"
  End If
End Sub

Private Sub DLInv1_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub Form_Activate()
  FormatoMaskCta MBoxCta_Inv
 'FormatoMaskCta MBoxCta
  FormatoMaskCta MBoxCta1
  FormatoMaskCta MBoxCta_Ing
  FormatoMaskCta MBoxCta_Ing0
  FormatoMaskCodK MBoxCodigo
  If Modulo = "INVENTARIO" Then Command5.Enabled = True
  Si_No = False
  
  sSQL = "SELECT * " _
       & "FROM Catalogo_Productos " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' "
  SelectAdodc AdoAux, sSQL
  If AdoAux.Recordset.RecordCount <= 0 Then
     SetAdoAddNew "Catalogo_Productos"
     SetAdoFields "Item", NumEmpresa
     SetAdoFields "Periodo", Periodo_Contable
     SetAdoFields "TC", "I"
     SetAdoFields "Codigo_Inv", "01"
     SetAdoFields "Producto", "INVENTARIO"
     SetAdoUpdate
  End If
  
  sSQL = "SELECT * " _
       & "FROM Catalogo_Marcas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND CodMar = '.' "
  SelectAdodc AdoAux, sSQL
  With AdoAux.Recordset
   If .RecordCount <= 0 Then
       SetAdoAddNew "Catalogo_Marcas"
       SetAdoFields "CodMar", Ninguno
       SetAdoFields "Marca", Ninguno
       SetAdoFields "Item", NumEmpresa
       SetAdoFields "Periodo", Periodo_Contable
       SetAdoUpdate
   End If
  End With

  sSQL = "UPDATE Catalogo_Productos " _
       & "SET Cta_Inventario = '0' " _
       & "WHERE Cta_Inventario = '.' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' "
  ConectarAdoExecute sSQL
  
  sSQL = "UPDATE Catalogo_Productos " _
       & "SET Cta_Ventas = '0' " _
       & "WHERE Cta_Ventas = '.' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' "
  ConectarAdoExecute sSQL
  
  sSQL = "UPDATE Catalogo_Productos " _
       & "SET Cta_Ventas_0 = '0' " _
       & "WHERE Cta_Ventas_0 = '.' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' "
  ConectarAdoExecute sSQL
  
  sSQL = "UPDATE Catalogo_Productos " _
       & "SET Cta_Ventas_Ant = '0' " _
       & "WHERE Cta_Ventas_Ant = '.' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' "
  ConectarAdoExecute sSQL
  
  sSQL = "UPDATE Catalogo_Productos " _
       & "SET Producto = Ucase$(Producto) " _
       & "WHERE TC = 'I' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND Producto <> Ucase$(Producto) "
  ConectarAdoExecute sSQL
  
  sSQL = "UPDATE Catalogo_Productos " _
       & "SET Producto = Ucase$(Producto) " _
       & "WHERE TC = 'I' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND Producto <> Ucase$(Producto) "
  ConectarAdoExecute sSQL
  
  sSQL = "SELECT Item,Codigo_Inv,Periodo " _
       & "FROM Catalogo_Productos " _
       & "WHERE TC = 'P' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "ORDER BY Codigo_Inv "
  SelectAdodc AdoTInv, sSQL
  If AdoTInv.Recordset.RecordCount > 0 Then
     sSQL = "SELECT * " _
          & "FROM Catalogo_Productos " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "ORDER BY Codigo_Inv "
     SelectAdodc AdoInv, sSQL
     If AdoInv.Recordset.RecordCount > 0 Then
        ReDim CodigoCtas(AdoInv.Recordset.RecordCount + 1) As String
        For I = 0 To AdoInv.Recordset.RecordCount
            CodigoCtas(I) = ""
        Next I
     End If
     Contador = 0
     Do While Not AdoTInv.Recordset.EOF
        Codigo = AdoTInv.Recordset.Fields("Codigo_Inv")
        Cta_Sup = CodigoCuentaSup(Codigo)
        With AdoInv.Recordset
         If .RecordCount > 0 Then
             Do While (Cta_Sup <> "0")
               .MoveFirst
               .Find ("Codigo_Inv Like '" & Cta_Sup & "' ")
                If Not .EOF And Cta_Sup <> "0" Then
                   Cta_Sup = CodigoCuentaSup(Cta_Sup)
                Else
                   Si_No = True: Evaluar = True
                   For I = 0 To AdoInv.Recordset.RecordCount
                       If CodigoCtas(I) = Cta_Sup Then Evaluar = False
                   Next I
                   If Evaluar Then
                      SetAdoAddNew "Catalogo_Productos"
                      SetAdoFields "Item", NumEmpresa
                      SetAdoFields "Codigo_Inv", Cta_Sup
                      SetAdoFields "Producto", "NINGUN PRODUCTO"
                      SetAdoFields "Periodo", Periodo_Contable
                      SetAdoFields "TC", "I"
                      SetAdoUpdate
                      CodigoCtas(Contador) = Cta_Sup
                      Contador = Contador + 1
                   End If
                   Cta_Sup = CodigoCuentaSup(Cta_Sup)
                End If
             Loop
         End If
        End With
        AdoTInv.Recordset.MoveNext
     Loop
  End If
  
  RatonNormal
  If Si_No Then
     Cadena = vbCrLf
     For I = 0 To Contador
         Cadena = Cadena & CodigoCtas(I) & vbCrLf
     Next I
     MsgBox "Los siguientes Codigos no se han creado: " & vbCrLf _
            & Cadena & "ADVERTENCIA: REVIZAR."
  End If
  Label5.Caption = " &DETALLE COMPLETO DEL PRODUCTO" & vbCrLf
  For I = 1 To 9
    For J = 1 To 9
        Label5.Caption = Label5.Caption & CStr(J)
    Next J
    Label5.Caption = Label5.Caption & "^"
  Next I
    
  sSQL = "SELECT * " _
       & "FROM Catalogo_Productos " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND TC = 'P' " _
       & "ORDER BY Codigo_Inv "
  SelectDBList DLInv1, AdoInv1, sSQL, "Producto"
    
  sSQL = "SELECT * " _
       & "FROM Catalogo_Productos " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "ORDER BY Codigo_Inv "
  SelectAdodc AdoInv, sSQL
  RatonReloj
  With AdoInv.Recordset
   If .RecordCount > 0 Then
'      Codigo = "C" & .Fields("Codigo_Inv")
'      Cta_Sup = "C" & CodigoCuentaSup(.Fields("Codigo_Inv"))
       .MoveFirst
        Do While Not .EOF
           If Len(.Fields("Codigo_Inv")) = 2 Then
              Codigo = "C" & .Fields("Codigo_Inv")
              Cta_Sup = "C" & .Fields("Codigo_Inv")
              Cuenta = .Fields("Codigo_Inv") & " - " & .Fields("Producto")
              AddNewCtaInv .Fields("TC")
           Else
              Codigo = "C" & .Fields("Codigo_Inv")
              Cta_Sup = "C" & CodigoCuentaSup(.Fields("Codigo_Inv"))
              Cuenta = .Fields("Codigo_Inv") & " - " & .Fields("Producto")
              AddNewCtaInv .Fields("TC")
           End If
          .MoveNext
        Loop
    End If
   End With
  MBoxCodigo.SetFocus
  RatonNormal
End Sub

Private Sub Form_Load()
  ConectarAdodc AdoAux
  ConectarAdodc AdoInv
  ConectarAdodc AdoInv1
  ConectarAdodc AdoTInv
  ConectarAdodc AdoCodInv
  ConectarAdodc AdoReceta
End Sub

Private Sub MBoxCodigo_GotFocus()
  MarcarTexto MBoxCodigo
End Sub

Private Sub MBoxCodigo_LostFocus()
   MBoxCodigo.Text = UCase$(MBoxCodigo.Text)
End Sub

Private Sub MBoxCta_Ing_GotFocus()
  MarcarTexto MBoxCta_Ing
End Sub

Private Sub MBoxCta_Ing_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBoxCta_IngAnt_GotFocus()
  MarcarTexto MBoxCta_IngAnt
End Sub

Private Sub MBoxCta_IngAnt_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBoxCta_Ing0_GotFocus()
  MarcarTexto MBoxCta_Ing0
End Sub

Private Sub MBoxCta_Ing0_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBoxCta_Inv_GotFocus()
  MarcarTexto MBoxCta_Inv
End Sub

Private Sub MBoxCta_Inv_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBoxCta1_GotFocus()
  MarcarTexto MBoxCta1
End Sub

Private Sub MBoxCta1_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextBarra_GotFocus()
  MarcarTexto TextBarra
End Sub

Private Sub TextBarra_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  If AltDown And KeyCode = vbKeyF2 Then
     RatonReloj
     sSQL = "SELECT MAX(Codigo_Barra) As Maximo " _
          & "FROM Catalogo_Productos " _
          & "WHERE TC = 'P' " _
          & "AND Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' "
     SelectData AdoCodInv, sSQL
     With AdoCodInv.Recordset
      If .RecordCount > 0 Then
          TextBarra = Format(Val(.Fields("Maximo")) + 1, "0000")
      Else
          TextBarra = "0001"
      End If
     End With
     RatonNormal
  End If
  PresionoEnter KeyCode
End Sub

Private Sub TextBarra_LostFocus()
  TextoValido TextBarra, , True
  TextBarra = Replace(TextBarra, "Ñ", "N")
  TextBarra = Replace(TextBarra, "Á", "A")
  TextBarra = Replace(TextBarra, "É", "E")
  TextBarra = Replace(TextBarra, "Í", "I")
  TextBarra = Replace(TextBarra, "Ó", "O")
  TextBarra = Replace(TextBarra, "Ú", "U")
  TextBarra = Replace(TextBarra, "Ü", "U")
  TextBarra = Replace(TextBarra, "&", "Y")
End Sub

Private Sub TextMaximo_GotFocus()
 MarcarTexto TextMaximo
End Sub

Private Sub TextMaximo_KeyDown(KeyCode As Integer, Shift As Integer)
PresionoEnter KeyCode
End Sub

Private Sub TextMaximo_LostFocus()
  TextoValido TextMaximo, True
End Sub

Private Sub TextMinimo_GotFocus()
  MarcarTexto TextMinimo
End Sub

Private Sub TextMinimo_KeyDown(KeyCode As Integer, Shift As Integer)
PresionoEnter KeyCode
End Sub

Private Sub TextMinimo_LostFocus()
  TextoValido TextMinimo
End Sub

Private Sub TextPVP_GotFocus()
  MarcarTexto TextPVP
End Sub

Private Sub TextPVP_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextPVP_LostFocus()
  TextoValido TextPVP, True, , 4
End Sub

Private Sub TextSubCta_GotFocus()
  MarcarTexto TextSubCta
End Sub

Private Sub TextSubCta_LostFocus()
  TextoValido TextSubCta
End Sub

Private Sub TextUnidad_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextUnidad_LostFocus()
  TextoValido TextUnidad, , True
End Sub

Public Sub LlenarInv()
   FormatoMaskCta MBoxCta_Inv
   'FormatoMaskCta MBoxCta
   FormatoMaskCta MBoxCta1
   FormatoMaskCta MBoxCta_Ing
   FormatoMaskCta MBoxCta_Ing0
   FormatoMaskCta MBoxCta_IngAnt
   FormatoMaskCodK MBoxCodigo
   TextSubCta.Text = ""
   TextUnidad.Text = "0"
   TextMaximo.Text = "0"
   TextMinimo.Text = "0"
   TxtAyuda.Text = Ninguno
   With AdoInv.Recordset
    If .RecordCount > 0 Then
        Codigo = SinEspaciosIzq(TVCatalogo.SelectedItem)
        'Codigo = SinEspaciosIzq(DCInv.Text)
        'MsgBox Codigo & vbCrLf & CodigosSinPuntos(Codigo)
       .MoveFirst
        TextoBusqueda = "Codigo_Inv Like '" & Codigo & "' "
       .Find (TextoBusqueda)
        If Not .EOF Then
           TextSubCta.Text = .Fields("Producto")
           TextUnidad.Text = .Fields("Unidad")
           TextMaximo.Text = .Fields("Maximo")
           TextMinimo.Text = .Fields("Minimo")
           TxtGramaje.Text = .Fields("Gramaje")
           TextPVP.Text = .Fields("PVP")
           TxtAyuda.Text = .Fields("Ayuda")
           TextBarra.Text = .Fields("Codigo_Barra")
           If TextBarra.Text = Ninguno Then
              TextBarra.Text = CodigosSinPuntos(.Fields("Codigo_Inv"))
           End If
           TxtPX.Text = .Fields("PX")
           TxtPY.Text = .Fields("PY")
           TxtItem_Banco.Text = .Fields("Item_Banco")
           TxtDesc_Item.Text = .Fields("Desc_Item")
           TxtCorte.Text = .Fields("Utilidad") * 100
           MBoxCodigo.Text = FormatoCodigoKardex(.Fields("Codigo_Inv"))
           MBoxCta_Inv.Text = FormatoCodigoCta(.Fields("Cta_Inventario"))
           MBoxCta1.Text = FormatoCodigoCta(.Fields("Cta_Costo_Venta"))
           MBoxCta_Ing.Text = FormatoCodigoCta(.Fields("Cta_Ventas"))
           MBoxCta_Ing0.Text = FormatoCodigoCta(.Fields("Cta_Ventas_0"))
           MBoxCta_IngAnt.Text = FormatoCodigoCta(.Fields("Cta_Ventas_Ant"))
           If Len(.Fields("Cta_Inventario")) > 1 Then Label18.Caption = " UTILIDAD %" Else Label18.Caption = " COMISION %"
           If .Fields("TC") = "P" Then
               Option1.value = False
               Option2.value = True
           Else
               Option1.value = True
               Option2.value = False
               TextPVP.Text = "0"
               TextBarra.Text = "0000000000"
           End If
           ColocarCodigoBarra TextBarra.Text, .Fields("PVP")
           If .Fields("IVA") Then CheqIVA.value = 1 Else CheqIVA.value = 0
           If .Fields("INV") Then CheqInv.value = 1 Else CheqInv.value = 0
           If .Fields("Agrupacion") Then CheqAgrupacion.value = 1 Else CheqAgrupacion.value = 0
           If .Fields("Div") Then OpcDiv.value = 1 Else OpcMul.value = 1
           sSQL = "SELECT * " _
                & "FROM Catalogo_Recetas " _
                & "WHERE Codigo_PP = '" & .Fields("Codigo_Inv") & "' " _
                & "AND Item = '" & NumEmpresa & "' " _
                & "AND Periodo = '" & Periodo_Contable & "' " _
                & "ORDER BY Codigo_Receta "
           SelectDataGrid DGReceta, AdoReceta, sSQL
           TxtDetalle.Text = .Fields("Detalle")
        Else
            MsgBox "No existe"
        End If
    Else
        Nuevo = True
        TextSubCta.SetFocus
    End If
   End With
End Sub

Public Sub GrabarInv()
  RatonReloj
  Nuevo = False
  'TextoValido TextPVP, True
  TextoValido TxtPX, True
  TextoValido TxtPY, True
  TextoValido TextMaximo, True
  TextoValido TextMinimo, True
  TextoValido TextUnidad
  TextoValido TextBarra
  TextoValido TxtItem_Banco, , True
  TextoValido TxtDesc_Item, , True
  TextoValido TxtAyuda
  TextoValido TextPVP, True, , 5
  TextoValido TxtCorte, True
  If Len(TxtDetalle.Text) <= 1 Then TxtDetalle.Text = Ninguno
  'CampoBusqueda = DGBusq.Columns(DGBusq.Col).Caption
  Codigo = UCase$(CambioCodigoCta(MBoxCodigo))
  If Option1.value Then TextSubCta.Text = UCase$(TextSubCta.Text)
  Codigo1 = "C" & Codigo
  Cta_Sup = "C" & CodigoCuentaSup(Codigo)
  Cuenta = Codigo & " - " & TextSubCta.Text
  
  sSQL = "SELECT * " _
       & "FROM Catalogo_Productos " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "ORDER BY Codigo_Inv "
  SelectAdodc AdoInv, sSQL
  With AdoInv.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       TextoBusqueda = "Codigo_Inv Like '" & Codigo & "' "
      .Find (TextoBusqueda)
       If .EOF Then
           SetAddNew AdoInv
           Nuevo = True
       End If
   Else
      SetAddNew AdoInv
      Nuevo = True
   End If
   'MsgBox Nuevo & vbCrLf & Codigo
   SetFields AdoInv, "Codigo_Inv", Codigo
   SetFields AdoInv, "Producto", TextSubCta.Text
   If Option1.value Then Cadena = "I" Else Cadena = "P"
   TipoCta = Cadena
   SetFields AdoInv, "TC", Cadena
   SetFields AdoInv, "Unidad", TextUnidad.Text
   SetFields AdoInv, "Maximo", TextMaximo.Text
   SetFields AdoInv, "Minimo", TextMinimo.Text
   SetFields AdoInv, "Gramaje", TxtGramaje.Text
  'MsgBox CCur(TextPVP.Text)
   SetFields AdoInv, "PVP", Val(CCur(TextPVP.Text))
   SetFields AdoInv, "Codigo_Barra", TextBarra.Text
   SetFields AdoInv, "Item", NumEmpresa
   SetFields AdoInv, "Cta_Inventario", "0"
   SetFields AdoInv, "Cta_Costo_Venta", "0"
   SetFields AdoInv, "Cta_Ventas", "0"
   SetFields AdoInv, "Cta_Ventas_0", "0"
   SetFields AdoInv, "Cta_Ventas_Ant", "0"
   SetFields AdoInv, "Cta_Venta_Anticipada", "0"
   SetFields AdoInv, "Detalle", TxtDetalle.Text
   SetFields AdoInv, "PX", TxtPX.Text
   SetFields AdoInv, "PY", TxtPY.Text
   SetFields AdoInv, "Item_Banco", TxtItem_Banco.Text
   SetFields AdoInv, "Desc_Item", TxtDesc_Item.Text
   SetFields AdoInv, "Periodo", Periodo_Contable
   SetFields AdoInv, "Utilidad", CCur(TxtCorte.Text) / 100
   SetFields AdoInv, "Ayuda", TxtAyuda.Text
   Si_No = False
   If CheqIVA.value = 1 Then Si_No = True
   SetFields AdoInv, "IVA", Si_No
   Si_No = False
   If CheqInv.value = 1 Then Si_No = True
   If OpcDiv.value = 1 Then .Fields("Div") = True Else .Fields("Div") = False
   'If Cadena <> "I" Then
      SetFields AdoInv, "INV", Si_No
      SetFields AdoInv, "Cta_Inventario", CambioCodigoCta(MBoxCta_Inv.Text)
      SetFields AdoInv, "Cta_Costo_Venta", CambioCodigoCta(MBoxCta1.Text)
      SetFields AdoInv, "Cta_Ventas", CambioCodigoCta(MBoxCta_Ing.Text)
      SetFields AdoInv, "Cta_Ventas_0", CambioCodigoCta(MBoxCta_Ing0.Text)
      SetFields AdoInv, "Cta_Ventas_Ant", CambioCodigoCta(MBoxCta_IngAnt.Text)
   'Else
   '   SetFields AdoInv, "INV", False
   'End If
   Si_No = False
   If CheqAgrupacion.value = 1 Then Si_No = True
   SetFields AdoInv, "Agrupacion", Si_No
   
   SetUpdate AdoInv
   If Nuevo Then
      Codigo2 = Codigo
      Codigo = Codigo1
      AddNewCtaInv TipoCta
      Codigo = Codigo2
   Else
      IE = TVCatalogo.SelectedItem.Index
      TVCatalogo.Nodes(IE).Text = Codigo & " - " & TextSubCta.Text
      TVCatalogo.Refresh
   End If
  End With
  RatonNormal
End Sub

Private Sub TVCatalogo_DblClick()
  SiguienteControl
End Sub

Private Sub TVCatalogo_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  PresionoEnter KeyCode
  If KeyCode = vbKeyDelete Then
     Codigo = SinEspaciosIzq(TVCatalogo.SelectedItem)
     Cuenta = SinEspaciosDer(TVCatalogo.SelectedItem)
     sSQL = "SELECT * " _
          & "FROM Trans_Kardex " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "AND Codigo_Inv = '" & Codigo & "' "
     SelectAdodc AdoAux, sSQL
     If AdoAux.Recordset.RecordCount > 0 Then
        MsgBox "No se puede eliminar este codigo: " & Codigo & vbCrLf _
               & "Detalle: " & Cuenta & vbCrLf _
               & "existen datos procesados"
     Else
        Mensajes = "Seguro de Eliminar el Codigo:" & Codigo & vbCrLf _
                 & "?"
        Titulo = "ELIMINACION"
        If BoxMensaje = vbYes Then
           sSQL = "DELETE * " _
                & "FROM Catalogo_Productos " _
                & "WHERE Item = '" & NumEmpresa & "' " _
                & "AND Periodo = '" & Periodo_Contable & "' " _
                & "AND Codigo_Inv = '" & Codigo & "' "
           ConectarAdoExecute sSQL
           TVCatalogo.Nodes.Remove TVCatalogo.SelectedItem.Index
        End If
     End If
  End If
  If CtrlDown And KeyCode = vbKeyP Then ImprimirAdodc AdoInv, True, 2, 8
  If CtrlDown And KeyCode = vbKeyU Then Unload IngProdInv
End Sub

Private Sub TVCatalogo_LostFocus()
  LlenarInv
End Sub

Private Sub TxtAyuda_GotFocus()
  MarcarTexto TxtAyuda
End Sub

Private Sub TxtAyuda_LostFocus()
  TextoValido TxtAyuda
End Sub

Private Sub TxtCantReceta_GotFocus()
   MarcarTexto TxtCantReceta
End Sub

Private Sub TxtCantReceta_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtCantReceta_LostFocus()
  TextoValido TxtCantReceta, True, , 6
End Sub

Private Sub TxtCorte_GotFocus()
  MarcarTexto TxtCorte
End Sub

Private Sub TxtCorte_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter TxtCorte
End Sub

Private Sub TxtCorte_LostFocus()
  TextoValido TxtCorte, True, , 2
End Sub

Private Sub TxtDetalle_GotFocus()
  MarcarTexto TxtDetalle
End Sub

Private Sub TxtGramaje_GotFocus()
  MarcarTexto TxtGramaje
End Sub

Private Sub TxtGramaje_LostFocus()
  TextoValido TxtGramaje, True
End Sub

