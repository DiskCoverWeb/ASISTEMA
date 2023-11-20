VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.4#0"; "comctl32.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Begin VB.Form FAsientos 
   Caption         =   "ASIENTOS"
   ClientHeight    =   7380
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12810
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7380
   ScaleWidth      =   12810
   WindowState     =   2  'Maximized
   Begin MSDataListLib.DataCombo DCUsuario 
      Bindings        =   "FAsiento.frx":0000
      DataSource      =   "AdoUsuario"
      Height          =   315
      Left            =   1785
      TabIndex        =   55
      Top             =   840
      Width           =   4845
      _ExtentX        =   8546
      _ExtentY        =   556
      _Version        =   393216
      Text            =   "DataCombo1"
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   5580
      Left            =   105
      TabIndex        =   15
      Top             =   1260
      Width           =   12300
      _ExtentX        =   21696
      _ExtentY        =   9843
      _Version        =   393216
      Tabs            =   6
      TabsPerRow      =   6
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Página No. &1"
      TabPicture(0)   =   "FAsiento.frx":0019
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label6"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "DGAsientos4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "DGAsientos1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "TextConcepto4"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "TextConcepto"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "TextConcepto1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "DGAsientos"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "Página No. &2"
      TabPicture(1)   =   "FAsiento.frx":0035
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "TextConcepto2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "TextConcepto5"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "TextConcepto3"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "DGAsientos3"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "DGAsientos5"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "DGAsientos2"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label10"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label3"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Label5"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).ControlCount=   9
      TabCaption(2)   =   "Página No. &3"
      TabPicture(2)   =   "FAsiento.frx":0051
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "TextConcepto8"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "TextConcepto7"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "TextConcepto6"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "DGAsientos6"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "DGAsientos7"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "DGAsientos8"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Label13"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "Label12"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "Label11"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).ControlCount=   9
      TabCaption(3)   =   "Página No. &4"
      TabPicture(3)   =   "FAsiento.frx":006D
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "TextConcepto11"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "TextConcepto10"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "TextConcepto9"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "DGAsientos9"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "DGAsientos10"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "DGAsientos11"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "Label16"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).Control(7)=   "Label15"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).Control(8)=   "Label14"
      Tab(3).Control(8).Enabled=   0   'False
      Tab(3).ControlCount=   9
      TabCaption(4)   =   "Página No. &5"
      TabPicture(4)   =   "FAsiento.frx":0089
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "TextConcepto14"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "TextConcepto13"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "TextConcepto12"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "DGAsientos12"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).Control(4)=   "DGAsientos13"
      Tab(4).Control(4).Enabled=   0   'False
      Tab(4).Control(5)=   "DGAsientos14"
      Tab(4).Control(5).Enabled=   0   'False
      Tab(4).Control(6)=   "Label19"
      Tab(4).Control(6).Enabled=   0   'False
      Tab(4).Control(7)=   "Label18"
      Tab(4).Control(7).Enabled=   0   'False
      Tab(4).Control(8)=   "Label17"
      Tab(4).Control(8).Enabled=   0   'False
      Tab(4).ControlCount=   9
      TabCaption(5)   =   "Página No. &6"
      TabPicture(5)   =   "FAsiento.frx":00A5
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "TextConcepto17"
      Tab(5).Control(1)=   "TextConcepto16"
      Tab(5).Control(2)=   "TextConcepto15"
      Tab(5).Control(3)=   "DGAsientos15"
      Tab(5).Control(4)=   "DGAsientos16"
      Tab(5).Control(5)=   "DGAsientos17"
      Tab(5).Control(6)=   "Label22"
      Tab(5).Control(7)=   "Label21"
      Tab(5).Control(8)=   "Label20"
      Tab(5).ControlCount=   9
      Begin VB.TextBox TextConcepto17 
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
         Left            =   -73740
         TabIndex        =   70
         Top             =   3780
         Width           =   10935
      End
      Begin VB.TextBox TextConcepto16 
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
         Left            =   -73740
         TabIndex        =   67
         Top             =   2100
         Width           =   10935
      End
      Begin VB.TextBox TextConcepto15 
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
         Left            =   -73740
         TabIndex        =   64
         Top             =   420
         Width           =   10935
      End
      Begin VB.TextBox TextConcepto14 
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
         Left            =   -73740
         TabIndex        =   61
         Top             =   3780
         Width           =   10935
      End
      Begin VB.TextBox TextConcepto13 
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
         Left            =   -73740
         TabIndex        =   58
         Top             =   2100
         Width           =   10935
      End
      Begin MSDataGridLib.DataGrid DGAsientos 
         Bindings        =   "FAsiento.frx":00C1
         Height          =   1380
         Left            =   105
         TabIndex        =   44
         Top             =   2415
         Width           =   12090
         _ExtentX        =   21325
         _ExtentY        =   2434
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
      Begin VB.TextBox TextConcepto11 
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
         Left            =   -73740
         TabIndex        =   20
         Top             =   3780
         Width           =   10935
      End
      Begin VB.TextBox TextConcepto10 
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
         Left            =   -73740
         TabIndex        =   21
         Top             =   2100
         Width           =   10935
      End
      Begin VB.TextBox TextConcepto2 
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
         Left            =   -73740
         TabIndex        =   28
         Top             =   2085
         Width           =   10935
      End
      Begin VB.TextBox TextConcepto1 
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
         Left            =   1260
         TabIndex        =   16
         Top             =   3765
         Width           =   10935
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
         Left            =   1260
         TabIndex        =   17
         Top             =   2100
         Width           =   10935
      End
      Begin VB.TextBox TextConcepto12 
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
         Left            =   -73740
         TabIndex        =   26
         Top             =   420
         Width           =   10935
      End
      Begin VB.TextBox TextConcepto9 
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
         Left            =   -73740
         TabIndex        =   29
         Top             =   420
         Width           =   10935
      End
      Begin VB.TextBox TextConcepto8 
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
         Left            =   -73740
         TabIndex        =   41
         Top             =   3780
         Width           =   10935
      End
      Begin VB.TextBox TextConcepto7 
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
         Left            =   -73740
         TabIndex        =   39
         Top             =   2100
         Width           =   10935
      End
      Begin VB.TextBox TextConcepto6 
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
         Left            =   -73740
         TabIndex        =   35
         Top             =   420
         Width           =   10935
      End
      Begin VB.TextBox TextConcepto5 
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
         Left            =   -73740
         TabIndex        =   32
         Top             =   3765
         Width           =   10935
      End
      Begin VB.TextBox TextConcepto3 
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
         Left            =   -73740
         TabIndex        =   25
         Top             =   405
         Width           =   10935
      End
      Begin VB.TextBox TextConcepto4 
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
         Left            =   1260
         TabIndex        =   18
         Top             =   420
         Width           =   10935
      End
      Begin MSDataGridLib.DataGrid DGAsientos1 
         Bindings        =   "FAsiento.frx":00DB
         Height          =   1380
         Left            =   105
         TabIndex        =   45
         Top             =   4095
         Width           =   12090
         _ExtentX        =   21325
         _ExtentY        =   2434
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
      Begin MSDataGridLib.DataGrid DGAsientos3 
         Bindings        =   "FAsiento.frx":00F6
         Height          =   1380
         Left            =   -74895
         TabIndex        =   46
         Top             =   735
         Width           =   12090
         _ExtentX        =   21325
         _ExtentY        =   2434
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
      Begin MSDataGridLib.DataGrid DGAsientos5 
         Bindings        =   "FAsiento.frx":0111
         Height          =   1380
         Left            =   -74895
         TabIndex        =   47
         Top             =   4095
         Width           =   12090
         _ExtentX        =   21325
         _ExtentY        =   2434
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
      Begin MSDataGridLib.DataGrid DGAsientos6 
         Bindings        =   "FAsiento.frx":012C
         Height          =   1380
         Left            =   -74895
         TabIndex        =   48
         Top             =   735
         Width           =   12090
         _ExtentX        =   21325
         _ExtentY        =   2434
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
      Begin MSDataGridLib.DataGrid DGAsientos7 
         Bindings        =   "FAsiento.frx":0147
         Height          =   1380
         Left            =   -74895
         TabIndex        =   49
         Top             =   2415
         Width           =   12090
         _ExtentX        =   21325
         _ExtentY        =   2434
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
      Begin MSDataGridLib.DataGrid DGAsientos8 
         Bindings        =   "FAsiento.frx":0162
         Height          =   1380
         Left            =   -74895
         TabIndex        =   50
         Top             =   4095
         Width           =   12090
         _ExtentX        =   21325
         _ExtentY        =   2434
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
      Begin MSDataGridLib.DataGrid DGAsientos9 
         Bindings        =   "FAsiento.frx":017D
         Height          =   1380
         Left            =   -74895
         TabIndex        =   51
         Top             =   735
         Width           =   12090
         _ExtentX        =   21325
         _ExtentY        =   2434
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
      Begin MSDataGridLib.DataGrid DGAsientos10 
         Bindings        =   "FAsiento.frx":0198
         Height          =   1380
         Left            =   -74895
         TabIndex        =   52
         Top             =   2415
         Width           =   12090
         _ExtentX        =   21325
         _ExtentY        =   2434
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
      Begin MSDataGridLib.DataGrid DGAsientos11 
         Bindings        =   "FAsiento.frx":01B4
         Height          =   1380
         Left            =   -74895
         TabIndex        =   53
         Top             =   4095
         Width           =   12090
         _ExtentX        =   21325
         _ExtentY        =   2434
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
      Begin MSDataGridLib.DataGrid DGAsientos12 
         Bindings        =   "FAsiento.frx":01D0
         Height          =   1380
         Left            =   -74895
         TabIndex        =   54
         Top             =   735
         Width           =   12090
         _ExtentX        =   21325
         _ExtentY        =   2434
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
      Begin MSDataGridLib.DataGrid DGAsientos2 
         Bindings        =   "FAsiento.frx":01EC
         Height          =   1380
         Left            =   -74895
         TabIndex        =   56
         Top             =   2415
         Width           =   12090
         _ExtentX        =   21325
         _ExtentY        =   2434
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
      Begin MSDataGridLib.DataGrid DGAsientos4 
         Bindings        =   "FAsiento.frx":0207
         Height          =   1380
         Left            =   105
         TabIndex        =   57
         Top             =   735
         Width           =   12090
         _ExtentX        =   21325
         _ExtentY        =   2434
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
      Begin MSDataGridLib.DataGrid DGAsientos13 
         Bindings        =   "FAsiento.frx":0222
         Height          =   1380
         Left            =   -74895
         TabIndex        =   59
         Top             =   2415
         Width           =   12090
         _ExtentX        =   21325
         _ExtentY        =   2434
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
      Begin MSDataGridLib.DataGrid DGAsientos14 
         Bindings        =   "FAsiento.frx":023E
         Height          =   1380
         Left            =   -74895
         TabIndex        =   62
         Top             =   4095
         Width           =   12090
         _ExtentX        =   21325
         _ExtentY        =   2434
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
      Begin MSDataGridLib.DataGrid DGAsientos15 
         Bindings        =   "FAsiento.frx":025A
         Height          =   1380
         Left            =   -74895
         TabIndex        =   65
         Top             =   735
         Width           =   12090
         _ExtentX        =   21325
         _ExtentY        =   2434
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
      Begin MSDataGridLib.DataGrid DGAsientos16 
         Bindings        =   "FAsiento.frx":0276
         Height          =   1380
         Left            =   -74895
         TabIndex        =   68
         Top             =   2415
         Width           =   12090
         _ExtentX        =   21325
         _ExtentY        =   2434
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
      Begin MSDataGridLib.DataGrid DGAsientos17 
         Bindings        =   "FAsiento.frx":0292
         Height          =   1380
         Left            =   -74895
         TabIndex        =   71
         Top             =   4095
         Width           =   12090
         _ExtentX        =   21325
         _ExtentY        =   2434
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
      Begin VB.Label Label22 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " CONCEPTO"
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
         Left            =   -74895
         TabIndex        =   72
         Top             =   3780
         Width           =   1170
      End
      Begin VB.Label Label21 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " CONCEPTO"
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
         Left            =   -74895
         TabIndex        =   69
         Top             =   2100
         Width           =   1170
      End
      Begin VB.Label Label20 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " CONCEPTO"
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
         Left            =   -74895
         TabIndex        =   66
         Top             =   420
         Width           =   1170
      End
      Begin VB.Label Label19 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " CONCEPTO"
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
         Left            =   -74895
         TabIndex        =   63
         Top             =   3780
         Width           =   1170
      End
      Begin VB.Label Label18 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " CONCEPTO"
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
         Left            =   -74895
         TabIndex        =   60
         Top             =   2100
         Width           =   1170
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " CONCEPTO"
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
         TabIndex        =   22
         Top             =   3765
         Width           =   1170
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " CONCEPTO"
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
         Top             =   2085
         Width           =   1170
      End
      Begin VB.Label Label17 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " CONCEPTO"
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
         Left            =   -74895
         TabIndex        =   31
         Top             =   420
         Width           =   1170
      End
      Begin VB.Label Label16 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " CONCEPTO"
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
         Left            =   -74895
         TabIndex        =   34
         Top             =   3780
         Width           =   1170
      End
      Begin VB.Label Label15 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " CONCEPTO"
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
         Left            =   -74895
         TabIndex        =   38
         Top             =   2100
         Width           =   1170
      End
      Begin VB.Label Label14 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " CONCEPTO"
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
         Left            =   -74895
         TabIndex        =   42
         Top             =   420
         Width           =   1170
      End
      Begin VB.Label Label13 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " CONCEPTO"
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
         Left            =   -74895
         TabIndex        =   43
         Top             =   3780
         Width           =   1170
      End
      Begin VB.Label Label12 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " CONCEPTO"
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
         Left            =   -74895
         TabIndex        =   40
         Top             =   2100
         Width           =   1170
      End
      Begin VB.Label Label11 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " CONCEPTO"
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
         Left            =   -74895
         TabIndex        =   36
         Top             =   420
         Width           =   1170
      End
      Begin VB.Label Label10 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " CONCEPTO"
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
         Left            =   -74895
         TabIndex        =   33
         Top             =   3765
         Width           =   1170
      End
      Begin VB.Label Label3 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " CONCEPTO"
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
         Left            =   -74895
         TabIndex        =   30
         Top             =   2085
         Width           =   1170
      End
      Begin VB.Label Label5 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " CONCEPTO"
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
         Left            =   -74895
         TabIndex        =   27
         Top             =   405
         Width           =   1170
      End
      Begin VB.Label Label6 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " CONCEPTO"
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
         Top             =   405
         Width           =   1170
      End
   End
   Begin ComctlLib.ProgressBar ProgBar 
      Height          =   330
      Left            =   105
      TabIndex        =   19
      Top             =   6930
      Width           =   7155
      _ExtentX        =   12621
      _ExtentY        =   582
      _Version        =   327682
      Appearance      =   0
   End
   Begin MSAdodcLib.Adodc AdoAsientos 
      Height          =   330
      Left            =   210
      Top             =   1365
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
   Begin VB.CheckBox CheckUsuario 
      Caption         =   "Por Cajero(a):"
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
      Top             =   840
      Width           =   1590
   End
   Begin VB.TextBox TextValor 
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
      Left            =   2940
      MaxLength       =   14
      MultiLine       =   -1  'True
      TabIndex        =   6
      Text            =   "FAsiento.frx":02AE
      Top             =   420
      Width           =   1590
   End
   Begin VB.Frame Frame1 
      Height          =   750
      Left            =   1470
      TabIndex        =   2
      Top             =   0
      Width           =   1380
      Begin VB.OptionButton OpcF 
         Caption         =   "Faltante"
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
         TabIndex        =   4
         Top             =   420
         Width           =   1170
      End
      Begin VB.OptionButton OpcS 
         Caption         =   "Sobrante"
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
         TabIndex        =   3
         Top             =   210
         Value           =   -1  'True
         Width           =   1170
      End
   End
   Begin VB.CommandButton Command5 
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
      Height          =   960
      Left            =   9975
      Picture         =   "FAsiento.frx":02B2
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   105
      Width           =   1485
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Grabar Asientos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   8295
      Picture         =   "FAsiento.frx":05BC
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   105
      Width           =   1590
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "&Procesar Asientos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   6720
      Picture         =   "FAsiento.frx":09FE
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   105
      Width           =   1485
   End
   Begin MSMask.MaskEdBox MBoxFecha 
      Height          =   330
      Left            =   105
      TabIndex        =   1
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
   Begin MSMask.MaskEdBox MBoxCtaI 
      Height          =   330
      Left            =   4620
      TabIndex        =   8
      Top             =   420
      Width           =   1485
      _ExtentX        =   2619
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
   Begin MSAdodcLib.Adodc AdoAsientos1 
      Height          =   330
      Left            =   210
      Top             =   1680
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
      Caption         =   "Asientos1"
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
   Begin MSAdodcLib.Adodc AdoAsientos2 
      Height          =   330
      Left            =   210
      Top             =   1995
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
      Caption         =   "Asientos2"
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
   Begin MSAdodcLib.Adodc AdoAsientos3 
      Height          =   330
      Left            =   210
      Top             =   2310
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
      Caption         =   "Asientos3"
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
   Begin MSAdodcLib.Adodc AdoAsientos4 
      Height          =   330
      Left            =   210
      Top             =   2625
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
      Caption         =   "Asientos4"
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
   Begin MSAdodcLib.Adodc AdoAsientos5 
      Height          =   330
      Left            =   210
      Top             =   2940
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
      Caption         =   "Asientos5"
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
   Begin MSAdodcLib.Adodc AdoAsientos6 
      Height          =   330
      Left            =   210
      Top             =   3255
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
      Caption         =   "Asientos6"
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
   Begin MSAdodcLib.Adodc AdoAsientos7 
      Height          =   330
      Left            =   210
      Top             =   3570
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
      Caption         =   "Asientos7"
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
   Begin MSAdodcLib.Adodc AdoAsientos8 
      Height          =   330
      Left            =   210
      Top             =   3885
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
      Caption         =   "Asientos8"
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
   Begin MSAdodcLib.Adodc AdoAsientos9 
      Height          =   330
      Left            =   210
      Top             =   4200
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
      Caption         =   "Asientos9"
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
   Begin MSAdodcLib.Adodc AdoAsientos10 
      Height          =   330
      Left            =   210
      Top             =   4515
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
      Caption         =   "Asientos10"
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
   Begin MSAdodcLib.Adodc AdoAsientos11 
      Height          =   330
      Left            =   210
      Top             =   4830
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
      Caption         =   "Asientos11"
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
   Begin MSAdodcLib.Adodc AdoAsientos12 
      Height          =   330
      Left            =   210
      Top             =   5145
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
      Caption         =   "Asientos12"
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
   Begin MSAdodcLib.Adodc AdoFCajaCheq 
      Height          =   330
      Left            =   2520
      Top             =   2310
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
      Caption         =   "FCajaCheq"
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
   Begin MSAdodcLib.Adodc AdoFCajaEfec 
      Height          =   330
      Left            =   2520
      Top             =   2625
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
      Caption         =   "FCajaEfec"
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
   Begin MSAdodcLib.Adodc AdoCaja 
      Height          =   330
      Left            =   2520
      Top             =   2940
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
      Caption         =   "Caja"
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
   Begin MSAdodcLib.Adodc AdoUsuario 
      Height          =   330
      Left            =   2520
      Top             =   3255
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
      Caption         =   "Usuario"
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
   Begin MSAdodcLib.Adodc AdoAsientos13 
      Height          =   330
      Left            =   210
      Top             =   5460
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
      Caption         =   "Asientos13"
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
   Begin MSAdodcLib.Adodc AdoAsientos14 
      Height          =   330
      Left            =   210
      Top             =   5775
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
      Caption         =   "Asientos14"
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
   Begin MSAdodcLib.Adodc AdoAsientos15 
      Height          =   330
      Left            =   210
      Top             =   6090
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
      Caption         =   "Asientos15"
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
   Begin MSAdodcLib.Adodc AdoAsientos16 
      Height          =   330
      Left            =   210
      Top             =   6405
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
      Caption         =   "Asientos16"
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
   Begin MSAdodcLib.Adodc AdoAsientos17 
      Height          =   330
      Left            =   2520
      Top             =   1365
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
      Caption         =   "Asientos17"
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
      Left            =   2520
      Top             =   3570
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
   Begin VB.Label LabelEgresos 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   10500
      TabIndex        =   12
      Top             =   6930
      Width           =   1905
   End
   Begin VB.Label LabelIngresos 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   8505
      TabIndex        =   13
      Top             =   6930
      Width           =   1905
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " TOTALES"
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
      TabIndex        =   14
      Top             =   6930
      Width           =   1065
   End
   Begin VB.Label Label9 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Cuenta Inicial"
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
      TabIndex        =   7
      Top             =   105
      Width           =   1485
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "V A L O R"
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
      TabIndex        =   5
      Top             =   105
      Width           =   1590
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " FECHA:"
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
      Width           =   1275
   End
End
Attribute VB_Name = "FAsientos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub SumarIngEgrCaja(DtaCajaCheq As Adodc, DtaCajaEfec As Adodc)
  RatonReloj
  Saldo = 0: Total = 0
  Debe = 0: Haber = 0
  Debe_ME = 0: Haber_ME = 0
  With DtaCajaCheq.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       Do While Not .EOF
          If .Fields("ME") Then
              Saldo = Saldo + .Fields("Debitos")
          Else
              Total = Total + .Fields("Debitos")
          End If
         .MoveNext
       Loop
      .MoveFirst
   End If
  End With
  With DtaCajaEfec.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       Do While Not .EOF
          Select Case .Fields("TP")
            Case "APER", "BOVE", "RET", "RETS", "DEP", "DEPS"
                 If .Fields("ME") Then
                     Debe_ME = Debe_ME + .Fields("Debitos")
                     Haber_ME = Haber_ME + .Fields("Creditos")
                 Else
                     Debe = Debe + .Fields("Debitos")
                     Haber = Haber + .Fields("Creditos")
                 End If
          End Select
         .MoveNext
       Loop
      .MoveFirst
   End If
  End With
  RatonNormal
End Sub

Public Sub SumarIngEgr(DtaCajaCheq As Adodc, DtaCajaEfec As Adodc)
  RatonReloj
  Saldo = 0: Total = 0
  Debe = 0: Haber = 0
  Debe_ME = 0: Haber_ME = 0
  With DtaCajaCheq.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       Do While Not .EOF
          If .Fields("ME") Then
              Saldo = Saldo + .Fields("Debitos")
          Else
              Total = Total + .Fields("Debitos")
          End If
         .MoveNext
       Loop
      .MoveFirst
   End If
  End With
  With DtaCajaEfec.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       Do While Not .EOF
          If .Fields("ME") Then
              Debe_ME = Debe_ME + .Fields("Debitos")
              Haber_ME = Haber_ME + .Fields("Creditos")
          Else
              Debe = Debe + .Fields("Debitos")
              Haber = Haber + .Fields("Creditos")
          End If
         .MoveNext
       Loop
      .MoveFirst
   End If
  End With
  LabelIngCheqMN.Caption = Format(Total, "#,##0.00")
  LabelIngCheqME.Caption = Format(Saldo, "#,##0.00")
  LabelIngresos.Caption = Format(Debe, "#,##0.00")
  LabelEgresos.Caption = Format(Haber, "#,##0.00")
  LabelSaldo.Caption = Format(Debe - Haber, "#,##0.00")
  LabelIngresosME.Caption = Format(Debe_ME, "#,##0.00")
  LabelEgresosME.Caption = Format(Haber_ME, "#,##0.00")
  LabelSaldoME.Caption = Format(Debe_ME - Haber_ME, "#,##0.00")
  RatonNormal
End Sub

Private Sub Command1_Click()
  RatonReloj
  FechaValida MBoxFecha
  ProgBar.Min = 0: ProgBar.Max = 100
  ProgBar.value = 0
  Mifecha = BuscarFecha(MBoxFecha)
  FechaTexto = BuscarFecha(MBoxFecha)
  Fecha_Vence = MBoxFecha
  SumaDebe = 0: SumaHaber = 0
  IniciarAsientosCaja
 'Retiros Sin libreta: REA,REAC
  TextConcepto9.Text = "(" & NumEmpresa & ") Retiros de Clientes en agencias"
  ConceptoComp = TextConcepto9
  sSQL = "SELECT TP,Debitos,Creditos " _
       & "FROM Trans_Libretas " _
       & "WHERE Fecha = #" & Mifecha & "# " _
       & "AND T <> 'A' " _
       & "AND TP = 'REA' " _
       & "AND AC = " & Val(adFalse) & " "
  SelectAdodc AdoCaja, sSQL
  Debe = 0: Haber = 0
  With AdoCaja.Recordset
   If .RecordCount > 0 Then
       Codigo = MidStrg(.Fields("TP"), 4, 1)
       Do While Not .EOF
          Debe = Debe + .Fields("Debitos")
          Haber = Haber + .Fields("Debitos")
         .MoveNext
       Loop
   End If
  End With
  sSQL = "SELECT TP,Debitos,Creditos " _
       & "FROM Trans_Libretas " _
       & "WHERE Fecha = #" & Mifecha & "# " _
       & "AND T <> 'A' " _
       & "AND TP = 'REAC' " _
       & "AND AC = " & Val(adFalse) & " "
  SelectAdodc AdoCaja, sSQL
  With AdoCaja.Recordset
   If .RecordCount > 0 Then
       Codigo = MidStrg(.Fields("TP"), 4, 1)
       Do While Not .EOF
          Debe = Debe + .Fields("Debitos")
          Haber = Haber + .Fields("Debitos")
         .MoveNext
       Loop
   End If
  End With
  ProgBar.value = 1
  Debe = Round(Debe, 2): Haber = Round(Haber, 2)
 'MsgBox Codigo & Chr(13) & Haber
  Trans_No = 39
  InsertarAsientos AdoAsientos9, Cta_Libretas, 0, Debe, 0
  InsertarAsientos AdoAsientos9, Cta_Suspenso, 0, 0, Haber
  SumaDebe = SumaDebe + Debe
  SumaHaber = SumaHaber + Debe
 'Depositos Sin libreta: DEA,DEAC
  TextConcepto7.Text = "(" & NumEmpresa & ") Depositos de Clientes en agencias"
  ConceptoComp = TextConcepto7
  sSQL = "SELECT TP,Debitos,Creditos " _
       & "FROM Trans_Libretas " _
       & "WHERE Fecha = #" & Mifecha & "# " _
       & "AND T <> 'A' " _
       & "AND TP = 'DEA' " _
       & "AND AC = " & Val(adFalse) & " "
  SelectAdodc AdoCaja, sSQL
  Debe = 0: Haber = 0
  With AdoCaja.Recordset
   If .RecordCount > 0 Then
       Codigo = MidStrg(.Fields("TP"), 4, 1)
       Do While Not .EOF
          Debe = Debe + .Fields("Creditos")
          Haber = Haber + .Fields("Creditos")
         .MoveNext
       Loop
   End If
  End With
  sSQL = "SELECT TP,Debitos,Creditos " _
       & "FROM Trans_Libretas " _
       & "WHERE Fecha = #" & Mifecha & "# " _
       & "AND T <> 'A' " _
       & "AND TP = 'DEAC' " _
       & "AND AC = " & Val(adFalse) & " "
  SelectAdodc AdoCaja, sSQL
  With AdoCaja.Recordset
   If .RecordCount > 0 Then
       Codigo = MidStrg(.Fields("TP"), 4, 1)
       Do While Not .EOF
          Debe = Debe + .Fields("Creditos")
          Haber = Haber + .Fields("Creditos")
         .MoveNext
       Loop
   End If
  End With
  ProgBar.value = 2
  Debe = Round(Debe, 2): Haber = Round(Haber, 2)
  Trans_No = 37
  InsertarAsientos AdoAsientos7, Cta_Suspenso, 0, Haber, 0
  InsertarAsientos AdoAsientos7, Cta_Libretas, 0, 0, Debe
  SumaDebe = SumaDebe + Debe
  SumaHaber = SumaHaber + Debe
 'Cheques en Transito: DEPC,APEC,DDAC
  TextConcepto8.Text = "(" & NumEmpresa & ") Cheques en Tránsito en Boveda"
  ConceptoComp = TextConcepto8
  sSQL = "SELECT TP,Debitos,Creditos " _
       & "FROM Trans_Libretas " _
       & "WHERE Fecha = #" & Mifecha & "# " _
       & "AND T <> 'A' " _
       & "AND TP = 'DEPC' " _
       & "AND AC = " & Val(adFalse) & " " _
       & "AND CHT <> " & Val(adFalse) & " "
  SelectAdodc AdoCaja, sSQL
  Debe = 0: Haber = 0
  With AdoCaja.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          Debe = Debe + .Fields("Creditos")
          Haber = Haber + .Fields("Creditos")
          'MsgBox Debe
         .MoveNext
       Loop
   End If
  End With
  sSQL = "SELECT TP,Debitos,Creditos " _
       & "FROM Trans_Libretas " _
       & "WHERE Fecha = #" & Mifecha & "# " _
       & "AND T <> 'A' " _
       & "AND TP = 'APEC' " _
       & "AND AC = " & Val(adFalse) & " " _
       & "AND CHT <> " & Val(adFalse) & " "
  SelectAdodc AdoCaja, sSQL
  With AdoCaja.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          Debe = Debe + .Fields("Creditos")
          Haber = Haber + .Fields("Creditos")
         .MoveNext
       Loop
   End If
  End With
  sSQL = "SELECT TP,Debitos,Creditos " _
       & "FROM Trans_Libretas " _
       & "WHERE Fecha = #" & Mifecha & "# " _
       & "AND T <> 'A' " _
       & "AND TP = 'DDAC' " _
       & "AND AC = " & Val(adFalse) & " " _
       & "AND CHT <> " & Val(adFalse) & " "
  SelectAdodc AdoCaja, sSQL
  With AdoCaja.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          Debe = Debe + .Fields("Creditos")
          Haber = Haber + .Fields("Creditos")
         .MoveNext
       Loop
   End If
  End With
  ProgBar.value = 3
  Debe = Round(Debe, 2): Haber = Round(Haber, 2)
  Trans_No = 38
  InsertarAsientos AdoAsientos8, Cta_Cheque_Transito, 0, Debe, 0
  InsertarAsientos AdoAsientos8, Cta_CajaG, 0, 0, Haber    '
  SumaDebe = SumaDebe + Debe
  SumaHaber = SumaHaber + Debe
 'Retiros sin libreta: RDA
  TextConcepto5.Text = "(" & NumEmpresa & ") Retiros de Agencias"
  ConceptoComp = TextConcepto5
  sSQL = "SELECT TP,Debitos,Creditos " _
       & "FROM Trans_Libretas " _
       & "WHERE Fecha = #" & Mifecha & "# " _
       & "AND T <> 'A' " _
       & "AND TP = 'RDA' " _
       & "AND AC = " & Val(adFalse) & " "
  SelectAdodc AdoCaja, sSQL
  Debe = 0: Haber = 0
  With AdoCaja.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          Haber = Haber + .Fields("Debitos")
         .MoveNext
       Loop
   End If
  End With
  ProgBar.value = 4
  Debe = Round(Debe, 2): Haber = Round(Haber, 2)
  Trans_No = 35
  InsertarAsientos AdoAsientos5, Cta_Suspenso, 0, Haber, 0
  InsertarAsientos AdoAsientos5, Cta_CajaG, 0, 0, Haber
  SumaDebe = SumaDebe + Haber
  SumaHaber = SumaHaber + Haber
 'Depositos Sin libreta: DDA,DDAC
  TextConcepto6.Text = "(" & NumEmpresa & ") Depósito de Agencia"
  ConceptoComp = TextConcepto6
  sSQL = "SELECT TP,Debitos,Creditos " _
       & "FROM Trans_Libretas " _
       & "WHERE Fecha = #" & Mifecha & "# " _
       & "AND T <> 'A' " _
       & "AND TP = 'DDA' " _
       & "AND AC = " & Val(adFalse) & " "
  SelectAdodc AdoCaja, sSQL
  Debe = 0: Haber = 0
  With AdoCaja.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          Haber = Haber + .Fields("Creditos")
         .MoveNext
       Loop
   End If
  End With
  sSQL = "SELECT TP,Debitos,Creditos " _
       & "FROM Trans_Libretas " _
       & "WHERE Fecha = #" & Mifecha & "# " _
       & "AND T <> 'A' " _
       & "AND TP = 'DDAC' " _
       & "AND AC = " & Val(adFalse) & " "
  SelectAdodc AdoCaja, sSQL
  With AdoCaja.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          Haber = Haber + .Fields("Creditos")
         .MoveNext
       Loop
   End If
  End With
  ProgBar.value = 5
  Debe = Round(Debe, 2): Haber = Round(Haber, 2)
  Trans_No = 36
  InsertarAsientos AdoAsientos6, Cta_CajaG, 0, Haber, 0
  InsertarAsientos AdoAsientos6, Cta_Suspenso, 0, 0, Haber
  SumaDebe = SumaDebe + Haber
  SumaHaber = SumaHaber + Haber
 'Depositos: DEP,DEPC,APER,APEC
  TextConcepto.Text = "(" & NumEmpresa & ") Depósitos de Clientes"
  ConceptoComp = TextConcepto
  sSQL = "SELECT TP,Debitos,Creditos " _
       & "FROM Trans_Libretas " _
       & "WHERE Fecha = #" & Mifecha & "# " _
       & "AND T <> 'A' " _
       & "AND TP IN ('DEP','DEPP','N/CE') " _
       & "AND AC = " & Val(adFalse) & " "
  SelectAdodc AdoCaja, sSQL
  Debe = 0: Haber = 0
  With AdoCaja.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          Haber = Haber + .Fields("Creditos")
         .MoveNext
       Loop
   End If
  End With
  sSQL = "SELECT TP,Debitos,Creditos " _
       & "FROM Trans_Libretas " _
       & "WHERE Fecha = #" & Mifecha & "# " _
       & "AND T <> 'A' " _
       & "AND TP = 'DEPC' " _
       & "AND ACC = " & Val(adFalse) & " " _
       & "AND AC = " & Val(adFalse) & " " _
       & "AND CHT <> " & Val(adFalse) & " "
  SelectAdodc AdoCaja, sSQL
  With AdoCaja.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          Haber = Haber + .Fields("Creditos")
         .MoveNext
       Loop
   End If
  End With
  sSQL = "SELECT TP,Debitos,Creditos " _
       & "FROM Trans_Libretas " _
       & "WHERE Fecha = #" & Mifecha & "# " _
       & "AND T <> 'A' " _
       & "AND TP = 'APER' " _
       & "AND AC = " & Val(adFalse) & " "
  SelectAdodc AdoCaja, sSQL
  With AdoCaja.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          Haber = Haber + .Fields("Creditos")
         .MoveNext
       Loop
   End If
  End With
  sSQL = "SELECT TP,Debitos,Creditos " _
       & "FROM Trans_Libretas " _
       & "WHERE Fecha = #" & Mifecha & "# " _
       & "AND T <> 'A' " _
       & "AND TP = 'APEC' " _
       & "AND ACC = " & Val(adFalse) & " " _
       & "AND AC = " & Val(adFalse) & " " _
       & "AND CHT <> " & Val(adFalse) & " "
  SelectAdodc AdoCaja, sSQL
  With AdoCaja.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          Haber = Haber + .Fields("Creditos")
         .MoveNext
       Loop
   End If
  End With
  ProgBar.value = 6
  Debe = Round(Debe, 2): Haber = Round(Haber, 2)
  Trans_No = 30
  InsertarAsientos AdoAsientos, Cta_CajaG, 0, Haber, 0
  InsertarAsientos AdoAsientos, Cta_Libretas, 0, 0, Haber
  SumaDebe = SumaDebe + Haber
  SumaHaber = SumaHaber + Haber
 'Retiros: RET,CIER
  TextConcepto1.Text = "(" & NumEmpresa & ") Retiros de Clientes"
  ConceptoComp = TextConcepto1
  sSQL = "SELECT TP,Debitos,Creditos " _
       & "FROM Trans_Libretas " _
       & "WHERE Fecha = #" & Mifecha & "# " _
       & "AND T <> 'A' " _
       & "AND TP = 'RET' " _
       & "AND AC = " & Val(adFalse) & " "
  SelectAdodc AdoCaja, sSQL
  Debe = 0: Haber = 0
  With AdoCaja.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          Debe = Debe + .Fields("Debitos")
         .MoveNext
       Loop
   End If
  End With
  sSQL = "SELECT TP,Debitos,Creditos " _
       & "FROM Trans_Libretas " _
       & "WHERE Fecha = #" & Mifecha & "# " _
       & "AND T <> 'A' " _
       & "AND TP = 'CIER' " _
       & "AND AC = " & Val(adFalse) & " "
  SelectAdodc AdoCaja, sSQL
  With AdoCaja.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          Debe = Debe + .Fields("Debitos")
         .MoveNext
       Loop
   End If
  End With
  ProgBar.value = 7
  Debe = Round(Debe, 2): Haber = Round(Haber, 2)
  Trans_No = 31
  InsertarAsientos AdoAsientos1, Cta_Libretas, 0, Debe, 0
  InsertarAsientos AdoAsientos1, Cta_CajaG, 0, 0, Debe
  SumaDebe = SumaDebe + Debe
  SumaHaber = SumaHaber + Debe
 'Apertura/Certificados: N/DC,N/DG,NCCA
  TextConcepto2.Text = "(" & NumEmpresa & ") Aperturas y Certificados de Aportación"
  ConceptoComp = TextConcepto2
  sSQL = "SELECT TP,Debitos,Creditos " _
       & "FROM Trans_Libretas " _
       & "WHERE Fecha = #" & Mifecha & "# " _
       & "AND T <> 'A' " _
       & "AND TP = 'N/DC' " _
       & "AND AC = " & Val(adFalse) & " "
  SelectAdodc AdoCaja, sSQL
  Debe = 0: Haber = 0: TotalPasivo = 0
  With AdoCaja.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          Debe = Debe + .Fields("Debitos")
         .MoveNext
       Loop
   End If
  End With
  sSQL = "SELECT TP,Debitos,Creditos " _
       & "FROM Trans_Libretas " _
       & "WHERE Fecha = #" & Mifecha & "# " _
       & "AND T <> 'A' " _
       & "AND TP = 'N/DG' " _
       & "AND AC = " & Val(adFalse) & " "
  SelectAdodc AdoCaja, sSQL
  With AdoCaja.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          Haber = Haber + .Fields("Debitos")
         .MoveNext
       Loop
   End If
  End With
  sSQL = "SELECT TP,Debitos,Creditos " _
       & "FROM Trans_Libretas " _
       & "WHERE Fecha = #" & Mifecha & "# " _
       & "AND T <> 'A' " _
       & "AND TP = 'NCCA' " _
       & "AND AC = " & Val(adFalse) & " "
  SelectAdodc AdoCaja, sSQL
  With AdoCaja.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          TotalPasivo = TotalPasivo + .Fields("Debitos")
         .MoveNext
       Loop
   End If
  End With
  ProgBar.value = 8
  Debe = Round(Debe, 2): Haber = Round(Haber, 2): TotalPasivo = Round(TotalPasivo, 2)
  Trans_No = 32
  InsertarAsientos AdoAsientos2, Cta_Libretas, 0, Debe + Haber + TotalPasivo, 0
  InsertarAsientos AdoAsientos2, Cta_Certificado_Apor, 0, 0, TotalPasivo
  InsertarAsientos AdoAsientos2, Cta_Certificado, 0, 0, Debe
  InsertarAsientos AdoAsientos2, Cta_Apertura, 0, 0, Haber
  SumaDebe = SumaDebe + Debe + Haber
  SumaHaber = SumaHaber + Debe + Haber
 'Retiro de Certificados de Aportacion: RECA
  Debe = 0: Haber = 0
  sSQL = "SELECT TP,Debitos,Creditos " _
       & "FROM Trans_Libretas " _
       & "WHERE Fecha = #" & Mifecha & "# " _
       & "AND T <> 'A' " _
       & "AND TP = 'RECA' " _
       & "AND AC = " & Val(adFalse) & " "
  SelectAdodc AdoCaja, sSQL
  With AdoCaja.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          Haber = Haber + .Fields("Creditos")
         .MoveNext
       Loop
   End If
  End With
  ProgBar.value = 8
  Debe = Round(Debe, 2): Haber = Round(Haber, 2)
  Trans_No = 47
  
  InsertarAsientos AdoAsientos17, Cta_Libretas, 0, 0, Haber
  InsertarAsientos AdoAsientos17, Cta_Certificado, 0, Haber, 0
  TextConcepto17.Text = "(" & NumEmpresa & ") Transferencia de Certificados de Aportación a Depositos de Ahorro"
  SumaDebe = SumaDebe + Debe + Haber
  SumaHaber = SumaHaber + Debe + Haber
 'Depositos de Intereses Ganados: INT
  TextConcepto10.Text = "(" & NumEmpresa & ") Intereses Ganados"
  ConceptoComp = TextConcepto10
  sSQL = "SELECT TP,SUM(Creditos) As TCreditos " _
       & "FROM Trans_Libretas " _
       & "WHERE Fecha = #" & Mifecha & "# " _
       & "AND T <> 'A' " _
       & "AND TP = 'INT' " _
       & "AND AC = " & Val(adFalse) & " " _
       & "GROUP BY TP "
  SelectAdodc AdoCaja, sSQL
  Debe = 0: Haber = 0
  With AdoCaja.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          Haber = Haber + .Fields("TCreditos")
         .MoveNext
       Loop
   End If
  End With
  ProgBar.value = 9
  Debe = Round(Debe, 2): Haber = Round(Haber, 2)
  Trans_No = 40
  InsertarAsientos AdoAsientos10, Cta_Interes1, 0, Haber, 0
  InsertarAsientos AdoAsientos10, Cta_Libretas, 0, 0, Haber
  SumaDebe = SumaDebe + Haber
  SumaHaber = SumaHaber + Haber
 'Intereses Diarios
  TextConcepto3.Text = "(" & NumEmpresa & ") Provisiones de Intereses por Pagar en ahorros"
  ConceptoComp = TextConcepto3
  Debe = 0: Haber = 0
  'MsgBox Mifecha
  sSQL = "SELECT Fecha,SUM(Interes) As TInteres " _
       & "FROM Trans_Intereses " _
       & "WHERE Fecha = #" & Mifecha & "# " _
       & "AND AC = " & Val(adFalse) & " " _
       & "GROUP BY Fecha "
  'MsgBox sSQL
  SelectAdodc AdoCaja, sSQL
  Contador = 0: Saldo = 0: TotalInteres = 0
  With AdoCaja.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          Debe = Debe + .Fields("TInteres")
          Mifecha = .Fields("Fecha")
         .MoveNext
       Loop
   End If
  End With
  'MsgBox Debe
  ProgBar.value = 10
  Trans_No = 33
  'MsgBox Debe
  InsertarAsientos AdoAsientos3, Cta_Interes, 0, Debe, 0
  InsertarAsientos AdoAsientos3, Cta_Interes1, 0, 0, Debe
  SumaDebe = SumaDebe + Debe
  SumaHaber = SumaHaber + Debe
 'Faltante/Sobrante
  Debe = 0: Haber = 0
  If OpcS.value Then
     Cta_Sobrantes = CambioCodigoCta(MBoxCtaI.Text)
     Haber = Val(TextValor.Text)
     Debe = Haber
     TextConcepto4.Text = "(" & NumEmpresa & ") Sobrante de Caja"
     ConceptoComp = TextConcepto4
     Trans_No = 34
     InsertarAsientos AdoAsientos4, Cta_CajaG, 0, Debe, 0
     InsertarAsientos AdoAsientos4, Cta_Sobrantes, 0, 0, Haber
     SumaDebe = SumaDebe + Debe
     SumaHaber = SumaHaber + Debe
  Else
     Cta_Faltantes = CambioCodigoCta(MBoxCtaI.Text)
     Debe = Round(Val(TextValor.Text), 2)
     Haber = Debe
     TextConcepto4.Text = "(" & NumEmpresa & ") Faltante de Caja"
     ConceptoComp = TextConcepto4
     Trans_No = 34
     InsertarAsientos AdoAsientos4, Cta_Faltantes, 0, Debe, 0
     InsertarAsientos AdoAsientos4, Cta_CajaG, 0, 0, Haber
     SumaDebe = SumaDebe + Debe
     SumaHaber = SumaHaber + Haber
  End If
 'Mantenimiento: NDMT
  Mifecha = BuscarFecha(MBoxFecha)
  TextConcepto11.Text = "(" & NumEmpresa & ") Nota de dédito por Mantenimiento de Cuentas"
  ConceptoComp = TextConcepto11
  sSQL = "SELECT TP,Debitos,Creditos " _
       & "FROM Trans_Libretas " _
       & "WHERE Fecha = #" & Mifecha & "# " _
       & "AND T <> 'A' " _
       & "AND TP = 'NDMT' " _
       & "AND AC = " & Val(adFalse) & " "
  SelectAdodc AdoCaja, sSQL
  'MsgBox sSQL
  Debe = 0: Haber = 0
  With AdoCaja.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          Debe = Debe + .Fields("Debitos")
         .MoveNext
       Loop
   End If
  End With
  ProgBar.value = 11
  Haber = Debe
  Debe = Round(Debe, 2): Haber = Round(Haber, 2)
  Trans_No = 41
  InsertarAsientos AdoAsientos11, Cta_Libretas, 0, Debe, 0
  InsertarAsientos AdoAsientos11, Cta_Mantenimiento, 0, 0, Debe
  SumaDebe = SumaDebe + Debe
  SumaHaber = SumaHaber + Debe
  
 'Mantenimiento: NDFM
  Mifecha = BuscarFecha(MBoxFecha)
  TextConcepto13.Text = "(" & NumEmpresa & ") Nota de dédito por Fondo Mortuorio"
  ConceptoComp = TextConcepto13
  sSQL = "SELECT TP,Debitos,Creditos " _
       & "FROM Trans_Libretas " _
       & "WHERE Fecha = #" & Mifecha & "# " _
       & "AND T <> 'A' " _
       & "AND TP = 'NDFM' " _
       & "AND AC = " & Val(adFalse) & " "
  SelectAdodc AdoCaja, sSQL
  'MsgBox sSQL
  Debe = 0: Haber = 0
  With AdoCaja.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          Debe = Debe + .Fields("Debitos")
         .MoveNext
       Loop
   End If
  End With
  ProgBar.value = 11
  Haber = Debe
  Debe = Round(Debe, 2): Haber = Round(Haber, 2)
  Trans_No = 43
  InsertarAsientos AdoAsientos13, Cta_Libretas, 0, Debe, 0
  InsertarAsientos AdoAsientos13, Cta_Fondo_Mortuorio, 0, 0, Debe
  SumaDebe = SumaDebe + Debe
  SumaHaber = SumaHaber + Debe
 'Nota de Credito por Fondos de Reserva
  Mifecha = BuscarFecha(MBoxFecha)
  TextConcepto14.Text = "(" & NumEmpresa & ") Nota de Credito Fondo de Reserva"
  ConceptoComp = TextConcepto14
  sSQL = "SELECT TP,Debitos,Creditos " _
       & "FROM Trans_Libretas " _
       & "WHERE Fecha = #" & Mifecha & "# " _
       & "AND T <> 'A' " _
       & "AND TP = 'DEFR' " _
       & "AND AC = " & Val(adFalse) & " "
  SelectAdodc AdoCaja, sSQL
  'MsgBox sSQL
  Debe = 0: Haber = 0
  With AdoCaja.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          Haber = Haber + .Fields("Creditos")
         .MoveNext
       Loop
   End If
  End With
  ProgBar.value = 11
  Debe = Haber
  Debe = Round(Debe, 2): Haber = Round(Haber, 2)
  Trans_No = 44
  InsertarAsientos AdoAsientos14, Cta_CajaG, 0, Haber, 0
  InsertarAsientos AdoAsientos14, Cta_Libretas, 0, 0, Haber
  SumaDebe = SumaDebe + Debe
  SumaHaber = SumaHaber + Debe
  
 'Sumatoria de Vencidos
  Mifecha = BuscarFecha(CLongFecha(CFechaLong(MBoxFecha) - 30))
  TextConcepto12.Text = "(" & NumEmpresa & ") Traspaso de Créditos Vigentes a Vencidos del " & Mifecha
  ConceptoComp = TextConcepto12
  
  sSQL = "SELECT * " _
       & "FROM Catalogo_Prestamo " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Cta_P_1_30 <> '" & Ninguno & "' " _
       & "ORDER BY CTP "
  SelectAdodc AdoAux, sSQL
  With AdoAux.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          Codigo = .Fields("CTP")
          FechaFin = BuscarFecha(CLongFecha(CFechaLong(MBoxFecha)))
          FechaIni = BuscarFecha(CLongFecha(CFechaLong(MBoxFecha) - 15))
          sSQL = "UPDATE Trans_Prestamos " _
               & "SET Cta_No_Devenga_Int = '" & .Fields("Cta_No_Devenga_Int") & "' " _
               & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
               & "AND Item = '" & NumEmpresa & "' " _
               & "AND TP = '" & Codigo & "' " _
               & "AND T = 'P' "
          ConectarAdoExecute sSQL
          
          FechaFin = BuscarFecha(CLongFecha(CFechaLong(MBoxFecha) - 16))
          FechaIni = BuscarFecha(CLongFecha(CFechaLong(MBoxFecha) - 30))
          sSQL = "UPDATE Trans_Prestamos " _
               & "SET Cta_V_1_30 = '" & .Fields("Cta_V_1_30") & "' " _
               & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
               & "AND Item = '" & NumEmpresa & "' " _
               & "AND V_1_30 = " & Val(adFalse) & " " _
               & "AND TP = '" & Codigo & "' " _
               & "AND T = 'P' "
          ConectarAdoExecute sSQL
          FechaFin = BuscarFecha(CLongFecha(CFechaLong(MBoxFecha) - 31))
          FechaIni = BuscarFecha(CLongFecha(CFechaLong(MBoxFecha) - 90))
          sSQL = "UPDATE Trans_Prestamos " _
               & "SET Cta_V_31_90 = '" & .Fields("Cta_V_31_90") & "' " _
               & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
               & "AND Item = '" & NumEmpresa & "' " _
               & "AND V_31_90 = " & Val(adFalse) & " " _
               & "AND TP = '" & Codigo & "' " _
               & "AND T = 'P' "
          ConectarAdoExecute sSQL
          
          FechaFin = BuscarFecha(CLongFecha(CFechaLong(MBoxFecha) - 91))
          FechaIni = BuscarFecha(CLongFecha(CFechaLong(MBoxFecha) - 180))
          sSQL = "UPDATE Trans_Prestamos " _
               & "SET Cta_V_91_180 = '" & .Fields("Cta_V_91_180") & "' " _
               & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
               & "AND Item = '" & NumEmpresa & "' " _
               & "AND V_91_180 = " & Val(adFalse) & " " _
               & "AND TP = '" & Codigo & "' " _
               & "AND T = 'P' "
          ConectarAdoExecute sSQL
          
          FechaFin = BuscarFecha(CLongFecha(CFechaLong(MBoxFecha) - 181))
          FechaIni = BuscarFecha(CLongFecha(CFechaLong(MBoxFecha) - 360))
          sSQL = "UPDATE Trans_Prestamos " _
               & "SET Cta_V_181_360 = '" & .Fields("Cta_V_181_360") & "' " _
               & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
               & "AND Item = '" & NumEmpresa & "' " _
               & "AND V_181_360 = " & Val(adFalse) & " " _
               & "AND TP = '" & Codigo & "' " _
               & "AND T = 'P' "
          ConectarAdoExecute sSQL
          FechaFin = BuscarFecha(CLongFecha(CFechaLong(MBoxFecha) - 361))
          sSQL = "UPDATE Trans_Prestamos " _
               & "SET Cta_V_Mas_360 = '" & .Fields("Cta_V_Mas_360") & "' " _
               & "WHERE Fecha <= #" & FechaFin & "# " _
               & "AND Item = '" & NumEmpresa & "' " _
               & "AND V_Mas_360 = " & Val(adFalse) & " " _
               & "AND TP = '" & Codigo & "' " _
               & "AND T = 'P' "
          ConectarAdoExecute sSQL
         .MoveNext
       Loop
   End If
  End With
 'Inicia los asientos de Vencidos
  Trans_No = 42
  Debe = 0: Haber = 0: Valor = 0
  Insertar_Prestamos_Vencidos "Cta_No_Devenga_Int", ""
  Insertar_Prestamos_Vencidos "Cta_V_1_30", "V_1_30"
  Insertar_Prestamos_Vencidos "Cta_V_31_90", "V_31_90"
  Insertar_Prestamos_Vencidos "Cta_V_91_180", "V_91_180"
  Insertar_Prestamos_Vencidos "Cta_V_181_360", "V_181_360"
  Insertar_Prestamos_Vencidos "Cta_V_Mas_360", "V_Mas_360"
   
''sSQL = "SELECT TP.TP,Cta_Prestamo,Cta_Vencidos,Capital " _
''     & "FROM Trans_Prestamos As TP," _
''     & "Catalogo_Prestamo As CP," _
''     & "Clientes_Datos_Extras As CDE " _
''     & "WHERE TP.Fecha <= #" & Mifecha & "# " _
''     & "AND TP.V = " & Val(adFalse) & " " _
''     & "AND TP.T = 'P' " _
''     & "AND TP.TP = CP.CTP " _
''     & "AND TP.Cuenta_No = CDE.Cuenta_No " _
''     & "AND CP.Item = '" & NumEmpresa & "' " _
''     & "ORDER BY TP.TP,Cta_Prestamo,Cta_Vencidos "
''SelectAdodc AdoCaja, sSQL
'''MsgBox "..--.."
''Debe = 0: Haber = 0: Valor = 0
''With AdoCaja.Recordset
'' If .RecordCount > 0 Then
''     Codigo = .Fields("Cta_Vencidos")
''     Codigo1 = .Fields("Cta_Prestamo")
''     Trans_No = 42
''     Do While Not .EOF
''        If Codigo1 <> .Fields("Cta_Prestamo") Then
''           InsertarAsientos AdoAsientos12, Codigo, 0, Valor, 0
''           InsertarAsientos AdoAsientos12, Codigo1, 0, 0, Valor
''           Debe = Debe + Valor
''           Codigo = .Fields("Cta_Vencidos")
''           Codigo1 = .Fields("Cta_Prestamo")
''           Valor = 0
''        End If
''        Valor = Valor + .Fields("Capital")
''       .MoveNext
''     Loop
''     InsertarAsientos AdoAsientos12, Codigo, 0, Valor, 0
''     InsertarAsientos AdoAsientos12, Codigo1, 0, 0, Valor
''     Debe = Debe + Valor
'' End If
''End With
  Haber = Debe
  Debe = Round(Debe, 2)
  Haber = Round(Haber, 2)
  Mifecha = BuscarFecha(MBoxFecha)
  SumaDebe = SumaDebe + Debe
  SumaHaber = SumaHaber + Debe
  ProgBar.value = 13
  
 'Notas de Debitos por Servicios basicos
  Trans_No = 45
  Mifecha = BuscarFecha(MBoxFecha)
  TextConcepto15.Text = "(" & NumEmpresa & ") Nota de dédito por Servicios Básicos de Luz, Agua y Teléfono"
  ConceptoComp = TextConcepto15
  sSQL = "SELECT TP,Debitos,Creditos " _
       & "FROM Trans_Libretas " _
       & "WHERE Fecha = #" & Mifecha & "# " _
       & "AND T <> 'A' " _
       & "AND TP IN ('NDAG','NDLZ','NDTF') " _
       & "AND AC = " & Val(adFalse) & " "
  SelectAdodc AdoCaja, sSQL
  'MsgBox sSQL
  Debe = 0: Haber = 0
  With AdoCaja.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          Debe = Debe + .Fields("Debitos")
         .MoveNext
       Loop
   End If
  End With
  ProgBar.value = 15
  Haber = Debe
  Debe = Round(Debe, 2): Haber = Round(Haber, 2)
  InsertarAsientos AdoAsientos15, Cta_Libretas, 0, Debe, 0
  InsertarAsientos AdoAsientos15, Cta_Servicios_Basicos, 0, 0, Debe
  SumaDebe = SumaDebe + Debe
  SumaHaber = SumaHaber + Debe
  
 'Ingresos y Egresos de Boveda
  Trans_No = 46
  Mifecha = BuscarFecha(MBoxFecha)
  TextConcepto16.Text = "(" & NumEmpresa & ") Ingresos y Egresos de Boveda"
  ConceptoComp = TextConcepto16
  sSQL = "SELECT TP,Debitos,Creditos " _
       & "FROM Trans_Libretas " _
       & "WHERE Fecha = #" & Mifecha & "# " _
       & "AND T <> 'A' " _
       & "AND TP = 'BOVE' " _
       & "AND Cuenta_No = 'BOVEDA' " _
       & "AND AC = " & Val(adFalse) & " "
  SelectAdodc AdoCaja, sSQL
  'MsgBox sSQL
  Debe = 0: Haber = 0
  With AdoCaja.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          Debe = Debe + .Fields("Debitos")
          Haber = Haber + .Fields("Creditos")
         .MoveNext
       Loop
   End If
  End With
  ProgBar.value = 16
  Debe = Round(Debe, 2): Haber = Round(Haber, 2)
  Diferencia = Haber - Debe
  InsertarAsientos AdoAsientos16, Cta_CajaG, 0, 0, Debe
  InsertarAsientos AdoAsientos16, Cta_CajaG, 0, Haber, 0
  'MsgBox Diferencia
  If Diferencia > 0 Then
     InsertarAsientos AdoAsientos16, Cta_Transito, 0, 0, Diferencia
  Else
     InsertarAsientos AdoAsientos16, Cta_Transito, 0, -Diferencia, 0
  End If
  SumaDebe = SumaDebe + Debe
  SumaHaber = SumaHaber + Debe
' Totales de los asientos:
' =======================
  sSQL = "SELECT ME,Fecha,TP,Cuenta_No,Debitos,Creditos,CodigoU " _
       & "FROM Trans_Libretas " _
       & "WHERE Fecha = #" & Mifecha & "# " _
       & "AND T = 'P' "
  If CheckUsuario.value = 1 Then sSQL = sSQL & "AND CodigoU = '" & SinEspaciosDer(DCUsuario.Text) & "' "
  sSQL = sSQL & "ORDER BY ME,Fecha,TP,Cuenta_No,Debitos DESC,Creditos "
  SelectAdodc AdoFCajaCheq, sSQL
  
  sSQL = "SELECT ME,Fecha,TP,Cuenta_No,Debitos,Creditos,CodigoU " _
       & "FROM Trans_Libretas " _
       & "WHERE Fecha = #" & Mifecha & "# " _
       & "AND T <> 'P' "
  If CheckUsuario.value = 1 Then sSQL = sSQL & "AND CodigoU = '" & SinEspaciosDer(DCUsuario.Text) & "' "
  sSQL = sSQL & "ORDER BY ME,Fecha,TP,Cuenta_No,Debitos DESC,Creditos "
  SelectAdodc AdoFCajaEfec, sSQL
  SumarIngEgrCaja AdoFCajaCheq, AdoFCajaEfec
  LabelIngresos.Caption = Format(SumaDebe, "#,##0.00")
  LabelEgresos.Caption = Format(SumaHaber, "#,##0.00")
  ProgBar.value = 100
  Command2.Enabled = True
  RatonNormal
  MsgBox "ASIENTOS AUTOMATICOS PROCESADOS" & vbCrLf & vbCrLf & vbCrLf _
       & "PROCEDA A COMPROBAR LOS ASIENTOS Y GRABELOS"
End Sub

Private Sub Command2_Click()
Dim Si_Grabe As Boolean
  Si_Grabe = True
  'sSQL = "SELECT * " _
  '     & "FROM Trans_Libretas " _
  '     & "WHERE Fecha = #" & MiFecha & "# " _
  '     & "AND AC = False "
  ' SelectAdodc AdoCaja, sSQL
  'If AdoCaja.Recordset.RecordCount > 0 Then
  'MsgBox Round(SumaDebe - SumaHaber)
  If SumaDebe = 0 And SumaHaber = 0 Then Si_Grabe = False
  If (Si_Grabe) And (Round(SumaDebe - SumaHaber, 2) = 0) Then
  RatonReloj
  Mifecha = BuscarFecha(MBoxFecha.Text)
  FechaComp = MBoxFecha.Text
  Co.T = Normal
  Co.TP = CompDiario
  Co.Fecha = MBoxFecha.Text
  Co.CodigoB = Ninguno
  Co.Efectivo = 0
  Co.Monto_Total = 0
  Co.Item = NumEmpresa
  Co.Usuario = CodigoUsuario
' Iniciar a agrabar comprobantes
  If AdoAsientos.Recordset.RecordCount > 0 Then
     Co.Numero = ReadSetDataNum("Diario", True, True)
     Co.Concepto = MidStrg(TextConcepto.Text, 1, 120)
     Co.T_No = 30
     GrabarComprobante Co
  End If
  If AdoAsientos1.Recordset.RecordCount > 0 Then
     Co.Numero = ReadSetDataNum("Diario", True, True)
     Co.Concepto = MidStrg(TextConcepto1.Text, 1, 120)
     Co.T_No = 31
     GrabarComprobante Co
  End If
  If AdoAsientos2.Recordset.RecordCount > 0 Then
     Co.Numero = ReadSetDataNum("Diario", True, True)
     Co.Concepto = MidStrg(TextConcepto2.Text, 1, 120)
     Co.T_No = 32
     GrabarComprobante Co
  End If
  If AdoAsientos3.Recordset.RecordCount > 0 Then
     Co.Numero = ReadSetDataNum("Diario", True, True)
     Co.Concepto = MidStrg(TextConcepto3.Text, 1, 120)
     Co.T_No = 33
     GrabarComprobante Co
  End If
  If AdoAsientos4.Recordset.RecordCount > 0 Then
     Co.Numero = ReadSetDataNum("Diario", True, True)
     Co.Concepto = MidStrg(TextConcepto4.Text, 1, 120)
     Co.T_No = 34
     GrabarComprobante Co
  End If
  If AdoAsientos5.Recordset.RecordCount > 0 Then
     Co.Numero = ReadSetDataNum("Diario", True, True)
     Co.Concepto = MidStrg(TextConcepto5.Text, 1, 120)
     Co.T_No = 35
     GrabarComprobante Co
  End If
  If AdoAsientos6.Recordset.RecordCount > 0 Then
     Co.Numero = ReadSetDataNum("Diario", True, True)
     Co.Concepto = MidStrg(TextConcepto6.Text, 1, 120)
     Co.T_No = 36
     GrabarComprobante Co
  End If
  If AdoAsientos7.Recordset.RecordCount > 0 Then
     Co.Numero = ReadSetDataNum("Diario", True, True)
     Co.Concepto = MidStrg(TextConcepto7.Text, 1, 120)
     Co.T_No = 37
     GrabarComprobante Co
  End If
  If AdoAsientos8.Recordset.RecordCount > 0 Then
     Co.Numero = ReadSetDataNum("Diario", True, True)
     Co.Concepto = MidStrg(TextConcepto8.Text, 1, 120)
     Co.T_No = 38
     GrabarComprobante Co
  End If
  If AdoAsientos9.Recordset.RecordCount > 0 Then
     Co.Numero = ReadSetDataNum("Diario", True, True)
     Co.Concepto = MidStrg(TextConcepto9.Text, 1, 120)
     Co.T_No = 39
     GrabarComprobante Co
  End If
  If AdoAsientos10.Recordset.RecordCount > 0 Then
     Co.Numero = ReadSetDataNum("Diario", True, True)
     Co.Concepto = MidStrg(TextConcepto10.Text, 1, 120)
     Co.T_No = 40
     GrabarComprobante Co
  End If
  If AdoAsientos11.Recordset.RecordCount > 0 Then
     Co.Numero = ReadSetDataNum("Diario", True, True)
     Co.Concepto = MidStrg(TextConcepto11.Text, 1, 120)
     Co.T_No = 41
     GrabarComprobante Co
  End If
  If AdoAsientos12.Recordset.RecordCount > 0 Then
     Co.Numero = ReadSetDataNum("Diario", True, True)
     Co.Concepto = MidStrg(TextConcepto12.Text, 1, 120)
     Co.T_No = 42
     GrabarComprobante Co
  End If
  If AdoAsientos13.Recordset.RecordCount > 0 Then
     Co.Numero = ReadSetDataNum("Diario", True, True)
     Co.Concepto = MidStrg(TextConcepto13.Text, 1, 120)
     Co.T_No = 43
     'MsgBox Co.Numero
     GrabarComprobante Co
  End If
  If AdoAsientos14.Recordset.RecordCount > 0 Then
     Co.Numero = ReadSetDataNum("Diario", True, True)
     Co.Concepto = MidStrg(TextConcepto14.Text, 1, 120)
     Co.T_No = 44
     'MsgBox Co.Numero
     GrabarComprobante Co
  End If
  If AdoAsientos15.Recordset.RecordCount > 0 Then
     Co.Numero = ReadSetDataNum("Diario", True, True)
     Co.Concepto = MidStrg(TextConcepto15.Text, 1, 120)
     Co.T_No = 45
     'MsgBox Co.Numero
     GrabarComprobante Co
  End If
  If AdoAsientos16.Recordset.RecordCount > 0 Then
     Co.Numero = ReadSetDataNum("Diario", True, True)
     Co.Concepto = MidStrg(TextConcepto16.Text, 1, 120)
     Co.T_No = 46
     'MsgBox Co.Numero
     GrabarComprobante Co
  End If
  If AdoAsientos17.Recordset.RecordCount > 0 Then
     Co.Numero = ReadSetDataNum("Diario", True, True)
     Co.Concepto = MidStrg(TextConcepto17.Text, 1, 120)
     Co.T_No = 47
     'MsgBox Co.Numero
     GrabarComprobante Co
  End If
  Mifecha = BuscarFecha(MBoxFecha.Text)
  sSQL = "SELECT AC " _
       & "FROM Trans_Libretas " _
       & "WHERE Fecha = #" & Mifecha & "# " _
       & "AND AC = " & Val(adFalse) & " "
  SelectAdodc AdoCaja, sSQL
  If AdoCaja.Recordset.RecordCount > 0 Then
     sSQL = "UPDATE Trans_Libretas " _
          & "SET AC = " & Val(adTrue) & " " _
          & "WHERE Fecha = #" & Mifecha & "# "
     ConectarAdoExecute sSQL
     
     sSQL = "UPDATE Trans_Saldo_Libretas " _
          & "SET AC =  " & Val(adTrue) & " " _
          & "WHERE Fecha <= #" & Mifecha & "# "
     ConectarAdoExecute sSQL
     
     sSQL = "UPDATE Trans_Bloqueos " _
          & "SET Dias = Dias - 1 " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND T <> 'A' " _
          & "AND Dias > 0 "
     ConectarAdoExecute sSQL
     
     sSQL = "UPDATE Trans_Libretas " _
          & "SET Dias = Dias - 1 " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND T <> 'A' " _
          & "AND Dias > 0 "
     ConectarAdoExecute sSQL
     
     sSQL = "UPDATE Trans_Bloqueos " _
          & "SET T = 'A' " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Banco <> '.' " _
          & "AND T <> 'A' " _
          & "AND Cheque <> 'FR' " _
          & "AND Dias <= 0 "
     ConectarAdoExecute sSQL
     
     Mifecha = BuscarFecha(CLongFecha(CFechaLong(MBoxFecha.Text) - 30))
     sSQL = "UPDATE Trans_Prestamos " _
          & "SET V = " & Val(adTrue) & ", " _
          & "Fecha_V = '" & MBoxFecha.Text & "' " _
          & "WHERE Fecha <= #" & Mifecha & "# " _
          & "AND T = 'P' " _
          & "AND V =  " & Val(adFalse) & " "
     ConectarAdoExecute sSQL
     
  End If
  sSQL = "UPDATE Trans_Intereses " _
       & "SET AC = " & Val(adTrue) & " " _
       & "WHERE Fecha = #" & Mifecha & "# " _
       & "AND AC = " & Val(adFalse) & " "
  ConectarAdoExecute sSQL
  RatonNormal
  Unload FAsientos
'  End If
  Else
     MsgBox "Este día ya esta cerrado o " & vbCrLf & vbCrLf _
          & "No cuadran las transacciones"
  End If
End Sub

Private Sub Command5_Click()
  Unload FAsientos
End Sub

Private Sub Form_Activate()
  Mifecha = BuscarFecha(FechaSistema)
  Command2.Enabled = False
  TipoDoc = CompDiario
  If Supervisor = False Then
     If CNivel(3) Then
        Command2.Enabled = False
     End If
  End If
  IniciarAsientosCaja
  NumComp = ReadSetDataNum("Diario", True, False)
  FormatoMaskCta MBoxCtaI
  sSQL = "SELECT Nombre_Completo & '  ' & Codigo As NombUsuario " _
       & "FROM Accesos " _
       & "WHERE MidStrg(Codigo,1,5) <> 'ACCES' " _
       & "ORDER BY Nombre_Completo "
  SelectDBCombo DCUsuario, AdoUsuario, sSQL, "NombUsuario"
  'SeteosCtas AdoEmp
  Label9.Caption = "Cta. Sobrante"
  RatonNormal
  MBoxFecha.SetFocus
  FAsientos.Caption = "ASIENTOS CONTABLES DEL CIERRE"
  FAsientos.WindowState = 2
End Sub

Private Sub Form_Load()
'CentrarForm FAsientos
FAsientos.Caption = "Espere uno segundos... "
FAsientos.WindowState = 1
ConectarAdodc AdoAux
ConectarAdodc AdoAsientos
ConectarAdodc AdoAsientos1
ConectarAdodc AdoAsientos2
ConectarAdodc AdoAsientos3
ConectarAdodc AdoAsientos4
ConectarAdodc AdoAsientos5
ConectarAdodc AdoAsientos6
ConectarAdodc AdoAsientos7
ConectarAdodc AdoAsientos8
ConectarAdodc AdoAsientos9
ConectarAdodc AdoAsientos10
ConectarAdodc AdoAsientos11
ConectarAdodc AdoAsientos12
ConectarAdodc AdoAsientos13
ConectarAdodc AdoAsientos14
ConectarAdodc AdoAsientos15
ConectarAdodc AdoAsientos16
ConectarAdodc AdoAsientos17
ConectarAdodc AdoCaja
ConectarAdodc AdoUsuario
ConectarAdodc AdoFCajaCheq
ConectarAdodc AdoFCajaEfec
End Sub

Private Sub MBoxFecha_GotFocus()
  MarcarTexto MBoxFecha
End Sub

Private Sub MBoxFecha_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBoxFecha_LostFocus()
  FechaValida MBoxFecha
End Sub

Private Sub OpcF_Click()
  Label9.Caption = "Cta. Faltante"
End Sub

Private Sub OpcS_Click()
  Label9.Caption = "Cta. Sobrante"
End Sub

Private Sub TextValor_GotFocus()
  TextValor.Text = ""
End Sub

Private Sub TextValor_LostFocus()
  TextoValido TextValor, True
End Sub

Public Sub IniciarAsientosCaja()
  Trans_No = 30: IniciarAsientosDe DGAsientos, AdoAsientos
  Trans_No = 31: IniciarAsientosDe DGAsientos1, AdoAsientos1
  Trans_No = 32: IniciarAsientosDe DGAsientos2, AdoAsientos2
  Trans_No = 33: IniciarAsientosDe DGAsientos3, AdoAsientos3
  Trans_No = 34: IniciarAsientosDe DGAsientos4, AdoAsientos4
  Trans_No = 35: IniciarAsientosDe DGAsientos5, AdoAsientos5
  Trans_No = 36: IniciarAsientosDe DGAsientos6, AdoAsientos6
  Trans_No = 37: IniciarAsientosDe DGAsientos7, AdoAsientos7
  Trans_No = 38: IniciarAsientosDe DGAsientos8, AdoAsientos8
  Trans_No = 39: IniciarAsientosDe DGAsientos9, AdoAsientos9
  Trans_No = 40: IniciarAsientosDe DGAsientos10, AdoAsientos10
  Trans_No = 41: IniciarAsientosDe DGAsientos11, AdoAsientos11
  Trans_No = 42: IniciarAsientosDe DGAsientos12, AdoAsientos12
  Trans_No = 43: IniciarAsientosDe DGAsientos13, AdoAsientos13
  Trans_No = 44: IniciarAsientosDe DGAsientos14, AdoAsientos14
  Trans_No = 45: IniciarAsientosDe DGAsientos15, AdoAsientos15
  Trans_No = 46: IniciarAsientosDe DGAsientos16, AdoAsientos16
  Trans_No = 47: IniciarAsientosDe DGAsientos17, AdoAsientos17
End Sub

Public Sub Insertar_Prestamos_Vencidos(Cta_Vencido As String, BVencido As String)
  sSQL = "SELECT TP," & Cta_Vencido & ", SUM(Capital) As TCapital " _
       & "FROM Trans_Prestamos " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND " & Cta_Vencido & " <> '" & Ninguno & "' "
  If Len(BVencido) > 1 Then sSQL = sSQL & "AND " & BVencido & " = " & Val(adFalse) & " "
  sSQL = sSQL & "AND T = 'P' " _
       & "GROUP BY TP," & Cta_Vencido & " "
  'MsgBox sSQL
  SelectAdodc AdoCaja, sSQL
  With AdoCaja.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          Codigo = .Fields(Cta_Vencido)
          Valor = .Fields("TCapital")
          If Len(BVencido) > 1 Then
             DetalleComp = "Crédito Vencido de " & .Fields("TP") & " de " & Replace(MidStrg(Cta_Vencido, 7, Len(Cta_Vencido)), "_", " a ") & " días"
          Else
             DetalleComp = "Crédito Vencido que No devengan Interes de " & .Fields("TP")
          End If
          If Len(Codigo) > 1 And Valor > 0 Then
             InsertarAsientos AdoAsientos12, Codigo, 0, Valor, 0
             Debe = Debe + Valor
          End If
         .MoveNext
       Loop
   End If
  End With
  sSQL = "SELECT TP,Cta, SUM(Capital) As TCapital " _
       & "FROM Trans_Prestamos " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND " & Cta_Vencido & " <> '" & Ninguno & "' "
  If Len(BVencido) > 1 Then sSQL = sSQL & "AND " & BVencido & " = " & Val(adFalse) & " "
  sSQL = sSQL & "AND T = 'P' " _
       & "GROUP BY TP,Cta "
  SelectAdodc AdoCaja, sSQL
  With AdoCaja.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          Codigo1 = .Fields("Cta")
          Valor = .Fields("TCapital")
          If Len(BVencido) > 1 Then
             DetalleComp = "Crédito Vencido de " & .Fields("TP") & " de " & Replace(MidStrg(Cta_Vencido, 7, Len(Cta_Vencido)), "_", " a ") & " días"
          Else
             DetalleComp = "Crédito Vencido que No devengan Interes de " & .Fields("TP")
          End If
          If Len(Codigo1) > 1 And Valor > 0 Then
             InsertarAsientos AdoAsientos12, Codigo1, 0, 0, Valor
             Haber = Haber + Valor
          End If
         .MoveNext
       Loop
   End If
  End With
End Sub
