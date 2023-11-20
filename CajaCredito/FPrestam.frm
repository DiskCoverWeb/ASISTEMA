VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Begin VB.Form FPrestamo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "LIQUIDACION DE PRESTAMOS"
   ClientHeight    =   7710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11250
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7710
   ScaleWidth      =   11250
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Caption         =   "Cuentas Abiertas"
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
      Height          =   1485
      Left            =   4305
      TabIndex        =   43
      Top             =   525
      Visible         =   0   'False
      Width           =   6840
      Begin MSDataGridLib.DataGrid DGCtaAhorro 
         Bindings        =   "FPrestam.frx":0000
         Height          =   1170
         Left            =   105
         TabIndex        =   44
         Top             =   210
         Width           =   6630
         _ExtentX        =   11695
         _ExtentY        =   2064
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
               LCID            =   3082
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
               LCID            =   3082
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
   End
   Begin VB.TextBox TxtRUCS 
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
      MaxLength       =   30
      TabIndex        =   12
      Top             =   1680
      Width           =   1800
   End
   Begin MSDataListLib.DataCombo DCTipoPrestamo 
      Bindings        =   "FPrestam.frx":001B
      DataSource      =   "AdoTipoPrestamo"
      Height          =   315
      Left            =   1995
      TabIndex        =   14
      Top             =   1680
      Width           =   5055
      _ExtentX        =   8916
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
   Begin VB.CommandButton Command3 
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
      Height          =   1065
      Left            =   9870
      Picture         =   "FPrestam.frx":0039
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   3465
      Width           =   1170
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4530
      Left            =   105
      TabIndex        =   21
      Top             =   3045
      Width           =   11040
      _ExtentX        =   19473
      _ExtentY        =   7990
      _Version        =   393216
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
      TabCaption(0)   =   "&Tabla de Pagos"
      TabPicture(0)   =   "FPrestam.frx":0A2F
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "DGTabla"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Command4"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "&Datos de los Garantes"
      TabPicture(1)   =   "FPrestam.frx":0A4B
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label19"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label17"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label15"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label16"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label14"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Command2"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Command1"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "TxtNombresC"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "MBoxCIC"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "TxtTelefonoC"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "TxtLugarTrabC"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "TxtDirC"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Frame1"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "DGGarantes"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "DLGarantes"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).ControlCount=   15
      TabCaption(2)   =   "&Prestamos Otorgados"
      TabPicture(2)   =   "FPrestam.frx":0A67
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "DGListPrest"
      Tab(2).Control(1)=   "DGResumen"
      Tab(2).ControlCount=   2
      Begin VB.CommandButton Command4 
         Caption         =   "&Grabar Prestamo Antiguo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1065
         Left            =   9765
         Picture         =   "FPrestam.frx":0A83
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   1575
         Width           =   1170
      End
      Begin MSDataListLib.DataList DLGarantes 
         Bindings        =   "FPrestam.frx":0EC5
         DataSource      =   "AdoLGarantes"
         Height          =   2400
         Left            =   -74895
         TabIndex        =   42
         Top             =   1680
         Visible         =   0   'False
         Width           =   7995
         _ExtentX        =   14102
         _ExtentY        =   4233
         _Version        =   393216
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
      Begin MSDataGridLib.DataGrid DGGarantes 
         Bindings        =   "FPrestam.frx":0EE0
         Height          =   1485
         Left            =   -74895
         TabIndex        =   41
         Top             =   2940
         Width           =   7995
         _ExtentX        =   14102
         _ExtentY        =   2619
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
      Begin MSDataGridLib.DataGrid DGListPrest 
         Bindings        =   "FPrestam.frx":0EFA
         Height          =   2115
         Left            =   -74895
         TabIndex        =   40
         Top             =   2310
         Width           =   9570
         _ExtentX        =   16880
         _ExtentY        =   3731
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
      Begin MSDataGridLib.DataGrid DGResumen 
         Bindings        =   "FPrestam.frx":0F1A
         Height          =   1905
         Left            =   -74895
         TabIndex        =   39
         Top             =   420
         Width           =   9570
         _ExtentX        =   16880
         _ExtentY        =   3360
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
      Begin MSDataGridLib.DataGrid DGTabla 
         Bindings        =   "FPrestam.frx":0F34
         Height          =   4005
         Left            =   105
         TabIndex        =   38
         Top             =   420
         Width           =   9570
         _ExtentX        =   16880
         _ExtentY        =   7064
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
      Begin VB.Frame Frame1 
         Height          =   540
         Left            =   -74895
         TabIndex        =   22
         Top             =   420
         Width           =   7995
         Begin VB.OptionButton OpcC 
            Caption         =   "Conyugue"
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
            Left            =   2625
            TabIndex        =   33
            Top             =   210
            Width           =   1905
         End
         Begin VB.OptionButton OpcG 
            Caption         =   "Garante"
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
            Left            =   210
            TabIndex        =   37
            Top             =   210
            Value           =   -1  'True
            Width           =   2115
         End
      End
      Begin VB.TextBox TxtDirC 
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
         MaxLength       =   50
         TabIndex        =   32
         Top             =   2625
         Width           =   7995
      End
      Begin VB.TextBox TxtLugarTrabC 
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
         Left            =   -70170
         MaxLength       =   30
         TabIndex        =   30
         Top             =   1995
         Width           =   3270
      End
      Begin VB.TextBox TxtTelefonoC 
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
         Left            =   -73425
         MaxLength       =   10
         TabIndex        =   29
         Top             =   1995
         Width           =   3270
      End
      Begin MSMask.MaskEdBox MBoxCIC 
         Height          =   330
         Left            =   -74895
         TabIndex        =   28
         ToolTipText     =   "Formato de Fecha: DD/MM/AA"
         Top             =   1995
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
         Format          =   "CCCCCCCCCC"
         Mask            =   "##########"
         PromptChar      =   "0"
      End
      Begin VB.TextBox TxtNombresC 
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
         MaxLength       =   60
         TabIndex        =   24
         Top             =   1365
         Width           =   7995
      End
      Begin VB.CommandButton Command1 
         Caption         =   "G&rabar Garante"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1065
         Left            =   -65235
         Picture         =   "FPrestam.frx":0F4B
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   1575
         Width           =   1170
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Grabar Prestamo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1065
         Left            =   -65235
         Picture         =   "FPrestam.frx":138D
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   2730
         Width           =   1170
      End
      Begin VB.Label Label14 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Direccion del Trabajo"
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
         Top             =   2310
         Width           =   7995
      End
      Begin VB.Label Label16 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Lugar de Trabajo"
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
         Left            =   -70170
         TabIndex        =   27
         Top             =   1680
         Width           =   3270
      End
      Begin VB.Label Label15 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Telefono"
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
         Left            =   -73425
         TabIndex        =   26
         Top             =   1680
         Width           =   3270
      End
      Begin VB.Label Label17 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Identificacion"
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
         TabIndex        =   25
         Top             =   1680
         Width           =   1485
      End
      Begin VB.Label Label19 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " &Nombres                [Ctrl-B: Buscar Garantes]"
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
         TabIndex        =   23
         Top             =   1050
         Width           =   7995
      End
   End
   Begin MSAdodcLib.Adodc AdoPrestamo 
      Height          =   330
      Left            =   420
      Top             =   3360
      Visible         =   0   'False
      Width           =   2430
      _ExtentX        =   4286
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
      Caption         =   "Prestamo"
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
   Begin MSMask.MaskEdBox MBoxCuenta 
      Height          =   435
      Left            =   5145
      TabIndex        =   3
      ToolTipText     =   "Formato de Fecha: DD/MM/AA"
      Top             =   105
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   767
      _Version        =   393216
      ForeColor       =   192
      AllowPrompt     =   -1  'True
      AutoTab         =   -1  'True
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "CCCCCCCC-C"
      Mask            =   "########-#"
      PromptChar      =   "0"
   End
   Begin MSMask.MaskEdBox MBoxFecha 
      Height          =   435
      Left            =   1365
      TabIndex        =   1
      ToolTipText     =   "Formato de Fecha: DD/MM/AA"
      Top             =   105
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   767
      _Version        =   393216
      AllowPrompt     =   -1  'True
      AutoTab         =   -1  'True
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   14.25
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
   Begin VB.TextBox TxtRazonSocial 
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
      Left            =   7140
      MaxLength       =   49
      TabIndex        =   10
      Top             =   945
      Width           =   3585
   End
   Begin VB.TextBox TxtNombresS 
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
      MaxLength       =   60
      TabIndex        =   9
      Top             =   945
      Width           =   6945
   End
   Begin VB.TextBox TextInt 
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
      Left            =   7140
      MaxLength       =   49
      MultiLine       =   -1  'True
      TabIndex        =   16
      Text            =   "FPrestam.frx":17CF
      Top             =   1680
      Width           =   750
   End
   Begin VB.TextBox TextMonto 
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
      Left            =   8820
      MaxLength       =   49
      MultiLine       =   -1  'True
      TabIndex        =   20
      Text            =   "FPrestam.frx":17D3
      Top             =   1680
      Width           =   1905
   End
   Begin VB.TextBox TextMeses 
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
      Left            =   7980
      MaxLength       =   49
      MultiLine       =   -1  'True
      TabIndex        =   18
      Text            =   "FPrestam.frx":17D7
      Top             =   1680
      Width           =   750
   End
   Begin MSAdodcLib.Adodc AdoTipoPrestamo 
      Height          =   330
      Left            =   420
      Top             =   3675
      Visible         =   0   'False
      Width           =   2430
      _ExtentX        =   4286
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
      Caption         =   "TipoPrestamo"
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
      Left            =   420
      Top             =   3990
      Visible         =   0   'False
      Width           =   2430
      _ExtentX        =   4286
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
   Begin MSAdodcLib.Adodc AdoGarantes 
      Height          =   330
      Left            =   420
      Top             =   4305
      Visible         =   0   'False
      Width           =   2430
      _ExtentX        =   4286
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
      Caption         =   "Garantes"
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
   Begin MSAdodcLib.Adodc AdoTabla 
      Height          =   330
      Left            =   420
      Top             =   4620
      Visible         =   0   'False
      Width           =   2430
      _ExtentX        =   4286
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
      Caption         =   "Tabla"
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
   Begin MSAdodcLib.Adodc AdoCtaNo 
      Height          =   330
      Left            =   420
      Top             =   4935
      Visible         =   0   'False
      Width           =   2430
      _ExtentX        =   4286
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
      Caption         =   "CtaNo"
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
   Begin MSAdodcLib.Adodc AdoCta 
      Height          =   330
      Left            =   2835
      Top             =   3360
      Visible         =   0   'False
      Width           =   2430
      _ExtentX        =   4286
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
      Caption         =   "Cta"
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
   Begin MSAdodcLib.Adodc AdoResumenP 
      Height          =   330
      Left            =   2835
      Top             =   3675
      Visible         =   0   'False
      Width           =   2430
      _ExtentX        =   4286
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
      Caption         =   "ResumenP"
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
   Begin MSAdodcLib.Adodc AdoListarPrestamo 
      Height          =   330
      Left            =   2835
      Top             =   3990
      Visible         =   0   'False
      Width           =   2430
      _ExtentX        =   4286
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
      Caption         =   "ListarPrestamo"
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
   Begin MSAdodcLib.Adodc AdoLGarantes 
      Height          =   330
      Left            =   2835
      Top             =   4305
      Visible         =   0   'False
      Width           =   2430
      _ExtentX        =   4286
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
      Caption         =   "LGarantes"
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
   Begin MSAdodcLib.Adodc AdoCtaAhorro 
      Height          =   330
      Left            =   2835
      Top             =   4620
      Visible         =   0   'False
      Width           =   2430
      _ExtentX        =   4286
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
      Caption         =   "CtaAhorro"
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
   Begin MSAdodcLib.Adodc AdoTTabla 
      Height          =   330
      Left            =   2835
      Top             =   4935
      Visible         =   0   'False
      Width           =   2430
      _ExtentX        =   4286
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
      Caption         =   "TTabla"
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
   Begin MSDataGridLib.DataGrid DGTTabla 
      Bindings        =   "FPrestam.frx":17DB
      Height          =   855
      Left            =   1995
      TabIndex        =   46
      Top             =   2100
      Width           =   9150
      _ExtentX        =   16140
      _ExtentY        =   1508
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   19
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
         Size            =   9.75
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
   Begin VB.Label LabelCredNo 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "999999999"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   435
      Left            =   105
      TabIndex        =   6
      Top             =   2520
      Width           =   1800
   End
   Begin VB.Label Label11 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Credito No."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   105
      TabIndex        =   5
      Top             =   2100
      Width           =   1800
   End
   Begin VB.Label LabelEstado 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Nombres"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   435
      Left            =   6930
      TabIndex        =   4
      Top             =   105
      Width           =   4215
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " CUENTA No."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3045
      TabIndex        =   2
      Top             =   105
      Width           =   2115
   End
   Begin VB.Label Label25 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Representante"
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
      Left            =   7140
      TabIndex        =   8
      Top             =   630
      Width           =   3585
   End
   Begin VB.Label Label9 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TIPO DE PRESTAMO"
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
      Left            =   1995
      TabIndex        =   13
      Top             =   1365
      Width           =   5055
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Interes"
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
      Left            =   7140
      TabIndex        =   15
      Top             =   1365
      Width           =   750
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Monto de Prestamo"
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
      Left            =   8820
      TabIndex        =   19
      Top             =   1365
      Width           =   1905
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Meses"
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
      Left            =   7980
      TabIndex        =   17
      Top             =   1365
      Width           =   750
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Nombres"
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
      Top             =   630
      Width           =   6945
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " C.I./R.U.C."
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
      TabIndex        =   11
      Top             =   1365
      Width           =   1800
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " FECHA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   1275
   End
End
Attribute VB_Name = "FPrestamo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub Lista_De_Garantes()
  sSQL = "SELECT * " _
       & "FROM Clientes_Datos_Extras " _
       & "WHERE Cuenta_No = '" & MBoxCuenta & "' " _
       & "AND Tipo_Dato = 'GARANTES' " _
       & "AND TP = '" & TipoDoc & "' " _
       & "AND Credito_No = '" & Contrato_No & "' " _
       & "ORDER BY Num,GC "
  SelectDataGrid DGGarantes, AdoGarantes, sSQL
  
  sSQL = "SELECT Beneficiario As NombGarantes " _
       & "FROM Clientes_Datos_Extras " _
       & "WHERE Cuenta_No = '" & MBoxCuenta & "' " _
       & "AND Tipo_Dato = 'GARANTES' "
  SelectDBList DLGarantes, AdoLGarantes, sSQL, "NombGarantes"
  RatonNormal
End Sub

Public Sub ListarCuenta(Cuenta_No As String)
   TipoDoc = SinEspaciosIzq(DCTipoPrestamo)
   Contrato_No = LabelCredNo.Caption
   LabelEstado.Caption = ""
   TxtNombresS.Text = ""
   TxtRUCS.Text = "0000000000000"
   TxtNombresS.Text = ""
   TxtRazonSocial.Text = ""
   sSQL = "SELECT Cl.Cliente,Cl.Direccion,Cl.CI_RUC,Cl.Representante,Cl.Fecha_N,C.* " _
        & "FROM Clientes_Datos_Extras As C,Clientes As Cl " _
        & "WHERE C.Item = '" & NumEmpresa & "' " _
        & "AND C.Cuenta_No = '" & Cuenta_No & "' " _
        & "AND C.Tipo_Dato = 'LIBRETAS' " _
        & "AND C.Codigo = Cl.Codigo "
   SelectAdodc AdoCta, sSQL
   With AdoCta.Recordset
    If .RecordCount > 0 Then
        Select Case .Fields("T")
          Case Anulado: LabelEstado.Caption = "ANULADA"
          Case Procesado: LabelEstado.Caption = "ABIERTA"
          Case Normal: LabelEstado.Caption = "NORMAL"
        End Select
        'MBoxFecha.Text = .Fields("Fecha")
        TxtNombresS.Text = .Fields("Cliente")
        TxtRUCS.Text = .Fields("CI_RUC")
        TxtRazonSocial.Text = .Fields("Representante")
        Edad_Persona = Year(FechaSistema) - Year(.Fields("Fecha_N"))
        Lista_De_Garantes
        sSQL = "SELECT T,TP,Credito_No,Cuenta_No,Dia,Fecha,Saldo_Pendiente " _
             & "FROM Prestamos " _
             & "WHERE Cuenta_No = '" & Cuenta_No & "' " _
             & "ORDER BY T,TP,Credito_No,Fecha "
        SelectDataGrid DGResumen, AdoResumenP, sSQL
        If LabelEstado.Caption = "ANULADA" Then
            MsgBox "Cuenta Cerrada, No podra realizar Créditos"
            MBoxCuenta.SetFocus
        Else
            TxtNombresS.SetFocus
        End If
    Else
        MsgBox "Cuenta No existente, No podra realizar Créditos"
        Frame2.Visible = True
        DGCtaAhorro.Visible = True
    End If
   End With
End Sub

Private Sub Command1_Click()
TipoDoc = SinEspaciosIzq(DCTipoPrestamo.Text)
Contrato_No = LabelCredNo.Caption
Contador = 1
Titulo = "Pregunta de Grabacion"
Mensajes = "Seguro de Grabar Garante"
If BoxMensaje = vbYes Then
   RatonReloj
   If OpcG.value Then Si_No = True Else Si_No = False
   sSQL = "SELECT * " _
        & "FROM Clientes_Datos_Extras " _
        & "WHERE Cuenta_No = '" & MBoxCuenta & "' " _
        & "AND TP = '" & TipoDoc & "' " _
        & "AND Tipo_Dato = 'GARANTES' " _
        & "AND Credito_No = '" & Contrato_No & "' " _
        & "ORDER BY Num,GC DESC "
   SelectAdodc AdoCta, sSQL
   With AdoCta.Recordset
       If .RecordCount > 0 Then
          .MoveLast
           If Si_No Then
              Contador = .Fields("Num") + 1
           Else
              Contador = .Fields("Num")
           End If
       End If
      .AddNew
       SetFields AdoCta, "GC", Si_No
       SetFields AdoCta, "Num", Contador
       SetFields AdoCta, "TP", TipoDoc
       SetFields AdoCta, "Fecha_Registro", FechaSistema
       SetFields AdoCta, "Credito_No", Contrato_No
       SetFields AdoCta, "Cuenta_No", MBoxCuenta.Text
       SetFields AdoCta, "Beneficiario", TxtNombresC.Text
       SetFields AdoCta, "CI", MBoxCIC.Text
       SetFields AdoCta, "Lugar_Trabajo", TxtLugarTrabC.Text
       SetFields AdoCta, "Telefono", TxtTelefonoC.Text
       SetFields AdoCta, "Direccion", TxtDirC.Text
      .Update
   End With
   RatonNormal
   MsgBox "Grabacion Exitosa"
   Lista_De_Garantes
   RatonNormal
End If
End Sub

Private Sub Command2_Click()
TipoDoc = SinEspaciosIzq(DCTipoPrestamo.Text)
Contrato_No = LabelCredNo.Caption
Titulo = "Pregunta de Grabacion"
Mensajes = "Seguro de Grabar Liquidacion de Prestamo"
If BoxMensaje = vbYes Then
   RatonReloj
   With AdoTabla.Recordset
    If .RecordCount > 0 Then
       .MoveFirst
        Numero = ReadSetDataNum("Liquidacion", True, True)
        Contrato_No = NumEmpresa & Format(Numero, "0000000")
        RatonReloj
        SetAdoAddNew "Prestamos"
        SetAdoFields "T", Normal
        SetAdoFields "Tasa", Round(CCur(TextInt), 2)
        SetAdoFields "TP", TipoDoc
        SetAdoFields "ME", False
        SetAdoFields "Credito_No", Contrato_No
        SetAdoFields "Cuenta_No", MBoxCuenta
        SetAdoFields "Meses", CCur(TextMeses)
        SetAdoFields "Dia", .Fields("Dia")
        SetAdoFields "Fecha", .Fields("Fecha")
        SetAdoFields "Interes", .Fields("Interes")
        SetAdoFields "Capital", .Fields("Capital")
        SetAdoFields "Pagos", .Fields("Pagos")
        SetAdoFields "Saldo_Pendiente", Round(CCur(TextMonto), 2)
        SetAdoFields "Plazo", CInt(TextMeses.Text) * 30
        SetAdoFields "Item", NumEmpresa
        SetAdoUpdate
        MBoxCuenta.SetFocus
        RatonNormal
        Numero = ReadSetDataNum("Prestamos", True, False)
        LabelCredNo.Caption = NumEmpresa & Format(Numero, "00000")
        Contrato_No = LabelCredNo.Caption
    End If
   End With
   sSQL = "DELETE * " _
        & "FROM Asiento_P " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND CodigoU = '" & CodigoUsuario & "' "
   ConectarAdoExecute sSQL
End If
sSQL = "SELECT * " _
     & "FROM Asiento_P " _
     & "WHERE Item = '" & NumEmpresa & "' " _
     & "AND CodigoU = '" & CodigoUsuario & "' " _
     & "ORDER BY T_No "
SelectDataGrid DGTabla, AdoTabla, sSQL
End Sub

Private Sub Command3_Click()
  RatonNormal
  Unload FPrestamo
End Sub

Private Sub Command4_Click()
If ClaveSupervisor Then
TipoDoc = SinEspaciosIzq(DCTipoPrestamo.Text)
Contrato_No = LabelCredNo.Caption
Titulo = "Pregunta de Grabacion"
Mensajes = "Seguro de Grabar Liquidacion de Prestamo"
If BoxMensaje = vbYes Then
   NumMeses = Val(InputBox("Mes Pendiente", "ACREDITACION DE PRESTAMOS", "0"))
   RatonReloj
   With AdoTabla.Recordset
    If .RecordCount > 0 Then
       .MoveFirst
        Numero = ReadSetDataNum("Prestamos", True, True)
        Contrato_No = NumEmpresa & Format(Numero, "0000000")
        RatonReloj
        Do While Not .EOF
           If .Fields("T_No") = 0 Then
               SetAdoAddNew "Prestamos"
               SetAdoFields "T", Procesado
               SetAdoFields "Tasa", Round(CCur(TextInt.Text) / 100, 2)
               SetAdoFields "TP", TipoDoc
               SetAdoFields "ME", False
               SetAdoFields "Credito_No", Contrato_No
               SetAdoFields "Cuenta_No", MBoxCuenta.Text
               SetAdoFields "Meses", CCur(TextMeses.Text)
               SetAdoFields "Dia", .Fields("Dia")
               SetAdoFields "Fecha", .Fields("Fecha")
               SetAdoFields "Interes", .Fields("Interes")
               SetAdoFields "Capital", .Fields("Capital")
               SetAdoFields "Pagos", .Fields("Pagos")
               SetAdoFields "Saldo_Pendiente", Round(CCur(TextMonto.Text), 2)
               SetAdoFields "Encaje", Round(CCur(TextMonto.Text) * 0.2, 2)
               SetAdoFields "Plazo", CInt(TextMeses.Text) * 30
               SetAdoFields "Item", NumEmpresa
               SetAdoUpdate
               Valor = CCur(TextMonto.Text) * 0.2
               SetAdoAddNew "Trans_Bloqueos"
               SetAdoFields "T", Normal
               SetAdoFields "Fecha", .Fields("Fecha")
               SetAdoFields "Cuenta_No", MBoxCuenta.Text
               SetAdoFields "Valor", Round(Valor, 2)
               SetAdoFields "Item", NumEmpresa
               SetAdoUpdate
           Else
               SetAdoAddNew "Trans_Prestamos"
               SetAdoFields "Fecha_V", Ninguno
               SetAdoFields "TP", TipoDoc
               SetAdoFields "ME", False
               SetAdoFields "Credito_No", Contrato_No
               SetAdoFields "Cuenta_No", MBoxCuenta.Text
               SetAdoFields "Cuota_No", .Fields("T_No")
               SetAdoFields "Dia", .Fields("Dia")
               SetAdoFields "Fecha", .Fields("Fecha")
               SetAdoFields "Fecha_C", FechaSistema
               SetAdoFields "Capital", .Fields("Capital")
               SetAdoFields "Interes", .Fields("Interes")
               SetAdoFields "Comision", .Fields("Comision")
               SetAdoFields "Pagos", .Fields("Pagos")
               SetAdoFields "Saldo", .Fields("Saldo")
               SetAdoFields "Item", NumEmpresa
               If .Fields("T_No") < NumMeses Then
                  SetAdoFields "T", Cancelado
               Else
                  SetAdoFields "T", Procesado
               End If
               SetAdoUpdate
           End If
          .MoveNext
        Loop
        MBoxCuenta.SetFocus
        RatonNormal
        Numero = ReadSetDataNum("Prestamos", True, False)
        LabelCredNo.Caption = NumEmpresa & Format(Numero, "00000")
        Contrato_No = LabelCredNo.Caption
    End If
   End With
   sSQL = "DELETE * " _
        & "FROM Asiento_P " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND CodigoU = '" & CodigoUsuario & "' "
   ConectarAdoExecute sSQL
End If
sSQL = "SELECT * " _
     & "FROM Asiento_P " _
     & "WHERE Item = '" & NumEmpresa & "' " _
     & "AND CodigoU = '" & CodigoUsuario & "' " _
     & "ORDER BY T_No "
SelectDataGrid DGTabla, AdoTabla, sSQL
End If
End Sub

Private Sub DCTipoPrestamo_LostFocus()
TipoDoc = SinEspaciosIzq(DCTipoPrestamo.Text)
Contrato_No = LabelCredNo.Caption
  sSQL = "DELETE * " _
       & "FROM Asiento_P " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' "
   ConectarAdoExecute sSQL
   With AdoTipoPrestamo.Recordset
    If .RecordCount > 0 Then
       .MoveFirst
       .Find ("CTP = '" & TipoDoc & "' ")
        If Not .EOF Then
           Si_No = .Fields("DM")
           If Si_No Then Label5.Caption = " Dias" Else Label5.Caption = " Meses"
        End If
    End If
   End With
End Sub

Private Sub DGCtaAhorro_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyEscape Then Frame2.Visible = False
  If KeyCode = vbKeyReturn Then
     MBoxCuenta.Text = DGCtaAhorro.Columns(1).Text
     Frame2.Visible = False
     MBoxCuenta.SetFocus
     SiguienteControl
  End If
End Sub

Private Sub DGResumen_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF1 Then
     DGResumen.Col = 2: CuentaBanco = DGResumen.Text
     DGResumen.Col = 3: Cuenta = DGResumen.Text
     DGResumen.Col = 5: FechaTexto = DGResumen.Text
     DGResumen.Col = 6: Total = CCur(DGResumen.Text)
     Mensajes = "Fecha Prestamo: " & FechaTexto _
              & ", Cuenta No.  " & Cuenta & Chr(13) _
              & "Credito No. " & CuentaBanco _
              & ",  Valor: " & Format(Total, "#,##0.00")
     MsgBox Mensajes
     sSQL = "SELECT * " _
          & "FROM Prestamos " _
          & "WHERE Cuenta_No = '" & Cuenta & "' " _
          & "AND Credito_No = '" & Contrato_No & "' " _
          & "AND T_No <> 0 " _
          & "ORDER BY Credito_No,TP,Mes_No,Fecha "
     SelectDataGrid DGListPrest, AdoListarPrestamo, sSQL
  End If
End Sub

Private Sub DGListPrest_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyInsert Then
     DGListPrest.Col = 3: CuentaBanco = DGListPrest.Text
     DGListPrest.Col = 4: Cuenta = DGListPrest.Text
     DGListPrest.Col = 5: NoDias = DGListPrest.Text
     DGListPrest.Col = 7: FechaTexto = DGListPrest.Text
     Mensajes = "Fecha de Pago: " & FechaTexto _
              & ", Cuenta No.  " & Cuenta & Chr(13) _
              & "Credito No. " & CuentaBanco _
              & ",  Pago No. " & NoDias
     MsgBox Mensajes
     sSQL = "UPDATE Prestamos " _
          & "SET T = 'P' " _
          & "WHERE Cuenta_No = '" & Cuenta & "' " _
          & "AND Credito_No = '" & CuentaBanco & "' " _
          & "AND Mes_No = " & NoDias & " "
     ConectarAdoExecute sSQL
     DGResumen.Col = 2: CuentaBanco = DGResumen.Text
     DGResumen.Col = 3: Cuenta = DGResumen.Text
     sSQL = "SELECT * " _
          & "FROM Prestamos " _
          & "WHERE Cuenta_No = '" & Cuenta & "' " _
          & "AND Credito_No = '" & CuentaBanco & "' " _
          & "AND Mes_No <> 0 " _
          & "ORDER BY Credito_No,TP,Mes_No,Fecha "
     SelectDataGrid DGListPrest, AdoListarPrestamo, sSQL
  End If
End Sub

Private Sub DGTabla_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF1 Then GenerarDataTexto FPrestamo, AdoTabla, True
End Sub

Private Sub DLGarantes_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DLGarantes_LostFocus()
  TxtNombresC.Text = ""
  TxtLugarTrabC.Text = ""
  TxtTelefonoC.Text = ""
  TxtDirC.Text = ""
  sSQL = "SELECT * " _
       & "FROM Clientes_Datos_Extras " _
       & "WHERE Beneficiario = '" & DLGarantes.Text & "' " _
       & "AND Tipo_Dato = 'GARANTES' "
  SelectAdodc AdoAux, sSQL
  With AdoAux.Recordset
   If .RecordCount > 0 Then
       TxtNombresC.Text = .Fields("Beneficiario")
       MBoxCIC.Text = .Fields("CI")
       TxtLugarTrabC.Text = .Fields("Lugar_Trabajo")
       TxtTelefonoC.Text = .Fields("Telefono")
       TxtDirC.Text = .Fields("Direccion")
       DLGarantes.Visible = False
       TxtNombresC.SetFocus
   Else
       DLGarantes.Visible = False
       TxtNombresC.SetFocus
   End If
  End With
End Sub

Private Sub Form_Activate()
   If Supervisor = False Then
     If CNivel(6) Then
        Command1.Enabled = False
        Command2.Enabled = False
     End If
   End If
   Numero = ReadSetDataNum("Liquidacion", True, False)
   Contrato_No = NumEmpresa & Format(Numero, "0000000")
   LabelCredNo.Caption = Contrato_No
   sSQL = "SELECT CTP & '  ' & Descripcion As TipoP,* " _
        & "FROM Catalogo_Prestamo " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND TC <> " & Val(adFalse) & " " _
        & "AND LEN(Cta_P_1_30) > 1 " _
        & "ORDER BY CTP DESC "
   SelectDBCombo DCTipoPrestamo, AdoTipoPrestamo, sSQL, "TipoP", False
   
   sSQL = "SELECT C.Cliente,L.Cuenta_No " _
        & "FROM Clientes_Datos_Extras As L,Clientes As C " _
        & "WHERE L.Item = '" & NumEmpresa & "' " _
        & "AND L.Tipo_Dato = 'LIBRETAS' " _
        & "AND C.Codigo = L.Codigo " _
        & "ORDER BY C.Cliente,L.Cuenta_No "
   SelectDataGrid DGCtaAhorro, AdoCtaAhorro, sSQL
   RatonNormal
End Sub

Private Sub Form_Load()
   CentrarForm FPrestamo
   ConectarAdodc AdoAux
   ConectarAdodc AdoCta
   ConectarAdodc AdoCtaNo
   ConectarAdodc AdoTabla
   ConectarAdodc AdoTTabla
   ConectarAdodc AdoGarantes
   ConectarAdodc AdoLGarantes
   ConectarAdodc AdoResumenP
   ConectarAdodc AdoPrestamo
   ConectarAdodc AdoListarPrestamo
   ConectarAdodc AdoTipoPrestamo
   ConectarAdodc AdoCtaAhorro
End Sub

Private Sub MBoxCuenta_GotFocus()
  MarcarTexto MBoxCuenta
End Sub

Private Sub MBoxCuenta_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBoxCuenta_LostFocus()
   If MBoxCuenta.Text = "000000000-0" Then
      MBoxCuenta.Text = "123456789-1"
   End If
   ListarCuenta MBoxCuenta
End Sub

Private Sub MBoxFecha_LostFocus()
  FechaValida MBoxFecha
End Sub

Private Sub OpcC_Click()
  Command1.Caption = "G&rabar Conyugue"
End Sub

Private Sub OpcG_Click()
  Command1.Caption = "G&rabar Garante"
End Sub

Private Sub TextInt_GotFocus()
  MarcarTexto TextInt
End Sub

Private Sub TextInt_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextMeses_GotFocus()
  MarcarTexto TextMeses
End Sub

Private Sub TextMeses_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextMeses_LostFocus()
  If Si_No And Val(TextMeses) > 90 Then
     MsgBox "NO SE PUEDE DAR UN CREDITOS EN ESTAS CONDICIONES"
     TextMeses = "0"
     DCTipoPrestamo.SetFocus
  End If
End Sub

Private Sub TextMonto_GotFocus()
  MarcarTexto TextMonto
End Sub

Private Sub TextMonto_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextMonto_LostFocus()
  'GenerarTablaPrestamo MBoxFecha, AdoTabla, DGTabla, TextInt, TextMeses, TextMonto, Si_No, SinEspaciosIzq(DCTipoPrestamo.Text), CalcComision
  Generar_Tabla_Prestamo_Sobre_Saldos MBoxFecha, AdoTabla, DGTabla, TextInt, TextMeses, TextMonto, Si_No, SinEspaciosIzq(DCTipoPrestamo), CalcComision
  sSQL = "SELECT TP,SUM(Capital) As Tot_Capital,SUM(Interes) As Tot_Interes,SUM(Comision) As Tot_Comision,SUM(Pagos) As Tot_Pagos " _
      & "FROM Asiento_P " _
      & "WHERE Item = '" & NumEmpresa & "' " _
      & "AND CodigoU = '" & CodigoUsuario & "' " _
      & "GROUP BY TP "
  SelectDataGrid DGTTabla, AdoTTabla, sSQL
  Lista_De_Garantes
End Sub

Private Sub TxtNombresC_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  PresionoEnter KeyCode
  If KeyCode = vbKeyB And CtrlDown Then
     DLGarantes.Visible = True
     DLGarantes.SetFocus
  End If
End Sub

Private Sub TxtNombresC_LostFocus()
  TextoValido TxtNombresC, , True
End Sub
