VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Begin VB.Form SaldoCtasEspeciales 
   Caption         =   "SALDO DE CUENTAS ESPECIALES Y FLUJO DE CAJA SEMANAL"
   ClientHeight    =   8385
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11280
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8385
   ScaleWidth      =   11280
   WindowState     =   1  'Minimized
   Begin TabDlg.SSTab SSTab1 
      Height          =   6315
      Left            =   105
      TabIndex        =   12
      Top             =   1155
      Width           =   11250
      _ExtentX        =   19844
      _ExtentY        =   11139
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "&RESUMEN DE GASTOS DE CAJA"
      TabPicture(0)   =   "SaSubCta.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "LabelConcepto"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label6"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "LabelHaber"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "LabelDebe"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label3"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label5"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label4"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "DGAsiento"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "DGGastos"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Command9"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Command8"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Command7"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Command6"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "TextSucursal"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "TxtCheqDep"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "LstCtaCajaBancos"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).ControlCount=   17
      TabCaption(1)   =   "&REPORTE DE &FLUJO DE CAJA"
      TabPicture(1)   =   "SaSubCta.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "LabelEgr"
      Tab(1).Control(1)=   "Label10"
      Tab(1).Control(2)=   "LabelIng"
      Tab(1).Control(3)=   "Label11"
      Tab(1).Control(4)=   "AdoFlujoCaja"
      Tab(1).Control(5)=   "Command10"
      Tab(1).Control(6)=   "Command5"
      Tab(1).Control(7)=   "DGFlujoCaja"
      Tab(1).ControlCount=   8
      Begin VB.ListBox LstCtaCajaBancos 
         Height          =   735
         Left            =   105
         Style           =   1  'Checkbox
         TabIndex        =   14
         Top             =   735
         Width           =   6105
      End
      Begin VB.TextBox TxtCheqDep 
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
         Left            =   6300
         MaxLength       =   16
         TabIndex        =   23
         Top             =   1050
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox TextSucursal 
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
         Left            =   7455
         MaxLength       =   3
         TabIndex        =   20
         Text            =   "000"
         Top             =   420
         Width           =   540
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Imprimir Flujo de Caja"
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
         Left            =   9660
         Picture         =   "SaSubCta.frx":0038
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   420
         Width           =   1485
      End
      Begin MSDataGridLib.DataGrid DGFlujoCaja 
         Bindings        =   "SaSubCta.frx":0902
         Height          =   5475
         Left            =   -74895
         TabIndex        =   31
         Top             =   420
         Width           =   9885
         _ExtentX        =   17436
         _ExtentY        =   9657
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
      Begin VB.CommandButton Command5 
         Caption         =   "Listar Flujo de Caja"
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
         Left            =   -64920
         Picture         =   "SaSubCta.frx":091D
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   420
         Width           =   1065
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Imprimir Flujo Caja"
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
         Left            =   -64920
         Picture         =   "SaSubCta.frx":0D5F
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   1470
         Width           =   1065
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Procesar Flujo de Caja"
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
         Left            =   8085
         Picture         =   "SaSubCta.frx":1629
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   420
         Width           =   1485
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Imprimir"
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
         Picture         =   "SaSubCta.frx":1A6B
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   4620
         Width           =   855
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Grabar"
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
         Picture         =   "SaSubCta.frx":22ED
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   3780
         Width           =   855
      End
      Begin MSDataGridLib.DataGrid DGGastos 
         Bindings        =   "SaSubCta.frx":272F
         Height          =   1680
         Left            =   105
         TabIndex        =   18
         Top             =   1995
         Width           =   11040
         _ExtentX        =   19473
         _ExtentY        =   2963
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
      Begin MSDataGridLib.DataGrid DGAsiento 
         Bindings        =   "SaSubCta.frx":2747
         Height          =   2010
         Left            =   105
         TabIndex        =   25
         Top             =   3690
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   3545
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
      Begin MSAdodcLib.Adodc AdoFlujoCaja 
         Height          =   330
         Left            =   -74895
         Top             =   5895
         Width           =   4110
         _ExtentX        =   7250
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
         Caption         =   "FlujoCaja"
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
      Begin VB.Label Label4 
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Cheque/Deposito"
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
         Left            =   6300
         TabIndex        =   22
         Top             =   735
         Width           =   1695
      End
      Begin VB.Label Label5 
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Sucursal "
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
         Left            =   6300
         TabIndex        =   19
         Top             =   420
         Width           =   1170
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CUENTA DE CAJA"
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
         Top             =   420
         Width           =   6105
      End
      Begin VB.Label Label11 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " INGRESOS"
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
         Left            =   -70800
         TabIndex        =   37
         Top             =   5895
         Width           =   1170
      End
      Begin VB.Label LabelIng 
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
         ForeColor       =   &H000000FF&
         Height          =   330
         Left            =   -69645
         TabIndex        =   36
         Top             =   5895
         Width           =   1800
      End
      Begin VB.Label Label10 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " EGRESOS"
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
         Left            =   -67860
         TabIndex        =   35
         Top             =   5895
         Width           =   1065
      End
      Begin VB.Label LabelEgr 
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
         ForeColor       =   &H000000FF&
         Height          =   330
         Left            =   -66810
         TabIndex        =   34
         Top             =   5895
         Width           =   1800
      End
      Begin VB.Label Label1 
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
         Left            =   4830
         TabIndex        =   30
         Top             =   5790
         Width           =   1170
      End
      Begin VB.Label LabelDebe 
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
         ForeColor       =   &H000000FF&
         Height          =   330
         Left            =   5985
         TabIndex        =   29
         Top             =   5790
         Width           =   2115
      End
      Begin VB.Label LabelHaber 
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
         ForeColor       =   &H000000FF&
         Height          =   330
         Left            =   8085
         TabIndex        =   28
         Top             =   5790
         Width           =   2115
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Concepto"
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
         TabIndex        =   27
         Top             =   1575
         Width           =   960
      End
      Begin VB.Label LabelConcepto 
         BackColor       =   &H00FFFFFF&
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
         Left            =   1050
         TabIndex        =   26
         Top             =   1575
         Width           =   10095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Procesar "
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
      Left            =   3885
      TabIndex        =   6
      Top             =   0
      Width           =   7470
      Begin MSDataListLib.DataCombo DCCtas 
         Bindings        =   "SaSubCta.frx":2760
         DataSource      =   "AdoSubCta"
         Height          =   345
         Left            =   1680
         TabIndex        =   11
         Top             =   630
         Visible         =   0   'False
         Width           =   5685
         _ExtentX        =   10028
         _ExtentY        =   609
         _Version        =   393216
         Text            =   "DataCombo1"
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
      Begin VB.CheckBox CheqIndiv 
         Caption         =   "SubModulo:"
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
         Left            =   210
         TabIndex        =   10
         Top             =   630
         Width           =   1380
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Ingresos y Egresos de &Caja"
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
         Left            =   4410
         TabIndex        =   9
         Top             =   210
         Value           =   -1  'True
         Width           =   2745
      End
      Begin VB.OptionButton OpcE 
         Caption         =   "&Egresos de Caja"
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
         Left            =   2205
         TabIndex        =   8
         Top             =   210
         Width           =   1800
      End
      Begin VB.OptionButton OpcI 
         Caption         =   "&Ingresos de Caja"
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
         Left            =   210
         TabIndex        =   7
         Top             =   210
         Width           =   1800
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Salir del modulo"
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
      Left            =   11445
      Picture         =   "SaSubCta.frx":2778
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   105
      Width           =   1590
   End
   Begin MSMask.MaskEdBox MBoxFechaF 
      Height          =   330
      Left            =   945
      TabIndex        =   3
      Top             =   735
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   582
      _Version        =   393216
      AllowPrompt     =   -1  'True
      AutoTab         =   -1  'True
      MaxLength       =   10
      OLEDragMode     =   1
      OLEDropMode     =   2
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
      OLEDragMode     =   1
      OLEDropMode     =   2
   End
   Begin MSMask.MaskEdBox MBoxFechaI 
      Height          =   330
      Left            =   945
      TabIndex        =   1
      Top             =   105
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   582
      _Version        =   393216
      AllowPrompt     =   -1  'True
      AutoTab         =   -1  'True
      MaxLength       =   10
      OLEDragMode     =   1
      OLEDropMode     =   2
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
      OLEDragMode     =   1
      OLEDropMode     =   2
   End
   Begin MSAdodcLib.Adodc AdoCtas 
      Height          =   330
      Left            =   210
      Top             =   1890
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
      Caption         =   "Ctas"
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
   Begin MSAdodcLib.Adodc AdoGastos 
      Height          =   330
      Left            =   210
      Top             =   2205
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
      Caption         =   "Gastos"
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
   Begin MSAdodcLib.Adodc AdoAsiento 
      Height          =   330
      Left            =   210
      Top             =   2520
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
   Begin MSAdodcLib.Adodc AdoCaja 
      Height          =   330
      Left            =   210
      Top             =   2835
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
   Begin MSAdodcLib.Adodc AdoSubCta 
      Height          =   330
      Left            =   210
      Top             =   3150
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
      Caption         =   "SubCta"
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
   Begin VB.Label LabelDiaH 
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
      Height          =   330
      Left            =   2310
      TabIndex        =   5
      Top             =   735
      Width           =   1485
   End
   Begin VB.Label LabelDiaD 
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
      Height          =   330
      Left            =   2310
      TabIndex        =   4
      Top             =   105
      Width           =   1485
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Desde"
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
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Hasta"
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
      TabIndex        =   2
      Top             =   735
      Width           =   855
   End
End
Attribute VB_Name = "SaldoCtasEspeciales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Lista_Ctas As String

Private Sub CheqIndiv_Click()
  If CheqIndiv.value = 1 Then DCCtas.Visible = True Else DCCtas.Visible = False
End Sub

Private Sub Command1_Click()
  Unload SaldoCtasEspeciales
End Sub

Private Sub Command10_Click()
  If OpcI.value Then
     SQLMsg1 = "INGRESOS DE CAJA CHICA"
     Documento = 1
  ElseIf OpcE.value Then
     SQLMsg1 = "EGRESOS DE CAJA CHICA"
     Documento = 2
  Else
     SQLMsg1 = "INGRESOS/EGRESOS DE CAJAS"
     Documento = 3
  End If
  SQLMsg1 = SQLMsg1 & " Desde " & MBoxFechaI & " al " & MBoxFechaF
  ImprimirLibroReciboCaja AdoFlujoCaja, 9, Documento
End Sub

Private Sub Command5_Click()
  FechaValida MBoxFechaI
  FechaValida MBoxFechaF
  FechaIni = BuscarFecha(MBoxFechaI)
  FechaFin = BuscarFecha(MBoxFechaF)
  I = CFechaLong(MBoxFechaI)
  J = CFechaLong(MBoxFechaF)
     SubCtaGen = Ninguno
     If AdoSubCta.Recordset.RecordCount > 0 Then
        AdoSubCta.Recordset.MoveFirst
        AdoSubCta.Recordset.Find ("Cliente = '" & DCCtas.Text & "' ")
        If Not AdoSubCta.Recordset.EOF Then
           SubCtaGen = AdoSubCta.Recordset.Fields("Codigo")
        End If
     End If
    'Mayorizar Cajas Indivuduales
     Saldo = 0: Contador = 1
     sSQL = "SELECT Fecha,Cta,Codigo,Numero,Ingreso,Egreso,Saldo " _
          & "FROM Trans_Gastos_Caja " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & Lista_Ctas
     If CheqIndiv.value = 1 Then sSQL = sSQL & "AND Codigo = '" & SubCtaGen & "' "
     sSQL = sSQL & "ORDER BY Fecha,Cta,Codigo,Numero "
     Select_Adodc AdoCtas, sSQL
     RatonReloj
     With AdoCtas.Recordset
      If .RecordCount > 0 Then
          RatonReloj
          Codigo = .Fields("Codigo")
          Saldo = 0: Contador = 1
          Do While Not .EOF
'''             If Codigo <> .Fields("Codigo") Then
'''                Codigo = .Fields("Codigo")
'''                Saldo = 0
'''             End If
             Saldo = Saldo + .Fields("Ingreso") - .Fields("Egreso")
             If .Fields("Saldo") <> Saldo Then
                .Fields("Saldo") = Saldo
                .Update
             End If
             Contador = Contador + 1
            .MoveNext
          Loop
      End If
     End With
     RatonNormal
  Debitos = 0: Creditos = 0
  sSQL = "SELECT Fecha,Beneficiario,Detalle,Concepto,Numero,Ingreso,Egreso,Saldo,IVA,CodRet,Serie,Secuencial,Autorizacion " _
       & "FROM Trans_Gastos_Caja As TGC, Catalogo_SubCtas As GC " _
       & "WHERE TGC.Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
       & "AND TGC.Item = '" & NumEmpresa & "' " _
       & "AND TGC.Periodo = '" & Periodo_Contable & "' " _
       & Lista_Ctas
  If CheqIndiv.value = 1 Then sSQL = sSQL & "AND TGC.Codigo = '" & SubCtaGen & "' "
  sSQL = sSQL & "AND TGC.Item = GC.Item " _
       & "AND TGC.Periodo = GC.Periodo " _
       & "AND TGC.Codigo = GC.Codigo " _
       & "AND TGC.TC = 'GC' " _
       & "ORDER BY TGC.Fecha,TGC.Cta,TGC.Codigo,TGC.Numero "
  Select_Adodc_Grid DGFlujoCaja, AdoFlujoCaja, sSQL
 'MsgBox sSQL
  With AdoFlujoCaja.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          SumaDebe = SumaDebe + .Fields("Ingreso")
          SumaHaber = SumaHaber + .Fields("Egreso")
         .MoveNext
       Loop
      .MoveFirst
   End If
  End With
  LabelIng.Caption = Format(SumaDebe, "#,##0.00")
  LabelEgr.Caption = Format(SumaHaber, "#,##0.00")
  RatonNormal
End Sub

Private Sub Command6_Click()
  DGGastos.Visible = False
  SQLMsg3 = ""
  If OpcI.value Then
     SQLMsg1 = "INGRESO DE CAJA"
  ElseIf OpcE.value Then
     SQLMsg1 = "GASTOS DE CAJA"
  End If
  SQLMsg2 = "Desde:  " & MBoxFechaI & "   al   " & MBoxFechaF
  ImprimirAdodc AdoGastos, 1, 8, True
  'ImprimirGastosCaja AdoGastos, 2
  DGGastos.Visible = True
End Sub

Private Sub Command7_Click()
Dim Contra_Ctas() As String
Dim TotalCaja() As Currency

 'Reportes de Gastos de Caja Chica
  ReDim Contra_Ctas(1) As String
  ReDim TotalCaja(1) As Currency
  Contra_Ctas(0) = ""
  TotalCaja(0) = 0
  Lista_Ctas = ""
  If AdoCaja.Recordset.RecordCount > 0 Then
     For I = 0 To LstCtaCajaBancos.ListCount - 1
         If LstCtaCajaBancos.Selected(I) Then
            AdoCaja.Recordset.MoveFirst
            AdoCaja.Recordset.Find ("Cuenta = '" & LstCtaCajaBancos.List(I) & "'")
            If Not AdoCaja.Recordset.EOF Then
               Lista_Ctas = Lista_Ctas & "'" & AdoCaja.Recordset.Fields("Contra_Cta") & "',"
            End If
         End If
     Next I
  End If
  If Len(Lista_Ctas) > 1 Then Lista_Ctas = "AND Contra_Cta IN (" & MidStrg(Lista_Ctas, 1, Len(Lista_Ctas) - 1) & ") "
  RatonReloj
  FechaValida MBoxFechaI
  FechaValida MBoxFechaF
  FechaIni = BuscarFecha(MBoxFechaI)
  FechaFin = BuscarFecha(MBoxFechaF)
  NumPagos = ReadSetDataNum("Caja Chica", False, False)
  IniciarAsientosDe DGAsiento, AdoAsiento
  If OpcI.value Then LabelConcepto.Caption = " Reposicion de Caja Chica desde el " & MBoxFechaI & " al " & MBoxFechaF
  If OpcE.value Then LabelConcepto.Caption = " Egresos de Caja Chica desde el " & MBoxFechaI & " al " & MBoxFechaF
  LabelDiaD.Caption = DiasLetras(Weekday(MBoxFechaI))
  LabelDiaH.Caption = DiasLetras(Weekday(MBoxFechaF))
  DGGastos.Visible = False
    SubCtaGen = Ninguno
    If AdoSubCta.Recordset.RecordCount > 0 Then
       AdoSubCta.Recordset.MoveFirst
       AdoSubCta.Recordset.Find ("Cliente = '" & DCCtas.Text & "' ")
       If Not AdoSubCta.Recordset.EOF Then
          SubCtaGen = AdoSubCta.Recordset.Fields("Codigo")
       End If
    End If
    
  sSQL = "DELETE * " _
       & "FROM Saldo_Diarios " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND TP = 'FCJ' "
  Ejecutar_SQL_SP sSQL
  
  SQL1 = "DELETE * " _
       & "FROM Asiento_SC " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND T_No = " & Trans_No & " "
  Ejecutar_SQL_SP SQL1
  
  Total_IVA = 0
  
  sSQL = "SELECT Contra_Cta,SUM(Egreso+IVA) As TEgreso " _
       & "FROM Trans_Gastos_Caja " _
       & "WHERE Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & Lista_Ctas _
       & "AND Egreso > 0 "
  If CheqIndiv.value = 1 Then sSQL = sSQL & "AND Codigo = '" & SubCtaGen & "' "
  sSQL = sSQL _
       & "GROUP BY Contra_Cta " _
       & "ORDER BY Contra_Cta "
  Select_Adodc AdoCtas, sSQL
  With AdoCtas.Recordset
   If .RecordCount > 0 Then
       ReDim Contra_Ctas(.RecordCount) As String
       ReDim TotalCaja(.RecordCount) As Currency
       
       IE = 0
       Do While Not .EOF
          Contra_Ctas(IE) = .Fields("Contra_Cta")
          TotalCaja(IE) = .Fields("TEgreso")
          IE = IE + 1
         .MoveNext
       Loop
   End If
  End With
  sSQL = "SELECT Cta,Contra_Cta,Codigo,Fecha,Ingreso,Egreso,IVA " _
       & "FROM Trans_Gastos_Caja " _
       & "WHERE Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & Lista_Ctas
  If OpcI.value Then
     sSQL = sSQL & "AND Ingreso > 0 "
  ElseIf OpcE.value Then
     sSQL = sSQL & "AND Egreso > 0 "
  End If
  If CheqIndiv.value = 1 Then sSQL = sSQL & "AND Codigo = '" & SubCtaGen & "' "
  sSQL = sSQL & "ORDER BY Cta,Codigo,Fecha "
  Select_Adodc AdoCtas, sSQL
  With AdoCtas.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       Codigo = .Fields("Codigo")
       Cta = .Fields("Cta")
       Cta1 = .Fields("Contra_Cta")
       Saldo = 0: Saldo_ME = 0
       Total_ME = 0
       Debe = 0: Haber = 0
       For I = 1 To 7
           TotalDia(I) = 0
       Next I
       Do While Not .EOF
          If Codigo <> .Fields("Codigo") Or Cta <> .Fields("Cta") Then
             Insertar_Totales_Gastos
             Codigo = .Fields("Codigo")
             Cta = .Fields("Cta")
             For I = 1 To 7
                 TotalDia(I) = 0
             Next I
          End If
          I = Weekday(.Fields("Fecha"))
          If OpcI.value Then
             TotalDia(I) = TotalDia(I) + .Fields("Ingreso")
          ElseIf OpcE.value Then
             TotalDia(I) = TotalDia(I) + .Fields("Egreso")
          End If
          Total_IVA = Total_IVA + .Fields("IVA")
         .MoveNext
       Loop
       Insertar_Totales_Gastos
   End If
  End With
 'Asiento Contable
  'If SQL_Server = False Then MsgBox "Verifique los Datos"
  NoCheque = Ninguno
  sSQL = "SELECT Cta,CodigoC,Total " _
       & "FROM Saldo_Diarios " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND TP = 'FCJ' " _
       & "ORDER BY Cta,CodigoC "
  Select_Adodc AdoCtas, sSQL
  With AdoCtas.Recordset
   If .RecordCount > 0 Then
       Cta = .Fields("Cta")
       Codigo = .Fields("CodigoC")
       Total = 0: Saldo = 0
       Do While Not .EOF
          If Cta <> .Fields("Cta") Then
             If OpcI.value Then
                InsertarAsientos AdoAsiento, Cta, 0, 0, Total
             ElseIf OpcE.value Then
                InsertarAsientos AdoAsiento, Cta, 0, Total, 0
             End If
             Cta = .Fields("Cta")
             Total = 0
          End If
          Total = Total + .Fields("Total")
          Saldo = Saldo + .Fields("Total")
         .MoveNext
       Loop
       If OpcI.value Then
          InsertarAsientos AdoAsiento, Cta, 0, 0, Total
       ElseIf OpcE.value Then
          InsertarAsientos AdoAsiento, Cta, 0, Total, 0
       End If
   End If
  End With
  If Total_IVA > 0 Then InsertarAsientos AdoAsiento, Cta_IVA_Inventario, 0, Total_IVA, 0
  If OpcI.value Then
'     Contra_Cta = SinEspaciosIzq(DLCtas.Text)
 '    InsertarAsientos AdoAsiento, Contra_Cta, 0, Saldo, 0
  ElseIf OpcE.value Then
     If TxtCheqDep.Visible Then NoCheque = UCaseStrg(TxtCheqDep)
        For IE = 0 To UBound(Contra_Ctas)
            InsertarAsientos AdoAsiento, Contra_Ctas(IE), 0, 0, TotalCaja(IE)
        Next IE
  End If
  sSQL = "SELECT GC.TC,GC.Detalle,"
  For NoDias = 1 To 7
      sSQL = sSQL & DiasLetras(NoDias) & ","
  Next NoDias
  sSQL = sSQL & "TGC.Total " _
       & "FROM Saldo_Diarios As TGC,Catalogo_SubCtas As GC " _
       & "WHERE TGC.Item = '" & NumEmpresa & "' " _
       & "AND TGC.CodigoU = '" & CodigoUsuario & "' " _
       & "AND GC.Periodo = '" & Periodo_Contable & "' " _
       & "AND GC.Codigo = TGC.CodigoC " _
       & "AND GC.Item = TGC.Item " _
       & "ORDER BY GC.Codigo "
  Select_Adodc_Grid DGGastos, AdoGastos, sSQL
  If SQL_Server = False Then MsgBox "Proceso Terminado, Verifique."
  
  sSQL = "SELECT * " _
       & "FROM Asiento " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND T_No = " & Trans_No & " " _
       & "ORDER BY DEBE DESC,HABER,CODIGO "
  Select_Adodc_Grid DGAsiento, AdoAsiento, sSQL
  SumaDebe = 0: SumaHaber = 0
  With AdoAsiento.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       NumTrans = 0
       Do While Not .EOF
          SumaDebe = SumaDebe + .Fields("DEBE")
          SumaHaber = SumaHaber + .Fields("HABER")
         .Fields("A_No") = NumTrans
          NumTrans = NumTrans + 1
         .MoveNext
       Loop
      .UpdateBatch
      .MoveFirst
   End If
  End With
  LabelDebe.Caption = Format(SumaDebe, "#,##0.00")
  LabelHaber.Caption = Format(SumaHaber, "#,##0.00")
  DGGastos.Visible = True
  RatonNormal
End Sub

Private Sub Command8_Click()
  SQLMsg3 = LabelConcepto.Caption
  If OpcI.value Then
     SQLMsg1 = "INGRESO DE CAJA"
  ElseIf OpcE.value Then
     SQLMsg1 = "GASTOS DE CAJA"
  End If
  SQLMsg2 = "Desde:  " & MBoxFechaI & "   al   " & MBoxFechaF
  ImprimirAdodc AdoAsiento, 1, 8, True

'''  If OpcI.Value Then
'''     SQLMsg1 = "INGRESOS DE CAJA"
'''     Documento = 1
'''     SQLMsg1 = SQLMsg1 & " Desde " & MBoxFechaI.Text & " al " & MBoxFechaF.Text
'''     ImprimirLibroReciboCaja AdoFlujoCaja, 8, Documento
'''  ElseIf OpcE.Value Then
'''     SQLMsg1 = "EGRESOS DE CAJA"
'''     Documento = 2
'''     SQLMsg1 = SQLMsg1 & " Desde " & MBoxFechaI.Text & " al " & MBoxFechaF.Text
'''     ImprimirLibroReciboCaja AdoFlujoCaja, 8, Documento
'''  End If
End Sub

Private Sub Command9_Click()
If ClaveSupervisor Then
   SumaDebe = 0: SumaHaber = 0
   With AdoAsiento.Recordset
    If .RecordCount > 0 Then
       .MoveFirst
        Do While Not .EOF
           SumaDebe = SumaDebe + .Fields("DEBE")
           SumaHaber = SumaHaber + .Fields("HABER")
          .MoveNext
        Loop
       .MoveFirst
    End If
   End With
   Saldo = Redondear(SumaDebe - SumaHaber, 2)
   LabelDebe.Caption = Format(SumaDebe, "#,##0.00")
   LabelHaber.Caption = Format(SumaHaber, "#,##0.00")
  'MsgBox Saldo
   If Saldo = 0 Then
      FechaTexto = FechaSistema
      If OpcI.value Then
         FechaTexto = MBoxFechaI
      ElseIf OpcE.value Then
         FechaTexto = MBoxFechaF
      End If
      FechaComp = FechaTexto
      NumComp = ReadSetDataNum("Diario", True, True)
      SQL1 = "DELETE * " _
           & "FROM Asiento_SC " _
           & "WHERE Item = '" & NumEmpresa & "' " _
           & "AND CodigoU = '" & CodigoUsuario & "' " _
           & "AND Codigo = '" & Ninguno & "' " _
           & "AND T_No = " & Trans_No & " "
      Ejecutar_SQL_SP SQL1

      Co.T = Normal
      Co.TP = CompDiario
      Co.Fecha = FechaTexto
      Co.Numero = NumComp
      Co.Concepto = LabelConcepto.Caption
      Co.CodigoB = Ninguno
      Co.RUC_CI = LimpiarRUC
      Co.Efectivo = 0
      Co.Monto_Total = SumaDebe
      Co.T_No = Trans_No
      Co.Usuario = CodigoUsuario
      Co.Item = NumEmpresa
      Grabar_Comprobante Co
      If OpcE.value Then
         sSQL = "SELECT Cta,Contra_Cta,CodRet,Serie,Autorizacion,Secuencial,Codigo,CodigoC,Fecha,Egreso,IVA " _
              & "FROM Trans_Gastos_Caja " _
              & "WHERE Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
              & "AND Item = '" & NumEmpresa & "' " _
              & "AND Periodo = '" & Periodo_Contable & "' " _
              & Lista_Ctas _
              & "AND Egreso > 0 " _
              & "AND CodRet <> '000' " _
              & "ORDER BY Cta,Codigo,Fecha "
         Select_Adodc AdoCtas, sSQL
         With AdoCtas.Recordset
          If .RecordCount > 0 Then
              NumTrans = .RecordCount
              Do While Not .EOF
                 TipoDoc = "01"
                 Total_Ret = 0
                 Total_Con_IVA = 0
                 Total_Sin_IVA = 0
                 Total_Sin_No_IVA = 0
                 Total_IVA = .Fields("IVA")
                 If Total_IVA > 0 Then
                    Total_Con_IVA = .Fields("Egreso")
                 Else
                    Total_Sin_IVA = .Fields("Egreso")
                 End If
                 SubTotal = Total_Con_IVA + Total_Sin_IVA
                 
                 Mifecha = .Fields("Fecha")
                 CodigoCli = .Fields("CodigoC")
                 CodigoP = .Fields("CodRet")
                 SerieFactura = .Fields("Serie")
                 Autorizacion = .Fields("Autorizacion")
                 Factura_No = .Fields("Secuencial")
                 Cta_CajaG = .Fields("Contra_Cta")
                 Cta_Gasto = .Fields("Cta")
                 
                 sSQL = "DELETE * " _
                      & "FROM Trans_Compras " _
                      & "WHERE Item = '" & NumEmpresa & "' " _
                      & "AND Periodo = '" & Periodo_Contable & "' " _
                      & "AND FechaEmision = #" & BuscarFecha(Mifecha) & "# " _
                      & "AND IdProv = '" & CodigoCli & "' " _
                      & "AND Establecimiento = '" & MidStrg(SerieFactura, 1, 3) & "' " _
                      & "AND PuntoEmision = '" & MidStrg(SerieFactura, 4, 3) & "' " _
                      & "AND Secuencial = " & Factura_No & " " _
                      & "AND TipoComprobante = " & Val(TipoDoc) & " " _
                      & "AND Autorizacion = '" & Autorizacion & "' "
                 Ejecutar_SQL_SP sSQL
                 sSQL = "DELETE * " _
                      & "FROM Trans_Air " _
                      & "WHERE Item = '" & NumEmpresa & "' " _
                      & "AND Periodo = '" & Periodo_Contable & "' " _
                      & "AND IdProv = '" & CodigoCli & "' " _
                      & "AND EstabRetencion = '000' " _
                      & "AND PtoEmiRetencion = '000' " _
                      & "AND SecRetencion = 0 " _
                      & "AND AutRetencion = '0' " _
                      & "AND EstabFactura = '" & MidStrg(SerieFactura, 1, 3) & "' " _
                      & "AND PuntoEmiFactura = '" & MidStrg(SerieFactura, 4, 3) & "' " _
                      & "AND Factura_No = " & Factura_No & " "
                 Ejecutar_SQL_SP sSQL
                'Empezamos a grabar los datos de la retencion
                 SetAdoAddNew "Trans_Compras"
                 SetAdoFields "IdProv", CodigoCli
                 SetAdoFields "DevIva", "N"
                 SetAdoFields "CodSustento", "01"
                 SetAdoFields "TipoComprobante", Val(TipoDoc)
                 SetAdoFields "Establecimiento", MidStrg(SerieFactura, 1, 3)
                 SetAdoFields "PuntoEmision", MidStrg(SerieFactura, 4, 3)
                 SetAdoFields "Secuencial", Factura_No
                 SetAdoFields "Autorizacion", Autorizacion
                 SetAdoFields "FechaEmision", Mifecha
                 SetAdoFields "FechaRegistro", Mifecha
                 SetAdoFields "FechaCaducidad", Mifecha
                 If Total_IVA > 0 Then
                    SetAdoFields "BaseImpGrav", Total_Con_IVA
                    SetAdoFields "MontoIva", Total_IVA
                    SetAdoFields "PorcentajeIva", 2
                 End If
                 SetAdoFields "BaseImponible", Total_Sin_IVA
                 SetAdoFields "BaseNoObjIVA", Total_Sin_No_IVA
                 SetAdoFields "Cta_Pago", Cta_CajaG
                 SetAdoFields "Cta_Gasto", Cta_Gasto
                 SetAdoFields "PagoLocExt", "01"
                 SetAdoFields "FormaPago", "01"
                 SetAdoFields "PaisEfecPago", "NA"
                 SetAdoFields "AplicConvDobTrib", "NA"
                 SetAdoFields "PagExtSujRetNorLeg", "NA"
                 SetAdoFields "Linea_SRI", 0
                 SetAdoFields "FechaEmiModificado", "000"
                 SetAdoFields "EstabModificado", "000"
                 SetAdoFields "PtoEmiModificado", "000"
                 SetAdoFields "SecModificado", "000"
                 SetAdoFields "AutModificado", "000"
                 SetAdoFields "T", Normal
                 SetAdoFields "TP", CompDiario
                 SetAdoFields "Numero", NumComp
                 SetAdoFields "Fecha", Mifecha
                 SetAdoFields "ID", NumTrans
                 SetAdoFields "CodigoU", CodigoUsuario
                 SetAdoFields "Periodo", Periodo_Contable
                 SetAdoFields "Item", NumEmpresa
                 SetAdoUpdate
                 'MsgBox Total_Sin_IVA & vbCrLf & Total_Con_IVA & vbCrLf & NombreCliente
                 
                'RETENCION EN LA FUENTE
                 SetAdoAddNew "Trans_Air"
                 SetAdoFields "CodRet", CodigoP
                 SetAdoFields "BaseImp", SubTotal
                 SetAdoFields "ValRet", Total_Ret
                 SetAdoFields "EstabRetencion", "000"
                 SetAdoFields "PtoEmiRetencion", "000"
                 SetAdoFields "SecRetencion", "0"
                 SetAdoFields "AutRetencion", "0"
                 SetAdoFields "Tipo_Trans", "C"
                 SetAdoFields "IdProv", CodigoCli
                 SetAdoFields "EstabFactura", MidStrg(SerieFactura, 1, 3)
                 SetAdoFields "PuntoEmiFactura", MidStrg(SerieFactura, 4, 3)
                 SetAdoFields "Factura_No", Factura_No
                 SetAdoFields "Linea_SRI", 0
                 SetAdoFields "T", Normal
                 SetAdoFields "TP", CompDiario
                 SetAdoFields "Numero", NumComp
                 SetAdoFields "Fecha", Mifecha
                 SetAdoFields "ID", NumTrans
                 SetAdoFields "CodigoU", CodigoUsuario
                 SetAdoFields "Periodo", Periodo_Contable
                 SetAdoFields "Item", NumEmpresa
                 SetAdoUpdate
                 NumTrans = NumTrans + 1
               .MoveNext
              Loop
           End If
         End With
      End If
      RatonNormal
      ImprimirComprobantesDe False, Co
      DGAsiento.Visible = False
      IniciarAsientosDe DGAsiento, AdoAsiento
      DGAsiento.Visible = True
      NumPagos = ReadSetDataNum("Caja Chica", True, True)
   Else
      MsgBox "Las transacciones no cuadran"
   End If
 End If
End Sub

Private Sub DGFlujoCaja_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF1 Then GenerarDataTexto SaldoCtasEspeciales, AdoFlujoCaja
End Sub

''Private Sub DLCtas_LostFocus()
''  TxtCheqDep.Visible = False
''  With AdoCaja.Recordset
''   If .RecordCount > 0 Then
''      .MoveFirst
''      .Find ("Codigo = '" & SinEspaciosIzq(DLCtas) & "' ")
''       If Not .EOF And .Fields("TC") = "BA" Then
''          TxtCheqDep.Visible = True
''          TxtCheqDep.SetFocus
''       End If
''   End If
''  End With
''End Sub

Private Sub Form_Activate()
  TipoDoc = CompDiario
  Trans_No = 15
  Lista_Ctas = ""
  IniciarAsientosDe DGAsiento, AdoAsiento
  TextSucursal.Text = Format(NumEmpresa, "00")
  sSQL = "SELECT C.Detalle As Cliente,TS.Codigo " _
       & "FROM Trans_Gastos_Caja As TS,Catalogo_SubCtas As C " _
       & "WHERE TS.Item = '" & NumEmpresa & "' " _
       & "AND TS.Periodo = '" & Periodo_Contable & "' " _
       & "AND TS.Codigo = C.Codigo " _
       & "AND TS.Item = C.Item " _
       & "AND TS.Periodo = C.Periodo " _
       & "GROUP BY C.Detalle,TS.Codigo " _
       & "ORDER BY C.Detalle,TS.Codigo "
  SelectDB_Combo DCCtas, AdoSubCta, sSQL, "Cliente"
  sSQL = "DELETE * " _
       & "FROM Saldo_Diarios " _
       & "WHERE CodigoU = '" & CodigoUsuario & "' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND TP = 'FCJ' "
  Ejecutar_SQL_SP sSQL
  sSQL = "SELECT GC.Contra_Cta,CC.Cuenta,CC.TC " _
       & "FROM Catalogo_Cuentas As CC,Trans_Gastos_Caja As GC " _
       & "WHERE CC.Item = '" & NumEmpresa & "' " _
       & "AND CC.Periodo = '" & Periodo_Contable & "' " _
       & "AND CC.Item = GC.Item " _
       & "AND CC.Periodo = GC.Periodo " _
       & "AND CC.Codigo = GC.Contra_Cta " _
       & "GROUP BY GC.Contra_Cta,CC.Cuenta,CC.TC " _
       & "ORDER BY CC.Cuenta "
  Select_Adodc AdoCaja, sSQL
  LstCtaCajaBancos.Clear
  LstCtaCajaBancos.AddItem "Todas"
  LstCtaCajaBancos.Text = "Todas"
  LstCtaCajaBancos.Selected(LstCtaCajaBancos.ListCount - 1) = True
  With AdoCaja.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          LstCtaCajaBancos.AddItem .Fields("Cuenta")
         .MoveNext
       Loop
   End If
  End With
  SaldoCtasEspeciales.WindowState = 2
  SaldoCtasEspeciales.Caption = "Listar Flujo de Caja"
  RatonNormal
  MBoxFechaI.SetFocus
End Sub

Private Sub Form_Load()
  SaldoCtasEspeciales.Caption = "Espere..."
 'CentrarForm SaldoCtasEspeciales
  ConectarAdodc AdoAsiento
  ConectarAdodc AdoCtas
  ConectarAdodc AdoCaja
  ConectarAdodc AdoGastos
  ConectarAdodc AdoSubCta
  ConectarAdodc AdoFlujoCaja
End Sub

Private Sub MBoxFechaF_GotFocus()
  MarcarTexto MBoxFechaF
End Sub

Private Sub MBoxFechaF_LostFocus()
  FechaValida MBoxFechaF
End Sub

Private Sub MBoxFechaI_GotFocus()
  MarcarTexto MBoxFechaI
End Sub

Private Sub MBoxFechaI_LostFocus()
  FechaValida MBoxFechaI
End Sub

Public Sub Insertar_Totales_Gastos()
  Total = 0
  For I = 1 To 7
      Total = Total + TotalDia(I)
  Next I
  Total = Redondear(Total, 2)
  If Total <> 0 Then
     SetAdoAddNew "Saldo_Diarios"
     SetAdoFields "CodigoC", Codigo
     SetAdoFields "TP", "FCJ"
     For I = 2 To 7
         SetAdoFields DiasLetras(Weekday(I)), Redondear(TotalDia(I), 2)
     Next I
     If OpcI.value Then
        SetAdoFields "TC", "C"
        SetAdoFields "Cta", Cta_CajaG
     ElseIf OpcE.value Then
        SetAdoFields "TC", "G"
        SetAdoFields "Cta", Cta
        SetAdoFields "Cta_Aux", Cta1
     End If
     SetAdoFields "Total", Total
     SetAdoFields "CodigoU", CodigoUsuario
     SetAdoFields "Item", NumEmpresa
     SetAdoUpdate
     If OpcE.value Then
        SetAdoAddNew "Asiento_SC"
        SetAdoFields "DH", "1"
        SetAdoFields "Factura", NumPagos
        SetAdoFields "TC", "G"
        SetAdoFields "Valor", Total
        SetAdoFields "Codigo", Codigo
        SetAdoFields "Cta", Cta
        SetAdoFields "T_No", Trans_No
        SetAdoFields "CodigoU", CodigoUsuario
        SetAdoFields "Item", NumEmpresa
        SetAdoUpdate
     End If
  End If
  For I = 1 To 7
      TotalDia(I) = 0
  Next I
End Sub

Private Sub TextSucursal_GotFocus()
  MarcarTexto TextSucursal
End Sub

Private Sub TextSucursal_LostFocus()
  TextoValido TextSucursal, True
  TextSucursal.Text = Format(TextSucursal.Text, "000")
End Sub

Public Sub InsertarResumen(TDetalle As String)
  SetAdoAddNew "Saldo_Diarios"
  SetAdoFields "Comprobante", TDetalle
  SetAdoFields "Saldo_Anterior", Saldo + Creditos - Debitos
  SetAdoFields "Ingresos", Debitos
  SetAdoFields "Egresos", Creditos
  SetAdoFields "Saldo_Anterior", Saldo
  SetAdoFields "Item", NumEmpresa
  SetAdoFields "CodigoU", CodigoUsuario
  SetAdoFields "TP", "RESU"
  SetAdoUpdate
End Sub

Private Sub TxtCheqDep_GotFocus()
  MarcarTexto TxtCheqDep
End Sub

Private Sub TxtCheqDep_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtCheqDep_LostFocus()
   If IsNumeric(TxtCheqDep) Then
      TxtCheqDep = Format(Val(TxtCheqDep), "00000000")
   End If
End Sub
