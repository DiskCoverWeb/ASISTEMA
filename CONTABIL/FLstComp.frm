VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Begin VB.Form FListComprobantes 
   Caption         =   "SubCtas"
   ClientHeight    =   8895
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   11280
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   15615
   ScaleWidth      =   28560
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "&Fecha    >>>"
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
      Style           =   1  'Graphical
      TabIndex        =   38
      ToolTipText     =   "Presone un Click para cambiar la fecha"
      Top             =   1365
      Width           =   1380
   End
   Begin VB.Frame Frame1 
      Height          =   645
      Left            =   105
      TabIndex        =   0
      Top             =   630
      Width           =   11460
      Begin VB.OptionButton OpcTP 
         Caption         =   "&1.- Diario"
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
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   210
         Value           =   -1  'True
         Width           =   1275
      End
      Begin VB.OptionButton OpcTP 
         Caption         =   "&2.- Ingreso"
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
         Left            =   2205
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   210
         Width           =   1275
      End
      Begin VB.OptionButton OpcTP 
         Caption         =   "&3.- Egreso"
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
         Left            =   3570
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   210
         Width           =   1275
      End
      Begin VB.OptionButton OpcTP 
         Caption         =   "&4.- N/D"
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
         Left            =   4935
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   210
         Width           =   1275
      End
      Begin VB.OptionButton OpcTP 
         Caption         =   "&5.- N/C"
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
         Index           =   4
         Left            =   6300
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   210
         Width           =   1275
      End
      Begin VB.ComboBox CMeses 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   345
         Left            =   8190
         TabIndex        =   8
         Text            =   "Combo1"
         Top             =   210
         Width           =   1590
      End
      Begin MSMask.MaskEdBox MBoxEmpNo 
         Height          =   330
         Left            =   105
         TabIndex        =   1
         ToolTipText     =   "Escriba el número de la Empresa o Sucursal"
         Top             =   210
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   582
         _Version        =   393216
         AllowPrompt     =   -1  'True
         AutoTab         =   -1  'True
         MaxLength       =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "000"
         Mask            =   "###"
         PromptChar      =   "0"
      End
      Begin MSDataListLib.DataCombo DCComp 
         Bindings        =   "FLstComp.frx":0000
         DataSource      =   "AdoCompNo"
         Height          =   345
         Left            =   9765
         TabIndex        =   9
         Top             =   210
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   609
         _Version        =   393216
         ForeColor       =   255
         Text            =   "0100000000"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " &No."
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
         Left            =   7665
         TabIndex        =   7
         Top             =   210
         Width           =   540
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&S"
      Height          =   330
      Left            =   10395
      TabIndex        =   35
      Top             =   7455
      Width           =   330
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   11550
      Top             =   1260
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FLstComp.frx":0018
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FLstComp.frx":08F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FLstComp.frx":11CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FLstComp.frx":1A5E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FLstComp.frx":1D78
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FLstComp.frx":2652
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FLstComp.frx":2F2C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FLstComp.frx":3806
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FLstComp.frx":40E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FLstComp.frx":4786
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FLstComp.frx":4AA0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   34
      Top             =   0
      Width           =   28560
      _ExtentX        =   50377
      _ExtentY        =   1164
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Salir"
            Object.ToolTipText     =   "Salir del modulo"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Imprimir_Comprobante"
            Object.ToolTipText     =   "Imprime comprobante en pantalla"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Imprimir_Retencion"
            Object.ToolTipText     =   "Imprime Comprobante de Retencion"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Imprimir_Cheque"
            Object.ToolTipText     =   "Imprime el Cheque"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Modificar_Comprobante"
            Object.ToolTipText     =   "Modificar el Comprobante"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Anular_Comprobante"
            Object.ToolTipText     =   "Anular el Comprobante"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Autorizar_Comprobante"
            Object.ToolTipText     =   "Autoriza Comprobante Procesado"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Copiar_Comprobante"
            Object.ToolTipText     =   "Realizar una copia del Comprobante"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Copiar_Item"
            Object.ToolTipText     =   "Copia a otra empresa el comprobante"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Anulacion_Masa"
            Object.ToolTipText     =   "Anular varios comprobantes"
            ImageIndex      =   11
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4320
      Left            =   105
      TabIndex        =   19
      Top             =   2940
      Width           =   11355
      _ExtentX        =   20029
      _ExtentY        =   7620
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "&4.- CONTABILIZACION"
      TabPicture(0)   =   "FLstComp.frx":537A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "DGAsientos"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&5.- RETENCIONES"
      TabPicture(1)   =   "FLstComp.frx":5396
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DGFAV"
      Tab(1).Control(1)=   "DGFAC"
      Tab(1).Control(2)=   "DGRet"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "&6.- SUBCUENTAS"
      TabPicture(2)   =   "FLstComp.frx":53B2
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "DGIxExCC"
      Tab(2).Control(1)=   "DGCxCxP"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "&7.- KARDEX"
      TabPicture(3)   =   "FLstComp.frx":53CE
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "LblTotalC"
      Tab(3).Control(1)=   "LblTotalComp"
      Tab(3).Control(2)=   "Label10"
      Tab(3).Control(3)=   "Label11"
      Tab(3).Control(4)=   "DGKardex"
      Tab(3).ControlCount=   5
      Begin MSDataGridLib.DataGrid DGAsientos 
         Bindings        =   "FLstComp.frx":53EA
         Height          =   2535
         Left            =   105
         TabIndex        =   36
         ToolTipText     =   "<Ctrl+D> Elimina Cuenta, <Ctrl+Z> Cambiar Cuenta, <Ctrl+X> Cambiar Valor del Debe. Haber, etc."
         Top             =   420
         Width           =   9990
         _ExtentX        =   17621
         _ExtentY        =   4471
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
      Begin MSDataGridLib.DataGrid DGRet 
         Bindings        =   "FLstComp.frx":5404
         Height          =   1695
         Left            =   -74895
         TabIndex        =   22
         Top             =   1365
         Width           =   9990
         _ExtentX        =   17621
         _ExtentY        =   2990
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
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
      Begin MSDataGridLib.DataGrid DGCxCxP 
         Bindings        =   "FLstComp.frx":5419
         Height          =   1800
         Left            =   -74895
         TabIndex        =   23
         Top             =   435
         Width           =   9675
         _ExtentX        =   17066
         _ExtentY        =   3175
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
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
      Begin MSDataGridLib.DataGrid DGKardex 
         Bindings        =   "FLstComp.frx":5430
         Height          =   3060
         Left            =   -74895
         TabIndex        =   24
         ToolTipText     =   "<Ctrl+L> Colocar el Lote"
         Top             =   420
         Width           =   9990
         _ExtentX        =   17621
         _ExtentY        =   5398
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
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
      Begin MSDataGridLib.DataGrid DGFAC 
         Bindings        =   "FLstComp.frx":5448
         Height          =   855
         Left            =   -74895
         TabIndex        =   20
         Top             =   420
         Width           =   9990
         _ExtentX        =   17621
         _ExtentY        =   1508
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
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
      Begin MSDataGridLib.DataGrid DGFAV 
         Bindings        =   "FLstComp.frx":545D
         Height          =   960
         Left            =   -74895
         TabIndex        =   21
         Top             =   3150
         Width           =   9990
         _ExtentX        =   17621
         _ExtentY        =   1693
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
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
      Begin MSDataGridLib.DataGrid DGIxExCC 
         Bindings        =   "FLstComp.frx":5472
         Height          =   1800
         Left            =   -74895
         TabIndex        =   37
         Top             =   2310
         Width           =   9675
         _ExtentX        =   17066
         _ExtentY        =   3175
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
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
      Begin VB.Label Label11 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Total Costo"
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
         TabIndex        =   27
         Top             =   3885
         Width           =   1275
      End
      Begin VB.Label Label10 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Total Compra"
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
         Left            =   -71010
         TabIndex        =   25
         Top             =   3885
         Width           =   1380
      End
      Begin VB.Label LblTotalComp 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   -69645
         TabIndex        =   26
         Top             =   3885
         Width           =   1695
      End
      Begin VB.Label LblTotalC 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   -66600
         TabIndex        =   28
         Top             =   3885
         Width           =   1695
      End
   End
   Begin MSAdodcLib.Adodc AdoRet 
      Height          =   330
      Left            =   315
      Top             =   3255
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
   Begin MSAdodcLib.Adodc AdoIngKar 
      Height          =   330
      Left            =   315
      Top             =   3570
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
      Caption         =   "IngKar"
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
   Begin MSAdodcLib.Adodc AdoIxExCC 
      Height          =   330
      Left            =   315
      Top             =   3885
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
      Caption         =   "IxExCC"
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
   Begin MSAdodcLib.Adodc AdoComp 
      Height          =   330
      Left            =   315
      Top             =   4200
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
      Caption         =   "Comp"
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
   Begin MSAdodcLib.Adodc AdoCompNo 
      Height          =   330
      Left            =   315
      Top             =   4515
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
      Caption         =   "CompNo"
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
   Begin MSAdodcLib.Adodc AdoTrans 
      Height          =   330
      Left            =   315
      Top             =   4830
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
      Caption         =   "Trans"
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
      Left            =   315
      Top             =   5775
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
   Begin MSAdodcLib.Adodc AdoCtas 
      Height          =   330
      Left            =   315
      Top             =   5460
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
   Begin MSAdodcLib.Adodc AdoBanco 
      Height          =   330
      Left            =   315
      Top             =   5145
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
   Begin MSAdodcLib.Adodc AdoCxCxP 
      Height          =   330
      Left            =   315
      Top             =   6090
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
      Caption         =   "CxCxP"
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
   Begin MSAdodcLib.Adodc AdoFAC 
      Height          =   330
      Left            =   315
      Top             =   6405
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
      Caption         =   "FAC"
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
   Begin MSAdodcLib.Adodc AdoFAV 
      Height          =   330
      Left            =   315
      Top             =   6720
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
      Caption         =   "FAV"
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
   Begin MSMask.MaskEdBox MBFecha 
      Height          =   330
      Left            =   1575
      TabIndex        =   39
      Top             =   1365
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   582
      _Version        =   393216
      AllowPrompt     =   -1  'True
      AutoTab         =   -1  'True
      Enabled         =   0   'False
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
   Begin VB.Label LabelEst 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Normal"
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
      Left            =   2835
      TabIndex        =   10
      Top             =   1365
      Width           =   2640
   End
   Begin VB.Label LabelUsuario 
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
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   1575
      TabIndex        =   30
      Top             =   7455
      Width           =   4320
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Elaborado por:"
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
      TabIndex        =   29
      Top             =   7455
      Width           =   1485
   End
   Begin VB.Label LabelHaber 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
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
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   8610
      TabIndex        =   33
      Top             =   7455
      Width           =   1695
   End
   Begin VB.Label LabelDebe 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
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
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   6930
      TabIndex        =   32
      Top             =   7455
      Width           =   1695
   End
   Begin VB.Label LabelFormaPago 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
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
      Left            =   9765
      TabIndex        =   16
      Top             =   1785
      Width           =   1695
   End
   Begin VB.Label LabelRecibi 
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
      Left            =   1470
      TabIndex        =   14
      Top             =   1785
      Width           =   7260
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
      Height          =   645
      Left            =   1470
      TabIndex        =   18
      Top             =   2205
      Width           =   9990
   End
   Begin VB.Label LabelCantidad 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
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
      Left            =   9765
      TabIndex        =   12
      Top             =   1365
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   " Pagado a:"
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
      TabIndex        =   13
      Top             =   1785
      Width           =   1275
   End
   Begin VB.Label Label4 
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
      Height          =   225
      Left            =   8820
      TabIndex        =   11
      Top             =   1470
      Width           =   855
   End
   Begin VB.Label Label6 
      Caption         =   " Efectivo"
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
      Left            =   8820
      TabIndex        =   15
      Top             =   1890
      Width           =   855
   End
   Begin VB.Label Label7 
      Caption         =   "Por concepto de:"
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
      Left            =   105
      TabIndex        =   17
      Top             =   2205
      Width           =   1275
   End
   Begin VB.Label Label19 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Totales"
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
      Left            =   5985
      TabIndex        =   31
      Top             =   7455
      Width           =   960
   End
End
Attribute VB_Name = "FListComprobantes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim IDTipoComp As Integer
Dim FechaTemp As String

Public Sub Imprimir_Cheques_En_Bloque()
On Error GoTo Errorhandler
RatonReloj
Mensajes = "Imprimir Cheques" & vbCrLf _
         & "De: " & NombreCliente & vbCrLf
Titulo = "IMPRESION"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
   Escala_Centimetro Orientacion_Pagina, TipoTimes, 10
   InicioX = 0.5: InicioY = 0
   Pagina = 1
   PosLinea = 0
  'Iniciamos la impresion
   Printer.FontBold = True
   With AdoAsientos.Recordset
    If .RecordCount > 0 Then
       .MoveFirst
        Printer.FontBold = True
        Printer.FontSize = 10
        Printer.FontName = TipoCourier
        Do While Not .EOF
           If IsNumeric(.fields("Cheq_Dep")) And .fields("Haber") > 0 Then
              Codigo = .fields("Cta")
              CCHQConLineas = ProcesarSeteos(Format(Val(MidStrg(Codigo, Len(Codigo) - 1, 3)), "00"))
              Total = .fields("Haber")
              Printer.FontSize = SetD(2).Tamaño
              PrinterTexto SetD(2).PosX, PosLinea + SetD(2).PosY, NombreCliente
              Printer.FontSize = SetD(3).Tamaño
              PrinterTexto SetD(3).PosX, PosLinea + SetD(3).PosY, Format(Total, "#,###.00")
              Printer.FontSize = SetD(10).Tamaño
              PrinterTexto SetD(10).PosX, PosLinea + SetD(10).PosY, ULCase(NombreCiudad)
              Printer.FontSize = SetD(6).Tamaño
              PrinterTexto SetD(6).PosX, PosLinea + SetD(6).PosY, Format(FechaTexto, "yyyy/MM/dd")
              If SetD(4).PosX > 0 And SetD(4).PosY > 0 Then
                 Printer.FontSize = SetD(4).Tamaño
                 PrinterNumCheque SetD(4).PosX, PosLinea + SetD(4).PosY, SetD(5).PosX, Total
              End If
              If SetD(9).PosX > 0 And SetD(9).PosY > 0 Then
                 Cadena = Empresa & " " & Moneda & " " & Format(Total, "#,##0.00") & "**"
                 Printer.FontSize = SetD(9).Tamaño
                 PrinterTexto SetD(9).PosX, PosLinea + SetD(9).PosY, Cadena
              End If
              Printer.NewPage
              Pagina = Pagina + 1
           End If
          .MoveNext
        Loop
       .MoveFirst
        Printer.FontBold = False
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

Private Function ObtenerNumeroDeComp() As Comprobantes
Dim NCo As Comprobantes
  NumItem = MBoxEmpNo
  If OpcTP(0).value Then NCo.TP = CompDiario
  If OpcTP(1).value Then NCo.TP = CompIngreso
  If OpcTP(2).value Then NCo.TP = CompEgreso
  If OpcTP(3).value Then NCo.TP = CompNotaDebito
  If OpcTP(4).value Then NCo.TP = CompNotaCredito
  NCo.Item = MBoxEmpNo
  NCo.Numero = Val(DCComp)
  NCo.Concepto = LabelConcepto.Caption
  ObtenerNumeroDeComp = NCo
End Function

Private Sub CMeses_GotFocus()
  MarcarTexto CMeses
End Sub

Private Sub CMeses_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub CMeses_LostFocus()
  ListarTipoComp IDTipoComp
End Sub

'''Private Sub Command5_Click()
'''  If Len(FA.Autorizacion_R) = 13 And FA.Estado_SRI <> "OK" Then
'''     SRI_Autorizacion = SRI_Generar_XML(FA.Autorizacion, FA.Estado_SRI)
'''    'Pasamos a Colocar el formato XML del Comprobante
'''    If SRI_Autorizacion.Estado_SRI <> "OK" Then TextoXML = Replace(TextoXML, "'", " ")
'''    sSQL = "UPDATE Trans_Compras " _
'''         & "SET Estado_SRI = '" & FA.Estado_SRI & "', " _
'''         & "Clave_Acceso = '" & FA.ClaveAcceso & "' " _
'''         & "WHERE Item = '" & NumEmpresa & "' " _
'''         & "AND Periodo = '" & Periodo_Contable & "' " _
'''         & "AND TP = '" & Co.TP & "' " _
'''         & "AND Numero = '" & Co.Numero & "' " _
'''         & "AND Fecha = #" & Co.Fecha & "# "
'''    Ejecutar_SQL_SP sSQL
'''    RatonReloj
'''    If SRI_Autorizacion.Estado_SRI = "OK" Then
'''      'Generar_PDF_FA_SRI TFA
'''       sSQL = "UPDATE Trans_Compras " _
'''            & "SET AutRetencion = '" & Autorizacion & "', " _
'''            & "Fecha_Aut = #" & BuscarFecha(Mifecha) & "#, " _
'''            & "Hora_Aut = '" & MiHora & "' " _
'''            & "WHERE Item = '" & NumEmpresa & "' " _
'''            & "AND Periodo = '" & Periodo_Contable & "' " _
'''            & "AND TP = '" & Co.TP & "' " _
'''            & "AND Numero = '" & Co.Numero & "' " _
'''            & "AND Fecha = #" & Co.Fecha & "# "
'''       Ejecutar_SQL_SP sSQL
'''       RatonReloj
'''       If EsUnEmail(FA.EmailC) Then
'''          TMail.Mensaje = "Cliente: " & FA.Cliente & vbCrLf _
'''                        & "Clave de Acceso: " & vbCrLf & FA.ClaveAcceso & vbCrLf _
'''                        & "Hora de Generacion: " & MiHora & vbCrLf _
'''                        & "Emision: " & Mifecha & vbCrLf _
'''                        & "Vencimiento: " & Mifecha & vbCrLf _
'''                        & "Autorizacion: " & vbCrLf & Autorizacion & vbCrLf _
'''                        & "Factura No. " & FA.Serie & "-" & Format(FA.Factura, "000000000") & vbCrLf
'''          TMail.Asunto = "Retencion No. " & FA.Serie_R & "-" & Format(FA.Retencion, "000000000")
'''          TMail.Adjunto = RutaSysBases & "\AT\Comprobantes Autorizados\" & FA.ClaveAcceso & ".xml"
'''          TMail.para = FA.EmailC
'''          FEnviarCorreos.Show 1
'''          TMail.Adjunto = RutaSysBases & "\AT\Comprobantes Autorizados PDF\" & FA.ClaveAcceso & ".pdf"
'''          FEnviarCorreos.Show 1
'''          RatonNormal
'''        End If
'''    Else
'''       RatonNormal
'''       MsgBox Cadena & "(" & FA.Estado_SRI & ")"
'''    End If
'''  Else
'''     MsgBox "Esta Factura ya esta autorizada"
'''  End If
'''End Sub

Private Sub Command1_Click()
   Control_Procesos Normal, "Salio Listar Comprobantes"
   Unload FListComprobantes
End Sub

Private Sub Command2_Click()
     Mensajes = "Seguro cambiar la Fecha del Comprobante"
     Titulo = "PREGUNTA DE MODIFICACION"
     If BoxMensaje = vbYes Then
        MBFecha.Enabled = True
        FechaTemp = MBFecha.Text
        MBFecha.SetFocus
     Else
        MBFecha.Enabled = False
     End If
End Sub

Private Sub DCComp_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCComp_LostFocus()
  ListarComprobantes
End Sub

Private Sub DGAsientos_KeyDown(KeyCode As Integer, Shift As Integer)
Dim ID_Temp As Long
  Keys_Especiales Shift
  If CtrlDown And KeyCode = vbKeyF1 Then
     DGAsientos.Visible = False
     GenerarDataTexto FListComprobantes, AdoAsientos
     DGAsientos.Visible = True
  End If
  If CtrlDown And KeyCode = vbKeyZ Then
     If ClaveContador Then
        RatonReloj
        'DGAsientos.Visible = False
        Producto = "Transacciones"
        Codigo1 = DGAsientos.Columns(0)
        Asiento = DGAsientos.Columns(13)
        Codigo3 = Codigo1 & " - " & DGAsientos.Columns(1)
        FChangeCta.Show 1
        RatonNormal
        'DGAsientos.Visible = True
     End If
  End If
  If CtrlDown And KeyCode = vbKeyD Then
     DGAsientos.Visible = False
     Cta = DGAsientos.Columns(0)
     CuentaBanco = DGAsientos.Columns(1)
     Debe = DGAsientos.Columns(3)
     Haber = DGAsientos.Columns(4)
     Asiento = DGAsientos.Columns(13)
     
     Mensajes = "Esta seguro de eliminar la cuenta:" & vbCrLf & vbCrLf & Cta & " - " & CuentaBanco
     Titulo = "PREGUNTA DE ELIMINACION"
     If BoxMensaje = vbYes Then
        Actualiza_Procesado_Tabla "Transacciones", True
        Actualiza_Procesado_Tabla "Trans_SubCtas", True
        Actualiza_Procesado_Tabla "Trans_Kardex", True
        
        Elimina_Cuenta_Tabla "Transacciones", "Cta", Cta, ID_Temp
        Elimina_Cuenta_Tabla "Trans_SubCtas", "Cta", Cta
        Elimina_Cuenta_Tabla "Trans_Air", "Cta_Retencion", Cta
        Elimina_Cuenta_Tabla "Trans_Compras", "Cta_Servicio", Cta
        Elimina_Cuenta_Tabla "Trans_Compras", "Cta_Bienes", Cta
        Elimina_Cuenta_Tabla "Trans_Kardex", "Cta_Inv", Cta
        Elimina_Cuenta_Tabla "Trans_Kardex", "Contra_Cta", Cta
        
        MsgBox "Proceso terminado con exito," & vbCrLf & vbCrLf & "vuelva a listar el comprobante"
    End If
       'MsgBox sSQL
     DGAsientos.Visible = True
     DCComp.SetFocus
  End If
  If CtrlDown And KeyCode = vbKeyX Then
     If ClaveContador Then
        RatonReloj
        NomCta = LabelConcepto.Caption
        Cta = DGAsientos.Columns(0)
        Cuenta_No = DGAsientos.Columns(1)
        Debe = DGAsientos.Columns(3)
        Haber = DGAsientos.Columns(4)
        NomCtaSup = DGAsientos.Columns(5)
        NoCheque = DGAsientos.Columns(6)
        Asiento = DGAsientos.Columns(13)
        FChangeValores.Show
        RatonNormal
     End If
  End If
End Sub

Private Sub DGCxCxP_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  If CtrlDown And KeyCode = vbKeyF1 Then
     DGCxCxP.Visible = False
     GenerarDataTexto FListComprobantes, AdoCxCxP
     DGCxCxP.Visible = True
  End If
  If CtrlDown And KeyCode = vbKeyP Then
     MsgBox "cambiar"
  End If
End Sub

Private Sub DGKardex_KeyDown(KeyCode As Integer, Shift As Integer)
Dim Lote As String
  Keys_Especiales Shift
  If CtrlDown And KeyCode = vbKeyL Then
     Lote = InputBox("Ingrese el Lote", "INGRESO DE LOTES", "0")
     If Len(Lote) > 1 Then
        sSQL = "UPDATE Trans_Kardex " _
             & "SET Lote_No = '" & Lote & "' " _
             & "WHERE Item = '" & NumEmpresa & "' " _
             & "AND Periodo = '" & Periodo_Contable & "' " _
             & "AND TP = '" & Co.TP & "' " _
             & "AND Numero = '" & Co.Numero & "' " _
             & "AND Fecha = #" & BuscarFecha(Co.Fecha) & "# "
        Ejecutar_SQL_SP sSQL
        MsgBox "Proceso Exitoso"
     End If
  End If
  If CtrlDown And KeyCode = vbKeyF1 Then
     DGKardex.Visible = False
     GenerarDataTexto FListComprobantes, AdoIngKar
     DGKardex.Visible = True
  End If
End Sub

''''Private Sub DGSubCtas_KeyDown(KeyCode As Integer, Shift As Integer)
''''  Keys_Especiales Shift
''''  If CtrlDown And KeyCode = vbKeyF1 Then
''''     DGSubCtas.Visible = False
''''     GenerarDataTexto FListComprobantes, AdoSubCtas
''''     DGSubCtas.Visible = True
''''  End If
''''End Sub

Private Sub Form_Activate()
   SSTab1.Height = MDI_Y_Max - SSTab1.Top - 300
   SSTab1.width = MDI_X_Max - SSTab1.Left
   
   DGCxCxP.width = SSTab1.width - 200
   DGIxExCC.width = SSTab1.width - 200
   DGCxCxP.Height = (SSTab1.Height - DGCxCxP.Top) / 2 - 100
   DGIxExCC.Height = (SSTab1.Height - DGCxCxP.Top) / 2 - 100
   DGIxExCC.Top = DGCxCxP.Top + DGCxCxP.Height + 10
   
   DGFAC.width = SSTab1.width - 200
   
   DGRet.Top = DGFAC.Top + DGFAC.Height + 10
   DGRet.width = SSTab1.width - 200
   DGRet.Height = (SSTab1.Height - DGRet.Top - DGRet.Height) / 2 - 100
   
   DGFAV.Top = DGRet.Top + DGRet.Height + 10
   DGFAV.width = SSTab1.width - 200
   DGFAV.Height = SSTab1.Height - DGFAV.Top - 100
      
   DGKardex.width = SSTab1.width - 200
   DGKardex.Height = SSTab1.Height - DGAsientos.Top - 500
   Label10.Top = DGKardex.Top + DGKardex.Height + 50
   Label11.Top = DGKardex.Top + DGKardex.Height + 50
   LblTotalComp.Top = DGKardex.Top + DGKardex.Height + 50
   LblTotalC.Top = DGKardex.Top + DGKardex.Height + 50
   
   DGAsientos.width = SSTab1.width - 200
   DGAsientos.Height = SSTab1.Height - DGAsientos.Top - 100
  
   Label2.Top = SSTab1.Top + SSTab1.Height + 50
   Label19.Top = SSTab1.Top + SSTab1.Height + 50
   LabelDebe.Top = SSTab1.Top + SSTab1.Height + 50
   LabelHaber.Top = SSTab1.Top + SSTab1.Height + 50
   LabelUsuario.Top = SSTab1.Top + SSTab1.Height + 50
   Command1.Top = SSTab1.Top + SSTab1.Height + 50
   
   CMeses.Clear
   CMeses.AddItem "Todos"
   CMeses.AddItem "Enero"
   CMeses.AddItem "Febrero"
   CMeses.AddItem "Marzo"
   CMeses.AddItem "Abril"
   CMeses.AddItem "Mayo"
   CMeses.AddItem "Junio"
   CMeses.AddItem "Julio"
   CMeses.AddItem "Agosto"
   CMeses.AddItem "Septiembre"
   CMeses.AddItem "Octubre"
   CMeses.AddItem "Noviembre"
   CMeses.AddItem "Diciembre"
   CMeses.Text = "Todos"
   
   MBoxEmpNo = Format(NumEmpresa, "000")

  IDTipoComp = 0
  If Modulo = "GERENCIA" Then
     Toolbar1.buttons("Modificar_Comprobante").Enabled = False
     Toolbar1.buttons("Anular_Comprobante").Enabled = False
     Toolbar1.buttons("Autorizar_Comprobante").Enabled = False
     Toolbar1.buttons("Copiar_Comprobante").Enabled = False
     Toolbar1.buttons("Copiar_Item").Enabled = False
  End If
  If Bloquear_Control Then
     Toolbar1.buttons("Imprimir_Comprobante").Enabled = False
     Toolbar1.buttons("Imprimir_Retencion").Enabled = False
     Toolbar1.buttons("Imprimir_Cheque").Enabled = False
     Toolbar1.buttons("Modificar_Comprobante").Enabled = False
     Toolbar1.buttons("Anular_Comprobante").Enabled = False
     Toolbar1.buttons("Copiar_Comprobante").Enabled = False
     Toolbar1.buttons("Copiar_Item").Enabled = False
     Toolbar1.buttons("Autorizar_Comprobante").Enabled = False
  End If
  ListarTipoComp IDTipoComp
 'Determinamos si es matriz o sucursal
  MBoxEmpNo.Enabled = False

  If ConSucursal Then
     MBoxEmpNo.Enabled = True
     Co.Item = MBoxEmpNo
     MBoxEmpNo.SetFocus
  Else
     Co.Item = NumEmpresa
     OpcTP(0).SetFocus
  End If
  FListComprobantes.WindowState = 2
  RatonNormal
End Sub

Private Sub Form_Deactivate()
  RatonNormal
  FListComprobantes.WindowState = 1
End Sub

Private Sub Form_Load()
  'Abriendo Bases Relacionadas
   IDTipoComp = 0
   ConectarAdodc AdoFAC
   ConectarAdodc AdoFAV
   ConectarAdodc AdoRet
   ConectarAdodc AdoComp
   ConectarAdodc AdoCtas
   ConectarAdodc AdoBanco
   ConectarAdodc AdoTrans
   ConectarAdodc AdoCompNo
   ConectarAdodc AdoCxCxP
   ConectarAdodc AdoIngKar
   ConectarAdodc AdoIxExCC
   ConectarAdodc AdoAsientos
End Sub

Private Sub MBFecha_GotFocus()
  MarcarTexto MBFecha
End Sub

Private Sub MBFecha_LostFocus()
  FechaValida MBFecha
  If MBFecha.Text <> FechaTemp Then
     Mensajes = "Seguro de realizar el cambio"
     Titulo = "PREGUNTA DE MODIFICACION"
     If BoxMensaje = vbYes Then
        Actualiza_Fecha_Tabla "Comprobantes", MBFecha
        Actualiza_Fecha_Tabla "Trans_Air", MBFecha
        Actualiza_Fecha_Tabla "Trans_Compras", MBFecha
        Actualiza_Fecha_Tabla "Trans_Exportaciones", MBFecha
        Actualiza_Fecha_Tabla "Trans_Importaciones", MBFecha
        Actualiza_Fecha_Tabla "Trans_Kardex", MBFecha
        Actualiza_Fecha_Tabla "Trans_Rol_Pagos", MBFecha
        Actualiza_Fecha_Tabla "Trans_SubCtas", MBFecha
        Actualiza_Fecha_Tabla "Trans_Ventas", MBFecha
        Actualiza_Fecha_Tabla "Transacciones", MBFecha

        Actualiza_Procesado_Tabla "Transacciones", True
        Actualiza_Procesado_Tabla "Trans_SubCtas", True
        Actualiza_Procesado_Tabla "Trans_Kardex", True
     
        MBFecha.Enabled = False
        MsgBox "Proceso terminado con exito," & vbCrLf & vbCrLf & "vuelva a listar el comprobante"
    End If
  End If
End Sub

Private Sub MBoxEmpNo_GotFocus()
  MarcarTexto MBoxEmpNo
End Sub

Private Sub MBoxEmpNo_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBoxEmpNo_LostFocus()
  If Val(MBoxEmpNo) = 0 Then MBoxEmpNo = Format(NumEmpresa, "000")
  Co.Item = MBoxEmpNo
End Sub

Public Sub ListarComprobantes()
  RatonReloj
  ModificarComp = False
  CopiarComp = False
  NuevoComp = False
  ExisteComp = False
  
  DGRet.Visible = False
  DGFAC.Visible = False
  DGFAV.Visible = False
  DGCxCxP.Visible = False
  DGIxExCC.Visible = False
  DGKardex.Visible = False
  DGAsientos.Visible = False
'''  DGSubCtas.Visible = False
  LabelEst.Caption = Ninguno
  MBFecha.Text = "00/00/0000"
  LabelRecibi.Caption = Ninguno
  LabelConcepto.Caption = Ninguno
  LabelFormaPago.Caption = Ninguno
  LabelUsuario.Caption = Ninguno
  SumaDebe = 0: SumaHaber = 0
  Co.Fecha = FechaSistema
  Co.Numero = Val(DCComp)
  Co.CodigoB = Ninguno
  Co.CodigoDr = Ninguno
 'Llenar Encabezado del Comprobante
  sSQL = "SELECT C.*,A.Nombre_Completo,Cl.CI_RUC,Cl.Direccion,Cl.Email,Cl.Telefono,Cl.Celular,Cl.Cliente,Cl.Ciudad " _
       & "FROM Comprobantes As C, Accesos As A, Clientes As Cl " _
       & "WHERE C.Numero = " & Co.Numero & " " _
       & "AND C.TP = '" & Co.TP & "' " _
       & "AND C.Item = '" & Co.Item & "' " _
       & "AND C.Periodo = '" & Periodo_Contable & "' " _
       & "AND C.CodigoU = A.Codigo " _
       & "AND C.Codigo_B = Cl.Codigo "
  Select_Adodc AdoComp, sSQL
  With AdoComp.Recordset
   If .RecordCount > 0 Then
       LabelEst.Caption = "Normal"
       Co.T = .fields("T")
       Co.Fecha = .fields("Fecha")
       Co.CodigoB = .fields("Codigo_B")
       Co.Beneficiario = .fields("Cliente")
       Co.Concepto = .fields("Concepto")
       Co.Efectivo = .fields("Efectivo")
       If Co.T = Anulado Then LabelEst.Caption = "Anulado"
       MBFecha.Text = Format(Co.Fecha, FormatoFechas)
       LabelRecibi.Caption = Co.Beneficiario
       LabelConcepto.Caption = Co.Concepto
       LabelFormaPago.Caption = Co.Efectivo
       LabelUsuario.Caption = " " & .fields("Nombre_Completo")
       ExisteComp = True
       RatonNormal
   Else
       MsgBox "El Comprobante no exite."
       If MBoxEmpNo.Enabled Then MBoxEmpNo.SetFocus Else OpcTP(IDTipoComp).SetFocus
   End If
  End With
  If ExisteComp Then
    'Llenar Cuentas de Transacciones
     Co.Cheque = Ninguno
     Co.Cta_Banco = Ninguno
     sSQL = "SELECT T.Cta,Ca.Cuenta,T.Parcial_ME,T.Debe,T.Haber,T.Detalle,T.Cheq_Dep,T.Fecha_Efec,T.Codigo_C,Ca.Item,T.TP,T.Numero,T.Fecha,T.ID " _
          & "FROM Transacciones As T, Catalogo_Cuentas As Ca " _
          & "WHERE T.TP = '" & Co.TP & "' " _
          & "AND T.Numero = " & Co.Numero & " " _
          & "AND T.Item = '" & Co.Item & "' " _
          & "AND T.Periodo = '" & Periodo_Contable & "' " _
          & "AND T.Item = Ca.Item " _
          & "AND T.Periodo = Ca.Periodo " _
          & "AND T.Cta = Ca.Codigo " _
          & "ORDER BY T.ID,Debe DESC,T.Cta "
     Select_Adodc_Grid DGAsientos, AdoAsientos, sSQL
     DGAsientos.Visible = False
     With AdoAsientos.Recordset
      If .RecordCount > 0 Then
         .MoveFirst
          Do While Not .EOF
             SumaDebe = SumaDebe + .fields("Debe")
             SumaHaber = SumaHaber + .fields("Haber")
             
             If IsNumeric(.fields("Cheq_Dep")) And .fields("Haber") > 0 Then
                Co.Cheque = .fields("Cheq_Dep")
                Co.Cta_Banco = .fields("Cta")
             End If
            .MoveNext
          Loop
         .MoveFirst
      End If
     End With
     SumaDebe = Redondear(SumaDebe, 2)
     SumaHaber = Redondear(SumaHaber, 2)
     DGAsientos.Visible = True
     
    'Llenar Inventario
     sSQL = "SELECT TK.Codigo_Inv,CC.Producto,TK.CodBodega,TK.Entrada,TK.Salida,TK.Valor_Unitario,TK.Valor_Total,TK.Costo,TK.Total," _
          & "TK.Fecha,TK.TC,TK.Serie,TK.Factura,TK.Orden_No,TK.Lote_No,TK.Fecha_Fab,TK.Fecha_Exp,CC.Reg_Sanitario,TK.Modelo," _
          & "TK.Procedencia,TK.Serie_No,TK.Codigo_Barra,TK.Cta_Inv,TK.Contra_Cta " _
          & "FROM Trans_Kardex As TK,Catalogo_Productos As CC " _
          & "WHERE TK.Periodo = '" & Periodo_Contable & "' " _
          & "AND TK.Item = '" & Co.Item & "' " _
          & "AND TK.TP = '" & Co.TP & "' " _
          & "AND TK.Numero = " & Co.Numero & " " _
          & "AND TK.Codigo_Inv = CC.Codigo_Inv " _
          & "AND TK.Periodo = CC.Periodo " _
          & "AND TK.Item = CC.Item " _
          & "ORDER BY TK.Codigo_Inv, TK.CodBodega, TK.Fecha "
     SQLDec = "Valor_Unitario " & CStr(Dec_Costo) & "|Costo " & CStr(Dec_Costo) & "|."
     Select_Adodc_Grid DGKardex, AdoIngKar, sSQL, SQLDec
     Total = 0: Saldo = 0
     With AdoIngKar.Recordset
      If .RecordCount > 0 Then
          Do While Not .EOF
             Total = Total + Redondear(.fields("Valor_Total"), 2)
             Saldo = Saldo + Redondear(.fields("Total"), 2)
            .MoveNext
          Loop
          DGKardex.Visible = True
      End If
     End With
     LblTotalComp.Caption = Format(Total, "#,##0.00")
     LblTotalC.Caption = Format(Saldo, "#,##0.00")
     
    'Listar SubCtas CxC/CxP
     sSQL = "SELECT TSC.Cta,TSC.TC,Be.Cliente As Detalles,Factura,Fecha_V,Parcial_ME,Debitos,Creditos,Prima," _
          & "TSC.Detalle_SubCta,Be.Codigo " _
          & "FROM Trans_SubCtas As TSC, Clientes As Be " _
          & "WHERE TSC.Item = '" & Co.Item & "' " _
          & "AND TSC.Periodo = '" & Periodo_Contable & "' " _
          & "AND TSC.TP = '" & Co.TP & "' " _
          & "AND TSC.Numero = " & Co.Numero & " " _
          & "AND TSC.TC IN ('C','P') " _
          & "AND TSC.Codigo = Be.Codigo " _
          & "ORDER BY TSC.Cta, TSC.TC "
     Select_Adodc_Grid DGCxCxP, AdoCxCxP, sSQL
    
     sSQL = "SELECT TSC.Cta,TSC.TC,Be.Detalle As Detalles,Factura,Fecha_V,Parcial_ME,Debitos,Creditos,Prima," _
          & "TSC.Detalle_SubCta,Be.Codigo " _
          & "FROM Trans_SubCtas As TSC, Catalogo_SubCtas As Be " _
          & "WHERE TSC.TP = '" & Co.TP & "' " _
          & "AND TSC.Numero = " & Co.Numero & " " _
          & "AND TSC.Item = '" & Co.Item & "' " _
          & "AND TSC.Periodo = '" & Periodo_Contable & "' " _
          & "AND TSC.Codigo = Be.Codigo " _
          & "AND TSC.Periodo = Be.Periodo " _
          & "AND TSC.Item = Be.Item " _
          & "ORDER BY TSC.Cta, TSC.TC "
     Select_Adodc_Grid DGIxExCC, AdoIxExCC, sSQL
     
     If AdoIxExCC.Recordset.RecordCount > 0 Then DGIxExCC.Visible = True
     If AdoCxCxP.Recordset.RecordCount > 0 Then DGCxCxP.Visible = True
     
    'Listar las Retenciones
     sSQL = "SELECT * " _
          & "FROM Trans_Air " _
          & "WHERE Periodo = '" & Periodo_Contable & "' " _
          & "AND Item = '" & Co.Item & "' " _
          & "AND Numero = " & Co.Numero & " " _
          & "AND TP = '" & Co.TP & "' "
     Select_Adodc_Grid DGRet, AdoRet, sSQL
     If AdoRet.Recordset.RecordCount > 0 Then DGRet.Visible = True
     
     'Compras
      FechaIni = BuscarFecha(Co.Fecha)
      FechaFin = BuscarFecha(Co.Fecha)
      sSQL = "SELECT TC.Linea_SRI,C.Cliente,C.TD,C.CI_RUC,TC.TipoComprobante,TC.CodSustento,TC.TP,TC.Numero," _
           & "TC.Establecimiento,TC.PuntoEmision,TC.Secuencial,TC.Autorizacion,TC.FechaEmision,TC.FechaRegistro," _
           & "TC.FechaCaducidad,TC.BaseImponible,TC.BaseImpGrav,TC.PorcentajeIva,TC.MontoIva,TC.BaseImpIce," _
           & "TC.PorcentajeIce,TC.MontoIce,TC.MontoIvaBienes,TC.PorRetBienes,TC.ValorRetBienes,TC.MontoIvaServicios," _
           & "TC.PorRetServicios,TC.ValorRetServicios,TC.ContratoPartidoPolitico,TC.MontoTituloOneroso," _
           & "TC.MontoTituloGratuito,TC.DevIva,TC.Clave_Acceso,Estado_SRI,AutRetencion " _
           & "FROM Trans_Compras As TC, Clientes As C " _
           & "WHERE TC.TP = '" & Co.TP & "' " _
           & "AND TC.Numero = " & Co.Numero & " " _
           & "AND TC.Item = '" & Co.Item & "' " _
           & "AND TC.Periodo = '" & Periodo_Contable & "' " _
           & "AND TC.IdProv = C.Codigo " _
           & "ORDER BY TC.Linea_SRI,C.Cliente,C.CI_RUC,C.TD "
      Select_Adodc_Grid DGFAC, AdoFAC, sSQL
      If AdoFAC.Recordset.RecordCount > 0 Then
         FA.ClaveAcceso = AdoFAC.Recordset.fields("Clave_Acceso")
         FA.Estado_SRI = AdoFAC.Recordset.fields("Estado_SRI")
         FA.Autorizacion_R = AdoFAC.Recordset.fields("AutRetencion")
         DGFAC.Visible = True
      End If
     'Ventas
      sSQL = "SELECT TV.Linea_SRI,C.Cliente,TV.TipoComprobante,C.CI_RUC,C.TD,TV.TP,TV.Numero,TV.Establecimiento," _
           & "TV.PuntoEmision,TV.NumeroComprobantes,TV.Secuencial,TV.FechaEmision,TV.FechaRegistro,TV.BaseImponible," _
           & "TV.BaseImpGrav,TV.PorcentajeIva,TV.MontoIva,TV.BaseImpIce,TV.PorcentajeIce,TV.MontoIce,TV.MontoIvaBienes," _
           & "TV.PorRetBienes,TV.ValorRetBienes,TV.MontoIvaServicios,TV.PorRetServicios,TV.ValorRetServicios," _
           & "TV.RetPresuntiva " _
           & "FROM Trans_Ventas As TV, Clientes As C " _
           & "WHERE TV.TP = '" & Co.TP & "' " _
           & "AND TV.Numero = " & Co.Numero & " " _
           & "AND TV.Item = '" & Co.Item & "' " _
           & "AND TV.Periodo = '" & Periodo_Contable & "' " _
           & "AND TV.IdProv = C.Codigo  " _
           & "ORDER BY TV.Linea_SRI,C.Cliente,C.CI_RUC,C.TD "
      Select_Adodc_Grid DGFAV, AdoFAV, sSQL
      If AdoFAV.Recordset.RecordCount > 0 Then DGFAV.Visible = True
  End If
  LabelDebe.Caption = Format(SumaDebe, "#,##0.00")
  LabelHaber.Caption = Format(SumaHaber, "#,##0.00")
  LabelCantidad.Caption = Format(SumaDebe, "#,##0.00")
End Sub

Public Sub ListarTipoComp(dTipoComp As Integer)
Dim MesNo As Integer
  RatonReloj
  Co.TP = CompDiario
  Select Case dTipoComp
    Case 0: Co.TP = CompDiario
    Case 1: Co.TP = CompIngreso
    Case 2: Co.TP = CompEgreso
    Case 3: Co.TP = CompNotaDebito
    Case 4: Co.TP = CompNotaCredito
  End Select
  MesNo = CMeses.ListIndex
  If MesNo < 0 Then MesNo = 0
  sSQL = "SELECT Numero " _
       & "FROM Comprobantes " _
       & "WHERE Item = '" & Co.Item & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND TP = '" & Co.TP & "' "
  If MesNo > 0 Then sSQL = sSQL & "AND MONTH(Fecha) = " & MesNo & " "
  sSQL = sSQL & "ORDER BY Numero "
  SelectDB_Combo DCComp, AdoCompNo, sSQL, "Numero", True
  RatonNormal
End Sub

Public Sub AnularComprobante()
Dim AnularComprobanteDe As String
Dim MotivoAnulacion As String
Dim EsCxC As Boolean
Dim UnaFecha As Boolean
Dim CompCierre As Boolean
Dim IdF As Long

If ClaveSupervisor Then
   CompCierre = False
  'Co = ObtenerNumeroDeComp
   FechaInicial = Co.Fecha
   FechaFinal = Co.Fecha
   If InStr(Co.Concepto, "Cierre de Caja de") > 0 Then CompCierre = True
   If CompCierre Then
      If InStr(Co.Concepto, "Cierre de Caja de Cuentas por Cobrar") > 0 Then EsCxC = True Else EsCxC = False
      UnaFecha = True
      For IdF = 1 To Len(Co.Concepto)
          Mifecha = MidStrg(Co.Concepto, IdF, 10)
          If IsDate(Mifecha) Then
             If Year(Mifecha) >= 1900 Then
                If UnaFecha Then
                   FechaInicial = Mifecha
                   FechaFinal = Mifecha
                   UnaFecha = False
                   IdF = IdF + 10
                Else
                   FechaFinal = Mifecha
                   IdF = IdF + 10
                End If
             End If
          End If
      Next IdF
   End If
   FechaIni = BuscarFecha(FechaInicial)
   FechaFin = BuscarFecha(FechaFinal)
   AnularComprobanteDe = "WHERE Periodo = '" & Periodo_Contable & "' " _
                       & "AND Item = '" & Co.Item & "' " _
                       & "AND TP = '" & Co.TP & "' " _
                       & "AND Numero = " & Co.Numero & " "
   Mensajes = "Seguro de Anular" & vbCrLf & "El Comprobante No. " & Co.TP & "-" & Co.Numero
   Titulo = "Pregunta de Anulacion"
   If BoxMensaje = vbYes Then
      Control_Procesos "A", "Anulo Comprobante de: " & Co.TP & " No. " & Co.Numero
      If InStr(LabelConcepto.Caption, "(ANULADO)") > 0 Then
         Contra_Cta = LabelConcepto.Caption
      Else
         MotivoAnulacion = UCase(InputBox("MOTIVO DE LA ANULACION:", "FORMULARIO DE ANULACION", ""))
         Contra_Cta = "(ANULADO) "
         If MotivoAnulacion <> "" Then Contra_Cta = Contra_Cta & "[MOTIVO: " & MotivoAnulacion & "], "
         Contra_Cta = Contra_Cta & LabelConcepto.Caption
      End If
      Contra_Cta = MidStrg(Contra_Cta, 1, 120)
     'Actualizamos Comprobante
      sSQL = "UPDATE Comprobantes " _
           & "SET T = '" & Anulado & "', Concepto = '" & Contra_Cta & "' " _
           & AnularComprobanteDe
      Ejecutar_SQL_SP sSQL
     'Actualizar Transacciones
      sSQL = "UPDATE Transacciones " _
           & "SET T = '" & Anulado & "',Debe = 0,Haber = 0,Saldo = 0 " _
           & AnularComprobanteDe
      Ejecutar_SQL_SP sSQL
     'Actualizar Trans_SubCtas
      sSQL = "UPDATE Trans_SubCtas " _
           & "SET T = '" & Anulado & "',Debitos = 0,Creditos = 0,Saldo_MN = 0,Saldo_ME = 0,Prima = 0 " _
           & AnularComprobanteDe
      Ejecutar_SQL_SP sSQL
     'Actualizar Retencion
      sSQL = "DELETE * " _
           & "FROM Trans_Air " _
           & AnularComprobanteDe
      Ejecutar_SQL_SP sSQL
      sSQL = "DELETE * " _
           & "FROM Trans_Compras " _
           & AnularComprobanteDe
      Ejecutar_SQL_SP sSQL
      sSQL = "DELETE * " _
           & "FROM Trans_Ventas " _
           & AnularComprobanteDe
      Ejecutar_SQL_SP sSQL
      sSQL = "DELETE * " _
           & "FROM Trans_Exportaciones " _
           & AnularComprobanteDe
      Ejecutar_SQL_SP sSQL
      sSQL = "DELETE * " _
           & "FROM Trans_Importaciones " _
           & AnularComprobanteDe
      Ejecutar_SQL_SP sSQL
     
     'Eliminamos el Rol de Pagos
      sSQL = "DELETE * " _
           & "FROM Trans_Rol_de_Pagos " _
           & AnularComprobanteDe
      Ejecutar_SQL_SP sSQL
     'Actualizar Kardex
      If CompCierre Then
         sSQL = "UPDATE Trans_Kardex " _
              & "SET Procesado = 0, TP = '.', Numero = 0 " _
              & AnularComprobanteDe _
              & "AND Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
              & "AND LEN(Serie) > 1 " _
              & "AND Factura <> 0 "
         Ejecutar_SQL_SP sSQL
        'MsgBox "CompCierre" & vbCrLf & sSQL
      Else
         sSQL = "SELECT Codigo_Inv " _
              & "FROM Trans_Kardex " _
              & AnularComprobanteDe
         Select_Adodc AdoCxCxP, sSQL
         With AdoCxCxP.Recordset
          If .RecordCount > 0 Then
              Do While Not .EOF
                 Codigo = .fields("Codigo_Inv")
                 sSQL = "UPDATE Trans_Kardex " _
                      & "SET Procesado = 0 " _
                      & "WHERE Item = '" & Co.Item & "' " _
                      & "AND Periodo = '" & Periodo_Contable & "' " _
                      & "AND Codigo_Inv = '" & Codigo & "' "
                 Ejecutar_SQL_SP sSQL
                'MsgBox "Sin CompCierre" & vbCrLf & sSQL
                .MoveNext
              Loop
          End If
         End With
         
         sSQL = "DELETE * " _
              & "FROM Trans_Kardex " _
              & AnularComprobanteDe _
              & "AND TC = '" & Ninguno & "' " _
              & "AND Serie = '" & Ninguno & "' " _
              & "AND Factura = 0 "
         Ejecutar_SQL_SP sSQL
      End If
     
     'Actualizar las Ctas a mayoriazar
      sSQL = "SELECT Cta " _
           & "FROM Transacciones " _
           & AnularComprobanteDe _
           & "GROUP BY Cta " _
           & "ORDER BY Cta "
      Select_Adodc AdoCxCxP, sSQL
      With AdoCxCxP.Recordset
       If .RecordCount > 0 Then
           Do While Not .EOF
             'Determinamos que la cuenta ya fue mayorizada
              SubCta = .fields("Cta")
              sSQL = "UPDATE Transacciones " _
                   & "SET Procesado = " & Val(adFalse) & " " _
                   & "WHERE Item = '" & NumEmpresa & "' " _
                   & "AND Periodo = '" & Periodo_Contable & "' " _
                   & "AND Cta = '" & SubCta & "' "
              Ejecutar_SQL_SP sSQL
             .MoveNext
           Loop
       End If
      End With
   End If
End If
End Sub

Private Sub OpcTP_LostFocus(Index As Integer)
   IDTipoComp = Index
   ListarTipoComp IDTipoComp
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim ItemAux As String
 'MsgBox Co.Item & vbCrLf & Co.Numero & vbCrLf & Co.TP
  Select Case Button.key
    Case "Salir"
         ModificarComp = False
         CopiarComp = False
         NuevoComp = False
         ExisteComp = False
         Control_Procesos Normal, "Salio Listar Comprobantes"
         Unload FListComprobantes
    Case "Imprimir_Comprobante"
         Control_Procesos "I", "Imprimio Comprobante de: " & Co.TP & ", No. " & NumComp
         ImprimirComprobantesDe False, Co
         ListarComprobantes
         DCComp.SetFocus
    Case "Imprimir_Retencion"
         ImprimirComprobantesDe True, Co
    Case "Imprimir_Cheque"
        'MsgBox Co.Cheque & vbCrLf & Co.Cta_Banco
         Imprimir_Bloque_Cheques Co.Cheque, Co.Cheque, Co.Cta_Banco, Co.Cheque, "."
         'Imprimir_Cheques_En_Bloque
    Case "Modificar_Comprobante"
         If ClaveContador Then
            NumeroComp = Co.Numero
            Mensajes = "Seguro que quiere Modificar Comprobante " & Co.TP & " No. " & Co.Numero & vbCrLf
            If Len(Co.Beneficiario) > 1 Then Mensajes = Mensajes & "De " & ULCase(Co.Beneficiario)
            Titulo = "Pregunta de Modificacion"
            If BoxMensaje = vbYes Then
               ModificarComp = True
               CopiarComp = False
               NuevoComp = False
               Trans_No = 1
               FechaComp = MBFecha.Text
               Co.Fecha = MBFecha.Text
               Unload FListComprobantes
               BorrarAsientos True
               FComprobantes.Show
            End If
         End If
    Case "Anular_Comprobante"
         AnularComprobante
         ListarComprobantes
    Case "Autorizar_Comprobante"
    
    Case "Copiar_Comprobante"
         NumeroComp = NumComp
         Mensajes = "Seguro que quiere Copiar Comprobante " & Co.TP & " No. " & Co.Numero
         Titulo = "Pregunta de Eliminacion"
         If BoxMensaje = vbYes Then
            FechaComp = FechaSistema
            ModificarComp = False
            NuevoComp = False
            CopiarComp = True
            Trans_No = 1
            IniciarAsientosDe DGAsientos, AdoAsientos
            Unload FListComprobantes
            BorrarAsientos True
            FComprobantes.Show
         Else
            ListarComprobantes
         End If
    Case "Copiar_Item"
         ItemAux = InputBox("COPIAR A LA EMPRESA:", "COPIAR COMPROBANTE No. " & Co.TP & "-" & Co.Numero, "")
         If Len(ItemAux) = 3 And IsNumeric(ItemAux) Then
            Copiar_Comprobantes_Item ItemAux, "Comprobantes"
            Copiar_Comprobantes_Item ItemAux, "Transacciones"
            Copiar_Comprobantes_Item ItemAux, "Trans_Air"
            Copiar_Comprobantes_Item ItemAux, "Trans_Compras"
            Copiar_Comprobantes_Item ItemAux, "Trans_Ventas"
            Copiar_Comprobantes_Item ItemAux, "Trans_SubCtas"
            Copiar_Comprobantes_Item ItemAux, "Trans_Kardex"
            MsgBox "Proceso exitoso"
         End If
    Case "Anulacion_Masa"
         If ClaveContador Then
            RatonReloj
            FAnularMasa.Show 1
         End If
  End Select
End Sub

Public Sub Copiar_Comprobantes_Item(ItemA As String, Tabla As String)
    sSQL = "DELETE * " _
         & "FROM " & Tabla & " " _
         & "WHERE Item = '" & ItemA & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND TP = '" & Co.TP & "' " _
         & "AND Numero = '" & Co.Numero & "' "
    Ejecutar_SQL_SP sSQL
    
    sSQL = "SELECT " & Full_Fields(Tabla) & " " _
         & "FROM " & Tabla & " " _
         & "WHERE Item = '" & Co.Item & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND TP = '" & Co.TP & "' " _
         & "AND Numero = '" & Co.Numero & "' "
    Select_Adodc AdoAsientos, sSQL
    RatonReloj
    With AdoAsientos.Recordset
     If .RecordCount > 0 Then
         Do While Not .EOF
            SetAdoAddNew Tabla
            For K = 0 To .fields.Count - 1
                If .fields(K).Name = "Item" Then
                    SetAdoFields .fields(K).Name, ItemA
                Else
                    SetAdoFields .fields(K).Name, .fields(K)
                End If
            Next K
            SetAdoUpdate
           .MoveNext
         Loop
     End If
    End With
    RatonNormal
End Sub

