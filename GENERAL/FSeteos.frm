VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.5#0"; "comctl32.Ocx"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.Ocx"
Begin VB.Form FSeteos 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "SETEOS PRINCIPALES Y MANTENIMIENTO"
   ClientHeight    =   8175
   ClientLeft      =   -15
   ClientTop       =   210
   ClientWidth     =   14355
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8175
   ScaleWidth      =   14355
   ShowInTaskbar   =   0   'False
   Begin InetCtlsObjects.Inet URLinet 
      Left            =   14490
      Top             =   315
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.CommandButton Command12 
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
      Left            =   13020
      Picture         =   "FSeteos.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   525
      Width           =   1170
   End
   Begin TabDlg.SSTab TabSeteos 
      Height          =   7995
      Left            =   105
      TabIndex        =   3
      Top             =   105
      Width           =   14190
      _ExtentX        =   25030
      _ExtentY        =   14102
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      BackColor       =   16761024
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Se&teos Generales"
      TabPicture(0)   =   "FSeteos.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "DGEducativo"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "DGTipoPrest"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "DGCuentas"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "DGCodigos"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "&Duplicados/Mantenimiento"
      TabPicture(1)   =   "FSeteos.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "LstTablas"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "CheqSCAlum"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "LstDuplicados"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Command15"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "&Niveles de Seguridad"
      TabPicture(2)   =   "FSeteos.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Command4"
      Tab(2).Control(1)=   "Command1"
      Tab(2).Control(2)=   "LstEmpresas"
      Tab(2).Control(3)=   "Command6"
      Tab(2).Control(4)=   "LstModulos"
      Tab(2).Control(5)=   "Command10"
      Tab(2).Control(6)=   "DGEmp1"
      Tab(2).Control(7)=   "Command11"
      Tab(2).Control(8)=   "Command5"
      Tab(2).Control(9)=   "Command13"
      Tab(2).Control(10)=   "DCUsuario"
      Tab(2).Control(11)=   "Command7"
      Tab(2).Control(12)=   "TxtUsuario"
      Tab(2).Control(13)=   "Command3"
      Tab(2).Control(14)=   "Command2"
      Tab(2).Control(15)=   "TxtItem"
      Tab(2).Control(16)=   "Frame2"
      Tab(2).Control(17)=   "TextClave"
      Tab(2).Control(18)=   "MBPeriodo"
      Tab(2).Control(19)=   "DCBodega"
      Tab(2).Control(20)=   "Label3"
      Tab(2).Control(21)=   "Label5"
      Tab(2).Control(22)=   "Label2"
      Tab(2).Control(23)=   "Label1"
      Tab(2).Control(24)=   "Label10"
      Tab(2).Control(25)=   "Label4"
      Tab(2).ControlCount=   26
      TabCaption(3)   =   "&Impresiones"
      TabPicture(3)   =   "FSeteos.frx":035E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Command8"
      Tab(3).Control(1)=   "Command17(3)"
      Tab(3).Control(2)=   "Command17(2)"
      Tab(3).Control(3)=   "Command17(1)"
      Tab(3).Control(4)=   "Command21"
      Tab(3).Control(5)=   "DGSeteosPRN"
      Tab(3).Control(6)=   "PictFormatos"
      Tab(3).Control(7)=   "Command16"
      Tab(3).Control(8)=   "Command17(0)"
      Tab(3).Control(9)=   "DGFormato"
      Tab(3).ControlCount=   10
      Begin VB.CommandButton Command4 
         Caption         =   "Migracion a MySQL"
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
         Left            =   -62925
         Picture         =   "FSeteos.frx":037A
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   4305
         Width           =   1905
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Copiar Catalogos de Periodos"
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
         Left            =   -64920
         Picture         =   "FSeteos.frx":0A10
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   4305
         Width           =   1905
      End
      Begin VB.ListBox LstEmpresas 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4140
         Left            =   -72795
         Style           =   1  'Checkbox
         TabIndex        =   49
         Top             =   3570
         Width           =   7785
      End
      Begin VB.CommandButton Command6 
         Caption         =   "&Bloquear"
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
         Left            =   -63240
         Picture         =   "FSeteos.frx":10A6
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   420
         Width           =   1065
      End
      Begin VB.ListBox LstModulos 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4140
         Left            =   -74895
         Style           =   1  'Checkbox
         TabIndex        =   48
         Top             =   3570
         Width           =   2115
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Cerrar Facturacion"
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
         Left            =   -64920
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   3465
         Visible         =   0   'False
         Width           =   1905
      End
      Begin MSDataGridLib.DataGrid DGEmp1 
         Bindings        =   "FSeteos.frx":1970
         Height          =   2325
         Left            =   -71430
         TabIndex        =   17
         Top             =   1155
         Width           =   6315
         _ExtentX        =   11139
         _ExtentY        =   4101
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   18
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
            Name            =   "Arial Narrow"
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
      Begin VB.CommandButton Command15 
         Caption         =   "Procesar Lista Seleccionada"
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
         Left            =   -63450
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   420
         Width           =   1275
      End
      Begin VB.ListBox LstDuplicados 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7350
         Left            =   -72060
         Style           =   1  'Checkbox
         TabIndex        =   45
         Top             =   420
         Width           =   8520
      End
      Begin VB.CheckBox CheqSCAlum 
         Caption         =   "Sin Codigo de Alumnos"
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
         Left            =   -63450
         TabIndex        =   41
         Top             =   1260
         Width           =   1800
      End
      Begin VB.ListBox LstTablas 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7260
         Left            =   -74895
         TabIndex        =   42
         Top             =   420
         Width           =   2745
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Imprimir Fuentes"
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
         Left            =   -66810
         Picture         =   "FSeteos.frx":1986
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   420
         Width           =   2220
      End
      Begin VB.CommandButton Command11 
         Caption         =   "&Privilejios"
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
         Left            =   -64605
         Picture         =   "FSeteos.frx":1C90
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   525
         Width           =   1275
      End
      Begin VB.CommandButton Command17 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Index           =   3
         Left            =   -71220
         Picture         =   "FSeteos.frx":20D2
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   420
         Width           =   540
      End
      Begin VB.CommandButton Command17 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Index           =   2
         Left            =   -71850
         Picture         =   "FSeteos.frx":2514
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   420
         Width           =   540
      End
      Begin VB.CommandButton Command17 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Index           =   1
         Left            =   -72480
         Picture         =   "FSeteos.frx":2956
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   420
         Width           =   540
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Limpiar Bases de Datos"
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
         Height          =   750
         Left            =   -62925
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   3465
         Width           =   1905
      End
      Begin VB.CommandButton Command21 
         Caption         =   "Actualizar Impresiones"
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
         Left            =   -69120
         Picture         =   "FSeteos.frx":2D98
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   420
         Width           =   2220
      End
      Begin VB.CommandButton Command13 
         Caption         =   "Cerrar Educativo"
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
         Left            =   -62925
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   2625
         Visible         =   0   'False
         Width           =   1905
      End
      Begin MSDataListLib.DataCombo DCUsuario 
         Bindings        =   "FSeteos.frx":30A2
         DataSource      =   "AdoUsuario"
         Height          =   345
         Left            =   -74895
         TabIndex        =   5
         ToolTipText     =   "<Ctrl+Supr> Reactiva Usuario Todos los Privilejios"
         Top             =   735
         Width           =   9780
         _ExtentX        =   17251
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
      Begin MSDataGridLib.DataGrid DGSeteosPRN 
         Bindings        =   "FSeteos.frx":30BB
         Height          =   6525
         Left            =   -73110
         TabIndex        =   28
         Top             =   1260
         Width           =   12195
         _ExtentX        =   21511
         _ExtentY        =   11509
         _Version        =   393216
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
         Caption         =   "SETEOS DE CUENTAS EN ASIENTOS"
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
      Begin VB.PictureBox PictFormatos 
         BackColor       =   &H00FFFFFF&
         Height          =   1275
         Left            =   -73005
         ScaleHeight     =   2.143
         ScaleMode       =   7  'Centimeter
         ScaleWidth      =   9.181
         TabIndex        =   31
         Top             =   1575
         Visible         =   0   'False
         Width           =   5265
      End
      Begin VB.CommandButton Command16 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   -70590
         Picture         =   "FSeteos.frx":30D6
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   420
         Width           =   540
      End
      Begin VB.CommandButton Command17 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Index           =   0
         Left            =   -73110
         Picture         =   "FSeteos.frx":33E0
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   420
         Width           =   540
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Eliminar Periodo"
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
         Left            =   -64920
         Picture         =   "FSeteos.frx":3822
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   2625
         Width           =   1905
      End
      Begin VB.TextBox TxtUsuario 
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
         Left            =   -73530
         MaxLength       =   15
         TabIndex        =   25
         Top             =   1155
         Visible         =   0   'False
         Width           =   2010
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Cambiar Numero"
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
         Left            =   -64920
         Picture         =   "FSeteos.frx":40EC
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   1785
         Width           =   1905
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Cambiar Periodo"
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
         Left            =   -62925
         Picture         =   "FSeteos.frx":49B6
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   1785
         Width           =   1905
      End
      Begin VB.TextBox TxtItem 
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
         Left            =   -62085
         TabIndex        =   21
         Text            =   "000"
         Top             =   1260
         Width           =   645
      End
      Begin VB.Frame Frame2 
         Caption         =   "NIVELES DE SEGURIDAD"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1590
         Left            =   -74895
         TabIndex        =   9
         Top             =   1890
         Width           =   3375
         Begin VB.OptionButton OpcS 
            Caption         =   "S"
            Height          =   195
            Left            =   2730
            TabIndex        =   56
            Top             =   630
            Width           =   435
         End
         Begin VB.OptionButton OpcP 
            Caption         =   "P"
            Height          =   225
            Left            =   2730
            TabIndex        =   55
            Top             =   210
            Value           =   -1  'True
            Width           =   435
         End
         Begin VB.OptionButton OpcB 
            Caption         =   "B"
            Height          =   225
            Left            =   2730
            TabIndex        =   54
            Top             =   1050
            Width           =   435
         End
         Begin VB.CheckBox CheckNivel 
            Caption         =   "No. &7"
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
            Index           =   6
            Left            =   105
            TabIndex        =   39
            Top             =   1155
            Width           =   855
         End
         Begin VB.CheckBox CheckNivel 
            Caption         =   "No. &1"
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
            TabIndex        =   16
            Top             =   210
            Width           =   855
         End
         Begin VB.CheckBox CheckNivel 
            Caption         =   "No. &2"
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
            Left            =   1155
            TabIndex        =   15
            Top             =   210
            Width           =   855
         End
         Begin VB.CheckBox CheckNivel 
            Caption         =   "No. &3"
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
            TabIndex        =   14
            Top             =   525
            Width           =   855
         End
         Begin VB.CheckBox CheckNivel 
            Caption         =   "No. &4"
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
            Left            =   1155
            TabIndex        =   13
            Top             =   525
            Width           =   855
         End
         Begin VB.CheckBox CheckNivel 
            Caption         =   "No. &5"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   4
            Left            =   105
            TabIndex        =   12
            Top             =   840
            Width           =   855
         End
         Begin VB.CheckBox CheckNivel 
            Caption         =   "No. &6"
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
            Index           =   5
            Left            =   1155
            TabIndex        =   11
            Top             =   840
            Width           =   855
         End
         Begin VB.CheckBox CheckSupervisor 
            Caption         =   "Supervisor"
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
            Left            =   1155
            TabIndex        =   10
            Top             =   1155
            Width           =   1275
         End
      End
      Begin MSDataGridLib.DataGrid DGCodigos 
         Bindings        =   "FSeteos.frx":5280
         Height          =   5265
         Left            =   9345
         TabIndex        =   0
         Top             =   2520
         Width           =   4740
         _ExtentX        =   8361
         _ExtentY        =   9287
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   19
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
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "CODIGOS DE PROCESOS"
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
      Begin VB.TextBox TextClave 
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
         Left            =   -73530
         MaxLength       =   10
         TabIndex        =   6
         Top             =   1470
         Visible         =   0   'False
         Width           =   2010
      End
      Begin MSDataGridLib.DataGrid DGCuentas 
         Bindings        =   "FSeteos.frx":5299
         Height          =   5265
         Left            =   105
         TabIndex        =   1
         Top             =   2520
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   9287
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   19
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
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "SETEOS DE CUENTAS EN ASIENTOS"
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
      Begin MSMask.MaskEdBox MBPeriodo 
         Height          =   330
         Left            =   -64080
         TabIndex        =   19
         Top             =   1260
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
      Begin MSDataGridLib.DataGrid DGFormato 
         Bindings        =   "FSeteos.frx":52B2
         Height          =   7365
         Left            =   -74895
         TabIndex        =   34
         Top             =   420
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   12991
         _Version        =   393216
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
         Caption         =   "Formato Propio"
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
      Begin MSDataGridLib.DataGrid DGTipoPrest 
         Bindings        =   "FSeteos.frx":52CB
         Height          =   975
         Left            =   105
         TabIndex        =   43
         Top             =   420
         Width           =   12720
         _ExtentX        =   22437
         _ExtentY        =   1720
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   19
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
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "SETEOS DE CUENTAS EN PRESTAMOS"
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
      Begin MSDataGridLib.DataGrid DGEducativo 
         Bindings        =   "FSeteos.frx":52E6
         Height          =   1065
         Left            =   105
         TabIndex        =   44
         Top             =   1470
         Width           =   13980
         _ExtentX        =   24659
         _ExtentY        =   1879
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   19
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
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "SETEOS GENERALES DEL PLANTEL EDUCATIVO"
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
      Begin MSDataListLib.DataCombo DCBodega 
         Bindings        =   "FSeteos.frx":5301
         DataSource      =   "AdoBodega"
         Height          =   345
         Left            =   -64920
         TabIndex        =   50
         ToolTipText     =   "<Ctrl+Supr> Reactiva Usuario Todos los Privilejios"
         Top             =   5565
         Width           =   3900
         _ExtentX        =   6879
         _ExtentY        =   609
         _Version        =   393216
         Text            =   "BODEGAS"
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
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "BODEGA DE FACTURACION"
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
         Left            =   -64920
         TabIndex        =   51
         Top             =   5145
         Width           =   3900
      End
      Begin VB.Label Label5 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " USUARIO"
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
         TabIndex        =   26
         Top             =   1155
         Visible         =   0   'False
         Width           =   1380
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Item"
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
         Left            =   -62715
         TabIndex        =   20
         Top             =   1260
         Width           =   645
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Periodo"
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
         Left            =   -64920
         TabIndex        =   18
         Top             =   1260
         Width           =   855
      End
      Begin VB.Label Label10 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " CLAVE"
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
         TabIndex        =   8
         Top             =   1470
         Visible         =   0   'False
         Width           =   1380
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "NOMBRE DE &USUARIO"
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
         TabIndex        =   4
         Top             =   420
         Width           =   9780
      End
   End
   Begin MSAdodcLib.Adodc AdoFactura 
      Height          =   330
      Left            =   420
      Top             =   2940
      Visible         =   0   'False
      Width           =   2325
      _ExtentX        =   4101
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
   Begin MSAdodcLib.Adodc AdoFact 
      Height          =   330
      Left            =   420
      Top             =   2625
      Visible         =   0   'False
      Width           =   2325
      _ExtentX        =   4101
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
      Caption         =   "Fact"
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
      Left            =   420
      Top             =   2310
      Visible         =   0   'False
      Width           =   2325
      _ExtentX        =   4101
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
   Begin MSAdodcLib.Adodc AdoCajaCred 
      Height          =   330
      Left            =   420
      Top             =   1995
      Visible         =   0   'False
      Width           =   2325
      _ExtentX        =   4101
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
      Caption         =   "CajaCred"
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
   Begin MSAdodcLib.Adodc AdoBancos 
      Height          =   330
      Left            =   420
      Top             =   1680
      Visible         =   0   'False
      Width           =   2325
      _ExtentX        =   4101
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
      Caption         =   "Bancos"
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
   Begin MSAdodcLib.Adodc AdoTrans_GC 
      Height          =   330
      Left            =   420
      Top             =   3570
      Visible         =   0   'False
      Width           =   2325
      _ExtentX        =   4101
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
      Caption         =   "Trans_GC"
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
   Begin MSAdodcLib.Adodc AdoTrans_SC 
      Height          =   330
      Left            =   420
      Top             =   3255
      Visible         =   0   'False
      Width           =   2325
      _ExtentX        =   4101
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
      Caption         =   "Trans_SC"
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
      Left            =   420
      Top             =   1365
      Visible         =   0   'False
      Width           =   2325
      _ExtentX        =   4101
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
   Begin MSAdodcLib.Adodc AdoComp 
      Height          =   330
      Left            =   420
      Top             =   1050
      Visible         =   0   'False
      Width           =   2325
      _ExtentX        =   4101
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
   Begin MSAdodcLib.Adodc AdoEmp1 
      Height          =   330
      Left            =   420
      Top             =   735
      Visible         =   0   'False
      Width           =   2325
      _ExtentX        =   4101
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
      Caption         =   "Emp1"
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
   Begin MSAdodcLib.Adodc AdoEmp 
      Height          =   330
      Left            =   420
      Top             =   420
      Visible         =   0   'False
      Width           =   2325
      _ExtentX        =   4101
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
      Caption         =   "Emp"
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
   Begin MSAdodcLib.Adodc AdoCodigos 
      Height          =   330
      Left            =   420
      Top             =   3885
      Visible         =   0   'False
      Width           =   2325
      _ExtentX        =   4101
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
      Caption         =   "Codigos"
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
   Begin MSAdodcLib.Adodc AdoCuentas 
      Height          =   330
      Left            =   2835
      Top             =   420
      Visible         =   0   'False
      Width           =   2325
      _ExtentX        =   4101
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
      Caption         =   "Cuentas"
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
   Begin MSAdodcLib.Adodc AdoTipoPrest 
      Height          =   330
      Left            =   2835
      Top             =   735
      Visible         =   0   'False
      Width           =   2325
      _ExtentX        =   4101
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
      Caption         =   "TipoPrest"
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
      Left            =   2835
      Top             =   1050
      Visible         =   0   'False
      Width           =   2325
      _ExtentX        =   4101
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
   Begin MSAdodcLib.Adodc AdoInv 
      Height          =   330
      Left            =   2835
      Top             =   1365
      Visible         =   0   'False
      Width           =   2325
      _ExtentX        =   4101
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
   Begin MSAdodcLib.Adodc AdoSeteosPRN 
      Height          =   330
      Left            =   2835
      Top             =   1680
      Visible         =   0   'False
      Width           =   2325
      _ExtentX        =   4101
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
      Caption         =   "SeteosPRN"
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
   Begin MSAdodcLib.Adodc AdoSetPRN 
      Height          =   330
      Left            =   2835
      Top             =   1995
      Visible         =   0   'False
      Width           =   2325
      _ExtentX        =   4101
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
      Caption         =   "SetPRN"
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
   Begin MSAdodcLib.Adodc AdoFormato 
      Height          =   330
      Left            =   2835
      Top             =   2310
      Visible         =   0   'False
      Width           =   2325
      _ExtentX        =   4101
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
      Caption         =   "Formato"
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
      Left            =   2835
      Top             =   2625
      Visible         =   0   'False
      Width           =   2325
      _ExtentX        =   4101
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
   Begin MSAdodcLib.Adodc AdoPaises 
      Height          =   330
      Left            =   2835
      Top             =   2940
      Visible         =   0   'False
      Width           =   2325
      _ExtentX        =   4101
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
      Caption         =   "Paises"
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
   Begin MSAdodcLib.Adodc AdoAutorizacion 
      Height          =   330
      Left            =   2835
      Top             =   3255
      Visible         =   0   'False
      Width           =   2325
      _ExtentX        =   4101
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
      Caption         =   "Autorizacion"
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
   Begin MSAdodcLib.Adodc AdoEducativo 
      Height          =   330
      Left            =   2835
      Top             =   3570
      Visible         =   0   'False
      Width           =   2325
      _ExtentX        =   4101
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
      Caption         =   "Educativo"
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
      Left            =   525
      Top             =   5040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin MSAdodcLib.Adodc AdoModulos 
      Height          =   330
      Left            =   2835
      Top             =   3885
      Visible         =   0   'False
      Width           =   2325
      _ExtentX        =   4101
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
      Caption         =   "Modulos"
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
   Begin MSAdodcLib.Adodc AdoPrueba 
      Height          =   330
      Left            =   2835
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
      Caption         =   "Prueba"
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
      Left            =   420
      Top             =   4200
      Visible         =   0   'False
      Width           =   2325
      _ExtentX        =   4101
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
   Begin ComctlLib.ImageList ImageList1 
      Left            =   1155
      Top             =   4935
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   2
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FSeteos.frx":5319
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FSeteos.frx":54F3
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FSeteos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim CtasProc() As CtasAsiento
Dim ContCtas As Integer
Dim AutorizaOld As String

Private Sub Command1_Click()
  If ClaveSupervisor Then
     RatonReloj
     Si_No = True
     FCopyCat.Show 1
  End If
End Sub
'
Private Sub Command10_Click()
Dim CamposF As String
Dim CamposDF As String
Dim CamposTA As String
Dim CamposCL As String
Dim CamposFF As String
  
  FechaValida MBPeriodo
  FechaFin = MBPeriodo
  If ClaveContador Then
     Titulo = "CIERRE DEL PERIODO DE FACTURACION"
     Mensajes = "Seguro de realizar el cierre del periodo de Facturacion del: " & FechaFin
     If BoxMensaje = vbYes Then
        RatonReloj
        Progreso_Barra.Valor_Maximo = 100
        Progreso_Barra.Mensaje_Box = "CERRANDO EL PERIODO DEL: " & FechaFin
        Progreso_Iniciar
        FA.Factura = 0
        FA.Fecha_Corte = MBPeriodo 'FechaSistema
        FA.Fecha_Desde = "01/01/2000"
        FA.Fecha_Hasta = FA.Fecha_Corte
        
        Actualizar_Abonos_Facturas_SP FA
        Progreso_Esperar
    
        sSQL = "UPDATE Facturas " _
             & "SET X = '.' " _
             & "WHERE Item = '" & NumEmpresa & "' " _
             & "AND Periodo = '" & Periodo_Contable & "' "
        Ejecutar_SQL_SP sSQL
        Progreso_Esperar
        
        sSQL = "UPDATE Detalle_Factura " _
             & "SET X = '.' " _
             & "WHERE Item = '" & NumEmpresa & "' " _
             & "AND Periodo = '" & Periodo_Contable & "' "
        Ejecutar_SQL_SP sSQL
        Progreso_Esperar

        sSQL = "UPDATE Trans_Abonos " _
             & "SET X = '.' " _
             & "WHERE Item = '" & NumEmpresa & "' " _
             & "AND Periodo = '" & Periodo_Contable & "' "
        Ejecutar_SQL_SP sSQL
        Progreso_Esperar
        
        sSQL = "UPDATE Facturas " _
             & "SET X = 'C' " _
             & "WHERE Item = '" & NumEmpresa & "' " _
             & "AND Periodo = '" & Periodo_Contable & "' " _
             & "AND Fecha <= #" & BuscarFecha(MBPeriodo) & "# " _
             & "AND Saldo_Actual = 0 "
        Ejecutar_SQL_SP sSQL
        Progreso_Esperar
        
        sSQL = "UPDATE Facturas " _
             & "SET X = 'C' " _
             & "WHERE Item = '" & NumEmpresa & "' " _
             & "AND Periodo = '" & Periodo_Contable & "' " _
             & "AND Fecha <= #" & BuscarFecha(MBPeriodo) & "# " _
             & "AND T = 'A' "
        Ejecutar_SQL_SP sSQL
        Progreso_Esperar
        
        If SQL_Server Then
           sSQL = "UPDATE Detalle_Factura " _
                & "SET X = F.X " _
                & "FROM Detalle_Factura As FX,Facturas As F "
        Else
           sSQL = "UPDATE Detalle_Factura As FX,Facturas As F " _
                & "SET FX.X = F.X "
        End If
        sSQL = sSQL _
             & "WHERE F.Item = '" & NumEmpresa & "' " _
             & "AND F.Periodo = '" & Periodo_Contable & "' " _
             & "AND FX.Item = F.Item " _
             & "AND FX.Periodo = F.Periodo " _
             & "AND FX.TC = F.TC " _
             & "AND FX.Serie = F.Serie " _
             & "AND FX.Autorizacion = F.Autorizacion " _
             & "AND FX.Factura = F.Factura " _
             & "AND FX.CodigoC = F.CodigoC "
        Ejecutar_SQL_SP sSQL
        Progreso_Esperar

        If SQL_Server Then
           sSQL = "UPDATE Trans_Abonos " _
                & "SET X = F.X " _
                & "FROM Trans_Abonos As FX,Facturas As F "
        Else
           sSQL = "UPDATE Trans_Abonos As FX,Facturas As F " _
                & "SET FX.X = F.X "
        End If
        sSQL = sSQL _
             & "WHERE F.Item = '" & NumEmpresa & "' " _
             & "AND F.Periodo = '" & Periodo_Contable & "' " _
             & "AND FX.Item = F.Item " _
             & "AND FX.Periodo = F.Periodo " _
             & "AND FX.TP = F.TC " _
             & "AND FX.Serie = F.Serie " _
             & "AND FX.Autorizacion = F.Autorizacion " _
             & "AND FX.Factura = F.Factura " _
             & "AND FX.CodigoC = F.CodigoC "
        Ejecutar_SQL_SP sSQL
        Progreso_Esperar
        
        sSQL = "UPDATE Facturas " _
             & "SET Periodo = '" & MBPeriodo & "' " _
             & "WHERE Item = '" & NumEmpresa & "' " _
             & "AND Periodo = '" & Periodo_Contable & "' " _
             & "AND X = 'C' "
        Ejecutar_SQL_SP sSQL
        Progreso_Esperar
        
        sSQL = "UPDATE Detalle_Factura " _
             & "SET Periodo = '" & MBPeriodo & "' " _
             & "WHERE Item = '" & NumEmpresa & "' " _
             & "AND Periodo = '" & Periodo_Contable & "' " _
             & "AND X = 'C' "
        Ejecutar_SQL_SP sSQL
        Progreso_Esperar
        
        sSQL = "UPDATE Trans_Abonos " _
             & "SET Periodo = '" & MBPeriodo & "' " _
             & "WHERE Item = '" & NumEmpresa & "' " _
             & "AND Periodo = '" & Periodo_Contable & "' " _
             & "AND X = 'C' "
        Ejecutar_SQL_SP sSQL
        Progreso_Esperar
        
        sSQL = "UPDATE Trans_Abonos " _
             & "SET Periodo = '" & MBPeriodo & "' " _
             & "WHERE Item = '" & NumEmpresa & "' " _
             & "AND Periodo = '" & Periodo_Contable & "' " _
             & "AND Fecha <= #" & BuscarFecha(MBPeriodo) & "# " _
             & "AND TP = 'TJ' "
        Ejecutar_SQL_SP sSQL
        Progreso_Esperar
        
        Total = 0
        
        sSQL = "UPDATE Facturas " _
             & "SET X = 'P' " _
             & "WHERE Item = '" & NumEmpresa & "' " _
             & "AND Periodo = '" & Periodo_Contable & "' " _
             & "AND Fecha <= #" & BuscarFecha(MBPeriodo) & "# " _
             & "AND X <> 'C' "
        Ejecutar_SQL_SP sSQL
        Progreso_Esperar
        
        sSQL = "UPDATE Detalle_Factura " _
             & "SET X = 'P' " _
             & "FROM Detalle_Factura As DF, Facturas As F " _
             & "WHERE F.Item = '" & NumEmpresa & "' " _
             & "AND F.Periodo = '" & Periodo_Contable & "' " _
             & "AND F.X = 'P' " _
             & "AND DF.Item = F.Item " _
             & "AND DF.Periodo = F.Periodo " _
             & "AND DF.TC = F.TC " _
             & "AND DF.Serie = F.Serie " _
             & "AND DF.Factura = F.Factura " _
             & "AND DF.Autorizacion = F.Autorizacion "
        Ejecutar_SQL_SP sSQL
        Progreso_Esperar
        
        sSQL = "UPDATE Trans_Abonos " _
             & "SET X = 'P' " _
             & "FROM Trans_Abonos As TA, Facturas As F " _
             & "WHERE F.Item = '" & NumEmpresa & "' " _
             & "AND F.Periodo = '" & Periodo_Contable & "' " _
             & "AND F.X = 'P' " _
             & "AND TA.Fecha <= #" & BuscarFecha(MBPeriodo) & "# " _
             & "AND TA.Item = F.Item " _
             & "AND TA.Periodo = F.Periodo " _
             & "AND TA.TP = F.TC " _
             & "AND TA.Serie = F.Serie " _
             & "AND TA.Factura = F.Factura " _
             & "AND TA.Autorizacion = F.Autorizacion "
        Ejecutar_SQL_SP sSQL
        Progreso_Esperar
        
        CamposF = ""
        CamposDF = ""
        CamposTA = ""
        CamposCL = ""
        CamposFF = ""
        
        sSQL = "SELECT * " _
             & "FROM Catalogo_Lineas " _
             & "WHERE 1 = 0 "
        Select_Adodc AdoPrueba, sSQL
        With AdoPrueba.Recordset
        For I = 0 To .fields.Count - 1
            Select Case .fields(I).Name
              Case "Periodo", "ID" 'nada
              Case Else: CamposCL = CamposCL & .fields(I).Name & ","
            End Select
        Next I
        End With
        Progreso_Esperar
                
        sSQL = "SELECT * " _
             & "FROM Facturas_Formatos " _
             & "WHERE 1 = 0 "
        Select_Adodc AdoPrueba, sSQL
        With AdoPrueba.Recordset
        For I = 0 To .fields.Count - 1
            Select Case .fields(I).Name
              Case "Periodo", "ID" 'nada
              Case Else: CamposFF = CamposFF & .fields(I).Name & ","
            End Select
        Next I
        End With
        Progreso_Esperar
        
        sSQL = "SELECT * " _
             & "FROM Facturas " _
             & "WHERE 1 = 0 "
        Select_Adodc AdoPrueba, sSQL
        With AdoPrueba.Recordset
        For I = 0 To .fields.Count - 1
            Select Case .fields(I).Name
              Case "Periodo", "ID" 'nada
              Case Else: CamposF = CamposF & .fields(I).Name & ","
            End Select
        Next I
        End With
        Progreso_Esperar
        
        sSQL = "SELECT * " _
             & "FROM Detalle_Factura " _
             & "WHERE 1 = 0 "
        Select_Adodc AdoPrueba, sSQL
        With AdoPrueba.Recordset
        For I = 0 To .fields.Count - 1
            Select Case .fields(I).Name
              Case "Periodo", "ID" 'nada
              Case Else: CamposDF = CamposDF & .fields(I).Name & ","
            End Select
        Next I
        End With
        Progreso_Esperar
        
        sSQL = "SELECT * " _
             & "FROM Trans_Abonos " _
             & "WHERE 1 = 0 "
        Select_Adodc AdoPrueba, sSQL
        With AdoPrueba.Recordset
        For I = 0 To .fields.Count - 1
            Select Case .fields(I).Name
              Case "Periodo", "ID" 'nada
              Case Else: CamposTA = CamposTA & .fields(I).Name & ","
            End Select
        Next I
        End With
        Progreso_Esperar
        
        sSQL = "INSERT INTO Facturas (" & CamposF & "Periodo) " _
             & "SELECT " & CamposF & "'" & MBPeriodo & "' " _
             & "FROM Facturas " _
             & "WHERE Item = '" & NumEmpresa & "' " _
             & "AND Periodo = '" & Periodo_Contable & "' " _
             & "AND X = 'P' " _
             & "ORDER BY TC,Serie,Factura "
        Ejecutar_SQL_SP sSQL
        Progreso_Esperar
        
        sSQL = "INSERT INTO Detalle_Factura (" & CamposDF & "Periodo) " _
             & "SELECT " & CamposDF & "'" & MBPeriodo & "' " _
             & "FROM Detalle_Factura " _
             & "WHERE Item = '" & NumEmpresa & "' " _
             & "AND Periodo = '" & Periodo_Contable & "' " _
             & "AND X = 'P' " _
             & "ORDER BY TC,Serie,Factura "
        Ejecutar_SQL_SP sSQL
        Progreso_Esperar
        
        sSQL = "INSERT INTO Trans_Abonos (" & CamposTA & "Periodo) " _
             & "SELECT " & CamposTA & "'" & MBPeriodo & "' " _
             & "FROM Trans_Abonos " _
             & "WHERE Item = '" & NumEmpresa & "' " _
             & "AND Periodo = '" & Periodo_Contable & "' " _
             & "AND X = 'P' " _
             & "ORDER BY TP,Serie,Factura "
        Ejecutar_SQL_SP sSQL
        Progreso_Esperar
        
        sSQL = "INSERT INTO Catalogo_Lineas (" & CamposCL & "Periodo) " _
             & "SELECT " & CamposCL & "'" & MBPeriodo & "' " _
             & "FROM Catalogo_Lineas " _
             & "WHERE Item = '" & NumEmpresa & "' " _
             & "AND Periodo = '" & Periodo_Contable & "' " _
             & "ORDER BY Fact,Serie,Secuencial "
        Ejecutar_SQL_SP sSQL
        Progreso_Esperar
        
        sSQL = "INSERT INTO Facturas_Formatos (" & CamposFF & "Periodo) " _
             & "SELECT " & CamposFF & "'" & MBPeriodo & "' " _
             & "FROM Facturas_Formatos " _
             & "WHERE Item = '" & NumEmpresa & "' " _
             & "AND Periodo = '" & Periodo_Contable & "' " _
             & "ORDER BY TC,Serie,Fecha_Inicio "
        Ejecutar_SQL_SP sSQL
        Progreso_Esperar
        
        Actualizar_Abonos_Facturas_SP FA
        Progreso_Esperar
        
'''        MsgBox "..................."
'''
'''        sSQL = "SELECT * " _
'''             & "FROM Facturas " _
'''             & "WHERE Item = '" & NumEmpresa & "' " _
'''             & "AND Periodo = '" & Periodo_Contable & "' " _
'''             & "AND Fecha <= #" & BuscarFecha(MBPeriodo) & "# " _
'''             & "AND X <> 'C' " _
'''             & "ORDER BY TC,Serie,Factura "
'''        Select_Adodc AdoPrueba, sSQL
'''        With AdoPrueba.Recordset
'''         If .RecordCount > 0 Then
'''             Progreso_Barra.Valor_Maximo = Progreso_Barra.Valor_Maximo + .RecordCount
'''             Do While Not .EOF
'''                RatonReloj
'''                Progreso_Barra.Mensaje_Box = "CERRANDO FACTURACION DEL DOCUMENTO " & .Fields("Factura")
'''                Progreso_Esperar
'''
'''               'Cerramos las facturas pendientes
'''                sSQL = "UPDATE Facturas " _
'''                     & "SET Periodo = '" & MBPeriodo & "' " _
'''                     & "WHERE Item = '" & NumEmpresa & "' " _
'''                     & "AND Periodo = '" & Periodo_Contable & "' " _
'''                     & "AND TC = '" & .Fields("TC") & "' " _
'''                     & "AND Serie = '" & .Fields("Serie") & "' " _
'''                     & "AND Autorizacion = '" & .Fields("Autorizacion") & "' " _
'''                     & "AND Factura = " & .Fields("Factura") & " " _
'''                     & "AND CodigoC = '" & .Fields("CodigoC") & "' "
'''                Ejecutar_SQL_SP sSQL
'''
'''                SetAdoAddNew "Facturas"
'''                For I = 0 To AdoPrueba.Recordset.Fields.Count - 1
'''                    If AdoPrueba.Recordset.Fields(I).Name = "Observacion" Then
'''                       SetAdoFields AdoPrueba.Recordset.Fields(I).Name, "[" & MBPeriodo & "] " & AdoPrueba.Recordset.Fields(I)
'''                    Else
'''                       SetAdoFields AdoPrueba.Recordset.Fields(I).Name, AdoPrueba.Recordset.Fields(I)
'''                    End If
'''                Next I
'''                SetAdoUpdate
'''
'''                sSQL = "SELECT * " _
'''                     & "FROM Detalle_Factura " _
'''                     & "WHERE Item = '" & NumEmpresa & "' " _
'''                     & "AND Periodo = '" & Periodo_Contable & "' " _
'''                     & "AND TC = '" & .Fields("TC") & "' " _
'''                     & "AND Serie = '" & .Fields("Serie") & "' " _
'''                     & "AND Autorizacion = '" & .Fields("Autorizacion") & "' " _
'''                     & "AND Factura = " & .Fields("Factura") & " " _
'''                     & "AND CodigoC = '" & .Fields("CodigoC") & "' "
'''                Select_Adodc AdoFact, sSQL
'''                If AdoFact.Recordset.RecordCount > 0 Then
'''                   sSQL = "UPDATE Detalle_Factura " _
'''                        & "SET Periodo = '" & MBPeriodo & "' " _
'''                        & "WHERE Item = '" & NumEmpresa & "' " _
'''                        & "AND Periodo = '" & Periodo_Contable & "' " _
'''                        & "AND TC = '" & .Fields("TC") & "' " _
'''                        & "AND Serie = '" & .Fields("Serie") & "' " _
'''                        & "AND Autorizacion = '" & .Fields("Autorizacion") & "' " _
'''                        & "AND Factura = " & .Fields("Factura") & " " _
'''                        & "AND CodigoC = '" & .Fields("CodigoC") & "' "
'''                   Ejecutar_SQL_SP sSQL
'''                   Do While Not AdoFact.Recordset.EOF
'''                      SetAdoAddNew "Detalle_Factura"
'''                      For I = 0 To AdoFact.Recordset.Fields.Count - 1
'''                          If AdoFact.Recordset.Fields(I).Name = "Ruta" Then
'''                             SetAdoFields AdoFact.Recordset.Fields(I).Name, "[" & MBPeriodo & "] " & AdoFact.Recordset.Fields(I)
'''                          Else
'''                             SetAdoFields AdoFact.Recordset.Fields(I).Name, AdoFact.Recordset.Fields(I)
'''                          End If
'''                      Next I
'''                      SetAdoUpdate
'''                      AdoFact.Recordset.MoveNext
'''                   Loop
'''                End If
'''
'''                sSQL = "SELECT * " _
'''                     & "FROM Trans_Abonos " _
'''                     & "WHERE Item = '" & NumEmpresa & "' " _
'''                     & "AND Periodo = '" & Periodo_Contable & "' " _
'''                     & "AND Fecha <= #" & BuscarFecha(MBPeriodo) & "# " _
'''                     & "AND TP = '" & .Fields("TC") & "' " _
'''                     & "AND Serie = '" & .Fields("Serie") & "' " _
'''                     & "AND Autorizacion = '" & .Fields("Autorizacion") & "' " _
'''                     & "AND Factura = " & .Fields("Factura") & " " _
'''                     & "AND CodigoC = '" & .Fields("CodigoC") & "' "
'''                Select_Adodc AdoFact, sSQL
'''                If AdoFact.Recordset.RecordCount > 0 Then
'''                   sSQL = "UPDATE Trans_Abonos " _
'''                        & "SET Periodo = '" & MBPeriodo & "' " _
'''                        & "WHERE Item = '" & NumEmpresa & "' " _
'''                        & "AND Periodo = '" & Periodo_Contable & "' " _
'''                        & "AND Fecha <= #" & BuscarFecha(MBPeriodo) & "# " _
'''                        & "AND TP = '" & .Fields("TC") & "' " _
'''                        & "AND Serie = '" & .Fields("Serie") & "' " _
'''                        & "AND Autorizacion = '" & .Fields("Autorizacion") & "' " _
'''                        & "AND Factura = " & .Fields("Factura") & " " _
'''                        & "AND CodigoC = '" & .Fields("CodigoC") & "' "
'''                   Ejecutar_SQL_SP sSQL
'''                   Do While Not AdoFact.Recordset.EOF
'''                      SetAdoAddNew "Trans_Abonos"
'''                      For I = 0 To AdoFact.Recordset.Fields.Count - 1
'''                          If AdoFact.Recordset.Fields(I).Name = "Comprobante" Then
'''                             SetAdoFields AdoFact.Recordset.Fields(I).Name, "[" & MBPeriodo & "] " & AdoFact.Recordset.Fields(I)
'''                          Else
'''                             SetAdoFields AdoFact.Recordset.Fields(I).Name, AdoFact.Recordset.Fields(I)
'''                          End If
'''                      Next I
'''                      SetAdoUpdate
'''                      AdoFact.Recordset.MoveNext
'''                   Loop
'''                End If
'''                Total = Total + .Fields("Saldo_Actual")
'''               .MoveNext
'''             Loop
'''         End If
'''        End With
        Progreso_Barra.Incremento = Progreso_Barra.Valor_Maximo
        Progreso_Barra.Mensaje_Box = "CERRANDO EL PERIODO DE FACTURACION DEL: " & FechaFin
        Progreso_Esperar
     End If
  End If
  RatonNormal
  Progreso_Final
  MsgBox "Proceso Terminado"
End Sub

'Procesar Lista seleccionada
Private Sub Command15_Click()
Dim IdProc As Byte
  If LstDuplicados.ListCount > 1 Then
  For IdProc = 0 To LstDuplicados.ListCount - 1
      If LstDuplicados.Selected(IdProc) Then
        'MsgBox LstDuplicados.List(IdProc)
         Select Case LstDuplicados.List(IdProc)
           Case "Duplicados de Clientes"
                Procesar_Duplicados_Clientes
           Case "Duplicados de Usuarios"
                Procesar_Duplicados_Usuarios
           Case "Duplicados de Representantes"
                Procesar_Duplicados_Representantes
           Case "Duplicados de Cuentas, SubModulos y Productos"
                Procesar_Duplicados_Catalogo_SubModulos
           Case "Duplicados de Rubros Clientes Facturacion"
                Procesar_Duplicados_Rubros_Clientes_Facturacion
           Case "Duplicados de Compras y Retenciones"
                Procesar_Duplicados_Compras_Retenciones
           Case "Duplicados de Clientes Facturacion"
                Procesar_Duplicados_Clientes_Facturacion
           Case "Actualizacion de CI/RUC Clientes Facturacion"
                Procesar_Actualizar_CI_AT
           Case "Actualizacion de CI/RUC de Garantes"
                Procesar_Actualizar_CI_Garantes
           Case "Actualizacion de Beneficiario en las bases"
                Actualizar_Codigo_Clientes_Bases
           Case "Actualizacion de Codigos Superiores en Productos"
                Actualizar_Codigo_Superiores
           Case "Actualizacion de Documentos Electronicos Autorizados"
                If CFechaLong(FechaSistema) <= CFechaLong(Fecha_CE) Then
                   Actualizar_Facturas_Electronicas
                   Actualizar_Retenciones_Electronicas
                Else
                   RatonNormal
                   MsgBox MensajeNoAutorizarCE
                End If
           Case "Actualizacion Facturas y Notas de Credito en Kardex"
                Actualizar_Facturas_NC_en_Kardex
           Case "Abonar Facturas Pendientes"
                Abonar_Facturas_Pendientes
           Case "Cambiar PV as FA"
                Procesar_Cambio_PV_FA
           Case "Renumerar CI/RUC/Pasaporte de Personas"
                Procesar_Renumerar_CIRUC
           Case "Renumerar Clave del Catalogo"
                Procesar_Renumerar_Claves
           Case "Listar Movimientos con Cuentas de Grupo"
                Procesar_Lista_Ctas_Grupo
           Case "Respaldar Bases de Datos Completa"
                Respaldar_Base_Total
           Case "Borrar Datos de Estudiantes"
                 Borrar_Datos_Estudiantes
           Case "Eliminar Basura en Base de Datos"
                Procesar_Limpiar_Basura
           Case "Eliminar Seteos por Default (000)"
                Procesar_Limpiar_Basura_000
           Case "Eliminar Indice en Base de Datos"
                Procesar_Indices_Base_Datos
           Case "Eliminar Tablas Vacias"
                Eliminar_Tabla_Vacias
           Case "Generar Documentos Electronicos"
                Generar_Documentos_Electronicos
           Case "Prueba de Envio de Correos"
                Prueba_Envio_de_Correos
           Case "Actualizar Avreviaturas Accesos Usuarios"
                Poner_Avreviatura_Accesos
                
           Case "Realizar Copia de Actualizacion"
                Procesar_Update_DB
         End Select
      End If
  Next IdProc
  End If
  RatonNormal
  MsgBox "FIN DE LOS PROCESOS DE LA LISTA SELECCIONADA"
  'Unload FSeteos
End Sub

Private Sub Command4_Click()
  Unload FSeteos
  FMigracion.Show
End Sub

Private Sub Command8_Click()
Dim Fuentes(21) As String
Dim Fuente As String
Dim Tamanio As Byte
Dim IFond As Byte
On Error GoTo Errorhandler
Mensajes = "Seguro de Imprimir en:" & vbCrLf & Printer.DeviceName & "?"
Titulo = "IMPRESION"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
   Escala_Centimetro 2, TipoTimes, 7
   RatonReloj
   Fuentes(0) = TipoArial
   Fuentes(1) = TipoArialNarrow
   Fuentes(2) = TipoArialBlack
   Fuentes(3) = TipoArialUnicode
   Fuentes(4) = TipoAvantGarde
   Fuentes(5) = TipoSerif
   Fuentes(6) = TipoSansSerif
   Fuentes(7) = TipoCondensed
   Fuentes(8) = TipoComicSans
   Fuentes(9) = TipoConsola
   Fuentes(10) = TipoCourier
   Fuentes(11) = TipoCourierNew
   Fuentes(12) = TipoTimes
   Fuentes(13) = TipoTerminal
   Fuentes(14) = TipoSystem
   Fuentes(15) = TipoHelvetica
   Fuentes(16) = TipoHelveticaBold
   Fuentes(17) = TipoTahoma
   Fuentes(18) = TipoVerdana
   Fuentes(19) = TipoWingdings
   Fuente = ""
   For I = 65 To 122
       Fuente = Fuente & Chr(I)
   Next I
   For IFond = 0 To 20
       PosLinea = 1
       Printer.FontSize = 8
       Printer.FontBold = False
       Printer.FontItalic = False
       Printer.FontUnderline = False
       Printer.FontName = Fuentes(0)
       PrinterTexto 1, PosLinea, Fuentes(IFond)
       PosLinea = PosLinea + Printer.TextHeight("H") + 0.05
       Printer.FontName = Fuentes(IFond)
       For Tamanio = 7 To 24
           Printer.FontSize = Tamanio
           Printer.FontBold = False
           Printer.FontItalic = False
           Printer.FontUnderline = False
           PrinterTexto 1, PosLinea, CStr(Tamanio)
           PrinterTexto 2, PosLinea, Fuente
           PrinterTexto 16, PosLinea, "1,234,567,890.00"
           PosLinea = PosLinea + Printer.TextHeight("H")
           Printer.FontBold = False
           Printer.FontItalic = True
           Printer.FontUnderline = False
           PrinterTexto 1, PosLinea, CStr(Tamanio)
           PrinterTexto 2, PosLinea, Fuente
           PrinterTexto 16, PosLinea, "1,234,567,890.00"
           PosLinea = PosLinea + Printer.TextHeight("H")
           Printer.FontBold = False
           Printer.FontItalic = False
           Printer.FontUnderline = True
           PrinterTexto 1, PosLinea, CStr(Tamanio)
           PrinterTexto 2, PosLinea, Fuente
           PrinterTexto 16, PosLinea, "1,234,567,890.00"
           PosLinea = PosLinea + Printer.TextHeight("H")
           Printer.FontBold = True
           Printer.FontItalic = False
           Printer.FontUnderline = False
           PrinterTexto 1, PosLinea, CStr(Tamanio)
           PrinterTexto 2, PosLinea, Fuente
           PrinterTexto 16, PosLinea, "1,234,567,890.00"
           PosLinea = PosLinea + Printer.TextHeight("H")
       Next Tamanio
       Printer.NewPage
   Next IFond
RatonNormal
MensajeEncabData = ""
Printer.EndDoc
RatonNormal
MsgBox "Proceso terminado"
Exit Sub
Errorhandler:
    RatonNormal
    ErrorDeImpresion
    Exit Sub
Else
    RatonNormal
End If
End Sub

Public Sub Procesar_Actualizar_CI_AT()
  RatonReloj
  
  sSQL = "DELETE * " _
       & "FROM Clientes " _
       & "WHERE Grupo = 'REPRESEN' "
 'Ejecutar_SQL_SP sSQL
  RatonReloj
  sSQL = "UPDATE Clientes_Matriculas " _
       & "SET Cedula_R = '.' " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND Cedula_R = '' "
  Ejecutar_SQL_SP sSQL
  RatonReloj
  sSQL = "SELECT Item,Periodo,Codigo,Cedula_R,TB " _
       & "FROM Clientes_Matriculas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "ORDER BY Grupo_No,Cedula_R "
  Select_Adodc AdoAux, sSQL
  With AdoAux.Recordset
   If .RecordCount > 0 Then
       RatonReloj
       Progreso_Barra.Incremento = 0
       Progreso_Barra.Valor_Maximo = .RecordCount
       Do While Not .EOF
          Progreso_Barra.Mensaje_Box = "Proceso de Alumnos: "
          Progreso_Esperar
          CICliente = Replace(.fields("Cedula_R"), "-", "")
          CICliente = Replace(CICliente, "|", "")
          DigVerif = Digito_Verificador(CICliente)
         .fields("TB") = Tipo_RUC_CI.Tipo_Beneficiario
         .fields("Cedula_R") = CICliente
         .Update
         .MoveNext
       Loop
   End If
  End With
''  RatonReloj
''  sSQL = "SELECT * " _
''       & "FROM Clientes " _
''       & "WHERE Codigo <> '.' " _
''       & "ORDER BY CI_RUC,Grupo "
''  Select_Adodc AdoAux, sSQL
''  With AdoAux.Recordset
''   If .RecordCount > 0 Then
''       RatonReloj
''       Progreso_Barra.Incremento = 0
''       Progreso_Barra.Valor_Maximo = .RecordCount
''       Do While Not .EOF
''          Progreso_Barra.Mensaje_Box = "Proceso de Clientes/Proveedores: "
''          progreso_esperar
''          CICliente = Replace(.Fields("CI_RUC"), "-", "")
''          CICliente = Replace(CICliente, "|", "")
''          DigVerif = Digito_Verificador(CICliente)
''         .Fields("TD") = Tipo_RUC_CI.Tipo_Beneficiario
''          If CICliente = Ninguno Then
''            .Fields("TD") = "O"
''            '.Fields("CI_RUC") = Ninguno
''          Else
''
''            '.Fields("CI_RUC") = CICliente
''          End If
''         .Update
''         .MoveNext
''       Loop
''   End If
''  End With
''  sSQL = "SELECT C.FA,C.Cliente As Alumnos,CM.Representante,CM.Cedula_R As CI_RUC,C.Grupo,C.Direccion,C.Telefono,CM.Telefono_R,CM.TB " _
''       & "FROM Clientes As C, Clientes_Matriculas As CM " _
''       & "WHERE CM.Item = '" & NumEmpresa & "' " _
''       & "AND CM.Periodo = '" & Periodo_Contable & "' " _
''       & "AND CM.TB NOT IN ('C','R') " _
''       & "AND C.Codigo = CM.Codigo " _
''       & "UNION " _
''       & "SELECT FA,Cliente As Alumnos,'Revize en Clientes o Proveedores' As Representante,CI_RUC,Grupo,Direccion,Telefono,'0000000000' As Telefono_R,TD AS TB " _
''       & "FROM Clientes " _
''       & "WHERE TD NOT IN ('C','R') " _
''       & "AND FA = " & Val(adFalse) & " " _
''       & "ORDER BY FA,Grupo,Cliente "
''  Select_Adodc_Grid DGAux, AdoAux, sSQL
  RatonNormal
  
  MsgBox "Proceso Terminado"
End Sub

Public Sub Procesar_Actualizar_CI_Garantes()
  RatonReloj
  
  RatonReloj
  sSQL = "SELECT * " _
       & "FROM Clientes_Datos_Extras " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Beneficiario <> '.' " _
       & "ORDER BY CI,Beneficiario "
  Select_Adodc AdoAux, sSQL
  With AdoAux.Recordset
   If .RecordCount > 0 Then
       RatonReloj
       Progreso_Barra.Incremento = 0
       Progreso_Barra.Valor_Maximo = .RecordCount
       Do While Not .EOF
          Progreso_Barra.Mensaje_Box = "Proceso del Garante: "
          Progreso_Esperar
          CICliente = Replace(.fields("CI"), "-", "")
          CICliente = Replace(CICliente, "|", "")
          DigVerif = Digito_Verificador(CICliente)
         .fields("Acreditacion") = Tipo_RUC_CI.Tipo_Beneficiario
         .fields("Codigo") = Tipo_RUC_CI.Codigo_RUC_CI
         .Update
          sSQL = "SELECT * " _
               & "FROM Clientes " _
               & "WHERE Codigo = '" & Tipo_RUC_CI.Codigo_RUC_CI & "' "
          Select_Adodc AdoPrueba, sSQL
          If AdoPrueba.Recordset.RecordCount <= 0 Then
             SetAdoAddNew "Clientes"
             SetAdoFields "Codigo", Tipo_RUC_CI.Codigo_RUC_CI
             SetAdoFields "TD", Tipo_RUC_CI.Tipo_Beneficiario
             SetAdoFields "CI_RUC", Tipo_RUC_CI.RUC_CI
             SetAdoFields "Fecha", .fields("Fecha_Registro")
             SetAdoFields "Fecha_N", .fields("Fecha_Registro")
             SetAdoFields "Cliente", UCaseStrg(TrimStrg(.fields("Beneficiario")))
             SetAdoFields "Direccion", .fields("Direccion") & ", " & .fields("Credito_No")
             SetAdoFields "DirNum", .fields("TP")
             SetAdoFields "Lugar_Trabajo", .fields("Lugar_Trabajo")
             SetAdoFields "Prov", CodigoProv
             SetAdoFields "Pais", CodigoPais
             SetAdoFields "Ciudad", NombreCiudad
             SetAdoFields "Grupo", NumEmpresa
             SetAdoFields "Telefono", Telefono1
             SetAdoFields "Archivo_Foto", "SINFOTO"
             SetAdoUpdate
          End If
         .MoveNext
       Loop
   End If
  End With
  RatonNormal
  
End Sub

Private Sub Respaldar_Base_Total()
Dim Nombre_Base_Respaldo As String
   'Nombre de la base de datos a respaldar
    Nombre_Base_Respaldo = strNombreBaseDatos & " " & Replace(Format(FechaSistema, "yyyy/MM/dd"), "/", "-")
    CDialogDir.Filename = Nombre_Base_Respaldo
    RutaOrigen = SelectDialogFile(CDialogDir, RutaSysBases & "\Datos\Total")
    'RutaOrigen = CDialogDir.Filename
    If RutaOrigen <> "" Then
        RatonReloj
        Progreso_Barra.Mensaje_Box = "RESPALDANDO BASE DE DATOS COMPLETA"
        Progreso_Iniciar
        Progreso_Barra.Incremento = 0
        Progreso_Barra.Valor_Maximo = 10
        
        Progreso_Esperar
       'Aplicamos la base a reslpaldar
        sSQL = "USE " & strNombreBaseDatos & "; "
        Ejecutar_SQL_SP sSQL
        
        Progreso_Esperar
       'Colocamos la base en modo simple
        sSQL = "ALTER DATABASE " & strNombreBaseDatos & " " _
             & "SET RECOVERY SIMPLE; "
        Ejecutar_SQL_SP sSQL
        
        Progreso_Esperar
       'Reducimos el log al minimo
        sSQL = "DBCC SHRINKFILE (" & strNombreBaseDatos & "_Log, 1); "
        Ejecutar_SQL_SP sSQL
        
        Progreso_Esperar
       'Volvemos a colocar la base a FULL
        sSQL = "ALTER DATABASE " & strNombreBaseDatos & "SET RECOVERY FULL; "
        Ejecutar_SQL_SP sSQL
        
        Progreso_Esperar
       'Empezando a respaldar la base de datos
        sSQL = "BACKUP DATABASE " & strNombreBaseDatos & "TO DISK = '" & RutaOrigen & ".bak' " & "WITH FORMAT; "
       'MsgBox sSQL
       
        sSQL = "BACKUP DATABASE " & strNombreBaseDatos & "TO DISK = '" & RutaOrigen & ".bak' WITH FORMAT, MEDIANAME = 'Z_SQLServerBackups', " _
             & "NAME = 'Full Backup of " & strNombreBaseDatos & "'; "
        Ejecutar_SQL_SP sSQL
       'Fin del Proceso respaldo de la base de datos
        Progreso_Final
        RatonNormal
        
        '--Enter this command at the PowerShell command prompt,
        'C:\PS> Backup-SqlDatabase -ServerInstance Computer\Instance -Database MyDB -BackupAction Database
        
    End If
End Sub

Private Sub DGCodigos_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
     With AdoCodigos.Recordset
      If .RecordCount > 0 Then
         .MoveNext
          If .EOF Then .MoveFirst
      End If
     End With
  End If
End Sub

Private Sub DGCuentas_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  If CtrlDown And KeyCode = vbKeyP Then ImprimirAdodc AdoCuentas, True, 1, 9
  If KeyCode = vbKeyF1 Then
     DGCuentas.Visible = False
     GenerarDataTexto FSeteos, AdoCuentas
     DGCuentas.Visible = True
  End If
End Sub

Private Sub DGFormato_DblClick()
  Si_No = True
  Titulo = "Seteos de Documentos"
  Mensajes = "Formato Propio?"
  sSQL = "SELECT TP,Campo,Encabezado,Pos_X,Pos_Y,Porte,Item " _
       & "FROM Seteos_Documentos "
  If BoxMensaje = vbYes Then
     Si_No = True
     sSQL = sSQL & "WHERE Item = '000' " _
          & "AND TP = '" & DGFormato.Columns(0) & "' "
     DGSeteosPRN.Caption = "SETEOS DE DOCUMENTOS FORMATO PROPIO"
     TipoDoc = DGFormato.Columns(0)
  Else
     Si_No = False
     sSQL = sSQL & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND TP = 'P" & DGFormato.Columns(0) & "' "
     DGSeteosPRN.Caption = "SETEOS DE DOCUMENTOS FORMATO PERSONAL"
     TipoDoc = "P" & DGFormato.Columns(0)
  End If
  sSQL = sSQL & "ORDER BY Campo "
  Select_Adodc_Grid DGSeteosPRN, AdoSeteosPRN, sSQL
End Sub

Private Sub DGFormato_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Public Sub SetearCtasInventario(CtaFields As String)
  Si_No = True
  For IE = 0 To ContCtas - 1
      If CtaFields = CtasProc(IE).Cta Then Si_No = False
  Next IE
  If Si_No Then
     IE = 0
     While IE < ContCtas
        If CtasProc(IE).Cta = Ninguno Then
           CtasProc(IE).Cta = CtaFields
           IE = ContCtas + 1
        End If
        IE = IE + 1
     Wend
  End If
End Sub

Public Sub InsValorCtaInv(NCta As String, _
                          NValor As Currency)
  For IE = 0 To ContCtas - 1
      If CtasProc(IE).Cta = NCta Then
         CtasProc(IE).Valor = CtasProc(IE).Valor + Round(NValor, 2)
      End If
  Next IE
End Sub

Public Sub GenerarTablaEnArchivoPlano(FechaResp As String, _
                                      NombreTabla As String, _
                                      DtaAux As Adodc)
Dim NumFile As Integer
Dim RutaGeneraFile As String
Dim CaptionOld As String
Dim NombreFile As String
Dim CadFileCampos As String
Dim CadFileReg As String
Dim ContadorReg As Long
Dim TotalCampo As Integer
Dim ValorBool As String
Dim ILng As Long
RatonReloj
ContadorReg = 0
If FileResp <= 0 Then FileResp = 1
With DtaAux.Recordset
 If .RecordCount > 0 Then
    .MoveLast
     TotalReg = .RecordCount
     TotalCampo = .fields.Count - 1
    .MoveFirst
     'NombreFile = "Z" & NombreTabla & ".dbs"
     NombreFile = NombreTabla & ".dbs"
     'TextoFileEmp = TextoFileEmp & vbCrLf & NombreFile & " => " & NombreTabla
     RutaGeneraFile = RutaSistema & "\BASES\UPDATE_DB\" & NombreFile
     Progreso_Barra.Incremento = 0
     Progreso_Barra.Valor_Maximo = TotalCampo
     Progreso_Barra.Mensaje_Box = NombreFile
     Progreso_Esperar
    'MsgBox NombreFile
     NumFile = FreeFile
     'Contador = 0
     Open RutaGeneraFile For Output As #NumFile ' Abre el archivo.
     FAConLineas = False
     'Print #NumFile, Format(TotalReg, "##0") & " - " & NombreTabla
     ReDim TipoC(TotalCampo) As Campos_Tabla
     For ILng = 0 To TotalCampo
         TipoC(ILng).Campo = CompilarString(.fields(ILng).Name)
         TipoC(ILng).Ancho = AnchoTipoCampoTexto(.fields(ILng))
     Next ILng
     CadFileCampos = ""
     For ILng = 0 To TotalCampo
         CadFileCampos = CadFileCampos & TrimStrg(TipoC(ILng).Campo) & ","
     Next ILng
     
     CadFileCampos = MidStrg(CadFileCampos, 1, Len(CadFileCampos) - 1)
     
    'Grabamos el encabezado de la tabla en el archivo plano
     'CadFileCampos = "INSERT INTO " & NombreTabla & " (" & CadFileCampos & ") VALUES "
     Print #NumFile, "INSERT INTO " & NombreTabla & " (" & CadFileCampos & ") VALUES "
     
     CadFileCampos = ""
    
     'FAConLineas = True
    .MoveFirst
     Do While Not .EOF
        ContadorReg = ContadorReg + 1
        'FSeteos.Caption = NombreTabla _
        '               & ": Procesando(" & Format(ContadorReg / TotalReg, "00%") _
        '               & ") " & String$(ContadorReg Mod 40, "|")
        CadFileReg = ""
        For ILng = 0 To TotalCampo
            If IsNull(.fields(ILng)) Or IsEmpty(.fields(ILng)) Then Codigo4 = "0" Else Codigo4 = CStr(.fields(ILng))
            Select Case .fields(ILng).Type
              Case TadBoolean
                   If Codigo4 = Ninguno Then Codigo4 = "0"
                   Codigo4 = CStr(CInt(CBool(Codigo4)))
                   Codigo4 = SetearBlancos(Codigo4, 2, 0, True, FAConLineas)
                   Codigo4 = Replace(Codigo4, "-1", "1")
              Case TadByte, TadInteger, TadLong
                   If .fields(ILng).Name = "Item" Then
                       Codigo4 = SetearBlancos(Format(.fields(ILng), "000"), 3, 0, False, FAConLineas)
                   Else
                       Codigo4 = SetearBlancos(Codigo4, 0, 0, True, FAConLineas)
                   End If
              Case TadSingle, TadDouble, TadCurrency
                   Codigo4 = SetearBlancos(Codigo4, 0, 0, True, FAConLineas, True)
              Case TadText
                   If UCaseStrg(.fields(ILng).Name) = "RUC_CI" Then Codigo4 = CompilarRUC_CI(.fields(ILng))
                   If UCaseStrg(.fields(ILng).Name) = "RUC" Then Codigo4 = CompilarRUC_CI(.fields(ILng))
                   If UCaseStrg(.fields(ILng).Name) = "CI" Then Codigo4 = CompilarRUC_CI(.fields(ILng))
                   If UCaseStrg(.fields(ILng).Name) = "CEDULA" Then Codigo4 = CompilarRUC_CI(.fields(ILng))
                   'If Codigo4 = "0" Then Codigo4 = Ninguno
                   If Codigo4 = " " Then Codigo4 = Ninguno
                   Codigo4 = "'" & Codigo4 & "'"
              Case TadDate, TadDate1
                   Codigo4 = "'" & BuscarFecha(.fields(ILng)) & "'"
              Case Else
                   If Codigo4 = "0" Then Codigo4 = Ninguno
                   If Codigo4 = " " Then Codigo4 = Ninguno
                   Codigo4 = "'" & Codigo4 & "'"
            End Select
            Codigo4 = Replace(Codigo4, vbCrLf, Chr(170))
            Codigo4 = Replace(Codigo4, Chr(34), Chr(239))
            Codigo4 = Replace(Codigo4, ";", ",")
            Codigo4 = Replace(Codigo4, Chr(255), " ")
            Codigo4 = Replace(Codigo4, Chr(254), " ")
'''            Codigo4 = Replace(Codigo4, "á", "a")
'''            Codigo4 = Replace(Codigo4, "é", "e")
'''            Codigo4 = Replace(Codigo4, "í", "i")
'''            Codigo4 = Replace(Codigo4, "ó", "o")
'''            Codigo4 = Replace(Codigo4, "ú", "u")
'''            Codigo4 = Replace(Codigo4, "ñ", "n")
'''            Codigo4 = Replace(Codigo4, "Á", "A")
'''            Codigo4 = Replace(Codigo4, "É", "E")
'''            Codigo4 = Replace(Codigo4, "Í", "I")
'''            Codigo4 = Replace(Codigo4, "Ó", "O")
'''            Codigo4 = Replace(Codigo4, "Ú", "U")
'''            Codigo4 = Replace(Codigo4, "Ñ", "N")
'''            Codigo4 = Replace(Codigo4, "ü", "u")
'''            Codigo4 = Replace(Codigo4, "Ü", "U")
            Codigo4 = Replace(Codigo4, "&", "Y")
            Codigo4 = Replace(Codigo4, "#", "No ")
            Codigo4 = TrimStrg(Codigo4) & ","
            CadFileReg = CadFileReg & Codigo4
        Next ILng
        CadFileReg = MidStrg(CadFileReg, 1, Len(CadFileReg) - 1)
        
        CadFileCampos = CadFileCampos & "(" & CadFileReg & ")," & vbCrLf
        
        Progreso_Esperar
       .MoveNext
     Loop
    .MoveFirst
     FileResp = FileResp + 1
     CadFileCampos = MidStrg(CadFileCampos, 1, Len(CadFileCampos) - 3) & ";"
     Print #NumFile, CadFileCampos;
 End If
End With
Close #NumFile
RatonNormal
'MsgBox "Archivo: (" & NombreTabla & ")" & RutaGeneraFile & " Procesado."
End Sub

'''Public Sub GenerarTablaEnArchivoPlano(FechaResp As String, _
'''                                      NombreTabla As String, _
'''                                      DtaAux As Adodc)
'''Dim NumFile As Integer
'''Dim RutaGeneraFile As String
'''Dim CaptionOld As String
'''Dim NombreFile As String
'''Dim CadFileReg As String
'''Dim ContadorReg As Long
'''Dim TotalCampo As Integer
'''Dim ValorBool As String
'''Dim ILng As Long
'''RatonReloj
'''ContadorReg = 0
'''If FileResp <= 0 Then FileResp = 1
'''With DtaAux.Recordset
''' If .RecordCount > 0 Then
'''    .MoveLast
'''     TotalReg = .RecordCount
'''     TotalCampo = .fields.Count - 1
'''    .MoveFirst
'''     'NombreFile = "Z" & NombreTabla & ".dbs"
'''     NombreFile = NombreTabla & ".dbs"
'''     'TextoFileEmp = TextoFileEmp & vbCrLf & NombreFile & " => " & NombreTabla
'''     RutaGeneraFile = RutaSistema & "\BASES\UPDATE_DB\" & NombreFile
'''     Progreso_Barra.Incremento = 0
'''     Progreso_Barra.Valor_Maximo = TotalCampo
'''     Progreso_Barra.Mensaje_Box = NombreFile
'''     Progreso_Esperar
'''    'MsgBox NombreFile
'''     NumFile = FreeFile
'''     'Contador = 0
'''     Open RutaGeneraFile For Output As #NumFile ' Abre el archivo.
'''     FAConLineas = False
'''     'Print #NumFile, Format(TotalReg, "##0") & " - " & NombreTabla
'''     ReDim TipoC(TotalCampo) As Campos_Tabla
'''     For ILng = 0 To TotalCampo
'''         TipoC(ILng).Campo = CompilarString(.fields(ILng).Name)
'''         TipoC(ILng).Ancho = AnchoTipoCampoTexto(.fields(ILng))
'''     Next ILng
'''     CadFileReg = ""
'''     For ILng = 0 To TotalCampo
'''         CadFileReg = CadFileReg & TrimStrg(TipoC(ILng).Campo) & ";"
'''     Next ILng
'''
'''     CadFileReg = MidStrg(CadFileReg, 1, Len(CadFileReg) - 1)
'''
'''    'Grabamos el encabezado de la tabla en el archivo plano
'''     Print #NumFile, CadFileReg
'''
'''     'FAConLineas = True
'''    .MoveFirst
'''     Do While Not .EOF
'''        ContadorReg = ContadorReg + 1
'''        'FSeteos.Caption = NombreTabla _
'''        '               & ": Procesando(" & Format(ContadorReg / TotalReg, "00%") _
'''        '               & ") " & String$(ContadorReg Mod 40, "|")
'''        CadFileReg = ""
'''        For ILng = 0 To TotalCampo
'''            If IsNull(.fields(ILng)) Or IsEmpty(.fields(ILng)) Then Codigo4 = "0" Else Codigo4 = CStr(.fields(ILng))
'''            Select Case .fields(ILng).Type
'''              Case TadBoolean
'''                   If Codigo4 = Ninguno Then Codigo4 = "0"
'''                   Codigo4 = CStr(CInt(CBool(Codigo4)))
'''                   Codigo4 = SetearBlancos(Codigo4, 2, 0, True, FAConLineas)
'''                   Codigo4 = Replace(Codigo4, "-1", "1")
'''              Case TadByte, TadInteger, TadLong
'''                   If .fields(ILng).Name = "Item" Then
'''                       Codigo4 = SetearBlancos(Format(.fields(ILng), "000"), 3, 0, False, FAConLineas)
'''                   Else
'''                       Codigo4 = SetearBlancos(Codigo4, 0, 0, True, FAConLineas)
'''                   End If
'''              Case TadSingle, TadDouble, TadCurrency
'''                   Codigo4 = SetearBlancos(Codigo4, 0, 0, True, FAConLineas, True)
'''              Case TadText
'''                   If UCaseStrg(.fields(ILng).Name) = "RUC_CI" Then Codigo4 = CompilarRUC_CI(.fields(ILng))
'''                   If UCaseStrg(.fields(ILng).Name) = "RUC" Then Codigo4 = CompilarRUC_CI(.fields(ILng))
'''                   If UCaseStrg(.fields(ILng).Name) = "CI" Then Codigo4 = CompilarRUC_CI(.fields(ILng))
'''                   If UCaseStrg(.fields(ILng).Name) = "CEDULA" Then Codigo4 = CompilarRUC_CI(.fields(ILng))
'''                   'If Codigo4 = "0" Then Codigo4 = Ninguno
'''                   If Codigo4 = " " Then Codigo4 = Ninguno
'''                   Codigo4 = "^" & Codigo4 & "^"
'''              Case TadDate, TadDate1
'''                   Codigo4 = "^" & BuscarFecha(.fields(ILng)) & "^"
'''              Case Else
'''                   If Codigo4 = "0" Then Codigo4 = Ninguno
'''                   If Codigo4 = " " Then Codigo4 = Ninguno
'''                   Codigo4 = "^" & Codigo4 & "^"
'''            End Select
'''            Codigo4 = Replace(Codigo4, vbCrLf, Chr(170))
'''            Codigo4 = Replace(Codigo4, Chr(34), Chr(239))
'''            Codigo4 = Replace(Codigo4, ";", ",")
'''            Codigo4 = Replace(Codigo4, Chr(255), " ")
'''            Codigo4 = Replace(Codigo4, Chr(254), " ")
''''''            Codigo4 = Replace(Codigo4, "á", "a")
''''''            Codigo4 = Replace(Codigo4, "é", "e")
''''''            Codigo4 = Replace(Codigo4, "í", "i")
''''''            Codigo4 = Replace(Codigo4, "ó", "o")
''''''            Codigo4 = Replace(Codigo4, "ú", "u")
''''''            Codigo4 = Replace(Codigo4, "ñ", "n")
''''''            Codigo4 = Replace(Codigo4, "Á", "A")
''''''            Codigo4 = Replace(Codigo4, "É", "E")
''''''            Codigo4 = Replace(Codigo4, "Í", "I")
''''''            Codigo4 = Replace(Codigo4, "Ó", "O")
''''''            Codigo4 = Replace(Codigo4, "Ú", "U")
''''''            Codigo4 = Replace(Codigo4, "Ñ", "N")
''''''            Codigo4 = Replace(Codigo4, "ü", "u")
''''''            Codigo4 = Replace(Codigo4, "Ü", "U")
'''            Codigo4 = Replace(Codigo4, "&", "Y")
'''            Codigo4 = Replace(Codigo4, "#", "No ")
'''            Codigo4 = TrimStrg(Codigo4) & ";"
'''            CadFileReg = CadFileReg & Codigo4
'''        Next ILng
'''
'''        CadFileReg = MidStrg(CadFileReg, 1, Len(CadFileReg) - 1)
'''
'''        Print #NumFile, CadFileReg
'''        Progreso_Esperar
'''       .MoveNext
'''     Loop
'''    .MoveFirst
'''     FileResp = FileResp + 1
''' End If
'''End With
'''Close #NumFile
'''RatonNormal
'''End Sub

Public Sub AbrirCamposSQL(NumFile As Integer)
    Cod_Emp = "": Cod_Base = "": Cod_Field = ""
    Line Input #NumFile, Cod_Base
    TotalReg = CLng(SinEspaciosIzq(Cod_Base))
    Cod_Base = SinEspaciosDer(Cod_Base)
    Line Input #NumFile, Cod_Field
    'MsgBox Cod_Base & vbCrLf & Cod_Field
    CantCampos = 0
    For I = 1 To Len(Cod_Field)
        If MidStrg(Cod_Field, I, 1) = "|" Then CantCampos = CantCampos + 1
    Next I
    ReDim TipoC(CantCampos) As Campos_Tabla
    No_Desde = 1: No_Hasta = 1
    Cadena = Cod_Field
    For I = 1 To CantCampos
        Do
           No_Hasta = No_Hasta + 1
        Loop Until MidStrg(Cadena, No_Hasta, 1) = "|"
        TipoC(I).Campo = TrimStrg(MidStrg(Cadena, No_Desde, No_Hasta - 1))
        Cadena = MidStrg(Cadena, No_Hasta + 1, Len(Cadena))
        No_Desde = 1: No_Hasta = 1
    Next I
End Sub

Private Sub CheckSupervisor_Click()
If CheckSupervisor.value Then
   For I = 0 To 5
       CheckNivel(I).value = False
   Next I
End If
End Sub

Private Sub Command11_Click()
Dim IdMod As Integer
Dim IdEmp As Integer
Dim CodMod As String
Dim CodEmp As String
    RatonReloj
    TextoValido TextClave
    Codigo = Codigo_Usuario(DCUsuario)
    sSQL = "DELETE * " _
         & "FROM Acceso_Empresa " _
         & "WHERE Codigo = '" & Codigo & "' "
    Ejecutar_SQL_SP sSQL
  
    For IdMod = 0 To LstModulos.ListCount - 1
        If LstModulos.Selected(IdMod) Then
           CodMod = Ninguno
           sSQL = "SELECT * " _
                & "FROM Modulos " _
                & "WHERE Aplicacion = '" & LstModulos.List(IdMod) & "' "
           Select_Adodc AdoPrueba, sSQL
           If AdoPrueba.Recordset.RecordCount > 0 Then CodMod = AdoPrueba.Recordset.fields("Modulo")
           For IdEmp = 0 To LstEmpresas.ListCount - 1
               If LstEmpresas.Selected(IdEmp) Then
                  CodEmp = Ninguno
                  sSQL = "SELECT * " _
                       & "FROM Empresas " _
                       & "WHERE Empresa = '" & LstEmpresas.List(IdEmp) & "' "
                  Select_Adodc AdoPrueba, sSQL
                  If AdoPrueba.Recordset.RecordCount > 0 Then CodEmp = AdoPrueba.Recordset.fields("Item")
                 'Insertamos si existe seteos
                  If (Len(CodMod) + Len(CodEmp)) > 2 Then
                     SetAdoAddNew "Acceso_Empresa"
                     SetAdoFields "Codigo", Codigo
                     SetAdoFields "Modulo", CodMod
                     SetAdoFields "Item", CodEmp
                     SetAdoUpdate
                  End If
               End If
           Next IdEmp
        End If
    Next IdMod
    Codigo1 = Ninguno
    With AdoBodega.Recordset
     If .RecordCount > 0 Then
        .MoveFirst
        .Find ("Bodega = '" & DCBodega & "' ")
         If Not .EOF Then Codigo1 = .fields("CodBod")
     End If
    End With
   
    sSQL = "SELECT * " _
         & "FROM Accesos " _
         & "WHERE Codigo = '" & Codigo & "' "
    Select_Adodc AdoAux, sSQL
    With AdoAux.Recordset
     If .RecordCount > 0 Then
        .fields("Clave") = TextClave
        .fields("Usuario") = TxtUsuario
        .fields("CodBod") = Codigo1
        .fields("TODOS") = CheckSupervisor.value
        .fields("Nivel_1") = CheckNivel(0).value
        .fields("Nivel_2") = CheckNivel(1).value
        .fields("Nivel_3") = CheckNivel(2).value
        .fields("Nivel_4") = CheckNivel(3).value
        .fields("Nivel_5") = CheckNivel(4).value
        .fields("Nivel_6") = CheckNivel(5).value
        .fields("Nivel_7") = CheckNivel(6).value
        .fields("Supervisor") = CheckSupervisor.value
         
           .fields("Primaria") = CBool(OpcP.value)
           .fields("Secundaria") = CBool(OpcS.value)
           .fields("Bachillerato") = CBool(OpcB.value)
         
        .Update
     End If
    End With
    RatonNormal
    MsgBox "Niveles de seguridad grabado"
End Sub

Private Sub Command12_Click()
 Unload FSeteos
End Sub

Private Sub Command13_Click()
  FechaValida MBPeriodo
  FechaFin = MBPeriodo
  If ClaveSupervisor Then
     Titulo = "CIERRE DEL PERIODO EDUCATIVO"
     Mensajes = "Seguro de realizar el cierre del periodo Educativo del: " & FechaFin
     If BoxMensaje = vbYes Then
        RatonReloj
        Progreso_Iniciar
        Progreso_Barra.Mensaje_Box = "CERRANDO PERIODO EDUCATIVO DEL: " & FechaFin
        
       'Cerrando Tablas de Datos Educativos
        Progreso_Barra.Incremento = Progreso_Barra.Incremento + 2
        Progreso_Esperar
        Cerrar_Periodo_Tablas "Trans_Actas", FechaFin
        Progreso_Barra.Incremento = Progreso_Barra.Incremento + 2
        Progreso_Esperar
        Cerrar_Periodo_Tablas "Trans_Asistencia", FechaFin
        Progreso_Barra.Incremento = Progreso_Barra.Incremento + 2
        Progreso_Esperar
        Cerrar_Periodo_Tablas "Trans_Notas", FechaFin
        Progreso_Barra.Incremento = Progreso_Barra.Incremento + 2
        Progreso_Esperar
        Cerrar_Periodo_Tablas "Trans_Notas_Auxiliares", FechaFin
        Progreso_Barra.Incremento = Progreso_Barra.Incremento + 2
        Progreso_Esperar
        Cerrar_Periodo_Tablas "Trans_Notas_Grado", FechaFin
        Progreso_Barra.Incremento = Progreso_Barra.Incremento + 2
        Progreso_Esperar
        Cerrar_Periodo_Tablas "Trans_Promedios", FechaFin
        
       'Cerramos las tablas de Catalogos
        Progreso_Barra.Incremento = Progreso_Barra.Incremento + 2
        Progreso_Esperar
        CopiarAdoTablaPeriodo "Catalogo_Cursos", FechaFin
        Progreso_Barra.Incremento = Progreso_Barra.Incremento + 2
        Progreso_Esperar
        CopiarAdoTablaPeriodo "Catalogo_Equivalencia", FechaFin
        Progreso_Barra.Incremento = Progreso_Barra.Incremento + 2
        Progreso_Esperar
        CopiarAdoTablaPeriodo "Catalogo_Estudiantil", FechaFin
        Progreso_Barra.Incremento = Progreso_Barra.Incremento + 2
        Progreso_Esperar
        CopiarAdoTablaPeriodo "Catalogo_Examen_Grado", FechaFin
        Progreso_Barra.Incremento = Progreso_Barra.Incremento + 2
        Progreso_Esperar
        CopiarAdoTablaPeriodo "Catalogo_Materias", FechaFin
        Progreso_Barra.Incremento = Progreso_Barra.Incremento + 2
        Progreso_Esperar
        CopiarAdoTablaPeriodo "Catalogo_Periodo_Lectivo", FechaFin
        Progreso_Barra.Incremento = Progreso_Barra.Incremento + 2
        Progreso_Esperar
        CopiarAdoTablaPeriodo "Clientes_Matriculas", FechaFin
        Progreso_Final
        RatonNormal
    End If
  End If
End Sub

Private Sub Command16_Click()
Dim PathDibujo As String
Dim NombFilePict As String
Dim PosXPict As Single
Dim AnchoPict() As CtasAsiento
  Pagina = 1
  Codigo4 = TipoDoc
  PictFormatos.AutoRedraw = True
  PictFormatos.AutoSize = True
  PictFormatos.DrawWidth = 1
  PictFormatos.Cls
  PictFormatos.FontBold = True
  PictFormatos.FontSize = 24
  PictFormatos.FontName = TipoTimes
  NombFilePict = CodigosSinPuntos(CodigoUsuario) & NumEmpresa & NumModulo & Format(Pagina, "00") & ".GIF"
  RutaOrigen = RutaSistema & "\PRINTER\" & NombFilePict
  PictFormatos.Picture = LoadPicture(RutaSistema & "\FORMATOS\general\PaginaOf.emf")
  SavePicture PictFormatos.Image, RutaOrigen
  PictFormatos.Picture = LoadPicture()
  PictFormatos.Picture = LoadPicture(RutaSistema & "\FORMATOS\general\PaginaOf.emf")
  Contador = 0
  With AdoSeteosPRN.Recordset
   If .RecordCount > 0 Then
       Progreso_Barra.Incremento = 0
       Progreso_Barra.Valor_Maximo = .RecordCount
       Progreso_Barra.Mensaje_Box = "Formato de Impresion de: " & Codigo4
       Progreso_Esperar
      .MoveFirst
       Do While Not .EOF
          If (.fields("Pos_X") > 0) And (.fields("Pos_Y") > 0) Then
             PictFormatos.FontSize = .fields("Tamaño")
             'PictPrint_Texto PictFormatos, .Fields("Pos_X"), .Fields("Pos_Y"), .Fields("Encabezado")
          End If
          Progreso_Esperar
         .MoveNext
       Loop
   End If
  End With
  SavePicture PictFormatos.Image, RutaOrigen
  Progreso_Esperar
  RatonNormal
 'FVerGrafico.Show 1
  sSQL = "SELECT TP,Campo,Encabezado,Pos_X,Pos_Y,Tamaño,Item " _
       & "FROM Seteos_Documentos "
  If Si_No Then
     sSQL = sSQL & "WHERE Item = '000' " _
          & "AND TP = '" & TipoDoc & "' "
     DGSeteosPRN.Caption = "SETEOS DE DOCUMENTOS FORMATO PROPIO"
  Else
     sSQL = sSQL & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND TP = '" & TipoDoc & "' "
     DGSeteosPRN.Caption = "SETEOS DE DOCUMENTOS FORMATO PERSONAL"
  End If
  sSQL = sSQL & "ORDER BY Campo "
  Select_Adodc_Grid DGSeteosPRN, AdoSeteosPRN, sSQL
  RatonNormal
End Sub

Private Sub Command17_Click(index As Integer)
  Distancia = Val(InputBox("Recorrer a la Izquierda:", "CAMBIO DE POSICIONES", "0"))
  If Distancia > 0 Then
    'Vambio de Valores
     sSQL = "UPDATE Seteos_Documentos "
     Select Case index
       Case 0: sSQL = sSQL & "SET Pos_X = Pos_X - " & Distancia & " "
       Case 1: sSQL = sSQL & "SET Pos_X = Pos_X + " & Distancia & " "
       Case 2: sSQL = sSQL & "SET Pos_Y = Pos_Y - " & Distancia & " "
       Case 3: sSQL = sSQL & "SET Pos_Y = Pos_Y + " & Distancia & " "
     End Select
     If Si_No Then
        sSQL = sSQL & "WHERE Item = '000' "
        DGSeteosPRN.Caption = "SETEOS DE DOCUMENTOS FORMATO PROPIO"
     Else
        sSQL = sSQL & "WHERE Item = '" & NumEmpresa & "' "
        DGSeteosPRN.Caption = "SETEOS DE DOCUMENTOS FORMATO PERSONAL"
     End If
     sSQL = sSQL & "AND TP = '" & TipoDoc & "' "
     Select Case index
       Case 0, 1: sSQL = sSQL & "AND Pos_X <> 0 "
       Case 2, 3: sSQL = sSQL & "AND Pos_Y <> 0 "
     End Select
     'MsgBox sSQL
     Ejecutar_SQL_SP sSQL
    'Volver a presentar valores
     sSQL = "SELECT TP,Campo,Encabezado,Pos_X,Pos_Y,Tamaño,Item " _
          & "FROM Seteos_Documentos "
     If Si_No Then
        sSQL = sSQL & "WHERE Item = '000' "
        DGSeteosPRN.Caption = "SETEOS DE DOCUMENTOS FORMATO PROPIO"
     Else
        sSQL = sSQL & "WHERE Item = '" & NumEmpresa & "' "
        DGSeteosPRN.Caption = "SETEOS DE DOCUMENTOS FORMATO PERSONAL"
     End If
     sSQL = sSQL & "AND TP = '" & TipoDoc & "' " _
          & "ORDER BY Campo "
     Select_Adodc_Grid DGSeteosPRN, AdoSeteosPRN, sSQL
     RatonNormal
  End If
End Sub

Private Sub Command2_Click()
Dim ITab As Long
Dim JCamp As Long
Dim KCamp As Long
Dim Proceder As Boolean
'Generamos Los Datos en Item = 000
Periodo = MBPeriodo
Proceder = True
If Periodo = "00/00/0000" Then
   Mensajes = "ESTE PERIODO ESTA MAL DIRECCIONADO, DESEA QUE LE CORRIJA, " _
            & "CASO CONTRARIO CANCELE Y LLAME A SU PROVEEDOR DEL SISTEMA."
   Titulo = "CAMBIO DE PERIODO"
   If BoxMensaje = vbYes Then
      Periodo = Ninguno
   Else
      Proceder = False
   End If
End If
If Proceder Then
    NumItem = TxtItem
    For ITab = 0 To LstTablas.ListCount - 1
        FSeteos.Caption = "Progeso del Cambio: " & Format(ITab / LstTablas.ListCount, "00%")
        Si_No = False
        Modificar = False
        Select Case MidStrg(LstTablas.List(ITab), 1, 4)
          Case "Asie", "Bala", "Empr", "Sete", "Form", "Tipo", "Tabl", "Modu"
              'No hace ninguna modificacion
          Case Else  ' Si pasamos a cambiar el numero de Item y el cierre
               If LstTablas.List(ITab) <> "Clientes_Facturacion" Then
                    sSQL = "SELECT " & Full_Fields(LstTablas.List(ITab)) & " " _
                         & "FROM " & LstTablas.List(ITab) & " " _
                         & "WHERE 0 = 1 "
                    Select_Adodc AdoComp, sSQL
                    RatonReloj
                    With AdoComp.Recordset
                     For JCamp = 0 To .fields.Count - 1
                         If .fields(JCamp).Name = "Item" Then Si_No = True
                         If .fields(JCamp).Name = "Periodo" Then Modificar = True
                     Next JCamp
                    End With
                    If Modificar And Si_No Then
                       sSQL = "UPDATE " & LstTablas.List(ITab) & " " _
                            & "SET Periodo = '" & Periodo & "' " _
                            & "WHERE Item = '" & NumEmpresa & "' " _
                            & "AND Periodo = '" & Periodo_Contable & "' "
                      'MsgBox sSQL
                       Ejecutar_SQL_SP sSQL
                    End If
    '''           If Si_No And Not (Modificar) Then
    '''              sSQL = "UPDATE " & LstTablas.List(ITab) & " " _
    '''                   & "SET Item = '" & TxtItem & "' " _
    '''                   & "WHERE Item = '" & NumEmpresa & "' "
    '''             'MsgBox sSQL
    '''              Ejecutar_SQL_SP sSQL
    '''           End If
               End If
        End Select
        RatonNormal
    Next ITab
    sSQL = "DELETE * " _
         & "FROM Empresas " _
         & "WHERE Item = '" & NumEmpresa & "' "
    'MsgBox sSQL
    ''Ejecutar_SQL_SP sSQL
    
    FSeteos.Caption = "Progeso del Cambio: 100%"
    MsgBox "Proceso Terminado"
End If
Unload FSeteos
RatonReloj
UnidadSistema
IngresarClave = False
ListEmp.Show 1
PonerDirEmpresa
End Sub

Private Sub Command21_Click()
Dim VectFields() As Variant
  RatonReloj
'  Progreso_Barra.Mensaje_Box = "Determinando Duplicados de: Seteos Documentos"
'  Progreso_Esperar
 'Campos de las Tablas a eliminar duplicados
  Eliminar_Duplicados_SP "Seteos_Documentos", "Item, TP, Campo"
  Eliminar_Duplicados_SP "Formato_Propio", "Item, Num, Codigo"
  Eliminar_Duplicados_SP "Formato", "Item, TP"
  Duplicar_Tabla_SP "Seteos_Documentos", "Seteos_Documentos_T"

 'Ahora pasamos a actualizar los formatos de Impresion de comprobantes
  sSQL = "INSERT INTO Seteos_Documentos_T (Item, TP, Campo, Encabezado, Pos_X, Pos_Y, Porte, X) " _
       & "SELECT Item, TP, Campo, Encabezado, Pos_X, Pos_Y, Porte, 'U' " _
       & "FROM Seteos_Documentos " _
       & "WHERE Item = '000' " _
       & "ORDER BY Item, TP, Campo, Encabezado "
  Ejecutar_SQL_SP sSQL
 
  sSQL = "UPDATE Seteos_Documentos_T " _
       & "SET X = '.' " _
       & "FROM Seteos_Documentos_T As SDT, Seteos_Documentos As SD " _
       & "WHERE SD.Item = '" & NumEmpresa & "' " _
       & "AND SUBSTRING(SD.TP,2,2) = SDT.TP " _
       & "AND SDT.Campo = SD.Campo "
  Ejecutar_SQL_SP sSQL
  
  sSQL = "DELETE * " _
       & "FROM Seteos_Documentos_T " _
       & "WHERE X = '.' "
  Ejecutar_SQL_SP sSQL
  
  sSQL = "INSERT INTO Seteos_Documentos (Item, TP, Campo, Encabezado, Pos_X, Pos_Y, Porte, X) " _
       & "SELECT '" & NumEmpresa & "', 'P'+TP, Campo, Encabezado, Pos_X, Pos_Y, Porte, X " _
       & "FROM Seteos_Documentos_T " _
       & "WHERE Item = '000' " _
       & "ORDER BY Item, TP, Campo, Encabezado "
  Ejecutar_SQL_SP sSQL
    
  sSQL = "DROP TABLE Seteos_Documentos_T "
  Ejecutar_SQL_SP sSQL

' Formatos
  sSQL = "DELETE * " _
       & "FROM Formato " _
       & "WHERE Item = '" & NumEmpresa & "' "
  Ejecutar_SQL_SP sSQL
  
  sSQL = "INSERT INTO Formato (Item, TP, Lineas, X) " _
       & "SELECT '" & NumEmpresa & "', TP, Lineas, X " _
       & "FROM Formato " _
       & "WHERE Item = '000' " _
       & "ORDER BY Item, TP "
  Ejecutar_SQL_SP sSQL
  
  sSQL = "SELECT TP,Lineas,Item " _
       & "FROM Formato " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "ORDER BY TP "
  Select_Adodc_Grid DGFormato, AdoFormato, sSQL
  
  RatonNormal
  MsgBox "Proceso Terminado"
End Sub

Private Sub Command3_Click()
Dim ITab As Long
Dim JCamp As Long
Dim KCamp As Long
'Generamos Los Datos en Item = 000
Periodo = MBPeriodo
NumItem = TxtItem

If Len(NumEmpresa) <= 1 Then NumEmpresa = ""

sSQL = "UPDATE Empresas " _
     & "SET Item = '" & TxtItem & "', Grupo = '" & TxtItem & "' " _
     & "WHERE Item = '" & NumEmpresa & "' "
Ejecutar_SQL_SP sSQL
For ITab = 0 To LstTablas.ListCount - 1
    Si_No = False
    Modificar = False
    If MidStrg(LstTablas.List(ITab), 1, 7) <> "Asiento" And MidStrg(LstTablas.List(ITab), 1, 8) <> "Balances" Then
       FSeteos.Caption = "Progeso del Cambio: " & Format(ITab / LstTablas.ListCount, "00%") & " -> " & LstTablas.List(ITab)
       sSQL = "SELECT * " _
            & "FROM " & LstTablas.List(ITab) & " "
       Select_Adodc AdoComp, sSQL
       RatonReloj
       With AdoComp.Recordset
        For JCamp = 0 To .fields.Count - 1
            If .fields(JCamp).Name = "Item" Then Si_No = True
            'If .Fields(JCamp).Name = "Periodo" Then Modificar = True
        Next JCamp
       End With
       If Si_No Then
          sSQL = "UPDATE " & LstTablas.List(ITab) & " " _
               & "SET Item = '" & TxtItem & "' " _
               & "WHERE Item = '" & NumEmpresa & "' "
       End If
'''       If NombreUsuario = "Administrador de Red" Then Modificar = False
'''       If Modificar Then sSQL = sSQL & "AND Periodo = '" & Periodo_Contable & "' "
       If Si_No Then Ejecutar_SQL_SP sSQL
    End If
    RatonNormal
Next ITab
FSeteos.Caption = "Progeso del Cambio: 100%"
MsgBox "Proceso Terminado"
Unload FSeteos
RatonReloj
UnidadSistema
IngresarClave = False
ListEmp.Show 1
'PonerDirEmpresa MDIAnexos, "CONTABILIDAD"
End Sub

Private Sub Command5_Click()
   If ClaveSupervisor Then
      Titulo = "Pregunta de Eliminar"
      Mensajes = "Seguro de Borrar el contenido de toda la base de datos"
      If BoxMensaje = vbYes Then
         RatonReloj
         BorrarBasesDatos False
         RatonNormal
      End If
      MsgBox "Proceso Terminado"
   End If
End Sub

Private Sub Command6_Click()
  Control_Procesos Normal, "Desactivar Usuario: " & DCUsuario
  sSQL = "UPDATE Accesos " _
       & "SET TODOS = " & Val(adFalse) & " " _
       & "WHERE Codigo = '" & Codigo & "' "
  Ejecutar_SQL_SP sSQL
End Sub

Private Sub Command7_Click()
Dim ITab As Long
Dim JCamp As Long
Dim KCamp As Long

 'Generamos Los Datos en Item = 000
  Mensajes = "Este proceso eliminara el periodo " & MBPeriodo & " de la " _
           & "empresa actual." & vbCrLf & vbCrLf _
           & "ESTA SEGURO DE PROSEGUIR?"
  Titulo = "ELIMINAR PERIODO"
  If BoxMensaje = vbYes Then Eliminar_Periodo_SP MBPeriodo

'''Periodo = MBPeriodo
'''NumItem = TxtItem
'''For ITab = 0 To LstTablas.ListCount - 1
'''    FSeteos.Caption = "Progeso del Cambio: " & Format(ITab / LstTablas.ListCount, "00%")
'''    Si_No = False
'''    Modificar = False
'''    If MidStrg(LstTablas.List(ITab), 1, 7) <> "Asiento" And MidStrg(LstTablas.List(ITab), 1, 8) <> "Balances" And _
'''       LstTablas.List(ITab) <> "Seteos_Documentos" And LstTablas.List(ITab) <> "Formato" Then
'''       sSQL = "SELECT * " _
'''            & "FROM " & LstTablas.List(ITab) & " "
'''       Select_Adodc AdoComp, sSQL
'''       RatonReloj
'''       With AdoComp.Recordset
'''        For JCamp = 0 To .Fields.Count - 1
'''            If .Fields(JCamp).Name = "Item" Then Si_No = True
'''            If .Fields(JCamp).Name = "Periodo" Then Modificar = True
'''        Next JCamp
'''       End With
'''       If Modificar And Si_No Then
'''          sSQL = "DELETE * " _
'''               & "FROM " & LstTablas.List(ITab) & " " _
'''               & "WHERE Item = '" & NumEmpresa & "' " _
'''               & "AND Periodo = '" & Periodo & "' "
'''          Ejecutar_SQL_SP sSQL
'''       End If
'''    End If
'''    RatonNormal
'''Next ITab
'''FSeteos.Caption = "Progeso del Cambio: 100%"
   MsgBox "Proceso Terminado"
'''    RatonReloj
'''    UnidadSistema
'''    IngresarClave = False
'''    ListEmp.Show 1
End Sub

Public Sub Procesar_Duplicados_Catalogo_SubModulos()
Dim vSingle As Single
Dim vInteger1 As Integer
Dim vInteger2 As Integer

 'CODIGO Y SUBMODULOS
  RatonReloj
 'Actualizacion de Facturas
  sSQL = "UPDATE Trans_Abonos " _
       & "SET TP = 'FA' " _
       & "WHERE TP = '.' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' "
  Ejecutar_SQL_SP sSQL
  sSQL = "UPDATE Detalle_Factura " _
       & "SET TC = 'FA' " _
       & "WHERE TC = '.' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' "
  Ejecutar_SQL_SP sSQL
  sSQL = "UPDATE Facturas " _
       & "SET TC = 'FA' " _
       & "WHERE TC = '.' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' "
  Ejecutar_SQL_SP sSQL
''  sSQL = "UPDATE Trans_Abonos " _
''       & "SET EstabRetencion = '001',PtoEmiRetencion = '001' " _
''       & "WHERE EstabRetencion = '.' " _
''       & "AND Item = '" & NumEmpresa & "' " _
''       & "AND Periodo = '" & Periodo_Contable & "' "
''  Ejecutar_SQL_SP sSQL
  
  sSQL = "UPDATE Trans_Abonos " _
       & "SET Banco = 'RETENCION IVA BIENES' " _
       & "WHERE Banco = 'RETENCION DEL IVA' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' "
  Ejecutar_SQL_SP sSQL
  
  sSQL = "UPDATE Trans_Abonos " _
       & "SET Banco = 'RETENCION FUENTE - 340' " _
       & "WHERE Banco = 'RETENCION EN LA FUENTE' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' "
  Ejecutar_SQL_SP sSQL
 'Lista de todos los clientes
  sSQL = "SELECT Codigo,Cliente " _
       & "FROM Clientes " _
       & "WHERE Grupo <> '.' " _
       & "ORDER BY Codigo "
  Select_Adodc AdoPrueba, sSQL
  
 'Catalogo_Interes
  sSQL = "SELECT * " _
       & "FROM Catalogo_Interes " _
       & "WHERE 0 = 1 "
  Select_Adodc AdoCuentas, sSQL
  ReDim VectFields(AdoCuentas.Recordset.fields.Count + 1) As Variant
  
  sSQL = "SELECT Item, Tipo, Interes, Desde, Hasta, COUNT(Tipo) As NumItem " _
       & "FROM Catalogo_Interes " _
       & "WHERE Item <> '.' " _
       & "GROUP BY Item, Tipo, Interes, Desde, Hasta " _
       & "HAVING COUNT(Tipo) > 1 "
  Select_Adodc AdoComp, sSQL
  With AdoComp.Recordset
   If .RecordCount > 0 Then
       Contador = 0
      .MoveFirst
       Do While Not .EOF
          Codigo = .fields("Tipo")
          Codigo2 = .fields("Item")
          vSingle = .fields("Interes")
          Contador = Contador + 1
          FSeteos.Caption = "Catalogo_CxCxP: " & Codigo & " -> " & Format(Contador / .RecordCount, "00%")
          sSQL = "SELECT * " _
               & "FROM Catalogo_Interes " _
               & "WHERE Tipo = '" & Codigo & "' " _
               & "AND Item = '" & Codigo2 & "' " _
               & "AND Interes = " & vSingle & " "
          Select_Adodc AdoCuentas, sSQL
          If AdoCuentas.Recordset.RecordCount > 0 Then
             For I = 0 To AdoCuentas.Recordset.fields.Count - 1
                VectFields(I) = AdoCuentas.Recordset.fields(I)
             Next I
             sSQL = "DELETE * " _
                  & "FROM Catalogo_Interes " _
                  & "WHERE Tipo = '" & Codigo & "' " _
                  & "AND Item = '" & Codigo2 & "' " _
                  & "AND Interes = " & vSingle & " "
             Ejecutar_SQL_SP sSQL
             SetAddNew AdoCuentas
             For I = 0 To AdoCuentas.Recordset.fields.Count - 1
                 SetFields AdoCuentas, AdoCuentas.Recordset.fields(I).Name, VectFields(I)
             Next I
             SetUpdate AdoCuentas
          End If
         .MoveNext
       Loop
   End If
  End With
  
 'Catalogo_Equivalencia
  sSQL = "SELECT * " _
       & "FROM Catalogo_Equivalencia " _
       & "WHERE 0 = 1 "
  Select_Adodc AdoCuentas, sSQL
  ReDim VectFields(AdoCuentas.Recordset.fields.Count + 1) As Variant
  
  sSQL = "SELECT Periodo, Item, Nivel, Equivalencia, COUNT(Equivalencia) As NumItem " _
       & "FROM Catalogo_Equivalencia " _
       & "WHERE Item <> '.' " _
       & "GROUP BY Periodo, Item, Nivel, Equivalencia " _
       & "HAVING COUNT(Equivalencia) > 1 "
  Select_Adodc AdoComp, sSQL
  With AdoComp.Recordset
   If .RecordCount > 0 Then
       Contador = 0
      .MoveFirst
       Do While Not .EOF
          Codigo1 = .fields("Item")
          Codigo2 = .fields("Periodo")
          Codigo3 = .fields("Nivel")
          Codigo4 = .fields("Equivalencia")
          Contador = Contador + 1
          FSeteos.Caption = "Catalogo_CxCxP: " & Codigo & " -> " & Format(Contador / .RecordCount, "00%")
          sSQL = "SELECT * " _
               & "FROM Catalogo_Equivalencia " _
               & "WHERE Item = '" & Codigo1 & "' " _
               & "AND Periodo = '" & Codigo2 & "' " _
               & "AND Nivel = '" & Codigo3 & "' " _
               & "AND Equivalencia = '" & Codigo4 & "' "
          Select_Adodc AdoCuentas, sSQL
          If AdoCuentas.Recordset.RecordCount > 0 Then
             For I = 0 To AdoCuentas.Recordset.fields.Count - 1
                VectFields(I) = AdoCuentas.Recordset.fields(I)
             Next I
             sSQL = "DELETE * " _
                  & "FROM Catalogo_Equivalencia " _
                  & "WHERE Item = '" & Codigo1 & "' " _
                  & "AND Periodo = '" & Codigo2 & "' " _
                  & "AND Nivel = '" & Codigo3 & "' " _
                  & "AND Equivalencia = '" & Codigo4 & "' "
             Ejecutar_SQL_SP sSQL
             SetAddNew AdoCuentas
             For I = 0 To AdoCuentas.Recordset.fields.Count - 1
                 SetFields AdoCuentas, AdoCuentas.Recordset.fields(I).Name, VectFields(I)
             Next I
             SetUpdate AdoCuentas
          End If
         .MoveNext
       Loop
   End If
  End With
  
 'Catalogo_Corresponsal
  sSQL = "SELECT * " _
       & "FROM Catalogo_Corresponsal " _
       & "WHERE 0 = 1 "
  Select_Adodc AdoCuentas, sSQL
  ReDim VectFields(AdoCuentas.Recordset.fields.Count + 1) As Variant
  
  sSQL = "SELECT Periodo, Item, Codigo_C, COUNT(Codigo_C) As NumItem " _
       & "FROM Catalogo_Corresponsal " _
       & "WHERE Item <> '.' " _
       & "GROUP BY Periodo, Item, Codigo_C " _
       & "HAVING COUNT(Codigo_C) > 1 "
  Select_Adodc AdoComp, sSQL
  With AdoComp.Recordset
   If .RecordCount > 0 Then
       Contador = 0
      .MoveFirst
       Do While Not .EOF
          Codigo1 = .fields("Item")
          Codigo2 = .fields("Periodo")
          Codigo3 = .fields("Codigo_C")
          Contador = Contador + 1
          FSeteos.Caption = "Catalogo_CxCxP: " & Codigo & " -> " & Format(Contador / .RecordCount, "00%")
          sSQL = "SELECT * " _
               & "FROM Catalogo_Corresponsal " _
               & "WHERE Item = '" & Codigo1 & "' " _
               & "AND Periodo = '" & Codigo2 & "' " _
               & "AND Codigo_C = '" & Codigo3 & "' "
          Select_Adodc AdoCuentas, sSQL
          If AdoCuentas.Recordset.RecordCount > 0 Then
             For I = 0 To AdoCuentas.Recordset.fields.Count - 1
                VectFields(I) = AdoCuentas.Recordset.fields(I)
             Next I
             sSQL = "DELETE * " _
                  & "FROM Catalogo_Corresponsal " _
                  & "WHERE Item = '" & Codigo1 & "' " _
                  & "AND Periodo = '" & Codigo2 & "' " _
                  & "AND Codigo_C = '" & Codigo3 & "' "
             Ejecutar_SQL_SP sSQL
             SetAddNew AdoCuentas
             For I = 0 To AdoCuentas.Recordset.fields.Count - 1
                 SetFields AdoCuentas, AdoCuentas.Recordset.fields(I).Name, VectFields(I)
             Next I
             SetUpdate AdoCuentas
          End If
         .MoveNext
       Loop
   End If
  End With
  
 'Catalogo_Cyber_Tiempo
  sSQL = "SELECT * " _
       & "FROM Catalogo_Cyber_Tiempo " _
       & "WHERE 0 = 1 "
  Select_Adodc AdoCuentas, sSQL
  ReDim VectFields(AdoCuentas.Recordset.fields.Count + 1) As Variant
  
  sSQL = "SELECT Item, Desde, Hasta, COUNT(Item) As NumItem " _
       & "FROM Catalogo_Cyber_Tiempo " _
       & "WHERE Item <> '.' " _
       & "GROUP BY Item, Desde, Hasta " _
       & "HAVING COUNT(Item) > 1 "
  Select_Adodc AdoComp, sSQL
  With AdoComp.Recordset
   If .RecordCount > 0 Then
       Contador = 0
      .MoveFirst
       Do While Not .EOF
          Codigo1 = .fields("Item")
          vInteger1 = .fields("Desde")
          vInteger2 = .fields("Hasta")
          Contador = Contador + 1
          FSeteos.Caption = "Catalogo_Cyber_Tiempo: " & Codigo & " -> " & Format(Contador / .RecordCount, "00%")
          sSQL = "SELECT * " _
               & "FROM Catalogo_Cyber_Tiempo " _
               & "WHERE Item = '" & Codigo1 & "' " _
               & "AND Desde = " & vInteger1 & " " _
               & "AND Hasta = " & vInteger2 & " "
          Select_Adodc AdoCuentas, sSQL
          If AdoCuentas.Recordset.RecordCount > 0 Then
             For I = 0 To AdoCuentas.Recordset.fields.Count - 1
                VectFields(I) = AdoCuentas.Recordset.fields(I)
             Next I
             sSQL = "DELETE * " _
                  & "FROM Catalogo_Cyber_Tiempo " _
                  & "WHERE Item = '" & Codigo1 & "' " _
                  & "AND Desde = " & vInteger1 & " " _
                  & "AND Hasta = " & vInteger2 & " "
             Ejecutar_SQL_SP sSQL
             SetAddNew AdoCuentas
             For I = 0 To AdoCuentas.Recordset.fields.Count - 1
                 SetFields AdoCuentas, AdoCuentas.Recordset.fields(I).Name, VectFields(I)
             Next I
             SetUpdate AdoCuentas
          End If
         .MoveNext
       Loop
   End If
  End With
  
 'Catalogo_Formularios
  sSQL = "SELECT * " _
       & "FROM Catalogo_Formularios " _
       & "WHERE 0 = 1 "
  Select_Adodc AdoCuentas, sSQL
  ReDim VectFields(AdoCuentas.Recordset.fields.Count + 1) As Variant
  
  sSQL = "SELECT Item, Formulario, Codigo, Detalle, COUNT(Detalle) As NumItem " _
       & "FROM Catalogo_Formularios " _
       & "WHERE Item <> '.' " _
       & "GROUP BY Item, Formulario, Codigo, Detalle " _
       & "HAVING COUNT(Detalle) > 1 "
  Select_Adodc AdoComp, sSQL
  With AdoComp.Recordset
   If .RecordCount > 0 Then
       Contador = 0
      .MoveFirst
       Do While Not .EOF
          Codigo1 = .fields("Item")
          Codigo2 = .fields("Formulario")
          Codigo3 = .fields("Codigo")
          Codigo4 = .fields("Detalle")
          Contador = Contador + 1
          FSeteos.Caption = "Catalogo_CxCxP: " & Codigo & " -> " & Format(Contador / .RecordCount, "00%")
          sSQL = "SELECT * " _
               & "FROM Catalogo_Formularios " _
               & "WHERE Item = '" & Codigo1 & "' " _
               & "AND Formulario = '" & Codigo2 & "' " _
               & "AND Codigo = '" & Codigo3 & "' " _
               & "AND Detalle = '" & Codigo4 & "' "
          Select_Adodc AdoCuentas, sSQL
          If AdoCuentas.Recordset.RecordCount > 0 Then
             For I = 0 To AdoCuentas.Recordset.fields.Count - 1
                VectFields(I) = AdoCuentas.Recordset.fields(I)
             Next I
             sSQL = "DELETE * " _
                  & "FROM Catalogo_Formularios " _
                  & "WHERE Item = '" & Codigo1 & "' " _
                  & "AND Formulario = '" & Codigo2 & "' " _
                  & "AND Codigo = '" & Codigo3 & "' " _
                  & "AND Detalle = '" & Codigo4 & "' "
             Ejecutar_SQL_SP sSQL
             SetAddNew AdoCuentas
             For I = 0 To AdoCuentas.Recordset.fields.Count - 1
                 SetFields AdoCuentas, AdoCuentas.Recordset.fields(I).Name, VectFields(I)
             Next I
             SetUpdate AdoCuentas
          End If
         .MoveNext
       Loop
   End If
  End With
  
 'Codigos Catalogo CxCxP
  sSQL = "SELECT * " _
       & "FROM Catalogo_CxCxP " _
       & "WHERE 0 = 1 "
  Select_Adodc AdoCuentas, sSQL
  ReDim VectFields(AdoCuentas.Recordset.fields.Count + 1) As Variant
  
  sSQL = "SELECT Periodo,Item,Codigo,Cta,COUNT(Cta) As NumItem " _
       & "FROM Catalogo_CxCxP " _
       & "WHERE Item <> '.' " _
       & "GROUP BY Periodo,Item,Codigo,Cta " _
       & "HAVING COUNT(Cta) > 1 "
  Select_Adodc AdoComp, sSQL
  With AdoComp.Recordset
   If .RecordCount > 0 Then
       Contador = 0
      .MoveFirst
       Do While Not .EOF
          Codigo = .fields("Codigo")
          Codigo1 = .fields("Periodo")
          Codigo2 = .fields("Item")
          Codigo3 = .fields("Cta")
          Contador = Contador + 1
          FSeteos.Caption = "Catalogo_CxCxP: " & Codigo & " -> " & Format(Contador / .RecordCount, "00%")
          sSQL = "SELECT * " _
               & "FROM Catalogo_CxCxP " _
               & "WHERE Codigo = '" & Codigo & "' " _
               & "AND Periodo = '" & Codigo1 & "' " _
               & "AND Item = '" & Codigo2 & "' " _
               & "AND Cta = '" & Codigo3 & "' "
          Select_Adodc AdoCuentas, sSQL
          If AdoCuentas.Recordset.RecordCount > 0 Then
             For I = 0 To AdoCuentas.Recordset.fields.Count - 1
                VectFields(I) = AdoCuentas.Recordset.fields(I)
             Next I
             sSQL = "DELETE * " _
                  & "FROM Catalogo_CxCxP " _
                  & "WHERE Codigo = '" & Codigo & "' " _
                  & "AND Periodo = '" & Codigo1 & "' " _
                  & "AND Item = '" & Codigo2 & "' " _
                  & "AND Cta = '" & Codigo3 & "' "
             Ejecutar_SQL_SP sSQL
             SetAddNew AdoCuentas
             For I = 0 To AdoCuentas.Recordset.fields.Count - 1
                 SetFields AdoCuentas, AdoCuentas.Recordset.fields(I).Name, VectFields(I)
             Next I
             SetUpdate AdoCuentas
          End If
         .MoveNext
       Loop
   End If
  End With
  
 'Codigos Catalogo SubCtas
  sSQL = "SELECT * " _
       & "FROM Catalogo_SubCtas " _
       & "WHERE 0 = 1 "
  Select_Adodc AdoCuentas, sSQL
  ReDim VectFields(AdoCuentas.Recordset.fields.Count + 1) As Variant
  
  sSQL = "SELECT TC,Periodo,Item,Codigo,Detalle,COUNT(Codigo) As NumItem " _
       & "FROM Catalogo_SubCtas " _
       & "WHERE Item <> '.' " _
       & "GROUP BY TC,Periodo,Item,Codigo,Detalle " _
       & "HAVING COUNT(Codigo) > 1 "
  Select_Adodc AdoComp, sSQL
  With AdoComp.Recordset
   If .RecordCount > 0 Then
       Contador = 0
      .MoveFirst
       Do While Not .EOF
          Codigo = .fields("Codigo")
          Codigo1 = .fields("Periodo")
          Codigo2 = .fields("Item")
          Codigo3 = .fields("Detalle")
          Codigo4 = .fields("TC")
          Contador = Contador + 1
          FSeteos.Caption = "Catalogo_SubCtas: " & Codigo & " -> " & Format(Contador / .RecordCount, "00%")
          sSQL = "SELECT * " _
               & "FROM Catalogo_SubCtas " _
               & "WHERE Codigo = '" & Codigo & "' " _
               & "AND Periodo = '" & Codigo1 & "' " _
               & "AND Item = '" & Codigo2 & "' " _
               & "AND Detalle = '" & Codigo3 & "' " _
               & "AND TC = '" & Codigo4 & "' "
          Select_Adodc AdoCuentas, sSQL
          If AdoCuentas.Recordset.RecordCount > 0 Then
             For I = 0 To AdoCuentas.Recordset.fields.Count - 1
                VectFields(I) = AdoCuentas.Recordset.fields(I)
             Next I
             sSQL = "DELETE * " _
                  & "FROM Catalogo_SubCtas " _
                  & "WHERE Codigo = '" & Codigo & "' " _
                  & "AND Periodo = '" & Codigo1 & "' " _
                  & "AND Item = '" & Codigo2 & "' " _
                  & "AND Detalle = '" & Codigo3 & "' " _
                  & "AND TC = '" & Codigo4 & "' "
             Ejecutar_SQL_SP sSQL
             SetAddNew AdoCuentas
             For I = 0 To AdoCuentas.Recordset.fields.Count - 1
                 SetFields AdoCuentas, AdoCuentas.Recordset.fields(I).Name, VectFields(I)
             Next I
             SetUpdate AdoCuentas
          End If
         .MoveNext
       Loop
   End If
  End With
  
  sSQL = "SELECT TC,Periodo,Item,Detalle,COUNT(Detalle) As NumItem " _
       & "FROM Catalogo_SubCtas " _
       & "WHERE Item <> '.' " _
       & "GROUP BY TC,Periodo,Item,Detalle " _
       & "HAVING COUNT(Detalle) > 1 "
  Select_Adodc AdoComp, sSQL
  With AdoComp.Recordset
   If .RecordCount > 0 Then
       Contador = 0
      .MoveFirst
       Do While Not .EOF
          'Codigo = .Fields("Codigo")
          Codigo1 = .fields("Periodo")
          Codigo2 = .fields("Item")
          Codigo3 = .fields("Detalle")
          Codigo4 = .fields("TC")
          Contador = Contador + 1
          FSeteos.Caption = "Catalogo_SubCtas: " & Codigo & " -> " & Format(Contador / .RecordCount, "00%")
          sSQL = "SELECT * " _
               & "FROM Catalogo_SubCtas " _
               & "WHERE Periodo = '" & Codigo1 & "' " _
               & "AND Item = '" & Codigo2 & "' " _
               & "AND Detalle = '" & Codigo3 & "' " _
               & "ORDER BY Codigo "
          Select_Adodc AdoCuentas, sSQL
          If AdoCuentas.Recordset.RecordCount > 0 Then
             Codigo = AdoCuentas.Recordset.fields("Codigo")
             For I = 0 To AdoCuentas.Recordset.fields.Count - 1
                 VectFields(I) = AdoCuentas.Recordset.fields(I)
             Next I
             Do While Not AdoCuentas.Recordset.EOF
                CodigoA = AdoCuentas.Recordset.fields("Codigo")
                sSQL = "DELETE * " _
                     & "FROM Catalogo_SubCtas " _
                     & "WHERE Codigo = '" & CodigoA & "' " _
                     & "AND Periodo = '" & Codigo1 & "' " _
                     & "AND Item = '" & Codigo2 & "' "
                Ejecutar_SQL_SP sSQL
                If Codigo <> CodigoA Then
                   sSQL = "UPDATE Trans_SubCtas " _
                        & "SET Codigo = '" & Codigo & "' " _
                        & "WHERE Codigo = '" & CodigoA & "' "
                   Ejecutar_SQL_SP sSQL
                End If
                AdoCuentas.Recordset.MoveNext
             Loop
             SetAddNew AdoCuentas
             For I = 0 To AdoCuentas.Recordset.fields.Count - 1
                 SetFields AdoCuentas, AdoCuentas.Recordset.fields(I).Name, VectFields(I)
             Next I
             SetUpdate AdoCuentas
          End If
         .MoveNext
       Loop
   End If
  End With
 
 'Codigos Catalogo de Cuentas
  sSQL = "SELECT * " _
       & "FROM Catalogo_Cuentas " _
       & "WHERE 0 = 1 "
  Select_Adodc AdoCuentas, sSQL
  ReDim VectFields(AdoCuentas.Recordset.fields.Count + 1) As Variant
  
  sSQL = "SELECT Periodo,Item,Codigo,COUNT(Codigo) As NumItem " _
       & "FROM Catalogo_Cuentas " _
       & "WHERE Item <> '.' " _
       & "GROUP BY Periodo,Item,Codigo " _
       & "HAVING COUNT(Codigo) > 1 "
  Select_Adodc AdoComp, sSQL
  With AdoComp.Recordset
   If .RecordCount > 0 Then
       Contador = 0
      .MoveFirst
       Do While Not .EOF
          Codigo = .fields("Codigo")
          Codigo1 = .fields("Periodo")
          Codigo2 = .fields("Item")
          Contador = Contador + 1
          FSeteos.Caption = "Catalogo_Cuentas: " & Codigo & " -> " & Format(Contador / .RecordCount, "00%")
          sSQL = "SELECT * " _
               & "FROM Catalogo_Cuentas " _
               & "WHERE Codigo = '" & Codigo & "' " _
               & "AND Periodo = '" & Codigo1 & "' " _
               & "AND Item = '" & Codigo2 & "' "
          Select_Adodc AdoCuentas, sSQL
          If AdoCuentas.Recordset.RecordCount > 0 Then
             For I = 0 To AdoCuentas.Recordset.fields.Count - 1
                 VectFields(I) = AdoCuentas.Recordset.fields(I)
             Next I
             sSQL = "DELETE * " _
                  & "FROM Catalogo_Cuentas " _
                  & "WHERE Codigo = '" & Codigo & "' " _
                  & "AND Periodo = '" & Codigo1 & "' " _
                  & "AND Item = '" & Codigo2 & "' "
             Ejecutar_SQL_SP sSQL
             SetAddNew AdoCuentas
             For I = 0 To AdoCuentas.Recordset.fields.Count - 1
                 SetFields AdoCuentas, AdoCuentas.Recordset.fields(I).Name, VectFields(I)
             Next I
             SetUpdate AdoCuentas
          End If
         .MoveNext
       Loop
   End If
  End With
 'Codigos Catalogo de Productos
  sSQL = "SELECT * " _
       & "FROM Catalogo_Productos " _
       & "WHERE 0 = 1 "
  Select_Adodc AdoCuentas, sSQL
  ReDim VectFields(AdoCuentas.Recordset.fields.Count + 1) As Variant
  
  sSQL = "SELECT Periodo,Item,Codigo_Inv,COUNT(Codigo_Inv) As NumItem " _
       & "FROM Catalogo_Productos " _
       & "WHERE Item <> '.' " _
       & "GROUP BY Periodo,Item,Codigo_Inv " _
       & "HAVING COUNT(Codigo_Inv) > 1 "
  Select_Adodc AdoComp, sSQL
  With AdoComp.Recordset
   If .RecordCount > 0 Then
       Contador = 0
      .MoveFirst
       Do While Not .EOF
          Codigo = .fields("Codigo_Inv")
          Codigo1 = .fields("Periodo")
          Codigo2 = .fields("Item")
          Contador = Contador + 1
          FSeteos.Caption = "Catalogo_Productos: " & Codigo & " -> " & Format(Contador / .RecordCount, "00%")
          sSQL = "SELECT * " _
               & "FROM Catalogo_Productos " _
               & "WHERE Codigo_Inv = '" & Codigo & "' " _
               & "AND Periodo = '" & Codigo1 & "' " _
               & "AND Item = '" & Codigo2 & "' "
          Select_Adodc AdoCuentas, sSQL
          If AdoCuentas.Recordset.RecordCount > 0 Then
             For I = 0 To AdoCuentas.Recordset.fields.Count - 1
                 VectFields(I) = AdoCuentas.Recordset.fields(I)
             Next I
             sSQL = "DELETE * " _
                  & "FROM Catalogo_Productos " _
                  & "WHERE Codigo_Inv = '" & Codigo & "' " _
                  & "AND Periodo = '" & Codigo1 & "' " _
                  & "AND Item = '" & Codigo2 & "' "
             Ejecutar_SQL_SP sSQL
             SetAddNew AdoCuentas
             For I = 0 To AdoCuentas.Recordset.fields.Count - 1
                 SetFields AdoCuentas, AdoCuentas.Recordset.fields(I).Name, VectFields(I)
             Next I
             SetUpdate AdoCuentas
          End If
         .MoveNext
       Loop
   End If
  End With
  
  sSQL = "SELECT Periodo,Item,Producto,COUNT(Codigo_Inv) As NumItem " _
       & "FROM Catalogo_Productos " _
       & "WHERE Item <> '.' " _
       & "GROUP BY Periodo,Item,Producto " _
       & "HAVING COUNT(Producto) > 1 "
  Select_Adodc AdoComp, sSQL
  With AdoComp.Recordset
   If .RecordCount > 0 Then
       Contador = 0
      .MoveFirst
       Do While Not .EOF
          Codigo = .fields("Producto")
          Codigo1 = .fields("Periodo")
          Codigo2 = .fields("Item")
          Contador = Contador + 1
          FSeteos.Caption = "Catalogo_Productos: " & Codigo & " -> " & Format(Contador / .RecordCount, "00%")
          sSQL = "SELECT * " _
               & "FROM Catalogo_Productos " _
               & "WHERE Producto = '" & Codigo & "' " _
               & "AND Periodo = '" & Codigo1 & "' " _
               & "AND Item = '" & Codigo2 & "' "
          Select_Adodc AdoCuentas, sSQL
          If AdoCuentas.Recordset.RecordCount > 0 Then
             For I = 0 To AdoCuentas.Recordset.fields.Count - 1
                 VectFields(I) = AdoCuentas.Recordset.fields(I)
             Next I
             sSQL = "DELETE * " _
                  & "FROM Catalogo_Productos " _
                  & "WHERE Producto = '" & Codigo & "' " _
                  & "AND Periodo = '" & Codigo1 & "' " _
                  & "AND Item = '" & Codigo2 & "' "
             Ejecutar_SQL_SP sSQL
             SetAddNew AdoCuentas
             For I = 0 To AdoCuentas.Recordset.fields.Count - 1
                 SetFields AdoCuentas, AdoCuentas.Recordset.fields(I).Name, VectFields(I)
             Next I
             SetUpdate AdoCuentas
          End If
         .MoveNext
       Loop
   End If
  End With
  
 'Codigos de Procesos de Libretas
  sSQL = "SELECT * " _
       & "FROM Catalogo_Proceso " _
       & "WHERE 0 = 1 "
  Select_Adodc AdoCuentas, sSQL
  ReDim VectFields(AdoCuentas.Recordset.fields.Count + 1) As Variant
  
  sSQL = "SELECT TP,Proceso,Count(TP) " _
       & "FROM Catalogo_Proceso " _
       & "WHERE TP <> '.' " _
       & "GROUP BY TP,Proceso " _
       & "HAVING Count(TP) > 1 "
  Select_Adodc AdoComp, sSQL
  With AdoComp.Recordset
   If .RecordCount > 0 Then
       Contador = 0
      .MoveFirst
       Do While Not .EOF
          Codigo = .fields("TP")
          Codigo1 = .fields("Proceso")
          Contador = Contador + 1
          FSeteos.Caption = "Catalogo_Proceso: " & Codigo & " -> " & Format(Contador / .RecordCount, "00%")
          sSQL = "SELECT * " _
               & "FROM Catalogo_Proceso " _
               & "WHERE TP = '" & Codigo & "' " _
               & "ORDER BY Cmds desc "
          Select_Adodc AdoCuentas, sSQL
          If AdoCuentas.Recordset.RecordCount > 0 Then
             For I = 0 To AdoCuentas.Recordset.fields.Count - 1
                 VectFields(I) = AdoCuentas.Recordset.fields(I)
             Next I
             sSQL = "DELETE * " _
                  & "FROM Catalogo_Proceso " _
                  & "WHERE TP = '" & Codigo & "' "
             Ejecutar_SQL_SP sSQL
             SetAddNew AdoCuentas
             For I = 0 To AdoCuentas.Recordset.fields.Count - 1
                 SetFields AdoCuentas, AdoCuentas.Recordset.fields(I).Name, VectFields(I)
             Next I
             SetUpdate AdoCuentas
          End If
         .MoveNext
       Loop
   End If
  End With
 'Fecha de Balances
  sSQL = "SELECT * " _
       & "FROM Fechas_Balance " _
       & "WHERE 0 = 1 "
  Select_Adodc AdoCuentas, sSQL
  ReDim VectFields(AdoCuentas.Recordset.fields.Count + 1) As Variant
  
  sSQL = "SELECT Periodo,Item,Detalle,Count(Detalle) " _
       & "FROM Fechas_Balance " _
       & "WHERE Detalle <> '.' " _
       & "GROUP BY Periodo,Item,Detalle " _
       & "HAVING Count(Detalle) > 1 "
  Select_Adodc AdoComp, sSQL
  With AdoComp.Recordset
   If .RecordCount > 0 Then
       Contador = 0
      .MoveFirst
       Do While Not .EOF
          Codigo = .fields("Periodo")
          Codigo1 = .fields("Item")
          Codigo2 = .fields("Detalle")
          Contador = Contador + 1
          FSeteos.Caption = "Catalogo_Proceso: " & Codigo2 & " -> " & Format(Contador / .RecordCount, "00%")
          sSQL = "SELECT * " _
               & "FROM Fechas_Balance " _
               & "WHERE Detalle = '" & Codigo2 & "' " _
               & "AND Periodo = '" & Codigo & "' " _
               & "AND Item = '" & Codigo1 & "' " _
               & "ORDER BY Fecha_Inicial "
          Select_Adodc AdoCuentas, sSQL
          If AdoCuentas.Recordset.RecordCount > 0 Then
             For I = 0 To AdoCuentas.Recordset.fields.Count - 1
                 VectFields(I) = AdoCuentas.Recordset.fields(I)
             Next I
             sSQL = "DELETE * " _
                  & "FROM Fechas_Balance " _
                  & "WHERE Detalle = '" & Codigo2 & "' " _
                  & "AND Periodo = '" & Codigo & "' " _
                  & "AND Item = '" & Codigo1 & "' "
             Ejecutar_SQL_SP sSQL
             SetAddNew AdoCuentas
             For I = 0 To AdoCuentas.Recordset.fields.Count - 1
                 SetFields AdoCuentas, AdoCuentas.Recordset.fields(I).Name, VectFields(I)
             Next I
             SetUpdate AdoCuentas
          End If
         .MoveNext
       Loop
   End If
  End With
  
 'Formato Propio
  sSQL = "SELECT * " _
       & "FROM Formato_Propio " _
       & "WHERE 0 = 1 "
  Select_Adodc AdoCuentas, sSQL
  ReDim VectFields(AdoCuentas.Recordset.fields.Count + 1) As Variant
  
  sSQL = "SELECT Item,TP,Texto,Tipo_Objeto,Codigo,COUNT(Codigo) " _
       & "FROM Formato_Propio " _
       & "WHERE Codigo <> '.' " _
       & "GROUP BY Item,TP,Texto,Tipo_Objeto,Codigo " _
       & "HAVING COUNT(Codigo) > 1 "
  Select_Adodc AdoComp, sSQL
  With AdoComp.Recordset
   If .RecordCount > 0 Then
       Contador = 0
      .MoveFirst
       Do While Not .EOF
          TipoCta = .fields("TP")
          Codigo = .fields("Texto")
          Codigo1 = .fields("Item")
          Codigo2 = .fields("Codigo")
          Codigo3 = .fields("Tipo_Objeto")
          Contador = Contador + 1
          FSeteos.Caption = "Catalogo_Proceso: " & Codigo2 & " -> " & Format(Contador / .RecordCount, "00%")
          sSQL = "SELECT * " _
               & "FROM Formato_Propio " _
               & "WHERE Texto = '" & Codigo & "' " _
               & "AND Item = '" & Codigo1 & "' " _
               & "AND Codigo = '" & Codigo2 & "' " _
               & "AND Tipo_Objeto = '" & Codigo3 & "' " _
               & "AND TP = '" & TipoCta & "' " _
               & "ORDER BY Texto "
          Select_Adodc AdoCuentas, sSQL
          If AdoCuentas.Recordset.RecordCount > 0 Then
             For I = 0 To AdoCuentas.Recordset.fields.Count - 1
                 VectFields(I) = AdoCuentas.Recordset.fields(I)
             Next I
             sSQL = "DELETE * " _
                  & "FROM Formato_Propio " _
                  & "WHERE Texto = '" & Codigo & "' " _
                  & "AND Item = '" & Codigo1 & "' " _
                  & "AND Codigo = '" & Codigo2 & "' " _
                  & "AND Tipo_Objeto = '" & Codigo3 & "' " _
                  & "AND TP = '" & TipoCta & "' "
             Ejecutar_SQL_SP sSQL
             SetAddNew AdoCuentas
             For I = 0 To AdoCuentas.Recordset.fields.Count - 1
                 SetFields AdoCuentas, AdoCuentas.Recordset.fields(I).Name, VectFields(I)
             Next I
             SetUpdate AdoCuentas
          End If
         .MoveNext
       Loop
   End If
  End With
 
 'Catalogo_Cursos
  sSQL = "SELECT * " _
       & "FROM Catalogo_Cursos " _
       & "WHERE 0 = 1 "
  Select_Adodc AdoCuentas, sSQL
  ReDim VectFields(AdoCuentas.Recordset.fields.Count + 1) As Variant
  
  sSQL = "SELECT Periodo,Item,Curso,COUNT(Curso) As NumItem " _
       & "FROM Catalogo_Cursos " _
       & "WHERE Item <> '.' " _
       & "GROUP BY Periodo,Item,Curso " _
       & "HAVING COUNT(Curso) > 1 "
  Select_Adodc AdoComp, sSQL
  With AdoComp.Recordset
   If .RecordCount > 0 Then
       Contador = 0
      .MoveFirst
       Do While Not .EOF
          Codigo = .fields("Curso")
          Codigo1 = .fields("Periodo")
          Codigo2 = .fields("Item")
          Contador = Contador + 1
          FSeteos.Caption = "Catalogo_Cursos: " & Codigo & " -> " & Format(Contador / .RecordCount, "00%")
          sSQL = "SELECT * " _
               & "FROM Catalogo_Cursos " _
               & "WHERE Curso = '" & Codigo & "' " _
               & "AND Periodo = '" & Codigo1 & "' " _
               & "AND Item = '" & Codigo2 & "' "
          Select_Adodc AdoCuentas, sSQL
          If AdoCuentas.Recordset.RecordCount > 0 Then
             For I = 0 To AdoCuentas.Recordset.fields.Count - 1
                 VectFields(I) = AdoCuentas.Recordset.fields(I)
             Next I
             sSQL = "DELETE * " _
                  & "FROM Catalogo_Cursos " _
                  & "WHERE Curso = '" & Codigo & "' " _
                  & "AND Periodo = '" & Codigo1 & "' " _
                  & "AND Item = '" & Codigo2 & "' "
             Ejecutar_SQL_SP sSQL
             SetAddNew AdoCuentas
             For I = 0 To AdoCuentas.Recordset.fields.Count - 1
                 SetFields AdoCuentas, AdoCuentas.Recordset.fields(I).Name, VectFields(I)
             Next I
             SetUpdate AdoCuentas
          End If
         .MoveNext
       Loop
   End If
  End With
   
 'Catalogo_Prestamo
  sSQL = "SELECT * " _
       & "FROM Catalogo_Prestamo " _
       & "WHERE 0 = 1 "
  Select_Adodc AdoCuentas, sSQL
  ReDim VectFields(AdoCuentas.Recordset.fields.Count + 1) As Variant
  
  sSQL = "SELECT Periodo, Item, Codigo, Descripcion, COUNT(Descripcion) As NumItem " _
       & "FROM Catalogo_Prestamo " _
       & "WHERE Item <> '.' " _
       & "GROUP BY Periodo, Item, Codigo, Descripcion " _
       & "HAVING COUNT(Descripcion) > 1 "
  Select_Adodc AdoComp, sSQL
  With AdoComp.Recordset
   If .RecordCount > 0 Then
       Contador = 0
      .MoveFirst
       Do While Not .EOF
          Codigo1 = .fields("Periodo")
          Codigo2 = .fields("Item")
          Codigo3 = .fields("Codigo")
          Codigo4 = .fields("Descripcion")
          Contador = Contador + 1
          FSeteos.Caption = "Catalogo_Cursos: " & Codigo & " -> " & Format(Contador / .RecordCount, "00%")
          sSQL = "SELECT * " _
               & "FROM Catalogo_Prestamo " _
               & "WHERE Periodo = '" & Codigo1 & "' " _
               & "AND Item = '" & Codigo2 & "' " _
               & "AND Codigo = '" & Codigo3 & "' " _
               & "AND Descripcion = '" & Codigo4 & "' "
          Select_Adodc AdoCuentas, sSQL
          If AdoCuentas.Recordset.RecordCount > 0 Then
             For I = 0 To AdoCuentas.Recordset.fields.Count - 1
                 VectFields(I) = AdoCuentas.Recordset.fields(I)
             Next I
             sSQL = "DELETE * " _
                  & "FROM Catalogo_Prestamo " _
                  & "WHERE Periodo = '" & Codigo1 & "' " _
                  & "AND Item = '" & Codigo2 & "' " _
                  & "AND Codigo = '" & Codigo3 & "' " _
                  & "AND Descripcion = '" & Codigo4 & "' "
             Ejecutar_SQL_SP sSQL
             SetAddNew AdoCuentas
             For I = 0 To AdoCuentas.Recordset.fields.Count - 1
                 SetFields AdoCuentas, AdoCuentas.Recordset.fields(I).Name, VectFields(I)
             Next I
             SetUpdate AdoCuentas
          End If
         .MoveNext
       Loop
   End If
  End With
   
 'Eliminar CxC,CxP,CxI,CxE que esten mal asignadas
  sSQL = "DELETE * " _
       & "FROM Catalogo_CxCxP " _
       & "WHERE Cta = '.' "
  Ejecutar_SQL_SP sSQL
  
  sSQL = "DELETE * " _
       & "FROM Catalogo_SubCtas " _
       & "WHERE Detalle = '.' "
  Ejecutar_SQL_SP sSQL
  
  sSQL = "DELETE * " _
       & "FROM Catalogo_Cuentas " _
       & "WHERE Codigo = '.' "
  Ejecutar_SQL_SP sSQL
  
  sSQL = "DELETE * " _
       & "FROM Catalogo_Productos " _
       & "WHERE Codigo_Inv = '.' "
  Ejecutar_SQL_SP sSQL
'''
''' 'Reprocesar Descuento/Becas de Facturas
'''  sSQL = "UPDATE Facturas " _
'''       & "SET Descuento = 0 " _
'''       & "WHERE Descuento <> 0 "
'''  Ejecutar_SQL_SP sSQL
'''  sSQL = "SELECT TC,Factura,SUM(Total_Desc) As TotalDesc " _
'''       & "FROM Detalle_Factura " _
'''       & "WHERE Total_Desc <> 0 " _
'''       & "AND Item = '" & NumEmpresa & "' " _
'''       & "GROUP BY TC,Factura "
'''  Select_Adodc AdoComp, sSQL
'''  With AdoComp.Recordset
'''   If .RecordCount > 0 Then
'''       Contador = 0
'''       Do While Not .EOF
'''          Factura_No = .Fields("Factura")
'''          'CodigoC = .Fields("CodigoC")
'''          Total = .Fields("TotalDesc")
'''          TipoDoc = .Fields("TC")
'''          'MsgBox Total
'''          sSQL = "UPDATE Facturas " _
'''               & "SET Descuento = " & Total & " " _
'''               & "WHERE Factura = " & Factura_No & " " _
'''               & "AND Item = '" & NumEmpresa & "' " _
'''               & "AND TC = '" & TipoDoc & "' "
'''          Ejecutar_SQL_SP sSQL
'''
'''          sSQL = "UPDATE Facturas " _
'''               & "SET Total_MN = Sin_IVA + Con_IVA - Descuento + IVA, SubTotal = Sin_IVA + Con_IVA " _
'''               & "WHERE Factura = " & Factura_No & " " _
'''               & "AND Item = '" & NumEmpresa & "' " _
'''               & "AND TC = '" & TipoDoc & "' "
'''          Ejecutar_SQL_SP sSQL
'''
'''          FSeteos.Caption = "Detalle_Factura: " & Factura_No & " -> " & Format(Contador / .RecordCount, "00%")
'''          Contador = Contador + 1
'''         .MoveNext
'''       Loop
'''   End If
'''  End With

  RatonNormal
End Sub

Private Sub DCUsuario_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  If CtrlDown And KeyCode = vbKeyDelete Then
     Control_Procesos Normal, "Activar Usuario: " & DCUsuario
     Codigo = Codigo_Usuario(DCUsuario)
     sSQL = "DELETE * " _
          & "FROM Acceso_Empresa " _
          & "WHERE Codigo = '" & Codigo & "' " _
          & "AND Item = '" & NumEmpresa & "' "
     Ejecutar_SQL_SP sSQL
    'Activar Usuario
     sSQL = "UPDATE Accesos " _
         & "SET TODOS = " & Val(adTrue) & " " _
         & "WHERE Codigo = '" & Codigo & "' "
     Ejecutar_SQL_SP sSQL
     MsgBox "Usuario Reactivado"
  End If
  PresionoEnter KeyCode
End Sub

Private Sub DCUsuario_LostFocus()
Dim IdMod As Integer
   For IdMod = 0 To LstModulos.ListCount - 1
       LstModulos.Selected(IdMod) = False
   Next IdMod
   For IdMod = 0 To LstEmpresas.ListCount - 1
       LstEmpresas.Selected(IdMod) = False
   Next IdMod
   For I = 1 To 7
      CNivel(I) = True
   Next I
   CSupervisor = False
   Codigo = Codigo_Usuario(DCUsuario)
   sSQL = "SELECT * " _
        & "FROM Accesos " _
        & "WHERE Codigo = '" & Codigo & "' "
   Select_Adodc AdoAux, sSQL
   If AdoAux.Recordset.RecordCount > 0 Then
      TextClave = AdoAux.Recordset.fields("Clave")
      TxtUsuario = AdoAux.Recordset.fields("Usuario")
      Codigo1 = AdoAux.Recordset.fields("CodBod")
      CNivel(1) = AdoAux.Recordset.fields("Nivel_1")
      CNivel(2) = AdoAux.Recordset.fields("Nivel_2")
      CNivel(3) = AdoAux.Recordset.fields("Nivel_3")
      CNivel(4) = AdoAux.Recordset.fields("Nivel_4")
      CNivel(5) = AdoAux.Recordset.fields("Nivel_5")
      CNivel(6) = AdoAux.Recordset.fields("Nivel_6")
      CNivel(7) = AdoAux.Recordset.fields("Nivel_7")
      CSupervisor = AdoAux.Recordset.fields("Supervisor")
      
      If AdoAux.Recordset.fields("Primaria") Then OpcP.value = True
      If AdoAux.Recordset.fields("Secundaria") Then OpcS.value = True
      If AdoAux.Recordset.fields("Bachillerato") Then OpcB.value = True
   Else
      TextClave.Text = ""
   End If
      
   With AdoBodega.Recordset
    If Codigo1 <> Ninguno And .RecordCount > 0 Then
       .MoveFirst
       .Find ("CodBod = '" & Codigo1 & "' ")
        If Not .EOF Then DCBodega = .fields("Bodega")
    End If
   End With
      
   sSQL = "SELECT AE.*,M.Aplicacion,E.Empresa " _
        & "FROM Acceso_Empresa As AE,Modulos As M,Empresas As E " _
        & "WHERE AE.Codigo = '" & Codigo & "' " _
        & "AND AE.Item = E.Item " _
        & "AND AE.Modulo = M.Modulo " _
        & "ORDER BY AE.Modulo,AE.Item "
   Select_Adodc AdoAux, sSQL
   With AdoAux.Recordset
    If .RecordCount > 0 Then
        Do While Not .EOF
           For IdMod = 0 To LstModulos.ListCount - 1
               If LstModulos.List(IdMod) = .fields("Aplicacion") Then
                  LstModulos.Selected(IdMod) = True
                  IdMod = LstModulos.ListCount
               End If
           Next IdMod
           For IdMod = 0 To LstEmpresas.ListCount - 1
               If LstEmpresas.List(IdMod) = .fields("Empresa") Then
                  LstEmpresas.Selected(IdMod) = True
                  IdMod = LstEmpresas.ListCount
               End If
           Next IdMod
          .MoveNext
        Loop
       .MoveFirst
        For I = 1 To 7
           If CNivel(I) <> 0 Then CNivel(I) = 1
        Next I
        If CSupervisor <> 0 Then CSupervisor = 1
    Else
        MsgBox "USUARIO SIN PERSONALIZACION" & vbCrLf _
             & "DE PRIVILEGIOS, TIENE ACCESO" & vbCrLf _
             & "A TODOS LOS PROCESOS."
    End If
   End With
   For I = 1 To 7
       If CNivel(I) Then CheckNivel(I - 1).value = 1 Else CheckNivel(I - 1).value = 0
   Next I
   If CSupervisor <> 0 Then CSupervisor = 1
   CheckSupervisor.value = CSupervisor
   CheckNivel(0).SetFocus
End Sub

Private Sub DGSeteosPRN_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  If CtrlDown And KeyCode = vbKeyP Then ImprimirAdodc AdoSeteosPRN, 1, 9, True
  If KeyCode = vbKeyF1 Then
     DGSeteosPRN.Visible = False
     GenerarDataTexto FSeteos, AdoSeteosPRN
     DGSeteosPRN.Visible = True
  End If
  
  If KeyCode = vbKeyReturn Then
     AdoSeteosPRN.Recordset.MoveNext
     If AdoSeteosPRN.Recordset.EOF Then AdoSeteosPRN.Recordset.MoveFirst
  End If
End Sub

Private Sub Form_Activate()
Dim AdoCon1 As ADODB.Connection
Dim RstSchema As ADODB.Recordset
Dim IdTime As Long
Dim strCnn As String
  RatonReloj
  Contador = 0
  LstTablas.Clear
  LstModulos.Clear
  LstEmpresas.Clear
  sSQL = "SELECT * " _
       & "FROM Modulos " _
       & "WHERE Modulo <> '" & Ninguno & "' " _
       & "ORDER BY Aplicacion "
  Select_Adodc AdoModulos, sSQL
  With AdoModulos.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          LstModulos.AddItem .fields("Aplicacion")
         .MoveNext
       Loop
   End If
  End With
  sSQL = "SELECT Empresa,Item " _
       & "FROM Empresas " _
       & "WHERE Item <> '" & Ninguno & "' " _
       & "ORDER BY Empresa,Item "
  Select_Adodc AdoModulos, sSQL
  With AdoModulos.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          LstEmpresas.AddItem .fields("Empresa")
         .MoveNext
       Loop
   End If
  End With
' Crea variables de objeto para los objetos de acceso a datos.
    Dim itmX As ListItem
    Set AdoCon1 = New ADODB.Connection
    AdoCon1.open AdoStrCnn
    Set RstSchema = AdoCon1.OpenSchema(adSchemaTables)
    Do Until RstSchema.EOF
       If RstSchema!TABLE_TYPE = "TABLE" And MidStrg(RstSchema!TABLE_NAME, 1, 1) <> "~" Then
         'Llenamos la lista de Tablas
          LstTablas.AddItem RstSchema!TABLE_NAME
          Contador = Contador + 1
       End If
       RstSchema.MoveNext
    Loop
'=======================================================================
  sSQL = "SELECT Nombre_Completo,Clave " _
       & "FROM Accesos " _
       & "WHERE MidStrg(Codigo,1,6) = 'ACCESO' " _
       & "ORDER BY Codigo "
  Select_Adodc_Grid DGEmp1, AdoEmp1, sSQL
  
' Lista de Usuarios en la red
  sSQL = "SELECT Nombre_Completo,Codigo " _
       & "FROM Accesos " _
       & "WHERE MidStrg(Codigo,1,6) <> 'ACCESO' " _
       & "ORDER BY Nombre_Completo,Codigo "
  SelectDB_Combo DCUsuario, AdoUsuario, sSQL, "Nombre_Completo"
  
  sSQL = "SELECT Concepto,Numero,Item,Periodo " _
       & "FROM Codigos " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "ORDER BY Concepto "
  Select_Adodc_Grid DGCodigos, AdoCodigos, sSQL
  
  sSQL = "SELECT Detalle,Codigo,DC,T_No,Item,Periodo " _
       & "FROM Ctas_Proceso " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "ORDER BY Detalle,T_No "
  Select_Adodc_Grid DGCuentas, AdoCuentas, sSQL
  
  sSQL = "SELECT TP,Lineas,Item " _
       & "FROM Formato " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "ORDER BY TP "
  Select_Adodc_Grid DGFormato, AdoFormato, sSQL
  
  sSQL = "SELECT * " _
       & "FROM Catalogo_Prestamo " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "ORDER BY Descripcion "
  Select_Adodc_Grid DGTipoPrest, AdoTipoPrest, sSQL
  
  sSQL = "SELECT * " _
       & "FROM Catalogo_Bodegas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "ORDER BY Bodega "
  SelectDB_Combo DCBodega, AdoBodega, sSQL, "Bodega"
  
  If NombreUsuario = "Administrador de Red" Then
     TabSeteos.TabVisible(2) = True
     Label5.Visible = True
     Label10.Visible = True
     TxtUsuario.Visible = True
     TextClave.Visible = True
     Command5.Enabled = True
     Command6.Visible = True
     Command7.Visible = True
     Command13.Visible = True
     Command10.Visible = True
     
''     CommandButton13.Visible = True
''     CommandButton15.Visible = True
''     CommandButton1.Visible = True
''     CommandButton4.Visible = True
''     CommandButton5.Visible = True
''     CommandButton8.Visible = True

  Else
     TabSeteos.TabVisible(2) = False
     Label5.Visible = False
     Label10.Visible = False
     TxtUsuario.Visible = False
     TextClave.Visible = False
     Command6.Visible = False
     Command7.Visible = False
     Command13.Visible = False

  End If
  
  sSQL = "SELECT * " _
       & "FROM Catalogo_Periodo_Lectivo " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' "
  Select_Adodc_Grid DGEducativo, AdoEducativo, sSQL
  
  LstDuplicados.AddItem "Duplicados de Clientes"
  LstDuplicados.AddItem "Duplicados de Usuarios"
  LstDuplicados.AddItem "Duplicados de Representantes"
  LstDuplicados.AddItem "Duplicados de Cuentas, SubModulos y Productos"
 'LstDuplicados.AddItem "Duplicados de Rubros Clientes Facturacion"
  LstDuplicados.AddItem "Duplicados de Compras y Retenciones"
  LstDuplicados.AddItem "Duplicados de Clientes Facturacion"
  LstDuplicados.AddItem "Actualizar Abreviaturas Accesos Usuarios"
  LstDuplicados.AddItem "Actualizacion de Inicio de Sección"
  LstDuplicados.AddItem "Actualizacion de CI/RUC Clientes Facturacion"
  LstDuplicados.AddItem "Actualizacion de CI/RUC de Garantes"
  LstDuplicados.AddItem "Actualizacion de Beneficiario en las bases"
  LstDuplicados.AddItem "Actualizacion de Codigos Superiores en Productos"
  LstDuplicados.AddItem "Actualizacion de Documentos Electronicos Autorizados"
  LstDuplicados.AddItem "Actualizacion Facturas y Notas de Credito en Kardex"
  LstDuplicados.AddItem "Abonar Facturas Pendientes"
  LstDuplicados.AddItem "Cambiar PV as FA"
  LstDuplicados.AddItem "Listar Movimientos con Cuentas de Grupo"
  LstDuplicados.AddItem "Renumerar CI/RUC/Pasaporte de Personas"
  LstDuplicados.AddItem "Renumerar Clave del Catalogo"
  'If Modo_Educativo Then LstDuplicados.AddItem "Borrar Datos de Estudiantes"
  LstDuplicados.AddItem "Respaldar Bases de Datos Completa"
  LstDuplicados.AddItem "Generar Documentos Electronicos"
  LstDuplicados.AddItem "Eliminar Seteos por Default (000)"
  LstDuplicados.AddItem "Eliminar Basura en Base de Datos"
  LstDuplicados.AddItem "Eliminar Indice en Base de Datos"
  LstDuplicados.AddItem "Eliminar Tablas Vacias"
  LstDuplicados.AddItem "Prueba de Envio de Correos"

  LstDuplicados.AddItem "Realizar Copia de Actualizacion"
  If CodigoUsuario = "ACCESO01" Then Command1.Enabled = True
  RatonNormal
End Sub

Private Sub Form_Load()
  CentrarForm FSeteos
' Direccion de Bases de la empresa
  ConectarAdodc AdoInv
  ConectarAdodc AdoAux
  ConectarAdodc AdoEmp
  ConectarAdodc AdoEmp1
  ConectarAdodc AdoTrans
  ConectarAdodc AdoTrans_SC
  ConectarAdodc AdoTrans_GC
  ConectarAdodc AdoComp
  ConectarAdodc AdoRet
  ConectarAdodc AdoFact
  ConectarAdodc AdoUsuario
  ConectarAdodc AdoFactura
  ConectarAdodc AdoCajaCred
  ConectarAdodc AdoBancos
  ConectarAdodc AdoCodigos
  ConectarAdodc AdoCuentas
  ConectarAdodc AdoSeteosPRN
  ConectarAdodc AdoSetPRN
  ConectarAdodc AdoTipoPrest
  ConectarAdodc AdoPrueba
  ConectarAdodc AdoBodega
  ConectarAdodc AdoFormato
  ConectarAdodc AdoModulos
  ConectarAdodc AdoEducativo
  ConectarAdodc AdoAutorizacion
End Sub

Private Sub TabSeteos_Click(PreviousTab As Integer)
  'If TabSeteos.Tab = 1 Then TextEmpresa.SetFocus
End Sub

Public Sub BorrarBaseDeAdodc(NombreTabla As String, Optional PorFecha As Boolean)
Dim JCamp As Integer
Dim Item_Ok As Boolean
Dim Fecha_Ok As Boolean
Dim Where_Ok As Boolean
  If Len(NombreTabla) > 1 Then
     FSeteos.Caption = "Borrando la Base: " & NombreTabla & "..."
     Item_Ok = False
     Fecha_Ok = False
     Where_Ok = False
     sSQL = "SELECT " & Full_Fields(NombreTabla) & " " _
          & "FROM " & TrimStrg(NombreTabla) & " " _
          & "WHERE 1 = 0 "
     Select_Adodc AdoAux, sSQL
     For JCamp = 0 To AdoAux.Recordset.fields.Count - 1
         If AdoAux.Recordset.fields(JCamp).Name = "Item" Then Item_Ok = True
         If AdoAux.Recordset.fields(JCamp).Name = "Fecha" Then Fecha_Ok = True
     Next JCamp
     sSQL = "DELETE * " _
          & "FROM " & TrimStrg(NombreTabla) & " "
     If Item_Ok Then
        sSQL = sSQL & "WHERE Item = '" & NumEmpresa & "' "
        Where_Ok = True
     End If
     If Fecha_Ok Then
        If Where_Ok Then
           If PorFecha Then sSQL = sSQL & "AND Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# "
        Else
           If PorFecha Then sSQL = sSQL & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# "
        End If
     End If
     Ejecutar_SQL_SP sSQL
  End If
End Sub

Public Sub CambiarNumeroAdodc(NombreTabla As String, ItemCambio As String)
  If Len(NombreTabla) > 1 Then
     FSeteos.Caption = "Cambiando la Base: " & NombreTabla & "..."
     sSQL = "UPDATE " & TrimStrg(NombreTabla) & " " _
          & "SET Item = '" & ItemCambio & "' "
     If NombreTabla = "Empresas" Then sSQL = sSQL & ", Grupo = '" & ItemCambio & "' "
     sSQL = sSQL & "WHERE Item = '" & NumEmpresa & "' "
     Ejecutar_SQL_SP sSQL
  End If
End Sub

Public Sub BorrarBasesDatos(BorrarPorFecha As Boolean)
    If BorrarPorFecha = False Then
       BorrarBaseDeAdodc "Catalogo_Auditor", BorrarPorFecha
       BorrarBaseDeAdodc "Catalogo_Cuentas", BorrarPorFecha
       BorrarBaseDeAdodc "Catalogo_CxCxP", BorrarPorFecha
       BorrarBaseDeAdodc "Catalogo_Productos", BorrarPorFecha
       BorrarBaseDeAdodc "Catalogo_Rol_Pagos", BorrarPorFecha
       BorrarBaseDeAdodc "Catalogo_SubCtas", BorrarPorFecha
       BorrarBaseDeAdodc "Catalogo_Prestamo", BorrarPorFecha
       BorrarBaseDeAdodc "Catalogo_Materias", BorrarPorFecha
       BorrarBaseDeAdodc "Catalogo_Lineas", BorrarPorFecha
       BorrarBaseDeAdodc "Catalogo_Habitacion", BorrarPorFecha
       BorrarBaseDeAdodc "Catalogo_Estudiantil", BorrarPorFecha
       BorrarBaseDeAdodc "Catalogo_Corresponsal", BorrarPorFecha
       BorrarBaseDeAdodc "Clientes_Facturacion", BorrarPorFecha
    End If
    BorrarBaseDeAdodc "Ctas_Proceso"
    BorrarBaseDeAdodc "Balance_Consolidado"
    BorrarBaseDeAdodc "Seteos_Documentos"
    BorrarBaseDeAdodc "Formato"
   'BorrarBaseDeAdodc "Empresas"
    BorrarBaseDeAdodc "Acceso_Empresa", BorrarPorFecha
   'BorrarBaseDeAdodc "Accesos", BorrarPorFecha
   'BorrarBaseDeAdodc "Clientes", BorrarPorFecha
    BorrarBaseDeAdodc "Clientes_Datos_Extras", BorrarPorFecha
    BorrarBaseDeAdodc "Comprobantes", BorrarPorFecha
    BorrarBaseDeAdodc "Detalle_Factura", BorrarPorFecha
    BorrarBaseDeAdodc "Detalle_Nota_Credito", BorrarPorFecha
    BorrarBaseDeAdodc "Facturas", BorrarPorFecha
    BorrarBaseDeAdodc "Prestamos", BorrarPorFecha
    BorrarBaseDeAdodc "Saldo_Diarios", BorrarPorFecha
    BorrarBaseDeAdodc "Trans_Abonos", BorrarPorFecha
    BorrarBaseDeAdodc "Trans_Air", BorrarPorFecha
    BorrarBaseDeAdodc "Trans_Aplicacion", BorrarPorFecha
    BorrarBaseDeAdodc "Trans_Bloqueos", BorrarPorFecha
    BorrarBaseDeAdodc "Trans_Compras", BorrarPorFecha
    BorrarBaseDeAdodc "Trans_Abonos", BorrarPorFecha
    BorrarBaseDeAdodc "Trans_Entrada_Salida", BorrarPorFecha
    BorrarBaseDeAdodc "Trans_Gastos_Caja", BorrarPorFecha
    BorrarBaseDeAdodc "Trans_Kardex", BorrarPorFecha
    BorrarBaseDeAdodc "Trans_Libretas", BorrarPorFecha
    BorrarBaseDeAdodc "Trans_Prestamos", BorrarPorFecha
    BorrarBaseDeAdodc "Trans_Presupuestos"
    BorrarBaseDeAdodc "Trans_Rol_Horas", BorrarPorFecha
    BorrarBaseDeAdodc "Trans_Rol_Pagos", BorrarPorFecha
    BorrarBaseDeAdodc "Trans_SubCtas", BorrarPorFecha
    BorrarBaseDeAdodc "Trans_Suscripciones", BorrarPorFecha
    BorrarBaseDeAdodc "Trans_Conciliacion", BorrarPorFecha
    BorrarBaseDeAdodc "Trans_Ventas", BorrarPorFecha
    BorrarBaseDeAdodc "Transacciones", BorrarPorFecha
    BorrarBaseDeAdodc "Trans_Notas"
End Sub

Private Sub TxtItem_GotFocus()
  MarcarTexto TxtItem
End Sub

Private Sub TxtItem_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtItem_LostFocus()
  TextoValido TxtItem, True, , 0
  If Val(TxtItem) <= 0 Then TxtItem = "999"
  TxtItem = Format(Val(TxtItem), "000")
End Sub

Public Sub Cambio_De_Codigo(NombreTabla As String, _
                            Campo As String, _
                            CodOld As String, _
                            CodNew As String)
  If CodNew <> CodOld Then
     sSQL = "UPDATE " & NombreTabla & " " _
          & "SET " & Campo & " = '" & CodNew & "' " _
          & "WHERE " & Campo & " = '" & CodOld & "' "
     Ejecutar_SQL_SP sSQL
    'MsgBox "Actualizar: " & sSQL
  End If
End Sub

'''Public Sub ActualizarCodigoCliente(CodigoNew As String, _
'''                                   CodigoOld As String, _
'''                                   Optional Client As Boolean)
'''   If Client Then
'''      sSQL = "UPDATE Clientes " _
'''           & "SET Codigo = '" & CodigoNew & "'," _
'''           & "TD = '" & TipoBenef & "' " _
'''           & "WHERE Codigo = '" & CodigoOld & "' "
'''      Ejecutar_SQL_SP sSQL
'''   End If
'''
'''   sSQL = "UPDATE Comprobantes " _
'''        & "SET Codigo_B = '" & CodigoNew & "' " _
'''        & "WHERE Codigo_B = '" & CodigoOld & "' "
'''   Ejecutar_SQL_SP sSQL
'''
'''   sSQL = "UPDATE Catalogo_CxCxP " _
'''        & "SET Codigo = '" & CodigoNew & "' " _
'''        & "WHERE Codigo = '" & CodigoOld & "' "
'''   Ejecutar_SQL_SP sSQL
'''
'''   sSQL = "UPDATE Catalogo_SubCtas " _
'''        & "SET Codigo = '" & CodigoNew & "' " _
'''        & "WHERE Codigo = '" & CodigoOld & "' "
'''   Ejecutar_SQL_SP sSQL
'''
'''   sSQL = "UPDATE Catalogo_Rol_Pagos " _
'''        & "SET Codigo = '" & CodigoNew & "' " _
'''        & "WHERE Codigo = '" & CodigoOld & "' "
'''   Ejecutar_SQL_SP sSQL
'''
'''   sSQL = "UPDATE Detalle_Factura " _
'''        & "SET CodigoC = '" & CodigoNew & "' " _
'''        & "WHERE CodigoC = '" & CodigoOld & "' "
'''   Ejecutar_SQL_SP sSQL
'''
'''   sSQL = "UPDATE Facturas " _
'''        & "SET CodigoC = '" & CodigoNew & "' " _
'''        & "WHERE CodigoC = '" & CodigoOld & "' "
'''   Ejecutar_SQL_SP sSQL
'''
'''   sSQL = "UPDATE Clientes_Matriculas " _
'''        & "SET Codigo = '" & CodigoNew & "' " _
'''        & "WHERE Codigo = '" & CodigoOld & "' "
'''   Ejecutar_SQL_SP sSQL
'''
'''   sSQL = "UPDATE Clientes_Facturacion " _
'''        & "SET Codigo = '" & CodigoNew & "' " _
'''        & "WHERE Codigo = '" & CodigoOld & "' "
'''   Ejecutar_SQL_SP sSQL
'''
'''   sSQL = "UPDATE Trans_Abonos " _
'''        & "SET CodigoC = '" & CodigoNew & "' " _
'''        & "WHERE CodigoC = '" & CodigoOld & "' "
'''   Ejecutar_SQL_SP sSQL
'''
'''   sSQL = "UPDATE Trans_Aduanas " _
'''        & "SET CodigoC = '" & CodigoNew & "' " _
'''        & "WHERE CodigoC = '" & CodigoOld & "' "
'''   Ejecutar_SQL_SP sSQL
'''
'''   sSQL = "UPDATE Trans_Comision " _
'''        & "SET CodigoC = '" & CodigoNew & "' " _
'''        & "WHERE CodigoC = '" & CodigoOld & "' "
'''   Ejecutar_SQL_SP sSQL
'''
'''   sSQL = "UPDATE Trans_Cuotas " _
'''        & "SET CodigoC = '" & CodigoNew & "' " _
'''        & "WHERE CodigoC = '" & CodigoOld & "' "
'''   Ejecutar_SQL_SP sSQL
'''
'''   sSQL = "UPDATE Trans_Kardex " _
'''        & "SET Codigo_P = '" & CodigoNew & "' " _
'''        & "WHERE Codigo_P = '" & CodigoOld & "' "
'''   Ejecutar_SQL_SP sSQL
'''
'''   sSQL = "UPDATE Trans_Actas " _
'''        & "SET Codigo = '" & CodigoNew & "' " _
'''        & "WHERE Codigo = '" & CodigoOld & "' "
'''   Ejecutar_SQL_SP sSQL
'''
'''   sSQL = "UPDATE Trans_Notas " _
'''        & "SET Codigo = '" & CodigoNew & "' " _
'''        & "WHERE Codigo = '" & CodigoOld & "' "
'''   Ejecutar_SQL_SP sSQL
'''
'''   sSQL = "UPDATE Trans_Pedidos " _
'''        & "SET CodigoC = '" & CodigoNew & "' " _
'''        & "WHERE CodigoC = '" & CodigoOld & "' "
'''   Ejecutar_SQL_SP sSQL
'''
'''   sSQL = "UPDATE Trans_Rol_Horas " _
'''        & "SET Codigo = '" & CodigoNew & "' " _
'''        & "WHERE Codigo = '" & CodigoOld & "' "
'''   Ejecutar_SQL_SP sSQL
'''
'''   sSQL = "UPDATE Trans_Rol_Pagos " _
'''        & "SET Codigo = '" & CodigoNew & "' " _
'''        & "WHERE Codigo = '" & CodigoOld & "' "
'''   Ejecutar_SQL_SP sSQL
'''
'''   sSQL = "UPDATE Trans_SubCtas " _
'''        & "SET Codigo = '" & CodigoNew & "' " _
'''        & "WHERE Codigo = '" & CodigoOld & "' "
'''   Ejecutar_SQL_SP sSQL
'''
'''   sSQL = "UPDATE Prestamos " _
'''        & "SET Cuenta_No = '" & CodigoNew & "' " _
'''        & "WHERE Cuenta_No = '" & CodigoOld & "' "
'''   Ejecutar_SQL_SP sSQL
'''
'''   sSQL = "UPDATE Trans_Prestamos " _
'''        & "SET Cuenta_No = '" & CodigoNew & "' " _
'''        & "WHERE Cuenta_No = '" & CodigoOld & "' "
'''   Ejecutar_SQL_SP sSQL
'''
'''   sSQL = "UPDATE Transacciones " _
'''        & "SET Codigo_C = '" & CodigoNew & "' " _
'''        & "WHERE Codigo_C = '" & CodigoOld & "' "
'''   Ejecutar_SQL_SP sSQL
'''
'''   sSQL = "UPDATE Trans_Air " _
'''        & "SET IdProv = '" & CodigoNew & "' " _
'''        & "WHERE IdProv = '" & CodigoOld & "' "
'''   Ejecutar_SQL_SP sSQL
'''
'''   sSQL = "UPDATE Trans_Compras " _
'''        & "SET IdProv = '" & CodigoNew & "' " _
'''        & "WHERE IdProv = '" & CodigoOld & "' "
'''   Ejecutar_SQL_SP sSQL
'''
'''   sSQL = "UPDATE Trans_Ventas " _
'''        & "SET IdProv = '" & CodigoNew & "' " _
'''        & "WHERE IdProv = '" & CodigoOld & "' "
'''   Ejecutar_SQL_SP sSQL
'''
'''   sSQL = "UPDATE Trans_Exportaciones " _
'''        & "SET IdFiscalProv = '" & CodigoNew & "' " _
'''        & "WHERE IdFiscalProv = '" & CodigoOld & "' "
'''   Ejecutar_SQL_SP sSQL
'''
'''   sSQL = "UPDATE Trans_Importaciones " _
'''        & "SET IdFiscalProv = '" & CodigoNew & "' " _
'''        & "WHERE IdFiscalProv = '" & CodigoOld & "' "
'''   Ejecutar_SQL_SP sSQL
'''End Sub

Public Sub Actualizar_Facturas_NC_en_Kardex()
Dim AdoDBAut As ADODB.Recordset
Dim AdoDBDetFA As ADODB.Recordset
Dim AdoDBComp As ADODB.Recordset
Dim IdF As Long
Dim UnaFecha As Boolean
Dim EsCxC As Boolean
Dim TDec_Costo As Byte
'Dim Codigos_Inv() As String
Dim Mes_FA As Long

'Cierre de Caja de Cuentas por Cobrar
'Cierre de Caja de Abonos
    DatInv.Fecha_Stock = FechaSistema
    Progreso_Iniciar
    FEsperar.Show
    
'''    Progreso_Barra.Mensaje_Box = "Act. SubTotal NC"
'''    Progreso_Esperar
'''    sSQL = "UPDATE Detalle_Factura " _
'''         & "SET Cantidad_NC = ROUND(SubTotal_NC/Precio,2,0) " _
'''         & "WHERE Item = '" & NumEmpresa & "' " _
'''         & "AND Periodo = '" & Periodo_Contable & "' " _
'''         & "AND Cantidad_NC = 0 " _
'''         & "AND LEN(Serie_NC) = 6 " _
'''         & "AND Secuencial_NC > 0 " _
'''         & "AND Precio > 0 "
'''    Ejecutar_SQL_SP sSQL
        
'''    Progreso_Barra.Mensaje_Box = "Act. Series NC"
'''    Progreso_Esperar
'''    If SQL_Server Then
'''       sSQL = "UPDATE Detalle_Factura " _
'''            & "SET Fecha_NC=TA.Fecha, Serie_NC=TA.Serie_NC, Autorizacion_NC=TA.Autorizacion_NC, Secuencial_NC=TA.Secuencial_NC " _
'''            & "FROM Detalle_Factura As DF, Trans_Abonos As TA "
'''    Else
'''       sSQL = "UPDATE Detalle_Factura As DF, Trans_Abonos As TA " _
'''            & "SET DF.Fecha_NC=TA.Fecha, DF.Serie_NC=TA.Serie_NC, DF.Autorizacion_NC=TA.Autorizacion_NC, DF.Secuencial_NC=TA.Secuencial_NC "
'''    End If
'''    sSQL = sSQL _
'''         & "WHERE DF.Item = '" & NumEmpresa & "' " _
'''         & "AND DF.Periodo = '" & Periodo_Contable & "' " _
'''         & "AND TA.Banco = 'NOTA DE CREDITO' " _
'''         & "AND DF.Item = TA.Item " _
'''         & "AND DF.Periodo = TA.Periodo " _
'''         & "AND DF.TC = TA.TP " _
'''         & "AND DF.Serie = TA.Serie " _
'''         & "AND DF.Factura = TA.Factura " _
'''         & "AND DF.Autorizacion = TA.Autorizacion "
'''    Ejecutar_SQL_SP sSQL
    
    FechaIni = BuscarFecha("01/08/2019")
    Progreso_Barra.Mensaje_Box = "Determinando Cierre Caja"
    Progreso_Esperar
    Imagen_Esperar
    TDec_Costo = Dec_Costo
    If TDec_Costo > 6 Then TDec_Costo = 6
    sSQL = "UPDATE Trans_Kardex " _
         & "SET X = '.' " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' "
    Ejecutar_SQL_SP sSQL
    
    If SQL_Server Then
       sSQL = "UPDATE Trans_Kardex " _
            & "SET X = 'C' " _
            & "FROM Trans_Kardex As TK, Comprobantes As C "
    Else
       sSQL = "UPDATE Trans_Kardex As TK, Comprobantes As C " _
            & "SET DF.X = 'C' "
    End If
    sSQL = sSQL _
         & "WHERE TK.Item = '" & NumEmpresa & "' " _
         & "AND TK.Periodo = '" & Periodo_Contable & "' " _
         & "AND C.Concepto LIKE 'Cierre de Caja de%' " _
         & "AND TK.Item = C.Item " _
         & "AND TK.Periodo = C.Periodo " _
         & "AND TK.TP = C.TP " _
         & "AND TK.Numero = C.Numero " _
         & "AND TK.Fecha = C.Fecha "
    Ejecutar_SQL_SP sSQL
    
    Imagen_Esperar
    sSQL = "DELETE " _
         & "FROM Trans_Kardex " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND X = 'C' "
    Ejecutar_SQL_SP sSQL
    
    Imagen_Esperar
    sSQL = "DELETE " _
         & "FROM Trans_Kardex " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND LEN(TC) = 2 " _
         & "AND LEN(Serie) = 6 " _
         & "AND Factura <> 0 " _
         & "AND SUBSTRING(Detalle,1,3) <> 'FA:' "
    Ejecutar_SQL_SP sSQL
    
    Imagen_Esperar
    sSQL = "UPDATE Trans_Kardex " _
         & "SET Procesado = 0 " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' "
    Ejecutar_SQL_SP sSQL
    
    Mayorizar_Inventario_SP
    
    Progreso_Barra.Mensaje_Box = "Determinando Codigos con Inventario"
    Progreso_Esperar
    FEsperar.Show
    sSQL = "UPDATE Detalle_Factura " _
         & "SET X = '.', Costo = 0 " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' "
    Ejecutar_SQL_SP sSQL
    
    Imagen_Esperar
    If SQL_Server Then
       sSQL = "UPDATE Detalle_Factura " _
            & "SET X = 'I' " _
            & "FROM Detalle_Factura As DF, Catalogo_Productos As CP "
    Else
       sSQL = "UPDATE Detalle_Factura As DF, Catalogo_Productos As CP " _
            & "SET DF.X = 'I' "
    End If
    sSQL = sSQL _
         & "WHERE DF.Item = '" & NumEmpresa & "' " _
         & "AND DF.Periodo = '" & Periodo_Contable & "' " _
         & "AND CP.TC = 'P' " _
         & "AND LEN(CP.Cta_Inventario) > 2 " _
         & "AND LEN(CP.Cta_Costo_Venta) > 2 " _
         & "AND DF.T <> '" & Anulado & "' " _
         & "AND DF.Item = CP.Item " _
         & "AND DF.Periodo = CP.Periodo " _
         & "AND DF.Codigo = CP.Codigo_Inv "
    Ejecutar_SQL_SP sSQL
    
    Progreso_Barra.Mensaje_Box = "Actualizando Costos"
    Progreso_Esperar
    Imagen_Esperar
    sSQL = "SELECT * " _
         & "FROM Detalle_Factura " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND Fecha >= #" & FechaIni & "# " _
         & "AND T <> '" & Anulado & "' " _
         & "AND Costo = 0 " _
         & "AND X = 'I' " _
         & "ORDER BY Fecha, TC, Serie, Factura, Codigo, ID "
    Select_AdoDB AdoDBDetFA, sSQL
    With AdoDBDetFA
     If .RecordCount > 0 Then
         Progreso_Barra.Valor_Maximo = Progreso_Barra.Valor_Maximo + .RecordCount
         Do While Not .EOF
            Imagen_Esperar
            DatInv.PVP = Redondear(.fields("Precio"), TDec_Costo)
               DatInv.Costo = 0  'DatInv.PVP
            DatInv.Fecha_Stock = .fields("Fecha")
            DatInv.Codigo_Inv = .fields("Codigo")
            DatInv.Detalle = .fields("CodigoC")
         
            Progreso_Barra.Mensaje_Box = DatInv.Fecha_Stock & ", Actualizando Costo de: " & DatInv.Codigo_Inv
            Progreso_Esperar
            sSQL = "SELECT Cliente " _
                 & "FROM Clientes " _
                 & "WHERE Codigo = '" & DatInv.Detalle & "' "
            Select_AdoDB AdoDBAut, sSQL
            If AdoDBAut.RecordCount > 0 Then DatInv.Detalle = AdoDBAut.fields("Cliente") Else DatInv.Detalle = ""
            AdoDBAut.Close
            
            Imagen_Esperar
            sSQL = "SELECT TOP 2 Costo, Valor_Unitario, Existencia " _
                 & "FROM Trans_Kardex " _
                 & "WHERE Item = '" & NumEmpresa & "' " _
                 & "AND Periodo = '" & Periodo_Contable & "' " _
                 & "AND Fecha <= #" & BuscarFecha(DatInv.Fecha_Stock) & "# " _
                 & "AND Codigo_Inv = '" & DatInv.Codigo_Inv & "' " _
                 & "ORDER BY Fecha DESC, Entrada, Salida DESC,TP DESC, Numero DESC, ID DESC "
            Select_AdoDB AdoDBAut, sSQL
            If AdoDBAut.RecordCount > 0 Then
               DatInv.Costo = AdoDBAut.fields("Costo")
               DatInv.Stock = AdoDBAut.fields("Existencia")
               AdoDBAut.MoveNext
               If Not AdoDBAut.EOF Then DatInv.Costo = (DatInv.Costo + AdoDBAut.fields("Costo")) / 2
            End If
            AdoDBAut.Close
            
           'Actualizamos el costo de Venta
           'If DatInv.Costo > DatInv.PVP Then DatInv.Costo = DatInv.PVP
            DatInv.Costo = Redondear(DatInv.Costo, TDec_Costo)
           .fields("Costo") = DatInv.Costo
           .Update
           
            Progreso_Barra.Mensaje_Box = "Insertamos FA al Kardex"
            Progreso_Esperar True
            Imagen_Esperar
            SetAdoAddNew "Trans_Kardex"
            SetAdoFields "Codigo_P", .fields("CodigoC")
            SetAdoFields "Item", .fields("Item")
            SetAdoFields "Periodo", .fields("Periodo")
            SetAdoFields "T", Normal
            SetAdoFields "CodBodega", .fields("CodBodega")
            SetAdoFields "Codigo_Inv", .fields("Codigo")
            SetAdoFields "Fecha", .fields("Fecha")
            SetAdoFields "TC", .fields("TC")
            SetAdoFields "Serie", .fields("Serie")
            SetAdoFields "Factura", .fields("Factura")
            SetAdoFields "CodigoL", .fields("CodigoL")
            SetAdoFields "Salida", .fields("Cantidad")
            SetAdoFields "CodMarca", .fields("CodMarca")
            SetAdoFields "Lote_No", .fields("Lote_No")
            SetAdoFields "Fecha_Fab", .fields("Fecha_Fab")
            SetAdoFields "Fecha_Exp", .fields("Fecha_Exp")
            SetAdoFields "Modelo", .fields("Modelo")
            SetAdoFields "Procedencia", .fields("Procedencia")
            SetAdoFields "Serie_No", .fields("Serie_No")
            SetAdoFields "Total_IVA", .fields("Total_IVA")
            SetAdoFields "Porc_C", .fields("Porc_C")
            SetAdoFields "CodigoU", .fields("CodigoU")
            SetAdoFields "PVP", .fields("Precio")
            SetAdoFields "Detalle", TrimStrg(MidStrg("FA: " & DatInv.Detalle, 1, 100))
            SetAdoFields "Existencia", DatInv.Stock - .fields("Cantidad")
            SetAdoFields "Valor_Unitario", .fields("Precio")
            SetAdoFields "Valor_Total", Redondear(.fields("Precio") * .fields("Cantidad"), 2)
            SetAdoFields "Costo", DatInv.Costo
            SetAdoFields "Total", Redondear(DatInv.Costo * .fields("Cantidad"), 2)
            SetAdoUpdate
'''            If .Fields("SubTotal_NC") <> 0 Then
'''
'''                'If .Fields("Factura") = 5778 Then MsgBox .Fields("Factura")
'''
'''                Progreso_Barra.Mensaje_Box = "Insertamos NC al Kardex"
'''                Progreso_Esperar True
'''
'''                SetAdoAddNew "Trans_Kardex"
'''                SetAdoFields "Codigo_P", .Fields("CodigoC")
'''                SetAdoFields "Item", .Fields("Item")
'''                SetAdoFields "Periodo", .Fields("Periodo")
'''                SetAdoFields "T", Normal
'''                SetAdoFields "CodBodega", .Fields("CodBodega")
'''                SetAdoFields "Codigo_Inv", .Fields("Codigo")
'''                SetAdoFields "Fecha", .Fields("Fecha")
'''                SetAdoFields "TC", .Fields("TC")
'''                SetAdoFields "Serie", .Fields("Serie")
'''                SetAdoFields "Factura", .Fields("Factura")
'''                SetAdoFields "CodigoL", .Fields("CodigoL")
'''                SetAdoFields "Entrada", .Fields("Cantidad_NC")
'''                SetAdoFields "CodMarca", .Fields("CodMarca")
'''                SetAdoFields "Lote_No", .Fields("Lote_No")
'''                SetAdoFields "Fecha_Fab", .Fields("Fecha_Fab")
'''                SetAdoFields "Fecha_Exp", .Fields("Fecha_Exp")
'''                SetAdoFields "Modelo", .Fields("Modelo")
'''                SetAdoFields "Procedencia", .Fields("Procedencia")
'''                SetAdoFields "Serie_No", .Fields("Serie_No")
'''                SetAdoFields "Total_IVA", .Fields("Total_IVA")
'''                SetAdoFields "Porc_C", .Fields("Porc_C")
'''                SetAdoFields "CodigoU", .Fields("CodigoU")
'''                SetAdoFields "PVP", .Fields("Precio")
'''                SetAdoFields "Detalle", TrimStrg(MidStrg("NC: " & DatInv.Detalle, 1, 100))
'''                SetAdoFields "Existencia", DatInv.Stock + .Fields("Cantidad_NC")
'''                SetAdoFields "Valor_Unitario", .Fields("Precio")
'''                SetAdoFields "Valor_Total", Redondear(.Fields("Precio") * .Fields("Cantidad_NC"), 2)
'''                SetAdoFields "Costo", DatInv.Costo
'''                SetAdoFields "Total", Redondear(DatInv.Costo * .Fields("Cantidad_NC"), 2)
'''                SetAdoUpdate
'''                'MsgBox Progreso_Barra.Mensaje_Box & vbCrLf & .Fields("Fecha")
'''            End If
           .MoveNext
         Loop
     End If
    End With
    AdoDBDetFA.Close
            
    Progreso_Barra.Mensaje_Box = "Ctas. de Sal. Inv."
    Progreso_Esperar
    Imagen_Esperar
    If SQL_Server Then
       sSQL = "UPDATE Trans_Kardex " _
            & "SET Cta_Inv=CP.Cta_Inventario, Contra_Cta=CP.Cta_Costo_Venta " _
            & "FROM Trans_Kardex As TK, Catalogo_Productos As CP "
    Else
       sSQL = "UPDATE Trans_Kardex As TK, Catalogo_Productos As CP " _
            & "SET TK.Cta_Inv=CP.Cta_Inventario, TK.Contra_Cta=CP.Cta_Costo_Venta "
    End If
    sSQL = sSQL _
         & "WHERE TK.Item = '" & NumEmpresa & "' " _
         & "AND TK.Periodo = '" & Periodo_Contable & "' " _
         & "AND LEN(CP.Cta_Inventario) > 2 " _
         & "AND LEN(CP.Cta_Costo_Venta) > 2 " _
         & "AND SUBSTRING(TK.Detalle,1,3) = 'FA:' " _
         & "AND TK.Item = CP.Item " _
         & "AND TK.Periodo = CP.Periodo " _
         & "AND TK.Codigo_Inv = CP.Codigo_Inv "
    Ejecutar_SQL_SP sSQL
    
    Progreso_Barra.Mensaje_Box = "Ctas. de Ent. Inv. NC"
    Progreso_Esperar
    Imagen_Esperar
    If SQL_Server Then
       sSQL = "UPDATE Trans_Kardex " _
            & "SET Cta_Inv=CP.Cta_Inventario " _
            & "FROM Trans_Kardex As TK, Catalogo_Productos As CP "
    Else
       sSQL = "UPDATE Trans_Kardex As TK, Catalogo_Productos As CP " _
            & "SET TK.Cta_Inv=CP.Cta_Inventario "
    End If
    sSQL = sSQL _
         & "WHERE TK.Item = '" & NumEmpresa & "' " _
         & "AND TK.Periodo = '" & Periodo_Contable & "' " _
         & "AND Cta_Inv = '.' " _
         & "AND LEN(CP.Cta_Inventario) > 2 " _
         & "AND LEN(CP.Cta_Costo_Venta) > 2 " _
         & "AND SUBSTRING(TK.Detalle,1,3) = 'NC:' " _
         & "AND TK.Item = CP.Item " _
         & "AND TK.Periodo = CP.Periodo " _
         & "AND TK.Codigo_Inv = CP.Codigo_Inv "
    Ejecutar_SQL_SP sSQL

    Imagen_Esperar
    If SQL_Server Then
       sSQL = "UPDATE Trans_Kardex " _
            & "SET Contra_Cta=TA.Cta " _
            & "FROM Trans_Kardex As TK, Trans_Abonos As TA "
    Else
       sSQL = "UPDATE Trans_Kardex As TK, Trans_Abonos As TA " _
            & "SET TK.Contra_Cta=TA.Cta "
    End If
    sSQL = sSQL _
         & "WHERE TK.Item = '" & NumEmpresa & "' " _
         & "AND TK.Periodo = '" & Periodo_Contable & "' " _
         & "AND TK.Contra_Cta = '.' " _
         & "AND TA.Banco = 'NOTA DE CREDITO' " _
         & "AND TA.Cheque = 'VENTAS' " _
         & "AND SUBSTRING(TK.Detalle,1,3) = 'NC:' " _
         & "AND TK.Item = TA.Item " _
         & "AND TK.Periodo = TA.Periodo " _
         & "AND TK.TC = TA.TP " _
         & "AND TK.Serie = TA.Serie " _
         & "AND TK.Factura = TA.Factura "
    Ejecutar_SQL_SP sSQL
        
    Progreso_Barra.Incremento = 0
    Progreso_Barra.Valor_Maximo = 100
    Progreso_Barra.Mensaje_Box = "Determinando Registro de Cierre de Caja"
    Progreso_Esperar
    Imagen_Esperar
    Cadena = ""
    EsCxC = False
    sSQL = "SELECT TP,Numero,Fecha,Concepto " _
         & "FROM Comprobantes " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND Concepto LIKE 'Cierre de Caja de%' " _
         & "ORDER BY Fecha "
    Select_AdoDB AdoDBComp, sSQL
    If AdoDBComp.RecordCount > 0 Then
       Progreso_Barra.Valor_Maximo = Progreso_Barra.Valor_Maximo + AdoDBComp.RecordCount
       Do While Not AdoDBComp.EOF
          Imagen_Esperar
          Co.TP = AdoDBComp.fields("TP")
          Co.Numero = AdoDBComp.fields("Numero")
          Co.Concepto = AdoDBComp.fields("Concepto")
          If MidStrg(Co.Concepto, 1, 36) = "Cierre de Caja de Cuentas por Cobrar" Then EsCxC = True Else EsCxC = False
          UnaFecha = True
          For IdF = 1 To Len(Co.Concepto)
             'MsgBox "'" & MidStrg(Co.Concepto, IdF, 10) & "'"
              If IsDate(MidStrg(Co.Concepto, IdF, 10)) Then
                 If Year(MidStrg(Co.Concepto, IdF, 10)) > 2000 Then
                    If UnaFecha Then
                       FechaInicial = MidStrg(Co.Concepto, IdF, 10)
                       FechaFinal = MidStrg(Co.Concepto, IdF, 10)
                       UnaFecha = False
                       'MsgBox "I: " & Co.Concepto & vbCrLf & "'" & MidStrg(Co.Concepto, IdF, 10) & "'"
                       IdF = IdF + 10
                    Else
                       FechaFinal = MidStrg(Co.Concepto, IdF, 10)
                       'MsgBox "F: " & Co.Concepto & vbCrLf & "'" & MidStrg(Co.Concepto, IdF, 10) & "'"
                       IdF = IdF + 10
                    End If
                 End If
              End If
          Next IdF
          
          FechaIni = BuscarFecha(FechaInicial)
          FechaFin = BuscarFecha(FechaFinal)
          'MsgBox AdoDBComp.Fields("Fecha") & " - " & FechaInicial & " - " & FechaFinal & vbCrLf
          Imagen_Esperar
          sSQL = "DELETE * " _
               & "FROM Trans_Kardex " _
               & "WHERE Item = '" & NumEmpresa & "' " _
               & "AND Periodo = '" & Periodo & "' " _
               & "AND TP = '" & Co.TP & "' " _
               & "AND Numero = " & Co.Numero & " "
          Ejecutar_SQL_SP sSQL
          'MsgBox sSQL
          Imagen_Esperar
          If EsCxC Then
             sSQL = "UPDATE Trans_Kardex " _
                  & "SET TP = '" & Co.TP & "', Numero = " & Co.Numero & " " _
                  & "WHERE Item = '" & NumEmpresa & "' " _
                  & "AND Periodo = '" & Periodo_Contable & "' " _
                  & "AND Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
                  & "AND SUBSTRING(Detalle,1,3) ='FA:' "
             Ejecutar_SQL_SP sSQL
             'MsgBox sSQL
             Cadena = Cadena & "CxC: " & AdoDBComp.fields("Fecha") & " - " & FechaInicial & " - " & FechaFinal & vbCrLf
          Else
             sSQL = "UPDATE Trans_Kardex " _
                  & "SET TP = '" & Co.TP & "', Numero = " & Co.Numero & " " _
                  & "WHERE Item = '" & NumEmpresa & "' " _
                  & "AND Periodo = '" & Periodo_Contable & "' " _
                  & "AND Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
                  & "AND SUBSTRING(Detalle,1,3) ='NC:' "
             Ejecutar_SQL_SP sSQL
             'MsgBox sSQL
             
             Cadena = Cadena & "Abonos: " & AdoDBComp.fields("Fecha") & " - " & FechaInicial & " - " & FechaFinal & vbCrLf
          End If
          AdoDBComp.MoveNext
       Loop
    End If
    AdoDBComp.Close
    
    Progreso_Barra.Mensaje_Box = "Reindexando Inventario"
    Progreso_Esperar
    Imagen_Esperar
    sSQL = "UPDATE Trans_Kardex " _
         & "SET Procesado = 0 " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' "
    Ejecutar_SQL_SP sSQL
    
    Unload FEsperar
    Mayorizar_Inventario_SP
    Progreso_Final
End Sub

Public Sub Actualizar_Facturas_Electronicas()
'Dim ObjAutori As New WS_Autorizacion
Dim URLAutorizacion As String
Dim RutaXMLAutorizado As String
Dim RutaXMLRechazado As String
Dim ArrayAutorizacion() As String
Dim Tiempo_Espera As Integer
Dim Tiempo_SRI As Integer
Dim SRI_Aut As Tipo_Estado_SRI
Dim AdoDBAut As ADODB.Recordset

'Certificado para Firmar el documento
 RatonReloj
 Progreso_Barra.Mensaje_Box = "Consultado Documentos no autorizados"
 Progreso_Iniciar

'Pagina de Conexion con el SRI
 URLAutorizacion = Leer_Campo_Empresa("Web_SRI_Autorizado")
 
'Listar fechas de facturas no autorizadas
 FechaIni = FechaSistema
 FechaFin = FechaSistema
 sSQL = "SELECT Autorizacion, MIN(Fecha) As Fecha_Min, MAX(Fecha) As Fecha_Max " _
      & "FROM Facturas " _
      & "WHERE Item = '" & NumEmpresa & "' " _
      & "AND Periodo = '" & Periodo_Contable & "' " _
      & "AND T <> '" & Anulado & "' " _
      & "AND LEN(Autorizacion) = 13 " _
      & "GROUP BY Autorizacion "
 Select_AdoDB AdoDBAut, sSQL
 If AdoDBAut.RecordCount > 0 Then
    FechaIni = AdoDBAut.fields("Fecha_Min")
    FechaFin = AdoDBAut.fields("Fecha_Max")
 End If
 FechaIni = BuscarFecha(FechaIni)
 FechaFin = BuscarFecha(FechaFin)
 AdoDBAut.Close
 Contador = 0
 sSQL = "SELECT CodigoC,Clave_Acceso,Estado_SRI,TC,Fecha,Serie,Factura,Autorizacion " _
      & "FROM Facturas " _
      & "WHERE Item = '" & NumEmpresa & "' " _
      & "AND Periodo = '" & Periodo_Contable & "' " _
      & "AND Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
      & "AND LEN(Autorizacion) = 13 " _
      & "ORDER BY Fecha,TC,Serie,Factura "
 Select_AdoDB AdoDBAut, sSQL
 If AdoDBAut.RecordCount > 0 Then
    Progreso_Barra.Valor_Maximo = Progreso_Barra.Valor_Maximo + AdoDBAut.RecordCount
    
    RutaXMLRechazado = RutaDocumentos & "\Comprobantes no Autorizados\*.xml"
    If Dir$(RutaXMLRechazado) <> "" Then Kill RutaXMLRechazado
    
   'Empezamos a autorizar
    Do While Not AdoDBAut.EOF
       RatonReloj
       Contador = Contador + 1
       FA.Estado_SRI = "CN"
       FA.CodigoC = AdoDBAut.fields("CodigoC")
       FA.TC = AdoDBAut.fields("TC")
       FA.Serie = AdoDBAut.fields("Serie")
       FA.Fecha = AdoDBAut.fields("Fecha")
       FA.Factura = AdoDBAut.fields("Factura")
       FA.Autorizacion = AdoDBAut.fields("Autorizacion")
       FA.ClaveAcceso = AdoDBAut.fields("Clave_Acceso")
       If Len(FA.ClaveAcceso) < 13 Then
         'MsgBox "Ant: " & FA.ClaveAcceso
          FA.ClaveAcceso = Format$(FA.Fecha, "ddmmyyyy") & "01" & RUC & Ambiente & FA.Serie _
                         & Format$(FA.Factura, String(9, "0")) _
                         & "123456781"
          FA.ClaveAcceso = FA.ClaveAcceso & Digito_Verificador_Modulo11(FA.ClaveAcceso)
       End If
       Progreso_Barra.Mensaje_Box = "[" & Contador & " de " & AdoDBAut.RecordCount & " - " _
                                  & Format(FA.Fecha, "MM/yyyy") & "] " _
                                  & "Consultando " & FA.TC _
                                  & " No. " & FA.Serie & "-" & Format(FA.Factura, "000000000")
       Progreso_Esperar
       SRI_Aut.Clave_De_Acceso = FA.ClaveAcceso
       SRI_Aut.Estado_SRI = "CF"
       SRI_Aut.Error_SRI = ""
       RutaXMLAutorizado = RutaDocumentos & "\Comprobantes Autorizados\" & FA.ClaveAcceso & ".xml"
       RutaXMLRechazado = RutaDocumentos & "\Comprobantes no Autorizados\" & FA.ClaveAcceso & ".xml"
       For Tiempo_Espera = 1 To 3
           Progreso_Esperar True
           For Tiempo_SRI = 0 To 300
               Progreso_Esperar True
           Next Tiempo_SRI
'           ArrayAutorizacion = ObjAutori.FF_ObtieneNumAutorizado(URLAutorizacion, FA.ClaveAcceso, RutaXMLAutorizado, RutaXMLRechazado)
           If ArrayAutorizacion(0) = "AUTORIZADO" Then Exit For
       Next Tiempo_Espera
      'Ok Documento Firmado y Autorizado
       If ArrayAutorizacion(0) = "AUTORIZADO" Then
          Progreso_Barra.Mensaje_Box = "[Ok] " & "Actualizando " & FA.TC _
                                     & " No. " & FA.Serie & "-" & Format(FA.Factura, "000000000")
          Progreso_Esperar True
          SRI_Aut.Estado_SRI = "OK"
          SRI_Aut.Error_SRI = "OK"
          SRI_Aut.Autorizacion = ArrayAutorizacion(1)
          SRI_Aut.Fecha_Autorizacion = Format$(MidStrg(ArrayAutorizacion(2), 1, 10), "dd/MM/yyyy")
          SRI_Aut.Hora_Autorizacion = MidStrg(ArrayAutorizacion(2), 12, 8)
          SRI_Aut.Documento_XML = Leer_Archivo_Texto(RutaXMLAutorizado)
          SRI_Actualizar_Documento_XML SRI_Aut
          SRI_Actualizar_Autorizacion_Comprobante FA.TC, SRI_Aut, FA
       Else
          SRI_Aut.Error_SRI = ArrayAutorizacion(0) & " " & ArrayAutorizacion(1)
          'SRI_Actualizar_XML_Factura FA, SRI_Aut
       End If
       AdoDBAut.MoveNext
    Loop
 End If
 AdoDBAut.Close
 RatonNormal
 Progreso_Final
End Sub

Public Sub Actualizar_Retenciones_Electronicas()
'Dim ObjAutori As New WS_Autorizacion
Dim URLAutorizacion As String
Dim RutaXMLAutorizado As String
Dim RutaXMLRechazado As String
Dim ArrayAutorizacion() As String
Dim Tiempo_Espera As Integer
Dim Tiempo_SRI As Integer
Dim SRI_Aut As Tipo_Estado_SRI
Dim AdoDBAut As ADODB.Recordset

'Certificado para Firmar el documento
 RatonReloj
 Progreso_Barra.Mensaje_Box = "Consultado Documentos no autorizados"
 Progreso_Iniciar

'Pagina de Conexion con el SRI
 URLAutorizacion = Leer_Campo_Empresa("Web_SRI_Autorizado")
 
'Listar fechas de facturas no autorizadas
 FechaIni = FechaSistema
 FechaFin = FechaSistema
 
  sSQL = "SELECT AutRetencion, MIN(Fecha) As Fecha_Min, MAX(Fecha) As Fecha_Max  " _
       & "FROM Trans_Compras " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND LEN(Serie_Retencion) = 6 " _
       & "AND LEN(AutRetencion) BETWEEN 13 and 49 " _
       & "AND LEN(Clave_Acceso) > 1 " _
       & "AND Estado_SRI <> 'OK' " _
       & "GROUP BY AutRetencion " _
       & "ORDER BY AutRetencion "
  Select_AdoDB AdoDBAut, sSQL
  If AdoDBAut.RecordCount > 0 Then
     FechaIni = AdoDBAut.fields("Fecha_Min")
     FechaFin = AdoDBAut.fields("Fecha_Max")
  End If
  FechaIni = BuscarFecha(FechaIni)
  FechaFin = BuscarFecha(FechaFin)
  AdoDBAut.Close
  Contador = 0
 
   sSQL = "SELECT C.Cliente,C.CI_RUC,C.TD,C.Direccion,C.Email,C.Email2,C.Ciudad,C.DirNumero,C.Telefono,TC.* " _
        & "FROM Trans_Compras As TC,Clientes As C " _
        & "WHERE TC.Item = '" & NumEmpresa & "' " _
        & "AND TC.Periodo = '" & Periodo_Contable & "' " _
        & "AND TC.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
        & "AND LEN(TC.Serie_Retencion) = 6 " _
        & "AND LEN(TC.AutRetencion) = 13 " _
        & "AND TC.Estado_SRI <> 'OK' " _
        & "AND TC.IdProv = C.Codigo " _
        & "ORDER BY TC.Fecha,Cta_Servicio,Cta_Bienes "
 Select_AdoDB AdoDBAut, sSQL
 If AdoDBAut.RecordCount > 0 Then
    Progreso_Barra.Valor_Maximo = Progreso_Barra.Valor_Maximo + AdoDBAut.RecordCount
    
    RutaXMLRechazado = RutaDocumentos & "\Comprobantes no Autorizados\*.xml"
    If Dir$(RutaXMLRechazado) <> "" Then Kill RutaXMLRechazado
    
   'Empezamos a autorizar
    Do While Not AdoDBAut.EOF
       RatonReloj
       Contador = Contador + 1
       FechaTexto = AdoDBAut.fields("Fecha")
       Co.Fecha = AdoDBAut.fields("Fecha")
       Co.Beneficiario = AdoDBAut.fields("Cliente")
       Co.RUC_CI = AdoDBAut.fields("CI_RUC")
       Co.Direccion = AdoDBAut.fields("Direccion")
       Co.TD = AdoDBAut.fields("TD")
       Co.Email = AdoDBAut.fields("Email")
       Co.TP = AdoDBAut.fields("TP")
       Co.Numero = AdoDBAut.fields("Numero")
             
       FA.EmailC = AdoDBAut.fields("Email")
       FA.EmailR = AdoDBAut.fields("Email2")
       FA.TP = Co.TP
       FA.Numero = Co.Numero
       FA.Fecha = AdoDBAut.fields("FechaEmision")
       FA.Serie_R = AdoDBAut.fields("Serie_Retencion")
       FA.Retencion = AdoDBAut.fields("SecRetencion")
       FA.ClaveAcceso = AdoDBAut.fields("Clave_Acceso")
       FA.Estado_SRI = "CN"
       FA.Autorizacion_R = AdoDBAut.fields("AutRetencion")
       FA.Serie = AdoDBAut.fields("Establecimiento") & AdoDBAut.fields("PuntoEmision")
       FA.Factura = AdoDBAut.fields("Secuencial")
       If Len(FA.ClaveAcceso) = 13 Then
        '  MsgBox "Ant RE: '" & FA.ClaveAcceso & "'"
           FA.ClaveAcceso = Format$(FA.Fecha, "ddmmyyyy") & "07" & RUC & Ambiente & FA.Serie_R & Format$(FA.Retencion, String(9, "0")) _
                          & "123456781"
           FA.ClaveAcceso = Replace(FA.ClaveAcceso, ".", "1")
           FA.ClaveAcceso = FA.ClaveAcceso & Digito_Verificador_Modulo11(FA.ClaveAcceso)
       End If
       Progreso_Barra.Mensaje_Box = "[" & Contador & " de " & AdoDBAut.RecordCount & " - " _
                                  & Format(FA.Fecha, "MM/yyyy") & "] " _
                                  & "Consultando RE " _
                                  & " No. " & FA.Serie_R & "-" & Format(FA.Retencion, "000000000")
       Progreso_Esperar
       SRI_Aut.Clave_De_Acceso = FA.ClaveAcceso
       SRI_Aut.Estado_SRI = "CF"
       SRI_Aut.Error_SRI = ""
       RutaXMLAutorizado = RutaDocumentos & "\Comprobantes Autorizados\" & FA.ClaveAcceso & ".xml"
       RutaXMLRechazado = RutaDocumentos & "\Comprobantes no Autorizados\" & FA.ClaveAcceso & ".xml"
       
       For Tiempo_Espera = 1 To 3
           Progreso_Esperar True
'           ArrayAutorizacion = ObjAutori.FF_ObtieneNumAutorizado(URLAutorizacion, FA.ClaveAcceso, RutaXMLAutorizado, RutaXMLRechazado)
           If ArrayAutorizacion(0) = "AUTORIZADO" Then Exit For
           For Tiempo_SRI = 0 To 50
               Progreso_Esperar True
           Next Tiempo_SRI
       Next Tiempo_Espera
      'Ok Documento Firmado y Autorizado
       If ArrayAutorizacion(0) = "AUTORIZADO" Then
          Progreso_Barra.Mensaje_Box = "[Ok] " & "Actualizando RE " _
                                     & " No. " & FA.Serie_R & "-" & Format(FA.Retencion, "000000000")
          Progreso_Esperar True
          SRI_Aut.Estado_SRI = "OK"
          SRI_Aut.Error_SRI = "OK"
          SRI_Aut.Autorizacion = ArrayAutorizacion(1)
          SRI_Aut.Fecha_Autorizacion = Format$(MidStrg(ArrayAutorizacion(2), 1, 10), "dd/MM/yyyy")
          SRI_Aut.Hora_Autorizacion = MidStrg(ArrayAutorizacion(2), 12, 8)
          SRI_Aut.Documento_XML = Leer_Archivo_Texto(RutaXMLAutorizado)
          SRI_Actualizar_Documento_XML SRI_Aut
          SRI_Actualizar_Autorizacion_Comprobante "RE", SRI_Aut, FA
       Else
          SRI_Aut.Error_SRI = ArrayAutorizacion(0) & " " & ArrayAutorizacion(1)
          'SRI_Actualizar_XML_Factura FA, SRI_Aut
       End If
       AdoDBAut.MoveNext
    Loop
 End If
 AdoDBAut.Close
 RatonNormal
 Progreso_Final
End Sub

Public Sub Actualizar_Codigo_Superiores()
    Contador = 0
    sSQL = "SELECT * " _
         & "FROM Catalogo_Productos " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "ORDER BY Periodo,Codigo_Inv "
    Select_Adodc AdoAux, sSQL
    With AdoAux.Recordset
     If .RecordCount > 0 Then
         Do While Not .EOF
            FSeteos.Caption = "Procesando el " & Format(Contador / .RecordCount, "00.00%") & " - Periodo: " & .fields("Periodo") & ", Codigo: " & .fields("Codigo_Inv")
           .fields("Codigo_Sup") = CambioCodigoCtaSup(.fields("Codigo_Inv"))
           .Update
           .MoveNext
            Contador = Contador + 1
         Loop
     End If
    End With
    FSeteos.Caption = "SETEOS PRINCIPALES Y MANTENIMIENTO"
End Sub

Public Sub Actualizar_Codigo_Clientes(Nombre_Tabla As String, Codigo_Cliente As String)
  RatonReloj
  sSQL = "UPDATE " & Nombre_Tabla & " " _
       & "SET X = '.' " _
       & "WHERE X <> '.' "
  Ejecutar_SQL_SP sSQL

  If SQL_Server Then
     sSQL = "UPDATE " & Nombre_Tabla & " " _
          & "SET X = 'X' " _
          & "FROM " & Nombre_Tabla & " As X,Clientes As C "
  Else
     sSQL = "UPDATE " & Nombre_Tabla & " As X,Clientes As C " _
          & "SET X.X = 'X' "
  End If
  sSQL = sSQL & "WHERE X." & Codigo_Cliente & " = C.Codigo "
  Ejecutar_SQL_SP sSQL
  
  sSQL = "UPDATE " & Nombre_Tabla & " " _
       & "SET " & Codigo_Cliente & " = '.' " _
       & "WHERE X = '.' "
 'MsgBox sSQL
  Ejecutar_SQL_SP sSQL
  RatonNormal
End Sub

Public Sub Actualizar_Codigo_Clientes_Bases()
  RatonReloj
  sSQL = "UPDATE Clientes " _
       & "SET X = 'X' " _
       & "WHERE X <> 'X' "
  Ejecutar_SQL_SP sSQL

  If SQL_Server Then
     sSQL = "UPDATE Clientes " _
          & "SET X = '.' " _
          & "FROM Clientes As C,Catalogo_CxCxP As XB "
  Else
     sSQL = "UPDATE Clientes As C,Catalogo_CxCxP As XB " _
          & "SET C.X = '.' "
  End If
  sSQL = sSQL & "WHERE C.Codigo = XB.Codigo "
  Ejecutar_SQL_SP sSQL
  
  If SQL_Server Then
     sSQL = "UPDATE Clientes " _
          & "SET X = '.' " _
          & "FROM Clientes As C,Catalogo_SubCtas As XB "
  Else
     sSQL = "UPDATE Clientes As C,Catalogo_SubCtas As XB " _
          & "SET C.X = '.' "
  End If
  sSQL = sSQL & "WHERE C.Codigo = XB.Codigo "
  Ejecutar_SQL_SP sSQL
  
  If SQL_Server Then
     sSQL = "UPDATE Clientes " _
          & "SET X = '.' " _
          & "FROM Clientes As C,Clientes_Facturacion As XB "
  Else
     sSQL = "UPDATE Clientes As C,Clientes_Facturacion As XB " _
          & "SET C.X = '.' "
  End If
  sSQL = sSQL & "WHERE C.Codigo = XB.Codigo "
  Ejecutar_SQL_SP sSQL
  
  If SQL_Server Then
     sSQL = "UPDATE Clientes " _
          & "SET X = '.' " _
          & "FROM Clientes As C,Clientes_Matriculas As XB "
  Else
     sSQL = "UPDATE Clientes As C,Clientes_Matriculas As XB " _
          & "SET C.X = '.' "
  End If
  sSQL = sSQL & "WHERE C.Codigo = XB.Codigo "
  Ejecutar_SQL_SP sSQL
  
  If SQL_Server Then
     sSQL = "UPDATE Clientes " _
          & "SET X = '.' " _
          & "FROM Clientes As C,Comprobantes As XB "
  Else
     sSQL = "UPDATE Clientes As C,Comprobantes As XB " _
          & "SET C.X = '.' "
  End If
  sSQL = sSQL & "WHERE C.Codigo = XB.Codigo_B "
  Ejecutar_SQL_SP sSQL
  
  If SQL_Server Then
     sSQL = "UPDATE Clientes " _
          & "SET X = '.' " _
          & "FROM Clientes As C,Comprobantes As XB "
  Else
     sSQL = "UPDATE Clientes As C,Comprobantes As XB " _
          & "SET C.X = '.' "
  End If
  sSQL = sSQL & "WHERE C.Codigo = XB.CodigoU "
  Ejecutar_SQL_SP sSQL
  
  If SQL_Server Then
     sSQL = "UPDATE Clientes " _
          & "SET X = '.' " _
          & "FROM Clientes As C,Detalle_Factura As XB "
  Else
     sSQL = "UPDATE Clientes As C,Detalle_Factura As XB " _
          & "SET C.X = '.' "
  End If
  sSQL = sSQL & "WHERE C.Codigo = XB.CodigoC "
  Ejecutar_SQL_SP sSQL
  
  If SQL_Server Then
     sSQL = "UPDATE Clientes " _
          & "SET X = '.' " _
          & "FROM Clientes As C,Facturas As XB "
  Else
     sSQL = "UPDATE Clientes As C,Facturas As XB " _
          & "SET C.X = '.' "
  End If
  sSQL = sSQL & "WHERE C.Codigo = XB.CodigoC "
  Ejecutar_SQL_SP sSQL
  
  If SQL_Server Then
     sSQL = "UPDATE Clientes " _
          & "SET X = '.' " _
          & "FROM Clientes As C,Trans_Abonos As XB "
  Else
     sSQL = "UPDATE Clientes As C,Trans_Abonos As XB " _
          & "SET C.X = '.' "
  End If
  sSQL = sSQL & "WHERE C.Codigo = XB.CodigoC "
  Ejecutar_SQL_SP sSQL
  
  If SQL_Server Then
     sSQL = "UPDATE Clientes " _
          & "SET X = '.' " _
          & "FROM Clientes As C,Trans_Abonos As XB "
  Else
     sSQL = "UPDATE Clientes As C,Trans_Abonos As XB " _
          & "SET C.X = '.' "
  End If
  sSQL = sSQL & "WHERE C.Codigo = XB.CodigoU "
  Ejecutar_SQL_SP sSQL
  
  If SQL_Server Then
     sSQL = "UPDATE Clientes " _
          & "SET X = '.' " _
          & "FROM Clientes As C,Trans_Air As XB "
  Else
     sSQL = "UPDATE Clientes As C,Trans_Air As XB " _
          & "SET C.X = '.' "
  End If
  sSQL = sSQL & "WHERE C.Codigo = XB.IdProv "
  Ejecutar_SQL_SP sSQL
  
  If SQL_Server Then
     sSQL = "UPDATE Clientes " _
          & "SET X = '.' " _
          & "FROM Clientes As C,Trans_Compras As XB "
  Else
     sSQL = "UPDATE Clientes As C,Trans_Compras As XB " _
          & "SET C.X = '.' "
  End If
  sSQL = sSQL & "WHERE C.Codigo = XB.IdProv "
  Ejecutar_SQL_SP sSQL
  
  If SQL_Server Then
     sSQL = "UPDATE Clientes " _
          & "SET X = '.' " _
          & "FROM Clientes As C,Trans_Gastos_Caja As XB "
  Else
     sSQL = "UPDATE Clientes As C,Trans_Gastos_Caja As XB " _
          & "SET C.X = '.' "
  End If
  sSQL = sSQL & "WHERE C.Codigo = XB.Codigo "
  Ejecutar_SQL_SP sSQL
  
  If SQL_Server Then
     sSQL = "UPDATE Clientes " _
          & "SET X = '.' " _
          & "FROM Clientes As C,Trans_Gastos_Caja As XB "
  Else
     sSQL = "UPDATE Clientes As C,Trans_Gastos_Caja As XB " _
          & "SET C.X = '.' "
  End If
  sSQL = sSQL & "WHERE C.Codigo = XB.CodigoC "
  Ejecutar_SQL_SP sSQL
  
  If SQL_Server Then
     sSQL = "UPDATE Clientes " _
          & "SET X = '.' " _
          & "FROM Clientes As C,Trans_Ventas As XB "
  Else
     sSQL = "UPDATE Clientes As C,Trans_Ventas As XB " _
          & "SET C.X = '.' "
  End If
  sSQL = sSQL & "WHERE C.Codigo = XB.IdProv "
  Ejecutar_SQL_SP sSQL
  
  If SQL_Server Then
     sSQL = "UPDATE Clientes " _
          & "SET X = '.' " _
          & "FROM Clientes As C,Transacciones As XB "
  Else
     sSQL = "UPDATE Clientes As C,Transacciones As XB " _
          & "SET C.X = '.' "
  End If
  sSQL = sSQL & "WHERE C.Codigo = XB.Codigo_C "
  Ejecutar_SQL_SP sSQL
  
  sSQL = "DELETE * " _
       & "FROM Clientes " _
       & "WHERE X = 'X' "
  Ejecutar_SQL_SP sSQL
  RatonNormal
End Sub

Public Sub Actualizar_Cursos_Seteos(Periodo As String, Curso As String, NombreCampo As String, ValorCampo As String)
    If ValorCampo <> Ninguno Then
       sSQL = "UPDATE Catalogo_Cursos " _
            & "SET " & NombreCampo & " = '" & ValorCampo & "' " _
            & "WHERE Item = '" & NumEmpresa & "' " _
            & "AND Periodo = '" & Periodo & "' " _
            & "AND Curso = '" & Curso & "' "
      'MsgBox sSQL
       Ejecutar_SQL_SP sSQL
    End If
End Sub

Public Sub Listar_Tablas()
Dim AdoCon1 As ADODB.Connection
Dim RstSchema As ADODB.Recordset
Dim IJ As Long
Dim IdTime As Long
Dim strCnn As String
Dim NúmeroError
On Error GoTo Errorhandler
' Consultamos las cuentas de la tabla
  RatonReloj
  LstTablas.Clear
  Set AdoCon1 = New ADODB.Connection
  AdoCon1.open AdoStrCnn
  Set RstSchema = AdoCon1.OpenSchema(adSchemaTables)
  Do Until RstSchema.EOF
     If RstSchema!TABLE_TYPE = "TABLE" And MidStrg(RstSchema!TABLE_NAME, 1, 1) <> "~" Then
        LstTablas.AddItem RstSchema!TABLE_NAME
     End If
     RstSchema.MoveNext
  Loop
  AdoCon1.Close
  RatonNormal
  Exit Sub
Errorhandler:
    RatonNormal
    MsgBox "Error:(" & Err & ")" & vbCrLf _
             & "Error de Conexion: " & Printer.DeviceName & vbCrLf _
             & "No pudo Abrir correctamente la base"
    Exit Sub
End Sub

Public Sub Abonar_Facturas_Pendientes()
  FechaFin = FechaSistema
  If ClaveContador Then
     Titulo = "ABONO FACTURAS PENDIENTES"
     Mensajes = "Seguro de realizar los abonos de Facturacion del: " & FechaFin
     If BoxMensaje = vbYes Then
        RatonReloj
        Progreso_Barra.Mensaje_Box = "ABONANDO FACTURAS PENDIENTES AL " & FechaFin
        Progreso_Iniciar
        FA.Factura = 0
        FA.Fecha_Corte = FechaSistema
        Actualizar_Abonos_Facturas_SP FA
        Progreso_Esperar
        
        Total = 0
        sSQL = "SELECT TC, Serie, Autorizacion, CodigoC, Factura, Fecha, Razon_Social, Saldo_MN, Cta_CxP " _
             & "FROM Facturas " _
             & "WHERE Item = '" & NumEmpresa & "' " _
             & "AND Periodo = '" & Periodo_Contable & "' " _
             & "AND Saldo_MN > 0 " _
             & "AND T <> 'A' " _
             & "ORDER BY TC,Serie,Factura "
        Select_Adodc AdoPrueba, sSQL
        With AdoPrueba.Recordset
         If .RecordCount > 0 Then
             MDIFormulario.ProgressBarEstado.value = 0
             MDIFormulario.ProgressBarEstado.Max = .RecordCount
             Progreso_Barra.Incremento = 1
             Progreso_Barra.Valor_Maximo = Progreso_Barra.Valor_Maximo + .RecordCount
            'MsgBox "Empezar: " & .RecordCount
             TA.Cta = Cta_General
             TA.Banco = "ABONO AUTOMATIZADO"
             TA.Cheque = "POR CAJA"
             TA.T = Cancelado
             Do While Not .EOF
                RatonReloj
               'Seteos de Abonos Generales
                TA.Cta_CxP = .fields("Cta_CxP")
                TA.TP = .fields("TC")
                TA.Serie = .fields("Serie")
                TA.Autorizacion = .fields("Autorizacion")
                TA.CodigoC = .fields("CodigoC")
                TA.Factura = .fields("Factura")
                TA.Fecha = .fields("Fecha")
                TA.Recibi_de = .fields("Razon_Social")
                TA.Abono = .fields("Saldo_MN")
               'MsgBox TA.Recibi_de
                Grabar_Abonos TA
                Total = Total + TA.Abono
                Progreso_Barra.Mensaje_Box = "ABONANDO DOCUMENTO " & .fields("TC") & " No. " & .fields("Serie") & "-" & .fields("Factura")
                Progreso_Esperar
                MDIFormulario.ProgressBarEstado.value = MDIFormulario.ProgressBarEstado.value + 1
               'MsgBox MDIFormulario.ProgressBarEstado.value & vbCrLf & MDIFormulario.ProgressBarEstado.Max
               .MoveNext
             Loop
         End If
        End With
     End If
  End If
  RatonNormal
  Progreso_Final
End Sub

Public Sub Procesar_Cambio_PV_FA()
Dim ITab As Long
Dim JCamp As Long
Dim KCamp As Long
Dim ListaItem As String


Progreso_Barra.Incremento = 0
Progreso_Barra.Valor_Maximo = LstTablas.ListCount - 1
Progreso_Barra.Mensaje_Box = "SETEOS"
Progreso_Esperar
Progreso_Barra.Incremento = 0
Progreso_Barra.Valor_Maximo = LstTablas.ListCount - 1
Progreso_Barra.Mensaje_Box = "SETEOS"
Progreso_Esperar
Ln_No = 0
sSQL = "SELECT * " _
     & "FROM Detalle_Factura " _
     & "WHERE Item = '" & NumEmpresa & "' " _
     & "AND Periodo = '" & Periodo_Contable & "' " _
     & "AND TC = 'PV' " _
     & "AND T <> 'A' " _
     & "ORDER BY Fecha,Factura "
Select_Adodc AdoCuentas, sSQL
With AdoCuentas.Recordset
 If .RecordCount > 0 Then
     Progreso_Barra.Valor_Maximo = .RecordCount
     Factura_No = .fields("Factura")
     Comp_No = 1
     Do While Not .EOF
        If Factura_No <> .fields("Factura") Then
           Factura_No = .fields("Factura")
           Comp_No = Comp_No + 1
        End If
        SetAdoAddNew "Trans_Ticket"
        SetAdoFields "TC", .fields("TC")
        SetAdoFields "Fecha", .fields("Fecha")
        SetAdoFields "CodigoC", .fields("CodigoC")
        SetAdoFields "Factura", .fields("Factura")
        SetAdoFields "Ticket", Comp_No
        SetAdoFields "Codigo_Inv", .fields("Codigo")
        SetAdoFields "Producto", TrimStrg(MidStrg(.fields("Producto"), 1, 50))
        SetAdoFields "Cantidad", .fields("Cantidad")
        SetAdoFields "Precio", .fields("Precio")
        SetAdoFields "Total", .fields("Total")
        SetAdoFields "Descuento", .fields("Total_Desc")
        SetAdoFields "Item", NumEmpresa
        SetAdoFields "CodigoU", CodigoUsuario
        SetAdoFields "Periodo", Periodo_Contable
        SetAdoFields "D_No", Ln_No
        SetAdoUpdate
        Progreso_Barra.Mensaje_Box = "Conversion a Punto de Venta"
        Progreso_Esperar
        Ln_No = Ln_No + 1
       .MoveNext
     Loop
 End If
End With
sSQL = "DELETE * " _
     & "FROM Facturas " _
     & "WHERE Item = '" & NumEmpresa & "' " _
     & "AND Periodo = '" & Periodo_Contable & "' " _
     & "AND TC = 'PV' "
Ejecutar_SQL_SP sSQL
sSQL = "DELETE * " _
     & "FROM Detalle_Factura " _
     & "WHERE Item = '" & NumEmpresa & "' " _
     & "AND Periodo = '" & Periodo_Contable & "' " _
     & "AND TC = 'PV' "
Ejecutar_SQL_SP sSQL
sSQL = "DELETE * " _
     & "FROM Trans_Abonos " _
     & "WHERE Item = '" & NumEmpresa & "' " _
     & "AND Periodo = '" & Periodo_Contable & "' " _
     & "AND TP = 'PV' "
Ejecutar_SQL_SP sSQL

MsgBox "Proceso terminado con exito"
Unload FSeteos
End Sub

Public Sub Procesar_Lista_Ctas_Grupo()
  sSQL = "SELECT T.Item,T.Fecha,T.TP,T.Numero,C.Codigo,C.Cuenta,T.Cheq_Dep,T.Debe,T.Haber,T.Parcial_ME " _
       & "FROM Transacciones As T,Catalogo_Cuentas As C " _
       & "WHERE C.Item = T.Item " _
       & "AND C.Codigo = T.Cta " _
       & "AND C.DG = 'G' " _
       & "ORDER BY T.Item,C.Codigo "
  Select_Adodc AdoAux, sSQL
  'Select_Adodc_Grid DGAux, AdoAux, sSQL
End Sub

Public Sub Procesar_Renumerar_Claves()
 '
  sSQL = "UPDATE Catalogo_Cuentas " _
       & "SET Clave = 0 " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' "
  Ejecutar_SQL_SP sSQL
  
  sSQL = "UPDATE Catalogo_Cuentas " _
       & "SET TC = 'N' " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND TC = '.' "
  Ejecutar_SQL_SP sSQL
  
  sSQL = "SELECT * " _
       & "FROM Catalogo_Cuentas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND DG = 'D' " _
       & "ORDER BY Codigo "
  Select_Adodc AdoAux, sSQL
  RatonReloj
  With AdoAux.Recordset
   If .RecordCount > 0 Then
       Contador = 0
       Do While Not .EOF
          Contador = Contador + 1
         .fields("Clave") = Contador
         .Update
         .MoveNext
       Loop
   End If
  End With
  RatonNormal
 '
End Sub

Public Sub Procesar_Renumerar_CIRUC()
    sSQL = "SELECT Codigo,CI_RUC,TD,Cliente,Direccion,Grupo " _
         & "FROM Clientes " _
         & "WHERE Codigo <> '.' " _
         & "ORDER BY TD,Cliente,Grupo,Codigo "
    Select_Adodc AdoComp, sSQL
    RatonReloj
    With AdoComp.Recordset
     If .RecordCount > 0 Then
         Progreso_Iniciar
         Progreso_Barra.Valor_Maximo = .RecordCount * 2
         Progreso_Barra.Mensaje_Box = "SETEOS"
         Progreso_Esperar
         Do While Not .EOF
            Progreso_Barra.Mensaje_Box = "Actualizando: " & .fields("Cliente")
            Progreso_Esperar
            Codigo1 = .fields("Codigo")
            Cadena = Digito_Verificador(.fields("CI_RUC"))
            Codigo2 = Tipo_RUC_CI.Codigo_RUC_CI
          
            sSQL = "UPDATE Clientes " _
                 & "SET TD = '" & Tipo_RUC_CI.Tipo_Beneficiario & "' " _
                 & "WHERE Codigo = '" & Codigo1 & "' "
            Ejecutar_SQL_SP sSQL
            
            Progreso_Barra.Mensaje_Box = "Cambio de Codigo de " & .fields("Cliente")
            Progreso_Esperar
            
            Cambio_De_Codigo "Acceso_Empresa", "Codigo", Codigo1, Codigo2
            Cambio_De_Codigo "Accesos", "Codigo", Codigo1, Codigo2
            Cambio_De_Codigo "Catalogo_CxCxP", "Codigo", Codigo1, Codigo2
            Cambio_De_Codigo "Catalogo_Rol_Pagos", "Codigo", Codigo1, Codigo2
            Cambio_De_Codigo "Catalogo_Rol_Rubros", "Codigo", Codigo1, Codigo2
            Cambio_De_Codigo "Catalogo_SubCtas", "Codigo", Codigo1, Codigo2
            Cambio_De_Codigo "Clientes", "Codigo", Codigo1, Codigo2
            Cambio_De_Codigo "Clientes_Datos_Extras", "Codigo", Codigo1, Codigo2
            Cambio_De_Codigo "Clientes_Facturacion", "Codigo", Codigo1, Codigo2
            Cambio_De_Codigo "Clientes_Matriculas", "Codigo", Codigo1, Codigo2
            Cambio_De_Codigo "Comprobantes", "Codigo_B", Codigo1, Codigo2
            Cambio_De_Codigo "Detalle_Factura", "CodigoC", Codigo1, Codigo2
            Cambio_De_Codigo "Facturas", "CodigoC", Codigo1, Codigo2
            Cambio_De_Codigo "Trans_Abonos", "CodigoC", Codigo1, Codigo2
            Cambio_De_Codigo "Trans_Actas", "Codigo", Codigo1, Codigo2
            Cambio_De_Codigo "Trans_Activos", "Codigo_R", Codigo1, Codigo2
            Cambio_De_Codigo "Trans_Aduanas", "CodigoC", Codigo1, Codigo2
            Cambio_De_Codigo "Trans_Air", "IdProv", Codigo1, Codigo2
            Cambio_De_Codigo "Trans_Asistencia", "Codigo", Codigo1, Codigo2
            Cambio_De_Codigo "Trans_Comision", "CodigoC", Codigo1, Codigo2
            Cambio_De_Codigo "Trans_Compras", "IdProv", Codigo1, Codigo2
            Cambio_De_Codigo "Trans_Cuotas", "CodigoC", Codigo1, Codigo2
            Cambio_De_Codigo "Trans_Exportaciones", "IdFiscalProv", Codigo1, Codigo2
            Cambio_De_Codigo "Trans_Fideicomiso", "Codigo", Codigo1, Codigo2
            Cambio_De_Codigo "Trans_Fletes", "CodigoC", Codigo1, Codigo2
            Cambio_De_Codigo "Trans_Gastos_Caja", "Codigo", Codigo1, Codigo2
            Cambio_De_Codigo "Trans_Importaciones", "IdFiscalProv", Codigo1, Codigo2
            Cambio_De_Codigo "Trans_Kardex", "Codigo_P", Codigo1, Codigo2
            Cambio_De_Codigo "Trans_Memos", "Codigo", Codigo1, Codigo2
            Cambio_De_Codigo "Trans_Memos", "CC1", Codigo1, Codigo2
            Cambio_De_Codigo "Trans_Memos", "CC2", Codigo1, Codigo2
            Cambio_De_Codigo "Trans_Notas", "Codigo", Codigo1, Codigo2
            Cambio_De_Codigo "Trans_Notas_Grado", "Codigo", Codigo1, Codigo2
            Cambio_De_Codigo "Trans_Pedidos", "Codigo", Codigo1, Codigo2
            Cambio_De_Codigo "Trans_Promedios", "Codigo", Codigo1, Codigo2
            Cambio_De_Codigo "Trans_Rol_de_Pagos", "Codigo", Codigo1, Codigo2
            Cambio_De_Codigo "Trans_Rol_Horas", "Codigo", Codigo1, Codigo2
            Cambio_De_Codigo "Trans_Rol_Pagos", "Codigo", Codigo1, Codigo2
            Cambio_De_Codigo "Trans_SubCtas", "Codigo", Codigo1, Codigo2
            Cambio_De_Codigo "Trans_Ticket", "CodigoC", Codigo1, Codigo2
            Cambio_De_Codigo "Trans_Ventas", "IdProv", Codigo1, Codigo2
            Cambio_De_Codigo "Transacciones", "Codigo_C", Codigo1, Codigo2
            Cambio_De_Codigo "Prestamos", "Cuenta_No", Codigo1, Codigo2
           .MoveNext
         Loop
     End If
    End With
    Progreso_Final
    RatonNormal
    MsgBox "Proceso Terminado"
    Unload FSeteos
End Sub

Public Sub Procesar_Limpiar_Basura()
Dim AdoCon1 As ADODB.Connection
Dim RstSchema As ADODB.Recordset
Dim SiFecha As Boolean
Dim IdTime As Long
Dim MesNo As Integer
Dim ItemMax As Integer
Dim Items As String
Dim strCnn As String
Dim itmX As ListItem

  RatonNormal
  Mensajes = "Este proceso eliminara basura procesada " _
           & "en la base de Datos, y solo se debera ejecutar " _
           & "por un administrador de la base de datos, o dirijida " _
           & "por el mismo." & vbCrLf & vbCrLf _
           & "ESTA SEGURO DE PROSEGUIR?"
  Titulo = "ELIMINAR BASURA"
  If BoxMensaje = vbYes Then
     RatonReloj
     Control_Procesos Normal, "Ejecuta Eliminar Basura "
     Items = ""
     Progreso_Barra.Incremento = 0
     Progreso_Barra.Mensaje_Box = "Reabriendo Periodo"
     Progreso_Iniciar
     Progreso_Barra.Valor_Maximo = LstTablas.ListCount + 10
     
     sSQL = "SELECT Item " _
          & "FROM Empresas " _
          & "WHERE Item <> '.' " _
          & "ORDER BY Item "
     Select_Adodc AdoEmp, sSQL
     With AdoEmp.Recordset
      If .RecordCount > 0 Then
          Do While Not .EOF
             Items = Items & "'" & .fields("Item") & "', "
            .MoveNext
          Loop
      End If
     End With
     
     Items = Items & "'000' "
     For I = 0 To LstTablas.ListCount - 1
         Si_No = False
         Progreso_Barra.Mensaje_Box = "LBD: " & LstTablas.List(I)
         Progreso_Esperar
         sSQL = "SELECT " & Full_Fields(LstTablas.List(I)) & " " _
              & "FROM " & LstTablas.List(I) & " " _
              & "WHERE 1 = 0 "
         Select_Adodc AdoAux, sSQL
         RatonReloj
         With AdoAux.Recordset
          For J = 0 To .fields.Count - 1
              If .fields(J).Name = "Item" Then Si_No = True
          Next J
         End With
         If LstTablas.List(I) = "Clientes" Then Si_No = False
         If LstTablas.List(I) = "Empresas" Then Si_No = False
         If Si_No Then
            sSQL = "DELETE * " _
                 & "FROM " & LstTablas.List(I) & " " _
                 & "WHERE NOT Item IN (" & Items & ") "
            Ejecutar_SQL_SP sSQL
         End If
     Next I
     RatonNormal
     Progreso_Final
  End If
End Sub

Public Sub Procesar_Limpiar_Basura_000()
Dim AdoCon1 As ADODB.Connection
Dim RstSchema As ADODB.Recordset
Dim SiFecha As Boolean
Dim IdTime As Long
Dim MesNo As Integer
Dim ItemMax As Integer
Dim Items As String
Dim strCnn As String
Dim itmX As ListItem

  RatonNormal
  Mensajes = "Este proceso eliminara basura procesada " _
           & "en la base de Datos, y solo se debera ejecutar " _
           & "por un administrador de la base de datos, o dirijida " _
           & "por el mismo." & vbCrLf & vbCrLf _
           & "ESTA SEGURO DE PROSEGUIR?"
  Titulo = "ELIMINAR BASURA"
  If BoxMensaje = vbYes Then
     RatonReloj
     Control_Procesos Normal, "Ejecuta Eliminar Basura 000"
     Items = ""
     Progreso_Barra.Incremento = 0
     Progreso_Barra.Mensaje_Box = "Reabriendo Periodo"
     Progreso_Iniciar
     Progreso_Barra.Valor_Maximo = LstTablas.ListCount + 10
     Items = Items & "'000' "
     For I = 0 To LstTablas.ListCount - 1
         Si_No = False
         Modificar = False
         Progreso_Barra.Mensaje_Box = "LBD: " & LstTablas.List(I)
         Progreso_Esperar
         sSQL = "SELECT " & Full_Fields(LstTablas.List(I)) & " " _
              & "FROM " & LstTablas.List(I) & " " _
              & "WHERE 1 = 0 "
         Select_Adodc AdoAux, sSQL
         RatonReloj
         With AdoAux.Recordset
          For J = 0 To .fields.Count - 1
              If .fields(J).Name = "Item" Then Si_No = True
          Next J
         End With
         If LstTablas.List(I) = "Clientes" Then Si_No = False
         If LstTablas.List(I) = "Empresas" Then Si_No = False
         If MidStrg(LstTablas.List(I), 1, 6) = "Tabla_" Then Modificar = True
         If MidStrg(LstTablas.List(I), 1, 5) = "Tipo_" Then Modificar = True
         If Si_No Then
            sSQL = "DELETE * " _
                 & "FROM " & LstTablas.List(I) & " " _
                 & "WHERE Item = '000' "
            Ejecutar_SQL_SP sSQL
         End If
         If Modificar Then
            sSQL = "DELETE * " _
                 & "FROM " & LstTablas.List(I) & " " _
                 & "WHERE ID >= 0 "
            Ejecutar_SQL_SP sSQL
         End If
     Next I
     RatonNormal
     Progreso_Final
  End If
End Sub

Public Sub Eliminar_Tabla_Vacias()
  RatonNormal
  Mensajes = "Este proceso eliminara todas las tablas en la base de Datos " _
           & "que se encuentren vacias, y solo se debera ejecutar " _
           & "por un administrador de la base de datos, o dirijida " _
           & "por el mismo." & vbCrLf & vbCrLf _
           & "ESTA SEGURO DE PROSEGUIR?"
  Titulo = "ELIMINAR TABLAS"
  If BoxMensaje = vbYes Then
     RatonReloj
     Control_Procesos Normal, "Ejecuta Eliminar Tablas vacias"
     Progreso_Barra.Incremento = 0
     Progreso_Barra.Mensaje_Box = "Verificando Tablas"
     Progreso_Iniciar
     Progreso_Barra.Valor_Maximo = LstTablas.ListCount + 10
     For I = 0 To LstTablas.ListCount - 1
         Progreso_Barra.Mensaje_Box = "LBD: " & LstTablas.List(I)
         Progreso_Esperar
         sSQL = "SELECT COUNT(*) As Contador " _
              & "FROM " & LstTablas.List(I) & " "
         Select_Adodc AdoAux, sSQL
         RatonReloj
         If AdoAux.Recordset.fields("Contador") = 0 Then Ejecutar_SQL_SP "DROP TABLE IF EXISTS " & LstTablas.List(I) & ";"
     Next I
     RatonNormal
     Progreso_Final
  End If
End Sub

Public Sub Procesar_Indices_Base_Datos()
  RatonNormal
  Titulo = "ELIMINAR INDICES"
  Mensajes = "Este proceso eliminara los Indices " _
           & "en la base de Datos." & vbCrLf & vbCrLf _
           & "ESTA SEGURO DE PROSEGUIR?"
  If BoxMensaje = vbYes Then
     RatonReloj
     Control_Procesos Normal, "Ejecuta Eliminar Indices"
     Eliminar_Indices_SP
     Progreso_Barra.Incremento = 0
     Progreso_Barra.Mensaje_Box = "Reabriendo Periodo"
     Progreso_Iniciar
     Progreso_Barra.Valor_Maximo = LstTablas.ListCount + 10
     For I = 0 To LstTablas.ListCount - 1
         Si_No = False
         Progreso_Barra.Mensaje_Box = "LBD: " & LstTablas.List(I)
         Progreso_Esperar
         sSQL = "SELECT " & Full_Fields(LstTablas.List(I)) & " " _
              & "FROM " & LstTablas.List(I) & " " _
              & "WHERE 1 = 0 "
         Select_Adodc AdoAux, sSQL
         RatonReloj
         With AdoAux.Recordset
          For J = 0 To .fields.Count - 1
              If .fields(J).Name = "ID" Then Si_No = True
          Next J
         End With
         If Si_No Then
            sSQL = "ALTER TABLE " & LstTablas.List(I) & " " _
                 & "DROP COLUMN [ID];"
            Ejecutar_SQL_SP sSQL
         End If
     Next I
     RatonNormal
     Progreso_Final
  End If
End Sub

Public Sub Procesar_Update_DB()
Dim AdoCon1 As ADODB.Connection
Dim RstSchema As ADODB.Recordset
Dim ITab As Long
Dim JCamp As Long
Dim KCamp As Long
Dim NombreTabla As String
Dim NombreTablaE As String
Dim CamposTabla As String
Dim Si_Periodo As Boolean
Dim Si_Fecha As Boolean
Dim Si_Fecha_I As Boolean
Dim Si_TP As Boolean
Dim Si_ID As Boolean
Dim Si_Campo As Boolean
Dim Si_Codigo As Boolean
Dim Si_Update As Boolean
'
FileResp = 0
If Existe_File(RutaSistema & "\BASES\UPDATE_DB\*.csv") Then Kill RutaSistema & "\BASES\UPDATE_DB\*.csv"
If Existe_File(RutaSistema & "\BASES\UPDATE_DB\*.dbs") Then Kill RutaSistema & "\BASES\UPDATE_DB\*.dbs"
If Existe_File(RutaSistema & "\BASES\UPDATE_DB\*.txt") Then Kill RutaSistema & "\BASES\UPDATE_DB\*.txt"
If Existe_File(RutaSistema & "\BASES\UPDATE_DB\*.upd") Then Kill RutaSistema & "\BASES\UPDATE_DB\*.upd"
'MsgBox "..."
 sSQL = "DELETE * " _
      & "FROM Trans_Entrada_Salida " _
      & "WHERE Item IN ('.','000') "
 Ejecutar_SQL_SP sSQL
 
 sSQL = "DELETE * " _
      & "FROM Seteos_Documentos " _
      & "WHERE LEN(TP) > 2 " _
      & "AND Item = '000' "
 Ejecutar_SQL_SP sSQL
 
 sSQL = "DELETE * " _
      & "FROM Trans_Documentos " _
      & "WHERE Item = '000' "
 Ejecutar_SQL_SP sSQL
 
 Ejecutar_SQL_SP "DROP TABLE IF EXISTS Actualizacion"
 
 ' Crea variables de objeto para los objetos de acceso a datos.
'    Dim itmX As ListItem
    LstTablas.Clear
    Contador = 0
    Set AdoCon1 = New ADODB.Connection
    AdoCon1.open AdoStrCnn
    Set RstSchema = AdoCon1.OpenSchema(adSchemaTables)
    Do Until RstSchema.EOF
       If RstSchema!TABLE_TYPE = "TABLE" And MidStrg(RstSchema!TABLE_NAME, 1, 1) <> "~" Then
         'Llenamos la lista de Tablas
          
          If RstSchema!TABLE_NAME <> "Actualizacion" Then
             LstTablas.AddItem RstSchema!TABLE_NAME
             Contador = Contador + 1
          End If
       End If
       RstSchema.MoveNext
    Loop

 'MsgBox "Desktop Test: " & Contador
 
 Progreso_Barra.Incremento = 0
 Progreso_Barra.Valor_Maximo = LstTablas.ListCount - 1
 Progreso_Barra.Mensaje_Box = "SETEOS"
 Progreso_Esperar
 
''' For ITab = 0 To LstTablas.ListCount - 1
'''     sSQL = "SELECT COLUMN_NAME " _
'''          & "FROM Information_Schema.Columns " _
'''          & "WHERE TABLE_NAME = '" & LstTablas.List(ITab) & "' " _
'''          & "AND COLUMN_NAME IN ('Item','Periodo') "
'''     Select_Adodc AdoComp, sSQL
'''     If AdoComp.Recordset.RecordCount > 0 Then
'''        sSQL = "DELETE * " _
'''             & "FROM " & LstTablas.List(ITab) & " " _
'''             & "WHERE Item = '000' " _
'''             & "AND Periodo <> '.' "
'''       'MsgBox sSQL
'''        'Ejecutar_SQL_SP sSQL
'''     End If
''' Next ITab

 For ITab = 0 To LstTablas.ListCount - 1
    LstTablas.Text = LstTablas.List(ITab)
    Si_No = False
    sSQL = "SELECT * " _
         & "FROM " & LstTablas.List(ITab) & " " _
         & "WHERE 1 = 0 "
    Select_Adodc AdoComp, sSQL
    RatonReloj
    RutaGeneraFile = RutaSistema & "\BASES\UPDATE_DB\" & LstTablas.List(ITab) & ".upd"
   'FSeteos.Caption = RutaGeneraFile
    NumFile = FreeFile
    Open RutaGeneraFile For Output As #NumFile
    With AdoComp.Recordset
     Print #NumFile, "INSERT INTO Estructura_Tabla (Campo, Tipo_SQL, X) VALUES "
     Cadena = ""
     For JCamp = 0 To .fields.Count - 1
         Cadena1 = FieldTypeSQL(.fields(JCamp).Type)
         Cadena2 = FieldTypeAccess(.fields(JCamp).Type)
         KCamp = .fields(JCamp).DefinedSize
        'If .Fields(JCamp).Name = "Ruta_Certificado" Then MsgBox KCamp
         If KCamp >= 65535 Then KCamp = 0
         Select Case .fields(JCamp).Type
           Case adLongVarWChar
                Cadena1 = Cadena1 & "(MAX)"
           Case adVarWChar
                Cadena1 = Cadena1 & "(" & CStr(KCamp) & ")"
                Cadena2 = Cadena2 & "(" & CStr(KCamp) & ")"
           Case adNumeric: Cadena1 = Cadena1 & "(18,0)"
         End Select
         Cadena = Cadena & "('" & .fields(JCamp).Name & "','" & Cadena1 & "','.')," & vbCrLf
         
''         Cadena1 = Cadena1 & " NULL"
''         Cadena2 = Cadena2 & " NULL"
''
''         Print #NumFile, .fields(JCamp).Name;
''         Print #NumFile, "|";
''         Print #NumFile, Cadena1;
''         Print #NumFile, "|";
''         Print #NumFile, Cadena2;
''         Print #NumFile, "|"
        'MsgBox "Desktop Test: " & Cadena1 & " - " & Cadena2
     Next JCamp
     Cadena = MidStrg(Cadena, 1, Len(Cadena) - 3) & ";"
     Print #NumFile, Cadena
    End With
    Close #NumFile
    RatonNormal
    Progreso_Barra.Mensaje_Box = LstTablas.List(ITab)
    Progreso_Esperar
Next ITab

'Generamos Los Datos en Item = 000
 Progreso_Barra.Incremento = 0
 Progreso_Barra.Valor_Maximo = LstTablas.ListCount - 1
 Progreso_Barra.Mensaje_Box = "SETEOS"
 Progreso_Esperar
 For ITab = 0 To LstTablas.ListCount - 1
     RatonReloj
     LstTablas.Text = LstTablas.List(ITab)
     Select Case LstTablas.List(ITab)
       Case "Accesos", "Clientes", "Clientes_Auxiliar", "Empresas_Externas": Si_Update = False
       Case Else: Si_Update = True
     End Select
     If Si_Update Then
        Si_Periodo = False
        Si_No = False
        Si_Fecha = False
        Si_Fecha_I = False
        Si_TP = False
        Si_Campo = False
        Si_Codigo = False
        Si_ID = False
        CamposTabla = ""
        sSQL = "SELECT * " _
             & "FROM " & LstTablas.List(ITab) & " " _
             & "WHERE 1 = 0 "
        Select_Adodc AdoComp, sSQL
        RatonReloj
        With AdoComp.Recordset
         For JCamp = 0 To .fields.Count - 1
             If .fields(JCamp).Name = "Item" Then Si_No = True
             If .fields(JCamp).Name = "Fecha" Then Si_Fecha = True
             If .fields(JCamp).Name = "Fecha_Inicio" Then Si_Fecha_I = True
             If .fields(JCamp).Name = "TP" Then Si_TP = True
             If .fields(JCamp).Name = "Campo" Then Si_Campo = True
             If .fields(JCamp).Name = "Codigo" Then Si_Codigo = True
             If .fields(JCamp).Name = "ID" Then Si_ID = True
             If .fields(JCamp).Name = "Periodo" Then Si_Periodo = True
             If .fields(JCamp).Name <> "ID" Then CamposTabla = CamposTabla & .fields(JCamp).Name & ","
         Next JCamp
        End With
        CamposTabla = MidStrg(CamposTabla, 1, Len(CamposTabla) - 1)
        sSQL = "SELECT " & CamposTabla & " " _
             & "FROM " & LstTablas.List(ITab) & " "
        If Si_No Then
           sSQL = sSQL & "WHERE Item = '000' "
           If Si_Periodo Then sSQL = sSQL & "AND Periodo = '.' "
        Else
           If LstTablas.List(ITab) = "Clientes" Then sSQL = sSQL & "WHERE Grupo = '999999' "
           If LstTablas.List(ITab) = "Accesos" Then sSQL = sSQL & "WHERE Codigo IN ('ACCESO01','ACCESO02','ACCESO03','ACCESO04','ACCESO05','ACCESO06','ACCESO07','ACCESO08','ACCESO09','ACCESO10','0702164179') "
        End If
        Cadena = ""
        If Si_Fecha Then Cadena = Cadena & "Fecha,"
        If Si_Fecha_I Then Cadena = Cadena & "Fecha_Inicio,"
        If Si_TP Then Cadena = Cadena & "TP,"
        If Si_Campo Then Cadena = Cadena & "Campo,"
        If Si_Codigo Then Cadena = Cadena & "Codigo,"
        If Si_ID Then Cadena = Cadena & "ID,"
        If Cadena <> "" Then
           Cadena = MidStrg(Cadena, 1, Len(Cadena) - 1)
           sSQL = sSQL & "ORDER BY " & Cadena & " "
        End If
        Select_Adodc AdoAux, sSQL
        GenerarTablaEnArchivoPlano FechaSistema, LstTablas.List(ITab), AdoAux
     End If
     RatonNormal
     Progreso_Barra.Mensaje_Box = LstTablas.List(ITab)
     Progreso_Esperar
 Next ITab
Progreso_Barra.Incremento = 0
Progreso_Barra.Valor_Maximo = LstTablas.ListCount - 1
Progreso_Barra.Mensaje_Box = "SETEOS"
Progreso_Esperar

Progreso_Barra.Incremento = 0
Progreso_Barra.Valor_Maximo = 0
Progreso_Barra.Mensaje_Box = "ABASES"
Progreso_Esperar

RutaGeneraFile = RutaSistema & "\BASES\UPDATE_DB\ABASES.txt"
NumFile = FreeFile
Open RutaGeneraFile For Output As #NumFile

RutaGeneraFile = RutaSistema & "\BASES\UPDATE_DB\ACAMPOS.txt"
NumFile2 = FreeFile
Open RutaGeneraFile For Output As #NumFile2


Print #NumFile, "+";
Print #NumFile, String$(89, "=");
Print #NumFile, "+"
NombreTabla = "CANTIDAD DE TABLAS PARA ACTUALIZAR DE LA BASE DE DATOS: " & LstTablas.ListCount
Print #NumFile, "|";
Print #NumFile, NombreTabla;
Print #NumFile, String$(89 - Len(NombreTabla), " ");
Print #NumFile, "|"
Print #NumFile, "+";
Print #NumFile, String$(89, "=");
Print #NumFile, "+"

Print #NumFile2, "+";
Print #NumFile2, String$(89, "=");
Print #NumFile2, "+"
NombreTabla = "CANTIDAD DE TABLAS PARA ACTUALIZAR DE LA BASE DE DATOS: " & LstTablas.ListCount
Print #NumFile2, "|";
Print #NumFile2, NombreTabla;
Print #NumFile2, String$(89 - Len(NombreTabla), " ");
Print #NumFile2, "|"
Print #NumFile2, "+";
Print #NumFile2, String$(89, "=");
Print #NumFile2, "+"

For I = 0 To LstTablas.ListCount - 1
    LstTablas.Text = LstTablas.List(I)
    sSQL = "SELECT * " _
         & "FROM " & LstTablas.List(I) & " " _
         & "WHERE 1 = 0 "
    Select_Adodc AdoComp, sSQL
    RatonReloj
    With AdoComp.Recordset
    'FSeteos.Caption = "(" & RutaGeneraFile & ") => TABLA: " & LstTablas.List(I)
     NombreTablaE = LstTablas.List(I)
     NombreTabla = "TABLA: " & Format(.fields.Count, "000") & " - " & LstTablas.List(I)
     Print #NumFile, "|";
     Print #NumFile, NombreTabla;
     Print #NumFile, String$(89 - Len(NombreTabla), " ");
     Print #NumFile, "|"
     Print #NumFile, "|CAMPO";
     Print #NumFile, String$(20, " ");
     Print #NumFile, "|TIPO";
     Print #NumFile, String$(1, " ");
     Print #NumFile, "|ADODC";
     Print #NumFile, String$(18 - 5, " ");
     Print #NumFile, "|SQL";
     Print #NumFile, String$(22 - 3, " ");
     Print #NumFile, "|ACCESS";
     Print #NumFile, String$(15 - 6, " ");
     Print #NumFile, "|"
     For J = 0 To .fields.Count - 1
         Cadena = FieldType("", .fields(J).Type)
         Cadena1 = FieldTypeSQL(.fields(J).Type)
         Cadena2 = FieldTypeAccess(.fields(J).Type)
         K = .fields(J).DefinedSize
         If K > 65535 Then K = 0
         Select Case .fields(J).Type
           Case adLongVarWChar
                Cadena1 = Cadena1 & "(MAX)"
           Case adVarWChar
                Cadena2 = Cadena2 & "(" & CStr(K) & ")"
                Cadena1 = Cadena1 & "(" & CStr(K) & ")"
         End Select
         If .fields(J).Type = adNumeric Then Cadena1 = Cadena1 & "(18,0)"
         Cadena2 = Cadena2 & " NULL"
         Cadena1 = Cadena1 & " NULL"
         Print #NumFile, "|";
         Print #NumFile, .fields(J).Name;
         Print #NumFile, String$(30 - Len(.fields(J).Name), " ");
         Print #NumFile, "|";
         Print #NumFile, CStr(.fields(J).Type);
         Print #NumFile, String$(5 - Len(CStr(.fields(J).Type)), " ");
         Print #NumFile, "|";
         Print #NumFile, Cadena;
         Print #NumFile, String$(18 - Len(Cadena), " ");
         Print #NumFile, "|";
         Print #NumFile, Cadena1;
         Print #NumFile, String$(22 - Len(Cadena1), " ");
         Print #NumFile, "|";
         Print #NumFile, Cadena2;
         Print #NumFile, String$(15 - Len(Cadena2), " ");
         Print #NumFile, "|"
         
         If InStr(UCase(.fields(J).Name), "Ñ") Then Print #NumFile2, NombreTablaE & "." & .fields(J).Name
         If InStr(UCase(.fields(J).Name), "Á") Then Print #NumFile2, NombreTablaE & "." & .fields(J).Name
         If InStr(UCase(.fields(J).Name), "É") Then Print #NumFile2, NombreTablaE & "." & .fields(J).Name
         If InStr(UCase(.fields(J).Name), "Í") Then Print #NumFile2, NombreTablaE & "." & .fields(J).Name
         If InStr(UCase(.fields(J).Name), "Ó") Then Print #NumFile2, NombreTablaE & "." & .fields(J).Name
         If InStr(UCase(.fields(J).Name), "Ú") Then Print #NumFile2, NombreTablaE & "." & .fields(J).Name
     Next J
    End With
    Print #NumFile, "+";
    Print #NumFile, String$(89, "=");
    Print #NumFile, "+"
    RatonNormal
    Progreso_Barra.Mensaje_Box = LstTablas.List(I)
    Progreso_Esperar
    'MsgBox RutaSistema & "\BASES\UPDATE_DB\ABASES.TXT"
Next I
Close #NumFile2
Close #NumFile
End Sub

Public Sub Procesar_Duplicados_Clientes()
  RatonReloj
  Progreso_Barra.Incremento = 0
  Progreso_Barra.Valor_Maximo = 100
  
 'Eliminar Duplicados en el Catalogo de Cuentas
  Progreso_Barra.Incremento = Progreso_Barra.Incremento + 5
  Progreso_Barra.Mensaje_Box = "Determinando Duplicados de: Catalogo de Cuentas "
  Progreso_Esperar
  Eliminar_Duplicados_SP "Catalogo_Cuentas", "Codigo"
  
 'Codigo Repetidos de Clientes
  Progreso_Barra.Incremento = Progreso_Barra.Incremento + 5
  Progreso_Barra.Mensaje_Box = "Determinando Duplicados de: Codigos de Clientes "
  Progreso_Esperar
  Eliminar_Duplicados_SP "Clientes", "Codigo"
  
 'CI/RUC Repetidos
  Progreso_Barra.Incremento = Progreso_Barra.Incremento + 5
  Progreso_Barra.Mensaje_Box = "Determinando Duplicados de: Cédula o RUC "
  Progreso_Esperar
  Eliminar_Duplicados_SP "Clientes", "CI_RUC"
 
  Progreso_Barra.Incremento = Progreso_Barra.Incremento + 5
  Progreso_Barra.Mensaje_Box = "Determinando Duplicados de Nombres "
  Progreso_Esperar
  Eliminar_Duplicados_SP "Clientes", "Cliente"
  
  Progreso_Barra.Incremento = Progreso_Barra.Incremento + 5
  Progreso_Barra.Mensaje_Box = "Determinando Duplicados de: Notas "
  Progreso_Esperar
  Eliminar_Duplicados_SP "Trans_Notas", "Item, Periodo, Codigo, CodE, CodMat"
  
 'Codigo Repetidos de Clientes Matriculados
  Progreso_Barra.Incremento = Progreso_Barra.Incremento + 3
  Progreso_Barra.Mensaje_Box = "Determinando Duplicados de: Clientes de Matriculas"
  Progreso_Esperar
  Eliminar_Duplicados_SP "Clientes_Matriculas", "Item, Periodo, Codigo"
    
 'Codigo Repetidos de Catalogo Rol Rubros
  Progreso_Barra.Mensaje_Box = "Determinando Duplicados de: Catalogo Rol Rubros"
  Progreso_Esperar
  Eliminar_Duplicados_SP "Catalogo_Rol_Rubros", "Codigo, Cod_Rol_Pago, Cta,Valor, I_E"
  
 'Codigo Repetidos de Trans Notas
  Progreso_Barra.Mensaje_Box = "Determinando Duplicados de: Clientes de Matriculas"
  Progreso_Esperar
  sSQL = "SELECT * " _
       & "FROM Trans_Notas " _
       & "WHERE Item = '" & NumEmpresa & "' "
  Select_Adodc AdoCuentas, sSQL
  ReDim VectFields(AdoCuentas.Recordset.fields.Count + 1) As Variant
  
  sSQL = "SELECT Item,Periodo,Codigo,CodE,CodMat,COUNT(CodMat) As NumItem " _
       & "FROM Trans_Notas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "GROUP BY Item,Periodo,Codigo,CodE,CodMat " _
       & "HAVING COUNT(CodMat) > 1 "
  Select_Adodc AdoComp, sSQL
  With AdoComp.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       Progreso_Barra.Incremento = 0
       Progreso_Barra.Valor_Maximo = .RecordCount
       Progreso_Barra.Mensaje_Box = "Codigo de Clientes de Notas"
       Do While Not .EOF
          'MsgBox ".............. "
          Codigo = .fields("Codigo")
          Codigo1 = .fields("Item")
          Codigo2 = .fields("Periodo")
          Codigo3 = .fields("CodE")
          Codigo4 = .fields("CodMat")
          sSQL = "SELECT * " _
               & "FROM Trans_Notas " _
               & "WHERE Codigo = '" & Codigo & "' " _
               & "AND Item = '" & Codigo1 & "' " _
               & "AND Periodo = '" & Codigo2 & "' " _
               & "AND CodE = '" & Codigo3 & "' " _
               & "AND CodMat = '" & Codigo4 & "' " _
               & "ORDER BY Item,Periodo,Codigo,CodE,CodMat "
          Select_Adodc AdoCuentas, sSQL
          If AdoCuentas.Recordset.RecordCount > 0 Then
             Progreso_Barra.Mensaje_Box = "Codigo Matricula: " & Codigo
             For I = 0 To AdoCuentas.Recordset.fields.Count - 1
                VectFields(I) = AdoCuentas.Recordset.fields(I)
             Next I
             sSQL = "DELETE * " _
                  & "FROM Clientes_Matriculas " _
                  & "WHERE Codigo = '" & Codigo & "' " _
                  & "AND Item = '" & Codigo1 & "' " _
                  & "AND Periodo = '" & Codigo2 & "' " _
                  & "AND CodE = '" & Codigo3 & "' " _
                  & "AND CodMat = '" & Codigo4 & "' "
             Ejecutar_SQL_SP sSQL
             SetAddNew AdoCuentas
             For I = 0 To AdoCuentas.Recordset.fields.Count - 1
                 SetFields AdoCuentas, AdoCuentas.Recordset.fields(I).Name, VectFields(I)
             Next I
             SetUpdate AdoCuentas
          End If
          Progreso_Esperar
         .MoveNext
       Loop
   End If
  End With
 'Determinamos que todas las facturas estan bien el en Cliente/Proveedor
  Progreso_Barra.Incremento = Progreso_Barra.Incremento + 3
  Progreso_Barra.Mensaje_Box = "Verificación de Facturas sin Clientes "
  Progreso_Esperar
  Actualizar_Codigo_Clientes "Facturas", "CodigoC"
  Progreso_Barra.Incremento = Progreso_Barra.Incremento + 5
  Progreso_Barra.Mensaje_Box = "Verificación de Detalle Facturas sin Clientes "
  Progreso_Esperar
  Actualizar_Codigo_Clientes "Detalle_Factura", "CodigoC"
  Progreso_Barra.Incremento = Progreso_Barra.Incremento + 5
  Progreso_Barra.Mensaje_Box = "Verificación de Abonos sin Clientes "
  Progreso_Esperar
  Actualizar_Codigo_Clientes "Trans_Abonos", "CodigoC"
  Progreso_Barra.Incremento = Progreso_Barra.Incremento + 5
  Progreso_Barra.Mensaje_Box = "Verificación de Compras sin Clientes "
  Progreso_Esperar
  Actualizar_Codigo_Clientes "Trans_Compras", "IdProv"
  Progreso_Barra.Incremento = Progreso_Barra.Incremento + 5
  Progreso_Barra.Mensaje_Box = "Verificación de Ventas sin Clientes "
  Progreso_Esperar
  Actualizar_Codigo_Clientes "Trans_Ventas", "IdProv"
  Progreso_Barra.Incremento = Progreso_Barra.Incremento + 5
  Progreso_Barra.Mensaje_Box = "Verificación de Importaciones sin Clientes "
  Progreso_Esperar
  Actualizar_Codigo_Clientes "Trans_Importaciones", "IdFiscalProv"
  Progreso_Barra.Incremento = Progreso_Barra.Incremento + 5
  Progreso_Barra.Mensaje_Box = "Verificación de Exportaciones sin Clientes "
  Progreso_Esperar
  Actualizar_Codigo_Clientes "Trans_Exportaciones", "IdFiscalProv"
  Progreso_Barra.Incremento = Progreso_Barra.Incremento + 5
  Progreso_Barra.Mensaje_Box = "Verificación de Air sin Clientes "
  Progreso_Esperar
  Actualizar_Codigo_Clientes "Trans_Air", "IdProv"
  Progreso_Barra.Incremento = Progreso_Barra.Incremento + 5
  Progreso_Barra.Mensaje_Box = "Verificación de consistencia en Facturas "
  Progreso_Esperar
  Progreso_Barra.Incremento = Progreso_Barra.Incremento + 5
  Progreso_Barra.Mensaje_Box = "Verificación de Rol de Pagos "
  Progreso_Esperar
  Actualizar_Codigo_Clientes "Catalogo_Rol_Pagos", "Codigo"
  Progreso_Barra.Incremento = Progreso_Barra.Incremento + 5
  Progreso_Esperar
  Actualizar_Codigo_Clientes "Catalogo_Rol_Rubros", "Codigo"
  Progreso_Barra.Incremento = Progreso_Barra.Incremento + 5
  Progreso_Esperar
  Actualizar_Codigo_Clientes "Trans_Rol_de_Pagos", "Codigo"
  Progreso_Barra.Incremento = Progreso_Barra.Incremento + 5
  Progreso_Esperar
  Actualizar_Codigo_Clientes "Trans_Rol_Horas", "Codigo"
  Progreso_Barra.Incremento = Progreso_Barra.Incremento + 5
  Progreso_Esperar
 'Enceramos los codigos
  sSQL = "UPDATE Trans_Abonos " _
       & "SET X = '.' " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' "
  Ejecutar_SQL_SP sSQL
  sSQL = "UPDATE Detalle_Factura " _
       & "SET X = '.' " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' "
  Ejecutar_SQL_SP sSQL
  sSQL = "UPDATE Facturas " _
       & "SET X = '.' " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' "
  Ejecutar_SQL_SP sSQL
 'Verificamos si existe clientes en la factura
  If SQL_Server Then
     sSQL = "UPDATE Facturas " _
          & "SET X = 'S' " _
          & "FROM Facturas As F,Clientes As C "
  Else
     sSQL = "UPDATE Facturas As F,Clientes As C " _
          & "SET F.X = 'S' "
  End If
  sSQL = sSQL _
       & "WHERE F.Item = '" & NumEmpresa & "' " _
       & "AND F.Periodo = '" & Periodo_Contable & "' " _
       & "AND F.CodigoC = C.Codigo "
  Ejecutar_SQL_SP sSQL
 'Colocamos Consumidor Final a las facturas inconsistentes
  sSQL = "UPDATE Facturas " _
       & "SET CodigoC = '9999999999' " _
       & "WHERE X = '.' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' "
  Ejecutar_SQL_SP sSQL
 'Determinamos si existe el Cliente en el Detalle de la Factura
  If SQL_Server Then
     sSQL = "UPDATE Detalle_Factura " _
          & "SET X = F.X " _
          & "FROM Detalle_Factura As DF,Facturas As F "
  Else
     sSQL = "UPDATE Detalle_Factura As DF,Facturas As F " _
          & "SET DF.X = F.X "
  End If
  sSQL = sSQL & "WHERE F.Item = '" & NumEmpresa & "' " _
       & "AND F.Periodo = '" & Periodo_Contable & "' " _
       & "AND F.CodigoC = DF.CodigoC " _
       & "AND F.Factura = DF.Factura " _
       & "AND F.Fecha = DF.Fecha " _
       & "AND F.Item = DF.Item " _
       & "AND F.Periodo = DF.Periodo " _
       & "AND F.TC = DF.TC "
  Ejecutar_SQL_SP sSQL
 'Determinamos si existe el Cliente en Abonno de la Factura
  If SQL_Server Then
     sSQL = "UPDATE Trans_Abonos " _
          & "SET X = F.X " _
          & "FROM Trans_Abonos As DF,Facturas As F "
  Else
     sSQL = "UPDATE Trans_Abonos As DF,Facturas As F " _
          & "SET DF.X = F.X "
  End If
  sSQL = sSQL & "WHERE F.Item = '" & NumEmpresa & "' " _
       & "AND F.Periodo = '" & Periodo_Contable & "' " _
       & "AND F.CodigoC = DF.CodigoC " _
       & "AND F.Factura = DF.Factura " _
       & "AND F.Item = DF.Item " _
       & "AND F.Periodo = DF.Periodo " _
       & "AND F.TC = DF.TP "
  Ejecutar_SQL_SP sSQL
 'Colocamos Consumidor Final a las facturas inconsistentes
  sSQL = "UPDATE Detalle_Factura " _
       & "SET CodigoC = '9999999999' " _
       & "WHERE X = '.' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' "
  Ejecutar_SQL_SP sSQL
 'Colocamos Consumidor Final a las facturas inconsistentes
  sSQL = "UPDATE Trans_Abonos " _
       & "SET CodigoC = '9999999999' " _
       & "WHERE X = '.' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' "
  Ejecutar_SQL_SP sSQL
 'Codigo Generales del Sistema que deben estar obligatoriamente en el Sistema
 'para su correcto funcionamento y control de procesos contables y modulos.
  sSQL = "UPDATE Clientes " _
       & "SET Prov = '" & CodigoProv & "'," _
       & "    Pais = '" & CodigoPais & "'," _
       & "    Ciudad = '" & UCaseStrg(NombreCiudad) & "' " _
       & "WHERE Prov IN ('.','00') "
  Ejecutar_SQL_SP sSQL
  sSQL = "DELETE * " _
       & "FROM Clientes " _
       & "WHERE Codigo IN ('.','..','9999999999','8888888888') "
  Ejecutar_SQL_SP sSQL
  sSQL = "DELETE * " _
       & "FROM Clientes " _
       & "WHERE Cliente IN ('.','BODEGUERO(A)','CONSUMIDOR FINAL') "
  Ejecutar_SQL_SP sSQL
 'Generamos los códigos obligatorios del sistema
  SetAdoAddNew "Clientes", True
  SetAdoFields "Codigo", "."
  SetAdoFields "Cliente", "."
  SetAdoFields "TD", "O"
  SetAdoFields "CI_RUC", String$(9, "0")
  SetAdoFields "T", "N"
  SetAdoFields "Ciudad", UCaseStrg(NombreCiudad)
  SetAdoFields "Grupo", String$(6, "9")
  SetAdoFields "Telefono", Telefono1
  SetAdoFields "TelefonoT", Telefono2
  SetAdoFields "FAX", FAX
  SetAdoFields "Celular", "099000000"
  SetAdoFields "Direccion", "SD"
  SetAdoFields "DirNumero", "SN"
  SetAdoFields "Pais", CodigoPais
  SetAdoFields "Prov", CodigoProv
  SetAdoUpdate
     
  SetAdoAddNew "Clientes", True
  SetAdoFields "Codigo", ".."
  SetAdoFields "Cliente", "T O T A L E S"
  SetAdoFields "TD", "O"
  SetAdoFields "CI_RUC", String$(9, "1")
  SetAdoFields "T", "N"
  SetAdoFields "Ciudad", UCaseStrg(NombreCiudad)
  SetAdoFields "Grupo", String$(6, "9")
  SetAdoFields "Telefono", Telefono1
  SetAdoFields "TelefonoT", Telefono2
  SetAdoFields "FAX", FAX
  SetAdoFields "Celular", "099000000"
  SetAdoFields "Direccion", "SD"
  SetAdoFields "DirNumero", "SN"
  SetAdoFields "Pais", CodigoPais
  SetAdoFields "Prov", CodigoProv
  SetAdoUpdate
     
  SetAdoAddNew "Clientes", True
  SetAdoFields "Codigo", String$(10, "8")
  SetAdoFields "Cliente", "BODEGUERO(A)"
  SetAdoFields "TD", "O"
  SetAdoFields "CI_RUC", String$(9, "8")
  SetAdoFields "T", "N"
  SetAdoFields "Ciudad", UCaseStrg(NombreCiudad)
  SetAdoFields "Grupo", String$(6, "9")
  SetAdoFields "Telefono", Telefono1
  SetAdoFields "TelefonoT", Telefono2
  SetAdoFields "FAX", FAX
  SetAdoFields "Celular", "099000000"
  SetAdoFields "Direccion", "SD"
  SetAdoFields "DirNumero", "SN"
  SetAdoFields "Pais", CodigoPais
  SetAdoFields "Prov", CodigoProv
  SetAdoUpdate
     
  SetAdoAddNew "Clientes", True
  SetAdoFields "Codigo", String$(10, "9")
  SetAdoFields "Cliente", "CONSUMIDOR FINAL"
  SetAdoFields "TD", "R"
  SetAdoFields "CI_RUC", String$(13, "9")
  SetAdoFields "T", "N"
  SetAdoFields "FA", True
  SetAdoFields "Ciudad", UCaseStrg(NombreCiudad)
  SetAdoFields "Grupo", String$(6, "9")
  SetAdoFields "Telefono", Telefono1
  SetAdoFields "TelefonoT", Telefono2
  SetAdoFields "FAX", FAX
  SetAdoFields "Celular", "099000000"
  SetAdoFields "Direccion", "SD"
  SetAdoFields "DirNumero", "SN"
  SetAdoFields "Pais", CodigoPais
  SetAdoFields "Prov", CodigoProv
  SetAdoUpdate
  RatonNormal
  Progreso_Barra.Incremento = Progreso_Barra.Valor_Maximo
  Progreso_Esperar
  Progreso_Barra.Incremento = 100
  Progreso_Esperar
End Sub

Public Sub Procesar_Duplicados_Rubros_Clientes_Facturacion()
Dim CantCar As Byte
Dim NombreRep As String
  Progreso_Barra.Mensaje_Box = "Duplicados Rubros Clientes Facturacion"
  Progreso_Iniciar
  sSQL = "SELECT * " _
       & "FROM Clientes_Facturacion " _
       & "WHERE Codigo = '.' "
  Select_Adodc AdoCuentas, sSQL
  ReDim VectFields(AdoCuentas.Recordset.fields.Count + 1) As Variant

  sSQL = "SELECT Codigo, Periodo, Num_Mes, Codigo_Inv, COUNT(Codigo_Inv) As CantRubros " _
       & "FROM Clientes_Facturacion " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "GROUP BY Codigo, Periodo, Num_Mes, Codigo_Inv " _
       & "HAVING COUNT(Codigo_Inv) > 1 " _
       & "ORDER BY Codigo, Periodo, Num_Mes, Codigo_Inv "
  Select_Adodc AdoComp, sSQL
  With AdoComp.Recordset
   If .RecordCount > 0 Then
       Contador = 0
       Progreso_Barra.Valor_Maximo = Progreso_Barra.Valor_Maximo + .RecordCount
       Progreso_Barra.Mensaje_Box = "Clientes Facturacion"
       Do While Not .EOF
          CantCar = 0
          Codigo = .fields("Codigo")
          Codigo1 = .fields("Periodo")
          Codigo2 = .fields("Codigo_Inv")
          NumComp = .fields("Num_Mes")
          Progreso_Barra.Mensaje_Box = "Duplicado Clientes Facturacion: " & Codigo
          Progreso_Esperar
          sSQL = "SELECT * " _
               & "FROM Clientes_Facturacion " _
               & "WHERE Codigo = '" & Codigo & "' " _
               & "AND Periodo = '" & Codigo1 & "' " _
               & "AND Codigo_Inv = '" & Codigo2 & "' " _
               & "AND Num_Mes = " & NumComp & " " _
               & "AND Item = '" & NumEmpresa & "' "
          Select_Adodc AdoCuentas, sSQL
          If AdoCuentas.Recordset.RecordCount > 0 Then
             For I = 0 To AdoCuentas.Recordset.fields.Count - 1
                VectFields(I) = AdoCuentas.Recordset.fields(I)
             Next I
             sSQL = "DELETE * " _
                  & "FROM Clientes_Facturacion " _
                  & "WHERE Codigo = '" & Codigo & "' " _
                  & "AND Periodo = '" & Codigo1 & "' " _
                  & "AND Codigo_Inv = '" & Codigo2 & "' " _
                  & "AND Num_Mes = " & NumComp & " " _
                  & "AND Item = '" & NumEmpresa & "' "
             Ejecutar_SQL_SP sSQL
             
            'Creamos el Rubro no duplicado
             SetAddNew AdoCuentas
             For I = 0 To AdoCuentas.Recordset.fields.Count - 1
                 SetFields AdoCuentas, AdoCuentas.Recordset.fields(I).Name, VectFields(I)
             Next I
             SetUpdate AdoCuentas
          End If
         .MoveNext
       Loop
   End If
  End With
  Progreso_Final
End Sub

Public Sub Procesar_Duplicados_Clientes_Facturacion()
Dim CantCar As Byte
Dim NombreRep As String
  Progreso_Barra.Mensaje_Box = "Duplicados Facturas y Abonos"
  Progreso_Iniciar
  sSQL = "SELECT TOP 1 * " _
       & "FROM Clientes_Facturacion " _
       & "WHERE Valor = 0 "
  Select_Adodc AdoCuentas, sSQL
  ReDim VectFields(AdoCuentas.Recordset.fields.Count + 1) As Variant

  sSQL = "SELECT Codigo, Periodo, Num_Mes, Codigo_Inv, Valor, COUNT (Periodo) As Duplicados " _
       & "FROM Clientes_Facturacion " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "GROUP BY Codigo, Periodo, Num_Mes, Codigo_Inv, Valor " _
       & "HAVING Count(Periodo) > 1 " _
       & "ORDER BY Codigo, Periodo, Num_Mes, Codigo_Inv, Valor "
  Select_Adodc AdoComp, sSQL
  With AdoComp.Recordset
   If .RecordCount > 0 Then
       Contador = 0
       Progreso_Barra.Valor_Maximo = Progreso_Barra.Valor_Maximo + .RecordCount
       Progreso_Barra.Mensaje_Box = "Facturas"
       Do While Not .EOF
          Progreso_Barra.Mensaje_Box = "Duplicado Clientes Facturacion: " & .fields("Codigo")
          Progreso_Esperar
          sSQL = "SELECT * " _
               & "FROM Clientes_Facturacion " _
               & "WHERE Item = '" & NumEmpresa & "' " _
               & "AND Periodo = '" & .fields("Periodo") & "' " _
               & "AND Codigo = '" & .fields("Codigo") & "' " _
               & "AND Codigo_Inv = '" & .fields("Codigo_Inv") & "' " _
               & "AND Num_Mes = " & .fields("Num_Mes") & " "
          Select_Adodc AdoCuentas, sSQL
          If AdoCuentas.Recordset.RecordCount > 0 Then
             For I = 0 To AdoCuentas.Recordset.fields.Count - 1
                VectFields(I) = AdoCuentas.Recordset.fields(I)
             Next I
             sSQL = "DELETE * " _
                  & "FROM Clientes_Facturacion " _
                  & "WHERE Item = '" & NumEmpresa & "' " _
                  & "AND Periodo = '" & .fields("Periodo") & "' " _
                  & "AND Codigo = '" & .fields("Codigo") & "' " _
                  & "AND Codigo_Inv = '" & .fields("Codigo_Inv") & "' " _
                  & "AND Num_Mes = " & .fields("Num_Mes") & " "
             Ejecutar_SQL_SP sSQL
             
            'Creamos el Rubro no duplicado
             SetAddNew AdoCuentas
             For I = 0 To AdoCuentas.Recordset.fields.Count - 1
                 SetFields AdoCuentas, AdoCuentas.Recordset.fields(I).Name, VectFields(I)
             Next I
             SetUpdate AdoCuentas
          End If
         .MoveNext
       Loop
   End If
  End With
  Progreso_Final
End Sub

Public Sub Procesar_Duplicados_Representantes()
Dim CantCar As Byte
Dim NombreRep As String
Dim Fecha_Rep As String
  
  Fecha_Rep = BuscarFecha("01/01/2016")
  
  Progreso_Barra.Mensaje_Box = "Clientes de Matriculas"
  Progreso_Iniciar
  sSQL = "SELECT Cedula_R, COUNT(Cedula_R) As CIRUC " _
       & "FROM Clientes_Matriculas " _
       & "WHERE Periodo = '" & Periodo_Contable & "' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "GROUP BY Cedula_R " _
       & "HAVING COUNT(Cedula_R) > 1 "
  Select_Adodc AdoComp, sSQL
  With AdoComp.Recordset
   If .RecordCount > 0 Then
       Contador = 0
       Progreso_Barra.Valor_Maximo = Progreso_Barra.Valor_Maximo + .RecordCount + 2
       Progreso_Barra.Mensaje_Box = "Clientes de Matriculas"
       Do While Not .EOF
          Progreso_Barra.Mensaje_Box = "Representante: " & MidStrg(NombreRep, 1, 25) & "..."
          Progreso_Esperar
          CantCar = 0
          NombreRep = Ninguno
          Codigo = .fields("Cedula_R")
          sSQL = "SELECT Cedula_R,Representante,Item,Periodo,Codigo " _
               & "FROM Clientes_Matriculas " _
               & "WHERE Cedula_R = '" & Codigo & "' " _
               & "AND Periodo = '" & Periodo_Contable & "' " _
               & "AND Item = '" & NumEmpresa & "' " _
               & "ORDER BY Representante "
          Select_Adodc AdoCuentas, sSQL
          If AdoCuentas.Recordset.RecordCount > 0 Then
             Do While Not AdoCuentas.Recordset.EOF
                If Len(AdoCuentas.Recordset.fields("Representante")) > CantCar Then
                   NombreRep = AdoCuentas.Recordset.fields("Representante")
                   CantCar = Len(AdoCuentas.Recordset.fields("Representante"))
                End If
                AdoCuentas.Recordset.MoveNext
             Loop
          End If
          
          sSQL = "UPDATE Clientes_Matriculas " _
               & "SET Representante = '" & UCaseStrg(NombreRep) & "' " _
               & "WHERE Periodo = '" & Periodo_Contable & "' " _
               & "AND Item = '" & NumEmpresa & "' " _
               & "AND Cedula_R = '" & Codigo & "' " _
               & "AND Representante <> '" & UCaseStrg(NombreRep) & "' "
          Ejecutar_SQL_SP sSQL
         .MoveNext
       Loop
   End If
  End With
  
  Progreso_Barra.Mensaje_Box = "CONSUMIDOR FINAL"
  Progreso_Esperar
  sSQL = "UPDATE Facturas " _
       & "SET RUC_CI = '9999999999999', Razon_Social = 'CONSUMIDOR FINAL', TB = 'R' " _
       & "WHERE Fecha >= #" & Fecha_Rep & "# " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND Item = '" & NumEmpresa & "' "
  Ejecutar_SQL_SP sSQL
  
  sSQL = "UPDATE Trans_Ventas " _
       & "SET RUC_CI = '9999999999999', Razon_Social = 'CONSUMIDOR FINAL', TB = 'R' " _
       & "WHERE Fecha >= #" & Fecha_Rep & "# " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND Item = '" & NumEmpresa & "' "
  Ejecutar_SQL_SP sSQL
  
  Progreso_Barra.Mensaje_Box = "ACTUALIZANDO VENTAS"
  Progreso_Esperar
  If SQL_Server Then
     sSQL = "UPDATE Trans_Ventas " _
          & "SET RUC_CI = C.CI_RUC, Razon_Social = C.Cliente, TB = C.TD " _
          & "FROM Trans_Ventas As TV, Clientes As C "
  Else
     sSQL = "UPDATE Trans_Ventas As TV, Clientes As C " _
          & "SET TV.RUC_CI = C.CI_RUC, TV.Razon_Social = C.Cliente, TV.TB = C.TD "
  End If
  sSQL = sSQL _
       & "WHERE TV.Fecha >= #" & Fecha_Rep & "# " _
       & "AND TV.Periodo = '" & Periodo_Contable & "' " _
       & "AND TV.Item = '" & NumEmpresa & "' " _
       & "AND C.TD IN ('C','R','P') " _
       & "AND TV.IdProv = C.Codigo "
  Ejecutar_SQL_SP sSQL
  
  Progreso_Barra.Mensaje_Box = "ACTUALIZANDO FACTURAS"
  Progreso_Esperar
  If SQL_Server Then
     sSQL = "UPDATE Facturas " _
          & "SET RUC_CI = C.CI_RUC, Razon_Social = C.Cliente, TB = C.TD " _
          & "FROM Facturas As F, Clientes As C "
  Else
     sSQL = "UPDATE Facturas As F, Clientes As C " _
          & "SET F.RUC_CI = C.CI_RUC, F.Razon_Social = C.Cliente, F.TB = C.TD "
  End If
  sSQL = sSQL _
       & "WHERE F.Fecha >= #" & Fecha_Rep & "# " _
       & "AND F.Periodo = '" & Periodo_Contable & "' " _
       & "AND F.Item = '" & NumEmpresa & "' " _
       & "AND C.TD IN ('C','R','P') " _
       & "AND F.CodigoC = C.Codigo "
  Ejecutar_SQL_SP sSQL
  
  sSQL = "SELECT TOP 1 * " _
       & "FROM Clientes_Matriculas " _
       & "WHERE Periodo = '" & Periodo_Contable & "' " _
       & "AND Item = '" & NumEmpresa & "' "
  Select_Adodc AdoAux, sSQL
  If AdoAux.Recordset.RecordCount > 0 Then
     If SQL_Server Then
        sSQL = "UPDATE Facturas " _
             & "SET RUC_CI = CM.Cedula_R, Razon_Social = CM.Representante, TB = CM.TD " _
             & "FROM Facturas As F,Clientes_Matriculas As CM "
     Else
        sSQL = "UPDATE Facturas As F,Clientes_Matriculas As CM " _
             & "SET F.RUC_CI = CM.Cedula_R, F.Razon_Social = CM.Representante, F.TB = CM.TD "
     End If
     sSQL = sSQL _
          & "WHERE F.Fecha >= #" & Fecha_Rep & "# " _
          & "AND F.Periodo = '" & Periodo_Contable & "' " _
          & "AND F.Item = '" & NumEmpresa & "' " _
          & "AND CM.TD IN ('C','R','P') " _
          & "AND F.Periodo = CM.Periodo " _
          & "AND F.Item = CM.Item " _
          & "AND F.CodigoC = CM.Codigo "
     Ejecutar_SQL_SP sSQL
  End If
  Progreso_Final
End Sub

Public Sub Procesar_Duplicados_Usuarios()
Dim ITab As Long
Dim JCamp As Integer
Dim IdCampo As Integer

  RatonReloj
 'Codigos Catalogo Ctas_Proceso
  Progreso_Barra.Mensaje_Box = "Determinando Duplicados de: Ctas_Proceso"
  Progreso_Esperar
  Eliminar_Duplicados_SP "Ctas_Proceso", "Periodo, Item, Detalle"
 
 'Codigos Catalogo Seteos_Documentos
  Progreso_Barra.Mensaje_Box = "Determinando Duplicados de: Seteos Documentos"
  Progreso_Esperar
  Eliminar_Duplicados_SP "Seteos_Documentos", "Item, TP, Campo"
   
 'Codigos Catalogo Modulos
  sSQL = "SELECT * " _
       & "FROM Modulos " _
       & "WHERE Modulo <> '.' "
  Select_Adodc AdoCuentas, sSQL
  ReDim VectFields(AdoCuentas.Recordset.fields.Count + 1) As Variant
  sSQL = "SELECT Modulo, COUNT(Modulo) As NumItem " _
       & "FROM Modulos " _
       & "WHERE Modulo <> '.' " _
       & "GROUP BY Modulo " _
       & "HAVING COUNT(Modulo) > 1 "
  Select_Adodc AdoComp, sSQL
  With AdoComp.Recordset
   If .RecordCount > 0 Then
       Contador = 0
      .MoveFirst
       Do While Not .EOF
          Codigo1 = .fields("Modulo")
          Contador = Contador + 1
          FSeteos.Caption = Codigo1 & "-" & Codigo3 & " -> " & Format(Contador / .RecordCount, "00%")
          sSQL = "SELECT * " _
               & "FROM Modulos " _
               & "WHERE Modulo = '" & Codigo1 & "' "
          Select_Adodc AdoCuentas, sSQL
          If AdoCuentas.Recordset.RecordCount > 0 Then
             For I = 0 To AdoCuentas.Recordset.fields.Count - 1
                VectFields(I) = AdoCuentas.Recordset.fields(I)
             Next I
             sSQL = "DELETE * " _
                  & "FROM Modulos " _
                  & "WHERE Modulo = '" & Codigo1 & "' "
             Ejecutar_SQL_SP sSQL
             SetAddNew AdoCuentas
             For I = 0 To AdoCuentas.Recordset.fields.Count - 1
                 SetFields AdoCuentas, AdoCuentas.Recordset.fields(I).Name, VectFields(I)
             Next I
             SetUpdate AdoCuentas
          End If
         .MoveNext
       Loop
   End If
  End With
  
 
 'Formato
  sSQL = "SELECT * " _
       & "FROM Formato " _
       & "WHERE Item <> '.' "
  Select_Adodc AdoCuentas, sSQL
  ReDim VectFields(AdoCuentas.Recordset.fields.Count + 1) As Variant
  sSQL = "SELECT Item, TP, COUNT(TP) As NumItem " _
       & "FROM Formato " _
       & "WHERE Item <> '.' " _
       & "GROUP BY Item, TP " _
       & "HAVING COUNT(TP) > 1 "
  Select_Adodc AdoComp, sSQL
  With AdoComp.Recordset
   If .RecordCount > 0 Then
       Contador = 0
      .MoveFirst
       Do While Not .EOF
          Codigo1 = .fields("Item")
          Codigo2 = .fields("TP")
          Contador = Contador + 1
          FSeteos.Caption = Codigo1 & "-" & Codigo2 & " -> " & Format(Contador / .RecordCount, "00%")
          sSQL = "SELECT * " _
               & "FROM Formato " _
               & "WHERE Item = '" & Codigo1 & "' " _
               & "AND TP = '" & Codigo2 & "' "
          Select_Adodc AdoCuentas, sSQL
          If AdoCuentas.Recordset.RecordCount > 0 Then
             For I = 0 To AdoCuentas.Recordset.fields.Count - 1
                VectFields(I) = AdoCuentas.Recordset.fields(I)
             Next I
             sSQL = "DELETE * " _
                  & "FROM Formato " _
                  & "WHERE Item = '" & Codigo1 & "' " _
                  & "AND TP = '" & Codigo2 & "' "
             Ejecutar_SQL_SP sSQL
             SetAddNew AdoCuentas
             For I = 0 To AdoCuentas.Recordset.fields.Count - 1
                 SetFields AdoCuentas, AdoCuentas.Recordset.fields(I).Name, VectFields(I)
             Next I
             SetUpdate AdoCuentas
          End If
         .MoveNext
       Loop
   End If
  End With
  
 'Codigos Repetidos
  sSQL = "SELECT * " _
       & "FROM Codigos " _
       & "WHERE Item <> '.' "
  Select_Adodc AdoCuentas, sSQL
  ReDim VectFields(AdoCuentas.Recordset.fields.Count + 1) As Variant
  
  sSQL = "SELECT Periodo,Item,Concepto,COUNT(Concepto) As NumItem " _
       & "FROM Codigos " _
       & "WHERE Item <> '.' " _
       & "GROUP BY Periodo,Item,Concepto " _
       & "HAVING COUNT(Concepto) > 1 "
  Select_Adodc AdoComp, sSQL
  With AdoComp.Recordset
   If .RecordCount > 0 Then
       Contador = 0
      .MoveFirst
       Do While Not .EOF
          Codigo = .fields("Concepto")
          Codigo1 = .fields("Periodo")
          Codigo2 = .fields("Item")
          Contador = Contador + 1
          FSeteos.Caption = Codigo & " -> " & Format(Contador / .RecordCount, "00%")
          sSQL = "SELECT * " _
               & "FROM Codigos " _
               & "WHERE Concepto = '" & Codigo & "' " _
               & "AND Periodo = '" & Codigo1 & "' " _
               & "AND Item = '" & Codigo2 & "' "
          Select_Adodc AdoCuentas, sSQL
          If AdoCuentas.Recordset.RecordCount > 0 Then
             For I = 0 To AdoCuentas.Recordset.fields.Count - 1
                VectFields(I) = AdoCuentas.Recordset.fields(I)
             Next I
             sSQL = "DELETE * " _
                  & "FROM Codigos " _
                  & "WHERE Concepto = '" & Codigo & "' " _
                  & "AND Periodo = '" & Codigo1 & "' " _
                  & "AND Item = '" & Codigo2 & "' "
             Ejecutar_SQL_SP sSQL
             SetAddNew AdoCuentas
             For I = 0 To AdoCuentas.Recordset.fields.Count - 1
                 SetFields AdoCuentas, AdoCuentas.Recordset.fields(I).Name, VectFields(I)
             Next I
             SetUpdate AdoCuentas
          End If
         .MoveNext
       Loop
   End If
  End With
 
 'Campos de la Tabla a eliminar duplicados
  sSQL = "SELECT * " _
       & "FROM Accesos " _
       & "WHERE Codigo = '.' "
  Select_Adodc AdoCuentas, sSQL
  ReDim VectFields(AdoCuentas.Recordset.fields.Count + 1) As Variant
 'Codigo Repetidos
  sSQL = "SELECT Codigo,COUNT(Codigo) As NumItem " _
       & "FROM Accesos " _
       & "WHERE Usuario <> '*' " _
       & "GROUP BY Codigo " _
       & "HAVING COUNT(Codigo) > 1 "
  Select_Adodc AdoComp, sSQL
  With AdoComp.Recordset
   If .RecordCount > 0 Then
       Contador = 0
      .MoveFirst
       Do While Not .EOF
          Contador = Contador + 1
          Codigo = .fields("Codigo")
          FSeteos.Caption = Codigo & " -> " & Format(Contador / .RecordCount, "00%")
          sSQL = "SELECT * " _
               & "FROM Accesos " _
               & "WHERE Codigo = '" & Codigo & "'  "
          Select_Adodc AdoCuentas, sSQL
          If AdoCuentas.Recordset.RecordCount > 0 Then
             For I = 0 To AdoCuentas.Recordset.fields.Count - 1
                VectFields(I) = AdoCuentas.Recordset.fields(I)
             Next I
             sSQL = "DELETE * " _
                  & "FROM Accesos " _
                  & "WHERE Codigo = '" & Codigo & "'  "
             Ejecutar_SQL_SP sSQL
             SetAddNew AdoCuentas
             For I = 0 To AdoCuentas.Recordset.fields.Count - 1
                 SetFields AdoCuentas, AdoCuentas.Recordset.fields(I).Name, VectFields(I)
             Next I
             SetUpdate AdoCuentas
          End If
         .MoveNext
       Loop
   End If
  End With

' Usuarios y Claves
  sSQL = "SELECT Usuario,Clave,COUNT(Usuario) As NumItem " _
       & "FROM Accesos " _
       & "WHERE Usuario <> '*' " _
       & "GROUP BY Usuario,Clave " _
       & "HAVING COUNT(Usuario) > 1 "
  Select_Adodc AdoComp, sSQL
  With AdoComp.Recordset
   If .RecordCount > 0 Then
       Contador = 0
      .MoveFirst
       Do While Not .EOF
          Contador = Contador + 1
          Codigo = .fields("Usuario")
          FSeteos.Caption = Codigo & " -> " & Format(Contador / .RecordCount, "00%")
          
          sSQL = "SELECT * " _
               & "FROM Accesos " _
               & "WHERE Usuario = '" & Codigo & "'  "
          Select_Adodc AdoCuentas, sSQL
          If AdoCuentas.Recordset.RecordCount > 0 Then
             For I = 0 To AdoCuentas.Recordset.fields.Count - 1
                VectFields(I) = AdoCuentas.Recordset.fields(I)
             Next I
             sSQL = "DELETE * " _
                  & "FROM Accesos " _
                  & "WHERE Usuario = '" & Codigo & "'  "
             Ejecutar_SQL_SP sSQL
             SetAddNew AdoCuentas
             For I = 0 To AdoCuentas.Recordset.fields.Count - 1
                 SetFields AdoCuentas, AdoCuentas.Recordset.fields(I).Name, VectFields(I)
             Next I
             SetUpdate AdoCuentas
          End If
         .MoveNext
       Loop
   End If
  End With
  
  sSQL = "UPDATE Facturas " _
       & "SET X = '.' " _
       & "WHERE X <> '.' "
  Ejecutar_SQL_SP sSQL
  sSQL = "UPDATE Detalle_Factura " _
       & "SET X = '.' " _
       & "WHERE X <> '.' "
  Ejecutar_SQL_SP sSQL
  sSQL = "UPDATE Trans_Abonos " _
       & "SET X = '.' " _
       & "WHERE X <> '.' "
  Ejecutar_SQL_SP sSQL
  
  If SQL_Server Then
     sSQL = "UPDATE Facturas " _
          & "SET X= 'X' " _
          & "FROM Facturas As X,Accesos As C "
  Else
     sSQL = "UPDATE Facturas As X,Accesos As C " _
          & "SET X.X = 'X' "
  End If
  sSQL = sSQL & "WHERE X.CodigoU = C.Codigo "
  Ejecutar_SQL_SP sSQL
  If SQL_Server Then
     sSQL = "UPDATE Detalle_Factura " _
          & "SET X= 'X' " _
          & "FROM Detalle_Factura As X,Accesos As C "
  Else
     sSQL = "UPDATE Detalle_Factura As X,Accesos As C " _
          & "SET X.X = 'X' "
  End If
  sSQL = sSQL & "WHERE X.CodigoU = C.Codigo "
  Ejecutar_SQL_SP sSQL
  If SQL_Server Then
     sSQL = "UPDATE Trans_Abonos " _
          & "SET X = 'X' " _
          & "FROM Trans_Abonos As X,Accesos As C "
  Else
     sSQL = "UPDATE Trans_Abonos As X,Accesos As C " _
          & "SET X.X = 'X' "
  End If
  sSQL = sSQL & "WHERE X.CodigoU = C.Codigo "
  Ejecutar_SQL_SP sSQL
  
  sSQL = "UPDATE Facturas " _
       & "SET CodigoU = '.' " _
       & "WHERE X = '.' "
  Ejecutar_SQL_SP sSQL
  sSQL = "UPDATE Detalle_Factura " _
       & "SET CodigoU = '.' " _
       & "WHERE X = '.' "
  Ejecutar_SQL_SP sSQL
  sSQL = "UPDATE Trans_Abonos " _
       & "SET CodigoU = '.' " _
       & "WHERE X = '.' "
  Ejecutar_SQL_SP sSQL
 'Eliminamos los usuario incorrecto
  sSQL = "UPDATE Accesos " _
       & "SET X = '.' " _
       & "WHERE X <> '.' "
  Ejecutar_SQL_SP sSQL
 'Determinamos si el usuario esta activo
  If SQL_Server Then
     sSQL = "UPDATE Accesos " _
          & "SET X = 'X' " _
          & "FROM Accesos As A,Clientes As X "
  Else
     sSQL = "UPDATE Accesos As A,Clientes As X " _
          & "SET A.X = 'X' "
  End If
  sSQL = sSQL & "WHERE A.Codigo = X.Codigo "
  Ejecutar_SQL_SP sSQL
 'Ahora buscamos en las otras tablas
  RatonReloj
  For ITab = 0 To LstTablas.ListCount - 1
      Si_No = False
      sSQL = "SELECT * " _
           & "FROM " & LstTablas.List(ITab) & " "
      Select_Adodc AdoComp, sSQL
      With AdoComp.Recordset
       For JCamp = 0 To .fields.Count - 1
           If .fields(JCamp).Name = "CodigoU" Then Si_No = True
       Next JCamp
      End With
      If MidStrg(LstTablas.List(ITab), 1, 4) = "Tipo" Then Si_No = False
      If MidStrg(LstTablas.List(ITab), 1, 5) = "Tabla" Then Si_No = False
      If MidStrg(LstTablas.List(ITab), 1, 7) = "Asiento" Then Si_No = False
      If Si_No Then
         If SQL_Server Then
            sSQL = "UPDATE Accesos " _
                 & "SET X = 'X' " _
                 & "FROM Accesos As A," & LstTablas.List(ITab) & " As X "
         Else
            sSQL = "UPDATE Accesos As A," & LstTablas.List(ITab) & " As X " _
                 & "SET A.X = 'X' "
         End If
         sSQL = sSQL & "WHERE A.Codigo = X.CodigoU "
         Ejecutar_SQL_SP sSQL
      End If
  Next ITab
  sSQL = "DELETE * " _
       & "FROM Accesos " _
       & "WHERE X <> 'X' "
  Ejecutar_SQL_SP sSQL
  RatonNormal
End Sub

Public Sub Procesar_Duplicados_Compras_Retenciones()
 'Procesamos Retenciones
  Progreso_Barra.Mensaje_Box = "Duplicados Compras y Retenciones"
  Progreso_Iniciar
  sSQL = "SELECT * " _
       & "FROM Trans_Air " _
       & "WHERE Numero = 0 " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND Item = '" & NumEmpresa & "' "
  Select_Adodc AdoCuentas, sSQL
  ReDim VectFields(AdoCuentas.Recordset.fields.Count + 1) As Variant

  sSQL = "SELECT Item, Periodo, Fecha, TP, Numero, Factura_No, CodRet, IdProv, COUNT(Item) " _
       & "FROM Trans_Air " _
       & "WHERE Periodo = '" & Periodo_Contable & "' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "GROUP BY Item, Periodo, Fecha, TP, Numero, Factura_No, CodRet, IdProv " _
       & "HAVING COUNT(Item) > 1 " _
       & "ORDER BY Item, Periodo, Fecha, TP, Numero, Factura_No, CodRet, IdProv "
  Select_Adodc AdoComp, sSQL
  With AdoComp.Recordset
   If .RecordCount > 0 Then
       Contador = 0
       Progreso_Barra.Valor_Maximo = Progreso_Barra.Valor_Maximo + .RecordCount
       Progreso_Barra.Mensaje_Box = "Retenciones"
       Do While Not .EOF
          Factura_No = .fields("Factura_No")
          Mifecha = .fields("Fecha")
          Codigo = .fields("Item")
          Codigo1 = .fields("Periodo")
          Codigo2 = .fields("CodRet")
          Codigo3 = .fields("TP")
          CodigoB = .fields("IdProv")
          Numero = .fields("Numero")
          
          Progreso_Barra.Mensaje_Box = "Duplicado Retencion: " & Format(Factura_No, "000000000") & " - " & CodigoB
          Progreso_Esperar
          sSQL = "SELECT * " _
               & "FROM Trans_Air " _
               & "WHERE Item = '" & Codigo & "' " _
               & "AND Periodo = '" & Codigo1 & "' " _
               & "AND CodRet = '" & Codigo2 & "' " _
               & "AND IdProv = '" & CodigoB & "' " _
               & "AND Factura_No = " & Factura_No & " " _
               & "AND TP = '" & Codigo3 & "' " _
               & "AND Numero = " & Numero & " " _
               & "AND Fecha = #" & BuscarFecha(Mifecha) & "# "
          Select_Adodc AdoCuentas, sSQL
          If AdoCuentas.Recordset.RecordCount > 0 Then
             For I = 0 To AdoCuentas.Recordset.fields.Count - 1
                VectFields(I) = AdoCuentas.Recordset.fields(I)
             Next I
             sSQL = "DELETE * " _
                  & "FROM Trans_Air " _
                  & "WHERE Item = '" & Codigo & "' " _
                  & "AND Periodo = '" & Codigo1 & "' " _
                  & "AND CodRet = '" & Codigo2 & "' " _
                  & "AND IdProv = '" & CodigoB & "' " _
                  & "AND Factura_No = " & Factura_No & " " _
                  & "AND TP = '" & Codigo3 & "' " _
                  & "AND Numero = " & Numero & " " _
                  & "AND Fecha = #" & BuscarFecha(Mifecha) & "# "
             Ejecutar_SQL_SP sSQL
             
            'Creamos el Rubro no duplicado
             SetAddNew AdoCuentas
             For I = 0 To AdoCuentas.Recordset.fields.Count - 1
                 SetFields AdoCuentas, AdoCuentas.Recordset.fields(I).Name, VectFields(I)
             Next I
             SetUpdate AdoCuentas
          End If
         .MoveNext
       Loop
   End If
  End With
  
 'Procesamos Compras
  Progreso_Barra.Mensaje_Box = "Duplicados Compras"
  Progreso_Iniciar
  sSQL = "SELECT * " _
       & "FROM Trans_Compras " _
       & "WHERE Numero = 0 " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND Item = '" & NumEmpresa & "' "
  Select_Adodc AdoCuentas, sSQL
  ReDim VectFields(AdoCuentas.Recordset.fields.Count + 1) As Variant

  sSQL = "SELECT Item, Periodo, Fecha, TP, Numero, Secuencial, Autorizacion, IdProv, COUNT(Item) " _
       & "FROM Trans_Compras " _
       & "WHERE Periodo = '" & Periodo_Contable & "' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "GROUP BY Item, Periodo, Fecha, TP, Numero, Secuencial, Autorizacion, IdProv " _
       & "HAVING COUNT(Item) > 1 " _
       & "ORDER BY Item, Periodo, Fecha, TP, Numero, Secuencial, Autorizacion, IdProv "
  Select_Adodc AdoComp, sSQL
  With AdoComp.Recordset
   If .RecordCount > 0 Then
       Contador = 0
       Progreso_Barra.Valor_Maximo = Progreso_Barra.Valor_Maximo + .RecordCount
       Progreso_Barra.Mensaje_Box = "Compras"
       Do While Not .EOF
          Factura_No = .fields("Secuencial")
          Mifecha = .fields("Fecha")
          Codigo = .fields("Item")
          Codigo1 = .fields("Periodo")
          Codigo2 = .fields("Autorizacion")
          Codigo3 = .fields("TP")
          CodigoB = .fields("IdProv")
          Numero = .fields("Numero")
          
          Progreso_Barra.Mensaje_Box = "Duplicado Compras: " & Format(Factura_No, "000000000") & " - " & CodigoB
          Progreso_Esperar
          sSQL = "SELECT * " _
               & "FROM Trans_Compras " _
               & "WHERE Item = '" & Codigo & "' " _
               & "AND Periodo = '" & Codigo1 & "' " _
               & "AND Autorizacion = '" & Codigo2 & "' " _
               & "AND IdProv = '" & CodigoB & "' " _
               & "AND Secuencial = " & Factura_No & " " _
               & "AND TP = '" & Codigo3 & "' " _
               & "AND Numero = " & Numero & " " _
               & "AND Fecha = #" & BuscarFecha(Mifecha) & "# "
          Select_Adodc AdoCuentas, sSQL
          If AdoCuentas.Recordset.RecordCount > 0 Then
             For I = 0 To AdoCuentas.Recordset.fields.Count - 1
                VectFields(I) = AdoCuentas.Recordset.fields(I)
             Next I
             sSQL = "DELETE * " _
                  & "FROM Trans_Compras " _
                  & "WHERE Item = '" & Codigo & "' " _
                  & "AND Periodo = '" & Codigo1 & "' " _
                  & "AND Autorizacion = '" & Codigo2 & "' " _
                  & "AND IdProv = '" & CodigoB & "' " _
                  & "AND Secuencial = " & Factura_No & " " _
                  & "AND TP = '" & Codigo3 & "' " _
                  & "AND Numero = " & Numero & " " _
                  & "AND Fecha = #" & BuscarFecha(Mifecha) & "# "
             Ejecutar_SQL_SP sSQL
             
            'Creamos el Rubro no duplicado
             SetAddNew AdoCuentas
             For I = 0 To AdoCuentas.Recordset.fields.Count - 1
                 SetFields AdoCuentas, AdoCuentas.Recordset.fields(I).Name, VectFields(I)
             Next I
             SetUpdate AdoCuentas
          End If
         .MoveNext
       Loop
   End If
  End With
  Progreso_Final
End Sub

Public Sub Borrar_Datos_Estudiantes()
Dim IdEst As Integer
Dim IdCampo As Integer
Dim Si_Item As Boolean
  For IdEst = 0 To LstTablas.ListCount - 1
      Si_Item = False
      sSQL = "SELECT * " _
           & "FROM " & LstTablas.List(IdEst) & " "
      Select_Adodc AdoAux, sSQL
      With AdoAux.Recordset
       For IdCampo = 0 To .fields.Count - 1
           If .fields(IdCampo).Name = "Item" Then Si_Item = True
       Next IdCampo
      End With
      If Si_Item Then
         sSQL = "DELETE * " _
              & "FROM " & LstTablas.List(IdEst) & " " _
              & "WHERE Item <> '000' "
         Ejecutar_SQL_SP sSQL
      End If
  Next IdEst
  sSQL = "UPDATE Accesos " _
       & "SET X = '.' "
  Ejecutar_SQL_SP sSQL
  
  sSQL = "UPDATE Accesos " _
       & "SET X = 'N' " _
       & "WHERE MidStrg(Codigo,1,6) = 'ACCESO' "
  Ejecutar_SQL_SP sSQL
  sSQL = "UPDATE Accesos " _
       & "SET X = 'N' " _
       & "WHERE Codigo = '0702164179' "
  Ejecutar_SQL_SP sSQL
  sSQL = "DELETE * " _
       & "FROM Accesos " _
       & "WHERE X = '.' "
  Ejecutar_SQL_SP sSQL

End Sub

Public Function Codigo_Usuario(Nombre_User As String) As String
Dim CodigoUser As String
  CodigoUser = Ninguno
  If IsNull(Nombre_User) Then Nombre_User = ""
  If Len(Nombre_User) < 1 Then Nombre_User = ""
  With AdoUsuario.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Nombre_Completo = '" & Nombre_User & "' ")
       If Not .EOF Then CodigoUser = .fields("Codigo")
   End If
  End With
  Codigo_Usuario = CodigoUser
End Function


Public Sub Generar_Documentos_Electronicos()
Dim FileNext As String
Dim DatosXMLA As String
Dim Clave_Acceso As String
Dim RutaXMLAutorizado As String
Dim IdxFile As Long
Dim ContFile As Long
   RatonReloj
   Progreso_Barra.Incremento = 0
   Progreso_Barra.Valor_Maximo = 1000
   Progreso_Barra.Mensaje_Box = "Determinando cantidad de Archivos"
   Progreso_Esperar
   ContFile = 0
   RutaXMLAutorizado = RutaSysBases & "\CE\CE" & NumEmpresa & "\Comprobantes Autorizados\"
   FileNext = Dir$(RutaXMLAutorizado, vbNormal) 'Recupera la primera entrada.
   Do While FileNext <> ""
      If FileNext <> "." And FileNext <> ".." Then
         If (GetAttr(RutaSysBases & "\CE\CE" & NumEmpresa & "\Comprobantes Autorizados\" & FileNext) And vbNormal) = vbNormal Then
            If UCaseStrg(MidStrg(FileNext, Len(FileNext) - 2, 3)) = "XML" Then ContFile = ContFile + 1
         End If
      End If
      FileNext = Dir$
   Loop
   Progreso_Barra.Valor_Maximo = ContFile + 100
   Progreso_Barra.Mensaje_Box = "Verificando Archivos"
   Progreso_Esperar
   IdxFile = 0
   ReDim ListaFile(ContFile) As String
   RutaXMLAutorizado = RutaSysBases & "\CE\CE" & NumEmpresa & "\Comprobantes Autorizados\"
   FileNext = Dir$(RutaXMLAutorizado, vbNormal) 'Recupera la primera entrada.
   Do While FileNext <> ""
      If FileNext <> "." And FileNext <> ".." Then
         If (GetAttr(RutaXMLAutorizado & FileNext) And vbNormal) = vbNormal Then
            If UCaseStrg(MidStrg(FileNext, Len(FileNext) - 2, 3)) = "XML" Then
               ListaFile(IdxFile) = FileNext
               IdxFile = IdxFile + 1
            End If
         End If
      End If
      FileNext = Dir$
   Loop
   For IdxFile = 0 To ContFile - 1
       FileNext = ListaFile(IdxFile)
       Progreso_Barra.Mensaje_Box = "Importando Archivos: " & FileNext
       Progreso_Esperar
      'ddmmyyyy01RUCAmbienteSerieFacturaddmmyyyy1DigVerif
         Clave_Acceso = MidStrg(FileNext, 1, 49)
         SerieFactura = MidStrg(Clave_Acceso, 25, 6)
         Factura_No = MidStrg(Clave_Acceso, 31, 9)
         Select Case MidStrg(Clave_Acceso, 9, 2)
           Case "01": TipoDoc = "FA"
           Case "04": TipoDoc = "NC"
           Case "07": TipoDoc = "RE"
           Case Else: TipoDoc = "XX"
         End Select
        'MsgBox RutaXMLAutorizado & FileNext & vbCrLf & TipoDoc & vbCrLf & SerieFactura & vbCrLf & Factura_No
         DatosXMLA = Leer_Archivo_Texto(RutaXMLAutorizado & FileNext)
         sSQL = "SELECT * " _
              & "FROM Trans_Documentos " _
              & "WHERE Item = '" & NumEmpresa & "' " _
              & "AND Periodo = '" & Periodo_Contable & "' " _
              & "AND Clave_Acceso = '" & Clave_Acceso & "' "
         Select_Adodc AdoAux, sSQL
         If AdoAux.Recordset.RecordCount <= 0 Then
            If Len(DatosXMLA) > 1 Then
               AdoAux.Recordset.AddNew
               AdoAux.Recordset.fields("Item") = NumEmpresa
               AdoAux.Recordset.fields("Periodo") = Periodo_Contable
               AdoAux.Recordset.fields("TD") = TipoDoc
               AdoAux.Recordset.fields("Serie") = SerieFactura
               AdoAux.Recordset.fields("Documento") = Factura_No
               AdoAux.Recordset.fields("Clave_Acceso") = Clave_Acceso
               AdoAux.Recordset.fields("Documento_Autorizado") = DatosXMLA
               AdoAux.Recordset.Update
            End If
         End If
   Next IdxFile
   sSQL = "UPDATE Trans_Documentos " _
        & "SET TD = 'FA', Serie = MidStrg(Clave_Acceso, 25, 6), Documento = MidStrg(Clave_Acceso, 31, 9) " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND Periodo = '" & Periodo_Contable & "' " _
        & "AND LEN(TD) <= 1 " _
        & "AND MidStrg(Clave_Acceso, 9, 2) = '01' "
   Ejecutar_SQL_SP sSQL
   
   sSQL = "UPDATE Trans_Documentos " _
        & "SET TD = 'NC', Serie = MidStrg(Clave_Acceso, 25, 6), Documento = MidStrg(Clave_Acceso, 31, 9) " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND Periodo = '" & Periodo_Contable & "' " _
        & "AND LEN(TD) <= 1 " _
        & "AND MidStrg(Clave_Acceso, 9, 2) = '04' "
   Ejecutar_SQL_SP sSQL
   
   sSQL = "UPDATE Trans_Documentos " _
        & "SET TD = 'RE', Serie = MidStrg(Clave_Acceso, 25, 6), Documento = MidStrg(Clave_Acceso, 31, 9) " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND Periodo = '" & Periodo_Contable & "' " _
        & "AND LEN(TD) <= 1 " _
        & "AND MidStrg(Clave_Acceso, 9, 2) = '07' "
   Ejecutar_SQL_SP sSQL
   Progreso_Final
End Sub

Public Sub Poner_Avreviatura_Accesos()
     RatonReloj
     sSQL = "SELECT Codigo, Nombre_Completo, Cod_Ejec, ID " _
          & "FROM Accesos " _
          & "WHERE Cod_Ejec = '.' "
     Select_Adodc AdoAux, sSQL
     With AdoAux.Recordset
      If .RecordCount > 0 Then
          Do While Not .EOF
            .fields("Cod_Ejec") = Abreviatura_Texto(.fields("Nombre_Completo"))
            .Update
            .MoveNext
          Loop
      End If
     End With
End Sub
