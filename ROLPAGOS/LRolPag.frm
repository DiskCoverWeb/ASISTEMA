VERSION 5.00
Object = "{05BFD3F1-6319-4F30-B752-C7A22889BCC4}#1.0#0"; "AcroPDF.dll"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.5#0"; "comctl32.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Begin VB.Form LRolPagos 
   Caption         =   "ROL DE PAGOS"
   ClientHeight    =   10035
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14775
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10035
   ScaleWidth      =   14775
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14775
      _ExtentX        =   26061
      _ExtentY        =   1164
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   18
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Salir"
            Object.ToolTipText     =   "Salir del Rol"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "CxC"
            Object.ToolTipText     =   "Ingresa las CxC"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "CxP"
            Object.ToolTipText     =   "Ingresa las CxP"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Quincena"
            Object.ToolTipText     =   "Procesa la Quincena del mes"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Nomina"
            Object.ToolTipText     =   "Procesa la nomina del Mes"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Rol_Individual"
            Object.ToolTipText     =   "Procesa Rol Individual"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Rol_Colectivo"
            Object.ToolTipText     =   "Procesa Rol Colectivo"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Grabar"
            Object.ToolTipText     =   "Grabar el Rol de Pagos a Contabilidad"
            Object.Tag             =   ""
            ImageIndex      =   10
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Limpiar"
            Object.ToolTipText     =   "Encerar los datos para empezar otromaes"
            Object.Tag             =   ""
            ImageIndex      =   11
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Excel"
            Object.ToolTipText     =   "Cuadre del Rol Vs Banco"
            Object.Tag             =   ""
            ImageIndex      =   9
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Email"
            Object.ToolTipText     =   "Envia por mails el rol actual"
            Object.Tag             =   ""
            ImageIndex      =   18
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Emails"
            Object.ToolTipText     =   "Envia por mails los roles"
            Object.Tag             =   ""
            ImageIndex      =   17
         EndProperty
         BeginProperty Button13 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button14 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Primero"
            Object.ToolTipText     =   "Ir al Inicio"
            Object.Tag             =   ""
            ImageIndex      =   12
         EndProperty
         BeginProperty Button15 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Antes"
            Object.ToolTipText     =   "Ir uno antes"
            Object.Tag             =   ""
            ImageIndex      =   13
         EndProperty
         BeginProperty Button16 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Despues"
            Object.ToolTipText     =   "Ir uno Despues"
            Object.Tag             =   ""
            ImageIndex      =   14
         EndProperty
         BeginProperty Button17 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Ultimo"
            Object.ToolTipText     =   "Ir al Final"
            Object.Tag             =   ""
            ImageIndex      =   15
         EndProperty
         BeginProperty Button18 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Buscar"
            Object.ToolTipText     =   "Buscar un Rol por Cédula"
            Object.Tag             =   ""
            ImageIndex      =   16
         EndProperty
      EndProperty
      Begin VB.Frame Frame2 
         Caption         =   "Mes a procesar"
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
         Height          =   645
         Left            =   10185
         TabIndex        =   1
         Top             =   0
         Width           =   7680
         Begin VB.ComboBox CMes 
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
            Left            =   2625
            TabIndex        =   5
            Text            =   "Todos"
            Top             =   210
            Width           =   1485
         End
         Begin VB.ComboBox CAnio 
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
            Left            =   735
            TabIndex        =   3
            Text            =   "2000"
            Top             =   210
            Width           =   1065
         End
         Begin VB.ComboBox CmbGrupos 
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
            Left            =   5250
            TabIndex        =   7
            Text            =   "Grupo"
            Top             =   210
            Width           =   2325
         End
         Begin VB.Label Label2 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Mes:"
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
            TabIndex        =   4
            Top             =   210
            Width           =   750
         End
         Begin VB.Label Label8 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Año:"
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
            Top             =   210
            Width           =   645
         End
         Begin VB.Label Label4 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Por Grupo"
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
            TabIndex        =   6
            Top             =   210
            Width           =   1065
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Ordenar por "
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
      Height          =   855
      Left            =   105
      TabIndex        =   8
      Top             =   630
      Width           =   1485
      Begin VB.OptionButton OpcEmpleado 
         Caption         =   "Empleado"
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
         TabIndex        =   10
         Top             =   525
         Width           =   1170
      End
      Begin VB.OptionButton OpcGrupo 
         Caption         =   "Grupo"
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
         TabIndex        =   9
         Top             =   210
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin MSDataListLib.DataCombo DCCxP 
      Bindings        =   "LRolPag.frx":0000
      DataSource      =   "AdoBanco"
      Height          =   315
      Left            =   7140
      TabIndex        =   16
      Top             =   1155
      Width           =   4845
      _ExtentX        =   8546
      _ExtentY        =   556
      _Version        =   393216
      Text            =   "CxP"
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
   Begin VB.CheckBox CheqCxP 
      Caption         =   "Generar Nomina sin alcance de efectivo (CxP Empleados)"
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
      Left            =   1680
      TabIndex        =   15
      ToolTipText     =   "Para usar esta opción debe crear en Contabilidad una Cuenta por Pagar sin módulo"
      Top             =   1155
      Width           =   5370
   End
   Begin VB.CheckBox CheqHoras 
      Caption         =   "Con Horas Trabajadas"
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
      TabIndex        =   13
      Top             =   735
      Value           =   1  'Checked
      Width           =   2325
   End
   Begin VB.CheckBox CheqCD 
      Caption         =   "Generar Rol con comprobante de Egreso"
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
      Left            =   9555
      TabIndex        =   14
      Top             =   735
      Value           =   1  'Checked
      Width           =   3795
   End
   Begin VB.CommandButton Command3 
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
      Left            =   12600
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   1155
      Width           =   330
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   9150
      Left            =   105
      TabIndex        =   19
      Top             =   1575
      Width           =   20175
      _ExtentX        =   35586
      _ExtentY        =   16140
      _Version        =   393216
      Tabs            =   6
      Tab             =   1
      TabsPerRow      =   6
      TabHeight       =   520
      BackColor       =   -2147483637
      TabCaption(0)   =   "ROL INDIVIDUAL"
      TabPicture(0)   =   "LRolPag.frx":0017
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "APDFRol"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "ROL DE PAGOS"
      TabPicture(1)   =   "LRolPag.frx":0033
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "AdoNomina"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "DGTotNomina"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "DGNomina"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "CxC/CxP Empleados"
      TabPicture(2)   =   "LRolPag.frx":004F
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "DGSubCtas"
      Tab(2).Control(1)=   "DGNomina1"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "OTROS INGRESOS/EGRESOS"
      TabPicture(3)   =   "LRolPag.frx":006B
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "DGI_E_Empleado"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "CONTABILIDAD"
      TabPicture(4)   =   "LRolPag.frx":0087
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "DGAsiento(0)"
      Tab(4).Control(1)=   "Label1"
      Tab(4).Control(2)=   "Label19"
      Tab(4).Control(3)=   "LabelDebe"
      Tab(4).Control(4)=   "LabelHaber"
      Tab(4).Control(5)=   "LabelDiferencia"
      Tab(4).Control(6)=   "LblConcepto(0)"
      Tab(4).ControlCount=   7
      TabCaption(5)   =   "PROVISIONES"
      TabPicture(5)   =   "LRolPag.frx":00A3
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "DGAsiento(2)"
      Tab(5).Control(1)=   "DGAsiento(1)"
      Tab(5).Control(2)=   "LblConcepto(2)"
      Tab(5).Control(3)=   "LblConcepto(1)"
      Tab(5).ControlCount=   4
      Begin AcroPDFLibCtl.AcroPDF APDFRol 
         Height          =   3690
         Left            =   -74895
         TabIndex        =   24
         Top             =   420
         Width           =   7680
         _cx             =   5080
         _cy             =   5080
      End
      Begin MSDataGridLib.DataGrid DGNomina 
         Bindings        =   "LRolPag.frx":00BF
         Height          =   3900
         Left            =   105
         TabIndex        =   20
         ToolTipText     =   $"LRolPag.frx":00D7
         Top             =   420
         Width           =   10305
         _ExtentX        =   18177
         _ExtentY        =   6879
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
      Begin MSDataGridLib.DataGrid DGTotNomina 
         Bindings        =   "LRolPag.frx":0172
         Height          =   1065
         Left            =   105
         TabIndex        =   22
         ToolTipText     =   "<Ctrl + F9>: Comisiones y el I.E.S.S."
         Top             =   4410
         Width           =   10305
         _ExtentX        =   18177
         _ExtentY        =   1879
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
      Begin MSAdodcLib.Adodc AdoNomina 
         Height          =   330
         Left            =   105
         Top             =   5460
         Width           =   10305
         _ExtentX        =   18177
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   1
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
         Caption         =   "Nomina"
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
      Begin MSDataGridLib.DataGrid DGNomina1 
         Bindings        =   "LRolPag.frx":018D
         Height          =   2745
         Left            =   -74895
         TabIndex        =   21
         ToolTipText     =   "<Ctrl + F9>: Comisiones y el I.E.S.S."
         Top             =   420
         Width           =   10305
         _ExtentX        =   18177
         _ExtentY        =   4842
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
      Begin MSDataGridLib.DataGrid DGSubCtas 
         Bindings        =   "LRolPag.frx":01A6
         Height          =   2640
         Left            =   -74895
         TabIndex        =   23
         ToolTipText     =   "<Ctrl + F9>: Comisiones y el I.E.S.S."
         Top             =   3150
         Width           =   10305
         _ExtentX        =   18177
         _ExtentY        =   4657
         _Version        =   393216
         AllowUpdate     =   -1  'True
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
         Bindings        =   "LRolPag.frx":01C1
         Height          =   2850
         Index           =   2
         Left            =   -74895
         TabIndex        =   25
         Top             =   3885
         Width           =   10305
         _ExtentX        =   18177
         _ExtentY        =   5027
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
      Begin MSDataGridLib.DataGrid DGAsiento 
         Bindings        =   "LRolPag.frx":01DB
         Height          =   1905
         Index           =   1
         Left            =   -74895
         TabIndex        =   26
         Top             =   735
         Width           =   10305
         _ExtentX        =   18177
         _ExtentY        =   3360
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
      Begin MSDataGridLib.DataGrid DGAsiento 
         Bindings        =   "LRolPag.frx":01F5
         Height          =   4635
         Index           =   0
         Left            =   -74895
         TabIndex        =   29
         Top             =   735
         Width           =   10305
         _ExtentX        =   18177
         _ExtentY        =   8176
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
      Begin MSDataGridLib.DataGrid DGI_E_Empleado 
         Bindings        =   "LRolPag.frx":020E
         Height          =   3900
         Left            =   -74895
         TabIndex        =   36
         Top             =   420
         Width           =   10305
         _ExtentX        =   18177
         _ExtentY        =   6879
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
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Diferencia"
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
         TabIndex        =   35
         Top             =   5565
         Width           =   1170
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
         Left            =   -71955
         TabIndex        =   34
         Top             =   5565
         Width           =   1065
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
         ForeColor       =   &H000000FF&
         Height          =   330
         Left            =   -70905
         TabIndex        =   33
         Top             =   5565
         Width           =   1800
      End
      Begin VB.Label LabelHaber 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H000000FF&
         Height          =   330
         Left            =   -69120
         TabIndex        =   32
         Top             =   5565
         Width           =   1800
      End
      Begin VB.Label LabelDiferencia 
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
         ForeColor       =   &H000000FF&
         Height          =   330
         Left            =   -73740
         TabIndex        =   31
         Top             =   5565
         Width           =   1695
      End
      Begin VB.Label LblConcepto 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "."
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
         Left            =   -74895
         TabIndex        =   30
         Top             =   420
         Width           =   10305
      End
      Begin VB.Label LblConcepto 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "."
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
         Left            =   -74895
         TabIndex        =   28
         Top             =   3570
         Width           =   10305
      End
      Begin VB.Label LblConcepto 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "."
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
         Left            =   -74895
         TabIndex        =   27
         Top             =   420
         Width           =   10305
      End
   End
   Begin VB.TextBox TxtCheque 
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
      Left            =   5775
      MaxLength       =   8
      TabIndex        =   12
      Text            =   "0"
      Top             =   735
      Width           =   1275
   End
   Begin MSAdodcLib.Adodc AdoAsiento 
      Height          =   330
      Left            =   1260
      Top             =   5565
      Visible         =   0   'False
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
   Begin MSAdodcLib.Adodc AdoSubCta 
      Height          =   330
      Left            =   1260
      Top             =   2415
      Visible         =   0   'False
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
   Begin MSAdodcLib.Adodc AdoSubCta1 
      Height          =   330
      Left            =   1260
      Top             =   2730
      Visible         =   0   'False
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
      Caption         =   "SubCta1"
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
      Left            =   1260
      Top             =   3045
      Visible         =   0   'False
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
      CommandType     =   1
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
   Begin MSAdodcLib.Adodc AdoBanco 
      Height          =   330
      Left            =   1260
      Top             =   3360
      Visible         =   0   'False
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
      CommandType     =   1
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
   Begin MSAdodcLib.Adodc AdoCaja 
      Height          =   330
      Left            =   1260
      Top             =   3675
      Visible         =   0   'False
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
      CommandType     =   1
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
   Begin MSAdodcLib.Adodc AdoAsientoB 
      Height          =   330
      Left            =   1260
      Top             =   3990
      Visible         =   0   'False
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
      CommandType     =   1
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
      Caption         =   "AsientoB"
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
   Begin MSAdodcLib.Adodc AdoNomina1 
      Height          =   330
      Left            =   1260
      Top             =   4305
      Visible         =   0   'False
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
      CommandType     =   1
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
      Caption         =   "Nomina1"
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
   Begin MSAdodcLib.Adodc AdoTotNomina 
      Height          =   330
      Left            =   1260
      Top             =   4620
      Visible         =   0   'False
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
      CommandType     =   1
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
      Caption         =   "TotNomina"
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
   Begin MSAdodcLib.Adodc AdoCertificado 
      Height          =   330
      Left            =   1260
      Top             =   4935
      Visible         =   0   'False
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
      Caption         =   "Certificado"
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
   Begin MSAdodcLib.Adodc AdoAsientoSC 
      Height          =   330
      Left            =   1260
      Top             =   5250
      Visible         =   0   'False
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
      CommandType     =   1
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
      Caption         =   "AsientoSC"
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
   Begin MSAdodcLib.Adodc AdoAsiento1 
      Height          =   330
      Left            =   1260
      Top             =   5880
      Visible         =   0   'False
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
      Caption         =   "Asiento1"
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
   Begin MSAdodcLib.Adodc AdoAsiento2 
      Height          =   330
      Left            =   1260
      Top             =   6195
      Visible         =   0   'False
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
      Caption         =   "Asiento2"
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
      Left            =   1260
      Top             =   6510
      Visible         =   0   'False
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
   Begin MSAdodcLib.Adodc AdoAsientoRol 
      Height          =   330
      Left            =   1260
      Top             =   6825
      Visible         =   0   'False
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
      Caption         =   "AsientoRol"
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
   Begin MSAdodcLib.Adodc AdoNominaProv 
      Height          =   330
      Left            =   1260
      Top             =   2100
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
      CommandType     =   1
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
      Caption         =   "NominaProv"
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
   Begin MSAdodcLib.Adodc AdoNovedades 
      Height          =   330
      Left            =   1260
      Top             =   7140
      Visible         =   0   'False
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
      Caption         =   "Novedades"
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
   Begin MSAdodcLib.Adodc AdoImpRenta 
      Height          =   330
      Left            =   3570
      Top             =   2100
      Visible         =   0   'False
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
      Caption         =   "ImpRenta"
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
   Begin MSAdodcLib.Adodc AdoCtaCat 
      Height          =   330
      Left            =   3570
      Top             =   2520
      Visible         =   0   'False
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
      CommandType     =   1
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
      Caption         =   "CtaCat"
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
   Begin MSAdodcLib.Adodc AdoI_E_Empleado 
      Height          =   330
      Left            =   3570
      Top             =   2940
      Visible         =   0   'False
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
      Caption         =   "I_E_Empleado"
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
      Left            =   11760
      Top             =   1575
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   18
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LRolPag.frx":022C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LRolPag.frx":0546
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LRolPag.frx":0860
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LRolPag.frx":0B7A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LRolPag.frx":0E94
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LRolPag.frx":11AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LRolPag.frx":14C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LRolPag.frx":17E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LRolPag.frx":1AFC
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LRolPag.frx":4C06
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LRolPag.frx":4F20
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LRolPag.frx":523A
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LRolPag.frx":5464
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LRolPag.frx":56A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LRolPag.frx":58E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LRolPag.frx":5B0A
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LRolPag.frx":5D34
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LRolPag.frx":604E
            Key             =   ""
         EndProperty
      EndProperty
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
      Height          =   330
      Left            =   12075
      TabIndex        =   17
      Top             =   1155
      Width           =   435
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Número de Cheque con el que empieza el Rol"
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
      Left            =   1680
      TabIndex        =   11
      Top             =   735
      Width           =   4110
   End
End
Attribute VB_Name = "LRolPagos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim oDocument As Object
Dim DocumentoXML As MSXML2.DOMDocument30

Dim CantCtas As Long
Dim MyTime As Single
Dim Lista_Emails As String
Dim PrimerDia As String
Dim UltimoDia As String

Dim TRol_Pago As Tipo_Rol_Pago_Individual
Dim CtasRol() As CtasAsiento
Dim CtasPro() As CtasAsiento
Dim CtasPat() As CtasAsiento

Public Sub Imprimir_Pagina(Optional Impresora As Boolean)
Dim AnchoPict As Single
Dim AltoPict As Single
Dim NombFilePict As String
On Error GoTo Errorhandler
'''AcrovPDF.LoadFile ""
'''AcrovPDF.setZoom 120
'''AcrovPDF.Visible = True

Mensajes = "Seguro de Imprimir en:"
Titulo = "IMPRESION"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
   RatonReloj
  'Generamos el documento
   NombFilePict = "Rol Pagos " & CAnio & "-" & Format$(NumMeses, "00") & " R-" & RUC & " " & CodigoUsuario
   tPrint.TipoImpresion = cPrint.cPrinter
   tPrint.NombreArchivo = NombFilePict
   tPrint.TituloArchivo = "Rol de Pagos " & CAnio & "-" & Format(NumMeses, "00") & " " & RUC
   tPrint.TipoLetra = TipoHelvetica
   tPrint.OrientacionPagina = 1
   tPrint.PaginaA4 = True
   tPrint.EsCampoCorto = False
   tPrint.VerDocumento = False
   
   Set cPrint = New cImpresion
   cPrint.iniciaImpresion
   
'''   InicioX = 0: InicioY = 0
'''   Escala_Centimetro Orientacion_Pagina, TipoTimes, 8
'''   Pagina = 1
'''   If Impresora Then
'''      If Medio_Rol Then
'''         Generar_Rol_Printer_Medio AdoNomina.Recordset.Fields("Codigo"), 0.01, 0.01
'''         Generar_Rol_Printer_Medio AdoNomina.Recordset.Fields("Codigo"), 10.5, 0.01
'''      Else
'''         Generar_Rol_Printer AdoNomina.Recordset.Fields("Codigo"), 0.01, 0.01
'''      End If
'''   Else
'''      AnchoPict = Redondear(Printer.ScaleWidth, 5)
'''      AltoPict = Redondear(Printer.ScaleHeight, 5)
'''      Printer.PaintPicture PictRol.Image, InicioX, InicioY, AnchoPict, AltoPict
'''   End If
   MensajeEncabData = ""
   
  'fin del documento
'''   cPrint.finalizaImpresion
'''   AcrovPDF.LoadFile RutaSysBases & "\TEMP\" & NombFilePict & ".pdf"
'''   AcrovPDF.setZoom 120
'''   AcrovPDF.Visible = True

   RatonNormal
   Exit Sub
Errorhandler:
             RatonNormal
             ErrorDeImpresion
             Exit Sub
Else
   RatonNormal
End If
End Sub

Public Sub InsertarCertificado(CtaProc As String, Valor As Currency, TipoDeCta As String)
  SetAdoAddNew "Asiento_SC"
  SetAdoFields "Codigo", CodigoCliente
  SetAdoFields "Beneficiario", NombreCliente
  SetAdoFields "Cta", CtaProc
  SetAdoFields "Valor", Redondear(Valor, 2)
  SetAdoFields "Fecha_V", FechaFinal
  SetAdoFields "TC", TipoDeCta
  SetAdoFields "Factura", Factura_No
  SetAdoFields "DH", "2"
  SetAdoFields "Valor_ME", 0
  SetAdoFields "TM", "1"
  SetAdoFields "Item", NumEmpresa
  SetAdoFields "T_No", Trans_No
  SetAdoFields "CodigoU", CodigoUsuario
  If Valor > 0 Then SetAdoUpdate
End Sub

Public Sub Llenar_Rol_Pagos_Individual(CodigoRol As String, Optional General_PDF As Boolean)
Dim AdoAuxRolDB As ADODB.Recordset
Dim ContLineas As Integer
Dim Es_Vacaciones As Boolean
Dim Aporte_Patronal As Single
Dim NombFilePict As String
Dim NombFilehtml As String

    If Len(CodigoRol) > 1 Then
      'Datos del Encabezadodel Rol Individual
       No_Personal = Ninguno
       FechaTexto = Ninguno
       CICliente = Ninguno
       NomCtaSup = Ninguno
       NumCheque = Ninguno
       TextoBanco = Ninguno
       CodigoB = "OTROS"
       
       Grupo_No = CmbGrupos
       sSQL = "SELECT C.Cliente, C.CI_RUC, C.Direccion, C.Telefono, C.Actividad, C.Email, C.Email2, " _
            & "CR.Fecha, CR.No_Personal, CR.FechaVI, CR.FechaVF, CR.Mes, CR.Cta_Transferencia, CR.Codigo_Banco " _
            & "FROM Clientes As C, Catalogo_Rol_Pagos As CR " _
            & "WHERE CR.Item = '" & NumEmpresa & "' " _
            & "AND CR.Periodo = '" & Periodo_Contable & "' " _
            & "AND CR.Salario > 0 " _
            & "AND CR.Fecha <= #" & BuscarFecha(FechaFinal) & "# " _
            & "AND CR.Codigo = '" & CodigoRol & "' " _
            & "AND CR.Codigo = C.Codigo "
       Select_AdoDB AdoAuxRolDB, sSQL
       With AdoAuxRolDB
        If .RecordCount > 0 Then
            NombreCliente = Replace(.fields("Cliente"), ".", "")
            No_Personal = .fields("No_Personal")
            FechaTexto = .fields("Fecha")
            CICliente = .fields("CI_RUC")
            FA.CxC_Clientes = .fields("Actividad")
            FechaInicial = .fields("FechaVI")
            FechaFinal = .fields("FechaVF")
            NoMeses = .fields("Mes")
            NomCtaSup = .fields("Cta_Transferencia")
            'Cta_IESS = .fields("Cta_IESS_Personal")
            TextoBanco = .fields("Codigo_Banco")
           'Enviamos lista de mails
            Lista_Emails = ""
            If Len(.fields("Email")) > 1 Then Lista_Emails = Lista_Emails & TrimStrg(.fields("Email")) & ";"
            If .fields("Email") <> .fields("Email2") And Len(.fields("Email2")) > 1 Then
                Lista_Emails = Lista_Emails & TrimStrg(.fields("Email2")) & ";"
            End If
        End If
       End With
       AdoAuxRolDB.Close
       
       sSQL = "SELECT Descripcion " _
            & "FROM Tabla_Referenciales_SRI " _
            & "WHERE Codigo = '" & TextoBanco & "' " _
            & "AND Tipo_Referencia = 'BANCOS Y COOP' "
       Select_AdoDB AdoAuxRolDB, sSQL
       If AdoAuxRolDB.RecordCount > 0 Then TextoBanco = ULCase(AdoAuxRolDB.fields("Descripcion"))
       AdoAuxRolDB.Close
       
       Fecha_Del_AT CMes, CAnio
       
      'Presentamos el rol individual del empleado
       SQL2 = "SELECT " & Full_Fields("Trans_Rol_de_Pagos") & " " _
            & "FROM Trans_Rol_de_Pagos " _
            & "WHERE Fecha_D >= #" & FechaIni & "# " _
            & "AND Fecha_H <= #" & FechaFin & "# " _
            & "AND Item = '" & NumEmpresa & "' " _
            & "AND Periodo = '" & Periodo_Contable & "' " _
            & "AND Codigo = '" & CodigoRol & "' "
       If CmbGrupos <> "TODOS" Then SQL2 = SQL2 & "AND Grupo_Rol = '" & CmbGrupos & "' "
       SQL2 = SQL2 & "ORDER BY Grupo_Rol,Codigo,Tipo_Rubro,ID,Ingresos desc,Egresos,Detalle "
       Select_Adodc AdoAsientoRol, SQL2
      'MsgBox SQL2 & vbCrLf & vbCrLf & General_PDF
      'Generamos el documento
       If Not (General_PDF) Then
         'SetNombrePRN = ""
          SetNombrePRN = Impresota_PDF
          NombFilePict = "Rol_Pagos de " & NombreCliente & " " & CAnio & "-" & Format$(NumMeses, "00") & " " & CodigoUsuario
          NombFilePict = Replace(NombFilePict, " ", "_")
          NombFilehtml = NombFilePict & ".html"
          tPrint.TipoImpresion = Es_PDF
          tPrint.NombreArchivo = NombFilePict
          tPrint.TituloArchivo = "Rol de Pagos de " & NombreCliente & " " & CAnio & "-" & Format(NumMeses, "00")
          tPrint.TipoLetra = TipoHelvetica
          tPrint.OrientacionPagina = 1
          tPrint.PaginaA4 = True
          tPrint.EsCampoCorto = False
          tPrint.VerDocumento = False
          Set cPrint = New cImpresion
          cPrint.iniciaImpresion
          If Medio_Rol Then
            'Si es medio rol izquierda y derecha
             Generar_Rol_Medio CodigoRol, 1.3, 1
             Generar_Rol_Medio CodigoRol, 11.3, 1
          Else
            'Si es rol completo arriba y abajo
             Generar_Rol CodigoRol, 1.3, 1
          End If
          
         'fin del documento
          cPrint.finalizaImpresion
          Set oDocument = Nothing
          'WebBPDF.navigate RutaDocumentoPDF
          'Generar_Rol_html CodigoRol, NombFilehtml
          'MsgBox "...." & vbCrLf & NombFilePict
          APDFRol.Object.src = RutaSysBases & "\TEMP\" & NombFilePict & ".pdf"
          'WebBPDF.navigate RutaSysBases & "\TEMP\" & NombFilehtml
          'Presentar_PDF RPPDF, RutaDocumentoPDF, 125
       End If
    End If
End Sub

Private Sub CheqCxP_Click()
  Cta = "0"
  If CheqCxP.value = 1 Then DCCxP.Visible = True Else DCCxP.Visible = False
End Sub

Private Sub CmbGrupos_GotFocus()
  LRolPagos.Caption = "ROL DE PAGOS MES DE " & UCase(MesesLetras(Month(FechaFinal)))
  FechaInicial = PrimerDiaMes(FechaFinal)
  PrimerDia = BuscarFecha(PrimerDiaMes(FechaFinal))
  UltimoDia = BuscarFecha(UltimoDiaMes(FechaFinal))
End Sub

Private Sub CmbGrupos_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub CmbGrupos_LostFocus()
'''  Listar_Cuentas_Rol "R"
  '''Listar_Empleados
End Sub

Private Sub CMes_LostFocus()
  Fecha_Del_AT CMes, CAnio
  Datos_IESS FechaFinal
  
  sSQL = "SELECT C.Cliente, CRR.I_E, CRR.Detalle, CRR.Cta, CRR.Valor, CRR.Calc_IESS, CRR.Cod_Rol_Pago " _
       & "FROM Catalogo_Rol_Rubros As CRR, Clientes As C " _
       & "WHERE CRR.Item = '" & NumEmpresa & "' " _
       & "AND CRR.Periodo = '" & Periodo_Contable & "' " _
       & "AND CRR.Mes = " & Month(FechaInicial) & " " _
       & "AND CRR.Codigo = C.Codigo " _
       & "ORDER BY C.Cliente, CRR.Codigo, CRR.I_E, CRR.Cta "
  Select_Adodc_Grid DGI_E_Empleado, AdoI_E_Empleado, sSQL

  sSQL = "SELECT Grupo_Rol " _
       & "FROM Catalogo_Rol_Pagos " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND Salario > 0 " _
       & "AND Fecha <= #" & BuscarFecha(FechaFinal) & "# " _
       & "GROUP BY Grupo_Rol " _
       & "ORDER BY Grupo_Rol "
  Select_Adodc AdoAux, sSQL
  CmbGrupos.Clear
  CmbGrupos.AddItem "TODOS"
  With AdoAux.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          CmbGrupos.AddItem .fields("Grupo_Rol")
         .MoveNext
       Loop
   End If
  End With
  CmbGrupos.Text = "TODOS"
  'Listar_Empleados
  Listar_CxCxP_SubMod
  CmbGrupos.SetFocus
End Sub

Public Sub Procesar_Rol_De_Pagos_Mes()
Dim Rol_I As Long
Dim Rol_M As Long
Dim Rol_F As Long
Dim CxP_RolPagos As String
Dim Fecha_Rol_Mes As String
Dim OrdenAlfabetico As Boolean
  
  RatonReloj
  Progreso_Barra.Incremento = 0
  Progreso_Barra.Valor_Maximo = 100
  Progreso_Barra.Mensaje_Box = "Encerando Asientos"
  Progreso_Esperar
  
  SQL1 = "DELETE FROM Tabla_Temporal " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' "
  Ejecutar_SQL_SP SQL1
  
  If CheqCxP.value <> 0 Then CxP_RolPagos = SinEspaciosIzq(DCCxP) Else CxP_RolPagos = "0"
  
 'Procedemos a encerar el rol a procesar o consultar
  TextoImprimio = ""
  Inicializar_Cero_Asientos
  Si_No = Leer_Campo_Empresa("Rol_2_Pagina")
  Medio_Rol = Leer_Campo_Empresa("Medio_Rol")
  If OpcGrupo.value Then OrdenAlfabetico = False Else OrdenAlfabetico = True
  Fecha_Del_AT CMes, CAnio

 'Procedemos a procesar el Rol pedido del mes o quincena
  Opcion = 1
  SQL2 = "SELECT Item, Periodo, Codigo " _
       & "FROM Trans_Rol_de_Pagos " _
       & "WHERE Fecha_D >= #" & FechaIni & "# " _
       & "AND Fecha_H <= #" & FechaFin & "# " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND TP <> '" & Ninguno & "' " _
       & "AND Numero <> 0 "
  Select_Adodc AdoAux, SQL2
 'MsgBox AdoAux.Recordset.RecordCount & vbCrLf & vbCrLf & SQL2
  If AdoAux.Recordset.RecordCount > 0 Then
     MsgBox "Este Rol ya fue Procesado. Se procedera solo a presentar el Rol."
  Else
    'Procesamos el Rol de Pagos del Mes
     Procesar_Rol_Pagos_del_Mes_SP FechaIni, FechaFin, CmbGrupos.Text, CxP_RolPagos, CLng(TxtCheque)
    '---------------------------------------------------------------------------------------
  End If
  Trans_No = 100
  DGAsiento(0).Visible = False
  DGAsiento(1).Visible = False
  DGAsiento(2).Visible = False
  
  Progreso_Barra.Mensaje_Box = "Procesar Asientos"
  Progreso_Esperar
  Procesar_Rol_Pagos_Asientos_SP FechaIni, FechaFin
  
  Progreso_Barra.Mensaje_Box = "LLenar Rol Pagos"
  Progreso_Esperar

  Reporte_Rol_Pagos_Colectivo_SP FechaIni, FechaFin, CmbGrupos, OrdenAlfabetico, SubSQL, SQL2
  
  sSQL = "SELECT " & SubSQL & " " _
       & "FROM Reporte_Rol_Colectivo " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' "
  Select_Adodc_Grid DGNomina, AdoNomina, sSQL, 2, True
  
  sSQL = "SELECT " & SQL2 & " " _
       & "FROM Reporte_Rol_Colectivo " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "GROUP BY Grupo_Rol " _
       & "ORDER BY Grupo_Rol "
  Select_Adodc_Grid DGTotNomina, AdoTotNomina, sSQL, 2, True
  
 'Listar_Empleados
 
  LblConcepto(0).Caption = "Registro de Nómina correspondiente al mes de " & MesesLetras(Month(FechaFinal))
  LblConcepto(1).Caption = "Provision IESS Patronal correspondiente al mes de " & MesesLetras(Month(FechaFinal))
  LblConcepto(2).Caption = "Provision Decimo 3er., Decimo 4to., Vacaciones y de Fondos de Reserva correspondiente al mes de " & MesesLetras(Month(FechaFinal))
  If AdoNomina.Recordset.RecordCount > 0 Then
     AdoNomina.Recordset.MoveFirst
     Llenar_Rol_Pagos_Individual AdoNomina.Recordset.fields("Codigo")
  End If
    
  DGAsiento(0).Visible = True
  DGAsiento(1).Visible = True
  DGAsiento(2).Visible = True
    
  Listar_CxCxP_SubMod
  
  Inicializar_Cero_Asientos
  
  Progreso_Barra.Incremento = Progreso_Barra.Valor_Maximo
  Progreso_Barra.Mensaje_Box = "Fin del Proceso del Rol"
  Progreso_Final
  
  DGNomina1.Caption = "EMPLEADOS SIN ALCANCE DE REMUNERACION"
  sSQL = "SELECT C.Cliente As Empleado, TRP.Egresos as Neto_a_Recibir " _
       & "FROM Clientes as C,Trans_Rol_de_Pagos As TRP " _
       & "WHERE TRP.Item = '" & NumEmpresa & "' " _
       & "AND TRP.Periodo = '" & Periodo_Contable & "' " _
       & "AND TRP.Fecha_D = #" & FechaIni & "# " _
       & "AND TRP.Cod_Rol_Pago = 'Neto_Recibir' " _
       & "AND TRP.Egresos <= 0 " _
       & "AND TRP.Codigo = C.Codigo " _
       & "ORDER BY Cliente "
  Select_Adodc_Grid DGNomina1, AdoNomina1, sSQL
  LRolPagos.Caption = "NOMINA O ROL DE PAGOS DE " & UCase(MesesLetras(Month(FechaFinal))) & " - " & Format(Time - MyTime, "mm:ss")
  MsgBox "Fin del Proceso, Revise los resultados."
  FInfoError.Show
End Sub

'CxC de Quincena
Public Sub Procesar_Quincena()
Dim Rol_I As Long
Dim Rol_M As Long
Dim Rol_F As Long
' Procesamos los Ingresos/Egresos de Rol de Pagos
  RatonReloj
  Ln_No = 1
  Opcion = 2
  DetalleComp = Ninguno
  Fecha_Vence = FechaFinal
  Trans_No = 100
  'TotalAbonos1 = Val(CCur(TxtMonto))
  Comp_No = Val(TxtCheque)
 'Asientos y Submodulos de CxC por quincena
  SQL1 = "DELETE " _
       & "FROM Asiento " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND T_No = " & Trans_No & " "
  Ejecutar_SQL_SP SQL1
  SQL2 = "DELETE * " _
       & "FROM Asiento_SC " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND T_No = " & Trans_No & " " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND TC = 'C' "
  Ejecutar_SQL_SP SQL2
  SQL2 = "DELETE * " _
       & "FROM Trans_Rol_Pagos " _
       & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' "
  If CmbGrupos.Text <> "TODOS" Then SQL2 = SQL2 & "AND Grupo_Rol = '" & CmbGrupos & "' "
  Ejecutar_SQL_SP SQL2

 'Grilla de Asientos
  SQL2 = "SELECT * " _
       & "FROM Asiento_SC " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND T_No = " & Trans_No & " " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND TC = 'C' " _
       & "ORDER BY Codigo,TC,Cta "
  Select_Adodc AdoAsientoSC, SQL2
  SQL2 = "SELECT * " _
       & "FROM Asiento " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND T_No = " & Trans_No & " " _
       & "AND CodigoU = '" & CodigoUsuario & "' "
  Select_Adodc_Grid DGAsiento(0), AdoAsiento, SQL2
  Total = 0
  Contador = 0
  With AdoClientes.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       Do While Not .EOF
          Contador = Contador + 1
          CodigoCliente = .fields("Codigo")
          NombreCliente = .fields("Cliente")
          Grupo_No = .fields("Grupo_Rol")
          Cta = .fields("Cta_Quincena")
          Abono = Redondear(.fields("Salario") / 2, 2)
          If Len(Cta) <= 1 Then Abono = 0
          NoCheque = Ninguno
          DetalleComp = Format(Contador, "00") & ".- " & NombreCliente
          If Len(.fields("Cta_Transferencia")) > 1 Then
             'Cta1 = SinEspaciosIzq(DLBanco1.Text)
             NoCheque = "TRANSFERENCIA"
          ElseIf Abono > TotalAbonos1 And Comp_No > 0 Then
             'Cta1 = SinEspaciosIzq(DLBanco.Text)
             NoCheque = Format(Comp_No, "00000000")
             Comp_No = Comp_No + 1
          Else
             'Cta1 = SinEspaciosIzq(DLCtas.Text)
          End If
          If Abono > 0 Then
             InsertarAsientos AdoAsiento, Cta1, 0, 0, Abono
            'Insertamos el submodulo de CxC
             SetAdoAddNew "Asiento_SC"
             SetAdoFields "Codigo", CodigoCliente
             SetAdoFields "Beneficiario", NombreCliente
             SetAdoFields "Cta", Cta
             SetAdoFields "Valor", Abono
             SetAdoFields "FECHA_V", FechaFinal
             SetAdoFields "TC", "C"
             SetAdoFields "Factura", 0
             SetAdoFields "DH", "1"
             SetAdoFields "Valor_ME", 0
             SetAdoFields "TM", "1"
             SetAdoFields "Item", NumEmpresa
             SetAdoFields "T_No", Trans_No
             SetAdoFields "SC_No", Ln_No
             SetAdoFields "CodigoU", CodigoUsuario
             SetAdoUpdate
             'InsValorCtaRol Cta, Abono
          End If
          
         'Rol Pago Quincena
          SetAdoAddNew "Trans_Rol_Pagos"
          SetAdoFields "SN", "1"
          SetAdoFields "T", Normal
          SetAdoFields "Codigo", CodigoCliente
          SetAdoFields "Fecha", FechaFinal
          SetAdoFields "Dias", Day(FechaFinal)
          SetAdoFields "InLiquido", Abono
          SetAdoFields "Neto_Recibir", Abono
          SetAdoFields "Cheque_No", NoCheque
          SetAdoFields "CodigoU", CodigoUsuario
          SetAdoFields "Item", NumEmpresa
          SetAdoFields "Quincena", True
          If CmbGrupos.Text <> "TODOS" Then SetAdoFields "Grupo_Rol", CmbGrupos
          SetAdoUpdate
         .MoveNext
       Loop
   End If
  End With
  SQL2 = "SELECT * " _
       & "FROM Asiento_SC " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND T_No = " & Trans_No & " " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND TC = 'C' " _
       & "ORDER BY TC,Cta,DH "
  Select_Adodc AdoAsientoSC, SQL2
  Ln_No = 0
  NoCheque = Ninguno
  DetalleComp = "Anticipo Empleados por quincena del mes de " & MesesLetras(Month(FechaFinal))
  For IE = 0 To CantCtas - 1
      If CtasRol(IE).Cta <> "0" Then
         Select Case MidStrg(CtasRol(IE).Cta, 1, 1)
           Case "1": InsertarAsientos AdoAsiento, CtasRol(IE).Cta, 0, CtasRol(IE).Valor, 0
           Case "2": InsertarAsientos AdoAsiento, CtasRol(IE).Cta, 0, 0, CtasRol(IE).Valor
         End Select
      End If
  Next IE
  Debe = 0
  Haber = 0
  SQL2 = "SELECT * " _
       & "FROM Asiento " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND T_No = " & Trans_No & " " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "ORDER BY A_No "
  Select_Adodc_Grid DGAsiento(0), AdoAsiento, SQL2
  With AdoAsiento.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          Debe = Debe + .fields("Haber")
          Haber = Haber + .fields("Haber")
         .MoveNext
       Loop
   End If
  End With
  LabelDebe.Caption = Format(Debe, "#,##0.00")
  LabelHaber.Caption = Format(Haber, "#,##0.00")
  LabelDiferencia.Caption = Format(Debe - Haber, "#,##0.00")
  LblConcepto(0).Caption = "Pago primera quincena del mes de " & MesesLetras(Month(FechaFinal))
  RatonNormal
  MsgBox "Quincena procesada, revizar antes de grabar"
End Sub

Public Sub Procesar_Limpiar()
  RatonReloj
  Trans_No = 100
  SQL2 = "DELETE * " _
       & "FROM Asiento_SC " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND T_No = " & Trans_No & " "
  Ejecutar_SQL_SP SQL2
  RatonNormal
End Sub

Public Sub Procesar_Rol_Individual()
Dim AuxPosLinea As Single
Dim AntPosLinea As Single
Dim AnchoPict As Single
Dim AltoPict As Single
Dim PosCopiaY As Single
Dim NumEmpleados As Integer
Dim NombFilePict As String

On Error GoTo Errorhandler
   Si_No = Leer_Campo_Empresa("Rol_2_Pagina")
   Medio_Rol = Leer_Campo_Empresa("Medio_Rol")
   Fecha_Del_AT CMes, CAnio
   Bandera = False
   
   Mensajes = "Seguro de Generar los Roles Individuales?"
   Titulo = "GENERACION ROLES INDIVIDUALES"
   If BoxMensaje Then
      RatonReloj
     'Generamos el documento
      NombFilePict = "Roles_de_Pagos_" & CAnio & "-" & Format(NumMeses, "00") & "_" & RUC
      tPrint.TipoImpresion = Es_PDF
      tPrint.NombreArchivo = NombFilePict
      tPrint.TituloArchivo = NombFilePict
      tPrint.TipoLetra = TipoHelvetica
      tPrint.OrientacionPagina = 1
      tPrint.PaginaA4 = True
      tPrint.EsCampoCorto = False
      tPrint.VerDocumento = False
      Set cPrint = New cImpresion
      cPrint.iniciaImpresion
      InicioX = 0.5: InicioY = 0
      PosCopiaY = Redondear(SetPapelLargo / 2, 2) ' Largo del papel
     'MsgBox SetPapelCopia & vbCrLf & SetPapelAncho & vbCrLf & SetPapelLargo & vbCrLf & PosCopiaY & vbCrLf & Printer.ScaleHeight
      Pagina = 1
      Contador = 1
      NumEmpleados = 0
      PosLinea = 0.1
      'IR = Val(InputBox("Empezar desde: ", "IMPRESION", "0"))
      IR = 0
      NombreBanco1 = UCase(MidStrg(NombreBanco1, Len(SinEspaciosIzq(NombreBanco1)) + 1, Len(NombreBanco1)))
      With AdoNomina.Recordset
       If .RecordCount > 0 Then
          .MoveFirst
           Progreso_Barra.Incremento = 0
           Progreso_Barra.Valor_Maximo = .RecordCount
           Do While Not .EOF
              'MsgBox Medio_Rol & vbCrLf & .Fields("Codigo") & vbCrLf & NumEmpleados & vbCrLf & IR
              If NumEmpleados >= IR And Len(.fields("Codigo")) > 1 Then
                 Llenar_Rol_Pagos_Individual .fields("Codigo"), True
                 Progreso_Barra.Mensaje_Box = "Imprimiendos Rol Individual de: (" & Pagina & ") " & .fields("Nombre_Empleado")
                 Progreso_Esperar
                 If Medio_Rol Then
                   'Si es medio rol izquierda y derecha
                    Generar_Rol_Medio .fields("Codigo"), 1.3, 1
                    Generar_Rol_Medio .fields("Codigo"), 11.3, 1
                    cPrint.paginaNueva
                 Else
                   'Si es rol completo arriba y abajo
                    If Si_No Then
                       If Contador = 1 Then PosLinea = 0.5
                       If Contador = 2 Then PosLinea = PosCopiaY
                    Else
                       PosLinea = 0.5
                       Contador = 3
                    End If
                   'MsgBox PosLinea & ".........."
                    Generar_Rol .fields("Codigo"), 1.3, PosLinea
                    Contador = Contador + 1
                    If Contador > 2 Then
                       Contador = 1
                       cPrint.paginaNueva
                       Pagina = Pagina + 1
                    End If
                 End If
              End If
              NumEmpleados = NumEmpleados + 1
             .MoveNext
           Loop
       End If
      End With
      MensajeEncabData = ""
     'fin del documento
      cPrint.finalizaImpresion
     ' Presentar_PDF fPDF
     
      Set oDocument = Nothing
'      WebBPDF.navigate RutaDocumentoPDF
     'Presentar_PDF RPPDF, RutaDocumentoPDF
      RatonNormal
      MsgBox "Proceso Terminado con exito. Busque el archivo es:" & vbCrLf & vbCrLf & RutaSysBases & "\TEMP\" & NombFilePict & ".pdf"
      Presenta_Archivo_PDF RutaSysBases & "\TEMP\" & NombFilePict & ".pdf"
   Else
      RatonNormal
   End If
Exit Sub
Errorhandler:
    RatonNormal
    ErrorDeImpresion
    Exit Sub
End Sub

Public Sub Procesar_Rol_Individual_Emails(Optional Por_CI_RUC As String)
Dim AuxPosLinea As Single
Dim AntPosLinea As Single
Dim AnchoPict As Single
Dim AltoPict As Single
Dim PosCopiaY As Single
Dim NumEmpleados As Integer
Dim NombFilePict As String
Dim NombFilehtml As String
Dim Email As String
Dim posPuntoComa As Integer
Dim Un_Solo_Rol As Boolean

On Error GoTo Errorhandler
   If Len(Por_CI_RUC) > 1 Then Un_Solo_Rol = True Else Un_Solo_Rol = False
   Fecha_Del_AT CMes, CAnio
   Mensajes = "Seguro de Enviar Rol Pago por Email"
   Titulo = "ENVIO POR MAILS"
   If BoxMensaje Then
      RatonReloj
      TMail.ListaMail = 0
      TMail.Credito_No = ""
      InicioX = 0.5: InicioY = 0
      PosCopiaY = Redondear(SetPapelLargo / 2, 2) ' Largo del papel
     'MsgBox SetPapelCopia & vbCrLf & SetPapelAncho & vbCrLf & SetPapelLargo & vbCrLf & PosCopiaY & vbCrLf & Printer.ScaleHeight
      Pagina = 1
      Contador = 1
      NumEmpleados = 0
      PosLinea = 0.1
      IR = 0
      TMail.ListaMail = 0
      NombreBanco1 = UCase(MidStrg(NombreBanco1, Len(SinEspaciosIzq(NombreBanco1)) + 1, Len(NombreBanco1)))
      With AdoNomina.Recordset
       If .RecordCount > 0 Then
           Progreso_Barra.Incremento = 0
           Progreso_Barra.Valor_Maximo = .RecordCount
           If Un_Solo_Rol Then
              Generar_Rol_html Por_CI_RUC
    
             'Enviamos por mail el rol
              TMail.Asunto = "Rol Pagos " & CAnio & "-" & Format$(NumMeses, "00") & " De: " & NombreCliente
              TMail.Mensaje = "Rol Pagos " & CAnio & "-" & Format$(NumMeses, "00") & " " & vbCrLf _
                            & "Correspondiente a: " & NombreCliente & ". "
              TMail.Adjunto = ""

             'Enviamos lista de mails
              TMail.para = ""
              Insertar_Mail TMail.para, FA.EmailC
              Insertar_Mail TMail.para, FA.EmailC2
              Insertar_Mail TMail.para, FA.EmailR
              If Email_CE_Copia Then Insertar_Mail TMail.para, EmailProcesos
              FEnviarCorreos.Show vbModal
           Else
             .MoveFirst
              Do While Not .EOF
                 Progreso_Barra.Mensaje_Box = "Generando Rol Individual de: (" & Pagina & ") " & .fields("Nombre_Empleado")
                 Progreso_Esperar
                'Generamos el Rol por HTML
                 Generar_Rol_html .fields("Codigo")
    
                'Enviamos por mail el rol
                 TMail.Asunto = "Rol Pagos " & CAnio & "-" & Format$(NumMeses, "00") & " De: " & NombreCliente
                 TMail.Mensaje = "Rol Pagos " & CAnio & "-" & Format$(NumMeses, "00") & " " & vbCrLf _
                               & "Correspondiente a: " & NombreCliente & ". "
                 TMail.Adjunto = ""
                 
                'Enviamos lista de mails
                 TMail.para = ""
                 Insertar_Mail TMail.para, FA.EmailC
                 Insertar_Mail TMail.para, FA.EmailC2
                 Insertar_Mail TMail.para, FA.EmailR
                 If Email_CE_Copia Then Insertar_Mail TMail.para, EmailProcesos
                 FEnviarCorreos.Show vbModal

                 NumEmpleados = NumEmpleados + 1
                .MoveNext
              Loop
             .MoveFirst
           End If
       End If
      End With
      MensajeEncabData = ""
   Else
      RatonNormal
   End If
Exit Sub
Errorhandler:
    RatonNormal
    ErrorDeImpresion
    Exit Sub
End Sub

Public Sub Procesar_CxCxP(vTC As String)
  RatonReloj
  Fecha_Del_AT CMes, CAnio
  
 'MsgBox FechaInicial & vbCrLf & FechaFinal
  Trans_No = 100
  DGNomina.Visible = False
  DGAsiento(0).Visible = False
  
'''  sSQL = "UPDATE Trans_SubCtas " _
'''       & "SET X = '.' " _
'''       & "WHERE Item = '" & NumEmpresa & "' " _
'''       & "AND Periodo = '" & Periodo_Contable & "' "
'''  Ejecutar_SQL_SP sSQL
'''
'''  sSQL = "UPDATE Trans_SubCtas " _
'''       & "SET X = 'R' " _
'''       & "FROM Trans_SubCtas As TS, Catalogo_Rol_Pagos As CRP " _
'''       & "WHERE TS.Item = '" & NumEmpresa & "' " _
'''       & "AND TS.Periodo = '" & Periodo_Contable & "' " _
'''       & "AND TS.TC = '" & vTC & "' " _
'''       & "AND CRP.T = '" & Normal & "' " _
'''       & "AND TS.Item = CRP.Item " _
'''       & "AND TS.Periodo = CRP.Periodo " _
'''       & "AND TS.Codigo = CRP.Codigo "
'''  Ejecutar_SQL_SP sSQL

  SQL2 = "DELETE * " _
       & "FROM Asiento_SC " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND T_No = " & Trans_No & " " _
       & "AND TC = '" & vTC & "' "
  Ejecutar_SQL_SP SQL2
  
 'Procedemos a insertar las CxC o CxP del Empleado
  sSQL = "INSERT INTO Asiento_SC (Codigo, TC, Cta, Serie, Factura, CodigoU, Item, Fecha_V, TM, T_No, SC_No, DH, Valor, Detalle_SubCta) " _
       & "SELECT Codigo, TC, Cta, Serie, Factura, '" & CodigoUsuario & "', '" & NumEmpresa & "', #" & BuscarFecha(FechaFinal) & "#, '1', " _
       & Trans_No & ", ROW_NUMBER() OVER(ORDER BY Codigo, Cta ASC), "
  Select Case vTC
    Case "C": sSQL = sSQL & "2, SUM(Debitos)-SUM(Creditos), "
    Case "P": sSQL = sSQL & "1, SUM(Creditos)-SUM(Debitos), "
    Case "G": sSQL = sSQL & "1, SUM(Creditos)-SUM(Debitos), "
    Case "CC": sSQL = sSQL & "1, SUM(Creditos)-SUM(Debitos), "
  End Select
  sSQL = sSQL _
       & "'Cx" & vTC & " Empleado' " _
       & "FROM Trans_SubCtas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND Fecha_V <= #" & BuscarFecha(FechaFinal) & "# " _
       & "AND T = '" & Normal & "' " _
       & "AND TC = '" & vTC & "' " _
       & "AND Codigo IN (SELECT Codigo FROM Catalogo_Rol_Pagos WHERE Item = '" & NumEmpresa & "' AND Periodo = '" & Periodo_Contable & "' AND T = '" & Normal & "') " _
       & "GROUP BY Codigo, TC, Cta, Serie, Factura "
  If vTC = "C" Then sSQL = sSQL & "HAVING SUM(Debitos)-SUM(Creditos) > 0 " Else sSQL = sSQL & "HAVING SUM(Creditos)-SUM(Debitos) > 0 "
  sSQL = sSQL & "ORDER BY Codigo, TC, Cta, Serie, Factura "
  Ejecutar_SQL_SP sSQL, , "CxCxP_Empleado"

  sSQL = "UPDATE Asiento_SC " _
       & "SET Beneficiario = SUBSTRING(CS.Detalle,1,60) " _
       & "FROM Asiento_SC As A, Catalogo_SubCtas As CS " _
       & "WHERE CS.Item = '" & NumEmpresa & "' " _
       & "AND CS.Periodo = '" & Periodo_Contable & "' " _
       & "AND A.CodigoU = '" & CodigoUsuario & "' " _
       & "AND A.T_No = " & Trans_No & " " _
       & "AND A.Item = CS.Item " _
       & "AND A.Codigo = CS.Codigo "
  Ejecutar_SQL_SP sSQL

  sSQL = "UPDATE Asiento_SC " _
       & "SET Beneficiario = SUBSTRING(C.Cliente,1,60) " _
       & "FROM Asiento_SC As A, Clientes As C " _
       & "WHERE A.Item = '" & NumEmpresa & "' " _
       & "AND A.CodigoU = '" & CodigoUsuario & "' " _
       & "AND A.T_No = " & Trans_No & " " _
       & "AND A.Codigo = C.Codigo "
  Ejecutar_SQL_SP sSQL
  
  Eliminar_Nulos_SP "Asiento_SC"
  
  Listar_CxCxP_SubMod
  DGNomina.Visible = True
  DGAsiento(0).Visible = True
  RatonNormal
End Sub

'''Public Sub Procesar_CxP()
'''  RatonReloj
'''  Inicializar_Cero_Asientos True
'''  Fecha_Del_AT CMes, CAnio
''' 'MsgBox FechaInicial & vbCrLf & FechaFinal
'''  Trans_No = 100
'''  Contador = 0
'''  Cadena1 = ""
'''  DGNomina.Visible = False
'''  DGAsiento(0).Visible = False
'''  SQL2 = "DELETE * " _
'''       & "FROM Asiento_SC " _
'''       & "WHERE Item = '" & NumEmpresa & "' " _
'''       & "AND T_No = " & Trans_No & " " _
'''       & "AND CodigoU = '" & CodigoUsuario & "' " _
'''       & "AND TC = 'P' "
'''  Ejecutar_SQL_SP SQL2
''' 'Select_Adodc_Grid DGAsiento(0), AdoAsiento, SQL2
''' 'IniciarAsientosDe DGAsiento(0), AdoAsiento
'''  Nota_No = ReadSetDataNum("Certificados", True, False)
'''
''' 'Lista todos los Empleados para ligar con su SubCta de Modulo de CxC y CxP
'''  Si_No = False
'''  Listar_Empleados
'''  With AdoClientes.Recordset
'''   If .RecordCount > 0 Then
'''       Do While Not .EOF
'''          Total = 0
'''          Contador = Contador + 1
'''          CodigoCliente = .Fields("Codigo")
'''          NombreCliente = .Fields("Cliente")
'''          Saldos_CxC_CxP CodigoCliente, "P"
'''         .MoveNext
'''       Loop
'''       Codigo = Leer_Cta_Catalogo(SubCtaGen)
'''       If Codigo = Ninguno Then
'''          Si_No = True
'''          Cadena1 = Cadena1 & SubCtaGen & vbCrLf
'''       End If
'''   End If
'''  End With
'''  Listar_CxCxP_SubMod
'''  DGNomina.Visible = True
'''  DGAsiento(0).Visible = True
'''  RatonNormal
'''  If Si_No Then MsgBox "Codigos Contables No existen: " & vbCrLf & Cadena1
'''End Sub

'Imprimir Rol de Pagos Colectivo
Public Sub Procesar_Rol_Colectivo()
 DGAsiento(0).Visible = False
 Fecha_Del_AT CMes, CAnio
 'En el control AdoNomina esa todo lo del rol de pagos
 Orientacion_Pagina = 2
 SQLMsg1 = ""
 SQLMsg2 = "R O L    D E    P A G O S"
 SQLMsg3 = "Desde el " & FechaInicial & " al " & FechaFinal
 IR = 0 'Val(InputBox("Empezar desde: ", "IMPRESION", "0"))
 Imprimir_Rol_Colectivo AdoNomina, AdoTotNomina, True
'' MensajeEncabData = "R O L    D E    P A G O S"
'' SQLMsg1 = "Desde el " & FechaInicial & " al " & FechaFinal
'' SQLMsg2 = "PROVISIONES DEL ROL DE PAGO"
'' SQLMsg3 = ""
'' Orientacion_Pagina = 1
'' Cuadricula = True
'' ImprimirAdo AdoNominaProv, True, 1, 7
'Imprimir_Rol_de_Pagos AdoNomina, AdoTotNomina, True, 1, CLng(IR)
 DGAsiento(0).Visible = True
End Sub

Public Sub Procesar_Excel()
Dim AdoRolBancos As ADODB.Recordset
Dim sSQL1 As String

Dim NFila As Integer
Dim NColumna As Integer
Dim NCelda As Integer
Dim RutaGeneraFile As String
Dim oExcel As Object
Dim oBook As Object
Dim oSheet As Object
 RatonReloj
 
'''           Select Case SSTab1.Tab
'''             Case 0
 
 Fecha_Del_AT CMes, CAnio
 Progreso_Barra.Incremento = 0
 Progreso_Barra.Valor_Maximo = 100
 Progreso_Barra.Mensaje_Box = ""
 Progreso_Esperar
 
 RutaGeneraFile = RutaSysBases & "\TEMP\ROL PAGOS CONTRA BANCOS " & Replace(FechaFinal, "/", "-") & ".xls"
 If Dir(RutaGeneraFile) <> "" Then Kill RutaGeneraFile

'FechaIni = BuscarFecha(FechaInicial)
'FechaMid = BuscarFecha(FechaMitad)
'FechaFin = BuscarFecha(FechaFinal)
 
 Set AdoRolBancos = New ADODB.Recordset
 AdoRolBancos.CursorType = adOpenDynamic
 AdoRolBancos.CursorLocation = adUseClient
 sSQL1 = "SELECT CRP.FP + '-' + CRP.TC As F_P, C.Cliente As Detalle_Rol, CRP.Cta_Transferencia As Forma_de_Pago," _
       & "TRP.Egresos As Neto_Recibir, CRP.Religioso " _
       & "FROM Clientes As C, Catalogo_Rol_Pagos As CRP, Trans_Rol_de_Pagos As TRP " _
       & "WHERE CRP.Item = '" & NumEmpresa & "' " _
       & "AND CRP.Periodo = '" & Periodo_Contable & "' " _
       & "AND TRP.Cod_Rol_Pago = 'Neto_Recibir' " _
       & "AND Fecha_D = #" & FechaIni & "# " _
       & "AND Fecha_H >= #" & FechaFin & "# " _
       & "AND CRP.Codigo = C.Codigo " _
       & "AND CRP.Codigo = TRP.Codigo " _
       & "AND CRP.Item = TRP.Item " _
       & "AND CRP.Periodo = TRP.Periodo " _
       & "ORDER BY CRP.FP, CRP.TC, C.Cliente "
 sSQL1 = CompilarSQL(sSQL1)
 AdoRolBancos.open sSQL1, AdoStrCnn, , , adCmdText
 With AdoRolBancos
  If .RecordCount > 0 Then
      Progreso_Barra.Valor_Maximo = Progreso_Barra.Valor_Maximo + .RecordCount
     'Start a new workbook in Excel
      Set oExcel = CreateObject("Excel.Application")
      Set oBook = oExcel.Workbooks.Add
     'Add data to cells of the first worksheet in the new workbook
      Set oSheet = oBook.Worksheets(1)
      RatonReloj
     'Encabezado de la hola
     'Ancho de las columnas
      oSheet.Columns("A").columnWidth = 5
      oSheet.Columns("B").columnWidth = 60
      oSheet.Columns("C").columnWidth = 40
      oSheet.Columns("D").columnWidth = 15
     'Detalle de las columnas
      oSheet.Range("A1").value = "F_P"
      oSheet.Range("B1").value = "Detalle_Rol"
      oSheet.Range("C1").value = "Forma_de_Pago"
      oSheet.Range("D1").value = "Neto_Recibir"
      oSheet.Range("A1:D1").Font.Bold = True
     'Datos de la hoja de calculo
      NFila = 1
      Codigo = .fields("F_P")
         Select Case Codigo
           Case "T-BA": TipoCta = "POR BANCOS"
           Case "E-CJ": TipoCta = "POR EFECTIVO"
           Case "": TipoCta = ""
           Case "": TipoCta = ""
           Case "": TipoCta = ""
           Case "": TipoCta = ""
           Case Else: TipoCta = ""
         End Select
      
      Do While Not .EOF
         NFila = NFila + 1
         
         Select Case .fields("F_P")
           Case "": TipoCta = ""
           Case "": TipoCta = ""
           Case "": TipoCta = ""
           Case "": TipoCta = ""
           Case "": TipoCta = ""
           Case "": TipoCta = ""
           Case Else: TipoCta = ""
         End Select

         
        'Religioso
         oSheet.Range("A" & CStr(NFila)).value = .fields("F_P")
         oSheet.Range("B" & CStr(NFila)).value = .fields("Detalle_Rol")
         oSheet.Range("C" & CStr(NFila)).value = .fields("Forma_de_Pago")
         oSheet.Range("D" & CStr(NFila)).value = .fields("Neto_Recibir")
        .MoveNext
      Loop
     'Bloqueamos las celdas que no se puden cambiar
'''      For NCelda = 1 To 6
'''          With oSheet.Cells(1, NCelda)          ' seleccionamos la 1ª celda
'''                  .Interior.ColorIndex = 41     ' Color fondo = azul '41
'''                  .Font.Size = 9                ' tamaño de letra
'''                  .Font.Bold = True             ' Fuente en negrita
'''                  .Font.ColorIndex = 2          ' Color fuente = blanco
'''          End With
'''          With oSheet.Cells(NFila + 1, NCelda)  ' seleccionamos la 1ª celda
'''              .Interior.ColorIndex = 41         ' Color fondo = azul '41
'''              .Font.Size = 9                    ' tamaño de letra
'''              .Font.Bold = True                 ' Fuente en negrita
'''              .Font.ColorIndex = 2              ' Color fuente = blanco
'''          End With
'''      Next NCelda
'''      oSheet.Unprotect "DiskCoverEducativo"
'''      oSheet.Range("B2:B" & CStr(NFila)).Locked = False
'''      oSheet.Protect "DiskCoverEducativo"
     'Save the Workbook and Quit Excel
      oBook.SaveAs RutaGeneraFile
      oExcel.Quit
  End If
 End With
 RatonNormal
 AdoRolBancos.Close
 Progreso_Final
 MsgBox "EL ARCHIVO SE GRABO EN: " & RutaGeneraFile
End Sub

Private Sub Command3_Click()
    Unload LRolPagos
End Sub

'Grabar Rol a la Contabilidad
Public Sub Procesar_Grabar()
  Trans_No = 100
  FechaComp = FechaFinal
  SQL2 = "SELECT * " _
       & "FROM Asiento " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND T_No = " & Trans_No & " " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "ORDER BY A_No "
  Select_Adodc_Grid DGAsiento(0), AdoAsiento, SQL2
  SumaDebe = 0: SumaHaber = 0
  With AdoAsiento.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          SumaDebe = SumaDebe + .fields("DEBE")
          SumaHaber = SumaHaber + .fields("HABER")
         .MoveNext
       Loop
      .MoveFirst
   End If
  End With
  SQL2 = "SELECT * " _
       & "FROM Asiento_SC " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND T_No = " & Trans_No & " " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "ORDER BY Codigo,TC,Cta "
  Select_Adodc AdoAsientoSC, SQL2
  'MsgBox SumaDebe & vbCrLf & SumaHaber
  If Redondear(SumaDebe - SumaHaber, 2) = 0 Then
     RatonReloj
     Co.T = Normal
     Co.Fecha = FechaFinal
     Co.CodigoB = Ninguno
     Co.Efectivo = TotalCajaMN
     Co.Monto_Total = TotalCajaMN + Total_Cheque + Total_Bancos
     Co.Item = NumEmpresa
     Co.Usuario = CodigoUsuario
     If CheqCD.value <> 1 Then
        Co.TP = CompDiario
        Co.Numero = ReadSetDataNum("Diario", True, True)
     Else
        Co.TP = CompEgreso
        Co.Numero = ReadSetDataNum("Egresos", True, True)
     End If
     Co.Concepto = LblConcepto(0).Caption
     Co.T_No = Trans_No
     Grabar_Comprobante Co
     ImprimirComprobantesDe False, Co
     SQL2 = "UPDATE Trans_Rol_de_Pagos " _
          & "SET TP = '" & Co.TP & "'," _
          & "Numero = " & Co.Numero & " " _
          & "WHERE Fecha_D >= #" & FechaIni & "# " _
          & "AND Fecha_H <= #" & FechaFin & "# " _
          & "AND Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' "
     If CmbGrupos <> "TODOS" Then SQL2 = SQL2 & "AND Grupo_Rol = '" & CmbGrupos & "' "
     Ejecutar_SQL_SP SQL2
     
    'Grabamos las provisiones
     Trans_No = 101
     Co.TP = CompDiario
     Co.Efectivo = 0
     Co.Monto_Total = 0
     Co.Numero = ReadSetDataNum("Diario", True, True)
     Co.Concepto = LblConcepto(1).Caption
     Co.T_No = Trans_No
     Grabar_Comprobante Co
     ImprimirComprobantesDe False, Co

     Trans_No = 102
     Co.TP = CompDiario
     Co.Numero = ReadSetDataNum("Diario", True, True)
     Co.Concepto = LblConcepto(2).Caption
     Co.T_No = Trans_No
     Grabar_Comprobante Co
     ImprimirComprobantesDe False, Co
     Unload LRolPagos
  Else
     MsgBox "Las Transacciones no cuadran"
  End If
End Sub


Private Sub DCCxP_LostFocus()
  Cta = Leer_Cta_Catalogo(SinEspaciosIzq(DCCxP))
  Label3.Caption = SubCta
End Sub

Private Sub DGAsiento_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF1 Then
     Select Case index
       Case 0: GenerarDataTexto LRolPagos, AdoAsiento
       Case 1: GenerarDataTexto LRolPagos, AdoAsiento1
       Case 2: GenerarDataTexto LRolPagos, AdoAsiento2
     End Select
  End If
End Sub

Private Sub DGNomina_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  
  If AdoNomina.Recordset.RecordCount > 0 Then
     CodigoCli = DGNomina.Columns(DGNomina.Columns.Count - 2)  ' Codigo del Empleado
     NombreCliente = DGNomina.Columns(2)                       ' Nombre del Empleado
     CodigoP = DGNomina.Columns(4)                             ' Grupo del Empleado
  Else
     CodigoCli = Ninguno
     NombreCliente = Ninguno
     CodigoP = Ninguno
  End If
  
  If CtrlDown And (vbKeyF9 = KeyCode) Then
     Mifecha = FechaFinal
     FComisiones.Show 1
     'Procesar_Asientos_Rol
  End If
  
  If CtrlDown And (vbKeyF10 = KeyCode) Then
     Mifecha = FechaFinal
     FRolPago.Show 1
  End If
  
  If CtrlDown And (vbKeyI = KeyCode) Then
     Valor = Val(InputBox("INGRESE EL IMPUESTO A LA RENTA MANUAL DE " & NombreCliente & ":", "IMPUESTO A LA RENTA", "0"))
     If Valor >= 0 Then
        SQL2 = "DELETE * " _
             & "FROM Trans_Rol_de_Pagos " _
             & "WHERE Item = '" & NumEmpresa & "' " _
             & "AND Periodo = '" & Periodo_Contable & "' " _
             & "AND Fecha_D = '" & FechaIni & "' " _
             & "AND Codigo = '" & CodigoCli & "' " _
             & "AND Cod_Rol_Pago = 'Imp_Renta' " _
             & "AND SubModulo = 'M' " _
             & "AND Tipo_Rubro = 'PER' "
        Ejecutar_SQL_SP SQL2
        
        SQL2 = "INSERT INTO Trans_Rol_de_Pagos (Item,Periodo,CodigoU,Grupo_Rol,Codigo,Tipo_Rubro,Cta,Egresos,Horas,Fecha_D,Fecha_H,SubModulo,Cod_Rol_Pago,IESS) " _
             & "VALUES ('" & NumEmpresa & "','" & Periodo_Contable & "','" & CodigoUsuario & "','" & CodigoP & "','" & CodigoCli & "','PER','" & Cta_Impuesto_Renta_Empleado & "'," _
             & Valor & ",0,'" & FechaIni & "','" & FechaFin & "','M','Imp_Renta',0)"
        Ejecutar_SQL_SP SQL2
        
        SQL2 = "UPDATE Catalogo_Rol_Pagos " _
             & "SET Calculo_Manual_IR = 1 " _
             & "WHERE Item = '" & NumEmpresa & "' " _
             & "AND Periodo = '" & Periodo_Contable & "' " _
             & "AND Codigo = '" & CodigoCli & "' "
        Ejecutar_SQL_SP SQL2
     End If
  End If
  
  If CtrlDown And (vbKeyD = KeyCode) Then
     Mensajes = "Seguro desactivar a " & NombreCliente & " del calculo manual del I.R."
     Titulo = "DESACTIVAR CALCULO MANUAL I.R."
     If BoxMensaje Then
        SQL2 = "UPDATE Catalogo_Rol_Pagos " _
             & "SET Calculo_Manual_IR = 0 " _
             & "WHERE Item = '" & NumEmpresa & "' " _
             & "AND Periodo = '" & Periodo_Contable & "' " _
             & "AND Codigo = '" & CodigoCli & "' "
        Ejecutar_SQL_SP SQL2
     End If
  End If
  
  If CtrlDown And (vbKey0 = KeyCode) Then
     Mensajes = "Seguro de activar a " & NombreCliente & " el calculo del I.R."
     Titulo = "ACTIVAR I.R. EN CERO"
     If BoxMensaje Then
        SQL2 = "UPDATE Catalogo_Rol_Pagos " _
             & "SET No_Calcular_IR = 1 " _
             & "WHERE Item = '" & NumEmpresa & "' " _
             & "AND Periodo = '" & Periodo_Contable & "' " _
             & "AND Codigo = '" & CodigoCli & "' "
        Ejecutar_SQL_SP SQL2
     End If
  End If
  
  If CtrlDown And (vbKeyN = KeyCode) Then
     Mensajes = "Seguro Activar a " & NombreCliente & " en calculo I.R."
     Titulo = "DESACTIVAR I.R. EN CERO"
     If BoxMensaje Then
        SQL2 = "UPDATE Catalogo_Rol_Pagos " _
             & "SET No_Calcular_IR = 0 " _
             & "WHERE Item = '" & NumEmpresa & "' " _
             & "AND Periodo = '" & Periodo_Contable & "' " _
             & "AND Codigo = '" & CodigoCli & "' "
        Ejecutar_SQL_SP SQL2
     End If
  End If
  
  If AdoNomina.Recordset.RecordCount > 0 Then
     If CtrlDown And (vbKeyI = KeyCode) Then AdoNomina.Recordset.MoveFirst
     If CtrlDown And (vbKeyF = KeyCode) Then AdoNomina.Recordset.MoveLast
  End If
  
  If KeyCode = vbKeyF1 Then GenerarDataTexto LRolPagos, AdoNomina
End Sub

Private Sub DGNomina1_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  If CtrlDown And (vbKeyP = KeyCode) Then ImprimirAdodc AdoNomina1, True, 2, 7
End Sub

Private Sub DGSubCtas_KeyDown(KeyCode As Integer, Shift As Integer)
   Keys_Especiales Shift
   If KeyCode = vbKeyF1 Then GenerarDataTexto LRolPagos, AdoAsientoSC
End Sub

Private Sub DGTotNomina_KeyDown(KeyCode As Integer, Shift As Integer)
   Keys_Especiales Shift
   If KeyCode = vbKeyF1 Then GenerarDataTexto LRolPagos, AdoTotNomina
End Sub

Private Sub Form_Activate()
Dim AltoTab As Single
Dim AnchoTab As Single
Dim InicioTab As Single
Dim MitadTab As Single
Dim CuatrupleTab As Single

  FechaInicial = FechaSistema
  FechaFinal = FechaSistema
  Datos_IESS FechaFinal
  'CtaImpRenta = Leer_Cta_Catalogo(Cta_Impuesto_Renta_Empleado)
  
  If Sueldo_Basico <= 0 Then
     RatonNormal
     MsgBox "Falta setear el sueldo Basico, comuniquese con el Administrador del Sistema"
     Unload Me
  Else
    'Presentar_PDF RPPDF, RutaDocumentoPDF
    'Ancho y Largo de la pantalla
     AnchoTab = MDI_X_Max - 100
     AltoTab = MDI_Y_Max - 1650
     MitadTab = (MDI_Y_Max - 2800) / 2
     CuatrupleTab = (MDI_Y_Max - 2450) / 4
     InicioTab = SSTab1.Top
    
     SSTab1.width = AnchoTab
     SSTab1.Height = AltoTab
     
     SSTab1.Tab = 5
     
     SetPapelLargo = 29
     CAnio.Clear
     For I = Year(FechaSistema) To Year(FechaSistema) - 10 Step -1
         CAnio.AddItem CStr(I)
     Next I
     CAnio.Text = CStr(Year(FechaSistema))
     
     CMes.Clear
     For IE = 12 To 1 Step -1
         CMes.AddItem MesesLetras(IE)
     Next IE
     CMes.Text = CStr(MesesLetras(Month(FechaSistema)))
    
     Inicializar_Cero_Asientos
     Trans_No = 100
    'Pagos sin alcance
     sSQL = "SELECT (Codigo & ' - ' & Cuenta) As Cuentas, * " _
          & "FROM Catalogo_Cuentas " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "AND TC IN ('P','PS') " _
          & "AND DG = 'D' " _
          & "ORDER BY Codigo "
     SelectDB_Combo DCCxP, AdoBanco, sSQL, "Cuentas"
     DCCxP.Visible = False
     Listar_CxCxP_SubMod
     
     FechaIni = BuscarFecha(PrimerDiaMes(FechaSistema))
     FechaFin = BuscarFecha(UltimoDiaMes(FechaSistema))
     
     Reporte_Rol_Pagos_Colectivo_SP FechaIni, FechaFin, "TODOS", True, SubSQL, SQL2
     
     DGNomina1.Caption = "EMPLEADOS SIN ALCANCE DE REMUNERACION"
     sSQL = "SELECT C.Cliente As Empleado, TRP.Egresos as Neto_a_Recibir " _
          & "FROM Clientes as C,Trans_Rol_de_Pagos As TRP " _
          & "WHERE Fecha_D = #20501231# " _
          & "AND Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "AND Egresos <= 0 " _
          & "AND TRP.Codigo = C.Codigo " _
          & "ORDER BY Cliente "
     Select_Adodc_Grid DGNomina1, AdoNomina1, sSQL
     'MsgBox sSQL
     
     sSQL = "SELECT " & SubSQL & " " _
          & "FROM Reporte_Rol_Colectivo " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND CodigoU = '" & CodigoUsuario & "' "
     Select_Adodc_Grid DGNomina, AdoNomina, sSQL, 2, True
     
     sSQL = "SELECT C.Cliente, CRR.I_E, CRR.Detalle, CRR.Cta, CRR.Valor, CRR.Calc_IESS, CRR.Cod_Rol_Pago " _
          & "FROM Catalogo_Rol_Rubros As CRR, Clientes As C " _
          & "WHERE CRR.Item = '" & NumEmpresa & "' " _
          & "AND CRR.Periodo = '" & Periodo_Contable & "' " _
          & "AND CRR.Mes = 0 " _
          & "AND CRR.Codigo = C.Codigo " _
          & "ORDER BY C.Cliente, CRR.Codigo, CRR.I_E, CRR.Cta "
     Select_Adodc_Grid DGI_E_Empleado, AdoI_E_Empleado, sSQL
     'MsgBox "..."
     'WebBPDF.navigate RutaSistema & "\FONDOS\index.html"
     APDFRol.Object.src = ""
     
     LblConcepto(1).width = AnchoTab - 350
     DGAsiento(1).width = AnchoTab - 350
     DGAsiento(1).Height = MitadTab - 1000
     LblConcepto(2).width = AnchoTab - 350
     DGAsiento(2).width = AnchoTab - 350
     DGAsiento(2).Top = DGAsiento(1).Height + 1200
     DGAsiento(2).Height = AltoTab - DGAsiento(2).Top - 150
     LblConcepto(2).Top = DGAsiento(1).Height + 850

     SSTab1.Tab = 4
     LblConcepto(0).width = AnchoTab - 350
     DGAsiento(0).width = AnchoTab - 350
     DGAsiento(0).Height = SSTab1.Height - 1300
     Label1.Top = SSTab1.Height - 500
     Label19.Top = SSTab1.Height - 500
     LabelDiferencia.Top = SSTab1.Height - 500
     LabelDebe.Top = SSTab1.Height - 500
     LabelHaber.Top = SSTab1.Height - 500
     
     SSTab1.Tab = 3
     DGI_E_Empleado.width = AnchoTab - 350
     DGI_E_Empleado.Height = AltoTab - 600
     
     SSTab1.Tab = 2
     DGNomina1.width = AnchoTab - 350
     DGNomina1.Height = MitadTab
     DGSubCtas.Top = DGNomina1.Height + 440
     DGSubCtas.width = AnchoTab - 350
     DGSubCtas.Height = AltoTab - DGSubCtas.Top - 150
     
     SSTab1.Tab = 1
     AdoNomina.Top = SSTab1.Height - 500
     DGNomina.width = AnchoTab - 350
     DGNomina.Height = CuatrupleTab * 3
     'DGTotNomina.Top = SSTab1.Height - 1250
     DGTotNomina.Height = CuatrupleTab
     DGTotNomina.Top = DGNomina.Top + DGNomina.Height
     DGTotNomina.width = AnchoTab - 350
          
     SSTab1.Tab = 0
     APDFRol.width = SSTab1.width - 250
     APDFRol.Height = SSTab1.Height - 650
     
     RatonNormal
     CAnio.SetFocus
  End If
End Sub

Private Sub Form_Load()
  ConectarAdodc AdoAux
  ConectarAdodc AdoCaja
  ConectarAdodc AdoBanco
  ConectarAdodc AdoNomina
  ConectarAdodc AdoCtaCat
  ConectarAdodc AdoNovedades
  ConectarAdodc AdoNominaProv
  ConectarAdodc AdoClientes
  ConectarAdodc AdoCertificado
  ConectarAdodc AdoTotNomina
  ConectarAdodc AdoNomina1
  ConectarAdodc AdoSubCta
  ConectarAdodc AdoSubCta1
  ConectarAdodc AdoAsiento
  ConectarAdodc AdoAsiento1
  ConectarAdodc AdoAsiento2
  ConectarAdodc AdoAsientoB
  ConectarAdodc AdoAsientoSC
  ConectarAdodc AdoAsientoRol
  ConectarAdodc AdoImpRenta
  ConectarAdodc AdoI_E_Empleado
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
   MyTime = Time
  'MsgBox Button.key
   Select Case Button.key
     Case "Salir"
          Unload Me
     Case "CxC"
          Procesar_CxCxP "C"
     Case "CxP"
          Procesar_CxCxP "P"
     Case "Quincena"
          Procesar_Quincena
     Case "Nomina"
          'Procesar_Nomina
          Procesar_Rol_De_Pagos_Mes
     Case "Rol_Individual"
          Procesar_Rol_Individual
     Case "Rol_Colectivo"
          Procesar_Rol_Colectivo
     Case "Grabar"
          Procesar_Grabar
     Case "Limpiar"
          Procesar_Limpiar
     Case "Excel"
          Procesar_Excel
     Case "Email"
          With AdoNomina.Recordset
           If .RecordCount > 0 Then Procesar_Rol_Individual_Emails .fields("Codigo")
          End With
     Case "Emails"
          Procesar_Rol_Individual_Emails
     Case "Primero"
          With AdoNomina.Recordset
           If .RecordCount > 0 Then
              .MoveFirst
               Llenar_Rol_Pagos_Individual .fields("Codigo")
           End If
          End With
     Case "Antes"
          With AdoNomina.Recordset
           If .RecordCount > 0 Then
              .MovePrevious
               If .BOF Then .MoveLast
               Llenar_Rol_Pagos_Individual .fields("Codigo")
           End If
          End With
     Case "Despues"
          With AdoNomina.Recordset
           If .RecordCount > 0 Then
              .MoveNext
               If .EOF Then .MoveFirst
               Llenar_Rol_Pagos_Individual .fields("Codigo")
           End If
          End With
     Case "Ultimo"
          With AdoNomina.Recordset
           If .RecordCount > 0 Then
              .MoveLast
               If .fields("Codigo") = "T O T A L " Then .MovePrevious
               Llenar_Rol_Pagos_Individual .fields("Codigo")
           End If
          End With
     Case "Buscar"
          TipoDatoBusq = TadText
          CampoBusqueda = "Codigo"
          FBusqueda.Show 1
          If TextoBusqueda <> "" Then
             With AdoNomina.Recordset
              If .RecordCount > 0 Then
                 .MoveFirst
                 .Find (CampoBusqueda & TextoBusqueda)
                  If .EOF Then
                      MsgBox "No existe este codigo"
                     .MoveFirst
                  End If
                  Llenar_Rol_Pagos_Individual .fields("Codigo")
              End If
             End With
          End If
   End Select
End Sub

Private Sub TxtCheque_GotFocus()
  MarcarTexto TxtCheque
End Sub

Private Sub TxtCheque_LostFocus()
  TextoValido TxtCheque, , True
End Sub

Public Sub Limpiar_Rol_Individual()
'.SubModulo = Ninguno
  With TRol_Pago
      .T = Normal
      .Cta = Ninguno
      .Detalle = Ninguno
      .Cheq_Dep_Transf = Ninguno
      .Tipo_Rubro = Ninguno
      .Ingresos = 0
      .Egresos = 0
      .Dias = 0
      .Horas = 0
      .Porc = 0
      .Retencion_No = 0
      .ID = 0
  End With
End Sub

Public Sub Generar_Rol(CodigoRol As String, Xo As Single, Yo As Single)
Dim ContLineas As Integer
Dim Es_Vacaciones As Boolean
Dim No_Recibe_Sueldo As Boolean
Dim PFilIni As Single
Dim Tot_Ingresos As Currency
Dim Tot_Egresos As Currency

'Empezamos a Escribir en papel grafico el Rol Individual
'Los rubros que se ingresaron anteriormente con el rol
 cPrint.tipoDeLetra = TipoCourier 'TipoTimesRoman
 cPrint.tipoNegrilla = False
 cPrint.PorteDeLetra = 10
 No_Recibe_Sueldo = True
 With AdoAsientoRol.Recordset
  If .RecordCount > 0 Then
     'Es_Vacaciones = .Fields("Vac")
      Tot_Ingresos = 0
      Tot_Egresos = 0
      
      cPrint.tipoNegrilla = True
      cPrint.printImagen LogoTipo, Xo, Yo, 3, 1.4
      RutaDestino = RutaSistema & "\FOTOS\" & CodigoRol & ".jpg"
     'cPrint.printImagen  RutaDestino, 7.5, 1, 2.5, 3
      cPrint.PorteDeLetra = 15
      If UCase$(RazonSocial) = UCase$(NombreComercial) Then
         cPrint.printTexto Xo + 3.5, Yo, UCase$(RazonSocial)
      Else
         cPrint.printTexto Xo + 3.5, Yo, UCase$(RazonSocial)
         cPrint.printTexto Xo + 3.5, Yo + 0.5, UCase$(NombreComercial)
      End If
      cPrint.PorteDeLetra = 10
      cPrint.printTexto Xo + 3.5, Yo + 0.95, "Direccion: " & Direccion
      cPrint.PorteDeLetra = 12
      cPrint.printTexto Xo + 3.5, Yo + 1.5, "ROL INDIVIDUAL DE PAGOS"
      cPrint.PorteDeLetra = 9
      cPrint.printTexto Xo + 12.5, Yo + 1.9, "Desde: " & FechaInicial & " al: " & FechaFinal
      
      cPrint.printLinea Xo, Yo + 2.4, 19.5, Yo + 2.4
      
      cPrint.printTexto Xo + 0.1, Yo + 2.5, "Fecha de Ingreso:"
      cPrint.printTexto Xo + 0.1, Yo + 3, "Beneficiario:"
      cPrint.printTexto Xo + 11.6, Yo + 2.5, "Codigo:"
      cPrint.printTexto Xo + 0.1, Yo + 3.5, "Periodo:"
      cPrint.printTexto Xo + 16.7, Yo + 2.5, "Días:"
      cPrint.printTexto Xo + 11.6, Yo + 3.5, "Forma de Pago:"
      cPrint.tipoNegrilla = False
      cPrint.printTexto Xo + 3.7, Yo + 2.5, FechaTexto
      cPrint.printTexto Xo + 2.8, Yo + 3, NombreCliente
      cPrint.printTexto Xo + 13.1, Yo + 2.5, CICliente  'CodigoRol
      cPrint.printTexto Xo + 1.8, Yo + 3.5, MesesLetras(Month(FechaFinal))
      cPrint.tipoNegrilla = True

      PFil = Yo + 4
      PFilIni = Yo + 4
      cPrint.printLinea Xo, Yo + 4, 19.5, Yo + 4
      PFil = PFil + 0.05
      cPrint.tipoNegrilla = True
      cPrint.printTexto Xo + 0.1, PFil, "D E T A L L E    D E L    E M P L E A D O"
      cPrint.printTexto Xo + 14.5, PFil, "INGRESOS"
      cPrint.printTexto Xo + 17, PFil, "EGRESOS"
      PFil = PFil + 0.4
      cPrint.printLinea Xo, PFil, 19.5, PFil
      PFil = PFil + 0.1
     .MoveFirst
      Do While Not .EOF
        'MsgBox .Fields("Tipo_Rubro") & vbCrLf & .Fields("Detalle") & vbCrLf & .Fields("Cod_Rol_Pago")
         If .fields("Tipo_Rubro") = "PER" Then
             If .fields("Cod_Rol_Pago") = "Neto_Recibir" Then
                 cPrint.tipoNegrilla = True
                 No_Recibe_Sueldo = False
                 cPrint.printLinea Xo, PFil, 19.5, PFil
                 PFil = PFil + 0.1
                 cPrint.printTexto Xo + 0.1, PFil, "SUBTOTALES DE INGRESOS Y EGRESOS"
                 cPrint.printVariable Xo + 12.7, PFil, Tot_Ingresos
                 cPrint.printVariable Xo + 15.2, PFil, Tot_Egresos
                 PFil = PFil + 0.4
                 cPrint.printLinea Xo, PFil, 19.5, PFil
                 PFil = PFil + 0.1
                 cPrint.printTexto Xo + 0.1, PFil, UCaseStrg(.fields("Detalle"))
             Else
                 cPrint.tipoNegrilla = False
                 cPrint.printTexto Xo + 0.1, PFil, UCase(.fields("Detalle"))
             End If
             If .fields("Ingresos") <> 0 Then
                 cPrint.printFields Xo + 12.7, PFil, .fields("Ingresos")
             End If
             If .fields("Egresos") <> 0 Then
                 cPrint.printFields Xo + 15.2, PFil, .fields("Egresos")
             End If
             If .fields("Cheq_Dep_Transf") <> Ninguno Then CodigoB = .fields("Cheq_Dep_Transf")
             If .fields("Dias") <> 0 Then I = .fields("Dias")
             If .fields("Horas") <> 0 Then J = .fields("Horas")
             Tot_Ingresos = Tot_Ingresos + .fields("Ingresos")
             Tot_Egresos = Tot_Egresos + .fields("Egresos")
             
             PFil = PFil + 0.4
         End If
        .MoveNext
      Loop
'''      If No_Recibe_Sueldo Then
'''         cPrint.tipoNegrilla = True
'''         cPrint.printTexto Xo, PFil, "TOTAL A RECIBIR"
'''         cPrint.printTexto Xo + 15.5, PFil, "0.00", True, 1.9
'''         PFil = PFil + 0.4
'''      End If
      cPrint.tipoNegrilla = False
      PFil = PFil - 0.1
      'cPrint.printTexto   Xo + 14.6, Yo + 4.5, Format(J, "#,##0.00")
      cPrint.printTexto Xo + 18, Yo + 2.5, Format(I, "#,##0") ' DIAS
      cPrint.printTexto Xo + 14.8, Yo + 3.5, CodigoB
      cPrint.tipoNegrilla = False
      'cPrint.printLinea Xo, PFil - 0.3, 19.5, PFil - 0.4
      cPrint.printLinea Xo, PFil + 0.05, 19.5, PFil + 0.05
      PFil = PFil + 0.1
      If Len(Lista_Emails) > 3 Then cPrint.printTexto Xo + 0.1, PFil, "EMAIL(S) DE ENVIO(S): " & Lista_Emails
      PFil = PFil + 0.4
      cPrint.PorteDeLetra = 10
      sSQL = "SELECT " & Full_Fields("Trans_Entrada_Salida") & " " _
           & "FROM Trans_Entrada_Salida " _
           & "WHERE Fecha BETWEEN #" & BuscarFecha(FechaInicial) & "# AND #" & BuscarFecha(FechaFinal) & "# " _
           & "AND Codigo = '" & CodigoRol & "' " _
           & "AND Item = '" & NumEmpresa & "' " _
           & "AND Periodo = '" & Periodo_Contable & "' " _
           & "AND ES = 'R' " _
           & "ORDER BY Fecha,Hora "
      Select_Adodc AdoNovedades, sSQL
      With AdoNovedades.Recordset
       If .RecordCount > 0 Then
           cPrint.tipoNegrilla = True
           cPrint.printTexto Xo + 1.1, PFil, "OBSERVACIONES:"
           cPrint.tipoNegrilla = False
           PFil = PFil + 0.35
           Do While Not .EOF
              cPrint.printTexto Xo + 1.1, PFil, .fields("Tarea")
              PFil = PFil + 0.35
             .MoveNext
           Loop
       End If
      End With
      cPrint.printTexto Xo + 4, Yo + 11.5, String(12, "_")
      cPrint.printTexto Xo + 9, Yo + 11.5, String(17, "_")
      cPrint.printTexto Xo + 4.2, Yo + 12, "Empleador"
      cPrint.printTexto Xo + 9.5, Yo + 12, "Recibi conforme"
      DetalleComp = ""
  End If
 End With
End Sub

Public Sub Generar_Rol_html(CodigoRol As String)
Dim AdoRolEmpleado As ADODB.Recordset
Dim AdoRolDetalle As ADODB.Recordset
Dim AdoRolNovedades As ADODB.Recordset
Dim Es_Vacaciones As Boolean
Dim No_Recibe_Sueldo As Boolean
Dim ContLineas As Integer
Dim PFilIni As Single
Dim Estilos As String

'Empezamos a generar Rol Individual en formato HTML
 TMail.MensajeHTML = Leer_Archivo_Texto(RutaSistema & "\JAVASCRIPT\f_rol_pago.html")
 sSQL = "SELECT C.Cliente, C.CI_RUC, C.Email, C.Email2, C.EmailR, CRP.Fecha, CRP.Grupo_Rol, CRP.Dias_Mes, CRP.Cta_Transferencia, CRP.Mes " _
      & "FROM Clientes As C, Catalogo_Rol_Pagos As CRP " _
      & "WHERE CRP.Item = '" & NumEmpresa & "' " _
      & "AND CRP.Periodo = '" & Periodo_Contable & "' " _
      & "AND C.Codigo = '" & CodigoRol & "' " _
      & "AND C.Codigo = CRP.Codigo "
 Select_AdoDB AdoRolEmpleado, sSQL
 With AdoRolEmpleado
  If .RecordCount > 0 Then
      FA.Cliente = .fields("Cliente")
      FA.Grupo = .fields("Grupo_Rol")
      FA.Fecha_R = .fields("Fecha")
      FA.CI_RUC = .fields("CI_RUC")
      FA.EmailC = .fields("Email")
      FA.EmailC2 = .fields("Email2")
      FA.EmailR = .fields("EmailR")
      FA.Observacion = .fields("Cta_Transferencia")
      FA.Fecha = FechaInicial
      FA.Fecha_V = FechaFinal
      NoMes = .fields("Mes")
      NombreCliente = FA.Cliente
  End If
 End With
 AdoRolEmpleado.Close
 
 html_Titulo_Mensaje = "ROL DE PAGOS INDIVIDUAL MES DE " & MesesLetras(CInt(Month(FechaInicial)), True)
 Mifecha = FechaFinal
 
'Presentamos el rol individual del empleado
 TextoXML = ""
 sSQL = "SELECT " & Full_Fields("Trans_Rol_de_Pagos") & " " _
      & "FROM Trans_Rol_de_Pagos " _
      & "WHERE Item = '" & NumEmpresa & "' " _
      & "AND Periodo = '" & Periodo_Contable & "' " _
      & "AND Fecha_D >= #" & FechaIni & "# " _
      & "AND Fecha_H <= #" & FechaFin & "# " _
      & "AND Codigo = '" & CodigoRol & "' " _
      & "AND Tipo_Rubro = 'PER' " _
      & "ORDER BY ID, Ingresos desc, Egresos, Detalle "
 Select_AdoDB AdoRolDetalle, sSQL
 SumaDebe = 0
 SumaHaber = 0
 With AdoRolDetalle
  If .RecordCount > 0 Then
      Do While Not .EOF
         Select Case .fields("Cod_Rol_Pago")
          Case "Neto_Recibir"
               If .fields("Cheq_Dep_Transf") <> Ninguno Then FA.Observacion = .fields("Cheq_Dep_Transf") Else FA.Observacion = "EN EFECTIVO"
               ValorTotal = .fields("Egresos")
          Case Else
               SumaDebe = SumaDebe + .fields("Ingresos")
               SumaHaber = SumaHaber + .fields("Egresos")
               If .fields("Cod_Rol_Pago") = "Salario" Then
                   If .fields("Dias") <> 0 Then I = .fields("Dias")
                   If .fields("Horas") <> 0 Then J = .fields("Horas")
               End If
               Insertar_Campo_XML AbrirXML("tr")
               Insertar_Campo_XML CampoIdXML("td", "", UCaseStrg(Sin_Signos_Especiales(.fields("Detalle"))))
               If .fields("Ingresos") <> 0 Then
                    Insertar_Campo_XML CampoIdXML("td", "class='row text-right'", Format(.fields("Ingresos"), "#,##0.00"))
                    Insertar_Campo_XML CampoIdXML("td", "class='row text-right'", Chr(255))
               End If
               If .fields("Egresos") <> 0 Then
                    Insertar_Campo_XML CampoIdXML("td", "class='row text-right'", Chr(255))
                    Insertar_Campo_XML CampoIdXML("td", "class='row text-right'", Format(.fields("Egresos"), "#,##0.00"))
               End If
               Insertar_Campo_XML CerrarXML("tr")
         End Select
        .MoveNext
      Loop
   End If
 End With
 AdoRolDetalle.Close
 
 html_Detalle_adicional = TextoXML
 
 html_Informacion_adicional = ""
 sSQL = "SELECT Fecha, Hora, Tarea " _
      & "FROM Trans_Entrada_Salida " _
      & "WHERE Fecha BETWEEN #" & BuscarFecha(FechaInicial) & "# AND #" & BuscarFecha(FechaFinal) & "# " _
      & "AND Codigo = '" & CodigoRol & "' " _
      & "AND Item = '" & NumEmpresa & "' " _
      & "AND Periodo = '" & Periodo_Contable & "' " _
      & "AND ES = 'R' " _
      & "ORDER BY Fecha,Hora "
 Select_AdoDB AdoRolNovedades, sSQL
 With AdoRolNovedades
  If .RecordCount > 0 Then
      html_Informacion_adicional = html_Informacion_adicional & "<a><B>OBSERVACIONES</B></a><BR>"
      Do While Not .EOF
         html_Informacion_adicional = html_Informacion_adicional & "=> " & .fields("Tarea") & vbCrLf
        .MoveNext
      Loop
  End If
 End With
 AdoRolNovedades.Close
End Sub

Public Sub Generar_Rol_Medio(CodigoRol As String, Xo As Single, Yo As Single)
Dim ContLineas As Integer
Dim Es_Vacaciones As Boolean

'Empezamos a Escribir en papel grafico el Rol Individual
'Los rubros que se ingresaron anteriormente con el rol
 cPrint.tipoNegrilla = False
 cPrint.PorteDeLetra = 7
 With AdoAsientoRol.Recordset
 'MsgBox CodigoRol & vbCrLf & .RecordCount
  If .RecordCount > 0 Then
      J = 0
      PFil = Yo
     'Es_Vacaciones = .Fields("Vac")
      cPrint.tipoNegrilla = True
      cPrint.printImagen LogoTipo, Xo, Yo, 2.4, 1.1
      RutaDestino = RutaSistema & "\FOTOS\" & CodigoRol & ".jpg"
     'cPrint.printImagen  RutaDestino, 7.5, 1, 2.5, 3
      If UCase$(RazonSocial) = UCase$(NombreComercial) Then
         cPrint.printTexto Xo + 2.5, PFil, UCase$(RazonSocial)
         PFil = PFil + 0.3
      Else
         cPrint.printTexto Xo + 2.5, PFil, UCase$(RazonSocial)
         PFil = PFil + 0.3
         cPrint.printTexto Xo + 2.5, PFil, UCase$(NombreComercial)
         PFil = PFil + 0.3
      End If
      cPrint.printTexto Xo + 2.5, PFil, "R.U.C. " & RUC
      PFil = PFil + 0.3
      cPrint.PorteDeLetra = 6
      cPrint.printTexto Xo + 2.5, PFil, "Direccion: " & ULCase(Direccion)
      PFil = Yo + 1.3
      cPrint.printLinea Xo, PFil, Xo + 8.5, PFil
      PFil = PFil + 0.1
      cPrint.PorteDeLetra = 9
      CodigoB = "ROL DE PAGOS INDIVIDUAL MES DE " & UCase(MesesLetras(Month(FechaFinal)))
      JR = cPrint.anchoTexto(CodigoB)
      cPrint.printTexto Xo + (9 - JR) / 2, PFil, CodigoB
      PFil = PFil + 0.5
     .MoveFirst
      CodigoB = "NINGUNO"
      Do While Not .EOF
         If .fields("Tipo_Rubro") = "PER" Then
             If .fields("Cheq_Dep_Transf") <> Ninguno Then CodigoB = .fields("Cheq_Dep_Transf")
             If .fields("Dias") <> 0 Then I = .fields("Dias")
             If .fields("Horas") <> 0 Then J = .fields("Horas")
         End If
        .MoveNext
      Loop
      cPrint.PorteDeLetra = 8
      cPrint.tipoNegrilla = True
      cPrint.printTexto Xo, PFil, "Beneficiario:"
      cPrint.tipoNegrilla = False
      cPrint.printTexto Xo + 1.7, PFil, NombreCliente
      'MsgBox Xo & vbCrLf & PFil
      PFil = PFil + 0.4
      cPrint.tipoNegrilla = True
      cPrint.printTexto Xo, PFil, "Forma de Pago:"
      cPrint.printTexto Xo + 5, PFil, "Codigo:"
      cPrint.printTexto Xo + 7.9, PFil, "Días:"
      cPrint.tipoNegrilla = False
      cPrint.printTexto Xo + 2.1, PFil, CodigoB
      cPrint.printTexto Xo + 6.05, PFil, CodigoRol
      cPrint.printTexto Xo + 8.65, PFil, Format(I, "#,##0")
      PFil = PFil + 0.35
      cPrint.tipoNegrilla = True
      cPrint.printTexto Xo, PFil, "Desde:"
      cPrint.printTexto Xo + 2.5, PFil, "al:"
      cPrint.printTexto Xo + 5, PFil, "Fecha de Ingreso:"
      cPrint.tipoNegrilla = False
      cPrint.printTexto Xo + 1, PFil, FechaInicial
      cPrint.printTexto Xo + 2.9, PFil, FechaFinal
      cPrint.printTexto Xo + 7.4, PFil, FechaTexto
      PFil = PFil + 0.35
      cPrint.tipoNegrilla = True
      cPrint.printTexto Xo, PFil, "Emails:"
      cPrint.tipoNegrilla = False
      cPrint.printTexto Xo + 1, PFil, Lista_Emails
      PFil = PFil + 0.35
      cPrint.printLinea Xo, PFil, Xo + 8.5, PFil
      PFil = PFil + 0.1
      cPrint.tipoNegrilla = True
      cPrint.printTexto Xo, PFil, "DETALLE DEL EMPLEADO"
      cPrint.printTexto Xo + 5.5, PFil, "INGRESOS"
      cPrint.printTexto Xo + 7.5, PFil, "EGRESOS"
      cPrint.tipoNegrilla = False
      PFil = PFil + 0.3
      cPrint.printLinea Xo, PFil, Xo + 8.5, PFil
      PFil = PFil + 0.1
     .MoveFirst
     'MsgBox .RecordCount & vbCrLf & CodigoRol
      Do While Not .EOF
         'MsgBox .Fields("Tipo_Rubro") & vbCrLf & .Fields("Detalle") & vbCrLf & .Fields("Cod_Rol_Pago")
         cPrint.tipoNegrilla = False
         'MsgBox .Fields("Tipo_Rubro") & vbCrLf & .Fields("Detalle") & vbCrLf & .Fields("Cod_Rol_Pago")
         If .fields("Tipo_Rubro") = "PER" Then
             If .fields("Cod_Rol_Pago") = "Neto_Recibir" Then
                 cPrint.printLinea Xo, PFil, Xo + 8.5, PFil
                 PFil = PFil + 0.1
                 cPrint.tipoNegrilla = True
                 cPrint.printTexto Xo, PFil, ULCase("TOTAL A RECIBIR")
                 cPrint.printLinea Xo, PFil + 0.3, Xo + 8.5, PFil + 0.3
             Else
                 cPrint.printTexto Xo, PFil, UCase(.fields("Detalle"))
             End If
             If .fields("Ingresos") <> 0 Then cPrint.printFields Xo + 4.7, PFil, .fields("Ingresos")
             If .fields("Cod_Rol_Pago") = "Neto_Recibir" And .fields("Egresos") = 0 Then
                 cPrint.printTexto Xo + 8.45, PFil, "0.00"
             Else
                 cPrint.printFields Xo + 6.6, PFil, .fields("Egresos")
             End If
             PFil = PFil + 0.35
         End If
        .MoveNext
      Loop
      PFil = PFil + 0.05
      cPrint.PorteDeLetra = 7
      sSQL = "SELECT " & Full_Fields("Trans_Entrada_Salida") & " " _
           & "FROM Trans_Entrada_Salida " _
           & "WHERE Fecha BETWEEN #" & BuscarFecha(FechaInicial) & "# AND #" & BuscarFecha(FechaFinal) & "# " _
           & "AND Codigo = '" & CodigoRol & "' " _
           & "AND Item = '" & NumEmpresa & "' " _
           & "AND Periodo = '" & Periodo_Contable & "' " _
           & "AND ES = 'R' " _
           & "ORDER BY Fecha,Hora "
      Select_Adodc AdoNovedades, sSQL
      If AdoNovedades.Recordset.RecordCount > 0 Then
         cPrint.tipoNegrilla = True
         cPrint.printTexto Xo, PFil, "OBSERVACIONES:"
         cPrint.tipoNegrilla = False
         PFil = PFil + 0.35
         Do While Not AdoNovedades.Recordset.EOF
            cPrint.printTexto Xo + 0.4, PFil, "=> " & AdoNovedades.Recordset.fields("Tarea")
            PFil = PFil + 0.35
            AdoNovedades.Recordset.MoveNext
         Loop
      End If
      cPrint.printLinea Xo, PFil, Xo + 8.5, PFil
      cPrint.printLinea Xo, PFil, Xo + 8.8, Yo + 9.7
      cPrint.printLinea Xo, 11, Xo + 8.5, 11
      cPrint.printTexto Xo + 1, Yo + 12.6, String(12, "_")
      cPrint.printTexto Xo + 6, Yo + 12.6, String(17, "_")
      cPrint.printTexto Xo + 1.2, Yo + 13, "Empleador"
      cPrint.printTexto Xo + 6.2, Yo + 13, "Recibi conforme"
      DetalleComp = ""
      'If CodigoRol = "0803666460" Then MsgBox CodigoRol
  End If
 End With
End Sub

Public Sub Listar_CxCxP_SubMod()
   Trans_No = 100
   SQL2 = "SELECT Codigo, Beneficiario, Serie, Factura, Valor, Detalle_SubCta, TC, Cta, FECHA_V, SC_No, TM, T_No, Fecha_D, Fecha_H, Bloquear, Item, CodigoU, Prima, DH, Valor_ME, ID " _
        & "FROM Asiento_SC " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND T_No IN (100,101,102) " _
        & "AND CodigoU = '" & CodigoUsuario & "' " _
        & "ORDER BY T_No, Cta, Beneficiario, DH, SC_No "
   Select_Adodc_Grid DGSubCtas, AdoAsientoSC, SQL2
End Sub

Public Sub Encabezado_Rol()
Dim Ancho_Maximo As Single
 PosLinea = 1
 Ancho_Maximo = cPrint.dAnchoPapel - 0.5
 cPrint.printImagen LogoTipo, 1, PosLinea, 4.5, 2
 RutaDestino = RutaSistema & "\LOGOS\DiskCover.gif"
 cPrint.printImagen RutaDestino, Ancho_Maximo - 1.8, PosLinea, 1.8, 0.6
 cPrint.letraTipo TipoHelvetica, 6
 cPrint.tipoNegrilla = True
 cPrint.printTexto Ancho_Maximo - 4.5, PosLinea, "Hora:"
 cPrint.printTexto Ancho_Maximo - 4.5, PosLinea + 0.3, "Pagina No."
 cPrint.printTexto Ancho_Maximo - 4.5, PosLinea + 0.6, "Fecha:"
 cPrint.printTexto Ancho_Maximo - 4.5, PosLinea + 0.9, "Usuario:"
 cPrint.tipoNegrilla = False
 cPrint.printTexto Ancho_Maximo - 3.2, PosLinea, Format(Time, "hh:mm:ss")
 cPrint.printTexto Ancho_Maximo - 3.2, PosLinea + 0.3, Format(Pagina, "0000")
 cPrint.printTexto Ancho_Maximo - 3.2, PosLinea + 0.6, FechaStrgDias(Date)
 cPrint.printTexto Ancho_Maximo - 3.2, PosLinea + 0.9, ULCase(NombreUsuario)
 cPrint.letraTipo TipoTimes
 cPrint.tipoNegrilla = True
 cPrint.PorteDeLetra = 14
 If UCase$(RazonSocial) = UCase$(NombreComercial) Then
    cPrint.printTexto 1, PosLinea, UCase$(RazonSocial), 14, "C", Ancho_Maximo
 Else
    cPrint.printTexto 1, PosLinea, UCase$(RazonSocial), 14, "C", Ancho_Maximo
    cPrint.printTexto 1, PosLinea + 0.5, UCase$(NombreComercial), 14, "C", Ancho_Maximo
 End If
 PosLinea = PosLinea + 0.8
 cPrint.PorteDeLetra = 9
 cPrint.tipoNegrilla = False
 cPrint.printTexto 1, PosLinea, ULCase(Direccion) & ". Teléfono: " & Telefono1, 9, "C", Ancho_Maximo
 PosLinea = PosLinea + 0.45
 cPrint.PorteDeLetra = 12
 cPrint.tipoNegrilla = True
 cPrint.printTexto 1, PosLinea, MensajeEncabData, 12, "C", Ancho_Maximo
 cPrint.tipoNegrilla = False
 cPrint.PorteDeLetra = 8
 Pagina = Pagina + 1
 cPrint.letraTipo TipoHelvetica
 PosLinea = PosLinea + 0.4
End Sub

Public Sub Imprimir_Rol_Colectivo(Datas As Adodc, _
                                  DatasT As Adodc, _
                                  Optional EsCampoCorto As Boolean)
'''On Error GoTo Errorhandler
Dim SizeLetra As Integer
Dim AnchoPict As Single
Dim AltoPict As Single
Dim PosLineaTemp As Single
Dim X_Max As Single
Dim Y_Max As Single
Dim NombFilePict As String
Dim TotValores(30) As Double
Dim CantCamposTemp As Integer

'''Mensajes = "Seguro de Imprimir en:" & vbCrLf & Printer.DeviceName & "?"
'''Titulo = "IMPRESION"
'''Bandera = False
'''SetPrinters.Show 1
Orientacion_Pagina = 2
SetNombrePRN = Impresota_PDF
If PonImpresoraDefecto(SetNombrePRN) Then
   RatonReloj
  'Generamos el documento
   NombFilePict = "Rol Pagos Colectivo " & CAnio & "-" & Format$(NumMeses, "00") & " R-" & RUC & " " & CodigoUsuario
   tPrint.TipoImpresion = Es_PDF
   tPrint.NombreArchivo = NombFilePict
   tPrint.TituloArchivo = "Rol de Pagos Colectivo " & CAnio & "-" & Format(NumMeses, "00") & " " & RUC
   tPrint.TipoLetra = TipoHelvetica
   tPrint.OrientacionPagina = Orientacion_Pagina
   tPrint.PaginaA4 = True
   tPrint.EsCampoCorto = False
   tPrint.VerDocumento = False
   
   Set cPrint = New cImpresion
   cPrint.iniciaImpresion
   cPrint.anchoRegistro InicioX, Datas, EsCampoCorto

InicioX = 0.5
InicioY = 0
SizeLetra = 6
'DataAnchoCampos InicioX, Datas, SizeLetra, TipoArialNarrow, Orientacion_Pagina, EsCampoCorto
Volver_Imprimir_Provision:

 Ancho(0) = 1    'No.
 Ancho(1) = 1.5  'CI
 Ancho(2) = 2.9  'Empleado
 Ancho(3) = 5.4  'Grupo_No
 Ancho(4) = 6.7  'Días
 Ancho(5) = 7.4  'Fecha_Ing
 Ancho(6) = 8.8  'FR
 Ancho(7) = 9.3  'Horas
 Ancho(8) = 10.3 'Cheque_No
 
 Pagina = 1
 PosLinea = 1
'Iniciamos la impresion
 cPrint.tipoNegrilla = False
 With Datas.Recordset
  If .RecordCount > 0 Then
     .MoveFirst
     'Ancho(6) = Salario
      Distancia = Ancho(8) + 1.4
      J = CantCampos
      For I = 9 To CantCampos - 1
          If .fields(I).Name = "II" Then J = I
      Next I
      CantCampos = J
      For I = 9 To CantCampos - 1
          If .fields(I).Name = "I" Or .fields(I).Name = "II" Then
              Ancho(I) = Distancia
              Distancia = Distancia + 0.02
          ElseIf .fields(I).Name = "Firma" Then
              Ancho(I) = Distancia
              Distancia = Distancia + 5
          ElseIf .fields(I).Name = "CodigoU" Then
              Ancho(I) = Distancia
              Distancia = Distancia + 7
          Else
              Ancho(I) = Distancia
              Distancia = Distancia + 1
          End If
          Ancho(CantCampos) = Ancho(I) + 0.1
      Next I
      X_Max = Ancho(CantCampos)

     'MsgBox Y_Max & " .... " & X_Max
      LimiteAlto = cPrint.dAltoPapel - 0.5
     'MsgBox LimiteAlto
      Ancho(CantCampos) = cPrint.dAnchoPapel - 1
      MensajeEncabData = "ROL DE PAGOS COLECTIVO CORRESPONDIENTE AL MES DE " & MesesLetras(Month(FechaInicial), True)
      Encabezado_Rol
      PosLineaTemp = PosLinea
      cPrint.PorteDeLetra = SizeLetra
      cPrint.tipoNegrilla = True
      For I = 0 To CantCampos - 1
          Cadena = Replace(.fields(I).Name, "_", " ")
          Select Case Cadena
            Case "Codigo", "CodigoU", "I", "II": 'cPrint.printTexto  Ancho(I) + Distancia, PosLinea, ""
            Case Else: PosLinea = cPrint.printTextoMultiple(Ancho(I), PosLineaTemp, Cadena, 1)
          End Select
      Next I
      PosLinea = PosLineaTemp
      PosLinea = PosLinea + 0.7
      PosLineaTemp = PosLinea
      cPrint.printLinea 0.5, PosLinea, 28.5, PosLinea
      PosLinea = PosLinea + 0.05
      SizeLetra = 5
      cPrint.tipoNegrilla = False
      cPrint.PorteDeLetra = SizeLetra
      Do While Not .EOF
         cPrint.tipoNegrilla = True
         For I = 0 To CantCampos - 1
             Distancia = cPrint.anchoFields(.fields(I), 2)
             If I = 0 Then
                cPrint.dStrgFormatoCampo = Format(Val(cPrint.dStrgFormatoCampo), "00")
                Distancia = 0
                'MsgBox I & ": " & cPrint.dStrgFormatoCampo
             End If
             'Distancia = CampoWidth(.Fields(I))
             If cPrint.dStrgFormatoCampo = Ninguno Then
                cPrint.dStrgFormatoCampo = " "
             ElseIf cPrint.dStrgFormatoCampo = "0" Or cPrint.dStrgFormatoCampo = "0.00" Then
                cPrint.dStrgFormatoCampo = " "
             End If
             Select Case .fields(I).Name
               Case "Codigo", "CodigoU", "I", "II": ' cPrint.printTexto  Ancho(I) + Distancia, PosLinea, ""
               Case "Nombre_Empleado"
                    cPrint.printTexto Ancho(I) + Distancia, PosLinea, Extraer_Apellidos(cPrint.dStrgFormatoCampo)
                    cPrint.printTexto Ancho(I) + Distancia, PosLinea + 0.3, Extraer_Nombres(cPrint.dStrgFormatoCampo)
               Case Else: cPrint.printTexto Ancho(I) + Distancia, PosLinea, cPrint.dStrgFormatoCampo
             End Select
         Next I
         PosLinea = PosLinea + 0.7
         cPrint.printLinea 0.5, PosLinea, 28.5, PosLinea 'Ancho(CantCampos) - 0.1
         PosLinea = PosLinea + 0.05
         If PosLinea >= LimiteAlto - 0.6 Then
'''            For I = 0 To CantCampos - 1
'''             Select Case .Fields(I).Name
'''               Case "Codigo", "CodigoU", "I", "II": ' cPrint.printTexto  Ancho(I) + Distancia, PosLinea, ""
'''               Case Else: cPrint.printLinea Ancho(I), 2.8, Ancho(I), PosLinea
'''             End Select
'''            Next I
        
            cPrint.printLinea 0.5, PosLinea, 28.5, PosLinea
            cPrint.paginaNueva
            Encabezado_Rol
            SizeLetra = 6
            PosLineaTemp = PosLinea
            cPrint.PorteDeLetra = SizeLetra
            cPrint.tipoNegrilla = True
            For I = 0 To CantCampos - 1
                Cadena = Replace(.fields(I).Name, "_", " ")
                Select Case Cadena
                  Case "Codigo", "CodigoU", "I", "II": 'cPrint.printTexto  Ancho(I) + Distancia, PosLinea, ""
                  Case Else: PosLinea = cPrint.printTextoMultiple(Ancho(I), PosLineaTemp, Cadena, 1)
                End Select
            Next I
            PosLinea = PosLineaTemp
            PosLinea = PosLinea + 0.7
            cPrint.printLinea 0.5, PosLinea, 28.5, PosLinea
            PosLinea = PosLinea + 0.1
            cPrint.tipoNegrilla = False
            SizeLetra = 5
         End If
        .MoveNext
      Loop
'''      For I = 0 To CantCampos - 1
'''       Select Case .Fields(I).Name
'''         Case "Codigo", "CodigoU", "I", "II": ' cPrint.printTexto  Ancho(I) + Distancia, PosLinea, ""
'''         Case Else: cPrint.printLinea Ancho(I) - 0.5, PosLineaTemp, Ancho(I) - 0.5, PosLinea
'''       End Select
'''      Next I
  End If
 End With
 cPrint.printLinea 28.5, PosLineaTemp, 28.5, PosLinea
 PosLinea = PosLinea + 0.5
 
'Resumen de los pagos por grupos
 With DatasT.Recordset
  If .RecordCount > 0 Then
     .MoveFirst
      If PosLinea >= LimiteAlto - 0.6 Then
         cPrint.paginaNueva
         Encabezado_Rol
      End If
      cPrint.anchoRegistro InicioX, DatasT, EsCampoCorto
      Cadena = ""
      Ancho(0) = 1
      Ancho(1) = 3
      InicioX = 5.25
      For I = 2 To CantCampos - 1
          Ancho(I) = InicioX
          InicioX = InicioX + 1.5
      Next I
      Ancho(I) = InicioX
      cPrint.PorteDeLetra = SizeLetra
      cPrint.tipoNegrilla = True
      For I = 0 To CantCampos - 1
          cPrint.printTexto Ancho(I), PosLinea, .fields(I).Name
          TotValores(I) = 0
      Next I
      PosLinea = PosLinea + 0.35
      PosLineaTemp = PosLinea
      cPrint.printLinea 0.5, PosLinea, Ancho(CantCampos) - 0.5, PosLinea
      PosLinea = PosLinea + 0.05
      cPrint.tipoNegrilla = False
      
      Do While Not .EOF
         For I = 0 To CantCampos - 1
             Distancia = cPrint.anchoFields(.fields(I), 2)
             If cPrint.dStrgFormatoCampo = Ninguno Then
                cPrint.dStrgFormatoCampo = " "
             ElseIf cPrint.dStrgFormatoCampo = "0" Or cPrint.dStrgFormatoCampo = "0.00" Then
                cPrint.dStrgFormatoCampo = " "
             End If
             cPrint.printTexto Ancho(I) + Distancia, PosLinea, cPrint.dStrgFormatoCampo
             If I <> 0 Then TotValores(I) = TotValores(I) + .fields(I)
         Next I
         PosLinea = PosLinea + 0.4
         cPrint.printLinea 0.5, PosLinea, Ancho(CantCampos) - 0.5, PosLinea
         PosLinea = PosLinea + 0.05
         If PosLinea >= LimiteAlto Then
            For I = 0 To CantCampos
                cPrint.printLinea Ancho(I) - 0.7, PosLineaTemp, Ancho(I) - 0.7, PosLinea + 0.1
            Next I
            
            cPrint.paginaNueva
            Encabezado_Rol
            PosLineaTemp = PosLinea + 0.2
            cPrint.PorteDeLetra = SizeLetra
            cPrint.tipoNegrilla = True
            For I = 0 To CantCampos - 1
                cPrint.printTexto Ancho(I), PosLinea, .fields(I).Name
            Next I
            PosLinea = PosLinea + 0.3
            cPrint.printLinea 0.5, PosLinea + 0.1, Ancho(CantCampos) - 0.5, PosLinea + 0.1
            PosLinea = PosLinea + 0.1
            cPrint.tipoNegrilla = False
         End If
        .MoveNext
      Loop
      For I = 0 To CantCampos
          cPrint.printLinea Ancho(I) - 0.7, PosLineaTemp, Ancho(I) - 0.7, PosLinea
      Next I
  End If
 End With
 cPrint.tipoNegrilla = True
 cPrint.printTexto Ancho(0), PosLinea, "T O T A L E S"
 For I = 1 To CantCampos - 1
     cPrint.printVariable Ancho(I), PosLinea, TotValores(I)
 Next I
 cPrint.PorteDeLetra = SizeLetra
 CantCamposTemp = CantCampos
 sSQL = "SELECT C.Cliente,TR.* " _
      & "FROM Trans_Entrada_Salida As TR,Clientes As C " _
      & "WHERE TR.Fecha BETWEEN #" & BuscarFecha(FechaInicial) & "# AND #" & BuscarFecha(FechaFinal) & "# " _
      & "AND TR.Codigo = C.Codigo " _
      & "AND TR.Item = '" & NumEmpresa & "' " _
      & "AND TR.Periodo = '" & Periodo_Contable & "' " _
      & "AND TR.ES = 'R' " _
      & "ORDER BY C.Cliente,TR.Fecha,TR.Hora "
 Select_Adodc AdoNovedades, sSQL
 With AdoNovedades.Recordset
  If .RecordCount > 0 Then
      cPrint.tipoNegrilla = True
      PosLinea = PosLinea + 0.4
      cPrint.printTexto Ancho(0), PosLinea, "OBSERVACIONES:"
      cPrint.tipoNegrilla = False
      PosLinea = PosLinea + 0.35
      Do While Not .EOF
         cPrint.printTexto Ancho(0), PosLinea, .fields("Tarea")
         PosLinea = PosLinea + 0.35
        .MoveNext
      Loop
  End If
  End With
  cPrint.tipoNegrilla = False
  CantCampos = CantCamposTemp
  RatonNormal
  MensajeEncabData = ""
 
 'fin del documento
  cPrint.finalizaImpresion
 'XXXYYYZZZ
  'Presentar_PDF RPPDF, RutaDocumentoPDF
  
'''Exit Sub
'''Errorhandler:
'''    RatonNormal
'''    ErrorDeImpresion
'''    Exit Sub
'''Else
'''    RatonNormal
End If
End Sub

Public Sub Inicializar_Cero_Asientos()
   'Inicializamos los Asientos de submodulos
    Trans_No = 102
    SQL2 = "SELECT * " _
         & "FROM Asiento " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND T_No = " & Trans_No & " " _
         & "AND CodigoU = '" & CodigoUsuario & "' "
    Select_Adodc_Grid DGAsiento(2), AdoAsiento2, SQL2
    
    Trans_No = 101
    SQL2 = "SELECT * " _
         & "FROM Asiento " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND T_No = " & Trans_No & " " _
         & "AND CodigoU = '" & CodigoUsuario & "' "
    Select_Adodc_Grid DGAsiento(1), AdoAsiento1, SQL2
        
    Trans_No = 100
    SQL2 = "SELECT * " _
         & "FROM Asiento " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND T_No = " & Trans_No & " " _
         & "AND CodigoU = '" & CodigoUsuario & "' "
    Select_Adodc_Grid DGAsiento(0), AdoAsiento, SQL2
End Sub

'''Public Sub In_Ctas_Del_Rol(ExisteLaCta As String)
'''  If InStr(CtasDelRol, ExisteLaCta) = 0 Then CtasDelRol = CtasDelRol & "'" & ExisteLaCta & "',"
'''End Sub

Private Sub WebBPDF_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
   On Error Resume Next
   Set oDocument = pDisp.Document

'   MsgBox "File opened by: " & oDocument.Application.Name
End Sub

