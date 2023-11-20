VERSION 5.00
Object = "{05BFD3F1-6319-4F30-B752-C7A22889BCC4}#1.0#0"; "AcroPDF.dll"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.4#0"; "comctl32.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Begin VB.Form LRolPagos 
   Caption         =   "Catalogo de Rol de Pagos"
   ClientHeight    =   8160
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11280
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8160
   ScaleWidth      =   11280
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11280
      _ExtentX        =   19897
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
            Size            =   8,25
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
               Size            =   8,25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2625
            TabIndex        =   5
            Text            =   "Combo1"
            Top             =   210
            Width           =   1485
         End
         Begin VB.ComboBox CAnio 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8,25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   735
            TabIndex        =   3
            Text            =   "Combo1"
            Top             =   210
            Width           =   1065
         End
         Begin VB.ComboBox CmbGrupos 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8,25
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
               Size            =   8,25
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
               Size            =   8,25
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
               Size            =   8,25
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
         Size            =   8,25
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
            Size            =   8,25
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
            Size            =   8,25
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
      Height          =   345
      Left            =   7140
      TabIndex        =   16
      Top             =   1155
      Width           =   4845
      _ExtentX        =   8546
      _ExtentY        =   609
      _Version        =   393216
      Text            =   "CxP"
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
   Begin VB.CheckBox CheqCxP 
      Caption         =   "Generar Nomina sin alcance de efectivo (CxP Empleados)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8,25
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
         Size            =   8,25
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
         Size            =   8,25
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
         Size            =   8,25
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
      Height          =   6525
      Left            =   105
      TabIndex        =   19
      Top             =   1575
      Width           =   11565
      _ExtentX        =   20399
      _ExtentY        =   11509
      _Version        =   393216
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      BackColor       =   -2147483637
      TabCaption(0)   =   "ROL INDIVIDUAL"
      TabPicture(0)   =   "LRolPag.frx":0017
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fPDF"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "ROL DE PAGOS"
      TabPicture(1)   =   "LRolPag.frx":0033
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DGNominaProv"
      Tab(1).Control(1)=   "AdoNomina"
      Tab(1).Control(2)=   "DGTotNomina"
      Tab(1).Control(3)=   "DGNomina"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "CxC/CxP Empleados"
      TabPicture(2)   =   "LRolPag.frx":004F
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "DGNomina1"
      Tab(2).Control(1)=   "DGSubCtas"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "CONTABILIDAD"
      TabPicture(3)   =   "LRolPag.frx":006B
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "DGAsiento(0)"
      Tab(3).Control(1)=   "LabelDiferencia"
      Tab(3).Control(2)=   "LabelHaber"
      Tab(3).Control(3)=   "LabelDebe"
      Tab(3).Control(4)=   "Label19"
      Tab(3).Control(5)=   "Label1"
      Tab(3).Control(6)=   "LblConcepto(0)"
      Tab(3).ControlCount=   7
      TabCaption(4)   =   "PROVISIONES"
      TabPicture(4)   =   "LRolPag.frx":0087
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "DGAsiento(1)"
      Tab(4).Control(1)=   "DGAsiento(2)"
      Tab(4).Control(2)=   "LblConcepto(1)"
      Tab(4).Control(3)=   "LblConcepto(2)"
      Tab(4).ControlCount=   4
      Begin AcroPDFLibCtl.AcroPDF fPDF 
         Height          =   750
         Left            =   105
         TabIndex        =   36
         Top             =   420
         Width           =   4320
         _cx             =   5080
         _cy             =   5080
      End
      Begin MSDataGridLib.DataGrid DGNomina 
         Bindings        =   "LRolPag.frx":00A3
         Height          =   1905
         Left            =   -74895
         TabIndex        =   20
         ToolTipText     =   "<Ctrl + F9>: Comisiones y el I.E.S.S."
         Top             =   420
         Width           =   10305
         _ExtentX        =   18177
         _ExtentY        =   3360
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8,25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8,25
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
         Bindings        =   "LRolPag.frx":00BB
         Height          =   1065
         Left            =   -74895
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
            Size            =   8,25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8,25
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
         Left            =   -74895
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
            Size            =   8,25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSDataGridLib.DataGrid DGNomina1 
         Bindings        =   "LRolPag.frx":00D6
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
            Size            =   8,25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8,25
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
         Bindings        =   "LRolPag.frx":00EF
         Height          =   4635
         Index           =   0
         Left            =   -74895
         TabIndex        =   23
         Top             =   735
         Width           =   10305
         _ExtentX        =   18177
         _ExtentY        =   8176
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8,25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8,25
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
         Bindings        =   "LRolPag.frx":0108
         Height          =   1905
         Index           =   1
         Left            =   -74895
         TabIndex        =   30
         Top             =   735
         Width           =   10305
         _ExtentX        =   18177
         _ExtentY        =   3360
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8,25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8,25
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
         Bindings        =   "LRolPag.frx":0122
         Height          =   2850
         Index           =   2
         Left            =   -74895
         TabIndex        =   31
         Top             =   2940
         Width           =   10305
         _ExtentX        =   18177
         _ExtentY        =   5027
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8,25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8,25
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
      Begin MSDataGridLib.DataGrid DGSubCtas 
         Bindings        =   "LRolPag.frx":013C
         Height          =   2640
         Left            =   -74895
         TabIndex        =   34
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
            Size            =   8,25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8,25
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
      Begin MSDataGridLib.DataGrid DGNominaProv 
         Bindings        =   "LRolPag.frx":0157
         Height          =   1905
         Left            =   -74895
         TabIndex        =   35
         Top             =   2415
         Width           =   10305
         _ExtentX        =   18177
         _ExtentY        =   3360
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8,25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8,25
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
      Begin VB.Label LabelDiferencia 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8,25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   330
         Left            =   -71010
         TabIndex        =   27
         Top             =   5445
         Width           =   1695
      End
      Begin VB.Label LabelHaber 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8,25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   330
         Left            =   -66390
         TabIndex        =   25
         Top             =   5445
         Width           =   1800
      End
      Begin VB.Label LabelDebe 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8,25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   330
         Left            =   -68175
         TabIndex        =   26
         Top             =   5445
         Width           =   1800
      End
      Begin VB.Label Label19 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Totales"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8,25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -69225
         TabIndex        =   29
         Top             =   5445
         Width           =   1065
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Diferencia"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8,25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -72165
         TabIndex        =   28
         Top             =   5445
         Width           =   1170
      End
      Begin VB.Label LblConcepto 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8,25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   -74895
         TabIndex        =   33
         Top             =   420
         Width           =   10305
      End
      Begin VB.Label LblConcepto 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8,25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   2
         Left            =   -74895
         TabIndex        =   32
         Top             =   2625
         Width           =   10305
      End
      Begin VB.Label LblConcepto 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8,25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   -74895
         TabIndex        =   24
         Top             =   420
         Width           =   10305
      End
   End
   Begin VB.TextBox TxtCheque 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8,25
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
         Size            =   8,25
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
         Size            =   8,25
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
         Size            =   8,25
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
         Size            =   8,25
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
         Size            =   8,25
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
         Size            =   8,25
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
         Size            =   8,25
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
         Size            =   8,25
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
         Size            =   8,25
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
         Size            =   8,25
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
         Size            =   8,25
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
         Size            =   8,25
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
         Size            =   8,25
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
         Size            =   8,25
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
         Size            =   8,25
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
         Size            =   8,25
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
         Size            =   8,25
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
      Caption         =   "Novedades"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8,25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   11865
      Top             =   2940
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
            Picture         =   "LRolPag.frx":0173
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LRolPag.frx":048D
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LRolPag.frx":07A7
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LRolPag.frx":0AC1
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LRolPag.frx":0DDB
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LRolPag.frx":10F5
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LRolPag.frx":140F
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LRolPag.frx":1729
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LRolPag.frx":1A43
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LRolPag.frx":4B4D
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LRolPag.frx":4E67
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LRolPag.frx":5181
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LRolPag.frx":53AB
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LRolPag.frx":55E9
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LRolPag.frx":5827
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LRolPag.frx":5A51
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LRolPag.frx":5C7B
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LRolPag.frx":5F95
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8,25
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
         Size            =   8,25
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

Dim ErrorInventario As String
Dim CtasRol() As CtasAsiento
Dim CtasPro() As CtasAsiento
Dim CtasPat() As CtasAsiento
Dim CantCtas As Long
Dim Rubros_Otros_Ingresos As String
Dim Lista_Emails As String

Dim TRol_Pago As Tipo_Rol_Pago_Individual

Public Sub Ctas_Asientos_Rol()
  RatonReloj
  sSQL = "SELECT Grupo_Rol,Cta_Diferencia,Cta_Vacacion,Cta_Sueldo,Cta_Horas_Ext," _
       & "Cta_Aporte_Patronal_G,Cta_Decimo_Cuarto_G,Cta_Decimo_Cuarto_P,Cta_Decimo_Tercer_P," _
       & "Cta_Fondo_Reserva_G,Cta_Fondo_Reserva_P,Cta_Vacaciones_G,Cta_Vacaciones_P," _
       & "Cta_IESS_Personal,Cta_Quincena,Cta_Decimo_Tercer_G," _
       & "Cta_IESS_Patronal,Cta_Antiguedad,Cta_Forma_Pago " _
       & "FROM Catalogo_Rol_Pagos " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "GROUP BY Grupo_Rol,Cta_Diferencia,Cta_Vacacion,Cta_Sueldo,Cta_Horas_Ext," _
       & "Cta_Aporte_Patronal_G,Cta_Decimo_Cuarto_G,Cta_Decimo_Cuarto_P,Cta_Decimo_Tercer_P," _
       & "Cta_Fondo_Reserva_G,Cta_Fondo_Reserva_P,Cta_Vacaciones_G,Cta_Vacaciones_P," _
       & "Cta_IESS_Personal,Cta_Quincena,Cta_Decimo_Tercer_G," _
       & "Cta_IESS_Patronal,Cta_Antiguedad,Cta_Forma_Pago " _
       & "ORDER BY Grupo_Rol,Cta_Diferencia,Cta_Vacacion,Cta_Sueldo,Cta_Horas_Ext," _
       & "Cta_Aporte_Patronal_G,Cta_Decimo_Cuarto_G,Cta_Decimo_Cuarto_P,Cta_Decimo_Tercer_P," _
       & "Cta_Fondo_Reserva_G,Cta_Fondo_Reserva_P,Cta_Vacaciones_G,Cta_Vacaciones_P," _
       & "Cta_IESS_Personal,Cta_Quincena,Cta_Decimo_Tercer_G," _
       & "Cta_IESS_Patronal,Cta_Antiguedad,Cta_Forma_Pago "
  SelectAdodc AdoAux, sSQL
  With AdoAux.Recordset
   If .RecordCount > 0 Then
       CantCtas = (.RecordCount * 16) + 10
       ReDim CtasRol(CantCtas) As CtasAsiento
       ReDim CtasPro(CantCtas) As CtasAsiento
       ReDim CtasPat(CantCtas) As CtasAsiento
       For IE = 0 To CantCtas - 1
           CtasRol(IE).Cta = "0"
           CtasRol(IE).Valor = 0
           CtasPro(IE).Cta = "0"
           CtasPro(IE).Valor = 0
           CtasPat(IE).Cta = "0"
           CtasPat(IE).Valor = 0
       Next IE
      'Seteamos las Cuentas del Rol Pagos
       Do While Not .EOF
          NivelNo = .Fields("Grupo_Rol")
          
         'Para el Rol
          SetearCtasCierreRol .Fields("Cta_Diferencia")
          SetearCtasCierreRol .Fields("Cta_Vacacion")
          SetearCtasCierreRol .Fields("Cta_Sueldo")
          SetearCtasCierreRol .Fields("Cta_Horas_Ext")
          SetearCtasCierreRol .Fields("Cta_Antiguedad")
          SetearCtasCierreRol .Fields("Cta_IESS_Personal")
          SetearCtasCierreRol .Fields("Cta_Quincena")
          SetearCtasCierreRol .Fields("Cta_Forma_Pago")
          SetearCtasCierreRol .Fields("Cta_Fondo_Reserva_G")
          SetearCtasCierreRol .Fields("Cta_Decimo_Tercer_G")
          SetearCtasCierreRol .Fields("Cta_Decimo_Cuarto_G")
          
         'Aporte Patronal
          SetearCtasCierrePat .Fields("Cta_Aporte_Patronal_G")
          SetearCtasCierrePat .Fields("Cta_IESS_Patronal")
          
         'Provisiones de Decimos y Fondo de Reserva
          SetearCtasCierrePro .Fields("Cta_Decimo_Cuarto_G")
          SetearCtasCierrePro .Fields("Cta_Decimo_Cuarto_P")
          SetearCtasCierrePro .Fields("Cta_Decimo_Tercer_G")
          SetearCtasCierrePro .Fields("Cta_Decimo_Tercer_P")
          SetearCtasCierrePro .Fields("Cta_Fondo_Reserva_G")
          SetearCtasCierrePro .Fields("Cta_Fondo_Reserva_P")
          SetearCtasCierrePro .Fields("Cta_Vacaciones_G")
          SetearCtasCierrePro .Fields("Cta_Vacaciones_P")
         .MoveNext
       Loop
   End If
  End With
  NivelNo = "Rubros Adicionales"
  sSQL = "SELECT Cta " _
       & "FROM Catalogo_Rol_Rubros " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND CPais = '" & CodigoPais & "' " _
       & "GROUP BY Cta "
  SelectAdodc AdoAux, sSQL
  With AdoAux.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          SetearCtasCierreRol .Fields("Cta")
         .MoveNext
       Loop
   End If
  End With
  
 'SubModulos
  NivelNo = "Submódulos"
  sSQL = "SELECT Cta " _
       & "FROM Asiento_SC " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND TC IN ('C','P') " _
       & "GROUP BY Cta "
  SelectAdodc AdoAux, sSQL
  With AdoAux.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          SetearCtasCierreRol .Fields("Cta")
         .MoveNext
       Loop
   End If
  End With
 'Activar la Cuenta de Impuesto a la Renta
  NivelNo = "Seteos"
  SetearCtasCierreRol Cta_Impuesto_Renta_Empleado
  RatonNormal
 End Sub

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

'''Public Sub EncabezadoRolPagos(Datas As Adodc)
'''Dim AuxPosLinea As Single
'''Dim InicX As Single
'''Dim InicY As Single
'''Dim Y0 As Single
'''Dim y1 As Single
'''Dim X0 As Single
'''Dim x1 As Single
'''   LimiteAncho = 19
'''   X0 = 0.1: x1 = LimiteAncho
'''   Y0 = PosLinea
'''   Ancho(CantCampos) = LimiteAncho
'''   PorteLetra = cPrint.porteDeLetra
'''   LetraAnterior = cPrint.tipoDeLetra
'''   cPrint.tipoDeLetra = TipoTimes
'''
'''   cPrint.porteDeLetra = 8
'''   If X0 <= 0 Then X0 = 0.1
'''   If Y0 <= 0 Then X0 = 0.1
'''   If Printer.Orientation = 2 Then
'''      Y0 = 1: y1 = 2.6
'''   End If
'''   If x1 > LimiteAncho Then x1 = LimiteAncho - 0.1
'''   PrinterPaint LogoTipo, X0, Y0, 3, 1.5
'''   Printer.FontBold = True: cPrint.porteDeLetra = 18: Printer.FontItalic = True
'''   Printer.CurrentX = CentrarTextoEncab(Empresa, X0, x1)
'''   Printer.CurrentY = PosLinea
'''   Printer.Print Empresa
'''   cPrint.porteDeLetra = 9: Printer.FontItalic = False
'''   PrinterTexto 17, PosLinea, "No. " & Format(Pagina, "0000")
'''   cPrint.porteDeLetra = 8
'''   PosLinea = PosLinea + 0.7
'''   Cadena = "R.U.C. " & RUC & " - " & Direccion & ". Teléfono: " & Telefono1 & "."
'''   Printer.CurrentX = CentrarTextoEncab(Cadena, X0, x1)
'''   Printer.CurrentY = PosLinea
'''   Printer.Print Cadena
'''   PosLinea = PosLinea + 0.5
'''   Printer.FontBold = False: cPrint.porteDeLetra = 10
'''   Printer.FontName = LetraAnterior
'''   Printer.FontName = TipoTimes
'''   Printer.FontBold = True
'''If SQLMsg1 <> "" Then
'''   cPrint.porteDeLetra = 12
'''   PrinterTexto CentrarTexto(SQLMsg1, Ancho(CantCampos)), PosLinea, SQLMsg1
'''   PosLinea = PosLinea + 0.7
'''End If
'''cPrint.porteDeLetra = 10
'''If SQLMsg2 <> "" Then
'''   PrinterTexto CentrarTexto(SQLMsg2, Ancho(CantCampos)), PosLinea, SQLMsg2
'''   PosLinea = PosLinea + 0.6
'''End If
'''If SQLMsg3 <> "" Then
'''   PrinterTexto Ancho(0), PosLinea, SQLMsg3
'''   PosLinea = PosLinea + 0.6
'''End If
'''cPrint.porteDeLetra = 9
''''========================================================================
'''Imprimir_Linea_H PosLinea, Ancho(0), LimiteAncho
'''PosLinea = PosLinea + 0.2
'''AuxPosLinea = PosLinea
'''With Datas.Recordset
''' If .RecordCount > 0 Then
'''     Printer.FontBold = True
'''     Codigo = .Fields("Codigo")
'''     PrinterTexto 1, PosLinea, "FECHA DE INGRESO:"
'''     PrinterTexto 11, PosLinea, "CUSP No."
'''     PosLinea = PosLinea + 0.4
'''     PrinterTexto 1, PosLinea, "BENEFICIARIO:"
'''     PrinterTexto 11, PosLinea, "D.I. No."
'''     PosLinea = PosLinea + 0.4
'''     PrinterTexto 1, PosLinea, "CODIGO:"
'''     PrinterTexto 5, PosLinea, "CARGO:"
'''     PrinterTexto 11, PosLinea, "TOTAL HORAS:"
'''     PosLinea = PosLinea + 0.4
'''     PrinterTexto 1, PosLinea, "PERIODO: "
'''     PrinterTexto 11, PosLinea, "VACACIONES DESDE:"
'''     PrinterTexto 17, PosLinea, "AL:"
'''     PosLinea = PosLinea + 0.5
'''     Printer.FontBold = False
'''     PosLinea = AuxPosLinea
'''     PrinterTexto 4.5, PosLinea, FechaTexto
'''     PrinterTexto 13, PosLinea, No_Personal
'''     PosLinea = PosLinea + 0.4
'''     PrinterFields 3.8, PosLinea, .Fields("Cliente")
'''     PrinterTexto 13, PosLinea, CICliente
'''     PosLinea = PosLinea + 0.4
'''     PrinterFields 3, PosLinea, .Fields("Codigo")
'''     PrinterTexto 6.7, PosLinea, CxC_Clientes
'''     PrinterFields 13, PosLinea, .Fields("Hora_Trab")
'''     PosLinea = PosLinea + 0.4
'''     PrinterTexto 3, PosLinea, UCase(MesesLetras(Month(FechaFinal)))
'''     If NoMeses = Month(FechaInicial) Then
'''         PrinterTexto 14.5, PosLinea, FechaInicial
'''         PrinterTexto 18, PosLinea, FechaFinal
'''     End If
''' End If
'''End With
'''PosLinea = PosLinea + 0.5
'''Imprimir_Linea_H PosLinea, Ancho(0), LimiteAncho
'''PosLinea = PosLinea + 0.2
'''Printer.FontBold = False
'''cPrint.porteDeLetra = PorteLetra
'''Printer.FontName = LetraAnterior
'''End Sub

'Inserta los SubModulos de CxC o CxP
Public Sub InsertarCxCxP(CodigoClient As String, _
                         CtaProc As String, _
                         Valor As Currency, _
                         TipoDeCta As String)
  ' MsgBox "...."
  If Len(CodigoClient) > 1 And Len(CtaProc) > 1 And Valor > 0 Then
     If LnSC_No < 0 Then LnSC_No = 0
     SetAdoAddNew "Asiento_SC"
     SetAdoFields "Codigo", CodigoClient
     SetAdoFields "Beneficiario", NombreCliente
     SetAdoFields "Cta", CtaProc
     SetAdoFields "Valor", Valor
     SetAdoFields "FECHA_V", FechaFinal
     SetAdoFields "Factura", Factura_No
     SetAdoFields "TC", TipoDeCta
     Select Case TipoDeCta
       Case "C": SetAdoFields "DH", "2"
       Case "P": SetAdoFields "DH", "1"
       Case "G": SetAdoFields "DH", "1"
     End Select
     SetAdoFields "TM", "1"
     SetAdoFields "Item", NumEmpresa
     SetAdoFields "T_No", Trans_No
     SetAdoFields "SC_No", LnSC_No
     SetAdoFields "CodigoU", CodigoUsuario
     SetAdoUpdate
     LnSC_No = LnSC_No + 1
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
Dim ContLineas As Integer
Dim Es_Vacaciones As Boolean
Dim Aporte_Patronal As Single
Dim NombFilePict As String

'Certificados Acumulados
 sSQL = "SELECT Codigo, SUM(Creditos) As Total_Certif " _
      & "FROM Trans_SubCtas " _
      & "WHERE Cta = '" & Cta_Dcts_Certif & "' " _
      & "AND Codigo = '" & CodigoRol & "' " _
      & "AND Item = '" & NumEmpresa & "' " _
      & "AND Periodo = '" & Periodo_Contable & "' " _
      & "GROUP BY Codigo " _
      & "ORDER BY Codigo "
 SelectAdodc AdoCertificado, sSQL
'Datos del Encabezadodel Rol Individual
 No_Personal = Ninguno
 FechaTexto = Ninguno
 CICliente = Ninguno
 NomCtaSup = Ninguno
 NumCheque = Ninguno
 CodigoB = "OTROS"
 With AdoClientes.Recordset
  If .RecordCount > 0 Then
     .MoveFirst
     .Find ("Codigo = '" & CodigoRol & "' ")
      If Not .EOF Then
         NombreCliente = .Fields("Cliente")
         No_Personal = .Fields("No_Personal")
         FechaTexto = .Fields("Fecha")
         CICliente = .Fields("CI_RUC")
         CxC_Clientes = .Fields("Actividad")
         FechaInicial = .Fields("FechaVI")
         FechaFinal = .Fields("FechaVF")
         NoMeses = .Fields("Mes")
         IEESS_Per = .Fields("IEESS_Per")
         IEESS_Pat = .Fields("IEESS_Pat")
         IEESS_Ext = .Fields("IEESS_ExtC")
         NomCtaSup = .Fields("Cta_Transferencia")
         Cta_IESS = .Fields("Cta_IESS_Personal")
         TextoBanco = Ninguno
        'Enviamos lista de mails
         Lista_Emails = ""
         If Len(.Fields("Email")) > 1 Then Lista_Emails = Trim$(.Fields("Email")) & ";"
         If .Fields("Email") <> .Fields("Email2") And Len(.Fields("Email2")) > 1 Then
             Lista_Emails = Lista_Emails & Trim$(.Fields("Email2")) & ";"
         End If
         sSQL = "SELECT * " _
              & "FROM Tabla_Referenciales_SRI " _
              & "WHERE Codigo = '" & .Fields("Codigo_Banco") & "' " _
              & "AND Tipo_Referencia = 'BANCOS Y COOP' "
        SelectAdodc AdoAux, sSQL
        If AdoAux.Recordset.RecordCount > 0 Then TextoBanco = ULCase(AdoAux.Recordset.Fields("Descripcion"))
      End If
  End If
 End With
 
 Fecha_Del_AT CMes, CAnio
'Presentamos el rol individual del empleado
 SQL2 = "SELECT * " _
      & "FROM Trans_Rol_de_Pagos " _
      & "WHERE Fecha_D >= #" & FechaIni & "# " _
      & "AND Fecha_H <= #" & FechaFin & "# " _
      & "AND Item = '" & NumEmpresa & "' " _
      & "AND Periodo = '" & Periodo_Contable & "' " _
      & "AND Codigo = '" & CodigoRol & "' "
 If CmbGrupos <> "TODOS" Then SQL2 = SQL2 & "AND Grupo_Rol = '" & CmbGrupos & "' "
 SQL2 = SQL2 & "ORDER BY Grupo_Rol,Codigo,Tipo_Rubro,ID,Ingresos desc,Egresos,Detalle "
 'MsgBox SQL2
 SelectAdodc AdoAsientoRol, SQL2
'Generamos el documento
 If Not (General_PDF) Then
    'SetNombrePRN = ""
    SetNombrePRN = Impresota_PDF
    NombFilePict = "Rol Pagos de " & NombreCliente & " " & CAnio & "-" & Format$(NumMeses, "00") & " " & CodigoUsuario
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
       Generar_Rol_Medio CodigoRol, 1, 0.5
       Generar_Rol_Medio CodigoRol, 10.5, 0.5
    Else
      'Si es rol completo arriba y abajo
       Generar_Rol CodigoRol, 1, 1
    End If
   'fin del documento
    cPrint.finalizaImpresion
    
    Presentar_PDF fPDF
 End If
End Sub

Public Sub Procesar_Asientos_Rol()
Dim VentasDia As Boolean
Dim Ctas_Catalogo As String
Dim Total_Aporte_Patronal As Currency
   RatonReloj
   CodigoCli = Ninguno
   I = CantCtas - 1
   For IE = 0 To I - 1
     For JE = IE + 1 To I
       If CtasRol(IE).Cta < CtasRol(JE).Cta Then
          Cta_Aux = CtasRol(IE).Cta
          Valor = CtasRol(IE).Valor
          CtasRol(IE).Cta = CtasRol(JE).Cta
          CtasRol(IE).Valor = Redondear(CtasRol(JE).Valor, 2)
          CtasRol(JE).Cta = Cta_Aux
          CtasRol(JE).Valor = Valor
       End If
     Next JE
   Next IE
   For IE = 0 To I - 1
     For JE = IE + 1 To I
       If CtasPro(IE).Cta < CtasPro(JE).Cta Then
          Cta_Aux = CtasPro(IE).Cta
          Valor = CtasPro(IE).Valor
          CtasPro(IE).Cta = CtasPro(JE).Cta
          CtasPro(IE).Valor = Redondear(CtasPro(JE).Valor, 2)
          CtasPro(JE).Cta = Cta_Aux
          CtasPro(JE).Valor = Valor
       End If
     Next JE
   Next IE
   For IE = 0 To I - 1
     For JE = IE + 1 To I
       If CtasPat(IE).Cta < CtasPat(JE).Cta Then
          Cta_Aux = CtasPat(IE).Cta
          Valor = CtasPat(IE).Valor
          CtasPat(IE).Cta = CtasPat(JE).Cta
          CtasPat(IE).Valor = Redondear(CtasPat(JE).Valor, 2)
          CtasPat(JE).Cta = Cta_Aux
          CtasPat(JE).Valor = Valor
       End If
     Next JE
   Next IE
   DetalleComp = Ninguno
   Trans_No = 101
   SQL1 = "DELETE " _
        & "FROM Asiento " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND CodigoU = '" & CodigoUsuario & "' " _
        & "AND T_No = " & Trans_No & " "
   ConectarAdoExecute SQL1
   
   SQL2 = "SELECT * " _
        & "FROM Asiento " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND T_No = " & Trans_No & " " _
        & "AND CodigoU = '" & CodigoUsuario & "' "
   SelectDataGrid DGAsiento(1), AdoAsiento1, SQL2
   
   Trans_No = 102
   SQL1 = "DELETE " _
        & "FROM Asiento " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND CodigoU = '" & CodigoUsuario & "' " _
        & "AND T_No = " & Trans_No & " "
   ConectarAdoExecute SQL1
   
   SQL2 = "SELECT * " _
        & "FROM Asiento " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND T_No = " & Trans_No & " " _
        & "AND CodigoU = '" & CodigoUsuario & "' "
   SelectDataGrid DGAsiento(2), AdoAsiento2, SQL2
   
   Trans_No = 100
   SQL1 = "DELETE " _
        & "FROM Asiento " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND CodigoU = '" & CodigoUsuario & "' " _
        & "AND T_No = " & Trans_No & " "
   ConectarAdoExecute SQL1
   
   SQL2 = "SELECT * " _
        & "FROM Asiento " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND T_No = " & Trans_No & " " _
        & "AND CodigoU = '" & CodigoUsuario & "' "
   SelectDataGrid DGAsiento(0), AdoAsiento, SQL2
      
   SQL2 = "SELECT * " _
        & "FROM Asiento_SC " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND T_No = " & Trans_No & " " _
        & "AND CodigoU = '" & CodigoUsuario & "' " _
        & "ORDER BY Codigo,Cta,DH "
   SelectAdodc AdoSubCta1, SQL2
  'Asiento del Rol de Pago
   DGNomina.Visible = False
   DGNomina1.Visible = False
   TotalCajaMN = 0
   Total_Cheque = 0
   Total_Pagar = 0
   TotalIngreso = 0
   Fecha_Vence = FechaFinal
  'Recolectamos informacion
   Ln_No = 0
   NoCheque = Ninguno
   DetalleComp = Ninguno
   Trans_No = 101
   LblConcepto(1).Caption = "(" & NumEmpresa & ") Registro de Aporte Patronal del " & FechaInicial & " al " & FechaFinal
   For IE = 0 To CantCtas - 1
      If CtasPat(IE).Cta <> "0" Then
         If CtasPat(IE).Valor >= 0 Then
            InsertarAsientos AdoAsiento1, CtasPat(IE).Cta, 0, CtasPat(IE).Valor, 0
         Else
            InsertarAsientos AdoAsiento1, CtasPat(IE).Cta, 0, 0, -CtasPat(IE).Valor
         End If
      End If
   Next IE
   Ln_No = 0
   Trans_No = 102
   LblConcepto(2).Caption = "(" & NumEmpresa & ") Registro de Provisiones de: 10cmo. 3ro., 10cmo. 4to., Vacaciones, Fondos de Reserva del " & FechaInicial & " al " & FechaFinal
   For IE = 0 To CantCtas - 1
      If CtasPro(IE).Cta <> "0" Then
         If CtasPro(IE).Valor >= 0 Then
            InsertarAsientos AdoAsiento2, CtasPro(IE).Cta, 0, CtasPro(IE).Valor, 0
         Else
            InsertarAsientos AdoAsiento2, CtasPro(IE).Cta, 0, 0, -CtasPro(IE).Valor
         End If
      End If
   Next IE
   
   Trans_No = 100
   Ln_No = 0
   For IE = 0 To CantCtas - 1
      If CtasRol(IE).Cta <> "0" Then
         If CtasRol(IE).Valor >= 0 Then
            InsertarAsientos AdoAsiento, CtasRol(IE).Cta, 0, CtasRol(IE).Valor, 0
         Else
            InsertarAsientos AdoAsiento, CtasRol(IE).Cta, 0, 0, -CtasRol(IE).Valor
         End If
      End If
   Next IE
   RatonReloj
   Contador = 0
  'Asignamos Codigo Contable segun el Abono
   TotalCajaMN = 0
   Total_Cheque = 0
   Total_Bancos = 0
   CodigoCli = Ninguno
   sSQL = "SELECT TRP.*,C.Cliente " _
        & "FROM Trans_Rol_de_Pagos As TRP,Clientes As C " _
        & "WHERE TRP.Fecha_D >= #" & FechaIni & "# " _
        & "AND TRP.Fecha_H <= #" & FechaFin & "# " _
        & "AND TRP.Item = '" & NumEmpresa & "' " _
        & "AND TRP.Periodo = '" & Periodo_Contable & "' " _
        & "AND TRP.Cod_Rol_Pago = 'Neto_Recibir' " _
        & "AND TRP.Codigo = C.Codigo "
   If CmbGrupos <> "TODOS" Then sSQL = sSQL & "AND TRP.Grupo_Rol = '" & CmbGrupos & "' "
   sSQL = sSQL & "ORDER BY C.Cliente,TRP.Cta,TRP.Codigo "
   SelectAdodc AdoAux, sSQL
   With AdoAux.Recordset
    If .RecordCount > 0 Then
       .MoveFirst
        Do While Not .EOF
           NoCheque = Ninguno
           Contador = Contador + 1
           DetalleComp = Format(Contador, "00") & ".- " & ULCase(.Fields("Cliente"))
          'Procesar Asiento de Efectivos, Cheques o Transferencias del Empleado
           CodigoCli = Ninguno
           If CheqCxP.value = 0 Then
              If UCase(Mid$(.Fields("Cheq_Dep_Transf"), 1, 3)) = "CHQ" Then
                 NoCheque = SinEspaciosDer(.Fields("Cheq_Dep_Transf"))
                 Total_Cheque = Total_Cheque + .Fields("Egresos")
                 CodigoCli = .Fields("Codigo")
                 InsertarAsientos AdoAsiento, .Fields("Cta"), 0, 0, .Fields("Egresos")
              ElseIf UCase(Mid$(.Fields("Cheq_Dep_Transf"), 1, 1)) = "C" Then
                 NoCheque = "Transf."
                 Total_Bancos = Total_Bancos + .Fields("Egresos")
                 CodigoCli = .Fields("Codigo")
                 InsertarAsientos AdoAsiento, .Fields("Cta"), 0, 0, .Fields("Egresos")
              Else
                 NoCheque = Ninguno
                 TotalCajaMN = TotalCajaMN + .Fields("Egresos")
                 InsertarAsientos AdoAsiento, .Fields("Cta"), 0, 0, .Fields("Egresos")
              End If
           Else
              'ElseIf UCase(Mid$(.Fields("Cheq_Dep_Transf"), 1, 2)) = "CP" Then
              NoCheque = "CP" & CStr(Year(FechaFinal) & Format(Month(FechaFinal), "00"))
              Total_Pagar = Total_Pagar + .Fields("Egresos")
              CodigoCli = .Fields("Codigo")
           End If
         .MoveNext
       Loop
    End If
   End With
   If CheqCxP.value <> 0 Then
      NoCheque = Ninguno
      DetalleComp = Ninguno
      InsertarAsientos AdoAsiento, SinEspaciosIzq(DCCxP), 0, 0, Total_Pagar
   End If
   CodigoCli = Ninguno
   SumaDebe = 0: SumaHaber = 0
   SQL2 = "SELECT * " _
        & "FROM Asiento " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND T_No = " & Trans_No & " " _
        & "AND CodigoU = '" & CodigoUsuario & "' " _
        & "ORDER BY A_No "
   SelectAdodc AdoAsiento, SQL2
   With AdoAsiento.Recordset
    If .RecordCount > 0 Then
       .MoveFirst
        Do While Not .EOF
           SumaDebe = SumaDebe + .Fields("DEBE")
           SumaHaber = SumaHaber + .Fields("HABER")
          .MoveNext
        Loop
    End If
   End With
   NoCheque = Ninguno
   DetalleComp = "Diferencia por Ctas incompletas"
   Cta_Diferencial = ReadAdoCta("Cta_Diferencial_Cambiario")
   Diferencia = Abs(SumaDebe - SumaHaber)
   If SumaDebe > SumaHaber Then
      InsertarAsientos AdoAsiento, Cta_Diferencial, 0, 0, Diferencia
   Else
      InsertarAsientos AdoAsiento, Cta_Diferencial, 0, Diferencia, 0
   End If
   SumaDebe = 0: SumaHaber = 0
   SQL2 = "SELECT * " _
        & "FROM Asiento " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND T_No = " & Trans_No & " " _
        & "AND CodigoU = '" & CodigoUsuario & "' " _
        & "ORDER BY A_No "
   SelectAdodc AdoAsiento, SQL2
   SelectDataGrid DGAsiento(0), AdoAsiento, SQL2
   With AdoAsiento.Recordset
    If .RecordCount > 0 Then
       .MoveFirst
        Do While Not .EOF
           SumaDebe = SumaDebe + .Fields("DEBE")
           SumaHaber = SumaHaber + .Fields("HABER")
          .MoveNext
        Loop
    End If
   End With
   LabelDebe.Caption = Format(SumaDebe, "#,##0.00")
   LabelHaber.Caption = Format(SumaHaber, "#,##0.00")
   LabelDiferencia.Caption = Format(SumaDebe - SumaHaber, "#,##0.00")
   DetalleComp = Ninguno
   
  'Presentamos los Empleados que no les alcanza el sueldo
   sSQL = "SELECT C.Cliente As Empleado,TRP.Egresos as Neto_a_Recibir " _
        & "FROM Clientes as C,Trans_Rol_de_Pagos As TRP " _
        & "WHERE Fecha_D >= #" & FechaIni & "# " _
        & "AND Fecha_H <= #" & FechaFin & "# " _
        & "AND Item = '" & NumEmpresa & "' " _
        & "AND Periodo = '" & Periodo_Contable & "' " _
        & "AND Egresos < 0 " _
        & "AND TRP.Codigo = C.Codigo " _
        & "ORDER BY Cliente "
  SelectDataGrid DGNomina1, AdoNomina1, sSQL
  RatonNormal
'''''   CxP Certificados
''''  Trans_No = 100
''''  Cta_Dcts_Certif = ReadAdoCta("Cta_Rol_Dcts_Certif")
''''  LeerCta Cta_Dcts_Certif
''''  If Codigo = Ninguno Then
''''     Si_No = True
''''     Cadena1 = Cadena1 & Cta_Dcts_Certif & vbCrLf
''''  End If
''''  sSQL = "SELECT CRP.Codigo,C.Cliente,SUM(CRP.Certificado) As T_D_C " _
''''       & "FROM Trans_Rol_Horas As CRP,Clientes As C,Catalogo_Rol_Pagos AS CR " _
''''       & "WHERE CRP.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
''''       & "AND CRP.Item = '" & NumEmpresa & "' " _
''''       & "AND CRP.Periodo = '" & Periodo_Contable & "' " _
''''       & "AND CR.Aporte_Cer > 0 "
''''  If CmbGrupos.Text <> "TODOS" Then sSQL = sSQL & "AND CR.Grupo_Rol = '" & CmbGrupos.Text & "' "
''''  sSQL = sSQL & "AND CRP.Codigo = C.Codigo " _
''''       & "AND CR.Codigo = CRP.Codigo " _
''''       & "AND CR.Item = CRP.Item " _
''''       & "AND CR.Periodo = CRP.Periodo " _
''''       & "GROUP BY CRP.Codigo,C.Cliente " _
''''       & "ORDER BY CRP.Codigo,C.Cliente "
''''  SelectAdodc AdoSubCta, sSQL
''''  Contador = 0
''''  Factura_No = Nota_No
''''  With AdoSubCta.Recordset
''''   If .RecordCount > 0 Then
''''       Do While Not .EOF
''''          Contador = Contador + 1
''''          CodigoCliente = .Fields("Codigo")
''''          NombreCliente = .Fields("Cliente")
''''          LRolPagos.Caption = "Certificados de Aportacion: " & Format(Contador / .RecordCount, "00%")
''''          InsertarCertificado Cta_Dcts_Certif, .Fields("T_D_C"), "P"
''''          Factura_No = Factura_No + 1
''''         .MoveNext
''''       Loop
''''   End If
''''  End With
  If CmbGrupos <> "TODOS" Then
     LblConcepto(0).Caption = LblConcepto(0).Caption & ", del Grupo " & CmbGrupos & " "
     LblConcepto(1).Caption = LblConcepto(1).Caption & ", del Grupo " & CmbGrupos & " "
     LblConcepto(2).Caption = LblConcepto(2).Caption & ", del Grupo " & CmbGrupos & " "
  End If
  DGNomina.Visible = True
  DGNomina1.Visible = True
End Sub

Private Sub CheqCxP_Click()
  If CheqCxP.value = 1 Then DCCxP.Visible = True Else DCCxP.Visible = False
End Sub

Private Sub CmbGrupos_GotFocus()
  LRolPagos.Caption = "ROL DE PAGOS MES DE " & UCase(MesesLetras(Month(FechaFinal)))
End Sub

Private Sub CmbGrupos_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub CmbGrupos_LostFocus()
  Listar_Empleados
End Sub

Private Sub CMes_LostFocus()
  Fecha_Del_AT CMes, CAnio
  sSQL = "SELECT Grupo_Rol " _
       & "FROM Catalogo_Rol_Pagos " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND Salario > 0 " _
       & "AND Fecha <= #" & BuscarFecha(FechaFinal) & "# " _
       & "GROUP BY Grupo_Rol " _
       & "ORDER BY Grupo_Rol "
  SelectAdodc AdoAux, sSQL
  CmbGrupos.Clear
  CmbGrupos.AddItem "TODOS"
  With AdoAux.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          CmbGrupos.AddItem .Fields("Grupo_Rol")
         .MoveNext
       Loop
   End If
  End With
  CmbGrupos.Text = "TODOS"
  Listar_Empleados
  Listar_CxCxP_SubMod
  CmbGrupos.SetFocus
End Sub

Public Sub Procesar_Nomina()
Dim Rol_I As Long
Dim Rol_M As Long
Dim Rol_F As Long
Dim Fecha_Rol_Mes As String
  Fecha_Del_AT CMes, CAnio
  Progreso_Barra.Incremento = 0
  Progreso_Barra.Valor_Maximo = 100
  Progreso_Barra.Mensaje_Box = ""
  Progreso_Esperar
  TextoImprimio = ""
  Rubros_Otros_Ingresos = ""
  SQL2 = "SELECT * " _
       & "FROM Catalogo_Cuentas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND LEN(Cod_Rol_Pago) > 1 " _
       & "ORDER BY Codigo "
  SelectAdodc AdoAux, SQL2
  With AdoAux.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          SQL2 = "DELETE * " _
               & "FROM Asiento_SC " _
               & "WHERE Item = '" & NumEmpresa & "' " _
               & "AND T_No = " & Trans_No & " " _
               & "AND CodigoU = '" & CodigoUsuario & "' " _
               & "AND Cta = '" & .Fields("Codigo") & "' "
          ConectarAdoExecute SQL2
         'MsgBox SQL2
          If .Fields("I_E_Emp") = "I" And .Fields("Con_IESS") Then
              Rubros_Otros_Ingresos = Rubros_Otros_Ingresos & "'" & .Fields("Cod_Rol_Pago") & "',"
          End If
         .MoveNext
       Loop
   End If
  End With
  
  SQL2 = "DELETE * " _
       & "FROM Asiento_SC " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND T_No = " & Trans_No & " " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND Mid$(Cta,1,1) = '5' "
  ConectarAdoExecute SQL2
  If CheqCxP.value Then
     SQL2 = "DELETE * " _
          & "FROM Asiento_SC " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND T_No = " & Trans_No & " " _
          & "AND CodigoU = '" & CodigoUsuario & "' " _
          & "AND Mid$(Detalle_SubCta,1,3) = 'SxP' "
     ConectarAdoExecute SQL2
     
     SQL2 = "DELETE * " _
          & "FROM Asiento_SC " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND T_No = " & Trans_No & " " _
          & "AND CodigoU = '" & CodigoUsuario & "' " _
          & "AND Cta = '" & SinEspaciosIzq(DCCxP) & "' "
     ConectarAdoExecute SQL2
  End If
' Procesamos los Ingresos/Egresos de Rol de Pagos
  RatonReloj
  Progreso_Barra.Mensaje_Box = "Encerando Asientos"
  Progreso_Esperar
  Inicializar_Cero_Asientos True
 
 'Borramos el rol mal procesado si este fue escrito mal la fecha
  Fecha_Rol_Mes = BuscarFecha("01/" & Format(Month(FechaFinal), "00") & "/" & Format(Year(FechaFinal), "0000"))
  SQL2 = "DELETE * " _
       & "FROM Trans_Rol_de_Pagos " _
       & "WHERE Fecha_D >= #" & Fecha_Rol_Mes & "# " _
       & "AND Fecha_H <= #" & FechaFin & "# " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND TP = '" & Ninguno & "' " _
       & "AND Numero = 0 "
  If CmbGrupos <> "TODOS" Then SQL2 = SQL2 & "AND Grupo_Rol = '" & CmbGrupos & "' "
  ConectarAdoExecute SQL2
 'Procedemos a procesar el Rol pedido del mes o quincena
  Opcion = 1
'''  FechaIni = BuscarFecha(fechainicial)
'''  FechaFin = BuscarFecha(fechafinal)
  
  SQL2 = "SELECT * " _
       & "FROM Trans_Rol_de_Pagos " _
       & "WHERE Fecha_D >= #" & FechaIni & "# " _
       & "AND Fecha_H <= #" & FechaFin & "# " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND TP <> '" & Ninguno & "' " _
       & "AND Numero <> 0 "
  If CmbGrupos <> "TODOS" Then SQL2 = SQL2 & "AND Grupo_Rol = '" & CmbGrupos & "' "
  SelectAdodc AdoAux, SQL2
 'MsgBox AdoAux.Recordset.RecordCount & vbCrLf & vbCrLf & SQL2
  If AdoAux.Recordset.RecordCount <= 0 Then
     Llenar_Rol_Pagos_Empleados False
     Listar_CxCxP_SubMod
  Else
     MsgBox "Este Rol ya fue Procesado" & vbCrLf & vbCrLf _
          & "Se procedera solo a presentar" & vbCrLf & vbCrLf _
          & "El Rol Procesado."
  End If
  Progreso_Barra.Mensaje_Box = "Procesar Asientos"
  Progreso_Esperar
  Procesar_Asientos_Rol
  
  Progreso_Barra.Mensaje_Box = "LLenar Rol Pagos"
  Progreso_Esperar
  Llenar_Rol_Pagos_Colectivo False
  Trans_No = 100
  DGAsiento(0).Visible = False
  DGAsiento(1).Visible = False
  DGAsiento(2).Visible = False
  LblConcepto(0).Caption = "Registro de Nómina correspondiente al mes de " & MesesLetras(Month(FechaFinal))
  LblConcepto(1).Caption = "Provision IESS Patronal correspondiente al mes de " & MesesLetras(Month(FechaFinal))
  LblConcepto(2).Caption = "Provision Decimo 3er., Decimo 4to., Vacaciones y de Fondos de Reserva correspondiente al mes de " & MesesLetras(Month(FechaFinal))
  If AdoNomina.Recordset.RecordCount > 0 Then
     AdoNomina.Recordset.MoveFirst
     Llenar_Rol_Pagos_Individual AdoNomina.Recordset.Fields("Codigo")
  End If
    
    DGAsiento(0).Visible = True
    DGAsiento(1).Visible = True
    DGAsiento(2).Visible = True
  
  Progreso_Barra.Incremento = Progreso_Barra.Valor_Maximo
  Progreso_Barra.Mensaje_Box = "Fin del Proceso del Rol"
  Progreso_Final
  
  Listar_CxCxP_SubMod
  If AdoNomina1.Recordset.RecordCount > 0 Then MsgBox "Existen empleados sin alcance a la nomina"
'''    If Len(TextoImprimio) > 1 Then
'''       TextoImprimio = "ADVERTENCIAS:" & vbCrLf & TextoImprimio
'''       FInfoError.Show
'''    End If
End Sub

'''Private Sub Command12_Click()
'''  Imprimir_Pagina True
'''End Sub

Public Sub Procesar_CxP()
  RatonReloj
  Inicializar_Cero_Asientos True
  Fecha_Del_AT CMes, CAnio
 'MsgBox FechaInicial & vbCrLf & FechaFinal
  Trans_No = 100
  Contador = 0
  Cadena1 = ""
  DGNomina.Visible = False
  DGAsiento(0).Visible = False
  SQL2 = "DELETE * " _
       & "FROM Asiento_SC " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND T_No = " & Trans_No & " " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND TC = 'P' "
  ConectarAdoExecute SQL2
 'SelectDataGrid DGAsiento(0), AdoAsiento, SQL2
 'IniciarAsientosDe DGAsiento(0), AdoAsiento
  Nota_No = ReadSetDataNum("Certificados", True, False)
 
 'Lista todos los Empleados para ligar con su SubCta de Modulo de CxC y CxP
  Si_No = False
  Listar_Empleados
  With AdoClientes.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          Total = 0
          Contador = Contador + 1
          CodigoCliente = .Fields("Codigo")
          NombreCliente = .Fields("Cliente")
          Saldos_CxC_CxP CodigoCliente, "P"
         .MoveNext
       Loop
       LeerCta SubCtaGen
       If Codigo = Ninguno Then
          Si_No = True
          Cadena1 = Cadena1 & SubCtaGen & vbCrLf
       End If
   End If
  End With
  Listar_CxCxP_SubMod
  DGNomina.Visible = True
  DGAsiento(0).Visible = True
  RatonNormal
  If Si_No Then MsgBox "Codigos Contables No existen: " & vbCrLf & Cadena1
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
  ConectarAdoExecute SQL1
  SQL2 = "DELETE * " _
       & "FROM Asiento_SC " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND T_No = " & Trans_No & " " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND TC = 'C' "
  ConectarAdoExecute SQL2
  SQL2 = "DELETE * " _
       & "FROM Trans_Rol_Pagos " _
       & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' "
  If CmbGrupos.Text <> "TODOS" Then SQL2 = SQL2 & "AND Grupo_Rol = '" & CmbGrupos & "' "
  ConectarAdoExecute SQL2

 'Grilla de Asientos
  SQL2 = "SELECT * " _
       & "FROM Asiento_SC " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND T_No = " & Trans_No & " " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND TC = 'C' " _
       & "ORDER BY Codigo,TC,Cta "
  SelectAdodc AdoAsientoSC, SQL2
  SQL2 = "SELECT * " _
       & "FROM Asiento " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND T_No = " & Trans_No & " " _
       & "AND CodigoU = '" & CodigoUsuario & "' "
  SelectDataGrid DGAsiento(0), AdoAsiento, SQL2
  Total = 0
  Contador = 0
  With AdoClientes.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       Do While Not .EOF
          Contador = Contador + 1
          CodigoCliente = .Fields("Codigo")
          NombreCliente = .Fields("Cliente")
          Grupo_No = .Fields("Grupo_Rol")
          Cta = .Fields("Cta_Quincena")
          Abono = Redondear(.Fields("Salario") / 2, 2)
          If Len(Cta) <= 1 Then Abono = 0
          NoCheque = Ninguno
          DetalleComp = Format(Contador, "00") & ".- " & NombreCliente
          If Len(.Fields("Cta_Transferencia")) > 1 Then
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
             InsValorCtaRol Cta, Abono
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
  SelectAdodc AdoAsientoSC, SQL2
  Ln_No = 0
  NoCheque = Ninguno
  DetalleComp = "Anticipo Empleados por quincena del mes de " & MesesLetras(Month(FechaFinal))
  For IE = 0 To CantCtas - 1
      If CtasRol(IE).Cta <> "0" Then
         Select Case Mid$(CtasRol(IE).Cta, 1, 1)
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
  SelectDataGrid DGAsiento(0), AdoAsiento, SQL2
  With AdoAsiento.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          Debe = Debe + .Fields("Haber")
          Haber = Haber + .Fields("Haber")
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
  ConectarAdoExecute SQL2
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
   Fecha_Del_AT CMes, CAnio
   Mensajes = "Seguro de Imprimir en:" & vbCrLf & Printer.DeviceName & "?"
   Titulo = "IMPRESION"
   Bandera = False
   SetPrinters.Show 1
   If PonImpresoraDefecto(SetNombrePRN) Then
      RatonReloj
     'Generamos el documento
      NombFilePict = "Rol Pagos " & CAnio & "-" & Format$(NumMeses, "00") & " R-" & RUC & " " & CodigoUsuario
      tPrint.TipoImpresion = Es_Printer
      tPrint.NombreArchivo = NombFilePict
      tPrint.TituloArchivo = "Rol de Pagos " & CAnio & "-" & Format(NumMeses, "00") & " " & RUC
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
      NombreBanco1 = UCase(Mid$(NombreBanco1, Len(SinEspaciosIzq(NombreBanco1)) + 1, Len(NombreBanco1)))
      With AdoNomina.Recordset
       If .RecordCount > 0 Then
          .MoveFirst
           Progreso_Barra.Incremento = 0
           Progreso_Barra.Valor_Maximo = .RecordCount
           Do While Not .EOF
              If NumEmpleados >= IR Then
                 Llenar_Rol_Pagos_Individual .Fields("Codigo"), True
                 Progreso_Barra.Mensaje_Box = "Imprimiendos Rol Individual de: (" & Pagina & ") " & .Fields("Nombre_Empleado")
                 Progreso_Esperar
                 If Medio_Rol Then
                    AnchoPict = cPrint.dAnchoPapel
                    AltoPict = cPrint.dAltoPapel
                    PosLinea = 0.5
                    Generar_Rol_Medio .Fields("Codigo"), 1, PosLinea
                    Generar_Rol_Medio .Fields("Codigo"), 10.5, PosLinea
                    cPrint.paginaNueva
                 Else
                    If Si_No Then
                       If Contador = 1 Then PosLinea = 0.5
                       If Contador = 2 Then PosLinea = PosCopiaY
                    Else
                       PosLinea = 0.5
                       Contador = 3
                    End If
                   'MsgBox PosLinea & ".........."
                    Generar_Rol .Fields("Codigo"), 1, PosLinea
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
      Presentar_PDF fPDF
      RatonNormal
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
      NombreBanco1 = UCase(Mid$(NombreBanco1, Len(SinEspaciosIzq(NombreBanco1)) + 1, Len(NombreBanco1)))
      With AdoNomina.Recordset
       If .RecordCount > 0 Then
          .MoveFirst
           Progreso_Barra.Incremento = 0
           Progreso_Barra.Valor_Maximo = .RecordCount
           If Un_Solo_Rol Then .Find ("Codigo = '" & Por_CI_RUC & "' ")
           Do While Not .EOF
              If NumEmpleados >= IR Then
                 Llenar_Rol_Pagos_Individual .Fields("Codigo"), True
                 Progreso_Barra.Mensaje_Box = "Imprimiendos Rol Individual de: (" & Pagina & ") " & .Fields("Nombre_Empleado")
                 Progreso_Esperar
                'Generamos el documento
                 NombFilePict = "R-" & RUC & " Periodo " & CAnio & "-" & Format$(NumMeses, "00") & " - NOMINA DE " & NombreCliente
                 tPrint.TipoImpresion = Es_PDF
                 tPrint.NombreArchivo = NombFilePict
                 tPrint.TituloArchivo = "Rol de Pagos " & CAnio & "-" & Format(NumMeses, "00") & " " & RUC
                 tPrint.TipoLetra = TipoHelvetica
                 tPrint.OrientacionPagina = 1
                 tPrint.PaginaA4 = True
                 tPrint.EsCampoCorto = False
                 tPrint.VerDocumento = False
                 Set cPrint = New cImpresion
                 cPrint.iniciaImpresion
                 PosLinea = 0.5
                 AnchoPict = cPrint.dAnchoPapel
                 AltoPict = cPrint.dAltoPapel
                 Generar_Rol .Fields("Codigo"), 1, PosLinea
                'fin del documento
                 cPrint.finalizaImpresion
                 
                'Enviamos por mail el rol
                 TMail.Asunto = "Rol Pagos " & CAnio & "-" & Format$(NumMeses, "00") & " De: " & NombreCliente
                 TMail.Mensaje = "Rol Pagos " & CAnio & "-" & Format$(NumMeses, "00") & " " & vbCrLf _
                               & "correspondiente a: " & NombreCliente & ". "
                 TMail.Adjunto = RutaSysBases & "\TEMP\" & tPrint.NombreArchivo & ".pdf"
                 
                'MsgBox "Remite " & TMail.de & vbCrLf & TMail.Adjunto & vbCrLf & .Fields("Codigo") & " - " & Lista_Emails
                 
                 If Email_CE_Copia Then
                    TMail.para = Lista_De_Correos(TMail.ListaMail).Correo_Electronico
                    FEnviarCorreos.Show 1
                 End If
                 
                 Do While Len(Lista_Emails) > 3
                    posPuntoComa = InStr(Lista_Emails, ";")
                    Email = Mid$(Lista_Emails, 1, posPuntoComa - 1)
                   'MsgBox "Lista: " & Email
                    If EsUnEmail(Email) Then
                       'MsgBox "Email: " & Email & vbCrLf & "File: " & TMail.Adjunto
                       TMail.para = Email
                       FEnviarCorreos.Show 1
                    End If
                    Lista_Emails = Mid$(Lista_Emails, posPuntoComa + 1, Len(Lista_Emails))
                 Loop
              End If
              NumEmpleados = NumEmpleados + 1
              If Un_Solo_Rol Then .MoveLast
             .MoveNext
           Loop
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

Public Sub Procesar_CxC()
  RatonReloj
  Progreso_Barra.Incremento = 0
  Progreso_Barra.Mensaje_Box = "Encerando Asientos"
  Progreso_Esperar
  Inicializar_Cero_Asientos True
  Fecha_Del_AT CMes, CAnio
  'MsgBox FechaInicial & vbCrLf & FechaFinal
  Contador = 0
  Cadena1 = ""
  LnSC_No = 0
  Trans_No = 100
  DGNomina.Visible = False
  DGAsiento(0).Visible = False
  SQL2 = "DELETE * " _
       & "FROM Asiento_SC " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND T_No = " & Trans_No & " " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND TC = 'C' "
  ConectarAdoExecute SQL2
  
' IniciarAsientosDe  DGAsiento(0), AdoAsiento
  Nota_No = ReadSetDataNum("Certificados", True, False)
 
 'Lista todos los Empleados para ligar con su SubCta de Modulo de CxC y CxP
  Si_No = False
  Listar_Empleados
  With AdoClientes.Recordset
   If .RecordCount > 0 Then
       Progreso_Barra.Valor_Maximo = Progreso_Barra.Valor_Maximo + .RecordCount
       Do While Not .EOF
          Progreso_Barra.Mensaje_Box = "Procesando CxC y CxP: "
          Progreso_Esperar
          Total = 0
          Contador = Contador + 1
          CodigoCliente = .Fields("Codigo")
          NombreCliente = .Fields("Cliente")
          Saldos_CxC_CxP CodigoCliente, "C"
         .MoveNext
       Loop
       LeerCta SubCtaGen
       If Codigo = Ninguno Then
          Si_No = True
          Cadena1 = Cadena1 & SubCtaGen & vbCrLf
       End If
   End If
  End With
  Progreso_Barra.Mensaje_Box = "Procesando CxC y CxP: "
  Progreso_Esperar
  Listar_CxCxP_SubMod
  DGNomina.Visible = True
  DGAsiento(0).Visible = True
  RatonNormal
  Progreso_Barra.Mensaje_Box = "Listando CxC y CxP: "
  Progreso_Final
  If Si_No Then MsgBox "Codigos Contables No existen: " & vbCrLf & Cadena1
End Sub

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
 Imprimir_Rol_Colectivo AdoNomina, AdoTotNomina
 MensajeEncabData = "R O L    D E    P A G O S"
 SQLMsg1 = "Desde el " & FechaInicial & " al " & FechaFinal
 SQLMsg2 = "PROVISIONES DEL ROL DE PAGO"
 SQLMsg3 = ""
 Orientacion_Pagina = 1
 Cuadricula = True
 ImprimirAdo AdoNominaProv, True, 1, 7
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
 AdoRolBancos.Open sSQL1, AdoStrCnn, , , adCmdText
 With AdoRolBancos
  If .RecordCount > 0 Then
      Progreso_Barra.Valor_Maximo = Progreso_Barra.Valor_Maximo + .RecordCount
     'Start a new workbook in Excel
      Set oExcel = CreateObject("Excel.Application")
      Set oBook = oExcel.workbooks.Add
     'Add data to cells of the first worksheet in the new workbook
      Set oSheet = oBook.Worksheets(1)
      RatonReloj
     'Encabezado de la hola
     'Ancho de las columnas
      oSheet.Columns("A").ColumnWidth = 5
      oSheet.Columns("B").ColumnWidth = 60
      oSheet.Columns("C").ColumnWidth = 40
      oSheet.Columns("D").ColumnWidth = 15
     'Detalle de las columnas
      oSheet.Range("A1").value = "F_P"
      oSheet.Range("B1").value = "Detalle_Rol"
      oSheet.Range("C1").value = "Forma_de_Pago"
      oSheet.Range("D1").value = "Neto_Recibir"
      oSheet.Range("A1:D1").Font.Bold = True
     'Datos de la hoja de calculo
      NFila = 1
      Codigo = .Fields("F_P")
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
         
         Select Case .Fields("F_P")
           Case "": TipoCta = ""
           Case "": TipoCta = ""
           Case "": TipoCta = ""
           Case "": TipoCta = ""
           Case "": TipoCta = ""
           Case "": TipoCta = ""
           Case Else: TipoCta = ""
         End Select

         
        'Religioso
         oSheet.Range("A" & CStr(NFila)).value = .Fields("F_P")
         oSheet.Range("B" & CStr(NFila)).value = .Fields("Detalle_Rol")
         oSheet.Range("C" & CStr(NFila)).value = .Fields("Forma_de_Pago")
         oSheet.Range("D" & CStr(NFila)).value = .Fields("Neto_Recibir")
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
  SelectDataGrid DGAsiento(0), AdoAsiento, SQL2
  SumaDebe = 0: SumaHaber = 0
  With AdoAsiento.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          SumaDebe = SumaDebe + .Fields("DEBE")
          SumaHaber = SumaHaber + .Fields("HABER")
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
  SelectAdodc AdoAsientoSC, SQL2
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
     GrabarComprobante Co
     ImprimirComprobantesDe False, Co
     SQL2 = "UPDATE Trans_Rol_de_Pagos " _
          & "SET TP = '" & Co.TP & "'," _
          & "Numero = " & Co.Numero & " " _
          & "WHERE Fecha_D >= #" & FechaIni & "# " _
          & "AND Fecha_H <= #" & FechaFin & "# " _
          & "AND Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' "
     If CmbGrupos <> "TODOS" Then SQL2 = SQL2 & "AND Grupo_Rol = '" & CmbGrupos & "' "
     ConectarAdoExecute SQL2
     
    'Grabamos las provisiones
     Trans_No = 101
     Co.TP = CompDiario
     Co.Efectivo = 0
     Co.Monto_Total = 0
     Co.Numero = ReadSetDataNum("Diario", True, True)
     Co.Concepto = LblConcepto(1).Caption
     Co.T_No = Trans_No
     GrabarComprobante Co
     ImprimirComprobantesDe False, Co
        
     Trans_No = 102
     Co.TP = CompDiario
     Co.Numero = ReadSetDataNum("Diario", True, True)
     Co.Concepto = LblConcepto(2).Caption
     Co.T_No = Trans_No
     GrabarComprobante Co
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

Private Sub DGAsiento_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF1 Then
     Select Case Index
       Case 0: GenerarDataTexto LRolPagos, AdoAsiento
       Case 1: GenerarDataTexto LRolPagos, AdoAsiento1
       Case 2: GenerarDataTexto LRolPagos, AdoAsiento2
     End Select
  End If
End Sub

Private Sub DGNomina_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  If CtrlDown And (vbKeyF9 = KeyCode) Then
     CodigoCli = DGNomina.Columns(1)
     NombreCliente = DGNomina.Columns(14)
     Mifecha = FechaFinal
     FComisiones.Show 1
     Procesar_Asientos_Rol
  End If
  If CtrlDown And (vbKeyF10 = KeyCode) Then
     CodigoCli = DGNomina.Columns(1)
     NombreCliente = DGNomina.Columns(10)
     Mifecha = FechaFinal
     FRolPago.Show 1
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

Private Sub DGNominaProv_KeyDown(KeyCode As Integer, Shift As Integer)
   Keys_Especiales Shift
   If KeyCode = vbKeyF1 Then GenerarDataTexto LRolPagos, AdoNominaProv
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
Dim TripleTab As Single
  FechaInicial = FechaSistema
  FechaFinal = FechaSistema
  AnchoTab = MDI_X_Max - 100
  AltoTab = MDI_Y_Max - 1500
  MitadTab = (MDI_Y_Max - 2800) / 2
  TripleTab = (MDI_Y_Max - 2450) / 3
  InicioTab = SSTab1.Top
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
 
 'Ancho y Largo de la pantalla
  SSTab1.width = AnchoTab
  SSTab1.Height = AltoTab
      
  fPDF.width = SSTab1.width - 250
  fPDF.Height = SSTab1.Height - 650
 
  AdoNomina.Top = SSTab1.Height - 500
  DGTotNomina.Top = SSTab1.Height - 1250
  
  DGNomina.Height = TripleTab
  DGTotNomina.Height = TripleTab
  DGNominaProv.Height = TripleTab
  
  DGNominaProv.Top = DGNomina.Top + DGNomina.Height
  DGTotNomina.Top = DGNominaProv.Top + DGNominaProv.Height
  
  DGNomina1.Height = MitadTab
  DGSubCtas.Top = DGNomina1.Height + 440
  DGSubCtas.Height = AltoTab - DGSubCtas.Top - 150
  
  DGAsiento(0).Height = SSTab1.Height - 1300
  DGAsiento(1).Height = MitadTab - 1000
  DGAsiento(2).Top = DGAsiento(1).Height + 1200
  DGAsiento(2).Height = AltoTab - DGAsiento(2).Top - 150
  LblConcepto(2).Top = DGAsiento(1).Height + 850
  Label1.Top = SSTab1.Height - 500
  Label19.Top = SSTab1.Height - 500
  LabelDiferencia.Top = SSTab1.Height - 500
  LabelDebe.Top = SSTab1.Height - 500
  LabelHaber.Top = SSTab1.Height - 500
  
  DGAsiento(0).width = AnchoTab - 350
  DGAsiento(1).width = AnchoTab - 350
  DGAsiento(2).width = AnchoTab - 350
  
  DGNomina.width = AnchoTab - 350
  DGNomina1.width = AnchoTab - 350
  DGNominaProv.width = AnchoTab - 350
  DGSubCtas.width = AnchoTab - 350
  DGTotNomina.width = AnchoTab - 350
  SetPapelLargo = 29
 
 'MsgBox PictRol.Width & vbCrLf & PictRol.Height
  Inicializar_Cero_Asientos True
  Trans_No = 100
 'Pagos sin alcance
  sSQL = "SELECT (Codigo & ' - ' & Cuenta) As Cuentas, * " _
       & "FROM Catalogo_Cuentas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND TC IN ('P','PS') " _
       & "AND DG = 'D' " _
       & "ORDER BY Codigo "
  SelectDBCombo DCCxP, AdoBanco, sSQL, "Cuentas"
  DCCxP.Visible = False
  Listar_Empleados
  Listar_CxCxP_SubMod
  RatonNormal
  CAnio.SetFocus
End Sub

Private Sub Form_Load()
  ConectarAdodc AdoAux
  ConectarAdodc AdoCaja
  ConectarAdodc AdoBanco
  ConectarAdodc AdoNomina
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
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
  'MsgBox Button.key
   Select Case Button.key
     Case "Salir"
          Unload LRolPagos
     Case "CxC"
          Procesar_CxC
     Case "CxP"
          Procesar_CxP
     Case "Quincena"
          Procesar_Quincena
     Case "Nomina"
          Procesar_Nomina
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
           If .RecordCount > 0 Then Procesar_Rol_Individual_Emails .Fields("Codigo")
          End With
     Case "Emails"
          Procesar_Rol_Individual_Emails
     Case "Primero"
          With AdoNomina.Recordset
           If .RecordCount > 0 Then
              .MoveFirst
               Llenar_Rol_Pagos_Individual .Fields("Codigo")
           End If
          End With
     Case "Antes"
          With AdoNomina.Recordset
           If .RecordCount > 0 Then
              .MovePrevious
               If .BOF Then .MoveLast
               Llenar_Rol_Pagos_Individual .Fields("Codigo")
           End If
          End With
     Case "Despues"
          With AdoNomina.Recordset
           If .RecordCount > 0 Then
               If .EOF Then .MoveFirst Else .MoveNext
               Llenar_Rol_Pagos_Individual .Fields("Codigo")
           End If
          End With
     Case "Ultimo"
          With AdoNomina.Recordset
           If .RecordCount > 0 Then
              .MoveLast
               If .Fields("Codigo") = "T O T A L " Then .MovePrevious
               Llenar_Rol_Pagos_Individual .Fields("Codigo")
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
                  Llenar_Rol_Pagos_Individual .Fields("Codigo")
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

Public Sub Listar_Empleados()
  Grupo_No = CmbGrupos
  sSQL = "SELECT C.Cliente,C.CI_RUC,C.Direccion,C.Telefono,C.Actividad,C.Email,C.Email2,CR.* " _
       & "FROM Clientes As C,Catalogo_Rol_Pagos As CR " _
       & "WHERE CR.Item = '" & NumEmpresa & "' " _
       & "AND CR.Periodo = '" & Periodo_Contable & "' " _
       & "AND CR.Salario > 0 " _
       & "AND CR.Fecha <= #" & BuscarFecha(FechaFinal) & "# "
  If Grupo_No <> "TODOS" Then sSQL = sSQL & "AND CR.Grupo_Rol = '" & Grupo_No & "' "
  sSQL = sSQL & "AND CR.Codigo = C.Codigo "
  If OpcGrupo.value Then
     sSQL = sSQL & "ORDER BY CR.Grupo_Rol,C.Cliente,CR.Codigo,CR.Cta_Transferencia "
  Else
     sSQL = sSQL & "ORDER BY C.Cliente,CR.Codigo,CR.Cta_Transferencia "
  End If
  SelectAdodc AdoClientes, sSQL
End Sub

Public Sub InsValorCtaRol(NCta As String, NValor As Currency)
  For IE = 0 To CantCtas - 1
      If CtasRol(IE).Cta = NCta Then
         CtasRol(IE).Valor = CtasRol(IE).Valor + Redondear(NValor, 2)
      End If
  Next IE
End Sub

Public Sub InsValorCtaPro(NCta As String, NValor As Currency)
  For IE = 0 To CantCtas - 1
      If CtasPro(IE).Cta = NCta Then
         CtasPro(IE).Valor = CtasPro(IE).Valor + Redondear(NValor, 2)
      End If
  Next IE
End Sub

Public Sub InsValorCtaPat(NCta As String, NValor As Currency)
  For IE = 0 To CantCtas - 1
      If CtasPat(IE).Cta = NCta Then
         CtasPat(IE).Valor = CtasPat(IE).Valor + Redondear(NValor, 2)
      End If
  Next IE
End Sub

Public Sub SetearCtasCierreRol(CtaFields As String)
  If CtaFields <> "0" Then
     Si_No = True
     For IE = 0 To CantCtas - 1
         If CtaFields = CtasRol(IE).Cta Then Si_No = False
     Next IE
     If Si_No Then
        IE = 0
        While IE < CantCtas
           If CtasRol(IE).Cta = "0" Then
              If Leer_Cta_Catalogo(CtaFields) = Ninguno Then
                 TextoImprimio = TextoImprimio & NivelNo & ", Cta Rol: " & CtaFields & vbCrLf
              Else
                 CtasRol(IE).Cta = CtaFields
              End If
              IE = CantCtas + 1
           End If
           IE = IE + 1
        Wend
     End If
  End If
End Sub

Public Sub SetearCtasCierrePro(CtaFields As String)
  If CtaFields <> "0" Then
     Si_No = True
     For IE = 0 To CantCtas - 1
         If CtaFields = CtasPro(IE).Cta Then Si_No = False
     Next IE
     If Si_No Then
        IE = 0
        While IE < CantCtas
           If CtasPro(IE).Cta = "0" Then
              If Leer_Cta_Catalogo(CtaFields) = Ninguno Then
                 TextoImprimio = TextoImprimio & NivelNo & ", Cta Prov.: " & CtaFields & vbCrLf
              Else
                 CtasPro(IE).Cta = CtaFields
              End If
              IE = CantCtas + 1
           End If
           IE = IE + 1
        Wend
     End If
  End If
End Sub

Public Sub SetearCtasCierrePat(CtaFields As String)
  If CtaFields <> "0" Then
     Si_No = True
     For IE = 0 To CantCtas - 1
         If CtaFields = CtasPat(IE).Cta Then Si_No = False
     Next IE
     If Si_No Then
        IE = 0
        While IE < CantCtas
           If CtasPat(IE).Cta = "0" Then
              If Leer_Cta_Catalogo(CtaFields) = Ninguno Then
                 TextoImprimio = TextoImprimio & NivelNo & ", Cta IESS: " & CtaFields & vbCrLf
              Else
                 CtasPat(IE).Cta = CtaFields
              End If
              IE = CantCtas + 1
           End If
           IE = IE + 1
        Wend
     End If
  End If
End Sub

Public Sub Insertar_Rol_Individual()
  LeerCta TRol_Pago.Cta
  If TRol_Pago.Detalle = Ninguno Then TRol_Pago.Detalle = Cuenta
  If TRol_Pago.Fecha_D = "" Then TRol_Pago.Fecha_D = FechaSistema
  If TRol_Pago.Fecha_H = "" Then TRol_Pago.Fecha_H = FechaSistema
  If TRol_Pago.Grupo_Rol = "" Then TRol_Pago.Grupo_Rol = Ninguno
  If TRol_Pago.Codigo = "" Then TRol_Pago.Codigo = Ninguno
  TRol_Pago.Ingresos = Redondear(TRol_Pago.Ingresos, 2)
  TRol_Pago.Egresos = Redondear(TRol_Pago.Egresos, 2)
 'Insertamos los rubros del rol individual
  If Len(Leer_Cta_Catalogo(TRol_Pago.Cta)) > 1 Then
     SetAdoAddNew "Trans_Rol_de_Pagos"
     SetAdoFields "T", TRol_Pago.T
     SetAdoFields "Codigo", TRol_Pago.Codigo
     SetAdoFields "Cta", TRol_Pago.Cta
     SetAdoFields "Detalle", TRol_Pago.Detalle
     SetAdoFields "Cheq_Dep_Transf", TRol_Pago.Cheq_Dep_Transf
     SetAdoFields "Ingresos", Redondear(TRol_Pago.Ingresos, 2)
     SetAdoFields "Egresos", Redondear(TRol_Pago.Egresos, 2)
     SetAdoFields "Dias", TRol_Pago.Dias
     SetAdoFields "Fecha_D", TRol_Pago.Fecha_D
     SetAdoFields "Fecha_H", TRol_Pago.Fecha_H
     SetAdoFields "Grupo_Rol", TRol_Pago.Grupo_Rol
     SetAdoFields "Horas", TRol_Pago.Horas
     SetAdoFields "Porc", TRol_Pago.Porc
     SetAdoFields "Retencion_No", TRol_Pago.Retencion_No
     SetAdoFields "Tipo_Rubro", TRol_Pago.Tipo_Rubro
    'SetAdoFields "ID", TRol_Pago.ID
     SetAdoFields "Cod_Rol_Pago", TRol_Pago.Cod_Rol_Pago
     SetAdoFields "Item", NumEmpresa
     SetAdoFields "CodigoU", CodigoUsuario
     If (TRol_Pago.Ingresos + TRol_Pago.Egresos) <> 0 Then SetAdoUpdate
  Else
     'MsgBox TRol_Pago.Cta & vbCrLf & TRol_Pago.Codigo & vbCrLf & TRol_Pago.Cod_Rol_Pago & vbCrLf & TRol_Pago.Ingresos & vbCrLf & TRol_Pago.Egresos
  End If
End Sub

Public Sub Limpiar_Rol_Individual()
  With TRol_Pago
      .T = Normal
      .Cta = Ninguno
      .Detalle = Ninguno
      .Cheq_Dep_Transf = Ninguno
      .Tipo_Rubro = Ninguno
      .SubModulo = Ninguno
      .Ingresos = 0
      .Egresos = 0
      .Dias = 0
      .Horas = 0
      .Porc = 0
      .Retencion_No = 0
      .ID = 0
  End With
End Sub

'''Public Sub Generar_Rol_Grafico(CodigoRol As String, Xo As Single, Yo As Single)
'''Dim ContLineas As Integer
'''Dim Es_Vacaciones As Boolean
'''Dim No_Recibe_Sueldo As Boolean
''''Empezamos a Escribir en papel grafico el Rol Individual
''' PictRol.FontName = TipoArial ' TipoArial - TipoArialNarrow - TipoComicSans
''''Los rubros que se ingresaron anteriormente con el rol
''' PictRol.FontBold = False
''' PictRol.FontSize = 11
''' No_Recibe_Sueldo = True
''' With AdoAsientoRol.Recordset
'''  If .RecordCount > 0 Then
'''      'Es_Vacaciones = .Fields("Vac")
'''      PictRol.FontBold = True
'''      PictPrint_Grafico PictRol, LogoTipo, Xo + 0.6, Yo + 0.2, 3, 1.5
'''      RutaDestino = RutaSistema & "\FOTOS\" & CodigoRol & ".jpg"
'''     'PictPrint_Grafico PictRol, RutaDestino, 7.5, 1, 2.5, 3
'''      PictRol.FontSize = 15
'''      PictPrint_Texto PictRol, Xo + 3.5, Yo + 0.2, UCase(Empresa)
'''      PictRol.FontSize = 10
'''      PictPrint_Texto PictRol, Xo + 3.5, Yo + 0.9, "Direccion: " & Direccion
'''      PictRol.FontSize = 12
'''      PictPrint_Texto PictRol, Xo + 3.5, Yo + 1.5, "ROL INDIVIDUAL DE PAGOS"
'''      PictRol.FontSize = 10
'''      PictPrint_Texto PictRol, Xo + 12.5, Yo + 1.9, "Desde: " & FechaInicial & " al: " & FechaFinal
'''      PictRol.Line (1, Yo + 2.4)-(19, Yo + 2.4)
'''      PictRol.FontSize = 11
'''      PictPrint_Texto PictRol, Xo + 1, Yo + 2.5, "Fecha de Ingreso:"
'''      PictPrint_Texto PictRol, Xo + 1, Yo + 3, "Beneficiario:"
'''      PictPrint_Texto PictRol, Xo + 13, Yo + 2.5, "Codigo:"
'''      PictPrint_Texto PictRol, Xo + 1, Yo + 3.5, "Periodo:"
'''      PictPrint_Texto PictRol, Xo + 13.4, Yo + 3, "Días:"
'''      PictPrint_Texto PictRol, Xo + 11.4, Yo + 3.5, "Forma de Pago:"
'''      PictRol.FontBold = False
'''      PictPrint_Texto PictRol, Xo + 4.8, Yo + 2.5, FechaTexto
'''      PictPrint_Texto PictRol, Xo + 3.6, Yo + 3, NombreCliente
'''      PictPrint_Texto PictRol, Xo + 14.6, Yo + 2.5, CICliente 'CodigoRol
'''      PictPrint_Texto PictRol, Xo + 2.7, Yo + 3.5, MesesLetras(Month(FechaFinal))
'''      PictPrint_Texto PictRol, Xo + 11.4, Yo + 3.95, TextoBanco
'''      PictRol.FontBold = True
'''      PFil = Yo + 4.5
''''''      PictRol.Line (1, PFil)-(11, PFil)
''''''      PFil = PFil + 0.05
''''''      PictPrint_Texto PictRol, Xo + 1.3, PFil, "D E T A L L E S     P A T R O N A L E S"
''''''      PFil = PFil + 0.6
''''''      PictRol.Line (1, PFil)-(11, PFil)
''''''      PFil = PFil + 0.1
''''''     .MoveFirst
''''''      Do While Not .EOF
''''''         If .Fields("Tipo_Rubro") = "PAT" And .Fields("Ingresos") <> 0 Then
''''''             PictRol.FontBold = True
''''''             PictPrint_Texto PictRol, 1.3, PFil, UCase(.Fields("Detalle"))
''''''             PictRol.FontBold = False
''''''             PictPrint_Texto PictRol, Xo + 9, PFil, Format(.Fields("Ingresos"), "#,###.00"), True, 1.9
''''''             PFil = PFil + 0.5
''''''         End If
''''''        .MoveNext
''''''      Loop
''''''      PFil = PFil + 0.1
'''      PictRol.Line (1, PFil)-(19, PFil)
'''      PFil = PFil + 0.05
'''      PictRol.FontBold = True
'''      PictPrint_Texto PictRol, Xo + 1.3, PFil, "D E T A L L E    D E L    E M P L E A D O"
'''      PictPrint_Texto PictRol, Xo + 13.5, PFil, "INGRESOS"
'''      PictPrint_Texto PictRol, Xo + 16.5, PFil, "EGRESOS"
'''      PFil = PFil + 0.6
'''      PictRol.Line (1, PFil)-(19, PFil)
'''      PFil = PFil + 0.1
'''     .MoveFirst
'''      Do While Not .EOF
'''         If .Fields("Tipo_Rubro") = "PER" Then
'''             If .Fields("Detalle") = "TOTAL A RECIBIR" Then
'''                 PictRol.FontBold = True
'''                 No_Recibe_Sueldo = False
'''             Else
'''                 PictRol.FontBold = False
'''             End If
'''             PictPrint_Texto PictRol, 1.3, PFil, UCase(.Fields("Detalle"))
'''             If .Fields("Ingresos") <> 0 Then
'''                 PictPrint_Texto PictRol, Xo + 13.5, PFil, Format(.Fields("Ingresos"), "#,###.00"), True, 1.9
'''             End If
'''             If .Fields("Egresos") <> 0 Then
'''                 PictPrint_Texto PictRol, Xo + 16.5, PFil, Format(.Fields("Egresos"), "#,###.00"), True, 1.9
'''             End If
'''             If .Fields("Cheq_Dep_Transf") <> Ninguno Then CodigoB = .Fields("Cheq_Dep_Transf")
'''             If .Fields("Dias") <> 0 Then I = .Fields("Dias")
'''             If .Fields("Horas") <> 0 Then J = .Fields("Horas")
'''             PFil = PFil + 0.5
'''         End If
'''        .MoveNext
'''      Loop
'''      If No_Recibe_Sueldo Then
'''         PictRol.FontBold = True
'''         PictPrint_Texto PictRol, 1.3, PFil, "TOTAL A RECIBIR"
'''         PictPrint_Texto PictRol, Xo + 16.5, PFil, Format(0, "#,##0.00"), True, 1.9
'''         PFil = PFil + 0.5
'''      End If
'''      PictRol.FontBold = False
'''      'PictPrint_Texto PictRol, Xo + 14.6, Yo + 4.5, Format(J, "#,##0.00")
'''      PictPrint_Texto PictRol, Xo + 14.6, Yo + 3, Format(I, "#,##0")
'''      PictPrint_Texto PictRol, Xo + 14.6, Yo + 3.5, CodigoB
'''  End If
''' End With
''' PictRol.FontBold = False
''' PictRol.Line (1, PFil - 0.5)-(19, PFil - 0.5), QBColor(Negro)
''' PictRol.Line (1, PFil)-(19, PFil), QBColor(Negro)
''' PictRol.FontSize = 10
''' sSQL = "SELECT * " _
'''      & "FROM Trans_Entrada_Salida " _
'''      & "WHERE Fecha BETWEEN #" & BuscarFecha(FechaInicial) & "# AND #" & BuscarFecha(FechaFinal) & "# " _
'''      & "AND Codigo = '" & CodigoRol & "' " _
'''      & "AND Item = '" & NumEmpresa & "' " _
'''      & "AND Periodo = '" & Periodo_Contable & "' " _
'''      & "AND ES = 'R' " _
'''      & "ORDER BY Fecha,Hora "
''' SelectAdodc AdoNovedades, sSQL
''' With AdoNovedades.Recordset
'''  If .RecordCount > 0 Then
'''      PictRol.FontBold = True
'''      PictPrint_Texto PictRol, Xo + 1.3, PFil, "OBSERVACIONES:"
'''      PictRol.FontBold = False
'''      PFil = PFil + 0.35
'''      Do While Not .EOF
'''         PictPrint_Texto PictRol, Xo + 1.3, PFil, .Fields("Tarea")
'''         PFil = PFil + 0.35
'''        .MoveNext
'''      Loop
'''  End If
''' End With
'''
''' PictPrint_Texto PictRol, Xo + 1, Yo + 11.5, String(12, "_")
''' PictPrint_Texto PictRol, Xo + 6, Yo + 11.5, String(17, "_")
''' PictPrint_Texto PictRol, Xo + 1.2, Yo + 12, "Empleador"
''' PictPrint_Texto PictRol, Xo + 6.5, Yo + 12, "Recibi conforme"
''' DetalleComp = ""
'''End Sub

Public Sub Generar_Rol(CodigoRol As String, Xo As Single, Yo As Single)
Dim ContLineas As Integer
Dim Es_Vacaciones As Boolean
Dim No_Recibe_Sueldo As Boolean
Dim PFilIni As Single
'Empezamos a Escribir en papel grafico el Rol Individual
'Los rubros que se ingresaron anteriormente con el rol
 cPrint.tipoDeLetra = TipoCourier 'TipoTimesRoman
 cPrint.tipoNegrilla = False
 cPrint.porteDeLetra = 10
 No_Recibe_Sueldo = True
 With AdoAsientoRol.Recordset
  If .RecordCount > 0 Then
     'Es_Vacaciones = .Fields("Vac")
      cPrint.tipoNegrilla = True
      cPrint.printImagen LogoTipo, Xo, Yo, 3, 1.4
      RutaDestino = RutaSistema & "\FOTOS\" & CodigoRol & ".jpg"
     'cPrint.printImagen  RutaDestino, 7.5, 1, 2.5, 3
      cPrint.porteDeLetra = 15
      If UCase$(RazonSocial) = UCase$(NombreComercial) Then
         cPrint.printTexto Xo + 3.5, Yo, UCase$(RazonSocial)
      Else
         cPrint.printTexto Xo + 3.5, Yo, UCase$(RazonSocial)
         cPrint.printTexto Xo + 3.5, Yo + 0.5, UCase$(NombreComercial)
      End If
      cPrint.porteDeLetra = 10
      cPrint.printTexto Xo + 3.5, Yo + 0.95, "Direccion: " & Direccion
      cPrint.porteDeLetra = 12
      cPrint.printTexto Xo + 3.5, Yo + 1.5, "ROL INDIVIDUAL DE PAGOS"
      cPrint.porteDeLetra = 9
      cPrint.printTexto Xo + 12.5, Yo + 1.9, "Desde: " & FechaInicial & " al: " & FechaFinal
      
      cPrint.printCuadroLinea Xo, Yo + 2.4, 19.5, Yo + 2.4
      
      cPrint.printTexto Xo, Yo + 2.5, "Fecha de Ingreso:"
      cPrint.printTexto Xo, Yo + 3, "Beneficiario:"
      cPrint.printTexto Xo + 11.5, Yo + 2.5, "Codigo:"
      cPrint.printTexto Xo, Yo + 3.5, "Periodo:"
      cPrint.printTexto Xo + 16.8, Yo + 2.5, "Días:"
      cPrint.printTexto Xo + 11.5, Yo + 3.5, "Forma de Pago:"
      cPrint.tipoNegrilla = False
      cPrint.printTexto Xo + 3.8, Yo + 2.5, FechaTexto
      cPrint.printTexto Xo + 2.6, Yo + 3, NombreCliente
      cPrint.printTexto Xo + 13, Yo + 2.5, CICliente  'CodigoRol
      cPrint.printTexto Xo + 1.7, Yo + 3.5, MesesLetras(Month(FechaFinal))
      cPrint.tipoNegrilla = True
      PFil = Yo + 4
      PFilIni = Yo + 4
      cPrint.printCuadroLinea Xo, Yo + 4, 19.5, Yo + 4
      PFil = PFil + 0.05
      cPrint.tipoNegrilla = True
      cPrint.printTexto Xo, PFil, "D E T A L L E    D E L    E M P L E A D O"
      cPrint.printTexto Xo + 14, PFil, "INGRESOS"
      cPrint.printTexto Xo + 16.5, PFil, "EGRESOS"
      PFil = PFil + 0.4
      cPrint.printCuadroLinea Xo, PFil, 19.5, PFil
      PFil = PFil + 0.1
     .MoveFirst
      Do While Not .EOF
         If .Fields("Tipo_Rubro") = "PER" Then
             If .Fields("Detalle") = "TOTAL A RECIBIR" Then
                 cPrint.tipoNegrilla = True
                 No_Recibe_Sueldo = False
             Else
                 cPrint.tipoNegrilla = False
             End If
             cPrint.printTexto Xo, PFil, UCase(.Fields("Detalle"))
             If .Fields("Ingresos") <> 0 Then
                 cPrint.printField Xo + 13, PFil, .Fields("Ingresos")
             End If
             If .Fields("Egresos") <> 0 Then
                 cPrint.printField Xo + 15.5, PFil, .Fields("Egresos")
             End If
             If .Fields("Cheq_Dep_Transf") <> Ninguno Then CodigoB = .Fields("Cheq_Dep_Transf")
             If .Fields("Dias") <> 0 Then I = .Fields("Dias")
             If .Fields("Horas") <> 0 Then J = .Fields("Horas")
             PFil = PFil + 0.4
         End If
        .MoveNext
      Loop
      If No_Recibe_Sueldo Then
         cPrint.tipoNegrilla = True
         cPrint.printTexto Xo, PFil, "TOTAL A RECIBIR"
         cPrint.printTexto Xo + 15.5, PFil, "0.00", True, 1.9
         PFil = PFil + 0.4
      End If
      cPrint.tipoNegrilla = False
      'cPrint.printTexto   Xo + 14.6, Yo + 4.5, Format(J, "#,##0.00")
      cPrint.printTexto Xo + 18, Yo + 2.5, Format(I, "#,##0")
      cPrint.printTexto Xo + 14.5, Yo + 3.5, CodigoB
  End If
 End With
 cPrint.tipoNegrilla = False
 cPrint.printCuadroLinea Xo, PFil - 0.4, 19.5, PFil - 0.4
 cPrint.printCuadroLinea Xo, PFil + 0.05, 19.5, PFil + 0.05
 cPrint.porteDeLetra = 10
 sSQL = "SELECT * " _
      & "FROM Trans_Entrada_Salida " _
      & "WHERE Fecha BETWEEN #" & BuscarFecha(FechaInicial) & "# AND #" & BuscarFecha(FechaFinal) & "# " _
      & "AND Codigo = '" & CodigoRol & "' " _
      & "AND Item = '" & NumEmpresa & "' " _
      & "AND Periodo = '" & Periodo_Contable & "' " _
      & "AND ES = 'R' " _
      & "ORDER BY Fecha,Hora "
 SelectAdodc AdoNovedades, sSQL
 With AdoNovedades.Recordset
  If .RecordCount > 0 Then
      cPrint.tipoNegrilla = True
      cPrint.printTexto Xo + 1.3, PFil, "OBSERVACIONES:"
      cPrint.tipoNegrilla = False
      PFil = PFil + 0.35
      Do While Not .EOF
         cPrint.printTexto Xo + 1.3, PFil, .Fields("Tarea")
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
End Sub

Public Sub Generar_Rol_Medio(CodigoRol As String, Xo As Single, Yo As Single)
Dim ContLineas As Integer
Dim Es_Vacaciones As Boolean

'Empezamos a Escribir en papel grafico el Rol Individual
'Los rubros que se ingresaron anteriormente con el rol
 cPrint.tipoNegrilla = False
 cPrint.porteDeLetra = 10
 With AdoAsientoRol.Recordset
  If .RecordCount > 0 Then
     'Es_Vacaciones = .Fields("Vac")
      cPrint.tipoNegrilla = True
      cPrint.printImagen LogoTipo, Xo + 0.5, Yo, 3, 1.5
      RutaDestino = RutaSistema & "\FOTOS\" & CodigoRol & ".jpg"
     'cPrint.printImagen  RutaDestino, 7.5, 1, 2.5, 3
      cPrint.porteDeLetra = 7
      If UCase$(RazonSocial) = UCase$(NombreComercial) Then
         cPrint.printTexto Xo + 3, Yo + 0.2, UCase$(RazonSocial)
      Else
         cPrint.printTexto Xo + 3, Yo, UCase$(RazonSocial)
         cPrint.printTexto Xo + 3, Yo + 0.4, UCase$(NombreComercial)
      End If
      
      cPrint.printTexto Xo + 3, Yo + 0.8, "Direccion: " & ULCase(Direccion)
      cPrint.porteDeLetra = 9
      cPrint.printTexto Xo + 3, Yo + 1.2, "ROL DE PAGOS INDIVIDUAL"
      cPrint.porteDeLetra = 8
      cPrint.printTexto Xo + 0.5, Yo + 1.6, "Desde: " & FechaInicial & " al: " & FechaFinal
      cPrint.printTexto Xo + 7, Yo + 1.6, "Periodo:"
      
      cPrint.printCuadroLinea Xo + 0.5, Yo + 2.05, Xo + 9.5, Yo + 2.05
      
      cPrint.porteDeLetra = 8
      cPrint.printTexto Xo + 0.5, Yo + 2.1, "Fecha de Ingreso:"
      cPrint.printTexto Xo + 5, Yo + 2.1, "Codigo:"
      cPrint.printTexto Xo + 8.5, Yo + 2.1, "Días:"
      cPrint.printTexto Xo + 0.5, Yo + 2.5, "Beneficiario:"
      
      cPrint.printTexto Xo + 0.5, Yo + 2.9, "Forma de Pago:"
      cPrint.tipoNegrilla = False
      cPrint.printTexto Xo + 3, Yo + 2.1, FechaTexto
      cPrint.printTexto Xo + 6.1, Yo + 2.1, CodigoRol
      cPrint.printTexto Xo + 8.2, Yo + 1.6, MesesLetras(Month(FechaFinal))
      cPrint.printTexto Xo + 2.2, Yo + 2.5, NombreCliente
      cPrint.tipoNegrilla = True
      PFil = Yo + 3.4
      cPrint.printCuadroLinea Xo + 0.5, PFil, Xo + 9.5, PFil
      PFil = PFil + 0.05
      cPrint.tipoNegrilla = True
      cPrint.printTexto Xo + 0.5, PFil, "DETALLE DEL EMPLEADO"
      cPrint.printTexto Xo + 6, PFil, "INGRESOS"
      cPrint.printTexto Xo + 8, PFil, "EGRESOS"
      PFil = PFil + 0.4
      cPrint.printCuadroLinea Xo + 0.5, PFil, Xo + 9.5, PFil
      PFil = PFil + 0.05
     .MoveFirst
      Do While Not .EOF
         If .Fields("Tipo_Rubro") = "PER" Then
             If .Fields("Detalle") = "TOTAL A RECIBIR" Then cPrint.tipoNegrilla = True Else cPrint.tipoNegrilla = False
             cPrint.printTexto Xo + 0.5, PFil, UCase(.Fields("Detalle"))
             If .Fields("Ingresos") <> 0 Then
                 cPrint.printField Xo + 5.3, PFil, .Fields("Ingresos")
             End If
             If .Fields("Egresos") <> 0 Then
                 cPrint.printField Xo + 7.3, PFil, .Fields("Egresos")
             End If
             If .Fields("Cheq_Dep_Transf") <> Ninguno Then CodigoB = .Fields("Cheq_Dep_Transf")
             If .Fields("Dias") <> 0 Then I = .Fields("Dias")
             If .Fields("Horas") <> 0 Then J = .Fields("Horas")
             PFil = PFil + 0.4
         End If
        .MoveNext
      Loop
      cPrint.tipoNegrilla = False
      'cPrint.printTexto   Xo + 14.6, Yo + 4.5, Format(J, "#,##0.00")
      cPrint.printTexto Xo + 9.2, Yo + 2.1, Format(I, "#,##0")
      cPrint.printTexto Xo + 2.7, Yo + 2.9, CodigoB
  End If
 End With
 cPrint.tipoNegrilla = False
 'PFil = PFil + 0.1
 cPrint.printCuadroLinea Xo + 0.5, PFil - 0.4, Xo + 9.5, PFil - 0.4
 cPrint.printCuadroLinea Xo + 0.5, PFil + 0.05, Xo + 9.5, PFil + 0.05
 cPrint.porteDeLetra = 8
 sSQL = "SELECT * " _
      & "FROM Trans_Entrada_Salida " _
      & "WHERE Fecha BETWEEN #" & BuscarFecha(FechaInicial) & "# AND #" & BuscarFecha(FechaFinal) & "# " _
      & "AND Codigo = '" & CodigoRol & "' " _
      & "AND Item = '" & NumEmpresa & "' " _
      & "AND Periodo = '" & Periodo_Contable & "' " _
      & "AND ES = 'R' " _
      & "ORDER BY Fecha,Hora "
 SelectAdodc AdoNovedades, sSQL
 With AdoNovedades.Recordset
  If .RecordCount > 0 Then
      cPrint.tipoNegrilla = True
      cPrint.printTexto Xo + 1.3, PFil, "OBSERVACIONES:"
      cPrint.tipoNegrilla = False
      PFil = PFil + 0.35
      Do While Not .EOF
         cPrint.printTexto Xo + 1.3, PFil, .Fields("Tarea")
         PFil = PFil + 0.35
        .MoveNext
      Loop
  End If
 End With
 cPrint.printTexto Xo + 1, Yo + 11.5, String(12, "_")
 cPrint.printTexto Xo + 6, Yo + 11.5, String(17, "_")
 cPrint.printTexto Xo + 1.2, Yo + 12, "Empleador"
 cPrint.printTexto Xo + 6.2, Yo + 12, "Recibi conforme"
 DetalleComp = ""
End Sub

'''Public Sub Generar_Rol_Grafico_Medio(CodigoRol As String, Xo As Single, Yo As Single)
'''Dim ContLineas As Integer
'''Dim Es_Vacaciones As Boolean
''' Fecha_Del_AT CMes, CAnio
''''Empezamos a Escribir en papel grafico el Rol Individual
''' PictRol.FontName = TipoArial ' TipoArial - TipoArialNarrow - TipoComicSans
''''Los rubros que se ingresaron anteriormente con el rol
''' PictRol.FontBold = False
''' PictRol.FontSize = 11
''' With AdoAsientoRol.Recordset
'''  If .RecordCount > 0 Then
'''      'Es_Vacaciones = .Fields("Vac")
'''      PictRol.FontBold = True
'''
'''      PictPrint_Grafico PictRol, LogoTipo, Xo + 0.1, Yo + 0.2, 3, 1.5
'''      RutaDestino = RutaSistema & "\FOTOS\" & CodigoRol & ".JPG"
'''     'PictPrint_Grafico PictRol, RutaDestino, 7.5, 1, 2.5, 3
'''      PictRol.FontSize = 10
'''      PictPrint_Texto PictRol, Xo + 2, Yo + 0.2, UCase(Empresa)
'''      PictRol.FontSize = 8
'''      PictPrint_Texto PictRol, Xo + 3, Yo + 0.8, "Direccion: " & Direccion
'''      PictRol.FontSize = 10
'''      PictPrint_Texto PictRol, Xo + 3, Yo + 1.4, "ROL INDIVIDUAL DE PAGOS"
'''      PictRol.FontSize = 8
'''      PictPrint_Texto PictRol, Xo + 5.5, Yo + 1.9, "Desde: " & FechaInicial & " al: " & FechaFinal
'''      PictRol.Line (Xo + 0.5, Yo + 2.4)-(Xo + 10, Yo + 2.4)
'''      PictRol.FontSize = 8
'''      PictPrint_Texto PictRol, Xo + 0.5, Yo + 2.5, "Fecha de Ingreso:"
'''      PictPrint_Texto PictRol, Xo + 0.5, Yo + 3, "Beneficiario:"
'''      PictPrint_Texto PictRol, Xo + 6.8, Yo + 2.5, "Codigo:"
'''      PictPrint_Texto PictRol, Xo + 0.5, Yo + 3.5, "Periodo:"
'''      PictPrint_Texto PictRol, Xo + 8.5, Yo + 3, "Días:"
'''      PictPrint_Texto PictRol, Xo + 5, Yo + 3.5, "Forma de Pago:"
'''      If TextoBanco <> Ninguno Then PictPrint_Texto PictRol, Xo + 0.5, Yo + 4, "Banco:"
'''      PictRol.FontBold = False
'''      PictPrint_Texto PictRol, Xo + 3.5, Yo + 2.5, FechaTexto
'''      PictPrint_Texto PictRol, Xo + 2.5, Yo + 3, NombreCliente
'''      PictPrint_Texto PictRol, Xo + 8.2, Yo + 2.5, CodigoRol
'''      PictPrint_Texto PictRol, Xo + 1.9, Yo + 3.5, MesesLetras(Month(FechaFinal))
'''      PictRol.FontBold = True
'''      If TextoBanco <> Ninguno Then PFil = Yo + 4.5 Else PFil = Yo + 4.1
''''''      PictRol.Line (1, PFil)-(11, PFil)
''''''      PFil = PFil + 0.05
''''''      PictPrint_Texto PictRol, Xo + 1.3, PFil, "D E T A L L E S     P A T R O N A L E S"
''''''      PFil = PFil + 0.6
''''''      PictRol.Line (1, PFil)-(11, PFil)
''''''      PFil = PFil + 0.1
''''''     .MoveFirst
''''''      Do While Not .EOF
''''''         If .Fields("Tipo_Rubro") = "PAT" And .Fields("Ingresos") <> 0 Then
''''''             PictRol.FontBold = True
''''''             PictPrint_Texto PictRol, 1.3, PFil, UCase(.Fields("Detalle"))
''''''             PictRol.FontBold = False
''''''             PictPrint_Texto PictRol, Xo + 9, PFil, Format(.Fields("Ingresos"), "#,###.00"), True, 1.9
''''''             PFil = PFil + 0.5
''''''         End If
''''''        .MoveNext
''''''      Loop
''''''      PFil = PFil + 0.1
'''      PictRol.Line (Xo + 0.5, PFil)-(Xo + 10, PFil)
'''      PFil = PFil + 0.05
'''      PictRol.FontBold = True
'''      PictPrint_Texto PictRol, Xo + 0.5, PFil, "D E T A L L E    D E L    E M P L E A D O"
'''      PictPrint_Texto PictRol, Xo + 6.6, PFil, "INGRESOS"
'''      PictPrint_Texto PictRol, Xo + 8.6, PFil, "EGRESOS"
'''      PFil = PFil + 0.6
'''      PictRol.Line (Xo + 0.5, PFil)-(Xo + 10, PFil)
'''      PFil = PFil + 0.1
'''     .MoveFirst
'''      Do While Not .EOF
'''         If .Fields("Tipo_Rubro") = "PER" Then
'''             If .Fields("Detalle") = "TOTAL A RECIBIR" Then PictRol.FontBold = True Else PictRol.FontBold = False
'''             PictPrint_Texto PictRol, Xo + 0.5, PFil, UCase(.Fields("Detalle"))
'''             If .Fields("Ingresos") <> 0 Then
'''                 PictPrint_Texto PictRol, Xo + 6, PFil, Format(.Fields("Ingresos"), "#,###.00"), True, 1.9
'''             End If
'''             If .Fields("Egresos") <> 0 Then
'''                 PictPrint_Texto PictRol, Xo + 8, PFil, Format(.Fields("Egresos"), "#,###.00"), True, 1.9
'''             End If
'''             If .Fields("Cheq_Dep_Transf") <> Ninguno Then CodigoB = .Fields("Cheq_Dep_Transf")
'''             If .Fields("Dias") <> 0 Then I = .Fields("Dias")
'''             If .Fields("Horas") <> 0 Then J = .Fields("Horas")
'''             PFil = PFil + 0.5
'''         End If
'''        .MoveNext
'''      Loop
'''      PictRol.FontBold = False
'''      'PictPrint_Texto PictRol, Xo + 14.6, Yo + 4.5, Format(J, "#,##0.00")
'''      PictPrint_Texto PictRol, Xo + 9.4, Yo + 3, Format(I, "#,##0")
'''      PictPrint_Texto PictRol, Xo + 7.5, Yo + 3.5, CodigoB
'''      If TextoBanco <> Ninguno Then PictPrint_Texto PictRol, Xo + 1.7, Yo + 4, TextoBanco
'''  End If
''' End With
''' PictRol.FontBold = False
''' PictRol.Line (Xo + 0.5, PFil - 0.5)-(Xo + 10, PFil - 0.5), QBColor(Negro)
''' PictRol.Line (Xo + 0.5, PFil)-(Xo + 10, PFil), QBColor(Negro)
''' PictRol.FontSize = 8
''' sSQL = "SELECT * " _
'''      & "FROM Trans_Entrada_Salida " _
'''      & "WHERE Fecha BETWEEN #" & BuscarFecha(FechaInicial) & "# AND #" & BuscarFecha(FechaFinal) & "# " _
'''      & "AND Codigo = '" & CodigoRol & "' " _
'''      & "AND Item = '" & NumEmpresa & "' " _
'''      & "AND Periodo = '" & Periodo_Contable & "' " _
'''      & "AND ES = 'R' " _
'''      & "ORDER BY Fecha,Hora "
''' SelectAdodc AdoNovedades, sSQL
''' With AdoNovedades.Recordset
'''  If .RecordCount > 0 Then
'''      PictRol.FontBold = True
'''      PictPrint_Texto PictRol, Xo + 0.5, PFil, "OBSERVACIONES:"
'''      PictRol.FontBold = False
'''      PFil = PFil + 0.35
'''      Do While Not .EOF
'''         PictPrint_Texto PictRol, Xo + 0.5, PFil, .Fields("Tarea")
'''         PFil = PFil + 0.35
'''        .MoveNext
'''      Loop
'''  End If
''' End With
''' PictPrint_Texto PictRol, Xo + 1, Yo + 12, String(12, "_")
''' PictPrint_Texto PictRol, Xo + 6, Yo + 12, String(17, "_")
''' PictPrint_Texto PictRol, Xo + 1.2, Yo + 12.4, "Empleador"
''' PictPrint_Texto PictRol, Xo + 6.5, Yo + 12.4, "Recibi conforme"
''' DetalleComp = ""
'''End Sub

Public Sub Llenar_Rol_Pagos_Empleados(Es_quincena As Boolean)
Dim Rol_I As Long
Dim Rol_M As Long
Dim Rol_F As Long
Dim Rol_D As Long
Dim Cheque_No As Long
Dim Total_GP As Currency
Dim Total_IR As Currency
Dim Total_IR_Mes_Ant As Currency
Dim Total_IR_Meses As Currency
Dim Total_Otros_Ing As Currency
Dim Total_DFR As Currency
Dim Total_CRR As Currency
Dim DH_SubCta As String
Dim Fecha_Empleado As String
Dim Dias_Temp As Integer
Dim Dias_Laborados As Integer
Dim Dias_Laborados_Emp As Integer
Dim Dias_Del_Mes As Integer
Dim Reingreso_FR As Boolean
Dim Fecha_IESS As String
Dim Meses_IR As Integer
Dim Meses_IR_Mes_Ant As Integer
Dim Dias_x_Mes As Byte
Dim Cta_SueldoV As String
Dim SueldoV As Currency

' Procesamos los Ingresos/Egresos de Rol de Pagos
  RatonReloj
  Progreso_Barra.Mensaje_Box = "Determinando Datos a procesar"
  Progreso_Iniciar
  Opcion = 1
  Ctas_Asientos_Rol
  Listar_Empleados
  Listar_CxCxP_SubMod
 'Verificamos las cuentas de proceso del Rol
  If Month(FechaFinal) = 2 Then
     Dias_Del_Mes = 28
     Fecha_IESS = "28/" & Format(Month(FechaFinal), "00") & "/" & Format(Year(FechaFinal), "0000")
  Else
     Dias_Del_Mes = 30
     Fecha_IESS = "30/" & Format(Month(FechaFinal), "00") & "/" & Format(Year(FechaFinal), "0000")
  End If
  TextoValido TxtCheque, , True
  Meses_Provision = 12
  If Sueldo_Basico <= 0 Then
     Sueldo_Basico = 0
     MsgBox "Falta setear el sueldo Basico"
  End If
  Cheque_No = Val(TxtCheque)
  Grupo_No = CmbGrupos
  Trans_No = 100
 'Eliminamos el rol antiguo si es necesario
  SQL2 = "DELETE * " _
       & "FROM Trans_Rol_de_Pagos " _
       & "WHERE Fecha_D >= #" & FechaIni & "# " _
       & "AND Fecha_H <= #" & FechaFin & "# " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' "
  If CmbGrupos <> "TODOS" Then SQL2 = SQL2 & "AND Grupo_Rol = '" & CmbGrupos & "' "
  ConectarAdoExecute SQL2
  
 'MsgBox SQL2
  sSQL = "SELECT CR.Grupo_Rol,TR.Codigo,C.Cliente,CR.Fecha,CR.Vivienda,CR.Salud,CR.Educacion,CR.Alimentacion," _
       & "CR.Vestimenta,CR.Discapacidad,CR.Tercera_Edad,CR.TiempoParcial,SUM(TR.Dias) As T_Dias," _
       & "SUM(TR.Horas) As T_Horas,SUM(TR.Horas_Exts) As T_Horas_Exts,SUM(TR.Ing_Liquido) As T_Ing_Liquido," _
       & "SUM(TR.Ing_Horas_Ext) As T_Ing_Horas_Ext " _
       & "FROM Trans_Rol_Horas As TR,Clientes As C,Catalogo_Rol_Pagos AS CR " _
       & "WHERE TR.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
       & "AND TR.Item = '" & NumEmpresa & "' " _
       & "AND TR.Periodo = '" & Periodo_Contable & "' " _
       & "AND CR.Salario > 0 "
  If CmbGrupos <> "TODOS" Then sSQL = sSQL & "AND CR.Grupo_Rol = '" & Grupo_No & "' "
  sSQL = sSQL & "AND TR.Codigo = C.Codigo " _
       & "AND CR.Codigo = TR.Codigo " _
       & "AND CR.Item = TR.Item " _
       & "AND CR.Periodo = TR.Periodo " _
       & "GROUP BY CR.Grupo_Rol,C.Cliente,CR.Fecha,TR.Codigo,CR.Vivienda,CR.Salud,CR.Educacion,CR.Alimentacion," _
       & "CR.Vestimenta,CR.Discapacidad,CR.Tercera_Edad,CR.TiempoParcial " _
       & "ORDER BY CR.Grupo_Rol,C.Cliente,TR.Codigo "
  SelectAdodc AdoSubCta, sSQL
  Contador = 0
  With AdoSubCta.Recordset
   If .RecordCount > 0 Then
       Progreso_Barra.Valor_Maximo = Progreso_Barra.Valor_Maximo + .RecordCount
       Do While Not .EOF
         'If Mid$(.Fields("Cliente"), 1, 4) = "PEÑA" Then MsgBox NombreCliente & vbCrLf & Dias_Del_Mes & vbCrLf & .Fields("TiempoParcial")
          
          Total_Otros_Ing = 0
          IEESS_Per = 0: IEESS_Pat = 0: IEESS_Ext = 0
          Debe = 0: Haber = 0: Diferencia = 0
          Contador = Contador + 1
          CodigoCliente = .Fields("Codigo")
          Total = Redondear(.Fields("T_Ing_Liquido"), 2)      ' Sueldo Liquido
          Saldo = Redondear(.Fields("T_Ing_Horas_Ext"), 2)    ' Horas Extras
          
          TotalIngreso = Total + Saldo
          Total_IR = TotalIngreso
          TotalAbonos = TotalIngreso
          Total_IESS = TotalIngreso
          
         'Datos Generales del Rol de Pago
          TRol_Pago.Grupo_Rol = .Fields("Grupo_Rol")
          Limpiar_Rol_Individual
          TRol_Pago.Codigo = CodigoCliente
          TRol_Pago.Fecha_D = FechaInicial
          TRol_Pago.Fecha_H = FechaFinal
          Fecha_Empleado = .Fields("Fecha")
          Dias_Laborados_Emp = CFechaLong(FechaFinal) - CFechaLong(Fecha_Empleado)
          
          If (Day(Fecha_Empleado) > 1) And (Dias_Laborados_Emp <= 30) Then
             Dias_Laborados = (CFechaLong(Fecha_IESS) - CFechaLong(Fecha_Empleado))
          Else
             Dias_Laborados = Day(Fecha_IESS)
          End If
          Dias_Temp = CFechaLong(Fecha_IESS) - CFechaLong(Fecha_Empleado)
         
         'Averiguamos si existe asignado el Empleado al Rol
          If AdoClientes.Recordset.RecordCount > 0 Then
             AdoClientes.Recordset.MoveFirst
             AdoClientes.Recordset.Find ("Codigo = '" & CodigoCliente & "' ")
             If Not AdoClientes.Recordset.EOF Then
                NombreCliente = AdoClientes.Recordset.Fields("Cliente")
                If AdoClientes.Recordset.Fields("Reingreso_FR") Then Dias_Temp = 365
                
                Progreso_Barra.Mensaje_Box = NombreCliente
                Progreso_Esperar
                Cta_SueldoV = Ninguno
                SueldoV = 0
                IEESS_Per = AdoClientes.Recordset.Fields("IEESS_Per")
                IEESS_Pat = AdoClientes.Recordset.Fields("IEESS_Pat")
                IEESS_Ext = AdoClientes.Recordset.Fields("IEESS_ExtC")
                Rol_I = CFechaLong(AdoClientes.Recordset.Fields("FechaVI"))
                Rol_F = CFechaLong(AdoClientes.Recordset.Fields("FechaVF"))
                Rol_M = CFechaLong(FechaFinal)
                Cta_Sueldo = AdoClientes.Recordset.Fields("Cta_Sueldo")
                If Rol_I <= Rol_M And Rol_M <= Rol_F Then
                   Cta_SueldoV = AdoClientes.Recordset.Fields("Cta_Vacacion")
                End If
                Cta_Horas_Extras = AdoClientes.Recordset.Fields("Cta_Horas_Ext")
                Reingreso_FR = AdoClientes.Recordset.Fields("Reingreso_FR")
                
               'Vacaciones
                If Len(AdoClientes.Recordset.Fields("Cta_Vacaciones_G")) > 1 And _
                   Len(AdoClientes.Recordset.Fields("Cta_Vacaciones_P")) > 1 Then
                   Total_DFR = Redondear(TotalIngreso / (Meses_Provision * 2), 2)
                   If Dias_Laborados < 30 Then Total_DFR = Redondear((Total_DFR / 30) * Dias_Laborados, 2)
                   Limpiar_Rol_Individual
                   TRol_Pago.ID = 5
                   TRol_Pago.Cod_Rol_Pago = "Vacaciones"
                   TRol_Pago.Tipo_Rubro = "PRO"
                   TRol_Pago.Ingresos = Total_DFR
                   TRol_Pago.Cta = AdoClientes.Recordset.Fields("Cta_Vacaciones_G")
                   Insertar_Rol_Individual
                   InsValorCtaPro TRol_Pago.Cta, TRol_Pago.Ingresos
                   Limpiar_Rol_Individual
                   TRol_Pago.ID = 5
                   TRol_Pago.Cod_Rol_Pago = "Vacaciones"
                   TRol_Pago.Tipo_Rubro = "PRO"
                   TRol_Pago.Egresos = Total_DFR
                   TRol_Pago.Cta = AdoClientes.Recordset.Fields("Cta_Vacaciones_P")
                   Insertar_Rol_Individual
                   InsValorCtaPro TRol_Pago.Cta, -TRol_Pago.Egresos
                End If
                
               'Sueldo Liquido
                If (Dias_Laborados < Dias_Del_Mes) And (Dias_Laborados_Emp < Dias_Del_Mes) Then
                   Total_DFR = Redondear((Total / Dias_Del_Mes) * Dias_Laborados, 2)
                Else
                   Total_DFR = Total
                End If
                If Rol_I <= Rol_M And Rol_M <= Rol_F And Cta_SueldoV <> Ninguno Then
                   Rol_D = Rol_F - Rol_I
                   SueldoV = Redondear((Total_DFR / Dias_Del_Mes) * Rol_D, 2)
                   Total_DFR = Total_DFR - SueldoV
                End If
                Limpiar_Rol_Individual
                If CheqHoras.value = 1 Then TRol_Pago.Horas = .Fields("T_Horas")
                TRol_Pago.ID = 10
                TRol_Pago.Cod_Rol_Pago = "Salario"
                TRol_Pago.Tipo_Rubro = "PER"
                TRol_Pago.Cta = Cta_Sueldo
                TRol_Pago.Dias = .Fields("T_Dias")
                TRol_Pago.Ingresos = Total_DFR
                Insertar_Rol_Individual
                InsValorCtaRol TRol_Pago.Cta, TRol_Pago.Ingresos
                Debe = Debe + TRol_Pago.Ingresos
                
                If Rol_I <= Rol_M And Rol_M <= Rol_F And Cta_SueldoV <> Ninguno Then
                   Limpiar_Rol_Individual
                   If CheqHoras.value = 1 Then TRol_Pago.Horas = .Fields("T_Horas")
                   TRol_Pago.ID = 11
                   TRol_Pago.Cod_Rol_Pago = "Salario"
                   TRol_Pago.Tipo_Rubro = "PER"
                   TRol_Pago.Cta = Cta_SueldoV
                   TRol_Pago.Dias = .Fields("T_Dias")
                   TRol_Pago.Ingresos = SueldoV
                   Insertar_Rol_Individual
                   InsValorCtaRol TRol_Pago.Cta, TRol_Pago.Ingresos
                   Debe = Debe + TRol_Pago.Ingresos
                End If
                
               'Ingreso de Horas Extras
                Limpiar_Rol_Individual
                TRol_Pago.ID = 12
                TRol_Pago.Cod_Rol_Pago = "Hor_Ext"
                TRol_Pago.Tipo_Rubro = "PER"
                TRol_Pago.Cta = Cta_Horas_Extras
                If CheqHoras.value = 1 Then TRol_Pago.Horas = .Fields("T_Horas_Exts")
                TRol_Pago.Ingresos = Saldo
                Debe = Debe + TRol_Pago.Ingresos
                Insertar_Rol_Individual
                InsValorCtaRol TRol_Pago.Cta, TRol_Pago.Ingresos
               'Debe = Debe + TRol_Pago.Ingresos  antes
               'If CodigoCliente = "1720073269" Then MsgBox NombreCliente & vbCrLf & CodigoCliente & " ..."
               '================================================================
               'Insertamos datos de los Rubros adicionales de Ingresos y Egresos
               'con o sin calculo al IESS
               '================================================================
                NoMes = Month(FechaFinal)
                sSQL = "SELECT * " _
                     & "FROM Catalogo_Rol_Rubros " _
                     & "WHERE Item = '" & NumEmpresa & "' " _
                     & "AND Periodo = '" & Periodo_Contable & "' " _
                     & "AND CPais = '" & CodigoPais & "' " _
                     & "AND Codigo = '" & TRol_Pago.Codigo & "' " _
                     & "AND Mes = " & NoMes & " " _
                     & "ORDER BY I_E Desc,Detalle "
                SelectAdodc AdoAux, sSQL
                
                If AdoAux.Recordset.RecordCount > 0 Then
                   Do While Not AdoAux.Recordset.EOF
                     'Datos iniciales para determinar si enviamos a CxC o CxP en submodulos
                      Si_No = AdoAux.Recordset.Fields("Calc_IESS")
                      TipoProc = AdoAux.Recordset.Fields("I_E")
                      
                      If AdoAux.Recordset.Fields("TV") = "%" Then
                         Tasa = AdoAux.Recordset.Fields("Valor") / 100
                         Valor = Redondear(TotalAbonos * Tasa, 2)
                      Else
                         Valor = AdoAux.Recordset.Fields("Valor")
                      End If
                      Cta = Leer_Cta_Catalogo(AdoAux.Recordset.Fields("Cta"))
                      If Cta <> Ninguno Then
                         Select Case SubCta
                           Case "C", "P"
                                SQL2 = "SELECT * " _
                                     & "FROM Catalogo_CxCxP " _
                                     & "WHERE Item = '" & NumEmpresa & "' " _
                                     & "AND TC = '" & SubCta & "' " _
                                     & "AND Cta = '" & Cta & "' " _
                                     & "AND Codigo = '" & CodigoCliente & "' "
                                SelectAdodc AdoAsientoSC, SQL2
                                If AdoAsientoSC.Recordset.RecordCount <= 0 Then
                                   SetAdoAddNew "Catalogo_CxCxP"
                                   SetAdoFields "Codigo", CodigoCliente
                                   SetAdoFields "Cta", Cta
                                   SetAdoFields "TC", SubCta
                                   SetAdoUpdate
                                End If
                                If TipoProc = "E" Then DH_SubCta = "2" Else DH_SubCta = "1"
                                Trans_No = 100
                                SQL2 = "SELECT * " _
                                     & "FROM Asiento_SC " _
                                     & "WHERE Item = '" & NumEmpresa & "' " _
                                     & "AND DH = '" & DH_SubCta & "' " _
                                     & "AND TC = '" & SubCta & "' " _
                                     & "AND Cta = '" & Cta & "' " _
                                     & "AND Codigo = '" & CodigoCliente & "' " _
                                     & "AND CodigoU = '" & CodigoUsuario & "' " _
                                     & "AND T_No = " & Trans_No & " "
                                SelectAdodc AdoAsientoSC, SQL2
                                If AdoAsientoSC.Recordset.RecordCount > 0 Then
                                   AdoAsientoSC.Recordset.Fields("Valor") = Valor
                                   AdoAsientoSC.Recordset.Update
                                Else
                                   SetAdoAddNew "Asiento_SC"
                                   SetAdoFields "Codigo", CodigoCliente
                                   SetAdoFields "Beneficiario", NombreCliente
                                   SetAdoFields "Cta", Cta
                                   SetAdoFields "Valor", Valor
                                   SetAdoFields "FECHA_V", FechaFinal
                                   SetAdoFields "TC", SubCta
                                   SetAdoFields "DH", "2"
                                   SetAdoFields "TM", "1"
                                   SetAdoFields "T_No", Trans_No
                                   SetAdoFields "SC_No", LnSC_No
                                   SetAdoFields "CodigoU", CodigoUsuario
                                   SetAdoUpdate
                                   LnSC_No = LnSC_No + 1
                                End If
                           Case Else
                               'Ingreso adicional que se calcula el Seguro Social
                                Limpiar_Rol_Individual
                                TRol_Pago.Cod_Rol_Pago = AdoAux.Recordset.Fields("Cod_Rol_Pago")
                                TRol_Pago.Tipo_Rubro = "PER"
                                TRol_Pago.Cta = AdoAux.Recordset.Fields("Cta")
                                If TipoProc = "I" Then
                                   If Si_No Then
                                      Total_Otros_Ing = Total_Otros_Ing + Valor
                                      Total_IESS = Total_IESS + Valor
                                   End If
                                   TRol_Pago.ID = 249
                                   TRol_Pago.Ingresos = Valor
                                   InsValorCtaRol TRol_Pago.Cta, TRol_Pago.Ingresos
                                   Debe = Debe + TRol_Pago.Ingresos
                                Else
                                   TRol_Pago.ID = 250
                                   TRol_Pago.Egresos = Valor
                                   InsValorCtaRol TRol_Pago.Cta, -TRol_Pago.Egresos
                                   Haber = Haber + TRol_Pago.Egresos
                                End If
                                Insertar_Rol_Individual
                         End Select
                      End If
                      AdoAux.Recordset.MoveNext
                   Loop
                End If
                              
               'Ingreso por CxP
                SQL2 = "SELECT * " _
                     & "FROM Asiento_SC " _
                     & "WHERE Item = '" & NumEmpresa & "' " _
                     & "AND T_No = 100 " _
                     & "AND CodigoU = '" & CodigoUsuario & "' " _
                     & "AND Codigo = '" & TRol_Pago.Codigo & "' " _
                     & "AND TC = 'P' " _
                     & "ORDER BY Beneficiario, Codigo, Cta, DH "
                SelectAdodc AdoAsientoSC, SQL2
                If AdoAsientoSC.Recordset.RecordCount > 0 Then
                   AdoAsientoSC.Recordset.MoveFirst
                   Do While Not AdoAsientoSC.Recordset.EOF
                      CodigoB = Replace(AdoAsientoSC.Recordset.Fields("Cta"), ".", "_")
                      Limpiar_Rol_Individual
                      TRol_Pago.ID = 251
                      TRol_Pago.Cod_Rol_Pago = "CxP_" & Mid$(CodigoB, Len(CodigoB) - 4, 5)
                      TRol_Pago.Tipo_Rubro = "PER"
                      TRol_Pago.Cta = AdoAsientoSC.Recordset.Fields("Cta")
                      TRol_Pago.Ingresos = AdoAsientoSC.Recordset.Fields("Valor")
                      Insertar_Rol_Individual
                      InsValorCtaRol TRol_Pago.Cta, TRol_Pago.Ingresos
                      Debe = Debe + TRol_Pago.Ingresos
                      Total_IR = Total_IR + TRol_Pago.Ingresos   'Imp Renta Emple.
                      AdoAsientoSC.Recordset.MoveNext
                   Loop
                End If
                
               'Egresos por CxC
                SQL2 = "SELECT * " _
                     & "FROM Asiento_SC " _
                     & "WHERE Item = '" & NumEmpresa & "' " _
                     & "AND T_No = 100 " _
                     & "AND CodigoU = '" & CodigoUsuario & "' " _
                     & "AND Codigo = '" & TRol_Pago.Codigo & "' " _
                     & "AND TC = 'C' " _
                     & "ORDER BY Beneficiario,Codigo,Cta,DH "
                SelectAdodc AdoAsientoSC, SQL2
                If AdoAsientoSC.Recordset.RecordCount > 0 Then
                   AdoAsientoSC.Recordset.MoveFirst
                   Do While Not AdoAsientoSC.Recordset.EOF
                      CodigoB = Replace(AdoAsientoSC.Recordset.Fields("Cta"), ".", "_")
                      Limpiar_Rol_Individual
                      TRol_Pago.ID = 252
                      TRol_Pago.Cod_Rol_Pago = "CxC_" & Mid$(CodigoB, Len(CodigoB) - 4, 5)
                      TRol_Pago.Tipo_Rubro = "PER"
                      TRol_Pago.Cta = AdoAsientoSC.Recordset.Fields("Cta")
                      TRol_Pago.Egresos = AdoAsientoSC.Recordset.Fields("Valor")
                      Insertar_Rol_Individual
                      InsValorCtaRol TRol_Pago.Cta, -TRol_Pago.Egresos
                      Haber = Haber + TRol_Pago.Egresos
                      AdoAsientoSC.Recordset.MoveNext
                   Loop
                End If
                 
                If Dias_Del_Mes >= 28 And Month(FechaFinal) = 2 Then Dias_Del_Mes = 30
                If Dias_Laborados >= 28 And Month(FechaFinal) = 2 Then Dias_Laborados = 30
                                  
               'Aporte Patronal 12.15%
                Total_DFR = Redondear(Total_IESS * IEESS_Pat, 2)
                If Dias_Laborados < Dias_Del_Mes Then Total_DFR = Redondear((Total_DFR / Dias_Del_Mes) * Dias_Laborados, 2)
                Total_IESS_Pat = Total_DFR
                Limpiar_Rol_Individual
                TRol_Pago.ID = 253
                TRol_Pago.Cod_Rol_Pago = "Aporte_Pat"
                TRol_Pago.Tipo_Rubro = "PAT"
                TRol_Pago.Detalle = "IESS Patronal " & Redondear(IEESS_Pat * 100, 2) & "%(G)"
                TRol_Pago.Cta = AdoClientes.Recordset.Fields("Cta_Aporte_Patronal_G")
                TRol_Pago.Porc = IEESS_Pat
                TRol_Pago.Ingresos = Total_IESS_Pat
                Insertar_Rol_Individual
                InsValorCtaPat TRol_Pago.Cta, TRol_Pago.Ingresos
                
                Limpiar_Rol_Individual
                TRol_Pago.ID = 253
                TRol_Pago.Cod_Rol_Pago = "Aporte_Pat"
                TRol_Pago.Tipo_Rubro = "PAT"
                TRol_Pago.Detalle = "IESS Patronal " & Redondear(IEESS_Pat * 100, 2) & "%(P)"
                TRol_Pago.Cta = AdoClientes.Recordset.Fields("Cta_IESS_Patronal")
                TRol_Pago.Porc = IEESS_Pat
                TRol_Pago.Egresos = Total_IESS_Pat
                Insertar_Rol_Individual
                InsValorCtaPat TRol_Pago.Cta, -TRol_Pago.Egresos
                
               'IESS Personal 9.45%
                Total_DFR = Redondear(Total_IESS * IEESS_Per, 2)
                If Dias_Laborados < Dias_Del_Mes Then Total_DFR = Redondear((Total_DFR / Dias_Del_Mes) * Dias_Laborados, 2)
                Total_IESS_Per = Total_DFR
               'If Mid$(NombreCliente, 1, 15) = "ALMEIDA CEVALLO" Then MsgBox NombreCliente & vbCrLf & Dias_Del_Mes
                Limpiar_Rol_Individual
                TRol_Pago.ID = 254
                TRol_Pago.Cod_Rol_Pago = "Aporte_Per"
                TRol_Pago.Tipo_Rubro = "PER"
                TRol_Pago.Detalle = "IESS Personal " & Redondear(IEESS_Per * 100, 2) & "%"
                TRol_Pago.Cta = AdoClientes.Recordset.Fields("Cta_IESS_Personal")
                TRol_Pago.Porc = IEESS_Per
                TRol_Pago.Egresos = Total_IESS_Per
                Insertar_Rol_Individual
                InsValorCtaRol TRol_Pago.Cta, -TRol_Pago.Egresos
                Haber = Haber + TRol_Pago.Egresos
                Total_IR = Total_IR - TRol_Pago.Egresos
                
               'IESS Personal 3.41%
                Total_DFR = Redondear(Total_IESS * IEESS_Ext, 2)
                If Dias_Laborados < Dias_Del_Mes Then Total_DFR = Redondear((Total_DFR / Dias_Del_Mes) * Dias_Laborados, 2)
                Total_IESS_Ext = Total_DFR
               'If Mid$(NombreCliente, 1, 15) = "ALMEIDA CEVALLO" Then MsgBox NombreCliente & vbCrLf & Dias_Del_Mes
                Limpiar_Rol_Individual
                TRol_Pago.ID = 254
                TRol_Pago.Cod_Rol_Pago = "Aporte_Ext_C"
                TRol_Pago.Tipo_Rubro = "PER"
                TRol_Pago.Detalle = "IESS Extensión Conyugue " & Redondear(IEESS_Ext * 100, 2) & "%"
                TRol_Pago.Cta = AdoClientes.Recordset.Fields("Cta_IESS_Personal")
                TRol_Pago.Porc = IEESS_Ext
                TRol_Pago.Egresos = Total_IESS_Ext
                Insertar_Rol_Individual
                InsValorCtaRol TRol_Pago.Cta, -TRol_Pago.Egresos
                Haber = Haber + TRol_Pago.Egresos
                Total_IR = Total_IR - TRol_Pago.Egresos
                
               'Fondos de Reserva al Gasto O directo al rol
                If Reingreso_FR Then Dias_Temp = 366
               ' if  MsgBox Dias_Temp & vbCrLf & Fecha_IESS & vbCrLf & Fecha_Empleado
                If Dias_Temp > 365 Then
                   Total_DFR = Redondear(Total_IESS * 0.0833, 2)
                   If (Dias_Temp - 365) < 255 Then Dias_x_Mes = Dias_Temp - 365 Else Dias_x_Mes = 255
                   If 1 < Dias_x_Mes And Dias_x_Mes < 30 Then Total_DFR = Redondear((Total_DFR / Dias_Del_Mes) * Dias_x_Mes, 2)
''                   If AdoClientes.Recordset.Fields("Codigo") = "1711939460" Then
''                      MsgBox Total_DFR & vbCrLf & AdoClientes.Recordset.Fields("Fecha")
''                   End If
                   'MsgBox Total_DFR
                   
                   If AdoClientes.Recordset.Fields("Pagar_Fondo_Reserva") Then
                      If Len(AdoClientes.Recordset.Fields("Cta_Fondo_Reserva_G")) > 1 Then
                         Limpiar_Rol_Individual
                         TRol_Pago.ID = 12
                         TRol_Pago.Cod_Rol_Pago = "Fon_Res_G"
                         TRol_Pago.Tipo_Rubro = "PER"
                         TRol_Pago.Cta = AdoClientes.Recordset.Fields("Cta_Fondo_Reserva_G")
                         TRol_Pago.Ingresos = Total_DFR   'TotalIngreso
                         Insertar_Rol_Individual
                         InsValorCtaRol TRol_Pago.Cta, TRol_Pago.Ingresos
                         Debe = Debe + TRol_Pago.Ingresos
                      End If
                   Else
                      If Len(AdoClientes.Recordset.Fields("Cta_Fondo_Reserva_G")) > 1 And _
                         Len(AdoClientes.Recordset.Fields("Cta_Fondo_Reserva_P")) > 1 Then
                         Limpiar_Rol_Individual
                         TRol_Pago.ID = 3
                         TRol_Pago.Cod_Rol_Pago = "Fon_Res_P"
                         TRol_Pago.Tipo_Rubro = "PRO"
                         TRol_Pago.Cta = AdoClientes.Recordset.Fields("Cta_Fondo_Reserva_G")
                         TRol_Pago.Ingresos = Total_DFR
                         Insertar_Rol_Individual
                         InsValorCtaPro TRol_Pago.Cta, TRol_Pago.Ingresos
                         
                         Limpiar_Rol_Individual
                         TRol_Pago.ID = 3
                         TRol_Pago.Cod_Rol_Pago = "Fon_Res_P"
                         TRol_Pago.Tipo_Rubro = "PRO"
                         TRol_Pago.Cta = AdoClientes.Recordset.Fields("Cta_Fondo_Reserva_P")
                         TRol_Pago.Egresos = Total_DFR
                         Insertar_Rol_Individual
                         InsValorCtaPro TRol_Pago.Cta, -TRol_Pago.Egresos
                      End If
                   End If
                End If
                
               'Decimo Tercer
                Total_DFR = Redondear(Total_IESS / Meses_Provision, 2)
                
               'Si el Empleado sale antes del mes
                If .Fields("T_Dias") < Dias_Laborados Then Dias_Laborados = .Fields("T_Dias")
                
                If Dias_Laborados < 30 Then Total_DFR = Redondear((Total_DFR / 30) * Dias_Laborados, 2)
                
               'If Mid$(NombreCliente, 1, 15) = "MACIAS INTRIAGO" Then MsgBox NombreCliente & vbCrLf & Dias_Del_Mes & vbCrLf & .Fields("T_Dias") & " = " & Total_DFR
                
                If AdoClientes.Recordset.Fields("Pagar_Decimos") Then
                   If Len(AdoClientes.Recordset.Fields("Cta_Decimo_Tercer_G")) > 1 Then
                      Limpiar_Rol_Individual
                      TRol_Pago.ID = 19
                      TRol_Pago.Cod_Rol_Pago = "Decimo_III_G"
                      TRol_Pago.Tipo_Rubro = "PER"
                      TRol_Pago.Ingresos = Total_DFR
                      TRol_Pago.Cta = AdoClientes.Recordset.Fields("Cta_Decimo_Tercer_G")
                      Insertar_Rol_Individual
                      InsValorCtaRol TRol_Pago.Cta, TRol_Pago.Ingresos
                      Debe = Debe + TRol_Pago.Ingresos
                   End If
                Else
                   If Len(AdoClientes.Recordset.Fields("Cta_Decimo_Tercer_G")) > 1 And _
                      Len(AdoClientes.Recordset.Fields("Cta_Decimo_Tercer_P")) Then
                      Limpiar_Rol_Individual
                      TRol_Pago.ID = 1
                      TRol_Pago.Cod_Rol_Pago = "Decimo_III"
                      TRol_Pago.Tipo_Rubro = "PRO"
                      TRol_Pago.Ingresos = Total_DFR
                      TRol_Pago.Cta = AdoClientes.Recordset.Fields("Cta_Decimo_Tercer_G")
                      Insertar_Rol_Individual
                      InsValorCtaPro TRol_Pago.Cta, TRol_Pago.Ingresos
                      
                      Limpiar_Rol_Individual
                      TRol_Pago.ID = 1
                      TRol_Pago.Cod_Rol_Pago = "Decimo_III"
                      TRol_Pago.Tipo_Rubro = "PRO"
                      TRol_Pago.Egresos = Total_DFR
                      TRol_Pago.Cta = AdoClientes.Recordset.Fields("Cta_Decimo_Tercer_P")
                      Insertar_Rol_Individual
                      InsValorCtaPro TRol_Pago.Cta, -TRol_Pago.Egresos
                   End If
                End If
                
               'Decimo Cuarto
                Total_DFR = Redondear(Sueldo_Basico / Meses_Provision, 2)
                If Dias_Laborados < 30 Then Total_DFR = Redondear((Total_DFR / 30) * Dias_Laborados, 2)
                
                If .Fields("TiempoParcial") Then
                    If .Fields("T_Horas") = 160 Then
                        Total_DFR = Redondear(Total_DFR / 2, 2)
                    Else
                        Total_DFR = Redondear((.Fields("T_Horas") * Total_DFR) / 160, 2)
                    End If
                End If
                
               ' If Mid$(NombreCliente, 1, 4) = "PEÑA" Then MsgBox NombreCliente & vbCrLf & Dias_Del_Mes & vbCrLf & .Fields("TiempoParcial")
                
                If AdoClientes.Recordset.Fields("Pagar_Decimos") Then
                   If Len(AdoClientes.Recordset.Fields("Cta_Decimo_Cuarto_G")) Then
                      Limpiar_Rol_Individual
                      TRol_Pago.ID = 20
                      TRol_Pago.Cod_Rol_Pago = "Decimo_IV_G"
                      TRol_Pago.Tipo_Rubro = "PER"
                      TRol_Pago.Ingresos = Total_DFR
                      TRol_Pago.Cta = AdoClientes.Recordset.Fields("Cta_Decimo_Cuarto_G")
                      Insertar_Rol_Individual
                      InsValorCtaRol TRol_Pago.Cta, TRol_Pago.Ingresos
                      Debe = Debe + TRol_Pago.Ingresos
                   End If
                Else
                   If Len(AdoClientes.Recordset.Fields("Cta_Decimo_Cuarto_G")) > 1 And _
                      Len(AdoClientes.Recordset.Fields("Cta_Decimo_Cuarto_P")) > 1 Then
                      Limpiar_Rol_Individual
                      TRol_Pago.ID = 2
                      TRol_Pago.Cod_Rol_Pago = "Decimo_IV"
                      TRol_Pago.Tipo_Rubro = "PRO"
                      TRol_Pago.Ingresos = Total_DFR
                      TRol_Pago.Cta = AdoClientes.Recordset.Fields("Cta_Decimo_Cuarto_G")
                      Insertar_Rol_Individual
                      InsValorCtaPro TRol_Pago.Cta, TRol_Pago.Ingresos
                      
                      Limpiar_Rol_Individual
                      TRol_Pago.ID = 2
                      TRol_Pago.Cod_Rol_Pago = "Decimo_IV"
                      TRol_Pago.Tipo_Rubro = "PRO"
                      TRol_Pago.Egresos = Total_DFR
                      TRol_Pago.Cta = AdoClientes.Recordset.Fields("Cta_Decimo_Cuarto_P")
                      TRol_Pago.Codigo_Banco = AdoClientes.Recordset.Fields("Codigo_Banco")
                      Insertar_Rol_Individual
                      InsValorCtaPro TRol_Pago.Cta, -TRol_Pago.Egresos
                   End If
                End If
                
               'IMPUESTO A LA RENTA
               'Datos Informativos de los Gastos deducibles
                Total_IR = Total + Saldo + Total_Otros_Ing - Total_IESS_Per

                If Leer_Cta_Catalogo(Cta_Impuesto_Renta_Empleado) <> Ninguno Then
                   Meses_IR = 0
                   Total_IR_Meses = 0
                    If Month(FechaInicial) > 1 Then
                       sSQL = "SELECT Codigo,COUNT(Codigo) As Meses_Trabajados " _
                            & "FROM Trans_Rol_de_Pagos " _
                            & "WHERE Item = '" & NumEmpresa & "' " _
                            & "AND Periodo = '" & Periodo_Contable & "' " _
                            & "AND Fecha_D < #" & BuscarFecha(FechaInicial) & "# " _
                            & "AND Codigo = '" & TRol_Pago.Codigo & "' " _
                            & "AND Cod_Rol_Pago = 'Salario' " _
                            & "AND Ingresos > 0 " _
                            & "GROUP BY Codigo "
                       SelectAdodc AdoAux, sSQL
                       If AdoAux.Recordset.RecordCount > 0 Then Meses_IR = AdoAux.Recordset.Fields("Meses_Trabajados")
                       
                       sSQL = "SELECT SUM(Ingresos) As Total_Ingresos " _
                            & "FROM Trans_Rol_de_Pagos " _
                            & "WHERE Item = '" & NumEmpresa & "' " _
                            & "AND Periodo = '" & Periodo_Contable & "' " _
                            & "AND Fecha_D < #" & BuscarFecha(FechaInicial) & "# " _
                            & "AND Codigo = '" & TRol_Pago.Codigo & "' " _
                            & "AND Cod_Rol_Pago IN (" & Rubros_Otros_Ingresos & "'Salario','Hora_Extr','Hor_Ext') " _
                            & "AND Ingresos > 0 "
                       SelectAdodc AdoAux, sSQL
                       If AdoAux.Recordset.RecordCount > 0 Then
                          If Not IsNull(AdoAux.Recordset.Fields("Total_Ingresos")) Then Total_IR_Meses = AdoAux.Recordset.Fields("Total_Ingresos")
                       End If
                       Total_IR_Mes_Ant = 0
                       Meses_IR_Mes_Ant = 0
                       sSQL = "SELECT SUM(Egresos) As Total_Egresos, COUNT(Codigo) As Cant_Mes_IR " _
                            & "FROM Trans_Rol_de_Pagos " _
                            & "WHERE Item = '" & NumEmpresa & "' " _
                            & "AND Periodo = '" & Periodo_Contable & "' " _
                            & "AND Fecha_D < #" & BuscarFecha(FechaInicial) & "# " _
                            & "AND Codigo = '" & TRol_Pago.Codigo & "' " _
                            & "AND Cod_Rol_Pago = 'Imp_Renta' " _
                            & "AND Tipo_Rubro = 'PER' " _
                            & "AND Egresos > 0 "
                       SelectAdodc AdoAux, sSQL
                       If AdoAux.Recordset.RecordCount > 0 Then
                          If Not IsNull(AdoAux.Recordset.Fields("Total_Egresos")) Then Total_IR_Mes_Ant = AdoAux.Recordset.Fields("Total_Egresos")
                          If Not IsNull(AdoAux.Recordset.Fields("Cant_Mes_IR")) Then Meses_IR_Mes_Ant = AdoAux.Recordset.Fields("Cant_Mes_IR")
                       End If
                       
                       sSQL = "SELECT SUM(Egresos) As Total_Egresos " _
                            & "FROM Trans_Rol_de_Pagos " _
                            & "WHERE Item = '" & NumEmpresa & "' " _
                            & "AND Periodo = '" & Periodo_Contable & "' " _
                            & "AND Fecha_D < #" & BuscarFecha(FechaInicial) & "# " _
                            & "AND Codigo = '" & TRol_Pago.Codigo & "' " _
                            & "AND Cod_Rol_Pago = 'Aporte_Per' " _
                            & "AND Egresos > 0 "
                       SelectAdodc AdoAux, sSQL
                       If AdoAux.Recordset.RecordCount > 0 Then
                          If Not IsNull(AdoAux.Recordset.Fields("Total_Egresos")) Then
                             Total_IR_Meses = Total_IR_Meses - AdoAux.Recordset.Fields("Total_Egresos")
                          End If
                       End If
                    End If
                    
                   'If TRol_Pago.Codigo = "0915316640" Then MsgBox "(" & Meses_IR & ") Total Meses = " & Total_IR_Meses
                    
                    Total_GP = .Fields("Vivienda") + .Fields("Salud") + .Fields("Educacion") + .Fields("Alimentacion")
                    Total_GP = Total_GP + .Fields("Vestimenta") + .Fields("Discapacidad") + .Fields("Tercera_Edad")
                    
                    NoMeses = (12 - Month(FechaInicial) + 1)
                    Meses_IR = Meses_IR + NoMeses
                    If Meses_IR <= 0 Then Meses_IR = 1
                   'Si el trabajador trabajo menos de 12 meses se prorratea
                    If Meses_IR < 12 Then
                       Total_GP = (Total_GP / 12) * Meses_IR
                    End If
                    Total_GP = Redondear(Total_GP, 2)
                    
                    Total_IR = Total_IR * NoMeses
                    Total_IR = Total_IR + Total_IR_Meses - Total_GP
                   'If TRol_Pago.Codigo = "1307048908" Then MsgBox Total_IR
                    Total_IR = Redondear(Total_IR, 2)
                    
                    SQL2 = "SELECT Desde, Hasta, Basico, Excede " _
                         & "FROM Tabla_Renta " _
                         & "WHERE Año = '" & CStr(Year(FechaInicial)) & "' " _
                         & "AND Desde < " & Total_IR & " " _
                         & "AND " & Total_IR & " <= Hasta " _
                         & "ORDER BY Desde,Hasta "
                    SelectAdodc AdoImpRenta, SQL2
                    'Por mes actual
                    If AdoImpRenta.Recordset.RecordCount > 0 Then
                       Total_Desc = Total_IR - AdoImpRenta.Recordset.Fields("Desde")
                      'If TRol_Pago.Codigo = "1302000383" Then MsgBox Total_IR & vbTab & Total_IESS & vbCrLf & Cta_Impuesto_Renta_Empleado & " ..."
                       Total_Desc = Total_Desc * Redondear(AdoImpRenta.Recordset.Fields("Excede") / 100, 2)
                      'If TRol_Pago.Codigo = "1302000383" Then MsgBox Total_IR & vbTab & Total_IESS & vbCrLf & Cta_Impuesto_Renta_Empleado & " ..."
                       Total_Desc = Redondear((Total_Desc + AdoImpRenta.Recordset.Fields("Basico")) / Meses_IR, 2)
                      'If TRol_Pago.Codigo = "1302000383" Then MsgBox Total_IR & vbTab & Total_IESS & vbCrLf & Cta_Impuesto_Renta_Empleado & " ..."
                    End If
                   'Diferencia del cambio de sueldo si existiera aumento de sueldo
                    Diferencia = 0
                    If 12 > Meses_IR_Mes_Ant Then Diferencia = ((Total_Desc * Meses_IR_Mes_Ant) - Total_IR_Mes_Ant) / (12 - Meses_IR_Mes_Ant)
                    If Diferencia > 0 Then Total_Desc = Total_Desc + Diferencia
                   'If TRol_Pago.Codigo = "1302000383" Then MsgBox Total_IR & vbTab & Total_IESS & vbCrLf & Cta_Impuesto_Renta_Empleado & vbCrLf & "I.R. = " & Total_Desc & " ..."
                    If Total_Desc > 0 Then
                       Limpiar_Rol_Individual
                       TRol_Pago.ID = 18
                       TRol_Pago.Cod_Rol_Pago = "Imp_Renta"
                       TRol_Pago.Tipo_Rubro = "PER"
                       TRol_Pago.Cta = Cta_Impuesto_Renta_Empleado
                       TRol_Pago.Egresos = Total_Desc
                       Insertar_Rol_Individual
                       InsValorCtaRol TRol_Pago.Cta, -TRol_Pago.Egresos
                       Haber = Haber + TRol_Pago.Egresos
                    End If
                End If
               
               'Neto a Recibir del Sueldo
                Cuenta_No = AdoClientes.Recordset.Fields("Cta_Transferencia")
                TipoCta = AdoClientes.Recordset.Fields("TC")
                TipoProc = AdoClientes.Recordset.Fields("FP")
                Limpiar_Rol_Individual
                TRol_Pago.ID = 255
                TRol_Pago.Cod_Rol_Pago = "Neto_Recibir"
                TRol_Pago.Tipo_Rubro = "PER"
                TRol_Pago.Detalle = "TOTAL A RECIBIR"
                TRol_Pago.Egresos = Debe - Haber
                TRol_Pago.Cta = AdoClientes.Recordset.Fields("Cta_Forma_Pago")
                Select Case TipoProc
                  Case "E": TRol_Pago.Cheq_Dep_Transf = "EFECTIVO"
                  Case "C": TRol_Pago.Cheq_Dep_Transf = "Chq. No. " & Format(Cheque_No, "00000000")
                            Cheque_No = Cheque_No + 1
                  Case "T": TRol_Pago.Cheq_Dep_Transf = Cuenta_No
                  Case "O": TRol_Pago.Cheq_Dep_Transf = Ninguno
                End Select
                If CheqCxP.value <> 0 Then
                   TRol_Pago.Cta = SinEspaciosIzq(DCCxP.Text)
                  'TRol_Pago.Cheq_Dep_Transf = "CP" & CStr(Year(fechafinal) & Format(Month(fechafinal), "00"))
                   Cta = Leer_Cta_Catalogo(TRol_Pago.Cta)
                   If Cta <> Ninguno And SubCta = "P" Then
                      SQL2 = "SELECT * " _
                           & "FROM Catalogo_CxCxP " _
                           & "WHERE Item = '" & NumEmpresa & "' " _
                           & "AND TC = '" & SubCta & "' " _
                           & "AND Cta = '" & Cta & "' " _
                           & "AND Codigo = '" & CodigoCliente & "' "
                      SelectAdodc AdoAsientoSC, SQL2
                      If AdoAsientoSC.Recordset.RecordCount <= 0 Then
                         SetAdoAddNew "Catalogo_CxCxP"
                         SetAdoFields "Codigo", CodigoCliente
                         SetAdoFields "Cta", Cta
                         SetAdoFields "TC", SubCta
                         SetAdoUpdate
                      End If
                      SetAdoAddNew "Asiento_SC"
                      SetAdoFields "Codigo", CodigoCliente
                      SetAdoFields "Beneficiario", NombreCliente
                      SetAdoFields "Cta", Cta
                      SetAdoFields "Valor", TRol_Pago.Egresos
                      SetAdoFields "FECHA_V", FechaFinal
                      If Len(TRol_Pago.Cheq_Dep_Transf) > 1 Then
                         SetAdoFields "Detalle_SubCta", "SxP: " & TRol_Pago.Cheq_Dep_Transf
                      Else
                         SetAdoFields "Detalle_SubCta", "SxP"
                      End If
                      SetAdoFields "TC", SubCta
                      SetAdoFields "DH", "2"
                      SetAdoFields "TM", "1"
                      SetAdoFields "T_No", Trans_No
                      SetAdoFields "SC_No", LnSC_No
                      SetAdoFields "CodigoU", CodigoUsuario
                      SetAdoUpdate
                      LnSC_No = LnSC_No + 1
                   End If
                End If
                Insertar_Rol_Individual
                Debe = 0: Haber = 0: Diferencia = 0
             End If
          End If
         .MoveNext
       Loop
   End If
  End With
  Listar_CxCxP_SubMod
''  Contador = 0
''  With AdoAsientoSC.Recordset
''   If .RecordCount > 0 Then
''       Do While Not .EOF
''          Contador = Contador + 1
''         .Fields("SC_No") = Contador
''         .MoveNext
''       Loop
''      .UpdateBatch
''   End If
''  End With
  Listar_CxCxP_SubMod
End Sub

Public Sub Llenar_Rol_Pagos_Colectivo(Es_quincena As Boolean)
  DGNomina.Visible = False
  DGTotNomina.Visible = False
 'Borrarmos la tabla temporal del rol
  SQL1 = "DELETE * " _
       & "FROM Asiento_Rol_Colectivo " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' "
  ConectarAdoExecute SQL1
 'Insertamos los Empleados del rol
  Contador = 0
  With AdoClientes.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       Do While Not .EOF
          Contador = Contador + 1
          SetAdoAddNew "Asiento_Rol_Colectivo"
          SetAdoFields "No_", Format(Contador, "00")
          SetAdoFields "Codigo", .Fields("Codigo")
          SetAdoFields "C_I", .Fields("CI_RUC")
          SetAdoFields "Nombre_Empleado", .Fields("Cliente")
          SetAdoFields "Grupo_Rol", .Fields("Grupo_Rol")
          SetAdoFields "Fecha", FechaFinal
          SetAdoFields "Porc_Apo_Pat", .Fields("IEESS_Pat")
          SetAdoFields "Porc_Apo_Per", .Fields("IEESS_Per")
          SetAdoFields "I", ""
          SetAdoFields "II", ""
          SetAdoFields "Item", NumEmpresa
          SetAdoFields "CodigoU", CodigoUsuario
          SetAdoFields "Grupo_Rol", .Fields("Grupo_Rol")
          SetAdoFields "Fecha_Ing", .Fields("Fecha")
          SetAdoFields "FR", .Fields("Pagar_Fondo_Reserva")
          SetAdoUpdate
         .MoveNext
       Loop
   End If
  End With
 'Llenamos todos los ingresos
  sSQL = "SELECT Cod_Rol_Pago " _
       & "FROM Trans_Rol_de_Pagos " _
       & "WHERE Fecha_D >= #" & FechaIni & "# " _
       & "AND Fecha_H <= #" & FechaFin & "# " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND Tipo_Rubro = 'PER' " _
       & "AND Ingresos > 0 " _
       & "GROUP BY Cod_Rol_Pago " _
       & "ORDER BY Cod_Rol_Pago "
  SelectAdodc AdoAux, sSQL
  SQL1 = "SELECT No_,C_I,Nombre_Empleado,Grupo_Rol,Dias,Fecha_Ing,FR,Horas,Cheque_No"
  SQL2 = "SELECT Grupo_Rol"
  SQL3 = "SELECT No_,C_I,Nombre_Empleado"
  With AdoAux.Recordset
   If .RecordCount > 0 Then
       I = 1
       Do While Not .EOF
          SQL1 = SQL1 & ",Ing_" & Format(I, "00") & " As " & .Fields("Cod_Rol_Pago")
          SQL2 = SQL2 & ",SUM(Ing_" & Format(I, "00") & ") As " & .Fields("Cod_Rol_Pago")
          I = I + 1
         .MoveNext
       Loop
   End If
  End With
 'Llenamos todos los Egresos
  sSQL = "SELECT Cod_Rol_Pago " _
       & "FROM Trans_Rol_de_Pagos " _
       & "WHERE Fecha_D >= #" & FechaIni & "# " _
       & "AND Fecha_H <= #" & FechaFin & "# " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND Tipo_Rubro = 'PER' " _
       & "AND Egresos > 0 " _
       & "GROUP BY Cod_Rol_Pago " _
       & "ORDER BY Cod_Rol_Pago "
  SelectAdodc AdoAux, sSQL
 'Llenamos todos los Egresos
  SQL1 = SQL1 & ",I"
  With AdoAux.Recordset
   If .RecordCount > 0 Then
       I = 1
       Do While Not .EOF
          If .Fields("Cod_Rol_Pago") <> "Neto_Recibir" Then
              SQL1 = SQL1 & ",Egr_" & Format(I, "00") & " As " & .Fields("Cod_Rol_Pago")
              SQL2 = SQL2 & ",SUM(Egr_" & Format(I, "00") & ") As " & .Fields("Cod_Rol_Pago")
              I = I + 1
          End If
         .MoveNext
       Loop
   End If
  End With
  sSQL = "SELECT Cod_Rol_Pago " _
       & "FROM Trans_Rol_de_Pagos " _
       & "WHERE Fecha_D >= #" & FechaIni & "# " _
       & "AND Fecha_H <= #" & FechaFin & "# " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND Tipo_Rubro IN ('PAT','PRO') " _
       & "AND Egresos > 0 " _
       & "GROUP BY Cod_Rol_Pago " _
       & "ORDER BY Cod_Rol_Pago "
  SelectAdodc AdoAux, sSQL
  SQL1 = SQL1 & ",Neto_Recibir,Firma,II,Aporte_Pat"
  With AdoAux.Recordset
   If .RecordCount > 0 Then
       I = 1
       Do While Not .EOF
          'MsgBox I & vbCrLf & .Fields("Cod_Rol_Pago")
          SQL1 = SQL1 & ",Inf_" & Format(I, "00") & " As " & .Fields("Cod_Rol_Pago")
          SQL2 = SQL2 & ",SUM(Inf_" & Format(I, "00") & ") As " & .Fields("Cod_Rol_Pago")
          SQL3 = SQL3 & ",Inf_" & Format(I, "00") & " As " & .Fields("Cod_Rol_Pago")
          I = I + 1
         'MsgBox I & " - " & .Fields("Cod_Rol_Pago")
         .MoveNext
       Loop
   End If
  End With
  SQL1 = SQL1 & ",Codigo,CodigoU " _
       & "FROM Asiento_Rol_Colectivo " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' "
  If CmbGrupos <> "TODOS" Then SQL1 = SQL1 & "AND Grupo_Rol = '" & CmbGrupos & "' "
  If OpcGrupo.value Then
     SQL1 = SQL1 & "ORDER BY Grupo_Rol,No_,Nombre_Empleado "
  Else
     SQL1 = SQL1 & "ORDER BY Nombre_Empleado "
  End If
  
  SQL2 = SQL2 & " FROM Asiento_Rol_Colectivo " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' "
  If CmbGrupos <> "TODOS" Then SQL2 = SQL2 & "AND Grupo_Rol = '" & CmbGrupos & "' "
  SQL2 = SQL2 & "GROUP BY Grupo_Rol " _
       & "ORDER BY Grupo_Rol "
  
  SQL3 = SQL3 & ",Porc_Apo_Pat As IESS_Pa,Porc_Apo_Per As IESS_Pe,Grupo_Rol " _
       & "FROM Asiento_Rol_Colectivo " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' "
  If CmbGrupos <> "TODOS" Then SQL3 = SQL3 & "AND Grupo_Rol = '" & CmbGrupos & "' "
  If OpcGrupo.value Then
     SQL3 = SQL3 & "ORDER BY Grupo_Rol,No_,Nombre_Empleado "
  Else
     SQL3 = SQL3 & "ORDER BY Nombre_Empleado "
  End If
  'MsgBox SQL1
  SelectDataGrid DGNomina, AdoNomina, SQL1
 'Guardamos temporalmente el tipo de consulta para presentar el rol lleno
 'Empezamos a llenar el rol colectivo
  Contador = 0
  sSQL = "SELECT * " _
       & "FROM Trans_Rol_de_Pagos " _
       & "WHERE Fecha_D >= #" & FechaIni & "# " _
       & "AND Fecha_H <= #" & FechaFin & "# " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "ORDER BY Periodo,T,Grupo_Rol,Codigo,Tipo_Rubro,ID "
  SelectAdodc AdoAux, sSQL
  With AdoAux.Recordset
   If .RecordCount > 0 Then
       Progreso_Barra.Valor_Maximo = Progreso_Barra.Valor_Maximo + .RecordCount
       Do While Not .EOF
          Contador = Contador + 1
          Progreso_Barra.Mensaje_Box = "ROL COLECTIVO DEL MES DE"
          Progreso_Esperar
          CodigoCli = .Fields("Codigo")
          Codigo = .Fields("Cod_Rol_Pago")
          If AdoNomina.Recordset.RecordCount > 0 Then
             AdoNomina.Recordset.MoveFirst
             AdoNomina.Recordset.Find ("Codigo = '" & CodigoCli & "' ")
             If Not AdoNomina.Recordset.EOF Then
               'MsgBox AdoNomina.Recordset.Fields(Codigo)
                If .Fields("Ingresos") <> 0 Then Valor = .Fields("Ingresos")
                If .Fields("Egresos") <> 0 Then Valor = .Fields("Egresos")
                If .Fields("Dias") <> 0 Then AdoNomina.Recordset.Fields("Dias") = .Fields("Dias")
                If .Fields("Cheq_Dep_Transf") <> Ninguno Then AdoNomina.Recordset.Fields("Cheque_No") = SinEspaciosDer(.Fields("Cheq_Dep_Transf"))
                If .Fields("Horas") <> 0 Then AdoNomina.Recordset.Fields("Horas") = .Fields("Horas")
                AdoNomina.Recordset.Fields(Codigo) = Valor
                AdoNomina.Recordset.Update
             End If
          End If
         .MoveNext
       Loop
   End If
  End With
  DGNomina.Visible = True
  DGTotNomina.Visible = True
  SelectDataGrid DGNomina, AdoNomina, SQL1
  SelectDataGrid DGNominaProv, AdoNominaProv, SQL3
  SelectDataGrid DGTotNomina, AdoTotNomina, SQL2
End Sub

Public Sub Listar_CxCxP_SubMod()
   Trans_No = 100
   SQL2 = "SELECT * " _
        & "FROM Asiento_SC " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND T_No = " & Trans_No & " " _
        & "AND CodigoU = '" & CodigoUsuario & "' " _
        & "ORDER BY Beneficiario, DH, Cta, SC_No "
   SelectDataGrid DGSubCtas, AdoAsientoSC, SQL2
End Sub

Public Sub Encabezado_Rol()
Dim Ancho_Maximo As Single
 PosLinea = 1
 Ancho_Maximo = cPrint.dAnchoPapel - 0.5
 cPrint.printImagen LogoTipo, 1, PosLinea, 4.5, 2
 RutaDestino = RutaSistema & "\LOGOS\DiskCover.gif"
 cPrint.printImagen RutaDestino, Ancho_Maximo - 1.8, PosLinea, 1.8, 0.6
 cPrint.letraTipo TipoHelvetica, 7
 cPrint.tipoNegrilla = True
 cPrint.printTexto Ancho_Maximo - 4.5, PosLinea, "Hora:"
 cPrint.printTexto Ancho_Maximo - 4.5, PosLinea + 0.3, "Pagina No."
 cPrint.printTexto Ancho_Maximo - 4.5, PosLinea + 0.6, "Fecha:"
 cPrint.printTexto Ancho_Maximo - 4.5, PosLinea + 0.9, "Usuario:"
 cPrint.tipoNegrilla = False
 cPrint.printTexto Ancho_Maximo - 3.2, PosLinea, Format(Time, "hh:mm:ss")
 cPrint.printTexto Ancho_Maximo - 3.2, PosLinea + 0.3, Format(Pagina, "0000")
 cPrint.printTexto Ancho_Maximo - 3.2, PosLinea + 0.6, FechaStrgDias(date)
 cPrint.printTexto Ancho_Maximo - 3.2, PosLinea + 0.9, ULCase(NombreUsuario)
 cPrint.letraTipo TipoTimes
 cPrint.tipoNegrilla = True
 cPrint.porteDeLetra = 14
 If UCase$(RazonSocial) = UCase$(NombreComercial) Then
    cPrint.printTexto 1, PosLinea, UCase$(RazonSocial), "C", Ancho_Maximo
 Else
    cPrint.printTexto 1, PosLinea, UCase$(RazonSocial), "C", Ancho_Maximo
    cPrint.printTexto 1, PosLinea + 0.5, UCase$(NombreComercial), "C", Ancho_Maximo
 End If
 PosLinea = PosLinea + 1
 cPrint.porteDeLetra = 9
 cPrint.tipoNegrilla = False
 cPrint.printTexto 1, PosLinea, ULCase(Direccion) & ". Teléfono: " & Telefono1, "C", Ancho_Maximo
 PosLinea = PosLinea + 0.5
 cPrint.porteDeLetra = 12
 cPrint.tipoNegrilla = True
 cPrint.printTexto 1, PosLinea, MensajeEncabData, "C", Ancho_Maximo
 cPrint.tipoNegrilla = False
 cPrint.porteDeLetra = 8
 Pagina = Pagina + 1
 cPrint.letraTipo TipoHelvetica
 PosLinea = PosLinea + 0.6
End Sub

Public Sub Imprimir_Rol_Colectivo(Datas As Adodc, _
                                  DatasT As Adodc, _
                                  Optional EsCampoCorto As Boolean)
On Error GoTo Errorhandler
Dim SizeLetra As Integer
Dim AnchoPict As Single
Dim AltoPict As Single
Dim PosLineaTemp As Single
Dim X_Max As Single
Dim Y_Max As Single
Dim NombFilePict As String
Dim TotValores(30) As Double
Dim CantCamposTemp As Integer

Mensajes = "Seguro de Imprimir en:" & vbCrLf & Printer.DeviceName & "?"
Titulo = "IMPRESION"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
   RatonReloj
  'Generamos el documento
   NombFilePict = "Rol Pagos Colectivo " & CAnio & "-" & Format$(NumMeses, "00") & " R-" & RUC & " " & CodigoUsuario
   tPrint.TipoImpresion = Es_Printer
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
 Ancho(3) = 6    'Grupo_No
 Ancho(4) = 7.45 'Días
 Ancho(5) = 8.1  'Fecha_Ing
 Ancho(6) = 9.45 'FR
 Ancho(7) = 9.8  'Horas
 Ancho(8) = 10.5 'Cheque_No
 
 Pagina = 1
 PosLinea = 1
'Iniciamos la impresion
 cPrint.tipoNegrilla = False
 With Datas.Recordset
  If .RecordCount > 0 Then
     .MoveFirst
     'Ancho(6) = Salario
      Distancia = Ancho(8) + 1.5
      J = CantCampos
      For I = 9 To CantCampos - 1
          If .Fields(I).Name = "II" Then J = I
      Next I
      CantCampos = J
      For I = 9 To CantCampos - 1
          If .Fields(I).Name = "I" Or .Fields(I).Name = "II" Then
              Ancho(I) = Distancia
              Distancia = Distancia + 0.05
          ElseIf .Fields(I).Name = "Firma" Then
              Ancho(I) = Distancia
              Distancia = Distancia + 5
          ElseIf .Fields(I).Name = "CodigoU" Then
              Ancho(I) = Distancia
              Distancia = Distancia + 7
          Else
              Ancho(I) = Distancia
              Distancia = Distancia + 1.6
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
      cPrint.porteDeLetra = SizeLetra
      cPrint.tipoNegrilla = True
      For I = 0 To CantCampos - 1
          Select Case .Fields(I).Name
            Case "Codigo", "CodigoU", "I", "II": 'cPrint.printTexto  Ancho(I) + Distancia, PosLinea, ""
            Case Else: cPrint.printTexto Ancho(I), PosLinea, .Fields(I).Name
          End Select
      Next I
      PosLinea = PosLinea + 0.4
      PosLineaTemp = PosLinea
      cPrint.printCuadroLinea Ancho(0), PosLinea, Ancho(CantCampos), PosLinea
      PosLinea = PosLinea + 0.05
      cPrint.tipoNegrilla = False
      Do While Not .EOF
         cPrint.tipoNegrilla = True
         For I = 0 To CantCampos - 1
             Distancia = cPrint.anchoFields(.Fields(I), 2)
             'Distancia = CampoWidth(.Fields(I))
             If cPrint.dStrgFormatoCampo = Ninguno Then
                cPrint.dStrgFormatoCampo = " "
             ElseIf cPrint.dStrgFormatoCampo = "0" Or cPrint.dStrgFormatoCampo = "0.00" Then
                cPrint.dStrgFormatoCampo = " "
             End If
             Select Case .Fields(I).Name
               Case "Codigo", "CodigoU", "I", "II": ' cPrint.printTexto  Ancho(I) + Distancia, PosLinea, ""
               Case "Nombre_Empleado"
                    cPrint.printTexto Ancho(I) + Distancia, PosLinea, Extraer_Apellidos(cPrint.dStrgFormatoCampo)
                    cPrint.printTexto Ancho(I) + Distancia, PosLinea + 0.3, Extraer_Nombres(cPrint.dStrgFormatoCampo)
               Case Else: cPrint.printTexto Ancho(I) + Distancia, PosLinea, cPrint.dStrgFormatoCampo
             End Select
         Next I
         PosLinea = PosLinea + 0.7
         cPrint.printCuadroLinea Ancho(0) - 0.1, PosLinea, Ancho(CantCampos) - 0.1, PosLinea
         PosLinea = PosLinea + 0.05
         If PosLinea >= LimiteAlto - 0.6 Then
            For I = 0 To CantCampos - 1
             Select Case .Fields(I).Name
               Case "Codigo", "CodigoU", "I", "II": ' cPrint.printTexto  Ancho(I) + Distancia, PosLinea, ""
               Case Else: cPrint.printCuadroLinea Ancho(I) - 0.1, 2.8, Ancho(I) - 0.1, PosLinea
             End Select
            Next I
            cPrint.printCuadroLinea Ancho(0) - 0.1, PosLinea, Ancho(CantCampos) - 0.1, PosLinea
            cPrint.paginaNueva
            Encabezado_Rol
            cPrint.porteDeLetra = SizeLetra
            cPrint.tipoNegrilla = True
            For I = 0 To CantCampos - 1
                Select Case .Fields(I).Name
                  Case "Codigo", "CodigoU", "I", "II": 'cPrint.printTexto  Ancho(I) + Distancia, PosLinea, ""
                  Case Else: cPrint.printTexto Ancho(I), PosLinea, .Fields(I).Name
                End Select
            Next I
            PosLinea = PosLinea + 0.4
            cPrint.printCuadroLinea Ancho(0) - 0.1, PosLinea, Ancho(CantCampos) - 0.1, PosLinea
            PosLinea = PosLinea + 0.1
            cPrint.tipoNegrilla = False
         End If
        .MoveNext
      Loop
      For I = 0 To CantCampos - 1
       Select Case .Fields(I).Name
         Case "Codigo", "CodigoU", "I", "II": ' cPrint.printTexto  Ancho(I) + Distancia, PosLinea, ""
         Case Else: cPrint.printCuadroLinea Ancho(I) - 0.1, PosLineaTemp, Ancho(I) - 0.1, PosLinea
       End Select
      Next I
  End If
 End With
 cPrint.printCuadroLinea Ancho(CantCampos) - 0.1, PosLineaTemp, Ancho(CantCampos) - 0.1, PosLinea
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
          InicioX = InicioX + 1.6
      Next I
      Ancho(I) = InicioX
      cPrint.porteDeLetra = SizeLetra
      cPrint.tipoNegrilla = True
      For I = 0 To CantCampos - 1
          cPrint.printTexto Ancho(I), PosLinea, .Fields(I).Name
          TotValores(I) = 0
      Next I
      PosLinea = PosLinea + 0.35
      PosLineaTemp = PosLinea
      cPrint.printCuadroLinea Ancho(0), PosLinea, Ancho(CantCampos), PosLinea
      PosLinea = PosLinea + 0.05
      cPrint.tipoNegrilla = False
      
      Do While Not .EOF
         For I = 0 To CantCampos - 1
             Distancia = cPrint.anchoFields(.Fields(I), 2)
             If cPrint.dStrgFormatoCampo = Ninguno Then
                cPrint.dStrgFormatoCampo = " "
             ElseIf cPrint.dStrgFormatoCampo = "0" Or cPrint.dStrgFormatoCampo = "0.00" Then
                cPrint.dStrgFormatoCampo = " "
             End If
             cPrint.printTexto Ancho(I) + Distancia, PosLinea, cPrint.dStrgFormatoCampo
             If I <> 0 Then TotValores(I) = TotValores(I) + .Fields(I)
         Next I
         PosLinea = PosLinea + 0.4
         cPrint.printCuadroLinea Ancho(0), PosLinea, Ancho(CantCampos), PosLinea
         PosLinea = PosLinea + 0.05
         If PosLinea >= LimiteAlto Then
            For I = 0 To CantCampos
                cPrint.printCuadroLinea Ancho(I) - 0.1, PosLineaTemp, Ancho(I) - 0.1, PosLinea + 0.1
            Next I
            
            cPrint.paginaNueva
            Encabezado_Rol
            PosLineaTemp = PosLinea + 0.2
            cPrint.porteDeLetra = SizeLetra
            cPrint.tipoNegrilla = True
            For I = 0 To CantCampos - 1
                cPrint.printTexto Ancho(I), PosLinea, .Fields(I).Name
            Next I
            PosLinea = PosLinea + 0.3
            cPrint.printCuadroLinea Ancho(0) - 0.1, PosLinea + 0.1, Ancho(CantCampos), PosLinea + 0.1
            PosLinea = PosLinea + 0.1
            cPrint.tipoNegrilla = False
         End If
        .MoveNext
      Loop
      For I = 0 To CantCampos
          cPrint.printCuadroLinea Ancho(I) - 0.1, PosLineaTemp, Ancho(I) - 0.1, PosLinea
      Next I
  End If
 End With
 cPrint.tipoNegrilla = True
 cPrint.printTexto Ancho(0), PosLinea, "T O T A L E S"
 For I = 1 To CantCampos - 1
     cPrint.printVariable Ancho(I), PosLinea, TotValores(I)
 Next I
 cPrint.porteDeLetra = SizeLetra
 CantCamposTemp = CantCampos
 sSQL = "SELECT C.Cliente,TR.* " _
      & "FROM Trans_Entrada_Salida As TR,Clientes As C " _
      & "WHERE TR.Fecha BETWEEN #" & BuscarFecha(FechaInicial) & "# AND #" & BuscarFecha(FechaFinal) & "# " _
      & "AND TR.Codigo = C.Codigo " _
      & "AND TR.Item = '" & NumEmpresa & "' " _
      & "AND TR.Periodo = '" & Periodo_Contable & "' " _
      & "AND TR.ES = 'R' " _
      & "ORDER BY C.Cliente,TR.Fecha,TR.Hora "
 SelectAdodc AdoNovedades, sSQL
 With AdoNovedades.Recordset
  If .RecordCount > 0 Then
      cPrint.tipoNegrilla = True
      PosLinea = PosLinea + 0.4
      cPrint.printTexto Ancho(0), PosLinea, "OBSERVACIONES:"
      cPrint.tipoNegrilla = False
      PosLinea = PosLinea + 0.35
      Do While Not .EOF
         cPrint.printTexto Ancho(0), PosLinea, .Fields("Tarea")
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
  Presentar_PDF fPDF
Exit Sub
Errorhandler:
    RatonNormal
    ErrorDeImpresion
    Exit Sub
Else
    RatonNormal
End If
End Sub

Public Sub Saldos_CxC_CxP(Codigo_Persona As String, TC As String)
Dim Fecha_Rol_I As String
Dim Fecha_Rol_F As String
Dim ContSC As Byte
  
   Fecha_Rol_I = BuscarFecha(FechaInicial)
   Fecha_Rol_F = BuscarFecha(FechaFinal)
   ContSC = 0
'''  'Saldos Pendientes antes del mes
'''   If TC = "C" Then
'''      sSQL = "SELECT TC,Cta,Factura,SUM(Debitos-Creditos) As TSaldo "
'''   Else
'''      sSQL = "SELECT TC,Cta,Factura,SUM(Creditos-Debitos) As TSaldo "
'''   End If
'''   sSQL = sSQL _
'''        & "FROM Trans_SubCtas " _
'''        & "WHERE Fecha < #" & Fecha_Rol_I & "# " _
'''        & "AND T <> '" & Anulado & "' " _
'''        & "AND TC = '" & TC & "' " _
'''        & "AND Item = '" & NumEmpresa & "' " _
'''        & "AND Periodo = '" & Periodo_Contable & "' " _
'''        & "AND Codigo = '" & Codigo_Persona & "' " _
'''        & "GROUP BY TC,Cta,Factura " _
'''        & "ORDER BY TC,Cta,Factura "
'''   SelectAdodc AdoSubCta1, sSQL
'''   With AdoSubCta1.Recordset
'''    If .RecordCount > 0 Then
'''        Saldo = 0
'''        TipoDoc = .Fields("TC")
'''        SubCtaGen = .Fields("Cta")
'''        Factura_No = .Fields("Factura")
'''        Do While Not .EOF
'''           If SubCtaGen <> .Fields("Cta") Or TipoDoc <> .Fields("TC") Or Factura_No <> .Fields("Factura") Then
'''              If Saldo > 0 Then
'''                 ContSC = ContSC + 1
'''                 Total = Total + Saldo
'''                 InsertarCxCxP CodigoCliente, SubCtaGen, Saldo, TipoDoc
'''                 LeerCta SubCtaGen
'''                 If Codigo = Ninguno Then
'''                    Si_No = True
'''                    Cadena1 = Cadena1 & SubCtaGen & vbCrLf
'''                 End If
'''              End If
'''              Saldo = 0
'''              TipoDoc = .Fields("TC")
'''              SubCtaGen = .Fields("Cta")
'''              Factura_No = .Fields("Factura")
'''           End If
'''           Saldo = Saldo + Redondear(.Fields("TSaldo"), 2)
'''          .MoveNext
'''        Loop
'''
'''        If Saldo > 0 Then
'''           Total = Total + Saldo
'''           InsertarCxCxP CodigoCliente, SubCtaGen, Saldo, TipoDoc
'''           LeerCta SubCtaGen
'''           If Codigo = Ninguno Then
'''              Si_No = True
'''              Cadena1 = Cadena1 & SubCtaGen & vbCrLf
'''           End If
'''        End If
'''    End If
'''   End With

  'Saldos Pendientes del mes
   'BETWEEN #" & Fecha_Rol_I & "# and
   If TC = "C" Then
      sSQL = "SELECT TC,Cta,Factura,(SUM(Debitos)-SUM(Creditos)) As TSaldo "
   Else
      sSQL = "SELECT TC,Cta,Factura,(SUM(Creditos)-SUM(Debitos)) As TSaldo "
   End If
   sSQL = sSQL _
        & "FROM Trans_SubCtas " _
        & "WHERE Fecha_V <= #" & Fecha_Rol_F & "# " _
        & "AND T <> '" & Anulado & "' " _
        & "AND TC = '" & TC & "' " _
        & "AND Item = '" & NumEmpresa & "' " _
        & "AND Periodo = '" & Periodo_Contable & "' " _
        & "AND Codigo = '" & Codigo_Persona & "' " _
        & "AND LEN(Cta) > 3 " _
        & "GROUP BY TC,Cta,Factura "
   If TC = "C" Then
      sSQL = sSQL _
           & "HAVING SUM(Debitos)-SUM(Creditos) > 0 "
   Else
      sSQL = sSQL _
           & "HAVING SUM(Creditos)-SUM(Debitos) > 0 "
   End If
   sSQL = sSQL _
        & "ORDER BY TC,Cta,Factura "
   SelectAdodc AdoSubCta1, sSQL
   With AdoSubCta1.Recordset
   'If Codigo_Persona = "1303410367" Then MsgBox sSQL & vbCrLf & vbCrLf & "Registros: " & .RecordCount
    If .RecordCount > 0 Then
        Saldo = 0
        TipoDoc = .Fields("TC")
        SubCtaGen = .Fields("Cta")
        Factura_No = .Fields("Factura")
        Do While Not .EOF
           If SubCtaGen <> .Fields("Cta") Or TipoDoc <> .Fields("TC") Or Factura_No <> .Fields("Factura") Then
              'If Codigo_Persona = "0400731824" Then MsgBox "Registro Medio: " & SubCtaGen & ", Saldo: " & Saldo
              If Saldo > 0 Then
                 Total = Total + Saldo
                 InsertarCxCxP CodigoCliente, SubCtaGen, Saldo, TipoDoc
                 LeerCta SubCtaGen
                 If Codigo = Ninguno Then
                    Si_No = True
                    Cadena1 = Cadena1 & SubCtaGen & vbCrLf
                 End If
              End If
              Saldo = 0
              TipoDoc = .Fields("TC")
              SubCtaGen = .Fields("Cta")
              Factura_No = .Fields("Factura")
           End If
           Saldo = Saldo + .Fields("TSaldo")
          .MoveNext
        Loop
       'If Codigo_Persona = "1308498649" Then MsgBox "Ultimo Registro: " & SubCtaGen & ", Saldo: " & Saldo
        If Saldo > 0 Then
           Total = Total + Saldo
           InsertarCxCxP CodigoCliente, SubCtaGen, Saldo, TipoDoc
           LeerCta SubCtaGen
           If Codigo = Ninguno Then
              Si_No = True
              Cadena1 = Cadena1 & SubCtaGen & vbCrLf
           End If
        End If
    End If
   End With
End Sub

Public Sub Inicializar_Cero_Asientos(EnCero As Boolean)
   'Inicializamos los Asientos de submodulos
    If EnCero Then
       SQL2 = "DELETE * " _
            & "FROM Asiento " _
            & "WHERE Item = '" & NumEmpresa & "' " _
            & "AND T_No = 100 " _
            & "AND CodigoU = '" & CodigoUsuario & "' "
       ConectarAdoExecute SQL2
       SQL2 = "DELETE * " _
            & "FROM Asiento " _
            & "WHERE Item = '" & NumEmpresa & "' " _
            & "AND T_No = 101 " _
            & "AND CodigoU = '" & CodigoUsuario & "' "
       ConectarAdoExecute SQL2
       SQL2 = "DELETE * " _
            & "FROM Asiento " _
            & "WHERE Item = '" & NumEmpresa & "' " _
            & "AND T_No = 102 " _
            & "AND CodigoU = '" & CodigoUsuario & "' "
       ConectarAdoExecute SQL2
    End If
    
   'Presentamos el resultado de los asientos
    Trans_No = 101
    SQL2 = "SELECT * " _
         & "FROM Asiento " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND T_No = " & Trans_No & " " _
         & "AND CodigoU = '" & CodigoUsuario & "' "
    SelectDataGrid DGAsiento(1), AdoAsiento1, SQL2
    
    Trans_No = 102
    SQL2 = "SELECT * " _
         & "FROM Asiento " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND T_No = " & Trans_No & " " _
         & "AND CodigoU = '" & CodigoUsuario & "' "
    SelectDataGrid DGAsiento(2), AdoAsiento2, SQL2
    
    Trans_No = 100
    SQL2 = "SELECT * " _
         & "FROM Asiento " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND T_No = " & Trans_No & " " _
         & "AND CodigoU = '" & CodigoUsuario & "' "
    SelectDataGrid DGAsiento(0), AdoAsiento, SQL2
End Sub

