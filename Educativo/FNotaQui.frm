VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.4#0"; "Comctl32.ocx"
Begin VB.Form FNotasQuimestre 
   BackColor       =   &H00C0FFFF&
   Caption         =   "CUADROS DE NOTAS"
   ClientHeight    =   8745
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12675
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "FNotaQui.frx":0000
   ScaleHeight     =   8745
   ScaleWidth      =   12675
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12675
      _ExtentX        =   22357
      _ExtentY        =   1164
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "ImgLstMenu"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   16
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Salir"
            Object.ToolTipText     =   "Salir del Módulo"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Imprimir"
            Object.ToolTipText     =   "Imprimir lo que presenta en pantalla"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Notas"
            Object.ToolTipText     =   "Resumen Notas"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Disciplina"
            Object.ToolTipText     =   "Disciplina"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "MNotaQuimestre"
            Object.ToolTipText     =   "Notas Quimestrales por Parcial"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Listas"
            Object.ToolTipText     =   "Listas"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "ListasMes"
            Object.ToolTipText     =   "Lista de Alumnios por meses"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "ListaAlumnos"
            Object.ToolTipText     =   "Lista de Alumnos"
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "NotasProfesor"
            Object.ToolTipText     =   "Imprime resumen de materia por profesor"
            Object.Tag             =   ""
            ImageIndex      =   9
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "MejorPromedio"
            Object.ToolTipText     =   "Mejor Promedio"
            Object.Tag             =   ""
            ImageIndex      =   11
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "MMejorEgresado"
            Object.ToolTipText     =   "Notas Actas Grado"
            Object.Tag             =   ""
            ImageIndex      =   12
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "ResumenExamenGrado"
            Object.ToolTipText     =   "Resum. Exam. Grado"
            Object.Tag             =   ""
            ImageIndex      =   13
         EndProperty
         BeginProperty Button13 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "MNominaOficial"
            Object.ToolTipText     =   "Nomina Oficial"
            Object.Tag             =   ""
            ImageIndex      =   14
         EndProperty
         BeginProperty Button14 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "BPromobidos"
            Object.ToolTipText     =   "Presenta la nomina de Aprobados o Reprobados"
            Object.Tag             =   ""
            ImageIndex      =   16
         EndProperty
         BeginProperty Button15 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "NotasBlanco"
            Object.ToolTipText     =   "Presentar Notas en Blanco"
            Object.Tag             =   ""
            ImageIndex      =   17
         EndProperty
         BeginProperty Button16 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "CuadroSubNotas"
            Object.ToolTipText     =   "Cuadro de Calificaciones de Sub-Notas (Ingles)"
            Object.Tag             =   ""
            ImageIndex      =   18
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSDataListLib.DataList DLCurso 
      Bindings        =   "FNotaQui.frx":0342
      DataSource      =   "AdoAux"
      Height          =   1815
      Left            =   4200
      TabIndex        =   6
      Top             =   1050
      Width           =   5475
      _ExtentX        =   9657
      _ExtentY        =   3201
      _Version        =   393216
      BackColor       =   8454143
      ForeColor       =   12582912
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.ListBox LstPeriodos 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   105
      TabIndex        =   2
      Top             =   1050
      Width           =   4005
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&S"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   210
      TabIndex        =   21
      Top             =   7875
      Width           =   330
   End
   Begin VB.TextBox TxtHasta 
      BackColor       =   &H00FFFFC0&
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
      Left            =   7980
      TabIndex        =   5
      Text            =   "0"
      Top             =   735
      Width           =   1695
   End
   Begin VB.TextBox TxtDesde 
      BackColor       =   &H00FFFFC0&
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
      Left            =   6300
      TabIndex        =   4
      Text            =   "0"
      Top             =   735
      Width           =   1695
   End
   Begin VB.Frame FrmPictCalif 
      BorderStyle     =   0  'None
      Height          =   6000
      Left            =   105
      TabIndex        =   15
      Top             =   3360
      Width           =   12510
      Begin VB.VScrollBar VScroll1 
         Height          =   4845
         Left            =   105
         TabIndex        =   19
         Top             =   105
         Width           =   330
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00E0E0E0&
         Height          =   4950
         Left            =   420
         ScaleHeight     =   8.625
         ScaleMode       =   7  'Centimeter
         ScaleWidth      =   18.441
         TabIndex        =   17
         Top             =   105
         Width           =   10515
         Begin VB.PictureBox PictCalif 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Height          =   4680
            Left            =   0
            ScaleHeight     =   8.149
            ScaleMode       =   7  'Centimeter
            ScaleWidth      =   17.833
            TabIndex        =   18
            Top             =   0
            Width           =   10170
         End
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   330
         Left            =   1785
         TabIndex        =   16
         Top             =   5460
         Width           =   9150
      End
      Begin VB.Label LblA4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0x0"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   420
         TabIndex        =   20
         Top             =   5460
         Width           =   1380
      End
   End
   Begin MSDataGridLib.DataGrid DGResumenNotas 
      Bindings        =   "FNotaQui.frx":0357
      Height          =   3690
      Left            =   105
      TabIndex        =   14
      Top             =   3360
      Visible         =   0   'False
      Width           =   3060
      _ExtentX        =   5398
      _ExtentY        =   6509
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
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Lista de Notas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9765
      TabIndex        =   9
      Top             =   630
      Width           =   1590
      Begin VB.OptionButton OpcProm 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Promediales"
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
         TabIndex        =   11
         Top             =   525
         Width           =   1380
      End
      Begin VB.OptionButton OpcNotas 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Por Curso"
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
         Top             =   210
         Value           =   -1  'True
         Width           =   1170
      End
   End
   Begin MSAdodcLib.Adodc AdoAux 
      Height          =   330
      Left            =   525
      Top             =   3990
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
   Begin MSAdodcLib.Adodc AdoPromedio1 
      Height          =   330
      Left            =   525
      Top             =   4305
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
      Caption         =   "Promedio1"
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
   Begin MSAdodcLib.Adodc AdoPromedio2 
      Height          =   330
      Left            =   525
      Top             =   4620
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
      Caption         =   "Promedio2"
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
   Begin MSAdodcLib.Adodc AdoResumen 
      Height          =   330
      Left            =   525
      Top             =   4935
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
      Caption         =   "Resumen"
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
      Left            =   9765
      TabIndex        =   12
      ToolTipText     =   "Formato de Fecha: DD/MM/AA"
      Top             =   1575
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
   Begin MSAdodcLib.Adodc AdoMaterias 
      Height          =   330
      Left            =   525
      Top             =   5250
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
      Caption         =   "Materias"
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
   Begin MSAdodcLib.Adodc AdoNotas 
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
      Caption         =   "Notas"
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
   Begin MSAdodcLib.Adodc AdoResumenNotas 
      Height          =   330
      Left            =   525
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
      Caption         =   "ResumenNotas"
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
   Begin MSAdodcLib.Adodc AdoMatriculas 
      Height          =   330
      Left            =   525
      Top             =   6195
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
      Caption         =   "Matriculas"
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
   Begin MSAdodcLib.Adodc AdoConducta 
      Height          =   330
      Left            =   525
      Top             =   6510
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
      Caption         =   "Conducta"
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
   Begin MSDataListLib.DataCombo DCMaterias 
      Bindings        =   "FNotaQui.frx":0375
      DataSource      =   "AdoMaterias"
      Height          =   315
      Left            =   2730
      TabIndex        =   8
      Top             =   2940
      Width           =   6945
      _ExtentX        =   12250
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
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
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " SELECCIONE LA MATERIA:"
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
      Top             =   2940
      Width           =   2640
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " SELECCIONE LA MATERIA:"
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
      Width           =   4005
   End
   Begin ComctlLib.ImageList ImgLstMenu 
      Left            =   11445
      Top             =   735
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
            Picture         =   "FNotaQui.frx":038F
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FNotaQui.frx":06A9
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FNotaQui.frx":09C3
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FNotaQui.frx":0CDD
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FNotaQui.frx":0FF7
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FNotaQui.frx":1339
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FNotaQui.frx":1653
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FNotaQui.frx":196D
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FNotaQui.frx":1C87
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FNotaQui.frx":1FA1
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FNotaQui.frx":227F
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FNotaQui.frx":23F5
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FNotaQui.frx":270F
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FNotaQui.frx":2A29
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FNotaQui.frx":2D43
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FNotaQui.frx":305D
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FNotaQui.frx":3377
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FNotaQui.frx":3691
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Rango Nomina Oficial"
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
      TabIndex        =   3
      Top             =   735
      Width           =   2115
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
      Left            =   12915
      TabIndex        =   13
      Top             =   1575
      Width           =   1590
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   13860
      Top             =   735
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   1
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FNotaQui.frx":39AB
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FNotasQuimestre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Column, Row As Integer
Dim Index1, Index2, Index3, Index4 As Integer

Dim SumaHoriz(50) As Single
Dim SumaHorizT(50) As Single
Dim SumaHorizPQ(50) As Single
Dim SumaHorizSQ(50) As Single
Dim SumaVerti(50) As Single
Dim VectNota(50) As Currency
Dim VectNotaP(50) As Currency
Dim VectMate(50) As String
Dim VectCodMat(50) As String
Dim VectCamposNotas() As String

Dim ContHoriz As Integer
Dim ContVert As Integer

Dim CadenaParcial As String
Dim Dec_Campos As String
Dim SQLPromQ As String
Dim AltoMaximo As Single
Dim AnchoMaximo As Single
Dim OpcionNotas As Byte

Public Function Print_Nota(Nota_No As Byte) As Boolean
Dim FPrint_Nota(23) As Boolean
    For I = 1 To 22
        FPrint_Nota(I) = False
    Next I
    If FormatoLibreta = "BIMESTRES" Then
       If OpcPeriodo("PQBim1", LstPeriodos) Then
          FPrint_Nota(1) = True
       End If
       If OpcPeriodo("PQBim2", LstPeriodos) Then
          For I = 1 To 4
              FPrint_Nota(I) = True
          Next I
       End If
       If OpcPeriodo("SQBim1", LstPeriodos) Then
          For I = 1 To 5
              FPrint_Nota(I) = True
          Next I
       End If
       If OpcPeriodo("SQBim1", LstPeriodos) Then
          For I = 1 To 8
              FPrint_Nota(I) = True
          Next I
          FPrint_Nota(10) = True
       End If
       If OpcPeriodo("PF", LstPeriodos) Then
          For I = 1 To 10
              FPrint_Nota(I) = True
          Next I
       End If
    Else
       If OpcPeriodo("PQBim1", LstPeriodos) Then
          FPrint_Nota(1) = True
       End If
       If OpcPeriodo("PQBim2", LstPeriodos) Then
          For I = 1 To 6
              FPrint_Nota(I) = True
          Next I
       End If
       If OpcPeriodo("SQBim1", LstPeriodos) Then
          For I = 1 To 7
              FPrint_Nota(I) = True
          Next I
       End If
       If OpcPeriodo("SQBim2", LstPeriodos) Then
          For I = 1 To 12
              FPrint_Nota(I) = True
          Next I
          FPrint_Nota(10) = True
       End If
       If OpcPeriodo("PF", LstPeriodos) Then
          For I = 1 To 15
              FPrint_Nota(I) = True
          Next I
       End If
    End If
    Print_Nota = FPrint_Nota(Nota_No)
End Function

Public Sub Imprimir_Pagina()
Dim AnchoPict As Single
Dim AltoPict As Single
Dim NombFilePict As String
On Error GoTo Errorhandler
Mensajes = "Seguro de Imprimir en:" & vbCrLf & Printer.DeviceName & "?"
Titulo = "IMPRESION"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
   RatonReloj
   InicioX = 0.01: InicioY = 0.01
   Escala_Centimetro Orientacion_Pagina, TipoTimes, 8
   Pagina = 1
  'vbTwips
   If Printer.ScaleMode = vbCentimeters Then
      Printer.ScaleWidth = Me.ScaleX(Printer.ScaleWidth, vbCentimeters, vbPixels)
      Printer.ScaleHeight = Me.ScaleX(Printer.ScaleHeight, vbCentimeters, vbPixels)
      Printer.ScaleMode = vbPixels
   End If
   Printer.PaintPicture PictCalif.Image, 0, 0, Printer.ScaleWidth, Printer.ScaleHeight
   MensajeEncabData = ""
   Printer.EndDoc
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

Private Sub Command1_Click()
   RatonNormal
   Unload Me
End Sub

Private Sub DGResumenNotas_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  If KeyCode = vbKeyF1 Then
     DGResumenNotas.Visible = False
     GenerarDataTexto FNotasQuimestre, AdoResumenNotas
     DGResumenNotas.Visible = True
  End If
End Sub

Private Sub DLCurso_DblClick()
  SiguienteControl
End Sub

Private Sub DLCurso_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DLCurso_LostFocus()
    Codigo = SinEspaciosIzq(DLCurso)
    CodigoA = ""
    CodigoB = ""
    CodigoL = ""
    Cadena = Leer_Datos_del_Curso(Codigo)
    CuentaBanco = Cadena
    If Len(Dato_Curso.Titulo) > 1 Then CodigoA = Dato_Curso.Titulo
    If Len(Dato_Curso.Tipo_Titulo) > 1 Then CodigoB = Dato_Curso.Tipo_Titulo
    If Len(Dato_Curso.Especialidad) > 1 Then CodigoL = Dato_Curso.Especialidad
    Listar_Alumnos_Notas_Del Codigo
    If AdoMaterias.Recordset.RecordCount > 0 Then
       DCMaterias.Visible = True
       DCMaterias.SetFocus
    Else
       DCMaterias.Visible = False
       MBFecha.SetFocus
    End If
End Sub

Private Sub LstPeriodos_DblClick()
  SiguienteControl
End Sub

Private Sub LstPeriodos_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub LstPeriodos_LostFocus()
Dim CantCampos As Byte
     SQLTAI = Ninguno
     SQLAIC = Ninguno
     SQLAGC = Ninguno
     SQLL = Ninguno
     SQLExaP = Ninguno
     SQLNotas = Ninguno
     SQLBim1 = Ninguno
     SQLBim2 = Ninguno
     SQLBim3 = Ninguno
     SQLExamen = Ninguno
     SQLPromQ = Ninguno
     CantCampos = 5
     CadenaParcial = LstPeriodos.Text
     If OpcPeriodo("PF", LstPeriodos) Then Cadena = "Periodo Final de Quimestres"
     If OpcPeriodo("PQBim1", LstPeriodos) Then
        SQLTAI = "PQTAI1"
        SQLAIC = "PQAIC1"
        SQLAGC = "PQAGC1"
        SQLL = "PQL1"
        SQLExaP = "PQExaP1"
        SQLNotas = "PQBim1"
        SQLBim1 = "PQBim1"
        SQLBim2 = "PQBim2"
        SQLBim3 = "PQBim3"
        SQLProm = "PQBim1"
        SQLExamen = "ExamenPQ"
        SQLPromQ = "PromPQ"
        OpcionNotas = 1
     End If
     If OpcPeriodo("PQBim2", LstPeriodos) Then
        SQLTAI = "PQTAI2"
        SQLAIC = "PQAIC2"
        SQLAGC = "PQAGC2"
        SQLL = "PQL2"
        SQLExaP = "PQExaP2"
        SQLNotas = "PQBim2"
        SQLBim1 = "PQBim1"
        SQLBim2 = "PQBim2"
        SQLBim3 = "PQBim3"
        SQLProm = "PQBim2"
        SQLExamen = "ExamenPQ"
        SQLPromQ = "PromPQ"
        OpcionNotas = 2
     End If
     If OpcPeriodo("PQBim3", LstPeriodos) Then
        SQLTAI = "PQTAI3"
        SQLAIC = "PQAIC3"
        SQLAGC = "PQAGC3"
        SQLL = "PQL3"
        SQLExaP = "PQExaP3"
        SQLNotas = "PQBim3"
        SQLBim1 = "PQBim1"
        SQLBim2 = "PQBim2"
        SQLBim3 = "PQBim3"
        SQLProm = "PQBim3"
        SQLExamen = "ExamenPQ"
        SQLPromQ = "PromPQ"
        OpcionNotas = 3
     End If
     If OpcPeriodo("PQ", LstPeriodos) Then
        SQLTAI = "PQTAI3"
        SQLAIC = "PQAIC3"
        SQLAGC = "PQAGC3"
        SQLL = "PQL3"
        SQLExaP = "PQExaP3"
        SQLNotas = "PQBim3"
        SQLBim1 = "PQBim1"
        SQLBim2 = "PQBim2"
        SQLBim3 = "PQBim3"
        SQLProm = "PQBim3"
        SQLExamen = "ExamenPQ"
        SQLPromQ = "PromPQ"
        CantCampos = 6
        OpcionNotas = 4
     End If
     If OpcPeriodo("SQBim1", LstPeriodos) Then
        SQLTAI = "SQTAI1"
        SQLAIC = "SQAIC1"
        SQLAGC = "SQAGC1"
        SQLL = "SQL1"
        SQLExaP = "SQExaP1"
        SQLNotas = "SQBim1"
        SQLBim1 = "SQBim1"
        SQLBim2 = "SQBim2"
        SQLBim3 = "SQBim3"
        SQLProm = "SQBim1"
        SQLExamen = "ExamenSQ"
        SQLPromQ = "PromSQ"
        OpcionNotas = 1
     End If
     If OpcPeriodo("SQBim2", LstPeriodos) Then
        SQLTAI = "SQTAI2"
        SQLAIC = "SQAIC2"
        SQLAGC = "SQAGC2"
        SQLL = "SQL2"
        SQLExaP = "SQExaP2"
        SQLNotas = "SQBim2"
        SQLBim1 = "SQBim1"
        SQLBim2 = "SQBim2"
        SQLBim3 = "SQBim3"
        SQLProm = "SQBim2"
        SQLExamen = "ExamenSQ"
        SQLPromQ = "PromSQ"
        OpcionNotas = 2
     End If
     If OpcPeriodo("SQBim3", LstPeriodos) Then
        SQLTAI = "SQTAI3"
        SQLAIC = "SQAIC3"
        SQLAGC = "SQAGC3"
        SQLL = "SQL3"
        SQLExaP = "SQExaP3"
        SQLNotas = "SQBim3"
        SQLBim1 = "SQBim1"
        SQLBim2 = "SQBim2"
        SQLBim3 = "SQBim3"
        SQLProm = "SQBim3"
        SQLExamen = "ExamenSQ"
        SQLPromQ = "PromSQ"
        OpcionNotas = 3
     End If
     If OpcPeriodo("SQ", LstPeriodos) Then
        SQLTAI = "SQTAI3"
        SQLAIC = "SQAIC3"
        SQLAGC = "SQAGC3"
        SQLL = "SQL3"
        SQLExaP = "SQExaP3"
        SQLNotas = "SQBim3"
        SQLBim1 = "SQBim1"
        SQLBim2 = "SQBim2"
        SQLBim3 = "SQBim3"
        SQLProm = "SQBim3"
        SQLExamen = "ExamenSQ"
        SQLPromQ = "PromSQ"
        CantCampos = 6
        OpcionNotas = 4
     End If
     If OpcPeriodo("TQBim1", LstPeriodos) Then
        SQLBim1 = "TQBim1"
        SQLBim2 = "TQBim2"
        SQLBim3 = "TQBim3"
        SQLExamen = "ExamenTQ"
        SQLPromQ = "PromTQ"
        OpcionNotas = 1
     End If
     If OpcPeriodo("TQBim2", LstPeriodos) Then
        SQLBim1 = "TQBim1"
        SQLBim2 = "TQBim2"
        SQLBim3 = "TQBim3"
        SQLExamen = "ExamenTQ"
        SQLPromQ = "PromTQ"
        OpcionNotas = 2
     End If
     If OpcPeriodo("TQBim3", LstPeriodos) Then
        SQLBim1 = "TQBim1"
        SQLBim2 = "TQBim2"
        SQLBim3 = "TQBim3"
        SQLExamen = "ExamenTQ"
        SQLPromQ = "PromTQ"
        OpcionNotas = 3
     End If
     If OpcPeriodo("TQ", LstPeriodos) Then
        SQLBim1 = "TQBim1"
        SQLBim2 = "TQBim2"
        SQLBim3 = "TQBim3"
        SQLExamen = "ExamenTQ"
        SQLPromQ = "PromTQ"
        OpcionNotas = 4
     End If
     
     ReDim VectCamposNotas(1 To CantCampos) As String
     VectCamposNotas(1) = SQLTAI
     VectCamposNotas(2) = SQLAIC
     VectCamposNotas(3) = SQLAGC
     VectCamposNotas(4) = SQLL
     VectCamposNotas(5) = SQLExaP
     If CantCampos = 6 Then VectCamposNotas(6) = SQLExamen
    'MsgBox CadenaParcial
End Sub

Private Sub MBFecha_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBFecha_LostFocus()
  FechaValida MBFecha
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
   DGResumenNotas.Visible = False
   FrmPictCalif.Visible = True
   DCMaterias.Visible = True
  
  'MsgBox "Opcion del Menu: " & vbCrLf & Button.key & vbCrLf & "Opcion de Impresion: " & Opciones
   Dec_Campos = "Tot_" & SQLPromQ & " 3|"
   Select Case Button.key
     Case "Salir"
          Unload FNotasQuimestre
          Opciones = 0
     Case "Imprimir"
          'MsgBox Opciones
          Select Case Opciones
            Case 0
                 Imprimir_Pagina
                 FNotasQuimestre.Caption = "RESUMEN NOTAS QUIMESTRALES"
            Case 1
                 Cuadricula = False
                'SQLMsg1 = ""
                 SQLMsg2 = ""
                 SQLMsg3 = ""
                 If OpcPeriodo("PQBim2", LstPeriodos) Then
                    Imprimir_Mejor_Promedio AdoResumenNotas, True, 1, 9, Opcion, 1, Dec_Campos
                 ElseIf OpcPeriodo("SQBim2", LstPeriodos) Then
                    Imprimir_Mejor_Promedio AdoResumenNotas, True, 1, 9, Opcion, 2, Dec_Campos
                 Else
                    Imprimir_Mejor_Promedio AdoResumenNotas, True, 1, 9, Opcion, 3, Dec_Campos
                 End If
            Case 2
                 Imprimr_Nomina_Oficial
                 Mensajes = "Generar Archivo para el Ministerio de Educación"
                 Titulo = "GENERACION ARCHIVO EXCEL"
                 If BoxMensaje = vbYes Then Imprimr_Nomina_Oficial_Excel
          End Select
     Case "Notas"
          Listar_Notas_Del_Curso
          Opciones = 0
     Case "MNotaQuimestre"
          Listar_Calificacion_Del_Curso PictCalif
          Opciones = 0
     Case "CuadroSubNotas"
          Listar_Calificacion_Del_Curso PictCalif, , True
          Opciones = 0
     Case "Disciplina"
          Listar_Disciplina_Del_Curso
          Opciones = 0
     Case "Listas"
          Lista_Alumnos_por_Curso False
          Opciones = 0
     Case "ListasMes"
          Lista_Alumnos_por_Curso True
          Opciones = 0
     Case "ListaAlumnos"
          Lista_Alumnos_por_Notas True
          Opciones = 0
     Case "NotasProfesor"
          Listar_Materias_x_Profesor
          Opciones = 0
     Case "MejorPromedio"
          Listar_Mejor_Promedio
          DGResumenNotas.Visible = True
          FrmPictCalif.Visible = False
          Opciones = 1
     Case "MMejorEgresado"
          Listar_Mejor_Egresado
          Opciones = 0
     Case "ResumenExamenGrado"
          Mensajes = "Listar los Supletorios"
          Titulo = "LISTA DE NOTAS DE EXAMENES DE GRADO"
          If BoxMensaje = vbYes Then
             Listar_Examenes_Grado True
          Else
             Listar_Examenes_Grado False
          End If
          Opciones = 0
     Case "MNominaOficial"
          Listar_Nomina_Oficial
          Opciones = 2
     Case "BPromobidos"
          Listar_Calificacion_Del_Curso PictCalif, True
          Opciones = 0
     Case "NotasBlanco"
          Codigo = SinEspaciosIzq(DLCurso)
          Listar_Notas_Blanco Codigo
   End Select
End Sub

Private Sub TxtDesde_GotFocus()
  MarcarTexto TxtDesde
End Sub

Private Sub TxtDesde_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtHasta_GotFocus()
  MarcarTexto TxtHasta
End Sub

Private Sub TxtHasta_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub VScroll1_Change()
  VScroll1_Scroll
End Sub

Private Sub VScroll1_Scroll()
  PictCalif.Top = -VScroll1.value
  LblA4.Caption = " " & Format(VScroll1.value, "00.00") & " - " & Format(HScroll1.value, "00.00")
End Sub

Private Sub HScroll1_Change()
  HScroll1_Scroll
End Sub

Private Sub HScroll1_Scroll()
  PictCalif.Left = -HScroll1.value
  LblA4.Caption = " " & Format(VScroll1.value, "00.00") & " - " & Format(HScroll1.value, "00.00")
End Sub

'''Private Sub Command2_Click()
'''  RatonReloj
'''  'RutaOrigen = "D:\Mis Archivos Walter\Fotos\BMP\BEBES.BMP"
'''  FVerGrafico.Show
'''End Sub

Private Sub Form_Activate()
  Label3.Caption = ""
  MBFecha.Text = FechaSistema
  Leer_Periodo_Lectivo
  
  FrmPictCalif.width = MDI_X_Max - 100
  FrmPictCalif.Height = MDI_Y_Max - 3400
  FrmPictCalif.Refresh
  Picture1.width = FrmPictCalif.width - 500
  Picture1.Height = FrmPictCalif.Height - 500
  DGResumenNotas.width = MDI_X_Max - 100
  DGResumenNotas.Height = MDI_Y_Max - 3400
  
  VScroll1.Height = Picture1.Height - 200
  HScroll1.width = Picture1.width - HScroll1.Left
  HScroll1.Top = FNotasQuimestre.Height - 1800
  LblA4.Top = Picture1.Height + Picture1.Top
  HScroll1.Top = Picture1.Height + Picture1.Top
  'LstVAlumnos.width = FLibretas.width - LstVAlumnos.Left - 300
  'MsgBox "..."
  Command1.Top = PictCalif.Height - PictCalif.Top
  VScroll1.value = 0
  HScroll1.value = 0
  HScroll1.Min = 0
  VScroll1.Min = 0
  VScroll1.Max = PictCalif.Height + 20 '- Picture1.Height
  HScroll1.Max = PictCalif.width  '- Picture1.Width
  VScroll1_Scroll
  HScroll1_Scroll
  LstPeriodos.Clear
    If Mid$(FormatoLibreta, 1, 9) = "TRIMESTRE" Then
       LstPeriodos.AddItem "Primer Trimestre Primer Periodo"
       LstPeriodos.AddItem "Primer Trimestre Segundo Periodo"
       LstPeriodos.AddItem "Promedio Primer Trimestre"
       
       LstPeriodos.AddItem "Segundo Trimestre Primer Periodo"
       LstPeriodos.AddItem "Segundo Trimestre Segundo Periodo"
       LstPeriodos.AddItem "Promedio Segundo Trimestre"
       
       LstPeriodos.AddItem "Tercer Trimestre Primer Periodo"
       LstPeriodos.AddItem "Tercer Trimestre Segundo Periodo"
       LstPeriodos.AddItem "Promedio Tercer Trimestre"
       
    ElseIf Mid$(FormatoLibreta, 1, 9) = "QUIMESTRE" Then
       LstPeriodos.AddItem "Primer Quimestre Primer Parcial"
       LstPeriodos.AddItem "Primer Quimestre Segundo Parcial"
       LstPeriodos.AddItem "Primer Quimestre Tercer Parcial"
       LstPeriodos.AddItem "Primer Quimestre Examen"
       LstPeriodos.AddItem "Promedio Primer Quimestre"
       
       LstPeriodos.AddItem "Segundo Quimestre Primer Parcial"
       LstPeriodos.AddItem "Segundo Quimestre Segundo Parcial"
       LstPeriodos.AddItem "Segundo Quimestre Tercer Parcial"
       LstPeriodos.AddItem "Segundo Quimestre Examen"
       LstPeriodos.AddItem "Promedio Segundo Quimestre"
       
       LstPeriodos.AddItem "Todos los Quimestres"
    Else
       LstPeriodos.AddItem "Primer Quimestre Primer Periodo"
       LstPeriodos.AddItem "Primer Quimestre Segundo Periodo"
       LstPeriodos.AddItem "Promedio Primer Quimestre"
       
       LstPeriodos.AddItem "Segundo Quimestre Primer Periodo"
       LstPeriodos.AddItem "Segundo Quimestre Segundo Periodo"
       LstPeriodos.AddItem "Promedio Segundo Quimestre"
    End If
  LstPeriodos.Text = LstPeriodos.List(0)
'''  Pos_Pict_X = PictCalif.Left
'''  Pos_Pict_Y = PictCalif.Top
  Cadena1 = Tipo_Acceso_Educativo("", "CodigoE")
  sSQL = "SELECT (CodigoE & ' - ' & Detalle) As Cursos,Seccion,Detalle As Paralelo,CodigoE " _
       & "FROM Catalogo_Estudiantil " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND TC = 'P' " _
       & Cadena1 _
       & "ORDER BY CodigoE "
       
  sSQL = "SELECT (Curso & ' - ' & Descripcion) As Cursos,Seccion,Paralelo,Curso " _
       & "FROM Catalogo_Cursos " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND LEN(Curso) > 4 " _
       & "ORDER BY Curso "
  SelectDBList DLCurso, AdoAux, sSQL, "Cursos"
  FNotasQuimestre.Caption = "RESUMEN NOTAS QUIMESTRALES"
  TxtDesde = "0"
  TxtHasta = "0"
  sSQL = "SELECT C.Cliente,CM.CI,TN.* " _
       & "FROM Trans_Actas As TN,Clientes As C,Clientes_Matriculas As CM " _
       & "WHERE TN.Item = '" & NumEmpresa & "' " _
       & "AND TN.Periodo = '" & Periodo_Contable & "' " _
       & "AND TN.CodE > '3.03' " _
       & "AND TN.Id_No <> 0 " _
       & "AND TN.Codigo = C.Codigo " _
       & "AND CM.Codigo = C.Codigo " _
       & "AND TN.Item = CM.Item " _
       & "AND TN.Periodo = CM.Periodo " _
       & "ORDER BY TN.Id_No "
  SelectAdodc AdoNotas, sSQL
  If AdoNotas.Recordset.RecordCount > 0 Then
     AdoNotas.Recordset.MoveFirst
     TxtDesde = Format(AdoNotas.Recordset.Fields("Id_No"), "000")
     AdoNotas.Recordset.MoveLast
     TxtHasta = Format(AdoNotas.Recordset.Fields("Id_No"), "000")
  End If
  sSQL = "SELECT CM.*,TA.ConductaPQ1,TA.ConductaPQ2,TA.ConductaSQ1,TA.ConductaSQ2 " _
       & "FROM Clientes_Matriculas As CM,Trans_Asistencia As TA " _
       & "WHERE CM.Item = '" & NumEmpresa & "' " _
       & "AND CM.Periodo = '" & Periodo_Contable & "' " _
       & "AND CM.Item = TA.Item " _
       & "AND CM.Periodo = TA.Periodo " _
       & "AND CM.Codigo = TA.Codigo " _
       & "ORDER BY CM.Codigo "
  SelectAdodc AdoMatriculas, sSQL
  
  Codigo = SinEspaciosIzq(DLCurso)
  Listar_Materias_Curso Codigo
  RatonNormal
  Listar_Alumnos_Notas_Del SinEspaciosIzq(DLCurso)
  LstPeriodos.SetFocus
End Sub

Private Sub Form_Load()
  ConectarAdodc AdoAux
  ConectarAdodc AdoNotas
  ConectarAdodc AdoResumen
  ConectarAdodc AdoConducta
  ConectarAdodc AdoMaterias
  ConectarAdodc AdoPromedio1
  ConectarAdodc AdoPromedio2
  ConectarAdodc AdoMatriculas
  ConectarAdodc AdoResumenNotas
End Sub

'''Public Sub Listar_Quimestres()
''''Procesamos las Notas de Curso
'''  PictCalif.Cls
'''  PictCalif.FontName = TipoArialNarrow
'''  PictCalif.FontSize = 10
'''  For I = 1 To 50
'''      SumaHoriz(I) = 0
'''      VectNota(I) = 0
'''      VectMate(I) = "."
'''      VectCodMat(I) = "."
'''  Next I
'''  RatonReloj
'''  Contador = 0
'''  ContNotas = 0
'''  TipoDoc = Ninguno
'''  Codigo = SinEspaciosIzq(DLCurso.Text)
'''  With AdoAux.Recordset
'''   If .RecordCount > 0 Then
'''       If Codigo = "" Then Codigo = "."
'''      .MoveFirst
'''      .Find ("CodigoE = '" & Codigo & "' ")
'''       If Not .EOF Then
'''          TipoDoc = .Fields("Seccion")
'''          Codigo1 = .Fields("Paralelo")
'''       End If
'''   End If
'''  End With
'''  sSQL = "SELECT C.Cliente,CE.Detalle As Materia,C.Grupo,CE.C,CE.P,TN.* " _
'''       & "FROM Clientes As C,Catalogo_Estudiantil As CE,Trans_Notas As TN " _
'''       & "WHERE C.Grupo = '" & Codigo & "' " _
'''       & "AND CE.Item = '" & NumEmpresa & "' " _
'''       & "AND CE.Periodo = '" & Periodo_Contable & "' " _
'''       & "AND CE.CodMat = TN.CodMat " _
'''       & "AND TN.Codigo = C.Codigo " _
'''       & "AND C.Grupo = Mid$(CE.CodigoE,1,7) " _
'''       & "AND CE.Item = TN.Item " _
'''       & "AND CE.Periodo = TN.Periodo " _
'''       & "ORDER BY C.Cliente,CE.CodigoE "
'''  'MsgBox sSQL
'''  SelectAdodc AdoResumen, sSQL
'''  PictCalif.FontSize = 8
'''  PictCalif.FontBold = False
'''  With AdoResumen.Recordset
'''   If .RecordCount > 0 Then
'''       CodigoCliente = .Fields("Codigo")
'''       NombreCliente = .Fields("Cliente")
'''       Contador = 0
'''       ContHoriz = 0
'''       ContVert = 0
'''       PFil = 4
'''      'Contamos las Notas y los Alumnos
'''       Do While Not .EOF
'''          If CodigoCliente <> .Fields("Codigo") Then
'''             PictPrint_Texto PictCalif, 1, PFil, Format(ContVert + 1, "00") & ".-"
'''             PictPrint_Texto PictCalif, 1.5, PFil, NombreCliente
'''             PFil = PFil + 0.45
'''             CodigoCliente = .Fields("Codigo")
'''             NombreCliente = .Fields("Cliente")
'''             If ContNotas < ContHoriz Then ContNotas = ContHoriz
'''             ContHoriz = 0
'''             ContVert = ContVert + 1
'''          End If
'''          ContHoriz = ContHoriz + 1
'''          VectMate(ContHoriz) = .Fields("Materia")
'''         .MoveNext
'''       Loop
'''       PictPrint_Texto PictCalif, 1, PFil, Format(ContVert + 1, "00") & ".-"
'''       PictPrint_Texto PictCalif, 1.5, PFil, NombreCliente
'''       If ContNotas < ContHoriz Then ContNotas = ContHoriz
'''       ContVert = ContVert + 1
'''       PictCalif.Height = (ContVert * 0.5) + 10
'''       PictCalif.Width = ContNotas * 3
'''       VScroll1.Min = 0
'''       HScroll1.Min = 0
'''       VScroll1.Max = PictCalif.Height
'''       HScroll1.Max = PictCalif.Width
'''       VScroll1.Value = 0
'''       HScroll1.Value = 0
'''       PCol = 10
'''       PFil = 3.5
'''       PictCalif.FontBold = True
'''       For I = 1 To ContNotas
'''           Cadena = SinEspaciosIzq(VectMate(I))
'''           Cadena1 = Trim(Mid$(VectMate(I), Len(Cadena) + 1, Len(VectMate(I))))
'''           If Cadena = "" Then Cadena = " "
'''           If Cadena1 = "" Then Cadena1 = " "
'''           PictPrint_Texto PictCalif, PCol, 0.3, Format(I, "00")
'''           cPrint.printTextoAngulo PictCalif, PCol, PFil, 90, 0.5, 8, Cadena
'''           PCol = PCol + 0.45
'''           cPrint.printTextoAngulo PictCalif, PCol, PFil, 90, 0.5, 8, Cadena1
'''           PCol = PCol + 1.5
'''       Next I
'''      'Generamos las Notas del Curso
'''      .MoveFirst
'''       Do While Not .EOF
'''          If CodigoCliente <> .Fields("Codigo") Then
'''             ContVert = ContVert + 1
'''             SetAdoAddNew "Balances_Mes"
'''             SetAdoFields "CodigoC", CodigoCliente
'''             SetAdoFields "Codigo", Codigo
'''             SetAdoFields "TC", "NO"
'''             Sumatoria = 0
'''             ContHoriz = 0
'''             Insertar_Notas_Curso Contador
'''             'SumaHoriz(ContVert) = Sumatoria
'''             If ContHoriz = 0 Then ContHoriz = 1
'''             Entrada = Round(Sumatoria / ContHoriz, 2)
'''             SetAdoFields "TOTAL", Entrada
'''             SetAdoUpdate
'''             'MsgBox CodigoCliente
'''             For I = 1 To Contador
'''                 VectNota(I) = 0
'''             Next I
'''             CodigoCliente = .Fields("Codigo")
'''             If ContNotas < Contador Then ContNotas = Contador
'''             Contador = 0
'''          End If
'''          Contador = Contador + 1
'''          If OpcPQBim1.Value Then VectNota(Contador) = .Fields("PQBim1")
'''          If OpcPQBim2.Value Then VectNota(Contador) = .Fields("PQBim2")
'''          If OpcPQ.Value Then VectNota(Contador) = .Fields("PromPQ")
'''
'''          If OpcSQBim1.Value Then VectNota(Contador) = .Fields("SQBim1")
'''          If OpcSQBim2.Value Then VectNota(Contador) = .Fields("SQBim2")
'''          If OpcSQ.Value Then VectNota(Contador) = .Fields("PromSQ")
'''
'''          VectMate(Contador) = .Fields("Materia")
'''          VectCodMat(Contador) = .Fields("CodMat")
'''          If .Fields("C") Then VectMate(Contador) = VectMate(Contador) & "_"
'''          If .Fields("P") Then VectMate(Contador) = VectMate(Contador) & "|"
'''         .MoveNext
'''       Loop
'''       ContVert = ContVert + 1
'''       'SumaHoriz(ContVert) = Sumatoria
'''       If ContHoriz = 0 Then ContHoriz = 1
'''       Entrada = Round(Sumatoria / ContHoriz, 2)
'''      .MoveFirst
'''   End If
'''  End With
'''  Total = 0
'''  For I = 1 To ContVert
'''    'Total = Total + SumaHoriz(I)
'''  Next I
'''  If ContNotas = 0 Then ContNotas = 1
'''  Total = Round(Total / ContNotas)
'''  RatonNormal
'''End Sub

Public Sub Listar_Alumnos_Notas_Del(CodCurso As String, Optional Tipo_PF As Boolean, Optional Lst_Supletorio As Boolean, Optional SubNotas As Boolean)
 'Notas Quimestrales
  CantAlumn = Dato_Curso.ContAlumnos
 
  sSQL = "SELECT CC.Curso,CC.Descripcion,C.Cliente,C.Sexo,TM.Materia,TM.C,TM.C2,TM.P,TM.I,TM.SDiv," _
       & "CM.Matricula_No,CM.Folio_No,CC.Especialidad,CC.Seccion,CM.Domicilio," _
       & "CM.Telefono_D,CM.Telefono_R,CM.Representante,CM.Aprobado,TN.* "
  If SubNotas Then
     sSQL = sSQL & "FROM Trans_Notas_Auxiliares As TN,"
  Else
     sSQL = sSQL & "FROM Trans_Notas As TN,"
  End If
  sSQL = sSQL & "Catalogo_Materias As TM," _
       & "Catalogo_Cursos As CC," _
       & "Clientes As C," _
       & "Clientes_Matriculas As CM " _
       & "WHERE TN.Item = '" & NumEmpresa & "' " _
       & "AND TN.Periodo = '" & Periodo_Contable & "' " _
       & "AND TN.CodE = '" & CodCurso & "' "
  If Tipo_PF Then sSQL = sSQL & "AND TM.I <> " & Val(adFalse) & " "
  If Lst_Supletorio Then sSQL = sSQL & "AND CM.Aprobado = " & Val(adFalse) & " "
  sSQL = sSQL _
       & "AND TN.Codigo = C.Codigo " _
       & "AND TN.Codigo = CM.Codigo " _
       & "AND TN.CodE = CC.Curso " _
       & "AND TN.CodMat = TM.CodMat " _
       & "AND TN.Item = CC.Item " _
       & "AND TN.Item = TM.Item " _
       & "AND TN.Item = CM.Item " _
       & "AND TN.Periodo = CC.Periodo " _
       & "AND TN.Periodo = TM.Periodo " _
       & "AND TN.Periodo = CM.Periodo " _
       & "ORDER BY CC.Curso,C.Cliente,TN.Id_No,TN.Orden "
  SelectAdodc AdoResumen, sSQL
 'Conducta Quimestrales
  sSQL = "SELECT CC.Curso,C.Codigo,TM.Materia,CM.Telefono_D,TN.* " _
       & "FROM Trans_Notas As TN," _
       & "Catalogo_Materias As TM," _
       & "Catalogo_Cursos As CC," _
       & "Clientes As C," _
       & "Clientes_Matriculas As CM " _
       & "WHERE TN.Item = '" & NumEmpresa & "' " _
       & "AND TN.Periodo = '" & Periodo_Contable & "' " _
       & "AND TN.CodE = '" & CodCurso & "' " _
       & "AND TN.CodMat IN ('997','999','998') " _
       & "AND TN.Codigo = C.Codigo " _
       & "AND TN.Codigo = CM.Codigo " _
       & "AND TN.CodE = CC.Curso " _
       & "AND TN.CodMat = TM.CodMat " _
       & "AND TN.Item = CC.Item " _
       & "AND TN.Item = TM.Item " _
       & "AND TN.Item = CM.Item " _
       & "AND TN.Periodo = CC.Periodo " _
       & "AND TN.Periodo = TM.Periodo " _
       & "AND TN.Periodo = CM.Periodo " _
       & "ORDER BY CC.Curso,C.Codigo "
  SelectAdodc AdoConducta, sSQL
 ' MsgBox AdoConducta.Recordset.RecordCount & vbCrLf & vbCrLf & sSQL
 'Asistencia
  sSQL = "SELECT CC.Curso,CC.Descripcion,C.Cliente,TM.Materia,TM.C,TM.P,TM.I,TN.* " _
       & "FROM Trans_Asistencia As TN,Catalogo_Materias As TM,Catalogo_Cursos As CC," _
       & "Clientes As C " _
       & "WHERE TN.Item = '" & NumEmpresa & "' " _
       & "AND TN.Periodo = '" & Periodo_Contable & "' " _
       & "AND TN.CodE = '" & CodCurso & "' " _
       & "AND TN.Codigo = C.Codigo " _
       & "AND TN.CodE = CC.Curso " _
       & "AND TN.CodMat = TM.CodMat " _
       & "AND TN.Item = CC.Item " _
       & "AND TN.Item = TM.Item " _
       & "AND TN.Periodo = CC.Periodo " _
       & "AND TN.Periodo = TM.Periodo " _
       & "ORDER BY CC.Curso,C.Cliente "
  SelectAdodc AdoPromedio1, sSQL
End Sub

Public Sub Listar_Disciplina_Del_Curso()
 NombreBanco = UCase(LstPeriodos.Text)
 PictCalif.Cls
 PictCalif.FontBold = True
 PCol = 10
 PictCalif.width = PCol + (ContNotas * 0.8) + 1.2
 PictCalif.Height = 12 + (AdoPromedio1.Recordset.RecordCount * 0.45)
 PictCalif.FontName = TipoTimes
 PosLinea = 1
 If LogoTipo <> "" Then PictCalif.PaintPicture LoadPicture(LogoTipo), 0.1, PosLinea, 5, 2.5
 PictCalif.FontSize = 20
 PictPrint_Texto PictCalif, 1, PosLinea, Institucion1, , PictCalif.width, True
 PosLinea = PosLinea + 0.7
 PictPrint_Texto PictCalif, 1, PosLinea, Institucion2, , PictCalif.width, True
 PosLinea = PosLinea + 0.7
 PictCalif.FontSize = 12
 PictCalif.FontBold = False
 PictPrint_Texto PictCalif, 1, PosLinea, Direccion & " Teléfono: " & Telefono1, , PictCalif.width, True
 PosLinea = PosLinea + 0.6
 PictCalif.FontSize = 16
 PictCalif.FontBold = True
 PictPrint_Texto PictCalif, 1, PosLinea, "C A L I F I C A C I O N E S    D E    D I S C I P L I N A", , PictCalif.width, True
 PictCalif.FontBold = False
 PosLinea = PosLinea + 0.7
 PictCalif.FontSize = 12
 PictPrint_Texto PictCalif, 1, PosLinea, "AÑO LECTIVO: " & Anio_Lectivo, , PictCalif.width, True
 PosLinea = PosLinea + 0.6
 PictPrint_Texto PictCalif, 0.8, PosLinea, NombreBanco
 PictPrint_Texto PictCalif, PictCalif.width - PictCalif.TextWidth(CuentaBanco) - 1.2, PosLinea, CuentaBanco
 PictCalif.FontSize = 28
 PictPrint_Texto PictCalif, 1.7, PosLinea + 1.5, "A L U M N O S"
 PictCalif.FontName = TipoArialNarrow
 PictCalif.FontSize = 9
 PCol = 10
 PosLinea = 7
 For I = 0 To ContNotas - 1
     Cadena2 = SinEspaciosDer(VectMate(I))
     J = Len(VectMate(I)) - Len(Cadena2)
     If J <= 0 Then
        Cadena1 = Cadena2
        Cadena2 = " "
     Else
        Cadena1 = Trim(Mid$(VectMate(I), 1, J))
     End If
     cPrint.printTextoAngulo PictCalif, PCol, PosLinea, 90, 5, PictCalif.FontSize, Cadena1
     PCol = PCol + 0.3
     cPrint.printTextoAngulo PictCalif, PCol, PosLinea, 90, 5, PictCalif.FontSize, Cadena2
     PCol = PCol + 0.5
 Next
 PictCalif.FontSize = 9
 PosLinea = PosLinea + 0.2
 PCol = 10.25
 Contador = 0
 For I = 0 To ContNotas - 1
     VectNota(I) = 0
     SumaHoriz(I) = 0
 Next I
 I = 0
 NumMeses = 0
 Cuota_No = 0
 PictCalif.FontBold = False
 With AdoResumen.Recordset
  If .RecordCount > 0 Then
     .MoveFirst
      Codigo = .Fields("Codigo")
      NombreCliente = .Fields("Cliente")
      Contador = Contador + 1
      PictPrint_Texto PictCalif, 0.8, PosLinea, Format(Contador, "00") & ".-"
      PictPrint_Texto PictCalif, 1.3, PosLinea, .Fields("Cliente")
      'PosLinea = PosLinea + 0.4
      Do While Not .EOF
         If Codigo <> .Fields("Codigo") Then
            Faltas_Just = 0
            Faltas_Injust = 0
            Atrasos = 0
            If AdoPromedio1.Recordset.RecordCount > 0 Then
               AdoPromedio1.Recordset.MoveFirst
               AdoPromedio1.Recordset.Find ("Codigo = '" & Codigo & "' ")
               If Not AdoPromedio1.Recordset.EOF Then
                  If OpcPeriodo("PQBim1", LstPeriodos) Then
                     NumFacturas = AdoPromedio1.Recordset.Fields("ConductaPQ1")
                     Faltas_Just = AdoPromedio1.Recordset.Fields("PQBFJ1")
                     Faltas_Injust = AdoPromedio1.Recordset.Fields("PQBFI1")
                     Atrasos = AdoPromedio1.Recordset.Fields("PQBA1")
                  End If
                  If OpcPeriodo("PQBim2", LstPeriodos) Then
                     NumFacturas = AdoPromedio1.Recordset.Fields("ConductaPQ2")
                     Faltas_Just = AdoPromedio1.Recordset.Fields("PQBFJ2")
                     Faltas_Injust = AdoPromedio1.Recordset.Fields("PQBFI2")
                     Atrasos = AdoPromedio1.Recordset.Fields("PQBA2")
                  End If
                  If OpcPeriodo("SQBim1", LstPeriodos) Then
                     NumFacturas = AdoPromedio1.Recordset.Fields("ConductaSQ1")
                     Faltas_Just = AdoPromedio1.Recordset.Fields("SQBFJ1")
                     Faltas_Injust = AdoPromedio1.Recordset.Fields("SQBFI1")
                     Atrasos = AdoPromedio1.Recordset.Fields("SQBA1")
                  End If
                  If OpcPeriodo("SQBim2", LstPeriodos) Then
                     NumFacturas = AdoPromedio1.Recordset.Fields("ConductaSQ2")
                     Faltas_Just = AdoPromedio1.Recordset.Fields("SQBFJ2")
                     Faltas_Injust = AdoPromedio1.Recordset.Fields("SQBFI2")
                     Atrasos = AdoPromedio1.Recordset.Fields("SQBA2")
                  End If
               End If
            End If
            Total = 0
            J = 0
            For I = 0 To ContNotas - 1
                If VectNota(I) > 0 Then
                   J = J + 1
                   Total = Total + VectNota(I)
                End If
            Next I
            If J = 0 Then
               Total = 0
            Else
               Total = Total / J
            End If
            Total = Round((Total + NumFacturas) / 2)
            SumaHoriz(ContNotas) = SumaHoriz(ContNotas) + Total
            PCol = PCol - 0.8
            If NumFacturas > 0 Then
               PictPrint_Texto PictCalif, PCol, PosLinea, Format(NumFacturas, "00")
               NumMeses = NumMeses + 1
               Cuota_No = Cuota_No + NumFacturas
            End If
            PCol = PCol + 0.8
            If Faltas_Just > 0 Then PictPrint_Texto PictCalif, PCol, PosLinea, Format(Faltas_Just, "00")
            PCol = PCol + 0.8
            If Faltas_Injust > 0 Then PictPrint_Texto PictCalif, PCol, PosLinea, Format(Faltas_Injust, "00")
            PCol = PCol + 0.8
            If Atrasos > 0 Then PictPrint_Texto PictCalif, PCol, PosLinea, Format(Atrasos, "00")
            PCol = PCol + 0.8
            PictPrint_Texto PictCalif, PCol, PosLinea, Format(Total, "00")
            PCol = PCol + 0.8
            PosLinea = PosLinea + 0.45
            Contador = Contador + 1
            PictCalif.Line (0.7, PosLinea)-(PCol, PosLinea)
            PosLinea = PosLinea + 0.05
            PictPrint_Texto PictCalif, 0.8, PosLinea, Format(Contador, "00") & ".-"
            PictPrint_Texto PictCalif, 1.3, PosLinea, .Fields("Cliente")
            PCol = 10.25
            Codigo = .Fields("Codigo")
            NombreCliente = .Fields("Cliente")
            For I = 0 To ContNotas - 1
                VectNota(I) = 0
            Next I
            I = 0
         End If
'         If OpcPeriodo("PQBim1", LstPeriodos) Then VectNota(I) = .Fields("ConductaPQ1")
'         If OpcPeriodo("PQBim2", LstPeriodos) Then VectNota(I) = .Fields("ConductaPQ2")
'         If OpcPeriodo("SQBim1", LstPeriodos) Then VectNota(I) = .Fields("ConductaSQ1")
'         If OpcPeriodo("SQBim2", LstPeriodos) Then VectNota(I) = .Fields("ConductaSQ2")
'         If OpcPeriodo("TQBim1", LstPeriodos) Then VectNota(I) = .Fields("ConductaTQ1")
'         If OpcPeriodo("TQBim2", LstPeriodos) Then VectNota(I) = .Fields("ConductaTQ2")
         If VectNota(I) > 0 Then PictPrint_Texto PictCalif, PCol, PosLinea, Format(VectNota(I), "00")
         SumaHoriz(I) = SumaHoriz(I) + VectNota(I)
         I = I + 1
         PCol = PCol + 0.8
        .MoveNext
      Loop
      Faltas_Just = 0
      Faltas_Injust = 0
      Atrasos = 0
      If AdoPromedio1.Recordset.RecordCount > 0 Then
         AdoPromedio1.Recordset.MoveFirst
         AdoPromedio1.Recordset.Find ("Codigo = '" & Codigo & "' ")
         If Not AdoPromedio1.Recordset.EOF Then
            If OpcPeriodo("PQBim1", LstPeriodos) Then
               NumFacturas = AdoPromedio1.Recordset.Fields("ConductaPQ1")
               Faltas_Just = AdoPromedio1.Recordset.Fields("PQBFJ1")
               Faltas_Injust = AdoPromedio1.Recordset.Fields("PQBFI1")
               Atrasos = AdoPromedio1.Recordset.Fields("PQBA1")
            End If
            If OpcPeriodo("PQBim2", LstPeriodos) Then
               NumFacturas = AdoPromedio1.Recordset.Fields("ConductaPQ2")
               Faltas_Just = AdoPromedio1.Recordset.Fields("PQBFJ2")
               Faltas_Injust = AdoPromedio1.Recordset.Fields("PQBFI2")
               Atrasos = AdoPromedio1.Recordset.Fields("PQBA2")
            End If
            If OpcPeriodo("SQBim1", LstPeriodos) Then
               NumFacturas = AdoPromedio1.Recordset.Fields("ConductaSQ1")
               Faltas_Just = AdoPromedio1.Recordset.Fields("SQBFJ1")
               Faltas_Injust = AdoPromedio1.Recordset.Fields("SQBFI1")
               Atrasos = AdoPromedio1.Recordset.Fields("SQBA1")
            End If
            If OpcPeriodo("SQBim2", LstPeriodos) Then
               NumFacturas = AdoPromedio1.Recordset.Fields("ConductaSQ2")
               Faltas_Just = AdoPromedio1.Recordset.Fields("SQBFJ2")
               Faltas_Injust = AdoPromedio1.Recordset.Fields("SQBFI2")
               Atrasos = AdoPromedio1.Recordset.Fields("SQBA2")
            End If
         End If
      End If
      Total = 0
      J = 0
      For I = 0 To ContNotas - 1
          If VectNota(I) > 0 Then
             J = J + 1
             Total = Total + VectNota(I)
          End If
      Next I
      If J = 0 Then
         Total = 0
      Else
         Total = Total / J
      End If
      Total = Round((Total + NumFacturas) / 2)
      SumaHoriz(ContNotas) = SumaHoriz(ContNotas) + Total
      PCol = PCol - 0.8
      If NumFacturas > 0 Then
         PictPrint_Texto PictCalif, PCol, PosLinea, Format(NumFacturas, "00")
         NumMeses = NumMeses + 1
         Cuota_No = Cuota_No + NumFacturas
      End If
      
      PCol = PCol + 0.8
      If Faltas_Just > 0 Then PictPrint_Texto PictCalif, PCol, PosLinea, Format(Faltas_Just, "00")
      PCol = PCol + 0.8
      If Faltas_Injust > 0 Then PictPrint_Texto PictCalif, PCol, PosLinea, Format(Faltas_Injust, "00")
      PCol = PCol + 0.8
      If Atrasos > 0 Then PictPrint_Texto PictCalif, PCol, PosLinea, Format(Atrasos, "00")
      PCol = PCol + 0.8
      PictPrint_Texto PictCalif, PCol, PosLinea, Format(Total, "00")
      PCol = PCol + 0.8
      PosLinea = PosLinea + 0.45
      PictCalif.Line (0.7, PosLinea)-(PCol, PosLinea)
      PictCalif.Line (0.7, 7.15)-(PCol, 7.15)
      PictCalif.Line (0.7, 4.3)-(PCol, 4.3)
      PictCalif.Line (PCol, 4.3)-(PCol, PosLinea)
      PictCalif.Line (0.7, 4.3)-(0.7, PosLinea)
      PCol = 9.9
      For I = 0 To ContNotas - 1
        PictCalif.Line (PCol, 4.3)-(PCol, PosLinea)
        PCol = PCol + 0.8
      Next I
      PictCalif.FontBold = True
      PCol = 10
      PictPrint_Texto PictCalif, PCol - 3.5, PosLinea, "TOTAL PROMEDIOS:"
      For I = 0 To ContNotas - 1
        If SumaHoriz(I) > 0 Then
           SumaHoriz(I) = SumaHoriz(I) / Contador
           PictPrint_Texto PictCalif, PCol, PosLinea, Format(SumaHoriz(I), "00.00")
        End If
        PCol = PCol + 0.8
      Next I
      PCol = PCol - 4
      If NumMeses <= 0 Then NumMeses = 1
      PictPrint_Texto PictCalif, PCol, PosLinea, Format(Cuota_No / NumMeses, "00.00")
      PCol = PCol + 4
      PCol = PCol - 0.55
      If SumaHoriz(ContNotas) > 0 Then
         SumaHoriz(ContNotas) = SumaHoriz(ContNotas) / Contador
         PictPrint_Texto PictCalif, PCol, PosLinea, Format(SumaHoriz(ContNotas), "00.00")
      End If
  End If
 End With
End Sub

Public Sub Listar_Calificacion_Del_Curso(TipoObjeto As Object, _
                                         Optional Lst_Supletorio As Boolean, _
                                         Optional SubNotas As Boolean)
Dim No_Disciplina As Currency
Dim CantProm As Byte
Dim TotalPQ As Currency
Dim TotalSQ As Currency
Dim Promedio As Currency
 'MsgBox ContNotas
  TipoObjeto.Visible = False
  Leer_Datos_del_Curso Dato_Curso.Curso, , SubNotas
  Listar_Alumnos_Notas_Del Dato_Curso.Curso, , Lst_Supletorio, SubNotas
  If TypeOf TipoObjeto Is PictureBox Then TipoObjeto.Cls
  TipoObjeto.width = 21
  TipoObjeto.Height = 29.7
  With AdoResumen.Recordset
  If .RecordCount > 0 Then
    'MsgBox UBound(VectMateria) & vbCrLf & ContNotas
     NombreBanco = UCase(LstPeriodos.Text)
     PCol = 10
     TipoObjeto.FontBold = True
     TipoObjeto.FontName = TipoVerdana
     TipoObjeto.width = 12 + Dato_Curso.ContMat
     If TipoObjeto.width < 21 Then TipoObjeto.width = 21
     If TypeOf TipoObjeto Is PictureBox Then
        TipoObjeto.Height = (CantAlumn * 0.5) + 10.5
        If TipoObjeto.Height < 29.7 Then TipoObjeto.Height = 29.7
     End If
     PosLinea = 0.5
     If LogoTipo <> "" Then TipoObjeto.PaintPicture LoadPicture(LogoTipo), 0.1, PosLinea, 5, 2.5
     TipoObjeto.FontSize = 16
     TipoObjeto.FontBold = True
     PictPrint_Texto 1, PosLinea, Institucion1, , TipoObjeto.width, True
     PosLinea = PosLinea + 0.7
     PictPrint_Texto 1, PosLinea, Institucion2, , TipoObjeto.width, True
     PosLinea = PosLinea + 0.7
     TipoObjeto.FontSize = 10
     TipoObjeto.FontBold = False
     PictPrint_Texto 1, PosLinea, ULCase(NombreCiudad), , TipoObjeto.width, True
     TipoObjeto.FontBold = True
     PosLinea = PosLinea + 0.5
     TipoObjeto.FontSize = 14
     If Lst_Supletorio Then
        PictPrint_Texto 1, PosLinea, "REPORTE DE ESTUDIANTES A SUPLETORIO/REMEDIAL", , TipoObjeto.width, True
     Else
        PictPrint_Texto 1, PosLinea, "ACTA DE CALIFICACIONES", , TipoObjeto.width, True
     End If
     TipoObjeto.FontBold = False
     PosLinea = PosLinea + 0.7
     TipoObjeto.FontSize = 10
     PictPrint_Texto 1, PosLinea, "AÑO LECTIVO: " & Anio_Lectivo, , TipoObjeto.width, True
     PosLinea = PosLinea + 0.5
     TipoObjeto.FontSize = 9
     PictPrint_Texto 0.8, PosLinea, NombreBanco
     PosLinea = PosLinea + 0.5
     PosLinea = PictPrint_Texto_Multiple(TipoObjeto, 0.8, PosLinea, CuentaBanco, 20)
     TipoObjeto.FontSize = 28
     PictPrint_Texto 1.3, PosLinea + 1, "A L U M N O S"
     TipoObjeto.FontName = TipoArialNarrow
     TipoObjeto.FontSize = 10
     PCol = 8.95
     PosLinea = 6.8
     
    'Encabezado de las materias
     For I = 1 To Dato_Curso.ContMat
'         If I < (ContNotas - 1) Then
            Cadena = Replace(Dato_Curso.Materia(I), "-", "")
            Cadena1 = SinEspaciosIzq(Cadena)
            Cadena = Trim(Mid$(Cadena, Len(Cadena1) + 1, Len(Cadena)))
            Cadena2 = SinEspaciosIzq(Cadena)
            Cadena = Trim(Mid$(Cadena, Len(Cadena2) + 1, Len(Cadena)))
            Cadena3 = SinEspaciosIzq(Cadena)
            If Len(Cadena2) <= 1 And Len(Cadena3) <= 1 Then
               PCol = PCol + 0.3
               cPrint.printTextoAngulo PCol, PosLinea, 90, 5, TipoObjeto.FontSize, Cadena1
               PCol = PCol + 0.7
            ElseIf Len(Cadena3) <= 1 Then
               PCol = PCol + 0.15
               cPrint.printTextoAngulo PCol, PosLinea, 90, 5, TipoObjeto.FontSize, Cadena1
               PCol = PCol + 0.3
               cPrint.printTextoAngulo PCol, PosLinea, 90, 5, TipoObjeto.FontSize, Cadena2
               PCol = PCol + 0.55
            Else
               cPrint.printTextoAngulo PCol, PosLinea, 90, 5, TipoObjeto.FontSize, Cadena1
               PCol = PCol + 0.3
               cPrint.printTextoAngulo PCol, PosLinea, 90, 5, TipoObjeto.FontSize, Cadena2
               PCol = PCol + 0.3
               cPrint.printTextoAngulo PCol, PosLinea, 90, 5, TipoObjeto.FontSize, Cadena3
               PCol = PCol + 0.4
            End If
'         Else
'            Cadena = Replace(VectMateria(I).Materias, "-", "")
'            PCol = PCol + 0.2
'            cPrint.printTextoAngulo PCol, PosLinea, 90, 6.5, TipoObjeto.FontSize + 7, Cadena
'            PCol = PCol + 0.8
'         End If
     Next
     PCol = PCol + 0.2
     cPrint.printTextoAngulo PCol, PosLinea, 90, 5, TipoObjeto.FontSize, "S U M A"
     PCol = PCol + 0.4
     cPrint.printTextoAngulo PCol, PosLinea, 90, 5, TipoObjeto.FontSize, "T O T A L"
     PCol = PCol + 0.55
     
     PCol = PCol + 0.2
     cPrint.printTextoAngulo PCol, PosLinea, 90, 5, TipoObjeto.FontSize, "PROMEDIO"
     PCol = PCol + 0.4
     cPrint.printTextoAngulo PCol, PosLinea, 90, 5, TipoObjeto.FontSize, "GENERAL"
     PCol = PCol + 0.55
     
     PosLinea = PosLinea + 0.05
     TipoObjeto.Line (0.7, PosLinea)-(PCol - 0.05, PosLinea)
     PosLinea = PosLinea + 0.05
     PCol = 9
     Contador = 0
     For I = 1 To Dato_Curso.ContAlumnos
         Dato_Curso.NotaPQ(I) = 0
         Dato_Curso.NotaSQ(I) = 0
         Dato_Curso.NotaTQ(I) = 0
         Dato_Curso.NotaFinal(I) = 0
     Next I
     For I = 0 To Dato_Curso.ContMat
         SumaHoriz(I) = 0
         SumaHorizT(I) = 0
     Next I
     I = 1
     NumMeses = 0
     Cuota_No = 0
     CantProm = 0
     TipoObjeto.FontBold = False
     Si_No = True
    'Empezamos a llenar los datos de las materias
     .MoveFirst
      Codigo = .Fields("Codigo")
      NombreCliente = .Fields("Cliente")
      Contador = Contador + 1
      TipoObjeto.FontSize = 8
      TipoObjeto.FontName = TipoArialNarrow
      PictPrint_Texto 0.8, PosLinea, Format(Contador, "00") & ".-"
      PictPrint_Texto 1.35, PosLinea, .Fields("Cliente")
      TipoObjeto.FontName = TipoVerdana
      J = 0
      Do While Not .EOF
         
         If Codigo <> .Fields("Codigo") Then
            'If Codigo = "0355200364" Then MsgBox "..."
            Total = 0
            TotalPQ = 0
            TotalSQ = 0
           'MsgBox OpcPeriodo("PF", LstPeriodos)
            If J > 0 Then
               If OpcPeriodo("PF", LstPeriodos) Then
                  For I = 0 To J - 1
                      TotalPQ = TotalPQ + SumaHorizPQ(I)
                      TotalSQ = TotalSQ + SumaHorizSQ(I)
                      Total = Total + SumaHoriz(I) 'SumaHorizPQ(I) + SumaHorizSQ(I)
                  Next I
                  TotalPQ = Redondear_2Dec(TotalPQ / J)
                  TotalSQ = Redondear_2Dec(TotalSQ / J)
                  Saldo = Redondear_2Dec(Total / J)
               Else
                  For I = 0 To J - 1
                      Total = Total + SumaHoriz(I)
                  Next I
                  Saldo = Redondear_2Dec(Total / J)
               End If
            Else
               Saldo = 0
            End If
            If Total > 0 Then
               PictPrint_Variables PCol + 0.1, PosLinea, Total, True, 1       'Sumatoria Total
               PCol = PCol + 1.2
               PictPrint_Nota_Materia PCol + 0.1, PosLinea, Saldo, False, 2   'Promedio Notas
            End If
            Contador = Contador + 1
            PosLinea = PosLinea + 0.45
            PCol = PCol + 1
            TipoObjeto.Line (0.7, PosLinea)-(PCol, PosLinea)
            PosLinea = PosLinea + 0.05
            TipoObjeto.FontName = TipoArialNarrow
            PictPrint_Texto 0.8, PosLinea, Format(Contador, "00") & ".-"
            PictPrint_Texto 1.35, PosLinea, .Fields("Cliente")
            TipoObjeto.FontName = TipoVerdana
            PCol = 9
            Codigo = .Fields("Codigo")
            NombreCliente = .Fields("Cliente")
            For I = 1 To Dato_Curso.ContAlumnos
                Dato_Curso.NotaPQ(I) = 0
                Dato_Curso.NotaSQ(I) = 0
                Dato_Curso.NotaTQ(I) = 0
                Dato_Curso.NotaFinal(I) = 0
            Next I
            For I = 0 To 49
                SumaHoriz(I) = 0
            Next I
            I = 1
            CantProm = 0
            Si_No = True
            J = 0
         End If
         
         If OpcPeriodo("PQBim1", LstPeriodos) Then Dato_Curso.NotaFinal(I) = .Fields("PQBim1")
         If OpcPeriodo("PQBim2", LstPeriodos) Then Dato_Curso.NotaFinal(I) = .Fields("PQBim2")
         If OpcPeriodo("PQBim3", LstPeriodos) Then Dato_Curso.NotaFinal(I) = .Fields("PQBim3")
         If OpcPeriodo("SQBim1", LstPeriodos) Then Dato_Curso.NotaFinal(I) = .Fields("SQBim1")
         If OpcPeriodo("SQBim2", LstPeriodos) Then Dato_Curso.NotaFinal(I) = .Fields("SQBim2")
         If OpcPeriodo("SQBim3", LstPeriodos) Then Dato_Curso.NotaFinal(I) = .Fields("SQBim3")
         If OpcPeriodo("TQBim1", LstPeriodos) Then Dato_Curso.NotaFinal(I) = .Fields("TQBim1")
         If OpcPeriodo("TQBim2", LstPeriodos) Then Dato_Curso.NotaFinal(I) = .Fields("TQBim2")
         If OpcPeriodo("TQBim3", LstPeriodos) Then Dato_Curso.NotaFinal(I) = .Fields("TQBim3")
         If OpcPeriodo("ExamenPQ", LstPeriodos) Then Dato_Curso.NotaFinal(I) = .Fields("ExamenPQ")
         If OpcPeriodo("ExamenSQ", LstPeriodos) Then Dato_Curso.NotaFinal(I) = .Fields("ExamenSQ")
         If OpcPeriodo("PQ", LstPeriodos) Then Dato_Curso.NotaFinal(I) = .Fields("PromPQ")
         If OpcPeriodo("SQ", LstPeriodos) Then Dato_Curso.NotaFinal(I) = .Fields("PromSQ")
         If OpcPeriodo("TQ", LstPeriodos) Then Dato_Curso.NotaFinal(I) = .Fields("PromTQ")
         If OpcPeriodo("PF", LstPeriodos) Then
            Dato_Curso.NotaFinal(I) = .Fields("PromFinal")
            If .Fields("C2") Then
                Dato_Curso.NotaPQ(I) = .Fields("ExamenPQ")
                Dato_Curso.NotaSQ(I) = .Fields("ExamenSQ")
            Else
                Dato_Curso.NotaPQ(I) = .Fields("PromPQ")
                Dato_Curso.NotaSQ(I) = .Fields("PromSQ")
            End If
         End If
        ' MsgBox J & vbCrLf & Dato_Curso.NotaFinal(I)
         Dato_Curso.NotaFinal(I) = Redondear(Dato_Curso.NotaFinal(I), 2)
        'Imprime nota a nota por materia
        'If Codigo = "0355200364" Then MsgBox .Fields("CodMat") & " = " & Dato_Curso.NotaFinal(I) & "..."
         If Dato_Curso.NotaFinal(I) > 0 Then
            If Val(.Fields("CodMat")) < 997 Then
               If .Fields("C") Then
                   PictPrint_Nota_Materia PCol + 0.3, PosLinea, Dato_Curso.NotaFinal(I), True, Dec_Nota
               Else
                   PictPrint_Nota_Materia PCol, PosLinea, Dato_Curso.NotaFinal(I), , Dec_Nota
                   If Not .Fields("C2") Then
                      SumaHoriz(J) = Dato_Curso.NotaFinal(I)
                      SumaHorizT(I) = SumaHorizT(I) + Dato_Curso.NotaFinal(I)
                      SumaHorizPQ(J) = Dato_Curso.NotaPQ(I)
                      SumaHorizSQ(J) = Dato_Curso.NotaSQ(I)
                      J = J + 1
                   End If
               End If
            Else
               PictPrint_Nota_Materia PCol + 0.3, PosLinea, Dato_Curso.NotaFinal(I), True, Dec_Nota
            End If
         End If
         'SumaHoriz(I) = SumaHoriz(I) + Dato_Curso.NotaFinal(I)
         I = I + 1
         PCol = PCol + 1
        .MoveNext
      Loop
     'Ultimo alumno
        Total = 0
        TotalPQ = 0
        TotalSQ = 0
        If J > 0 Then
           If OpcPeriodo("PF", LstPeriodos) Then
              For I = 0 To J - 1
                  TotalPQ = TotalPQ + SumaHorizPQ(I)
                  TotalSQ = TotalSQ + SumaHorizSQ(I)
                  Total = Total + SumaHoriz(I) 'SumaHorizPQ(I) + SumaHorizSQ(I)
              Next I
              TotalPQ = Redondear(TotalPQ / J, 2)
              TotalSQ = Redondear(TotalSQ / J, 2)
              Saldo = Redondear_2Dec(Total / J)
           Else
              For I = 0 To J - 1
                  Total = Total + SumaHoriz(I)
              Next I
              Saldo = Redondear(Total / J, 2)
           End If
        Else
           Saldo = 0
        End If
      If Total > 0 Then
         PictPrint_Variables PCol + 0.1, PosLinea, Total, True, 1
         PCol = PCol + 1.2
         PictPrint_Nota_Materia PCol + 0.1, PosLinea, Saldo, False, 2
      End If
      Contador = Contador + 1
      PosLinea = PosLinea + 0.45
      PCol = PCol + 1
     'Impresion de las rayas del reporte de calificaciones
      PCol = 8.9
      For I = 1 To Dato_Curso.ContMat
        TipoObjeto.Line (PCol, 4.5)-(PCol, PosLinea)
        PCol = PCol + 1
      Next I
      PCol = PCol + 0.1
      TipoObjeto.Line (PCol, 4.5)-(PCol, PosLinea)
      PCol = PCol + 1.2
      TipoObjeto.Line (PCol, 4.5)-(PCol, PosLinea)
      PCol = PCol + 1
      TipoObjeto.Line (PCol, 4.5)-(PCol, PosLinea)

      TipoObjeto.Line (0.7, PosLinea)-(PCol, PosLinea)
      TipoObjeto.Line (0.7, 4.5)-(PCol, 4.5)
      TipoObjeto.Line (0.7, 4.5)-(0.7, PosLinea)
      TipoObjeto.FontBold = True
  Else
      PosLinea = 1
      TipoObjeto.FontSize = 24
      TipoObjeto.FontBold = True
      PictPrint_Texto 1, PosLinea, "NO EXISTEN DATOS A MOSTRAR", , TipoObjeto.width, True
      PosLinea = PosLinea + 2
      PCol = 10
  End If
 End With
 Codigo4 = SinEspaciosIzq(DLCurso)
 PosLinea = PosLinea + 0.05
 PCol = 8.7
 Contador = 0
 Sumatoria = 0
 For I = 1 To Dato_Curso.ContMat
     PictPrint_Nota_Materia PCol + 0.3, PosLinea, SumaHorizT(I) / Dato_Curso.ContAlumnos, False, Dec_Nota
     If SumaHorizT(I) > 0 Then
        Contador = Contador + 1
        Sumatoria = Sumatoria + (SumaHorizT(I) / Dato_Curso.ContAlumnos)
     End If
     PCol = PCol + 1
 Next I
 PCol = PCol + 1.3
 If Contador <> 0 Then PictPrint_Nota_Materia PCol + 0.3, PosLinea, Sumatoria / Contador, False, Dec_Nota
 PCol = 21
 TipoObjeto.FontSize = 10
 PosLinea = PosLinea + 0.5
 Cadena = FechaStrgCiudad(MBFecha)
 PictPrint_Texto PCol - TipoObjeto.TextWidth(Cadena), PosLinea, Cadena
 If Not Lst_Supletorio Then
    PosLinea = PosLinea + 1.9
    Select Case Codigo4
       Case "0.00" To "1.99"
            PictPrint_Texto 1, PosLinea, Director             '31.8
            PictPrint_Texto 16, PosLinea, Secretario1
            PosLinea = PosLinea + 0.3
            PictPrint_Texto 1, PosLinea, TextoDirector        '31.5
            PictPrint_Texto 16, PosLinea, TextoSecretario1
       Case "2.00" To "3.99"
            PictPrint_Texto 1, PosLinea, Rector
            PictPrint_Texto 16, PosLinea, Secretario2
            PosLinea = PosLinea + 0.35
            PictPrint_Texto 1, PosLinea, TextoRector
            PictPrint_Texto 16, PosLinea, TextoSecretario2
       Case "4.00" To "5.99"
            PictPrint_Texto 1, PosLinea, Rector
            PictPrint_Texto 16, PosLinea, Secretario3
            PosLinea = PosLinea + 0.35
            PictPrint_Texto 1, PosLinea, TextoRector
            PictPrint_Texto 16, PosLinea, TextoSecretario2
     End Select
     'If TypeOf TipoObjeto Is PictureBox Then TipoObjeto.Height = 2 + PosLinea
  End If
  TipoObjeto.Visible = True
 'MsgBox TipoObjeto.Height
End Sub

Public Sub Listar_Aprobados_Reprobados()
Dim No_Disciplina As Currency
 'Recolectamos los Codigos y Nombre de las Materias
  Listar_Alumnos_Notas_Del SinEspaciosIzq(DLCurso), True
''  With AdoResumen.Recordset
''   If .RecordCount > 0 Then
''      .MoveFirst
''       ContNotas = 0
''       Codigo = .Fields("Codigo")
''       CodigoCli = .Fields("Codigo")
''       CuentaBanco = .Fields("Descripcion")
''       Do While Not .EOF
''          If Codigo <> .Fields("Codigo") Then Exit Do
''          VectMate(ContNotas) = .Fields("Materia")
''          VectCodMat(ContNotas) = .Fields("CodMat")
''          ContNotas = ContNotas + 1
''         .MoveNext
''       Loop
''       VectMate(ContNotas) = "COMPORTAMENTAL"
''       ContNotas = ContNotas + 1
''      .MoveFirst
''   End If
''  End With
  
 NombreBanco = UCase(LstPeriodos.Text)
 PictCalif.Cls
 PictCalif.FontBold = True
 PCol = 10
 PictCalif.FontName = TipoVerdana
 PictCalif.width = PCol + (ContNotas * 1.2)
 'PictCalif.Height = 12 + (AdoPromedio1.Recordset.RecordCount * 0.45)
 'MsgBox AdoPromedio1.Recordset.RecordCount
 PosLinea = 0.5
 If LogoTipo <> "" Then PictCalif.PaintPicture LoadPicture(LogoTipo), 0.1, PosLinea, 5, 2.5
 PictCalif.FontSize = 20
 PictCalif.FontBold = True
 PictPrint_Texto PictCalif, 1, PosLinea, Institucion1, , PictCalif.width, True
 PosLinea = PosLinea + 0.7
 PictPrint_Texto PictCalif, 1, PosLinea, Institucion2, , PictCalif.width, True
 PosLinea = PosLinea + 0.7
 PictCalif.FontSize = 12
 PictPrint_Texto PictCalif, 1, PosLinea, NombreCiudad, , PictCalif.width, True
 PosLinea = PosLinea + 0.6
 PictCalif.FontSize = 16
 
 PictPrint_Texto PictCalif, 1, PosLinea, "ACTA DE CALIFICACIONES", , PictCalif.width, True
 PictCalif.FontBold = False
 PosLinea = PosLinea + 0.7
 PictCalif.FontSize = 12
 PictPrint_Texto PictCalif, 1, PosLinea, "AÑO LECTIVO: " & Anio_Lectivo, , PictCalif.width, True
 PosLinea = PosLinea + 0.6
 PictPrint_Texto PictCalif, 0.8, PosLinea, NombreBanco
 PictPrint_Texto PictCalif, PictCalif.width - PictCalif.TextWidth(CuentaBanco) - 1.2, PosLinea, CuentaBanco
 PictCalif.FontSize = 28
 PictPrint_Texto PictCalif, 1.3, PosLinea + 1, "A L U M N O S"
 PictCalif.FontName = TipoArialNarrow
 PictCalif.FontSize = 10
 PCol = 9
 PosLinea = 6.6
 For I = 0 To ContNotas - 1
     If I <> (ContNotas - 1) Then
     Cadena = Replace(VectMate(I), "-", "")
     Cadena1 = SinEspaciosIzq(Cadena)
     Cadena = Trim(Mid$(Cadena, Len(Cadena1) + 1, Len(Cadena)))
     Cadena2 = SinEspaciosIzq(Cadena)
     Cadena = Trim(Mid$(Cadena, Len(Cadena2) + 1, Len(Cadena)))
     Cadena3 = SinEspaciosIzq(Cadena)
     PCol = PCol + 0.05
     cPrint.printTextoAngulo PictCalif, PCol, PosLinea, 90, 5, PictCalif.FontSize, Cadena1
     PCol = PCol + 0.3
     cPrint.printTextoAngulo PictCalif, PCol, PosLinea, 90, 5, PictCalif.FontSize, Cadena2
     PCol = PCol + 0.3
     cPrint.printTextoAngulo PictCalif, PCol, PosLinea, 90, 5, PictCalif.FontSize, Cadena3
     PCol = PCol + 0.55
     Else
     Cadena = Replace(VectMate(I), "-", "")
     PCol = PCol + 0.3
     cPrint.printTextoAngulo PictCalif, PCol, PosLinea, 90, 3.45, PictCalif.FontSize + 6, Cadena
     PCol = PCol + 0.8
     End If
 Next
 PictCalif.FontName = TipoVerdana
 PictCalif.FontSize = 9
 PosLinea = PosLinea + 0.05
 PictCalif.Line (0.7, PosLinea)-(PCol, PosLinea)
 PosLinea = PosLinea + 0.05
 PCol = 9
 Contador = 0
 For I = 0 To ContNotas - 1
     VectNota(I) = 0
     SumaHoriz(I) = 0
 Next I
 I = 0
 NumMeses = 0
 Cuota_No = 0
 PictCalif.FontBold = False
 With AdoResumen.Recordset
  If .RecordCount > 0 Then
     .MoveFirst
      Codigo = .Fields("Codigo")
      NombreCliente = .Fields("Cliente")
      Contador = Contador + 1
      PictCalif.FontName = TipoArialNarrow
      PictPrint_Texto PictCalif, 0.8, PosLinea, Format(Contador, "00") & ".-"
      PictPrint_Texto PictCalif, 1.35, PosLinea, .Fields("Cliente")
      PictCalif.FontName = TipoVerdana
      Do While Not .EOF
         If Codigo <> .Fields("Codigo") Then
            Total = 0
            J = 0
            For I = 0 To ContNotas - 1
                If VectNota(I) > 0 Then
                   J = J + 1
                   Total = Total + VectNota(I)
                End If
            Next I
            If J = 0 Then
               Total = 0
            Else
               Total = Total / J
            End If
            Total = Redondear((Total + NumFacturas) / 2, 2)
            'MsgBox NumFacturas
            SumaHoriz(ContNotas) = SumaHoriz(ContNotas) + Total
            PCol = PCol - 0.8
            If NumFacturas > 0 Then
               'PictPrint_Texto PictCalif, PCol, PosLinea, Format(NumFacturas, "00.00")
               NumMeses = NumMeses + 1
               Cuota_No = Cuota_No + NumFacturas
            End If
            No_Disciplina = 0
            Total = 0
            If AdoConducta.Recordset.RecordCount > 0 Then
               AdoConducta.Recordset.MoveFirst
               AdoConducta.Recordset.Find ("Codigo = '" & Codigo & "' ")
               If Not AdoConducta.Recordset.EOF Then
                  'MsgBox "...."
                  If OpcPeriodo("PQ", LstPeriodos) Then
                     If AdoConducta.Recordset.Fields("ConductaPQ1") > 0 Then No_Disciplina = No_Disciplina + 1
                     If AdoConducta.Recordset.Fields("ConductaPQ2") > 0 Then No_Disciplina = No_Disciplina + 1
                  ElseIf OpcPeriodo("SQ", LstPeriodos) Then
                     If AdoConducta.Recordset.Fields("ConductaSQ1") > 0 Then No_Disciplina = No_Disciplina + 1
                     If AdoConducta.Recordset.Fields("ConductaSQ2") > 0 Then No_Disciplina = No_Disciplina + 1
                  ElseIf OpcPeriodo("TQ", LstPeriodos) Then
                     If AdoConducta.Recordset.Fields("ConductaTQ1") > 0 Then No_Disciplina = No_Disciplina + 1
                     If AdoConducta.Recordset.Fields("ConductaTQ2") > 0 Then No_Disciplina = No_Disciplina + 1
                  ElseIf OpcPeriodo("PF", LstPeriodos) Then
                     If AdoConducta.Recordset.Fields("ConductaPQ1") > 0 Then No_Disciplina = No_Disciplina + 1
                     If AdoConducta.Recordset.Fields("ConductaPQ2") > 0 Then No_Disciplina = No_Disciplina + 1
                     If AdoConducta.Recordset.Fields("ConductaSQ1") > 0 Then No_Disciplina = No_Disciplina + 1
                     If AdoConducta.Recordset.Fields("ConductaSQ2") > 0 Then No_Disciplina = No_Disciplina + 1
                     If AdoConducta.Recordset.Fields("ConductaTQ1") > 0 Then No_Disciplina = No_Disciplina + 1
                     If AdoConducta.Recordset.Fields("ConductaTQ2") > 0 Then No_Disciplina = No_Disciplina + 1
                  End If
                  If OpcPeriodo("PQBim1", LstPeriodos) Then
                     SubTotal = AdoConducta.Recordset.Fields("ConductaPQ1")
                  ElseIf OpcPeriodo("PQBim2", LstPeriodos) Then
                     SubTotal = AdoConducta.Recordset.Fields("ConductaPQ2")
                  ElseIf OpcPeriodo("SQBim1", LstPeriodos) Then
                     SubTotal = AdoConducta.Recordset.Fields("ConductaSQ1")
                  ElseIf OpcPeriodo("SQBim2", LstPeriodos) Then
                     SubTotal = AdoConducta.Recordset.Fields("ConductaSQ2")
                  ElseIf OpcPeriodo("TQBim1", LstPeriodos) Then
                     SubTotal = AdoConducta.Recordset.Fields("ConductaTQ1")
                  ElseIf OpcPeriodo("TQBim2", LstPeriodos) Then
                     SubTotal = AdoConducta.Recordset.Fields("ConductaTQ2")
                  ElseIf OpcPeriodo("PQ", LstPeriodos) Then
                     SubTotal = AdoConducta.Recordset.Fields("ConductaPQ1") _
                              + AdoConducta.Recordset.Fields("ConductaPQ2")
                  ElseIf OpcPeriodo("SQ", LstPeriodos) Then
                     SubTotal = AdoConducta.Recordset.Fields("ConductaSQ1") _
                              + AdoConducta.Recordset.Fields("ConductaSQ2")
                  ElseIf OpcPeriodo("TQ", LstPeriodos) Then
                     SubTotal = AdoConducta.Recordset.Fields("ConductaTQ1") _
                              + AdoConducta.Recordset.Fields("ConductaTQ2")
                  ElseIf OpcPeriodo("PF", LstPeriodos) Then
                     SubTotal = AdoConducta.Recordset.Fields("ConductaPQ1") _
                              + AdoConducta.Recordset.Fields("ConductaPQ2") _
                              + AdoConducta.Recordset.Fields("ConductaSQ1") _
                              + AdoConducta.Recordset.Fields("ConductaSQ2") _
                              + AdoConducta.Recordset.Fields("ConductaTQ1") _
                              + AdoConducta.Recordset.Fields("ConductaTQ2")
                  End If
                  If No_Disciplina <= 0 Then No_Disciplina = 1
                  Total = Redondear(SubTotal / No_Disciplina, 2)
               End If
            End If
            PCol = PCol + 0.8
            If Total <> 0 Then PictPrint_Nota_Materia PictCalif, PCol + 0.3, PosLinea, Total, True, Dec_Nota
            'PictPrint_Texto PictCalif, PCol, PosLinea, Equivalencia(Total)     'Format(Total, "00.00")
            PCol = PCol + 1.1
            Contador = Contador + 1
            PosLinea = PosLinea + 0.45
            PictCalif.Line (0.7, PosLinea)-(PCol, PosLinea)
            PosLinea = PosLinea + 0.05
            PictCalif.FontName = TipoArialNarrow
            PictPrint_Texto PictCalif, 0.8, PosLinea, Format(Contador, "00") & ".-"
            PictPrint_Texto PictCalif, 1.35, PosLinea, .Fields("Cliente")
            PictCalif.FontName = TipoVerdana
            PCol = 9
            Codigo = .Fields("Codigo")
            NombreCliente = .Fields("Cliente")
            For I = 0 To ContNotas - 1
                VectNota(I) = 0
            Next I
            I = 0
         End If
         If OpcPeriodo("PQBim1", LstPeriodos) Then VectNota(I) = .Fields("PQBim1")
         If OpcPeriodo("PQBim2", LstPeriodos) Then VectNota(I) = .Fields("PQBim2")
         If OpcPeriodo("SQBim1", LstPeriodos) Then VectNota(I) = .Fields("SQBim1")
         If OpcPeriodo("SQBim2", LstPeriodos) Then VectNota(I) = .Fields("SQBim2")
         If OpcPeriodo("TQBim1", LstPeriodos) Then VectNota(I) = .Fields("TQBim1")
         If OpcPeriodo("TQBim2", LstPeriodos) Then VectNota(I) = .Fields("TQBim2")
         If OpcPeriodo("PQ", LstPeriodos) Then VectNota(I) = .Fields("PromPQ")
         If OpcPeriodo("SQ", LstPeriodos) Then VectNota(I) = .Fields("PromSQ")
         If OpcPeriodo("TQ", LstPeriodos) Then VectNota(I) = .Fields("PromTQ")
         If OpcPeriodo("PF", LstPeriodos) Then VectNota(I) = .Fields("PromFinal")
         If VectNota(I) > 0 Then PictPrint_Nota_Materia PictCalif, PCol, PosLinea, VectNota(I), , Dec_Nota
         SumaHoriz(I) = SumaHoriz(I) + VectNota(I)
         I = I + 1
         PCol = PCol + 1.2
        .MoveNext
      Loop
      
     'Ultimo alumno
      Total = 0
      J = 0
      For I = 0 To ContNotas - 1
          If VectNota(I) > 0 Then
             J = J + 1
             Total = Total + VectNota(I)
          End If
      Next I
      If J = 0 Then
         Total = 0
      Else
         Total = Total / J
      End If
      Total = Redondear((Total + NumFacturas) / 2)
      SumaHoriz(ContNotas) = SumaHoriz(ContNotas) + Total
      PCol = PCol - 0.8
      If NumFacturas > 0 Then
         'PictPrint_Texto PictCalif, PCol, PosLinea, Format(NumFacturas, "00")
         NumMeses = NumMeses + 1
         Cuota_No = Cuota_No + NumFacturas
      End If
      Total = 0
      No_Disciplina = 0
      If AdoConducta.Recordset.RecordCount > 0 Then
         AdoConducta.Recordset.MoveFirst
         AdoConducta.Recordset.Find ("Codigo = '" & Codigo & "' ")
         If Not AdoConducta.Recordset.EOF Then
            If OpcPeriodo("PQ", LstPeriodos) Then
               If AdoConducta.Recordset.Fields("ConductaPQ1") > 0 Then No_Disciplina = No_Disciplina + 1
               If AdoConducta.Recordset.Fields("ConductaPQ2") > 0 Then No_Disciplina = No_Disciplina + 1
            ElseIf OpcPeriodo("SQ", LstPeriodos) Then
               If AdoConducta.Recordset.Fields("ConductaSQ1") > 0 Then No_Disciplina = No_Disciplina + 1
               If AdoConducta.Recordset.Fields("ConductaSQ2") > 0 Then No_Disciplina = No_Disciplina + 1
            ElseIf OpcPeriodo("TQ", LstPeriodos) Then
               If AdoConducta.Recordset.Fields("ConductaTQ1") > 0 Then No_Disciplina = No_Disciplina + 1
               If AdoConducta.Recordset.Fields("ConductaTQ2") > 0 Then No_Disciplina = No_Disciplina + 1
            ElseIf OpcPeriodo("PF", LstPeriodos) Then
               If AdoConducta.Recordset.Fields("ConductaPQ1") > 0 Then No_Disciplina = No_Disciplina + 1
               If AdoConducta.Recordset.Fields("ConductaPQ2") > 0 Then No_Disciplina = No_Disciplina + 1
               If AdoConducta.Recordset.Fields("ConductaSQ1") > 0 Then No_Disciplina = No_Disciplina + 1
               If AdoConducta.Recordset.Fields("ConductaSQ2") > 0 Then No_Disciplina = No_Disciplina + 1
               If AdoConducta.Recordset.Fields("ConductaTQ1") > 0 Then No_Disciplina = No_Disciplina + 1
               If AdoConducta.Recordset.Fields("ConductaTQ2") > 0 Then No_Disciplina = No_Disciplina + 1
            ElseIf OpcPeriodo("PQ", LstPeriodos) Then
               If AdoConducta.Recordset.Fields("ConductaPQ1") > 0 Then No_Disciplina = No_Disciplina + 1
               If AdoConducta.Recordset.Fields("ConductaPQ2") > 0 Then No_Disciplina = No_Disciplina + 1
            ElseIf OpcPeriodo("SQ", LstPeriodos) Then
               If AdoConducta.Recordset.Fields("ConductaSQ1") > 0 Then No_Disciplina = No_Disciplina + 1
               If AdoConducta.Recordset.Fields("ConductaSQ2") > 0 Then No_Disciplina = No_Disciplina + 1
            ElseIf OpcPeriodo("TQ", LstPeriodos) Then
               If AdoConducta.Recordset.Fields("ConductaTQ1") > 0 Then No_Disciplina = No_Disciplina + 1
               If AdoConducta.Recordset.Fields("ConductaTQ2") > 0 Then No_Disciplina = No_Disciplina + 1
            ElseIf OpcPeriodo("PF", LstPeriodos) Then
               If AdoConducta.Recordset.Fields("ConductaPQ1") > 0 Then No_Disciplina = No_Disciplina + 1
               If AdoConducta.Recordset.Fields("ConductaPQ2") > 0 Then No_Disciplina = No_Disciplina + 1
               If AdoConducta.Recordset.Fields("ConductaSQ1") > 0 Then No_Disciplina = No_Disciplina + 1
               If AdoConducta.Recordset.Fields("ConductaSQ2") > 0 Then No_Disciplina = No_Disciplina + 1
               If AdoConducta.Recordset.Fields("ConductaTQ1") > 0 Then No_Disciplina = No_Disciplina + 1
               If AdoConducta.Recordset.Fields("ConductaTQ2") > 0 Then No_Disciplina = No_Disciplina + 1
            End If
            If OpcPeriodo("PQBim1", LstPeriodos) Then
               SubTotal = AdoConducta.Recordset.Fields("ConductaPQ1")
            ElseIf OpcPeriodo("PQBim2", LstPeriodos) Then
               SubTotal = AdoConducta.Recordset.Fields("ConductaPQ2")
            ElseIf OpcPeriodo("SQBim1", LstPeriodos) Then
               SubTotal = AdoConducta.Recordset.Fields("ConductaSQ1")
            ElseIf OpcPeriodo("SQBim2", LstPeriodos) Then
               SubTotal = AdoConducta.Recordset.Fields("ConductaSQ2")
            ElseIf OpcPeriodo("TQBim1", LstPeriodos) Then
               SubTotal = AdoConducta.Recordset.Fields("ConductaTQ1")
            ElseIf OpcPeriodo("TQBim2", LstPeriodos) Then
               SubTotal = AdoConducta.Recordset.Fields("ConductaTQ2")
            ElseIf OpcPeriodo("PQ", LstPeriodos) Then
               SubTotal = AdoConducta.Recordset.Fields("ConductaPQ1") _
                        + AdoConducta.Recordset.Fields("ConductaPQ2")
            ElseIf OpcPeriodo("SQ", LstPeriodos) Then
               SubTotal = AdoConducta.Recordset.Fields("ConductaSQ1") _
                        + AdoConducta.Recordset.Fields("ConductaSQ2")
            ElseIf OpcPeriodo("TQ", LstPeriodos) Then
               SubTotal = AdoConducta.Recordset.Fields("ConductaTQ1") _
                        + AdoConducta.Recordset.Fields("ConductaTQ2")
            ElseIf OpcPeriodo("PF", LstPeriodos) Then
               SubTotal = AdoConducta.Recordset.Fields("ConductaPQ1") _
                        + AdoConducta.Recordset.Fields("ConductaPQ2") _
                        + AdoConducta.Recordset.Fields("ConductaSQ1") _
                        + AdoConducta.Recordset.Fields("ConductaSQ2") _
                        + AdoConducta.Recordset.Fields("ConductaTQ1") _
                        + AdoConducta.Recordset.Fields("ConductaTQ2")
            End If
            If No_Disciplina <= 0 Then No_Disciplina = 1
            Total = Redondear(SubTotal / No_Disciplina)
         End If
      End If
      PCol = PCol + 0.8
      If Total <> 0 Then PictPrint_Nota_Materia PictCalif, PCol + 0.3, PosLinea, Total, True, Dec_Nota
      PCol = PCol + 1.1
      PosLinea = PosLinea + 0.45
      PictCalif.Line (0.7, PosLinea)-(PCol, PosLinea)
      PictCalif.Line (0.7, 7.15)-(PCol, 7.15)
      PictCalif.Line (0.7, 4.3)-(PCol, 4.3)
      PictCalif.Line (PCol, 4.3)-(PCol, PosLinea)
      PictCalif.Line (0.7, 4.3)-(0.7, PosLinea)
      PCol = 8.9
      For I = 0 To ContNotas - 1
        PictCalif.Line (PCol, 4.3)-(PCol, PosLinea)
        PCol = PCol + 1.2
      Next I
      PictCalif.FontBold = True
      'PCol = 10
  End If
 End With
 Codigo4 = SinEspaciosIzq(DLCurso)
 PictCalif.FontSize = 10
 PosLinea = PosLinea + 0.05
 Cadena = FechaStrgCiudad(MBFecha)
 PictPrint_Texto PictCalif, PCol - PictCalif.TextWidth(Cadena), PosLinea, Cadena
 
 PosLinea = PosLinea + 1.9
 Select Case Codigo4
    Case "0.00" To "1.99"
         PictPrint_Texto PictCalif, 1, PosLinea, Director             '31.8
         PictPrint_Texto PictCalif, 16, PosLinea, Secretario1
         PosLinea = PosLinea + 0.3
         PictPrint_Texto PictCalif, 1, PosLinea, TextoDirector        '31.5
         PictPrint_Texto PictCalif, 16, PosLinea, TextoSecretario1
    Case "2.00" To "3.99"
         PictPrint_Texto PictCalif, 1, PosLinea, Rector
         PictPrint_Texto PictCalif, 16, PosLinea, Secretario2
         PosLinea = PosLinea + 0.35
         PictPrint_Texto PictCalif, 1, PosLinea, TextoRector
         PictPrint_Texto PictCalif, 16, PosLinea, TextoSecretario2
    Case "4.00" To "5.99"
         PictPrint_Texto PictCalif, 1, PosLinea, Rector
         PictPrint_Texto PictCalif, 16, PosLinea, Secretario3
         PosLinea = PosLinea + 0.35
         PictPrint_Texto PictCalif, 1, PosLinea, TextoRector
         PictPrint_Texto PictCalif, 16, PosLinea, TextoSecretario2
  End Select
  PictCalif.Height = 2 + PosLinea
 'MsgBox PictCalif.Height
End Sub

Public Sub Listar_Materias_x_Profesor()
Dim AnchoAux As Single
 'Notas por Materia de Profesor
  Codigo = SinEspaciosIzq(DLCurso)
  'Listar_Materias_Curso Codigo
''' Codigo4 = .Fields("Materia")
''' NombreDocente = .Fields("Profesores")

  sSQL = "SELECT C.Cliente,C.Sexo,TN.* " _
       & "FROM Trans_Notas As TN,Clientes As C " _
       & "WHERE TN.Item = '" & NumEmpresa & "' " _
       & "AND TN.Periodo = '" & Periodo_Contable & "' " _
       & "AND TN.CodE = '" & Codigo & "' " _
       & "AND TN.CodMat = '" & Codigo3 & "' " _
       & "AND TN.Codigo = C.Codigo " _
       & "ORDER BY C.Cliente,TN.Id_No,TN.Orden "
  SelectAdodc AdoNotas, sSQL

  With AdoMaterias.Recordset
    If .RecordCount > 0 Then
       .MoveFirst
       .Find ("Materia = '" & DCMaterias & "' ")
       If Not .EOF Then
          Codigo4 = .Fields("Materia")
          Codigo3 = .Fields("CodMat")
          NombreDocente = .Fields("Profesores")
       End If
    End If
  End With

 NombreBanco = UCase(LstPeriodos.Text)
 PictCalif.Cls
 PictCalif.Picture = LoadPicture(RutaSistema & "\FORMATOS\GENERAL\Pagina_A4.GIF")
 PictCalif.Refresh
 AnchoAux = PictCalif.width
 PictCalif.FontBold = True
 PCol = 10
 PictCalif.Height = 12 + (AdoPromedio1.Recordset.RecordCount * 0.45)
 PictCalif.width = PCol + (ContNotas * 0.9) + 1.1
 If PictCalif.width < AnchoAux Then PictCalif.width = AnchoAux
 PictCalif.FontName = TipoTimes
 PosLinea = 0.3
 If LogoTipo <> "" Then PictCalif.PaintPicture LoadPicture(LogoTipo), 0.1, PosLinea, 5, 2.5
 PictCalif.FontSize = 22
 PictPrint_Texto PictCalif, 1, PosLinea, Empresa, , PictCalif.width, True
 PosLinea = PosLinea + 0.9
 PictCalif.FontSize = 12
 PictCalif.FontBold = False
 PictPrint_Texto PictCalif, 1, PosLinea, ULCase(Direccion) & " Teléfono: " & Telefono1, , PictCalif.width, True
 PosLinea = PosLinea + 0.5
 PictCalif.FontSize = 16
 PictCalif.FontBold = True
 PictPrint_Texto PictCalif, 1, PosLinea, "C A L I F I C A C I O N E S    P O R    Q U I M E S T R E S", , PictCalif.width, True
 PictCalif.FontBold = False
 PosLinea = PosLinea + 0.6
 PictCalif.FontSize = 12
 PictPrint_Texto PictCalif, 1, PosLinea, "AÑO LECTIVO: " & Anio_Lectivo, , PictCalif.width, True
 PosLinea = PosLinea + 0.5
 PictPrint_Texto PictCalif, 0.8, PosLinea, CuentaBanco
 PictPrint_Texto PictCalif, 0.8, PosLinea + 0.6, DCMaterias
 PictPrint_Texto PictCalif, 0.8, PosLinea + 1.1, ULCase(NombreDocente)
 PictCalif.FontSize = 28
 PictPrint_Texto PictCalif, 1.7, PosLinea + 1.9, "A L U M N O S"
 PictCalif.FontSize = 10
 PictCalif.FontBold = True
 PictCalif.FontName = TipoArialNarrow
 PictPrint_Texto PictCalif, 10.4, PosLinea + 0.6, "PRIMER QUIMESTRE"
 PictPrint_Texto PictCalif, 14.2, PosLinea + 0.6, "SEGUNDO QUIMESTRE"
 PictCalif.FontSize = 9
 PCol = 10
 PosLinea = 5.8
 PCol = PCol + 0.2
 cPrint.printTextoAngulo PictCalif, PCol, PosLinea, 90, 5, PictCalif.FontSize, "Bimensual 1"
 PCol = PCol + 1
 cPrint.printTextoAngulo PictCalif, PCol, PosLinea, 90, 5, PictCalif.FontSize, "Bimensual 2"
 PCol = PCol + 1
 cPrint.printTextoAngulo PictCalif, PCol, PosLinea, 90, 5, PictCalif.FontSize, "Examen 1Q"
 PCol = PCol + 1
 cPrint.printTextoAngulo PictCalif, PCol, PosLinea, 90, 5, PictCalif.FontSize, "Prom. 1Q"
 PCol = PCol + 1
 cPrint.printTextoAngulo PictCalif, PCol, PosLinea, 90, 5, PictCalif.FontSize, "Bimensual 1"
 PCol = PCol + 1
 cPrint.printTextoAngulo PictCalif, PCol, PosLinea, 90, 5, PictCalif.FontSize, "Bimensual 2"
 PCol = PCol + 1
 cPrint.printTextoAngulo PictCalif, PCol, PosLinea, 90, 5, PictCalif.FontSize, "Examen 2Q"
 PCol = PCol + 1
 cPrint.printTextoAngulo PictCalif, PCol, PosLinea, 90, 5, PictCalif.FontSize, "Prom. 2Q"
 PCol = PCol + 1
 cPrint.printTextoAngulo PictCalif, PCol, PosLinea, 90, 5, PictCalif.FontSize, "Promedio"
 PCol = PCol + 1
 cPrint.printTextoAngulo PictCalif, PCol, PosLinea, 90, 5, PictCalif.FontSize, "SUPLETORIO"
 PCol = PCol + 1
 cPrint.printTextoAngulo PictCalif, PCol, PosLinea, 90, 5, PictCalif.FontSize, "Promedio"
 PCol = PCol + 0.2
 cPrint.printTextoAngulo PictCalif, PCol, PosLinea, 90, 5, PictCalif.FontSize, "Final"
 PCol = PCol + 0.6
 PosColumna = PCol
 PictCalif.FontBold = False
 
 PictCalif.FontSize = 9
 PosLinea = PosLinea + 0.2
 PCol = 10.25
 Contador = 0
 For I = 0 To ContNotas - 1
     VectNota(I) = 0
     SumaHoriz(I) = 0
 Next I
 I = 0
 NumMeses = 0
 Cuota_No = 0
 Total = 0
 JR = 1
 PictCalif.FontBold = False
 With AdoNotas.Recordset
  If .RecordCount > 0 Then
     .MoveFirst
      Do While Not .EOF
         IR = PCol
         Codigo = .Fields("Codigo")
         NombreCliente = .Fields("Cliente")
         Contador = Contador + 1
         PictCalif.FontName = TipoArialNarrow
         PictPrint_Texto PictCalif, 0.8, PosLinea, Format(Contador, "00") & ".-"
         PictPrint_Texto PictCalif, 1.3, PosLinea, .Fields("Cliente")
         PictCalif.FontName = TipoArial
        'Imprimimos las notas
         If Print_Nota(1) Then PictPrint_Nota_Materia PictCalif, IR, PosLinea, .Fields("PQBim1")
         IR = IR + JR
         If Print_Nota(2) Then PictPrint_Nota_Materia PictCalif, IR, PosLinea, .Fields("PQBim2")
         IR = IR + JR
         If Print_Nota(4) Then PictPrint_Nota_Materia PictCalif, IR, PosLinea, .Fields("ExamenPQ")
         IR = IR + JR
         If Print_Nota(6) Then PictPrint_Nota_Materia PictCalif, IR, PosLinea, .Fields("PromPQ")
         IR = IR + JR
         If Print_Nota(7) Then PictPrint_Nota_Materia PictCalif, IR, PosLinea, .Fields("SQBim1")
         IR = IR + JR
         If Print_Nota(8) Then PictPrint_Nota_Materia PictCalif, IR, PosLinea, .Fields("SQBim2")
         IR = IR + JR
         If Print_Nota(10) Then PictPrint_Nota_Materia PictCalif, IR, PosLinea, .Fields("ExamenSQ")
         IR = IR + JR
         If Print_Nota(12) Then PictPrint_Nota_Materia PictCalif, IR, PosLinea, .Fields("PromSQ")
         IR = IR + JR
         Sumatoria = (.Fields("PromPQ") + .Fields("PromSQ")) / 2
         If Print_Nota(13) Then PictPrint_Nota_Materia PictCalif, IR, PosLinea, Sumatoria
         IR = IR + JR
         If Print_Nota(14) Then PictPrint_Nota_Materia PictCalif, IR, PosLinea, .Fields("Supletorio")
         IR = IR + JR
         If Print_Nota(15) Then PictPrint_Nota_Materia PictCalif, IR, PosLinea, .Fields("PromFinal")
         IR = IR + JR
         PosLinea = PosLinea + 0.45
         PictCalif.Line (0.7, PosLinea)-(PosColumna, PosLinea)
         PosLinea = PosLinea + 0.05
         Total = Total + .Fields("PromFinal")
        .MoveNext
      Loop
     'Fin de la impresion
      PosLinea = PosLinea + 0.45
     'recuadro externo
      PictCalif.Line (0.7, 3.3)-(PosColumna, PosLinea), , B
      PictCalif.Line (0.65, 3.35)-(PosColumna - 0.05, PosLinea - 0.05), , B
      PictCalif.Line (9.9, 3.95)-(PosColumna - 3.05, 3.95)
      PictCalif.Line (0.7, 5.9)-(PosColumna - 0.05, 5.9)
      PictCalif.Line (0.7, 4.6)-(9.9, 4.6)
      PCol = 9.9
      For I = 0 To 10
        Select Case I
          Case 0, 4, 8, 9, 10: PictCalif.Line (PCol, 3.3)-(PCol, PosLinea)
          Case Else: PictCalif.Line (PCol, 3.95)-(PCol, PosLinea)
        End Select
        PCol = PCol + 1
      Next I
      PictCalif.FontBold = True
      PCol = 10
      
'''      PictPrint_Texto PictCalif, PCol - 3.5, PosLinea, "TOTAL PROMEDIOS:"
'''      For I = 0 To ContNotas - 1
'''        If SumaHoriz(I) > 0 Then
'''           SumaHoriz(I) = SumaHoriz(I) / Contador
'''           PictPrint_Texto PictCalif, PCol, PosLinea, Format(SumaHoriz(I), "00.00")
'''        End If
'''        PCol = PCol + 0.8
'''      Next I
'''      PCol = PCol - 4
'''      If NumMeses <= 0 Then NumMeses = 1
'''      PictPrint_Texto PictCalif, PCol, PosLinea, Format(Cuota_No / NumMeses, "00.00")
'''      PCol = PCol + 4
'''      PCol = PCol - 0.55
'''      If SumaHoriz(ContNotas) > 0 Then
'''         SumaHoriz(ContNotas) = SumaHoriz(ContNotas) / Contador
'''         PictPrint_Texto PictCalif, PCol, PosLinea, Format(SumaHoriz(ContNotas), "00.00")
'''      End If
  End If
 End With
End Sub

Public Sub Listar_Examenes_Grado(Solo_Supletorio As Boolean)
Dim VCodMat(10) As TipoMaterias
Dim Cursos As String
Dim PosLinea1 As Single
 'Notas por Materia de Profesor
  DCMaterias.Visible = False
  Codigo = SinEspaciosIzq(DLCurso)
  sSQL = "SELECT * " _
       & "FROM Catalogo_Materias " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "ORDER BY CodMat "
  SelectAdodc AdoMaterias, sSQL
  
  sSQL = "SELECT CodMat " _
       & "FROM Trans_Notas_Grado " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND Mid$(CodE,1," & Len(Codigo) & ") = '" & Codigo & "' " _
       & "GROUP BY CodMat " _
       & "ORDER BY CodMat "
  SelectAdodc AdoNotas, sSQL
  With AdoNotas.Recordset
    If .RecordCount > 0 Then
        VCodMat(0).CantMat = .RecordCount
        I = 0
        Do While Not .EOF
           VCodMat(I).CodigoMat = .Fields("CodMat")
           I = I + 1
          .MoveNext
        Loop
    End If
  End With
  For I = 0 To VCodMat(0).CantMat - 1
      With AdoMaterias.Recordset
       If .RecordCount > 0 Then
          .MoveFirst
          .Find ("CodMat = '" & VCodMat(I).CodigoMat & "' ")
           If Not .EOF Then
              VCodMat(I).Materias = .Fields("Materia")
           End If
       End If
      End With
  Next I
  sSQL = "SELECT C.Cliente,C.Sexo,TN.* " _
       & "FROM Trans_Notas_Grado As TN,Clientes As C,Clientes_Matriculas As CM " _
       & "WHERE TN.Item = '" & NumEmpresa & "' " _
       & "AND TN.Periodo = '" & Periodo_Contable & "' " _
       & "AND Mid$(TN.CodE,1," & Len(Codigo) & ") = '" & Codigo & "' "
  If Solo_Supletorio Then sSQL = sSQL & "AND TN.Examen <= 11 "
  sSQL = sSQL _
       & "AND TN.Codigo = C.Codigo " _
       & "AND CM.Codigo = C.Codigo " _
       & "AND TN.Item = CM.Item " _
       & "AND TN.Periodo = CM.Periodo " _
       & "ORDER BY C.Cliente,TN.CodMat,TN.Id_No "
  SelectAdodc AdoNotas, sSQL
  PictCalif.Cls
  PictCalif.Picture = LoadPicture(RutaSistema & "\FORMATOS\GENERAL\Pagina_A4.GIF")
  PictCalif.FontBold = True
  PCol = 10
'''  PictCalif.width = PCol + (ContNotas * 0.8) + 1.2
'''  PictCalif.Height = 12 + (AdoPromedio1.Recordset.RecordCount * 0.45)
  PictCalif.FontName = TipoTimes
  PosLinea = 0.3
  If LogoTipo <> "" Then PictCalif.PaintPicture LoadPicture(LogoTipo), 0.1, PosLinea, 5, 2.5
  PictCalif.FontSize = 20
  PictPrint_Texto PictCalif, 1, PosLinea, Institucion1, , PictCalif.width, True
  PosLinea = PosLinea + 0.7
  PictPrint_Texto PictCalif, 1, PosLinea, Institucion2, , PictCalif.width, True
  PosLinea = PosLinea + 0.7
  PictCalif.FontSize = 12
  PictPrint_Texto PictCalif, 1, PosLinea, UCase(NombreCiudad), , PictCalif.width, True
  PosLinea = PosLinea + 1.5
  Contador = 0
  I = 0
  NumMeses = 0
  Cuota_No = 0
  Total = 0
  JR = 1
  With AdoNotas.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      'Encabezado del curso
       Cursos = Leer_Datos_del_Curso(.Fields("CodE"), 1)
       PictCalif.FontSize = 10
       PosLinea1 = PosLinea
       PictPrint_Texto PictCalif, 15, PosLinea, "NOTAS DE ACTAS DE GRADO"
       PosLinea = PosLinea + 0.4
       PictPrint_Texto PictCalif, 15, PosLinea, "AÑO LECTIVO: " & Anio_Lectivo
       PosLinea = PosLinea1
       PosLinea = PictPrint_Texto_Multiple(PictCalif, 2, PosLinea, Cursos, 11.5)
       PictCalif.FontSize = 16
       PictPrint_Texto PictCalif, 4, PosLinea + 1, "M O M I N A"
       PCol = 18 - (VCodMat(0).CantMat * 1)
       PosColumna = PCol
       PictCalif.FontSize = 12
       PosLinea = PosLinea + 2.4
       For I = 0 To VCodMat(0).CantMat - 1
           Cadena1 = VCodMat(I).Materias
           Cadena = Trim(SinEspaciosIzq(Cadena1))
           Cadena1 = Trim(Mid$(Cadena1, Len(Cadena) + 1, Len(Cadena1)))
           If Cadena1 <> "" Then
              cPrint.printTextoAngulo PictCalif, PCol - 0.15, PosLinea, 90, 5, PictCalif.FontSize, Cadena
              cPrint.printTextoAngulo PictCalif, PCol + 0.15, PosLinea, 90, 5, PictCalif.FontSize, Cadena1
           Else
              cPrint.printTextoAngulo PictCalif, PCol, PosLinea, 90, 5, PictCalif.FontSize, Cadena
           End If
           PCol = PCol + 1
       Next I
       
       cPrint.printTextoAngulo PictCalif, PCol, PosLinea, 90, 5, PictCalif.FontSize, "SUMA"
       PCol = PCol + 1
       cPrint.printTextoAngulo PictCalif, PCol, PosLinea, 90, 5, PictCalif.FontSize, "Promedio"
       PCol = PCol + 1
       PosLinea = PosLinea + 0.15
       PictCalif.FontBold = False
       PictCalif.FontSize = 10
       Contador = Contador + 1
       PictCalif.FontName = TipoArialNarrow
       PictPrint_Texto PictCalif, 2, PosLinea, Format(Contador, "00") & ".-"
       PictPrint_Texto PictCalif, 2.7, PosLinea, .Fields("Cliente")
       NombreCliente = .Fields("Cliente")
       Total = 0
       Do While Not .EOF
          If NombreCliente <> .Fields("Cliente") Then
            'Imprimimos las notas
             IR = PosColumna
             Saldo = 0
             For I = 0 To VCodMat(0).CantMat - 1
                 PictPrint_Texto PictCalif, IR, PosLinea, Format(VCodMat(I).Valor, "00"), False
                 IR = IR + JR
                 Saldo = Saldo + VCodMat(I).Valor
             Next I
             PictPrint_Texto PictCalif, IR, PosLinea, Format(Saldo, "00"), False
             IR = IR + JR
             PictPrint_Texto PictCalif, IR - 0.2, PosLinea, Format(Saldo / VCodMat(0).CantMat, "00.00"), False
             IR = IR + JR
             PosLinea = PosLinea + 0.45
             PictCalif.Line (1.9, PosLinea)-(19.6, PosLinea)
             PosLinea = PosLinea + 0.05
             Codigo = .Fields("Codigo")
             NombreCliente = .Fields("Cliente")
             Contador = Contador + 1
             PictPrint_Texto PictCalif, 2, PosLinea, Format(Contador, "00") & ".-"
             PictPrint_Texto PictCalif, 2.7, PosLinea, .Fields("Cliente")
          End If
          For I = 0 To VCodMat(0).CantMat - 1
              If VCodMat(I).CodigoMat = .Fields("CodMat") Then VCodMat(I).Valor = .Fields("Examen")
          Next I
          Total = Total + .Fields("Examen")
         .MoveNext
       Loop
        'Imprimimos las notas
         IR = PosColumna
         Saldo = 0
         For I = 0 To VCodMat(0).CantMat - 1
             PictPrint_Texto PictCalif, IR, PosLinea, Format(VCodMat(I).Valor, "00"), False
             IR = IR + JR
             Saldo = Saldo + VCodMat(I).Valor
         Next I
         PictPrint_Texto PictCalif, IR, PosLinea, Format(Saldo, "00"), False
         IR = IR + JR
         PictPrint_Texto PictCalif, IR - 0.2, PosLinea, Format(Saldo / VCodMat(0).CantMat, "00.00"), False
         IR = IR + JR
         PosLinea = PosLinea + 0.45
         PictCalif.Line (1.9, PosLinea)-(19.6, PosLinea)
         PosLinea = PosLinea + 0.05
      'Fin de la impresion
      'Recuadro externo
       PictCalif.Line (1.9, 4)-(19.6, 6.1), , B
       PictCalif.Line (1.9, 6.1)-(19.6, PosLinea), , B
       PCol = PosColumna - 0.4
       For I = 0 To VCodMat(0).CantMat + 1
           PictCalif.Line (PCol, 4)-(PCol, PosLinea)
           PCol = PCol + JR
       Next I
   End If
  End With
End Sub

Public Sub Listar_Mejor_Egresado()
Dim VCodMat(10) As TipoMaterias
Dim Cursos As String
Dim PosLinea1 As Single
 'Notas por Materia de Profesor
  DCMaterias.Visible = False
  Codigo = SinEspaciosIzq(DLCurso)
  VCodMat(0).CantMat = 4
  VCodMat(0).Materias = "Nota de 1 a 5"
  VCodMat(1).Materias = "Trabajo"
  VCodMat(2).Materias = "Inv Prom Sexto"
  VCodMat(3).Materias = "Nota de Grado"
  
  sSQL = "UPDATE Trans_Actas " _
       & "SET Notas = ROUND(Notas,2), " _
       & "Trabajo = ROUND(Trabajo,2)," _
       & "Investigacion = ROUND(Investigacion,2)," _
       & "Evaluacion = ROUND(Evaluacion,2) " _
       & "WHERE Periodo = '" & Periodo_Contable & "' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Mid$(CodE,1," & Len(Codigo) & ") = '" & Codigo & "' "
  ConectarAdoExecute sSQL

  sSQL = "SELECT C.Cliente,C.Sexo,TN.* " _
       & "FROM Trans_Actas As TN,Clientes As C,Clientes_Matriculas As CM " _
       & "WHERE TN.Item = '" & NumEmpresa & "' " _
       & "AND TN.Periodo = '" & Periodo_Contable & "' " _
       & "AND Mid$(TN.CodE,1," & Len(Codigo) & ") = '" & Codigo & "' " _
       & "AND TN.Codigo = C.Codigo " _
       & "AND CM.Codigo = C.Codigo " _
       & "AND TN.Item = CM.Item " _
       & "AND TN.Periodo = CM.Periodo " _
       & "ORDER BY C.Cliente "
  SelectAdodc AdoNotas, sSQL
  PictCalif.Cls
  PictCalif.Picture = LoadPicture(RutaSistema & "\FORMATOS\GENERAL\Pagina_A4.GIF")
  PictCalif.FontBold = True
  PCol = 10
'''  PictCalif.width = PCol + (ContNotas * 0.8) + 1.2
'''  PictCalif.Height = 12 + (AdoPromedio1.Recordset.RecordCount * 0.45)
  PictCalif.FontName = TipoTimes
  PosLinea = 0.3
  If LogoTipo <> "" Then PictCalif.PaintPicture LoadPicture(LogoTipo), 1, PosLinea, 3, 1.5
  PictCalif.FontSize = 22
  PictPrint_Texto PictCalif, 1, PosLinea, Institucion1, , PictCalif.width, True
  PosLinea = PosLinea + 0.7
  PictPrint_Texto PictCalif, 1, PosLinea, Institucion2, , PictCalif.width, True
  PosLinea = PosLinea + 0.7
  PictCalif.FontSize = 12
  PictPrint_Texto PictCalif, 1, PosLinea, UCase(NombreCiudad), , PictCalif.width, True
  PosLinea = PosLinea + 0.7
  Contador = 0
  I = 0
  NumMeses = 0
  Cuota_No = 0
  Total = 0
  JR = 1.1
  With AdoNotas.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      'Encabezado del curso
       Contra_Cta = ""
       Cursos = Leer_Datos_del_Curso(.Fields("CodE"), 1)
       PictCalif.FontSize = 10
       PosLinea1 = PosLinea
       PictPrint_Texto PictCalif, 14.8, PosLinea, "NOTAS DE ACTAS DE GRADO"
       PosLinea = PosLinea + 0.4
       PictPrint_Texto PictCalif, 14.8, PosLinea, "AÑO LECTIVO: " & Anio_Lectivo
       PosLinea = PosLinea1
       PosLinea = PictPrint_Texto_Multiple(PictCalif, 2, PosLinea, Cursos, 11.5)
       PictCalif.FontSize = 16
       
       PictPrint_Texto PictCalif, 4, PosLinea + 1, "M O M I N A"
       PCol = 17.7 - (VCodMat(0).CantMat * 1.1)
       PosColumna = PCol
       PictCalif.FontSize = 12
       PosLinea = PosLinea + 2.1
       For I = 0 To VCodMat(0).CantMat - 1
           Cadena1 = VCodMat(I).Materias
           Cadena = Trim(SinEspaciosIzq(Cadena1))
           Cadena1 = Trim(Mid$(Cadena1, Len(Cadena) + 1, Len(Cadena1)))
           If Cadena1 <> "" Then
              cPrint.printTextoAngulo PictCalif, PCol - 0.1, PosLinea, 90, 5, PictCalif.FontSize, Cadena
              cPrint.printTextoAngulo PictCalif, PCol + 0.25, PosLinea, 90, 5, PictCalif.FontSize, Cadena1
           Else
              cPrint.printTextoAngulo PictCalif, PCol + 0.1, PosLinea, 90, 5, PictCalif.FontSize, Cadena
           End If
           PCol = PCol + 1.1
       Next I
       cPrint.printTextoAngulo PictCalif, PCol + 0.1, PosLinea, 90, 5, PictCalif.FontSize, "SUMA"
       PCol = PCol + 1
       cPrint.printTextoAngulo PictCalif, PCol + 0.1, PosLinea, 90, 5, PictCalif.FontSize, "Promedio"
       PCol = PCol + 1
       PosLinea = PosLinea + 0.15
       PictCalif.FontBold = False
       PictCalif.FontSize = 10
       PictCalif.FontName = TipoArialNarrow
       Total = 0
       Do While Not .EOF
          Contador = Contador + 1
          PictPrint_Texto PictCalif, 2, PosLinea, Format(Contador, "00") & ".-"
          PictPrint_Texto PictCalif, 2.7, PosLinea, .Fields("Cliente")
          VCodMat(0).Valor = .Fields("Notas")
          VCodMat(1).Valor = .Fields("Trabajo")
          VCodMat(2).Valor = .Fields("Investigacion")
          VCodMat(3).Valor = .Fields("Evaluacion")
          'MsgBox .Fields("Notas")
          IR = PosColumna
          Saldo = 0
          For I = 0 To VCodMat(0).CantMat - 1
              PictPrint_Texto PictCalif, IR, PosLinea, Format(VCodMat(I).Valor, "00.00"), False
              IR = IR + JR
              Saldo = Saldo + VCodMat(I).Valor
          Next I
          PictPrint_Texto PictCalif, IR, PosLinea, Format(Saldo, "#,##0.00"), False
          IR = IR + JR
          PictPrint_Texto PictCalif, IR - 0.2, PosLinea, Format(Saldo / VCodMat(0).CantMat, "00.00"), False
          IR = IR + JR
          PosLinea = PosLinea + 0.45
          PictCalif.Line (1.9, PosLinea)-(19.6, PosLinea)
          PosLinea = PosLinea + 0.05
         .MoveNext
       Loop
      'Fin de la impresion
      'Recuadro externo
       PictCalif.Line (1.9, 3.2)-(19.6, 5), , B
       PictCalif.Line (1.9, 5)-(19.6, PosLinea), , B
       PCol = PosColumna - 0.3
       For I = 0 To VCodMat(0).CantMat + 1
           PictCalif.Line (PCol, 3.2)-(PCol, PosLinea)
           PCol = PCol + JR
       Next I
   End If
  End With
  PosLinea = PosLinea + 0.05
  PictPrint_Texto PictCalif, 15, PosLinea, FechaStrgCiudad(MBFecha)
  PictPrint_Texto PictCalif, 3.7, 26.8, String(Len(Rector) - 0.5, "_")
  PictPrint_Texto PictCalif, 11.2, 26.8, String(Len(Secretario2) - 0.5, "_")
  PictPrint_Texto PictCalif, 4, 27.2, Rector
  PictPrint_Texto PictCalif, 11.5, 27.2, Secretario2
  PictPrint_Texto PictCalif, 4, 27.6, "RECTOR(A)"
  PictPrint_Texto PictCalif, 11.5, 27.6, "SECRETARIO(A)"
End Sub

Public Sub Listar_Nomina_Oficial()
Dim VCodMat(10) As TipoMaterias
Dim Cursos As String
Dim SumaNotas As Currency
Dim PosLinea1 As Single
 'Notas por Materia de Profesor
  DCMaterias.Visible = False
  Codigo = SinEspaciosIzq(DLCurso)
  VCodMat(0).CantMat = 3
  VCodMat(0).Materias = "No. Acta Grado"
  VCodMat(1).Materias = "Cedula Identidad"
  VCodMat(2).Materias = "Sexo"
  TextoValido TxtDesde, True, , 0
  TextoValido TxtHasta, True, , 0
  sSQL = "SELECT C.Cliente,C.Sexo,CM.CI,TN.* " _
       & "FROM Trans_Actas As TN,Clientes As C,Clientes_Matriculas As CM " _
       & "WHERE TN.Id_No BETWEEN " & Val(TxtDesde) & " and " & Val(TxtHasta) & " " _
       & "AND TN.Item = '" & NumEmpresa & "' " _
       & "AND TN.Periodo = '" & Periodo_Contable & "' " _
       & "AND Mid$(TN.CodE,1," & Len(Codigo) & ") = '" & Codigo & "' " _
       & "AND TN.Id_No <> 0 " _
       & "AND TN.Codigo = C.Codigo " _
       & "AND CM.Codigo = C.Codigo " _
       & "AND TN.Item = CM.Item " _
       & "AND TN.Periodo = CM.Periodo " _
       & "ORDER BY C.Cliente "
  SelectAdodc AdoNotas, sSQL
  PictCalif.Cls
  PictCalif.Picture = LoadPicture(RutaSistema & "\FORMATOS\GENERAL\Pagina_A4.GIF")
  PictCalif.FontBold = True
  PCol = 10
'''  PictCalif.width = PCol + (ContNotas * 0.8) + 1.2
'''  PictCalif.Height = 12 + (AdoPromedio1.Recordset.RecordCount * 0.45)
  PictCalif.FontName = TipoTimes
  PosLinea = 0.3
  If LogoTipo <> "" Then PictCalif.PaintPicture LoadPicture(LogoTipo), 1, PosLinea, 3, 1.5
  PictCalif.FontSize = 22
  PictPrint_Texto PictCalif, 1, PosLinea, Institucion1, , PictCalif.width, True
  PosLinea = PosLinea + 0.7
  PictPrint_Texto PictCalif, 1, PosLinea, Institucion2, , PictCalif.width, True
  PosLinea = PosLinea + 0.7
  PictCalif.FontSize = 12
  PictPrint_Texto PictCalif, 1, PosLinea, UCase(NombreCiudad), , PictCalif.width, True
  PosLinea = PosLinea + 0.5
  Contador = 0
  I = 0
  NumMeses = 0
  Cuota_No = 0
  Total = 0
  JR = 1.1
  With AdoNotas.Recordset
   If .RecordCount > 0 Then
       'MsgBox .RecordCount
      .MoveFirst
      'Encabezado del curso
       Cursos = Leer_Datos_del_Curso(.Fields("CodE"), 1)
       PictCalif.FontSize = 10
       PosLinea1 = PosLinea
       PictPrint_Texto PictCalif, 14, PosLinea, "NOMINA OFICIAL DE GRADUADOS"
       PosLinea = PosLinea + 0.4
       PictPrint_Texto PictCalif, 15, PosLinea, "AÑO LECTIVO: " & Anio_Lectivo
       PosLinea = PosLinea1
       PosLinea = PictPrint_Texto_Multiple(PictCalif, 2, PosLinea, Cursos, 11.5)
       PictCalif.FontSize = 16
       PictPrint_Texto PictCalif, 2.5, PosLinea + 0.5, "MOMINA  DE GRADUADOS"
       PosColumna = PCol
       PictCalif.FontSize = 9
       PosLinea = PosLinea + 1
       PCol = 13.5
       Cadena1 = VCodMat(0).Materias
       Cadena = Trim(SinEspaciosIzq(Cadena1))
       Cadena1 = Trim(Mid$(Cadena1, Len(Cadena) + 1, Len(Cadena1)))
       PictPrint_Texto PictCalif, PCol, PosLinea, Cadena
       PictPrint_Texto PictCalif, PCol, PosLinea + 0.3, Cadena1
       PCol = 15
       Cadena1 = VCodMat(1).Materias
       Cadena = Trim(SinEspaciosIzq(Cadena1))
       Cadena1 = Trim(Mid$(Cadena1, Len(Cadena) + 1, Len(Cadena1)))
       PictPrint_Texto PictCalif, PCol, PosLinea, Cadena
       PictPrint_Texto PictCalif, PCol, PosLinea + 0.3, Cadena1
       PCol = 17.1
       Cadena1 = VCodMat(2).Materias
       Cadena = Trim(SinEspaciosIzq(Cadena1))
       Cadena1 = Trim(Mid$(Cadena1, Len(Cadena) + 1, Len(Cadena1)))
       PictPrint_Texto PictCalif, PCol, PosLinea + 0.1, Cadena
       PCol = 18.5
       PosLinea = PosLinea + 0.9
       cPrint.printTextoAngulo PictCalif, PCol + 0.1, PosLinea - 0.05, 90, 5, PictCalif.FontSize, "Promedio"
       PCol = PCol + 1
       PosLinea = PosLinea + 0.15
       PictCalif.FontBold = False
       PictCalif.FontSize = 10
       PictCalif.FontName = TipoArialNarrow
       Total = 0
       Do While Not .EOF
          Contador = Contador + 1
          PictPrint_Texto PictCalif, 2, PosLinea, Format(Contador, "00") & ".-"
          PictPrint_Texto PictCalif, 2.7, PosLinea, .Fields("Cliente")
          SumaNotas = .Fields("Notas") + .Fields("Trabajo") + .Fields("Investigacion") + .Fields("Evaluacion")
          SumaNotas = Val(Format(SumaNotas / 4, "00.00"))
          SumaNotas = Val(Format(SumaNotas, "00"))
          PictPrint_Texto PictCalif, 13.5, PosLinea, Format(.Fields("Id_No"), "000")
          PictPrint_Texto PictCalif, 15, PosLinea, .Fields("CI")
          PictPrint_Texto PictCalif, 17.5, PosLinea, .Fields("Sexo")
          PictPrint_Texto PictCalif, 18.7, PosLinea, Format(SumaNotas, "00"), False
          PosLinea = PosLinea + 0.45
          PictCalif.Line (1.9, PosLinea)-(19.6, PosLinea)
          PosLinea = PosLinea + 0.05
         .MoveNext
       Loop
      'Fin de la impresion
      'Recuadro externo
       PictCalif.Line (1.9, 3)-(19.6, 4.6), , B
       PictCalif.Line (1.9, 4.6)-(19.6, PosLinea), , B
       
       PictCalif.Line (13.4, 3)-(13.4, PosLinea)
       PictCalif.Line (14.9, 3)-(14.9, PosLinea)
       PictCalif.Line (17, 3)-(17, PosLinea)
       PictCalif.Line (18.2, 3)-(18.2, PosLinea)
   End If
  End With
  PosLinea = PosLinea + 0.05
  PictPrint_Texto PictCalif, 15, PosLinea, FechaStrgCiudad(MBFecha)
  PictPrint_Texto PictCalif, 3.7, 26.3, String(Len(Rector) - 0.5, "_")
  PictPrint_Texto PictCalif, 11.2, 26.3, String(Len(Secretario2) - 0.5, "_")
  PictPrint_Texto PictCalif, 4, 26.8, Rector
  PictPrint_Texto PictCalif, 11.5, 26.8, Secretario2
  PictPrint_Texto PictCalif, 4, 27.2, "RECTOR(A)"
  PictPrint_Texto PictCalif, 11.5, 27.2, "SECRETARIO(A)"
End Sub

Public Sub Imprimr_Nomina_Oficial()
Dim VNombres(10) As String
Dim Cursos As String
Dim AnchoPict As Single
Dim AltoPict As Single
Dim NombFilePict As String
Dim SumaNotas As Currency
On Error GoTo Errorhandler
   Mensajes = "Seguro de Imprimir en:" & vbCrLf & Printer.DeviceName & "?"
   Titulo = "IMPRESION"
   Bandera = False
   Orientacion_Pagina = 2
   SetPrinters.Show 1
   If PonImpresoraDefecto(SetNombrePRN) Then
      RatonReloj
      InicioX = 0: InicioY = 0
      Escala_Centimetro Orientacion_Pagina, TipoTimes, 8
      Pagina = 1
     'Notas por Materia de Profesor
      DCMaterias.Visible = False
      Codigo = SinEspaciosIzq(DLCurso)
      TextoValido TxtDesde, True, , 0
      TextoValido TxtHasta, True, , 0
      sSQL = "SELECT C.Cliente,C.Sexo,CM.Especialidad,CM.Titulo,CM.Tipo_Titulo,CM.Nivel,CM.Ciclo,CM.CI,CM.Codigo_Titulo,TN.* " _
           & "FROM Trans_Actas As TN,Clientes As C,Clientes_Matriculas As CM " _
           & "WHERE TN.Id_No BETWEEN " & Val(TxtDesde) & " and " & Val(TxtHasta) & " " _
           & "AND TN.Item = '" & NumEmpresa & "' " _
           & "AND TN.Periodo = '" & Periodo_Contable & "' " _
           & "AND Mid$(TN.CodE,1," & Len(Codigo) & ") = '" & Codigo & "' " _
           & "AND TN.Id_No <> 0 " _
           & "AND TN.Codigo = C.Codigo " _
           & "AND CM.Codigo = C.Codigo " _
           & "AND TN.Item = CM.Item " _
           & "AND TN.Periodo = CM.Periodo " _
           & "ORDER BY C.Cliente "
      SelectAdodc AdoNotas, sSQL
      PCol = 10
      Printer.FontName = TipoArialNarrow
      Printer.FontSize = 10
      Contador = 0
      With AdoNotas.Recordset
       If .RecordCount > 0 Then
          .MoveFirst
          'Encabezado del curso
           Encabezado_Nomina_Oficial
           Do While Not .EOF
              SumaNotas = .Fields("Notas") + .Fields("Trabajo") + .Fields("Investigacion") + .Fields("Evaluacion")
              SumaNotas = Val(Format(SumaNotas / 4, "00.00"))
              SumaNotas = Val(Format(SumaNotas, "00"))
              Printer.Line (1.9, PosLinea)-(3.3, PosLinea + 0.6), Negro, B    ' no Acta
              Printer.Line (3.4, PosLinea)-(5.8, PosLinea + 0.6), Negro, B    ' CI
              Printer.Line (5.9, PosLinea)-(17.7, PosLinea + 0.6), Negro, B   ' Nombres
              Printer.Line (17.8, PosLinea)-(19.2, PosLinea + 0.6), Negro, B  ' Sexo
              Printer.Line (19.3, PosLinea)-(20.2, PosLinea + 0.6), Negro, B  ' Calificacion
              Printer.Line (20.2, PosLinea)-(23.7, PosLinea + 0.6), Negro, B  ' En letras
              Printer.Line (23.8, PosLinea)-(24.8, PosLinea + 0.6), Negro, B  ' Fecha dia
              Printer.Line (24.8, PosLinea)-(25.8, PosLinea + 0.6), Negro, B  ' Fecha mes
              Printer.Line (25.8, PosLinea)-(27, PosLinea + 0.6), Negro, B  ' Fecha año
              PosLinea = PosLinea + 0.1
              PrinterTexto 2.2, PosLinea, Format(.Fields("Id_No"), "000")
              PrinterTexto 3.5, PosLinea, .Fields("CI")
              NombreCliente = .Fields("Cliente")
              
              VNombres(0) = Trim(SinEspaciosIzq(NombreCliente))
              NombreCliente = Trim(Mid$(NombreCliente, Len(VNombres(0)) + 1, Len(NombreCliente)))
              VNombres(1) = SinEspaciosIzq(NombreCliente)
              NombreCliente = Trim(Mid$(NombreCliente, Len(VNombres(1)) + 1, Len(NombreCliente)))
              
              PrinterTexto 6, PosLinea, Trim(VNombres(0))
              PrinterTexto 9, PosLinea, Trim(VNombres(1))
              PrinterTexto 12, PosLinea, Trim(NombreCliente)
              If .Fields("Sexo") = "M" Then
                  PrinterTexto 18.3, PosLinea, "1"
              Else
                  PrinterTexto 18.3, PosLinea, "2"
              End If
              PrinterTexto 19.5, PosLinea, Format(SumaNotas, "00")
              PrinterTexto 20.5, PosLinea, Cambio_Letras(SumaNotas, True)
              NoDias = Day(MBFecha)
              NoMeses = Month(MBFecha)
              NoAnio = Year(MBFecha)
              PrinterTexto 24, PosLinea, Format(NoDias, "00")
              PrinterTexto 25, PosLinea, Format(NoMeses, "00")
              PrinterTexto 26, PosLinea, Format(NoAnio, "0000")
              PosLinea = PosLinea + 0.5
              If PosLinea >= LimiteAlto Then
                 Printer.NewPage
                 Printer.FontName = TipoArialNarrow
                 Encabezado_Nomina_Oficial
              End If
             .MoveNext
           Loop
       End If
      End With
   MensajeEncabData = ""
   Printer.EndDoc
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

Public Sub Imprimr_Nomina_Oficial_Excel()
Dim VNombres(10) As String
Dim PosXPict As Single
Dim PosYPict As Single
Dim Disciplina As Single
Dim SumaPromX As Single
Dim SumaPromY As Single
Dim CantMaterias As Byte
Dim CantAlumnos As Byte
Dim NombreMateria As String
Dim SiguientePagina As Boolean
 
Dim Aprobado As Boolean
Dim apexcel As Variant
Dim SumaNotas As Currency
  RatonReloj
  Contador = 0
  Progreso_Iniciar
  Set apexcel = CreateObject("excel.application")
 'hace que excel se vea o no
  apexcel.Visible = False
 'agrega un nuevo libro
  apexcel.workbooks.Add
  SiguientePagina = True
  Pagina = 1
  SumaPromX = 0
  SumaPromY = 0
  CantAlumnos = 0
  CantMaterias = 0
    'Notas por Materia de Profesor
     DCMaterias.Visible = False
     TextoValido TxtDesde, True, , 0
     TextoValido TxtHasta, True, , 0
     Codigo = SinEspaciosIzq(DLCurso)
     sSQL = "SELECT C.Cliente,C.Sexo,CM.Especialidad,CM.Titulo,CM.Tipo_Titulo,CM.Nivel," _
          & "CM.Ciclo,CM.CI,CM.Codigo_Titulo,TN.* " _
          & "FROM Trans_Actas As TN,Clientes As C,Clientes_Matriculas As CM " _
          & "WHERE TN.Id_No BETWEEN " & Val(TxtDesde) & " and " & Val(TxtHasta) & " " _
          & "AND TN.Item = '" & NumEmpresa & "' " _
          & "AND TN.Periodo = '" & Periodo_Contable & "' " _
          & "AND TN.Id_No <> 0 " _
          & "AND TN.Codigo = C.Codigo " _
          & "AND CM.Codigo = C.Codigo " _
          & "AND TN.Item = CM.Item " _
          & "AND TN.Periodo = CM.Periodo " _
          & "ORDER BY TN.Id_No,C.Cliente "
     SelectAdodc AdoNotas, sSQL
     PorteLetra = 10
     Contador = 0
     With AdoNotas.Recordset
      If .RecordCount > 0 Then
          .MoveFirst
          'Encabezado del curso
           Contador = Contador + 1
           apexcel.cells(Contador, 1).formula = "DET_ACTA"
           apexcel.cells(Contador, 2).formula = "DET_CEDULA"
           apexcel.cells(Contador, 3).formula = "DET_NOMBRE"
           apexcel.cells(Contador, 4).formula = "DET_SEXO"
           apexcel.cells(Contador, 5).formula = "DET_COLEGI"
           apexcel.cells(Contador, 6).formula = "DET_TITULO"
           apexcel.cells(Contador, 7).formula = "DET_TIPTIT"
           apexcel.cells(Contador, 8).formula = "DET_DESTIT"
           apexcel.cells(Contador, 9).formula = "DET_PARALE"
           apexcel.cells(Contador, 10).formula = "DET_CALIFI"
           apexcel.cells(Contador, 11).formula = "DET_FECGRA"
           apexcel.cells(Contador, 12).formula = "DET_NUMREF"
           apexcel.cells(Contador, 13).formula = "DET_PAGINA"
           apexcel.cells(Contador, 14).formula = "DET_FECREF"
           Progreso_Barra.Valor_Maximo = Progreso_Barra.Valor_Maximo + .RecordCount
           Do While Not .EOF
              Progreso_Esperar
              Contador = Contador + 1
              SumaNotas = .Fields("Notas") + .Fields("Trabajo") + .Fields("Investigacion") + .Fields("Evaluacion")
              SumaNotas = Val(Format(SumaNotas / 4, "00.00"))
              SumaNotas = Val(Format(SumaNotas, "00"))
              apexcel.cells(Contador, 1).formula = Format(.Fields("Id_No"), "#.00000")
              apexcel.cells(Contador, 2).formula = "'" & .Fields("CI")
              apexcel.cells(Contador, 3).formula = UCase(.Fields("Cliente"))
              apexcel.cells(Contador, 4).formula = .Fields("Sexo")
              apexcel.cells(Contador, 5).formula = Format(Val(Codigo_Ministerio), "#.00000")
              apexcel.cells(Contador, 6).formula = "'" & Mid$(.Fields("Codigo_Titulo"), 1, 2)
              apexcel.cells(Contador, 7).formula = "'" & Mid$(.Fields("Codigo_Titulo"), 3, 2)
              apexcel.cells(Contador, 8).formula = "'" & Mid$(.Fields("Codigo_Titulo"), 5, 2)
              apexcel.cells(Contador, 9).formula = ""
              apexcel.cells(Contador, 10).formula = Format(SumaNotas, "#.00000")
              apexcel.cells(Contador, 11).formula = MBFecha
              apexcel.cells(Contador, 12).formula = "0.00000"
              apexcel.cells(Contador, 13).formula = "0.00000"
              apexcel.cells(Contador, 14).formula = "/  /"
             .MoveNext
           Loop
       End If
      End With
  
  apexcel.Visible = True
  Set apexcel = Nothing
  RatonNormal
  Progreso_Final
End Sub

Public Sub Listar_Mejor_Promedio()
Dim VCodMat(20) As TipoMaterias
Dim Cursos As String
   Opcion = 2
   Dec_Campos = ""
   Mensajes = "Listar Mejor Promedio" & vbCrLf _
            & "por Cursos"
   Titulo = "Pregunta de Confirmación"
   If BoxMensaje = vbYes Then Opcion = 1
          
 'Notas por Materia de Profesor
  Codigo = SinEspaciosIzq(DLCurso)
  DCMaterias.Visible = False
  sSQL = "SELECT * " _
       & "FROM Catalogo_Materias " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "ORDER BY CodMat "
  SelectAdodc AdoMaterias, sSQL
  
  sSQL = "SELECT CodMat " _
       & "FROM Trans_Notas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND Mid$(CodE,1," & Len(Codigo) & ") = '" & Codigo & "' " _
       & "GROUP BY CodMat " _
       & "ORDER BY CodMat "
  SelectAdodc AdoNotas, sSQL
  With AdoNotas.Recordset
    If .RecordCount > 0 Then
        VCodMat(0).CantMat = .RecordCount
        I = 0
        Do While Not .EOF
           VCodMat(I).CodigoMat = .Fields("CodMat")
           I = I + 1
          .MoveNext
        Loop
    End If
  End With
  For I = 0 To VCodMat(0).CantMat - 1
      With AdoMaterias.Recordset
       If .RecordCount > 0 Then
          .MoveFirst
          .Find ("CodMat = '" & VCodMat(I).CodigoMat & "' ")
           If Not .EOF Then
              VCodMat(I).Materias = .Fields("Materia")
           End If
       End If
      End With
  Next I
    MensajeEncabData = "CUADRO DE HONOR / HONOR ROLL"
    sSQL = "SELECT C.Cliente As Alumno,"
    If OpcPeriodo("PF", LstPeriodos) Then
       Cadena = "Periodo Final de Quimestres"
       SQLMsg1 = "P R O M E D I O     F I N A L"
    End If
    If OpcPeriodo("PQBim1", LstPeriodos) Or OpcPeriodo("SQBim1", LstPeriodos) Or OpcPeriodo("TQBim1", LstPeriodos) Then
       SQLMsg1 = "PRIMER PARCIAL / FIRST PARCIAL"
       sSQL = sSQL & "(SUM(TN." & SQLBim1 & ")/COUNT(C.Cliente)) As Tot_" & SQLBim1 & ","
       Dec_Campos = "Tot_" & SQLBim1
    End If
    If OpcPeriodo("PQBim2", LstPeriodos) Or OpcPeriodo("SQBim2", LstPeriodos) Or OpcPeriodo("TQBim2", LstPeriodos) Then
       SQLMsg1 = "SEGUNDO PARCIAL / SECOND PARCIAL"
       sSQL = sSQL & "(SUM(TN." & SQLBim2 & ")/COUNT(C.Cliente)) As Tot_" & SQLBim2 & ","
       Dec_Campos = "Tot_" & SQLBim2
    End If
    If OpcPeriodo("PQBim3", LstPeriodos) Or OpcPeriodo("SQBim3", LstPeriodos) Or OpcPeriodo("TQBim3", LstPeriodos) Then
       SQLMsg1 = "TERCER PARCIAL / SECOND PARCIAL"
       sSQL = sSQL & "(SUM(TN." & SQLBim3 & ")/COUNT(C.Cliente)) As Tot_" & SQLBim3 & ","
       Dec_Campos = "Tot_" & SQLBim3
    End If
    If OpcPeriodo("PQ", LstPeriodos) Or OpcPeriodo("SQ", LstPeriodos) Or OpcPeriodo("TQ", LstPeriodos) Then
       SQLMsg1 = "PROMEDIO QUIMESTRE / PROMED QUIMESTER"
       sSQL = sSQL & "(SUM(TN." & SQLBim1 & ")/COUNT(C.Cliente)) As Tot_" & SQLBim1 & ","
       sSQL = sSQL & "(SUM(TN." & SQLBim2 & ")/COUNT(C.Cliente)) As Tot_" & SQLBim2 & ","
       sSQL = sSQL & "(SUM(TN." & SQLBim3 & ")/COUNT(C.Cliente)) As Tot_" & SQLBim3 & ","
       sSQL = sSQL & "(SUM(TN." & SQLPromQ & ")/COUNT(C.Cliente)) As Tot_" & SQLPromQ & ","
       Dec_Campos = "Tot_" & SQLPromQ
    End If
    
    sSQL = sSQL & "C.Grupo,C.Direccion As Curso " _
         & "FROM Trans_Notas As TN,Catalogo_Materias As TM,Catalogo_Cursos As CC,Clientes As C " _
         & "WHERE TN.Item = '" & NumEmpresa & "' " _
         & "AND TN.Periodo = '" & Periodo_Contable & "' " _
         & "AND Mid$(TN.CodE,1,1) = '" & Mid$(Codigo, 1, 1) & "' " _
         & "AND TN.Orden = 0 " _
         & "AND TM.SDiv = " & Val(adFalse) & " " _
         & "AND TM.I <> " & Val(adFalse) & " " _
         & "AND TN.CodMat NOT IN ('999','998') " _
         & "AND TN.Codigo = C.Codigo " _
         & "AND TN.CodE = CC.Curso " _
         & "AND TN.CodMat = TM.CodMat " _
         & "AND TN.Item = CC.Item " _
         & "AND TN.Item = TM.Item " _
         & "AND TN.Periodo = CC.Periodo " _
         & "AND TN.Periodo = TM.Periodo " _
         & "GROUP BY C.Grupo,C.Direccion,C.Cliente "
    If OpcPeriodo("PQBim1", LstPeriodos) Or OpcPeriodo("SQBim1", LstPeriodos) Or OpcPeriodo("TQBim1", LstPeriodos) Then
       sSQL = sSQL & "HAVING (SUM(TN." & SQLBim1 & ")/COUNT(C.Cliente)) >= " & Mejor_Promedio & " "
    End If
    If OpcPeriodo("PQBim2", LstPeriodos) Or OpcPeriodo("SQBim2", LstPeriodos) Or OpcPeriodo("TQBim2", LstPeriodos) Then
       sSQL = sSQL & "HAVING (SUM(TN." & SQLBim2 & ")/COUNT(C.Cliente)) >= " & Mejor_Promedio & " "
    End If
    If OpcPeriodo("PQBim3", LstPeriodos) Or OpcPeriodo("SQBim3", LstPeriodos) Or OpcPeriodo("TQBim3", LstPeriodos) Then
       sSQL = sSQL & "HAVING (SUM(TN." & SQLBim3 & ")/COUNT(C.Cliente)) >= " & Mejor_Promedio & " "
    End If
    If OpcPeriodo("PQ", LstPeriodos) Or OpcPeriodo("SQ", LstPeriodos) Or OpcPeriodo("TQ", LstPeriodos) Then
       sSQL = sSQL & "HAVING (((SUM(TN." & SQLBim1 & ") + SUM(TN." & SQLBim2 & ") + SUM(TN." & SQLBim3 & "))/3" _
            & ")/COUNT(C.Cliente)) >= " & Mejor_Promedio & " "
    End If
  If Opcion = 1 Then
     sSQL = sSQL & "ORDER BY C.Grupo,(SUM(TN." & SQLPromQ & ")/COUNT(C.Cliente)) DESC,C.Cliente "
  Else
     sSQL = sSQL & "ORDER BY " & Dec_Campos & " DESC,C.Grupo,C.Cliente "
  End If
  '(SUM(TN." & SQLPromQ & ")/COUNT(C.Cliente))
 'MsgBox sSQL
  Dec_Campos = Dec_Campos & " 3|"
  SelectDataGrid DGResumenNotas, AdoResumenNotas, sSQL, Dec_Campos
  DGResumenNotas.Caption = MensajeEncabData
End Sub

Public Sub Listar_Notas_Del_Curso()
Dim PColMax As Single
Dim Nota1, Nota2, Nota3, Nota4, Nota5, Nota6 As Byte
Dim Not1, Not2, Not3, Not4, Not5, Not6 As Single
 RatonReloj
 Listar_Alumnos_Notas_Del SinEspaciosIzq(DLCurso), True
 NombreBanco = UCase(LstPeriodos.Text)
 PictCalif.Cls
 PictCalif.FontBold = True
 PCol = 16
 PictCalif.width = PCol + ((ContNotas - 4) * 2.46)
 PictCalif.Height = 12 + (AdoPromedio1.Recordset.RecordCount * 0.45)
 If PictCalif.Height < 30 Then PictCalif.Height = 30
 PictCalif.FontName = TipoTimes
 PosLinea = 1
 If LogoTipo <> "" Then PictCalif.PaintPicture LoadPicture(LogoTipo), 0.5, PosLinea, 5, 2.5
 PictCalif.FontSize = 22
 PictPrint_Texto PictCalif, 1, PosLinea, UCase(Institucion1 & " " & Institucion2), , PictCalif.width, True
 PosLinea = PosLinea + 0.9
 PictCalif.FontSize = 12
 PictCalif.FontBold = False
 PictPrint_Texto PictCalif, 1, PosLinea, "AÑO LECTIVO: " & Anio_Lectivo, , PictCalif.width, True
 PosLinea = PosLinea + 0.6
 PictCalif.FontSize = 11
 PictCalif.FontBold = True
 PictPrint_Texto PictCalif, 1, PosLinea, NombreBanco, , PictCalif.width, True
 PictCalif.FontBold = False
 PosLinea = PosLinea + 0.6
 PictCalif.FontSize = 12
 If OpcPeriodo("PF", LstPeriodos) Then
    PictPrint_Texto PictCalif, 1, PosLinea, CuentaBanco, , PictCalif.width, True
 Else
    PictPrint_Texto PictCalif, 1, PosLinea, "CUADRO DE JUNTAS PROMEDIALES DE APROVECHAMIENTO, CORRESPONDIENTE A LOS ALUMNOS DEL " & CuentaBanco, , PictCalif.width, True
 End If
 PosLinea = PosLinea + 0.6
 PictCalif.FontName = TipoArialNarrow
 PictCalif.FontSize = 8
 PorteLetra = PictCalif.FontSize
 PCol = 10
 PosLinea = 6.2
 For I = 0 To 50
     SumaHoriz(I) = 0
 Next I
 For I = 0 To ContNotas - 5
     Cadena1 = " ": Cadena2 = " "
     If NumeroDeEspacios(VectMate(I)) = 0 Then
        Cadena1 = VectMate(I)
     ElseIf NumeroDeEspacios(VectMate(I)) = 1 Then
        Cadena1 = SinEspaciosIzq(VectMate(I))
        Cadena2 = SinEspaciosDer(VectMate(I))
     Else
        Cadena1 = SinEspaciosIzq(VectMate(I))
        Cadena2 = Trim(Mid$(VectMate(I), Len(Cadena1) + 1, Len(VectMate(I))))
     End If
     PictPrint_Texto PictCalif, PCol + 0.05, PosLinea - 2.5, Cadena1
     PictPrint_Texto PictCalif, PCol + 0.05, PosLinea - 2.2, Cadena2
     cPrint.printTextoAngulo PictCalif, PCol, PosLinea, 90, 5, PorteLetra + 1, "Primer"
     PCol = PCol + 0.2
     cPrint.printTextoAngulo PictCalif, PCol, PosLinea, 90, 5, PorteLetra + 1, "Periodo"
     PCol = PCol + 0.4
     cPrint.printTextoAngulo PictCalif, PCol, PosLinea, 90, 5, PorteLetra + 1, "Segundo"
     PCol = PCol + 0.2
     cPrint.printTextoAngulo PictCalif, PCol, PosLinea, 90, 5, PorteLetra + 1, "Periodo"
     PCol = PCol + 0.4
     If Mid$(FormatoLibreta, 1, 9) = "QUIMESTRE" Then
        cPrint.printTextoAngulo PictCalif, PCol, PosLinea, 90, 5, PorteLetra + 1, "Tercer"
        PCol = PCol + 0.2
        cPrint.printTextoAngulo PictCalif, PCol, PosLinea, 90, 5, PorteLetra + 1, "Periodo"
        PCol = PCol + 0.4
     End If
     If OpcPeriodo("PF", LstPeriodos) Then
        cPrint.printTextoAngulo PictCalif, PCol, PosLinea, 90, 5, PorteLetra + 1, "Suma de los"
        PCol = PCol + 0.2
        cPrint.printTextoAngulo PictCalif, PCol, PosLinea, 90, 5, PorteLetra + 1, "Quimestres"
        PCol = PCol + 0.5
        cPrint.printTextoAngulo PictCalif, PCol, PosLinea, 90, 5, PorteLetra + 1, "Supletorio"
        PCol = PCol + 0.5
     Else
        cPrint.printTextoAngulo PictCalif, PCol, PosLinea, 90, 5, PorteLetra + 1, "Evaluacion"
        PCol = PCol + 0.2
        cPrint.printTextoAngulo PictCalif, PCol, PosLinea, 90, 5, PorteLetra + 1, "Acumulativa"
        PCol = PCol + 0.4
        cPrint.printTextoAngulo PictCalif, PCol, PosLinea, 90, 5, PorteLetra + 1, "Promedio"
        PCol = PCol + 0.2
        cPrint.printTextoAngulo PictCalif, PCol, PosLinea, 90, 5, PorteLetra + 1, "del Periodo"
        PCol = PCol + 0.4
     End If
 Next
 PictCalif.FontSize = 10
 If Not OpcPeriodo("PF", LstPeriodos) Then cPrint.printTextoAngulo PictCalif, PCol + 0.1, PosLinea, 90, 7, PorteLetra + 1, "PROMEDIO"
 PCol = PCol + 0.2
 If Not OpcPeriodo("PF", LstPeriodos) Then cPrint.printTextoAngulo PictCalif, PCol + 0.2, PosLinea, 90, 7, PorteLetra + 1, "FINAL"
 PCol = PCol + 0.6
 PColMax = PCol
 PictCalif.width = PColMax + 1
 PictCalif.FontSize = 9
 PosLinea = PosLinea + 0.2
 PCol = 10.25
 Contador = 0
 For I = 0 To ContNotas - 1
     VectNota(I) = 0
     SumaVerti(I) = 0
     SumaHoriz(I) = 0
 Next I
 Saldo = 0
 I = 0
 PCol = 10.1
 PictCalif.FontBold = False
'Empieza la impresion de las notas
 With AdoResumen.Recordset
  If .RecordCount > 0 Then
     .MoveFirst
      CodigoCli = .Fields("Codigo")
      NombreCliente = .Fields("Cliente")
      Contador = Contador + 1
      PictPrint_Texto PictCalif, 0.8, PosLinea, Format(Contador, "00") & ".-"
      PictPrint_Texto PictCalif, 1.3, PosLinea, .Fields("Cliente")
     'PosLinea = PosLinea + 0.4
      Do While Not .EOF
         If CodigoCli <> .Fields("Codigo") Then
            Total = 0
            J = 0
            Cadena = ""
            For I = 0 To ContNotas - 5
                If VectNota(I) > 0 Then
                   J = J + 1
                   Total = Total + VectNota(I)
                  'Cadena = Cadena & Contador & " >> " & I & " - " & VectNota(I) & " > " & VectMate(I) & vbCrLf
                End If
            Next I
            If J > 0 Then Total = Total / J
           'MsgBox Cadena & Redondear(Total)
            Saldo = Saldo + Total
            SumaHoriz(ContNotas) = SumaHoriz(ContNotas) + Total
            If Not OpcPeriodo("PF", LstPeriodos) Then PictPrint_Texto PictCalif, PCol, PosLinea, Format(Total, "00")
           'PCol = PCol + 0.8
            PosLinea = PosLinea + 0.45
            Contador = Contador + 1
            PictCalif.Line (0.7, PosLinea)-(PColMax, PosLinea)
            PosLinea = PosLinea + 0.05
            PictPrint_Texto PictCalif, 0.8, PosLinea, Format(Contador, "00") & ".-"
            PictPrint_Texto PictCalif, 1.3, PosLinea, .Fields("Cliente")
            PCol = 10.1
            For I = 0 To ContNotas - 1
                VectNota(I) = 0
            Next I
            I = 0
            CodigoCli = .Fields("Codigo")
            NombreCliente = .Fields("Cliente")
         End If
         Nota1 = 0
         Nota2 = 0
         Nota3 = 0
         Nota4 = 0
         Nota5 = 0
         Nota6 = 0
         If OpcPeriodo("PQBim1", LstPeriodos) Then
            Nota1 = .Fields("PQBim1")
            VectNota(I) = Nota1
         ElseIf OpcPeriodo("PQBim2", LstPeriodos) Then
            Nota1 = .Fields("PQBim1")
            Nota2 = .Fields("PQBim2")
            VectNota(I) = Nota2
         ElseIf OpcPeriodo("PQ", LstPeriodos) Then
            Nota1 = .Fields("PQBim1")
            Nota2 = .Fields("PQBim2")
            Nota3 = .Fields("ExamenPQ")
            Nota4 = .Fields("PromPQ")
            VectNota(I) = Nota4
         ElseIf OpcPeriodo("SQBim1", LstPeriodos) Then
            Nota1 = .Fields("SQBim1")
            VectNota(I) = Nota1
         ElseIf OpcPeriodo("SQBim2", LstPeriodos) Then
            Nota1 = .Fields("SQBim1")
            Nota2 = .Fields("SQBim2")
            VectNota(I) = Nota2
         ElseIf OpcPeriodo("SQ", LstPeriodos) Then
            Nota1 = .Fields("SQBim1")
            Nota2 = .Fields("SQBim2")
            Nota3 = .Fields("ExamenSQ")
            Nota4 = .Fields("PromSQ")
            VectNota(I) = Nota4
         ElseIf OpcPeriodo("TQBim1", LstPeriodos) Then
            Nota1 = .Fields("TQBim1")
            VectNota(I) = Nota1
         ElseIf OpcPeriodo("TQBim2", LstPeriodos) Then
            Nota1 = .Fields("TQBim1")
            Nota2 = .Fields("TQBim2")
            VectNota(I) = Nota2
         ElseIf OpcPeriodo("TQ", LstPeriodos) Then
            Nota1 = .Fields("TQBim1")
            Nota2 = .Fields("TQBim2")
            Nota3 = .Fields("ExamenTQ")
            Nota4 = .Fields("PromTQ")
            VectNota(I) = Nota4
'''         Else
'''            If Mid$(FormatoLibreta, 1, 9) = "QUIMESTRE" Then
'''                Nota1 = .Fields("PromPQ")
'''                Nota2 = .Fields("PromSQ")
'''
'''                Nota3 = .Fields("PromPQ") + .Fields("PromSQ")
'''                Nota4 = .Fields("Supletorio")
'''                Nota5 = .Fields("PromFinal")
'''                VectNota(I) = Nota5
'''            Else
'''                Nota1 = .Fields("PromPQ")
'''                Nota2 = .Fields("PromSQ")
'''                Nota3 = .Fields("PromTQ")
'''
'''                Nota4 = .Fields("PromPQ") + .Fields("PromSQ") + .Fields("PromTQ")
'''                Nota5 = .Fields("Supletorio")
'''                Nota6 = .Fields("PromFinal")
'''                VectNota(I) = Nota6
'''
'''            End If
         End If
         If Nota1 > 0 Then PictPrint_Nota_Materia PictCalif, PCol, PosLinea, Nota1, False
         PCol = PCol + 0.6
         If Nota2 > 0 Then PictPrint_Nota_Materia PictCalif, PCol, PosLinea, Nota2, False
         PCol = PCol + 0.6
         If Nota3 > 0 Then PictPrint_Nota_Materia PictCalif, PCol, PosLinea, Nota3, False
         PCol = PCol + 0.6
         If Nota4 > 0 Then PictPrint_Nota_Materia PictCalif, PCol, PosLinea, Nota4, False
         PCol = PCol + 0.6
'''         If Mid$(FormatoLibreta, 1, 9) = "TRIMESTRE" Then
'''            If Nota5 > 0 Then PictPrint_Nota_Materia PictCalif, PCol, PosLinea, Nota5, False
'''            PCol = PCol + 0.6
'''         End If
         SumaHoriz(I) = SumaHoriz(I) + VectNota(I)
         SumaVerti(I) = SumaVerti(I) + VectNota(I)
         I = I + 1
        .MoveNext
      Loop
     'MsgBox Contador & vbCrLf & ContNotas
      Total = 0
      J = 0
      Cadena = ""
      For I = 0 To ContNotas - 5
          If VectNota(I) > 0 Then
             J = J + 1
             Total = Total + VectNota(I)
          End If
      Next I
      If J > 0 Then Total = Total / J
      Saldo = Saldo + Total
      SumaHoriz(ContNotas) = SumaHoriz(ContNotas) + Total
        
      If Not OpcPeriodo("PF", LstPeriodos) Then PictPrint_Texto PictCalif, PCol, PosLinea, Format(Total, "00")
      PCol = PCol + 0.8
      'PosLinea = PosLinea + 0.45
      'Contador = Contador + 1
      PCol = 10: Total = 0
      PosLinea = PosLinea + 0.45
      PictCalif.FontBold = True
      If Contador <= 0 Then Contador = 1
      Saldo = Saldo / Contador
      If Not OpcPeriodo("PF", LstPeriodos) Then PictPrint_Texto PictCalif, PCol - 3.5, PosLinea, "TOTAL PROMEDIO:"
      For I = 0 To ContNotas - 5
         Not1 = 0: Not2 = 0: Not3 = 0: Not4 = 0: Not5 = 0
         If OpcPeriodo("PQBim1", LstPeriodos) Then Not1 = SumaVerti(I) / Contador
         If OpcPeriodo("PQBim2", LstPeriodos) Then Not4 = SumaVerti(I) / Contador
         If OpcPeriodo("SQBim1", LstPeriodos) Then Not1 = SumaVerti(I) / Contador
         If OpcPeriodo("SQBim2", LstPeriodos) Then Not4 = SumaVerti(I) / Contador
         If OpcPeriodo("TQBim1", LstPeriodos) Then Not1 = SumaVerti(I) / Contador
         If OpcPeriodo("TQBim2", LstPeriodos) Then Not4 = SumaVerti(I) / Contador
         If OpcPeriodo("PF", LstPeriodos) Then Not4 = SumaVerti(I) / Contador
         If Not1 <> 0 Then PictPrint_Texto PictCalif, PCol, PosLinea, Format(Not1, "00.00")
         PCol = PCol + 0.6
         If Not2 <> 0 Then PictPrint_Texto PictCalif, PCol, PosLinea, Format(Not2, "00.00")
         PCol = PCol + 0.6
         If Not3 <> 0 Then PictPrint_Texto PictCalif, PCol, PosLinea, Format(Not3, "00.00")
         PCol = PCol + 0.6
         If Not4 <> 0 Then PictPrint_Texto PictCalif, PCol, PosLinea, Format(Not4, "00.00")
         PCol = PCol + 0.6
'''         If Not5 <> 0 Then PictPrint_Texto PictCalif, PCol, PosLinea, Format(Not5, "00.00")
'''         PCol = PCol + 0.6
        'Cadena = Cadena & Contador & " >> " & I & " - " & SumaVerti(I) & " > " & VectMate(I) & vbCrLf
      Next I
     'MsgBox Cadena
      PCol = PCol + 0.15
      If Not OpcPeriodo("PF", LstPeriodos) Then PictPrint_Texto PictCalif, PCol, PosLinea, Format(Saldo, "00.00")
     'Rayas finales de la consulta
      PCol = 9.95
     'Rayas verticales de los Periodos
      For I = 0 To ContNotas - 5
        PictCalif.Line (PCol, 3.65)-(PCol, PosLinea)
        PCol = PCol + 0.6
        PictCalif.Line (PCol, 4.45)-(PCol, PosLinea)
        PCol = PCol + 0.6
        PictCalif.Line (PCol, 4.45)-(PCol, PosLinea)
        PCol = PCol + 0.6
        PictCalif.Line (PCol, 4.45)-(PCol, PosLinea)
        PCol = PCol + 0.6
'''        If Mid$(FormatoLibreta, 1, 9) = "TRIMESTRE" Then
'''           PictCalif.Line (PCol, 4.45)-(PCol, PosLinea)
'''           PCol = PCol + 0.6
'''        End If
      Next I
      PictCalif.Line (9.95, 4.45)-(PCol, 4.45)
      PictCalif.Line (PCol, 3.65)-(PCol, PosLinea)
      PCol = PCol + 0.8
      PictCalif.Line (0.7, PosLinea)-(PColMax, PosLinea)
      PictCalif.Line (0.7, 3.65)-(PCol, 3.65)
      PictCalif.Line (0.7, 6.3)-(PCol, 6.3)
      PictCalif.Line (0.7, 6.35)-(PCol, 6.35)
      PictCalif.Line (PCol, 3.65)-(PCol, PosLinea)
      PictCalif.Line (0.7, 3.65)-(0.7, PosLinea)
      PictCalif.FontBold = True
      PictCalif.FontSize = 20
      PictPrint_Texto PictCalif, 3, 4.5, "A P E L L I D O S"
      PictPrint_Texto PictCalif, 2.9, 5.3, "Y   N O M B R E S"
  End If
 End With
 RatonNormal
End Sub

Public Sub Lista_Alumnos_por_Curso(PorMeses As BookmarkEnum)
 NombreBanco = UCase(LstPeriodos.Text)
 FechaValida MBFecha
 PictCalif.Cls
 PictCalif.FontBold = True
 PCol = 10
 PictCalif.Height = 12 + (AdoPromedio1.Recordset.RecordCount * 0.45)
 If PorMeses Then
    PictCalif.width = 21
    If PictCalif.Height < 29 Then PictCalif.Height = 29
 Else
    PictCalif.width = 29
    If PictCalif.Height < 21 Then PictCalif.Height = 21
 End If
 PictCalif.FontName = TipoTimes
 PosLinea = 1
 If LogoTipo <> "" Then PictCalif.PaintPicture LoadPicture(LogoTipo), 0.1, PosLinea, 5, 2.5
 PictCalif.FontSize = 20
 PictPrint_Texto PictCalif, 1, PosLinea, Institucion1, , PictCalif.width, True
 PosLinea = PosLinea + 0.7
 PictPrint_Texto PictCalif, 1, PosLinea, Institucion2, , PictCalif.width, True
 PosLinea = PosLinea + 0.7
 PictCalif.FontSize = 12
 PictCalif.FontBold = False
 PictPrint_Texto PictCalif, 1, PosLinea, Direccion & " Teléfono: " & Telefono1, , PictCalif.width, True
 PosLinea = PosLinea + 0.6
 PictCalif.FontSize = 16
 PictCalif.FontBold = True
 PictPrint_Texto PictCalif, 1, PosLinea, "L I S T A    D E    A L U M N O S", , PictCalif.width, True
 PictCalif.FontBold = False
 PosLinea = PosLinea + 0.7
 PictCalif.FontSize = 12
 PictPrint_Texto PictCalif, 1, PosLinea, "AÑO LECTIVO: " & Anio_Lectivo, , PictCalif.width, True
 PosLinea = PosLinea + 0.6
 'PictPrint_Texto PictCalif, 0.8, PosLinea, NombreBanco
 PictPrint_Texto PictCalif, PictCalif.width - PictCalif.TextWidth(Dato_Curso.Nombre_Largo) - 2, PosLinea, Dato_Curso.Nombre_Largo
 PictCalif.FontName = TipoArialNarrow
 PictCalif.FontSize = 9
 PosLinea = 4.9 ' 4.4
 PictCalif.FontBold = True
 PictPrint_Texto PictCalif, 0.8, PosLinea, "MATRIC."
 PictPrint_Texto PictCalif, 3.5, PosLinea, "A P E L L I D O S   Y   N O M B R E S"
 PCol = 10
 If PorMeses Then
    NoMeses = Month(MBFecha)
    For I = 1 To 10
        Cadena = UCase(Mid$(MesesLetras(NoMeses), 1, 3))
        NoMeses = NoMeses + 1
        If NoMeses > 12 Then NoMeses = 1
        PictPrint_Texto PictCalif, PCol, PosLinea, Cadena
        PCol = PCol + 0.9
    Next
 Else
    PictPrint_Texto PictCalif, 10, PosLinea, "REPRESENTANTE"
    PictPrint_Texto PictCalif, 16, PosLinea, "DIRECCION"
    PictPrint_Texto PictCalif, 24.5, PosLinea, "TELEFONOS"
    PCol = 27.3
 End If
 AnchoFactura = PCol - 0.2
 PictCalif.FontSize = 9
 PosLinea = PosLinea + 0.5
 PCol = 10.25
 Contador = 0
 For I = 0 To ContNotas - 1
     VectNota(I) = 0
     SumaHoriz(I) = 0
 Next I
 I = 0
 PictCalif.FontBold = False
 With AdoResumen.Recordset
  If .RecordCount > 0 Then
     .MoveFirst
      Codigo = .Fields("Codigo")
      NombreCliente = .Fields("Cliente")
      Contador = Contador + 1
      PictPrint_Texto PictCalif, 0.8, PosLinea, Format(.Fields("Matricula_No"), "00000000")
      PictPrint_Texto PictCalif, 2, PosLinea, Format(Contador, "00") & ".-"
      PictPrint_Texto PictCalif, 2.6, PosLinea, .Fields("Cliente")
      If Not PorMeses Then
         PictPrint_Texto PictCalif, 10, PosLinea, .Fields("Representante")
         PictPrint_Texto PictCalif, 16, PosLinea, .Fields("Domicilio")
         Cadena = .Fields("Telefono_D") & " / " & .Fields("Telefono_R")
         PictPrint_Texto PictCalif, 24.5, PosLinea, Cadena
      End If
      Do While Not .EOF
         If Codigo <> .Fields("Codigo") Then
            PosLinea = PosLinea + 0.45
            Contador = Contador + 1
            PictCalif.Line (0.7, PosLinea)-(AnchoFactura, PosLinea)
            PosLinea = PosLinea + 0.05
            PictPrint_Texto PictCalif, 0.8, PosLinea, Format(.Fields("Matricula_No"), "00000000")
            PictPrint_Texto PictCalif, 2, PosLinea, Format(Contador, "00") & ".-"
            PictPrint_Texto PictCalif, 2.6, PosLinea, .Fields("Cliente")
            If Not PorMeses Then
               PictPrint_Texto PictCalif, 10, PosLinea, .Fields("Representante")
               PictPrint_Texto PictCalif, 16, PosLinea, .Fields("Domicilio")
               Cadena = .Fields("Telefono_D") & " / " & .Fields("Telefono_R")
               PictPrint_Texto PictCalif, 24.5, PosLinea, Cadena
            End If
            PCol = 10.25
            Codigo = .Fields("Codigo")
            NombreCliente = .Fields("Cliente")
            For I = 0 To ContNotas - 1
                VectNota(I) = 0
            Next I
            I = 0
         End If
         PCol = PCol + 0.8
        .MoveNext
      Loop
      PCol = PCol + 1.5
      PosLinea = PosLinea + 0.45
      PictCalif.Line (0.7, PosLinea)-(AnchoFactura, PosLinea)
      PictCalif.Line (0.7, 5.35)-(AnchoFactura, 5.35)
      PictCalif.Line (0.7, 4.8)-(AnchoFactura, 4.8)
      PictCalif.Line (0.7, 4.8)-(0.7, PosLinea)
      PictCalif.Line (1.9, 4.8)-(1.9, PosLinea)
      PictCalif.Line (AnchoFactura, 4.8)-(AnchoFactura, PosLinea)
      PCol = 9.9
      If PorMeses Then
         NoMeses = Month(MBFecha)
         For I = 1 To 10
             PictCalif.Line (PCol, 4.8)-(PCol, PosLinea)
             PCol = PCol + 0.9
         Next
      Else
         PictCalif.Line (9.9, 4.8)-(9.9, PosLinea)
         PictCalif.Line (15.9, 4.8)-(15.9, PosLinea)
         PictCalif.Line (24.4, 4.8)-(24.4, PosLinea)
      End If
  End If
 End With
End Sub

Public Sub Lista_Alumnos_por_Notas(PorMeses As BookmarkEnum)
 NombreBanco = UCase(LstPeriodos.Text)
 FechaValida MBFecha
 PictCalif.Cls
 PictCalif.FontBold = True
 PCol = 10
 PictCalif.Height = 12 + (AdoPromedio1.Recordset.RecordCount * 0.45)
 If PorMeses Then
    PictCalif.width = 21
    If PictCalif.Height < 29 Then PictCalif.Height = 29
 Else
    PictCalif.width = 29
    If PictCalif.Height < 21 Then PictCalif.Height = 21
 End If
 PictCalif.FontName = TipoTimes
 PosLinea = 1
 If LogoTipo <> "" Then PictCalif.PaintPicture LoadPicture(LogoTipo), 0.1, PosLinea, 5, 2.5
 PictCalif.FontSize = 20
 PictPrint_Texto PictCalif, 1, PosLinea, Institucion1, , PictCalif.width, True
 PosLinea = PosLinea + 0.7
 PictPrint_Texto PictCalif, 1, PosLinea, Institucion2, , PictCalif.width, True
 PosLinea = PosLinea + 0.7

 PictCalif.FontSize = 12
 PictCalif.FontBold = False
 PictPrint_Texto PictCalif, 1, PosLinea, Direccion & " Teléfono: " & Telefono1, , PictCalif.width, True
 PosLinea = PosLinea + 0.6
 PictCalif.FontSize = 16
 PictCalif.FontBold = True
 PictPrint_Texto PictCalif, 1, PosLinea, "L I S T A    D E    A L U M N O S", , PictCalif.width, True
 PictCalif.FontBold = False
 PosLinea = PosLinea + 0.7
 PictCalif.FontSize = 12
 PictPrint_Texto PictCalif, 1, PosLinea, "AÑO LECTIVO: " & Anio_Lectivo, , PictCalif.width, True
 PosLinea = PosLinea + 0.6
 PictPrint_Texto PictCalif, 0.8, PosLinea, Dato_Curso.Nombre_Largo
 PictCalif.FontName = TipoArialNarrow
 PictCalif.FontSize = 9
 PosLinea = 4.8
 PictCalif.FontBold = True
 PictPrint_Texto PictCalif, 0.8, PosLinea, "MATRIC."
 PictPrint_Texto PictCalif, 3.5, PosLinea, "A P E L L I D O S   Y   N O M B R E S"
 PCol = 10
 If PorMeses Then
    NoMeses = 1
    For I = 1 To 10
        Cadena = "N " & Format(NoMeses, "00")
        NoMeses = NoMeses + 1
        PictPrint_Texto PictCalif, PCol, PosLinea, Cadena
        PCol = PCol + 0.9
    Next
 Else
    PictPrint_Texto PictCalif, 10, PosLinea, "REPRESENTANTE"
    PictPrint_Texto PictCalif, 16, PosLinea, "DIRECCION"
    PictPrint_Texto PictCalif, 24.5, PosLinea, "TELEFONOS"
    PCol = 27.3
 End If
 AnchoFactura = PCol - 0.2
 PictCalif.FontSize = 9
 PosLinea = PosLinea + 0.5
 PCol = 10.25
 Contador = 0
 For I = 0 To ContNotas - 1
     VectNota(I) = 0
     SumaHoriz(I) = 0
 Next I
 I = 0
 PictCalif.FontBold = False
 With AdoResumen.Recordset
  If .RecordCount > 0 Then
     .MoveFirst
      Codigo = .Fields("Codigo")
      NombreCliente = .Fields("Cliente")
      Contador = Contador + 1
      PictPrint_Texto PictCalif, 0.8, PosLinea, Format(.Fields("Matricula_No"), "00000000")
      PictPrint_Texto PictCalif, 2, PosLinea, Format(Contador, "00") & ".-"
      PictPrint_Texto PictCalif, 2.6, PosLinea, .Fields("Cliente")
      If Not PorMeses Then
         PictPrint_Texto PictCalif, 10, PosLinea, .Fields("Representante")
         PictPrint_Texto PictCalif, 16, PosLinea, .Fields("Domicilio")
         Cadena = .Fields("Telefono_D") & " / " & .Fields("Telefono_R")
         PictPrint_Texto PictCalif, 24.5, PosLinea, Cadena
      End If
      Do While Not .EOF
         If Codigo <> .Fields("Codigo") Then
            PosLinea = PosLinea + 0.45
            Contador = Contador + 1
            PictCalif.Line (0.7, PosLinea)-(AnchoFactura, PosLinea)
            PosLinea = PosLinea + 0.05
            PictPrint_Texto PictCalif, 0.8, PosLinea, Format(.Fields("Matricula_No"), "00000000")
            PictPrint_Texto PictCalif, 2, PosLinea, Format(Contador, "00") & ".-"
            PictPrint_Texto PictCalif, 2.6, PosLinea, .Fields("Cliente")
            If Not PorMeses Then
               PictPrint_Texto PictCalif, 10, PosLinea, .Fields("Representante")
               PictPrint_Texto PictCalif, 16, PosLinea, .Fields("Domicilio")
               Cadena = .Fields("Telefono_D") & " / " & .Fields("Telefono_R")
               PictPrint_Texto PictCalif, 24.5, PosLinea, Cadena
            End If
            PCol = 10.25
            Codigo = .Fields("Codigo")
            NombreCliente = .Fields("Cliente")
            For I = 0 To ContNotas - 1
                VectNota(I) = 0
            Next I
            I = 0
         End If
         PCol = PCol + 0.8
        .MoveNext
      Loop
      PCol = PCol + 1.5
      PosLinea = PosLinea + 0.45
      PictCalif.Line (0.7, PosLinea)-(AnchoFactura, PosLinea)
      PictCalif.Line (0.7, 5.35)-(AnchoFactura, 5.35)
      PictCalif.Line (0.7, 4.8)-(AnchoFactura, 4.8)
      PictCalif.Line (0.7, 4.8)-(0.7, PosLinea)
      PictCalif.Line (1.9, 4.8)-(1.9, PosLinea)
      PictCalif.Line (AnchoFactura, 4.8)-(AnchoFactura, PosLinea)
      PCol = 9.9
      If PorMeses Then
         NoMeses = Month(MBFecha)
         For I = 1 To 10
             PictCalif.Line (PCol, 4.8)-(PCol, PosLinea)
             PCol = PCol + 0.9
         Next
      Else
         PictCalif.Line (9.9, 4.8)-(9.9, PosLinea)
         PictCalif.Line (15.9, 4.8)-(15.9, PosLinea)
         PictCalif.Line (24.4, 4.8)-(24.4, PosLinea)
      End If
  End If
 End With
 LimiteAlto = PictCalif.Height - 1.5
 PosLinea = PosLinea + 0.2
 PictCalif.FontBold = True
 PictPrint_Texto PictCalif, 0.8, PosLinea, "OBSERVACIONES:"
 PictCalif.FontBold = False
 PosLinea = PosLinea + 0.45
 PictCalif.Line (3.5, PosLinea)-(AnchoFactura, PosLinea)
 PosLinea = PosLinea + 0.6
 Do While PosLinea < LimiteAlto
    PictCalif.Line (0.7, PosLinea)-(AnchoFactura, PosLinea)
    PosLinea = PosLinea + 0.6
 Loop
 'MsgBox PosLinea & vbCrLf & LimiteAlto
End Sub

Public Sub Encabezado_Nomina_Oficial()
    PosLinea = 0.1
    'Institucion1
    Printer.FontSize = 18
    Printer.FontBold = True
    Printer.FontUnderline = True
    PrinterCentrarTexto 28, PosLinea, "NOMINA OFICIAL DE GRADUADOS"
    Printer.FontUnderline = False
    PosLinea = PosLinea + 1
    Printer.FontSize = 9
    Printer.Line (1.9, PosLinea)-(15, PosLinea + 1.05), Negro, B   ' Nombre del Colegio
    Printer.Line (1.9, PosLinea + 0.5)-(15, PosLinea + 0.5), Negro, B ' Linea
    PosLinea = PosLinea + 0.05
    PrinterTexto 2, PosLinea, "Nombre del Colegio"
    PosLinea = PosLinea + 0.3
    Printer.Line (24, PosLinea)-(25, PosLinea + 0.5), Negro, B  ' Hoja
    Printer.Line (24.5, PosLinea)-(24.5, PosLinea + 0.5), Negro ' Hoja
    Printer.Line (26, PosLinea)-(27, PosLinea + 0.5), Negro, B ' De
    Printer.Line (26.5, PosLinea)-(26.5, PosLinea + 0.5), Negro, B ' De
    PosLinea = PosLinea + 0.05
    PrinterTexto 2, PosLinea + 0.2, Institucion2
    PrinterTexto 23, PosLinea, "HOJA"
    PrinterTexto 25.2, PosLinea, "DE"
    PosLinea = PosLinea + 0.8
    PrinterTexto 2, PosLinea, "Titulo"
    PrinterTexto 9, PosLinea, "Tipo Titulo"
    PrinterTexto 19.4, PosLinea, "Especilización"
    PosLinea = PosLinea + 0.4
    Printer.Line (1.9, PosLinea)-(7.9, PosLinea + 0.6), Negro, B  ' Titulo
    Printer.Line (8, PosLinea)-(8.6, PosLinea + 0.6), Negro, B
    Printer.Line (7.9, PosLinea + 0.3)-(8, PosLinea + 0.3), Negro
    Printer.Line (8.9, PosLinea)-(18.3, PosLinea + 0.6), Negro, B  ' TipoTitulo
    Printer.Line (18.4, PosLinea)-(19, PosLinea + 0.6), Negro, B
    Printer.Line (18.3, PosLinea + 0.3)-(18.4, PosLinea + 0.3), Negro
    Printer.Line (19.3, PosLinea)-(26.3, PosLinea + 0.6), Negro, B  ' Especializacion
    Printer.Line (26.4, PosLinea)-(27, PosLinea + 0.6), Negro, B
    Printer.Line (26.3, PosLinea + 0.3)-(26.4, PosLinea + 0.3), Negro
    PosLinea = PosLinea + 0.05
    PrinterTexto 2, PosLinea, CodigoA
    PrinterTexto 9, PosLinea, CodigoB
    PrinterTexto 19.4, PosLinea, CodigoL
    
   'MsgBox CodigoA
    
    PosLinea = PosLinea + 0.8
    Printer.Line (1.9, PosLinea)-(3.3, PosLinea + 1), Negro, B    ' no Acta
    Printer.Line (3.4, PosLinea)-(5.8, PosLinea + 1), Negro, B    ' CI
    Printer.Line (5.9, PosLinea)-(17.7, PosLinea + 1), Negro, B   ' Nombres
    Printer.Line (17.8, PosLinea)-(19.2, PosLinea + 1), Negro, B  ' Sexo
    Printer.Line (19.3, PosLinea)-(23.7, PosLinea + 1), Negro, B  ' En letras
    Printer.Line (23.8, PosLinea)-(27, PosLinea + 1), Negro, B  ' Fecha año
    PosLinea = PosLinea + 0.05
    Printer.FontSize = 8
    PrinterTexto 1.9, PosLinea, "NÚMERO"
    PrinterTexto 3.4, PosLinea, "NÚMERO CEDULA"
    PrinterTexto 6, PosLinea, "APELLIDOS"
    PrinterTexto 12, PosLinea, "NOMBRES"
    PrinterTexto 18, PosLinea, "SEXO"
    PrinterTexto 19.5, PosLinea, "CALIFICACIONES"
    PrinterTexto 24, PosLinea, "FECHA DE GRADO"
    PosLinea = PosLinea + 0.5
    PrinterTexto 1.9, PosLinea, "DE ACTA"
    PrinterTexto 3.4, PosLinea, "DE IDENTIDAD"
    PrinterTexto 6, PosLinea, "Paterno"
    PrinterTexto 9, PosLinea, "Materno"
    PrinterTexto 18, PosLinea, "1H 2M"
    PrinterTexto 19.5, PosLinea, "NÚMERO EN LETRAS"
    PrinterTexto 24, PosLinea, "DD"
    PrinterTexto 25, PosLinea, "MM"
    PrinterTexto 26, PosLinea, "AAAA"
    PosLinea = PosLinea + 0.7
End Sub

Public Sub Listar_Materias_Curso(Curso As String)
  If Curso = "" Then Curso = Ninguno
  sSQL = "SELECT CE.TC,CE.CodigoE, CM.Materia, C.Cliente As Profesores, CE.CodMat " _
       & "FROM Catalogo_Estudiantil As CE,Clientes As C,Catalogo_Materias As CM " _
       & "WHERE CE.Item = '" & NumEmpresa & "' " _
       & "AND CE.Periodo = '" & Periodo_Contable & "' " _
       & "AND CE.TC = 'M' " _
       & "AND CE.Profesor = C.Codigo " _
       & "AND CE.CodMat = CM.CodMat " _
       & "AND CE.Item = CM.Item " _
       & "AND CE.Periodo = CM.Periodo " _
       & "AND Mid$(CE.CodigoE,1," & Len(Curso) & ") = '" & Curso & "' " _
       & "ORDER BY CE.CodigoE "
  SelectDBCombo DCMaterias, AdoMaterias, sSQL, "Materia"
End Sub

''Public Sub Imprimir_Actas_Calificaciones()
''Dim AnchoPict As Single
''Dim AltoPict As Single
''On Error GoTo Errorhandler
''Mensajes = "Seguro de Imprimir en:" & vbCrLf & Printer.DeviceName & "?"
''Titulo = "IMPRESION DE LIBRETAS"
''Bandera = False
''SetPrinters.Show 1
''If PonImpresoraDefecto(SetNombrePRN) Then
''   RatonReloj
''   Pagina = 1
''   InicioX = 0
''   InicioY = 0
''   Escala_Centimetro 1, TipoTimes, 9
''   Listar_Calificacion_Del_Curso Printer
''   RatonNormal
''   MensajeEncabData = ""
''   Printer.EndDoc
''   Exit Sub
''Errorhandler:
''             RatonNormal
''             ErrorDeImpresion
''             Exit Sub
''Else
''   RatonNormal
''End If
''End Sub

Public Sub Listar_Notas_Blanco(Curso As String)
Dim Listar_Curso As Boolean
Dim CantCampos As Integer
Dim IDCampos As Byte
   Progreso_Iniciar
   DGResumenNotas.Visible = False
   FrmPictCalif.Visible = True
   Contador = 0
   Listar_Curso = False
   Mensajes = "Listar por Cursos (Si/Yes)" & vbCrLf & vbCrLf _
            & "Listar Todo el Plantel (No/Not) "
   Titulo = "PRESENTACION"
   If BoxMensaje = vbYes Then Listar_Curso = True
       
   sSQL = "UPDATE Trans_Notas " _
        & "SET X = '.' " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND Periodo = '" & Periodo_Contable & "' "
   If Listar_Curso Then sSQL = sSQL & "AND CodE = '" & Curso & "' "
   ConectarAdoExecute sSQL
   sSQL = "UPDATE Trans_Notas_Auxiliares " _
        & "SET X = '.' " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND Periodo = '" & Periodo_Contable & "' "
   If Listar_Curso Then sSQL = sSQL & "AND CodE = '" & Curso & "' "
   ConectarAdoExecute sSQL
  
  'Verificamos las notas en blanco
   For IDCampos = 1 To UBound(VectCamposNotas)
    If Len(VectCamposNotas(IDCampos)) > 1 Then
       Progreso_Iniciar
       sSQL = "UPDATE Trans_Notas " _
            & "SET X = 'B' " _
            & "WHERE Item = '" & NumEmpresa & "' " _
            & "AND Periodo = '" & Periodo_Contable & "' " _
            & "AND " & VectCamposNotas(IDCampos) & " <=0 " _
            & "AND CodMat not IN ('998','999') "
       If Listar_Curso Then sSQL = sSQL & "AND CodE = '" & Curso & "' "
       ConectarAdoExecute sSQL
       Progreso_Iniciar
       sSQL = "UPDATE Trans_Notas_Auxiliares " _
            & "SET X = 'B' " _
            & "WHERE Item = '" & NumEmpresa & "' " _
            & "AND Periodo = '" & Periodo_Contable & "' " _
            & "AND " & VectCamposNotas(IDCampos) & " <=0 " _
            & "AND CodMat not IN ('998','999') "
       If Listar_Curso Then sSQL = sSQL & "AND CodE = '" & Curso & "' "
       ConectarAdoExecute sSQL
    End If
   Next
    
   sSQL = "SELECT TN.CodE,CM.Materia,C.Cliente As Estudiante,"
   For IDCampos = 1 To UBound(VectCamposNotas)
       If Len(VectCamposNotas(IDCampos)) > 1 Then sSQL = sSQL & "TN." & VectCamposNotas(IDCampos) & ","
   Next
   sSQL = sSQL & "TN.CodMat " _
        & "FROM Trans_Notas As TN,Clientes As C,Catalogo_Materias As CM " _
        & "WHERE TN.Item = '" & NumEmpresa & "' " _
        & "AND TN.Periodo = '" & Periodo_Contable & "' "
   If Listar_Curso Then
      sSQL = sSQL & "AND TN.CodE = '" & Curso & "' "
   Else
      sSQL = sSQL & "AND TN.CodE >= '1.02' "
   End If
   sSQL = sSQL _
        & "AND TN.X = 'B' " _
        & "AND TN.Codigo = C.Codigo " _
        & "AND TN.CodMat = CM.CodMat " _
        & "AND TN.Item = CM.Item " _
        & "AND TN.Periodo = CM.Periodo " _
        & "ORDER BY TN.CodE,CM.Materia,C.Cliente "
   SelectDataGrid DGResumenNotas, AdoResumenNotas, sSQL, , True
   DGResumenNotas.Visible = True
   FrmPictCalif.Visible = False
   Progreso_Final
End Sub
 
