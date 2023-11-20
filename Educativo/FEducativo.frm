VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.4#0"; "comctl32.Ocx"
Begin VB.Form FEducativo 
   Caption         =   "CATALOGO ESTUDIANTIL"
   ClientHeight    =   8730
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14640
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8730
   ScaleWidth      =   14640
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   30
      Top             =   0
      Width           =   14640
      _ExtentX        =   25823
      _ExtentY        =   1164
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   12
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Salir"
            Object.ToolTipText     =   "Salir del Modulo"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Imprimir"
            Object.ToolTipText     =   "Imprimir"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Grabar_Notas"
            Object.ToolTipText     =   "Grabar Notas"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Grabar_Notas_Grado"
            Object.ToolTipText     =   "Grabar Notas de Grado"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Grabar_Actas_Grado"
            Object.ToolTipText     =   "Grabar Actas de Grado"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Grabar_Promedio_Finales"
            Object.ToolTipText     =   "Grabar Promedio Finales"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Mejor_Puntaje"
            Object.ToolTipText     =   "Imprimir mejor puntaje"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Email"
            Object.ToolTipText     =   "Enviar por Correo"
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Actualizar_Cursos"
            Object.ToolTipText     =   "Actualiza Datos Generales de los Cursos"
            Object.Tag             =   ""
            ImageIndex      =   9
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Iniciar_Alumnos_Nuevos"
            Object.ToolTipText     =   "Inserta Notas de Alumnos Nuevos"
            Object.Tag             =   ""
            ImageIndex      =   10
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Actualiza_Cambio_Curso"
            Object.ToolTipText     =   "Actualiza el cambio de Curso o Paralelo"
            Object.Tag             =   ""
            ImageIndex      =   11
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Todas_Notas_Emails"
            Object.ToolTipText     =   "Enviar Notas de Todas las Materias"
            Object.Tag             =   ""
            ImageIndex      =   12
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.ListBox LstPeriodos 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1020
      Left            =   105
      TabIndex        =   25
      Top             =   1050
      Width           =   6735
   End
   Begin ComctlLib.TreeView TVNivel 
      Height          =   1800
      Left            =   105
      TabIndex        =   29
      Top             =   2100
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   3175
      _Version        =   327682
      Style           =   7
      ImageList       =   "ImgList"
      Appearance      =   1
   End
   Begin MSDataGridLib.DataGrid DGDetalle 
      Bindings        =   "FEducativo.frx":0000
      Height          =   7155
      Left            =   6930
      TabIndex        =   27
      Top             =   1050
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   12621
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
   Begin TabDlg.SSTab SSTabMaterias 
      Height          =   4020
      Left            =   105
      TabIndex        =   4
      Top             =   3990
      Width           =   6765
      _ExtentX        =   11933
      _ExtentY        =   7091
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Materias de Grado"
      TabPicture(0)   =   "FEducativo.frx":0019
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "DLMaterias"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Materias Asignadas del Curso"
      TabPicture(1)   =   "FEducativo.frx":0035
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DGCatalogoGrado"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "DATOS DE CURSO"
      TabPicture(2)   =   "FEducativo.frx":0051
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Command11"
      Tab(2).Control(1)=   "TxtTipo_Titulo"
      Tab(2).Control(2)=   "TxtTitulo"
      Tab(2).Control(3)=   "TxtEspecialidad"
      Tab(2).Control(4)=   "TxtBachiller"
      Tab(2).Control(5)=   "TxtCodigo_Titulo"
      Tab(2).Control(6)=   "TxtParalelo"
      Tab(2).Control(7)=   "TxtCiclo"
      Tab(2).Control(8)=   "TxtSeccion"
      Tab(2).Control(9)=   "Label13"
      Tab(2).Control(10)=   "Label12"
      Tab(2).Control(11)=   "Label11"
      Tab(2).Control(12)=   "Label9"
      Tab(2).Control(13)=   "Label8"
      Tab(2).Control(14)=   "Label7"
      Tab(2).Control(15)=   "Label6"
      Tab(2).Control(16)=   "Label5"
      Tab(2).Control(17)=   "Label2"
      Tab(2).ControlCount=   18
      Begin VB.CommandButton Command11 
         Caption         =   "Grabar Datos"
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
         Left            =   -69330
         TabIndex        =   24
         Top             =   735
         Width           =   855
      End
      Begin VB.TextBox TxtTipo_Titulo 
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
         TabIndex        =   21
         Top             =   3570
         Visible         =   0   'False
         Width           =   5475
      End
      Begin VB.TextBox TxtTitulo 
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
         TabIndex        =   19
         Top             =   2940
         Visible         =   0   'False
         Width           =   6420
      End
      Begin VB.TextBox TxtEspecialidad 
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
         TabIndex        =   17
         Top             =   2310
         Visible         =   0   'False
         Width           =   6420
      End
      Begin VB.TextBox TxtBachiller 
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
         TabIndex        =   15
         Top             =   1680
         Visible         =   0   'False
         Width           =   6420
      End
      Begin VB.TextBox TxtCodigo_Titulo 
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
         Left            =   -69435
         MaxLength       =   6
         TabIndex        =   23
         Text            =   "000000"
         Top             =   3570
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.TextBox TxtParalelo 
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
         Left            =   -70590
         MaxLength       =   5
         TabIndex        =   13
         Top             =   1050
         Width           =   1170
      End
      Begin VB.TextBox TxtCiclo 
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
         Left            =   -73215
         MaxLength       =   20
         TabIndex        =   11
         Top             =   1050
         Width           =   2640
      End
      Begin VB.TextBox TxtSeccion 
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
         MaxLength       =   12
         TabIndex        =   9
         Top             =   1050
         Width           =   1695
      End
      Begin MSDataGridLib.DataGrid DGCatalogoGrado 
         Bindings        =   "FEducativo.frx":006D
         Height          =   3435
         Left            =   -74895
         TabIndex        =   6
         Top             =   420
         Width           =   6525
         _ExtentX        =   11509
         _ExtentY        =   6059
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
         AllowDelete     =   -1  'True
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
      Begin MSDataListLib.DataList DLMaterias 
         Bindings        =   "FEducativo.frx":008C
         DataSource      =   "AdoMaterias"
         Height          =   3375
         Left            =   120
         TabIndex        =   5
         Top             =   420
         Visible         =   0   'False
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   5953
         _Version        =   393216
         ForeColor       =   128
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
      Begin VB.Label Label13 
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
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   -74895
         TabIndex        =   7
         Top             =   420
         Width           =   6420
      End
      Begin VB.Label Label12 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " CODIGO"
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
         Left            =   -69435
         TabIndex        =   22
         Top             =   3255
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Label Label11 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " TIPO TITULO DEL ACTA DE GRADO"
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
         Left            =   -74895
         TabIndex        =   20
         Top             =   3255
         Visible         =   0   'False
         Width           =   5475
      End
      Begin VB.Label Label9 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " TITULO DEL ACTA DE GRADO"
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
         Left            =   -74895
         TabIndex        =   18
         Top             =   2625
         Visible         =   0   'False
         Width           =   6420
      End
      Begin VB.Label Label8 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " ESPECIALIDAD"
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
         Left            =   -74895
         TabIndex        =   16
         Top             =   1995
         Visible         =   0   'False
         Width           =   6420
      End
      Begin VB.Label Label7 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " BACHILLER"
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
         Left            =   -74895
         TabIndex        =   14
         Top             =   1365
         Visible         =   0   'False
         Width           =   6420
      End
      Begin VB.Label Label6 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " PARALELO"
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
         Left            =   -70590
         TabIndex        =   12
         Top             =   735
         Width           =   1170
      End
      Begin VB.Label Label5 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " SECCION"
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
         Left            =   -74895
         TabIndex        =   8
         Top             =   735
         Width           =   1695
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " CICLO"
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
         Left            =   -73215
         TabIndex        =   10
         Top             =   735
         Width           =   2640
      End
   End
   Begin VB.TextBox Text1 
      Height          =   390
      Left            =   6930
      TabIndex        =   3
      Top             =   8295
      Width           =   435
   End
   Begin MSAdodcLib.Adodc AdoNivel 
      Height          =   330
      Left            =   315
      Top             =   1740
      Visible         =   0   'False
      Width           =   2640
      _ExtentX        =   4657
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
      Caption         =   "Nivel"
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
   Begin MSAdodcLib.Adodc AdoDetalle 
      Height          =   330
      Left            =   7455
      Top             =   8295
      Width           =   3165
      _ExtentX        =   5583
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
      Caption         =   "Detalle"
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
   Begin MSAdodcLib.Adodc AdoMaterias 
      Height          =   330
      Left            =   315
      Top             =   2055
      Visible         =   0   'False
      Width           =   2640
      _ExtentX        =   4657
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
   Begin MSAdodcLib.Adodc AdoAux 
      Height          =   330
      Left            =   315
      Top             =   2370
      Visible         =   0   'False
      Width           =   2640
      _ExtentX        =   4657
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
   Begin MSAdodcLib.Adodc AdoActas 
      Height          =   330
      Left            =   315
      Top             =   2685
      Visible         =   0   'False
      Width           =   2640
      _ExtentX        =   4657
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
      Caption         =   "Actas"
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
   Begin MSAdodcLib.Adodc AdoAsistencia 
      Height          =   330
      Left            =   315
      Top             =   3000
      Visible         =   0   'False
      Width           =   2640
      _ExtentX        =   4657
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
      Caption         =   "Asistencia"
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
   Begin MSAdodcLib.Adodc AdoAutorizar 
      Height          =   330
      Left            =   315
      Top             =   1425
      Visible         =   0   'False
      Width           =   2640
      _ExtentX        =   4657
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
      Caption         =   "Autorizar"
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
   Begin MSAdodcLib.Adodc AdoNGrado 
      Height          =   330
      Left            =   315
      Top             =   3315
      Visible         =   0   'False
      Width           =   2640
      _ExtentX        =   4657
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
      Caption         =   "NGrado"
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
   Begin MSAdodcLib.Adodc AdoEvalua 
      Height          =   330
      Left            =   315
      Top             =   3630
      Visible         =   0   'False
      Width           =   2640
      _ExtentX        =   4657
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
      Caption         =   "Evalua"
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
   Begin MSAdodcLib.Adodc AdoPromedios 
      Height          =   330
      Left            =   315
      Top             =   3945
      Visible         =   0   'False
      Width           =   2640
      _ExtentX        =   4657
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
      Caption         =   "Promedios"
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
   Begin MSAdodcLib.Adodc AdoCatalogoGrado 
      Height          =   330
      Left            =   315
      Top             =   4305
      Visible         =   0   'False
      Width           =   2640
      _ExtentX        =   4657
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
      Caption         =   "CatalogoGrado"
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
   Begin MSAdodcLib.Adodc AdoCurso 
      Height          =   330
      Left            =   315
      Top             =   4620
      Visible         =   0   'False
      Width           =   2640
      _ExtentX        =   4657
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
      Caption         =   "Curso"
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
   Begin MSAdodcLib.Adodc AdoCursos 
      Height          =   330
      Left            =   2940
      Top             =   1470
      Visible         =   0   'False
      Width           =   2640
      _ExtentX        =   4657
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
      Caption         =   "Cursos"
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
      Left            =   315
      Top             =   4935
      Visible         =   0   'False
      Width           =   2640
      _ExtentX        =   4657
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
   Begin MSAdodcLib.Adodc AdoNotasA 
      Height          =   330
      Left            =   315
      Top             =   5250
      Visible         =   0   'False
      Width           =   2640
      _ExtentX        =   4657
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
      Caption         =   "NotasA"
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
   Begin ComctlLib.ImageList ImgList 
      Left            =   13860
      Top             =   2205
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   6
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FEducativo.frx":00A6
            Key             =   "C"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FEducativo.frx":03C0
            Key             =   "N"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FEducativo.frx":06DA
            Key             =   "E"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FEducativo.frx":09F4
            Key             =   "M"
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FEducativo.frx":0D0E
            Key             =   "H"
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FEducativo.frx":1028
            Key             =   "Mj"
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   13755
      Top             =   1365
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   13
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FEducativo.frx":1342
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FEducativo.frx":165C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FEducativo.frx":1976
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FEducativo.frx":1C90
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FEducativo.frx":979A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FEducativo.frx":16B64
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FEducativo.frx":378B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FEducativo.frx":5CB2C
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FEducativo.frx":5CE46
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FEducativo.frx":5D160
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FEducativo.frx":5D47A
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FEducativo.frx":5D794
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FEducativo.frx":5D9D2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label LblMail 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " SELECCIONE EL &PERIODO"
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
      Left            =   6930
      TabIndex        =   28
      Top             =   735
      Width           =   6735
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " SELECCIONE EL &PERIODO"
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
      TabIndex        =   26
      Top             =   735
      Width           =   6735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "<Ctrl+F5> Modificar|<Ctrl+F6> No Modifica|<Ctrl+Ins> Insertar|<Ctrl+B> Buscar|<Ctrl+Supr> Eliminar|<Ctrl+V> Cambio de Valores"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   855
      Left            =   105
      TabIndex        =   2
      Top             =   8085
      Width           =   5265
   End
   Begin VB.Label LblValor 
      Alignment       =   2  'Center
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
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   5460
      TabIndex        =   0
      Top             =   8400
      Width           =   1275
   End
   Begin VB.Label Label10 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Valor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   330
      Left            =   5460
      TabIndex        =   1
      Top             =   8085
      Width           =   1275
   End
End
Attribute VB_Name = "FEducativo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Tipo_Insert_Materias As Byte
'
Public Sub Actualiza_Cambio_Curso()
    RatonReloj
    Progreso_Barra.Mensaje_Box = "Actualizacion de Cursos"
    Progreso_Iniciar
    DGDetalle.Visible = False
   'Actualizamos Los Alumnos Matriculados
    If Periodo_Contable = Ninguno Then
       Progreso_Barra.Mensaje_Box = "Actualizando Cursos"
       Progreso_Esperar
       RatonReloj
       If SQL_Server Then
          sSQL = "UPDATE Clientes_Matriculas " _
               & "SET Grupo_No = C.Grupo " _
               & "FROM Clientes_Matriculas As CM,Clientes As C "
       Else
          sSQL = "UPDATE Clientes_Matriculas As CM,Clientes As C " _
               & "SET CM.Grupo_No = C.Grupo "
       End If
       sSQL = sSQL _
            & "WHERE CM.Item = '" & NumEmpresa & "' " _
            & "AND CM.Periodo = '" & Periodo_Contable & "' " _
            & "AND C.FA <> " & Val(adFalse) & " " _
            & "AND LEN(C.Grupo) = 7 " _
            & "AND CM.Codigo = C.Codigo "
       ConectarAdoExecute sSQL
    
       Progreso_Barra.Mensaje_Box = "Actualizando Clientes"
       Progreso_Esperar
       sSQL = "UPDATE Clientes " _
            & "SET X = '.' " _
            & "WHERE FA <> " & Val(adFalse) & " " _
            & "AND LEN(Grupo) = 7 "
       ConectarAdoExecute sSQL
       
       Progreso_Esperar
       If SQL_Server Then
          sSQL = "UPDATE Clientes " _
               & "SET X = 'V' " _
               & "FROM Clientes As C, Clientes_Matriculas As CM "
       Else
          sSQL = "UPDATE Clientes As C, Clientes_Matriculas As CM " _
               & "SET X = 'V' "
       End If
       sSQL = sSQL _
            & "WHERE CM.Item = '" & NumEmpresa & "' " _
            & "AND CM.Periodo = '" & Periodo_Contable & "' " _
            & "AND C.FA <> " & Val(adFalse) & " " _
            & "AND LEN(C.Grupo) = 7 " _
            & "AND C.Codigo = CM.Codigo "
       ConectarAdoExecute sSQL
       
       Progreso_Barra.Mensaje_Box = "Actualizando Matriculas Clientes"
       Progreso_Esperar
       sSQL = "SELECT CC.Curso,C.Codigo,C.FA,C.Cliente " _
            & "FROM Clientes As C, Catalogo_Cursos As CC " _
            & "WHERE C.FA <> " & Val(adFalse) & " " _
            & "AND C.X = '.' " _
            & "AND CC.Item = '" & NumEmpresa & "' " _
            & "AND CC.Periodo = '" & Periodo_Contable & "' " _
            & "AND CC.Curso = C.Grupo " _
            & "ORDER BY CC.Curso, C.Cliente "
       SelectAdodc AdoAux, sSQL
       RatonReloj
       
       With AdoAux.Recordset
        If .RecordCount > 0 Then
            Progreso_Barra.Valor_Maximo = Progreso_Barra.Valor_Maximo + .RecordCount
            Do While Not .EOF
               Progreso_Barra.Mensaje_Box = "Actualizando " & .Fields("Cliente")
               Progreso_Esperar
            
               NivelNo = .Fields("Curso")
               CodigoCli = .Fields("Codigo")
               sSQL = "SELECT * " _
                    & "FROM Clientes_Matriculas " _
                    & "WHERE Item = '" & NumEmpresa & "' " _
                    & "AND Periodo = '" & Periodo_Contable & "' " _
                    & "AND Codigo = '" & CodigoCli & "' "
               SelectAdodc AdoDetalle, sSQL
               RatonReloj
               If AdoDetalle.Recordset.RecordCount <= 0 Then
                  SetAdoAddNew "Clientes_Matriculas"
                  SetAdoFields "T", Normal
                  SetAdoFields "Codigo", CodigoCli
                  SetAdoFields "Grupo_No", NivelNo
                  SetAdoUpdate
               End If
              .MoveNext
            Loop
        End If
       End With
    End If
    Progreso_Barra.Mensaje_Box = "Actualizando Notas"
    Progreso_Esperar
    
    sSQL = "UPDATE Trans_Notas " _
         & "SET Orden = 9 " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND CodMat IN ('997','998','999') " _
         & "AND Orden <> 9 "
    ConectarAdoExecute sSQL
    
    Progreso_Barra.Mensaje_Box = "Actualizando Catalogo Cursos"
    Progreso_Esperar
    
    sSQL = "UPDATE Catalogo_Estudiantil " _
         & "SET Orden = 9 " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND CodMat IN ('997','998','999') " _
         & "AND Orden <> 9 "
    ConectarAdoExecute sSQL
    RatonNormal
    DGDetalle.Visible = True
    Progreso_Final
    MsgBox "Proceso Terminado, Vuelva a ingresar al módulo"
    Unload Me
End Sub

Public Sub ListarActasAlunmos(Optional OpcList As Boolean)
  'MsgBox OpcList
  If OpcList Then
     If FormatoLibreta = "BIMESTRES" Then
        SQLDec = ""
        sSQL = "SELECT Alumno,Notas,Trabajo,Investigacion,Evaluacion,Cedula,Id_No,CodigoU "
     Else
        SQLDec = "N_1_5 3|Trabajo 3|Inv_Prom_Sexto 3|Nota_Grado 3|."
        sSQL = "SELECT Alumno,Notas As N_1_5,Trabajo,Investigacion As Inv_Prom_Sexto,Evaluacion As Nota_Grado,Cedula,Id_No,CodigoU "
     End If
     sSQL = sSQL & "FROM Asiento_A " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND CodigoU = '" & CodigoUsuario & "' " _
          & "ORDER BY Alumno,Id_No "
  Else
     SQLDec = ""
     sSQL = "SELECT C.Cliente As Alumno,TN.*,CM.CI " _
          & "FROM Clientes As C,Trans_Actas As TN,Clientes_Matriculas As CM " _
          & "WHERE TN.CodE = '" & Codigo & "' " _
          & "AND TN.Item = '" & NumEmpresa & "' " _
          & "AND TN.Periodo = '" & Periodo_Contable & "' " _
          & "AND C.Codigo = TN.Codigo " _
          & "AND C.Codigo = CM.Codigo " _
          & "AND TN.Item = CM.Item " _
          & "AND TN.Periodo = CM.Periodo " _
          & "ORDER BY C.Sexo DESC,C.Cliente "
  End If
End Sub

Public Sub ListarPromediosAlunmos(Optional OpcList As Boolean)
  'MsgBox OpcList
  If OpcList Then
     SQLDec = "Promedio 3|."
     sSQL = "SELECT Id_No,Alumno,N_1er,N_2do,N_3er,N_4to,N_5to,Total,Promedio,Codigo,CodigoU " _
          & "FROM Asiento_A " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND CodigoU = '" & CodigoUsuario & "' " _
          & "ORDER BY Id_No,Alumno "
  Else
     SQLDec = "Promedio 3|."
     sSQL = "SELECT C.Cliente As Alumno,TN.*,CM.CI " _
          & "FROM Clientes As C,Trans_Promedios As TN,Clientes_Matriculas As CM " _
          & "WHERE TN.CodE = '" & Codigo & "' " _
          & "AND TN.Item = '" & NumEmpresa & "' " _
          & "AND TN.Periodo = '" & Periodo_Contable & "' " _
          & "AND C.Codigo = TN.Codigo " _
          & "AND C.Codigo = CM.Codigo " _
          & "AND TN.Item = CM.Item " _
          & "AND TN.Periodo = CM.Periodo " _
          & "ORDER BY C.Sexo DESC,C.Cliente "
  End If
End Sub

Public Sub LlenarCodigos()
  Dim nodX As Node
  Cadena1 = Tipo_Acceso_Educativo("", "CodigoE")
' Establece propiedades del control ImageList.
  TVNivel.Nodes.Clear
 ' TVNivel.SingleSel = True
  TVNivel.LineStyle = tvwTreeLines
' Crea un árbol con varios objetos Node sin ordenar.
  sSQL = "SELECT CE.*,CM.Materia,CM.C,CM.I,CM.P,C.Cliente As Dirigente,C.Email,C.Email2 " _
       & "FROM Catalogo_Estudiantil As CE, Catalogo_Materias AS CM, Clientes As C " _
       & "WHERE CE.Item = '" & NumEmpresa & "' " _
       & "AND CE.Periodo = '" & Periodo_Contable & "' " _
       & Cadena1 _
       & "AND CE.Item = CM.Item " _
       & "AND CE.Periodo = CM.Periodo " _
       & "AND CE.CodMat = CM.CodMat " _
       & "AND CE.Profesor = C.Codigo " _
       & "ORDER BY CE.CodigoE "
  SelectAdodc AdoNivel, sSQL
  'MsgBox sSQL
  With ImgList
   If AdoNivel.Recordset.RecordCount > 0 Then
      Do While Not AdoNivel.Recordset.EOF
         Codigo = "C" & AdoNivel.Recordset.Fields("CodigoE")
         CodigoL = AdoNivel.Recordset.Fields("CodigoE")
         TipoDoc = AdoNivel.Recordset.Fields("TC")
         TipoProc = AdoNivel.Recordset.Fields("CodMat")
         Select Case TipoDoc
           Case "M": Cadena = AdoNivel.Recordset.Fields("Materia")
           Case Else
                If AdoCursos.Recordset.RecordCount > 0 Then
                   AdoCursos.Recordset.MoveFirst
                   AdoCursos.Recordset.Find ("Curso = '" & CodigoL & "'")
                   If Not AdoCursos.Recordset.EOF Then Cadena = AdoCursos.Recordset.Fields("Descripcion")
                End If
         End Select
         Cadena = Quitar_Signos_Especiales(Cadena)
         If Len(Codigo) = 2 Then
            Set nodX = TVNivel.Nodes.Add(, , Codigo, Cadena, .ListImages(1).key, .ListImages(1).key)
         Else
            Select Case TipoDoc
              Case "N": Set nodX = TVNivel.Nodes.Add(CambioCodigoCtaSup(Codigo), tvwChild, Codigo, Cadena, .ListImages(2).key, .ListImages(2).key)
              Case "P": Set nodX = TVNivel.Nodes.Add(CambioCodigoCtaSup(Codigo), tvwChild, Codigo, Cadena, .ListImages(3).key, .ListImages(3).key)
              Case "M": Set nodX = TVNivel.Nodes.Add(CambioCodigoCtaSup(Codigo), tvwChild, Codigo, Cadena, .ListImages(4).key, .ListImages(4).key)
            End Select
         End If
         AdoNivel.Recordset.MoveNext
      Loop
   End If
  End With
  RatonNormal
End Sub

Public Sub EliminarCta()
  Codigo1 = Mid$(TVNivel.SelectedItem.key, 2, Len(TVNivel.SelectedItem.key) - 1)
  Cadena = SinEspaciosIzq(TVNivel.SelectedItem.key)
  With AdoNivel.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
4      .Find ("Codigo like '" & Codigo1 & "' ")
       If Not .EOF Then
          Mensajes = "Esta seguro que desea eliminar la " & vbCrLf _
                   & "Cuenta No. [" & TVNivel.SelectedItem & "]"
          Titulo = "Pregunta de Eliminacion"
          If BoxMensaje = vbYes Then
            .Delete
             TVNivel.Nodes.Remove TVNivel.SelectedItem.Index
          End If
       End If
   End If
  End With
End Sub

Public Sub Grabar_Notas()
  Mensajes = "Esta seguro de Grabar Notas "
  Titulo = "Pregunta de grabación"
  If BoxMensaje = vbYes Then
     Actualizar_Notas_del_Curso TipoDoc, CodigoCuentaSup(Codigo)
     Listar_Notas_Alunmos AdoAutorizar, TipoDoc, CodigoCuentaSup(Codigo), LstPeriodos
     SelectDataGrid DGDetalle, AdoDetalle, sSQL
     MsgBox "Proceso Terminado"
  End If
End Sub

Public Sub Grabar_Notas_Grado()
    sSQL = "SELECT * " _
         & "FROM Asiento_NG " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND CodMat = '" & TipoDoc & "' " _
         & "AND CodigoU = '" & CodigoUsuario & "' " _
         & "ORDER BY Id_No "
    SelectAdodc AdoAux, sSQL
    With AdoAux.Recordset
     If .RecordCount > 0 Then
         Do While Not .EOF
            
            Real1 = .Fields("Examen")
            CodigoCli = .Fields("Codigo")
            Codigo1 = .Fields("CodE")
            Codigo2 = .Fields("CodMat")
           'MsgBox Codigo1
            sSQL = "SELECT * " _
                 & "FROM Trans_Notas_Grado " _
                 & "WHERE Item = '" & NumEmpresa & "' " _
                 & "AND Periodo = '" & Periodo_Contable & "' " _
                 & "AND Codigo = '" & CodigoCli & "' " _
                 & "AND CodE = '" & Codigo1 & "' " _
                 & "AND CodMat = '" & Codigo2 & "' "
            SelectAdodc AdoNGrado, sSQL
            If AdoNGrado.Recordset.RecordCount > 0 Then
               AdoNGrado.Recordset.Fields("Examen") = Real1
               AdoNGrado.Recordset.Update
            Else
               SetAdoAddNew "Trans_Notas_Grado"
               SetAdoFields "Id_No", Contador
               SetAdoFields "Codigo", CodigoCli
               SetAdoFields "CodMat", Codigo2
               SetAdoFields "CodE", Codigo1
               SetAdoFields "Examen", Real1
               SetAdoFields "Item", NumEmpresa
               SetAdoUpdate
            End If
           .MoveNext
         Loop
     End If
    End With
    MsgBox "Proceso de Grabación terminado"
End Sub

Private Sub Command11_Click()
    TextoValido TxtSeccion, , True
    TextoValido TxtCiclo, , True
    TextoValido TxtParalelo, , True
    TextoValido TxtBachiller, , True
    TextoValido TxtEspecialidad, , True
    TextoValido TxtTitulo, , True
    TextoValido TxtTipo_Titulo, , True
    TextoValido TxtCodigo_Titulo, , True
    If TxtBachiller = Ninguno Then TxtBachiller = Label13.Caption
    If TxtSeccion = Ninguno Then TxtSeccion = "MATUTINA"
    If TxtCiclo = Ninguno Then
       Select Case Mid$(Codigo, 1, 1)
         Case "1", "2": TxtCiclo = "CICLO BÁSICO"
         Case "3": TxtCiclo = "CICLO DIVERSIFICADO"
       End Select
    End If
    TxtCodigo_Titulo = Format(Val(TxtCodigo_Titulo), "000000")
    Label13.Caption = Cuenta
    sSQL = "DELETE * " _
         & "FROM Catalogo_Cursos " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND Curso = '" & Codigo & "' "
    ConectarAdoExecute sSQL
    SetAdoAddNew "Catalogo_Cursos"
    SetAdoFields "Curso", Codigo
    SetAdoFields "Descripcion", Label13.Caption
    SetAdoFields "Paralelo", TxtParalelo
    SetAdoFields "Bachiller", TxtBachiller
    SetAdoFields "Especialidad", TxtEspecialidad
    SetAdoFields "Ciclo", TxtCiclo
    SetAdoFields "Seccion", TxtSeccion
    SetAdoFields "Titulo", TxtTitulo
    SetAdoFields "Tipo_Titulo", TxtTipo_Titulo
    SetAdoFields "Codigo_Titulo", TxtCodigo_Titulo
    SetAdoFields "Periodo", Periodo_Contable
    SetAdoFields "Item", NumEmpresa
    SetAdoUpdate
    SSTabMaterias.Tab = 0
    TVNivel.SetFocus
End Sub

Public Sub Grabar_Actas_Grado()
  Mensajes = "Esta seguro de Grabar Actas de Grado "
  Titulo = "Pregunta de grabación"
  If BoxMensaje = vbYes Then
     If SQL_Server Then
        sSQL = "UPDATE Trans_Actas " _
             & "SET Notas = AN.Notas," _
             & "Evaluacion = AN.Evaluacion," _
             & "Trabajo = AN.Trabajo," _
             & "Investigacion = AN.Investigacion," _
             & "Id_No = AN.Id_No," _
             & "PromFinal = ROUND((AN.Notas+AN.Evaluacion+AN.Trabajo+AN.Investigacion)/4,3) " _
             & "FROM Trans_Actas as TN,Asiento_A As AN "
     Else
        sSQL = "UPDATE Trans_Actas as TN,Asiento_A As AN " _
             & "SET TN.Notas = AN.Notas," _
             & "TN.Evaluacion = AN.Evaluacion," _
             & "TN.Trabajo = AN.Trabajo," _
             & "TN.Investigacion = AN.Investigacion," _
             & "TN.Id_No = AN.Id_No," _
             & "TN.N_1er = AN.N_1er," _
             & "TN.N_2do = AN.N_2do," _
             & "TN.N_3er = AN.N_3er," _
             & "TN.N_4to = AN.N_4to," _
             & "TN.N_5to = AN.N_5to," _
             & "TN.PromFinal = ROUND((AN.Notas+AN.Evaluacion+AN.Trabajo+AN.Investigacion)/4,3) "
     End If
     sSQL = sSQL & "WHERE AN.Item = '" & NumEmpresa & "' " _
          & "AND TN.Periodo = '" & Periodo_Contable & "' " _
          & "AND AN.CodigoU = '" & CodigoUsuario & "' " _
          & "AND TN.Item = AN.Item " _
          & "AND TN.Codigo = AN.Codigo "
     ConectarAdoExecute sSQL
    'Actualizar Cedula de alumno(a)s
     If SQL_Server Then
        sSQL = "UPDATE Clientes_Matriculas " _
             & "SET CI = AN.Cedula " _
             & "FROM Clientes_Matriculas as TN,Asiento_A As AN "
     Else
        sSQL = "UPDATE Clientes_Matriculas as TN,Asiento_A As AN " _
             & "SET TN.CI = AN.Cedula "
     End If
     sSQL = sSQL & "WHERE AN.Item = '" & NumEmpresa & "' " _
          & "AND TN.Periodo = '" & Periodo_Contable & "' " _
          & "AND AN.CodigoU = '" & CodigoUsuario & "' " _
          & "AND TN.Item = AN.Item " _
          & "AND TN.Codigo = AN.Codigo "
     ConectarAdoExecute sSQL
          
     sSQL = "DELETE * " _
          & "FROM Asiento_A " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND CodigoU = '" & CodigoUsuario & "' "
     ConectarAdoExecute sSQL
     ListarActasAlunmos True
     SelectDataGrid DGDetalle, AdoDetalle, sSQL
     MsgBox "Proceso Terminado"
   End If
End Sub

Public Sub Iniciar_Alumnos_Nuevos()
Dim Id_NoE As Byte
Dim SumaCampos As String
Dim InsertarAsistencia As Boolean
'Actualizamos las Notas y las Actas de grado en blanco
  If ClaveSupervisor Then
    'Borramos Notas del Alumnos y Materias en Blanco
     Eliminar_Notas_Cero "Trans_Actas"
     Eliminar_Notas_Cero "Trans_Asistencia"
     Eliminar_Notas_Cero "Trans_Notas"
     Eliminar_Notas_Cero "Trans_Notas_Auxiliares"
     Eliminar_Notas_Cero "Trans_Notas_Grado"
     Eliminar_Notas_Cero "Trans_Promedios"
     RatonReloj
     Contador = 0
     sSQL = "SELECT * " _
          & "FROM Catalogo_Estudiantil " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "AND TC = 'M' " _
          & "ORDER BY CodigoE "
     SelectAdodc AdoNivel, sSQL
    With AdoNivel.Recordset
     If .RecordCount > 0 Then
         Do While Not .EOF
            Id_NoE = Val(Mid$(.Fields("CodigoE"), Len(.Fields("CodigoE")) - 1, 2))
            Codigo1 = .Fields("CodMat")                             ' Codigo Materia
            Codigo4 = .Fields("CodMatP")                            ' Codigo Materia Auxiliares
            Codigo2 = Trim(CambioCodigoCtaSup(.Fields("CodigoE")))  ' Curso o Grupo
            
           'Notas del Alumnos y Materias
            sSQL = "SELECT * " _
                 & "FROM Trans_Notas " _
                 & "WHERE Item = '" & NumEmpresa & "' " _
                 & "AND Periodo = '" & Periodo_Contable & "' " _
                 & "AND CodMat = '" & Codigo1 & "' " _
                 & "AND CodE = '" & Codigo2 & "' " _
                 & "ORDER BY CodE, CodMat "
            SelectAdodc AdoNotas, sSQL
            
           'Notas Auxiliares del Alumnos y Materias
            sSQL = "SELECT * " _
                 & "FROM Trans_Notas_Auxiliares " _
                 & "WHERE Item = '" & NumEmpresa & "' " _
                 & "AND Periodo = '" & Periodo_Contable & "' " _
                 & "AND CodMat = '" & Codigo1 & "' " _
                 & "AND CodE = '" & Codigo2 & "' " _
                 & "ORDER BY CodE, CodMat "
            SelectAdodc AdoNotasA, sSQL
            
           'Asistencias y Conductas
            sSQL = "SELECT * " _
                 & "FROM Trans_Asistencia " _
                 & "WHERE Item = '" & NumEmpresa & "' " _
                 & "AND Periodo = '" & Periodo_Contable & "' " _
                 & "AND CodMat = '" & Codigo1 & "' " _
                 & "AND CodE = '" & Codigo2 & "' " _
                 & "ORDER BY CodE, CodMat "
            SelectAdodc AdoAsistencia, sSQL
            
           'Actas de Grado Si estan en sexto curso
            sSQL = "SELECT * " _
                 & "FROM Trans_Actas " _
                 & "WHERE Item = '" & NumEmpresa & "' " _
                 & "AND Periodo = '" & Periodo_Contable & "' " _
                 & "ORDER BY Codigo "
            SelectAdodc AdoActas, sSQL
            
           'Promedios de 1 a 5 de los alumnos
            sSQL = "SELECT * " _
                 & "FROM Trans_Promedios " _
                 & "WHERE Item = '" & NumEmpresa & "' " _
                 & "AND Periodo = '" & Periodo_Contable & "' " _
                 & "ORDER BY Codigo "
            SelectAdodc AdoPromedios, sSQL
            
           'Alumnos del Curso
            sSQL = "SELECT * " _
                 & "FROM Clientes " _
                 & "WHERE Grupo = '" & Codigo2 & "' " _
                 & "AND FA <> " & Val(adFalse) & " " _
                 & "ORDER BY Cliente "
            SelectAdodc AdoAux, sSQL
            'MsgBox Codigo2 & ": " & AdoAux.Recordset.RecordCount
            If AdoAux.Recordset.RecordCount > 0 Then
               Do While Not AdoAux.Recordset.EOF
                  Codigo3 = AdoAux.Recordset.Fields("Codigo")
                  FEducativo.Caption = "Progreso: " & Format(Contador / .RecordCount, "00%") & " " & Codigo2 & " - " & Codigo1 & " - " & Codigo3 & "..."
                  sSQL = "SELECT * " _
                       & "FROM Clientes_Matriculas " _
                       & "WHERE Item = '" & NumEmpresa & "' " _
                       & "AND Periodo = '" & Periodo_Contable & "' " _
                       & "AND Codigo = '" & Codigo3 & "' "
                  SelectAdodc AdoDetalle, sSQL
                  If AdoDetalle.Recordset.RecordCount <= 0 Then
                     SetAdoAddNew "Clientes_Matriculas"
                     SetAdoFields "Codigo", Codigo3
                     SetAdoFields "Grupo_No", Codigo2
                     SetAdoFields "Periodo", Periodo_Contable
                     SetAdoFields "Item", NumEmpresa
                     SetAdoUpdate
                  End If
                  
                 'Verificacmos si es parte de SubCtas o no (CodMatP)
                  If Codigo4 = Ninguno Then
                    'Insertamos Notas en Blanco
                     If AdoNotas.Recordset.RecordCount > 0 Then
                        AdoNotas.Recordset.MoveFirst
                        AdoNotas.Recordset.Find ("Codigo = '" & Codigo3 & "' ")
                        If AdoNotas.Recordset.EOF Then
                          'If Codigo3 = "035501301" Then MsgBox "...."
                           SetAdoAddNew "Trans_Notas"
                           SetAdoFields "Id_No", Id_NoE
                           SetAdoFields "CodE", Codigo2
                           SetAdoFields "CodMat", Codigo1
                           SetAdoFields "Codigo", Codigo3
                           SetAdoFields "Periodo", Periodo_Contable
                           SetAdoFields "Item", NumEmpresa
                           SetAdoUpdate
                        End If
                     Else
                        SetAdoAddNew "Trans_Notas"
                        SetAdoFields "Id_No", Id_NoE
                        SetAdoFields "CodE", Codigo2
                        SetAdoFields "CodMat", Codigo1
                        SetAdoFields "Codigo", Codigo3
                        SetAdoFields "Periodo", Periodo_Contable
                        SetAdoFields "Item", NumEmpresa
                        SetAdoUpdate
                     End If
                  Else
                    'Insertamos Notas Auxiliares en Blanco
                     If AdoNotasA.Recordset.RecordCount > 0 Then
                        AdoNotasA.Recordset.MoveFirst
                        AdoNotasA.Recordset.Find ("Codigo = '" & Codigo3 & "' ")
                        If AdoNotasA.Recordset.EOF Then
                           SetAdoAddNew "Trans_Notas_Auxiliares"
                           SetAdoFields "Id_No", Id_NoE
                           SetAdoFields "CodE", Codigo2
                           SetAdoFields "CodMat", Codigo1
                           SetAdoFields "CodMatP", Codigo4
                           SetAdoFields "Codigo", Codigo3
                           SetAdoFields "Periodo", Periodo_Contable
                           SetAdoFields "Item", NumEmpresa
                           SetAdoUpdate
                        End If
                     Else
                        SetAdoAddNew "Trans_Notas_Auxiliares"
                        SetAdoFields "Id_No", Id_NoE
                        SetAdoFields "CodE", Codigo2
                        SetAdoFields "CodMat", Codigo1
                        SetAdoFields "CodMatP", Codigo4
                        SetAdoFields "Codigo", Codigo3
                        SetAdoFields "Periodo", Periodo_Contable
                        SetAdoFields "Item", NumEmpresa
                        SetAdoUpdate
                     End If
                  End If
                 'Promedio de 1ro a 5to
                  If Mid$(Codigo2, 1, 4) = "3.03" Or Mid$(Codigo2, 1, 4) = "5.03" Then
                       If AdoPromedios.Recordset.RecordCount > 0 Then
                          AdoPromedios.Recordset.MoveFirst
                          AdoPromedios.Recordset.Find ("Codigo = '" & Codigo3 & "' ")
                          If AdoPromedios.Recordset.EOF Then
                             SetAdoAddNew "Trans_Promedios"
                             SetAdoFields "Id_No", Id_NoE
                             SetAdoFields "CodE", Codigo2
                             SetAdoFields "Codigo", Codigo3
                             SetAdoFields "Periodo", Periodo_Contable
                             SetAdoFields "Item", NumEmpresa
                             SetAdoUpdate
                          End If
                       Else
                          SetAdoAddNew "Trans_Promedios"
                          SetAdoFields "Id_No", Id_NoE
                          SetAdoFields "CodE", Codigo2
                          SetAdoFields "Codigo", Codigo3
                          SetAdoFields "Periodo", Periodo_Contable
                          SetAdoFields "Item", NumEmpresa
                          SetAdoUpdate
                       End If
                      'Es Acta de Grado
                     If AdoActas.Recordset.RecordCount > 0 Then
                        AdoActas.Recordset.MoveFirst
                        AdoActas.Recordset.Find ("Codigo = '" & Codigo3 & "' ")
                        If AdoActas.Recordset.EOF Then
                           SetAdoAddNew "Trans_Actas"
                           SetAdoFields "Id_No", Id_NoE
                           SetAdoFields "CodE", Codigo2
                           SetAdoFields "Codigo", Codigo3
                           SetAdoFields "Periodo", Periodo_Contable
                           SetAdoFields "Item", NumEmpresa
                           SetAdoUpdate
                        End If
                     Else
                        SetAdoAddNew "Trans_Actas"
                        SetAdoFields "Id_No", Val(Mid$(Codigo2, Len(Codigo2) - 1, 2))
                        SetAdoFields "CodE", Codigo2
                        SetAdoFields "Codigo", Codigo3
                        SetAdoFields "Periodo", Periodo_Contable
                        SetAdoFields "Item", NumEmpresa
                        SetAdoUpdate
                     End If
                  End If
                  InsertarAsistencia = False
                  If Asistencias And Codigo2 > "2" Then
                     InsertarAsistencia = True
                  ElseIf "997" <= Codigo1 And Codigo1 <= "999" Then
                     InsertarAsistencia = True
                  End If
                 'Asistencia de Cursos
                  If InsertarAsistencia Then
                     If AdoAsistencia.Recordset.RecordCount > 0 Then
                        AdoAsistencia.Recordset.MoveFirst
                        AdoAsistencia.Recordset.Find ("Codigo = '" & Codigo3 & "' ")
                        If AdoAsistencia.Recordset.EOF Then
                           SetAdoAddNew "Trans_Asistencia"
                           SetAdoFields "Id_No", Id_NoE
                           SetAdoFields "CodE", Codigo2
                           SetAdoFields "Codigo", Codigo3
                           SetAdoFields "CodMat", Codigo1
                           SetAdoFields "Periodo", Periodo_Contable
                           SetAdoFields "Item", NumEmpresa
                           SetAdoUpdate
                        End If
                     Else
                        SetAdoAddNew "Trans_Asistencia"
                        SetAdoFields "Id_No", Id_NoE
                        SetAdoFields "CodE", Codigo2
                        SetAdoFields "Codigo", Codigo3
                        SetAdoFields "CodMat", Codigo1
                        SetAdoFields "Periodo", Periodo_Contable
                        SetAdoFields "Item", NumEmpresa
                        SetAdoUpdate
                     End If
                  End If
                  AdoAux.Recordset.MoveNext
               Loop
            End If
            Contador = Contador + 1
            FEducativo.Caption = "Progreso: " & Format(Contador / .RecordCount, "00%") & " " & Codigo2 & " - " & Codigo1 & " - " & Codigo3
            FEducativo.Refresh
            RatonReloj
           .MoveNext
         Loop
     End If
    End With
    sSQL = "UPDATE Trans_Notas " _
         & "SET Orden = 9 " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND CodMat IN ('997','998','999') " _
         & "AND Orden <> 9 "
    ConectarAdoExecute sSQL
    
    sSQL = "UPDATE Catalogo_Estudiantil " _
         & "SET Orden = 9 " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND CodMat IN ('997','998','999') " _
         & "AND Orden <> 9 "
    ConectarAdoExecute sSQL
    
  RatonNormal
  MsgBox "Fin del Proceso," & vbCrLf _
       & "vuelva a ingresar"
  Unload FEducativo
  End If
End Sub

Public Sub Email()
Dim EsperarMail As Integer
Dim ContMat As Long
    If Len(LblMail.Caption) > 1 Then
       DGDetalle.Visible = False
       TMail.Adjunto = Enviar_Notas_por_Materia(Codigo, CodigoN, TipoDoc, NombreDocente)
       TMail.Asunto = "Solicitud de Envio de Notas"
       TMail.Mensaje = "Estimado Docente, descargue el archivo, si aparece un mensaje que dice: " & vbCrLf _
            & "'Desea abrir el archivo ahora?', Presionar el boton SI." & vbCrLf & vbCrLf _
            & "Si presenta un mensaje como este: Vista Protegida, Presionar el Boton que dice:" & vbCrLf _
            & "Habilitar edición"
       TMail.para = LblMail.Caption
       FEnviarCorreos.Show 1
       DGDetalle.Visible = True
    End If
End Sub

Public Sub Imprimir()
    Cuadricula = True
    MensajeEncabData = "ACTA DE CALIFICACIONES POR MATERIA"
    SQLMsg1 = UCase(LstPeriodos.Text)
    SQLMsg2 = "Lcdo(a). " & ULCase(NombreDocente)
    SQLMsg3 = "Curso: " & Codigo & ", Materia: " & TVNivel.SelectedItem
    ImprimirAdo AdoDetalle, True, 1, 9, True
End Sub

Public Sub Grabar_Promedio_Finales()
  Mensajes = "Esta seguro de Grabar Promedios de Alumnos "
  Titulo = "Pregunta de grabación"
  If BoxMensaje = vbYes Then
     If SQL_Server Then
        sSQL = "UPDATE Trans_Promedios " _
             & "SET Total = ROUND(AN.N_1er+AN.N_2do+AN.N_3er+AN.N_4to+AN.N_5to,2,0)," _
             & "PromFinal = ROUND((AN.N_1er+AN.N_2do+AN.N_3er+AN.N_4to+AN.N_5to)/5,3,0)," _
             & "N_1er = AN.N_1er," _
             & "N_2do = AN.N_2do," _
             & "N_3er = AN.N_3er," _
             & "N_4to = AN.N_4to," _
             & "N_5to = AN.N_5to " _
             & "FROM Trans_Promedios as TN,Asiento_A As AN "
     Else
        sSQL = "UPDATE Trans_Promedios as TN,Asiento_A As AN " _
             & "SET TN.Total = ROUND(AN.N_1er+AN.N_2do+AN.N_3er+AN.N_4to+AN.N_5to,2,0)," _
             & "TN.PromFinal = ROUND((AN.N_1er+AN.N_2do+AN.N_3er+AN.N_4to+AN.N_5to)/5,3,0)," _
             & "TN.N_1er = AN.N_1er," _
             & "TN.N_2do = AN.N_2do," _
             & "TN.N_3er = AN.N_3er," _
             & "TN.N_4to = AN.N_4to," _
             & "TN.N_5to = AN.N_5to "
     End If
     sSQL = sSQL _
          & "WHERE AN.Item = '" & NumEmpresa & "' " _
          & "AND TN.Periodo = '" & Periodo_Contable & "' " _
          & "AND AN.CodigoU = '" & CodigoUsuario & "' " _
          & "AND TN.Item = AN.Item " _
          & "AND TN.Codigo = AN.Codigo "
     ConectarAdoExecute sSQL
          
     sSQL = "DELETE * " _
          & "FROM Asiento_A " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND CodigoU = '" & CodigoUsuario & "' "
     ConectarAdoExecute sSQL
    'MsgBox "......"
     ListarPromediosAlunmos True
     SelectDataGrid DGDetalle, AdoDetalle, sSQL
     MsgBox "Proceso Terminado"
   End If
End Sub

Public Sub Mejor_Puntaje()
   MensajeEncabData = "PUNTAJE DE MEJORES ALUMNOS SECCION: " & TVNivel.SelectedItem.Text
   Imprimir_Mejores_Alumnos AdoDetalle, 1, 9
End Sub

Private Sub DGDetalle_KeyDown(KeyCode As Integer, Shift As Integer)
Dim PromTotalCurso As Boolean
  Keys_Especiales Shift
  If KeyCode = vbKeyRight Then
     'KeyCode = vbKeyReturn
     If Es_Promedios Then
        If AdoDetalle.Recordset.Fields("N_1er") > 20 Then AdoDetalle.Recordset.Fields("N_1er") = 20
        If AdoDetalle.Recordset.Fields("N_2do") > 20 Then AdoDetalle.Recordset.Fields("N_2do") = 20
        If AdoDetalle.Recordset.Fields("N_3er") > 20 Then AdoDetalle.Recordset.Fields("N_3er") = 20
        If AdoDetalle.Recordset.Fields("N_4to") > 20 Then AdoDetalle.Recordset.Fields("N_4to") = 20
        If AdoDetalle.Recordset.Fields("N_5to") > 20 Then AdoDetalle.Recordset.Fields("N_5to") = 20
        AdoDetalle.Recordset.Update
        Total = AdoDetalle.Recordset.Fields("N_1er")
        Total = Total + AdoDetalle.Recordset.Fields("N_2do")
        Total = Total + AdoDetalle.Recordset.Fields("N_3er")
        Total = Total + AdoDetalle.Recordset.Fields("N_4to")
        Total = Total + AdoDetalle.Recordset.Fields("N_5to")
        AdoDetalle.Recordset.Fields("Total") = Round(Total, 3)
        AdoDetalle.Recordset.Fields("Promedio") = Round(Total / 5, 3)
        AdoDetalle.Recordset.Update
     End If
  End If
  If KeyCode = vbKeyReturn Then
     AdoDetalle.Recordset.MoveNext
     If AdoDetalle.Recordset.EOF Then AdoDetalle.Recordset.MoveFirst
  End If
  If CtrlDown And KeyCode = vbKeyDelete Then
     'Codigo
     Codigo1 = DGDetalle.Columns(0).Text
     Codigo2 = DGDetalle.Columns(1).Text
     Mensajes = "Eliminar al Alumno(a): " & UCase(Codigo2) & vbCrLf & vbCrLf _
              & "Codigo del Alumno(a): " & Codigo1 & vbCrLf & vbCrLf _
              & "Curso del Plantel: " & Codigo
     Titulo = "PREGUNTA DE ELIMINACION DE ALUMNOS"
     If BoxMensaje = vbYes Then
        RatonReloj
        sSQL = "DELETE * " _
             & "FROM Trans_Notas " _
             & "WHERE Item = '" & NumEmpresa & "' " _
             & "AND Periodo = '" & Periodo_Contable & "' " _
             & "AND Codigo = '" & Codigo1 & "' " _
             & "AND CodE = '" & Codigo & "' "
        ConectarAdoExecute sSQL
        sSQL = "DELETE * " _
             & "FROM Trans_Actas " _
             & "WHERE Item = '" & NumEmpresa & "' " _
             & "AND Periodo = '" & Periodo_Contable & "' " _
             & "AND Codigo = '" & Codigo1 & "' " _
             & "AND CodE = '" & Codigo & "' "
        ConectarAdoExecute sSQL
        sSQL = "UPDATE Clientes " _
             & "SET Grupo = 'RET' " _
             & "WHERE Codigo = '" & Codigo1 & "' "
        ConectarAdoExecute sSQL
        RatonNormal
     End If
  End If
  If CtrlDown And KeyCode = vbKeyP Then
     Es_Promedios = False
     PromTotalCurso = False
     MensajeEncabData = ""
     SQLMsg1 = ""
     SQLMsg2 = ""
     SQLMsg3 = ""
     SQLMsg4 = ""
     For I = 0 To AdoDetalle.Recordset.Fields.Count - 1
         If AdoDetalle.Recordset.Fields(I).Name = "N_1er" Then Es_Promedios = True
         If AdoDetalle.Recordset.Fields(I).Name = "Curso" Then PromTotalCurso = True
     Next I
     SQLMsg2 = ""
     Cadena = Mid$(TVNivel.SelectedItem.key, 2, Len(TVNivel.SelectedItem.key) - 1)
     Cadena = CambioCodigoCtaSup(Cadena)
     With AdoNivel.Recordset
      If .RecordCount > 0 Then
         .MoveFirst
         .Find ("CodigoE like '" & Cadena & "' ")
          If Not .EOF Then
             MensajeEncabData = .Fields("Detalle")
          End If
      End If
     End With
     Cuadricula = True
     If Es_Promedios And PromTotalCurso = False Then
        MensajeEncabData = DGDetalle.Caption
        Mensajes = "IMPRESION DE PROMEDIOS DE 1RO. A 5TO."
        SQLDec = "Promedio 3|."
        Imprimir_Promedio_Notas AdoDetalle, 9, True
     ElseIf PromTotalCurso And PromTotalCurso Then
        MensajeEncabData = DGDetalle.Caption
        Mensajes = "IMPRESION DE PROMEDIOS DE 1RO. A 5TO."
        SQLDec = "Promedio 3|."
        Imprimir_Promedio_Notas AdoDetalle, 8, True, PromTotalCurso
     Else
        Mensajes = "IMPRESION DE PROMEDIOS DE NOTAS FINALES"
        SQLMsg2 = TVNivel.SelectedItem
        SQLDec = "N_1_5 3|Trabajo 3|Inv_Prom_Sexto 3|Nota_Grado 3|."
        Imprimir_Nomina_Notas AdoDetalle, AdoAutorizar
     End If
  End If
End Sub

Private Sub DLMaterias_KeyDown(KeyCode As Integer, Shift As Integer)
  Dim nodX As Node
  If KeyCode = vbKeyEscape Then DLMaterias.Visible = False
  If KeyCode = vbKeyReturn Then
     Codigo = SinEspaciosDer(DLMaterias.Text)
     Cuenta = Mid$(DLMaterias.Text, 1, Len(DLMaterias.Text) - Len(Codigo) - 2)
     Codigo1 = Mid$(TVNivel.SelectedItem.key, 2, Len(TVNivel.SelectedItem.key) - 1)
     If Tipo_Insert_Materias = 1 Then
        sSQL = "SELECT * " _
             & "FROM Catalogo_Estudiantil " _
             & "WHERE Item = '" & NumEmpresa & "' " _
             & "AND Periodo = '" & Periodo_Contable & "' " _
             & "AND TC = 'M' " _
             & "AND Mid$(CodigoE,1," & Len(Codigo1) & ") = '" & Codigo1 & "' " _
             & "ORDER BY CodigoE "
     Else
        sSQL = "SELECT * " _
             & "FROM Catalogo_Examen_Grado " _
             & "WHERE Item = '" & NumEmpresa & "' " _
             & "AND Periodo = '" & Periodo_Contable & "' " _
             & "AND TC = 'M' " _
             & "AND Mid$(CodigoE,1," & Len(Codigo1) & ") = '" & Codigo1 & "' " _
             & "ORDER BY CodigoE "
     End If
     SelectAdodc AdoAux, sSQL
     Si_No = True
     With AdoAux.Recordset
      If .RecordCount > 0 Then
          Do While Not .EOF
             If .Fields("CodMat") = Codigo Then Si_No = False
             Codigo2 = Mid$(.Fields("CodigoE"), Len(.Fields("CodigoE")) - 1, 2)
             Codigo2 = Format(Val(Codigo2) + 1, "00")
             'MsgBox "-> " & Codigo2
            .MoveNext
          Loop
      Else
          Codigo2 = "01"
      End If
     End With
     'MsgBox Codigo2
     If Si_No Then
        If Tipo_Insert_Materias = 1 Then
           Set nodX = TVNivel.Nodes.Add("C" & Codigo1, tvwChild, "C" & Codigo1 & "." & Codigo2, Cuenta, ImgList.ListImages(4).key, ImgList.ListImages(4).key)
           SetAdoAddNew "Catalogo_Estudiantil"
           SetAdoFields "CodigoE", Codigo1 & "." & Codigo2
        Else
           SetAdoAddNew "Catalogo_Examen_Grado"
           SetAdoFields "CodigoE", Codigo1
        End If
        SetAdoFields "TC", "M"
        SetAdoFields "CodMat", Codigo
        SetAdoFields "Detalle", Cuenta
        SetAdoFields "Periodo", Periodo_Contable
        SetAdoFields "Item", NumEmpresa
        SetAdoUpdate
        If Tipo_Insert_Materias = 1 Then
           TVNivel.Refresh
          'LlenarCodigos
           sSQL = "SELECT * " _
                & "FROM Catalogo_Estudiantil " _
                & "WHERE Item = '" & NumEmpresa & "' " _
                & "AND Periodo = '" & Periodo_Contable & "' " _
                & "ORDER BY CodigoE "
           SelectData AdoNivel, sSQL
        Else
           sSQL = "SELECT CodMat,Detalle,CodigoE,Item,Periodo " _
                & "FROM Catalogo_Examen_Grado " _
                & "WHERE Item = '" & NumEmpresa & "' " _
                & "AND Periodo = '" & Periodo_Contable & "' " _
                & "AND TC = 'M' " _
                & "AND Mid$(CodigoE,1," & Len(Codigo1) & ") = '" & Codigo1 & "' " _
                & "ORDER BY CodigoE "
           SelectDataGrid DGCatalogoGrado, AdoCatalogoGrado, sSQL
        End If
        
        MsgBox "Proceso de Asignación exitoso"
     End If
     SSTabMaterias.TabCaption(1) = ""
     'DLMaterias.Visible = False
     'TVNivel.SetFocus
  End If
End Sub

Private Sub Form_Activate()
Dim IdCurso As Byte
Dim CodCurso As String
  DGDetalle.AllowUpdate = Not Cierre_Periodo
  DGDetalle.Height = MDI_Y_Max - DGDetalle.Top - 300
  DGDetalle.width = MDI_X_Max - DGDetalle.Left
  LblMail.width = MDI_X_Max - DGDetalle.Left
  AdoDetalle.width = MDI_X_Max - DGDetalle.Left - 500
  AdoDetalle.Top = DGDetalle.Top + DGDetalle.Height + 50
  Text1.Top = DGDetalle.Top + DGDetalle.Height + 50
 'Actualizamos Catalogo_Estudiantil
  sSQL = "UPDATE Catalogo_Estudiantil " _
       & "SET Id_No = 0 " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND Id_No IS NULL "
  ConectarAdoExecute sSQL
  If SQL_Server Then
     sSQL = "UPDATE Catalogo_Estudiantil " _
          & "SET Id_No=CONVERT(TINYINT,Mid$(CodigoE,9,2)) "
  Else
     sSQL = "UPDATE Catalogo_Estudiantil " _
          & "SET Id_No=Val(Mid$(CodigoE,9,2)) "
  End If
  sSQL = sSQL & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' "
  ConectarAdoExecute sSQL
  
  sSQL = "SELECT * " _
       & "FROM Catalogo_Cursos " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND LEN(Curso) = 1 " _
       & "ORDER BY Curso "
  SelectAdodc AdoCursos, sSQL
  With AdoCursos.Recordset
   If .RecordCount <= 0 Then
       For IdCurso = 1 To 3
           CodCurso = CStr(IdCurso)
           SetAdoAddNew "Catalogo_Cursos"
           SetAdoFields "Curso", CodCurso
           Select Case CStr(IdCurso)
             Case "1": SetAdoFields "Descripcion", "EDUCACIÓN BÁSICA"
             Case "2": SetAdoFields "Descripcion", "EDUCACIÓN SECUNDARIA"
             Case "3": SetAdoFields "Descripcion", "BACHILLERATO"
           End Select
           SetAdoFields "Periodo", Periodo_Contable
           SetAdoFields "Item", NumEmpresa
           SetAdoUpdate
       Next IdCurso
   End If
  End With
  sSQL = "SELECT * " _
       & "FROM Catalogo_Cursos " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND LEN(Curso) = 4 " _
       & "ORDER BY Curso "
  SelectAdodc AdoCursos, sSQL
  With AdoCursos.Recordset
   If .RecordCount <= 0 Then
       For IdCurso = 0 To 7
           CodCurso = "1." & Format(IdCurso, "00")
           SetAdoAddNew "Catalogo_Cursos"
           SetAdoFields "Curso", CodCurso
           Select Case CStr(IdCurso)
             Case "0": SetAdoFields "Descripcion", "INICIAL DE EDUCACIÓN BÁSICA"
             Case "1": SetAdoFields "Descripcion", "PRIMERO DE EDUCACIÓN BÁSICA"
             Case "2": SetAdoFields "Descripcion", "SEGUNDO DE EDUCACIÓN BÁSICA"
             Case "3": SetAdoFields "Descripcion", "TERCERO DE EDUCACIÓN BÁSICA"
             Case "4": SetAdoFields "Descripcion", "CUARTO DE EDUCACIÓN BÁSICA"
             Case "5": SetAdoFields "Descripcion", "QUINTO DE EDUCACIÓN BÁSICA"
             Case "6": SetAdoFields "Descripcion", "SEXTO DE EDUCACIÓN BÁSICA"
             Case "7": SetAdoFields "Descripcion", "SEPTIMO DE EDUCACIÓN BÁSICA"
           End Select
           SetAdoFields "Periodo", Periodo_Contable
           SetAdoFields "Item", NumEmpresa
           SetAdoUpdate
       Next IdCurso
       For IdCurso = 0 To 7
           CodCurso = "1." & Format(IdCurso, "00")
           SetAdoAddNew "Catalogo_Cursos"
           SetAdoFields "Curso", CodCurso
           Select Case CStr(IdCurso)
             Case "0": SetAdoFields "Descripcion", "INICIAL DE EDUCACIÓN BÁSICA"
             Case "1": SetAdoFields "Descripcion", "PRIMERO DE EDUCACIÓN BÁSICA"
             Case "2": SetAdoFields "Descripcion", "SEGUNDO DE EDUCACIÓN BÁSICA"
             Case "3": SetAdoFields "Descripcion", "TERCERO DE EDUCACIÓN BÁSICA"
             Case "4": SetAdoFields "Descripcion", "CUARTO DE EDUCACIÓN BÁSICA"
             Case "5": SetAdoFields "Descripcion", "QUINTO DE EDUCACIÓN BÁSICA"
             Case "6": SetAdoFields "Descripcion", "SEXTO DE EDUCACIÓN BÁSICA"
             Case "7": SetAdoFields "Descripcion", "SEPTIMO DE EDUCACIÓN BÁSICA"
           End Select
           SetAdoFields "Periodo", Periodo_Contable
           SetAdoFields "Item", NumEmpresa
           SetAdoUpdate
       Next IdCurso
       For IdCurso = 8 To 10
           CodCurso = "2." & Format(IdCurso, "00")
           SetAdoAddNew "Catalogo_Cursos"
           SetAdoFields "Curso", CodCurso
           Select Case CStr(IdCurso)
             Case "8": SetAdoFields "Descripcion", "OCTAVO DE EDUCACIÓN BÁSICA"
             Case "9": SetAdoFields "Descripcion", "NOVENO DE EDUCACIÓN BÁSICA"
             Case "10": SetAdoFields "Descripcion", "DÉCIMO DE EDUCACIÓN BÁSICA"
           End Select
           SetAdoFields "Periodo", Periodo_Contable
           SetAdoFields "Item", NumEmpresa
           SetAdoUpdate
       Next IdCurso
       For IdCurso = 1 To 3
           CodCurso = "3." & Format(IdCurso, "00")
           SetAdoAddNew "Catalogo_Cursos"
           SetAdoFields "Curso", CodCurso
           Select Case CStr(IdCurso)
             Case "1": SetAdoFields "Descripcion", "PRIMERO DE BACHILLERATO"
             Case "2": SetAdoFields "Descripcion", "SEGUNDO DE BACHILLERATO"
             Case "3": SetAdoFields "Descripcion", "TERCERO DE BACHILLERATO"
           End Select
           SetAdoFields "Periodo", Periodo_Contable
           SetAdoFields "Item", NumEmpresa
           SetAdoUpdate
       Next IdCurso
   End If
  End With
  
  sSQL = "SELECT * " _
       & "FROM Catalogo_Cursos " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "ORDER BY Curso "
  SelectAdodc AdoCursos, sSQL
    
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
       
       LstPeriodos.AddItem "Todos Trimestres"
    ElseIf Mid$(FormatoLibreta, 1, 9) = "QUIMESTRE" Then
       LstPeriodos.AddItem "Primer Quimestre Primer Parcial"
       LstPeriodos.AddItem "Primer Quimestre Segundo Parcial"
       LstPeriodos.AddItem "Primer Quimestre Tercer Parcial"
       LstPeriodos.AddItem "Promedio Primer Quimestre"
       
       LstPeriodos.AddItem "Segundo Quimestre Primer Parcial"
       LstPeriodos.AddItem "Segundo Quimestre Segundo Parcial"
       LstPeriodos.AddItem "Segundo Quimestre Tercer Parcial"
       LstPeriodos.AddItem "Promedio Segundo Quimestre"
       
       LstPeriodos.AddItem "Todos los Quimestres"
    Else
       LstPeriodos.AddItem "Primer Quimestre Primer Periodo"
       LstPeriodos.AddItem "Primer Quimestre Segundo Periodo"
       LstPeriodos.AddItem "Promedio Primer Quimestre"
       
       LstPeriodos.AddItem "Segundo Quimestre Primer Periodo"
       LstPeriodos.AddItem "Segundo Quimestre Segundo Periodo"
       LstPeriodos.AddItem "Promedio Segundo Quimestre"
       
       LstPeriodos.AddItem "Todos los Periodos"
    End If
  
 'Llenamos la malla curricular
  LlenarCodigos
  sSQL = "SELECT * " _
       & "FROM Catalogo_Periodo_Lectivo " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' "
  SelectAdodc AdoAutorizar, sSQL
  With AdoAutorizar.Recordset
   If .RecordCount > 0 Then
       Si_No = .Fields("NPQP1") Or .Fields("NPQP2") Or .Fields("NPQEX") Or .Fields("NSQP1")
       Si_No = Si_No Or .Fields("NSQP2") Or .Fields("NSQEX") Or .Fields("NSUPL")
   End If
  End With
 'If Periodo_Contable <> Ninguno Then DGDetalle.AllowUpdate = False
  TVNivel.ToolTipText = "<Ctrl + A>: Ingresar Actas de Grado," _
                      & "<Alt+P>: Promedios de 1ero. a 5to.," _
                      & "<Ctrl+F9>: Insertar Materias de Grado," _
                      & "<Ctrl+D>: Datos del Curso" _
                      & "<Ctrl+G>: Ingresar Nota de Examen de Grado"
  RatonNormal
  LstPeriodos.Text = LstPeriodos.List(0)
  LstPeriodos.SetFocus
End Sub

Private Sub Form_Load()
  ConectarAdodc AdoAux
  ConectarAdodc AdoNivel
  ConectarAdodc AdoActas
  ConectarAdodc AdoCurso
  ConectarAdodc AdoNotas
  ConectarAdodc AdoNotasA
  ConectarAdodc AdoCursos
  ConectarAdodc AdoNGrado
  ConectarAdodc AdoEvalua
  ConectarAdodc AdoDetalle
  ConectarAdodc AdoAutorizar
  ConectarAdodc AdoMaterias
  ConectarAdodc AdoPromedios
  ConectarAdodc AdoAsistencia
  ConectarAdodc AdoCatalogoGrado
End Sub

Private Sub LstPeriodos_DblClick()
  SiguienteControl
End Sub

Private Sub LstPeriodos_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub LstPeriodos_LostFocus()
  CadenaParcial = Visualizar_Notas_Periodo(LstPeriodos)
     'MsgBox CadenaParcial
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
  'MsgBox Button.key
   Select Case Button.key
     Case "Salir"
          Unload Me
          Opciones = 0
     Case "Imprimir"
          Imprimir
     Case "Grabar_Notas"
          Grabar_Notas
     Case "Grabar_Notas_Grado"
          Grabar_Notas_Grado
     Case "Grabar_Actas_Grado"
          Grabar_Actas_Grado
     Case "Grabar_Promedio_Finales"
          Grabar_Promedio_Finales
     Case "Mejor_Puntaje"
          Mejor_Puntaje
     Case "Email"
          Email
     Case "Actualizar_Cursos"
          'Actualizar_Malla_Cursos
          Actualiza_Cursos
     Case "Iniciar_Alumnos_Nuevos"
          Iniciar_Alumnos_Nuevos
     Case "Actualiza_Cambio_Curso"
          Actualiza_Cambio_Curso
     Case "Todas_Notas_Emails"
          Todas_Notas_Emails
   End Select
End Sub

Private Sub TVNivel_DblClick()
Dim OpcionNotas As Byte
  SiguienteControl
  NombreDocente = Ninguno
  LblMail.Caption = Ninguno
  Cadena = Mid$(TVNivel.SelectedItem.key, 2, Len(TVNivel.SelectedItem.key) - 1)
  With AdoNivel.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("CodigoE like '" & Cadena & "' ")
       Codigo = .Fields("CodigoE")
       TipoDoc = .Fields("CodMat")
       TipoCta = .Fields("TC")
       CodigoN = .Fields("Materia")
       NombreDocente = .Fields("Dirigente")
       If Len(.Fields("Email")) > 1 Then LblMail.Caption = .Fields("Email")
       If Len(.Fields("Email2")) > 1 And Len(LblMail.Caption) <= 1 Then LblMail.Caption = .Fields("Email2")
   End If
  End With
  Leer_Notas_Parciales TipoDoc, CodigoCuentaSup(Codigo), LstPeriodos
  DGDetalle.Caption = "PROFESOR(A). " & ULCase(NombreDocente) & ", Asignatura. " & TVNivel.SelectedItem & " DEL " & CodigoCuentaSup(Codigo)
  OpcionNotas = 0
  Listar_Notas_Alunmos AdoAutorizar, TipoDoc, CodigoCuentaSup(Codigo), LstPeriodos
 'Text1 = sSQL
  SQLDec = ""
  SelectDataGrid DGDetalle, AdoDetalle, sSQL
  RatonNormal
End Sub

Private Sub TVNivel_KeyDown(KeyCode As Integer, Shift As Integer)
Dim CI_Alum As String
  Keys_Especiales Shift
  CI_Alum = Ninguno
  Cadena = Mid$(TVNivel.SelectedItem.key, 2, Len(TVNivel.SelectedItem.key) - 1)
  With AdoNivel.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("CodigoE like '" & Cadena & "' ")
       Codigo = .Fields("CodigoE")
       Codigo4 = CodigoCuentaSup(Codigo)
       TipoDoc = .Fields("CodMat")
       TipoCta = .Fields("TC")
       Cuenta = .Fields("Materia")
   End If
  End With
  If CtrlDown And KeyCode = vbKeyInsert And TipoCta = "P" Then
    'Insertar Materias al Curso
     SSTabMaterias.TabCaption(1) = TVNivel.SelectedItem
     sSQL = "SELECT (Materia & ' - ' & CodMat) As CMaterias " _
          & "FROM Catalogo_Materias " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "ORDER BY Materia "
     SelectDBList DLMaterias, AdoMaterias, sSQL, "CMaterias"
     DLMaterias.Visible = True
     DLMaterias.SetFocus
     Tipo_Insert_Materias = 1
  End If
  If CtrlDown And KeyCode = vbKeyF9 And TipoCta = "P" Then
    'Insertar Materias de Examen de Grado al Curso
     SSTabMaterias.TabCaption(1) = "MATERIAS ASIGNADAS"
     DGCatalogoGrado.Caption = TVNivel.SelectedItem
     sSQL = "SELECT CodMat,Detalle,CodigoE,Item,Periodo " _
          & "FROM Catalogo_Examen_Grado " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "AND TC = 'M' " _
          & "AND Mid$(CodigoE,1," & Len(Codigo) & ") = '" & Codigo & "' " _
          & "ORDER BY CodigoE "
     SelectDataGrid DGCatalogoGrado, AdoCatalogoGrado, sSQL
     sSQL = "SELECT (CM.Materia & ' - ' & CM.CodMat) As CMaterias " _
          & "FROM Catalogo_Materias As CM,Catalogo_Estudiantil As CE " _
          & "WHERE CM.Item = '" & NumEmpresa & "' " _
          & "AND CM.Periodo = '" & Periodo_Contable & "' " _
          & "AND Mid$(CE.CodigoE,1," & Len(Codigo) & ") = '" & Codigo & "' " _
          & "AND CM.CodMat NOT IN ('998','999') " _
          & "AND CM.CodMat = CE.CodMat " _
          & "AND CM.Item = CE.Item " _
          & "AND CM.Periodo = CE.Periodo " _
          & "ORDER BY CM.Materia "
     SelectDBList DLMaterias, AdoMaterias, sSQL, "CMaterias"
     DLMaterias.Visible = True
     DLMaterias.SetFocus
     Tipo_Insert_Materias = 2
  End If
  If CtrlDown And KeyCode = vbKeyD And TipoCta = "P" Then
    'Actualizar los datos del curso
     Label2.Caption = " CICLO               (" & Codigo & ")"
     SSTabMaterias.Tab = 2
     TxtSeccion = "MATUTINA"
     TxtCiclo = "CICLO"
     TxtParalelo = ""
     TxtBachiller = ""
     TxtEspecialidad = ""
     TxtTitulo = ""
     TxtTipo_Titulo = ""
     TxtCodigo_Titulo = "000000"
     Label13.Caption = Cuenta
     
     sSQL = "SELECT * " _
          & "FROM Catalogo_Cursos " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "AND Curso = '" & Codigo & "' " _
          & "ORDER BY Curso "
     SelectAdodc AdoCurso, sSQL
     With AdoCurso.Recordset
      If .RecordCount > 0 Then
          TxtSeccion = .Fields("Seccion")
          TxtCiclo = .Fields("Ciclo")
          TxtParalelo = .Fields("Paralelo")
          TxtBachiller = .Fields("Bachiller")
          TxtEspecialidad = .Fields("Especialidad")
          TxtTitulo = .Fields("Titulo")
          TxtTipo_Titulo = .Fields("Tipo_Titulo")
          TxtCodigo_Titulo = .Fields("Codigo_Titulo")
      End If
     End With
     TxtBachiller.Visible = False
     TxtEspecialidad.Visible = False
     TxtTitulo.Visible = False
     TxtTipo_Titulo.Visible = False
     TxtCodigo_Titulo.Visible = False
     Label7.Visible = False
     Label8.Visible = False
     Label9.Visible = False
     Label11.Visible = False
     Label12.Visible = False
     If Mid$(Codigo, 1, 1) >= "2" Then
        TxtBachiller.Visible = True
        TxtEspecialidad.Visible = True
        TxtTitulo.Visible = True
        TxtTipo_Titulo.Visible = True
        TxtCodigo_Titulo.Visible = True
        Label7.Visible = True
        Label8.Visible = True
        Label9.Visible = True
        Label11.Visible = True
        Label12.Visible = True
     End If
     TxtSeccion.SetFocus
  End If
  If CtrlDown And KeyCode = vbKeyP Then
     sSQL = "SELECT CE.CodigoE,CE.Detalle,CE.CodMat,C.Cliente As Profesores " _
          & "FROM Catalogo_Estudiantil As CE,Clientes As C " _
          & "WHERE CE.Item = '" & NumEmpresa & "' " _
          & "AND CE.Periodo = '" & Periodo_Contable & "' " _
          & "AND CE.Profesor = C.Codigo " _
          & "ORDER BY C.Codigo "
     SelectData AdoAux, sSQL
     Cuadricula = True
     MensajeEncabData = TVNivel.SelectedItem
     ImprimirAdodc AdoAux, 1, 9, True
  End If
  If KeyCode = vbKeyDelete Then EliminarCta
  If CtrlDown And KeyCode = vbKeyL And TipoCta = "P" Then
     sSQL = "SELECT TN.Codigo,Cliente As Alumno,Direccion As Curso " _
          & "FROM Clientes As C,Trans_Notas As TN " _
          & "WHERE TN.CodE = '" & Codigo & "' " _
          & "AND TN.Item = '" & NumEmpresa & "' " _
          & "AND TN.Periodo = '" & Periodo_Contable & "' " _
          & "AND C.Codigo = TN.Codigo " _
          & "GROUP BY TN.Codigo,Cliente,Direccion,Sexo " _
          & "ORDER BY Sexo DESC,Cliente "
     SelectDataGrid DGDetalle, AdoDetalle, sSQL
     DGDetalle.Caption = "ALUMNOS DEL: " & TVNivel.SelectedItem
  End If
 'Insertar las notas del los alumnos
  If CtrlDown And KeyCode = vbKeyL And TipoCta = "M" Then
     Leer_Notas_Parciales TipoDoc, CodigoCuentaSup(Codigo), LstPeriodos
     DGDetalle.Caption = "ALUMNOS DEL: " & TVNivel.SelectedItem & " DEL " & CodigoCuentaSup(Codigo)
     Opcion = 0
'''     If OpcPQBim1.value Then Opcion = 1
'''     If OpcPQ.value Then Opcion = 2
'''     If OpcSQBim1.value Then Opcion = 3
'''     If OpcSQ.value Then Opcion = 4
     'MsgBox Opcion
     Listar_Notas_Alunmos AdoAutorizar, TipoDoc, CodigoCuentaSup(Codigo), LstPeriodos
     SQLDec = ""
     SelectDataGrid DGDetalle, AdoDetalle, sSQL
  End If
 'Insertar las notas de Grado del los alumnos
  If CtrlDown And KeyCode = vbKeyG And TipoCta = "M" Then
     sSQL = "DELETE * " _
          & "FROM Asiento_NG " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND CodMat = '" & TipoDoc & "' " _
          & "AND CodigoU = '" & CodigoUsuario & "' "
     ConectarAdoExecute sSQL
     'CodigoCuentaSup(.Fields("CodE"))
     sSQL = "SELECT CodMat,Detalle,CodigoE,Item,Periodo " _
          & "FROM Catalogo_Examen_Grado " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "AND TC = 'M' " _
          & "AND Mid$(CodigoE,1," & Len(Codigo4) & ") = '" & Codigo4 & "' " _
          & "AND CodMat = '" & TipoDoc & "' " _
          & "ORDER BY CodigoE "
     SelectDataGrid DGCatalogoGrado, AdoCatalogoGrado, sSQL
     If AdoCatalogoGrado.Recordset.RecordCount > 0 Then
        sSQL = "SELECT C.Codigo,C.Cliente " _
             & "FROM Clientes As C,Clientes_Matriculas As CM " _
             & "WHERE CM.Grupo_No = '" & Codigo4 & "' " _
             & "AND C.Codigo = CM.Codigo " _
             & "AND CM.Item = '" & NumEmpresa & "' " _
             & "AND CM.Periodo = '" & Periodo_Contable & "' " _
             & "ORDER BY C.Cliente "
        SelectData AdoAux, sSQL
        With AdoAux.Recordset
         If .RecordCount > 0 Then
             Contador = 1
             Do While Not .EOF
                Real1 = 0
                sSQL = "SELECT * " _
                     & "FROM Trans_Notas_Grado " _
                     & "WHERE Item = '" & NumEmpresa & "' " _
                     & "AND Periodo = '" & Periodo_Contable & "' " _
                     & "AND Codigo = '" & .Fields("Codigo") & "' " _
                     & "AND Mid$(CodE,1," & Len(Codigo4) & ") = '" & Codigo4 & "' " _
                     & "AND CodMat = '" & TipoDoc & "' "
                SelectAdodc AdoNGrado, sSQL
                If AdoNGrado.Recordset.RecordCount > 0 Then
                   Real1 = AdoNGrado.Recordset.Fields("Examen")
                End If
                SetAdoAddNew "Asiento_NG"
                SetAdoFields "Id_No", Contador
                SetAdoFields "Codigo", .Fields("Codigo")
                SetAdoFields "Alumno", .Fields("Cliente")
                SetAdoFields "CodMat", TipoDoc
                SetAdoFields "CodE", Codigo4
                SetAdoFields "Examen", Real1
                SetAdoFields "Item", NumEmpresa
                SetAdoUpdate
                Contador = Contador + 1
               .MoveNext
             Loop
         End If
        End With
        sSQL = "SELECT Id_No,Alumno,Examen,CodMat,Item,Codigo,CodE,CodigoU " _
             & "FROM Asiento_NG " _
             & "WHERE Item = '" & NumEmpresa & "' " _
             & "AND CodMat = '" & TipoDoc & "' " _
             & "AND CodigoU = '" & CodigoUsuario & "' " _
             & "ORDER BY Id_No "
        SQLDec = ""
        SelectDataGrid DGDetalle, AdoDetalle, sSQL
        DGDetalle.Caption = "ALUMNOS DEL: " & TVNivel.SelectedItem & " DEL " & CodigoCuentaSup(Codigo)
     Else
        MsgBox "Esta Materia no esta asignada para examen de Grado"
     End If
  End If
  If CtrlDown And KeyCode = vbKeyA And TipoCta = "P" Then
    'MsgBox "Ctrl+A y P: " & Codigo
     If Mid$(Codigo, 1, 4) = "3.03" Or Mid$(Codigo, 1, 4) = "5.03" Then
       'Verificamos el promedio de los alumnnos de 1 a 5 año
        sSQL = "SELECT Codigo,PromFinal " _
             & "FROM Trans_Promedios " _
             & "WHERE Item = '" & NumEmpresa & "' " _
             & "AND Periodo = '" & Periodo_Contable & "' " _
             & "AND CodE = '" & Codigo & "' " _
             & "AND PromFinal > 0 " _
             & "ORDER BY Codigo "
        SelectAdodc AdoPromedios, sSQL
       'Promedio de Notas del Examen de Grado
        sSQL = "SELECT Codigo, SUM(Examen) AS TExamen, COUNT(Codigo) AS CExamen " _
             & "FROM Trans_Notas_Grado " _
             & "WHERE Item = '" & NumEmpresa & "' " _
             & "AND Periodo = '" & Periodo_Contable & "' " _
             & "AND Mid$(CodE,1," & Len(Codigo) & ") = '" & Codigo & "' " _
             & "GROUP BY Codigo " _
             & "ORDER BY Codigo "
        SelectAdodc AdoNGrado, sSQL
       'Promedio de Notas del 6to. año
       '& "AND TM.P <> 0 "
        sSQL = "SELECT TN.Codigo,AVG(TN.PromFinal) As PromF " _
             & "FROM Trans_Notas As TN,Catalogo_Materias As TM " _
             & "WHERE TN.Item = '" & NumEmpresa & "' " _
             & "AND TN.Periodo = '" & Periodo_Contable & "' " _
             & "AND TN.CodE = '" & Codigo & "' " _
             & "AND TN.PromFinal > 0 " _
             & "AND TN.CodMatP = '.' " _
             & "AND TN.CodMat NOT IN ('998','999') " _
             & "AND TM.C = " & Val(adFalse) & " " _
             & "AND TN.CodMat = TM.CodMat " _
             & "AND TN.Item = TM.Item " _
             & "AND TN.Periodo = TM.Periodo " _
             & "GROUP BY TN.Codigo " _
             & "ORDER BY TN.Codigo "
        SelectAdodc AdoEvalua, sSQL
'''        MsgBox AdoPromedios.Recordset.RecordCount & vbCrLf _
'''               & AdoNGrado.Recordset.RecordCount & vbCrLf _
'''               & AdoEvalua.Recordset.RecordCount & vbCrLf & "...."
        sSQL = "DELETE * " _
             & "FROM Asiento_A " _
             & "WHERE Item = '" & NumEmpresa & "' " _
             & "AND CodigoU = '" & CodigoUsuario & "' "
        ConectarAdoExecute sSQL
        ListarActasAlunmos
        SelectAdodc AdoDetalle, sSQL
        DGDetalle.Caption = "ALUMNOS DEL: " & TVNivel.SelectedItem
       'If AdoDetalle.Recordset.RecordCount <= 0 Then
        sSQL = "SELECT C.Codigo,C.Cliente As Alumno,C.Direccion As Curso " _
             & "FROM Clientes As C,Clientes_Matriculas As CM " _
             & "WHERE CM.Grupo_No = '" & Codigo & "' " _
             & "AND C.Codigo = CM.Codigo " _
             & "AND CM.Item = '" & NumEmpresa & "' " _
             & "AND CM.Periodo = '" & Periodo_Contable & "' " _
             & "ORDER BY C.Cliente "
        SelectData AdoAux, sSQL
        'MsgBox AdoAux.Recordset.RecordCount & vbCrLf & AdoDetalle.Recordset.RecordCount
        If AdoAux.Recordset.RecordCount <> AdoDetalle.Recordset.RecordCount Then
           'MsgBox AdoAux.Recordset.RecordCount & vbCrLf & AdoDetalle.Recordset.RecordCount
           Do While Not AdoAux.Recordset.EOF
              Si_No = False
              Codigo1 = AdoAux.Recordset.Fields("Codigo")
              If AdoDetalle.Recordset.RecordCount > 0 Then
                 AdoDetalle.Recordset.MoveFirst
                 AdoDetalle.Recordset.Find ("Codigo = '" & Codigo1 & "' ")
                 If Not AdoDetalle.Recordset.EOF Then Si_No = True
              End If
             'MsgBox Si_No
              If Si_No = False Then
                'MsgBox Mid$(Codigo, Len(Codigo) - 1, 2)
                 SetAdoAddNew "Trans_Actas"
                 SetAdoFields "Id_No", 0
                 SetAdoFields "Codigo", Codigo1
                 SetAdoFields "Notas", 0
                 SetAdoFields "Trabajo", 0
                 SetAdoFields "Investigacion", 0
                 SetAdoFields "Evaluacion", 0
                 SetAdoFields "Periodo", Periodo_Contable
                 SetAdoFields "Item", NumEmpresa
                 SetAdoUpdate
              End If
              AdoAux.Recordset.MoveNext
           Loop
        End If
     'End If
     ListarActasAlunmos
     SelectAdodc AdoAux, sSQL
     Contador = 1
     If AdoAux.Recordset.RecordCount > 0 Then
        Do While Not AdoAux.Recordset.EOF
           CodigoCli = AdoAux.Recordset.Fields("Codigo")
           Total = 0
           Saldo = 0
           Diferencia = 0
          'Promedio de 1 a 5 año
           If AdoPromedios.Recordset.RecordCount > 0 Then
              AdoPromedios.Recordset.MoveFirst
              AdoPromedios.Recordset.Find ("Codigo = '" & CodigoCli & "' ")
              If Not AdoPromedios.Recordset.EOF Then
                 Diferencia = AdoPromedios.Recordset.Fields("PromFinal")
              End If
           End If
          'Promedio de Notas del Examen de Grado
           If AdoNGrado.Recordset.RecordCount > 0 Then
              AdoNGrado.Recordset.MoveFirst
              AdoNGrado.Recordset.Find ("Codigo = '" & CodigoCli & "' ")
              If Not AdoNGrado.Recordset.EOF Then
                 Cantidad = AdoNGrado.Recordset.Fields("CExamen")
                 If Cantidad <= 0 Then Cantidad = 1
                 Total = AdoNGrado.Recordset.Fields("TExamen") / Cantidad
              End If
           End If
          'Promedio de las Notas del 6 año
           If AdoEvalua.Recordset.RecordCount > 0 Then
              AdoEvalua.Recordset.MoveFirst
              AdoEvalua.Recordset.Find ("Codigo = '" & CodigoCli & "' ")
              If Not AdoEvalua.Recordset.EOF Then Saldo = AdoEvalua.Recordset.Fields("PromF")
           End If
           'If Total <= 0 Then Total = AdoAux.Recordset.Fields("Evaluacion")
           'If Saldo <= 0 Then Saldo = AdoAux.Recordset.Fields("Investigacion")
           'If Diferencia <= 0 Then Diferencia = AdoAux.Recordset.Fields("Notas")
''''           MsgBox Total & vbCrLf _
''''                  & Saldo & vbCrLf _
''''                  & Diferencia & vbCrLf

           SetAdoAddNew "Asiento_A"
           SetAdoFields "Id_No", AdoAux.Recordset.Fields("Id_No")
           SetAdoFields "Codigo", AdoAux.Recordset.Fields("Codigo")
           SetAdoFields "Alumno", AdoAux.Recordset.Fields("Alumno")
           SetAdoFields "Trabajo", AdoAux.Recordset.Fields("Trabajo")
           SetAdoFields "Cedula", AdoAux.Recordset.Fields("CI")
           SetAdoFields "Notas", Diferencia          'Promedio de 1 a 5 año
           SetAdoFields "Investigacion", Saldo       'Promedio de Notas 6 año
           SetAdoFields "Evaluacion", Total          'Promedio de Notas Examen de Grado
           SetAdoFields "Item", NumEmpresa
           SetAdoFields "CodigoU", CodigoUsuario
           SetAdoUpdate
           Contador = Contador + 1
           AdoAux.Recordset.MoveNext
        Loop
     End If
     ListarActasAlunmos True
     'MsgBox SQLDec
     SelectDataGrid DGDetalle, AdoDetalle, sSQL, SQLDec
     Else
        MsgBox "Este paralelo no es valido"
     End If
  End If
  If AltDown And KeyCode = vbKeyP And TipoCta = "P" Then
     'MsgBox "Ctrl+A y P: " & Codigo
        sSQL = "SELECT Codigo,AVG(Nota_Grado)As NotaG " _
             & "FROM Trans_Notas " _
             & "WHERE Item = '" & NumEmpresa & "' " _
             & "AND Periodo = '" & Periodo_Contable & "' " _
             & "AND CodE = '" & Codigo & "' " _
             & "AND Nota_Grado > 0 " _
             & "GROUP BY Codigo " _
             & "ORDER BY Codigo "
        SelectAdodc AdoNGrado, sSQL
        
        sSQL = "SELECT TN.Codigo,AVG(TN.PromFinal) As PromF " _
             & "FROM Trans_Notas As TN,Catalogo_Materias As TM " _
             & "WHERE TN.Item = '" & NumEmpresa & "' " _
             & "AND TN.Periodo = '" & Periodo_Contable & "' " _
             & "AND TN.CodE = '" & Codigo & "' " _
             & "AND TN.PromFinal > 0 " _
             & "AND TM.P <> 0 " _
             & "AND TN.CodMatP = '" & Ninguno & "' " _
             & "AND TN.CodMat = TM.CodMat " _
             & "AND TN.Item = TM.Item " _
             & "AND TN.Periodo = TM.Periodo " _
             & "GROUP BY TN.Codigo " _
             & "ORDER BY TN.Codigo "
        SelectAdodc AdoEvalua, sSQL
     
     sSQL = "DELETE * " _
          & "FROM Asiento_A " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND CodigoU = '" & CodigoUsuario & "' "
     ConectarAdoExecute sSQL
     ListarPromediosAlunmos
     SelectAdodc AdoDetalle, sSQL
     DGDetalle.Caption = "ALUMNOS DEL: " & TVNivel.SelectedItem
     'If AdoDetalle.Recordset.RecordCount <= 0 Then
        sSQL = "SELECT Codigo,Cliente As Alumno,Direccion As Curso " _
             & "FROM Clientes " _
             & "WHERE Grupo = '" & Codigo & "' " _
             & "ORDER BY Cliente "
        SelectData AdoAux, sSQL
        'MsgBox AdoAux.Recordset.RecordCount & vbCrLf & AdoDetalle.Recordset.RecordCount
        If AdoAux.Recordset.RecordCount <> AdoDetalle.Recordset.RecordCount Then
          'MsgBox AdoAux.Recordset.RecordCount & vbCrLf & AdoDetalle.Recordset.RecordCount
           Do While Not AdoAux.Recordset.EOF
              Si_No = False
              Codigo1 = AdoAux.Recordset.Fields("Codigo")
              If AdoDetalle.Recordset.RecordCount > 0 Then
                 AdoDetalle.Recordset.MoveFirst
                 AdoDetalle.Recordset.Find ("Codigo = '" & Codigo1 & "' ")
                 If Not AdoDetalle.Recordset.EOF Then Si_No = True
              End If
             'MsgBox Si_No
              If Si_No = False Then
                'MsgBox Codigo & " - " & Mid$(Codigo, Len(Codigo) - 1, 2)
                 SetAdoAddNew "Trans_Promedios"
                 SetAdoFields "Id_No", Val(Mid$(Codigo, Len(Codigo) - 1, 2))
                 SetAdoFields "Codigo", Codigo1
                 SetAdoFields "Periodo", Periodo_Contable
                 SetAdoFields "Item", NumEmpresa
                 SetAdoUpdate
              End If
              AdoAux.Recordset.MoveNext
           Loop
        End If
     'End If
     ListarPromediosAlunmos
     SelectAdodc AdoAux, sSQL
     Contador = 0
     If AdoAux.Recordset.RecordCount > 0 Then
        Do While Not AdoAux.Recordset.EOF
           Contador = Contador + 1
           CodigoCli = AdoAux.Recordset.Fields("Codigo")
           Total = 0: Saldo = 0
           If AdoNGrado.Recordset.RecordCount > 0 Then
              AdoNGrado.Recordset.MoveFirst
              AdoNGrado.Recordset.Find ("Codigo = '" & CodigoCli & "' ")
              If Not AdoNGrado.Recordset.EOF Then Total = AdoNGrado.Recordset.Fields("NotaG")
           End If
           If AdoEvalua.Recordset.RecordCount > 0 Then
              AdoEvalua.Recordset.MoveFirst
              AdoEvalua.Recordset.Find ("Codigo = '" & CodigoCli & "' ")
              If Not AdoEvalua.Recordset.EOF Then Saldo = AdoEvalua.Recordset.Fields("PromF")
           End If
           'MsgBox Total
           Total = AdoAux.Recordset.Fields("N_1er")
           Total = Total + AdoAux.Recordset.Fields("N_2do")
           Total = Total + AdoAux.Recordset.Fields("N_3er")
           Total = Total + AdoAux.Recordset.Fields("N_4to")
           Total = Total + AdoAux.Recordset.Fields("N_5to")
          'MsgBox CByte(Contador) & " - " & AdoAux.Recordset.Fields("Alumno")
           SetAdoAddNew "Asiento_A"
           SetAdoFields "Id_No", CByte(Contador)
           SetAdoFields "Codigo", AdoAux.Recordset.Fields("Codigo")
           SetAdoFields "Alumno", AdoAux.Recordset.Fields("Alumno")
           SetAdoFields "N_1er", AdoAux.Recordset.Fields("N_1er")
           SetAdoFields "N_2do", AdoAux.Recordset.Fields("N_2do")
           SetAdoFields "N_3er", AdoAux.Recordset.Fields("N_3er")
           SetAdoFields "N_4to", AdoAux.Recordset.Fields("N_4to")
           SetAdoFields "N_5to", AdoAux.Recordset.Fields("N_5to")
           SetAdoFields "Total", Total
           SetAdoFields "Promedio", Redondear(Total / 5, 3)
           SetAdoFields "Item", NumEmpresa
           SetAdoFields "CodigoU", CodigoUsuario
           SetAdoUpdate
           
           AdoAux.Recordset.MoveNext
        Loop
     ListarPromediosAlunmos True
    'MsgBox SQLDec
     SelectDataGrid DGDetalle, AdoDetalle, sSQL, SQLDec
     Else
        MsgBox "Datos No permitidos"
     End If
  End If
  If CtrlDown And KeyCode = vbKeyF11 And TipoCta = "N" Then
     sSQL = "SELECT TN.CodE As Curso,Cliente As Alumno," _
          & "TN.N_1er,TN.N_2do,TN.N_3er,TN.N_4to,TN.N_5to,TN.Total," _
          & "TN.PromFinal As Promedio " _
          & "FROM Clientes As C,Trans_Promedios As TN,Clientes_Matriculas As CM " _
          & "WHERE Mid$(TN.CodE,1," & Len(Codigo) & ") = '" & Codigo & "' " _
          & "AND TN.Item = '" & NumEmpresa & "' " _
          & "AND TN.Periodo = '" & Periodo_Contable & "' " _
          & "AND C.Codigo = TN.Codigo " _
          & "AND C.Codigo = CM.Codigo " _
          & "AND TN.Item = CM.Item " _
          & "AND TN.Periodo = CM.Periodo " _
          & "ORDER BY TN.PromFinal DESC,TN.CodE,C.Cliente "
     SQLDec = "Promedio 3|."
     SelectDataGrid DGDetalle, AdoDetalle, sSQL, SQLDec
  End If
'''  If CtrlDown And KeyCode = vbKeyF11 And TipoCta = "C" Then
'''    'Notas por Materia de Profesor
'''     If OpcPQ.value Then
'''        sSQL = "SELECT C.Cliente As Estudiante,TN.CodE As Curso,(SUM(TN.PromPQ)/COUNT(TN.CodE)) As Promedio "
'''     ElseIf OpcSQ.value Then
'''        sSQL = "SELECT C.Cliente As Estudiante,TN.CodE As Curso,(SUM(TN.PromSQ)/COUNT(TN.CodE)) As Promedio "
'''     Else
'''        sSQL = "SELECT C.Cliente As Estudiante,TN.CodE As Curso,(SUM(TN.PromFinal)/COUNT(TN.CodE)) As Promedio "
'''     End If
'''     sSQL = sSQL & "FROM Trans_Notas As TN,Clientes As C " _
'''          & "WHERE TN.Item = '" & NumEmpresa & "' " _
'''          & "AND TN.Periodo = '" & Periodo_Contable & "' " _
'''          & "AND Mid$(TN.CodE,1," & Len(Codigo) & ") = '" & Codigo & "' " _
'''          & "AND TN.CodMat <> '999' " _
'''          & "AND TN.Codigo = C.Codigo " _
'''          & "GROUP BY C.Cliente,TN.CodE "
'''     If OpcPQ.value Then
'''        sSQL = sSQL & "HAVING (SUM(TN.PromPQ)/COUNT(TN.CodE)) >= 19 "
'''     ElseIf OpcSQ.value Then
'''        sSQL = sSQL & "HAVING (SUM(TN.PromSQ)/COUNT(TN.CodE)) >= 19 "
'''     Else
'''        sSQL = sSQL & "HAVING (SUM(TN.PromFinal)/COUNT(TN.CodE)) >= 19 "
'''     End If
'''     sSQL = sSQL & "ORDER BY Promedio Desc,TN.CodE,C.Cliente "
'''     SQLDec = "Promedio 2 |."
'''     SelectDataGrid DGDetalle, AdoDetalle, sSQL, SQLDec
'''  End If
  
'''  If CtrlDown And KeyCode = vbKeyF12 And TipoCta = "C" Then
'''    'Notas por Materia de Profesor
'''     If OpcPQ.value Then
'''        sSQL = "SELECT C.Cliente As Estudiante,TN.CodE As Curso,(SUM(TN.PromPQ)/COUNT(TN.CodE)) As Promedio "
'''     ElseIf OpcSQ.value Then
'''        sSQL = "SELECT C.Cliente As Estudiante,TN.CodE As Curso,(SUM(TN.PromSQ)/COUNT(TN.CodE)) As Promedio "
'''     Else
'''        sSQL = "SELECT C.Cliente As Estudiante,TN.CodE As Curso,(SUM(TN.PromFinal)/COUNT(TN.CodE)) As Promedio "
'''     End If
'''     sSQL = sSQL & "FROM Trans_Notas As TN,Clientes As C " _
'''          & "WHERE TN.Item = '" & NumEmpresa & "' " _
'''          & "AND TN.Periodo = '" & Periodo_Contable & "' " _
'''          & "AND Mid$(TN.CodE,1," & Len(Codigo) & ") = '" & Codigo & "' " _
'''          & "AND TN.CodMat <> '999' " _
'''          & "AND TN.Codigo = C.Codigo " _
'''          & "GROUP BY C.Cliente,TN.CodE "
'''     If OpcPQ.value Then
'''        sSQL = sSQL & "HAVING (SUM(TN.PromPQ)/COUNT(TN.CodE)) >= 19 "
'''     ElseIf OpcSQ.value Then
'''        sSQL = sSQL & "HAVING (SUM(TN.PromSQ)/COUNT(TN.CodE)) >= 19 "
'''     Else
'''        sSQL = sSQL & "HAVING (SUM(TN.PromFinal)/COUNT(TN.CodE)) >= 19 "
'''     End If
'''     sSQL = sSQL & "ORDER BY TN.CodE,Promedio Desc,C.Cliente "
'''     SQLDec = "Promedio 2 |."
'''     SelectDataGrid DGDetalle, AdoDetalle, sSQL, SQLDec
'''  End If
  
  'PresionoEnter KeyCode
End Sub

Private Sub TxtBachiller_GotFocus()
   MarcarTexto TxtBachiller
End Sub

Private Sub TxtBachiller_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtCiclo_GotFocus()
  MarcarTexto TxtCiclo
End Sub

Private Sub TxtCiclo_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtCodigo_Titulo_GotFocus()
  MarcarTexto TxtCodigo_Titulo
End Sub

Private Sub TxtCodigo_Titulo_KeyDown(KeyCode As Integer, Shift As Integer)
    PresionoEnter KeyCode
End Sub

Private Sub TxtCodigo_Titulo_LostFocus()
   TextoValido TxtCodigo_Titulo, True, , 0
   TxtCodigo_Titulo = Format(Val(TxtCodigo_Titulo), "000000")
End Sub

Private Sub TxtEspecialidad_GotFocus()
  MarcarTexto TxtEspecialidad
End Sub

Private Sub TxtEspecialidad_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtParalelo_GotFocus()
   MarcarTexto TxtParalelo
End Sub

Private Sub TxtParalelo_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtSeccion_GotFocus()
  MarcarTexto TxtSeccion
End Sub

Private Sub TxtSeccion_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtTipo_Titulo_Change()
   MarcarTexto TxtTipo_Titulo
End Sub

Private Sub TxtTipo_Titulo_GotFocus()
   MarcarTexto TxtTipo_Titulo
End Sub

Private Sub TxtTipo_Titulo_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtTitulo_GotFocus()
  MarcarTexto TxtTitulo
End Sub

Private Sub TxtTitulo_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Public Sub Eliminar_Notas_Cero(Tabla_Eliminar As String)
Dim SumaCampos As String
  SumaCampos = ""
  sSQL = "SELECT * " _
       & "FROM " & Tabla_Eliminar & " " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND Codigo = '.' "
  SelectAdodc AdoNivel, sSQL
  With AdoNivel.Recordset
       For J = 0 To .Fields.Count - 1
           If .Fields(J).Type = TadCurrency Then SumaCampos = SumaCampos & .Fields(J).Name & " + "
       Next J
  End With
  SumaCampos = Trim(SumaCampos)
  SumaCampos = Mid$(SumaCampos, 1, Len(SumaCampos) - 1)
  SumaCampos = "(" & Trim(SumaCampos) & ") <= 0 "
  sSQL = "DELETE * " _
       & "FROM " & Tabla_Eliminar & " " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND " & SumaCampos & " "
  ConectarAdoExecute sSQL
End Sub

Public Function Enviar_Notas_por_Materia(Curso As String, _
                                         Materia As String, _
                                         CodigoMateria As String, _
                                         Profesor As String) As String
Dim NFila As Integer
Dim RutaGeneraFile As String
Dim Paralelo As String
Dim NotaExa As Boolean
Dim oExcel As Object
Dim oBook As Object
Dim oSheet As Object

  RatonReloj
  DGDetalle.Visible = False
  NotaExa = False
 'Start a new workbook in Excel
  Set oExcel = CreateObject("Excel.Application")
  Set oBook = oExcel.workbooks.Add
 'Add data to cells of the first worksheet in the new workbook
  Set oSheet = oBook.Worksheets(1)
  
  Profesor = Replace(Profesor, ".", " ")
  Profesor = Replace(Profesor, "/", " ")
  Profesor = Replace(Profesor, ":", "")
  Profesor = UCase(Trim(Profesor))
  
  Materia = Replace(Materia, ".", " ")
  Materia = Replace(Materia, "/", " ")
  Materia = Replace(Materia, ":", "")
  Materia = UCase(Trim(Materia))
  
  Paralelo = Trim(Mid$(CambioCodigoCtaSup(Curso), 3, 10))
  Paralelo = Trim(Replace(Paralelo, ".", "-"))
  
  Contador = 0
  With AdoDetalle.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       NFila = 1
       RutaGeneraFile = RutaSysBases & "\Emails\Notas del " & Paralelo & " " & Materia & " - " & Profesor & ".xls"
       If Dir(RutaGeneraFile) <> "" Then Kill RutaGeneraFile
       Codigo1 = Leer_Datos_del_Curso(Curso)
       Codigo1 = Dato_Curso.Especialidad
       oSheet.Range("A1").value = "ACTA DE CALIFICACIONES POR MATERIA"
       oSheet.Range("A2").value = UCase(LstPeriodos.Text)
       oSheet.Range("A3").value = "Lcdo(a). " & ULCase(NombreDocente)
       oSheet.Range("A4").value = "Curso: " & Codigo & ", Materia: " & TVNivel.SelectedItem
       For IE = 0 To AdoDetalle.Recordset.Fields.Count - 1
           oSheet.Range(Chr(65 + IE) & "5").value = .Fields(IE).Name
       Next IE
       NFila = 5
       Do While Not .EOF
          NFila = NFila + 1
          For IE = 0 To AdoDetalle.Recordset.Fields.Count - 1
              oSheet.Range(Chr(65 + IE) & CStr(NFila)).value = .Fields(IE)
          Next IE
         .MoveNext
       Loop
       NFila = NFila + 2
       oSheet.Range("A" & CStr(NFila)).value = "ENVIADO POR SECRETARÍA GENERAL EL DÍA " & FechaSistema & " A LAS " & Format(Time, "hh:mm:ss")
   End If
  End With
  DGDetalle.Visible = True
  RatonNormal
 'Save the Workbook and Quit Excel
  oBook.SaveAs RutaGeneraFile
  oExcel.Quit
  Enviar_Notas_por_Materia = RutaGeneraFile
End Function

Public Sub Todas_Notas_Emails()
Dim EsperarMail As Integer
Dim ContMat As Long
  RatonReloj
  NombreDocente = Ninguno
  LblMail.Caption = Ninguno
  With AdoNivel.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       Do While Not .EOF
          If .Fields("TC") = "M" And Len(.Fields("Dirigente")) > 1 Then
              Codigo = .Fields("CodigoE")
              TipoDoc = .Fields("CodMat")
              TipoCta = .Fields("TC")
              CodigoN = .Fields("Materia")
              NombreDocente = .Fields("Dirigente")
              If Len(.Fields("Email")) > 1 Then LblMail.Caption = .Fields("Email")
              If Len(.Fields("Email2")) > 1 And Len(LblMail.Caption) <= 1 Then LblMail.Caption = .Fields("Email2")
              OpcionNotas = 0
              Leer_Notas_Parciales TipoDoc, CodigoCuentaSup(Codigo), LstPeriodos
              Listar_Notas_Alunmos AdoAutorizar, TipoDoc, CodigoCuentaSup(Codigo), LstPeriodos
              SQLDec = ""
              SelectDataGrid DGDetalle, AdoDetalle, sSQL
              If Len(LblMail.Caption) > 1 Then
                 FEducativo.Caption = Codigo & ": " & NombreDocente & " <" & CodigoN & ">"
                 DGDetalle.Visible = False
                 TMail.Adjunto = Enviar_Notas_por_Materia(Codigo, CodigoN, TipoDoc, NombreDocente)
                 TMail.Asunto = "Solicitud de Envio de Notas"
                 TMail.Mensaje = "Estimado Docente, descargue el archivo, si aparece un mensaje que dice: " & vbCrLf _
                      & "'Desea abrir el archivo ahora?', Presionar el boton SI." & vbCrLf & vbCrLf _
                      & "Si presenta un mensaje como este: Vista Protegida, Presionar el Boton que dice:" & vbCrLf _
                      & "Habilitar edición"
                 TMail.para = LblMail.Caption
                 FEnviarCorreos.Show 1
                 DGDetalle.Visible = True
              End If
          End If
         .MoveNext
       Loop
   End If
  End With
  RatonNormal
  FEducativo.Caption = "CATALOGO ESTUDIANTIL"
  MsgBox "PROCESO TERMINADO"
End Sub
