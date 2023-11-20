VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.4#0"; "Comctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FLibretas 
   Caption         =   "CATALOGO ESTUDIANTIL"
   ClientHeight    =   7575
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13335
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7575
   ScaleWidth      =   13335
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   26
      Top             =   0
      Width           =   13335
      _ExtentX        =   23521
      _ExtentY        =   1164
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "ImgLstMenu"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   21
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Salir"
            Object.ToolTipText     =   "Salir del proceso"
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
            Key             =   "Matriculado"
            Object.ToolTipText     =   "Listar los Matriculados"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "ListadoAlumMatDP"
            Object.ToolTipText     =   "Alumnos Matriculados para la Dirección Provincial"
            Object.Tag             =   ""
            ImageIndex      =   16
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Representante"
            Object.ToolTipText     =   "Listar Representante de Alumnos"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Aptitud"
            Object.ToolTipText     =   "Aptitud"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Promocion"
            Object.ToolTipText     =   "Listar Promoción"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Libretas"
            Object.ToolTipText     =   "Imprimir Libretas del Curso"
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Actas"
            Object.ToolTipText     =   "Listar Acta de Grado"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "SegundaPg"
            Object.ToolTipText     =   "Listar segunda página del Acta de Grado"
            Object.Tag             =   ""
            ImageIndex      =   9
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Carnet"
            Object.ToolTipText     =   "Lista de Carnet"
            Object.Tag             =   ""
            ImageIndex      =   10
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "CuadroProm"
            Object.ToolTipText     =   "Cuadro Final Dirección Provincial"
            Object.Tag             =   ""
            ImageIndex      =   11
         EndProperty
         BeginProperty Button13 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Certificados"
            Object.ToolTipText     =   "Listar Cetificado de Matricula"
            Object.Tag             =   ""
            ImageIndex      =   12
         EndProperty
         BeginProperty Button14 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "SolExaGrado"
            Object.ToolTipText     =   "Imprime Solicitud para rendir examen de grado"
            Object.Tag             =   ""
            ImageIndex      =   13
         EndProperty
         BeginProperty Button15 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "AprobarExamGrado"
            Object.ToolTipText     =   "Aprobacion Exam. de Grado"
            Object.Tag             =   ""
            ImageIndex      =   14
         EndProperty
         BeginProperty Button16 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "NotasExamGrado"
            Object.ToolTipText     =   "Notas de Examen de Grado"
            Object.Tag             =   ""
            ImageIndex      =   15
         EndProperty
         BeginProperty Button17 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Supletorio"
            Object.ToolTipText     =   "Presentar Cuadro de Supletorios"
            Object.Tag             =   ""
            ImageIndex      =   18
         EndProperty
         BeginProperty Button18 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Recalcular"
            Object.ToolTipText     =   "Recalcula los totales de las notas ingresadas"
            Object.Tag             =   ""
            ImageIndex      =   17
         EndProperty
         BeginProperty Button19 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Lista_Estudiantes"
            Object.ToolTipText     =   "Listar Nomina de Estudiantes"
            Object.Tag             =   ""
            ImageIndex      =   19
         EndProperty
         BeginProperty Button20 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Nomina_Representante"
            Object.ToolTipText     =   "Nomina de Estudiantes con Representantes"
            Object.Tag             =   ""
            ImageIndex      =   20
         EndProperty
         BeginProperty Button21 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Nomina_Representante_Email"
            Object.ToolTipText     =   "Nomina de Estudiantes con Representantes con Email"
            Object.Tag             =   ""
            ImageIndex      =   21
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin ComctlLib.ListView LstVAlumnos 
      Height          =   1905
      Left            =   105
      TabIndex        =   25
      Top             =   5040
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   3360
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CheckBox CheqFirma 
      Caption         =   "Con Firma"
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
      Left            =   19005
      TabIndex        =   22
      Top             =   1155
      Value           =   1  'Checked
      Width           =   1170
   End
   Begin VB.TextBox TxtObservacion 
      BackColor       =   &H00FFC0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   330
      Left            =   15435
      MaxLength       =   20
      TabIndex        =   21
      Text            =   "OBSERVACION"
      Top             =   1155
      Width           =   3480
   End
   Begin VB.TextBox TxtTitulo 
      BackColor       =   &H00FFC0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   330
      Left            =   9030
      MaxLength       =   80
      TabIndex        =   20
      Text            =   "NOMINA DE ESTUDIANTES"
      Top             =   1155
      Width           =   5055
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
      Height          =   330
      Left            =   7350
      TabIndex        =   16
      Top             =   735
      Width           =   330
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
      Width           =   7575
   End
   Begin VB.PictureBox PictTotal 
      BackColor       =   &H00C0FFFF&
      Height          =   330
      Left            =   15540
      ScaleHeight     =   270
      ScaleWidth      =   4575
      TabIndex        =   6
      Top             =   735
      Width           =   4635
   End
   Begin MSAdodcLib.Adodc AdoAlumnos 
      Height          =   330
      Left            =   210
      Top             =   1575
      Visible         =   0   'False
      Width           =   1905
      _ExtentX        =   3360
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
      Caption         =   "Alumnos"
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
      Left            =   2100
      Top             =   1260
      Visible         =   0   'False
      Width           =   1905
      _ExtentX        =   3360
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
   Begin MSAdodcLib.Adodc AdoLibreta 
      Height          =   330
      Left            =   210
      Top             =   1260
      Visible         =   0   'False
      Width           =   1905
      _ExtentX        =   3360
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
      Caption         =   "Libreta"
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
   Begin MSAdodcLib.Adodc AdoPlantel 
      Height          =   330
      Left            =   2100
      Top             =   1575
      Visible         =   0   'False
      Width           =   1905
      _ExtentX        =   3360
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
      Caption         =   "Plantel"
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
      Left            =   210
      Top             =   1890
      Visible         =   0   'False
      Width           =   1905
      _ExtentX        =   3360
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
   Begin MSAdodcLib.Adodc AdoLectivo 
      Height          =   330
      Left            =   2100
      Top             =   1890
      Visible         =   0   'False
      Width           =   1905
      _ExtentX        =   3360
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
      Caption         =   "Lectivo"
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
      Left            =   210
      Top             =   2205
      Visible         =   0   'False
      Width           =   1905
      _ExtentX        =   3360
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
   Begin MSAdodcLib.Adodc AdoMaterias 
      Height          =   330
      Left            =   2100
      Top             =   2205
      Visible         =   0   'False
      Width           =   1905
      _ExtentX        =   3360
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
   Begin MSAdodcLib.Adodc AdoCursos 
      Height          =   330
      Left            =   3990
      Top             =   1260
      Visible         =   0   'False
      Width           =   1905
      _ExtentX        =   3360
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
   Begin MSAdodcLib.Adodc AdoEquivalencia 
      Height          =   330
      Left            =   3990
      Top             =   1575
      Visible         =   0   'False
      Width           =   1905
      _ExtentX        =   3360
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
      Caption         =   "Equivalencia"
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
      Left            =   14175
      TabIndex        =   5
      ToolTipText     =   "Formato de Fecha: DD/MM/AA"
      Top             =   735
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
   Begin MSAdodcLib.Adodc AdoNotasLibreta 
      Height          =   330
      Left            =   3990
      Top             =   1890
      Visible         =   0   'False
      Width           =   1905
      _ExtentX        =   3360
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
      Caption         =   "Equivalencia"
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   6315
      Left            =   7770
      TabIndex        =   7
      Top             =   1575
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   11139
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "PAGINA PRINCIPAL"
      TabPicture(0)   =   "FLibreta.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "LblA4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Picture1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "TxtCodigo"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Command2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Command3"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "HScroll1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "VScroll1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "LISTAR NOTAS POR EXCEL"
      TabPicture(1)   =   "FLibreta.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DGNotasLibreta"
      Tab(1).ControlCount=   1
      Begin VB.VScrollBar VScroll1 
         Height          =   3375
         Left            =   105
         TabIndex        =   14
         Top             =   735
         Width           =   330
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   330
         Left            =   3045
         TabIndex        =   13
         Top             =   4515
         Width           =   7785
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&V"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   105
         TabIndex        =   12
         Top             =   420
         Width           =   330
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&H"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   105
         TabIndex        =   11
         Top             =   4095
         Width           =   330
      End
      Begin VB.TextBox TxtCodigo 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1680
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   4515
         Width           =   1380
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00E0E0E0&
         Height          =   3585
         Left            =   525
         ScaleHeight     =   6.218
         ScaleMode       =   7  'Centimeter
         ScaleWidth      =   20.479
         TabIndex        =   8
         Top             =   420
         Width           =   11670
         Begin VB.PictureBox PictLibreta 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Height          =   1530
            Left            =   0
            ScaleHeight     =   2.593
            ScaleMode       =   7  'Centimeter
            ScaleWidth      =   17.463
            TabIndex        =   9
            Top             =   0
            Width           =   9960
         End
      End
      Begin MSDataGridLib.DataGrid DGNotasLibreta 
         Bindings        =   "FLibreta.frx":0038
         Height          =   5685
         Left            =   -74895
         TabIndex        =   17
         Top             =   420
         Width           =   9990
         _ExtentX        =   17621
         _ExtentY        =   10028
         _Version        =   393216
         AllowUpdate     =   0   'False
         Enabled         =   -1  'True
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
      Begin VB.Label LblA4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   420
         TabIndex        =   15
         Top             =   4515
         Width           =   1275
      End
   End
   Begin MSComctlLib.TreeView TVNivel 
      Height          =   1695
      Left            =   105
      TabIndex        =   24
      Top             =   2940
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   2990
      _Version        =   393217
      Style           =   7
      ImageList       =   "ImageList"
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   210
      Top             =   6825
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FLibreta.frx":0056
            Key             =   "C"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FLibreta.frx":0930
            Key             =   "N"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FLibreta.frx":0C4A
            Key             =   "M"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FLibreta.frx":1524
            Key             =   "E"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FLibreta.frx":1DFE
            Key             =   "H"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FLibreta.frx":2118
            Key             =   "Mj"
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList ImgLstHM 
      Left            =   1890
      Top             =   6825
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   2
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FLibreta.frx":2432
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FLibreta.frx":274C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label6 
      BackColor       =   &H00400000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " SELECCIONE EL &PERIODO"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   105
      TabIndex        =   23
      Top             =   4725
      Width           =   7575
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " MENSAJE:"
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
      Left            =   14175
      TabIndex        =   19
      Top             =   1155
      Width           =   1275
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " TITULO:"
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
      Left            =   7770
      TabIndex        =   18
      Top             =   1155
      Width           =   1275
   End
   Begin VB.Label LblFormato 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "FORMATO"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6195
      TabIndex        =   1
      Top             =   735
      Width           =   1170
   End
   Begin VB.Label LblDirigente 
      BackColor       =   &H00C0FFFF&
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
      ForeColor       =   &H008080FF&
      Height          =   330
      Left            =   9030
      TabIndex        =   4
      Top             =   735
      Width           =   5055
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " DIRIGENTE:"
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
      Left            =   7770
      TabIndex        =   3
      Top             =   735
      Width           =   1275
   End
   Begin VB.Label Label2 
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
      TabIndex        =   0
      Top             =   735
      Width           =   6105
   End
   Begin ComctlLib.ImageList ImgLstMenu 
      Left            =   1050
      Top             =   6825
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   21
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FLibreta.frx":29C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FLibreta.frx":2CE0
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FLibreta.frx":2FFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FLibreta.frx":3314
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FLibreta.frx":362E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FLibreta.frx":3970
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FLibreta.frx":3C8A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FLibreta.frx":3FA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FLibreta.frx":42BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FLibreta.frx":45D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FLibreta.frx":48B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FLibreta.frx":4A2C
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FLibreta.frx":4D46
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FLibreta.frx":5060
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FLibreta.frx":537A
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FLibreta.frx":5694
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FLibreta.frx":59AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FLibreta.frx":5CC8
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FLibreta.frx":5E7A
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FLibreta.frx":A9CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FLibreta.frx":23A1E
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FLibretas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Materias_Examenes() As TipoMaterias
Dim Y0 As Single
Dim y1 As Single
Dim X0 As Single
Dim x1 As Single
Dim Pos_Pict_X As Single
Dim Pos_Pict_Y As Single
Dim PosLineaX As Single
Dim TotalRegs(16) As Integer
Dim VPQBim1, VPQBim2, VPQBim3, VSQBim1, VSQBim2, VSQBim3, VTQBim1, VTQBim2, VTQBim3 As Currency
Dim VExamenPQ, VExamenSQ, VExamenTQ As Currency
Dim VProm, VPromT, VPromPQ, VPromSQ, VPromTQ, VPromFinal As Currency
Dim Directiva As String
Dim DirectorRegional As String
Dim AnchoMaxMaterias As Single
Dim AnchoMaximo As Single
Dim AltoMaximo As Single

Dim IdIni As Byte
Dim IdFin As Byte
Dim FJ As String
Dim FI As String
Dim Atrasos As String
Dim DiasL As String
Dim Evaluacion As String
Dim Imp_Informe As Boolean
Dim CodMatxPag(1 To 6) As Integer
Dim ContMaxPagina() As Integer
Dim ContMinPagina() As Integer

'Dim nodX As Node
' Create a simple PDF file using the mjwPDF class
Dim ObjPDF As New mjwPDF

Public Sub AddNewCta(TipoTC As String, Codigo As String, Detalle As String)
Dim SubInd As Integer
Dim inserteKey As Boolean
  If Len(Codigo) = 2 Then
     TVNivel.Nodes.Add , , Codigo, Detalle, ImageList.ListImages(1).key, ImageList.ListImages(1).key
     TVNivel.Tag = Codigo
  Else
     Select Case TipoTC
       Case "C": IE = 1
       Case "N": IE = 2
       Case "M": IE = 3
       Case "E": IE = 4
       Case "H": IE = 5
       Case "Mj": IE = 6
     End Select
     Cta_Sup = CodigoCuentaSup(Codigo)
     inserteKey = True
     For SubInd = 1 To TVNivel.Nodes.Count
         If TVNivel.Nodes(SubInd).key = Codigo Then inserteKey = False
     Next SubInd
     If inserteKey Then
        TVNivel.Nodes.Add Cta_Sup, tvwChild, Codigo, Detalle, ImageList.ListImages(IE).key, ImageList.ListImages(IE).key
        TVNivel.Tag = Codigo ' Mid$(Codigo, 2, Len(Codigo))
     End If
  End If
End Sub

Public Sub Grafico_Kinder(TopeDib As Single)
  PictLibreta.Cls
  PictLibreta.width = AnchoMaximo
  PictLibreta.Height = AltoMaximo

  PictLibreta.FontSize = 10
  PictLibreta.FontBold = True
  
  PictPrint_Texto PictLibreta, 3.5, 2.4, "EDUCADORA:"
           
  PictPrint_Texto PictLibreta, 0.7, 3.1, "EJES DE"
  PictPrint_Texto PictLibreta, 0.7, 3.5, "DESARROLLO"
  
  PictPrint_Texto PictLibreta, 0.7, 4.4, "DESARROLLO"
  PictPrint_Texto PictLibreta, 0.7, 4.8, "PERSONAL"
      
  PictPrint_Texto PictLibreta, 0.7, 5.7, "CONOCIMIENTO"
  PictPrint_Texto PictLibreta, 0.7, 6, "DEL ENTORNO"
  PictPrint_Texto PictLibreta, 0.7, 6.3, "INMEDIATO"

  PictPrint_Texto PictLibreta, 0.7, 7.3, "EXPRESION"
  PictPrint_Texto PictLibreta, 0.7, 7.7, "Y"
  PictPrint_Texto PictLibreta, 0.7, 8.1, "COMUNICACION"
  PictPrint_Texto PictLibreta, 0.7, 8.5, "CREATIVA"
  
  PictPrint_Texto PictLibreta, 0.7, 10, "OTRAS"
  PictPrint_Texto PictLibreta, 0.7, 10.5, "AREAS"
  
  PictLibreta.FontSize = 12
  PictPrint_Texto PictLibreta, 3.5, 3.25, "BLOQUES DE EXPERIENCIA"
  PictLibreta.FontSize = 11
  PictPrint_Texto PictLibreta, 3.5, 4.1, " 1.- Identidad y Autonomia personal"
  PictPrint_Texto PictLibreta, 3.5, 4.6, " 2.- Desarrollo físico, salud y Nutrición"
  PictPrint_Texto PictLibreta, 3.5, 5.1, " 3.- Desarrollo Social (socialización)"
      
  PictPrint_Texto PictLibreta, 3.5, 5.7, " 4.- Relaciones Lógico - Matemáticas"
  PictPrint_Texto PictLibreta, 3.5, 6.2, " 5.- Mundo Social, Cultura y Natural"
  
  PictPrint_Texto PictLibreta, 3.5, 6.9, " 6.- Expresión Corporal"
  PictPrint_Texto PictLibreta, 3.5, 7.4, " 7.- Expresión Lúdica"
  PictPrint_Texto PictLibreta, 3.5, 7.9, " 8.- Expresión Oral y Escrita"
  PictPrint_Texto PictLibreta, 3.5, 8.4, " 9.- Expresión Musical"
  PictPrint_Texto PictLibreta, 3.5, 8.9, "10.- Expresión Plástica"
  
  PictPrint_Texto PictLibreta, 3.5, 9.5, "1.- Cultura Física"
  PictPrint_Texto PictLibreta, 3.5, 10, "2.- Religión"
  PictPrint_Texto PictLibreta, 3.5, 10.5, "3.- Computación"
  PictPrint_Texto PictLibreta, 3.5, 11, "4.- Inglés"
  PictPrint_Texto PictLibreta, 3.5, 11.5, "5.- Tareas"
  
  PictLibreta.FontSize = 10
  PictPrint_Texto PictLibreta, 13.5, 1.5, "ASISTENCIA"
  PictPrint_Texto PictLibreta, 13.5, 2.2, "FALTAS JUSTIFICADAS"
  PictPrint_Texto PictLibreta, 13.5, 2.9, "FALTAS INJUSTIFICADAS"
  
  PictPrint_Texto PictLibreta, 14.5, 3.8, "N O M E N C L A T U R A"
  PictPrint_Texto PictLibreta, 13.5, 4.4, "MS"
  PictPrint_Texto PictLibreta, 13.5, 5, "S"
  PictPrint_Texto PictLibreta, 13.5, 5.6, "MDS"
  PictPrint_Texto PictLibreta, 13.5, 6.2, "PS"
  
      PictPrint_Texto PictLibreta, 14.5, 4.4, "Muy Satisfactorio"
      PictPrint_Texto PictLibreta, 14.5, 5, "Satisfactorio"
      PictPrint_Texto PictLibreta, 14.5, 5.6, "Medianamente Satisfactorio"
      PictPrint_Texto PictLibreta, 14.5, 6.2, "Poco Satisfactorio"
         
  PictPrint_Texto PictLibreta, 13.5, 7.5, "APROVECHAMIENTO"
  PictPrint_Texto PictLibreta, 15, 8.15, "DISCIPLINA"
  PictPrint_Texto PictLibreta, 13.5, 9.5, "Observación:"
  
 'Recuadro de la libreta
  PictLibreta.Line (0.5, 3)-(13, 12), QBColor(Negro), B
  PictLibreta.Line (0.5, 4.1)-(13, 4.1), QBColor(Negro)
  PictLibreta.Line (10.2, 3.6)-(13, 3.6), QBColor(Negro)
  PictLibreta.Line (3.3, 4.6)-(13, 4.6), QBColor(Negro)
  PictLibreta.Line (3.3, 5.1)-(13, 5.1), QBColor(Negro)
  PictLibreta.Line (0.5, 5.6)-(13, 5.6), QBColor(Negro)
  PictLibreta.Line (3.3, 6.2)-(13, 6.2), QBColor(Negro)
  PictLibreta.Line (0.5, 6.8)-(13, 6.8), QBColor(Negro)
  PictLibreta.Line (3.3, 7.4)-(13, 7.4), QBColor(Negro)
  PictLibreta.Line (3.3, 7.9)-(13, 7.9), QBColor(Negro)
  PictLibreta.Line (11.5, 7.9)-(11.5, 8.4), QBColor(Negro)
  PictLibreta.Line (3.3, 8.4)-(13, 8.4), QBColor(Negro)
  PictLibreta.Line (3.3, 8.9)-(13, 8.9), QBColor(Negro)
  PictLibreta.Line (0.5, 9.5)-(13, 9.5), QBColor(Negro)
  PictLibreta.Line (3.3, 10)-(13, 10), QBColor(Negro)
  PictLibreta.Line (3.3, 10.5)-(13, 10.5), QBColor(Negro)
  PictLibreta.Line (3.3, 11)-(13, 11), QBColor(Negro)
  PictLibreta.Line (3.3, 11.5)-(13, 11.5), QBColor(Negro)
  PictLibreta.Line (3.3, 3)-(3.3, 12), QBColor(Negro)
  PictLibreta.Line (10.2, 3)-(10.2, 12), QBColor(Negro)
  
  PictLibreta.Line (13.3, 1.4)-(19.5, 3.5), QBColor(Negro), B
  PictLibreta.Line (13.3, 2.1)-(19.5, 2.1), QBColor(Negro)
  PictLibreta.Line (13.3, 2.8)-(19.5, 2.8), QBColor(Negro)
  PictLibreta.Line (17.8, 1.4)-(17.8, 3.5), QBColor(Negro)
  
  PictLibreta.Line (13.3, 3.7)-(19.5, 6.8), QBColor(Negro), B
  PictLibreta.Line (13.3, 4.4)-(19.5, 4.4), QBColor(Negro)
  PictLibreta.Line (13.3, 5)-(19.5, 5), QBColor(Negro)
  PictLibreta.Line (13.3, 5.6)-(19.5, 5.6), QBColor(Negro)
  PictLibreta.Line (13.3, 6.2)-(19.5, 6.2), QBColor(Negro)
  PictLibreta.Line (14.35, 4.4)-(14.35, 6.8), QBColor(Negro)
  
  PictLibreta.Line (17, 7.4)-(19.5, 8.7), QBColor(Negro), B
  PictLibreta.Line (17, 8.05)-(19.5, 8.05), QBColor(Negro)
  
  PictLibreta.Line (15.8, 9.9)-(19.5, 9.9), QBColor(Negro)
  PictLibreta.Line (13.5, 10.6)-(19.5, 10.6), QBColor(Negro)
  PictLibreta.Line (13.5, 11.3)-(19.5, 11.3), QBColor(Negro)
  PictLibreta.Line (13.5, 12)-(19.5, 12), QBColor(Negro)
  
  PictPrint_Texto PictLibreta, 3.5, TopeDib + 3, TextoDirector
  PictPrint_Texto PictLibreta, 9.5, TopeDib + 3, "PROFESOR(A)"
  PictPrint_Texto PictLibreta, 15, TopeDib + 3, "REPRESENTANTE"
  PictLibreta.FontSize = 10
End Sub

Public Sub Grafico_Kinder_Informe_Final(TopeDib As Single)
  PictLibreta.Cls
  PictLibreta.width = AnchoMaximo
  PictLibreta.Height = AltoMaximo

  PictLibreta.FontSize = 10
  PictLibreta.FontBold = True
  
  PictPrint_Texto PictLibreta, 3, 3.4, "____ ALUMNO(A): "
  PictPrint_Texto PictLibreta, 3, 4.2, "HA CULMINADO EL: " & String(58, "_")
  PictPrint_Texto PictLibreta, 3, 5, "PASA AL: " & String(67, "_")

  PictPrint_Texto PictLibreta, 3, 6.2, "PAROVECHAMIENTO"
  PictPrint_Texto PictLibreta, 11, 6.2, "DISCIPLINA"
      
  PictPrint_Texto PictLibreta, 3, 7.2, "OBSERVACIONES: " & String(60, "_")
  PictPrint_Texto PictLibreta, 3, 8, String(76, "_")
  PictPrint_Texto PictLibreta, 3, 8.8, String(76, "_")
  
  PictPrint_Texto PictLibreta, 3, 9.8, "FECHA: " & FechaStrgCiudad(MBFecha)
 'Recuadro de la libreta
  PictLibreta.Line (1, 0.8)-(19, 13), QBColor(Negro), B
  
  PictLibreta.Line (6.5, 5.9)-(10.5, 6.8), QBColor(Negro), B
  PictLibreta.Line (13.1, 5.9)-(17, 6.8), QBColor(Negro), B
   
  PictPrint_Texto PictLibreta, 3, 11.5, "f.) ____________________"
  PictPrint_Texto PictLibreta, 12, 11.5, "f.) ____________________"
  
  PictPrint_Texto PictLibreta, 4, 12, TextoDirector
  PictPrint_Texto PictLibreta, 13, 12, "PROFESOR(A)"
  PictLibreta.FontSize = 10
End Sub

'''Public Sub Encabezado_Aprovechamiento1(TipoObjeto As Object, Optional OpcSupletorio As Boolean)
'''Dim PosLogo As Single
'''Dim Logo1 As String
'''Dim Ancho_Maya As Single
'''Dim PathDibujo As String
'''
'''  Ancho_Maya = Dato_Curso.ContMat * 1.5
'''  TipoObjeto.Cls
'''  TipoObjeto.FontBold = True
'''  TipoObjeto.FontSize = 12
'''  TipoObjeto.FontName = TipoTimes
'''
'''  PosLinea = 0.1
'''  Logo1 = RutaSistema & "\LOGOS\MINISEDU.JPG"
'''  TipoObjeto.PaintPicture LoadPicture(Logo1), 0.5, PosLinea, 3, 2.5
'''  Logo1 = RutaSistema & "\LOGOS\ECUADOR.GIF"
'''  TipoObjeto.PaintPicture LoadPicture(Logo1), (Ancho_Maya / 2) + 9.5, PosLinea, 2, 2
'''
'''  PosLinea = PosLinea + 2
'''  PictPrint_Texto 10.5, PosLinea, "REPÚBLICA DEL ECUADOR", , Ancho_Maya, True
'''  PosLinea = PosLinea + 0.45
'''  PictPrint_Texto 10.5, PosLinea, UCase$(Institucion1 & " " & Institucion2), , Ancho_Maya, True
'''  PosLinea = PosLinea + 0.45
'''  PictPrint_Texto 10.5, PosLinea, "CUADRO FINAL DE CALIFICACIONES", , Ancho_Maya, True
'''  PosLinea = PosLinea + 0.5
'''  TipoObjeto.FontSize = 10
'''  TipoObjeto2.FontSize = 10
'''  PictPrint_Texto 10.5, PosLinea, "A Ñ O   L E C T I V O:  " & Anio_Lectivo, , Ancho_Maya, True
'''  PosLinea = PosLinea + 0.45
''' 'Datos del Curso
'''  If Mid(Dato_Curso.Curso, 1, 1) < "3" Then
'''     PictPrint_Texto 1, PosLinea, "AÑO/CURSO: " & Dato_Curso.Bachiller
'''  Else
'''     PictPrint_Texto 1, PosLinea, "AÑO/CURSO: " & Dato_Curso.Curso_Texto
'''  End If
'''  PictPrint_Texto 14, PosLinea, "JORNADA: MATUTINA"
'''  PictPrint_Texto 19, PosLinea, "MODALIDAD: PRESENCIAL"
'''  PictPrint_Texto 26, PosLinea, "AMIE: " & Codigo_AMIE
'''  PosLinea = PosLinea + 0.4
'''  PictPrint_Texto 1, PosLinea, "PARALELO: " & Dato_Curso.Paralelo
'''  PictPrint_Texto 14, PosLinea, "ZONA: " & Zona
'''  PictPrint_Texto 19, PosLinea, "DISTRITO N° " & Distrito
'''  PosLinea = PosLinea + 0.4
'''  If Mid(Dato_Curso.Curso, 1, 1) >= "3" Then
'''     PictPrint_Texto 1, PosLinea, "TIPO DE BACHILLERATO: " & Dato_Curso.Bachiller
'''     PosLinea = PosLinea + 0.4
'''     PictPrint_Texto 1, PosLinea, Dato_Curso.Especialidad
'''     If Dato_Curso.Figura_Profesional <> Ninguno Then PictPrint_Texto 11, PosLinea, "FIGURA PROFESIONAL: " & Dato_Curso.Figura_Profesional
'''  End If
'''  TipoObjeto.FontBold = False
''' 'If Not OpcSupletorio Then PictPrint_Texto TipoObjeto.width - 5, PosLinea, FechaStrgCiudad(MBFecha)
'''  TipoObjeto.FontSize = 16
'''  PosLinea = PosLinea + 1.2
'''  TipoObjeto.FontBold = True
'''  PictPrint_Texto 1, PosLinea, "A P E L L I D O S"
'''  PosLinea = PosLinea + 0.8
'''  PictPrint_Texto 1, PosLinea, "Y   N O M B R E S"
'''  TipoObjeto.FontName = TipoCourier
'''  TipoObjeto.FontSize = 7
'''  TipoObjeto.FontBold = False
'''  PosLinea = 5.6
'''End Sub

Public Sub PDF_Encabezado_Aprovechamiento(TipoObjeto As Object)
Dim PosLogo As Single
Dim Logo1 As String
Dim PathDibujo As String
With TipoObjeto
    '.PDFSetFontStyle FONT_BOLD, True
    .PDFSetFontName FONT_TIMES
     PosLinea = 0.5
     Logo1 = RutaSistema & "\LOGOS\MINISEDU.JPG"
     cPrint.printImagen Logo1, 0.5, PosLinea, 1.8, 1
     Logo1 = RutaSistema & "\LOGOS\ECUADOR.JPG"
     cPrint.printImagen Logo1, 10, PosLinea, 1.2, 1
     PosLinea = PosLinea + 1.5
    .PDFSetFontSize 8
     PictPrint_Texto 1, PosLinea, "REPÚBLICA DEL ECUADOR", , 18.5, True
     PosLinea = PosLinea + 0.35
     PictPrint_Texto 1, PosLinea, UCase$(Institucion1 & " " & Institucion2), , 18.5, True
     PosLinea = PosLinea + 0.35
    .PDFSetFontSize 7
     PictPrint_Texto 1, PosLinea, "CUADRO FINAL DE CALIFICACIONES", , 18.5, True
     PosLinea = PosLinea + 0.35
     PictPrint_Texto 1, PosLinea, "A Ñ O   L E C T I V O:  " & Anio_Lectivo, , 18.5, True
     PosLinea = PosLinea + 0.4
     PictPrint_Texto 1, PosLinea, "JORNADA: MATUTINA"
     PictPrint_Texto 5, PosLinea, "MODALIDAD: PRESENCIAL"
     PictPrint_Texto 10, PosLinea, "AMIE: " & Codigo_AMIE
     PictPrint_Texto 14, PosLinea, "ZONA: " & Zona
     PictPrint_Texto 16, PosLinea, "DISTRITO N° " & Distrito
     
     PosLinea = PosLinea + 0.3
    'Datos del Curso
     If Mid(Dato_Curso.Curso, 1, 1) < "3" Then
        PictPrint_Texto 1, PosLinea, "AÑO/CURSO: " & Dato_Curso.Bachiller
     Else
        PictPrint_Texto 1, PosLinea, "AÑO/CURSO: " & Dato_Curso.Curso_Texto
     End If
     PictPrint_Texto 14, PosLinea, "PARALELO: " & Dato_Curso.Paralelo
     PosLinea = PosLinea + 0.3
     If Mid(Dato_Curso.Curso, 1, 1) >= "3" Then
        PictPrint_Texto 1, PosLinea, "TIPO DE BACHILLERATO: " & Dato_Curso.Tipo_Titulo
        PictPrint_Texto 14, PosLinea, Dato_Curso.Especialidad
        PosLinea = PosLinea + 0.3
        If Dato_Curso.Figura_Profesional <> Ninguno Then PictPrint_Texto 1, PosLinea, "FIGURA PROFESIONAL: " & Dato_Curso.Figura_Profesional
     End If
    .PDFSetFontSize 14
     PosLinea = PosLinea + 0.9
     PictPrint_Texto 1.2, PosLinea, "A P E L L I D O S"
     PosLinea = PosLinea + 0.8
     PictPrint_Texto 1.2, PosLinea, "Y   N O M B R E S"
End With
End Sub

Public Sub Notas_Promedio_Aprovechamiento(TipoObjeto As Object, NoPagina As Byte, SumaPromX As Single, CantMat As Long)
Dim IR_Temp As Single
 'Cuadro Final
  IR = 0
  If Pagina > 2 And (Pagina = NoPagina) And Si_No Then
     IR = Dato_Curso.PosXMat(Dato_Curso.ContMat - 1)
     If Dato_Curso.CantNotas > 5 Then IR = IR + (Dato_Curso.CantNotas * 0.7)
     Si_No = False
  ElseIf Pagina > 1 And (Pagina = NoPagina) And Si_No Then
     IR = Dato_Curso.PosXMat(Dato_Curso.ContMat - 1)
     If Dato_Curso.CantNotas > 5 Then IR = IR + (Dato_Curso.CantNotas * 0.7)
     Si_No = False
  End If
  If IR > 0 Then
     IR_Temp = IR
     PictPrint_Nota_Materia IR + 0.05, PosLinea, SumaPromX, , 2
     IR = IR + 1
     SumaPromX = Redondear(SumaPromX / CantMat, 2)
     PictPrint_Nota_Materia IR + 0.05, PosLinea, SumaPromX, , 2
     IR = IR + 0.7
     PictPrint_Texto IR + 0.05, PosLinea, Equivalencia(CCur(SumaPromX))
     IR = IR_Temp
     TipoObjeto.Line (IR, PosLinea - 0.05)-(IR + 2.6, PosLinea - 0.05), QBColor(Negro)
      
      TipoObjeto.Line (IR, PosLinea)-(IR, PosLinea + 2.6), QBColor(Negro)
      IR = IR + 1
      TipoObjeto.Line (IR, PosLinea)-(IR, PosLinea + 2.6), QBColor(Negro)
      IR = IR + 0.7
      TipoObjeto.Line (IR, PosLinea)-(IR, PosLinea + 2.6), QBColor(Negro)
      IR = IR + 0.9
      TipoObjeto.Line (IR, PosLinea)-(IR, PosLinea + 2.6), QBColor(Negro)
  End If
End Sub

Public Sub Encabezado_Aprovechamiento2(TipoObjeto As Object, NoPagina As Byte, Optional OpcSupletorio As Boolean)
Dim IR_Temp As Single
 'Imprimimos los encabezados de las materias y las Notas Finales
  PFil = PosLinea
  If NoPagina = 1 Then
     IR = Dato_Curso.PosXMat(1) - 0.7
     TipoObjeto.Line (IR, PosLinea)-(IR + 0.7, PosLinea + 2.6), QBColor(Negro), B
'''     cPrint.printTextoAngulo IR + 0.05, PFil + 2.5, 90, 4.5, 10, "EVALUACION DEL"
'''     cPrint.printTextoAngulo IR + 0.35, PFil + 2.5, 90, 4.5, 10, "COMPORTAMIENTO"
     If Dato_Curso.CantNotas > 5 Then
        IdIni = 1
        IdFin = 4
     Else
        IdIni = 1
        IdFin = 7
     End If
  ElseIf NoPagina = 2 Then
     If Dato_Curso.CantNotas > 5 Then
        IdIni = 5
        IdFin = 9
     Else
        IdIni = 8
        IdFin = 14
     End If
     If IdFin > Dato_Curso.ContMat Then IdFin = Dato_Curso.ContMat
  Else
     If Dato_Curso.CantNotas > 5 Then
        IdIni = 10
     Else
        IdIni = 15
     End If
     IdFin = Dato_Curso.ContMat
  End If
 'IR = IR + 1
  
  For I = IdIni To IdFin
      IR = Dato_Curso.PosXMat(I)
      If Dato_Curso.CodMat(I) < "997" Then
         TipoObjeto.FontBold = False
         TipoObjeto.FontName = TipoHelvetica
         TipoObjeto.FontSize = 7
         PictPrint_Texto IR + (Dato_Curso.CantNotas * 0.7) - 0.4, PosLinea + 0.4, Format(I, "00")
         If Dato_Curso.CantNotas > 5 Then
            PictPrint_Texto_Justifica IR + 0.1, IR + 4.8, PosLinea + 0.03, Dato_Curso.Materia(I)
         Else
            PictPrint_Texto_Justifica IR + 0.1, IR + 3, PosLinea + 0.03, Dato_Curso.Materia(I)
         End If
                           
         TipoObjeto.Line (IR, PosLinea)-(IR + (Dato_Curso.CantNotas * 0.7), PosLinea + 2.6), QBColor(Negro), B
         TipoObjeto.Line (IR, PosLinea + 0.8)-(IR + (Dato_Curso.CantNotas * 0.7), PosLinea + 0.8), QBColor(Negro)
         
         TipoObjeto.FontName = TipoHelvetica 'TipoArialNarrow
         TipoObjeto.FontBold = True
         Encabezado_Materias_Aprovechamiento IR, PFil + 2.5
      End If
      
     'MsgBox AnchoPict(I).Detalle
  Next I
 'Cuadro Final
  IR = 0
  If Pagina > 2 And (Pagina = NoPagina) And Si_No Then
     IR = Dato_Curso.PosXMat(Dato_Curso.ContMat - 1)
     If Dato_Curso.CantNotas > 5 Then IR = IR + (Dato_Curso.CantNotas * 0.7)
     Si_No = False
  ElseIf Pagina > 1 And (Pagina = NoPagina) And Si_No Then
     IR = Dato_Curso.PosXMat(Dato_Curso.ContMat - 1)
     If Dato_Curso.CantNotas > 5 Then IR = IR + (Dato_Curso.CantNotas * 0.7)
     Si_No = False
  End If
  If IR > 0 Then
     IR_Temp = IR
     'MsgBox IR & " - " & NoPagina
      cPrint.printTextoAngulo IR + 0.05, PFil + 2.5, 90, 4.5, 12, "S U M A"
      cPrint.printTextoAngulo IR + 0.45, PFil + 2.5, 90, 4.5, 12, "T O T A L"
      IR = IR + 1
      cPrint.printTextoAngulo IR + 0.05, PFil + 2.5, 90, 4.5, 10, "PROMEDIO"
      cPrint.printTextoAngulo IR + 0.3, PFil + 2.5, 90, 4.5, 10, "ANUAL"
      IR = IR + 0.7
      cPrint.printTextoAngulo IR + 0.05, PFil + 2.5, 90, 4.5, 12, "ESCALA"
      cPrint.printTextoAngulo IR + 0.45, PFil + 2.5, 90, 4.5, 12, "CUALITATIVA"
      
      IR = IR_Temp
      TipoObjeto.Line (IR, PosLinea)-(IR + 1, PosLinea + 2.6), QBColor(Negro), B
      IR = IR + 1
      TipoObjeto.Line (IR, PosLinea)-(IR + 0.7, PosLinea + 2.6), QBColor(Negro), B
      IR = IR + 0.7
      TipoObjeto.Line (IR, PosLinea)-(IR + 0.9, PosLinea + 2.6), QBColor(Negro), B
      
      If FormatoLibreta = "BIMESTRES" Then
         TipoObjeto.Line (IR + 1, PosLinea)-(IR + 2.3, PosLinea + 2.9), QBColor(Negro), B
         TipoObjeto.Line (IR + 2.3, PosLinea)-(IR + 4.2, PosLinea + 2.9), QBColor(Negro), B
         IR = IR + 0.15
         cPrint.printTextoAngulo IR, PFil, 90, 5, 26, "PROMEDIO ANUAL"
         IR = IR + 1.3
         cPrint.printTextoAngulo IR, PFil, 90, 7, 26, "DISCIPLINA"
         IR = IR + 1.1
         cPrint.printTextoAngulo IR, PFil, 90, 6, 52, "OBSERVACION"
      Else
         IR = IR + 0.2
    '''     cPrint.printTextoAngulo IR, PFil, 90, 5, 25, "PROMEDIO TOTAL"
      End If
  End If
  TipoObjeto.FontSize = 7
  PosLinea = 8.2
  PrimeraLinea = PosLinea
End Sub

Public Sub PDF_Encabezado_Aprovechamiento2(TipoObjeto As Object, NoPagina As Integer)
Dim IR_Temp As Single

 'Imprimimos los encabezados de las materias y las Notas Finales
  TipoObjeto.PDFSetFontSize 5
  PFil = PosLinea
  IdIni = ContMinPagina(NoPagina)
  IdFin = ContMaxPagina(NoPagina)
  If NoPagina = 1 Then
     IR = Dato_Curso.PosXMat(1) - 0.5
     cPrint.printTextoAngulo IR + 0.05, PFil + 0.2, 90, 0, 0, "EVALUACION DEL"
     cPrint.printTextoAngulo IR + 0.3, PFil + 0.2, 90, 0, 0, "COMPORTAMIENTO"
     PictPrint_Cuadro_Linea IR - 0.2, PosLinea - 0.9, IR + 0.4, PosLinea + 1.2, QBColor(Negro), "B"
  End If
 'MsgBox Dato_Curso.ContMat & vbCrLf & IdIni & vbCrLf & IdFin
  For I = IdIni To IdFin
      IR = Dato_Curso.PosXMat(I)
      If Dato_Curso.CodMat(I) < "997" Then
         TipoObjeto.PDFSetFontSize 6
         PictPrint_Texto IR + (Dato_Curso.CantNotas * 0.6) - 0.4, PosLinea - 0.7, Format(I, "00")
         TipoObjeto.PDFSetFontSize 7
         If Dato_Curso.CantNotas > 5 Then
            PictPrint_Texto IR + 0.05, PosLinea - 1, Dato_Curso.Materia(I), , 4.5
         Else
            PictPrint_Texto IR + 0.05, PosLinea - 1, Dato_Curso.Materia(I), , 2.5
         End If
         PictPrint_Cuadro_Linea IR - 0.1, PosLinea - 0.9, IR + (Dato_Curso.CantNotas * 0.6) - 0.1, PosLinea - 0.3, QBColor(Negro), "B"
         PFil = PosLinea + 0.2
         PDF_Encabezado_Materias_Aprovechamiento IR, PFil
      Else
        'Cuadro Final
         ObjPDF.PDFSetFontSize 8
         PosLineaX = IR
         PFil = PosLinea + 0.2
         IR = IR + 0.25
         cPrint.printTextoAngulo IR, PFil, 90, 4.5, 12, "S U M A"
         cPrint.printTextoAngulo IR + 0.3, PFil, 90, 4.5, 12, "T O T A L"
         IR = IR + 0.75
         cPrint.printTextoAngulo IR, PFil, 90, 4.5, 10, "PROMEDIO"
         cPrint.printTextoAngulo IR + 0.3, PFil, 90, 4.5, 10, "ANUAL"
         IR = IR + 0.75
         cPrint.printTextoAngulo IR, PFil, 90, 4.5, 12, "ESCALA"
         cPrint.printTextoAngulo IR + 0.3, PFil, 90, 4.5, 12, "CUALITATIVA"
         IR = IR + 0.3
         ObjPDF.PDFSetFontSize 6
      End If
     'MsgBox AnchoPict(I).Detalle
  Next I
  PosLinea = 8.2
  PrimeraLinea = PosLinea
  AnchoMaxMaterias = IR + 0.15
End Sub

Public Sub Listar_Alumnos_Curso(Curso As String)
Dim itmX As ListItem
Dim CadAux As String
 If Len(Curso) = 7 Then
    CadAux = Leer_Datos_del_Curso(Curso)
   'Colocamos el cero si hay notas en blanco o Nulos
    sSQL = "UPDATE Catalogo_Estudiantil " _
         & "SET Id_No = 0 " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND Mid$(CodigoE,1,7) = '" & Curso & "' " _
         & "AND Id_No IS NULL "
    ConectarAdoExecute sSQL
   'Actualizamos Catalogo_Estudiantil
    If SQL_Server Then
       sSQL = "UPDATE Catalogo_Estudiantil " _
            & "SET Id_No = CONVERT(TINYINT,Mid$(CodigoE,9,2)) "
    Else
       sSQL = "UPDATE Catalogo_Estudiantil " _
            & "SET Id_No = Val(Mid$(CodigoE,9,2)) "
    End If
    sSQL = sSQL & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND Mid$(CodigoE,1,7) = '" & Curso & "' " _
         & "AND TC = 'M' "
    If SQL_Server Then
       sSQL = sSQL & "AND Id_No <> CONVERT(TINYINT,Mid$(CodigoE,9,2)) "
    Else
       sSQL = sSQL & "AND Id_No <> Val(Mid$(CodigoE,9,2)) "
    End If
    'MsgBox sSQL
    ConectarAdoExecute sSQL
   'Iniciamos las materias que se promedian
    sSQL = "UPDATE Trans_Notas " _
         & "SET CodMatP = '.' " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND CodE = '" & Curso & "' "
    ConectarAdoExecute sSQL
   'Iniciamos las posiciones de las materias del curso
    sSQL = "UPDATE Trans_Notas " _
         & "SET Id_No = 0 " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND CodE = '" & Curso & "' "
    ConectarAdoExecute sSQL
   'Iniciamos las posiciones de las materias del curso
    sSQL = "UPDATE Trans_Notas " _
         & "SET Orden = 0 " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND CodE = '" & Curso & "' "
    ConectarAdoExecute sSQL
   'Actualiza el orden de libretas
     If SQL_Server Then
        sSQL = "UPDATE Trans_Notas " _
             & "SET Id_No=CE.Id_No,CodMatP=CE.CodMatP,Orden=CE.Orden " _
             & "FROM Trans_Notas As TN,Catalogo_Estudiantil As CE "
     Else
        sSQL = "UPDATE Trans_Notas As TN,Catalogo_Estudiantil As CE " _
             & "SET TN.Id_No=CE.Id_No,TN.CodMatP=CE.CodMatP,TN.Orden=CE.Orden "
     End If
     sSQL = sSQL & "WHERE TN.Item = '" & NumEmpresa & "' " _
          & "AND TN.Periodo = '" & Periodo_Contable & "' " _
          & "AND TN.CodE = '" & Curso & "' " _
          & "AND TN.CodMat = CE.CodMat " _
          & "AND TN.CodE = Mid$(CE.CodigoE,1,7) " _
          & "AND TN.Item = CE.Item " _
          & "AND TN.Periodo = CE.Periodo "
     ConectarAdoExecute sSQL
    'Borramos las notas con basura
     sSQL = "DELETE * " _
          & "FROM Trans_Notas " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "AND CodE = '" & Curso & "' " _
          & "AND Id_No = 0 "
     ConectarAdoExecute sSQL
 End If
  Contador = 0
  LstVAlumnos.ListItems.Clear
  LstVAlumnos.View = 3
  sSQL = "SELECT C.Cliente As Alumno, C.Sexo, C.Celular, C.Telefono, CM.* " _
       & "FROM Clientes As C,Clientes_Matriculas As CM " _
       & "WHERE CM.Item = '" & NumEmpresa & "' " _
       & "AND CM.Periodo = '" & Periodo_Contable & "' " _
       & "AND CM.Grupo_No = '" & Curso & "' " _
       & "AND C.Codigo = CM.Codigo " _
       & "ORDER BY C.Cliente,C.Sexo "
  'MsgBox sSQL
  SelectAdodc AdoAlumnos, sSQL
  With AdoAlumnos.Recordset
      'MsgBox sSQL & vbCrLf & .RecordCount
   If .RecordCount > 0 Then
       Codigo1 = Leer_Datos_del_Curso(Curso)
       Codigo1 = Dato_Curso.Especialidad
       Do While Not .EOF
          Contador = Contador + 1
          Codigo = .Fields("Codigo")
          TipoBenef = .Fields("Sexo")
          NombreCliente = .Fields("Alumno")
         'MsgBox Codigo & vbCrLf & NombreCliente
          Set itmX = LstVAlumnos.ListItems.Add(, "C" & Codigo, NombreCliente)
          If TipoBenef = "M" Then
             itmX.Icon = 1       ' Icono Hombre
             itmX.SmallIcon = 1
          Else
             itmX.Icon = 2       ' Icono Mujer
             itmX.SmallIcon = 2
          End If
         .MoveNext
       Loop
   End If
  End With
  LstVAlumnos.ColumnHeaders.Item(1).Text = "(" & Contador & ") Nombre del Alumno"
  NumAlumnos = Contador
End Sub

Public Sub Libreta_Del_Alumno_Bimestres(AdoLib As Adodc)
Dim AnchoDib As Single
Dim AltoDib As Single
Dim Curso As String
Dim Alumno As String
Dim Paralelo As String
Dim PosXPict As Single
Dim AnchoPict() As CtasAsiento
Dim Y0 As Single
Dim y1 As Single
Dim X0 As Single
Dim x1 As Single
Dim PosLineaX As Single
Dim TotalRegs(11) As Integer
Dim Presentar_Notas_Libreta As Boolean
  Presentar_Notas_Libreta = False
  With AdoLib.Recordset
   If .RecordCount > 0 Then
       Curso = .Fields("Curso")
       Paralelo = .Fields("Paralelo")
       Alumno = .Fields("Alumno")
       NombreCliente = .Fields("Alumno")
   End If
  End With
  PictLibreta.FontName = TipoArialNarrow   ' TipoTimes
  PictLibreta.ForeColor = QBColor(Negro)
  PosColumna = 9.4
  Select Case Mid$(Curso, 1, 4)
    Case "0.00" To "1.01" '
         AnchoDib = 20: AltoDib = 13
         If OpcPeriodo("PF", LstPeriodos) Then
            Grafico_Kinder_Informe_Final AltoDib
            PictLibreta.PaintPicture LoadPicture(LogoTipo), 1.3, 1, 4, 2
            PictLibreta.FontSize = 20
            PictPrint_Texto PictLibreta, 4, 1, Empresa
            PictLibreta.FontSize = 18
            PictPrint_Texto PictLibreta, 7.8, 1.8, "INFORME FINAL"
            PictLibreta.FontSize = 10
            PictLibreta.FontBold = True
            PictPrint_Texto PictLibreta, 6, 3.4, Alumno
            PictLibreta.FontBold = False
         Else
            Grafico_Kinder AltoDib
            PictLibreta.PaintPicture LoadPicture(LogoTipo), 0.1, 0.1, 3, 1.5
            PictLibreta.FontSize = 15
            PictPrint_Texto PictLibreta, 4, 0.1, Empresa
            PictLibreta.FontSize = 10
            PictPrint_Texto PictLibreta, 4, 0.7, "FICHA DE DESARROLLO DE DESTREZAS Y HABILIDADES"
            
            PictPrint_Texto PictLibreta, 14, 0.2, "NIVEL:"
            If AdoLib.Recordset.Fields("Sexo") = "M" Then
               PictPrint_Texto PictLibreta, 0.5, 1.8, "ALUMNO:"
            Else
               PictPrint_Texto PictLibreta, 0.5, 1.8, "ALUMNA:"
            End If
            PictLibreta.FontBold = False
            PictLibreta.FontSize = 10
            PictPrint_Texto PictLibreta, 15.3, 0.2, Curso
            PictPrint_Texto PictLibreta, 14, 0.7, Paralelo
            PictPrint_Texto PictLibreta, 2.7, 1.8, Alumno
            PosColumna = 0
            PosLinea = 0
         End If
    Case "1.02" To "1.99"
         AnchoDib = 20.3: AltoDib = 11.5
         PictLibreta.Cls
         PictLibreta.PaintPicture LoadPicture(LogoTipo), 0.45, 0.2, 3.9, 1.8
         PictLibreta.FontSize = 8
         PFil = AltoDib - 1.9
         PictLibreta.Line (0.3, PFil)-(4.1, PFil + 2.2), QBColor(Negro), B
         PictLibreta.FontUnderline = True
         PictPrint_Texto PictLibreta, 0.35, PFil + 0.05, "   ESCALA VALORATIVA   "
         PictLibreta.FontUnderline = False
         PictPrint_Texto PictLibreta, 0.5, PFil + 0.4, "19 - 20 SOBRESALIENTE"
         PictPrint_Texto PictLibreta, 0.5, PFil + 0.75, "16 - 18 MUY BUENO"
         PictPrint_Texto PictLibreta, 0.5, PFil + 1.1, "14 - 15 BUENO"
         PictPrint_Texto PictLibreta, 0.5, PFil + 1.45, "12 - 13 REGULAR"
         PictPrint_Texto PictLibreta, 0.5, PFil + 1.8, "00 - 11 INSUFICIENTE"
         PictLibreta.Line (0.3, 0.1)-(19, AltoDib - 2), QBColor(Negro), B
         PictLibreta.Line (0.3, 2.2)-(19, 2.2), QBColor(Negro)   ' horizontal
         PictLibreta.Line (0.3, 4.3)-(13.05, 4.3), QBColor(Negro)
         PCol = PosColumna - 0.35
         PictLibreta.Line (0.3, 3.8)-(PCol, 3), QBColor(Negro), B
         PictLibreta.Line (PCol, 2.2)-(PCol, AltoDib - 2), QBColor(Negro)
         PictLibreta.Line (PCol, 2.8)-(19, 2.8), QBColor(Negro)
         PictLibreta.FontSize = 7
         PictPrint_Texto PictLibreta, PCol + 0.15, 2.4, "I QUIMESTRE"
         PictPrint_Texto PictLibreta, PCol + 2.15, 2.4, "II QUIMESTRE"
         PictLibreta.FontSize = 9
         PictPrint_Texto PictLibreta, PCol + 4.6, 2.35, "OBSERVACIONES GENERALES"
         PictLibreta.FontSize = 13
         PictPrint_Texto PictLibreta, 1.5, 3.8, "M A T E R I A S"
         For I = 1 To 5
             If I = 3 Or I = 5 Then
                PictLibreta.Line (PCol, 2.2)-(PCol, AltoDib - 2), QBColor(Negro)
             Else
                PictLibreta.Line (PCol, 2.7)-(PCol, AltoDib - 2), QBColor(Negro)
             End If
             Select Case I
               Case 1, 3: cPrint.printTextoAngulo PictLibreta, PCol + 0.35, 4.2, 90, 4, 10, "Bimestre I"
               Case 2, 4: cPrint.printTextoAngulo PictLibreta, PCol + 0.35, 4.2, 90, 4, 10, "Bimestre II"
             End Select
             PCol = PCol + 1
         Next I
         PFil = 3.5
         For I = 1 To 9
             PictLibreta.Line (13.2, PFil)-(18.85, PFil), QBColor(Negro)
             PFil = PFil + 0.65
         Next I
         'PictLibreta.Line (PFil, 2.3)-(PFil, 2.8), QBColor(Negro)
         PictLibreta.FontSize = 20
         PictPrint_Texto PictLibreta, 3.7, 0.1, Empresa
         PictLibreta.FontSize = 11
         PictPrint_Texto PictLibreta, 4, 0.9, "LIBRETA DE CALIFICACIONES"
         PictLibreta.FontSize = 16
         PictPrint_Texto PictLibreta, 4, 1.4, "AÑO LECTIVO " & Anio_Lectivo
         PictLibreta.FontSize = 9
         PictPrint_Texto PictLibreta, 14, 1.7, FechaStrgCiudad(MBFecha)
         PosLinea = 1.4
         PictLibreta.FontSize = 8
         If AdoLib.Recordset.Fields("Sexo") = "M" Then
            PictPrint_Texto PictLibreta, 0.5, 2.25, "Alumno:"
         Else
            PictPrint_Texto PictLibreta, 0.5, 2.25, "Alumna:"
         End If
         PictPrint_Texto PictLibreta, 0.5, 3.05, "Curso:"
         PictLibreta.FontBold = False
         PictPrint_Texto PictLibreta, 7.6, 3.05, Curso
         PictLibreta.FontSize = 10
         PictPrint_Texto PictLibreta, 1.2, 2.55, Alumno
         PictPrint_Texto PictLibreta, 1.2, 3.35, Paralelo
         PictLibreta.FontBold = True
         PictPrint_Texto PictLibreta, 5, AltoDib + 1, "REPRESENTANTE"
         PictPrint_Texto PictLibreta, 9, AltoDib + 1, TextoDirector
         PictPrint_Texto PictLibreta, 12.5, AltoDib + 1, TextoSecretario1
         PictPrint_Texto PictLibreta, 16, AltoDib + 1, "PROFESOR(A)"
         PosLinea = 4.5
         Presentar_Notas_Libreta = True
    Case "2.00" To "3.99"
         AnchoDib = 20.3: AltoDib = 13.45
         PictLibreta.Cls
         PictLibreta.PaintPicture LoadPicture(LogoTipo), 0.45, 0.2, 3.9, 1.9
         PictLibreta.FontSize = 8
         PFil = AltoDib - 1.9
         PictLibreta.Line (0.3, PFil)-(4.1, PFil + 2.1), QBColor(Negro), B
         PictLibreta.FontUnderline = True
         PictPrint_Texto PictLibreta, 0.35, PFil + 0.05, "   ESCALA VALORATIVA   "
         PictLibreta.FontUnderline = False
         PictPrint_Texto PictLibreta, 0.5, PFil + 0.4, "19 - 20 SOBRESALIENTE"
         PictPrint_Texto PictLibreta, 0.5, PFil + 0.75, "16 - 18 MUY BUENO"
         PictPrint_Texto PictLibreta, 0.5, PFil + 1.1, "14 - 15 BUENO"
         PictPrint_Texto PictLibreta, 0.5, PFil + 1.45, "12 - 13 REGULAR"
         PictPrint_Texto PictLibreta, 0.5, PFil + 1.8, "00 - 11 INSUFICIENTE"

         PictLibreta.Line (0.3, 0.1)-(19, AltoDib - 2), QBColor(Negro), B
         PictLibreta.Line (0.3, 2.2)-(19, 4.3), QBColor(Negro), B
         PCol = PosColumna - 0.35
         PictLibreta.Line (0.3, 3)-(PCol, 3.8), QBColor(Negro), B
         PictLibreta.Line (PCol, 2.2)-(PCol, AltoDib - 2), QBColor(Negro)
         PictLibreta.Line (PCol, 2.7)-(17.05, 2.7), QBColor(Negro)
         PictLibreta.FontSize = 10
         PictPrint_Texto PictLibreta, PCol + 0.25, 2.25, "PRIMER QUIMESTRE"
         PictPrint_Texto PictLibreta, PCol + 4.15, 2.25, "SEGUNDO QUIMESTRE"
         PictLibreta.FontSize = 13
         PictPrint_Texto PictLibreta, 1.5, 3.8, "M A T E R I A S"
         For I = 1 To 10
             If I = 5 Or I = 9 Or I = 10 Then
                PictLibreta.Line (PCol, 2.2)-(PCol, AltoDib - 2), QBColor(Negro)
             Else
                PictLibreta.Line (PCol, 2.7)-(PCol, AltoDib - 2), QBColor(Negro)
             End If
             Select Case I
               Case 1, 5: cPrint.printTextoAngulo PictLibreta, PCol + 0.35, 4.2, 90, 4, 10, "Bimestre 1"
               Case 2, 6: cPrint.printTextoAngulo PictLibreta, PCol + 0.35, 4.2, 90, 4, 10, "Bimestre 2"
               Case 3, 7: cPrint.printTextoAngulo PictLibreta, PCol + 0.35, 4.2, 90, 4, 10, "Sumatoria"
               Case 4, 8: cPrint.printTextoAngulo PictLibreta, PCol + 0.35, 4.2, 90, 4, 10, "Promedio"
               Case 9: cPrint.printTextoAngulo PictLibreta, PCol + 0.25, 4.2, 90, 4, 9, "SUPLETORIO"
               Case 10: cPrint.printTextoAngulo PictLibreta, PCol + 0.05, 4.2, 90, 5, 10, "Promedio"
                        cPrint.printTextoAngulo PictLibreta, PCol + 0.4, 4.2, 90, 5, 10, "Final"
             End Select
             PCol = PCol + 1
         Next I
         'PictLibreta.Line (PFil, 2.3)-(PFil, 2.8), QBColor(Negro)
         PictLibreta.FontSize = 20
         PictPrint_Texto PictLibreta, 3.7, 0.1, Empresa
         PictLibreta.FontSize = 11
         PictPrint_Texto PictLibreta, 4, 0.9, "LIBRETA DE CALIFICACIONES"
         PictLibreta.FontSize = 16
         PictPrint_Texto PictLibreta, 4, 1.4, "AÑO LECTIVO " & Anio_Lectivo
         PictLibreta.FontSize = 9
         PictPrint_Texto PictLibreta, 14, 1.7, FechaStrgCiudad(MBFecha)
         PosLinea = 1.4
         PictLibreta.FontSize = 8
         If AdoLib.Recordset.Fields("Sexo") = "M" Then
            PictPrint_Texto PictLibreta, 0.5, 2.25, "Alumno:"
         Else
            PictPrint_Texto PictLibreta, 0.5, 2.25, "Alumna:"
         End If
         PictPrint_Texto PictLibreta, 0.5, 3.05, "Curso:"
         PictLibreta.FontBold = False
         PictPrint_Texto PictLibreta, 7.6, 3.05, Curso
         PictLibreta.FontSize = 10
         PictPrint_Texto PictLibreta, 1.2, 2.6, Alumno
         PictPrint_Texto PictLibreta, 1.2, 3.4, Paralelo
         PictLibreta.FontBold = True
         PictPrint_Texto PictLibreta, 5, AltoDib - 0.2, "REPRESENTANTE"
         PictPrint_Texto PictLibreta, 10, AltoDib - 0.2, TextoRector
         PictPrint_Texto PictLibreta, 14.5, AltoDib - 0.2, TextoSecretario2
         PosLinea = 4.35
         Presentar_Notas_Libreta = True
  End Select
 With AdoLib.Recordset
  If .RecordCount > 0 And Presentar_Notas_Libreta Then
      PictLibreta.FontSize = 9
      PictLibreta.FontName = TipoArial
      Codigo = .Fields("Codigo")
      Cadena1 = .Fields("Curso")
      Cadena2 = .Fields("Alumno")
      Codigo4 = .Fields("CodE")
      Select Case Mid$(Codigo4, 1, 4)
        Case "0.00" To "1.01": y1 = 11.35
        Case "1.02" To "1.99": y1 = 9.5
        Case "2.00" To "3.99": y1 = 11.45
      End Select
      For I = 0 To 10
          TotalRegs(I) = 0
      Next I
      VPQBim1 = 0: VPQBim2 = 0: VSQBim1 = 0: VSQBim2 = 0: VPromPQ = 0: VPromSQ = 0: VPromFinal = 0
      Do While Not .EOF
         Opciones = .Fields("Orden")
         JR = 1
         If PosColumna > 0 Then
            If .Fields("Orden") = 9 Then
                PosLinea = y1
                If TotalReg = 0 Then TotalReg = 1
                If TotalRegs(1) = 0 Then TotalRegs(1) = 1
                If TotalRegs(2) = 0 Then TotalRegs(2) = 1
                If TotalRegs(4) = 0 Then TotalRegs(4) = 1
                If TotalRegs(5) = 0 Then TotalRegs(5) = 1
                If TotalRegs(6) = 0 Then TotalRegs(6) = 1
                If TotalRegs(8) = 0 Then TotalRegs(8) = 1
                If TotalRegs(10) = 0 Then TotalRegs(10) = 1
                PictLibreta.FontBold = True
                PictLibreta.FontSize = 10
                PictPrint_Texto PictLibreta, 4.5, PosLinea, "SUMATORIA"
                PictLibreta.FontSize = 8
                PictLibreta.FontBold = False
                IR = PosColumna
                If ("2.00" <= Codigo4) And (Codigo4 <= "3.99") Then
                   If VPQBim1 > 0 And Print_Nota(1) Then PictPrint_Texto PictLibreta, IR - 0.1, PosLinea, Format(VPQBim1, "00")
                   IR = IR + JR
                   If VPQBim2 > 0 And Print_Nota(2) Then PictPrint_Texto PictLibreta, IR - 0.1, PosLinea, Format(VPQBim2, "00")
                   IR = IR + (JR * 2)
                   If VPromPQ > 0 And Print_Nota(4) Then PictPrint_Texto PictLibreta, IR - 0.1, PosLinea, Format(VPromPQ, "00")
                   IR = IR + JR
                   If VSQBim1 > 0 And Print_Nota(5) Then PictPrint_Texto PictLibreta, IR - 0.1, PosLinea, Format(VSQBim1, "00")
                   IR = IR + JR
                   If VSQBim2 > 0 And Print_Nota(6) Then PictPrint_Texto PictLibreta, IR - 0.1, PosLinea, Format(VSQBim2, "00")
                   IR = IR + (JR * 2)
                   If VPromSQ > 0 And Print_Nota(8) Then PictPrint_Texto PictLibreta, IR - 0.1, PosLinea, Format(VPromSQ, "00")
                   IR = IR + (JR * 2)
                   IR = IR - 0.4
                   If VPromFinal > 0 And Print_Nota(10) Then PictPrint_Texto PictLibreta, IR, PosLinea, Format(VPromFinal, "00.00")
                Else
                   If VPQBim1 > 0 Then PictPrint_Texto PictLibreta, IR - 0.1, PosLinea, Format(VPQBim1, "00")
                   IR = IR + JR
                   If VPQBim2 > 0 Then PictPrint_Texto PictLibreta, IR - 0.1, PosLinea, Format(VPQBim2, "00")
                   IR = IR + JR
                   If VSQBim1 > 0 Then PictPrint_Texto PictLibreta, IR - 0.1, PosLinea, Format(VSQBim1, "00")
                   IR = IR + JR
                   If VSQBim2 > 0 Then PictPrint_Texto PictLibreta, IR - 0.1, PosLinea, Format(VSQBim2, "00")
                   IR = IR + JR
                End If
                PosLinea = PosLinea + 0.4
                PictLibreta.FontBold = True
                PictLibreta.FontSize = 10
                PictPrint_Texto PictLibreta, 4.5, PosLinea, "APROVECHAMIENTO"
                PictLibreta.FontBold = False
                PictLibreta.FontSize = 8
                IR = PosColumna - 0.05
                If ("2.00" <= Codigo4) And (Codigo4 <= "3.99") Then
                   If VPQBim1 > 0 And Print_Nota(1) Then PictPrint_Texto PictLibreta, IR - 0.2, PosLinea, Format(VPQBim1 / TotalRegs(1), "00.00")
                   IR = IR + JR
                   If VPQBim2 > 0 And Print_Nota(2) Then PictPrint_Texto PictLibreta, IR - 0.2, PosLinea, Format(VPQBim2 / TotalRegs(2), "00.00")
                   IR = IR + (JR * 2)
                   If VPromPQ > 0 And Print_Nota(4) Then PictPrint_Texto PictLibreta, IR - 0.2, PosLinea, Format(VPromPQ / TotalRegs(4), "00.00")
                   IR = IR + JR
                   If VSQBim1 > 0 And Print_Nota(5) Then PictPrint_Texto PictLibreta, IR - 0.2, PosLinea, Format(VSQBim1 / TotalRegs(5), "00.00")
                   IR = IR + JR
                   If VSQBim2 > 0 And Print_Nota(6) Then PictPrint_Texto PictLibreta, IR - 0.2, PosLinea, Format(VSQBim2 / TotalRegs(6), "00.00")
                   IR = IR + (JR * 2)
                   If VPromSQ > 0 And Print_Nota(8) Then PictPrint_Texto PictLibreta, IR - 0.2, PosLinea, Format(VPromSQ / TotalRegs(8), "00.00")
                   IR = IR + (JR * 2)
                   IR = IR - 0.2
                   If VPromFinal > 0 And Print_Nota(10) Then PictPrint_Texto PictLibreta, IR, PosLinea, Format(VPromFinal / TotalRegs(10), "00.00")
                Else
                   If VPQBim1 > 0 Then PictPrint_Texto PictLibreta, IR - 0.2, PosLinea, Format(VPQBim1 / TotalRegs(1), "00.00")
                   IR = IR + JR
                   If VPQBim2 > 0 Then PictPrint_Texto PictLibreta, IR - 0.2, PosLinea, Format(VPQBim2 / TotalRegs(2), "00.00")
                   IR = IR + JR
                   If VSQBim1 > 0 Then PictPrint_Texto PictLibreta, IR - 0.2, PosLinea, Format(VSQBim1 / TotalRegs(5), "00.00")
                   IR = IR + JR
                   If VSQBim2 > 0 Then PictPrint_Texto PictLibreta, IR - 0.2, PosLinea, Format(VSQBim2 / TotalRegs(6), "00.00")
                   IR = IR + JR
                End If
                PosLinea = PosLinea + 0.4
                PictLibreta.FontBold = True
                PictLibreta.FontSize = 10
                PictPrint_Texto PictLibreta, 4.5, PosLinea, .Fields("Materia")
                PictLibreta.FontBold = False
                PictLibreta.FontSize = 8
                IR = PosColumna
                'MsgBox .Fields("Materia")
                If ("2.00" <= Codigo4) And (Codigo4 <= "3.99") Then
                   If Print_Nota(1) Then PictPrint_Nota_Materia PictLibreta, IR, PosLinea, .Fields("PQBim1")
                   IR = IR + JR
                   If Print_Nota(2) Then PictPrint_Nota_Materia PictLibreta, IR, PosLinea, .Fields("PQBim2")
                   IR = IR + JR
                   Sumatoria = .Fields("PQBim1") + .Fields("PQBim2")
                   If Print_Nota(3) Then PictPrint_Nota_Materia PictLibreta, IR, PosLinea, Sumatoria
                   IR = IR + JR
                   If Print_Nota(4) Then PictPrint_Nota_Materia PictLibreta, IR, PosLinea, .Fields("PromPQ")
                   IR = IR + JR
                   If Print_Nota(5) Then PictPrint_Nota_Materia PictLibreta, IR, PosLinea, .Fields("SQBim1")
                   IR = IR + JR
                   If Print_Nota(6) Then PictPrint_Nota_Materia PictLibreta, IR, PosLinea, .Fields("SQBim2")
                   IR = IR + JR
                   Sumatoria = .Fields("SQBim1") + .Fields("SQBim2")
                   If Print_Nota(7) Then PictPrint_Nota_Materia PictLibreta, IR, PosLinea, Sumatoria
                   IR = IR + JR
                   If Print_Nota(8) Then PictPrint_Nota_Materia PictLibreta, IR, PosLinea, .Fields("PromSQ")
                   IR = IR + JR
                   If Print_Nota(9) Then PictPrint_Nota_Materia PictLibreta, IR, PosLinea, .Fields("Supletorio")
                   IR = IR + JR
                   If Print_Nota(10) Then PictPrint_Nota_Materia PictLibreta, IR, PosLinea, .Fields("PromFinal")
                   IR = IR + JR
                Else
                   PictPrint_Nota_Materia PictLibreta, IR, PosLinea, .Fields("PQBim1")
                   IR = IR + JR
                   PictPrint_Nota_Materia PictLibreta, IR, PosLinea, .Fields("PQBim2")
                   IR = IR + JR
                   PictPrint_Nota_Materia PictLibreta, IR, PosLinea, .Fields("SQBim1")
                   IR = IR + JR
                   PictPrint_Nota_Materia PictLibreta, IR, PosLinea, .Fields("SQBim2")
                   PosLinea = PosLinea + 0.4
                   PictLibreta.FontBold = True
                   PictPrint_Texto PictLibreta, 4.5, PosLinea, "PROMEDIO FINAL"
                   PictLibreta.FontBold = False
                   Sumatoria = 0
                   Contador = 0
                   If VPQBim1 > 0 Then
                      Sumatoria = Sumatoria + (VPQBim1 / TotalRegs(1))
                      Contador = Contador + 1
                   End If
                   If VPQBim2 > 0 Then
                      Sumatoria = Sumatoria + (VPQBim2 / TotalRegs(2))
                      Contador = Contador + 1
                   End If
                   If VSQBim1 > 0 Then
                      Sumatoria = Sumatoria + (VSQBim1 / TotalRegs(5))
                      Contador = Contador + 1
                   End If
                   If VSQBim2 > 0 Then
                      Sumatoria = Sumatoria + (VSQBim2 / TotalRegs(6))
                      Contador = Contador + 1
                   End If
                   If Contador = 0 Then Contador = 1
                   Sumatoria = Format(Sumatoria / Contador, "00.00")
                   If Sumatoria > 0 Then PictPrint_Texto PictLibreta, IR - 0.2, PosLinea, Format(Sumatoria, "00.00")
                End If
                PictLibreta.FontBold = False
            Else
                Si_No = False
                ' If .Fields("CodMat") = "024" Then Si_No = True
                If .Fields("C") <> 0 Then Si_No = True
                PictLibreta.FontBold = False
                PictLibreta.FontUnderline = False
                ' MsgBox .Fields("CodMatP")
                If .Fields("CodMatP") = Ninguno Then
                    Contador = Contador + 1
                    PictPrint_Texto PictLibreta, 0.7, PosLinea, Format(Contador, "00") & ".-"
                End If
                PictPrint_Texto PictLibreta, 1.6, PosLinea, .Fields("Materia")
                IR = PosColumna
                If ("2.00" <= Codigo4) And (Codigo4 <= "3.99") Then
                   If Print_Nota(1) Then PictPrint_Nota_Materia PictLibreta, IR, PosLinea, .Fields("PQBim1"), Si_No
                   IR = IR + JR
                   If Print_Nota(2) Then PictPrint_Nota_Materia PictLibreta, IR, PosLinea, .Fields("PQBim2"), Si_No
                   IR = IR + JR
                   Sumatoria = .Fields("PQBim1") + .Fields("PQBim2")
                   If Print_Nota(3) Then PictPrint_Nota_Materia PictLibreta, IR, PosLinea, Sumatoria, Si_No
                   IR = IR + JR
                   If Print_Nota(4) Then PictPrint_Nota_Materia PictLibreta, IR, PosLinea, .Fields("PromPQ"), Si_No
                   IR = IR + JR
                   If Print_Nota(5) Then PictPrint_Nota_Materia PictLibreta, IR, PosLinea, .Fields("SQBim1"), Si_No
                   IR = IR + JR
                   If Print_Nota(6) Then PictPrint_Nota_Materia PictLibreta, IR, PosLinea, .Fields("SQBim2"), Si_No
                   IR = IR + JR
                   Sumatoria = .Fields("SQBim1") + .Fields("SQBim2")
                   If Print_Nota(7) Then PictPrint_Nota_Materia PictLibreta, IR, PosLinea, Sumatoria, Si_No
                   IR = IR + JR
                   If Print_Nota(8) Then PictPrint_Nota_Materia PictLibreta, IR, PosLinea, .Fields("PromSQ"), Si_No
                   IR = IR + JR
                   If Print_Nota(9) Then PictPrint_Nota_Materia PictLibreta, IR, PosLinea, .Fields("Supletorio"), Si_No
                   IR = IR + JR
                   If Print_Nota(10) Then PictPrint_Nota_Materia PictLibreta, IR, PosLinea, .Fields("PromFinal"), Si_No
                   IR = IR + JR
                Else
                   PictPrint_Nota_Materia PictLibreta, IR, PosLinea, .Fields("PQBim1"), Si_No
                   IR = IR + JR
                   PictPrint_Nota_Materia PictLibreta, IR, PosLinea, .Fields("PQBim2"), Si_No
                   IR = IR + JR
                   PictPrint_Nota_Materia PictLibreta, IR, PosLinea, .Fields("SQBim1"), Si_No
                   IR = IR + JR
                   PictPrint_Nota_Materia PictLibreta, IR, PosLinea, .Fields("SQBim2"), Si_No
                   IR = IR + JR
                End If
                'MsgBox .Fields("CodMat")
                If .Fields("Orden") <> 9 And .Fields("CodMatP") = Ninguno Then
                   'MsgBox .Fields("Materia") & vbCrLf & Si_No
                    If Not Si_No Then
                        VPQBim1 = VPQBim1 + .Fields("PQBim1")
                        VPQBim2 = VPQBim2 + .Fields("PQBim2")
                        VSQBim1 = VSQBim1 + .Fields("SQBim1")
                        VSQBim2 = VSQBim2 + .Fields("SQBim2")
                        VPromPQ = VPromPQ + .Fields("PromPQ")
                        VPromSQ = VPromSQ + .Fields("PromSQ")
                        VPromFinal = VPromFinal + .Fields("PromFinal")
                        If .Fields("PQBim1") > 0 Then TotalRegs(1) = TotalRegs(1) + 1
                        If .Fields("PQBim2") > 0 Then TotalRegs(2) = TotalRegs(2) + 1
                        If .Fields("PromPQ") > 0 Then TotalRegs(4) = TotalRegs(4) + 1
                        If .Fields("SQBim1") > 0 Then TotalRegs(5) = TotalRegs(5) + 1
                        If .Fields("SQBim2") > 0 Then TotalRegs(6) = TotalRegs(6) + 1
                        If .Fields("PromSQ") > 0 Then TotalRegs(8) = TotalRegs(8) + 1
                        If .Fields("PromFinal") > 0 Then TotalRegs(10) = TotalRegs(10) + 1
'''                        MsgBox .Fields("Materia") & vbCrLf & VPQBim1 & vbCrLf _
'''                             & VPQBim2 & vbCrLf _
'''                             & VSQBim1 & vbCrLf _
'''                             & VSQBim2 & vbCrLf _
'''                             & VPromPQ & vbCrLf _
'''                             & VPromSQ & vbCrLf _
'''                             & VPromFinal
                    End If
                End If
                PosLinea = PosLinea + 0.35
            End If
         End If
        .MoveNext
      Loop
      RatonNormal
      Cuadricula = False
      MensajeEncabData = "": SQLMsg1 = "": SQLMsg2 = "": SQLMsg3 = "": SQLMsg4 = ""
   End If
 End With
End Sub

Public Sub Libreta_Del_Alumno_Periodos(AdoLib As Adodc)
Dim AnchoDib As Single
Dim AltoDib As Single
Dim Curso As String
Dim Alumno As String
Dim Paralelo As String
Dim PosXPict As Single
Dim AnchoPict() As CtasAsiento
Dim Y0 As Single
Dim y1 As Single
Dim X0 As Single
Dim x1 As Single
Dim PosLineaX As Single
Dim TotalRegs(16) As Integer
Dim CanPromFinal As Byte

  PosLinea = 5.5
  With AdoLib.Recordset
   If .RecordCount > 0 Then
       Curso = .Fields("Curso")
       Paralelo = .Fields("Paralelo")
       Alumno = .Fields("Alumno")
       NombreCliente = .Fields("Alumno")
       Do While Not .EOF
          PosLinea = PosLinea + 0.36
         .MoveNext
       Loop
      .MoveFirst
   End If
  End With
  AltoDib = PosLinea
  PictLibreta.FontName = TipoTimes
  PictLibreta.ForeColor = QBColor(Negro)
  PosColumna = 7.5
  JR = 0.85
  Select Case Mid$(Curso, 1, 4)
    Case "0.00" To "1.01" '
         AnchoDib = 20
         'AltoDib = 13
         Grafico_Kinder AltoDib
         If Len(LogoTipo) > 1 Then PictLibreta.PaintPicture LoadPicture(LogoTipo), 0.1, 0.1, 2, 1
         PictLibreta.FontSize = 15
         PictPrint_Texto PictLibreta, 4, 0.1, Empresa
         PictLibreta.FontSize = 10
         PictPrint_Texto PictLibreta, 4, 0.7, "FICHA DE DESARROLLO DE DESTREZAS Y HABILIDADES"
         
         PictPrint_Texto PictLibreta, 14, 0.2, "NIVEL:"
         PictPrint_Texto PictLibreta, 0.5, 1.8, "ALUMNO(A):"
         
         PictLibreta.FontBold = False
         PictLibreta.FontSize = 10
         PictPrint_Texto PictLibreta, 15.3, 0.2, Curso
         PictPrint_Texto PictLibreta, 14, 0.7, Paralelo
         PictPrint_Texto PictLibreta, 2.7, 1.8, Alumno
         PosColumna = 0
         PosLinea = 0
    Case "1.02" To "1.99"
         AnchoDib = 20.3
         'AltoDib = 11.3
         PictLibreta.Cls
         PictLibreta.FontSize = 8
         PFil = 7
        'Cuadro Externo
         PictLibreta.Line (0.3, PFil + 0.5)-(19.9, PFil + AltoDib - 2), QBColor(Negro), B
         PictLibreta.Line (0.3, PFil + 1.4)-(19.9, PFil + 3.6), QBColor(Negro), B
         PCol = PosColumna - 0.35
         PictLibreta.Line (0.3, PFil + 2.2)-(PCol, PFil + 3), QBColor(Negro), B
         PictLibreta.Line (PCol, PFil + 1.4)-(PCol, PFil + AltoDib - 2), QBColor(Negro)
         PictLibreta.Line (PCol, PFil + 1.9)-(17.35, PFil + 1.9), QBColor(Negro)
         PictLibreta.FontSize = 10
         PictPrint_Texto PictLibreta, PCol + 1, PFil + 1.45, "PRIMER QUIMESTRE"
         PictPrint_Texto PictLibreta, PCol + 6, PFil + 1.45, "SEGUNDO QUIMESTRE"
         PictLibreta.FontSize = 13
         PictPrint_Texto PictLibreta, 1.5, PFil + 3, "M A T E R I A S"
         For I = 1 To 15
            'Lineas Verticales de la Libreta
             If I = 7 Or I = 13 Or I = 14 Or I = 15 Then
                PictLibreta.Line (PCol, PFil + 1.4)-(PCol, PFil + AltoDib - 2), QBColor(Negro)
             Else
                PictLibreta.Line (PCol, PFil + 1.9)-(PCol, PFil + AltoDib - 2), QBColor(Negro)
             End If
            'Encabezados de la libreta
             Select Case I
               Case 1, 7: cPrint.printTextoAngulo PictLibreta, PCol + 0.35, PFil + 3.4, 90, 4, 10, "Bimensual 1"
               Case 2, 8: cPrint.printTextoAngulo PictLibreta, PCol + 0.35, PFil + 3.4, 90, 4, 10, "Bimensual 2"
               Case 3, 9: cPrint.printTextoAngulo PictLibreta, PCol + 0.35, PFil + 3.4, 90, 4, 10, "70%"
               Case 4: cPrint.printTextoAngulo PictLibreta, PCol + 0.35, PFil + 3.4, 90, 4, 10, "Examen 1Q"
               Case 10: cPrint.printTextoAngulo PictLibreta, PCol + 0.35, PFil + 3.4, 90, 4, 10, "Examen 2Q"
               Case 5, 11: cPrint.printTextoAngulo PictLibreta, PCol + 0.35, PFil + 3.4, 90, 4, 10, "30%"
               Case 6: cPrint.printTextoAngulo PictLibreta, PCol + 0.35, PFil + 3.4, 90, 4, 10, "Prom. 1Q"
               Case 12: cPrint.printTextoAngulo PictLibreta, PCol + 0.35, PFil + 3.4, 90, 4, 10, "Prom. 2Q"
               Case 15: cPrint.printTextoAngulo PictLibreta, PCol + 0.05, PFil + 3.4, 90, 5, 10, "Promedio"
                        cPrint.printTextoAngulo PictLibreta, PCol + 0.4, PFil + 3.4, 90, 5, 10, "Final"
             End Select
             If ("2.00" <= Codigo4) And (Codigo4 <= "3.99") Then
               Select Case I
                 Case 13: cPrint.printTextoAngulo PictLibreta, PCol + 0.35, PFil + 3.4, 90, 5, 10, "Promedio"
                 Case 14: cPrint.printTextoAngulo PictLibreta, PCol + 0.25, PFil + 3.4, 90, 4, 9, "SUPLETORIO"
               End Select
             End If
             PCol = PCol + JR
         Next I
         PictLibreta.FontSize = 14
         PictPrint_Texto PictLibreta, 8.2, 3.7, "SECCIÓN PRIMARIA"
         PictPrint_Texto PictLibreta, 6.9, 4.5, "BOLETIN DE CALIFICACIONES"
         PictPrint_Texto PictLibreta, 7.6, 5.2, "AÑO LECTIVO " & Anio_Lectivo
         If OpcPeriodo("PQBim1", LstPeriodos) Then Cadena = "1er. Parcial"
         If OpcPeriodo("PQ", LstPeriodos) Then Cadena = "2do. Parcial"
         If OpcPeriodo("SQBim1", LstPeriodos) Then Cadena = "3er. Parcial"
         If OpcPeriodo("SQ", LstPeriodos) Then Cadena = "4to. Parcial"
         If OpcPeriodo("PF", LstPeriodos) Then Cadena = "Periodo Final"
         PictLibreta.FontSize = 9
         PictPrint_Texto PictLibreta, 9.8, 6, Cadena
         PictPrint_Texto PictLibreta, 14, PFil, FechaStrgCiudad(MBFecha)
         PosLinea = 1.4
         PictLibreta.FontSize = 8
         PictPrint_Texto PictLibreta, 0.5, PFil + 0.61, "Alumno:"
         PictPrint_Texto PictLibreta, 0.5, PFil + 1.45, "Paralelo:"
         PictPrint_Texto PictLibreta, 0.5, PFil + 2.25, "Curso:"
         PictLibreta.FontBold = False
         PictPrint_Texto PictLibreta, 0.5, PFil + 1.8, Curso
         PictLibreta.FontSize = 10
         PictPrint_Texto PictLibreta, 0.5, PFil + 0.95, Alumno
         PictPrint_Texto PictLibreta, 0.5, PFil + 2.6, Paralelo
         PosLinea = 10.6
    Case "2.00" To "3.99"
         AnchoDib = 20.3
         'AltoDib = 11.5
         PictLibreta.Cls
         If LogoTipo <> "" Then PictLibreta.PaintPicture LoadPicture(LogoTipo), 0.45, 0.2, 2, 1
         PictLibreta.FontSize = 8
         PFil = AltoDib - 1.9
        'Cuadro Externo
         PictLibreta.Line (0.3, 0.1)-(19.9, AltoDib - 2), QBColor(Negro), B
         PictLibreta.Line (0.3, 1.4)-(19.9, 3.6), QBColor(Negro), B
         PCol = PosColumna - 0.35
         PictLibreta.Line (0.3, 2.2)-(PCol, 3), QBColor(Negro), B
         PictLibreta.Line (PCol, 1.4)-(PCol, AltoDib - 2), QBColor(Negro)
         PictLibreta.Line (PCol, 1.9)-(17.35, 1.9), QBColor(Negro)
         PictLibreta.FontSize = 10
         PictPrint_Texto PictLibreta, PCol + 1, 1.45, "PRIMER QUIMESTRE"
         PictPrint_Texto PictLibreta, PCol + 6, 1.45, "SEGUNDO QUIMESTRE"
         PictLibreta.FontSize = 13
         PictPrint_Texto PictLibreta, 1.5, 3, "M A T E R I A S"
         For I = 1 To 15
            'Lineas Verticales de la Libreta
             If I = 7 Or I = 13 Or I = 14 Or I = 15 Then
                PictLibreta.Line (PCol, 1.4)-(PCol, AltoDib - 2), QBColor(Negro)
             Else
                PictLibreta.Line (PCol, 1.9)-(PCol, AltoDib - 2), QBColor(Negro)
             End If
            'Encabezados de la libreta
             Select Case I
               Case 1, 7: cPrint.printTextoAngulo PictLibreta, PCol + 0.35, 3.4, 90, 4, 10, "Bimensual 1"
               Case 2, 8: cPrint.printTextoAngulo PictLibreta, PCol + 0.35, 3.4, 90, 4, 10, "Bimensual 2"
               Case 3, 9: cPrint.printTextoAngulo PictLibreta, PCol + 0.35, 3.4, 90, 4, 10, "70%"
               Case 4: cPrint.printTextoAngulo PictLibreta, PCol + 0.35, 3.4, 90, 4, 10, "Examen 1Q"
               Case 10: cPrint.printTextoAngulo PictLibreta, PCol + 0.35, 3.4, 90, 4, 10, "Examen 2Q"
               Case 5, 11: cPrint.printTextoAngulo PictLibreta, PCol + 0.35, 3.4, 90, 4, 10, "30%"
               Case 6: cPrint.printTextoAngulo PictLibreta, PCol + 0.35, 3.4, 90, 4, 10, "Prom. 1Q"
               Case 12: cPrint.printTextoAngulo PictLibreta, PCol + 0.35, 3.4, 90, 4, 10, "Prom. 2Q"
               Case 15: cPrint.printTextoAngulo PictLibreta, PCol + 0.05, 3.4, 90, 5, 10, "Promedio"
                        cPrint.printTextoAngulo PictLibreta, PCol + 0.4, 3.4, 90, 5, 10, "Final"
             End Select
             If ("2.00" <= Codigo4) And (Codigo4 <= "3.99") Then
               Select Case I
                 Case 13: cPrint.printTextoAngulo PictLibreta, PCol + 0.35, 3.4, 90, 5, 10, "Promedio"
                 Case 14: cPrint.printTextoAngulo PictLibreta, PCol + 0.25, 3.4, 90, 4, 9, "SUPLETORIO"
               End Select
             End If
             PCol = PCol + JR
         Next I
         'PictLibreta.Line (PFil, 2.3)-(PFil, 2.8), QBColor(Negro)
         PictLibreta.FontSize = 16
         PictPrint_Texto PictLibreta, 2.5, 0.1, Institucion1
         PictLibreta.FontSize = 18
         PictPrint_Texto PictLibreta, 2.5, 0.7, UCase(Institucion2)
         PictLibreta.FontSize = 11
         PictPrint_Texto PictLibreta, 13.5, 0.1, "BOLETIN DE CALIFICACIONES"
         PictLibreta.FontSize = 11
         PictPrint_Texto PictLibreta, 13.5, 0.5, "AÑO LECTIVO " & Anio_Lectivo
         If OpcPeriodo("PQBim1", LstPeriodos) Then Cadena = "1er. Parcial"
         If OpcPeriodo("PQ", LstPeriodos) Then Cadena = "2do. Parcial"
         If OpcPeriodo("SQBim1", LstPeriodos) Then Cadena = "3er. Parcial"
         If OpcPeriodo("SQ", LstPeriodos) Then Cadena = "4to. Parcial"
         If OpcPeriodo("PF", LstPeriodos) Then Cadena = "Periodo Final"
         PictLibreta.FontSize = 9
         'PictPrint_Texto PictLibreta, 13.5, 0.9, Cadena
         PictPrint_Texto PictLibreta, 13.5, 1, FechaStrgCiudad(MBFecha)
         PosLinea = 1.4
         PictLibreta.FontSize = 8
         PictPrint_Texto PictLibreta, 0.5, 1.45, "Alumno:"
         PictPrint_Texto PictLibreta, 0.5, 2.25, "Curso:"
         PictLibreta.FontBold = False
         PictPrint_Texto PictLibreta, 5.8, 2.25, Curso
         PictLibreta.FontSize = 10
         PictPrint_Texto PictLibreta, 0.5, 1.8, Alumno
         PictPrint_Texto PictLibreta, 0.5, 2.6, Paralelo
         PosLinea = 3.6
  End Select
  PosColumna = 7.5
 With AdoLib.Recordset
  If .RecordCount > 0 Then
      PictLibreta.FontSize = 9
      PictLibreta.FontName = TipoCourierNew
      Codigo = .Fields("Codigo")
      Cadena1 = .Fields("Curso")
      Cadena2 = .Fields("Alumno")
      Codigo4 = .Fields("CodE")
      Select Case Mid$(Codigo4, 1, 4)
        Case "0.00" To "1.01": y1 = 9.35
        Case "1.02" To "1.99": y1 = 16.4
        Case "2.00" To "3.99": y1 = 9.6
      End Select
      For I = 0 To 14
          TotalRegs(I) = 0
      Next I
      VPQBim1 = 0: VPQBim2 = 0: VSQBim1 = 0: VSQBim2 = 0
      VPromPQ = 0: VPromSQ = 0: VPromFinal = 0
      VExamenPQ = 0: VExamenSQ = 0
      Do While Not .EOF
         PictLibreta.FontSize = 9
         Opciones = .Fields("Orden")
         JR = 0.85
         If PosColumna > 0 Then
           'MsgBox PosColumna & vbCrLf & .Fields("Orden")
            If .Fields("Orden") = 9 Then
                PosLinea = y1
                If TotalReg = 0 Then TotalReg = 1
                For I = 0 To 14
                    If TotalRegs(I) = 0 Then TotalRegs(I) = 1
                Next I
                PictLibreta.FontBold = True
                PictPrint_Texto PictLibreta, 0.6, PosLinea, "T O T A L"
                PictLibreta.FontBold = False
                IR = PosColumna
               'IMPRIME SUMATORIA TOTAL DE NOTAS
                If ("1.02" <= Codigo4) And (Codigo4 <= "3.99") Then
                   If VPQBim1 > 0 And Print_Nota(1) Then PictPrint_Texto PictLibreta, IR - 0.1, PosLinea, Format(VPQBim1, "00")
                   IR = IR + JR
                   If VPQBim2 > 0 And Print_Nota(2) Then PictPrint_Texto PictLibreta, IR - 0.1, PosLinea, Format(VPQBim2, "00")
                   IR = IR + (JR * 2)
                   If VExamenPQ > 0 And Print_Nota(4) Then PictPrint_Texto PictLibreta, IR - 0.1, PosLinea, Format(VExamenPQ, "00")
                   IR = IR + (JR * 2)
                   If VPromPQ > 0 And Print_Nota(6) Then PictPrint_Texto PictLibreta, IR - 0.1, PosLinea, Format(VPromPQ, "00")
                   IR = IR + JR
                   If VSQBim1 > 0 And Print_Nota(7) Then PictPrint_Texto PictLibreta, IR - 0.1, PosLinea, Format(VSQBim1, "00")
                   IR = IR + JR
                   If VSQBim2 > 0 And Print_Nota(8) Then PictPrint_Texto PictLibreta, IR - 0.1, PosLinea, Format(VSQBim2, "00")
                   IR = IR + (JR * 2)
                   If VExamenSQ > 0 And Print_Nota(10) Then PictPrint_Texto PictLibreta, IR - 0.1, PosLinea, Format(VExamenSQ, "00")
                   IR = IR + (JR * 2)
                   If VPromSQ > 0 And Print_Nota(12) Then PictPrint_Texto PictLibreta, IR - 0.1, PosLinea, Format(VPromSQ, "00")
                   IR = IR + (JR * 3)
                   IR = IR - 0.2
                   If VPromFinal > 0 And Print_Nota(15) Then PictPrint_Texto PictLibreta, IR, PosLinea, Format(VPromFinal, "00")
                End If
                PosLinea = PosLinea + 0.35
                PictLibreta.FontBold = True
                PictPrint_Texto PictLibreta, 0.6, PosLinea, "PROMEDIO DE RENDIMIENTO"
                PosLinea = PosLinea + 0.05
                PictLibreta.FontBold = False
                PictLibreta.FontSize = 7.5
                IR = PosColumna - 0.05
               'IMPRIME TOTAL PROMEDIOS
                If ("1.02" <= Codigo4) And (Codigo4 <= "3.99") Then
                   If VPQBim1 > 0 Then PictPrint_Texto PictLibreta, IR - 0.2, PosLinea, Format(VPQBim1 / TotalRegs(1), "00.00")
                   IR = IR + JR
                   If VPQBim2 > 0 Then PictPrint_Texto PictLibreta, IR - 0.2, PosLinea, Format(VPQBim2 / TotalRegs(2), "00.00")
                   IR = IR + (JR * 2)
                   If VExamenPQ > 0 Then PictPrint_Texto PictLibreta, IR - 0.1, PosLinea, Format(VExamenPQ / TotalRegs(4), "00.00")
                   IR = IR + (JR * 2)
                   If VPromPQ > 0 Then PictPrint_Texto PictLibreta, IR - 0.2, PosLinea, Format(VPromPQ / TotalRegs(6), "00.00")
                   IR = IR + JR
                   If VSQBim1 > 0 Then PictPrint_Texto PictLibreta, IR - 0.2, PosLinea, Format(VSQBim1 / TotalRegs(7), "00.00")
                   IR = IR + JR
                   If VSQBim2 > 0 Then PictPrint_Texto PictLibreta, IR - 0.2, PosLinea, Format(VSQBim2 / TotalRegs(8), "00.00")
                   IR = IR + (JR * 2)
                   If VExamenSQ > 0 Then PictPrint_Texto PictLibreta, IR - 0.1, PosLinea, Format(VExamenSQ / TotalRegs(10), "00.00")
                   IR = IR + (JR * 2)
                   If VPromSQ > 0 Then PictPrint_Texto PictLibreta, IR - 0.2, PosLinea, Format(VPromSQ / TotalRegs(12), "00.00")
                   IR = IR + (JR * 3)
                   IR = IR - 0.2
                   If VPromFinal > 0 And OpcPeriodo("PF", LstPeriodos) Then PictPrint_Texto PictLibreta, IR, PosLinea, Format(VPromFinal / TotalRegs(13), "00.00")
                End If
                PosLinea = PosLinea + 0.35
                PictLibreta.FontBold = True
                PictLibreta.FontSize = 9
                PictPrint_Texto PictLibreta, 0.6, PosLinea, UCase(.Fields("Materia"))
                PictLibreta.FontBold = False
                IR = PosColumna
               'IMPRIMIR CONDUCTA
                If ("1.02" <= Codigo4) And (Codigo4 <= "3.99") Then
                   'PictPrint_Nota_Materia PictLibreta, IR, PosLinea, .Fields("PQBim1")
                   'IR = IR + JR
                   'PictPrint_Nota_Materia PictLibreta, IR, PosLinea, .Fields("PQBim2")
                   IR = IR + (JR * 5)
                   'Sumatoria = .Fields("PQBim1") + .Fields("PQBim2") / 2
                   PictPrint_Nota_Materia PictLibreta, IR, PosLinea, .Fields("PromPQ")  'Sumatoria
                   IR = IR + (JR * 6)
                   'PictPrint_Nota_Materia PictLibreta, IR, PosLinea, .Fields("ConductaPQ1")
                   'IR = IR + JR
                   'PictPrint_Nota_Materia PictLibreta, IR, PosLinea, .Fields("SQBim1")
                   'IR = IR + JR
                   'PictPrint_Nota_Materia PictLibreta, IR, PosLinea, .Fields("SQBim2")
                   'IR = IR + (JR * 3)
                   'Sumatoria = .Fields("SQBim1") + .Fields("SQBim2") / 2
                   PictPrint_Nota_Materia PictLibreta, IR, PosLinea, .Fields("PromSQ")  'Sumatoria
                   IR = IR + (JR * 3)
                   'PictPrint_Nota_Materia PictLibreta, IR, PosLinea, .Fields("ConductaSQ1")
                   'IR = IR + JR
                   'PictPrint_Nota_Materia PictLibreta, IR, PosLinea, .Fields("Supletorio")
                   'IR = IR + (JR * 2)
                   If .Fields("PromPQ") > 0 And .Fields("PromSQ") > 0 Then
                      'Diferencia = Redondear((.Fields("ConductaPQ1") + .Fields("ConductaSQ1")) / 2)
                       Diferencia = Redondear((.Fields("PromPQ") + .Fields("PromSQ")) / 2)
                       PictPrint_Nota_Materia PictLibreta, IR, PosLinea, Diferencia     ' .Fields("PromFinal")
                   End If
                   IR = IR + JR
                   ' MsgBox .Fields("PromFinal")
                End If
                PictLibreta.FontBold = False
            Else
                Si_No = False
                If .Fields("C") <> 0 Then Si_No = True
                'MsgBox .Fields("C") & vbCrLf & Si_No
               'IMPRESION DE LAS NOTAS DE LAS MATERIAS
                PictLibreta.FontBold = False
                PictLibreta.FontUnderline = False
                Contador = Contador + 1
                If .Fields("CodMatP") <> Ninguno Then
                    PictPrint_Texto PictLibreta, 1.5, PosLinea, .Fields("Materia")
                Else
                    PictPrint_Texto PictLibreta, 0.5, PosLinea, .Fields("Materia")
                End If
                IR = PosColumna
                If ("1.02" <= Codigo4) And (Codigo4 <= "3.99") Then
                   If Print_Nota(1) Then PictPrint_Nota_Materia PictLibreta, IR, PosLinea, .Fields("PQBim1"), Si_No
                   IR = IR + JR
                   If Print_Nota(2) Then PictPrint_Nota_Materia PictLibreta, IR, PosLinea, .Fields("PQBim2"), Si_No
                   IR = IR + JR
                   If Print_Nota(3) And .Fields("PQ_PP") > 0 Then PictPrint_Texto PictLibreta, IR, PosLinea, Format(.Fields("PQ_PP"), "00")
                   IR = IR + JR
                   If Print_Nota(4) Then PictPrint_Nota_Materia PictLibreta, IR, PosLinea, .Fields("ExamenPQ"), Si_No
                   IR = IR + JR
                   If Print_Nota(5) And .Fields("PQ_PE") > 0 Then PictPrint_Texto PictLibreta, IR, PosLinea, Format(.Fields("PQ_PE"), "00")
                   IR = IR + JR
                   If Print_Nota(6) Then PictPrint_Nota_Materia PictLibreta, IR, PosLinea, .Fields("PromPQ"), Si_No
                   IR = IR + JR
                   If Print_Nota(7) Then PictPrint_Nota_Materia PictLibreta, IR, PosLinea, .Fields("SQBim1"), Si_No
                   IR = IR + JR
                   If Print_Nota(8) Then PictPrint_Nota_Materia PictLibreta, IR, PosLinea, .Fields("SQBim2"), Si_No
                   IR = IR + JR
                   If Print_Nota(9) And .Fields("SQ_PP") > 0 Then PictPrint_Texto PictLibreta, IR, PosLinea, Format(.Fields("SQ_PP"), "00")
                   IR = IR + JR
                   If Print_Nota(10) Then PictPrint_Nota_Materia PictLibreta, IR, PosLinea, .Fields("ExamenSQ"), Si_No
                   IR = IR + JR
                   If Print_Nota(11) And .Fields("SQ_PE") > 0 Then PictPrint_Texto PictLibreta, IR, PosLinea, Format(.Fields("SQ_PE"), "00")
                   IR = IR + JR
                   If Print_Nota(12) Then PictPrint_Nota_Materia PictLibreta, IR, PosLinea, .Fields("PromSQ"), Si_No
                   IR = IR + JR
                   If ("2.00" <= Codigo4) And (Codigo4 <= "3.99") Then
                      Sumatoria = (.Fields("PromPQ") + .Fields("PromSQ")) / 2
                      If Print_Nota(13) Then PictPrint_Nota_Materia PictLibreta, IR, PosLinea, Sumatoria, Si_No
                   End If
                   IR = IR + JR
                   If ("2.00" <= Codigo4) And (Codigo4 <= "3.99") Then
                      If Print_Nota(14) Then PictPrint_Nota_Materia PictLibreta, IR, PosLinea, .Fields("Supletorio"), Si_No
                   End If
                   IR = IR + JR
                   If Print_Nota(15) Then PictPrint_Nota_Materia PictLibreta, IR, PosLinea, .Fields("PromFinal"), Si_No
                   IR = IR + JR
                End If
                If .Fields("Orden") <> 9 And .Fields("I") <> False And .Fields("CodMatP") = Ninguno Then
                    If Not Si_No Then
                       VPQBim1 = VPQBim1 + .Fields("PQBim1")
                       VPQBim2 = VPQBim2 + .Fields("PQBim2")
                       VSQBim1 = VSQBim1 + .Fields("SQBim1")
                       VSQBim2 = VSQBim2 + .Fields("SQBim2")
                       VPromPQ = VPromPQ + .Fields("PromPQ")
                       VPromSQ = VPromSQ + .Fields("PromSQ")
                       VExamenPQ = VExamenPQ + .Fields("ExamenPQ")
                       VExamenSQ = VExamenSQ + .Fields("ExamenSQ")
                       VPromFinal = VPromFinal + .Fields("PromFinal")
                       If .Fields("PQBim1") > 0 Then TotalRegs(1) = TotalRegs(1) + 1
                       If .Fields("PQBim2") > 0 Then TotalRegs(2) = TotalRegs(2) + 1
                       If .Fields("ExamenPQ") > 0 Then TotalRegs(4) = TotalRegs(4) + 1
                       If .Fields("PromPQ") > 0 Then TotalRegs(6) = TotalRegs(6) + 1
                       If .Fields("SQBim1") > 0 Then TotalRegs(7) = TotalRegs(7) + 1
                       If .Fields("SQBim2") > 0 Then TotalRegs(8) = TotalRegs(8) + 1
                       If .Fields("ExamenSQ") > 0 Then TotalRegs(10) = TotalRegs(10) + 1
                       If .Fields("PromSQ") > 0 Then TotalRegs(12) = TotalRegs(12) + 1
                       If .Fields("PromFinal") > 0 Then TotalRegs(13) = TotalRegs(13) + 1
                    End If
                End If
                PosLinea = PosLinea + 0.36
            End If
         End If
        .MoveNext
      Loop
     'Faltas justificadas o atraso
      PictLibreta.FontSize = 9
      PosLinea = PosLinea + 0.4
      PictPrint_Texto PictLibreta, 0.6, PosLinea, "FALTAS JUSTIFICADAS"
      PictPrint_Texto PictLibreta, 5, PosLinea, Format(Faltas_Just, "00")
      
      PictPrint_Texto PictLibreta, 7.5, PosLinea, "FALTAS INJUSTIFICADAS"
      PictPrint_Texto PictLibreta, 12, PosLinea, Format(Faltas_Injust, "00")

      PictPrint_Texto PictLibreta, 16, PosLinea, "ATRASOS"
      PictPrint_Texto PictLibreta, 18, PosLinea, Format(Atrasos, "00")
      PosLinea = PosLinea + 0.4
      If ("1.02" <= Codigo4) And (Codigo4 <= "1.99") Then
         PFil = PosLinea
         PFil = PFil + 0.9
         PictLibreta.Line (0.3, PFil)-(19.9, PFil), QBColor(Negro)
         PFil = PFil + 0.05
         PictLibreta.FontBold = True
         PictLibreta.FontUnderline = True
         PictPrint_Texto PictLibreta, 0.6, PFil, "NOMENCLATURA"
         PictPrint_Texto PictLibreta, 5.5, PFil, "FELICITACIONES POR:"
         PictPrint_Texto PictLibreta, 13, PFil, "PUEDE MEJORAR:"
         PictLibreta.FontUnderline = False
         PictLibreta.FontBold = False
         PFil = PFil + 0.05
         PictPrint_Texto PictLibreta, 0.6, PFil + 0.4, "IN INCLUSION"
         PictPrint_Texto PictLibreta, 0.6, PFil + 0.75, "I  INSUFICIENTE"
         PictPrint_Texto PictLibreta, 0.6, PFil + 1.1, "B  BUENA"
         PictPrint_Texto PictLibreta, 0.6, PFil + 1.45, "R  REGULAR"
         PictPrint_Texto PictLibreta, 0.6, PFil + 1.8, "M  MUY BUENA"
         PictPrint_Texto PictLibreta, 0.6, PFil + 2.15, "S  SOBRESALIENTE"
         PFil = PFil + 0.6
         For I = 1 To 4
             PictPrint_Texto PictLibreta, 5.5, PFil, String(36, "_")
             PictPrint_Texto PictLibreta, 13, PFil, String(36, "_")
             PFil = PFil + 0.6
         Next I
         PFil = PFil + 2
         PictLibreta.FontBold = True
         PictLibreta.FontUnderline = False
         PictPrint_Texto PictLibreta, 4, PFil, String(17, "_")
         PictPrint_Texto PictLibreta, 10, PFil, String(28, "_")
         PFil = PFil + 0.4
         PictPrint_Texto PictLibreta, 4.5, PFil, "PROFESOR(A)"
         PictPrint_Texto PictLibreta, 10.5, PFil, "FIRMA DEL REPRESENTANTE"
         PictLibreta.FontBold = False
      End If
      If ("2.00" <= Codigo4) And (Codigo4 <= "3.99") Then
        'Lineas de Observacion
         Cuadricula = False
         PictLibreta.FontBold = True
         PictLibreta.FontUnderline = True
         PictLibreta.FontName = TipoArialNarrow
         PictPrint_Texto PictLibreta, 0.6, PosLinea, "OBSERVACIONES:"
         PictLibreta.FontUnderline = False
         PictPrint_Texto PictLibreta, 3.1, PosLinea, String(121, "_")
         PosLinea = PosLinea + 1.2
         PictLibreta.FontBold = True
         PictLibreta.FontUnderline = False
         PictPrint_Texto PictLibreta, 3, PosLinea, LblDirigente.Caption
         PosLinea = PosLinea + 0.1
         PictPrint_Texto PictLibreta, 3, PosLinea, String(35, "_")
         PictPrint_Texto PictLibreta, 9.5, PosLinea, String(35, "_")
         PosLinea = PosLinea + 0.4
         PictPrint_Texto PictLibreta, 3.5, PosLinea, "DIRIGENTE DEL CURSO"
         PictPrint_Texto PictLibreta, 10, PosLinea, "FIRMA DEL REPRESENTANTE"
         PictLibreta.FontBold = False
      End If
      RatonNormal
      Cuadricula = False
      MensajeEncabData = "": SQLMsg1 = "": SQLMsg2 = "": SQLMsg3 = "": SQLMsg4 = ""
   End If
 End With
End Sub

Public Sub Libreta_Del_Alumno_Trimestre1(AdoLib As Adodc)
Dim AnchoDib As Single
Dim AltoDib As Single
Dim Curso As String
Dim Alumno As String
Dim Paralelo As String
Dim PosXPict As Single
Dim AnchoPict() As CtasAsiento
Dim Y0 As Single
Dim y1 As Single
Dim X0 As Single
Dim x1 As Single
Dim PosLineaX As Single
Dim TotalRegs(18) As Integer
Dim CanPromFinal As Byte
Dim Formato_Nota As String
  PosLinea = 5.5
  With AdoLib.Recordset
   If .RecordCount > 0 Then
       Curso = .Fields("Curso")
       Paralelo = .Fields("Paralelo")
       Alumno = .Fields("Alumno")
       NombreCliente = .Fields("Alumno")
       Do While Not .EOF
          PosLinea = PosLinea + 0.36
         .MoveNext
       Loop
      .MoveFirst
   End If
  End With
  AltoDib = PosLinea
  PictLibreta.FontName = TipoTimes
  PictLibreta.ForeColor = QBColor(Negro)
  PosColumna = 8.2
  JR = 0.85
  Select Case Mid$(Curso, 1, 4)
    Case "0.00" To "1.01" '
         AnchoDib = 20
         'AltoDib = 13
         Grafico_Kinder AltoDib
         PictLibreta.PaintPicture LoadPicture(LogoTipo), 0.1, 0.1, 2, 1
         PictLibreta.FontSize = 15
         PictPrint_Texto PictLibreta, 4, 0.1, Empresa
         PictLibreta.FontSize = 10
         PictPrint_Texto PictLibreta, 4, 0.7, "FICHA DE DESARROLLO DE DESTREZAS Y HABILIDADES"
         
         PictPrint_Texto PictLibreta, 14, 0.2, "NIVEL:"
         PictPrint_Texto PictLibreta, 0.5, 1.8, "ALUMNO(A):"
         
         PictLibreta.FontBold = False
         PictLibreta.FontSize = 9
         PictLibreta.FontName = TipoArialNarrow
         PictPrint_Texto PictLibreta, 15.3, 0.2, Curso
         PictPrint_Texto PictLibreta, 14, 0.7, Paralelo
         PictPrint_Texto PictLibreta, 2.7, 1.8, Alumno
         PosColumna = 0
         PosLinea = 0
    Case "1.02" To "3.99"
         AnchoDib = 20.3
         PictLibreta.Cls
         If LogoTipo <> "" Then PictLibreta.PaintPicture LoadPicture(LogoTipo), 1, 0.2, 2, 1
         PictLibreta.FontSize = 8
         PFil = 0.1
        'Cuadro Externo
         PictLibreta.Line (1, 0.1)-(19.9, AltoDib - 2), QBColor(Negro), B
         PictLibreta.Line (1, 1.4)-(19.9, 3.6), QBColor(Negro), B
         PCol = PosColumna - 0.35
         PictLibreta.Line (1, 2.2)-(PCol, 3), QBColor(Negro), B
         PictLibreta.Line (PCol, 1.4)-(PCol, AltoDib - 2), QBColor(Negro)
         PictLibreta.Line (PCol, 2.2)-(18.05, 2.2), QBColor(Negro)
         PictLibreta.FontSize = 8
         PictPrint_Texto PictLibreta, PCol + 0.5, PFil + 1.4, "PRIMER"
         PictPrint_Texto PictLibreta, PCol + 0.3, PFil + 1.7, "TRIMESTRE"
         PictPrint_Texto PictLibreta, PCol + 3.55, PFil + 1.4, "SEGUNDO"
         PictPrint_Texto PictLibreta, PCol + 3.5, PFil + 1.7, "TRIMESTRE"
         PictPrint_Texto PictLibreta, PCol + 7.3, PFil + 1.4, "TERCER"
         PictPrint_Texto PictLibreta, PCol + 7.1, PFil + 1.7, "TRIMESTRE"

         PictLibreta.FontSize = 11
         PictPrint_Texto PictLibreta, 3, 3.1, "M A T E R I A S"
         PictLibreta.FontSize = 10
         For I = 1 To 14
            'Lineas Verticales de la Libreta
             If I = 5 Or I = 9 Or I = 13 Or I = 14 Then
                PictLibreta.Line (PCol, 1.4)-(PCol, AltoDib - 2), QBColor(Negro)
             Else
                PictLibreta.Line (PCol, 2.2)-(PCol, AltoDib - 2), QBColor(Negro)
             End If
            'Encabezados de la libreta
             Select Case I
               Case 1, 5, 9: cPrint.printTextoAngulo PictLibreta, PCol + 0.05, PFil + 3.4, 90, 4, 10, "Primer"
                             cPrint.printTextoAngulo PictLibreta, PCol + 0.4, PFil + 3.4, 90, 4, 10, "Parcial"
               Case 2, 6, 10: cPrint.printTextoAngulo PictLibreta, PCol + 0.05, PFil + 3.4, 90, 4, 10, "Segundo"
                              cPrint.printTextoAngulo PictLibreta, PCol + 0.4, PFil + 3.4, 90, 4, 10, "Parcial"
               Case 3, 7, 11
                              If FormatoLibreta = "TRIMESTRE2" Then
                                 cPrint.printTextoAngulo PictLibreta, PCol + 0.35, PFil + 3.3, 90, 4, 10, "Examen"
                              End If
               Case 4: cPrint.printTextoAngulo PictLibreta, PCol + 0.05, PFil + 3.4, 90, 4, 10, "Prom."
                       cPrint.printTextoAngulo PictLibreta, PCol + 0.4, PFil + 3.4, 90, 4, 10, "1er. T."
               Case 8: cPrint.printTextoAngulo PictLibreta, PCol + 0.05, PFil + 3.4, 90, 4, 10, "Prom."
                       cPrint.printTextoAngulo PictLibreta, PCol + 0.4, PFil + 3.4, 90, 4, 10, "2do. T."
               Case 12: cPrint.printTextoAngulo PictLibreta, PCol + 0.05, PFil + 3.4, 90, 4, 10, "Prom."
                        cPrint.printTextoAngulo PictLibreta, PCol + 0.4, PFil + 3.4, 90, 4, 10, "3er. T."
               Case 14: cPrint.printTextoAngulo PictLibreta, PCol + 0.1, PFil + 3.4, 90, 5, 10, "Promedio"
                        cPrint.printTextoAngulo PictLibreta, PCol + 0.45, PFil + 3.4, 90, 5, 10, "Final"
             End Select
             If ("2.00" <= Codigo4) And (Codigo4 <= "3.99") Then
               Select Case I
                 Case 13: cPrint.printTextoAngulo PictLibreta, PCol + 0.25, PFil + 3.4, 90, 4, 9, "SUPLETORIO"
                 Case 14: cPrint.printTextoAngulo PictLibreta, PCol + 0.1, PFil + 3.4, 90, 5, 10, "Promedio"
                          cPrint.printTextoAngulo PictLibreta, PCol + 0.45, PFil + 3.4, 90, 5, 10, "Final"
               End Select
             End If
             PCol = PCol + JR
         Next I
         'PictLibreta.Line (PFil, 2.3)-(PFil, 2.8), QBColor(Negro)
         PictLibreta.FontSize = 16
         PictPrint_Texto PictLibreta, 3, 0.1, Institucion1
         PictPrint_Texto PictLibreta, 3, 0.7, UCase(Institucion2)
         PictLibreta.FontSize = 11
         PictPrint_Texto PictLibreta, 13.5, 0.1, "BOLETÍN DE EVALUACIÓN"
         PictLibreta.FontSize = 11
         PictPrint_Texto PictLibreta, 13.5, 0.5, "AÑO LECTIVO " & Anio_Lectivo
         If OpcPeriodo("PQBim1", LstPeriodos) Then Cadena = "1er. Parcial"
         If OpcPeriodo("PQ", LstPeriodos) Then Cadena = "2do. Parcial"
         If OpcPeriodo("SQBim1", LstPeriodos) Then Cadena = "3er. Parcial"
         If OpcPeriodo("SQ", LstPeriodos) Then Cadena = "4to. Parcial"
         If OpcPeriodo("TQBim1", LstPeriodos) Then Cadena = "5er. Parcial"
         If OpcPeriodo("PF", LstPeriodos) Then Cadena = "Periodo Final"
         PictLibreta.FontSize = 9
         'PictPrint_Texto PictLibreta, 13.5, 0.9, Cadena
         PictPrint_Texto PictLibreta, 13.5, 1, FechaStrgCiudad(MBFecha)
         PosLinea = 1.4
         PictLibreta.FontSize = 8
         PictPrint_Texto PictLibreta, 1.2, 1.45, "Estudiante:"
         PictPrint_Texto PictLibreta, 1.2, 2.25, "Curso:"
         PictLibreta.FontBold = False
         PictLibreta.FontName = TipoArialNarrow
         PictPrint_Texto PictLibreta, 6.5, 2.25, Curso
         PictPrint_Texto PictLibreta, 1.2, 1.8, Alumno
         PictPrint_Texto PictLibreta, 1.2, 2.6, Paralelo
         PosLinea = 3.6
  End Select
  If Dec_Nota > 0 Then PosColumna = 8 Else PosColumna = 8.2
 With AdoLib.Recordset
  If .RecordCount > 0 Then
      If Dec_Nota > 0 Then
         PictLibreta.FontSize = 9
         Formato_Nota = "00." & String(Dec_Nota, "0")
         PictLibreta.FontName = TipoArialNarrow
      Else
         PictLibreta.FontSize = 10
         Formato_Nota = "00"
         PictLibreta.FontName = TipoCourierNew
      End If
      Codigo = .Fields("Codigo")
      Cadena1 = .Fields("Curso")
      Cadena2 = .Fields("Alumno")
      Codigo4 = .Fields("CodE")
      Select Case Mid$(Codigo4, 1, 4)
        Case "0.00" To "1.01": y1 = 0  '9.35
        Case "1.02" To "3.99": y1 = 9.6
      End Select
      For I = 0 To 16
          TotalRegs(I) = 0
      Next I
      VPQBim1 = 0: VPQBim2 = 0: VSQBim1 = 0: VSQBim2 = 0: VTQBim1 = 0: VTQBim2 = 0
      VPromPQ = 0: VPromSQ = 0: VPromTQ = 0: VPromFinal = 0
      VExamenPQ = 0: VExamenSQ = 0
      Do While Not .EOF
         PictLibreta.FontSize = 8
         Opciones = .Fields("Orden")
         JR = 0.85
         If PosColumna > 0 Then
           'MsgBox PosColumna & vbCrLf & .Fields("Orden")
            If .Fields("Orden") = 9 Then
               'PosLinea = y1
                PosLinea = PosLinea + 0.36
                If TotalReg = 0 Then TotalReg = 1
                For I = 0 To 16
                    If TotalRegs(I) = 0 Then TotalRegs(I) = 1
                Next I
                PictLibreta.FontSize = 9
                PictLibreta.FontBold = True
                PictPrint_Texto PictLibreta, 1.1, PosLinea, "T O T A L"
                PictLibreta.FontBold = False
                If Dec_Nota = 0 Then
                   PictLibreta.FontSize = 9
                   IR = PosColumna - 0.2
                Else
                   PictLibreta.FontSize = 7
                   IR = PosColumna - 0.1
                End If
               'MsgBox VPQBim1
               'IMPRIME SUMATORIA TOTAL DE NOTAS
                If ("1.02" <= Codigo4) And (Codigo4 <= "3.99") Then
                   If VPQBim1 > 0 And Print_Nota(1) Then PictPrint_Texto PictLibreta, IR, PosLinea, Format(VPQBim1, Formato_Nota)
                   IR = IR + JR
                   If VPQBim2 > 0 And Print_Nota(2) Then PictPrint_Texto PictLibreta, IR, PosLinea, Format(VPQBim2, Formato_Nota)
                   IR = IR + JR
                   If VExamenPQ > 0 And Print_Nota(3) Then PictPrint_Texto PictLibreta, IR, PosLinea, Format(VExamenPQ, Formato_Nota)
                   IR = IR + JR
                   If VPromPQ > 0 And Print_Nota(4) Then PictPrint_Texto PictLibreta, IR, PosLinea, Format(VPromPQ, Formato_Nota)
                   IR = IR + JR
                   If VSQBim1 > 0 And Print_Nota(5) Then PictPrint_Texto PictLibreta, IR, PosLinea, Format(VSQBim1, Formato_Nota)
                   IR = IR + JR
                   If VSQBim2 > 0 And Print_Nota(6) Then PictPrint_Texto PictLibreta, IR, PosLinea, Format(VSQBim2, Formato_Nota)
                   IR = IR + JR
                   If VExamenSQ > 0 And Print_Nota(7) Then PictPrint_Texto PictLibreta, IR, PosLinea, Format(VExamenSQ, Formato_Nota)
                   IR = IR + JR
                   If VPromSQ > 0 And Print_Nota(8) Then PictPrint_Texto PictLibreta, IR, PosLinea, Format(VPromSQ, Formato_Nota)
                   IR = IR + JR
                   If VSQBim1 > 0 And Print_Nota(9) Then PictPrint_Texto PictLibreta, IR, PosLinea, Format(VTQBim1, Formato_Nota)
                   IR = IR + JR
                   If VSQBim2 > 0 And Print_Nota(10) Then PictPrint_Texto PictLibreta, IR, PosLinea, Format(VTQBim2, Formato_Nota)
                   IR = IR + JR
                   If VExamenSQ > 0 And Print_Nota(11) Then PictPrint_Texto PictLibreta, IR, PosLinea, Format(VExamenTQ, Formato_Nota)
                   IR = IR + JR
                   If VPromSQ > 0 And Print_Nota(12) Then PictPrint_Texto PictLibreta, IR, PosLinea, Format(VPromTQ, Formato_Nota)
                   IR = IR + (JR * 2)
                   If VPromFinal > 0 And Print_Nota(14) Then PictPrint_Texto PictLibreta, IR, PosLinea, Format(VPromFinal, Formato_Nota)
                End If
                PosLinea = PosLinea + 0.35
                PictLibreta.FontBold = True
                PictLibreta.FontSize = 9
                PictPrint_Texto PictLibreta, 1.1, PosLinea, "DESEMPEÑO ACADÉMICO:"
                PosLinea = PosLinea + 0.05
                PictLibreta.FontBold = False
                PictLibreta.FontSize = 7
                IR = PosColumna - 0.1
               'IMPRIME TOTAL PROMEDIOS
                If ("1.02" <= Codigo4) And (Codigo4 <= "3.99") Then
                   If VPQBim1 > 0 Then PictPrint_Texto PictLibreta, IR, PosLinea, Format(VPQBim1 / TotalRegs(1), "00.00")
                   IR = IR + JR
                   If VPQBim2 > 0 Then PictPrint_Texto PictLibreta, IR, PosLinea, Format(VPQBim2 / TotalRegs(2), "00.00")
                   IR = IR + JR
                   If VExamenPQ > 0 Then PictPrint_Texto PictLibreta, IR, PosLinea, Format(VExamenPQ / TotalRegs(3), "00.00")
                   IR = IR + JR
                   If VPromPQ > 0 Then PictPrint_Texto PictLibreta, IR, PosLinea, Format(VPromPQ / TotalRegs(4), "00.00")
                   IR = IR + JR
                   If VSQBim1 > 0 Then PictPrint_Texto PictLibreta, IR, PosLinea, Format(VSQBim1 / TotalRegs(5), "00.00")
                   IR = IR + JR
                   If VSQBim2 > 0 Then PictPrint_Texto PictLibreta, IR, PosLinea, Format(VSQBim2 / TotalRegs(6), "00.00")
                   IR = IR + JR
                   If VExamenSQ > 0 Then PictPrint_Texto PictLibreta, IR, PosLinea, Format(VExamenSQ / TotalRegs(7), "00.00")
                   IR = IR + JR
                   If VPromSQ > 0 Then PictPrint_Texto PictLibreta, IR, PosLinea, Format(VPromSQ / TotalRegs(8), "00.00")
                   IR = IR + JR
                   If VTQBim1 > 0 Then PictPrint_Texto PictLibreta, IR, PosLinea, Format(VTQBim1 / TotalRegs(9), "00.00")
                   IR = IR + JR
                   If VTQBim2 > 0 Then PictPrint_Texto PictLibreta, IR, PosLinea, Format(VTQBim2 / TotalRegs(10), "00.00")
                   IR = IR + JR
                   If VExamenTQ > 0 Then PictPrint_Texto PictLibreta, IR, PosLinea, Format(VExamenTQ / TotalRegs(11), "00.00")
                   IR = IR + JR
                   If VPromTQ > 0 Then PictPrint_Texto PictLibreta, IR, PosLinea, Format(VPromTQ / TotalRegs(12), "00.00")
                   IR = IR + (JR * 2)
                   If VPromFinal > 0 And OpcPeriodo("PF", LstPeriodos) Then
                      If FormatoLibreta = "PERIODO" Then
                         PictPrint_Texto PictLibreta, IR, PosLinea, Format(VPromFinal / TotalRegs(14), "00.00")
                      Else
                         PictPrint_Texto PictLibreta, IR, PosLinea, Format(VPromFinal / TotalRegs(14), "00.000")
                      End If
                   End If
                End If
                PosLinea = PosLinea + 0.35
                PictLibreta.FontBold = True
                PictLibreta.FontSize = 9
                PictPrint_Texto PictLibreta, 1.1, PosLinea, "DISCIPLINA:"
                PictLibreta.FontBold = False
                IR = PosColumna
               'IMPRIMIR CONDUCTA
                If ("1.02" <= Codigo4) And (Codigo4 <= "3.99") Then
                   IR = IR + (JR * 3)
                   'Sumatoria = .Fields("PQBim1") + .Fields("PQBim2") / 2
                   PictPrint_Nota_Materia PictLibreta, IR, PosLinea, .Fields("PromPQ")  'Sumatoria
                   IR = IR + (JR * 4)
                   PictPrint_Nota_Materia PictLibreta, IR, PosLinea, .Fields("PromSQ")  'Sumatoria
                   IR = IR + (JR * 4)
                   PictPrint_Nota_Materia PictLibreta, IR, PosLinea, .Fields("PromTQ")  'Sumatoria
                   IR = IR + (JR * 4)
                   If .Fields("PromPQ") > 0 And .Fields("PromSQ") > 0 And .Fields("PromTQ") > 0 Then
                       Diferencia = Redondear((.Fields("PromPQ") + .Fields("PromSQ") + .Fields("PromTQ")) / 3)
                       PictPrint_Nota_Materia PictLibreta, IR, PosLinea, Diferencia     ' .Fields("PromFinal")
                   End If
                   IR = IR + JR
                   'MsgBox .Fields("PromFinal")
                End If
                PictLibreta.FontBold = False
                PosLinea = PosLinea + 0.4
            Else
                Si_No = False
                If .Fields("C") Then Si_No = True
               'MsgBox .Fields("C") & vbCrLf & Si_No
               'IMPRESION DE LAS NOTAS DE LAS MATERIAS
                PictLibreta.FontBold = False
                PictLibreta.FontUnderline = False
                Contador = Contador + 1
                If ("1.02" <= Codigo4) And (Codigo4 <= "3.99") Then
                    If .Fields("SDiv") Then
                        PictLibreta.FontUnderline = True
                        PictLibreta.FontBold = True
                        PictPrint_Texto PictLibreta, 1.2, PosLinea, .Fields("Materia")
                        PictLibreta.FontUnderline = False
                        PictLibreta.FontBold = False
                    ElseIf .Fields("CodMatP") <> Ninguno Then
                        PictPrint_Grafico PictLibreta, RutaSistema & "\ICONOS\vwicn115.ICO", 1.4, PosLinea + 0.1, 0.2, 0.2
                        PictPrint_Texto PictLibreta, 1.8, PosLinea, .Fields("Materia")
                    Else
                        PictPrint_Texto PictLibreta, 1.2, PosLinea, .Fields("Materia")
                    End If
                    IR = PosColumna
                   ' MsgBox .Fields("Materia") & " =  " & .Fields("PQBim1")
                   If Print_Nota(1) Then PictPrint_Nota_Materia PictLibreta, IR, PosLinea, .Fields("PQBim1"), Si_No, Dec_Nota
                   IR = IR + JR
                   If Print_Nota(2) Then PictPrint_Nota_Materia PictLibreta, IR, PosLinea, .Fields("PQBim2"), Si_No, Dec_Nota
                   IR = IR + JR
                   If Print_Nota(3) Then PictPrint_Nota_Materia PictLibreta, IR, PosLinea, .Fields("ExamenPQ"), Si_No, Dec_Nota
                   IR = IR + JR
                   If Print_Nota(4) Then PictPrint_Nota_Materia PictLibreta, IR, PosLinea, .Fields("PromPQ"), Si_No, Dec_Nota
                   IR = IR + JR
                   If Print_Nota(5) Then PictPrint_Nota_Materia PictLibreta, IR, PosLinea, .Fields("SQBim1"), Si_No, Dec_Nota
                   IR = IR + JR
                   If Print_Nota(6) Then PictPrint_Nota_Materia PictLibreta, IR, PosLinea, .Fields("SQBim2"), Si_No, Dec_Nota
                   IR = IR + JR
                   If Print_Nota(7) Then PictPrint_Nota_Materia PictLibreta, IR, PosLinea, .Fields("ExamenSQ"), Si_No, Dec_Nota
                   IR = IR + JR
                   If Print_Nota(8) Then PictPrint_Nota_Materia PictLibreta, IR, PosLinea, .Fields("PromSQ"), Si_No, Dec_Nota
                   IR = IR + JR
                   If Print_Nota(9) Then PictPrint_Nota_Materia PictLibreta, IR, PosLinea, .Fields("TQBim1"), Si_No, Dec_Nota
                   IR = IR + JR
                   If Print_Nota(10) Then PictPrint_Nota_Materia PictLibreta, IR, PosLinea, .Fields("TQBim2"), Si_No, Dec_Nota
                   IR = IR + JR
                   If Print_Nota(11) Then PictPrint_Nota_Materia PictLibreta, IR, PosLinea, .Fields("ExamenTQ"), Si_No, Dec_Nota
                   IR = IR + JR
                   If Print_Nota(12) Then PictPrint_Nota_Materia PictLibreta, IR, PosLinea, .Fields("PromTQ"), Si_No, Dec_Nota
                   IR = IR + JR
                   If ("2.00" <= Codigo4) And (Codigo4 <= "3.99") Then
                      If Print_Nota(13) Then PictPrint_Nota_Materia PictLibreta, IR, PosLinea, .Fields("Supletorio"), Si_No, Dec_Nota
                   End If
                   IR = IR + JR
                   'If ("2.00" <= Codigo4) And (Codigo4 <= "3.99") Then
                      Sumatoria = (.Fields("PromPQ") + .Fields("PromSQ") + .Fields("PromTQ")) / 3
                      If Print_Nota(14) Then PictPrint_Nota_Materia PictLibreta, IR, PosLinea, Sumatoria, Si_No, Dec_Nota
                   'End If
                   IR = IR + JR
                End If
                If .Fields("Orden") <> 9 And .Fields("I") And .Fields("CodMatP") = Ninguno Then
                   'MsgBox .Fields("Orden") & vbCrLf & Si_No
                    If Not Si_No Then
                       VPQBim1 = VPQBim1 + .Fields("PQBim1")
                       VPQBim2 = VPQBim2 + .Fields("PQBim2")
                       
                       VSQBim1 = VSQBim1 + .Fields("SQBim1")
                       VSQBim2 = VSQBim2 + .Fields("SQBim2")
                       
                       VTQBim1 = VTQBim1 + .Fields("TQBim1")
                       VTQBim2 = VTQBim2 + .Fields("TQBim2")
                       
                       VPromPQ = VPromPQ + .Fields("PromPQ")
                       VPromSQ = VPromSQ + .Fields("PromSQ")
                       VPromTQ = VPromTQ + .Fields("PromTQ")
                       
                       VExamenPQ = VExamenPQ + .Fields("ExamenPQ")
                       VExamenSQ = VExamenSQ + .Fields("ExamenSQ")
                       VExamenTQ = VExamenTQ + .Fields("ExamenTQ")
                       
                       VPromFinal = VPromFinal + .Fields("PromFinal")
                       If .Fields("PQBim1") > 0 Then TotalRegs(1) = TotalRegs(1) + 1
                       If .Fields("PQBim2") > 0 Then TotalRegs(2) = TotalRegs(2) + 1
                       If .Fields("ExamenPQ") > 0 Then TotalRegs(3) = TotalRegs(3) + 1
                       If .Fields("PromPQ") > 0 Then TotalRegs(4) = TotalRegs(4) + 1
                       
                       If .Fields("SQBim1") > 0 Then TotalRegs(5) = TotalRegs(5) + 1
                       If .Fields("SQBim2") > 0 Then TotalRegs(6) = TotalRegs(6) + 1
                       If .Fields("ExamenSQ") > 0 Then TotalRegs(7) = TotalRegs(7) + 1
                       If .Fields("PromSQ") > 0 Then TotalRegs(8) = TotalRegs(8) + 1
                       
                       If .Fields("TQBim1") > 0 Then TotalRegs(9) = TotalRegs(9) + 1
                       If .Fields("TQBim2") > 0 Then TotalRegs(10) = TotalRegs(10) + 1
                       If .Fields("ExamenTQ") > 0 Then TotalRegs(11) = TotalRegs(11) + 1
                       If .Fields("PromTQ") > 0 Then TotalRegs(12) = TotalRegs(12) + 1
                       
                       If .Fields("PromFinal") > 0 Then TotalRegs(14) = TotalRegs(14) + 1
                    End If
                End If
                PosLinea = PosLinea + 0.36
            End If
         End If
        .MoveNext
      Loop
     'Faltas justificadas o atraso
      If ("1.02" <= Codigo4) And (Codigo4 <= "3.99") Then
          PictLibreta.FontSize = 9
          PictPrint_Texto PictLibreta, 1.1, PosLinea, "FALTAS JUSTIFICADAS"
          PictPrint_Texto PictLibreta, 5.5, PosLinea, Format(Faltas_Just, "00")
          
          PictPrint_Texto PictLibreta, 8, PosLinea, "FALTAS INJUSTIFICADAS"
          PictPrint_Texto PictLibreta, 12.5, PosLinea, Format(Faltas_Injust, "00")
    
          PictPrint_Texto PictLibreta, 16.5, PosLinea, "ATRASOS"
          PictPrint_Texto PictLibreta, 18.5, PosLinea, Format(Atrasos, "00")
          PosLinea = PosLinea + 0.4
      End If
      If ("1.02" <= Codigo4) And (Codigo4 <= "3.99") Then
        'Lineas de Observacion
         PictLibreta.FontSize = 7
         PictLibreta.FontBold = True
         PictLibreta.FontUnderline = True
         PictPrint_Texto PictLibreta, 1, PosLinea, "NOMENCLATURA"
         PictLibreta.FontUnderline = False
         PictLibreta.FontBold = False
         PictPrint_Texto PictLibreta, 1, PosLinea + 0.4, "S  SOBRESALIENTE"
         PictPrint_Texto PictLibreta, 1, PosLinea + 0.7, "M  MUY BUENA"
         PictPrint_Texto PictLibreta, 1, PosLinea + 1, "R  REGULAR"
         PictPrint_Texto PictLibreta, 1, PosLinea + 1.3, "B  BUENA"
         PictPrint_Texto PictLibreta, 1, PosLinea + 1.6, "I  INSUFICIENTE"
         Cuadricula = False
         PictLibreta.FontBold = True
         PictLibreta.FontUnderline = True
         PictLibreta.FontName = TipoArialNarrow
         PictLibreta.FontSize = 8
         PictPrint_Texto PictLibreta, 3.5, PosLinea, "OBSERVACIONES:"
         PictLibreta.FontUnderline = False
         PictPrint_Texto PictLibreta, 5.8, PosLinea, String(105, "_")
         PosLinea = PosLinea + 1
         PictLibreta.FontBold = True
         PictLibreta.FontUnderline = False
         PictPrint_Texto PictLibreta, 3.5, PosLinea, String(35, "_")
         PictPrint_Texto PictLibreta, 9.2, PosLinea, String(35, "_")
         PictPrint_Texto PictLibreta, 15, PosLinea, String(35, "_")
         PosLinea = PosLinea + 0.35
         Select Case Codigo4
           Case "1.02" To "1.99"
                PictPrint_Texto PictLibreta, 4.5, PosLinea, Director
                PictPrint_Texto PictLibreta, 9.5, PosLinea, "PROFESOR(A)"
           Case "2.00" To "3.99"
                PictPrint_Texto PictLibreta, 4.5, PosLinea, Rector
                PictPrint_Texto PictLibreta, 9.5, PosLinea, ULCase(Trim(LblDirigente.Caption))
         End Select
         PictPrint_Texto PictLibreta, 15.5, PosLinea, "FIRMA DEL REPRESENTANTE"
         PosLinea = PosLinea + 0.35
         Select Case Codigo4
           Case "1.02" To "1.99"
                PictPrint_Texto PictLibreta, 4.5, PosLinea, ULCase(TextoDirector)
           Case "2.00" To "3.99"
                PictPrint_Texto PictLibreta, 5, PosLinea, ULCase(TextoRector)
                PictPrint_Texto PictLibreta, 9.5, PosLinea, "Dirigente del Curso"
         End Select
         PictLibreta.FontBold = False
      End If
      RatonNormal
      Cuadricula = False
      MensajeEncabData = "": SQLMsg1 = "": SQLMsg2 = "": SQLMsg3 = "": SQLMsg4 = ""
   End If
 End With
End Sub

Public Sub Libreta_Del_Alumno_Quimestres(TipoObjeto As Object, AdoLib As Adodc)
Dim AnchoDib As Single
Dim AltoDib As Single
Dim Curso As String
Dim Alumno As String
Dim Paralelo As String
Dim LogoMinisterio As String
Dim Notas() As String
Dim PosXPict As Single
Dim AnchoPict() As CtasAsiento
Dim Y0 As Single
Dim y1 As Single
Dim X0 As Single
Dim x1 As Single
Dim PosLineaX As Single
Dim PosLineaXF As Single
Dim TotalRegs(22) As Integer
Dim ValorRegs(22) As Currency
Dim Promedio As Currency
Dim Promedio_General As Currency
Dim CanPromFinal As Byte
Dim Formato_Nota As String
Dim UnaVezOpta As Boolean
Dim EsPreBa As Boolean
Dim FinDeLibreta As Boolean
Dim Listar_Informe As Boolean
Dim Cualitativa1 As Boolean
Dim Cualitativa2 As Boolean
Dim TroncoAux As Boolean
FinDeLibreta = False
EsPreBa = False
UnaVezOpta = True
Listar_Informe = False
TroncoAux = False

''If TypeOf TipoObjeto Is Printer Then
''   MsgBox ("Printer")
''ElseIf TypeOf TipoObjeto Is PictureBox Then
''   MsgBox ("PictureBox")
''ElseIf TypeOf TipoObjeto Is mjwPDF Then
''   MsgBox ("ObjetoPDF")
''End If
PosLinea = 0
TipoLetra = TipoArial
 With AdoLib.Recordset
  If .RecordCount > 0 Then
      Curso = .Fields("Curso")
      Paralelo = .Fields("Paralelo")
      Alumno = .Fields("Alumno")
      NombreCliente = .Fields("Alumno")
      Do While Not .EOF
         If Len(SQLInforme) > 1 Then
            If Len(.Fields(SQLInforme)) > 1 Then Listar_Informe = True
         End If
         PosLinea = PosLinea + 0.36
        .MoveNext
      Loop
     .MoveFirst
    If OpcionNotas = 4 Then
       ReDim Notas(8) As String
       Notas(0) = " P. I"
       Notas(1) = " P. II"
       Notas(2) = " P. III"
       Notas(3) = " 80%"
       Notas(4) = " Exam"
       Notas(5) = " 20%"
       Notas(6) = " Prom"
       If Asistencias And Codigo4 > "2" Then Notas(7) = " EQUIV." Else Notas(7) = " EQUIVALENCIA"
    Else
       ReDim Notas(7) As String
       Notas(0) = " TAI"
       Notas(1) = " AIC"
       Notas(2) = " AGC"
       Notas(3) = " LOE"
       Notas(4) = " EXA"
       Notas(5) = " PRO"
       If Asistencias And Codigo4 > "2" Then Notas(6) = " CUALIT." Else Notas(6) = " CUALITATIVA"
    End If
    If Mid$(Curso, 1, 4) <= "1.01" Then
       If OpcionNotas = 5 Then
          Notas(0) = " P. I"
          Notas(1) = " P. II"
          Notas(2) = " P. III"
          Notas(3) = " Prom"
          Notas(4) = " S. I"
          Notas(5) = " S. II"
          Notas(6) = " S. III"
          Notas(7) = " Prom"
          Notas(8) = " EQUIVALENCIA"
       Else
          Notas(0) = " P. I"
          Notas(1) = " P. II"
          Notas(2) = " P. III"
          Notas(3) = " Prom"
          Notas(4) = " "
          Notas(5) = " "
          Notas(6) = " EQUIVALENCIA"
          Notas(7) = " "
       End If
       EsPreBa = True
    End If
    CadenaParcial = Visualizar_Notas_Periodo(LstPeriodos)
    AltoDib = PosLinea
    If TypeOf TipoObjeto Is PictureBox Then TipoObjeto.Cls
    'PictPrint_Tipo_Letra TipoArial, PorteLetra  'TipoHelvetica
    'PictPrint_Color_Letra QBColor(Negro)
    PosLinea = 0.2
    PosColumna = 8
    JR = 1
     AnchoDib = 20
     LogoMinisterio = RutaSistema & "\LOGOS\MINISEDU.JPG"
     
     cPrint.printImagen LogoTipo, 1, 0.5, 4.5, 2.25
     cPrint.printImagen LogoMinisterio, 16.5, 0.5, 2.5, 1.7
     'PictPrint_Color_Letra QBColor(Negro)
     'PictPrint_Estilo_Letra FONT_BOLD, True
     'PictPrint_Tipo_Letra TipoArial, 12
     
     PictPrint_Texto 1, PosLinea, Institucion1, , 18, True
     PosLinea = PosLinea + 0.6
     'PictPrint_Porte_Letra 9
     PictPrint_Texto 1, PosLinea, Institucion2, , 18, True
     PosLinea = PosLinea + 0.5
     'PictPrint_Porte_Letra 8
     Cadena = Direccion & " - Teléfono: " & Telefono1
     PictPrint_Texto 1, PosLinea, Cadena, , 18, True
     PosLinea = PosLinea + 0.5
     Cadena = EmailEmpresa
     If Len(Codigo_AMIE) > 1 Then Cadena = Cadena & String(20, " ") & "Codigo AMIE: " & Codigo_AMIE
     PictPrint_Texto 1, PosLinea, Cadena, , 18, True
     PosLinea = PosLinea + 0.5
     'PictPrint_Porte_Letra 9
     PictPrint_Texto 1, PosLinea, UCase(NombreCiudad) & "-" & UCase(NombrePais), , 18, True
     'PictPrint_Porte_Letra 12
     PosLinea = PosLinea + 0.5
     PictPrint_Texto 1, PosLinea, TextoWeb, , 18, True
     'PictPrint_Porte_Letra 9
     PosLinea = PosLinea + 0.8
     PictPrint_Texto 1, PosLinea, "Año Lectivo " & Anio_Lectivo
     PictPrint_Texto 11.2, PosLinea, TextoLeyenda
     'PictPrint_Porte_Letra 10
     PosLinea = PosLinea + 0.5
     'PictPrint_Estilo_Letra FONT_BOLD, True
     PictPrint_Texto 1, PosLinea, "INFORME ACADÉMICO DEL ESTUDIANTE", , 18, True
     PosLinea = PosLinea + 0.4
     If OpcionNotas = 4 Then
        PictPrint_Texto 1, PosLinea + 0.05, UCase(CadenaParcial)
     Else
        Cadena = UCase(SinEspaciosIzqNoBlancos(CadenaParcial, 1) & " " & SinEspaciosIzqNoBlancos(CadenaParcial, 2))
        PictPrint_Texto 1, PosLinea + 0.05, Cadena
        If Mid$(Curso, 1, 4) > "1.01" Then
           Cadena = UCase(SinEspaciosIzqNoBlancos(CadenaParcial, 3) & " " & SinEspaciosIzqNoBlancos(CadenaParcial, 4))
           PictPrint_Texto 14.5, PosLinea + 0.05, Cadena
        End If
     End If
     'PictPrint_Estilo_Letra FONT_BOLD, False
     PosLinea = PosLinea + 0.5
     PictPrint_Texto 1, PosLinea, Alumno
     PosLinea = PosLinea + 0.5
'     PictPrint_Estilo_Letra FONT_BOLD, True
     PictPrint_Texto 1, PosLinea, Curso & " " & Paralelo
     PosLinea = PosLinea + 0.5
     PictPrint_Texto 1, PosLinea, "Docente Tutor:"
     'PictPrint_Estilo_Letra FONT_BOLD, False
     PictPrint_Texto 3.5, PosLinea, ULCase(LblDirigente.Caption)
     PictPrint_Texto 13.1, PosLinea, FechaStrgCiudad(MBFecha)
     PosLinea = PosLinea + 0.5
     PFil = PosLinea
    'Cuadro Externo
     'PictPrint_Porte_Letra 9
     'PictPrint_Cuadro_Linea 1, PosLinea, 19, PosLinea + 0.45, QBColor(Negro), "B"
     'PictPrint_Color_Letra QBColor(Negro)
     PictPrint_Texto 1, PosLinea + 0.05, "ÁMBITOS Y ASIGNATURAS", , PosColumna - 1, True
    'Imprimimos las columnas de las materias
     IR = PosColumna
     If OpcionNotas = 4 Then J = 7 Else J = 6
     For I = 0 To J
         PictPrint_Texto IR + 0.05, PosLinea + 0.05, Notas(I)
         'PictPrint_Cuadro_Linea IR, PosLinea, IR, PosLinea + 0.5, QBColor(Negro)
         IR = IR + JR
     Next I
     If Asistencias And Codigo4 >= "2" Then
        IR = IR + 0.4
        'PictPrint_Cuadro_Linea IR, PosLinea, IR, PosLinea + 0.5, QBColor(Negro)
        IR = IR + 0.1
        PictPrint_Texto IR + 0.05, PosLinea + 0.05, "FJ"
        IR = IR + 0.75
        'PictPrint_Cuadro_Linea IR, PosLinea, IR, PosLinea + 0.5, QBColor(Negro)
        IR = IR + 0.1
        PictPrint_Texto IR + 0.05, PosLinea + 0.05, "FI"
        IR = IR + 0.75
        'PictPrint_Cuadro_Linea IR, PosLinea, IR, PosLinea + 0#, QBColor(Negro)
        IR = IR + 0.1
        PictPrint_Texto IR + 0.05, PosLinea + 0.05, "A"
     End If
     'PictPrint_Cuadro_Linea 19, PosLinea, 19, PosLinea + 0.5, QBColor(Negro), "B"
     PosLinea = PosLinea + 0.5
     'PictPrint_Estilo_Letra FONT_BOLD, False
      If Dec_Nota > 0 Then Formato_Nota = "00." & String(Dec_Nota, "0") Else Formato_Nota = "00"
      Codigo = .Fields("Codigo")
      Cadena1 = .Fields("Curso")
      Cadena2 = .Fields("Alumno")
      Codigo4 = .Fields("CodE")
      For I = 0 To 20
          TotalRegs(I) = 0
          ValorRegs(I) = 0
      Next I
      'TipoObjeto.FillColor = QBColor(Blanco_Brillante)
      'TipoObjeto.ForeColor = QBColor(Negro)
      VPQBim1 = 0: VPQBim2 = 0: VPQBim3 = 0
      VSQBim1 = 0: VSQBim2 = 0: VSQBim3 = 0
      VPromPQ = 0: VPromSQ = 0: VPromTQ = 0
      VPromFinal = 0: VProm = 0: VExamenPQ = 0: VExamenSQ = 0: VExamenTQ = 0
      Promedio_General = 0
      PosLineaX = PosLinea
      Do While Not .EOF
         Cualitativa1 = False
         Cualitativa2 = False
         If .Fields("C") Then Cualitativa1 = True
         If .Fields("C2") Then Cualitativa2 = True
         If EsPreBa Then Cualitativa1 = True
         SinImprimir = .Fields("SinImprimir")
         'PictPrint_Porte_Letra 9
        'MsgBox .Fields("Materia") & vbCrLf & Cualitativa1 & vbCrLf & Cualitativa2
         If .Fields("CodMatP") <> Ninguno Then
             If OpcionNotas = 4 Then
                'PictPrint_Cuadro_Linea 8, PosLinea, 15, PosLinea + 0.45, RGB(230, 246, 246), "BF"
             Else
                'PictPrint_Cuadro_Linea 8, PosLinea, 14, PosLinea + 0.45, RGB(230, 246, 246), "BF"
             End If
             'PictPrint_Porte_Letra 8
         End If
         'PictPrint_Color_Letra QBColor(Negro)
         Opciones = .Fields("Orden")
         JR = 1
         If PosColumna > 0 Then
           'MsgBox PosColumna & vbCrLf & .Fields("Orden")
            If .Fields("Orden") = 9 Then
               'Clubes
                TroncoAux = True
               'Imprimimos lineas al final de las notas para los promedios
                'PictPrint_Cuadro_Linea 1, PosLineaX, 1, PosLineaXF - 0.05, QBColor(Negro)
                IR = PosColumna
                If OpcionNotas = 4 Then I = 8 Else I = 7
                For J = 1 To I
                    'PictPrint_Cuadro_Linea IR, PosLineaX, IR, PosLineaXF - 0.05, QBColor(Negro)
                    IR = IR + JR
                Next J
                If Asistencias And Codigo4 > "2" Then
                   IR = IR + 0.4
                   PictPrint_Cuadro_Linea IR, PosLineaX, IR, PosLineaXF - 0.05, QBColor(Negro)
                   IR = IR + 0.85
                   PictPrint_Cuadro_Linea IR, PosLineaX, IR, PosLineaXF - 0.05, QBColor(Negro)
                   IR = IR + 0.85
                   PictPrint_Cuadro_Linea IR, PosLineaX, IR, PosLineaXF - 0.05, QBColor(Negro)
                End If
                PictPrint_Cuadro_Linea 19, PosLineaX, 19, PosLineaXF, QBColor(Negro)
                PictPrint_Cuadro_Linea 1, PosLineaXF - 0.05, 19, PosLineaXF - 0.05, QBColor(Negro)
                If TotalReg = 0 Then TotalReg = 1
                For I = 0 To 16
                    If TotalRegs(I) = 0 Then TotalRegs(I) = 1
                Next I
                PosLinea = PosLinea - 0.05
                PictPrint_Cuadro_Linea 1, PosLinea, 19, PosLinea, QBColor(Negro)
                PosLinea = PosLinea + 0.05
                PictPrint_Cuadro_Linea 1, PosLinea, 8, PosLinea + 0.5, QBColor(Negro), "B"
                PictPrint_Cuadro_Linea 8, PosLinea, 10, PosLinea + 0.5, QBColor(Negro), "B"
                PictPrint_Cuadro_Linea 10, PosLinea, 19, PosLinea + 0.5, QBColor(Negro), "B"
                PictPrint_Estilo_Letra FONT_BOLD, True
                PictPrint_Texto 1.1, PosLinea + 0.1, "T O T A L"
                PictPrint_Estilo_Letra FONT_BOLD, False
                IR = PosColumna
                If OpcionNotas = 4 Then J = 6 Else J = 5
                For I = 0 To J
                    PictPrint_Cuadro_Linea IR, PosLinea + 0.05, IR, PosLinea + 0.5, QBColor(Negro), "B"
                    If EsPreBa Then
                       PictPrint_Porte_Letra 9
                       If I < 4 Then PictPrint_Nota_Materia IR + 0.1, PosLinea + 0.1, ValorRegs(I) / TotalRegs(I), True, 0, EsPreBa
                    Else
                       PictPrint_Porte_Letra 8
                       PictPrint_Texto IR + 0.1, PosLinea + 0.1, Format(ValorRegs(I), "00.00")
                    End If
                    IR = IR + JR
                Next I
                PictPrint_Cuadro_Linea IR, PosLinea + 0.05, IR, PosLinea + 0.5, QBColor(Negro), "B"
                PosLinea = PosLinea + 0.5
               'Promedio General
                PictPrint_Porte_Letra 9
                Promedio_General = Redondear(ValorRegs(J) / TotalRegs(J), Tot_Dec_Nota)
                PictPrint_Cuadro_Linea 1, PosLinea, 8, PosLinea + 0.5, QBColor(Negro), "B"
                PictPrint_Cuadro_Linea 8, PosLinea, 10, PosLinea + 0.5, QBColor(Negro), "B"
                PictPrint_Cuadro_Linea 10, PosLinea, 19, PosLinea + 0.5, QBColor(Negro), "B"
                PictPrint_Estilo_Letra FONT_BOLD, True
                PictPrint_Texto 1.1, PosLinea + 0.1, "Promedio General"
                PictPrint_Estilo_Letra FONT_BOLD, False
                PictPrint_Estilo_Letra FONT_ITALIC, False
                IR = PosColumna
                If EsPreBa Then
                   PictPrint_Porte_Letra 10
                   PictPrint_Nota_Materia IR + 0.7, PosLinea + 0.05, Promedio_General, True, Dec_Nota, EsPreBa
                Else
                   PictPrint_Porte_Letra 9
                   PictPrint_Texto IR + 0.5, PosLinea + 0.05, Format(Promedio_General, "00." & String(Tot_Dec_Nota, "0"))
                   IR = IR + (JR * 2)
                   PictPrint_Texto IR, PosLinea + 0.05, Equivalencia(Promedio_General), , 2, True
                End If
                PosLinea = PosLinea + 0.5
                PictPrint_Cuadro_Linea 1, PosLinea, 8, PosLinea + 0.5, QBColor(Negro), "B"
                PictPrint_Cuadro_Linea 8, PosLinea, 10, PosLinea + 0.5, QBColor(Negro), "B"
                PictPrint_Cuadro_Linea 10, PosLinea, 19, PosLinea + 0.5, QBColor(Negro), "B"
                PictPrint_Estilo_Letra FONT_BOLD, True
                PictPrint_Texto 1.1, PosLinea + 0.1, "Evaluación del Comportamiento"
                
               'IMPRIMIR CONDUCTA
               'MsgBox Valor
                PictPrint_Estilo_Letra FONT_BOLD, True
                PictPrint_Nota_Materia PosColumna + 0.8, PosLinea + 0.05, Valor, True, Dec_Nota, , True
                PictPrint_Estilo_Letra FONT_BOLD, False
                PosLinea = PosLinea + 0.5
                PictPrint_Cuadro_Linea 1, PosLinea, 19.01, PosLinea, QBColor(Negro)
                PosLinea = PosLinea + 0.05
                PictPrint_Cuadro_Linea 1, PosLinea, 19.01, PosLinea, QBColor(Negro)
                PosLinea = PosLinea + 0.1
                PosLineaX = PosLinea
                FinDeLibreta = True
            Else
               'IMPRESION DE LAS NOTAS DE LAS MATERIAS
                PictPrint_Estilo_Letra FONT_NORMAL, False
                Contador = Contador + 1
                  If .Fields("CodMatP") <> Ninguno Then
                      cPrint.printImagen RutaSistema & "\ICONOS\vwicn115.ICO", 1.4, PosLinea + 0.1, 0.2, 0.2
                      PictPrint_Estilo_Letra FONT_UNDERLINE, True
                      PosLinea = PictPrint_Texto_Multiple(TipoObjeto, 1.8, PosLinea, .Fields("Materia"), 6)
                     'PictPrint_Texto 1.8, PosLinea, .Fields("Materia")
                      PictPrint_Estilo_Letra FONT_UNDERLINE, False
                      PictPrint_Estilo_Letra FONT_ITALIC, True
                      
                      PictPrint_Estilo_Letra FONT_BOLD, True
                  ElseIf .Fields("SDiv") Then
                      cPrint.printImagen RutaSistema & "\ICONOS\Visto.ICO", 1.4, PosLinea + 0.1, 0.2, 0.2
                      PictPrint_Estilo_Letra FONT_UNDERLINE, True
                      PosLinea = PictPrint_Texto_Multiple(TipoObjeto, 1.8, PosLinea, .Fields("Materia"), 6)
                      'PictPrint_Texto 1.8, PosLinea, .Fields("Materia")
                      PictPrint_Estilo_Letra FONT_UNDERLINE, False
                      If UnaVezOpta Then
                         IR = PosColumna
                         If OpcionNotas = 4 Then J = 7 Else J = 6
                         For I = 0 To J
                             PictPrint_Cuadro_Linea IR, PosLinea - 0.5, IR, PosLinea - 0.05, QBColor(Negro), "B"
                             'PictPrint_Texto IR + 0.05, PosLinea - 0.5, Notas(I)
                             IR = IR + JR
                         Next I
                          UnaVezOpta = False
                      End If
                  Else
                      PosLinea = PictPrint_Texto_Multiple(TipoObjeto, 1.2, PosLinea, .Fields("Materia"), 7)
                  End If
                 IR = PosColumna + 0.1
                'Notas Parciales
                 
                 If .Fields("P") Then
                     If EsPreBa Then     ' ES PREBASICA
                        PictPrint_Nota_Materia IR, PosLinea, .Fields(SQLBim1), Cualitativa1, Dec_Nota, EsPreBa, Cualitativa2
                        IR = IR + JR
                        PictPrint_Nota_Materia IR, PosLinea, .Fields(SQLBim2), Cualitativa1, Dec_Nota, EsPreBa, Cualitativa2
                        IR = IR + JR
                        PictPrint_Nota_Materia IR, PosLinea, .Fields(SQLBim3), Cualitativa1, Dec_Nota, EsPreBa, Cualitativa2
                        IR = IR + JR
                        PictPrint_Nota_Materia IR, PosLinea, .Fields(SQLPromQ), Cualitativa1, Dec_Nota, EsPreBa, Cualitativa2
                        IR = IR + JR
                        IR = IR + JR
                        IR = IR + JR
                        'MsgBox .Fields("CodMat") & vbCrLf & .Fields(SQLPromQ)
                        PictPrint_Estilo_Letra FONT_ITALIC, False
                        PictPrint_Nota_Materia IR, PosLinea, .Fields(SQLPromQ), Cualitativa1, Dec_Nota, EsPreBa, Cualitativa2
                     ElseIf OpcionNotas = 4 Then
                        PictPrint_Nota_Materia IR, PosLinea, .Fields(SQLBim1), Cualitativa1, Dec_Nota, EsPreBa, Cualitativa2
                        IR = IR + JR
                        PictPrint_Nota_Materia IR, PosLinea, .Fields(SQLBim2), Cualitativa1, Dec_Nota, EsPreBa, Cualitativa2
                        IR = IR + JR
                        PictPrint_Nota_Materia IR, PosLinea, .Fields(SQLBim3), Cualitativa1, Dec_Nota, EsPreBa, Cualitativa2
                        IR = IR + JR
                        VPromT = Redondear((.Fields(SQLBim1) + .Fields(SQLBim2) + .Fields(SQLBim3)) / 3, 3)
                        PictPrint_Nota_Materia IR, PosLinea, VPromT * 0.8, Cualitativa1, Dec_Nota, EsPreBa, Cualitativa2
                        IR = IR + JR
                        PictPrint_Nota_Materia IR, PosLinea, .Fields(SQLExamen), Cualitativa1, Dec_Nota, EsPreBa, Cualitativa2
                        IR = IR + JR
                        PictPrint_Nota_Materia IR, PosLinea, .Fields(SQLExamen) * 0.2, Cualitativa1, Dec_Nota, EsPreBa, Cualitativa2
                        IR = IR + JR
                        VPromFinal = Redondear((VPromT * 0.8) + (.Fields(SQLExamen) * 0.2), Dec_Nota)
                        PictPrint_Nota_Materia IR, PosLinea, .Fields(SQLPromQ), Cualitativa1, Dec_Nota, EsPreBa, Cualitativa2
                        IR = IR + JR
                        PictPrint_Estilo_Letra FONT_ITALIC, False
                        PictPrint_Texto IR, PosLinea, Equivalencia(VPromFinal), , 1, True
                     Else
                       'MsgBox Cualitativa1 & vbCrLf & Cualitativa2
                        PictPrint_Nota_Materia IR, PosLinea, .Fields(SQLTAI), Cualitativa1, Dec_Nota, EsPreBa, Cualitativa2
                        IR = IR + JR
                        PictPrint_Nota_Materia IR, PosLinea, .Fields(SQLAIC), Cualitativa1, Dec_Nota, EsPreBa, Cualitativa2
                        IR = IR + JR
                        PictPrint_Nota_Materia IR, PosLinea, .Fields(SQLAGC), Cualitativa1, Dec_Nota, EsPreBa, Cualitativa2
                        IR = IR + JR
                        PictPrint_Nota_Materia IR, PosLinea, .Fields(SQLL), Cualitativa1, Dec_Nota, EsPreBa, Cualitativa2
                        IR = IR + JR
                        PictPrint_Nota_Materia IR, PosLinea, .Fields(SQLExaP), Cualitativa1, Dec_Nota, EsPreBa, Cualitativa2
                        IR = IR + JR
                        
                        PictPrint_Nota_Materia IR, PosLinea, .Fields(SQLProm), Cualitativa1, Dec_Nota, EsPreBa, Cualitativa2
                        IR = IR + JR
                        PictPrint_Estilo_Letra FONT_ITALIC, False
                        PictPrint_Texto IR, PosLinea, Equivalencia(.Fields(SQLProm), , , Cualitativa2), , 1, True
                        
                     End If
                 Else
                     If OpcionNotas = 1 Then
                        PictPrint_Nota_Materia IR, PosLinea, .Fields(SQLBim1), Cualitativa1, Dec_Nota, EsPreBa, Cualitativa2
                        IR = IR + (JR * 6) + 0.4
                        PictPrint_Estilo_Letra FONT_ITALIC, False
                        PictPrint_Texto IR, PosLinea, Equivalencia(.Fields(SQLBim1), , , Cualitativa2), , 1, True
                     End If
                     If OpcionNotas = 2 Then
                        PictPrint_Nota_Materia IR, PosLinea, .Fields(SQLBim2), Cualitativa1, Dec_Nota, EsPreBa, Cualitativa2
                        IR = IR + (JR * 6) + 0.4
                        PictPrint_Estilo_Letra FONT_ITALIC, False
                        PictPrint_Texto IR, PosLinea, Equivalencia(.Fields(SQLBim2), , , Cualitativa2), , 1, True
                     End If
                     If OpcionNotas = 3 Then
                        PictPrint_Nota_Materia IR, PosLinea, .Fields(SQLBim3), Cualitativa1, Dec_Nota, EsPreBa, Cualitativa2
                        IR = IR + (JR * 6) + 0.4
                        PictPrint_Estilo_Letra FONT_ITALIC, False
                        PictPrint_Texto IR, PosLinea, Equivalencia(.Fields(SQLBim3), , , Cualitativa2), , 1, True
                     End If
                     If OpcionNotas = 4 Then
                        If TroncoAux Then
                           IR = IR + (JR * 6)
                           PictPrint_Nota_Materia IR, PosLinea, .Fields(SQLExamen), Cualitativa1, Dec_Nota, EsPreBa, Cualitativa2
                           VPromT = .Fields(SQLExamen)
                        Else
                           VPromT = Redondear((.Fields(SQLBim1) + .Fields(SQLBim2) + .Fields(SQLBim3)) / 3, 3)
                           PictPrint_Nota_Materia IR, PosLinea, VPromT, Cualitativa1, Dec_Nota, EsPreBa, Cualitativa2
                           IR = IR + JR
                           PictPrint_Nota_Materia IR, PosLinea, VPromT * 0.8, Cualitativa1, Dec_Nota, EsPreBa, Cualitativa2
                           IR = IR + JR
                           PictPrint_Nota_Materia IR, PosLinea, .Fields(SQLExamen), Cualitativa1, Dec_Nota, EsPreBa, Cualitativa2
                           IR = IR + JR
                           PictPrint_Nota_Materia IR, PosLinea, .Fields(SQLExamen) * 0.2, Cualitativa1, Dec_Nota, EsPreBa, Cualitativa2
                           IR = IR + JR
                           VPromT = Redondear((VPromT * 0.8) + (.Fields(SQLExamen) * 0.2), Dec_Nota)
                           PictPrint_Nota_Materia IR, PosLinea, .Fields(SQLPromQ), Cualitativa1, Dec_Nota, EsPreBa, Cualitativa2
                           'PictPrint_Nota_Materia IR, PosLinea, VPromT, Cualitativa1, Dec_Nota
                        End If
                        IR = IR + JR
                        PictPrint_Estilo_Letra FONT_ITALIC, False
                        PictPrint_Texto IR + 0.3, PosLinea, Equivalencia(CCur(VPromT), , , Cualitativa2), , 1, True
                        IR = PosColumna
                     End If
                End If
                If .Fields("Orden") <> 9 And .Fields("I") And .Fields("CodMatP") = Ninguno Then
                   'MsgBox .Fields("Orden") & vbCrLf & Cualitativa1 & vbCrLf & VPQBim3
                    'If Not Cualitativa1 Then
                       If OpcionNotas = 4 Then
                          If EsPreBa Then
                             ValorRegs(0) = ValorRegs(0) + .Fields(SQLBim1)
                             ValorRegs(1) = ValorRegs(1) + .Fields(SQLBim2)
                             ValorRegs(2) = ValorRegs(2) + .Fields(SQLBim3)
                             VPromT = Redondear((.Fields(SQLBim1) + .Fields(SQLBim2) + .Fields(SQLBim3)) / 3)
                             ValorRegs(3) = ValorRegs(3) + VPromT
                             ValorRegs(4) = ValorRegs(4) + .Fields(SQLExamen)
                             ValorRegs(5) = ValorRegs(5) + .Fields(SQLExamen)
                             VPromFinal = VPromT
                             ValorRegs(6) = ValorRegs(6) + .Fields(SQLPromQ)   'VPromFinal
                             
                             If .Fields(SQLBim1) Then TotalRegs(0) = TotalRegs(0) + 1
                             If .Fields(SQLBim2) Then TotalRegs(1) = TotalRegs(1) + 1
                             If .Fields(SQLBim3) Then TotalRegs(2) = TotalRegs(2) + 1
                             If VPromT Then TotalRegs(3) = TotalRegs(3) + 1
                             If .Fields(SQLExamen) Then TotalRegs(4) = TotalRegs(4) + 1
                             If .Fields(SQLExamen) Then TotalRegs(5) = TotalRegs(5) + 1
                             If VPromFinal Then TotalRegs(6) = TotalRegs(6) + 1
                          
                          Else
                             ValorRegs(0) = ValorRegs(0) + .Fields(SQLBim1)
                             ValorRegs(1) = ValorRegs(1) + .Fields(SQLBim2)
                             ValorRegs(2) = ValorRegs(2) + .Fields(SQLBim3)
                             VPromT = Redondear((.Fields(SQLBim1) + .Fields(SQLBim2) + .Fields(SQLBim3)) / 3, 3)
                             ValorRegs(3) = ValorRegs(3) + (VPromT * 0.8)
                             ValorRegs(4) = ValorRegs(4) + .Fields(SQLExamen)
                             ValorRegs(5) = ValorRegs(5) + (.Fields(SQLExamen) * 0.2)
                             VPromFinal = Redondear((VPromT * 0.8) + (.Fields(SQLExamen) * 0.2))
                             ValorRegs(6) = ValorRegs(6) + .Fields(SQLPromQ)   'VPromFinal
                             
                             If .Fields(SQLBim1) Then TotalRegs(0) = TotalRegs(0) + 1
                             If .Fields(SQLBim2) Then TotalRegs(1) = TotalRegs(1) + 1
                             If .Fields(SQLBim3) Then TotalRegs(2) = TotalRegs(2) + 1
                             If VPromT Then TotalRegs(3) = TotalRegs(3) + 1
                             If .Fields(SQLExamen) Then TotalRegs(4) = TotalRegs(4) + 1
                             If .Fields(SQLExamen) Then TotalRegs(5) = TotalRegs(5) + 1
                             If VPromFinal Then TotalRegs(6) = TotalRegs(6) + 1
                          End If
                       Else
                          ValorRegs(0) = ValorRegs(0) + .Fields(SQLTAI)
                          ValorRegs(1) = ValorRegs(1) + .Fields(SQLAIC)
                          ValorRegs(2) = ValorRegs(2) + .Fields(SQLAGC)
                          ValorRegs(3) = ValorRegs(3) + .Fields(SQLL)
                          ValorRegs(4) = ValorRegs(4) + .Fields(SQLExaP)
                          ValorRegs(5) = ValorRegs(5) + .Fields(SQLProm)
                          
                          If .Fields(SQLTAI) Then TotalRegs(0) = TotalRegs(0) + 1
                          If .Fields(SQLAIC) Then TotalRegs(1) = TotalRegs(1) + 1
                          If .Fields(SQLAGC) Then TotalRegs(2) = TotalRegs(2) + 1
                          If .Fields(SQLL) Then TotalRegs(3) = TotalRegs(3) + 1
                          If .Fields(SQLExaP) Then TotalRegs(4) = TotalRegs(4) + 1
                          If .Fields(SQLProm) Then TotalRegs(5) = TotalRegs(5) + 1
                       End If
                    'End If
                End If
                If Asistencias And Codigo4 >= "2" Then
                   Leer_Asistencia_Materia Codigo4, Codigo, .Fields("CodMat")
                   IR = IR + 1.5
                   If Faltas_Just_PorMat > 0 Then PictPrint_Texto IR + 0.05, PosLinea + 0.05, Format(Faltas_Just_PorMat, "00")
                   IR = IR + 0.85
                   If Faltas_Injust_PorMat > 0 Then PictPrint_Texto IR + 0.05, PosLinea + 0.05, Format(Faltas_Injust_PorMat, "00")
                   IR = IR + 0.85
                   If Atrasos_PorMat > 0 Then PictPrint_Texto IR + 0.05, PosLinea + 0.05, Format(Atrasos_PorMat, "00")
                End If
                PosLinea = PosLinea + 0.45
                PictPrint_Cuadro_Linea PosColumna, PosLinea, 19, PosLinea, QBColor(Negro)
                PosLinea = PosLinea + 0.05
                PosLineaXF = PosLinea
            End If
         End If
        .MoveNext
      Loop
      If (PosLineaXF - PosLineaX) > 0 Then
         PictPrint_Cuadro_Linea 1, PosLineaX - 0.1, 19, PosLineaXF - 0.05, QBColor(Negro), "B"
         PictPrint_Cuadro_Linea 1, PosLineaXF, 19.03, PosLineaXF, QBColor(Negro)
         IR = PosColumna
         If OpcionNotas = 4 Then I = 8 Else I = 7
         For J = 1 To I
             PictPrint_Cuadro_Linea IR, PosLineaX - 0.05, IR, PosLineaXF - 0.05, QBColor(Negro)
             IR = IR + JR
         Next J
         If Asistencias And Codigo4 > "2" Then
            IR = IR + 0.4
            PictPrint_Cuadro_Linea IR, PosLineaX, IR, PosLineaXF - 0.05, QBColor(Negro)
            IR = IR + 0.85
            PictPrint_Cuadro_Linea IR, PosLineaX, IR, PosLineaXF - 0.05, QBColor(Negro)
            IR = IR + 0.85
            PictPrint_Cuadro_Linea IR, PosLineaX, IR, PosLineaXF - 0.05, QBColor(Negro)
         End If
         PosLinea = PosLinea + 0.1
      End If
        PosLineaX = PosLinea
       'RECOMENDACIONES
        PictPrint_Porte_Letra 8
        PosLinea = PosLinea + 0.3
        PictPrint_Estilo_Letra FONT_BOLD, True
        PictPrint_Estilo_Letra FONT_UNDERLINE, True
        PictPrint_Texto 1, PosLinea, "RECOMENDACIONES:"
        PictPrint_Estilo_Letra FONT_UNDERLINE, False
        PictPrint_Estilo_Letra FONT_BOLD, False
        PictPrint_Texto 1, PosLinea, String(60, "_")
        PosLinea = PosLinea + 0.6
        PictPrint_Texto 1, PosLinea, String(60, "_")
        PosLinea = PosLinea + 0.6
        PictPrint_Estilo_Letra FONT_BOLD, True
        PictPrint_Estilo_Letra FONT_UNDERLINE, True
        PictPrint_Texto 1, PosLinea, "PLAN DE MEJORA ACADÉMICA:"
        PictPrint_Estilo_Letra FONT_UNDERLINE, False
        PictPrint_Estilo_Letra FONT_BOLD, False
        PictPrint_Texto 1, PosLinea, String(60, "_")
        PosLinea = PosLinea + 0.6
        PictPrint_Texto 1, PosLinea, String(60, "_")
        PosLinea = PosLineaX + 0.2
        PictPrint_Cuadro_Linea 11.2, PosLinea, 19, PosLinea + 0.5, QBColor(Blanco), "BF"
        PictPrint_Cuadro_Linea 11.2, PosLinea, 16, PosLinea + 0.5, QBColor(Negro), "B"
        PictPrint_Porte_Letra 9
        PictPrint_Estilo_Letra FONT_BOLD, True
        PictPrint_Texto 11.2, PosLinea + 0.05, "Asistencia", , 4, True
        PictPrint_Estilo_Letra FONT_BOLD, False
        PictPrint_Cuadro_Linea 15, PosLinea, 16, PosLinea + 0.5, QBColor(Negro), "B"
        PictPrint_Cuadro_Linea 16, PosLinea, 17, PosLinea + 0.5, QBColor(Negro), "B"
        PictPrint_Cuadro_Linea 17, PosLinea, 18, PosLinea + 0.5, QBColor(Negro), "B"
        PictPrint_Cuadro_Linea 18, PosLinea, 19, PosLinea + 0.5, QBColor(Negro), "B"
        PictPrint_Texto 15, PosLinea + 0.05, "IP", , 1, True
        PictPrint_Texto 16, PosLinea + 0.05, "IIP", , 1, True
        PictPrint_Texto 17, PosLinea + 0.05, "IIIP", , 1, True
        PictPrint_Texto 18, PosLinea + 0.05, "Total", , 1, True
        PosLinea = PosLinea + 0.5
        PictPrint_Porte_Letra 8
        PictPrint_Cuadro_Linea 11.2, PosLinea, 15, PosLinea + 0.5, QBColor(Negro), "B"
        PictPrint_Cuadro_Linea 15, PosLinea, 16, PosLinea + 0.5, QBColor(Negro), "B"
        PictPrint_Cuadro_Linea 16, PosLinea, 17, PosLinea + 0.5, QBColor(Negro), "B"
        PictPrint_Cuadro_Linea 17, PosLinea, 18, PosLinea + 0.5, QBColor(Negro), "B"
        PictPrint_Cuadro_Linea 18, PosLinea, 19, PosLinea + 0.5, QBColor(Negro), "B"
        PictPrint_Texto 11.3, PosLinea + 0.05, "Días laborados del alumno"
        IR = 15.3
        PictPrint_Texto IR, PosLinea + 0.05, Format(Dias_Laborados1, "00")
        IR = IR + 1
        PictPrint_Texto IR, PosLinea + 0.05, Format(Dias_Laborados2, "00")
        IR = IR + 1
        PictPrint_Texto IR, PosLinea + 0.05, Format(Dias_Laborados3, "00")
        IR = IR + 1
        If OpcionNotas = 4 Then PictPrint_Texto 14.3 + OpcionNotas, PosLinea + 0.05, Format(Dias_Laborados, "00")
        PosLinea = PosLinea + 0.5
        PictPrint_Cuadro_Linea 11.2, PosLinea, 15, PosLinea + 0.5, QBColor(Negro), "B"
        PictPrint_Cuadro_Linea 15, PosLinea, 16, PosLinea + 0.5, QBColor(Negro), "B"
        PictPrint_Cuadro_Linea 16, PosLinea, 17, PosLinea + 0.5, QBColor(Negro), "B"
        PictPrint_Cuadro_Linea 17, PosLinea, 18, PosLinea + 0.5, QBColor(Negro), "B"
        PictPrint_Cuadro_Linea 18, PosLinea, 19, PosLinea + 0.5, QBColor(Negro), "B"
        PictPrint_Texto 11.3, PosLinea + 0.05, "Faltas justificadas"
        IR = 15.3
        PictPrint_Texto IR, PosLinea + 0.05, Format(Faltas_Just1, "00")
        IR = IR + 1
        PictPrint_Texto IR, PosLinea + 0.05, Format(Faltas_Just2, "00")
        IR = IR + 1
        PictPrint_Texto IR, PosLinea + 0.05, Format(Faltas_Just3, "00")
        IR = IR + 1
        If OpcionNotas = 4 Then PictPrint_Texto IR, PosLinea + 0.05, Format(Faltas_Just, "00")
        PosLinea = PosLinea + 0.5
        PictPrint_Cuadro_Linea 11.2, PosLinea, 15, PosLinea + 0.5, QBColor(Negro), "B"
        PictPrint_Cuadro_Linea 15, PosLinea, 16, PosLinea + 0.5, QBColor(Negro), "B"
        PictPrint_Cuadro_Linea 16, PosLinea, 17, PosLinea + 0.5, QBColor(Negro), "B"
        PictPrint_Cuadro_Linea 17, PosLinea, 18, PosLinea + 0.5, QBColor(Negro), "B"
        PictPrint_Cuadro_Linea 18, PosLinea, 19, PosLinea + 0.5, QBColor(Negro), "B"
        PictPrint_Texto 11.3, PosLinea + 0.05, "Faltas no justificadas"
        IR = 15.3
        PictPrint_Texto IR, PosLinea + 0.05, Format(Faltas_Injust1, "00")
        IR = IR + 1
        PictPrint_Texto IR, PosLinea + 0.05, Format(Faltas_Injust2, "00")
        IR = IR + 1
        PictPrint_Texto IR, PosLinea + 0.05, Format(Faltas_Injust3, "00")
        IR = IR + 1
        If OpcionNotas = 4 Then PictPrint_Texto IR, PosLinea + 0.05, Format(Faltas_Injust, "00")
        PosLinea = PosLinea + 0.5
        PictPrint_Cuadro_Linea 11.2, PosLinea, 15, PosLinea + 0.5, QBColor(Negro), "B"
        PictPrint_Cuadro_Linea 15, PosLinea, 16, PosLinea + 0.5, QBColor(Negro), "B"
        PictPrint_Cuadro_Linea 16, PosLinea, 17, PosLinea + 0.5, QBColor(Negro), "B"
        PictPrint_Cuadro_Linea 17, PosLinea, 18, PosLinea + 0.5, QBColor(Negro), "B"
        PictPrint_Cuadro_Linea 18, PosLinea, 19, PosLinea + 0.5, QBColor(Negro), "B"
        PictPrint_Texto 11.3, PosLinea + 0.05, "Atrasos"
        IR = 15.3
        PictPrint_Texto IR, PosLinea + 0.05, Format(Atrasos1, "00")
        IR = IR + 1
        PictPrint_Texto IR, PosLinea + 0.05, Format(Atrasos2, "00")
        IR = IR + 1
        PictPrint_Texto IR, PosLinea + 0.05, Format(Atrasos3, "00")
        IR = IR + 1
        If OpcionNotas = 4 Then PictPrint_Texto IR, PosLinea + 0.05, Format(Atrasos, "00")
     
     'Tabla Cualitativa y cuantitativa
      PosLinea = PosLinea + 0.7
      PictPrint_Cuadro_Linea 1, PosLinea, 19, PosLinea + 0.5, QBColor(Blanco), "BF"
      PictPrint_Cuadro_Linea 1, PosLinea, 19, PosLinea + 0.5, QBColor(Negro), "B"
      
      PictPrint_Porte_Letra 9
      PictPrint_Texto 1, PosLinea + 0.05, "EVALUACIÓN DEL COMPORTAMIENTO", , 10, True
      PictPrint_Texto 10, PosLinea + 0.05, "EQUIVALENCIAS", , 8, True
      PosLinea = PosLinea + 0.5
      PosLineaX = PosLinea
      PictPrint_Porte_Letra 7
      CanPromFinal = UBound(Equivalencias)
      Do While CanPromFinal > 0
         CanPromFinal = CanPromFinal - 1
         With Equivalencias(CanPromFinal)
              PictPrint_Cuadro_Linea 1, PosLinea, 10.7, PosLinea, QBColor(Negro)
              PosLinea = PosLinea + 0.04
              If Mid$(Codigo4, 1, 4) <= "1.01" Then
                 PictPrint_Texto 1.1, PosLinea + 0.05, .Cualitativa
              Else
                 PictPrint_Texto 1.1, PosLinea + 0.05, .Letras
              End If
              PictPrint_Texto 1.5, PosLinea + 0.05, "= " & .Significado_Letras
              PosLinea = PictPrint_Texto_Multiple(TipoObjeto, 3.9, PosLinea + 0.05, .Significado_Evaluacion, 6.7)
              PosLinea = PosLinea + 0.4
         End With
      Loop
      PosLinea = PosLineaX
      CanPromFinal = UBound(Equivalencias)
      Do While CanPromFinal > 0
         CanPromFinal = CanPromFinal - 1
         With Equivalencias(CanPromFinal)
              If EsPreBa Then
                 PictPrint_Texto 10.8, PosLinea + 0.05, .Cualitativa
                 PictPrint_Texto 11.5, PosLinea + 0.05, ": " & .Significado_Letras
              Else
                 PictPrint_Texto 10.8, PosLinea + 0.05, .Equivalencia
                 PictPrint_Texto 11.5, PosLinea + 0.05, ": " & .Significado_Equivalencia
                 PictPrint_Texto 17.5, PosLinea + 0.05, "= " & .Rango
              End If
              PosLinea = PosLinea + 0.35
         End With
      Loop
      PosLinea = PosLinea + 0.1
      PictPrint_Cuadro_Linea 10.7, PosLinea, 19, PosLinea, QBColor(Negro)
      PosLinea = PosLinea + 0.1
      PictPrint_Texto 10.8, PosLinea, "TAI"
      PictPrint_Texto 11.5, PosLinea, ": Trabajos Académicos Independientes (Tareas)"
      PosLinea = PosLinea + 0.35
      PictPrint_Texto 10.8, PosLinea, "AIC"
      PictPrint_Texto 11.5, PosLinea, ": Actividades Individuales en Clase"
      PosLinea = PosLinea + 0.35
      PictPrint_Texto 10.8, PosLinea, "AGC"
      PictPrint_Texto 11.5, PosLinea, ": Actividades Grupales en Clase"
      PosLinea = PosLinea + 0.35
      PictPrint_Texto 10.8, PosLinea, "LOE"
      PictPrint_Texto 11.5, PosLinea, ": Lecciones Orales o Escritas"
      PosLinea = PosLinea + 0.35
      PictPrint_Texto 10.8, PosLinea, "EXA"
      PictPrint_Texto 11.5, PosLinea, ": Evaluación Sumativa"
      PosLinea = PosLinea + 0.35
      PictPrint_Texto 10.8, PosLinea, "PRO"
      PictPrint_Texto 11.5, PosLinea, ": Promedio de Parciales"
      PosLinea = PosLinea + 0.35
      
     'Fin del Cuadro de Evaluacion
      PictPrint_Cuadro_Linea 1, PosLineaX, 19, PosLinea, QBColor(Negro), "B"
      PictPrint_Cuadro_Linea 10.7, PosLineaX - 0.5, 10.7, PosLinea, QBColor(Negro)
      If Listar_Informe Then
         PosLinea = PosLinea + 0.1
         PictPrint_Porte_Letra 10
         PictPrint_Texto 1, PosLinea, "Revise el informe académico del Alumno."
      End If
      PosLinea = 25
      cPrint.printImagen RutaSistema & "\FORMATOS\FIRMALIBRETA.jpg", 1.5, PosLinea + 1.3, 5, 1.5
      PosLinea = PosLinea + 2.5
      PictPrint_Porte_Letra 9
      PictPrint_Cuadro_Linea 1.5, PosLinea, 6.5, PosLinea, QBColor(Negro)
      PictPrint_Cuadro_Linea 13.5, PosLinea, 18.5, PosLinea, QBColor(Negro)
      PosLinea = PosLinea + 0.1
      If ("1.00" <= Codigo4) And (Codigo4 <= "3.99") Then
         Select Case Codigo4
           Case "1.00" To "1.99"
                PictPrint_Texto 1.5, PosLinea, Director, , 5, True
                PictPrint_Texto 13.5, PosLinea, ULCase(LblDirigente.Caption), , 5, True
           Case "2.00" To "3.99"
                PictPrint_Texto 1.5, PosLinea, ULCase(Rector), , 5, True
                PictPrint_Texto 13.5, PosLinea, ULCase(LblDirigente.Caption), , 5, True
         End Select
         PosLinea = PosLinea + 0.4
         Select Case Codigo4
           Case "1.00" To "1.99"
                PictPrint_Texto 1.5, PosLinea, TextoDirector, , 5, True
                PictPrint_Texto 13.5, PosLinea, "PROFESOR(A)", , 5, True
           Case "2.00" To "3.99"
                PictPrint_Texto 1.5, PosLinea, TextoRector, , 5, True
                PictPrint_Texto 13.5, PosLinea, "Docente Tutor", , 5, True
         End Select
         PictPrint_Estilo_Letra FONT_BOLD, False
      End If
      RatonNormal
      Cuadricula = False
      MensajeEncabData = "": SQLMsg1 = "": SQLMsg2 = "": SQLMsg3 = "": SQLMsg4 = ""
   End If
 End With
End Sub

Public Sub Libreta_Del_Alumno_Trimestre2(AdoLib As Adodc)
Dim AnchoDib As Single
Dim AltoDib As Single
Dim Curso As String
Dim Alumno As String
Dim Paralelo As String
Dim PosXPict As Single
Dim AnchoPict() As CtasAsiento
Dim Y0 As Single
Dim y1 As Single
Dim X0 As Single
Dim x1 As Single
Dim PosLineaX As Single
Dim TotalRegs(22) As Integer
Dim CanPromFinal As Byte
Dim Formato_Nota As String

  PosLinea = 5.5
  With AdoLib.Recordset
   If .RecordCount > 0 Then
       Curso = .Fields("Curso")
       Paralelo = .Fields("Paralelo")
       Alumno = .Fields("Alumno")
       NombreCliente = .Fields("Alumno")
       Do While Not .EOF
          PosLinea = PosLinea + 0.36
         .MoveNext
       Loop
      .MoveFirst
   End If
  End With
  AltoDib = PosLinea
  PictLibreta.FontName = TipoTimes
  PictLibreta.ForeColor = QBColor(Negro)
  PosColumna = 8.2
  JR = 0.85
  Select Case Mid$(Curso, 1, 4)
    Case "0.00" To "1.01" '
         AnchoDib = 20
         'AltoDib = 13
         Grafico_Kinder AltoDib
         PictLibreta.PaintPicture LoadPicture(LogoTipo), 0.1, 0.1, 2, 1
         PictLibreta.FontSize = 15
         PictPrint_Texto PictLibreta, 4, 0.1, Empresa
         PictLibreta.FontSize = 10
         PictPrint_Texto PictLibreta, 4, 0.7, "FICHA DE DESARROLLO DE DESTREZAS Y HABILIDADES"
         
         PictPrint_Texto PictLibreta, 14, 0.2, "NIVEL:"
         PictPrint_Texto PictLibreta, 0.5, 1.8, "ALUMNO(A):"
         
         PictLibreta.FontBold = False
         PictLibreta.FontSize = 9
         PictLibreta.FontName = TipoArialNarrow
         PictPrint_Texto PictLibreta, 15.3, 0.2, Curso
         PictPrint_Texto PictLibreta, 14, 0.7, Paralelo
         PictPrint_Texto PictLibreta, 2.7, 1.8, Alumno
         PosColumna = 0
         PosLinea = 0
    Case "1.02" To "1.99"
         AnchoDib = 20.3
         'AltoDib = 11.3
         PictLibreta.Cls
         If Encabezado_Prim Then
            If LogoTipo <> "" Then PictLibreta.PaintPicture LoadPicture(LogoTipo), 17 / 2, 0.2, 3.6, 1.8
            PictLibreta.FontSize = 16
            PictPrint_Texto PictLibreta, 1.2, 2.2, UCase(Institucion1), , 19, True
            PictPrint_Texto PictLibreta, 1.2, 2.9, UCase(Institucion2), , 19, True
         End If
         PictLibreta.FontSize = 8
         PFil = 7
        'Cuadro Externo
         PictLibreta.Line (1, PFil + 0.5)-(19.9, PFil + AltoDib - 2), QBColor(Negro), B
         PictLibreta.Line (1, PFil + 1.4)-(19.9, PFil + 3.6), QBColor(Negro), B
         PCol = PosColumna - 0.35
         PictLibreta.Line (1, PFil + 2.2)-(PCol, PFil + 3), QBColor(Negro), B
         PictLibreta.Line (PCol, PFil + 1.4)-(PCol, PFil + AltoDib - 2), QBColor(Negro)
         PictLibreta.Line (PCol, PFil + 2.2)-(18.05, PFil + 2.2), QBColor(Negro)
         PictLibreta.FontSize = 8
         PictPrint_Texto PictLibreta, PCol + 0.5, PFil + 1.45, "PRIMER"
         PictPrint_Texto PictLibreta, PCol + 0.3, PFil + 1.75, "TRIMESTRE"
         PictPrint_Texto PictLibreta, PCol + 3.55, PFil + 1.45, "SEGUNDO"
         PictPrint_Texto PictLibreta, PCol + 3.5, PFil + 1.75, "TRIMESTRE"
         PictPrint_Texto PictLibreta, PCol + 7.3, PFil + 1.45, "TERCER"
         PictPrint_Texto PictLibreta, PCol + 7.1, PFil + 1.75, "TRIMESTRE"
         PictLibreta.FontSize = 11
         PictPrint_Texto PictLibreta, 3, PFil + 3.1, "M A T E R I A S"
         PictLibreta.FontSize = 10
         For I = 1 To 14
            'Lineas Verticales de la Libreta
             If I = 5 Or I = 9 Or I = 13 Or I = 14 Then
                PictLibreta.Line (PCol, PFil + 1.4)-(PCol, PFil + AltoDib - 2), QBColor(Negro)
             Else
                PictLibreta.Line (PCol, PFil + 2.2)-(PCol, PFil + AltoDib - 2), QBColor(Negro)
             End If
            'Encabezados de la libreta
             Select Case I
               Case 1, 5, 9: cPrint.printTextoAngulo PictLibreta, PCol + 0.05, PFil + 3.4, 90, 4, 10, "Primer"
                             cPrint.printTextoAngulo PictLibreta, PCol + 0.4, PFil + 3.4, 90, 4, 10, "Parcial"
               Case 2, 6, 10: cPrint.printTextoAngulo PictLibreta, PCol + 0.05, PFil + 3.4, 90, 4, 10, "Segundo"
                              cPrint.printTextoAngulo PictLibreta, PCol + 0.4, PFil + 3.4, 90, 4, 10, "Parcial"
               Case 3, 7, 11
                              If FormatoLibreta = "TRIMESTRE2" Then
                                 cPrint.printTextoAngulo PictLibreta, PCol + 0.05, PFil + 3.4, 90, 4, 10, "Prueba"
                                 cPrint.printTextoAngulo PictLibreta, PCol + 0.4, PFil + 3.4, 90, 4, 10, "Trimestral"
                              End If
               Case 4: cPrint.printTextoAngulo PictLibreta, PCol + 0.05, PFil + 3.4, 90, 4, 10, "Prom."
                       cPrint.printTextoAngulo PictLibreta, PCol + 0.4, PFil + 3.4, 90, 4, 10, "1er. T."
               Case 8: cPrint.printTextoAngulo PictLibreta, PCol + 0.05, PFil + 3.4, 90, 4, 10, "Prom."
                       cPrint.printTextoAngulo PictLibreta, PCol + 0.4, PFil + 3.4, 90, 4, 10, "2do. T."
               Case 12: cPrint.printTextoAngulo PictLibreta, PCol + 0.05, PFil + 3.4, 90, 4, 10, "Prom."
                        cPrint.printTextoAngulo PictLibreta, PCol + 0.4, PFil + 3.4, 90, 4, 10, "3er. T."
               Case 14: cPrint.printTextoAngulo PictLibreta, PCol + 0.05, PFil + 3.4, 90, 5, 10, "Promedio"
                        cPrint.printTextoAngulo PictLibreta, PCol + 0.4, PFil + 3.4, 90, 5, 10, "Final"
             End Select
             If ("2.00" <= Codigo4) And (Codigo4 <= "3.99") Then
               Select Case I
                 Case 13: cPrint.printTextoAngulo PictLibreta, PCol + 0.25, PFil + 3.4, 90, 4, 9, "Supletorio"
                 Case 14: cPrint.printTextoAngulo PictLibreta, PCol + 0.05, PFil + 3.4, 90, 5, 10, "Promedio"
                          cPrint.printTextoAngulo PictLibreta, PCol + 0.4, PFil + 3.4, 90, 5, 10, "Final"
               End Select
             End If
             PCol = PCol + JR
         Next I
         PictLibreta.FontSize = 14
         PictPrint_Texto PictLibreta, 8.2, 3.7, "SECCIÓN PRIMARIA"
         PictPrint_Texto PictLibreta, 6.9, 4.5, "BOLETIN DE CALIFICACIONES"
         PictPrint_Texto PictLibreta, 7.6, 5.2, "AÑO LECTIVO " & Anio_Lectivo
         If OpcPeriodo("PQBim1", LstPeriodos) Then Cadena = "1er. Parcial"
         If OpcPeriodo("PQ", LstPeriodos) Then Cadena = "2do. Parcial"
         If OpcPeriodo("SQBim1", LstPeriodos) Then Cadena = "3er. Parcial"
         If OpcPeriodo("SQ", LstPeriodos) Then Cadena = "4to. Parcial"
         If OpcPeriodo("PF", LstPeriodos) Then Cadena = "Periodo Final"
         PictLibreta.FontSize = 9
         PictPrint_Texto PictLibreta, 9.8, 6, Cadena
         PictPrint_Texto PictLibreta, 14, PFil, FechaStrgCiudad(MBFecha)
         PosLinea = 1.4
         PictLibreta.FontSize = 8
         PictPrint_Texto PictLibreta, 1.2, PFil + 0.61, "Estudiante:"
         PictPrint_Texto PictLibreta, 1.2, PFil + 1.45, "Paralelo:"
         PictPrint_Texto PictLibreta, 1.2, PFil + 2.25, "Curso:"
         PictLibreta.FontBold = False
         PictLibreta.FontName = TipoArialNarrow
         PictPrint_Texto PictLibreta, 1.2, PFil + 1.8, Curso
         PictPrint_Texto PictLibreta, 1.2, PFil + 0.95, Alumno
         PictPrint_Texto PictLibreta, 1.2, PFil + 2.6, Paralelo
         PosLinea = 10.6
    Case "2.00" To "3.99"
         AnchoDib = 20.3
         PictLibreta.Cls
         If LogoTipo <> "" Then PictLibreta.PaintPicture LoadPicture(LogoTipo), 1, 0.2, 2, 1
         PictLibreta.FontSize = 8
         PFil = 0.1
        'Cuadro Externo
         PictLibreta.Line (1, 0.1)-(19.9, AltoDib - 2), QBColor(Negro), B
         PictLibreta.Line (1, 1.4)-(19.9, 3.6), QBColor(Negro), B
         PCol = PosColumna - 0.35
         PictLibreta.Line (1, 2.2)-(PCol, 3), QBColor(Negro), B
         PictLibreta.Line (PCol, 1.4)-(PCol, AltoDib - 2), QBColor(Negro)
         PictLibreta.Line (PCol, 2.2)-(18.05, 2.2), QBColor(Negro)
         PictLibreta.FontSize = 8
         PictPrint_Texto PictLibreta, PCol + 0.5, PFil + 1.4, "PRIMER"
         PictPrint_Texto PictLibreta, PCol + 0.3, PFil + 1.7, "TRIMESTRE"
         PictPrint_Texto PictLibreta, PCol + 3.55, PFil + 1.4, "SEGUNDO"
         PictPrint_Texto PictLibreta, PCol + 3.5, PFil + 1.7, "TRIMESTRE"
         PictPrint_Texto PictLibreta, PCol + 7.3, PFil + 1.4, "TERCER"
         PictPrint_Texto PictLibreta, PCol + 7.1, PFil + 1.7, "TRIMESTRE"

         PictLibreta.FontSize = 11
         PictPrint_Texto PictLibreta, 3, 3.1, "M A T E R I A S"
         PictLibreta.FontSize = 10
         For I = 1 To 14
            'Lineas Verticales de la Libreta
             If I = 5 Or I = 9 Or I = 13 Or I = 14 Then
                PictLibreta.Line (PCol, 1.4)-(PCol, AltoDib - 2), QBColor(Negro)
             Else
                PictLibreta.Line (PCol, 2.2)-(PCol, AltoDib - 2), QBColor(Negro)
             End If
            'Encabezados de la libreta
             Select Case I
               Case 1, 5, 9: cPrint.printTextoAngulo PictLibreta, PCol + 0.05, PFil + 3.4, 90, 4, 10, "Primer"
                             cPrint.printTextoAngulo PictLibreta, PCol + 0.4, PFil + 3.4, 90, 4, 10, "Parcial"
               Case 2, 6, 10: cPrint.printTextoAngulo PictLibreta, PCol + 0.05, PFil + 3.4, 90, 4, 10, "Segundo"
                              cPrint.printTextoAngulo PictLibreta, PCol + 0.4, PFil + 3.4, 90, 4, 10, "Parcial"
               Case 3, 7, 11
                              If FormatoLibreta = "TRIMESTRE2" Then
                                 cPrint.printTextoAngulo PictLibreta, PCol + 0.35, PFil + 3.3, 90, 4, 10, "Examen"
                              End If
               Case 4: cPrint.printTextoAngulo PictLibreta, PCol + 0.05, PFil + 3.4, 90, 4, 10, "Prom."
                       cPrint.printTextoAngulo PictLibreta, PCol + 0.4, PFil + 3.4, 90, 4, 10, "1er. T."
               Case 8: cPrint.printTextoAngulo PictLibreta, PCol + 0.05, PFil + 3.4, 90, 4, 10, "Prom."
                       cPrint.printTextoAngulo PictLibreta, PCol + 0.4, PFil + 3.4, 90, 4, 10, "2do. T."
               Case 12: cPrint.printTextoAngulo PictLibreta, PCol + 0.05, PFil + 3.4, 90, 4, 10, "Prom."
                        cPrint.printTextoAngulo PictLibreta, PCol + 0.4, PFil + 3.4, 90, 4, 10, "3er. T."
               Case 14: cPrint.printTextoAngulo PictLibreta, PCol + 0.1, PFil + 3.4, 90, 5, 10, "Promedio"
                        cPrint.printTextoAngulo PictLibreta, PCol + 0.45, PFil + 3.4, 90, 5, 10, "Final"
             End Select
             If ("2.00" <= Codigo4) And (Codigo4 <= "3.99") Then
               Select Case I
                 Case 13: cPrint.printTextoAngulo PictLibreta, PCol + 0.25, PFil + 3.4, 90, 4, 9, "SUPLETORIO"
                 Case 14: cPrint.printTextoAngulo PictLibreta, PCol + 0.1, PFil + 3.4, 90, 5, 10, "Promedio"
                          cPrint.printTextoAngulo PictLibreta, PCol + 0.45, PFil + 3.4, 90, 5, 10, "Final"
               End Select
             End If
             PCol = PCol + JR
         Next I
         'PictLibreta.Line (PFil, 2.3)-(PFil, 2.8), QBColor(Negro)
         PictLibreta.FontSize = 16
         PictPrint_Texto PictLibreta, 3, 0.1, Institucion1
         PictPrint_Texto PictLibreta, 3, 0.7, UCase(Institucion2)
         PictLibreta.FontSize = 11
         PictPrint_Texto PictLibreta, 13.5, 0.1, "BOLETÍN DE EVALUACIÓN"
         PictLibreta.FontSize = 11
         PictPrint_Texto PictLibreta, 13.5, 0.5, "AÑO LECTIVO " & Anio_Lectivo
         If OpcPeriodo("PQBim1", LstPeriodos) Then Cadena = "1er. Parcial"
         If OpcPeriodo("PQ", LstPeriodos) Then Cadena = "2do. Parcial"
         If OpcPeriodo("SQBim1", LstPeriodos) Then Cadena = "3er. Parcial"
         If OpcPeriodo("SQ", LstPeriodos) Then Cadena = "4to. Parcial"
         If OpcPeriodo("PF", LstPeriodos) Then Cadena = "Periodo Final"
         PictLibreta.FontSize = 9
         'PictPrint_Texto PictLibreta, 13.5, 0.9, Cadena
         PictPrint_Texto PictLibreta, 13.5, 1, FechaStrgCiudad(MBFecha)
         PosLinea = 1.4
         PictLibreta.FontSize = 8
         PictPrint_Texto PictLibreta, 1.2, 1.45, "Estudiante:"
         PictPrint_Texto PictLibreta, 1.2, 2.25, "Curso:"
         PictLibreta.FontBold = False
         PictLibreta.FontName = TipoArialNarrow
         PictPrint_Texto PictLibreta, 6.5, 2.25, Curso
         PictPrint_Texto PictLibreta, 1.2, 1.8, Alumno
         PictPrint_Texto PictLibreta, 1.2, 2.6, Paralelo
         PosLinea = 3.6
  End Select
  If Dec_Nota > 0 Then PosColumna = 8 Else PosColumna = 8.2
 With AdoLib.Recordset
  If .RecordCount > 0 Then
      If Dec_Nota > 0 Then
         PictLibreta.FontSize = 9
         Formato_Nota = "00." & String(Dec_Nota, "0")
         PictLibreta.FontName = TipoArialNarrow
      Else
         PictLibreta.FontSize = 10
         Formato_Nota = "00"
         PictLibreta.FontName = TipoCourierNew
      End If
      Codigo = .Fields("Codigo")
      Cadena1 = .Fields("Curso")
      Cadena2 = .Fields("Alumno")
      Codigo4 = .Fields("CodE")
      Select Case Mid$(Codigo4, 1, 4)
        Case "0.00" To "1.01": y1 = 0  '9.35
        Case "1.02" To "1.99": y1 = 16.4
        Case "2.00" To "3.99": y1 = 9.6
      End Select
      For I = 0 To 21
          TotalRegs(I) = 0
      Next I
      VPQBim1 = 0: VPQBim2 = 0: VSQBim1 = 0: VSQBim2 = 0: VTQBim1 = 0: VTQBim2 = 0
      VPromPQ = 0: VPromSQ = 0: VPromTQ = 0: VPromFinal = 0
      VExamenPQ = 0: VExamenSQ = 0: VExamenTQ = 0
      Do While Not .EOF
         PictLibreta.FontSize = 8
         Opciones = .Fields("Orden")
         JR = 0.85

         If PosColumna > 0 Then
           'MsgBox PosColumna & vbCrLf & .Fields("Orden")
            If .Fields("Orden") = 9 Then
               'PosLinea = y1
                PosLinea = PosLinea + 0.36
                If TotalReg = 0 Then TotalReg = 1
                For I = 0 To 21
                    If TotalRegs(I) = 0 Then TotalRegs(I) = 1
                Next I
                PictLibreta.FontSize = 9
                PictLibreta.FontBold = True
                PictPrint_Texto PictLibreta, 1.1, PosLinea, "T O T A L"
                PictLibreta.FontBold = False
                If Dec_Nota = 0 Then
                   PictLibreta.FontSize = 8
                   IR = PosColumna - 0.2
                Else
                   PictLibreta.FontSize = 7
                   IR = PosColumna - 0.3
                End If
               'MsgBox VPQBim1
               'IMPRIME SUMATORIA TOTAL DE NOTAS
                If ("1.02" <= Codigo4) And (Codigo4 <= "3.99") Then
                   If VPQBim1 > 0 And Print_Nota(1) Then PictPrint_Texto PictLibreta, IR, PosLinea, Format(VPQBim1, Formato_Nota)
                   IR = IR + JR
                   If VPQBim2 > 0 And Print_Nota(2) Then PictPrint_Texto PictLibreta, IR, PosLinea, Format(VPQBim2, Formato_Nota)
                   IR = IR + JR
                   If VExamenPQ > 0 And Print_Nota(3) Then PictPrint_Texto PictLibreta, IR, PosLinea, Format(VExamenPQ, Formato_Nota)
                   IR = IR + JR
                   If VPromPQ > 0 And Print_Nota(4) Then PictPrint_Texto PictLibreta, IR, PosLinea, Format(VPromPQ, Formato_Nota)
                   IR = IR + JR
                   
                   If VSQBim1 > 0 And Print_Nota(5) Then PictPrint_Texto PictLibreta, IR, PosLinea, Format(VSQBim1, Formato_Nota)
                   IR = IR + JR
                   If VSQBim2 > 0 And Print_Nota(6) Then PictPrint_Texto PictLibreta, IR, PosLinea, Format(VSQBim2, Formato_Nota)
                   IR = IR + JR
                   If VExamenSQ > 0 And Print_Nota(7) Then PictPrint_Texto PictLibreta, IR, PosLinea, Format(VExamenSQ, Formato_Nota)
                   IR = IR + JR
                   If VPromSQ > 0 And Print_Nota(8) Then PictPrint_Texto PictLibreta, IR, PosLinea, Format(VPromSQ, Formato_Nota)
                   IR = IR + JR
                   
                   If VTQBim1 > 0 And Print_Nota(9) Then PictPrint_Texto PictLibreta, IR, PosLinea, Format(VTQBim1, Formato_Nota)
                   IR = IR + JR
                   If VTQBim2 > 0 And Print_Nota(10) Then PictPrint_Texto PictLibreta, IR, PosLinea, Format(VTQBim2, Formato_Nota)
                   IR = IR + JR
                   If VExamenTQ > 0 And Print_Nota(11) Then PictPrint_Texto PictLibreta, IR, PosLinea, Format(VExamenTQ, Formato_Nota)
                   IR = IR + JR
                   If VPromTQ > 0 And Print_Nota(12) Then PictPrint_Texto PictLibreta, IR, PosLinea, Format(VPromTQ, Formato_Nota)
                   IR = IR + JR
                   IR = IR + JR
                   If VPromFinal > 0 And Print_Nota(14) Then PictPrint_Texto PictLibreta, IR, PosLinea, Format(VPromFinal, Formato_Nota)
                End If
                PosLinea = PosLinea + 0.35
                PictLibreta.FontBold = True
                PictLibreta.FontSize = 9
                PictPrint_Texto PictLibreta, 1.1, PosLinea, "DESEMPEÑO ACADÉMICO:"
'''                Select Case Codigo4
'''                  Case "1.02" To "1.99": PictPrint_Texto PictLibreta, 1.1, PosLinea, "PROMEDIO DE RENDIMIENTO:"
'''                  Case "2.00" To "3.99": PictPrint_Texto PictLibreta, 1.1, PosLinea, "DESEMPEÑO ACADÉMICO:"
'''                End Select
                PosLinea = PosLinea + 0.05
                PictLibreta.FontBold = False
                PictLibreta.FontSize = 7
                IR = PosColumna - 0.3
               'IMPRIME TOTAL PROMEDIOS
                If ("1.02" <= Codigo4) And (Codigo4 <= "3.99") Then
                   If VPQBim1 > 0 And Print_Nota(1) Then PictPrint_Texto PictLibreta, IR, PosLinea, Format(VPQBim1 / TotalRegs(1), "00.00")
                   IR = IR + JR
                   If VPQBim2 > 0 And Print_Nota(2) Then PictPrint_Texto PictLibreta, IR, PosLinea, Format(VPQBim2 / TotalRegs(2), "00.00")
                   IR = IR + JR
                   If VExamenPQ > 0 And Print_Nota(3) Then PictPrint_Texto PictLibreta, IR, PosLinea, Format(VExamenPQ / TotalRegs(3), "00.00")
                   IR = IR + JR
                   If VPromPQ > 0 And Print_Nota(4) Then PictPrint_Texto PictLibreta, IR, PosLinea, Format(VPromPQ / TotalRegs(4), "00.00")
                   IR = IR + JR
                   
                   If VSQBim1 > 0 And Print_Nota(5) Then PictPrint_Texto PictLibreta, IR, PosLinea, Format(VSQBim1 / TotalRegs(5), "00.00")
                   IR = IR + JR
                   If VSQBim2 > 0 And Print_Nota(6) Then PictPrint_Texto PictLibreta, IR, PosLinea, Format(VSQBim2 / TotalRegs(6), "00.00")
                   IR = IR + JR
                   If VExamenSQ > 0 And Print_Nota(7) Then PictPrint_Texto PictLibreta, IR, PosLinea, Format(VExamenSQ / TotalRegs(7), "00.00")
                   IR = IR + JR
                   If VPromSQ > 0 And Print_Nota(8) Then PictPrint_Texto PictLibreta, IR, PosLinea, Format(VPromSQ / TotalRegs(8), "00.00")
                   IR = IR + JR
                   
                   If VTQBim1 > 0 And Print_Nota(9) Then PictPrint_Texto PictLibreta, IR, PosLinea, Format(VTQBim1 / TotalRegs(9), "00.00")
                   IR = IR + JR
                   If VTQBim2 > 0 And Print_Nota(10) Then PictPrint_Texto PictLibreta, IR, PosLinea, Format(VTQBim2 / TotalRegs(10), "00.00")
                   IR = IR + JR
                   If VExamenTQ > 0 And Print_Nota(11) Then PictPrint_Texto PictLibreta, IR, PosLinea, Format(VExamenTQ / TotalRegs(11), "00.00")
                   IR = IR + JR
                   If VPromTQ > 0 And Print_Nota(12) Then PictPrint_Texto PictLibreta, IR, PosLinea, Format(VPromTQ / TotalRegs(12), "00.00")
                   IR = IR + JR
                   IR = IR + JR
                   If VPromFinal > 0 And Print_Nota(14) Then
                      If FormatoLibreta = "PERIODO" Then
                         PictPrint_Texto PictLibreta, IR, PosLinea, Format(VPromFinal / TotalRegs(19), "00.00")
                      Else
                         PictPrint_Texto PictLibreta, IR, PosLinea, Format(VPromFinal / TotalRegs(14), "00.000")
                      End If
                   End If
                End If
                PosLinea = PosLinea + 0.35
                PictLibreta.FontBold = True
                PictLibreta.FontSize = 9
                PictPrint_Texto PictLibreta, 1.1, PosLinea, "DISCIPLINA:"
                PictLibreta.FontBold = False
                IR = PosColumna
               'IMPRIMIR CONDUCTA
                If ("1.02" <= Codigo4) And (Codigo4 <= "3.99") Then
                   IR = IR + (JR * 3)
                   If .Fields("PromPQ") > 0 And Print_Nota(4) Then PictPrint_Nota_Materia PictLibreta, IR, PosLinea, .Fields("PromPQ")  'Sumatoria
                   IR = IR + (JR * 4)
                   If .Fields("PromSQ") > 0 And Print_Nota(8) Then PictPrint_Nota_Materia PictLibreta, IR, PosLinea, .Fields("PromSQ")  'Sumatoria
                   IR = IR + (JR * 4)
                   If .Fields("PromTQ") > 0 And Print_Nota(12) Then PictPrint_Nota_Materia PictLibreta, IR, PosLinea, .Fields("PromTQ")  'Sumatoria
                   IR = IR + (JR * 2)
                   If .Fields("PromPQ") > 0 And .Fields("PromSQ") > 0 And .Fields("PromTQ") > 0 And Print_Nota(14) Then
                       Diferencia = Redondear((.Fields("PromPQ") + .Fields("PromSQ") + .Fields("PromTQ")) / 3)
                       PictPrint_Nota_Materia PictLibreta, IR, PosLinea, Diferencia     ' .Fields("PromFinal")
                   End If
                   IR = IR + JR
                   'MsgBox .Fields("PromFinal")
                End If
                PictLibreta.FontBold = False
                PosLinea = PosLinea + 0.4
            Else
                Si_No = False
                If .Fields("C") Then Si_No = True
               'MsgBox .Fields("C") & vbCrLf & Si_No
               'IMPRESION DE LAS NOTAS DE LAS MATERIAS
                PictLibreta.FontBold = False
                PictLibreta.FontUnderline = False
                Contador = Contador + 1
                
                If ("1.02" <= Codigo4) And (Codigo4 <= "3.99") Then
                    If .Fields("CodMatP") <> Ninguno Then
                        PictPrint_Texto PictLibreta, 1.8, PosLinea, .Fields("Materia")
                    Else
                        PictPrint_Texto PictLibreta, 1.2, PosLinea, .Fields("Materia")
                    End If
                   IR = PosColumna
                   If Print_Nota(1) Then PictPrint_Nota_Materia PictLibreta, IR, PosLinea, .Fields("PQBim1"), Si_No, Dec_Nota
                   IR = IR + JR
                   If Print_Nota(2) Then PictPrint_Nota_Materia PictLibreta, IR, PosLinea, .Fields("PQBim2"), Si_No, Dec_Nota
                   IR = IR + JR
                   If Print_Nota(3) Then PictPrint_Nota_Materia PictLibreta, IR, PosLinea, .Fields("ExamenPQ"), Si_No, Dec_Nota
                   IR = IR + JR
                   If Print_Nota(4) Then PictPrint_Nota_Materia PictLibreta, IR, PosLinea, .Fields("PromPQ"), Si_No, Dec_Nota
                   IR = IR + JR
                   If Print_Nota(5) Then PictPrint_Nota_Materia PictLibreta, IR, PosLinea, .Fields("SQBim1"), Si_No, Dec_Nota
                   IR = IR + JR
                   If Print_Nota(6) Then PictPrint_Nota_Materia PictLibreta, IR, PosLinea, .Fields("SQBim2"), Si_No, Dec_Nota
                   IR = IR + JR
                   If Print_Nota(7) Then PictPrint_Nota_Materia PictLibreta, IR, PosLinea, .Fields("ExamenSQ"), Si_No, Dec_Nota
                   IR = IR + JR
                   If Print_Nota(8) Then PictPrint_Nota_Materia PictLibreta, IR, PosLinea, .Fields("PromSQ"), Si_No, Dec_Nota
                   IR = IR + JR
                   If Print_Nota(9) Then PictPrint_Nota_Materia PictLibreta, IR, PosLinea, .Fields("TQBim1"), Si_No, Dec_Nota
                   IR = IR + JR
                   If Print_Nota(10) Then PictPrint_Nota_Materia PictLibreta, IR, PosLinea, .Fields("TQBim2"), Si_No, Dec_Nota
                   IR = IR + JR
                   If Print_Nota(11) Then PictPrint_Nota_Materia PictLibreta, IR, PosLinea, .Fields("ExamenTQ"), Si_No, Dec_Nota
                   IR = IR + JR
                   If Print_Nota(12) Then PictPrint_Nota_Materia PictLibreta, IR, PosLinea, .Fields("PromTQ"), Si_No, Dec_Nota
                   IR = IR + JR
                   If ("2.00" <= Codigo4) And (Codigo4 <= "3.99") Then
                      If Print_Nota(13) Then PictPrint_Nota_Materia PictLibreta, IR, PosLinea, .Fields("Supletorio"), Si_No, Dec_Nota
                   End If
                   IR = IR + JR
                   If Print_Nota(14) Then PictPrint_Nota_Materia PictLibreta, IR, PosLinea, .Fields("PromFinal"), Si_No, Dec_Nota
                End If
                If .Fields("Orden") <> 9 And .Fields("I") <> False And .Fields("CodMatP") = Ninguno Then
                   'MsgBox .Fields("Orden") & vbCrLf & Si_No
                    If Not Si_No Then
                       VPQBim1 = VPQBim1 + Redondear(.Fields("PQBim1"), Dec_Nota)
                       VPQBim2 = VPQBim2 + Redondear(.Fields("PQBim2"), Dec_Nota)
                       VSQBim1 = VSQBim1 + Redondear(.Fields("SQBim1"), Dec_Nota)
                       VSQBim2 = VSQBim2 + Redondear(.Fields("SQBim2"), Dec_Nota)
                       VTQBim1 = VTQBim1 + Redondear(.Fields("TQBim1"), Dec_Nota)
                       VTQBim2 = VTQBim2 + Redondear(.Fields("TQBim2"), Dec_Nota)
                       VPromPQ = VPromPQ + Redondear(.Fields("PromPQ"), Dec_Nota)
                       VPromSQ = VPromSQ + Redondear(.Fields("PromSQ"), Dec_Nota)
                       VPromTQ = VPromTQ + Redondear(.Fields("PromTQ"), Dec_Nota)
                       VExamenPQ = VExamenPQ + Redondear(.Fields("ExamenPQ"), Dec_Nota)
                       VExamenSQ = VExamenSQ + Redondear(.Fields("ExamenSQ"), Dec_Nota)
                       VExamenTQ = VExamenTQ + Redondear(.Fields("ExamenTQ"), Dec_Nota)
                       VPromFinal = VPromFinal + Redondear(.Fields("PromFinal"), Dec_Nota)
                       If .Fields("PQBim1") > 0 Then TotalRegs(1) = TotalRegs(1) + 1
                       If .Fields("PQBim2") > 0 Then TotalRegs(2) = TotalRegs(2) + 1
                       If .Fields("ExamenPQ") > 0 Then TotalRegs(3) = TotalRegs(3) + 1
                       If .Fields("PromPQ") > 0 Then TotalRegs(4) = TotalRegs(4) + 1
                       
                       If .Fields("SQBim1") > 0 Then TotalRegs(5) = TotalRegs(5) + 1
                       If .Fields("SQBim2") > 0 Then TotalRegs(6) = TotalRegs(6) + 1
                       If .Fields("ExamenSQ") > 0 Then TotalRegs(7) = TotalRegs(7) + 1
                       If .Fields("PromSQ") > 0 Then TotalRegs(8) = TotalRegs(8) + 1
                       
                       If .Fields("TQBim1") > 0 Then TotalRegs(9) = TotalRegs(9) + 1
                       If .Fields("TQBim2") > 0 Then TotalRegs(10) = TotalRegs(10) + 1
                       If .Fields("ExamenTQ") > 0 Then TotalRegs(11) = TotalRegs(11) + 1
                       If .Fields("PromTQ") > 0 Then TotalRegs(12) = TotalRegs(12) + 1
                       
                       If .Fields("PromFinal") > 0 Then TotalRegs(14) = TotalRegs(14) + 1
                    End If
                End If
                PosLinea = PosLinea + 0.36
            End If
         End If
        .MoveNext
      Loop
     'Faltas justificadas o atraso
      If ("1.02" <= Codigo4) And (Codigo4 <= "3.99") Then
          PictLibreta.FontSize = 9
          PictPrint_Texto PictLibreta, 1.1, PosLinea, "FALTAS JUSTIFICADAS"
          PictPrint_Texto PictLibreta, 5.5, PosLinea, Format(Faltas_Just, "00")
          
          PictPrint_Texto PictLibreta, 8, PosLinea, "FALTAS INJUSTIFICADAS"
          PictPrint_Texto PictLibreta, 12.5, PosLinea, Format(Faltas_Injust, "00")
    
          PictPrint_Texto PictLibreta, 16.5, PosLinea, "ATRASOS"
          PictPrint_Texto PictLibreta, 18.5, PosLinea, Format(Atrasos, "00")
          PosLinea = PosLinea + 0.4
      End If
      If ("1.02" <= Codigo4) And (Codigo4 <= "1.99") Then
         PFil = PosLinea
         PFil = PFil + 0.9
         PictLibreta.Line (0.3, PFil)-(19.9, PFil), QBColor(Negro)
         PFil = PFil + 0.05
         PictLibreta.FontBold = True
         PictLibreta.FontUnderline = True
         PictPrint_Texto PictLibreta, 0.6, PFil, "NOMENCLATURA"
         PictPrint_Texto PictLibreta, 5.5, PFil, "FELICITACIONES POR:"
         PictPrint_Texto PictLibreta, 13, PFil, "PUEDE MEJORAR:"
         PictLibreta.FontUnderline = False
         PictLibreta.FontBold = False
         PFil = PFil + 0.05
         PictPrint_Texto PictLibreta, 1.3, PFil + 0.4, "S  SOBRESALIENTE"
         PictPrint_Texto PictLibreta, 1.3, PFil + 0.75, "M  MUY BUENA"
         PictPrint_Texto PictLibreta, 1.3, PFil + 1.1, "R  REGULAR"
         PictPrint_Texto PictLibreta, 1.3, PFil + 1.45, "B  BUENA"
         PictPrint_Texto PictLibreta, 1.3, PFil + 1.8, "I  INSUFICIENTE"
         PictPrint_Texto PictLibreta, 1.3, PFil + 2.15, "IN INCLUSION"
         PFil = PFil + 0.6
         For I = 1 To 4
             PictPrint_Texto PictLibreta, 5.5, PFil, String(36, "_")
             PictPrint_Texto PictLibreta, 13, PFil, String(36, "_")
             PFil = PFil + 0.6
         Next I
         PFil = PFil + 2
         PictLibreta.FontBold = True
         PictLibreta.FontUnderline = False
         PictPrint_Texto PictLibreta, 1.5, PFil, String(35, "_")
         PictPrint_Texto PictLibreta, 8, PFil, String(35, "_")
         PictPrint_Texto PictLibreta, 15, PFil, String(35, "_")
         PFil = PFil + 0.4
         PictPrint_Texto PictLibreta, 2, PFil, Director
         PictPrint_Texto PictLibreta, 8.2, PFil, "PROFESOR(A)"
         PictPrint_Texto PictLibreta, 15.5, PFil, "FIRMA DEL REPRESENTANTE"
         PictLibreta.FontBold = False
      End If
      If ("2.00" <= Codigo4) And (Codigo4 <= "3.99") Then
        'Lineas de Observacion
         Cuadricula = False
         PictLibreta.FontBold = True
         PictLibreta.FontUnderline = True
         PictLibreta.FontName = TipoArialNarrow
         PictPrint_Texto PictLibreta, 1.1, PosLinea, "OBSERVACIONES:"
         PictLibreta.FontUnderline = False
         PictPrint_Texto PictLibreta, 3.6, PosLinea, String(121, "_")
         PosLinea = PosLinea + 1
         PictLibreta.FontBold = True
         PictLibreta.FontUnderline = False
         PictPrint_Texto PictLibreta, 1.5, PosLinea, String(35, "_")
         PictPrint_Texto PictLibreta, 8, PosLinea, String(35, "_")
         PictPrint_Texto PictLibreta, 15, PosLinea, String(35, "_")
         PosLinea = PosLinea + 0.35
         PictPrint_Texto PictLibreta, 2, PosLinea, Rector
         PictPrint_Texto PictLibreta, 8.2, PosLinea, ULCase(LblDirigente.Caption)
         PictPrint_Texto PictLibreta, 15.5, PosLinea, "FIRMA DEL REPRESENTANTE"
         PosLinea = PosLinea + 0.35
         PictPrint_Texto PictLibreta, 3.5, PosLinea, ULCase(TextoRector)
         PictPrint_Texto PictLibreta, 9.2, PosLinea, "Dirigente del Curso"
         PictLibreta.FontBold = False
      End If
      RatonNormal
      Cuadricula = False
      MensajeEncabData = "": SQLMsg1 = "": SQLMsg2 = "": SQLMsg3 = "": SQLMsg4 = ""
   End If
 End With
End Sub

Public Sub Libreta_Del_Alumno(TipoObjeto As Object, Curso As String, CodigoAlumno As String)
  VScroll1.Visible = True
  HScroll1.Visible = True
  ImpCeros = False
  TipoLetra = TipoArial   'Narrow
  Picture1.ScaleMode = vbCentimeters
  PictLibreta.Cls
  PictLibreta.width = 21
  PictLibreta.Height = 29.7
 'Listamos las notas del Alumnno
  Notas_Del_Alumno Curso, CodigoAlumno
  With AdoLibreta.Recordset
   If .RecordCount > 0 Then
       RatonReloj
       I = 0
       Contador = 0
       PorteLetra = 10
      'Geneeramos el documento
       Set cPrint = New cImpresion
       cPrint.dtipoImpresion = cPrint.cPDF
       cPrint.dNombreArchivo = NombreCliente
       cPrint.dTituloArchivo = "Libreta de " & NombreCliente
       cPrint.dTipoLetra = TipoArial
       cPrint.dOrientacionPagina = 1
       cPrint.dPaginaA4 = True
       cPrint.dEsCampoCorto = False
       cPrint.dVerDocumento = True
       cPrint.iniciaImpresion cPrint
      
       cPrint.printImagen LogoTipo, 1, 1, 5, 2
      'printImagen PathCodigoBarra, 1.5, 1.5, 18, 20
       PosLinea = 3.5
       cPrint.colorDeLetra = Negro
       cPrint.tipoNegrilla = True
       cPrint.letraTipo TipoArial, 8
    
       VScroll1.Max = PictLibreta.ScaleHeight
       HScroll1.Max = PictLibreta.ScaleWidth
       If LstVAlumnos.ListItems.Count <= 0 Then PictLibreta.Cls
       Picture1.Refresh
       Select Case FormatoLibreta
         Case "TRIMESTRE1": Libreta_Del_Alumno_Trimestre1 AdoLibreta
         Case "TRIMESTRE2": Libreta_Del_Alumno_Trimestre2 AdoLibreta
         Case "BIMESTRES":  Libreta_Del_Alumno_Bimestres AdoLibreta
         Case "QUIMESTRE":  Libreta_Del_Alumno_Quimestres AdoLibreta
         Case Else:         Libreta_Del_Alumno_Periodos AdoLibreta
       End Select
   Else
       MsgBox "Este Alumno(a) no tiene datos para procesar"
   End If
  End With
  cPrint.finPagina
  cPrint.finalizaImpresion
  RatonNormal
End Sub

Public Sub Imprimir_Carnet()
 GenerarArchivoPlano FLibretas, AdoAlumnos, "CARNET.TXT", True
 MsgBox "SE GENERO EL SIGUIENTE ARCHIVO:" & vbCrLf _
        & Left(RutaSysBases, 2) & "\SYSBASES\TEXTOS\CARNET.TXT"
End Sub

Public Sub Aptitud_Promocion(TipoObjeto As Object, EsPromocion As Boolean, Curso As String, CodigoAlumno As String)
  TipoObjeto.width = AnchoMaximo
  TipoObjeto.Height = AltoMaximo
 'Listamos las notas del Alumnno
  Notas_Del_Alumno Curso, CodigoAlumno
' MsgBox FormatoLibreta
  Select Case FormatoLibreta
    Case "TRIMESTRE1": Aptitud_Promocion_Trimestre1 EsPromocion, Curso, CodigoAlumno
    Case "TRIMESTRE2": Aptitud_Promocion_Periodos EsPromocion, Curso, CodigoAlumno
    Case "BIMESTRES":  Aptitud_Promocion_Bimestres EsPromocion, Curso, CodigoAlumno
    Case "QUIMESTRE":  Aptitud_Promocion_Quimestre EsPromocion, Curso, CodigoAlumno
    Case Else:         Aptitud_Promocion_Periodos EsPromocion, Curso, CodigoAlumno
  End Select
End Sub

Public Sub Aptitud_Promocion_Bimestres(EsPromocion As Boolean, Curso As String, CodigoAlumno As String)
Dim LineaIni As Single
Dim PuntajeTotal As Single
RatonReloj
PuntajeTotal = 0
InicioX = 0.5: InicioY = 0.1
'Pagina = 1
'Iniciamos la impresion
PictLibreta.Cls
PictLibreta.FontName = TipoTimes
PictLibreta.FontBold = False

With AdoLibreta.Recordset
 If .RecordCount > 0 Then
     CodigoCliente = .Fields("Codigo")
     
     'MsgBox CodigoL
     If Mid$(CodigoL, 1, 4) < "1.02" Then
        SQLMsg1 = "AÑO LECTIVO: " & Anio_Lectivo
        SQLMsg2 = "I N F O R M E    F I N A L"
        SQLMsg3 = ""
        PictPrint_Grafico PictLibreta, RutaSistema & "\FORMATOS\INFORME.GIF", 0.5, 0.1, 19, 13
        'PictEncabezado PictLibreta, 1, 0.5
        PictLibreta.FontBold = True
        PictLibreta.FontSize = 14
        PictPrint_Texto PictLibreta, 7.1, 4, .Fields("Alumno")
     Else
     SQLMsg1 = "AÑO LECTIVO: " & Anio_Lectivo
     SQLMsg2 = ""
     SQLMsg3 = ""
     'PictEncabezado PictLibreta, 1, 0.5
     PictLibreta.FontSize = 14
     Cadena = "CODIGO AMIE " & Codigo_AMIE
     PictPrint_Texto PictLibreta, 1, PosLinea, Cadena, , 19, True
     PosLinea = PosLinea + 0.6
     PictLibreta.FontSize = 12
     Cadena = "MANTA - MANABI - ECUADOR"
     PictPrint_Texto PictLibreta, 1, PosLinea, Cadena, , 19, True
     PosLinea = PosLinea + 0.6
     Cadena = "Correo : " & Mail_Colegio & ", Teléfono: " & Telefono1
     PictPrint_Texto PictLibreta, 1, PosLinea, Cadena, , 19, True
     If EsPromocion Then SQLMsg2 = "CERTIFICADO DE PROMOCION" Else SQLMsg2 = "CERTIFICADO DE APTITUD"
     PictLibreta.FontBold = True
     PosLinea = PosLinea + 0.8
     PictLibreta.FontSize = 18
     PictPrint_Texto PictLibreta, 1, PosLinea, SQLMsg2, , 19, True
     PosLinea = PosLinea + 0.7
     PictLibreta.FontUnderline = False
     PictLibreta.FontBold = False
     PosLinea = PosLinea + 0.5
     PictLibreta.FontSize = 12
     If EsPromocion Then
        Cadena = "La secretaría de La " & UCase(Empresa) & " certifica que "
        If .Fields("Sexo") = "M" Then
            Cadena = Cadena & "el señor " & .Fields("Alumno") & ", alumno del "
        Else
            Cadena = Cadena & "la señorita " & .Fields("Alumno") & ", alumna del "
        End If
        
        Select Case Mid$(Curso, 1, 4)
          Case "1.00" To "1.99"
               Cadena = Cadena & Dato_Curso.Descripcion & " "
          Case "2.00" To "2.99"
               Cadena = Cadena & Dato_Curso.Bachiller & " "
          Case "3.00" To "3.99"
               Cadena = Cadena & Dato_Curso.Descripcion & " "
        End Select
        Cadena = Cadena _
               & "durante el año lectivo " & Anio_Lectivo & " ha obtenido las calificaciones que a continuación se transcriben. "
        Select Case Mid$(Curso, 1, 4)
          Case "1.00" To "1.99"
               Cadena = Cadena & "Las mismas que le acreditan para ser promovido al curso inmediato superior."
          Case "2.00" To "2.99"
               Cadena = Cadena & "Las mismas que le acreditan para ser promovido al curso inmediato superior."
        End Select
        PosLinea = PictPrint_Texto_Justifica(PictLibreta, 2, 18.5, PosLinea, Cadena)
        PosLinea = PosLinea + 0.7
        PictLibreta.FontSize = 11
        PictLibreta.FontBold = True
        PictPrint_Texto PictLibreta, 2, PosLinea, "ASIGNATURAS"
        PictLibreta.FontSize = 8
        LineaIni = PosLinea
        PictPrint_Texto PictLibreta, 8.6, PosLinea, "I"
        PictPrint_Texto PictLibreta, 10, PosLinea, "II"
        PictPrint_Texto PictLibreta, 11, PosLinea, "PUNTAJE"
        PictPrint_Texto PictLibreta, 12.6, PosLinea, "PROME_"
        PictPrint_Texto PictLibreta, 14.1, PosLinea, "SUPLE_"
        PictPrint_Texto PictLibreta, 15.5, PosLinea, "PROMED."
        PictPrint_Texto PictLibreta, 17, PosLinea, "OBSERVACIONES"
        PosLinea = PosLinea + 0.35
        PictPrint_Texto PictLibreta, 8.2, PosLinea, "QUIM"
        PictPrint_Texto PictLibreta, 9.7, PosLinea, "QUIM"
        PictPrint_Texto PictLibreta, 11.1, PosLinea, "TOTAL"
        PictPrint_Texto PictLibreta, 12.6, PosLinea, "DIO"
        PictPrint_Texto PictLibreta, 14.1, PosLinea, "TORIO"
        PictPrint_Texto PictLibreta, 15.5, PosLinea, "FINAL"
     Else
        PictPrint_Texto PictLibreta, 2, PosLinea, "CURSO: " & .Fields("Paralelo")   '.Fields("Curso")
        PosLinea = PosLinea + 1
        If .Fields("Sexo") = "M" Then
            PictPrint_Texto PictLibreta, 2, PosLinea, "NOMBRE DEL ALUMNO: " & .Fields("Alumno")
        Else
            PictPrint_Texto PictLibreta, 2, PosLinea, "NOMBRE DE LA ALUMNA: " & .Fields("Alumno")
        End If
        PosLinea = PosLinea + 0.5
        PictLibreta.FontSize = 12
        PictLibreta.FontBold = True
        PictPrint_Texto PictLibreta, 2, PosLinea, "ASIGNATURAS"
        PictLibreta.FontSize = 8
        PictPrint_Texto PictLibreta, 8, PosLinea, "I"
        PictPrint_Texto PictLibreta, 9.4, PosLinea, "II"
        PictPrint_Texto PictLibreta, 10.4, PosLinea, "PUNTAJE"
        PictPrint_Texto PictLibreta, 12.2, PosLinea, "PROMEDIO"
        PictPrint_Texto PictLibreta, 14, PosLinea, "SUPLE_"
        PictPrint_Texto PictLibreta, 16.5, PosLinea, "PROMEDIO"
        PosLinea = PosLinea + 0.35
        PictPrint_Texto PictLibreta, 7.7, PosLinea, "QUIM"
        PictPrint_Texto PictLibreta, 9.1, PosLinea, "QUIM"
        PictPrint_Texto PictLibreta, 10.4, PosLinea, "TOTAL"
        PictPrint_Texto PictLibreta, 12.2, PosLinea, "QUIMEST."
        PictPrint_Texto PictLibreta, 14, PosLinea, "TORIO"
        PictPrint_Texto PictLibreta, 16.5, PosLinea, "FINAL"
     End If
     PictLibreta.FontBold = False
     PosLinea = PosLinea + 0.5
     PictLibreta.FontSize = 11
     PictLibreta.Line (1.9, PosLinea)-(19.7, PosLinea), QBColor(0)
     PosLinea = PosLinea + 0.05
     Valor_Prom = 0: Saldo = 0: Contador = 0: Abono = 0
     PictLibreta.FontUnderline = False
     PictLibreta.FontBold = False
     Contador = 0
     Do While Not .EOF
        'MsgBox .Fields("Orden") & " - " & .Fields("Materia") & " - " & .Fields("C") & " - " & .Fields("P") & " - " & .Fields("I")
        Contador = Contador + 1
        If .Fields("CodMatP") = Ninguno Then
        If .Fields("I") Then
           'Sumatoria
            Select Case Mid$(Curso, 1, 4)
              Case "2.00" To "3.99" '
                   'If .Fields("P") And .Fields("Orden") <> 9 Then Saldo = Saldo + Valor_Prom         'Saldo = Saldo + Redondear(.Fields("PromFinal"))
                   If .Fields("PromPQ") > 0 And .Fields("PromSQ") <= 0 Then
                       Abono_ME = .Fields("PromPQ")
                   ElseIf .Fields("PromPQ") <= 0 And .Fields("PromSQ") > 0 Then
                       Abono_ME = .Fields("PromSQ")
                   Else
                       Abono_ME = (.Fields("PromPQ") + .Fields("PromSQ")) / 2
                   End If
                   Abono_ME = Redondear(Abono_ME, 3)
                   Total = .Fields("PromPQ") + .Fields("PromSQ")
              Case Else
                   Abono_ME = (.Fields("PQBim1") + .Fields("PQBim2") + .Fields("SQBim1") + .Fields("SQBim2")) / 4
                   Total = .Fields("PQBim1") + .Fields("PQBim2") + .Fields("SQBim1") + .Fields("SQBim2")
                   'If .Fields("P") And .Fields("Orden") <> 9 Then Saldo = Saldo + Valor_Prom
            End Select
           'Disciplina
            
            If .Fields("Orden") = 9 Then Abono = Redondear(.Fields("PromFinal"), 3)
            If .Fields("Orden") <> 9 Then
              'If .Fields("CodMat") <> "019" And .Fields("CodMat") <> "024" Then
              PictPrint_Texto PictLibreta, 2, PosLinea, UCase(.Fields("Materia"))
              PictLibreta.Line (8, PosLinea)-(19, PosLinea + 0.4), Blanco_Brillante, BF
              
              If EsPromocion Then
                 Valor_Prom = Abono_ME
                 If .Fields("PromPQ") > 0 Then PictPrint_Texto PictLibreta, 8.2, PosLinea, Format(.Fields("PromPQ"), "##0.000")
                 If .Fields("PromSQ") > 0 Then PictPrint_Texto PictLibreta, 9.7, PosLinea, Format(.Fields("PromSQ"), "##0.000")
                 If Total > 0 Then PictPrint_Texto PictLibreta, 11.2, PosLinea, Format(Total, "##0.000")
                 If Abono_ME > 0 Then PictPrint_Texto PictLibreta, 12.7, PosLinea, Format(Abono_ME, "##0.000")
                 If .Fields("Supletorio") > 0 Then
                     PictPrint_Texto PictLibreta, 14.2, PosLinea, Format(.Fields("Supletorio"), "##0.000")
                     Valor_Prom = Redondear((Valor_Prom + .Fields("Supletorio")) / 2, 3)
                 End If
                 'MsgBox Total & vbCrLf & Abono_ME & vbCrLf & Redondear(Abono_ME) & vbCrLf & .Fields("PromFinal")
                 'PictPrint_Texto PictLibreta, 14, PosLinea, Format(Abono_ME, "##.00")
                 If Valor_Prom > 0 Then PictPrint_Texto PictLibreta, 15.6, PosLinea, Format(Valor_Prom, "##0.000")
                 If Valor_Prom >= 12 Then
                    PictPrint_Texto PictLibreta, 17, PosLinea, "APROBADO"
                 Else
                    PictPrint_Texto PictLibreta, 17, PosLinea, "REPROBADO"
                 End If
                 
                 
'''                 If .Fields("PromFinal") > 0 Then PictPrint_Texto PictLibreta, 15.6, PosLinea, Format(.Fields("PromFinal"), "##0.000")
'''                 If Redondear(.Fields("PromFinal")) >= 14 Then
'''                    PictPrint_Texto PictLibreta, 17, PosLinea, "APROBADO"
'''                 Else
'''                    PictPrint_Texto PictLibreta, 17, PosLinea, "REPROBADO"
'''                 End If
                 If .Fields("P") And .Fields("Orden") <> 9 Then Saldo = Saldo + Valor_Prom
              Else
                 PictPrint_Texto PictLibreta, 8, PosLinea, Format(.Fields("PromPQ"), "##")
                 PictPrint_Texto PictLibreta, 9.5, PosLinea, Format(.Fields("PromSQ"), "##")
                 PictPrint_Texto PictLibreta, 11, PosLinea, Format(Total, "##")
                 PictPrint_Texto PictLibreta, 12.5, PosLinea, Format(Abono_ME, "##")
                 PictPrint_Texto PictLibreta, 14, PosLinea, Format(.Fields("Supletorio"), "##")
                 PictPrint_Texto PictLibreta, 15.5, PosLinea, Format(.Fields("PromFinal"), "##")
              End If
              PosLinea = PosLinea + 0.45
              PictLibreta.Line (1.9, PosLinea)-(19.7, PosLinea), QBColor(0)
              PosLinea = PosLinea + 0.05
            End If
        End If
        End If
       .MoveNext
     Loop
     PictLibreta.Line (1.9, LineaIni)-(19.7, LineaIni), QBColor(0)
     PictLibreta.Line (1.9, LineaIni)-(1.9, PosLinea), QBColor(0)
     PictLibreta.Line (8, LineaIni)-(8, PosLinea), QBColor(0)
     PictLibreta.Line (9.4, LineaIni)-(9.4, PosLinea), QBColor(0)
     PictLibreta.Line (10.9, LineaIni)-(10.9, PosLinea), QBColor(0)
     PictLibreta.Line (12.4, LineaIni)-(12.4, PosLinea), QBColor(0)
     PictLibreta.Line (13.9, LineaIni)-(13.9, PosLinea), QBColor(0)
     PictLibreta.Line (15.4, LineaIni)-(15.4, PosLinea), QBColor(0)
     PictLibreta.Line (16.9, LineaIni)-(16.9, PosLinea), QBColor(0)
     PictLibreta.Line (19.7, LineaIni)-(19.7, PosLinea), QBColor(0)
     PictLibreta.FontBold = True
     PosLinea = PosLinea + 0.4
     'Contador = Contador - 1
     If EsPromocion Then
        PictLibreta.FontSize = 12
        PictPrint_Texto PictLibreta, 2, PosLinea, "PUNTAJE TOTAL:   " & Format(Saldo, "##0.000")
        PosLinea = PosLinea + 0.8
        PictPrint_Texto PictLibreta, 2, PosLinea, "APROVECHAMIENTO:   " & Format(Saldo / Contador, "##0.000")
        DirCliente = "."
        Select Case Redondear(Saldo / Contador)
          Case 0 To 11:  DirCliente = "INSUFICIENTE"
          Case 12 To 13: DirCliente = "REGULAR"
          Case 14 To 15: DirCliente = "BUENA"
          Case 16 To 18: DirCliente = "MUY BUENA"
          Case 19 To 20: DirCliente = "SOBRESALIENTE"
        End Select
        PictPrint_Texto PictLibreta, 9, PosLinea, "Equivalente a:   " & DirCliente
        PosLinea = PosLinea + 0.8
        PictPrint_Texto PictLibreta, 2, PosLinea, "DISCIPLINA:   " & Format(Abono, "##0.000")
        DirCliente = "."
        Select Case Redondear(Abono, 2)
          Case 0 To 11:  DirCliente = "INSUFICIENTE"
          Case 12 To 13: DirCliente = "REGULAR"
          Case 14 To 15: DirCliente = "BUENA"
          Case 16 To 18: DirCliente = "MUY BUENA"
          Case 19 To 20: DirCliente = "SOBRESALIENTE"
        End Select
        PictPrint_Texto PictLibreta, 9, PosLinea, "Equivalente a:   " & DirCliente
        PosLinea = PosLinea + 0.8
     Else
        PictLibreta.FontSize = 12
        PictPrint_Texto PictLibreta, 2, PosLinea, "PUNTAJE TOTAL:   " & Format(Saldo, "##.##")
        PosLinea = PosLinea + 0.8
        PictPrint_Texto PictLibreta, 2, PosLinea, "APROVECHAMIENTO:   " & Format(Saldo / Contador, "##.##")
        PosLinea = PosLinea + 0.8
        PictPrint_Texto PictLibreta, 2, PosLinea, "DISCIPLINA:   " & Format(Abono, "##.##")
        PosLinea = PosLinea + 0.8
     End If
     PictLibreta.FontSize = 10
     Select Case Codigo4
       Case "0.00" To "1.99"
            PictPrint_Texto PictLibreta, 5, 24, Director
            PictPrint_Texto PictLibreta, 13.5, 24, Secretario1
            PictPrint_Texto PictLibreta, 5, 24.5, TextoDirector
            PictPrint_Texto PictLibreta, 13.5, 24.5, TextoSecretario1
       Case "2.00" To "3.99"
            PictPrint_Texto PictLibreta, 5, 24, Rector
            PictPrint_Texto PictLibreta, 13.5, 24, Secretario2
            PictPrint_Texto PictLibreta, 5, 24.5, TextoRector
            PictPrint_Texto PictLibreta, 13.5, 24.5, TextoSecretario2
     End Select
     NoMeses = FechaMes(MBFecha)
     NoAnio = FechaAnio(MBFecha)
     PictPrint_Texto PictLibreta, 1.5, 26.5, FechaStrgCiudad(MBFecha)
     'MesesLetras(NoMeses) & " del " & NoAnio
     End If
 End If
End With
RatonNormal
MensajeEncabData = ""
End Sub

Public Sub Aptitud_Promocion_Trimestre1(EsPromocion As Boolean, Curso As String, CodigoAlumno As String)
Dim LineaIni As Single
Dim PuntajeTotal As Single
Dim CursoSup As String
Dim LogoAux  As String
Dim SiPasa As Boolean
Dim PromPT As Currency
Dim PromST As Currency
Dim PromTT As Currency
RatonReloj
PuntajeTotal = 0
PromPT = 0
PromST = 0
PromTT = 0
InicioX = 0.5: InicioY = 0.1
'Pagina = 1
'Iniciamos la impresion
PictLibreta.Cls
PictLibreta.FontName = TipoArial   'TipoTimes
PictLibreta.FontBold = False
'Listamos las notas del Alumnno
Notas_Del_Alumno Curso, CodigoAlumno
CursoSup = Mid$(Curso, 1, 2) & Format(Val(Mid$(Curso, 3, 2)) + 1, "00") & Mid$(Curso, 5, 3)
SiPasa = True
PosColumna = 8.1
With AdoLibreta.Recordset
 If .RecordCount > 0 Then
     CodigoCliente = .Fields("Codigo")
     'MsgBox CodigoL
     If Mid$(CodigoL, 1, 4) < "1.02" Then
        SQLMsg1 = "AÑO LECTIVO: " & Anio_Lectivo
        SQLMsg2 = "I N F O R M E    F I N A L"
        SQLMsg3 = ""
        PictPrint_Grafico PictLibreta, RutaSistema & "\FORMATOS\INFORME.GIF", 0.5, 0.1, 19, 13
        PictEncabezado PictLibreta, 1, 0.5
        PictLibreta.FontBold = True
        PictLibreta.FontSize = 14
        PictPrint_Texto PictLibreta, 7.1, 4, .Fields("Alumno")
     Else
     PosLinea = 1
     If LogoTipo <> "" Then PictPrint_Grafico PictLibreta, LogoTipo, 2, 0.5, 4, 2
     LogoAux = RutaSistema & "\LOGOS\MINISEDU.GIF"
     PictPrint_Grafico PictLibreta, LogoAux, 16, 0.5, 3, 1.2
     PictLibreta.FontBold = True
     PictLibreta.FontSize = 16
     PictPrint_Texto PictLibreta, 1, PosLinea, Institucion1, , 19, True
     PosLinea = PosLinea + 0.6
     PictLibreta.FontSize = 12
     PictPrint_Texto PictLibreta, 1, PosLinea, Institucion2, , 19, True
     PictLibreta.FontSize = 9
     PosLinea = PosLinea + 0.6
     PictLibreta.FontBold = False
     If EsPromocion Then
        PictPrint_Texto PictLibreta, 1, PosLinea, "Dirección: " & Direccion, , 19, True
        PosLinea = PosLinea + 0.6
        PictLibreta.FontUnderline = True
        Cadena = "E-Mail: " & Mail_Colegio & " Telf. " & Telefono1 & " Fax: " & FAX
        PictPrint_Texto PictLibreta, 1, PosLinea, "Dirección: " & Cadena, , 19, True
        PosLinea = PosLinea + 0.5
        PictLibreta.FontUnderline = False
        Cadena = "CODIGO AMIE " & Codigo_AMIE
        PictPrint_Texto PictLibreta, 1, PosLinea, Cadena, , 19, True
        PosLinea = PosLinea + 0.5
        PictLibreta.FontSize = 10
        Cadena = "MANTA - MANABI - ECUADOR"
        PictPrint_Texto PictLibreta, 1, PosLinea, Cadena, , 19, True
        PosLinea = PosLinea + 1
        PictLibreta.FontBold = True
        PictLibreta.FontSize = 12
        If Mid$(Curso, 1, 4) >= "3.03" Then
           PictPrint_Texto PictLibreta, 1, PosLinea, "CERTIFICADO DE APTITUD", , 19, True
        Else
           PictPrint_Texto PictLibreta, 1, PosLinea, "CERTIFICADO ANUAL DE PROMOCION", , 19, True
        End If
        PosLinea = PosLinea + 0.5
        PictPrint_Texto PictLibreta, 1, PosLinea, Anio_Lectivo, , 19, True
        PosLinea = PosLinea + 1
        PictLibreta.FontSize = 10
        PictLibreta.FontBold = False
       'Descipcion del Certificado
        If Mid$(Curso, 1, 4) >= "3.03" Then
           If .Fields("Sexo") = "M" Then
               Cadena = "El Alumno " & .Fields("Alumno") & ", "
           Else
               Cadena = "La Alumna " & .Fields("Alumno") & ", "
           End If
           Cadena = Cadena _
                  & "de conformidad con el Reglamento General de la Ley de Educación, ha cumplido con los " _
                  & "requisitos respectivos y ha obtenido el puntaje necesario, según consta en los libros " _
                  & "que reposan en el archivo de la Secretaría de este Plantel, por lo tanto se encuentra apto " _
                  & "para presentarse a rendir los exámenes de grado."
           PosLinea = PictPrint_Texto_Justifica(PictLibreta, 2, 18.5, PosLinea, Cadena)
           PosLinea = PosLinea + 1
           PictLibreta.FontBold = True
           PictPrint_Texto PictLibreta, 2, PosLinea, "CURSO:"
           PictPrint_Texto PictLibreta, 5.2, PosLinea, Dato_Curso.Descripcion
           PictPrint_Texto PictLibreta, 14.2, PosLinea, "JORNADA:"
           PictPrint_Texto PictLibreta, 16.8, PosLinea, "MATUTINA"
           PosLinea = PosLinea + 0.5
           PictPrint_Texto PictLibreta, 5.2, PosLinea, Dato_Curso.Bachiller
           PictPrint_Texto PictLibreta, 14.2, PosLinea, "AÑO LECTIVO:"
           PictPrint_Texto PictLibreta, 16.8, PosLinea, Anio_Lectivo
        Else
           Cadena = "LA " & UCase(Institucion1) & ", confiere el presente certificado de Promoción "
           If .Fields("Sexo") = "M" Then
               Cadena = Cadena & "al estudiante " & .Fields("Alumno")
           Else
               Cadena = Cadena & "a la estudiante " & .Fields("Alumno")
           End If
           Cadena = Cadena & " del " & Dato_Curso.Curso_Anio & " AÑO " & Dato_Curso.Paralelo & " DE " & Dato_Curso.Especialidad & " "
           PosLinea = PictPrint_Texto_Justifica(PictLibreta, 2, 18.5, PosLinea, Cadena)
           PosLinea = PosLinea + 0.7
           Cadena = "Luego de haberse presentado a las evaluaciones correspondientes al año escolar " _
                  & Anio_Lectivo & " obteniendo las siguientes calificaciones:"
           PosLinea = PictPrint_Texto_Justifica(PictLibreta, 2, 18.5, PosLinea, Cadena)
        End If
        PosLinea = PosLinea + 0.7
        PictLibreta.FontSize = 11
        PictLibreta.FontBold = True
        If Mid$(Curso, 1, 4) >= "3.03" Then
            PictPrint_Texto PictLibreta, 2, PosLinea + 0.1, "ASIGNATURAS"
        Else
            PictPrint_Texto PictLibreta, 2, PosLinea + 0.1, "A R E A S"
        End If
        PictLibreta.FontSize = 8
        LineaIni = PosLinea
        PictPrint_Texto PictLibreta, 8.5, PosLinea, "I"
        PictPrint_Texto PictLibreta, 9.7, PosLinea, "II"
        PictPrint_Texto PictLibreta, 10.8, PosLinea, "III"
        PictPrint_Texto PictLibreta, 11.7, PosLinea, "TOTAL"
        If Mid$(Curso, 1, 4) >= "2.00" Then
           PictPrint_Texto PictLibreta, 12.95, PosLinea, "SUPLE_"
           PictPrint_Texto PictLibreta, 14.1, PosLinea, "PROM."
           PictPrint_Texto PictLibreta, 15.4, PosLinea, "OBSERVACIONES"
        Else
           PictPrint_Texto PictLibreta, 12.95, PosLinea, "PROM."
           PictPrint_Texto PictLibreta, 14.1, PosLinea, "OBSERVACIONES"
        End If
        PosLinea = PosLinea + 0.35
        PictLibreta.FontSize = 7
        PictPrint_Texto PictLibreta, 8.1, PosLinea, "Trimestre"
        PictPrint_Texto PictLibreta, 9.3, PosLinea, "Trimestre"
        PictPrint_Texto PictLibreta, 10.5, PosLinea, "Trimestre"
        PictLibreta.FontSize = 8
        If Mid$(Curso, 1, 4) >= "2.00" Then
           PictPrint_Texto PictLibreta, 12.95, PosLinea, "TORIO"
           PictPrint_Texto PictLibreta, 14.1, PosLinea, "FINAL"
        Else
           PictPrint_Texto PictLibreta, 12.95, PosLinea, "FINAL"
        End If
     Else
        PosLinea = PosLinea + 0.5
        PictLibreta.FontBold = True
        PictLibreta.FontSize = 10
        PictPrint_Texto PictLibreta, 1, PosLinea, "CERTIFICADO APTITUD", , 19, True
        PictLibreta.FontBold = False
        PosLinea = PosLinea + 1
        If .Fields("Sexo") = "M" Then
            Cadena = "El Alumno " & .Fields("Alumno") & ", "
        Else
            Cadena = "La Alumna " & .Fields("Alumno") & ", "
        End If
        Cadena = Cadena _
               & "de conformidad con el Reglamento General de la Ley de Educación, ha cumplido con los " _
               & "requisitos respectivos y ha obtenido el puntaje necesario, según consta en los libros " _
               & "que reposan en el archivo de la Secretaría de este Plantel, por lo tanto se encuentra apto " _
               & "para presentarse a rendir los exámenes de grado."
        PosLinea = PictPrint_Texto_Justifica(PictLibreta, 2, 18.5, PosLinea, Cadena)
        PosLinea = PosLinea + 1
        PictLibreta.FontBold = True
        PictPrint_Texto PictLibreta, 2, PosLinea, "CURSO:"
        PictPrint_Texto PictLibreta, 5.2, PosLinea, Dato_Curso.Descripcion
        PictPrint_Texto PictLibreta, 14.2, PosLinea, "JORNADA:"
        PictPrint_Texto PictLibreta, 16.8, PosLinea, "MATUTINA"
        PosLinea = PosLinea + 0.5
        PictPrint_Texto PictLibreta, 5.2, PosLinea, Dato_Curso.Bachiller
        PictPrint_Texto PictLibreta, 14.2, PosLinea, "AÑO LECTIVO:"
        PictPrint_Texto PictLibreta, 16.8, PosLinea, Anio_Lectivo
        PosLinea = PosLinea + 0.8
        LineaIni = PosLinea
        PictLibreta.Line (1.9, PosLinea)-(18.6, PosLinea), QBColor(0)
        PosLinea = PosLinea + 0.1
        PictLibreta.FontSize = 12
        PictPrint_Texto PictLibreta, 2.1, PosLinea + 0.15, "ASIGNATURAS"
        PictLibreta.FontSize = 8
        PictPrint_Texto PictLibreta, 6.8, PosLinea, "Calificación"
        PictLibreta.FontSize = 12
        PictPrint_Texto PictLibreta, 8.7, PosLinea + 0.15, "CALIFICACION EN LETRAS"
        PictLibreta.FontSize = 10
        PictPrint_Texto PictLibreta, 16.45, PosLinea + 0.1, "APROBADO"
        PosLinea = PosLinea + 0.35
        PictLibreta.FontSize = 9
        PictPrint_Texto PictLibreta, 6.9, PosLinea, "NUMERO"
        PictLibreta.FontBold = False
     End If
     PictLibreta.FontBold = False
     PosLinea = PosLinea + 0.5
     PictLibreta.FontSize = 10
     PictLibreta.Line (1.9, PosLinea)-(18.6, PosLinea), QBColor(0)
     PosLinea = PosLinea + 0.05
     Valor_Prom = 0: Saldo = 0: Contador = 0: Abono = 0: JR = 1.2
     PictLibreta.FontUnderline = False
     PictLibreta.FontBold = False
     Do While Not .EOF
        'MsgBox .Fields("Orden") & " - " & .Fields("Materia") & " - " & .Fields("C") & " - " & .Fields("P") & " - " & .Fields("I")
        If .Fields("CodMatP") = Ninguno Then
            If .Fields("I") Then
               'Sumatoria
                Select Case Mid$(Curso, 1, 4)
                  Case "1.02" To "3.99" '
                       If .Fields("PromPQ") > 0 And .Fields("PromSQ") <= 0 Then
                           Abono_ME = .Fields("PromPQ")
                       ElseIf .Fields("PromPQ") <= 0 And .Fields("PromSQ") > 0 Then
                           Abono_ME = .Fields("PromSQ")
                       Else
                           Abono_ME = (.Fields("PromPQ") + .Fields("PromSQ")) / 2
                       End If
                       Abono_ME = Redondear(Abono_ME, 3)
                       Total = .Fields("PromPQ") + .Fields("PromSQ") + .Fields("PromTQ")
                  Case Else
                       Abono_ME = (.Fields("PQBim1") + .Fields("PQBim2") + .Fields("SQBim1") + .Fields("SQBim2")) / 4
                       Total = .Fields("PQBim1") + .Fields("PQBim2") + .Fields("SQBim1") + .Fields("SQBim2")
                       'If .Fields("P") And .Fields("Orden") <> 9 Then Saldo = Saldo + Valor_Prom
                End Select
               'Disciplina
                Contador = Contador + 1
                If .Fields("Orden") = 9 Then Abono = Redondear(.Fields("PromFinal"), 3)
                If .Fields("Orden") <> 9 Then
                    Si_No = .Fields("C")
                    PictPrint_Texto PictLibreta, 2, PosLinea, ULCase(.Fields("Materia"))
                    PictLibreta.Line (8, PosLinea)-(18.5, PosLinea + 0.4), QBColor(Blanco_Brillante), BF
                    Valor_Prom = Redondear(Total / 3, 3)
                    If .Fields("Supletorio") > 0 Then Valor_Prom = Redondear((Valor_Prom + .Fields("Supletorio")) / 2, 3)
                    If EsPromocion Then
                       IR = PosColumna
                       PictPrint_Nota_Materia PictLibreta, IR, PosLinea, .Fields("PromPQ"), Si_No, 3
                       IR = IR + JR
                       PictPrint_Nota_Materia PictLibreta, IR, PosLinea, .Fields("PromSQ"), Si_No, 3
                       IR = IR + JR
                       PictPrint_Nota_Materia PictLibreta, IR, PosLinea, .Fields("PromTQ"), Si_No, 3
                       IR = IR + JR
                       If Total > 0 And Si_No = False Then PictPrint_Texto PictLibreta, IR, PosLinea, Format(Total, "00.000")
                       IR = IR + JR
                       PromPT = PromPT + .Fields("PromPQ")
                       PromST = PromST + .Fields("PromSQ")
                       PromTT = PromTT + .Fields("PromTQ")
                       If Mid$(Curso, 1, 4) >= "2.00" Then
                          If Si_No = False And .Fields("Supletorio") > 0 Then PictPrint_Texto PictLibreta, IR, PosLinea, Format(.Fields("Supletorio"), "00.000")
                          IR = IR + JR
                       End If
                       If Valor_Prom > 0 And Si_No = False Then PictPrint_Texto PictLibreta, IR, PosLinea, Format(Valor_Prom, "00.000")
                       IR = IR + JR
                       If Valor_Prom >= 12 Then
                          PictPrint_Texto PictLibreta, 15.5, PosLinea, "APROBADO"
                       Else
                          PictPrint_Texto PictLibreta, 15.5, PosLinea, "REPROBADO"
                          SiPasa = False
                       End If
                    Else
                       PictPrint_Nota_Materia PictLibreta, 7, PosLinea, Valor_Prom, Si_No, 3
                       PictLibreta.FontSize = 7
                      'Cambio a letras el valor del promedio final
                       Cadena = Cambio_Letras_Decimales(Valor_Prom, 3)
                       If Si_No = False Then PosLinea = PictPrint_Texto_Justifica(PictLibreta, 8.6, 16, PosLinea + 0.1, Cadena)
                       PictLibreta.FontSize = 10
                       PosLinea = PosLinea - 0.1
                       If Valor_Prom >= 12 Then
                          PictPrint_Texto PictLibreta, 16.4, PosLinea, "APROBADO"
                       Else
                          PictPrint_Texto PictLibreta, 16.4, PosLinea, "REPROBADO"
                       End If
                    End If
                    If .Fields("P") And .Fields("Orden") <> 9 Then Saldo = Saldo + Valor_Prom   '.Fields("PromFinal")
                    PosLinea = PosLinea + 0.45
                    PictLibreta.Line (1.9, PosLinea)-(18.6, PosLinea), QBColor(0)
                    PosLinea = PosLinea + 0.05
                End If
            End If
        End If
       .MoveNext
     Loop
     If EsPromocion Then
        PictLibreta.Line (1.9, LineaIni)-(18.6, LineaIni), QBColor(0)
        PictLibreta.Line (1.9, LineaIni)-(1.9, PosLinea), QBColor(0)
        PictLibreta.Line (8, LineaIni)-(8, PosLinea), QBColor(0)
        PictLibreta.Line (9.2, LineaIni)-(9.2, PosLinea), QBColor(0)
        PictLibreta.Line (10.4, LineaIni)-(10.4, PosLinea), QBColor(0)
        PictLibreta.Line (11.6, LineaIni)-(11.6, PosLinea), QBColor(0)
        PictLibreta.Line (12.8, LineaIni)-(12.8, PosLinea), QBColor(0)
        PictLibreta.Line (14, LineaIni)-(14, PosLinea), QBColor(0)
        If Mid$(Curso, 1, 4) >= "2.00" Then PictLibreta.Line (15.3, LineaIni)-(15.3, PosLinea), QBColor(0)
        PictLibreta.Line (18.6, LineaIni)-(18.6, PosLinea), QBColor(0)
     Else
        PosLinea = PosLinea - 0.05
        PictLibreta.Line (1.9, LineaIni)-(1.9, PosLinea), QBColor(0)
        PictLibreta.Line (6.7, LineaIni)-(6.7, PosLinea), QBColor(0)
        PictLibreta.Line (8.5, LineaIni)-(8.5, PosLinea), QBColor(0)
        PictLibreta.Line (16.3, LineaIni)-(16.3, PosLinea), QBColor(0)
        PictLibreta.Line (18.6, LineaIni)-(18.6, PosLinea), QBColor(0)
     End If
     PictLibreta.FontBold = True
     Contador = Contador - 1
     If EsPromocion Then
        PictLibreta.FontSize = 8
        PictLibreta.Line (1.9, PosLinea)-(18.6, PosLinea), QBColor(0)
        PosLinea = PosLinea + 0.05
        PictPrint_Texto PictLibreta, 2, PosLinea, "T O T A L:"
        IR = PosColumna
        PictPrint_Nota_Materia PictLibreta, IR, PosLinea, PromPT, Si_No, 3
        IR = IR + JR
        PictPrint_Nota_Materia PictLibreta, IR, PosLinea, PromST, Si_No, 3
        IR = IR + JR
        PictPrint_Nota_Materia PictLibreta, IR, PosLinea, PromTT, Si_No, 3
        If Mid$(Curso, 1, 4) >= "2.00" Then
           IR = IR + (JR * 3)
           PictPrint_Texto PictLibreta, IR, PosLinea, Format(Saldo / Contador, "00.000")
        Else
           IR = IR + (JR * 2)
           PictPrint_Texto PictLibreta, IR, PosLinea, Format(Saldo / Contador, "00.000")
        End If
        PictLibreta.FontSize = 10
        PosLinea = PosLinea + 0.6
        If Mid$(Curso, 1, 4) >= "3.03" Then
            PictLibreta.FontSize = 12
'''            PictPrint_Texto PictLibreta, 2, PosLinea, "PUNTAJE TOTAL:"
'''            PictPrint_Texto PictLibreta, 9.5, PosLinea, Format(Saldo, "00.000")
'''            PosLinea = PosLinea + 0.6
            PictPrint_Texto PictLibreta, 2, PosLinea, "PROMEDIO GENERAL:"
            PictPrint_Texto PictLibreta, 9.5, PosLinea, Format(Saldo / Contador, "00.000")
            PictPrint_Texto PictLibreta, 12, PosLinea, "APROBADO"
            PosLinea = PosLinea + 0.6
            PictPrint_Texto PictLibreta, 2, PosLinea, "DISCIPLINA:"
            PictPrint_Texto PictLibreta, 9.5, PosLinea, Format(Abono, "00.00")
            PictPrint_Texto PictLibreta, 12, PosLinea, "APROBADO"
            PosLinea = PosLinea + 1
            PictLibreta.FontBold = False
        Else
            PictLibreta.FontSize = 12
'''            PictPrint_Texto PictLibreta, 2, PosLinea, "T O T A L:"
'''            PictPrint_Texto PictLibreta, 11, PosLinea, Format(Saldo, "00.000")
'''            PosLinea = PosLinea + 0.6
            
            PictPrint_Texto PictLibreta, 2, PosLinea, "PROMEDIO DE DISCIPLINA:"
            PictPrint_Texto PictLibreta, 11, PosLinea, Format(Abono, "00.000")
            DirCliente = "."
            Select Case Redondear(Abono, 2)
              Case 0 To 11.49:  DirCliente = "INSUFICIENTE"
              Case 11.5 To 13.49: DirCliente = "REGULAR"
              Case 13.5 To 15.49: DirCliente = "BUENA"
              Case 15.5 To 18.49: DirCliente = "MUY BUENA"
              Case 18.5 To 20: DirCliente = "SOBRESALIENTE"
            End Select
            PictPrint_Texto PictLibreta, 13, PosLinea, DirCliente
            PosLinea = PosLinea + 0.6
            
            PictPrint_Texto PictLibreta, 2, PosLinea, "PROMEDIO GLOBAL DE RENDIMIENTO:"
            PictPrint_Texto PictLibreta, 11, PosLinea, Format(Saldo / Contador, "##0.000")
            DirCliente = "."
            Select Case Redondear(Saldo / Contador)
              Case 0 To 11.99:  DirCliente = "INSUFICIENTE"
              Case 11.5 To 13.49: DirCliente = "REGULAR"
              Case 13.5 To 15.49: DirCliente = "BUENA"
              Case 15.5 To 18.49: DirCliente = "MUY BUENA"
              Case 18.5 To 20: DirCliente = "SOBRESALIENTE"
            End Select
            PictPrint_Texto PictLibreta, 13, PosLinea, DirCliente
            PosLinea = PosLinea + 0.6
            PictLibreta.FontBold = False
            Cadena = "Por lo tanto "
            If SiPasa Then Cadena = Cadena & "SI se lo promueve al " Else Cadena = Cadena & "NO se lo promueve al "
            If Mid$(Curso, 1, 4) >= "2.10" Then
               Cadena = Cadena & Dato_Curso.Curso_Superior & " " & " AÑO DE " & Dato_Curso.Bachiller & " "
            Else
               Cadena = Cadena & Dato_Curso.Curso_Superior & " " & " AÑO DE " & Dato_Curso.Especialidad & " "
            End If
            PosLinea = PictPrint_Texto_Justifica(PictLibreta, 2, 18.5, PosLinea, Cadena)
            PosLinea = PosLinea + 1
            PictPrint_Texto PictLibreta, 2, PosLinea, "Asi consta en los libros de calificaciones de la Secretaria del Plantel."
            NoMeses = FechaMes(MBFecha)
            NoAnio = FechaAnio(MBFecha)
            PosLinea = PosLinea + 1
        End If
     Else
        PictLibreta.FontSize = 11
        PictPrint_Texto PictLibreta, 2, PosLinea, "PUNTAJE TOTAL:"
        PictPrint_Texto PictLibreta, 6.8, PosLinea, Format(Saldo, "00.000")
        PosLinea = PosLinea + 0.6
        PictPrint_Texto PictLibreta, 2, PosLinea, "PROMEDIO GENERAL:"
        PictPrint_Texto PictLibreta, 6.8, PosLinea, Format(Saldo / Contador, "00.000")
        
        PictPrint_Texto PictLibreta, 8.5, PosLinea, "APROBADO"
        PosLinea = PosLinea + 0.6
        PictPrint_Texto PictLibreta, 2, PosLinea, "DISCIPLINA:"
        PictPrint_Texto PictLibreta, 6.8, PosLinea, Format(Abono, "00.00")
        PictPrint_Texto PictLibreta, 8.5, PosLinea, "APROBADO"
        PosLinea = PosLinea + 1
        PictLibreta.FontBold = False
     End If
     If EsPromocion Then
        PictPrint_Texto PictLibreta, 13.5, PosLinea, FechaStrgCiudad(MBFecha)
     Else
        PictPrint_Texto PictLibreta, 2, PosLinea, FechaStrgCiudad(MBFecha)
     End If
     PictLibreta.FontSize = 10
     PosLinea = 24
     Select Case Codigo4
       Case "0.00" To "1.99"
            PictPrint_Texto PictLibreta, 2, PosLinea, Director
            PictPrint_Texto PictLibreta, 10.5, PosLinea, Secretario1
            PosLinea = PosLinea + 0.4
            PictPrint_Texto PictLibreta, 2, PosLinea, TextoDirector
            PictPrint_Texto PictLibreta, 10.5, PosLinea, TextoSecretario1
       Case "2.00" To "3.99"
            PictPrint_Texto PictLibreta, 2, PosLinea, Rector
            PictPrint_Texto PictLibreta, 10.5, PosLinea, Secretario2
            PosLinea = PosLinea + 0.4
            PictPrint_Texto PictLibreta, 2, PosLinea, TextoRector
            PictPrint_Texto PictLibreta, 10.5, PosLinea, TextoSecretario2
     End Select
     PosLinea = PosLinea + 1.5
     End If
 End If
End With
RatonNormal
MensajeEncabData = ""
End Sub

Public Sub Aptitud_Promocion_Quimestre(TipoObjeto As Object, EsPromocion As Boolean, Curso As String, CodigoAlumno As String)
Dim PosLineaIni As Single
Dim PosLineaFin As Single
Dim AnchoPict As Single
Dim LogoAux As String
Dim CantPromConducta As Byte
Dim Disciplina As Single
Dim SiPromuebe As Boolean
Dim DisciplinaPQ As Currency
Dim DisciplinaSQ As Currency
Dim ParteEntera As Currency
Dim ParteDecimal As Currency
Dim Promedio As Currency
Dim Cambio_a_Leras As String
Dim LongLinea As Byte
Dim EsPreBasica As Boolean
Dim EsCualitativa As Boolean
Dim EsCualitativa2 As Boolean
Dim EsPromovido As Boolean
Dim FinCalculo As Boolean

'Dim Especialidad As String
  RatonReloj
  EsPromovido = True
  FinCalculo = True
 'MsgBox EsPromocion
  EsPreBasica = False
  If Mid$(Curso, 1, 4) <= "1.01" Then EsPreBasica = True
  With AdoPlantel.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("CodigoE like '" & Curso & "' ")
       If Not .EOF Then
          NombreDocente = .Fields("Dirigente")
          SexoDocente = .Fields("Sexo")
       End If
   End If
  End With
  
SiPromuebe = True
Disciplina = Procesar_Disciplinas(CodigoAlumno, Curso)
'Listamos las notas del Alumnno
Notas_Del_Alumno Curso, CodigoAlumno
InicioX = 0.5: InicioY = 0.1: Abono = 0
'Pagina = 1
'Iniciamos la impresion
If TypeOf TipoObjeto Is PictureBox Then TipoObjeto.Cls
TipoObjeto.FontName = TipoVerdana 'TipoHelvetica
'TipoObjeto.ForeColor = QBColor(Negro)
TipoObjeto.width = 21
TipoObjeto.Height = 29.7

AnchoPict = 20
TipoObjeto.FontBold = False

With AdoLibreta.Recordset
 If .RecordCount > 0 Then
     CodigoCliente = .Fields("Codigo")
    'MsgBox CodigoL
     If Mid$(CodigoL, 1, 4) < "1.01" Then
        EsPreBasica = True
        SQLMsg1 = "AÑO LECTIVO: " & Anio_Lectivo
        SQLMsg2 = "I N F O R M E    F I N A L"
        SQLMsg3 = ""
        cPrint.printImagen RutaSistema & "\FORMATOS\INFORME.GIF", 0.5, 0.1, 19, 13
        PictEncabezado 1, 0.5
        'TipoObjeto.FontName = TipoHelvetica
        'TipoObjeto.ForeColor = QBColor(Negro)
        TipoObjeto.FontBold = True
        TipoObjeto.FontSize = 14
        PictPrint_Texto 7.1, 4, .Fields("Alumno")
     Else
     With AdoMatriculas.Recordset
      If .RecordCount > 0 Then
         .MoveFirst
         .Find ("Codigo = '" & CodigoCliente & "' ")
          If Not .EOF Then
             CantPromConducta = 0
             If .Fields("ConductaPQ1") > 0 Then CantPromConducta = CantPromConducta + 1
             If .Fields("ConductaPQ2") > 0 Then CantPromConducta = CantPromConducta + 1
             If .Fields("ConductaPQ3") > 0 Then CantPromConducta = CantPromConducta + 1
             If .Fields("ConductaSQ1") > 0 Then CantPromConducta = CantPromConducta + 1
             If .Fields("ConductaSQ2") > 0 Then CantPromConducta = CantPromConducta + 1
             If .Fields("ConductaSQ3") > 0 Then CantPromConducta = CantPromConducta + 1
             If CantPromConducta <= 0 Then CantPromConducta = 1
             DisciplinaPQ = Redondear((.Fields("ConductaPQ1") + .Fields("ConductaPQ2") + .Fields("ConductaPQ3")) / 3, Dec_Nota)
             Select Case Anio_Lectivo
               Case Is >= "2014 - 2015", Ninguno
                    DisciplinaSQ = Redondear(.Fields("ConductaSQ3"), Dec_Nota)
               Case Else
                    DisciplinaSQ = Redondear((.Fields("ConductaSQ1") + .Fields("ConductaSQ2") + .Fields("ConductaSQ3")) / 3, Dec_Nota)
             End Select
             Abono = Redondear((.Fields("ConductaPQ1") + .Fields("ConductaPQ2") + .Fields("ConductaPQ3") + .Fields("ConductaSQ1") + .Fields("ConductaSQ2") + .Fields("ConductaSQ3")) / CantPromConducta, Dec_Nota)
          End If
      End If
     End With
     SQLMsg2 = "AÑO LECTIVO: " & Anio_Lectivo
     If EsPromocion Then
        SQLMsg1 = "CERTIFICADO DE PROMOCIÓN"
     Else
        SQLMsg1 = "INFORME FINAL ACADEMICO"
     End If
     SQLMsg3 = "JORNADA MATUTINA"
     TipoObjeto.FontBold = True
     TipoObjeto.FontSize = 19
     PosLinea = 0.5
     
     If Mid$(CodigoL, 1, 4) >= "1.01" Then
        LogoAux = RutaSistema & "\LOGOS\LOGOPROMOCIONIZQ.gif"
        cPrint.printImagen LogoAux, 2, 0.3, 4, 2
        
        LogoAux = RutaSistema & "\LOGOS\LOGOPROMOCIONCEN.gif"
        cPrint.printImagen LogoAux, 8.5, 0.3, 4, 2
        
        LogoAux = RutaSistema & "\LOGOS\LOGOPROMOCIONDER.gif"
        cPrint.printImagen LogoAux, 16, 0.3, 4, 2
        
        TipoObjeto.FontSize = 14
        PosLinea = PosLinea + 2.2
        PictPrint_Texto 0.1, PosLinea, Institucion1, , 21, True
        PosLinea = PosLinea + 0.7
        TipoObjeto.FontSize = 14
        If TipoObjeto.TextWidth(Institucion2) > 18 Then TipoObjeto.FontSize = 12
        PictPrint_Texto 0.1, PosLinea, Institucion2, , 21, True
        PosLinea = PosLinea + 0.8
        TipoObjeto.FontSize = 10
        TipoObjeto.FontBold = False
        PictPrint_Texto 0.1, PosLinea, "Codigo AMIE: " & Codigo_AMIE, , AnchoPict, True
        PosLinea = PosLinea + 0.5
        Cadena = ULCase(NombreCiudad & ", " & Direccion)
        PictPrint_Texto 0.1, PosLinea, Cadena, , AnchoPict, True
        PosLinea = PosLinea + 0.45
        Cadena = "Teléfonos: " & Telefono1 & "/" & Telefono2 & " - Correo: " & Mail_Colegio
        PictPrint_Texto 0.1, PosLinea, Cadena, , AnchoPict, True
        PosLinea = PosLinea + 0.6
     Else
        PosLinea = PosLinea + 4
     End If
     TipoObjeto.FontBold = True
     TipoObjeto.FontSize = 14
     If SQLMsg1 <> "" Then
        PictPrint_Texto 0.1, PosLinea, SQLMsg1, , AnchoPict, True
        PosLinea = PosLinea + 0.6
     End If
     TipoObjeto.FontBold = False
     TipoObjeto.FontSize = 10
     If SQLMsg2 <> "" Then
        PictPrint_Texto 0.1, PosLinea, SQLMsg2, , AnchoPict, True
        PosLinea = PosLinea + 0.5
     End If
     TipoObjeto.FontBold = False
     If SQLMsg3 <> "" Then
        PictPrint_Texto 0.1, PosLinea, SQLMsg3, , AnchoPict, True
        PosLinea = PosLinea + 0.6
     End If
     TipoObjeto.FontUnderline = False
     PosLinea = PosLinea + 0.4
     TipoObjeto.FontSize = 10
     Contador = 0
     Do While Not .EOF
        If .Fields("I") And .Fields("P") And .Fields("CodMatP") = Ninguno Then Contador = Contador + 1
       .MoveNext
     Loop
    .MoveFirst
    'Inicio del cuadro
     PosLineaIni = PosLinea
     If EsPromocion Then
        TipoObjeto.FontBold = False
        Cadena1 = "De Conformidad con lo prescrito en el Art. 197 del Reglamento General a la Ley Orgánica de Educación " _
                & "Intercultural y demás normativas vigentes, certifico que "
        If .Fields("Sexo") = "M" Then Cadena1 = Cadena1 & "el " Else Cadena1 = Cadena1 & " la "
        Cadena1 = Cadena1 & "estudiante: ^" & .Fields("Alumno") & "~, del ^"
        'If Mid$(Curso, 1, 1) >= "3" Then Cadena1 = Cadena1 & Dato_Curso.Curso_Texto & " DE "
        Cadena1 = Cadena1 & Dato_Curso.Nombre_Largo & "~, obtuvo las siguientes calificaciones durante el presente año lectivo: "
        PosLinea = PictPrint_Texto_Justifica(TipoObjeto, 2, 19.5, PosLinea, Cadena1)
        PosLinea = PosLinea + 0.8
        PosLineaIni = PosLinea
        PosLineaFin = PosLineaIni + (Contador * 0.45) + 0.6
        TipoObjeto.FontBold = True
        PictPrint_Texto 3, PosLinea + 0.3, "A S I G N A T U R A S"
        TipoObjeto.FontSize = 10
        PictPrint_Texto 13, PosLinea + 0.05, "C A L I F I C A C I O N E S"
        PosLinea = PosLinea + 0.55
        PictPrint_Texto 11, PosLinea, "N Ú M E R O"
        PictPrint_Texto 15.5, PosLinea, "L E T R A S"
        PosLinea = PosLinea + 0.5
     Else
        'PictPrint_Texto 2, PosLinea, "CURSO: " & Cursos
        PosLinea_Aux = PictPrint_Texto_Multiple(TipoObjeto, 2, PosLinea, "CURSO: " & Dato_Curso.Nombre_Largo, 18)
        PosLinea = PosLinea + 0.9
        PictPrint_Texto 2, PosLinea, "NOMBRE DEL ESTUDIANTE: " & .Fields("Alumno")
        PosLinea = PosLinea + 0.6
        PosLineaIni = PosLinea - 0.1
        PosLineaFin = PosLineaIni + (Contador * 0.45) + 0.6
        
        TipoObjeto.FontSize = 12
        TipoObjeto.FontBold = True
        PictPrint_Texto 2, PosLinea + 0.15, "ASIGNATURAS"
        TipoObjeto.FontSize = 7
        PictPrint_Texto 9.4, PosLinea, "I"
        PictPrint_Texto 10.9, PosLinea, "II"
        PictPrint_Texto 12, PosLinea, "PUNTAJE"
        PictPrint_Texto 13.5, PosLinea, "PROMEDIO"
        PictPrint_Texto 15.5, PosLinea, "SUPLE_"
        PictPrint_Texto 17, PosLinea, "REMEDIAL"
        PictPrint_Texto 19, PosLinea, "PROMEDIO"
        PosLinea = PosLinea + 0.35
        PictPrint_Texto 9.1, PosLinea, "QUIM"
        PictPrint_Texto 10.6, PosLinea, "QUIM"
        PictPrint_Texto 12, PosLinea, "TOTAL"
        PictPrint_Texto 13.5, PosLinea, "QUIMESTRE"
        PictPrint_Texto 15.5, PosLinea, "TORIO"
        PictPrint_Texto 19, PosLinea, "FINAL"
        PosLinea = PosLinea + 0.5
     End If
     TipoObjeto.FontBold = False
     PosLinea = PosLinea + 0.05
     Saldo = 0: Contador = 0: Debe = 0
     TipoObjeto.FontUnderline = False
     TipoObjeto.FontBold = False
     Contador = 0
     EsPromovido = True
     Do While Not .EOF
        EsCualitativa = .Fields("C")
        EsCualitativa2 = .Fields("C2")
        If Mid$(Curso, 1, 4) <= "1.01" Then EsCualitativa = True
        TipoObjeto.FontSize = 10
        Total = .Fields("PromPQ") + .Fields("PromSQ")
        Abono_ME = .Fields("PromFinal")
        If Len(CStr((Abono_ME - Int(Abono_ME)))) >= 5 Then Abono_ME = Truncar(Abono_ME, 2)
        ParteEntera = Int(Abono_ME)
        ParteDecimal = (Abono_ME - ParteEntera) * 100
        'MsgBox Abono_ME & vbCrLf & ParteEntera & vbCrLf & ParteDecimal
        Cambio_a_Leras = Cambio_Letras(ParteEntera) & " COMA "
        If ParteDecimal < 10 Then
           Cambio_a_Leras = Cambio_a_Leras & "CERO " & Cambio_Letras(ParteDecimal)
        Else
           If ParteDecimal = 0 Then
              Cambio_a_Leras = Cambio_a_Leras & "CERO CERO"
           Else
              Cambio_a_Leras = Cambio_a_Leras & Cambio_Letras(ParteDecimal)
           End If
        End If
         
       'MsgBox .Fields("CodMatP") & ", Orden: " & .Fields("Orden") & " - " & .Fields("Materia") & " - C: " & .Fields("C") & " - P: " & .Fields("P") & " - I: " & .Fields("I")
        
        If .Fields("CodMatP") = Ninguno Then
            If .Fields("I") Then
               'Sumatoria .Fields("P") And
                If .Fields("Orden") = 9 Then
                    FinCalculo = False
                    If Contador <= 0 Then Contador = 1
                     If (Debe / Contador) < Nota_Rojo Then EsPromovido = False
                    'Fin del listado de las materias
                     TipoObjeto.FontSize = 11
                     TipoObjeto.FontBold = True
                     PosLinea = PosLinea + 0.1
                     If EsPromocion Then
                        If Contador = 0 Then Contador = 1
                        PosLinea = PosLinea + 0.1
                        TipoObjeto.FontBold = True
                        TipoObjeto.FontSize = 12
                        PictPrint_Texto 2, PosLinea + 0.1, "PROMEDIO GENERAL"
                        
                        TipoObjeto.FontSize = 11
                        TipoObjeto.FontBold = False
                        Promedio = Debe / Contador
                       'MsgBox Debe & vbCrLf & Contador & vbCrLf & "Nota = " & Promedio & vbCrLf & "Redondeada = " & Redondear_2Dec(Promedio)
                        
                        If Mid$(Curso, 1, 4) <= "1.01" Then
                           PictPrint_Nota_Materia 11.7, PosLinea, Promedio, EsCualitativa, Dec_Nota, EsPreBasica, EsCualitativa2
                        Else
                           PictPrint_Texto 11.7, PosLinea + 0.1, Cambio_Punto_x_Coma(Promedio), True, 1
                        End If
                        Promedio = Redondear_2Dec(Promedio)
                        
                        ParteEntera = Int(Promedio)
                        
                        ParteDecimal = (Promedio - ParteEntera) * 100
                        
                        Cambio_a_Leras = Cambio_Letras(ParteEntera) & " COMA "
                        
                        If ParteDecimal < 10 Then
                           Cambio_a_Leras = Cambio_a_Leras & "CERO " & Cambio_Letras(ParteDecimal)
                        Else
                           Cambio_a_Leras = Cambio_a_Leras & Cambio_Letras(ParteDecimal)
                        End If
                        TipoObjeto.FontSize = 10
                        PictPrint_Texto 13.5, PosLinea + 0.1, Cambio_a_Leras
                        PosLinea = PosLinea + 1
                        TipoObjeto.FontBold = True
                        TipoObjeto.FontSize = 11
                        PictPrint_Texto 2, PosLinea, "EVALUACIÓN DEL COMPORTAMIENTO:"
                        CantPromConducta = 0
                        If Real1 > 0 Then CantPromConducta = CantPromConducta + 1
                        If Real2 > 0 Then CantPromConducta = CantPromConducta + 1
                        If Real3 > 0 Then CantPromConducta = CantPromConducta + 1
                        If Real4 > 0 Then CantPromConducta = CantPromConducta + 1
                        If CantPromConducta <= 0 Then CantPromConducta = 1
                        Abono = Redondear((Real1 + Real2 + Real3 + Real4) / CantPromConducta)
                        Select Case Anio_Lectivo
                          Case Is >= "2014 - 2015", Ninguno
                               Disciplina = Redondear(DisciplinaSQ, Dec_Nota)
                          Case Else
                               Disciplina = Redondear((DisciplinaPQ + DisciplinaSQ) / 2, Dec_Nota)
                        End Select
                        TipoObjeto.FontSize = 14
                        PictPrint_Nota_Materia 11.8, PosLinea, Disciplina, True, Dec_Nota, , True
                        TipoObjeto.FontBold = False
                        TipoObjeto.FontSize = 7
                        PosLinea = PictPrint_Texto_Multiple(TipoObjeto, 13.6, PosLinea, Equivalencia(CCur(Disciplina), , True), 6.3)
                        PosLinea = PosLinea - 1.5
                        TipoObjeto.FontSize = 11
                     Else
                        TipoObjeto.Line (1.9, PosLinea)-(20.5, PosLinea), QBColor(0)
                        PosLinea = PosLinea + 0.1
                        If Contador = 0 Then Contador = 1
                        If Mid$(Curso, 1, 4) > "1.01" Then
                           PictPrint_Texto 2, PosLinea, "PUNTAJE TOTAL:"
                           TipoObjeto.FontSize = 10
                           If Len(Format(Debe, "##.00")) > 5 Then
                              PictPrint_Texto 18.8, PosLinea, Format(Debe, "##.00")
                           Else
                              PictPrint_Texto 19, PosLinea, Format(Debe, "##.00")
                           End If
                        End If
                        PosLinea = PosLinea + 0.8
                        TipoObjeto.FontSize = 11
                        PictPrint_Texto 2, PosLinea, "Promedio de Aprovechamiento:"
                        TipoObjeto.FontSize = 10
                        
                        If Mid$(Curso, 1, 4) <= "1.01" Then
                           PictPrint_Nota_Materia 19, PosLinea, (Debe / Contador), EsCualitativa, Dec_Nota, EsPreBasica
                        Else
                           PictPrint_Texto 19, PosLinea, Cambio_Punto_x_Coma(Debe / Contador), True, 1
                        End If
                        PosLinea = PosLinea + 0.8
                        TipoObjeto.FontSize = 11
                        PictPrint_Texto 2, PosLinea, "Evaluación del Comportamiento:"
                        TipoObjeto.FontSize = 10
                        If FormatoLibreta = "QUIMESTRE" Then
                           Disciplina = Redondear(DisciplinaSQ, Dec_Nota)
                        Else
                           Disciplina = Redondear((DisciplinaPQ + DisciplinaSQ) / 2, Dec_Nota)
                        End If
                        PictPrint_Nota_Materia 9.3, PosLinea, DisciplinaPQ, True
                        PictPrint_Nota_Materia 10.8, PosLinea, DisciplinaSQ, True
                        PictPrint_Nota_Materia 13.8, PosLinea, Disciplina, True
                        PosLinea = PosLinea - 1.7
                     End If
                    'Fin del cuadro
                     PosLineaFin = PosLinea
                     If EsPromocion Then
                        TipoObjeto.Line (1.9, PosLineaIni)-(20, PosLineaIni + 1), QBColor(Negro), B
                        TipoObjeto.Line (10.9, PosLineaIni + 0.5)-(20, PosLineaIni + 0.5), QBColor(Negro)
                        TipoObjeto.Line (10.9, PosLineaIni)-(10.9, PosLineaFin + 2), QBColor(Negro)
                        TipoObjeto.Line (13.4, PosLineaIni + 0.5)-(13.4, PosLineaFin + 2), QBColor(Negro)
                        
                        TipoObjeto.Line (1.9, PosLineaIni + 1)-(20, PosLineaFin + 2), QBColor(Negro), B
                        TipoObjeto.Line (1.9, PosLineaFin)-(20, PosLineaFin), QBColor(Negro)
                        TipoObjeto.Line (1.9, PosLineaFin + 1)-(20, PosLineaFin + 1), QBColor(Negro)
                     Else
                        'TipoObjeto.Line (1.9, PosLinea)-(20.5, PosLinea), QBColor(0)
                        TipoObjeto.Line (1.9, PosLineaIni)-(20.5, PosLineaFin), QBColor(Negro), B
                        TipoObjeto.Line (8.9, PosLineaIni)-(10.4, PosLineaFin), QBColor(Negro), B
                        TipoObjeto.Line (11.9, PosLineaIni)-(13.4, PosLineaFin), QBColor(Negro), B
                        TipoObjeto.Line (15.2, PosLineaIni)-(16.9, PosLineaFin), QBColor(Negro), B
                        TipoObjeto.Line (18.9, PosLineaIni)-(20.5, PosLineaFin), QBColor(Negro), B
                        TipoObjeto.Line (1.9, PosLineaIni)-(20.5, PosLineaIni + 0.9), QBColor(Negro), B
                        TipoObjeto.Line (1.9, PosLineaFin)-(20.5, PosLineaFin + 2.3), QBColor(Negro), B
                        TipoObjeto.Line (1.9, PosLineaFin + 0.7)-(20.5, PosLineaFin + 1.5), QBColor(Negro), B
                     End If
                     PosLinea = PosLinea + 2.4
                Else
                    If Not .Fields("C") Then Saldo = Saldo + Abono_ME
                   'PictPrint_Texto 2, PosLinea, .Fields("Materia")
                    If EsPromocion Then
                       PosLinea = PictPrint_Texto_Multiple(TipoObjeto, 2, PosLinea, .Fields("Materia"), 8.5)
                       If Abono_ME > 0 Then
                          If Mid$(Curso, 1, 4) <= "1.01" Then
                             PictPrint_Nota_Materia 11.85, PosLinea, Abono_ME, EsCualitativa, Dec_Nota, EsPreBasica, EsCualitativa2
                          Else
                             If EsCualitativa Then
                                PictPrint_Nota_Materia 11.85, PosLinea, Abono_ME, True, Dec_Nota, EsPreBasica, False
                             ElseIf EsCualitativa2 Then
                                PictPrint_Nota_Materia 11.85, PosLinea, Abono_ME, False, Dec_Nota, EsPreBasica, True
                             ElseIf EsCualitativa And EsCualitativa2 Then
                                PictPrint_Nota_Materia 11.85, PosLinea, Abono_ME, True, Dec_Nota, EsPreBasica, True
                             Else
                                PictPrint_Texto 11.7, PosLinea, Cambio_Punto_x_Coma(Abono_ME), True, 1
                             End If
                          End If
                          If Not EsCualitativa And Not EsCualitativa2 Then
                             TipoObjeto.FontSize = 10
                             PictPrint_Texto 13.5, PosLinea, Cambio_a_Leras
                          End If
                          If EsCualitativa2 Then
                             TipoObjeto.FontSize = 7
                             PosLinea = PictPrint_Texto_Multiple(TipoObjeto, 13.6, PosLinea - 0.3, Equivalencia(CCur(Disciplina), , , , True), 6.3)
                             PosLinea = PosLinea - 0.3
                          End If
                          
                       End If
                    Else
                       PosLinea = PictPrint_Texto_Multiple(TipoObjeto, 2, PosLinea, .Fields("Materia"), 6.5)
                       PictPrint_Nota_Materia 9.1, PosLinea, .Fields("PromPQ"), EsCualitativa, Dec_Nota, EsPreBasica, EsCualitativa2
                       PictPrint_Nota_Materia 10.6, PosLinea, .Fields("PromSQ"), EsCualitativa, Dec_Nota, EsPreBasica, EsCualitativa2
                       PictPrint_Nota_Materia 12.1, PosLinea, Total, EsCualitativa, Dec_Nota, EsPreBasica, EsCualitativa2
                       PictPrint_Nota_Materia 13.7, PosLinea, Total / 2, EsCualitativa, Dec_Nota, EsPreBasica, EsCualitativa2
                       PictPrint_Nota_Materia 15.5, PosLinea, .Fields("Supletorio"), EsCualitativa, Dec_Nota, EsPreBasica, EsCualitativa2
                       PictPrint_Nota_Materia 17.2, PosLinea, .Fields("Remedial"), EsCualitativa, Dec_Nota, EsPreBasica, EsCualitativa2
                       PictPrint_Nota_Materia 19.1, PosLinea, Abono_ME, EsCualitativa, Dec_Nota, EsPreBasica, EsCualitativa2
                    End If
                    If FinCalculo And Mid$(Curso, 1, 4) >= "2.00" Then
                       If Abono_ME < Nota_Rojo Then EsPromovido = False
                    End If
                    PosLinea = PosLinea + 0.45
                   'MsgBox Redondear(.Fields("PromFinal"), Dec_Nota)
                    If .Fields("P") Then
                        If Abono_ME > 0 Then
                           Contador = Contador + 1
                           Debe = Debe + Abono_ME
                           'MsgBox Contador & vbCrLf & .Fields("Materia") & vbCrLf & Abono_ME & vbCrLf & Debe
                        End If
                    End If
                End If
            End If
        End If
        If Abono_ME < Nota_Rojo And Val(.Fields("CodMat")) <= 997 Then SiPromuebe = False
       .MoveNext
     Loop
     TipoObjeto.FontSize = 10
     TipoObjeto.FontBold = False
     Select Case Codigo4
       Case "0.00" To "2.99"
            PosLinea = PosLinea + 0.35
     End Select
     
     If EsPromovido Then
        Cadena1 = "Por lo tanto es promovido al ^" & Dato_Curso.Curso_Superior _
                & "~. Para Certificar suscriben en unidad de acto "
     Else
        Cadena1 = "Por lo tanto ^no es promovido al curso inmediato superior~. Para Certificar suscriben en unidad de acto "
     End If
     Select Case Codigo4
       Case "0.00" To "1.99"
            If SexoDirector = "M" Then
               Cadena1 = Cadena1 & "el "
            Else
               Cadena1 = Cadena1 & "la "
            End If
            Cadena1 = Cadena1 & ULCase(TextoDirector) & " "
            If SexoDocente = "M" Then
               Cadena1 = Cadena1 & "y el "
            Else
               Cadena1 = Cadena1 & "y la "
            End If
            'Cadena1 = Cadena1 & "Docente del Plantel del ^" & Dato_Curso.Nombre_Largo & "~. "
            Cadena1 = Cadena1 & ULCase(TextoSecretario1) & " del Plantel."
       Case "2.00" To "5.99"
            If SexoRector = "M" Then
               Cadena1 = Cadena1 & "el "
            Else
               Cadena1 = Cadena1 & "la "
            End If
            Cadena1 = Cadena1 & ULCase(TextoRector) & " "
            If SexoSecre2 = "M" Then
               Cadena1 = Cadena1 & "y el "
            Else
               Cadena1 = Cadena1 & "y la "
            End If
            Cadena1 = Cadena1 & ULCase(TextoSecretario2) & " del Plantel."
     End Select
     PosLinea = PictPrint_Texto_Justifica(TipoObjeto, 2, 19.5, PosLinea, Cadena1)
     NoMeses = FechaMes(MBFecha)
     NoAnio = FechaAnio(MBFecha)
     PosLinea = PosLinea + 0.8
     PictPrint_Texto 13.5, PosLinea, FechaStrgCiudad(MBFecha)
     PosLinea = PosLinea + 0.55
     Select Case Codigo4
       Case "0.00" To "1.99"
            LongLinea = Len(NombreDocente)
            If LongLinea < 2 Then LongLinea = 3
            PictPrint_Texto 2.5, 25, String(30, "_")
            PictPrint_Texto 11, 25, String(30, "_")
            PictPrint_Texto 3, 25.5, Director
            PictPrint_Texto 11.5, 25.5, ULCase(Secretario1)
            PictPrint_Texto 3, 26, TextoDirector
            PictPrint_Texto 11.5, 26, TextoSecretario1
       Case "2.00" To "5.99"
            PictPrint_Texto 3, 25.5, String(30, "_")
            PictPrint_Texto 12, 25.5, String(30, "_")
            PictPrint_Texto 3.5, 26, Rector
            PictPrint_Texto 12.5, 26, ULCase(Secretario2)
            PictPrint_Texto 3.5, 26.5, TextoRector
            PictPrint_Texto 12.5, 26.5, TextoSecretario2
     End Select
     End If
 End If
End With
RatonNormal
MensajeEncabData = ""
End Sub

Public Sub Aptitud_Promocion_Periodos(EsPromocion As Boolean, Curso As String, CodigoAlumno As String)
Dim PosLineaIni As Single
Dim AnchoPict As Single
Dim LogoAux As String
Dim Cursos As String
Dim CantPromConducta As Byte
Dim Disciplina As Single
'Dim Especialidad As String
RatonReloj
InicioX = 0.5: InicioY = 0.1: Abono = 0
'Pagina = 1
'Iniciamos la impresion

PictLibreta.Cls
PictLibreta.FontName = TipoTimes
PictLibreta.FontBold = False
'Notas del Alumnos
Disciplina = Procesar_Disciplinas(CodigoAlumno, Curso)
Notas_Del_Alumno Curso, CodigoAlumno
With AdoLibreta.Recordset
 If .RecordCount > 0 Then
     CodigoCliente = .Fields("Codigo")
     'MsgBox CodigoL
     If Mid$(CodigoL, 1, 4) < "1.02" Then
        SQLMsg1 = "AÑO LECTIVO: " & Anio_Lectivo
        SQLMsg2 = "I N F O R M E    F I N A L"
        SQLMsg3 = ""
        PictPrint_Grafico PictLibreta, RutaSistema & "\FORMATOS\INFORME.GIF", 0.5, 0.1, 19, 13
        PictEncabezado PictLibreta, 1, 0.5
        PictLibreta.FontBold = True
        PictLibreta.FontSize = 14
        PictPrint_Texto PictLibreta, 7.1, 4, .Fields("Alumno")
     Else
     Contra_Cta = Ninguno
     Cursos = "Estudiante del: "
     With AdoMatriculas.Recordset
      If .RecordCount > 0 Then
         .MoveFirst
         .Find ("Codigo = '" & CodigoCliente & "' ")
          If Not .EOF Then
             Contra_Cta = Dato_Curso.Especialidad
             CantPromConducta = 0
             If .Fields("ConductaPQ1") > 0 Then CantPromConducta = CantPromConducta + 1
             If .Fields("ConductaPQ2") > 0 Then CantPromConducta = CantPromConducta + 1
             If .Fields("ConductaSQ1") > 0 Then CantPromConducta = CantPromConducta + 1
             If .Fields("ConductaSQ2") > 0 Then CantPromConducta = CantPromConducta + 1
             If CantPromConducta <= 0 Then CantPromConducta = 1
             Abono = Redondear((.Fields("ConductaPQ1") + .Fields("ConductaPQ2") + .Fields("ConductaSQ1") + .Fields("ConductaSQ2")) / CantPromConducta)
             Cursos = Cursos & Leer_Datos_del_Curso(Curso, 1)
          End If
      End If
     End With
     SQLMsg2 = "AÑO LECTIVO: " & Anio_Lectivo
     If EsPromocion Then
        SQLMsg1 = "CERTIFICADO DE PROMOCIÓN"
     Else
        SQLMsg1 = "CERTIFICADO DE APTITUD"
     End If
     SQLMsg3 = "JORNADA MATUTINA"
     AnchoPict = Redondear(PictLibreta.width, 2)
     PictLibreta.FontName = TipoTimes
     PictLibreta.FontBold = True
     PictLibreta.FontSize = 19
     PosLinea = 0.5
     If Mid$(CodigoL, 1, 4) >= "2" Then
        LogoAux = RutaSistema & "\LOGOS\MINISEDU.GIF"
        PictPrint_Grafico PictLibreta, LogoAux, 1, 0.5, 4, 2
        'PictPrint_Grafico PictLibreta, LogoTipo, 16, 0.5, 4, 2
        PictLibreta.FontSize = 16
        PictPrint_Texto PictLibreta, 1, PosLinea, "REPÚBLICA DEL ECUADOR", , AnchoPict, True
        PosLinea = PosLinea + 0.7
        PictPrint_Texto PictLibreta, 1, PosLinea, "MINISTERIO DE EDUCACIÓN", , AnchoPict, True
        PosLinea = PosLinea + 1.4
        PictPrint_Texto PictLibreta, 1, PosLinea, Institucion1, , AnchoPict, True
        PosLinea = PosLinea + 0.7
        PictLibreta.FontSize = 18
        PictPrint_Texto PictLibreta, 1, PosLinea, Institucion2, , AnchoPict, True
        PosLinea = PosLinea + 1
     Else
        PosLinea = PosLinea + 4
     End If
     PictLibreta.FontSize = 10
     PictPrint_Texto PictLibreta, 1, PosLinea, UCase(NombreCiudad), , AnchoPict, True
     PosLinea = PosLinea + 0.8
     PictLibreta.FontBold = True
     If SQLMsg1 <> "" Then
        PictLibreta.FontSize = 14
        PictPrint_Texto PictLibreta, 1, PosLinea, SQLMsg1, , AnchoPict, True
        PosLinea = PosLinea + 0.6
     End If
     PictLibreta.FontSize = 16
     PictLibreta.FontBold = False
     If SQLMsg2 <> "" Then
        PictLibreta.FontBold = True
        PictPrint_Texto PictLibreta, 1, PosLinea, SQLMsg2, , AnchoPict, True
        PosLinea = PosLinea + 0.7
     End If
     PictLibreta.FontSize = 9
     If SQLMsg3 <> "" Then
        PictPrint_Texto PictLibreta, 2, PosLinea, SQLMsg3
        PosLinea = PosLinea + 0.5
     End If
     PosLinea = PosLinea + 0.1
     'PictEncabezado PictLibreta, 1, 0.5
     PictLibreta.FontUnderline = False
     PosLinea = PosLinea + 0.5
     PictLibreta.FontSize = 12
     If EsPromocion Then
        PictLibreta.FontBold = False
        If Mid$(Codigo4, 1, 1) < "2" Then
           Cadena1 = "La " & Institucion1 & " " & Institucion2 & ", " _
                   & "Conforme al Art. 315 del Reglamento General de la Ley de Educación, confiere el " _
                   & "presente CERTIFICADO DE PROMOCIÓN a la niña."
        Else
        Cadena1 = "La " & Institucion1 & " " & Institucion2 & ", " _
                & "Conforme al Art. 315 del Reglamento General de la Ley de Educación, confiere el " _
                & "presente CERTIFICADO DE PROMOCIÓN a la Srta."
        End If
        PosLinea = PictPrint_Texto_Multiple(PictLibreta, 2, PosLinea, Cadena1, 17)
        PosLinea = PosLinea + 0.8
        PictLibreta.FontSize = 16
        PictLibreta.FontBold = True
        PictPrint_Texto PictLibreta, 3.5, PosLinea, .Fields("Alumno")
        PictLibreta.FontBold = False
        PictLibreta.FontSize = 12
        PosLinea = PosLinea + 1
        PosLinea = PictPrint_Texto_Multiple(PictLibreta, 2, PosLinea, Cursos, 17)
        PosLinea = PosLinea + 0.6
        PictPrint_Texto PictLibreta, 2, PosLinea, "Luego de haberse presentado a exámenes, ha obtenido los siguientes promedios:"
        PosLinea = PosLinea + 0.6
        PosLineaIni = PosLinea
        PictLibreta.FontBold = True
        PictPrint_Texto PictLibreta, 3, PosLinea + 0.3, "A S I G N A T U R A S"
        PictLibreta.FontSize = 10
        PictPrint_Texto PictLibreta, 12, PosLinea + 0.05, "P R O M E D I O   D E   A P R O B A C I Ó N"
        PosLinea = PosLinea + 0.55
        PictPrint_Texto PictLibreta, 11, PosLinea, "NÚMEROS"
        PictPrint_Texto PictLibreta, 13.5, PosLinea, "LETRAS"
     Else
        PictPrint_Texto PictLibreta, 2, PosLinea, "CURSO: " & .Fields("Curso")
        PosLinea = PosLinea + 1
        PictPrint_Texto PictLibreta, 2, PosLinea, "NOMBRE DEL ESTUDIANTE: " & .Fields("Alumno")
        PosLinea = PosLinea + 0.4
        PictLibreta.FontSize = 12
        PictLibreta.FontBold = True
        PictPrint_Texto PictLibreta, 2, PosLinea, "ASIGNATURAS"
        PictLibreta.FontSize = 8
        PictPrint_Texto PictLibreta, 9, PosLinea, "I"
        PictPrint_Texto PictLibreta, 10.4, PosLinea, "II"
        PictPrint_Texto PictLibreta, 11.4, PosLinea, "PUNTAJE"
        PictPrint_Texto PictLibreta, 12.2, PosLinea, "PROMEDIO"
        PictPrint_Texto PictLibreta, 15, PosLinea, "SUPLE_"
        PictPrint_Texto PictLibreta, 17.5, PosLinea, "PROMEDIO"
        PosLinea = PosLinea + 0.35
        PictPrint_Texto PictLibreta, 8.7, PosLinea, "QUIM"
        PictPrint_Texto PictLibreta, 10.1, PosLinea, "QUIM"
        PictPrint_Texto PictLibreta, 11.4, PosLinea, "TOTAL"
        PictPrint_Texto PictLibreta, 13.2, PosLinea, "QUIMEST."
        PictPrint_Texto PictLibreta, 15, PosLinea, "TORIO"
        PictPrint_Texto PictLibreta, 17.5, PosLinea, "FINAL"
     End If
     PictLibreta.FontBold = False
     PosLinea = PosLinea + 0.5
     
     PictLibreta.Line (1.9, PosLinea)-(19.5, PosLinea), QBColor(0)
     PosLinea = PosLinea + 0.05
     Saldo = 0: Contador = 0: Debe = 0
     PictLibreta.FontUnderline = False
     PictLibreta.FontBold = False
     PictLibreta.FontSize = 12
     Do While Not .EOF
        'MsgBox "Orden: " & .Fields("Orden") & " - " & .Fields("Materia") & " - C: " & .Fields("C") & " - P: " & .Fields("P") & " - I: " & .Fields("I")
        If .Fields("CodMatP") = Ninguno Then
        If .Fields("I") Then
           'Disciplina
           'If .Fields("Orden") = 9 Then                Abono = Redondear(.Fields("PromFinal"))
            Abono_ME = Redondear(.Fields("PromFinal"))
            Total = .Fields("PromPQ") + .Fields("PromSQ")
           'Sumatoria .Fields("P") And
            If .Fields("Orden") <> 9 Then
                If Not .Fields("C") Then
                    Saldo = Saldo + Redondear(.Fields("PromFinal"))
                    Contador = Contador + 1
                End If
                PictPrint_Texto PictLibreta, 2, PosLinea, UCase(.Fields("Materia"))
                If EsPromocion Then
                   PictPrint_Nota_Materia PictLibreta, 11.6, PosLinea, .Fields("PromFinal"), .Fields("C")
                   If .Fields("C") = False Then PictPrint_Texto PictLibreta, 13.5, PosLinea, Cambio_Letras(Abono_ME)
                Else
                   PictPrint_Texto PictLibreta, 9, PosLinea, Format(.Fields("PromPQ"), "##")
                   PictPrint_Texto PictLibreta, 10.5, PosLinea, Format(.Fields("PromSQ"), "##")
                   PictPrint_Texto PictLibreta, 12, PosLinea, Format(Total, "##")
                   PictPrint_Texto PictLibreta, 13.5, PosLinea, Format(Abono_ME, "##")
                   PictPrint_Texto PictLibreta, 15.5, PosLinea, Format(.Fields("Supletorio"), "##")
                   PictPrint_Texto PictLibreta, 17.5, PosLinea, Format(.Fields("PromFinal"), "##")
                End If
                PosLinea = PosLinea + 0.5
               ' MsgBox Redondear(.Fields("PromFinal"), Dec_Nota)
                Debe = Debe + Redondear(.Fields("PromFinal"), Dec_Nota)
            End If
        End If
        End If
       .MoveNext
     Loop
     PictLibreta.FontBold = True
     PosLinea = PosLinea + 0.1
     If EsPromocion Then
        If Contador = 0 Then Contador = 1
        'PictLibreta.FontSize = 12
        PictLibreta.FontBold = True
        PictLibreta.Line (1.9, PosLinea)-(19.5, PosLinea), QBColor(0)
        PosLinea = PosLinea + 0.1
        PictPrint_Texto PictLibreta, 2, PosLinea, "T O T A L"
        PictLibreta.FontBold = False
        If Dec_Nota > 0 Then
           PictPrint_Texto PictLibreta, 11.3, PosLinea, Format(Debe, "00.00")
        Else
           PictPrint_Texto PictLibreta, 11.3, PosLinea, Format(Debe, "00")
        End If
        PosLinea = PosLinea + 0.6
        PictLibreta.Line (1.9, PosLinea)-(19.5, PosLinea), QBColor(0)
        PosLinea = PosLinea + 0.1
        PictLibreta.FontBold = True
        PictPrint_Texto PictLibreta, 2, PosLinea, "PROMEDIO DE RENDIMIENTO"
        PictLibreta.FontBold = False
        PictPrint_Texto PictLibreta, 11.3, PosLinea, Format(Saldo / Contador, "00.00")
        PictLibreta.FontSize = 11
        PictPrint_Texto PictLibreta, 13.5, PosLinea, Cambio_Letras_Decimales(Saldo / Contador, 2)
        PictLibreta.FontSize = 12
        PosLinea = PosLinea + 0.6
        PictLibreta.Line (1.9, PosLinea)-(19.5, PosLinea), QBColor(0)
        PosLinea = PosLinea + 0.1
        PictLibreta.FontBold = True
        PictPrint_Texto PictLibreta, 2, PosLinea, "PROMEDIO DE CONDUCTA"
        CantPromConducta = 0
        If Real1 > 0 Then CantPromConducta = CantPromConducta + 1
        If Real2 > 0 Then CantPromConducta = CantPromConducta + 1
        If Real3 > 0 Then CantPromConducta = CantPromConducta + 1
        If Real4 > 0 Then CantPromConducta = CantPromConducta + 1
        If CantPromConducta <= 0 Then CantPromConducta = 1
        Abono = Redondear((Real1 + Real2 + Real3 + Real4) / CantPromConducta)
        PictLibreta.FontBold = False
        PictPrint_Texto PictLibreta, 11.3, PosLinea, Format(Disciplina, "00.00")
        'PictPrint_Texto PictLibreta, 11.7, PosLinea, Cambio_Letras(Abono, True)
        PosLinea = PosLinea + 0.6
        PictLibreta.Line (1.9, PosLinea)-(19.5, PosLinea), QBColor(0)
     Else
        If Contador = 0 Then Contador = 1
        PictPrint_Texto PictLibreta, 2, PosLinea, "PUNTAJE TOTAL:   " & Format(Saldo, "##.##")
        PosLinea = PosLinea + 0.8
        PictPrint_Texto PictLibreta, 2, PosLinea, "Promedio de Aprovechamiento:   " & Format(Saldo / Contador, "##.##")
        PosLinea = PosLinea + 0.8
        PictPrint_Texto PictLibreta, 2, PosLinea, "Promedio de Disciplina:   " & Format(Disciplina, "##.##")
        PosLinea = PosLinea + 0.8
     End If
     PictLibreta.Line (1.9, PosLineaIni)-(19.5, PosLineaIni), QBColor(0)
     PictLibreta.Line (10.9, PosLineaIni + 0.5)-(19.5, PosLineaIni + 0.5), QBColor(0)
     PictLibreta.Line (1.9, PosLineaIni)-(1.9, PosLinea), QBColor(0)
     PictLibreta.Line (10.9, PosLineaIni)-(10.9, PosLinea), QBColor(0)
     PictLibreta.Line (12.8, PosLineaIni + 0.55)-(12.8, PosLinea), QBColor(0)
     PictLibreta.Line (19.5, PosLineaIni)-(19.5, PosLinea), QBColor(0)
     PosLinea = PosLinea + 0.2
     If Mid$(Codigo4, 1, 1) < "2" Then
        PictPrint_Texto PictLibreta, 2, PosLinea, "Por lo tanto se lo promueve al " & String(58, "_") & "."
     Else
        PictPrint_Texto PictLibreta, 2, PosLinea, "Por lo tanto se la promueve al curso inmediato superior."
     End If
     PosLinea = PosLinea + 0.6
     PictPrint_Texto PictLibreta, 2, PosLinea, "Así consta en los libros de calificaciones de la Secretaria de esta Unidad Educativa."
     NoMeses = FechaMes(MBFecha)
     NoAnio = FechaAnio(MBFecha)
     PosLinea = PosLinea + 1
     PictPrint_Texto PictLibreta, 13, PosLinea, FechaStrgCiudad(MBFecha)
     PosLinea = PosLinea + 0.55
     Select Case Codigo4
       Case "0.00" To "1.99"
            PictPrint_Texto PictLibreta, 3.5, 25, String(Len(Rector) - 2, "_")
            PictPrint_Texto PictLibreta, 11, 25, String(Len(Secretario2) - 2, "_")
            PictPrint_Texto PictLibreta, 4, 25.5, Director
            PictPrint_Texto PictLibreta, 11.5, 25.5, Secretario1
            PictPrint_Texto PictLibreta, 4, 26, TextoDirector
            PictPrint_Texto PictLibreta, 11.5, 26, TextoSecretario1
       Case "2.00" To "5.99"
            PictPrint_Texto PictLibreta, 3.5, 25, String(Len(Rector) - 2, "_")
            PictPrint_Texto PictLibreta, 11, 25, String(Len(Secretario2) - 2, "_")
            PictPrint_Texto PictLibreta, 4, 25.5, Rector
            PictPrint_Texto PictLibreta, 11.5, 25.5, Secretario2
            PictPrint_Texto PictLibreta, 4, 26, TextoRector
            PictPrint_Texto PictLibreta, 11.5, 26, TextoSecretario2
     End Select
     End If
 End If
End With
RatonNormal
MensajeEncabData = ""
End Sub

Public Sub Certificado_Matricula()
Dim UltimaLinea As Single
RatonReloj
Notas_Del_Alumno CodigoL, CodigoCliente
InicioX = 0.5: InicioY = 0.1
'Pagina = 1
'Iniciamos la impresion
PictLibreta.Cls
PictLibreta.width = AnchoMaximo
PictLibreta.Height = AltoMaximo
PictLibreta.FontName = TipoTimes
PictLibreta.FontBold = False
With AdoAlumnos.Recordset
 If .RecordCount > 0 Then
     'CodigoCliente = .Fields("Codigo")
     'MsgBox CodigoL
     SQLMsg1 = "AÑO LECTIVO: " & Anio_Lectivo
     SQLMsg2 = "CERTIFICADO DE MATRICULA"
     SQLMsg3 = ""
     PictLibreta.FontBold = True
     PictLibreta.FontSize = 18
     PictPrint_Texto PictLibreta, 1.5, 0.5, "REPÚBLICA DEL ECUADOR", , 20, True
     PictEncabezado PictLibreta, 1, 1.3
     PictLibreta.FontUnderline = False
     PosLinea = PosLinea + 0.5
     PictLibreta.FontSize = 12
    .MoveFirst
    .Find ("Codigo = '" & CodigoCliente & "' ")
     If Not .EOF Then
        Cadena = Leer_Datos_del_Curso(.Fields("Grupo_No"))
         PictPrint_Texto PictLibreta, 1.5, PosLinea, "La " & Empresa & " declara que el(la) Alumno(a):"
         PosLinea = PosLinea + 1
         PictLibreta.FontSize = 16
         PictLibreta.FontBold = True
         PictPrint_Texto PictLibreta, 1.5, PosLinea, .Fields("Alumno"), , 20, True
         PictLibreta.FontBold = False
         PosLinea = PosLinea + 1
         PictLibreta.FontSize = 12
         PictPrint_Texto PictLibreta, 1.5, PosLinea, "Previo cumplimiento de los correspondientes requisitos legales reglamentarios, ha sido matriculado(a) en esta"
         PosLinea = PosLinea + 0.6
         PictPrint_Texto PictLibreta, 1.5, PosLinea, "Unidad Educativa con los siguientes datos:"
         PosLinea = PosLinea + 1
         UltimaLinea = PosLinea
         PictPrint_Texto PictLibreta, 4, PosLinea, "CURSO"
         PictPrint_Texto PictLibreta, 8.8, PosLinea, ":"
         PosLinea = PosLinea + 0.7
         PictPrint_Texto PictLibreta, 4, PosLinea, "ESPECIALIDAD"
         PictPrint_Texto PictLibreta, 8.8, PosLinea, ":"
         PosLinea = PosLinea + 0.7
    ''     PictPrint_Texto PictLibreta, 4, PosLinea, "CICLO"
    ''     PictPrint_Texto PictLibreta, 8.8, PosLinea, ":"
    ''     PosLinea = PosLinea + 0.7
         PictPrint_Texto PictLibreta, 4, PosLinea, "NIVEL"
         PictPrint_Texto PictLibreta, 8.8, PosLinea, ":"
         PosLinea = PosLinea + 0.7
         PictPrint_Texto PictLibreta, 4, PosLinea, "SECCION"
         PictPrint_Texto PictLibreta, 8.8, PosLinea, ":"
         PosLinea = PosLinea + 0.7
         PictPrint_Texto PictLibreta, 4, PosLinea, "MATRICULA No."
         PictPrint_Texto PictLibreta, 8.8, PosLinea, ":"
         PosLinea = PosLinea + 0.7
         PictPrint_Texto PictLibreta, 4, PosLinea, "FOLIO No."
         PictPrint_Texto PictLibreta, 8.8, PosLinea, ":"
         PosLinea = PosLinea + 0.7
         PictPrint_Texto PictLibreta, 4, PosLinea, "FECHA DE MATRICULA"
         PictPrint_Texto PictLibreta, 8.8, PosLinea, ":"
         PosLinea = UltimaLinea
         PictLibreta.FontBold = True
         PictPrint_Texto PictLibreta, 9.3, PosLinea, Dato_Curso.Descripcion
         PosLinea = PosLinea + 0.7
         PictPrint_Texto PictLibreta, 9.3, PosLinea, Dato_Curso.Especialidad
         PosLinea = PosLinea + 0.7
    ''     PictPrint_Texto PictLibreta, 1, PosLinea, .Fields("")
    ''     PosLinea = PosLinea + 0.7
         Select Case Codigo4
           Case "0.00" To "1.99": Cadena = "PRIMARIA"
           Case "2.00" To "5.99": Cadena = "SECUNDARIA"
         End Select
         PictPrint_Texto PictLibreta, 9.3, PosLinea, Cadena
         PosLinea = PosLinea + 0.7
         PictPrint_Texto PictLibreta, 9.3, PosLinea, Dato_Curso.Seccion
         PosLinea = PosLinea + 0.7
         PictPrint_Texto PictLibreta, 9.3, PosLinea, Format(.Fields("Matricula_No"), "000000")
         PosLinea = PosLinea + 0.7
         PictPrint_Texto PictLibreta, 9.3, PosLinea, Format(.Fields("Folio_No"), "000000")
         PosLinea = PosLinea + 0.7
         PictPrint_Texto PictLibreta, 9.3, PosLinea, .Fields("Fecha_M")
         PosLinea = PosLinea + 1
         PictLibreta.FontBold = False
         PictLibreta.FontUnderline = False
         PictPrint_Texto PictLibreta, 1.5, PosLinea, "Para constancia y fines consiguientes, se le confiere el presente CERTIFICADO en la Ciudad de:"
         PosLinea = PosLinea + 0.7
         PictPrint_Texto PictLibreta, 1.5, PosLinea, FechaStrgCiudad(MBFecha)
         PosLinea = PosLinea + 1.2
         PictLibreta.FontBold = True
         PictPrint_Texto PictLibreta, 1.5, PosLinea, "OBSERVACIONES:"
         PictLibreta.FontBold = False
         PictPrint_Texto PictLibreta, 5.5, PosLinea, .Fields("Observaciones")
         PosLinea = PosLinea + 0.6
         Select Case Codigo4
           Case "0.00" To "1.99"
                PictPrint_Texto PictLibreta, 5, 20, Director
                PictPrint_Texto PictLibreta, 13.5, 20, Secretario1
                PictPrint_Texto PictLibreta, 5, 20.5, TextoDirector
                PictPrint_Texto PictLibreta, 13.5, 20.5, TextoSecretario1
           Case "2.00" To "3.99"
                PictPrint_Texto PictLibreta, 5, 20, Rector
                PictPrint_Texto PictLibreta, 13.5, 20, Secretario2
                PictPrint_Texto PictLibreta, 5, 20.5, TextoRector
                PictPrint_Texto PictLibreta, 13.5, 20.5, TextoSecretario2
          Case "4.00" To "5.99"
                PictPrint_Texto PictLibreta, 5, 20, Rector
                PictPrint_Texto PictLibreta, 13.5, 20, Secretario3
                PictPrint_Texto PictLibreta, 5, 20.5, TextoRector
                PictPrint_Texto PictLibreta, 13.5, 20.5, TextoSecretario3
         End Select
     End If
 End If
End With
RatonNormal
MensajeEncabData = ""
End Sub

Public Sub Alumnos_Matriculados(ElCurso As String)
Dim Y0 As Single
Dim y1 As Single
Dim X0 As Single
Dim x1 As Single
Dim Curso As String
Dim PosLineaX As Single
Dim TotalRegs(11) As Integer
Dim NumFileAlumnos As Long
On Error GoTo Errorhandler
With AdoAlumnos.Recordset
 If .RecordCount > 0 Then
    .MoveFirst
   RatonReloj
   MensajeEncabData = "LISTADO DE ALUMNOS MATRICULADOS POR CURSO"
   Mensajes = "Seguro de Imprimir en:" & vbCrLf & Printer.DeviceName & "?"
   Titulo = "IMPRESION"
   Bandera = False
   'MsgBox CodigoAlumno
   SetPrinters.Show 1
   If PonImpresoraDefecto(SetNombrePRN) Then
      InicioX = 0.5: InicioY = 0
      DataAnchoCampos InicioX, AdoAlumnos, 8, TipoTimes, 1, True
      PosLinea = 3.6
      Contador = 0
     .MoveFirst
      Codigo1 = .Fields("Alumno")
      CodigoCli = .Fields("Codigo")
      CICliente = .Fields("Matricula_No")
      Cadena = Leer_Datos_del_Curso(ElCurso)
      SQLMsg1 = "AÑO LECTIVO " & Anio_Lectivo
      SQLMsg2 = "Listado de Alumnos Matriculados por Curso"
      SQLMsg3 = Dato_Curso.Curso & " - " & Dato_Curso.Descripcion
      SQLMsg4 = "Fecha: " & FechaStrg(MBFecha)
      Curso = Dato_Curso.Curso
      PrinterFontBold True
      Encabezado_Documento 2, 0.5, 19
      PrinterFontBold True
      Y0 = PosLinea
      PrinterTexto 2, PosLinea, "NUM."
      PrinterTexto 3, PosLinea, "NOMBRE DEL ALUMNO"
      PrinterTexto 14.5, PosLinea, "FOLIO"
      PrinterTexto 17, PosLinea, "MATRICULA"
      PosLinea = PosLinea + 0.35
      Imprimir_Linea_H PosLinea, 2, 19, Negro
      PosLinea = PosLinea + 0.1
      PrinterFontBold False
      PrinterFontSize 8
      Do While Not .EOF
         Contador = Contador + 1
         PrinterTexto 2, PosLinea, Format(Contador, "00")
         PrinterTexto 3, PosLinea, .Fields("Alumno")
         PrinterTexto 14.5, PosLinea, .Fields("Folio_No")
         PrinterTexto 17, PosLinea, .Fields("Matricula_No")
         PosLinea = PosLinea + 0.35
         Imprimir_Linea_H PosLinea, 2, 19, Negro
         PosLinea = PosLinea + 0.05
         If PosLinea >= LimiteAlto Then
            Imprimir_Linea_H PosLinea, 2, 19, Negro, True
            Imprimir_Linea_V 1.9, Y0, PosLinea, Negro
            Imprimir_Linea_V 2.9, Y0, PosLinea, Negro
            Imprimir_Linea_V 14.4, Y0, PosLinea, Negro
            Imprimir_Linea_V 16.9, Y0, PosLinea, Negro
            Imprimir_Linea_V 19, Y0, PosLinea, Negro
            Printer.NewPage
            PosLinea = 3.6
            PrinterFontBold True
            Encabezado_Documento 2, 0.5, 19
            PrinterFontBold True
            Y0 = PosLinea
            PrinterTexto 2, PosLinea, "NUM."
            PrinterTexto 3, PosLinea, "NOMBRE DEL ALUMNO"
            PrinterTexto 14.5, PosLinea, "FOLIO"
            PrinterTexto 17, PosLinea, "MATRICULA"
            PosLinea = PosLinea + 0.35
            Imprimir_Linea_H PosLinea, 2, 19, Negro
            PosLinea = PosLinea + 0.1
            PrinterFontBold False
            PrinterFontSize 8
         End If
        .MoveNext
      Loop
      Imprimir_Linea_V 1.9, Y0, PosLinea, Negro
      Imprimir_Linea_V 2.9, Y0, PosLinea, Negro
      Imprimir_Linea_V 14.4, Y0, PosLinea, Negro
      Imprimir_Linea_V 16.9, Y0, PosLinea, Negro
      Imprimir_Linea_V 19, Y0, PosLinea, Negro
      Cuadricula = False
      MensajeEncabData = "": SQLMsg1 = "": SQLMsg2 = "": SQLMsg3 = "": SQLMsg4 = ""
      PrinterFontSize 6
      Cadena = ""
      For I = 1 To 5
          Cadena = Cadena & String(35, "-") & " x "
      Next I
      Cadena = Cadena & String(35, "-")
      PrinterTexto CentrarTexto(Cadena), PosLinea, Cadena
      PrinterFontSize 8
      PosLinea = LimiteAlto
      Select Case Curso
        Case "0.00" To "1.99"
             PrinterTexto 3, PosLinea, Director
             PrinterTexto 10, PosLinea, Secretario1
             PosLinea = PosLinea + 0.4
             PrinterTexto 3, PosLinea, TextoDirector
             PrinterTexto 10, PosLinea, TextoSecretario1
        Case "2.00" To "3.99"
             PrinterTexto 3, PosLinea, Rector
             PrinterTexto 10, PosLinea, Secretario2
             PosLinea = PosLinea + 0.4
             PrinterTexto 3, PosLinea, TextoRector
             PrinterTexto 10, PosLinea, TextoSecretario2
      End Select
      PrinterEndDoc
   End If
  'Comenzamos a generar el archivo del curso a csv
   RatonNormal
   Mensajes = "Exportar Lista de Matriculados a Excel"
   Titulo = "Pregunta de Exportación"
   If BoxMensaje = vbYes Then
      RatonReloj
      Contador = 0
      If .RecordCount > 0 Then
         .MoveFirst
          Codigo1 = .Fields("Alumno")
          CodigoCli = .Fields("Codigo")
          CICliente = .Fields("Matricula_No")
          Cadena = Leer_Datos_del_Curso(ElCurso)
          SQLMsg1 = "AÑO LECTIVO " & Anio_Lectivo
          SQLMsg2 = "Listado de Alumnos Matriculados por Curso"
          SQLMsg3 = Dato_Curso.Curso & " - " & Dato_Curso.Descripcion
          SQLMsg4 = "Fecha: " & FechaStrg(MBFecha)
          Curso = Dato_Curso.Curso
          
          NombreArchivo = RutaSysBases & "\Excel\" & Trim(Replace(Dato_Curso.Descripcion, Chr(34), "")) & ".csv"
          NumFileAlumnos = FreeFile
          Open NombreArchivo For Output As #NumFileAlumnos
          Print #NumFileAlumnos, Empresa
          Print #NumFileAlumnos, ""
          Print #NumFileAlumnos, SQLMsg1
          Print #NumFileAlumnos, SQLMsg2
          Print #NumFileAlumnos, SQLMsg3
          Print #NumFileAlumnos, SQLMsg4
          Print #NumFileAlumnos, "NUM.;";
          Print #NumFileAlumnos, "NOMBRE DEL ALUMNO;";
          Print #NumFileAlumnos, "FOLIO;";
          Print #NumFileAlumnos, "MATRICULA;"
          Do While Not .EOF
              Contador = Contador + 1
              Print #NumFileAlumnos, Format(Contador, "00") & ";";
              Print #NumFileAlumnos, .Fields("Alumno") & ";";
              Print #NumFileAlumnos, .Fields("Folio_No") & ";";
              Print #NumFileAlumnos, .Fields("Matricula_No") & ";"
             .MoveNext
          Loop
          Select Case Curso
             Case "0.00" To "1.99"
                  Print #NumFileAlumnos, ";" & Director & ";;";
                  Print #NumFileAlumnos, Secretario1 & ";"
                  
                  Print #NumFileAlumnos, ";" & TextoDirector & ";;";
                  Print #NumFileAlumnos, TextoSecretario1 & ";"
             Case "2.00" To "3.99"
                  Print #NumFileAlumnos, ";" & Rector & ";;";
                  Print #NumFileAlumnos, Secretario2 & ";"
                  
                  Print #NumFileAlumnos, ";" & TextoRector & ";;";
                  Print #NumFileAlumnos, TextoSecretario2 & ";"
          End Select
          Close #NumFileAlumnos
          RatonNormal
      End If
   End If
      Exit Sub
Errorhandler:
    RatonNormal
    ErrorDeImpresion
    Exit Sub
Else
    RatonNormal
End If

  
End With
End Sub

Public Sub Alumnos_Matriculados_Seccion(ElCurso As String)
Dim Y0 As Single
Dim Yo As Single
Dim y1 As Single
Dim X0 As Single
Dim x1 As Single
Dim Curso As String
Dim PosLineaX As Single
Dim TotalRegs(11) As Integer
On Error GoTo Errorhandler
sSQL = "SELECT CC.Descripcion, CC.Curso, C.Cliente As Alumno, C.Sexo, C.Celular, C.Telefono, CM.* " _
       & "FROM Catalogo_Cursos As CC,Clientes As C,Clientes_Matriculas As CM " _
       & "WHERE CC.Item = '" & NumEmpresa & "' " _
       & "AND CC.Periodo = '" & Periodo_Contable & "' " _
       & "AND Mid$(CC.Curso,1,1) >= '" & Mid$(ElCurso, 1, 1) & "' " _
       & "AND CC.Item = CM.Item " _
       & "AND CC.Periodo = CM.Periodo " _
       & "AND CC.Curso = CM.Grupo_No " _
       & "AND C.Codigo = CM.Codigo " _
       & "ORDER BY CC.Curso,C.Cliente,C.Sexo "
SelectAdodc AdoAux, sSQL
With AdoAux.Recordset
 If .RecordCount > 0 Then
    .MoveFirst
   RatonReloj
   MensajeEncabData = "LISTADO DE ALUMNOS MATRICULADOS"
   Mensajes = "Seguro de Imprimir en:" & vbCrLf & Printer.DeviceName & "?"
   Titulo = "IMPRESION"
   Bandera = False
   'MsgBox CodigoAlumno
   SetPrinters.Show 1
   If PonImpresoraDefecto(SetNombrePRN) Then
      InicioX = 0.5: InicioY = 0
      DataAnchoCampos InicioX, AdoAlumnos, 8, TipoTimes, 1, True
      Pagina = 1
      PosLinea = 3.6
      Contador = 0
     .MoveFirst
      Codigo1 = .Fields("Alumno")
      CodigoCli = .Fields("Codigo")
      CICliente = .Fields("Matricula_No")
      SQLMsg1 = "AÑO LECTIVO " & Anio_Lectivo
      SQLMsg2 = ""
      SQLMsg3 = ""
      SQLMsg4 = ""
      Codigo2 = Ninguno
      Codigo3 = Ninguno
      Curso = .Fields("Curso")
      Codigo3 = Leer_Datos_del_Curso(.Fields("Curso"), 1)
      Codigo2 = Dato_Curso.Especialidad
      PrinterFontBold True
      Encabezado_Institucion 2, 19
      PrinterFontBold True
      NumeroLineas = PrinterLineasMayor(2, PosLinea, Codigo3, 12)
      PrinterTexto 15.5, PosLinea, "JORNADA MATUTINA"
      PosLinea = PosLinea + (NumeroLineas * 0.4)
      Imprimir_Linea_H PosLinea, 2, 19, Negro
      PosLinea = PosLinea + 0.1
      Yo = PosLinea
      Y0 = PosLinea
      PrinterTexto 2, PosLinea, "NUM."
      PrinterTexto 3, PosLinea, "NOMBRE DEL ALUMNO"
      PrinterTexto 10.8, PosLinea, "FOLIO"
      PrinterTexto 12.6, PosLinea, "MATRICULA"
      PosLinea = PosLinea + 0.35
      Imprimir_Linea_H PosLinea, 2, 19, Negro
      PosLinea = PosLinea + 0.1
      PrinterFontBold False
      PrinterFontSize 8
      Do While Not .EOF
         If Curso <> .Fields("Curso") Then
            Curso = .Fields("Curso")
            PosLinea = PosLinea - 0.05
            Imprimir_Linea_H PosLinea, 2, 14.6, Negro
            Imprimir_Linea_V 1.9, Y0, PosLinea, Negro
            Imprimir_Linea_V 2.9, Y0, PosLinea, Negro
            Imprimir_Linea_V 10.5, Y0, PosLinea, Negro
            Imprimir_Linea_V 12.5, Y0, PosLinea, Negro
            Imprimir_Linea_V 14.6, Y0, PosLinea, Negro
            PosLinea = PosLinea + 0.2
            Printer.FontBold = True
            SQLMsg2 = ""
            SQLMsg3 = ""
            SQLMsg4 = ""
            Codigo2 = Ninguno
            Codigo3 = Ninguno
            Codigo3 = Leer_Datos_del_Curso(.Fields("Curso"), 1)
            Codigo2 = Dato_Curso.Especialidad
            NumeroLineas = PrinterLineasMayor(2, PosLinea, Codigo3, 12)
            PosLinea = PosLinea + (NumeroLineas * 0.4)
            Imprimir_Linea_H PosLinea, 2, 14.6, Negro
            PosLinea = PosLinea + 0.1
            Y0 = PosLinea
            If PosLinea >= LimiteAlto Then
                PosLinea = PosLinea - 0.05
                Imprimir_Linea_H PosLinea, 2, 19, Negro
                Imprimir_Linea_V 1.9, Y0, PosLinea, Negro
                Imprimir_Linea_V 2.9, Y0, PosLinea, Negro
                Imprimir_Linea_V 10.5, Y0, PosLinea, Negro
                Imprimir_Linea_V 12.5, Y0, PosLinea, Negro
                Imprimir_Linea_V 14.6, Yo, PosLinea, Negro
                Imprimir_Linea_V 19, Yo, PosLinea, Negro
                PosLineaX = Redondear((PosLinea - Yo) / 3)
                PosLinea = PosLineaX + Yo
                Imprimir_Linea_H PosLinea - 0.7, 15, 18.5, Negro
                PrinterTexto 15, PosLinea - 0.5, "Supervisor de Educación Media"
                Imprimir_Linea_H PosLinea, 14.6, 19, Negro
                PosLinea = PosLinea + PosLineaX
                Imprimir_Linea_H PosLinea - 0.7, 15, 18.5, Negro
                Select Case Curso
                  Case "0.00" To "1.99"
                       PrinterTexto 15.5, PosLinea - 0.5, TextoDirector
                  Case "2.00" To "3.99"
                       PrinterTexto 15.5, PosLinea - 0.5, TextoRector
                End Select
                Imprimir_Linea_H PosLinea, 14.6, 19, Negro
                PosLinea = PosLinea + PosLineaX
                Imprimir_Linea_H PosLinea - 0.9, 15, 18.5
                Select Case Curso
                  Case "0.00" To "1.99"
                       PrinterTexto 15.5, PosLinea - 0.7, TextoSecretario1
                  Case "2.00" To "3.99"
                       PrinterTexto 15.5, PosLinea - 0.7, TextoSecretario2
                End Select
                Printer.NewPage
                PosLinea = 3.6
                PrinterFontBold True
                Encabezado_Institucion 2, 19
                PrinterFontBold True
                NumeroLineas = PrinterLineasMayor(2, PosLinea, Codigo3, 12)
                PrinterTexto 15.5, PosLinea, "JORNADA MATUTINA"
                PosLinea = PosLinea + (NumeroLineas * 0.4)
                Imprimir_Linea_H PosLinea, 2, 19, Negro
                PosLinea = PosLinea + 0.1
                Y0 = PosLinea
                Yo = PosLinea
                PrinterTexto 2, PosLinea, "NUM."
                PrinterTexto 3, PosLinea, "NOMBRE DEL ALUMNO"
                PrinterTexto 10.8, PosLinea, "FOLIO"
                PrinterTexto 12.6, PosLinea, "MATRICULA"
                PosLinea = PosLinea + 0.35
                Imprimir_Linea_H PosLinea, 2, 19, Negro
                PosLinea = PosLinea + 0.1
                PrinterFontBold False
                PrinterFontSize 8
            End If
         End If
         Contador = Contador + 1
         Printer.FontBold = False
         PrinterTexto 2, PosLinea, Format(Contador, "000")
         PrinterTexto 3.2, PosLinea, .Fields("Alumno")
         PrinterTexto 11, PosLinea, .Fields("Folio_No")
         PrinterTexto 13, PosLinea, .Fields("Matricula_No")
         PosLinea = PosLinea + 0.35
         Imprimir_Linea_H PosLinea, 2, 14.6, Negro
         PosLinea = PosLinea + 0.05
         If PosLinea >= LimiteAlto Then
            PosLinea = PosLinea - 0.05
            Imprimir_Linea_H PosLinea, 2, 19, Negro
            Imprimir_Linea_V 1.9, Y0, PosLinea, Negro
            Imprimir_Linea_V 2.9, Y0, PosLinea, Negro
            Imprimir_Linea_V 10.5, Y0, PosLinea, Negro
            Imprimir_Linea_V 12.5, Y0, PosLinea, Negro
            Imprimir_Linea_V 14.6, Yo, PosLinea, Negro
            Imprimir_Linea_V 19, Yo, PosLinea, Negro
            PosLineaX = Redondear((PosLinea - Yo) / 3)
            PosLinea = PosLineaX + Yo
            Imprimir_Linea_H PosLinea - 0.7, 15, 18.5, Negro
            PrinterTexto 15, PosLinea - 0.5, "Supervisor de Educación Media"
            Imprimir_Linea_H PosLinea, 14.6, 19, Negro
            PosLinea = PosLinea + PosLineaX
            Imprimir_Linea_H PosLinea - 0.7, 15, 18.5, Negro
            Select Case Curso
              Case "0.00" To "1.99"
                   PrinterTexto 15.5, PosLinea - 0.5, TextoDirector
              Case "2.00" To "3.99"
                   PrinterTexto 15.5, PosLinea - 0.5, TextoRector
            End Select
            Imprimir_Linea_H PosLinea, 14.6, 19, Negro
            PosLinea = PosLinea + PosLineaX
            Imprimir_Linea_H PosLinea - 0.9, 15, 18.5
            Select Case Curso
              Case "0.00" To "1.99"
                   PrinterTexto 15.5, PosLinea - 0.7, TextoSecretario1
              Case "2.00" To "3.99"
                   PrinterTexto 15.5, PosLinea - 0.7, TextoSecretario2
            End Select
            Printer.NewPage
            PosLinea = 3.6
            PrinterFontBold True
            Encabezado_Institucion 2, 19
            PrinterFontBold True
            NumeroLineas = PrinterLineasMayor(2, PosLinea, Codigo3, 12)
            PrinterTexto 15.5, PosLinea, "JORNADA MATUTINA"
            PosLinea = PosLinea + (NumeroLineas * 0.4)
            Imprimir_Linea_H PosLinea, 2, 19, Negro
            PosLinea = PosLinea + 0.1
            Y0 = PosLinea
            Yo = PosLinea
            PrinterTexto 2, PosLinea, "NUM."
            PrinterTexto 3, PosLinea, "NOMBRE DEL ALUMNO"
            PrinterTexto 10.8, PosLinea, "FOLIO"
            PrinterTexto 12.6, PosLinea, "MATRICULA"
            PosLinea = PosLinea + 0.35
            Imprimir_Linea_H PosLinea, 2, 19, Negro
            PosLinea = PosLinea + 0.1
            PrinterFontBold False
            PrinterFontSize 8
         End If
        .MoveNext
      Loop
      PosLinea = PosLinea - 0.05
      Imprimir_Linea_H PosLinea, 2, 14.6, Negro
      Imprimir_Linea_V 1.9, Y0, PosLinea, Negro
      Imprimir_Linea_V 2.9, Y0, PosLinea, Negro
      Imprimir_Linea_V 10.5, Y0, PosLinea, Negro
      Imprimir_Linea_V 12.5, Y0, PosLinea, Negro
      PosLinea = PosLinea + 0.1
      PrinterTexto 2, PosLinea, FechaStrgCiudad(FechaComp)
      PosLinea = LimiteAlto
      Imprimir_Linea_H PosLinea, 14.6, 19, Negro, True
      Imprimir_Linea_V 14.6, Yo, PosLinea, Negro
      Imprimir_Linea_V 19, Yo, PosLinea, Negro
      PosLineaX = Redondear((PosLinea - Yo) / 3)
      PosLinea = PosLineaX + Yo
      Imprimir_Linea_H PosLinea - 0.7, 15, 18.5, Negro
      PrinterTexto 15, PosLinea - 0.5, "Supervisor de Educación Media"
      Imprimir_Linea_H PosLinea, 14.6, 19, Negro
      PosLinea = PosLinea + PosLineaX
      Imprimir_Linea_H PosLinea - 0.7, 15, 18.5, Negro
      Select Case Curso
        Case "0.00" To "1.99"
             PrinterTexto 15.5, PosLinea - 0.5, TextoDirector
        Case "2.00" To "3.99"
             PrinterTexto 15.5, PosLinea - 0.5, TextoRector
      End Select
      Imprimir_Linea_H PosLinea, 14.6, 19, Negro
      PosLinea = PosLinea + PosLineaX
      Imprimir_Linea_H PosLinea - 0.9, 15, 18.5
      Select Case Curso
        Case "0.00" To "1.99"
             PrinterTexto 15.5, PosLinea - 0.6, TextoSecretario1
        Case "2.00" To "3.99"
             PrinterTexto 15.5, PosLinea - 0.6, TextoSecretario2
      End Select
      Cuadricula = False
      MensajeEncabData = "": SQLMsg1 = "": SQLMsg2 = "": SQLMsg3 = "": SQLMsg4 = ""
      PrinterFontSize 8
      PosLinea = LimiteAlto
      PrinterEndDoc
      RatonNormal
      Exit Sub
Errorhandler:
    RatonNormal
    ErrorDeImpresion
    Exit Sub
Else
    RatonNormal
End If
  End If
End With
End Sub

Public Sub Listar_Representantes(ElCurso As String)
Dim Y0 As Single
Dim y1 As Single
Dim X0 As Single
Dim x1 As Single
Dim PosLineaX As Single
On Error GoTo Errorhandler
With AdoAlumnos.Recordset
 If .RecordCount > 0 Then
   RatonReloj
   Mensajes = "Seguro de Imprimir en:" & vbCrLf & Printer.DeviceName & "?"
   Titulo = "IMPRESION"
   Bandera = False
   'MsgBox CodigoAlumno
   SetPrinters.Show 1
   If PonImpresoraDefecto(SetNombrePRN) Then
      InicioX = 0.5: InicioY = 0
      DataAnchoCampos InicioX, AdoAlumnos, 8, TipoCourier, Orientacion_Pagina, True
      PosLinea = 3.6
      Contador = 0
     .MoveFirst
      Cadena = Leer_Datos_del_Curso(ElCurso)
      MensajeEncabData = "Listado de Representante por Curso"
      SQLMsg1 = "AÑO LECTIVO " & Anio_Lectivo
      SQLMsg2 = Dato_Curso.Curso & " - " & Dato_Curso.Descripcion
      SQLMsg3 = ""
      PrinterFontBold True
      Encabezados
      PrinterFontBold True
      PrinterTexto 0.5, PosLinea, "No."
      PrinterTexto 1.2, PosLinea, "ALUMNO(A)"
      PrinterTexto 9, PosLinea, "REPRESENTANTE"
      PrinterTexto 13, PosLinea, "DIRECCION"
      PrinterTexto 23, PosLinea, "TELEFONOS"
      PosLinea = PosLinea + 0.4
      Printer.Line (0.5, PosLinea)-(28, PosLinea), QBColor(0)
      PosLinea = PosLinea + 0.05
      Printer.Line (0.5, PosLinea)-(28, PosLinea), QBColor(0)
      PosLinea = PosLinea + 0.1
      'PrinterFontBold False
      PrinterFontSize 8
      Printer.FontBold = False
      Do While Not .EOF
         Contador = Contador + 1
         Codigo2 = ""
         Printer.FontItalic = False
         If Val(.Fields("Telefono")) <> 0 Then Codigo2 = Codigo2 & .Fields("Telefono")
         If Val(.Fields("Telefono_R")) <> 0 Then Codigo2 = Codigo2 & "/" & .Fields("Telefono_R")
         If Val(.Fields("Celular")) Then Codigo2 = Codigo2 & "/" & .Fields("Celular")
         PrinterTexto 0.5, PosLinea, "|" & Format(Contador, "00") & ".-"
         PrinterTexto 1.2, PosLinea, "|" & .Fields("Alumno")
         PrinterTexto 9, PosLinea, "|" & .Fields("Representante_Alumno")
         PrinterTexto 13, PosLinea, "|" & .Fields("Domicilio")
         PrinterTexto 23, PosLinea, "|" & Codigo2
         PosLinea = PosLinea + 0.35
         If PosLinea >= LimiteAlto Then
            Imprimir_Linea_H PosLinea, 1, 19, Negro, True
            Printer.NewPage
            PosLinea = 3.6
            MensajeEncabData = "Listado de Representante por Curso"
            SQLMsg1 = "AÑO LECTIVO " & Anio_Lectivo
            SQLMsg2 = .Fields("Curso") & " - " & .Fields("Descripcion")
            SQLMsg3 = ""
            PrinterFontBold True
            Encabezados
            PrinterFontBold True
            PrinterTexto 0.5, PosLinea, "No."
            PrinterTexto 1.2, PosLinea, "ALUMNO(A)"
            PrinterTexto 9, PosLinea, "REPRESENTANTE"
            PrinterTexto 13, PosLinea, "DIRECCION"
            PrinterTexto 23, PosLinea, "TELEFONOS"
            PosLinea = PosLinea + 0.4
            Printer.Line (0.5, PosLinea)-(28, PosLinea), QBColor(0)
            PosLinea = PosLinea + 0.05
            Printer.Line (0.5, PosLinea)-(28, PosLinea), QBColor(0)
            PosLinea = PosLinea + 0.1
           'PrinterFontBold False
            PrinterFontSize 8
            Printer.FontBold = False
         End If
        .MoveNext
      Loop
      Printer.Line (0.5, PosLinea)-(28, PosLinea), QBColor(0)
      PosLinea = PosLinea + 0.05
      Printer.Line (0.5, PosLinea)-(28, PosLinea), QBColor(0)
      Cuadricula = False
      MensajeEncabData = "": SQLMsg1 = "": SQLMsg2 = "": SQLMsg3 = "": SQLMsg4 = ""
      RatonNormal
      PrinterEndDoc
      Exit Sub
Errorhandler:
    RatonNormal
    ErrorDeImpresion
    Exit Sub
Else
    RatonNormal
End If
  End If
End With
End Sub

Public Sub Llenar_Catalogo_Estudiantil()
  
' Establece propiedades del control ImageList.
  TVNivel.Nodes.Clear
  TVNivel.LineStyle = tvwTreeLines
' Crea un árbol con varios objetos Node sin ordenar.
  'Cadena1 = Tipo_Acceso_Educativo("", "CodigoE")
  '& Cadena1
  Contador = 0
  sSQL = "SELECT CE.*,CC.Descripcion,CM.Materia,CM.C,CM.I,CM.P,C.Cliente As Dirigente,C.Sexo " _
       & "FROM Catalogo_Estudiantil As CE, Catalogo_Cursos As CC, Catalogo_Materias AS CM, Clientes As C " _
       & "WHERE CE.Item = '" & NumEmpresa & "' " _
       & "AND CE.Periodo = '" & Periodo_Contable & "' " _
       & "AND CE.TC <> 'M' " _
       & "AND CE.Item = CC.Item " _
       & "AND CE.Item = CM.Item " _
       & "AND CE.Periodo = CC.Periodo " _
       & "AND CE.Periodo = CM.Periodo " _
       & "AND CE.CodigoE = CC.Curso " _
       & "AND CE.CodMat = CM.CodMat " _
       & "AND CE.Profesor = C.Codigo " _
       & "ORDER BY CE.CodigoE "
  SelectAdodc AdoPlantel, sSQL
  With ImgLstMenu 'ImgList
   If AdoPlantel.Recordset.RecordCount > 0 Then
      Do While Not AdoPlantel.Recordset.EOF
         Contador = Contador + 1
        'MsgBox Contador & " - " & AdoPlantel.Recordset.Fields("CodigoE")
         Codigo = "C" & AdoPlantel.Recordset.Fields("CodigoE")
         CodigoL = AdoPlantel.Recordset.Fields("CodigoE")
         TipoDoc = AdoPlantel.Recordset.Fields("TC")
         TipoProc = AdoPlantel.Recordset.Fields("CodMat")
         Select Case TipoDoc
           Case "M": Cadena = AdoPlantel.Recordset.Fields("Materia")
           Case Else
                Cadena = AdoPlantel.Recordset.Fields("Descripcion")
         End Select
         Codigos = CambioCodigoCtaSup(Codigo)
         If Len(Codigo) = 2 Then
            AddNewCta "C", Codigo, Cadena
            'Set nodX = TVNivel.Nodes.Add(, , Codigo, Cadena, ImgLstMenu.ListImages(1).key, ImgLstMenu.ListImages(1).key)
         Else
            AddNewCta TipoDoc, Codigo, Cadena
''            Select Case TipoDoc
''              Case "N": Set nodX = TVNivel.Nodes.Add(Codigos, tvwChild, Codigo, Cadena, .ListImages(2).key, .ListImages(2).key)
''              Case "P": Set nodX = TVNivel.Nodes.Add(Codigos, tvwChild, Codigo, Cadena, .ListImages(3).key, .ListImages(3).key)
''              Case "M": Set nodX = TVNivel.Nodes.Add(Codigos, tvwChild, Codigo, Cadena, .ListImages(4).key, .ListImages(4).key)
''            End Select
         End If
         AdoPlantel.Recordset.MoveNext
      Loop
   Else
      RatonNormal
      'Unload FLibretas
   End If
  End With
End Sub

Public Sub EliminarCta()
  Codigo1 = Mid$(TVNivel.SelectedItem.key, 2, Len(TVNivel.SelectedItem.key) - 1)
  Cadena = SinEspaciosIzq(TVNivel.SelectedItem.key)
  With AdoLibreta.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("CodigoE like '" & Codigo1 & "' ")
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

Private Sub Command1_Click()
  RatonNormal
  Unload Me
End Sub

Private Sub Command2_Click()
  HScroll1.SetFocus
End Sub

Private Sub Command3_Click()
  VScroll1.SetFocus
End Sub

Public Sub Nota_Maxima_Periodo()
Dim CamposProm(13) As String
Dim CampoAux As String
Dim ContProm As Integer
   'Asignacion de Notas Valores Maximos
    CamposProm(0) = SQLTAI
    CamposProm(1) = SQLAIC
    CamposProm(2) = SQLAGC
    CamposProm(3) = SQLL
    CamposProm(4) = SQLExaP
    CamposProm(5) = SQLBim1
    CamposProm(6) = SQLBim2
    CamposProm(7) = SQLBim3
    CamposProm(8) = SQLExamen
    CamposProm(9) = "" 'SQLConductaQ
    CamposProm(10) = "Nota_Grado"
    CamposProm(11) = "Supletorio"
    For ContProm = 0 To 12
        CampoAux = CamposProm(ContProm)
        If Len(CampoAux) > 1 Then
           Progreso_Barra.Mensaje_Box = "Verificando el Valor de la Nota: " & CampoAux
           Progreso_Esperar
           sSQL = "UPDATE Trans_Notas " _
                & "SET " & CampoAux & " = " & Nota_Mayor & " " _
                & "WHERE Item = '" & NumEmpresa & "' " _
                & "AND Periodo = '" & Periodo_Contable & "' " _
                & "AND " & CampoAux & " > " & Nota_Mayor & " "
           ConectarAdoExecute sSQL
           sSQL = "UPDATE Trans_Notas_Auxiliares " _
                & "SET " & CampoAux & " = " & Nota_Mayor & " " _
                & "WHERE Item = '" & NumEmpresa & "' " _
                & "AND Periodo = '" & Periodo_Contable & "' " _
                & "AND " & CampoAux & " > " & Nota_Mayor & " "
           ConectarAdoExecute sSQL
        End If
    Next ContProm
End Sub

Public Sub Recalcular_Notas_CodMatP()
Dim CamposProm(7) As String
Dim CampoAux As String
Dim TxtSQLSuma As String
Dim ContProm As Integer
Dim Valor_Nota As Currency
    Progreso_Barra.Mensaje_Box = "Actualizano Sub Nota con notas finales: "
    Progreso_Esperar
    sSQL = "UPDATE Trans_Notas " _
         & "SET X = '.' " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND X <> '.' "
    ConectarAdoExecute sSQL
    
    Progreso_Esperar
    sSQL = "UPDATE Trans_Notas " _
         & "SET X = 'S' " _
         & "FROM Trans_Notas As TN, Trans_Notas_Auxiliares As TNA " _
         & "WHERE TN.Item = '" & NumEmpresa & "' " _
         & "AND TN.Periodo = '" & Periodo_Contable & "' " _
         & "AND TN.Item = TNA.Item " _
         & "AND TN.Periodo = TNA.Periodo " _
         & "AND TN.CodMat = TNA.CodMatP " _
         & "AND TN.Codigo = TNA.Codigo "
    ConectarAdoExecute sSQL
   'Asignacion de Notas promediadas
    CamposProm(0) = SQLTAI
    CamposProm(1) = SQLAIC
    CamposProm(2) = SQLAGC
    CamposProm(3) = SQLL
    CamposProm(4) = SQLExaP
    If OpcionNotas < 4 Then
       sSQL = "UPDATE Trans_Notas_Auxiliares " _
            & "SET " & SQLProm & " = ROUND((" & SQLTAI & "+" & SQLAIC & "+" & SQLAGC & "+" & SQLL & "+" & SQLExaP & ")/ContNotas,2,0) " _
            & "WHERE Item = '" & NumEmpresa & "' " _
            & "AND Periodo = '" & Periodo_Contable & "' "
       ConectarAdoExecute sSQL
    ElseIf OpcionNotas = 4 Then
       Select Case FormatoLibreta
         Case "QUIMESTRE"
              TxtSQLSuma = "(" & SQLBim1 & " + " & SQLBim2 & " + " & SQLBim3 & ")/3"
         Case "TRIMESTRE2"
              TxtSQLSuma = "(" & SQLBim1 & " + " & SQLBim2 & " + " & SQLExamen & ")/3"
         Case "PERIODO"
              TxtSQLSuma = "(" & SQLBim1 & " + " & SQLBim2 & ")/2"
         Case Else
              TxtSQLSuma = "(" & SQLBim1 & " + " & SQLBim2 & ")/2"
       End Select
       Progreso_Esperar
       sSQL = "UPDATE Trans_Notas_Auxiliares " _
            & "SET " & SQLQPX & " = ROUND((" & TxtSQLSuma & ") * " & CStr(Q_PX) & "," & Dec_Nota & ",0) " _
            & "WHERE Item = '" & NumEmpresa & "' " _
            & "AND Periodo = '" & Periodo_Contable & "' " _
            & "AND " & SQLQPX & " <> ROUND((" & TxtSQLSuma & ") * " & CStr(Q_PX) & "," & Dec_Nota & ",0) "
       ConectarAdoExecute sSQL
       
       Progreso_Esperar
       sSQL = "UPDATE Trans_Notas_Auxiliares " _
            & "SET " & SQLQEX & " = ROUND((" & SQLExamen & ") * " & CStr(Q_EX) & "," & Dec_Nota & ",0) " _
            & "WHERE Item = '" & NumEmpresa & "' " _
            & "AND Periodo = '" & Periodo_Contable & "' " _
            & "AND " & SQLQEX & " <> ROUND((" & TxtSQLSuma & ") * " & CStr(Q_EX) & "," & Dec_Nota & ",0) "
       ConectarAdoExecute sSQL
       Progreso_Esperar
       sSQL = "UPDATE Trans_Notas_Auxiliares " _
            & "SET " & SQLPromQ & " = ROUND(" & SQLQPX & " + " & SQLQEX & "," & Dec_Nota & ",0) " _
            & "WHERE Item = '" & NumEmpresa & "' " _
            & "AND Periodo = '" & Periodo_Contable & "' " _
            & "AND " & SQLPromQ & " <> ROUND(" & SQLQPX & " + " & SQLQEX & "," & Dec_Nota & ",0) "
       ConectarAdoExecute sSQL
       
        sSQL = "UPDATE Trans_Notas_Auxiliares " _
             & "SET " & SQLPromQ & " = ROUND((" & SQLBim1 & "+" & SQLBim2 & "+" & SQLBim3 & ")/3," & Dec_Nota & ",0) " _
             & "WHERE Item = '" & NumEmpresa & "' " _
             & "AND Periodo = '" & Periodo_Contable & "' " _
             & "AND " & SQLBim1 & " > 0 " _
             & "AND " & SQLBim2 & " > 0 " _
             & "AND " & SQLBim3 & " > 0 " _
             & "AND MID(CodE,1,4) <= '1.01' "
        ConectarAdoExecute sSQL
    End If
    If OpcionNotas <> 4 Then CamposProm(6) = SQLExamen Else CamposProm(6) = Ninguno
    For ContProm = 0 To 6
        CampoAux = CamposProm(ContProm)
         If Len(CampoAux) > 1 Then
            Progreso_Barra.Mensaje_Box = "Actualizano Sub Nota: " & CampoAux
            Progreso_Esperar
            
            SubSQL = "SELECT ROUND(AVG(" & CampoAux & ")," & Dec_Nota & ",0) As T" & CampoAux & " " _
                   & "FROM Trans_Notas_Auxiliares " _
                   & "WHERE Trans_Notas_Auxiliares." & CampoAux & " <> 0 " _
                   & "AND Trans_Notas_Auxiliares.Item = Trans_Notas.Item " _
                   & "AND Trans_Notas_Auxiliares.Periodo = Trans_Notas.Periodo " _
                   & "AND Trans_Notas_Auxiliares.CodMatP = Trans_Notas.CodMat " _
                   & "AND Trans_Notas_Auxiliares.Codigo = Trans_Notas.Codigo "
                   
            sSQL = "UPDATE Trans_Notas " _
                 & "SET " & CampoAux & " = (" & SubSQL & ") " _
                 & "WHERE Item = '" & NumEmpresa & "' " _
                 & "AND Periodo = '" & Periodo_Contable & "' " _
                 & "AND X = 'S' "
            ConectarAdoExecute sSQL
            
            sSQL = "UPDATE Trans_Notas " _
                 & "SET " & CamposProm(ContProm) & " = 0 " _
                 & "WHERE Item = '" & NumEmpresa & "' " _
                 & "AND Periodo = '" & Periodo_Contable & "' " _
                 & "AND " & CamposProm(ContProm) & " IS NULL "
            ConectarAdoExecute sSQL
         End If
    Next ContProm
End Sub

Public Sub Recalcular_Notas()
Dim ContPrim As Long
Dim ContSecD As Long
Dim ContSecV As Long
Dim ContMatr As Long
Dim Imp_Mat As Byte
Dim TxtBim As String
Dim TxtSQLSuma As String
Dim TxtBimCon As String
Dim TxtBimExa As String
Dim Periodo_Notas As Byte
Dim Periodo_Notas_S As String
Dim SDI As String
Dim Valor_Nota As Currency
Dim SQLNotaParcial(5) As String
Dim IdNotaP As Integer
    
    MiTiempo = Time
    SDI = FLibretas.Caption
   'Tipo de Recalculacion de Notas
    RatonReloj
    Periodo_Notas = 0
    Periodo_Notas_S = "Todas"
    If OpcPeriodo("PQBim1", LstPeriodos) Or OpcPeriodo("PQ", LstPeriodos) Then
       Periodo_Notas = 1
       Periodo_Notas_S = "Empezando a Recalcular Primer Periodo"
    End If
    If OpcPeriodo("SQBim1", LstPeriodos) Or OpcPeriodo("SQ", LstPeriodos) Then
       Periodo_Notas = 2
       Periodo_Notas_S = "Empezando a Recalcular Segundo Periodo"
    End If
    If OpcPeriodo("TQBim1", LstPeriodos) Or OpcPeriodo("TQ", LstPeriodos) Then
       Periodo_Notas = 3
       Periodo_Notas_S = "Empezando a Recalcular Tercer Periodo"
    End If
    FLibretas.Caption = Periodo_Notas_S
    
    Progreso_Barra.Incremento = 0
    Progreso_Barra.Valor_Maximo = 100
    
    Progreso_Barra.Mensaje_Box = "RECALCULAR SUB-NOTAS"
    Progreso_Esperar
    
    sSQL = "SELECT * " _
         & "FROM Clientes " _
         & "WHERE FA <> " & Val(adFalse) & " " _
         & "ORDER BY Grupo,Cliente "
    SelectAdodc AdoAux, sSQL
   'If AdoAux.Recordset.RecordCount > 0 Then Progreso_Barra.Valor_Maximo = Progreso_Barra.Valor_Maximo + AdoAux.Recordset.RecordCount
    
   'Verificamos que la nota no sea mayor que 20
    Progreso_Barra.Mensaje_Box = "Verificando Notas Incorrectas"
    Progreso_Esperar
    Nota_Maxima_Periodo
    
    SQLNotaParcial(0) = SQLTAI
    SQLNotaParcial(1) = SQLAIC
    SQLNotaParcial(2) = SQLAGC
    SQLNotaParcial(3) = SQLL
    SQLNotaParcial(4) = SQLExaP
    
   'Actualizando cuantas notas estan ingresadas
    sSQL = "UPDATE Trans_Notas " _
         & "SET ContNotas = 5 " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' "
    ConectarAdoExecute sSQL
    If SQL_Server Then
       sSQL = "UPDATE Trans_Notas " _
            & "SET Horas = CE.Horas_Clase " _
            & "FROM Trans_Notas As TN, Catalogo_Estudiantil As CE "
    Else
       sSQL = "UPDATE Trans_Notas As TN, Catalogo_Estudiantil As CE " _
            & "SET TN.Horas = CE.Horas_Clase "
    End If
    sSQL = sSQL _
         & "WHERE TN.Item = '" & NumEmpresa & "' " _
         & "AND TN.Periodo = '" & Periodo_Contable & "' " _
         & "AND TN.CodMat = CE.CodMat " _
         & "AND TN.CodE = Mid$(CE.CodigoE,1,7) " _
         & "AND TN.Item = CE.Item " _
         & "AND TN.Periodo = CE.Periodo "
    ConectarAdoExecute sSQL
    
    For IdNotaP = 0 To 4
        If SQLNotaParcial(IdNotaP) <> Ninguno Then
           sSQL = "UPDATE Trans_Notas " _
                & "SET ContNotas = ContNotas - 1 " _
                & "WHERE Item = '" & NumEmpresa & "' " _
                & "AND Periodo = '" & Periodo_Contable & "' " _
                & "AND " & SQLNotaParcial(IdNotaP) & " <= 0 " _
                & "AND Horas < " & Horas_Min & " "
           ConectarAdoExecute sSQL
        End If
    Next IdNotaP
    sSQL = "UPDATE Trans_Notas " _
         & "SET ContNotas = 1 " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND ContNotas <= 0 "
    ConectarAdoExecute sSQL
    
    sSQL = "UPDATE Trans_Notas_Auxiliares " _
         & "SET ContNotas = 5 " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' "
    ConectarAdoExecute sSQL
    If SQL_Server Then
       sSQL = "UPDATE Trans_Notas_Auxiliares " _
            & "SET Horas = CE.Horas_Clase " _
            & "FROM Trans_Notas_Auxiliares As TN, Catalogo_Estudiantil As CE "
    Else
       sSQL = "UPDATE Trans_Notas_Auxiliares As TN, Catalogo_Estudiantil As CE " _
            & "SET TN.Horas = CE.Horas_Clase "
    End If
    sSQL = sSQL _
         & "WHERE TN.Item = '" & NumEmpresa & "' " _
         & "AND TN.Periodo = '" & Periodo_Contable & "' " _
         & "AND TN.CodMat = CE.CodMat " _
         & "AND TN.CodE = Mid$(CE.CodigoE,1,7) " _
         & "AND TN.Item = CE.Item " _
         & "AND TN.Periodo = CE.Periodo "
    ConectarAdoExecute sSQL
    For IdNotaP = 0 To 4
        If SQLNotaParcial(IdNotaP) <> Ninguno Then
            sSQL = "UPDATE Trans_Notas_Auxiliares " _
                 & "SET ContNotas = ContNotas - 1 " _
                 & "WHERE Item = '" & NumEmpresa & "' " _
                 & "AND Periodo = '" & Periodo_Contable & "' " _
                 & "AND " & SQLNotaParcial(IdNotaP) & " <= 0 " _
                 & "AND Horas < " & Horas_Min & " "
            ConectarAdoExecute sSQL
        End If
    Next IdNotaP
    sSQL = "UPDATE Trans_Notas_Auxiliares " _
         & "SET ContNotas = 1 " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND ContNotas <= 0 "
    ConectarAdoExecute sSQL
    
   'Aprobar por defaulto todos los estudiantes
    sSQL = "UPDATE Clientes_Matriculas " _
         & "SET Aprobado = " & Val(adTrue) & " " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' "
    ConectarAdoExecute sSQL
     
   'Actualizamos los subtotales de las Sub-Notas
    Progreso_Barra.Mensaje_Box = "Recalculando Sub-Notas"
    Progreso_Esperar
    Recalcular_Notas_CodMatP
   
   'Enceramos la materia de disciplina
    If OpcionNotas < 4 Then
       Progreso_Barra.Mensaje_Box = "Actualizando Disciplinas"
       Progreso_Esperar
       If SQL_Server Then
          sSQL = "UPDATE Trans_Notas " _
               & "SET " & SQLProm & " = TA." & SQLConductaQ & " " _
               & "FROM Trans_Notas As TN, Trans_Asistencia As TA "
       Else
          sSQL = "UPDATE Trans_Notas As TN, Trans_Asistencia As TA " _
               & "SET TN." & SQLProm & " = TA." & SQLConductaQ & " "
       End If
       sSQL = sSQL _
            & "WHERE TN.Item = '" & NumEmpresa & "' " _
            & "AND TN.Periodo = '" & Periodo_Contable & "' " _
            & "AND TA.CodMat IN ('997','998','999') " _
            & "AND TN.CodMat = TA.CodMat " _
            & "AND TN.Codigo = TA.Codigo " _
            & "AND TN.Item = TA.Item " _
            & "AND TN.Periodo = TA.Periodo " _
            & "AND TN.CodE = TA.CodE "
       ConectarAdoExecute sSQL
      'Calculamos Promedio de Parciales en Quimestres o Trimestres
      
       TxtSQLSuma = SQLTAI & " + " & SQLAIC & " + " & SQLAGC & " + " & SQLL & " + " & SQLExaP
       sSQL = "UPDATE Trans_Notas " _
            & "SET " & SQLNotas & " = ROUND((" & TxtSQLSuma & ")/ContNotas," & Dec_Nota & ",0) " _
            & "WHERE Item = '" & NumEmpresa & "' " _
            & "AND Periodo = '" & Periodo_Contable & "' " _
            & "AND CodMat NOT IN ('997','998','999') " _
            & "AND " & SQLNotas & " <> ROUND((" & TxtSQLSuma & ")/ContNotas," & Dec_Nota & ",0) "
       ConectarAdoExecute sSQL
        
'      'Calculamos los promedios si es Pre-Basica
'       SQLNotaParcial(0) = SQLTAI
'       SQLNotaParcial(1) = SQLAIC
'       SQLNotaParcial(2) = SQLAGC
'       SQLNotaParcial(3) = SQLL
'       SQLNotaParcial(4) = SQLExaP
'       For IdNotaP = 0 To 4
'           sSQL = "UPDATE Trans_Notas " _
'                & "SET " & SQLNotaParcial(IdNotaP) & " = ROUND(" & SQLNotaParcial(IdNotaP) & "," & Dec_Nota & ",0) " _
'                & "WHERE Item = '" & NumEmpresa & "' " _
'                & "AND Periodo = '" & Periodo_Contable & "' " _
'                & "AND CodMat NOT IN ('997','998','999') " _
'                & "AND " & SQLNotaParcial(IdNotaP) & " <> ROUND(" & SQLNotaParcial(IdNotaP) & "," & Dec_Nota & ",0) " _
'                & "AND CodE <= '1.01' "
'           ConectarAdoExecute sSQL
'       Next IdNotaP
       
'       TxtSQLSuma = SQLTAI & " + " & SQLAIC & " + " & SQLAGC & " + " & SQLL & " + " & SQLExaP
'       sSQL = "UPDATE Trans_Notas " _
'            & "SET " & SQLNotas & " = ROUND((" & TxtSQLSuma & ")/5," & Dec_Nota & ",0) " _
'            & "WHERE Item = '" & NumEmpresa & "' " _
'            & "AND Periodo = '" & Periodo_Contable & "' " _
'            & "AND CodMat NOT IN ('997','998','999') " _
'            & "AND " & SQLNotas & " <> ROUND((" & TxtSQLSuma & ")/5," & Dec_Nota & ",0) " _
'            & "AND CodE <= '1.01' "
'       ConectarAdoExecute sSQL
      
        sSQL = "UPDATE Trans_Notas " _
             & "SET " & SQLPromQ & " = ROUND(" & SQLBim1 & "," & Dec_Nota & ",0) " _
             & "WHERE Item = '" & NumEmpresa & "' " _
             & "AND Periodo = '" & Periodo_Contable & "' " _
             & "AND " & SQLBim1 & " > 0 " _
             & "AND " & SQLBim2 & " <= 0 " _
             & "AND " & SQLBim3 & " <= 0 " _
             & "AND CodE <= '1.01' "
        ConectarAdoExecute sSQL
        
        sSQL = "UPDATE Trans_Notas " _
             & "SET " & SQLPromQ & " = ROUND((" & SQLBim1 & "+" & SQLBim2 & ")/2," & Dec_Nota & ",0) " _
             & "WHERE Item = '" & NumEmpresa & "' " _
             & "AND Periodo = '" & Periodo_Contable & "' " _
             & "AND " & SQLBim1 & " > 0 " _
             & "AND " & SQLBim2 & " > 0 " _
             & "AND " & SQLBim3 & " <= 0 " _
             & "AND CodE <= '1.01' "
        ConectarAdoExecute sSQL
    ElseIf OpcionNotas = 4 Then
      'Actualizamos la Promedios de disciplina
       Progreso_Barra.Mensaje_Box = "Actualizano Promedio de Conducta"
       Progreso_Esperar
       sSQL = "UPDATE Trans_Notas " _
            & "SET " & SQLPromQ & " = ROUND(" & SQLBim1 & "," & Dec_Nota & ",0) " _
            & "WHERE Item = '" & NumEmpresa & "' " _
            & "AND Periodo = '" & Periodo_Contable & "' " _
            & "AND " & SQLBim1 & " > 0 " _
            & "AND " & SQLBim2 & " <= 0 " _
            & "AND CodMat IN ('997','998','999') "
       'MsgBox sSQL
       ConectarAdoExecute sSQL
         
       sSQL = "UPDATE Trans_Notas " _
            & "SET " & SQLPromQ & " = ROUND(" & SQLBim2 & "," & Dec_Nota & ",0) " _
            & "WHERE Item = '" & NumEmpresa & "' " _
            & "AND Periodo = '" & Periodo_Contable & "' " _
            & "AND " & SQLBim1 & " <= 0 " _
            & "AND " & SQLBim2 & " > 0 " _
            & "AND CodMat IN ('997','998','999') "
       ConectarAdoExecute sSQL
          
       sSQL = "UPDATE Trans_Notas " _
            & "SET " & SQLPromQ & " = ROUND((" & SQLBim1 & "+" & SQLBim2 & ")/2," & Dec_Nota & ",0) " _
            & "WHERE Item = '" & NumEmpresa & "' " _
            & "AND Periodo = '" & Periodo_Contable & "' " _
            & "AND " & SQLBim1 & " > 0 " _
            & "AND " & SQLBim2 & " > 0 " _
            & "AND " & SQLBim3 & " <= 0 " _
            & "AND CodMat IN ('997','998','999') "
       ConectarAdoExecute sSQL
       Select Case Anio_Lectivo
         Case Is >= "2014 - 2015", Ninguno
              sSQL = "UPDATE Trans_Notas " _
                   & "SET " & SQLPromQ & " = ROUND(" & SQLBim3 & "," & Dec_Nota & ",0) " _
                   & "WHERE Item = '" & NumEmpresa & "' " _
                   & "AND Periodo = '" & Periodo_Contable & "' " _
                   & "AND " & SQLBim1 & " > 0 " _
                   & "AND " & SQLBim2 & " > 0 " _
                   & "AND " & SQLBim3 & " > 0 " _
                   & "AND CodMat IN ('997','998','999') "
              ConectarAdoExecute sSQL
             'MsgBox OpcionNotas & vbCrLf & sSQL
         Case Else
              sSQL = "UPDATE Trans_Notas " _
                   & "SET " & SQLPromQ & " = ROUND((" & SQLBim1 & "+" & SQLBim2 & "+" & SQLBim3 & ")/3," & Dec_Nota & ",0) " _
                   & "WHERE Item = '" & NumEmpresa & "' " _
                   & "AND Periodo = '" & Periodo_Contable & "' " _
                   & "AND " & SQLBim1 & " > 0 " _
                   & "AND " & SQLBim2 & " > 0 " _
                   & "AND " & SQLBim3 & " > 0 " _
                   & "AND CodMat IN ('997','998','999') "
              ConectarAdoExecute sSQL
       End Select
      'MsgBox Anio_Lectivo & vbCrLf & sSQL
       Select Case FormatoLibreta
         Case "QUIMESTRE"
              TxtSQLSuma = "(" & SQLBim1 & " + " & SQLBim2 & " + " & SQLBim3 & ")/3"
         Case "TRIMESTRE2"
              TxtSQLSuma = "(" & SQLBim1 & " + " & SQLBim2 & " + " & SQLExamen & ")/3"
         Case "PERIODO"
              TxtSQLSuma = "(" & SQLBim1 & " + " & SQLBim2 & ")/2"
         Case Else
              TxtSQLSuma = "(" & SQLBim1 & " + " & SQLBim2 & ")/2"
       End Select
       Progreso_Esperar
       
       sSQL = "UPDATE Trans_Notas " _
            & "SET " & SQLQPX & " = ROUND((" & TxtSQLSuma & ") * " & CStr(Q_PX) & "," & Dec_Nota & ",0) " _
            & "WHERE Item = '" & NumEmpresa & "' " _
            & "AND Periodo = '" & Periodo_Contable & "' " _
            & "AND CodMat NOT IN ('997','998','999') " _
            & "AND " & SQLQPX & " <> ROUND((" & TxtSQLSuma & ") * " & CStr(Q_PX) & "," & Dec_Nota & ",0) "
       ConectarAdoExecute sSQL
       
       Progreso_Esperar
       sSQL = "UPDATE Trans_Notas " _
            & "SET " & SQLQEX & " = ROUND((" & SQLExamen & ") * " & CStr(Q_EX) & "," & Dec_Nota & ",0) " _
            & "WHERE Item = '" & NumEmpresa & "' " _
            & "AND Periodo = '" & Periodo_Contable & "' " _
            & "AND CodMat NOT IN ('997','998','999') " _
            & "AND " & SQLQEX & " <> ROUND((" & TxtSQLSuma & ") * " & CStr(Q_EX) & "," & Dec_Nota & ",0) "
       ConectarAdoExecute sSQL
       
       Progreso_Esperar
       sSQL = "UPDATE Trans_Notas " _
            & "SET " & SQLPromQ & " = ROUND(" & SQLQPX & " + " & SQLQEX & "," & Dec_Nota & ",0) " _
            & "WHERE Item = '" & NumEmpresa & "' " _
            & "AND Periodo = '" & Periodo_Contable & "' " _
            & "AND CodMat NOT IN ('997','998','999') " _
            & "AND " & SQLPromQ & " <> ROUND(" & SQLQPX & " + " & SQLQEX & "," & Dec_Nota & ",0) "
       ConectarAdoExecute sSQL
       
        sSQL = "UPDATE Trans_Notas " _
             & "SET " & SQLPromQ & " = ROUND((" & SQLBim1 & "+" & SQLBim2 & "+" & SQLBim3 & ")/3," & Dec_Nota & ",0) " _
             & "WHERE Item = '" & NumEmpresa & "' " _
             & "AND Periodo = '" & Periodo_Contable & "' " _
             & "AND " & SQLBim1 & " > 0 " _
             & "AND " & SQLBim2 & " > 0 " _
             & "AND " & SQLBim3 & " > 0 " _
             & "AND MID(CodE,1,4) <= '1.01' "
        ConectarAdoExecute sSQL
        'MsgBox sSQL
    ElseIf OpcionNotas = 5 Then
       Progreso_Barra.Mensaje_Box = "Actualizando Disciplinas"
       Progreso_Esperar
       Select Case Anio_Lectivo
         Case Is >= "2014 - 2015", Ninguno:
              sSQL = "UPDATE Trans_Notas " _
                   & "SET PromSQ = ROUND(SQBim3," & Dec_Nota & ",0) " _
                   & "WHERE Item = '" & NumEmpresa & "' " _
                   & "AND Periodo = '" & Periodo_Contable & "' " _
                   & "AND CodMat IN ('997','998','999') "
              ConectarAdoExecute sSQL
              sSQL = "UPDATE Trans_Notas " _
                   & "SET PromFinal = ROUND(SQBim3," & Dec_Nota & ",0) " _
                   & "FROM Trans_Notas " _
                   & "WHERE Item = '" & NumEmpresa & "' " _
                   & "AND Periodo = '" & Periodo_Contable & "' " _
                   & "AND CodMat IN ('997','998','999') "
              ConectarAdoExecute sSQL
             'MsgBox OpcionNotas & vbCrLf & sSQL
         Case Else
              sSQL = "UPDATE Trans_Notas " _
                   & "SET PromFinal = ROUND((PromPQ+PromSQ)/2," & Dec_Nota & ",0) " _
                   & "FROM Trans_Notas " _
                   & "WHERE Item = '" & NumEmpresa & "' " _
                   & "AND Periodo = '" & Periodo_Contable & "' " _
                   & "AND CodMat IN ('997','998','999') "
              ConectarAdoExecute sSQL
       End Select
       Progreso_Barra.Mensaje_Box = "Actualizano Promedios Finales"
       Progreso_Esperar
       sSQL = "UPDATE Trans_Notas " _
            & "SET PromFinal = PromPQ " _
            & "WHERE Item = '" & NumEmpresa & "' " _
            & "AND Periodo = '" & Periodo_Contable & "' " _
            & "AND PromPQ > 0 " _
            & "AND PromSQ <= 0 "
       ConectarAdoExecute sSQL
     
       sSQL = "UPDATE Trans_Notas " _
            & "SET PromFinal = PromSQ " _
            & "WHERE Item = '" & NumEmpresa & "' " _
            & "AND Periodo = '" & Periodo_Contable & "' " _
            & "AND PromPQ <= 0 " _
            & "AND PromSQ > 0 "
       ConectarAdoExecute sSQL
              
       sSQL = "UPDATE Trans_Notas " _
            & "SET PromFinal = ROUND((PromPQ + PromSQ)/2," & Dec_Nota & ",1) " _
            & "WHERE Item = '" & NumEmpresa & "' " _
            & "AND Periodo = '" & Periodo_Contable & "' " _
            & "AND PromPQ > 0 " _
            & "AND PromSQ > 0 " _
            & "AND PromTQ <= 0 "
       ConectarAdoExecute sSQL
     
       sSQL = "UPDATE Trans_Notas " _
            & "SET PromFinal = ROUND((PromPQ + PromSQ + PromTQ)/3," & Dec_Nota & ",1) " _
            & "WHERE Item = '" & NumEmpresa & "' " _
            & "AND Periodo = '" & Periodo_Contable & "' " _
            & "AND PromPQ > 0 " _
            & "AND PromSQ > 0 " _
            & "AND PromTQ > 0 "
       ConectarAdoExecute sSQL

      'Calculando los promedios si hay supletorio o remedial
       sSQL = "UPDATE Trans_Notas " _
            & "SET PromFinal = ROUND((PromFinal + Supletorio)/2," & Dec_Nota & ",1) " _
            & "WHERE CodE BETWEEN '1.02' and '5.99' " _
            & "AND Item = '" & NumEmpresa & "' " _
            & "AND Periodo = '" & Periodo_Contable & "' " _
            & "AND PromFinal < " & Nota_Rojo & " " _
            & "AND Supletorio > 0 " _
            & "AND Remedial <= 0 "
       ConectarAdoExecute sSQL
       
       sSQL = "UPDATE Trans_Notas " _
            & "SET PromFinal = " & Nota_Rojo & " " _
            & "WHERE CodE BETWEEN '1.02' and '5.99' " _
            & "AND Item = '" & NumEmpresa & "' " _
            & "AND Periodo = '" & Periodo_Contable & "' " _
            & "AND Supletorio >= " & Nota_Rojo & " "
       ConectarAdoExecute sSQL
          
       sSQL = "UPDATE Trans_Notas " _
            & "SET PromFinal = ROUND((PromFinal + Remedial)/2," & Dec_Nota & ",1) " _
            & "WHERE CodE BETWEEN '1.02' and '5.99' " _
            & "AND Item = '" & NumEmpresa & "' " _
            & "AND Periodo = '" & Periodo_Contable & "' " _
            & "AND PromFinal < " & Nota_Rojo & " " _
            & "AND Supletorio >= " & Nota_Rojo & " " _
            & "AND Remedial > 0 "
       ConectarAdoExecute sSQL

       sSQL = "UPDATE Trans_Notas " _
            & "SET PromFinal = " & Nota_Rojo & " " _
            & "WHERE CodE BETWEEN '1.02' and '5.99' " _
            & "AND Item = '" & NumEmpresa & "' " _
            & "AND Periodo = '" & Periodo_Contable & "' " _
            & "AND Remedial >= " & Nota_Rojo & " "
       ConectarAdoExecute sSQL
     
       If SQL_Server Then
          sSQL = "UPDATE Clientes_Matriculas " _
               & "SET Aprobado = " & Val(adFalse) & " " _
               & "FROM Clientes_Matriculas As CM, Trans_Notas As TN "
       Else
          sSQL = "UPDATE Clientes_Matriculas As CM, Trans_Notas As TN " _
               & "SET CM.Aprobado = " & Val(adFalse) & " "
       End If
       sSQL = sSQL _
            & "WHERE CM.Item = '" & NumEmpresa & "' " _
            & "AND CM.Periodo = '" & Periodo_Contable & "' " _
            & "AND TN.PromFinal < " & Nota_Rojo & " " _
            & "AND CM.Codigo = TN.Codigo " _
            & "AND CM.Item = TN.Item " _
            & "AND CM.Periodo = TN.Periodo "
       ConectarAdoExecute sSQL
    End If
 
   'Procesamos la Conducta/Disciplina
    Progreso_Barra.Mensaje_Box = "Actualizano La Conducta o Disciplina"
    Progreso_Esperar
    Contador = 0
    
    Consultamos_Disciplinas AdoMatriculas
    Progreso_Esperar
  
    Progreso_Barra.Mensaje_Box = "Recalculo de notas exitoso " & Format(Time - MiTiempo, "hh:mm:ss")
    Progreso_Barra.Incremento = Progreso_Barra.Valor_Maximo
    Progreso_Esperar
    FLibretas.Caption = "CATALOGO ESTUDIANTIL"
    RatonNormal
    MsgBox "Proceso terminado," & vbCrLf & "puede Listar Libretas "
    LstPeriodos.SetFocus
End Sub

'Public Sub PrintObjFields(TipoObjeto As Object, _
'                          Xo As Single, _
'                          Yo As Single, _
'                          Texto As String, _
'                          Optional PonerLineas As Boolean, _
'                          Optional ImpLineaCero As Boolean)
''If Yo <= LimiteAlto Then
'If ((Xo > 0) And (Yo > 0)) Then
'   TipoObjeto.CurrentX = Xo
'   TipoObjeto.CurrentY = Yo
'   TipoObjeto.Print Texto
'End If
'End Sub

Private Sub DGNotasLibreta_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  If CtrlDown And KeyCode = vbKeyF1 Then
     DGNotasLibreta.Visible = False
     GenerarDataTexto FLibretas, AdoNotasLibreta
     DGNotasLibreta.Visible = True
  End If
End Sub

Private Sub Form_Activate()
Dim clmX As ColumnHeader
  'AltoMaximo = 29.7
  'AnchoMaximo = 21
  SSTab1.Tab = 1
  SSTab1.width = MDI_X_Max - SSTab1.Left - 40
  SSTab1.Height = MDI_Y_Max - SSTab1.Top - 40  '3650
  If Modulo = "TUTORIA" Then
     Toolbar1.Buttons("Imprimir").Enabled = False
     Toolbar1.Buttons("Actas").Enabled = False
     Toolbar1.Buttons("Matriculado").Enabled = False
     Toolbar1.Buttons("SegundaPg").Enabled = False
     Toolbar1.Buttons("Libretas").Enabled = False
     Toolbar1.Buttons("Promocion").Enabled = False
     Toolbar1.Buttons("Aptitud").Enabled = False
     Toolbar1.Buttons("CuadroProm").Enabled = False
     Toolbar1.Buttons("Certificados").Enabled = False
     Toolbar1.Buttons("SolExaGrado").Enabled = False
     Toolbar1.Buttons("AprobarExamGrado").Enabled = False
     Toolbar1.Buttons("NotasExamGrado").Enabled = False
     Toolbar1.Buttons("Supletorio").Enabled = False
     Toolbar1.Buttons("Recalcular").Enabled = False
  End If
  LblFormato.Caption = FormatoLibreta
  MBFecha = FechaSistema
           
  Consultamos_Disciplinas AdoMatriculas
  
  sSQL = "SELECT * " _
       & "FROM Catalogo_Materias " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "ORDER BY Materia "
  SelectAdodc AdoMaterias, sSQL
  
  sSQL = "SELECT * " _
       & "FROM Catalogo_Cursos " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "ORDER BY Curso "
  SelectAdodc AdoCursos, sSQL
  
  sSQL = "SELECT * " _
       & "FROM Catalogo_Equivalencia " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "ORDER BY Desde,Hasta "
  SelectAdodc AdoEquivalencia, sSQL

  Contador = 0
  Leer_Periodo_Lectivo
  
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
' Etiqueta controles OptionButton con opciones de View.
'  Option1(0).Caption = "Icono"
'  Option1(1).Caption = "Pequeño"
'  Option1(2).Caption = "Lista"
'  Option1(3).Caption = "Reporte"
'  TVNivel.ToolTipText = "<F4> Cuadro de Notas Promediales, " _
'                      & "<Enter> Listar Alumnos."
'  LstVAlumnos.ToolTipText = "<F1> Lista Libreta, <Ctrl+I> Informe Final, <Ctrl+A> Certificado de Aptitud, " _
'                         & "<Ctrl+P> Certificado de Promocion, <F5> Alumnos Matriculados "
  Llenar_Catalogo_Estudiantil
  Set clmX = LstVAlumnos.ColumnHeaders.Add(, , "Nombre del Alumno", LstVAlumnos.width)
  LstVAlumnos.BorderStyle = ccFixedSingle ' Establece la propiedad BorderStyle.
  LstVAlumnos.Icons = ImgLstHM 'ImgList
  LstVAlumnos.SmallIcons = ImgLstHM 'ImgList
  LstVAlumnos.View = lvwList
  With PictLibreta
      .AutoSize = True
      .AutoRedraw = True
      .ScaleMode = vbCentimeters
      .Left = 0
      .Top = 0
      .width = 20.5
      .Height = 29.5
  End With

 'Set PrinterView = PictLibreta
  PictLibreta.Cls
  LstVAlumnos.Height = MDI_Y_Max - LstVAlumnos.Top - 40 '3650

  DGNotasLibreta.width = SSTab1.width - 200
  DGNotasLibreta.Height = SSTab1.Height - 600
  
  SSTab1.Tab = 0
  Picture1.width = SSTab1.width - 600
  Picture1.Height = SSTab1.Height - 800
  VScroll1.Height = Picture1.Height - 1000
  HScroll1.width = Picture1.width - HScroll1.Left
  HScroll1.Top = Picture1.Height + Picture1.Top + 10
  LblA4.Top = Picture1.Height + Picture1.Top + 10
  
  TxtCodigo.Top = Picture1.Height + Picture1.Top + 10
  Command2.Top = Picture1.Height + Picture1.Top + 10
  
 'LstVAlumnos.width = FLibretas.width - LstVAlumnos.Left - 300
  VScroll1.value = 0
  HScroll1.value = 0
  VScroll1.Max = PictLibreta.Height  '- Picture1.Height
  HScroll1.Max = PictLibreta.width  '- Picture1.Width
  VScroll1_Scroll
  HScroll1_Scroll
  
'  Pos_Pict_X = PictLibreta.Left
'  Pos_Pict_Y = PictLibreta.Top
  
  Directiva = Leer_Archivo_Texto(RutaSistema & "\DOCUMENT\Directiva.txt")
  Directiva = Replace(Directiva, vbCrLf, "")
  DirectorRegional = Leer_Archivo_Texto(RutaSistema & "\DOCUMENT\DirectorRegional.txt")
  DirectorRegional = Replace(DirectorRegional, vbCrLf, "")
  LstPeriodos.Text = LstPeriodos.List(0)
  Opcion = 0
  LstPeriodos.SetFocus
End Sub

Private Sub Form_Load()
  ConectarAdodc AdoAux
  ConectarAdodc AdoCursos
  ConectarAdodc AdoDetalle
  ConectarAdodc AdoAlumnos
  ConectarAdodc AdoLibreta
  ConectarAdodc AdoPlantel
  ConectarAdodc AdoLectivo
  ConectarAdodc AdoMaterias
  ConectarAdodc AdoMatriculas
  ConectarAdodc AdoNotasLibreta
  ConectarAdodc AdoEquivalencia
End Sub

Private Sub LstPeriodos_DblClick()
  SiguienteControl
End Sub

Private Sub LstPeriodos_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub LstPeriodos_LostFocus()
  OpcionNotas = Seleccionar_Periodo(LstPeriodos)
 'MsgBox CadenaParcial
End Sub

Private Sub LstVAlumnos_DblClick()
  SiguienteControl
End Sub

Private Sub LstVAlumnos_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  CodigoCliente = Ninguno
  CodigoL = TVNivel.SelectedItem.key
  CodigoL = Mid$(CodigoL, 2, Len(CodigoL))
  If LstVAlumnos.ListItems.Count > 0 Then
     CodigoCliente = LstVAlumnos.SelectedItem.key
     CodigoCliente = Mid$(CodigoCliente, 2, Len(CodigoCliente))
  End If
  If CtrlDown And KeyCode = vbKeyP Then Imprimir_Libretas True
  If CtrlDown And KeyCode = vbKeyG Then Listar_Solicitud_Examen_Grado PictLibreta, CodigoCliente, 1
  If CtrlDown And KeyCode = vbKeyA Then Listar_Solicitud_Examen_Grado PictLibreta, CodigoCliente, 2
  If CtrlDown And KeyCode = vbKeyE Then Listar_Solicitud_Examen_Grado PictLibreta, CodigoCliente, 3
  If CtrlDown And KeyCode = vbKeyR Then Informe_Del_Alumno PictLibreta, CodigoCliente
 'If CtrlDown And KeyCode = vbKeyA Then Listar_Acta_Grado CodigoCliente
  PresionoEnter KeyCode
End Sub

Private Sub LstVAlumnos_LostFocus()
  Pagina = 1
  CodigoCliente = Ninguno
  CodigoL = TVNivel.SelectedItem.key
  CodigoL = Mid$(CodigoL, 2, Len(CodigoL))
  If LstVAlumnos.ListItems.Count > 0 Then
    'MsgBox CodigoL & vbCrLf & LstVAlumnos.SelectedItem
     CodigoCliente = LstVAlumnos.SelectedItem.key
     NombreCliente = LstVAlumnos.SelectedItem
     CodigoCliente = Mid$(CodigoCliente, 2, Len(CodigoCliente))
     TxtCodigo = CodigoCliente
     Leer_Datos_del_Curso CodigoL
     If OpcionNotas = 5 Then
        Aptitud_Promocion PictLibreta, False, CodigoL, CodigoCliente
     Else
        Libreta_Del_Alumno PictLibreta, CodigoL, CodigoCliente
     End If
  End If
End Sub

Private Sub MBFecha_GotFocus()
  MarcarTexto MBFecha
End Sub

Private Sub MBFecha_LostFocus()
  FechaValida MBFecha
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
Dim AnchoPict As Single
Dim AltoPict As Single
'''On Error GoTo Errorhandler
  RatonNormal
  Leer_Datos_del_Curso CodigoL
'''  CodigoCliente = Ninguno
'''  CodigoL = Ninguno
'''  If TVNivel.Nodes.Count > 0 Then
'''     TVNivel.SetFocus
'''     CodigoL = TVNivel.SelectedItem.key
'''     CodigoL = Mid$(CodigoL, 2, Len(CodigoL))
'''  End If
'''  If LstVAlumnos.ListItems.Count > 0 Then
'''     CodigoCliente = LstVAlumnos.SelectedItem.key
'''     CodigoCliente = Mid$(CodigoCliente, 2, Len(CodigoCliente))
'''  End If
 'MsgBox Button.key & vbCrLf & vbCrLf & TipoCta
  Select Case Button.key
    Case "Salir"
         Unload Me
    Case "Imprimir"
         Imprimir_Pagina
    Case "Matriculado"
         Alumnos_Matriculados CodigoL
    Case "ListadoAlumMatDP"
         Alumnos_Matriculados_Seccion CodigoL
    Case "Representante"
         Listar_Representantes CodigoL
    Case "Actas"
         Mensajes = "Actas Individual"
         Titulo = "IMPRESION DE ACTAS"
         If BoxMensaje = vbYes Then
            Listar_Acta_Grado CodigoCliente, True
         Else
            Imprimir_Actas CodigoL
         End If
         'If BoxMensaje = vbYes Then Listar_Actas True Else Listar_Actas False
    Case "SegundaPg"
         Mensajes = "Actas Individual"
         Titulo = "IMPRESION DE ACTAS"
         If BoxMensaje = vbYes Then
            Listar_Acta_Grado_Pag2 CodigoCliente, True
         Else
            Imprimir_Actas_Pag2 CodigoL
         End If
    Case "Libretas"
''         Mensajes = "Libreta Individual"
''         Titulo = "IMPRESION DE LIBRETAS"
''         If BoxMensaje = vbYes Then
''            TxtCodigo = "(" & Pagina & ") " & CodigoCliente
''            Libreta_Del_Alumno CodigoL, CodigoCliente
''         Else
            Imprimir_Libretas
            Imprimir_Informes
''         End If
    Case "Promocion"
         Mensajes = "Promocion Individual"
         Titulo = "IMPRESION DE PROMOCION"
         If BoxMensaje = vbYes Then
            Aptitud_Promocion PictLibreta, True, CodigoL, CodigoCliente
         Else
            Imprimir_Aptitudes_Promociones True
         End If
    Case "Aptitud"
        'MsgBox Button.key
         Mensajes = "Promocion Individual"
         Titulo = "IMPRESION DE APTITUD"
         If BoxMensaje = vbYes Then
            Aptitud_Promocion PictLibreta, False, CodigoL, CodigoCliente
         Else
            Imprimir_Aptitudes_Promociones False
         End If
    Case "Carnet"
         Imprimir_Carnet
    Case "CuadroProm"
         PDF_Procesar_Aprovechamiento Dato_Curso.Curso
         Mensajes = "Generar Reporte del Cuadro" & vbCrLf _
                  & "de Notas en Excel"
         Titulo = "Pregunta de Confirmación"
         If BoxMensaje = vbYes Then Procesar_Aprovechamiento_Excel CodigoL
    Case "Certificados"
         Mensajes = "Certificado Individual"
         Titulo = "IMPRESION DE CERTIFICADOS"
         If BoxMensaje = vbYes Then Certificado_Matricula Else Imprimir_Certificado_Matriculas
    Case "SolExaGrado"
         Mensajes = "Solicitud de Examen de Grado Individual"
         Titulo = "IMPRESION DE SOLICITUD"
         If BoxMensaje = vbYes Then
            Listar_Solicitud_Examen_Grado PictLibreta, CodigoCliente, 1
         Else
            Imprimir_Solicitudes_Examenes_Grado False, 1
         End If
    Case "AprobarExamGrado"
         Mensajes = "Aprovación de Examen de Grado Individual"
         Titulo = "IMPRESION DE APROBACION"
         If BoxMensaje = vbYes Then
            Listar_Solicitud_Examen_Grado PictLibreta, CodigoCliente, 2
         Else
            Imprimir_Solicitudes_Examenes_Grado False, 2
         End If
    Case "NotasExamGrado"
         Mensajes = "Notas de Examen de Grado Individual"
         Titulo = "IMPRESION DE NOTAS DE EXAMEN DE GRADO"
         If BoxMensaje = vbYes Then
            Listar_Solicitud_Examen_Grado PictLibreta, CodigoCliente, 3
         Else
            Imprimir_Solicitudes_Examenes_Grado False, 3
         End If
    Case "Supletorio"
         PDF_Procesar_Aprovechamiento CodigoL, True
    Case "Recalcular"
         'If Mid$(FormatoLibreta, 1, 9) = "QUIMESTRE" Then Recalcular_Notas_CodMatP Else
         Recalcular_Notas
    Case "Lista_Estudiantes"
         PDF_Lista_Estudiantes
    Case "Nomina_Representante"
         PDF_Nomina_Representante CodigoL
    Case "Nomina_Representante_Email"
         PDF_Nomina_Representante_Email CodigoL
  End Select
'''
'''Errorhandler:
'''             PictLibreta.Visible = True
'''             RatonNormal
'''             ErrorDeImpresion
'''             Exit Sub
End Sub

Private Sub TVNivel_DblClick()
  SiguienteControl
End Sub

Private Sub TVNivel_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TVNivel_LostFocus()
  Label6.Caption = TVNivel.SelectedItem
  CodigoL = TVNivel.SelectedItem.key
  CodigoL = Mid$(CodigoL, 2, Len(CodigoL))
  With AdoPlantel.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("CodigoE like '" & CodigoL & "' ")
      'MsgBox CodigoL
       If Not .EOF Then
          Codigo = .Fields("CodigoE")
          Codigo4 = Mid$(Codigo, 1, 4)
          TipoDoc = .Fields("CodMat")
          TipoCta = .Fields("TC")
          LblDirigente.Caption = " " & .Fields("Dirigente")
       End If
   End If
  End With
  Listar_Alumnos_Curso CodigoL
End Sub

Public Sub Imprimir_Total_Libreta(PosicionTotal As Single, _
                                  AdoConducta As Adodc, _
                                  Optional Recomendacion As Boolean, _
                                  Optional Escalas As Boolean)
  If TotalReg = 0 Then TotalReg = 1
  Printer.FontName = TipoArial
  PrinterFontBold True
  PrinterFontSize 8
  If TotalRegs(1) = 0 Then TotalRegs(1) = 1
  If TotalRegs(2) = 0 Then TotalRegs(2) = 1
  If TotalRegs(4) = 0 Then TotalRegs(4) = 1
  If TotalRegs(5) = 0 Then TotalRegs(5) = 1
  If TotalRegs(6) = 0 Then TotalRegs(6) = 1
  If TotalRegs(8) = 0 Then TotalRegs(8) = 1
  If TotalRegs(10) = 0 Then TotalRegs(10) = 1
  PosLinea = PosicionTotal + 0.05
  Imprimir_Linea_H PosLinea, 0.8, 19.3
  PosLinea = PosLinea + 0.1
  If Escalas Then PrinterPaint RutaSistema & "\FORMATOS\ESVANOTA.GIF", 0.8, PosLinea, 3, 2
  PrinterTexto 4.5, PosLinea, "SUMATORIA"
  IR = PosColumna
  If VPQBim1 > 0 Then PrinterTexto IR, PosLinea, Format(VPQBim1, "00")
  IR = IR + JR
  If VPQBim2 > 0 Then PrinterTexto IR, PosLinea, Format(VPQBim2, "00")
  IR = IR + (JR * 3)
  If VPromPQ > 0 Then PrinterTexto IR, PosLinea, Format(VPromPQ, "00")
  IR = IR + JR
  If VSQBim1 > 0 Then PrinterTexto IR, PosLinea, Format(VSQBim1, "00")
  IR = IR + JR
  If VSQBim2 > 0 Then PrinterTexto IR, PosLinea, Format(VSQBim2, "00")
  IR = IR + (JR * 3)
  If VPromSQ > 0 Then PrinterTexto IR, PosLinea, Format(VPromSQ, "00")
  If ("2.00" <= Codigo4) And (Codigo4 <= "5.99") Then
     IR = IR + (JR * 2) + 0.2
  Else
     IR = IR + JR + 0.2
  End If
  IR = IR - 0.2
  If VPromFinal > 0 Then PrinterTexto IR, PosLinea, Format(VPromFinal, "00.00")
  PosLinea = PosLinea + 0.35
  PrinterTexto 4.5, PosLinea, "APROVECHAMIENTO"
  IR = PosColumna
  If VPQBim1 > 0 Then PrinterTexto IR, PosLinea, Format(VPQBim1 / TotalRegs(1), "00.00")
  IR = IR + JR
  If VPQBim2 > 0 Then PrinterTexto IR, PosLinea, Format(VPQBim2 / TotalRegs(2), "00.00")
  IR = IR + (JR * 3)
  If VPromPQ > 0 Then PrinterTexto IR, PosLinea, Format(VPromPQ / TotalRegs(4), "00.00")
  IR = IR + JR
  If VSQBim1 > 0 Then PrinterTexto IR, PosLinea, Format(VSQBim1 / TotalRegs(5), "00.00")
  IR = IR + JR
  If VSQBim2 > 0 Then PrinterTexto IR, PosLinea, Format(VSQBim2 / TotalRegs(6), "00.00")
  IR = IR + (JR * 3)
  If VPromSQ > 0 Then PrinterTexto IR, PosLinea, Format(VPromSQ / TotalRegs(8), "00.00")
  If ("2.00" <= Codigo4) And (Codigo4 <= "5.99") Then
     IR = IR + (JR * 2) + 0.2
  Else
     IR = IR + JR + 0.2
  End If
  IR = IR - 0.2
  If VPromFinal > 0 And VPromPQ > 0 And VPromSQ > 0 Then PrinterTexto IR, PosLinea, Format(VPromFinal / TotalRegs(10), "00.00")
  With AdoConducta.Recordset
   If .RecordCount > 0 Then
       If .Fields("Orden") = 9 Then
           PosLinea = PosLinea + 0.4
           PrinterFontBold True
           PrinterTexto 4.5, PosLinea, .Fields("Materia")
           IR = PosColumna
           PrinterTextoNegrilla IR, PosLinea, .Fields("PQBim1"), , .Fields("C")
           IR = IR + JR
           PrinterTextoNegrilla IR, PosLinea, .Fields("PQBim2"), , .Fields("C")
           IR = IR + JR
           Sumatoria = .Fields("PQBim1") + .Fields("PQBim2")
           PrinterTextoNegrilla IR, PosLinea, Sumatoria
           IR = IR + JR
           
           IR = IR + JR
           PrinterTextoNegrilla IR, PosLinea, .Fields("PromPQ"), , .Fields("C")
           IR = IR + JR
           PrinterTextoNegrilla IR, PosLinea, .Fields("SQBim1"), , .Fields("C")
           IR = IR + JR
           PrinterTextoNegrilla IR, PosLinea, .Fields("SQBim2"), , .Fields("C")
           IR = IR + JR
           Sumatoria = .Fields("SQBim1") + .Fields("SQBim2")
           PrinterTextoNegrilla IR, PosLinea, Sumatoria
           IR = IR + JR
           
           IR = IR + JR
           PrinterTextoNegrilla IR, PosLinea, .Fields("PromSQ")
           If ("2.00" <= Codigo4) And (Codigo4 <= "5.99") Then
              IR = IR + JR
              If .Fields("Supletorio") > 0 Then PrinterTexto IR, PosLinea, Format(.Fields("Supletorio"), "00")
           End If
           IR = IR + JR + 0.2
           IR = IR - 0.2
           If .Fields("PromFinal") > 0 And .Fields("PromPQ") > 0 And .Fields("PromSQ") > 0 Then
               PrinterTextoNegrilla IR, PosLinea, .Fields("PromFinal"), 2, .Fields("C")
           End If
           PrinterFontBold False
       End If
   End If
  End With
 'Imprimimos el Resto de la libreta Recomendaciones
  If Recomendacion Then
     PosLinea = PosLinea + 1
     If AdoAux.Recordset.RecordCount > 0 Then
        AdoAux.Recordset.MoveFirst
        AdoAux.Recordset.Find ("Codigo = '" & Codigo & "' ")
        If Not AdoAux.Recordset.EOF Then
           Printer.FontBold = True
           PrinterTexto 1.5, PosLinea, "FALTAS JUSTIFICADAS:"
           PrinterTexto 7, PosLinea, "PERIODOS DE CLASE"
           Printer.FontBold = False
           If AdoAux.Recordset.Fields("PQBFJ1") > 0 Then
              PrinterVariables 6, PosLinea, AdoAux.Recordset.Fields("PQBFJ1")
           End If
           PosLinea = PosLinea + 0.5
           Printer.FontBold = True
           PrinterTexto 1.5, PosLinea, "FALTAS INJUSTIFICADAS:"
           PrinterTexto 7, PosLinea, "PERIODOS DE CLASE"
           Printer.FontBold = False
           If AdoAux.Recordset.Fields("PQBFI1") > 0 Then
              PrinterVariables 6, PosLinea, AdoAux.Recordset.Fields("PQBFI1")
           End If
           PosLinea = PosLinea + 0.5
           Printer.FontBold = True
           PrinterTexto 1.5, PosLinea, "ATRASOS:"
           PrinterTexto 7, PosLinea, "PERIODOS DE CLASE"
           Printer.FontBold = False
           If AdoAux.Recordset.Fields("PQBA1") > 0 Then
              PrinterVariables 6, PosLinea, AdoAux.Recordset.Fields("PQBA1")
           End If
           PosLinea = PosLinea + 0.5
        End If
     End If
     PrinterFontSize 10
     PosLinea = PosLinea + 1
     If Codigo4 >= "2" Then
        Printer.FontBold = True
        PrinterTexto 1.5, PosLinea, "OBSERVACIONES:"
        PosLinea = PosLinea + 0.5
        Printer.FontBold = False
        CodigoP = "Las asignaturas que no tienen nota, son terminales en un quimestre."
        PrinterTexto 2.5, PosLinea, CodigoP
        PosLinea = PosLinea + 1
     End If
     Printer.FontBold = True
     PrinterTexto 1.5, PosLinea, "RECOMENDACIONES:"
     PosLinea = PosLinea + 1
     Imprimir_Linea_H PosLinea, 2.5, 18
     PosLinea = PosLinea + 0.6
     Imprimir_Linea_H PosLinea, 2.5, 18
     PosLinea = PosLinea + 0.6
     Imprimir_Linea_H PosLinea, 2.5, 18
     PosLinea = PosLinea + 0.6
     Imprimir_Linea_H PosLinea, 2.5, 18
     PosLinea = PosLinea + 0.6
     Imprimir_Linea_H PosLinea, 2.5, 18
     PosLinea = PosLinea + 0.6
     Imprimir_Linea_H PosLinea, 2.5, 18
     PosLinea = PosLinea + 1
  End If
  PosLinea = PosLinea + 1.2
  Printer.FontSize = 9
  Printer.FontBold = True
  If ("2.00" <= Codigo4) And (Codigo4 <= "5.99") Then
     PrinterTexto 4, PosLinea, "REPRESENTANTE"
     PrinterTexto 10.5, PosLinea, "RECTORA"
     PrinterTexto 16, PosLinea, "SECRETARIA"
  Else
     PrinterTexto 4, PosLinea, "REPRESENTANTE"
     PrinterTexto 9, PosLinea, "RECTORA"
     PrinterTexto 13, PosLinea, "SECRETARIA"
     PrinterTexto 17, PosLinea, "PROFESOR(A)"
  End If
  Printer.FontBold = False
End Sub

Public Sub Imprimir_Libretas(Optional ImpIndividual As Boolean)
Dim AnchoPict As Single
Dim AltoPict As Single
On Error GoTo Errorhandler
Mensajes = "Seguro de Imprimir en:" & vbCrLf & Printer.DeviceName & "?"
Titulo = "IMPRESION DE LIBRETAS"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
   RatonReloj
   
   Pagina = 1
   InicioX = 0
   InicioY = 0
   PictLibreta.Visible = False
   Escala_Centimetro 1, TipoTimes, 9
   AnchoPict = Round(Printer.ScaleWidth, 5)
   AltoPict = Round(Printer.ScaleHeight, 5)
   If ImpIndividual Then
     TxtCodigo = CodigoCliente
     Leer_Datos_del_Curso CodigoL
     Libreta_Del_Alumno Printer, CodigoL, CodigoCliente
   Else
      CodigoL = TVNivel.SelectedItem.key
      CodigoL = Mid$(CodigoL, 2, Len(CodigoL))
      Leer_Datos_del_Curso CodigoL
      With AdoAlumnos.Recordset
       If .RecordCount > 0 Then
          .MoveFirst
           Do While Not .EOF
              CodigoCliente = .Fields("Codigo")
              TxtCodigo = "(" & Pagina & ") " & CodigoCliente
              Libreta_Del_Alumno Printer, CodigoL, CodigoCliente
             'Printer.PaintPicture PictLibreta.Image, InicioX, InicioY, AnchoPict, AltoPict
              Printer.NewPage
              Pagina = Pagina + 1
             .MoveNext
           Loop
       End If
      End With
   End If
   RatonNormal
   MensajeEncabData = ""
   Printer.EndDoc
   PictLibreta.Visible = True
   Exit Sub
Errorhandler:
             PictLibreta.Visible = True
             RatonNormal
             ErrorDeImpresion
             Exit Sub
Else
   RatonNormal
End If
End Sub

Public Sub Imprimir_Informes(Optional ImpIndividual As Boolean)
Dim AnchoPict As Single
Dim AltoPict As Single
On Error GoTo Errorhandler
Mensajes = "Seguro de Imprimir en:" & vbCrLf & Printer.DeviceName & "?"
Titulo = "IMPRESION DE LIBRETAS"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
   RatonReloj
   Pagina = 1
   InicioX = 0
   InicioY = 0
   PictLibreta.Visible = False
   Escala_Centimetro 1, TipoTimes, 9
   AnchoPict = Round(Printer.ScaleWidth, 5)
   AltoPict = Round(Printer.ScaleHeight, 5)
   If ImpIndividual Then
     TxtCodigo = CodigoCliente
     Informe_Del_Alumno Printer, CodigoCliente
   Else
      CodigoL = TVNivel.SelectedItem.key
      CodigoL = Mid$(CodigoL, 2, Len(CodigoL))
      Leer_Datos_del_Curso CodigoL
      With AdoAlumnos.Recordset
       If .RecordCount > 0 Then
          .MoveFirst
           Do While Not .EOF
              CodigoCliente = .Fields("Codigo")
              TxtCodigo = "(" & Pagina & ") " & CodigoCliente
              Informe_Del_Alumno Printer, CodigoCliente
              If Imp_Informe Then Printer.NewPage
              Pagina = Pagina + 1
             .MoveNext
           Loop
       End If
      End With
   End If
   RatonNormal
   MensajeEncabData = ""
   Printer.EndDoc
   PictLibreta.Visible = True
   Exit Sub
Errorhandler:
             PictLibreta.Visible = True
             RatonNormal
             ErrorDeImpresion
             Exit Sub
Else
   RatonNormal
End If
End Sub

Public Sub Imprimir_Aptitudes_Promociones(EsProm As Boolean)
Dim AnchoPict As Single
Dim AltoPict As Single
On Error GoTo Errorhandler
Mensajes = "Seguro de Imprimir en:" & vbCrLf & Printer.DeviceName & "?"
Titulo = "IMPRESION DE LIBRETAS"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
   RatonReloj
   Pagina = 1
   InicioX = 0
   InicioY = 0
  'PictLibreta.Visible = False
   Escala_Centimetro 1, TipoTimes, 9
   AnchoPict = Round(Printer.ScaleWidth, 5)
   AltoPict = Round(Printer.ScaleHeight, 5)
   CodigoL = TVNivel.SelectedItem.key
   CodigoL = Mid$(CodigoL, 2, Len(CodigoL))
   With AdoAlumnos.Recordset
    If .RecordCount > 0 Then
       .MoveFirst
        Do While Not .EOF
           CodigoCliente = .Fields("Codigo")
           TxtCodigo = "(" & Pagina & ") " & CodigoCliente
           Aptitud_Promocion Printer, EsProm, CodigoL, CodigoCliente
          'Printer.PaintPicture PictLibreta.Image, InicioX, InicioY, AnchoPict, AltoPict
           Printer.NewPage
           Pagina = Pagina + 1
          .MoveNext
        Loop
    End If
   End With
   RatonNormal
   MensajeEncabData = ""
   Printer.EndDoc
   'PictLibreta.Visible = True
   Exit Sub
Errorhandler:
             PictLibreta.Visible = True
             RatonNormal
             ErrorDeImpresion
             Exit Sub
Else
   RatonNormal
End If
End Sub

Public Sub Imprimir_Certificado_Matriculas()
Dim AnchoPict As Single
Dim AltoPict As Single
On Error GoTo Errorhandler
Mensajes = "Seguro de Imprimir en:" & vbCrLf & Printer.DeviceName & "?"
Titulo = "IMPRESION DE LIBRETAS"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
   RatonReloj
   Pagina = 1
   InicioX = 0
   InicioY = 0
   PictLibreta.Visible = False
   Escala_Centimetro 1, TipoTimes, 9
   AnchoPict = Round(Printer.ScaleWidth, 5)
   AltoPict = Round(Printer.ScaleHeight, 5)
   CodigoL = TVNivel.SelectedItem.key
   CodigoL = Mid$(CodigoL, 2, Len(CodigoL))
   With AdoAlumnos.Recordset
    If .RecordCount > 0 Then
       .MoveFirst
        Do While Not .EOF
           CodigoCliente = .Fields("Codigo")
           TxtCodigo = "(" & Pagina & ") " & CodigoCliente
           Certificado_Matricula
           Printer.PaintPicture PictLibreta.Image, InicioX, InicioY, AnchoPict, AltoPict
           Printer.NewPage
           Pagina = Pagina + 1
          .MoveNext
        Loop
    End If
   End With
   RatonNormal
   MensajeEncabData = ""
   Printer.EndDoc
   PictLibreta.Visible = True
   Exit Sub
Errorhandler:
             PictLibreta.Visible = True
             RatonNormal
             ErrorDeImpresion
             Exit Sub
Else
   RatonNormal
End If
End Sub

Public Sub Encabezado_Libretas(CodigoNivel As String, _
                               Curso As String, _
                               Alumno As String)
Dim Xo As Single
Dim Yo As Single
 'MsgBox CodigoNivel
 'Tipo de Libreta
  Printer.FontBold = True
  Select Case Mid$(CodigoNivel, 1, 4)
    Case "0.00" To "1.01" '
         PrinterPaint RutaSistema & "\FORMATOS\KINDER.GIF", 0.5, 2.5, 19.5, 13
         MensajeEncabData = "FICHA DE DESARROLLO DE DESTREZAS Y HABILIDADES"
         SQLMsg1 = "NIVEL:"
         Printer.FontBold = False
         Printer.FontSize = 9
         PrinterTexto 0.6, 3.35, Alumno
         PrinterTexto 0.6, 4.1, Curso
         PosColumna = 0
    Case "1.02" To "1.99"
         PrinterPaint RutaSistema & "\FORMATOS\PRIMARIA.BMP", 0.5, 3, 19, 2.7
         MensajeEncabData = "LIBRETA DE CALIFICACIONES"
         Printer.FontSize = 6
         PrinterTexto 0.6, 3.03, "ALUMNO(A):"
         PrinterTexto 0.6, 3.8, "CURSO:"
         Printer.FontBold = False
         Printer.FontSize = 9
         PrinterTexto 0.6, 3.35, Alumno
         PrinterTexto 0.6, 4.1, Curso
         PosColumna = 8.45
    Case "2.00" To "5.99"
         PrinterPaint RutaSistema & "\FORMATOS\SECUNDARIA.BMP", 0.5, 3, 19, 2.7
         MensajeEncabData = "LIBRETA DE CALIFICACIONES"
         Printer.FontSize = 6
         PrinterTexto 0.6, 3.03, "ALUMNO(A):"
         PrinterTexto 0.6, 3.8, "CURSO:"
         Printer.FontBold = False
         Printer.FontSize = 9
         PrinterTexto 0.6, 3.35, Alumno
         PrinterTexto 0.6, 4.1, Curso
         PosColumna = 8.45
  End Select
 'Encabezado Libreta
  Printer.FontBold = True
  PrinterPaint LogoTipo, 0.1, 0.5, 4.5, 2.25
  Printer.ForeColor = QBColor(Negro)
  Printer.FontName = TipoTimes
  Printer.FontSize = 15
  Xo = CentrarTextoEncab(Empresa, 1, 19)
  PrinterTexto Xo, 0.7, Empresa
  Printer.FontSize = 12
  Xo = CentrarTextoEncab(MensajeEncabData, 1, 19)
  PrinterTexto Xo, 1.6, MensajeEncabData
  SQLMsg2 = "AÑO LECTIVO " & Anio_Lectivo
  Xo = CentrarTextoEncab(SQLMsg2, 1, 19)
  Printer.FontSize = 10
  PrinterTexto Xo, 2.4, SQLMsg2
  PrinterTexto 0.5, 3.8, SQLMsg1
  'MsgBox ".."
  Printer.FontSize = 7
  Printer.FontBold = False
  PrinterTexto 15, 2.65, "FECHA DE ENTREGA " & MBFecha
End Sub

Public Function Procesar_Disciplinas(CodigoAlum As String, CodCurso As String) As Single
Dim TotalDisciplina As Single
Dim N1, N2 As Byte
  Dias_Laborados = 0
  Atrasos = 0
  Faltas_Just = 0
  Faltas_Injust = 0
  TotalDisciplina = 0
  sSQL = "SELECT * " _
       & "FROM Trans_Notas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND Codigo = '" & CodigoAlum & "' " _
       & "AND CodE = '" & CodCurso & "' " _
       & "AND CodMat IN ('998','999') "
  SelectAdodc AdoLibreta, sSQL
  With AdoLibreta.Recordset
   If .RecordCount > 0 Then
       Select Case FormatoLibreta
         Case "TRIMESTRE1": TotalDisciplina = Redondear((.Fields("PromPQ") + .Fields("PromSQ") + .Fields("PromTQ")) / 3)
         Case "TRIMESTRE2": TotalDisciplina = Redondear((.Fields("PromPQ") + .Fields("PromSQ") + .Fields("PromTQ")) / 3)
         Case "BIMESTRES":  TotalDisciplina = Redondear((.Fields("PromPQ") + .Fields("PromSQ")) / 2)
         Case "QUIMESTRE"
                         If OpcionNotas = 5 Then
                            If Periodo_Contable <= "31/12/2013" Then
                               TotalDisciplina = .Fields("PromFinal")
                            Else
                               TotalDisciplina = Redondear(.Fields("PromSQ"))
                            End If
                         Else
                            TotalDisciplina = Redondear((.Fields("PromPQ") + .Fields("PromSQ")) / 2)
                         End If
         Case Else: TotalDisciplina = 0
       End Select
   End If
  End With
  
  sSQL = "SELECT * " _
       & "FROM Trans_Asistencia " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND Codigo = '" & CodigoAlum & "' " _
       & "AND CodE = '" & CodCurso & "' "
  SelectAdodc AdoLibreta, sSQL
  With AdoLibreta.Recordset
   If .RecordCount > 0 Then
       Dias_Laborados = .Fields("PQBFJ1") + .Fields("PQBFJ2") + .Fields("SQBFJ1") + .Fields("SQBFJ2")
       Faltas_Just = .Fields("PQBFJ1") + .Fields("PQBFJ2") + .Fields("SQBFJ1") + .Fields("SQBFJ2")
       Faltas_Injust = .Fields("PQBFI1") + .Fields("PQBFI2") + .Fields("SQBFI1") + .Fields("SQBFI2")
       Atrasos = .Fields("PQBA1") + .Fields("PQBA2") + .Fields("SQBA1") + .Fields("SQBA2")
''       N1 = Redondear((.Fields("ConductaPQ1") + .Fields("ConductaPQ2")) / 2)
''       N2 = Redondear((.Fields("ConductaSQ1") + .Fields("ConductaSQ2")) / 2)
''       TotalDisciplina = Redondear((N1 + N2) / 2)
   End If
  End With
  Procesar_Disciplinas = TotalDisciplina
End Function

'''Public Sub Procesar_Aprovechamiento(ElCurso As String, Optional OpcSupletorio As Boolean, Optional OpcRemedial As Boolean)
'''Dim PathDibujo As String
'''Dim PosXPict As Single
'''Dim PosYPict As Single
'''Dim Disciplina As Single
'''Dim SumaPromX As Single
'''Dim SumaPromY As Single
'''Dim PromMat As Single
''''Dim CantMaterias As Byte
'''Dim IniMaterias As Byte
'''Dim FinMaterias As Byte
'''Dim CantMaxMaterias As Byte
'''Dim CantAlumnos As Byte
'''Dim NombFilePict As String
'''Dim SiguientePagina As Boolean
'''Dim Aprobado As Boolean
'''  SiguientePagina = True
'''  Pagina = 1
'''  SumaPromX = 0
'''  SumaPromY = 0
'''  CantAlumnos = 0
'''  CantMaxMaterias = 0
'''  Si_No = True
'''  Progreso_Barra.Incremento = 0
'''  Progreso_Barra.Valor_Maximo = 100
'''  Progreso_Barra.Mensaje_Box = "PROCESANDO APROVECHAMIENTO"
'''  Progreso_esperar
''' 'Nomina de Alumnos del Curso
'''  If Mid(Dato_Curso.Curso, 1, 1) <= "1" Then
'''     Dato_Curso.CantNotas = 5
'''     Dato_Curso.MateriasPorPagina = 7
'''  Else
'''     Dato_Curso.CantNotas = 8
'''     Dato_Curso.MateriasPorPagina = 4
'''  End If
'''  PosXPict = 7.5
'''  Pagina = 0
'''  IniMaterias = 1
'''  FinMaterias = Dato_Curso.MateriasPorPagina
'''  If FinMaterias > Dato_Curso.ContMat Then FinMaterias = Dato_Curso.ContMat
'''  Do While IniMaterias <= Dato_Curso.ContMat
'''     Dato_Curso.PosXMat(IniMaterias) = PosXPict
'''     PosXPict = PosXPict + (Dato_Curso.CantNotas * 0.7)
'''     If IniMaterias >= FinMaterias Then
'''        If SiguientePagina Then
'''           If Mid(Dato_Curso.Curso, 1, 1) <= "1" Then
'''              Dato_Curso.MateriasPorPagina = Dato_Curso.MateriasPorPagina + 4
'''           Else
'''              Dato_Curso.MateriasPorPagina = Dato_Curso.MateriasPorPagina + 1
'''           End If
'''           SiguientePagina = False
'''        End If
'''        FinMaterias = FinMaterias + Dato_Curso.MateriasPorPagina
'''        If FinMaterias > Dato_Curso.ContMat Then FinMaterias = Dato_Curso.ContMat
'''        PosXPict = 0.1   'Posicion en la siguiente pagina
'''        Pagina = Pagina + 1
'''     End If
'''     IniMaterias = IniMaterias + 1
'''  Loop
'''
'''  Cadena = "POSICIONES: Pg. " & Pagina & vbCrLf
'''  For I = 1 To Dato_Curso.ContMat
'''      Cadena = Cadena & "(" & I & ") = " & Dato_Curso.PosXMat(I) & vbCrLf
'''  Next I
''' 'MsgBox Cadena
'''
'''  CantAlumnos = Dato_Curso.ContAlumnos
'''  PosYPict = (CantAlumnos * 0.5) + 4    'Establesco el Alto del cuadro
'''
'''  PictLibreta.AutoSize = True
'''  PictLibreta.AutoRedraw = True
'''  PictLibreta.DrawWidth = 1
'''  PictLibreta.width = Round(AltoMaximo)
'''  PictLibreta.Height = Round(AnchoMaximo + 7) ' PosYPict + 5
''' 'If PictLibreta.Height < 31 Then PictLibreta.Height = 31
'''
'''  PictLibreta2.AutoSize = True
'''  PictLibreta2.AutoRedraw = True
'''  PictLibreta2.DrawWidth = 1
'''  PictLibreta2.width = Round(AltoMaximo)
'''  PictLibreta2.Height = Round(AnchoMaximo + 7) 'PosYPict + 5
''' 'If PictLibreta2.Height < 31 Then PictLibreta2.Height = 31
'''
'''  PictLibreta3.AutoSize = True
'''  PictLibreta3.AutoRedraw = True
'''  PictLibreta3.DrawWidth = 1
'''  PictLibreta3.width = Round(AltoMaximo)
'''  PictLibreta3.Height = Round(AnchoMaximo + 7) 'PosYPict + 5
'''
'''  VScroll1.Max = PictLibreta.Height
'''  HScroll1.Max = PictLibreta.width
'''  VScroll2.Max = PictLibreta2.Height
'''  HScroll2.Max = PictLibreta2.width
'''  VScroll3.Max = PictLibreta3.Height
'''  HScroll3.Max = PictLibreta3.width
'''
'''  sSQL = "SELECT TM.Materia,C.Sexo,C.Cliente As Alumno,CC.Descripcion As Curso,TM.C,TM.P,TM.I,TN.* " _
'''       & "FROM Trans_Notas As TN," _
'''       & "Catalogo_Materias As TM," _
'''       & "Catalogo_Cursos As CC," _
'''       & "Clientes As C," _
'''       & "Clientes_Matriculas As CM " _
'''       & "WHERE TN.Item = '" & NumEmpresa & "' " _
'''       & "AND TN.Periodo = '" & Periodo_Contable & "' " _
'''       & "AND TN.CodE = '" & ElCurso & "' " _
'''       & "AND TN.CodMatP = '" & Ninguno & "' " _
'''       & "AND TN.CodMat NOT IN ('997','998','999') "
''' If OpcSupletorio Then sSQL = sSQL & "AND PromFinal < " & Nota_Rojo & " "
''' If OpcRemedial Then sSQL = sSQL & "AND PromFinal < " & Nota_Rojo & " "
''' sSQL = sSQL _
'''       & "AND TN.Codigo = C.Codigo " _
'''       & "AND C.Codigo = CM.Codigo " _
'''       & "AND TN.CodE = CC.Curso " _
'''       & "AND TN.CodE = CM.Grupo_No " _
'''       & "AND TN.CodMat = TM.CodMat " _
'''       & "AND TN.Item = CC.Item " _
'''       & "AND TN.Item = TM.Item " _
'''       & "AND TN.Item = CM.Item " _
'''       & "AND TN.Periodo = CC.Periodo " _
'''       & "AND TN.Periodo = TM.Periodo " _
'''       & "AND TN.Periodo = CM.Periodo "
'''  If Alfabetico Then
'''     sSQL = sSQL & "ORDER BY C.Cliente,TN.Orden,TN.Id_No "
'''  Else
'''     sSQL = sSQL & "ORDER BY C.Sexo,C.Cliente,TN.Orden,TN.Id_No "
'''  End If
'''  SelectAdodc AdoAux, sSQL
'''  With AdoAux.Recordset
'''  'MsgBox .RecordCount
'''   If .RecordCount > 0 Then
'''      'MsgBox PictLibreta.Height
'''       Encabezado_Aprovechamiento1 OpcSupletorio
'''       Contador = 0
'''       PosLinea = 5.6
'''       Encabezado_Aprovechamiento2 PictLibreta, 1, OpcSupletorio
'''       PosLinea = 5.6
'''       If Pagina > 1 Then Encabezado_Aprovechamiento2 PictLibreta2, 2, OpcSupletorio
'''       PosLinea = 5.6
'''       If Pagina > 2 Then Encabezado_Aprovechamiento2 PictLibreta3, 3, OpcSupletorio
'''       Progreso_Barra.Valor_Maximo = .RecordCount + 1
'''      .MoveFirst
'''       J = 1
'''       K = 0
'''       PosLinea = 8.2
'''       PosXPict = 8.5
'''
'''       PosLinea = PosLinea + 0.05
'''       PictLibreta.FontBold = False
'''       PictLibreta2.FontBold = False
'''       PictLibreta3.FontBold = False
'''       CodigoCli = .Fields("Codigo")
'''       NombreCliente = .Fields("Alumno")
'''       Codigo = .Fields("CodMat")
'''       Si_No = CBool(.Fields("I"))
'''       Contador = Contador + 1
'''      'MsgBox NombreCliente
'''       PictLibreta.FontName = TipoHelvetica                 ' TipoArialNarrow/TipoCourier/TipoVerdana/TipoHelvetica
'''       PictLibreta.FontSize = 7
'''       PictLibreta2.FontName = TipoHelvetica
'''       PictLibreta2.FontSize = 7
'''       PictLibreta3.FontName = TipoHelvetica
'''       PictLibreta3.FontSize = 7
'''       PictLibreta.Line (0.1, PosLinea - 0.05)-(7.6, PosLinea - 0.05), QBColor(Negro)
'''       PictPrint_Texto PictLibreta, 0.2, PosLinea, Format(Contador, "00") & ".-"
'''       PictPrint_Texto PictLibreta, 0.65, PosLinea, NombreCliente
'''      'Determinamos a cuantos decimales se necesita para presentar el cuadro
'''       IR = PosXPict
'''       Aprobado = True
'''      .MoveFirst
'''       Do While Not .EOF
'''          PictLibreta.FontBold = False
'''          PictLibreta.FontUnderline = False
'''          PictLibreta2.FontBold = False
'''          PictLibreta2.FontUnderline = False
'''          Si_No = CBool(.Fields("I"))
'''          If Si_No Then
'''            If CodigoCli <> .Fields("Codigo") Then
'''              'Colocamos la Disciplina y el Promedio
'''               Disciplina = Procesar_Disciplinas(CodigoCli, ElCurso)
'''               PictPrint_Nota_Materia PictLibreta, 7, PosLinea, Disciplina, True, Dec_Nota
'''               Notas_Promedio_Aprovechamiento PictLibreta, 1, SumaPromX, K
'''               If Pagina > 1 Then Notas_Promedio_Aprovechamiento PictLibreta2, 2, SumaPromX, K
'''               If Pagina > 2 Then Notas_Promedio_Aprovechamiento PictLibreta3, 3, SumaPromX, K
'''
'''               SumaPromY = SumaPromY + SumaPromX
'''               If FormatoLibreta = "BIMESTRES" Then
'''                  If Aprobado Then
'''                     PictPrint_Texto PictLibreta2, PosXPict + 2.4, PosLinea, "Aprobado"
'''                  Else
'''                     PictPrint_Texto PictLibreta2, PosXPict + 2.4, PosLinea, "Reprobado"
'''                  End If
'''               Else
''''''                  PictPrint_Texto PictLibreta, PosXPict + 0.1, PosLinea, Format(SumaPromX, "00.000")
'''               End If
'''               SumaPromX = 0
'''               If FormatoLibreta = "BIMESTRES" Then
'''                  PictLibreta.Line (PosXPict + 2.3, PosLinea)-(PosXPict + 2.3, PosLinea + 0.45), QBColor(Negro)
'''                  PictLibreta.Line (PosXPict + 4.2, PosLinea)-(PosXPict + 4.2, PosLinea + 0.45), QBColor(Negro)
'''               End If
'''               If FormatoLibreta = "BIMESTRES" Then
'''                  PictLibreta.Line (0.1, PosLinea)-(PosXPict + 4.2, PosLinea), QBColor(Negro)
'''               Else
''''                  PictLibreta.Line (0.1, PosLinea)-(Dato_Curso.PosXMat(6) + 8, PosLinea), QBColor(Negro)
''' '                 PictLibreta2.Line (1, PosLinea)-(Dato_Curso.PosXMat(Dato_Curso.ContMat - 1) + 10, PosLinea), QBColor(Negro)
'''               End If
'''               'PosLinea = PosLinea + 0.05
'''               J = 1
'''               K = 0
'''               PosXPict = 7.5
'''               Contador = Contador + 1
'''               CodigoCli = .Fields("Codigo")
'''               NombreCliente = .Fields("Alumno")
'''               Si_No = CBool(.Fields("I"))
'''               PosLinea = PosLinea + 0.4
'''               PictLibreta.Line (0.1, PosLinea - 0.05)-(7.6, PosLinea - 0.05), QBColor(Negro)
'''               PictPrint_Texto PictLibreta, 0.2, PosLinea, Format(Contador, "00") & ".-"
'''               PictPrint_Texto PictLibreta, 0.65, PosLinea, NombreCliente
'''               Aprobado = True
'''            End If
'''          'Si_No = True
'''
'''             If .Fields("PromPQ") > 0 And .Fields("PromSQ") <= 0 Then
'''                 Abono_ME = .Fields("PromPQ")
'''             ElseIf .Fields("PromPQ") <= 0 And .Fields("PromSQ") > 0 Then
'''                 Abono_ME = .Fields("PromSQ")
'''             Else
'''                 If Mid$(FormatoLibreta, 1, 9) = "TRIMESTRE" Then
'''                    Abono_ME = (.Fields("PromPQ") + .Fields("PromSQ") + .Fields("PromTQ")) / 3
'''                 Else
'''                    Abono_ME = (.Fields("PromPQ") + .Fields("PromSQ")) / 2
'''                 End If
'''             End If
'''             Abono_ME = Redondear_2Dec(Abono_ME)
'''''                  If Abono_ME < Nota_Rojo Then
'''''                     PictLibreta.FontBold = True
'''''                     PictLibreta.FontUnderline = True
'''''                    'Aprobado = False
'''''                  End If
'''            'Empezamos a imprimir las notas
'''             If FormatoLibreta = "BIMESTRES" Then
'''                IR = CSng(Dato_Curso.PosXMat(J)) + 0.05
'''             Else
'''                If Dec_Nota > 0 Then IR = CSng(Dato_Curso.PosXMat(J)) + 0.05 Else IR = CSng(Dato_Curso.PosXMat(J)) + 0.2
'''             End If
'''             If .Fields("C") = False Then
'''                 If Mid$(FormatoLibreta, 1, 9) = "TRIMESTRE" Then
'''                    PromMat = .Fields("PromPQ") + .Fields("PromSQ") + .Fields("PromTQ")
'''                    If Dec_Nota > 0 Then
'''                       Cadena = Format(PromMat, "00." & String(Dec_Nota, "0"))
'''                    Else
'''                       Cadena = Format(PromMat, "00")
'''                    End If
'''                    If PromMat > 0 Then PictPrint_Texto PictLibreta, IR, PosLinea, Cadena
'''                 Else
'''''                         PromMat = .Fields("PromPQ") + .Fields("PromSQ")
'''''                         If PromMat > 0 Then PictPrint_Texto PictLibreta, IR, PosLinea, Format(PromMat, "00.00")
'''                 End If
'''             End If
'''''                  IR = IR + 1  '0.7
'''             SinImprimir = True
'''             'If Contador = 8 Then MsgBox Contador & vbCrLf & .Fields("CodMat")
'''             If Dato_Curso.CantNotas > 5 Then
'''                Select Case Indice_Materia(.Fields("CodMat"))
'''                  Case 1 To 4      ' Si es la priemr pagina
'''                       Notas_Materias_Aprovechamiento PictLibreta, IR, PosLinea, AdoAux
'''                  Case 5 To 9     ' Si es la segunda pagina
'''                       Notas_Materias_Aprovechamiento PictLibreta2, IR, PosLinea, AdoAux
'''                  Case Else  ' Si es la Tercera pagina
'''                       Notas_Materias_Aprovechamiento PictLibreta3, IR, PosLinea, AdoAux
'''                End Select
'''             Else
'''                Select Case Indice_Materia(.Fields("CodMat"))
'''                  Case 1 To 7      ' Si es la priemr pagina
'''                       Notas_Materias_Aprovechamiento PictLibreta, IR, PosLinea, AdoAux
'''                  Case Else    ' Si es la segunda pagina
'''                       Notas_Materias_Aprovechamiento PictLibreta2, IR, PosLinea, AdoAux
'''                End Select
'''             End If
'''            'IR = IR + 1
'''             If .Fields("PromFinal") > 0 And .Fields("C") = False Then
'''                 SumaPromX = SumaPromX + Redondear(.Fields("PromFinal"), Dec_Nota)
'''                 K = K + 1
'''             End If
'''             If Redondear(.Fields("PromFinal"), Dec_Nota) < Nota_Rojo Then Aprobado = False
'''            'PosXPict = PosXPict + 8
'''            'MsgBox J
'''             J = J + 1
'''          End If
'''          Progreso_esperar
'''         .MoveNext
'''       Loop
'''    'Colocamos la Disciplina y el Promedio
'''     Disciplina = Procesar_Disciplinas(CodigoCli, ElCurso)
'''     PictPrint_Nota_Materia PictLibreta, 7, PosLinea, Disciplina, True, Dec_Nota
'''     Notas_Promedio_Aprovechamiento PictLibreta, 1, SumaPromX, K
'''     If Pagina > 1 Then Notas_Promedio_Aprovechamiento PictLibreta2, 2, SumaPromX, K
'''     If Pagina > 2 Then Notas_Promedio_Aprovechamiento PictLibreta3, 3, SumaPromX, K
'''     SumaPromY = SumaPromY + SumaPromX
'''
'''     If FormatoLibreta = "BIMESTRES" Then
'''        If Aprobado Then
'''           PictPrint_Texto PictLibreta, PosXPict + 2.4, PosLinea, "Aprobado"
'''        Else
'''           PictPrint_Texto PictLibreta, PosXPict + 2.4, PosLinea, "Reprobado"
'''        End If
'''     Else
'''''                  PictPrint_Texto PictLibreta, PosXPict + 0.1, PosLinea, Format(SumaPromX, "00.000")
'''     End If
'''     SumaPromX = 0
'''     If FormatoLibreta = "BIMESTRES" Then
'''        PictLibreta.Line (PosXPict + 2.3, PosLinea)-(PosXPict + 2.3, PosLinea + 0.45), QBColor(Negro)
'''        PictLibreta.Line (PosXPict + 4.2, PosLinea)-(PosXPict + 4.2, PosLinea + 0.45), QBColor(Negro)
'''     End If
'''     PosLinea = PosLinea + 0.45
'''     If FormatoLibreta = "BIMESTRES" Then
'''        PictLibreta.Line (0.1, PosLinea)-(PosXPict + 4.2, PosLinea), QBColor(Negro)
'''     Else
''''        PictLibreta.Line (0.1, PosLinea)-(Dato_Curso.PosXMat(6) + 8, PosLinea), QBColor(Negro)
''''        PictLibreta2.Line (1, PosLinea)-(Dato_Curso.PosXMat(Dato_Curso.ContMat - 1) + 10, PosLinea), QBColor(Negro)
'''     End If
'''
'''     PosLinea = PosLinea + 0.05
'''   End If
'''  End With
'''
''' 'Cuadro de Firmas de Materias
'''  If FormatoLibreta = "BIMESTRES" Then UltimaLinea = PosLinea + 2 Else UltimaLinea = PosLinea
'''
'''  PictLibreta.Line (0.1, PrimeraLinea)-(0.1, UltimaLinea), QBColor(Negro)
'''  IR = 6.8
'''  PictLibreta.Line (IR, PrimeraLinea)-(IR, UltimaLinea), QBColor(Negro)
'''
'''
'''
'''
'''
'''
'''''  For I = 1 To 6
'''''      IR = Dato_Curso.PosXMat(I)
'''''      For J = 0 To 8
'''''          PictLibreta.Line (IR, PrimeraLinea)-(IR, UltimaLinea), QBColor(Negro)
'''''          IR = IR + 0.6
'''''      Next J
'''''  Next I
'''
'''''  For I = 7 To Dato_Curso.ContMat - 1
'''''      IR = Dato_Curso.PosXMat(I)
'''''      For J = 0 To 8
'''''          PictLibreta2.Line (IR, PrimeraLinea)-(IR, UltimaLinea), QBColor(Negro)
'''''          IR = IR + 0.6
'''''      Next J
'''''  Next I
'''''  PictLibreta2.Line (IR, PrimeraLinea)-(IR, UltimaLinea), QBColor(Negro)
'''''  IR = IR + 1
'''''  PictLibreta2.Line (IR, PrimeraLinea)-(IR, UltimaLinea), QBColor(Negro)
'''
'''  'PictLibreta.Line (0.05, PrimeraLinea)-(Dato_Curso.PosXMat(7) + 8, UltimaLinea), QBColor(Negro), B
'''
'''  If FormatoLibreta = "BIMESTRES" Then
'''     PictLibreta.Line (0.1, UltimaLinea)-(IR, UltimaLinea), QBColor(Negro)
'''     PictPrint_Texto PictLibreta, 3, PosLinea + 0.5, "F I R M A S"
'''  Else
'''
'''  End If
'''  PictPrint_Texto PictLibreta2, Dato_Curso.PosXMat(Dato_Curso.ContMat - 1), UltimaLinea + 0.05, FechaStrgCiudad(MBFecha)
'''  PosLinea = UltimaLinea + 2
'''  PCol = 2
'''    Select Case Codigo4
'''      Case "0.00" To "1.99"
'''           PictPrint_Texto PictLibreta, Dato_Curso.PosXMat(1), PosLinea, Director                 '31.8
'''           PictPrint_Texto PictLibreta2, Dato_Curso.PosXMat(Dato_Curso.ContMat), PosLinea, Secretario1
'''           PosLinea = PosLinea + 0.35
'''           PictPrint_Texto PictLibreta, Dato_Curso.PosXMat(1), PosLinea, TextoDirector        '31.5
'''           PictPrint_Texto PictLibreta2, Dato_Curso.PosXMat(Dato_Curso.ContMat), PosLinea, TextoSecretario1
'''      Case "2.00" To "3.99"
'''           PictPrint_Texto PictLibreta, Dato_Curso.PosXMat(1), PosLinea, Rector
'''           PictPrint_Texto PictLibreta2, Dato_Curso.PosXMat(Dato_Curso.ContMat), PosLinea, Secretario2
'''           PosLinea = PosLinea + 0.35
'''           PictPrint_Texto PictLibreta, Dato_Curso.PosXMat(1), PosLinea, TextoRector
'''           PictPrint_Texto PictLibreta2, Dato_Curso.PosXMat(Dato_Curso.ContMat), PosLinea, TextoSecretario2
'''      Case "4.00" To "5.99"
'''           PictPrint_Texto PictLibreta, Dato_Curso.PosXMat(1), PosLinea, Rector
'''           PictPrint_Texto PictLibreta2, Dato_Curso.PosXMat(Dato_Curso.ContMat), PosLinea, Secretario3
'''           PosLinea = PosLinea + 0.35
'''           PictPrint_Texto PictLibreta, Dato_Curso.PosXMat(1), PosLinea, TextoRector
'''           PictPrint_Texto PictLibreta2, Dato_Curso.PosXMat(Dato_Curso.ContMat), PosLinea, TextoSecretario2
'''    End Select
'''  Progreso_Barra.Incremento = Progreso_Barra.Valor_Maximo
'''  Progreso_esperar
'''  RatonNormal
'''  Pagina = 1
'''  Opcion = 2
'''  PictLibreta.SetFocus
'''End Sub

Public Sub PDF_Procesar_Aprovechamiento(ElCurso As String, _
                                        Optional OpcSupletorio As Boolean, _
                                        Optional OpcRemedial As Boolean)
Dim PathDibujo As String
Dim PosXPict As Single
Dim PosYPict As Single
Dim Disciplina As Single
Dim SumaPromX As Single
Dim SumaPromY As Single
Dim PromMat As Single
Dim IniMaterias As Byte
Dim FinMaterias As Byte
Dim CantMaxMaterias As Byte
Dim CantAlumnos As Byte
Dim NombFilePict As String
Dim SiguientePagina As Boolean
Dim Aprobado As Boolean
    
RatonReloj
Consultamos_Disciplinas AdoMatriculas, ElCurso

Set ObjPDF = New mjwPDF
ObjPDF.PDFTitle = "Cuadro de Calificacion Anual"
ObjPDF.PDFFileName = RutaSysBases & "\TEMP\" & Replace(ElCurso, ".", "-") & ".PDF"
ObjPDF.PDFLoadAfm = RutaSistema & "\FONTSPDF"
ObjPDF.PDFSetUnit = UNIT_CM
ObjPDF.PDFSetLayoutMode = LAYOUT_DEFAULT
ObjPDF.PDFFormatPage = FORMAT_A4
ObjPDF.PDFOrientation = ORIENT_PORTRAIT
ObjPDF.PDFView = True
ObjPDF.PDFBeginDoc
ObjPDF.PDFSetBookmark " "
ObjPDF.PDFSetFontName FONT_HELVETICA
  SiguientePagina = True
  SumaPromX = 0
  SumaPromY = 0
  CantAlumnos = 0
  CantMaxMaterias = 0
  Si_No = True
  Progreso_Barra.Incremento = 0
  Progreso_Barra.Valor_Maximo = 100
  Progreso_Barra.Mensaje_Box = "PROCESANDO APROVECHAMIENTO"
  Progreso_Esperar
  PosXPict = 6.2
 'Cambiamos la ultima posicion de la materia
  I = Dato_Curso.ContMat
  If Dato_Curso.CodMat(I) < "997" Then
     Cadena = Dato_Curso.CodMat(I)
     Dato_Curso.CodMat(I) = Dato_Curso.CodMat(I - 1)
     Dato_Curso.CodMat(I - 1) = Cadena
     Cadena = Dato_Curso.Materia(I)
     Dato_Curso.Materia(I - 1) = Dato_Curso.Materia(I)
     Dato_Curso.Materia(I) = Cadena
  End If
  
  Progreso_Esperar
 'Determinamos cuantas materias por pagina alcanzan
  If Mid(Dato_Curso.Curso, 1, 1) <= "1" Then
     Dato_Curso.CantNotas = 5
     Dato_Curso.MateriasPorPagina = 4
  Else
     Dato_Curso.CantNotas = 8
     Dato_Curso.MateriasPorPagina = 3
  End If
  
  For I = 1 To 6
      CodMatxPag(I) = 0
  Next I
  
  Pagina = 0
  IniMaterias = 1
  FinMaterias = Dato_Curso.MateriasPorPagina
  If FinMaterias > Dato_Curso.ContMat Then FinMaterias = Dato_Curso.ContMat
  CodMatxPag(Pagina + 1) = IniMaterias
  Do While IniMaterias <= Dato_Curso.ContMat
     Dato_Curso.PosXMat(IniMaterias) = PosXPict  ' Ancho de Cada Bloque de Materias
     PosXPict = PosXPict + (Dato_Curso.CantNotas * 0.6)  'Cantidad de notas por materia
     If IniMaterias >= FinMaterias Then
        If SiguientePagina Then
           If Mid(Dato_Curso.Curso, 1, 1) <= "1" Then
              Dato_Curso.MateriasPorPagina = Dato_Curso.MateriasPorPagina + 2
           Else
              Dato_Curso.MateriasPorPagina = Dato_Curso.MateriasPorPagina + 1
           End If
           SiguientePagina = False
        End If
        FinMaterias = FinMaterias + Dato_Curso.MateriasPorPagina
        If FinMaterias > Dato_Curso.ContMat Then FinMaterias = Dato_Curso.ContMat
        PosXPict = 1.5   'Posicion en la siguiente pagina
        Pagina = Pagina + 1
        If IniMaterias < Dato_Curso.ContMat Then CodMatxPag(Pagina + 1) = IniMaterias + 1
     End If
     IniMaterias = IniMaterias + 1
  Loop
    
  Progreso_Esperar
  Cadena = "POSICIONES: Pg. " & Pagina & vbCrLf
  For I = 1 To Dato_Curso.ContMat
      Cadena = Cadena & "(" & I & ") = " & Dato_Curso.PosXMat(I) & vbCrLf
  Next I
  
  J = 0
  For I = 1 To 6
      If CodMatxPag(I) <> 0 Then J = J + 1
  Next I
  ReDim ContMaxPagina(1 To J) As Integer
  ReDim ContMinPagina(1 To J) As Integer

  For I = 1 To UBound(ContMaxPagina)
      ContMinPagina(I) = CodMatxPag(I)
      ContMaxPagina(I) = CodMatxPag(I + 1) - 1
      Cadena = Cadena & ContMinPagina(I) & " - " & ContMaxPagina(I) & vbCrLf
  Next I
  ContMaxPagina(I - 1) = Dato_Curso.ContMat
  
  'MsgBox Cadena & vbCrLf & ContMaxPagina(I - 1)
  
  CantAlumnos = Dato_Curso.ContAlumnos
  sSQL = "SELECT TM.Materia,C.Sexo,C.Cliente As Alumno,CC.Descripcion As Curso,TM.C,TM.C2,TM.P,TM.I,TN.* " _
       & "FROM Trans_Notas As TN,Catalogo_Materias As TM,Catalogo_Cursos As CC,Clientes As C,Clientes_Matriculas As CM " _
       & "WHERE TN.Item = '" & NumEmpresa & "' " _
       & "AND TN.Periodo = '" & Periodo_Contable & "' " _
       & "AND TN.CodE = '" & ElCurso & "' " _
       & "AND TN.CodMatP = '" & Ninguno & "' " _
       & "AND TN.CodMat NOT IN ('997','998','999') "
  If OpcSupletorio Then sSQL = sSQL & "AND PromFinal < " & Nota_Rojo & " "
  If OpcRemedial Then sSQL = sSQL & "AND PromFinal < " & Nota_Rojo & " "
  sSQL = sSQL _
       & "AND TN.Codigo = C.Codigo " _
       & "AND C.Codigo = CM.Codigo " _
       & "AND TN.CodE = CC.Curso " _
       & "AND TN.CodE = CM.Grupo_No " _
       & "AND TN.CodMat = TM.CodMat " _
       & "AND TN.Item = CC.Item " _
       & "AND TN.Item = TM.Item " _
       & "AND TN.Item = CM.Item " _
       & "AND TN.Periodo = CC.Periodo " _
       & "AND TN.Periodo = TM.Periodo " _
       & "AND TN.Periodo = CM.Periodo "
  If Alfabetico Then
     sSQL = sSQL & "ORDER BY C.Cliente,TN.Orden,TN.Id_No "
  Else
     sSQL = sSQL & "ORDER BY C.Sexo,C.Cliente,TN.Orden,TN.Id_No "
  End If
  SelectAdodc AdoAux, sSQL
  
  ObjPDF.PDFSetFontName FONT_HELVETICA
  ObjPDF.PDFSetFontSize 5
  'ObjPDF.PDFSetFontStyle FONT_NORMAL, False
  With AdoAux.Recordset
  'MsgBox .RecordCount
   If .RecordCount > 0 Then
       Progreso_Barra.Valor_Maximo = Progreso_Barra.Valor_Maximo + (.RecordCount * UBound(ContMaxPagina))
       Progreso_Esperar
       For Pagina = 1 To UBound(ContMaxPagina)
           Progreso_Barra.Mensaje_Box = "Procesando Cuadro Anual: Pagina No. " & Pagina
           PictPrint_Cuadro_Linea ObjPDF, 1, 1, 1, 1, QBColor(Negro)
          .MoveFirst
           Contador = 0
           If Pagina = 1 Then PDF_Encabezado_Aprovechamiento ObjPDF
           PosLinea = 5.55
           PosYPict = PosLinea
           PDF_Encabezado_Aprovechamiento2 ObjPDF, Pagina
           PosLinea = 6.35
           PrimeraLinea = PosLinea
           J = 1
           K = 0
           SumaPromX = 0
           CodigoCli = .Fields("Codigo")
           NombreCliente = .Fields("Alumno")
           Codigo = .Fields("CodMat")
           Si_No = CBool(.Fields("I"))
           Contador = Contador + 1
'          If Pagina = UBound(ContMaxPagina) Then AnchoMaxMaterias = AnchoMaxMaterias - 0.4
'          If AnchoMaxMaterias < 1.1 Then AnchoMaxMaterias = 1.1
           PictPrint_Cuadro_Linea ObjPDF, 1, PosLinea, AnchoMaxMaterias, PosLinea, QBColor(Negro)
           PosLinea = PosLinea + 0.3
           ObjPDF.PDFSetFontSize 5
           If Pagina = 1 Then
              ObjPDF.PDFSetFontSize 5
              PictPrint_Texto ObjPDF, 1.1, PosLinea, Format(Contador, "00") & ".-"
              PictPrint_Texto ObjPDF, 1.4, PosLinea, NombreCliente
           Else
              PictPrint_Texto ObjPDF, 1.1, PosLinea, Format(Contador, "00")
           End If
           ObjPDF.PDFSetFontSize 6
          'Determinamos a cuantos decimales se necesita para presentar el cuadro
           IR = PosXPict
           Aprobado = True
          .MoveFirst
           Do While Not .EOF
              Si_No = CBool(.Fields("I"))
              If Si_No Then
                If CodigoCli <> .Fields("Codigo") Then
                  'Colocamos la Disciplina y el Promedio
                   If Pagina = 1 Then
                      With AdoMatriculas.Recordset
                       If .RecordCount > 0 Then
                          .MoveFirst
                          .Find ("Codigo = '" & CodigoCli & "' ")
                           If Not .EOF Then
                              Disciplina = Redondear((.Fields("ConductaPQ1") + .Fields("ConductaPQ2") + .Fields("ConductaPQ3")) / 3, Dec_Nota)
                              Select Case Anio_Lectivo
                                Case Is >= "2014 - 2015", Ninguno
                                     Disciplina = Redondear(.Fields("ConductaSQ3"), Dec_Nota)
                                Case Else
                                     Disciplina = Redondear((.Fields("ConductaSQ1") + .Fields("ConductaSQ2") + .Fields("ConductaSQ3")) / 3, Dec_Nota)
                              End Select
                           End If
                       End If
                      End With
                      'Disciplina = Procesar_Disciplinas(CodigoCli, ElCurso)
                      PictPrint_Nota_Materia ObjPDF, 5.7, PosLinea, Disciplina, True, Dec_Nota
                   End If
                   If Pagina = UBound(ContMaxPagina) Then
                      IR = PosLineaX
                     'MsgBox Pagina & vbCrLf & UBound(ContMaxPagina) & vbCrLf & SumaPromX & vbCrLf & PosLineaX
                      PictPrint_Variables ObjPDF, IR - 0.45, PosLinea - 0.05, CCur(SumaPromX), True, 1
                      'PictPrint_Nota_Materia ObjPDF, IR, PosLinea, SumaPromX, , 1     'IR
                      IR = IR + 0.8
                      If K <= 0 Then K = 1
                      SumaPromX = SumaPromX / K
                      SumaPromX = Redondear_2Dec(SumaPromX)
                      PictPrint_Nota_Materia ObjPDF, IR, PosLinea, SumaPromX, , 2
                      IR = IR + 0.7
                      PictPrint_Texto ObjPDF, IR, PosLinea, Equivalencia(CCur(SumaPromX))
                      'AnchoMaxMaterias = IR
                   End If
                   J = 1
                   K = 0
                   SumaPromX = 0
                   PosXPict = 7.5
                   Contador = Contador + 1
                   CodigoCli = .Fields("Codigo")
                   NombreCliente = .Fields("Alumno")
                   Si_No = CBool(.Fields("I"))
                   PosLinea = PosLinea + 0.05

                   PictPrint_Cuadro_Linea ObjPDF, 1, PosLinea, AnchoMaxMaterias, PosLinea, QBColor(Negro)
                   PosLinea = PosLinea + 0.3
                   If Pagina = 1 Then
                      ObjPDF.PDFSetFontSize 5
                      PictPrint_Texto ObjPDF, 1.1, PosLinea, Format(Contador, "00") & ".-"
                      PictPrint_Texto ObjPDF, 1.4, PosLinea, NombreCliente
                   Else
                      PictPrint_Texto ObjPDF, 1.1, PosLinea, Format(Contador, "00")
                   End If
                   ObjPDF.PDFSetFontSize 6
                   Aprobado = True
                End If
                
               'Empezamos a imprimir las notas
                If Dec_Nota > 0 Then IR = CSng(Dato_Curso.PosXMat(J)) + 0.05 Else IR = CSng(Dato_Curso.PosXMat(J)) + 0.2
                If Not .Fields("C") And Not .Fields("C2") Then
                    PromMat = .Fields("PromPQ") + .Fields("PromSQ")
                    'If PromMat > 0 Then PictPrint_Texto ObjPDF, IR, PosLinea, Format(PromMat, "00.00")
                End If
                SinImprimir = True
               'If Contador = 8 Then MsgBox Contador & vbCrLf & .Fields("CodMat")
                IniMaterias = Indice_Materia(.Fields("CodMat"))
                IR = Dato_Curso.PosXMat(IniMaterias)
                If ContMinPagina(Pagina) <= IniMaterias And IniMaterias <= ContMaxPagina(Pagina) Then
                   Notas_Materias_Aprovechamiento ObjPDF, IR, PosLinea, AdoAux
                End If
               'IR = IR + 1
                If .Fields("PromFinal") > 0 And Not .Fields("C") And Not .Fields("C2") Then
                    SumaPromX = SumaPromX + Redondear(.Fields("PromFinal"), Dec_Nota)
                    K = K + 1
                End If
                If Redondear(.Fields("PromFinal"), Dec_Nota) < Nota_Rojo Then Aprobado = False
               'PosXPict = PosXPict + 8
               'MsgBox J
                J = J + 1
              End If
              Progreso_Esperar
             .MoveNext
           Loop
           
          'Colocamos la Disciplina y el Promedio
           If Pagina = 1 Then
              With AdoMatriculas.Recordset
               If .RecordCount > 0 Then
                  .MoveFirst
                  .Find ("Codigo = '" & CodigoCli & "' ")
                   If Not .EOF Then
                      Disciplina = Redondear((.Fields("ConductaPQ1") + .Fields("ConductaPQ2") + .Fields("ConductaPQ3")) / 3, Dec_Nota)
                      Select Case Anio_Lectivo
                        Case Is >= "2014 - 2015", Ninguno
                             Disciplina = Redondear(.Fields("ConductaSQ3"), Dec_Nota)
                        Case Else
                             Disciplina = Redondear((.Fields("ConductaSQ1") + .Fields("ConductaSQ2") + .Fields("ConductaSQ3")) / 3, Dec_Nota)
                      End Select
                   End If
               End If
              End With
              'Disciplina = Procesar_Disciplinas(CodigoCli, ElCurso)
              PictPrint_Nota_Materia ObjPDF, 5.7, PosLinea, Disciplina, True, Dec_Nota
           End If
           If Pagina = UBound(ContMaxPagina) Then
              IR = PosLineaX
              PictPrint_Variables ObjPDF, IR - 0.45, PosLinea - 0.05, CCur(SumaPromX), True, 1
              IR = IR + 0.8
              If K <= 0 Then K = 1
              
              SumaPromX = SumaPromX / K
              SumaPromX = Redondear_2Dec(SumaPromX)
              PictPrint_Nota_Materia ObjPDF, IR, PosLinea, SumaPromX, , 2
              IR = IR + 0.7
              PictPrint_Texto ObjPDF, IR, PosLinea, Equivalencia(CCur(SumaPromX))
              'AnchoMaxMaterias = IR
           End If
           SumaPromY = SumaPromY + SumaPromX
           SumaPromX = 0
           PosLinea = PosLinea + 0.05
           UltimaLinea = PosLinea
           PictPrint_Cuadro_Linea ObjPDF, 1, PosLinea, AnchoMaxMaterias, PosLinea, QBColor(Negro)
           If Pagina = 1 Then PictPrint_Cuadro_Linea ObjPDF, 1, PrimeraLinea, 1, UltimaLinea, QBColor(Negro)
          'MsgBox IR & vbCrLf & AnchoMaxMaterias & vbCrLf & PosLineaX
             'Lineas de Promedios
              If Pagina = UBound(ContMaxPagina) Then
                 IR = PosLineaX - 0.1
                 PictPrint_Cuadro_Linea ObjPDF, IR, PrimeraLinea - 2.1, IR, UltimaLinea, QBColor(Negro)
                 IR = IR + 0.8
                 PictPrint_Cuadro_Linea ObjPDF, IR, PrimeraLinea - 2.1, IR, UltimaLinea, QBColor(Negro)
                 IR = IR + 0.7
                 PictPrint_Cuadro_Linea ObjPDF, IR, PrimeraLinea - 2.1, IR, UltimaLinea, QBColor(Negro)
                 IR = IR + 0.8
                 PictPrint_Cuadro_Linea ObjPDF, IR, PrimeraLinea - 2.1, IR, UltimaLinea, QBColor(Negro)
                 'Recuadro de totales
                 PictPrint_Cuadro_Linea ObjPDF, PosLineaX - 0.1, PrimeraLinea - 1.7, IR, PrimeraLinea + 0.4, QBColor(Negro), "B"
                 AnchoMaxMaterias = AnchoMaxMaterias - 2.3
              End If
              If Pagina = 1 Then IR = Dato_Curso.PosXMat(1) - 0.7 Else IR = 1.4
              Do While IR < AnchoMaxMaterias
                 PictPrint_Cuadro_Linea ObjPDF, IR, PrimeraLinea, IR, UltimaLinea, QBColor(Negro)
                 IR = IR + 0.6
              Loop
              PictPrint_Cuadro_Linea ObjPDF, AnchoMaxMaterias, PrimeraLinea, AnchoMaxMaterias, UltimaLinea, QBColor(Negro)
           
           PosLinea = UltimaLinea + 2
          'Cuadro de Firmas de Materias
           ObjPDF.PDFSetFontSize 8
           PCol = 2
           Select Case Pagina
             Case 1
                  PictPrint_Texto ObjPDF, 1, UltimaLinea + 0.6, FechaStrgCiudad(MBFecha)
                  ObjPDF.PDFSetFontSize 9
                  PosLinea = UltimaLinea + 2
                  Select Case Codigo4
                    Case "0.00" To "1.99"
                         PictPrint_Texto ObjPDF, Dato_Curso.PosXMat(1), PosLinea, Director
                         PosLinea = PosLinea + 0.35
                         PictPrint_Texto ObjPDF, Dato_Curso.PosXMat(1), PosLinea, TextoDirector
                    Case "2.00" To "3.99"
                         PictPrint_Texto ObjPDF, Dato_Curso.PosXMat(1), PosLinea, Rector
                         PosLinea = PosLinea + 0.35
                         PictPrint_Texto ObjPDF, Dato_Curso.PosXMat(1), PosLinea, TextoRector
                    Case "4.00" To "5.99"
                         PictPrint_Texto ObjPDF, Dato_Curso.PosXMat(1), PosLinea, Rector
                         PosLinea = PosLinea + 0.35
                         PictPrint_Texto ObjPDF, Dato_Curso.PosXMat(1), PosLinea, TextoRector
                  End Select
             Case 2
                  ObjPDF.PDFSetFontSize 9
                  PCol = 2
                  Select Case Codigo4
                    Case "0.00" To "1.99"
                         PictPrint_Texto ObjPDF, Dato_Curso.PosXMat(Dato_Curso.ContMat), PosLinea, Secretario1
                         PosLinea = PosLinea + 0.35
                         PictPrint_Texto ObjPDF, Dato_Curso.PosXMat(Dato_Curso.ContMat), PosLinea, TextoSecretario1
                    Case "2.00" To "3.99"
                         PictPrint_Texto ObjPDF, Dato_Curso.PosXMat(Dato_Curso.ContMat), PosLinea, Secretario2
                         PosLinea = PosLinea + 0.35
                         PictPrint_Texto ObjPDF, Dato_Curso.PosXMat(Dato_Curso.ContMat), PosLinea, TextoSecretario2
                    Case "4.00" To "5.99"
                         PictPrint_Texto ObjPDF, Dato_Curso.PosXMat(Dato_Curso.ContMat), PosLinea, Secretario3
                         PosLinea = PosLinea + 0.35
                         PictPrint_Texto ObjPDF, Dato_Curso.PosXMat(Dato_Curso.ContMat), PosLinea, TextoSecretario2
                  End Select
           End Select
           ObjPDF.PDFEndPage
           If Pagina < UBound(ContMaxPagina) Then ObjPDF.PDFNewPage
       Next Pagina
   End If
  End With
  Progreso_Barra.Incremento = Progreso_Barra.Valor_Maximo
  Progreso_Esperar
'Fin del PDF
ObjPDF.PDFEndPage
ObjPDF.PDFEndDoc
RatonNormal
End Sub

Public Sub Procesar_Aprovechamiento_Excel(ElCurso As String, Optional OpcSupletorio As Boolean, Optional OpcRemedial As Boolean)
Dim PathDibujo As String
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
Dim color(8) As Long
  RatonReloj
  color(0) = vbBlack
  color(1) = vbRed
  color(2) = vbGreen
  color(3) = vbYellow
  color(4) = vbBlue
  color(5) = vbMagenta
  color(6) = vbCyan
  color(7) = vbWhite

  Contador = 0
  Progreso_Iniciar
  Set apexcel = CreateObject("excel.application")
 'hace que excel se vea o no
  apexcel.Visible = False
 'agrega un nuevo libro
  apexcel.workbooks.Add
  apexcel.cells(1, 1).Select
  PathDibujo = RutaSistema & "\LOGOS\MINISEDU.JPG"
  apexcel.ActiveSheet.Pictures.Insert(PathDibujo).Select
  apexcel.cells(1, 10).Select
  PathDibujo = RutaSistema & "\LOGOS\ECUADOR.GIF"
  apexcel.ActiveSheet.Pictures.Insert(PathDibujo).Select
  
  SiguientePagina = True
  Pagina = 1
  SumaPromX = 0
  SumaPromY = 0
  CantAlumnos = 0
  CantMaterias = 0
  Progreso_Barra.Mensaje_Box = "PROCESANDO APROVECHAMIENTO POR EXCEL"
  Progreso_Esperar
 'Nomina de Alumnos del Curso
  sSQL = "SELECT TM.Materia,C.Sexo,C.Cliente As Alumno,CC.Descripcion As Curso,TM.C,TM.P,TM.I,TN.* " _
       & "FROM Trans_Notas As TN," _
       & "Catalogo_Materias As TM," _
       & "Catalogo_Cursos As CC," _
       & "Clientes As C," _
       & "Clientes_Matriculas As CM " _
       & "WHERE TN.Item = '" & NumEmpresa & "' " _
       & "AND TN.Periodo = '" & Periodo_Contable & "' " _
       & "AND TN.CodE = '" & ElCurso & "' " _
       & "AND TN.CodMatP = '" & Ninguno & "' " _
       & "AND TN.CodMat NOT IN ('997','998','999') " _
       & "AND TM.P <> " & Val(adFalse) & " "
 If OpcSupletorio Then sSQL = sSQL & "AND PromFinal < " & Nota_Rojo & " "
 If OpcRemedial Then sSQL = sSQL & "AND PromFinal < " & Nota_Rojo & " "
 sSQL = sSQL _
       & "AND TN.Codigo = C.Codigo " _
       & "AND C.Codigo = CM.Codigo " _
       & "AND TN.CodE = CC.Curso " _
       & "AND TN.CodE = CM.Grupo_No " _
       & "AND TN.CodMat = TM.CodMat " _
       & "AND TN.Item = CC.Item " _
       & "AND TN.Item = TM.Item " _
       & "AND TN.Item = CM.Item " _
       & "AND TN.Periodo = CC.Periodo " _
       & "AND TN.Periodo = TM.Periodo " _
       & "AND TN.Periodo = CM.Periodo "
  If Alfabetico Then
     sSQL = sSQL & "ORDER BY C.Cliente,TN.Orden,TN.Id_No "
  Else
     sSQL = sSQL & "ORDER BY C.Sexo,C.Cliente,TN.Orden,TN.Id_No "
  End If
  SelectAdodc AdoAux, sSQL
  With AdoAux.Recordset
   If .RecordCount > 0 Then
       CodigoCli = .Fields("Codigo")
       Evaluar = True
       Progreso_Barra.Valor_Maximo = Progreso_Barra.Valor_Maximo + .RecordCount
       Do While Not .EOF
          Progreso_Esperar
          If CodigoCli <> .Fields("Codigo") Then
             CantAlumnos = CantAlumnos + 1
             Evaluar = False
             CodigoCli = .Fields("Codigo")
          End If
          If Evaluar Then
             Si_No = True
             If .Fields("I") = False Then Si_No = False
             If .Fields("CodMat") = "998" Then Si_No = False
             If .Fields("CodMat") = "999" Then Si_No = False
             If Si_No Then CantMaterias = CantMaterias + 1
          End If
         .MoveNext
       Loop
       CantMaterias = CantMaterias + 1
       CantAlumnos = CantAlumnos + 1
       ReDim AnchoPict(CantMaterias + 4) As CtasAsiento
       For I = 0 To CantMaterias
           AnchoPict(I).Valor = 0
           AnchoPict(I).Cta = Ninguno
           AnchoPict(I).Detalle = Ninguno
       Next I
       J = 0
       PosXPict = 8.5
      .MoveFirst
       CodigoCli = .Fields("Codigo")
      'Determino el ancho del Grafico
       Do While Not .EOF
          If CodigoCli <> .Fields("Codigo") Then Exit Do
          Si_No = True
          If .Fields("I") = False Then Si_No = False
          If .Fields("CodMat") = "998" Then Si_No = False
          If .Fields("CodMat") = "999" Then Si_No = False
          If Si_No Then
             AnchoPict(J).Cta = .Fields("CodMat")
             AnchoPict(J).Detalle = .Fields("Materia")
             AnchoPict(J).Valor = PosXPict
             PosXPict = PosXPict + 2.9
             J = J + 1
          End If
         .MoveNext
       Loop
     ''If LogoTipo <> "" Then PictLibreta.PaintPicture LoadPicture(LogoTipo), 2, PosLinea, 5, 2.5
       apexcel.cells(2, 1).Font.Bold = True
       apexcel.cells(2, 1).Font.Size = 16
       apexcel.cells(2, 1).formula = "REPÚBLICA DEL ECUADOR"
       apexcel.cells(3, 1).Font.Bold = True
       apexcel.cells(3, 1).Font.Size = 16
       apexcel.cells(3, 1).formula = Institucion1 & " " & Institucion2
       apexcel.cells(4, 1).Font.Bold = True
       apexcel.cells(4, 1).Font.Size = 16
       apexcel.cells(4, 1).formula = "CUADRO FINAL DE CALIFICACIONES"
       apexcel.cells(5, 1).Font.Size = 16
       apexcel.cells(5, 1).Font.Bold = True
       apexcel.cells(5, 1).formula = "A Ñ O   L E C T I V O:  " & Anio_Lectivo
       apexcel.cells(6, 1).Font.Size = 16
       apexcel.cells(6, 1).Font.Bold = True
       apexcel.cells(6, 1).formula = "RÉGIMEN COSTA"
       
      'Datos del Curso
       apexcel.cells(7, 1).Font.Size = 10
       apexcel.cells(7, 1).Font.Bold = True
       apexcel.cells(7, 1).formula = "ZONA: " & Zona
       apexcel.cells(7, 6).Font.Size = 10
       apexcel.cells(7, 6).Font.Bold = True
       apexcel.cells(7, 6).formula = "DISTRITO N° " & Distrito
       apexcel.cells(7, 14).Font.Size = 10
       apexcel.cells(7, 14).Font.Bold = True
       apexcel.cells(7, 14).formula = "AMIE: " & Codigo_AMIE
       
       apexcel.cells(8, 1).Font.Size = 10
       apexcel.cells(8, 1).Font.Bold = True
       apexcel.cells(8, 1).formula = "AÑO/CURSO: " & Dato_Curso.Nombre_Largo
       apexcel.cells(8, 12).Font.Size = 10
       apexcel.cells(8, 12).Font.Bold = True
       apexcel.cells(8, 12).formula = "JORNADA: MATUTINA"
       apexcel.cells(8, 14).Font.Size = 10
       apexcel.cells(8, 14).Font.Bold = True
       apexcel.cells(8, 14).formula = "MODALIDAD: PRESENCIAL"
       
       apexcel.cells(9, 1).Font.Size = 10
       apexcel.cells(9, 1).Font.Bold = True
       apexcel.cells(9, 1).formula = "PARALELO: " & Dato_Curso.Paralelo
       apexcel.cells(10, 1).Font.Size = 10
       apexcel.cells(10, 1).Font.Bold = True
       apexcel.cells(10, 1).formula = "TIPO DE BACHILLERATO: " & Dato_Curso.Titulo
       apexcel.cells(11, 1).Font.Size = 10
       apexcel.cells(11, 1).Font.Bold = True
       apexcel.cells(11, 1).formula = "FIGURA PROFESIONAL: "
       
       apexcel.cells(16, 1).Font.Size = 14
       apexcel.cells(16, 1).Font.Bold = True
       apexcel.cells(16, 1).formula = "A P E L L I D O S   Y   N O M B R E S"
       apexcel.Columns(1).columnWidth = 50
       J = 2
       K = 0
       For I = 0 To CantMaterias
        If AnchoPict(I).Cta <> Ninguno Then
           Cadena1 = AnchoPict(I).Detalle
           Cadena = Trim(SinEspaciosIzq(Cadena1))
          'Insertamos la Materia
           apexcel.cells(13, J).Interior.color = color(K + 1)
           apexcel.cells(13, J + 1).Interior.color = color(K + 1)
           apexcel.cells(13, J + 2).Interior.color = color(K + 1)
           apexcel.cells(13, J + 3).Interior.color = color(K + 1)
           apexcel.cells(13, J + 4).Interior.color = color(K + 1)
           apexcel.cells(13, J + 5).Interior.color = color(K + 1)
           apexcel.cells(13, J + 6).Interior.color = color(K + 1)
           apexcel.cells(13, J + 6).Interior.color = color(K + 1)
           apexcel.cells(13, J).Font.color = color(K)
           apexcel.cells(13, J).Font.Name = TipoArialNarrow
           apexcel.cells(13, J).Font.Size = 8
           apexcel.cells(13, J).Font.Bold = True
           apexcel.cells(13, J).formula = Cadena
          'Contador Materias
           apexcel.cells(13, J + 5).formula = "'" & Format(I + 1, "00")
           
           Cadena1 = Trim(Mid$(Cadena1, Len(Cadena) + 1, Len(Cadena1)))
           Cadena = Trim(SinEspaciosIzq(Cadena1))
''           apexcel.Cells(14, J).Interior.Color = Color(K + 1)
''           apexcel.Cells(14, J + 1).Interior.Color = Color(K + 1)
''           apexcel.Cells(14, J + 2).Interior.Color = Color(K + 1)
''           apexcel.Cells(14, J + 3).Interior.Color = Color(K + 1)
''           apexcel.Cells(14, J + 4).Interior.Color = Color(K + 1)
''           apexcel.Cells(14, J + 5).Interior.Color = Color(K + 1)
           apexcel.cells(14, J).Font.color = color(K)
           apexcel.cells(14, J).Font.Name = TipoArialNarrow
           apexcel.cells(14, J).Font.Size = 8
           apexcel.cells(14, J).Font.Bold = True
           apexcel.cells(14, J).formula = Cadena
           
           Cadena1 = Trim(Mid$(Cadena1, Len(Cadena) + 1, Len(Cadena1)))
''           apexcel.Cells(15, J).Interior.Color = Color(K + 1)
''           apexcel.Cells(15, J + 1).Interior.Color = Color(K + 1)
''           apexcel.Cells(15, J + 2).Interior.Color = Color(K + 1)
''           apexcel.Cells(15, J + 3).Interior.Color = Color(K + 1)
''           apexcel.Cells(15, J + 4).Interior.Color = Color(K + 1)
''           apexcel.Cells(15, J + 5).Interior.Color = Color(K + 1)
           apexcel.cells(15, J).Font.color = color(K)
           apexcel.cells(15, J).Font.Name = TipoArialNarrow
           apexcel.cells(15, J).Font.Size = 8
           apexcel.cells(15, J).Font.Bold = True
           apexcel.cells(15, J).formula = Cadena1
           
          'Insertamos Encabezado de la Materia
           apexcel.cells(16, J).Font.Name = TipoArialNarrow
           apexcel.cells(16, J).Font.Size = 8
           apexcel.cells(16, J).Interior.color = color(4)
           apexcel.cells(16, J).Font.color = color(7)
           apexcel.cells(16, J).Orientation = 90
           apexcel.cells(16, J).formula = "PRIMER" & vbLf & "QUIMESTRE"
           apexcel.Columns(J).columnWidth = 5
           J = J + 1
           apexcel.cells(16, J).Font.Name = TipoArialNarrow
           apexcel.cells(16, J).Font.Size = 8
           apexcel.cells(16, J).Interior.color = color(4)
           apexcel.cells(16, J).Font.color = color(7)
           apexcel.cells(16, J).Orientation = 90
           apexcel.cells(16, J).formula = "SEGUNDO" & vbLf & "QUIMESTRE"
           apexcel.Columns(J).columnWidth = 5
           J = J + 1
           apexcel.cells(16, J).Font.Name = TipoArialNarrow
           apexcel.cells(16, J).Font.Size = 8
           apexcel.cells(16, J).Interior.color = color(5)
           apexcel.cells(16, J).Font.color = color(0)
           apexcel.cells(16, J).Orientation = 90
           apexcel.cells(16, J).formula = "PROMEDIO" & vbLf & "GLOBAL"
           apexcel.Columns(J).columnWidth = 5
           J = J + 1
           apexcel.cells(16, J).Font.Name = TipoArialNarrow
           apexcel.cells(16, J).Font.Size = 8
           apexcel.cells(16, J).Interior.color = color(6)
           apexcel.cells(16, J).Font.color = color(0)
           apexcel.cells(16, J).Orientation = 90
           apexcel.cells(16, J).formula = "EXAMEN" & vbLf & "SUPLETORIO"
           apexcel.Columns(J).columnWidth = 5
           J = J + 1
           apexcel.cells(16, J).Font.Name = TipoArialNarrow
           apexcel.cells(16, J).Font.Size = 8
           apexcel.cells(16, J).Interior.color = color(6)
           apexcel.cells(16, J).Font.color = color(0)
           apexcel.cells(16, J).Orientation = 90
           apexcel.cells(16, J).formula = "EXAMEN" & vbLf & "RECUPERACION"
           apexcel.Columns(J).columnWidth = 5
           J = J + 1
           apexcel.cells(16, J).Font.Name = TipoArialNarrow
           apexcel.cells(16, J).Font.Size = 8
           apexcel.cells(16, J).Interior.color = color(7)
           apexcel.cells(16, J).Font.color = color(0)
           apexcel.cells(16, J).Orientation = 90
           apexcel.cells(16, J).formula = "EXAMEN" & vbLf & "REMEDIAL"
           apexcel.Columns(J).columnWidth = 5
           J = J + 1
           apexcel.cells(16, J).Font.Name = TipoArialNarrow
           apexcel.cells(16, J).Font.Size = 8
           apexcel.cells(16, J).Interior.color = color(7)
           apexcel.cells(16, J).Font.color = color(0)
           apexcel.cells(16, J).Orientation = 90
           apexcel.cells(16, J).formula = "EXAMEN" & vbLf & "DE GRACIA"
           apexcel.Columns(J).columnWidth = 5
           J = J + 1
           apexcel.cells(16, J).Font.Name = TipoArialNarrow
           apexcel.cells(16, J).Font.Size = 8
           apexcel.cells(16, J).Interior.color = color(3)
           apexcel.cells(16, J).Font.color = color(0)
           apexcel.cells(16, J).Orientation = 90
           apexcel.cells(16, J).formula = "PROMEDIO" & vbLf & "TOTAL"
           
           apexcel.Columns(J).columnWidth = 5
           J = J + 1
           K = K + 1
           If K > 6 Then K = 0
        End If
       Next I
       apexcel.cells(16, J).Font.Name = TipoArialNarrow
       apexcel.cells(16, J).Font.Size = 7
       apexcel.cells(16, J).Font.color = color(0)
       apexcel.cells(16, J).Orientation = 90
       apexcel.cells(16, J).formula = "EVALUACION DEL" & vbCrLf & "COMPORTAMIENTO"
       apexcel.Columns(J).columnWidth = 5
       J = J + 1
       apexcel.cells(16, J).Font.Name = TipoArialNarrow
       apexcel.cells(16, J).Font.Size = 8
       apexcel.cells(16, J).Font.color = color(1)
       apexcel.cells(16, J).Orientation = 90
       apexcel.cells(16, J).formula = "PROMEDIO" & vbLf & "TOTAL"
       apexcel.Columns(J).columnWidth = 7
       J = J + 1
       apexcel.cells(16, J).Font.Name = TipoArialNarrow
       apexcel.cells(16, J).Font.Size = 7
       apexcel.cells(16, J).Font.color = color(4)
       apexcel.cells(16, J).formula = "OBSERVACION"
       apexcel.Columns(J).columnWidth = 10
       J = J + 1
       Progreso_Barra.Valor_Maximo = .RecordCount + 1
      .MoveFirst
       CodigoCli = .Fields("Codigo")
       NombreCliente = .Fields("Alumno")
       Codigo = .Fields("CodMat")
       Contador = 1
       apexcel.cells(16, J).Font.Name = TipoArialNarrow
       apexcel.cells(16, J).Font.Size = 8
       apexcel.cells(Contador + 16, 1).formula = Format(Contador, "00") & ".- " & NombreCliente
       Aprobado = True
       J = 2
       K = 0
       Do While Not .EOF
          Si_No = CBool(.Fields("I"))
          If Si_No Then
          If CodigoCli <> .Fields("Codigo") Then
            'Colocamos la Disciplina y el Promedio
             If K <= 0 Then K = 1
             'MsgBox SumaPromX
             SumaPromX = SumaPromX / K
             SumaPromY = SumaPromY + SumaPromX
             Disciplina = Procesar_Disciplinas(CodigoCli, ElCurso)
             apexcel.cells(Contador + 16, J).formula = Format(Disciplina, "00")
             J = J + 1
             apexcel.cells(Contador + 16, J).formula = Format(SumaPromX, "00.000")
             'MsgBox SumaPromX
             J = J + 1
             If SumaPromX >= Nota_Rojo Then
                apexcel.cells(Contador + 16, J).formula = "Aprobado"
             Else
                apexcel.cells(Contador + 16, J).Font.color = vbRed
                apexcel.cells(Contador + 16, J).formula = "Reprobado"
             End If
             SumaPromX = 0
             J = 2
             K = 0
             PosXPict = 8.5
             Contador = Contador + 1
             CodigoCli = .Fields("Codigo")
             NombreCliente = .Fields("Alumno")
             PictLibreta.FontBold = False
             apexcel.cells(Contador + 16, 1).formula = Format(Contador, "00") & ".- " & NombreCliente
             Aprobado = True
          End If
             Select Case .Fields("CodMat")
             Case "998", "999"
             
             Case Else
                If .Fields("PromPQ") > 0 And .Fields("PromSQ") <= 0 Then
                    Abono_ME = .Fields("PromPQ")
                    If Abono_ME <= 14 Then
                       'Aprobado = False
                    End If
                ElseIf .Fields("PromPQ") <= 0 And .Fields("PromSQ") > 0 Then
                    Abono_ME = .Fields("PromSQ")
                    If Abono_ME <= 14 Then
                       'Aprobado = False
                    End If
                Else
                    If Mid$(FormatoLibreta, 1, 9) = "TRIMESTRE" Then
                       Abono_ME = .Fields("PromPQ") + .Fields("PromSQ") + .Fields("PromTQ")
                       Abono_ME = Abono_ME / 3
                    Else
                       Abono_ME = .Fields("PromPQ") + .Fields("PromSQ")
                       Abono_ME = Abono_ME / 2
                    End If
                End If
                Abono_ME = Redondear_2Dec(Abono_ME)
                If .Fields("PromPQ") < Nota_Rojo Then
                   apexcel.cells(Contador + 16, J).Font.color = vbRed
                   apexcel.cells(Contador + 16, J).Font.Bold = True
                End If
                If .Fields("PromPQ") > 0 Then apexcel.cells(Contador + 16, J).formula = Format(.Fields("PromPQ"), "00.00")
                J = J + 1
                If .Fields("PromSQ") < Nota_Rojo Then
                   apexcel.cells(Contador + 16, J).Font.color = vbRed
                   apexcel.cells(Contador + 16, J).Font.Bold = True
                End If
                If .Fields("PromSQ") > 0 Then apexcel.cells(Contador + 16, J).formula = Format(.Fields("PromSQ"), "00.00")
                J = J + 1
                If Abono_ME < Nota_Rojo Then
                   apexcel.cells(Contador + 16, J).Font.color = vbRed
                   apexcel.cells(Contador + 16, J).Font.Bold = True
                End If
                If Abono_ME > 0 Then apexcel.cells(Contador + 16, J).formula = Format(Abono_ME, "00.00")
                J = J + 1
                If .Fields("Supletorio") < Nota_Rojo Then
                   apexcel.cells(Contador + 16, J).Font.color = vbRed
                   apexcel.cells(Contador + 16, J).Font.Bold = True
                End If
                J = J + 1
                'Recuperacion
                If .Fields("Supletorio") > 0 Then apexcel.cells(Contador + 16, J).formula = Format(.Fields("Supletorio"), "00.00")
                J = J + 1
                If .Fields("Remedial") < Nota_Rojo Then
                   apexcel.cells(Contador + 16, J).Font.color = vbRed
                   apexcel.cells(Contador + 16, J).Font.Bold = True
                End If
                If .Fields("Remedial") > 0 Then apexcel.cells(Contador + 16, J).formula = Format(.Fields("Remedial"), "00.00")
                J = J + 1
                
                If .Fields("PromFinal") < Nota_Rojo Then
                   apexcel.cells(Contador + 16, J).Font.color = vbRed
                   apexcel.cells(Contador + 16, J).Font.Bold = True
                End If
                J = J + 1
                'Examen de Gracia
                If .Fields("PromFinal") > 0 Then apexcel.cells(Contador + 16, J).formula = Format(.Fields("PromFinal"), "00.00")
                J = J + 1
                If .Fields("PromFinal") > 0 And .Fields("C") = False Then
                    SumaPromX = SumaPromX + .Fields("PromFinal")
                    K = K + 1
                End If
                If Redondear(.Fields("PromFinal")) < Nota_Rojo Then Aprobado = False
             End Select
          End If
          Progreso_Esperar
         .MoveNext
       Loop
       If K <= 0 Then K = 1
       SumaPromX = SumaPromX / K
       SumaPromY = SumaPromY + SumaPromX
       Disciplina = Procesar_Disciplinas(CodigoCli, ElCurso)
       apexcel.cells(Contador + 16, J).formula = Format(Disciplina, "00")
       J = J + 1
       apexcel.cells(Contador + 16, J).formula = Format(SumaPromX, "00.000")
       J = J + 1
       If SumaPromX >= 15 Then
          apexcel.cells(Contador + 16, J).formula = "Aprobado"
       Else
          apexcel.cells(Contador + 16, J).Font.color = vbRed
          apexcel.cells(Contador + 16, J).formula = "Reprobado"
       End If
       SumaPromX = 0
   End If
  End With
  Contador = Contador + 4
  PCol = 8.5
  Select Case Codigo4
    Case "0.00" To "1.99"
         apexcel.cells(Contador + 16, 2).formula = "DIRECTOR(A)"       '31.5
         apexcel.cells(Contador + 16, 16).formula = "SECRETARIO(A)"
         Contador = Contador + 1
         apexcel.cells(Contador + 16, 2).formula = Director            '31.8
         apexcel.cells(Contador + 16, 16).formula = Secretario1
    Case "2.00" To "3.99"
         apexcel.cells(Contador + 16, 2).formula = "RECTOR(A)"
         apexcel.cells(Contador + 16, 16).formula = "SECRETARIO(A)"
         apexcel.cells(Contador + 16, 30).formula = "LEGALIZADO SEGUN DECRETO EJECUTIVO"
         Contador = Contador + 1
         apexcel.cells(Contador + 16, 2).formula = Rector
         apexcel.cells(Contador + 16, 16).formula = Secretario2
         apexcel.cells(Contador + 16, 30).formula = "No. 1734 DEL 6/8/90 del Min. de Ecucación"
    Case "4.00" To "5.99"
         apexcel.cells(Contador + 16, 2).formula = "RECTOR(A)"
         apexcel.cells(Contador + 16, 16).formula = "SECRETARIO(A)"
         apexcel.cells(Contador + 16, 30).formula = "LEGALIZADO SEGUN DECRETO EJECUTIVO"
         Contador = Contador + 1
         apexcel.cells(Contador + 16, 2).formula = Rector
         apexcel.cells(Contador + 16, 16).formula = Secretario3
         apexcel.cells(Contador + 16, 30).formula = "No. 1734 DEL 6/8/90 del Min. de Ecucación"
  End Select
 'SavePicture PictLibreta.Image, RutaOrigen
  Progreso_Barra.Incremento = Progreso_Barra.Valor_Maximo
  Progreso_Esperar
  apexcel.Visible = True
  Pagina = 1
  Set apexcel = Nothing
  RatonNormal
  Progreso_Final
End Sub

Public Sub Imprimir_Actas(Curso As String)
Dim AnchoPict As Single
Dim AltoPict As Single
On Error GoTo Errorhandler
    Mensajes = "Imprimir Acta de Grado"
    Titulo = "Pregunta de Imprimir"
    Bandera = False
    SetPrinters.Show 1
    If PonImpresoraDefecto(SetNombrePRN) Then
       InicioX = 0: InicioY = 0
       Escala_Centimetro Orientacion_Pagina, TipoTimes, 8
       Pagina = 1
       Printer.ScaleMode = vbTwips
       AnchoPict = Round(Printer.ScaleWidth, 6)
       AltoPict = Round(Printer.ScaleHeight, 6)
        sSQL = "SELECT C.Cliente,C.Direccion,TA.Codigo " _
             & "FROM Clientes As C,Trans_Actas As TA " _
             & "WHERE TA.Periodo = '" & Periodo_Contable & "' " _
             & "AND TA.Id_No <> 0 " _
             & "AND TA.Item = '" & NumEmpresa & "' " _
             & "AND C.Grupo = '" & Curso & "' " _
             & "AND C.Codigo = TA.Codigo " _
             & "ORDER BY C.Direccion,C.Cliente "
        SelectAdodc AdoDetalle, sSQL
        With AdoDetalle.Recordset
         If .RecordCount > 0 Then
             Do While Not .EOF
                Listar_Acta_Grado .Fields("Codigo"), True
                Printer.PaintPicture PictLibreta.Image, InicioX, InicioY, AnchoPict, AltoPict
                Printer.NewPage
               .MoveNext
             Loop
         End If
        End With
        MensajeEncabData = ""
        Printer.EndDoc
        RatonNormal
        Exit Sub
Errorhandler:
             PictLibreta.Visible = True
             RatonNormal
             ErrorDeImpresion
             Exit Sub
    Else
       RatonNormal
    End If
End Sub

Public Sub Imprimir_Actas_Pag2(Curso As String)
Dim AnchoPict As Single
Dim AltoPict As Single
On Error GoTo Errorhandler
    Mensajes = "Imprimir Acta de Grado"
    Titulo = "Pregunta de Imprimir"
    Bandera = False
    SetPrinters.Show 1
    If PonImpresoraDefecto(SetNombrePRN) Then
       InicioX = 0: InicioY = 0
       Escala_Centimetro Orientacion_Pagina, TipoTimes, 8
       Pagina = 1
       Printer.ScaleMode = vbTwips
       AnchoPict = Round(Printer.ScaleWidth, 6)
       AltoPict = Round(Printer.ScaleHeight, 6)
        sSQL = "SELECT C.Cliente,C.Direccion,TA.Codigo " _
             & "FROM Clientes As C,Trans_Actas As TA " _
             & "WHERE TA.Periodo = '" & Periodo_Contable & "' " _
             & "AND TA.Item = '" & NumEmpresa & "' " _
             & "AND C.Grupo = '" & Curso & "' " _
             & "AND C.Codigo = TA.Codigo " _
             & "ORDER BY C.Direccion,C.Cliente "
        SelectAdodc AdoDetalle, sSQL
        With AdoDetalle.Recordset
         If .RecordCount > 0 Then
             Do While Not .EOF
                Listar_Acta_Grado_Pag2 .Fields("Codigo"), True
                Printer.PaintPicture PictLibreta.Image, InicioX, InicioY, AnchoPict, AltoPict
                Printer.NewPage
               .MoveNext
             Loop
         End If
        End With
        MensajeEncabData = ""
        Printer.EndDoc
        RatonNormal
        Exit Sub
Errorhandler:
             PictLibreta.Visible = True
             RatonNormal
             ErrorDeImpresion
             Exit Sub
    Else
       RatonNormal
    End If
End Sub

Public Sub Listar_Acta_Grado(CodigoAlumno As String, ActaOriginal As Boolean)
Dim SumaNotas As Currency
Dim VCalif(5) As String
Dim LogoAux As String
Dim Cursos  As String
    InicioX = 0.5: InicioY = 0.1: Abono = 0
    'Pagina = 1
    'Iniciamos la impresion
    PictLibreta.Cls
    PictLibreta.width = AnchoMaximo
    PictLibreta.Height = AltoMaximo
    
    PictLibreta.FontName = TipoTimes
    PictLibreta.FontBold = False
    FechaValida MBFecha
    Select Case Mid$(CodigoL, 1, 4)   ' Curso del Alumno
      Case "2.00" To "3.99"
           Unidad = Secretario2
           Carta_Porte = TextoVicerrector1
           Codigo2 = TextoBachiller1
      Case "4.00" To "5.99"
           Unidad = Secretario3
           Carta_Porte = TextoVicerrector2
           Codigo2 = TextoBachiller2
    End Select
    Contra_Cta = Ninguno
    Cursos = ""
    With AdoMatriculas.Recordset
      If .RecordCount > 0 Then
         .MoveFirst
         .Find ("Codigo = '" & CodigoCliente & "' ")
          If Not .EOF Then
             Cursos = Leer_Datos_del_Curso(CodigoL, 2)
             Contra_Cta = Dato_Curso.Especialidad
          End If
      End If
    End With
    
    sSQL = "UPDATE Trans_Actas " _
         & "SET Notas = ROUND(Notas,2), " _
         & "Trabajo = ROUND(Trabajo,2)," _
         & "Investigacion = ROUND(Investigacion,2)," _
         & "Evaluacion = ROUND(Evaluacion,2) " _
         & "WHERE Periodo = '" & Periodo_Contable & "' " _
         & "AND Item = '" & NumEmpresa & "' " _
         & "AND Codigo = '" & CodigoAlumno & "' "
    ConectarAdoExecute sSQL
         
    sSQL = "SELECT Cliente As Alumno,Direccion As Curso,Sexo,TA.* " _
         & "FROM Clientes As C,Trans_Actas As TA " _
         & "WHERE TA.Periodo = '" & Periodo_Contable & "' " _
         & "AND TA.Item = '" & NumEmpresa & "' " _
         & "AND C.Codigo = '" & CodigoAlumno & "' " _
         & "AND C.Codigo = TA.Codigo "
    SelectAdodc AdoDetalle, sSQL
    With AdoDetalle.Recordset
     If .RecordCount > 0 Then
         RatonReloj
        Cod_Bodega = Codigo1
        Mifecha = MBFecha
        NoDias = FechaDia(MBFecha)
        NoMeses = FechaMes(MBFecha)
        NoAnio = FechaAnio(MBFecha)
        If ActaOriginal Then
           SQLMsg2 = "ACTA DE GRADO  No. " & Format(.Fields("Id_No"), "000")
        Else
           SQLMsg2 = "COPIA DE ACTA DE GRADO " & Space(20) & " No. " & Format(.Fields("Id_No"), "000")
        End If
        SQLMsg3 = ""
        SQLMsg2 = ""
        SQLMsg1 = NombreCiudad & " - " & ULCase(NombrePais)
        Total = .Fields("PromFinal")
        ValorUnit = Redondear(.Fields("PromFinal"))
        SumaNotas = .Fields("Notas") + .Fields("Trabajo") + .Fields("Investigacion") + .Fields("Evaluacion")
        DirCliente = "."
        Select Case Redondear(SumaNotas / 4)
          Case 0 To 11:  DirCliente = "INSUFICIENTE"
          Case 12 To 13: DirCliente = "REGULAR"
          Case 14 To 15: DirCliente = "BUENA"
          Case 16 To 18: DirCliente = "MUY BUENA"
          Case 19 To 20: DirCliente = "SOBRESALIENTE"
        End Select
        If FormatoLibreta = "BIMESTRES" Then
           VCalif(0) = "a) Promedio de la Notas Finales 8vo. Año a 2do. Año de Bachillerato:"   'Notas
           VCalif(3) = "b) Evaluación Escrita del Perfil de Bachiller en Ciencias:"   'Investigacion
           VCalif(2) = "c) Total Final del trabajo de Alfabetización y Defensa Civíl:"   'Trabajo
           VCalif(1) = "d) Promedio Tercer año Bachillerato:"   'Evaluacion
           VCalif(4) = "El promedio Final de las cuatro notas es de:"    'SumaNotas
        Else
           VCalif(0) = "a) Promedio de la Notas Globales de primero a tercero del ciclo básico y de primero " _
                     & "   y segundo curso del ciclo diversificado."  'Notas
           VCalif(1) = "b) Promedio Global Correspondiente al tercer curso del ciclo diversificado."  'Investigacion
           VCalif(2) = "c) Nota final del trabajo de investigación, o práctico. (participación estudiantil)." 'Trabajo
           VCalif(3) = "d) Promedio de exámenes escritos de Grado."  'Evaluacion
           VCalif(4) = "d) Promedio General:"   'SumaNotas
        End If
        CodigoCliente = .Fields("Codigo")
        NombreCliente = UCase(.Fields("Alumno"))
        LogoAux = RutaSistema & "\LOGOS\ECUADOR.GIF"
        PictPrint_Grafico PictLibreta, LogoAux, 10, 0.1, 1.5, 1.5
        LogoAux = RutaSistema & "\LOGOS\MINISEDU.GIF"
        PictPrint_Grafico PictLibreta, LogoAux, 1.8, 0.1, 3.5, 1.5
        
        PosLinea = 1.7
        PictLibreta.FontSize = 16
        PictPrint_Texto PictLibreta, 1, PosLinea, "REPÚBLICA DEL ECUADOR", , 19, True
        PosLinea = PosLinea + 0.7
        PictPrint_Texto PictLibreta, 1, PosLinea, "MINISTERIO DE ECUCACION", , 19, True
        PosLinea = PosLinea + 1
        PictLibreta.FontSize = 14
        PictPrint_Texto PictLibreta, 11, PosLinea, "ACTA DE GRADO No. "
        PictPrint_Texto PictLibreta, 17, PosLinea, Format(.Fields("Id_No"), "000")
        PosLinea = PosLinea + 0.7
        PictLibreta.FontSize = 12
        Cadena = "La suscrita Secretaria Titular de la " & Institucion1 & " " & Institucion2 _
               & " en forma legal confiere el Acta de Grado correspondiente a "
        If .Fields("Sexo") = "M" Then
            Cadena = Cadena & "el Señor:"
        Else
            Cadena = Cadena & "la Señorita:"
        End If
        PosLinea = PictPrint_Texto_Justifica(PictLibreta, 2, 18, PosLinea, Cadena)
        PosLinea = PosLinea + 0.6
        PictPrint_Texto PictLibreta, 1, PosLinea, NombreCliente, , 19, True
        PosLinea = PosLinea + 0.7
        PictLibreta.FontSize = 12
        
        Cadena = "En el Cantón " & ULCase(NombreCiudad) & ", Provincia de " & ULCase(NombreProvincia) & ", a los " _
               & LCase(Cambio_Letras(NoDias, True)) & " " _
               & "días del mes de " & LCase(MesesLetras(NoMeses)) & " de " & LCase(Cambio_Letras(NoAnio, True)) & ", " _
               & "el Consejo Directivo de la " & Institucion1 & " " & Institucion2 _
               & " integrado por los siguientes miembros: "
        PosLinea = PictPrint_Texto_Justifica(PictLibreta, 2, 18, PosLinea, Cadena)
        PosLinea = PosLinea + 0.6
        PosLinea = PictPrint_Texto_Justifica(PictLibreta, 2, 18, PosLinea, Directiva)
        PosLinea = PosLinea + 0.6
        Cadena = "de conformidad con lo dispuesto en el Art. 248 del Reglamento General de la Ley " _
               & "de Educación, certifica que:"
        PosLinea = PictPrint_Texto_Justifica(PictLibreta, 2, 18, PosLinea, Cadena)
        PosLinea = PosLinea + 0.6
        If .Fields("Sexo") = "M" Then
            PictPrint_Texto PictLibreta, 2, PosLinea, "El Señor:"
        Else
            PictPrint_Texto PictLibreta, 2, PosLinea, "La Señorita:"
        End If
        PictPrint_Texto PictLibreta, 1, PosLinea, NombreCliente, , 19, True
        PosLinea = PosLinea + 0.6
        PictPrint_Texto PictLibreta, 2, PosLinea, "ha obtenido las siguientes calificaciones:"
        PosLinea = PosLinea + 0.6
        PosLinea = PictPrint_Texto_Justifica(PictLibreta, 2, 17, PosLinea, VCalif(0))
        PictPrint_Texto PictLibreta, 17, PosLinea, Format(.Fields("Notas"), "00.00")
        PosLinea = PosLinea + 0.6
        PosLinea = PictPrint_Texto_Justifica(PictLibreta, 2, 17, PosLinea, VCalif(1))
        PictPrint_Texto PictLibreta, 17, PosLinea, Format(.Fields("Investigacion"), "00.00")
        PosLinea = PosLinea + 0.6
        PosLinea = PictPrint_Texto_Justifica(PictLibreta, 2, 17, PosLinea, VCalif(2))
        PictPrint_Texto PictLibreta, 17, PosLinea, Format(.Fields("Trabajo"), "00.00")
        PosLinea = PosLinea + 0.6
        PosLinea = PictPrint_Texto_Justifica(PictLibreta, 2, 17, PosLinea, VCalif(3))
        PictPrint_Texto PictLibreta, 17, PosLinea, Format(.Fields("Evaluacion"), "00.00")
        PosLinea = PosLinea + 0.6
        PictLibreta.Line (16.5, PosLinea)-(18.2, PosLinea), Negro
        PosLinea = PosLinea + 0.1
        PictPrint_Texto PictLibreta, 14, PosLinea, "T O T A L:"
        PictPrint_Texto PictLibreta, 17, PosLinea, Format(SumaNotas, "00.00")
        PosLinea = PosLinea + 0.6
        PictPrint_Texto PictLibreta, 12.8, PosLinea, "Promedio General:"
        SumaNotas = Val(Format(SumaNotas / 4, "00.00"))
        PictPrint_Texto PictLibreta, 17, PosLinea, Format(SumaNotas, "00.00")
        PosLinea = PosLinea + 0.6
        SumaNotas = Val(Format(SumaNotas, "00"))
        PictPrint_Texto PictLibreta, 2, PosLinea, "NOTA DEFINITIVA DE GRADO:"
        PictPrint_Texto PictLibreta, 9, PosLinea, "(" & Format(SumaNotas, "00") & ") " & Cambio_Letras(SumaNotas, True)
        PosLinea = PosLinea + 0.6
        PosLinea = PictPrint_Texto_Multiple(PictLibreta, 2, PosLinea, "EQUIVALENTE A:", 17)
        PictPrint_Texto PictLibreta, 9, PosLinea, DirCliente
        PosLinea = PosLinea + 0.6
        Cadena = "En virtud de la aprobación el Consejo Directivo le confiere el título de " & Cursos
        PosLinea = PictPrint_Texto_Justifica(PictLibreta, 2, 18, PosLinea, Cadena)
        PosLinea = PosLinea + 0.6
        Cadena = "Por todo lo actuado, los Miembros del Consejo Directivo se ratifican y " _
               & "firman en unidad de acto conjuntamente con la secretaria que da fe y certifica:"
        PosLinea = PictPrint_Texto_Justifica(PictLibreta, 2, 18, PosLinea, Cadena)
        PosLinea = PosLinea + 1.5
        PictPrint_Texto PictLibreta, 3, PosLinea, String(20, "_")
        PictPrint_Texto PictLibreta, 11, PosLinea, String(20, "_")
        PosLinea = PosLinea + 0.5
        PictPrint_Texto PictLibreta, 4, PosLinea, TextoRector
        PictPrint_Texto PictLibreta, 11.5, PosLinea, "VICE-RECTORA"
        PosLinea = PosLinea + 1.5
        PictPrint_Texto PictLibreta, 3, PosLinea, String(20, "_")
        PictPrint_Texto PictLibreta, 11, PosLinea, String(20, "_")
        PosLinea = PosLinea + 0.5
        PictPrint_Texto PictLibreta, 4, PosLinea, "PRIMER VOCAL"
        PictPrint_Texto PictLibreta, 11.5, PosLinea, "SEGUNDO VOCAL"
        PosLinea = PosLinea + 1.5
        PictPrint_Texto PictLibreta, 3, PosLinea, String(20, "_")
        PictPrint_Texto PictLibreta, 11, PosLinea, String(20, "_")
        PosLinea = PosLinea + 0.5
        PictPrint_Texto PictLibreta, 4, PosLinea, "TERCER VOCAL"
        PictPrint_Texto PictLibreta, 11.5, PosLinea, TextoSecretario2
     End If
    End With
    RatonNormal
End Sub

Public Sub Listar_Acta_Grado_Pag2(CodigoAlumno As String, ActaOriginal As Boolean)
Dim SumaNotas As Currency
Dim VCalif(5) As String
Dim LogoAux As String
Dim Cursos  As String
    InicioX = 0.5: InicioY = 0.1: Abono = 0
    'Pagina = 1
    'Iniciamos la impresion
    PictLibreta.Cls
    PictLibreta.FontName = TipoTimes
    PictLibreta.FontBold = False
    FechaValida MBFecha
    sSQL = "SELECT Cliente As Alumno,Direccion As Curso,TA.* " _
         & "FROM Clientes As C,Trans_Actas As TA " _
         & "WHERE TA.Periodo = '" & Periodo_Contable & "' " _
         & "AND TA.Item = '" & NumEmpresa & "' " _
         & "AND C.Codigo = '" & CodigoAlumno & "' " _
         & "AND C.Codigo = TA.Codigo "
    SelectAdodc AdoDetalle, sSQL
    With AdoDetalle.Recordset
     If .RecordCount > 0 Then
         RatonReloj
         Cod_Bodega = Codigo1
         Mifecha = MBFecha
         NoDias = FechaDia(MBFecha)
         NoMeses = FechaMes(MBFecha)
         NoAnio = FechaAnio(MBFecha)
        PosLinea = 1.7
        PictLibreta.FontSize = 12
        Cadena = "Puede el interesado(a) hacer uso de la presente Acta de Grado en la forma que " _
               & "estime necesaria, remitiéndome, si el caso requiere, a los libros y registros que " _
               & "reposan en el archivo a mi cargo."
        PosLinea = PictPrint_Texto_Multiple(PictLibreta, 2, PosLinea, Cadena, 16)
        PosLinea = PosLinea + 1
        PictPrint_Texto PictLibreta, 9, PosLinea, "USO SECCIÓN REFRENDACIÓN"
        PosLinea = PosLinea + 1
        PictLibreta.Line (2, PosLinea)-(9.5, PosLinea + 9), Negro, B
        PictLibreta.Line (2.1, PosLinea + 0.1)-(9.4, PosLinea + 8.9), Negro, B
        
        PictLibreta.Line (10, PosLinea)-(18, PosLinea + 9), Negro, B
        PictLibreta.Line (10.1, PosLinea + 0.1)-(17.9, PosLinea + 8.9), Negro, B
        
        PictPrint_Texto PictLibreta, 3, PosLinea + 0.3, FechaStrgCiudad(Mifecha)
        PosLinea = PosLinea + 9.5
        Cadena = "Se certifica que la firma de la " & DirectorRegional & " es auténtica."
        PosLinea = PictPrint_Texto_Multiple(PictLibreta, 2, PosLinea, Cadena, 16)
     End If
    End With
    RatonNormal
End Sub

Public Sub xxxListar_Acta_Grado(CodigoClient As String)
Dim TLN As String
  FechaValida MBFecha
  sSQL = "SELECT Cliente As Alumno,Direccion As Curso,TA.* " _
       & "FROM Clientes As C,Trans_Actas As TA " _
       & "WHERE C.Codigo = '" & CodigoClient & "' " _
       & "AND TA.Periodo = '" & Periodo_Contable & "' " _
       & "AND TA.Item = '" & NumEmpresa & "' " _
       & "AND C.Codigo = TA.Codigo " _
       & "ORDER BY C.Sexo DESC,C.Direccion,C.Cliente "
  SelectAdodc AdoDetalle, sSQL
  'MsgBox sSQL
  With AdoDetalle.Recordset
   If .RecordCount > 0 Then
       RatonReloj
       PictLibreta.Cls
       PictLibreta.PaintPicture LoadPicture(LogoTipo), 0.45, 0.2, 3.9, 2
       PictLibreta.ForeColor = QBColor(Negro)
       PictLibreta.FontName = TipoTimes
       PictLibreta.FontSize = 20
       PictPrint_Texto PictLibreta, 4, 0.1, Empresa
       PictLibreta.FontSize = 11
       PictPrint_Texto PictLibreta, 4, 1, "A C T A     D E     G R A D O"
       PictLibreta.FontSize = 16
       PictPrint_Texto PictLibreta, 4, 1.5, "AÑO LECTIVO " & Anio_Lectivo
       PictLibreta.FontSize = 9
       PictPrint_Texto PictLibreta, 17.5, 1.8, MBFecha
       PosLinea = 1.4
       PictLibreta.FontName = TipoCourier
       PictLibreta.FontSize = 10
'''       Rector = .Fields("Rector")
'''       Director = .Fields("Director")
'''       Secretario1 = .Fields("Secretario1")
'''       Secretario2 = .Fields("Secretario2")
'''       Secretario3 = .Fields("Secretario3")
'''       Rector = .Fields("Rector")
'''       Director = .Fields("Director")
'''       TextoRector = .Fields("Texto_Rector")
'''       TextoDirector = .Fields("Texto_Director")
'''       Anio_Lectivo = .Fields("Anio_Lectivo")
'''       FormatoLibreta = .Fields("Formato")
'''       Recomen = .Fields("Recomendacion")
'''       Escalas = .Fields("Escala")

       TLN = "En la ciudad de " & NombreCiudad & " Capital de la Republica del Ecuador, Provincia de Pichincha, " _
           & "a los " & Day(MBFecha) & " días del mes de " & MesesLetras(Month(MBFecha)) & " del año " & Cambio_Letras(Year(MBFecha), True) & "." _
           & "El Concejo Directivo de la " & Empresa & " Conformados por los siguientes Piembros: "
       
       Mifecha = MBFecha.Text
       NoDias = FechaDia(MBFecha.Text)
       NoMeses = FechaMes(MBFecha.Text)
       NoAnio = FechaAnio(MBFecha.Text)
       SQLMsg2 = "A C T A     D E     G R A D O"
       SQLMsg3 = ""
       SQLMsg1 = UCase(NombreCiudad) & " - ECUADOR"
       Total = Redondear(.Fields("PromFinal"))
       ValorUnit = Redondear(.Fields("PromFinal"))
       DirCliente = "."
       Select Case Redondear(.Fields("PromFinal"))
            Case 0 To 11:  DirCliente = "INSUFICIENTE"
            Case 12 To 13: DirCliente = "REGULAR"
            Case 14 To 15: DirCliente = "BUENA"
            Case 16 To 18: DirCliente = "MUY BUENA"
            Case 19 To 20: DirCliente = "SOBRESALIENTE"
        End Select
        NombreCliente = UCase(.Fields("Alumno"))
        
        'MsgBox NombreCliente
        Producto = "a) Promedio de notas Finales 8vo. Año Básico a 3ero. año Bachillerato ... " & Format(.Fields("Notas"), "00.00") & vbCrLf _
                 & "b) Trabajo Nuevo Rumbo Cultural ......................................... " & Format(.Fields("Trabajo"), "00.00") & vbCrLf _
                 & "c) Nota Final de Trabajo de Investigación y Defensa ..................... " & Format(.Fields("Investigacion"), "00.00") & vbCrLf _
                 & "d) Evaluación Escrita del Perfil de Bachiller en Ciencias ............... " & Format(.Fields("Evaluacion"), "00.00") & vbCrLf _
                 & "   El promedio final de las Cuatro notas es de .......................... " & Format(.Fields("PromFinal"), "00.00")
        PosLinea = PictPrint_Texto_Multiple(PictLibreta, 2, 5, Producto, 18)
        


TLN = "Primer, segundo y tercer vocal respectivamente, con el objeto de legalizar las pruebas rendidas por el(a)  Sr(ita)"
TLN = "a quien de acuerdo con lo que dispone el Articulo 239 del Reglamentado General de la Ley de Educación y Cultura, del Articulo 19 del reglamento para la aplicación de la Reforma Curricular del Bachillerato propuesta por la Universidad, Andina Simón Bolívar (Acuerdo Ministerial No. 1382), se procedió a señalarse la nota final de la siguiente forma:"

TLN = "Que de acuerdo con el Art. 239 del reglamento es de [N] [n]"
TLN = "Equivalente a [D], Calificación que habilita a El(a) Sr(ita)"
TLN = "Para recibir la investidura de BACHILLER EN CIENCIAS. Por tal efecto. La Hna RECTORA EN NOMBRE DE LA REPUBLICA Y POR AUTORIDAD DE LA LEY."

TLN = "Le confiere La investidura en forma reglamentaria, en fe de lo cual el H. Consejo Directivo, representado por sus miembros procede a suscribir la presente Acta, Con la intervención de la Secretaria que certifica."
TLN = "f) Rectora Presidenta   _________________        f) Vice- Rector   _______________"
TLN = "f) Primer Vocal         _________________        f) Segundo Vocal  _______________"
TLN = "f) Tercer Vocal         _________________        f) Secretaria     _______________"

        
        
        
   End If
  End With
  RatonNormal
End Sub

Private Sub TxtTitulo_GotFocus()
  MarcarTexto TxtTitulo
End Sub

Private Sub TxtTitulo_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtTitulo_LostFocus()
  TextoValido TxtTitulo, , True
End Sub

Private Sub TxtObservacion_GotFocus()
  MarcarTexto TxtObservacion
End Sub

Private Sub TxtObservacion_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtObservacion_LostFocus()
  TextoValido TxtObservacion, , True
End Sub

Private Sub VScroll1_Change()
  VScroll1_Scroll
End Sub

Private Sub VScroll1_Scroll()
  PictLibreta.Top = -VScroll1.value
  LblA4.Caption = " " & Format(VScroll1.value, "00.00") & " - " & Format(HScroll1.value, "00.00")
End Sub

Private Sub HScroll1_Change()
  HScroll1_Scroll
End Sub

Private Sub HScroll1_Scroll()
  PictLibreta.Left = -HScroll1.value
  LblA4.Caption = " " & Format(VScroll1.value, "00.00") & " - " & Format(HScroll1.value, "00.00")
End Sub

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
   Escala_Centimetro Orientacion_Pagina, TipoTimes, 8
   Pagina = 1
  'vbTwips
   If Printer.ScaleMode = vbCentimeters Then
      Printer.ScaleWidth = Me.ScaleX(Printer.ScaleWidth, vbCentimeters, vbPixels)
      Printer.ScaleHeight = Me.ScaleX(Printer.ScaleHeight, vbCentimeters, vbPixels)
      Printer.ScaleMode = vbPixels
   End If

   Printer.PaintPicture PictLibreta.Image, 0, 0, Printer.ScaleWidth, Printer.ScaleHeight

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

Public Sub Leer_Asistencia_Materia(CodCurso As String, CodigoAlum As String, CodMat As String)
  Atrasos_PorMat = 0
  Faltas_Just_PorMat = 0
  Faltas_Injust_PorMat = 0
  
  sSQL = "SELECT * " _
       & "FROM Trans_Asistencia " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND Codigo = '" & CodigoAlum & "' " _
       & "AND CodE = '" & CodCurso & "' " _
       & "AND CodMat = '" & CodMat & "' "
  SelectAdodc AdoLibreta, sSQL
  With AdoLibreta.Recordset
   If .RecordCount > 0 Then
       If FormatoLibreta = "QUIMESTRE" Then
            If OpcPeriodo("PQBim1", LstPeriodos) Then
               Faltas_Just_PorMat = .Fields("PQBFJ1")
               Faltas_Injust_PorMat = .Fields("PQBFI1")
               Atrasos_PorMat = .Fields("PQBA1")
            End If
            If OpcPeriodo("PQBim2", LstPeriodos) Then
               Faltas_Just_PorMat = .Fields("PQBFJ2")
               Faltas_Injust_PorMat = .Fields("PQBFI2")
               Atrasos_PorMat = .Fields("PQBA2")
            End If
            If OpcPeriodo("PQBim3", LstPeriodos) Then
               Faltas_Just_PorMat = .Fields("PQBFJ3")
               Faltas_Injust_PorMat = .Fields("PQBFI3")
               Atrasos_PorMat = .Fields("PQBA3")
            End If
            If OpcPeriodo("PQ", LstPeriodos) Then
               Faltas_Just_PorMat = .Fields("PQBFJ1") + .Fields("PQBFJ2") + .Fields("PQBFJ3")
               Faltas_Injust_PorMat = .Fields("PQBFI1") + .Fields("PQBFI2") + .Fields("PQBFI3")
               Atrasos_PorMat = .Fields("PQBA1") + .Fields("PQBA2") + .Fields("PQBA3")
            End If
            If OpcPeriodo("SQBim1", LstPeriodos) Then
               Faltas_Just_PorMat = .Fields("SQBFJ1")
               Faltas_Injust_PorMat = .Fields("SQBFI1")
               Atrasos_PorMat = .Fields("SQBA1")
            End If
            If OpcPeriodo("SQBim2", LstPeriodos) Then
               Faltas_Just_PorMat = .Fields("SQBFJ2")
               Faltas_Injust_PorMat = .Fields("SQBFI2")
               Atrasos_PorMat = .Fields("SQBA2")
            End If
            If OpcPeriodo("SQBim3", LstPeriodos) Then
               Faltas_Just_PorMat = .Fields("SQBFJ3")
               Faltas_Injust_PorMat = .Fields("SQBFI3")
               Atrasos_PorMat = .Fields("SQBA3")
            End If
            If OpcPeriodo("SQ", LstPeriodos) Then
               Faltas_Just_PorMat = .Fields("SQBFJ1") + .Fields("SQBFJ2") + .Fields("SQBFJ3")
               Faltas_Injust_PorMat = .Fields("SQBFI1") + .Fields("SQBFI2") + .Fields("SQBFI3")
               Atrasos_PorMat = .Fields("SQBA1") + .Fields("SQBA2") + .Fields("SQBA3")
            End If
       Else
           If OpcPeriodo("PQBim1", LstPeriodos) Or OpcPeriodo("PQ", LstPeriodos) Then
              Faltas_Just_PorMat = .Fields("PQBFJ1") + .Fields("PQBFJ2") + .Fields("PQBFJ3")
              Faltas_Injust_PorMat = .Fields("PQBFI1") + .Fields("PQBFI2") + .Fields("PQBFI3")
              Atrasos_PorMat = .Fields("PQBA1") + .Fields("PQBA2") + .Fields("PQBA3")
           ElseIf OpcPeriodo("SQBim1", LstPeriodos) Or OpcPeriodo("SQ", LstPeriodos) Or OpcPeriodo("PF", LstPeriodos) Then
              Faltas_Just_PorMat = .Fields("SQBFJ1") + .Fields("SQBFJ2") + .Fields("SQBFJ3")
              Faltas_Injust_PorMat = .Fields("SQBFI1") + .Fields("SQBFI2") + .Fields("SQBFI3")
              Atrasos_PorMat = .Fields("SQBA1") + .Fields("SQBA2") + .Fields("SQBA3")
           End If
       End If
   End If
  End With
End Sub

Public Sub Notas_Del_Alumno(CodCurso As String, CodigoAlum As String)
Dim SumaPeriodos As String
  Valor = 0
  Atrasos = 0
  Faltas_Just = 0
  Faltas_Injust = 0
  Dias_Laborados = 0
  
  Atrasos1 = 0
  Faltas_Just1 = 0
  Faltas_Injust1 = 0
  Dias_Laborados1 = 0
  
  Atrasos2 = 0
  Faltas_Just2 = 0
  Faltas_Injust2 = 0
  Dias_Laborados2 = 0
  
  Atrasos3 = 0
  Faltas_Just3 = 0
  Faltas_Injust3 = 0
  Dias_Laborados3 = 0
  
  sSQL = "SELECT * " _
       & "FROM Trans_Asistencia " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND Codigo = '" & CodigoAlum & "' " _
       & "AND CodE = '" & CodCurso & "' " _
       & "AND CodMat BETWEEN '997' and '999' "
  SelectAdodc AdoLibreta, sSQL
  With AdoLibreta.Recordset
   If .RecordCount > 0 Then
       'MsgBox FormatoLibreta
       If FormatoLibreta = "QUIMESTRE" Then
            If OpcPeriodo("PF", LstPeriodos) Then
            End If
            If OpcPeriodo("PQBim1", LstPeriodos) Then
               Valor = .Fields("ConductaPQ1")
               Dias_Laborados1 = .Fields("PQDias1")
               Faltas_Just1 = .Fields("PQBFJ1")
               Faltas_Injust1 = .Fields("PQBFI1")
               Atrasos1 = .Fields("PQBA1")
            End If
            If OpcPeriodo("PQBim2", LstPeriodos) Then
               Valor = .Fields("ConductaPQ2")
               Dias_Laborados2 = .Fields("PQDias2")
               Faltas_Just2 = .Fields("PQBFJ2")
               Faltas_Injust2 = .Fields("PQBFI2")
               Atrasos2 = .Fields("PQBA2")
            End If
            If OpcPeriodo("PQBim3", LstPeriodos) Then
               Valor = .Fields("ConductaPQ3")
               Dias_Laborados3 = .Fields("PQDias3")
               Faltas_Just3 = .Fields("PQBFJ3")
               Faltas_Injust3 = .Fields("PQBFI3")
               Atrasos3 = .Fields("PQBA3")
            End If
            If OpcPeriodo("PQ", LstPeriodos) Then
               Select Case Anio_Lectivo
                 Case Is >= "2014 - 2015", Ninguno
                      Valor = .Fields("ConductaPQ3")
                 Case Else
                      Valor = Redondear((.Fields("ConductaPQ1") + .Fields("ConductaPQ2") + .Fields("ConductaPQ3")) / 3, 0)
               End Select
               Dias_Laborados1 = .Fields("PQDias1")
               Faltas_Just1 = .Fields("PQBFJ1")
               Faltas_Injust1 = .Fields("PQBFI1")
               Atrasos1 = .Fields("PQBA1")
               
               Dias_Laborados2 = .Fields("PQDias2")
               Faltas_Just2 = .Fields("PQBFJ2")
               Faltas_Injust2 = .Fields("PQBFI2")
               Atrasos2 = .Fields("PQBA2")
               
               Dias_Laborados3 = .Fields("PQDias3")
               Faltas_Just3 = .Fields("PQBFJ3")
               Faltas_Injust3 = .Fields("PQBFI3")
               Atrasos3 = .Fields("PQBA3")
               
               Dias_Laborados = .Fields("PQDias1") + .Fields("PQDias2") + .Fields("PQDias3")
               Faltas_Just = .Fields("PQBFJ1") + .Fields("PQBFJ2") + .Fields("PQBFJ3")
               Faltas_Injust = .Fields("PQBFI1") + .Fields("PQBFI2") + .Fields("PQBFI3")
               Atrasos = .Fields("PQBA1") + .Fields("PQBA2") + .Fields("PQBA3")
            End If
            
            If OpcPeriodo("SQBim1", LstPeriodos) Then
               Valor = .Fields("ConductaSQ1")
               Dias_Laborados1 = .Fields("SQDias1")
               Faltas_Just1 = .Fields("SQBFJ1")
               Faltas_Injust1 = .Fields("SQBFI1")
               Atrasos1 = .Fields("SQBA1")
            End If
            If OpcPeriodo("SQBim2", LstPeriodos) Then
               Valor = .Fields("ConductaSQ2")
               Dias_Laborados2 = .Fields("SQDias2")
               Faltas_Just2 = .Fields("SQBFJ2")
               Faltas_Injust2 = .Fields("SQBFI2")
               Atrasos2 = .Fields("SQBA2")
            End If
            If OpcPeriodo("SQBim3", LstPeriodos) Then
               Valor = .Fields("ConductaSQ3")
               Dias_Laborados3 = .Fields("SQDias3")
               Faltas_Just3 = .Fields("SQBFJ3")
               Faltas_Injust3 = .Fields("SQBFI3")
               Atrasos3 = .Fields("SQBA3")
            End If
            If OpcPeriodo("SQ", LstPeriodos) Then
               Select Case Anio_Lectivo
                 Case Is >= "2014 - 2015", Ninguno
                      Valor = .Fields("ConductaSQ3")
                 Case Else
                      Valor = Redondear((.Fields("ConductaSQ1") + .Fields("ConductaSQ2") + .Fields("ConductaSQ3")) / 3, 0)
               End Select
               Dias_Laborados1 = .Fields("SQDias1")
               Faltas_Just1 = .Fields("SQBFJ1")
               Faltas_Injust1 = .Fields("SQBFI1")
               Atrasos1 = .Fields("SQBA1")
               
               Dias_Laborados2 = .Fields("SQDias2")
               Faltas_Just2 = .Fields("SQBFJ2")
               Faltas_Injust2 = .Fields("SQBFI2")
               Atrasos2 = .Fields("SQBA2")
               
               Dias_Laborados3 = .Fields("SQDias3")
               Faltas_Just3 = .Fields("SQBFJ3")
               Faltas_Injust3 = .Fields("SQBFI3")
               Atrasos3 = .Fields("SQBA3")
               
               Dias_Laborados = .Fields("SQDias3") + .Fields("SQDias3") + .Fields("SQDias3")
               Faltas_Just = .Fields("SQBFJ1") + .Fields("SQBFJ2") + .Fields("SQBFJ3")
               Faltas_Injust = .Fields("SQBFI1") + .Fields("SQBFI2") + .Fields("SQBFI3")
               Atrasos = .Fields("SQBA1") + .Fields("SQBA2") + .Fields("SQBA3")
            End If
       Else
           If OpcPeriodo("PQBim1", LstPeriodos) Or OpcPeriodo("PQ", LstPeriodos) Then
              Faltas_Just = .Fields("PQBFJ1") + .Fields("PQBFJ2") + .Fields("PQBFJ3")
              Faltas_Injust = .Fields("PQBFI1") + .Fields("PQBFI2") + .Fields("PQBFI3")
              Atrasos = .Fields("PQBA1") + .Fields("PQBA2") + .Fields("PQBA3")
           ElseIf OpcPeriodo("SQBim1", LstPeriodos) Or OpcPeriodo("SQ", LstPeriodos) Or OpcPeriodo("PF", LstPeriodos) Then
              Faltas_Just = .Fields("SQBFJ1") + .Fields("SQBFJ2") + .Fields("SQBFJ3")
              Faltas_Injust = .Fields("SQBFI1") + .Fields("SQBFI2") + .Fields("SQBFI3")
              Atrasos = .Fields("SQBA1") + .Fields("SQBA2") + .Fields("SQBA3")
    '''       Else
    '''          Faltas_Just = .Fields("PQBFJ1") + .Fields("PQBFJ2") + .Fields("SQBFJ1") + .Fields("SQBFJ2")
    '''          Faltas_Injust = .Fields("PQBFI1") + .Fields("PQBFI2") + .Fields("SQBFI1") + .Fields("SQBFI2")
    '''          Atrasos = .Fields("PQBA1") + .Fields("PQBA2") + .Fields("SQBA1") + .Fields("SQBA2")
           End If
           Real1 = .Fields("ConductaPQ1")
           Real2 = .Fields("ConductaPQ2")
           Real5 = .Fields("ConductaPQ3")
           Real3 = .Fields("ConductaSQ1")
           Real4 = .Fields("ConductaSQ2")
           Real6 = .Fields("ConductaSQ3")
       End If
   End If
  End With
  SQL1 = "SELECT CC.Curso,C.Sexo,CC.Descripcion As Paralelo,C.Cliente As Alumno," _
       & "CM.Materia,CM.C,CM.C2,CM.P,CM.I,CM.SDiv,CM.SinImprimir,TN.* " _
       & "FROM Trans_Notas As TN, Catalogo_Materias As CM, Catalogo_Cursos As CC, Clientes As C " _
       & "WHERE TN.Item = '" & NumEmpresa & "' " _
       & "AND TN.Periodo = '" & Periodo_Contable & "' " _
       & "AND TN.CodE = '" & CodCurso & "' " _
       & "AND TN.Codigo = '" & CodigoAlum & "' " _
       & "AND TN.Codigo = C.Codigo " _
       & "AND TN.CodE = CC.Curso " _
       & "AND TN.CodMat = CM.CodMat " _
       & "AND TN.Item = CC.Item " _
       & "AND TN.Item = CM.Item " _
       & "AND TN.Periodo = CC.Periodo " _
       & "AND TN.Periodo = CM.Periodo "
       
  SQL2 = "SELECT CC.Curso,C.Sexo,CC.Descripcion As Paralelo,C.Cliente As Alumno," _
       & "CM.Materia,CM.C,CM.C2,CM.P,CM.I,CM.SDiv,CM.SinImprimir,TN.* " _
       & "FROM Trans_Notas_Auxiliares As TN, Catalogo_Materias As CM, Catalogo_Cursos As CC, Clientes As C " _
       & "WHERE TN.Item = '" & NumEmpresa & "' " _
       & "AND TN.Periodo = '" & Periodo_Contable & "' " _
       & "AND TN.CodE = '" & CodCurso & "' " _
       & "AND TN.Codigo = '" & CodigoAlum & "' " _
       & "AND TN.Codigo = C.Codigo " _
       & "AND TN.CodE = CC.Curso " _
       & "AND TN.CodMat = CM.CodMat " _
       & "AND TN.Item = CC.Item " _
       & "AND TN.Item = CM.Item " _
       & "AND TN.Periodo = CC.Periodo " _
       & "AND TN.Periodo = CM.Periodo "
  sSQL = SQL1
'  If OpcionNotas <> 5 Then
     sSQL = sSQL _
          & "UNION " _
          & SQL2
'  End If
  sSQL = sSQL & "ORDER BY C.Cliente,TN.Id_No "      ' C.Sexo DESC,
  SelectAdodc AdoLibreta, sSQL
  'MsgBox AdoLibreta.Recordset.RecordCount & vbCrLf & sSQL
  SQL1 = "SELECT CM.Materia,"
  If OpcionNotas = 4 Then
     SumaPeriodos = "TN." & SQLBim1 & " + TN." & SQLBim2 & " + TN." & SQLBim3
     SQL1 = SQL1 _
          & "TN." & SQLBim1 & "," _
          & "TN." & SQLBim2 & "," _
          & "TN." & SQLBim3 & "," _
          & "ROUND((" & SumaPeriodos & ")/3,2,0) As Promedio," _
          & "ROUND(((" & SumaPeriodos & ")/3) * 0.80,2,0) As Prom_80," _
          & "TN." & SQLExamen & "," _
          & "ROUND(TN." & SQLExamen & " * 0.20,2,0) As Exam_20," _
          & "TN." & SQLPromQ & "," _
          & ""
  ElseIf OpcionNotas <> 5 Then
     SQL1 = SQL1 _
          & "TN." & SQLTAI & "," _
          & "TN." & SQLAIC & "," _
          & "TN." & SQLAGC & "," _
          & "TN." & SQLL & "," _
          & "TN." & SQLExaP & "," _
          & "TN." & SQLProm & ","
  Else
     SQL1 = SQL1 & "TN.*,"
  End If
  SQL1 = SQL1 & "TN.CodMat,TN.CodMatP,CC.Curso,TN.Orden,CM.SDiv,TN.Id_No "
  
  SQL2 = "FROM Trans_Notas As TN,Catalogo_Materias As CM,Catalogo_Cursos As CC,Clientes As C " _
       & "WHERE TN.Item = '" & NumEmpresa & "' " _
       & "AND TN.Periodo = '" & Periodo_Contable & "' " _
       & "AND TN.CodE = '" & CodCurso & "' " _
       & "AND TN.Codigo = '" & CodigoAlum & "' " _
       & "AND TN.Codigo = C.Codigo " _
       & "AND TN.CodE = CC.Curso " _
       & "AND TN.CodMat = CM.CodMat " _
       & "AND TN.Item = CC.Item " _
       & "AND TN.Item = CM.Item " _
       & "AND TN.Periodo = CC.Periodo " _
       & "AND TN.Periodo = CM.Periodo "
       
  SQL3 = "FROM Trans_Notas_Auxiliares As TN,Catalogo_Materias As CM,Catalogo_Cursos As CC,Clientes As C " _
       & "WHERE TN.Item = '" & NumEmpresa & "' " _
       & "AND TN.Periodo = '" & Periodo_Contable & "' " _
       & "AND TN.CodE = '" & CodCurso & "' " _
       & "AND TN.Codigo = '" & CodigoAlum & "' " _
       & "AND TN.Codigo = C.Codigo " _
       & "AND TN.CodE = CC.Curso " _
       & "AND TN.CodMat = CM.CodMat " _
       & "AND TN.Item = CC.Item " _
       & "AND TN.Item = CM.Item " _
       & "AND TN.Periodo = CC.Periodo " _
       & "AND TN.Periodo = CM.Periodo "
  sSQL = SQL1 & SQL2
    If OpcionNotas <> 5 Then
     sSQL = sSQL _
          & "UNION " _
          & SQL1 & SQL3
  End If
  sSQL = sSQL & "ORDER BY TN.CodMatP,TN.Id_No,CM.SDiv "
  SelectDataGrid DGNotasLibreta, AdoNotasLibreta, sSQL, , True
End Sub

Public Function Print_Nota(Nota_No As Byte) As Boolean
Dim FPrint_Nota(1 To 22) As Boolean
    For I = 1 To 22
        FPrint_Nota(I) = False
    Next I
    If FormatoLibreta = "BIMESTRES" Then
       If OpcPeriodo("PQBim1", LstPeriodos) Then
          FPrint_Nota(1) = True
       End If
       If OpcPeriodo("PQ", LstPeriodos) Then
          For I = 1 To 4
              FPrint_Nota(I) = True
          Next I
       End If
       If OpcPeriodo("SQBim1", LstPeriodos) Then
          For I = 1 To 5
              FPrint_Nota(I) = True
          Next I
       End If
       If OpcPeriodo("SQ", LstPeriodos) Then
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
    ElseIf Mid$(FormatoLibreta, 1, 9) = "TRIMESTRE" Then
       If OpcPeriodo("PQBim1", LstPeriodos) Then
          FPrint_Nota(1) = True
       End If
       If OpcPeriodo("PQBim2", LstPeriodos) Then
          For I = 1 To 2
              FPrint_Nota(I) = True
          Next I
       End If
       If OpcPeriodo("PQ", LstPeriodos) Then
          For I = 1 To 4
              FPrint_Nota(I) = True
          Next I
       End If
       If OpcPeriodo("SQBim1", LstPeriodos) Then
          For I = 1 To 5
              FPrint_Nota(I) = True
          Next I
       End If
       If OpcPeriodo("SQBim2", LstPeriodos) Then
          For I = 1 To 6
              FPrint_Nota(I) = True
          Next I
       End If
       If OpcPeriodo("SQ", LstPeriodos) Then
          For I = 1 To 8
              FPrint_Nota(I) = True
          Next I
       End If
       If OpcPeriodo("TQBim1", LstPeriodos) Then
          For I = 1 To 9
              FPrint_Nota(I) = True
          Next I
       End If
       If OpcPeriodo("TQBim2", LstPeriodos) Then
          For I = 1 To 10
              FPrint_Nota(I) = True
          Next I
       End If
       If OpcPeriodo("TQ", LstPeriodos) Then
          For I = 1 To 12
              FPrint_Nota(I) = True
          Next I
          FPrint_Nota(14) = True
       End If
       If OpcPeriodo("PF", LstPeriodos) Then
          For I = 1 To 14
              FPrint_Nota(I) = True
          Next I
       End If
    Else
       If OpcPeriodo("PQBim1", LstPeriodos) Then
          FPrint_Nota(1) = True
       End If
       If OpcPeriodo("PQBim2", LstPeriodos) Then
          For I = 1 To 2
              FPrint_Nota(I) = True
          Next I
       End If
       If OpcPeriodo("PQ", LstPeriodos) Then
          For I = 1 To 4
              FPrint_Nota(I) = True
          Next I
       End If
       If OpcPeriodo("SQBim1", LstPeriodos) Then
          For I = 1 To 5
              FPrint_Nota(I) = True
          Next I
       End If
       If OpcPeriodo("SQBim2", LstPeriodos) Then
          For I = 1 To 6
              FPrint_Nota(I) = True
          Next I
       End If
       If OpcPeriodo("SQ", LstPeriodos) Then
          For I = 1 To 12
              FPrint_Nota(I) = True
          Next I
          FPrint_Nota(10) = True
       End If
       If OpcPeriodo("PF", LstPeriodos) Then
          For I = 1 To 19
              FPrint_Nota(I) = True
          Next I
       End If
    End If
    Print_Nota = FPrint_Nota(Nota_No)
End Function

Public Sub Listar_Solicitud_Examen_Grado(TipoObjeto As Object, CodigoAlumno As String, Tipo_Solicitud As Byte)
Dim InicioLinea As Single
Dim UltimaLinea As Single
Dim Materias_Examen As String
Dim ContMatExam As Byte
RatonReloj
Notas_Del_Alumno CodigoL, CodigoCliente
InicioX = 0.5: InicioY = 0.1
'Iniciamos la impresion
If TypeOf TipoObjeto Is PictureBox Then TipoObjeto.Cls
TipoObjeto.FontName = TipoGeorgia   'TipoTimes
TipoObjeto.FontBold = False
TipoObjeto.width = 21
TipoObjeto.Height = 29
Materias_Examen = ""
ContMatExam = 0
sSQL = "SELECT * " _
     & "FROM Catalogo_Examen_Grado " _
     & "WHERE Item = '" & NumEmpresa & "' " _
     & "AND Periodo = '" & Periodo_Contable & "' " _
     & "AND CodigoE = '" & CodigoL & "' " _
     & "ORDER BY Detalle "
SelectAdodc AdoAux, sSQL
With AdoAux.Recordset
 If .RecordCount > 0 Then
     Do While Not .EOF
        ContMatExam = ContMatExam + 1
       .MoveNext
     Loop
    .MoveFirst
     IE = 0
     ReDim Materias_Examenes(ContMatExam) As TipoMaterias
     Do While Not .EOF
        Materias_Examenes(IE).CodigoMat = .Fields("CodMat")
        Materias_Examenes(IE).Materias = .Fields("Detalle")
        Materias_Examenes(IE).Valor = 0
        IE = IE + 1
        Materias_Examen = Materias_Examen & Trim(.Fields("Detalle")) & ", "
       .MoveNext
     Loop
 End If
End With

With AdoAlumnos.Recordset
 If .RecordCount > 0 Then
     'CodigoCliente = .Fields("Codigo")
     'MsgBox CodigoL
     SQLMsg1 = "AÑO LECTIVO: " & Anio_Lectivo
     SQLMsg2 = "CERTIFICADO DE MATRICULA"
     SQLMsg3 = ""
    .MoveFirst
    .Find ("Codigo = '" & CodigoCliente & "' ")
     If Not .EOF Then
        Listar_Notas_Examen_Grado CodigoCliente, ContMatExam
        PosLinea = 0.1
        TipoObjeto.FontBold = True
        If LogoTipo <> "" Then TipoObjeto.PaintPicture LoadPicture(LogoTipo), 10, PosLinea, 4, 2
        PosLinea = PosLinea + 2.5
        TipoObjeto.FontSize = 16
        PictPrint_Texto 1.5, PosLinea, Institucion1, , 20, True
        TipoObjeto.FontSize = 18
        PosLinea = PosLinea + 0.7
        PictPrint_Texto 1.5, PosLinea, Institucion2, , 20, True
        PosLinea = PosLinea + 0.7
        TipoObjeto.FontSize = 14
        PictPrint_Texto 10, PosLinea, NombreCiudad
        PosLinea = PosLinea + 1
        TipoObjeto.FontItalic = True
        TipoObjeto.FontSize = 12
        PictPrint_Texto 7.8, PosLinea, TextoLeyenda
        TipoObjeto.FontItalic = False
        PosLinea = PosLinea + 1
        Cadena1 = Leer_Datos_del_Curso(.Fields("Grupo_No"), 1)
        Select Case Tipo_Solicitud
          Case 1
                TipoObjeto.FontBold = False
                PictPrint_Texto 2.5, PosLinea, FechaStrgCiudad(MBFecha)
                PosLinea = PosLinea + 1.6
                PictPrint_Texto 2.5, PosLinea, Rector & ","
                PosLinea = PosLinea + 0.6
                PictPrint_Texto 2.5, PosLinea, "Rectora de la " & Institucion1 & " " & Institucion2
                PosLinea = PosLinea + 0.6
                PictPrint_Texto 2.5, PosLinea, "Ciudad."
                PosLinea = PosLinea + 1
                PictPrint_Texto 2.5, PosLinea, "Yo "
                TipoObjeto.FontBold = True
                PictPrint_Texto 2.5, PosLinea, .Fields("Alumno"), , 20, True
                TipoObjeto.FontBold = False
                PosLinea = PosLinea + 0.6
''                PosLinea = PictPrint_Texto_Multiple(TipoObjeto, 1.5, PosLinea, "Estudiante del " & Cadena1, 17)
''                PosLinea = PosLinea + 1
                Cadena = "Estudiante del " & Cadena1 & " " _
                       & "del Colegio de su digna rectoría, a usted respetuosamente expongo: Que he terminado " _
                       & "los estudios de segunda enseñanza, como consta en la documentación que acompaño " _
                       & "y por lo tanto solicito a usted, se sirva declararme apta para presentarme a los " _
                       & "Exámenes Escritos de Grado, previo a la obtención del título de BACHILLER en las " _
                       & "materias optativas de: "
                PosLinea = PictPrint_Texto_Justifica(TipoObjeto, 2.5, 19, PosLinea, Cadena)
                PosLinea = PosLinea + 0.6
                TipoObjeto.FontBold = True
                Materias_Examen = Trim(Materias_Examen)
                Materias_Examen = Mid$(Materias_Examen, 1, Len(Materias_Examen) - 1) & "."
                If Len(Materias_Examen) > 1 Then
                   PosLinea = PictPrint_Texto_Justifica(TipoObjeto, 2.5, 19, PosLinea, Materias_Examen)
                End If
                TipoObjeto.FontBold = False
                PosLinea = PosLinea + 0.6
                PictPrint_Texto 2.5, PosLinea, "Conforme lo dispone el artículo 243 del Reglamento General de la Ley de Educación."
                PosLinea = PosLinea + 1
                PictPrint_Texto 2.5, PosLinea, "Acompaño los siguientes requisitos:"
                PosLinea = PosLinea + 0.6
                PictPrint_Texto 2.5, PosLinea, "1.- Partida de nacimiento."
                PosLinea = PosLinea + 0.6
                PictPrint_Texto 2.5, PosLinea, "2.- Certificado de haber Terminado la Primaria."
                PosLinea = PosLinea + 0.6
                PictPrint_Texto 2.5, PosLinea, "3.- Matrículas del Octavo al Tercero de Bachillerato."
                PosLinea = PosLinea + 0.6
                PictPrint_Texto 2.5, PosLinea, "4.- Promociones del Octavo al Tercero de Bachillerato."
                PosLinea = PosLinea + 0.6
                PictPrint_Texto 2.5, PosLinea, "5.- Acta con la calificación de la Participación Estudiantil."
                PosLinea = PosLinea + 0.6
                PictPrint_Texto 2.5, PosLinea, "6.- Cédula de identidad (fotocopia)."
                PosLinea = PosLinea + 0.6
                PictPrint_Texto 2.5, PosLinea, "7.- Copia del Comprobante del ENES (Examen Nacional para la Educación Superior)."
                PosLinea = PosLinea + 1.6
                PictPrint_Texto 2.5, PosLinea, "Muy Atentamente"
                PosLinea = PosLinea + 2.5
                PictPrint_Texto 2.5, PosLinea, String(Len(.Fields("Alumno")) - 5, "_")
                PosLinea = PosLinea + 0.6
                PictPrint_Texto 2.5, PosLinea, ULCase(.Fields("Alumno"))
          Case 2
                PosLinea = PosLinea + 5
                TipoObjeto.FontBold = False
                Cadena = Rector & ", Rectora de la " & Institucion1 & " " & Institucion2 & " de " _
                       & "conformidad con el artículo 243 del Reglamento General de la Ley de Educación, una " _
                       & "vez revisada la documentación de la señorita:"
                PosLinea = PictPrint_Texto_Justifica(TipoObjeto, 2.5, 19, PosLinea, Cadena)
                PosLinea = PosLinea + 0.6
                TipoObjeto.FontBold = True
                PictPrint_Texto 2.5, PosLinea, .Fields("Alumno"), , 20, True
                TipoObjeto.FontBold = False
                PosLinea = PosLinea + 1
                PosLinea = PictPrint_Texto_Multiple(TipoObjeto, 2.5, PosLinea, "Estudiante del " & Cadena1 & ", resolvió declararla apta para presentarse a rendir los exámenes escritos de grado.", 17)
''                PosLinea = PosLinea + 0.6
''                PictPrint_Texto 2.5, PosLinea, "resolvió declararla apta para presentarse a rendir los exámenes escritos de grado."
                PosLinea = PosLinea + 1.2
                PictPrint_Texto 2.5, PosLinea, FechaStrgCiudad(MBFecha)
                PosLinea = PosLinea + 2.5
                PictPrint_Texto 2.5, PosLinea, "_______________________"
                PosLinea = PosLinea + 0.6
                PictPrint_Texto 2.5, PosLinea, Rector
                PosLinea = PosLinea + 0.6
                PictPrint_Texto 2.5, PosLinea, "      R E C T O R A"
          Case 3
                TipoObjeto.FontBold = False
                PictPrint_Texto 1.5, PosLinea, "CALIFICACIONES DE LOS EXAMENES ESCRITOS DE GRADO", , 20, True
                PosLinea = PosLinea + 1.2
                PictPrint_Texto 2.5, PosLinea, "Estudiante "
                TipoObjeto.FontBold = True
                PictPrint_Texto 2.5, PosLinea, .Fields("Alumno"), , 20, True
                TipoObjeto.FontBold = False
                PosLinea = PosLinea + 0.6
                PosLinea = PictPrint_Texto_Multiple(TipoObjeto, 2.5, PosLinea, "Del " & Cadena1, 17)
                PosLinea = PosLinea + 1.2
                PictPrint_Texto 3.2, PosLinea, "ASIGNATURAS"
                PictPrint_Texto 11.5, PosLinea, "NOTAS"
                PictPrint_Texto 14, PosLinea, "FIRMA DEL PROFESOR"
                PosLinea = PosLinea + 0.8
                TipoObjeto.Line (2.5, PosLinea)-(20, PosLinea), QBColor(Negro)
                InicioLinea = PosLinea
                PosLinea = PosLinea + 0.6
                Total = 0
                For IE = 0 To ContMatExam - 1
                    PictPrint_Texto 3, PosLinea, Materias_Examenes(IE).Materias
                    PictPrint_Texto 11.8, PosLinea, Format(Materias_Examenes(IE).Valor, "00.00")
                    PosLinea = PosLinea + 1
                    TipoObjeto.Line (2.5, PosLinea)-(20, PosLinea), QBColor(Negro)
                    UltimaLinea = PosLinea
                    PosLinea = PosLinea + 0.5
                    Total = Total + Materias_Examenes(IE).Valor
                Next IE
                TipoObjeto.Line (2.5, InicioLinea)-(2.5, UltimaLinea), QBColor(Negro)
                TipoObjeto.Line (11.3, InicioLinea)-(11.3, UltimaLinea), QBColor(Negro)
                TipoObjeto.Line (13.5, InicioLinea)-(13.5, UltimaLinea), QBColor(Negro)
                TipoObjeto.Line (20, InicioLinea)-(20, UltimaLinea), QBColor(Negro)
                PictPrint_Texto 2.5, PosLinea, "T O T A L"
                PictPrint_Texto 11.8, PosLinea, Format(Total, "00.00")
                PosLinea = PosLinea + 0.6
                PictPrint_Texto 2.5, PosLinea, "Promedio de Exámenes de Grado"
                PictPrint_Texto 11.8, PosLinea, Format(Total / ContMatExam, "00.00")
                PosLinea = PosLinea + 1.2
                TipoObjeto.FontBold = False
                PictPrint_Texto 2.5, PosLinea, FechaStrgCiudad(MBFecha)
                PosLinea = PosLinea + 2.5
                PictPrint_Texto 2.5, PosLinea, "____________________"
                PictPrint_Texto 14.5, PosLinea, "____________________"
                PosLinea = PosLinea + 0.6
                PictPrint_Texto 2.5, PosLinea, Rector
                PictPrint_Texto 14.5, PosLinea, Secretario2
                PosLinea = PosLinea + 0.6
                PictPrint_Texto 2.5, PosLinea, TextoRector
                PictPrint_Texto 14.5, PosLinea, TextoSecretario1
        End Select
     End If
 End If
End With
RatonNormal
MensajeEncabData = ""
End Sub

Public Sub Imprimir_Solicitudes_Examenes_Grado(EsProm As Boolean, Tipo_Solicitud As Byte)
Dim AnchoPict As Single
Dim AltoPict As Single
On Error GoTo Errorhandler
Mensajes = "Seguro de Imprimir en:" & vbCrLf & Printer.DeviceName & "?"
Titulo = "IMPRESION DE LIBRETAS"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
   RatonReloj
   Pagina = 1
   InicioX = 0
   InicioY = 0
   PictLibreta.Visible = False
   Escala_Centimetro 1, TipoTimes, 9
   AnchoPict = Round(Printer.ScaleWidth, 5)
   AltoPict = Round(Printer.ScaleHeight, 5)
   CodigoL = TVNivel.SelectedItem.key
   CodigoL = Mid$(CodigoL, 2, Len(CodigoL))
   With AdoAlumnos.Recordset
    If .RecordCount > 0 Then
       .MoveFirst
        Do While Not .EOF
           CodigoCliente = .Fields("Codigo")
           TxtCodigo = "(" & Pagina & ") " & CodigoCliente
           Listar_Solicitud_Examen_Grado Printer, CodigoCliente, Tipo_Solicitud
''           Printer.PaintPicture PictLibreta.Image, InicioX, InicioY, AnchoPict, AltoPict
           Printer.NewPage
           Pagina = Pagina + 1
          .MoveNext
        Loop
    End If
   End With
   RatonNormal
   MensajeEncabData = ""
   Printer.EndDoc
   PictLibreta.Visible = True
   Exit Sub
Errorhandler:
             PictLibreta.Visible = True
             RatonNormal
             ErrorDeImpresion
             Exit Sub
Else
   RatonNormal
End If
End Sub

Public Sub Listar_Notas_Examen_Grado(CodigoC As String, CountMatExam)
    If CountMatExam > 0 Then
       For IE = 0 To CountMatExam
           Materias_Examenes(IE).Valor = 0
       Next IE
       sSQL = "SELECT * " _
            & "FROM Trans_Notas_Grado " _
            & "WHERE Item = '" & NumEmpresa & "' " _
            & "AND Periodo = '" & Periodo_Contable & "' " _
            & "AND Codigo = '" & CodigoC & "' " _
            & "AND Mid$(CodE,1," & Len(CodigoL) & ") = '" & CodigoL & "' " _
            & "ORDER BY CodMat "
       SelectAdodc AdoAux, sSQL
       With AdoAux.Recordset
        If .RecordCount > 0 Then
            Do While Not .EOF
               For IE = 0 To CountMatExam
                   If Materias_Examenes(IE).CodigoMat = .Fields("CodMat") Then
                      Materias_Examenes(IE).Valor = Materias_Examenes(IE).Valor + .Fields("Examen")
                   End If
               Next IE
              .MoveNext
            Loop
        End If
       End With
   End If
End Sub

Public Sub Informe_Del_Alumno(TipoObjeto As Object, CodigoAlum As String)
Dim AnchoDib As Single
Dim AltoDib As Single
Dim Curso As String
Dim Alumno As String
Dim Paralelo As String
Dim PosXPict As Single
Dim AnchoPict() As CtasAsiento
Dim Y0 As Single
Dim y1 As Single
Dim X0 As Single
Dim x1 As Single
Dim PosLineaX As Single
Dim PosLineaXF As Single
Dim CanPromFinal As Byte
Dim Formato_Nota As String
Dim UnaVezOpta As Boolean
Dim EsPreBa As Boolean
Dim FinDeLibreta As Boolean

  Imp_Informe = False
  PosLinea = 0
  TipoLetra = TipoArial
  
  SQL1 = "SELECT CC.Curso,C.Sexo,CC.Descripcion As Paralelo,C.Cliente As Alumno," _
       & "CM.Materia,CM.SDiv,TN." & SQLInforme & ",TN.Id_No,TN.CodMatP,TN.Orden " _
       & "FROM Trans_Notas As TN,Catalogo_Materias As CM,Catalogo_Cursos As CC,Clientes As C " _
       & "WHERE TN.Item = '" & NumEmpresa & "' " _
       & "AND TN.Periodo = '" & Periodo_Contable & "' " _
       & "AND TN.Codigo = '" & CodigoAlum & "' " _
       & "AND LEN(TN." & SQLInforme & ") > 1 " _
       & "AND TN.Codigo = C.Codigo " _
       & "AND TN.CodE = CC.Curso " _
       & "AND TN.CodMat = CM.CodMat " _
       & "AND TN.Item = CC.Item " _
       & "AND TN.Item = CM.Item " _
       & "AND TN.Periodo = CC.Periodo " _
       & "AND TN.Periodo = CM.Periodo "
       
  SQL2 = "SELECT CC.Curso,C.Sexo,CC.Descripcion As Paralelo,C.Cliente As Alumno," _
       & "CM.Materia,CM.SDiv,TN." & SQLInforme & ",TN.Id_No,TN.CodMatP,TN.Orden " _
       & "FROM Trans_Notas_Auxiliares As TN,Catalogo_Materias As CM,Catalogo_Cursos As CC,Clientes As C " _
       & "WHERE TN.Item = '" & NumEmpresa & "' " _
       & "AND TN.Periodo = '" & Periodo_Contable & "' " _
       & "AND TN.Codigo = '" & CodigoAlum & "' " _
       & "AND LEN(TN." & SQLInforme & ") > 1 " _
       & "AND TN.Codigo = C.Codigo " _
       & "AND TN.CodE = CC.Curso " _
       & "AND TN.CodMat = CM.CodMat " _
       & "AND TN.Item = CC.Item " _
       & "AND TN.Item = CM.Item " _
       & "AND TN.Periodo = CC.Periodo " _
       & "AND TN.Periodo = CM.Periodo "
       
  sSQL = SQL1 & "UNION " & SQL2 _
       & "ORDER BY C.Cliente,TN.Id_No "
  SelectAdodc AdoNotasLibreta, sSQL
 If TypeOf TipoObjeto Is PictureBox Then TipoObjeto.Cls
 With AdoNotasLibreta.Recordset
  If .RecordCount > 0 Then
      Imp_Informe = True
      Curso = .Fields("Curso")
      Paralelo = .Fields("Paralelo")
      Alumno = .Fields("Alumno")
      NombreCliente = .Fields("Alumno")
      CadenaParcial = Visualizar_Notas_Periodo(LstPeriodos)
      AltoDib = PosLinea
    
    PictPrint_Tipo_Letra TipoArial, PorteLetra  'TipoHelvetica
    PictPrint_Color_Letra QBColor(Negro)
    PosLinea = 0.2
    PosColumna = 8
    JR = 1
     AnchoDib = 20
     cPrint.printImagen LogoTipo, 1, 0.5, 4.5, 2.25
     PictPrint_Color_Letra QBColor(Negro)
     PictPrint_Estilo_Letra FONT_BOLD, True
     PictPrint_Tipo_Letra TipoArial, 16
     PictPrint_Porte_Letra 16
     PictPrint_Texto 1, PosLinea, Institucion1, , 18, True
     PosLinea = PosLinea + 0.7
     PictPrint_Porte_Letra 14
     PictPrint_Texto 1, PosLinea, Institucion2, , 18, True
     PosLinea = PosLinea + 0.6
     PictPrint_Porte_Letra 8
     Cadena = "Teléfono: " & Telefono1
     PictPrint_Texto 1, PosLinea, Cadena, , 18, True
     PictPrint_Porte_Letra 10
     PosLinea = PosLinea + 0.5
     PictPrint_Texto 1, PosLinea, "Año Lectivo " & Anio_Lectivo, , 18, True
     PosLinea = PosLinea + 0.5
     PictPrint_Estilo_Letra FONT_BOLD, True
     PictPrint_Texto 1, PosLinea, "INFORME ACADÉMICO DEL ESTUDIANTE", , 18, True
     PosLinea = PosLinea + 0.4
     If OpcionNotas = 4 Then
        PictPrint_Texto 1, PosLinea + 0.05, UCase(CadenaParcial)
     Else
        Cadena = UCase(SinEspaciosIzqNoBlancos(CadenaParcial, 1) & " " & SinEspaciosIzqNoBlancos(CadenaParcial, 2))
        PictPrint_Texto 1, PosLinea + 0.05, Cadena
        If Mid$(Curso, 1, 4) > "1.01" Then
           Cadena = UCase(SinEspaciosIzqNoBlancos(CadenaParcial, 3) & " " & SinEspaciosIzqNoBlancos(CadenaParcial, 4))
           PictPrint_Texto 14.5, PosLinea + 0.05, Cadena
        End If
     End If
     PictPrint_Estilo_Letra FONT_BOLD, False
     PosLinea = PosLinea + 0.5
     PictPrint_Texto 1, PosLinea, Alumno
     PosLinea = PosLinea + 0.5
     PictPrint_Estilo_Letra FONT_BOLD, True
     PictPrint_Texto 1, PosLinea, Curso & " " & Paralelo
     PosLinea = PosLinea + 0.5
     PictPrint_Texto 1, PosLinea, "Docente Tutor:"
     PictPrint_Estilo_Letra FONT_BOLD, False
     PictPrint_Texto 3.5, PosLinea, ULCase(LblDirigente.Caption)
     PictPrint_Texto 13.1, PosLinea, FechaStrgCiudad(MBFecha)
     PosLinea = PosLinea + 0.5
     PFil = PosLinea
    'Cuadro Externo
     PictPrint_Porte_Letra 9
     PictPrint_Cuadro_Linea 1, PosLinea, 19, PosLinea + 0.45, QBColor(Negro), "B"
     PictPrint_Color_Letra QBColor(Negro)
     PictPrint_Texto 1, PosLinea + 0.05, "ÁMBITOS Y ASIGNATURAS", , PosColumna - 1, True
     PictPrint_Texto PosColumna, PosLinea + 0.05, "INFORME ACADEMICO", , PosColumna - 1, True
    'Imprimimos las columnas de las materias
     PosLinea = PosLinea + 0.6
     IR = PosColumna
      PosLineaX = PosLinea
      Do While Not .EOF
         If .Fields("Orden") <> 9 Then
             PictPrint_Porte_Letra 9
             If .Fields("CodMatP") <> Ninguno Then PictPrint_Porte_Letra 8
             PictPrint_Color_Letra QBColor(Negro)
             JR = 1
            'IMPRESION DE LAS NOTAS DE LAS MATERIAS
             PictPrint_Estilo_Letra FONT_BOLD, True
             PosLineaX = PosLinea
             Contador = Contador + 1
               If .Fields("CodMatP") <> Ninguno Then
                   cPrint.printImagen RutaSistema & "\ICONOS\vwicn115.ICO", 1.4, PosLinea + 0.1, 0.2, 0.2
                   PictPrint_Estilo_Letra FONT_UNDERLINE, True
                   PosLinea = PictPrint_Texto_Multiple(TipoObjeto, 1.8, PosLinea, .Fields("Materia"), 6)
                  'PictPrint_Texto 1.8, PosLinea, .Fields("Materia")
                   PictPrint_Estilo_Letra FONT_UNDERLINE, False
                   PictPrint_Estilo_Letra FONT_ITALIC, True
               ElseIf .Fields("SDiv") Then
                   cPrint.printImagen RutaSistema & "\ICONOS\Visto.ICO", 1.4, PosLinea + 0.1, 0.2, 0.2
                   PictPrint_Estilo_Letra FONT_UNDERLINE, True
                   PosLinea = PictPrint_Texto_Multiple(TipoObjeto, 1.8, PosLinea, .Fields("Materia"), 6)
                   'PictPrint_Texto 1.8, PosLinea, .Fields("Materia")
                   PictPrint_Estilo_Letra FONT_UNDERLINE, False
               Else
                   PosLinea = PictPrint_Texto_Multiple(TipoObjeto, 1.2, PosLinea, .Fields("Materia"), 7)
               End If
              If PosLineaX < PosLinea Then PosLinea = PosLinea - 0.4
              IR = PosColumna + 0.7
              PictPrint_Estilo_Letra FONT_NORMAL, False
              PosLinea = PictPrint_Texto_Multiple(TipoObjeto, IR, PosLinea, .Fields(SQLInforme), 10)
              PosLinea = PosLinea + 0.4
              PictPrint_Cuadro_Linea 1, PosLinea, 19.03, PosLinea, QBColor(Negro)
              PosLinea = PosLinea + 0.1
          End If
         .MoveNext
      Loop
      If (PosLinea - PFil) > 0 Then
         PictPrint_Cuadro_Linea 1, PFil, 19, PosLinea, QBColor(Negro), "B"
         PictPrint_Cuadro_Linea PosColumna + 0.5, PFil, 19, PosLinea, QBColor(Negro), "B"
      End If
      PosLineaX = PosLinea
      PosLinea = 25
      PictPrint_Porte_Letra 9
      PictPrint_Cuadro_Linea 1.5, PosLinea, 6.5, PosLinea, QBColor(Negro)
      If ("1.00" <= Codigo4) And (Codigo4 <= "3.99") Then
         PosLinea = PosLinea + 0.1
         PictPrint_Texto 1.5, PosLinea, ULCase(LblDirigente.Caption), , 5, True
         PosLinea = PosLinea + 0.4
         Select Case Codigo4
           Case "1.00" To "1.99"
                PictPrint_Texto 1.5, PosLinea, "PROFESOR(A)", , 5, True
           Case "2.00" To "3.99"
                PictPrint_Texto 1.5, PosLinea, "Docente Tutor", , 5, True
         End Select
         PictPrint_Estilo_Letra FONT_BOLD, False
      End If
      RatonNormal
      Cuadricula = False
      MensajeEncabData = "": SQLMsg1 = "": SQLMsg2 = "": SQLMsg3 = "": SQLMsg4 = ""
   End If
 End With
End Sub

Public Sub Encabezado_Materias_Aprovechamiento(TipoObjeto As Object, Pos_IR As Single, Pos_Fil As Single)
Dim IR As Single
    IR = Pos_IR
    If Mid$(FormatoLibreta, 1, 9) = "TRIMESTRE" Then
       Pos_IR = Pos_IR + 0.05
       cPrint.printTextoAngulo Pos_IR, Pos_Fil, 90, 4.5, 10, "SUMA DE"
       Pos_IR = Pos_IR + 0.25
       cPrint.printTextoAngulo Pos_IR, Pos_Fil, 90, 4.5, 10, "TRIMESTRES"
       Pos_IR = Pos_IR + 0.35
    Else
       Pos_IR = Pos_IR + 0.05
       cPrint.printTextoAngulo Pos_IR, Pos_Fil, 90, 4.5, 10, "PRIMER"
       Pos_IR = Pos_IR + 0.25
       cPrint.printTextoAngulo Pos_IR, Pos_Fil, 90, 4.5, 10, "QUIMESTRE"
       Pos_IR = Pos_IR + 0.35
       cPrint.printTextoAngulo Pos_IR, Pos_Fil, 90, 4.5, 10, "SEGUNDO"
       Pos_IR = Pos_IR + 0.25
       cPrint.printTextoAngulo Pos_IR, Pos_Fil, 90, 4.5, 10, "QUIMESTRE"
       Pos_IR = Pos_IR + 0.4
    End If
    Pos_IR = Pos_IR + 0.05
    cPrint.printTextoAngulo Pos_IR, Pos_Fil, 90, 4.5, 10, "PROMEDIO"
    Pos_IR = Pos_IR + 0.25
    cPrint.printTextoAngulo Pos_IR, Pos_Fil, 90, 4.5, 10, "GLOBAL"
    Pos_IR = Pos_IR + 0.45
    
    Pos_IR = Pos_IR + 0.2
    cPrint.printTextoAngulo Pos_IR, Pos_Fil, 90, 3.4, 14, "RECUPERACION"
    Pos_IR = Pos_IR + 0.5
    
    If Dato_Curso.CantNotas > 5 Then
        cPrint.printTextoAngulo Pos_IR, Pos_Fil, 90, 4.5, 10, "EXAMEN"
        Pos_IR = Pos_IR + 0.25
        cPrint.printTextoAngulo Pos_IR, Pos_Fil, 90, 4.5, 10, "SUPLETORIO"
        Pos_IR = Pos_IR + 0.45
        cPrint.printTextoAngulo Pos_IR, Pos_Fil, 90, 4.5, 10, "EXAMEN"
        Pos_IR = Pos_IR + 0.25
        cPrint.printTextoAngulo Pos_IR, Pos_Fil, 90, 4.5, 10, "REMEDIAL"
        Pos_IR = Pos_IR + 0.45
        cPrint.printTextoAngulo Pos_IR, Pos_Fil, 90, 4.5, 10, "EXAMEN"
        Pos_IR = Pos_IR + 0.25
        cPrint.printTextoAngulo Pos_IR, Pos_Fil, 90, 4.5, 10, "GRACIA"
        Pos_IR = Pos_IR + 0.45
    End If
    cPrint.printTextoAngulo Pos_IR, Pos_Fil, 90, 4.5, 10, "PROMEDIO"
    Pos_IR = Pos_IR + 0.25
    cPrint.printTextoAngulo Pos_IR, Pos_Fil, 90, 4.5, 10, "FINAL"
    
    For J = 1 To Dato_Curso.CantNotas
        IR = IR + 0.7
        TipoObjeto.Line (IR, Pos_Fil - 1.7)-(IR, Pos_Fil + 0.1), QBColor(Negro)
    Next J
End Sub

Public Sub PDF_Encabezado_Materias_Aprovechamiento(TipoObjeto As Object, Pos_IR As Single, Pos_Fil As Single)
Dim IR As Single
    TipoObjeto.PDFSetFontSize 5
    IR = Pos_IR
   'MsgBox Pos_IR & " - " & Pos_Fil
    If Mid$(FormatoLibreta, 1, 9) = "TRIMESTRE" Then
       Pos_IR = Pos_IR + 0.15
       cPrint.printTextoAngulo Pos_IR, Pos_Fil, 90, 0, 0, "SUMA DE"
       Pos_IR = Pos_IR + 0.2
       cPrint.printTextoAngulo Pos_IR, Pos_Fil, 90, 0, 0, "TRIMESTRES"
       Pos_IR = Pos_IR + 0.3
    Else
       Pos_IR = Pos_IR + 0.15
       cPrint.printTextoAngulo Pos_IR, Pos_Fil, 90, 0, 0, "PRIMER"
       Pos_IR = Pos_IR + 0.25
       cPrint.printTextoAngulo Pos_IR, Pos_Fil, 90, 0, 0, "QUIMESTRE"
       Pos_IR = Pos_IR + 0.35
       cPrint.printTextoAngulo Pos_IR, Pos_Fil, 90, 0, 0, "SEGUNDO"
       Pos_IR = Pos_IR + 0.25
       cPrint.printTextoAngulo Pos_IR, Pos_Fil, 90, 0, 0, "QUIMESTRE"
       Pos_IR = Pos_IR + 0.3
    End If
    Pos_IR = Pos_IR + 0.05
    cPrint.printTextoAngulo Pos_IR, Pos_Fil, 90, 0, 0, "PROMEDIO"
    Pos_IR = Pos_IR + 0.25
    cPrint.printTextoAngulo Pos_IR, Pos_Fil, 90, 0, 0, "GLOBAL"
    Pos_IR = Pos_IR + 0.3
    
    Pos_IR = Pos_IR + 0.2
    cPrint.printTextoAngulo Pos_IR, Pos_Fil, 90, 0, 0, "RECUPERACION"
    Pos_IR = Pos_IR + 0.45
    
    If Dato_Curso.CantNotas > 5 Then
        cPrint.printTextoAngulo Pos_IR, Pos_Fil, 90, 0, 0, "EXAMEN"
        Pos_IR = Pos_IR + 0.25
        cPrint.printTextoAngulo Pos_IR, Pos_Fil, 90, 0, 0, "SUPLETORIO"
        Pos_IR = Pos_IR + 0.35
        cPrint.printTextoAngulo Pos_IR, Pos_Fil, 90, 0, 0, "EXAMEN"
        Pos_IR = Pos_IR + 0.25
        cPrint.printTextoAngulo Pos_IR, Pos_Fil, 90, 0, 0, "REMEDIAL"
        Pos_IR = Pos_IR + 0.35
        cPrint.printTextoAngulo Pos_IR, Pos_Fil, 90, 0, 0, "EXAMEN"
        Pos_IR = Pos_IR + 0.25
        cPrint.printTextoAngulo Pos_IR, Pos_Fil, 90, 0, 0, "GRACIA"
        Pos_IR = Pos_IR + 0.35
    End If
    cPrint.printTextoAngulo Pos_IR, Pos_Fil, 90, 0, 0, "PROMEDIO"
    Pos_IR = Pos_IR + 0.2
    cPrint.printTextoAngulo Pos_IR, Pos_Fil, 90, 0, 0, "FINAL"
    IR = IR - 0.1
    For J = 1 To Dato_Curso.CantNotas
        PictPrint_Cuadro_Linea IR, PosLinea - 0.3, IR + 0.6, PosLinea + 1.2, QBColor(Negro), "B"
        IR = IR + 0.6
    Next J
End Sub

Public Sub Notas_Materias_Aprovechamiento(TipoObjeto As Object, Pos_IR As Single, Pos_Fil As Single, AdoNotas As Adodc)
Dim IR_Temp As Single
   IR_Temp = Pos_IR
   With AdoNotas.Recordset
        If .Fields("PromPQ") > 0 And .Fields("PromSQ") <= 0 Then
            Abono_ME = .Fields("PromPQ")
        ElseIf .Fields("PromPQ") <= 0 And .Fields("PromSQ") > 0 Then
            Abono_ME = .Fields("PromSQ")
        Else
            If Mid$(FormatoLibreta, 1, 9) = "TRIMESTRE" Then
               Abono_ME = (.Fields("PromPQ") + .Fields("PromSQ") + .Fields("PromTQ")) / 3
            Else
               Abono_ME = (.Fields("PromPQ") + .Fields("PromSQ")) / 2
            End If
        End If
        Abono_ME = Redondear_2Dec(Abono_ME)
        If .Fields("C2") Then
            PictPrint_Nota_Materia Pos_IR, PosLinea, .Fields("ExamenPQ"), .Fields("C"), Dec_Nota, , .Fields("C2")
            Pos_IR = Pos_IR + 0.6
            PictPrint_Nota_Materia Pos_IR, PosLinea, .Fields("ExamenSQ"), .Fields("C"), Dec_Nota, , .Fields("C2")
            Pos_IR = Pos_IR + 0.6
        Else
            PictPrint_Nota_Materia Pos_IR, PosLinea, .Fields("PromPQ"), .Fields("C"), Dec_Nota, , .Fields("C2")
            Pos_IR = Pos_IR + 0.6
            PictPrint_Nota_Materia Pos_IR, PosLinea, .Fields("PromSQ"), .Fields("C"), Dec_Nota, , .Fields("C2")
            Pos_IR = Pos_IR + 0.6
        End If
        PictPrint_Nota_Materia Pos_IR, PosLinea, Abono_ME, .Fields("C"), Dec_Nota, , .Fields("C2")
        Pos_IR = Pos_IR + 0.6
       'PictPrint_Nota_Materia Pos_IR, PosLinea, .Fields("Recuperacion"), False, Dec_Nota
        Pos_IR = Pos_IR + 0.6
        If Dato_Curso.CantNotas > 5 Then
           PictPrint_Nota_Materia Pos_IR, PosLinea, .Fields("Supletorio"), .Fields("C"), Dec_Nota, , .Fields("C2")
           Pos_IR = Pos_IR + 0.6
           PictPrint_Nota_Materia Pos_IR, PosLinea, .Fields("Remedial"), .Fields("C"), Dec_Nota, , .Fields("C2")
           Pos_IR = Pos_IR + 0.6
          'PictPrint_Nota_Materia Pos_IR, PosLinea, .Fields("Gracia"), False, Dec_Nota
           Pos_IR = Pos_IR + 0.6
        End If
        PictPrint_Nota_Materia Pos_IR, PosLinea, .Fields("PromFinal"), .Fields("C"), Dec_Nota, , .Fields("C2")
        Pos_IR = Pos_IR + 0.6
        If TypeOf TipoObjeto Is mjwPDF Then    ' Si la Impresion es en PDF
        
        Else
           TipoObjeto.Line (IR_Temp, PosLinea - 0.05)-(Pos_IR, PosLinea - 0.05), QBColor(Negro)
           IR_Temp = IR_Temp - 0.05
           Do While IR_Temp <= Pos_IR
              TipoObjeto.Line (IR_Temp, PosLinea - 0.05)-(IR_Temp, PosLinea + 0.35), QBColor(Negro)
              IR_Temp = IR_Temp + 0.6
           Loop
        End If
   End With
End Sub

Public Sub PDF_Nomina_Representante(Curso As String)
Dim Logo1 As String
Dim IR As Single
RatonReloj
  sSQL = "SELECT CM.Matricula_No, CM.CI As Cedula, C.Cliente As Estudiante, CM.Representante_Alumno," _
       & "CM.CI_R As Cedula_Rep, CM.Telefono_R, C.Celular, CM.Domicilio As Direccion_Rep,C.Codigo " _
       & "FROM Clientes As C,Clientes_Matriculas As CM " _
       & "WHERE CM.Item = '" & NumEmpresa & "' " _
       & "AND CM.Periodo = '" & Periodo_Contable & "' " _
       & "AND CM.Grupo_No = '" & Curso & "' " _
       & "AND C.Codigo = CM.Codigo " _
       & "ORDER BY C.Cliente,C.Sexo "
  SelectAdodc AdoAlumnos, sSQL
  Progreso_Barra.Incremento = 0
  Progreso_Barra.Valor_Maximo = AdoAlumnos.Recordset.RecordCount
  Progreso_Barra.Mensaje_Box = "Generando Lista"
  Progreso_Esperar

Set ObjPDF = New mjwPDF
ObjPDF.PDFTitle = "Lista de Estudiantes"
ObjPDF.PDFFileName = RutaSysBases & "\TEMP\Nomina_Representante_del_" & Replace(Dato_Curso.Curso, ".", "-") & ".PDF"
ObjPDF.PDFLoadAfm = RutaSistema & "\FONTSPDF"
ObjPDF.PDFSetUnit = UNIT_CM
ObjPDF.PDFSetLayoutMode = LAYOUT_DEFAULT
ObjPDF.PDFFormatPage = FORMAT_A4
ObjPDF.PDFOrientation = ORIENT_PAYSAGE 'ORIENT_PORTRAIT
ObjPDF.PDFView = True
ObjPDF.PDFBeginDoc
ObjPDF.PDFSetBookmark " "
ObjPDF.PDFSetFontName FONT_HELVETICA
    ObjPDF.PDFSetFontSize 10
    PosLinea = 1.5
    Pagina = 1
    Logo1 = RutaSistema & "\LOGOS\MINISEDU.JPG"
    PictPrint_Grafico ObjPDF, Logo1, 1.3, PosLinea, 2, 1.3
    PictPrint_Grafico ObjPDF, LogoTipo, 27, PosLinea, 2, 1.3
    ObjPDF.PDFSetFontSize 10
    PosLinea = PosLinea + 0.2
    PictPrint_Texto ObjPDF, 1, PosLinea, UCase$(Institucion1), , 28, True
    PosLinea = PosLinea + 0.5
    PictPrint_Texto ObjPDF, 1, PosLinea, UCase$(Institucion2), , 28, True
    PosLinea = PosLinea + 0.4
    ObjPDF.PDFSetFontSize 8
    PictPrint_Texto ObjPDF, 1, PosLinea, "AÑO LECTIVO: " & Anio_Lectivo, , 28, True
    PosLinea = PosLinea + 0.6
    PictPrint_Texto ObjPDF, 1, PosLinea, "CURSO: " & Dato_Curso.Bachiller
    PictPrint_Texto ObjPDF, 16, PosLinea, "PARALELO: " & Dato_Curso.Paralelo
    PictPrint_Texto ObjPDF, 20.7, PosLinea, "FECHA MATRICULA: " & FechaStrg(MBFecha)
    PosLinea = PosLinea + 0.6
    
    ObjPDF.PDFSetFontSize 8
    PrimeraLinea = PosLinea - 0.4
    PictPrint_Cuadro_Linea ObjPDF, 1, PosLinea - 0.4, 29, PosLinea - 0.4, QBColor(Negro)
    PictPrint_Cuadro_Linea ObjPDF, 1, PosLinea + 0.01, 29, PosLinea + 0.46, QBColor(Blanco), "BF"
    PictPrint_Texto ObjPDF, 1.05, PosLinea, "No"
    PictPrint_Texto ObjPDF, 1.55, PosLinea, "Matric."
    PictPrint_Texto ObjPDF, 2.5, PosLinea, "CEDULA"
    PictPrint_Texto ObjPDF, 4.4, PosLinea, "ESTUDIANTE"
    PictPrint_Texto ObjPDF, 11.5, PosLinea, "REPRESENTANTE DEL ESTUDIANTE"
    PictPrint_Texto ObjPDF, 17.2, PosLinea, "CEDULA R."
    PictPrint_Texto ObjPDF, 19, PosLinea, "TELEFONO"
    PictPrint_Texto ObjPDF, 20.7, PosLinea, "CELULAR"
    PictPrint_Texto ObjPDF, 22.5, PosLinea, "DIRECCION"
    PosLinea = PosLinea + 0.5
    Contador = 0
    With AdoAlumnos.Recordset
     If .RecordCount > 0 Then
         Codigo1 = Leer_Datos_del_Curso(Curso)
         Codigo1 = Dato_Curso.Especialidad
         Do While Not .EOF
            Progreso_Barra.Mensaje_Box = .Fields("Estudiante")
            Progreso_Esperar
            PictPrint_Cuadro_Linea ObjPDF, 1, PosLinea - 0.4, 29, PosLinea - 0.4, QBColor(Negro)
            Contador = Contador + 1
            PictPrint_Texto ObjPDF, 1.05, PosLinea, Format(Contador, "00")
            PictPrint_Texto ObjPDF, 1.5, PosLinea, .Fields("Matricula_No")
            PictPrint_Texto ObjPDF, 2.5, PosLinea, .Fields("Cedula")
            PictPrint_Texto ObjPDF, 4.4, PosLinea, .Fields("Estudiante")
            PictPrint_Texto ObjPDF, 11.5, PosLinea, ULCase(.Fields("Representante_Alumno"))
            PictPrint_Texto ObjPDF, 17.2, PosLinea, .Fields("Cedula_Rep")
            PictPrint_Texto ObjPDF, 19, PosLinea, .Fields("Telefono_R")
            PictPrint_Texto ObjPDF, 20.7, PosLinea, .Fields("Celular")
            PictPrint_Texto ObjPDF, 22.5, PosLinea, ULCase(.Fields("Direccion_Rep"))
            PosLinea = PosLinea + 0.5
            If PosLinea >= 20 Then
               PictPrint_Texto ObjPDF, 1.05, PosLinea, "Página No. " & Pagina
               PictPrint_Cuadro_Linea ObjPDF, 1, PosLinea - 0.4, 29, PosLinea - 0.4, QBColor(Negro)
               PictPrint_Cuadro_Linea ObjPDF, 1, PrimeraLinea, 1, PosLinea - 0.4, QBColor(Negro)
               PictPrint_Cuadro_Linea ObjPDF, 1.45, PrimeraLinea, 1.45, PosLinea - 0.4, QBColor(Negro)
               PictPrint_Cuadro_Linea ObjPDF, 2.45, PrimeraLinea, 2.45, PosLinea - 0.4, QBColor(Negro)
               PictPrint_Cuadro_Linea ObjPDF, 4.35, PrimeraLinea, 4.35, PosLinea - 0.4, QBColor(Negro)
               PictPrint_Cuadro_Linea ObjPDF, 11.45, PrimeraLinea, 11.45, PosLinea - 0.4, QBColor(Negro)
               PictPrint_Cuadro_Linea ObjPDF, 17.15, PrimeraLinea, 17.15, PosLinea - 0.4, QBColor(Negro)
               PictPrint_Cuadro_Linea ObjPDF, 18.95, PrimeraLinea, 18.95, PosLinea - 0.4, QBColor(Negro)
               PictPrint_Cuadro_Linea ObjPDF, 20.65, PrimeraLinea, 20.65, PosLinea - 0.4, QBColor(Negro)
               PictPrint_Cuadro_Linea ObjPDF, 22.45, PrimeraLinea, 22.45, PosLinea - 0.4, QBColor(Negro)
               PictPrint_Cuadro_Linea ObjPDF, 29, PrimeraLinea, 29, PosLinea - 0.4, QBColor(Negro)
               ObjPDF.PDFEndPage
               ObjPDF.PDFNewPage
               PosLinea = 1.5
               PrimeraLinea = PosLinea - 0.4
               Pagina = Pagina + 1
            End If
           .MoveNext
         Loop
     End If
    End With
    PictPrint_Texto ObjPDF, 1.05, PosLinea, "Página No. " & Pagina
    PictPrint_Cuadro_Linea ObjPDF, 1, PosLinea - 0.4, 29, PosLinea - 0.4, QBColor(Negro)
    PictPrint_Cuadro_Linea ObjPDF, 1, PrimeraLinea, 1, PosLinea - 0.4, QBColor(Negro)
    PictPrint_Cuadro_Linea ObjPDF, 1.45, PrimeraLinea, 1.45, PosLinea - 0.4, QBColor(Negro)
    PictPrint_Cuadro_Linea ObjPDF, 2.45, PrimeraLinea, 2.45, PosLinea - 0.4, QBColor(Negro)
    PictPrint_Cuadro_Linea ObjPDF, 4.35, PrimeraLinea, 4.35, PosLinea - 0.4, QBColor(Negro)
    PictPrint_Cuadro_Linea ObjPDF, 11.45, PrimeraLinea, 11.45, PosLinea - 0.4, QBColor(Negro)
    PictPrint_Cuadro_Linea ObjPDF, 17.15, PrimeraLinea, 17.15, PosLinea - 0.4, QBColor(Negro)
    PictPrint_Cuadro_Linea ObjPDF, 18.95, PrimeraLinea, 18.95, PosLinea - 0.4, QBColor(Negro)
    PictPrint_Cuadro_Linea ObjPDF, 20.65, PrimeraLinea, 20.65, PosLinea - 0.4, QBColor(Negro)
    PictPrint_Cuadro_Linea ObjPDF, 22.45, PrimeraLinea, 22.45, PosLinea - 0.4, QBColor(Negro)
    PictPrint_Cuadro_Linea ObjPDF, 29, PrimeraLinea, 29, PosLinea - 0.4, QBColor(Negro)
    PictPrint_Texto ObjPDF, 22.5, PosLinea, FechaStrgCiudad(FechaSistema)
    PosLinea = PosLinea + 1.5
    PictPrint_Texto ObjPDF, 4.4, PosLinea, Rector
    PictPrint_Texto ObjPDF, 17.2, PosLinea, Secretario1
    PosLinea = PosLinea + 0.4
    PictPrint_Texto ObjPDF, 4.4, PosLinea, TextoRector
    PictPrint_Texto ObjPDF, 17.2, PosLinea, TextoSecretario1
'Fin del PDF
ObjPDF.PDFEndPage
ObjPDF.PDFEndDoc
Progreso_Barra.Incremento = Progreso_Barra.Valor_Maximo
Progreso_Esperar
RatonNormal
GenerarDataTexto FLibretas, AdoAlumnos
RatonNormal
End Sub

Public Sub PDF_Nomina_Representante_Email(Curso As String)
Dim Logo1 As String
Dim IR As Single
RatonReloj
  sSQL = "SELECT CM.CI As Cedula, C.Cliente As Estudiante, CM.Fecha_N, CM.Representante_Alumno," _
       & "CM.CI_R As Cedula_Rep, C.Celular, CM.Email_R As Correo_Representante " _
       & "FROM Clientes As C,Clientes_Matriculas As CM " _
       & "WHERE CM.Item = '" & NumEmpresa & "' " _
       & "AND CM.Periodo = '" & Periodo_Contable & "' " _
       & "AND CM.Grupo_No = '" & Curso & "' " _
       & "AND C.Codigo = CM.Codigo " _
       & "ORDER BY C.Cliente,C.Sexo "
  'MsgBox sSQL
  SelectAdodc AdoAlumnos, sSQL
  Progreso_Barra.Incremento = 0
  Progreso_Barra.Valor_Maximo = AdoAlumnos.Recordset.RecordCount
  Progreso_Barra.Mensaje_Box = "Generando Lista"
  Progreso_Esperar

  
Set ObjPDF = New mjwPDF
ObjPDF.PDFTitle = "Lista de Estudiantes"
ObjPDF.PDFFileName = RutaSysBases & "\TEMP\Nomina_Representante_email_del_" & Replace(Dato_Curso.Curso, ".", "-") & ".PDF"
ObjPDF.PDFLoadAfm = RutaSistema & "\FONTSPDF"
ObjPDF.PDFSetUnit = UNIT_CM
ObjPDF.PDFSetLayoutMode = LAYOUT_DEFAULT
ObjPDF.PDFFormatPage = FORMAT_A4
ObjPDF.PDFOrientation = ORIENT_PAYSAGE 'ORIENT_PORTRAIT
ObjPDF.PDFView = True
ObjPDF.PDFBeginDoc
ObjPDF.PDFSetBookmark " "
ObjPDF.PDFSetFontName FONT_HELVETICA
    ObjPDF.PDFSetFontSize 10
    PosLinea = 1.5
    Pagina = 1
    Logo1 = RutaSistema & "\LOGOS\MINISEDU.JPG"
    PictPrint_Grafico ObjPDF, Logo1, 1.3, PosLinea, 2, 1.3
    PictPrint_Grafico ObjPDF, LogoTipo, 27, PosLinea, 2, 1.3
    ObjPDF.PDFSetFontSize 10
    PosLinea = PosLinea + 0.2
    PictPrint_Texto ObjPDF, 1, PosLinea, UCase$(Institucion1), , 28, True
    PosLinea = PosLinea + 0.5
    PictPrint_Texto ObjPDF, 1, PosLinea, UCase$(Institucion2), , 28, True
    PosLinea = PosLinea + 0.4
    ObjPDF.PDFSetFontSize 8
    PictPrint_Texto ObjPDF, 1, PosLinea, "AÑO LECTIVO: " & Anio_Lectivo, , 28, True
    PosLinea = PosLinea + 0.6
    PictPrint_Texto ObjPDF, 1, PosLinea, "CURSO: " & Dato_Curso.Bachiller
    PictPrint_Texto ObjPDF, 16, PosLinea, "PARALELO: " & Dato_Curso.Paralelo
    PictPrint_Texto ObjPDF, 20.7, PosLinea, "FECHA MATRICULA: " & FechaStrg(MBFecha)
    PosLinea = PosLinea + 0.6
    
    ObjPDF.PDFSetFontSize 8
    PrimeraLinea = PosLinea - 0.4
    PictPrint_Cuadro_Linea ObjPDF, 1, PosLinea - 0.4, 29, PosLinea - 0.4, QBColor(Negro)
    PictPrint_Cuadro_Linea ObjPDF, 1, PosLinea + 0.01, 29, PosLinea + 0.46, QBColor(Blanco), "BF"
    PictPrint_Texto ObjPDF, 1.05, PosLinea, "No"
    PictPrint_Texto ObjPDF, 1.55, PosLinea, "CEDULA"
    PictPrint_Texto ObjPDF, 3.3, PosLinea, "ESTUDIANTE"
    PictPrint_Texto ObjPDF, 10.5, PosLinea, "FECHA N."
    PictPrint_Texto ObjPDF, 12, PosLinea, "REPRESENTANTE DEL ESTUDIANTE"
    PictPrint_Texto ObjPDF, 19, PosLinea, "CEDULA R."
    PictPrint_Texto ObjPDF, 20.7, PosLinea, "CELULAR"
    PictPrint_Texto ObjPDF, 22.5, PosLinea, "CORREO ELECTRONICO"
    PosLinea = PosLinea + 0.5
    Contador = 0
    With AdoAlumnos.Recordset
     If .RecordCount > 0 Then
         Codigo1 = Leer_Datos_del_Curso(Curso)
         Codigo1 = Dato_Curso.Especialidad
         Do While Not .EOF
            Progreso_Barra.Mensaje_Box = .Fields("Estudiante")
            Progreso_Esperar
            PictPrint_Cuadro_Linea ObjPDF, 1, PosLinea - 0.4, 29, PosLinea - 0.4, QBColor(Negro)
            Contador = Contador + 1
            PictPrint_Texto ObjPDF, 1.05, PosLinea, Format(Contador, "00")
            PictPrint_Texto ObjPDF, 1.5, PosLinea, .Fields("Cedula")
            PictPrint_Texto ObjPDF, 3.3, PosLinea, .Fields("Estudiante")
            PictPrint_Texto ObjPDF, 10.5, PosLinea, .Fields("Fecha_N")
            PictPrint_Texto ObjPDF, 12, PosLinea, .Fields("Representante_Alumno")
            PictPrint_Texto ObjPDF, 19, PosLinea, .Fields("Cedula_Rep")
            PictPrint_Texto ObjPDF, 20.7, PosLinea, .Fields("Celular")
            PictPrint_Texto ObjPDF, 22.5, PosLinea, .Fields("Correo_Representante")
            PosLinea = PosLinea + 0.5
            If PosLinea >= 20 Then
               PictPrint_Texto ObjPDF, 1.05, PosLinea, "Página No. " & Pagina
               PictPrint_Cuadro_Linea ObjPDF, 1, PosLinea - 0.4, 29, PosLinea - 0.4, QBColor(Negro)
               PictPrint_Cuadro_Linea ObjPDF, 1, PrimeraLinea, 1, PosLinea - 0.4, QBColor(Negro)
               PictPrint_Cuadro_Linea ObjPDF, 1.45, PrimeraLinea, 1.45, PosLinea - 0.4, QBColor(Negro)
               PictPrint_Cuadro_Linea ObjPDF, 3.25, PrimeraLinea, 3.25, PosLinea - 0.4, QBColor(Negro)
               PictPrint_Cuadro_Linea ObjPDF, 10.45, PrimeraLinea, 10.45, PosLinea - 0.4, QBColor(Negro)
               PictPrint_Cuadro_Linea ObjPDF, 11.95, PrimeraLinea, 11.95, PosLinea - 0.4, QBColor(Negro)
               PictPrint_Cuadro_Linea ObjPDF, 18.95, PrimeraLinea, 18.95, PosLinea - 0.4, QBColor(Negro)
               PictPrint_Cuadro_Linea ObjPDF, 20.65, PrimeraLinea, 20.65, PosLinea - 0.4, QBColor(Negro)
               PictPrint_Cuadro_Linea ObjPDF, 22.45, PrimeraLinea, 22.45, PosLinea - 0.4, QBColor(Negro)
               PictPrint_Cuadro_Linea ObjPDF, 29, PrimeraLinea, 29, PosLinea - 0.4, QBColor(Negro)
               ObjPDF.PDFEndPage
               ObjPDF.PDFNewPage
               PosLinea = 1.5
               PrimeraLinea = PosLinea - 0.4
               Pagina = Pagina + 1
            End If
           .MoveNext
         Loop
     End If
    End With
    PictPrint_Texto ObjPDF, 1.05, PosLinea, "Página No. " & Pagina
    PictPrint_Cuadro_Linea ObjPDF, 1, PosLinea - 0.4, 29, PosLinea - 0.4, QBColor(Negro)
    PictPrint_Cuadro_Linea ObjPDF, 1, PrimeraLinea, 1, PosLinea - 0.4, QBColor(Negro)
    PictPrint_Cuadro_Linea ObjPDF, 1.45, PrimeraLinea, 1.45, PosLinea - 0.4, QBColor(Negro)
    PictPrint_Cuadro_Linea ObjPDF, 3.25, PrimeraLinea, 3.25, PosLinea - 0.4, QBColor(Negro)
    PictPrint_Cuadro_Linea ObjPDF, 10.45, PrimeraLinea, 10.45, PosLinea - 0.4, QBColor(Negro)
    PictPrint_Cuadro_Linea ObjPDF, 11.95, PrimeraLinea, 11.95, PosLinea - 0.4, QBColor(Negro)
    PictPrint_Cuadro_Linea ObjPDF, 18.95, PrimeraLinea, 18.95, PosLinea - 0.4, QBColor(Negro)
    PictPrint_Cuadro_Linea ObjPDF, 20.65, PrimeraLinea, 20.65, PosLinea - 0.4, QBColor(Negro)
    PictPrint_Cuadro_Linea ObjPDF, 22.45, PrimeraLinea, 22.45, PosLinea - 0.4, QBColor(Negro)
    PictPrint_Cuadro_Linea ObjPDF, 29, PrimeraLinea, 29, PosLinea - 0.4, QBColor(Negro)
    PictPrint_Texto ObjPDF, 22.5, PosLinea, FechaStrgCiudad(FechaSistema)
    PosLinea = PosLinea + 1.5
    PictPrint_Texto ObjPDF, 4.4, PosLinea, Rector
    PictPrint_Texto ObjPDF, 17.2, PosLinea, Secretario1
    PosLinea = PosLinea + 0.4
    PictPrint_Texto ObjPDF, 4.4, PosLinea, TextoRector
    PictPrint_Texto ObjPDF, 17.2, PosLinea, TextoSecretario1
'Fin del PDF
ObjPDF.PDFEndPage
ObjPDF.PDFEndDoc
Progreso_Barra.Incremento = Progreso_Barra.Valor_Maximo
Progreso_Esperar
RatonNormal
GenerarDataTexto FLibretas, AdoAlumnos
RatonNormal
End Sub

Public Sub PDF_Lista_Estudiantes()
Dim Logo1 As String
Dim IR As Single
RatonReloj
Set ObjPDF = New mjwPDF
ObjPDF.PDFTitle = "Lista de Estudiantes"
ObjPDF.PDFFileName = RutaSysBases & "\TEMP\Lista_Estudiantes_del_" & Replace(Dato_Curso.Curso, ".", "-") & ".PDF"
ObjPDF.PDFLoadAfm = RutaSistema & "\FONTSPDF"
ObjPDF.PDFSetUnit = UNIT_CM
ObjPDF.PDFSetLayoutMode = LAYOUT_DEFAULT
ObjPDF.PDFFormatPage = FORMAT_A4
ObjPDF.PDFOrientation = ORIENT_PORTRAIT
ObjPDF.PDFView = True
ObjPDF.PDFBeginDoc
ObjPDF.PDFSetBookmark " "
ObjPDF.PDFSetFontName FONT_TIMES
    Progreso_Barra.Incremento = 0
    Progreso_Barra.Valor_Maximo = Dato_Curso.ContAlumnos
    Progreso_Barra.Mensaje_Box = "LISTA DE ALUMNOS"
    Progreso_Esperar
    PosLinea = 1.5
    Logo1 = RutaSistema & "\LOGOS\MINISEDU.JPG"
    PictPrint_Grafico ObjPDF, Logo1, 1.3, PosLinea, 2.5, 1.7
    PictPrint_Grafico ObjPDF, LogoTipo, 18, PosLinea, 2.4, 1.7
    ObjPDF.PDFSetFontSize 14
    PosLinea = PosLinea + 0.2
    PictPrint_Texto ObjPDF, 1, PosLinea, UCase$(Institucion1), , 18.5, True
    PosLinea = PosLinea + 0.6
    PictPrint_Texto ObjPDF, 1, PosLinea, UCase$(Institucion2), , 18.5, True
    PosLinea = PosLinea + 0.5
    ObjPDF.PDFSetFontSize 8
    PictPrint_Texto ObjPDF, 1, PosLinea, "AÑO LECTIVO: " & Anio_Lectivo, , 18.5, True
    PosLinea = PosLinea + 0.5
    ObjPDF.PDFSetFontSize 11
    PictPrint_Texto ObjPDF, 1, PosLinea, TxtTitulo, , 18.5, True
    PosLinea = PosLinea + 0.5
    ObjPDF.PDFSetFontSize 8
    PictPrint_Texto ObjPDF, 1, PosLinea, "TUTOR: " & LblDirigente.Caption
    PictPrint_Texto ObjPDF, 15.5, PosLinea, "DIA: " & FechaStrg(MBFecha)
    PosLinea = PosLinea + 0.4
    PictPrint_Texto ObjPDF, 1, PosLinea, "CURSO: " & Dato_Curso.Bachiller
    PictPrint_Texto ObjPDF, 15.5, PosLinea, "PARALELO: " & Dato_Curso.Paralelo
    PosLinea = PosLinea + 0.5
    ObjPDF.PDFSetFontSize 10
    PrimeraLinea = PosLinea
    PictPrint_Cuadro_Linea ObjPDF, 1, PosLinea + 0.03, 20, PosLinea + 0.46, QBColor(Blanco), "BF"
    PictPrint_Texto ObjPDF, 1.5, PosLinea, "A P E L L I D O S   Y   N O M B R E S"
    PictPrint_Texto ObjPDF, 11.2, PosLinea, TxtObservacion
    PosLinea = PosLinea + 0.5
    For I = 1 To Dato_Curso.ContAlumnos
        PictPrint_Cuadro_Linea ObjPDF, 1, PosLinea - 0.4, 20, PosLinea - 0.4, QBColor(Negro)
        Progreso_Barra.Mensaje_Box = Dato_Curso.Alumno(I)
        Progreso_Esperar
        PosLinea = PosLinea + 0.3
        PictPrint_Texto ObjPDF, 1.05, PosLinea, Format(I, "00")
        PictPrint_Texto ObjPDF, 1.6, PosLinea, Dato_Curso.Alumno(I)
        PosLinea = PosLinea + 0.8
        If PosLinea >= 29 Then
           PictPrint_Cuadro_Linea ObjPDF, 1, PrimeraLinea, 20, PosLinea, QBColor(Negro), "B"
           PictPrint_Cuadro_Linea ObjPDF, 1.5, PrimeraLinea + 0.5, 1.5, PosLinea, QBColor(Negro), "B"
           PictPrint_Cuadro_Linea ObjPDF, 1, PosLinea - 0.4, 20, PosLinea - 0.4, QBColor(Negro)
           PictPrint_Cuadro_Linea ObjPDF, 11, PrimeraLinea - 0.4, 11, PosLinea - 0.4, QBColor(Negro)
           PosLinea = 2
           PrimeraLinea = PosLinea
           ObjPDF.PDFEndPage
           ObjPDF.PDFNewPage
        End If
    Next I
    PictPrint_Cuadro_Linea ObjPDF, 1, PrimeraLinea, 20, PosLinea, QBColor(Negro), "B"
    PictPrint_Cuadro_Linea ObjPDF, 1.5, PrimeraLinea, 1.5, PosLinea, QBColor(Negro), "B"
    PictPrint_Cuadro_Linea ObjPDF, 1, PosLinea - 0.4, 20, PosLinea - 0.4, QBColor(Negro)
    PictPrint_Cuadro_Linea ObjPDF, 11, PrimeraLinea - 0.4, 11, PosLinea - 0.4, QBColor(Negro)
    PictPrint_Texto ObjPDF, 1.5, PosLinea, FechaStrgCiudad(FechaSistema)
    If CheqFirma.value = 1 Then
       PosLinea = PosLinea + 2.5
       PictPrint_Texto ObjPDF, 3, PosLinea, Rector
       PictPrint_Texto ObjPDF, 13, PosLinea, Secretario1
       PosLinea = PosLinea + 0.4
       PictPrint_Texto ObjPDF, 3, PosLinea, TextoRector
       PictPrint_Texto ObjPDF, 13, PosLinea, TextoSecretario1
    End If
'Fin del PDF
ObjPDF.PDFEndPage
ObjPDF.PDFEndDoc
RatonNormal
End Sub

