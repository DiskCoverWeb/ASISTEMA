VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Begin VB.Form FCatalogoRolPagos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "RUBROS DE ROL NOMINA"
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   7695
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command3 
      Caption         =   "&Eliminar"
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
      Left            =   5775
      Picture         =   "FCtaRolP.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   3675
      Width           =   855
   End
   Begin VB.Frame FrmGrupos 
      Caption         =   "SELECCIONE EL GRUPO DE ROL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2115
      Left            =   3885
      TabIndex        =   4
      Top             =   630
      Visible         =   0   'False
      Width           =   3270
      Begin MSDataListLib.DataList DLGrupos 
         Bindings        =   "FCtaRolP.frx":08CA
         DataSource      =   "AdoGrupos"
         Height          =   1740
         Left            =   105
         TabIndex        =   5
         Top             =   210
         Width           =   3060
         _ExtentX        =   5398
         _ExtentY        =   3069
         _Version        =   393216
         BackColor       =   12648447
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin MSDataListLib.DataCombo CGrupo 
      Bindings        =   "FCtaRolP.frx":08E2
      DataSource      =   "AdoRubros"
      Height          =   315
      Left            =   3465
      TabIndex        =   1
      Top             =   105
      Width           =   4110
      _ExtentX        =   7250
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
   Begin VB.CommandButton Command1 
      Caption         =   "&Grabar"
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
      Left            =   4830
      Picture         =   "FCtaRolP.frx":08FA
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   3675
      Width           =   855
   End
   Begin VB.CommandButton Command2 
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
      Left            =   6720
      Picture         =   "FCtaRolP.frx":0C04
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   3675
      Width           =   855
   End
   Begin MSAdodcLib.Adodc AdoRubros 
      Height          =   330
      Left            =   105
      Top             =   3675
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
      Caption         =   "Rubros"
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
   Begin MSMask.MaskEdBox MBCta_Quincena 
      Height          =   330
      Left            =   2205
      TabIndex        =   11
      Top             =   1575
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   582
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
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MBCta_Antig 
      Height          =   330
      Left            =   5985
      TabIndex        =   23
      Top             =   630
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   582
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
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MBCta_IESS_Personal 
      Height          =   330
      Left            =   5985
      TabIndex        =   29
      Top             =   1575
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   582
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
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MBCta_Vacacion 
      Height          =   330
      Left            =   2205
      TabIndex        =   7
      Top             =   945
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   582
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
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MBCta_Horas_Ext 
      Height          =   330
      Left            =   2205
      TabIndex        =   9
      Top             =   1260
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   582
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
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MBCta_Sueldo 
      Height          =   330
      Left            =   2205
      TabIndex        =   3
      ToolTipText     =   "<Ctrl+R> Seleccionar Cuentas de otros Grupo del Rol"
      Top             =   630
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   582
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
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MBCta_Fondo_Reserva_G 
      Height          =   330
      Left            =   2205
      TabIndex        =   19
      Top             =   2940
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   582
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
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MBCta_Fondo_Reserva_P 
      Height          =   330
      Left            =   5985
      TabIndex        =   37
      Top             =   2940
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   582
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
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MBCta_Vacaciones_P 
      Height          =   330
      Left            =   5985
      TabIndex        =   39
      Top             =   3255
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   582
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
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MBCta_Vacaciones_G 
      Height          =   330
      Left            =   2205
      TabIndex        =   21
      Top             =   3255
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   582
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
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MBCta_Aporte_Patronal_G 
      Height          =   330
      Left            =   2205
      TabIndex        =   13
      Top             =   1995
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   582
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
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MBCta_IESS_Patronal 
      Height          =   330
      Left            =   5985
      TabIndex        =   31
      Top             =   1995
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   582
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
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MBCta_Diferencia 
      Height          =   330
      Left            =   5985
      TabIndex        =   25
      Top             =   945
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   582
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
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MBCta_Decimo_Cuarto_P 
      Height          =   330
      Left            =   5985
      TabIndex        =   35
      Top             =   2625
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   582
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
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MBCta_Decimo_Cuarto_G 
      Height          =   330
      Left            =   2205
      TabIndex        =   17
      Top             =   2625
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   582
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
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MBCta_Decimo_Tercer_P 
      Height          =   330
      Left            =   5985
      TabIndex        =   33
      Top             =   2310
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   582
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
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MBCta_Decimo_Tercer_G 
      Height          =   330
      Left            =   2205
      TabIndex        =   15
      Top             =   2310
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   582
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
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MBCta_ExtConyugue 
      Height          =   330
      Left            =   5985
      TabIndex        =   27
      Top             =   1260
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   582
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
      PromptChar      =   "_"
   End
   Begin MSAdodcLib.Adodc AdoGrupos 
      Height          =   330
      Left            =   105
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
      Caption         =   "Grupos"
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
   Begin VB.Label Label36 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Decimo Tercer (G)"
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
      TabIndex        =   14
      Top             =   2310
      Width           =   2115
   End
   Begin VB.Label Label39 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Decimo Tercer (P)"
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
      Left            =   3885
      TabIndex        =   32
      Top             =   2310
      Width           =   2115
   End
   Begin VB.Label Label35 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Decimo Cuarto (G)"
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
      TabIndex        =   16
      Top             =   2625
      Width           =   2115
   End
   Begin VB.Label Label38 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Decimo Cuarto (P)"
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
      Left            =   3885
      TabIndex        =   34
      Top             =   2625
      Width           =   2115
   End
   Begin VB.Label Label30 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Hora no Trabajadas"
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
      Left            =   3885
      TabIndex        =   24
      Top             =   945
      Width           =   2115
   End
   Begin VB.Label Label12 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Ext. de Conyugue (P)"
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
      Left            =   3885
      TabIndex        =   26
      Top             =   1260
      Width           =   2115
   End
   Begin VB.Label Label40 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Fondo de Reserva (P)"
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
      Left            =   3885
      TabIndex        =   36
      Top             =   2940
      Width           =   2115
   End
   Begin VB.Label Label37 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Fondo de Reserva (G)"
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
      TabIndex        =   18
      Top             =   2940
      Width           =   2115
   End
   Begin VB.Label Label32 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Aporte Patronal (P)"
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
      Left            =   3885
      TabIndex        =   30
      Top             =   1995
      Width           =   2115
   End
   Begin VB.Label Label34 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Aporte Patronal (G)"
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
      TabIndex        =   12
      Top             =   1995
      Width           =   2115
   End
   Begin VB.Label Label53 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Prov. Vacaciones (G)"
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
      TabIndex        =   20
      Top             =   3255
      Width           =   2115
   End
   Begin VB.Label Label52 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Prov. Vacaciones (P)"
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
      Left            =   3885
      TabIndex        =   38
      Top             =   3255
      Width           =   2115
   End
   Begin VB.Label Label42 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Grupo de Rol Pago"
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
      Width           =   3375
   End
   Begin VB.Label Label20 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Sueldo Normal (G) "
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
      Top             =   630
      Width           =   2115
   End
   Begin VB.Label Label21 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Horas Extras (G)"
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
      TabIndex        =   8
      Top             =   1260
      Width           =   2115
   End
   Begin VB.Label Label23 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Sueldo Vacacion (G)"
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
      TabIndex        =   6
      Top             =   945
      Width           =   2115
   End
   Begin VB.Label Label33 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Aporte Personal (P)"
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
      Left            =   3885
      TabIndex        =   28
      Top             =   1575
      Width           =   2115
   End
   Begin VB.Label Label22 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Antigüedad"
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
      Left            =   3885
      TabIndex        =   22
      Top             =   630
      Width           =   2115
   End
   Begin VB.Label Label41 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Quincena (A-CxC)"
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
      TabIndex        =   10
      Top             =   1575
      Width           =   2115
   End
End
Attribute VB_Name = "FCatalogoRolPagos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CGrupo_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub CGrupo_LostFocus()
  Leer_Catalogo_Rol_Pagos CGrupo.Text
End Sub

Private Sub Command1_Click()
Dim EsNuevo As Boolean
Dim GrupoRol As String
  GrupoRol = UCaseStrg(CGrupo)
  If Len(GrupoRol) > 1 Then
     EsNuevo = False
     If AdoRubros.Recordset.RecordCount > 0 Then
        AdoRubros.Recordset.MoveFirst
        AdoRubros.Recordset.Find ("Grupo_Rol = '" & GrupoRol & "' ")
        If AdoRubros.Recordset.EOF Then EsNuevo = True
     Else
        EsNuevo = True
     End If
     If EsNuevo Then
        SetAddNew AdoRubros
        SetFields AdoRubros, "Grupo_Rol", GrupoRol
     End If
     SetFields AdoRubros, "Cta_Sueldo", CambioCodigoCta(MBCta_Sueldo)
     SetFields AdoRubros, "Cta_Horas_Ext", CambioCodigoCta(MBCta_Horas_Ext)
     SetFields AdoRubros, "Cta_Antiguedad", CambioCodigoCta(MBCta_Antig)
     SetFields AdoRubros, "Cta_Diferencia", CambioCodigoCta(MBCta_Diferencia)
     SetFields AdoRubros, "Cta_Vacacion", CambioCodigoCta(MBCta_Vacacion)
     SetFields AdoRubros, "Cta_Aporte_Patronal_G", CambioCodigoCta(MBCta_Aporte_Patronal_G)
     SetFields AdoRubros, "Cta_Decimo_Cuarto_G", CambioCodigoCta(MBCta_Decimo_Cuarto_G)
     SetFields AdoRubros, "Cta_Decimo_Cuarto_P", CambioCodigoCta(MBCta_Decimo_Cuarto_P)
     SetFields AdoRubros, "Cta_Decimo_Tercer_G", CambioCodigoCta(MBCta_Decimo_Tercer_G)
     SetFields AdoRubros, "Cta_Decimo_Tercer_P", CambioCodigoCta(MBCta_Decimo_Tercer_P)
     SetFields AdoRubros, "Cta_Fondo_Reserva_G", CambioCodigoCta(MBCta_Fondo_Reserva_G)
     SetFields AdoRubros, "Cta_Fondo_Reserva_P", CambioCodigoCta(MBCta_Fondo_Reserva_P)
     SetFields AdoRubros, "Cta_Ext_Conyugue_P", CambioCodigoCta(MBCta_ExtConyugue)
     SetFields AdoRubros, "Cta_Vacaciones_G", CambioCodigoCta(MBCta_Vacaciones_G)
     SetFields AdoRubros, "Cta_Vacaciones_P", CambioCodigoCta(MBCta_Vacaciones_P)
     SetFields AdoRubros, "Cta_IESS_Patronal", CambioCodigoCta(MBCta_IESS_Patronal)
     SetFields AdoRubros, "Cta_IESS_Personal", CambioCodigoCta(MBCta_IESS_Personal)
     SetFields AdoRubros, "Cta_Quincena", CambioCodigoCta(MBCta_Quincena)
     SetUpdate AdoRubros
     
     Llenar_Grupos
     Leer_Catalogo_Rol_Pagos GrupoRol
     MsgBox "Proceso grabado exitosamente"
  Else
     MsgBox "No se puede grabar este tipo de Grupo"
  End If
End Sub

Private Sub Command2_Click()
  Unload FCatalogoRolPagos
End Sub

Private Sub Command3_Click()
Dim Grupo_Rol As String
  Grupo_Rol = CGrupo.Text
  sSQL = "SELECT * " _
       & "FROM Catalogo_Rol_Pagos " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND Grupo_Rol = '" & Grupo_Rol & "' "
  Select_Adodc AdoGrupos, sSQL
  If AdoGrupos.Recordset.RecordCount <= 0 Then
     sSQL = "DELETE * " _
          & "FROM Catalogo_Rol_Cuentas " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "AND Grupo_Rol = '" & Grupo_Rol & "' "
     Ejecutar_SQL_SP sSQL
     MsgBox "Proceso exitoso"
  Else
     MsgBox "No se puede eliminar, hay procesos vinculados"
  End If
  Llenar_Grupos
End Sub

Private Sub DLGrupos_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DLGrupos_LostFocus()
   Leer_Catalogo_Rol_Pagos DLGrupos
   FrmGrupos.Visible = False
   'MBCta_Sueldo.SetFocus
End Sub

Private Sub Form_Activate()
  FormatoMaskCta MBCta_Antig
  FormatoMaskCta MBCta_Sueldo
  FormatoMaskCta MBCta_Vacacion
  FormatoMaskCta MBCta_Horas_Ext
  FormatoMaskCta MBCta_Diferencia
  FormatoMaskCta MBCta_Aporte_Patronal_G
  FormatoMaskCta MBCta_Decimo_Cuarto_G
  FormatoMaskCta MBCta_Decimo_Cuarto_P
  FormatoMaskCta MBCta_Decimo_Tercer_G
  FormatoMaskCta MBCta_Decimo_Tercer_P
  FormatoMaskCta MBCta_Fondo_Reserva_G
  FormatoMaskCta MBCta_Fondo_Reserva_P
  FormatoMaskCta MBCta_IESS_Patronal
  FormatoMaskCta MBCta_IESS_Personal
  FormatoMaskCta MBCta_Vacaciones_G
  FormatoMaskCta MBCta_Vacaciones_P
  FormatoMaskCta MBCta_Quincena
  FormatoMaskCta MBCta_ExtConyugue
  
  Llenar_Grupos
  Leer_Catalogo_Rol_Pagos CGrupo.Text
  
  RatonNormal
  CGrupo.SetFocus
End Sub

Private Sub Form_Load()
 RatonReloj
 CentrarForm FCatalogoRolPagos
 ConectarAdodc AdoGrupos
 ConectarAdodc AdoRubros
End Sub

Private Sub MBCta_Antig_GotFocus()
    MarcarTexto MBCta_Antig
End Sub

Private Sub MBCta_Antig_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBCta_Aporte_Patronal_G_GotFocus()
  MarcarTexto MBCta_Aporte_Patronal_G
End Sub

Private Sub MBCta_Aporte_Patronal_G_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBCta_Decimo_Cuarto_G_GotFocus()
  MarcarTexto MBCta_Decimo_Cuarto_G
End Sub

Private Sub MBCta_Decimo_Cuarto_G_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBCta_Decimo_Cuarto_P_GotFocus()
  MarcarTexto MBCta_Decimo_Cuarto_P
End Sub

Private Sub MBCta_Decimo_Cuarto_P_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBCta_Decimo_Tercer_G_GotFocus()
  MarcarTexto MBCta_Decimo_Tercer_G
End Sub

Private Sub MBCta_Decimo_Tercer_G_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBCta_Decimo_Tercer_P_GotFocus()
  MarcarTexto MBCta_Decimo_Tercer_P
End Sub

Private Sub MBCta_Decimo_Tercer_P_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBCta_Diferencia_GotFocus()
  MarcarTexto MBCta_Diferencia
End Sub

Private Sub MBCta_ExtConyugue_GotFocus()
  MarcarTexto MBCta_ExtConyugue
End Sub

Private Sub MBCta_ExtConyugue_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBCta_Fondo_Reserva_G_GotFocus()
  MarcarTexto MBCta_Fondo_Reserva_G
End Sub

Private Sub MBCta_Fondo_Reserva_G_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBCta_Fondo_Reserva_P_GotFocus()
   MarcarTexto MBCta_Fondo_Reserva_P
End Sub

Private Sub MBCta_Fondo_Reserva_P_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBCta_Horas_Ext_GotFocus()
  MarcarTexto MBCta_Horas_Ext
End Sub

Private Sub MBCta_IESS_Patronal_GotFocus()
  MarcarTexto MBCta_IESS_Patronal
End Sub

Private Sub MBCta_IESS_Patronal_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBCta_IESS_Personal_GotFocus()
  MarcarTexto MBCta_IESS_Personal
End Sub

Private Sub MBCta_IESS_Personal_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBCta_Quincena_GotFocus()
  MarcarTexto MBCta_Quincena
End Sub

Private Sub MBCta_Quincena_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBCta_Sueldo_GotFocus()
  MarcarTexto MBCta_Sueldo
End Sub

Private Sub MBCta_Sueldo_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  PresionoEnter KeyCode
  If CtrlDown And KeyCode = vbKeyC Then
     FrmGrupos.Visible = True
     DLGrupos.SetFocus
  End If
  
  If CtrlDown And KeyCode = vbKeyD Then
     sSQL = "DELETE * " _
          & "FROM Catalogo_Rol_Cuentas " _
          & "WHERE Item <> '000' "
     Ejecutar_SQL_SP sSQL
  
     sSQL = "INSERT INTO Catalogo_Rol_Cuentas (X, Fecha, Item, Periodo, Grupo_Rol, Cta_Diferencia, Cta_Vacacion, Cta_Sueldo, Cta_Horas_Ext, Cta_Aporte_Patronal_G, " _
          & "Cta_Decimo_Cuarto_G, Cta_Decimo_Cuarto_P, Cta_Decimo_Tercer_P, Cta_Fondo_Reserva_G, " _
          & "Cta_Fondo_Reserva_P, Cta_IESS_Personal, Cta_Quincena, Cta_Decimo_Tercer_G, Cta_IESS_Patronal, " _
          & "Cta_Antiguedad, Cta_Vacaciones_G, Cta_Vacaciones_P, Cta_Ext_Conyugue_P, " _
          & "Cta_Utilidades_G, Cta_Utilidades_P) " _
          & "SELECT '.', Fecha, Item, Periodo, Grupo_Rol, Cta_Diferencia, Cta_Vacacion, Cta_Sueldo, Cta_Horas_Ext, Cta_Aporte_Patronal_G, " _
          & "Cta_Decimo_Cuarto_G, Cta_Decimo_Cuarto_P, Cta_Decimo_Tercer_P, Cta_Fondo_Reserva_G, " _
          & "Cta_Fondo_Reserva_P, Cta_IESS_Personal, Cta_Quincena, Cta_Decimo_Tercer_G, Cta_IESS_Patronal, " _
          & "Cta_Antiguedad, Cta_Vacaciones_G, Cta_Vacaciones_P, Cta_Ext_Conyugue_P, " _
          & "Cta_Utilidades_G, Cta_Utilidades_P " _
          & "FROM Catalogo_Rol_Pagos " _
          & "WHERE Item <> '000' " _
          & "GROUP BY Item, Periodo, Grupo_Rol, Cta_Diferencia, Cta_Vacacion, Cta_Sueldo, Cta_Horas_Ext, Cta_Aporte_Patronal_G, " _
          & "Cta_Decimo_Cuarto_G, Cta_Decimo_Cuarto_P, Cta_Decimo_Tercer_P, Cta_Fondo_Reserva_G, " _
          & "Cta_Fondo_Reserva_P, Cta_IESS_Personal, Cta_Quincena, Cta_Decimo_Tercer_G, Cta_IESS_Patronal, " _
          & "Cta_Antiguedad, Cta_Vacaciones_G, Cta_Vacaciones_P, Cta_Ext_Conyugue_P, " _
          & "Cta_Utilidades_G, Cta_Utilidades_P " _
          & "ORDER BY Item, Periodo, Grupo_Rol, Cta_Diferencia, Cta_Vacacion, Cta_Sueldo, Cta_Horas_Ext, Cta_Aporte_Patronal_G, " _
          & "Cta_Decimo_Cuarto_G, Cta_Decimo_Cuarto_P, Cta_Decimo_Tercer_P, Cta_Fondo_Reserva_G, " _
          & "Cta_Fondo_Reserva_P, Cta_IESS_Personal, Cta_Quincena, Cta_Decimo_Tercer_G, Cta_IESS_Patronal, " _
          & "Cta_Antiguedad, Cta_Vacaciones_G, Cta_Vacaciones_P, Cta_Ext_Conyugue_P, " _
          & "Cta_Utilidades_G, Cta_Utilidades_P "
     Ejecutar_SQL_SP sSQL
     MsgBox "Proceso exitoso"
  End If
End Sub

Private Sub MBCta_Vacacion_GotFocus()
  MarcarTexto MBCta_Vacacion
End Sub

Private Sub MBCta_Vacacion_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBCta_Vacaciones_G_GotFocus()
   MarcarTexto MBCta_Vacaciones_G
End Sub

Private Sub MBCta_Vacaciones_G_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBCta_Vacaciones_P_GotFocus()
   MarcarTexto MBCta_Vacaciones_P
End Sub

Private Sub MBCta_Horas_Ext_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBCta_Diferencia_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBCta_Vacaciones_P_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Public Sub Llenar_Grupos()
  sSQL = "SELECT * " _
       & "FROM Catalogo_Rol_Cuentas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "ORDER BY Grupo_Rol "
  SelectDB_Combo CGrupo, AdoRubros, sSQL, "Grupo_Rol"
  SelectDB_List DLGrupos, AdoGrupos, sSQL, "Grupo_Rol"
End Sub

Public Sub Leer_Catalogo_Rol_Pagos(GrupoRol As String)
    MBCta_Diferencia = FormatoCodigoCta("0")
    MBCta_Vacacion = FormatoCodigoCta("0")
    MBCta_Sueldo = FormatoCodigoCta("0")
    MBCta_Horas_Ext = FormatoCodigoCta("0")
    MBCta_Antig = FormatoCodigoCta("0")
    MBCta_Aporte_Patronal_G = FormatoCodigoCta("0")
    MBCta_Decimo_Cuarto_G = FormatoCodigoCta("0")
    MBCta_Decimo_Cuarto_P = FormatoCodigoCta("0")
    MBCta_Decimo_Tercer_G = FormatoCodigoCta("0")
    MBCta_Decimo_Tercer_P = FormatoCodigoCta("0")
    MBCta_Fondo_Reserva_G = FormatoCodigoCta("0")
    MBCta_Fondo_Reserva_P = FormatoCodigoCta("0")
    MBCta_Vacaciones_G = FormatoCodigoCta("0")
    MBCta_Vacaciones_P = FormatoCodigoCta("0")
    MBCta_IESS_Patronal = FormatoCodigoCta("0")
    MBCta_IESS_Personal = FormatoCodigoCta("0")
    MBCta_Quincena = FormatoCodigoCta("0")
    MBCta_ExtConyugue = FormatoCodigoCta("0")
    
    If Len(GrupoRol) > 1 Then
       With AdoGrupos.Recordset
        If .RecordCount > 0 Then
           .MoveFirst
           .Find ("Grupo_Rol = '" & GrupoRol & "' ")
            If Not .EOF Then
               MBCta_Diferencia = FormatoCodigoCta(.Fields("Cta_Diferencia"))
               MBCta_Vacacion = FormatoCodigoCta(.Fields("Cta_Vacacion"))
               MBCta_Sueldo = FormatoCodigoCta(.Fields("Cta_Sueldo"))
               MBCta_Horas_Ext = FormatoCodigoCta(.Fields("Cta_Horas_Ext"))
               MBCta_Antig = FormatoCodigoCta(.Fields("Cta_Antiguedad"))
               MBCta_Aporte_Patronal_G = FormatoCodigoCta(.Fields("Cta_Aporte_Patronal_G"))
               MBCta_Decimo_Cuarto_G = FormatoCodigoCta(.Fields("Cta_Decimo_Cuarto_G"))
               MBCta_Decimo_Cuarto_P = FormatoCodigoCta(.Fields("Cta_Decimo_Cuarto_P"))
               MBCta_Decimo_Tercer_G = FormatoCodigoCta(.Fields("Cta_Decimo_Tercer_G"))
               MBCta_Decimo_Tercer_P = FormatoCodigoCta(.Fields("Cta_Decimo_Tercer_P"))
               MBCta_Fondo_Reserva_G = FormatoCodigoCta(.Fields("Cta_Fondo_Reserva_G"))
               MBCta_Fondo_Reserva_P = FormatoCodigoCta(.Fields("Cta_Fondo_Reserva_P"))
               MBCta_Vacaciones_G = FormatoCodigoCta(.Fields("Cta_Vacaciones_G"))
               MBCta_Vacaciones_P = FormatoCodigoCta(.Fields("Cta_Vacaciones_P"))
               MBCta_IESS_Patronal = FormatoCodigoCta(.Fields("Cta_IESS_Patronal"))
               MBCta_IESS_Personal = FormatoCodigoCta(.Fields("Cta_IESS_Personal"))
               MBCta_Quincena = FormatoCodigoCta(.Fields("Cta_Quincena"))
               MBCta_ExtConyugue = FormatoCodigoCta(.Fields("Cta_Ext_Conyugue_P"))
               MBCta_Sueldo.SetFocus
            End If
        End If
       End With
    End If
End Sub

