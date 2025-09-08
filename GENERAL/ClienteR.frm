VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.5#0"; "comctl32.Ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form FClientesRazonSocial 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ASIENTO DE MATRICULA"
   ClientHeight    =   7335
   ClientLeft      =   -15
   ClientTop       =   270
   ClientWidth     =   12945
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   12945
   ShowInTaskbar   =   0   'False
   Begin ComctlLib.Toolbar TBarCliente 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   35
      Top             =   0
      Width           =   12945
      _ExtentX        =   22834
      _ExtentY        =   1164
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   14
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Salir"
            Object.ToolTipText     =   "Salir del Módulo"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Modificar"
            Object.ToolTipText     =   "Modifica el Registro"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Grabar"
            Object.ToolTipText     =   "Grabar los Datos del Registro"
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "FactMult"
            Object.ToolTipText     =   "Asignar a Facturacion Multiple"
            Object.Tag             =   ""
            ImageIndex      =   15
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "UnFacMult"
            Object.ToolTipText     =   "Desabilitar Asignacion de Facturacion"
            Object.Tag             =   ""
            ImageIndex      =   19
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Actualizar"
            Object.ToolTipText     =   "Actualiza Alumnos Recien Ingresados"
            Object.Tag             =   ""
            ImageIndex      =   14
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Activar"
            Object.ToolTipText     =   "Activar el Estudiante para asiento de matricula"
            Object.Tag             =   ""
            ImageIndex      =   24
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Actualizar_DBF"
            Object.ToolTipText     =   "Actualizar Alumnos al Educativo (DBF)"
            Object.Tag             =   ""
            ImageIndex      =   25
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Ahorros"
            Object.ToolTipText     =   "Presenta solo Debitos a Cuentas de Ahorros"
            Object.Tag             =   ""
            ImageIndex      =   9
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Corrientes"
            Object.ToolTipText     =   "Presenta solo Debitos a Cuentas Corrientes"
            Object.Tag             =   ""
            ImageIndex      =   11
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Tarjetas"
            Object.ToolTipText     =   "Presenta solo los de Tarjetas"
            Object.Tag             =   ""
            ImageIndex      =   12
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "-"
            Object.Tag             =   ""
            Style           =   3
         EndProperty
         BeginProperty Button13 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Certificado"
            Object.ToolTipText     =   "Imprime Certificado de no Adeudar"
            Object.Tag             =   ""
            ImageIndex      =   20
         EndProperty
         BeginProperty Button14 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Certificados"
            Object.ToolTipText     =   "Imprime Certificados de no adeudar"
            Object.Tag             =   ""
            ImageIndex      =   25
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.Frame Frame3 
         Height          =   645
         Left            =   7770
         TabIndex        =   36
         Top             =   0
         Width           =   2955
         Begin MSMask.MaskEdBox MBFechaCorte 
            Height          =   330
            Left            =   1575
            TabIndex        =   37
            Top             =   210
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
         Begin VB.Label Label33 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Fecha de &Corte"
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
            TabIndex        =   38
            Top             =   210
            Width           =   1485
         End
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&S"
      Height          =   330
      Left            =   840
      TabIndex        =   34
      Top             =   7455
      Width           =   330
   End
   Begin VB.CheckBox CheqPorDeposito 
      Caption         =   "Depositar al Banco"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7875
      TabIndex        =   33
      Top             =   6930
      Width           =   2745
   End
   Begin VB.TextBox TxtDescuento 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
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
      Left            =   11130
      MaxLength       =   12
      TabIndex        =   15
      Text            =   "0.00"
      Top             =   4620
      Width           =   1695
   End
   Begin VB.TextBox TxtCtaNo 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   330
      Left            =   1890
      MaxLength       =   20
      TabIndex        =   27
      Text            =   "."
      Top             =   6825
      Width           =   4530
   End
   Begin VB.ComboBox CTipoCta 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   105
      TabIndex        =   25
      Text            =   "TARJETA"
      Top             =   6825
      Width           =   1695
   End
   Begin VB.TextBox TxtEmailRS 
      BackColor       =   &H00C0FFFF&
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
      Left            =   3360
      MaxLength       =   60
      TabIndex        =   21
      Top             =   5565
      Width           =   9465
   End
   Begin VB.TextBox TxtLugarTrabajoR 
      BackColor       =   &H00C0FFFF&
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
      Left            =   3360
      MaxLength       =   50
      TabIndex        =   19
      Top             =   5250
      Width           =   9465
   End
   Begin VB.TextBox TxtRazonSocial 
      BackColor       =   &H00C0FFFF&
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
      Left            =   3360
      MaxLength       =   50
      TabIndex        =   17
      Top             =   4935
      Width           =   9465
   End
   Begin VB.TextBox TxtTelefonoRS 
      BackColor       =   &H00C0FFFF&
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
      Left            =   8085
      MaxLength       =   10
      TabIndex        =   13
      Top             =   4620
      Width           =   1695
   End
   Begin VB.TextBox TxtCedulaR 
      BackColor       =   &H00C0FFFF&
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
      Left            =   3360
      MaxLength       =   13
      TabIndex        =   10
      Top             =   4620
      Width           =   3165
   End
   Begin VB.TextBox TxtCI_RUC 
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
      Left            =   10605
      MaxLength       =   13
      TabIndex        =   8
      ToolTipText     =   "<Ctrl+M> Codigo de Matrícula"
      Top             =   4095
      Width           =   2220
   End
   Begin VB.TextBox TxtApellidosS 
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
      TabIndex        =   6
      Top             =   4095
      Width           =   10410
   End
   Begin VB.Frame Frame1 
      Caption         =   " NOMBRE DEL ALUMNO(A)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2955
      Left            =   105
      TabIndex        =   0
      Top             =   735
      Width           =   12720
      Begin VB.TextBox TxtCodigo 
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
         Left            =   10500
         MaxLength       =   13
         TabIndex        =   4
         ToolTipText     =   "<Ctrl+M> Codigo de Matrícula"
         Top             =   2520
         Width           =   2115
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "FOTO ALUMNO(A)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   2220
         Left            =   10500
         TabIndex        =   3
         Top             =   210
         Width           =   2115
         Begin VB.Image ImgFoto 
            Height          =   1935
            Left            =   105
            Picture         =   "ClienteR.frx":0000
            Stretch         =   -1  'True
            Top             =   210
            Width           =   1875
         End
      End
      Begin MSDataListLib.DataCombo DCCliente 
         Bindings        =   "ClienteR.frx":171ED
         DataSource      =   "AdoListCtas"
         Height          =   2520
         Left            =   2205
         TabIndex        =   2
         ToolTipText     =   "Ctrl+B: Buscar datos en forma general"
         Top             =   315
         Width           =   8205
         _ExtentX        =   14473
         _ExtentY        =   4445
         _Version        =   393216
         Style           =   1
         ForeColor       =   8388608
         Text            =   "Cliente"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DCGrupo 
         Bindings        =   "ClienteR.frx":17207
         DataSource      =   "AdoGrupo"
         Height          =   2520
         Left            =   105
         TabIndex        =   1
         Top             =   315
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   4445
         _Version        =   393216
         Style           =   1
         ForeColor       =   16711680
         Text            =   "123456789012345"
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
   End
   Begin MSAdodcLib.Adodc AdoBanco 
      Height          =   330
      Left            =   840
      Top             =   2310
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
      Caption         =   "Tarjetas"
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
   Begin MSAdodcLib.Adodc AdoListCtas 
      Height          =   330
      Left            =   840
      Top             =   2625
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
      Caption         =   "ListCtas"
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
      Left            =   840
      Top             =   1365
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
   Begin MSAdodcLib.Adodc AdoCuentas 
      Height          =   330
      Left            =   840
      Top             =   1995
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
   Begin MSAdodcLib.Adodc AdoCreditos 
      Height          =   330
      Left            =   3045
      Top             =   1050
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
      Caption         =   "Creditos"
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
      Left            =   840
      Top             =   1050
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
   Begin MSAdodcLib.Adodc AdoEducativo 
      Height          =   330
      Left            =   840
      Top             =   1680
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
   Begin MSAdodcLib.Adodc AdoGrupo 
      Height          =   330
      Left            =   3045
      Top             =   1365
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
      Caption         =   "Grupo"
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
   Begin MSAdodcLib.Adodc AdoDireccion 
      Height          =   330
      Left            =   3045
      Top             =   1680
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
      Caption         =   "Direccion"
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
   Begin MSDataListLib.DataCombo DCBanco 
      Bindings        =   "ClienteR.frx":1721E
      DataSource      =   "AdoBanco"
      Height          =   315
      Left            =   3360
      TabIndex        =   23
      Top             =   6090
      Width           =   9465
      _ExtentX        =   16695
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   12648447
      ForeColor       =   12582912
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
   Begin MSMask.MaskEdBox MBFecha 
      Height          =   330
      Left            =   6510
      TabIndex        =   29
      ToolTipText     =   "Formato de Fecha: DD/MM/AA"
      Top             =   6825
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   582
      _Version        =   393216
      ForeColor       =   12582912
      AllowPrompt     =   -1  'True
      AutoTab         =   -1  'True
      MaxLength       =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "MM/yyyy"
      Mask            =   "##/####"
      PromptChar      =   "0"
   End
   Begin VB.Label Label7 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " CADUCIDAD"
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
      Left            =   6510
      TabIndex        =   28
      Top             =   6510
      Width           =   1275
   End
   Begin VB.Label Label5 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " DESCUENTO"
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
      TabIndex        =   14
      Top             =   4620
      Width           =   1380
   End
   Begin VB.Label Label3 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " NUMERO DE CUENTA"
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
      TabIndex        =   26
      Top             =   6510
      Width           =   4530
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " TIPO DE CUENTA"
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
      TabIndex        =   24
      Top             =   6510
      Width           =   1695
   End
   Begin VB.Label Label27 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " CORREO FACTURA"
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
      Left            =   105
      TabIndex        =   20
      Top             =   5565
      Width           =   3270
   End
   Begin VB.Label Label31 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " LUGAR TRABAJO"
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
      Top             =   5250
      Width           =   3270
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " DEBITO AUTOMATICO"
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
      TabIndex        =   22
      Top             =   6090
      Width           =   3270
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00004080&
      BorderWidth     =   3
      X1              =   0
      X2              =   12915
      Y1              =   5985
      Y2              =   5985
   End
   Begin VB.Label Label25 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " RAZON SOCIAL REPRESENTANTE"
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
      Top             =   4935
      Width           =   3270
   End
   Begin VB.Label Label42 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " TELEFONO"
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
      Left            =   6825
      TabIndex        =   12
      Top             =   4620
      Width           =   1275
   End
   Begin VB.Label LblTD 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "R"
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
      Left            =   6510
      TabIndex        =   11
      Top             =   4620
      Width           =   330
   End
   Begin VB.Label Label17 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " CEDULA DE IDENTIDAD"
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
      TabIndex        =   9
      Top             =   4620
      Width           =   3270
   End
   Begin VB.Label Label37 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   10605
      TabIndex        =   32
      Top             =   6510
      Width           =   2220
   End
   Begin VB.Label Label39 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " DEUDA PENDIENTE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   7875
      TabIndex        =   31
      Top             =   6510
      Width           =   2745
   End
   Begin VB.Label Label18 
      Alignment       =   1  'Right Justify
      Caption         =   "NO EXISTE DATOS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   10605
      TabIndex        =   30
      Top             =   6930
      Width           =   2220
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00004080&
      BorderWidth     =   3
      X1              =   0
      X2              =   12915
      Y1              =   4515
      Y2              =   4515
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " CODIGO BANCO"
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
      Left            =   10605
      TabIndex        =   7
      Top             =   3780
      Width           =   2220
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " APELLIDOS Y NOMBRES"
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
      Left            =   105
      TabIndex        =   5
      Top             =   3780
      Width           =   10410
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   105
      Top             =   7350
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   26
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ClienteR.frx":17235
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ClienteR.frx":1754F
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ClienteR.frx":17869
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ClienteR.frx":17B83
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ClienteR.frx":17E9D
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ClienteR.frx":181B7
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ClienteR.frx":184D1
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ClienteR.frx":187EB
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ClienteR.frx":18B05
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ClienteR.frx":18E1F
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ClienteR.frx":19139
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ClienteR.frx":19453
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ClienteR.frx":281E5
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ClienteR.frx":284FF
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ClienteR.frx":28819
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ClienteR.frx":28B33
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ClienteR.frx":28CE5
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ClienteR.frx":294FF
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ClienteR.frx":2973D
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ClienteR.frx":29A57
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ClienteR.frx":29D71
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ClienteR.frx":29F4B
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ClienteR.frx":2A265
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ClienteR.frx":2A57F
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ClienteR.frx":2A899
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ClienteR.frx":2ABB3
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FClientesRazonSocial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim AdoRegs As ADODB.Recordset
Dim Archivo_Foto As String
Dim Cliente_Ant As String
Dim NombFilePict As String
Dim Imprime As Boolean
Dim Estudiante As String
Dim SexoEst As String
Dim FechaC As String
Dim Strgs As String
Dim TipoReporte As String

Public Sub Imprimir_Certificados_Pagos()
With AdoListCtas.Recordset
 If .RecordCount > 0 Then
     RatonReloj
     tPrint.TipoImpresion = Es_PDF
     tPrint.NombreArchivo = NombFilePict
     tPrint.TituloArchivo = "Certificado " & Estudiante
     tPrint.TipoLetra = TipoTimes
     tPrint.OrientacionPagina = 1
     tPrint.PaginaA4 = True
     tPrint.EsCampoCorto = False
     tPrint.VerDocumento = True
     Set cPrint = New cImpresion
     cPrint.iniciaImpresion
    .MoveFirst
     Set AdoRegs = New ADODB.Recordset
     AdoRegs.CursorType = adOpenStatic
     AdoRegs.CursorLocation = adUseClient
     
     Do While Not .EOF
        Imprime = True
        CodigoCliente = .fields("Codigo")
        Estudiante = .fields("Cliente")
        SexoEst = .fields("Sexo")
        TipoReporte = "PREFA"
        Select Case TipoReporte
          Case "PREFA"
               Strgs = "SELECT * " _
                     & "FROM Clientes_Facturacion " _
                     & "WHERE Item = '" & NumEmpresa & "' " _
                     & "AND Fecha <= #" & BuscarFecha(FechaC) & "# " _
                     & "AND Codigo = '" & CodigoCliente & "' "
          Case "FA"
               'No piden todavia
        End Select
        Strgs = CompilarSQL(Strgs)
        AdoRegs.open Strgs, AdoStrCnn, , , adCmdText
        If AdoRegs.RecordCount > 0 Then Imprime = False
        AdoRegs.Close
        If Imprime Then
           SetNombrePRN = Impresota_PDF
           NombFilePict = "Certificado " & Estudiante
           Imprimir_Certificado_Pago CodigoCliente, 0.5, 0.5, MBFechaCorte
           Imprimir_Certificado_Pago CodigoCliente, 11, 0.5, MBFechaCorte
           cPrint.paginaNueva
        End If
       .MoveNext
     Loop
    .MoveFirst
    'fin del documento
     cPrint.finalizaImpresion
     RatonNormal
     MsgBox "Proceso Terminado. Proceda a imprimir el PDF"
 End If
End With
End Sub

Public Sub Imprimir_Certificado_Pago(CodigoCli As String, _
                                     Xo As Single, _
                                     Yo As Single, _
                                     Fecha_Corte As String)
Dim CadAux As String
    cPrint.tipoNegrilla = True
    cPrint.PorteDeLetra = 10
    cPrint.printImagen LogoTipo, Xo + 5, Yo, 3, 1.5
    If UCaseStrg(RazonSocial) = UCaseStrg(NombreComercial) Then
       cPrint.printTexto Xo + 1.5, Yo + 0.4, UCaseStrg(RazonSocial)
       CadAux = RazonSocial
    ElseIf Len(RazonSocial) > 1 And Len(NombreComercial) > 1 Then
       cPrint.printTexto Xo + 1.5, Yo + 0.4, UCaseStrg(RazonSocial)
       cPrint.printTexto Xo + 1.5, Yo + 0.8, UCaseStrg(NombreComercial)
       CadAux = RazonSocial & ", " & NombreComercial
    Else
       cPrint.printTexto Xo + 1.5, Yo + 0.4, UCaseStrg(Empresa)
       CadAux = Empresa
    End If
    cPrint.PorteDeLetra = 10
    cPrint.printTexto Xo + 1.5, Yo + 1.2, "CERTIFICADO DE NO ADEUDAR"
    cPrint.printTexto Xo + 1.5, Yo + 1.8, "Corte al: " & FechaSistema
    
    cPrint.printLinea Xo + 0.5, Yo + 3, Xo + 9.5, Yo + 3
    
    cPrint.PorteDeLetra = 11
    cPrint.tipoNegrilla = False
    Cadena = "La " & CadAux & ". Certifica que "
    If SexoEst = "M" Then
       Cadena = Cadena & "El "
    Else
       Cadena = Cadena & "La "
    End If
    Cadena = Cadena & "Estudiante " & UCaseStrg(Estudiante) & ", " _
           & "está al día en sus pagos, hasta el " & FechaStrg(Fecha_Corte) & ". El Representante puede hacer " _
           & "uso del presente para legalizar la matrícula del período 2017-2018, de su Representado. "

    PosLinea = cPrint.printTextoMultiple(Xo + 0.5, Yo + 4, Cadena, 8.5)

    cPrint.printTexto Xo + 1, Yo + 10.5, String(16, "_")
    cPrint.printTexto Xo + 1.2, Yo + 11, "Depto. Colecturia"
End Sub

Public Sub GrabarCliente()
  Si_No = False
  T = Normal
  TextoValido TxtRazonSocial, , True
  TextoValido TxtApellidosS, , True
  TextoValido TxtTelefonoRS, , True
  TextoValido TxtCedulaR, , True
  TextoValido TxtLugarTrabajoR, , True
  TextoValido TxtEmailRS
  TextoValido TxtCtaNo
  TextoValido TxtDescuento, True, , 2
  
  Mensajes = "Esta seguro de Grabar datos de:" & vbCrLf _
           & TxtApellidosS & vbCrLf _
           & "del Codigo: " & TxtCodigo
  Titulo = "Pregunta de grabación"
  If BoxMensaje = vbYes And AdoListCtas.Recordset.RecordCount > 0 Then
     Codigo = TxtCodigo
     If Codigo <> Ninguno Then
        RatonReloj
        sSQL = "SELECT Codigo,T,Cliente,TelefonoT,DireccionT,Email,CodigoU,DirNumero, ID " _
             & "FROM Clientes " _
             & "WHERE Codigo = '" & Codigo & "' "
        If Mas_Grupos Then sSQL = sSQL & "AND DirNumero = '" & NumEmpresa & "' "
        Select_Adodc AdoAux, sSQL
        With AdoAux.Recordset
         If .RecordCount > 0 Then
            .fields("T") = Normal
            .fields("Cliente") = TxtApellidosS
            .fields("TelefonoT") = TxtTelefonoRS
            .fields("DireccionT") = TxtLugarTrabajoR
            .fields("Email") = TxtEmailRS
            .fields("CodigoU") = CodigoUsuario
             If Mas_Grupos Then .fields("DirNumero") = NumEmpresa
            .Update
         End If
        End With
       'Grabamos Datos de los Alumnos
        sSQL = "SELECT * " _
             & "FROM Clientes_Matriculas " _
             & "WHERE Codigo = '" & Codigo & "' " _
             & "AND Periodo = '" & Periodo_Contable & "' " _
             & "AND Item = '" & NumEmpresa & "' "
        Select_Adodc AdoAux, sSQL
        With AdoAux.Recordset
             If .RecordCount <= 0 Then
                 SetAddNew AdoAux
                 SetFields AdoAux, "Codigo", Codigo
                 SetFields AdoAux, "Grupo_No", Grupo_No
                 SetFields AdoAux, "Periodo", Periodo_Contable
                 SetFields AdoAux, "Item", NumEmpresa
             End If
             SetFields AdoAux, "T", Normal
             SetFields AdoAux, "TD", LblTD.Caption
             SetFields AdoAux, "Cedula_R", TxtCedulaR
             SetFields AdoAux, "Lugar_Trabajo_R", TxtLugarTrabajoR
             SetFields AdoAux, "Representante", TxtRazonSocial
             SetFields AdoAux, "Telefono_RS", TxtTelefonoRS
             SetFields AdoAux, "Cta_Numero", TxtCtaNo
             SetFields AdoAux, "Descuento", TxtDescuento
             SetFields AdoAux, "Tipo_Cta", CTipoCta
             SetFields AdoAux, "Caducidad", UltimoDiaMes("01/" & MBFecha)
             SetFields AdoAux, "Por_Deposito", CBool(CheqPorDeposito.value)
             If AdoBanco.Recordset.RecordCount > 0 Then
                AdoBanco.Recordset.MoveFirst
                AdoBanco.Recordset.Find ("Descripcion = '" & DCBanco & "' ")
                If Not AdoBanco.Recordset.EOF Then
                   SetFields AdoAux, "Cod_Banco", AdoBanco.Recordset.fields("Codigo")
                Else
                   SetFields AdoAux, "Cod_Banco", 0
                End If
             End If
             SetUpdate AdoAux
        End With
        TBeneficiario.Codigo = Codigo
        TBeneficiario.Direccion = TxtLugarTrabajoR
        'TBeneficiario.Telefono1 = TxtTelefonoR
        TBeneficiario.Representante = TxtRazonSocial
        TBeneficiario.CI_RUC = TxtCedulaR
        RatonNormal
     Else
        MsgBox "No se puede grabar este Codigo"
     End If
  End If
  Estudiante_DBF.codest = Codigo
  Estudiante_DBF.cedula = Codigo
  Estudiante_DBF.cedular = TxtCedulaR
  'Estudiante_DBF.fonopaga = TxtTelefonoR
  Estudiante_DBF.pagador = TxtRazonSocial
  Estudiante_DBF.direcpaga = TxtLugarTrabajoR
  'Actualizar_Pagos
  FA.CodigoC = Codigo
  TBeneficiario = Leer_Datos_Cliente_SP(FA.CodigoC)
  ListarClientes CliFact
End Sub

Public Sub ListarClientes(Optional BuscarCliente As Boolean)
Dim TextosCliente As String
  RatonReloj
  
  sSQL = "SELECT TOP 50 Cliente, CI_RUC, TD, Codigo, Cta_CxP, Grupo, Cod_Ejec, Email, Archivo_Foto, Direccion " _
       & "FROM Clientes " _
       & "WHERE FA <> " & Val(adFalse) & " "
  If Mas_Grupos Then sSQL = sSQL & "AND DirNumero = '" & NumEmpresa & "' "
  If BuscarCliente Then sSQL = sSQL & "AND Grupo = '" & DCGrupo & "' "
  sSQL = sSQL _
       & "ORDER BY Cliente "
  SelectDB_Combo DCCliente, AdoListCtas, sSQL, "Cliente"
  Frame1.Caption = " NOMBRE DEL CLIENTE" & Space(30) & "Total Clientes: " & Format$(AdoListCtas.Recordset.RecordCount, "000000")
  RatonNormal
 'DCCliente.SetFocus
End Sub

Public Sub ListarCuenta(TextoBusqueda As String)
  TxtCodigo = Ninguno
  TxtLugarTrabajoR = Ninguno
  TxtTelefonoRS = Ninguno
  TxtCedulaR = Ninguno
  TxtRazonSocial = Ninguno
  DCBanco = Ninguno
  CTipoCta = Ninguno
  TxtCtaNo = Ninguno
  TxtDescuento = "0.00"
  CheqPorDeposito.value = 0
  With AdoListCtas.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Cliente = '" & TextoBusqueda & "' ")
       If Not .EOF Then
          TxtCodigo = .fields("Codigo")
          CodigoCliente = .fields("Codigo")
          TxtApellidosS = .fields("Cliente")
          DCCliente = TxtApellidosS
          TxtCI_RUC = .fields("CI_RUC")
          Grupo_No = .fields("Grupo")
          TxtEmailRS = .fields("Email")
          Archivo_Foto = .fields("Archivo_Foto")
          Label4.Caption = " APELLIDOS Y NOMBRE: (" & .fields("Direccion") & ")"
          DCGrupo.Text = Grupo_No
          TxtApellidosS.Enabled = False
          ImgFoto.Picture = LoadPicture()
          RutaDestino = RutaSistema & "\FOTOS\" & Archivo_Foto & ".JPG"
          If Dir(RutaDestino) <> "" Then
             ImgFoto.Picture = LoadPicture(RutaDestino)
          Else
             RutaDestino = RutaSistema & "\FOTOS\" & Archivo_Foto & ".GIF"
             If Dir(RutaDestino) <> "" Then ImgFoto.Picture = LoadPicture(RutaDestino)
          End If
         'Lista de Alumnos Matriculados
          sSQL = "SELECT * " _
               & "FROM Clientes_Matriculas " _
               & "WHERE Periodo = '" & Periodo_Contable & "' " _
               & "AND Item = '" & NumEmpresa & "' " _
               & "AND Codigo = '" & CodigoCliente & "' "
          Select_Adodc AdoCuentas, sSQL
          If AdoCuentas.Recordset.RecordCount > 0 Then
             TxtLugarTrabajoR = AdoCuentas.Recordset.fields("Lugar_Trabajo_R")
             TxtTelefonoRS = AdoCuentas.Recordset.fields("Telefono_RS")
             TxtCedulaR = AdoCuentas.Recordset.fields("Cedula_R")
             TxtRazonSocial = AdoCuentas.Recordset.fields("Representante")
             LblTD.Caption = AdoCuentas.Recordset.fields("TD")
             TxtCtaNo = AdoCuentas.Recordset.fields("Cta_Numero")
             TxtDescuento = AdoCuentas.Recordset.fields("Descuento")
             CTipoCta = AdoCuentas.Recordset.fields("Tipo_Cta")
             DCBanco = AdoCuentas.Recordset.fields("Cod_Banco")
             MBFecha = Format(AdoCuentas.Recordset.fields("Caducidad"), "MM/yyyy")
             If AdoCuentas.Recordset.fields("Por_Deposito") Then CheqPorDeposito.value = 1
             If AdoBanco.Recordset.RecordCount > 0 Then
                AdoBanco.Recordset.MoveFirst
                AdoBanco.Recordset.Find ("Codigo = " & Val(DCBanco) & " ")
                If Not AdoBanco.Recordset.EOF Then
                   DCBanco = AdoBanco.Recordset.fields("Descripcion")
                Else
                   DCBanco = Ninguno
                End If
             End If
          End If
       Else
          MsgBox "No Existe"
       End If
   Else
      MsgBox "No Existe"
   End If
  End With
  Label37.Caption = "0.00"
  sSQL = "SELECT CodigoC,SUM(Saldo_MN) As Deuda_Pendiente,COUNT(CodigoC) As CFacturas " _
       & "FROM Facturas " _
       & "WHERE Fecha <= #" & BuscarFecha(FechaSistema) & "# " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND CodigoC = '" & CodigoCliente & "' " _
       & "AND Saldo_MN > 0 " _
       & "AND T <> 'A' " _
       & "GROUP BY CodigoC "
  Select_Adodc AdoAux, sSQL
  With AdoAux.Recordset
   If .RecordCount > 0 Then
       Label37.Caption = Format$(.fields("Deuda_Pendiente"), "#,##0.00")
       Label39.Caption = "DEUDA PENDIENTE (" & .fields("CFacturas") & ")"
       If .fields("CFacturas") > 1 Then
         'TBarCliente.Buttons.Item(3).Enabled = False
          MsgBox "NO DEBE MATRICULAR EL ESTUDIANTE" & vbCrLf & vbCrLf _
               & "PORQUE TIENE UNA DEUDA PENDIENTE" & vbCrLf & vbCrLf _
               & "DE " & .fields("CFacturas") & " FACTURAS" & vbCrLf & vbCrLf _
               & "QUE SUMAN USD " & Label37.Caption
       End If
   End If
  End With
End Sub

Private Sub Command1_Click()
  Unload FClientesRazonSocial
End Sub

Private Sub DCCliente_DblClick(Area As Integer)
  SiguienteControl
End Sub

Private Sub DCCliente_KeyDown(KeyCode As Integer, Shift As Integer)
Dim NombreTabla(20) As String
  Keys_Especiales Shift
'''  If CtrlDown And KeyCode = vbKeyB Then
'''     ListarClientes False
'''     MsgBox "Busque el dato"
'''  End If
  PresionoEnter KeyCode
End Sub

Private Sub DCCliente_KeyPress(KeyAscii As Integer)
Dim Busqueda As String
    Busqueda = DCCliente.Text
    If Len(Busqueda) > 1 Then
       sSQL = "SELECT TOP 50 Cliente, CI_RUC, TD, Codigo, Cta_CxP, Grupo, Cod_Ejec, Email, Archivo_Foto, Direccion, ID " _
            & "FROM Clientes "
       If IsNumeric(Busqueda) Then sSQL = sSQL & "WHERE CI_RUC LIKE '" & Busqueda & "%' " Else sSQL = sSQL & "WHERE Cliente LIKE '%" & Busqueda & "%' "
       If Mas_Grupos Then sSQL = sSQL & "AND DirNumero = '" & NumEmpresa & "' "
       sSQL = sSQL _
            & "AND FA <> " & Val(adFalse) & " " _
            & "ORDER BY Cliente "
       Select_Adodc AdoListCtas, sSQL
    End If
End Sub

Private Sub DCCliente_LostFocus()
  ListarCuenta DCCliente
  TipoDoc = "M"
End Sub

Private Sub DCGrupo_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCGrupo_LostFocus()
  ListarClientes True
End Sub

Private Sub Form_Activate()
  MBFecha = Format(FechaSistema, "MM/yyyy")
  FechaComp = FechaSistema
  MBFechaCorte = "30/06/2017"
  Actualiza_Cursos
  
  CTipoCta.Clear
  CTipoCta.AddItem "CORRIENTE"
  CTipoCta.AddItem "AHORROS"
  CTipoCta.AddItem "TARJETA"
  CTipoCta.AddItem Ninguno
  CTipoCta.Text = Ninguno
  
  sSQL = "SELECT * " _
       & "FROM Tabla_Referenciales_SRI " _
       & "WHERE Tipo_Referencia = 'BANCOS Y COOP' " _
       & "AND Codigo >= '0' " _
       & "ORDER BY Descripcion "
  SelectDB_Combo DCBanco, AdoBanco, sSQL, "Descripcion"
 
  sSQL = "SELECT Grupo " _
       & "FROM Clientes " _
       & "WHERE Codigo <> '.' " _
       & "AND FA <> " & Val(adFalse) & " "
  If Mas_Grupos Then sSQL = sSQL & "AND DirNumero = '" & NumEmpresa & "' "
  sSQL = sSQL _
       & "GROUP BY Grupo " _
       & "ORDER BY Grupo "
  SelectDB_Combo DCGrupo, AdoGrupo, sSQL, "Grupo"
  ListarClientes CliFact
  RatonNormal
  DCGrupo.SetFocus
End Sub

Private Sub Form_Load()
   CentrarForm FClientesRazonSocial
   ConectarAdodc AdoAux
   ConectarAdodc AdoCta
   ConectarAdodc AdoGrupo
   ConectarAdodc AdoCuentas
   ConectarAdodc AdoListCtas
   ConectarAdodc AdoBanco
   ConectarAdodc AdoCreditos
   ConectarAdodc AdoEducativo
   ConectarAdodc AdoDireccion
   FClientesRazonSocial.Caption = "CREACION DEL CLIENTE"
End Sub

Private Sub TBarCliente_ButtonClick(ByVal Button As ComctlLib.Button)
Dim Secundaria1 As Boolean
Dim Secundaria2 As Boolean
 FechaComp = "01/" & MBFecha
 FechaC = UltimoDiaMes(MBFechaCorte)
 
'MsgBox Button.key
 Select Case Button.key
   Case "Salir"
        RatonNormal
        Unload FClientesRazonSocial
   Case "Modificar"
        TxtApellidosS.Enabled = True
        TxtRazonSocial.Enabled = True
   Case "Grabar"
        GrabarCliente
   Case "FactMult"
        Mensajes = "Asignar Facturacion " & TxtApellidosS & "."
        Titulo = "Pregunta de Facturacion Multiple"
        If BoxMensaje = vbYes Then
           CodigoCliente = TxtCodigo
           NombreCliente = TxtApellidosS
           sSQL = "UPDATE Clientes " _
                & "SET FA = " & Val(adTrue) & " " _
                & "WHERE Codigo = '" & CodigoCliente & "' "
           If Mas_Grupos Then sSQL = sSQL & "AND DirNumero = '" & NumEmpresa & "' "
           Ejecutar_SQL_SP sSQL
        End If
   Case "UnFacMult"
        Mensajes = "Des-activar a Facturacion a: " & TxtApellidosS & "."
        Titulo = "Pregunta de Facturacion Multiple"
        If BoxMensaje = vbYes Then
           CodigoCliente = TxtCodigo
           NombreCliente = TxtApellidosS
           sSQL = "UPDATE Clientes " _
                & "SET FA = " & Val(adFalse) & " " _
                & "WHERE Codigo = '" & CodigoCliente & "' "
           If Mas_Grupos Then sSQL = sSQL & "AND DirNumero = '" & NumEmpresa & "' "
           Ejecutar_SQL_SP sSQL
        End If
   Case "Actualizar"
        ListarClientes True 'CliFact
   Case "Activar"
        If ClaveAdministrador Then
           RatonReloj
           Cadena = InputBox("INGRESE EL MOTIVO DE LA ACTIVACION", "ACTIVACION DE MATRICULA", "ORDEN SUPERIOR")
           Control_Procesos "E", "MATRICULAR POR: " & Cadena
           sSQL = "UPDATE Clientes_Matriculas " _
                & "SET Matricular = " & Val(adTrue) & " " _
                & "WHERE Codigo = '" & TxtCodigo & "' " _
                & "AND Item = '" & NumEmpresa & "' " _
                & "AND Periodo = '" & Periodo_Contable & "' "
           Ejecutar_SQL_SP sSQL
           RatonNormal
           MsgBox "PROCESO REALIZADO CON EXITO"
        End If
   Case "Actualizar_DBF"
        Actualizar_Alumnos_DBF
   Case "Ahorros"
        sSQL = "SELECT C.*,CM.Tipo_Cta " _
             & "FROM Clientes As C,Clientes_Matriculas As CM " _
             & "WHERE C.FA <> " & Val(adFalse) & " " _
             & "AND CM.Item = '" & NumEmpresa & "' " _
             & "AND CM.Periodo = '" & Periodo_Contable & "' " _
             & "AND CM.Tipo_Cta = 'AHORROS' " _
             & "AND C.Codigo = CM.Codigo " _
             & "ORDER BY Cliente "
        SelectDB_Combo DCCliente, AdoListCtas, sSQL, "Cliente"
        Frame1.Caption = " NOMBRE DEL CLIENTE" & Space(30) & "Total Clientes: " & Format$(AdoListCtas.Recordset.RecordCount, "000000")
        DCCliente.SetFocus
   Case "Corrientes"
        sSQL = "SELECT C.*,CM.Tipo_Cta " _
             & "FROM Clientes As C,Clientes_Matriculas As CM " _
             & "WHERE C.FA <> " & Val(adFalse) & " " _
             & "AND CM.Item = '" & NumEmpresa & "' " _
             & "AND CM.Periodo = '" & Periodo_Contable & "' " _
             & "AND CM.Tipo_Cta = 'CORRIENTE' " _
             & "AND C.Codigo = CM.Codigo " _
             & "ORDER BY Cliente "
        SelectDB_Combo DCCliente, AdoListCtas, sSQL, "Cliente"
        Frame1.Caption = " NOMBRE DEL CLIENTE" & Space(30) & "Total Clientes: " & Format$(AdoListCtas.Recordset.RecordCount, "000000")
        DCCliente.SetFocus
   Case "Tarjetas"
        sSQL = "SELECT C.*,CM.Tipo_Cta " _
             & "FROM Clientes As C,Clientes_Matriculas As CM " _
             & "WHERE C.FA <> " & Val(adFalse) & " " _
             & "AND CM.Item = '" & NumEmpresa & "' " _
             & "AND CM.Periodo = '" & Periodo_Contable & "' " _
             & "AND CM.Tipo_Cta = 'TARJETA' " _
             & "AND C.Codigo = CM.Codigo " _
             & "ORDER BY Cliente "
        SelectDB_Combo DCCliente, AdoListCtas, sSQL, "Cliente"
        Frame1.Caption = " NOMBRE DEL CLIENTE" & Space(30) & "Total Clientes: " & Format$(AdoListCtas.Recordset.RecordCount, "000000")
        DCCliente.SetFocus
   Case "Certificado"
        Imprime = True
        TipoReporte = "PREFA"
        Select Case TipoReporte
          Case "PREFA"
               Strgs = "SELECT * " _
                     & "FROM Clientes_Facturacion " _
                     & "WHERE Item = '" & NumEmpresa & "' " _
                     & "AND Fecha <= #" & BuscarFecha(FechaC) & "# " _
                     & "AND Codigo = '" & CodigoCliente & "' "
          Case "FA"
               'No piden todavia
        End Select
        Strgs = CompilarSQL(Strgs)
        
        Set AdoRegs = New ADODB.Recordset
        AdoRegs.CursorType = adOpenStatic
        AdoRegs.CursorLocation = adUseClient
        AdoRegs.open Strgs, AdoStrCnn, , , adCmdText
        If AdoRegs.RecordCount > 0 Then Imprime = False
        AdoRegs.Close
        If Imprime Then
           Strgs = "SELECT Sexo,Cliente " _
                 & "FROM Clientes " _
                 & "WHERE Codigo = '" & CodigoCliente & "' "
           Strgs = CompilarSQL(Strgs)
           AdoRegs.open Strgs, AdoStrCnn, , , adCmdText
           If AdoRegs.RecordCount > 0 Then
              Estudiante = AdoRegs.fields("Cliente")
              SexoEst = AdoRegs.fields("Sexo")
           End If
           AdoRegs.Close
        
           SetNombrePRN = Impresota_PDF
           NombFilePict = "Certificado " & Estudiante
           tPrint.TipoImpresion = Es_PDF
           tPrint.NombreArchivo = NombFilePict
           tPrint.TituloArchivo = "Certificado " & Estudiante
           tPrint.TipoLetra = TipoTimes
           tPrint.OrientacionPagina = 1
           tPrint.PaginaA4 = True
           tPrint.EsCampoCorto = False
           tPrint.VerDocumento = True
           Set cPrint = New cImpresion
           cPrint.iniciaImpresion
           Imprimir_Certificado_Pago CodigoCliente, 0.5, 0.5, MBFechaCorte
           Imprimir_Certificado_Pago CodigoCliente, 11, 0.5, MBFechaCorte
          'fin del documento
           cPrint.finalizaImpresion
        Else
            MsgBox "El Estudiante tiene Deuda"
            End If
   Case "Certificados"
        Imprimir_Certificados_Pagos
 End Select
'If Button.key <> "Salir" Then ListarClientes True
 RatonNormal
End Sub

Private Sub TxtCedulaR_GotFocus()
  MarcarTexto TxtCedulaR
End Sub

Private Sub TxtCedulaR_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

'Private Sub TxtCedulaR_KeyPress(KeyAscii As Integer)
'   KeyAscii = Solo_Letras_Numeros(KeyAscii)
'End Sub

Private Sub TxtCedulaR_LostFocus()
    If Len(TxtCedulaR) > 1 Then
       DigVerif = Digito_Verificador(TxtCedulaR)
       If Tipo_RUC_CI.Tipo_Beneficiario = "P" Then
          Mensajes = "Este código es un Pasaporte"
          Titulo = "CONFIRMACION DE PASAPORTE"
          If BoxMensaje <> vbYes Then Tipo_RUC_CI.Tipo_Beneficiario = "O"
       End If
       If DigVerif = "-" And Tipo_RUC_CI.Tipo_Beneficiario <> "P" Then
          MsgBox "RUC/CEDULA INCORRECTA"
          TxtCedulaR.SetFocus
       Else
          LblTD.Caption = Tipo_RUC_CI.Tipo_Beneficiario
          If Tipo_RUC_CI.Tipo_Beneficiario = "R" Then
             Label25.Caption = " R A Z O N    S O C I A L"
             Label17.Caption = " R.U.C."
          Else
             Label25.Caption = " APELLIDOS Y NOMBRES"
             Label17.Caption = " C.I./Pasaporte/Otros"
          End If
          LblTD.Caption = Tipo_RUC_CI.Tipo_Beneficiario
       End If
       sSQL = "SELECT CM.Codigo,CM.Representante,CM.Cedula_R,CM.Telefono_RS,C.Email " _
            & "FROM Clientes_Matriculas As CM, Clientes As C " _
            & "WHERE CM.Item = '" & NumEmpresa & "' " _
            & "AND CM.Periodo = '" & Periodo_Contable & "' " _
            & "AND CM.Cedula_R = '" & TxtCedulaR & "' " _
            & "AND CM.Codigo = C.Codigo "
       Select_Adodc AdoAux, sSQL
       With AdoAux.Recordset
        If .RecordCount > 0 Then
            TxtRazonSocial = .fields("Representante")
            TxtTelefonoRS = .fields("Telefono_RS")
            TxtEmailRS = .fields("Email")
        End If
       End With
    End If
End Sub

Private Sub TxtCI_RUC_GotFocus()
  MarcarTexto TxtCI_RUC
End Sub

Private Sub TxtCI_RUC_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  If CtrlDown And KeyCode = vbKeyA Then
     sSQL = "UPDATE Clientes " _
          & "SET T = 'N' " _
          & "WHERE T = '.' "
     Ejecutar_SQL_SP sSQL
     
     sSQL = "UPDATE Clientes_Matriculas " _
          & "SET T = 'N' " _
          & "WHERE T = '.' "
     Ejecutar_SQL_SP sSQL

     MsgBox "Actualización realizado con éxito"
  End If
  PresionoEnter KeyCode
End Sub

Private Sub TxtApellidosS_GotFocus()
  MarcarTexto TxtApellidosS
End Sub

Private Sub TxtApellidosS_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtApellidosS_LostFocus()
  TextoValido TxtApellidosS, , True
  With AdoListCtas.Recordset
   If .RecordCount > 0 And TxtApellidosS <> Ninguno Then
       RatonReloj
      .MoveFirst
      .Find ("Cliente Like '" & TxtApellidosS & "' ")
       RatonNormal
       If Not .EOF Then
          MsgBox "El Cliente " & TxtApellidosS _
               & ", ya existe, está asignado a " & vbCrLf & vbCrLf _
               & .fields("Cliente") & vbCrLf & vbCrLf _
               & "Codigo: " & .fields("CI_RUC")
          DCCliente.SetFocus
       End If
   End If
  End With
End Sub

Private Sub TxtCI_RUC_LostFocus()
  TextoValido TxtCI_RUC, , True
  With AdoListCtas.Recordset
   If .RecordCount > 0 And TxtCI_RUC.Text <> Ninguno Then
       RatonReloj
      .MoveFirst
      .Find ("CI_RUC Like '" & TxtCI_RUC.Text & "' ")
       RatonNormal
       If Not .EOF Then
          If .fields("Cliente") <> TxtApellidosS Then
              MsgBox "Este Código, está asignado a " & vbCrLf & vbCrLf & .fields("Cliente")
              TxtCI_RUC.SetFocus
          Else
              TipoBenef = .fields("TD")
          End If
       Else
          DigVerif = Digito_Verificador(TxtCI_RUC.Text)
          Caracter = MidStrg(TxtCI_RUC.Text, 10, 1)
          Select Case Tipo_RUC_CI.Tipo_Beneficiario
            Case "C", "D", "R"
                 If DigVerif <> Caracter Then
                    MsgBox "Codigo Incorrecto"
                    TxtCI_RUC.SetFocus
                 End If
          End Select
       End If
   End If
  End With
  'Label4.Caption = "* APELLIDOS Y NOMBRES"
End Sub

Private Sub TxtCodigo_GotFocus()
  MarcarTexto TxtCodigo
End Sub

Private Sub TxtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtCtaNo_GotFocus()
  MarcarTexto TxtCtaNo
End Sub

Private Sub TxtCtaNo_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtDescuento_GotFocus()
   MarcarTexto TxtDescuento
End Sub

Private Sub TxtDescuento_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtEmailRS_GotFocus()
  MarcarTexto TxtEmailRS
End Sub

Private Sub TxtEmailRS_LostFocus()
  TextoValido TxtEmailRS
End Sub

Private Sub TxtEmailRS_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtLugarTrabajoR_GotFocus()
  MarcarTexto TxtLugarTrabajoR
End Sub

Private Sub TxtLugarTrabajoR_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtRazonSocial_GotFocus()
  MarcarTexto TxtRazonSocial
End Sub

Private Sub TxtRazonSocial_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtRazonSocial_LostFocus()
  TextoValido TxtRazonSocial, , True
End Sub

Private Sub TxtTelefonoRS_GotFocus()
   MarcarTexto TxtTelefonoRS
End Sub

Private Sub TxtTelefonoRS_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Public Sub Actualizar_Alumnos_DBF()
  Progreso_Barra.Mensaje_Box = "Actualizando Alumnos a Educativo"
  Progreso_Iniciar
  ListarClientes False
  With AdoListCtas.Recordset
   If .RecordCount > 0 Then
       Progreso_Barra.Valor_Maximo = Progreso_Barra.Valor_Maximo + .RecordCount
       Do While Not .EOF
          RatonReloj
          TBeneficiario.Patron_Busqueda = .fields("CI_RUC")
          TBeneficiario = Leer_Datos_Cliente_SP(TBeneficiario.Patron_Busqueda)
          Progreso_Barra.Mensaje_Box = "Actualizando: " & TBeneficiario.Cliente
          Progreso_Esperar
          Estudiante_DBF.codest = TBeneficiario.CI_RUC
          Estudiante_DBF.cedula = TBeneficiario.CI_RUC
          Estudiante_DBF.cedular = TBeneficiario.RUC_CI_Rep
          Estudiante_DBF.fonopaga = TBeneficiario.TelefonoT
          Estudiante_DBF.pagador = TBeneficiario.Representante
          Estudiante_DBF.direcpaga = TBeneficiario.Direccion_Rep
          'MsgBox "..."
          'Actualizar_Pagos
          RatonNormal
         .MoveNext
       Loop
   End If
  End With
  Progreso_Final
  RatonNormal
  MsgBox "Proceso Terminado"
End Sub

Private Sub MBFechaCorte_GotFocus()
   MarcarTexto MBFechaCorte
End Sub

Private Sub MBFechaCorte_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub MBFechaCorte_LostFocus()
  FechaValida MBFechaCorte
End Sub

