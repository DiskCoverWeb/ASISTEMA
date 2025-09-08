VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.5#0"; "COMCTL32.OCX"
Begin VB.Form ISubCtas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ingreso/Modificacion de SubCuentas"
   ClientHeight    =   6450
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9900
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6450
   ScaleWidth      =   9900
   ShowInTaskbar   =   0   'False
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   29
      Top             =   0
      Width           =   9900
      _ExtentX        =   17463
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   7
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Eliminar"
            Object.ToolTipText     =   "Elimina una Cuenta Contable"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Nuevo"
            Object.ToolTipText     =   "Nueva Cuenta Contable"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Grabar"
            Object.ToolTipText     =   "Grabar Cuenta"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Primero"
            Object.ToolTipText     =   "Primera Cuenta"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Anterior"
            Object.ToolTipText     =   "Anterior Cuenta"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Siguiente"
            Object.ToolTipText     =   "Siguiente Cuenta"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Ultimo"
            Object.ToolTipText     =   "Ultima Cuenta"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.CheckBox CheqBloquear 
      Caption         =   "Bloquear Codigo"
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
      Top             =   4200
      Width           =   1905
   End
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
      Left            =   1470
      MaxLength       =   10
      TabIndex        =   8
      Text            =   "000"
      Top             =   3360
      Width           =   1905
   End
   Begin VB.TextBox TxtNivel 
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
      Left            =   4620
      MaxLength       =   2
      TabIndex        =   10
      Text            =   "00"
      Top             =   3360
      Width           =   960
   End
   Begin VB.CheckBox CheqNivel 
      Caption         =   "Agrupación nivel"
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
      Left            =   5670
      TabIndex        =   11
      Top             =   3360
      Width           =   1905
   End
   Begin VB.TextBox TxtReembolso 
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
      Left            =   1575
      MaxLength       =   80
      TabIndex        =   20
      Text            =   "0"
      Top             =   4725
      Width           =   6210
   End
   Begin VB.CheckBox CheqCaja 
      Caption         =   "Gasto de Caja"
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
      Left            =   3570
      TabIndex        =   23
      Top             =   5145
      Visible         =   0   'False
      Width           =   2535
   End
   Begin MSDataListLib.DataList DLCtas 
      Bindings        =   "Isubctas.frx":0000
      DataSource      =   "AdoSubCta"
      Height          =   1980
      Left            =   105
      TabIndex        =   6
      Top             =   1260
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   3493
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSAdodcLib.Adodc AdoSubCta 
      Height          =   330
      Left            =   210
      Top             =   1470
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
   Begin VB.TextBox TextPresupuesto 
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
      Left            =   1575
      MaxLength       =   12
      MultiLine       =   -1  'True
      TabIndex        =   22
      Text            =   "Isubctas.frx":0018
      Top             =   5145
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipo de Cuenta"
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
      Left            =   105
      TabIndex        =   0
      Top             =   420
      Width           =   8415
      Begin VB.OptionButton OpcCC 
         Caption         =   "Centro de Costos"
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
         Left            =   6510
         TabIndex        =   4
         Top             =   210
         Width           =   1800
      End
      Begin VB.OptionButton OpcPM 
         Caption         =   "Modulo de Primas"
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
         Left            =   4410
         TabIndex        =   3
         Top             =   210
         Width           =   1905
      End
      Begin VB.OptionButton OpcI 
         Caption         =   "Modulo de Ingresos"
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
         TabIndex        =   1
         Top             =   210
         Value           =   -1  'True
         Width           =   2010
      End
      Begin VB.OptionButton OpcG 
         Caption         =   "Modulo de Gastos"
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
         Left            =   2310
         TabIndex        =   2
         Top             =   210
         Width           =   1905
      End
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
      Height          =   855
      Left            =   8610
      Picture         =   "Isubctas.frx":001A
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   525
      Width           =   960
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
      Height          =   855
      Left            =   8610
      Picture         =   "Isubctas.frx":045C
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   1470
      Width           =   960
   End
   Begin VB.TextBox TextSubCta 
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
      MaxLength       =   60
      TabIndex        =   13
      Text            =   "0"
      Top             =   3780
      Width           =   6210
   End
   Begin MSAdodcLib.Adodc AdoSubCta1 
      Height          =   330
      Left            =   210
      Top             =   1785
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
   Begin MSMask.MaskEdBox MBoxCta 
      Height          =   330
      Left            =   3465
      TabIndex        =   25
      Top             =   5565
      Width           =   1800
      _ExtentX        =   3175
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
   Begin MSAdodcLib.Adodc AdoCatalogo 
      Height          =   330
      Left            =   210
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
      Caption         =   "Catalogo"
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
   Begin MSMask.MaskEdBox MBFechaF 
      Height          =   330
      Left            =   4935
      TabIndex        =   18
      Top             =   4200
      Visible         =   0   'False
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
   Begin MSMask.MaskEdBox MBFechaI 
      Height          =   330
      Left            =   2835
      TabIndex        =   16
      Top             =   4200
      Visible         =   0   'False
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
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&Desde"
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
      Left            =   2100
      TabIndex        =   15
      Top             =   4200
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&Hasta"
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
      TabIndex        =   17
      Top             =   4200
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      DrawMode        =   2  'Blackness
      X1              =   0
      X2              =   9870
      Y1              =   4620
      Y2              =   4620
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " NIVEL No."
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
      Left            =   3465
      TabIndex        =   9
      Top             =   3360
      Width           =   1170
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " REEMBOLSO"
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
      TabIndex        =   19
      Top             =   4725
      Width           =   1380
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "NOMBRE DE LA CUENTA"
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
      Left            =   210
      TabIndex        =   26
      Top             =   5985
      Width           =   7575
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " CUENTA RELACIONADA"
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
      TabIndex        =   24
      Top             =   5565
      Width           =   3270
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " VALOR"
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
      TabIndex        =   21
      Top             =   5145
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SUBCUENTA DE BLOQUE"
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
      TabIndex        =   5
      Top             =   1050
      Width           =   8415
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " SUBCUENTA"
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
      Top             =   3780
      Width           =   1380
   End
   Begin VB.Label Label2 
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
      Height          =   330
      Left            =   105
      TabIndex        =   7
      Top             =   3360
      Width           =   1380
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   8820
      Top             =   2520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   7
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Isubctas.frx":0766
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Isubctas.frx":0878
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Isubctas.frx":098A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Isubctas.frx":0A9C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Isubctas.frx":0FAE
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Isubctas.frx":14C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Isubctas.frx":19D2
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "ISubCtas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CheqBloquear_Click()
    If CheqBloquear.value = 0 Then
       Label9.Visible = False
       Label10.Visible = False
       MBFechaI.Visible = False
       MBFechaF.Visible = False
    Else
       Label9.Visible = True
       Label10.Visible = True
       MBFechaI.Visible = True
       MBFechaF.Visible = True
    End If
End Sub

Private Sub Command1_Click()
  GrabarCta TxtCodigo
  If OpcI.value Then
     ListarSubCtas "I"
  ElseIf OpcG.value Then
     ListarSubCtas "G"
  ElseIf OpcPM.value Then
     ListarSubCtas "PM"
  Else
     ListarSubCtas "CC"
  End If
End Sub

Private Sub Command2_Click()
  Unload ISubCtas
End Sub

Private Sub DLCtas_DblClick()
  SiguienteControl
End Sub

Private Sub DLCtas_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DLCtas_LostFocus()
  Cadena = SinEspaciosIzq(DLCtas)
  Codigo1 = SinEspaciosIzq(DLCtas)
  LlenarCta Cadena
End Sub

Private Sub Form_Activate()
  Label1.Visible = True
  FormatoMaskCta MBoxCta
  TextPresupuesto.Visible = True
  ListarSubCtas "I"
  RatonNormal
End Sub

Private Sub Form_Load()
  CentrarForm ISubCtas
  ConectarAdodc AdoSubCta
  ConectarAdodc AdoSubCta1
  ConectarAdodc AdoCatalogo
End Sub

Private Sub MBFechaI_GotFocus()
  MarcarTexto MBFechaI
End Sub

Private Sub MBFechaI_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBFechaI_LostFocus()
  FechaValida MBFechaI
End Sub

Private Sub MBFechaF_GotFocus()
  MarcarTexto MBFechaF
End Sub

Private Sub MBFechaF_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBFechaF_LostFocus()
  FechaValida MBFechaF
End Sub

Private Sub OpcCC_Click()
  ListarSubCtas "CC"
  Label1.Visible = False
  TextPresupuesto.Visible = False
  TxtCodigo.Enabled = True
  TxtNivel.Enabled = True 'False
  CheqCaja.Visible = False
End Sub

Private Sub OpcPM_Click()
  ListarSubCtas "PM"
  Label1.Visible = False
  TextPresupuesto.Visible = False
  TxtCodigo.Enabled = False
  TxtNivel.Enabled = True
  CheqCaja.Visible = False
End Sub

Private Sub OpcG_Click()
  ListarSubCtas "G"
  Label1.Visible = True
  TextPresupuesto.Visible = True
  TxtCodigo.Enabled = False
  TxtNivel.Enabled = True
  CheqCaja.Visible = True
End Sub

Private Sub OpcI_Click()
  ListarSubCtas "I"
  Label1.Visible = True
  TextPresupuesto.Visible = True
  TxtCodigo.Enabled = False
  TxtNivel.Enabled = True
  CheqCaja.Visible = False
End Sub

Private Sub TextPresupuesto_LostFocus()
  TextPresupuesto.Text = Format(CDbl(TextPresupuesto.Text), "#,##0.00")
End Sub

Private Sub TextSubCta_LostFocus()
  TextoValido TextSubCta
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
 With AdoSubCta.Recordset
 Select Case Button.key
   Case "Eliminar"
        If DLCtas.Enabled Then
           Cadena = SinEspaciosIzq(DLCtas.Text)
           sSQL = "SELECT Codigo " _
                & "FROM Trans_SubCtas " _
                & "WHERE  Codigo = '" & Cadena & "' " _
                & "AND Item = '" & NumEmpresa & "' " _
                & "AND Periodo = '" & Periodo_Contable & "' "
           Select_Adodc AdoSubCta1, sSQL
           If AdoSubCta1.Recordset.RecordCount > 0 Then
              Mensajes = "No se puede eliminar esta SubCuenta," & vbCrLf _
                       & "porque tiene cuentas procesables."
              MsgBox Mensajes
           Else
              Mensajes = "Esta seguro que desea eliminar la " & vbCrLf _
                       & "Cuenta No. [" & Cadena & "]"
              Titulo = "Pregunta de Eliminacion"
              If BoxMensaje = vbYes Then
                 sSQL = "DELETE * " _
                      & "FROM Catalogo_SubCtas " _
                      & "WHERE Codigo = '" & Cadena & "' " _
                      & "AND Periodo = '" & Periodo_Contable & "' " _
                      & "AND Item = '" & NumEmpresa & "' "
                 Ejecutar_SQL_SP sSQL
              End If
         End If
         If OpcG.value Then TipoCta = "G"
         If OpcI.value Then TipoCta = "I"
         If OpcPM.value Then TipoCta = "PM"
         If OpcCC.value Then TipoCta = "CC"
         ListarSubCtas TipoCta
        End If
   Case "Nuevo"
        NuevaCta
        Nuevo = True
        TxtNivel = "00"
        CheqNivel.value = 0
        DLCtas.Enabled = False
        If OpcCC.value Then
           MarcarTexto TxtCodigo
           TxtCodigo.SetFocus
        Else
           TxtNivel.SetFocus
        End If
   Case "Grabar"
        GrabarCta TxtCodigo
   Case "Primero"
        Nuevo = False
       .MoveFirst
        DLCtas.Text = .fields("Nombre_Cta")
   Case "Anterior"
        Nuevo = False
       .MovePrevious
        If .BOF Then .MoveFirst
        DLCtas.Text = .fields("Nombre_Cta")
   Case "Siguiente"
        Nuevo = False
       .MoveNext
        If .EOF Then .MoveLast
        DLCtas.Text = .fields("Nombre_Cta")
   Case "Ultimo"
        Nuevo = False
       .MoveLast
        DLCtas.Text = .fields("Nombre_Cta")
 End Select
 If Nuevo = False Then
    Cadena = SinEspaciosIzq(DLCtas.Text)
    LlenarCta Cadena
 End If
 End With
End Sub

Public Sub LlenarCta(CodigoCta As String)
   TxtCodigo = NumEmpresa & "0000000"
   sSQL = "SELECT * " _
        & "FROM Catalogo_SubCtas " _
        & "WHERE Codigo = '" & CodigoCta & "' " _
        & "AND Item = '" & NumEmpresa & "' " _
        & "AND Periodo = '" & Periodo_Contable & "' "
   Select_Adodc AdoSubCta1, sSQL
   With AdoSubCta1.Recordset
    If .RecordCount > 0 Then
        TextSubCta.Text = .fields("Detalle")
        TxtCodigo = .fields("Codigo")
        Select Case .fields("TC")
          Case "G": OpcG.value = True
          Case "I": OpcI.value = True
          Case "PM": OpcPM.value = True
          Case "CC": OpcCC.value = True
        End Select
        MBoxCta.Text = FormatoCodigoCta(.fields("Cta_Reembolso"))
        Label5.Caption = " " & Listar_Detalle_Cta(.fields("Cta_Reembolso"))
        'MsgBox Listar_Detalle_Cta(.Fields("Cta"))
        TxtReembolso = .fields("Reembolso")
        TxtNivel = .fields("Nivel")
        MBFechaI = .fields("Fecha_D")
        MBFechaF = .fields("Fecha_H")
        TextPresupuesto = Format(.fields("Presupuesto"), "#,##0.00")
        If .fields("Caja") Then CheqCaja.value = 1 Else CheqCaja.value = 0
        If .fields("Agrupacion") Then CheqNivel.value = 1 Else CheqNivel.value = 0
        If .fields("Bloquear") Then CheqBloquear.value = 1 Else CheqBloquear.value = 0
    Else
        DLCtas.Enabled = False
        TextSubCta.Text = ""
        TxtCodigo = NumEmpresa & "0000000"
        Nuevo = True
        TextSubCta.SetFocus
    End If
   End With
   DLCtas.Enabled = True
End Sub

Public Sub NuevaCta()
  DLCtas.Enabled = False
  TextSubCta.Text = ""
  TxtCodigo = NumEmpresa & "0000000"
End Sub

Public Sub GrabarCta(CodigoCta As String)
  If OpcG.value Then TipoCta = "G"
  If OpcI.value Then TipoCta = "I"
  If OpcPM.value Then TipoCta = "PM"
  If OpcCC.value Then TipoCta = "CC"
  If TxtNivel = Ninguno Then TxtNivel = "00"
  
  sSQL = "SELECT * " _
       & "FROM Catalogo_SubCtas " _
       & "WHERE Codigo = '" & CodigoCta & "' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND TC = '" & TipoCta & "' "
  Select_Adodc AdoSubCta1, sSQL
  If AdoSubCta1.Recordset.RecordCount > 0 Then
     Codigo = AdoSubCta1.Recordset.fields("Codigo")
  Else
     If TipoCta = "CC" Then
        Codigo = CodigoCta
     Else
        Numero = ReadSetDataNum("SubCtas", True, True)
        Codigo = FormatoCodigo(TextSubCta.Text, Numero)
     End If
  End If
  sSQL = "DELETE * " _
       & "FROM Catalogo_SubCtas " _
       & "WHERE Codigo = '" & Codigo & "' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND TC = '" & TipoCta & "' "
  Ejecutar_SQL_SP sSQL
  
  SetAddNew AdoSubCta1
  SetFields AdoSubCta1, "TC", TipoCta
  SetFields AdoSubCta1, "Codigo", Codigo
  SetFields AdoSubCta1, "Detalle", TextSubCta
  SetFields AdoSubCta1, "Presupuesto", CCur(TextPresupuesto)
  SetFields AdoSubCta1, "Caja", adFalse
  SetFields AdoSubCta1, "Agrupacion", adFalse
  SetFields AdoSubCta1, "Bloquear", adFalse
  SetFields AdoSubCta1, "Nivel", TxtNivel
  SetFields AdoSubCta1, "Item", NumEmpresa
  SetFields AdoSubCta1, "Periodo", Periodo_Contable
  SetFields AdoSubCta1, "Cta_Reembolso", CambioCodigoCta(MBoxCta.Text)
  SetFields AdoSubCta1, "Reembolso", TxtReembolso
  SetFields AdoSubCta1, "Fecha_D", MBFechaI
  SetFields AdoSubCta1, "Fecha_H", MBFechaF
  If CheqCaja.value = 1 Then SetFields AdoSubCta1, "Caja", adTrue
  If CheqNivel.value = 1 Then SetFields AdoSubCta1, "Agrupacion", adTrue
  If CheqBloquear.value = 1 Then SetFields AdoSubCta1, "Bloquear", adTrue
  SetUpdate AdoSubCta1
  Nuevo = False
  MsgBox "Grabación Exitosa"
End Sub

Public Sub ListarSubCtas(TipoCta As String)
  TxtCodigo = NumEmpresa & "0000000"
  TextSubCta.Text = ""
  sSQL = "SELECT * " _
       & "FROM Catalogo_Cuentas " _
       & "WHERE MidStrg(Codigo,1,1) IN ('4', '5')  " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "ORDER BY Codigo "
  Select_Adodc AdoCatalogo, sSQL
  
  sSQL = "SELECT (Codigo & Space(5) & Detalle & Space(60-LEN(Detalle)) & Nivel) As Nombre_Cta " _
       & "FROM Catalogo_SubCtas " _
       & "WHERE TC = '" & TipoCta & "' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' "
  If TipoCta = "CC" Then
     sSQL = sSQL & "ORDER BY Codigo, Detalle "
  Else
     sSQL = sSQL & "ORDER BY Nivel,Agrupacion DESC,Detalle,Codigo "
  End If
  SelectDB_List DLCtas, AdoSubCta, sSQL, "Nombre_Cta"
  If AdoSubCta.Recordset.RecordCount > 0 Then
     DLCtas.Enabled = True
     DLCtas.SetFocus
  End If
End Sub

Public Function Listar_Detalle_Cta(FCta As String) As String
  With AdoCatalogo.Recordset
   If .RecordCount > 0 Then
       If FCta = "" Then FCta = Ninguno
      .MoveFirst
      .Find ("Codigo = '" & FCta & "' ")
       If Not .EOF Then
          Listar_Detalle_Cta = .fields("Cuenta")
       Else
          Listar_Detalle_Cta = ""
       End If
   End If
  End With
End Function

Private Sub TxtCodigo_GotFocus()
  MarcarTexto TxtCodigo
End Sub

Private Sub TxtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtCodigo_LostFocus()
  TextoValido TxtCodigo, , True
  If Len(TxtCodigo) > 1 And DLCtas.Enabled = False Then
     sSQL = "SELECT Codigo " _
          & "FROM Catalogo_SubCtas " _
          & "WHERE Codigo = '" & TxtCodigo & "' " _
          & "AND Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' "
     Select_Adodc AdoSubCta1, sSQL
     If AdoSubCta1.Recordset.RecordCount > 0 Then MsgBox "Este codigo ya existe, seleccione otro"
  End If
End Sub

Private Sub TxtNivel_GotFocus()
  MarcarTexto TxtNivel
End Sub

Private Sub TxtNivel_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtNivel_LostFocus()
  'TxtNivel = Format(CByte(TxtNivel), "00")
End Sub

Private Sub TxtReembolso_GotFocus()
  MarcarTexto TxtReembolso
End Sub

Private Sub TxtReembolso_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtReembolso_LostFocus()
  TextoValido TxtReembolso, , True
End Sub
