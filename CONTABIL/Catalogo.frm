VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.4#0"; "comctl32.Ocx"
Begin VB.Form CatalogoCtas 
   Caption         =   "PRESENTACION DEL CATALOGO DE CUENTA"
   ClientHeight    =   11235
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   19770
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11235
   ScaleWidth      =   19770
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   19770
      _ExtentX        =   34872
      _ExtentY        =   1164
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   4
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Salir"
            Object.ToolTipText     =   "Salir del modulo"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Consultar"
            Object.ToolTipText     =   "Procesar consulta"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Imprimir"
            Object.ToolTipText     =   "Imprimir Resultados"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Excel"
            Object.ToolTipText     =   "Exportar a Excel"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
      EndProperty
      Begin VB.Frame Frame1 
         Caption         =   "Tipo de Consulta"
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
         Left            =   8925
         TabIndex        =   6
         Top             =   0
         Width           =   3795
         Begin VB.OptionButton OpcG 
            Caption         =   "De Grupo"
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
            Left            =   1155
            TabIndex        =   8
            Top             =   315
            Width           =   1170
         End
         Begin VB.OptionButton OpcD 
            Caption         =   "De Detalle"
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
            Left            =   2415
            TabIndex        =   9
            Top             =   315
            Width           =   1275
         End
         Begin VB.OptionButton OpcT 
            Caption         =   "Todas"
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
            TabIndex        =   7
            Top             =   315
            Value           =   -1  'True
            Width           =   960
         End
      End
      Begin VB.Frame Frame2 
         Height          =   645
         Left            =   2415
         TabIndex        =   1
         Top             =   0
         Width           =   6420
         Begin MSMask.MaskEdBox MBoxCtaI 
            Height          =   330
            Left            =   1470
            TabIndex        =   3
            Top             =   210
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   582
            _Version        =   393216
            AllowPrompt     =   -1  'True
            AutoTab         =   -1  'True
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
         Begin MSMask.MaskEdBox MBoxCtaF 
            Height          =   330
            Left            =   4620
            TabIndex        =   5
            Top             =   210
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   582
            _Version        =   393216
            AllowPrompt     =   -1  'True
            AutoTab         =   -1  'True
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
         Begin VB.Label Label3 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Cuenta &Final"
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
            Left            =   3255
            TabIndex        =   4
            Top             =   210
            Width           =   1380
         End
         Begin VB.Label Label4 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Cuenta I&nicial"
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
            Width           =   1380
         End
      End
   End
   Begin MSDataGridLib.DataGrid DGCatalogo 
      Bindings        =   "Catalogo.frx":0000
      Height          =   4950
      Left            =   105
      TabIndex        =   10
      Top             =   735
      Width           =   10620
      _ExtentX        =   18733
      _ExtentY        =   8731
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   21
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
         Name            =   "Courier New"
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
   Begin MSAdodcLib.Adodc AdoCatalogo 
      Height          =   330
      Left            =   105
      Top             =   5985
      Width           =   10200
      _ExtentX        =   17992
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
   Begin ComctlLib.ImageList ImageList1 
      Left            =   315
      Top             =   945
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   4
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Catalogo.frx":001A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Catalogo.frx":0334
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Catalogo.frx":064E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Catalogo.frx":0968
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "CatalogoCtas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub ListarCatalogoCuentas()
  Codigo = CambioCodigoCta(MBoxCtaI.Text)
  Codigo1 = CambioCodigoCta(MBoxCtaF.Text)
  If Codigo = "0" Or Codigo = "" Then Codigo = "1"
  If Codigo1 = "0" Or Codigo1 = "" Then Codigo1 = "9"
  
  SQLMsg1 = "PLAN DE CUENTAS"
  sSQL = "SELECT Clave,TC,ME,DG,Codigo,Cuenta,Presupuesto,Codigo_Ext,Cod_Rol_Pago,I_E_Emp, Con_IESS " _
       & "FROM Catalogo_Cuentas " _
       & "WHERE Cuenta <> '" & Ninguno & "' " _
       & "AND Codigo BETWEEN '" & Codigo & "' and '" & Codigo1 & "' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' "
  If OpcG.value Then
     sSQL = sSQL & "AND DG = 'G' "
  ElseIf OpcD.value Then
     sSQL = sSQL & "AND DG = 'D' "
  End If
  sSQL = sSQL & "ORDER BY Codigo "
  Select_Adodc_Grid DGCatalogo, AdoCatalogo, sSQL
End Sub

Private Sub Form_Activate()
  RatonReloj
  If Supervisor = False Then Command1.Enabled = CNivel(6)
  FormatoMaskCta MBoxCtaI
  FormatoMaskCta MBoxCtaF
  ListarCatalogoCuentas
  MBoxCtaI.SetFocus
  RatonNormal
End Sub

Private Sub Form_Load()
  'CentrarForm CatalogoCtas
  ConectarAdodc AdoCatalogo
  DGCatalogo.Height = MDI_Y_Max - DGCatalogo.Top - 300
  DGCatalogo.width = MDI_X_Max - DGCatalogo.Left
  AdoCatalogo.Top = DGCatalogo.Top + DGCatalogo.Height
  AdoCatalogo.width = MDI_X_Max - DGCatalogo.Left
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
   'MsgBox Button.key
    Select Case Button.key
      Case "Salir"
           Unload CatalogoCtas
      Case "Consultar"
           ListarCatalogoCuentas
      Case "Imprimir"
           RatonReloj
           DGCatalogo.Visible = False
           Imprimir_Catalogo AdoCatalogo
           ListarCatalogoCuentas
           DGCatalogo.Visible = True
           RatonNormal
      Case "Excel"
           DGCatalogo.Visible = False
           GenerarDataTexto CatalogoCtas, AdoCatalogo
           DGCatalogo.Visible = True
    End Select
End Sub
