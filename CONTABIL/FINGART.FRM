VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.1#0"; "COMCTL32.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FIngArt 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ingreso de Articulo"
   ClientHeight    =   5565
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11355
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5565
   ScaleWidth      =   11355
   ShowInTaskbar   =   0   'False
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   25
      Top             =   0
      Width           =   11355
      _ExtentX        =   20029
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327680
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
      MouseIcon       =   "FIngArt.frx":0000
   End
   Begin VB.TextBox TextProveedor 
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
      Left            =   7665
      TabIndex        =   43
      Top             =   1575
      Width           =   3585
   End
   Begin VB.Frame Frame3 
      Caption         =   "Datos del Articulo para Facturación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2640
      Left            =   4095
      TabIndex        =   33
      Top             =   2730
      Width           =   3690
      Begin VB.CheckBox CheckIVA 
         Caption         =   "Facturar con I.V.A."
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
         Left            =   1890
         TabIndex        =   36
         Top             =   1890
         Width           =   1695
      End
      Begin VB.TextBox TextPVP 
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
         Left            =   105
         MaxLength       =   10
         MultiLine       =   -1  'True
         TabIndex        =   35
         Text            =   "FIngArt.frx":001C
         Top             =   2100
         Width           =   1275
      End
      Begin VB.TextBox TextLinea 
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
         MaxLength       =   35
         TabIndex        =   34
         Top             =   630
         Width           =   3480
      End
      Begin MSMask.MaskEdBox MBoxCtaInv 
         Height          =   330
         Left            =   105
         TabIndex        =   37
         Top             =   1365
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   582
         _Version        =   327680
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
      Begin MSMask.MaskEdBox MBoxContCta 
         Height          =   330
         Left            =   1890
         TabIndex        =   38
         Top             =   1365
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   582
         _Version        =   327680
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
      Begin VB.Label Label12 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " CTA. DE INV.:"
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
         TabIndex        =   42
         Top             =   1050
         Width           =   1695
      End
      Begin VB.Label Label15 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " P.V.P.:"
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
         TabIndex        =   41
         Top             =   1785
         Width           =   1275
      End
      Begin VB.Label Label16 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " LINEA DEL PRODUCTO:"
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
         TabIndex        =   40
         Top             =   315
         Width           =   3480
      End
      Begin VB.Label Label11 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " CONTRA CTA.:"
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
         TabIndex        =   39
         Top             =   1050
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Conversion para Costos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2010
      Left            =   7875
      TabIndex        =   20
      Top             =   2730
      Width           =   3375
      Begin VB.TextBox TextUnidad_Conv 
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
         TabIndex        =   31
         Top             =   1575
         Width           =   1170
      End
      Begin VB.TextBox TextUnidades 
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
         Left            =   2205
         MaxLength       =   10
         MultiLine       =   -1  'True
         TabIndex        =   26
         Text            =   "FIngArt.frx":001E
         Top             =   630
         Width           =   1065
      End
      Begin VB.Frame Frame2 
         Height          =   855
         Left            =   1365
         TabIndex        =   28
         Top             =   1050
         Width           =   1905
         Begin VB.Line Line1 
            X1              =   840
            X2              =   1785
            Y1              =   525
            Y2              =   525
         End
         Begin VB.Label Label19 
            Caption         =   "                U * G"
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
            TabIndex        =   30
            Top             =   525
            Width           =   1695
         End
         Begin VB.Label Label18 
            Caption         =   " VUnit = VUnit * D"
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
            TabIndex        =   29
            Top             =   315
            Width           =   1695
         End
      End
      Begin VB.TextBox TextGramos 
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
         Left            =   1155
         MaxLength       =   10
         MultiLine       =   -1  'True
         TabIndex        =   23
         Text            =   "FIngArt.frx":0020
         Top             =   630
         Width           =   960
      End
      Begin VB.TextBox TexTDensidad 
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
         Left            =   105
         MaxLength       =   10
         MultiLine       =   -1  'True
         TabIndex        =   21
         Text            =   "FIngArt.frx":0022
         Top             =   630
         Width           =   960
      End
      Begin VB.Label Label20 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " UNIDAD A CONVERTIR:"
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
         Left            =   105
         TabIndex        =   32
         Top             =   1155
         Width           =   1170
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "&Unidades:"
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
         Left            =   2205
         TabIndex        =   27
         Top             =   315
         Width           =   1065
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "&Gramos:"
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
         TabIndex        =   24
         Top             =   315
         Width           =   960
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "&Densidad:"
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
         Top             =   315
         Width           =   960
      End
   End
   Begin MSDBCtls.DBList DBLArt 
      Bindings        =   "FIngArt.frx":0024
      DataSource      =   "DataArticulo"
      Height          =   4545
      Left            =   105
      TabIndex        =   0
      Top             =   840
      Width           =   3900
      _ExtentX        =   6879
      _ExtentY        =   8017
      _Version        =   327680
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
   Begin VB.Data DataArtFact 
      Caption         =   "ArtFact"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1890
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4620
      Visible         =   0   'False
      Width           =   1905
   End
   Begin MSMask.MaskEdBox MBoxCta 
      Height          =   330
      Left            =   4095
      TabIndex        =   7
      Top             =   2310
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   582
      _Version        =   327680
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
   Begin VB.TextBox TextMaximo 
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
      Left            =   8820
      TabIndex        =   9
      Top             =   2310
      Width           =   1170
   End
   Begin VB.TextBox TextMinimo 
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
      Left            =   10080
      TabIndex        =   17
      Top             =   2310
      Width           =   1170
   End
   Begin VB.TextBox TextBodega 
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
      Left            =   7665
      TabIndex        =   10
      Top             =   2310
      Width           =   1065
   End
   Begin VB.TextBox TextUnidad 
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
      Left            =   6300
      TabIndex        =   11
      Top             =   2310
      Width           =   1275
   End
   Begin VB.TextBox TextArticulo 
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
      Left            =   4095
      TabIndex        =   8
      Top             =   1575
      Width           =   3480
   End
   Begin MSDBCtls.DBCombo DBCTipoInv 
      Bindings        =   "FIngArt.frx":003B
      DataSource      =   "DataInv"
      Height          =   315
      Left            =   5565
      TabIndex        =   1
      Top             =   840
      Width           =   3480
      _ExtentX        =   6138
      _ExtentY        =   556
      _Version        =   327680
      Text            =   " "
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
   Begin VB.Data DataSQL 
      Caption         =   "SQL"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   105
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4620
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Data DataInv 
      Caption         =   "Inv"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1890
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4935
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.Data DataArticulo 
      Caption         =   "Articulo"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   105
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4305
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Data DataCodigos 
      Caption         =   "Codigos"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1890
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4305
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.Data DataArt 
      Caption         =   "Art"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   105
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4935
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
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
      Left            =   9975
      TabIndex        =   19
      Top             =   4935
      Width           =   1275
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Grabar"
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
      Left            =   8610
      TabIndex        =   18
      Top             =   4935
      Width           =   1275
   End
   Begin VB.Label Label10 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " PROVEEDOR DEL ARTICULO:"
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
      Left            =   7665
      TabIndex        =   44
      Top             =   1260
      Width           =   3585
   End
   Begin VB.Label Label9 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " MINIMO:"
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
      Left            =   10080
      TabIndex        =   16
      Top             =   1995
      Width           =   1170
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " BODEGA:"
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
      Left            =   7665
      TabIndex        =   6
      Top             =   1995
      Width           =   1065
   End
   Begin VB.Label LabelCodigo 
      BackColor       =   &H80000009&
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
      Left            =   4095
      TabIndex        =   12
      Top             =   840
      Width           =   1380
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " CUENTA DE VENTAS:"
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
      Left            =   4095
      TabIndex        =   14
      Top             =   1995
      Width           =   2115
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " MAXIMO:"
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
      Left            =   8820
      TabIndex        =   5
      Top             =   1995
      Width           =   1170
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " UNIDAD:"
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
      Left            =   6300
      TabIndex        =   4
      Top             =   1995
      Width           =   1275
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " NOMBRE DEL ARTICULO:"
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
      Left            =   4095
      TabIndex        =   3
      Top             =   1260
      Width           =   3480
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " TIPO DE INVENTARIO:"
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
      Left            =   5565
      TabIndex        =   15
      Top             =   525
      Width           =   3480
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327680
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   7
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FIngArt.frx":004D
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FIngArt.frx":015F
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FIngArt.frx":0271
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FIngArt.frx":0383
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FIngArt.frx":0895
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FIngArt.frx":0DA7
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FIngArt.frx":12B9
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " CODIGO:"
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
      Left            =   4095
      TabIndex        =   13
      Top             =   525
      Width           =   1380
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " ARTICULO"
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
      Top             =   525
      Width           =   3900
   End
End
Attribute VB_Name = "FIngArt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub LlenarArticulo(CodArt As String)
  RatonReloj MDIConta
  Invisible
  LabelCodigo.Visible = True
  DBLArt.Enabled = True
  sSQL = "SELECT * FROM Articulo "
  sSQL = sSQL & "WHERE Codigo = '" & CodArt & "' "
  SelectData DataArt, sSQL, False
  With DataArt.Recordset
   If .RecordCount > 0 Then
       TipoInventario = .Fields("TP")
       LabelCodigo.Caption = .Fields("Codigo")
       TextArticulo.Text = .Fields("Articulo")
       TextUnidad.Text = .Fields("Unidad")
       TextBodega.Text = .Fields("Bodega")
       TextMinimo.Text = .Fields("Minimo")
       TextMaximo.Text = .Fields("Maximo")
       TextProveedor.Text = .Fields("Proveedor")
       MBoxCta.Text = FormatoCodigoCta(.Fields("Codigo_Cta"))
       sSQL = "SELECT * FROM Inventario "
       sSQL = sSQL & "WHERE TP = '" & TipoInventario & "' "
       SelectData DataArt, sSQL, False
       DBCTipoInv.Text = DataArt.Recordset.Fields("Concepto")
       If TipoInventario = PT Then
          sSQL = "SELECT * FROM Articulo "
          sSQL = sSQL & "WHERE Codigo = '" & Codigo & "' "
          SelectData DataArtFact, sSQL, False
          With DataArtFact.Recordset
           If .RecordCount > 0 Then
               ArticuloPT
               TextLinea.Text = .Fields("Linea")
               TextPVP.Text = .Fields("PVP")
               MBoxCta.Text = FormatoCodigoCta(.Fields("Cta_C"))
               MBoxCtaInv.Text = FormatoCodigoCta(.Fields("Cta_K"))
               MBoxContCta.Text = FormatoCodigoCta(.Fields("Cta_PP"))
               If .Fields("IVA") Then CheckIVA.Value = 1 Else CheckIVA.Value = 0
           End If
          End With
       End If
    End If
  End With
  Nuevo = False
  RatonNormal MDIConta
End Sub

Public Sub NuevoArticulo()
  Nuevo = True
  LabelCodigo.Caption = ""
  TextArticulo.Text = ""
  TextUnidad.Text = ""
  TextBodega.Text = ""
  TextMinimo.Text = ""
  TextMaximo.Text = ""
  TextProveedor.Text = ""
  Invisible
  FormatoMaskCta MBoxCta
End Sub

Public Sub GrabarArticulo(CodArt As String)
  RatonReloj MDIConta
  DBLArt.Enabled = True
  Nuevo = False
  LabelCodigo.Caption = CodArt
  If CodArt = "" Then CodArt = Ninguno
  sSQL = "SELECT * FROM Articulo "
  sSQL = sSQL & "WHERE Codigo = '" & CodArt & "' "
  SelectData DataArtFact, sSQL, False
  sSQL = "SELECT * FROM Articulo "
  sSQL = sSQL & "WHERE Codigo = '" & CodArt & "' "
  SelectData DataArt, sSQL, False
  With DataArt.Recordset
       If .RecordCount > 0 Then
          .Edit
           If TipoInventario = PT Then DataArtFact.Recordset.Edit
       Else
          .AddNew
           Numero = ReadSetDataNum(DataCodigos, "Articulos", True)
           Codigo = FormatoCodigo(TextArticulo.Text, Numero)
          .Fields("Codigo") = Codigo
           If TipoInventario = PT Then
              DataArtFact.Recordset.AddNew
              DataArtFact.Recordset.Fields("Codigo") = Codigo
           End If
       End If
      .Fields("Articulo") = TextArticulo.Text
      .Fields("Unidad") = TextUnidad.Text
      .Fields("Bodega") = TextBodega.Text
      .Fields("Minimo") = TextMinimo.Text
      .Fields("Maximo") = TextMaximo.Text
      .Fields("Proveedor") = TextProveedor.Text
      .Fields("Codigo_Cta") = CambioCodigoCta(MBoxCta.Text)
       sSQL = "SELECT * FROM Inventario "
       sSQL = sSQL & "WHERE Concepto = '" & DBCTipoInv.Text & "' "
       SelectData DataSQL, sSQL, False
      .Fields("TP") = DataSQL.Recordset.Fields("TP")
      .Update
       If TipoInventario = PT Then
          DataArtFact.Recordset.Fields("Articulo") = TextArticulo.Text
          DataArtFact.Recordset.Fields("Linea") = TextLinea.Text
          DataArtFact.Recordset.Fields("PVP") = TextPVP.Text
          DataArtFact.Recordset.Fields("IVA") = CheckIVA.Value
          DataArtFact.Recordset.Fields("Stock") = 0
          DataArtFact.Recordset.Fields("Cta_C") = CambioCodigoCta(MBoxCta.Text)
          DataArtFact.Recordset.Fields("Cta_K") = CambioCodigoCta(MBoxCtaInv.Text)
          DataArtFact.Recordset.Fields("Cta_PP") = CambioCodigoCta(MBoxContCta.Text)
          DataArtFact.Recordset.Update
       End If
       sSQL = "SELECT Codigo & ' ' & Articulo As NombArt FROM Articulo "
       sSQL = sSQL & "ORDER BY TP,Articulo "
       SelectData DataArticulo, sSQL, False
       LabelCodigo.Caption = Codigo
  End With
  LabelCodigo.Visible = True
  RatonNormal MDIConta
End Sub

Private Sub Command1_Click()
  GrabarArticulo Codigo
End Sub

Private Sub Command2_Click()
  Unload FIngArt
End Sub

Private Sub DBCTipoInv_LostFocus()
   TipoInventario = MP
   sSQL = "SELECT * FROM Inventario "
   sSQL = sSQL & "WHERE Concepto = '" & DBCTipoInv.Text & "' "
   SelectData DataSQL, sSQL, False
   If DataSQL.Recordset.RecordCount > 0 Then TipoInventario = DataSQL.Recordset.Fields("TP")
   If TipoInventario = PT Then ArticuloPT Else Invisible
End Sub

Private Sub DBLArt_DblClick()
  SendKeys "{TAB}", False
End Sub

Private Sub DBLArt_KeyDown(KeyCode As Integer, Shift As Integer)
 PresionoEnter KeyCode
End Sub

Private Sub DBLArt_LostFocus()
 If (Nuevo = False) And (DataArticulo.Recordset.RecordCount > 0) Then
    Codigo = SinEspaciosIzq(DBLArt.Text)
    LlenarArticulo SinEspaciosIzq(Codigo)
 End If
End Sub

Private Sub Form_Activate()
   sSQL = "SELECT * FROM Articulo "
   SelectData DataArtFact, sSQL, False
   sSQL = "SELECT Codigo & ' ' & Articulo As NombArt FROM Articulo "
   sSQL = sSQL & "ORDER BY TP,Articulo "
   SelectDBList DBLArt, DataArticulo, sSQL, "NombArt"
   sSQL = "SELECT Concepto FROM Inventario "
   sSQL = sSQL & "ORDER BY Concepto "
   SelectDBCombo DBCTipoInv, DataInv, sSQL, "Concepto"
   FormatoMaskCta MBoxCta
   FormatoMaskCta MBoxCtaInv
   FormatoMaskCta MBoxContCta
   Invisible
   RatonNormal MDIConta
End Sub

Private Sub Form_Load()
   CentrarForm FIngArt
   DataArt.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
   DataInv.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
   DataSQL.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
   DataArticulo.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
   DataArtFact.DatabaseName = RutaEmpresa & "\PRODUCC.MDB"
   DataCodigos.DatabaseName = RutaEmpresa & "\SETEOS.MDB"
End Sub

Private Sub TextArticulo_LostFocus()
 If TextArticulo.Text = "" Then TextArticulo.Text = Ninguno
End Sub

Private Sub TextBodega_LostFocus()
  If TextBodega.Text = "" Then TextBodega.Text = Ninguno
End Sub

Private Sub TexTDensidad_LostFocus()
  If TexTDensidad.Text = "" Then TexTDensidad.Text = "1"
End Sub

Private Sub TextGramos_LostFocus()
  If TextGramos.Text = "" Then TextGramos.Text = "1"
End Sub

Private Sub TextLinea_LostFocus()
  If TextLinea.Text = "" Then TextLinea.Text = niguno
End Sub

Private Sub TextProveedor_LostFocus()
  If TextProveedor.Text = "" Then TextProveedor.Text = niguno
End Sub

Private Sub TextUnidad_Conv_LostFocus()
  If TextUnidad_Conv.Text = "" Then TextUnidad_Conv.Text = Ninguno
End Sub

Private Sub TextUnidad_LostFocus()
  If TextUnidad.Text = "" Then TextUnidad.Text = Ninguno
End Sub

Private Sub TextUnidades_GotFocus()
  TextUnidades.Text = ""
End Sub

Private Sub TextUnidades_LostFocus()
  If TextUnidades.Text = "" Then TextUnidades.Text = "1"
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
 Select Case Button.Key
   Case "Eliminar"
      If DBLArt.Enabled Then
         sSQL = "SELECT Codigo FROM Kardex "
         sSQL = sSQL & "WHERE  Codigo = '" & Codigo & "' "
         SelectData DataArt, sSQL, False
         If DataArt.Recordset.RecordCount > 0 Then
            Mensajes = "No se puede eliminar este Producto (" & Codigo & ")," & Chr(13)
            Mensajes = Mensajes & "porque tiene valores procesables."
            MsgBox Mensajes
         Else
            Mensajes = "Esta seguro que desea eliminar" & Chr(13)
            Mensajes = Mensajes & "El Producto [" & Codigo & "]"
            Titulo = "Pregunta de Eliminacion"
            If BoxMensaje = 6 Then
               sSQL = "DELETE * FROM Articulo "
               sSQL = sSQL & "WHERE Codigo = '" & Codigo & "' "
               DeleteData DataArt, sSQL
               sSQL = "DELETE * FROM Articulo "
               sSQL = sSQL & "WHERE Codigo = '" & Codigo & "' "
               DeleteData DataArtFact, sSQL
               sSQL = "SELECT Codigo & ' ' & Articulo As NombArt FROM Articulo "
               sSQL = sSQL & "ORDER BY TP,Articulo "
               SelectData DataArticulo, sSQL, False
            End If
         End If
         DBLArt.SetFocus
      End If
   Case "Nuevo"
       NuevoArticulo
       DBLArt.Enabled = False
       LabelCodigo.Visible = False
       DBCTipoInv.SetFocus
       Codigo = Ninguno
   Case "Grabar"
       GrabarArticulo Codigo
   Case "Primero"
       DataArticulo.Recordset.MoveFirst
   Case "Anterior"
       DataArticulo.Recordset.MovePrevious
       If DataArticulo.Recordset.BOF Then DataArticulo.Recordset.MoveFirst
   Case "Siguiente"
       DataArticulo.Recordset.MoveNext
       If DataArticulo.Recordset.EOF Then DataArticulo.Recordset.MoveLast
   Case "Ultimo"
       DataArticulo.Recordset.MoveLast
 End Select
 If (Nuevo = False) And (DataArticulo.Recordset.RecordCount > 0) Then
    DBLArt.Text = DataArticulo.Recordset.Fields(0)
    Codigo = SinEspaciosIzq(DBLArt.Text)
    LlenarArticulo Codigo
 End If
End Sub

Public Sub Invisible()
  TextLinea.Text = ""
  TextPVP.Text = "0"
  FormatoMaskCta MBoxCtaInv
  FormatoMaskCta MBoxContCta
  CheckIVA.Visible = 0
  Frame1.Visible = False
  Frame3.Visible = False
End Sub

Public Sub ArticuloPT()
  Frame1.Visible = True
  Frame3.Visible = True
End Sub
