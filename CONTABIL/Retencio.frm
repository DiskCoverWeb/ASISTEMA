VERSION 5.00
Object = "{05BFD3F1-6319-4F30-B752-C7A22889BCC4}#1.0#0"; "AcroPDF.dll"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.5#0"; "comctl32.Ocx"
Begin VB.Form Retenciones 
   BackColor       =   &H00C0FFC0&
   Caption         =   "LISTAR COMPROBANTE DE RETENCION DE EGRESOS"
   ClientHeight    =   9435
   ClientLeft      =   180
   ClientTop       =   465
   ClientWidth     =   15960
   DrawMode        =   1  'Blackness
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9435
   ScaleWidth      =   15960
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   15960
      _ExtentX        =   28152
      _ExtentY        =   1164
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   10
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Salir"
            Object.ToolTipText     =   "SAlir del modullo"
            Object.Tag             =   ""
            ImageIndex      =   11
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Imprimir"
            Object.ToolTipText     =   "Imprimir Retencion Fisica"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Primero"
            Object.ToolTipText     =   "Primer Comprobante"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Anterior"
            Object.ToolTipText     =   "Anterior"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Siguiente"
            Object.ToolTipText     =   "Siguiente"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Ultimo"
            Object.ToolTipText     =   "Ultimo Comprobante"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Cambio_Retencion"
            Object.ToolTipText     =   "Cambia Numero de Retención"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Autorizar"
            Object.ToolTipText     =   "Autorizar Retencion Pendiente"
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Autorizar_Grupo"
            Object.ToolTipText     =   "Autoriza Retenciones en Lote"
            Object.Tag             =   ""
            ImageIndex      =   9
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Enviar_Mails"
            Object.ToolTipText     =   "Envia Retencion por Emails"
            Object.Tag             =   ""
            ImageIndex      =   10
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.Frame Frame1 
         Height          =   750
         Left            =   5880
         TabIndex        =   17
         Top             =   -105
         Width           =   7260
         Begin VB.TextBox TxtDocHasta 
            Alignment       =   1  'Right Justify
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
            Left            =   5460
            TabIndex        =   19
            Text            =   "0"
            Top             =   210
            Width           =   1695
         End
         Begin VB.TextBox TxtDocDesde 
            Alignment       =   1  'Right Justify
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
            Left            =   1890
            TabIndex        =   18
            Text            =   "0"
            Top             =   210
            Width           =   1800
         End
         Begin VB.Label Label8 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Documento Hasta"
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
            Left            =   3780
            TabIndex        =   21
            Top             =   210
            Width           =   1695
         End
         Begin VB.Label Label4 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Documento  Desde"
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
            Top             =   210
            Width           =   1800
         End
      End
   End
   Begin AcroPDFLibCtl.AcroPDF APDFRetencion 
      Height          =   4635
      Left            =   105
      TabIndex        =   23
      Top             =   2625
      Width           =   6840
      _cx             =   5080
      _cy             =   5080
   End
   Begin VB.CheckBox CheqClaveAcceso 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      Caption         =   "Clave de Accceso:"
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
      Left            =   420
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   1260
      Width           =   3060
   End
   Begin VB.OptionButton OpcManuales 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Manuales"
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
      Left            =   3360
      TabIndex        =   2
      Top             =   735
      Width           =   1170
   End
   Begin VB.OptionButton OpcAut 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Autorizados"
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
      Left            =   105
      TabIndex        =   0
      Top             =   735
      Value           =   -1  'True
      Width           =   1380
   End
   Begin VB.ComboBox CMeses 
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
      Height          =   315
      Left            =   5145
      TabIndex        =   4
      Text            =   "Diciembre"
      Top             =   735
      Width           =   1380
   End
   Begin VB.OptionButton OpcNoAut 
      BackColor       =   &H00C0FFC0&
      Caption         =   "No Autorizados"
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
      Left            =   1575
      TabIndex        =   1
      Top             =   735
      Width           =   1695
   End
   Begin VB.TextBox TxtAutorizacion 
      Height          =   330
      Left            =   9555
      MaxLength       =   50
      TabIndex        =   15
      ToolTipText     =   "<Ctrl+A> Autorizar manualmente desde el SRI"
      Top             =   1260
      Width           =   4740
   End
   Begin VB.TextBox TxtClave 
      Height          =   330
      Left            =   3465
      MaxLength       =   50
      TabIndex        =   13
      ToolTipText     =   "<Ctrl+A> Volver a Generar y Firmar el Documento Electronico"
      Top             =   1260
      Width           =   4740
   End
   Begin MSDataListLib.DataCombo DCComp 
      Bindings        =   "Retencio.frx":0000
      DataSource      =   "AdoComp"
      Height          =   345
      Left            =   9135
      TabIndex        =   8
      Top             =   735
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   609
      _Version        =   393216
      Text            =   "999999999"
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
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
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
      Left            =   105
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1260
      Width           =   330
   End
   Begin MSAdodcLib.Adodc AdoComp1 
      Height          =   330
      Left            =   15225
      Top             =   1785
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
      Caption         =   "Comp1"
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
   Begin MSAdodcLib.Adodc AdoComp 
      Height          =   330
      Left            =   15225
      Top             =   2100
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
      Caption         =   "Comp"
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
   Begin MSAdodcLib.Adodc AdoDetRet 
      Height          =   330
      Left            =   17220
      Top             =   1785
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
      Caption         =   "DetRet"
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
   Begin MSAdodcLib.Adodc AdoDetCom 
      Height          =   330
      Left            =   17220
      Top             =   2100
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
      Caption         =   "DetCom"
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
   Begin MSDataListLib.DataCombo DCSerie 
      Bindings        =   "Retencio.frx":0016
      DataSource      =   "AdoSerie"
      Height          =   345
      Left            =   7560
      TabIndex        =   6
      Top             =   735
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   609
      _Version        =   393216
      Text            =   "999999"
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
   Begin MSAdodcLib.Adodc AdoSerie 
      Height          =   330
      Left            =   19215
      Top             =   1785
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
      Caption         =   "Serie"
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
   Begin MSAdodcLib.Adodc AdoTP_Num 
      Height          =   330
      Left            =   19215
      Top             =   2100
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
      Caption         =   "TP_Num"
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
   Begin MSDataListLib.DataCombo DCTP_Num 
      Bindings        =   "Retencio.frx":002D
      DataSource      =   "AdoTP_Num"
      Height          =   345
      Left            =   11340
      TabIndex        =   10
      Top             =   735
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   609
      _Version        =   393216
      Text            =   "999999999"
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
   Begin VB.Label LblResultado 
      BackColor       =   &H0000FF00&
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
      Height          =   855
      Left            =   105
      TabIndex        =   24
      Top             =   1680
      Width           =   14190
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &T.C."
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
      Left            =   10815
      TabIndex        =   9
      Top             =   735
      Width           =   540
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &No."
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
      Left            =   8715
      TabIndex        =   7
      Top             =   735
      Width           =   435
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &No."
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
      Left            =   4620
      TabIndex        =   3
      Top             =   735
      Width           =   540
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Autorizacion:"
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
      Left            =   8295
      TabIndex        =   14
      Top             =   1260
      Width           =   1275
   End
   Begin VB.Label LabelEst 
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
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   13230
      TabIndex        =   11
      Top             =   735
      Width           =   1590
   End
   Begin VB.Label Label12 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &Serie No."
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
      Left            =   6615
      TabIndex        =   5
      Top             =   735
      Width           =   960
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   14490
      Top             =   1785
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   11
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Retencio.frx":0045
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Retencio.frx":035F
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Retencio.frx":DC71
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Retencio.frx":1B583
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Retencio.frx":28E95
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Retencio.frx":367A7
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Retencio.frx":36AC1
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Retencio.frx":64973
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Retencio.frx":64C8D
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Retencio.frx":64FA7
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Retencio.frx":A6EA9
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Retenciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Retencion_No As Long
Dim Serie_R As String
 
'Dim ObjAutori As New WS_Autorizacion

Dim ArrayAutorizacion() As String
Dim DocDesde As Long
Dim DocHasta As Long
'Dim SRI_Aut As Tipo_Estado_SRI
Dim Resultado_Ret As String
 
Private Sub Command1_Click()
   Unload Retenciones
End Sub

Private Sub TxtDocDesde_GotFocus()
  MarcarTexto TxtDocDesde
End Sub

Private Sub TxtDocDesde_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtDocDesde_LostFocus()
  TextoValido TxtDocDesde, True, , 0
End Sub

Private Sub TxtDocHasta_GotFocus()
  MarcarTexto TxtDocHasta
End Sub

Private Sub TxtDocHasta_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtDocHasta_LostFocus()
  TextoValido TxtDocHasta, True, , 0
End Sub

Private Sub DCComp_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCComp_LostFocus()
    Listar_Tipo_Comp_Retencion
End Sub

Private Sub DCSerie_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub DCSerie_LostFocus()
   Listar_Tipo_Retencion DCSerie
End Sub

Private Sub DCTP_Num_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  PresionoEnter KeyCode
End Sub

Private Sub DCTP_Num_LostFocus()
   If Not IsNumeric(DCComp) Then DCComp = "0"
   Listar_Retencion DCSerie, Val(DCComp), SinEspaciosIzq(DCTP_Num), Val(SinEspaciosDer(DCTP_Num))
End Sub

Private Sub Form_Activate()
   sSQL = "UPDATE Trans_Compras " _
        & "SET Estado_SRI = 'OK' " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND Periodo = '" & Periodo_Contable & "' " _
        & "AND LEN(Clave_Acceso) = 49 " _
        & "AND LEN(AutRetencion) = 37 " _
        & "AND Estado_SRI <> 'OK' "
   Ejecutar_SQL_SP sSQL
   
   LblResultado.width = Retenciones.width - 400
   APDFRetencion.width = Retenciones.width - 400
   APDFRetencion.Height = MDI_Y_Max - APDFRetencion.Top - 100
   
   APDFRetencion.Object.src = ""
   RatonNormal
End Sub

Private Sub Form_Load()
  'LstResultado.width = MDI_X_Max - 250
   ConectarAdodc AdoComp
   ConectarAdodc AdoDetRet
   ConectarAdodc AdoDetCom
   ConectarAdodc AdoComp1
   ConectarAdodc AdoSerie
   ConectarAdodc AdoTP_Num
   
   CMeses.Clear
   CMeses.AddItem "Todos"
   CMeses.AddItem "Enero"
   CMeses.AddItem "Febrero"
   CMeses.AddItem "Marzo"
   CMeses.AddItem "Abril"
   CMeses.AddItem "Mayo"
   CMeses.AddItem "Junio"
   CMeses.AddItem "Julio"
   CMeses.AddItem "Agosto"
   CMeses.AddItem "Septiembre"
   CMeses.AddItem "Octubre"
   CMeses.AddItem "Noviembre"
   CMeses.AddItem "Diciembre"
   CMeses.Text = "Todos"
   
  'Pagina de Conexion con el SRI
   SRI_Obtener_Datos_Comprobantes_Electronicos
  'CentrarForm Retenciones
End Sub

Public Sub Listar_Retencion(Serie_R As String, Comp_No As Long, TP As String, TP_No As Long)
Dim TipoProc As String
   RatonReloj
  'Listar las Retenciones del IVA
   TxtAutorizacion = ""
   TxtClave = ""
   sSQL = "SELECT C.Cliente,C.CI_RUC,C.TD,C.Direccion,C.Email,C.Email2,C.Ciudad,C.DirNumero,C.Telefono,TC.* " _
        & "FROM Trans_Compras As TC, Clientes As C " _
        & "WHERE TC.Item = '" & NumEmpresa & "' " _
        & "AND TC.Periodo = '" & Periodo_Contable & "' " _
        & "AND TC.SecRetencion = " & Comp_No & " " _
        & "AND TC.Serie_Retencion = '" & Serie_R & "' " _
        & "AND TC.Numero = " & Val(TP_No) & " " _
        & "AND TC.TP = '" & TP & "' " _
        & "AND TC.IdProv = C.Codigo " _
        & "ORDER BY Cta_Servicio,Cta_Bienes "
   Select_Adodc AdoDetCom, sSQL
   With AdoDetCom.Recordset
    If .RecordCount > 0 Then
       'MsgBox .RecordCount
        Co.Item = NumEmpresa
        Co.Fecha = .Fields("Fecha")
        Co.Beneficiario = .Fields("Cliente")
        Co.RUC_CI = .Fields("CI_RUC")
        Co.Direccion = .Fields("Direccion")
        Co.TD = .Fields("TD")
        Co.Email = .Fields("Email")
        Co.TP = .Fields("TP")
        Co.Numero = .Fields("Numero")
        
        FA.EmailC = .Fields("Email")
        FA.EmailR = .Fields("Email2")
        FA.TP = Co.TP
        FA.Numero = Co.Numero
        FA.TR = "RE"
        FA.TC = "FA"
        FA.Fecha = .Fields("FechaEmision")
        FA.Fecha_R = .Fields("FechaEmision")
        FA.Serie_R = .Fields("Serie_Retencion")
        FA.Retencion = .Fields("SecRetencion")
        FA.ClaveAcceso = .Fields("Clave_Acceso")
        FA.Estado_SRI = .Fields("Estado_SRI")
        FA.Autorizacion_R = .Fields("AutRetencion")
        FA.Serie = .Fields("Establecimiento") & .Fields("PuntoEmision")
        FA.Factura = .Fields("Secuencial")
        
        FechaTexto = Co.Fecha
        TxtAutorizacion = FA.Autorizacion_R
        TxtClave = FA.ClaveAcceso
        CheqClaveAcceso.Caption = " Clave de Acceso (" & FA.Estado_SRI & ")"
        SetNombrePRN = ""
        SRI_Generar_PDF_RE FA, False
        APDFRetencion.Object.src = RutaDocumentoPDF
      ' Presentar_PDF ATSPDF, RutaDocumentoPDF
        RatonNormal
    End If
   End With
   LblResultado.Caption = SRI_Leer_Comprobantes_no_Autorizados(FA.Autorizacion_R)
  'LblResultado.Caption = SRI_Leer_Comprobantes_no_Autorizados("2604202507070216417900110010030000000421234567814")   ' FA.Autorizacion_R
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
  TextoImprimio = ""
  DocDesde = Val(TxtDocDesde.Text)
  DocHasta = Val(TxtDocHasta.Text)
  Select Case Button.key
    Case "Salir"
         Unload Retenciones
    Case "Imprimir"
         NumComp = Co.Numero
         Co.Item = NumEmpresa
         ImprimirComprobantesDe True, Co
         DCComp.SetFocus
    Case "Primero"
         AdoComp.Recordset.MoveFirst
         DCComp = AdoComp.Recordset.Fields("SecRetencion")
         Listar_Tipo_Comp_Retencion
         Listar_Retencion DCSerie, Val(DCComp), SinEspaciosIzq(DCTP_Num), SinEspaciosDer(DCTP_Num)
    Case "Anterior"
         AdoComp.Recordset.MovePrevious
         If AdoComp.Recordset.BOF Then AdoComp.Recordset.MoveFirst
         DCComp = AdoComp.Recordset.Fields("SecRetencion")
         Listar_Tipo_Comp_Retencion
         Listar_Retencion DCSerie, Val(DCComp), SinEspaciosIzq(DCTP_Num), SinEspaciosDer(DCTP_Num)
    Case "Siguiente"
         AdoComp.Recordset.MoveNext
         If AdoComp.Recordset.EOF Then AdoComp.Recordset.MoveLast
         DCComp = AdoComp.Recordset.Fields("SecRetencion")
         Listar_Tipo_Comp_Retencion
         Listar_Retencion DCSerie, Val(DCComp), SinEspaciosIzq(DCTP_Num), SinEspaciosDer(DCTP_Num)
    Case "Ultimo"
         AdoComp.Recordset.MoveLast
         DCComp = AdoComp.Recordset.Fields("SecRetencion")
         Listar_Tipo_Comp_Retencion
         Listar_Retencion DCSerie, Val(DCComp), SinEspaciosIzq(DCTP_Num), SinEspaciosDer(DCTP_Num)
    Case "Cambio_Retencion"
         Cambio_Retencion
    Case "Autorizar"
         FA.TP = Co.TP
         FA.Numero = Co.Numero
         FA.Fecha = Co.Fecha
         FA.Serie_R = DCSerie.Text
         FA.Retencion = Val(DCComp.Text)
         SRI_Crear_Clave_Acceso_Retenciones FA, False
         LblResultado.Caption = SRI_Leer_Comprobantes_no_Autorizados(SRI_Autorizacion.Clave_De_Acceso)
    Case "Autorizar_Grupo"
         If Len(DCSerie) < 6 Then DCSerie = "001001"
         sSQL = "SELECT C.Cliente,C.CI_RUC,C.TD,C.Direccion,C.Email,C.Ciudad,C.DirNumero,C.Telefono,TC.* " _
              & "FROM Trans_Compras As TC,Clientes As C " _
              & "WHERE TC.Item = '" & NumEmpresa & "' " _
              & "AND TC.Periodo = '" & Periodo_Contable & "' " _
              & "AND TC.Serie_Retencion = '" & DCSerie & "' " _
              & "AND LEN(AutRetencion) = 13 "
         If DocDesde > 0 And DocHasta > 0 And DocDesde <= DocHasta Then sSQL = sSQL & "AND TC.SecRetencion BETWEEN " & DocDesde & " and " & DocHasta & " "
         sSQL = sSQL _
              & "AND TC.IdProv = C.Codigo " _
              & "ORDER BY Serie_Retencion,SecRetencion "
         Select_Adodc AdoComp1, sSQL
        'MsgBox sSQL
         With AdoComp1.Recordset
          If .RecordCount > 0 Then
             'MsgBox .RecordCount
              Do While Not .EOF
                 RatonReloj
                 Co.TP = .Fields("TP")
                 Co.Numero = .Fields("Numero")
                 Co.Fecha = .Fields("Fecha")
                 
                 FA.TP = Co.TP
                 FA.Numero = Co.Numero
                 FA.Fecha = .Fields("Fecha")
                 FA.Serie_R = .Fields("Serie_Retencion")
                 FA.Retencion = .Fields("SecRetencion")
                 
                 SRI_Crear_Clave_Acceso_Retenciones FA, False
                 
                'Leemos el Error si lo tubiera
                 LblResultado.Caption = SRI_Leer_Comprobantes_no_Autorizados(SRI_Autorizacion.Clave_De_Acceso)
                 If Len(LblResultado.Caption) > 1 Then TextoImprimio = TextoImprimio & LblResultado.Caption & vbCrLf & String(80, "-") & vbCrLf
                 RatonNormal
                .MoveNext
              Loop
          End If
         End With
         
         Generar_Informe_Errores
         Retenciones.Caption = "LISTAR COMPROBANTES DE RETENCION"
    Case "Enviar_Mails"
'         WBPDF.Navigate "file:///C:/SISTEMA/FONDOS/index_pdf.html"
 '        WBPDF.Refresh
         Titulo = "ENVIAR RETENCION POR MAIL"
         Mensajes = "Enviar por mail el documento:" & vbCrLf _
                  & "Fecha: " & FA.Fecha & ", Retencion No. " & FA.Serie_R & "-" & Format$(FA.Retencion, "000000000") & vbCrLf _
                  & "Clave de Acceso: " & vbCrLf & FA.ClaveAcceso & vbCrLf _
                  & "Autorizacion: " & vbCrLf & FA.Autorizacion_R & vbCrLf & vbCrLf _
                  & "A los correos siguientes: " & vbCrLf
         If Len(FA.EmailC) > 1 Then Mensajes = Mensajes & TrimStrg(FA.EmailC) & ";"
         If FA.EmailC <> FA.EmailR And Len(FA.EmailR) > 1 Then Mensajes = Mensajes & TrimStrg(FA.EmailR) & ";"
         If BoxMensaje = vbYes Then
            FA.Numero = Co.Numero
            FA.TP = Co.TP
            SRI_Autorizacion.Fecha_Autorizacion = Co.Fecha
            SRI_Autorizacion.Autorizacion = FA.Autorizacion_R
            SRI_Enviar_Mails FA, SRI_Autorizacion
            DCComp = AdoComp.Recordset.Fields("SecRetencion")
         End If
  End Select
End Sub

Public Sub Cambio_Retencion()
Dim Cambio_Ret As Long
Dim TPrompt As String
   TPrompt = "CAMBIAR LA RETENCION No. " & Format(Retencion_No, "00000000") & vbCrLf & vbCrLf _
           & "POR EL NUEVO NUMERO DE RETENCION:"
   Cambio_Ret = Val(InputBox(TPrompt, "CAMBIO DE NUMERO DE RETENCION", CStr(Retencion_No)))
   If Cambio_Ret > 0 And (Cambio_Ret <> Retencion_No) Then
      RatonReloj
      sSQL = "UPDATE Trans_Compras " _
           & "SET SecRetencion = " & Cambio_Ret & " " _
           & "WHERE Item = '" & NumEmpresa & "' " _
           & "AND Periodo = '" & Periodo_Contable & "' " _
           & "AND TP = '" & FA.TP & "' " _
           & "AND Numero = " & FA.Numero & " " _
           & "AND Serie_Retencion = '" & FA.Serie_R & "' " _
           & "AND SecRetencion = " & FA.Retencion & " "
      Ejecutar_SQL_SP sSQL
      RatonReloj
      sSQL = "UPDATE Trans_Air " _
           & "SET SecRetencion = " & Cambio_Ret & " " _
           & "WHERE Item = '" & NumEmpresa & "' " _
           & "AND Tipo_Trans IN ('C','I') " _
           & "AND TP = '" & FA.TP & "' " _
           & "AND Numero = " & FA.Numero & " " _
           & "AND EstabRetencion+PtoEmiRetencion = '" & FA.Serie_R & "' " _
           & "AND SecRetencion = " & FA.Retencion & " "
      Ejecutar_SQL_SP sSQL
      RatonNormal
      MsgBox "Proceso Exitoso, vuelva consultar"
   End If
End Sub

Public Sub Listar_Tipo_Retencion(Serie_R As String)
Dim MesNo As Integer
  MesNo = CMeses.ListIndex
  If MesNo < 0 Then MesNo = 0
  If Len(Serie_R) < 6 Then Serie_R = "001001"
  sSQL = "SELECT SecRetencion " _
       & "FROM Trans_Compras " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND Serie_Retencion = '" & Serie_R & "' " _
       & "AND SecRetencion > 0 "
  If OpcNoAut.value Then
      sSQL = sSQL _
          & "AND LEN(AutRetencion) = 13 " _
          & "AND LEN(Clave_Acceso) >= 1 " _
          & "AND Estado_SRI <> 'OK' "
  End If
  If MesNo > 0 Then sSQL = sSQL & "AND MONTH(Fecha) = " & MesNo & " "
  sSQL = sSQL _
       & "GROUP BY SecRetencion " _
       & "ORDER BY SecRetencion "
  SelectDB_Combo DCComp, AdoComp, sSQL, "SecRetencion", True
End Sub

Public Sub Listar_Tipo_Serie_Retencion()
Dim MesNo As Integer
  MesNo = CMeses.ListIndex
  If MesNo < 0 Then MesNo = 0
  sSQL = "SELECT Serie_Retencion " _
       & "FROM Trans_Compras " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND LEN(Serie_Retencion) = 6 "
  If OpcNoAut.value Then
      sSQL = sSQL _
          & "AND LEN(AutRetencion) BETWEEN 13 and 49 " _
          & "AND LEN(Clave_Acceso) >= 1 " _
          & "AND Estado_SRI <> 'OK' "
  ElseIf OpcManuales.value Then
    sSQL = sSQL _
          & "AND LEN(AutRetencion) < 13 " _
          & "AND LEN(Clave_Acceso) = 1 "
  Else
    sSQL = sSQL _
          & "AND LEN(AutRetencion) > 13 " _
          & "AND LEN(Clave_Acceso) > 13 " _
          & "AND Estado_SRI = 'OK' "
  End If
  If MesNo > 0 Then sSQL = sSQL & "AND MONTH(Fecha) = " & MesNo & " "
  sSQL = sSQL _
       & "GROUP BY Serie_Retencion " _
       & "ORDER BY Serie_Retencion "
  SelectDB_Combo DCSerie, AdoSerie, sSQL, "Serie_Retencion"
End Sub

Public Sub Listar_Tipo_Comp_Retencion()
  sSQL = "SELECT (TP + ' ' + CAST(Numero AS VARCHAR)) As TC_Num " _
       & "FROM Trans_Compras " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND Serie_Retencion = '" & DCSerie & "' " _
       & "AND SecRetencion = " & Val(DCComp) & " " _
       & "ORDER BY TP,Numero "
  SelectDB_Combo DCTP_Num, AdoTP_Num, sSQL, "TC_Num", True
End Sub

Private Sub CMeses_GotFocus()
  MarcarTexto CMeses
End Sub

Private Sub CMeses_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub CMeses_LostFocus()
  Listar_Tipo_Serie_Retencion
  RatonNormal
End Sub

Private Sub TxtAutorizacion_KeyDown(KeyCode As Integer, Shift As Integer)
Dim Secuencial As String
    Keys_Especiales Shift
    If CtrlDown And KeyCode = vbKeyA Then
       If CFechaLong(FechaSistema) <= CFechaLong(Fecha_CE) Then
         Secuencial = "INGRESE LA AUTORIZACION DEL" & vbCrLf _
                    & "DOCUMENTO " & Co.Fecha & "-" & Co.TP & "-" & Format$(Co.Numero, "000000000") & vbCrLf _
                    & "Autorizacion: " & TxtAutorizacion & ":"
         Autorizacion = InputBox(Secuencial, "CAMBIO DE AUTORIZACION", TxtAutorizacion)
         If IsNumeric(Autorizacion) And Len(Autorizacion) >= 3 Then
            SQL1 = "UPDATE Trans_Air " _
                 & "SET AutRetencion = '" & Autorizacion & "' " _
                 & "WHERE Item = '" & NumEmpresa & "' " _
                 & "AND Periodo = '" & Periodo_Contable & "' " _
                 & "AND Numero = " & Co.Numero & " " _
                 & "AND TP = '" & Co.TP & "' " _
                 & "AND AutRetencion = '" & TxtAutorizacion & "' " _
                 & "AND EstabRetencion = '" & MidStrg(FA.Serie_R, 1, 3) & "' " _
                 & "AND PtoEmiRetencion = '" & MidStrg(FA.Serie_R, 4, 3) & "' " _
                 & "AND SecRetencion = '" & FA.Retencion & "' "
            Ejecutar_SQL_SP SQL1
        
            SQL1 = "UPDATE Trans_Compras " _
                 & "SET AutRetencion = '" & Autorizacion & "' " _
                 & "WHERE Item = '" & NumEmpresa & "' " _
                 & "AND Periodo = '" & Periodo_Contable & "' " _
                 & "AND Numero = " & Co.Numero & " " _
                 & "AND TP = '" & Co.TP & "' " _
                 & "AND AutRetencion = '" & TxtAutorizacion & "' " _
                 & "AND Serie_Retencion = '" & FA.Serie_R & "' " _
                 & "AND SecRetencion = '" & FA.Retencion & "' "
            Ejecutar_SQL_SP SQL1
            MsgBox "Proceso terminado"
         End If
       Else
         RatonNormal
         MsgBox MensajeNoAutorizarCE
       End If
    End If
End Sub

Private Sub TxtClave_KeyDown(KeyCode As Integer, Shift As Integer)
    Keys_Especiales Shift
    If CtrlDown And KeyCode = vbKeyA Then
       If CFechaLong(FechaSistema) <= CFechaLong(Fecha_CE) Then
          RatonReloj
               '& "AND Estado_SRI <> 'OK' "
          TextoImprimio = ""
          sSQL = "UPDATE Trans_Compras " _
               & "SET Estado_SRI = '.', Clave_Acceso = '.', AutRetencion = '" & RUC & "' " _
               & "WHERE Item = '" & NumEmpresa & "' " _
               & "AND Periodo = '" & Periodo_Contable & "' " _
               & "AND TP = '" & Co.TP & "' " _
               & "AND Numero = " & Co.Numero & " "
          Ejecutar_SQL_SP sSQL
          RutaXMLFirmado = RutaDocumentos & "\Comprobantes Generados\" & FA.ClaveAcceso & ".xml"
          If Dir$(RutaXMLFirmado) <> "" Then Kill RutaXMLFirmado
          
          RutaXMLFirmado = RutaDocumentos & "\Comprobantes no Autorizados\" & FA.ClaveAcceso & ".xml"
          If Dir$(RutaXMLFirmado) <> "" Then Kill RutaXMLFirmado
                    
          RutaXMLFirmado = RutaDocumentos & "\Comprobantes Firmados\" & FA.ClaveAcceso & ".xml"
          If Dir$(RutaXMLFirmado) <> "" Then Kill RutaXMLFirmado

          FA.Numero = Co.Numero
          FA.TP = Co.TP
          FA.Serie_R = DCSerie.Text
          FA.Retencion = Val(DCComp.Text)
          SRI_Crear_Clave_Acceso_Retenciones FA, False, CBool(CheqClaveAcceso.value)
          RatonNormal
          LblResultado.Caption = SRI_Leer_Comprobantes_no_Autorizados(SRI_Autorizacion.Clave_De_Acceso)
          Generar_Informe_Errores
       Else
          RatonNormal
          MsgBox MensajeNoAutorizarCE
       End If
    End If
'''    If CtrlDown And KeyCode = vbKeyW Then
'''       MsgBox Fecha_CE
'''       If CFechaLong(FechaSistema) <= CFechaLong(Fecha_CE) Then
'''          FA.Numero = Co.Numero
'''          FA.TP = Co.TP
'''
'''          SRI_Crear_Clave_Acceso_Retenciones FA, True
'''          RatonNormal
'''       Else
'''          RatonNormal
'''          MsgBox MensajeNoAutorizarCE
'''       End If
'''    End If
End Sub
