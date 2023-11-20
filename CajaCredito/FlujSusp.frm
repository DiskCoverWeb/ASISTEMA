VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FlujoSuspenso 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "FLUJO DE SUSPENSOS"
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10620
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   10620
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   960
      Left            =   210
      TabIndex        =   0
      Top             =   525
      Width           =   1485
      Begin MSMask.MaskEdBox MBoxFecha 
         Height          =   330
         Left            =   105
         TabIndex        =   2
         Top             =   525
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
      Begin VB.Label Label4 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " &FECHA:"
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
         Top             =   210
         Width           =   1275
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6315
      Left            =   105
      TabIndex        =   3
      Top             =   105
      Width           =   10410
      _ExtentX        =   18362
      _ExtentY        =   11139
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "&1.- SUSPENSOS"
      TabPicture(0)   =   "FlujSusp.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Command1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Command4"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "DGSuspensos"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "&2.- CHEQUES"
      TabPicture(1)   =   "FlujSusp.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DGCheques"
      Tab(1).Control(1)=   "Command2"
      Tab(1).Control(2)=   "Command3"
      Tab(1).ControlCount=   3
      Begin MSDataGridLib.DataGrid DGSuspensos 
         Bindings        =   "FlujSusp.frx":0038
         Height          =   4740
         Left            =   105
         TabIndex        =   9
         Top             =   1470
         Width           =   10200
         _ExtentX        =   17992
         _ExtentY        =   8361
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
      Begin VB.CommandButton Command3 
         Caption         =   "&Imprimir"
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
         Left            =   -72165
         Picture         =   "FlujSusp.frx":0053
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   525
         Width           =   1065
      End
      Begin VB.CommandButton Command2 
         Caption         =   "C&onsultar"
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
         Left            =   -73320
         Picture         =   "FlujSusp.frx":06BD
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   525
         Width           =   1065
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&Imprimir"
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
         Left            =   2835
         Picture         =   "FlujSusp.frx":0AFF
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   525
         Width           =   1065
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Consultar"
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
         Left            =   1680
         Picture         =   "FlujSusp.frx":1169
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   525
         Width           =   1065
      End
      Begin MSDataGridLib.DataGrid DGCheques 
         Bindings        =   "FlujSusp.frx":15AB
         Height          =   4740
         Left            =   -74895
         TabIndex        =   10
         Top             =   1470
         Width           =   10200
         _ExtentX        =   17992
         _ExtentY        =   8361
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
   End
   Begin MSAdodcLib.Adodc AdoCheques 
      Height          =   330
      Left            =   630
      Top             =   2625
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
      Caption         =   "Cheques"
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
   Begin MSAdodcLib.Adodc AdoSuspensos 
      Height          =   330
      Left            =   630
      Top             =   2310
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
      Caption         =   "Suspensos"
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
   Begin VB.CommandButton Command5 
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
      Left            =   4095
      Picture         =   "FlujSusp.frx":15C4
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   630
      Width           =   1065
   End
End
Attribute VB_Name = "FlujoSuspenso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
  FechaValida MBoxFecha
  Mifecha = BuscarFecha(MBoxFecha.Text)
  sSQL = "SELECT TC.ME,TC.Fecha,TC.TP,TC.Cuenta_No,TC.Debitos,TC.Creditos,TC.CodigoU " _
       & "FROM Trans_Libretas As TC,Catalogo_Proceso As TP " _
       & "WHERE TC.Fecha = #" & Mifecha & "# " _
       & "AND TC.T <> 'A' " _
       & "AND TP.Mi_Cta = " & Val(adFalse) & " " _
       & "AND TC.TP = TP.TP " _
       & "AND TC.ACC = " & Val(adFalse) & " " _
       & "ORDER BY TC.ME,TC.Fecha,TC.TP,TC.Cuenta_No "
  SelectDataGrid DGSuspensos, AdoSuspensos, sSQL
End Sub

Private Sub Command2_Click()
  FechaValida MBoxFecha
  Mifecha = BuscarFecha(MBoxFecha.Text)
  sSQL = "SELECT ME,Fecha,TP,Cuenta_No,Debitos,Creditos,CodigoU " _
       & "FROM Trans_Libretas " _
       & "WHERE Fecha = #" & Mifecha & "# " _
       & "AND T <> 'A' " _
       & "AND CHT <> " & Val(adFalse) & " " _
       & "ORDER BY ME,Fecha,TP,Cuenta_No "
  SelectDataGrid DGCheques, AdoCheques, sSQL
End Sub

Private Sub Command3_Click()
  SQLMsg1 = "REPORTE DE DEPOSITOS EN CHEQUES"
  ImprimirFlujoCajaCoop AdoCheques, True, 1, 9, True, True, 0, False
End Sub

Private Sub Command4_Click()
  SQLMsg1 = "REPORTE DE SUSPENSOS"
  ImprimirFlujoCajaCoop AdoSuspensos, True, 1, 9, True, True, 0, False
End Sub

Private Sub Command5_Click()
  Unload FlujoSuspenso
End Sub

Private Sub DGCheques_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF1 Then GenerarDataTexto FlujoSuspenso, AdoCheques
End Sub

Private Sub DGSuspensos_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF1 Then GenerarDataTexto FlujoSuspenso, AdoSuspensos
End Sub

Private Sub Form_Activate()
  RatonNormal
End Sub

Private Sub Form_Load()
  CentrarForm FlujoSuspenso
  ConectarAdodc AdoCheques
  ConectarAdodc AdoSuspensos
End Sub

Private Sub MBoxFecha_GotFocus()
  MarcarTexto MBoxFecha
End Sub

Private Sub MBoxFecha_LostFocus()
  FechaValida MBoxFecha
End Sub

