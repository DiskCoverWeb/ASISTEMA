VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form FResSusc 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "RESUMEN DE SUSCRIPCIONES"
   ClientHeight    =   7050
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10515
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7050
   ScaleWidth      =   10515
   ShowInTaskbar   =   0   'False
   Begin MSDataGridLib.DataGrid DGQuery 
      Bindings        =   "FResSusc.frx":0000
      Height          =   5685
      Left            =   105
      TabIndex        =   19
      Top             =   1260
      Width           =   8730
      _ExtentX        =   15399
      _ExtentY        =   10028
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      RowDividerStyle =   1
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
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
            LCID            =   2058
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
            LCID            =   2058
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
      Caption         =   "Por:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1170
      Left            =   6405
      TabIndex        =   12
      Top             =   0
      Width           =   2430
      Begin VB.OptionButton OpcSemestral 
         Caption         =   "Sem&estral"
         Height          =   210
         Left            =   1260
         TabIndex        =   22
         Top             =   840
         Width           =   1080
      End
      Begin VB.OptionButton OpcTrimestral 
         Caption         =   "&Trimestral"
         Height          =   210
         Left            =   1260
         TabIndex        =   21
         Top             =   525
         Width           =   1080
      End
      Begin VB.OptionButton OpcAnual 
         Caption         =   "&Anual"
         Height          =   210
         Left            =   1260
         TabIndex        =   20
         Top             =   210
         Width           =   975
      End
      Begin VB.OptionButton OpcMensual 
         Caption         =   "&Mensual"
         Height          =   210
         Left            =   105
         TabIndex        =   13
         Top             =   210
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton OpcQuincenal 
         Caption         =   "&Quincenal"
         Height          =   210
         Left            =   105
         TabIndex        =   14
         Top             =   525
         Width           =   1080
      End
      Begin VB.OptionButton OpcSemanal 
         Caption         =   "Se&manal"
         Height          =   210
         Left            =   105
         TabIndex        =   15
         Top             =   840
         Width           =   975
      End
   End
   Begin VB.CommandButton Command3 
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
      Left            =   8925
      Picture         =   "FResSusc.frx":0017
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   1995
      Width           =   1065
   End
   Begin VB.Frame Frame3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1170
      Left            =   105
      TabIndex        =   0
      Top             =   0
      Width           =   6315
      Begin MSMask.MaskEdBox MBFechaF 
         Height          =   330
         Left            =   840
         TabIndex        =   4
         Top             =   630
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
      Begin VB.OptionButton OpcNueva 
         Caption         =   "&Nuevas"
         Height          =   210
         Left            =   2205
         TabIndex        =   7
         Top             =   840
         Width           =   975
      End
      Begin VB.OptionButton OpcAnul 
         Caption         =   "&Anuladas"
         Height          =   210
         Left            =   2205
         TabIndex        =   5
         Top             =   210
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton OpcSusp 
         Caption         =   "&Suspensas"
         Height          =   210
         Left            =   2205
         TabIndex        =   6
         Top             =   525
         Width           =   1185
      End
      Begin VB.OptionButton OpcCanc 
         Caption         =   "&Terminadas"
         Height          =   210
         Left            =   3465
         TabIndex        =   9
         Top             =   525
         Width           =   1185
      End
      Begin VB.OptionButton OpcRenov 
         Caption         =   "&Renovaciones"
         Height          =   210
         Left            =   3465
         TabIndex        =   8
         Top             =   210
         Width           =   1395
      End
      Begin VB.OptionButton OpcTodas 
         Caption         =   "Ac&tivas"
         Height          =   210
         Left            =   3465
         TabIndex        =   10
         Top             =   840
         Width           =   990
      End
      Begin VB.OptionButton OpcListarT 
         Caption         =   "Lista Todas"
         Height          =   210
         Left            =   5040
         TabIndex        =   11
         Top             =   210
         Width           =   1200
      End
      Begin MSMask.MaskEdBox MBFechaI 
         Height          =   330
         Left            =   840
         TabIndex        =   2
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
      Begin VB.Label Label6 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " &Hasta:"
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
         TabIndex        =   3
         Top             =   630
         Width           =   750
      End
      Begin VB.Label Label5 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " &Desde:"
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
         Width           =   750
      End
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Resumir S&uscrip."
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
      Left            =   8925
      Picture         =   "FResSusc.frx":0A0D
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   105
      Width           =   1065
   End
   Begin VB.CommandButton Command9 
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
      Left            =   8925
      Picture         =   "FResSusc.frx":0D17
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   1050
      Width           =   1065
   End
   Begin MSAdodcLib.Adodc AdoQuery 
      Height          =   330
      Left            =   420
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
      Caption         =   "Listado de Facturas"
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
End
Attribute VB_Name = "FResSusc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command3_Click()
  Unload FResSusc
End Sub

Private Sub Command8_Click()
RatonReloj
DGQuery.Visible = False
FechaValida MBFechaI
FechaValida MBFechaF
FechaIni = BuscarFecha(MBFechaI.Text)
FechaFin = BuscarFecha(MBFechaF.Text)
If OpcMensual.value Then TipoComp = "MENS"
If OpcQuincenal.value Then TipoComp = "QUNC"
If OpcSemanal.value Then TipoComp = "SEMA"
If OpcAnual.value Then TipoComp = "ANUA"
If OpcTrimestral.value Then TipoComp = "TRIM"
If OpcSemestral.value Then TipoComp = "SEME"
If OpcListarT.value <> True Then
   If OpcNueva.value Then
      SQLMsg1 = "LISTADO DE SUSCRIPCIONES NUEVAS"
      TipoDoc = "N"
   End If
   If OpcRenov.value Then
      SQLMsg1 = "LISTADO DE SUSCRIPCIONES RENOVACIONES"
      TipoDoc = "R"
   End If
   If OpcAnul.value Then
      SQLMsg1 = "LISTADO DE SUSCRIPCIONES ANULADAS"
      TipoDoc = "A"
   End If
   If OpcSusp.value Then
      SQLMsg1 = "LISTADO DE SUSCRIPCIONES SUSPENSAS"
      TipoDoc = "S"
   End If
   If OpcCanc.value Then
      SQLMsg1 = "LISTADO DE SUSCRIPCIONES TERMINADAS"
      TipoDoc = "T"
   End If
   If OpcTodas.value Then
      SQLMsg1 = "LISTADO DE SUSCRIPCIONES ACTIVAS"
      TipoDoc = "X"
   End If
Else
   SQLMsg1 = "LISTADO DE TODAS LAS SUSCRIPCIONES"
   TipoDoc = "Y"
End If
sSQL = "SELECT C.Ciudad,Count(P.Credito_No) As Cantidad " _
     & "FROM Prestamos As P,Clientes As C " _
     & "WHERE P.Item = '" & NumEmpresa & "' " _
     & "AND P.Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
     & "AND P.TP = '" & TipoComp & "' " _
     & "AND P.Cuenta_No = C.Codigo "
If OpcListarT.value <> True Then
   If TipoDoc <> "X" Then
      sSQL = sSQL & "AND P.T = '" & TipoDoc & "' "
   Else
      sSQL = sSQL & "AND P.T <> 'A' " _
           & "AND P.T <> 'S' " _
           & "AND P.T <> 'T' "
   End If
End If
sSQL = sSQL & "GROUP BY C.Ciudad " _
     & "ORDER BY C.Ciudad "
Select_Adodc_Grid DGQuery, AdoQuery, sSQL
DGQuery.Caption = SQLMsg1
DGQuery.Visible = True
RatonNormal
End Sub

Private Sub Command9_Click()
   DGQuery.Visible = False
   SQLMsg2 = "Del " & MBFechaI.Text & " al " & MBFechaF.Text
   ImprimirAdodc AdoQuery, 1, 10
   DGQuery.Visible = True
End Sub

Private Sub Form_Activate()
   RatonNormal
End Sub

Private Sub Form_Load()
  CentrarForm FResSusc
  ConectarAdodc AdoQuery
End Sub

Private Sub MBFechaF_GotFocus()
  MarcarTexto MBFechaF
End Sub

Private Sub MBFechaF_LostFocus()
  FechaValida MBFechaF
End Sub

Private Sub MBFechaI_GotFocus()
  MarcarTexto MBFechaI
End Sub

Private Sub MBFechaI_LostFocus()
  FechaValida MBFechaI
End Sub
