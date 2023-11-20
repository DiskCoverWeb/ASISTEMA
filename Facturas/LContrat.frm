VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form ListarContrato 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cartera de Clientes"
   ClientHeight    =   6810
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11265
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6810
   ScaleWidth      =   11265
   ShowInTaskbar   =   0   'False
   Begin MSDataListLib.DataCombo DCGrupo 
      Bindings        =   "LContrat.frx":0000
      DataSource      =   "AdoGrupo"
      Height          =   315
      Left            =   1200
      TabIndex        =   23
      Top             =   1440
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
   End
   Begin MSAdodcLib.Adodc AdoQuery 
      Height          =   330
      Left            =   240
      Top             =   6300
      Width           =   3495
      _ExtentX        =   6165
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
      Caption         =   "Listado de facturas"
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
      Height          =   4260
      Left            =   120
      TabIndex        =   12
      Top             =   2040
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   7514
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "&1.- Prestamos/Contratos"
      TabPicture(0)   =   "LContrat.frx":0017
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "LabelAbonado"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "LabelFacturado"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "&2.- Tablas de Prestamos"
      TabPicture(1)   =   "LContrat.frx":0033
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "AdoTablaPrestamo"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin MSAdodcLib.Adodc AdoTablaPrestamo 
         Height          =   330
         Left            =   -74880
         Top             =   3840
         Width           =   11055
         _ExtentX        =   19500
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
         Caption         =   "Lista de Facturas"
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
      Begin VB.Label LabelFacturado 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   330
         Left            =   5355
         TabIndex        =   16
         Top             =   4305
         Width           =   1905
      End
      Begin VB.Label LabelAbonado 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   330
         Left            =   8925
         TabIndex        =   14
         Top             =   4305
         Width           =   1905
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Total Facturado"
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
         Left            =   3570
         TabIndex        =   17
         Top             =   4305
         Width           =   1800
      End
      Begin VB.Label Label4 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Total x Cobrar"
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
         Left            =   7245
         TabIndex        =   15
         Top             =   4320
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1065
      Left            =   2310
      TabIndex        =   18
      Top             =   0
      Width           =   8625
      Begin MSDataListLib.DataCombo DCCliente 
         Bindings        =   "LContrat.frx":004F
         DataSource      =   "AdoCliente"
         Height          =   315
         Left            =   2520
         TabIndex        =   24
         Top             =   480
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.OptionButton OpcA 
         Caption         =   "&Abonos de Contratos Mensuales"
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
         Left            =   5040
         TabIndex        =   13
         Top             =   210
         Width           =   3270
      End
      Begin VB.OptionButton OpcC 
         Caption         =   "&Contratos Mensuales"
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
         Left            =   1995
         TabIndex        =   20
         Top             =   210
         Width           =   2220
      End
      Begin VB.OptionButton OpcP 
         Caption         =   "&Prestamos"
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
         TabIndex        =   19
         Top             =   210
         Value           =   -1  'True
         Width           =   1275
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "NOMBRE DEL CLIENTE"
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
         TabIndex        =   21
         Top             =   525
         Width           =   2325
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Facturas"
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
      Left            =   2310
      TabIndex        =   9
      Top             =   1155
      Width           =   3900
      Begin VB.OptionButton OpcPend 
         Caption         =   "Pendientes"
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
         Top             =   210
         Value           =   -1  'True
         Width           =   1380
      End
      Begin VB.OptionButton OpcTodas 
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
         Height          =   330
         Left            =   2835
         TabIndex        =   11
         Top             =   210
         Width           =   960
      End
      Begin VB.OptionButton OpcCanc 
         Caption         =   "Elaborados"
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
         TabIndex        =   4
         Top             =   210
         Width           =   1380
      End
   End
   Begin MSMask.MaskEdBox MBoxFechaF 
      Height          =   330
      Left            =   945
      TabIndex        =   1
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
      Height          =   750
      Left            =   9975
      Picture         =   "LContrat.frx":0068
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1155
      Width           =   960
   End
   Begin VB.CommandButton Command2 
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
      Height          =   750
      Left            =   8925
      Picture         =   "LContrat.frx":02EA
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1155
      Width           =   960
   End
   Begin VB.CheckBox CheckTodos 
      Caption         =   "&Listar Todos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   105
      TabIndex        =   2
      Top             =   945
      Value           =   1  'Checked
      Width           =   2100
   End
   Begin MSMask.MaskEdBox MBoxFechaI 
      Height          =   330
      Left            =   945
      TabIndex        =   0
      Top             =   105
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
      Height          =   750
      Left            =   7875
      Picture         =   "LContrat.frx":0954
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1155
      Width           =   960
   End
   Begin MSAdodcLib.Adodc AdoGrupo 
      Height          =   330
      Left            =   120
      Top             =   2280
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
   Begin MSAdodcLib.Adodc AdoCliente 
      Height          =   330
      Left            =   120
      Top             =   2640
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
      Caption         =   "Cliente"
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
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Grupo No."
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
      Top             =   1365
      Width           =   1065
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
      TabIndex        =   8
      Top             =   525
      Width           =   855
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
      TabIndex        =   7
      Top             =   105
      Width           =   855
   End
End
Attribute VB_Name = "ListarContrato"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
RatonReloj
FechaValida MBoxFechaI, False
FechaValida MBoxFechaF, False
Grupo_No = Val(DCGrupo.Text)
FechaIni = BuscarFecha(MBoxFechaI.Text)
FechaFin = BuscarFecha(MBoxFechaF.Text)
Codigo = SinEspaciosIzq(DCCliente.Text)
Total = 0: Saldo = 0
If OpcC.Value Then
   sSQL = "SELECT Co.T,Cliente,Co.Contrato_No,Co.Fecha,Co.Pago_No,Co.Concepto,Monto_MN,Monto_ME "
   sSQL = sSQL & "FROM Contratos_Meses As Co,Clientes As Cl "
   sSQL = sSQL & "WHERE Co.Codigo_C = Cl.Codigo "
   sSQL = sSQL & "AND Cl.Grupo = " & Grupo_No & " "
   If CheckTodos.Value = 0 Then
      sSQL = sSQL & "AND (Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "#) "
      sSQL = sSQL & "AND Contrato_No = '" & Codigo & "' "
   End If
   If OpcPend.Value Then sSQL = sSQL & "AND T = '" & Procesado & "' "
   If OpcCanc.Value Then sSQL = sSQL & "AND T = '" & Cancelado & "' "
   sSQL = sSQL & "ORDER BY Cl.Cliente,Contrato_No,Fecha "
ElseIf OpcA.Value Then
   sSQL = "SELECT Co.T,Cliente,AM.Contrato_No,AM.Fecha,AM.Cuota_No,Co.Concepto,AM.Abono "
   sSQL = sSQL & "FROM Contratos_Meses As Co,Clientes As Cl,Abono_Meses As AM "
   sSQL = sSQL & "WHERE Cl.Grupo = " & Grupo_No & " "
   sSQL = sSQL & "AND Co.Contrato_No = AM.Contrato_No "
   sSQL = sSQL & "AND Co.Codigo_C = Cl.Codigo "
   sSQL = sSQL & "AND Co.Pago_No = AM.Cuota_No "
   If CheckTodos.Value = 0 Then
      sSQL = sSQL & "AND (AM.Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "#) "
      sSQL = sSQL & "AND AM.Contrato_No = '" & Codigo & "' "
   End If
   If OpcPend.Value Then sSQL = sSQL & "AND T = '" & Procesado & "' "
   If OpcCanc.Value Then sSQL = sSQL & "AND T = '" & Cancelado & "' "
   sSQL = sSQL & "ORDER BY Cl.Cliente,AM.Contrato_No,AM.Fecha,AM.Cuota_No "
Else
   sSQL = "SELECT Cliente,Co.* "
   sSQL = sSQL & "FROM Prestamos As Co,Clientes As Cl "
   sSQL = sSQL & "WHERE Co.Codigo_C = Cl.Codigo "
   sSQL = sSQL & "AND Pago_No=0 "
   sSQL = sSQL & "AND Cl.Grupo = " & Grupo_No & " "
   If CheckTodos.Value = 0 Then
      sSQL = sSQL & "AND (Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "#) "
      sSQL = sSQL & "AND Contrato_No = " & Codigo & " "
   End If
   If OpcPend.Value Then sSQL = sSQL & "AND T = '" & Pendiente & "' "
   If OpcCanc.Value Then sSQL = sSQL & "AND T = '" & Cancelado & "' "
   sSQL = sSQL & "ORDER BY Cl.Cliente,Fecha,Contrato_No "
End If
'SelectDataGrid DGQuery, AdoQuery, sSQL
With AdoQuery.Recordset
  Do While Not .EOF
     If OpcC.Value Then
        Total = Total + .Fields("Monto_MN")
        Saldo = Saldo + .Fields("Monto_ME")
     ElseIf OpcA.Value Then
        'Total = Total + .Fields("Monto_MN")
        Saldo = Saldo + .Fields("Abono")
     Else
        Total = Total + .Fields("Abono")
        Saldo = Saldo + .Fields("Saldo")
     End If
    .MoveNext
  Loop
End With

LabelFacturado.Caption = Format(Total, "#,##0")
LabelAbonado.Caption = Format(Saldo, "#,##0")
RatonNormal
End Sub

Private Sub Command2_Click()
   DGQuery.Visible = False
   With Heads
       .MsgObjetivo = "Corte de Contratos: "
       .TextoObjetivo = " desde el " & MBoxFechaI.Text & " al " & MBoxFechaF.Text
       .MsgConcepto = ""
       .TextoConcepto = ""
        If OpcPend.Value Then .MsgTitulo = "LISTADO DE CONTRATOS PENDIENTES"
        If OpcCanc.Value Then .MsgTitulo = "LISTADO DE CONTRATOS ELABORADOS"
        If OpcTodas.Value Then .MsgTitulo = "LISTADO DE TODOS LOS CONTRATOS"
        MiFecha = MBoxFechaF.Text
   End With
   If OpcA.Value Then
      ' ImprimirData AdoQuery, True, 1, 8
   Else
      'ImprimirContratos AdoQuery, True, 1, 8
   End If
   DGQuery.Visible = True
End Sub

Private Sub Command3_Click()
  Unload ListarContrato
End Sub

Private Sub Command4_Click()

End Sub

Private Sub DCGrupo_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCGrupo_LostFocus()
   Grupo_No = Val(DCGrupo.Text)
   sSQL = "SELECT Contrato_No & ' de: ' & Cliente As Contratos "
   sSQL = sSQL & "FROM Prestamos As CM,Clientes As Cl "
   sSQL = sSQL & "WHERE CM.Codigo_C = Cl.Codigo "
   sSQL = sSQL & "AND Cl.Grupo = " & Grupo_No & " "
   sSQL = sSQL & "GROUP BY Contrato_No & ' de: ' & Cliente "
   SelectDBCombo DCCliente, AdoCliente, sSQL, "Contratos", False
End Sub

Private Sub DGQuery_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  If CtrlDown And KeyCode = vbKeyT Then
     DGQuery.Col = 6
     Numero = Val(DGQuery.Text)
     MsgBox "Revise la Tabla del Credito No. " & DGQuery.Text & Chr(13) & Chr(13) & " en la siguiente Viñeta."
     sSQL = "SELECT T,Dia,Pago_No,Fecha,Capital,Interes,"
     sSQL = sSQL & "IVA,Abono,Saldo "
     sSQL = sSQL & "FROM Prestamos "
     sSQL = sSQL & "WHERE Contrato_No = " & Numero & " "
     sSQL = sSQL & "AND Pago_No<>0 "
     sSQL = sSQL & "ORDER BY Pago_No,Fecha "
     SelectDataGrid DGTablaPrestamo, AdoTablaPrestamo, sSQL
  End If
End Sub

Private Sub Form_Activate()
   sSQL = "SELECT Grupo "
   sSQL = sSQL & "FROM Clientes "
   sSQL = sSQL & "GROUP BY Grupo "
   SelectDBCombo DCGrupo, AdoGrupo, sSQL, "Grupo"
   Grupo_No = Val(DCGrupo.Text)
  
   sSQL = "SELECT Contrato_No & ' de: ' & Cliente As Contratos "
   sSQL = sSQL & "FROM Prestamos As CM,Clientes As Cl "
   sSQL = sSQL & "WHERE CM.Codigo_C = Cl.Codigo "
   sSQL = sSQL & "AND Cl.Grupo = " & Grupo_No & " "
   sSQL = sSQL & "GROUP BY Contrato_No & ' de: ' & Cliente "
   SelectDBCombo DCCliente, AdoCliente, sSQL, "Contratos", False
   Command1.Caption = "P&restamos"
   RatonNormal
End Sub

Private Sub Form_Load()
  'FACTURAS.MDB
   CentrarForm ListarContrato
   ConectarAdodc AdoGrupo
   ConectarAdodc AdoQuery
   ConectarAdodc AdoCliente
   ConectarAdodc AdoTablaPrestamo
End Sub

Private Sub MBoxFechaF_LostFocus()
  FechaValida MBoxFechaF, False
End Sub

Private Sub MBoxFechaI_LostFocus()
  FechaValida MBoxFechaI, False
End Sub

Private Sub OpcA_Click()
   Grupo_No = Val(DCGrupo.Text)
   sSQL = "SELECT Contrato_No & ' de: ' & Cliente As Contratos "
   sSQL = sSQL & "FROM Contratos_Meses As CM,Clientes As Cl "
   sSQL = sSQL & "WHERE CM.Codigo_C = Cl.Codigo "
   sSQL = sSQL & "AND Cl.Grupo = " & Grupo_No & " "
   sSQL = sSQL & "GROUP BY Contrato_No & ' de: ' & Cliente "
   SelectDBCombo DCCliente, AdoCliente, sSQL, "Contratos", False
   Command1.Caption = "C&ontratos"
End Sub

Private Sub OpcC_Click()
   Grupo_No = Val(DCGrupo.Text)
   sSQL = "SELECT Contrato_No & ' de: ' & Cliente As Contratos "
   sSQL = sSQL & "FROM Contratos_Meses As CM,Clientes As Cl "
   sSQL = sSQL & "WHERE CM.Codigo_C = Cl.Codigo "
   sSQL = sSQL & "AND Cl.Grupo = " & Grupo_No & " "
   sSQL = sSQL & "GROUP BY Contrato_No & ' de: ' & Cliente "
   SelectDBCombo DCCliente, AdoCliente, sSQL, "Contratos", False
   Command1.Caption = "C&ontratos"
End Sub

Private Sub OpcP_Click()
   Grupo_No = Val(DCGrupo.Text)
   sSQL = "SELECT Contrato_No & ' de: ' & Cliente As Contratos "
   sSQL = sSQL & "FROM Prestamos As CM,Clientes As Cl "
   sSQL = sSQL & "WHERE CM.Codigo_C = Cl.Codigo "
   sSQL = sSQL & "AND Cl.Grupo = " & Grupo_No & " "
   sSQL = sSQL & "GROUP BY Contrato_No & ' de: ' & Cliente "
   SelectDBCombo DCCliente, AdoCliente, sSQL, "Contratos", False
   Command1.Caption = "P&restamos"
End Sub

