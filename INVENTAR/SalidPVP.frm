VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Begin VB.Form SalidaPVP 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SALIDAS EN COSTOS/PVP"
   ClientHeight    =   6690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12330
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   12330
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command5 
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
      Height          =   960
      Left            =   11130
      Picture         =   "SalidPVP.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1155
      Width           =   1065
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
      Height          =   960
      Left            =   11130
      Picture         =   "SalidPVP.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2205
      Width           =   1065
   End
   Begin VB.CommandButton Command3 
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
      Height          =   960
      Left            =   11130
      Picture         =   "SalidPVP.frx":1194
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   105
      Width           =   1065
   End
   Begin MSDataGridLib.DataGrid DGQuery 
      Bindings        =   "SalidPVP.frx":15D6
      Height          =   5160
      Left            =   105
      TabIndex        =   12
      Top             =   1050
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   9102
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
   Begin MSAdodcLib.Adodc AdoProd 
      Height          =   330
      Left            =   240
      Top             =   5280
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
      Caption         =   "Prod"
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
   Begin MSAdodcLib.Adodc AdoDetKardex 
      Height          =   330
      Left            =   105
      Top             =   6300
      Width           =   2640
      _ExtentX        =   4657
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
      Caption         =   "DetKardex"
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
   Begin MSAdodcLib.Adodc AdoOrden 
      Height          =   330
      Left            =   240
      Top             =   4560
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
      Caption         =   "Orden"
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
   Begin MSAdodcLib.Adodc AdoProducto 
      Height          =   330
      Left            =   240
      Top             =   4920
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
      Caption         =   "Producto"
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
   Begin VB.Frame Frame1 
      Height          =   1065
      Left            =   105
      TabIndex        =   0
      Top             =   0
      Width           =   10935
      Begin MSDataListLib.DataCombo DCProducto 
         Bindings        =   "SalidPVP.frx":15F1
         DataSource      =   "AdoProducto"
         Height          =   315
         Left            =   2730
         TabIndex        =   6
         Top             =   630
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo DCOrden 
         Bindings        =   "SalidPVP.frx":160B
         DataSource      =   "AdoOrden"
         Height          =   315
         Left            =   8400
         TabIndex        =   9
         Top             =   630
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.OptionButton OpcCD 
         Caption         =   "&Diario No."
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
         Left            =   9555
         TabIndex        =   8
         Top             =   210
         Width           =   1275
      End
      Begin VB.OptionButton OpcO 
         Caption         =   "&Guía No."
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
         Left            =   8400
         TabIndex        =   7
         Top             =   210
         Width           =   1170
      End
      Begin VB.OptionButton OpcR 
         Caption         =   "&Responsable"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2730
         TabIndex        =   5
         Top             =   210
         Value           =   -1  'True
         Width           =   1485
      End
      Begin MSMask.MaskEdBox MBoxFechaI 
         Height          =   330
         Left            =   1365
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
      Begin MSMask.MaskEdBox MBoxFechaF 
         Height          =   330
         Left            =   1365
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
      Begin VB.Label Label5 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Fecha I&nicial"
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
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Fecha &Final"
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
         Width           =   1275
      End
   End
   Begin MSAdodcLib.Adodc AdoAux 
      Height          =   330
      Left            =   210
      Top             =   4200
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
   Begin VB.Label LabelUtilidad 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
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
      Left            =   8505
      TabIndex        =   16
      Top             =   6300
      Width           =   1275
   End
   Begin VB.Label LabelPVP 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
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
      TabIndex        =   19
      Top             =   6300
      Width           =   1275
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " PVP"
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
      Left            =   4830
      TabIndex        =   18
      Top             =   6300
      Width           =   855
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Utilidad/Perdida"
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
      Left            =   6930
      TabIndex        =   17
      Top             =   6300
      Width           =   1590
   End
   Begin VB.Label LabelSalida 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
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
      TabIndex        =   14
      Top             =   6300
      Width           =   1275
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Salidas"
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
      Left            =   2730
      TabIndex        =   15
      Top             =   6300
      Width           =   855
   End
End
Attribute VB_Name = "SalidaPVP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ListarDetallesInv()
  RatonReloj
  FechaValida MBoxFechaI
  FechaValida MBoxFechaF
  If OpcR.value Then
     Codigo = SinEspaciosIzq(DCProducto.Text)
     TextoValidoVar Codigo
  ElseIf OpcCD.value Then
     Codigo = DCOrden.Text
     TextoValidoVar Codigo, True
  Else
     Codigo = DCOrden.Text
     TextoValidoVar Codigo
  End If
  FechaIni = BuscarFecha(MBoxFechaI)
  FechaFin = BuscarFecha(MBoxFechaF)

  sSQL = "SELECT K.Codigo_Barra,K.Codigo_Inv,P.Producto,K.Fecha,K.TP,K.Numero,K.Salida,K.Valor_Total As Total_Costo,(K.Salida * P.PVP) As Total_PVP,K.Orden_No " _
       & "FROM Trans_Kardex As K,Catalogo_Productos As P " _
       & "WHERE K.Item = '" & NumEmpresa & "' " _
       & "AND K.Periodo = '" & Periodo_Contable & "' "
  If OpcR.value Then
     sSQL = sSQL & "AND K.Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
          & "AND Codigo_P = '" & Codigo & "' "
  ElseIf OpcCD.value Then
     sSQL = sSQL & "AND K.TP = 'CD' " _
          & "AND K.Numero = " & Val(Codigo) & " "
  Else
     sSQL = sSQL & "AND Orden_No = '" & Codigo & "' "
  End If
  sSQL = sSQL & "AND K.Codigo_Inv = P.Codigo_Inv " _
       & "AND K.Salida > 0 " _
       & "AND K.Item = P.Item " _
       & "AND K.Periodo = P.Periodo " _
       & "ORDER BY K.Orden_No,K.Codigo_Inv,K.Fecha,K.TP,K.Numero,K.ID "
  Select_Adodc_Grid DGQuery, AdoDetKardex, sSQL
  Total = 0: Saldo = 0
  With AdoDetKardex.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          Total = Total + .Fields("Total_Costo")
          Saldo = Saldo + .Fields("Total_PVP")
         .MoveNext
       Loop
      .MoveFirst
   End If
  End With
  LabelSalida.Caption = Format(Total, "#,##0.00")
  LabelPVP.Caption = Format(Saldo, "#,##0.00")
  LabelUtilidad.Caption = Format(Saldo - Total, "#,##0.00")
  RatonNormal
End Sub

Private Sub Command2_Click()
 Unload SalidaPVP
End Sub

Private Sub Command3_Click()
  ListarDetallesInv
End Sub

Private Sub Command5_Click()
  DGQuery.Visible = False
  MensajeEncabData = "CONTROL DE COSTO Y P.V.P. "
  If OpcR.value Then
     SQLMsg1 = "Beneficiario: " & DCProducto.Text
  ElseIf OpcCD.value Then
     SQLMsg1 = "Diario No. " & DCOrden.Text
  Else
     SQLMsg1 = "Orden No. " & DCOrden.Text
  End If
  ImprimirCostoInventario AdoDetKardex, 1, 8
  DGQuery.Visible = True
End Sub

Private Sub DGQuery_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF1 Then GenerarDataTexto KardexSQLs, AdoDetKardex
End Sub

Private Sub Form_Activate()
  sSQL = "SELECT Codigo_Inv,Fecha,Codigo_P,SUM(Entrada) As Compras " _
       & "FROM Trans_Kardex  " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND Entrada > 0 " _
       & "GROUP BY Codigo_Inv,Fecha,Codigo_P " _
       & "ORDER BY Codigo_Inv,Fecha,Codigo_P "
  Select_Adodc AdoAux, sSQL
  RatonReloj
  With AdoAux.Recordset
   If .RecordCount > 0 Then
       CodigoInv = .Fields("Codigo_Inv")
       FechaFinal = .Fields("Fecha")
       CodigoP = .Fields("Codigo_P")
       Do While Not .EOF
          If CodigoInv <> .Fields("Codigo_Inv") Then
             Actualizar_Ventas_Proveedor
             CodigoInv = .Fields("Codigo_Inv")
             FechaFinal = .Fields("Fecha")
             CodigoP = .Fields("Codigo_P")
          End If
          If FechaFinal <> .Fields("Fecha") Then
             FechaFinal = .Fields("Fecha")
             CodigoP = .Fields("Codigo_P")
             Actualizar_Ventas_Proveedor
          End If
          FechaFinal = .Fields("Fecha")
         .MoveNext
       Loop
       Actualizar_Ventas_Proveedor
   End If
  End With

  sSQL = "SELECT TK.Codigo_P & Space(5) & C.Cliente As Nombre_Cta " _
       & "FROM Trans_Kardex As TK,Clientes As C " _
       & "WHERE TK.Item = '" & NumEmpresa & "' " _
       & "AND TK.Periodo = '" & Periodo_Contable & "' " _
       & "AND C.Codigo = TK.Codigo_P " _
       & "GROUP BY TK.Codigo_P,C.Cliente "
  SelectDB_Combo DCProducto, AdoProducto, sSQL, "Nombre_Cta"
  RatonNormal
  MBoxFechaI.SetFocus
End Sub

Private Sub Form_Load()
  CentrarForm SalidaPVP
  ConectarAdodc AdoAux
  ConectarAdodc AdoProd
  ConectarAdodc AdoOrden
  ConectarAdodc AdoProducto
  ConectarAdodc AdoDetKardex
End Sub

Private Sub MBoxFechaF_GotFocus()
  MarcarTexto MBoxFechaF
End Sub

Private Sub MBoxFechaF_LostFocus()
  FechaValida MBoxFechaF
End Sub

Private Sub MBoxFechaI_GotFocus()
  MarcarTexto MBoxFechaI
End Sub

Private Sub MBoxFechaI_LostFocus()
  FechaValida MBoxFechaI
End Sub

Private Sub OpcCD_Click()
  sSQL = "SELECT Numero " _
       & "FROM Trans_Kardex " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "GROUP BY Numero "
  SelectDB_Combo DCOrden, AdoOrden, sSQL, "Numero"
  DCOrden.SetFocus
End Sub

Private Sub OpcO_Click()
  sSQL = "SELECT Orden_No " _
       & "FROM Trans_Kardex " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "GROUP BY Orden_No "
  SelectDB_Combo DCOrden, AdoOrden, sSQL, "Orden_No"
  DCOrden.SetFocus
End Sub

Private Sub OpcR_Click()
  DCProducto.SetFocus
End Sub

Public Sub Actualizar_Ventas_Proveedor()
  sSQL = "UPDATE Trans_Kardex " _
       & "SET Codigo_P = '" & CodigoP & "' " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND Codigo_Inv = '" & CodigoInv & "' " _
       & "AND Salida > 0 " _
       & "AND Fecha >= #" & BuscarFecha(FechaFinal) & "# "
  Ejecutar_SQL_SP sSQL
End Sub
