VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form TotalKardex 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "INVENTARIO DE KARDEX"
   ClientHeight    =   7365
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7365
   ScaleWidth      =   10935
   ShowInTaskbar   =   0   'False
   Begin MSDataListLib.DataCombo DCTInv 
      Bindings        =   "TotalKar.frx":0000
      DataSource      =   "AdoTInv"
      Height          =   315
      Left            =   2520
      TabIndex        =   4
      Top             =   525
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16777215
      Text            =   "DC"
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
      Left            =   9870
      Picture         =   "TotalKar.frx":0016
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   105
      Width           =   960
   End
   Begin VB.CommandButton Command1 
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
      Picture         =   "TotalKar.frx":08E0
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   105
      Width           =   960
   End
   Begin VB.OptionButton OpcE 
      Caption         =   "Ventas"
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
      TabIndex        =   6
      Top             =   525
      Width           =   1065
   End
   Begin VB.OptionButton OpcI 
      Caption         =   "Compras"
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
      TabIndex        =   5
      Top             =   105
      Value           =   -1  'True
      Width           =   1065
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Costos"
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
      Left            =   7980
      Picture         =   "TotalKar.frx":11AA
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   105
      Width           =   960
   End
   Begin VB.CheckBox CheqGrupo 
      Caption         =   "Por Grupo de Inventario"
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
      Top             =   525
      Width           =   2430
   End
   Begin MSDataGridLib.DataGrid DGQuery 
      Bindings        =   "TotalKar.frx":15EC
      Height          =   6000
      Left            =   105
      TabIndex        =   10
      Top             =   945
      Width           =   10725
      _ExtentX        =   18918
      _ExtentY        =   10583
      _Version        =   393216
      AllowUpdate     =   0   'False
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
   Begin MSAdodcLib.Adodc AdoDetKardex 
      Height          =   330
      Left            =   105
      Top             =   6930
      Width           =   7575
      _ExtentX        =   13361
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
   Begin MSAdodcLib.Adodc AdoProducto 
      Height          =   330
      Left            =   420
      Top             =   2835
      Visible         =   0   'False
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
   Begin MSAdodcLib.Adodc AdoProd 
      Height          =   330
      Left            =   420
      Top             =   3150
      Visible         =   0   'False
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
   Begin MSMask.MaskEdBox MBoxFechaF 
      Height          =   330
      Left            =   3885
      TabIndex        =   2
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
   Begin MSAdodcLib.Adodc AdoTInv 
      Height          =   330
      Left            =   420
      Top             =   2520
      Visible         =   0   'False
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
      Caption         =   "TInv"
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
   Begin MSMask.MaskEdBox MBoxFechaI 
      Height          =   330
      Left            =   2520
      TabIndex        =   1
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
   Begin MSAdodcLib.Adodc AdoAux 
      Height          =   330
      Left            =   420
      Top             =   3465
      Visible         =   0   'False
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
   Begin VB.Label LabelTot 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.0"
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
      TabIndex        =   11
      Top             =   6930
      Width           =   1800
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Valor Total"
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
      TabIndex        =   12
      Top             =   6930
      Width           =   1170
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fecha de Reporte"
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
      Width           =   2430
   End
End
Attribute VB_Name = "TotalKardex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
  If OpcI.value Then MensajeEncabData = "RESUMEN COSTOS DE COMPRAS"
  If OpcE.value Then MensajeEncabData = "RESUMEN COSTOS DE VENTAS"
  If CheqGrupo.value Then
     SQLMsg2 = DCTInv.Text
  Else
     SQLMsg2 = ""
  End If
  SQLMsg1 = "RESUMEN DEL " & MBoxFechaI.Text & " AL " & MBoxFechaF.Text
  ImprimirAdoCostos AdoDetKardex, True, 1, 10, Total
End Sub

Private Sub Command2_Click()
  Unload TotalKardex
End Sub

Private Sub Command4_Click()
  RatonReloj
  sSQL = "DELETE * " _
       & "FROM Saldo_Diarios " _
       & "WHERE CodigoU = '" & CodigoUsuario & "' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND TP = 'INVX' "
  ConectarAdoExecute sSQL
  DGQuery.Visible = False
  FechaValida MBoxFechaF
  Codigo = SinEspaciosIzq(DCTInv.Text)
  FechaValida MBoxFechaI
  FechaValida MBoxFechaF
  FechaIni = BuscarFecha(MBoxFechaI)
  FechaFin = BuscarFecha(MBoxFechaF)
  Codigo = SinEspaciosIzq(DCTInv)
  If OpcI.value Then
     sSQL = "SELECT CP.Codigo_Inv,CP.Producto,SUM(DF.Entrada) As Cant " _
          & "FROM Catalogo_Productos As CP, Trans_Kardex AS DF " _
          & "WHERE DF.Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
          & "AND DF.Item = '" & NumEmpresa & "' " _
          & "AND DF.Periodo = '" & Periodo_Contable & "' " _
          & "AND CP.Item = DF.Item " _
          & "AND CP.Periodo = DF.Periodo " _
          & "AND CP.Codigo_Inv = DF.Codigo_Inv "
     If CheqGrupo.value = 1 Then sSQL = sSQL & "AND Mid$(CP.Codigo_Inv,1," & Len(Codigo) & ") = '" & Codigo & "' "
     sSQL = sSQL & "GROUP BY CP.Codigo_Inv,CP.Producto " _
          & "HAVING SUM(DF.Entrada) > 0 " _
          & "ORDER BY CP.Codigo_Inv "
  Else
     sSQL = "SELECT CP.Codigo_Inv,CP.Producto,SUM(DF.Cantidad) As Cant " _
          & "FROM Catalogo_Productos As CP, Detalle_Factura AS DF " _
          & "WHERE DF.Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
          & "AND DF.Item = '" & NumEmpresa & "' " _
          & "AND DF.Periodo = '" & Periodo_Contable & "' " _
          & "AND CP.Item = DF.Item " _
          & "AND CP.Periodo = DF.Periodo " _
          & "AND CP.Codigo_Inv = DF.Codigo "
     If CheqGrupo.value = 1 Then sSQL = sSQL & "AND Mid$(CP.Codigo_Inv,1," & Len(Codigo) & ") = '" & Codigo & "' "
     sSQL = sSQL & "GROUP BY CP.Codigo_Inv,CP.Producto " _
          & "HAVING SUM(DF.Cantidad) > 0 " _
          & "ORDER BY CP.Codigo_Inv "
  End If
  'MsgBox sSQL
  SelectDataGrid DGQuery, AdoDetKardex, sSQL
  Total = 0: Contador = 0
  With AdoDetKardex.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          Contador = Contador + 1
          TotalKardex.Caption = Format(Contador / .RecordCount, "00%")
          Producto = .Fields("Producto")
          CodigoInv = .Fields("Codigo_Inv")
          Cantidad = .Fields("Cant")
          Precio = 0
          sSQL = "SELECT TOP 1 * " _
               & "FROM Trans_Kardex " _
               & "WHERE Codigo_Inv = '" & CodigoInv & "' " _
               & "AND Fecha <= #" & FechaIni & "# " _
               & "AND Item = '" & NumEmpresa & "' " _
               & "AND Periodo = '" & Periodo_Contable & "' " _
               & "AND Entrada > 0 " _
               & "ORDER BY Fecha DESC,Kardex DESC "
          SelectAdodc AdoAux, sSQL
          If AdoAux.Recordset.RecordCount > 0 Then Precio = AdoAux.Recordset.Fields("Valor_Unitario")
          If Precio <= 0 Then
             sSQL = "SELECT TOP 1 * " _
                  & "FROM Trans_Kardex " _
                  & "WHERE Codigo_Inv = '" & CodigoInv & "' " _
                  & "AND Fecha <= #" & FechaFin & "# " _
                  & "AND Item = '" & NumEmpresa & "' " _
                  & "AND Periodo = '" & Periodo_Contable & "' " _
                  & "AND Entrada > 0 " _
                  & "ORDER BY Fecha DESC,Kardex DESC "
             SelectAdodc AdoAux, sSQL
             If AdoAux.Recordset.RecordCount > 0 Then Precio = AdoAux.Recordset.Fields("Valor_Unitario")
          End If
          If Precio > 0 Then
             Total = Precio * Cantidad
             SetAdoAddNew "Saldo_Diarios"
             SetAdoFields "T", Normal
             SetAdoFields "Cta", CodigoInv
             SetAdoFields "Comprobante", Producto
             SetAdoFields "TP", "INVX"
             SetAdoFields "Fecha", MBoxFechaF
             SetAdoFields "Saldo_Actual", Cantidad
             SetAdoFields "Hoy", Precio
             SetAdoFields "Total", Total
             SetAdoFields "Item", NumEmpresa
             SetAdoFields "CodigoU", CodigoUsuario
             SetAdoUpdate
          End If
         .MoveNext
       Loop
   End If
  End With
  sSQL = "SELECT Cta As Codigo,Comprobante As Producto,Saldo_Actual As Cantidad,Hoy As Precio,"
  If OpcI.value Then
     sSQL = sSQL & "Total As Compras "
  Else
     sSQL = sSQL & "Total As Ventas "
  End If
  sSQL = sSQL & "FROM Saldo_Diarios " _
       & "WHERE Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' "
  SelectDataGrid DGQuery, AdoDetKardex, sSQL
  Total = 0
  With AdoDetKardex.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          If OpcI.value Then
             Total = Total + .Fields("Compras")
          Else
             Total = Total + .Fields("Ventas")
          End If
         .MoveNext
       Loop
   End If
  End With
  LabelTot.Caption = Format(Total, "#,##0.00")
  DGQuery.Visible = True
  TotalKardex.Caption = "RESUMEN DE COMPRAS/VENTAS PROMEDIADAS"
  RatonNormal
End Sub

Private Sub DGQuery_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF1 Then GenerarDataTexto TotalKardex, AdoDetKardex
End Sub

Private Sub Form_Activate()
  sSQL = "SELECT Codigo_Inv & '  ' & Producto As NomProd " _
       & "FROM Catalogo_Productos " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND TC = 'I' " _
       & "ORDER BY Codigo_Inv "
  SelectDBCombo DCTInv, AdoTInv, sSQL, "NomProd"
  DGQuery.Visible = False
  TotalKardex.Caption = "RESUMEN DE COMPRAS/VENTAS PROMEDIADAS"
  MBoxFechaI.SetFocus
  RatonNormal
End Sub

Private Sub Form_Load()
  CentrarForm TotalKardex
  ConectarAdodc AdoAux
  ConectarAdodc AdoProd
  ConectarAdodc AdoTInv
  ConectarAdodc AdoProducto
  ConectarAdodc AdoDetKardex
End Sub

Private Sub MBoxFechaF_GotFocus()
  MarcarTexto MBoxFechaF
End Sub

Private Sub MBoxFechaF_LostFocus()
  FechaValida MBoxFechaF, False
End Sub

Private Sub MBoxFechaI_GotFocus()
  MarcarTexto MBoxFechaI
End Sub

Private Sub MBoxFechaI_LostFocus()
  FechaValida MBoxFechaI, False
End Sub

