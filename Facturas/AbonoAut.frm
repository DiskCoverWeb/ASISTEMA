VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form AbonoAutomatico 
   Caption         =   "INGRESO DE CAJA"
   ClientHeight    =   8400
   ClientLeft      =   5040
   ClientTop       =   4395
   ClientWidth     =   13365
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8400
   ScaleWidth      =   13365
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command4 
      Caption         =   "&Todas las Pendientes"
      Height          =   750
      Left            =   8085
      Picture         =   "AbonoAut.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   105
      Width           =   1170
   End
   Begin VB.OptionButton OpcNC 
      Caption         =   "Notas de Crédito"
      Height          =   330
      Left            =   2835
      TabIndex        =   12
      Top             =   945
      Width           =   1800
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Por Autorización"
      Height          =   750
      Left            =   6825
      Picture         =   "AbonoAut.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   105
      Width           =   1275
   End
   Begin VB.OptionButton OpcCierre 
      Caption         =   "Cierre Periodo"
      Height          =   330
      Left            =   1155
      TabIndex        =   11
      Top             =   945
      Width           =   1590
   End
   Begin VB.OptionButton OpcAbonos 
      Caption         =   "Abonos"
      Height          =   330
      Left            =   105
      TabIndex        =   10
      Top             =   945
      Value           =   -1  'True
      Width           =   960
   End
   Begin MSDataGridLib.DataGrid DGFactura 
      Bindings        =   "AbonoAut.frx":0884
      Height          =   5790
      Left            =   105
      TabIndex        =   19
      Top             =   1365
      Width           =   13140
      _ExtentX        =   23178
      _ExtentY        =   10213
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
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
            LCID            =   3082
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
            LCID            =   3082
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
   Begin MSAdodcLib.Adodc AdoIngCaja 
      Height          =   330
      Left            =   315
      Top             =   2835
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
      Caption         =   "IngCaja"
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
   Begin MSAdodcLib.Adodc AdoFactura 
      Height          =   330
      Left            =   105
      Top             =   7140
      Width           =   6630
      _ExtentX        =   11695
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
      Caption         =   "Factura"
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
   Begin VB.CommandButton Command2 
      Caption         =   "&Salir"
      Height          =   750
      Left            =   10395
      Picture         =   "AbonoAut.frx":089D
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   105
      Width           =   960
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   750
      Left            =   9345
      Picture         =   "AbonoAut.frx":1167
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   105
      Width           =   960
   End
   Begin MSMask.MaskEdBox MBFecha 
      Height          =   330
      Left            =   4095
      TabIndex        =   7
      Top             =   420
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
   Begin MSAdodcLib.Adodc AdoBanco 
      Height          =   330
      Left            =   315
      Top             =   3150
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
      Caption         =   "Banco"
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
      Bindings        =   "AbonoAut.frx":15A9
      DataSource      =   "AdoBanco"
      Height          =   315
      Left            =   4725
      TabIndex        =   13
      Top             =   945
      Width           =   6630
      _ExtentX        =   11695
      _ExtentY        =   556
      _Version        =   393216
      Text            =   "Banco"
   End
   Begin MSMask.MaskEdBox MBFechaA 
      Height          =   330
      Left            =   5460
      TabIndex        =   9
      Top             =   420
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
   Begin MSAdodcLib.Adodc AdoSerie 
      Height          =   330
      Left            =   315
      Top             =   3465
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
   Begin MSDataListLib.DataCombo DCTipo 
      Bindings        =   "AbonoAut.frx":15C0
      DataSource      =   "AdoTipo"
      Height          =   360
      Left            =   3150
      TabIndex        =   5
      Top             =   420
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   635
      _Version        =   393216
      Text            =   "FA"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo DCSerie 
      Bindings        =   "AbonoAut.frx":15D6
      DataSource      =   "AdoSerie"
      Height          =   360
      Left            =   1890
      TabIndex        =   3
      Top             =   420
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   635
      _Version        =   393216
      Text            =   "001001"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo DCAutorizacion 
      Bindings        =   "AbonoAut.frx":15ED
      DataSource      =   "AdoAutorizacion"
      Height          =   360
      Left            =   105
      TabIndex        =   1
      Top             =   420
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   635
      _Version        =   393216
      Text            =   "1234567890"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSAdodcLib.Adodc AdoAutorizacion 
      Height          =   330
      Left            =   315
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
      Caption         =   "Autorizacion"
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
   Begin MSAdodcLib.Adodc AdoTipo 
      Height          =   330
      Left            =   315
      Top             =   3780
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
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
      Caption         =   "Tipo"
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
   Begin VB.Label Label11 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Autorización"
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   1695
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Serie"
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   1890
      TabIndex        =   2
      Top             =   105
      Width           =   1170
   End
   Begin VB.Label Label9 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Tipo"
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   3150
      TabIndex        =   4
      Top             =   105
      Width           =   855
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &ASIENTO:"
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   5460
      TabIndex        =   8
      Top             =   105
      Width           =   1275
   End
   Begin VB.Label LabelSaldo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Left            =   8295
      TabIndex        =   17
      Top             =   7140
      Width           =   1800
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Saldo Pendiente"
      Height          =   330
      Left            =   6720
      TabIndex        =   18
      Top             =   7140
      Width           =   1590
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &CIERRE:"
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   4095
      TabIndex        =   6
      Top             =   105
      Width           =   1275
   End
End
Attribute VB_Name = "AbonoAutomatico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
  FechaValida MBFecha
  FechaValida MBFechaA
  TA.Fecha = MBFechaA
  FA.Fecha_Corte = MBFechaA
  FA.Fecha_Desde = MBFechaA
  FA.Fecha_Hasta = MBFechaA
  Mensajes = "Esta Seguro que desea grabar Abonos"
  Titulo = "Formulario de Grabación."
  If BoxMensaje = vbYes Then
     FechaTexto = MBFechaA ' FechaSistema
     DGFactura.Visible = False
     Total = 0: Contador = 0
     TA.Cta = SinEspaciosIzq(DCBanco.Text)
     With AdoFactura.Recordset
      If .RecordCount > 0 Then
         .MoveFirst
          FA.TC = .Fields("TC")
          FA.Serie = .Fields("Serie")
          CodigoP = Ninguno
          TipoDoc = Ninguno
          Do While Not .EOF
            'Abono de Factura
             TA.T = Cancelado
             TA.TP = .Fields("TC")
             TA.Serie = .Fields("Serie")
             TA.Factura = .Fields("Factura")
             TA.Autorizacion = .Fields("Autorizacion")
             TA.Cta_CxP = .Fields("Cta_CxP")
             TA.CodigoC = .Fields("CodigoC")
             TA.Abono = .Fields("Saldo_MN")
             TA.Cheque = .Fields("Grupo")
             If OpcAbonos.value Then
                TA.Banco = "Abonos de Cierre"
             ElseIf OpcNC.value Then
                TA.Banco = "NOTA DE CREDITO"
                TA.Cheque = "VENTAS"
             Else
                TA.Banco = "Cierre de Periodo"
             End If
             Grabar_Abonos TA
             sSQL = "UPDATE Facturas " _
                  & "SET Saldo_MN = 0, T = 'C' " _
                  & "WHERE Fecha <= #" & BuscarFecha(MBFecha) & "# " _
                  & "AND Item = '" & NumEmpresa & "' " _
                  & "AND Periodo = '" & Periodo_Contable & "' " _
                  & "AND TC = '" & TA.TP & "' " _
                  & "AND Serie = '" & TA.Serie & "' " _
                  & "AND Autorizacion = '" & TA.Autorizacion & "' " _
                  & "AND T <> 'A' " _
                  & "AND Saldo_MN > 0 "
             Ejecutar_SQL_SP sSQL
             Contador = Contador + 1
             AbonoAutomatico.Caption = "Factura No. " & TA.Serie & "-" & TA.Factura & " " & Format$(Contador / .RecordCount, "00%")
            .MoveNext
          Loop
      End If
     End With
     'Procesar_Saldo_De_Facturas AbonoAutomatico, AdoFactura
     Actualizar_Abonos_Facturas_SP FA
     DGFactura.Visible = True
     RatonNormal
  End If
  Unload Me
End Sub

Private Sub Command2_Click()
   Control_Procesos Normal, "Salir de abonos de facturas"
   Unload Me
End Sub

Private Sub Command3_Click()
  Saldos_Pendiente_Facturas True
End Sub

Private Sub Command4_Click()
  Saldos_Pendiente_Facturas False
End Sub

Private Sub DCAutorizacion_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCAutorizacion_LostFocus()
  FA.Autorizacion = DCAutorizacion
  sSQL = "SELECT Serie " _
       & "FROM Facturas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND Autorizacion = '" & FA.Autorizacion & "' " _
       & "GROUP BY Serie " _
       & "ORDER BY Serie "
  SelectDB_Combo DCSerie, AdoSerie, sSQL, "Serie"
End Sub

Private Sub DCSerie_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCSerie_LostFocus()
  FA.Serie = DCSerie
  sSQL = "SELECT TC " _
       & "FROM Facturas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND Autorizacion = '" & FA.Autorizacion & "' " _
       & "AND Serie = '" & FA.Serie & "' " _
       & "GROUP BY TC " _
       & "ORDER BY TC DESC "
  SelectDB_Combo DCTipo, AdoTipo, sSQL, "TC"
End Sub

Private Sub DCTipo_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCTipo_LostFocus()
  FA.TC = DCTipo
  If TipoFactura = "" Then TipoFactura = "FA"
End Sub

Private Sub Form_Activate()
  DGFactura.Visible = False
  MBFecha.Text = FechaSistema
  sSQL = "SELECT Codigo & Space(5) & Cuenta As NomCuenta " _
       & "FROM Catalogo_Cuentas " _
       & "WHERE MidStrg(Codigo,1,1) = '1' " _
       & "AND TC IN ('CJ','BA') " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "ORDER BY Codigo "
  SelectDB_Combo DCBanco, AdoBanco, sSQL, "NomCuenta"
   
  sSQL = "SELECT Autorizacion " _
       & "FROM Facturas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND LEN(Autorizacion) > 8 " _
       & "AND T = 'P' " _
       & "GROUP BY Autorizacion " _
       & "ORDER BY Autorizacion "
  SelectDB_Combo DCAutorizacion, AdoAutorizacion, sSQL, "Autorizacion"
  FA.TC = Ninguno
  FA.Serie = Ninguno
  FA.Autorizacion = Ninguno
  RatonNormal
  Saldos_Pendiente_Facturas True
End Sub

Private Sub Form_Load()
   'CentrarForm AbonoAutomatico
   ConectarAdodc AdoTipo
   ConectarAdodc AdoBanco
   ConectarAdodc AdoSerie
   ConectarAdodc AdoFactura
   ConectarAdodc AdoAutorizacion
   ConectarAdodc AdoIngCaja
End Sub

Private Sub MBFecha_GotFocus()
  MarcarTexto MBFecha
End Sub

Private Sub MBFecha_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBFecha_LostFocus()
  FechaValida MBFecha
End Sub

Private Sub MBFechaA_GotFocus()
  MarcarTexto MBFechaA
End Sub

Private Sub MBFechaA_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBFechaA_LostFocus()
  FechaValida MBFechaA
End Sub

Private Sub OpcAbonos_Click()
  DCBanco.Visible = True
  sSQL = "SELECT Codigo & Space(5) & Cuenta As NomCuenta " _
       & "FROM Catalogo_Cuentas " _
       & "WHERE MidStrg(Codigo,1,1) = '1' " _
       & "AND TC IN ('CJ','BA') " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "ORDER BY Codigo "
  SelectDB_Combo DCBanco, AdoBanco, sSQL, "NomCuenta"
End Sub

Private Sub OpcCierre_Click()
  DCBanco.Visible = True
  sSQL = "SELECT Codigo & Space(5) & Cuenta As NomCuenta " _
       & "FROM Catalogo_Cuentas " _
       & "WHERE MidStrg(Codigo,1,1) = '4' " _
       & "AND DG = 'D' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "ORDER BY Codigo "
  SelectDB_Combo DCBanco, AdoBanco, sSQL, "NomCuenta"
End Sub

Public Sub Saldos_Pendiente_Facturas(PorAutorizacion As Boolean)
  DGFactura.Visible = False
  RatonReloj
  FechaValida MBFecha
  FechaFin = BuscarFecha(MBFecha)
  Cadena = ""
  sSQL = "SELECT TC,Autorizacion,Serie,MAX(Vencimiento) As Fecha_Venc " _
       & "FROM Facturas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND LEN(Autorizacion) <= 13 " _
       & "GROUP BY TC,Autorizacion,Serie " _
       & "ORDER BY TC,Autorizacion,Serie "
  Select_Adodc AdoFactura, sSQL
  Total = 0: Contador = 0
  With AdoFactura.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          Mifecha = BuscarFecha(.Fields("Fecha_Venc"))
          sSQL = "UPDATE Facturas " _
               & "SET Vencimiento = #" & Mifecha & "# " _
               & "WHERE Item = '" & NumEmpresa & "' " _
               & "AND Periodo = '" & Periodo_Contable & "' " _
               & "AND TC = '" & .Fields("TC") & "' " _
               & "AND Serie = '" & .Fields("Serie") & "' " _
               & "AND Autorizacion = '" & .Fields("Autorizacion") & "' " _
               & "AND Serie = '" & .Fields("Serie") & "' " _
               & "AND Vencimiento <> #" & Mifecha & "# "
          Ejecutar_SQL_SP sSQL
         .MoveNext
       Loop
   End If
  End With
  sSQL = "SELECT F.T,F.TC,F.Serie,F.Autorizacion,F.Factura,F.CodigoC,C.Grupo,F.Cta_CxP,F.Saldo_MN " _
       & "FROM Facturas As F,Clientes As C " _
       & "WHERE F.Fecha <= #" & FechaFin & "# " _
       & "AND F.Item = '" & NumEmpresa & "' " _
       & "AND F.Periodo = '" & Periodo_Contable & "' "
  If PorAutorizacion Then
     sSQL = sSQL _
          & "AND F.TC = '" & FA.TC & "' " _
          & "AND F.Serie = '" & FA.Serie & "' " _
          & "AND F.Autorizacion = '" & FA.Autorizacion & "' "
  End If
  sSQL = sSQL _
       & "AND F.Saldo_MN > 0 " _
       & "AND F.T <> '" & Anulado & "' " _
       & "AND F.CodigoC = C.Codigo " _
       & "ORDER BY F.Autorizacion,F.Serie,F.TC,F.Factura,F.CodigoC,C.Grupo,F.Cta_CxP "
  Select_Adodc AdoFactura, sSQL
  Total = 0: Contador = 0
  With AdoFactura.Recordset
   If .RecordCount > 0 Then
       Factura_Desde = .Fields("Factura")
       Do While Not .EOF
          Contador = Contador + 1
          AbonoAutomatico.Caption = Format$(Contador / .RecordCount, "00%")
          Total = Total + .Fields("Saldo_MN")
          Factura_Hasta = .Fields("Factura")
         .MoveNext
       Loop
   End If
  End With
  LabelSaldo.Caption = Format$(Total, "#,##0.00")
  DGFactura.Visible = True
  RatonNormal
End Sub

Private Sub OpcNC_Click()
  DCBanco.Visible = True
  sSQL = "SELECT Codigo & Space(5) & Cuenta As NomCuenta " _
       & "FROM Catalogo_Cuentas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND DG = 'D' " _
       & "ORDER BY Codigo "
  SelectDB_Combo DCBanco, AdoBanco, sSQL, "NomCuenta"
End Sub
