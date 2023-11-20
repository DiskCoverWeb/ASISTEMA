VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form BalanceConsolidado 
   Caption         =   "BALANCE GENERAL"
   ClientHeight    =   7110
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11625
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7110
   ScaleWidth      =   11625
   WindowState     =   2  'Maximized
   Begin MSDataGridLib.DataGrid DataGridConsolidado 
      Bindings        =   "BalanceC.frx":0000
      Height          =   5895
      Left            =   105
      TabIndex        =   11
      Top             =   840
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   10398
      _Version        =   393216
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
   Begin MSAdodcLib.Adodc AdodcConsolidado 
      Height          =   330
      Left            =   105
      Top             =   6720
      Width           =   10095
      _ExtentX        =   17806
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
      Caption         =   "Consolidado"
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
      Left            =   8820
      TabIndex        =   4
      Top             =   420
      Width           =   1380
      _ExtentX        =   2434
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
   Begin VB.Data DataPromedio 
      Caption         =   "Promedio"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   330
      Left            =   210
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1155
      Visible         =   0   'False
      Width           =   2220
   End
   Begin VB.CommandButton Command3 
      Caption         =   "S&aldos Promedios"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   10290
      Picture         =   "BalanceC.frx":001F
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2205
      Width           =   1170
   End
   Begin VB.TextBox TextCotiza 
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
      Left            =   10290
      MaxLength       =   11
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   420
      Width           =   1170
   End
   Begin VB.CommandButton Command4 
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
      Height          =   1065
      Left            =   10290
      Picture         =   "BalanceC.frx":0461
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4515
      Width           =   1170
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Imprimir Resultados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   10290
      Picture         =   "BalanceC.frx":06E3
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3360
      Width           =   1170
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Listar Balance"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   10290
      Picture         =   "BalanceC.frx":0D4D
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1050
      Width           =   1170
   End
   Begin VB.Data DataBalanceG 
      Caption         =   "BalanceG"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   210
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1785
      Visible         =   0   'False
      Width           =   2220
   End
   Begin VB.Data DataFechaBal 
      Caption         =   "FechaBal"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   330
      Left            =   210
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1470
      Visible         =   0   'False
      Width           =   2220
   End
   Begin VB.Data DataCtas 
      Caption         =   "Ctas"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   330
      Left            =   210
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   840
      Visible         =   0   'False
      Width           =   2220
   End
   Begin MSMask.MaskEdBox MBoxFechaI 
      Height          =   330
      Left            =   8820
      TabIndex        =   3
      Top             =   105
      Width           =   1380
      _ExtentX        =   2434
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
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &Hasta"
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
      Left            =   8085
      TabIndex        =   2
      Top             =   420
      Width           =   750
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &Desde"
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
      Left            =   8085
      TabIndex        =   1
      Top             =   105
      Width           =   750
   End
   Begin VB.Label Label11 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Cotizacion:"
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
      Left            =   10290
      TabIndex        =   5
      Top             =   105
      Width           =   1170
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   570
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   7905
   End
End
Attribute VB_Name = "BalanceConsolidado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
  RatonReloj
  DBGBalance.Visible = False
  Dolar = Round(Dolar)
  sSQL = "UPDATE Balance_Consolidado " _
       & "SET TOTAL = 0.004 "
  UpdateData DataBalanceG, sSQL
  sSQL = "UPDATE Balance_General " _
       & "SET Total = 0.004 "
  UpdateData DataBalanceG, sSQL
  sSQL = "UPDATE Estado_Resultado " _
       & "SET Total = 0.004 "
  UpdateData DataBalanceG, sSQL
  sSQL = "UPDATE Balance_Consolidado " _
       & "SET TOTAL = Saldo_MN + (Saldo_ME * " & Dolar & ") " _
       & "WHERE Saldo_ME <> 0.004 " _
       & "AND Saldo_MN <> 0.004 "
  UpdateData DataBalanceG, sSQL
  sSQL = "UPDATE Balance_General " _
       & "SET Total = Parcial + (Analitico * " & Dolar & ") " _
       & "WHERE Parcial <> 0.004 " _
       & "AND Analitico <> 0.004 "
  UpdateData DataBalanceG, sSQL
  sSQL = "UPDATE Estado_Resultado " _
       & "SET Total = Parcial + (Analitico * " & Dolar & ") " _
       & "WHERE Parcial <> 0.004 " _
       & "AND Analitico <> 0.004 "
  UpdateData DataBalanceG, sSQL
  sSQL = "SELECT Codigo,Cuenta,Saldo_ME,Saldo_MN,TOTAL " _
       & "FROM Balance_Consolidado " _
       & "WHERE DG = 'G' "
  SelectAdodcGrid datagrisconsidado, AdodcConsolidado, sSQL
  Opcion = 1
  DBGBalance.Visible = True
  RatonNormal
End Sub

Private Sub Command2_Click()
  DBGBalance.Visible = False
  If Opcion = 1 Then
     SQLMsg1 = "BALANCE CONSOLIDADO"
  Else
     SQLMsg1 = "BALANCE DE PROMEDIOS"
  End If
  SQLMsg2 = "AL " & FechaStrg(FechaFin)
  ImprimirGeneralCon DataBalanceG, Opcion, True
  DBGBalance.Visible = True
End Sub

Private Sub Command3_Click()
Dim NumDia As Integer
Dim Saldos_Prom_MN(31) As Double
Dim Saldos_Prom_ME(31) As Double
  RatonReloj
  FechaValida MBoxFechaI, False
  FechaValida MBoxFechaF, False
  NumDia = Day(MBoxFechaF.Text)
  FechaIni = BuscarFecha(MBoxFechaI.Text)
  FechaFin = BuscarFecha(MBoxFechaF.Text)
  DBGBalance.Visible = False
  sSQL = "DELETE * " _
       & "FROM Saldo_Promedios_" & CodigoUsuario & " "
  DeleteData DataPromedio, sSQL
  sSQL = "SELECT * " _
       & "FROM Saldo_Promedios_" & CodigoUsuario & " "
  SelectAdodc DataPromedio, sSQL
  sSQL = "SELECT Cta, Fecha, Saldo, Saldo_ME " _
       & "FROM Transacciones " _
       & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
       & "AND T <> 'A' " _
       & "ORDER BY Cta, Fecha "
  SelectAdodc DataCtas, sSQL
  With DataCtas.Recordset
   If .RecordCount > 0 Then
       For I = 1 To 31
           Saldos_Prom_MN(I) = 0
           Saldos_Prom_ME(I) = 0
       Next
       Codigo = .Fields("Cta")
       Do While Not .EOF
          If Codigo <> .Fields("Cta") Then
             Cadena = "Codigo: " & Codigo & vbCrLf
             Saldo = Saldos_Prom_MN(1)
             Saldo_ME = Saldos_Prom_ME(1)
             For I = 1 To NumDia
                 If Saldos_Prom_MN(I) <> 0 Then
                    Saldo = Saldos_Prom_MN(I)
                 Else
                    Saldos_Prom_MN(I) = Saldo
                 End If
                 If Saldos_Prom_ME(I) <> 0 Then
                    Saldo_ME = Saldos_Prom_ME(I)
                 Else
                    Saldos_Prom_ME(I) = Saldo_ME
                 End If
             Next
             Saldo = 0
             Saldo_ME = 0
             For I = 1 To 31
                 Saldo = Saldo + Saldos_Prom_MN(I)
                 Saldo_ME = Saldo_ME + Saldos_Prom_ME(I)
             Next
             DataPromedio.Recordset.AddNew
             DataPromedio.Recordset.Fields("Codigo") = Codigo
             DataPromedio.Recordset.Fields("Saldo_MN") = Round(Saldo / NumDia)
             DataPromedio.Recordset.Fields("Saldo_ME") = Round(Saldo_ME / NumDia)
             DataPromedio.Recordset.Update
             For I = 1 To 31
                 Saldos_Prom_MN(I) = 0
                 Saldos_Prom_ME(I) = 0
             Next
             Codigo = .Fields("Cta")
          End If
          Dia = Day(.Fields("Fecha"))
          Saldos_Prom_MN(Dia) = .Fields("Saldo")
          Saldos_Prom_ME(Dia) = .Fields("Saldo_ME")
         .MoveNext
       Loop
   End If
  End With
  sSQL = "SELECT SP.Codigo,C.Cuenta,SP.Saldo_MN,SP.Saldo_ME FROM Saldo_Promedios_" & CodigoUsuario & " As SP, Catalogo As C "
  sSQL = sSQL & "WHERE C.Codigo = SP.Codigo "
  sSQL = sSQL & "ORDER BY SP.Codigo "
  SelectDataGrid datagrisconsidado, AdodcConsolidado, sSQL
  Opcion = 2
  DataGridConsolidado.Visible = True
  RatonNormal
End Sub

Private Sub Command4_Click()
  Unload BalanceConsolidado
End Sub

Private Sub Form_Activate()
  sSQL = "SELECT * FROM FechaBalance "
  SelectAdodc DataFechaBal, sSQL
  If DataFechaBal.Recordset.RecordCount > 0 Then
     FechaIni = DataFechaBal.Recordset.Fields("Fecha_Inicial")
     FechaFin = DataFechaBal.Recordset.Fields("Fecha_Final")
  End If
  Label5.Caption = Empresa & " : BALANCE GENERAL" & vbCrLf & "AL " & FechaStrg(FechaFin)
  RatonNormal
End Sub

Private Sub Form_Load()
  CentrarForm BalanceConsolidado
  DataCtas.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
  DataPromedio.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
  DataFechaBal.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
  DataBalanceG.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
  ConectarAdodc AdodcConsolidado, RutaEmpresa & "\CONTABIL.MDB"
End Sub

Private Sub MBoxFechaF_GotFocus()
  MarcarTexto MBoxFechaF
End Sub

Private Sub MBoxFechaF_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBoxFechaF_LostFocus()
  FechaValida MBoxFechaF, False
End Sub

Private Sub MBoxFechaI_GotFocus()
  MarcarTexto MBoxFechaI
End Sub

Private Sub MBoxFechaI_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBoxFechaI_LostFocus()
  FechaValida MBoxFechaI, False
End Sub

Private Sub TextCotiza_GotFocus()
   TextCotiza.Text = Dolar
End Sub

Private Sub TextCotiza_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub TextCotiza_LostFocus()
   If Val(TextCotiza.Text) <= 0 Then TextCotiza.Text = "0"
   Dolar = Val(TextCotiza.Text)
End Sub

