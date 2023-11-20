VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Begin VB.Form FListarSaldoCtas 
   Caption         =   "SALDOS DE DEPOSITOS DE AHORROS"
   ClientHeight    =   7860
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11790
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7860
   ScaleWidth      =   11790
   WindowState     =   2  'Maximized
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
      Height          =   960
      Left            =   7980
      Picture         =   "FlistSal.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   525
      Width           =   1065
   End
   Begin MSDataGridLib.DataGrid DGCuentas 
      Bindings        =   "FlistSal.frx":08CA
      Height          =   5790
      Left            =   105
      TabIndex        =   6
      Top             =   1680
      Width           =   11460
      _ExtentX        =   20214
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
      Left            =   9135
      Picture         =   "FlistSal.frx":08E3
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   525
      Width           =   1065
   End
   Begin VB.Frame Frame1 
      Height          =   1065
      Left            =   210
      TabIndex        =   1
      Top             =   420
      Width           =   2325
      Begin MSMask.MaskEdBox MBoxFechaF 
         Height          =   330
         Left            =   945
         TabIndex        =   5
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
      Begin MSMask.MaskEdBox MBoxFecha 
         Height          =   330
         Left            =   945
         TabIndex        =   3
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
      Begin VB.Label Label4 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Desde:"
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
         Width           =   855
      End
      Begin VB.Label Label3 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Hasra:"
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
         TabIndex        =   4
         Top             =   630
         Width           =   855
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   1485
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   11460
      _ExtentX        =   20214
      _ExtentY        =   2619
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "SALDO DE LIBRETAS"
      TabPicture(0)   =   "FlistSal.frx":11AD
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label16"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Command3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "TextSaldoCont"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "TextSaldoDisp"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "TextInt"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "SALDO DE LIBRETAS POR RUBROS"
      TabPicture(1)   =   "FlistSal.frx":11C9
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Command5"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "SALDOS E INCONSISTENCIA"
      TabPicture(2)   =   "FlistSal.frx":11E5
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Command4"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "PRESMATOS E INCONSISTENCIA"
      TabPicture(3)   =   "FlistSal.frx":1201
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Command6"
      Tab(3).ControlCount=   1
      Begin VB.CommandButton Command6 
         Caption         =   "&Inconsistencia"
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
         Left            =   -68700
         Picture         =   "FlistSal.frx":121D
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   420
         Width           =   1485
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&Inconsistencia"
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
         Left            =   -68700
         Picture         =   "FlistSal.frx":165F
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   420
         Width           =   1485
      End
      Begin VB.TextBox TextInt 
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
         ForeColor       =   &H00404040&
         Height          =   360
         Left            =   4620
         MaxLength       =   14
         MultiLine       =   -1  'True
         TabIndex        =   7
         Text            =   "FlistSal.frx":1AA1
         Top             =   1020
         Width           =   2010
      End
      Begin VB.TextBox TextSaldoDisp 
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
         ForeColor       =   &H00404040&
         Height          =   360
         Left            =   4620
         MaxLength       =   14
         MultiLine       =   -1  'True
         TabIndex        =   9
         Text            =   "FlistSal.frx":1AA8
         Top             =   705
         Width           =   2010
      End
      Begin VB.CommandButton Command5 
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
         Left            =   -68280
         Picture         =   "FlistSal.frx":1AAF
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   420
         Width           =   1065
      End
      Begin VB.TextBox TextSaldoCont 
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
         ForeColor       =   &H00404040&
         Height          =   360
         Left            =   4620
         MaxLength       =   14
         MultiLine       =   -1  'True
         TabIndex        =   10
         Text            =   "FlistSal.frx":1EF1
         Top             =   390
         Width           =   2010
      End
      Begin VB.CommandButton Command3 
         Caption         =   "S&aldos del día"
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
         Left            =   6720
         Picture         =   "FlistSal.frx":1EF8
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   420
         Width           =   1065
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Interes Acumulado"
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
         Left            =   2520
         TabIndex        =   11
         Top             =   1020
         Width           =   2115
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Saldo Disponible"
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
         Left            =   2520
         TabIndex        =   12
         Top             =   705
         Width           =   2115
      End
      Begin VB.Label Label16 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Saldo Contable"
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
         Left            =   2520
         TabIndex        =   13
         Top             =   390
         Width           =   2115
      End
   End
   Begin MSAdodcLib.Adodc AdoCtas 
      Height          =   330
      Left            =   315
      Top             =   2520
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
      Caption         =   "Ctas"
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
   Begin MSAdodcLib.Adodc AdoSaldo 
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
      Caption         =   "Saldo"
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
   Begin MSAdodcLib.Adodc AdoCuentas 
      Height          =   330
      Left            =   105
      Top             =   7455
      Width           =   11460
      _ExtentX        =   20214
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
      Caption         =   "Cuentas"
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
   Begin MSAdodcLib.Adodc AdoAux 
      Height          =   330
      Left            =   315
      Top             =   2205
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
End
Attribute VB_Name = "FListarSaldoCtas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
   MensajeEncabado = "SALDO DE LIBRETAS DEL " & MBoxFecha.Text
   DGCuentas.Visible = False
   ImprimirSaldosLibretas AdoCuentas, CCur(SaldoDisp), CCur(SaldoCont)
   DGCuentas.Visible = True
End Sub

Private Sub Command2_Click()
 Unload FListarSaldoCtas
End Sub

Private Sub Command3_Click()
  SaldoCont = 0: SaldoDisp = 0: TotalInteres = 0
  FechaValida MBoxFecha
  FechaValida MBoxFechaF
  FechaIni = BuscarFecha(MBoxFecha)
  FechaFin = BuscarFecha(MBoxFechaF)
  sSQL = "SELECT C.T,Cta.Cuenta_No,Cliente,SL.Fecha,SL.Saldo_Disp,SL.Saldo_Cont,SL.Interes " _
       & "FROM Clientes_Datos_Extras As Cta,Trans_Intereses As SL,Clientes C " _
       & "WHERE SL.Fecha Between #" & FechaIni & "# and #" & FechaFin & "# " _
       & "AND Cta.Tipo_Dato = 'LIBRETAS' " _
       & "AND Cta.Cuenta_No = SL.Cuenta_No " _
       & "AND Cta.Codigo = C.Codigo " _
       & "ORDER BY Cta.Cuenta_No "
  SelectDataGrid DGCuentas, AdoCuentas, sSQL
  DGCuentas.Visible = False
  With AdoCuentas.Recordset
   If .RecordCount > 0 Then
       RatonReloj
       Do While Not .EOF
          SaldoDisp = SaldoDisp + .Fields("Saldo_Disp")
          SaldoCont = SaldoCont + .Fields("Saldo_Cont")
          TotalInteres = TotalInteres + .Fields("Interes")
         .MoveNext
       Loop
      .MoveFirst
   End If
  End With
  TextSaldoDisp.Text = Format(SaldoDisp, "#,##0.00")
  TextSaldoCont.Text = Format(SaldoCont, "#,##0.00")
  TextInt.Text = Format(TotalInteres, "#,##0.00")
  DGCuentas.Visible = True
  RatonNormal
End Sub

Private Sub Command4_Click()
  FechaValida MBoxFechaF
  sSQL = "DELETE * " _
       & "FROM Trans_Saldo_Libretas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Fecha = #" & BuscarFecha(MBoxFechaF) & "# "
  ConectarAdoExecute sSQL
  sSQL = "SELECT * " _
       & "FROM Clientes_Datos_Extras " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "ORDER BY Cuenta_No "
  SelectAdodc AdoAux, sSQL
  Contador = 0
  With AdoAux.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          Cuenta_No = .Fields("Cuenta_No")
          CodigoCli = .Fields("Codigo")
          FListarSaldoCtas.Caption = Cuenta_No & " - " & Format(Contador / .RecordCount, "#00%")
          Contador = Contador + 1
          SaldoCont = 0
          SaldoDisp = 0
          Saldo = 0
          sSQL = "SELECT TOP 1 * " _
               & "FROM Trans_Libretas " _
               & "WHERE Cuenta_No = '" & Cuenta_No & "' " _
               & "AND Fecha <= #" & BuscarFecha(MBoxFechaF) & "# " _
               & "ORDER BY Fecha DESC,IDT DESC,Hora DESC,ID DESC "
          SelectAdodc AdoSaldo, sSQL
          If AdoSaldo.Recordset.RecordCount > 0 Then
             SaldoCont = AdoSaldo.Recordset.Fields("Saldo_Cont")
             SaldoDisp = AdoSaldo.Recordset.Fields("Saldo_Disp")
          End If
          sSQL = "SELECT Cuenta_No, SUM(Creditos) AS TCreditos, SUM(Debitos) AS TDebitos " _
               & "FROM Trans_Libretas " _
               & "WHERE Cuenta_No = '" & Cuenta_No & "' " _
               & "AND TP <> 'EFCH' " _
               & "AND Fecha <= #" & BuscarFecha(MBoxFechaF) & "# " _
               & "GROUP BY Cuenta_No "
          SelectAdodc AdoSaldo, sSQL
          If AdoSaldo.Recordset.RecordCount > 0 Then
             Saldo = AdoSaldo.Recordset.Fields("TCreditos") - AdoSaldo.Recordset.Fields("TDebitos")
          End If
          SetAdoAddNew "Trans_Saldo_Libretas"
          SetAdoFields "Cuenta_No", Cuenta_No
          SetAdoFields "Codigo", CodigoCli
          SetAdoFields "Saldo_Cont", SaldoCont
          SetAdoFields "Saldo_Disp", SaldoDisp
          SetAdoFields "Inconsistencia", Saldo
          SetAdoFields "Item", NumEmpresa
          SetAdoUpdate
         .MoveNext
       Loop
   End If
  End With
  sSQL = "SELECT C.Cliente,TSL.Cuenta_No,TSL.Saldo_Cont,TSL.Saldo_Disp,TSL.Inconsistencia " _
       & "FROM Trans_Saldo_Libretas As TSL,Clientes As C " _
       & "WHERE TSL.Item = '" & NumEmpresa & "' " _
       & "AND TSL.Fecha = #" & BuscarFecha(MBoxFechaF) & "# " _
       & "AND TSL.Saldo_Cont <> TSL.Inconsistencia " _
       & "AND TSL.Codigo=C.Codigo " _
       & "ORDER BY TSL.Cuenta_No "
  SelectDataGrid DGCuentas, AdoCuentas, sSQL
End Sub

Private Sub Command5_Click()
Dim TotCodTP() As Campos_Tabla
Dim ContTP As Integer
  FechaValida MBoxFecha
  FechaValida MBoxFechaF
  FechaIni = BuscarFecha(MBoxFecha)
  FechaFin = BuscarFecha(MBoxFechaF)
  sSQL = "DELETE * " _
       & "FROM Saldo_Proceso " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' "
  ConectarAdoExecute sSQL
  ContTP = 1
  sSQL = "SELECT * " _
       & "FROM Catalogo_Proceso " _
       & "WHERE DC <> '.' " _
       & "ORDER BY DC,TP "
  SelectAdodc AdoSaldo, sSQL
  If AdoSaldo.Recordset.RecordCount > 0 Then ContTP = AdoSaldo.Recordset.RecordCount
  ReDim TotCodTP(ContTP) As Campos_Tabla
  With AdoSaldo.Recordset
   If .RecordCount > 0 Then
       ContTP = 0
       Do While Not .EOF
          TotCodTP(ContTP).Campo = Replace(.Fields("TP"), "/", "")
          TotCodTP(ContTP).Valor = 0
          ContTP = ContTP + 1
         .MoveNext
       Loop
       TotCodTP(ContTP).Campo = "Saldo"
       TotCodTP(ContTP).Valor = 0
   End If
  End With
  Ln_No = 1
  sSQL = "SELECT Cuenta_No,TP,SUM(Creditos-Debitos) As Monto " _
       & "FROM Trans_Libretas " _
       & "WHERE Fecha <= #" & FechaFin & "# " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "GROUP BY Cuenta_No,TP " _
       & "ORDER BY Cuenta_No,TP "
  SelectAdodc AdoAux, sSQL
  With AdoAux.Recordset
   If .RecordCount > 0 Then
       Cuenta_No = .Fields("Cuenta_No")
       Do While Not .EOF
          Codigo = Replace(.Fields("TP"), "/", "")
          If Cuenta_No <> .Fields("Cuenta_No") Then
             SetAdoAddNew "Saldo_Proceso"
             SetAdoFields "Cuenta_No", Cuenta_No
             Saldo = 0
             For I = 0 To ContTP
                 SetAdoFields TotCodTP(I).Campo, TotCodTP(I).Valor
                 Saldo = Saldo + TotCodTP(I).Valor
                 TotCodTP(I).Valor = 0
             Next I
             SetAdoFields "Saldo", Saldo
             SetAdoFields "CodigoU", CodigoUsuario
             SetAdoFields "Item", NumEmpresa
             SetAdoFields "Ln_No", Ln_No
             SetAdoUpdate
             Cuenta_No = .Fields("Cuenta_No")
             Ln_No = Ln_No + 1
          End If
          For I = 0 To ContTP
              If TotCodTP(I).Campo = Codigo Then
                 TotCodTP(I).Valor = TotCodTP(ContTP).Valor + .Fields("Monto")
              End If
          Next I
         .MoveNext
       Loop
       SetAdoAddNew "Saldo_Proceso"
       SetAdoFields "Cuenta_No", Cuenta_No
       Saldo = 0
       For I = 0 To ContTP
           SetAdoFields TotCodTP(I).Campo, TotCodTP(I).Valor
           Saldo = Saldo + TotCodTP(I).Valor
           TotCodTP(I).Valor = 0
       Next I
       SetAdoFields "Saldo", Saldo
       SetAdoFields "CodigoU", CodigoUsuario
       SetAdoFields "Item", NumEmpresa
       SetAdoFields "Ln_No", Ln_No
       SetAdoUpdate
   End If
  End With
  Ln_No = 9999
  sSQL = "SELECT "
       For I = 0 To ContTP
           sSQL = sSQL & "SUM(" & TotCodTP(I).Campo & ") As T" & TotCodTP(I).Campo & ","
       Next I
  sSQL = sSQL & "Item " _
       & "FROM Saldo_Proceso " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "GROUP BY Item "
  SelectAdodc AdoAux, sSQL
  With AdoAux.Recordset
   If .RecordCount > 0 Then
       For I = 0 To .Fields.Count - 1
           Codigo1 = MidStrg(.Fields(I).Name, 2, Len(.Fields(I).Name))
           Total = .Fields(I)
           For J = 0 To ContTP
               If Codigo1 = TotCodTP(J).Campo Then TotCodTP(J).Valor = Total
           Next J
       Next I
   End If
  End With
  sSQL = "SELECT C.Cliente,SP.Cuenta_No,"
  For J = 0 To ContTP
      'MsgBox TotCodTP(J).Campo
      If TotCodTP(J).Valor <> 0 Then sSQL = sSQL & TotCodTP(J).Campo & ","
  Next J
  sSQL = sSQL & "Item " _
       & "FROM Saldo_Proceso As SP,Clientes As C " _
       & "WHERE SP.Codigo = C.Codigo " _
       & "AND SP.Item = '" & NumEmpresa & "' " _
       & "AND SP.CodigoU = '" & CodigoUsuario & "' " _
       & "ORDER BY C.Cliente,SP.Cuenta_No "
  SelectDataGrid DGCuentas, AdoCuentas, sSQL
End Sub

Private Sub Command6_Click()
  sSQL = "SELECT * " _
       & "FROM Trans_Prestamos " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND T = 'P' " _
       & "ORDER BY Credito_No,Cuenta_No,Cuota_No,Fecha "
  SelectAdodc AdoAux, sSQL
  Contador = 0
  I = 1
  With AdoAux.Recordset
   If .RecordCount > 0 Then
       Credito_No = .Fields("Credito_No")
       I = .Fields("Cuota_No")
       Do While Not .EOF
          If Credito_No <> .Fields("Credito_No") Then
             Credito_No = .Fields("Credito_No")
             I = .Fields("Cuota_No")
          End If
         .Fields("Numero") = I
          FListarSaldoCtas.Caption = Cuenta_No & " - " & Format(Contador / .RecordCount, "#00%")
          Contador = Contador + 1
          I = I + 1
         .Update
         .MoveNext
       Loop
   End If
  End With
  sSQL = "SELECT * " _
       & "FROM Trans_Prestamos " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND T = 'P' " _
       & "AND Cuota_No <> Numero " _
       & "ORDER BY Cuenta_No,Credito_No,Fecha "
  SelectDataGrid DGCuentas, AdoCuentas, sSQL
End Sub

Private Sub DGCuentas_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF1 Then GenerarDataTexto FListarSaldoCtas, AdoCuentas, True
End Sub

Private Sub Form_Activate()
   sSQL = "SELECT * " _
        & "FROM Catalogo_Proceso " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND DC <> '.' " _
        & "ORDER BY DC,TP "
   SelectAdodc AdoSaldo, sSQL
   With AdoSaldo.Recordset
    If .RecordCount > 0 Then
        If Existe_Tabla("Saldo_Proceso") Then
           NombreCampo = ""
           sSQL = "SELECT * " _
                & "FROM Saldo_Proceso " _
                & "WHERE Ln_No = 0 "
           SelectAdodc AdoCtas, sSQL
           Do While Not .EOF
              NombreCampo = Replace(.Fields("TP"), "/", "")
              Evaluar = False
              For J = 0 To AdoCtas.Recordset.Fields.Count - 1
                 If NombreCampo = AdoCtas.Recordset.Fields(J).Name Then Evaluar = True
              Next J
              If Evaluar = False Then
                 SQL1 = "ALTER TABLE Saldo_Proceso "
                 If SQL_Server Then
                    SQL1 = SQL1 & "ADD [" & NombreCampo & "] MONEY NULL; "
                 Else
                    SQL1 = SQL1 & "ADD [" & NombreCampo & "] CURRENCY NULL; "
                 End If
                 ConectarAdoExecute SQL1
              End If
             .MoveNext
           Loop
           
        Else
           SQL1 = "CREATE TABLE Saldo_Proceso ("
           If SQL_Server Then
              SQL1 = SQL1 _
                   & "[Cuenta_No] NVARCHAR (10) NULL," _
                   & "[Año] NVARCHAR (4) NULL," _
                   & "[Mes] NVARCHAR (3) NULL,"
           Else
              SQL1 = SQL1 _
                   & "[Cuenta_No] TEXT (10) NULL," _
                   & "[Año] TEXT (4) NULL," _
                   & "[Mes] TEXT (3) NULL,"
           End If
           Do While Not .EOF
              NombreCampo = Replace(.Fields("TP"), "/", "")
              If SQL_Server Then
                 SQL1 = SQL1 & "[" & NombreCampo & "] MONEY NULL,"
              Else
                 SQL1 = SQL1 & "[" & NombreCampo & "] CURRENCY NULL,"
              End If
             .MoveNext
           Loop
           If SQL_Server Then
              SQL1 = SQL1 _
                   & "[Saldo] MONEY NULL," _
                   & "[Ln_No] SMALLINT NULL," _
                   & "[Codigo] NVARCHAR (10) NULL," _
                   & "[CodigoU] NVARCHAR (10) NULL," _
                   & "[Item] NVARCHAR (3) NULL); "
           Else
              SQL1 = SQL1 _
                   & "[Saldo] CURRENCY NULL," _
                   & "[Ln_No] SHORT NULL," _
                   & "[Codigo] TEXT (10) NULL," _
                   & "[CodigoU] TEXT (10) NULL," _
                   & "[Item] TEXT (3) NULL); "
           End If
           ConectarAdoExecute SQL1
        End If
    End If
   End With
   
   SSTab1.width = MDI_X_Max - 30
   AdoCuentas.width = MDI_X_Max - 30
   DGCuentas.width = MDI_X_Max - 30
   DGCuentas.Height = MDI_Y_Max - DGCuentas.Top - 300
   AdoCuentas.Top = DGCuentas.Top + DGCuentas.Height + 30
    
   RatonNormal
   MBoxFecha.SetFocus
End Sub

Private Sub Form_Load()
  ConectarAdodc AdoAux
  ConectarAdodc AdoCtas
  ConectarAdodc AdoSaldo
  ConectarAdodc AdoCuentas
End Sub

Private Sub MBoxFecha_GotFocus()
  MarcarTexto MBoxFecha
End Sub

Private Sub MBoxFecha_LostFocus()
   FechaValida MBoxFecha
End Sub

Private Sub MBoxFechaF_GotFocus()
  MarcarTexto MBoxFechaF
End Sub

Private Sub MBoxFechaF_LostFocus()
   FechaValida MBoxFechaF
End Sub
