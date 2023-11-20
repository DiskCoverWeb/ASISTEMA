VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form ChequesTransito 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CHEQUES EN TRANSITO"
   ClientHeight    =   7155
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11280
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7155
   ScaleWidth      =   11280
   ShowInTaskbar   =   0   'False
   Begin VB.Data DataTP 
      Caption         =   "FlujoCaja"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   7665
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   420
      Visible         =   0   'False
      Width           =   2010
   End
   Begin MSDBCtls.DBCombo DBCTP 
      Bindings        =   "CheqTran.frx":0000
      DataSource      =   "DataTP"
      Height          =   315
      Left            =   8610
      TabIndex        =   24
      Top             =   1155
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   556
      _Version        =   327680
      Text            =   "DBCombo1"
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
   Begin VB.CheckBox CheckTP 
      Caption         =   "Tipo Proc"
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
      Left            =   8610
      TabIndex        =   25
      Top             =   840
      Width           =   1275
   End
   Begin MSDBGrid.DBGrid DBGFlujoCajaCheq 
      Bindings        =   "CheqTran.frx":0011
      Height          =   5055
      Left            =   105
      OleObjectBlob   =   "CheqTran.frx":002D
      TabIndex        =   12
      Top             =   1575
      Width           =   9885
   End
   Begin VB.Frame Frame3 
      Height          =   645
      Left            =   105
      TabIndex        =   21
      Top             =   840
      Width           =   4845
      Begin VB.OptionButton OpcG 
         Caption         =   "Ordenar por Grupo"
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
         TabIndex        =   23
         Top             =   210
         Value           =   -1  'True
         Width           =   1905
      End
      Begin VB.OptionButton OpcT 
         Caption         =   "Ordenar por Transaccion"
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
         TabIndex        =   22
         Top             =   210
         Width           =   2535
      End
   End
   Begin VB.Data DataUsuario 
      Caption         =   "Usuario"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2730
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2730
      Visible         =   0   'False
      Width           =   2430
   End
   Begin MSDBCtls.DBCombo DBCUsuario 
      Bindings        =   "CheqTran.frx":0A00
      DataSource      =   "DataUsuario"
      Height          =   315
      Left            =   5040
      TabIndex        =   20
      Top             =   1155
      Width           =   3480
      _ExtentX        =   6138
      _ExtentY        =   556
      _Version        =   327680
      Text            =   "DBCombo1"
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
   Begin VB.CheckBox CheckUsuario 
      Caption         =   "Por Cajero(a):"
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
      Left            =   5040
      TabIndex        =   19
      Top             =   840
      Width           =   3480
   End
   Begin VB.Frame Frame2 
      Height          =   750
      Left            =   3885
      TabIndex        =   14
      Top             =   0
      Width           =   1590
      Begin VB.OptionButton OpcME 
         Caption         =   "ME"
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
         TabIndex        =   4
         Top             =   315
         Value           =   -1  'True
         Width           =   645
      End
      Begin VB.OptionButton OpcMN 
         Caption         =   "MN"
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
         Left            =   840
         TabIndex        =   5
         Top             =   315
         Width           =   645
      End
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
      Height          =   960
      Left            =   10080
      Picture         =   "CheqTran.frx":0A16
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4620
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
      Height          =   960
      Left            =   10080
      Picture         =   "CheqTran.frx":0C98
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3570
      Width           =   1065
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Grabar"
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
      Left            =   10080
      Picture         =   "CheqTran.frx":1302
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1470
      Width           =   1065
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Caja Anterior"
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
      Left            =   10080
      Picture         =   "CheqTran.frx":1744
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2520
      Width           =   1065
   End
   Begin VB.Frame Frame1 
      Height          =   750
      Left            =   1575
      TabIndex        =   13
      Top             =   0
      Width           =   2220
      Begin VB.OptionButton OpcE 
         Caption         =   "Egreso"
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
         Left            =   1155
         TabIndex        =   3
         Top             =   315
         Width           =   960
      End
      Begin VB.OptionButton OpcI 
         Caption         =   "Ingreso"
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
         TabIndex        =   2
         Top             =   315
         Value           =   -1  'True
         Width           =   960
      End
   End
   Begin VB.TextBox TextCant 
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
      ForeColor       =   &H80000002&
      Height          =   360
      Left            =   5565
      MaxLength       =   14
      MultiLine       =   -1  'True
      TabIndex        =   7
      Text            =   "CheqTran.frx":1B86
      Top             =   420
      Width           =   2010
   End
   Begin VB.Data DataCaja 
      Caption         =   "FlujoCaja"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   7665
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   105
      Visible         =   0   'False
      Width           =   2010
   End
   Begin VB.Data DataFlujoCajaCheq 
      Caption         =   "FlujoCajaCheq"
      Connect         =   "Access"
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
      Top             =   2730
      Visible         =   0   'False
      Width           =   2430
   End
   Begin MSMask.MaskEdBox MBoxFecha 
      Height          =   330
      Left            =   105
      TabIndex        =   1
      Top             =   420
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   582
      _Version        =   327680
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
   Begin VB.Data DataFlujoCajaEfec 
      Caption         =   "FlujoCajaEfec"
      Connect         =   "Access"
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
      Top             =   2625
      Visible         =   0   'False
      Width           =   2430
   End
   Begin VB.Label LabelIngCheqME 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
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
      Left            =   8190
      TabIndex        =   17
      Top             =   6615
      Width           =   1800
   End
   Begin VB.Label LabelIngCheqMN 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
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
      Left            =   3570
      TabIndex        =   15
      Top             =   6615
      Width           =   1800
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "INGRESOS CHEQUE M/N "
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
      TabIndex        =   16
      Top             =   6615
      Width           =   3480
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "INGRESOS CHEQUE M/E "
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
      Left            =   5355
      TabIndex        =   18
      Top             =   6615
      Width           =   2850
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " FECHA:"
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
      Width           =   1380
   End
   Begin VB.Label Label16 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Cantidad"
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
      Left            =   5565
      TabIndex        =   6
      Top             =   105
      Width           =   2010
   End
End
Attribute VB_Name = "ChequesTransito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub SumarIngEgr(DtaCajaCheq As Data, DtaCajaEfec As Data)
  RatonReloj
  Saldo = 0: Total = 0
  Debe = 0: Haber = 0
  Debe_ME = 0: Haber_ME = 0
  With DtaCajaCheq.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       Do While Not .EOF
          If .Fields("ME") Then
              Saldo = Saldo + .Fields("Debitos")
          Else
              Total = Total + .Fields("Debitos")
          End If
         .MoveNext
       Loop
      .MoveFirst
   End If
  End With
  With DtaCajaEfec.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       Do While Not .EOF
          Select Case .Fields("TP")
            Case "APER", "BOVE"
                 If .Fields("ME") Then
                     Debe_ME = Debe_ME + .Fields("Debitos")
                     Haber_ME = Haber_ME + .Fields("Creditos")
                 Else
                     Debe = Debe + .Fields("Debitos")
                     Haber = Haber + .Fields("Creditos")
                 End If
          End Select
          Select Case Mid(.Fields("TP"), 1, 3)
            Case "RET", "DEP"
                 If .Fields("ME") Then
                     Debe_ME = Debe_ME + .Fields("Debitos")
                     Haber_ME = Haber_ME + .Fields("Creditos")
                 Else
                     Debe = Debe + .Fields("Debitos")
                     Haber = Haber + .Fields("Creditos")
                 End If
          End Select
         .MoveNext
       Loop
      .MoveFirst
   End If
  End With
  LabelIngCheqMN.Caption = Format(Total, "#,##0.00")
  LabelIngCheqME.Caption = Format(Saldo, "#,##0.00")
  LabelIngresos.Caption = Format(Debe, "#,##0.00")
  LabelEgresos.Caption = Format(Haber, "#,##0.00")
  LabelSaldo.Caption = Format(Debe - Haber, "#,##0.00")
  LabelIngresosME.Caption = Format(Debe_ME, "#,##0.00")
  LabelEgresosME.Caption = Format(Haber_ME, "#,##0.00")
  LabelSaldoME.Caption = Format(Debe_ME - Haber_ME, "#,##0.00")
  RatonNormal
End Sub

Private Sub Command1_Click()
  MiFecha = BuscarFecha(MBoxFecha.Text)
  sSQL = "SELECT ME,Fecha,TP,Cuenta_No,Debitos,Creditos,Usuario "
  sSQL = sSQL & "FROM Transacciones "
  sSQL = sSQL & "WHERE Fecha = #" & MiFecha & "# AND T = 'P' "
  sSQL = sSQL & "AND CL = False "
  If CheckUsuario.Value = 1 Then sSQL = sSQL & "AND Usuario = '" & DBCUsuario.Text & "' "
  If CheckTP.Value = 1 Then sSQL = sSQL & "AND TP = '" & DBCTP.Text & "' "
  If OpcG.Value Then
     sSQL = sSQL & "ORDER BY ME,Fecha,TP,Cuenta_No "
  Else
     sSQL = sSQL & "ORDER BY Fecha,Hora "
  End If
  SelectDBGrid DBGFlujoCajaCheq, DataFlujoCajaCheq, sSQL
  sSQL = "SELECT ME,Fecha,TP,Cuenta_No,Debitos,Creditos,Usuario "
  sSQL = sSQL & "FROM Transacciones "
  sSQL = sSQL & "WHERE Fecha = #" & MiFecha & "# AND T <> 'A' "
  sSQL = sSQL & "AND CL = False "
  If CheckUsuario.Value = 1 Then sSQL = sSQL & "AND Usuario = '" & DBCUsuario.Text & "' "
  If CheckTP.Value = 1 Then sSQL = sSQL & "AND TP = '" & DBCTP.Text & "' "
  If OpcG.Value Then
     sSQL = sSQL & "ORDER BY ME,Fecha,TP,Cuenta_No "
  Else
     sSQL = sSQL & "ORDER BY Fecha,Hora "
  End If
  SelectDBGrid DBGFlujoCajaEfec, DataFlujoCajaEfec, sSQL
  SumarIngEgr DataFlujoCajaCheq, DataFlujoCajaEfec
End Sub

Private Sub Command2_Click()
Mensajes = "Desea Grabar"
Titulo = "Formulario de Grabacion"
If BoxMensaje = 6 Then
  Debe = 0: Haber = 0
  If OpcI.Value Then Debe = Round_ME(Val(TextCant.Text)) Else Haber = Round_ME(Val(TextCant.Text))
  sSQL = "SELECT * FROM Transacciones "
  SelectData DataCaja, sSQL, False
  With DataCaja.Recordset
      .AddNew
      .Fields("T") = Normal
      .Fields("Fecha") = MBoxFecha.Text
      .Fields("Cuenta_No") = Ninguno
      .Fields("TP") = "BOVE"
      .Fields("Debitos") = 0
      .Fields("Creditos") = 0
       'If OpcME.Value Then
         .Fields("Debitos") = Debe
         .Fields("Creditos") = Haber
      ' Else
      '   .Fields("Debitos_MN") = Debe
       '  .Fields("Creditos_MN") = Haber
       'End If
      .Fields("Usuario") = DBCUsuario.Text
      .Fields("ID") = 0
      .Fields("Hora") = Format(Time, FormatoTimes)
      .Fields("Item") = NumEmpresa
      .Fields("Cheque") = Ninguno
      .Update
  End With
  sSQL = "SELECT ME,Fecha,TP,Cuenta_No,Debitos,Creditos,Usuario "
  sSQL = sSQL & "FROM Transacciones "
  sSQL = sSQL & "WHERE Fecha = #" & MiFecha & "# AND T = 'P' "
  sSQL = sSQL & "AND CL = False "
  If CheckUsuario.Value = 1 Then sSQL = sSQL & "AND Usuario = '" & DBCUsuario.Text & "' "
  sSQL = sSQL & "ORDER BY ME,Fecha,TP,Cuenta_No,Debitos DESC,Creditos "
  SelectDBGrid DBGFlujoCajaCheq, DataFlujoCajaCheq, sSQL
  sSQL = "SELECT ME,Fecha,TP,Cuenta_No,Debitos,Creditos,Usuario "
  sSQL = sSQL & "FROM Transacciones "
  sSQL = sSQL & "WHERE Fecha = #" & MiFecha & "# AND T <> 'P' "
  sSQL = sSQL & "AND CL = False "
  If CheckUsuario.Value = 1 Then sSQL = sSQL & "AND Usuario = '" & DBCUsuario.Text & "' "
  sSQL = sSQL & "ORDER BY ME,Fecha,TP,Cuenta_No,Debitos DESC,Creditos "
  SelectDBGrid DBGFlujoCajaEfec, DataFlujoCajaEfec, sSQL
  SumarIngEgr DataFlujoCajaCheq, DataFlujoCajaEfec
End If
End Sub

Private Sub Command4_Click()
  SQLMsg1 = "REPORTE DE FLUJO DE CAJA"
  ImprimirFlujoCajaCoop DataFlujoCajaEfec, True, 1, 9, OpcG.Value, True
  'SQLMsg1 = "REPORTE DE FLUJO DE CAJA CHEQUE"
  'ImprimirFlujoCajaCoop DataFlujoCajaEfec, True, 1, 9, True
End Sub

Private Sub Command5_Click()
  Unload FlujoDeCaja
End Sub

Private Sub Form_Activate()
  'MsgBox OpcG.Value
  MiFecha = BuscarFecha(FechaSistema)
  sSQL = "SELECT ME,Fecha,TP,Cuenta_No,Debitos,Creditos,Usuario "
  sSQL = sSQL & "FROM Transacciones "
  sSQL = sSQL & "WHERE Fecha = #" & MiFecha & "# AND T = 'P' "
  sSQL = sSQL & "AND CL = False "
  If CheckUsuario.Value = 1 Then sSQL = sSQL & "AND Usuario = '" & DBCUsuario.Text & "' "
  sSQL = sSQL & "AND TP = 'DEPC' "
  If OpcG.Value Then
     sSQL = sSQL & "ORDER BY ME,Fecha,TP,Cuenta_No "
  Else
     sSQL = sSQL & "ORDER BY Fecha,Hora "
  End If
  SelectDBGrid DBGFlujoCajaCheq, DataFlujoCajaCheq, sSQL
  sSQL = "SELECT ME,Fecha,TP,Cuenta_No,Debitos,Creditos,Usuario "
  sSQL = sSQL & "FROM Transacciones "
  sSQL = sSQL & "WHERE Fecha = #" & MiFecha & "# AND T <> 'A' "
  sSQL = sSQL & "AND CL = False "
  If CheckUsuario.Value = 1 Then sSQL = sSQL & "AND Usuario = '" & DBCUsuario.Text & "' "
  If CheckTP.Value = 1 Then sSQL = sSQL & "AND TP = '" & DBCTP.Text & "' "
  If OpcG.Value Then
     sSQL = sSQL & "ORDER BY ME,Fecha,TP,Cuenta_No "
  Else
     sSQL = sSQL & "ORDER BY Fecha,Hora "
  End If
  SelectDBGrid DBGFlujoCajaEfec, DataFlujoCajaEfec, sSQL
  sSQL = "SELECT TP FROM Transacciones "
  sSQL = sSQL & "GROUP BY TP "
  SelectDBCombo DBCTP, DataTP, sSQL, "TP", False
  sSQL = "SELECT Usuario FROM Transacciones "
  sSQL = sSQL & "GROUP BY Usuario "
  SelectDBCombo DBCUsuario, DataUsuario, sSQL, "Usuario", False
  SumarIngEgr DataFlujoCajaCheq, DataFlujoCajaEfec
  RatonNormal
  MBoxFecha.SetFocus
End Sub

Private Sub Form_Load()
CentrarForm ChequesTransito
DataTP.DatabaseName = RutaEmpresa & "\CAJACRED.MDB"
DataCaja.DatabaseName = RutaEmpresa & "\CAJACRED.MDB"
DataUsuario.DatabaseName = RutaEmpresa & "\CAJACRED.MDB"
DataFlujoCajaCheq.DatabaseName = RutaEmpresa & "\CAJACRED.MDB"
DataFlujoCajaEfec.DatabaseName = RutaEmpresa & "\CAJACRED.MDB"
End Sub

Private Sub MBoxFecha_GotFocus()
  MarcarTexto MBoxFecha
End Sub

Private Sub MBoxFecha_LostFocus()
  FechaValida MBoxFecha, False
End Sub

Private Sub OpcE_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub OpcI_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub OpcME_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub OpcMN_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextCant_GotFocus()
  TextCant.Text = ""
End Sub

Private Sub TextCant_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextCant_LostFocus()
  TextoValido TextCant, True
End Sub

