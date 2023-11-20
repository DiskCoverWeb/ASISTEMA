VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form DiarioESI 
   Caption         =   "ESTADO DE SITUASION INICIAL"
   ClientHeight    =   7185
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   11160
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7185
   ScaleWidth      =   11160
   WindowState     =   2  'Maximized
   Begin MSDBGrid.DBGrid DBGAsientos 
      Bindings        =   "DiariESI.frx":0000
      Height          =   6105
      Left            =   105
      OleObjectBlob   =   "DiariESI.frx":0017
      TabIndex        =   0
      Top             =   105
      Width           =   10935
   End
   Begin VB.Data DataCatalogo 
      Caption         =   "Catalogo"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   105
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5880
      Visible         =   0   'False
      Width           =   2010
   End
   Begin VB.Data DataCtas 
      Caption         =   "Ctas"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5460
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "&Grabar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   210
      TabIndex        =   1
      Top             =   6615
      Width           =   1275
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1575
      TabIndex        =   2
      Top             =   6615
      Width           =   1275
   End
   Begin VB.Data DataAsientos 
      Caption         =   "Asientos"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   105
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6195
      Width           =   10935
   End
   Begin VB.Label LabelHaber 
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
      Left            =   9240
      TabIndex        =   5
      Top             =   6615
      Width           =   1695
   End
   Begin VB.Label LabelDebe 
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
      Left            =   7455
      TabIndex        =   4
      Top             =   6615
      Width           =   1695
   End
   Begin VB.Label Label19 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Totales"
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
      Left            =   6090
      TabIndex        =   3
      Top             =   6615
      Width           =   1275
   End
End
Attribute VB_Name = "DiarioESI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdGrabar_Click()
  DBGAsientos.Visible = False
  CalculosTotalAsientos DataAsientos, LabelDebe, LabelHaber
  DBGAsientos.Visible = True
  If SumaDebe <> SumaHaber Then
     Mensajes = "Las transacciones no cuadran correctamente" & Chr(13)
     Mensajes = Mensajes & "corrija los resultados de las cuentas"
     MsgBox Mensajes
     DBGAsientos.SetFocus
  Else
     Mensajes = "Esta seguro de Grabar El Estado de Situacion Inicial"
     Titulo = "Pregunta de Grabación"
     If BoxMensaje = 6 Then
       With DataAsientos.Recordset
        RatonReloj MDIConta
        SelectData DataCatalogo, "Cuentas", False
        If .RecordCount > 0 Then
           .MoveFirst
            Do While Not .EOF
               Codigo = .Fields("CODIGO")
               Parcial_ME = .Fields("PARCIAL_ME")
               Debe = .Fields("DEBE")
               Haber = .Fields("HABER")
               If (Debe - Haber) <> 0 Then
                  sSQL = "UPDATE Cuentas SET "
                  sSQL = sSQL & "S_Parcial_ME = " & Parcial_ME & ", "
                  sSQL = sSQL & "S_Debe = " & Debe & ", "
                  sSQL = sSQL & "S_Haber = " & Haber & " "
                  sSQL = sSQL & "WHERE Codigo = '" & Codigo & "' "
                  UpdateData DataCatalogo, sSQL
               End If
              .MoveNext
            Loop
        End If
        RatonNormal MDIConta
       End With
    End If
  End If
End Sub

Private Sub DBGAsientos_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = Key_Esc Then
     DBGAsientos.Visible = False
     CalculosTotalAsientos DataAsientos, LabelDebe, LabelHaber
     DBGAsientos.Visible = True
     CmdGrabar.SetFocus
  End If
End Sub

Private Sub MBoxFecha_GotFocus()
  MBoxFecha.Text = FechaSistema
  'MarcarTexto MBoxFecha
End Sub

Private Sub MBoxFecha_LostFocus()
  'FechaValida MBoxFecha
End Sub

Private Sub TextCotiza_GotFocus()
   TextCotiza.Text = Dolar
End Sub

Private Sub TextCotiza_LostFocus()
   Dolar = Val(TextCotiza.Text)
End Sub

Private Sub CmdCancelar_Click()
   Unload DiarioESI
End Sub

Private Sub Form_Activate()
  sSQL = "DELETE * FROM Asientos_D "
  DeleteData DataAsientos, sSQL
  sSQL = "SELECT * FROM Asientos_D "
  DBMallaAsientos DiarioESI, DBGAsientos, DataAsientos, sSQL
  DBGAsientos.Visible = False
  sSQL = "SELECT * FROM Cuentas "
  sSQL = sSQL & "WHERE '1' <= Codigo AND Codigo < '4' "
  sSQL = sSQL & "AND DG = 'D' "
  sSQL = sSQL & "ORDER BY Codigo "
  SelectData DataCtas, sSQL, False
  With DataCtas.Recordset
   If .RecordCount > 0 Then
       TotalBarra = .RecordCount
       Contador = 0
      .MoveFirst
       Do While Not .EOF
          DiarioESI.Caption = "ESTADO DE SITUACION INICIAL: '" & .Fields("Codigo") & "' "
          DataAsientos.Recordset.AddNew
          DataAsientos.Recordset.Fields("CODIGO") = .Fields("Codigo")
          DataAsientos.Recordset.Fields("CUENTA") = .Fields("Cuenta")
          DataAsientos.Recordset.Fields("PARCIAL_ME") = .Fields("S_Parcial_ME")
          DataAsientos.Recordset.Fields("DEBE") = .Fields("S_Debe")
          DataAsientos.Recordset.Fields("HABER") = .Fields("S_Haber")
          DataAsientos.Recordset.Update
          BoxPrint.Show 0
          Contador = Contador + 1
         .MoveNext
       Loop
   End If
  End With
  DBGAsientos.Visible = True
  DiarioESI.Caption = "ESTADO DE SITUACION INICIAL"
  RatonNormal MDIConta
  DBGAsientos.SetFocus
End Sub

Private Sub Form_Load()
   CentrarForm DiarioESI
   'Abriendo bases relacionadas
   DataCtas.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
   DataAsientos.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
   DataCatalogo.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
End Sub

