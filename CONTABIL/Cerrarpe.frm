VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form CerrarPeriodo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CERRAR PERIODO"
   ClientHeight    =   5010
   ClientLeft      =   -300
   ClientTop       =   -15
   ClientWidth     =   5685
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5010
   ScaleWidth      =   5685
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "CANCELAR"
      Height          =   540
      Left            =   3150
      TabIndex        =   15
      Top             =   2415
      Width           =   2325
   End
   Begin VB.TextBox TextFechaAF 
      Height          =   330
      Left            =   5145
      MaxLength       =   2
      TabIndex        =   8
      Top             =   735
      Width           =   435
   End
   Begin VB.TextBox TextFechaMF 
      Height          =   330
      Left            =   4725
      MaxLength       =   2
      TabIndex        =   7
      Text            =   "12"
      Top             =   735
      Width           =   435
   End
   Begin VB.TextBox TextFechaDF 
      Height          =   330
      Left            =   4305
      MaxLength       =   2
      TabIndex        =   6
      Text            =   "31"
      Top             =   735
      Width           =   435
   End
   Begin VB.TextBox TextFechaA 
      Alignment       =   2  'Center
      Height          =   330
      Left            =   3885
      MaxLength       =   2
      TabIndex        =   5
      Top             =   735
      Width           =   435
   End
   Begin VB.TextBox TextFechaM 
      Alignment       =   2  'Center
      Height          =   330
      Left            =   3465
      MaxLength       =   2
      TabIndex        =   4
      Text            =   "01"
      Top             =   735
      Width           =   435
   End
   Begin VB.TextBox TextFechaD 
      Alignment       =   2  'Center
      Height          =   330
      Left            =   3045
      MaxLength       =   2
      TabIndex        =   3
      Text            =   "01"
      Top             =   735
      Width           =   435
   End
   Begin VB.FileListBox FileSistema 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2970
      Left            =   105
      TabIndex        =   12
      Top             =   525
      Width           =   2535
   End
   Begin VB.DirListBox DirSistema 
      Height          =   540
      Left            =   3045
      TabIndex        =   11
      Top             =   3045
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "PROCESAR CIERRE"
      Height          =   540
      Left            =   3150
      TabIndex        =   9
      Top             =   1470
      Width           =   2325
   End
   Begin ComctlLib.ProgressBar ProgBarra 
      Height          =   330
      Left            =   105
      TabIndex        =   10
      Top             =   4515
      Width           =   5475
      _ExtentX        =   9657
      _ExtentY        =   582
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Data DataFechaBal 
      Caption         =   "FechaBal"
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
      Top             =   0
      Visible         =   0   'False
      Width           =   2685
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Copiando y Procesando:"
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
      Left            =   105
      TabIndex        =   16
      Top             =   3570
      Width           =   2220
   End
   Begin VB.Label LabelResult 
      Caption         =   "Espere un momento..."
      Height          =   645
      Left            =   2415
      TabIndex        =   14
      Top             =   3780
      Width           =   3165
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nuevo Periodo"
      Height          =   225
      Left            =   4305
      TabIndex        =   2
      Top             =   525
      Width           =   1275
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Periodo Final"
      Height          =   225
      Left            =   3045
      TabIndex        =   1
      Top             =   525
      Width           =   1275
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ARCHIVOS A PROCESAR"
      Height          =   225
      Left            =   105
      TabIndex        =   13
      Top             =   315
      Width           =   2535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "FECHAS DD/MM/AA"
      Height          =   225
      Left            =   3045
      TabIndex        =   0
      Top             =   315
      Width           =   2535
   End
End
Attribute VB_Name = "CerrarPeriodo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub TextFechaA_Change()
  If Len(TextFechaA.Text) >= TextFechaA.MaxLength Then TextFechaDF.SetFocus
End Sub

Private Sub TextFechaA_GotFocus()
   TextFechaA.Text = ""
End Sub

Private Sub TextFechaA_LostFocus()
  If TextFechaA.Text = "" Then TextFechaA.Text = Anio
End Sub

Private Sub TextFechaAF_GotFocus()
  TextFechaAF.Text = ""
End Sub

Private Sub TextFechaD_Change()
  If Len(TextFechaD.Text) >= TextFechaD.MaxLength Then TextFechaM.SetFocus
End Sub

Private Sub TextFechaD_GotFocus()
  TextFechaD.Text = ""
End Sub

Private Sub TextFechaD_LostFocus()
  If TextFechaD.Text = "" Then TextFechaD.Text = Dia
End Sub

Private Sub TextFechaDF_GotFocus()
  TextFechaDF.Text = ""
End Sub

Private Sub TextFechaM_Change()
  If Len(TextFechaM.Text) >= TextFechaM.MaxLength Then TextFechaA.SetFocus
End Sub

Private Sub TextFechaM_GotFocus()
  TextFechaM.Text = ""
End Sub

Private Sub TextFechaM_LostFocus()
  If TextFechaM.Text = "" Then TextFechaM.Text = Mes
End Sub

Private Sub TextFechaAF_Change()
  If Len(TextFechaAF.Text) >= TextFechaAF.MaxLength Then SendKeys "{TAB}", True
End Sub

Private Sub TextFechaAF_LostFocus()
  If TextFechaAF.Text = "" Then TextFechaAF.Text = Anio
End Sub

Private Sub TextFechaDF_Change()
  If Len(TextFechaDF.Text) >= TextFechaDF.MaxLength Then TextFechaMF.SetFocus
End Sub

Private Sub TextFechaDF_LostFocus()
  If TextFechaDF.Text = "" Then TextFechaDF.Text = Dia
End Sub

Private Sub TextFechaMF_Change()
  If Len(TextFechaMF.Text) >= TextFechaMF.MaxLength Then TextFechaAF.SetFocus
End Sub

Private Sub TextFechaMF_GotFocus()
  TextFechaMF.Text = ""
End Sub

Private Sub TextFechaMF_LostFocus()
  If TextFechaMF.Text = "" Then TextFechaMF.Text = Mes
End Sub

Private Sub Command1_Click()
Dim dbsOrigen As Database
Dim dbsDestino As Database
  ChDir RutaEmpresa
  FileSistema.Pattern = "*.*"
  sSQL = "SELECT * FROM FechaBalance "
  DataFechaBal.RecordSource = sSQL: DataFechaBal.Refresh
  MiFecha = DataFechaBal.Recordset.Fields("Fecha_Final")
  FechaTexto = DataFechaBal.Recordset.Fields("Fecha_Inicial")
  FechaTexto1 = DataFechaBal.Recordset.Fields("Fecha_Final")
  DataFechaBal.Database.Close
  SumaDebe = 0: SumaHaber = 0
  FechaIni = FormatoFecha(TextFechaD.Text, TextFechaM.Text, TextFechaA.Text)
  FechaFin = FormatoFecha(TextFechaDF.Text, TextFechaMF.Text, TextFechaAF.Text)
  If (MiFecha = FechaIni) Then
     Mensajes = "Esta seguro que desea cerrar EL PERIODO: " & Chr(13)
     Mensajes = Mensajes & FechaIni & " e iniciar un nuevo periodo."
     Titulo = "CIERRE DEL PERIODO"
     TipoDeCaja = 4 + 32: ResultBox = MsgBox(Mensajes, TipoDeCaja, Titulo)
     If (ResultBox = 6) And (FechaTexto <> FechaTexto1) Then
        MousePointer = vbHourglass
        RutaOrigen = UCase(RutaEmpresa & "\")
        RutaDestino = UCase(RutaEmpresa & "1\")
        Cadena = RutaOrigen & Chr(13) & RutaDestino
        LabelResult.Caption = Cadena
        ProgBarra.Min = 0: ProgBarra.Max = FileSistema.ListCount - 1
        TotalBarra = 0
        For I = 0 To FileSistema.ListCount - 1
            RutaOrigen = UCase(RutaEmpresa & "\" & FileSistema.List(I))
            RutaDestino = UCase(RutaEmpresa & "1\" & FileSistema.List(I))
            FileCopy RutaOrigen, RutaDestino
            ProgBarra.Value = TotalBarra: TotalBarra = TotalBarra + 1
        Next I
        FileSistema.Pattern = "*.MDB"
        ProgBarra.Min = 0: ProgBarra.Max = FileSistema.ListCount - 1
        TotalBarra = 0
        For I = 0 To FileSistema.ListCount - 1
            ProgBarra.Value = TotalBarra
            RutaOrigen = UCase(RutaEmpresa & "\" & FileSistema.List(I))
            RutaDestino = UCase(RutaEmpresa & "1\" & FileSistema.List(I))
            If UCase(FileSistema.List(I)) = "CONTABIL.MDB" Then
               'procesando actualizaciones del periodo a cerrar
               FechaFin = FormatoFecha(TextFechaDF.Text, TextFechaMF.Text, TextFechaAF.Text)
               sSQL = "SELECT * FROM FechaBalance "
               DataFechaBal.RecordSource = sSQL: DataFechaBal.Refresh
               sSQL = "UPDATE FechaBalance SET Fecha_Inicial = #" & FechaFin & "# "
               DataFechaBal.Database.Execute sSQL
               
               FechaIni = BuscarFecha(TextFechaD.Text, TextFechaM.Text, TextFechaA.Text)
               FechaFin = BuscarFecha(TextFechaDF.Text, TextFechaMF.Text, TextFechaAF.Text)
               
               sSQL = "DELETE * FROM Transacciones "
               sSQL = sSQL & "WHERE Fecha <= #" & FechaFin & "# "
               Set dbsOrigen = OpenDatabase(RutaOrigen)
               dbsOrigen.Execute sSQL: dbsOrigen.Close
               
               sSQL = "DELETE * FROM Ingresos "
               sSQL = sSQL & "WHERE Fecha <= #" & FechaFin & "# "
               Set dbsOrigen = OpenDatabase(RutaOrigen)
               dbsOrigen.Execute sSQL: dbsOrigen.Close
               
               sSQL = "DELETE * FROM Egresos "
               sSQL = sSQL & "WHERE Fecha <= #" & FechaFin & "# "
               Set dbsOrigen = OpenDatabase(RutaOrigen)
               dbsOrigen.Execute sSQL: dbsOrigen.Close
               
               sSQL = "DELETE * FROM Diario "
               sSQL = sSQL & "WHERE Fecha <= #" & FechaFin & "# "
               Set dbsOrigen = OpenDatabase(RutaOrigen)
               dbsOrigen.Execute sSQL: dbsOrigen.Close
               
               sSQL = "DELETE * FROM Detalle_Retencion "
               sSQL = sSQL & "WHERE Fecha <= #" & FechaFin & "# "
               Set dbsOrigen = OpenDatabase(RutaOrigen)
               dbsOrigen.Execute sSQL: dbsOrigen.Close
               
               'procesando actualizaciones del nuevo periodo
               sSQL = "DELETE * FROM Transacciones "
               sSQL = sSQL & "WHERE Fecha >= #" & FechaIni & "# "
               Set dbsDestino = OpenDatabase(RutaDestino)
               dbsDestino.Execute sSQL: dbsDestino.Close
               
               sSQL = "DELETE * FROM Ingresos "
               sSQL = sSQL & "WHERE Fecha >= #" & FechaIni & "# "
               Set dbsDestino = OpenDatabase(RutaDestino)
               dbsDestino.Execute sSQL: dbsDestino.Close
               
               sSQL = "DELETE * FROM Egresos "
               sSQL = sSQL & "WHERE Fecha >= #" & FechaIni & "# "
               Set dbsDestino = OpenDatabase(RutaDestino)
               dbsDestino.Execute sSQL: dbsDestino.Close
               
               sSQL = "DELETE * FROM Diario "
               sSQL = sSQL & "WHERE Fecha >= #" & FechaIni & "# "
               Set dbsDestino = OpenDatabase(RutaDestino)
               dbsDestino.Execute sSQL: dbsDestino.Close
               
               sSQL = "DELETE * FROM Detalle_Retencion "
               sSQL = sSQL & "WHERE Fecha >= #" & FechaIni & "# "
               Set dbsDestino = OpenDatabase(RutaDestino)
               dbsDestino.Execute sSQL: dbsDestino.Close
               
            ElseIf UCase(FileSistema.List(I)) = "PRODUCC.MDB" Then
               'todo lo relacionado a facturacion
            End If
            TotalBarra = TotalBarra + 1
        Next I
        
        MousePointer = vbDefault
     Else
         MsgBox "Ya se proceso el cierre del periodo."
     End If
  Else
     Mensaje = "WARNING: " & Chr(13)
     Mensaje = Mensaje & "Para Procesar el cierre debe procesar primero MAYORIZACION Y BALANCE DE COMPROBACION de la opcion CONTABILIDAD "
     Mensaje = Mensaje & "del menú principal, además las fecha de cierre debe ser igual a la fecha final de la mayorización, "
     Mensaje = Mensaje & "por lo tanto no se procesará el cierre."
     FMensaje.Show
  End If
End Sub

Private Sub Command2_Click()
  Unload CerrarPeriodo
End Sub

Private Sub Form_Load()
  CentrarForm CerrarPeriodo
  DataFechaBal.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
  Cadena = "Copiando y Procesando: " & Chr(13)
  Cadena = Cadena & "Desde: " & Chr(13)
  Cadena = Cadena & "a: "
  Label4.Caption = Cadena
  MDIConta.MousePointer = vbDefault
End Sub

