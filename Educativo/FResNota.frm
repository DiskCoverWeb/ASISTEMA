VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSChrt20.ocx"
Begin VB.Form FResumenNotas 
   Caption         =   "RESUMEN DE NOTAS"
   ClientHeight    =   6765
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11805
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6765
   ScaleWidth      =   11805
   WindowState     =   2  'Maximized
   Begin MSChart20Lib.MSChart MSChrNotas 
      Height          =   3270
      Left            =   105
      OleObjectBlob   =   "FResNota.frx":0000
      TabIndex        =   6
      Top             =   735
      Width           =   11565
   End
   Begin MSDataGridLib.DataGrid DGResumen 
      Bindings        =   "FResNota.frx":2356
      Height          =   2535
      Left            =   105
      TabIndex        =   0
      Top             =   4095
      Width           =   11565
      _ExtentX        =   20399
      _ExtentY        =   4471
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
   Begin VB.Frame Frame1 
      Height          =   645
      Left            =   105
      TabIndex        =   1
      Top             =   0
      Width           =   11565
      Begin MSDataListLib.DataList DLCurso 
         Bindings        =   "FResNota.frx":236F
         DataSource      =   "AdoAux"
         Height          =   300
         Left            =   0
         TabIndex        =   5
         Top             =   210
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   529
         _Version        =   393216
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
      Begin VB.OptionButton OpcT 
         Caption         =   "Los Dos Quimestres"
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
         Left            =   9345
         TabIndex        =   4
         Top             =   210
         Width           =   2115
      End
      Begin VB.OptionButton OpcSQ 
         Caption         =   "Segundo Quimestre"
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
         Left            =   7245
         TabIndex        =   3
         Top             =   210
         Width           =   2010
      End
      Begin VB.OptionButton OpcPQ 
         Caption         =   "Primer Quimestre"
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
         TabIndex        =   2
         Top             =   210
         Width           =   1800
      End
   End
   Begin MSAdodcLib.Adodc AdoResumen 
      Height          =   330
      Left            =   315
      Top             =   1575
      Visible         =   0   'False
      Width           =   2115
      _ExtentX        =   3731
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
      Caption         =   "Resumen"
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
      Top             =   1260
      Visible         =   0   'False
      Width           =   2115
      _ExtentX        =   3731
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
Attribute VB_Name = "FResumenNotas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
  DGResumen.Height = FResumenNotas.Height - 1700
  DGResumen.width = FResumenNotas.width - 300
  MSChrNotas.width = FResumenNotas.width - 300
  sSQL = "SELECT (CodigoE & ' - ' & Detalle) As Cursos " _
       & "FROM Catalogo_Estudiantil " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND TC = 'P' " _
       & "ORDER BY CodigoE "
  SelectDBList DLCurso, AdoAux, sSQL, "Cursos"
  RatonNormal
End Sub

Private Sub Form_Load()
  ConectarAdodc AdoAux
  ConectarAdodc AdoResumen
End Sub

Public Sub Listar_Quimestres()
Dim Column, Row As Integer
Dim Index1, Index2, Index3, Index4 As Integer
  RatonReloj
  sSQL = "SELECT CE.CodigoE,CE.Detalle,"
  If OpcPQ.value Then
     sSQL = sSQL & "TN.PQBim1,COUNT(TN.PQBim1) As Cant_1er_QBim1," _
          & "TN.PQBim2,COUNT(TN.PQBim2) As Cant_1er_QBim2," _
          & "TN.PromPQ,SUM(TN.PromPQ) As Cant_Prom_1erQ," _
          & "TN.PromFinal,SUM(TN.PromFinal) As Cant_Prom_Final "
  ElseIf OpcSQ.value Then
     sSQL = sSQL & "TN.SQBim1,COUNT(TN.SQBim1) As Cant_2do_QBim1," _
          & "TN.SQBim2,COUNT(TN.SQBim2) As Cant_2do_QBim2," _
          & "TN.PromSQ,SUM(TN.PromSQ) As Cant_Prom_2doQ," _
          & "TN.PromFinal,SUM(TN.PromFinal) As Cant_Prom_Final "
  Else
     sSQL = sSQL & "TN.PQBim1,COUNT(TN.PQBim1) As Cant_1er_QBim1," _
          & "TN.PQBim2,COUNT(TN.PQBim2) As Cant_1er_QBim2," _
          & "TN.SQBim1,COUNT(TN.SQBim1) As Cant_2do_QBim1," _
          & "TN.SQBim2,COUNT(TN.SQBim2) As Cant_2do_QBim2," _
          & "TN.PromPQ,SUM(TN.PromPQ) As Cant_Prom_1erQ," _
          & "TN.PromSQ,SUM(TN.PromSQ) As Cant_Prom_2doQ," _
          & "TN.PromFinal,SUM(TN.PromFinal) As Cant_Prom_Final "
  End If
  sSQL = sSQL & "FROM Trans_Notas As TN,Catalogo_Estudiantil As CE,Clientes As C " _
       & "WHERE TN.Item = '" & NumEmpresa & "' " _
       & "AND C.Codigo = TN.Codigo " _
       & "AND C.Grupo = CE.CodigoE " _
       & "AND TN.Item = CE.Item " _
       & "AND CE.CodigoE = '" & SinEspaciosIzq(DLCurso.Text) & "' "
  If OpcPQ.value Then
     sSQL = sSQL & "GROUP BY CE.CodigoE,CE.Detalle,TN.PQBim1,TN.PQBim2,TN.PromPQ,TN.PromFinal " _
          & "ORDER BY CE.CodigoE,CE.Detalle,TN.PQBim1,TN.PQBim2 "
  ElseIf OpcSQ.value Then
     sSQL = sSQL & "GROUP BY CE.CodigoE,CE.Detalle,TN.SQBim1,TN.SQBim2,TN.PromSQ,TN.PromFinal " _
          & "ORDER BY CE.CodigoE,CE.Detalle,TN.SQBim1,TN.SQBim2 "
  Else
     sSQL = sSQL & "GROUP BY CE.CodigoE,CE.Detalle,TN.PQBim1,TN.PQBim2,TN.SQBim1,TN.SQBim2,TN.PromPQ,TN.PromSQ,TN.PromFinal " _
       & "ORDER BY CE.CodigoE,CE.Detalle,TN.PQBim1,TN.PQBim2,TN.SQBim1,TN.SQBim2 "
  End If
  SelectDataGrid DGResumen, AdoResumen, sSQL, , True
  If OpcPQ.value Or OpcSQ.value Then
     MSChrNotas.chartType = VtChChartType2dBar   'VtChChartType3dBar
     MSChrNotas.ColumnCount = 4
     MSChrNotas.RowCount = 20
     MSChrNotas.Column = 1: MSChrNotas.ColumnLabel = "Bimestre 1"
     MSChrNotas.Column = 2: MSChrNotas.ColumnLabel = "Bimestre 2"
     MSChrNotas.Column = 3: MSChrNotas.ColumnLabel = "Promedios"
     MSChrNotas.Column = 4: MSChrNotas.ColumnLabel = "Prom. Final"
     For Row = 1 To 20
         MSChrNotas.Row = Row: MSChrNotas.RowLabel = "N." & Format(Row, "00")
     Next Row
     With AdoResumen.Recordset
      If .RecordCount > 0 Then
          Do While Not .EOF
             If OpcPQ.value Then
                If .Fields("PQBim1") > 0 Then
                    'MsgBox .Fields("PQBim1")
                    MSChrNotas.Row = .Fields("PQBim1"): MSChrNotas.Column = 1
                    MSChrNotas.Data = .Fields("Cant_1er_QBim1")
                End If
                If .Fields("PQBim2") > 0 Then
                    MSChrNotas.Row = .Fields("PQBim2"): MSChrNotas.Column = 2
                    MSChrNotas.Data = .Fields("Cant_1er_QBim2")
                End If
                If .Fields("PromPQ") > 0 Then
                    MSChrNotas.Row = .Fields("PromPQ"): MSChrNotas.Column = 3
                    MSChrNotas.Data = .Fields("Cant_Prom_1erQ")
                End If
             End If
             If OpcSQ.value Then
             
             End If
            .MoveNext
          Loop
      End If
     End With
'''     For Column = 1 To 4
'''         For Row = 1 To 20
'''            MSChrNotas.Row = Row
'''            MSChrNotas.Column = Column
'''            MSChrNotas.Data = Int((20 - 0 + 1) * Rnd + 0)
'''         Next Row
'''      Next Column
      ' Utiliza el gráfico como fondo de la leyenda.
      MSChrNotas.ShowLegend = True
      MSChrNotas.SelectPart VtChPartTypePlot, Index1, Index2, Index3, Index4
      MSChrNotas.EditCopy
      MSChrNotas.SelectPart VtChPartTypeLegend, Index1, Index2, Index3, Index4
      MSChrNotas.EditPaste
  End If
  RatonNormal
End Sub

Private Sub OpcPQ_DblClick()
  Listar_Quimestres
End Sub

Private Sub OpcSQ_DblClick()
  Listar_Quimestres
End Sub

Private Sub OpcT_DblClick()
  Listar_Quimestres
End Sub
