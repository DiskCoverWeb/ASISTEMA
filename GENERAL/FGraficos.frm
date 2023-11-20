VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form FGraficos 
   BackColor       =   &H00FFFFFF&
   Caption         =   "GRAFICOS DE RESULTADOS"
   ClientHeight    =   5880
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   9330
   Icon            =   "FGraficos.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5880
   ScaleWidth      =   9330
   WindowState     =   2  'Maximized
   Begin MSChart20Lib.MSChart MSChart1 
      Height          =   5580
      Left            =   105
      OleObjectBlob   =   "FGraficos.frx":030A
      TabIndex        =   0
      Top             =   0
      Width           =   9150
   End
   Begin VB.Menu mhola 
      Caption         =   "hola"
   End
End
Attribute VB_Name = "FGraficos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Resize()
  With MSChart1
      .Left = 0
      .Top = 0
      .Width = Me.ScaleWidth
      .Height = Me.ScaleHeight
  End With
End Sub

Private Sub MSChart1_DblClick()
Dim intCol As Integer
Dim intFila As Integer
  With MSChart1
      .chartType = VtChChartType3dStep
      .ColumnCount = 12
      .RowCount = 1
      .TitleText = "ESTADOS DE RESULTADOS EN BARRAS"
      .FootnoteText = "RESULTADOS CONTRA PRESUPUESTOS"
       For intCol = 1 To 12
           For intFila = 1 To 12
              .Column = intCol
              .Row = 1
              .Data = Rnd(intCol) * 12
               Cadena = Mid(MesesLetras(intCol), 1, 3)
              .ColumnLabel = Cadena
           Next intFila
       Next intCol
      .ShowLegend = True
  End With
End Sub

