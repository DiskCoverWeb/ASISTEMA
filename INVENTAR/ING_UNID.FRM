VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form Ing_Unid 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ingreso de Unidades de Medidas"
   ClientHeight    =   2370
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4065
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   4065
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "SALIR"
      Height          =   420
      Left            =   2565
      TabIndex        =   1
      Top             =   945
      Width           =   1230
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "Ing_Unid.frx":0000
      Height          =   1905
      Left            =   135
      OleObjectBlob   =   "Ing_Unid.frx":0015
      TabIndex        =   0
      Top             =   270
      Width           =   2175
   End
   Begin VB.Data DataUnidad 
      Caption         =   "Unidad"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   135
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   2175
   End
End
Attribute VB_Name = "Ing_Unid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
  Unload Ing_Unid
End Sub

Private Sub DBGrid1_BeforeDelete(Cancel As Integer)
    Mensajes = "¿Realmente desea eliminar La Unidad " & Chr(13)
    Mensajes = Mensajes & DataUnidad.Recordset.Fields("Unidad") & "?"
    Titulo = "Confirmación de eliminación"
    TipoDeCaja = 4 + 32: J = MsgBox(Mensajes, TipoDeCaja, Titulo)
    If J = 6 Then Cancel = False Else Cancel = True
End Sub

Private Sub Form_Activate()
   SSQL = "SELECT * FROM Unidades ORDER BY Unidad "
   DataUnidad.RecordSource = SSQL: DataUnidad.Refresh
End Sub

Private Sub Form_Load()
   'Centrar el formulario
   Left = (Screen.Width - Width) / 2
   Top = (Screen.Height - Height) / 2 - 600
   DataUnidad.DatabaseName = RutaEmpresa & "\KARDEX.MDB"
End Sub
