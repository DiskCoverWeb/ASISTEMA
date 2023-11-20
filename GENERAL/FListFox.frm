VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Begin VB.Form FListFox 
   Caption         =   "LISTADO DE ALUMNOS DEL SISTEMA EDUCATIVO"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin MSFlexGridLib.MSFlexGrid MSFGFOXP 
      Bindings        =   "FListFox.frx":0000
      Height          =   2430
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   4320
      _ExtentX        =   7620
      _ExtentY        =   4286
      _Version        =   393216
   End
   Begin VB.Data DtaFoxP 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   105
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2730
      Width           =   1140
   End
End
Attribute VB_Name = "FListFox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()

Dim I As Integer
Dim J As Integer
 
'TBeneficiario As Tipo_Beneficiarios
''' sSQL = "UPDATE " & Dato_DBF.Antiguos & " " _
'''      & "SET Bus = 'N' " _
'''      & "WHERE Bus <> 'N' "
''' ConectarDataExecute sSQL
 
 If Dato_DBF.Carpeta <> "." Then
    MsgBox Dato_DBF.Carpeta
    sSQL = "SELECT * " _
         & "FROM " & Dato_DBF.Antiguos & " " _
         & "WHERE NOMBRES <> '.' " _
         & "ORDER BY NOMBRES "
          
    Set Dato_DBF.Base_Datos = OpenDatabase(Dato_DBF.Carpeta, False, False, "FoxPro 3.0;")
    Set Dato_DBF.Registo = Dato_DBF.Base_Datos.OpenRecordset(sSQL)
    Set DtaFoxP.Recordset = Dato_DBF.Registo
    
    RatonReloj
    I = 0
    With DtaFoxP.Recordset
     If .RecordCount > 0 Then
        .MoveFirst
         Cadena = ""
         For J = 0 To .Fields.Count - 1
             Cadena = Cadena & .Fields(J).Name & vbTab
         Next J
     End If
    End With
    RatonNormal
  End If
End Sub

Private Sub Form_Load()
   MSFGFOXP.Height = MDI_Y_Max - MSFGFOXP.Top - 350
   MSFGFOXP.width = MDI_X_Max - MSFGFOXP.Left - 100
   DtaFoxP.width = MDI_X_Max - MSFGFOXP.Left - 100
   DtaFoxP.Top = MSFGFOXP.Top + MSFGFOXP.Height
 End Sub
