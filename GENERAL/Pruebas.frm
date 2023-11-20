VERSION 5.00
Begin VB.Form Pruebas 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Pruebas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

'Variable de objeto para utilizar Microsoft Jet and Replication Objects
Dim obj_Jet As New JRO.JetEngine
 'Variables para las cadenas de conexión de la base de datos origen y destino
  Dim cFuente As String, cDestino As String

  'Cadena de conexión de la base de deatos que se va a compactar
  cFuente = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\SISTEMA\EMPRESA\DiskCoveOk.mdb"
  'cDestino = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\SISTEMA\EMPRESA\destino.mdb;Jet OLEDB:Engine Type=5"
  cDestino = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\\DISKCOVER-VAIO\Backup DiskCover\destino.mdb;Jet OLEDB:Engine Type=5"


  'Compacta la base de datos con el método CompactDatabase
  obj_Jet.CompactDatabase cFuente, cDestino
  
  RatonNormal
  'Compactar_Base = True

'Rutina de error
 
  'Compactar_Base = False
 If Err.Description = 0 Then
    MsgBox "Respaldo Exitoso"
 Else
    MsgBox Err.Description, vbExclamation
 End If
' FileCopy "C:\SISTEMA\EMPRESA\DiskCoveOk.mdb", "\\DISKCOVER-VAIO\Backup DiskCover\destino.mdb"
End Sub
