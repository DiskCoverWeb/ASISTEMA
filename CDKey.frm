VERSION 5.00
Begin VB.Form FCDKey 
   BorderStyle     =   0  'None
   Caption         =   "CDKey"
   ClientHeight    =   270
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2445
   Icon            =   "CDKey.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   270
   ScaleWidth      =   2445
   ShowInTaskbar   =   0   'False
End
Attribute VB_Name = "FCDKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Dim NumFile As Integer
Dim RutaGeneraFile As String
Dim ValorBool As String
Dim MiRegistro As Record   ' Declara una variable.
RutaGeneraFile = Left(CurDir$, 2) & "\ASISTEMA\CDKey.BIN"
NumFile = FreeFile


Open RutaGeneraFile For Random As #NumFile Len = Len(MiRegistro)
MiRegistro.Id = 1
MiRegistro.Nombre = "Walter Vaca Prieto"
Print #NumFile, MiRegistro;
Print #NumFile, MiRegistro.Nombre
MiRegistro.Id = 2
MiRegistro.Nombre = "XXXXXXXXXXXXXXXXXX"
Print #NumFile, MiRegistro.Id
Print #NumFile, MiRegistro.Nombre
' Cierra antes de volver a abrir en otro modo.
Close #NumFile


''Open RutaGeneraFile For Output As #NumFile ' Abre el archivo.
''Print #NumFile, "Walter Vaca Prieto"
''Print #NumFile, "XXXXXXXXXXXXXXXXXX"
''Close #NumFile
  End
End Sub
