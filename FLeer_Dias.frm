VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Begin VB.Form FLeerDias 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "LEER DIAS DE RESPALDOS"
   ClientHeight    =   375
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   2700
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   375
   ScaleWidth      =   2700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc AdoRespaldo 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   2745
      _ExtentX        =   4842
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
      Caption         =   "Respaldo"
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
Attribute VB_Name = "FLeerDias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
    RatonReloj
    ConectarAdodc AdoRespaldo
   'Averiguamos los dias de Respaldo
    sSQL = "SELECT * " _
         & "FROM Tabla_Dias_Meses " _
         & "WHERE No_D_M > 0 " _
         & "AND Tipo = 'D' " _
         & "ORDER BY No_D_M "
    SelectAdodc AdoRespaldo, sSQL
    With AdoRespaldo.Recordset
     If .RecordCount > 0 Then
         Do While Not .EOF
            DiaZip(.Fields("No_D_M")) = .Fields("Zip")
            'MsgBox .Fields("No_D_M") & vbCrLf & DiaZip(.Fields("No_D_M"))
           .MoveNext
         Loop
     End If
    End With
   'Averiguamos la Carpeta de Respaldo
    sSQL = "SELECT Item,Empresa,Email_Respaldos " _
         & "FROM Empresas " _
         & "WHERE LEN(Email_Respaldos) > 1 " _
         & "ORDER BY Item,Empresa,Email_Respaldos "
    SelectAdodc AdoRespaldo, sSQL
    With AdoRespaldo.Recordset
     If .RecordCount > 0 Then
         Email_Respaldo = .Fields("Email_Respaldos")
         ReDim NombreDeEmpresas(.RecordCount) As String
         Contador = 0
         Do While Not .EOF
            NombreDeEmpresas(Contador) = .Fields("Empresa")
            Contador = Contador + 1
           .MoveNext
         Loop
         Contador = 0
     End If
    End With
    RatonNormal
    Unload FLeerDias
End Sub

Private Sub Form_Load()
   CentrarForm FLeerDias
End Sub
