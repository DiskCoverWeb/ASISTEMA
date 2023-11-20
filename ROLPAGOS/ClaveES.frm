VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Begin VB.Form EntradasSalidas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ENTRADA Y SALIDA DE EMPLEADOS"
   ClientHeight    =   1275
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7155
   Icon            =   "ClaveES.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1275
   ScaleWidth      =   7155
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TxtTarea 
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
      IMEMode         =   3  'DISABLE
      Left            =   945
      MaxLength       =   50
      MultiLine       =   -1  'True
      TabIndex        =   7
      Top             =   840
      Width           =   6105
   End
   Begin MSDataListLib.DataCombo DCProceso 
      Bindings        =   "ClaveES.frx":0696
      DataSource      =   "AdoProceso"
      Height          =   345
      Left            =   3465
      TabIndex        =   5
      Top             =   420
      Width           =   3585
      _ExtentX        =   6324
      _ExtentY        =   609
      _Version        =   393216
      Text            =   "DataCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox TxtClave 
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
      IMEMode         =   3  'DISABLE
      Left            =   1785
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   420
      Width           =   1695
   End
   Begin VB.TextBox TxtUsuario 
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
      TabIndex        =   1
      Top             =   420
      Width           =   1695
   End
   Begin MSAdodcLib.Adodc AdoCxCxP 
      Height          =   330
      Left            =   3150
      Top             =   840
      Visible         =   0   'False
      Width           =   2010
      _ExtentX        =   3545
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
      Caption         =   "CxCxP"
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
   Begin MSAdodcLib.Adodc AdoProceso 
      Height          =   330
      Left            =   1155
      Top             =   840
      Visible         =   0   'False
      Width           =   2010
      _ExtentX        =   3545
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
      Caption         =   "Proceso"
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
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " PROCESO:"
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
      Left            =   3465
      TabIndex        =   4
      Top             =   105
      Width           =   3585
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " CLAVE:"
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
      Left            =   1785
      TabIndex        =   2
      Top             =   105
      Width           =   1695
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " TAREA:"
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
      TabIndex        =   6
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " USUARIO:"
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
      Width           =   1695
   End
End
Attribute VB_Name = "EntradasSalidas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub DCProceso_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
  If KeyCode = vbKeyEscape Then Unload EntradasSalidas
End Sub

Private Sub Form_Activate()
  sSQL = "SELECT Proceso " _
       & "FROM Trans_Entrada_Salida " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Proceso <> '.' " _
       & "GROUP BY Proceso " _
       & "ORDER BY Proceso "
  SelectDB_Combo DCProceso, AdoProceso, sSQL, "Proceso"

  sSQL = "SELECT * " _
       & "FROM Catalogo_Rol_Pagos " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "ORDER BY Usuario,Clave "
  Select_Adodc AdoCxCxP, sSQL
End Sub

Private Sub Form_Load()
  CentrarForm EntradasSalidas
  ConectarAdodc AdoCxCxP
  ConectarAdodc AdoProceso
  EntradasSalidas.Caption = "REGISTRO DE ENTRADA/SALIDA"
End Sub

Private Sub TxtUsuario_GotFocus()
 TxtUsuario = ""
End Sub

Private Sub TxtUsuario_LostFocus()
 CodigoCli = Ninguno
 TextoValido TxtUsuario, , True
 If TxtUsuario <> Ninguno Then
    If AdoCxCxP.Recordset.RecordCount > 0 Then
       AdoCxCxP.Recordset.MoveFirst
       AdoCxCxP.Recordset.Find ("Usuario = '" & TxtUsuario & "' ")
       If AdoCxCxP.Recordset.EOF Then
          MsgBox "Usuario No asignado"
          TxtUsuario.SetFocus
       End If
    End If
 End If
End Sub

Private Sub TxtUsuario_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
  If KeyCode = vbKeyEscape Then Unload EntradasSalidas
End Sub

Private Sub TxtClave_GotFocus()
  MarcarTexto TxtClave
End Sub

Private Sub TxtClave_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
  If KeyCode = vbKeyEscape Then Unload EntradasSalidas
End Sub

Private Sub TxtClave_LostFocus()
 CodigoCli = Ninguno
 TextoValido TxtClave, , True
 If TxtClave.Text <> Ninguno Then
    If AdoCxCxP.Recordset.RecordCount > 0 Then
       AdoCxCxP.Recordset.MoveFirst
       AdoCxCxP.Recordset.Find ("Clave = '" & TxtClave.Text & "' ")
       If Not AdoCxCxP.Recordset.EOF Then
          If TxtUsuario.Text = AdoCxCxP.Recordset.Fields("Usuario") Then
             CodigoCli = AdoCxCxP.Recordset.Fields("Codigo")
          Else
             MsgBox "La clave no corresponde"
             TxtClave.SetFocus
          End If
       Else
          MsgBox "Clave incorrecta"
          TxtClave.SetFocus
       End If
    End If
 End If
End Sub

Private Sub TxtTarea_GotFocus()
  MarcarTexto TxtTarea
End Sub

Private Sub TxtTarea_LostFocus()
  TextoValido TxtTarea, , True
  If CodigoCli <> Ninguno And TxtTarea.Text <> Ninguno Then
     MiTiempo = Time
     SetAdoAddNew "Trans_Entrada_Salida"
     SetAdoFields "ES", "H"
     SetAdoFields "Codigo", CodigoCli
     SetAdoFields "Hora", Format(MiTiempo, FormatoTimes)
     SetAdoFields "Fecha", FechaSistema
     SetAdoFields "Proceso", UCase(MidStrg(DCProceso.Text, 1, 30))
     SetAdoFields "Tarea", TxtTarea
     SetAdoFields "CodigoU", CodigoUsuario
     SetAdoFields "Item", NumEmpresa
     SetAdoFields "Periodo", Periodo_Contable
     SetAdoUpdate
     MsgBox "Datos Registrados Correctamente" & vbCrLf & "Hora: " & Format(MiTiempo, FormatoTimes)
     Unload EntradasSalidas
  Else
     MsgBox "Datos incompletos, No se registro proceso"
     TxtUsuario.SetFocus
  End If
End Sub

Private Sub TxtTarea_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
  If KeyCode = vbKeyEscape Then Unload EntradasSalidas
End Sub

