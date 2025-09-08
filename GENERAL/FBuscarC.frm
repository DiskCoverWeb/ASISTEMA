VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Begin VB.Form FBuscarClientes 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "BUSCAR BENEFICIARIO: <Enter> Acepta, <Esc> Cancela"
   ClientHeight    =   4500
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7065
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   7065
   ShowInTaskbar   =   0   'False
   Begin MSDataListLib.DataCombo DCClientes 
      Bindings        =   "FBuscarC.frx":0000
      DataSource      =   "AdoClientes"
      Height          =   4470
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "<Enter> Acepta seleccion del Beneficiario, <Esc> Salir sin seleccionar Beneficiario"
      Top             =   0
      Width           =   7050
      _ExtentX        =   12435
      _ExtentY        =   7885
      _Version        =   393216
      Style           =   1
      BackColor       =   16761024
      ForeColor       =   16777215
      Text            =   "Buscar Clientes"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSAdodcLib.Adodc AdoClientes 
      Height          =   330
      Left            =   105
      Top             =   420
      Visible         =   0   'False
      Width           =   2220
      _ExtentX        =   3916
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
      BackColor       =   16744576
      ForeColor       =   16777215
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
      Caption         =   "Clientes"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
End
Attribute VB_Name = "FBuscarClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub DCClientes_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyEscape Then
      TBeneficiario.Patron_Busqueda = Ninguno
      TBeneficiario = Leer_Datos_Cliente_SP(TBeneficiario.Patron_Busqueda)
      Unload FBuscarClientes
   End If
   If KeyCode = vbKeyReturn Then
      TBeneficiario.Patron_Busqueda = DCClientes.Text
      If Len(TBeneficiario.Patron_Busqueda) <= 0 Then TBeneficiario.Patron_Busqueda = Ninguno
      TBeneficiario = Leer_Datos_Cliente_SP(TBeneficiario.Patron_Busqueda)
      Unload FBuscarClientes
   End If
End Sub

Private Sub Form_Activate()
  If Len(TBeneficiario.Patron_Busqueda) > 0 Then
     sSQL = "SELECT Cliente, CI_RUC, Codigo " _
          & "FROM Clientes "
     If IsNumeric(TBeneficiario.Patron_Busqueda) Then
        sSQL = sSQL & "WHERE CI_RUC LIKE '%" & TBeneficiario.Patron_Busqueda & "%' "
     Else
        sSQL = sSQL & "WHERE Cliente LIKE '%" & TBeneficiario.Patron_Busqueda & "%' "
     End If
     sSQL = sSQL & "ORDER BY Cliente "
     SelectDB_Combo DCClientes, AdoClientes, sSQL, "Cliente"
     DCClientes.SetFocus
  Else
     MsgBox "No hay datos que buscar"
     Unload FBuscarClientes
  End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  PresionoEnter KeyCode
  If KeyCode = vbKeyEscape Then
     TBeneficiario.Codigo = Ninguno
     Unload FBuscarClientes
  End If
End Sub

Private Sub Form_Load()
  CentrarForm FBuscarClientes
  ConectarAdodc AdoClientes
End Sub
