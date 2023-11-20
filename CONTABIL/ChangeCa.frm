VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FChangeCta 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CAMBIO DE VALORES DE LA CUENTA"
   ClientHeight    =   3570
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10095
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   10095
   ShowInTaskbar   =   0   'False
   Begin MSDataListLib.DataCombo DCCuenta 
      Bindings        =   "ChangeCa.frx":0000
      DataSource      =   "AdoCta"
      Height          =   2655
      Left            =   105
      TabIndex        =   3
      Top             =   735
      Width           =   8835
      _ExtentX        =   15584
      _ExtentY        =   4683
      _Version        =   393216
      Style           =   1
      BackColor       =   12648447
      Text            =   ""
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
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
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
      Left            =   9030
      Picture         =   "ChangeCa.frx":0015
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   105
      Width           =   960
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Salir"
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
      Left            =   9030
      Picture         =   "ChangeCa.frx":08DF
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1050
      Width           =   960
   End
   Begin MSAdodcLib.Adodc AdoCta 
      Height          =   330
      Left            =   210
      Top             =   840
      Width           =   2430
      _ExtentX        =   4286
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
      Caption         =   "Cta"
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
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ELIJA LA EMPRESA A COPIAR EL CATALOGO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   8835
   End
End
Attribute VB_Name = "FChangeCta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
 'MsgBox Producto
  Codigo2 = SinEspaciosIzq(DCCuenta.Text)
  If Codigo2 = "" Then Codigo2 = Ninguno
  If Codigo1 <> Ninguno Then
     Select Case Producto
       Case "Catalogo"
            Actualiza_Cuenta_Tabla "Transacciones", "Cta", Codigo1, Codigo2
            Actualiza_Cuenta_Tabla "Trans_SubCtas", "Cta", Codigo1, Codigo2
            Actualiza_Cuenta_Tabla "Trans_Compras", "Cta_Servicio", Codigo1, Codigo2
            Actualiza_Cuenta_Tabla "Trans_Compras", "Cta_Bienes", Codigo1, Codigo2
            Actualiza_Cuenta_Tabla "Trans_Air", "Cta_Retencion", Codigo1, Codigo2
            Actualiza_Cuenta_Tabla "Trans_Kardex", "Cta_Inv", Codigo1, Codigo2
            Actualiza_Cuenta_Tabla "Trans_Kardex", "Contra_Cta", Codigo1, Codigo2
            Actualiza_Cuenta_Tabla "Facturas", "Cta_CxP", Codigo1, Codigo2
            Actualiza_Cuenta_Tabla "Trans_Abonos", "Cta_CxP", Codigo1, Codigo2
            Actualiza_Cuenta_Tabla "Trans_Abonos", "Cta", Codigo1, Codigo2
            Actualiza_Cuenta_Tabla "Catalogo_CxCxP", "Cta", Codigo1, Codigo2
            Actualiza_Cuenta_Tabla "Catalogo_Lineas", "CxC", Codigo1, Codigo2
            Actualiza_Cuenta_Tabla "Catalogo_Lineas", "CxC_Anterior", Codigo1, Codigo2
       Case Else
            Actualiza_Cuenta_Tabla "Transacciones", "Cta", Codigo1, Codigo2, True, Asiento
            Actualiza_Cuenta_Tabla "Trans_SubCtas", "Cta", Codigo1, Codigo2, True
            Actualiza_Cuenta_Tabla "Trans_Compras", "Cta_Servicio", Codigo1, Codigo2, True
            Actualiza_Cuenta_Tabla "Trans_Compras", "Cta_Bienes", Codigo1, Codigo2, True
            Actualiza_Cuenta_Tabla "Trans_Air", "Cta_Retencion", Codigo1, Codigo2, True
            Actualiza_Cuenta_Tabla "Trans_Kardex", "Cta_Inv", Codigo1, Codigo2, True
            Actualiza_Cuenta_Tabla "Trans_Kardex", "Contra_Cta", Codigo1, Codigo2, True
     End Select
     Actualiza_Procesado_Tabla "Transacciones", , "Cta", Codigo1
     Actualiza_Procesado_Tabla "Transacciones", , "Cta", Codigo2
     Actualiza_Procesado_Tabla "Trans_SubCtas", , "Cta", Codigo1
     Actualiza_Procesado_Tabla "Trans_SubCtas", , "Cta", Codigo2
     Actualiza_Procesado_Tabla "Trans_Kardex", , "Cta_Inv", Codigo1
     Actualiza_Procesado_Tabla "Trans_Kardex", , "Cta_Inv", Codigo2
     Actualiza_Procesado_Tabla "Trans_Kardex", , "Contra_Cta", Codigo1
     Actualiza_Procesado_Tabla "Trans_Kardex", , "Contra_Cta", Codigo2
     MsgBox "Proceso terminado, vuelva a listar el comprobante"
  Else
     MsgBox "No se puede cambiar la Cuenta"
  End If
  Unload FChangeCta
End Sub

Private Sub Command2_Click()
  Unload FChangeCta
End Sub

Private Sub DCCuenta_KeyDown(KeyCode As Integer, Shift As Integer)
  'PresionoEnter KeyCode
End Sub

Private Sub DCCuenta_KeyPress(KeyAscii As Integer)
Dim Busqueda As String
    Busqueda = DCCuenta.Text
    If Len(Busqueda) > 1 Then
       sSQL = "SELECT TOP 50 Codigo & ' - ' & Cuenta As Ctas, Codigo, Cuenta " _
            & "FROM Catalogo_Cuentas " _
            & "WHERE Item = '" & NumEmpresa & "' " _
            & "AND Periodo = '" & Periodo_Contable & "' " _
            & "AND DG = 'D' " _
            & "AND Codigo <> '" & Codigo1 & "' "
       If IsNumeric(Replace(Busqueda, ".", "")) Then
          sSQL = sSQL & "AND Codigo LIKE '" & Busqueda & "%' "
       Else
          sSQL = sSQL & "AND Cuenta LIKE '%" & Busqueda & "%' "
       End If
       sSQL = sSQL & "ORDER BY Codigo "
       Select_Adodc AdoCta, sSQL
    End If
End Sub

Private Sub Form_Activate()
  If Codigo1 <> Ninguno Then
     Label6.Caption = Codigo3 & vbCrLf & "SELECCIONE LA CUENTA A CAMBIAR"
     sSQL = "SELECT TOP 50 Codigo & ' - ' & Cuenta As Ctas, Codigo, Cuenta " _
          & "FROM Catalogo_Cuentas " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "AND DG = 'D' " _
          & "AND Codigo <> '" & Codigo1 & "' " _
          & "ORDER BY Codigo "
     SelectDB_Combo DCCuenta, AdoCta, sSQL, "Ctas"
     If AdoCta.Recordset.RecordCount <= 0 Then
        Unload FChangeCta
     Else
        DCCuenta.SetFocus
     End If
  Else
     MsgBox "No existe Cuenta para cambiar"
     Unload FChangeCta
  End If
  RatonNormal
End Sub

Private Sub Form_Load()
  CentrarForm FChangeCta
  ConectarAdodc AdoCta
End Sub
