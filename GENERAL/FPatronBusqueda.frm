VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FPatronBusqueda 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "."
   ClientHeight    =   3210
   ClientLeft      =   0
   ClientTop       =   15
   ClientWidth     =   5985
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawStyle       =   5  'Transparent
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FPatronBusqueda.frx":0000
   ScaleHeight     =   3210
   ScaleWidth      =   5985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox CmbCampos 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   735
      TabIndex        =   1
      Top             =   315
      Width           =   4530
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   5250
      Picture         =   "FPatronBusqueda.frx":9A46
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   315
      Width           =   435
   End
   Begin MSDataListLib.DataCombo DCValorBusqueda 
      Bindings        =   "FPatronBusqueda.frx":9D50
      DataSource      =   "AdoBuscar"
      Height          =   1350
      Left            =   210
      TabIndex        =   2
      Top             =   1155
      Width           =   5475
      _ExtentX        =   9657
      _ExtentY        =   2381
      _Version        =   393216
      Style           =   1
      BackColor       =   16761024
      Text            =   ""
   End
   Begin MSAdodcLib.Adodc AdoBuscar 
      Height          =   330
      Left            =   210
      Top             =   2520
      Width           =   5475
      _ExtentX        =   9657
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
      Caption         =   "Buscar"
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
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF8080&
      Caption         =   "&X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   315
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   315
      Width           =   435
   End
End
Attribute VB_Name = "FPatronBusqueda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim HayID As Boolean
Dim HayItem As Boolean
Dim HayPeriodo As Boolean

Dim vTipoBusqueda As String
Dim vAliasTabla As String
Dim vTablaBusqueda As String
Dim vCampoBusqueda As String
Dim vValorBusqueda As String

Private Sub CmbCampos_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
  If KeyCode = vbKeyEscape Then
     SQLPatron = ""
     Unload FPatronBusqueda
  End If
End Sub

Private Sub CmbCampos_LostFocus()
  HayID = False
  HayItem = False
  HayPeriodo = False
  vAliasTabla = ""
  vTablaBusqueda = ""
  vCampoBusqueda = CmbCampos.Text
  If InStr("Cta, Cheq_Dep, Debe, Haber, Saldo, Parcial_ME, Saldo_ME, Detalle, Codigo_C, C_Costo", vCampoBusqueda) Then
     vAliasTabla = "T."
     vTablaBusqueda = "Transacciones"
     HayID = True
     HayItem = True
     HayPeriodo = True
  ElseIf InStr("TC, Fecha_V, Fecha_E, Factura, Debitos, Creditos, Prima, Comp_No, Autorizacion, Serie, Detalle_SubCta, Procesado", vCampoBusqueda) Then
     vAliasTabla = "TS."
     vTablaBusqueda = "Trans_SubCtas"
     HayID = True
     HayItem = True
     HayPeriodo = True
  ElseIf InStr("TP, Fecha, Numero, Codigo_B, Presupuesto, Concepto, Autorizado, CodigoU", vCampoBusqueda) Then
     vAliasTabla = "CO."
     vTablaBusqueda = "Comprobantes"
     HayID = True
     HayItem = True
     HayPeriodo = True
  ElseIf InStr("Clave, TC, DG, Codigo, Cuenta, Presupuesto, ME, TB, Codigo_Ext, Cta_Acreditar, Mod_Gastos, I_E_Emp, Con_IESS, Cod_Rol_Pago, Tipo_Pago", vCampoBusqueda) Then
     vAliasTabla = "CC."
     vTablaBusqueda = "Catalogo_Cuentas"
     HayID = True
     HayItem = True
     HayPeriodo = True
  ElseIf InStr("Codigo, Cliente, CI_RUC, Representante, Grupo, Actividad, Profesion, Email, Direccion, DirNumero, Telefono, " _
             & "TelefonoT, Celular, Ciudad, Prov, Pais, No_Dep, RISE, Especial, Email2, Contacto, Telefono_R, Cta_CxP, Tipo_Cta, Cod_Banco, " _
             & "Cta_Numero, Fecha_Cad, EmailR, Tipo_Cliente, Barrio, Canton, Parroquia, Referencia, Estado, Beneficiario", vCampoBusqueda) Then
     vAliasTabla = "CL."
     vTablaBusqueda = "Clientes"
     HayID = True
  End If
  If vTablaBusqueda <> "" And Len(vCampoBusqueda) > 1 Then
     Select Case vCampoBusqueda
       Case "Beneficiario"
            sSQL = "SELECT TOP 10 Cliente " _
                 & "FROM " & vTablaBusqueda & " " _
                 & "GROUP BY Cliente " _
                 & "ORDER BY Cliente "
            DCValorBusqueda.Text = ""
       Case Else
            sSQL = "SELECT TOP 10 " & vCampoBusqueda & " " _
                 & "FROM " & vTablaBusqueda & " " _
                 & "GROUP BY " & vCampoBusqueda & " " _
                 & "ORDER BY " & vCampoBusqueda & " "
            DCValorBusqueda.Text = ""
     End Select
     Select_Adodc AdoBuscar, sSQL
     Select Case vCampoBusqueda
       Case "Beneficiario"
            DCValorBusqueda.ListField = "Cliente"
       Case Else
            DCValorBusqueda.ListField = vCampoBusqueda
     End Select
     DCValorBusqueda.Refresh
  End If
End Sub

Private Sub Command1_Click()
   SQLPatron = ""
   Unload FPatronBusqueda
End Sub

Private Sub Command2_Click()
  If vAliasTabla <> "" Then
     Select Case vCampoBusqueda
       Case "Beneficiario"
            vCampoBusqueda = vAliasTabla & "Cliente"
       Case Else
            vCampoBusqueda = vAliasTabla & vCampoBusqueda
     End Select
     If vTipoBusqueda <> "" And vValorBusqueda <> "" Then
        SQLPatron = "AND " & vCampoBusqueda & " = "
        Select Case vTipoBusqueda
          Case "N": SQLPatron = SQLPatron & vValorBusqueda & " "
          Case "D": SQLPatron = SQLPatron & "#" & BuscarFecha(vValorBusqueda) & "# "
          Case Else: SQLPatron = SQLPatron & "'" & vValorBusqueda & "' "
        End Select
     End If
  Else
     SQLPatron = ""
  End If
  Unload FPatronBusqueda
End Sub

Private Sub DCValorBusqueda_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
  If KeyCode = vbKeyEscape Then
     SQLPatron = ""
     Unload FPatronBusqueda
  End If
End Sub

Private Sub DCValorBusqueda_KeyPress(KeyAscii As Integer)
Dim IdB As Integer
  vValorBusqueda = DCValorBusqueda.Text
  If vTablaBusqueda <> "" And Len(vCampoBusqueda) > 1 Then
     Select Case vCampoBusqueda
       Case "Beneficiario"
            sSQL = "SELECT TOP 30 Cliente "
       Case Else
            sSQL = "SELECT TOP 30 " & vCampoBusqueda & " "
     End Select
     sSQL = sSQL _
          & "FROM " & vTablaBusqueda & " "
     For IdB = 0 To UBound(Campos_Patron_Busqueda) - 1
         If vCampoBusqueda = Campos_Patron_Busqueda(IdB).Campo Then
            vTipoBusqueda = Campos_Patron_Busqueda(IdB).TipoCampo
            Select Case vTipoBusqueda
              Case "N"
                   If vValorBusqueda = "" Then vValorBusqueda = "0"
                   sSQL = sSQL & "WHERE " & vCampoBusqueda & " = " & vValorBusqueda & " "
              Case "D"
                   If IsDate(vValorBusqueda) And Len(vValorBusqueda) >= 10 Then
                      sSQL = sSQL & "WHERE " & vCampoBusqueda & " = #" & BuscarFecha(vValorBusqueda) & "# "
                   Else
                      sSQL = sSQL & "WHERE " & vCampoBusqueda & " = #" & BuscarFecha("01/01/2000") & "# "
                   End If
              Case Else
                   If vValorBusqueda = "" Then vValorBusqueda = Ninguno
                   Select Case vCampoBusqueda
                     Case "Beneficiario"
                          sSQL = sSQL & "WHERE Cliente LIKE '%" & vValorBusqueda & "%' "
                     Case Else
                          sSQL = sSQL & "WHERE " & vCampoBusqueda & " LIKE '%" & vValorBusqueda & "%' "
                   End Select
            End Select
         End If
     Next IdB
     If HayItem Then sSQL = sSQL & "AND Item = '" & NumEmpresa & "' "
     If HayPeriodo Then sSQL = sSQL & "AND Periodo = '" & Periodo_Contable & "' "
     Select Case vCampoBusqueda
       Case "Beneficiario"
            sSQL = sSQL _
                 & "GROUP BY Cliente " _
                 & "ORDER BY Cliente "
       Case Else
            sSQL = sSQL _
                 & "GROUP BY " & vCampoBusqueda & " " _
                 & "ORDER BY " & vCampoBusqueda & " "
     End Select
     Select_Adodc AdoBuscar, sSQL
     'ElseIf InStr("Beneficiario", vCampoBusqueda) Then
  End If
'  Label1.Caption = vTipoBusqueda & vbCrLf & vTablaBusqueda & vbCrLf & vCampoBusqueda & vbCrLf & vValorBusqueda & vbCrLf & sSQL
'  Label1.Refresh
End Sub

Private Sub DCValorBusqueda_LostFocus()
   vValorBusqueda = DCValorBusqueda.Text
End Sub

Private Sub Form_Activate()
    CmbCampos.SetFocus
End Sub

'Empieza enviar correos
Private Sub Form_Load()
Dim IdB As Integer
    RatonReloj
    CentrarForm FPatronBusqueda
    Redondear_Formulario FPatronBusqueda, 40
    ConectarAdodc AdoBuscar
    SQLPatron = ""
    CmbCampos.Clear
    For IdB = 0 To UBound(Campos_Patron_Busqueda) - 1
        If MidStrg(Campos_Patron_Busqueda(IdB).Campo, 1, 5) <> "Fecha" Then CmbCampos.AddItem Campos_Patron_Busqueda(IdB).Campo
    Next IdB
    CmbCampos.Text = CmbCampos.List(0)
    RatonNormal
End Sub
