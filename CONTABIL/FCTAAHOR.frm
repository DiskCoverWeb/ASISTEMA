VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Begin VB.Form FCtaAhorro 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Adulto"
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8340
   Icon            =   "FCTAAHOR.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   8340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSDataListLib.DataCombo DCTipoLibreta 
      Bindings        =   "FCTAAHOR.frx":014A
      DataSource      =   "AdoTipoLibreta"
      Height          =   345
      Left            =   105
      TabIndex        =   1
      Top             =   525
      Width           =   5475
      _ExtentX        =   9657
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
   Begin VB.CheckBox CheqLib 
      Caption         =   "Actualizacion Libreta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   105
      TabIndex        =   14
      Top             =   2415
      Width           =   2325
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
      Height          =   750
      Left            =   7350
      Picture         =   "FCTAAHOR.frx":0167
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   105
      Width           =   855
   End
   Begin VB.TextBox TxtArea 
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
      Left            =   5145
      MaxLength       =   5
      TabIndex        =   9
      Top             =   945
      Width           =   1065
   End
   Begin VB.TextBox TxtSector 
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
      Left            =   3570
      MaxLength       =   20
      TabIndex        =   11
      Top             =   1260
      Width           =   2640
   End
   Begin VB.TextBox TxtNoSoc 
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
      Left            =   3570
      MaxLength       =   2
      MultiLine       =   -1  'True
      TabIndex        =   7
      Text            =   "FCTAAHOR.frx":0A31
      Top             =   945
      Width           =   540
   End
   Begin MSAdodcLib.Adodc AdoCxCxP 
      Height          =   330
      Left            =   4620
      Top             =   2415
      Visible         =   0   'False
      Width           =   1800
      _ExtentX        =   3175
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
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   7350
      Picture         =   "FCTAAHOR.frx":0A33
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   945
      Width           =   855
   End
   Begin MSAdodcLib.Adodc AdoTipoLibreta 
      Height          =   330
      Left            =   2835
      Top             =   2415
      Visible         =   0   'False
      Width           =   1800
      _ExtentX        =   3175
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
      Caption         =   "TipoLibreta"
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
   Begin MSAdodcLib.Adodc AdoClientes 
      Height          =   330
      Left            =   6405
      Top             =   2415
      Visible         =   0   'False
      Width           =   1800
      _ExtentX        =   3175
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
   Begin MSDataListLib.DataCombo DCClientes 
      Bindings        =   "FCTAAHOR.frx":12FD
      DataSource      =   "AdoClientes"
      Height          =   345
      Left            =   105
      TabIndex        =   13
      Top             =   1995
      Width           =   7155
      _ExtentX        =   12621
      _ExtentY        =   609
      _Version        =   393216
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
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " BENEFICIARIO"
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
      TabIndex        =   12
      Top             =   1680
      Width           =   7155
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " AREA"
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
      Left            =   4095
      TabIndex        =   8
      Top             =   945
      Width           =   1065
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " SECTOR"
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
      Left            =   2310
      TabIndex        =   10
      Top             =   1260
      Width           =   1275
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " No. SOCIOS"
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
      Left            =   2310
      TabIndex        =   6
      Top             =   945
      Width           =   1275
   End
   Begin VB.Label LblCtaNo 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " XXXXXXXXXX"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   105
      TabIndex        =   5
      Top             =   1260
      Width           =   2115
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " CODIGO"
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
      Left            =   5670
      TabIndex        =   2
      Top             =   105
      Width           =   1590
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " CUENTA No."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   105
      TabIndex        =   4
      Top             =   945
      Width           =   2115
   End
   Begin VB.Label LblCliente 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ELIJA LA CUENTA DE ASIGNACION"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   330
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   5475
   End
   Begin VB.Label LblCodigo 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "XXXXXXXXXX"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   330
      Left            =   5670
      TabIndex        =   3
      Top             =   525
      Width           =   1590
   End
End
Attribute VB_Name = "FCtaAhorro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()

  TextoValido TxtNoSoc, True, True
  TextoValido TxtArea, , True
  TextoValido TxtSector, , True
  CodigoCli = LblCodigo.Caption
  Cuenta_No = LblCtaNo.Caption
    
 'MsgBox NumEmpresa
  With AdoClientes.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Cliente = '" & DCClientes & "' ")
       If Not .EOF Then
          CodigoB = .fields("Codigo")
       End If
   End If
  End With
  
  sSQL = "SELECT * " _
       & "FROM Clientes_Datos_Extras " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Tipo_Dato = 'LIBRETAS' " _
       & "AND Cuenta_No = '" & Cuenta_No & "' " _
       & "AND Codigo = '" & CodigoCli & "' "
  Select_Adodc AdoCxCxP, sSQL
  If AdoCxCxP.Recordset.RecordCount <= 0 Then
     SetAddNew AdoCxCxP
     SetFields AdoCxCxP, "Item", NumEmpresa
     SetFields AdoCxCxP, "Tipo_Dato", "LIBRETAS"
     SetFields AdoCxCxP, "Codigo", CodigoCli
     SetFields AdoCxCxP, "Num", Val(TxtNoSoc.Text)
     SetFields AdoCxCxP, "Cuenta_No", Cuenta_No
     SetFields AdoCxCxP, "Area", TxtArea.Text
     SetFields AdoCxCxP, "Ciudad", TxtSector.Text
     SetFields AdoCxCxP, "Fecha_Registro", FechaSistema
     SetFields AdoCxCxP, "Fecha_B", FechaSistema
     SetFields AdoCxCxP, "CodigoU", CodigoUsuario
     SetFields AdoCxCxP, "CodigoA", Ninguno
     SetFields AdoCxCxP, "Tipo", TipoCta
     SetFields AdoCxCxP, "CodigoB", CodigoB
     SetFields AdoCxCxP, "Acreditacion", TipoDoc
     If TipoDoc = "A" Then
        FechaStr = CLongFecha(CFechaLong(FechaSistema) + 365)
        SetFields AdoCxCxP, "Fecha_Ret", FechaStr
     End If
     If TipoDoc = "2A" Then
        FechaStr = CLongFecha(CFechaLong(FechaSistema) + 730)
        SetFields AdoCxCxP, "Fecha_Ret", FechaStr
     End If
     If CheqLib.value = 1 Then
        SetFields AdoCxCxP, "T", Normal
     Else
        SetFields AdoCxCxP, "T", Pendiente
     End If
     SetFields AdoCxCxP, "ME", False
     SetUpdate AdoCxCxP
  End If
'  Mensajes = "Imprimir"
 'Imprimir_Apertura LblCtaNo.Caption
  Unload FCtaAhorro
End Sub

Private Sub Command2_Click()
  Unload FCtaAhorro
End Sub

Private Sub DCTipoLibreta_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCTipoLibreta_LostFocus()
  TipoCta = Ninguno
  TipoDoc = Ninguno
  With AdoTipoLibreta.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Tipo = '" & DCTipoLibreta.Text & "' ")
       If Not .EOF Then
          TipoCta = .fields("Tipo")
          TipoDoc = .fields("Acreditacion")
          Tipo_Cuenta .fields("Tipo_Cta")
       End If
   End If
  End With
  
End Sub

Private Sub Form_Activate()
  LblCodigo.Caption = CodigoCli
  LblCliente.Caption = NombreCliente
  
  sSQL = "SELECT * " _
       & "FROM Clientes " _
       & "WHERE Cliente <> '.' " _
       & "ORDER BY Cliente "
  SelectDB_Combo DCClientes, AdoClientes, sSQL, "Cliente"
  
  sSQL = "SELECT Tipo,Acreditacion,Tipo_Cta " _
       & "FROM Catalogo_Interes " _
       & "WHERE TP = 'C' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "GROUP BY Tipo,Acreditacion,Tipo_Cta " _
       & "ORDER BY Tipo "
  SelectDB_Combo DCTipoLibreta, AdoTipoLibreta, sSQL, "Tipo"
  
End Sub

Private Sub Form_Load()
  CentrarForm FCtaAhorro
  ConectarAdodc AdoCxCxP
  ConectarAdodc AdoClientes
  ConectarAdodc AdoTipoLibreta
  FCtaAhorro.Caption = "ASIGNACION DE CUENTA DE AHORRO"
End Sub

Private Sub TxtArea_GotFocus()
  MarcarTexto TxtArea
End Sub

Private Sub TxtArea_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtArea_LostFocus()
  TextoValido TxtArea, , True
End Sub

Private Sub TxtNoSoc_GotFocus()
  MarcarTexto TxtNoSoc
End Sub

Private Sub TxtNoSoc_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtNoSoc_LostFocus()
  TextoValido TxtNoSoc, True, True
End Sub

Private Sub TxtSector_GotFocus()
  MarcarTexto TxtSector
End Sub

Private Sub TxtSector_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtSector_LostFocus()
  TextoValido TxtSector, , True
End Sub

Public Sub Tipo_Cuenta(Tipo_Cta As String)
  sSQL = "SELECT * " _
       & "FROM Clientes_Datos_Extras " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "ORDER BY Cuenta_No "
  Select_Adodc AdoCxCxP, sSQL
  If AdoCxCxP.Recordset.RecordCount > 0 Then
     AdoCxCxP.Recordset.MoveLast
     Cuenta_No = MidStrg(AdoCxCxP.Recordset.fields("Cuenta_No"), 1, 8)
  Else
     Cuenta_No = NumEmpresa & "00000"
  End If
  Cuenta_No = Format(CLng(Cuenta_No + 1), "00000000")
  Cuenta_No = Cuenta_No & "-" & Tipo_Cta
  LblCtaNo.Caption = Cuenta_No
End Sub
