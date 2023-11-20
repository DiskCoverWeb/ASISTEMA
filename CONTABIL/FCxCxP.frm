VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FCxCxP 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CUENTAS POR COBRAR / CUENTAS POR PAGAR"
   ClientHeight    =   5070
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12195
   Icon            =   "FCxCxP.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   12195
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox CheqPorcentajes 
      Caption         =   "PORCENTAJES DE RETENCION"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   105
      TabIndex        =   10
      Top             =   4410
      Width           =   2010
   End
   Begin VB.Frame FrmPorcentaje 
      Height          =   645
      Left            =   2205
      TabIndex        =   11
      Top             =   4305
      Width           =   8940
      Begin VB.TextBox TxtRetIVAS 
         Height          =   330
         Left            =   8085
         TabIndex        =   17
         Text            =   "0"
         Top             =   210
         Width           =   750
      End
      Begin VB.TextBox TxtRetIVAB 
         Height          =   330
         Left            =   5040
         TabIndex        =   15
         Text            =   "0"
         Top             =   210
         Width           =   750
      End
      Begin VB.TextBox TxtCodRet 
         Height          =   330
         Left            =   1890
         TabIndex        =   13
         Top             =   210
         Width           =   855
      End
      Begin VB.Label Label5 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Retencion IVA Servicio"
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
         Left            =   5880
         TabIndex        =   16
         Top             =   210
         Width           =   2220
      End
      Begin VB.Label Label4 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Retencion IVA Bienes"
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
         Left            =   2835
         TabIndex        =   14
         Top             =   210
         Width           =   2220
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Codigo Retencion"
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
         Top             =   210
         Width           =   1800
      End
   End
   Begin VB.CheckBox CheqGasto 
      Caption         =   "ASIGNAR A LA CUENTA DE GASTO"
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
      TabIndex        =   4
      Top             =   2730
      Width           =   5790
   End
   Begin MSDataListLib.DataList DLCxCxP 
      Bindings        =   "FCxCxP.frx":014A
      DataSource      =   "AdoCxCxP"
      Height          =   1815
      Left            =   105
      TabIndex        =   3
      Top             =   735
      Width           =   11040
      _ExtentX        =   19473
      _ExtentY        =   3201
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Consolas"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSAdodcLib.Adodc AdoCxCxP 
      Height          =   330
      Left            =   210
      Top             =   945
      Visible         =   0   'False
      Width           =   1905
      _ExtentX        =   3360
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
      Left            =   11235
      Picture         =   "FCxCxP.frx":0161
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   945
      Width           =   855
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
      Left            =   11235
      Picture         =   "FCxCxP.frx":0A2B
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   105
      Width           =   855
   End
   Begin MSDataListLib.DataList DLGasto 
      Bindings        =   "FCxCxP.frx":12F5
      DataSource      =   "AdoGasto"
      Height          =   1230
      Left            =   105
      TabIndex        =   5
      Top             =   3045
      Visible         =   0   'False
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   2170
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Consolas"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSAdodcLib.Adodc AdoGasto 
      Height          =   330
      Left            =   210
      Top             =   1260
      Visible         =   0   'False
      Width           =   1905
      _ExtentX        =   3360
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
      Caption         =   "Gasto"
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
   Begin MSAdodcLib.Adodc AdoSubModulo 
      Height          =   330
      Left            =   210
      Top             =   1575
      Visible         =   0   'False
      Width           =   1905
      _ExtentX        =   3360
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
      Caption         =   "SubModulo"
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
   Begin MSDataListLib.DataList DLSubModulo 
      Bindings        =   "FCxCxP.frx":130C
      DataSource      =   "AdoSubModulo"
      Height          =   840
      Left            =   6090
      TabIndex        =   9
      Top             =   3360
      Visible         =   0   'False
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   1482
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Consolas"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSAdodcLib.Adodc AdoAux 
      Height          =   330
      Left            =   210
      Top             =   1890
      Visible         =   0   'False
      Width           =   1905
      _ExtentX        =   3360
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
      Caption         =   "Aux"
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
   Begin MSMask.MaskEdBox MBoxCta 
      Height          =   330
      Left            =   9345
      TabIndex        =   7
      Top             =   2625
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Asignar Cuenta del IVA al Gasto"
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
      Left            =   6090
      TabIndex        =   6
      Top             =   2625
      Width           =   3270
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " CUENTA DE SUBMODULO"
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
      Left            =   6090
      TabIndex        =   8
      Top             =   3045
      Visible         =   0   'False
      Width           =   5055
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " SE ASIGNARA A:"
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
      TabIndex        =   2
      Top             =   420
      Width           =   11040
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
      Left            =   1575
      TabIndex        =   1
      Top             =   105
      Width           =   9570
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
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   1485
   End
End
Attribute VB_Name = "FCxCxP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim CtaGastoCosto As String
Dim SubmoduloGastoCosto As String

Private Sub CheqGasto_Click()
    If CheqGasto.value <> 0 Then
       DLGasto.Visible = True
       If AdoSubModulo.Recordset.RecordCount > 0 Then
          Label1.Visible = True
          DLSubModulo.Visible = True
       Else
          Label1.Visible = False
          DLSubModulo.Visible = False
       End If
       DLGasto.SetFocus
    Else
       DLGasto.Visible = False
       Label1.Visible = False
       DLSubModulo.Visible = False
    End If
End Sub

Private Sub CheqPorcentajes_Click()
  If CheqPorcentajes.value = 0 Then FrmPorcentaje.Visible = False Else FrmPorcentaje.Visible = True
End Sub

Private Sub Command1_Click()
  Cta_Aux = SinEspaciosIzq(DLCxCxP.Text)
  SubmoduloGastoCosto = Ninguno
  If AdoSubModulo.Recordset.RecordCount > 0 And DLSubModulo.Visible Then
     AdoSubModulo.Recordset.MoveFirst
     AdoSubModulo.Recordset.Find ("Detalle = '" & DLSubModulo.Text & "' ")
     If Not AdoSubModulo.Recordset.EOF Then SubmoduloGastoCosto = AdoSubModulo.Recordset.fields("Codigo")
  End If
  
  If CheqGasto.value Then CtaGastoCosto = SinEspaciosIzq(DLGasto.Text) Else Cta1 = Ninguno
  sSQL = "SELECT * " _
       & "FROM Catalogo_CxCxP " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND Cta = '" & Cta_Aux & "' " _
       & "AND Codigo = '" & CodigoCliente & "' " _
       & "AND TC = '" & SubCta & "' "
  Select_Adodc AdoCxCxP, sSQL
  If AdoCxCxP.Recordset.RecordCount <= 0 Then
     SetAddNew AdoCxCxP
     SetFields AdoCxCxP, "Item", NumEmpresa
     SetFields AdoCxCxP, "Periodo", Periodo_Contable
     SetFields AdoCxCxP, "Codigo", CodigoCliente
     SetFields AdoCxCxP, "Cta", Cta_Aux
     SetFields AdoCxCxP, "TC", SubCta
     SetFields AdoCxCxP, "Importaciones", 0
  End If
  If CheqGasto.value <> 0 Then
     SetFields AdoCxCxP, "Cta_Gasto", CtaGastoCosto
     SetFields AdoCxCxP, "SubModulo", SubmoduloGastoCosto
  End If
  If CheqPorcentajes.value <> 0 Then
     SetFields AdoCxCxP, "Porc_IVAB", Val(TxtRetIVAB) / 100
     SetFields AdoCxCxP, "Porc_IVAS", Val(TxtRetIVAS) / 100
     SetFields AdoCxCxP, "Cod_Ret", TxtCodRet
  End If
  
  Codigo = CambioCodigoCta(MBoxCta.Text)
  If Len(Codigo) > 1 Then SetFields AdoCxCxP, "Cta_IVA_Gasto", Codigo
  SetUpdate AdoCxCxP
  AXML.Cta_Credito = Cta_Aux
  Unload FCxCxP
End Sub

Private Sub Command2_Click()
  Unload FCxCxP
End Sub

Private Sub DLCxCxP_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DLCxCxP_LostFocus()
    Cta_Aux = SinEspaciosIzq(DLCxCxP.Text)
    Llenar_Cta_CxP Cta_Aux
End Sub

Private Sub DLGasto_LostFocus()
  Label1.Visible = False
  DLSubModulo.Visible = False
  With AdoGasto.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Nombre_Cta = '" & DLGasto & "' ")
       If Not .EOF Then
          sSQL = "SELECT Detalle, Codigo " _
               & "FROM Catalogo_SubCtas " _
               & "WHERE Item = '" & NumEmpresa & "' " _
               & "AND Periodo = '" & Periodo_Contable & "' " _
               & "AND TC = '" & .fields("TC") & "' " _
               & "ORDER BY Detalle "
          SelectDB_List DLSubModulo, AdoSubModulo, sSQL, "Detalle"
          If AdoSubModulo.Recordset.RecordCount > 0 Then
             Label1.Visible = True
             DLSubModulo.Visible = True
             DLSubModulo.SetFocus
          End If
       End If
   End If
  End With
End Sub


Private Sub Form_Activate()
  LblCodigo.Caption = CodigoCliente
  LblCliente.Caption = NombreCliente
  CheqPorcentajes.value = 0
  FrmPorcentaje.Visible = False
  FormatoMaskCta MBoxCta
  Llenar_Cta_CxP
End Sub

Private Sub Form_Load()
  CentrarForm FCxCxP
  ConectarAdodc AdoAux
  ConectarAdodc AdoCxCxP
  ConectarAdodc AdoGasto
  ConectarAdodc AdoSubModulo
  
  sSQL = "SELECT Codigo & Space(19-LEN(Codigo)) & Cuenta As Nombre_Cta, TC, Codigo " _
       & "FROM Catalogo_Cuentas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND DG = 'D' "
  If SubCta = "C" Then sSQL = sSQL & "AND Codigo LIKE '4%' " Else sSQL = sSQL & "AND Codigo LIKE '5%' "
  sSQL = sSQL & "ORDER BY Codigo "
  SelectDB_List DLGasto, AdoGasto, sSQL, "Nombre_Cta"
  
  sSQL = "SELECT Detalle, Codigo " _
       & "FROM Catalogo_SubCtas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' "
  If SubCta = "C" Then sSQL = sSQL & "AND TC IN ('CC','I') " Else sSQL = sSQL & "AND TC IN ('CC','G') "
  sSQL = sSQL & "ORDER BY Detalle "
  SelectDB_List DLSubModulo, AdoSubModulo, sSQL, "Detalle"
  
  sSQL = "SELECT Codigo & Space(19-LEN(Codigo)) & Cuenta As Nombre_Cta, Codigo " _
       & "FROM Catalogo_Cuentas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND DG = 'D' " _
       & "AND TC = '" & SubCta & "' " _
       & "ORDER BY Codigo "
  SelectDB_List DLCxCxP, AdoCxCxP, sSQL, "Nombre_Cta"
  
  If SubCta = "C" Then
     FCxCxP.Caption = "ASIGNACION DE CUENTAS POR COBRAR"
     CheqGasto.Caption = "ASIGNAR A LA CUENTA DE INGRESO"
     DLGasto.Enabled = False
  Else
     FCxCxP.Caption = "ASIGNACION DE CUENTAS POR PAGAR"
     CheqGasto.Caption = "ASIGNAR A LA CUENTA DE GASTO"
     DLGasto.Enabled = True
  End If
End Sub

Private Sub MBoxCta_GotFocus()
    MarcarTexto MBoxCta
End Sub

Private Sub MBoxCta_KeyDown(KeyCode As Integer, Shift As Integer)
    PresionoEnter KeyCode
End Sub

Private Sub TxtCodRet_GotFocus()
    MarcarTexto TxtCodRet
End Sub

Private Sub TxtCodRet_KeyDown(KeyCode As Integer, Shift As Integer)
    PresionoEnter KeyCode
End Sub

Private Sub TxtCodRet_LostFocus()
    TextoValido TxtCodRet, , True
End Sub

Private Sub TxtRetIVAB_GotFocus()
    MarcarTexto TxtRetIVAB
End Sub

Private Sub TxtRetIVAB_KeyDown(KeyCode As Integer, Shift As Integer)
    PresionoEnter KeyCode
End Sub

Private Sub TxtRetIVAB_LostFocus()
    TextoValido TxtRetIVAB, True, , 2
End Sub

Private Sub TxtRetIVAS_GotFocus()
    MarcarTexto TxtRetIVAS
End Sub

Private Sub TxtRetIVAS_KeyDown(KeyCode As Integer, Shift As Integer)
    PresionoEnter KeyCode
End Sub

Private Sub TxtRetIVAS_LostFocus()
    TextoValido TxtRetIVAS, True, , 2
End Sub

Private Sub Llenar_Cta_CxP(Optional CtaCxP As String)
    sSQL = "SELECT " & Full_Fields("Catalogo_CxCxP") & " " _
         & "FROM Catalogo_CxCxP " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND Codigo = '" & CodigoCliente & "' " _
         & "AND TC = '" & SubCta & "' "
    If Len(CtaCxP) > 1 Then sSQL = sSQL & "AND Cta = '" & CtaCxP & "' "
    Select_Adodc AdoAux, sSQL
    With AdoAux.Recordset
     If .RecordCount > 0 Then
         CtaGastoCosto = .fields("Cta_Gasto")
         SubmoduloGastoCosto = .fields("SubModulo")
         TxtCodRet = .fields("Cod_Ret")
         TxtRetIVAB = .fields("Porc_IVAB") * 100
         TxtRetIVAS = .fields("Porc_IVAS") * 100
         MBoxCta.Text = FormatoCodigoCta(.fields("Cta_IVA_Gasto"))
         If Len(TxtCodRet) > 1 And (Val(TxtRetIVAB) + Val(TxtRetIVAS)) > 0 Then
            CheqPorcentajes.value = 1
            FrmPorcentaje.Visible = True
         Else
            CheqPorcentajes.value = 0
            FrmPorcentaje.Visible = False
         End If
         
         If AdoCxCxP.Recordset.RecordCount > 0 And CtaCxP = "" Then DLCxCxP.Text = AdoCxCxP.Recordset.fields("Nombre_Cta")
         
         If AdoGasto.Recordset.RecordCount > 0 Then
            AdoGasto.Recordset.MoveFirst
            AdoGasto.Recordset.Find ("Codigo = '" & CtaGastoCosto & "' ")
            If Not AdoGasto.Recordset.EOF Then
               DLGasto.Text = AdoGasto.Recordset.fields("Nombre_Cta")
               CheqGasto.value = 1
               DLGasto.Visible = True
            Else
               CheqGasto.value = 0
               DLGasto.Visible = False
            End If
         End If
         If AdoSubModulo.Recordset.RecordCount > 0 Then
            AdoSubModulo.Recordset.MoveFirst
            AdoSubModulo.Recordset.Find ("Codigo = '" & SubmoduloGastoCosto & "' ")
            If Not AdoSubModulo.Recordset.EOF Then
               DLSubModulo.Text = AdoSubModulo.Recordset.fields("Detalle")
               Label1.Visible = True
               DLSubModulo.Visible = True
            Else
               Label1.Visible = False
               DLSubModulo.Visible = False
            End If
         End If
     End If
    End With
End Sub

