VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Begin VB.Form LCxCxP 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ingreso/Modificacion de SubCuentas"
   ClientHeight    =   4260
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6765
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   6765
   ShowInTaskbar   =   0   'False
   Begin MSDataListLib.DataList DLCtas 
      Bindings        =   "LCxCxP.frx":0000
      DataSource      =   "AdoSubCta"
      Height          =   2985
      Left            =   105
      TabIndex        =   5
      Top             =   1155
      Width           =   6525
      _ExtentX        =   11509
      _ExtentY        =   5265
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSAdodcLib.Adodc AdoSubCta 
      Height          =   330
      Left            =   210
      Top             =   1365
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
      Caption         =   "SubCta"
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
   Begin VB.Frame Frame1 
      Caption         =   "Tipo de Cuenta"
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
      Left            =   105
      TabIndex        =   0
      Top             =   0
      Width           =   5475
      Begin VB.OptionButton OpcP 
         Caption         =   "Modulo de Ctas x Pagar"
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
         Left            =   2730
         TabIndex        =   1
         Top             =   315
         Width           =   2535
      End
      Begin VB.OptionButton OpcC 
         Caption         =   "Modulo de Ctas x Cobrar"
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
         TabIndex        =   2
         Top             =   315
         Value           =   -1  'True
         Width           =   2535
      End
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
      Height          =   645
      Left            =   5670
      Picture         =   "LCxCxP.frx":0018
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   105
      Width           =   960
   End
   Begin MSAdodcLib.Adodc AdoSubCta1 
      Height          =   330
      Left            =   2205
      Top             =   1995
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
      Caption         =   "SubCta1"
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
      Caption         =   "SUBCUENTA DE BLOQUE"
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
      TabIndex        =   3
      Top             =   840
      Width           =   6525
   End
End
Attribute VB_Name = "LCxCxP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command2_Click()
  Unload LCxCxP
End Sub

Private Sub DLCtas_DblClick()
  SiguienteControl
End Sub

Private Sub DLCtas_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
  If KeyCode = vbKeyDelete Then
     Codigo = SinEspaciosIzq(DLCtas.Text)
     Cadena = SinEspaciosDer(DLCtas.Text)
     sSQL = "SELECT Codigo " _
          & "FROM Trans_SubCtas " _
          & "WHERE Codigo = '" & Cadena & "' " _
          & "AND Cta = '" & Codigo & "' " _
          & "AND Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' "
     Select_Adodc AdoSubCta1, sSQL
     If AdoSubCta1.Recordset.RecordCount > 0 Then
        Mensajes = "No se puede eliminar esta SubCuenta," & vbCrLf _
                 & "porque tiene cuentas procesables."
        MsgBox Mensajes
     Else
        Mensajes = "Esta seguro que desea eliminar la " & vbCrLf _
                 & "Cuenta No. [" & Cadena & "]"
        Titulo = "Pregunta de Eliminacion"
        If BoxMensaje = vbYes Then
           sSQL = "DELETE * " _
                & "FROM Catalogo_CxCxP " _
                & "WHERE Codigo = '" & Cadena & "' " _
                & "AND Cta = '" & Codigo & "' " _
                & "AND Item = '" & NumEmpresa & "' " _
                & "AND Periodo = '" & Periodo_Contable & "' "
           Ejecutar_SQL_SP sSQL
           If OpcC.value Then
              ListarSubCtas "C"
           Else
              ListarSubCtas "P"
           End If
        End If
     End If
     DLCtas.SetFocus
  End If
End Sub

Private Sub DLCtas_LostFocus()
  Cadena = SinEspaciosIzq(DLCtas.Text)
  Codigo1 = SinEspaciosIzq(DLCtas.Text)
End Sub

Private Sub Form_Activate()
  ListarSubCtas "C"
  RatonNormal
End Sub

Private Sub Form_Load()
  CentrarForm LCxCxP
  ConectarAdodc AdoSubCta
  ConectarAdodc AdoSubCta1
End Sub

Private Sub OpcC_Click()
  ListarSubCtas "C"
End Sub

Private Sub OpcP_Click()
  ListarSubCtas "P"
End Sub

Public Sub ListarSubCtas(TipoCta As String)
  sSQL = "SELECT CP.Cta & ' => ' & Cl.Cliente & Space(3) & CP.Codigo As Nombre_Cta " _
       & "FROM Catalogo_CxCxP As CP,Clientes As Cl " _
       & "WHERE CP.TC = '" & TipoCta & "' " _
       & "AND CP.Codigo = Cl.Codigo " _
       & "AND CP.Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "ORDER BY CP.Cta,Cl.Cliente "
  SelectDB_List DLCtas, AdoSubCta, sSQL, "Nombre_Cta"
  DLCtas.SetFocus
End Sub
