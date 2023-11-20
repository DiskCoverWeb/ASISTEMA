VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FBenefic 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ingreso/Modificacion de SubCuentas"
   ClientHeight    =   5085
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   7575
   ShowInTaskbar   =   0   'False
   Begin MSDataListLib.DataList DLCtas 
      Bindings        =   "Benefic.frx":0000
      DataSource      =   "AdoSubCta"
      Height          =   1815
      Left            =   2160
      TabIndex        =   23
      Top             =   720
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   3201
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc AdoSubCta1 
      Height          =   330
      Left            =   3960
      Top             =   2160
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
   Begin MSAdodcLib.Adodc AdoSubCta 
      Height          =   330
      Left            =   4080
      Top             =   1800
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   7
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Eliminar"
            Object.ToolTipText     =   "Elimina una Cuenta Contable"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Nuevo"
            Object.ToolTipText     =   "Nueva Cuenta Contable"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Grabar"
            Object.ToolTipText     =   "Grabar Cuenta"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Primero"
            Object.ToolTipText     =   "Primera Cuenta"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Anterior"
            Object.ToolTipText     =   "Anterior Cuenta"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Siguiente"
            Object.ToolTipText     =   "Siguiente Cuenta"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Ultimo"
            Object.ToolTipText     =   "Ultima Cuenta"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Frame Frame1 
      Height          =   960
      Left            =   105
      TabIndex        =   20
      Top             =   420
      Width           =   1905
      Begin VB.OptionButton OpcR 
         Caption         =   "Responsable"
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
         Left            =   210
         TabIndex        =   22
         Top             =   525
         Width           =   1485
      End
      Begin VB.OptionButton OpcB 
         Caption         =   "Proveedor"
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
         Left            =   210
         TabIndex        =   21
         Top             =   210
         Value           =   -1  'True
         Width           =   1380
      End
   End
   Begin VB.TextBox TextDirecc 
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
      Left            =   1680
      MaxLength       =   60
      TabIndex        =   16
      Top             =   4620
      Width           =   4425
   End
   Begin VB.TextBox TextCiudad 
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
      MaxLength       =   15
      TabIndex        =   15
      Top             =   4620
      Width           =   1485
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Grabar"
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
      Left            =   6300
      Picture         =   "Benefic.frx":0018
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   735
      Width           =   1065
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
      Left            =   6300
      Picture         =   "Benefic.frx":045A
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   1680
      Width           =   1065
   End
   Begin VB.TextBox TextSubCta 
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
      MaxLength       =   35
      TabIndex        =   3
      Text            =   "0"
      Top             =   3150
      Width           =   4635
   End
   Begin MSMask.MaskEdBox MBoxRUC 
      Height          =   330
      Left            =   105
      TabIndex        =   9
      Top             =   3885
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   582
      _Version        =   393216
      AllowPrompt     =   -1  'True
      AutoTab         =   -1  'True
      MaxLength       =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#########-#-###"
      Mask            =   "#########-#-###"
      PromptChar      =   "0"
   End
   Begin MSMask.MaskEdBox MBoxTelefono 
      Height          =   330
      Left            =   2100
      TabIndex        =   10
      Top             =   3885
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   582
      _Version        =   393216
      AllowPrompt     =   -1  'True
      AutoTab         =   -1  'True
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "##-###-###"
      Mask            =   "##-###-###"
      PromptChar      =   "0"
   End
   Begin MSMask.MaskEdBox MBoxCelular 
      Height          =   330
      Left            =   3465
      TabIndex        =   11
      Top             =   3885
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   582
      _Version        =   393216
      AllowPrompt     =   -1  'True
      AutoTab         =   -1  'True
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "##-###-###"
      Mask            =   "##-###-###"
      PromptChar      =   "0"
   End
   Begin MSMask.MaskEdBox MBoxFAX 
      Height          =   330
      Left            =   4830
      TabIndex        =   12
      Top             =   3885
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   582
      _Version        =   393216
      AllowPrompt     =   -1  'True
      AutoTab         =   -1  'True
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "##-###-###"
      Mask            =   "##-###-###"
      PromptChar      =   "0"
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "NOMBRE DE PROVEEDORES"
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
      Left            =   2100
      TabIndex        =   0
      Top             =   525
      Width           =   4005
   End
   Begin VB.Label Label12 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " CIUDAD:"
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
      TabIndex        =   13
      Top             =   4305
      Width           =   1485
   End
   Begin VB.Label Label11 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " R.U.C. / C.I."
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
      TabIndex        =   5
      Top             =   3570
      Width           =   1905
   End
   Begin VB.Label Label9 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " TELEFONO:"
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
      Left            =   2100
      TabIndex        =   6
      Top             =   3570
      Width           =   1275
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " CELULAR:"
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
      TabIndex        =   7
      Top             =   3570
      Width           =   1275
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " FAX:"
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
      Left            =   4830
      TabIndex        =   8
      Top             =   3570
      Width           =   1275
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " DIRECCION:"
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
      Left            =   1680
      TabIndex        =   14
      Top             =   4305
      Width           =   4425
   End
   Begin VB.Label LabelCodigo 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0000"
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
      Left            =   4830
      TabIndex        =   4
      Top             =   3150
      Width           =   1275
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "BENEFICIARIO"
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
      Top             =   2835
      Width           =   4635
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Codigo:"
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
      Left            =   4830
      TabIndex        =   2
      Top             =   2835
      Width           =   1275
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   7
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Benefic.frx":06DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Benefic.frx":07EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Benefic.frx":0900
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Benefic.frx":0A12
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Benefic.frx":0F24
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Benefic.frx":1436
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Benefic.frx":1948
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FBenefic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
  GrabarCta LabelCodigo.Caption
  AdoSubCtas.Refresh
End Sub

Private Sub Command2_Click()
  Unload FBenefic
End Sub

Private Sub DLCtas_DblClick()
  SiguienteControl
End Sub

Private Sub DLCtas_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DLCtas_LostFocus()
  Cadena = SinEspaciosIzq(DLCtas.Text)
  LlenarCta Cadena
End Sub

Private Sub DLCtas_Click()

End Sub

Private Sub Form_Activate()
  sSQL = "SELECT Codigo & Space(5) & Beneficiario As Nombre_Cta "
  sSQL = sSQL & "FROM Beneficiarios "
  sSQL = sSQL & "WHERE TC = 'P' "
  sSQL = sSQL & "ORDER BY Codigo "
  SelectDBList DLCtas, AdoSubCtas, sSQL, "Nombre_Cta"
  DLCtas.SetFocus
  RatonNormal
End Sub

Private Sub Form_Load()
  CentrarForm FBenefic
  ConectarAdodc AdoSubCtas
  ConectarAdodc AdoSubCta1
End Sub

Private Sub OpcB_Click()
  sSQL = "SELECT Codigo & Space(5) & Beneficiario As Nombre_Cta "
  sSQL = sSQL & "FROM Beneficiarios "
  sSQL = sSQL & "WHERE TC = 'P' "
  sSQL = sSQL & "ORDER BY Codigo "
  SelectDBList DLCtas, AdoSubCtas, sSQL, "Nombre_Cta"
End Sub

Private Sub OpcR_Click()
  sSQL = "SELECT Codigo & Space(5) & Beneficiario As Nombre_Cta "
  sSQL = sSQL & "FROM Beneficiarios "
  sSQL = sSQL & "WHERE TC = 'R' "
  sSQL = sSQL & "ORDER BY Codigo "
  SelectDBList DLCtas, AdoSubCtas, sSQL, "Nombre_Cta"
End Sub

Private Sub TextSubCta_LostFocus()
  If TextSubCta.Text = "" Then TextSubCta.Text = Ninguno
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
 With AdoSubCtas.Recordset
 Select Case Button.key
   Case "Eliminar"
      If DLCtas.Enabled Then
         Cadena = SinEspaciosIzq(DLCtas.Text)
         sSQL = "SELECT Codigo_P FROM Kardex "
         sSQL = sSQL & " WHERE  Codigo_P = '" & Cadena & "' "
         SelectData AdoSubCta1, sSQL, False
         If AdoSubCta1.Recordset.RecordCount > 0 Then
            Mensajes = "No se puede eliminar esta SubCuenta," & Chr(13)
            Mensajes = Mensajes & "porque tiene cuentas procesables."
            MsgBox Mensajes
         Else
            Mensajes = "Esta seguro que desea eliminar la " & Chr(13)
            Mensajes = Mensajes & "Cuenta No. [" & Cadena & "]"
            Titulo = "Pregunta de Eliminacion"
            TipoDeCaja = 4 + 32: ResultBox = MsgBox(Mensajes, TipoDeCaja, Titulo)
            If ResultBox = 6 Then
               sSQL = "DELETE * FROM Beneficiarios "
               sSQL = sSQL & "WHERE Codigo = '" & Cadena & "' "
               DeleteData AdoSubCta1, sSQL
            End If
         End If
         AdoSubCtas.Refresh
         DLCtas.SetFocus
      End If
   Case "Nuevo"
       NuevaCta
       Nuevo = True
       DLCtas.Enabled = False
       TextSubCta.SetFocus
   Case "Grabar"
       GrabarCta LabelCodigo.Caption
   Case "Primero"
       Nuevo = False
       .MoveFirst
   Case "Anterior"
       Nuevo = False
       .MovePrevious
       If .BOF Then .MoveFirst
   Case "Siguiente"
       Nuevo = False
       .MoveNext
       If .EOF Then .MoveLast
   Case "Ultimo"
       Nuevo = False
       .MoveLast
 End Select
 If Nuevo = False Then
    'DLCtas.Text = .Fields(0)
    Cadena = SinEspaciosIzq(DLCtas.Text)
    LlenarCta Cadena
 End If
 End With
End Sub

Public Sub LlenarCta(CodigoCta As String)
   sSQL = "SELECT * FROM Beneficiarios "
   sSQL = sSQL & "WHERE Codigo = '" & CodigoCta & "' "
   If OpcB.Value Then
      sSQL = sSQL & "AND TC = 'P' "
   Else
      sSQL = sSQL & "AND TC = 'R' "
   End If
   SelectData AdoSubCta1, sSQL, False
   With AdoSubCta1.Recordset
    If .RecordCount > 0 Then
        TextSubCta.Text = .Fields("Beneficiario")
        LabelCodigo.Caption = .Fields("Codigo")
        TextCiudad.Text = .Fields("Ciudad")
        TextDirecc.Text = .Fields("Direccion")
        MBoxRUC.Text = .Fields("RUC_CI")
        MBoxTelefono.Text = .Fields("Telefono")
        MBoxCelular.Text = .Fields("Celular")
        MBoxFAX.Text = .Fields("FAX")
    Else
        DLCtas.Enabled = False
        TextSubCta.Text = ""
        LabelCodigo.Caption = "0000"
        Nuevo = True
        TextSubCta.SetFocus
    End If
   End With
   DLCtas.Enabled = True
End Sub

Public Sub NuevaCta()
  DLCtas.Enabled = False
  TextSubCta.Text = ""
  LabelCodigo.Caption = "0000"
End Sub

Public Sub GrabarCta(CodigoCta As String)
  TextoValido TextCiudad
  TextoValido TextDirecc
  sSQL = "SELECT * FROM Beneficiarios "
  sSQL = sSQL & "WHERE Codigo = '" & CodigoCta & "' "
  If OpcB.Value Then
     sSQL = sSQL & "AND TC = 'P' "
  Else
     sSQL = sSQL & "AND TC = 'R' "
  End If
  SelectData AdoSubCta1, sSQL, False
  With AdoSubCta1.Recordset
   If .RecordCount > 0 Then
      .Edit
       Codigo = .Fields("Codigo")
   Else
      .AddNew
       Numero = ReadSetDataNum("Beneficiarios", True, True)
       Codigo = FormatoCodigo(TextSubCta.Text, Numero)
   End If
   If OpcB.Value Then
     .Fields("TC") = "P"
   Else
     .Fields("TC") = "R"
   End If
  .Fields("Codigo") = Codigo
  .Fields("Beneficiario") = Mid(TextSubCta.Text, 1, 35)
  .Fields("Ciudad") = TextCiudad.Text
  .Fields("Direccion") = TextDirecc.Text
  .Fields("RUC_CI") = MBoxRUC.Text
  .Fields("Telefono") = MBoxTelefono.Text
  .Fields("Celular") = MBoxCelular.Text
  .Fields("FAX") = MBoxFAX.Text
  .Update
   Nuevo = False
   AdoSubCtas.Refresh
   DLCtas.Enabled = True
  End With
End Sub

Private Sub MBoxCelular_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBoxFAX_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBoxRUC_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBoxTelefono_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextCiudad_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextCiudad_LostFocus()
  TextoValido TextCiudad, False
End Sub

Private Sub TextDirecc_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextDirecc_LostFocus()
  TextoValido TextDirecc, False
End Sub

