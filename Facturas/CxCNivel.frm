VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.4#0"; "COMCTL32.OCX"
Begin VB.Form CxCNivel 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5070
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   5580
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   5580
   Begin ComctlLib.TreeView TVNivel 
      Height          =   4740
      Left            =   0
      TabIndex        =   1
      Top             =   315
      Width           =   5580
      _ExtentX        =   9843
      _ExtentY        =   8361
      _Version        =   327682
      HideSelection   =   0   'False
      Style           =   7
      ImageList       =   "ImgList"
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&X"
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
      Left            =   5250
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   330
   End
   Begin MSAdodcLib.Adodc AdoNivel 
      Height          =   330
      Left            =   105
      Top             =   1995
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
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
      Caption         =   "Nivel"
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
   Begin VB.Label LabelRuta 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CxC"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   330
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   5265
   End
   Begin ComctlLib.ImageList ImgList 
      Left            =   105
      Top             =   420
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   2
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "CxCNivel.frx":0000
            Key             =   "Abrir"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "CxCNivel.frx":031A
            Key             =   "Cerrar"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "CxCNivel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
  Unload CxCNivel
End Sub

Private Sub Form_Activate()
Dim nodX As Node
' Establece propiedades del control ImageList.
  TVNivel.LineStyle = tvwTreeLines
' Crea un árbol con varios objetos Node sin ordenar.
  sSQL = "SELECT * " _
       & "FROM Catalogo_Lineas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND TL = False " _
       & "ORDER BY Codigo "
  SelectData AdoNivel, sSQL
  With ImgList
   If AdoNivel.Recordset.RecordCount > 0 Then
      Do While Not AdoNivel.Recordset.EOF
         Codigo = AdoNivel.Recordset.Fields("Codigo")
         Cadena = AdoNivel.Recordset.Fields("Concepto")
         If Len(Codigo) = 2 Then
            Set nodX = TVNivel.Nodes.Add(, , Codigo, Cadena, .ListImages(2).key, .ListImages(1).key)
         Else
            Set nodX = TVNivel.Nodes.Add(CambioCodigoCtaSup(Codigo), tvwChild, Codigo, Cadena, .ListImages(2).key, .ListImages(1).key)
         End If
         AdoNivel.Recordset.MoveNext
      Loop
      nodX.EnsureVisible
      RatonNormal
      TVNivel.SetFocus
   Else
      RatonNormal
      Unload CxCNivel
   End If
  End With
End Sub

Private Sub Form_Load()
   CentrarForm CxCNivel
   ConectarAdodc AdoNivel
End Sub

Private Sub TVNivel_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyReturn
         If NivelNo <> Ninguno Then
          ' Encontro el tipo de Factura
            sSQL = "SELECT * " _
                 & "FROM Catalogo_Lineas " _
                 & "WHERE Codigo = '" & NivelNo & "' "
            SelectData AdoNivel, sSQL
            With AdoNivel.Recordset
             If .RecordCount > 0 Then
                 NivelNo = .Fields("Codigo")
                 Cta_General = .Fields("CxC")
                 Cta_Ventas = .Fields("Cta_Venta")
                 TipoProc = .Fields("TP")
             End If
            End With
            Unload CxCNivel
            'FacturasTours.Show
         End If
    Case vbKeyEscape: Unload CxCNivel
  End Select
End Sub

Private Sub TVNivel_NodeClick(ByVal Node As ComctlLib.Node)
' Obtenemos el codigo del menu
  TipoFacturas = " " & Node.FullPath
  LabelRuta.Caption = " " & Node.FullPath
  NivelNo = Ninguno
  If Node.Children = 0 Then NivelNo = Mid$(Node.key, 2, Len(Node.key) - 1)
End Sub
