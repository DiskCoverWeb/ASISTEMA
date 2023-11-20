VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.4#0"; "comctl32.Ocx"
Begin VB.Form FHabitacion 
   Caption         =   "DETALLE DE HABITACIONES"
   ClientHeight    =   7485
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11700
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7485
   ScaleWidth      =   11700
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture3 
      Height          =   540
      Left            =   5775
      Picture         =   "FHabitac.frx":0000
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   22
      Top             =   2310
      Width           =   540
   End
   Begin VB.PictureBox Picture2 
      Height          =   540
      Left            =   5775
      Picture         =   "FHabitac.frx":030A
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   21
      Top             =   1785
      Width           =   540
   End
   Begin VB.PictureBox Picture1 
      Height          =   540
      Left            =   5775
      Picture         =   "FHabitac.frx":0BD4
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   20
      Top             =   1260
      Width           =   540
   End
   Begin MSDataGridLib.DataGrid DGDetalle 
      Bindings        =   "FHabitac.frx":1016
      Height          =   4005
      Left            =   5880
      TabIndex        =   3
      Top             =   2940
      Width           =   5790
      _ExtentX        =   10213
      _ExtentY        =   7064
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   12298
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   12298
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox PictMatriz 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFC0C0&
      Height          =   4005
      Left            =   5775
      ScaleHeight     =   6.959
      ScaleMode       =   7  'Centimeter
      ScaleWidth      =   10.107
      TabIndex        =   2
      Top             =   2940
      Visible         =   0   'False
      Width           =   5790
   End
   Begin ComctlLib.TreeView TVNivel 
      Height          =   6840
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   5580
      _ExtentX        =   9843
      _ExtentY        =   12065
      _Version        =   327682
      HideSelection   =   0   'False
      Style           =   7
      ImageList       =   "ImgList"
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSAdodcLib.Adodc AdoNivel 
      Height          =   330
      Left            =   315
      Top             =   945
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
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
   Begin MSAdodcLib.Adodc AdoDetalle 
      Height          =   330
      Left            =   315
      Top             =   1365
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
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
      Caption         =   "Detalle"
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
   Begin VB.Label LblTotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   10185
      TabIndex        =   11
      Top             =   2520
      Width           =   1380
   End
   Begin VB.Label Label9 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Total Consum."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   330
      Left            =   8610
      TabIndex        =   10
      Top             =   2520
      Width           =   1590
   End
   Begin VB.Label LblServicio 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   10185
      TabIndex        =   19
      Top             =   2205
      Width           =   1380
   End
   Begin VB.Label Label17 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Servicio 10%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   330
      Left            =   8610
      TabIndex        =   18
      Top             =   2205
      Width           =   1590
   End
   Begin VB.Label LblIVA 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   10185
      TabIndex        =   17
      Top             =   1890
      Width           =   1380
   End
   Begin VB.Label LblIVA1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " I.V.A. 12%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   330
      Left            =   8610
      TabIndex        =   16
      Top             =   1890
      Width           =   1590
   End
   Begin VB.Label LblValorHab 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   10185
      TabIndex        =   26
      Top             =   1575
      Width           =   1380
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Habitacion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   330
      Left            =   8610
      TabIndex        =   27
      Top             =   1575
      Width           =   1590
   End
   Begin VB.Label LblSubTotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   10185
      TabIndex        =   15
      Top             =   1260
      Width           =   1380
   End
   Begin VB.Label Label13 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " SubTotal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   330
      Left            =   8610
      TabIndex        =   14
      Top             =   1260
      Width           =   1590
   End
   Begin VB.Label LblDias 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   10185
      TabIndex        =   9
      Top             =   945
      Width           =   1380
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Días de Estadia"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   330
      Left            =   8610
      TabIndex        =   8
      Top             =   945
      Width           =   1590
   End
   Begin VB.Label LblValor 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   7665
      TabIndex        =   28
      Top             =   945
      Width           =   960
   End
   Begin VB.Label Label10 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Valor de Habitacion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   330
      Left            =   5775
      TabIndex        =   29
      Top             =   945
      Width           =   1905
   End
   Begin VB.Label Label5 
      Caption         =   " Habitacion Reservada"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   330
      Left            =   6300
      TabIndex        =   25
      Top             =   2415
      Width           =   2220
   End
   Begin VB.Label Label3 
      Caption         =   " Habitacion Ocupada"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   330
      Left            =   6300
      TabIndex        =   24
      Top             =   1890
      Width           =   2220
   End
   Begin VB.Label Label1 
      Caption         =   " Habitacion Libre"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   330
      Left            =   6300
      TabIndex        =   23
      Top             =   1365
      Width           =   2220
   End
   Begin VB.Label LblCliente 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   6930
      TabIndex        =   7
      Top             =   525
      Width           =   4635
   End
   Begin VB.Label LblFecha 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "00/00/0000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   10395
      TabIndex        =   13
      Top             =   105
      Width           =   1170
   End
   Begin VB.Label LblHabitacion 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   6930
      TabIndex        =   5
      Top             =   105
      Width           =   1275
   End
   Begin VB.Label Label11 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Fecha Ingreso"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   330
      Left            =   8925
      TabIndex        =   12
      Top             =   105
      Width           =   1485
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Cliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   330
      Left            =   5775
      TabIndex        =   6
      Top             =   525
      Width           =   1170
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Habitación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   330
      Left            =   5775
      TabIndex        =   4
      Top             =   105
      Width           =   1170
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
      Left            =   105
      TabIndex        =   1
      Top             =   7035
      Width           =   11460
   End
   Begin ComctlLib.ImageList ImgList 
      Left            =   315
      Top             =   210
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   4
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FHabitac.frx":102F
            Key             =   "Libre"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FHabitac.frx":1349
            Key             =   "Piso"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FHabitac.frx":1573
            Key             =   "Ocupado"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FHabitac.frx":188D
            Key             =   "Reservado"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FHabitacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
  Dim nodX As node
' Establece propiedades del control ImageList.
  TVNivel.LineStyle = tvwTreeLines
' Crea un árbol con varios objetos Node sin ordenar.
  sSQL = "SELECT CH.*,C.Cliente " _
       & "FROM Catalogo_Habitacion As CH,Clientes As C " _
       & "WHERE CH.Item = '" & NumEmpresa & "' " _
       & "AND CH.CodigoC = C.Codigo " _
       & "ORDER BY CH.Codigo "
  Select_Adodc AdoNivel, sSQL
  With ImgList
   If AdoNivel.Recordset.RecordCount > 0 Then
      Do While Not AdoNivel.Recordset.EOF
         Codigo = AdoNivel.Recordset.Fields("Codigo")
         Codigo1 = AdoNivel.Recordset.Fields("Habitacion")
         If AdoNivel.Recordset.Fields("Lleno") Then
            Cadena = AdoNivel.Recordset.Fields("Cliente")
         Else
            Cadena = AdoNivel.Recordset.Fields("Detalle")
         End If
         If AdoNivel.Recordset.Fields("TC") = "H" Then Cadena = Codigo1 & " - " & Cadena
         If Len(Codigo) = 3 Then
            Set nodX = TVNivel.Nodes.Add(, , Codigo, Cadena, .ListImages(2).key, .ListImages(2).key)
         Else
                                   'Hijo de                     es hijo   su papa, Detalle,
            Select Case AdoNivel.Recordset.Fields("Est")
              Case "L": Set nodX = TVNivel.Nodes.Add(CambioCodigoCtaSup(Codigo), tvwChild, Codigo, Cadena, .ListImages(1).key, .ListImages(1).key)
              Case "O": Set nodX = TVNivel.Nodes.Add(CambioCodigoCtaSup(Codigo), tvwChild, Codigo, Cadena, .ListImages(3).key, .ListImages(2).key)
              Case "R": Set nodX = TVNivel.Nodes.Add(CambioCodigoCtaSup(Codigo), tvwChild, Codigo, Cadena, .ListImages(4).key, .ListImages(4).key)
            End Select
         End If
         AdoNivel.Recordset.MoveNext
      Loop
      nodX.EnsureVisible
      RatonNormal
      TVNivel.SetFocus
   Else
      RatonNormal
      Unload FHabitacion
   End If
  End With
  RatonNormal
End Sub

Private Sub Form_Load()
  ConectarAdodc AdoNivel
  ConectarAdodc AdoDetalle
End Sub

Private Sub TVNivel_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyReturn
         If NivelNo <> Ninguno Then
            Mifecha = FechaSistema
          ' Encontro el tipo de Factura
            sSQL = "SELECT * " _
                 & "FROM Catalogo_Habitacion " _
                 & "WHERE Habitacion = '" & NivelNo & "' " _
                 & "AND Item = '" & NumEmpresa & "' "
            Select_Adodc AdoNivel, sSQL
            If AdoNivel.Recordset.RecordCount > 0 Then
               Mifecha = AdoNivel.Recordset.Fields("Fecha")
               Valor = AdoNivel.Recordset.Fields("Valor")
            End If
            LblValor.Caption = Format$(Valor, "#,##0.00")
            LblFecha.Caption = Mifecha
            LblHabitacion.Caption = NivelNo
            sSQL = "SELECT TH.Fecha,TH.Cantidad,CP.Producto,CP.PVP,(CP.PVP*TH.Cantidad) As Total " _
                 & "FROM Trans_Habitacion TH,Catalogo_Productos As CP " _
                 & "WHERE TH.Habitacion = '" & NivelNo & "' " _
                 & "AND TH.Item = '" & NumEmpresa & "' " _
                 & "AND TH.Fecha >= #" & BuscarFecha(Mifecha) & "# " _
                 & "AND TH.Item = CP.Item " _
                 & "AND TH.Codigo_Inv = CP.Codigo_Inv " _
                 & "ORDER BY TH.Fecha,CP.Producto "
            Select_Adodc_Grid DGDetalle, AdoDetalle, sSQL
            Total = 0: PorcEnvio = 0
            With AdoDetalle.Recordset
             If .RecordCount > 0 Then
                 Do While Not .EOF
                    Total = Total + .Fields("Total")
                   .MoveNext
                 Loop
             End If
            End With
            PorcEnvio = CFechaLong(FechaSistema) - CFechaLong(Mifecha)
            LblDias.Caption = PorcEnvio
            LblSubTotal.Caption = Format$(Total, "#,##0.00")
            LblValorHab.Caption = Format$(Valor * PorcEnvio, "#,##0.00")
            Total = Total + (Valor * PorcEnvio)
            LblIVA.Caption = Format$(Total * 0.12, "#,##0.00")
            LblServicio.Caption = Format$(Total * 0.1, "#,##0.00")
            LblTotal.Caption = Format$(Total * 1.22, "#,##0.00")
            DGDetalle.Caption = "PEDIDOS DE LA HABITACION: " & NivelNo
         End If
    Case vbKeyEscape: Unload Me
  End Select
End Sub

Private Sub TVNivel_NodeClick(ByVal node As ComctlLib.node)
' Obtenemos el codigo del menu
  TipoFacturas = " " & node.FullPath
  LabelRuta.Caption = " " & node.FullPath
  NivelNo = Ninguno
  If node.children = 0 Then
     DGDetalle.Visible = True
     PictMatriz.Visible = False
  Else
     DGDetalle.Visible = False
     PictMatriz.Visible = True
  End If
  NivelNo = SinEspaciosIzq(node.Text)
End Sub


