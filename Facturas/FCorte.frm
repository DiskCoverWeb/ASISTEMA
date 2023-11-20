VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FCorte 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "VALORES DEL CORTE"
   ClientHeight    =   4230
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4650
   Icon            =   "FCorte.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   4650
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox TxtDias 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   1995
      MaxLength       =   6
      MultiLine       =   -1  'True
      TabIndex        =   19
      Text            =   "FCorte.frx":030A
      Top             =   3045
      Width           =   1485
   End
   Begin VB.TextBox Txt2PorcServ 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   1995
      MaxLength       =   10
      MultiLine       =   -1  'True
      TabIndex        =   9
      Text            =   "FCorte.frx":030C
      Top             =   1365
      Width           =   1485
   End
   Begin VB.TextBox TxtTransporte 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   1995
      MaxLength       =   10
      MultiLine       =   -1  'True
      TabIndex        =   7
      Text            =   "FCorte.frx":030E
      Top             =   1050
      Width           =   1485
   End
   Begin VB.TextBox Txt2Porc 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   1995
      MaxLength       =   10
      MultiLine       =   -1  'True
      TabIndex        =   5
      Text            =   "FCorte.frx":0310
      Top             =   735
      Width           =   1485
   End
   Begin VB.TextBox TxtTotal 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   330
      Left            =   1995
      MaxLength       =   12
      MultiLine       =   -1  'True
      TabIndex        =   23
      Text            =   "FCorte.frx":0312
      Top             =   3780
      Width           =   1485
   End
   Begin VB.TextBox TxtPVPCorte 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   1995
      MaxLength       =   10
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "FCorte.frx":0314
      Top             =   420
      Width           =   1485
   End
   Begin VB.TextBox TxtCantCorte 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   1995
      MaxLength       =   8
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "FCorte.frx":0316
      Top             =   105
      Width           =   1485
   End
   Begin VB.OptionButton Option2 
      Caption         =   "1/2 Corte"
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
      Left            =   1995
      TabIndex        =   11
      Top             =   1785
      Width           =   1170
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Corte Entero"
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
      TabIndex        =   10
      Top             =   1785
      Value           =   -1  'True
      Width           =   1590
   End
   Begin VB.TextBox TxtNumCorte 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   1995
      MaxLength       =   6
      MultiLine       =   -1  'True
      TabIndex        =   13
      Text            =   "FCorte.frx":0318
      Top             =   2100
      Width           =   1485
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Aceptar"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   3570
      Picture         =   "FCorte.frx":031A
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   105
      Width           =   960
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   3570
      Picture         =   "FCorte.frx":075C
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   945
      Width           =   960
   End
   Begin MSAdodcLib.Adodc AdoCorte 
      Height          =   330
      Left            =   105
      Top             =   4095
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
      Caption         =   "Corte"
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
   Begin MSAdodcLib.Adodc AdoAux 
      Height          =   330
      Left            =   1890
      Top             =   4095
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
   Begin VB.Label Label11 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Interés de Crédito"
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
      TabIndex        =   18
      Top             =   3045
      Width           =   1905
   End
   Begin VB.Label Label9 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Días de Crédito"
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
      TabIndex        =   16
      Top             =   2730
      Width           =   1905
   End
   Begin VB.Label LblDias 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   1995
      TabIndex        =   17
      Top             =   2730
      Width           =   1485
   End
   Begin VB.Label Label10 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Nuevo P.V.P."
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
      TabIndex        =   20
      Top             =   3360
      Width           =   1905
   End
   Begin VB.Label LblPVP 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   1995
      TabIndex        =   21
      Top             =   3360
      Width           =   1485
   End
   Begin VB.Label LblCorte 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   1995
      TabIndex        =   15
      Top             =   2415
      Width           =   1485
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Valor del Corte"
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
      TabIndex        =   14
      Top             =   2415
      Width           =   1905
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " 2% por Servicio"
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
      TabIndex        =   8
      Top             =   1365
      Width           =   1905
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Transporte"
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
      Top             =   1050
      Width           =   1905
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Si menor a 500 2%"
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
      TabIndex        =   4
      Top             =   735
      Width           =   1905
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " T O T A L"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   330
      Left            =   105
      TabIndex        =   22
      Top             =   3780
      Width           =   1905
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Precio por Pliego"
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
      Width           =   1905
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Cantidad"
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
      Width           =   1905
   End
   Begin VB.Label Label16 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Número de Corte"
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
      Top             =   2100
      Width           =   1905
   End
End
Attribute VB_Name = "FCorte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
  CantAnterior = 1: VUnitTemp = 0
  Unload FCorte
End Sub

Private Sub Command2_Click()
 CantAnterior = Cantidad
 VUnitTemp = CDbl(TxtTotal)
 Unload FCorte
End Sub

Private Sub Form_Activate()
  LblDias = NoDias & " día(s)"
  TxtCantCorte = Format$(Cantidad, "#,##0")
  TxtPVPCorte = Format$(Precio, "#,##0." & String(Dec_PVP, "0"))
  SubTotal = Cantidad * Precio
  sSQL = "SELECT * " _
       & "FROM Tabla_Costos " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "ORDER BY Concepto "
  SelectAdodc AdoCorte, sSQL
End Sub

Private Sub Form_Load()
  CentrarForm FCorte
  ConectarAdodc AdoCorte
  ConectarAdodc AdoAux
End Sub

Private Sub TxtDias_GotFocus()
  MarcarTexto TxtDias
End Sub

Private Sub Txt2Porc_GotFocus()
  If Val(CCur(TxtCantCorte)) < 500 Then SubTotal = SubTotal + (SubTotal * 0.02)
  Txt2Porc = Format$(SubTotal, "#,##0.00")
End Sub

Private Sub Txt2Porc_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub Txt2PorcServ_GotFocus()
  If SubTotal < 100 Then SubTotal = SubTotal + (SubTotal * 0.02)
  Txt2PorcServ = Format$(SubTotal, "#,##0.00")
End Sub

Private Sub TxtCantCorte_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtCantCorte_LostFocus()
  TxtTotal = Format$(SubTotal, "#,##0.00")
  Cantidad = Val(CCur(TxtCantCorte))
End Sub

Private Sub TxtDias_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtDias_LostFocus()
  SubTotal = SubTotal + Val(CCur(TxtDias))
  LblPVP = Format$(SubTotal / Cantidad, "#,##0." & String(Dec_PVP, "0"))
  TxtTotal = Format$(SubTotal, "#,##0.00")
  TxtTotal.SetFocus
End Sub

Private Sub TxtNumCorte_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtNumCorte_GotFocus()
  MarcarTexto TxtNumCorte
End Sub

Private Sub TxtNumCorte_LostFocus()
  Tasa = 0: VUnitTemp = 0
  With AdoCorte.Recordset
   If .RecordCount > 0 Then
       NumMeses = CInt(TxtNumCorte)
       If NumMeses > 0 Then
          If Option1.value Then
            .MoveFirst
            .Find ("Concepto = 'CORTE' ")
             If Not .EOF Then Tasa = .Fields("Valor")
          Else
            .MoveFirst
            .Find ("Concepto = 'MEDIOCORTE' ")
             If Not .EOF Then Tasa = .Fields("Valor")
          End If
          VUnitTemp = NumMeses * Tasa       'Tasa = Valor Corte
       End If
   End If
  End With
  LblCorte = Format$(VUnitTemp, "#,##0." & String(Dec_PVP, "0"))       'Tasa = Valor Corte
  SubTotal = SubTotal + VUnitTemp
  LblPVP = Format$(SubTotal / Cantidad, "#,##0." & String(Dec_PVP, "0"))
  TxtTotal = Format$(SubTotal, "#,##0.00")
End Sub

Private Sub TxtPVPCorte_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtPVPCorte_LostFocus()
  SubTotal = Val(CSng(TxtCantCorte)) * Val(CSng(TxtPVPCorte))
  TxtTotal = Format$(SubTotal, "#,##0.00")
End Sub

Private Sub TxtTransporte_GotFocus()
  TxtTransporte = "5.00"
  MarcarTexto TxtTransporte
End Sub

Private Sub TxtTransporte_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtTransporte_LostFocus()
   SubTotal = SubTotal + Val(CCur(TxtTransporte))
   TxtTotal = Format$(SubTotal, "#,##0.00")
End Sub
