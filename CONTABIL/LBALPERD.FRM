VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form ListBalPerd 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mayor Analitico"
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7950
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   7950
   ShowInTaskbar   =   0   'False
   Begin VB.Data DataFechaBal 
      Caption         =   "FechaBal"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Data DataResult 
      Caption         =   "Result"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "IMPRIMIR RESULTADOS"
      Height          =   735
      Left            =   3120
      TabIndex        =   4
      Top             =   3360
      Width           =   1815
   End
   Begin MSDBGrid.DBGrid DBGResult 
      Bindings        =   "LBalPerd.frx":0000
      Height          =   2535
      Left            =   120
      OleObjectBlob   =   "LBalPerd.frx":0015
      TabIndex        =   3
      Top             =   600
      Width           =   7695
   End
   Begin VB.Label LabelTotReg 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
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
      Height          =   255
      Left            =   1920
      TabIndex        =   8
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label LabelTotalReg 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Total de Registros:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   3360
      Width           =   1815
   End
   Begin VB.Label LabelTotPag 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
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
      Height          =   255
      Left            =   1920
      TabIndex        =   6
      Top             =   3840
      Width           =   1095
   End
   Begin VB.Label LabelTotalPagina 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Total de Página(s):"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   3840
      Width           =   1815
   End
   Begin VB.Label LabelTotSaldo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      TabIndex        =   2
      Top             =   3840
      Width           =   2775
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Resultado de Perdidas o Ganancias"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5040
      TabIndex        =   1
      Top             =   3240
      Width           =   2775
   End
   Begin VB.Label Label5 
      Caption         =   "Label5"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3495
   End
End
Attribute VB_Name = "ListBalPerd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command2_Click()
Dim C, F As Long
MousePointer = vbHourglass
On Error GoTo Errorhandler
Pagina = 1: FinImp = 25: Printer.Orientation = 1
TipoEscala = 7: Printer.ScaleMode = TipoEscala
Printer.DrawWidth = 1: Printer.FontSize = 8: Printer.Font.Bold = True

F = DataResult.Recordset.RecordCount
C = DataResult.Recordset.Fields.Count
ReDim Campos(C + 1) As String
ReDim Anchos(C + 1) As Single
ReDim Tipos(C + 1)
ReDim Sizes(C + 1)
For J = 0 To C - 1
  Tipos(J) = DataResult.Recordset.Fields(J).Type
  Sizes(J) = DataResult.Recordset.Fields(J).Size
  Campos(J) = DataResult.Recordset.Fields(J).Name
  Anchos(J) = Printer.TextWidth(Campos(J)) + 0.2
  If Anchos(J) < 1 Then Anchos(J) = 1
Next J
Printer.Font.Bold = False
For J = 0 To C - 1
  Ancho = AnchoTipoCampo(Tipos(J), Sizes(J)) + 0.2
  If Ancho > Anchos(J) Then Anchos(J) = Ancho
Next J
Anchos(C) = Anchos(C - 1)
For J = C - 1 To 1 Step -1
  Anchos(J) = Anchos(J - 1)
Next J
Anchos(0) = 1.5
For J = C To 0 Step -1
  Suma1 = 0
  For I = J To 0 Step -1
     Suma1 = Suma1 + Anchos(I)
  Next
  Anchos(J) = Suma1
Next
' Iniciamos la impresion
PosLinea = 1
Printer.FontSize = 8: Printer.Font.Bold = True
Printer.CurrentX = 2: Printer.CurrentY = PosLinea
Printer.Print Empresa
PosLinea = PosLinea + 0.4
Printer.CurrentX = 2: Printer.CurrentY = PosLinea
Printer.Print Direccion
PosLinea = PosLinea + 0.4
Printer.CurrentX = 2: Printer.CurrentY = PosLinea
Printer.Print "Teléfono(s): " & Telefono1 & "/" & Telefono2
PosLinea = 2.2: Printer.FontSize = 18
Cadena = "RESULTADOS DE"
Printer.CurrentX = CentrarTexto(Cadena): Printer.CurrentY = PosLinea
Printer.Print Cadena
PosLinea = PosLinea + 0.8
Cadena = "PERDIDAS O GANACIAS"
Printer.CurrentX = CentrarTexto(Cadena): Printer.CurrentY = PosLinea
Printer.Print Cadena
'========================================================================
PosLinea = PosLinea + 1
Printer.Line (1.5, PosLinea)-(Anchos(C), PosLinea), QBColor(0)
PosLinea = PosLinea + 0.1
Printer.Line (1.5, PosLinea)-(Anchos(C), PosLinea), QBColor(0)
PosLinea = PosLinea + 0.5
'============================================================
' Comenzamos a escribir en la impresora los encabezados
'============================================================
Printer.Font.Size = 8
For J = 0 To C - 1
   Printer.Line (Anchos(J), PosLinea - 0.1)-(Anchos(J), PosLinea + 0.4), QBColor(0)
   Printer.CurrentX = Anchos(J) + 0.1: Printer.CurrentY = PosLinea
   Printer.Print Campos(J)
Next J
Printer.Line (Anchos(C), PosLinea - 0.1)-(Anchos(C), PosLinea + 0.4), QBColor(0)
Printer.Line (1.5, PosLinea - 0.1)-(Anchos(C), PosLinea - 0.1), QBColor(0)
Printer.Line (1.5, PosLinea + 0.4)-(Anchos(C), PosLinea + 0.4), QBColor(0)
PosLinea = PosLinea + 0.5
Printer.Font.Bold = False
' comenzamos a imprimir la base
'==============================
DataResult.Recordset.MoveFirst
Do While Not DataResult.Recordset.EOF
  Ancho = InicioX
  For J = 0 To C - 1
    Printer.Line (Anchos(J), PosLinea - 0.1)-(Anchos(J), PosLinea + 0.4), QBColor(0)
    Printer.CurrentX = Anchos(J) + CampoWidth(DataResult.Recordset.Fields(Campos(J)), Tipos(J), Sizes(J)) + 0.1
    Printer.CurrentY = PosLinea
    Printer.Print FormatoTipoCampo(DataResult.Recordset.Fields(Campos(J)), Tipos(J))
  Next J
  Printer.Line (Anchos(C), PosLinea - 0.1)-(Anchos(C), PosLinea + 0.4), QBColor(0)
  PosLinea = PosLinea + 0.4
  DataResult.Recordset.MoveNext
Loop
Printer.Line (1.5, PosLinea)-(Anchos(C), PosLinea), QBColor(0)
PosLinea = PosLinea + 0.1
Printer.Line (1.5, PosLinea)-(Anchos(C), PosLinea), QBColor(0)
PosLinea = PosLinea + 1: Printer.Font.Size = 12
Printer.CurrentX = 1.5: Printer.CurrentY = PosLinea
Printer.Print "RESULTADO TOTAL: "
Printer.CurrentX = 8: Printer.CurrentY = PosLinea
Printer.Print Format(Sumatoria, "#,##0.00")
Printer.EndDoc
MousePointer = vbDefault
MsgBox "Se ha enviado a la impresora: " & Pagina & " pagina(s)."
Exit Sub
Errorhandler:
    MousePointer = vbDefault
    MsgBox "Error: No pudo imprimir correctamente."
    Exit Sub
End Sub

Private Sub Form_Activate()
  MousePointer = vbHourglass
  Sumatoria = 0
  SSQL = "SELECT * FROM Resultados "
  DataResult.RecordSource = SSQL: DataResult.Refresh
  If DataResult.Recordset.RecordCount > 0 Then
     DataResult.Recordset.MoveLast
     DataResult.Recordset.Edit
     DataResult.Recordset.Fields("Total") = 0
     DataResult.Recordset.Update
     DataResult.Recordset.MoveFirst
     Do While Not DataResult.Recordset.EOF
        Codigos = DataResult.Recordset.Fields("OP")
        Saldo = DataResult.Recordset.Fields("Total")
        If Codigos = "+" Then Sumatoria = Sumatoria + Saldo Else Sumatoria = Sumatoria - Saldo
        DataResult.Recordset.MoveNext
     Loop
     DataResult.Recordset.MoveLast
     DataResult.Recordset.Edit
     DataResult.Recordset.Fields("Total") = Sumatoria
     DataResult.Recordset.Update
  End If
  SSQL = "SELECT Codigo,Cuenta,Total FROM Resultados "
  DataResult.RecordSource = SSQL: DataResult.Refresh
  DBGResult.Columns(DataResult.Recordset.Fields.Count - 2).Width = TextWidth(String(40, "X"))
  DBGResult.Columns(DataResult.Recordset.Fields.Count - 1).NumberFormat = "#,##0.00"
  LabelTotSaldo.Caption = Format(Sumatoria, "#,##0.00")
  LabelTotReg.Caption = Format(DataResult.Recordset.RecordCount, "#,##0")
  LabelTotPag.Caption = "1"
  If Sumatoria >= 0 Then
     Cadena = "La Empresa puede continuar, su estado es de Ganancia."
  Else
     Cadena = "La Empresa tiene un saldo negativo, Su estado es de Perdida."
  End If
  DBGResult.Caption = Cadena
  MousePointer = vbDefault
End Sub

Private Sub Form_Load()
  'Centrar el formulario
  Left = (Screen.Width - Width) / 2
  Top = (Screen.Height - Height) / 2 - 600
  Label5.Caption = Empresa
  DataResult.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
  DataFechaBal.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
  MDIConta.MousePointer = vbDefault
End Sub

