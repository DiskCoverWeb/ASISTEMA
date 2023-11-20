VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Begin VB.Form FAsignaFact 
   BackColor       =   &H00808080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CUENTAS POR COBRAR / CUENTAS POR PAGAR"
   ClientHeight    =   6585
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11580
   Icon            =   "AsigFact.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   11580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TxtDesc2 
      Alignment       =   1  'Right Justify
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
      Left            =   10395
      MaxLength       =   8
      TabIndex        =   12
      Text            =   "0.00"
      Top             =   3465
      Width           =   1065
   End
   Begin VB.TextBox TxtDia 
      Alignment       =   2  'Center
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
      Left            =   2730
      MaxLength       =   2
      TabIndex        =   6
      Text            =   "31"
      Top             =   3465
      Width           =   540
   End
   Begin VB.TextBox TxtDesc 
      Alignment       =   1  'Right Justify
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
      Left            =   7770
      MaxLength       =   8
      TabIndex        =   10
      Text            =   "0.00"
      Top             =   3465
      Width           =   1065
   End
   Begin MSDataListLib.DataCombo DCInv 
      Bindings        =   "AsigFact.frx":014A
      DataSource      =   "AdoCxCxP"
      Height          =   2910
      Left            =   2205
      TabIndex        =   4
      Top             =   525
      Width           =   8100
      _ExtentX        =   14288
      _ExtentY        =   5133
      _Version        =   393216
      Style           =   1
      BackColor       =   12632256
      ForeColor       =   8388608
      Text            =   "Productos"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.ListBox LstMeses 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2985
      Left            =   105
      Style           =   1  'Checkbox
      TabIndex        =   3
      Top             =   840
      Width           =   2010
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Modificar"
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
      Left            =   10395
      Picture         =   "AsigFact.frx":0161
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   945
      Width           =   1065
   End
   Begin MSDataGridLib.DataGrid DGRubros 
      Bindings        =   "AsigFact.frx":07F7
      Height          =   2535
      Left            =   105
      TabIndex        =   16
      Top             =   3885
      Width           =   11355
      _ExtentX        =   20029
      _ExtentY        =   4471
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
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
            LCID            =   3082
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
            LCID            =   3082
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
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Insertar"
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
      Left            =   10395
      Picture         =   "AsigFact.frx":080F
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   105
      Width           =   1065
   End
   Begin VB.TextBox TxtArea 
      Alignment       =   1  'Right Justify
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
      Left            =   5355
      MaxLength       =   8
      TabIndex        =   8
      Text            =   "0.00"
      Top             =   3465
      Width           =   1065
   End
   Begin MSAdodcLib.Adodc AdoCxCxP 
      Height          =   330
      Left            =   210
      Top             =   4620
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
      BackColor       =   &H00E0E0E0&
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
      Left            =   10395
      Picture         =   "AsigFact.frx":10D9
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   1785
      Width           =   1065
   End
   Begin MSAdodcLib.Adodc AdoRubros 
      Height          =   330
      Left            =   2100
      Top             =   4620
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
      Caption         =   "Rubros"
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
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " DESCUENTO 2"
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
      Left            =   8925
      TabIndex        =   11
      Top             =   3465
      Width           =   1485
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &Dia"
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
      Left            =   2205
      TabIndex        =   5
      Top             =   3465
      Width           =   540
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " DESCUENTO"
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
      Left            =   6510
      TabIndex        =   9
      Top             =   3465
      Width           =   1275
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   105
      TabIndex        =   2
      Top             =   525
      Width           =   2010
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
      Left            =   8820
      TabIndex        =   1
      Top             =   105
      Width           =   1485
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " VALOR A FACTURAR "
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
      Left            =   3360
      TabIndex        =   7
      Top             =   3465
      Width           =   2010
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
      Width           =   8625
   End
End
Attribute VB_Name = "FAsignaFact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim FechaSigMes As String
  TextoValido TxtArea, True, True
  Valor = CCur(TxtArea.Text)
  CodigoP = SinEspaciosIzq(DCInv.Text)
  If CodigoP <> Ninguno And CodigoCliente <> Ninguno Then
     sSQL = "DELETE * " _
          & "FROM Clientes_Facturacion " _
          & "WHERE Codigo_Inv = '" & CodigoP & "' " _
          & "AND Codigo = '" & CodigoCliente & "' " _
          & "AND Num_Mes = 0 " _
          & "AND Item = '" & NumEmpresa & "' "
     Ejecutar_SQL_SP sSQL
     For I = 1 To 12
         If LstMeses.Selected(I) = True Then
            Dia = Val(TxtDia)
            Anio = Val(Label1.Caption)
            Select Case I
              Case 1, 3, 5, 7, 8, 10, 12
                   If (Dia > 31) Then Dia = 31
              Case 2
                   If ((Anio Mod 4 <> 0) And (Dia > 28)) Then Dia = 28
                   If ((Anio Mod 4 = 0) And (Dia > 29)) Then Dia = 29
              Case 4, 6, 9, 11
                   If (Dia > 30) Then Dia = 30
              Case Else
                   Dia = 30
            End Select
            FechaSigMes = Format$(Dia, "00") & "/" & Format$(I, "00") & "/" & Format$(Anio, "0000")
            sSQL = "DELETE * " _
                 & "FROM Clientes_Facturacion " _
                 & "WHERE Codigo_Inv = '" & CodigoP & "' " _
                 & "AND Codigo = '" & CodigoCliente & "' " _
                 & "AND Periodo = '" & Label1.Caption & "' " _
                 & "AND Num_Mes = " & CByte(I) & " " _
                 & "AND Item = '" & NumEmpresa & "' "
            Ejecutar_SQL_SP sSQL
            SetAdoAddNew "Clientes_Facturacion"
            SetAdoFields "T", Normal
            SetAdoFields "Codigo", CodigoCliente
            SetAdoFields "Valor", CCur(TxtArea)
            SetAdoFields "Descuento", CCur(TxtDesc)
            SetAdoFields "Descuento2", CCur(TxtDesc2)
            SetAdoFields "Codigo_Inv", CodigoP
            SetAdoFields "Num_Mes", I
            SetAdoFields "Mes", MesesLetras(CInt(I))
            SetAdoFields "GrupoNo", Codigo2
            SetAdoFields "Fecha", FechaSigMes
            SetAdoFields "Item", NumEmpresa
            SetAdoFields "Periodo", Label1.Caption
            SetAdoFields "CodigoU", CodigoUsuario
            SetAdoUpdate
         End If
     Next I
  Else
     MsgBox "No hay datos para ingresar"
  End If
  Unload FAsignaFact
End Sub

Private Sub Command2_Click()
  Codigo = ""
  CodigoP = ""
  Unload FAsignaFact
End Sub

Private Sub Command3_Click()
  Codigo = ""
  CodigoP = ""
  Unload FAsignaFact
End Sub

Private Sub DCInv_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DGRubros_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  If KeyCode = vbKeyReturn Then
     AdoRubros.Recordset.MoveNext
     If AdoRubros.Recordset.EOF Then AdoRubros.Recordset.MoveFirst
  End If
  If AdoRubros.Recordset.RecordCount > 0 Then
     If CtrlDown And KeyCode = vbKeyV Then
        Codigos = DGRubros.Columns(1).Text
        Valor = Val(DGRubros.Columns(2).Text)
        Valor = Val(InputBox("INGRESE EL VALOR A" & vbCrLf & vbCrLf & "MODIFICAR DEL CODIGO: " & Codigos, "CAMBIO DE VALORES", Format$(Valor, "##0.00")))
        If Valor > 0 Then
           sSQL = "UPDATE Clientes_Facturacion " _
                & "SET Valor = " & Val(Valor) & " " _
                & "WHERE Codigo = '" & CodigoCliente & "' " _
                & "AND Codigo_Inv = '" & Codigos & "' " _
                & "AND Item = '" & NumEmpresa & "' "
           Ejecutar_SQL_SP sSQL
           Listar_Rubros_Grupo
        End If
     End If
     If CtrlDown And KeyCode = vbKeyD Then
        Codigos = DGRubros.Columns(1).Text
        Valor = Val(DGRubros.Columns(3).Text)
        Valor = Val(InputBox("INGRESE EL VALOR A" & vbCrLf & vbCrLf & "MODIFICAR DEL CODIGO: " & Codigos, "CAMBIO DE VALORES", Format$(Valor, "##0.00")))
        If Valor >= 0 Then
           sSQL = "UPDATE Clientes_Facturacion " _
                & "SET Descuento = " & Val(Valor) & " " _
                & "WHERE Codigo = '" & CodigoCliente & "' " _
                & "AND Codigo_Inv = '" & Codigos & "' " _
                & "AND Item = '" & NumEmpresa & "' "
           Ejecutar_SQL_SP sSQL
           Listar_Rubros_Grupo
        End If
     End If
     If CtrlDown And KeyCode = vbKey2 Then
        Codigos = DGRubros.Columns(1).Text
        Valor = Val(DGRubros.Columns(3).Text)
        Valor = Val(InputBox("INGRESE EL VALOR A" & vbCrLf & vbCrLf & "MODIFICAR DEL CODIGO: " & Codigos, "CAMBIO DE VALORES", Format$(Valor, "##0.00")))
        If Valor >= 0 Then
           sSQL = "UPDATE Clientes_Facturacion " _
                & "SET Descuento2 = " & Val(Valor) & " " _
                & "WHERE Codigo = '" & CodigoCliente & "' " _
                & "AND Codigo_Inv = '" & Codigos & "' " _
                & "AND Item = '" & NumEmpresa & "' "
           Ejecutar_SQL_SP sSQL
           Listar_Rubros_Grupo
        End If
     End If
  End If
End Sub

Private Sub Form_Activate()
  TxtArea = "0.00"
  TxtDesc = "0.00"
  TxtDesc2 = "0.00"
  LblCodigo.Caption = CodigoCliente
  LblCliente.Caption = Codigo1
  FAsignaFact.Caption = "CUENTAS POR COBRAR EXTRACONTABLE DEL: " & Codigo2
  Label1.Caption = Codigo4
  Label5.Caption = "VALOR A FACTURAR"
  RatonReloj
  LstMeses.Clear
  LstMeses.AddItem "Todos los Meses"
  sSQL = "SELECT * " _
       & "FROM Tabla_Dias_Meses " _
       & "WHERE Tipo = 'M' " _
       & "AND No_D_M > 0 " _
       & "ORDER BY No_D_M "
  Select_Adodc AdoRubros, sSQL
  With AdoRubros.Recordset
   If .RecordCount Then
       Do While Not .EOF
          LstMeses.AddItem .fields("Dia_Mes")
         .MoveNext
       Loop
   End If
  End With
  Listar_Rubros_Grupo
  sSQL = "SELECT Codigo_Inv & '  ' & Producto As NomProd,* " _
       & "FROM Catalogo_Productos " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND TC = 'P' " _
       & "AND LEN(Cta_Inventario) = 1 " _
       & "AND INV <> " & Val(adFalse) & " " _
       & "ORDER BY Codigo_Inv "
  SelectDB_Combo DCInv, AdoCxCxP, sSQL, "NomProd"
End Sub

Private Sub Form_Load()
  CentrarForm FAsignaFact
  ConectarAdodc AdoCxCxP
  ConectarAdodc AdoRubros
  FAsignaFact.Caption = "ASIGNACION DE CODIGO DE FACTURACION"
End Sub

Private Sub LstMeses_LostFocus()
  If LstMeses.Selected(0) = True Then
     For I = 1 To 12
         LstMeses.Selected(I) = True
     Next I
     LstMeses.Selected(0) = False
  End If
End Sub

Private Sub TxtDesc_GotFocus()
  MarcarTexto TxtDesc
End Sub

Private Sub TxtDesc_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtDesc_LostFocus()
  TextoValido TxtDesc, True, , 2
End Sub

Private Sub TxtDesc2_GotFocus()
  MarcarTexto TxtDesc2
End Sub

Private Sub TxtDesc2_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtDesc2_LostFocus()
  TextoValido TxtDesc2, True, , 2
End Sub

Private Sub TxtArea_GotFocus()
  MarcarTexto TxtArea
End Sub

Private Sub TxtArea_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtArea_LostFocus()
  TextoValido TxtArea, True, , 2
End Sub

Public Sub Listar_Rubros_Grupo()
    sSQL = "SELECT Mes, Codigo_Inv, Valor, Descuento, Descuento2, Periodo, Fecha, Codigo, Item " _
         & "FROM Clientes_Facturacion " _
         & "WHERE Codigo = '" & CodigoCliente & "' " _
         & "AND Item = '" & NumEmpresa & "' " _
         & "ORDER BY Periodo, Num_Mes, Codigo_Inv "
    Select_Adodc_Grid DGRubros, AdoRubros, sSQL
End Sub

Private Sub TxtDia_GotFocus()
  MarcarTexto TxtDia
End Sub

Private Sub TxtDia_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtDia_LostFocus()
  If Not IsNumeric(TxtDia) Then
     MsgBox "Dato Incorrecto"
     TxtDia.SetFocus
  Else
     If Val(TxtDia) > 31 Then
        MsgBox "Numero incorrecto"
        TxtDia.SetFocus
     Else
        TxtDia = Format$(Val(TxtDia), "00")
     End If
  End If
End Sub

