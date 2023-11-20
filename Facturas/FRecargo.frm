VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FRecargo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10230
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   10230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command6 
      Caption         =   "&Multas"
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
      Left            =   8610
      Picture         =   "FRecargo.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3465
      Width           =   1485
   End
   Begin VB.CheckBox CheqRangos 
      Caption         =   "Procesar &Por Rangos Grupos:"
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
      Value           =   1  'Checked
      Width           =   2955
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Insertar &Todos"
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
      Left            =   8610
      Picture         =   "FRecargo.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   105
      Width           =   1485
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Eli&minar Todos"
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
      Left            =   8610
      Picture         =   "FRecargo.frx":1194
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2625
      Width           =   1485
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
      Left            =   8610
      Picture         =   "FRecargo.frx":15D6
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4305
      Width           =   1485
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
      Left            =   4200
      MaxLength       =   8
      TabIndex        =   7
      Top             =   4305
      Width           =   1485
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
      Height          =   3435
      Left            =   105
      Style           =   1  'Checkbox
      TabIndex        =   4
      Top             =   840
      Width           =   2010
   End
   Begin MSDataListLib.DataCombo DCInv 
      Bindings        =   "FRecargo.frx":1EA0
      DataSource      =   "AdoCxCxP"
      Height          =   3690
      Left            =   2205
      TabIndex        =   5
      Top             =   525
      Width           =   6315
      _ExtentX        =   11139
      _ExtentY        =   6509
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
   Begin MSAdodcLib.Adodc AdoCxCxP 
      Height          =   330
      Left            =   210
      Top             =   1470
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
   Begin MSAdodcLib.Adodc AdoRubros 
      Height          =   330
      Left            =   210
      Top             =   1155
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
   Begin MSDataListLib.DataCombo DCGrupoI 
      Bindings        =   "FRecargo.frx":1EB7
      DataSource      =   "AdoGrupo"
      Height          =   360
      Left            =   3150
      TabIndex        =   1
      Top             =   105
      Width           =   2640
      _ExtentX        =   4657
      _ExtentY        =   635
      _Version        =   393216
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo DCGrupoF 
      Bindings        =   "FRecargo.frx":1ECE
      DataSource      =   "AdoGrupo"
      Height          =   360
      Left            =   5880
      TabIndex        =   2
      Top             =   105
      Width           =   2640
      _ExtentX        =   4657
      _ExtentY        =   635
      _Version        =   393216
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSAdodcLib.Adodc AdoGrupo 
      Height          =   330
      Left            =   210
      Top             =   1785
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
      Caption         =   "Grupo"
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
   Begin MSMask.MaskEdBox MBFechaT 
      Height          =   330
      Left            =   7035
      TabIndex        =   9
      Top             =   4725
      Width           =   1485
      _ExtentX        =   2619
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
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "0"
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&Fecha Tope de Pagos "
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
      TabIndex        =   8
      Top             =   4725
      Width           =   4845
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
      TabIndex        =   3
      Top             =   525
      Width           =   2010
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "VALOR A FACTURAR "
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
      TabIndex        =   6
      Top             =   4305
      Width           =   2010
   End
End
Attribute VB_Name = "FRecargo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Mas_Grupos As Boolean

Private Sub CheqRangos_Click()
 If CheqRangos.Value = 0 Then
    DCGrupoI.Enabled = False
    DCGrupoF.Enabled = False
 Else
    DCGrupoI.Enabled = True
    DCGrupoF.Enabled = True
 End If
End Sub

Private Sub Command1_Click()
  TextoValido TxtArea, True, True
  Valor = CCur(TxtArea.Text)
  CodigoP = SinEspaciosIzq(DCInv.Text)
  sSQL = "SELECT * " _
       & "FROM Clientes " _
       & "WHERE FA <> " & Val(adFalse) & " " _
       & "AND Grupo = '" & Codigo1 & "' "
  If Mas_Grupos Then sSQL = sSQL & "AND DirNumero = '" & NumEmpresa & "' "
  sSQL = sSQL & "ORDER BY Grupo,Cliente,Sexo "
  SelectAdodc AdoRubros, sSQL
  With AdoRubros.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          CodigoCliente = .Fields("Codigo")
          Codigo1 = .Fields("Grupo")
          For I = 1 To 12
           If LstMeses.Selected(I) = True Then
              sSQL = "DELETE * " _
                   & "FROM Clientes_Facturacion " _
                   & "WHERE Codigo_Inv = '" & CodigoP & "' " _
                   & "AND Codigo = '" & CodigoCliente & "' " _
                   & "AND Num_Mes = " & CByte(I) & " " _
                   & "AND Periodo = '" & Label1.Caption & "' " _
                   & "AND Item = '" & NumEmpresa & "' "
              ConectarAdoExecute sSQL
              SetAdoAddNew "Clientes_Facturacion"
              SetAdoFields "T", Normal
              SetAdoFields "Codigo", CodigoCliente
              SetAdoFields "Valor", Valor
              SetAdoFields "Descuento", CCur(TxtDesc)
              SetAdoFields "Codigo_Inv", CodigoP
              SetAdoFields "Num_Mes", I
              SetAdoFields "GrupoNo", Codigo1
              SetAdoFields "Mes", MesesLetras(CInt(I))
              If CFechaLong(MBFechaT) < CFechaLong("01/" & Format(I, "00") & "/" & Label1.Caption) Then
                 SetAdoFields "Fecha", "01/" & Format(I, "00") & "/" & Label1.Caption
              Else
                 SetAdoFields "Fecha", MBFechaT
              End If
              SetAdoFields "Item", NumEmpresa
              SetAdoFields "Periodo", Label1.Caption
              SetAdoFields "CodigoU", CodigoUsuario
              SetAdoUpdate
           End If
          Next I
         .MoveNext
       Loop
   End If
  End With
  Unload FPensiones
End Sub

Private Sub Command2_Click()
  Unload FPensiones
End Sub

Private Sub Command3_Click()
  If ClaveSupervisor Then
  CodigoP = SinEspaciosIzq(DCInv.Text)
  sSQL = "SELECT * " _
       & "FROM Clientes " _
       & "WHERE FA <> " & Val(adFalse) & " "
  If Mas_Grupos Then sSQL = sSQL & "AND DirNumero = '" & NumEmpresa & "' "
  If CheqRangos.Value <> 0 Then
     Codigo1 = DCGrupoI.Text
     Codigo2 = DCGrupoF.Text
     If Codigo1 = "" Then Codigo1 = Ninguno
     If Codigo2 = "" Then Codigo2 = Ninguno
     sSQL = sSQL & "AND Grupo BETWEEN '" & Codigo1 & "' and '" & Codigo2 & "' "
  End If
  sSQL = sSQL & "ORDER BY Grupo,Cliente,Sexo "
  'MsgBox sSQL
  SelectAdodc AdoRubros, sSQL
  With AdoRubros.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          CodigoCliente = .Fields("Codigo")
          Codigo1 = .Fields("Grupo")
          FPensiones.Caption = Format(Contador / .RecordCount, "00%") & " - ELIMINACION DE CODIGOS DE FACTURARION - GRUPO: " & Codigo1
          For I = 1 To 12
           If LstMeses.Selected(I) = True Then
              sSQL = "DELETE * " _
                   & "FROM Clientes_Facturacion " _
                   & "WHERE Codigo_Inv = '" & CodigoP & "' " _
                   & "AND Codigo = '" & CodigoCliente & "' " _
                   & "AND Num_Mes = " & CByte(I) & " " _
                   & "AND Periodo = '" & Label1.Caption & "' " _
                   & "AND Item = '" & NumEmpresa & "' "
              ConectarAdoExecute sSQL
           End If
          Next I
         .MoveNext
       Loop
   End If
  End With
  Unload FPensiones
  End If
End Sub

Private Sub Command4_Click()
  If ClaveSupervisor Then
     If LstMeses.ListCount > 1 Then
     TextoValido TxtArea, True, True
  Valor = CCur(TxtArea.Text)
  CodigoP = SinEspaciosIzq(DCInv.Text)
  Contador = 0
  sSQL = "SELECT * " _
       & "FROM Clientes " _
       & "WHERE FA <> " & Val(adFalse) & " "
  If Mas_Grupos Then sSQL = sSQL & "AND DirNumero = '" & NumEmpresa & "' "
  If CheqRangos.Value <> 0 Then
     Codigo1 = DCGrupoI.Text
     Codigo2 = DCGrupoF.Text
     If Codigo1 = "" Then Codigo1 = Ninguno
     If Codigo2 = "" Then Codigo2 = Ninguno
     sSQL = sSQL & "AND Grupo BETWEEN '" & Codigo1 & "' and '" & Codigo2 & "' "
  End If
  sSQL = sSQL & "ORDER BY Grupo,Cliente,Sexo "
  'MsgBox sSQL
  SelectAdodc AdoRubros, sSQL
  With AdoRubros.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          Contador = Contador + 1
          CodigoCliente = .Fields("Codigo")
          Codigo1 = .Fields("Grupo")
          FPensiones.Caption = Format(Contador / .RecordCount, "00%") & " - ASIGNACION DE CODIGOS DE FACTURARION - GRUPO: " & Codigo1
          For I = 1 To 12
           If LstMeses.Selected(I) = True Then
              sSQL = "DELETE * " _
                   & "FROM Clientes_Facturacion " _
                   & "WHERE Codigo_Inv = '" & CodigoP & "' " _
                   & "AND Codigo = '" & CodigoCliente & "' " _
                   & "AND Num_Mes = " & CByte(I) & " " _
                   & "AND Periodo = '" & Label1.Caption & "' " _
                   & "AND Item = '" & NumEmpresa & "' "
              ConectarAdoExecute sSQL
              SetAdoAddNew "Clientes_Facturacion"
              SetAdoFields "T", Normal
              SetAdoFields "Codigo", CodigoCliente
              SetAdoFields "Valor", Valor
              SetAdoFields "Descuento", CCur(TxtDesc)
              SetAdoFields "Codigo_Inv", CodigoP
              SetAdoFields "Num_Mes", I
              SetAdoFields "GrupoNo", Codigo1
              SetAdoFields "Mes", MesesLetras(CInt(I))
              If CFechaLong(MBFechaT) < CFechaLong("01/" & Format(I, "00") & "/" & Label1.Caption) Then
                 SetAdoFields "Fecha", "01/" & Format(I, "00") & "/" & Label1.Caption
              Else
                 SetAdoFields "Fecha", MBFechaT
              End If
              SetAdoFields "Item", NumEmpresa
              SetAdoFields "Periodo", Label1.Caption
              SetAdoFields "CodigoU", CodigoUsuario
              SetAdoUpdate
           End If
          Next I
         .MoveNext
       Loop
   End If
  End With
     Else
        Contador = 0
        CodigoP = SinEspaciosIzq(DCInv.Text)
        sSQL = "SELECT Codigo,GrupoNo,COUNT(Codigo) As Cantidad,SUM(Valor) As Valor_Fact " _
             & "FROM Clientes_Facturacion " _
             & "WHERE Item = '" & NumEmpresa & "' " _
             & "GROUP BY Codigo,GrupoNo "
        SelectAdodc AdoRubros, sSQL
        With AdoRubros.Recordset
         If .RecordCount > 0 Then
             Do While Not .EOF
                Contador = Contador + 1
                CodigoCliente = .Fields("Codigo")
                Codigo1 = .Fields("GrupoNo")
                Valor = .Fields("Valor_Fact")
                Cantidad = .Fields("Cantidad")
                FPensiones.Caption = "Deuda Pendiente: " & Format(Contador / .RecordCount, "00%") & " - ASIGNACION DE CODIGOS DE FACTURARION - GRUPO: " & Codigo1
                SetAdoAddNew "Clientes_Facturacion"
                SetAdoFields "T", Normal
                SetAdoFields "Codigo", CodigoCliente
                SetAdoFields "Valor", Valor
                SetAdoFields "Descuento", CCur(TxtDesc)
                SetAdoFields "Codigo_Inv", CodigoP
                SetAdoFields "Num_Mes", 0
                SetAdoFields "GrupoNo", Codigo1
                SetAdoFields "Mes", "(" & Format(Cantidad, "00") & ")"
                SetAdoFields "Item", NumEmpresa
                SetAdoFields "Periodo", Label1.Caption
                SetAdoFields "CodigoU", CodigoUsuario
                SetAdoUpdate
                sSQL = "DELETE * " _
                     & "FROM Clientes_Facturacion " _
                     & "WHERE Codigo = '" & CodigoCliente & "' " _
                     & "AND Num_Mes <> 0 " _
                     & "AND Item = '" & NumEmpresa & "' "
                ConectarAdoExecute sSQL
               .MoveNext
             Loop
        End If
       End With
     End If
  End If
  Unload FPensiones
End Sub

Private Sub Command5_Click()
  CodigoP = SinEspaciosIzq(DCInv.Text)
  sSQL = "SELECT * " _
       & "FROM Clientes " _
       & "WHERE FA <> " & Val(adFalse) & " " _
       & "AND Grupo = '" & Codigo1 & "' "
  If Mas_Grupos Then sSQL = sSQL & "AND DirNumero = '" & NumEmpresa & "' "
  sSQL = sSQL & "ORDER BY Grupo,Cliente,Sexo "
  SelectAdodc AdoRubros, sSQL
  With AdoRubros.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          CodigoCliente = .Fields("Codigo")
          Codigo1 = .Fields("Grupo")
          For I = 1 To 12
           If LstMeses.Selected(I) = True Then
              sSQL = "DELETE * " _
                   & "FROM Clientes_Facturacion " _
                   & "WHERE Codigo_Inv = '" & CodigoP & "' " _
                   & "AND Codigo = '" & CodigoCliente & "' " _
                   & "AND Num_Mes = " & CByte(I) & " " _
                   & "AND Periodo = '" & Label1.Caption & "' " _
                   & "AND Item = '" & NumEmpresa & "' "
              ConectarAdoExecute sSQL
           End If
          Next I
         .MoveNext
       Loop
   End If
  End With
  Unload FPensiones
End Sub

Private Sub Command6_Click()
  If ClaveSupervisor Then
  If LstMeses.ListCount > 1 Then
     TextoValido TxtArea, True, True
     FechaValida MBFechaT
     Actualizar_Saldos_Facturas MBFechaT, "FA"
     Actualizar_Saldos_Facturas MBFechaT, "FM"
     Actualizar_Saldos_Facturas MBFechaT, "NV"
     FechaIni = BuscarFecha("01/" & Month(MBFechaT) & "/" & Year(MBFechaT))
     FechaFin = BuscarFecha(UltimoDiaMes(MBFechaT))
     'MsgBox UltimoDiaMes(MBFechaT)
     Valor = CCur(TxtArea.Text)
     CodigoP = SinEspaciosIzq(DCInv.Text)
     Contador = 0
     sSQL = "SELECT C.Codigo,C.Grupo,C.Cliente,F.Saldo_Actual,F.T " _
          & "FROM Clientes As C, Facturas As F " _
          & "WHERE C.Codigo = F.CodigoC " _
          & "AND F.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
          & "AND F.Saldo_Actual > 0 "
     If Mas_Grupos Then sSQL = sSQL & "AND C.DirNumero = '" & NumEmpresa & "' "
     If CheqRangos.Value <> 0 Then
        Codigo1 = DCGrupoI.Text
        Codigo2 = DCGrupoF.Text
        If Codigo1 = "" Then Codigo1 = Ninguno
        If Codigo2 = "" Then Codigo2 = Ninguno
        sSQL = sSQL & "AND C.Grupo BETWEEN '" & Codigo1 & "' and '" & Codigo2 & "' "
     End If
     sSQL = sSQL & "ORDER BY C.Grupo,C.Cliente "
    'MsgBox sSQL
     SelectAdodc AdoRubros, sSQL
     With AdoRubros.Recordset
      If .RecordCount > 0 Then
          Do While Not .EOF
             Contador = Contador + 1
             CodigoCliente = .Fields("Codigo")
             Codigo1 = .Fields("Grupo")
             FPensiones.Caption = Format(Contador / .RecordCount, "00%") & " - ASIGNACION DE CODIGOS DE FACTURARION - GRUPO: " & Codigo1
             For I = 1 To 12
              If LstMeses.Selected(I) = True Then
                 sSQL = "DELETE * " _
                      & "FROM Clientes_Facturacion " _
                      & "WHERE Codigo_Inv = '" & CodigoP & "' " _
                      & "AND Codigo = '" & CodigoCliente & "' " _
                      & "AND Num_Mes = " & CByte(I) & " " _
                      & "AND Periodo = '" & Label1.Caption & "' " _
                      & "AND Item = '" & NumEmpresa & "' "
                 ConectarAdoExecute sSQL
                 SetAdoAddNew "Clientes_Facturacion"
                 SetAdoFields "T", Normal
                 SetAdoFields "Codigo", CodigoCliente
                 SetAdoFields "Valor", Valor
                 SetAdoFields "Codigo_Inv", CodigoP
                 SetAdoFields "Num_Mes", I
                 SetAdoFields "GrupoNo", Codigo1
                 SetAdoFields "Mes", MesesLetras(CInt(I))
                 SetAdoFields "Item", NumEmpresa
                 SetAdoFields "Periodo", Label1.Caption
                 SetAdoFields "CodigoU", CodigoUsuario
                 SetAdoUpdate
              End If
             Next I
            .MoveNext
          Loop
      End If
     End With
  End If
  End If
  'MsgBox "....."
  Unload FPensiones
End Sub


Private Sub Form_Activate()
  Mas_Grupos = Leer_Campo_Empresa("Separar_Grupos")
  FPensiones.Caption = "CUENTAS POR COBRAR EXTRACONTABLE"
  Label1.Caption = Codigo4
  LstMeses.Clear
  If Si_No Then
     LstMeses.AddItem "Deuda Pendiente"
     Command1.Enabled = False
     Command3.Enabled = False
     Label5.Visible = False
     TxtArea.Visible = False
  Else
     LstMeses.AddItem "Todos los Meses"
     sSQL = "SELECT * " _
          & "FROM Tabla_Meses " _
          & "WHERE NoMes <> 0 " _
          & "ORDER BY NoMes "
     SelectAdodc AdoRubros, sSQL
     With AdoRubros.Recordset
      If .RecordCount Then
          Do While Not .EOF
             LstMeses.AddItem .Fields("Mes")
            .MoveNext
          Loop
      End If
     End With
  End If
  sSQL = "SELECT Codigo_Inv & '  ' & Producto As NomProd,* " _
       & "FROM Catalogo_Productos " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND TC = 'P' " _
       & "ORDER BY Codigo_Inv "
  SelectDBCombo DCInv, AdoCxCxP, sSQL, "NomProd"
  sSQL = "SELECT Grupo " _
       & "FROM Clientes " _
       & "WHERE FA <> " & Val(adFalse) & " "
  If Mas_Grupos Then sSQL = sSQL & "AND DirNumero = '" & NumEmpresa & "' "
  sSQL = sSQL & "GROUP BY Grupo " _
       & "ORDER BY Grupo "
  SelectDBCombo DCGrupoI, AdoGrupo, sSQL, "Grupo"
  SelectDBCombo DCGrupoF, AdoGrupo, sSQL, "Grupo"
  If AdoGrupo.Recordset.RecordCount > 0 Then
     AdoGrupo.Recordset.MoveLast
     DCGrupoF.Text = AdoGrupo.Recordset.Fields("Grupo")
  End If
  RatonNormal
End Sub

Private Sub Form_Load()
  CentrarForm FPensiones
  ConectarAdodc AdoGrupo
  ConectarAdodc AdoCxCxP
  ConectarAdodc AdoRubros
  FPensiones.Caption = "ASIGNACION DE CODIGO DE FACTURACION   GRUPO: " & Codigo1
End Sub

Private Sub LstMeses_LostFocus()
  If LstMeses.ListCount > 1 Then
     If LstMeses.Selected(0) = True Then
        For I = 1 To 12
            LstMeses.Selected(I) = True
        Next I
        LstMeses.Selected(0) = False
     End If
  End If
End Sub

Private Sub MBFechaT_GotFocus()
  MarcarTexto MBFechaT
End Sub

Private Sub MBFechaT_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBFechaT_LostFocus()
  FechaValida MBFechaT
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

Private Sub TxtDesc_GotFocus()
  MarcarTexto TxtDesc
End Sub

Private Sub TxtDesc_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtDesc_LostFocus()
  TextoValido TxtDesc, True, , 2
End Sub
