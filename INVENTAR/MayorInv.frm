VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Begin VB.Form MayorizarInv 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   450
   ClientLeft      =   15
   ClientTop       =   -30
   ClientWidth     =   7260
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   450
   ScaleWidth      =   7260
   Begin VB.PictureBox PictTotal 
      Height          =   435
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   7200
      TabIndex        =   0
      Top             =   0
      Width           =   7260
   End
   Begin MSAdodcLib.Adodc AdoCtas 
      Height          =   330
      Left            =   3780
      Top             =   0
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
      Caption         =   "Ctas"
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
   Begin MSAdodcLib.Adodc AdoTrans 
      Height          =   330
      Left            =   0
      Top             =   0
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
      Caption         =   "Trans"
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
   Begin MSAdodcLib.Adodc AdoSubCtas 
      Height          =   330
      Left            =   1785
      Top             =   0
      Visible         =   0   'False
      Width           =   2010
      _ExtentX        =   3545
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
      Caption         =   "SubCtas"
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
End
Attribute VB_Name = "MayorizarInv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Mayorizavion de los Inventarios

Private Sub Form_Activate()
Dim Primero As Boolean
Dim Cod_Cta As String
Dim Num_Comp As Long
Dim CantBarra As Integer
Dim CantBodega As Integer
Dim UpdateCodInv As Boolean
Dim ValorUnitAntOld As Currency
  RatonReloj
  LineasDeTexto = ""
  MiTiempo = Time
  CantBarra = 0: CantBodega = 0
  Progreso_Barra.Incremento = 0
  Progreso_Barra.Valor_Maximo = 100
  Progreso_Barra.Mensaje_Box = "Mayorizando Existencias: Espere un momento..."
  Progreso_Esperar
  
  sSQL = "UPDATE Trans_Kardex " _
       & "SET X = 'X' " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' "
  ConectarAdoExecute sSQL
    
  If SQL_Server Then
     sSQL = "UPDATE Trans_Kardex " _
          & "SET X = '.' " _
          & "FROM Trans_Kardex As TK,Comprobantes As C "
  Else
     sSQL = "UPDATE Trans_Kardex As TK,Comprobantes As C " _
          & "SET TK.X = '.' "
  End If
  sSQL = sSQL & "WHERE C.Item = '" & NumEmpresa & "' " _
       & "AND C.Periodo = '" & Periodo_Contable & "' " _
       & "AND C.T <> 'A' " _
       & "AND TK.TP = C.TP " _
       & "AND TK.Numero = C.Numero " _
       & "AND TK.Fecha = C.Fecha " _
       & "AND TK.Item = C.Item " _
       & "AND TK.Periodo = C.Periodo "
  ConectarAdoExecute sSQL

  sSQL = "DELETE * " _
       & "FROM Trans_Kardex " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND X <> '.' "
  ConectarAdoExecute sSQL
        
  sSQL = "DELETE * " _
       & "FROM Catalogo_Productos " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND LEN(Codigo_Inv) < 2 "
  ConectarAdoExecute sSQL

  sSQL = "UPDATE Catalogo_Productos " _
       & "SET Cta_Inventario = '0', Cta_Costo_Venta = '0', " _
       & "Cta_Ventas = '0', Cta_Ventas_0 = '0' " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND TC = 'I' "
  ConectarAdoExecute sSQL
  
'''  sSQL = "UPDATE Catalogo_Productos " _
'''       & "SET INV = " & Val(adFalse) & " " _
'''       & "WHERE Item = '" & NumEmpresa & "' " _
'''       & "AND LEN(Cta_Inventario) <= 1 "
'''  ConectarAdoExecute sSQL
  
  sSQL = "UPDATE Catalogo_Productos " _
       & "SET INV = " & Val(adTrue) & " " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND TC = 'I' "
  ConectarAdoExecute sSQL
  
  sSQL = "UPDATE Trans_Kardex " _
       & "SET Valor_Total = Valor_Unitario * Entrada, Total = Valor_Unitario * Entrada " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND Entrada > 0 " _
       & "AND T <> 'A' "
  ConectarAdoExecute sSQL
  
  sSQL = "SELECT CP.Codigo_Inv,CP.Item,COUNT(TK.Codigo_Inv) As CantMov " _
       & "FROM Catalogo_Productos As CP,Trans_Kardex As TK " _
       & "WHERE CP.Item = '" & NumEmpresa & "' " _
       & "AND CP.Periodo = '" & Periodo_Contable & "' " _
       & "AND CP.TC = 'P' " _
       & "AND TK.T <> 'A' " _
       & "AND CP.Item = TK.Item " _
       & "AND CP.Periodo = TK.Periodo " _
       & "AND CP.Codigo_Inv = TK.Codigo_Inv " _
       & "GROUP BY CP.Codigo_Inv,CP.Item " _
       & "ORDER BY CP.Codigo_Inv,CP.Item "
  SelectAdodc AdoCtas, sSQL
' (1) Mayorizando: Todo el Kardex
  If AdoCtas.Recordset.RecordCount > 0 Then
     AdoCtas.Recordset.MoveFirst
     Progreso_Barra.Valor_Maximo = AdoCtas.Recordset.RecordCount
     Do While Not AdoCtas.Recordset.EOF
        CodigoInv = AdoCtas.Recordset.Fields("Codigo_Inv")
        Progreso_Barra.Mensaje_Box = "(1) Mayorizando: [" & Format$(Time - MiTiempo, "HH:MM:SS") & "] - " & CodigoInv
        Progreso_Esperar
        sSQL = "SELECT * " _
             & "FROM Trans_Kardex " _
             & "WHERE T <> '" & Anulado & "' " _
             & "AND Codigo_Inv = '" & CodigoInv & "' " _
             & "AND Item = '" & NumEmpresa & "' " _
             & "AND Periodo = '" & Periodo_Contable & "' " _
             & "ORDER BY Fecha,Entrada DESC,Salida,TP,Numero,Kardex "
        SelectAdodc AdoSubCtas, sSQL
        With AdoSubCtas.Recordset
         If .RecordCount > 0 Then
             Total = 0
             Cantidad = 0
             SaldoAnterior = 0
             ValorUnit = .Fields("Valor_Unitario")
             ValorUnitAnt = ValorUnit  ' Costo
             IniciarStock = True
             If .Fields("Salida") > 0 Then LineasDeTexto = LineasDeTexto & "Total " & CodigoInv & vbCrLf
             Do While Not .EOF
                Cantidad = Redondear(Cantidad + .Fields("Entrada") - .Fields("Salida"), 2)
                If .Fields("Entrada") > 0 Then
                    ValorUnit = Redondear(.Fields("Valor_Unitario"), Dec_Costo)
                    ValorTotal = Redondear(.Fields("Entrada") * ValorUnit, 4)
                    Total = Total + ValorTotal
                  ' Calculamos el costeo
                    If Cantidad > 0 Then ValorUnitAnt = Redondear(Total / Cantidad, Dec_Costo)
                    If ValorUnitAnt <= 0 Then ValorUnitAnt = ValorUnit
                End If
                If .Fields("Salida") > 0 Then
                    ValorTotal = Redondear(.Fields("Salida") * ValorUnitAnt, 4)
                    Total = Total - ValorTotal
                    ValorUnit = ValorUnitAnt
                End If
                UpdateCodInv = False
                If .Fields("Valor_Total") <> ValorTotal Then UpdateCodInv = True
                If .Fields("Existencia") <> Cantidad Then UpdateCodInv = True
                If .Fields("Costo") <> ValorUnitAnt Then UpdateCodInv = True
                If .Fields("Total") <> Total Then UpdateCodInv = True
                If .Fields("Valor_Unitario") <> ValorUnit Then UpdateCodInv = True
                If UpdateCodInv Then
                  .Fields("Valor_Total") = ValorTotal
                  .Fields("Existencia") = Cantidad
                  .Fields("Costo") = ValorUnitAnt
                  .Fields("Total") = Total
                  .Fields("Valor_Unitario") = ValorUnit
                  .Update
                End If
               .MoveNext
             Loop
         End If
        End With
        AdoCtas.Recordset.MoveNext
     Loop
     Progreso_Barra.Mensaje_Box = "(1) Mayorizando..."
     Progreso_Esperar
  End If
 '(2) Mayorizando: Bodegas
 '(3) Mayorizando: Codigo de Barras
  If LineasDeTexto <> "" And Presentar_Inventario Then MsgBox "Codigo sin Mayorizar correctamente:" & vbCrLf & LineasDeTexto & "No existe asiento inicial."
  RatonNormal
  Unload MayorizarInv
End Sub

Private Sub Form_Load()
  CentrarForm MayorizarInv
  ConectarAdodc AdoCtas
  ConectarAdodc AdoTrans
  ConectarAdodc AdoSubCtas
End Sub

Private Sub ProcBar_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
  'Forma de mayorizar
End Sub

Private Sub PictTotal_Click()

End Sub
