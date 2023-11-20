VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form FImporFA 
   Caption         =   "Importar Datos"
   ClientHeight    =   8850
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12285
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8850
   ScaleWidth      =   12285
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command3 
      Caption         =   "&Salir"
      Height          =   750
      Left            =   2415
      TabIndex        =   2
      Top             =   105
      Width           =   1065
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   4140
      Left            =   105
      TabIndex        =   3
      Top             =   945
      Width           =   11670
      _ExtentX        =   20585
      _ExtentY        =   7303
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSAdodcLib.Adodc AdoAux 
      Height          =   330
      Left            =   210
      Top             =   2730
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
   Begin VB.CommandButton Command2 
      Caption         =   "Subir al Sistema"
      Height          =   750
      Left            =   1260
      TabIndex        =   1
      Top             =   105
      Width           =   1065
   End
   Begin MSComDlg.CommonDialog CDialogDir 
      Left            =   11760
      Top             =   7455
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin MSAdodcLib.Adodc AdoAct 
      Height          =   330
      Left            =   210
      Top             =   2310
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
      Caption         =   "Act"
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
   Begin MSAdodcLib.Adodc AdoLinea 
      Height          =   330
      Left            =   210
      Top             =   1890
      Visible         =   0   'False
      Width           =   2115
      _ExtentX        =   3731
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
      Caption         =   "Linea"
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
   Begin MSAdodcLib.Adodc AdoAsiento 
      Height          =   330
      Left            =   105
      Top             =   8400
      Width           =   4530
      _ExtentX        =   7990
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
      Caption         =   "Asiento"
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
   Begin MSForms.CommandButton CommandButton1 
      Height          =   750
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   1065
      Size            =   "1879;1323"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
End
Attribute VB_Name = "FImporFA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim CtasProc() As CtasAsiento
Dim ContCtas As Integer

'Variables para acceder a la hoja excel
Dim obj_Excel As Object, obj_Workbook As Object, obj_Worksheet As Object

'Type para el Rango
Private Type T_Rango
        NumFila1 As Long
        NumFila2 As Long
        NumCol1 As Long
        NumCol2 As Long
End Type

'Variable para el UDT que almacena cuatro variables par el Rango de datos de la hoha Excel
Dim rango As T_Rango

'función que recibe el Apth del ArchivoExcel y una variable de tipo T_Rango para almacenar los valores y retornarlos
Private Function Obtener_Rango_Excel(path_excel As String, rango As T_Rango) As Boolean
On Error GoTo ErrSub
'Variables de objeto Excel
Dim objExcel As Object, obj_Hoja As Object
Dim Primera_Fila As Integer, Primera_Column As Integer
Dim num_Filas As Integer, num_Col As Integer
   'Crear la instancia de la aplicación Excel
    Set objExcel = CreateObject("Excel.Application")
    objExcel.workbooks.Open FileName:=path_excel
    If Val(objExcel.Application.Version) >= 8 Then
        Set obj_Hoja = objExcel.ActiveSheet
    Else
        Set obj_Hoja = objExcel
    End If
   'Almacenamos en la variable Type el rango, es decir la primera fila la ultima fila, La primer Columna y la ultima
    With rango
        .NumFila1 = Format$(obj_Hoja.UsedRange.Row)
        .NumFila2 = Format$(obj_Hoja.UsedRange.Row + obj_Hoja.UsedRange.Rows.Count - 1)
        .NumCol1 = Format$(obj_Hoja.UsedRange.Column)
        .NumCol2 = Format$(obj_Hoja.UsedRange.Column + obj_Hoja.UsedRange.Columns.Count - 1)
    End With
    objExcel.ActiveWorkbook.Close False
   'cierra el Excel
    objExcel.Quit
   'Elimina las referencias y objetos
    Set obj_Hoja = Nothing
    Set objExcel = Nothing
Exit Function
ErrSub:
MsgBox " Número de Error: " & Err.Number & vbNewLine & "Descripción: " & "err.Description"
End Function

Private Sub Excel_FlexGrid(sPath As String, FlexGrid As Object, rango As T_Rango)

Dim I As Long
Dim N As Long
Dim No_Bancos As Long
    On Error GoTo error_sub
    If Len(Dir(sPath)) = 0 Then
       MsgBox "El archivo no existe", vbCritical
       Exit Sub
    End If
    RatonReloj
    No_Bancos = 0
    Set obj_Excel = CreateObject("Excel.Application")
    'obj_Excel.Visible = True
    Set obj_Workbook = obj_Excel.workbooks.Open(sPath)
    Set obj_Worksheet = obj_Workbook.ActiveSheet
   'MsgBox "..................>>>"
   'MSFlexGrid1.Visible = False
    MSHFlexGrid1.Visible = False
    MSHFlexGrid1.Clear
    With MSHFlexGrid1
       ' Especificar  acá la cantidad de filas y columnas
         Select Case Modulo
           Case "FACTURACION": .Cols = 20
           Case "CONTABILIDAD": .Cols = 30
           Case Else
               .Cols = 18
         End Select
        .Rows = rango.NumFila2
        'MsgBox rango.NumCol2 & vbCrLf & rango.NumFila2
       ' Recorremos las filas del FlexGrid para agregar los datos
         For N = 0 To .Cols - 4
          If TextWidth(CStr(obj_Worksheet.cells(1, N + 1).Value)) > .ColWidth(N) Then
            .ColWidth(N) = TextWidth(CStr(obj_Worksheet.cells(1, N + 1).Value) & "XXXX")
          End If
         Next
        
        For I = 0 To .Rows - 1
           'Recorremos las columnas y Fila del FlexGrid
            Si_No = True
            .Row = I
            .Col = 0
            .Text = Format(I, "0000")
            For N = 0 To .Cols - 4
                If N = 0 And Trim(CStr(obj_Worksheet.cells(I + 1, N + 1).Value)) = "" Then
                   Si_No = False
                   No_Bancos = No_Bancos + 1
                End If
                'MsgBox Si_No & vbCrLf & Trim(CStr(obj_Worksheet.cells(I + 1, N + 1).Value))
                If Si_No Then
                If TextWidth(CStr(obj_Worksheet.cells(I + 1, N + 1).Value)) > .ColWidth(N) Then
                  .ColWidth(N) = TextWidth(CStr(obj_Worksheet.cells(I + 1, N + 1).Value) & "XXXX")
                End If
               .Col = N + 1
               'Asignamos el Texto de la celda del Flex el contenido de la celda del excel
                If N = 5 Then Codigo1 = obj_Worksheet.cells(I + 1, N + 1).Value
                If N = 8 Then Codigo2 = obj_Worksheet.cells(I + 1, N + 1).Value
               .Text = obj_Worksheet.cells(I + 1, N + 1).Value
                End If
            Next
            If Si_No Then
            Codigo3 = Trim(SinEspaciosDer(Codigo2))
            If Len(Codigo1) < 3 Then Codigo1 = Codigo1 & "XX"
            If Len(Codigo2) < 3 Then Codigo2 = Codigo2 & "XX"
            If Len(Codigo3) < 3 Then Codigo3 = Codigo3 & "XX"
           .Col = .Cols - 3
            Codigo1 = Replace(Codigo1, ".", "X")
            Codigo2 = Replace(Codigo2, ".", "X")
            Codigo3 = Replace(Codigo3, ".", "X")
            Codigo = Mid(Codigo1, 1, 2)
            Codigo = Replace(Codigo, ".", "X")
            If Mid(Codigo2, 1, 3) = Mid(Codigo3, 1, 3) Then Codigo3 = "XXX"
            Codigo = Codigo & "." & Mid(Codigo2, 1, 3) & Mid(Codigo3, 1, 3) & "." & Format(I, "0000")
            'MsgBox Codigo
            If TextWidth(CStr(Codigo)) > .ColWidth(.Cols - 3) Then .ColWidth(.Cols - 3) = TextWidth(CStr(Codigo & "XX"))
           .Text = Codigo
           .Col = .Cols - 2
           .Text = Mid(Codigo1, 1, 2) & Mid(Codigo2, 1, 3) & Mid(Codigo3, 1, 3) & Format(I, "0000")
            End If
            Me.Caption = "(" & rango.NumCol2 & ") Importar Excel a FlexGrid " & I & " de " & rango.NumFila2
        Next
       .Row = 0
       .Col = 0
       .Text = "No"
        rango.NumFila2 = rango.NumFila2 - No_Bancos
       .Rows = rango.NumFila2
    End With
    MSHFlexGrid1.Visible = True
    obj_Workbook.Close
    obj_Excel.Quit
    Descargar
    RatonNormal
Exit Sub
error_sub:
MsgBox Err.Description
Descargar
End Sub

Private Sub Descargar()
    On Local Error Resume Next
    Set obj_Workbook = Nothing
    Set obj_Excel = Nothing
    Set obj_Worksheet = Nothing
    Me.MousePointer = vbDefault
End Sub

Private Sub Command2_Click()
  RatonReloj
  Importar_Abonos
  RatonNormal
  MsgBox "Proceso Terminado"
  Unload Me
End Sub

Private Sub Command3_Click()
 Unload FImporFA
End Sub

Private Sub CommandButton1_Click()
  CDialogDir.InitDir = "C:\"
  RutaOrigen = UCase(SelectZipFile(CDialogDir, SelectAll))
  If RutaOrigen <> "" Then
    'Le pasamos el Path del Libro y una variable de tipo T_Rango para retornar los valores
     Call Obtener_Rango_Excel(RutaOrigen, rango)
     Call Excel_FlexGrid(RutaOrigen, MSHFlexGrid1, rango)
  End If
End Sub

Private Sub Form_Activate()
  Trans_No = 199
  
  RatonNormal
  sSQL = "SELECT * " _
       & "FROM Catalogo_Lineas " _
       & "WHERE TL <> " & Val(adFalse) & " " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "ORDER BY Codigo "
   
   
End Sub

Private Sub Form_Load()
   Me.Caption = "Importar Excel a FlexGrid"
   CommandButton1.Caption = "Importar" & vbCrLf & "a" & vbCrLf & "Flexgrid"
   ConectarAdodc AdoAux
   ConectarAdodc AdoAct
   ConectarAdodc AdoLinea
   ConectarAdodc AdoAsiento
  
   MSHFlexGrid1.width = MDI_X_Max - 100
   MSHFlexGrid1.Height = (MDI_Y_Max - MSHFlexGrid1.Top - 600)

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Descargar
End Sub

Public Sub Importar_Abonos()
Dim I As Long
Dim N As Long
Dim Tot_Propinas As Currency
 
  Encerar_Facturas
  
  FechaTexto = FechaSistema
  Bandera = False
  Evaluar = True
  DiarioCaja = ReadSetDataNum("Recibo_No", True, True)
 'Empezamos la importacion de las facturas
  With MSHFlexGrid1
       'MsgBox .Rows & vbCrLf & .Cols
       For I = 1 To .Rows - 1
          .Row = I
           For N = 1 To .Cols - 1
              .Col = N
               Codigo = Trim(.Text)
               Codigo1 = Trim(.Text)
               TA.Banco = ""
               Select Case N
                 Case 1: TA.Fecha = Codigo
                 Case 2: TA.Factura = Val(Codigo)
                 Case 3: TA.Autorizacion = Codigo
                 Case 4: TA.Abono = Redondear(Val(Codigo), 2)
                 Case 5: TA.Banco = Codigo
                 Case 6: TA.Banco = TA.Banco & ", Cod. " & Codigo
                 Case 7: TA.Cheque = Codigo
               End Select
           Next N
           If TA.Banco = "" Then TA.Banco = Ninguno
           TA.Cta_CxP = Cta_Cobrar
           TA.CodigoC = CodigoCliente
           
''''    'Abono de Factura Caja MN
''''     TA.T = Normal
''''     TA.TP = TipoFactura
''''     TA.Fecha = MBFecha
''''     TA.Cta = Cta_CajaG
''''     TA.Cta_CxP = Cta_Cobrar
''''     TA.Banco = "EFECTIVO MN"
''''     TA.Cheque = Grupo_No
''''     TA.Factura = Factura_No
''''     TA.Abono = TotalCajaMN
''''     Grabar_Abonos TA
''''    'Abono de Factura Banco
''''     TA.T = Normal
''''     TA.TP = TipoFactura
''''     TA.Fecha = MBFecha
''''     TA.Cta = Trim(SinEspaciosIzq(DCBanco))
''''     TA.Cta_CxP = Cta_Cobrar
''''     TA.Banco = TextBanco
''''     TA.Cheque = TextCheqNo
''''     TA.Factura = Factura_No
''''     TA.Abono = Total_Bancos
''''     Grabar_Abonos TA
''''    'Abono de Factura Tarjeta
''''     TA.T = Normal
''''     TA.TP = TipoFactura
''''     TA.Fecha = MBFecha
''''     TA.Cta = Trim(SinEspaciosIzq(DCTarjeta))
''''     TA.Cta_CxP = Cta_Cobrar
''''     TA.Banco = NombreBanco1
''''     TA.Cheque = TextBaucher
''''     TA.Factura = Factura_No
''''     TA.Abono = Total_Tarjeta
''''     Grabar_Abonos TA
''''    'Abono de Factura Interes Tarjeta
''''     TA.T = Normal
''''     TA.TP = "TJ"
''''     TA.Fecha = MBFecha
''''     TA.Cta = Cta1
''''     TA.Cta_CxP = Cta_Tarjetas
''''     TA.Banco = "INTERES POR TARJETA"
''''     TA.Cheque = TextBaucher
''''     TA.Factura = Factura_No
''''     TA.Abono = Val(TextInteres)
''''     Grabar_Abonos TA
''''     Codigo1 = Format(Val(Trim(Mid(TxtSerieRet, 1, 3))), "000")
''''     Codigo2 = Format(Val(Trim(Mid(TxtSerieRet, 4, 3))), "000")
''''    'Abono de Factura Rete. IVA Bienes
''''     TA.T = Normal
''''     TA.TP = TipoFactura
''''     TA.Fecha = MBFecha
''''     TA.Cta = Trim(SinEspaciosIzq(DCRetIBienes))
''''     TA.Cta_CxP = Cta_Cobrar
''''     TA.Banco = "RETENCION IVA BIENES"
''''     TA.Cheque = TextCompRet
''''     TA.Factura = Factura_No
''''     TA.Abono = Total_RetIVAB
''''     TA.AutorizacionR = TxtAutoRet
''''     TA.Establecimiento = Codigo1
''''     TA.Emision = Codigo2
''''     TA.Porcentaje = Val(CBienes)
''''     Grabar_Abonos_Retenciones TA
''''    'Abono de Factura Ret IVA Servicio
''''     TA.T = Normal
''''     TA.TP = TipoFactura
''''     TA.Fecha = MBFecha
''''     TA.Cta = Trim(SinEspaciosIzq(DCRetISer))
''''     TA.Cta_CxP = Cta_Cobrar
''''     TA.Banco = "RETENCION IVA SERVICIO"
''''     TA.Cheque = TextCompRet
''''     TA.Factura = Factura_No
''''     TA.Abono = Total_RetIVAS
''''     TA.AutorizacionR = TxtAutoRet
''''     TA.Establecimiento = Codigo1
''''     TA.Emision = Codigo2
''''     TA.Porcentaje = Val(CServicio)
''''     Grabar_Abonos_Retenciones TA
''''    'Abono de Factura Ret. Fuente
''''     TA.T = Normal
''''     TA.TP = TipoFactura
''''     TA.Fecha = MBFecha
''''     TA.Cta = Trim(SinEspaciosIzq(DCRetFuente))
''''     TA.Cta_CxP = Cta_Cobrar
''''     TA.Banco = "RETENCION FUENTE - " & DCCodRet
''''     TA.Cheque = TextCompRet
''''     TA.Factura = Factura_No
''''     TA.Abono = Total_Ret
''''     TA.AutorizacionR = TxtAutoRet
''''     TA.Establecimiento = Codigo1
''''     TA.Emision = Codigo2
''''     TA.Porcentaje = Val(TextPorc)
''''     Grabar_Abonos_Retenciones TA
     
     'MsgBox TipoFactura
     
     T = "P"
     If SaldoDisp <= 0 Then
        T = "C"
        SaldoDisp = 0
     End If
     sSQL = "UPDATE Facturas " _
          & "SET Saldo_MN = " & SaldoDisp & ",T = '" & T & "' " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Factura = " & Factura_No & " " _
          & "AND TC = '" & TipoFactura & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "AND CodigoC = '" & CodigoCliente & "' "
     ConectarAdoExecute sSQL


                 Actualizar_Facturas "Facturas", Mifecha
                 Actualizar_Facturas "Detalle_Factura", Mifecha
           Me.Caption = "Importar de FlexGrid a Sistema de Facturacion El Numero: " & FA.Factura & ": " & I & " de " & rango.NumFila2
           'MsgBox "..."
      Next I
  End With
End Sub

Public Sub Eliminar_Abonos(Mi_Fecha As String)
   sSQL = "DELETE * " _
        & "FROM Trans_Abonos " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND Periodo = '" & Periodo_Contable & "' " _
        & "AND TP = '" & FA.TC & "' " _
        & "AND Fecha = #" & BuscarFecha(Mi_Fecha) & "# "
   ConectarAdoExecute sSQL
End Sub

Public Sub Actualizar_Facturas(Nom_Tabla As String, Mi_Fecha As String)
   sSQL = "UPDATE " & Nom_Tabla & " " _
        & "SET T = 'C' " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND Periodo = '" & Periodo_Contable & "' " _
        & "AND Fecha = #" & BuscarFecha(Mi_Fecha) & "# " _
        & "AND T <> 'C' "
   ConectarAdoExecute sSQL
End Sub

