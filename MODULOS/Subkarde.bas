Attribute VB_Name = "SubKadex"

Public Sub Stock_Actual_Inventario(Fecha_Inv As String, _
                                   Codigo_Inventario As String, _
                                   Optional Por_Bodega As String)
Dim AdoStock As ADODB.Recordset
Dim CantBodegas As Integer
  CantBodegas = 0
  Cantidad = 0
  ValorUnit = 0
  SaldoAnterior = 0
  If Len(Fecha_Inv) < 10 Then Fecha_Inv = FechaSistema
  If Len(Codigo_Inventario) <= 1 Then
     Codigo_Inventario = Ninguno
  Else
     If Len(Por_Bodega) > 1 Then
        sSQL = "SELECT Codigo_Inv,AVG(Costo) As TCosto,SUM(Entrada-Salida) As Stock,(AVG(Costo)*SUM(Entrada-Salida)) As Saldo_Inv " _
             & "FROM Trans_Kardex " _
             & "WHERE Fecha <= #" & BuscarFecha(Fecha_Inv) & "# " _
             & "AND Codigo_Inv = '" & Codigo_Inventario & "' " _
             & "AND CodBodega = '" & Por_Bodega & "' " _
             & "AND Item = '" & NumEmpresa & "' " _
             & "AND Periodo = '" & Periodo_Contable & "' " _
             & "AND T <> 'A' " _
             & "AND TP <> '" & Ninguno & "' " _
             & "AND Numero > 0 " _
             & "GROUP BY Codigo_Inv "
     Else
        sSQL = "SELECT TOP 1 Codigo_Inv,Costo As TCosto,Existencia As Stock,Total As Saldo_Inv " _
             & "FROM Trans_Kardex " _
             & "WHERE Fecha <= #" & BuscarFecha(Fecha_Inv) & "# " _
             & "AND Codigo_Inv = '" & Codigo_Inventario & "' " _
             & "AND Item = '" & NumEmpresa & "' " _
             & "AND Periodo = '" & Periodo_Contable & "' " _
             & "AND T <> 'A' " _
             & "AND TP <> '" & Ninguno & "' " _
             & "AND Numero > 0 " _
             & "ORDER BY Fecha DESC,TP DESC, Numero DESC,ID DESC "
    End If
    Select_AdoDB AdoStock, sSQL
    With AdoStock
     If .RecordCount > 0 Then
         Cantidad = .Fields("Stock")
         ValorUnit = Redondear(.Fields("TCosto"), Dec_Costo)
         SaldoAnterior = .Fields("Saldo_Inv")
     End If
    End With
  End If
End Sub

Public Sub EncabezadoKardex(Datas As Adodc)
Dim InicX As Single
Dim InicY As Single
Encabezado Ancho(0), Ancho(11)
PorteLetra = Printer.FontSize
LetraAnterior = Printer.FontName
Printer.FontName = TipoTimes
Printer.FontSize = 18: Printer.FontBold = True
SQLMsg1 = "CONTROL DE EXISTENCIAS"
PrinterTexto CentrarTexto(SQLMsg1), PosLinea + 0.1, SQLMsg1
PosLinea = PosLinea + 0.9
'========================================================================
'Dibujo = RutaSistema & "\FORMATOS\KARDEX.GIF"
'PrinterPaint Dibujo, Ancho(0), PosLinea, 25.5, 2
Printer.FontSize = 10
Printer.FontBold = False
PrinterFields 1.5, PosLinea + 0.1, Datas.Recordset.Fields("Codigo_Inv"), False
PrinterFields 1.5, PosLinea + 0.6, Datas.Recordset.Fields("Producto"), False
PrinterFields 5.5, PosLinea + 0.1, Datas.Recordset.Fields("Unidad"), False
PrinterFields 12, PosLinea + 0.1, Datas.Recordset.Fields("Minimo"), False
PrinterFields 12, PosLinea + 0.5, Datas.Recordset.Fields("Maximo"), False
'PrinterFields 17.5, PosLinea + 0.1, Datas.Recordset.Fields("Proveedor"), True, False
'PrinterFields 17, PosLinea + 0.5, Datas.Recordset.Fields("Bodega"), False
Printer.FontBold = True
PrinterTexto Ancho(0), PosLinea + 1.5, "Bod."
PrinterTexto Ancho(1), PosLinea + 1.5, "Fecha"
PrinterTexto Ancho(2), PosLinea + 1.5, "TP"
PrinterTexto Ancho(3), PosLinea + 1.5, "Numero"
PrinterTexto Ancho(4), PosLinea + 1.5, "Detalle"
PrinterTexto Ancho(5), PosLinea + 1.5, "Entrada"
PrinterTexto Ancho(6), PosLinea + 1.5, "Salida"
PrinterTexto Ancho(7), PosLinea + 1.5, "Valor_Unit"
PrinterTexto Ancho(8), PosLinea + 1.5, "Valor_Total"
PrinterTexto Ancho(9), PosLinea + 1.5, "Stock Act."
PrinterTexto Ancho(10), PosLinea + 1.5, "Costo_Prom"
PrinterTexto Ancho(11), PosLinea + 1.5, "Saldo_Total"
Printer.Line (InicioX, PosLinea + 1.9)-(Ancho(CantCampos), PosLinea + 1.9), QBColor(0)
Printer.FontSize = PorteLetra
Printer.FontName = LetraAnterior
PosLinea = PosLinea + 2.05
Printer.FontBold = False
End Sub

Public Sub EncabezadoKardexCantidad(Datas As Adodc)
Dim InicX As Single
Dim InicY As Single
Encabezado Ancho(0), Ancho(10)
PorteLetra = Printer.FontSize
LetraAnterior = Printer.FontName
Printer.FontName = TipoTimes
Printer.FontSize = 18: Printer.FontBold = True
SQLMsg1 = "CONTROL DE EXITENCIAS"
PrinterTexto CentrarTexto(SQLMsg1), PosLinea + 0.1, SQLMsg1
PosLinea = PosLinea + 0.9
'========================================================================
Printer.FontSize = 10: Printer.FontBold = True
With Datas.Recordset
PrinterFields 3, PosLinea + 0.1, .Fields("Producto"), False
PrinterFields 2.5, PosLinea + 0.5, .Fields("Unidad"), False
PrinterFields 7.5, PosLinea + 0.5, .Fields("Codigo_Inv"), False
PrinterFields 12, PosLinea + 0.1, .Fields("Minimo"), False
PrinterFields 12, PosLinea + 0.5, .Fields("Maximo"), False
PrinterFields 17, PosLinea + 0.5, .Fields("Bodega"), False
End With
PrinterTexto Ancho(0), PosLinea + 1, "Fecha"
PrinterTexto Ancho(1), PosLinea + 1, "Comp_No"
PrinterTexto Ancho(2), PosLinea + 1, "D e t a l l e"
PrinterTexto Ancho(7), PosLinea + 1, "Entrada"
PrinterTexto Ancho(8), PosLinea + 1, "Salida"
PrinterTexto Ancho(9), PosLinea + 1, "Stock"
PosLinea = PosLinea + 1.5
Printer.Line (Ancho(0), PosLinea)-(Ancho(10), PosLinea), QBColor(0)
PosLinea = PosLinea + 0.5
Printer.Line (Ancho(0), PosLinea)-(Ancho(10), PosLinea), QBColor(0)
PosLinea = PosLinea + 0.1
Printer.FontBold = False
Printer.FontSize = PorteLetra
Printer.FontName = LetraAnterior
End Sub

Public Sub EncabNotaInv(DtaProv As Adodc, Numero As Long, OrdenPNo As String, Det_Ctas() As String, Cont_Ctas As Long, TipoDoc As String)
Dim InicX As Single
Dim InicY As Single
Encabezado Ancho(0), 19.5
PorteLetra = Printer.FontSize
LetraAnterior = Printer.FontName
Printer.FontItalic = False
Printer.FontName = TipoArial
Printer.FontBold = True
SQLMsg1 = "NOTA DE ENTRADA/SALIDA DE INVENTARIO"
Cadena = "TP: " & TipoDoc & ",  No. " & Format$(Numero, "000000")
Printer.FontSize = 14
PrinterTexto CentrarTexto(SQLMsg1), PosLinea, SQLMsg1
PosLinea = PosLinea + 0.8
Printer.FontSize = 10
Cadena1 = Ninguno
With DtaProv.Recordset
 If .RecordCount > 0 Then
     PrinterTexto 2, PosLinea, "FECHA: " & .Fields("Fecha")
     PrinterTexto 14, PosLinea, Cadena
     PosLinea = PosLinea + 0.5
     PrinterTexto 2, PosLinea, "BENEFICIARIO:"
     Printer.FontBold = False
     PrinterTexto 5, PosLinea, .Fields("Cliente")
     Cadena1 = .Fields("Concepto")
 Else
     PrinterTexto 2, PosLinea, "FECHA: " & FechaTexto
     PrinterTexto 14, PosLinea, Cadena
     PosLinea = PosLinea + 0.5
     PrinterTexto 2, PosLinea, "BENEFICIARIO:"
     Printer.FontBold = False
     PrinterTexto 5, PosLinea, NombreCliente
 End If
End With
'========================================================================
Printer.FontSize = 9
Printer.FontBold = True
PosLinea = PosLinea + 0.5
Printer.Line (Ancho(0), PosLinea)-(19.5, PosLinea), QBColor(0)
PosLinea = PosLinea + 0.05
Printer.FontBold = True
PrinterTexto 2, PosLinea, "DETALLES:"
Printer.FontBold = False
NumeroLineas = PrinterLineasMayor(4, PosLinea, Cadena1, 15.5)
PosLinea = PosLinea + (NumeroLineas * 0.4)
'PrinterTexto 2, PosLinea, Cadena1
Printer.FontBold = True
PrinterTexto 2, PosLinea, "CUENTAS INVOLUCRADAS:"
Printer.FontBold = False
For I = 0 To Cont_Ctas
    PrinterTexto 6.5, PosLinea, Det_Ctas(I)
    PosLinea = PosLinea + 0.4
Next I
Printer.FontBold = True
Printer.Line (Ancho(0), PosLinea)-(19.5, PosLinea), QBColor(0)
PosLinea = PosLinea + 0.05
PrinterTexto 1, PosLinea, "BOD."
PrinterTexto 2, PosLinea, "CODIGO/SERIE"
PrinterTexto 5.5, PosLinea, "CODIGO"
PrinterTexto 8, PosLinea, "PRODUCTO"
PrinterTexto 15.5, PosLinea, "ENTRADA"
PrinterTexto 18, PosLinea, "SALIDA"
PosLinea = PosLinea + 0.4
Printer.Line (Ancho(0), PosLinea)-(19.5, PosLinea), QBColor(0)
PosLinea = PosLinea + 0.1
Printer.FontBold = False
Printer.FontSize = PorteLetra
Printer.FontName = LetraAnterior
End Sub

Public Sub Imprimir_Nota_Inventario(DtaProv As Adodc, _
                                    Datas As Adodc, _
                                    Numero As Long, _
                                    OrdenNo As String, _
                                    OpcPrint As String, _
                                    SFechaI As String, _
                                    SFechaF As String, _
                                    Total_Inv As Currency)
Dim AdoDBKardex As ADODB.Recordset
Dim FormaImp As Byte
Dim SizeLetra As Integer
Dim Detalle_Ctas(20) As String
Dim Lote_No As String
Dim Fecha_Exp As String
Dim Fecha_Fab As String
Dim Reg_Sanitario As String
Dim Procedencia As String

On Error GoTo Errorhandler
'IMPRIMIR S/N
Mensajes = "Desea imprimir Nota de Inventario"
Titulo = "IMPRESION"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
  RatonReloj
  Select Case OpcPrint
    Case "R"
         Codigo = SinEspaciosIzq(OrdenNo)
         TextoValidoVar Codigo
    Case "CD"
         Codigo = CStr(Numero)
         TextoValidoVar Codigo, True
    Case "NC"
         Codigo = CStr(Numero)
         TextoValidoVar Codigo, True
    Case "G"
         Codigo = OrdenNo
         TextoValidoVar Codigo
    Case "B"
         Codigo = OrdenNo
         TextoValidoVar Codigo
  End Select
  FechaIni = BuscarFecha(SFechaI)
  FechaFin = BuscarFecha(SFechaF)
  sSQL = "SELECT Co.Fecha,P.Producto,K.TP,K.Numero,C.Cliente,C.CI_RUC,C.Direccion,C.Telefono," _
       & "Lote_No,Fecha_Exp,Fecha_Fab,P.Reg_Sanitario,Modelo,Serie_No,Procedencia,Concepto,Entrada,Salida " _
       & "FROM Trans_Kardex As K,Catalogo_Productos As P,Comprobantes As Co,Clientes As C " _
       & "WHERE K.Item = '" & NumEmpresa & "' " _
       & "AND K.Periodo = '" & Periodo_Contable & "' "
  Select Case OpcPrint
    Case "R"
         sSQL = sSQL & "AND K.Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
              & "AND K.Codigo_P = '" & Codigo & "' "
    Case "CD"
         sSQL = sSQL & "AND K.TP = 'CD' " _
              & "AND K.Numero = " & Val(Codigo) & " "
    Case "NC"
         sSQL = sSQL & "AND K.TP = 'NC' " _
              & "AND K.Numero = " & Val(Codigo) & " "
    Case "G"
         sSQL = sSQL & "AND K.Orden_No = '" & Codigo & "' "
    Case "B"
         sSQL = sSQL & "AND K.Codigo_Barra = '" & Codigo & "' "
  End Select
  sSQL = sSQL & "AND K.Codigo_Inv = P.Codigo_Inv " _
       & "AND K.TP = Co.TP " _
       & "AND K.Numero = Co.Numero " _
       & "AND K.Item = Co.Item " _
       & "AND K.Item = P.Item " _
       & "AND K.Periodo = Co.Periodo  " _
       & "AND K.Periodo = P.Periodo  " _
       & "AND C.Codigo = Co.Codigo_B " _
       & " " _
       & "ORDER BY P.Producto,K.Fecha,K.TP,K.Numero,K.ID "
  Select_Adodc DtaProv, sSQL
 'MsgBox sSQL & vbCrLf & DtaProv.Recordset.RecordCount
  With DtaProv.Recordset
   If .RecordCount > 0 Then
       'MsgBox .RecordCount & vbCrLf & .Fields("Cliente")
       FechaTexto = .Fields("Fecha")
       Numero = .Fields("Numero")
       NombreCliente = .Fields("Cliente")
   End If
  End With
  'MsgBox DtaProv.Recordset.RecordCount
  
  sSQL = "SELECT K.Orden_No,K.Codigo_Barra,K.CodBodega,K.Codigo_Inv,P.Producto,K.Fecha,K.TP,K.Numero,Entrada,Salida," _
       & "Lote_No,Fecha_Exp,Fecha_Fab,P.Reg_Sanitario,Modelo,Serie_No,Procedencia " _
       & "FROM Trans_Kardex As K,Catalogo_Productos As P " _
       & "WHERE K.Item = '" & NumEmpresa & "' " _
       & "AND K.Periodo = '" & Periodo_Contable & "' "
  Select Case OpcPrint
    Case "R"
         sSQL = sSQL & "AND K.Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
              & "AND Codigo_P = '" & Codigo & "' "
    Case "CD"
         sSQL = sSQL & "AND K.TP = 'CD' " _
              & "AND K.Numero = " & Val(Codigo) & " "
    Case "NC"
         sSQL = sSQL & "AND K.TP = 'NC' " _
              & "AND K.Numero = " & Val(Codigo) & " "
    Case "G"
         sSQL = sSQL & "AND K.Orden_No = '" & Codigo & "' "
    Case "B"
         sSQL = sSQL & "AND K.Codigo_Barra = '" & Codigo & "' "
  End Select
  sSQL = sSQL & "AND K.Codigo_Inv = P.Codigo_Inv " _
       & "AND K.Item = P.Item " _
       & "AND K.Periodo = P.Periodo  " _
       & "ORDER BY K.Orden_No,K.Codigo_Inv,K.Fecha,K.TP,K.Numero,K.ID "
  Select_Adodc Datas, sSQL
' Cuentas Involucradas
  Set AdoDBKardex = New ADODB.Recordset
  AdoDBKardex.CursorType = adOpenStatic
  AdoDBKardex.CursorLocation = adUseClient
  sSQL = "SELECT K.Cta_Inv As Ctas ,C.Cuenta " _
       & "FROM Trans_Kardex As K,Catalogo_Cuentas As C " _
       & "WHERE K.Item = '" & NumEmpresa & "' " _
       & "AND K.Periodo = '" & Periodo_Contable & "' "
  Select Case OpcPrint
    Case "R"
         sSQL = sSQL & "AND K.Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# "
    Case "CD"
         sSQL = sSQL & "AND K.TP = 'CD' " _
              & "AND K.Numero = " & Val(Codigo) & " "
    Case "NC"
         sSQL = sSQL & "AND K.TP = 'NC' " _
              & "AND K.Numero = " & Val(Codigo) & " "
    Case "G"
         sSQL = sSQL & "AND K.Orden_No = '" & Codigo & "' "
    Case "B"
         sSQL = sSQL & "AND K.Codigo_Barra = '" & Codigo & "' "
  End Select
  sSQL = sSQL & "AND K.Cta_Inv = C.Codigo " _
       & "AND K.Item = C.Item " _
       & "AND K.Periodo = C.Periodo  " _
       & "GROUP BY K.Cta_Inv,C.Cuenta " _
       & "UNION " _
       & "SELECT K.Contra_Cta AS Ctas,C.Cuenta " _
       & "FROM Trans_Kardex As K,Catalogo_Cuentas As C " _
       & "WHERE K.Item = '" & NumEmpresa & "' " _
       & "AND K.Periodo = '" & Periodo_Contable & "' "
  Select Case OpcPrint
    Case "R"
         sSQL = sSQL & "AND K.Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# "
    Case "CD"
         sSQL = sSQL & "AND K.TP = 'CD' " _
              & "AND K.Numero = " & Val(Codigo) & " "
    Case "NC"
         sSQL = sSQL & "AND K.TP = 'NC' " _
              & "AND K.Numero = " & Val(Codigo) & " "
    Case "G"
         sSQL = sSQL & "AND K.Orden_No = '" & Codigo & "' "
    Case "B"
         sSQL = sSQL & "AND K.Codigo_Barra = '" & Codigo & "' "
  End Select
  sSQL = sSQL & "AND K.Contra_Cta = C.Codigo " _
       & "AND K.Item = C.Item " _
       & "AND K.Periodo = C.Periodo  " _
       & "GROUP BY K.Contra_Cta,C.Cuenta "
  sSQL = CompilarSQL(sSQL)
  'MsgBox sSQL
  AdoDBKardex.Open sSQL, AdoStrCnn, , , adCmdText
 'Seteamos los encabezados para las facturas
  Contador = 0
  If AdoDBKardex.RecordCount > 0 Then
     Do While Not AdoDBKardex.EOF
        Detalle_Ctas(Contador) = AdoDBKardex.Fields("Cuenta")
        Contador = Contador + 1
        AdoDBKardex.MoveNext
     Loop
  End If
InicioX = 0.5: InicioY = 0: FormaImp = 1: SizeLetra = 12
DataAnchoCampos InicioX, Datas, SizeLetra, TipoArial, FormaImp
CantCampos = 6
Pagina = 1
Entrada = 0
Salida = 0
With Datas.Recordset
 If .RecordCount > 0 Then
    .MoveFirst
     Do While Not .EOF
        Entrada = Entrada + .Fields("Entrada")
        Salida = Salida + .Fields("Salida")
       .MoveNext
     Loop
 End If
End With

Mensajes = "En Impresora Pequeña?" & vbCrLf
Titulo = "Formulario de Impresion"
If BoxMensaje = vbYes Then
   RatonReloj
   PosLinea = 0.01
   SizeLetra = 7
   Printer.FontName = TipoArial
   PrinterPaint LogoTipo, 0.01, PosLinea, 2.5, 1      'Ancho = 4.7
   Printer.FontSize = SizeLetra + 2
   PrinterTexto 3.5, PosLinea, "R.U.C."
   PosLinea = PosLinea + 0.36
   PrinterTexto 3.5, PosLinea, RUC
   PosLinea = PosLinea + 0.36
   Printer.FontSize = SizeLetra
   PrinterTexto 3.5, PosLinea, "Teléfono: " & Telefono1
   PosLinea = PosLinea + 0.36
   Printer.FontSize = SizeLetra + 5
   Printer.FontBold = True
   PrinterCentrarTexto 7, PosLinea, Empresa
   Printer.FontBold = False
   PosLinea = PosLinea + 0.7
   If Len(NombreComercial) > 1 And Dir(LogoTipo) = "" Then
      Printer.FontSize = SizeLetra + 3
      PrinterCentrarTexto 7, PosLinea, NombreComercial
      PosLinea = PosLinea + 0.36
   End If
   Printer.FontBold = True
   Printer.FontSize = SizeLetra + 3
   If Entrada > 0 Then PrinterCentrarTexto 7, PosLinea, "NOTA DE ENTRADA"
   If Salida > 0 Then PrinterCentrarTexto 7, PosLinea, "NOTA DE SALIDA"
   If Entrada > 0 And Salida > 0 Then PrinterCentrarTexto 7, PosLinea, "NOTA DE ENTRADA Y SALIDA"
   Printer.FontBold = False
   PosLinea = PosLinea + 0.6
   Printer.FontSize = SizeLetra + 1
   
   Cadena = "TP: " & OpcPrint & ", No. " & Format$(Numero, "000000")
   
   PrinterTexto 0.01, PosLinea, "HORA DE IMPRESION DOCUMENTO: " & Time
   PosLinea = PosLinea + 0.36
   Printer.FontSize = SizeLetra + 2
   Cadena1 = Ninguno
   With DtaProv.Recordset
    If .RecordCount > 0 Then
        PrinterTexto 0.01, PosLinea, "FECHA: " & .Fields("Fecha")
        PrinterTexto 3.4, PosLinea, Cadena
        PosLinea = PosLinea + 0.46
        Printer.FontSize = SizeLetra
        PrinterTexto 0.01, PosLinea, "BENEFICIARIO:"
        Printer.FontBold = False
        PrinterTexto 2.1, PosLinea, .Fields("Cliente")
        Cadena1 = .Fields("Concepto")
    Else
        PrinterTexto 0.01, PosLinea, "FECHA: " & FechaTexto
        PrinterTexto 3.4, PosLinea, Cadena
        Printer.FontSize = SizeLetra
        PosLinea = PosLinea + 0.36
        'PrinterTexto 0.01, PosLinea, "BENEFICIARIO:"
        Printer.FontBold = False
        PrinterTexto 3, PosLinea, NombreCliente
    End If
   End With
   Printer.FontSize = SizeLetra + 3
   PosLinea = PosLinea + 0.36
   Printer.FontBold = True
   PrinterTexto 0.01, PosLinea, "CUENTAS INVOLUCRADAS:"
   PosLinea = PosLinea + 0.4
   Printer.FontBold = False
   For I = 0 To Contador
       PrinterTexto 0.5, PosLinea, Detalle_Ctas(I)
       PosLinea = PosLinea + 0.4
   Next I
   
   With Datas.Recordset
    If .RecordCount > 0 Then
       .MoveFirst
        Printer.FontSize = SizeLetra + 1
        PrinterTexto 0.01, PosLinea, Cadena1
        PosLinea = PosLinea + 0.36
        Printer.FontSize = SizeLetra
        Lote_No = .Fields("Lote_No")
        Fecha_Exp = .Fields("Fecha_Exp")
        Fecha_Fab = .Fields("Fecha_Fab")
        Reg_Sanitario = "Reg. Sanitario: " & .Fields("Reg_Sanitario") & ", Procedencia: " & .Fields("Procedencia")
        Procedencia = "Modelo: " & .Fields("Modelo") & ", Serie No. " & .Fields("Serie_No")
        If Len(Lote_No) > 1 Then
           PrinterTexto 0.01, PosLinea, String(80, "-")
           PosLinea = PosLinea + 0.2
           PrinterTexto 0.01, PosLinea, "LOTE No. " & Lote_No & UCaseStrg("          FAB: " & Format(Fecha_Fab, "MMM-yyyy") & "          EXP: " & Format(Fecha_Exp, "MMM-yyyy"))
           PosLinea = PosLinea + 0.36
           PrinterTexto 0.01, PosLinea, Reg_Sanitario
           PosLinea = PosLinea + 0.36
           PrinterTexto 0.01, PosLinea, Procedencia
           PosLinea = PosLinea + 0.36
        End If
        PrinterTexto 0.01, PosLinea, String(80, "-")
        PosLinea = PosLinea + 0.2
        Printer.FontBold = True
        PrinterTexto 0.01, PosLinea, "CANT."
        PrinterTexto 1.1, PosLinea, "CODIGO"
        PrinterTexto 3, PosLinea, "P R O D U C T O"
        PosLinea = PosLinea + 0.26
        Printer.FontBold = False
        PrinterTexto 0.01, PosLinea, String(80, "-")
        PosLinea = PosLinea + 0.3
        
        Do While Not .EOF
           If Lote_No <> .Fields("Lote_No") Then
              Lote_No = .Fields("Lote_No")
              Fecha_Exp = .Fields("Fecha_Exp")
              Fecha_Fab = .Fields("Fecha_Fab")
              Reg_Sanitario = "Reg. Sanitario: " & .Fields("Reg_Sanitario") & ", Procedencia: " & .Fields("Procedencia")
              Procedencia = "Modelo: " & .Fields("Modelo") & ", Serie No. " & .Fields("Serie_No")
              If Len(Lote_No) > 1 Then
                 PrinterTexto 0.01, PosLinea, String(80, "-")
                 PosLinea = PosLinea + 0.2
                 PrinterTexto 0.01, PosLinea, "LOTE No. " & Lote_No & UCaseStrg("          FAB: " & Format(Fecha_Fab, "MMM-yyyy") & "          EXP: " & Format(Fecha_Exp, "MMM-yyyy"))
                 PosLinea = PosLinea + 0.36
                 PrinterTexto 0.01, PosLinea, Reg_Sanitario
                 PosLinea = PosLinea + 0.36
                 PrinterTexto 0.01, PosLinea, Procedencia
                 PosLinea = PosLinea + 0.36
                 PrinterTexto 0.01, PosLinea, String(80, "-")
                 PosLinea = PosLinea + 0.2
              End If
           End If
           If .Fields("Entrada") > 0 Then PrinterTexto 0.01, PosLinea, .Fields("Entrada")
           If .Fields("Salida") > 0 Then PrinterTexto 0.01, PosLinea, .Fields("Salida")
           PrinterTexto 1.1, PosLinea, .Fields("Codigo_Inv")
           PrinterTexto 3, PosLinea, .Fields("Producto")
           PosLinea = PosLinea + 0.36
          .MoveNext
        Loop
    End If
   End With
   Printer.FontSize = SizeLetra
   PrinterTexto 0.01, PosLinea, String(80, "-")
   PosLinea = PosLinea + 1.5
   Printer.FontSize = SizeLetra
   PrinterTexto 0.01, PosLinea, String(20, "_")
   PrinterTexto 4, PosLinea, String(20, "_")
   PrinterTexto 0.01, PosLinea + 0.5, "ENTREGUE CONFORME"
   PrinterTexto 4, PosLinea + 0.5, "  RECIBI CONFORME"
   PosLinea = PosLinea + 2
   PrinterTexto 4, PosLinea + 0.5, "_"
Else
    RatonReloj
    SizeLetra = 9
    Printer.FontName = TipoArial
   'Iniciamos la impresion
    Printer.FontBold = False
    With Datas.Recordset
     If .RecordCount > 0 Then
         .MoveFirst
          Lote_No = .Fields("Lote_No")
          Fecha_Exp = .Fields("Fecha_Exp")
          Fecha_Fab = .Fields("Fecha_Fab")
          Reg_Sanitario = .Fields("Reg_Sanitario")
          Procedencia = "Procedencia: " & .Fields("Procedencia") & ", Modelo: " & .Fields("Modelo") & ", Serie No. " & .Fields("Serie_No")
          EncabNotaInv DtaProv, Numero, OrdenNo, Detalle_Ctas, Contador, OpcPrint
          Printer.FontSize = SizeLetra
          Grupo_No = .Fields("Orden_No")
          Printer.FontBold = True
          PrinterTexto 2, PosLinea, "Orden No. "
          Printer.FontBold = False
          PrinterTexto 3.7, PosLinea, Grupo_No
          
          If Len(Lote_No) > 1 Then
             Cadena = "LOTE No. " & Lote_No _
                    & ", FAB: " & UCaseStrg(Format(Fecha_Fab, "MM-yyyy")) _
                    & ", EXP: " & UCaseStrg(Format(Fecha_Exp, "MM-yyyy")) _
                    & ", REG. SANITARIO: " & Reg_Sanitario
             PrinterTexto 5.45, PosLinea, Cadena
             PosLinea = PosLinea + 0.4
             PrinterTexto 5.45, PosLinea, Procedencia
          End If
          PosLinea = PosLinea + 0.4
          
          Do While Not .EOF
             If Grupo_No <> .Fields("Orden_No") Then
                Printer.FontBold = True
                PrinterTexto 2, PosLinea, "Orden No. "
                Printer.FontBold = False
                PrinterTexto 3.5, PosLinea, Grupo_No
                PosLinea = PosLinea + 0.4
                Grupo_No = .Fields("Orden_No")
             End If
             If Lote_No <> .Fields("Lote_No") Then
                Lote_No = .Fields("Lote_No")
                Fecha_Exp = .Fields("Fecha_Exp")
                Fecha_Fab = .Fields("Fecha_Fab")
                Reg_Sanitario = .Fields("Reg_Sanitario")
                Procedencia = "Procedencia: " & .Fields("Procedencia") & ", Modelo: " & .Fields("Modelo") & ", Serie No. " & .Fields("Serie_No")
                If Len(Lote_No) > 1 Then
                   Cadena = "LOTE No. " & Lote_No _
                          & ", FAB: " & UCaseStrg(Format(Fecha_Fab, "MM-yyyy")) _
                          & ", EXP: " & UCaseStrg(Format(Fecha_Exp, "MM-yyyy")) _
                          & ", REG. SANITARIO: " & Reg_Sanitario
                   PrinterTexto 5.45, PosLinea, Cadena
                   PosLinea = PosLinea + 0.4
                   PrinterTexto 5.45, PosLinea, Procedencia
                   PosLinea = PosLinea + 0.4
                End If
             End If
             PrinterFields 1, PosLinea, .Fields("CodBodega"), False
             PrinterFields 2, PosLinea, .Fields("Codigo_Barra")
             PrinterFields 5.5, PosLinea, .Fields("Codigo_Inv"), False
             PrinterFields 8, PosLinea, .Fields("Producto"), False
             PrinterFields 15, PosLinea, .Fields("Entrada"), False
             PrinterFields 17.5, PosLinea, .Fields("Salida"), False
             PosLinea = PosLinea + 0.4
             If PosLinea >= LimiteAlto Then
                'Printer.Line (Ancho(0), PosLinea)-(Ancho(CantCampos), PosLinea), QBColor(0)
                Printer.NewPage
                EncabNotaInv DtaProv, Numero, OrdenNo, Detalle_Ctas, Contador, OpcPrint
                Printer.FontSize = SizeLetra
             End If
            .MoveNext
          Loop
      End If
    End With
    PosLinea = PosLinea + 0.1
    ''Printer.Line (InicioX, PosLinea)-(19, PosLinea), QBColor(0)
    ''PosLinea = PosLinea + 0.05
    ''Printer.Line (InicioX, PosLinea)-(19, PosLinea), QBColor(0)
    PosLinea = PosLinea + 0.05
    If Total_Inv > 0 Then
       PrinterTexto 1, PosLinea, "TOTAL INVENTARIO"
       PrinterVariables 6, PosLinea, Total_Inv
       PosLinea = PosLinea + 0.4
    End If
    PosLinea = PosLinea + 1.5
    PrinterTexto 3.5, PosLinea, String(20, "_")
    PrinterTexto 8, PosLinea, String(20, "_")
    PrinterTexto 14, PosLinea, String(20, "_")
    
    PrinterTexto 4, PosLinea + 0.5, "AUTORIZADO"
    PrinterTexto 8, PosLinea + 0.5, "ENTREGUE CONFORME"
    PrinterTexto 14, PosLinea + 0.5, "  RECIBI CONFORME"
    End If
    RatonNormal
    MensajeEncabData = ""
    Printer.EndDoc
End If
Exit Sub
Errorhandler:
    RatonNormal
    ErrorDeImpresion
    Exit Sub
End Sub

Public Sub ImprimirNotaEntradaSalida(DtaProv As Adodc, _
                                     Datas As Adodc, _
                                     Numero As Long, _
                                     OrdenNo As String, _
                                     OpcPrint As String, _
                                     SFechaI As String, _
                                     SFechaF As String)
Dim FormaImp As Byte
Dim SizeLetra As Integer
Dim xx(0) As String
xx(0) = ""
On Error GoTo Errorhandler
'IMPRIMIR S/N
Mensajes = "Desea imprimir Nota de Entrada/Salida"
Titulo = "IMPRESION"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
  RatonReloj
  Select Case OpcPrint
    Case "R"
         Codigo = SinEspaciosIzq(OrdenNo)
         TextoValidoVar Codigo
    Case "BO"
         Codigo = CStr(Numero)
         TextoValidoVar Codigo, True
    Case "G"
         Codigo = OrdenNo
         TextoValidoVar Codigo
  End Select
  FechaIni = BuscarFecha(SFechaI)
  FechaFin = BuscarFecha(SFechaF)
  sSQL = "SELECT K.Fecha,P.Producto,K.TP,K.Numero,C.Cliente,C.CI_RUC,C.Direccion,C.Telefono,Entrada,Salida " _
       & "FROM Trans_Kardex As K,Catalogo_Productos As P,Clientes As C " _
       & "WHERE K.Item = '" & NumEmpresa & "' "
  Select Case OpcPrint
    Case "R"
         sSQL = sSQL & "AND K.Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
              & "AND K.Codigo_P = '" & Codigo & "' "
    Case "BO"
         sSQL = sSQL & "AND K.TP = 'BO' " _
              & "AND K.Numero = " & Val(Codigo) & " "
    Case "G"
         sSQL = sSQL & "AND K.Orden_No = '" & Codigo & "' "
  End Select
  sSQL = sSQL & "AND K.Codigo_Inv = P.Codigo_Inv " _
       & "AND K.Periodo = '" & Periodo_Contable & "' " _
       & "AND K.Periodo = P.Periodo " _
       & "AND K.Item = P.Item " _
       & "AND C.Codigo = K.Codigo_P " _
       & "ORDER BY P.Producto,K.Fecha,TP,Numero,K.ID "
  Select_Adodc DtaProv, sSQL
  With DtaProv.Recordset
   If .RecordCount > 0 Then
       'MsgBox .RecordCount & vbCrLf & .Fields("Cliente")
       Numero = .Fields("Numero")
   End If
  End With
  sSQL = "SELECT K.Orden_No,K.Codigo_Barra,K.CodBodega,K.Codigo_Inv,P.Producto,K.Fecha,K.TP,K.Numero,Entrada,Salida " _
       & "FROM Trans_Kardex As K,Catalogo_Productos As P " _
       & "WHERE K.Item = '" & NumEmpresa & "' "
  Select Case OpcPrint
    Case "R"
         sSQL = sSQL & "AND K.Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
              & "AND Codigo_P = '" & Codigo & "' "
    Case "BO"
         sSQL = sSQL & "AND K.TP = 'BO' " _
              & "AND K.Numero = " & Val(Codigo) & " "
    Case "G"
         sSQL = sSQL & "AND Orden_No = '" & Codigo & "' "
  End Select
  sSQL = sSQL & "AND K.Codigo_Inv = P.Codigo_Inv " _
       & "AND K.Periodo = '" & Periodo_Contable & "' " _
       & "AND K.Periodo = P.Periodo " _
       & "AND K.Item = P.Item " _
       & "ORDER BY K.Orden_No,K.Codigo_Inv,K.Fecha,K.TP,K.Numero,K.ID "
  'MsgBox sSQL
  Select_Adodc Datas, sSQL
InicioX = 0.5: InicioY = 0: FormaImp = 1: SizeLetra = 12
DataAnchoCampos InicioX, Datas, SizeLetra, TipoTimes, FormaImp
CantCampos = 6
Pagina = 1
SizeLetra = 8
'Iniciamos la impresion
Printer.FontBold = False
With Datas.Recordset
 If .RecordCount > 0 Then
     .MoveFirst
      EncabNotaInv DtaProv, Numero, OrdenNo, xx, 0, ""
      Printer.FontSize = SizeLetra
      Grupo_No = .Fields("Orden_No")
      Printer.FontBold = True
      PrinterTexto 2, PosLinea, "Orden No. "
      Printer.FontBold = False
      PrinterTexto 3.5, PosLinea, Grupo_No
      PosLinea = PosLinea + 0.4
      Do While Not .EOF
         If Grupo_No <> .Fields("Orden_No") Then
            Printer.FontBold = True
            PrinterTexto 2, PosLinea, "Orden No. "
            Printer.FontBold = False
            PrinterTexto 3.5, PosLinea, Grupo_No
            PosLinea = PosLinea + 0.4
            Grupo_No = .Fields("Orden_No")
         End If
         PrinterFields 1, PosLinea, .Fields("CodBodega"), False
         PrinterFields 2, PosLinea, .Fields("Codigo_Barra")
         PrinterFields 5.5, PosLinea, .Fields("Codigo_Inv"), False
         PrinterFields 8, PosLinea, .Fields("Producto"), False
         PrinterFields 16, PosLinea, .Fields("Entrada"), False
         PrinterFields 18, PosLinea, .Fields("Salida"), False
         PosLinea = PosLinea + 0.4
         If PosLinea >= LimiteAlto Then
            Printer.Line (Ancho(0), PosLinea)-(Ancho(CantCampos), PosLinea), QBColor(0)
            Printer.NewPage
            EncabNotaInv DtaProv, Numero, OrdenNo, xx, 0, ""
            Printer.FontSize = SizeLetra
         End If
        .MoveNext
      Loop
  End If
End With
PosLinea = PosLinea + 0.1
Printer.Line (InicioX, PosLinea)-(Ancho(CantCampos), PosLinea), QBColor(0)
PosLinea = PosLinea + 0.05
Printer.Line (InicioX, PosLinea)-(Ancho(CantCampos), PosLinea), QBColor(0)
PosLinea = PosLinea + 1.5
PrinterTexto 3.5, PosLinea, String(20, "_")
PrinterTexto 8, PosLinea, String(20, "_")
PrinterTexto 14, PosLinea, String(20, "_")

PrinterTexto 4, PosLinea + 0.5, "AUTORIZADO"
PrinterTexto 8, PosLinea + 0.5, "ENTREGUE CONFORME"
PrinterTexto 14, PosLinea + 0.5, "  RECIBI CONFORME"
RatonNormal
MensajeEncabData = ""
Printer.EndDoc

End If
Exit Sub
Errorhandler:
    RatonNormal
    ErrorDeImpresion
    Exit Sub
End Sub

Public Sub ImprimirKardex(Datas As Adodc, Datas1 As Adodc)
Dim PosInicio As Single
On Error GoTo Errorhandler
RatonNormal
Mensajes = "Seguro de Imprimir en:" & vbCrLf & Printer.DeviceName & "?"
Titulo = "IMPRESION"
Bandera = False
Orientacion_Pagina = 2
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then

RatonReloj
InicioX = 0.5: InicioY = 0
Escala_Centimetro 2, TipoArialNarrow, 9
Pagina = 1
'Iniciamos la impresion
With Datas.Recordset
     CantCampos = 12
     ReDim Ancho(CantCampos + 1) As Single
     Ancho(0) = 0.2    'Bodega
     Ancho(1) = 1.3    'Fecha
     Ancho(2) = 2.9    'TP
     Ancho(3) = 3.6    'Comp_No = Numero
     Ancho(4) = 5.2    'Detalle
     Ancho(5) = 12.9   'Entrada
     Ancho(6) = 14.9   'Salida
     Ancho(7) = 16.9   'Valor_Unit
     Ancho(8) = 19.2   'Valor_Total
     Ancho(9) = 21.5   'Cantidad
     Ancho(10) = 23.5  'Costo_Prom
     Ancho(11) = 25.7  'Saldo_Total
     Ancho(12) = 27.9
     EncabezadoKardex Datas1
     Printer.FontBold = False
    .MoveFirst
     PosInicio = PosLinea - 0.05
     Printer.FontSize = 9
     
     Mifecha = .Fields("Fecha")
     TipoDoc = .Fields("TP")
     Numero = .Fields("Numero")
     Producto = .Fields("Detalle")
     PrinterTexto Ancho(1) + 0.05, PosLinea, Mifecha
     PrinterTexto Ancho(2) + 0.05, PosLinea, TipoDoc
     PrinterVariables Ancho(3) + 0.05, PosLinea, Numero
     PosLinea = PrinterLineasTexto(Ancho(4) + 0.05, PosLinea, Producto, 6.2)
      Do While Not .EOF
         Printer.FontBold = False
         If (TipoDoc <> .Fields("TP")) Or (Numero - .Fields("Numero") <> 0) Then
            'MsgBox TipoDoc & " - " & .Fields("TP") & vbCrLf & Numero & " - " & .Fields("Comp_No")
            TipoDoc = .Fields("TP")
            Numero = .Fields("Numero")
            Mifecha = .Fields("Fecha")
            Producto = .Fields("Detalle")
            PrinterTexto Ancho(1) + 0.05, PosLinea, Mifecha
            PrinterTexto Ancho(2) + 0.05, PosLinea, TipoDoc
            PrinterVariables Ancho(3) + 0.05, PosLinea, Numero
            PosLinea = PrinterLineasTexto(Ancho(4) + 0.05, PosLinea, Producto, 6.2)
         End If
         PrinterFields Ancho(0) + 0.05, PosLinea, .Fields("Bodega")
         PrinterFields Ancho(5) + 0.05, PosLinea, .Fields("Entrada")
         PrinterFields Ancho(6) + 0.05, PosLinea, .Fields("Salida")
         PrinterFields Ancho(7) + 0.05, PosLinea, .Fields("Valor_Unitario")
         PrinterFields Ancho(8) + 0.05, PosLinea, .Fields("Valor_Total")
         PrinterFields Ancho(9) + 0.05, PosLinea, .Fields("Stock")
         PrinterFields Ancho(10) + 0.05, PosLinea, .Fields("Costo")
         PrinterFields Ancho(11) + 0.05, PosLinea, .Fields("Saldo")
         PosLinea = PosLinea + 0.4
         If PosLinea >= LimiteAlto Then
            For I = 0 To CantCampos
                Printer.Line (Ancho(I), PosInicio)-(Ancho(I), PosLinea), QBColor(0)
            Next I
            Printer.Line (InicioX, PosLinea)-(Ancho(CantCampos), PosLinea), QBColor(0)
            Printer.NewPage
            PosLinea = 0
            EncabezadoKardex Datas1
            Printer.FontSize = 9
            PosInicio = PosLinea - 0.05
         End If
        .MoveNext
      Loop
End With
For I = 0 To CantCampos
    Printer.Line (Ancho(I), PosInicio)-(Ancho(I), PosLinea), QBColor(0)
Next I
Printer.Line (InicioX, PosLinea)-(Ancho(CantCampos), PosLinea), QBColor(0)
PosLinea = PosLinea + 0.05
Printer.Line (InicioX, PosLinea)-(Ancho(CantCampos), PosLinea), QBColor(0)
UltimaLinea = PosLinea + 0.5
RatonNormal
MensajeEncabData = ""
Printer.EndDoc
End If
Exit Sub
Errorhandler:
    RatonNormal
    ErrorDeImpresion
    Exit Sub
End Sub

Public Sub ImprimirKardexCantidad(Datas As Adodc, Datas1 As Adodc)
On Error GoTo Errorhandler
RatonReloj
InicioX = 0.5: InicioY = 0
Escala_Centimetro 1, TipoCondensed, 8
Pagina = 1
'Iniciamos la impresion
With Datas.Recordset
     CantCampos = 10
     ReDim Ancho(CantCampos) As Single
     Ancho(0) = 0.5   'Fecha
     Ancho(1) = 1.7   'Comp_No
     Ancho(2) = 3.5   'Detalle
     Ancho(3) = 11    '
     Ancho(4) = 13    '
     Ancho(5) = 15    '
     Ancho(6) = 17    '
     Ancho(7) = 13    'Entrada
     Ancho(8) = 15    'Salida
     Ancho(9) = 17    'Stock
     Ancho(10) = 19
     EncabezadoKardexCantidad Datas1
     Printer.FontBold = False
    .MoveFirst
     Printer.FontSize = 8
      Do While Not .EOF
         PrinterFields Ancho(0), PosLinea, .Fields("Fecha")
         PrinterFields Ancho(1), PosLinea, .Fields("Comp_No")
         PrinterFields Ancho(2), PosLinea, .Fields("Detalle")
         PrinterFields Ancho(7), PosLinea, .Fields("Entrada")
         PrinterFields Ancho(8), PosLinea, .Fields("Salida")
         PrinterFields Ancho(9), PosLinea, .Fields("Cantidad")
        'PrinterAllFields CantCampos, PosLinea, Datas, True
         PosLinea = PosLinea + 0.4
         If PosLinea > LimiteAlto Then
            Printer.Line (InicioX, PosLinea)-(Ancho(CantCampos), PosLinea), QBColor(0)
            Printer.NewPage
            PosLinea = 0
            EncabezadoKardexCantidad Datas1
            Printer.FontSize = 8
         End If
        .MoveNext
      Loop
End With
Printer.Line (InicioX, PosLinea)-(Ancho(CantCampos), PosLinea), QBColor(0)
PosLinea = PosLinea + 0.05
Printer.Line (InicioX, PosLinea)-(Ancho(CantCampos), PosLinea), QBColor(0)
UltimaLinea = PosLinea + 0.5
RatonNormal
MensajeEncabData = ""
Printer.EndDoc
Exit Sub
Errorhandler:
    RatonNormal
    ErrorDeImpresion
    Exit Sub
End Sub

Public Sub ImprimirAdoCostos(Datas As Adodc, _
                             FinDoc As Boolean, _
                             FormaImp As Byte, _
                             SizeLetra As Integer, _
                             TotalCosto As Currency)
On Error GoTo Errorhandler
RatonReloj
Mensajes = "Seguro de Imprimir en:" & vbCrLf & Printer.DeviceName & "?"
Titulo = "IMPRESION"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
InicioX = 0.5: InicioY = 0.1
DataAnchoCampos InicioX, Datas, SizeLetra, TipoTimes, FormaImp
Pagina = 1
'Iniciamos la impresion
Printer.FontBold = False
With Datas.Recordset
 If .RecordCount > 0 Then
     EncabezadoData Datas
     Printer.FontSize = SizeLetra
     .MoveFirst
     Do While Not .EOF
        PrinterAllFields CantCampos, PosLinea, Datas, True, False
        PosLinea = PosLinea + 0.36
        If PosLinea >= LimiteAlto Then
           Printer.Line (Ancho(0), PosLinea)-(Ancho(CantCampos), PosLinea), QBColor(0)
           Printer.NewPage
           EncabezadoData Datas
           Printer.FontSize = SizeLetra
        End If
       .MoveNext
     Loop
  End If
End With
Printer.Line (InicioX, PosLinea)-(Ancho(CantCampos), PosLinea), QBColor(0)
PosLinea = PosLinea + 0.05
Printer.Line (InicioX, PosLinea)-(Ancho(CantCampos), PosLinea), QBColor(0)
PosLinea = PosLinea + 0.05
PrinterTexto Ancho(3), PosLinea, "TOTAL"
PrinterVariables Ancho(4), PosLinea, TotalCosto
RatonNormal
MensajeEncabData = ""
If FinDoc Then Printer.EndDoc
Exit Sub
Errorhandler:
    RatonNormal
    ErrorDeImpresion
    Exit Sub
Else
    RatonNormal
End If
End Sub

Public Sub Imprimir_Resumen_Kardex(Datas As Adodc, _
                                   SizeLetra As Integer, _
                                   Resumido As Boolean, _
                                   Total_Inv As Currency, _
                                   Optional EsCampoCorto As Boolean)
                         

On Error GoTo Errorhandler
'MsgBox Resumido
RatonReloj
Total = 0
Mensajes = "Seguro de Imprimir en:" & vbCrLf & Printer.DeviceName & "?"
Titulo = "IMPRESION"
Bandera = False
Orientacion_Pagina = 2
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
InicioX = 0.5: InicioY = 0
DataAnchoCampos InicioX, Datas, SizeLetra, TipoArialNarrow, Orientacion_Pagina, EsCampoCorto
Ancho(2) = Ancho(2) - 1
For I = 3 To CantCampos - 1
    Ancho(I) = Ancho(I) + 0.7
Next I
Pagina = 1
Ancho(CantCampos) = LimiteAncho
'Iniciamos la impresion
Printer.FontBold = False
Debitos = 0: Creditos = 0
With Datas.Recordset
 If .RecordCount > 0 Then
    .MoveFirst
     EncabezadoData Datas
     If Cuadricula Then
        If EnDosPaginas > 0 Then
           Imprimir_Linea_H PosLinea, Ancho(0), LimiteAncho, Gris
        Else
           Imprimir_Linea_H PosLinea, Ancho(0), Ancho(CantCampos), Gris
        End If
        PosLinea = PosLinea + 0.05
     End If
     Printer.FontSize = SizeLetra - 1
     Do While Not .EOF
        If .Fields("TC") = "I" Then
            Printer.FontBold = True
        Else
            Printer.FontBold = False
            Total = Total + .Fields("Total")
        End If
        If EnDosPaginas > 0 Then
           PrinterAllFields EnDosPaginas, PosLinea, Datas, True, False
        Else
           PrinterAllFields CantCampos, PosLinea, Datas, True, False
        End If
        PosLinea = PosLinea + 0.44
        If Cuadricula Then
           If EnDosPaginas > 0 Then
              Imprimir_Linea_H PosLinea, Ancho(0), LimiteAncho, Gris
           Else
              Imprimir_Linea_H PosLinea, Ancho(0), Ancho(CantCampos), Gris
           End If
           PosLinea = PosLinea + 0.05
        End If
        If PosLinea >= LimiteAlto Then
           If EnDosPaginas > 0 Then
              Imprimir_Linea_H PosLinea, Ancho(0), LimiteAncho, QBColor(0)
           Else
              Imprimir_Linea_H PosLinea, Ancho(0), Ancho(CantCampos), QBColor(0)
           End If
           Printer.NewPage
           EncabezadoData Datas
           Printer.FontSize = SizeLetra - 1
        End If
        Debitos = .Fields("Entradas")
       .MoveNext
     Loop
 End If
End With
PrinterTexto 1, PosLinea, "TOTAL INVENTARIO"
PrinterVariables 6, PosLinea, Total_Inv
RatonNormal
MensajeEncabData = "": SQLMsg1 = "": SQLMsg2 = "": SQLMsg3 = "": SQLMsg4 = ""
Printer.EndDoc
Exit Sub
Errorhandler:
    RatonNormal
    ErrorDeImpresion
    Exit Sub
Else
    RatonNormal
End If
End Sub

Public Sub ImprimirProv(Datas As Adodc, FinDoc As Boolean, FormaImp As Byte)
Dim Anchos As Single
On Error GoTo Errorhandler
RatonReloj
'Escala_Centimetro FormaImp, TipoCondensed, 9
DataAnchoCampos 0.5, Datas, 9, TipoCondensed, FormaImp
PosLinea = 0: C = CantCampos
'Iniciamos la impresion
Pagina = 1
EncabezadoData Datas
With Datas.Recordset
    .MoveFirst
Do While Not .EOF
  Anchos = InicioX
  Printer.FontBold = False
  For J = 0 To C - 1
    Printer.Line (Ancho(J), PosLinea - 0.1)-(Ancho(J), PosLinea + 0.4), QBColor(0)
    CampoWidth .Fields(J)
    Printer.CurrentX = Ancho(J) + Distancia + 0.1
    Printer.CurrentY = PosLinea
    Printer.Print StrgFormatoCampo
  Next J
  Printer.Line (Ancho(C), PosLinea - 0.1)-(Ancho(C), PosLinea + 0.4), QBColor(0)
  PosLinea = PosLinea + 0.4
  If PosLinea > LimiteAlto Then
    Printer.NewPage
    EncabezadoData Datas
  End If
  .MoveNext
Loop
End With
If PosLinea > LimiteAlto Then PosLinea = 0.5
Printer.Line (0.5, PosLinea)-(Ancho(C), PosLinea), QBColor(0)
PosLinea = PosLinea + 0.1
Printer.Line (0.5, PosLinea)-(Ancho(C), PosLinea), QBColor(0)
UltimaLinea = PosLinea + 1
RatonNormal
MensajeEncabData = ""
If FinDoc Then Printer.EndDoc
Exit Sub
Errorhandler:
    RatonNormal
    ErrorDeImpresion
    Exit Sub
End Sub

Public Sub ImprimirCatalogoInv(Datas As Adodc)
Dim SizeLetra As Integer
On Error GoTo Errorhandler
RatonReloj
Mensajes = "Seguro de Imprimir en:" & vbCrLf & Printer.DeviceName & "?"
Titulo = "IMPRESION"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
InicioX = 0.5: InicioY = 0
SizeLetra = 9
'Escala_Centimetro 1, TipoTimes, SizeLetra
DataAnchoCampos InicioX, Datas, SizeLetra, TipoTimes, 1
'''Ancho(0) = 0.5  'TP
'''Ancho(1) = 1.2  'Codigo_Inv
'''Ancho(2) = 3    'Producto
'''Ancho(3) = 10.5 'PVP /Unidad
''''Ancho(4) = 9.5  'Codigo_Barra
'''Ancho(5) = 12   'Cta_Inventario
'''Ancho(6) = 14.5 'Cta_Costo_Venta
'''Ancho(7) = 17   'Cta_Ventas
'''Ancho(8) = 19.5
Pagina = 1
CantCampos = 8
'Iniciamos la impresion
EncabezadoData Datas
Printer.FontBold = False
With Datas.Recordset
     .MoveFirst
      Do While Not .EOF
'''         Select Case Nivel(.Fields("Codigo_Inv"))
'''           Case 1: PCol = 0
'''           Case 2: PCol = 0.3
'''           Case 3: PCol = 0.6
'''           Case 4: PCol = 0.9
'''           Case 5: PCol = 1.2
'''           Case 6: PCol = 1.5
'''         End Select
'''         Printer.FontBold = False: Printer.FontSize = SizeLetra
'''         If .Fields("TC") <> "P" Then Printer.FontBold = True
'''         PrinterFields Ancho(0), PosLinea, .Fields("TC"), True
'''         PrinterFields Ancho(1), PosLinea, .Fields("Codigo_Inv"), True
'''         PrinterFields Ancho(2) + PCol, PosLinea, .Fields("Producto"), False
'''         If .Fields("TC") <> "I" Then
'''             'PrinterFields Ancho(3), PosLinea, .Fields("Unidad"), False
'''             PrinterFields Ancho(3), PosLinea, .Fields("PVP"), False
'''             'PrinterFields Ancho(4), PosLinea, .Fields("Codigo_Barra"), False
'''             PrinterFields Ancho(5), PosLinea, .Fields("Cta_Inventario"), False
'''             PrinterFields Ancho(6), PosLinea, .Fields("Cta_Costo_Venta"), False
'''             PrinterFields Ancho(7), PosLinea, .Fields("Cta_Ventas"), False
'''         End If
'''         For I = 3 To CantCampos
'''             Printer.Line (Ancho(I), PosLinea - 0.1)-(Ancho(I), PosLinea + 0.4), QBColor(0)
'''         Next I
         PrinterAllFields CantCampos, PosLinea, Datas, True, False
         PosLinea = PosLinea + 0.35
         If PosLinea > LimiteAlto Then
            Printer.Line (Ancho(0), PosLinea)-(Ancho(CantCampos), PosLinea), QBColor(0)
            Printer.NewPage
            PosLinea = 0
            EncabezadoData Datas
         End If
        .MoveNext
      Loop
End With
Printer.Line (Ancho(0), PosLinea)-(Ancho(CantCampos), PosLinea), QBColor(0)
PosLinea = PosLinea + 0.05
Printer.Line (Ancho(0), PosLinea)-(Ancho(CantCampos), PosLinea), QBColor(0)
RatonNormal
MensajeEncabData = ""
Printer.EndDoc
End If
Exit Sub
Errorhandler:
    RatonNormal
    ErrorDeImpresion
    Exit Sub
End Sub

Public Sub ImprimirCatalogoActivos(Datas As Adodc)
Dim SizeLetra As Integer
On Error GoTo Errorhandler
RatonReloj
Mensajes = "Seguro de Imprimir en:" & vbCrLf & Printer.DeviceName & "?"
Titulo = "IMPRESION"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
InicioX = 0.5: InicioY = 0
SizeLetra = 8
'Escala_Centimetro 1, TipoTimes, SizeLetra
DataAnchoCampos InicioX, Datas, SizeLetra, TipoArialNarrow, 1
'''Ancho(0) = 0.5  'TP
'''Ancho(1) = 1.2  'Codigo_Inv
'''Ancho(2) = 3    'Producto
'''Ancho(3) = 10.5 'PVP /Unidad
''''Ancho(4) = 9.5  'Codigo_Barra
'''Ancho(5) = 12   'Cta_Inventario
'''Ancho(6) = 14.5 'Cta_Costo_Venta
'''Ancho(7) = 17   'Cta_Ventas
'''Ancho(8) = 19.5
Pagina = 1
'CantCampos = 8
'Iniciamos la impresion
EncabezadoData Datas
Printer.FontBold = False
With Datas.Recordset
 If .RecordCount > 0 Then
    .MoveFirst
     Codigo = .Fields("Ubicacion")
     Codigo1 = .Fields("Nombre_Responsable")
     Printer.FontBold = True
     PrinterTexto Ancho(0), PosLinea, "RESPONSABLE:"
     Printer.FontBold = False
     PrinterTexto Ancho(0) + 2, PosLinea, UCaseStrg(Codigo1)
     PosLinea = PosLinea + 0.4
     Do While Not .EOF
'''         Select Case Nivel(.Fields("Codigo_Inv"))
'''           Case 1: PCol = 0
'''           Case 2: PCol = 0.3
'''           Case 3: PCol = 0.6
'''           Case 4: PCol = 0.9
'''           Case 5: PCol = 1.2
'''           Case 6: PCol = 1.5
'''         End Select
'''         Printer.FontBold = False: Printer.FontSize = SizeLetra
'''         If .Fields("TC") <> "P" Then Printer.FontBold = True
'''         PrinterFields Ancho(0), PosLinea, .Fields("TC"), True
'''         PrinterFields Ancho(1), PosLinea, .Fields("Codigo_Inv"), True
'''         PrinterFields Ancho(2) + PCol, PosLinea, .Fields("Producto"), False
'''         If .Fields("TC") <> "I" Then
'''             'PrinterFields Ancho(3), PosLinea, .Fields("Unidad"), False
'''             PrinterFields Ancho(3), PosLinea, .Fields("PVP"), False
'''             'PrinterFields Ancho(4), PosLinea, .Fields("Codigo_Barra"), False
'''             PrinterFields Ancho(5), PosLinea, .Fields("Cta_Inventario"), False
'''             PrinterFields Ancho(6), PosLinea, .Fields("Cta_Costo_Venta"), False
'''             PrinterFields Ancho(7), PosLinea, .Fields("Cta_Ventas"), False
'''         End If
'''         For I = 3 To CantCampos
'''             Printer.Line (Ancho(I), PosLinea - 0.1)-(Ancho(I), PosLinea + 0.4), QBColor(0)
'''         Next I
         SizeLetra = 8
         If Codigo <> .Fields("Ubicacion") Then
            PosLinea = PosLinea + 0.35
            Codigo = .Fields("Ubicacion")
         End If
         PrinterAllFields CantCampos, PosLinea, Datas, False, False
         PosLinea = PosLinea + 0.35
         If PosLinea >= LimiteAlto Then
            Printer.Line (Ancho(0), PosLinea)-(Ancho(CantCampos), PosLinea), QBColor(0)
            Printer.NewPage
            PosLinea = 0
            EncabezadoData Datas
            Codigo1 = .Fields("Nombre_Responsable")
            Printer.FontBold = True
            PrinterTexto Ancho(0), PosLinea, "RESPONSABLE:"
            Printer.FontBold = False
            PrinterTexto Ancho(0) + 3, PosLinea, UCaseStrg(Codigo1)
            PosLinea = PosLinea + 0.4
         End If
        .MoveNext
      Loop
 End If
End With
Printer.Line (Ancho(0), PosLinea)-(Ancho(CantCampos), PosLinea), QBColor(0)
PosLinea = PosLinea + 0.05
Printer.Line (Ancho(0), PosLinea)-(Ancho(CantCampos), PosLinea), QBColor(0)
PosLinea = PosLinea + 1.5
Printer.FontName = TipoArial
Printer.FontBold = True
PrinterTexto Ancho(0) + 3, PosLinea, "_________________"
PrinterTexto Ancho(3), PosLinea, "__________________"
PosLinea = PosLinea + 0.4
PrinterTexto Ancho(0) + 3.5, PosLinea, "Recibido por"
PrinterTexto Ancho(3) + 0.5, PosLinea, "Entregado por"
Printer.FontBold = False
RatonNormal
MensajeEncabData = ""
Printer.EndDoc
End If
Exit Sub
Errorhandler:
    RatonNormal
    ErrorDeImpresion
    Exit Sub
End Sub

Public Sub Imprimir_Resumen_Barra(Datas As Adodc, _
                                  SizeLetra As Integer, _
                                  Optional EsCampoCorto As Boolean)
On Error GoTo Errorhandler
RatonReloj
Mensajes = "Seguro de Imprimir en:" & vbCrLf & Printer.DeviceName & "?"
Titulo = "IMPRESION"
Bandera = False
Orientacion_Pagina = 2
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
InicioX = 0.5: InicioY = 0
DataAnchoCampos InicioX, Datas, SizeLetra, TipoTimes, Orientacion_Pagina, EsCampoCorto
Ancho(1) = Ancho(1) - 1
Ancho(2) = Ancho(2) - 1
Ancho(3) = Ancho(3) - 1
Ancho(4) = Ancho(4) - 1
Ancho(5) = Ancho(5) - 1
Ancho(6) = Ancho(6) - 1
Ancho(7) = Ancho(7) - 1
Ancho(CantCampos) = 19.15
Pagina = 1
'Iniciamos la impresion
Printer.FontBold = False
With Datas.Recordset
    .MoveFirst
     EncabezadoData Datas
     Printer.FontSize = SizeLetra
     Do While Not .EOF
        If MidStrg(.Fields("Detalle"), 1, 3) = "---" Then
           PosLinea = PosLinea + 0.07
           Printer.Line (InicioX, PosLinea)-(Ancho(CantCampos), PosLinea), QBColor(0)
           PosLinea = PosLinea + 0.05
           Printer.Line (InicioX, PosLinea)-(Ancho(CantCampos), PosLinea), QBColor(0)
           PosLinea = PosLinea + 0.07
        Else
           PrinterAllFields CantCampos, PosLinea, Datas, True, False
           PosLinea = PosLinea + 0.36
        End If
        
        If PosLinea >= LimiteAlto Then
           Printer.Line (Ancho(0), PosLinea)-(Ancho(CantCampos), PosLinea), QBColor(0)
           Printer.NewPage
           EncabezadoData Datas
           Printer.FontSize = SizeLetra
        End If
       .MoveNext
     Loop
End With
RatonNormal
MensajeEncabData = ""
Printer.EndDoc
Exit Sub
Errorhandler:
    RatonNormal
    ErrorDeImpresion
    Exit Sub
Else
    RatonNormal
End If
End Sub

Public Sub ImprimirCostoInventario(Datas As Adodc, _
                                   FormaImp As Byte, _
                                   SizeLetra As Integer, _
                                   Optional EsCampoCorto As Boolean, _
                                   Optional PiePagina As String)
On Error GoTo Errorhandler
RatonReloj
Mensajes = "Seguro de Imprimir en:" & vbCrLf & Printer.DeviceName & "?"
Titulo = "IMPRESION"
Bandera = False
Orientacion_Pagina = 2
SetPrinters.Show 1

If PonImpresoraDefecto(SetNombrePRN) Then
InicioX = 0.5: InicioY = 0
DataAnchoCampos InicioX, Datas, SizeLetra, TipoArialNarrow, Orientacion_Pagina, EsCampoCorto
'''Cadena = "Ancho:" & vbCrLf
'''For I = 0 To CantCampos
'''    Cadena = Cadena & Ancho(I) & vbCrLf
'''Next I
'''MsgBox Cadena & AnchoPapel & ", Paginas = " & EnDosPaginas
Pagina = 1
'Iniciamos la impresion
Printer.FontBold = False
With Datas.Recordset
 If .RecordCount > 0 Then
    .MoveFirst
     EncabezadoData Datas
     Printer.FontName = TipoArialNarrow
     If Cuadricula Then
        If EnDosPaginas > 0 Then
           Imprimir_Linea_H PosLinea, Ancho(0), AnchoPapel, Gris
        Else
           Imprimir_Linea_H PosLinea, Ancho(0), Ancho(CantCampos), Gris
        End If
        PosLinea = PosLinea + 0.05
     End If
     Printer.FontSize = SizeLetra
     Total = 0
     Saldo = 0
     Do While Not .EOF
        Total = Total + .Fields("Total_Costo")
        Saldo = Saldo + .Fields("Total_PVP")
        If EnDosPaginas > 0 Then
           PrinterAllFields EnDosPaginas, PosLinea, Datas, True, False
        Else
           PrinterAllFields CantCampos, PosLinea, Datas, True, False
        End If
        PosLinea = PosLinea + 0.36
        If Cuadricula Then
           If EnDosPaginas > 0 Then
              Imprimir_Linea_H PosLinea, Ancho(0), AnchoPapel, Gris
           Else
              Imprimir_Linea_H PosLinea, Ancho(0), Ancho(CantCampos), Gris
           End If
           PosLinea = PosLinea + 0.05
        End If
        If PosLinea >= LimiteAlto Then
           If Cuadricula Then
              If EnDosPaginas > 0 Then
                 Imprimir_Linea_H PosLinea, Ancho(0), AnchoPapel, Gris
              Else
                 Imprimir_Linea_H PosLinea, Ancho(0), Ancho(CantCampos), Gris
              End If
              PosLinea = PosLinea + 0.05
           End If
           Printer.NewPage
           EncabezadoData Datas
           Printer.FontSize = SizeLetra
           Printer.FontName = TipoArialNarrow
        End If
       .MoveNext
     Loop
     If EnDosPaginas > 0 Then
        Imprimir_Linea_H PosLinea, Ancho(0), AnchoPapel, Negro, True
     Else
        Imprimir_Linea_H PosLinea, Ancho(0), Ancho(CantCampos), Negro, True
     End If
    'Imprimimos en la segunda pagina el restante de la impresion
     If EnDosPaginas > 0 Then
        Printer.NewPage
        EncabezadoData Datas, True
        If Cuadricula Then
           Imprimir_Linea_H PosLinea, Ancho(0), Ancho(CantCampos), Gris
           PosLinea = PosLinea + 0.05
        End If
        Printer.FontSize = SizeLetra
        Printer.FontBold = False
       .MoveFirst
        Do While Not .EOF
           Printer.FontName = TipoArialNarrow
           PrinterAllFields CantCampos, PosLinea, Datas, True, False, True
           PosLinea = PosLinea + 0.36
           If Cuadricula Then
              Imprimir_Linea_H PosLinea, Ancho(0), Ancho(CantCampos), Gris
              PosLinea = PosLinea + 0.05
           End If
           If PosLinea >= LimiteAlto Then
              Imprimir_Linea_H PosLinea, Ancho(0), Ancho(CantCampos)
              Printer.NewPage
              EncabezadoData Datas, True
              Printer.FontSize = SizeLetra
           End If
          .MoveNext
        Loop
       .MoveFirst
        Imprimir_Linea_H PosLinea, Ancho(0), Ancho(CantCampos), Negro, True
     End If
 End If
End With
PosLinea = PosLinea + 0.05
PrinterTexto Ancho(5), PosLinea, "T O T A L E S"
PrinterVariables Ancho(7), PosLinea, Total
PrinterVariables Ancho(8), PosLinea, Saldo
PrinterVariables Ancho(9), PosLinea, Saldo - Total

PosLinea = PosLinea + 0.05
If PiePagina <> "" Then PrinterTexto Ancho(0), PosLinea, PiePagina
Cuadricula = False
MensajeEncabData = "": SQLMsg1 = "": SQLMsg2 = "": SQLMsg3 = "": SQLMsg4 = ""
RatonNormal
Printer.EndDoc
Exit Sub
Errorhandler:
    RatonNormal
    ErrorDeImpresion
    Exit Sub
Else
    RatonNormal
End If
End Sub

Public Function Generar_Salidas_Excel(Co As Comprobantes) As String
Dim NFila As Integer
Dim NCelda As Integer
Dim RutaGeneraFile As String
Dim Archivo As String
Dim Doctor As String
Dim Tratamiento As String
Dim oExcel As Object
Dim oBook As Object
Dim oSheet As Object
Dim DBKardex As ADODB.Recordset
Dim DBCliente As ADODB.Recordset

  RatonReloj
  Set DBKardex = New ADODB.Recordset
  Set DBCliente = New ADODB.Recordset
  DBKardex.CursorType = adOpenStatic
  DBCliente.CursorType = adOpenStatic
  DBKardex.CursorLocation = adUseClient
  DBCliente.CursorLocation = adUseClient
  
 'Start a new workbook in Excel
  Set oExcel = CreateObject("Excel.Application")
  Set oBook = oExcel.Workbooks.Add
  
 'Add data to cells of the first worksheet in the new workbook
  Set oSheet = oBook.Worksheets(1)
  Archivo = "Insumo " & Co.TP & " " & Format(Co.Numero, "00000000")
   
  sSQL = "SELECT P.Producto, P.IVA, K.Fecha, K.TP, K.Numero, C.Cliente, Co.Concepto, K.Salida, K.Valor_Unitario, " _
       & "K.Valor_Total, K.Orden_No, K.Codigo_Dr, K.Codigo_Tra, K.No_Refrendo, K.Descuento " _
       & "FROM Trans_Kardex As K, Catalogo_Productos As P, Comprobantes As Co, Clientes As C " _
       & "WHERE K.Item = '" & NumEmpresa & "' " _
       & "AND K.Periodo = '" & Periodo_Contable & "' " _
       & "AND K.TP = '" & Co.TP & "' " _
       & "AND K.Numero = " & Co.Numero & " " _
       & "AND K.Codigo_Inv = P.Codigo_Inv " _
       & "AND K.TP = Co.TP " _
       & "AND K.Numero = Co.Numero " _
       & "AND K.Item = Co.Item " _
       & "AND K.Item = P.Item " _
       & "AND Co.Codigo_B = C.Codigo " _
       & "AND K.Periodo = Co.Periodo " _
       & "AND K.Periodo = P.Periodo " _
       & "ORDER BY K.Orden_No, P.Producto, K.Fecha, K.TP, K.Numero, K.ID "
  Select_AdoDB DBKardex, sSQL
  RatonReloj
  Contador = 0
  RutaGeneraFile = RutaSysBases & "\TEMP\" & TrimStrg(Archivo) & ".xls"
  If Dir(RutaGeneraFile) <> "" Then Kill RutaGeneraFile
  With DBKardex
   If .RecordCount > 0 Then
       Validar_Porc_IVA (.Fields("Fecha"))
       Doctor = Ninguno
       sSQL = "SELECT Cliente " _
            & "FROM Clientes " _
            & "WHERE Codigo = '" & .Fields("Codigo_Dr") & "' "
       Select_AdoDB DBCliente, sSQL
       If DBCliente.RecordCount > 0 Then Doctor = DBCliente.Fields("Cliente")
       DBCliente.Close
       
       Tratamiento = Ninguno
       sSQL = "SELECT Producto " _
            & "FROM Catalogo_Productos " _
            & "WHERE Codigo_Inv = '" & .Fields("Codigo_Tra") & "' "
       Select_AdoDB DBCliente, sSQL
       If DBCliente.RecordCount > 0 Then Tratamiento = DBCliente.Fields("Producto")
       DBCliente.Close
       
      'Ancho de cada columna
       oSheet.Columns("A").columnWidth = 40
      'oSheet.Cells(Fila,Columna)
       For NCelda = 1 To 11
           oSheet.Columns(Chr(65 + NCelda)).columnWidth = 12
           oSheet.cells(13, NCelda).HorizontalAlignment = vbCenter
           oSheet.cells(13, NCelda).VerticalAlignment = vbCenter
           oSheet.cells(13, NCelda).WrapText = True
           oSheet.cells(13, NCelda).Interior.ColorIndex = 41    ' Color fondo = azul '41
           oSheet.cells(13, NCelda).Font.Size = 9               ' tamaño de letra
           oSheet.cells(13, NCelda).Font.Bold = True            ' Fuente en negrita
           oSheet.cells(13, NCelda).Font.ColorIndex = 2         ' Color fuente = blanco
       Next NCelda
       oSheet.rows(13).RowHeight = 60
      'Encabezado
       For NFila = 5 To 11
           oSheet.cells(NFila, 1).Interior.ColorIndex = 41   ' Color fondo = azul '41
           oSheet.cells(NFila, 1).Font.Size = 10             ' tamaño de letra
           oSheet.cells(NFila, 1).Font.Bold = True           ' Fuente en negrita
           oSheet.cells(NFila, 1).Font.ColorIndex = 2        ' Color fuente = blanco
       Next NFila
       
       For NCelda = 1 To 11
           oSheet.cells(12, NCelda).Interior.ColorIndex = Magenta    ' Color fondo = azul '41
           oSheet.cells(12, NCelda).Font.Size = 10              ' tamaño de letra
           oSheet.cells(12, NCelda).Font.Bold = True            ' Fuente en negrita
           oSheet.cells(12, NCelda).Font.ColorIndex = 2         ' Color fuente = blanco
       Next NCelda
       
       oSheet.Range("A5").value = "CLIENTE:"
       oSheet.Range("B5").value = "IESS POR PAGAR"
       oSheet.Range("A6").value = "FECHA:"
       oSheet.Range("B6").value = "'" & Format(.Fields("Fecha"), "dd/MM/yyyy")
       oSheet.Range("A7").value = "PACIENTE:"
       oSheet.Range("B7").value = .Fields("Cliente")
       oSheet.Range("A8").value = "TRATAMIENTO:"
       oSheet.Range("B8").value = Tratamiento
       oSheet.Range("A9").value = "OJO A OPERAR:"
       oSheet.Range("B9").value = ""
       oSheet.Range("A10").value = "HISTORIA CLINICA:"
       oSheet.Range("B10").value = ""
       oSheet.Range("A11").value = "MEDICO:"
       oSheet.Range("B11").value = Doctor
       oSheet.Range("B12").value = "HOJA DE INSUMOS UTILIZADOS EN LA SESION QUIRURGICA"
       
      'Detalle de isumos
       oSheet.Range("A13").value = "INSUMO"
       oSheet.Range("B13").value = "CANT."
       oSheet.Range("C13").value = "PRECIO UNITORIO BASE IVA 0%"
       oSheet.Range("D13").value = "PRECIO UNITARIO BASE IVA " & Format(Porc_IVA, "00%")
       oSheet.Range("E13").value = "SUBTOTAL"
       oSheet.Range("F13").value = "10% DE GESTION AUTORIZADO IESS DEL SUBTOTAL"
       oSheet.Range("G13").value = "SUBTOTAL"
       oSheet.Range("H13").value = "IVA " & Format(Porc_IVA, "00%")
       oSheet.Range("I13").value = "TOTAL"
       oSheet.Range("J13").value = "PROVEEDOR"
       oSheet.Range("K13").value = "FACTURA No"
       NFila = 14
       Saldo = 0
       Do While Not .EOF
          oSheet.Range("A" & CStr(NFila)).value = .Fields("Producto")
          oSheet.Range("B" & CStr(NFila)).value = .Fields("Salida")
          Total_IVA = 0
          PVP = .Fields("Valor_Unitario")
          Total = .Fields("Valor_Total")
          If Not .Fields("IVA") Then
             oSheet.Range("C" & CStr(NFila)).value = Total
          Else
             oSheet.Range("D" & CStr(NFila)).value = Total
             Total_IVA = Redondear(Total * Porc_IVA, 2)
          End If
          SubTotal = (Total * (.Fields("Descuento") / 100)) + Total
          oSheet.Range("E" & CStr(NFila)).value = Total
          oSheet.Range("F" & CStr(NFila)).value = Format(.Fields("Descuento") / 100, "00%")
          oSheet.Range("G" & CStr(NFila)).value = SubTotal
          oSheet.Range("H" & CStr(NFila)).value = Total_IVA
          oSheet.Range("I" & CStr(NFila)).value = SubTotal + Total_IVA
          Total = Total + SubTotal + Total_IVA
          oSheet.Range("J" & CStr(NFila)).value = .Fields("No_Refrendo")        'PROVEEDOR
          oSheet.Range("K" & CStr(NFila)).value = .Fields("Orden_No")           'FACTURA No"
          Saldo = Saldo + SubTotal + Total_IVA
          NFila = NFila + 1
         .MoveNext
         'oSheet.Cells(NFila, 14).formula = "=SUM(B" & CStr(NFila) & ":" & "F" & CStr(NFila) & ")/5"
       Loop
       
      'Totales del comprobantes
       For NCelda = 1 To 11
           oSheet.cells(NFila, NCelda).Interior.ColorIndex = Magenta    ' Color fondo = azul '41
           oSheet.cells(NFila, NCelda).Font.Size = 10              ' tamaño de letra
           oSheet.cells(NFila, NCelda).Font.Bold = True            ' Fuente en negrita
           oSheet.cells(NFila, NCelda).Font.ColorIndex = 2         ' Color fuente = blanco
       Next NCelda
 
       oSheet.Range("E" & CStr(NFila)).value = "TOTAL A COBRAR AL IESS"
       oSheet.Range("I" & CStr(NFila)).value = Saldo
       NFila = NFila + 6
       oSheet.Range("E" & CStr(NFila)).value = "________________________"
       NFila = NFila + 1
       oSheet.Range("E" & CStr(NFila)).value = "    FIRMA DEL CIRUJANO"
      'Bloqueamos las celdas que no se puden cambiar
       
       oSheet.Unprotect "DiskCoverSystem"
       oSheet.Range("B2:B" & CStr(NFila)).Locked = False
       oSheet.Range("C2:C" & CStr(NFila)).Locked = False
       oSheet.Range("D2:D" & CStr(NFila)).Locked = False
       oSheet.Range("E2:E" & CStr(NFila)).Locked = False
       oSheet.Range("F2:F" & CStr(NFila)).Locked = False
       oSheet.Range("G2:F" & CStr(NFila)).Locked = False
       oSheet.Range("H2:F" & CStr(NFila)).Locked = False
       oSheet.Range("I2:F" & CStr(NFila)).Locked = False
       oSheet.Protect "DiskCoverSystem"
      'Save the Workbook and Quit Excel
        
       oBook.SaveAs RutaGeneraFile
       oExcel.Quit
   Else
      RutaGeneraFile = ""
   End If
  End With
  DBKardex.Close
  RatonNormal
  Generar_Salidas_Excel = RutaGeneraFile
End Function


