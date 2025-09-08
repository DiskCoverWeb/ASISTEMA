Attribute VB_Name = "SubContabil"
Option Explicit

Public Sub Actualiza_Fecha_Tabla(Tabla As String, Fecha As String)
    If IsDate(Fecha) Then
       sSQL = "UPDATE " & Tabla & " " _
            & "SET Fecha = #" & BuscarFecha(Fecha) & "# " _
            & "WHERE Item = '" & NumEmpresa & "' " _
            & "AND Periodo = '" & Periodo_Contable & "' " _
            & "AND TP = '" & Co.TP & "' " _
            & "AND Numero = '" & Co.Numero & "' "
       Ejecutar_SQL_SP sSQL
    End If
End Sub

Public Sub Actualiza_Procesado_Tabla(Tabla As String, Optional ConTP As Boolean, Optional Cuenta As String, Optional Valor As String)
    sSQL = "UPDATE " & Tabla & " " _
         & "SET Procesado = 0 " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' "
    If ConTP Then
       sSQL = sSQL _
            & "AND TP = '" & Co.TP & "' " _
            & "AND Numero = '" & Co.Numero & "' "
    End If
    If Len(Cuenta) > 1 And Len(Valor) > 1 Then sSQL = sSQL & "AND " & Cuenta & " = '" & Valor & "' "
    Ejecutar_SQL_SP sSQL
End Sub

Public Sub Actualiza_Cuenta_Tabla(Tabla As String, Campo As String, CtaOld As String, CtaNew As String, Optional ConTP As Boolean, Optional IDTemp As Long)
    sSQL = "UPDATE " & Tabla & " " _
         & "SET " & Campo & " = '" & CtaNew & "' " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND " & Campo & " = '" & CtaOld & "' "
    If IDTemp > 0 Then sSQL = sSQL & "AND ID = " & IDTemp & " "
    If ConTP Then
       sSQL = sSQL _
            & "AND TP = '" & Co.TP & "' " _
            & "AND Numero = '" & Co.Numero & "' "
    End If
    Ejecutar_SQL_SP sSQL
End Sub

Public Sub Elimina_Cuenta_Tabla(Tabla As String, Cuenta As String, Valor As String, Optional IDTemp As Long)
    sSQL = "DELETE * " _
         & "FROM Transacciones " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND TP = '" & Co.TP & "' " _
         & "AND Numero = '" & Co.Numero & "' " _
         & "AND " & Cuenta & " = '" & Valor & "' "
    If IDTemp > 0 Then sSQL = sSQL & "AND ID = " & IDTemp & " "
    Ejecutar_SQL_SP sSQL
End Sub

Public Sub Imprimir_CxCxP_Meses(AdoBanco As Adodc, Cta As String, FechaCorte As String)
Dim AdoBenefDB As ADODB.Recordset
Dim TC As String
Dim Codigo As String
Dim Cuenta As String
Dim Beneficiario As String
Dim NombFilePict As String
Dim PosCampo(6) As Double
Dim TotalSubModulo As Currency

    RatonReloj
    TotalSubModulo = 0
    Codigo = Ninguno
    Cuenta = Ninguno
    TC = Ninguno
    sSQL = "SELECT TC, Codigo, Cuenta " _
         & "FROM Catalogo_Cuentas " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND Codigo = '" & Cta & "' "
    Select_AdoDB AdoBenefDB, sSQL
    If AdoBenefDB.RecordCount > 0 Then
       TC = AdoBenefDB.fields("TC")
       Codigo = AdoBenefDB.fields("Codigo")
       Cuenta = UCaseStrg(AdoBenefDB.fields("Cuenta"))
    End If
    AdoBenefDB.Close

   With AdoBanco.Recordset
    If .RecordCount > 0 Then
       .MoveFirst
        RatonReloj
        RutaDocumentoPDF = ""
        Select Case TC
          Case "C", "P": MensajeEncabData = "REPORTE DE " & Cuenta
          Case Else: MensajeEncabData = "NINGUN REPORTE"
        End Select
        SQLMsg1 = ""
        SQLMsg2 = ""
        SQLMsg3 = ""
         NombFilePict = "RCXCXP_" & NumEmpresa & "_" & BuscarFecha(FechaSistema) & "-" & CodigoUsuario
         SetNombrePRN = Impresota_PDF
        'Geneeramos el documento
         tPrint.TipoImpresion = Es_PDF
         tPrint.NombreArchivo = NombFilePict
         tPrint.TituloArchivo = MensajeEncabData
         tPrint.TipoLetra = TipoHelvetica
         tPrint.OrientacionPagina = 1
         tPrint.PaginaA4 = True
         tPrint.EsCampoCorto = True
         tPrint.VerDocumento = True
         Set cPrint = New cImpresion
         cPrint.iniciaImpresion
         PosLinea = 2
         PosCampo(1) = 1.5    ' Beneficiario
         PosCampo(2) = 11     ' Anio
         PosCampo(3) = 12     ' Mes
         PosCampo(4) = 13     ' Valor_x_Mes
         PosCampo(5) = 15     ' Categoria
         cPrint.printEncabezado 1.5, PosLinea, TipoHelvetica
        'Pagina No. 1
         cPrint.fondoDeLetra = Negro
         cPrint.tipoNegrilla = True
         cPrint.PorteDeLetra = 8
         cPrint.colorDeLetra = Negro
         Cadena = "La informacion presente reposa en la base de dato de la Institucion, corte realizado al " & FechaStrg(FechaCorte) & ", " _
                & "cualquier informacion adicional comuniquese al personal encargado."
         PosLinea = cPrint.printTextoMultiple(1.5, PosLinea, Cadena, 18.9)
         PosLinea = PosLinea + 0.5
         cPrint.printTexto PosCampo(1), PosLinea, "Beneficiario"
         cPrint.printTexto PosCampo(2), PosLinea, "Anio"
         cPrint.printTexto PosCampo(3), PosLinea, "Mes"
         cPrint.printTexto PosCampo(4), PosLinea, "Valor_x_Mes"
         cPrint.printTexto PosCampo(5), PosLinea, "Categoria"
         PosLinea = PosLinea + 0.4
         cPrint.printLinea 1.2, PosLinea - 0.55, cPrint.dAnchoPapel - 1, PosLinea - 0.55
         cPrint.printLinea 1.2, PosLinea - 0.05, cPrint.dAnchoPapel - 1, PosLinea - 0.05
         PosLinea = PosLinea + 0.1
         cPrint.colorDeLetra = Negro
         cPrint.tipoNegrilla = False
         Beneficiario = .fields("Beneficiario")
         cPrint.printTexto PosCampo(1), PosLinea, Beneficiario
         Do While Not .EOF
            If InStr(.fields("Anio"), "TOTAL") Then
               PosLinea = PosLinea + 0.1
               cPrint.printLinea 1.2, PosLinea - 0.1, cPrint.dAnchoPapel - 1, PosLinea - 0.1
               cPrint.printLinea 1.2, PosLinea + 0.3, cPrint.dAnchoPapel - 1, PosLinea + 0.3
               TotalSubModulo = TotalSubModulo + .fields("Valor_x_Mes")
            End If
            If Beneficiario <> .fields("Beneficiario") Then
               Beneficiario = .fields("Beneficiario")
               cPrint.printTexto PosCampo(1), PosLinea, Beneficiario
            End If
            cPrint.printFields PosCampo(2), PosLinea, .fields("Anio")
            cPrint.printFields PosCampo(3), PosLinea, .fields("Mes")
            cPrint.printFields PosCampo(4), PosLinea, .fields("Valor_x_Mes")
            cPrint.printFields PosCampo(5), PosLinea, .fields("Categoria")
            PosLinea = PosLinea + 0.4
            If InStr(.fields("Anio"), "TOTAL") Then PosLinea = PosLinea + 0.1
            'Siguiente Pagina
            If PosLinea > (cPrint.dAltoPapel - 1.5) Then
               cPrint.paginaNueva
               PosLinea = 1
               cPrint.printEncabezado 1.2, PosLinea, TipoHelvetica
               PosLinea = PosLinea - 0.15
               cPrint.colorDeLetra = Negro
               cPrint.tipoNegrilla = True
               cPrint.PorteDeLetra = 8
               cPrint.printTexto PosCampo(1), PosLinea, "Beneficiario"
               cPrint.printTexto PosCampo(2), PosLinea, "Anio"
               cPrint.printTexto PosCampo(3), PosLinea, "Mes"
               cPrint.printTexto PosCampo(4), PosLinea, "Valor_x_Mes"
               cPrint.printTexto PosCampo(5), PosLinea, "Categoria"
               PosLinea = PosLinea + 0.4
               cPrint.printLinea 1.2, PosLinea - 0.05, cPrint.dAnchoPapel - 1, PosLinea - 0.05
               PosLinea = PosLinea + 0.1
               cPrint.tipoNegrilla = False
            End If
           .MoveNext
         Loop
         'PosLinea = PosLinea + 0.05
         cPrint.printLinea 1.2, PosLinea - 0.1, cPrint.dAnchoPapel - 1, PosLinea - 0.1
         cPrint.printLinea 1.2, PosLinea + 0.3, cPrint.dAnchoPapel - 1, PosLinea + 0.3
         cPrint.printTexto PosCampo(1), PosLinea, "TOTAL " & Cuenta
         cPrint.printVariable PosCampo(4), PosLinea, TotalSubModulo
         RatonNormal
        'fin del documento
         cPrint.finalizaImpresion
    Else
       RatonNormal
       MsgBox "No existe informacion que presentar"
    End If
   End With
End Sub

Public Sub Mayorizar_SubModulos(Cta_SubModulo As String, TNumEmpresa As String)
Dim AdoSubCtasDB As ADODB.Recordset
Dim Cod_Cta As String

  Progreso_Barra.Mensaje_Box = Mifecha & ", Cta: " & Cta_SubModulo & " - Mayorizando la SubCuenta..."
  Progreso_Esperar
  SubCta = Cta_SubModulo
  sSQL = "SELECT * " _
       & "FROM Trans_SubCtas " _
       & "WHERE TP IN ('CD','CE','CI','ND','NC') " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND Item = '" & TNumEmpresa & "' " _
       & "AND Cta = '" & Cta_SubModulo & "' " _
       & "AND T <> '" & Anulado & "' " _
       & "ORDER BY Codigo,Fecha,TP,Numero,Factura,Debitos DESC,Creditos,ID "
  Select_AdoDB AdoSubCtasDB, sSQL
  RatonReloj
  SumaDebe = 0: SumaHaber = 0
  With AdoSubCtasDB
  If .RecordCount > 0 Then
     .MoveFirst
      Cod_Cta = .fields("Codigo")
      FechaTexto = .fields("Fecha")
      Mifecha = .fields("Fecha")
      SumaCta = 0: Suma_ME = 0
      Do While Not .EOF
         If Cod_Cta <> .fields("Codigo") Then
           'Determinamos que la cuenta ya fue mayorizada
            Cod_Cta = .fields("Codigo")
            Mifecha = .fields("Fecha")
            SumaCta = 0: Suma_ME = 0
         End If
         Debe = Redondear(.fields("Debitos"), 2)
         Haber = Redondear(.fields("Creditos"), 2)
         If .fields("Parcial_ME") >= 0 Then
             Debe_ME = Redondear(.fields("Parcial_ME"), 2)
             Haber_ME = 0
         Else
             Haber_ME = Redondear(-.fields("Parcial_ME"), 2)
             Debe_ME = 0
         End If
         Mayorizar_Saldos .fields("Cta")
        'MsgBox .Fields("ID")
         If ((SumaCta <> .fields("Saldo_MN")) Or (Suma_ME <> .fields("Saldo_ME"))) Then
            .fields("Saldo_MN") = SumaCta
            .fields("Saldo_ME") = Suma_ME
            .Update
         End If
        .MoveNext
      Loop
  End If
  End With
  AdoSubCtasDB.Close
End Sub

Public Function Leer_Codigo_Inv(CodigoDeInv As String, _
                                FechaInventario As String, _
                                Optional CodBodega As String, _
                                Optional CodMarca As String) As Boolean
Dim SQL As String
Dim BuscarCodigoInv As String
Dim Codigo_Ok As Boolean
Dim Con_Kardex As Boolean
Dim AdoCodigo As ADODB.Recordset
  
  RatonReloj
 'Datos por default
  If CodBodega = "" Then CodBodega = Ninguno
  If CodMarca = "" Then CodMarca = Ninguno
  Codigo_Ok = False
  Con_Kardex = False
  DatInv.Stock = 0
  DatInv.Costo = 0
  DatInv.Con_Kardex = False
  DatInv.Codigo_Barra = Ninguno
  DatInv.Tipo_SubMod = Ninguno
  DatInv.Cta_Inventario = Ninguno
  DatInv.Cta_Costo_Venta = Ninguno
  DatInv.Codigo_Inv = CodigoDeInv
  DatInv.Fecha_Stock = FechaInventario
  DatInv.Serie_No = Ninguno
  
 'Validacion de datos correctos
  If Not IsDate(DatInv.Fecha_Stock) Then DatInv.Fecha_Stock = FechaSistema
  If Len(DatInv.TC) <= 1 Then DatInv.TC = "FA"
  BuscarCodigoInv = CodigoDeInv
  Leer_Codigo_Inv_SP BuscarCodigoInv, DatInv.Fecha_Stock, CodBodega, CodMarca, DatInv.Codigo_Inv
  
 '-----------------------------------------------------------------
 'Si existe el producto pasamos a recolectar los datos del producto
 '-----------------------------------------------------------------
  If DatInv.Codigo_Inv <> Ninguno Then
     SQL = "SELECT Producto, Detalle, Codigo_Barra_K, Unidad, Minimo, Maximo, Cta_Inventario, Cta_Costo_Venta, Cta_Ventas, Cta_Ventas_0, Cta_Venta_Anticipada, " _
         & "Utilidad, Div, PVP_2, Por_Reservas, Reg_Sanitario, IVA, PVP, Tipo_SubMod, Stock, Costo, Valor_Unit, Con_Kardex " _
         & "FROM Catalogo_Productos " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND Codigo_Inv = '" & DatInv.Codigo_Inv & "' "
     Select_AdoDB AdoCodigo, SQL
     With AdoCodigo
      If .RecordCount > 0 Then
          DatInv.Producto = .fields("Producto")
          DatInv.Detalle = .fields("Detalle")
          DatInv.Codigo_Barra = .fields("Codigo_Barra_K")
          DatInv.Unidad = .fields("Unidad")
          DatInv.Minimo = .fields("Minimo")
          DatInv.Maximo = .fields("Maximo")
          DatInv.Cta_Inventario = .fields("Cta_Inventario")
          DatInv.Cta_Costo_Venta = .fields("Cta_Costo_Venta")
          DatInv.Cta_Ventas = .fields("Cta_Ventas")
          DatInv.Cta_Ventas_0 = .fields("Cta_Ventas_0")
          DatInv.Cta_Venta_Anticipada = .fields("Cta_Venta_Anticipada")
          DatInv.Utilidad = .fields("Utilidad")
          DatInv.Div = .fields("Div")
          DatInv.PVP2 = .fields("PVP_2")
          DatInv.Por_Reservas = .fields("Por_Reservas")
          DatInv.Reg_Sanitario = .fields("Reg_Sanitario")
          DatInv.Stock = .fields("Stock")
          DatInv.Costo = Redondear(.fields("Costo"), Dec_Costo)
          DatInv.Valor_Unit = Redondear(.fields("Valor_Unit"), Dec_Costo)
          DatInv.Tipo_SubMod = .fields("Tipo_SubMod")
          DatInv.Con_Kardex = .fields("Con_Kardex")
          Select Case DatInv.TC
            Case "NV", "PV"
                 If .fields("IVA") Then DatInv.PVP = .fields("PVP") * (1 + Porc_IVA) Else DatInv.PVP = .fields("PVP")
                 DatInv.IVA = False
            Case Else
                 DatInv.PVP = .fields("PVP")
                 DatInv.IVA = .fields("IVA")
          End Select
          'If Len(DatInv.Cta_Inventario) <= 1 Then DatInv.Stock = 1
          Codigo_Ok = True
      Else
          MsgBox "Producto no Asignado"
      End If
     End With
     AdoCodigo.Close
  Else
     MsgBox "No existen datos"
  End If
  RatonNormal
  Leer_Codigo_Inv = Codigo_Ok
End Function

'''Public Function Leer_Codigo_Inv(Codigo_de_Inv As String, _
'''                                FechaInventario As String, _
'''                                Optional CodBodega As String, _
'''                                Optional CodMarca As String) As Boolean
'''Dim SQL As String
'''Dim Codigo_Ok As Boolean
'''Dim Con_Kardex As Boolean
'''Dim AdoStock As ADODB.Recordset
'''Dim AdoCodigo As ADODB.Recordset
'''
'''  RatonReloj
''' 'Datos por default
'''  Codigo_Ok = False
'''  Con_Kardex = False
'''  DatInv.Stock = 0
'''  DatInv.Costo = 0
'''  DatInv.Codigo_Barra = Ninguno
'''  DatInv.Tipo_SubMod = Ninguno
'''  DatInv.Cta_Inventario = Ninguno
'''  DatInv.Cta_Costo_Venta = Ninguno
'''  DatInv.Codigo_Inv = Ninguno
'''  DatInv.Fecha_Stock = FechaInventario
'''
''' 'Validacion de datos correctos
'''  If Not IsDate(DatInv.Fecha_Stock) Then DatInv.Fecha_Stock = FechaSistema
'''  If Len(DatInv.TC) <= 1 Then DatInv.TC = "FA"
'''
'''  Leer_Codigo_Inv_SP DatInv.Codigo_Inv, FechaInventario, CodBodega, CodMarca
'''
''' 'Buscamos por Producto
'''  SQL = "SELECT Codigo_Inv " _
'''      & "FROM Catalogo_Productos " _
'''      & "WHERE Item = '" & NumEmpresa & "' " _
'''      & "AND Periodo = '" & Periodo_Contable & "' " _
'''      & "AND Producto LIKE '" & Codigo_de_Inv & "' "
'''  Select_AdoDB AdoCodigo, SQL
'''  If AdoCodigo.RecordCount > 0 Then DatInv.Codigo_Inv = AdoCodigo.Fields("Codigo_Inv")
'''  AdoCodigo.Close
'''
''' 'Por Codigo_Inv
'''  If DatInv.Codigo_Inv = Ninguno Then
'''     SQL = "SELECT Codigo_Inv " _
'''         & "FROM Catalogo_Productos " _
'''         & "WHERE Item = '" & NumEmpresa & "' " _
'''         & "AND Periodo = '" & Periodo_Contable & "' " _
'''         & "AND Codigo_Inv = '" & Codigo_de_Inv & "' "
'''     Select_AdoDB AdoCodigo, SQL
'''     If AdoCodigo.RecordCount > 0 Then DatInv.Codigo_Inv = AdoCodigo.Fields("Codigo_Inv")
'''     AdoCodigo.Close
'''  End If
'''
''' 'Por Codigo_Inv Izquierdo
'''  If DatInv.Codigo_Inv = Ninguno Then
'''     SQL = "SELECT Codigo_Inv " _
'''         & "FROM Catalogo_Productos " _
'''         & "WHERE Item = '" & NumEmpresa & "' " _
'''         & "AND Periodo = '" & Periodo_Contable & "' " _
'''         & "AND Codigo_Inv = '" & SinEspaciosIzq(Codigo_de_Inv) & "' "
'''     Select_AdoDB AdoCodigo, SQL
'''     If AdoCodigo.RecordCount > 0 Then DatInv.Codigo_Inv = AdoCodigo.Fields("Codigo_Inv")
'''     AdoCodigo.Close
'''  End If
'''
''' 'Por Codigo_Inv Derecho
'''  If DatInv.Codigo_Inv = Ninguno Then
'''     SQL = "SELECT Codigo_Inv " _
'''         & "FROM Catalogo_Productos " _
'''         & "WHERE Item = '" & NumEmpresa & "' " _
'''         & "AND Periodo = '" & Periodo_Contable & "' " _
'''         & "AND Codigo_Inv = '" & SinEspaciosDer(Codigo_de_Inv) & "' "
'''     Select_AdoDB AdoCodigo, SQL
'''     If AdoCodigo.RecordCount > 0 Then DatInv.Codigo_Inv = AdoCodigo.Fields("Codigo_Inv")
'''     AdoCodigo.Close
'''  End If
'''
''' 'Por Codigo de Barra en Kardex
'''  If DatInv.Codigo_Inv = Ninguno Then
'''     sSQL = "SELECT Codigo_Inv, SUM(Entrada-Salida) As TStock " _
'''          & "FROM Trans_Kardex " _
'''          & "WHERE Item = '" & NumEmpresa & "' " _
'''          & "AND Periodo = '" & Periodo_Contable & "' " _
'''          & "AND Fecha <= #" & BuscarFecha(DatInv.Fecha_Stock) & "# " _
'''          & "AND Codigo_Barra = '" & Codigo_de_Inv & "' " _
'''          & "AND T <> 'A' " _
'''          & "GROUP BY Codigo_Inv "
'''     Select_AdoDB AdoCodigo, sSQL
'''    'MsgBox AdoCodigo.RecordCount
'''     If AdoCodigo.RecordCount > 0 Then
'''        DatInv.Codigo_Barra = Codigo_de_Inv
'''        DatInv.Codigo_Inv = AdoCodigo.Fields("Codigo_Inv")
'''     End If
'''     AdoCodigo.Close
'''  End If
'''
''' 'Por Codigo de Barra
'''  If DatInv.Codigo_Inv = Ninguno Then
'''     SQL = "SELECT Codigo_Inv " _
'''         & "FROM Catalogo_Productos " _
'''         & "WHERE Item = '" & NumEmpresa & "' " _
'''         & "AND Periodo = '" & Periodo_Contable & "' " _
'''         & "AND Codigo_Barra = '" & Codigo_de_Inv & "' "
'''     Select_AdoDB AdoCodigo, SQL
'''     If AdoCodigo.RecordCount > 0 Then DatInv.Codigo_Inv = AdoCodigo.Fields("Codigo_Inv")
'''     AdoCodigo.Close
'''  End If
'''
''' '-----------------------------------------------------------------
''' 'Si existe el producto pasamos a recolectar los datos del producto
''' '-----------------------------------------------------------------
'''  If DatInv.Codigo_Inv <> Ninguno Then
'''     SQL = "SELECT * " _
'''         & "FROM Catalogo_Productos " _
'''         & "WHERE Item = '" & NumEmpresa & "' " _
'''         & "AND Periodo = '" & Periodo_Contable & "' " _
'''         & "AND Codigo_Inv = '" & DatInv.Codigo_Inv & "' "
'''     Select_AdoDB AdoCodigo, SQL
'''     With AdoCodigo
'''      If .RecordCount > 0 Then
'''          Codigo_Ok = True
'''          DatInv.Producto = .Fields("Producto")
'''          DatInv.Detalle = .Fields("Detalle")
'''          If DatInv.Codigo_Barra = Ninguno Then DatInv.Codigo_Barra = .Fields("Codigo_Barra")
'''          DatInv.Unidad = .Fields("Unidad")
'''          DatInv.Minimo = .Fields("Minimo")
'''          DatInv.Maximo = .Fields("Maximo")
'''          DatInv.Cta_Inventario = .Fields("Cta_Inventario")
'''          DatInv.Cta_Costo_Venta = .Fields("Cta_Costo_Venta")
'''          DatInv.Cta_Ventas = .Fields("Cta_Ventas")
'''          DatInv.Cta_Ventas_0 = .Fields("Cta_Ventas_0")
'''          DatInv.Cta_Venta_Anticipada = .Fields("Cta_Venta_Anticipada")
'''          DatInv.Utilidad = .Fields("Utilidad")
'''          DatInv.Div = .Fields("Div")
'''          DatInv.PVP2 = .Fields("PVP_2")
'''          DatInv.Por_Reservas = .Fields("Por_Reservas")
'''          DatInv.Reg_Sanitario = .Fields("Reg_Sanitario")
'''
'''          If CodBodega = "" Then CodBodega = Ninguno
'''          If CodMarca = "" Then CodMarca = Ninguno
'''          Cta_Ventas = DatInv.Cta_Ventas
'''          Select Case DatInv.TC
'''            Case "NV", "PV"
'''                 If .Fields("IVA") Then DatInv.PVP = .Fields("PVP") * (1 + Porc_IVA) Else DatInv.PVP = .Fields("PVP")
'''                 DatInv.IVA = False
'''            Case Else
'''                 DatInv.PVP = .Fields("PVP")
'''                 DatInv.IVA = .Fields("IVA")
'''          End Select
'''
'''         'Sacamos el stock del producto
'''          sSQL = "SELECT Codigo_Inv, SUM(Entrada-Salida) As TStock " _
'''               & "FROM Trans_Kardex " _
'''               & "WHERE Item = '" & NumEmpresa & "' " _
'''               & "AND Periodo = '" & Periodo_Contable & "' " _
'''               & "AND Fecha <= #" & BuscarFecha(DatInv.Fecha_Stock) & "# " _
'''               & "AND Codigo_Inv = '" & DatInv.Codigo_Inv & "' " _
'''               & "AND T <> 'A' "
'''          If Len(CodBodega) > 1 Then sSQL = sSQL & "AND CodBodega = '" & CodBodega & "' "
'''          If Len(CodMarca) > 1 Then sSQL = sSQL & "AND CodMarca = '" & CodMarca & "' "
'''          sSQL = sSQL & "GROUP BY Codigo_Inv "
'''          Select_AdoDB AdoStock, sSQL
'''          If AdoStock.RecordCount > 0 Then If Not IsNull(AdoStock.Fields("TStock")) Then DatInv.Stock = AdoStock.Fields("TStock")
'''          AdoStock.Close
'''
'''         'Sacamos el costo del producto
'''          sSQL = "SELECT TOP 1 Codigo_Inv,Costo,Valor_Unitario,Existencia,Total,T " _
'''               & "FROM Trans_Kardex " _
'''               & "WHERE Item = '" & NumEmpresa & "' " _
'''               & "AND Periodo = '" & Periodo_Contable & "' " _
'''               & "AND Fecha <= #" & BuscarFecha(DatInv.Fecha_Stock) & "# " _
'''               & "AND Codigo_Inv = '" & DatInv.Codigo_Inv & "' " _
'''               & "AND T <> 'A' " _
'''               & "ORDER BY Fecha DESC,TP DESC, Numero DESC,ID DESC "
'''          Select_AdoDB AdoStock, sSQL
'''          If AdoStock.RecordCount > 0 Then
'''             DatInv.Costo = Redondear(AdoStock.Fields("Costo"), Dec_Costo)
'''             DatInv.Valor_Unit = Redondear(AdoStock.Fields("Valor_Unitario"), Dec_Costo)
'''          End If
'''
'''          If Len(DatInv.Cta_Inventario) <= 1 Then DatInv.Stock = 1
'''
'''          sSQL = "SELECT Codigo, TC " _
'''               & "FROM Catalogo_Cuentas " _
'''               & "WHERE Item = '" & NumEmpresa & "' " _
'''               & "AND Periodo = '" & Periodo_Contable & "' " _
'''               & "AND Codigo IN ('" & DatInv.Cta_Ventas & "','" & DatInv.Cta_Ventas_0 & "') " _
'''               & "AND TC <> 'N' " _
'''               & "ORDER BY Codigo "
'''          Select_AdoDB AdoStock, sSQL
'''          If AdoStock.RecordCount > 0 Then
'''             Do While Not AdoStock.EOF
'''                DatInv.Tipo_SubMod = AdoStock.Fields("TC")
'''                AdoStock.MoveNext
'''             Loop
'''          End If
'''          AdoStock.Close
'''      Else
'''          MsgBox "Producto no Asignado"
'''      End If
'''     End With
'''     AdoCodigo.Close
'''  Else
'''     MsgBox "No existen datos"
'''  End If
'''  RatonNormal
'''  Leer_Codigo_Inv = Codigo_Ok
'''End Function

Public Sub EncabezadoSaldosTemp()
PosLinea = PosLinea + 0.05
Printer.FontSize = 10
Printer.FontBold = False
PrinterTexto Ancho(0), PosLinea, "CLIENTE: " & NombreCliente
Printer.FontSize = 9
Printer.FontBold = True
PosLinea = PosLinea + 0.4
'PrinterTexto Ancho(0), PosLinea, "T"
PrinterTexto Ancho(1), PosLinea, "Fecha"
PrinterTexto Ancho(2), PosLinea, "Fecha_Venc"
PrinterTexto Ancho(3), PosLinea, "Factura"
PrinterTexto Ancho(4), PosLinea, "Total"
PrinterTexto Ancho(5), PosLinea, "Vencido"
PrinterTexto Ancho(6), PosLinea, "Ven_1_a_30"
PrinterTexto Ancho(7), PosLinea, "Ven_31_a_60"
PrinterTexto Ancho(8), PosLinea, "Ven_61_a_90"
PrinterTexto Ancho(9), PosLinea, "Mas_91"
PrinterTexto Ancho(10), PosLinea, "Cta"
Printer.FontBold = False
PosLinea = PosLinea + 0.36
Imprimir_Linea_H PosLinea, Ancho(0), Ancho(CantCampos)
PosLinea = PosLinea + 0.05
End Sub

Public Sub Encab_Estdo_Cuentas()
  PorteLetra = Printer.FontSize
  LetraAnterior = Printer.FontName
  Printer.FontName = TipoTimes
  Printer.FontBold = True
  InicioY = PosLinea
 'Iniciamos la impresion
  Printer.FontSize = 10
  PosLinea = PosLinea + 0.05
  PrinterTexto Ancho(0), PosLinea, NombreCliente & ":"
  Printer.FontUnderline = False
  PrinterTexto Ancho(6), PosLinea, "PROGRAMA: [" & Contrato_No & "]"
  Printer.FontUnderline = True
  PosLinea = PosLinea + 0.45
  Printer.FontSize = 9
  PrinterTexto Ancho(1), PosLinea, "Tipo Rubro"
  PrinterTexto Ancho(2), PosLinea, "Fecha"
  PrinterTexto Ancho(3), PosLinea, "D e t a l l e"
  PrinterTexto Ancho(4), PosLinea, "Recibo"
  PrinterTexto Ancho(6), PosLinea, "Por Cobrar"
  PrinterTexto Ancho(7), PosLinea, "A b o n o"
  PrinterTexto Ancho(8), PosLinea, "Saldo_Actual"
  Printer.FontBold = False
  Printer.FontUnderline = False
  PosLinea = PosLinea + 0.5
  Printer.FontSize = PorteLetra
  Printer.FontName = LetraAnterior
End Sub

'''Public Sub Insertar_Saldo_Facturas_CxCxP()
''' 'TipoCta = .Fields("TC")
''' 'FechaTexto = .Fields("Fecha")
''' 'FechaTexto1 = .Fields("Fecha")
''' 'MiFecha = .Fields("Fecha_V")
''' 'Cta = .Fields("Cta")
''' 'Codigo = .Fields("Codigo")
''' 'Factura_No = .Fields("Factura")
''' 'Debe = 0: Haber = 0
'''  T = "P"
'''  Select Case TipoCta
'''    Case "C"
'''         Total = Debe
'''         Saldo = Debe - Haber
'''    Case "P"
'''         Total = Haber
'''         Saldo = Haber - Debe
'''  End Select
'''  If Saldo = 0 Then T = "C"
'''  SetAdoAddNew "Facturas"
'''  SetAdoFields "T", T
'''  SetAdoFields "TC", TipoCta
'''  SetAdoFields "Factura", Factura_No
'''  SetAdoFields "CodigoC", Codigo
'''  SetAdoFields "Fecha", FechaTexto
'''  SetAdoFields "Fecha_C", FechaTexto1
'''  SetAdoFields "Fecha_V", Mifecha
'''  SetAdoFields "SubTotal", Total
'''  SetAdoFields "Total_MN", Total
'''  SetAdoFields "Saldo_MN", Saldo
'''  SetAdoFields "Cta_CxP", Cta
'''  SetAdoFields "Comercial", Beneficiario
'''  SetAdoFields "Item", NumEmpresa
'''  SetAdoFields "CodigoU", CodigoUsuario
'''  SetAdoFields "Periodo", Periodo_Contable
'''  SetAdoUpdate
'''End Sub

'''Public Sub Saldo_Facturas_CxCxP(FechaI As String, _
'''                                FechaF As String)
'''Dim AdoDBRecordset As ADODB.Recordset
'''Dim Proceso_de_Barras As Progreso_Barras
'''Dim Fecha_I As String
'''Dim Fecha_F As String
'''  Contador = 0
'''  Progreso_Iniciar
'''  Fecha_I = BuscarFecha(FechaI)
'''  Fecha_F = BuscarFecha(FechaF)
'''  Set AdoDBRecordset = New ADODB.Recordset
'''  AdoDBRecordset.CursorType = adOpenStatic
'''  AdoDBRecordset.CursorLocation = adUseClient
'''  sSQL = "DELETE * " _
'''       & "FROM Facturas " _
'''       & "WHERE Item = '" & NumEmpresa & "' " _
'''       & "AND Periodo = '" & Periodo_Contable & "' "
'''  If Cta_Sup <> Ninguno Then sSQL = sSQL & "AND Cta_CxP = '" & Cta_Sup & "' "
'''  If CodigoB <> Ninguno Then sSQL = sSQL & "AND CodigoC = '" & CodigoB & "' "
'''  Select Case TipoCta
'''    Case "C", "P"
'''         sSQL = sSQL & "AND TC = '" & TipoCta & "' "
'''    Case Else
'''         sSQL = sSQL & "AND TC IN ('C','P') "
'''  End Select
'''  Ejecutar_SQL_SP sSQL
''' 'SubCtas
'''  sSQL = "SELECT TC,Cta,Codigo,Factura,Fecha,Fecha_V,Debitos,Creditos,Detalle_SubCta " _
'''       & "FROM Trans_SubCtas " _
'''       & "WHERE Item = '" & NumEmpresa & "' " _
'''       & "AND Periodo = '" & Periodo_Contable & "' " _
'''       & "AND Fecha <= #" & Fecha_F & "# "
'''  If Cta_Sup <> Ninguno Then sSQL = sSQL & "AND Cta = '" & Cta_Sup & "' "
'''  If CodigoB <> Ninguno Then sSQL = sSQL & "AND Codigo = '" & CodigoB & "' "
'''  Select Case TipoCta
'''    Case "C", "P"
'''         sSQL = sSQL & "AND TC = '" & TipoCta & "' "
'''    Case Else
'''         sSQL = sSQL & "AND TC IN ('C','P') "
'''  End Select
'''  sSQL = sSQL & "ORDER BY TC,Cta,Codigo,Factura,Fecha,Fecha_V "
'''  sSQL = CompilarSQL(sSQL)
'''  'MsgBox sSQL
'''  AdoDBRecordset.Open sSQL, AdoStrCnn, , , adCmdText
'''  With AdoDBRecordset
'''   If .RecordCount > 0 Then
'''       TipoCta = .Fields("TC")
'''       FechaTexto = .Fields("Fecha")
'''       FechaTexto1 = .Fields("Fecha")
'''       Mifecha = .Fields("Fecha_V")
'''       Cta = .Fields("Cta")
'''       Codigo = .Fields("Codigo")
'''       Factura_No = .Fields("Factura")
'''       Beneficiario = .Fields("Detalle_SubCta")
'''       Debe = 0: Haber = 0
'''       'MsgBox sSQL
'''       RatonReloj
'''       Proceso_de_Barras.Valor_Maximo = .RecordCount
'''       Proceso_de_Barras.Mensaje_Box = "SALDO DE FACTURAS"
'''       Progreso_Esperar
'''       Do While Not .EOF
'''          'SaldoSubCtasVence.Caption = SQLMsg1 & " - " & Format$(Contador / .RecordCount, "00%")
'''          Contador = Contador + 1
'''          'MiFecha = .Fields("Fecha_V")
'''          If Cta <> .Fields("Cta") Or _
'''             Codigo <> .Fields("Codigo") Or _
'''             Factura_No <> .Fields("Factura") Or _
'''             TipoCta <> .Fields("TC") Then
'''             Insertar_Saldo_Facturas_CxCxP
'''             TipoCta = .Fields("TC")
'''             FechaTexto = .Fields("Fecha")
'''             FechaTexto1 = .Fields("Fecha")
'''             Mifecha = .Fields("Fecha_V")
'''             Cta = .Fields("Cta")
'''             Codigo = .Fields("Codigo")
'''             Factura_No = .Fields("Factura")
'''             Beneficiario = .Fields("Detalle_SubCta")
'''             Debe = 0: Haber = 0
'''          End If
'''          FechaTexto1 = .Fields("Fecha")
'''          Debe = Debe + .Fields("Debitos")
'''          Haber = Haber + .Fields("Creditos")
'''          Progreso_Esperar
'''         .MoveNext
'''       Loop
'''       Insertar_Saldo_Facturas_CxCxP
'''   End If
'''  End With
''''  Unload FHola
'''  Progreso_Final
'''End Sub

Public Sub Imprimir_Estado_Cuentas(Datas As Adodc, _
                                  FormaImp As Byte, _
                                  SizeLetra As Integer, _
                                  Optional EsCampoCorto As Boolean)
On Error GoTo Errorhandler
RatonReloj
Mensajes = "Seguro de Imprimir en:" & vbCrLf & Printer.DeviceName & "?"
Titulo = "IMPRESION"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
InicioX = 1: InicioY = 0
DataAnchoCampos InicioX, Datas, SizeLetra, TipoTimes, FormaImp, EsCampoCorto
Ancho(0) = 1      ' Cliente
Ancho(1) = 1      ' Contrato
Ancho(2) = 4      ' Fecha
Ancho(3) = 5.5    ' Detalle
Ancho(4) = 11.5   ' Recibo_No
Ancho(5) = 11     ' IVA
Ancho(6) = 13     ' Ventas
Ancho(7) = 15     ' Abono
Ancho(8) = 17     ' Saldo_Actual
Ancho(9) = 19     ' Fin
Pagina = 1
Total = 0
'Iniciamos la impresion
Printer.FontBold = False
With Datas.Recordset
    .MoveFirst
     Encabezado Ancho(0), Ancho(9)
     NombreCliente = .fields("Cliente")
     Contrato_No = .fields("Contrato_No")
     CodigoInv = .fields("Producto")
     Printer.FontSize = 10
     PrinterTexto Ancho(0), PosLinea, SQLMsg1
     PosLinea = PosLinea + 0.45
     Encab_Estdo_Cuentas
     Printer.FontSize = SizeLetra
     PrinterTexto Ancho(1), PosLinea, UCaseStrg(CodigoInv) & ":"
     PosLinea = PosLinea + 0.4
     Do While Not .EOF
        If NombreCliente <> .fields("Cliente") Or Contrato_No <> .fields("Contrato_No") Then
           Imprimir_Linea_H PosLinea, Ancho(0), Ancho(9)
           PosLinea = PosLinea + 0.05
           PrinterTexto Ancho(7), PosLinea, "SALDO"
           PrinterVariables Ancho(8), PosLinea, Saldo
           Total = Total + Saldo
           PosLinea = PosLinea + 0.4
           Imprimir_Linea_H PosLinea, Ancho(0), Ancho(9), Negro, True
           PosLinea = PosLinea + 0.05
           PrinterTexto Ancho(6) + 1, PosLinea, "SALDO TOTAL"
           PrinterVariables Ancho(8), PosLinea, Total

           Printer.NewPage
           Encabezado Ancho(0), Ancho(9)
           Printer.FontSize = 10
           PrinterTexto Ancho(0), PosLinea, SQLMsg1
           PosLinea = PosLinea + 0.45
           NombreCliente = .fields("Cliente")
           Contrato_No = .fields("Contrato_No")
           CodigoInv = .fields("Producto")
           Encab_Estdo_Cuentas
           Printer.FontSize = SizeLetra
           PrinterTexto Ancho(1), PosLinea, UCaseStrg(CodigoInv) & ":"
           PosLinea = PosLinea + 0.4
           Total = 0
        End If
        If CodigoInv <> .fields("Producto") Then
           CodigoInv = .fields("Producto")
           Imprimir_Linea_H PosLinea, Ancho(0), Ancho(9)
           PosLinea = PosLinea + 0.05
           PrinterTexto Ancho(7), PosLinea, "SALDO"
           PrinterVariables Ancho(8), PosLinea, Saldo
           Total = Total + Saldo
           PosLinea = PosLinea + 0.4
           PrinterTexto Ancho(1), PosLinea, UCaseStrg(CodigoInv) & ":"
           PosLinea = PosLinea + 0.4
        End If
        PrinterFields Ancho(2), PosLinea, .fields("Fecha")
        PrinterFields Ancho(3), PosLinea, .fields("Detalle")
        PrinterFields Ancho(4), PosLinea, .fields("Comprobante")
        PrinterFields Ancho(6), PosLinea, .fields("Por_Cobrar")
        PrinterFields Ancho(7), PosLinea, .fields("Abono")
        PrinterFields Ancho(8), PosLinea, .fields("Saldo")
        Saldo = .fields("Saldo")
        PosLinea = PosLinea + 0.36
        If PosLinea >= LimiteAlto Then
           Imprimir_Linea_H PosLinea, Ancho(0), Ancho(9)
           Printer.NewPage
           Encabezado Ancho(0), Ancho(9)
           Printer.FontSize = 10
           PrinterTexto Ancho(0), PosLinea, SQLMsg1
           PosLinea = PosLinea + 0.45
           Contrato_No = .fields("Contrato_No")
           CodigoInv = .fields("Producto")
           Encab_Estdo_Cuentas
           Printer.FontSize = SizeLetra
           PrinterFields Ancho(1), PosLinea, .fields("Contrato_No")
        End If
        'If Pagina > 20 Then Exit Do
       .MoveNext
     Loop
End With
Imprimir_Linea_H PosLinea, Ancho(0), Ancho(9)
PosLinea = PosLinea + 0.05
PrinterTexto Ancho(7), PosLinea, "SALDO"
PrinterVariables Ancho(8), PosLinea, Saldo
PosLinea = PosLinea + 0.4
Total = Total + Saldo
Imprimir_Linea_H PosLinea, Ancho(0), Ancho(9), Negro, True
PosLinea = PosLinea + 0.05
PrinterTexto Ancho(6) + 1, PosLinea, "SALDO TOTAL"
PrinterVariables Ancho(8), PosLinea, Total
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

Public Sub ImprimirPlanillaIESS(Datas As Adodc, _
                                FinDoc As Boolean, _
                                FormaImp As Byte, _
                                SizeLetra As Integer, _
                                Optional EsCampoCorto As Boolean)
On Error GoTo Errorhandler
RatonReloj
Mensajes = "Seguro de Imprimir en:" & vbCrLf & Printer.DeviceName & "?"
Titulo = "IMPRESION"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
InicioX = 0.5: InicioY = 0
DataAnchoCampos InicioX, Datas, SizeLetra, TipoTimes, FormaImp, EsCampoCorto
Pagina = 1
'Iniciamos la impresion
Printer.FontBold = False
With Datas.Recordset
    .MoveFirst
     Dibujo = RutaSistema & "\FORMATOS\LIBBANCO.GIF"
     PrinterPaint Dibujo, 0.5, PosLinea, 25, 1
     Printer.FontSize = SizeLetra
     Do While Not .EOF
        'PrinterFields
        PosLinea = PosLinea + 0.36
        If PosLinea >= LimiteAlto Then
           Imprimir_Linea_H PosLinea, Ancho(0), Ancho(CantCampos)
           Printer.NewPage
           EncabezadoData Datas
           Printer.FontSize = SizeLetra
        End If
       .MoveNext
     Loop
End With
Imprimir_Linea_H PosLinea, Ancho(0), Ancho(CantCampos), Negro, True
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

Public Sub GenerarTablaPrestamoCxCP(BoxMiFecha As String, _
                                    DtaTabla As Adodc, _
                                    DBG_Tabla As DataGrid, _
                                    TxtInt As TextBox, _
                                    TxtMeses As TextBox, _
                                    TxtMonto As TextBox, _
                                    TipoPrest As String, _
                                    Optional SinComis As Boolean)
  Interes = Redondear(CCur(Val(CCur(TxtInt.Text))) / 100, 4)
  'MsgBox Interes
  Numero = CInt(Val(TxtMeses.Text))
  Total = Redondear(Val(CCur(TxtMonto.Text)), 2)
  Mifecha = BoxMiFecha
  Saldo = Total:  Valor_ME = 0:  Total_ME = 0:  Valor = 0
  sSQL = "DELETE * " _
       & "FROM Asiento_P " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' "
  Ejecutar_SQL_SP sSQL
  
  sSQL = "SELECT * " _
       & "FROM Asiento_P " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' "
  Select_Adodc_Grid DBG_Tabla, DtaTabla, sSQL
  DBG_Tabla.Visible = False
  With DtaTabla.Recordset
       Total = Redondear(Total + (Total * ((Numero / 12) * Interes)), 2)
       Tasa = 0
       If Numero <= 0 Then Numero = 1
       Do
         Tasa = Redondear(Tasa + 0.0001, 4)
         Cuota = Redondear(((Saldo * Tasa) / 12) / (1 - (1 + (Tasa / 12)) ^ -Numero), 2)
       Loop Until (Cuota * Numero) >= Total
       Contador = 1: Total = Saldo
       Valor = Redondear(((12 * Total) + (Total * Interes * Numero)) / (12 * Numero), 2)
       Valor_ME = 0: Total_ME = 0: Comision = 0
       For I = 0 To Numero
           SetAddNew DtaTabla
          .fields("Cuotas") = I
          '.Fields("Dia") = Weekday(MiFecha)
           If I = 0 Then
             .fields("Fecha") = Mifecha
             .fields("Capital") = 0
             .fields("Interes") = 0
             .fields("Comision") = 0
             .fields("Pagos") = 0
              Mifecha = SiguienteMes(Mifecha)
           Else
              If I = Numero Then
                 'MsgBox Saldo
                 Total_ME = Saldo
                 Valor_ME = Valor - Saldo
                 If Interes <= 0 Then Valor_ME = 0
                 Valor = Total_ME + Valor_ME
                 Total = 0
              End If
             .fields("Fecha") = Mifecha
             .fields("Capital") = Total_ME
             .fields("Interes") = Valor_ME
             .fields("Comision") = Comision
             .fields("Pagos") = Valor
              Mifecha = SiguienteMes(Mifecha)
           End If
          .fields("Saldo") = Total
          .fields("CodigoU") = CodigoUsuario
          .fields("Cta") = TipoPrest
          .fields("T_No") = Trans_No
           SetUpdate DtaTabla
          'Comision del 1%
           If SinComis = False Then Comision = Redondear(Total * 0.012, 2)
          'Interes Inicial
           Valor_ME = Redondear(Total * (Tasa / 12), 2)
           If Interes = 0 Then Valor_ME = 0
          'Amortizacion o Capital
           Total_ME = Redondear(Valor - Valor_ME, 2)
          'Saldo Pendiente
           Saldo = Total
           Total = Redondear(Total - Total_ME, 2)
          'Interes Final
           Valor_ME = Redondear(Valor - Total_ME - Comision, 2)
           Contador = Contador + 1
       Next I
  End With
  DBG_Tabla.Visible = True
End Sub

Public Sub Imprimir_Rol_de_Pagos(Datas As Adodc, _
                                 DataTot As Adodc, _
                                 FinDoc As Boolean, _
                                 FormaImp As Byte, _
                                 Optional EmpezarImp As Long)
Dim SizeLetra As Integer
Dim AuxPosLinea As Single
Dim VectCurr() As Double

On Error GoTo Errorhandler
RatonReloj
SizeLetra = 8
Mensajes = "Seguro de Imprimir en:" & vbCrLf & Printer.DeviceName & "?"
Titulo = "IMPRESION"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
InicioX = 0.5: InicioY = 0
'Escala_Centimetro FormaImp, TipoTimes, SizeLetra
'TipoTimes
DataAnchoCampos InicioX, Datas, SizeLetra, TipoArialNarrow, FormaImp
LimiteAlto = LimiteAlto - 0.5
LimiteAncho = LimiteAncho - 0.5
Ancho(CantCampos) = LimiteAncho
ReDim VectCurr(LimiteAncho + 1) As Double
For I = 0 To LimiteAncho
   VectCurr(I) = 0
Next I
Pagina = 1
Contador = 1
'Iniciamos la impresion
Printer.FontBold = False
With Datas.Recordset
    .MoveFirst
     If EmpezarImp > 0 Then
        Contador = EmpezarImp
        Do While Not .EOF
           EmpezarImp = EmpezarImp - 1
           If EmpezarImp > 0 Then .MoveNext Else GoTo Salir_Busq
        Loop
Salir_Busq:
     End If
     EncabezadoData Datas
     AuxPosLinea = PosLinea
     Printer.FontSize = SizeLetra
     Printer.FontName = TipoArialNarrow 'TipoCondensed
     
     Printer.FontBold = False
     Do While Not .EOF
        For I = 3 To CantCampos - 3
            If IsNumeric(.fields(I)) Then VectCurr(I) = VectCurr(I) + .fields(I)
        Next I
        PrinterTexto 0.01, PosLinea + 0.2, Val(Contador) & ".-"
        PrinterAllFields CantCampos, PosLinea + 0.2, Datas, False, False
        Contador = Contador + 1
        PosLinea = PosLinea + 0.5
        Imprimir_Linea_H PosLinea + 0.3, Ancho(0), Ancho(CantCampos)
        PosLinea = PosLinea + 0.4
        If PosLinea > LimiteAlto Then
           For I = 0 To CantCampos
               Imprimir_Linea_V Ancho(I), AuxPosLinea, PosLinea
           Next I
           Imprimir_Linea_H PosLinea + 0.3, Ancho(0), Ancho(CantCampos)
           Printer.NewPage
           'MsgBox "Fin Rol"
           EncabezadoData Datas
           AuxPosLinea = PosLinea
           Printer.FontSize = SizeLetra
           Printer.FontName = TipoArialNarrow 'TipoCondensed
           Printer.FontBold = False
        End If
       .MoveNext
     Loop
     For I = 0 To CantCampos
         Imprimir_Linea_V Ancho(I), AuxPosLinea, PosLinea
     Next I
     Imprimir_Linea_H PosLinea, Ancho(0), Ancho(CantCampos), Negro, True
End With
PosLinea = PosLinea + 0.05
For I = 0 To CantCampos - 1
    PrinterVariables Ancho(I) + 0.2, PosLinea, VectCurr(I)
Next I
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

Public Sub ImprimirResultSemana(Datas As Adodc, _
                                FinDoc As Boolean, _
                                FormaImp As Byte, _
                                SizeLetra As Integer)
On Error GoTo Errorhandler
Mensajes = "Seguro de Imprimir en:" & vbCrLf & vbCrLf _
         & Printer.DeviceName & "?"
Titulo = "IMPRESION"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
RatonReloj
InicioX = 0.5: InicioY = 0
DataAnchoCampos InicioX, Datas, SizeLetra, TipoTimes, FormaImp
Ancho(0) = 1
Ancho(1) = 3.5
Ancho(2) = 7.5
Ancho(3) = 9.5
Ancho(4) = 11.5
Ancho(5) = 13.5
Ancho(6) = 15.5
Ancho(7) = 17.5
Ancho(8) = 19.5
Pagina = 1
'Iniciamos la impresion
Printer.FontBold = False
With Datas.Recordset
     .MoveFirst
      EncabezadoData Datas
      Printer.FontSize = SizeLetra
      Do While Not .EOF
         If MidStrg(.fields("Codigo"), 1, 1) = " " Then
            Imprimir_Linea_H PosLinea, Ancho(0), Ancho(CantCampos)
            PosLinea = PosLinea + 0.05
            Printer.FontBold = True
         Else
            Printer.FontBold = False
         End If
         PrinterAllFields CantCampos, PosLinea, Datas, True, False
         PosLinea = PosLinea + 0.36
         If MidStrg(.fields("Codigo"), 1, 1) = " " Then
            Imprimir_Linea_H PosLinea, Ancho(0), Ancho(CantCampos)
            PosLinea = PosLinea + 0.05
         End If
         If PosLinea >= LimiteAlto Then
            Imprimir_Linea_H PosLinea, Ancho(0), Ancho(CantCampos)
            Printer.NewPage
            EncabezadoData Datas
            Printer.FontSize = SizeLetra
         End If
        .MoveNext
      Loop
End With
Imprimir_Linea_H PosLinea, Ancho(0), Ancho(CantCampos)
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

Public Sub Imprimir_Recibo_De_Caja(Datas As Adodc, _
                                   NumRecibo As Long, _
                                   OpcI As Boolean, _
                                   No_Empresa As String)
On Error GoTo Errorhandler
Mensajes = "Imprimir Recibo de Caja No. " & NumRecibo & vbCrLf & vbCrLf _
         & "de la Empresa No. " & No_Empresa
Titulo = "IMPRESION"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
RatonReloj
InicioX = 0.5: InicioY = 0
sSQL = "SELECT GC.Detalle,TGC.* " _
     & "FROM Trans_Gastos_Caja As TGC,Catalogo_SubCtas As GC " _
     & "WHERE TGC.Item = '" & No_Empresa & "' " _
     & "AND TGC.Periodo = '" & Periodo_Contable & "' "
If OpcI Then
   sSQL = sSQL & "AND Ingreso > 0 "
Else
   sSQL = sSQL & "AND Egreso > 0 "
End If
sSQL = sSQL & "AND Numero = " & NumRecibo & " " _
     & "AND TGC.Codigo = GC.Codigo " _
     & "AND TGC.Periodo = GC.Periodo "
Select_Adodc Datas, sSQL
DataAnchoCampos InicioX, Datas, 8, TipoTimes, 1
Pagina = 1: PosLinea = 1
'Iniciamos la impresion

'MsgBox Datas.Recordset.RecordCount & vbCrLf & sSQL

With Datas.Recordset
 If .RecordCount > 0 Then
     FA.LogoFactura = RutaSistema & "\FORMATOS\RECIBOC.GIF"
     PrinterPaint FA.LogoFactura, PosLinea, 0.5, 16, 8
     PrinterPaint LogoTipo, 1.2, PosLinea + 0.01, 3.5, 1.5
     Printer.FontSize = 10
     Printer.FontBold = True
     If .fields("Ingreso") <> 0 Then
         PrinterTexto 1.2, PosLinea + 2.5, "Recibi de: "
     Else
         PrinterTexto 1.2, PosLinea + 2.5, "Pagado a: "
     End If
     Printer.FontBold = False
     PrinterFields 14.4, PosLinea + 0.2, .fields("Fecha")
     PrinterFields 3.1, PosLinea + 2.5, .fields("Beneficiario")
     Valor = Redondear(.fields("Ingreso") + .fields("Egreso"), 2)
     Cadena = Cambio_Letras(Valor)
     NumeroLineas = PrinterLineasMayor(4, PosLinea + 3.3, Cadena, 10)
     Cadena = ""
     If Len(.fields("Detalle")) > 1 Then Cadena = Cadena & .fields("Detalle") & ": "
     If Len(.fields("Concepto")) > 1 Then Cadena = Cadena & .fields("Concepto")
     NumeroLineas = PrinterLineasMayor(4.1, PosLinea + 4.4, Cadena, 12)
     Printer.FontSize = 12
     PrinterTexto 14.5, PosLinea + 1.4, Format$(.fields("Numero"), "0000000")
     PrinterVariables 13.5, PosLinea + 2.5, Valor
     Printer.FontBold = True
     If Empresa = NombreComercial Then
        PrinterTexto 4, PosLinea + 0.01, Empresa
     Else
        PrinterTexto 4, PosLinea + 0.01, Empresa
        PrinterTexto 4, PosLinea + 0.5, NombreComercial
     End If
     Printer.FontSize = 8: Printer.FontBold = False
     PrinterTexto 4, PosLinea + 1, "Dir. " & Direccion
     Printer.FontSize = 12
     Printer.FontBold = True
     If .fields("Ingreso") <> 0 Then
         PrinterTexto 5.1, PosLinea + 1.4, "COMPROBANTE DE INGRESO DE CAJA"
     Else
         PrinterTexto 5.1, PosLinea + 1.4, "COMPROBANTE DE EGRESO DE CAJA"
     End If
     
 End If
End With
Printer.FontBold = False
MensajeEncabData = ""
Printer.EndDoc
RatonNormal
End If
Exit Sub
Errorhandler:
    RatonNormal
    ErrorDeImpresion
    Exit Sub
End Sub

Public Sub Imprimir_Libro_Dependencia(Datas As Adodc, _
                                      SizeLetra As Integer, _
                                      Optional EsCampoCorto As Boolean)
On Error GoTo Errorhandler
Mensajes = "Seguro de Imprimir en:" & vbCrLf & Printer.DeviceName & "?"
Titulo = "IMPRESION"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
RatonReloj
InicioX = 1: InicioY = 0
DataAnchoCampos InicioX, Datas, SizeLetra, TipoTimes, 1
Ancho(0) = 1    'Fecha
Ancho(1) = 1    'Cliente
Ancho(2) = 7.3  'CI_RUC
Ancho(3) = 9    'Salario_Neto
Ancho(4) = 9.5  'Aporte_Personal
Ancho(5) = 11.5 'Porcentaje_Aporte
Ancho(6) = 13   'Ingresos_Liquidos
Ancho(7) = 15   'Base_Imponible
Ancho(8) = 17   'Retenido
Ancho(9) = 19
Pagina = 1
'Iniciamos la impresion
Printer.FontBold = False
With Datas.Recordset
    .MoveFirst
     EncabezadoData Datas
     Printer.FontSize = SizeLetra
     Mifecha = .fields("Fecha")
     PrinterTexto Ancho(1), PosLinea, " (*) MES DE: " & UCaseStrg(MesesLetras(FechaMes(Mifecha)))
     Imprimir_Lineas_Campos PosLinea
     PosLinea = PosLinea + 0.36
     Do While Not .EOF
        If Mifecha <> .fields("Fecha") Then
           Mifecha = .fields("Fecha")
           PrinterTexto Ancho(1), PosLinea, " (*) MES DE: " & UCaseStrg(MesesLetras(FechaMes(Mifecha)))
           Imprimir_Lineas_Campos PosLinea
           PosLinea = PosLinea + 0.36
        End If
        PrinterFields Ancho(1), PosLinea, .fields("Cliente")
        PrinterFields Ancho(2), PosLinea, .fields("Cedula")
        Si_No = .fields("SN")
        If Si_No Then
           PrinterTexto Ancho(3), PosLinea, "Si"
        Else
           PrinterTexto Ancho(3), PosLinea, "No"
        End If
        PrinterFields Ancho(4), PosLinea, .fields("Aporte_Per")
        PrinterFields Ancho(5), PosLinea, .fields("Porc_Aporte")
        PrinterFields Ancho(6), PosLinea, .fields("Ingresos_Liq")
        PrinterFields Ancho(7), PosLinea, .fields("Base_Imponible")
        PrinterFields Ancho(8), PosLinea, .fields("Retenido")
        Imprimir_Lineas_Campos PosLinea
        PosLinea = PosLinea + 0.36
        If PosLinea >= LimiteAlto Then
           Imprimir_Linea_H PosLinea, Ancho(0), Ancho(CantCampos)
           Printer.NewPage
           EncabezadoData Datas
           Printer.FontSize = SizeLetra
        End If
       .MoveNext
     Loop
End With
Imprimir_Linea_H PosLinea, Ancho(0), Ancho(CantCampos), Negro, True
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

Public Sub Imprimir_Libro_Retenciones(Datas As Adodc)
Dim SizeLetra As Integer
On Error GoTo Errorhandler
Mensajes = "Seguro de Imprimir en:" & vbCrLf & Printer.DeviceName & "?"
Titulo = "IMPRESION"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
SizeLetra = 8
RatonReloj
InicioX = 0.5: InicioY = 0
'Escala_Centimetro   FormaImp, TipoTimes, SizeLetra
DataAnchoCampos InicioX, Datas, SizeLetra, TipoTimes, 1, True
Ancho(0) = 0.5   'Cuenta
Ancho(1) = 0.5   'T
Ancho(2) = 0.8   'Fecha
Ancho(3) = 2.3   'Autorizacion
Ancho(4) = 3.8   'Secuencial
Ancho(5) = 4.9   'Cliente
Ancho(6) = 10.5  'RUC_CI
Ancho(7) = 12.6  'TP
Ancho(8) = 13.2  'Numero
Ancho(9) = 14.3  'Valor_Fact
Ancho(10) = 15.8 'I.V.A.
Ancho(11) = 17.2 'Valor_Ret
Ancho(12) = 18.5 'Porc
Ancho(13) = 19.5
CantCampos = 13
Pagina = 1
'Iniciamos la impresion
Printer.FontBold = False
With Datas.Recordset
 If .RecordCount > 0 Then
     .MoveFirst
      Encabezado Ancho(0), Ancho(13)
      Printer.FontName = TipoCondensed
      PosLinea = PosLinea + 0.1
      Cuenta = .fields("Cuenta")
      Total_Ret = .fields("Porc")
      Cadena = Cuenta & ": " & Format$(Total_Ret, "00%")
      Total = 0
      Printer.FontBold = False
      Printer.FontSize = 12
      PrinterTexto Ancho(0), PosLinea, Cadena
      PosLinea = PosLinea + 0.6
      Printer.FontSize = 8
      Imprimir_Linea_H PosLinea, Ancho(0), Ancho(CantCampos)
      PosLinea = PosLinea + 0.1
      PrinterTexto Ancho(1), PosLinea, "T"
      PrinterTexto Ancho(2), PosLinea, "Fecha"
      PrinterTexto Ancho(3), PosLinea, "Autorizacion"
      PrinterTexto Ancho(4), PosLinea, "Reten."
      PrinterTexto Ancho(5), PosLinea, "Beneficiario"
      PrinterTexto Ancho(6), PosLinea, "RUC_CI"
      PrinterTexto Ancho(7), PosLinea, "TP"
      PrinterTexto Ancho(8), PosLinea, "Numero"
      PrinterTexto Ancho(9), PosLinea, "Valor_Fa"
      PrinterTexto Ancho(10), PosLinea, "I.V.A."
      PrinterTexto Ancho(11), PosLinea, "Valor_Re"
      PrinterTexto Ancho(12), PosLinea, "Porc"
      PosLinea = PosLinea + 0.5
      Imprimir_Linea_H PosLinea, Ancho(0), Ancho(CantCampos)
      PosLinea = PosLinea + 0.1
      Printer.FontBold = False
      Do While Not .EOF
         If Cuenta <> .fields("Cuenta") Or Total_Ret <> .fields("Porc") Then
            Printer.FontBold = False
            Imprimir_Linea_H PosLinea, Ancho(0), Ancho(CantCampos)
            PosLinea = PosLinea + 0.1
            PrinterTexto Ancho(9), PosLinea, "T O T A L"
            PrinterVariables Ancho(10), PosLinea, Total
            Cuenta = .fields("Cuenta")
            Total_Ret = .fields("Porc")
            PosLinea = PosLinea + 0.5
            Cadena = Cuenta & ": " & Format$(Total_Ret, "00%")
            Printer.FontBold = False
            Printer.FontSize = 12
            PrinterTexto Ancho(0), PosLinea, Cadena
            PosLinea = PosLinea + 0.6
            Printer.FontSize = 8
            Imprimir_Linea_H PosLinea, Ancho(0), Ancho(CantCampos)
            PosLinea = PosLinea + 0.1
            PrinterTexto Ancho(1), PosLinea, "T"
            PrinterTexto Ancho(2), PosLinea, "Fecha"
            PrinterTexto Ancho(3), PosLinea, "Autorizacion"
            PrinterTexto Ancho(4), PosLinea, "Reten."
            PrinterTexto Ancho(5), PosLinea, "Beneficiario"
            PrinterTexto Ancho(6), PosLinea, "RUC_CI"
            PrinterTexto Ancho(7), PosLinea, "TP"
            PrinterTexto Ancho(8), PosLinea, "Numero"
            PrinterTexto Ancho(9), PosLinea, "Valor_Fa"
            PrinterTexto Ancho(10), PosLinea, "I.V.A."
            PrinterTexto Ancho(11), PosLinea, "Valor_Re"
            PrinterTexto Ancho(12), PosLinea, "Porc"
            PosLinea = PosLinea + 0.5
            Imprimir_Linea_H PosLinea, Ancho(0), Ancho(CantCampos)
            PosLinea = PosLinea + 0.1
            Printer.FontBold = False
            Total = 0
         End If
         Printer.FontBold = False
         Printer.FontSize = SizeLetra
         PrinterFields Ancho(1), PosLinea, .fields("T")
         PrinterFields Ancho(2), PosLinea, .fields("Fecha")
         PrinterFields Ancho(3), PosLinea, .fields("Autorizacion")
         PrinterFields Ancho(4), PosLinea, .fields("Secuencial")
         PrinterFields Ancho(5), PosLinea, .fields("Cliente")
         PrinterFields Ancho(6), PosLinea, .fields("CI_RUC")
         PrinterFields Ancho(7), PosLinea, .fields("TP")
         PrinterFields Ancho(8), PosLinea, .fields("Numero")
         PrinterFields Ancho(9), PosLinea, .fields("Valor_Fact")
         PrinterFields Ancho(10), PosLinea, .fields("IVA")
         PrinterFields Ancho(11), PosLinea, .fields("Valor_Ret")
         PrinterFields Ancho(12), PosLinea, .fields("Porc")
         PosLinea = PosLinea + 0.35
         If PosLinea >= LimiteAlto Then
            Imprimir_Linea_H PosLinea, Ancho(0), Ancho(CantCampos)
            Printer.NewPage
            Encabezado Ancho(0), Ancho(11)
            Printer.FontName = TipoCondensed
            PosLinea = PosLinea + 0.1
            Printer.FontSize = 8
            Imprimir_Linea_H PosLinea, Ancho(0), Ancho(CantCampos)
            PosLinea = PosLinea + 0.1
            PrinterTexto Ancho(1), PosLinea, "T"
            PrinterTexto Ancho(2), PosLinea, "Fecha"
            PrinterTexto Ancho(3), PosLinea, "Autorizacion"
            PrinterTexto Ancho(4), PosLinea, "Reten."
            PrinterTexto Ancho(5), PosLinea, "Beneficiario"
            PrinterTexto Ancho(6), PosLinea, "RUC_CI"
            PrinterTexto Ancho(7), PosLinea, "TP"
            PrinterTexto Ancho(8), PosLinea, "Numero"
            PrinterTexto Ancho(9), PosLinea, "Valor_Fa"
            PrinterTexto Ancho(10), PosLinea, "I.V.A."
            PrinterTexto Ancho(10), PosLinea, "Valor_Re"
            PrinterTexto Ancho(11), PosLinea, "Porc"
            PosLinea = PosLinea + 0.5
            Imprimir_Linea_H PosLinea, Ancho(0), Ancho(CantCampos)
            PosLinea = PosLinea + 0.1
            Printer.FontBold = False
            Printer.FontSize = SizeLetra
         End If
         Total = Total + .fields("Valor_Ret")
        .MoveNext
      Loop
  End If
End With
Printer.FontBold = False
Imprimir_Linea_H PosLinea, Ancho(0), Ancho(CantCampos)
PosLinea = PosLinea + 0.1
PrinterTexto Ancho(9), PosLinea, "T O T A L"
PrinterVariables Ancho(10), PosLinea, Total
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

Public Sub ImprimirLibroReciboCaja(Datas As Adodc, _
                                   SizeLetra As Integer, _
                                   Opc As Integer, _
                                   Optional SubTot As Boolean)
On Error GoTo Errorhandler
Mensajes = "Imprimir Flujo de Caja"
Titulo = "IMPRESION"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then

RatonReloj
InicioX = 0.5: InicioY = 0
'Escala_Centimetro   FormaImp, TipoTimes, SizeLetra
DataAnchoCampos InicioX, Datas, SizeLetra, TipoTimes, 1
'Fecha,Beneficiario,Detalle,Concepto,Numero,Ingreso,Egreso,Saldo
Ancho(0) = 0.5 'Fecha
Ancho(1) = 0.7 'Beneficiario
Ancho(2) = 2.2 'Detalle
Ancho(3) = 3.4 'Concepto
Ancho(4) = 7   'Numero
Select Case Opc
  Case 1, 2: Ancho(5) = 17.5  'Ingreso / Egreso
             Ancho(6) = 19.5  'Fin
             CantCampos = 6
  Case 3: Ancho(5) = 13.5  'Ingreso
          Ancho(6) = 15.5  'Egreso
          Ancho(7) = 17.5  'Saldo
          Ancho(8) = 19.5  'Fin
          CantCampos = 8
End Select
Pagina = 1
'Iniciamos la impresion
Printer.FontBold = False
With Datas.Recordset
     .MoveFirst
      If Opc = 3 Then
         Encabezado Ancho(0), Ancho(8)
      Else
         Encabezado Ancho(0), Ancho(6)
      End If
      TituloLibroCaja
      Mifecha = Ninguno
      'Cuenta = .Fields("Codigo")
      'Cadena = Cuenta & " - " & .Fields("Detalle") & ": "
''      Cadena = .Fields("Detalle") & ": "
''      Printer.FontBold = True
''      Printer.FontSize = 12
''      Printer.FontUnderline = True
''      PrinterTexto Ancho(0), PosLinea, Cadena
''      Printer.FontUnderline = False
''      PosLinea = PosLinea + 0.5
      Printer.FontSize = SizeLetra
      EncabezadoLibroReciboCaja Opc
      Printer.FontBold = False
      If Opc = 3 Then
         Printer.FontBold = True
         Saldo = Redondear(.fields("Saldo") + .fields("Egreso") - .fields("Ingreso"), 2)
         PrinterTexto Ancho(4), PosLinea, "S A L D O    A N T E R I O R"
         PrinterVariables Ancho(7) - 0.5, PosLinea, Saldo
         PosLinea = PosLinea + 0.4
         Printer.FontBold = False
      End If
      Debe = 0: Haber = 0: Total = 0: Saldo = 0
      Codigo = SinEspaciosIzq(.fields("Concepto"))
      Do While Not .EOF
         Printer.FontSize = SizeLetra
         Printer.FontBold = False
         If Mifecha <> .fields("Fecha") Then
            PrinterFields Ancho(1), PosLinea, .fields("Fecha")
            Mifecha = .fields("Fecha")
         End If
         PrinterTexto Ancho(2), PosLinea, Format$(.fields("Numero"), "000000")
         PrinterFields Ancho(3), PosLinea, .fields("Beneficiario")
         PrinterFields Ancho(4), PosLinea, .fields("Concepto")
         Select Case Opc
           Case 1
                PrinterFields Ancho(5), PosLinea, .fields("Ingreso")
                Total = Total + .fields("Ingreso")
                Saldo = Saldo + .fields("Ingreso")
           Case 2
                PrinterFields Ancho(5), PosLinea, .fields("Egreso")
                Total = Total + .fields("Egreso")
                Saldo = Saldo + .fields("Egreso")
           Case 3
                PrinterFields Ancho(7), PosLinea, .fields("Saldo")
                PrinterFields Ancho(6), PosLinea, .fields("Egreso")
                PrinterFields Ancho(5), PosLinea, .fields("Ingreso")
                Debe = Debe + .fields("Ingreso")
                Haber = Haber + .fields("Egreso")
                Saldo = Saldo + .fields("Saldo")
                Total = .fields("Saldo")
         End Select
         PosLinea = PosLinea + 0.35
         If PosLinea >= LimiteAlto Then
            PosLinea = PosLinea + 0.05
            Imprimir_Linea_H PosLinea, Ancho(0), Ancho(CantCampos)
            Printer.NewPage
            If Opc = 3 Then
               Encabezado Ancho(0), Ancho(8)
            Else
               Encabezado Ancho(0), Ancho(6)
            End If
            TituloLibroCaja
            Printer.FontBold = True
            Printer.FontSize = 12
            PrinterTexto Ancho(0), PosLinea, Cadena
            PosLinea = PosLinea + 0.6
            Printer.FontSize = SizeLetra
            EncabezadoLibroReciboCaja Opc
            Printer.FontBold = False
            Mifecha = Ninguno
         End If
        .MoveNext
      Loop
End With
Printer.FontBold = True
PosLinea = PosLinea + 0.05
Imprimir_Linea_H PosLinea, Ancho(0), Ancho(CantCampos)
PosLinea = PosLinea + 0.1
If Opc = 3 Then
   PrinterTexto Ancho(4), PosLinea, "T O T A L"
   PrinterVariables Ancho(7), PosLinea, CDbl(Total)
   PrinterVariables Ancho(6), PosLinea, CDbl(Haber)
   PrinterVariables Ancho(5), PosLinea, CDbl(Debe)
Else
   PrinterTexto Ancho(4) + 7, PosLinea, "T O T A L"
   PrinterVariables Ancho(5), PosLinea, CDbl(Total)
End If
PosLinea = PosLinea + 0.5
Imprimir_Linea_H PosLinea, Ancho(0), Ancho(CantCampos), Negro, True
PosLinea = PosLinea + 0.1
If Opc <> 3 Then
   PrinterTexto Ancho(4) + 6, PosLinea, "G R A N   T O T A L"
   PrinterVariables Ancho(5) - 0.5, PosLinea, Saldo
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

'''Public Sub Procesar_Balance_Comprobacion(FormPadre As Form, _
'''                                         FechaInicial As String, _
'''                                         FechaFinal As String, _
'''                                         DataCtas As Adodc, _
'''                                         DataTrans As Adodc, _
'''                                         Optional EsBalanceMes As Boolean)
'''Dim CantCtas As Long
'''Dim TotalCtas As Long
'''  RatonReloj
'''  FechaIniN = CFechaLong(FechaInicial)
'''  FechaFinN = CFechaLong(FechaFinal)
''' 'FechaIni = "01/01/" & Format$(FechaAnio(FechaInicial), "0000")
'''  FechaIni = BuscarFecha(FechaInicial)
'''  FechaFin = BuscarFecha(FechaFinal)
'''  MiTiempo1 = Time
'''
'''  sSQL = "SELECT * " _
'''       & "FROM Catalogo_Cuentas " _
'''       & "WHERE DG <> 'G' " _
'''       & "AND Item = '" & NumEmpresa & "' " _
'''       & "AND Periodo = '" & Periodo_Contable & "' " _
'''       & "ORDER BY Codigo "
'''  Select_Adodc DataCtas, sSQL
'''
'''  sSQL = "UPDATE Catalogo_Cuentas " _
'''       & "SET Saldo_Anterior=0,Debitos=0,Creditos=0,Saldo_Total=0,Saldo_Total_ME=0," _
'''       & "Total_N6=0,Total_N5=0,Total_N4=0,Total_N3=0,Total_N2=0,Total_N1=0 " _
'''       & "WHERE Item = '" & NumEmpresa & "' " _
'''       & "AND Periodo = '" & Periodo_Contable & "' "
'''  Ejecutar_SQL_SP sSQL
'''
'''  sSQL = "UPDATE Catalogo_Cuentas " _
'''       & "SET Total_N6 = 0.0001 " _
'''       & "WHERE Item = '" & NumEmpresa & "' " _
'''       & "AND Periodo = '" & Periodo_Contable & "' " _
'''       & "AND LEN(Codigo)= 1 "
'''  Ejecutar_SQL_SP sSQL
'''
'''  sSQL = "UPDATE Fechas_Balance " _
'''       & "SET Fecha_Inicial = #" & FechaIni & "#, Fecha_Final = #" & FechaFin & "#, Cerrado = 0 "
'''  If EsBalanceMes Then
'''     sSQL = sSQL & "WHERE Detalle = 'Balance Mes' "
'''  Else
'''     sSQL = sSQL & "WHERE Detalle = 'Balance' "
'''  End If
'''  sSQL = sSQL _
'''       & "AND Item = '" & NumEmpresa & "' " _
'''       & "AND Periodo = '" & Periodo_Contable & "' "
'''  Ejecutar_SQL_SP sSQL
'''
'''  SumaDebe = 0: SumaHaber = 0
'''  Debe = 0: Haber = 0: Saldo = 0
'''  Debe_ME = 0: Haber_ME = 0: Saldo_ME = 0
'''  sSQL = "SELECT T.Cta,T.Fecha,T.TP,T.Numero," _
'''       & "Debe,Haber,Saldo,Parcial_ME,Saldo_ME,T.ID " _
'''       & "FROM Transacciones As T,Comprobantes As C "
'''  If EsBalanceMes Then
'''     sSQL = sSQL & "WHERE T.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# "
'''  Else
'''     sSQL = sSQL & "WHERE T.Fecha <= #" & FechaFin & "# "
'''  End If
'''  sSQL = sSQL _
'''       & "AND T.Periodo = '" & Periodo_Contable & "' " _
'''       & "AND T.T = '" & Normal & "' " _
'''       & "AND T.TP = C.TP " _
'''       & "AND T.Numero = C.Numero " _
'''       & "AND T.Fecha = C.Fecha " _
'''       & "AND T.Item = C.Item " _
'''       & "AND T.Periodo = C.Periodo "
'''  If ConSucursal = False Then sSQL = sSQL & "AND T.Item = '" & NumEmpresa & "' "
'''  sSQL = sSQL & "ORDER BY T.Cta,T.Fecha,T.TP,T.Numero,Debe DESC,Haber,T.ID "
'''  Select_Adodc DataTrans, sSQL
'''  RatonReloj
'''  With DataTrans.Recordset
'''  If .RecordCount > 0 Then
'''     .MoveFirst
'''      CantCtas = 0
'''      TotalCtas = .RecordCount
'''      Cuenta = .fields("Cta")
'''      Do While Not .EOF
'''         If Cuenta <> .fields("Cta") Then
'''            MiTiempo = Time - MiTiempo1
'''            Cadena = Format$(CantCtas * 100 / TotalCtas, "00") & "% " _
'''                   & " Tiempo: " & Format$(MiTiempo, "HH:MM:SS") _
'''                   & ". Cta No. " & Cuenta & "."
'''            FormPadre.Caption = Cadena
'''           'Actualizamos el total de la cuenta mayorizada
'''            CalculosTotalesCtas DataCtas, Cuenta, Debe, Haber, Saldo, Debe_ME, Haber_ME, Saldo_ME
'''            Cuenta = .fields("Cta")
'''            Debe = 0: Haber = 0: Saldo = 0
'''            Debe_ME = 0: Haber_ME = 0: Saldo_ME = 0
'''         End If
'''         FechaN = CFechaLong(.fields("Fecha"))
'''         If ((FechaIniN <= FechaN) And (FechaN <= FechaFinN)) Then
'''            Debe = Debe + Redondear(.fields("Debe"), 2)
'''            Haber = Haber + Redondear(.fields("Haber"), 2)
'''            If .fields("Parcial_ME") >= 0 Then
'''                Debe_ME = Debe_ME + Redondear(.fields("Parcial_ME"), 2)
'''            Else
'''                Haber_ME = Haber_ME + Redondear(-.fields("Parcial_ME"), 2)
'''            End If
'''            SumaDebe = SumaDebe + Redondear(.fields("Debe"), 2)
'''            SumaHaber = SumaHaber + Redondear(.fields("Haber"), 2)
'''         Else
'''            If Not EsBalanceMes Then
'''               Saldo = Redondear(.fields("Saldo"), 2)
'''               Saldo_ME = Redondear(.fields("Saldo_ME"), 2)
'''            End If
'''         End If
'''         CantCtas = CantCtas + 1
'''         DataTrans.Recordset.MoveNext
'''      Loop
'''  End If
'''  End With
'''  CalculosTotalesCtas DataCtas, Cuenta, Debe, Haber, Saldo, Debe_ME, Haber_ME, Saldo_ME
'''  TotalActivo = 0
'''  TotalPasivo = 0
'''  TotalCapital = 0
'''  TotalIngreso = 0
'''  TotalEgreso = 0
'''  TotalCostos = 0
'''  sSQL = "SELECT Codigo, Saldo_Total " _
'''       & "FROM Catalogo_Cuentas " _
'''       & "WHERE LEN(Codigo) = 1 " _
'''       & "AND Item = '" & NumEmpresa & "' " _
'''       & "AND Periodo = '" & Periodo_Contable & "' " _
'''       & "ORDER BY Codigo "
'''  Select_Adodc DataTrans, sSQL
'''  With DataTrans.Recordset
'''   If .RecordCount > 0 Then
'''       Do While Not .EOF
'''          Select Case .fields("Codigo")
'''            Case "1": TotalActivo = .fields("Saldo_Total")
'''            Case "2": TotalPasivo = .fields("Saldo_Total")
'''            Case "3": TotalCapital = .fields("Saldo_Total")
'''            Case "4": If OpcCoop Then TotalEgreso = .fields("Saldo_Total") Else TotalIngreso = .fields("Saldo_Total")
'''            Case "5": If OpcCoop Then TotalIngreso = .fields("Saldo_Total") Else TotalEgreso = .fields("Saldo_Total")
'''            Case "6": TotalCostos = .fields("Saldo_Total")
'''          End Select
'''         .MoveNext
'''       Loop
'''   End If
'''  End With
'''
''' 'Enceramos las cuentas Generales en los Movimientos
'''  sSQL = "UPDATE Catalogo_Cuentas " _
'''       & "SET Debitos = 0.00,Creditos = 0.00 " _
'''       & "WHERE DG = 'G' " _
'''       & "AND Item = '" & NumEmpresa & "' " _
'''       & "AND Periodo = '" & Periodo_Contable & "' "
'''  Ejecutar_SQL_SP sSQL
'''
'''  sSQL = "UPDATE Catalogo_Cuentas " _
'''       & "SET Total_N1 = " & Redondear(TotalPasivo + TotalCapital, 2) & " " _
'''       & "WHERE Item = '" & NumEmpresa & "' " _
'''       & "AND Periodo = '" & Periodo_Contable & "' " _
'''       & "AND TB = 'ES' " _
'''       & "AND Codigo = 'x' "
'''  Ejecutar_SQL_SP sSQL
'''
'''  sSQL = "UPDATE Catalogo_Cuentas " _
'''       & "SET Total_N1 = " & Redondear(TotalIngreso - TotalEgreso - TotalCostos, 2) & " " _
'''       & "WHERE Item = '" & NumEmpresa & "' " _
'''       & "AND Periodo = '" & Periodo_Contable & "' " _
'''       & "AND TB = 'ES' " _
'''       & "AND Codigo = 'xx' "
'''  Ejecutar_SQL_SP sSQL
'''
'''  sSQL = "UPDATE Catalogo_Cuentas " _
'''       & "SET Total_N1 = " & Redondear(TotalIngreso - TotalEgreso - TotalCostos, 2) & " " _
'''       & "WHERE Item = '" & NumEmpresa & "' " _
'''       & "AND Periodo = '" & Periodo_Contable & "' " _
'''       & "AND TB = 'ER' " _
'''       & "AND Codigo = 'x' "
'''  Ejecutar_SQL_SP sSQL
'''  RatonNormal
'''End Sub

Public Sub ProcesarBalanceMes(FormPadre As Form, _
                              FechaInicial As String, _
                              FechaFinal As String, _
                              DataCtas As Adodc, _
                              DataTrans As Adodc, _
                              DataSem As Adodc, _
                              BalanceFinal As Boolean)
Dim Semana(5) As Single
Dim TotalSemana(5) As Single
Dim Domingos(5) As Integer
Dim CantCtas As Long
Dim TotalCtas As Long
  RatonReloj
  Progreso_Iniciar
  Cuenta = Ninguno
  FechaIni = BuscarFecha(FechaInicial)
  FechaFin = BuscarFecha(FechaFinal)
  'FechaIni = "01/01/" & Format$(FechaAnio(FechaInicial), "0000")
  'FechaFin = BuscarFecha(FechaFinal)
  MiTiempo1 = Time
  sSQL = "UPDATE Catalogo_Cuentas " _
       & "SET Saldo_Anterior=0,Debitos=0,Creditos=0,Saldo_Total=0,Saldo_Total_ME=0 " _
       & "WHERE Codigo>='4' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' "
  Ejecutar_SQL_SP sSQL
  SumaDebe = 0: SumaHaber = 0
  Debe = 0: Haber = 0: Saldo = 0
  Debe_ME = 0: Haber_ME = 0: Saldo_ME = 0
  sSQL = "DELETE * " _
       & "FROM Balances_Mes " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' "
  'Ejecutar_SQL_SP sSQL
  sSQL = "SELECT * " _
       & "FROM Balances_Mes " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' "
  Select_Adodc DataSem, sSQL
  
  sSQL = "SELECT Cta,T.Fecha,T.TP,T.Numero,Debe,Haber,Saldo,Parcial_ME,Saldo_ME,T.ID " _
       & "FROM Transacciones As T,Comprobantes As C " _
       & "WHERE T.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
       & "AND T.TP = C.TP " _
       & "AND T.Numero = C.Numero " _
       & "AND T.Fecha = C.Fecha " _
       & "AND T.Item = C.Item " _
       & "AND T.Periodo = C.Periodo " _
       & "AND T.T = '" & Normal & "' " _
       & "AND Cta >='4' " _
       & "AND Cta < '6' "
  If ConSucursal = False Then sSQL = sSQL & "AND T.Item = '" & NumEmpresa & "' "
  '& "AND T.Item = '" & NumEmpresa & "' "
  sSQL = sSQL & "AND T.Periodo = '" & Periodo_Contable & "' " _
       & "ORDER BY T.Cta,T.Fecha,T.TP,T.Numero,Debe DESC,Haber,T.ID "
  Select_Adodc DataTrans, sSQL
  With DataTrans.Recordset
  If .RecordCount > 0 Then
      Progreso_Barra.Valor_Maximo = Progreso_Barra.Valor_Maximo + .RecordCount
      Progreso_Esperar
     .MoveFirst
      CantCtas = 0
      TotalCtas = .RecordCount
      Cuenta = .fields("Cta")
      Caracter = MidStrg(.fields("Cta"), 1, 1)
      For I = 0 To 4: Semana(I) = 0: Next I
      For I = 0 To 4: Domingos(I) = 0: Next I
      For I = 0 To 4: TotalSemana(I) = 0: Next I
      J = 4
      Mifecha = FechaFinal
      For I = Day(FechaFinal) To Day(FechaInicial) Step -1
         If Weekday(Mifecha) = 1 Then
            Domingos(J) = I
            J = J - 1
         End If
         Mifecha = CLongFecha(CFechaLong(Mifecha) - 1)
      Next I
      Do While Not .EOF
         Progreso_Barra.Mensaje_Box = "Cuenta: " & Cuenta
         Progreso_Esperar
         If Caracter <> MidStrg(.fields("Cta"), 1, 1) Then
            Cantidad = 0
            For I = 0 To 4
                Cantidad = Cantidad + TotalSemana(I)
            Next I
            Caracter = MidStrg(.fields("Cta"), 1, 1)
            For I = 0 To 4: TotalSemana(I) = 0: Next I
         End If
         If Cuenta <> .fields("Cta") Then
            MiTiempo = Time - MiTiempo1
            
            'FormPadre.Caption = Cadena
             
            CalculosTotalesCtasMes DataCtas, Cuenta, Debe, Haber, Saldo, Debe_ME, Haber_ME, Saldo_ME
             
            For I = 0 To 4
                TotalSemana(I) = TotalSemana(I) + Semana(I)
            Next I
            Cantidad = 0
            For I = 0 To 4
                Cantidad = Cantidad + Semana(I)
            Next I
            Cuenta = .fields("Cta")
            Debe = 0: Haber = 0: Saldo = 0
            Debe_ME = 0: Haber_ME = 0: Saldo_ME = 0
            For I = 0 To 4: Semana(I) = 0: Next I
         End If
         NoDias = Day(.fields("Fecha"))
         Select Case NoDias
           Case Domingos(0) To Domingos(1) - 1
                If MidStrg(.fields("Cta"), 1, 1) = "4" Then
                   Semana(0) = Semana(0) + .fields("Haber")
                Else
                   Semana(0) = Semana(0) + .fields("Debe")
                End If
           Case Domingos(1) To Domingos(2) - 1
                If MidStrg(.fields("Cta"), 1, 1) = "4" Then
                   Semana(1) = Semana(1) + .fields("Haber")
                Else
                   Semana(1) = Semana(1) + .fields("Debe")
                End If
           Case Domingos(2) To Domingos(3) - 1
                If MidStrg(.fields("Cta"), 1, 1) = "4" Then
                   Semana(2) = Semana(2) + .fields("Haber")
                Else
                   Semana(2) = Semana(2) + .fields("Debe")
                End If
           Case Domingos(3) To Domingos(4) - 1
                If MidStrg(.fields("Cta"), 1, 1) = "4" Then
                   Semana(3) = Semana(3) + .fields("Haber")
                Else
                   Semana(3) = Semana(3) + .fields("Debe")
                End If
           Case Domingos(4) To 31
                If MidStrg(.fields("Cta"), 1, 1) = "4" Then
                   Semana(4) = Semana(4) + .fields("Haber")
                Else
                   Semana(4) = Semana(4) + .fields("Debe")
                End If
         End Select
         Debe = Debe + Redondear(.fields("Debe"), 2)
         Haber = Haber + Redondear(.fields("Haber"), 2)
         If .fields("Parcial_ME") > 0 Then
             Debe_ME = Debe_ME + Redondear(.fields("Parcial_ME"), 2)
         Else
             Haber_ME = Haber_ME + Redondear(-.fields("Parcial_ME"), 2)
         End If
         SumaDebe = SumaDebe + Redondear(.fields("Debe"), 2)
         SumaHaber = SumaHaber + Redondear(.fields("Haber"), 2)
         CantCtas = CantCtas + 1
         Cadena = "BALANCE DE COMPROBACION MENSUAL: " & Format$(CantCtas * 100 / TotalCtas, "00") & "% " _
                   & " Tiempo: " & Format$(MiTiempo, "HH:MM:SS") _
                   & ". Cta No. " & Cuenta & "."
        .MoveNext
      Loop
  End If
  End With
  CalculosTotalesCtasMes DataCtas, Cuenta, Debe, Haber, Saldo, Debe_ME, Haber_ME, Saldo_ME
  For I = 0 To 4
      TotalSemana(I) = TotalSemana(I) + Semana(I)
  Next I
  Cantidad = 0
  For I = 0 To 4
      Cantidad = Cantidad + Semana(I)
  Next I
  Cantidad = 0
  For I = 0 To 4
      Cantidad = Cantidad + TotalSemana(I)
  Next I
  Cadena = "BALANCE DE COMPROBACION MENSUAL"
  FormPadre.Caption = Cadena
  RatonNormal
  Progreso_Final
  'MsgBox "FIN DEL PROCESO"
End Sub

'''Public Sub ProcesarBalance12Meses(FormPadre As Form, _
'''                                  FechaInicial As String, _
'''                                  FechaFinal As String, _
'''                                  BalanceTotal As Boolean, _
'''                                  NoEmpresa As String, _
'''                                  Tipo_Reporte As String)
'''Dim AdoCtas As ADODB.Recordset
'''Dim AdoTrans As ADODB.Recordset
'''Dim AdoTransSC As ADODB.Recordset
'''Dim CantCtas As Long
'''Dim TotalCtas As Long
'''
'''Dim VCTotales(13) As Currency
'''Dim Totales_1(13) As Currency
'''Dim Totales_2(13) As Currency
'''Dim Totales_3(13) As Currency
'''Dim Totales_4(13) As Currency
'''Dim Totales_5(13) As Currency
'''Dim Totales_6(13) As Currency
'''Dim Totales_7(13) As Currency
'''Dim Totales_8(13) As Currency
'''Dim Totales_9(13) As Currency
'''
'''  RatonReloj
'''  Progreso_Iniciar
'''  FechaIni = BuscarFecha(FechaInicial)
'''  FechaFin = BuscarFecha(FechaFinal)
'''  MiTiempo1 = Time
'''  NumItemTemp = NoEmpresa
'''  SumaDebe = 0: SumaHaber = 0
'''  Debe = 0: Haber = 0: Saldo = 0
'''  Debe_ME = 0: Haber_ME = 0: Saldo_ME = 0
'''  Contador = 0
'''  For NoMeses = 1 To 12
'''      VCTotales(NoMeses) = 0
'''      Totales_1(NoMeses) = 0
'''      Totales_2(NoMeses) = 0
'''      Totales_3(NoMeses) = 0
'''      Totales_4(NoMeses) = 0
'''      Totales_5(NoMeses) = 0
'''      Totales_6(NoMeses) = 0
'''      Totales_7(NoMeses) = 0
'''      Totales_8(NoMeses) = 0
'''      Totales_9(NoMeses) = 0
'''  Next NoMeses
'''
'''  sSQL = "SELECT * " _
'''       & "FROM Saldo_Diarios " _
'''       & "WHERE Item = '" & NoEmpresa & "' " _
'''       & "AND CodigoU = '" & CodigoUsuario & "' " _
'''       & "AND TB = '" & Tipo_Reporte & "' " _
'''       & "AND DG = 'D' " _
'''       & "AND CodigoC = '" & Ninguno & "' " _
'''       & "AND TP = 'E12M' " _
'''       & "ORDER BY Cta "
'''  Select_AdoDB AdoCtas, sSQL
'''  With AdoCtas
'''   If .RecordCount > 0 Then
'''       Do While Not .EOF
'''          Cta = .Fields("Cta_Aux")
'''          For NoMeses = 0 To 12
'''              VCTotales(NoMeses) = 0
'''          Next NoMeses
'''         'Saldo de Cuentas Contables
'''          sSQL = "SELECT T.Fecha,T.TP,T.Numero,T.Debe,T.Haber,T.Saldo,T.ID,CC.TC " _
'''               & "FROM Transacciones As T, Catalogo_Cuentas As CC " _
'''               & "WHERE T.TP IN ('CD','CE','CI','ND','NC') " _
'''               & "AND T.Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
'''               & "AND T.Periodo = '" & Periodo_Contable & "' "
'''          If ConSucursal = False Then sSQL = sSQL & "AND T.Item = '" & NoEmpresa & "' "
'''          sSQL = sSQL _
'''               & "AND T.Cta = '" & Cta & "' " _
'''               & "AND T.T <> '" & Anulado & "' " _
'''               & "AND T.Item = CC.Item " _
'''               & "AND T.Periodo = CC.Periodo " _
'''               & "AND T.Cta = CC.Codigo " _
'''               & "ORDER BY T.Fecha,T.TP,T.Numero,T.Debe DESC,T.Haber,T.ID "
'''          ' & "AND T.Procesado = " & Val(adFalse) & " "
'''          Select_AdoDB AdoTrans, sSQL
'''          If AdoTrans.RecordCount > 0 Then
'''             Do While Not AdoTrans.EOF
'''                Mifecha = AdoTrans.Fields("Fecha")
'''                NoMeses = FechaMes(Mifecha)
'''                Debe = AdoTrans.Fields("Debe")
'''                Haber = AdoTrans.Fields("Haber")
'''                Saldo = AdoTrans.Fields("Saldo")
'''                TipoCta = AdoTrans.Fields("TC")
'''                If OpcCoop Then
'''                   Select Case MidStrg(Cta, 1, 1)
'''                     Case "4": VCTotales(NoMeses) = VCTotales(NoMeses) + Debe - Haber
'''                     Case "5", "6", "7": VCTotales(NoMeses) = VCTotales(NoMeses) + Haber - Debe
'''                   End Select
'''                Else
'''                   Select Case MidStrg(Cta, 1, 1)
'''                     Case "4": VCTotales(NoMeses) = VCTotales(NoMeses) + Haber - Debe
'''                     Case "5", "6", "7": VCTotales(NoMeses) = VCTotales(NoMeses) + Debe - Haber
'''                   End Select
'''                End If
'''                Select Case MidStrg(Cta, 1, 1)
'''                  Case "1", "2", "3": VCTotales(NoMeses) = Saldo
'''                End Select
'''                AdoTrans.MoveNext
'''             Loop
'''          End If
'''          AdoTrans.Close
'''
'''          Insertar_SubCta_12_Meses Cta, Ninguno, VCTotales
'''
'''         'Mayorizamos los SubModulos
'''          Cta = .Fields("Cta_Aux")
'''          CodigoCli = Ninguno
'''          If BalanceTotal Then
'''          For NoMeses = 1 To 12
'''              VCTotales(NoMeses) = 0
'''          Next NoMeses
'''          Select Case MidStrg(Cta, 1, 1) ' TipoCta
'''            Case "1"
'''                 sSQL = "SELECT Codigo,Factura As Facturas,SUM(Debitos) As TDebitos,SUM(Creditos) As TCreditos,SUM(Debitos-Creditos) As TSaldo,"
'''                 SQL1 = "HAVING SUM(Debitos-Creditos) <> 0 "
'''            Case "2", "3"
'''                 sSQL = "SELECT Codigo,Factura As Facturas,SUM(Creditos) As TCreditos,SUM(Debitos) As TDebitos,SUM(Creditos-Debitos) As TSaldo,"
'''                 SQL1 = "HAVING SUM(Creditos-Debitos) <> 0 "
'''            Case "5", "6", "7"
'''                 sSQL = "SELECT Codigo,'0' As Facturas,SUM(Debitos) As TDebitos,SUM(Creditos) As TCreditos,SUM(Debitos-Creditos) As TSaldo,"
'''                 SQL1 = "HAVING SUM(Debitos-Creditos) <> 0 "
'''            Case "4"
'''                 sSQL = "SELECT Codigo,'0' As Facturas,SUM(Creditos) As TCreditos,SUM(Debitos) As TDebitos,SUM(Creditos-Debitos) As TSaldo,"
'''                 SQL1 = "HAVING SUM(Creditos-Debitos) <> 0 "
'''          End Select
'''          sSQL = sSQL & "TC,MONTH(Fecha) As Mes " _
'''               & "FROM Trans_SubCtas " _
'''               & "WHERE Periodo = '" & Periodo_Contable & "' "
'''          If ConSucursal = False Then sSQL = sSQL & "AND Item = '" & NoEmpresa & "' "
'''          sSQL = sSQL _
'''               & "AND TP IN ('CD','CE','CI','ND','NC') " _
'''               & "AND Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
'''               & "AND T <> '" & Anulado & "' " _
'''               & "AND Cta = '" & Cta & "' "
'''          Select Case TipoCta
'''            Case "C", "P"
'''                 sSQL = sSQL _
'''                      & "GROUP BY Codigo,Factura,TC,MONTH(Fecha) " _
'''                      & SQL1 _
'''                      & "ORDER BY Codigo,Factura,TC,MONTH(Fecha) "
'''            Case "I", "G", "CC"
'''                 sSQL = sSQL _
'''                      & "GROUP BY Codigo,TC,MONTH(Fecha) " _
'''                      & SQL1 _
'''                      & "ORDER BY Codigo,TC,MONTH(Fecha) "
'''          End Select
'''        ''          sSQL = "SELECT * " _
'''        ''               & "FROM Trans_SubCtas " _
'''        ''               & "WHERE TP IN ('CD','CE','CI','ND','NC') " _
'''        ''               & "AND Periodo = '" & Periodo_Contable & "' "
'''        ''          If ConSucursal = False Then sSQL = sSQL & "AND Item = '" & NoEmpresa & "' "
'''        ''          sSQL = sSQL _
'''        ''               & "AND Cta = '" & Cta & "' " _
'''        ''               & "AND Procesado = " & Val(adFalse) & " " _
'''        ''               & "AND T <> '" & Anulado & "' " _
'''        ''               & "ORDER BY Codigo,Fecha,TP,Numero,Factura,Debitos DESC,Creditos,ID "
'''          Select Case TipoCta
'''            Case "C", "P", "G", "I", "CC"
'''                 Select_AdoDB AdoTrans, sSQL
'''                 'MsgBox sSQL
'''                 If AdoTrans.RecordCount > 0 Then
'''                    CodigoCli = AdoTrans.Fields("Codigo")
'''                    Do While Not AdoTrans.EOF
'''                       If CodigoCli <> AdoTrans.Fields("Codigo") Then
'''                          Actualizar_SubCta_12_Meses Cta, CodigoCli, VCTotales
'''                          For NoMeses = 1 To 12
'''                              VCTotales(NoMeses) = 0
'''                          Next NoMeses
'''                          CodigoCli = AdoTrans.Fields("Codigo")
'''                          FormPadre.Caption = "Procesando Cuentas SubModulos: " & Format$(Contador / .RecordCount, "00%")
'''                       End If
''''                       Mifecha = AdoTrans.Fields("Fecha")
''''                       NoMeses = FechaMes(Mifecha)
'''                       NoMeses = AdoTrans.Fields("Mes")
'''                       Debe = AdoTrans.Fields("TDebitos")
'''                       Haber = AdoTrans.Fields("TCreditos")
'''                       Saldo = AdoTrans.Fields("TSaldo")
'''                       If OpcCoop Then
'''                          Select Case MidStrg(Cta, 1, 1)
'''                            Case "4": VCTotales(NoMeses) = VCTotales(NoMeses) + Debe - Haber
'''                            Case "5", "6", "7": VCTotales(NoMeses) = VCTotales(NoMeses) + Haber - Debe
'''                          End Select
'''                       Else
'''                          Select Case MidStrg(Cta, 1, 1)
'''                            Case "4": VCTotales(NoMeses) = VCTotales(NoMeses) + Haber - Debe
'''                            Case "5", "6", "7": VCTotales(NoMeses) = VCTotales(NoMeses) + Debe - Haber
'''                          End Select
'''                       End If
'''                       Select Case MidStrg(Cta, 1, 1)
'''                         Case "1", "2", "3": VCTotales(NoMeses) = Saldo
'''                       End Select
'''                       AdoTrans.MoveNext
'''                     Loop
'''                 End If
'''                 Actualizar_SubCta_12_Meses Cta, CodigoCli, VCTotales
'''                 AdoTrans.Close
'''          End Select
'''          End If
'''          Progreso_Barra.Mensaje_Box = "Procesando Cuentas: " & Format$(Contador / .RecordCount, "00%")
'''          Progreso_Esperar
'''          FormPadre.Caption = "Procesando Cuentas: " & Format$(Contador / .RecordCount, "00%")
'''          Contador = Contador + 1
'''         .MoveNext
'''       Loop
'''   End If
'''  End With
'''  sSQL = "SELECT * " _
'''       & "FROM Saldo_Diarios " _
'''       & "WHERE Item = '" & NoEmpresa & "' " _
'''       & "AND CodigoU = '" & CodigoUsuario & "' " _
'''       & "AND TB = '" & Tipo_Reporte & "' " _
'''       & "AND LEN(Cta) = 1 " _
'''       & "AND DG = 'G' " _
'''       & "AND TP = 'E12M' " _
'''       & "ORDER BY Cta "
'''  Select_AdoDB AdoCtas, sSQL
'''  With AdoCtas
'''   If .RecordCount > 0 Then
'''       Do While Not .EOF
'''          For NoMeses = 1 To 12
'''              Mes = MesesLetras(NoMeses)
'''              Select Case .Fields("Cta")
'''                Case "1": Totales_1(NoMeses) = .Fields(Mes)
'''                          Totales_1(0) = .Fields("Total")
'''                Case "2": Totales_2(NoMeses) = .Fields(Mes)
'''                          Totales_2(0) = .Fields("Total")
'''                Case "3": Totales_3(NoMeses) = .Fields(Mes)
'''                          Totales_3(0) = .Fields("Total")
'''                Case "4": Totales_4(NoMeses) = .Fields(Mes)
'''                          Totales_4(0) = .Fields("Total")
'''                Case "5": Totales_5(NoMeses) = .Fields(Mes)
'''                          Totales_5(0) = .Fields("Total")
'''                Case "6": Totales_6(NoMeses) = .Fields(Mes)
'''                          Totales_6(0) = .Fields("Total")
'''                Case "7": Totales_7(NoMeses) = .Fields(Mes)
'''                          Totales_7(0) = .Fields("Total")
'''                Case "8": Totales_8(NoMeses) = .Fields(Mes)
'''                          Totales_8(0) = .Fields("Total")
'''                Case "9": Totales_9(NoMeses) = .Fields(Mes)
'''                          Totales_9(0) = .Fields("Total")
'''              End Select
'''          Next NoMeses
'''         .MoveNext
'''       Loop
'''   End If
'''  End With
'''  AdoCtas.Close
'''
'''  sSQL = "UPDATE Saldo_Diarios " _
'''       & "SET "
'''  For NoMeses = 1 To 12
'''      Mes = MesesLetras(NoMeses)
'''      Select Case Tipo_Reporte
'''        Case "ER": sSQL = sSQL & Mes & " = " & Totales_4(NoMeses) - (Totales_5(NoMeses) + Totales_6(NoMeses)) & ", "
'''        Case "ES": sSQL = sSQL & Mes & " = " & Totales_1(NoMeses) - Totales_2(NoMeses) - Totales_3(NoMeses) & ", "
'''      End Select
'''  Next NoMeses
'''  Select Case Tipo_Reporte
'''    Case "ER": sSQL = sSQL & "Total = " & Totales_4(0) - (Totales_5(0) + Totales_6(0)) & " "
'''    Case "ES": sSQL = sSQL & "Total = " & Totales_1(0) - Totales_2(0) - Totales_3(0) & " "
'''  End Select
'''
'''  sSQL = sSQL _
'''       & "WHERE Item = '" & NoEmpresa & "' " _
'''       & "AND CodigoU = '" & CodigoUsuario & "' " _
'''       & "AND Cta = '(+/-)' " _
'''       & "AND TP = 'E12M' "
'''  Ejecutar_SQL_SP sSQL
'''
'''  sSQL = "DELETE * " _
'''       & "FROM Saldo_Diarios " _
'''       & "WHERE Item = '" & NoEmpresa & "' " _
'''       & "AND CodigoU = '" & CodigoUsuario & "' " _
'''       & "AND TB = '" & Tipo_Reporte & "' " _
'''       & "AND TP = 'E12M' " _
'''       & "AND ("
'''  For NoMeses = 1 To 12
'''      sSQL = sSQL & " ABS(" & MesesLetras(NoMeses) & ") + "
'''  Next NoMeses
'''  sSQL = sSQL _
'''       & "ABS(Total)) = 0 "
'''  Ejecutar_SQL_SP sSQL
'''
'''  sSQL = "UPDATE Saldo_Diarios " _
'''       & "SET Presupuesto = 0, Diferencia = 0 " _
'''       & "WHERE Cta = '(+/-)' " _
'''       & "AND Item = '" & NoEmpresa & "' " _
'''       & "AND CodigoU = '" & CodigoUsuario & "' " _
'''       & "AND TP = 'E12M' "
'''  Ejecutar_SQL_SP sSQL
'''
'''  Cadena = "ESTADO DE SITUACION ANALITICO ANUAL"
'''  FormPadre.Caption = Cadena
''''  AdoTrans.Close
''''  AdoTransSC.Close
'''  RatonNormal
'''  Progreso_Final
''' 'MsgBox NoEmpresa & vbCrLf & ConSucursal
'''End Sub

Public Sub IniciarAsientos(DBGAsiento As DataGrid, _
                           DBGSubCta As DataGrid, _
                           DBGBanco As DataGrid, _
                           DBGRets As DataGrid, _
                           DtaAsiento As Adodc, _
                           DtaBanco As Adodc, _
                           DtaSubCta As Adodc, _
                           DtaRet As Adodc)
  If Trans_No <= 0 Then Trans_No = 1
  If NuevoComp Then Eliminar_Asientos_SP True
  SQL2 = "SELECT * " _
       & "FROM Asiento " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND T_No = " & Trans_No & " "
  Select_Adodc_Grid DBGAsiento, DtaAsiento, SQL2
  SQL2 = "SELECT * " _
       & "FROM Asiento_SC " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND T_No = " & Trans_No & " "
  Select_Adodc_Grid DBGSubCta, DtaSubCta, SQL2
  SQL2 = "SELECT * " _
       & "FROM Asiento_B " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND T_No = " & Trans_No & " "
  Select_Adodc_Grid DBGBanco, DtaBanco, SQL2
  SQL2 = "SELECT * " _
       & "FROM Asiento_Air " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND T_No = " & Trans_No & " "
  Select_Adodc_Grid DBGRets, DtaRet, SQL2
  NuevoComp = False
End Sub

Public Sub IniciarAsientosDe(DBGAsiento As DataGrid, _
                             DtaAsiento As Adodc)
  If Trans_No <= 0 Then Trans_No = 1
  Ln_No = 1
  Eliminar_Asientos_SP True
'''  SQL2 = "SELECT * " _
'''       & "FROM Asiento_SC " _
'''       & "WHERE Item = '" & NumEmpresa & "' " _
'''       & "AND T_No = " & Trans_No & " " _
'''       & "AND CodigoU = '" & CodigoUsuario & "' "
'''  Select_Adodc DtaAsiento, SQL2
'''  SQL2 = "SELECT * " _
'''       & "FROM Asiento_B " _
'''       & "WHERE Item = '" & NumEmpresa & "' " _
'''       & "AND T_No = " & Trans_No & " " _
'''       & "AND CodigoU = '" & CodigoUsuario & "' "
'''  Select_Adodc DtaAsiento, SQL2
'''  SQL2 = "SELECT * " _
'''       & "FROM Asiento_R " _
'''       & "WHERE Item = '" & NumEmpresa & "' " _
'''       & "AND T_No = " & Trans_No & " " _
'''       & "AND CodigoU = '" & CodigoUsuario & "' "
'''  Select_Adodc DtaAsiento, SQL2
'''  SQL2 = "SELECT * " _
'''       & "FROM Asiento_K " _
'''       & "WHERE Item = '" & NumEmpresa & "' " _
'''       & "AND T_No = " & Trans_No & " " _
'''       & "AND CodigoU = '" & CodigoUsuario & "' "
'''  Select_Adodc DtaAsiento, SQL2
  SQL2 = "SELECT * " _
       & "FROM Asiento " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND T_No = " & Trans_No & " " _
       & "AND CodigoU = '" & CodigoUsuario & "' "
  Select_Adodc_Grid DBGAsiento, DtaAsiento, SQL2
End Sub

Public Sub IniciarAsientosAdo(DtaAsiento As Adodc)
  If Trans_No <= 0 Then Trans_No = 1
  Eliminar_Asientos_SP True
  SQL2 = "SELECT * " _
       & "FROM Asiento_SC " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND T_No = " & Trans_No & " "
  Select_Adodc DtaAsiento, SQL2
  SQL2 = "SELECT * " _
       & "FROM Asiento_B " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND T_No = " & Trans_No & " "
  Select_Adodc DtaAsiento, SQL2
  SQL2 = "SELECT * " _
       & "FROM Asiento_Air " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND T_No = " & Trans_No & " "
  Select_Adodc DtaAsiento, SQL2
  SQL2 = "SELECT * " _
       & "FROM Asiento_Compras " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND T_No = " & Trans_No & " "
  Select_Adodc DtaAsiento, SQL2
  SQL2 = "SELECT * " _
       & "FROM Asiento_Ventas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND T_No = " & Trans_No & " "
  Select_Adodc DtaAsiento, SQL2
  SQL2 = "SELECT * " _
       & "FROM Asiento_Importaciones " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND T_No = " & Trans_No & " "
  Select_Adodc DtaAsiento, SQL2
  SQL2 = "SELECT * " _
       & "FROM Asiento_Exportaciones " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND T_No = " & Trans_No & " "
  Select_Adodc DtaAsiento, SQL2
  SQL2 = "SELECT * " _
       & "FROM Asiento_K " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND T_No = " & Trans_No & " "
  Select_Adodc DtaAsiento, SQL2
  SQL2 = "SELECT * " _
       & "FROM Asiento " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND T_No = " & Trans_No & " "
  Select_Adodc DtaAsiento, SQL2
End Sub

'''Public Sub Eliminar_Asientos_SP(B_Asiento As Boolean)
'''  If Trans_No <= 0 Then Trans_No = 1
'''  If B_Asiento Then
'''     SQL1 = "DELETE " _
'''          & "FROM Asiento " _
'''          & "WHERE Item = '" & NumEmpresa & "' " _
'''          & "AND CodigoU = '" & CodigoUsuario & "' " _
'''          & "AND T_No = " & Trans_No & " "
'''     Ejecutar_SQL_SP SQL1
'''  End If
'''  SQL1 = "DELETE " _
'''       & "FROM Asiento_SC " _
'''       & "WHERE Item = '" & NumEmpresa & "' " _
'''       & "AND CodigoU = '" & CodigoUsuario & "' " _
'''       & "AND T_No = " & Trans_No & " "
'''  Ejecutar_SQL_SP SQL1
'''  SQL1 = "DELETE " _
'''       & "FROM Asiento_B " _
'''       & "WHERE Item = '" & NumEmpresa & "' " _
'''       & "AND CodigoU = '" & CodigoUsuario & "' " _
'''       & "AND T_No = " & Trans_No & " "
'''  Ejecutar_SQL_SP SQL1
'''  SQL1 = "DELETE " _
'''       & "FROM Asiento_R " _
'''       & "WHERE Item = '" & NumEmpresa & "' " _
'''       & "AND CodigoU = '" & CodigoUsuario & "' " _
'''       & "AND T_No = " & Trans_No & " "
'''  Ejecutar_SQL_SP SQL1
'''  SQL1 = "DELETE " _
'''       & "FROM Asiento_RP " _
'''       & "WHERE Item = '" & NumEmpresa & "' " _
'''       & "AND CodigoU = '" & CodigoUsuario & "' " _
'''       & "AND T_No = " & Trans_No & " "
'''  Ejecutar_SQL_SP SQL1
'''  SQL1 = "DELETE " _
'''       & "FROM Asiento_K " _
'''       & "WHERE Item = '" & NumEmpresa & "' " _
'''       & "AND CodigoU = '" & CodigoUsuario & "' " _
'''       & "AND T_No = " & Trans_No & " "
'''  Ejecutar_SQL_SP SQL1
'''  SQL1 = "DELETE " _
'''       & "FROM Asiento_P " _
'''       & "WHERE Item = '" & NumEmpresa & "' " _
'''       & "AND CodigoU = '" & CodigoUsuario & "' " _
'''       & "AND T_No = " & Trans_No & " "
'''  Ejecutar_SQL_SP SQL1
'''  SQL1 = "DELETE " _
'''       & "FROM Asiento_Air " _
'''       & "WHERE Item = '" & NumEmpresa & "' " _
'''       & "AND CodigoU = '" & CodigoUsuario & "' " _
'''       & "AND T_No = " & Trans_No & " "
'''  Ejecutar_SQL_SP SQL1
'''  SQL1 = "DELETE " _
'''       & "FROM Asiento_Compras " _
'''       & "WHERE Item = '" & NumEmpresa & "' " _
'''       & "AND CodigoU = '" & CodigoUsuario & "' " _
'''       & "AND T_No = " & Trans_No & " "
'''  Ejecutar_SQL_SP SQL1
'''  SQL1 = "DELETE " _
'''       & "FROM Asiento_Exportaciones " _
'''       & "WHERE Item = '" & NumEmpresa & "' " _
'''       & "AND CodigoU = '" & CodigoUsuario & "' " _
'''       & "AND T_No = " & Trans_No & " "
'''  Ejecutar_SQL_SP SQL1
'''  SQL1 = "DELETE " _
'''       & "FROM Asiento_Importaciones " _
'''       & "WHERE Item = '" & NumEmpresa & "' " _
'''       & "AND CodigoU = '" & CodigoUsuario & "' " _
'''       & "AND T_No = " & Trans_No & " "
'''  Ejecutar_SQL_SP SQL1
'''  SQL1 = "DELETE " _
'''       & "FROM Asiento_Ventas " _
'''       & "WHERE Item = '" & NumEmpresa & "' " _
'''       & "AND CodigoU = '" & CodigoUsuario & "' " _
'''       & "AND T_No = " & Trans_No & " "
'''  Ejecutar_SQL_SP SQL1
'''End Sub

Public Function ExistenMovimientos() As Boolean
Dim AdoCtas As ADODB.Recordset
Dim ExisteMov As Boolean
  ExisteMov = False
  If Trans_No <= 0 Then Trans_No = 1
  SQL1 = "SELECT Item, CodigoU, T_No " _
       & "FROM Asiento " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND T_No = " & Trans_No & " " _
       & "UNION " _
       & "SELECT Item, CodigoU, T_No " _
       & "FROM Asiento_SC " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND T_No = " & Trans_No & " " _
       & "UNION " _
       & "SELECT Item, CodigoU, T_No " _
       & "FROM Asiento_B " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND T_No = " & Trans_No & " " _
       & "UNION " _
       & "SELECT Item, CodigoU, T_No " _
       & "FROM Asiento_R " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND T_No = " & Trans_No & " "
       
  SQL2 = "SELECT Item, CodigoU, T_No " _
       & "FROM Asiento_RP " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND T_No = " & Trans_No & " " _
       & "UNION " _
       & "SELECT Item, CodigoU, T_No " _
       & "FROM Asiento_K " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND T_No = " & Trans_No & " " _
       & "UNION " _
       & "SELECT Item, CodigoU, T_No " _
       & "FROM Asiento_P " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND T_No = " & Trans_No & " " _
       & "UNION " _
       & "SELECT Item, CodigoU, T_No " _
       & "FROM Asiento_Air " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND T_No = " & Trans_No & " "
       
  SQL3 = "SELECT Item, CodigoU, T_No " _
       & "FROM Asiento_Compras " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND T_No = " & Trans_No & " " _
       & "UNION " _
       & "SELECT Item, CodigoU, T_No " _
       & "FROM Asiento_Exportaciones " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND T_No = " & Trans_No & " " _
       & "UNION " _
       & "SELECT Item, CodigoU, T_No " _
       & "FROM Asiento_Importaciones " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND T_No = " & Trans_No & " " _
       & "UNION " _
       & "SELECT Item, CodigoU, T_No " _
       & "FROM Asiento_Ventas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND T_No = " & Trans_No & " "
  
  sSQL = SQL1 & "UNION " & SQL2 & "UNION " & SQL3
  Select_AdoDB AdoCtas, sSQL
  If AdoCtas.RecordCount > 0 Then ExisteMov = True
  AdoCtas.Close
  ExistenMovimientos = ExisteMov
End Function

Public Sub Encab_Libro_Diario(Datas As Adodc)
With Datas.Recordset
  If Not .EOF Then
     If OpcCoop Then
        Printer.Line (Ancho(0), PosLinea)-(Ancho(5), PosLinea + 1), Negro, B
        Printer.Line (Ancho(0), PosLinea + 0.5)-(Ancho(5), PosLinea + 0.5), Negro
        For J = 1 To C - 1
            If J <> 2 Then Printer.Line (Ancho(J), PosLinea + 0.5)-(Ancho(J), PosLinea + 1), Negro
        Next J
     Else
        Printer.Line (Ancho(0), PosLinea)-(Ancho(5), PosLinea + 1.4), Negro, B
        For J = 1 To C - 1
            Printer.Line (Ancho(J), PosLinea + 0.9)-(Ancho(J), PosLinea + 1.4), Negro
        Next J
     End If
     Printer.FontBold = True
     PrinterTexto Ancho(0), PosLinea + 0.1, "FECHA:"
     If OpcCoop Then
        PrinterTexto Ancho(0), PosLinea + 0.6, "CODIGO"
        PrinterTexto Ancho(1), PosLinea + 0.6, "CUENTA"
        PrinterTexto Ancho(3), PosLinea + 0.6, "DEBE"
        PrinterTexto Ancho(4), PosLinea + 0.6, "HABER"
     Else
        PrinterTexto Ancho(0) + 3.5, PosLinea + 0.1, "ELABORADO POR:"
        PrinterTexto Ancho(2), PosLinea + 0.1, "TP:"
        PrinterTexto Ancho(3), PosLinea + 0.1, "NUMERO No."
        PrinterTexto Ancho(0), PosLinea + 0.5, "CONCEPTO:"
        PrinterTexto Ancho(0), PosLinea + 1, "CODIGO"
        PrinterTexto Ancho(1), PosLinea + 1, "CUENTA"
        PrinterTexto Ancho(2), PosLinea + 1, "Parcial_ME"
        PrinterTexto Ancho(3), PosLinea + 1, "DEBE"
        PrinterTexto Ancho(4), PosLinea + 1, "HABER"
     End If
    'Encabezados del libro diario
     Printer.FontBold = False
     PrinterFields Ancho(0) + 1.1, PosLinea + 0.1, .fields("Fecha")
     If OpcCoop Then
        PosLinea = PosLinea + 1.1
     Else
        NumeroLineas = PrinterLineasMayor(Ancho(0) + 1.8, PosLinea + 0.5, .fields("Concepto"), 15)
        'MsgBox PosLinea & vbCrLf & NumeroLineas & vbCrLf & PosLinea_Aux
        'PrinterFields Ancho(0) + 1.8, PosLinea + 0.5, .Fields("Concepto")
        'PosLinea = PosLinea + 0.8
        PrinterFields Ancho(0) + 6.2, PosLinea + 0.1, .fields("CodigoU")
        PrinterFields Ancho(2) + 0.6, PosLinea + 0.1, .fields("TP")
        PrinterTexto Ancho(3) + 1.8, PosLinea + 0.1, Format$(.fields("Numero"), "00000000")
        Printer.Line (Ancho(5), PosLinea + 0.4)-(Ancho(5), PosLinea + 1.4), Negro
        Printer.Line (Ancho(0), PosLinea + 0.9)-(Ancho(5), PosLinea + 0.9), Negro
        PosLinea = PosLinea + 1.5
     End If
  End If
End With
End Sub

Public Sub EliminarSubCta(CodigoSC As String, OpcDH As Byte)
  sSQL = "DELETE * " _
       & "FROM Asiento_SC " _
       & "WHERE Cta = '" & CodigoSC & "' " _
       & "AND T_No = " & Trans_No & " " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' "
  If OpcDH = 1 Then sSQL = sSQL & "AND DH = '1' "
  If OpcDH = 2 Then sSQL = sSQL & "AND DH = '2' "
  Ejecutar_SQL_SP sSQL
  sSQL = "DELETE * " _
       & "FROM Asiento_Compras " _
       & "WHERE Cta_Bienes = '" & CodigoSC & "' " _
       & "AND T_No = " & Trans_No & " " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' "
  Ejecutar_SQL_SP sSQL
  sSQL = "DELETE * " _
       & "FROM Asiento_Compras " _
       & "WHERE Cta_Servicio = '" & CodigoSC & "' " _
       & "AND T_No = " & Trans_No & " " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' "
  Ejecutar_SQL_SP sSQL
  sSQL = "DELETE * " _
       & "FROM Asiento_Air " _
       & "WHERE Cta_Retencion = '" & CodigoSC & "' " _
       & "AND T_No = " & Trans_No & " " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' "
  Ejecutar_SQL_SP sSQL
  sSQL = "DELETE * " _
       & "FROM Asiento_RP " _
       & "WHERE Cta = '" & CodigoSC & "' " _
       & "AND T_No = " & Trans_No & " " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' "
  Ejecutar_SQL_SP sSQL
End Sub

Public Sub CalculosTotalesCtas(DataCta As Adodc, _
                               CodigoCta As String, _
                               TDebe As Currency, _
                               THaber As Currency, _
                               TSaldo As Currency, _
                               TDebeME As Currency, _
                               THaberME As Currency, _
                               TSaldoME As Currency)
Dim SaldoTotal As Currency
Dim SaldoTotalME As Currency
Dim Nivel_Cta As Byte
Dim CodigoAux As String

   'If MidStrg(CodigoCta, 1, 4) = "2.01" Then MsgBox Abs(TSaldo)
   Cta_Sup = CodigoCuentaSup(CodigoCta)
   If OpcCoop Then
      Select Case MidStrg(CodigoCta, 1, 1)
        Case "1", "4", "6", "8"
             SaldoTotal = Redondear(TSaldo + TDebe - THaber, 2)
             SaldoTotalME = Redondear(TSaldoME + TDebeME - THaberME, 2)
        Case "2", "3", "5", "7", "9"
             SaldoTotal = Redondear(TSaldo + THaber - TDebe, 2)
             SaldoTotalME = Redondear(TSaldoME + THaberME - TDebeME, 2)
      End Select
   Else
      Select Case MidStrg(CodigoCta, 1, 1)
        Case "1", "5", "6", "8"
             SaldoTotal = Redondear(TSaldo + TDebe - THaber, 2)
             SaldoTotalME = Redondear(TSaldoME + TDebeME - THaberME, 2)
        Case "2", "3", "4", "7", "9"
             SaldoTotal = Redondear(TSaldo + THaber - TDebe, 2)
             SaldoTotalME = Redondear(TSaldoME + THaberME - TDebeME, 2)
      End Select
   End If
   CodigoAux = CodigoCta
   Nivel_Cta = Niveles(CodigoAux)
  'MsgBox CodigoAux & vbCrLf & Nivel_Cta
   sSQL = "UPDATE Catalogo_Cuentas " _
        & "SET Total_N" & Nivel_Cta & " = " & SaldoTotal & "," _
        & "Saldo_Anterior = " & TSaldo & "," _
        & "Debitos = " & TDebe & "," _
        & "Creditos = " & THaber & "," _
        & "Saldo_Total = " & SaldoTotal & "," _
        & "Saldo_Total_ME = " & SaldoTotalME & " " _
        & "WHERE Codigo = '" & CodigoCta & "' " _
        & "AND Item = '" & NumEmpresa & "' " _
        & "AND Periodo = '" & Periodo_Contable & "' "
   Ejecutar_SQL_SP sSQL
  'Actualizamos las cuentas generales
   Do While (Cta_Sup <> "0")
      CodigoAux = CodigoCuentaSup(CodigoAux)
      Nivel_Cta = Niveles(CodigoAux)
     'MsgBox Nivel_Cta
      sSQL = "UPDATE Catalogo_Cuentas " _
           & "SET Saldo_Total = Saldo_Total + " & SaldoTotal & "," _
           & "Saldo_Total_ME = Saldo_Total_ME + " & SaldoTotalME & "," _
           & "Total_N" & Nivel_Cta & " = Total_N" & Nivel_Cta & " + " & SaldoTotal & " " _
           & "WHERE Codigo = '" & Cta_Sup & "' " _
           & "AND Item = '" & NumEmpresa & "' " _
           & "AND Periodo = '" & Periodo_Contable & "' "
      Ejecutar_SQL_SP sSQL
      Cta_Sup = CodigoCuentaSup(Cta_Sup)
   Loop
End Sub

Public Sub CalculosTotalesInventario(DataCta As Adodc, _
                                     CodigoCta As String, _
                                     TEntradas As Currency, _
                                     TSalidas As Currency, _
                                     TSaldoAnt As Single, _
                                     TSaldoAct As Single, _
                                     TSaldo As Currency)
   sSQL = "SELECT * " _
        & "FROM Catalogo_Productos " _
        & "WHERE TC <> 'I' " _
        & "AND Codigo_Inv = '" & CodigoCta & "' " _
        & "AND Item = '" & NumEmpresa & "' " _
        & "AND Periodo = '" & Periodo_Contable & "' "
   Select_Adodc DataCta, sSQL
   With DataCta.Recordset
    If .RecordCount > 0 Then
        Cta_Sup = CodigoCuentaSup(CodigoCta)
       .fields("Stock_Anterior") = TSaldoAnt
       .fields("Entradas") = TEntradas
       .fields("Salidas") = TSalidas
       .fields("Stock_Actual") = TSaldoAct
        .fields("Promedio") = Valor_Prom
        'If Contador > 0 Then .Fields("Promedio") = Redondear(Precio / Contador, 4)
       .fields("Valor_Total") = TSaldo
       .Update
        'If .Fields("Codigo_Inv") = "05.01.02" Then MsgBox Valor_Prom & vbCrLf & TSaldoAct
    End If
   End With
   Do While (Cta_Sup <> "0")
     'Actualizamos las cuentas generales
      sSQL = "UPDATE Catalogo_Productos " _
           & "SET Stock_Anterior = Stock_Anterior + " & TSaldoAnt & "," _
           & "Stock_Actual = Stock_Actual + " & TSaldoAct & "," _
           & "Entradas = Entradas + " & TEntradas & "," _
           & "Salidas = Salidas + " & TSalidas & "," _
           & "Valor_Total = Valor_Total + " & TSaldo & " " _
           & "WHERE Codigo_Inv = '" & Cta_Sup & "' " _
           & "AND Item = '" & NumEmpresa & "' " _
           & "AND Periodo = '" & Periodo_Contable & "' "
      Ejecutar_SQL_SP sSQL
      Cta_Sup = CodigoCuentaSup(Cta_Sup)
   Loop
End Sub

Public Sub CalculosTotalesCtasMes(DataCta As Adodc, _
                                  CodigoCta As String, _
                                  TDebe As Currency, _
                                  THaber As Currency, _
                                  TSaldo As Currency, _
                                  TDebeME As Currency, _
                                  THaberME As Currency, _
                                  TSaldoME As Currency)
Dim SaldoTotal As Currency
Dim SaldoTotalME As Currency
   TSaldo = 0
   TSaldoME = 0
   sSQL = "SELECT * " _
        & "FROM Catalogo_Cuentas " _
        & "WHERE DG <> 'G' " _
        & "AND Codigo ='" & CodigoCta & "' " _
        & "AND Item = '" & NumEmpresa & "' " _
        & "AND Periodo = '" & Periodo_Contable & "' "
   Select_Adodc DataCta, sSQL
   With DataCta.Recordset
    If .RecordCount > 0 Then
        Cta_Sup = CodigoCuentaSup(CodigoCta)
        If OpcCoop Then
           Select Case MidStrg(.fields("Codigo"), 1, 1)
             Case "1", "4", "6", "8"
                  SaldoTotal = Redondear(TSaldo + TDebe - THaber, 2)
                  SaldoTotalME = Redondear(TSaldoME + TDebeME - THaberME, 2)
             Case "2", "3", "5", "7", "9"
                  SaldoTotal = Redondear(TSaldo + THaber - TDebe, 2)
                  SaldoTotalME = Redondear(TSaldoME + THaberME - TDebeME, 2)
           End Select
        Else
           Select Case MidStrg(.fields("Codigo"), 1, 1)
             Case "1", "5"
                  SaldoTotal = Redondear(TSaldo + TDebe - THaber, 2)
                  SaldoTotalME = Redondear(TSaldoME + TDebeME - THaberME, 2)
             Case "2", "3", "4"
                  SaldoTotal = Redondear(TSaldo + THaber - TDebe, 2)
                  SaldoTotalME = Redondear(TSaldoME + THaberME - TDebeME, 2)
           End Select
        End If
        'If .Fields("ME") = False Then SaldoTotalME = 0
       .fields("Saldo_Anterior") = TSaldo
       .fields("Debitos") = TDebe
       .fields("Creditos") = THaber
       .fields("Saldo_Total") = SaldoTotal
       .fields("Saldo_Total_ME") = SaldoTotalME
       .Update
    End If
   End With
   Do While (Cta_Sup <> "0")
     'Actualizamos las cuentas generales
      sSQL = "UPDATE Catalogo_Cuentas " _
           & "SET Saldo_Total = Saldo_Total + " & SaldoTotal & "," _
           & "Saldo_Total_ME = Saldo_Total_ME + " & SaldoTotalME & " " _
           & "WHERE Codigo = '" & Cta_Sup & "' " _
           & "AND Item = '" & NumEmpresa & "' " _
           & "AND Periodo = '" & Periodo_Contable & "' "
      Ejecutar_SQL_SP sSQL
      Cta_Sup = CodigoCuentaSup(Cta_Sup)
   Loop
End Sub

Public Function CalculosSaldoAnt(TipoCod As String, _
                                 TDebe As Currency, _
                                 THaber As Currency, _
                                 TSaldo As Currency) As Currency
Dim TotSaldoAnt As Currency
  If OpcCoop Then
     Select Case MidStrg(TipoCod, 1, 1)
       Case "1", "4", "6", "8"
            TotSaldoAnt = Redondear(TSaldo - TDebe + THaber, 2)
       Case "2", "3", "5", "7", "9"
            TotSaldoAnt = Redondear(TSaldo - THaber + TDebe, 2)
     End Select
  Else
     Select Case MidStrg(TipoCod, 1, 1)
       Case "1", "5", "7", "9"
            TotSaldoAnt = Redondear(TSaldo - TDebe + THaber, 2)
       Case "2", "3", "4", "6", "8"
            TotSaldoAnt = Redondear(TSaldo - THaber + TDebe, 2)
     End Select
  End If
  CalculosSaldoAnt = TotSaldoAnt
End Function

Public Sub EncabezadoEstResult12Meses(Datas As Adodc, _
                                      IRR As Integer, _
                                      JRR As Integer)
Dim IX As Single
Dim NumeroDeMeses As Integer
  PorteLetra = Printer.FontSize
  LetraAnterior = Printer.FontName
  Printer.FontName = TipoTimes: Printer.FontBold = True
  Printer.FontSize = 9: PosLinea = 0.5
  Cadena = "Página No. " & Pagina
  PrinterTexto Ancho(0), PosLinea + 0.2, Cadena
  Printer.FontSize = 18
  PrinterTexto CentrarTexto(Empresa), PosLinea + 0.1, Empresa
  Printer.FontSize = 15
  PrinterTexto CentrarTexto(SQLMsg1), PosLinea + 0.8, SQLMsg1
  Printer.FontSize = 10
  PrinterTexto CentrarTexto(SQLMsg2), PosLinea + 1.4, SQLMsg2
  Printer.FontSize = 12: PosLinea = 2.5: IX = 0
 '========================================================================
  PrinterTexto Ancho(0), PosLinea + 0.2, "D E T A L L E"
  Printer.FontSize = 7.5
  Si_No = False
  
  'NumeroDeMeses = Month(DFechaMes)
  
  
  With Datas.Recordset
  KE = 2
  For J = IRR To JRR
      PrinterTexto Ancho(KE), PosLinea, UCaseStrg(.fields(J).Name)
      PrinterTexto Ancho(KE), PosLinea + 0.5, "Parcial"
      PrinterTexto Ancho(KE) + 1.8, PosLinea + 0.5, "Total"
      KE = KE + 1
  Next J
  For J = 0 To 8
   If J <> 1 Then Printer.Line (Ancho(J), PosLinea - 0.1)-(Ancho(J), PosLinea + 0.9), Negro
  Next J
  Printer.Line (Ancho(0), PosLinea - 0.1)-(Ancho(8), PosLinea - 0.1), Negro
  Printer.Line (Ancho(2), PosLinea + 0.4)-(Ancho(8), PosLinea + 0.4), Negro
  Printer.Line (Ancho(0), PosLinea + 0.9)-(Ancho(8), PosLinea + 0.9), Negro
  PosLinea = 3.45
 '========================================================================
  End With
  Printer.FontSize = PorteLetra
  Printer.FontName = LetraAnterior
  Printer.FontBold = False
  Pagina = Pagina + 1
End Sub

Public Sub EncabezadoEstResult12MesesP(Datas As Adodc)
Dim IX As Single
Dim NumeroDeMeses As Integer
  PorteLetra = Printer.FontSize
  LetraAnterior = Printer.FontName
  Printer.FontName = TipoTimes
  Printer.FontBold = True
  Printer.FontSize = 9: PosLinea = 0.5
  Cadena = "Página No. " & Pagina
  PrinterTexto Ancho(0), PosLinea + 0.2, Cadena
  Printer.FontSize = 18
  PrinterTexto CentrarTexto(Empresa), PosLinea + 0.1, Empresa
  Printer.FontSize = 15
  PrinterTexto CentrarTexto(SQLMsg1), PosLinea + 0.8, SQLMsg1
  Printer.FontSize = 10
  PrinterTexto CentrarTexto(SQLMsg2), PosLinea + 1.4, SQLMsg2
  Printer.FontSize = 8: PosLinea = 2.5: IX = 0
 '========================================================================
  PrinterTexto Ancho(0), PosLinea, "D E T A L L E"
  Printer.FontSize = 6: NoMeses = 1
  With Datas.Recordset
  For J = 3 To CantCampos - 1
      PrinterTexto Ancho(J), PosLinea, UCaseStrg(MidStrg(.fields(J).Name, 1, 10))
      NoMeses = NoMeses + 1
  Next J
  Printer.Line (Ancho(0), PosLinea - 0.1)-(Ancho(0), PosLinea + 0.35), Negro
  For J = 0 To CantCampos
   If J > 2 Then Printer.Line (Ancho(J), PosLinea - 0.1)-(Ancho(J), PosLinea + 0.35), Negro
  Next J
  Printer.Line (Ancho(0), PosLinea - 0.1)-(Ancho(CantCampos), PosLinea - 0.1), Negro
  Printer.Line (Ancho(0), PosLinea + 0.35)-(Ancho(CantCampos), PosLinea + 0.35), Negro
  End With
  PosLinea = 2.9
 '========================================================================
  Printer.FontSize = PorteLetra
  Printer.FontName = LetraAnterior
  Printer.FontBold = False
  Pagina = Pagina + 1
End Sub

Public Sub EncabezadoDataGeneral(Datas As Adodc, _
                                 Opcion As Byte, _
                                 BG As Boolean)
Dim IX As Single
Dim Son_Iguales As Boolean
Son_Iguales = False
PorteLetra = Printer.FontSize
LetraAnterior = Printer.FontName
Printer.FontName = TipoTimes: Printer.FontBold = True
Printer.FontSize = 9: PosLinea = 0
Cadena = "Página No. " & Pagina
PrinterTexto Ancho(0), PosLinea + 0.2, Cadena
   If UCaseStrg(Empresa) = UCaseStrg(NombreComercial) Then Son_Iguales = True
   PosLinea = PosLinea + 0.1
   If Son_Iguales Then
      Printer.FontSize = 18
      PrinterTexto CentrarTexto(Empresa), PosLinea, Empresa
   Else
      Printer.FontSize = 14
      PrinterTexto CentrarTexto(Empresa), PosLinea, Empresa
      PosLinea = PosLinea + 0.5
      Printer.FontSize = 16
      PrinterTexto CentrarTexto(NombreComercial), PosLinea, NombreComercial
   End If
PosLinea = PosLinea + 0.6
Printer.FontSize = 16
PrinterTexto CentrarTexto(SQLMsg1), PosLinea, SQLMsg1
Printer.FontSize = 10
PosLinea = PosLinea + 0.6
PrinterTexto CentrarTexto(SQLMsg2), PosLinea, SQLMsg2
PosLinea = PosLinea + 0.6
If Opcion = 1 Then PrinterTexto Ancho(0), PosLinea, "CODIGO"
IX = 1
'========================================================================
PrinterTexto Ancho(0 + IX), PosLinea, "CUENTA"
If OpcCoop Then
   If BG Then PrinterTexto Ancho(1 + IX), PosLinea, "S A L D O  M/E"
   PrinterTexto Ancho(2 + IX), PosLinea, "S A L D O  M/N"
Else
   PrinterTexto Ancho(1 + IX), PosLinea, "ANALITICO"
   PrinterTexto Ancho(2 + IX), PosLinea, "PARCIAL"
End If
PrinterTexto Ancho(3 + IX), PosLinea, "T O T A L"
Printer.Line (Ancho(0), PosLinea - 0.1)-(Ancho(4 + IX), PosLinea - 0.1), Negro
Printer.Line (Ancho(0), PosLinea + 0.4)-(Ancho(4 + IX), PosLinea + 0.4), Negro
PosLinea = PosLinea + 0.5
'========================================================================
Printer.FontBold = False
Printer.FontSize = PorteLetra
Printer.FontName = LetraAnterior
Pagina = Pagina + 1
End Sub

Public Sub EncabDiarioGeneralSimple()
Dim IX As Single
PorteLetra = Printer.FontSize
LetraAnterior = Printer.FontName
Printer.FontName = TipoTimes
Printer.FontBold = True
Printer.FontSize = 9: PosLinea = 0.1
Cadena = "Página No. " & Pagina
PrinterTexto Ancho(0), PosLinea + 0.2, Cadena
Printer.FontSize = 20
PrinterTexto CentrarTexto(Empresa), PosLinea, Empresa
PosLinea = PosLinea + 0.9
Printer.FontSize = 18
PrinterTexto CentrarTexto(SQLMsg1), PosLinea, SQLMsg1
PosLinea = PosLinea + 0.7
Printer.FontSize = 10
PrinterTexto CentrarTexto(SQLMsg2), PosLinea, SQLMsg2
PosLinea = PosLinea + 0.6
IX = 0: Printer.FontSize = 8
'========================================================================
Imprimir_Linea_H PosLinea, Ancho(0), Ancho(7)
PosLinea = PosLinea + 0.05
PrinterTexto Ancho(0), PosLinea, "COMPROB."
'PrinterTexto Ancho(1), PosLinea, "NUMERO"
PrinterTexto Ancho(2), PosLinea, "CONCEPTO"
PrinterTexto Ancho(3), PosLinea, "CODIGO"
PrinterTexto Ancho(4), PosLinea, "CUENTA"
PrinterTexto Ancho(5), PosLinea, "DEBE"
PrinterTexto Ancho(6), PosLinea, "HABER"
PosLinea = PosLinea + 0.35
Imprimir_Linea_H PosLinea, Ancho(0), Ancho(7)
PosLinea = PosLinea + 0.1
'========================================================================
Printer.FontBold = False
Printer.FontSize = PorteLetra
Printer.FontName = LetraAnterior
Pagina = Pagina + 1
End Sub

Public Sub EncabezadoDataSubCta(Datas1 As Adodc, _
                                MN_ME As Boolean, _
                                ImpResumen As Boolean)
Dim InicX As Single
Dim InicY As Single
PorteLetra = Printer.FontSize
LetraAnterior = Printer.FontName
Printer.FontName = TipoTimes
InicX = 0.5: InicY = 0.1
PrinterPaint LogoTipo, 0.5, 0, 3, 2
Printer.FontSize = 10
Cadena = "Página No. " & Pagina & "."
PrinterTexto TextoDerecha(Cadena), InicY, Cadena
Cadena = "Fecha: " & Date$ & "."
PrinterTexto TextoDerecha(Cadena), InicY + 1, Cadena
PosLinea = InicY
Printer.FontBold = True
Printer.FontSize = 20
PrinterTexto CentrarTexto(Empresa), PosLinea + 0.5, Empresa
PosLinea = 2
Printer.FontSize = 10
Printer.Line (Ancho(0), PosLinea)-(Ancho(CantCampos), PosLinea + 2), Negro, B
PosLinea = PosLinea + 0.1
PrinterTexto 2, PosLinea, "GRUPO:"
With Datas1.Recordset
 If .RecordCount > 0 Then
     If (.fields("PAX") <> 0) And (.fields("Guia") <> Ninguno) Then
         PrinterTexto 2, PosLinea + 0.5, "GUIA:"
         PrinterFields 3.5, PosLinea + 0.5, Datas1, "Guia"
         PrinterTexto 9, PosLinea + 0.5, "No. PAX:"
         PrinterFields 10.5, PosLinea + 0.5, Datas1, "PAX"
     End If
     PrinterTexto 2, PosLinea + 1, "PERIODO:"
     PrinterTexto 3.5, PosLinea, SQLMsg1
     Cadena = "DESDE " & .fields("Fecha_De") & " AL " & .fields("Fecha_Al")
     PrinterTexto 4, PosLinea + 1, Cadena
 End If
End With
Printer.FontSize = 8: PosLinea = PosLinea + 2
'========================================================================
Printer.FontUnderline = True
If ImpResumen = False Then
PrinterTexto Ancho(0), PosLinea, "FECHA"
PrinterTexto Ancho(1), PosLinea, "BENEFICIARIO"
PrinterTexto Ancho(2), PosLinea, "CUENTA"
PrinterTexto Ancho(3), PosLinea, "TP"
PrinterTexto Ancho(4), PosLinea, "NUMERO"
PrinterTexto Ancho(5), PosLinea, "DEBITOS"
PrinterTexto Ancho(6), PosLinea, "CREDITOS"
If MN_ME Then PrinterTexto Ancho(7), PosLinea, "VALOR/ME"
End If
Printer.FontUnderline = False
Printer.FontBold = False
Pagina = Pagina + 1
PosLinea = PosLinea + 0.5
Printer.FontSize = PorteLetra
Printer.FontName = LetraAnterior
End Sub

Public Function CalculosTotalBancos(DataBanco As Adodc) As Currency
Dim SumaBank As Currency
With DataBanco.Recordset
     SumaBank = 0
     If .RecordCount > 0 Then
        .MoveFirst
         Do While Not DataBanco.Recordset.EOF
            SumaBank = SumaBank + .fields("VALOR")
           .MoveNext
         Loop
     End If
End With
CalculosTotalBancos = SumaBank
End Function

Public Sub Select_Cuentas(DBLists As DataList, DtaCtas As Adodc, Optional TipoCtas As String)
Dim DetalleCta As String
Dim NumeroCta As String
  TipoCtas = UCaseStrg(TipoCtas)
  DetalleCta = TipoCtas
  NumeroCta = Replace(TipoCtas, ".", "")
  sSQL = "SELECT (Codigo & Space(19-LEN(Codigo)) & TC & Space(3-LEN(TC)) & STR(Clave) & ' - ' & Cuenta) As Nombre_Cuenta " _
       & "FROM Catalogo_Cuentas " _
       & "WHERE DG = 'D' " _
       & "AND Cuenta <> '" & Ninguno & "' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' "
  If TipoCtas <> "" Then
     LongStrg = Len(TipoCtas)
     If IsNumeric(NumeroCta) Then
        sSQL = sSQL & "AND MidStrg(Codigo,1," & LongStrg & ") = '" & TipoCtas & "' "
     Else
        If 1 <= LongStrg And LongStrg <= 2 Then
           sSQL = sSQL & "AND TC  = '" & MidStrg(TipoCtas, 1, 2) & "' "
        Else
           sSQL = sSQL & "AND Cuenta LIKE '%" & TipoCtas & "%' "
        End If
     End If
  End If
  sSQL = sSQL & "ORDER BY Codigo, Cuenta "
  'MsgBox sSQL
  SelectDB_List DBLists, DtaCtas, sSQL, "Nombre_Cuenta"
End Sub

Public Sub SelectBenef(DBComboBenef As DataCombo, _
                       DtaBenef As Adodc)
  sSQL = "SELECT Detalle " _
       & "FROM Catalogo_SubCtas " _
       & "WHERE TC = '" & Normal & "' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "ORDER BY Detalle "
  SelectDB_Combo DBComboBenef, DtaBenef, sSQL, "Detalle", False
End Sub

Public Sub SumaTotalAsientos(DataAsiento As Adodc) ', LabelD As Control, LabelH As Control)
   SumaDebe = 0: SumaHaber = 0
   With DataAsiento.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       Do While Not .EOF
          SumaDebe = SumaDebe + .fields("Debe")
          SumaHaber = SumaHaber + .fields("Haber")
         .MoveNext
       Loop
      .MoveFirst
   End If
   End With
   SumaDebe = Redondear(SumaDebe, 2)
   SumaHaber = Redondear(SumaHaber, 2)
End Sub

Public Sub CalculosTotalAsientos(DataAsiento As Adodc, _
                                 LabelD As Control, _
                                 LabelH As Control, _
                                 LabelDi As Control)
   SumaDebe = 0: SumaHaber = 0: Total_RetCta = 0
   LlenarRetencion = False
   With DataAsiento.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       Do While Not .EOF
          'If InStr(Co.Ctas_Modificar, .fields("CODIGO")) <= 0 Then Co.Ctas_Modificar = Co.Ctas_Modificar & .fields("CODIGO") & ","
          If Cta_Ret_Egreso = .fields("CODIGO") Then
             LlenarRetencion = True
             Total_RetCta = .fields("DEBE")
             If Total_RetCta <= 0 Then Total_RetCta = .fields("HABER")
          End If
          SumaDebe = SumaDebe + .fields("DEBE")
          SumaHaber = SumaHaber + .fields("HABER")
         'MsgBox "hola"
         .MoveNext
       Loop
   End If
   End With
   
   LabelD.Caption = Format$(SumaDebe, "#,##0.00")
   LabelH.Caption = Format$(SumaHaber, "#,##0.00")
   LabelDi.Caption = Format$(SumaDebe - SumaHaber, "#,##0.00")
End Sub

Public Sub Grabar_Comprobante(C1 As Comprobantes)
Dim ConBodegas As Boolean
Dim NumTrans As Long
Dim CodSustento As String
Dim AdoTemp As ADODB.Recordset
Dim AdoRetAut As ADODB.Recordset

  RatonReloj
 'Grabamos los datos de la transaccion en la tabla definitiva de almacenamiento
  TMail.TipoDeEnvio = Ninguno
  If Len(C1.Autorizacion_LC) >= 13 Then TMail.TipoDeEnvio = "CE"
  If Len(C1.Autorizacion_R) > 13 Then TMail.TipoDeEnvio = "CE"
  FA.TP = C1.TP
  FA.Fecha = C1.Fecha
  FA.Numero = C1.Numero
  FA.ClaveAcceso = Ninguno
  FA.Autorizacion_R = C1.Autorizacion_R
  FA.Autorizacion_LC = C1.Autorizacion_LC
  FA.Retencion = C1.Retencion
  FA.Serie_R = C1.Serie_R
  FA.Serie_LC = C1.Serie_LC
  
  Grabar_Comprobante_SP C1, CtaConciliada
  FA.TP = C1.TP
  FA.Numero = C1.Numero
 'MsgBox C1.Autorizacion_R & vbCrLf & C1.Autorizacion_LC & vbCrLf & C1.TP & "-" & C1.Numero & vbCrLf & C1.GrabadoExitoso
  If C1.GrabadoExitoso Then
     Control_Procesos Normal, "Grabar Comprobante de: " & C1.TP & " No. " & C1.Numero
    'Actualizar las Ctas a mayoriazar
    'Actualiza_Procesado_Kardex .fields("Codigo_Inv")
    'Pasamos a Autorizar la retencion si es electronica
     RatonReloj
     'MsgBox FA.Autorizacion_R & vbCrLf & FA.Autorizacion_LC
     If Len(C1.Autorizacion_R) >= 13 Then SRI_Crear_Clave_Acceso_Retenciones FA, True
     If Len(C1.Autorizacion_LC) >= 13 Then SRI_Crear_Clave_Acceso_Liquidacion FA, True
  End If
  RatonNormal
End Sub

Public Sub InsertarAsientosC(DtaAsiento As Adodc)
  If IsEmpty(CodigoCli) Then CodigoCli = Ninguno
  If IsNull(CodigoCli) Then CodigoCli = Ninguno
  If Codigo <> Ninguno Then
     Debe = 0: Haber = 0
     Select Case OpcDH
       Case 1: Debe = ValorDH
       Case 2: Haber = ValorDH
     End Select
     If ValorDH <> 0 Then
        With DtaAsiento.Recordset
         .AddNew
         .fields("CODIGO") = Codigo
         .fields("CUENTA") = Cuenta
         .fields("DETALLE") = DetalleComp
         .fields("Item") = NumEmpresa
          If OpcCoop And Moneda_US Then
             Debe = Debe / Dolar
             Haber = Haber / Dolar
          Else
            .fields("PARCIAL_ME") = 0
             If Moneda_US Then .fields("PARCIAL_ME") = (Debe - Haber) / Dolar
          End If
         .fields("DEBE") = Debe
         .fields("HABER") = Haber
         .fields("EFECTIVIZAR") = Fecha_Vence
         .fields("CHEQ_DEP") = NoCheque
         .fields("ME") = Moneda_US
         .fields("T_No") = Trans_No
         .fields("CODIGO_C") = CodigoCli
         .fields("CODIGO_CC") = CodigoCC
         .fields("Item") = NumEmpresa
         .fields("CodigoU") = CodigoUsuario
         .fields("TC") = SubCta
         .fields("A_No") = Ln_No
          Ln_No = Ln_No + 1
          If Cuenta <> Ninguno Then .Update
        End With
     End If
  End If
End Sub

Public Sub ImprimirAsientos(DataT As ADODB.Recordset, _
                            DataC As ADODB.Recordset, _
                            Datas1 As ADODB.Recordset, _
                            Datas2 As ADODB.Recordset, _
                            AsientosX As Long, _
                            SumaTotales As Long, _
                            PonerLineas As Boolean, _
                            TipoComp As String, _
                            EstaAnulado As String, _
                            Dolares As Single, _
                            Optional Pos_Y_chq As Single)
Dim NumAsientos As Long
Dim PosAsiento(6) As Single
Dim PosLineaInicio As Single
Dim PosLineaTotal As Single
Dim TamanoLetra As Single
Dim NUsuario As String
Dim NUser1 As String
Dim NUser2 As String
Dim NUser3 As String
Dim NUser4 As String
Dim NDolares As String
Dim EsPropio As Single
Dim DetSubCta As String
Dim ImpDetSubCta As Boolean
For I = 0 To 4
    PosAsiento(I) = SetD(AsientosX + I).PosX
Next I
PosAsiento(5) = PosAsiento(4) + 3
TamanoLetra = SetD(AsientosX).Tamaño
Printer.FontSize = TamanoLetra
PosLinea = SetD(AsientosX).PosY
PosLinea = PosLinea + 0.1
PosLineaInicio = PosLinea - 0.1
SumaDebe = 0: SumaHaber = 0
If PonerLineas Then NoColor = 0 Else NoColor = Blanco
With DataT
If .RecordCount > 0 Then
   .MoveLast
    NumAsientos = .RecordCount
    If PonerLineas Then
       If NumAsientos >= 8 Then PosLineaTotal = SetD(AsientosX).PosY + (NumAsientos + 1) * 0.4 Else PosLineaTotal = 10.5 + Pos_Y_chq
    Else
       PosLineaTotal = SetD(SumaTotales).PosY - 0.1
    End If
    If PosLineaTotal >= 23 Then PosLineaTotal = 23
    I = 0
   .MoveFirst
    Printer.FontBold = False
    Do While Not .EOF
       Printer.FontItalic = False
       Printer.FontSize = TamanoLetra
    
      Valor_ME = .fields("Parcial_ME")
      Cta_Sup = .fields("Cta")
      Valor_Prom = .fields("Debe") - .fields("Haber")
      PrinterFields PosAsiento(0) + 0.1, PosLinea, .fields("Cta")
      PrinterFields PosAsiento(1) + 0.1, PosLinea, .fields("Cuenta")
      If OpcCoop Then
         If Valor_ME <> 0 Then
            PrinterFields PosAsiento(3) + 0.1, PosLinea, .fields("Debe_ME")
            PrinterFields PosAsiento(4) + 0.1, PosLinea, .fields("Haber_ME")
            SumaDebe = SumaDebe + .fields("Debe")
            SumaHaber = SumaHaber + .fields("Haber")
         Else
            PrinterFields PosAsiento(3) + 0.1, PosLinea, .fields("Debe")
            PrinterFields PosAsiento(4) + 0.1, PosLinea, .fields("Haber")
            SumaDebe = SumaDebe + .fields("Debe")
            SumaHaber = SumaHaber + .fields("Haber")
         End If
      Else
         If Valor_ME <> 0 Then PrinterVariables PosAsiento(2) + 0.1, PosLinea, Valor_ME
         Printer.FontItalic = False
         PrinterFields PosAsiento(3) + 0.1, PosLinea, .fields("Debe")
         PrinterFields PosAsiento(4) + 0.1, PosLinea, .fields("Haber")
         For J = 0 To 5
             Printer.Line (PosAsiento(J), PosLinea - 0.05)-(PosAsiento(J), PosLinea + 0.4), NoColor
         Next J
         SumaDebe = SumaDebe + .fields("Debe")
         SumaHaber = SumaHaber + .fields("Haber")
      End If
      PosLinea = PosLinea + 0.35
      If Len(.fields("Detalle")) > 1 Then
         Printer.FontSize = 7
         PrinterFields PosAsiento(1) + 0.1, PosLinea, .fields("Detalle")
         If .fields("Cheq_Dep") <> Ninguno Then
             If IsNumeric(.fields("Cheq_Dep")) Then
                PrinterTexto PosAsiento(1) + 5.1, PosLinea, " No. " & .fields("Cheq_Dep")
             Else
                PrinterTexto PosAsiento(1) + 5.1, PosLinea, " " & .fields("Cheq_Dep")
             End If
         End If
         PosLinea = PosLinea + 0.35
      Else
         If .fields("Cheq_Dep") <> Ninguno Then
             Printer.FontSize = 7
             If IsNumeric(.fields("Cheq_Dep")) Then
                Select Case TipoComp
                  Case "CD": PrinterTexto PosAsiento(1) + 0.1, PosLinea, "No. " & Format$(Val(.fields("Cheq_Dep")), "00000000")
                  Case "CI": PrinterTexto PosAsiento(1) + 0.1, PosLinea, "Deposito No. " & Format$(Val(.fields("Cheq_Dep")), "00000000")
                  Case "CE": PrinterTexto PosAsiento(1) + 0.1, PosLinea, "Cheque No. " & Format$(Val(.fields("Cheq_Dep")), "00000000")
                  Case Else: PrinterTexto PosAsiento(1) + 0.1, PosLinea, "Documento No. " & Format$(Val(.fields("Cheq_Dep")), "00000000")
                End Select
             Else
                PrinterTexto PosAsiento(1) + 0.1, PosLinea, "Referencia: " & .fields("Cheq_Dep")
             End If
             PosLinea = PosLinea + 0.35
         End If
      End If
      
      If Datas1.RecordCount > 0 Then
         Printer.FontSize = 7
         'PosLinea = PosLinea + 0.35
         Datas1.MoveFirst
         Datas1.Find ("Cta = '" & Cta_Sup & "' ")
         If Not Datas1.EOF Then
            DetSubCta = Ninguno
            Do While Not Datas1.EOF
               Printer.FontItalic = True
               If Cta_Sup = Datas1.fields("Cta") And Datas1.fields("Codigo") <> Ninguno Then
                  If Datas1.fields("Debitos") > 0 And Valor_Prom > 0 Then
                     ImpDetSubCta = False
                     If DetSubCta <> Datas1.fields("Cliente") Then
                        PrinterTexto PosAsiento(1) + 0.1, PosLinea, "-" & Datas1.fields("Cliente")
                        DetSubCta = Datas1.fields("Cliente")
                        ImpDetSubCta = True
                     End If
                     Printer.FontItalic = False
                     
                     If Datas1.fields("Detalle_SubCta") <> Ninguno Then
                        If ImpDetSubCta Then PosLinea = PosLinea + 0.35
                        PrinterTexto PosAsiento(1) + 0.2, PosLinea, Datas1.fields("Detalle_SubCta")
                     End If
                     
                     If Datas1.fields("Factura") <> 0 Then
                        PrinterTexto PosAsiento(1) + 5.1, PosLinea, " No. " & Format$(Datas1.fields("Factura"), "0000000")
                     ElseIf Datas1.fields("Prima") <> 0 Then
                        PrinterTexto PosAsiento(1) + 5.1, PosLinea, " Valor Prima"
                     Else
                        If Datas1.fields("TC") <> "CC" Then PrinterTexto PosAsiento(1) + 4.6, PosLinea, " - Venc. " & Datas1.fields("Fecha_V")
                     End If
                     
                     If Datas1.fields("Prima") <> 0 Then
                        PrinterVariables PosAsiento(2) + 0.1, PosLinea, Datas1.fields("Prima")
                     Else
                        PrinterFields PosAsiento(2) + 0.1, PosLinea, Datas1.fields("Debitos")
                     End If
                     PosLinea = PosLinea + 0.35
                  End If
                  If Datas1.fields("Creditos") > 0 And Valor_Prom < 0 Then
                     ImpDetSubCta = False
                     If DetSubCta <> Datas1.fields("Cliente") Then
                        PrinterTexto PosAsiento(1) + 0.1, PosLinea, "-" & Datas1.fields("Cliente")
                        DetSubCta = Datas1.fields("Cliente")
                        ImpDetSubCta = True
                     End If
                     Printer.FontItalic = False
                     
                     If Datas1.fields("Detalle_SubCta") <> Ninguno Then
                        If ImpDetSubCta Then PosLinea = PosLinea + 0.35
                        PrinterTexto PosAsiento(1) + 0.2, PosLinea, Datas1.fields("Detalle_SubCta")
                     End If
                     If Datas1.fields("Factura") <> 0 Then
                        PrinterTexto PosAsiento(1) + 5.1, PosLinea, " No. " & Format$(Datas1.fields("Factura"), "0000000")
                     ElseIf Datas1.fields("Prima") <> 0 Then
                        PrinterTexto PosAsiento(1) + 5.1, PosLinea, " Valor Prima"
                     Else
                        If Datas1.fields("TC") <> "CC" Then PrinterTexto PosAsiento(1) + 4.6, PosLinea, " - Venc. " & Datas1.fields("Fecha_V")
                     End If
                     If Datas1.fields("Prima") <> 0 Then
                        PrinterVariables PosAsiento(2) + 0.1, PosLinea, Datas1.fields("Prima")
                     Else
                        PrinterFields PosAsiento(2) + 0.9, PosLinea, Datas1.fields("Creditos")
                     End If
                     PosLinea = PosLinea + 0.35
                  End If
                  
                  If PosLinea >= 23 Then
                     For J = 0 To 5
                         Printer.Line (PosAsiento(J), PosLineaInicio - 0.05)-(PosAsiento(J), PosLinea), NoColor
                     Next J
                     Printer.Line (0.5, PosLinea)-(PosAsiento(5), PosLinea), NoColor
                     Printer.FontItalic = False
                     Printer.FontBold = False
                     Printer.FontSize = SetD(1).Tamaño
                     PrinterTexto SetD(1).PosX, SetD(1).PosY, msg
                     Printer.FontSize = 7
                     Printer.FontItalic = True
                     PosLinea = PosLinea + 0.1
                     PrinterTexto PosAsiento(1), PosLinea, " CONTINUA EN LA SIGUIENTE PAGINA..."
                     PosLineaTotal = 23
                     Printer.NewPage
                     Dibujo = RutaSistema & "\FORMATOS\DIARIO1.GIF"
                     PrinterPaint Dibujo, 0.5, 0, 19, 3.5
                     EncabezadoEmpresa 0.1
                     Printer.FontItalic = False
                     Printer.FontBold = False
                     Printer.FontSize = TamanoLetra
                     PrinterTexto SetD(2).PosX, SetD(2).PosY, Mifecha
                     PrinterTexto 17.5, 1.9, Pagina & "."
                     Printer.FontName = LetraAnterior
                     PosLinea = 3.6
                     PosLineaInicio = PosLinea
                     I = 0
                     Pagina = Pagina + 1
                     Printer.FontSize = 7
                  End If
               End If
               Datas1.MoveNext
            Loop
         End If
      End If
      
      If Datas2.RecordCount > 0 Then
         Printer.FontSize = 7
         'PosLinea = PosLinea + 0.35
         Datas2.MoveFirst
         Datas2.Find ("Cta = '" & Cta_Sup & "' ")
         If Not Datas2.EOF Then
            DetSubCta = Ninguno
            Do While Not Datas2.EOF
               Printer.FontItalic = True
               If Cta_Sup = Datas2.fields("Cta") And Datas2.fields("Codigo") <> Ninguno Then
                  If Datas2.fields("Debitos") > 0 And Valor_Prom > 0 Then
                     ImpDetSubCta = False
                     If DetSubCta <> Datas2.fields("Cliente") Then
                        PrinterTexto PosAsiento(1) + 0.1, PosLinea, "-" & Datas2.fields("Cliente")
                        DetSubCta = Datas2.fields("Cliente")
                        ImpDetSubCta = True
                     End If
                     Printer.FontItalic = False
                     
                     If Datas2.fields("Detalle_SubCta") <> Ninguno Then
                        If ImpDetSubCta Then PosLinea = PosLinea + 0.35
                        PrinterTexto PosAsiento(1) + 0.2, PosLinea, Datas2.fields("Detalle_SubCta")
                     End If
                     
                     If Datas2.fields("Factura") <> 0 Then
                        PrinterTexto PosAsiento(1) + 5.1, PosLinea, " No. " & Format$(Datas2.fields("Factura"), "0000000")
                     ElseIf Datas2.fields("Prima") <> 0 Then
                        PrinterTexto PosAsiento(1) + 5.1, PosLinea, " Valor Prima"
                     Else
                        If Datas2.fields("TC") <> "CC" Then PrinterTexto PosAsiento(1) + 4.6, PosLinea, " - Venc. " & Datas2.fields("Fecha_V")
                     End If
                     
                     If Datas2.fields("Prima") <> 0 Then
                        PrinterVariables PosAsiento(2) + 0.1, PosLinea, Datas2.fields("Prima")
                     Else
                        PrinterFields PosAsiento(2) + 0.1, PosLinea, Datas2.fields("Debitos")
                     End If
                     PosLinea = PosLinea + 0.35
                  End If
                  If Datas2.fields("Creditos") > 0 And Valor_Prom < 0 Then
                     ImpDetSubCta = False
                     If DetSubCta <> Datas2.fields("Cliente") Then
                        PrinterTexto PosAsiento(1) + 0.1, PosLinea, "-" & Datas2.fields("Cliente")
                        DetSubCta = Datas2.fields("Cliente")
                        ImpDetSubCta = True
                     End If
                     Printer.FontItalic = False
                     
                     If Datas2.fields("Detalle_SubCta") <> Ninguno Then
                        If ImpDetSubCta Then PosLinea = PosLinea + 0.35
                        PrinterTexto PosAsiento(1) + 0.2, PosLinea, Datas2.fields("Detalle_SubCta")
                     End If
                     If Datas2.fields("Factura") <> 0 Then
                        PrinterTexto PosAsiento(1) + 5.1, PosLinea, " No. " & Format$(Datas2.fields("Factura"), "0000000")
                     ElseIf Datas2.fields("Prima") <> 0 Then
                        PrinterTexto PosAsiento(1) + 5.1, PosLinea, " Valor Prima"
                     Else
                        If Datas2.fields("TC") <> "CC" Then PrinterTexto PosAsiento(1) + 4.6, PosLinea, " - Venc. " & Datas2.fields("Fecha_V")
                     End If
                     If Datas2.fields("Prima") <> 0 Then
                        PrinterVariables PosAsiento(2) + 0.1, PosLinea, Datas2.fields("Prima")
                     Else
                        PrinterFields PosAsiento(2) + 0.9, PosLinea, Datas2.fields("Creditos")
                     End If
                     PosLinea = PosLinea + 0.35
                  End If
                  
                  If PosLinea >= 23 Then
                     For J = 0 To 5
                         Printer.Line (PosAsiento(J), PosLineaInicio - 0.05)-(PosAsiento(J), PosLinea), NoColor
                     Next J
                     Printer.Line (0.5, PosLinea)-(PosAsiento(5), PosLinea), NoColor
                     Printer.FontItalic = False
                     Printer.FontBold = False
                     Printer.FontSize = SetD(1).Tamaño
                     PrinterTexto SetD(1).PosX, SetD(1).PosY, msg
                     Printer.FontSize = 7
                     Printer.FontItalic = True
                     PosLinea = PosLinea + 0.1
                     PrinterTexto PosAsiento(1), PosLinea, " CONTINUA EN LA SIGUIENTE PAGINA..."
                     PosLineaTotal = 23
                     Printer.NewPage
                     Dibujo = RutaSistema & "\FORMATOS\DIARIO1.GIF"
                     PrinterPaint Dibujo, 0.5, 0, 19, 3.5
                     EncabezadoEmpresa 0.1
                     Printer.FontItalic = False
                     Printer.FontBold = False
                     Printer.FontSize = TamanoLetra
                     PrinterTexto SetD(2).PosX, SetD(2).PosY, Mifecha
                     PrinterTexto 17.5, 1.9, Pagina & "."
                     Printer.FontName = LetraAnterior
                     PosLinea = 3.6
                     PosLineaInicio = PosLinea
                     I = 0
                     Pagina = Pagina + 1
                     Printer.FontSize = 7
                  End If
               End If
               Datas2.MoveNext
            Loop
         End If
      End If
      
      'If Datas1.RecordCount = 0 And Datas2.RecordCount = 0 Then PosLinea = PosLinea + 0.35
      
      If PosLinea >= 23 Then
         PosLinea = PosLinea + 0.35
         For J = 0 To 5
             Printer.Line (PosAsiento(J), PosLineaInicio - 0.05)-(PosAsiento(J), PosLinea), NoColor
         Next J
         Printer.Line (0.5, PosLinea)-(PosAsiento(5), PosLinea), NoColor
         Printer.FontItalic = False
         Printer.FontBold = False
         Printer.FontSize = SetD(1).Tamaño
         PrinterTexto SetD(1).PosX, SetD(1).PosY, msg
         Printer.FontSize = 7
         Printer.FontItalic = True
         PosLinea = PosLinea + 0.1
         PrinterTexto PosAsiento(1), PosLinea, " CONTINUA EN LA SIGUIENTE PAGINA..."
         PosLineaTotal = 23
         Printer.NewPage
         Dibujo = RutaSistema & "\FORMATOS\DIARIO1.GIF"
         PrinterPaint Dibujo, 0.5, 0, 19, 3.5
         EncabezadoEmpresa 0.1
         Printer.FontItalic = False
         Printer.FontBold = False
         Printer.FontSize = TamanoLetra
         PrinterTexto SetD(2).PosX, SetD(2).PosY, Mifecha
         PrinterTexto 17.5, 1.9, Pagina & "."
         Printer.FontName = LetraAnterior
         PosLinea = 3.6
         PosLineaInicio = PosLinea
         I = 0
         Pagina = Pagina + 1
      End If
      I = I + 1
     .MoveNext
    Loop
    Printer.FontItalic = False
    Printer.FontBold = False
    Printer.FontSize = SetD(1).Tamaño
    PrinterTexto SetD(1).PosX, SetD(1).PosY, msg
    Printer.FontSize = TamanoLetra
    If PosLinea > PosLineaTotal Then PosLineaTotal = PosLinea
    For J = 0 To 5
        Printer.Line (PosAsiento(J), PosLineaInicio - 0.05)-(PosAsiento(J), PosLineaTotal), NoColor
    Next J
    If PonerLineas Then
       Dibujo = RutaSistema & "\FORMATOS\TOTALES.GIF"
       PrinterPaint Dibujo, 0.5, PosLineaTotal, 19, 2
       If EstaAnulado = "A" Then
          Dibujo = RutaSistema & "\FORMATOS\ANULADO.GIF"
          PrinterPaint Dibujo, 1, PosLineaTotal - 1.7, 6, 1.5
       End If
    End If
    NUsuario = ULCase(DataC.fields("Nombre_Completo"))
    NDolares = " "
    If Redondear(DataC.fields("Cotizacion"), 2) <> 0 Then
       NDolares = "U.S.$ " & Format$(DataC.fields("Cotizacion"), "#,###.##")
    End If
    Printer.FontItalic = False
    Printer.FontSize = SetD(AsientosX + 3).Tamaño
    PrinterVariables SetD(AsientosX + 3).PosX, PosLineaTotal + 0.1, SumaDebe
    PrinterVariables SetD(AsientosX + 4).PosX, PosLineaTotal + 0.1, SumaHaber
   'Abreviamos el nombre del usuario
    NUser1 = TrimStrg(SinEspaciosIzq(NUsuario))
    NUsuario = TrimStrg(MidStrg(NUsuario, Len(NUser1) + 1, Len(NUsuario)))
    
    NUser2 = TrimStrg(SinEspaciosIzq(NUsuario))
    NUsuario = TrimStrg(MidStrg(NUsuario, Len(NUser2) + 1, Len(NUsuario)))
    
    NUser3 = TrimStrg(SinEspaciosIzq(NUsuario))
    NUsuario = TrimStrg(MidStrg(NUsuario, Len(NUser3) + 1, Len(NUsuario)))
    
    NUser4 = TrimStrg(SinEspaciosIzq(NUsuario))
    NUsuario = TrimStrg(MidStrg(NUsuario, Len(NUser4) + 1, Len(NUsuario)))
    
    If NUser3 <> "" Then NUser3 = MidStrg(NUser3, 1, 1) & "."
    If NUser4 <> "" Then NUser4 = MidStrg(NUser4, 1, 1) & "."
    
    NUsuario = TrimStrg(NUser1 & " " & NUser2 & " " & NUser3 & " " & NUser4)
    
    
    EsPropio = 0
    Select Case TipoComp
      Case CompDiario
           TamanoLetra = SetD(10).Tamaño
           EsPropio = SetD(10).PosX
      Case CompEgreso, CompIngreso
           TamanoLetra = SetD(19).Tamaño
           EsPropio = SetD(19).PosX
    End Select
    If TamanoLetra <= 0 Then TamanoLetra = 8
    Printer.FontSize = TamanoLetra
    If EsPropio > 0 Then
       Select Case TipoComp
         Case CompDiario
              PrinterVariables SetD(10).PosX, SetD(10).PosY, NUsuario
         Case CompEgreso, CompIngreso
              PrinterVariables SetD(19).PosX, SetD(19).PosY, NUsuario
       End Select
    Else
       If TamanoLetra > 0 Then
          Printer.FontSize = TamanoLetra
          
          PrinterVariables 2.8, PosLineaTotal + 1, NUsuario
          PrinterVariables 1, PosLineaTotal + 1.5, NDolares
       End If
    End If
End If
End With
End Sub

Public Sub Imprimir_Balance(Datas As Adodc)
On Error GoTo Errorhandler
Dim NombFilePict As String
Dim PosLinInic As Long
Mensajes = "Imprimir Balance"
Titulo = "IMPRESION"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
   RatonReloj
   'Segunda_Pag = False
   If SetNombrePRN = Impresota_PDF Then tPrint.TipoImpresion = Es_PDF Else tPrint.TipoImpresion = Es_Printer
  'Generamos el documento
   tPrint.NombreArchivo = MensajeEncabData & " " & RUC
   tPrint.TituloArchivo = MensajeEncabData & " " & RUC
   tPrint.TipoLetra = TipoArialNarrow ' TipoHelvetica
   tPrint.OrientacionPagina = 1
   tPrint.PaginaA4 = True
   tPrint.EsCampoCorto = False
   tPrint.VerDocumento = False
   Set cPrint = New cImpresion
   cPrint.iniciaImpresion
   InicioX = 0.5: InicioY = 0
    C = 7
    ReDim Ancho(C) As Single
    Ancho(0) = 1
    Ancho(1) = 3.5
    Ancho(2) = 10.5
    Ancho(3) = 13
    Ancho(4) = 15.5
    Ancho(5) = 18
    Ancho(6) = 20.5
    Pagina = 1
    CantCampos = 6
    HoraSistema = Time
    MensajeEncabData = ""
    SQLMsg1 = ""
    SQLMsg2 = ""
    SQLMsg3 = ""
    cPrint.printEncabezado Ancho(0), 1, TipoArialNarrow
    EncabBalance
    PosLinea = PosLinea + 0.35
    PosLinInic = PosLinea
    With Datas.Recordset
     If .RecordCount > 0 Then
        .MoveFirst
         Do While Not .EOF
            cPrint.PorteDeLetra = 9
            cPrint.tipoNormal = True
            If .fields("DG") = "G" Then cPrint.tipoNegrilla = True
            cPrint.printFields Ancho(0), PosLinea, .fields("Codigo")
            If .fields("TC") <> "N" Then cPrint.tipoItalica = True
            Select Case .fields("TC")
              Case "BA", "CJ", "C", "P": cPrint.tipoSubrayado = True
            End Select
            cPrint.printFields Ancho(1), PosLinea, .fields("Cuenta")
            cPrint.tipoItalica = False
            cPrint.tipoSubrayado = False
            cPrint.printFields Ancho(2), PosLinea, .fields("Saldo_Anterior")
            cPrint.printFields Ancho(3), PosLinea, .fields("Debitos")
            cPrint.printFields Ancho(4), PosLinea, .fields("Creditos")
            cPrint.printFields Ancho(5), PosLinea, .fields("Saldo_Total")
            If .fields("DG") = "G" Then
                cPrint.printLinea Ancho(0) - 0.3, PosLinea - 0.4, Ancho(6) - 0.3, PosLinea - 0.4, Negro
                cPrint.printLinea Ancho(0) - 0.3, PosLinea + 0.1, Ancho(6) - 0.3, PosLinea + 0.1, Negro
            End If
            PosLinea = PosLinea + 0.5
            If PosLinea > LimiteAlto Then
               For I = 0 To 6
                   cPrint.printLinea Ancho(I) - 0.3, PosLinInic - 0.3, Ancho(I) - 0.3, PosLinea + 0.4, Negro
               Next I
               cPrint.printLinea Ancho(0), PosLinea, Ancho(6), PosLinea, Negro
               Printer.NewPage
               cPrint.printEncabezado Ancho(0), 1
               EncabBalance
               PosLinea = PosLinea + 0.35
               PosLinInic = PosLinea
               Printer.FontName = TipoArialNarrow
            End If
           .MoveNext
         Loop
     End If
    End With
    For I = 0 To 6
        cPrint.printLinea Ancho(I) - 0.3, PosLinInic - 0.3, Ancho(I) - 0.3, PosLinea + 0.4, Negro
    Next I
    cPrint.printLinea Ancho(0) - 0.3, PosLinea, Ancho(6) - 0.3, PosLinea, Negro
    PosLinea = PosLinea + 0.05
    cPrint.printLinea Ancho(0) - 0.3, PosLinea, Ancho(6) - 0.3, PosLinea, Negro
    PosLinea = PosLinea + 0.1
    Printer.FontBold = True
    PrinterVariables Ancho(3), PosLinea, SumaDebe
    PrinterVariables Ancho(4), PosLinea, SumaHaber
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

'''Public Sub Imprimir_Balance(Datas As Adodc)
'''On Error GoTo Errorhandler
'''Mensajes = "Imprimir Balance"
'''Titulo = "IMPRESION"
'''Bandera = False
'''SetPrinters.Show 1
'''If PonImpresoraDefecto(SetNombrePRN) Then
'''RatonReloj
'''
'''Escala_Centimetro 1, TipoArialNarrow, 9
'''InicioX = 0.5: InicioY = 0
'''C = 7
'''ReDim Ancho(C) As Single
'''Ancho(0) = 0.5
'''Ancho(1) = 3
'''Ancho(2) = 10
'''Ancho(3) = 12.5
'''Ancho(4) = 15
'''Ancho(5) = 17.5
'''Ancho(6) = 20
'''Pagina = 1
'''CantCampos = 6
'''HoraSistema = Time
'''MensajeEncabData = ""
'''Encabezado Ancho(0), Ancho(6)
'''EncabBalance
'''With Datas.Recordset
''' If .RecordCount > 0 Then
'''    .MoveFirst
'''     Printer.FontName = TipoArialNarrow
'''     Do While Not .EOF
'''        Printer.FontSize = 9
'''        Printer.FontBold = False
'''        Printer.FontItalic = False
'''        Printer.FontUnderline = False
'''        If .Fields("DG") = "G" Then Printer.FontBold = True
'''        PrinterFields Ancho(0), PosLinea, .Fields("Codigo")
'''        If .Fields("TC") <> "N" Then Printer.FontItalic = True
'''        Select Case .Fields("TC")
'''          Case "BA", "CJ", "C", "P": Printer.FontUnderline = True
'''        End Select
'''        PrinterFields Ancho(1), PosLinea, .Fields("Cuenta")
'''        Printer.FontItalic = False
'''        Printer.FontUnderline = False
'''        PrinterFields Ancho(2), PosLinea, .Fields("Saldo_Anterior")
'''        PrinterFields Ancho(3), PosLinea, .Fields("Debitos")
'''        PrinterFields Ancho(4), PosLinea, .Fields("Creditos")
'''        PrinterFields Ancho(5), PosLinea, .Fields("Saldo_Total")
'''        If .Fields("DG") = "G" Then
'''            Printer.Line (Ancho(0), PosLinea - 0.1)-(Ancho(6), PosLinea - 0.1), Negro
'''            Printer.Line (Ancho(0), PosLinea + 0.4)-(Ancho(6), PosLinea + 0.4), Negro
'''        End If
'''        For I = 0 To 6
'''            Printer.Line (Ancho(I), PosLinea - 0.1)-(Ancho(I), PosLinea + 0.4), Negro
'''        Next I
'''        PosLinea = PosLinea + 0.5
'''        If PosLinea >= LimiteAlto Then
'''           Printer.Line (Ancho(0), PosLinea)-(Ancho(6), PosLinea), Negro
'''           Printer.NewPage
'''           Encabezado Ancho(0), Ancho(6)
'''           EncabBalance
'''           Printer.FontName = TipoArialNarrow
'''        End If
'''       .MoveNext
'''     Loop
''' End If
'''End With
'''Printer.Line (Ancho(0), PosLinea)-(Ancho(6), PosLinea), Negro
'''PosLinea = PosLinea + 0.05
'''Printer.Line (Ancho(0), PosLinea)-(Ancho(6), PosLinea), Negro
'''PosLinea = PosLinea + 0.1
'''Printer.FontBold = True
'''PrinterVariables Ancho(3), PosLinea, SumaDebe
'''PrinterVariables Ancho(4), PosLinea, SumaHaber
'''RatonNormal
'''MensajeEncabData = ""
'''Printer.EndDoc
'''End If
'''Exit Sub
'''Errorhandler:
'''    RatonNormal
'''    ErrorDeImpresion
'''    Exit Sub
'''End Sub

Public Sub EncabMayor(Texto1 As String, _
                      Texto2 As String, _
                      MargenDer As Single)
'Iniciamos la impresion
PorteLetra = Printer.FontSize
LetraAnterior = Printer.FontName
Printer.FontName = TipoTimes
Printer.FontBold = True
'Printer.Line (0.5, PosLinea)-(MargenDer, PosLinea + 1.5), NEGRO, B
'Printer.Line (0.6, PosLinea + 0.1)-(MargenDer - 0.1, PosLinea + 1.4), NEGRO, B
Printer.FontSize = 18
Printer.CurrentX = CentrarTextoEncab(Texto1, 0.5, MargenDer)
Printer.CurrentY = PosLinea
Printer.Print Texto1
Printer.FontSize = 10
Printer.CurrentX = CentrarTextoEncab(Texto2, 0.5, MargenDer)
Printer.CurrentY = PosLinea + 0.6
Printer.Print Texto2
'Dibujo = RutaSistema & "\FORMATOS\LIBMAYOR.WMF"
'PrinterPaint Dibujo, 0.5, PosLinea, 19.5, 1
PosLinea = PosLinea + 1.1
Printer.FontSize = PorteLetra
Printer.FontName = LetraAnterior
Printer.FontBold = False
End Sub

Public Sub EncabMayor1(Cta As String, _
                       CtaSup As String, _
                       DtaSaldos As Adodc)
Dim SaldoAnt As Currency
Dim SaldoAntME As Currency
'DebeAct As Currency , HaberAct As Currency , SaldoAct As Currency
'Iniciamos la impresion
PorteLetra = Printer.FontSize
LetraAnterior = Printer.FontName
Printer.FontName = TipoArialNarrow
Printer.FontBold = True: Printer.FontItalic = False: Printer.FontSize = 9
With DtaSaldos.Recordset
TipoCta = MidStrg(.fields("Cta"), 1, 1)
Select Case TipoCta
  Case "1", "5"
       'SaldoAnt = SaldoAct - DebeAct + HaberAct
       SaldoAnt = .fields("Saldo") - .fields("Debe") + .fields("Haber")
       SaldoAntME = .fields("Saldo_ME") - .fields("Parcial_ME")
  Case "2", "3", "4"
       'SaldoAnt = SaldoAct - HaberAct + DebeAct
       SaldoAnt = .fields("Saldo") - .fields("Haber") + .fields("Debe")
       SaldoAntME = .fields("Saldo_ME") - .fields("Parcial_ME")
End Select
End With
PrinterTexto Ancho(0), PosLinea + 0.1, "CUENTA:"
PrinterTexto Ancho(4) - 1, PosLinea + 0.1, "GRUPO:"
PrinterTexto Ancho(4) - 2.1, PosLinea + 0.5, "Saldo Anterior S/."
Printer.FontBold = False
PrinterTexto Ancho(0) + 1.2, PosLinea + 0.1, Cta
PrinterTexto Ancho(4), PosLinea + 0.1, CtaSup
PrinterVariables Ancho(4) - 0.1, PosLinea + 0.5, SaldoAntME
PrinterVariables Ancho(7) - 0.1, PosLinea + 0.5, SaldoAnt
Printer.FontBold = True
Printer.FontUnderline = True
PrinterTexto Ancho(0), PosLinea + 0.95, "FECHA"
PrinterTexto Ancho(1), PosLinea + 0.95, "TD"
PrinterTexto Ancho(2), PosLinea + 0.95, "NUMERO"
PrinterTexto Ancho(3), PosLinea + 0.95, "C O N C E P T O"
If OpcCoop = False Then
   'Printer.Line (Ancho(4), PosLinea + 1)-(Ancho(4), PosLinea + 1.5), NEGRO
   PrinterVariables Ancho(4), PosLinea + 0.95, "PARCIAL M/E"
End If
PrinterTexto Ancho(5), PosLinea + 0.95, "D E B E"
PrinterTexto Ancho(6), PosLinea + 0.95, "H A B E R"
PrinterTexto Ancho(7), PosLinea + 0.95, "S A L D O"
Printer.Line (InicioX, PosLinea)-(Ancho(8), PosLinea + 1.4), Negro, B
Printer.Line (InicioX, PosLinea + 0.9)-(Ancho(8), PosLinea + 0.9), Negro
Printer.FontBold = False
Printer.FontUnderline = False
PosLinea = PosLinea + 1.5
Printer.FontSize = PorteLetra
Printer.FontName = LetraAnterior
End Sub

Public Sub EncabMayorM(Cta As String, _
                       CtaSup As String, _
                       DtaTrans As Adodc, _
                       Optional OpcMayor As Boolean)
' DebeAct As Currency , HaberAct As Currency , SaldoAct As Currency
Dim SaldoAnt As Currency
Dim SaldoAnt_ME As Currency
'Iniciamos la impresion
PorteLetra = Printer.FontSize
LetraAnterior = Printer.FontName
Printer.FontName = TipoTimes
If OpcMayor = True Then
   Dibujo = RutaSistema & "\FORMATOS\MAYOR2.GIF"
   PrinterPaint Dibujo, 0.5, PosLinea, 26, 1.5
Else
   Dibujo = RutaSistema & "\FORMATOS\MAYOR.GIF"
   PrinterPaint Dibujo, 0.5, PosLinea, 19.5, 1.5
End If
Printer.FontBold = True: Printer.FontSize = 9
With DtaTrans.Recordset
TipoCta = MidStrg(.fields("Cta"), 1, 1)
Select Case TipoCta
  Case "1", "5"
       SaldoAnt = .fields("Saldo") - .fields("Parcial_MN")
       SaldoAnt_ME = .fields("Saldo_ME") - .fields("Parcial_ME")
  Case "2", "3", "4"
       SaldoAnt = .fields("Saldo") + .fields("Parcial_MN")
       SaldoAnt_ME = .fields("Saldo_ME") + .fields("Parcial_ME")
End Select
If .fields("Parcial_ME") = 0 Then SaldoAnt_ME = 0
End With
PrinterTexto Ancho(4) + 7.5, PosLinea + 0.5, "Saldo Anterior:"
Printer.FontBold = False
PrinterTexto 2.4, PosLinea + 0.1, Cta
PrinterTexto 11.8, PosLinea + 0.1, CtaSup
PrinterVariables Ancho(5), PosLinea + 0.5, SaldoAnt_ME
PrinterVariables Ancho(8), PosLinea + 0.5, SaldoAnt
PosLinea = PosLinea + 1.6
Printer.FontSize = PorteLetra
Printer.FontName = LetraAnterior
End Sub

Public Sub EncabMayorMA(Cta As String, _
                        CtaSup As String, _
                        DtaTrans As Adodc, _
                        Optional OpcMayor As Boolean)
' DebeAct As Currency , HaberAct As Currency , SaldoAct As Currency
Dim SaldoAnt As Currency
Dim SaldoAnt_ME As Currency
'Iniciamos la impresion
PorteLetra = Printer.FontSize
LetraAnterior = Printer.FontName
Printer.FontName = TipoArialNarrow
Printer.FontBold = True
Printer.FontItalic = False
Printer.FontSize = 8
With DtaTrans.Recordset
If Not .EOF Then
TipoCta = MidStrg(.fields("Cta"), 1, 1)
'MsgBox Cta & vbCrLf & CtaSup
Select Case TipoCta
  Case "1", "5"
       SaldoAnt = .fields("Saldo_MN") - .fields("Debitos") + .fields("Creditos")
       'SaldoAnt_ME = .Fields("Saldo_ME") + .Fields("Parcial_ME")
  Case "2", "3", "4"
       SaldoAnt = .fields("Saldo_MN") - .fields("Creditos") + .fields("Debitos")
       'SaldoAnt_ME = .Fields("Saldo_ME") - .Fields("Parcial_ME")
End Select
'If .Fields("Parcial_ME") = 0 Then SaldoAnt_ME = 0
End If
End With
PrinterTexto Ancho(0), PosLinea + 0.1, "CUENTA:"
PrinterTexto Ancho(5) - 1.1, PosLinea + 0.1, "GRUPO:"
PrinterTexto Ancho(0), PosLinea + 0.5, "Mayor de Submódulo:"
PrinterTexto Ancho(5), PosLinea + 0.5, "Saldo Anterior S/."
Printer.FontBold = False
PrinterTexto Ancho(0) + 1.2, PosLinea + 0.1, Cta
PrinterTexto Ancho(5), PosLinea + 0.1, CtaSup
PrinterTexto Ancho(0) + 2.6, PosLinea + 0.5, TipoDoc
PrinterVariables Ancho(8) - 0.1, PosLinea + 0.5, SaldoAnt
Printer.FontBold = True
Printer.FontUnderline = True
PrinterTexto Ancho(0), PosLinea + 0.95, "FECHA"
PrinterTexto Ancho(1), PosLinea + 0.95, "FACTURA"
PrinterTexto Ancho(2), PosLinea + 0.95, "TD"
PrinterTexto Ancho(3), PosLinea + 0.95, "NUMERO"
PrinterTexto Ancho(4), PosLinea + 0.95, "C O N C E P T O"
PrinterTexto Ancho(5), PosLinea + 0.95, "PARCIAL M/E"
PrinterTexto Ancho(6), PosLinea + 0.95, "D E B I T O S"
PrinterTexto Ancho(7), PosLinea + 0.95, "C R E D I T O S"
PrinterTexto Ancho(8), PosLinea + 0.95, "S A L D O"
Printer.Line (InicioX, PosLinea)-(Ancho(9), PosLinea + 1.4), Negro, B
Printer.Line (InicioX, PosLinea + 0.9)-(Ancho(9), PosLinea + 0.9), Negro
Printer.FontBold = False
Printer.FontUnderline = False
PosLinea = PosLinea + 1.5
Printer.FontSize = PorteLetra
Printer.FontName = LetraAnterior
End Sub

Public Sub EncabLibroBanco1(Cta As String, _
                            CtaSup As String, _
                            DtaLibBanc As Adodc)
Dim SaldoAnt As Currency
Dim SaldoAntME As Currency
'Iniciamos la impresion
PorteLetra = Printer.FontSize
LetraAnterior = Printer.FontName
Printer.FontName = TipoTimes
Printer.FontBold = True
Printer.FontItalic = False
SaldoAnt = 0: SaldoAntME = 0
With DtaLibBanc.Recordset
     SaldoAnt = .fields("Saldo") - .fields("Debe") + .fields("Haber")
     SaldoAntME = .fields("Saldo_ME") - .fields("Parcial_ME")
End With
PosLinea = PosLinea - 0.5
Printer.FontSize = 20
Printer.CurrentX = CentrarTextoEncab("L I B R O    B A N C O", 0.5, Ancho(10))
Printer.CurrentY = PosLinea
Printer.Print "L I B R O    B A N C O"
PosLinea = PosLinea + 0.8
Printer.FontSize = 9
Printer.FontName = TipoArialNarrow
PrinterTexto 0.5, PosLinea + 0.1, "CUENTA SUPERIOR:"
PrinterTexto 0.5, PosLinea + 0.5, "CUENTA ACTUAL:"
PrinterTexto 15, PosLinea + 0.5, "Saldo US$."
PrinterTexto 21.5, PosLinea + 0.5, "Saldo S/."
Printer.FontBold = False

PrinterTexto 4, PosLinea + 0.1, CtaSup
PrinterTexto 4, PosLinea + 0.5, Cta

PrinterVariables Ancho(6), PosLinea + 0.5, SaldoAntME
PrinterVariables Ancho(9), PosLinea + 0.5, SaldoAnt

'''PrinterVariables Ancho(6), PosLinea + 0.5, Suma_ME
'''PrinterVariables Ancho(9), PosLinea + 0.5, SumaSaldo

          
Printer.FontBold = True
Printer.FontUnderline = True
PrinterTexto Ancho(0), PosLinea + 0.95, "FECHA"
PrinterTexto Ancho(1), PosLinea + 0.95, "TD"
PrinterTexto Ancho(2), PosLinea + 0.95, "NUMERO"
PrinterTexto Ancho(3), PosLinea + 0.95, "CHEQ/DEP"
PrinterTexto Ancho(4), PosLinea + 0.95, "BENEFICIARIO"
PrinterTexto Ancho(5), PosLinea + 0.95, "C O N C E P T O"
If OpcCoop = False Then PrinterVariables Ancho(6), PosLinea + 0.95, "PARCIAL M/E"
PrinterTexto Ancho(7), PosLinea + 0.95, "D E B E"
PrinterTexto Ancho(8), PosLinea + 0.95, "H A B E R"
PrinterTexto Ancho(9), PosLinea + 0.95, "S A L D O"
Printer.Line (InicioX, PosLinea)-(Ancho(10), PosLinea + 1.4), Negro, B
Printer.Line (InicioX, PosLinea + 0.9)-(Ancho(10), PosLinea + 0.9), Negro
Printer.FontBold = False
Printer.FontUnderline = False
PosLinea = PosLinea + 1.5
Printer.FontSize = PorteLetra
Printer.FontName = LetraAnterior
End Sub

Public Sub Encab_Diario_General(Datas As Adodc)
'Iniciamos la impresion
PorteLetra = Printer.FontSize
LetraAnterior = Printer.FontName
Printer.Line (Ancho(0), PosLinea)-(Ancho(5), PosLinea + 1), Negro, B
Printer.FontName = TipoTimes
Printer.FontBold = True: Printer.FontSize = 18
msg = "L I B R O     D I A R I O"
PrinterTexto CentrarTexto(msg), PosLinea + 0.1, msg
Printer.FontBold = False
PosLinea = PosLinea + 1
Printer.FontSize = PorteLetra
Printer.FontName = LetraAnterior
Encab_Libro_Diario Datas
End Sub

Public Sub InsertarAsientoBanco(DtaAsientoBanco As Adodc, _
                                Total_Cheq_Dep As Currency)
With DtaAsientoBanco.Recordset
 If .RecordCount < 5 Then
    .AddNew
     If Moneda_US Then Total_Cheq_Dep = Total_Cheq_Dep * Dolar
    .fields("ME") = Moneda_US
    .fields("CTA_BANCO") = CuentaBanco
    .fields("BANCO") = NombreBanco
    .fields("CHEQ_DEP") = NoCheque
    .fields("EFECTIVIZAR") = Fecha_Vence
    .fields("VALOR") = Total_Cheq_Dep
    .fields("T_No") = Trans_No
    .fields("Item") = NumEmpresa
    .fields("CodigoU") = CodigoUsuario
    .Update
 End If
End With
End Sub

Public Sub InsertarTotales(TB As String, _
                           cod As String, _
                           CtaDG As String, _
                           Cta As String, _
                           Tot As Currency, _
                           Tot_ME As Currency)
Dim AdoCon1 As ADODB.Connection
Dim SQLTabla As String
Dim Nivel_Cta As Byte
  If Dolar = 0 Then Tot_ME = 0
  Ln_No = Ln_No + 1
  Nivel_Cta = Niveles(CStr(cod))
  SQLTabla = "INSERT INTO Balances_Mes " _
           & "(Codigo,Cuenta,Total_N1,Total_N2,Total_N3,Total_N4,Total_N5,Total_N6,DG,Ln,TB,Item,TC) " _
           & "VALUES ('" & cod & "','" & Cta & "'"
  Select Case Nivel_Cta
    Case 1: SQLTabla = SQLTabla & "," & Tot & ",0,0,0,0,0 "
    Case 2: SQLTabla = SQLTabla & ",0," & Tot & ",0,0,0,0 "
    Case 3: SQLTabla = SQLTabla & ",0,0," & Tot & ",0,0,0 "
    Case 4: SQLTabla = SQLTabla & ",0,0,0," & Tot & ",0,0 "
    Case 5: SQLTabla = SQLTabla & ",0,0,0,0," & Tot & ",0 "
    Case 6: SQLTabla = SQLTabla & ",0,0,0,0,0," & Tot & " "
  End Select
  SQLTabla = SQLTabla & ",'" & CtaDG & "' " _
           & "," & Ln_No & " " _
           & ",'" & TB & "' " _
           & ",'" & NumEmpresa & "','" & TipoDoc & "');"
  'MsgBox SQLTabla
  Set AdoCon1 = New ADODB.Connection
  AdoCon1.open AdoStrCnn
  AdoCon1.Execute SQLTabla, , adCmdText
  AdoCon1.Close
End Sub

'Variables Generales de Entrada
'CodigoCli
'ValorDH
'Dolar
'OpcTM
'OpcDH
'Codigo
'Cuenta
'Fecha_Vence
'NoCheque
'Trans_No
'DetalleComp
Public Sub InsertarAsiento(DtaAsiento As Adodc)
Dim AdoRegSC As ADODB.Recordset
Dim ValorDHAux As Currency
Dim InsertarCta As Boolean
Dim Ln_No_A As Integer

  InsertarCta = True
  Ln_No_A = 0
  If IsEmpty(CodigoCli) Then CodigoCli = Ninguno
  If IsNull(CodigoCli) Then CodigoCli = Ninguno
  If NoCheque = Ninguno Then CodigoCli = Ninguno
  
  ValorDHAux = Redondear(ValorDH, 2)
  'MsgBox ValorDHAux
  If Codigo <> Ninguno Then
     Debe = 0: Haber = 0
     'And Moneda_US = False Then ValorDH = Redondear(ValorDH * Dolar,2)
     If OpcTM = 2 Or Moneda_US Then
        If Opcion_Mulp Then
           ValorDH = Val(ValorDH * Dolar)
        Else
           If Dolar <= 0 Then
              MsgBox "No se puede Dividir para cero," & vbCrLf & "cambie la Cotización."
              ValorDH = 0
           Else
              ValorDH = Val(ValorDH / Dolar)
           End If
        End If
     End If
     Select Case OpcDH
       Case 1: Debe = ValorDH
       Case 2: Haber = ValorDH
     End Select
     If ValorDH <> 0 And Cuenta <> Ninguno Then
        Select Case SubCta
          Case "C", "P", "G", "I", "CP", "PM", "CC"
               sSQL = "SELECT * " _
                    & "FROM Asiento " _
                    & "WHERE TC = '" & SubCta & "' " _
                    & "AND CODIGO = '" & Codigo & "' " _
                    & "AND T_No = " & Trans_No & " " _
                    & "AND Item = '" & NumEmpresa & "' " _
                    & "AND CodigoU = '" & CodigoUsuario & "' "
               Select Case OpcDH
                 Case 1: sSQL = sSQL & "AND DEBE > 0 "
                 Case 2: sSQL = sSQL & "AND HABER > 0 "
               End Select
               Select_AdoDB AdoRegSC, sSQL
               If AdoRegSC.RecordCount > 0 Then
                  InsertarCta = False
                  Ln_No_A = AdoRegSC.fields("A_No")
               End If
               AdoRegSC.Close
        End Select
        
        With DtaAsiento.Recordset
             If InsertarCta Then
               .AddNew
             Else
               .MoveFirst
               .Find ("A_No = " & Ln_No_A & " ")
                If .EOF Then .AddNew
             End If
            .fields("PARCIAL_ME") = 0
            .fields("ME") = False
            .fields("CODIGO") = Codigo
            .fields("CUENTA") = Cuenta
            .fields("DETALLE") = TrimStrg(MidStrg(DetalleComp, 1, 60))
             If OpcCoop Then
                If Moneda_US Then
                   Debe = Redondear(Debe / Dolar, 2)
                   Haber = Redondear(Haber / Dolar, 2)
                Else
                   Debe = Redondear(Debe, 2)
                   Haber = Redondear(Haber, 2)
                End If
             Else
               .fields("PARCIAL_ME") = 0
                If Moneda_US Or OpcTM = 2 Then
                   If (Debe - Haber) < 0 Then ValorDHAux = -ValorDHAux
                  .fields("PARCIAL_ME") = ValorDHAux
                  .fields("ME") = True
                End If
                Debe = Redondear(Debe, 2)
                Haber = Redondear(Haber, 2)
             End If
            .fields("DEBE") = Debe
            .fields("HABER") = Haber
            .fields("EFECTIVIZAR") = Fecha_Vence
            .fields("CHEQ_DEP") = NoCheque
            .fields("CODIGO_C") = CodigoCli
            .fields("CODIGO_CC") = CodigoCC
            .fields("T_No") = Trans_No
            .fields("Item") = NumEmpresa
            .fields("CodigoU") = CodigoUsuario
            .fields("TC") = SubCta
             If InsertarCta Then
               .fields("A_No") = Ln_No
                Ln_No = Ln_No + 1
             End If
            .Update
             
        End With
     End If
  End If
End Sub

Public Sub InsertarAsientos(Datas As Adodc, _
                            CodCta As String, _
                            Parcial_MEs As Currency, _
                            Debes As Currency, _
                            Habers As Currency)
Dim AdoReg As ADODB.Recordset
Dim InsAsiento As Boolean
Dim Cuenta As String

' Cadena de Conección a la base de datos
  InsAsiento = False
  If IsEmpty(CodigoCli) Then CodigoCli = Ninguno
  If IsNull(CodigoCli) Then CodigoCli = Ninguno
  
  If Debes > 0 Or Habers > 0 And CodCta <> "" Then
     If CodCta <> "0" Then
        Set AdoReg = New ADODB.Recordset
        AdoReg.CursorType = adOpenStatic
        AdoReg.CursorLocation = adUseClient
        SQL1 = "SELECT TC, Codigo, Cuenta " _
             & "FROM Catalogo_Cuentas " _
             & "WHERE Codigo = '" & CodCta & "' " _
             & "AND Item = '" & NumEmpresa & "' " _
             & "AND Periodo = '" & Periodo_Contable & "' "
        AdoReg.open SQL1, AdoStrCnn, , , adCmdText
        If AdoReg.RecordCount > 0 Then
           InsAsiento = True
           Cuenta = AdoReg.fields("Cuenta")
           SubCta = AdoReg.fields("TC")
        End If
        AdoReg.Close
        
        If Not InsAsiento And Len(CodCta) > 2 Then
           SQL1 = "SELECT  TC, Codigo, Cuenta  " _
                & "FROM Catalogo_Cuentas " _
                & "WHERE Codigo_Ext LIKE '%" & CodCta & "' " _
                & "AND Item = '" & NumEmpresa & "' " _
                & "AND Periodo = '" & Periodo_Contable & "' "
           AdoReg.open SQL1, AdoStrCnn, , , adCmdText
           If AdoReg.RecordCount > 0 Then
              InsAsiento = True
              CodCta = AdoReg.fields("Codigo")
              Cuenta = AdoReg.fields("Cuenta")
              SubCta = AdoReg.fields("TC")
           End If
           AdoReg.Close
        End If
       
       'MsgBox .RecordCount & vbCrLf & Cod & vbCrLf & Debes & vbCrLf & Habers
        If InsAsiento Then
           SetAddNew Datas
           SetFields Datas, "CODIGO", CodCta
           SetFields Datas, "CUENTA", Cuenta
           SetFields Datas, "DETALLE", DetalleComp
           If OpcCoop = False Then SetFields Datas, "PARCIAL_ME", Redondear(Parcial_MEs, 2)
           SetFields Datas, "DEBE", Redondear(Debes, 2)
           SetFields Datas, "HABER", Redondear(Habers, 2)
           SetFields Datas, "Item", NumEmpresa
           SetFields Datas, "T_No", Trans_No
           SetFields Datas, "ME", False
           SetFields Datas, "EFECTIVIZAR", Fecha_Vence
           SetFields Datas, "CODIGO_C", CodigoCli
           SetFields Datas, "CODIGO_CC", CodigoCC
           SetFields Datas, "CHEQ_DEP", NoCheque
           SetFields Datas, "CodigoU", CodigoUsuario
           SetFields Datas, "A_No", Ln_No
           SetFields Datas, "TC", SubCta
           SetUpdate Datas
           Ln_No = Ln_No + 1
   '      Else
   '          Cadena = "El Codigo de Cuenta: " & Cod & ", " & Abs(Debes - Habers) & vbCrLf & "No existe en el Catalogo."
   '          If ConceptoComp <> Ninguno Then Cadena = Cadena & vbCrLf & ConceptoComp
   '          Debes = 0: Habers = 0
        End If
     End If
  End If
End Sub

Public Function Si_Partida_Doble() As Boolean
Dim AdoDBAsiento As ADODB.Recordset
Dim SiCuadra As Boolean
Dim SQL As String
   SiCuadra = False
   SumaDebe = 0
   SumaHaber = 1
   SQL = "SELECT CODIGO, CUENTA, PARCIAL_ME, DEBE, HABER " _
       & "FROM Asiento " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND T_No = " & Trans_No & " "
   Select_AdoDB AdoDBAsiento, SQL
   With AdoDBAsiento
    If .RecordCount > 0 Then
        SumaDebe = 0
        SumaHaber = 0
        Do While Not .EOF
           SumaDebe = SumaDebe + .fields("DEBE")
           SumaHaber = SumaHaber + .fields("HABER")
          .MoveNext
        Loop
        If (SumaDebe - SumaHaber) = 0 Then SiCuadra = True
    End If
   End With
   AdoDBAsiento.Close
  'MsgBox SQL
   Si_Partida_Doble = SiCuadra
End Function

Public Function ImprimirBancos(TipoComp As String, _
                          DtaBanco As ADODB.Recordset, _
                          Item As Byte) As Currency
Dim Total_de_Banco As Currency
Total_de_Banco = 0
With DtaBanco
 If .RecordCount > 0 Then
    .MoveFirst
     'MsgBox .RecordCount & vbCrLf & Item
     IR = SetD(Item).PosY
     Printer.FontSize = SetD(Item).Tamaño
     If .RecordCount > 2 Then
         Cadena = .fields("Cuenta")
         Do While Not .EOF
            Cadena = .fields("Cuenta")
            Select Case TipoComp
              Case CompIngreso
                   Total_de_Banco = Total_de_Banco + .fields("Debe")
                   Cadena1 = "Depósitos: En bloque"
              Case CompEgreso
                   Total_de_Banco = Total_de_Banco + .fields("Haber")
                   Cadena1 = "Cheques: En bloque"
            End Select
           .MoveNext
         Loop
         PrinterVariables SetD(Item - 1).PosX, IR, Total_de_Banco
         'PrinterTexto SetD(Item + 2).PosX - 2, IR, Cadena1
         Printer.FontBold = False
         PrinterTexto SetD(Item).PosX, IR, Cadena
         PrinterTexto SetD(Item + 2).PosX - 2.1, IR, Cadena1 '" En bloque"
     Else
         Do While Not .EOF
            Printer.FontBold = True
            Select Case TipoComp
              Case CompIngreso
                   If .fields("Debe") > 0 Then
                       Total_de_Banco = Total_de_Banco + .fields("Debe")
                       PrinterTexto SetD(Item + 2).PosX - 2.1, IR, "Depósito No."
                       Printer.FontBold = False
                       'PrinterTexto SetD(Item - 1).PosX, IR, .Fields("Debe")
                       PrinterTexto SetD(Item).PosX, IR, .fields("Cuenta")
                       PrinterFields SetD(Item + 2).PosX, IR, .fields("Cheq_Dep")
                       IR = IR + 0.4
                   End If
              Case CompEgreso
                   If .fields("Haber") > 0 Then
                       Total_de_Banco = Total_de_Banco + .fields("Haber")
                       PrinterTexto SetD(Item + 2).PosX - 2.1, IR, "Cheque No."
                       Printer.FontBold = False
                       'PrinterTexto SetD(Item - 1).PosX, IR, .Fields("Haber")
                       PrinterTexto SetD(Item).PosX, IR, .fields("Cuenta")
                       PrinterFields SetD(Item + 2).PosX, IR, .fields("Cheq_Dep")
                       IR = IR + 0.4
                   End If
            End Select
           'MsgBox Cadena
           .MoveNext
         Loop
     End If
 End If
End With
ImprimirBancos = Total_de_Banco
End Function

Public Sub Mayorizar_Saldos(TipoCta1 As String)
   If OpcCoop Then
      Select Case MidStrg(TipoCta1, 1, 1)
        Case "4"
             SumaCta = SumaCta + Debe - Haber
             Suma_ME = Suma_ME + Debe_ME - Haber_ME
        Case "5"
             SumaCta = SumaCta + Haber - Debe
             Suma_ME = Suma_ME + Haber_ME - Debe_ME
      End Select
   Else
      Select Case MidStrg(TipoCta1, 1, 1)
        Case "4"
             SumaCta = SumaCta + Haber - Debe
             Suma_ME = Suma_ME + Haber_ME - Debe_ME
        Case "5"
             SumaCta = SumaCta + Debe - Haber
             Suma_ME = Suma_ME + Debe_ME - Haber_ME
      End Select
   End If
   Select Case MidStrg(TipoCta1, 1, 1)
     Case "1", "6", "8"
          SumaCta = SumaCta + Debe - Haber
          Suma_ME = Suma_ME + Debe_ME - Haber_ME
     Case "2", "3", "7", "9"
          SumaCta = SumaCta + Haber - Debe
          Suma_ME = Suma_ME + Haber_ME - Debe_ME
   End Select
End Sub

Public Sub Imprimir_Catalogo(Datas As Adodc)
Dim SizeLetra As Integer
On Error GoTo Errorhandler
Mensajes = "Seguro de Imprimir en:" & vbCrLf & Printer.DeviceName & "?"
Titulo = "IMPRESION"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
RatonReloj
InicioX = 0.5: InicioY = 0
SizeLetra = 9
SQLMsg1 = "": SQLMsg2 = "": SQLMsg3 = ""
MensajeEncabData = "P L A N    D E    C U E N T A S"
DataAnchoCampos InicioX, Datas, SizeLetra, TipoCourierNew, Orientacion_Pagina
Ancho(0) = 0.5 'Clave
Ancho(1) = 1.7  'TC
Ancho(2) = 2.3  'ME
Ancho(3) = 3    'DG
Ancho(4) = 3.6  'Codigo
Ancho(5) = 6.7  'Cuenta
Ancho(6) = 17   'Presupuesto
Ancho(7) = 20 '
Pagina = 1
'Iniciamos la impresion
EncabezadoData Datas
Printer.FontBold = False
With Datas.Recordset
     .MoveFirst
      Printer.FontName = TipoCourierNew
      Do While Not .EOF
         Select Case Niveles(.fields("Codigo"))
           Case 1: PCol = 0
           Case 2: PCol = 0.3
           Case 3: PCol = 0.6
           Case 4: PCol = 0.9
           Case 5: PCol = 1.2
           Case 6: PCol = 1.5
         End Select
         Printer.FontBold = False
         Printer.FontItalic = False
         Printer.FontSize = SizeLetra
         If .fields("DG") = "G" Then Printer.FontBold = True
         If .fields("Clave") <> 0 Then PrinterTexto Ancho(0), PosLinea, String$(3, " ") & Str(.fields("Clave")), True
         PrinterFields Ancho(1), PosLinea, .fields("TC"), True
         PrinterFields Ancho(2), PosLinea, .fields("ME"), True
         PrinterFields Ancho(3), PosLinea, .fields("DG"), True
         PrinterFields Ancho(4), PosLinea, .fields("Codigo"), True
         If .fields("TC") <> "N" Then Printer.FontItalic = True
         PrinterFields Ancho(5) + PCol, PosLinea, .fields("Cuenta"), False
         Printer.FontItalic = False
         PrinterFields Ancho(6), PosLinea, .fields("Presupuesto"), True
         Printer.Line (Ancho(CantCampos), PosLinea - 0.1)-(Ancho(CantCampos), PosLinea + 0.4), Negro
         PosLinea = PosLinea + 0.36
         If PosLinea >= LimiteAlto Then
            Printer.Line (Ancho(0), PosLinea)-(Ancho(CantCampos), PosLinea), Negro
            Printer.NewPage
            PosLinea = 0
            EncabezadoData Datas
            Printer.FontName = TipoCourierNew
         End If
        .MoveNext
      Loop
End With
MensajeEncabData = ""
Printer.Line (Ancho(0), PosLinea)-(Ancho(CantCampos), PosLinea), Negro
PosLinea = PosLinea + 0.05
Printer.Line (Ancho(0), PosLinea)-(Ancho(CantCampos), PosLinea), Negro
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

Public Sub ImprimirCompDiario(DataComp As ADODB.Recordset, _
                              DataTrans As ADODB.Recordset, _
                              DataFact As ADODB.Recordset, _
                              DataRets As ADODB.Recordset, _
                              DataSubC1 As ADODB.Recordset, _
                              DataSubC2 As ADODB.Recordset, _
                              ImpSoloRet As Boolean, _
                              Optional NuevaPagina As Boolean, _
                              Optional NoImpRet As Boolean)
'Establecemos Espacios y seteos de impresion
On Error GoTo Errorhandler
RatonReloj
LetraAnterior = Printer.FontName
CDConLineas = ProcesarSeteos(CompDiario)
'Escala_Centimetro 1, TipoTimes, 10
Pagina = 1
'Iniciamos la impresion
'MsgBox DataComp.Fields("Fecha")
With DataComp
 If .RecordCount > 0 Then
     Mifecha = .fields("Fecha")
     If ImpSoloRet = False Then
        ImprimirFormatoComprobantes CompDiario, FechaAnio(Mifecha), .fields("Numero"), CDConLineas, SetD(1).PosY
        Mifecha = FechaStrgCorta(.fields("Fecha"))
        Printer.FontBold = False
        Printer.FontSize = SetD(2).Tamaño
        PrinterTexto SetD(2).PosX, SetD(2).PosY, FechaStrgCorta(.fields("Fecha"))
        Printer.FontSize = SetD(3).Tamaño
        PrinterLineas SetD(3).PosX, SetD(3).PosY, .fields("Concepto"), 16
        ImprimirAsientos DataTrans, DataComp, DataSubC1, DataSubC2, 4, 9, CDConLineas, CompDiario, .fields("T"), .fields("Cotizacion")
     End If
     If NoImpRet = False Then
        If DataRets.RecordCount > 0 Then
           Mensajes = "Imprimir Retencion"
           Titulo = "Pregunta de Impresion"
           If BoxMensaje = vbYes Then ImprimirCompRetencion ImpSoloRet, DataFact, DataComp, DataRets
        End If
     End If
     Printer.FontName = LetraAnterior
     If NuevaPagina Then Printer.NewPage Else Printer.EndDoc
 End If
End With
RatonNormal
'MsgBox "El Comprobante se ha imprimido correctamente."
Exit Sub
Errorhandler:
    RatonNormal
    ErrorDeImpresion
    Exit Sub
End Sub

Public Sub ImprimirCompNota_D_C(DataComp As ADODB.Recordset, _
                                DataTrans As ADODB.Recordset, _
                                DataSubC1 As ADODB.Recordset, _
                                DataSubC2 As ADODB.Recordset, _
                                Tipo_D_C As String, _
                                Optional NuevaPagina As Boolean)
'Establecemos Espacios y seteos de impresion
On Error GoTo Errorhandler
RatonReloj
LetraAnterior = Printer.FontName
CDConLineas = ProcesarSeteos("DC")
'Escala_Centimetro 1, TipoTimes, 10
Pagina = 1
'Iniciamos la impresion
'MsgBox DataComp.Fields("Fecha")
With DataComp
 If .RecordCount > 0 Then
     Mifecha = .fields("Fecha")
     If Tipo_D_C = "ND" Then
        ImprimirFormatoComprobantes CompNotaDebito, FechaAnio(Mifecha), .fields("Numero"), CDConLineas, SetD(1).PosY
     Else
        ImprimirFormatoComprobantes CompNotaCredito, FechaAnio(Mifecha), .fields("Numero"), CDConLineas, SetD(1).PosY
     End If
     Mifecha = FechaStrgCorta(.fields("Fecha"))
     Printer.FontBold = False
     Printer.FontSize = SetD(2).Tamaño
     PrinterTexto SetD(2).PosX, SetD(2).PosY, FechaStrgCorta(.fields("Fecha"))
     Printer.FontSize = SetD(3).Tamaño
     PrinterLineas SetD(3).PosX, SetD(3).PosY, .fields("Concepto"), 16
     ImprimirAsientos DataTrans, DataComp, DataSubC1, DataSubC2, 4, 9, CDConLineas, "DC", .fields("T"), .fields("Cotizacion")
     Printer.FontName = LetraAnterior
     If NuevaPagina Then Printer.NewPage Else Printer.EndDoc
 End If
End With
RatonNormal
'MsgBox "El Comprobante se ha imprimido correctamente."
Exit Sub
Errorhandler:
    RatonNormal
    ErrorDeImpresion
    Exit Sub
End Sub

Public Sub Imprimir_Conciliacion(DtaNotaDC As Adodc, _
                                 DtaDebCred As Adodc, _
                                 DtaTransito As Adodc, _
                                 TipoCta As Boolean)
Dim Supervisor As String

Dim SizeLetra As Integer

Dim SizeLetra1 As Single
Dim LenT As Single

On Error GoTo Errorhandler
Mensajes = "Seguro de Imprimir en:" & vbCrLf & Printer.DeviceName & "?"
Titulo = "IMPRESION"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then

RatonReloj
Supervisor = Leer_Campo_Empresa("Supervisor")
NombreContador = Leer_Campo_Empresa("Contador")
'FormaImp = 1:
SizeLetra = 8: SizeLetra1 = 10
InicioX = 0.5: InicioY = 0
Escala_Centimetro 1, TipoTimes, SizeLetra
Pagina = 1
'Iniciamos la impresion
RatonReloj
Printer.FontBold = False
Encabezado 0.5, 19.5
     Printer.FontBold = True
     With Heads
          Printer.FontSize = 16
      If .MsgTitulo <> "" Then
          PrinterTexto CentrarTexto(.MsgTitulo), PosLinea, .MsgTitulo
          PosLinea = PosLinea + 0.8
      End If
      Printer.FontSize = 10
      If .MsgObjetivo <> "" And .TextoObjetivo <> "" Then
          PrinterTexto 0.5 + 0.3, PosLinea, .MsgObjetivo
          LenT = Printer.TextWidth(.MsgObjetivo)
          PrinterTexto 0.5 + LenT + 0.4, PosLinea, .TextoObjetivo
          PosLinea = PosLinea + 0.5
      End If
      If .MsgConcepto <> "" And .TextoConcepto <> "" Then
          PrinterTexto 0.5 + 0.3, PosLinea, .MsgConcepto
          LenT = Printer.TextWidth(.MsgConcepto)
          PrinterTexto 0.5 + LenT + 0.4, PosLinea, .TextoConcepto
          PosLinea = PosLinea + 0.6
      End If
     End With

With DtaNotaDC.Recordset
 If .RecordCount > 0 Then
    .MoveFirst
     Printer.FontSize = SizeLetra1
     Printer.FontUnderline = True
     PrinterTexto 0.5, PosLinea, "SALDO SEGUN LIBRO BANCOS:"
     Printer.FontUnderline = False
     PrinterVariables 13.5, PosLinea, Saldo
     PosLinea = PosLinea + 0.5
     Printer.FontUnderline = True
     PrinterTexto 0.5, PosLinea, "NOTAS DE DEBITOS (no contabilizadas):"
     PosLinea = PosLinea + 0.5
     Printer.FontSize = SizeLetra
     PrinterTexto 0.5, PosLinea, "F e c h a"
     PrinterTexto 2.5, PosLinea, "Concepto"
     PrinterTexto 12, PosLinea, "TP"
     PrinterTexto 13, PosLinea, "Documento"
     'PrinterTexto 15, PosLinea, "Parcial_ME"
     PrinterTexto 17, PosLinea, "D E B I T O S"
     PosLinea = PosLinea + 0.4
     Printer.FontUnderline = False
     Printer.FontBold = False
     Do While Not .EOF
        If .fields("Debitos") <> 0 Then
            PrinterFields 0.5, PosLinea, .fields("Fecha"), False
            PrinterFields 2.5, PosLinea, .fields("Concepto"), False
            PrinterFields 12, PosLinea, .fields("TP"), False
            PrinterFields 13, PosLinea, .fields("Documento"), False
            PrinterFields 17, PosLinea, .fields("Debitos"), False
            PosLinea = PosLinea + 0.4
            If PosLinea > LimiteAlto Then
               Printer.Line (0.5, PosLinea)-(19.5, PosLinea), Negro
               Printer.NewPage
               Encabezado 0.5, 19.5
               Printer.FontSize = SizeLetra
            End If
        End If
       .MoveNext
     Loop
     Printer.Line (InicioX, PosLinea)-(19.5, PosLinea), Negro
     PosLinea = PosLinea + 0.1
     PrinterTexto 14.5, PosLinea, "TOTAL N/D"
     PrinterVariables 17, PosLinea, SumaDebe
     PosLinea = PosLinea + 0.4
    .MoveFirst
     Printer.FontSize = SizeLetra1
     Printer.FontBold = True
     Printer.FontUnderline = True
     PrinterTexto 0.5, PosLinea, "NOTAS DE CREDITOS (no contabilizadas):"
     PosLinea = PosLinea + 0.5
     Printer.FontSize = SizeLetra
     PrinterTexto 0.5, PosLinea, "F e c h a"
     PrinterTexto 2.5, PosLinea, "Concepto"
     PrinterTexto 12, PosLinea, "TP"
     PrinterTexto 13, PosLinea, "Documento"
     'PrinterTexto 15, PosLinea, "Parcial_ME"
     PrinterTexto 17, PosLinea, "C R E D I T O S"
     Printer.FontUnderline = False
     PosLinea = PosLinea + 0.4
     Printer.FontBold = False
     Printer.FontSize = SizeLetra
     Do While Not .EOF
        If .fields("Creditos") <> 0 Then
            PrinterFields 0.5, PosLinea, .fields("Fecha"), False
            PrinterFields 2.5, PosLinea, .fields("Concepto"), False
            PrinterFields 12, PosLinea, .fields("TP"), False
            PrinterFields 13, PosLinea, .fields("Documento"), False
            PrinterFields 17, PosLinea, .fields("Creditos"), False
            PosLinea = PosLinea + 0.4
            If PosLinea > LimiteAlto Then
               Printer.Line (0.5, PosLinea)-(19.5, PosLinea), Negro
               Printer.NewPage
               Encabezado 0.5, 19.5
               Printer.FontSize = SizeLetra
            End If
        End If
       .MoveNext
     Loop
     Printer.Line (InicioX, PosLinea)-(19.5, PosLinea), Negro
     PosLinea = PosLinea + 0.1
     PrinterTexto 14.5, PosLinea, "TOTAL N/C"
     PrinterVariables 17, PosLinea, SumaHaber
     PosLinea = PosLinea + 0.4
     Printer.FontSize = SizeLetra1
     Printer.FontBold = True
     Printer.FontUnderline = True
     PrinterTexto 0.5, PosLinea, "SALDO REAL CONCILIADO DE LIBRO BANCOS:"
     Printer.FontUnderline = False
     PrinterVariables 16.5, PosLinea, Saldo + SumaDebe - SumaHaber
     Printer.FontBold = False
     PosLinea = PosLinea + 1
 Else
     Printer.FontSize = SizeLetra1
     Printer.FontUnderline = True
     PrinterTexto 0.5, PosLinea, "SALDO SEGUN LIBRO BANCOS:"
     Printer.FontUnderline = False
     PrinterVariables 13.5, PosLinea, Saldo
     PosLinea = PosLinea + 0.5
     
     Printer.FontBold = True
     Printer.FontUnderline = True
     PrinterTexto 0.5, PosLinea, "SALDO REAL CONCILIADO DE LIBRO BANCOS:"
     Printer.FontUnderline = False
     PrinterVariables 16.5, PosLinea, Saldo + SumaDebe - SumaHaber
     Printer.FontBold = False
 End If
End With

With DtaDebCred.Recordset
 If .RecordCount > 0 Then
     If PosLinea > LimiteAlto Then
        Printer.Line (0.5, PosLinea)-(19.5, PosLinea), Negro
        Printer.NewPage
        Encabezado 0.5, 19.5
        Printer.FontSize = SizeLetra
     End If
     Printer.FontSize = SizeLetra1
     Printer.FontBold = True
     Printer.FontUnderline = True
     PrinterTexto 0.5, PosLinea, "SALDO SEGUN ESTADO DE CUENTA:"
     Printer.FontUnderline = False
     PrinterVariables 13.5, PosLinea, Total_Bancos
     PosLinea = PosLinea + 0.5
     Printer.FontUnderline = True
     PrinterTexto 0.5, PosLinea, "DEPOSITOS EN TRANSITO (no registrados por el banco o la compañía):"
     PosLinea = PosLinea + 0.5
     Printer.FontSize = SizeLetra
     PrinterTexto 0.5, PosLinea, "F e c h a"
     PrinterTexto 2.5, PosLinea, "Beneficiario"
     PrinterTexto 10, PosLinea, "TP"
     PrinterTexto 11, PosLinea, "Numero"
     PrinterTexto 12.5, PosLinea, "Depósito No."
     PrinterTexto 14.5, PosLinea, "Parcial_ME"
     PrinterTexto 17, PosLinea, "V A L O R"
     PosLinea = PosLinea + 0.4
     Printer.FontUnderline = False
     Printer.FontBold = False
    .MoveFirst
     Printer.FontSize = SizeLetra
     Do While Not .EOF
        If .fields("Debe") <> 0 Then
            PrinterFields 0.5, PosLinea, .fields("Fecha"), False
            PrinterFields 2.5, PosLinea, .fields("Cliente"), False
            PrinterFields 10, PosLinea, .fields("TP"), False
            PrinterFields 11, PosLinea, .fields("Numero"), False
            PrinterFields 12.5, PosLinea, .fields("Cheq_Dep"), False
            PrinterFields 17, PosLinea, .fields("Debe"), False
            PosLinea = PosLinea + 0.4
            If PosLinea > LimiteAlto Then
               Printer.Line (0.5, PosLinea)-(19.5, PosLinea), Negro
               Printer.NewPage
               Encabezado 0.5, 19.5
               Printer.FontSize = SizeLetra
            End If
        End If
       .MoveNext
     Loop
     Printer.Line (InicioX, PosLinea)-(19.5, PosLinea), Negro
     PosLinea = PosLinea + 0.1
     PrinterTexto 14.5, PosLinea, "TOTAL DEPOSITO"
     PrinterVariables 17, PosLinea, Debe
     PosLinea = PosLinea + 0.4
     Printer.FontSize = SizeLetra1
     Printer.FontBold = True
     Printer.FontUnderline = True
     PrinterTexto 0.5, PosLinea, "CHEQUES GIRADOS Y NO COBRADOS:"
     PosLinea = PosLinea + 0.5
     Printer.FontSize = SizeLetra
     PrinterTexto 0.5, PosLinea, "F e c h a"
     PrinterTexto 2.5, PosLinea, "Beneficiario"
     PrinterTexto 10, PosLinea, "TP"
     PrinterTexto 11, PosLinea, "Numero"
     PrinterTexto 12.5, PosLinea, "Cheque No."
     PrinterTexto 14.5, PosLinea, "Parcial_ME"
     PrinterTexto 17, PosLinea, "V A L O R"
     PosLinea = PosLinea + 0.4
     Printer.FontUnderline = False
     Printer.FontBold = False
    .MoveFirst
     Printer.FontSize = SizeLetra
     Do While Not .EOF
        If .fields("Haber") <> 0 Then
            PrinterFields 0.5, PosLinea, .fields("Fecha"), False
            PrinterFields 2.5, PosLinea, .fields("Cliente"), False
            PrinterFields 10, PosLinea, .fields("TP"), False
            PrinterFields 11, PosLinea, .fields("Numero"), False
            PrinterFields 12.5, PosLinea, .fields("Cheq_Dep"), False
            'If TipoCta Then
            '   PrinterFields 17, PosLinea, .Fields("Haber_ME"), False
            'Else
            PrinterFields 17, PosLinea, .fields("Haber"), False
            'End If
            PosLinea = PosLinea + 0.4
            If PosLinea > LimiteAlto Then
               Printer.Line (0.5, PosLinea)-(19.5, PosLinea), Negro
               Printer.NewPage
               Encabezado 0.5, 19.5
               Printer.FontSize = SizeLetra
            End If
        End If
       .MoveNext
     Loop
     Printer.Line (InicioX, PosLinea)-(19.5, PosLinea), Negro
     PosLinea = PosLinea + 0.1
     PrinterTexto 14.5, PosLinea, "TOTAL CHEQUES"
     PrinterVariables 17, PosLinea, Haber
     PosLinea = PosLinea + 0.4
'''     Printer.FontSize = SizeLetra1
'''     Printer.FontBold = True
'''     Printer.FontUnderline = True
'''     PrinterTexto 0.5, PosLinea, "SALDO REAL CONCILIADO DEL ESTADO DE LA CUENTA:"
'''     Printer.FontUnderline = False
'''     PrinterVariables 16.5, PosLinea, Total_Bancos + Debe - Haber
'''     Printer.FontBold = False
'''     PosLinea = PosLinea + 1
 End If
End With

With DtaTransito.Recordset
 If .RecordCount > 0 Then
     If PosLinea > LimiteAlto Then
        Printer.Line (0.5, PosLinea)-(19.5, PosLinea), Negro
        Printer.NewPage
        Encabezado 0.5, 19.5
        Printer.FontSize = SizeLetra
     End If
     Printer.FontSize = SizeLetra1
     Printer.FontBold = True
     Printer.FontUnderline = True
     PrinterTexto 0.5, PosLinea, "SALDO SEGUN ESTADO DE CUENTA:"
     Printer.FontUnderline = False
     PrinterVariables 13.5, PosLinea, Total_Bancos
     PosLinea = PosLinea + 0.5
     Printer.FontUnderline = True
     PrinterTexto 0.5, PosLinea, "NOTAS DE DEBITOS (registrados por el banco y no en la compañía):"
     PosLinea = PosLinea + 0.5
     Printer.FontSize = SizeLetra
     PrinterTexto 0.5, PosLinea, "F e c h a"
     PrinterTexto 2.5, PosLinea, "Concepto"
     PrinterTexto 10, PosLinea, "TP"
     PrinterTexto 11, PosLinea, "Documento"
     PrinterTexto 17, PosLinea, "V A L O R"
     PosLinea = PosLinea + 0.4
     Printer.FontUnderline = False
     Printer.FontBold = False
    .MoveFirst
     Printer.FontSize = SizeLetra
     Do While Not .EOF
        If .fields("Debe") <> 0 Then
            PrinterFields 0.5, PosLinea, .fields("Fecha"), False
            PrinterFields 2.5, PosLinea, .fields("Concepto"), False
            PrinterFields 10, PosLinea, .fields("TP"), False
            PrinterFields 11, PosLinea, .fields("Documento"), False
            PrinterFields 17, PosLinea, .fields("Debe"), False
            PosLinea = PosLinea + 0.4
            If PosLinea > LimiteAlto Then
               Printer.Line (0.5, PosLinea)-(19.5, PosLinea), Negro
               Printer.NewPage
               Encabezado 0.5, 19.5
               Printer.FontSize = SizeLetra
            End If
        End If
       .MoveNext
     Loop
     Printer.Line (InicioX, PosLinea)-(19.5, PosLinea), Negro
     PosLinea = PosLinea + 0.1
     PrinterTexto 14.5, PosLinea, "TOTAL DEPOSITO"
     PrinterVariables 17, PosLinea, Debe
     PosLinea = PosLinea + 0.4
     Printer.FontSize = SizeLetra1
     Printer.FontBold = True
     Printer.FontUnderline = True
     PrinterTexto 0.5, PosLinea, "DEPOSITOS EN TRANSITO (registrados por el banco y no en la compañía):"
     PosLinea = PosLinea + 0.5
     Printer.FontSize = SizeLetra
     PrinterTexto 0.5, PosLinea, "F e c h a"
     PrinterTexto 2.5, PosLinea, "Beneficiario"
     PrinterTexto 10, PosLinea, "TP"
     PrinterTexto 11, PosLinea, "Numero"
     PrinterTexto 12.5, PosLinea, "Cheque No."
     PrinterTexto 14.5, PosLinea, "Parcial_ME"
     PrinterTexto 17, PosLinea, "V A L O R"
     PosLinea = PosLinea + 0.4
     Printer.FontUnderline = False
     Printer.FontBold = False
    .MoveFirst
     Printer.FontSize = SizeLetra
     Do While Not .EOF
        If .fields("Haber") <> 0 Then
            PrinterFields 0.5, PosLinea, .fields("Fecha"), False
            PrinterFields 2.5, PosLinea, .fields("Concepto"), False
            PrinterFields 10, PosLinea, .fields("TP"), False
            PrinterFields 11, PosLinea, .fields("Documento"), False
            'If TipoCta Then
            '   PrinterFields 17, PosLinea, .Fields("Haber_ME"), False
            'Else
               PrinterFields 17, PosLinea, .fields("Haber"), False
            'End If
            PosLinea = PosLinea + 0.4
            If PosLinea > LimiteAlto Then
               Printer.Line (0.5, PosLinea)-(19.5, PosLinea), Negro
               Printer.NewPage
               Encabezado 0.5, 19.5
               Printer.FontSize = SizeLetra
            End If
        End If
       .MoveNext
     Loop
 End If
End With

''     Printer.Line (InicioX, PosLinea)-(19.5, PosLinea), Negro
''     PosLinea = PosLinea + 0.1
''     PrinterTexto 14.5, PosLinea, "TOTAL CHEQUES"
''     PrinterVariables 17, PosLinea, Haber
''     PosLinea = PosLinea + 0.4
''     Printer.FontSize = SizeLetra1
''     Printer.FontBold = True
''     Printer.FontUnderline = True
''     PrinterTexto 0.5, PosLinea, "SALDO REAL CONCILIADO DEL ESTADO DE LA CUENTA:"
''     Printer.FontUnderline = False
''     PrinterVariables 16.5, PosLinea, Total_Bancos + Debe - Haber
''     Printer.FontBold = False
''     PosLinea = PosLinea + 1

PosLinea = PosLinea + 0.4
Printer.FontSize = SizeLetra1
Printer.FontBold = True
Printer.FontUnderline = True
PrinterTexto 0.5, PosLinea, "SALDO REAL CONCILIADO DEL ESTADO DE LA CUENTA:"
Printer.FontUnderline = False
PrinterVariables 16.5, PosLinea, Total_Bancos + Debe - Haber
Printer.FontBold = False
PosLinea = PosLinea + 1

Printer.Line (InicioX, PosLinea)-(19.5, PosLinea), Negro
PosLinea = PosLinea + 0.1
Printer.FontBold = True
PrinterTexto 0.5, PosLinea, "DIFERENCIA ENTRE SALDOS REALES DE LIBRO BANCOS Y ESTADO DE CUENTAS:"
PrinterVariables 16.5, PosLinea, Total_Saldos
PosLinea = PosLinea + 0.5
Printer.FontBold = False
Printer.Line (InicioX, PosLinea)-(19.5, PosLinea), Negro
PosLinea = PosLinea + 0.1
Printer.Line (InicioX, PosLinea)-(19.5, PosLinea), Negro
PosLinea = PosLinea + 2
PrinterTexto 3, PosLinea, String(Len(NombreContador), "_")
PrinterTexto 11, PosLinea, String(Len(Supervisor), "_")
PosLinea = PosLinea + 0.5
PrinterTexto 3, PosLinea, String(Len(NombreContador) / 2, " ") & "Elaborado por"
PrinterTexto 11, PosLinea, String(Len(Supervisor) / 2, " ") & "Revisado por"
PosLinea = PosLinea + 0.5
PrinterTexto 3, PosLinea, NombreContador
PrinterTexto 11, PosLinea, Supervisor
RatonNormal
MensajeEncabData = ""
'MsgBox "____"
Printer.EndDoc
RatonNormal
End If
Exit Sub
Errorhandler:
    RatonNormal
    ErrorDeImpresion
    Exit Sub
End Sub

Public Sub ImprimirComprobantesDe(ImpSoloReten As Boolean, _
                                  Co As Comprobantes)
Dim AdoComp As ADODB.Recordset
Dim AdoTrans As ADODB.Recordset
Dim AdoBanc As ADODB.Recordset
Dim AdoFact As ADODB.Recordset
Dim AdoRet As ADODB.Recordset
Dim AdoSubC1 As ADODB.Recordset
Dim AdoSubC2 As ADODB.Recordset
 Select Case Co.TP
   Case CompIngreso: Mensajes = "Imprimir Comprobante de Ingreso No. "
   Case CompEgreso: Mensajes = "Imprimir Comprobante de Egreso No. "
   Case CompDiario: Mensajes = "Imprimir Comprobante de Diario No. "
   Case CompNotaDebito: Mensajes = "Imprimir Nota de Debito No. "
   Case CompNotaCredito: Mensajes = "Imprimir Nota de Crédito No. "
 End Select
 Orientacion_Pagina = 1
 Mensajes = Mensajes & Format$(Co.Numero, "00000000") & " en:" & vbCrLf & Printer.DeviceName & "?"
 Titulo = "IMPRESION DE " & Co.TP
 Bandera = False
 SetPrinters.Show 1
 If PonImpresoraDefecto(SetNombrePRN) Then
    Escala_Centimetro 1, TipoTimes, 10
    Co.Fecha = FechaSistema
   'Listar el Comprobante
    sSQL = "SELECT C.*,A.Nombre_Completo,Cl.CI_RUC,Cl.Direccion,Cl.Email," _
         & "Cl.Telefono,Cl.Celular,Cl.FAX,Cl.Cliente,Cl.Codigo,Cl.Ciudad " _
         & "FROM Comprobantes As C,Accesos As A,Clientes As Cl " _
         & "WHERE C.Numero = " & Co.Numero & " " _
         & "AND C.TP = '" & Co.TP & "' " _
         & "AND C.Item = '" & Co.Item & "' " _
         & "AND C.Periodo = '" & Periodo_Contable & "' " _
         & "AND C.CodigoU = A.Codigo " _
         & "AND C.Codigo_B = Cl.Codigo "
    Select_AdoDB AdoComp, sSQL
    If AdoComp.RecordCount > 0 Then Co.Fecha = AdoComp.fields("Fecha")
   'Listar las Transacciones
    sSQL = "SELECT T.Cta,Ca.Cuenta,Parcial_ME,Debe,Haber,Detalle,Cheq_Dep,T.Fecha_Efec,Ca.Item " _
         & "FROM Transacciones As T,Catalogo_Cuentas As Ca " _
         & "WHERE T.TP = '" & Co.TP & "' " _
         & "AND T.Numero = " & Co.Numero & " " _
         & "AND T.Item = '" & Co.Item & "' " _
         & "AND T.Periodo = '" & Periodo_Contable & "' " _
         & "AND T.Item = Ca.Item " _
         & "AND T.Cta = Ca.Codigo " _
         & "AND T.Periodo = Ca.Periodo " _
         & "ORDER BY T.ID,Debe DESC,T.Cta "
    Select_AdoDB AdoTrans, sSQL
   'Llenar Bancos
    sSQL = "SELECT T.Cta,C.TC,C.Cuenta,Co.Fecha,Cl.Cliente,T.Cheq_Dep,T.Debe,T.Haber,T.Fecha_Efec " _
         & "FROM Transacciones As T,Comprobantes As Co,Catalogo_Cuentas As C,Clientes As Cl " _
         & "WHERE T.TP = '" & Co.TP & "' " _
         & "AND T.Numero = " & Co.Numero & " " _
         & "AND T.Item = '" & Co.Item & "' " _
         & "AND T.Periodo = '" & Periodo_Contable & "' " _
         & "AND T.Numero = Co.Numero " _
         & "AND T.TP = Co.TP " _
         & "AND T.Cta = C.Codigo " _
         & "AND T.Item = C.Item " _
         & "AND T.Item = Co.Item " _
         & "AND T.Periodo = C.Periodo " _
         & "AND T.Periodo = Co.Periodo " _
         & "AND C.TC = 'BA' " _
         & "AND Co.Codigo_B = Cl.Codigo "
    Select_AdoDB AdoBanc, sSQL
   'Listar las Retenciones del IVA
    sSQL = "SELECT * " _
         & "FROM Trans_Compras " _
         & "WHERE Numero = " & Co.Numero & " " _
         & "AND TP = '" & Co.TP & "' " _
         & "AND Item = '" & Co.Item & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "ORDER BY Cta_Servicio,Cta_Bienes "
   Select_AdoDB AdoFact, sSQL
  'Listar las Retenciones de la Fuente
   sSQL = "SELECT R.*,TIV.Concepto " _
         & "FROM Trans_Air As R,Tipo_Concepto_Retencion As TIV " _
         & "WHERE R.Numero = " & Co.Numero & " " _
         & "AND R.TP = '" & Co.TP & "' " _
         & "AND R.Item = '" & Co.Item & "' " _
         & "AND TIV.Fecha_Inicio <= #" & BuscarFecha(Co.Fecha) & "# " _
         & "AND TIV.Fecha_Final >= #" & BuscarFecha(Co.Fecha) & "# " _
         & "AND R.Periodo = '" & Periodo_Contable & "' " _
         & "AND R.Tipo_Trans IN ('C','I') " _
         & "AND R.CodRet = TIV.Codigo " _
         & "ORDER BY R.Cta_Retencion "
    Select_AdoDB AdoRet, sSQL
    
   'Llenar SubCtas
    sSQL = "SELECT T.Cta,T.TC,T.Factura,C.Cliente,T.Detalle_SubCta,T.Debitos,T.Creditos,T.Fecha_V,T.Codigo,T.Prima " _
         & "FROM Trans_SubCtas As T,Clientes As C " _
         & "WHERE T.TP = '" & Co.TP & "' " _
         & "AND T.Numero = " & Co.Numero & " " _
         & "AND T.Item = '" & Co.Item & "' " _
         & "AND T.Periodo = '" & Periodo_Contable & "' " _
         & "AND T.TC IN ('C','P') " _
         & "AND T.Codigo = C.Codigo " _
         & "ORDER BY T.Cta,C.Cliente,T.Fecha_V,T.Factura "
    Select_AdoDB AdoSubC1, sSQL
    sSQL = "SELECT T.Cta,T.TC,T.Factura,C.Detalle As Cliente,T.Detalle_SubCta,T.Debitos,T.Creditos,T.Fecha_V,T.Codigo,T.Prima " _
         & "FROM Trans_SubCtas As T,Catalogo_SubCtas As C " _
         & "WHERE T.TP = '" & Co.TP & "' " _
         & "AND T.Numero = " & Co.Numero & " " _
         & "AND T.Item = '" & Co.Item & "' " _
         & "AND T.Periodo = '" & Periodo_Contable & "' " _
         & "AND T.TC NOT IN ('C','P') " _
         & "AND T.TC = C.TC " _
         & "AND T.Item = C.Item " _
         & "AND T.Periodo = C.Periodo " _
         & "AND T.Codigo = C.Codigo " _
         & "ORDER BY T.Cta,C.Detalle,T.Fecha_V,T.Factura "
    Select_AdoDB AdoSubC2, sSQL
    ConceptoComp = Ninguno
    If AdoComp.RecordCount > 0 Then ConceptoComp = AdoComp.fields("Concepto")
    'MsgBox AdoRet.RecordCount
    TipoComp = Co.TP
    Select Case Co.TP
      Case CompIngreso: ImprimirCompIngreso AdoComp, AdoBanc, AdoTrans, AdoSubC1, AdoSubC2
      Case CompEgreso: ImprimirCompEgreso AdoComp, AdoBanc, AdoTrans, AdoFact, AdoRet, AdoSubC1, AdoSubC2, ImpSoloReten
      Case CompDiario: ImprimirCompDiario AdoComp, AdoTrans, AdoFact, AdoRet, AdoSubC1, AdoSubC2, ImpSoloReten
      Case CompNotaDebito: ImprimirCompNota_D_C AdoComp, AdoTrans, AdoSubC1, AdoSubC2, "ND"
      Case CompNotaCredito: ImprimirCompNota_D_C AdoComp, AdoTrans, AdoSubC1, AdoSubC2, "NC"
    End Select
    AdoComp.Close
    AdoTrans.Close
    AdoBanc.Close
    AdoRet.Close
    AdoSubC1.Close
    AdoSubC2.Close
 End If
End Sub

Public Sub ImprimirCompIngreso(DataComp As ADODB.Recordset, _
                               DataBanco As ADODB.Recordset, _
                               DataTrans As ADODB.Recordset, _
                               DataSubC1 As ADODB.Recordset, _
                               DataSubC2 As ADODB.Recordset, _
                               Optional NuevaPagina As Boolean)
On Error GoTo Errorhandler
'Establecemos Espacios y seteos de impresion
RatonReloj
LetraAnterior = Printer.FontName
CIConLineas = ProcesarSeteos(CompIngreso)
'Escala_Centimetro 1, TipoTimes, 10
'Iniciamos la impresion
With DataComp
 If .RecordCount > 0 Then
     Mifecha = .fields("Fecha")
     ImprimirFormatoComprobantes CompIngreso, FechaAnio(.fields("Fecha")), .fields("Numero"), CIConLineas, SetD(1).PosY
     Printer.FontBold = False
     Printer.FontSize = SetD(2).Tamaño
     PrinterTexto SetD(2).PosX, SetD(2).PosY, FechaStrgCorta(.fields("Fecha"))
     Printer.FontSize = SetD(3).Tamaño
     PrinterFields SetD(3).PosX, SetD(3).PosY, .fields("Cliente")
     Printer.FontSize = SetD(4).Tamaño
     PrinterFields SetD(4).PosX, SetD(4).PosY, .fields("CI_RUC")
     Printer.FontSize = SetD(5).Tamaño
     PrinterLineas SetD(5).PosX, SetD(5).PosY, .fields("Concepto"), 16
     Printer.FontSize = SetD(6).Tamaño
     PrinterTexto SetD(6).PosX, SetD(6).PosY, Format$(.fields("Monto_Total"), "#,##0.00")
     Printer.FontSize = SetD(7).Tamaño
     PrinterNum SetD(7).PosX, SetD(7).PosY, .fields("Monto_Total")
     Printer.FontSize = SetD(8).Tamaño
     PrinterFields SetD(8).PosX, SetD(8).PosY, .fields("Efectivo")
     'Diferencia = .Fields("Monto_Total") - .Fields("Efectivo")
     'Printer.FontSize = SetD(9).Tamaño
     'PrinterVariables SetD(9).PosX, SetD(9).PosY, Diferencia
     Diferencia = .fields("Monto_Total")
     Diferencia = Diferencia - ImprimirBancos(CompIngreso, DataBanco, 10)
     Printer.FontSize = SetD(9).Tamaño
     PrinterVariables SetD(9).PosX, SetD(9).PosY, Diferencia
     ImprimirAsientos DataTrans, DataComp, DataSubC1, DataSubC2, 13, 18, CIConLineas, CompIngreso, .fields("T"), .fields("Cotizacion")
     Printer.FontName = LetraAnterior
 End If
End With
If NuevaPagina Then Printer.NewPage Else Printer.EndDoc
RatonNormal
Exit Sub
Errorhandler:
    RatonNormal
    ErrorDeImpresion
    Exit Sub
End Sub

Public Sub ImprimirCompEgreso(DataComp As ADODB.Recordset, _
                              DataBanco As ADODB.Recordset, _
                              DataTrans As ADODB.Recordset, _
                              DataFact As ADODB.Recordset, _
                              DataRets As ADODB.Recordset, _
                              DataSubC1 As ADODB.Recordset, _
                              DataSubC2 As ADODB.Recordset, _
                              ImpSoloRet As Boolean, _
                              Optional NuevaPagina As Boolean, _
                              Optional NoImpRet As Boolean)
'Establecemos Espacios y seteos de impresion
Dim ImpCheque As Boolean
Dim Benef As String
Dim TipoBank As String
Dim PosX_Chq As Single
Dim PosY_Chq As Single
On Error GoTo Errorhandler
RatonReloj
LetraAnterior = Printer.FontName
CEConLineas = ProcesarSeteos(CompEgreso)
NombreCiudad = Leer_Campo_Empresa("Ciudad")
Pagina = 1
'Iniciamos la impresion
If DataComp.RecordCount > 0 Then
With DataComp
 If .RecordCount > 0 Then
     Mifecha = .fields("Fecha")
     If ImpSoloRet = False Then
        ImprimirFormatoComprobantes CompEgreso, FechaAnio(.fields("Fecha")), .fields("Numero"), CEConLineas, SetD(1).PosY
        Printer.FontBold = False
        Printer.FontSize = SetD(2).Tamaño
        PrinterTexto SetD(2).PosX, SetD(2).PosY, FechaStrgCorta(.fields("Fecha"))
        Printer.FontSize = SetD(3).Tamaño
        PrinterFields SetD(3).PosX, SetD(3).PosY, .fields("Cliente")
        Printer.FontSize = SetD(4).Tamaño
        PrinterFields SetD(4).PosX, SetD(4).PosY, .fields("CI_RUC")
        Printer.FontSize = SetD(5).Tamaño
        PrinterLineas SetD(5).PosX, SetD(5).PosY, .fields("Concepto"), 16
        Printer.FontSize = SetD(6).Tamaño
        PrinterTexto SetD(6).PosX, SetD(6).PosY, Format$(.fields("Monto_Total"), "#,##0.00")
        Printer.FontSize = SetD(7).Tamaño
        PrinterNum SetD(7).PosX, SetD(7).PosY, .fields("Monto_Total")
        Printer.FontSize = SetD(8).Tamaño
        PrinterFields SetD(8).PosX, SetD(8).PosY, .fields("Efectivo")
        Diferencia = .fields("Monto_Total") - .fields("Efectivo")
        Printer.FontSize = SetD(9).Tamaño
        PrinterVariables SetD(9).PosX, SetD(9).PosY, Diferencia
     End If
 End If
End With
If ImpSoloRet = False Then
   Diferencia = Diferencia - ImprimirBancos(CompEgreso, DataBanco, 10)
   'Printer.FontSize = SetD(9).Tamaño
   'PrinterVariables SetD(9).PosX, SetD(9).PosY, Diferencia
   ImprimirAsientos DataTrans, DataComp, DataSubC1, DataSubC2, 13, 18, CEConLineas, CompEgreso, DataComp.fields("T"), DataComp.fields("Cotizacion"), SetD(1).PosY
   Valor = 0
   If NuevaPagina = False And Modulo <> "CAJACREDITO" Then
      With DataBanco
       If .RecordCount > 0 Then
          .MoveFirst
           Mensajes = "Imprimir Cheque No. " & .fields("Cheq_Dep")
           Titulo = "Pregunta de Impresion"
           If BoxMensaje = vbYes Then
              Printer.FontBold = True
              Do While Not .EOF
                 If .fields("Haber") > 0 Then
                    'Impresion de constancia del cheque, no importa si no tiene papel copia o quimico
                     If SetD(20).PosX > 0 And SetD(20).PosY > 0 Then
                        PosX_Chq = SetD(20).PosX
                        PosY_Chq = SetD(20).PosY
                        TipoBank = TrimStrg(Format$(MidStrg(.fields("Cta"), Len(.fields("Cta")) - 1, 3), "00"))
                        CCHQConLineas = ProcesarSeteos(TipoBank)
                        Printer.FontSize = SetD(2).Tamaño
                        PrinterFields SetD(2).PosX + PosX_Chq, SetD(2).PosY + PosY_Chq, .fields("Cliente"), False
                        Printer.FontSize = SetD(3).Tamaño
                        PrinterTexto SetD(3).PosX + PosX_Chq, SetD(3).PosY + PosY_Chq, Format$(.fields("Haber"), "#,###.00")
                        Printer.FontSize = SetD(10).Tamaño
                        PrinterTexto SetD(10).PosX + PosX_Chq, SetD(10).PosY + PosY_Chq, ULCase(NombreCiudad)
                        Printer.FontSize = SetD(6).Tamaño
                        If CFechaLong(.fields("Fecha")) < CFechaLong(.fields("Fecha_Efec")) Then
                           PrinterTexto SetD(6).PosX + PosX_Chq, SetD(6).PosY + PosY_Chq, Format$(.fields("Fecha_Efec"), "yyyy/MM/dd")
                        Else
                           PrinterTexto SetD(6).PosX + PosX_Chq, SetD(6).PosY + PosY_Chq, Format$(.fields("Fecha"), "yyyy/MM/dd")
                        End If
                        If SetD(4).PosX > 0 And SetD(4).PosY > 0 Then
                           Printer.FontSize = SetD(4).Tamaño
                           PrinterNumCheque SetD(4).PosX + PosX_Chq, SetD(4).PosY + PosY_Chq, SetD(5).PosX, .fields("Haber")
                        End If
                        If SetD(9).PosX > 0 And SetD(9).PosY > 0 Then
                           Cadena = Empresa & " " & Moneda & " " & Format$(.fields("Haber"), "#,##0.00") & "**"
                           Printer.FontSize = SetD(9).Tamaño
                           PrinterTexto SetD(9).PosX + PosX_Chq, SetD(9).PosY + PosY_Chq, Cadena
                        End If
                     End If
                     TipoBank = TrimStrg(Format$(MidStrg(.fields("Cta"), Len(.fields("Cta")) - 1, 3), "00"))
                     CCHQConLineas = ProcesarSeteos(TipoBank)
                     
                     If SetD(7).PosX = 1 Then Printer.NewPage
                     Printer.FontSize = SetD(2).Tamaño
                     PrinterFields SetD(2).PosX, SetD(2).PosY, .fields("Cliente"), False
                     Printer.FontSize = SetD(3).Tamaño
                     PrinterTexto SetD(3).PosX, SetD(3).PosY, Format$(.fields("Haber"), "#,###.00")
                     Printer.FontSize = SetD(10).Tamaño
                     PrinterTexto SetD(10).PosX + PosX_Chq, SetD(10).PosY + PosY_Chq, ULCase(NombreCiudad)
                     Printer.FontSize = SetD(6).Tamaño
                     If CFechaLong(.fields("Fecha")) < CFechaLong(.fields("Fecha_Efec")) Then
                        PrinterTexto SetD(6).PosX, SetD(6).PosY, Format$(.fields("Fecha_Efec"), "yyyy/MM/dd")
                     Else
                        PrinterTexto SetD(6).PosX, SetD(6).PosY, Format$(.fields("Fecha"), "yyyy/MM/dd")
                     End If
                     If SetD(4).PosX > 0 And SetD(4).PosY > 0 Then
                        Printer.FontSize = SetD(4).Tamaño
                        PrinterNumCheque SetD(4).PosX, SetD(4).PosY, SetD(5).PosX, .fields("Haber")
                     End If
                     If SetD(9).PosX > 0 And SetD(9).PosY > 0 Then
                        Cadena = Empresa & " " & Moneda & " " & Format$(.fields("Haber"), "#,##0.00") & "**"
                        Printer.FontSize = SetD(9).Tamaño
                        PrinterTexto SetD(9).PosX, SetD(9).PosY, Cadena
                     End If
                 End If
                .MoveNext
              Loop
              Printer.FontBold = False
           End If
       End If
      End With
   End If
End If
If NoImpRet = False Then
   If DataRets.RecordCount > 0 Then
      Mensajes = "Imprimir Retencion"
      Titulo = "Pregunta de Impresion"
      If BoxMensaje = vbYes Then ImprimirCompRetencion ImpSoloRet, DataFact, DataComp, DataRets
   End If
End If
Printer.FontName = LetraAnterior
If NuevaPagina Then Printer.NewPage Else Printer.EndDoc
End If
RatonNormal
Exit Sub
Errorhandler:
    RatonNormal
    ErrorDeImpresion
    Exit Sub
End Sub

Public Sub ImprimirCompRetencion(SoloRet As Boolean, _
                                 DataFact As ADODB.Recordset, _
                                 DataComp As ADODB.Recordset, _
                                 DataRets As ADODB.Recordset)
Dim Copias As Boolean
Dim ConceptoRet As String
Dim PorcentajeRet As String
Dim PosXRet As Single
Dim PosYRet As Single
'Establecemos Espacios y seteos de impresion
PorteLetra = Printer.FontSize
LetraAnterior = Printer.FontName
Printer.FontName = TipoTimes
CRConLineas = ProcesarSeteos(CompRetencion)
PosXRet = 0: PosYRet = 0
Volver_Imp_Ret:
Printer.FontBold = True
  'MsgBox PosXRet & "-" & PosYRet
  'MsgBox "Retencion: " & SetD(23).PosX & vbCrLf & SoloRet
If SetD(23).PosX <> 0 And SoloRet = False Then Printer.NewPage

If CRConLineas Then
  'Formato de Retencion
   If SetD(24).PosX > 0 And SetD(24).PosY > 0 And SetD(25).PosX > 0 And SetD(25).PosY > 0 Then
      Dibujo = RutaSistema & "\FORMATOS\RETENCIO.GIF"
      PrinterPaint Dibujo, PosXRet + SetD(24).PosX, PosYRet + SetD(24).PosY, SetD(25).PosX, SetD(25).PosY
   End If
   If SetD(1).PosX > 0 And SetD(1).PosY > 0 Then
     'Formato de Logotipo
      PrinterPaint LogoTipo, PosXRet + 0.2, PosYRet + SetD(1).PosY - 0.1, PosXRet + 4, SetD(1).PosY + 2
      Printer.FontSize = 18
      PrinterTexto CentrarTexto(Empresa), PosYRet + SetD(1).PosY, Empresa
      Printer.FontSize = 12
      PrinterTexto PosXRet + SetD(1).PosX, PosYRet + SetD(1).PosY, "R.U.C." & RUC
      Printer.FontSize = 8
      Cadena = "Dirección: " & Direccion
      PrinterTexto PosXRet + SetD(1).PosX, PosYRet + SetD(1).PosY, Cadena
      Cadena = NombreCiudad & " - " & NombrePais & ".  Teléfono: " & Telefono1 & "/FAX: " & FAX
      PrinterTexto PosXRet + SetD(1).PosX, PosYRet + SetD(1).PosY, Cadena
      Printer.FontBold = False
      PosLinea = InicioY
   End If
End If
PosLinea = 0.2
Printer.FontBold = False
Pagina = 1
'Iniciamos la impresion
With DataComp
 If .RecordCount > 0 Then
    .MoveFirst
     Printer.FontSize = SetD(3).Tamaño
     If SetD(3).PosX > 0 Then PrinterFields PosXRet + SetD(3).PosX, PosYRet + SetD(3).PosY, .fields("Cliente")
     Printer.FontSize = SetD(9).Tamaño
     If SetD(9).PosX > 0 Then PrinterFields PosXRet + SetD(9).PosX, PosYRet + SetD(9).PosY, .fields("CI_RUC")
     Printer.FontSize = SetD(4).Tamaño
     If SetD(4).PosX > 0 Then PrinterFields PosXRet + SetD(4).PosX, PosYRet + SetD(4).PosY, .fields("Direccion")
     Printer.FontSize = SetD(6).Tamaño
     If SetD(6).PosX > 0 Then PrinterFields PosXRet + SetD(6).PosX, PosYRet + SetD(6).PosY, .fields("Telefono")
     Printer.FontSize = SetD(10).Tamaño
     If SetD(10).PosX > 0 Then PrinterTexto PosXRet + SetD(10).PosX, PosYRet + SetD(10).PosY, FechaAnio(.fields("Fecha"))
     Printer.FontSize = SetD(2).Tamaño
     If SetD(2).PosX > 0 Then PrinterFields PosXRet + SetD(2).PosX, PosYRet + SetD(2).PosY, .fields("Fecha")
     Printer.FontSize = SetD(29).Tamaño
     If SetD(29).PosX > 0 Then PrinterTexto PosXRet + SetD(29).PosX, PosYRet + SetD(29).PosY, Day(.fields("Fecha"))
     Printer.FontSize = SetD(30).Tamaño
     If SetD(30).PosX > 0 Then PrinterTexto PosXRet + SetD(30).PosX, PosYRet + SetD(30).PosY, MesesLetras(Month(.fields("Fecha")))
     Printer.FontSize = SetD(31).Tamaño
     If SetD(31).PosX > 0 Then PrinterTexto PosXRet + SetD(31).PosX, PosYRet + SetD(31).PosY, Year(.fields("Fecha"))
     If Len(NombreCiudad) > 1 Then
        Cadena = ULCase(NombreCiudad) & ", " & FechaStrg(.fields("Fecha"))
     Else
        Cadena = FechaStrg(.fields("Fecha"))
     End If
     Printer.FontSize = SetD(35).Tamaño
     If SetD(35).PosX > 0 Then PrinterTexto PosXRet + SetD(35).PosX, PosYRet + SetD(35).PosY, Cadena
     Printer.FontSize = SetD(5).Tamaño
     If SetD(5).PosX > 0 Then PrinterFields PosXRet + SetD(5).PosX, PosYRet + SetD(5).PosY, .fields("Ciudad")
     Printer.FontSize = SetD(1).Tamaño
     If SetD(1).PosX > 0 Then PrinterTexto PosXRet + SetD(1).PosX, PosYRet + SetD(1).PosY, Autorizacion
 End If
End With
PosLinea = PosYRet + SetD(13).PosY: SumaSaldo = 0
With DataFact
'MsgBox .RecordCount
 If .RecordCount > 0 Then
    .MoveFirst
    'Encabezado de la Retencion
     Printer.FontSize = SetD(1).Tamaño
     If SetD(1).PosX > 0 Then PrinterTexto PosXRet + SetD(1).PosX, PosYRet + SetD(1).PosY, Format$(.fields("Secuencial"), "000000000")
     Cadena = "Factura"
     Select Case .fields("TipoComprobante")
       Case "1": Cadena = "Factura"
       Case "2": Cadena = "Nota de Venta"
       Case "3": Cadena = "Liquidación de Compra de Bienes o Prestación de Servicios"
       Case "4": Cadena = "Notas de Crédito"
       Case "5": Cadena = "Notas de Débito"
       Case "8": Cadena = "Boletos a entradas a espectaculos públicos"
       Case "9": Cadena = "Tiquetes o Vales emitidos por máquinas registradoras"
     End Select
     Printer.FontSize = SetD(7).Tamaño
     If SetD(7).PosX > 0 Then PrinterTexto PosXRet + SetD(7).PosX, PosYRet + SetD(7).PosY, UCaseStrg(Cadena)
     Printer.FontSize = SetD(11).Tamaño
     If SetD(11).PosX > 0 Then PrinterLineas PosXRet + SetD(11).PosX, PosYRet + SetD(11).PosY, ConceptoComp, SetD(12).PosX
     PosLinea = SetD(13).PosY: SumaSaldo = 0
     Printer.FontBold = False
     Printer.FontSize = SetD(8).Tamaño
     If SetD(8).PosX > 0 Then PrinterTexto PosXRet + SetD(8).PosX, PosYRet + SetD(8).PosY, .fields("Establecimiento") & .fields("PuntoEmision") & "-" & Format$(.fields("Secuencial"), "000000000")
     Printer.FontSize = SetD(33).Tamaño
     If SetD(33).PosX > 0 Then PrinterTexto PosXRet + SetD(33).PosX, PosYRet + SetD(33).PosY, .fields("Autorizacion")
     Printer.FontSize = SetD(34).Tamaño
     If SetD(34).PosX > 0 Then PrinterTexto PosXRet + SetD(34).PosX, PosYRet + SetD(34).PosY, .fields("FechaCaducidad")
    '-----------------------------------------------------------------------------------
     Total_DetRet = .fields("ValorRetBienes") + .fields("ValorRetServicios")
     SumaSaldo = SumaSaldo + .fields("ValorRetBienes") + .fields("ValorRetServicios")
     If SetD(26).PosY > 0 Then PosLinea = SetD(26).PosY
     If .fields("ValorRetBienes") > 0 Then
         Total_RetCta = .fields("MontoIvaBienes")
         Total_DetRet = .fields("ValorRetBienes")
         PorcentajeRet = .fields("Porc_Bienes") & "%"
         ConceptoRet = "Retención por Bienes"
        '-------------------------------------------------------------------------
         Printer.FontSize = SetD(15).Tamaño
         If SetD(15).PosX > 0 Then PrinterTexto PosXRet + SetD(15).PosX, PosLinea, "R. IVA(B)"
         Printer.FontSize = SetD(17).Tamaño
         If SetD(17).PosX > 0 Then PrinterTexto PosXRet + SetD(17).PosX, PosLinea, ConceptoRet      'Concepto de Retencion
         Printer.FontSize = SetD(19).Tamaño
         If SetD(19).PosX > 0 Then PrinterVariables PosXRet + SetD(19).PosX, PosLinea, Total_RetCta 'Base Imponible
         Printer.FontSize = SetD(21).Tamaño
         If SetD(21).PosX > 0 Then PrinterVariables PosXRet + SetD(21).PosX, PosLinea, Total_DetRet 'Valor Retenido
         Printer.FontSize = SetD(20).Tamaño
         If SetD(20).PosX > 0 Then PrinterVariables PosXRet + SetD(20).PosX, PosLinea, PorcentajeRet 'Porcentaje de Ret.
         PosLinea = PosLinea + 0.4
     End If
     If .fields("ValorRetServicios") > 0 Then
         Total_RetCta = .fields("MontoIvaServicios")
         Total_DetRet = .fields("ValorRetServicios")
         PorcentajeRet = .fields("Porc_Servicios") & "%"
         ConceptoRet = "Retención por Servicios"
        '-------------------------------------------------------------------------
         Printer.FontSize = SetD(15).Tamaño
         If SetD(15).PosX > 0 Then PrinterTexto PosXRet + SetD(15).PosX, PosLinea, "R. IVA(S)"
         Printer.FontSize = SetD(17).Tamaño
         If SetD(17).PosX > 0 Then PrinterTexto PosXRet + SetD(17).PosX, PosLinea, ConceptoRet      'Concepto de Retencion
         Printer.FontSize = SetD(19).Tamaño
         If SetD(19).PosX > 0 Then PrinterVariables PosXRet + SetD(19).PosX, PosLinea, Total_RetCta 'Base Imponible
         Printer.FontSize = SetD(21).Tamaño
         If SetD(21).PosX > 0 Then PrinterVariables PosXRet + SetD(21).PosX, PosLinea, Total_DetRet 'Valor Retenido
         Printer.FontSize = SetD(20).Tamaño
         If SetD(20).PosX > 0 Then PrinterVariables PosXRet + SetD(20).PosX, PosLinea, PorcentajeRet 'Porcentaje de Ret.
         PosLinea = PosLinea + 0.4
     End If
 End If
End With
If SetD(27).PosY > 0 Then PosLinea = PosYRet + SetD(27).PosY
With DataRets
 If .RecordCount > 0 Then
    .MoveFirst
     Printer.FontBold = False
     Do While Not .EOF
        Total_RetCta = .fields("BaseImp")
        Total_DetRet = .fields("ValRet")
        SumaSaldo = SumaSaldo + .fields("ValRet")
        Cadena1 = .fields("TP") & "-" & .fields("Numero")
        ConceptoRet = MidStrg(.fields("Concepto"), 1, 40) & "..."
        PorcentajeRet = Format$(.fields("Porcentaje"), "#0.00%")
       '-------------------------------------------------------------------------------
        Printer.FontSize = SetD(14).Tamaño
        If SetD(14).PosX > 0 Then PrinterTexto PosXRet + SetD(14).PosX, PosLinea, "R. Renta "          'Tipo de Retencion
        Printer.FontSize = SetD(17).Tamaño
        If SetD(17).PosX > 0 Then PrinterTexto PosXRet + SetD(17).PosX, PosLinea, ConceptoRet      'Concepto de Retencion
        Printer.FontSize = SetD(18).Tamaño
        If SetD(18).PosX > 0 Then PrinterTexto PosXRet + SetD(18).PosX, PosLinea, Cadena1          'Referencia
        Printer.FontSize = SetD(19).Tamaño
        If SetD(19).PosX > 0 Then PrinterVariables PosXRet + SetD(19).PosX, PosLinea, Total_RetCta 'Base Imponible
        Printer.FontSize = SetD(21).Tamaño
        If SetD(21).PosX > 0 Then PrinterVariables PosXRet + SetD(21).PosX, PosLinea, Total_DetRet 'Valor Retenido
        Printer.FontSize = SetD(16).Tamaño
        If SetD(16).PosX > 0 Then PrinterFields PosXRet + SetD(16).PosX, PosLinea, .fields("CodRet") 'Codigo de Retencion
        Printer.FontSize = SetD(20).Tamaño
        If SetD(20).PosX > 0 Then PrinterVariables PosXRet + SetD(20).PosX, PosLinea, PorcentajeRet 'Porcentaje de Retencion
        PosLinea = PosLinea + 0.4
       .MoveNext
     Loop
     Printer.FontSize = SetD(22).Tamaño
     If SetD(22).PosX > 0 Then PrinterVariables PosXRet + SetD(22).PosX, PosYRet + SetD(22).PosY, SumaSaldo 'Total Retenido
     Printer.FontSize = SetD(32).Tamaño
     If SetD(32).PosX > 0 Then PrinterTexto PosXRet + SetD(32).PosX, PosYRet + SetD(32).PosY, UCaseStrg(Cambio_Letras(SumaSaldo)) 'Total Retenido en letras
 End If
End With
Mensajes = "Imprimir Copia de Retencion"
Titulo = "Pregunta de Impresion"
If BoxMensaje = vbYes Then
   If SetD(28).PosX > 0 And SetD(28).PosY > 0 Then
      PosXRet = SetD(28).PosX
      PosYRet = SetD(28).PosY
      SoloRet = True
   Else
      PosXRet = 0
      PosYRet = 0
   End If
   GoTo Volver_Imp_Ret
End If
'If SetD(23).PosX = 0 And SoloRet = False Then Printer.NewPage
Printer.FontSize = PorteLetra
Printer.FontName = LetraAnterior
End Sub

Function Niveles(CodigoCta As String) As Byte
Dim NumNivel As Byte
   NumNivel = 0
  'MsgBox CodigoCta
   If Len(CodigoCta) > 12 Then
      Niveles = 6
   Else
      For I = 1 To Len(CodigoCta)
          If MidStrg(CodigoCta, I, 1) = "." Then NumNivel = NumNivel + 1
      Next I
      Niveles = NumNivel + 1
   End If
   If CodigoCta = " " Then Niveles = 1
End Function

Public Function TiposCtaStrg(TipoCuenta As String) As String
Dim Resultado As String
   Select Case MidStrg(TipoCuenta, 1, 1)
      Case "1":  Resultado = "ACTIVO"
      Case "2":  Resultado = "PASIVO"
      Case "3":  Resultado = "CAPITAL"
      Case "4":  Resultado = "INGRESO"
      Case "5":  Resultado = "EGRESO"
      Case Else: Resultado = "NINGUNA"
   End Select
   TiposCtaStrg = Resultado
End Function

Public Sub ImprimirCheques(Datas As Adodc, _
                           FormaImp As Byte)
Dim NombreBanco As String
On Error GoTo Errorhandler
RatonReloj
InicioX = 0.5: InicioY = 0
DataAnchoCampos InicioX, Datas, 8, TipoTimes, FormaImp
Ancho(0) = 0.5
Ancho(1) = 6.5
Ancho(2) = 11
Ancho(3) = 13
Ancho(4) = 13.8
Ancho(5) = 15.3
Ancho(6) = 17
Ancho(7) = 19.5
Pagina = 1
'Iniciamos la impresion
Printer.FontBold = False
With Datas.Recordset
     .MoveFirst
      NombreBanco = .fields("Cuenta")
      EncabezadoData Datas
      Printer.FontSize = 8
      Total = 0
      PrinterFields Ancho(0), PosLinea, .fields("Cuenta"), False
      Do While Not .EOF
         If NombreBanco <> .fields("Cuenta") Then
            Printer.Line (Ancho(0), PosLinea)-(Ancho(CantCampos), PosLinea), Negro
            PosLinea = PosLinea + 0.1
            PrinterTexto Ancho(5), PosLinea, "TOTAL"
            PrinterVariables Ancho(6), PosLinea, Total
            PosLinea = PosLinea + 0.6
            Printer.Line (Ancho(0), PosLinea)-(Ancho(CantCampos), PosLinea), Negro
            PosLinea = PosLinea + 0.1
            NombreBanco = .fields("Cuenta")
            PrinterFields Ancho(0), PosLinea, .fields("Cuenta"), False
            Total = 0
         End If
         PrinterFields Ancho(1), PosLinea, .fields("Beneficiario"), False
         PrinterFields Ancho(2), PosLinea, .fields("Fecha"), False
         PrinterFields Ancho(3), PosLinea, .fields("TP"), False
         PrinterFields Ancho(4), PosLinea, .fields("Numero"), False
         PrinterFields Ancho(5), PosLinea, .fields("Cheq_Dep"), False
         PrinterFields Ancho(6), PosLinea, .fields("Valor"), False
         For I = 0 To CantCampos
            Printer.Line (Ancho(I), PosLinea - 0.1)-(Ancho(I), PosLinea + 0.4), Negro
         Next I
         Total = Total + .fields("Valor")
         PosLinea = PosLinea + 0.4
         If PosLinea > LimiteAlto Then
            Printer.Line (Ancho(0), PosLinea)-(Ancho(CantCampos), PosLinea), Negro
            Printer.NewPage
            EncabezadoData Datas
            Printer.FontSize = 8
         End If
        .MoveNext
      Loop
End With
Printer.Line (Ancho(0), PosLinea)-(Ancho(CantCampos), PosLinea), Negro
PosLinea = PosLinea + 0.1
PrinterTexto Ancho(5), PosLinea, "TOTAL"
PrinterVariables Ancho(6), PosLinea, Total
UltimaLinea = PosLinea + 0.6
Printer.Line (Ancho(0), PosLinea)-(Ancho(CantCampos), PosLinea), Negro
RatonNormal
MensajeEncabData = ""
Printer.EndDoc
Exit Sub
Errorhandler:
    RatonNormal
    ErrorDeImpresion
    Exit Sub
End Sub

Public Sub ImprimirSubCtas(Datas As Adodc, _
                           Datas1 As Adodc, _
                           Datas2 As Adodc, _
                           FinDoc As Boolean, _
                           FormaImp As Byte, _
                           SizeLetra As Integer, _
                           MN_ME As Boolean, _
                           ImpResum As Boolean)
On Error GoTo Errorhandler
Mensajes = "Seguro de Imprimir en:" & vbCrLf & Printer.DeviceName & "?"
Titulo = "IMPRESION"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then

RatonReloj
InicioX = 0.5: InicioY = 0
Escala_Centimetro FormaImp, TipoVerdana, SizeLetra
C = 9
CantCampos = C - 1
ReDim Ancho(C) As Single
Ancho(0) = 0.5  ' Fecha
Ancho(1) = 2.2  ' Beneficiario
If MN_ME Then
   Ancho(2) = 6.5  ' Cuenta
   Ancho(3) = 9.8  ' TP
   Ancho(4) = 10.5 ' Numero
   Ancho(5) = 12   ' Debitos
   Ancho(6) = 14.5 ' Creditos
   Ancho(7) = 17   ' Valor_ME
   Ancho(8) = 19.5
   CantCampos = 8
Else
   Ancho(2) = 8    ' Cuenta
   Ancho(3) = 12.6   ' TP
   Ancho(4) = 13.2   ' Numero
   Ancho(5) = 15     ' Debitos
   Ancho(6) = 17.5   ' Creditos
   Ancho(7) = 20
   CantCampos = 7
End If
Pagina = 1
Debe = 0: Haber = 0: Parcial_ME = 0: Valor = 0
'Iniciamos la impresion
EncabezadoDataSubCta Datas2, MN_ME, ImpResum
Printer.FontBold = False
'If ImpResum Then GoTo Solo_Resumen
 With Datas.Recordset
  If .RecordCount > 0 Then
     .MoveFirst
      Printer.FontSize = SizeLetra
      Do While Not .EOF
         If ImpResum = False Then
         PrinterFields Ancho(0), PosLinea, .fields("Fecha")
         PrinterTexto Ancho(1), PosLinea, ULCase(.fields("Beneficiario"))
         PrinterFields Ancho(2), PosLinea, .fields("Cuenta")
         PrinterFields Ancho(3), PosLinea, .fields("Tipo")
         PrinterFields Ancho(4), PosLinea, .fields("Numero")
         PrinterFields Ancho(5), PosLinea, .fields("Debitos")
         PrinterFields Ancho(6), PosLinea, .fields("Creditos")
         End If
         If MN_ME Then
            If ImpResum = False Then PrinterFields Ancho(7), PosLinea, .fields("Valor_ME"), True
            Parcial_ME = Parcial_ME - .fields("Valor_ME")
            If .fields("Valor_ME") >= 0 Then Valor = Valor + .fields("Valor_ME")
            If ImpResum = False Then Printer.Line (Ancho(8), PosLinea - 0.1)-(Ancho(8), PosLinea + 0.4), Negro
         Else
            'If ImpResum = False Then Printer.Line (Ancho(7), PosLinea - 0.1)-(Ancho(7), PosLinea + 0.4), NEGRO
         End If
         Debe = Debe + .fields("Debitos")
         Haber = Haber + .fields("Creditos")
         If ImpResum = False Then
            PosLinea = PosLinea + 0.4
            If PosLinea >= LimiteAlto Then
               Printer.Line (InicioX, PosLinea)-(Ancho(CantCampos), PosLinea), Negro
               Printer.NewPage
               PosLinea = 0.1
               EncabezadoDataSubCta Datas2, MN_ME, ImpResum
               Printer.FontSize = SizeLetra
            End If
         End If
        .MoveNext
      Loop
      If ImpResum = False Then
         Printer.Line (InicioX, PosLinea)-(Ancho(CantCampos), PosLinea), Negro
         PosLinea = PosLinea + 0.05
         Printer.Line (InicioX, PosLinea)-(Ancho(CantCampos), PosLinea), Negro
         PosLinea = PosLinea + 0.1
         PrinterTexto Ancho(2), PosLinea, "T O T A L E S"
         PrinterVariables Ancho(5), PosLinea, Debe
         PrinterVariables Ancho(6), PosLinea, Haber
         If MN_ME Then
            PrinterVariables Ancho(7), PosLinea, Parcial_ME
         End If
      End If
  End If
 End With
If ImpResum Then
   PosLinea = PosLinea + 0.4
Else
   PosLinea = PosLinea + 0.8
End If
 With Datas1.Recordset
  If .RecordCount > 0 Then
     .MoveFirst
      'Printer.NewPage
      'PosLinea = 0.5
      Printer.FontBold = True
      Printer.FontSize = 14
      Cadena = "E S T A D O     D E     R E S U L T A D O S"
      PrinterTexto CentrarTexto(Cadena), PosLinea, Cadena
      PosLinea = PosLinea + 0.8
      Printer.FontSize = SizeLetra
      Printer.FontUnderline = True
      PrinterTexto Ancho(1), PosLinea, "C U E N T A"
      PrinterTexto Ancho(2), PosLinea, "DEBITOS"
      PrinterTexto Ancho(3) - 0.5, PosLinea, "CREDITOS"
      PrinterTexto Ancho(5), PosLinea, "PRESUPUESTO"
      If MN_ME Then
         PrinterTexto Ancho(6), PosLinea, "DEBITOS/ME"
         PrinterTexto Ancho(7), PosLinea, "DIFERENCIA(+/-)"
      Else
         PrinterTexto Ancho(6), PosLinea, "DIFERENCIA(+/-)"
      End If
      Printer.FontUnderline = False
      Printer.FontBold = False
      PosLinea = PosLinea + 0.5
      SumaDebe = 0: SumaHaber = 0: Diferencia = 0: Sumatoria = 0
      Do While Not .EOF
         PrinterTexto Ancho(0), PosLinea, ULCase(.fields("Beneficiario"))
         PrinterFields Ancho(2), PosLinea, .fields("Suma_Debitos")
         PrinterFields Ancho(3) - 0.5, PosLinea, .fields("Suma_Creditos")
         'PrinterFields Ancho(5), PosLinea, .Fields("Presupuesto"), True
         If MN_ME Then
            PrinterFields Ancho(6), PosLinea, .fields("Valor_ME")
            PrinterFields Ancho(7), PosLinea, .fields("Diferencia")
            Printer.Line (Ancho(8), PosLinea - 0.1)-(Ancho(8), PosLinea + 0.4), Negro
         Else
            'PrinterFields Ancho(6), PosLinea, .Fields("Diferencia"), True
            'Printer.Line (Ancho(7), PosLinea - 0.1)-(Ancho(7), PosLinea + 0.4), NEGRO
         End If
         PosLinea = PosLinea + 0.4
         If PosLinea >= LimiteAlto Then
            PosLinea = PosLinea + 0.1
            Printer.Line (InicioX, PosLinea)-(Ancho(CantCampos), PosLinea), Negro
            Printer.NewPage
            PosLinea = 0.5
            Printer.FontSize = SizeLetra
            Printer.FontBold = True
            Printer.FontUnderline = True
            PrinterTexto Ancho(1), PosLinea, "C U E N T A"
            PrinterTexto Ancho(2), PosLinea, "DEBITOS"
            PrinterTexto Ancho(3) - 0.5, PosLinea, "CREDITOS"
            PrinterTexto Ancho(5), PosLinea, "PRESUPUESTO"
            If MN_ME Then
               PrinterTexto Ancho(6), PosLinea, "DEBITOS/ME"
               PrinterTexto Ancho(7), PosLinea, "DIFERENCIA(+/-)"
            Else
               PrinterTexto Ancho(6), PosLinea, "DIFERENCIA(+/-)"
            End If
            Printer.FontUnderline = False
            Printer.FontBold = False
            PosLinea = PosLinea + 0.5
         End If
         SumaDebe = SumaDebe + .fields("Suma_Debitos")
         SumaHaber = SumaHaber + .fields("Suma_Creditos")
         'Diferencia = Diferencia + .Fields("Diferencia")
         'Sumatoria = Sumatoria + .Fields("Presupuesto")
        .MoveNext
      Loop
      If MN_ME = False Then
         PosLinea = PosLinea + 0.05
         Printer.Line (InicioX, PosLinea)-(Ancho(CantCampos), PosLinea), Negro
         PosLinea = PosLinea + 0.1
         PrinterVariables Ancho(1) + 1, PosLinea, "T O T A L E S"
         PrinterVariables Ancho(2), PosLinea, SumaDebe
         PrinterVariables Ancho(3) - 0.5, PosLinea, SumaHaber
         PrinterVariables Ancho(5), PosLinea, Sumatoria
         PrinterVariables Ancho(6), PosLinea, Diferencia
      End If
  End If
 End With
PosLinea = PosLinea + 1
Printer.Line (InicioX, PosLinea)-(Ancho(CantCampos), PosLinea), Negro
PosLinea = PosLinea + 0.05
Printer.Line (InicioX, PosLinea)-(Ancho(CantCampos), PosLinea), Negro
PosLinea = PosLinea + 0.2
If PosLinea + 3 > LimiteAlto Then
   Printer.NewPage
   PosLinea = 0
   EncabezadoDataSubCta Datas2, MN_ME, ImpResum
   Printer.FontSize = SizeLetra
End If
Printer.FontSize = 10
Printer.FontUnderline = True
If MN_ME Then
   PrinterTexto Ancho(4) + 1, PosLinea, "SUCRES"
   PrinterTexto Ancho(6), PosLinea, "DOLARES"
Else
   PrinterTexto Ancho(4) + 1, PosLinea, "DOLARES"
End If
Printer.FontUnderline = False
PosLinea = PosLinea + 0.5
PrinterTexto Ancho(1) + 2, PosLinea, "UTILIDAD(Perdida) TOTAL"
If MN_ME Then
   PrinterVariables Ancho(4), PosLinea, Haber - Debe
   PrinterVariables Ancho(6) - 1, PosLinea, Parcial_ME
   PosLinea = PosLinea + 0.5
   PrinterTexto Ancho(1) + 2, PosLinea, "UTILIDAD(Perdida) PAX"
   PrinterVariables Ancho(4), PosLinea, (Haber - Debe) / Datas2.Recordset.fields("PAX")
   PrinterVariables Ancho(6) - 1, PosLinea, Parcial_ME / Datas2.Recordset.fields("PAX")
   PosLinea = PosLinea + 0.5
   PrinterTexto Ancho(1) + 2, PosLinea, "GASTOS PAX"
   PrinterVariables Ancho(4), PosLinea, Debe / Datas2.Recordset.fields("PAX")
   PrinterVariables Ancho(6) - 1, PosLinea, Valor / Datas2.Recordset.fields("PAX")
Else
   PrinterVariables Ancho(4), PosLinea, Haber - Debe
End If
PosLinea = PosLinea + 0.5
Printer.Line (InicioX, PosLinea)-(Ancho(CantCampos), PosLinea), Negro
PosLinea = PosLinea + 0.05
Printer.Line (InicioX, PosLinea)-(Ancho(CantCampos), PosLinea), Negro
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

Public Sub ImprimirBusqueda(Datas As Adodc, _
                            FormaImp As Byte, _
                            SizeLetra As Integer)
On Error GoTo Errorhandler
RatonReloj
InicioX = 0.5: InicioY = 0
'Escala_Centimetro FormaImp, TipoTimes, SizeLetra
DataAnchoCampos InicioX, Datas, SizeLetra, TipoTimes, FormaImp
Ancho(0) = 0.5   'T
Ancho(1) = 0.9   'Fecha
Ancho(2) = 2.6   'TP
Ancho(3) = 3.3   'Numero
Ancho(4) = 4.7   'Cheq_Dep
Ancho(5) = 6.2   'Beneficiario
Ancho(6) = 11    'RUC_CI
Ancho(7) = 13.5  'Usuario
Ancho(8) = 16.5  'Cotizacion
Ancho(9) = 18.5  'Parcial
Ancho(10) = 21   'Debe
Ancho(11) = 23.5 'Haber
Ancho(12) = 26   '*Final
Pagina = 1
'Iniciamos la impresion
Printer.FontBold = False
With Datas.Recordset
     .MoveFirst
      EncabezadoData Datas
      Printer.FontSize = SizeLetra
      Do While Not .EOF
         PrinterAllFields CantCampos, PosLinea, Datas, True, False
         PosLinea = PosLinea + 0.4
         If PosLinea > LimiteAlto Then
            Printer.Line (Ancho(0), PosLinea)-(Ancho(CantCampos), PosLinea), Negro
            Printer.NewPage
            EncabezadoData Datas
            Printer.FontSize = SizeLetra
         End If
        .MoveNext
      Loop
End With
Printer.Line (InicioX, PosLinea)-(Ancho(CantCampos), PosLinea), Negro
PosLinea = PosLinea + 0.05
Printer.Line (InicioX, PosLinea)-(Ancho(CantCampos), PosLinea), Negro
RatonNormal
MensajeEncabData = ""
Printer.EndDoc
Exit Sub
Errorhandler:
    RatonNormal
    ErrorDeImpresion
    Exit Sub
End Sub

Public Sub ImprimirSaldosBancos(Datas As Adodc, _
                                FormaImp As Byte)
On Error GoTo Errorhandler
RatonReloj
InicioX = 0.5: InicioY = 0
Saldo = 0
'Escala_Centimetro FormaImp, TipoTimes, 9
DataAnchoCampos InicioX, Datas, 8, TipoTimes, FormaImp
Ancho(0) = 0.5   'TC
Ancho(1) = 1.1   'Cuenta
Ancho(2) = 9.2   'ME
Ancho(3) = 9.8   'Saldo Anterior
Ancho(4) = 12.1  'Debitos
Ancho(5) = 14.4  'Creditos
Ancho(6) = 16.7  'Saldo Actual
Ancho(7) = 19
Pagina = 1
'Iniciamos la impresion
Printer.FontBold = False
With Datas.Recordset
     .MoveFirst
      Moneda_US = .fields("ME")
      TipoProc = .fields("TC")
      EncabezadoData Datas
      Printer.FontSize = 8
      Do While Not .EOF
         If Moneda_US <> .fields("ME") Or TipoProc <> .fields("TC") Then
            Printer.FontBold = True
            'PosLinea = PosLinea + 0.1
            Printer.Line (InicioX, PosLinea)-(Ancho(CantCampos), PosLinea), Negro
            PosLinea = PosLinea + 0.1
            PrinterVariables Ancho(5), PosLinea, "T O T A L"
            PrinterVariables Ancho(6), PosLinea, Saldo
            PosLinea = PosLinea + 0.6
            Printer.Line (InicioX, PosLinea)-(Ancho(CantCampos), PosLinea), Negro
            PosLinea = PosLinea + 0.1
            Moneda_US = .fields("ME")
            TipoProc = .fields("TC")
            Saldo = 0
         End If
         Printer.FontBold = False
         PrinterFields Ancho(1), PosLinea, .fields("Cuenta"), False
         Cadena = " "
         If .fields("ME") Then Cadena = "Si"
         PrinterTexto Ancho(2), PosLinea, Cadena
         PrinterFields Ancho(3), PosLinea, .fields("Saldo_Anterior"), False
         PrinterFields Ancho(4), PosLinea, .fields("Debitos"), False
         PrinterFields Ancho(5), PosLinea, .fields("Creditos"), False
         PrinterFields Ancho(6), PosLinea, .fields("Saldo_Actual"), False
         For I = 0 To CantCampos
            Printer.Line (Ancho(I), PosLinea - 0.1)-(Ancho(I), PosLinea + 0.4), Negro
         Next
         PosLinea = PosLinea + 0.35
         If PosLinea >= LimiteAlto Then
            Printer.Line (Ancho(0), PosLinea)-(Ancho(CantCampos), PosLinea), Negro
            Printer.NewPage
            EncabezadoData Datas
            Printer.FontSize = 8
         End If
         Saldo = Saldo + .fields("Saldo_Actual")
        .MoveNext
      Loop
End With
Printer.FontBold = True
'PosLinea = PosLinea + 0.1
Printer.Line (InicioX, PosLinea)-(Ancho(CantCampos), PosLinea), Negro
PosLinea = PosLinea + 0.1
PrinterVariables Ancho(5), PosLinea, "T O T A L"
PrinterVariables Ancho(6), PosLinea, Saldo
PosLinea = PosLinea + 0.5
Printer.Line (InicioX, PosLinea)-(Ancho(CantCampos), PosLinea), Negro
PosLinea = PosLinea + 0.05
Printer.Line (InicioX, PosLinea)-(Ancho(CantCampos), PosLinea), Negro
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

Public Sub ImprimirGastosCaja(Datas As Adodc, _
                              FormaImp As Byte)
On Error GoTo Errorhandler
RatonReloj
InicioX = 0.5: InicioY = 0
Saldo = 0
'Escala_Centimetro FormaImp, TipoTimes, 9
DataAnchoCampos InicioX, Datas, 8, TipoTimes, FormaImp, True
'''Ancho(0) = 0.5
'''Ancho(1) = 0.5
'''Ancho(2) = 4.9
'''Ancho(3) = 5.5
'''Ancho(4) = 7.5
'''Ancho(5) = 9.5
'''Ancho(6) = 11.5
'''Ancho(7) = 13.5
'''Ancho(8) = 15.5
'''Ancho(9) = 17.5
'''Ancho(10) = 19.5
Pagina = 1
'Iniciamos la impresion
Printer.FontBold = False
With Datas.Recordset
 If .RecordCount > 0 Then
     .MoveFirst
      'Moneda_US = .Fields("ME")
      TipoProc = .fields("TC")
      EncabezadoData Datas
      Printer.FontSize = 8
      For I = 1 To 7
          TotalDia(I) = 0
      Next I
      Do While Not .EOF
         If TipoProc <> .fields("TC") Then
            'PosLinea = PosLinea + 0.1
            Printer.Line (InicioX, PosLinea)-(Ancho(CantCampos), PosLinea), Negro
            PosLinea = PosLinea + 0.1
            PrinterVariables Ancho(1) + 2.5, PosLinea, "T O T A L E S"
            PrinterVariables Ancho(9) - 0.7, PosLinea, Saldo
            For NoDias = 2 To 7
                PrinterVariables Ancho(1 + NoDias), PosLinea, TotalDia(NoDias)
            Next NoDias
            PosLinea = PosLinea + 0.6
            Printer.Line (InicioX, PosLinea)-(Ancho(CantCampos), PosLinea), Negro
            PosLinea = PosLinea + 0.1
            'Moneda_US = .Fields("ME")
            TipoProc = .fields("TC")
            For I = 1 To 7
                TotalDia(I) = 0
            Next I
            Saldo = 0
         End If
         PrinterFields Ancho(1), PosLinea, .fields("Detalle")
         'PrinterFields Ancho(2), PosLinea, .Fields("ME")
         For NoDias = 2 To 7
             PrinterFields Ancho(1 + NoDias), PosLinea, .fields(DiasLetras(NoDias))
             TotalDia(NoDias) = TotalDia(NoDias) + .fields(DiasLetras(NoDias))
         Next NoDias
         PrinterFields Ancho(9), PosLinea, .fields("Total")
         For I = 0 To 10
            Printer.Line (Ancho(I), PosLinea - 0.1)-(Ancho(I), PosLinea + 0.4), Negro
         Next
         PosLinea = PosLinea + 0.4
         If PosLinea > LimiteAlto Then
            Printer.Line (Ancho(0), PosLinea)-(Ancho(10), PosLinea), Negro
            Printer.NewPage
            EncabezadoData Datas
            Printer.FontSize = 8
         End If
         Saldo = Saldo + .fields("Total")
        .MoveNext
      Loop
  End If
End With
'PosLinea = PosLinea + 0.1
Printer.Line (Ancho(0), PosLinea)-(Ancho(10), PosLinea), Negro
PosLinea = PosLinea + 0.1
PrinterVariables Ancho(1) + 2.5, PosLinea, "T O T A L E S"
PrinterVariables Ancho(9) - 0.5, PosLinea, Saldo
For NoDias = 2 To 7
    PrinterVariables Ancho(1 + NoDias), PosLinea, TotalDia(NoDias)
Next NoDias
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

Public Sub Tipo_Flujo()
  Select Case TipoProc
        Case "CJ"
             If Moneda_US Then
                PrinterTexto Ancho(1), PosLinea, "CAJA MONEDA EXTRANJERA"
             Else
                PrinterTexto Ancho(1), PosLinea, "CAJA MONEDA NACIONAL"
             End If
        Case "BA"
             If Moneda_US Then
                PrinterTexto Ancho(1), PosLinea, "BANCO MONEDA EXTRANJERA"
             Else
                PrinterTexto Ancho(1), PosLinea, "BANCO MONEDA NACIONAL"
             End If
        Case "C", "CS"
             If Moneda_US Then
                PrinterTexto Ancho(1), PosLinea, "CUENTAS POR COBRAR MONEDA EXTRANJERA"
             Else
                PrinterTexto Ancho(1), PosLinea, "CUENTAS POR COBRAR MONEDA NACIONAL"
             End If
        Case "P", "PS"
             If Moneda_US Then
                PrinterTexto Ancho(1), PosLinea, "CUENTAS POR PAGAR MONEDA EXTRANJERA"
             Else
                PrinterTexto Ancho(1), PosLinea, "CUENTAS POR PAGAR MONEDA NACIONAL"
             End If
        Case "G"
             If Moneda_US Then
                PrinterTexto Ancho(1), PosLinea, "CUENTAS DE GASTOS MONEDA EXTRANJERA"
             Else
                PrinterTexto Ancho(1), PosLinea, "CUENTAS DE GASTOS MONEDA NACIONAL"
             End If
        Case "I"
             If Moneda_US Then
                PrinterTexto Ancho(1), PosLinea, "CUENTAS DE INGRESO MONEDA EXTRANJERA"
             Else
                PrinterTexto Ancho(1), PosLinea, "CUENTAS DE INGRESO MONEDA NACIONAL"
             End If
        Case "RF"
             If Moneda_US Then
                PrinterTexto Ancho(1), PosLinea, "CUENTAS DE RETENCION FUENTE MONEDA EXTRANJERA"
             Else
                PrinterTexto Ancho(1), PosLinea, "CUENTAS DE RETENCION FUENTE MONEDA NACIONAL"
             End If
        Case "RI"
             If Moneda_US Then
                PrinterTexto Ancho(1), PosLinea, "CUENTAS DE RETENCION IVA MONEDA EXTRANJERA"
             Else
                PrinterTexto Ancho(1), PosLinea, "CUENTAS DE RETENCION IVA MONEDA NACIONAL"
             End If
  End Select
End Sub

Public Sub Imprimir_Saldos_Flujo(Datas As Adodc, _
                                 FormaImp As Byte)
 
On Error GoTo Errorhandler
Mensajes = "Seguro de Imprimir en:" & vbCrLf & Printer.DeviceName & "?"
Titulo = "IMPRESION"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
RatonReloj
InicioX = 0.5: InicioY = 0
Saldo = 0
Diferencia = 0
TotalActivo = 0
TotalPasivo = 0
TipoLetra = TipoArialNarrow
DataAnchoCampos InicioX, Datas, 8, TipoLetra, FormaImp
Ancho(0) = 0.5
Ancho(1) = 1.1
Ancho(2) = 7.2
Ancho(3) = 8
Ancho(4) = 10
Ancho(5) = 12
Ancho(6) = 14
Ancho(7) = 16
Ancho(8) = 18
Ancho(9) = 20
Pagina = 1
'Iniciamos la impresion
Printer.FontBold = False
With Datas.Recordset
     .MoveFirst
      Moneda_US = .fields("ME")
      TipoProc = .fields("TC")
      Encabezados
      Printer.FontSize = 8
      Printer.FontName = TipoLetra
      Printer.FontBold = True
      PrinterTexto Ancho(0), PosLinea, "ACTIVOS"
      Printer.FontBold = False
      PrinterAllFields CantCampos, PosLinea, Datas, False, True
      PosLinea = PosLinea + 0.5
      Do While Not .EOF
         If Moneda_US <> .fields("ME") Or TipoProc <> .fields("TC") Then
            Printer.FontBold = True
            Tipo_Flujo
            PrinterVariables Ancho(5), PosLinea, "T O T A L"
            PrinterVariables Ancho(6), PosLinea, Saldo
            Moneda_US = .fields("ME")
            TipoProc = .fields("TC")
            PosLinea = PosLinea + 0.6
            Saldo = 0
         End If
         Printer.FontBold = False
         PrinterAllFields CantCampos, PosLinea, Datas, False, False
         PosLinea = PosLinea + 0.4
         If PosLinea >= LimiteAlto Then
            Printer.NewPage
            Encabezados
            Printer.FontSize = 8
            Printer.FontName = TipoLetra
         End If
         Saldo = Saldo + .fields("Saldo_Actual")
        .MoveNext
      Loop
End With
Printer.FontBold = True
Tipo_Flujo
PrinterTexto Ancho(5), PosLinea, "T O T A L"
PrinterVariables Ancho(6), PosLinea, Saldo
PosLinea = PosLinea + 0.5
'''If TipoProc = "G" Then
'''   PrinterTexto Ancho(1), PosLinea, "TOTAL DIFERENCIA PRESUPUESTO"
'''   PrinterVariables Ancho(5), PosLinea, "T O T A L"
'''   PrinterVariables Ancho(6), PosLinea, Saldo - Sumatoria
'''   PosLinea = PosLinea + 0.5
'''End If
Printer.Line (InicioX, PosLinea)-(Ancho(CantCampos), PosLinea), Negro
PosLinea = PosLinea + 0.05
Printer.Line (InicioX, PosLinea)-(Ancho(CantCampos), PosLinea), Negro
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

Public Sub ImprimirSaldosSubCtas(Datas As Adodc, _
                                 FormaImp As Byte, _
                                 Optional NivelB As Boolean)
On Error GoTo Errorhandler
RatonReloj
InicioX = 1: InicioY = 0
Saldo = 0: Saldo_ME = 0: Valor = 0
Total = 0: Total_ME = 0: Valor_ME = 0
'Escala_Centimetro FormaImp, TipoTimes, 9
DataAnchoCampos InicioX, Datas, 10, TipoTimes, FormaImp
Ancho(0) = 1 'TC
Ancho(1) = 1.8 'Cuenta
Ancho(2) = 10 'Saldo_ME
Ancho(3) = 13 'Saldo_ME
Ancho(4) = 16 'Saldo
Ancho(5) = 19 'Final
Pagina = 1
'Iniciamos la impresion
Printer.FontBold = False
With Datas.Recordset
     .MoveFirst
      EncabezadoData Datas
      Printer.FontSize = 10
      Codigo = CodigoIzqGuion(.fields("Cuenta"))
      Do While Not .EOF
         If Codigo <> CodigoIzqGuion(.fields("Cuenta")) And NivelB Then
            PosLinea = PosLinea + 0.1
            Printer.Line (Ancho(0), PosLinea)-(Ancho(CantCampos), PosLinea), Negro
            PosLinea = PosLinea + 0.1
            PrinterVariables Ancho(1) + 1, PosLinea, "SUBTOTAL " & Codigo
            PrinterVariables Ancho(2), PosLinea, Total_ME
            PrinterVariables Ancho(3), PosLinea, Total
            PrinterVariables Ancho(4), PosLinea, Valor_ME
            PosLinea = PosLinea + 0.6
            Printer.Line (Ancho(0), PosLinea)-(Ancho(CantCampos), PosLinea), Negro
            PosLinea = PosLinea + 0.1
            Codigo = CodigoIzqGuion(.fields("Cuenta"))
            Total = 0: Total_ME = 0:  Valor_ME = 0
         End If
         PrinterAllFields CantCampos, PosLinea, Datas, True, False
         PosLinea = PosLinea + 0.4
         If PosLinea > LimiteAlto Then
            Printer.Line (Ancho(0), PosLinea)-(Ancho(CantCampos), PosLinea), Negro
            Printer.NewPage
            EncabezadoData Datas
            Printer.FontSize = 10
         End If
        'Total
         Saldo_ME = Saldo_ME + .fields("Saldo_ME")
         Saldo = Saldo + .fields("Saldo_MN")
         Valor = Valor + .fields("Saldo")
        'SubTotal
         Total_ME = Total_ME + .fields("Saldo_ME")
         Total = Total + .fields("Saldo_MN")
         Valor_ME = Valor_ME + .fields("Saldo")
        .MoveNext
      Loop
End With
PosLinea = PosLinea + 0.1
If NivelB Then
Printer.Line (Ancho(0), PosLinea)-(Ancho(CantCampos), PosLinea), Negro
PosLinea = PosLinea + 0.1
PrinterVariables Ancho(1) + 1, PosLinea, "SUBTOTAL " & Codigo
PrinterVariables Ancho(2), PosLinea, Total_ME
PrinterVariables Ancho(3), PosLinea, Total
PrinterVariables Ancho(4), PosLinea, Valor_ME
PosLinea = PosLinea + 0.6
End If
Printer.Line (Ancho(0), PosLinea)-(Ancho(CantCampos), PosLinea), Negro
PosLinea = PosLinea + 0.1
PrinterVariables Ancho(1) + 3, PosLinea, "T O T A L"
PrinterVariables Ancho(2), PosLinea, Saldo_ME
PrinterVariables Ancho(3), PosLinea, Saldo
PrinterVariables Ancho(4), PosLinea, Valor
PosLinea = PosLinea + 0.5
Printer.Line (InicioX, PosLinea)-(Ancho(CantCampos), PosLinea), Negro
PosLinea = PosLinea + 0.1

Printer.Line (InicioX, PosLinea)-(Ancho(CantCampos), PosLinea), Negro
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

Public Sub Imprimir_Saldos_SubCtas_Vence(Datas As Adodc, _
                                         FormaImp As Byte, _
                                         Optional NivelB As Boolean, _
                                         Optional Vertical As Boolean)
SizeLetra = 6
Mensajes = "Seguro de Imprimir en:" & vbCrLf & Printer.DeviceName & "?"
Titulo = "IMPRESION"
Bandera = True
If Vertical Then Orientacion_Pagina = 1 Else Orientacion_Pagina = 2
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
RatonReloj
InicioX = 1: InicioY = 0
Saldo = 0: Saldo_ME = 0: Valor = 0
Total = 0: Total_ME = 0: Valor_ME = 0
Debe = 0: Haber = 0
Debitos = 0: Creditos = 0
Ln_No = 1
DataAnchoCampos InicioX, Datas, 6, TipoVerdana, FormaImp, True
Ancho(0) = 0.5    'Cuenta
Ancho(1) = 0.7   'Cliente
Ancho(2) = 6.7    'Telefono
Ancho(3) = 8.2    'Serie-Factura
Ancho(4) = 10.6   'TP
Ancho(5) = 11.1   'Numero
Ancho(6) = 12.4   'Fecha
Ancho(7) = 13.9   'Fecha_V
If Vertical Then
   Ancho(8) = 15.4   'Total
   Ancho(9) = 17     'Abonos
   Ancho(10) = 18.6  'Saldo
   Ancho(11) = 20.2  '
   CantCampos = 11
Else
   Ancho(8) = 15.7   'Beneficiario
   Ancho(9) = 23.7   'Total
   Ancho(10) = 25.3  'Abonos
   Ancho(11) = 26.9  'Saldo
   Ancho(12) = 28.5  '
   CantCampos = 12
End If
Pagina = 1
'Iniciamos la impresion
Printer.FontBold = False
 With Datas.Recordset
  If .RecordCount > 0 Then
      Progreso_Barra.Mensaje_Box = "Imprimiendo/Generando Reporte"
      Progreso_Iniciar
      Progreso_Barra.Valor_Maximo = .RecordCount
 
     .MoveFirst
      Encabezado_SubCta_Venc Vertical
      Cuenta = UCaseStrg(.fields("Cuenta"))
      SubCta = ULCase(.fields("Cliente"))
      Si_No = True
      Printer.FontName = TipoVerdana
      Printer.FontItalic = False
      Printer.FontBold = False
      Printer.FontSize = SizeLetra
      PrinterTexto Ancho(0), PosLinea, Cuenta & ":"
      PosLinea = PosLinea + 0.3
      Printer.Line (Ancho(0), PosLinea)-(Ancho(CantCampos), PosLinea), Negro
      PosLinea = PosLinea + 0.1
      Do While Not .EOF
         If Cuenta <> UCaseStrg(.fields("Cuenta")) And NivelB Then
            PosLinea = PosLinea + 0.1
            Printer.Line (Ancho(0), PosLinea)-(Ancho(CantCampos), PosLinea), Negro
            PosLinea = PosLinea + 0.05
            If Vertical Then
               PrinterVariables Ancho(2), PosLinea, "Subtotal de " & SubCta
               PrinterVariables Ancho(8), PosLinea, Debitos
               PrinterVariables Ancho(9), PosLinea, Debitos - Creditos
               PrinterVariables Ancho(10), PosLinea, Creditos
            Else
               PrinterVariables Ancho(1), PosLinea, "SUBTOTAL"
               PrinterVariables Ancho(8), PosLinea, Debitos
               PrinterVariables Ancho(9), PosLinea, Debitos - Creditos
               PrinterVariables Ancho(10), PosLinea, Creditos
            End If
            PosLinea = PosLinea + 0.4
            
            'PosLinea = PosLinea + 0.1
            'Printer.Line (Ancho(0), PosLinea)-(Ancho(CantCampos), PosLinea), Negro
            'PosLinea = PosLinea + 0.1
            If Vertical Then
               PrinterVariables Ancho(2), PosLinea, "SUBTOTAL " & Cuenta
               PrinterVariables Ancho(8), PosLinea, Debe
               PrinterVariables Ancho(9), PosLinea, Debe - Haber
               PrinterVariables Ancho(10), PosLinea, Haber
            Else
               PrinterVariables Ancho(7), PosLinea, "SUBTOTAL"
               PrinterVariables Ancho(8), PosLinea, Debe
               PrinterVariables Ancho(9), PosLinea, Debe - Haber
               PrinterVariables Ancho(10), PosLinea, Haber
            End If
            PosLinea = PosLinea + 0.4
            Printer.Line (Ancho(0), PosLinea)-(Ancho(CantCampos), PosLinea), Negro
            PosLinea = PosLinea + 0.05
            Debe = 0: Haber = 0
            Debitos = 0: Creditos = 0
            Debe_ME = 0: Haber_ME = 0
            Cuenta = UCaseStrg(.fields("Cuenta"))
            SubCta = ULCase(.fields("Cliente"))
            Si_No = True
            PrinterTexto Ancho(0), PosLinea, Cuenta & ":"
            PosLinea = PosLinea + 0.3
            Printer.Line (Ancho(0), PosLinea)-(Ancho(CantCampos), PosLinea), Negro
            PosLinea = PosLinea + 0.05
            Bandera = True
            Ln_No = 1
         End If
         If SubCta <> ULCase(.fields("Cliente")) Then
            PosLinea = PosLinea + 0.1
            Printer.Line (Ancho(0), PosLinea)-(Ancho(CantCampos), PosLinea), Negro
            PosLinea = PosLinea + 0.05
            If Vertical Then
               PrinterVariables Ancho(2), PosLinea, "Subtotal de " & SubCta
               PrinterVariables Ancho(8), PosLinea, Debitos
               PrinterVariables Ancho(9), PosLinea, Debitos - Creditos
               PrinterVariables Ancho(10), PosLinea, Creditos
            Else
               PrinterVariables Ancho(1), PosLinea, "SUBTOTAL"
               PrinterVariables Ancho(8), PosLinea, Debitos
               PrinterVariables Ancho(9), PosLinea, Debitos - Creditos
               PrinterVariables Ancho(10), PosLinea, Creditos
            End If
            Debitos = 0: Creditos = 0
            Cuenta = UCaseStrg(.fields("Cuenta"))
            SubCta = ULCase(.fields("Cliente"))
            PosLinea = PosLinea + 0.4
            Printer.Line (Ancho(0), PosLinea)-(Ancho(CantCampos), PosLinea), Negro
            PosLinea = PosLinea + 0.05
            Bandera = True
            Ln_No = 1
         End If
         
         'PrinterTexto Ancho(1), PosLinea, ULCase(.fields(1))
         If Bandera Then
            PrinterTexto Ancho(1), PosLinea, ULCase(SubCta)
            Bandera = False
         End If
         
         'PrinterTexto Ancho(2) - 0.5, PosLinea, Format(Ln_No, "00")

         PrinterFields Ancho(2), PosLinea, .fields(2), True
         PrinterTexto Ancho(3), PosLinea, .fields(3) & "-" & Format(.fields(4), "000000000")
         PrinterFields Ancho(4), PosLinea, .fields(5), True
         PrinterTexto Ancho(5), PosLinea, Format(.fields(6), "00000000")
         PrinterFields Ancho(6), PosLinea, .fields(7), True
         PrinterFields Ancho(7), PosLinea, .fields(8), True
         If Vertical Then
            PrinterFields Ancho(8), PosLinea, .fields(9), True
            PrinterFields Ancho(9), PosLinea, .fields(10), True
            PrinterFields Ancho(10), PosLinea, .fields(11), True
         Else
            PrinterFields Ancho(8), PosLinea, .fields(9), True
            PrinterFields Ancho(9), PosLinea, .fields(10), True
            PrinterFields Ancho(10), PosLinea, .fields(11), True
            PrinterFields Ancho(11), PosLinea, .fields(12), True
         End If
         PosLinea = PosLinea + 0.3
         For I = 0 To CantCampos
             If I <> 1 Then Printer.Line (Ancho(I), PosLinea - 0.3)-(Ancho(I), PosLinea), Negro
         Next I
        'Total
         Total = Total + .fields("Total")
         Saldo = Saldo + .fields("Saldo")
        'SubTotal
         Debe = Debe + .fields("Total")
         Haber = Haber + .fields("Saldo")
        'SubTotal SubModulo
         Debitos = Debitos + .fields("Total")
         Creditos = Creditos + .fields("Saldo")
         Ln_No = Ln_No + 1
         If PosLinea >= (LimiteAlto - 0.3) Then
            Printer.Line (Ancho(0), PosLinea)-(Ancho(CantCampos), PosLinea), Negro
            Printer.NewPage
            Encabezado_SubCta_Venc Vertical
            Printer.FontName = TipoVerdana
            Printer.FontItalic = False
            Printer.FontBold = False
            Printer.FontSize = SizeLetra
            PrinterTexto Ancho(0), PosLinea, Cuenta & ":"
            PosLinea = PosLinea + 0.3
            Printer.Line (Ancho(0), PosLinea)-(Ancho(CantCampos), PosLinea), Negro
            PosLinea = PosLinea + 0.05
            Bandera = True
         End If
         Progreso_Barra.Mensaje_Box = "Progreso => " & Cuenta & ": " & SubCta
         Progreso_Esperar
        .MoveNext
      Loop
     .MoveFirst
 End If
End With
If NivelB Then
   PosLinea = PosLinea + 0.1
   Printer.Line (Ancho(0), PosLinea)-(Ancho(CantCampos), PosLinea), Negro
   PosLinea = PosLinea + 0.05
   PrinterVariables Ancho(2), PosLinea, "Subtotal de " & SubCta
   PrinterVariables Ancho(8), PosLinea, Debitos
   PrinterVariables Ancho(9), PosLinea, Debitos - Creditos
   PrinterVariables Ancho(10), PosLinea, Creditos
   PosLinea = PosLinea + 0.4
   Printer.Line (Ancho(0), PosLinea)-(Ancho(CantCampos), PosLinea), Negro
   PosLinea = PosLinea + 0.05
   If Vertical Then
      PrinterVariables Ancho(2), PosLinea, "SUBTOTAL " & Cuenta
      PrinterVariables Ancho(8), PosLinea, Debe
      PrinterVariables Ancho(9), PosLinea, Debe - Haber
      PrinterVariables Ancho(10), PosLinea, Haber
   Else
      PrinterVariables Ancho(7), PosLinea, "SUBTOTAL"
      PrinterVariables Ancho(8), PosLinea, Debe
      PrinterVariables Ancho(9), PosLinea, Debe - Haber
      PrinterVariables Ancho(10), PosLinea, Haber
   End If
   PosLinea = PosLinea + 0.4
   Printer.Line (Ancho(0), PosLinea)-(Ancho(CantCampos), PosLinea), Negro
   PosLinea = PosLinea + 0.05
End If
Printer.Line (Ancho(0), PosLinea)-(Ancho(CantCampos), PosLinea), Negro
PosLinea = PosLinea + 0.1
If Vertical Then
   PrinterVariables Ancho(7), PosLinea, "T O T A L"
   PrinterVariables Ancho(8), PosLinea, Total
   PrinterVariables Ancho(9), PosLinea, Total - Saldo
   PrinterVariables Ancho(10), PosLinea, Saldo
Else
   PrinterVariables Ancho(6), PosLinea, "T O T A L"
   PrinterVariables Ancho(7), PosLinea, Total
   PrinterVariables Ancho(8), PosLinea, Total - Saldo
   PrinterVariables Ancho(9), PosLinea, Saldo
End If
'PosLinea = PosLinea + 0.4
'Printer.Line (Ancho(0), PosLinea)-(Ancho(CantCampos), PosLinea), Negro
'PosLinea = PosLinea + 0.05
'Printer.Line (Ancho(0), PosLinea)-(Ancho(CantCampos), PosLinea), Negro
RatonNormal
Progreso_Final
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

Public Sub Imprimir_Saldos_SubCtas_IE(Datas As Adodc, _
                                      FormaImp As Byte, _
                                      Optional NivelB As Boolean, _
                                      Optional Benef As Boolean)
SizeLetra = 6
Mensajes = "Seguro de Imprimir en:" & vbCrLf & Printer.DeviceName & "?"
Titulo = "IMPRESION"
Bandera = False
Orientacion_Pagina = 1
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
RatonReloj
InicioX = 1: InicioY = 0
Saldo = 0: Saldo_ME = 0: Valor = 0
Total = 0: Total_ME = 0: Valor_ME = 0
DataAnchoCampos InicioX, Datas, 6, TipoVerdana, FormaImp
Ancho(0) = 1.5    'Cuenta
If Benef Then
   Ancho(1) = 6.5  'SubCuenta
   Ancho(2) = 11.5 'Detalle Auxiliar
   Ancho(3) = 18   'Total
   Ancho(4) = 20   '
   CantCampos = 4
Else
   Ancho(1) = 8    'SubCuenta
   Ancho(2) = 18   'Total
   Ancho(3) = 20   '
   CantCampos = 3
End If
Pagina = 1
'Iniciamos la impresion
Printer.FontBold = False
Total = 0: Saldo = 0: Debe = 0: Haber = 0
With Datas.Recordset
 If .RecordCount > 0 Then
     .MoveFirst
      Encabezado_SubCta_IE Benef
      Cuenta = .fields("Cuenta")
      SubCta = .fields("Sub_Modulos")
      Printer.FontName = TipoVerdana
      Printer.FontItalic = False
      Printer.FontBold = False
      Printer.FontSize = SizeLetra
      Printer.Line (Ancho(0), PosLinea)-(Ancho(0), PosLinea + 0.3), Negro
      PrinterTexto Ancho(0) + 0.1, PosLinea, Cuenta
      Printer.Line (Ancho(1), PosLinea)-(Ancho(1), PosLinea + 0.3), Negro
      PrinterTexto Ancho(1), PosLinea, SubCta
      Do While Not .EOF
         If Cuenta <> .fields("Cuenta") And SubCta <> .fields("Sub_Modulos") Then
            PosLinea = PosLinea + 0.05
            Printer.Line (Ancho(0), PosLinea)-(Ancho(CantCampos), PosLinea), Negro
            PosLinea = PosLinea + 0.05
            If Benef Then
               PrinterVariables Ancho(2), PosLinea, "SUBTOTAL DETALLE"
               PrinterVariables Ancho(3), PosLinea, Haber
            Else
               PrinterVariables Ancho(1), PosLinea, "SUBTOTAL DETALLE"
               PrinterVariables Ancho(2), PosLinea, Haber
            End If
            PosLinea = PosLinea + 0.4
            If Benef Then
               PrinterVariables Ancho(2), PosLinea, "SUBTOTAL"
               PrinterVariables Ancho(3), PosLinea, Debe
            Else
               PrinterVariables Ancho(1), PosLinea, "SUBTOTAL"
               PrinterVariables Ancho(2), PosLinea, Debe
            End If
            PosLinea = PosLinea + 0.4
            Printer.Line (Ancho(0), PosLinea)-(Ancho(CantCampos), PosLinea), Negro
            PosLinea = PosLinea + 0.1
            Debe = 0
            Haber = 0
            Cuenta = .fields("Cuenta")
            SubCta = .fields("Sub_Modulos")
            PrinterTexto Ancho(0), PosLinea, Cuenta
            PrinterTexto Ancho(1), PosLinea, SubCta
         End If
         If Cuenta <> .fields("Cuenta") Then
            PosLinea = PosLinea + 0.05
            Printer.Line (Ancho(0), PosLinea)-(Ancho(CantCampos), PosLinea), Negro
            PosLinea = PosLinea + 0.05
            If Benef Then
               PrinterVariables Ancho(2), PosLinea, "SUBTOTAL"
               PrinterVariables Ancho(3), PosLinea, Debe
            Else
               PrinterVariables Ancho(1), PosLinea, "SUBTOTAL"
               PrinterVariables Ancho(2), PosLinea, Debe
            End If
            PosLinea = PosLinea + 0.4
            Printer.Line (Ancho(0), PosLinea)-(Ancho(CantCampos), PosLinea), Negro
            PosLinea = PosLinea + 0.1
            Debe = 0
            Haber = 0
            Cuenta = .fields("Cuenta")
            SubCta = .fields("Sub_Modulos")
            PrinterTexto Ancho(0), PosLinea, Cuenta
            PrinterTexto Ancho(1), PosLinea, SubCta
         End If
         
         If SubCta <> .fields("Sub_Modulos") Then
            PosLinea = PosLinea + 0.05
            Printer.Line (Ancho(1), PosLinea)-(Ancho(CantCampos), PosLinea), Negro
            PosLinea = PosLinea + 0.05
            If Benef Then
               PrinterVariables Ancho(2), PosLinea, "SUBTOTAL DETALLE"
               PrinterVariables Ancho(3), PosLinea, Haber
            Else
               PrinterVariables Ancho(1), PosLinea, "SUBTOTAL DETALLE"
               PrinterVariables Ancho(2), PosLinea, Haber
            End If
            PosLinea = PosLinea + 0.4
            Printer.Line (Ancho(1), PosLinea)-(Ancho(CantCampos), PosLinea), Negro
            PosLinea = PosLinea + 0.1
            Haber = 0
            SubCta = .fields("Sub_Modulos")
            PrinterTexto Ancho(1), PosLinea, SubCta
         End If
         'PrinterTexto Ancho(1), PosLinea, ULCase(.Fields(1))
         
         If Benef Then
            PrinterFields Ancho(2), PosLinea, .fields(2), True
            PrinterFields Ancho(3), PosLinea, .fields(3), True
         Else
            PrinterFields Ancho(2), PosLinea, .fields(2), True
            'PrinterFields Ancho(3), PosLinea, .Fields(3), True
         End If
         For I = 1 To CantCampos
             Printer.Line (Ancho(I), PosLinea - 0.1)-(Ancho(I), PosLinea + 0.3), Negro
         Next I
         PosLinea = PosLinea + 0.3
         If PosLinea >= LimiteAlto Then
            Printer.Line (Ancho(0), PosLinea)-(Ancho(CantCampos), PosLinea), Negro
            Printer.NewPage
            Encabezado_SubCta_IE Benef
            Printer.FontName = TipoVerdana
            Printer.FontItalic = False
            Printer.FontBold = False
            Printer.FontSize = SizeLetra
            PrinterTexto Ancho(0), PosLinea, Cuenta
         End If
        'Total
         Total = Total + .fields("Total")
        'SubTotal
         Debe = Debe + .fields("Total")
        'SubTotal Detalle
         Haber = Haber + .fields("Total")
        .MoveNext
      Loop
     .MoveFirst
 End If
End With
            PosLinea = PosLinea + 0.05
            Printer.Line (Ancho(0), PosLinea)-(Ancho(CantCampos), PosLinea), Negro
            PosLinea = PosLinea + 0.05
            If Benef Then
               PrinterVariables Ancho(2), PosLinea, "SUBTOTAL"
               PrinterVariables Ancho(3), PosLinea, Debe
            Else
               PrinterVariables Ancho(1), PosLinea, "SUBTOTAL"
               PrinterVariables Ancho(2), PosLinea, Debe
            End If

''If NivelB Then
''   PosLinea = PosLinea + 0.05
''   Printer.Line (Ancho(0), PosLinea)-(Ancho(CantCampos), PosLinea), NEGRO
''   PosLinea = PosLinea + 0.05
''   If Vertical Then
''      PrinterVariables Ancho(2), PosLinea, "SUBTOTAL"
''      PrinterVariables Ancho(3), PosLinea, Debe
''   Else
''      PrinterVariables Ancho(1), PosLinea, "SUBTOTAL"
''      PrinterVariables Ancho(2), PosLinea, Debe
''   End If
   PosLinea = PosLinea + 0.4
   Printer.Line (Ancho(0), PosLinea)-(Ancho(CantCampos), PosLinea), Negro
   PosLinea = PosLinea + 0.05
''End If
Printer.Line (Ancho(0), PosLinea)-(Ancho(CantCampos), PosLinea), Negro
PosLinea = PosLinea + 0.05
If Benef Then
   PrinterVariables Ancho(2), PosLinea, "T O T A L"
   PrinterVariables Ancho(3), PosLinea, Total
Else
   PrinterVariables Ancho(1), PosLinea, "T O T A L"
   PrinterVariables Ancho(2), PosLinea, Total
End If
PosLinea = PosLinea + 0.4
Printer.Line (Ancho(0), PosLinea)-(Ancho(CantCampos), PosLinea), Negro
PosLinea = PosLinea + 0.05
Printer.Line (Ancho(0), PosLinea)-(Ancho(CantCampos), PosLinea), Negro
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

Public Sub Imprimir_Diario_General(DataT As Adodc, AdoSubCtas As Adodc, AdoConceptos As Adodc)
Dim ConceptoDe As String
 
Dim NoItem As Long
On Error GoTo Errorhandler
Mensajes = "Seguro de Imprimir en:" & vbCrLf & Printer.DeviceName & "?"
Titulo = "IMPRESION DIARIO GENERAL"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
RatonReloj
ReDim Ancho(6) As Single
InicioX = 0.5: InicioY = 0
PorteLetra = 9
TipoLetra = TipoArialNarrow
DataAnchoCampos InicioX, DataT, 8, TipoLetra, 1
LimiteAlto = LimiteAlto - 2
Pagina = 1
HoraSistema = Time
C = 5
Ancho(0) = 1      'Codigo
Ancho(1) = 3.4    'Cuenta
Ancho(2) = 13.1   'Parcial
Ancho(3) = 15.4   'Debe
Ancho(4) = 17.7   'Haber
Ancho(5) = 20     'Final
'============================================================
' Comenzamos a escribir en la impresora los encabezados
'============================================================
Printer.FontBold = False
With DataT.Recordset
    .MoveFirst
     Encabezado Ancho(0), Ancho(5)
     Printer.FontSize = PorteLetra
     Printer.FontName = TipoLetra
     Encab_Diario_General DataT
     TipoCta = .fields("TP")
     Numero = .fields("Numero")
     Mifecha = .fields("Fecha")
     ConceptoDe = .fields("Concepto")
     NoItem = .fields("Item")
     Debe = 0: Haber = 0
     Do While Not .EOF
        PorteLetra = 9
        If TipoCta <> .fields("TP") Or Numero <> .fields("Numero") Or Mifecha <> .fields("Fecha") Then
           Printer.FontSize = PorteLetra
           Printer.FontBold = True
           Printer.FontItalic = False
           PosLinea = PosLinea + 0.05
           Printer.Line (Ancho(0), PosLinea)-(Ancho(C), PosLinea), Negro
           PosLinea = PosLinea + 0.05
           PrinterVariables Ancho(0), PosLinea, Format$(NoItem, "000")
           PrinterVariables Ancho(2), PosLinea, "T O T A L E S"
           Printer.FontBold = False
           PrinterVariables Ancho(3), PosLinea, Debe
           PrinterVariables Ancho(4), PosLinea, Haber
           PosLinea = PosLinea + 0.6
           Printer.FontSize = PorteLetra
           Printer.FontName = TipoLetra
           Encab_Libro_Diario DataT
           If PosLinea >= LimiteAlto Then
              PosLinea = PosLinea + 0.05
              Printer.Line (Ancho(0), PosLinea)-(Ancho(C), PosLinea), Negro
              Printer.NewPage
              Encabezado Ancho(0), Ancho(5)
              Printer.FontSize = PorteLetra
              Printer.FontName = TipoLetra
              Encab_Diario_General DataT
           End If
           TipoCta = .fields("TP")
           Numero = .fields("Numero")
           Mifecha = .fields("Fecha")
           ConceptoDe = .fields("Concepto")
           NoItem = .fields("Item")
           Debe = 0: Haber = 0
        End If
        Printer.FontSize = PorteLetra
        Printer.FontName = TipoLetra
        Printer.FontBold = False
        Printer.FontItalic = False
        ID_Trans = .fields("ID")
        Cta_Aux = .fields("Cta")
        PrinterFields Ancho(0), PosLinea, .fields("Cta"), False
        PrinterFields Ancho(1), PosLinea, .fields("Cuenta"), False
        If .fields("Parcial_ME") <> 0 Then PrinterFields Ancho(2), PosLinea, .fields("Parcial_ME"), False
        PrinterFields Ancho(3), PosLinea, .fields("Debe"), False
        PrinterFields Ancho(4), PosLinea, .fields("Haber"), False
        Printer.FontSize = 7
        If Len(.fields("Detalle")) > 1 Then
           For I = 0 To C
               Printer.Line (Ancho(I), PosLinea - 0.1)-(Ancho(I), PosLinea + 0.3), Negro
           Next I
           PosLinea = PosLinea + 0.32
           PrinterFields Ancho(1), PosLinea, .fields("Detalle"), False
        End If
        Printer.FontItalic = True
        Evaluar = False
       'SubCtas
        If AdoSubCtas.Recordset.RecordCount > 0 Then
           AdoSubCtas.Recordset.MoveFirst
           AdoSubCtas.Recordset.Find ("TP = '" & TipoCta & "' ")
           If Not AdoSubCtas.Recordset.EOF Then
              AdoSubCtas.Recordset.Find ("Numero = " & Numero & " ")
              If Not AdoSubCtas.Recordset.EOF Then
                 AdoSubCtas.Recordset.Find ("Cta = '" & Cta_Aux & "' ")
                 If Not AdoSubCtas.Recordset.EOF Then
                    Si_No = True
                    Do While Not AdoSubCtas.Recordset.EOF And Si_No
                       If Cta_Aux = AdoSubCtas.Recordset.fields("Cta") And _
                          Numero = AdoSubCtas.Recordset.fields("Numero") And _
                          TipoCta = AdoSubCtas.Recordset.fields("TP") Then
                          'MsgBox "..."
                          For I = 0 To C
                              Printer.Line (Ancho(I), PosLinea - 0.1)-(Ancho(I), PosLinea + 0.3), Negro
                          Next I
                          PosLinea = PosLinea + 0.32
                          Printer.FontItalic = True
                          PrinterFields Ancho(1) + 0.4, PosLinea, AdoSubCtas.Recordset.fields("Cliente"), False
                          Printer.FontItalic = False
                          If AdoSubCtas.Recordset.fields("Prima") <> 0 Then
                             PrinterTexto Ancho(1) + 5.9, PosLinea, " Pr. No. " & Format$(AdoSubCtas.Recordset.fields("Prima"), "000000")
                          End If
                          If AdoSubCtas.Recordset.fields("Factura") <> 0 Then
                             PrinterTexto Ancho(1) + 7.7, PosLinea, " F. No. " & Format$(AdoSubCtas.Recordset.fields("Factura"), "000000")
                          End If
                          PrinterVariables Ancho(2), PosLinea, Abs(AdoSubCtas.Recordset.fields("Debitos") - AdoSubCtas.Recordset.fields("Creditos"))
                          Evaluar = True
                       Else
                          Si_No = False
                          AdoSubCtas.Recordset.MoveLast
                       End If
                       AdoSubCtas.Recordset.MoveNext
                    Loop
                 End If
              End If
           End If
        End If
        For I = 0 To C
            Printer.Line (Ancho(I), PosLinea - 0.1)-(Ancho(I), PosLinea + 0.38), Negro
        Next I
        Debe = Debe + .fields("Debe")
        Haber = Haber + .fields("Haber")
        If Evaluar Then PosLinea = PosLinea + 0.32 Else PosLinea = PosLinea + 0.37
        If PosLinea >= LimiteAlto Then
           PosLinea = PosLinea + 0.05
           Printer.Line (Ancho(0), PosLinea)-(Ancho(C), PosLinea), Negro
           Printer.NewPage
           Encabezado Ancho(0), Ancho(5)
           Printer.FontSize = PorteLetra
           Printer.FontName = TipoLetra
           Encab_Diario_General DataT
           'MsgBox Printer.FontSize & vbCrLf & Printer.FontName
        End If
       .MoveNext
     Loop
End With
PosLinea = PosLinea + 0.05
Printer.Line (InicioX, PosLinea)-(Ancho(C), PosLinea), Negro
Printer.FontSize = PorteLetra
Printer.FontBold = True
Printer.FontItalic = False
PosLinea = PosLinea + 0.05
PrinterVariables Ancho(2), PosLinea, "T O T A L E S"
Printer.FontBold = False
PrinterVariables Ancho(3), PosLinea, Debe
PrinterVariables Ancho(4), PosLinea, Haber
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

Public Sub ImprimirDiarioGeneralSimple(Datas As Adodc)
On Error GoTo Errorhandler
Mensajes = "Seguro de Imprimir en:" & vbCrLf & Printer.DeviceName & "?"
Titulo = "IMPRESION DIARIO GENERAL"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
RatonReloj
InicioX = 0.5: InicioY = 0.01
'DataAnchoCampos InicioX, Datas, 9, TipoTimes, 1
DataAnchoCampos InicioX, Datas, 9, TipoArial, 1
Ancho(0) = 0.5   'Fecha
Ancho(1) = 1.7   'Numero
Ancho(2) = 2.3   'Concepto
Ancho(3) = 7.9   'Codigo
Ancho(4) = 10.2  'Cuenta
Ancho(5) = 15.3  'Debe
Ancho(6) = 17.5  'Haber
Ancho(7) = 20
CantCampos = 7
Pagina = 1
LimiteAlto = LimiteAlto - 1
'Iniciamos la impresion
Printer.FontBold = False
With Datas.Recordset
     .MoveFirst
      EncabDiarioGeneralSimple
      Printer.FontSize = 8
      Printer.FontBold = False
      Debe = 0: Haber = 0
      Fecha = .fields("Fecha")
      Numero = .fields("Numero")
      TP = .fields("TP")
      PrinterFields Ancho(0), PosLinea, .fields("Fecha")
      PrinterVariables Ancho(0), PosLinea + 0.35, TP & "-" & Format$(Numero, "000000")
      NumeroLineas = PrinterLineasMayor(Ancho(2), PosLinea, .fields("Concepto"), 5.5)
      PFil = 0
      If NumeroLineas > 1 Then PFil = PosLinea + (NumeroLineas * 0.4)
      Do While Not .EOF
         If Fecha <> .fields("Fecha") Or Numero <> .fields("Numero") Or TP <> .fields("TP") Then
            PosLinea = PosLinea + 0.2
            Fecha = .fields("Fecha")
            Numero = .fields("Numero")
            TP = .fields("TP")
            If PosLinea < PFil Then PosLinea = PFil
            PrinterFields Ancho(0), PosLinea, .fields("Fecha")
            PrinterVariables Ancho(0), PosLinea + 0.35, TP & "-" & Format$(Numero, "000000")
            NumeroLineas = PrinterLineasMayor(Ancho(2), PosLinea, .fields("Concepto"), 5.5)
            PFil = 0
            If NumeroLineas > 1 Then PFil = PosLinea + (NumeroLineas * 0.4)
         End If
         Printer.FontBold = False
         PrinterFields Ancho(3), PosLinea, .fields("Cta")
         PrinterFields Ancho(4), PosLinea, .fields("Cuenta")
         PrinterFields Ancho(5), PosLinea, .fields("Debe")
         PrinterFields Ancho(6), PosLinea, .fields("Haber")
         PosLinea = PosLinea + 0.35
         If PosLinea > LimiteAlto Then
            If PosLinea < PFil Then PosLinea = PFil
            Printer.Line (Ancho(0), PosLinea)-(Ancho(CantCampos), PosLinea), Negro
            PFil = 0: PosLinea = PosLinea + 0.1
            PrinterVariables Ancho(4), PosLinea, "TOTAL PARCIAL"
            PrinterVariables Ancho(5), PosLinea, Debe
            PrinterVariables Ancho(6), PosLinea, Haber
            Printer.NewPage
            EncabDiarioGeneralSimple
            Printer.FontSize = 8
            Printer.FontBold = False
            PrinterVariables Ancho(4), PosLinea, "SUBTOTAL ANTERIOR"
            PrinterVariables Ancho(5), PosLinea, Debe
            PrinterVariables Ancho(6), PosLinea, Haber
            PosLinea = PosLinea + 0.4
            Printer.Line (Ancho(0), PosLinea)-(Ancho(CantCampos), PosLinea), Negro
            PosLinea = PosLinea + 0.1
         End If
         Debe = Debe + .fields("Debe")
         Haber = Haber + .fields("Haber")
        .MoveNext
      Loop
End With
Printer.Line (InicioX, PosLinea)-(Ancho(CantCampos), PosLinea), Negro
PosLinea = PosLinea + 0.05
Printer.FontBold = True
PrinterVariables Ancho(4), PosLinea, "T O T A L"
PrinterVariables Ancho(5), PosLinea, Debe
PrinterVariables Ancho(6), PosLinea, Haber
PosLinea = PosLinea + 0.4
Printer.Line (InicioX, PosLinea)-(Ancho(CantCampos), PosLinea), Negro
RatonNormal
MensajeEncabData = ""
Printer.FontBold = False
Printer.EndDoc
End If
Exit Sub
Errorhandler:
    RatonNormal
    ErrorDeImpresion
    Exit Sub
End Sub

Public Sub Imprimir_Diario_General_Coop(DataT As Adodc)
Dim ConceptoDe As String
Dim AutorizadoPor  As String
On Error GoTo Errorhandler
Mensajes = "Seguro de Imprimir en:" & vbCrLf & Printer.DeviceName & "?"
Titulo = "IMPRESION DIARIO GENERAL"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then

RatonReloj
ReDim Ancho(6) As Single
InicioX = 0.5: InicioY = 0
Escala_Centimetro 1, TipoTimes, 9
Pagina = 1
HoraSistema = Time
C = 5
Ancho(0) = 0.5
Ancho(1) = 2.9
Ancho(2) = 12.5
Ancho(3) = 14
Ancho(4) = 17
Ancho(5) = 20
'============================================================
' Comenzamos a escribir en la impresora los encabezados
'============================================================
Printer.FontBold = False
With DataT.Recordset
    .MoveFirst
     Encabezado Ancho(0), Ancho(5)
     Encab_Diario_General DataT
     TipoCta = .fields("TP")
     Numero = .fields("Numero")
     Mifecha = .fields("Fecha")
     ConceptoDe = .fields("Concepto")
     AutorizadoPor = .fields("Autorizado")
     Debe = 0: Haber = 0
     Do While Not .EOF
        Printer.FontSize = 9
        If PosLinea + 0.5 > LimiteAlto Then
           Printer.Line (Ancho(0), PosLinea)-(Ancho(C), PosLinea), Negro
           PosLinea = PosLinea + 0.1
           PrinterVariables Ancho(2), PosLinea, "T O T A L E S"
           PrinterVariables Ancho(3), PosLinea, Debe
           PrinterVariables Ancho(4), PosLinea, Haber
           Printer.NewPage
           Encabezado Ancho(0), Ancho(5)
           PrinterVariables Ancho(2), PosLinea, "T O T A L E S"
           PrinterVariables Ancho(3), PosLinea, Debe
           PrinterVariables Ancho(4), PosLinea, Haber
           PosLinea = PosLinea + 0.5
           Printer.Line (Ancho(0), PosLinea)-(Ancho(C), PosLinea), Negro
           PosLinea = PosLinea + 0.1
           Encab_Diario_General DataT
        End If
        If TipoCta <> .fields("TP") Or Numero <> .fields("Numero") Or Mifecha <> .fields("Fecha") Then
           PrinterVariables Ancho(0), PosLinea, "CONCEPTO: " & ConceptoDe
           PosLinea = PosLinea + 0.4
           If Mifecha <> .fields("Fecha") Then
              PrinterVariables Ancho(0), PosLinea, "F E C H A:"
              PrinterFields Ancho(1), PosLinea, .fields("Fecha"), False
              PosLinea = PosLinea + 0.4
              Printer.Line (Ancho(0), PosLinea)-(Ancho(C), PosLinea), Negro
              PosLinea = PosLinea + 0.1
           End If
           If PosLinea + 0.5 > LimiteAlto Then
              Printer.Line (Ancho(0), PosLinea)-(Ancho(C), PosLinea), Negro
              PosLinea = PosLinea + 0.1
              PrinterVariables Ancho(2), PosLinea, "T O T A L E S"
              PrinterVariables Ancho(3), PosLinea, Debe
              PrinterVariables Ancho(4), PosLinea, Haber
              Printer.NewPage
              Encabezado Ancho(0), Ancho(5)
              PrinterVariables Ancho(2), PosLinea, "T O T A L E S"
              PrinterVariables Ancho(3), PosLinea, Debe
              PrinterVariables Ancho(4), PosLinea, Haber
              PosLinea = PosLinea + 0.5
              Printer.Line (Ancho(0), PosLinea)-(Ancho(C), PosLinea), Negro
              PosLinea = PosLinea + 0.1
              Encab_Diario_General DataT
           End If
           TipoCta = .fields("TP")
           Numero = .fields("Numero")
           Mifecha = .fields("Fecha")
           ConceptoDe = .fields("Concepto")
        End If
        PrinterFields Ancho(0), PosLinea, .fields("Cta")
        PrinterFields Ancho(1), PosLinea, .fields("Cuenta")
        PrinterFields Ancho(3), PosLinea, .fields("Debe")
        PrinterFields Ancho(4), PosLinea, .fields("Haber")
        Debe = Debe + .fields("Debe")
        Haber = Haber + .fields("Haber")
        If AutorizadoPor = Ninguno Then AutorizadoPor = .fields("Autorizado")
        PosLinea = PosLinea + 0.4
       .MoveNext
     Loop
End With
PrinterVariables Ancho(0), PosLinea, "CONCEPTO: " & ConceptoDe
PosLinea = PosLinea + 0.4
Printer.Line (InicioX, PosLinea)-(Ancho(C), PosLinea), Negro
PosLinea = PosLinea + 0.1
PrinterVariables Ancho(2), PosLinea, "T O T A L E S"
PrinterVariables Ancho(3), PosLinea, Debe
PrinterVariables Ancho(4), PosLinea, Haber
PosLinea = PosLinea + 1.5
Printer.Line (InicioX, PosLinea)-(InicioX + 3, PosLinea), Negro
PosLinea = PosLinea + 0.1
PrinterVariables InicioX, PosLinea, AutorizadoPor
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

Public Sub Imprimir_General(Datas As Adodc, Opcion As Byte)
Dim TextoTotal As String
Dim IXX As Integer
Dim IParcial As Single
Dim Tesorero As String
Dim CIT As String
On Error GoTo Errorhandler
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
RatonReloj
Escala_Centimetro 1, TipoArialNarrow, 8
CantCampos = 6: C = 6
ReDim Ancho(CantCampos) As Single
IXX = 0
If Opcion = 1 Then
   Ancho(0) = 0.5
   Ancho(1) = 2.5
Else
   Ancho(0) = 0.5
   Ancho(1) = 1
End If
Ancho(2) = 10.5
Ancho(3) = 13.5
Ancho(4) = 16.5
Ancho(5) = 19.5
IXX = 1
Pagina = 1
'Iniciamos la impresion
EncabezadoDataGeneral Datas, Opcion, True
Printer.FontBold = False
With Datas.Recordset
    .MoveFirst
     Printer.FontName = TipoArialNarrow
     Do While Not .EOF
        Printer.FontSize = 9
        NivelCta = Niveles(.fields("Codigo"))
        Printer.FontBold = False
        Printer.FontItalic = False
        Printer.FontUnderline = False
        If .fields("DG") = "G" Then Printer.FontBold = True
        PrinterFields Ancho(0), PosLinea, .fields("Codigo")
        If .fields("TC") <> "N" Then Printer.FontItalic = True
        Select Case .fields("TC")
          Case "C", "P": Printer.FontUnderline = True
        End Select
        If Opcion = 1 Then PrinterFields Ancho(0), PosLinea, .fields("Codigo")
        PCol = Ancho(1) + (NivelCta * 0.3)
        PrinterFields PCol, PosLinea, .fields("Cuenta")
        Printer.FontItalic = False
        Printer.FontUnderline = False
        If MidStrg(.fields("Cuenta"), 1, 3) = " - " Then
           Printer.Line (17, PosLinea)-(19.5, PosLinea), Negro
           PosLinea = PosLinea + 0.05
           Printer.Line (17, PosLinea)-(19.5, PosLinea), Negro
           PosLinea = PosLinea + 0.05
        End If
        PrinterVariables 12.5, PosLinea, .fields("Total_N6")
        PrinterVariables 13.5, PosLinea, .fields("Total_N5")
        PrinterVariables 14.5, PosLinea, .fields("Total_N4")
        PrinterVariables 15.5, PosLinea, .fields("Total_N3")
        PrinterVariables 16.5, PosLinea, .fields("Total_N2")
        PrinterVariables 17.5, PosLinea, .fields("Total_N1")
        If MidStrg(.fields("Cuenta"), 1, 3) = " - " Then
           PosLinea = PosLinea + 0.45
        Else
           PosLinea = PosLinea + 0.35
        End If
        If PosLinea >= LimiteAlto Then
           Printer.NewPage
           PosLinea = 0
           EncabezadoDataGeneral Datas, Opcion, True
           Printer.FontName = TipoArialNarrow
        End If
       .MoveNext
     Loop
End With
PosLinea = PosLinea + 1.2
If PosLinea + 1.2 >= LimiteAlto Then
   Printer.NewPage
   PosLinea = 0
   EncabezadoDataGeneral Datas, Opcion, True
   PosLinea = PosLinea + 2
   Printer.FontName = TipoArialNarrow
End If

Tesorero = Leer_Campo_Empresa("Tesorero")
CIT = Leer_Campo_Empresa("CIT")
Printer.FontBold = True: Printer.FontSize = 10
If Len(Tesorero) > 1 Then
    PrinterTexto 1.5, PosLinea, String$(Len(NombreGerente) + 1, "_")
    PrinterTexto 8.5, PosLinea, String$(Len(NombreContador) + 1, "_")
    PrinterTexto 14.5, PosLinea, String$(Len(NombreContador) + 1, "_")
    PosLinea = PosLinea + 0.4
    PrinterTexto 1.5, PosLinea, " " & NombreGerente
    PrinterTexto 8.5, PosLinea, " " & Tesorero
    PrinterTexto 14.5, PosLinea, " " & NombreContador
    PosLinea = PosLinea + 0.4
    PrinterTexto 1.5, PosLinea, " C.I. " & CI_Representante
    PrinterTexto 8.5, PosLinea, " C.I. " & CIT
    PrinterTexto 14.5, PosLinea, " R.U.C. " & RUC_Contador
    PosLinea = PosLinea + 0.4
    PrinterTexto 1.5, PosLinea, " REPRESENTANTE LEGAL"
    PrinterTexto 8.5, PosLinea, " TESORERO"
    PrinterTexto 14.5, PosLinea, " CONTADOR"
Else
    PrinterTexto 3, PosLinea, String$(Len(NombreGerente) + 1, "_")
    PrinterTexto 11, PosLinea, String$(Len(NombreContador) + 1, "_")
    PosLinea = PosLinea + 0.4
    PrinterTexto 3, PosLinea, " " & NombreGerente
    PrinterTexto 11, PosLinea, " " & NombreContador
    PosLinea = PosLinea + 0.4
    PrinterTexto 3, PosLinea, " C.I. " & CI_Representante
    PrinterTexto 11, PosLinea, " R.U.C. " & RUC_Contador
    PosLinea = PosLinea + 0.4
    PrinterTexto 3, PosLinea, " REPRESENTANTE LEGAL"
    PrinterTexto 11, PosLinea, " CONTADOR"
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

Public Sub ImprimirEstResultAnalitico(Datas As Adodc, _
                                      Fecha_Final As String)
Dim NumeroDeMeses As Integer
Dim Id_Campos As Integer
On Error GoTo Errorhandler
Escala_Centimetro 2, TipoTimes, 7.5, True
CantCampos = 9: C = 9
ReDim Ancho(CantCampos) As Single
Ancho(0) = 0.2  ' Cta
Ancho(1) = 0.4  ' SubCta
Ancho(2) = 5.4  ' Ene   Jun   Nov
Ancho(3) = 9    ' Feb   Jul   Dic
Ancho(4) = 12.6 ' Mar   Ago   Presup
Ancho(5) = 16.2 ' Abr   Sep   Dif
Ancho(6) = 19.8 ' May   Oct
Ancho(7) = 23.4 ' Total
Ancho(8) = 27   ' Fin de la impresion
Pagina = 1
RatonReloj
NumeroDeMeses = Month(Fecha_Final)
'Iniciamos la impresion
With Datas.Recordset
 If .RecordCount > 0 Then
     Id_Campos = .fields.Count - 1
     IE = 2
     JE = IE + 5
     If JE > Id_Campos Then JE = Id_Campos
Volver_Imprimir:
    .MoveFirst
     EncabezadoEstResult12Meses Datas, IE, JE
     Printer.FontBold = False
     Printer.FontSize = 7.5
     Do While Not .EOF
        Printer.FontBold = False
        Select Case MidStrg(.fields(0), 1, 1)
          Case "-", "+", "(", " "
               Printer.FontBold = True
        End Select
        PrinterFields Ancho(0), PosLinea, .fields(1), False
        J = 2
        For I = IE To JE
            If MidStrg(TrimStrg(.fields(1)), 1, 1) = "*" Then
               Select Case MidStrg(.fields(0), 1, 1)
                 Case "-", "+", "("
                      PrinterFields Ancho(J) + 1.8, PosLinea, .fields(I), False
                 Case Else
                      PrinterFields Ancho(J), PosLinea, .fields(I), False
               End Select
            Else
               PrinterFields Ancho(J) + 1.8, PosLinea, .fields(I), False
            End If
            J = J + 1
        Next I
        Select Case MidStrg(.fields(0), 1, 1)
          Case "-", "+", "("
               Printer.Line (Ancho(0), PosLinea - 0.05)-(Ancho(8), PosLinea - 0.05), Negro
               Printer.Line (Ancho(0), PosLinea + 0.36)-(Ancho(8), PosLinea + 0.36), Negro
        End Select
        Select Case MidStrg(TrimStrg(.fields(1)), 1, 1)
          Case "-"
               Printer.Line (Ancho(0), PosLinea + 0.36)-(Ancho(8), PosLinea + 0.36), Negro
        End Select
        Printer.Line (Ancho(0), PosLinea - 0.05)-(Ancho(0), PosLinea + 0.4), Negro
        For J = 2 To 8
            Printer.Line (Ancho(J), PosLinea - 0.05)-(Ancho(J), PosLinea + 0.4), Negro
            Printer.Line (Ancho(J) + 1.8, PosLinea - 0.05)-(Ancho(J) + 1.8, PosLinea + 0.4), Negro
        Next J
        PosLinea = PosLinea + 0.4
        If PosLinea > LimiteAlto Then
           Printer.Line (Ancho(0), PosLinea)-(Ancho(8), PosLinea), Negro
           Printer.NewPage
           PosLinea = 0
           EncabezadoEstResult12Meses Datas, IE, JE
           Printer.FontSize = 7.5
        End If
       .MoveNext
     Loop
     PosLinea = PosLinea + 0.02
     Printer.Line (Ancho(0), PosLinea)-(Ancho(8), PosLinea), Negro
     IE = JE + 1
     If IE > Id_Campos Then IE = Id_Campos
     JE = IE + 5
     If JE > Id_Campos Then JE = Id_Campos
     If IE < Id_Campos Then
        Printer.NewPage
        PosLinea = 0
        GoTo Volver_Imprimir
     End If
  End If
End With
Printer.EndDoc
MensajeEncabData = ""
RatonNormal
Exit Sub
Errorhandler:
    RatonNormal
    ErrorDeImpresion
    Exit Sub
End Sub

Public Sub ImprimirEstResultPresupuesto(Datas As Adodc, _
                                        Fecha_Final As String)
Dim NumeroDeMeses As Integer
On Error GoTo Errorhandler
Mensajes = "Seguro de Imprimir en:" & vbCrLf & Printer.DeviceName & "?"
Titulo = "IMPRESION"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
RatonReloj
InicioX = 0.5: InicioY = 0
Pagina = 1
NumeroDeMeses = Month(Fecha_Final)
'Iniciamos la impresion
With Datas.Recordset
 If .RecordCount > 0 Then
     CantCampos = .fields.Count
     ReDim Ancho(CantCampos + 1) As Single
     Ancho(0) = 0.1   ' Cta
     Ancho(1) = 0.1   ' SubCta
     Ancho(2) = 0.1   ' SubCta
     Distancia = 4
     For I = 3 To CantCampos  ' Ene,Feb,..., etc.
         Ancho(I) = Distancia
         Distancia = Distancia + 1.6
     Next I
     'Ancho(I + 1) = Distancia 'fin de la impresion
     If Distancia <= 19.5 Then
        Escala_Centimetro 1, TipoTimes, 8, True
     Else
        Escala_Centimetro 2, TipoTimes, 8, True
     End If
     EncabezadoEstResult12MesesP Datas
     Printer.FontBold = False
    .MoveFirst
     Printer.FontSize = 6
     
     Do While Not .EOF
        Printer.FontSize = 6
        Printer.FontBold = False
        Printer.FontItalic = False
        Cadena = Space(Len(.fields("Codigo")) - 1) & .fields("Cuenta")
        If .fields("DG") = "G" Then
            Printer.FontBold = True
            Printer.Line (Ancho(0), PosLinea)-(Ancho(CantCampos), PosLinea), Negro
            PosLinea = PosLinea + 0.05
        End If
        If MidStrg(.fields("Cuenta"), 1, 3) = " * " Then
           Printer.FontSize = 5.5
           'Printer.FontBold = True
           Printer.FontItalic = True
           'Printer.Line (Ancho(0), PosLinea)-(Ancho(CantCampos), PosLinea), NEGRO
           'PosLinea = PosLinea + 0.05
           Cadena = Space(Len(.fields("Codigo"))) & .fields("Cuenta")
        End If
        PrinterTexto Ancho(0), PosLinea, Cadena
        Printer.Line (Ancho(3), PosLinea)-(Ancho(CantCampos), PosLinea + 0.35), Blanco, BF
        For I = 3 To CantCampos - 1
            PrinterFields Ancho(I), PosLinea, .fields(I), False
        Next I
        Printer.Line (Ancho(0), PosLinea - 0.1)-(Ancho(0), PosLinea + 0.35), Negro
        For J = 3 To CantCampos
            Printer.Line (Ancho(J), PosLinea - 0.1)-(Ancho(J), PosLinea + 0.35), Negro
        Next J
        PosLinea = PosLinea + 0.31
        If PosLinea > LimiteAlto Then
           Printer.Line (Ancho(0), PosLinea)-(Ancho(CantCampos), PosLinea), Negro
           Printer.NewPage
           PosLinea = 0
           EncabezadoEstResult12MesesP Datas
           Printer.FontSize = 6
        End If
       .MoveNext
     Loop
 End If
End With
Printer.Line (Ancho(0), PosLinea)-(Ancho(CantCampos), PosLinea), Negro
PosLinea = PosLinea + 0.05
Printer.Line (Ancho(0), PosLinea)-(Ancho(CantCampos), PosLinea), Negro
Printer.EndDoc
RatonNormal
MensajeEncabData = ""
Exit Sub
Errorhandler:
    RatonNormal
    ErrorDeImpresion
    Exit Sub
Else
    RatonNormal
End If
End Sub

Public Sub Imprimir_General_Con(Datas As Adodc, _
                              Opcion As Integer, _
                              BG As Boolean)
Dim TextoTotal As String
Dim IXX As Integer
Dim IParcial As Single
On Error GoTo Errorhandler
SizeLetra = 8
Mensajes = "Seguro de Imprimir en:" & vbCrLf & Printer.DeviceName & "?"
Titulo = "IMPRESION"
Bandera = False
Orientacion_Pagina = 2
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
RatonReloj
Escala_Centimetro 1, TipoTimes, 9
IXX = 0
   CantCampos = 6: C = 6
   ReDim Ancho(CantCampos) As Single
   Ancho(0) = 0.5
   Ancho(1) = 3
   Ancho(2) = 10.5
   Ancho(3) = 13.5
   Ancho(4) = 16.5
   Ancho(5) = 19.5
   IXX = 1
Pagina = 1
'Iniciamos la impresion
EncabezadoDataGeneral Datas, 1, BG
Printer.FontBold = False
With Datas.Recordset
    .MoveFirst
     Printer.FontBold = False
     Printer.FontItalic = False
     Do While Not .EOF
        Printer.FontSize = 10
        IParcial = 0
        If .fields("DG") = "G" Then Printer.FontBold = True
        PrinterFields Ancho(0), PosLinea, .fields("Codigo"), False
        If .fields("TC") <> "N" Then Printer.FontItalic = True
        PrinterFields Ancho(1), PosLinea, .fields("Cuenta"), False
        Printer.FontBold = False
        Printer.FontItalic = False
        Total = .fields("Total_N6") + .fields("Total_N5") + .fields("Total_N4") + .fields("Total_N3")
        PrinterVariables Ancho(2), PosLinea, Total
        PrinterFields Ancho(3), PosLinea, .fields("Total_N2"), False
        PrinterFields Ancho(4), PosLinea, .fields("Total_N1"), False
        
''        Select Case Opcion
''          Case 1
''               PrinterFields Ancho(2), PosLinea, .Fields("Saldo_Total_ME"), False
''               PrinterFields Ancho(3), PosLinea, .Fields("Saldo_Total"), False
''               PrinterFields Ancho(4), PosLinea, .Fields("Total"), False
''          Case 2
               'PrinterFields Ancho(2), PosLinea, .Fields("Saldo_Total_ME"), False
''              PrinterFields Ancho(3), PosLinea, .Fields("Saldo_Total"), False
''        End Select
        PosLinea = PosLinea + 0.4
        If PosLinea >= LimiteAlto Then
           Printer.NewPage
           PosLinea = 0
           EncabezadoDataGeneral Datas, 1, BG
        End If
       .MoveNext
     Loop
End With
PosLinea = PosLinea + 1.5
If PosLinea + 1 >= LimiteAlto Then
   Printer.NewPage
   PosLinea = 0
   EncabezadoDataGeneral Datas, 1, BG
   PosLinea = PosLinea + 2
End If
Printer.FontBold = True: Printer.FontSize = 10
PrinterTexto 3, PosLinea, "_________________________"
PrinterTexto 8, PosLinea, "_____________________"
PrinterTexto 13, PosLinea, "______________________"
PosLinea = PosLinea + 0.4
PrinterTexto 3, PosLinea, " REPRESENTANTE LEGAL"
PrinterTexto 8, PosLinea, " AUDITOR INTERNO"
PrinterTexto 13, PosLinea, " C O N T A D O R"
'MsgBox "hOLA"
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

Public Sub Leer_Cta_Superior(Cuenta As String, Optional CodigoC As String)
Dim AdoReg As ADODB.Recordset
Dim CtaSup As String
  CtaSup = Ninguno
  NomCtaSup = Ninguno
  
  sSQL = "SELECT * " _
       & "FROM Catalogo_Cuentas " _
       & "WHERE Codigo = '" & Cuenta & "' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' "
  Select_AdoDB AdoReg, sSQL
  If AdoReg.RecordCount > 0 Then
     NomCta = AdoReg.fields("Cuenta")
     TipoDoc = AdoReg.fields("TC")
     TipoBenef = AdoReg.fields("TC")
     CtaSup = CambioCodigoCtaSup(Cuenta)
  End If
  AdoReg.Close
  
  sSQL = "SELECT * " _
       & "FROM Catalogo_Cuentas " _
       & "WHERE Codigo = '" & CtaSup & "' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' "
  Select_AdoDB AdoReg, sSQL
  If AdoReg.RecordCount > 0 Then NomCtaSup = AdoReg.fields("Cuenta")
  AdoReg.Close
       
  If CodigoC <> "" Then
     NombreCliente = Ninguno
     sSQL = "SELECT * " _
          & "FROM Clientes " _
          & "WHERE Codigo = '" & CodigoC & "' "
     Select_AdoDB AdoReg, sSQL
     If AdoReg.RecordCount > 0 Then NombreCliente = AdoReg.fields("Cliente")
     AdoReg.Close
  End If
End Sub

Public Sub Imprimir_Mayor(DataT As Adodc, Optional ConSubMod As Boolean)
Dim RegSubCtas As ADODB.Recordset
Dim sSQLSC As String
Dim CadCta As String
Dim CadCtaSup As String
Dim NuevoDoc As Boolean
Dim MesActual As Integer
 
Dim NoItem As Long
Dim PosLineaIni As Single
Dim CtaSup As String
Dim Concepto_Aux As String
PorteLetra = 9
TipoLetra = TipoArialNarrow

On Error GoTo Errorhandler
Mensajes = "Seguro de Imprimir en:" & vbCrLf & Printer.DeviceName & "?"
Titulo = "IMPRESION"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
RatonReloj
ReDim Ancho(9) As Single
Escala_Centimetro 1, TipoLetra, PorteLetra
InicioX = 1: InicioY = 0
Pagina = 1: Documento = 1
C = 8
Ancho(0) = 1      ' Fecha
Ancho(1) = 2.5    ' TP
Ancho(2) = 3.1    ' Numero
Ancho(3) = 4.6    ' Concepto
Ancho(4) = 10.8   ' Parcial ME
Ancho(5) = 13.1   ' Debe
Ancho(6) = 15.4   ' Haber
Ancho(7) = 17.7   ' Saldo
Ancho(8) = 20     ' Fin
LimiteAlto = LimiteAlto - 1
'MsgBox LimiteAlto

Cta = Ninguno: NomCta = "": CtaSup = "": NomCtaSup = ""
If DataT.Recordset.RecordCount > 0 Then
   DataT.Recordset.MoveFirst
   Cta = DataT.Recordset.fields("Cta")
End If
Leer_Cta_Superior Cta
CadCtaSup = CtaSup & Space(5) & NomCtaSup
CadCta = Cta & Space(5) & NomCta
'============================================================
If DataT.Recordset.RecordCount > 0 Then
With DataT.Recordset
.MoveFirst
'MsgBox "Encabezado"
Encabezado Ancho(0), Ancho(8)
'PosLinea = PosLinea - 0.5
EncabMayor "L I B R O    M A Y O R", FechaCorte, Ancho(8)
EncabMayor1 CadCta, CadCtaSup, DataT
'.Fields ("Debe"), .Fields("Haber"), .Fields("Saldo")
Printer.FontBold = False
Printer.FontItalic = False
Printer.FontName = TipoLetra
Printer.FontSize = PorteLetra
MesActual = FechaMes(.fields("Fecha"))
Suma_ME = 0: SumaDebe = 0: SumaHaber = 0: SumaSaldo = 0
Mifecha = .fields("Fecha")
TipoDoc = .fields("TP")
Numero = .fields("Numero")
NoItem = .fields("Item")
PrinterFields Ancho(0), PosLinea, .fields("Fecha"), False
PrinterFields Ancho(1), PosLinea, .fields("TP"), False
PrinterFields Ancho(2), PosLinea, .fields("Numero"), False
PosLineaIni = PosLinea
Do While Not .EOF
   PorteLetra = 9
   Printer.FontBold = False
   Printer.FontItalic = False
   Printer.FontName = TipoLetra
   Printer.FontSize = PorteLetra
   'MsgBox "Hola"
   If Cta <> .fields("Cta") Then
      For J = 0 To C
          Printer.Line (Ancho(J), PosLineaIni - 0.1)-(Ancho(J), PosLinea + 0.1), Negro
      Next J
      Printer.FontBold = True
      PosLinea = PosLinea + 0.05
      Printer.Line (InicioX, PosLinea)-(Ancho(C), PosLinea), Negro
      PosLinea = PosLinea - 0.05
      Printer.Line (InicioX, PosLinea)-(Ancho(C), PosLinea), Negro
      PosLinea = PosLinea + 0.05
      PrinterVariableTexto Ancho(0), PosLinea, "Fin de: ", MesesLetras(MesActual)
      MesActual = FechaMes(.fields("Fecha"))
      PrinterTexto Ancho(4) - 2, PosLinea, "TOTALES"
      PrinterVariables Ancho(4), PosLinea, Suma_ME
      PrinterVariables Ancho(5), PosLinea, SumaDebe
      PrinterVariables Ancho(6), PosLinea, SumaHaber
      PrinterVariables Ancho(7), PosLinea, SumaSaldo
      Documento = Documento + 1
      NomCta = "": CtaSup = "": NomCtaSup = ""
      Cta = .fields("Cta")
      Leer_Cta_Superior Cta
      
      SumaDebe = 0: SumaHaber = 0: Suma_ME = 0
      PosLinea = PosLinea + 0.5
      CadCtaSup = CtaSup & Space(5) & NomCtaSup
      CadCta = Cta & Space(5) & NomCta
      If PosLinea >= LimiteAlto Then
         Printer.Line (InicioX, PosLinea + 0.1)-(Ancho(C), PosLinea + 0.1), Negro
         Printer.NewPage
         Encabezado Ancho(0), Ancho(8)
         EncabMayor "L I B R O    M A Y O R", FechaCorte, Ancho(8)
      End If
      EncabMayor1 CadCta, CadCtaSup, DataT
      PosLineaIni = PosLinea
      '.Fields("Debe"), .Fields("Haber"), .Fields("Saldo")
      PrinterFields Ancho(0), PosLinea, .fields("Fecha"), False
      Mifecha = .fields("Fecha")
      Printer.FontBold = False
      Printer.FontItalic = False
      Printer.FontName = TipoLetra
      Printer.FontSize = PorteLetra
   End If
   If MesActual <> FechaMes(.fields("Fecha")) Then
      For J = 0 To C
          Printer.Line (Ancho(J), PosLineaIni - 0.1)-(Ancho(J), PosLinea + 0.1), Negro
      Next J
      Printer.FontBold = True
      PosLinea = PosLinea + 0.1
      Printer.Line (InicioX, PosLinea)-(Ancho(C), PosLinea), Negro
      PosLinea = PosLinea - 0.05
      Printer.Line (InicioX, PosLinea)-(Ancho(C), PosLinea), Negro
      PosLinea = PosLinea + 0.1
      PrinterVariableTexto Ancho(0), PosLinea, "Fin de: ", MesesLetras(MesActual)
      MesActual = FechaMes(DataT.Recordset.fields("Fecha"))
      PrinterTexto Ancho(4) - 2, PosLinea, "TOTALES"
      PrinterVariables Ancho(4), PosLinea, Suma_ME
      PrinterVariables Ancho(5), PosLinea, SumaDebe
      PrinterVariables Ancho(6), PosLinea, SumaHaber
      PrinterVariables Ancho(7), PosLinea, SumaSaldo
      SumaDebe = 0: SumaHaber = 0: Suma_ME = 0
      PosLinea = PosLinea + 0.35
      PrinterVariableTexto Ancho(0), PosLinea, "Inicio de: ", MesesLetras(MesActual)
      PosLinea = PosLinea + 0.4
      Printer.Line (InicioX, PosLinea)-(Ancho(C), PosLinea), Negro
      PosLinea = PosLinea + 0.1
      Printer.FontBold = False
      Printer.FontItalic = False
      PosLineaIni = PosLinea
   End If
   If Mifecha <> .fields("Fecha") Then
      PrinterFields Ancho(0), PosLinea, .fields("Fecha"), False
      Mifecha = .fields("Fecha")
   End If
   If TipoDoc <> .fields("TP") Or Numero <> .fields("Numero") Then
      PrinterFields Ancho(1), PosLinea, .fields("TP"), False
      PrinterFields Ancho(2), PosLinea, .fields("Numero"), False
      Numero = .fields("Numero")
      TipoDoc = .fields("TP")
   End If
   If Len(.fields("Cliente")) > 1 Then
      Printer.FontBold = True
      Printer.FontSize = PorteLetra - 2
      PrinterFields Ancho(3), PosLinea, .fields("Cliente"), False
      Printer.FontSize = PorteLetra
      Printer.FontBold = False
      PosLinea = PosLinea + 0.3
   End If
   Concepto_Aux = .fields("Concepto")
   If Len(.fields("Cheq_Dep")) > 1 Then Concepto_Aux = Concepto_Aux & ". Doc. " & .fields("Cheq_Dep")
   If OpcCoop Then
      NumeroLineas = PrinterLineasMayor(Ancho(3), PosLinea, Concepto_Aux, 9.5)
      PosLinea = PosLinea_Aux
      PrinterFields Ancho(4), PosLinea, .fields("Parcial_ME"), False
      PrinterFields Ancho(5), PosLinea, .fields("Debe"), False
      PrinterFields Ancho(6), PosLinea, .fields("Haber"), False
      PrinterFields Ancho(7), PosLinea, .fields("Saldo"), False
      
'''      If .Fields("ME") Then
'''          PrinterFields Ancho(5), PosLinea, .Fields("Debe_ME"), False
'''          PrinterFields Ancho(6), PosLinea, .Fields("Haber_ME"), False
'''          PrinterFields Ancho(7), PosLinea, .Fields("Saldo_ME"), False
'''          SumaSaldo = .Fields("Saldo_ME")
'''          SumaDebe = SumaDebe + .Fields("Debe_ME")
'''          SumaHaber = SumaHaber + .Fields("Haber_ME")
'''      Else
'''          PrinterFields Ancho(5), PosLinea, .Fields("Debe"), False
'''          PrinterFields Ancho(6), PosLinea, .Fields("Haber"), False
'''          PrinterFields Ancho(7), PosLinea, .Fields("Saldo"), False
'''          SumaSaldo = .Fields("Saldo")
'''          SumaDebe = SumaDebe + .Fields("Debe")
'''          SumaHaber = SumaHaber + .Fields("Haber")
'''      End If
   Else
      NumeroLineas = PrinterLineasMayor(Ancho(3), PosLinea, Concepto_Aux, Ancho(4) - Ancho(3) - 0.2)
      PosLinea = PosLinea_Aux
      PrinterFields Ancho(4), PosLinea, .fields("Parcial_ME"), False
      PrinterFields Ancho(5), PosLinea, .fields("Debe"), False
      PrinterFields Ancho(6), PosLinea, .fields("Haber"), False
      PrinterFields Ancho(7), PosLinea, .fields("Saldo"), False
      SumaSaldo = .fields("Saldo")
      Suma_ME = .fields("Saldo_ME")
      SumaDebe = SumaDebe + .fields("Debe")
      SumaHaber = SumaHaber + .fields("Haber")
   End If
   PosLinea = PosLinea + AltoLetra
   If ConSubMod Then
     'Llenar SubCtas
      sSQLSC = "SELECT T.TC,T.Factura,C.Cliente,T.Detalle_SubCta,T.Debitos,T.Creditos,T.Fecha_V,T.Prima,T.Codigo " _
             & "FROM Trans_SubCtas As T,Clientes As C " _
             & "WHERE T.TP = '" & .fields("TP") & "' " _
             & "AND T.Numero = " & .fields("Numero") & " " _
             & "AND T.Item = '" & .fields("Item") & "' " _
             & "AND T.Cta = '" & .fields("Cta") & "' " _
             & "AND T.Periodo = '" & Periodo_Contable & "' " _
             & "AND T.TC IN ('C','P') " _
             & "AND T.Codigo = C.Codigo " _
             & "UNION " _
             & "SELECT T.TC,T.Factura,C.Detalle As Cliente,T.Detalle_SubCta,T.Debitos,T.Creditos,T.Fecha_V,T.Prima,T.Codigo " _
             & "FROM Trans_SubCtas As T,Catalogo_SubCtas As C " _
             & "WHERE T.TP = '" & .fields("TP") & "' " _
             & "AND T.Numero = " & .fields("Numero") & " " _
             & "AND T.Item = '" & .fields("Item") & "' " _
             & "AND T.Cta = '" & .fields("Cta") & "' " _
             & "AND T.Periodo = '" & Periodo_Contable & "' " _
             & "AND T.TC = C.TC " _
             & "AND T.Item = C.Item " _
             & "AND T.Periodo = C.Periodo " _
             & "AND T.Codigo = C.Codigo " _
             & "ORDER BY Cliente,T.Fecha_V,T.Factura "
      Select_AdoDB RegSubCtas, sSQLSC
      
     'MsgBox .Fields("Cta") & " - " & .Fields("TP") & " - " & .Fields("Numero") & vbCrLf & RegSubCtas.RecordCount
      
      If RegSubCtas.RecordCount > 0 Then
         Printer.FontBold = False
         Printer.FontName = TipoArialNarrow  'TipoCourierNew
         Printer.FontSize = PorteLetra - 3
         Do While Not RegSubCtas.EOF
            If PosLinea >= LimiteAlto Then
               For J = 0 To C
                   Printer.Line (Ancho(J), PosLineaIni - 0.1)-(Ancho(J), PosLinea + 0.1), Negro
               Next J
               Printer.Line (InicioX, PosLinea + 0.1)-(Ancho(C), PosLinea + 0.1), Negro
               Printer.NewPage
               CadCtaSup = CtaSup & Space(5) & NomCtaSup
               CadCta = Cta & Space(5) & NomCta
               Encabezado Ancho(0), Ancho(8)
               PosLinea = PosLinea - 0.5
               EncabMayor "L I B R O    M A Y O R", FechaCorte, Ancho(8)
               EncabMayor1 CadCta, CadCtaSup, DataT
             '.Fields("Debe"), .Fields("Haber"), .Fields("Saldo")
               Printer.FontBold = False
               Printer.FontItalic = False
               Printer.FontName = TipoLetra
               Printer.FontSize = PorteLetra
               PosLineaIni = PosLinea
            End If
            Printer.FontItalic = True
            If RegSubCtas.fields("TC") = "CC" Then
               sSQLSC = " * " & RegSubCtas.fields("Codigo") & ": " & RegSubCtas.fields("Cliente")
            Else
               sSQLSC = " * " & RegSubCtas.fields("Cliente")
            End If
            PrinterTexto Ancho(3), PosLinea, sSQLSC
            Printer.FontItalic = False
            PrinterFields Ancho(4) - 0.3, PosLinea, RegSubCtas.fields("Debitos"), False
            PrinterFields Ancho(4) + 0.65, PosLinea, RegSubCtas.fields("Creditos"), False
            PosLinea = PosLinea + AltoLetra
            RegSubCtas.MoveNext
         Loop
         Printer.FontBold = False
         Printer.FontItalic = False
         Printer.FontName = TipoLetra
         Printer.FontSize = PorteLetra
      End If
      RegSubCtas.Close
   End If
   
   If PosLinea >= LimiteAlto Then
      For J = 0 To C
          Printer.Line (Ancho(J), PosLineaIni - 0.1)-(Ancho(J), PosLinea + 0.1), Negro
      Next J
      Printer.Line (InicioX, PosLinea + 0.1)-(Ancho(C), PosLinea + 0.1), Negro
      Printer.NewPage
      CadCtaSup = CtaSup & Space(5) & NomCtaSup
      CadCta = Cta & Space(5) & NomCta
      Encabezado Ancho(0), Ancho(8)
      PosLinea = PosLinea - 0.5
      EncabMayor "L I B R O    M A Y O R", FechaCorte, Ancho(8)
      EncabMayor1 CadCta, CadCtaSup, DataT
    '.Fields("Debe"), .Fields("Haber"), .Fields("Saldo")
      Printer.FontBold = False
      Printer.FontItalic = False
      Printer.FontName = TipoLetra
      Printer.FontSize = PorteLetra
      PosLineaIni = PosLinea
   End If
  .MoveNext
Loop
End With
For J = 0 To C
    Printer.Line (Ancho(J), PosLineaIni - 0.1)-(Ancho(J), PosLinea + 0.1), Negro
Next J
PosLinea = PosLinea + 0.1
Printer.Line (InicioX, PosLinea)-(Ancho(C), PosLinea), Negro
PosLinea = PosLinea - 0.05
Printer.Line (InicioX, PosLinea)-(Ancho(C), PosLinea), Negro
Printer.FontBold = True
PosLinea = PosLinea + 0.1
PrinterVariableTexto Ancho(0), PosLinea, "Fin de: ", MesesLetras(MesActual)
PrinterTexto Ancho(4) - 2, PosLinea, "TOTALES"
PrinterVariables Ancho(4), PosLinea, Suma_ME
PrinterVariables Ancho(5), PosLinea, SumaDebe
PrinterVariables Ancho(6), PosLinea, SumaHaber
PrinterVariables Ancho(7), PosLinea, SumaSaldo
PosLinea = PosLinea + 0.1
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

Public Sub Imprimir_Mayor_Aux(DataT As Adodc, _
                              NombreCtaSuperior As String)
Dim CadCta As String
Dim CadCtaSup As String
Dim NuevoDoc As Boolean
Dim MesActual As Integer
Dim Total_Prima As Double
 
Dim CtaSup As String
Dim Detalle_Concepto As String

PorteLetra = 8
TipoLetra = TipoArialNarrow

On Error GoTo Errorhandler
Mensajes = "Seguro de Imprimir en:" & vbCrLf & Printer.DeviceName & "?"
Titulo = "IMPRESION"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
RatonReloj
ReDim Ancho(11) As Single
Escala_Centimetro 1, TipoLetra, PorteLetra
InicioX = 1: InicioY = 0
Pagina = 1: Documento = 1
C = 9
Ancho(0) = 1     'Fecha
Ancho(1) = 2.3   'Factura/Prima
Ancho(2) = 3.6   'TP
Ancho(3) = 4.1   'Numero
Ancho(4) = 5.4   'Concepto
Ancho(5) = 11.7  'Parcial_ME
Ancho(6) = 13.7  'Debitos
Ancho(7) = 15.7  'Creditos
Ancho(8) = 17.7  'Saldo
Ancho(9) = 19.7  'Fin

NomCta = "": CtaSup = ""
NomCtaSup = ""
DataT.Recordset.MoveFirst
Cta = DataT.Recordset.fields("Cta")
Codigo = DataT.Recordset.fields("Codigo")
Leer_Cta_Superior Cta, Codigo
Select Case TipoDoc
  Case "C": TipoDoc = "Ctas. por Cobrar"
  Case "P": TipoDoc = "Ctas. por Pagar"
  Case "I": TipoDoc = "Ctas. de Ingresos"
  Case "G": TipoDoc = "Ctas. de Gastos"
  Case "PM": TipoDoc = "Valores de Primas"
End Select
CadCta = Cta & " - " & NomCta
CadCtaSup = Codigo & " - " & NombreCliente
'============================================================

With DataT.Recordset
If .RecordCount > 0 Then
   .MoveFirst
Encabezado Ancho(0), Ancho(9)
    
PosLinea = PosLinea - 0.5
EncabMayor "Modulos de Subcuentas de Bloque", FechaCorte, Ancho(9)
EncabMayorMA CadCta, CadCtaSup, DataT, True
MesActual = FechaMes(.fields("Fecha"))
Mifecha = .fields("Fecha")
Total_Prima = 0: Suma_ME = 0: Suma_MN = 0: SumaDebe = 0: SumaHaber = 0: SumaSaldo = 0: SumaParcial_ME = 0
Printer.FontBold = False
Printer.FontItalic = False
Printer.FontName = TipoLetra
Printer.FontSize = PorteLetra
PrinterFields Ancho(0), PosLinea, .fields("Fecha"), True
Do While Not .EOF
   If PosLinea > LimiteAlto Then
      Printer.Line (InicioX, PosLinea + 0.05)-(Ancho(C), PosLinea + 0.05), Negro
      Printer.NewPage
      CadCta = Cta & " - " & NomCta
      CadCtaSup = Codigo & " - " & NombreCliente
      Encabezado Ancho(0), Ancho(9)
      PosLinea = PosLinea - 0.5
      EncabMayor "Modulos de Subcuentas de Bloque", FechaCorte, Ancho(7)
      EncabMayorMA CadCta, CadCtaSup, DataT, True
      Printer.FontBold = False
      Printer.FontItalic = False
      Printer.FontName = TipoLetra
      Printer.FontSize = PorteLetra
   End If
   If Cta <> .fields("Cta") Then
      Printer.Line (InicioX, PosLinea)-(Ancho(C), PosLinea), Negro
      Printer.Line (InicioX, PosLinea + 0.05)-(Ancho(C), PosLinea + 0.05), Negro
      PosLinea = PosLinea + 0.15
      'PrinterVariableTexto Ancho(0), PosLinea, "Fin de: ", MesesLetras(MesActual), False
      MesActual = FechaMes(.fields("Fecha"))
      Printer.FontBold = True
      PrinterTexto Ancho(5), PosLinea, "T O T A L E S"
      If TipoBenef = "PM" Then PrinterVariables Ancho(1), PosLinea, Total_Prima
      PrinterVariables Ancho(6), PosLinea, SumaDebe
      PrinterVariables Ancho(7), PosLinea, SumaHaber
      PrinterVariables Ancho(8), PosLinea, Saldo
      Documento = Documento + 1
      Cta = .fields("Cta")
      Leer_Cta_Superior Cta, Codigo
      Select Case TipoDoc
        Case "C": TipoDoc = "Ctas. por Cobrar"
        Case "P": TipoDoc = "Ctas. por Pagar"
        Case "I": TipoDoc = "Ctas. de Ingresos"
        Case "G": TipoDoc = "Ctas. de Gastos"
        Case "PM": TipoDoc = "Valores de Primas"
      End Select
      CadCta = Cta & " - " & NomCta
      CadCtaSup = Codigo & " - " & NombreCliente
      Total_Prima = 0: SumaDebe = 0: SumaHaber = 0: Suma_ME = 0: Suma_MN = 0: SumaParcial_ME = 0
      PosLinea = PosLinea + 0.6
      If PosLinea + 2 > LimiteAlto Then
         Printer.Line (InicioX, PosLinea + 0.1)-(Ancho(C), PosLinea + 0.1), Negro
         Printer.NewPage
         Encabezado Ancho(0), Ancho(9)
         PosLinea = PosLinea - 0.5
         EncabMayor "Modulos de Subcuentas de Bloque", FechaCorte, Ancho(9)
      End If
      EncabMayorMA CadCta, CadCtaSup, DataT, True '.Fields("Debitos"), .Fields("Creditos"), .Fields("Saldo")
      Mifecha = .fields("Fecha")
      Printer.FontBold = False
      Printer.FontItalic = False
      Printer.FontName = TipoLetra
      Printer.FontSize = PorteLetra
      PrinterFields Ancho(0), PosLinea, .fields("Fecha"), True
   End If
   If Codigo <> .fields("Codigo") Then
      Printer.Line (InicioX, PosLinea)-(Ancho(C), PosLinea), Negro
      Printer.Line (InicioX, PosLinea + 0.05)-(Ancho(C), PosLinea + 0.05), Negro
      PosLinea = PosLinea + 0.15
      'PrinterVariableTexto Ancho(0), PosLinea, "Fin de: ", MesesLetras(MesActual), False
      MesActual = FechaMes(.fields("Fecha"))
      Printer.FontBold = True
      PrinterTexto Ancho(5), PosLinea, "T O T A L E S"
      PrinterVariables Ancho(6), PosLinea, SumaDebe
      PrinterVariables Ancho(7), PosLinea, SumaHaber
      PrinterVariables Ancho(8), PosLinea, Saldo
      Documento = Documento + 1
      Codigo = .fields("Codigo")
      Leer_Cta_Superior Cta, Codigo
      CadCta = Cta & " - " & NomCta
      CadCtaSup = Codigo & " - " & NombreCliente

      SumaDebe = 0: SumaHaber = 0: Suma_ME = 0: Suma_MN = 0: SumaParcial_ME = 0
      PosLinea = PosLinea + 0.6
      If PosLinea + 2 > LimiteAlto Then
         Printer.Line (InicioX, PosLinea + 0.05)-(Ancho(C), PosLinea + 0.05), Negro
         Printer.NewPage
         Encabezado Ancho(0), Ancho(9)
         PosLinea = PosLinea - 0.5
         EncabMayor "Modulos de Subcuentas de Bloque", FechaCorte, Ancho(9)
      End If
      EncabMayorMA CadCta, CadCtaSup, DataT, True '.Fields("Debitos"), .Fields("Creditos"), .Fields("Saldo")
      Mifecha = .fields("Fecha")
      Printer.FontBold = False
      Printer.FontItalic = False
      Printer.FontName = TipoLetra
      Printer.FontSize = PorteLetra
      PrinterFields Ancho(0), PosLinea, .fields("Fecha"), True
   End If
   If Mifecha <> .fields("Fecha") Then
      PrinterFields Ancho(0), PosLinea, .fields("Fecha"), True
      Mifecha = .fields("Fecha")
   End If
   If TipoBenef = "PM" Then
      PrinterFields Ancho(1), PosLinea, .fields("Prima"), True
   Else
      PrinterFields Ancho(1), PosLinea, .fields("Factura"), True
   End If
   PrinterFields Ancho(2), PosLinea, .fields("TP"), True
   PrinterFields Ancho(3), PosLinea, .fields("Numero"), True
   Detalle_Concepto = ""
   If .fields("Concepto") <> Ninguno And .fields("Detalle_SubCta") <> Ninguno Then
       Detalle_Concepto = .fields("Concepto") & ", " & .fields("Detalle_SubCta")
   Else
       If .fields("Concepto") <> Ninguno Then Detalle_Concepto = .fields("Concepto")
       If .fields("Detalle_SubCta") <> Ninguno Then Detalle_Concepto = Detalle_Concepto & " " & .fields("Detalle_SubCta")
   End If
   Detalle_Concepto = TrimStrg(Detalle_Concepto)
   NumeroLineas = PrinterLineasMayor(Ancho(4), PosLinea, Detalle_Concepto, Ancho(5) - Ancho(4) - 0.2)
   If NumeroLineas > 1 Then
      For I = 1 To NumeroLineas
         For J = 0 To C
           Printer.Line (Ancho(J), PosLinea - 0.1)-(Ancho(J), PosLinea + 0.5), Negro
         Next J
         PosLinea = PosLinea + 0.4
      Next
      PosLinea = PosLinea - 0.4
   End If
   PrinterFields Ancho(5), PosLinea, .fields("Parcial_ME"), True
   PrinterFields Ancho(6), PosLinea, .fields("Debitos"), True
   PrinterFields Ancho(7), PosLinea, .fields("Creditos"), True
   PrinterFields Ancho(8), PosLinea, .fields("Saldo_MN"), True
   If TipoBenef = "PM" Then Total_Prima = Total_Prima + .fields("Prima")
   SumaDebe = SumaDebe + .fields("Debitos")
   SumaHaber = SumaHaber + .fields("Creditos")
   Saldo = .fields("Saldo_MN")
   For J = 0 To C
      Printer.Line (Ancho(J), PosLinea - 0.1)-(Ancho(J), PosLinea + 0.4), Negro
   Next J
   PosLinea = PosLinea + AltoLetra
  .MoveNext
Loop
Printer.Line (InicioX, PosLinea)-(Ancho(C), PosLinea), Negro
Printer.Line (InicioX, PosLinea + 0.05)-(Ancho(C), PosLinea + 0.05), Negro
PosLinea = PosLinea + 0.15
'PrinterVariableTexto Ancho(0), PosLinea, "Fin de: ", MesesLetras(MesActual), False
Printer.FontBold = True
PrinterTexto Ancho(5), PosLinea, "T O T A L E S"
If TipoBenef = "PM" Then PrinterVariables Ancho(1), PosLinea, Total_Prima
PrinterVariables Ancho(6), PosLinea, SumaDebe
PrinterVariables Ancho(7), PosLinea, SumaHaber
PrinterVariables Ancho(8), PosLinea, Saldo
End If
End With
RatonNormal
MensajeEncabData = ""
Printer.EndDoc
'MsgBox "Hola"
Exit Sub
Errorhandler:
    RatonNormal
    ErrorDeImpresion
    Exit Sub
Else
    RatonNormal
End If
End Sub

Public Sub ImprimirMayorAux1(DataT As Adodc, _
                             NombreCtaSuperior As String)
Dim DataReg As ADODB.Recordset
Dim CadCta As String
Dim CadCtaSup As String
Dim NuevoDoc As Boolean
Dim MesActual As Integer
On Error GoTo Errorhandler
RatonReloj
ReDim Ancho(10) As Single
Escala_Centimetro 2, TipoCondensed, 8
InicioX = 0.5: InicioY = 0
Set DataReg = New ADODB.Recordset
DataReg.CursorType = adOpenStatic
DataReg.CursorLocation = adUseClient
Pagina = 1: Documento = 1
C = 9
Ancho(0) = 0.5  'Fecha
Ancho(1) = 2.2  'Factura
Ancho(2) = 3.7  'TP
Ancho(3) = 4.4  'Numero
Ancho(4) = 5.9  'Concepto
Ancho(5) = 16.5 'Parcial_ME
Ancho(6) = 19   'Debitos
Ancho(7) = 21.5 'Creditos
Ancho(8) = 24   'Saldo
Ancho(9) = 26.5
NomCta = "": Cta_Sup = "": NomCtaSup = NombreCtaSuperior
If DataT.Recordset.RecordCount > 0 Then
DataT.Recordset.MoveFirst
Cta = DataT.Recordset.fields("Codigo")
sSQL = "SELECT * " _
     & "FROM Catalogo_SubCtas " _
     & "WHERE Codigo = '" & Cta & "' " _
     & "AND Item = '" & NumEmpresa & "' " _
     & "AND Periodo = '" & Periodo_Contable & "' "
DataReg.open sSQL, AdoStrCnn, , , adCmdText
If DataReg.RecordCount > 0 Then
   NomCta = DataReg.fields("Beneficiario")
   Cta_Sup = CambioCodigoCtaSup(Cta)
End If
CadCtaSup = NomCtaSup
CadCta = Cta & Space(10) & NomCta
'============================================================
With DataT.Recordset
.MoveFirst
Encabezado Ancho(0), Ancho(9)
EncabMayor "Modulos de Subcuentas de Bloque", FechaCorte, Ancho(9)
EncabMayorM CadCta, CadCtaSup, DataT, True
Printer.Font.bold = False
MesActual = FechaMes(.fields("Fecha"))
Suma_ME = 0: SumaDebe = 0: SumaHaber = 0: SumaSaldo = 0: SumaParcial_ME = 0
Do While Not .EOF
   If PosLinea > LimiteAlto Then
      Printer.Line (InicioX, PosLinea + 0.1)-(Ancho(C), PosLinea + 0.1), Negro
      Printer.NewPage
      CadCtaSup = NomCtaSup
      CadCta = Cta & Space(10) & NomCta
      Encabezado Ancho(0), Ancho(9)
      EncabMayor "Modulos de Subcuentas de Bloque", FechaCorte, Ancho(9)
      EncabMayorM CadCta, CadCtaSup, DataT, True
   End If
   If Cta <> .fields("Codigo") Then
      Printer.FontBold = True
      Printer.Line (InicioX, PosLinea)-(Ancho(C), PosLinea), Negro
      Printer.Line (InicioX, PosLinea + 0.1)-(Ancho(C), PosLinea + 0.1), Negro
      PosLinea = PosLinea + 0.2
      PrinterVariableTexto Ancho(0), PosLinea, "Fin de: ", MesesLetras(MesActual)
      MesActual = FechaMes(.fields("Fecha"))
      Printer.FontBold = False
      PrinterTexto Ancho(4) + 3.8, PosLinea, "T O T A L E S"
      PrinterVariables Ancho(5), PosLinea, Suma_ME
      PrinterVariables Ancho(6), PosLinea, SumaDebe
      PrinterVariables Ancho(7), PosLinea, SumaHaber
      PrinterVariables Ancho(8), PosLinea, SumaSaldo
      Documento = Documento + 1
      NomCta = "": Cta_Sup = ""
      Cta = .fields("Codigo")
      sSQL = "SELECT * " _
           & "FROM Catalogo_SubCtas " _
           & "WHERE Codigo = '" & Cta & "' " _
           & "AND Item = '" & NumEmpresa & "' " _
           & "AND Periodo = '" & Periodo_Contable & "' "
      DataReg.open sSQL, AdoStrCnn, , , adCmdText
      If DataReg.RecordCount > 0 Then
         NomCta = DataReg.fields("Beneficiario")
         Cta_Sup = CambioCodigoCtaSup(Cta)
      End If
      SumaDebe = 0: SumaHaber = 0: Suma_ME = 0: SumaParcial_ME = 0
      PosLinea = PosLinea + 1
      CadCtaSup = NomCtaSup
      CadCta = Cta & Space(10) & NomCta
      If PosLinea + 3 > LimiteAlto Then
         Printer.Line (InicioX, PosLinea + 0.1)-(Ancho(C), PosLinea + 0.1), Negro
         Printer.NewPage
         Encabezado Ancho(0), Ancho(9)
         EncabMayor "Modulos de Subcuentas de Bloque", FechaCorte, Ancho(9)
      End If
      EncabMayorM CadCta, CadCtaSup, DataT, True '.Fields("Debitos"), .Fields("Creditos"), .Fields("Saldo")
   End If
   If MesActual <> FechaMes(.fields("Fecha")) Then
      Printer.FontBold = True
      Printer.Line (InicioX, PosLinea)-(Ancho(C), PosLinea), Negro
      PosLinea = PosLinea + 0.1
      PrinterVariableTexto Ancho(0), PosLinea, "Fin de: ", MesesLetras(MesActual)
      MesActual = FechaMes(DataT.Recordset.fields("Fecha"))
      Printer.FontBold = False
      PrinterTexto Ancho(3) + 2, PosLinea, "TOTALES"
      PrinterVariables Ancho(4), PosLinea, SumaDebe
      PrinterVariables Ancho(5), PosLinea, SumaHaber
      PrinterVariables Ancho(6), PosLinea, SumaSaldo
      Printer.FontBold = True
      SumaDebe = 0: SumaHaber = 0: Suma_ME = 0
      PosLinea = PosLinea + 0.4
      PrinterVariableTexto Ancho(0), PosLinea, "Inicio de: ", MesesLetras(MesActual)
      PosLinea = PosLinea + 0.4
      Printer.Line (InicioX, PosLinea)-(Ancho(C), PosLinea), Negro
      PosLinea = PosLinea + 0.1
      Printer.FontBold = False
   End If
   Printer.FontSize = 8
   PrinterFields Ancho(0), PosLinea, .fields("Fecha"), True
   PrinterFields Ancho(1), PosLinea, .fields("Factura"), True
   PrinterFields Ancho(2), PosLinea, .fields("TP"), True
   PrinterFields Ancho(3), PosLinea, .fields("Numero"), True
   NumeroLineas = PrinterLineasMayor(Ancho(4), PosLinea, .fields("Concepto"), 10.4)
   If NumeroLineas > 1 Then
      For I = 1 To NumeroLineas
         For J = 0 To C
           Printer.Line (Ancho(J), PosLinea - 0.1)-(Ancho(J), PosLinea + 0.5), Negro
         Next J
         PosLinea = PosLinea + 0.4
      Next
      PosLinea = PosLinea - 0.4
   End If
   PrinterFields Ancho(5), PosLinea, .fields("Parcial_MN"), True
   PrinterFields Ancho(6), PosLinea, .fields("Parcial_ME"), True
   PrinterFields Ancho(7), PosLinea, .fields("Saldo"), True
   PrinterFields Ancho(8), PosLinea, .fields("Saldo_ME"), True
   For J = 0 To C
      Printer.Line (Ancho(J), PosLinea - 0.1)-(Ancho(J), PosLinea + 0.4), Negro
   Next J
   SumaSaldo = .fields("Saldo")
   If .fields("Parcial_ME") = 0 Then Suma_ME = 0 Else Suma_ME = .fields("Saldo_ME")
   SumaParcial_ME = SumaParcial_ME + .fields("Parcial_ME")
   SumaDebe = SumaDebe + .fields("Parcial_MN")
   SumaHaber = SumaHaber + .fields("Parcial_ME")
   PosLinea = PosLinea + AltoLetra
  .MoveNext
Loop
End With
Printer.Line (InicioX, PosLinea)-(Ancho(C), PosLinea), Negro
Printer.Line (InicioX, PosLinea + 0.1)-(Ancho(C), PosLinea + 0.1), Negro
Printer.FontBold = True
PosLinea = PosLinea + 0.2
PrinterVariableTexto Ancho(0), PosLinea, "Fin de: ", MesesLetras(MesActual)
Printer.FontBold = False
PrinterTexto Ancho(4) + 3.8, PosLinea, "T O T A L E S"
PrinterVariables Ancho(5), PosLinea, Suma_ME
PrinterVariables Ancho(6), PosLinea, SumaDebe
PrinterVariables Ancho(7), PosLinea, SumaHaber
PrinterVariables Ancho(8), PosLinea, SumaSaldo
PosLinea = PosLinea + 0.1
End If
DataReg.Close
RatonNormal
MensajeEncabData = ""
Printer.EndDoc
Exit Sub
Errorhandler:
    RatonNormal
    ErrorDeImpresion
    Exit Sub
End Sub

Public Sub Imprimir_Libro_Banco(DataT As Adodc)
Dim CadCta As String
Dim CtaSup As String
Dim CadCtaSup As String
Dim NuevoDoc As Boolean
Dim MesActual As Integer
Dim NumeroLineas As Single
 
Dim I As Integer

PorteLetra = 9
TipoLetra = TipoArialNarrow

On Error GoTo Errorhandler
Mensajes = "Seguro de Imprimir en:" & vbCrLf & Printer.DeviceName & "?"
Titulo = "IMPRESION DE LIBRO BANCO"
Bandera = False
Orientacion_Pagina = 2
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
RatonReloj
' Consultamos las cuentas de la tabla
ReDim Ancho(11) As Single
If DataT.Recordset.RecordCount > 0 Then
   DataT.Recordset.MoveFirst
Escala_Centimetro 2, TipoLetra, PorteLetra
InicioX = 0.5: InicioY = 0
Pagina = 1: Documento = 1
HoraSistema = Time
C = 10
Ancho(0) = 0.5   ' Fecha
Ancho(1) = 2     ' TP
Ancho(2) = 2.6   ' Numero
Ancho(3) = 4.1   ' Cheque/Deposito
Ancho(4) = 6     ' Beneficiario
Ancho(5) = 10    ' Concepto
Ancho(10) = LimiteAncho - 1
For I = 9 To 6 Step -1
    Ancho(I) = Ancho(I + 1) - 2.2
Next I
'''Ancho(6) = 17.5  ' Parcial ME
'''Ancho(7) = 19.5  ' Debe
'''Ancho(8) = 21.5  ' Haber
'''Ancho(9) = 23.5  ' Saldo
'''Ancho(10) = 25.5 ' Fin
'MsgBox Ancho(6) - Ancho(5)
NomCta = "": CtaSup = "": NomCtaSup = ""
Cta = DataT.Recordset.fields("Cta")
End If
Leer_Cta_Superior Cta
CadCtaSup = CtaSup & Space(10) & NomCtaSup
CadCta = Cta & Space(10) & NomCta

'MsgBox LimiteAlto & " ..."
'============================================================
If DataT.Recordset.RecordCount > 0 Then
With DataT.Recordset
.MoveFirst
Suma_ME = 0: SumaDebe = 0: SumaHaber = 0: SumaSaldo = 0
Encabezado Ancho(0), Ancho(10)
EncabLibroBanco1 CadCta, CadCtaSup, DataT
Printer.FontBold = False
Printer.FontItalic = False
Printer.FontName = TipoLetra
Printer.FontSize = PorteLetra
MesActual = FechaMes(.fields("Fecha"))
Mifecha = .fields("Fecha")
PrinterFields Ancho(0), PosLinea, .fields("Fecha"), True
Do While Not .EOF
   If PosLinea > LimiteAlto Then
      'MsgBox LimiteAlto
      PosLinea = PosLinea + 0.1
      Printer.Line (InicioX, PosLinea)-(Ancho(C), PosLinea), Negro
      Printer.NewPage
      CadCtaSup = CtaSup & Space(10) & NomCtaSup
      CadCta = Cta & Space(10) & NomCta
      Encabezado Ancho(0), Ancho(10)
      EncabLibroBanco1 CadCta, CadCtaSup, DataT
      Printer.FontBold = False
      Printer.FontItalic = False
      Printer.FontName = TipoLetra
      Printer.FontSize = PorteLetra
   End If
   Printer.FontBold = False
   Printer.FontItalic = False
   Printer.FontName = TipoLetra
   Printer.FontSize = PorteLetra
   If MesActual <> FechaMes(.fields("Fecha")) Then
      PosLinea = PosLinea + 0.1
      Printer.Line (InicioX, PosLinea)-(Ancho(C), PosLinea), Negro
      PosLinea = PosLinea + 0.05
      Printer.FontBold = True
      PrinterVariableTexto Ancho(0), PosLinea, "Fin de: ", MesesLetras(MesActual)
      MesActual = FechaMes(DataT.Recordset.fields("Fecha"))
      PrinterTexto Ancho(6) - 2, PosLinea, "TOTALES"
      If OpcCoop = False Then PrinterVariables Ancho(6), PosLinea, Suma_ME
      PrinterVariables Ancho(7), PosLinea, SumaDebe
      PrinterVariables Ancho(8), PosLinea, SumaHaber
      PrinterVariables Ancho(9), PosLinea, SumaSaldo
      SumaDebe = 0: SumaHaber = 0: Suma_ME = 0
      PosLinea = PosLinea + 0.4
      PrinterVariableTexto Ancho(0), PosLinea, "Inicio de: ", MesesLetras(MesActual)
      PosLinea = PosLinea + 0.4
      Printer.Line (InicioX, PosLinea)-(Ancho(C), PosLinea), Negro
      PosLinea = PosLinea + 0.1
      Printer.FontBold = False
      Printer.FontItalic = False
      Printer.FontName = TipoLetra
      Printer.FontSize = PorteLetra
   End If
   If Mifecha <> .fields("Fecha") Then
      PrinterFields Ancho(0), PosLinea, .fields("Fecha"), True
      Mifecha = .fields("Fecha")
   End If
   PrinterFields Ancho(1), PosLinea, .fields("TP"), True
   PrinterFields Ancho(2), PosLinea, .fields("Numero"), True
   PrinterFields Ancho(3), PosLinea, .fields("Cheq_Dep"), True
   PrinterFields Ancho(4), PosLinea, .fields("Cliente"), True
   If OpcCoop Then
      NumeroLineas = PrinterLineasMayor(Ancho(5), PosLinea, .fields("Concepto"), Redondear(Ancho(6) - Ancho(5)))
      If NumeroLineas > 1 Then
         For I = 1 To NumeroLineas
            For J = 0 To C
              If J <> 6 Then Printer.Line (Ancho(J), PosLinea - 0.1)-(Ancho(J), PosLinea + 0.5), Negro
            Next J
            PosLinea = PosLinea + 0.35
         Next
         PosLinea = PosLinea - 0.35
      End If
      If (.fields("Parcial_ME")) <> 0 Then
         PrinterFields Ancho(7), PosLinea, .fields("Debe_ME"), True
         PrinterFields Ancho(8), PosLinea, .fields("Haber_ME"), True
         PrinterFields Ancho(9), PosLinea, .fields("Saldo_ME"), True
         SumaDebe = SumaDebe + .fields("Debe_ME")
         SumaHaber = SumaHaber + .fields("Haber_ME")
      Else
         PrinterFields Ancho(7), PosLinea, .fields("Debe"), True
         PrinterFields Ancho(8), PosLinea, .fields("Haber"), True
         PrinterFields Ancho(9), PosLinea, .fields("Saldo"), True
         SumaDebe = SumaDebe + .fields("Debe")
         SumaHaber = SumaHaber + .fields("Haber")
      End If
      For J = 0 To C
         If J <> 6 Then Printer.Line (Ancho(J), PosLinea - 0.1)-(Ancho(J), PosLinea + 0.4), Negro
      Next J
      If .fields("T") <> Anulado Then
          SumaSaldo = .fields("Saldo")
          Suma_ME = .fields("Saldo_ME")
      End If
      If Suma_ME <> 0 Then SumaSaldo = Suma_ME
   Else
      NumeroLineas = PrinterLineasMayor(Ancho(5), PosLinea, .fields("Concepto"), Redondear(Ancho(6) - Ancho(5)))
      If NumeroLineas > 1 Then
      For I = 1 To NumeroLineas
         For J = 0 To C
           Printer.Line (Ancho(J), PosLinea - 0.1)-(Ancho(J), PosLinea + 0.5), Negro
         Next J
         PosLinea = PosLinea + 0.35
      Next
      PosLinea = PosLinea - 0.35
      End If
      PrinterFields Ancho(6), PosLinea, .fields("Parcial_ME"), True
      PrinterFields Ancho(7), PosLinea, .fields("Debe"), True
      PrinterFields Ancho(8), PosLinea, .fields("Haber"), True
      PrinterFields Ancho(9), PosLinea, .fields("Saldo"), True
      For J = 0 To C
         Printer.Line (Ancho(J), PosLinea - 0.1)-(Ancho(J), PosLinea + 0.4), Negro
      Next J
      SumaDebe = SumaDebe + .fields("Debe")
      SumaHaber = SumaHaber + .fields("Haber")
      If .fields("T") <> Anulado Then
          SumaSaldo = .fields("Saldo")
          Suma_ME = .fields("Saldo_ME")
      End If
   End If
   PosLinea = PosLinea + AltoLetra
  .MoveNext
Loop
End With
PosLinea = PosLinea + 0.1
Printer.Line (InicioX, PosLinea)-(Ancho(C), PosLinea), Negro
Printer.FontBold = True
PosLinea = PosLinea + 0.05
PrinterVariableTexto Ancho(0), PosLinea, "Fin de: ", MesesLetras(MesActual)
PrinterTexto Ancho(6) - 2, PosLinea, "TOTALES"
If OpcCoop = False Then PrinterVariables Ancho(6), PosLinea, Suma_ME
PrinterVariables Ancho(7), PosLinea, SumaDebe
PrinterVariables Ancho(8), PosLinea, SumaHaber
PrinterVariables Ancho(9), PosLinea, SumaSaldo
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

Public Sub ImprimirFormatoComprobantes(TComp As String, _
                                       AnioComp As Long, _
                                       Comp_No As Long, _
                                       PonerLineas As Boolean, _
                                       Optional Pos_Y_Comp As Single)
'Comienza las impresiones del formato del Comp en lineas
If PonerLineas Then
MargenIzq = 0.5: MargenSup = Pos_Y_Comp - 0.3: MargenDer = 19
If MargenSup <= 0 Then MargenSup = 0.01
InicioX = MargenIzq: InicioY = MargenSup
msg = Format$(MidStrg(CStr(AnioComp), 3, 2), "00") & "-" & Format$(Comp_No, "00000000")
'Imprimimos el formato del comprobante
'PrinterPaint MarcaAgua, 2, 3 + Pos_Y_Comp, 16, 8
Select Case TComp
'''  Case CompFactura
'''       MargenInf = AltoFactura
'''       PrinterPaint FA.LogoFactura, MargenIzq, MargenSup, MargenDer, MargenInf
'''       Msg = Format$(Comp_No, "000000")
  Case CompIngreso
       MargenInf = 7
       Dibujo = RutaSistema & "\FORMATOS\INGRESO.GIF"
       PrinterPaint Dibujo, MargenIzq, MargenSup, MargenDer, MargenInf
  Case CompEgreso
       MargenInf = 7
       Dibujo = RutaSistema & "\FORMATOS\EGRESO.GIF"
       PrinterPaint Dibujo, MargenIzq, MargenSup, MargenDer, MargenInf
  Case CompDiario
       MargenInf = 4.5
       Dibujo = RutaSistema & "\FORMATOS\DIARIO.GIF"
       PrinterPaint Dibujo, MargenIzq, MargenSup, MargenDer, MargenInf
       Printer.FontSize = 10
       PrinterVariables 17, 1.5 + Pos_Y_Comp, Pagina
  Case CompNotaDebito
       MargenInf = 4.5
       Dibujo = RutaSistema & "\FORMATOS\ND.GIF"
       PrinterPaint Dibujo, MargenIzq, MargenSup, MargenDer, MargenInf
       Printer.FontSize = 10
       PrinterVariables 17, 1.5 + Pos_Y_Comp, Pagina
  Case CompNotaCredito
       MargenInf = 4.5
       Dibujo = RutaSistema & "\FORMATOS\NC.GIF"
       PrinterPaint Dibujo, MargenIzq, MargenSup, MargenDer, MargenInf
       Printer.FontSize = 10
       PrinterVariables 17, 1.5 + Pos_Y_Comp, Pagina
  Case CompRetencion
       MargenInf = 9
       Dibujo = RutaSistema & "\FORMATOS\RETENCIO.GIF"
       PrinterPaint Dibujo, MargenIzq, MargenSup, MargenDer, MargenInf
End Select
'Imprimimos el Logotipo de la Empresa
PrinterPaint LogoTipo, 0.7, Pos_Y_Comp + 0.1, 4, 1.9
'imprimimos negrillas de titulos
EncabezadoEmpresa Pos_Y_Comp

'========================================================================
''Printer.FontSize = 12
''Printer.CurrentX = SetD(1).PosX - 0.2: Printer.CurrentY = SetD(1).PosY
''Printer.Print Msg
End If
Printer.FontBold = False: Printer.FontSize = 10
Pagina = Pagina + 1
End Sub

Public Sub Encabezado_Venc(Datas As Adodc, _
                              FormaImp As Byte, _
                              TipoRep As Integer)
Dim InicX As Single
Dim InicY As Single
Encabezado Ancho(0), Ancho(CantCampos)
PorteLetra = Printer.FontSize
LetraAnterior = Printer.FontName
Printer.FontName = TipoTimes
Printer.FontBold = True
If SQLMsg1 <> "" Then
   Printer.FontSize = 14
   PrinterTexto CentrarTexto(SQLMsg1), PosLinea, SQLMsg1
   PosLinea = PosLinea + 0.7
End If
Printer.FontSize = 10
If SQLMsg2 <> "" Then
   PrinterTexto CentrarTexto(SQLMsg2), PosLinea, SQLMsg2
   PosLinea = PosLinea + 0.6
End If
If SQLMsg3 <> "" Then
   PrinterTexto Ancho(0), PosLinea, SQLMsg3
   PosLinea = PosLinea + 0.6
End If
Printer.FontSize = 9
Printer.FontUnderline = True
'========================================================================
PrinterTexto Ancho(0), PosLinea, "Credito"
PrinterTexto Ancho(1), PosLinea, "Apellidos y Nombres"
PrinterTexto Ancho(2), PosLinea, "Cuenta_No"
If FormaImp = 2 Then
  PrinterTexto Ancho(3), PosLinea, "D i r e c c i ó n"
  PrinterTexto Ancho(4), PosLinea, "Sector"
  PrinterTexto Ancho(5), PosLinea, "Area"
End If
PrinterTexto Ancho(6), PosLinea, "T e l e f o n o s"
PrinterTexto Ancho(7), PosLinea, "Fecha"
PrinterTexto Ancho(8), PosLinea, "Cuota No"
'If TipoRep = 1 Then
   PrinterTexto Ancho(9), PosLinea, "A b o n o"
'Else
'   PrinterTexto Ancho(9), PosLinea, "C a p i t a l"
'End If
Printer.FontUnderline = False
'PrinterAllFieldsName CantCampos, PosLinea, Datas, True
PosLinea = PosLinea + 0.5
Printer.FontBold = False
Printer.FontSize = PorteLetra
Printer.FontName = LetraAnterior
End Sub

Public Sub Encabezado_SubCta_Venc(Vertical As Boolean)
Dim InicX As Single
Dim InicY As Single
Encabezado Ancho(0), Ancho(CantCampos)
PorteLetra = Printer.FontSize
LetraAnterior = Printer.FontName
Printer.FontName = TipoVerdana
Printer.FontBold = True
Printer.FontItalic = False
If SQLMsg1 <> "" Then
   Printer.FontSize = 10
   PrinterTexto CentrarTexto(SQLMsg1), PosLinea, SQLMsg1
   PosLinea = PosLinea + 0.5
End If
Printer.FontSize = 8
If SQLMsg2 <> "" And SQLMsg3 <> "" Then
   Cadena = TrimStrg(SQLMsg3 & " " & SQLMsg2)
   PrinterTexto CentrarTexto(Cadena), PosLinea, Cadena
   PosLinea = PosLinea + 0.5
End If
Printer.FontSize = 7
Printer.FontUnderline = True
'========================================================================
PrinterTexto Ancho(0), PosLinea, "C U E N T A"
PrinterTexto Ancho(1), PosLinea, "P E R S O N A"
PrinterTexto Ancho(2), PosLinea, "Telefono"
PrinterTexto Ancho(3), PosLinea, "Factura"
PrinterTexto Ancho(4), PosLinea, "TP"
PrinterTexto Ancho(5), PosLinea, "Numero"
PrinterTexto Ancho(6), PosLinea, "Fecha"
PrinterTexto Ancho(7), PosLinea, "Fecha_V"
If Vertical Then
   PrinterTexto Ancho(8), PosLinea, "Total"
   PrinterTexto Ancho(9), PosLinea, "Abonos"
   PrinterTexto Ancho(10), PosLinea, "Saldo"
Else
   PrinterTexto Ancho(8), PosLinea, "Beneficiario"
   PrinterTexto Ancho(9), PosLinea, "Total"
   PrinterTexto Ancho(10), PosLinea, "Abonos"
   PrinterTexto Ancho(11), PosLinea, "Saldo"
End If
PosLinea = PosLinea + 0.4
Printer.FontUnderline = False
Printer.FontBold = False
Printer.FontSize = PorteLetra
Printer.FontName = LetraAnterior
End Sub

Public Sub Encabezado_SubCta_IE(Benef As Boolean)
Dim InicX As Single
Dim InicY As Single
Encabezado Ancho(0), Ancho(CantCampos)
PorteLetra = Printer.FontSize
LetraAnterior = Printer.FontName
Printer.FontName = TipoVerdana
Printer.FontBold = True
Printer.FontItalic = False
If SQLMsg1 <> "" Then
   Printer.FontSize = 10
   PrinterTexto CentrarTexto(SQLMsg1), PosLinea, SQLMsg1
   PosLinea = PosLinea + 0.5
End If
Printer.FontSize = 8
If SQLMsg2 <> "" And SQLMsg3 <> "" Then
   Cadena = TrimStrg(SQLMsg3 & " " & SQLMsg2)
   PrinterTexto CentrarTexto(Cadena), PosLinea, Cadena
   PosLinea = PosLinea + 0.5
End If
Printer.FontSize = 7
Printer.FontUnderline = True
'========================================================================
PrinterTexto Ancho(0), PosLinea, "C U E N T A"
PrinterTexto Ancho(1), PosLinea, "S U B C U E N T A"
If Benef Then
   PrinterTexto Ancho(2), PosLinea, "Detalle Auxiliar"
   PrinterTexto Ancho(3), PosLinea, "Total"
Else
   PrinterTexto Ancho(2), PosLinea, "Total"
End If
PosLinea = PosLinea + 0.4
Printer.FontUnderline = False
Printer.FontBold = False
Printer.FontSize = PorteLetra
Printer.FontName = LetraAnterior
End Sub

Public Sub EncabBalance()
'Iniciamos la impresion
PorteLetra = Printer.FontSize
LetraAnterior = Printer.FontName
Printer.FontName = TipoTimes
Dibujo = RutaSistema & "\FORMATOS\BALANCE.GIF"
PrinterPaint Dibujo, 0.7, PosLinea, 19.5, 2
Printer.FontBold = True: Printer.FontSize = 16
Cadena = "BALANCE DE COMPROBACION"
PrinterTexto CentrarTexto(Cadena), PosLinea + 0.1, Cadena
Printer.FontSize = 10
SQLMsg1 = "Periodo Desde: " & FechaStrgCorta(FechaIni) & "   hasta: " & FechaStrgCorta(FechaFin)
PrinterTexto CentrarTexto(SQLMsg1), PosLinea + 0.7, SQLMsg1
Printer.FontBold = False
Printer.FontSize = PorteLetra
Printer.FontName = LetraAnterior
PosLinea = PosLinea + 2
End Sub

Public Sub LeerCodigoCta(TCodigo As TextBox, _
                         TCuenta As TextBox, _
                         TValor As TextBox, _
                         DBLists As DataList, _
                         FrmAsigna As Frame, _
                         TOpcTM As TextBox, _
                         TOpcDH As TextBox)
Dim AdoBuscarCta As ADODB.Recordset
Dim SSQLCodigo As String
Dim CodExterno As String
    Cuenta = Ninguno: Codigo = Ninguno: TipoCta = "G"
    SubCta = Ninguno: Moneda_US = False
    If UCaseStrg(MidStrg(TCodigo.Text, 1, 1)) = "E" Then
       CodExterno = TrimStrg(MidStrg(TCodigo.Text, 2, Len(TCodigo.Text)))
       If CodExterno = "" Then CodExterno = ".."
       sSQL = "SELECT * " _
            & "FROM Catalogo_Cuentas " _
            & "WHERE SUBSTRING(Codigo_Ext,4,LEN(Codigo_Ext)) = '" & CodExterno & "' " _
            & "AND Item = '" & NumEmpresa & "' " _
            & "AND Periodo = '" & Periodo_Contable & "' "
       Select_AdoDB AdoBuscarCta, sSQL
       If AdoBuscarCta.RecordCount > 0 Then
          Codigo = AdoBuscarCta.fields("Codigo")
          Cuenta = AdoBuscarCta.fields("Cuenta")
          SubCta = AdoBuscarCta.fields("TC")
          Moneda_US = AdoBuscarCta.fields("ME")
          TCodigo.Text = Codigo
          TCuenta.Text = Cuenta
          FrmAsigna.Visible = True
          If OpcCoop Then
             TOpcDH.SetFocus
          Else
             If SubCta = "BA" Then
             Else
                TOpcTM.SetFocus
             End If
          End If
       Else
          MsgBox "Warning:" & vbCrLf & " Esta Cuenta o Codigo Externo No existen, vuelva a ingresar."
          TCodigo.SetFocus
       End If
    ElseIf Val(TCodigo.Text) = 0 Then
       DBLists.Visible = True
       DBLists.SetFocus
       TCodigo.Text = Codigo
       TCuenta.Text = Cuenta
    ElseIf Val(TCodigo.Text) <> -1 Then
       sSQL = "SELECT * " _
            & "FROM Catalogo_Cuentas " _
            & "WHERE Codigo = '" & TCodigo.Text & "' " _
            & "AND DG <> 'G' " _
            & "AND Item = '" & NumEmpresa & "' " _
            & "AND Periodo = '" & Periodo_Contable & "' "
       Select_AdoDB AdoBuscarCta, sSQL
       If AdoBuscarCta.RecordCount > 0 Then
          Codigo = AdoBuscarCta.fields("Codigo")
          Cuenta = AdoBuscarCta.fields("Cuenta")
          SubCta = AdoBuscarCta.fields("TC")
          Moneda_US = AdoBuscarCta.fields("ME")
          TCodigo.Text = Codigo
          TCuenta.Text = Cuenta
          FrmAsigna.Visible = True
          If OpcCoop Then TOpcDH.SetFocus Else TOpcTM.SetFocus
       ElseIf IsNumeric(TCodigo.Text) Then
          
          If Val(TCodigo.Text) > 32000 Then TCodigo.Text = "32000"
          sSQL = "SELECT * " _
               & "FROM Catalogo_Cuentas " _
               & "WHERE Clave = " & CInt(TCodigo.Text) & " " _
               & "AND Item = '" & NumEmpresa & "' " _
               & "AND Periodo = '" & Periodo_Contable & "' "
          Select_AdoDB AdoBuscarCta, sSQL
          If AdoBuscarCta.RecordCount > 0 Then
            'MsgBox TCodigo.Text
             Codigo = AdoBuscarCta.fields("Codigo")
             Cuenta = AdoBuscarCta.fields("Cuenta")
             SubCta = AdoBuscarCta.fields("TC")
             Moneda_US = AdoBuscarCta.fields("ME")
             TCodigo.Text = Codigo
             TCuenta.Text = Cuenta
             FrmAsigna.Visible = True
             If OpcCoop Then
                TOpcDH.SetFocus
             Else
                If SubCta = "BA" Then
                Else
                   TOpcTM.SetFocus
                End If
             End If
          Else
             MsgBox "Warning:" & vbCrLf & " Esta Cuenta o Clave No existen, vuelva a ingresar."
             TCodigo.SetFocus
          End If
       Else
          MsgBox "Warning:" & vbCrLf & " Este dato no es correcto, vuelva a ingresar."
          TCodigo.SetFocus
       End If
       AdoBuscarCta.Close
    End If
End Sub

Public Sub LeerBanco(Datas As Adodc, _
                     CodigoBanco As String)
  NombreBanco = Ninguno: CuentaBanco = Ninguno
  CtaCteNo = Ninguno: Moneda_US = False
  Codigo = SinEspaciosIzq(CodigoBanco)
  sSQL = "SELECT * " _
       & "FROM Catalogo_Cuentas " _
       & "WHERE Codigo = '" & Codigo & "' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' "
  Select_Adodc Datas, sSQL
  With Datas.Recordset
   If .RecordCount > 0 Then
      'CtaCteNo = .Fields("No_Cta")
       CuentaBanco = .fields("Codigo")
       NombreBanco = .fields("Cuenta")
       Moneda_US = .fields("ME")
   End If
  End With
End Sub

'Codigo_Catalogo = Cuenta, TipoCta, SubCta, TipoPago = "01", Moneda_US
'Retorna el Codigo del Catalogo de Cuentas
'--------------------------------------------------------
Function Leer_Cta_Catalogo(CodigoCta As String) As String
Dim DataReg As ADODB.Recordset
Dim Codigo_Catalogo As String
  RatonReloj
  Cuenta = Ninguno
  Codigo_Catalogo = Ninguno
  CodRolPago = Ninguno
  TipoCta = "G"
  SubCta = "N"
  TipoPago = "01"
  Moneda_US = False
 'MsgBox CodigoCtaDim DataReg As ADODB.Recordset
  If Val(MidStrg(CodigoCta, 1, 1)) >= 1 Then
     Set DataReg = New ADODB.Recordset
     DataReg.CursorType = adOpenStatic
     DataReg.CursorLocation = adUseClient
     sSQL = "SELECT Codigo, Cuenta, TC, ME, DG, Tipo_Pago, Cod_Rol_Pago " _
          & "FROM Catalogo_Cuentas " _
          & "WHERE '" & CodigoCta & "' IN (Codigo, Codigo_Ext) " _
          & "AND Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' "
     DataReg.open sSQL, AdoStrCnn, , , adCmdText
     With DataReg
      If .RecordCount > 0 Then
          Codigo_Catalogo = .fields("Codigo")
          Cuenta = .fields("Cuenta")
          SubCta = .fields("TC")
          Moneda_US = .fields("ME")
          TipoCta = .fields("DG")
          TipoPago = .fields("Tipo_Pago")
          CodRolPago = .fields("Cod_Rol_Pago")
          If Val(TipoPago) <= 0 Then TipoPago = "01"
      End If
     End With
     DataReg.Close
  End If
  Leer_Cta_Catalogo = Codigo_Catalogo
  RatonNormal
End Function

Function Leer_SubCta_Modulo(CodigoCta As String) As String
Dim DataReg As ADODB.Recordset
Dim Codigo_Catalogo As String
  RatonReloj
  Sub_Cuenta = Ninguno
  Codigo_Catalogo = Ninguno
  If Len(CodigoCta) > 1 Then
     Set DataReg = New ADODB.Recordset
     DataReg.CursorType = adOpenStatic
     DataReg.CursorLocation = adUseClient
     sSQL = "SELECT Codigo,Detalle " _
          & "FROM Catalogo_SubCtas " _
          & "WHERE Codigo = '" & CodigoCta & "' " _
          & "AND Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' "
     DataReg.open sSQL, AdoStrCnn, , , adCmdText
     With DataReg
      If .RecordCount > 0 Then
          Codigo_Catalogo = .fields("Codigo")
          Sub_Cuenta = .fields("Detalle")
      End If
     End With
     DataReg.Close
      
     If Codigo_Catalogo = Ninguno Then
        Set DataReg = New ADODB.Recordset
        DataReg.CursorType = adOpenStatic
        DataReg.CursorLocation = adUseClient
        sSQL = "SELECT Codigo,Detalle " _
             & "FROM Catalogo_SubCtas " _
             & "WHERE Detalle = '" & CodigoCta & "' " _
             & "AND Item = '" & NumEmpresa & "' " _
             & "AND Periodo = '" & Periodo_Contable & "' "
        DataReg.open sSQL, AdoStrCnn, , , adCmdText
        With DataReg
         If .RecordCount > 0 Then
             Codigo_Catalogo = .fields("Codigo")
             Sub_Cuenta = .fields("Detalle")
         End If
        End With
        DataReg.Close
     End If
     If Codigo_Catalogo = Ninguno Then
        Set DataReg = New ADODB.Recordset
        DataReg.CursorType = adOpenStatic
        DataReg.CursorLocation = adUseClient
        sSQL = "SELECT Codigo,Cliente " _
             & "FROM Clientes " _
             & "WHERE Codigo = '" & CodigoCta & "' "
        DataReg.open sSQL, AdoStrCnn, , , adCmdText
        With DataReg
         If .RecordCount > 0 Then
             Codigo_Catalogo = .fields("Codigo")
             Sub_Cuenta = .fields("Cliente")
         End If
        End With
        DataReg.Close
     End If
  End If
  Leer_SubCta_Modulo = Codigo_Catalogo
  RatonNormal
End Function

'ExisteCtas(4) As String

Public Sub VerSiExisteCta(ExisteCtas() As String)
Dim DataReg As ADODB.Recordset
Dim Indc1 As Byte
Dim NCtas As String

  RatonReloj
  NCtas = ""
  For Indc1 = 0 To UBound(ExisteCtas) - 1
      If Len(ExisteCtas(Indc1)) > 1 And InStr(NCtas, ExisteCtas(Indc1)) = 0 Then NCtas = NCtas & "'" & ExisteCtas(Indc1) & "',"
  Next Indc1
  If NCtas <> "" Then NCtas = MidStrg(NCtas, 1, Len(NCtas) - 1)
  If Len(NCtas) > 1 Then
     Set DataReg = New ADODB.Recordset
     DataReg.CursorType = adOpenStatic
     DataReg.CursorLocation = adUseClient
     sSQL = "SELECT Codigo " _
          & "FROM Catalogo_Cuentas " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "AND Codigo IN (" & NCtas & ") "
     DataReg.open sSQL, AdoStrCnn, , , adCmdText
     NCtas = ""
     With DataReg
      If .RecordCount > 0 Then
          For Indc1 = 0 To UBound(ExisteCtas) - 1
              If Len(ExisteCtas(Indc1)) > 1 Then
                 .MoveFirst
                 .Find ("Codigo = '" & ExisteCtas(Indc1) & "' ")
                  If .EOF Then NCtas = NCtas & "'" & ExisteCtas(Indc1) & "'" & vbCrLf
              End If
          Next Indc1
      End If
     End With
     DataReg.Close
  End If
  RatonNormal
  If Len(NCtas) > 1 Then
     MsgBox "Falta de setear la(s) cuenta(s) siguiente(s):" & vbCrLf _
          & NCtas & "No existe en el CATALOGO DE CUENTAS, " _
          & "debe Setearla en el Módulo de SETEOS en la opcion de MANTENIMIENTO"
  End If
End Sub

Public Sub TituloLibroCaja()
   If SQLMsg1 <> "" Then
      Printer.FontSize = 14
      PrinterTexto CentrarTexto(SQLMsg1), PosLinea, SQLMsg1
      PosLinea = PosLinea + 0.7
   End If
   Printer.FontSize = 10
   If SQLMsg2 <> "" Then
      PrinterTexto CentrarTexto(SQLMsg2), PosLinea, SQLMsg2
      PosLinea = PosLinea + 0.6
   End If
End Sub

Public Sub EncabezadoLibroReciboCaja(Opc As Integer)
  'Printer.Line (Ancho(0), PosLinea)-(Ancho(CantCampos), PosLinea), NEGRO
  'PosLinea = PosLinea + 0.05
  Printer.FontUnderline = True
  PrinterTexto Ancho(1), PosLinea, "F e c h a "
  PrinterTexto Ancho(2), PosLinea, "Número "
  PrinterTexto Ancho(3), PosLinea, "Beneficiario  "
  PrinterTexto Ancho(4), PosLinea, "C o n c e p t o  "
  If Opc = 1 Then
     PrinterTexto Ancho(5) + 0.1, PosLinea, "I n g r e s o"
  ElseIf Opc = 2 Then
     PrinterTexto Ancho(5) + 0.1, PosLinea, "E g r e s o  "
  Else
     PrinterTexto Ancho(5) + 0.1, PosLinea, "I n g r e s o"
     PrinterTexto Ancho(6) + 0.1, PosLinea, "E g r e s o  "
     PrinterTexto Ancho(7) + 0.1, PosLinea, "S a l d o    "
  End If
  Printer.FontUnderline = False
  PosLinea = PosLinea + 0.4
  'Printer.Line (Ancho(0), PosLinea)-(Ancho(CantCampos), PosLinea), NEGRO
  'PosLinea = PosLinea + 0.05
End Sub

Public Sub Actualizar_SubCta_12_Meses(Cta12 As String, _
                                      SubCta12 As String, _
                                      VSTot() As Currency)
Dim CampoMes As String
  
  Total = 0
  For NoMeses = 1 To 12
      Total = Total + VSTot(NoMeses)
  Next NoMeses
  Total = Redondear(Total, 2)
  
  sSQL = "UPDATE Saldo_Diarios " _
       & "SET Total = Total + " & Total & ", "
  For NoMeses = 1 To 12
      CampoMes = MesesLetras(NoMeses)
      sSQL = sSQL & CampoMes & " = " & CampoMes & " + " & VSTot(NoMeses) & ", "
  Next NoMeses
  sSQL = sSQL & "Diferencia = Presupuesto - " & Total & " " _
       & "WHERE Cta_Aux = '" & Cta12 & "' " _
       & "AND CodigoC = '" & SubCta12 & "' " _
       & "AND Item = '" & NumItemTemp & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND TP = 'E12M' "
  Ejecutar_SQL_SP sSQL
  If Cta12 = "5.1.08.02" Then MsgBox "Actualizar_SubCta_12_Meses: " & vbCrLf & sSQL
End Sub

Public Sub Insertar_SubCta_12_Meses(Codigo12 As String, _
                                    Cuenta12 As String, _
                                    VSTot() As Currency)
     Cta_Sup = Codigo12
     Do While Len(Cta_Sup) > 0
       'MsgBox Cta_Sup & vbCrLf & Len(Cta_Sup)
        Actualizar_SubCta_12_Meses Cta_Sup, Cuenta12, VSTot
        Cta_Sup = CodigoCuentaSup(Cta_Sup)
        If Cta_Sup = "0" Then Cta_Sup = ""
     Loop
End Sub

Public Sub EliminarComprobantes(C1 As Comprobantes)
  SQL1 = "DELETE * " _
       & "FROM Comprobantes " _
       & "WHERE TP = '" & C1.TP & "' " _
       & "AND Numero = " & C1.Numero & " " _
       & "AND Item = '" & C1.Item & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' "
  Ejecutar_SQL_SP SQL1
  SQL1 = "DELETE * " _
       & "FROM Transacciones " _
       & "WHERE TP = '" & C1.TP & "' " _
       & "AND Numero = " & C1.Numero & " " _
       & "AND Item = '" & C1.Item & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' "
  Ejecutar_SQL_SP SQL1
  SQL1 = "DELETE * " _
       & "FROM Trans_SubCtas " _
       & "WHERE TP = '" & C1.TP & "' " _
       & "AND Numero = " & C1.Numero & " " _
       & "AND Item = '" & C1.Item & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' "
  Ejecutar_SQL_SP SQL1
  SQL1 = "DELETE * " _
       & "FROM Trans_Kardex " _
       & "WHERE TP = '" & C1.TP & "' " _
       & "AND Numero = " & C1.Numero & " " _
       & "AND Item = '" & C1.Item & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND LEN(TC) <= 1 "
  Ejecutar_SQL_SP SQL1
  SQL1 = "DELETE * " _
       & "FROM Trans_Compras " _
       & "WHERE TP = '" & C1.TP & "' " _
       & "AND Numero = " & C1.Numero & " " _
       & "AND Item = '" & C1.Item & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' "
  Ejecutar_SQL_SP SQL1
  SQL1 = "DELETE * " _
       & "FROM Trans_Air " _
       & "WHERE TP = '" & C1.TP & "' " _
       & "AND Numero = " & C1.Numero & " " _
       & "AND Item = '" & C1.Item & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' "
  Ejecutar_SQL_SP SQL1
  SQL1 = "DELETE * " _
       & "FROM Trans_Ventas " _
       & "WHERE TP = '" & C1.TP & "' " _
       & "AND Numero = " & C1.Numero & " " _
       & "AND Item = '" & C1.Item & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' "
  Ejecutar_SQL_SP SQL1
  SQL1 = "DELETE * " _
       & "FROM Trans_Exportaciones " _
       & "WHERE TP = '" & C1.TP & "' " _
       & "AND Numero = " & C1.Numero & " " _
       & "AND Item = '" & C1.Item & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' "
  Ejecutar_SQL_SP SQL1
  SQL1 = "DELETE * " _
       & "FROM Trans_Importaciones " _
       & "WHERE TP = '" & C1.TP & "' " _
       & "AND Numero = " & C1.Numero & " " _
       & "AND Item = '" & C1.Item & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' "
  Ejecutar_SQL_SP SQL1
  
  SQL1 = "DELETE * " _
       & "FROM Trans_Rol_Pagos " _
       & "WHERE TP = '" & C1.TP & "' " _
       & "AND Numero = " & C1.Numero & " " _
       & "AND Item = '" & C1.Item & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' "
  Ejecutar_SQL_SP SQL1
End Sub

Public Sub ImprimirBalanceSubCta(Datas As Adodc, _
                                 ValorTotal As Currency, _
                                 FormaImp As Byte, _
                                 SizeLetra As Integer)
On Error GoTo Errorhandler
RatonReloj
Mensajes = "Seguro de Imprimir en:" & vbCrLf & Printer.DeviceName & "?"
Titulo = "IMPRESION"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
InicioX = 0.5: InicioY = 0
DataAnchoCampos InicioX, Datas, SizeLetra, TipoTimes, FormaImp
Ancho(0) = 0.5  'TC
Ancho(1) = 0.5  'Detalles
Ancho(2) = 4.5  'Cta
Ancho(3) = 6.5  'Cuentas
Ancho(4) = 11   'Saldo Anterior
Ancho(5) = 13   'Ingresos
Ancho(6) = 15   'Egresos
Ancho(7) = 17   'Saldo Actual
Ancho(8) = 19   'Fin
Pagina = 1
'Iniciamos la impresion
Printer.FontBold = False
With Datas.Recordset
    .MoveFirst
     EncabezadoData Datas
     Printer.FontSize = SizeLetra
     Cadena = .fields("Detalles")
     Printer.FontBold = True
     Printer.FontUnderline = True
     Printer.FontSize = 6.5
     PrinterFields Ancho(1), PosLinea, .fields("Detalles")
     Printer.FontSize = SizeLetra
     Printer.FontBold = False
     Printer.FontUnderline = False
     Do While Not .EOF
        Printer.FontBold = False
        If Cadena <> .fields("Detalles") Then
           Printer.FontBold = True
           Printer.FontUnderline = True
           Printer.FontSize = 6.5
           PrinterFields Ancho(1), PosLinea, .fields("Detalles")
           Printer.FontSize = SizeLetra
           Printer.FontBold = False
           Printer.FontUnderline = False
           Cadena = .fields("Detalles")
        End If
        If MidStrg(.fields("Cta"), 1, 1) <> "I" Then
           PrinterFields Ancho(2), PosLinea, .fields("Cta")
        Else
           Printer.Line (Ancho(0), PosLinea)-(Ancho(CantCampos), PosLinea), Negro
           PosLinea = PosLinea + 0.1
           Printer.FontBold = True
        End If
        PrinterFields Ancho(3), PosLinea, .fields("Cuentas")
        PrinterFields Ancho(4), PosLinea, .fields("Saldo_Anterior")
        PrinterFields Ancho(5), PosLinea, .fields("Ingresos")
        PrinterFields Ancho(6), PosLinea, .fields("Egresos")
        PrinterFields Ancho(7), PosLinea, .fields("Saldo_Actual")
        If MidStrg(.fields("Cta"), 1, 1) = "I" Then PosLinea = PosLinea + 0.1
        'PrinterAllFields CantCampos, PosLinea, Datas, True
        PosLinea = PosLinea + 0.36
        If PosLinea >= LimiteAlto Then
           Printer.Line (Ancho(0), PosLinea)-(Ancho(CantCampos), PosLinea), Negro
           Printer.NewPage
           EncabezadoData Datas
           Printer.FontSize = SizeLetra
        End If
       .MoveNext
     Loop
    .MoveFirst
End With
Printer.Line (InicioX, PosLinea)-(Ancho(CantCampos), PosLinea), Negro
PosLinea = PosLinea + 0.05
Printer.Line (InicioX, PosLinea)-(Ancho(CantCampos), PosLinea), Negro
PosLinea = PosLinea + 0.05
Printer.FontBold = True
Printer.FontSize = SizeLetra + 1
PrinterVariables Ancho(7) - 0.2, PosLinea, ValorTotal
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

Public Sub Imprimir_Entradas_Salidas(Datas As Adodc, _
                                     FormaImp As Byte, _
                                     SizeLetra As Integer, _
                                     Opcion_Imp As Integer, _
                                     Optional EsCampoCorto As Boolean)
On Error GoTo Errorhandler
RatonReloj
Mensajes = "Seguro de Imprimir en:" & vbCrLf & Printer.DeviceName & "?"
Titulo = "IMPRESION"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
InicioX = 0.5: InicioY = 0
DataAnchoCampos InicioX, Datas, SizeLetra, TipoCondensed, Orientacion_Pagina, EsCampoCorto
If FormaImp <= 1 Then
   Ancho(0) = 0.5
   Ancho(1) = 5
   Ancho(2) = 5.5
   Tasa = 5.8
   For I = 3 To CantCampos
       Tasa = Tasa + 0.7
       Ancho(I) = Tasa
   Next I
   Ancho(CantCampos) = Tasa + 0.1
End If
Pagina = 1
'Iniciamos la impresion
Printer.FontBold = False
With Datas.Recordset
    .MoveFirst
     EncabezadoData Datas
     Printer.FontName = TipoCondensed
     Printer.FontSize = SizeLetra
     NombreCliente = .fields("Cliente")
     PrinterTexto Ancho(0), PosLinea, NombreCliente
     Do While Not .EOF
        If NombreCliente <> .fields("Cliente") Then
           NombreCliente = .fields("Cliente")
           PrinterTexto Ancho(0), PosLinea, NombreCliente
        End If
        For I = 1 To CantCampos - 1
            If Opcion_Imp <= 2 Then
               PrinterTexto Ancho(I), PosLinea, UCaseStrg(MidStrg(.fields(I), 1, 3))
            Else
               If I <= 2 Then
                  PrinterFields Ancho(I), PosLinea, .fields(I)
               Else
                  Cod_Benef = " "
                 'MsgBox "..."
                  If Opcion_Imp = 0 Then
                     If .fields(I) <> Ninguno Then Cod_Benef = "Ok"
                     PrinterTexto Ancho(I), PosLinea, Cod_Benef
                  ElseIf Opcion_Imp = 1 Then
                     PrinterFields Ancho(I), PosLinea, .fields(I)
                  Else
                     PrinterFields Ancho(I), PosLinea, .fields(I)
                  End If
               End If
            End If
        Next I
        PosLinea = PosLinea + 0.36
        If PosLinea >= LimiteAlto Then
           Printer.Line (Ancho(0), PosLinea)-(Ancho(CantCampos), PosLinea), Negro
           Printer.NewPage
           EncabezadoData Datas
           Printer.FontSize = SizeLetra
        End If
       .MoveNext
     Loop
End With
Printer.Line (InicioX, PosLinea)-(Ancho(CantCampos), PosLinea), Negro
PosLinea = PosLinea + 0.05
Printer.Line (InicioX, PosLinea)-(Ancho(CantCampos), PosLinea), Negro
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

Public Sub Imprimir_Saldos_SubCtas_Vence_Temporizada(Datas As Adodc, _
                                                     Optional EsCampoCorto As Boolean)
Dim SizeLetra As Integer
On Error GoTo Errorhandler
SizeLetra = 6
Orientacion_Pagina = 2
Mensajes = "Seguro de Imprimir en:" & vbCrLf & Printer.DeviceName & "?"
Titulo = "IMPRESION"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
RatonReloj
InicioX = 0.5: InicioY = 0
'Orientacion_Pagina
DataAnchoCampos InicioX, Datas, SizeLetra, TipoVerdana, 1, EsCampoCorto
Ancho(0) = 0.5    ' Cuenta
Ancho(1) = 0.5    ' Cliente
Ancho(2) = 2.2    ' Fecha
Ancho(3) = 3.9    ' Factura
Ancho(4) = 5.5    ' 1  -  7
Ancho(5) = 7.9    ' 8  - 30
Ancho(6) = 9.9    ' 31 - 60
Ancho(7) = 11.9   ' 61 - 90
Ancho(8) = 13.9   ' 91 - 180
Ancho(9) = 15.9   ' 181 - 360
Ancho(10) = 17.9  ' > 360
Ancho(11) = 19.5  ' Fin
CantCampos = 11
Ancho(CantCampos) = Ancho(11)
Pagina = 1
'Iniciamos la impresion
Printer.FontBold = False
SizeLetra = 9
Dim VSubTotales(6) As Double
Dim VTotales(6) As Double
LimiteAlto = LimiteAlto - 0.4
With Datas.Recordset
    .MoveFirst
     For I = 0 To 5
         VTotales(I) = 0
         VSubTotales(I) = 0
     Next I
     Total = 0: Monto_Total = 0
     Encabezados
     NombreCliente = .fields("Cliente")
     EncabezadoSaldosTemp
     Printer.FontSize = SizeLetra
     Do While Not .EOF
        If TrimStrg(NombreCliente) <> TrimStrg(.fields("Cliente")) Then
           PosLinea = PosLinea + 0.05
           Imprimir_Linea_H PosLinea, Ancho(0), Ancho(CantCampos), Negro
           PosLinea = PosLinea + 0.05
           PrinterTexto Ancho(1), PosLinea, "S U B T O T A L E S"
           PrinterVariables Ancho(4), PosLinea, Total
           PrinterVariables Ancho(5), PosLinea, VSubTotales(1)
           PrinterVariables Ancho(6), PosLinea, VSubTotales(2)
           PrinterVariables Ancho(7), PosLinea, VSubTotales(3)
           PrinterVariables Ancho(8), PosLinea, VSubTotales(4)
           PrinterVariables Ancho(9), PosLinea, VSubTotales(5)
           PosLinea = PosLinea + 0.4
           NombreCliente = .fields("Cliente")
           EncabezadoSaldosTemp
           Printer.FontSize = SizeLetra
           For I = 0 To 6
               VSubTotales(I) = 0
           Next I
           Total = 0
        End If
        If PosLinea >= LimiteAlto Then
           Imprimir_Linea_H PosLinea, Ancho(0), Ancho(CantCampos)
           Printer.NewPage
           Encabezados
           EncabezadoSaldosTemp
           Printer.FontSize = SizeLetra
        End If
        'PrinterFields Ancho(0), PosLinea, .Fields("T")
        PrinterFields Ancho(1), PosLinea, .fields("Fecha")
        PrinterFields Ancho(2), PosLinea, .fields("Fecha_Venc")
        PrinterFields Ancho(3), PosLinea, .fields("Factura")
        PrinterFields Ancho(4), PosLinea, .fields("Total")
        PrinterFields Ancho(5), PosLinea, .fields("Vencido")
        PrinterFields Ancho(6), PosLinea, .fields("Ven_1_a_30")
        PrinterFields Ancho(7), PosLinea, .fields("Ven_31_a_60")
        PrinterFields Ancho(8), PosLinea, .fields("Ven_61_a_90")
        PrinterFields Ancho(9), PosLinea, .fields("Mas_91")
        PrinterFields Ancho(10), PosLinea, .fields("Cta")
        
        Total = Total + .fields("Total")
        Monto_Total = Monto_Total + .fields("Total")
        VTotales(1) = VTotales(1) + .fields("Vencido")
        VTotales(2) = VTotales(2) + .fields("Ven_1_a_30")
        VTotales(3) = VTotales(3) + .fields("Ven_31_a_60")
        VTotales(4) = VTotales(4) + .fields("Ven_61_a_90")
        VTotales(5) = VTotales(5) + .fields("Mas_91")
        
        VSubTotales(1) = VSubTotales(1) + .fields("Vencido")
        VSubTotales(2) = VSubTotales(2) + .fields("Ven_1_a_30")
        VSubTotales(3) = VSubTotales(3) + .fields("Ven_31_a_60")
        VSubTotales(4) = VSubTotales(4) + .fields("Ven_61_a_90")
        VSubTotales(5) = VSubTotales(5) + .fields("Mas_91")
        For J = 1 To 11
            Printer.Line (Ancho(J), PosLinea - 0.05)-(Ancho(J), PosLinea + 0.36), Negro
        Next J
        PosLinea = PosLinea + 0.35
        If PosLinea >= LimiteAlto Then
           Imprimir_Linea_H PosLinea, Ancho(0), Ancho(CantCampos)
           Printer.NewPage
           Encabezados
           EncabezadoSaldosTemp
           Printer.FontSize = SizeLetra
        End If
       .MoveNext
     Loop
End With
PosLinea = PosLinea + 0.05
Imprimir_Linea_H PosLinea, Ancho(0), Ancho(CantCampos), Negro
PosLinea = PosLinea + 0.05
PrinterTexto Ancho(1), PosLinea, "S U B T O T A L E S"
PrinterVariables Ancho(4), PosLinea, Total
PrinterVariables Ancho(5), PosLinea, VSubTotales(1)
PrinterVariables Ancho(6), PosLinea, VSubTotales(2)
PrinterVariables Ancho(7), PosLinea, VSubTotales(3)
PrinterVariables Ancho(8), PosLinea, VSubTotales(4)
PrinterVariables Ancho(9), PosLinea, VSubTotales(5)
PosLinea = PosLinea + 0.35
Imprimir_Linea_H PosLinea, Ancho(0), Ancho(CantCampos), Negro, True
PosLinea = PosLinea + 0.05
PrinterTexto Ancho(1), PosLinea, "T O T A L E S"
PrinterVariables Ancho(4), PosLinea, Monto_Total
PrinterVariables Ancho(5), PosLinea, VTotales(1)
PrinterVariables Ancho(6), PosLinea, VTotales(2)
PrinterVariables Ancho(7), PosLinea, VTotales(3)
PrinterVariables Ancho(8), PosLinea, VTotales(4)
PrinterVariables Ancho(9), PosLinea, VTotales(5)
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

Public Sub Imprimir_Cheque_No(PosLinea1 As Single)
'Impresion solo cheque
     NombreCiudad = Leer_Campo_Empresa("Ciudad")
     Pagina = 1
     Printer.FontBold = True
     Printer.FontSize = 10
     Printer.FontName = TipoCourier
     Total = Valor
     Printer.FontSize = SetD(2).Tamaño
     PrinterTexto SetD(2).PosX, PosLinea1 + SetD(2).PosY, Beneficiario
     Printer.FontSize = SetD(3).Tamaño
     PrinterTexto SetD(3).PosX, PosLinea1 + SetD(3).PosY, Format$(Total, "#,###.00")
     If SetD(4).PosX > 0 And SetD(4).PosY > 0 Then
        Printer.FontSize = SetD(4).Tamaño
        PrinterNumCheque SetD(4).PosX, PosLinea1 + SetD(4).PosY, SetD(5).PosX, Total
     End If
     If SetD(9).PosX > 0 And SetD(9).PosY > 0 Then
        Cadena = Empresa & " " & Moneda & " " & Format$(Total, "#,##0.00") & "**"
        Printer.FontSize = SetD(9).Tamaño
        PrinterTexto SetD(9).PosX, PosLinea1 + SetD(9).PosY, Cadena
     End If
     Printer.FontSize = SetD(10).Tamaño
     PrinterTexto SetD(10).PosX, PosLinea1 + SetD(10).PosY, ULCase(NombreCiudad)
     
     Printer.FontSize = SetD(6).Tamaño
     PrinterTexto SetD(6).PosX, PosLinea1 + SetD(6).PosY, Format$(FechaTexto, "yyyy/MM/dd")
     Printer.FontBold = False
End Sub

Public Sub Imprimir_Bloque_Cheques(Desde As String, _
                                   Hasta As String, _
                                   Cta_Banco As String, _
                                   Cheque_No As String, _
                                   Detalle_Cheque As String)
Dim TipoBank As String
Dim AdoComp As ADODB.Recordset
Dim AdoCliente As ADODB.Recordset
On Error GoTo Errorhandler
RatonReloj
  sSQL = "SELECT Codigo,Cliente,CI_RUC " _
       & "FROM Clientes " _
       & "WHERE Codigo <> '.' " _
       & "ORDER BY Cliente "
  Select_AdoDB AdoCliente, sSQL
       
Cuenta = SinEspaciosIzq(Cta_Banco)  ' Cta
Codigo1 = SinEspaciosIzq(Cta_Banco)   ' Cta
Codigo2 = SinEspaciosIzq(Cheque_No)   ' Cheque No
If Desde = Hasta Then
   Mensajes = "Imprimir Cheque No. " & Desde & vbCrLf _
            & UCaseStrg(Detalle_Cheque) & vbCrLf _
            & "Cuenta No. " & Cuenta
Else
   Mensajes = "Imprimir Cheques Desde " & Desde & " al " & Hasta & vbCrLf _
            & UCaseStrg(Detalle_Cheque) & vbCrLf _
            & "Cuenta No. " & Cuenta
End If

  sSQL = "SELECT C.T,C.TP,C.Numero,Cl.Cliente,C.Codigo_B,C.Fecha,C1.Cuenta,Cl.CI_RUC," _
       & "Cl.Direccion,T.Cta,T.Cheq_Dep,T.Haber,T.Codigo_C,T.Fecha_Efec " _
       & "FROM Transacciones As T,Comprobantes As C,Clientes As Cl,Catalogo_Cuentas As C1 " _
       & "WHERE IsNumeric(T.Cheq_Dep) <> " & adFalse & " " _
       & "AND T.Cheq_Dep BETWEEN '" & Desde & "' and '" & Hasta & "' " _
       & "AND T.Item = '" & NumEmpresa & "' " _
       & "AND T.Periodo = '" & Periodo_Contable & "' " _
       & "AND T.Cta = '" & Cuenta & "' " _
       & "AND T.Haber > 0 " _
       & "AND T.Item = C.Item " _
       & "AND T.Item = C1.Item " _
       & "AND T.Periodo = C.Periodo " _
       & "AND T.Periodo = C1.Periodo " _
       & "AND T.TP = C.TP " _
       & "AND T.Numero = C.Numero " _
       & "AND C.Codigo_B = Cl.Codigo " _
       & "AND C1.Codigo = T.Cta " _
       & "ORDER BY T.Cheq_Dep "
  Select_AdoDB AdoComp, sSQL
  
  
Titulo = "IMPRESION"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
   Escala_Centimetro Orientacion_Pagina, TipoTimes, 10
   TipoBank = Format$(Val(MidStrg(Codigo1, Len(Codigo1) - 1, 2)), "00")
  'MsgBox TipoBank
   CCHQConLineas = ProcesarSeteos(TipoBank)
   InicioX = 0.5: InicioY = 0
   'DataAnchoCampos InicioX, AdoComp, SizeLetra, TipoTimes, 1
   Pagina = 1
   PosLinea = 0
  'Iniciamos la impresion
   Printer.FontBold = True
   With AdoComp
    If .RecordCount > 0 Then
        If Desde = Hasta Then
           FechaTexto = .fields("Fecha")
           If CFechaLong(.fields("Fecha")) < CFechaLong(.fields("Fecha_Efec")) Then
              FechaTexto = .fields("Fecha_Efec")
           End If
           Beneficiario = .fields("Cliente")
           CodigoCli = .fields("Codigo_C")
           If Beneficiario = Ninguno Then
              If AdoCliente.RecordCount > 0 Then
                 AdoCliente.MoveFirst
                 AdoCliente.Find ("Codigo = '" & CodigoCli & "' ")
                 If Not AdoCliente.EOF Then Beneficiario = AdoCliente.fields("Cliente")
              End If
           End If
           Valor = .fields("Haber")
           NumCheque = .fields("Cheq_Dep")
           Cuenta = .fields("Cta")
           Imprimir_Cheque_No PosLinea
        Else
          .MoveFirst
          .Find ("Cta = '" & Codigo1 & "' ")
           If Not .EOF Then
              Do While Not .EOF
                 If (Desde <= .fields("Cheq_Dep")) And (.fields("Cheq_Dep") <= Hasta) Then
                    FechaTexto = .fields("Fecha")
                    If CFechaLong(.fields("Fecha")) < CFechaLong(.fields("Fecha_Efec")) Then
                       FechaTexto = .fields("Fecha_Efec")
                    End If
                    Beneficiario = .fields("Cliente")
                    CodigoCli = .fields("Codigo_C")
                    If Beneficiario = Ninguno Then
                       If AdoCliente.RecordCount > 0 Then
                          AdoCliente.MoveFirst
                          AdoCliente.Find ("Codigo = '" & CodigoCli & "' ")
                          If Not AdoCliente.EOF Then Beneficiario = AdoCliente.fields("Cliente")
                       End If
                    End If
                    Valor = .fields("Haber")
                    NumCheque = .fields("Cheq_Dep")
                    Cuenta = .fields("Cta")
                    Imprimir_Cheque_No PosLinea
                    PosLinea = PosLinea + 7.5
                    If PosLinea >= LimiteAlto Then
                       Printer.NewPage
                       PosLinea = 0
                    End If
                 End If
                .MoveNext
              Loop
           End If
        End If
    End If
   End With
RatonNormal
MensajeEncabData = ""
Printer.EndDoc
AdoComp.Close
AdoCliente.Close
Exit Sub
Errorhandler:
    RatonNormal
    ErrorDeImpresion
    Exit Sub
Else
    RatonNormal
End If
End Sub

Public Sub Imprimir_Saldos_SubCtas_Costos(Datas As Adodc, Optional PorCta As CheckBox)
SizeLetra = 6
Mensajes = "Seguro de Imprimir en:" & vbCrLf & Printer.DeviceName & "?"
Titulo = "IMPRESION"
Bandera = False
Orientacion_Pagina = 1
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
RatonReloj
InicioX = 1: InicioY = 0
SubTotal = 0: Total = 0
DataAnchoCampos InicioX, Datas, 6, TipoVerdana, Orientacion_Pagina
If PorCta.value = 1 Then
   Ancho(0) = 1     'Codigo Cta
   Ancho(1) = 3     'Cuenta
Else
   Ancho(0) = 1     'Codigo Cta
   Ancho(1) = 7     'Cuenta
End If
Ancho(2) = 9     'SubCuenta
Ancho(3) = 17    'Total
Ancho(4) = 19    'Fin
Pagina = 1
'Iniciamos la impresion
Printer.FontBold = False
With Datas.Recordset
 If .RecordCount > 0 Then
     .MoveFirst
      Encabezado_SubCta_Costos PorCta
      If PorCta.value = 1 Then
         Cuenta = .fields("Cta") & " - " & .fields("Cuenta")
      Else
         Cuenta = .fields("Sub_Modulos")
      End If
      Si_No = True
      Printer.FontName = TipoVerdana
      Printer.FontItalic = False
      Printer.FontBold = False
      Printer.FontSize = SizeLetra
      PrinterTexto Ancho(0), PosLinea, UCaseStrg(Cuenta)
      Do While Not .EOF
         If PorCta.value = 1 Then
            If Cuenta <> .fields("Cta") & " - " & .fields("Cuenta") Then
               PosLinea = PosLinea + 0.05
               Printer.Line (Ancho(0), PosLinea)-(Ancho(CantCampos), PosLinea), Negro
               PosLinea = PosLinea + 0.05
               PrinterVariables Ancho(3) - 2, PosLinea, "SUBTOTAL"
               PrinterVariables Ancho(3), PosLinea, SubTotal
               PosLinea = PosLinea + 0.5
               Printer.Line (Ancho(0), PosLinea)-(Ancho(CantCampos), PosLinea), Negro
               PosLinea = PosLinea + 0.1
               SubTotal = 0
               Cuenta = .fields("Cta") & " - " & .fields("Cuenta")
               Si_No = True
               PrinterTexto Ancho(0), PosLinea, UCaseStrg(Cuenta)
            End If
            PrinterFields Ancho(2), PosLinea, .fields(2)
            PrinterFields Ancho(3), PosLinea, .fields(3)
         Else
            If Cuenta <> .fields("Sub_Modulos") Then
               PosLinea = PosLinea + 0.05
               Printer.Line (Ancho(0), PosLinea)-(Ancho(CantCampos), PosLinea), Negro
               PosLinea = PosLinea + 0.05
               PrinterVariables Ancho(3) - 2, PosLinea, "SUBTOTAL"
               PrinterVariables Ancho(3), PosLinea, SubTotal
               PosLinea = PosLinea + 0.5
               Printer.Line (Ancho(0), PosLinea)-(Ancho(CantCampos), PosLinea), Negro
               PosLinea = PosLinea + 0.1
               SubTotal = 0
               Cuenta = .fields("Sub_Modulos")
               Si_No = True
               PrinterTexto Ancho(0), PosLinea, UCaseStrg(Cuenta)
            End If
            PrinterFields Ancho(1), PosLinea, .fields(1)
            PrinterFields Ancho(2), PosLinea, .fields(2)
            PrinterFields Ancho(3), PosLinea, .fields(3)
         End If
         For I = 0 To CantCampos
             If PorCta.value = 1 Then
                If I <> 1 Then Printer.Line (Ancho(I), PosLinea - 0.1)-(Ancho(I), PosLinea + 0.3), Negro
             Else
                Printer.Line (Ancho(I), PosLinea - 0.1)-(Ancho(I), PosLinea + 0.3), Negro
             End If
         Next I
         PosLinea = PosLinea + 0.3
         If PosLinea >= LimiteAlto Then
            Printer.Line (Ancho(0), PosLinea)-(Ancho(CantCampos), PosLinea), Negro
            Printer.NewPage
            Encabezado_SubCta_Costos PorCta
            Printer.FontName = TipoVerdana
            Printer.FontItalic = False
            Printer.FontBold = False
            Printer.FontSize = SizeLetra
            PrinterTexto Ancho(0), PosLinea, Cuenta
         End If
        'Total
         Total = Total + .fields("Total")
         SubTotal = SubTotal + .fields("Total")
        .MoveNext
      Loop
     .MoveFirst
 End If
End With
   PosLinea = PosLinea + 0.05
   Printer.Line (Ancho(0), PosLinea)-(Ancho(CantCampos), PosLinea), Negro
   PosLinea = PosLinea + 0.05
   PrinterVariables Ancho(3) - 2, PosLinea, "SUBTOTAL"
   PrinterVariables Ancho(3), PosLinea, SubTotal
   
   PosLinea = PosLinea + 0.5
   Printer.Line (Ancho(0), PosLinea)-(Ancho(CantCampos), PosLinea), Negro
   PosLinea = PosLinea + 0.05

Printer.Line (Ancho(0), PosLinea)-(Ancho(CantCampos), PosLinea), Negro
PosLinea = PosLinea + 0.05
   PrinterVariables Ancho(3) - 2, PosLinea, "T O T A L"
   PrinterVariables Ancho(3), PosLinea, Total

PosLinea = PosLinea + 0.4
Printer.Line (Ancho(0), PosLinea)-(Ancho(CantCampos), PosLinea), Negro
PosLinea = PosLinea + 0.05
Printer.Line (Ancho(0), PosLinea)-(Ancho(CantCampos), PosLinea), Negro
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

Public Sub Encabezado_SubCta_Costos(Optional PorCta As CheckBox)
Dim InicX As Single
Dim InicY As Single
Encabezado Ancho(0), Ancho(CantCampos)
PorteLetra = Printer.FontSize
LetraAnterior = Printer.FontName
Printer.FontName = TipoVerdana
Printer.FontBold = True
Printer.FontItalic = False
If SQLMsg1 <> "" Then
   Printer.FontSize = 10
   PrinterTexto CentrarTexto(SQLMsg1), PosLinea, SQLMsg1
   PosLinea = PosLinea + 0.5
End If
Printer.FontSize = 8
If SQLMsg2 <> "" And SQLMsg3 <> "" Then
   Cadena = TrimStrg(SQLMsg3 & " " & SQLMsg2)
   PrinterTexto CentrarTexto(Cadena), PosLinea, Cadena
   PosLinea = PosLinea + 0.5
End If
Printer.FontSize = 7
Printer.FontUnderline = True
'========================================================================
If PorCta.value = 1 Then
   PrinterTexto Ancho(0), PosLinea, "C U E N T A"
   PrinterTexto Ancho(2), PosLinea, "S U B C U E N T A"
   PrinterTexto Ancho(3), PosLinea, "SUBTOTAL"
Else
   PrinterTexto Ancho(0), PosLinea, "S U B C U E N T A"
   PrinterTexto Ancho(1), PosLinea, "C U E N T A"
   PrinterTexto Ancho(2), PosLinea, "D E T A L L E"
   PrinterTexto Ancho(3), PosLinea, "SUBTOTAL"
End If
PosLinea = PosLinea + 0.4
Printer.FontUnderline = False
Printer.FontBold = False
Printer.FontSize = PorteLetra
Printer.FontName = LetraAnterior
End Sub

Public Function SQL_Tipo_Balance(TipoBalance As String, _
                                 TipoPyGCC As String) As String
Dim sSQL As String
Dim AdoCCDB As ADODB.Recordset
   
   Progreso_Barra.Valor_Maximo = 100
   Progreso_Barra.Mensaje_Box = "Iniciando la mayorizacion. Espere un momento..."
   Progreso_Iniciar
   CantCtas = 0
   Select Case TipoBalance
     Case "BR"
          If TipoPyGCC = "01" Or TipoPyGCC = "04" Then
            sSQL = "SELECT Codigo " _
                 & "FROM Catalogo_SubCtas " _
                 & "WHERE Item = '" & NumEmpresa & "' " _
                 & "AND Periodo = '" & Periodo_Contable & "' " _
                 & "AND Nivel = '" & TipoPyGCC & "' " _
                 & "AND Agrupacion = 0 " _
                 & "AND TC = 'CC' " _
                 & "ORDER BY Codigo "
            Select_AdoDB AdoCCDB, sSQL
            If AdoCCDB.RecordCount > 0 Then
               sSQL = "SELECT Item, Codigo, Cuenta, "
               Do While Not AdoCCDB.EOF
                  sSQL = sSQL & "Total_" & AdoCCDB.fields("Codigo") & ", "
                  AdoCCDB.MoveNext
               Loop
               sSQL = sSQL & "Saldo_Total, DG "
            End If
            AdoCCDB.Close
          Else
             sSQL = "SELECT Item, Codigo, Cuenta, Saldo_Anterior, Debitos, Creditos, Saldo_Mes, Saldo_Total, DG "
          End If
     Case Else
          sSQL = "SELECT Item, Codigo, Cuenta, Saldo_Anterior, Debitos, Creditos, Saldo_Mes, Saldo_Total, DG "
    End Select
    sSQL = sSQL _
         & "FROM Catalogo_Cuentas_Exterior " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND ABS(Saldo_Anterior) + ABS(Debitos) + ABS(Creditos) + ABS(Saldo_Mes) + ABS(Saldo_Total) <> 0 "
    Select Case TipoBalance
      Case "BS": sSQL = sSQL & "AND SUBSTRING(Codigo,1,1) BETWEEN '1' and '3' "
      Case "BR": sSQL = sSQL & "AND SUBSTRING(Codigo,1,1) BETWEEN '4' and '9' "
    End Select
    sSQL = sSQL & "ORDER BY Codigo "
    
    sSQLTotales = "SELECT DG, SUM(Debitos) As TDebitos, SUM(Creditos) As TCreditos " _
                & "FROM Catalogo_Cuentas_Exterior " _
                & "WHERE Item = '" & NumEmpresa & "' " _
                & "AND Periodo = '" & Periodo_Contable & "' " _
                & "AND (Debitos + Creditos) <> 0 " _
                & "AND DG = 'D' "
    Select Case TipoBalance
      Case "BS": sSQLTotales = sSQLTotales & "AND SUBSTRING(Codigo,1,1) BETWEEN '1' and '3' "
      Case "BR": sSQLTotales = sSQLTotales & "AND SUBSTRING(Codigo,1,1) BETWEEN '4' and '9' "
    End Select
    sSQLTotales = sSQLTotales & "GROUP BY DG "
    SQL_Tipo_Balance = sSQL
End Function
