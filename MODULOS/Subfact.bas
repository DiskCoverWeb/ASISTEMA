Attribute VB_Name = "SubFacturas"
Option Explicit

Public Sub Enviar_Recibo_Abonos_x_Email(TFA As Tipo_Facturas, FechaAbonos As String)
Dim AdoCliDB As ADODB.Recordset
   'Enviar por mail el Abono receptado
''    sSQL = "SELECT Cliente,CI_RUC,TD,Email,EmailR,Direccion,DireccionT,Ciudad,Telefono,Telefono_R,Grupo,Representante,CI_RUC_R,TD_R " _
''         & "FROM Clientes " _
''         & "WHERE Codigo = '" & .CodigoC & "' "
''    Select_AdoDB AdoCliDB, sSQL
   
    SRI_Autorizacion.Autorizacion = TA.Autorizacion
    SRI_Autorizacion.Tipo_Doc_SRI = "AB"
         FA.Nota = "Abono Recibido en la Institucion:" & vbCrLf _
                 & "Hora" & vbTab & vbTab & ": " & Format(Time, "HH:MM:SS") & vbCrLf _
                 & "Documento" & vbTab & ": " & TA.Recibo_No & vbCrLf _
                 & "Valor Recibdo USD " & Format(TotalAbonos, "#,##0.00") & vbCrLf
         SRI_Enviar_Mails FA, SRI_Autorizacion

End Sub

'Facturas            : RUC_CI, TB, Razon_Social
'Clientes_Matriculas : Cedula_R, TD, Representante
Public Function Leer_Datos_Cliente_FA(TFA As Tipo_Facturas) As Tipo_Facturas
'    If Len(Codigo_CIRUC_Cliente) <= 0 Then Codigo_CIRUC_Cliente = Ninguno
    With TFA
     If .CodigoC = Ninguno Then
        .Cliente = Ninguno
        .CI_RUC = "000000000"
        .TD = "O"
        .EmailC = Ninguno
        .EmailR = Ninguno
        .TelefonoC = "099000000"
        .DireccionC = "SD"
        .Grupo = "999999"
        .Curso = Ninguno
        .CiudadC = Ninguno
        .Razon_Social = Ninguno
        .RUC_CI = "000000000"
        .TB = "O"
      Else
        'Verificamos la informacion del Clienete
         TBeneficiario = Leer_Datos_Cliente_SP(TFA.CodigoC)
        .Cliente = TBeneficiario.Cliente
        .Razon_Social = TBeneficiario.Representante
        .TD = TBeneficiario.TD
        .TB = TBeneficiario.TD_Rep
        .TD_R = TBeneficiario.TD_Rep
        .RUC_CI = TBeneficiario.RUC_CI_Rep
        .CI_RUC = TBeneficiario.CI_RUC
        .TelefonoC = TBeneficiario.Telefono1
         If Len(TBeneficiario.Direccion_Rep) > 1 Then
           .DireccionC = TBeneficiario.Direccion_Rep
         Else
           .DireccionC = TBeneficiario.Direccion
         End If
        .Curso = TBeneficiario.Direccion
        .Grupo = TBeneficiario.Grupo_No
        .EmailC = TBeneficiario.Email1
        .EmailR = TBeneficiario.EmailR
        
        .Tipo_Cta = TBeneficiario.Tipo_Cta
        .Cod_Banco = TBeneficiario.Cod_Banco
        .Cta_Numero = TBeneficiario.Cta_Numero
        .Fecha_Cad = Format(Month(TBeneficiario.Fecha_Cad), "00") & "/" & Format(Year(TBeneficiario.Fecha_Cad), "0000")
        .Por_Deposito = TBeneficiario.Por_Deposito
     End If
    End With
   'MsgBox ".... " & TFA.DireccionC
    Leer_Datos_Cliente_FA = TFA
End Function

'''Public Sub Generar_XML_Facturas(TFA As Tipo_Facturas)
'''Dim AdoDBFA As ADODB.Recordset
'''Dim AdoDBDet As ADODB.Recordset
'''Dim AdoDBProd As ADODB.Recordset
'''Dim NumFile As Integer
'''Dim RutaGeneraFile As String
'''
'''Dim TipoIdent As String
'''Dim Cod_Aux As String
'''Dim Cod_Bar As String
'''Dim TotalDescuento As Currency
'''Dim TotalDescuento_0 As Currency
'''Dim TotalDescuento_X As Currency
'''Dim Autorizar_XML As Boolean
'''
'''    RatonReloj
'''    Autorizar_XML = False
'''    TextoXML = ""
'''   'Averiguamos si la Factura esta a nombre del Representante
'''    TBeneficiario = Leer_Datos_Cliente_SP(TFA.CodigoC)
'''   'MsgBox TBeneficiario.RUC_CI_Rep & vbCrLf & TBeneficiario.Representante & vbCrLf & TBeneficiario.TD_Rep
'''
'''    TFA.Cliente = TBeneficiario.Representante
'''    TFA.TD = TBeneficiario.TD_Rep
'''    TFA.CI_RUC = TBeneficiario.RUC_CI_Rep
'''    TFA.TelefonoC = TBeneficiario.Telefono1
'''    TFA.DireccionC = TBeneficiario.Direccion_Rep
'''    TFA.Curso = TBeneficiario.Direccion
'''    TFA.Grupo = TBeneficiario.Grupo_No
'''    TFA.EmailC = TBeneficiario.Email1
'''    TFA.EmailR = TBeneficiario.Email2
'''    TFA.Contacto = TBeneficiario.Contacto
'''
'''   'Detalle de descuentos
'''    TotalDescuento_0 = 0
'''    TotalDescuento_X = 0
'''    sSQL = "SELECT DF.*,CP.Reg_Sanitario,CP.Marca " _
'''         & "FROM Detalle_Factura As DF, Catalogo_Productos As CP " _
'''         & "WHERE DF.Item = '" & NumEmpresa & "' " _
'''         & "AND DF.Periodo = '" & Periodo_Contable & "' " _
'''         & "AND DF.TC = '" & TFA.TC & "' " _
'''         & "AND DF.Serie = '" & TFA.Serie & "' " _
'''         & "AND DF.Autorizacion = '" & TFA.Autorizacion & "' " _
'''         & "AND DF.Factura = " & TFA.Factura & " " _
'''         & "AND LEN(DF.Autorizacion) >= 13 " _
'''         & "AND DF.T <> 'A' " _
'''         & "AND DF.Item = CP.Item " _
'''         & "AND DF.Periodo = CP.Periodo " _
'''         & "AND DF.Codigo = CP.Codigo_Inv " _
'''         & "ORDER BY DF.ID,DF.Codigo "
'''    Select_AdoDB AdoDBDet, sSQL
'''    RatonReloj
'''    With AdoDBDet
'''     If .RecordCount > 0 Then
'''         Do While Not .EOF
'''            If .Fields("Total_IVA") = 0 Then
'''                TotalDescuento_0 = TotalDescuento_0 + .Fields("Total_Desc") + .Fields("Total_Desc2")
'''            Else
'''                TotalDescuento_X = TotalDescuento_X + .Fields("Total_Desc") + .Fields("Total_Desc2")
'''            End If
'''           .MoveNext
'''         Loop
'''        .MoveFirst
'''     End If
'''    End With
'''    TotalDescuento_0 = Redondear(TotalDescuento_0, 2)
'''    TotalDescuento_X = Redondear(TotalDescuento_X, 2)
'''    TotalDescuento = TotalDescuento_0 + TotalDescuento_X
'''
'''   'Encabezado de la Factura
'''    sSQL = "SELECT * " _
'''         & "FROM Facturas " _
'''         & "WHERE Item = '" & NumEmpresa & "' " _
'''         & "AND Periodo = '" & Periodo_Contable & "' " _
'''         & "AND TC = '" & TFA.TC & "' " _
'''         & "AND Serie = '" & TFA.Serie & "' " _
'''         & "AND Factura = " & TFA.Factura & " " _
'''         & "AND Autorizacion = '" & TFA.Autorizacion & "' " _
'''         & "AND T <> 'A' "
'''    Select_AdoDB AdoDBFA, sSQL
'''    RatonReloj
'''    With AdoDBFA
'''     If .RecordCount > 0 Then
'''         Autorizar_XML = True
'''        'TFA.CodigoA = .Fields("")
'''         TFA.T = .Fields("T")
'''         TFA.SP = .Fields("SP")
'''         TFA.Porc_IVA = .Fields("Porc_IVA")
'''         TFA.Imp_Mes = .Fields("Imp_Mes")
'''         TFA.Fecha = .Fields("Fecha")
'''         TFA.Vencimiento = .Fields("Vencimiento")
'''         TFA.SubTotal = .Fields("SubTotal")
'''         TFA.Sin_IVA = .Fields("Sin_IVA")
'''         TFA.Con_IVA = .Fields("Con_IVA")
'''         TFA.Descuento = .Fields("Descuento")
'''         TFA.Descuento2 = .Fields("Descuento2")
'''         TFA.Total_IVA = .Fields("IVA")
'''         TFA.Total_MN = .Fields("Total_MN")
'''         TFA.Razon_Social = .Fields("Razon_Social")
'''         TFA.RUC_CI = .Fields("RUC_CI")
'''         TFA.TB = .Fields("TB")
'''
'''        'MsgBox "Validar Porc IVA"
'''
'''         Validar_Porc_IVA TFA.Fecha
'''        'Generamos la Clave de acceso
'''         TFA.ClaveAcceso = Format$(TFA.Fecha, "ddmmyyyy") & "50850"
'''         Select Case TFA.TC
'''           Case "FA": TFA.ClaveAcceso = TFA.ClaveAcceso & "01"
'''           Case "NV": TFA.ClaveAcceso = TFA.ClaveAcceso & "02"
'''           Case "NC": TFA.ClaveAcceso = TFA.ClaveAcceso & "04"
'''         End Select
'''         TFA.ClaveAcceso = TFA.ClaveAcceso & TFA.Serie & Format$(TFA.Factura, String(9, "0")) & RUC & "003"
'''         TFA.ClaveAcceso = TFA.ClaveAcceso & Digito_Verificador_Modulo11(TFA.ClaveAcceso)
'''
'''         TFA.Hora = Format$(Time, FormatoTimes)
'''         TipoIdent = "P"
'''         Select Case TFA.TB
'''           Case "R": If TFA.CI_RUC = String(13, "9") Then TipoIdent = "07" Else TipoIdent = "04"
'''           Case "C": TipoIdent = "05"
'''           Case "P": TipoIdent = "06"
'''           Case Else: TipoIdent = "07"
'''         End Select
'''
'''        'ENCABEZADO XML PARA EL SRI DE LA FACTURA/NOTA DE VENTA
'''        'standalone=""yes""
'''         Insertar_Campo_XML "<?xml version=""1.0"" encoding=""UTF-8""?>"
'''         Select Case TFA.TC
'''           Case "FA": Insertar_Campo_XML "<factura id=""comprobante"" version=""1.1.0"">"
'''           Case "NV": Insertar_Campo_XML AbrirXML("<notaVenta>")
'''           Case Else: Insertar_Campo_XML AbrirXML("<puntoVenta>")
'''         End Select
'''        'Encabezado de la Factura
'''         Insertar_Campo_XML AbrirXML("infoFactura")
'''            Insertar_Campo_XML CampoXML("claveAcceso", TFA.ClaveAcceso)
'''            Insertar_Campo_XML CampoXML("codtda", "50850")
'''            Insertar_Campo_XML CampoXML("numint", "00129219")
'''            Insertar_Campo_XML CampoXML("idnota", "00129219")
'''            Select Case TFA.TC
'''              Case "FA": Insertar_Campo_XML CampoXML("tipdoc", "01")
'''              Case "NV": Insertar_Campo_XML CampoXML("tipdoc", "02")
'''              Case "NC": Insertar_Campo_XML CampoXML("tipdoc", "07")
'''              Case Else: Insertar_Campo_XML CampoXML("tipdoc", "00")
'''            End Select
'''            Insertar_Campo_XML CampoXML("serie", TFA.Serie)
'''            Insertar_Campo_XML CampoXML("numero", Format$(TFA.Factura, String(9, "0")))
'''            Insertar_Campo_XML CampoXML("fecemi", TFA.Fecha)
'''            Insertar_Campo_XML CampoXML("nomcli", TFA.Razon_Social)
'''            Insertar_Campo_XML CampoXML("dircli", TFA.DireccionC)
'''            Insertar_Campo_XML CampoXML("ruccli", TFA.RUC_CI)
'''            Insertar_Campo_XML CampoXML("telefcli", TFA.TelefonoC)
'''            Insertar_Campo_XML CampoXML("email", "comprobantes@diskcoversystem.com")
'''            Insertar_Campo_XML CampoXML("moneda", "02")
'''            Insertar_Campo_XML CampoXML("prevta", TFA.SubTotal)
'''            Insertar_Campo_XML CampoXML("dscto", TFA.Descuento + TFA.Descuento2)
'''            Insertar_Campo_XML CampoXML("valbrut", TFA.RUC_CI)
'''            Insertar_Campo_XML CampoXML("valven", TFA.SubTotal - TFA.Descuento - TFA.Descuento2)
'''            Insertar_Campo_XML CampoXML("igv", TFA.Total_IVA)
'''            Insertar_Campo_XML CampoXML("Total", TFA.Total_MN)
'''         Insertar_Campo_XML CerrarXML("infoFactura")
'''     End If
'''    End With
'''
'''   'Detalle de la Factura/Nota de Venta
'''    RatonReloj
'''    With AdoDBDet
'''     If .RecordCount > 0 Then
'''         Insertar_Campo_XML AbrirXML("detalles")
'''         Do While Not .EOF
'''            SubTotal = (.Fields("Cantidad") * .Fields("Precio")) - (.Fields("Total_Desc") + .Fields("Total_Desc2"))
'''            Insertar_Campo_XML AbrirXML("detalle")
'''                Insertar_Campo_XML CampoXML("codtda", "50850")
'''                Insertar_Campo_XML CampoXML("numint", "00129219")
'''                Insertar_Campo_XML CampoXML("codart", "5515322")
'''                Insertar_Campo_XML CampoXML("calidad", "1")
'''                Insertar_Campo_XML CampoXML("tallaid", "40")
'''                Insertar_Campo_XML CampoXML("detallep", "Colocar el nombre del producto que se vendio")
'''                Insertar_Campo_XML CampoXML("canti", .Fields("Cantidad"))
'''                Insertar_Campo_XML CampoXML("prevta", .Fields("Precio"))
'''                Insertar_Campo_XML CampoXML("dscto", .Fields("Total_Desc") + .Fields("Total_Desc2"))
'''                Insertar_Campo_XML CampoXML("valbrut", .Fields("Total"))
'''                Insertar_Campo_XML CampoXML("valven", .Fields("Precio"))
'''                Insertar_Campo_XML CampoXML("igv", .Fields("Total_IVA"))
'''                Insertar_Campo_XML CampoXML("total", .Fields("Total"))
'''            Insertar_Campo_XML CerrarXML("detalle")
'''           .MoveNext
'''         Loop
'''         Insertar_Campo_XML CerrarXML("detalles")
'''     End If
'''    End With
'''
'''       Insertar_Campo_XML AbrirXML("infoAbonos")
'''            Insertar_Campo_XML CampoXML("codtda", "50850")
'''            Insertar_Campo_XML CampoXML("numint", "00129219")
'''            Insertar_Campo_XML CampoXML("moneda", "03")
'''            Insertar_Campo_XML CampoXML("tippago", "02")
'''            Insertar_Campo_XML CampoXML("tasa", "0.00")
'''            Insertar_Campo_XML CampoXML("numrefe", "5554545555555555")
'''            Insertar_Campo_XML CampoXML("montopag", TFA.Total_MN)
'''            Insertar_Campo_XML CampoXML("logcrea", "20190601 14:55:33 VEN")
'''       Insertar_Campo_XML CerrarXML("infoAbonos")
'''
'''      'Fin del Archivo Xml
'''       Select Case TFA.TC
'''         Case "FA": Insertar_Campo_XML CerrarXML("factura")
'''         Case "NV": Insertar_Campo_XML CerrarXML("notaVenta")
'''         Case Else: Insertar_Campo_XML CerrarXML("puntoVenta")
'''       End Select
'''
'''      'Grabamos el comprobante XML
'''       RutaGeneraFile = RutaSysBases & "\TEMP\" & TFA.ClaveAcceso & ".xml"
'''       NumFile = FreeFile
'''       Open RutaGeneraFile For Output As #NumFile ' Abre el archivo.
'''            Print #NumFile, TextoXML;
'''       Close #NumFile
'''       RatonReloj
'''  ' Cerramos tablas abiertas
'''    AdoDBDet.Close
'''    AdoDBFA.Close
'''    RatonNormal
'''End Sub

Public Sub Encerar_Factura(TFA As Tipo_Facturas)
    With TFA
        .SP = False
        .Si_Existe_Doc = False
        .C = False
        .p = False
        .ME_ = False
        .Com_Pag = False
        .T = Normal
        .CodigoC = Ninguno
        .CodigoB = Ninguno
        .CodigoU = Ninguno
        .Codigo_T = Ninguno
        .Cliente = Ninguno
        .Contacto = Ninguno
        .CI_RUC = Ninguno
        .Razon_Social = Ninguno
        .RUC_CI = Ninguno
        .TB = Ninguno
        .DirNumero = Ninguno
        .DireccionC = Ninguno
        .DireccionS = Ninguno
        .CiudadC = Ninguno
        .EmailR = Ninguno
        .Curso = Ninguno
        .Grupo = Ninguno
        .Cod_Ejec = Ninguno
        .Forma_Pago = Ninguno
        .Imp_Mes = False
        .Fecha = FechaSistema
        .Fecha_V = FechaSistema
        .Fecha_C = FechaSistema
        .Fecha_Aut = FechaSistema
        .Vencimiento = FechaSistema
        .Hora = "00:00:00"
        .Tipo_Pago = 0
        .Tipo_Pago_Det = Ninguno
        .Cod_CxC = Ninguno
        .Nivel = Ninguno
        .Nota = Ninguno
        .Observacion = Ninguno
        .Definitivo = Ninguno
        .Declaracion = Ninguno
        .SubCta = Ninguno
        
        'SubTotales de la Factura
        .Factura = 0
        '.Porc_IVA = 0
        .TDT = 18
        .Total_Sin_No_IVA = 0
        .Descuento = 0
        .Descuento2 = 0
        .Total_Descuento = 0
        .SubTotal = 0
        .Total_IVA = 0
        .Con_IVA = 0
        .Sin_IVA = 0
        .Total_MN = 0
        .Saldo_MN = 0
        .Comision = 0
        .Servicio = 0
        .Total_ME = 0
        .Saldo_ME = 0
        .Cantidad = 0
        .Kilos = 0
        .Saldo_Actual = 0
        .Efectivo = 0
        .Saldo_Pend = 0
        .Saldo_Pend_MN = 0
        .Saldo_Pend_ME = 0
        .Ret_Fuente = 0
        .Ret_IVA = 0
        .Porc_C = 0
        .Cotizacion = 0
        .DAU = 0
        .FUE = 0
        .Solicitud = 0
        .Retencion = 0
        'Guia de Remision
        .ClaveAcceso_GR = Ninguno
        .Autorizacion_GR = Ninguno
        .Serie_GR = Ninguno
        .Remision = 0
        .Comercial = Ninguno
        .CIRUCComercial = Ninguno
        .Entrega = Ninguno
        .CIRUCEntrega = Ninguno
        .CiudadGRI = Ninguno
        .CiudadGRF = Ninguno
        .Placa_Vehiculo = Ninguno
        .FechaGRE = FechaSistema
        .FechaGRI = FechaSistema
        .FechaGRF = FechaSistema
        .Pedido = Ninguno
        .Zona = Ninguno
        .Orden_Compra = "0"
        .Serie_GR = Ninguno
        .Digitador = Ninguno
        .Error_SRI = Ninguno
        .Estado_SRI = Ninguno
        .Dir_PartidaGR = Ninguno
        .Dir_EntregaGR = Ninguno
        
        'CxC Linea de Produccion
        .CxC_Clientes = Ninguno
        .Cta_Venta = Ninguno
        .LogoFactura = Ninguno
        .LogoNotaCredito = Ninguno
        .AltoFactura = 0
        .AnchoFactura = 0
        .EspacioFactura = 0
        .Pos_Factura = 0
        .DireccionEstab = Ninguno
        .NombreEstab = Ninguno
        .TelefonoEstab = Ninguno
        .LogoTipoEstab = Ninguno
        .Autorizacion_R = Ninguno
        .Serie_R = "001001"
        .Fecha_Tours = Ninguno
    End With
End Sub

Public Sub Imprimir_CxC_Grupos(Datas As Adodc, _
                              SizeLetra As Integer, _
                              Optional EsCampoCorto As Boolean)
Dim PInicio As Single
Dim PFinal As Single
Dim GrupoNo As String

On Error GoTo Errorhandler

Mensajes = "Seguro de Imprimir en:" & vbCrLf & Printer.DeviceName & "?"
Titulo = "IMPRESION"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
RatonReloj
MensajeEncabData = "RESUMEN DE PENSIONES POR MESES"
InicioX = 0.5: InicioY = 0
'TipoTimes
DataAnchoCampos InicioX, Datas, SizeLetra, TipoCondensed, Orientacion_Pagina, EsCampoCorto
'MsgBox CantCampos
''Ancho(1) = 1
''Ancho(2) = 4
'Ancho(CantCampos - 2) = 0
Pagina = 1
Total = 0
'Iniciamos la impresion
Printer.FontBold = False
With Datas.Recordset
If .RecordCount > 0 Then
    .MoveFirst
     GrupoNo = .fields("GrupoNo")
     SQLMsg2 = .fields("Detalle_Grupo") & " - " & .fields("GrupoNo")
     'SQLMsg3 = "TOTAL POR COBRAR USD " & Format$(Total, "#,##0.00")
     EncabezadoData Datas
     PInicio = PosLinea
     Printer.FontBold = False
     Printer.FontSize = SizeLetra
     'MsgBox CantCampos
     Do While Not .EOF
        If GrupoNo <> .fields("GrupoNo") Then
           For I = 0 To CantCampos - 2
              Imprimir_Linea_V Ancho(I), PInicio, PFinal, Negro
           Next I
           Imprimir_Linea_H PosLinea, Ancho(0), Ancho(CantCampos - 3) + 1.5
           PosLinea = PosLinea + 0.05
           PrinterVariables Ancho(CantCampos - 2), PosLinea, Total
           Printer.NewPage
           GrupoNo = .fields("GrupoNo")
           SQLMsg2 = .fields("Detalle_Grupo") & " - " & .fields("GrupoNo")
           EncabezadoData Datas
           Printer.FontSize = SizeLetra
           PInicio = PosLinea
           Total = 0
        End If
        PrinterAllFields CantCampos - 2, PosLinea, Datas, False, False
        PosLinea = PosLinea + 0.38
        Imprimir_Linea_H PosLinea, Ancho(0), Ancho(CantCampos - 3) + 1.5
        PosLinea = PosLinea + 0.05
        PFinal = PosLinea
        Total = Total + .fields("Total")
        If PosLinea >= LimiteAlto Then
           For I = 0 To CantCampos - 2
               Imprimir_Linea_V Ancho(I), PInicio, PFinal, Negro
           Next I
           Imprimir_Linea_H PosLinea, Ancho(0), Ancho(CantCampos)
           Printer.NewPage
           GrupoNo = .fields("GrupoNo")
           SQLMsg2 = .fields("Detalle_Grupo") & " - " & .fields("GrupoNo")
           EncabezadoData Datas
           PInicio = PosLinea
           Printer.FontBold = False
           Printer.FontSize = SizeLetra
        End If
       .MoveNext
     Loop
End If
End With
For I = 0 To CantCampos - 2
    Imprimir_Linea_V Ancho(I), PInicio, PFinal, Negro
Next I
Imprimir_Linea_H PosLinea, Ancho(0), Ancho(CantCampos - 2) + 1.5, Negro, True
PosLinea = PosLinea + 0.05
PrinterVariables Ancho(CantCampos - 2), PosLinea, Total
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

Public Sub ReCalcular_Totales_Factura(TFA As Tipo_Facturas)
Dim FacturaDB As ADODB.Recordset
Dim VPorcIVA As Currency
Dim RUC_Rec As String
    VPorcIVA = 0.12
    sSQL = "SELECT * " _
         & "FROM Facturas " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND Factura = " & TFA.Factura & " " _
         & "AND Autorizacion = '" & TFA.Autorizacion & "' " _
         & "AND Serie = '" & TFA.Serie & "' " _
         & "AND TC = '" & TFA.TC & "' "
    Select_AdoDB FacturaDB, sSQL
    With FacturaDB
     If .RecordCount > 0 Then
         RUC_Rec = .fields("RUC_CI")
         If .fields("Con_IVA") = 0 Then
             VPorcIVA = 0.12
         Else
             VPorcIVA = Redondear(.fields("IVA") / .fields("Con_IVA"), 2)
            'MsgBox "<= " & VPorcIVA
             Select Case VPorcIVA
               Case Is > 0.12: VPorcIVA = 0.14
               Case Else: VPorcIVA = 0.12
             End Select
         End If
     End If
    End With
    FacturaDB.Close
   'MsgBox VPorcIVA
   
    sSQL = "UPDATE Detalle_Factura " _
         & "SET Total = ROUND(Cantidad*Precio,2,0) " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND Factura = " & TFA.Factura & " " _
         & "AND Autorizacion = '" & TFA.Autorizacion & "' " _
         & "AND Serie = '" & TFA.Serie & "' " _
         & "AND TC = '" & TFA.TC & "' "
    Ejecutar_SQL_SP sSQL
    '& "SET Total_IVA = ROUND((Total-(Total_Desc+Total_Desc2))* ROUND(Total_IVA/(Total-(Total_Desc+Total_Desc2)),2,0),4,0) "
    sSQL = "UPDATE Detalle_Factura " _
         & "SET Total_IVA = ROUND((Total-(Total_Desc+Total_Desc2))* " & CStr(VPorcIVA) & ",4,0) " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND Factura = " & TFA.Factura & " " _
         & "AND Autorizacion = '" & TFA.Autorizacion & "' " _
         & "AND Serie = '" & TFA.Serie & "' " _
         & "AND TC = '" & TFA.TC & "' " _
         & "AND Total_IVA <> 0 "
    Ejecutar_SQL_SP sSQL
    
    sSQL = "UPDATE Facturas " _
         & "SET IVA = (SELECT ROUND(SUM(Total_IVA),2,0) " _
         & "           FROM Detalle_Factura " _
         & "           WHERE Detalle_Factura.Total_IVA > 0 " _
         & "           AND Detalle_Factura.TC = Facturas.TC " _
         & "           AND Detalle_Factura.Item = Facturas.Item " _
         & "           AND Detalle_Factura.Periodo = Facturas.Periodo " _
         & "           AND Detalle_Factura.Factura = Facturas.Factura " _
         & "           AND Detalle_Factura.CodigoC = Facturas.CodigoC " _
         & "           AND Detalle_Factura.Serie = Facturas.Serie " _
         & "           AND Detalle_Factura.Autorizacion = Facturas.Autorizacion) " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND Factura = " & TFA.Factura & " " _
         & "AND Autorizacion = '" & TFA.Autorizacion & "' " _
         & "AND Serie = '" & TFA.Serie & "' " _
         & "AND TC = '" & TFA.TC & "' "
    Ejecutar_SQL_SP sSQL
                  
    sSQL = "UPDATE Facturas " _
         & "SET Con_IVA = (SELECT ROUND(SUM(Total),2,0) " _
         & "               FROM Detalle_Factura " _
         & "               WHERE Detalle_Factura.Total_IVA > 0 " _
         & "               AND Detalle_Factura.TC = Facturas.TC " _
         & "               AND Detalle_Factura.Item = Facturas.Item " _
         & "               AND Detalle_Factura.Periodo = Facturas.Periodo " _
         & "               AND Detalle_Factura.Factura = Facturas.Factura " _
         & "               AND Detalle_Factura.CodigoC = Facturas.CodigoC " _
         & "               AND Detalle_Factura.Serie = Facturas.Serie " _
         & "               AND Detalle_Factura.Autorizacion = Facturas.Autorizacion) " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND Factura = " & TFA.Factura & " " _
         & "AND Autorizacion = '" & TFA.Autorizacion & "' " _
         & "AND Serie = '" & TFA.Serie & "' " _
         & "AND TC = '" & TFA.TC & "' "
    Ejecutar_SQL_SP sSQL
    
    sSQL = "UPDATE Facturas " _
         & "SET Sin_IVA = (SELECT ROUND(SUM(Total),2,0) " _
         & "               FROM Detalle_Factura " _
         & "               WHERE Detalle_Factura.Total_IVA <= 0 " _
         & "               AND Detalle_Factura.TC = Facturas.TC " _
         & "               AND Detalle_Factura.Item = Facturas.Item " _
         & "               AND Detalle_Factura.Periodo = Facturas.Periodo " _
         & "               AND Detalle_Factura.Factura = Facturas.Factura " _
         & "               AND Detalle_Factura.CodigoC = Facturas.CodigoC " _
         & "               AND Detalle_Factura.Serie = Facturas.Serie " _
         & "               AND Detalle_Factura.Autorizacion = Facturas.Autorizacion) " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND Factura = " & TFA.Factura & " " _
         & "AND Autorizacion = '" & TFA.Autorizacion & "' " _
         & "AND Serie = '" & TFA.Serie & "' " _
         & "AND TC = '" & TFA.TC & "' "
    Ejecutar_SQL_SP sSQL
    
    sSQL = "UPDATE Facturas " _
         & "SET Descuento = (SELECT ROUND(SUM(Total_Desc),2,0) " _
         & "                 FROM Detalle_Factura " _
         & "                 WHERE Detalle_Factura.Total_Desc > 0 " _
         & "                 AND Detalle_Factura.TC = Facturas.TC " _
         & "                 AND Detalle_Factura.Item = Facturas.Item " _
         & "                 AND Detalle_Factura.Periodo = Facturas.Periodo " _
         & "                 AND Detalle_Factura.Factura = Facturas.Factura " _
         & "                 AND Detalle_Factura.CodigoC = Facturas.CodigoC " _
         & "                 AND Detalle_Factura.Serie = Facturas.Serie " _
         & "                 AND Detalle_Factura.Autorizacion = Facturas.Autorizacion) " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND Factura = " & TFA.Factura & " " _
         & "AND Autorizacion = '" & TFA.Autorizacion & "' " _
         & "AND Serie = '" & TFA.Serie & "' " _
         & "AND TC = '" & TFA.TC & "' "
    Ejecutar_SQL_SP sSQL
    
    sSQL = "UPDATE Facturas " _
         & "SET Descuento2 = (SELECT ROUND(SUM(Total_Desc2),2,0) " _
         & "                  FROM Detalle_Factura " _
         & "                  WHERE Detalle_Factura.Total_Desc2 > 0 " _
         & "                  AND Detalle_Factura.TC = Facturas.TC " _
         & "                  AND Detalle_Factura.Item = Facturas.Item " _
         & "                  AND Detalle_Factura.Periodo = Facturas.Periodo " _
         & "                  AND Detalle_Factura.Factura = Facturas.Factura " _
         & "                  AND Detalle_Factura.CodigoC = Facturas.CodigoC " _
         & "                  AND Detalle_Factura.Serie = Facturas.Serie " _
         & "                  AND Detalle_Factura.Autorizacion = Facturas.Autorizacion) " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND Factura = " & TFA.Factura & " " _
         & "AND Autorizacion = '" & TFA.Autorizacion & "' " _
         & "AND Serie = '" & TFA.Serie & "' " _
         & "AND TC = '" & TFA.TC & "' "
    Ejecutar_SQL_SP sSQL
    
    sSQL = "UPDATE Facturas " _
         & "SET Desc_0 = (SELECT SUM(Total_Desc+Total_Desc2) " _
         & "              FROM Detalle_Factura " _
         & "              WHERE Detalle_Factura.Total_IVA = 0 " _
         & "              AND Detalle_Factura.TC = Facturas.TC " _
         & "              AND Detalle_Factura.Item = Facturas.Item " _
         & "              AND Detalle_Factura.Periodo = Facturas.Periodo " _
         & "              AND Detalle_Factura.Fecha = Facturas.Fecha " _
         & "              AND Detalle_Factura.Factura = Facturas.Factura " _
         & "              AND Detalle_Factura.CodigoC = Facturas.CodigoC " _
         & "              AND Detalle_Factura.Serie = Facturas.Serie " _
         & "              AND Detalle_Factura.Autorizacion = Facturas.Autorizacion) " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND Factura = " & TFA.Factura & " " _
         & "AND Autorizacion = '" & TFA.Autorizacion & "' " _
         & "AND Serie = '" & TFA.Serie & "' " _
         & "AND TC = '" & TFA.TC & "' "
    Ejecutar_SQL_SP sSQL
          
    sSQL = "UPDATE Facturas " _
         & "SET Desc_X = (SELECT SUM(Total_Desc+Total_Desc2) " _
         & "              FROM Detalle_Factura " _
         & "              WHERE Detalle_Factura.Total_IVA > 0 " _
         & "              AND Detalle_Factura.TC = Facturas.TC " _
         & "              AND Detalle_Factura.Item = Facturas.Item " _
         & "              AND Detalle_Factura.Periodo = Facturas.Periodo " _
         & "              AND Detalle_Factura.Fecha = Facturas.Fecha " _
         & "              AND Detalle_Factura.Factura = Facturas.Factura " _
         & "              AND Detalle_Factura.CodigoC = Facturas.CodigoC " _
         & "              AND Detalle_Factura.Serie = Facturas.Serie " _
         & "              AND Detalle_Factura.Autorizacion = Facturas.Autorizacion) " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND Factura = " & TFA.Factura & " " _
         & "AND Autorizacion = '" & TFA.Autorizacion & "' " _
         & "AND Serie = '" & TFA.Serie & "' " _
         & "AND TC = '" & TFA.TC & "' "
    Ejecutar_SQL_SP sSQL
        
    sSQL = "UPDATE Facturas " _
         & "SET IVA = 0 " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND Factura = " & TFA.Factura & " " _
         & "AND Autorizacion = '" & TFA.Autorizacion & "' " _
         & "AND Serie = '" & TFA.Serie & "' " _
         & "AND TC = '" & TFA.TC & "' " _
         & "AND IVA IS NULL "
    Ejecutar_SQL_SP sSQL
    
    sSQL = "UPDATE Facturas " _
         & "SET Con_IVA = 0 " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND Factura = " & TFA.Factura & " " _
         & "AND Autorizacion = '" & TFA.Autorizacion & "' " _
         & "AND Serie = '" & TFA.Serie & "' " _
         & "AND TC = '" & TFA.TC & "' " _
         & "AND Con_IVA IS NULL "
    Ejecutar_SQL_SP sSQL
    
    sSQL = "UPDATE Facturas " _
         & "SET Sin_IVA = 0 " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND Factura = " & TFA.Factura & " " _
         & "AND Autorizacion = '" & TFA.Autorizacion & "' " _
         & "AND Serie = '" & TFA.Serie & "' " _
         & "AND TC = '" & TFA.TC & "' " _
         & "AND Sin_IVA IS NULL "
    Ejecutar_SQL_SP sSQL
    
    sSQL = "UPDATE Facturas " _
         & "SET Descuento = 0 " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND Factura = " & TFA.Factura & " " _
         & "AND Autorizacion = '" & TFA.Autorizacion & "' " _
         & "AND Serie = '" & TFA.Serie & "' " _
         & "AND TC = '" & TFA.TC & "' " _
         & "AND Descuento IS NULL "
    Ejecutar_SQL_SP sSQL
    
    sSQL = "UPDATE Facturas " _
         & "SET Descuento2 = 0 " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND Factura = " & TFA.Factura & " " _
         & "AND Autorizacion = '" & TFA.Autorizacion & "' " _
         & "AND Serie = '" & TFA.Serie & "' " _
         & "AND TC = '" & TFA.TC & "' " _
         & "AND Descuento2 IS NULL "
    Ejecutar_SQL_SP sSQL
    
    sSQL = "UPDATE Facturas " _
         & "SET Desc_0 = 0 " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND Factura = " & TFA.Factura & " " _
         & "AND Autorizacion = '" & TFA.Autorizacion & "' " _
         & "AND Serie = '" & TFA.Serie & "' " _
         & "AND TC = '" & TFA.TC & "' " _
         & "AND Desc_0 IS NULL "
    Ejecutar_SQL_SP sSQL
    
    sSQL = "UPDATE Facturas " _
         & "SET Desc_X = 0 " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND Factura = " & TFA.Factura & " " _
         & "AND Autorizacion = '" & TFA.Autorizacion & "' " _
         & "AND Serie = '" & TFA.Serie & "' " _
         & "AND TC = '" & TFA.TC & "' " _
         & "AND Desc_X IS NULL "
    Ejecutar_SQL_SP sSQL
    
    sSQL = "UPDATE Facturas " _
         & "SET Total_MN = ROUND(Con_IVA + Sin_IVA + IVA + Servicio - Descuento - Descuento2,2,0)," _
         & "SubTotal = ROUND(Con_IVA + Sin_IVA + Servicio,2,0) " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND Factura = " & TFA.Factura & " " _
         & "AND Autorizacion = '" & TFA.Autorizacion & "' " _
         & "AND Serie = '" & TFA.Serie & "' " _
         & "AND TC = '" & TFA.TC & "' "
    Ejecutar_SQL_SP sSQL
    RatonNormal
    
   'Averiguamos si la Factura esta a nombre del Representante
    TBeneficiario = Leer_Datos_Cliente_SP(TFA.CodigoC)
    If TBeneficiario.RUC_CI_Rep <> RUC_Rec Then
       Titulo = "ACTUALIZACION DE DATOS"
       Mensajes = "El RUC asignado originalmente no es el mismo, desea actualizar el nuevo"
       If BoxMensaje = vbYes Then
          sSQL = "UPDATE Facturas " _
               & "SET Estado_SRI = '.', " _
               & "RUC_CI = '" & TBeneficiario.RUC_CI_Rep & "', " _
               & "TB = '" & TBeneficiario.TD_Rep & "', " _
               & "Razon_Social = '" & TBeneficiario.Representante & "' " _
               & "WHERE Item = '" & NumEmpresa & "' " _
               & "AND Periodo = '" & Periodo_Contable & "' " _
               & "AND TC = '" & TFA.TC & "' " _
               & "AND Serie = '" & TFA.Serie & "' " _
               & "AND Factura = " & TFA.Factura & " " _
               & "AND CodigoC = '" & TFA.CodigoC & "' " _
               & "AND Autorizacion = '" & TFA.Autorizacion & "' "
          Ejecutar_SQL_SP sSQL
       End If
    End If
    RatonNormal
End Sub

Public Sub ImprimirAdo_SRI(Datas As Adodc, _
                           SizeLetra As Integer, _
                           Optional EsCampoCorto As Boolean)
On Error GoTo Errorhandler

Mensajes = "Seguro de Imprimir en:" & vbCrLf & Printer.DeviceName & "?"
Titulo = "IMPRESION"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
RatonReloj
InicioX = 0.5: InicioY = 0
DataAnchoCampos InicioX, Datas, SizeLetra, TipoArialNarrow, Orientacion_Pagina, EsCampoCorto
Pagina = 1
Contador = 0
Total_IVA = 0
Total_Con_IVA = 0
Total_Sin_IVA = 0
Total_Desc = 0
SubTotal = 0
'Iniciamos la impresion
Printer.FontBold = False
Codigo = CStr(Porc_IVA * 100)
With Datas.Recordset
If .RecordCount > 0 Then
    .MoveFirst
     EncabezadoData Datas
     Printer.FontSize = SizeLetra
     Printer.FontName = TipoArialNarrow
     Printer.FontBold = False
     Do While Not .EOF
        PrinterAllFields CantCampos, PosLinea, Datas, True, False
        PosLinea = PosLinea + 0.36
        If PosLinea >= LimiteAlto Then
           PosLinea = PosLinea + 0.05
           Imprimir_Linea_H PosLinea, Ancho(0), Ancho(15), Negro
           Printer.NewPage
           EncabezadoData Datas
           Printer.FontSize = SizeLetra
           Printer.FontName = TipoArialNarrow
           Printer.FontBold = False
        End If
        Contador = Contador + 1
        SubTotal = SubTotal + .fields("SubTotal")
        Total_Con_IVA = Total_Con_IVA + .fields("Base_" & Codigo)
        Total_Sin_IVA = Total_Sin_IVA + .fields("Base_0")
        Total_Desc = Total_Desc + .fields("Descuento")
        Total_IVA = Total_IVA + .fields("IVA_" & Codigo)
       .MoveNext
     Loop
End If
End With
PosLinea = PosLinea + 0.05
Imprimir_Linea_H PosLinea, Ancho(0), Ancho(15), Negro, True
PosLinea = PosLinea + 0.05
Total = Total_Con_IVA + Total_Sin_IVA - Total_Desc + Total_IVA
PrinterTexto Ancho(7), PosLinea, "T O T A L E S"
PrinterVariables Ancho(9), PosLinea, SubTotal
PrinterVariables Ancho(10), PosLinea, Total_Desc
PrinterVariables Ancho(11), PosLinea, Total_Con_IVA
PrinterVariables Ancho(12), PosLinea, Total_Sin_IVA
PrinterVariables Ancho(13), PosLinea, Total_IVA
PrinterVariables Ancho(14), PosLinea, Total
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

Public Sub Encabezado_Lista_Alumnos(SizeLet As Integer, FechaInic As String, Grupo As String, Curso As String, Curso1 As String, Optional TipoLineas As Byte)
     PosLinea = 0.01
     Pagina = 0
     Encabezado_Institucion 2, 19
     Printer.FontName = TipoTimes
     Printer.FontBold = True
     Printer.FontSize = 10
     Curso = Leer_Datos_del_Curso(Grupo, 1)
     If Len(Dato_Curso.Paralelo) > 1 Then Curso = Curso & " " & Dato_Curso.Paralelo
     NumeroLineas = PrinterLineasMayor(2.5, PosLinea, Dato_Curso.Descripcion, 16, 0.5)   'Curso
     PosLinea = PosLinea_Aux + 0.5
'''     If Curso1 <> Ninguno Then
'''        PrinterTexto 2.5, PosLinea, Curso1
'''        PosLinea = PosLinea + 0.5
'''     End If
     If TipoLineas = 3 Then
        Printer.FontBold = True
        PrinterTexto 2.5, PosLinea, "PROFESOR(A):"
        Printer.FontBold = False
        PrinterTexto 5.2, PosLinea, Codigo4
        PosLinea = PosLinea + 0.5
        Printer.FontBold = True
        PrinterTexto 2.5, PosLinea, "MATERIA:"
        Printer.FontBold = False
        PrinterTexto 4.6, PosLinea, Codigo3
        PosLinea = PosLinea + 0.5
     End If
    'PrinterTexto 3.2, PosLinea, Grupo & " - " & Curso
     Printer.FontBold = True
     Printer.FontSize = 10
     InicioY = PosLinea
     Imprimir_Linea_H PosLinea, 2, 19.5
     PosLinea = PosLinea + 0.05
     PrinterTexto 3, PosLinea, "N O M I N A "
     Printer.FontSize = SizeLet - 1
     PrinterTexto 2, PosLinea, "No."
     Printer.FontSize = SizeLet
     If J > 1 Then
        Contador = Month(FechaInic)
        PosColumna = 9.6
        For I = 1 To 17
            If PosColumna <= 19 Then
               PrinterTexto PosColumna, PosLinea, UCaseStrg(MidStrg(MesesLetras(CInt(Contador)), 1, 3))
            End If
            PosColumna = PosColumna + 1
            Contador = Contador + 1
            If Contador > 12 Then Contador = 1
        Next I
'''     Else
'''        If TipoLineas = 0 Then PrinterTexto 9.7, PosLinea, "O B S E R V A C I O N E S"
     End If
     PosLinea = PosLinea + 0.5
     Imprimir_Linea_H PosLinea, 2, 19.5
     PosLinea = PosLinea + 0.05
     Printer.FontBold = False
     Printer.FontItalic = False
End Sub

Public Sub Imprimir_Lista_Alumnos(Datas As Adodc, _
                                  FechaInicial As String, _
                                  FechaFinal As String, _
                                  Optional EsCampoCorto As Boolean, _
                                  Optional Sin_Salto_Pagina As Boolean)
On Error GoTo Errorhandler
Dim SizeLetra As Integer
SizeLetra = 8
Mensajes = "Seguro de Imprimir en:" & vbCrLf & Printer.DeviceName & "?"
Titulo = "IMPRESION"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
RatonReloj
J = Redondear((CFechaLong(FechaFinal) - CFechaLong(FechaInicial)) / 30) + 1
InicioX = 0.5: InicioY = 0
DataAnchoCampos InicioX, Datas, SizeLetra, TipoTimes, Orientacion_Pagina, EsCampoCorto
Pagina = 1
'Iniciamos la impresion
Printer.FontBold = False
'MsgBox Sin_Salto_Pagina
With Datas.Recordset
If .RecordCount > 0 Then
    .MoveFirst
     Codigo = .fields("Grupo")
     Codigo1 = .fields("Curso")
     'MsgBox ".."
     Encabezado_Lista_Alumnos SizeLetra, FechaInicial, Codigo, Codigo1, ""
     Do While Not .EOF
        If Codigo <> .fields("Grupo") Then
           PosLinea = PosLinea - 0.1
           InicioY = InicioY + 0.05
           Imprimir_Linea_V 1, InicioY, PosLinea
           Imprimir_Linea_V 1.55, InicioY, PosLinea
           Imprimir_Linea_V 19.5, InicioY, PosLinea
           PosColumna = 9.5
           Imprimir_Linea_V PosColumna, InicioY, PosLinea
           If J > 1 Then
              For I = 1 To 10
                  Imprimir_Linea_V PosColumna, InicioY, PosLinea
                  PosColumna = PosColumna + 1
              Next I
           End If
           'MsgBox Sin_Salto_Pagina
           Printer.NewPage
           Codigo = .fields("Grupo")
           Codigo1 = .fields("Curso")
           Encabezado_Lista_Alumnos SizeLetra, FechaInicial, Codigo, Codigo1, ""
        End If
        Printer.FontSize = SizeLetra
        Printer.FontName = TipoTimes
        PrinterTexto 1, PosLinea, .fields("No_")
        PrinterTexto 1.6, PosLinea, .fields("Cliente")
        PosLinea = PosLinea + 0.35
        Imprimir_Linea_H PosLinea, 1, 19.5
        PosLinea = PosLinea + 0.05
        K = .fields("No_") + 1
        If PosLinea >= LimiteAlto Then
           PosLinea = PosLinea - 0.1
           InicioY = InicioY + 0.05
           Imprimir_Linea_V 1, InicioY, PosLinea
           Imprimir_Linea_V 1.55, InicioY, PosLinea
           Imprimir_Linea_V 19.5, InicioY, PosLinea
           PosColumna = 9.5
           Imprimir_Linea_V PosColumna, InicioY, PosLinea
           If J > 1 Then
              For I = 1 To 10
                  Imprimir_Linea_V PosColumna, InicioY, PosLinea
                  PosColumna = PosColumna + 1
              Next I
           End If
           Printer.NewPage
           Codigo = .fields("Grupo")
           Codigo1 = .fields("Curso")
           Encabezado_Lista_Alumnos SizeLetra, FechaInicial, Codigo, Codigo1, ""
        End If
       .MoveNext
     Loop
End If
End With
PosLinea = PosLinea - 0.1
InicioY = InicioY + 0.05
Imprimir_Linea_V 1, InicioY, PosLinea
Imprimir_Linea_V 1.55, InicioY, PosLinea
Imprimir_Linea_V 19.5, InicioY, PosLinea
PosColumna = 9.5
Imprimir_Linea_V PosColumna, InicioY, PosLinea
If J > 1 Then
   For I = 1 To 10
       Imprimir_Linea_V PosColumna, InicioY, PosLinea
       PosColumna = PosColumna + 1
   Next I
End If
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

Public Sub Encabezado_Lista_Alumnos_Dir(SizeLet As Integer, Grupo As String)
     Encabezados
     Printer.FontBold = True
     Printer.FontSize = 11
     PrinterTexto 1.1, PosLinea, " - " & UCaseStrg(Grupo)
     PosLinea = PosLinea + 0.5
     InicioY = PosLinea
     Imprimir_Linea_H PosLinea, 1, 20
     PosLinea = PosLinea + 0.05
     PrinterTexto 2, PosLinea, "N O M I N A "
     Printer.FontSize = SizeLet - 1
     PrinterTexto 1, PosLinea, "No."
     Printer.FontSize = SizeLet
     'Grado y Direccion
     PrinterTexto 9.7, PosLinea, "GRUPO"
     PrinterTexto 11, PosLinea, "D I R E C C I O N"
     PrinterTexto 18.5, PosLinea, "TELEFONO"
     PosLinea = PosLinea + 0.5
     Imprimir_Linea_H PosLinea, 1, 20
     PosLinea = PosLinea + 0.05
     Printer.FontBold = False
End Sub

Public Sub Imprimir_Lista_Alumnos_Dir(Datas As Adodc, _
                                      NombreGrupo As String, _
                                      Optional Sin_Salto_Pagina As Boolean)
On Error GoTo Errorhandler
Dim SizeLetra As Integer
SizeLetra = 8
Mensajes = "Seguro de Imprimir en:" & vbCrLf & Printer.DeviceName & "? "
Titulo = "IMPRESION"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
RatonReloj
InicioX = 0.5: InicioY = 0
DataAnchoCampos InicioX, Datas, SizeLetra, TipoTimes, Orientacion_Pagina
Pagina = 1
'Iniciamos la impresion
Printer.FontBold = False
With Datas.Recordset
If .RecordCount > 0 Then
    .MoveFirst
     Codigo = .fields("Grupo")
     NombreGrupo = .fields("Curso")
     Encabezado_Lista_Alumnos_Dir SizeLetra, NombreGrupo
     Do While Not .EOF
        If Codigo <> .fields("Grupo") Then
           PosLinea = PosLinea - 0.1
           InicioY = InicioY + 0.05
           Imprimir_Linea_V 1, InicioY, PosLinea
           Imprimir_Linea_V 1.55, InicioY, PosLinea
           Imprimir_Linea_V 9.7, InicioY, PosLinea
           Imprimir_Linea_V 11, InicioY, PosLinea
           Imprimir_Linea_V 18.5, InicioY, PosLinea
           Imprimir_Linea_V 20.1, InicioY, PosLinea
           Printer.NewPage
           Codigo = .fields("Grupo")
           NombreGrupo = .fields("Curso")
           Encabezado_Lista_Alumnos_Dir SizeLetra, NombreGrupo
        End If
        Printer.FontSize = SizeLetra
        Printer.FontName = TipoTimes
        If .fields("T") <> Ninguno Then PrinterTexto 0.6, PosLinea, .fields("T")
        PrinterTexto 1, PosLinea, .fields("No_")
        PrinterTexto 1.6, PosLinea, .fields("Cliente")
        PrinterTexto 9.7, PosLinea, .fields("Grupo")
        PrinterTexto 11, PosLinea, .fields("Domicilio")
        PrinterTexto 18.5, PosLinea, .fields("Telefono")
        PosLinea = PosLinea + 0.35
        Imprimir_Linea_H PosLinea, 1, 20
        PosLinea = PosLinea + 0.05
        K = .fields("No_") + 1
        If PosLinea >= LimiteAlto Then
           PosLinea = PosLinea - 0.1
           InicioY = InicioY + 0.05
           Imprimir_Linea_V 1, InicioY, PosLinea
           Imprimir_Linea_V 1.55, InicioY, PosLinea
           Imprimir_Linea_V 9.7, InicioY, PosLinea
           Imprimir_Linea_V 11, InicioY, PosLinea
           Imprimir_Linea_V 18.5, InicioY, PosLinea
           Imprimir_Linea_V 20.1, InicioY, PosLinea
           Printer.NewPage
           Codigo = .fields("Grupo")
           NombreGrupo = .fields("Curso")
           Encabezado_Lista_Alumnos_Dir SizeLetra, NombreGrupo
        End If
       .MoveNext
     Loop
End If
End With
PosLinea = PosLinea - 0.1
InicioY = InicioY + 0.05
Imprimir_Linea_V 1, InicioY, PosLinea
Imprimir_Linea_V 1.55, InicioY, PosLinea
Imprimir_Linea_V 9.7, InicioY, PosLinea
Imprimir_Linea_V 11, InicioY, PosLinea
Imprimir_Linea_V 18.5, InicioY, PosLinea
Imprimir_Linea_V 20.1, InicioY, PosLinea
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

Public Sub ImprimirCAE(Datas As Adodc, _
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
Ancho(0) = 0.5   ' Manifiesto
Ancho(1) = 2     ' Carta
Ancho(2) = 3.5   ' Cliente
Ancho(3) = 8     ' Mercaderia
Ancho(4) = 11    ' FOB
Ancho(5) = 13.2  ' Flete
Ancho(6) = 14.7  ' Cantidad
Ancho(7) = 16.3  ' Clase
Ancho(8) = 17.5  ' Fecha I
Ancho(9) = 19    ' Fecha DUI
Ancho(10) = 20.5 ' Fecha E
Ancho(11) = 22   ' DUI
Ancho(12) = 23.6 ' Registro
Ancho(13) = 25.5 ' Observacion
Ancho(14) = 27.5 ' Fin
Pagina = 1
'SegundaPagina = False
EnDosPaginas = 0
'MsgBox CantCampos
'Iniciamos la impresion
Printer.FontBold = False
With Datas.Recordset
    .MoveFirst
        EncabezadoData Datas
     Printer.FontSize = SizeLetra
     Do While Not .EOF
        PrinterAllFields CantCampos, PosLinea, Datas, True, False
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
Imprimir_Linea_H PosLinea, InicioX, Ancho(CantCampos), Negro, True
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

Public Sub Imprimir_Pendientes_Facturacion(Datas As Adodc, _
                                           Tipo_Opc As Integer, _
                                           Optional EsCampoCorto As Boolean)
Dim SizeLetra As Integer
Dim LimAncho As Single
Dim CantCampFact As Integer
Dim SepararPorGrupos As Boolean
On Error GoTo Errorhandler
RatonReloj
SepararPorGrupos = False
SizeLetra = 7.5
Mensajes = "Seguro de Imprimir en:" & vbCrLf & Printer.DeviceName & "?"
Titulo = "IMPRESION"
Bandera = False

For IE = 0 To Datas.Recordset.fields.Count - 1
    CantCampos = IE
    If Datas.Recordset.fields(IE).Name = "Total" Then IE = Datas.Recordset.fields.Count + 1
Next IE
CantCampos = CantCampos + 1

If CantCampos >= 9 Then Orientacion_Pagina = 2 Else Orientacion_Pagina = 1
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
InicioX = 0.5: InicioY = 0
DataAnchoCampos InicioX, Datas, SizeLetra, TipoArialNarrow, Orientacion_Pagina, EsCampoCorto

For IE = 0 To Datas.Recordset.fields.Count - 1
    CantCampos = IE
    If Datas.Recordset.fields(IE).Name = "Total" Then IE = Datas.Recordset.fields.Count + 1
Next IE
CantCampos = CantCampos + 1

'MsgBox Datas.Recordset.Fields(CantCampos).Name & vbCrLf & Ancho(CantCampos)
LimAncho = LimiteAncho - 1.5

Pagina = 1
'Iniciamos la impresion
Printer.FontBold = False
Total = 0

Mensajes = "Imprimir separando en grupos la lista?"
Titulo = "FORMA DE IMPRESION"
If BoxMensaje = vbYes Then SepararPorGrupos = True
RatonReloj
With Datas.Recordset
    .MoveFirst
     Encabezado Ancho(0), Ancho(CantCampos)
     Printer.FontItalic = False
     Printer.FontSize = 11
     PrinterTexto CentrarTextoEncab(SQLMsg2, Ancho(0), Ancho(CantCampos)), PosLinea, SQLMsg2
     PosLinea = PosLinea + 0.05
     Printer.FontSize = SizeLetra
     Printer.FontBold = True
     If SepararPorGrupos Then
        PrinterTexto Ancho(0), PosLinea, .fields("Direccion")
        PrinterTexto Ancho(CantCampos - 1), PosLinea, "(" & .fields("Grupo") & ")"
        PosLinea = PosLinea + 0.35
        Imprimir_Linea_H PosLinea, Ancho(0), Ancho(CantCampos)
        PosLinea = PosLinea + 0.05
        EnDosPaginas = 0
        PrinterAllFields CantCampos, PosLinea, Datas, False, True
        For I = 0 To CantCampos
            Printer.Line (Ancho(I), PosLinea - 0.05)-(Ancho(I), PosLinea + 0.45), Negro
        Next I
        PosLinea = PosLinea + 0.35
        Imprimir_Linea_H PosLinea, Ancho(0), Ancho(CantCampos)
        PosLinea = PosLinea + 0.05
        TipoProc = .fields("Grupo")
     Else
        EnDosPaginas = 0
        PrinterAllFields CantCampos, PosLinea, Datas, False, True
     End If
     'MsgBox CantCampFact
     Do While Not .EOF
        If SepararPorGrupos Then
           If TipoProc <> .fields("Grupo") Then
              ' PosLinea = PosLinea + 0.05
              Imprimir_Linea_H PosLinea, Ancho(0), Ancho(CantCampos)
              PosLinea = PosLinea + 0.5
              Printer.FontSize = SizeLetra
              Printer.FontBold = True
              PrinterTexto Ancho(0), PosLinea, .fields("Direccion")
              PrinterTexto Ancho(CantCampos - 1), PosLinea, "(" & .fields("Grupo") & ")"
              PosLinea = PosLinea + 0.35
              Imprimir_Linea_H PosLinea, Ancho(0), Ancho(CantCampos)
              PosLinea = PosLinea + 0.05
              EnDosPaginas = 0
              PrinterAllFields CantCampos, PosLinea, Datas, False, True
              For I = 0 To CantCampos
                  Printer.Line (Ancho(I), PosLinea - 0.05)-(Ancho(I), PosLinea + 0.45), Negro
              Next I
              PosLinea = PosLinea + 0.35
              Imprimir_Linea_H PosLinea, Ancho(0), Ancho(CantCampos)
              PosLinea = PosLinea + 0.05
              TipoProc = .fields("Grupo")
           End If
        End If
        Printer.FontBold = False
        For I = 0 To CantCampos - 1
            PrinterFields Ancho(I), PosLinea, .fields(I), True
        Next I
        Printer.Line (Ancho(CantCampos), PosLinea - 0.05)-(Ancho(CantCampos), PosLinea + 0.45), Negro
        PosLinea = PosLinea + 0.35
        Imprimir_Linea_H PosLinea, Ancho(0), Ancho(CantCampos), Negro
        PosLinea = PosLinea + 0.05
        If Tipo_Opc = 11 Then
           Total = Total + .fields("Total")
        Else
           Total = Total + .fields("Total_MN")
        End If
        If (PosLinea + 0.8) >= LimiteAlto Then
           Printer.NewPage
           Encabezado Ancho(0), Ancho(CantCampos)
           Printer.FontItalic = False
           Printer.FontSize = 11
           PrinterTexto CentrarTextoEncab(SQLMsg2, Ancho(0), Ancho(CantCampos)), PosLinea, SQLMsg2
           PosLinea = PosLinea + 0.05
           Printer.FontSize = SizeLetra
           Printer.FontBold = True
           If SepararPorGrupos Then
              PrinterTexto Ancho(0), PosLinea, .fields("Direccion")
              PrinterTexto Ancho(CantCampos - 1), PosLinea, "(" & .fields("Grupo") & ")"
              PosLinea = PosLinea + 0.35
              Imprimir_Linea_H PosLinea, Ancho(0), Ancho(CantCampos)
              PosLinea = PosLinea + 0.05
              EnDosPaginas = 0
              PrinterAllFields CantCampos, PosLinea, Datas, False, True
              For I = 0 To CantCampos
                  Printer.Line (Ancho(I), PosLinea - 0.05)-(Ancho(I), PosLinea + 0.45), Negro
              Next I
              PosLinea = PosLinea + 0.35
              Imprimir_Linea_H PosLinea, Ancho(0), Ancho(CantCampos)
              PosLinea = PosLinea + 0.05
           Else
              EnDosPaginas = 0
              PrinterAllFields CantCampos, PosLinea, Datas, False, True
           End If
           TipoProc = .fields("Grupo")
        End If
       .MoveNext
     Loop
End With
Imprimir_Linea_H PosLinea, Ancho(0), Ancho(CantCampos), Negro
PosLinea = PosLinea + 0.05
Printer.FontBold = True
'If Tipo_Opc = 3 Then
PrinterTexto Ancho(CantCampos - 2), PosLinea, "T O T A L"
PrinterVariables Ancho(CantCampos - 1), PosLinea, Total
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

Public Sub Calculos_Totales_Factura(TFA As Tipo_Facturas)
Dim AdoDBFA As ADODB.Recordset
  Set AdoDBFA = New ADODB.Recordset
  AdoDBFA.CursorType = adOpenStatic
  AdoDBFA.CursorLocation = adUseClient
  
  TFA.SubTotal = 0
  TFA.Con_IVA = 0
  TFA.Sin_IVA = 0
  TFA.Descuento = 0
  TFA.Total_IVA = 0
  TFA.Total_MN = 0
  TFA.Total_ME = 0
  TFA.Descuento2 = 0
  TFA.Descuento_0 = 0
  TFA.Descuento_X = 0
  TFA.Servicio = 0
  TFA.Utilidad = 0
  
 'Miramos de cuanto es la factura para los calculos de los totales
  Total_Desc_ME = 0
  sSQL = "SELECT * " _
       & "FROM Asiento_F " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' "
  AdoDBFA.open sSQL, AdoStrCnn, , , adCmdText
  With AdoDBFA
   If .RecordCount > 0 Then
       Do While Not .EOF
          TFA.Servicio = TFA.Servicio + .fields("SERVICIO")
          TFA.Descuento = TFA.Descuento + .fields("Total_Desc")
          TFA.Descuento2 = TFA.Descuento2 + .fields("Total_Desc2")
          TFA.Total_IVA = TFA.Total_IVA + .fields("Total_IVA")
          If .fields("Total_IVA") > 0 Then
              TFA.Con_IVA = TFA.Con_IVA + .fields("TOTAL")
              TFA.Descuento_X = TFA.Descuento_X + TFA.Descuento + TFA.Descuento2
          Else
              TFA.Sin_IVA = TFA.Sin_IVA + .fields("TOTAL")
              TFA.Descuento_0 = TFA.Descuento_0 + TFA.Descuento + TFA.Descuento2
          End If
          TFA.Utilidad = TFA.Utilidad + .fields("Utilidad")
         'MsgBox TFA.Sin_IVA
         .MoveNext
       Loop
   End If
  End With
  AdoDBFA.Close
  
  'If Total_Desc_ME > 0 Then Total_Desc = Total_Desc_ME
  TFA.Total_IVA = Redondear(TFA.Total_IVA, 2)
  TFA.Con_IVA = Redondear(TFA.Con_IVA, 2)
  TFA.Sin_IVA = Redondear(TFA.Sin_IVA, 2)
  TFA.Servicio = Redondear(TFA.Servicio, 2)
  TFA.Utilidad = Redondear(TFA.Utilidad, 2)
  TFA.SubTotal = TFA.Sin_IVA + TFA.Con_IVA - TFA.Descuento - TFA.Descuento2
  TFA.Total_MN = TFA.Sin_IVA + TFA.Con_IVA - TFA.Descuento - TFA.Descuento2 + TFA.Total_IVA + TFA.Servicio
End Sub

Public Sub ReCalcular_PVP_Factura(AdoAsientoF As Adodc, Total_FA As Currency)
Dim CantRubros As Single
Dim PVPTemp As Single

  CantRubros = 0
 'Miramos de cuanto es la factura para los calculos de los totales
  With AdoAsientoF.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       Do While Not .EOF
          CantRubros = CantRubros + .fields("CANT")
         .MoveNext
       Loop
       If CantRubros = 0 Then CantRubros = 1
       PVPTemp = Redondear(Total_FA / CantRubros, 8)
      .MoveFirst
       Do While Not .EOF
         .fields("PRECIO") = PVPTemp
         .fields("TOTAL") = Redondear(PVPTemp * .fields("CANT"), 4)
         .MoveNext
       Loop
      .UpdateBatch
   End If
  End With
End Sub

Public Function Digito_Verificador_Modulo11(CadenaNumerica As String) As String
Dim Digito As Long
Dim Mod_11 As Byte
Dim Numero As Integer
Dim Total_Mod_11 As Long
  Mod_11 = 2
  Total_Mod_11 = 0
  CadenaNumerica = Replace(CadenaNumerica, ".", "0")
  For Digito = Len(CadenaNumerica) To 1 Step -1
      Numero = CInt(MidStrg(CadenaNumerica, Digito, 1)) * Mod_11
      Total_Mod_11 = Total_Mod_11 + Numero
      Mod_11 = Mod_11 + 1
      If Mod_11 > 7 Then Mod_11 = 2
  Next Digito
  Mod_11 = Total_Mod_11 Mod 11
  Mod_11 = 11 - Mod_11
  If Mod_11 = 10 Then Mod_11 = 1
  If Mod_11 = 11 Then Mod_11 = 0
  Digito_Verificador_Modulo11 = CStr(Mod_11)
End Function

Public Sub Grabar_Factura(TFA As Tipo_Facturas, VerFactura As Boolean, Optional NoRegTrans As Boolean)
  RatonReloj
 'TFA.CodigoC = Parameto de entrada
 
  'TFA = Leer_Datos_Cliente_FA(TFA)
  
  TFA.T = Pendiente
  If TFA.TC = "FR" Then TFA.TC = "FA"
  If Len(TFA.Autorizacion) >= 13 Then TMail.TipoDeEnvio = "CE"
  Grabar_Factura_SP TFA
  With TFA
       Cadena = ""
      .Hora = Format$(Time, FormatoTimes)
       If Not .Existe_Cliente Then Cadena = Cadena & "No se puedo grabar porque no existe Beneficiario" & vbCrLf
       If .Cantidad_Rubros = 0 Then Cadena = Cadena & "No se puedo grabar porque no existe rubros para facturar" & vbCrLf
       If .GrabadoExitoso Then
           Control_Procesos "G", "Grabar " & TFA.TC & " No. " & TFA.Serie & "-" & Format$(TFA.Factura, "000000000") & " [" & TFA.Hora & "]"
       Else
           Control_Procesos "E", "No se pudo grabar " & TFA.TC & " No. " & TFA.Serie & "-" & Format$(TFA.Factura, "000000000") & " [" & TFA.Hora & "]"
       End If
  End With
  RatonNormal
End Sub

'''Public Sub Grabar_Abonos_Retenciones(FTA As Tipo_Abono)
'''  Control_Procesos "P", FTA.Banco & " " & FTA.TP & " No. " & FTA.Serie & "-" & Format$(FTA.Factura, "0000000") & ", Por: " & Format$(FTA.Abono, "#,##0.00")
'''  With FTA
'''   If .Abono > 0 Then
'''       If .T = "" Or .T = Ninguno Then .T = Normal
'''       If .Cta_CxP = "" Or .Cta_CxP = Ninguno Then .Cta_CxP = Cta_Cobrar
'''       If .CodigoC = "" Or .CodigoC = Ninguno Then .CodigoC = CodigoCliente
'''       If .Comprobante = "" Then .Comprobante = Ninguno
'''       If .Codigo_Inv = "" Then .Codigo_Inv = Ninguno
'''       If .Fecha = Ninguno Then .Fecha = FechaSistema
'''       If .Serie = Ninguno Then .Serie = "001001"
'''       If .Autorizacion = Ninguno Then .Autorizacion = "1234567890"
'''       If .Cheque = Ninguno And DiarioCaja > 0 Then .Cheque = Format$(DiarioCaja, "00000000")
'''       If DiarioCaja > 0 Then .Recibo_No = Format$(DiarioCaja, "0000000000") Else .Recibo_No = "0000000000"
'''       FTA.Tipo_Cta = Leer_Cta_Catalogo(.Cta)
'''       FTA.Tipo_Cta = SubCta
'''       SetAdoAddNew "Trans_Abonos"
'''       SetAdoFields "T", .T
'''       SetAdoFields "TP", .TP
'''       SetAdoFields "Fecha", .Fecha
'''       SetAdoFields "Recibo_No", .Recibo_No
'''       SetAdoFields "Tipo_Cta", .Tipo_Cta
'''       SetAdoFields "Cta", .Cta
'''       SetAdoFields "Cta_CxP", .Cta_CxP
'''       SetAdoFields "Factura", .Factura
'''       SetAdoFields "CodigoC", .CodigoC
'''       SetAdoFields "Abono", .Abono
'''       SetAdoFields "Banco", .Banco
'''       SetAdoFields "Cheque", .Cheque
'''       SetAdoFields "Codigo_Inv", .Codigo_Inv
'''       SetAdoFields "Comprobante", .AutorizacionR
'''       SetAdoFields "EstabRetencion", .Establecimiento
'''       SetAdoFields "PtoEmiRetencion", .Emision
'''       SetAdoFields "Porc", .Porcentaje
'''       SetAdoFields "Serie", .Serie
'''       SetAdoFields "Autorizacion", .Autorizacion
'''       SetAdoFields "Autorizacion_R", .AutorizacionR
'''       SetAdoFields "CodigoU", CodigoUsuario
'''       SetAdoFields "Item", NumEmpresa
'''       SetAdoUpdate
'''   End If
'''  End With
'''  Actualizar_Saldos_Facturas_SP FTA.TP, FTA.Serie, FTA.Factura
''' 'Grabar_Abonos_Periodo_Superior FTA
'''End Sub

'Parametros de entrada: CodigoCliente
Public Sub Grabar_Abonos(FTA As Tipo_Abono, Optional NoRegTrans As Boolean)
Dim CodigoCta As String
  
  With FTA
   CodigoCta = Leer_Cta_Catalogo(.Cta_CxP)
   If Len(CodigoCta) > 1 Then
      CodigoCta = Leer_Cta_Catalogo(.Cta)
     'MsgBox "Abono: " & .Banco & " - " & .Cheque & " - " & .Abono
      If .Abono > 0 And Len(CodigoCta) > 1 And TipoCta = "D" Then
          If .T = "" Or .T = Ninguno Or .T = "A" Then .T = Normal
          If .Cta_CxP = "" Or .Cta_CxP = Ninguno Then .Cta_CxP = Cta_Cobrar
          If .CodigoC = "" Or .CodigoC = Ninguno Then .CodigoC = CodigoCliente
          If .Comprobante = "" Then .Comprobante = Ninguno
          If .Codigo_Inv = "" Then .Codigo_Inv = Ninguno
          If .Fecha = Ninguno Then .Fecha = FechaSistema
          If .Serie = Ninguno Then .Serie = "001001"
          If .Autorizacion = Ninguno Then .Autorizacion = "1234567890"
          If .Cheque = Ninguno And DiarioCaja > 0 Then .Cheque = Format$(DiarioCaja, "00000000")
          If DiarioCaja > 0 Then .Recibo_No = Format$(DiarioCaja, "0000000000") Else .Recibo_No = "0000000000"
         .Tipo_Cta = SubCta
        
          SetAdoAddNew "Trans_Abonos"
          SetAdoFields "T", .T
          SetAdoFields "TP", .TP
          SetAdoFields "Fecha", .Fecha
          SetAdoFields "Recibo_No", .Recibo_No
          SetAdoFields "Tipo_Cta", .Tipo_Cta
          SetAdoFields "Cta", .Cta
          SetAdoFields "Cta_CxP", .Cta_CxP
          SetAdoFields "Factura", .Factura
          SetAdoFields "CodigoC", .CodigoC
          SetAdoFields "Abono", .Abono
          SetAdoFields "Banco", .Banco
          SetAdoFields "Cheque", .Cheque
          SetAdoFields "Codigo_Inv", .Codigo_Inv
          SetAdoFields "Comprobante", .Comprobante
          SetAdoFields "Serie", .Serie
          SetAdoFields "Autorizacion", .Autorizacion
          SetAdoFields "Item", NumEmpresa
          SetAdoFields "CodigoU", CodigoUsuario
          SetAdoFields "Cod_Ejec", CodigoVen
          If .Banco = "NOTA DE CREDITO" Then
              SetAdoFields "Serie_NC", .Serie_NC
              SetAdoFields "Autorizacion_NC", .Autorizacion_NC
              SetAdoFields "Secuencial_NC", .Nota_Credito
          End If
          If Len(.Serie_R) = 6 Then
             SetAdoFields "Serie_R", .Serie_R
             SetAdoFields "Autorizacion_R", .AutorizacionR
             SetAdoFields "Secuencial_R", .Secuencial_R
             SetAdoFields "Porc", .Porcentaje
          End If
          SetAdoUpdate
          If Not NoRegTrans Then
             If .Banco = "NOTA DE CREDITO" Then
                 Control_Procesos "A", "Anulación por " & .Banco & " de " & .TP & " No. " & .Serie & "-" & Format$(.Factura, "000000000")
             Else
                 Control_Procesos "P", "Abono de " & .TP & " No. " & .Serie & "-" & Format$(.Factura, "000000000") & ", Por: " & Format$(.Abono, "#,##0.00")
             End If
          End If
         'Grabar_Abonos_Periodo_Superior FTA
          If FTA.TP = "TJ" Then Imprimir_FA_NV_TJ FTA
      Else
        Control_Procesos "P", "El importe a la Cta (" & .Cta & ") " & .TP & " No. " & .Serie & "-" & Format$(.Factura, "000000000") & ", Por: " & Format$(.Abono, "#,##0.00"), "No se realizo con exito " & .Banco & " " & .Cheque
      End If
   Else
      Control_Procesos "P", "El importe a la Cta (" & .Cta_CxP & ") " & .TP & " No. " & .Serie & "-" & Format$(.Factura, "000000000") & ", Por: " & Format$(.Abono, "#,##0.00"), "No se realizo con exito " & .Banco & " " & .Cheque
   End If
  End With
End Sub

Public Sub Grabar_Anticipos(FTA As Tipo_Abono)
  Control_Procesos "P", "Abono de " & FTA.TP & " No. " & FTA.Serie & "-" & Format$(FTA.Factura, "0000000") & ", Por: " & Format$(FTA.Abono, "#,##0.00")
  With FTA
   If .Abono > 0 Then
       If .T = "" Or .T = Ninguno Then .T = Normal
       If .Cta_CxP = "" Or .Cta_CxP = Ninguno Then .Cta_CxP = Cta_Cobrar
       If .CodigoC = "" Or .CodigoC = Ninguno Then .CodigoC = CodigoCliente
       If .Comprobante = "" Then .Comprobante = Ninguno
       If .Codigo_Inv = "" Then .Codigo_Inv = Ninguno
       If .Fecha = Ninguno Then .Fecha = FechaSistema
       If .Serie = Ninguno Then .Serie = "001001"
       If .Autorizacion = Ninguno Then .Autorizacion = "1234567890"
       If .Cheque = Ninguno And DiarioCaja > 0 Then .Cheque = Format$(DiarioCaja, "00000000")
      .Recibo_No = "0000000000"
       If DiarioCaja > 0 Then .Recibo_No = Format$(DiarioCaja, "0000000000")
       If .CodigoC <> Ninguno And Cta_Aux <> Ninguno Then
           SetAdoAddNew "Trans_SubCtas"
           SetAdoFields "T", "N"
           SetAdoFields "TC", "P"
           SetAdoFields "Fecha", .Fecha
           SetAdoFields "Fecha_V", .Fecha
           SetAdoFields "Cta", .Cta
           SetAdoFields "Codigo", .CodigoC
           SetAdoFields "Comp_No", .Factura
           SetAdoFields "Debitos", .Abono
           SetAdoFields "Serie", .Serie
           SetAdoFields "Autorizacion", .Autorizacion
           SetAdoFields "Item", NumEmpresa
           SetAdoFields "CodigoU", CodigoUsuario
           SetAdoUpdate
       End If
   End If
  End With
End Sub

Public Sub ImprimirRubrosFactura(Datas As Adodc, _
                                 Optional EsCampoCorto As Boolean)
Dim AdoReg As ADODB.Recordset
Dim FormaImp As Byte
Dim SizeLetra As Integer
On Error GoTo Errorhandler
RatonReloj
FormaImp = 1
SizeLetra = 9
Mensajes = "Seguro de Imprimir en:" & vbCrLf & Printer.DeviceName & "?"
Titulo = "IMPRESION"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
InicioX = 0.5: InicioY = 0
DataAnchoCampos InicioX, Datas, SizeLetra, TipoArialNarrow, FormaImp, EsCampoCorto
Set AdoReg = New ADODB.Recordset
AdoReg.CursorType = adOpenStatic
AdoReg.CursorLocation = adUseClient
SQL1 = "SELECT * " _
     & "FROM Catalogo_Productos " _
     & "WHERE TC = 'P' " _
     & "AND Item = '" & NumEmpresa & "' " _
     & "AND Periodo = '" & Periodo_Contable & "' " _
     & "ORDER BY Codigo_Inv "
AdoReg.open SQL1, AdoStrCnn, , , adCmdText
MensajeEncabData = "NOMINA PARA FACTURAR"
SQLMsg2 = "Codigos: "
With AdoReg
 If .RecordCount > 0 Then
     Do While Not .EOF
        SQLMsg2 = SQLMsg2 & .fields("Codigo_Inv") & " - " & .fields("Producto") & " | "
       .MoveNext
     Loop
 End If
End With
Pagina = 1
IR = 0
'Iniciamos la impresion
Printer.FontBold = False
With Datas.Recordset
    .MoveFirst
     Encabezado 1, 19
     Printer.FontBold = True
     PrinterTexto 1, PosLinea, SQLMsg1
     PosLinea = PosLinea + 0.5
     PrinterTexto 1, PosLinea, SQLMsg2
     PosLinea = PosLinea + 0.4
     Printer.Line (0.5, PosLinea)-(19, PosLinea), Negro
     PosLinea = PosLinea + 0.05
     Printer.FontBold = False
     Printer.FontSize = SizeLetra
     Codigo = .fields("Codigo")
     NombreCliente = .fields("Cliente")
     Codigo4 = "| "
     Contador = 0
     Do While Not .EOF
        If Codigo <> .fields("Codigo") Then
           IR = IR + 1
           PrinterTexto 0.5, PosLinea, Format$(IR, "00") & ".-"
           PrinterTexto 1, PosLinea, NombreCliente
           PrinterTexto 7.5, PosLinea, Codigo4
           PosLinea = PosLinea + 0.4
           Printer.Line (0.5, PosLinea)-(19, PosLinea), Negro
           PosLinea = PosLinea + 0.05
           Codigo = .fields("Codigo")
           NombreCliente = .fields("Cliente")
           Codigo4 = "| "
           Contador = 0
        End If
        Contador = Contador + 1
        Codigo4 = Codigo4 & CStr(Contador) & ".- " & .fields("Codigo_Inv") & "=" & Format$(.fields("Valor"), "000.00") & " | "
        If PosLinea >= LimiteAlto Then
           Printer.Line (0.5, PosLinea)-(19, PosLinea), Negro
           Printer.NewPage
           Encabezado 1, 19
           Printer.FontBold = True
           PrinterTexto 1, PosLinea, SQLMsg1
           PosLinea = PosLinea + 0.5
           PrinterTexto 1, PosLinea, SQLMsg2
           PosLinea = PosLinea + 0.4
           Printer.Line (0.5, PosLinea)-(19, PosLinea), Negro
           PosLinea = PosLinea + 0.05
           Printer.FontBold = False
           Printer.FontSize = SizeLetra
        End If
       .MoveNext
     Loop
     IR = IR + 1
     PrinterTexto 0.5, PosLinea, Format$(IR, "00") & ".-"
     PrinterTexto 1, PosLinea, NombreCliente
     PrinterTexto 7.5, PosLinea, Codigo4
     PosLinea = PosLinea + 0.4
    .MoveFirst
End With
Imprimir_Linea_H PosLinea, 0.5, 19, Negro, True
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

Public Sub ImprimirRubrosFacturaGrupo(GrupoNo As String, Optional EsCampoCorto As Boolean)
Dim AdoAlumnos As ADODB.Recordset
Dim AdoDetalle As ADODB.Recordset
Dim FormaImp As Byte
Dim SizeLetra As Integer
On Error GoTo Errorhandler
RatonReloj
FormaImp = 1
SizeLetra = 9
EsCampoCorto = True
Mensajes = "Seguro de Imprimir en:" & vbCrLf & Printer.DeviceName & "?"
Titulo = "IMPRESION"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
InicioX = 0.5: InicioY = 0
Escala_Centimetro 1, TipoArialNarrow, 8
'DataAnchoCampos InicioX, Datas, SizeLetra, TipoArialNarrow, FormaImp, EsCampoCorto
MensajeEncabData = "NOMINA PARA FACTURAR"
SQL1 = "SELECT * " _
     & "FROM Catalogo_Productos " _
     & "WHERE TC = 'P' " _
     & "AND Item = '" & NumEmpresa & "' " _
     & "AND Periodo = '" & Periodo_Contable & "' " _
     & "ORDER BY Codigo_Inv "
Set AdoDetalle = New ADODB.Recordset
AdoDetalle.CursorType = adOpenStatic
AdoDetalle.CursorLocation = adUseClient
AdoDetalle.open SQL1, AdoStrCnn, , , adCmdText
SQLMsg2 = "Codigos: "
With AdoDetalle
 If .RecordCount > 0 Then
     Do While Not .EOF
        SQLMsg2 = SQLMsg2 & .fields("Codigo_Inv") & " - " & .fields("Producto") & " | "
       .MoveNext
     Loop
 End If
End With
AdoDetalle.Close
Pagina = 1
IR = 0
Encabezado 1, 19
Total = 0
SQL1 = "SELECT * " _
     & "FROM Clientes " _
     & "WHERE Grupo = '" & GrupoNo & "' " _
     & "AND FA <> " & Val(adFalse) & " " _
     & "ORDER BY Cliente "
Set AdoAlumnos = New ADODB.Recordset
AdoAlumnos.CursorType = adOpenStatic
AdoAlumnos.CursorLocation = adUseClient
AdoAlumnos.open SQL1, AdoStrCnn, , , adCmdText
Contador = 0
With AdoAlumnos
 If .RecordCount > 0 Then
     Printer.FontBold = True
     PosLinea = PosLinea + 0.1
     PrinterTexto 1, PosLinea, "CURSO: " & .fields("Direccion") & " (" & .fields("Grupo") & ")"
     PosLinea = PosLinea + 0.4
     PrinterTexto 1, PosLinea, SQLMsg2
     PosLinea = PosLinea + 0.4
     Printer.Line (0.5, PosLinea)-(19, PosLinea), Negro
     PosLinea = PosLinea + 0.05
     Printer.FontBold = False
     Do While Not .EOF
        Codigo = .fields("Codigo")
        Codigo1 = .fields("CI_RUC")
        NombreCliente = .fields("Cliente")
        Contador = Contador + 1
        PrinterTexto 0.5, PosLinea, Format$(Contador, "00") & ".-"
        PrinterTexto 1, PosLinea, NombreCliente
        PosColumna = 8
        SQL1 = "SELECT * " _
             & "FROM Clientes_Facturacion " _
             & "WHERE Codigo = '" & Codigo & "' " _
             & "AND Item = '" & NumEmpresa & "' " _
             & "AND Periodo = '" & Periodo_Contable & "' " _
             & "AND Mes = '" & Ninguno & "' " _
             & "ORDER BY Codigo_Inv "
        Set AdoDetalle = New ADODB.Recordset
        AdoDetalle.CursorType = adOpenStatic
        AdoDetalle.CursorLocation = adUseClient
        AdoDetalle.open SQL1, AdoStrCnn, , , adCmdText
        If AdoDetalle.RecordCount > 0 Then
           Do While Not AdoDetalle.EOF
              PrinterTexto PosColumna, PosLinea, AdoDetalle.fields("Codigo_Inv") & " ="
              PosColumna = PosColumna + 1
              PrinterTexto PosColumna, PosLinea, AdoDetalle.fields("Valor")
              PosColumna = PosColumna + 0.6
              PrinterTexto PosColumna, PosLinea, "|"
              PosColumna = PosColumna + 0.2
              Total = Total + AdoDetalle.fields("Valor")
              AdoDetalle.MoveNext
           Loop
        End If
        PrinterTexto 17.5, PosLinea, "|" & Codigo1
        PosLinea = PosLinea + 0.35
        Printer.Line (0.5, PosLinea)-(19, PosLinea), Negro
        PosLinea = PosLinea + 0.05
        If PosLinea >= LimiteAlto Then
           Printer.NewPage
           Encabezado 1, 19
           Printer.FontBold = True
           PosLinea = PosLinea + 0.1
           PrinterTexto 1, PosLinea, .fields("Direccion") & " (" & .fields("Grupo") & ")"
           PosLinea = PosLinea + 0.4
           PrinterTexto 1, PosLinea, SQLMsg2
           PosLinea = PosLinea + 0.4
           Printer.Line (0.5, PosLinea)-(19, PosLinea), Negro
           PosLinea = PosLinea + 0.05
           Printer.FontBold = False
        End If
       .MoveNext
     Loop
 End If
End With
PosLinea = PosLinea + 0.05
PrinterTexto 1, PosLinea, "TOTAL A FACTURAR"
PrinterVariables 7.8, PosLinea, Total
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

Public Sub ImprimirResumenAsientoCaja(Datas As Adodc)
Dim FormaImp As Byte
Dim SizeLetra As Integer
On Error GoTo Errorhandler
RatonReloj
FormaImp = 1
SizeLetra = 8
Mensajes = "Seguro de Imprimir en:" & vbCrLf & Printer.DeviceName & "?"
Titulo = "IMPRESION"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
InicioX = 0.5: InicioY = 0
DataAnchoCampos InicioX, Datas, SizeLetra, TipoTimes, FormaImp
Pagina = 1
'Iniciamos la impresion
Printer.FontBold = False
With Datas.Recordset
    .MoveFirst
     EncabezadoData Datas
     Printer.FontSize = SizeLetra
     Do While Not .EOF
        PrinterFields Ancho(0), PosLinea, .fields("CODIGO")
        PrinterFields Ancho(1), PosLinea, .fields("CUENTA")
        PrinterFields Ancho(2), PosLinea, .fields("PARCIAL_ME")
        PrinterFields Ancho(3), PosLinea, .fields("DEBE")
        PrinterFields Ancho(4), PosLinea, .fields("HABER")
        'PrinterAllFields CantCampos, PosLinea, Datas, True, False
        PosLinea = PosLinea + 0.36
        If PosLinea >= LimiteAlto Then
           Imprimir_Linea_H PosLinea, InicioX, Ancho(CantCampos)
           Printer.NewPage
           EncabezadoData Datas
           Printer.FontSize = SizeLetra
        End If
       .MoveNext
     Loop
End With
Imprimir_Linea_H PosLinea, InicioX, Ancho(CantCampos), Negro, True
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

Public Sub Imprimir_Por_Buses(AdoBus As Adodc, _
                              Encabezado As String)
RatonReloj
Mensajes = "Seguro de Imprimir en:" & vbCrLf & Printer.DeviceName & "?"
Titulo = "IMPRESION"
Bandera = False
Orientacion_Pagina = 2
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
   PosLinea = 2
   With AdoBus.Recordset
       'Encabezado
       .MoveFirst
       'Geneeramos el documento
        tPrint.TipoImpresion = Es_PDF
        tPrint.NombreArchivo = "Por Buses"
        tPrint.TituloArchivo = "Impresion por buses"
        tPrint.TipoLetra = TipoArial
        tPrint.PorteLetra = 8
        tPrint.OrientacionPagina = Orientacion_Pagina
        tPrint.PaginaA4 = True
        tPrint.EsCampoCorto = False
        tPrint.VerDocumento = True
        
        Set cPrint = New cImpresion
        cPrint.iniciaImpresion
    
        cPrint.anchoRegistro 1.7, AdoBus
        Ancho(1) = Ancho(1) + 0.6
        Ancho(2) = Ancho(2) + 1
        Ancho(3) = Ancho(3) + 0.3
        Ancho(4) = Ancho(4) + 3.5
        Ancho(5) = 28.9
        cPrint.tipoNegrilla = True
        cPrint.printImagen LogoTipo, 1, 1, 3, 1.5
        cPrint.printTexto 2.5, 1, "LISTADO DEL " & Encabezado, "C", 16
        cPrint.PorteDeLetra = 8
        cPrint.printTexto 1, PosLinea, "No."
        cPrint.printTexto Ancho(0), PosLinea, "Cliente"
        cPrint.printTexto Ancho(1), PosLinea, "Teléfono"
        cPrint.printTexto Ancho(2), PosLinea, "Paralelo"
        cPrint.printTexto Ancho(3), PosLinea, "Direccion"
        cPrint.printTexto Ancho(4), PosLinea, "Ruta"
        cPrint.tipoNegrilla = False
        PosLinea = PosLinea + 0.4
        cPrint.printLinea 1, PosLinea, Ancho(5), PosLinea, Negro
        PosLinea = PosLinea + 0.05
       'Detalle del Reporte
        Contador = 1
        Do While Not .EOF
           cPrint.printLinea 1, PosLinea, 1, PosLinea + 0.4, Negro
           cPrint.printTexto 1, PosLinea, CStr(Contador)
           For I = 0 To .fields.Count - 1
               cPrint.printFields Ancho(I), PosLinea, .fields(I)
               cPrint.printLinea Ancho(I), PosLinea, Ancho(I), PosLinea + 0.4, Negro
           Next I
           cPrint.printLinea Ancho(I), PosLinea, Ancho(I), PosLinea + 0.4, Negro
           PosLinea = PosLinea + 0.4
           cPrint.printLinea 1, PosLinea, Ancho(I), PosLinea, Negro
           PosLinea = PosLinea + 0.05
           Contador = Contador + 1
          .MoveNext
        Loop
        cPrint.finalizaImpresion
   End With
End If
End Sub

Public Sub Imprimir_Saldo_Clientes(Datas As Adodc, _
                                   SizeLetra As Integer)
On Error GoTo Errorhandler
RatonReloj
Mensajes = "Seguro de Imprimir en:" & vbCrLf & Printer.DeviceName & "?"
Titulo = "IMPRESION"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
InicioX = 1: InicioY = 0
DataAnchoCampos InicioX, Datas, SizeLetra, TipoTimes, Orientacion_Pagina
Ancho(0) = 1      ' Cliente
Ancho(1) = 5.6    ' Factura
Ancho(2) = 7      ' Fecha
Ancho(3) = 8.1    ' Cantidad
Ancho(4) = 10     ' Producto
Ancho(5) = 15.9   ' Ventas
Ancho(6) = 17.7   ' Abonos
Ancho(7) = 19.5   ' Fin
Pagina = 1
TotalIngreso = 0
TotalAbonos = 0
Total = 0
Abono = 0
Saldo = 0
'Iniciamos la impresion
Printer.FontBold = False
Printer.FontSize = SizeLetra
Printer.FontName = TipoArialNarrow
With Datas.Recordset
    .MoveFirst
     Encabezado Ancho(0), Ancho(7)
     NombreCliente = .fields("Cliente")
     Printer.FontSize = 10
     PrinterTexto Ancho(0), PosLinea, SQLMsg1
     PosLinea = PosLinea + 0.45
     Encabezado_Saldo_Clientes
     Printer.FontName = TipoArialNarrow
     Printer.FontSize = SizeLetra
     Factura_No = .fields("Factura")
     PrinterTexto Ancho(1), PosLinea, Format$(.fields("Factura"), "0000000")
     Do While Not .EOF
        If Factura_No <> .fields("Factura") Then
           Saldo = Total - Abono
           PosLinea = PosLinea + 0.05
           Imprimir_Linea_H PosLinea, Ancho(5), Ancho(7), Negro
           PosLinea = PosLinea + 0.05
           PrinterTexto Ancho(5), PosLinea, "SALDO"
           PrinterVariables Ancho(6), PosLinea, Saldo
           PosLinea = PosLinea + 0.36
           Imprimir_Linea_H PosLinea, Ancho(1), Ancho(7), Negro
           If NombreCliente = .fields("Cliente") Then
              PosLinea = PosLinea + 0.05
              PrinterTexto Ancho(1), PosLinea, Format$(.fields("Factura"), "0000000")
           End If
           Factura_No = .fields("Factura")
           Total = 0
           Abono = 0
        End If
        If NombreCliente <> .fields("Cliente") Then
           NombreCliente = .fields("Cliente")
           Encabezado_Saldo_Clientes
           Printer.FontSize = SizeLetra
           PrinterTexto Ancho(1), PosLinea, Format$(.fields("Factura"), "0000000")
           Factura_No = .fields("Factura")
        End If
        PrinterFields Ancho(2), PosLinea, .fields("Fecha")
        PrinterFields Ancho(3), PosLinea, .fields("Cantidad")
        PrinterTexto Ancho(4), PosLinea, PrinterTextoMaximo(.fields("Producto"), 5.5)
        PrinterFields Ancho(5), PosLinea, .fields("Ventas")
        PrinterFields Ancho(6), PosLinea, .fields("Abonos")
        Total = Total + .fields("Ventas")
        Abono = Abono + .fields("Abonos")
        
        TotalIngreso = TotalIngreso + .fields("Ventas")
        TotalAbonos = TotalAbonos + .fields("Abonos")
        PosLinea = PosLinea + 0.36
        If PosLinea >= LimiteAlto Then
           Imprimir_Linea_H PosLinea, InicioX, Ancho(CantCampos)
           Printer.NewPage
           Encabezado Ancho(0), Ancho(7)
           Printer.FontSize = 10
           PrinterTexto Ancho(0), PosLinea, SQLMsg1
           PosLinea = PosLinea + 0.45
           Encabezado_Saldo_Clientes
           Printer.FontName = TipoArialNarrow
           Printer.FontSize = SizeLetra
           Factura_No = .fields("Factura")
           PrinterTexto Ancho(1), PosLinea, Format$(.fields("Factura"), "0000000")
           PrinterFields Ancho(2), PosLinea, .fields("Fecha")
           PrinterFields Ancho(3), PosLinea, .fields("Cantidad")
           PrinterTexto Ancho(4), PosLinea, PrinterTextoMaximo(.fields("Producto"), 5.5)
           PrinterFields Ancho(5), PosLinea, .fields("Ventas")
           PrinterFields Ancho(6), PosLinea, .fields("Abonos")
           PosLinea = PosLinea + 0.36
        End If
       .MoveNext
     Loop
     Saldo = Total - Abono
     PosLinea = PosLinea + 0.05
     Imprimir_Linea_H PosLinea, Ancho(5), Ancho(7), Negro
     PosLinea = PosLinea + 0.05
     PrinterTexto Ancho(5), PosLinea, "SALDO"
     PrinterVariables Ancho(6), PosLinea, Saldo
     PosLinea = PosLinea + 0.36
End With
Saldo = TotalIngreso - TotalAbonos
Imprimir_Linea_H PosLinea, InicioX, Ancho(CantCampos), Negro, True
PosLinea = PosLinea + 0.05
PrinterTexto Ancho(5), PosLinea, "T O T A L"
PrinterVariables Ancho(6), PosLinea, Saldo
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

Public Sub Imprimir_Resumen_Productos(Datas As Adodc, _
                                      SizeLetra As Integer)
On Error GoTo Errorhandler
RatonReloj
Mensajes = "Seguro de Imprimir en:" & vbCrLf & Printer.DeviceName & "?"
Titulo = "IMPRESION"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
InicioX = 1: InicioY = 0
DataAnchoCampos InicioX, Datas, SizeLetra, TipoTimes, Orientacion_Pagina
Ancho(0) = 1      ' Cliente
Ancho(1) = 6      ' Cantidad
Ancho(2) = 7.5    ' Producto
Ancho(3) = 13     ' Codigo
Ancho(4) = 16     ' IVA
Ancho(5) = 17.5   ' Ventas
Ancho(6) = 19.5   ' Fin
Pagina = 1
Total = 0
'Iniciamos la impresion
Printer.FontBold = False
Printer.FontSize = SizeLetra
Printer.FontName = TipoArialNarrow
With Datas.Recordset
    .MoveFirst
     Encabezado Ancho(0), Ancho(6)
     NombreCliente = .fields("Cliente")
     Printer.FontSize = 10
     PrinterTexto Ancho(0), PosLinea, SQLMsg1
     PosLinea = PosLinea + 0.45
     Encabezado_Resumen_Productos
     Printer.FontSize = SizeLetra
     Printer.FontName = TipoArialNarrow
     Do While Not .EOF
        If NombreCliente <> .fields("Cliente") Then
           NombreCliente = .fields("Cliente")
           Encabezado_Resumen_Productos
           Printer.FontSize = SizeLetra
        End If
        PrinterFields Ancho(1), PosLinea, .fields("Cant_Prod")
        PrinterTexto Ancho(2), PosLinea, PrinterTextoMaximo(.fields("Producto"), 4)
        PrinterFields Ancho(3), PosLinea, .fields("Codigo")
        PrinterFields Ancho(4), PosLinea, .fields("IVA")
        PrinterFields Ancho(5), PosLinea, .fields("Ventas")
        Total = Total + .fields("Ventas")
        PosLinea = PosLinea + 0.36
        If PosLinea >= LimiteAlto Then
           Imprimir_Linea_H PosLinea, InicioX, Ancho(CantCampos)
           Printer.NewPage
           Encabezado Ancho(0), Ancho(6)
           Printer.FontSize = 10
           PrinterTexto Ancho(0), PosLinea, SQLMsg1
           PosLinea = PosLinea + 0.45
           Encabezado_Resumen_Productos
           Printer.FontSize = SizeLetra
           Printer.FontName = TipoArialNarrow
           PrinterFields Ancho(1), PosLinea, .fields("Cant_Prod")
           PrinterTexto Ancho(2), PosLinea, PrinterTextoMaximo(.fields("Producto"), 5)
           PrinterFields Ancho(3), PosLinea, .fields("Codigo")
           PrinterFields Ancho(4), PosLinea, .fields("IVA")
           PrinterFields Ancho(5), PosLinea, .fields("Ventas")
           PosLinea = PosLinea + 0.36
        End If
       .MoveNext
     Loop
End With
Imprimir_Linea_H PosLinea, Ancho(0), Ancho(6), Negro, True
PosLinea = PosLinea + 0.05
PrinterTexto Ancho(4), PosLinea, "T O T A L"
PrinterVariables Ancho(5), PosLinea, Total
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

Public Sub Encabezado_Cartera_Resumen(Datas As Adodc, Optional SegundaPagina As Boolean)
Dim InicX As Single
Dim InicY As Single
Encabezado Ancho(0), AnchoPapel
PorteLetra = 8
LetraAnterior = Printer.FontName
Printer.FontName = TipoArialNarrow
Printer.FontBold = False
If SQLMsg1 <> "" Then
   Printer.FontSize = 12
   PrinterTexto Ancho(0), PosLinea, SQLMsg1
   PosLinea = PosLinea + 0.45
End If
If SQLMsg2 <> "" Then
   Printer.FontSize = 10
   PrinterTexto Ancho(0), PosLinea, SQLMsg2
   PosLinea = PosLinea + 0.4
End If
If SQLMsg3 <> "" Then
   Printer.FontSize = 8
   PrinterTexto Ancho(0), PosLinea, SQLMsg3
   PosLinea = PosLinea + 0.4
End If
PosLinea = PosLinea + 0.05
Printer.FontSize = PorteLetra
Printer.FontBold = True
Printer.FontUnderline = True
Printer.FontItalic = True
PrinterTexto Ancho(0), PosLinea, "Nombre del Cliente"
PrinterTexto Ancho(1), PosLinea, "Telefono"
PrinterTexto Ancho(2), PosLinea, "T"
PrinterTexto Ancho(3), PosLinea, "Fecha"
PrinterTexto Ancho(4), PosLinea, "Fecha V."
PrinterTexto Ancho(5), PosLinea, "Factura"
PrinterTexto Ancho(6), PosLinea, "Total"
PrinterTexto Ancho(7), PosLinea, "Abono"
PrinterTexto Ancho(8), PosLinea, "Saldo"
PrinterTexto Ancho(9), PosLinea, "D. Mora"
PrinterTexto Ancho(10), PosLinea, "Sector"
PrinterTexto Ancho(11), PosLinea, "Chq_Posf"
Printer.FontUnderline = False
Printer.FontBold = False
Printer.FontItalic = False
PosLinea = PosLinea + 0.4
Printer.FontSize = PorteLetra
Printer.FontName = LetraAnterior
End Sub

Public Sub Encabezado_Cartera_Resumen_Vendedor(Datas As Adodc, Optional SegundaPagina As Boolean)
Dim InicX As Single
Dim InicY As Single
PorteLetra = 7
LetraAnterior = Printer.FontName
Printer.FontName = TipoArialNarrow
Printer.FontItalic = False
Printer.FontBold = True
Printer.FontUnderline = False
If SQLMsg1 <> "" Then
   Printer.FontSize = 12
   PrinterTexto Ancho(0), PosLinea, SQLMsg1
   PosLinea = PosLinea + 0.45
End If
If SQLMsg2 <> "" Then
   Printer.FontSize = 9
   PrinterTexto Ancho(0), PosLinea, SQLMsg2
   If SQLMsg3 <> "" Then
      PrinterTexto Ancho(10), PosLinea, SQLMsg3
   End If
   PosLinea = PosLinea + 0.4
End If
PosLinea = PosLinea + 0.05
Printer.FontSize = PorteLetra
Printer.FontUnderline = True
Printer.FontItalic = True
PrinterTexto Ancho(0), PosLinea, "Nombre del Cliente"
PrinterTexto Ancho(1), PosLinea, "Telefono"
PrinterTexto Ancho(2), PosLinea, "T"
PrinterTexto Ancho(3), PosLinea, "Fecha"
PrinterTexto Ancho(4), PosLinea, "Fecha V."
PrinterTexto Ancho(5), PosLinea, "Serie"
PrinterTexto Ancho(6), PosLinea, "Factura"
PrinterTexto Ancho(7), PosLinea, "Total"
PrinterTexto Ancho(8), PosLinea, "Abono"
PrinterTexto Ancho(9), PosLinea, "Saldo"
PrinterTexto Ancho(10), PosLinea, "D. Mora"
PrinterTexto Ancho(11), PosLinea, "Chq_Posf"
Printer.FontUnderline = False
Printer.FontBold = False
Printer.FontItalic = False
PosLinea = PosLinea + 0.4
Printer.FontSize = PorteLetra
Printer.FontName = LetraAnterior
End Sub

Public Sub ImprimirNotaEgreso(DataFact As Adodc, DataDet As Adodc, NumFact As Long)
'Establecemos Espacios y seteos de impresion
On Error GoTo Errorhandler
Mensajes = "Imprmir Nota de Egreso de la Factura No. " & NumFact
Titulo = "IMPRESION"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
   Escala_Centimetro 1, TipoTimes, 10
   LetraAnterior = Printer.FontName
   RatonReloj
   Printer.FontBold = True
   Printer.FontName = TipoTimes
   sSQL = "SELECT F.*,C.Cliente,C.CI_RUC,C.Telefono,C.Direccion,C.Grupo " _
        & "FROM Facturas As F,Clientes As C " _
        & "WHERE F.Factura = " & NumFact & " " _
        & "AND C.Codigo = F.CodigoC " _
        & "AND F.TC <> 'C' " _
        & "AND F.TC <> 'P' " _
        & "AND F.Item = '" & NumEmpresa & "' " _
        & "AND F.Periodo = '" & Periodo_Contable & "' "
   Select_Adodc DataFact, sSQL
   With DataFact.Recordset
    If .RecordCount > 0 Then
        Printer.FontSize = 24
        PrinterTexto 1, 0.1, Empresa
        Printer.FontSize = 14
        Cadena = "Nota de Entrega No. " & FechaAnio(.fields("Fecha")) & "-" & Format$(.fields("Factura"), "000000") & "  "
        PrinterTexto TextoDerecha(Cadena), 1, Cadena
        Cadena = "Fecha: " & FechaStrgCorta(.fields("Fecha")) & "."
        PrinterTexto 1, 1, Cadena
        Printer.FontSize = 10
        PrinterFieldText 1, 2.2, "Cliente: ", .fields("Cliente")
        Cadena = "Factura No. " & Format$(.fields("Factura"), "0000000") & "  "
        PrinterTexto TextoDerecha(Cadena), 2.1, Cadena
    End If
   End With
   sSQL = "SELECT * " _
        & "FROM Detalle_Factura " _
        & "WHERE Factura = " & NumFact & " " _
        & "AND TC <> 'C' " _
        & "AND TC <> 'P' " _
        & "AND Item = '" & NumEmpresa & "' " _
        & "AND Periodo = '" & Periodo_Contable & "' "
   Select_Adodc DataDet, sSQL
   With DataDet.Recordset
    If .RecordCount > 0 Then
        PFil = 3
        Printer.FontSize = 10
        PrinterTexto 1.5, PFil, "CODIGO"
        PrinterTexto 4, PFil, "CANTIDAD"
        PrinterTexto 7, PFil, "P R O D U C T O "
        Printer.Line (1, PFil)-(19, PFil), Negro
        Printer.Line (1, PFil + 0.5)-(19, PFil + 0.5), Negro
        PFil = 3.6: Printer.FontBold = False
        Printer.FontBold = False
       .MoveFirst
        Do While (Not .EOF)
           PrinterFields 1.5, PFil, .fields("Codigo")
           PrinterFields 4, PFil, .fields("Cantidad")
           PrinterFields 7, PFil, .fields("Producto")
           PFil = PFil + 0.4
          .MoveNext
        Loop
    End If
   End With
   PFil = 10
   Printer.Line (2.5, PFil)-(6.5, PFil), Negro
   Printer.Line (9.5, PFil)-(13.5, PFil), Negro
   PFil = PFil + 0.1: Printer.FontSize = 10
   PrinterTexto 3, PFil, "DESPACHADO POR."
   PrinterTexto 10, PFil, "RECIBI CONFORME."
   Printer.FontName = LetraAnterior
   RatonNormal
   MensajeEncabData = ""
   Printer.EndDoc  ' La impresión ha terminado.
End If
Exit Sub
Errorhandler:
    RatonNormal
    ErrorDeImpresion
    Exit Sub
End Sub

'''Public Sub ImprimirEtiqueta(PosInic As Single, _
'''                            PosLinea1 As Single, _
'''                            DtaE As Adodc, _
'''                            PictBar As PictureBox)
'''  Code39.AlturaBarra = 20
'''  Code39.TamBarra = 1
'''  Code39.ColorCodigo = "N"
'''  With DtaE.Recordset
'''       Printer.FontSize = 6
'''       PrinterFields PosInic + 0.1, PosLinea1 + 0.75, .Fields("Cliente")
'''       Cadena = "AT. "
'''       If .Fields("Atencion") <> Ninguno Then Cadena = Cadena & .Fields("Atencion")
'''       If Cadena <> "" Then PrinterTexto PosInic + 0.05, PosLinea1 + 1, Cadena
'''       PrinterFields PosInic + 4.2, PosLinea1 + 0.75, .Fields("Telefono")
'''       PrinterFields PosInic + 4.2, PosLinea1 + 1, .Fields("Ciudad")
'''       PrinterFields PosInic + 0.1, PosLinea1 + 1.25, .Fields("Direccion")
'''       PrinterTexto PosInic + 0.05, PosLinea1 + 1.5, "No. (" & .Fields("DirNumero") & ")"
'''       PrinterFields PosInic + 0.1, PosLinea1 + 1.75, .Fields("Sector")
'''       Code39.ValorCodigo = .Fields("Contrato_No")
'''       Code39.RealizarCodigo
'''       ImprimirCodigoBarra PosInic + 1.5, PosLinea1 + 1.5, Code39.ValorCodigo, PictBar
'''       PrinterPaint LogoTipo, PosInic + 4.1, PosLinea1 + 0.01, 1.4, 0.65
'''  End With
'''End Sub

Public Sub ImprimirEtiquetas(Datas As Adodc, PictBar As PictureBox)
Dim SizeLetra As Integer
SizeLetra = 6
On Error GoTo Errorhandler
RatonReloj
Mensajes = "Seguro de Imprimir en:" & vbCrLf & Printer.DeviceName & "?"
Titulo = "IMPRESION"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
InicioX = 1: InicioY = 0
DataAnchoCampos InicioX, Datas, SizeLetra, TipoTimes, 1
'Iniciamos la impresion
Printer.FontBold = True
With Datas.Recordset
    .MoveFirst
     Printer.FontSize = SizeLetra
     Pagina = 1
     PosLinea = 1
     PosColumna = 0.4
     Do While Not .EOF
     
        ''ImprimirEtiqueta PosColumna, PosLinea, Datas, PictBar
        
        PosColumna = PosColumna + 6.95
        Pagina = Pagina + 1
        If Pagina > 3 Then
           PosLinea = PosLinea + 2.55
           Pagina = 1
           PosColumna = 0.4
        End If
        'If PosLinea >= 26 Then PosLinea = PosLinea + 2.5
        If PosLinea >= 26 Then
           Printer.NewPage
           Pagina = 1
           PosLinea = 1
           PosColumna = 0.4
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

Public Sub ImprimirSuscriptores(Datas As Adodc)
Dim SizeLetra As Integer
SizeLetra = 6
On Error GoTo Errorhandler
RatonReloj
Bandera = False
Mensajes = "Seguro de Imprimir en:" & vbCrLf & Printer.DeviceName & "?"
Titulo = "IMPRESION"
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
InicioX = 0.5: InicioY = 0
DataAnchoCampos InicioX, Datas, SizeLetra, TipoTimes, 2
Ancho(0) = 0.5    'T
Ancho(1) = 0.5    'Fecha_I
Ancho(2) = 1.6    'Fecha_F
Ancho(3) = 2.7    'Contrato
Ancho(4) = 4      'Cliente
Ancho(5) = 9.5    'Atencion
Ancho(6) = 14.5   'Telefono
Ancho(7) = 15.7   'Ciudad
Ancho(8) = 17.5   'Sector
Ancho(9) = 18.5   'Direccion
Ancho(10) = 24.5  'DirNumero
Ancho(11) = 24.5  'Contador
Ancho(12) = 24.5  'Email
Ancho(13) = 24.5  'Firma
Ancho(14) = 27.5  'Fin
Pagina = 1
Contador = 0
'Iniciamos la impresion
With Datas.Recordset
 If .RecordCount > 0 Then
    .MoveFirst
     EncabezadoData Datas
     TipoDoc = .fields("Ciudad")
     TipoComp = .fields("Sector")
     Do While Not .EOF
        If TipoDoc <> .fields("Ciudad") Or TipoComp <> .fields("Sector") Then
           For I = 0 To CantCampos
               Printer.Line (Ancho(I), 3.7)-(Ancho(I), PosLinea), Negro
           Next I
           Printer.Line (Ancho(0), PosLinea)-(Ancho(CantCampos), PosLinea), Negro
           PosLinea = PosLinea + 0.05
           PrinterTexto Ancho(0), PosLinea, "TOTAL SUSCRIPCIONES"
           PrinterVariables Ancho(3) + 1, PosLinea, Contador
           Contador = 0
           Printer.NewPage
           EncabezadoData Datas
           Printer.FontSize = SizeLetra
           TipoDoc = .fields("Ciudad")
           TipoComp = .fields("Sector")
        End If
        Printer.FontBold = False
        Printer.FontSize = SizeLetra
        PrinterFields Ancho(1), PosLinea, .fields("Fecha_I")
        PrinterFields Ancho(2), PosLinea, .fields("Fecha_F")
        PrinterFields Ancho(3), PosLinea, .fields("Contrato_No")
        PrinterFields Ancho(4), PosLinea, .fields("Cliente")
        PrinterFields Ancho(5), PosLinea, .fields("Atencion")
        PrinterFields Ancho(6), PosLinea, .fields("Telefono")
        PrinterFields Ancho(7), PosLinea, .fields("Ciudad")
        PrinterFields Ancho(8), PosLinea, .fields("Sector")
        PrinterFields Ancho(13), PosLinea, .fields("Firma")
        Cadena1 = "[" & .fields("Contador") & "] " & .fields("Direccion")
        If Len(.fields("DirNumero")) > 1 Then Cadena1 = Cadena1 & ", No. " & .fields("DirNumero")
        NumeroLineas = PrinterLineasMayor(Ancho(9), PosLinea, Cadena1, 6, 0.3)
        PosLinea = PosLinea + 0.6
        'PrinterAllFields CantCampos, PosLinea, Datas, True
        Printer.Line (Ancho(0), PosLinea)-(Ancho(CantCampos), PosLinea), Negro
        PosLinea = PosLinea + 0.05
        If PosLinea >= LimiteAlto Then
           For I = 0 To CantCampos
               Printer.Line (Ancho(I), 3.7)-(Ancho(I), PosLinea), Negro
           Next I
           Imprimir_Linea_H PosLinea, InicioX, Ancho(CantCampos)
           Printer.NewPage
           EncabezadoData Datas
           Printer.FontSize = SizeLetra
        End If
        Contador = Contador + 1
       .MoveNext
     Loop
     For I = 0 To CantCampos
         Printer.Line (Ancho(I), 3.7)-(Ancho(I), PosLinea), Negro
     Next I
     Imprimir_Linea_H PosLinea, InicioX, Ancho(CantCampos)
 End If
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

Public Sub ImprimirCobrosAbonos(Datas As Adodc, _
                                FormaImp As Byte, _
                                SizeLetra As Integer)
On Error GoTo Errorhandler
RatonReloj
Mensajes = "Seguro de Imprimir en:" & Chr(13) & Chr(13) & Printer.DeviceName & "?"
Titulo = "IMPRESION"
If BoxMensaje = vbYes Then
InicioX = 0.5: InicioY = 0
'Escala_Centimetro FormaImp, TipoTimes, SizeLetra
DataAnchoCampos InicioX, Datas, SizeLetra, TipoTimes, FormaImp
Pagina = 1: Total = 0: TotalInteres = 0
'Iniciamos la impresion
Printer.FontBold = False
With Datas.Recordset
  If .RecordCount > 0 Then
     .MoveFirst
      EncabezadoData Datas
      Printer.FontSize = SizeLetra
      PrinterFields Ancho(1), PosLinea, .fields("Cliente")
      PrinterFields Ancho(6), PosLinea, .fields("Usuario")
      Codigo = .fields("Usuario")
      Codigo1 = .fields("Cliente")
      Do While Not .EOF
        'PrinterAllFields CantCampos, PosLinea, Datas, True
         If Codigo <> .fields("Usuario") Then
            PosLinea = PosLinea + 0.05
            Imprimir_Linea_H PosLinea, InicioX, Ancho(CantCampos)
            PosLinea = PosLinea + 0.1
            PrinterVariables Ancho(4), PosLinea, Total
            PrinterVariables Ancho(5), PosLinea, CSng(TotalInteres)
            PrinterVariables Ancho(6), PosLinea, "T O T A L: " & Format$(Total + TotalInteres, "#,##0.00")
            PosLinea = PosLinea + 0.4
            Imprimir_Linea_H PosLinea, InicioX, Ancho(CantCampos)
            PosLinea = PosLinea + 0.1
            PrinterFields Ancho(1), PosLinea, .fields("Cliente")
            PrinterFields Ancho(6), PosLinea, .fields("Usuario")
            Codigo = .fields("Usuario")
            Codigo1 = .fields("Cliente")
            Total = 0
            TotalInteres = 0
         End If
         If Codigo1 <> .fields("Cliente") Then
            PrinterFields Ancho(1), PosLinea, .fields("Cliente")
            Codigo1 = .fields("Cliente")
         End If
         PrinterFields Ancho(0), PosLinea, .fields("CH")
         PrinterFields Ancho(2), PosLinea, .fields("Contrato_No")
         PrinterFields Ancho(3), PosLinea, .fields("Cuota_No")
         PrinterFields Ancho(4), PosLinea, .fields("Abono")
         PrinterFields Ancho(5), PosLinea, .fields("Interes")
         Total = Total + .fields("Abono")
         TotalInteres = TotalInteres + .fields("Interes")
         PosLinea = PosLinea + 0.36
         If PosLinea >= LimiteAlto Then
            Imprimir_Linea_H PosLinea, InicioX, Ancho(CantCampos)
            Printer.NewPage
            EncabezadoData Datas
            Printer.FontSize = SizeLetra
         End If
        .MoveNext
      Loop
      PosLinea = PosLinea + 0.05
      Imprimir_Linea_H PosLinea, InicioX, Ancho(CantCampos)
      PosLinea = PosLinea + 0.1
      PrinterVariables Ancho(4), PosLinea, Total
      PrinterVariables Ancho(5), PosLinea, CSng(TotalInteres)
      PrinterVariables Ancho(6), PosLinea, "T O T A L: " & Format$(Total + TotalInteres, "#,##0.00")
      PosLinea = PosLinea + 0.4
  End If
End With
Imprimir_Linea_H PosLinea, InicioX, Ancho(CantCampos), Negro, True
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

Public Sub ImprimirAbonoPagos(Datas As Data, FechaI As String, CodigoC As String, Contrato As String)
On Error GoTo Errorhandler
Dim SizeLetra As Integer
RatonReloj
Mensajes = "Seguro de Imprimir en:" & Chr(13) & Chr(13) & Printer.DeviceName & "?"
Titulo = "IMPRESION"
If BoxMensaje = vbYes Then
InicioX = 0.5: InicioY = 0
SizeLetra = 10
Escala_Centimetro 1, TipoTimes, SizeLetra
Pagina = 1: Total = 0: TotalInteres = 0
'Iniciamos la impresion
sSQL = "SELECT * " _
     & "FROM Clientes " _
     & "WHERE Codigo = '" & CodigoC & "' "
Select_Adodc Datas, sSQL
With Datas.Recordset
 If .RecordCount > 0 Then
     Codigo2 = MensajeEncabData
     MensajeEncabData = "RECIBO DE CAJA"
     Encabezado 2, 17
     Printer.FontBold = True
     PrinterTexto 3, 2, "NOMBRES: "
     PrinterTexto 3, 2.5, "DIRECCION:"
     PrinterTexto 3, 3, "TELEFONO:"
     PrinterTexto 10, 3, Codigo2
     Printer.FontBold = False
     PrinterTexto 5.5, 2, .fields("Cliente")
     PrinterTexto 5.5, 2.5, .fields("Direccion")
     PrinterTexto 5.5, 3, .fields("Telefono")
     sSQL = "SELECT AM.Contrato_No,AM.Cuota_No,AM.Abono,AM.Interes,CM.Concepto " _
          & "FROM Abono_Meses As AM,Contratos_Meses As CM " _
          & "WHERE AM.Codigo = '" & CodigoC & "' " _
          & "AND AM.Fecha = #" & BuscarFecha(FechaI) & "# " _
          & "AND AM.Contrato_No = CM.Contrato_No " _
          & "AND AM.Cuota_No = CM.Pago_No " _
          & "ORDER BY AM.Contrato_No,AM.Cuota_No "
     Select_Adodc Datas, sSQL
     If Datas.Recordset.RecordCount > 0 Then
      Printer.FontBold = False
      PosLinea = 3.5
      Printer.Line (3, PosLinea)-(17, PosLinea), Negro
      PosLinea = PosLinea + 0.1
      PrinterTexto 3, PosLinea, "CONTRATO No."
      PrinterTexto 6, PosLinea, "CUOTAS"
      PrinterTexto 8, PosLinea, "D E T A L L E"
      PrinterTexto 14, PosLinea, "A B O N O"
      PosLinea = PosLinea + 0.4
      Printer.Line (3, PosLinea)-(17, PosLinea), Negro
      PosLinea = PosLinea + 0.1
     .MoveFirst
      Total = 0: TotalInteres = 0
      Printer.FontSize = SizeLetra
      Do While Not Datas.Recordset.EOF
         PrinterFields 3, PosLinea, Datas.Recordset.fields("Contrato_No")
         PrinterFields 6.5, PosLinea, Datas.Recordset.fields("Cuota_No")
         PrinterFields 8, PosLinea, Datas.Recordset.fields("Concepto")
         PrinterFields 14, PosLinea, Datas.Recordset.fields("Abono")
         Total = Total + Datas.Recordset.fields("Abono")
         TotalInteres = TotalInteres + Datas.Recordset.fields("Interes")
         PosLinea = PosLinea + 0.36
         If PosLinea >= LimiteAlto Then
            Imprimir_Linea_H PosLinea, 3, 17
            Printer.NewPage
            Encabezado 1, 16
            Printer.FontSize = SizeLetra
         End If
         Datas.Recordset.MoveNext
      Loop
      PosLinea = PosLinea + 0.1
      Imprimir_Linea_H PosLinea, 3, 17
      PosLinea = PosLinea + 0.1
      PrinterTexto 13.5, PosLinea, "T O T A L"
      PrinterVariables 14, PosLinea, Total
      If TotalInteres > 0 Then
         PosLinea = PosLinea + 0.5
         PrinterTexto 13.5, PosLinea, "I N T E R E S"
         PrinterVariables 14, PosLinea, TotalInteres
      End If
      PosLinea = PosLinea + 1.2
      PrinterTexto 4, PosLinea, "_____________"
      PrinterTexto 10, PosLinea, "________________"
      PosLinea = PosLinea + 0.4
      PrinterTexto 4.5, PosLinea, "C A J A"
      PrinterTexto 10.5, PosLinea, "C L I E N T E"
     End If
 End If
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

Public Sub GeneraTablaPrestamoFijo(BoxMiFecha As String, _
                                   DtaTabla As Adodc, _
                                   DBG_Tabla As DataGrid, _
                                   TxtInt As String, _
                                   TxtMeses As String, _
                                   TxtMonto As String, _
                                   Meses_Dias As Boolean, _
                                   ConIVA As Boolean, _
                                   TipoPrestamo As String)
  Interes = Redondear(Val(TxtInt) / 100, 4)
 'MsgBox Interes
  Numero = Redondear(Val(TxtMeses))
  Total = Redondear(Val(TxtMonto))
  Mifecha = BoxMiFecha
  Saldo = Total:  Valor_ME = 0
  Total_ME = 0:  Valor = 0: Comision = 0
  RatonReloj
  DBG_Tabla.Visible = False
  If Total > 0 And Numero > 0 Then
  With DtaTabla.Recordset
    If Meses_Dias Then   'Si_No = True Dias else Meses
       For I = 0 To 1
          .AddNew
          .fields("Cuotas") = I
          .fields("Dia") = DiasLetras(Weekday(Mifecha))
          .fields("TP") = TipoPrestamo
           NoDias = Numero
           If I = 0 Then
              If Interes > 0 Then
                 Valor_ME = Redondear(((Total * Interes) / 360) * NoDias)
              Else
                 Valor_ME = 0
              End If
              Total_ME = Redondear(Total)
              
              Valor = Redondear(Total + Valor_ME)
              
             .fields("Fecha") = Mifecha
             .fields("Capital") = Total_ME
             .fields("Interes") = Valor_ME
             .fields("Abono") = 0
             .fields("Saldo") = Redondear(Total + Valor_ME)
              Mifecha = CLongFecha(CFechaLong(Mifecha) + Numero)
           Else
             .fields("Fecha") = Mifecha
             .fields("Capital") = 0
             .fields("Interes") = 0
             .fields("Abono") = Total
             .fields("Saldo") = 0
           End If
          .Update
       Next I
    Else
       Total = Redondear(Total + (Total * ((Numero / 12) * Interes)))
       Tasa = 0
       Do
         Tasa = Redondear(Tasa + 0.0001, 4)
         Cuota = Redondear(((Saldo * Tasa) / 12) / (1 - (1 + (Tasa / 12)) ^ -Numero))
       Loop Until (Cuota * Numero) >= Total
       Contador = 1: Total = Saldo
       Valor = Redondear(((12 * Total) + (Total * Interes * Numero)) / (12 * Numero))
       Valor_ME = 0: Total_ME = 0: Comision = 0
       For I = 0 To Numero
          .AddNew
          .fields("CodigoU") = CodigoUsuario
          .fields("T_No") = Trans_No
          .fields("Item") = NumEmpresa
          .fields("Cuotas") = I
          .fields("Dia") = DiasLetras(Weekday(Mifecha))
           If I = 0 Then
             .fields("Fecha") = Mifecha
             .fields("Capital") = 0
             .fields("Interes") = 0
             '.Fields("Abono") = 0
              Mifecha = SiguienteMes(Mifecha)
           Else
              If ConIVA Then
                 Comision = Redondear((Total_ME + Valor_ME) * Porc_IVA)
              Else
                 Comision = 0
              End If
             .fields("Fecha") = Mifecha
             .fields("Capital") = Total_ME
             .fields("Interes") = Valor_ME
             .fields("Pagos") = Total_ME + Valor_ME
              If ConIVA = 1 Then
                .fields("Abono") = Valor + (Valor * Porc_IVA)
              Else
                '.Fields("Abono") = Valor
              End If
              Mifecha = SiguienteMes(Mifecha)
           End If
          .fields("Saldo") = Total
          .fields("TP") = TipoPrestamo
          .Update
          'Comision del 1%
          'Comision = Redondear((Total_ME + Valor_ME) * Porc_IVA)
          'Interes Inicial
           If Interes > 0 Then
              Valor_ME = Redondear(Total * (Tasa / 12))
           Else
              Valor_ME = 0
           End If
          'Amortizacion o Capital
           Total_ME = Redondear(Valor - Valor_ME)
          'Saldo Pendiente
           Total = Redondear(Total - Total_ME)
          'Interes Final
           Valor_ME = Redondear(Valor - Total_ME)
          'Total = Redondear(Total - Total_ME)
           Contador = Contador + 1
       Next I
    End If
  End With
  If Meses_Dias = False Then
     With DtaTabla.Recordset
      If .RecordCount > 0 Then
         .MoveLast
         'Comision del 1%
          Valor = Redondear(.fields("Interes"))
          Total = Redondear(.fields("Capital"))
          'Abono = Redondear(.Fields("Abono"))
          Saldo = Redondear(.fields("Saldo"))
          If ConIVA Then
             Comision = Redondear((Total + Saldo) * Porc_IVA)
          Else
             Comision = 0
          End If
          If Valor > 0 Then
            .fields("Interes") = Abono - Total - Saldo - Comision
          Else
            .fields("Interes") = 0
            '.Fields("Abono") = Total + Saldo + Comision
          End If
         .fields("Capital") = Total + Saldo
          
         '.Fields("IVA") = Comision
         .fields("Saldo") = 0
         .Update
      End If
     End With
  End If
  End If
  RatonNormal
  DBG_Tabla.Visible = True
End Sub

Public Sub EncabReciboCaja(OpcionImp As Byte, Cliente As String, Datas As Adodc)
PorteLetra = Printer.FontSize
LetraAnterior = Printer.FontName
Printer.FontName = TipoTimes
Printer.FontSize = 10
InicioY = PosLinea
Printer.FontBold = False
'Iniciamos la impresion
With Datas.Recordset
  Select Case OpcionImp
    Case 1
         Dibujo = RutaSistema & "\FORMATOS\RECIBO.GIF"
         PrinterPaint Dibujo, 0.5, PosLinea, 16.5, 3
         PrinterTexto 2, PosLinea + 0.7, .fields("Fecha")
         PrinterTexto 2.5, PosLinea + 1.1, Cliente
         PrinterTexto 3.6, PosLinea + 1.5, "Pago de Facturas."
         PosLinea = PosLinea + 3
    Case 2
         Dibujo = RutaSistema & "\FORMATOS\RECIBO1.GIF"
         PrinterPaint Dibujo, 0.5, PosLinea, 16.5, 2
         PosLinea = PosLinea + 2
  End Select
  Printer.FontBold = True: Printer.FontSize = 14
  PrinterTexto 14.5, InicioY + 0.2, Format$(.fields("Recibo_No"), "000000")
  Printer.FontBold = False
End With
Printer.FontSize = PorteLetra
Printer.FontName = LetraAnterior
End Sub

Public Sub Encabezado_Diario_Caja_Ventas(Optional Pos_X As Single, Optional PorteLetraFin As Integer)
  Printer.FontBold = True
  Printer.FontItalic = True
  Printer.FontUnderline = True
  InicioY = PosLinea
  If Orientacion_Pagina = 1 Then Printer.FontSize = 9 Else Printer.FontSize = 8
 'Iniciamos la impresion
  If Len(Beneficiario) > 1 Then
     PrinterTexto Ancho(0) + Pos_X, PosLinea, "Cajero(a):"
     Printer.FontUnderline = False
     PrinterTexto Ancho(1) + 1.5 + Pos_X, PosLinea, Beneficiario
     Printer.FontUnderline = True
     PosLinea = PosLinea + 0.4
  End If
  PrinterTexto Ancho(0) + Pos_X, PosLinea, "Fecha/C l i e n t e"
  PrinterTexto Ancho(2) + Pos_X, PosLinea, "Factura"
  PrinterTexto Ancho(3) + Pos_X, PosLinea, "Total_IVA"
  PrinterTexto Ancho(4) + Pos_X, PosLinea, "Descuentos"
  PrinterTexto Ancho(5) + Pos_X, PosLinea, "Servicio"
  PrinterTexto Ancho(6) + Pos_X, PosLinea, "Total Fact."
  Printer.FontUnderline = False
  Printer.FontItalic = False
  Printer.FontBold = False
  If PorteLetraFin <= 0 Then PorteLetraFin = 12
  Printer.FontSize = PorteLetraFin
  If Orientacion_Pagina = 1 Then PosLinea = PosLinea + 0.5 Else PosLinea = PosLinea + 0.36
End Sub

Public Sub Encabezado_Diario_Caja_CxC(Optional Pos_X As Single, Optional PorteLetraFin As Integer)
  Printer.FontBold = True
  Printer.FontItalic = True
  Printer.FontUnderline = True
  InicioY = PosLinea
  If Orientacion_Pagina = 1 Then Printer.FontSize = 9 Else Printer.FontSize = 8
 'Iniciamos la impresion
  If Len(Beneficiario) > 1 Then
     PrinterTexto Ancho(0) + Pos_X, PosLinea, "Cajero(a):"
     Printer.FontUnderline = False
     PrinterTexto Ancho(1) + Pos_X + 1.5, PosLinea, Beneficiario
     Printer.FontUnderline = True
     PosLinea = PosLinea + 0.4
  End If
  PrinterTexto Ancho(0) + Pos_X, PosLinea, "Fecha/C l i e n t e"
  PrinterTexto Ancho(2) + Pos_X, PosLinea, "Factura"
  PrinterTexto Ancho(3) + Pos_X, PosLinea, "Detalle del Abono"
  PrinterTexto Ancho(5) + Pos_X, PosLinea, "Chq/Dep/Bau"
  PrinterTexto Ancho(6) + Pos_X, PosLinea, "Total Abono"
  Printer.FontUnderline = False
  Printer.FontItalic = False
  Printer.FontBold = False
  If PorteLetraFin <= 0 Then PorteLetraFin = 12
  Printer.FontSize = PorteLetraFin
  If Orientacion_Pagina = 1 Then PosLinea = PosLinea + 0.5 Else PosLinea = PosLinea + 0.36
End Sub

Public Sub Encabezado_Saldo_Clientes()
Dim NombreCli As String
  PorteLetra = Printer.FontSize
  LetraAnterior = Printer.FontName
  Printer.FontName = TipoTimes
  Printer.FontBold = True
  Printer.FontItalic = False
  Printer.FontUnderline = True
  InicioY = PosLinea
 'Iniciamos la impresion
  Printer.FontSize = 8
  PosLinea = PosLinea + 0.05
  PrinterTexto Ancho(0), PosLinea, PrinterTextoMaximo(NombreCliente, 4)
  PrinterTexto Ancho(1), PosLinea, "Factura"
  PrinterTexto Ancho(2), PosLinea, "F e c h a"
  PrinterTexto Ancho(3), PosLinea, "Cantidad"
  PrinterTexto Ancho(4), PosLinea, "D e t a l l e"
  PrinterTexto Ancho(5), PosLinea, "V e n t a s"
  PrinterTexto Ancho(6), PosLinea, "A b o n o s"
  Printer.FontBold = False
  Printer.FontUnderline = False
  Printer.FontSize = PorteLetra
  Printer.FontName = LetraAnterior
  PosLinea = PosLinea + 0.4
End Sub

Public Sub Encabezado_Resumen_Productos()
  Printer.FontBold = True
  Printer.FontUnderline = True
  Printer.FontItalic = True
  InicioY = PosLinea
  Printer.FontSize = 10
 'Iniciamos la impresion
  PosLinea = PosLinea + 0.05
  PrinterTexto Ancho(0), PosLinea, PrinterTextoMaximo(NombreCliente, 4.5)
  PrinterTexto Ancho(1), PosLinea, "Cantidad"
  PrinterTexto Ancho(2), PosLinea, "Producto"
  PrinterTexto Ancho(3), PosLinea, "Codigo"
  PrinterTexto Ancho(4), PosLinea, "I.V.A."
  PrinterTexto Ancho(5), PosLinea, "V E N T A S"
  Printer.FontBold = False
  Printer.FontUnderline = False
  Printer.FontItalic = False
  Printer.FontSize = 12
  PosLinea = PosLinea + 0.4
End Sub

Public Sub Encabezado_Diario_Caja_Inv(Optional Pos_X As Single, Optional PorteLetraFin As Integer)
  Printer.FontBold = True
  Printer.FontItalic = True
  Printer.FontUnderline = True
  InicioY = PosLinea
 'Iniciamos la impresion
  If Orientacion_Pagina = 1 Then Printer.FontSize = 9 Else Printer.FontSize = 8
  PrinterTexto Ancho(0) + Pos_X, PosLinea, "Codigo Inventario"
  PrinterTexto Ancho(1) + 1 + Pos_X, PosLinea, "P r o d u c t o"
  PrinterTexto Ancho(5) + Pos_X, PosLinea, "Entrada"
  PrinterTexto Ancho(6) + Pos_X, PosLinea, "Salida"
  Printer.FontUnderline = False
  Printer.FontItalic = False
  Printer.FontBold = False
  If PorteLetraFin <= 0 Then PorteLetraFin = 12
  Printer.FontSize = PorteLetraFin
  If Orientacion_Pagina = 1 Then PosLinea = PosLinea + 0.5 Else PosLinea = PosLinea + 0.36
End Sub

Public Sub Encabezado_Diario_Caja_Prod(Optional Pos_X As Single, Optional PorteLetraFin As Integer)
  Printer.FontBold = True
  Printer.FontItalic = True
  Printer.FontUnderline = True
  InicioY = PosLinea
  If Orientacion_Pagina = 1 Then Printer.FontSize = 9 Else Printer.FontSize = 8
 'Iniciamos la impresion
  PrinterTexto Ancho(1) + Pos_X, PosLinea, "P r o d u c t o"
  PrinterTexto Ancho(2) + Pos_X, PosLinea, "Codigo Inv."
  PrinterTexto Ancho(3) + Pos_X + 0.5, PosLinea, "Cantidad"
  PrinterTexto Ancho(4) + Pos_X, PosLinea, "SubTotal"
  PrinterTexto Ancho(5) + Pos_X, PosLinea, "Total IVA"
  PrinterTexto Ancho(6) + Pos_X, PosLinea, "TOTALES"
  Printer.FontUnderline = False
  Printer.FontItalic = False
  Printer.FontBold = False
  If PorteLetraFin <= 0 Then PorteLetraFin = 12
  Printer.FontSize = PorteLetraFin
  If Orientacion_Pagina = 1 Then PosLinea = PosLinea + 0.5 Else PosLinea = PosLinea + 0.36
End Sub

Public Sub Encabezado_Diario_Anticipos(Optional Pos_X As Single, Optional PorteLetraFin As Integer)
  Printer.FontBold = True
  Printer.FontItalic = True
  Printer.FontUnderline = True
  InicioY = PosLinea
  If Orientacion_Pagina = 1 Then Printer.FontSize = 9 Else Printer.FontSize = 8
 'Iniciamos la impresion
  PrinterTexto Ancho(0) + Pos_X, PosLinea, "Fecha/Cuenta"
  PrinterTexto Ancho(2) + Pos_X, PosLinea, "Comp. No."
  PrinterTexto Ancho(3) + Pos_X + 0.3, PosLinea, "Cliente"
  PrinterTexto Ancho(6) + Pos_X, PosLinea, "Total Abono"
  Printer.FontUnderline = False
  Printer.FontItalic = False
  Printer.FontBold = False
  If PorteLetraFin <= 0 Then PorteLetraFin = 12
  Printer.FontSize = PorteLetraFin
  If Orientacion_Pagina = 1 Then PosLinea = PosLinea + 0.5 Else PosLinea = PosLinea + 0.36
End Sub

Public Sub CalcularSaldos(DataIngC As Data, SaldoFact As Double)
  With DataIngC.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       SumaSaldo = SaldoFact
      .Edit
       Efectivo = .fields("Efectivo")
       Retencion = .fields("Retencion")
       Cheque = .fields("Cheque")
       SumaSaldo = SumaSaldo - (Efectivo + Retencion + Cheque)
      .fields("Valor") = SaldoFact
      .fields("Total_Abono") = Efectivo + Cheque + Retencion
      .fields("Saldo") = SumaSaldo
      .Update
      .MoveNext
       Do While Not .EOF
         .Edit
          Efectivo = .fields("Efectivo")
          Retencion = .fields("Retencion")
          Cheque = .fields("Cheque")
          Abono = SumaSaldo
          SumaSaldo = SumaSaldo - (Efectivo + Retencion + Cheque)
         .fields("Valor") = Abono
         .fields("Total_Abono") = Efectivo + Cheque + Retencion
         .fields("Saldo") = SumaSaldo
         .Update
         .MoveNext
       Loop
   End If
  End With
End Sub

Public Sub Imprimir_Facturas_CxC(MiForm As Form, _
                                 TFA As Tipo_Facturas, _
                                 Optional ReImp As Boolean, _
                                 Optional EsMatricula As Boolean, _
                                 Optional PorOrdenFactura As Boolean, _
                                 Optional Imprimir_Asc As Boolean, _
                                 Optional SinCopia As Boolean, _
                                 Optional CheqSinCodigo As Boolean)
Dim AdoDBFac As ADODB.Recordset
Dim AdoDBDet As ADODB.Recordset
Dim AdoDBAux As ADODB.Recordset

Dim Posicion As Single
Dim Salto_de_Factura As Single
Dim Imp_No_Factuas As String
Dim Cad_Tipo_Pago As String
Dim AND_BETWEEN_Facturas As String

On Error GoTo Errorhandler

RatonReloj
With TFA
 If .Desde < .Hasta Then
     Imp_No_Factuas = "Desde la " & .Desde & " hasta la " & .Hasta
     If .TC = "PV" Then
         AND_BETWEEN_Facturas = "AND F.Ticket BETWEEN " & .Desde & " and " & .Hasta & " "
     Else
         AND_BETWEEN_Facturas = "AND F.Factura BETWEEN " & .Desde & " and " & .Hasta & " "
     End If
 Else
     Imp_No_Factuas = "No. " & .Factura
     If .TC = "PV" Then
         AND_BETWEEN_Facturas = "AND F.Ticket = " & .Factura & " "
     Else
         AND_BETWEEN_Facturas = "AND F.Factura = " & .Factura & " "
     End If
 End If
End With
'Espacios entre las Facturas
 Salto_de_Factura = TFA.AltoFactura + TFA.EspacioFactura
 If Salto_de_Factura <= 0 Then Salto_de_Factura = 0

Mensajes = "Imprimir Facturas" & vbCrLf & Imp_No_Factuas
Titulo = "IMPRESION"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
   Escala_Centimetro Orientacion_Pagina, TipoArialNarrow, 9
   
   RutaOrigen = RutaSistema & "\FORMATOS\" & TFA.LogoFactura & ".GIF"
  'MsgBox LogoFactura & vbCrLf & AnchoFactura & vbCrLf & AltoFactura
   If TFA.LogoFactura = "MATRICIA" Then
      sSQL = "UPDATE Formato_Propio " _
           & "SET Texto = 'CLIENTE:' " _
           & "WHERE TP = 'IF' " _
           & "AND Num = 2 "
      Ejecutar_SQL_SP sSQL
      sSQL = "UPDATE Formato_Propio " _
           & "SET Texto = 'ALUMNA:' " _
           & "WHERE TP = 'IF' " _
           & "AND Num = 6 "
      Ejecutar_SQL_SP sSQL
'''      Else
'''          sSQL = "UPDATE Formato_Propio " _
'''               & "SET Texto = 'ALUMNA:' " _
'''               & "WHERE TP = 'IF' " _
'''               & "AND Num = 2 "
'''          Ejecutar_SQL_SP sSQL
'''          sSQL = "UPDATE Formato_Propio " _
'''               & "SET Texto = 'CLIENTE:' " _
'''               & "WHERE TP = 'IF' " _
'''               & "AND Num = 6 "
'''          Ejecutar_SQL_SP sSQL
'''      End If
   End If
   
   Posicion = 0
   Cadena1 = MiForm.Caption
   If ReImp Then Imp_No_Factuas = "Reimpresion " Else Imp_No_Factuas = "Impresión "
   Imp_No_Factuas = TFA.TC & "/" & TFA.Serie & "/" & TFA.Autorizacion & "/" & Imp_No_Factuas
  'MsgBox TFA.Tipo_PRN
   Select Case TFA.Tipo_PRN
     Case "CP", "FM": CEConLineas = ProcesarSeteos("FM")
     Case "OP": CEConLineas = ProcesarSeteos("OP")
     Case Else: CEConLineas = ProcesarSeteos("FA")
   End Select
   
   Control_Procesos "F", Imp_No_Factuas
   CopiarComp = False
   If Len(TFA.Autorizacion) < 13 Then
      If Not SinCopia Then
         Mensajes = "Imprimir con copia"
         Titulo = "Pregunta de Impresion"
         If BoxMensaje = vbYes Then CopiarComp = True
      End If
   End If
   Pagina = 1: PosLinea = 0.01: PosColumna = 0.01
  'Iniciamos la impresion
   If TFA.TC = "PV" Then
      sSQL = "SELECT F.*,C.Cliente,C.CI_RUC,C.Telefono,C.Direccion,C.Ciudad,C.Grupo,C.Email " _
           & "FROM Trans_Ticket As F,Clientes As C " _
           & "WHERE F.Item = '" & NumEmpresa & "' " _
           & "AND F.Periodo = '" & Periodo_Contable & "' " _
           & "AND F.TC = '" & TFA.TC & "' " _
           & AND_BETWEEN_Facturas _
           & "AND C.Codigo = F.CodigoC "
      If PorOrdenFactura Then
         If Imprimir_Asc Then
            sSQL = sSQL & "ORDER BY F.Ticket,C.Grupo,C.Cliente "
         Else
            sSQL = sSQL & "ORDER BY F.Ticket DESC,C.Grupo,C.Cliente "
         End If
      Else
         sSQL = sSQL & "ORDER BY C.Grupo,C.Cliente,F.Ticket "
      End If
   Else
      sSQL = "SELECT F.*,C.Cliente,C.CI_RUC,C.Telefono,C.TelefonoT,C.Direccion,C.DireccionT," _
           & "C.Grupo,C.Codigo,C.Ciudad,C.Email,C.TD,C.DirNumero " _
           & "FROM Facturas As F,Clientes As C " _
           & "WHERE F.Item = '" & NumEmpresa & "' " _
           & "AND F.Periodo = '" & Periodo_Contable & "' " _
           & "AND F.TC = '" & TFA.TC & "' " _
           & "AND F.Serie = '" & TFA.Serie & "' " _
           & AND_BETWEEN_Facturas _
           & "AND C.Codigo = F.CodigoC "
      If PorOrdenFactura Then
         If Imprimir_Asc Then
            sSQL = sSQL & "ORDER BY F.Factura,C.Grupo,C.Cliente "
         Else
            sSQL = sSQL & "ORDER BY F.Factura DESC,C.Grupo,C.Cliente "
         End If
      Else
         sSQL = sSQL & "ORDER BY C.Grupo,C.Cliente,F.Factura "
      End If
   End If
   Select_AdoDB AdoDBFac, sSQL
   
   'MsgBox sSQL & vbCrLf & AdoDBFac.RecordCount
   
   Printer.FontBold = False
   Printer.FontName = TipoArialNarrow
   With AdoDBFac
    If .RecordCount > 0 Then
        Do While Not .EOF
           TFA.Fecha = .fields("Fecha")
           TFA.Cta_CxP = .fields("Cta_CxP")
           TFA.Cod_CxC = .fields("Cod_CxC")
           TFA.Vencimiento = .fields("Vencimiento")
           TFA.Fecha_Aut = .fields("Fecha_Aut")
           TFA.Serie = .fields("Serie")
           TFA.Autorizacion = .fields("Autorizacion")
           TFA.Factura = .fields("Factura")
           TFA.CodigoC = .fields("Codigo")
           TFA.Saldo_Pend = .fields("Saldo_MN")
           TFA.Imp_Mes = .fields("Imp_Mes")
           ReImp = CBool(.fields("P"))
           TFA.Tipo_Pago = .fields("Tipo_Pago")
           Cad_Tipo_Pago = Ninguno
           Leer_Datos_FA_NV TFA
           
          'MsgBox TFA.LogoFactura
           If Len(TFA.Autorizacion) >= 13 Then            'Si es electronica en formato PDF
             'MsgBox SetPapelPRN
              Imprimir_FA_NV_Electronica TFA
              'If SetPapelPRN > 100 Then Imprimir_FA_NV_Electronica TFA Else SRI_Generar_PDF_FA TFA, True, True
           Else                                          'Caso contrario en impresion normal
               sSQL = "SELECT * " _
                    & "FROM Tabla_Referenciales_SRI " _
                    & "WHERE Tipo_Referencia = 'FORMA DE PAGO' " _
                    & "AND Codigo = '" & TFA.Tipo_Pago & "' "
               Select_AdoDB AdoDBAux, sSQL
               If AdoDBAux.RecordCount > 0 Then Cad_Tipo_Pago = AdoDBAux.fields("Descripcion")
               AdoDBAux.Close
               
               TextoBanco = Ninguno
               TextoCheque = Ninguno
               TextoFormaPago = ""
               
               TFA.Educativo = False
    '''           sSQL = "SELECT * " _
    '''                & "FROM Clientes_Matriculas " _
    '''                & "WHERE Periodo = '" & Periodo_Contable & "' " _
    '''                & "AND Item = '" & NumEmpresa & "' " _
    '''                & "AND Codigo = '" & TFA.CodigoC & "' "
    '''           Select_AdoDB AdoDBAux, sSQL
    '''           If AdoDBAux.RecordCount > 0 Then TFA.Educativo = True
    '''           AdoDBAux.Close
               
               sSQL = "SELECT CC.TC,CC.Cuenta,TA.Fecha,TA.CodigoC,TA.Abono,TA.Banco,TA.Cheque " _
                    & "FROM Catalogo_Cuentas CC, Trans_Abonos As TA " _
                    & "WHERE CC.Item = '" & NumEmpresa & "' " _
                    & "AND CC.Periodo = '" & Periodo_Contable & "' " _
                    & "AND TA.TP = '" & TFA.TC & "' " _
                    & "AND TA.Serie = '" & TFA.Serie & "' " _
                    & "AND TA.CodigoC = '" & TFA.CodigoC & "' " _
                    & "AND TA.Factura = " & TFA.Factura & " " _
                    & "AND TA.Fecha >= #" & BuscarFecha(TFA.Fecha) & "# " _
                    & "AND CC.Codigo = TA.Cta " _
                    & "AND CC.Item = TA.Item " _
                    & "AND CC.Periodo = TA.Periodo " _
                    & "ORDER BY CC.Codigo "
               Select_AdoDB AdoDBAux, sSQL
               If AdoDBAux.RecordCount > 0 Then
                  Do While Not AdoDBAux.EOF
                     TextoFormaPago = TextoFormaPago & AdoDBAux.fields("Fecha") & " " & AdoDBAux.fields("Banco") & " "
                     If AdoDBAux.fields("TC") = "BA" Then
                        TextoBanco = AdoDBAux.fields("Banco")
                        TextoCheque = AdoDBAux.fields("Cheque")
                     End If
                     AdoDBAux.MoveNext
                  Loop
               End If
               AdoDBAux.Close
               
               SaldoPendiente = 0
               If EsMatricula = False Then
                  sSQL = "SELECT CodigoC, SUM(Saldo_MN) As Pendiente " _
                       & "FROM Facturas " _
                       & "WHERE Item = '" & NumEmpresa & "' " _
                       & "AND Periodo = '" & Periodo_Contable & "' " _
                       & "AND CodigoC = '" & TFA.CodigoC & "' " _
                       & "AND TC = '" & TFA.TC & "' " _
                       & "AND T <> 'A' " _
                       & "GROUP BY CodigoC "
                  Select_AdoDB AdoDBAux, sSQL
                  If AdoDBAux.RecordCount > 0 Then SaldoPendiente = AdoDBAux.fields("Pendiente")
                  AdoDBAux.Close
               End If
               'MsgBox "."
               If SaldoPendiente <= 0 Then SaldoPendiente = .fields("Total_MN")
               Diferencia = SaldoPendiente - .fields("Total_MN")
               If Diferencia < 0 Then Diferencia = 0
               sSQL = "SELECT * " _
                    & "FROM Detalle_Factura " _
                    & "WHERE Item = '" & NumEmpresa & "' " _
                    & "AND Periodo = '" & Periodo_Contable & "' " _
                    & "AND Fecha >= #" & BuscarFecha(TFA.Fecha) & "# " _
                    & "AND TC = '" & TFA.TC & "' " _
                    & "AND Serie = '" & TFA.Serie & "' " _
                    & "AND Factura = " & TFA.Factura & " " _
                    & "ORDER BY ID, Codigo "
               Select_AdoDB AdoDBDet, sSQL
               MiForm.Caption = "Imprimiendo Factura No. " & Factura_No
               If (TFA.Hasta - TFA.Desde) <= 0 Then
                  If SetD(35).PosY > 0 Then PosLinea = PosLinea + SetD(35).PosY
                  If SetD(35).PosX > 0 Then PosColumna = PosColumna + SetD(35).PosX
               End If
               'MsgBox TFA.Factura & vbCrLf & PosColumna & vbCrLf & PosLinea & vbCrLf & "Pag. No. " & Pagina
               
               If (Pagina > TFA.CantFact) And (TFA.CantFact > 1) Then
                  'MsgBox "Pg. " & Pagina & vbCrLf & PosColumna & vbCrLf & PosLinea & vbCrLf & CantFact
                  Printer.NewPage
                  Pagina = 1: PosLinea = 0.01: PosColumna = 0.01
               End If
               Imprimir_FAM TFA, PosColumna, PosLinea, AdoDBFac, AdoDBDet, Cad_Tipo_Pago, ReImp, , CheqSinCodigo
               If CopiarComp Then
                  PosColumna = TFA.Pos_Factura
                  If (TFA.Hasta - TFA.Desde) <= 0 Then PosColumna = PosColumna + SetD(35).PosX
                  PosColumna = TFA.Pos_Factura
                  If TFA.Pos_Copia > 0 And CantFact = 1 Then PosLinea = TFA.Pos_Copia
                  'MsgBox PosLinea & vbCrLf & Factura_No
                  Imprimir_FAM TFA, PosColumna, PosLinea, AdoDBFac, AdoDBDet, Cad_Tipo_Pago, ReImp, , CheqSinCodigo
               End If
               AdoDBDet.Close
               
               PosColumna = 0.01
               Pagina = Pagina + 1
              'MsgBox Factura_No & vbCrLf & TFA.CantFact & vbCrLf & Pagina
               If TFA.CantFact = 1 Then
                  Printer.NewPage
                  Pagina = 1: PosLinea = 0.01
               Else
                  PosLinea = PosLinea + Salto_de_Factura
               End If
           End If
           
           sSQL = "UPDATE Facturas " _
                & "SET P = 1 " _
                & "WHERE Item = '" & NumEmpresa & "' " _
                & "AND Periodo = '" & Periodo_Contable & "' " _
                & "AND Factura = " & TFA.Factura & " " _
                & "AND TC = '" & TFA.TC & "' " _
                & "AND Serie = '" & TFA.Serie & "' " _
                & "AND Autorizacion = '" & TFA.Autorizacion & "' "
           Ejecutar_SQL_SP sSQL
           
          .MoveNext
        Loop
     End If
     MiForm.Caption = Cadena1
   End With
End If
MensajeEncabData = ""
Printer.EndDoc
RatonNormal
Exit Sub
Errorhandler:
    RatonNormal
    ErrorDeImpresion
    Exit Sub
End Sub

Public Sub Imprimir_Facturas_Copias_CxC(MiForm As Form, _
                                        DtaF As Adodc, _
                                        DtaD As Adodc, _
                                        FactDesde As Long, _
                                        FactHasta As Long, _
                                        TFA As Tipo_Facturas, _
                                        Optional ReImp As Boolean, _
                                        Optional EsMatricula As Boolean, _
                                        Optional PorOrdenFactura As Boolean)
Dim Posicion As Single
Dim Salto_de_Factura As Single
Dim Cont_Imp As Byte
Dim Cad_Tipo_Pago As String
On Error GoTo Errorhandler
RatonReloj
CEConLineas = ProcesarSeteos("FM")
Posicion = 0
Cont_Imp = 0
Cadena1 = MiForm.Caption
Mensajes = "Imprimir Copias de las Facturas" & vbCrLf & "Desde la " & FactDesde & " hasta la " & FactHasta
Titulo = "IMPRESION"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
   Escala_Centimetro Orientacion_Pagina, TipoArialNarrow, 9
   Pagina = 1: PosLinea = 0.01: PosColumna = 0.01
  'Iniciamos la impresion
   Printer.FontName = TipoArialNarrow
   Printer.FontBold = False
   sSQL = "SELECT F.*,C.Cliente,C.CI_RUC,C.Telefono,C.Direccion,C.DireccionT," _
        & "C.Grupo,C.Codigo,C.Ciudad,C.Email,C.TD " _
        & "FROM Facturas As F,Clientes As C " _
        & "WHERE F.Factura BETWEEN " & FactDesde & " and " & FactHasta & " " _
        & "AND F.CodigoC = C.Codigo " _
        & "AND F.TC = '" & TFA.TC & "' " _
        & "AND F.Serie = '" & TFA.Serie & "' " _
        & "AND F.Item = '" & NumEmpresa & "' " _
        & "AND F.Periodo = '" & Periodo_Contable & "' "
   If PorOrdenFactura Then
      sSQL = sSQL & "ORDER BY F.Factura,C.Grupo,C.Cliente "
   Else
      sSQL = sSQL & "ORDER BY C.Grupo,C.Cliente,F.Factura "
   End If
   Select_Adodc DtaF, sSQL
   'MsgBox sSQL
   With DtaF.Recordset
    If .RecordCount > 0 Then
        Validar_Porc_IVA .fields("Fecha")
        Cta_Cobrar = .fields("Cta_CxP")
        CodigoL = .fields("Cod_CxC")
        FA.Cod_CxC = .fields("Cod_CxC")
      ' Lineas de CxC Clientes
        Lineas_De_CxC FA
        Fecha_Vence = .fields("Vencimiento")
        Autorizacion = .fields("Autorizacion")
        SerieFactura = .fields("Serie")
      ' Espacios entre las Facturas
        Salto_de_Factura = TFA.AltoFactura + TFA.EspacioFactura
        If Salto_de_Factura <= 0 Then Salto_de_Factura = 0
      ' MsgBox CodigoL
        TFA.Autorizacion = Autorizacion
     Do While Not .EOF
        Cont_Imp = Cont_Imp + 1
        ReImp = CBool(.fields("P"))
        Factura_No = .fields("Factura")
        CodigoCliente = .fields("Codigo")
        SaldoPendiente = .fields("Saldo_MN")
           
        Cad_Tipo_Pago = Ninguno
        sSQL = "SELECT * " _
             & "FROM Tabla_Referenciales_SRI " _
             & "WHERE Tipo_Referencia = 'FORMA DE PAGO' " _
             & "AND Codigo = '" & TFA.Tipo_Pago & "' "
        Select_Adodc DtaD, sSQL
        If DtaD.Recordset.RecordCount > 0 Then Cad_Tipo_Pago = DtaD.Recordset.fields("Descripcion")
        
        If EsMatricula = False Then
           sSQL = "SELECT CodigoC, SUM(Saldo_MN) As Pendiente " _
                & "FROM Facturas " _
                & "WHERE Item = '" & NumEmpresa & "' " _
                & "AND Periodo = '" & Periodo_Contable & "' " _
                & "AND CodigoC = '" & CodigoCliente & "' " _
                & "AND TC = '" & TFA.TC & "' " _
                & "AND T <> 'A' " _
                & "GROUP BY CodigoC "
           Select_Adodc DtaD, sSQL
           If DtaD.Recordset.RecordCount > 0 Then SaldoPendiente = DtaD.Recordset.fields("Pendiente")
        End If
        If SaldoPendiente <= 0 Then SaldoPendiente = .fields("Total_MN")
        Diferencia = SaldoPendiente - .fields("Total_MN")
        If Diferencia < 0 Then Diferencia = 0
        sSQL = "SELECT * " _
             & "FROM Detalle_Factura " _
             & "WHERE Factura = " & Factura_No & " " _
             & "AND Item = '" & NumEmpresa & "' " _
             & "AND Periodo = '" & Periodo_Contable & "' " _
             & "AND TC = '" & TFA.TC & "' " _
             & "AND Serie = '" & TFA.Serie & "' " _
             & "ORDER BY Codigo "
        Select_Adodc DtaD, sSQL
        MiForm.Caption = "Imprimiendo Factura No. " & Factura_No
        
        Select Case Cont_Imp
          Case 1: PosLinea = 0.01: PosColumna = 0.01
                  Imprimir_FAM TFA, PosColumna, PosLinea, DtaF, DtaD, Cad_Tipo_Pago, ReImp, True
          Case 2: PosLinea = 0.01: PosColumna = TFA.Pos_Factura
                  Imprimir_FAM TFA, PosColumna, PosLinea, DtaF, DtaD, Cad_Tipo_Pago, ReImp, True
          Case 3: PosLinea = TFA.AltoFactura + TFA.EspacioFactura: PosColumna = 0.01
                  Imprimir_FAM TFA, PosColumna, PosLinea, DtaF, DtaD, Cad_Tipo_Pago, ReImp, True
          Case 4: PosLinea = TFA.AltoFactura + TFA.EspacioFactura: PosColumna = TFA.Pos_Factura
                  Imprimir_FAM TFA, PosColumna, PosLinea, DtaF, DtaD, Cad_Tipo_Pago, ReImp, True
                  Printer.NewPage
                  Pagina = Pagina + 1
                  Cont_Imp = 0
        End Select
       .MoveNext
     Loop
     End If
    MiForm.Caption = Cadena1
End With
End If
MensajeEncabData = ""
Printer.EndDoc
RatonNormal
Exit Sub
Errorhandler:
    RatonNormal
    ErrorDeImpresion
    Exit Sub
End Sub

Public Sub Imprimir_Recibos_CxC_PreFA(MiForm As Form, _
                                      DtaF As Adodc, _
                                      DtaD As Adodc, _
                                      FechaDesde As String, _
                                      FechaHasta As String, _
                                      Grupo1 As String, _
                                      Grupo2 As String, _
                                      TFA As Tipo_Facturas)
Dim Posicion As Single
Dim Salto_de_Factura As Single
Dim Cont_Imp As Byte
Dim LogoFacturaTemp As String
On Error GoTo Errorhandler
RatonReloj
CEConLineas = ProcesarSeteos("RB")
LogoFacturaTemp = TFA.LogoFactura
Posicion = 0
Cont_Imp = 0
Cadena1 = MiForm.Caption
Mensajes = "Imprimir Recibo: " & vbCrLf & "Desde la " & FechaDesde & " hasta la " & FechaHasta
Titulo = "IMPRESION"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
   Escala_Centimetro Orientacion_Pagina, TipoArialNarrow, 9
   Pagina = 1: PosLinea = 0.01: PosColumna = 0.01
  'Iniciamos la impresion
   Printer.FontName = TipoArialNarrow
   Printer.FontBold = False

   FechaIni = BuscarFecha(FechaDesde)
   FechaFin = BuscarFecha(FechaHasta)
   sSQL = "SELECT SUM(CF.Valor-CF.Descuento) As Total_MN,SUM(CF.Valor) As SaldoPend,SUM(CF.Descuento) as Descuento,0 As Factura,C.Cliente,C.CI_RUC,C.Telefono,C.Direccion," _
        & "C.Representante,C.Grupo,C.Codigo,C.Ciudad,C.CI_RUC_R,'" & FechaDesde & "' As Fecha,'" & FechaHasta & "' As Fecha_V,C.DireccionT,C.TD " _
        & "FROM Clientes_Facturacion As CF,Clientes As C " _
        & "WHERE CF.Item = '" & NumEmpresa & "' " _
        & "AND CF.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
        & "AND C.Grupo BETWEEN '" & Grupo1 & "' and '" & Grupo2 & "' " _
        & "AND C.Codigo = CF.Codigo " _
        & "GROUP BY C.Grupo,C.Cliente,C.CI_RUC,C.Telefono,C.Direccion,C.Representante,C.Codigo,C.Ciudad,C.CI_RUC_R,C.DireccionT,C.TD " _
        & "ORDER BY C.Grupo,C.Cliente,C.CI_RUC,C.Telefono,C.Direccion,C.Representante,C.Codigo,C.Ciudad,C.CI_RUC_R,C.DireccionT,C.TD "
   Select_Adodc DtaF, sSQL
   With DtaF.Recordset
    If .RecordCount > 0 Then
        Cta_Cobrar = Ninguno
        CodigoL = Ninguno
      ' Lineas de CxC Clientes
      ' Espacios entre las Facturas
        TFA.Pos_Factura = SetD(67).PosX
        TFA.AltoFactura = SetD(67).PosY
        TFA.EspacioFactura = SetD(68).PosY
        Salto_de_Factura = TFA.AltoFactura + TFA.EspacioFactura
        If Salto_de_Factura <= 0 Then Salto_de_Factura = 0
      ' MsgBox CodigoL

     Do While Not .EOF
        Cont_Imp = Cont_Imp + 1
        CodigoCliente = .fields("Codigo")
        SaldoPendiente = 0
        Diferencia = 0
        sSQL = "SELECT 1 As Cantidad,CF.Codigo_Inv As Codigo,CF.Valor As Precio,CP.Producto,CF.Mes," _
             & "CF.Valor As Total,'.' As Ticket " _
             & "FROM Clientes_Facturacion As CF, Catalogo_Productos As CP " _
             & "WHERE CF.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
             & "AND CF.Codigo = '" & CodigoCliente & "' " _
             & "AND CP.Item = '" & NumEmpresa & "' " _
             & "AND CP.Periodo = '" & Periodo_Contable & "' " _
             & "AND CF.Codigo_Inv = CP.Codigo_Inv " _
             & "AND CF.Item = CP.Item " _
             & "ORDER BY CF.Codigo_Inv "
        Select_Adodc DtaD, sSQL
        MiForm.Caption = "Imprimiendo Factura No. " & Factura_No
        TFA.LogoFactura = "RECIBOS"
        Select Case Cont_Imp
          Case 1: PosLinea = 0.01: PosColumna = 0.01
                  Imprimir_RCB PosColumna, PosLinea, DtaF, DtaD
          Case 2: PosLinea = 0.01: PosColumna = TFA.Pos_Factura
                  Imprimir_RCB PosColumna, PosLinea, DtaF, DtaD
          Case 3: PosLinea = TFA.AltoFactura + TFA.EspacioFactura: PosColumna = 0.01
                  Imprimir_RCB PosColumna, PosLinea, DtaF, DtaD
          Case 4: PosLinea = TFA.AltoFactura + TFA.EspacioFactura: PosColumna = TFA.Pos_Factura
                  Imprimir_RCB PosColumna, PosLinea, DtaF, DtaD
                  Printer.NewPage
                  Pagina = Pagina + 1
                  Cont_Imp = 0
        End Select
       .MoveNext
     Loop
     End If
    MiForm.Caption = Cadena1
End With
End If
MensajeEncabData = ""
TFA.LogoFactura = LogoFacturaTemp
Printer.EndDoc
RatonNormal
Exit Sub
Errorhandler:
    RatonNormal
    ErrorDeImpresion
    Exit Sub
End Sub

Public Sub Imprimir_Nota_Credito(MiForm As Form, _
                                 DtaF As Adodc, _
                                 DtaD As Adodc, _
                                 FactDesde As Long, _
                                 FactHasta As Long, _
                                 TFA As Tipo_Facturas, _
                                 Optional ReImp As Boolean, _
                                 Optional EsMatricula As Boolean)
Dim Posicion As Single
Dim Salto_de_Factura As Single
Dim PorcNC As Integer
On Error GoTo Errorhandler
RatonReloj
CEConLineas = ProcesarSeteos("NC")
Posicion = 0
Cadena1 = MiForm.Caption
Mensajes = "Imprimir las Facturas" & vbCrLf & "Desde la " & FactDesde & " hasta la " & FactHasta
Titulo = "IMPRESION"
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
   Escala_Centimetro Orientacion_Pagina, TipoArialNarrow, 9
   CopiarComp = False
   Mensajes = "Imprimir con copia"
   Titulo = "Pregunta de Impresion"
   If BoxMensaje = vbYes Then CopiarComp = True
   Pagina = 1: PosLinea = 0.01: PosColumna = 0.01
  'Iniciamos la impresion
   Printer.FontName = TipoArialNarrow
   Printer.FontBold = False
   
   sSQL = "UPDATE Facturas " _
        & "SET X = '.' " _
        & "WHERE Factura BETWEEN " & FactDesde & " and " & FactHasta & " " _
        & "AND TC = '" & TFA.TC & "' " _
        & "AND Serie = '" & TFA.Serie & "' " _
        & "AND Autorizacion = '" & TFA.Autorizacion & "' " _
        & "AND Item = '" & NumEmpresa & "' " _
        & "AND Periodo = '" & Periodo_Contable & "' "
   Ejecutar_SQL_SP sSQL
   
   If SQL_Server Then
      sSQL = "UPDATE Facturas " _
           & "SET X = 'C' " _
           & "FROM Facturas As F, Trans_Abonos As TA "
   Else
      sSQL = "UPDATE Facturas As F, Trans_Abonos As TA " _
           & "SET F.X = 'C' "
   End If
   sSQL = sSQL & "WHERE F.Factura BETWEEN " & FactDesde & " and " & FactHasta & " " _
        & "AND F.TC = '" & TFA.TC & "' " _
        & "AND F.Serie = '" & TFA.Serie & "' " _
        & "AND F.Autorizacion = '" & TFA.Autorizacion & "' " _
        & "AND F.Item = '" & NumEmpresa & "' " _
        & "AND F.Periodo = '" & Periodo_Contable & "' " _
        & "AND TA.Banco = 'NOTA DE CREDITO' " _
        & "AND F.TC = TA.TP " _
        & "AND F.Serie = TA.Serie " _
        & "AND F.Autorizacion = TA.Autorizacion " _
        & "AND F.Item = TA.Item " _
        & "AND F.Periodo = TA.Periodo "
   Ejecutar_SQL_SP sSQL
      
   sSQL = "SELECT F.*,C.Cliente,C.CI_RUC,C.Telefono,C.Direccion,C.DireccionT,C.Representante," _
        & "C.Grupo,C.Codigo,C.Ciudad,C.Email,C.Cedula,C.TD " _
        & "FROM Facturas As F,Clientes As C " _
        & "WHERE F.Factura BETWEEN " & FactDesde & " and " & FactHasta & " " _
        & "AND F.X = 'C' " _
        & "AND F.TC = '" & TFA.TC & "' " _
        & "AND F.Serie = '" & TFA.Serie & "' " _
        & "AND F.Autorizacion = '" & TFA.Autorizacion & "' " _
        & "AND F.Item = '" & NumEmpresa & "' " _
        & "AND F.Periodo = '" & Periodo_Contable & "' " _
        & "AND C.Codigo = F.CodigoC " _
        & "ORDER BY C.Grupo,C.Cliente,F.Factura "
   Select_Adodc DtaF, sSQL
  'MsgBox sSQL
   PorcNC = 0
   With DtaF.Recordset
    If .RecordCount > 0 Then
       'MsgBox FactDesde & "     " & FactHasta & "     " & .RecordCount
        Cta_Cobrar = .fields("Cta_CxP")
        CodigoL = .fields("Cod_CxC")
        TFA.Cod_CxC = .fields("Cod_CxC")
      ' Espacios entre las Facturas
        TFA.Pos_Factura = SetD(78).PosX
        TFA.AnchoFactura = SetD(79).PosX
        TFA.AltoFactura = SetD(79).PosY
        Salto_de_Factura = TFA.AltoFactura + TFA.EspacioFactura
        If Salto_de_Factura <= 0 Then Salto_de_Factura = 0
     Do While Not .EOF
        TFA.Nivel = .fields("Grupo")
        Factura_No = .fields("Factura")
        CodigoCliente = .fields("Codigo")
        SaldoPendiente = .fields("Saldo_MN")
        If SaldoPendiente <= 0 Then SaldoPendiente = .fields("Total_MN")
        Diferencia = SaldoPendiente - .fields("Total_MN")
        If Diferencia < 0 Then Diferencia = 0
        IVA_NC = 0
        SubTotal_NC = 0
        sSQL = "SELECT Fecha,Serie_NC,Cheque,SUM(Abono) As Abonos " _
             & "FROM Trans_Abonos " _
             & "WHERE Factura = " & Factura_No & " " _
             & "AND Item = '" & NumEmpresa & "' " _
             & "AND Periodo = '" & Periodo_Contable & "' " _
             & "AND TP = '" & TFA.TC & "' " _
             & "AND Serie = '" & TFA.Serie & "' " _
             & "AND Autorizacion = '" & TFA.Autorizacion & "' " _
             & "AND Banco = 'NOTA DE CREDITO' " _
             & "GROUP BY Fecha,Serie_NC, Cheque " _
             & "ORDER BY Fecha,Serie_NC, Cheque "
        Select_Adodc DtaD, sSQL
        'MsgBox DtaD.Recordset.RecordCount
        If DtaD.Recordset.RecordCount > 0 Then
           NC.Serie_NC = DtaD.Recordset.fields("Serie_NC")
           NC.Fecha_NC = DtaD.Recordset.fields("Fecha")
           Lineas_De_CxC NC
          'Lineas_De_NC DtaD.Recordset.Fields("Serie_NC"), DtaD.Recordset.Fields("Fecha")
           Do While Not DtaD.Recordset.EOF
              If DtaD.Recordset.fields("Cheque") = "VENTAS" Then SubTotal_NC = DtaD.Recordset.fields("Abonos")
              If DtaD.Recordset.fields("Cheque") = "I.V.A." Then IVA_NC = DtaD.Recordset.fields("Abonos")
              DtaD.Recordset.MoveNext
           Loop
        End If
        sSQL = "SELECT Fecha,Cheque,SUM(Abono) As Abonos " _
             & "FROM Trans_Abonos " _
             & "WHERE Factura = " & Factura_No & " " _
             & "AND Item = '" & NumEmpresa & "' " _
             & "AND Periodo = '" & Periodo_Contable & "' " _
             & "AND TP = '" & TFA.TC & "' " _
             & "AND Serie = '" & TFA.Serie & "' " _
             & "AND Autorizacion = '" & TFA.Autorizacion & "' " _
             & "AND Banco = 'NOTA DE CREDITO' " _
             & "GROUP BY Fecha,Cheque " _
             & "ORDER BY Fecha,Cheque "
        Select_Adodc DtaD, sSQL
        'MsgBox DtaD.Recordset.RecordCount
        If DtaD.Recordset.RecordCount > 0 Then
           Do While Not DtaD.Recordset.EOF
              Mifecha = DtaD.Recordset.fields("Fecha")
              DtaD.Recordset.MoveNext
           Loop
        End If
        sSQL = "SELECT * " _
             & "FROM Detalle_Factura " _
             & "WHERE Factura = " & Factura_No & " " _
             & "AND Item = '" & NumEmpresa & "' " _
             & "AND Periodo = '" & Periodo_Contable & "' " _
             & "AND TC = '" & TFA.TC & "' " _
             & "AND Serie = '" & TFA.Serie & "' " _
             & "AND Autorizacion = '" & TFA.Autorizacion & "' " _
             & "ORDER BY Codigo "
        Select_Adodc DtaD, sSQL
        'MsgBox "..."
        MiForm.Caption = Format$(PorcNC / .RecordCount, "00%") & " Imprimiendo " & TFA.Nivel & " Factura No. " & Factura_No
        PorcNC = PorcNC + 1
        If (FactHasta - FactDesde) <= 0 Then
           If SetD(35).PosY > 0 Then PosLinea = PosLinea + SetD(35).PosY
           If SetD(35).PosX > 0 Then PosColumna = PosColumna + SetD(35).PosX
        End If
       'MsgBox IVA_NC + SubTotal_NC
        If (IVA_NC + SubTotal_NC) > 0 Then
           Imprimir_FNC PosColumna, PosLinea, DtaF, DtaD, ReImp
           If CopiarComp Then
              PosColumna = TFA.Pos_Factura
              If (FactHasta - FactDesde) <= 0 Then PosColumna = PosColumna + SetD(35).PosX
              PosColumna = TFA.Pos_Factura
              If TFA.Pos_Copia > 0 And CantFact = 1 Then PosLinea = TFA.Pos_Copia
             'MsgBox PosLinea & vbCrLf & Factura_No
              Imprimir_FNC PosColumna, PosLinea, DtaF, DtaD, ReImp
           End If
           PosLinea = PosLinea + TFA.Pos_Copia + Salto_de_Factura
           PosColumna = 0.01
          'MsgBox Pagina & vbCrLf & Factura_No
           Pagina = Pagina + 1
        End If
        If (Pagina > CantFact) And (CantFact > 1) Then
           Printer.NewPage
           Pagina = 1: PosLinea = 0.01
        End If
        If CantFact = 1 Then
           Printer.NewPage
           Pagina = 1: PosLinea = 0.01
        End If
       .MoveNext
     Loop
     End If
    MiForm.Caption = Cadena1
End With
End If
MensajeEncabData = ""
Printer.EndDoc
RatonNormal
Exit Sub
Errorhandler:
    RatonNormal
    ErrorDeImpresion
    Exit Sub
End Sub

Public Sub Imprimir_Facturas(TFA As Tipo_Facturas)
Dim AdoDBDetalle As ADODB.Recordset
Dim CadenaMoneda As String
Dim Numero_Letras As String
Dim NombUusuario As String
Dim Cad_Tipo_Pago As String
Dim PVP_Desc As Currency
Dim Orden_No_S As String
Dim PFilT As Single
Dim PFilTemp As Single
Dim Desc_Sin_IVA As Currency
Dim Desc_Con_IVA As Currency

'Establecemos Espacios y seteos de impresion
On Error GoTo Errorhandler
  'MsgBox TipoFact
   CEConLineas = ProcesarSeteos(TFA.TC)
CantFils = 0
Mensajes = "Imprmir Factura No. " & TFA.Factura
Titulo = "IMPRESION"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
   
   tPrint.TipoImpresion = Es_Printer
   tPrint.NombreArchivo = TFA.TC & "-" & TFA.Serie & "-" & Format(TFA.Factura, "000000000")
   tPrint.TituloArchivo = TFA.TC & "-" & TFA.Serie & "-" & Format(TFA.Factura, "000000000")
   tPrint.TipoLetra = TipoCourier ' TipoArialNarrow
   tPrint.OrientacionPagina = Orientacion_Pagina
   tPrint.PaginaA4 = True
   tPrint.EsCampoCorto = False
   tPrint.VerDocumento = True
   Set cPrint = New cImpresion
   cPrint.iniciaImpresion
   RatonReloj
   Orden_No_S = ""
   cPrint.tipoDeLetra = TipoCourier ' TipoArialNarrow
   cPrint.tipoNegrilla = True
   Leer_Datos_FA_NV TFA
   Set AdoDBDetalle = Leer_Datos_FA_NV_Detalle(TFA)
   With AdoDBDetalle
    If .RecordCount > 0 Then
        Orden_No = .fields("Orden_No")
        Do While (Not .EOF)
           If Val(Orden_No) <> Val(TFA.Orden_Compra) And Val(TFA.Orden_Compra) > 0 Then
              Orden_No_S = Orden_No_S & CStr(Orden_No) & " "
              Orden_No = TFA.Orden_Compra
           End If
          .MoveNext
        Loop
        Orden_No_S = Orden_No_S & CStr(Orden_No) & " "
       .MoveFirst
    End If
   End With
  'MsgBox FA.Si_Existe_Doc
   If FA.Si_Existe_Doc Then
     '--------------------------------------------------------------------------------------------------------
     'Encabezado de la factura
     '--------------------------------------------------------------------------------------------------------
     'Imagen de la factura
      If TFA.LogoFactura <> "NINGUNO" And TFA.AnchoFactura > 0 And TFA.AltoFactura > 0 Then
         If SetD(1).PosX > 0 And SetD(1).PosY > 0 Then
            Codigo4 = Format$(TFA.Factura, "000000000")
            If TFA.LogoFactura = "MATRICIA" Then
               Imprimir_Formato_Propio "IF", SetD(1).PosX, SetD(1).PosY
            Else
               Cadena = RutaSistema & "\FORMATOS\" & TFA.LogoFactura & ".gif"
               cPrint.printImagen Cadena, SetD(1).PosX, SetD(1).PosY, TFA.AnchoFactura, TFA.AltoFactura
               cPrint.printImagen LogoTipo, SetD(34).PosX, SetD(34).PosY, 2.5, 1.25
            End If
         End If
      End If
      If SetD(2).PosX > 0 And SetD(2).PosY > 0 Then
         cPrint.PorteDeLetra = SetD(2).Porte
         cPrint.printTexto SetD(2).PosX, SetD(2).PosY, TFA.Serie & "-" & Format$(TFA.Factura, "000000000")
      End If
      If SetD(8).PosX > 0 And SetD(8).PosY > 0 Then
         cPrint.PorteDeLetra = SetD(8).Porte
         If TFA.Razon_Social = TFA.Cliente Then Cadena = TFA.Cliente Else Cadena = TFA.Razon_Social
         cPrint.printTexto SetD(8).PosX, SetD(8).PosY, Cadena
      End If
      If SetD(64).PosX > 0 And SetD(64).PosY > 0 Then
         cPrint.PorteDeLetra = SetD(64).Porte
         If TFA.Razon_Social <> TFA.Cliente Then
            cPrint.printTexto SetD(64).PosX, SetD(64).PosY, TFA.Cliente
         End If
      End If
      If SetD(11).PosX > 0 And SetD(11).PosY > 0 Then
         DireccionCli = TFA.DireccionC
         If Len(TFA.DirNumero) > 1 And TFA.DirNumero <> Ninguno Then
            If TFA.DirNumero <> "S/N" Then DireccionCli = DireccionCli & " (" & TFA.DirNumero & ")"
         End If
         cPrint.PorteDeLetra = SetD(11).Porte
         cPrint.printTexto SetD(11).PosX, SetD(11).PosY, DireccionCli
      End If
      If SetD(10).PosX > 0 And SetD(10).PosY > 0 Then
         cPrint.PorteDeLetra = SetD(10).Porte
         cPrint.printTexto SetD(10).PosX, SetD(10).PosY, TFA.Grupo
      End If
     'Codigo abreviado del Usuario
      If SetD(18).PosX > 0 And SetD(18).PosY > 0 Then
         NombUusuario = TFA.Digitador
         Cadena = Cambio_Usuario_Inicial(NombUusuario)
         cPrint.PorteDeLetra = SetD(18).Porte
         cPrint.printTexto SetD(18).PosX, SetD(18).PosY, Cadena
      End If
      If SetD(21).PosX > 0 And SetD(21).PosY > 0 Then
         cPrint.PorteDeLetra = SetD(21).Porte
         cPrint.printTexto SetD(21).PosX, SetD(21).PosY, TFA.Ejecutivo_Venta
      End If
      If SetD(3).PosX > 0 And SetD(3).PosY > 0 Then
         cPrint.PorteDeLetra = SetD(3).Porte
         cPrint.printTexto SetD(3).PosX, SetD(3).PosY, FechaStrgCorta(TFA.Fecha)
      End If
      If SetD(4).PosX > 0 And SetD(4).PosY > 0 Then
         Cadena = FechaDia(TFA.Fecha) & Space(10) & FechaMes(TFA.Fecha) & Space(10) & FechaAnio(TFA.Fecha)
         cPrint.PorteDeLetra = SetD(4).Porte
         cPrint.printTexto SetD(4).PosX, SetD(4).PosY, Cadena
      End If
      If SetD(7).PosX > 0 And SetD(7).PosY > 0 Then
         cPrint.PorteDeLetra = SetD(7).Porte
         cPrint.printTexto SetD(7).PosX, SetD(7).PosY, FechaStrgCiudad(TFA.Fecha)
      End If
      If SetD(5).PosX > 0 And SetD(5).PosY > 0 Then
         cPrint.PorteDeLetra = SetD(5).Porte
         cPrint.printTexto SetD(5).PosX, SetD(5).PosY, FechaStrgCorta(TFA.Fecha_V)
      End If
      If SetD(6).PosX > 0 And SetD(6).PosY > 0 Then
         Cadena = FechaDia(TFA.Fecha_V) & Space(10) & FechaMes(TFA.Fecha_V) & Space(10) & FechaAnio(TFA.Fecha_V)
         cPrint.PorteDeLetra = SetD(6).Porte
         cPrint.printTexto SetD(6).PosX, SetD(6).PosY, Cadena
      End If
      If SetD(14).PosX > 0 And SetD(14).PosY > 0 Then
         cPrint.PorteDeLetra = SetD(14).Porte
         cPrint.printTexto SetD(14).PosX, SetD(14).PosY, TFA.TelefonoC
      End If
      If SetD(12).PosX > 0 And SetD(12).PosY > 0 Then
         cPrint.PorteDeLetra = SetD(12).Porte
         cPrint.printTexto SetD(12).PosX, SetD(12).PosY, TFA.CiudadC
      End If
      If SetD(13).PosX > 0 And SetD(13).PosY > 0 Then
         cPrint.PorteDeLetra = SetD(13).Porte
         cPrint.printTexto SetD(13).PosX, SetD(13).PosY, TFA.CI_RUC
      End If
      If SetD(15).PosX > 0 And SetD(15).PosY > 0 Then
         cPrint.PorteDeLetra = SetD(15).Porte
         cPrint.printTexto SetD(15).PosX, SetD(15).PosY, TFA.EmailC
      End If
      If SetD(51).PosX > 0 And SetD(71).PosY > 0 Then
         cPrint.PorteDeLetra = SetD(51).Porte
         cPrint.printTexto SetD(51).PosX, SetD(51).PosY, TFA.DAU
      End If
      If SetD(52).PosX > 0 And SetD(52).PosY > 0 Then
         cPrint.PorteDeLetra = SetD(52).Porte
         cPrint.printTexto SetD(52).PosX, SetD(52).PosY, TFA.FUE
      End If
      If SetD(53).PosX > 0 And SetD(53).PosY > 0 Then
         cPrint.PorteDeLetra = SetD(53).Porte
         cPrint.printTexto SetD(53).PosX, SetD(53).PosY, TFA.Declaracion
      End If
      If SetD(54).PosX > 0 And SetD(54).PosY > 0 Then
         cPrint.PorteDeLetra = SetD(54).Porte
         cPrint.printTexto SetD(54).PosX, SetD(54).PosY, TFA.Remision
      End If
      If SetD(55).PosX > 0 And SetD(55).PosY > 0 Then
         cPrint.PorteDeLetra = SetD(55).Porte
         cPrint.printTexto SetD(55).PosX, SetD(55).PosY, TFA.Comercial
      End If
      If SetD(56).PosX > 0 And SetD(56).PosY > 0 Then
        cPrint.PorteDeLetra = SetD(56).Porte
        cPrint.printTexto SetD(56).PosX, SetD(56).PosY, TFA.Solicitud
      End If
      If SetD(57).PosX > 0 And SetD(57).PosY > 0 Then
        cPrint.PorteDeLetra = SetD(57).Porte
        cPrint.printTexto SetD(57).PosX, SetD(57).PosY, TFA.Cantidad
      End If
      If SetD(58).PosX > 0 And SetD(58).PosY > 0 Then
         cPrint.PorteDeLetra = SetD(58).Porte
         cPrint.printTexto SetD(58).PosX, SetD(58).PosY, TFA.Kilos
      End If
      If SetD(60).PosX > 0 And SetD(60).PosY > 0 Then
         cPrint.PorteDeLetra = SetD(60).Porte
         cPrint.printTexto SetD(60).PosX, SetD(60).PosY, Format$(Day(TFA.Fecha), "00")
      End If
      If SetD(61).PosX > 0 And SetD(61).PosY > 0 Then
         Cadena = UCaseStrg(MidStrg(MesesLetras(CInt(Month(TFA.Fecha))), 1, 3))
         cPrint.PorteDeLetra = SetD(61).Porte
         cPrint.printTexto SetD(61).PosX, SetD(61).PosY, Cadena
      End If
      If SetD(62).PosX > 0 And SetD(62).PosY > 0 Then
         cPrint.PorteDeLetra = SetD(62).Porte
         cPrint.printTexto SetD(62).PosX, SetD(62).PosY, Format$(Year(TFA.Fecha), "0000")
      End If
      If SetD(67).PosX > 0 And SetD(67).PosY > 0 Then
        cPrint.PorteDeLetra = SetD(67).Porte
        cPrint.printTexto SetD(67).PosX, SetD(67).PosY, FechaStrgCorta(TFA.Fecha_C)
      End If
      If SetD(16).PosX > 0 And SetD(16).PosY > 0 Then
        cPrint.PorteDeLetra = SetD(16).Porte
        NumeroLineas = cPrint.printTextoMultiple(SetD(16).PosX, SetD(16).PosY, TFA.Observacion, SetD(26).PosX)
      End If
      If SetD(17).PosX > 0 And SetD(17).PosY > 0 Then
        cPrint.PorteDeLetra = SetD(17).Porte
        NumeroLineas = cPrint.printTextoMultiple(SetD(17).PosX, SetD(17).PosY, TFA.Nota, SetD(26).PosX)
      End If
     '--------------------------------------------------------------------------------------------------------
     'Pie de factura
     '--------------------------------------------------------------------------------------------------------
      cPrint.tipoNegrilla = True
      Total = Redondear(TFA.Total_MN, 2)
      Total_ME = Redondear(TFA.Total_ME, 2)
     'Porcentaje Con IVA
      If SetD(39).PosX > 0 And SetD(39).PosY > 0 Then
         cPrint.PorteDeLetra = SetD(39).Porte
         cPrint.printTexto SetD(39).PosX, SetD(39).PosY, CStr(TFA.Porc_IVA * 100)
      End If
     'Sin IVA
      If SetD(37).PosX > 0 And SetD(37).PosY > 0 Then
         cPrint.PorteDeLetra = SetD(37).Porte
         Diferencia = TFA.Sin_IVA - TFA.Descuento_0
         cPrint.printVariable SetD(37).PosX, SetD(37).PosY, Diferencia
      End If
     'Con IVA
      If SetD(38).PosX > 0 And SetD(38).PosY > 0 Then
         cPrint.PorteDeLetra = SetD(38).Porte
         Diferencia = TFA.Con_IVA - TFA.Descuento_X
         cPrint.printVariable SetD(38).PosX, SetD(38).PosY, Diferencia
      End If
     'Descuento palabra
      If SetD(43).PosX > 0 And SetD(43).PosY > 0 Then
         cPrint.PorteDeLetra = SetD(43).Porte
         cPrint.printTexto SetD(43).PosX, SetD(43).PosY, "Descuento"
      End If
     'Total Descuento
      If SetD(42).PosX > 0 And SetD(42).PosY > 0 Then
         cPrint.PorteDeLetra = SetD(42).Porte
         cPrint.printVariable SetD(42).PosX, SetD(42).PosY, TFA.Total_Descuento
      End If
     'Total Comision
      If SetD(48).PosX > 0 And SetD(48).PosY > 0 Then
         cPrint.PorteDeLetra = SetD(48).Porte
         cPrint.printVariable SetD(48).PosX, SetD(48).PosY, TFA.Comision
      End If
     'Total Servicio
      If SetD(49).PosX > 0 And SetD(49).PosY > 0 Then
         cPrint.PorteDeLetra = SetD(49).Porte
         cPrint.printVariable SetD(49).PosX, SetD(49).PosY, TFA.Servicio
      End If
     'IVA Porcentaje
      If SetD(41).PosX > 0 And SetD(41).PosY > 0 Then
         cPrint.PorteDeLetra = SetD(41).Porte
         cPrint.printTexto SetD(41).PosX, SetD(41).PosY, CStr(TFA.Porc_IVA * 100) & " "
      End If
     'Total IVA
      If SetD(40).PosX > 0 And SetD(40).PosY > 0 Then
         cPrint.PorteDeLetra = SetD(40).Porte
         cPrint.printVariable SetD(40).PosX, SetD(40).PosY, TFA.Total_IVA
      End If
     'SubTotal
      If SetD(36).PosX > 0 And SetD(36).PosY > 0 Then
         cPrint.PorteDeLetra = SetD(36).Porte
         cPrint.printVariable SetD(36).PosX, SetD(36).PosY, TFA.SubTotal
      End If
     'SubTotal - Descuentos
      If SetD(66).PosX > 0 And SetD(66).PosY > 0 Then
         cPrint.PorteDeLetra = SetD(66).Porte
         cPrint.printVariable SetD(66).PosX, SetD(66).PosY, TFA.SubTotal - TFA.Total_Descuento
      End If
     'Total Facturado
      If SetD(44).PosX > 0 And SetD(44).PosY > 0 Then
         cPrint.PorteDeLetra = SetD(44).Porte
         cPrint.printVariable SetD(44).PosX, SetD(44).PosY, TFA.Total_MN
      End If
     'Total Facturado en letras
      If SetD(45).PosX > 0 And SetD(45).PosY > 0 Then
         Numero_Letras = Cambio_Letras(TFA.Total_MN, 2)
         cPrint.PorteDeLetra = SetD(45).Porte
         PrinterLineas SetD(45).PosX, SetD(45).PosY, Numero_Letras, 11.5
      End If
     'CxC Clientes: Linea de Produccion
      If SetD(50).PosX > 0 And SetD(50).PosY > 0 Then
         cPrint.PorteDeLetra = SetD(50).Porte
         cPrint.printTexto SetD(50).PosX, SetD(50).PosY, TFA.CxC_Clientes
      End If
     'Hora de Proceso
      If SetD(63).PosX > 0 And SetD(63).PosY > 0 Then
         cPrint.PorteDeLetra = SetD(63).Porte
         cPrint.printTexto SetD(63).PosX, SetD(63).PosY, TFA.Hora
      End If
      If SetD(68).PosX > 0 And SetD(68).PosY > 0 Then
         cPrint.PorteDeLetra = SetD(68).Porte
         cPrint.printTexto SetD(68).PosX, SetD(68).PosY, TrimStrg(Orden_No_S)
      End If
     'Tipo_Pago
      If Len(TFA.Tipo_Pago_Det) > 1 Then
         If SetD(79).PosX > 0 And SetD(79).PosY > 0 Then
            cPrint.PorteDeLetra = SetD(79).Porte
            cPrint.printVariable SetD(79).PosX, SetD(79).PosY, TFA.Fecha_V
         End If
         If SetD(78).PosX > 0 And SetD(78).PosY > 0 Then
            cPrint.PorteDeLetra = SetD(78).Porte
            cPrint.printVariable SetD(78).PosX, SetD(78).PosY, TFA.Total_MN
         End If
         If SetD(77).PosX > 0 And SetD(77).PosY > 0 Then
            RutaOrigen = RutaSistema & "\FORMATOS\Vistofp.jpg"
            cPrint.printImagen RutaOrigen, SetD(77).PosX, SetD(77).PosY, SetD(77).Porte, SetD(77).Porte
         End If
         If SetD(76).PosX > 0 And SetD(76).PosY > 0 Then
            cPrint.PorteDeLetra = SetD(76).Porte
            cPrint.printTexto SetD(76).PosX, SetD(76).PosY, "TIPO PAGO:"
         End If
         If SetD(75).PosX > 0 And SetD(75).PosY > 0 Then
            cPrint.PorteDeLetra = SetD(75).Porte
            cPrint.printTexto SetD(75).PosX, SetD(75).PosY, TrimStrg(MidStrg(TFA.Tipo_Pago_Det, 15, Len(TFA.Tipo_Pago_Det)))
         End If
      End If
     'MsgBox Orden_No_S & " ...."
   End If
  'Printer.FontName = TipoConsola
   SaldoInic = 0: SaldoFinal = 0
   Orden_No_S = ""
   cPrint.tipoNegrilla = False
  'Comenzamos a recoger los detalles de la factura
   sSQL = "SELECT DF.*,CP.Detalle,CP.Codigo_Barra,CP.Unidad,CP.Reg_Sanitario,CM.Marca " _
        & "FROM Detalle_Factura As DF,Catalogo_Productos As CP,Catalogo_Marcas As CM " _
        & "WHERE DF.Factura = " & TFA.Factura & " " _
        & "AND DF.TC = '" & TFA.TC & "' " _
        & "AND DF.Serie = '" & TFA.Serie & "' " _
        & "AND DF.Item = '" & NumEmpresa & "' " _
        & "AND DF.Periodo = '" & Periodo_Contable & "' " _
        & "AND DF.Periodo = CP.Periodo " _
        & "AND DF.Periodo = CM.Periodo " _
        & "AND DF.Item = CP.Item " _
        & "AND DF.Item = CM.Item " _
        & "AND DF.Codigo = CP.Codigo_Inv " _
        & "AND DF.CodMarca = CM.CodMar " _
        & "ORDER BY DF.ID,DF.Codigo "
   Select_AdoDB AdoDBDetalle, sSQL
   With AdoDBDetalle
    If .RecordCount > 0 Then
        PFil = SetD(22).PosY
        Orden_No = .fields("Orden_No")
        Do While (Not .EOF)
          'MsgBox .RecordCount
           SaldoInic = SaldoInic + .fields("Cantidad")
           SaldoFinal = SaldoFinal + .fields("Tonelaje")
           cPrint.PorteDeLetra = SetD(23).Porte
           cPrint.printFields SetD(23).PosX, PFil, .fields("Codigo")
           cPrint.PorteDeLetra = SetD(30).Porte
           cPrint.printFields SetD(30).PosX, PFil, .fields("Codigo_Barra")
           If .fields("CodMarca") <> Ninguno Then
               cPrint.PorteDeLetra = SetD(31).Porte
               cPrint.printTexto SetD(31).PosX, PFil, "(" & TrimStrg(MidStrg(.fields("Marca"), 1, 3)) & ")"
           End If
           cPrint.PorteDeLetra = SetD(33).Porte
           cPrint.printFields SetD(33).PosX, PFil, .fields("Ruta")
           cPrint.PorteDeLetra = SetD(32).Porte
           cPrint.printFields SetD(32).PosX, PFil, .fields("Unidad")
           If .fields("Cant_Hab") > 0 And Len(.fields("Tipo_Hab")) > 1 Then
               cPrint.PorteDeLetra = SetD(70).Porte
               cPrint.printFields SetD(70).PosX, PFil, .fields("Fecha_IN")
               cPrint.PorteDeLetra = SetD(71).Porte
               cPrint.printFields SetD(71).PosX, PFil, .fields("Fecha_OUT")
               cPrint.PorteDeLetra = SetD(72).Porte
               cPrint.printFields SetD(72).PosX, PFil, .fields("Cant_Hab")
               cPrint.PorteDeLetra = SetD(73).Porte
               cPrint.printFields SetD(73).PosX, PFil, .fields("Tipo_Hab")
           End If
           If SetD(24).PosX <= SetD(25).PosX Then
              cPrint.PorteDeLetra = SetD(24).Porte
              cPrint.printTexto SetD(24).PosX, PFil, " " & CStr(.fields("Cantidad"))
           End If
           cPrint.PorteDeLetra = SetD(25).Porte
           PFilTemp = PFil
           PFilT = PrinterLineasTexto(SetD(25).PosX, PFil - 0.2, .fields("Producto"), SetD(26).PosX)
           If PVP_Al_Inicio Then PFil = PFilTemp Else PFil = PFilT
           If Len(.fields("Lote_No")) > 1 And Len(.fields("Reg_Sanitario")) > 1 Then
              PFil = PFil + Printer.TextHeight("H")
              Cadena = "-LOTE No. " & .fields("Lote_No") _
                     & ", REG. SANITARIO: " & .fields("Reg_Sanitario")
              'cPrint.printTexto SetD(25).PosX, PFil, Cadena
              PFilT = PrinterLineasTexto(SetD(25).PosX, PFil, Cadena, SetD(26).PosX)
              PFil = PFil + Printer.TextHeight("H")
              Cadena = UCaseStrg("-FAB: " & Format(.fields("Fecha_Fab"), "MMM-yyyy") & ", EXP: " & Format(.fields("Fecha_Exp"), "MMM-yyyy"))
              'cPrint.printTexto SetD(25).PosX, PFil, Cadena
              PFilT = PrinterLineasTexto(SetD(25).PosX, PFil, Cadena, SetD(26).PosX)
           End If
           PFil = PFil + 0.2
           If SetD(24).PosX > SetD(25).PosX Then
              cPrint.PorteDeLetra = SetD(24).Porte
              cPrint.printTexto SetD(24).PosX, PFil, " " & CStr(.fields("Cantidad"))
           End If
           cPrint.PorteDeLetra = SetD(29).Porte
           cPrint.printFields SetD(29).PosX, PFil, .fields("Total")
           PVP_Desc = .fields("Total") - .fields("Total_Desc") - .fields("Total_Desc2")
           cPrint.PorteDeLetra = SetD(69).Porte
           cPrint.printVariable SetD(69).PosX, PFil, PVP_Desc
           cPrint.PorteDeLetra = SetD(28).Porte
           cPrint.printFields SetD(28).PosX, PFil, .fields("Precio"), , , , Dec_PVP
           cPrint.PorteDeLetra = SetD(27).Porte
           If .fields("Precio") > 0 And .fields("Cantidad") > 0 Then
               cPrint.printTexto SetD(27).PosX, PFil, Format$(.fields("Total_Desc") / (.fields("Cantidad") * .fields("Precio")), "00.00%")
           End If
           cPrint.PorteDeLetra = SetD(34).Porte
           cPrint.printTexto SetD(34).PosX, PFil, Format$(.fields("Tonelaje"), "#,##0.00")
          'Descuentos en Ventas
           If .fields("Total_IVA") > 0 Then
               Desc_Con_IVA = Desc_Con_IVA + .fields("Total_Desc")
           Else
               Desc_Sin_IVA = Desc_Sin_IVA + .fields("Total_Desc")
           End If
           If PVP_Al_Inicio Then PFil = PFilT
           PFil = PFil + Printer.TextHeight("H")
           If Val(Orden_No) <> Val(TFA.Orden_Compra) And Val(TFA.Orden_Compra) > 0 Then
              Orden_No_S = Orden_No_S & CStr(Orden_No) & " "
              Orden_No = TFA.Orden_Compra
           End If
          .MoveNext
        Loop
        Orden_No_S = Orden_No_S & CStr(Orden_No) & " "
    End If
   End With
   cPrint.finalizaImpresion
End If
'Printer.FontName = LetraAnterior
MensajeEncabData = ""
RatonNormal
Exit Sub
Errorhandler:
    RatonNormal
    ErrorDeImpresion
    Exit Sub
End Sub

Public Function Lista_Facturas_Text(TFA As Tipo_Facturas) As String
Dim AdoFacturaDB As ADODB.Recordset
Dim AdoDetalleDB As ADODB.Recordset
Dim TextDoc As String
Dim NombUsuario As String
Dim NombEjecutivo As String

Dim CadenaMoneda As String
Dim Numero_Letras As String
Dim Cad_Tipo_Pago As String
Dim Desc_Sin_IVA As Currency
Dim Desc_Con_IVA As Currency
Dim PVP_Desc As Currency
Dim Orden_No_S As String
Dim CxC_Clientes As String
Dim Imp_Mes As Boolean

CantFils = 0
TextDoc = ""
If TFA.Factura > 0 And Len(TFA.Serie) = 6 And Len(TFA.Autorizacion) > 3 Then
   RatonReloj
   CxC_Clientes = Ninguno
   NombEjecutivo = Ninguno
   sSQL = "SELECT F.*,C.Cliente,C.CI_RUC,C.Telefono,C.Direccion,C.Grupo,C.Codigo," _
        & "C.Ciudad,C.Email,C.DirNumero,C.Fecha As Fecha_A " _
        & "FROM Facturas As F,Clientes As C " _
        & "WHERE F.Factura = " & TFA.Factura & " " _
        & "AND F.TC = '" & TFA.TC & "' " _
        & "AND F.Serie = '" & TFA.Serie & "' " _
        & "AND F.Periodo = '" & Periodo_Contable & "' " _
        & "AND F.Item = '" & NumEmpresa & "' " _
        & "AND F.CodigoC = C.Codigo "
   Select_AdoDB AdoFacturaDB, sSQL
  'Iniciamos la consulta de impresion
   With AdoFacturaDB
    If .RecordCount > 0 Then
        Validar_Porc_IVA .fields("Fecha")
        TFA.Tipo_Pago = .fields("Tipo_Pago")
       'Encabezado de la Factura
        Cta_Cobrar = .fields("Cta_CxP")
        CodigoL = .fields("Cod_CxC")
        Codigo = .fields("CodigoU")
        Imp_Mes = .fields("Imp_Mes")
        
        SQL2 = "SELECT * " _
             & "FROM Accesos " _
             & "WHERE Codigo = '" & Codigo & "' "
        Select_AdoDB AdoDetalleDB, SQL2
        If AdoDetalleDB.RecordCount > 0 Then NombUsuario = AdoDetalleDB.fields("Nombre_Completo")
        AdoDetalleDB.Close
        
        SQL2 = "SELECT * " _
             & "FROM Clientes " _
             & "WHERE Codigo = '" & .fields("Cod_Ejec") & "' "
        Select_AdoDB AdoDetalleDB, SQL2
        If AdoDetalleDB.RecordCount > 0 Then NombEjecutivo = AdoDetalleDB.fields("Cliente")
        AdoDetalleDB.Close
        
        Cad_Tipo_Pago = Ninguno
        SQL2 = "SELECT * " _
             & "FROM Tabla_Referenciales_SRI " _
             & "WHERE Tipo_Referencia = 'FORMA DE PAGO' " _
             & "AND Codigo = '" & TFA.Tipo_Pago & "' "
        Select_AdoDB AdoDetalleDB, SQL2
        If AdoDetalleDB.RecordCount > 0 Then Cad_Tipo_Pago = AdoDetalleDB.fields("Descripcion")
        AdoDetalleDB.Close
        
        SQL2 = "SELECT * " _
             & "FROM Catalogo_Lineas " _
             & "WHERE Item = '" & NumEmpresa & "' " _
             & "AND Periodo = '" & Periodo_Contable & "' " _
             & "AND TL <> " & Val(adFalse) & " " _
             & "AND CxC = '" & Cta_Cobrar & "' " _
             & "AND Codigo = '" & CodigoL & "' " _
             & "ORDER BY Codigo,CxC "
        Select_AdoDB AdoDetalleDB, SQL2
        If AdoDetalleDB.RecordCount > 0 Then
           CxC_Clientes = AdoDetalleDB.fields("Concepto")
           Cta_Cobrar = AdoDetalleDB.fields("CxC")
           Cta_Ventas = AdoDetalleDB.fields("Cta_Venta")
           TFA.LogoFactura = AdoDetalleDB.fields("Logo_Factura")
           TFA.AltoFactura = AdoDetalleDB.fields("Largo")
           TFA.AnchoFactura = AdoDetalleDB.fields("Ancho")
           TFA.EspacioFactura = AdoDetalleDB.fields("Espacios")
           TFA.Pos_Factura = AdoDetalleDB.fields("Pos_Factura")
          'Individual = AdoDetalleDB.Fields("Individual")
           CodigoL = AdoDetalleDB.fields("Codigo")
           CantFact = AdoDetalleDB.fields("Fact_Pag")
        End If
        AdoDetalleDB.Close
        
        Cadena = ""
        If Len(TFA.Serie) >= 6 Then Cadena = TFA.Serie & "-"
        Cadena = Cadena & Format$(TFA.Factura, "00000000")
        PrinterTexto SetD(2).PosX, SetD(2).PosY, Cadena
        If SetD(8).PosX > 0 And SetD(8).PosY > 0 Then
           Printer.FontSize = SetD(8).Porte
           If .fields("Razon_Social") = .fields("Cliente") Then
               PrinterFields SetD(8).PosX, SetD(8).PosY, .fields("Cliente")
           Else
               PrinterFields SetD(8).PosX, SetD(8).PosY, .fields("Razon_Social")
           End If
        End If
        If SetD(64).PosX > 0 And SetD(64).PosY > 0 Then
           Printer.FontSize = SetD(64).Porte
           If .fields("Razon_Social") <> .fields("Cliente") Then
               PrinterTexto SetD(64).PosX, SetD(64).PosY, .fields("Cliente")
           End If
        End If
        DireccionCli = .fields("Direccion")
        If .fields("DirNumero") <> Ninguno Then
            If .fields("DirNumero") <> "S/N" Then DireccionCli = DireccionCli & " (" & .fields("DirNumero") & ")"
        End If
        Printer.FontSize = SetD(11).Porte
        PrinterTexto SetD(11).PosX, SetD(11).PosY, DireccionCli
        Printer.FontSize = SetD(10).Porte
        PrinterFields SetD(10).PosX, SetD(10).PosY, .fields("Grupo")
       'Codigo del Usuario
        Cadena = ""
        'MsgBox NombUusuario
        Do While NombUsuario <> ""
           Cadena1 = SinEspaciosIzq(NombUsuario)
           If Len(Cadena1) > 0 Then
              NombUsuario = TrimStrg(MidStrg(NombUsuario, Len(Cadena1) + 1, Len(NombUsuario)))
              Cadena = Cadena & MidStrg(Cadena1, 1, 1) & "."
              'MsgBox Cadena
           Else
              NombUsuario = ""
           End If
        Loop
        Printer.FontSize = SetD(18).Porte
        PrinterTexto SetD(18).PosX, SetD(18).PosY, Cadena
        Printer.FontSize = SetD(21).Porte
        PrinterTexto SetD(21).PosX, SetD(21).PosY, NombEjecutivo
        Printer.FontSize = SetD(3).Porte
        PrinterTexto SetD(3).PosX, SetD(3).PosY, FechaStrgCorta(.fields("Fecha"))
        Cadena = FechaDia(.fields("Fecha")) & Space(10) & FechaMes(.fields("Fecha")) & Space(10) & FechaAnio(.fields("Fecha"))
        Printer.FontSize = SetD(4).Porte
        PrinterTexto SetD(4).PosX, SetD(4).PosY, Cadena
        Printer.FontSize = SetD(7).Porte
        PrinterTexto SetD(7).PosX, SetD(7).PosY, FechaStrgCiudad(.fields("Fecha"))
        Printer.FontSize = SetD(5).Porte
        PrinterTexto SetD(5).PosX, SetD(5).PosY, FechaStrgCorta(.fields("Fecha_V"))
        Cadena = FechaDia(.fields("Fecha_V")) & Space(10) & FechaMes(.fields("Fecha_V")) & Space(10) & FechaAnio(.fields("Fecha_V"))
        Printer.FontSize = SetD(6).Porte
        PrinterTexto SetD(6).PosX, SetD(6).PosY, Cadena
        Printer.FontSize = SetD(14).Porte
        PrinterFields SetD(14).PosX, SetD(14).PosY, .fields("Telefono")
        Printer.FontSize = SetD(12).Porte
        PrinterFields SetD(12).PosX, SetD(12).PosY, .fields("Ciudad")
        Printer.FontSize = SetD(13).Porte
        PrinterFields SetD(13).PosX, SetD(13).PosY, .fields("CI_RUC")
        Printer.FontSize = SetD(15).Porte
        PrinterFields SetD(15).PosX, SetD(15).PosY, .fields("Email")
        Printer.FontSize = SetD(51).Porte
        PrinterFields SetD(51).PosX, SetD(51).PosY, .fields("DAU")
        Printer.FontSize = SetD(52).Porte
        PrinterFields SetD(52).PosX, SetD(52).PosY, .fields("FUE")
        Printer.FontSize = SetD(53).Porte
        PrinterFields SetD(53).PosX, SetD(53).PosY, .fields("Declaracion")
        Printer.FontSize = SetD(54).Porte
        PrinterFields SetD(54).PosX, SetD(54).PosY, .fields("Remision")
        Printer.FontSize = SetD(55).Porte
        PrinterFields SetD(55).PosX, SetD(55).PosY, .fields("Comercial")
        Printer.FontSize = SetD(56).Porte
        PrinterFields SetD(56).PosX, SetD(56).PosY, .fields("Solicitud")
        Printer.FontSize = SetD(57).Porte
        PrinterFields SetD(57).PosX, SetD(57).PosY, .fields("Cantidad")
        Printer.FontSize = SetD(58).Porte
        PrinterFields SetD(58).PosX, SetD(58).PosY, .fields("Kilos")
        Printer.FontSize = SetD(60).Porte
        PrinterTexto SetD(60).PosX, SetD(60).PosY, Format$(Day(.fields("Fecha")), "00")
        Cadena = UCaseStrg(MidStrg(MesesLetras(CInt(Month(.fields("Fecha")))), 1, 3))
        Printer.FontSize = SetD(61).Porte
        PrinterTexto SetD(61).PosX, SetD(61).PosY, Cadena
        Printer.FontSize = SetD(62).Porte
        PrinterTexto SetD(62).PosX, SetD(62).PosY, Format$(Year(.fields("Fecha")), "0000")
        Printer.FontSize = SetD(67).Porte
        PrinterTexto SetD(67).PosX, SetD(67).PosY, FechaStrgCorta(.fields("Fecha_A"))
        Printer.FontSize = SetD(16).Porte
        NumeroLineas = PrinterLineasMayor(SetD(16).PosX, SetD(16).PosY, .fields("Observacion"), SetD(26).PosX)
        Printer.FontSize = SetD(17).Porte
        NumeroLineas = PrinterLineasMayor(SetD(17).PosX, SetD(17).PosY, .fields("Nota"), SetD(26).PosX)
        'MsgBox "Encabezado...."
    End If
   End With
   'Printer.FontName = TipoConsola
   SaldoInic = 0: SaldoFinal = 0
   Desc_Sin_IVA = 0
   Desc_Con_IVA = 0
   Orden_No_S = ""
   
  'Comenzamos a recoger los detalles de la factura
   sSQL = "SELECT DF.*,CP.Detalle,CP.Codigo_Barra,CP.Unidad,CM.Marca " _
        & "FROM Detalle_Factura As DF,Catalogo_Productos As CP,Catalogo_Marcas As CM " _
        & "WHERE DF.Factura = " & TFA.Factura & " " _
        & "AND DF.TC = '" & TFA.TC & "' " _
        & "AND DF.Serie = '" & TFA.Serie & "' " _
        & "AND DF.Item = '" & NumEmpresa & "' " _
        & "AND DF.Periodo = '" & Periodo_Contable & "' " _
        & "AND DF.Periodo = CP.Periodo " _
        & "AND DF.Item = CP.Item " _
        & "AND DF.Periodo = CM.Periodo " _
        & "AND DF.Item = CM.Item " _
        & "AND DF.Codigo = CP.Codigo_Inv " _
        & "AND DF.CodMarca = CM.CodMar " _
        & "ORDER BY DF.D_No,DF.Codigo "
   Select_AdoDB AdoDetalleDB, sSQL
   With AdoDetalleDB
    If .RecordCount > 0 Then
        PFil = SetD(22).PosY
        Orden_No = .fields("Orden_No")
        Do While (Not .EOF)
           'MsgBox .RecordCount
           SaldoInic = SaldoInic + .fields("Cantidad")
           SaldoFinal = SaldoFinal + .fields("Tonelaje")
           Printer.FontSize = SetD(23).Porte
           PrinterFields SetD(23).PosX, PFil, .fields("Codigo")
           Printer.FontSize = SetD(30).Porte
           PrinterFields SetD(30).PosX, PFil, .fields("Codigo_Barra")
           If .fields("CodMarca") <> Ninguno Then
               Printer.FontSize = SetD(31).Porte
               PrinterTexto SetD(31).PosX, PFil, "(" & TrimStrg(MidStrg(.fields("Marca"), 1, 3)) & ")"
           End If
           Printer.FontSize = SetD(33).Porte
           PrinterFields SetD(33).PosX, PFil, .fields("Ruta")
           Printer.FontSize = SetD(32).Porte
           PrinterFields SetD(32).PosX, PFil, .fields("Unidad")
           If .fields("Cant_Hab") > 0 And Len(.fields("Tipo_Hab")) > 1 Then
               Printer.FontSize = SetD(70).Porte
               PrinterFields SetD(70).PosX, PFil, .fields("Fecha_IN")
               Printer.FontSize = SetD(71).Porte
               PrinterFields SetD(71).PosX, PFil, .fields("Fecha_OUT")
               Printer.FontSize = SetD(72).Porte
               PrinterFields SetD(72).PosX, PFil, .fields("Cant_Hab")
               Printer.FontSize = SetD(73).Porte
               PrinterFields SetD(73).PosX, PFil, .fields("Tipo_Hab")
           End If
           If SetD(24).PosX <= SetD(25).PosX Then
              Printer.FontSize = SetD(24).Porte
              PrinterTexto SetD(24).PosX, PFil, " " & CStr(.fields("Cantidad"))
           End If
           Printer.FontSize = SetD(25).Porte
           PFil = PrinterLineasTexto(SetD(25).PosX, PFil, .fields("Producto"), SetD(26).PosX)
           If SetD(24).PosX > SetD(25).PosX Then
              Printer.FontSize = SetD(24).Porte
              PrinterTexto SetD(24).PosX, PFil, " " & CStr(.fields("Cantidad"))
           End If
           Printer.FontSize = SetD(29).Porte
           PrinterFields SetD(29).PosX, PFil, .fields("Total")
           PVP_Desc = .fields("Total") - .fields("Total_Desc")
           Printer.FontSize = SetD(69).Porte
           PrinterVariables SetD(69).PosX, PFil, PVP_Desc
           Printer.FontSize = SetD(28).Porte
           PrinterFields SetD(28).PosX, PFil, .fields("Precio")
           Printer.FontSize = SetD(27).Porte
           If .fields("Precio") > 0 And .fields("Cantidad") > 0 Then
               PrinterTexto SetD(27).PosX, PFil, Format$(.fields("Total_Desc") / (.fields("Cantidad") * .fields("Precio")), "00.00%")
           End If
           Printer.FontSize = SetD(34).Porte
           PrinterTexto SetD(34).PosX, PFil, Format$(.fields("Tonelaje"), "#,##0.00")
          'Descuentos en Ventas
           If .fields("Total_IVA") > 0 Then
               Desc_Con_IVA = Desc_Con_IVA + .fields("Total_Desc")
           Else
               Desc_Sin_IVA = Desc_Sin_IVA + .fields("Total_Desc")
           End If
           
           'PFil = PFil + 0.5
           PFil = PFil + Printer.TextHeight("H")
           If Orden_No <> .fields("Orden_No") And .fields("Orden_No") > 0 Then
              Orden_No_S = Orden_No_S & CStr(Orden_No) & " "
              Orden_No = .fields("Orden_N")
           End If
          .MoveNext
        Loop
        Orden_No_S = Orden_No_S & CStr(Orden_No) & " "
    End If
   End With
   AdoDetalleDB.Close
   
  'Pie de factura
   With AdoFacturaDB
    If .RecordCount > 0 Then
        Total = Redondear(.fields("Total_MN"), 2)
        Total_ME = Redondear(.fields("Total_ME"), 2)
       'Porcentaje Con IVA
        Printer.FontSize = SetD(39).Porte
        PrinterTexto SetD(39).PosX, SetD(39).PosY, CStr(Porc_IVA * 100) & " "
       'Con IVA
        Printer.FontSize = SetD(38).Porte
        Desc_Con_IVA = .fields("Con_IVA") - Desc_Con_IVA
        PrinterVariables SetD(38).PosX, SetD(38).PosY, Desc_Con_IVA   '.Fields("Con_IVA")
       'Sin IVA
        Printer.FontSize = SetD(37).Porte
        Desc_Sin_IVA = .fields("Sin_IVA") - Desc_Sin_IVA
        PrinterVariables SetD(37).PosX, SetD(37).PosY, Desc_Sin_IVA   '.Fields("Sin_IVA")
       'Descuento palabra
        Printer.FontSize = SetD(43).Porte
        PrinterTexto SetD(43).PosX, SetD(43).PosY, "Descuento"
       'Total Descuento
        Printer.FontSize = SetD(42).Porte
        PrinterFields SetD(42).PosX, SetD(42).PosY, .fields("Descuento")
       'Total Comision
        Printer.FontSize = SetD(48).Porte
        PrinterFields SetD(48).PosX, SetD(48).PosY, .fields("Comision")
       'Total Servicio
        Printer.FontSize = SetD(49).Porte
        PrinterFields SetD(49).PosX, SetD(49).PosY, .fields("Servicio")
       'IVA Porcentaje
        Printer.FontSize = SetD(41).Porte
        PrinterTexto SetD(41).PosX, SetD(41).PosY, CStr(Porc_IVA * 100) & " "
       'IVA
        Printer.FontSize = SetD(40).Porte
        PrinterFields SetD(40).PosX, SetD(40).PosY, .fields("IVA")
       'SubTotal
        Printer.FontSize = SetD(36).Porte
        PrinterFields SetD(36).PosX, SetD(36).PosY, .fields("SubTotal")
       'SubTotal
        Printer.FontSize = SetD(66).Porte
        PrinterVariables SetD(66).PosX, SetD(66).PosY, .fields("SubTotal") - .fields("Descuento")
       'Total Facturado
        Printer.FontSize = SetD(44).Porte
        PrinterVariables SetD(44).PosX, SetD(44).PosY, Total
       'Total Facturador en letras
        Numero_Letras = Cambio_Letras(Total, 2)
        Printer.FontSize = SetD(45).Porte
        PrinterLineas SetD(45).PosX, SetD(45).PosY, Numero_Letras, 11.5
       'CxC Clientes: Linea de Produccion
        Printer.FontSize = SetD(50).Porte
        PrinterTexto SetD(50).PosX, SetD(50).PosY, TFA.CxC_Clientes
       'Hora de Proceso
        Printer.FontSize = SetD(63).Porte
        PrinterTexto SetD(63).PosX, SetD(63).PosY, .fields("Hora")
        Printer.FontSize = SetD(68).Porte
        PrinterTexto SetD(68).PosX, SetD(68).PosY, TrimStrg(Orden_No_S)
       'Tipo_Pago
        If SetD(78).PosX > 0 And SetD(78).PosY > 0 Then
           Printer.FontSize = SetD(78).Porte
           PrinterVariables SetD(78).PosX, SetD(78).PosY, Total
        End If
        If SetD(77).PosX > 0 And SetD(77).PosY > 0 Then
           RutaOrigen = RutaSistema & "\FORMATOS\Vistofp.jpg"
           PrinterPaint RutaOrigen, SetD(77).PosX, SetD(77).PosY, SetD(77).Porte, SetD(77).Porte
        End If
        If Len(Cad_Tipo_Pago) > 1 Then
           If SetD(76).PosX > 0 And SetD(76).PosY > 0 Then
              Printer.FontSize = SetD(76).Porte
              PrinterTexto SetD(76).PosX, SetD(76).PosY, "TIPO PAGO:"
           End If
           If SetD(75).PosX > 0 And SetD(75).PosY > 0 Then
              Printer.FontSize = SetD(75).Porte
              PrinterTexto SetD(75).PosX, SetD(75).PosY, Cad_Tipo_Pago
           End If
        End If
       'MsgBox Orden_No_S & " ...."
    End If
   End With
   AdoFacturaDB.Close
End If
 
MensajeEncabData = ""
RatonNormal
Lista_Facturas_Text = TextDoc
End Function

Public Sub Imprimir_Guia_Remision(DtaFactura As Adodc, _
                                  DtaDetalle As Adodc, _
                                  TFA As Tipo_Facturas)
Dim CadenaMoneda As String
Dim Numero_Letras As String
Dim CxC_Clientes As String
'Establecemos Espacios y seteos de impresion
On Error GoTo Errorhandler
'MsgBox TipoFact
CEConLineas = ProcesarSeteos("GR")
CantFils = 0
Mensajes = "IMPRIMIR GUIA DE REMISION DE LA FACTURA No. " & Format$(TFA.Factura, "0000000")
Titulo = "IMPRESION"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
   Escala_Centimetro 1, TipoCourier, 9
   RatonReloj
   Printer.FontName = TipoCourier ' TipoArialNarrow
   Printer.FontBold = True
   CxC_Clientes = Ninguno
   ' Printer.FontName = TipoTimes
   sSQL = "SELECT F.*,C.Cliente,C.CI_RUC,C.Telefono,C.Direccion,C.DirNumero,C.Ciudad,C.Grupo,C.Email " _
        & "FROM Facturas As F,Clientes As C " _
        & "WHERE F.Factura = " & TFA.Factura & " " _
        & "AND F.Serie = '" & TFA.Serie & "' " _
        & "AND F.Autorizacion = '" & TFA.Autorizacion & "' " _
        & "AND F.TC = '" & TFA.TC & "' " _
        & "AND F.Periodo = '" & Periodo_Contable & "' " _
        & "AND F.Item = '" & NumEmpresa & "' " _
        & "AND C.Codigo = F.CodigoC "
   Select_Adodc DtaFactura, sSQL
  'Iniciamos la consulta de impresion
   With DtaFactura.Recordset
    If .RecordCount > 0 Then
       'Encabezado de la Factura
        Cta_Cobrar = .fields("Cta_CxP")
        CodigoL = .fields("Cod_CxC")
        SQL2 = "SELECT * " _
             & "FROM Catalogo_Lineas " _
             & "WHERE Item = '" & NumEmpresa & "' " _
             & "AND Periodo = '" & Periodo_Contable & "' " _
             & "AND TL <> " & Val(adFalse) & " " _
             & "AND CxC = '" & Cta_Cobrar & "' " _
             & "AND Codigo = '" & CodigoL & "' " _
             & "ORDER BY Codigo,CxC "
        Select_Adodc DtaDetalle, SQL2
        If DtaDetalle.Recordset.RecordCount > 0 Then
           CxC_Clientes = DtaDetalle.Recordset.fields("Concepto")
           Cta_Cobrar = DtaDetalle.Recordset.fields("CxC")
           Cta_Ventas = DtaDetalle.Recordset.fields("Cta_Venta")
           TFA.LogoFactura = DtaDetalle.Recordset.fields("Logo_Factura")
           TFA.AltoFactura = DtaDetalle.Recordset.fields("Largo")
           TFA.AnchoFactura = DtaDetalle.Recordset.fields("Ancho")
           TFA.EspacioFactura = DtaDetalle.Recordset.fields("Espacios")
           TFA.Pos_Factura = DtaDetalle.Recordset.fields("Pos_Factura")
           'Individual = DtaDetalle.Recordset.Fields("Individual")
           CodigoL = DtaDetalle.Recordset.fields("Codigo")
           CantFact = DtaDetalle.Recordset.fields("Fact_Pag")
        End If
        'MsgBox LogoFactura
        If TFA.LogoFactura <> "NINGUNO" And TFA.AnchoFactura > 0 And TFA.AltoFactura > 0 Then
           If SetD(1).PosX > 0 And SetD(1).PosY > 0 Then
'''              Codigo4 = Format$(.Fields("Factura"), "0000000")
'''              Cadena = RutaSistema & "\FORMATOS\" & LogoFactura & ".GIF"
'''              PrinterPaint Cadena, SetD(1).PosX, SetD(1).PosY, AnchoFactura, AltoFactura
'''              PrinterPaint LogoTipo, SetD(34).PosX, SetD(34).PosY, 2.5, 1.25
           End If
        End If
        'NumFact = .Fields("Remision")
        DireccionCli = .fields("Direccion")
        DireccionGuia = .fields("Comercial")
        Printer.FontSize = SetD(2).Porte
        PrinterTexto SetD(2).PosX, SetD(2).PosY, Format$(.fields("Remision"), "0000000")
        Printer.FontSize = SetD(8).Porte
        PrinterFields SetD(8).PosX, SetD(8).PosY, .fields("Cliente")
        Printer.FontSize = SetD(11).Porte
        PrinterTexto SetD(11).PosX, SetD(11).PosY, DireccionGuia
        Printer.FontSize = SetD(10).Porte
        PrinterFields SetD(10).PosX, SetD(10).PosY, .fields("Grupo")
        Cadena = "Elab.[" & .fields("CodigoU") & "]"
        Printer.FontSize = SetD(18).Porte
        PrinterTexto SetD(18).PosX, SetD(18).PosY, Cadena
        Printer.FontSize = SetD(3).Porte
        PrinterTexto SetD(3).PosX, SetD(3).PosY, FechaStrgCorta(.fields("Fecha"))
        Cadena = FechaDia(.fields("Fecha")) & Space(10) & FechaMes(.fields("Fecha")) & Space(10) & FechaAnio(.fields("Fecha"))
        Printer.FontSize = SetD(4).Porte
        PrinterTexto SetD(4).PosX, SetD(4).PosY, Cadena
        Printer.FontSize = SetD(7).Porte
        PrinterTexto SetD(7).PosX, SetD(7).PosY, FechaStrgCiudad(.fields("Fecha"))
        Printer.FontSize = SetD(5).Porte
        PrinterTexto SetD(5).PosX, SetD(5).PosY, FechaStrgCorta(.fields("Fecha_V"))
        Cadena = FechaDia(.fields("Fecha_V")) & Space(10) & FechaMes(.fields("Fecha_V")) & Space(10) & FechaAnio(.fields("Fecha_V"))
        Printer.FontSize = SetD(6).Porte
        PrinterTexto SetD(6).PosX, SetD(6).PosY, Cadena
        Printer.FontSize = SetD(14).Porte
        PrinterFields SetD(14).PosX, SetD(14).PosY, .fields("Telefono")
        Printer.FontSize = SetD(12).Porte
        PrinterFields SetD(12).PosX, SetD(12).PosY, .fields("Ciudad")
        Printer.FontSize = SetD(13).Porte
        PrinterFields SetD(13).PosX, SetD(13).PosY, .fields("CI_RUC")
        Printer.FontSize = SetD(15).Porte
        PrinterFields SetD(15).PosX, SetD(15).PosY, .fields("Email")
        Printer.FontSize = SetD(51).Porte
        PrinterFields SetD(51).PosX, SetD(51).PosY, .fields("DAU")
        Printer.FontSize = SetD(52).Porte
        PrinterFields SetD(52).PosX, SetD(52).PosY, .fields("FUE")
        Printer.FontSize = SetD(53).Porte
        PrinterFields SetD(53).PosX, SetD(53).PosY, .fields("Declaracion")
        Printer.FontSize = SetD(56).Porte
        PrinterFields SetD(56).PosX, SetD(56).PosY, .fields("Solicitud")
        Printer.FontSize = SetD(57).Porte
        PrinterFields SetD(57).PosX, SetD(57).PosY, .fields("Cantidad")
        Printer.FontSize = SetD(58).Porte
        PrinterFields SetD(58).PosX, SetD(58).PosY, .fields("Kilos")
        Printer.FontSize = SetD(60).Porte
        PrinterTexto SetD(60).PosX, SetD(60).PosY, Format$(Day(.fields("Fecha")), "00")
        Cadena = UCaseStrg(MidStrg(MesesLetras(CInt(Month(.fields("Fecha")))), 1, 3))
        Printer.FontSize = SetD(61).Porte
        PrinterTexto SetD(61).PosX, SetD(61).PosY, Cadena
        Printer.FontSize = SetD(62).Porte
        PrinterTexto SetD(62).PosX, SetD(62).PosY, Format$(Year(.fields("Fecha")), "0000")
        Printer.FontSize = SetD(16).Porte
        NumeroLineas = PrinterLineasMayor(SetD(16).PosX, SetD(16).PosY, .fields("Observacion"), SetD(26).PosX)
        Printer.FontSize = SetD(17).Porte
        NumeroLineas = PrinterLineasMayor(SetD(17).PosX, SetD(17).PosY, .fields("Nota"), SetD(26).PosX)
        'MsgBox "Encabezado...."
    End If
   End With
   'Printer.FontName = TipoConsola
   SaldoInic = 0: SaldoFinal = 0
  'Comenzamos a recoger los detalles de la factura
   sSQL = "SELECT DF.*,CP.Detalle,CP.Codigo_Barra " _
        & "FROM Detalle_Factura As DF,Catalogo_Productos As CP " _
        & "WHERE DF.Factura = " & TFA.Factura & " " _
        & "AND DF.Serie = '" & TFA.Serie & "' " _
        & "AND DF.Autorizacion = '" & TFA.Autorizacion & "' " _
        & "AND DF.TC = '" & TFA.TC & "' " _
        & "AND DF.Item = '" & NumEmpresa & "' " _
        & "AND DF.Periodo = '" & Periodo_Contable & "' " _
        & "AND DF.Periodo = CP.Periodo " _
        & "AND DF.Item = CP.Item " _
        & "AND DF.Codigo = CP.Codigo_Inv " _
        & "ORDER BY DF.D_No,DF.Codigo "
   Select_Adodc DtaDetalle, sSQL
   'MsgBox sSQL
   With DtaDetalle.Recordset
    If .RecordCount > 0 Then
        PFil = SetD(22).PosY
        'MsgBox PFil
        Do While (Not .EOF)
           SaldoInic = SaldoInic + .fields("Cantidad")
           SaldoFinal = SaldoFinal + .fields("Tonelaje")
           Printer.FontSize = SetD(23).Porte
           PrinterFields SetD(23).PosX, PFil, .fields("Codigo")
           Printer.FontSize = SetD(30).Porte
           PrinterFields SetD(30).PosX, PFil, .fields("Codigo_Barra")
           Printer.FontSize = SetD(33).Porte
           PrinterFields SetD(33).PosX, PFil, .fields("Ruta")
           If SetD(24).PosX <= SetD(25).PosX Then
              Printer.FontSize = SetD(24).Porte
              PrinterTexto SetD(24).PosX, PFil, " " & CStr(.fields("Cantidad"))
           End If
           Printer.FontSize = SetD(25).Porte
           PFil = PrinterLineasTexto(SetD(25).PosX, PFil, .fields("Producto"), SetD(26).PosX)
           If SetD(24).PosX > SetD(25).PosX Then
              Printer.FontSize = SetD(24).Porte
              PrinterTexto SetD(24).PosX, PFil, " " & CStr(.fields("Cantidad"))
           End If
           Printer.FontSize = SetD(29).Porte
           PrinterFields SetD(29).PosX, PFil, .fields("Total")
           Printer.FontSize = SetD(28).Porte
           PrinterFields SetD(28).PosX, PFil, .fields("Precio")
           Printer.FontSize = SetD(34).Porte
           PrinterTexto SetD(34).PosX, PFil, Format$(.fields("Tonelaje"), "#,##0.00")
           'MsgBox "Detalle...."
           PFil = PFil + Printer.TextHeight("H") + 0.05
           'PFil = PFil + 0.5
          .MoveNext
        Loop
    End If
   End With
   'Printer.FontName = TipoTimes
  'Pie de factura
   With DtaFactura.Recordset
    If .RecordCount > 0 Then
        Total = Redondear(.fields("Total_MN"), 2)
        Total_ME = Redondear(.fields("Total_ME"), 2)
       'Porcentaje Con IVA
        Printer.FontSize = SetD(39).Porte
        PrinterTexto SetD(39).PosX, SetD(39).PosY, CStr(Porc_IVA * 100) & " "
       'Sin IVA
        Printer.FontSize = SetD(37).Porte
        PrinterFields SetD(37).PosX, SetD(37).PosY, .fields("Sin_IVA")
       'Con IVA
        Printer.FontSize = SetD(38).Porte
        PrinterFields SetD(38).PosX, SetD(38).PosY, .fields("Con_IVA")
       'Descuento palabra
        Printer.FontSize = SetD(43).Porte
        PrinterTexto SetD(43).PosX, SetD(43).PosY, "Descuento"
       'Total Descuento
        Printer.FontSize = SetD(42).Porte
        PrinterFields SetD(42).PosX, SetD(42).PosY, .fields("Descuento")
       'Total Comision
        Printer.FontSize = SetD(48).Porte
        PrinterFields SetD(48).PosX, SetD(48).PosY, .fields("Comision")
       'Total Servicio
        Printer.FontSize = SetD(49).Porte
        PrinterFields SetD(49).PosX, SetD(49).PosY, .fields("Servicio")
       'IVA Porcentaje
        Printer.FontSize = SetD(41).Porte
        PrinterTexto SetD(41).PosX, SetD(41).PosY, CStr(Porc_IVA * 100) & " "
       'IVA
        Printer.FontSize = SetD(40).Porte
        PrinterFields SetD(40).PosX, SetD(40).PosY, .fields("IVA")
       'SubTotal
        Printer.FontSize = SetD(36).Porte
        PrinterFields SetD(36).PosX, SetD(36).PosY, .fields("SubTotal")
       'Total Facturado
        Printer.FontSize = SetD(44).Porte
        PrinterVariables SetD(44).PosX, SetD(44).PosY, Total
       'Total Facturador en letras
        Numero_Letras = Cambio_Letras(Total)
        Printer.FontSize = SetD(45).Porte
        PrinterLineas SetD(45).PosX, SetD(45).PosY, Numero_Letras, 11.5
       'CxC Clientes: Linea de Produccion
        Printer.FontSize = SetD(50).Porte
        PrinterTexto SetD(50).PosX, SetD(50).PosY, TFA.CxC_Clientes
        'MsgBox "...."
    End If
   End With
End If
'Printer.FontName = LetraAnterior
MensajeEncabData = ""
RatonNormal
Printer.EndDoc
Exit Sub
Errorhandler:
    RatonNormal
    ErrorDeImpresion
    Exit Sub
End Sub

'''Public Sub Imprimir_Punto_Venta(TFA As Tipo_Facturas)
'''Dim AdoDBFactura As ADODB.Recordset
'''Dim AdoDBDetalle As ADODB.Recordset
'''Dim CadenaMoneda As String
'''Dim Numero_Letras As String
'''Dim Cant_Ln As Byte
'''Dim CantGuion As Byte
'''Dim CantBlancos As String
'''
'''On Error GoTo Errorhandler
'''Mensajes = "Imprmir Factura No. " & TFA.Factura
'''Titulo = "IMPRESION"
'''Bandera = False
'''SetPrinters.Show 1
'''If PonImpresoraDefecto(SetNombrePRN) Then
'''
'''   Escala_Centimetro 1, TipoConsola, 8  'TipoTerminal
'''
'''   RatonReloj
'''   CantGuion = CByte(Leer_Campo_Empresa("Cant_Ancho_PV"))
'''   If CantGuion < 26 Then CantGuion = 26
'''   Total = 0: Total_IVA = 0
'''   Cant_Ln = 0
'''   PosLinea = 0.1
'''   Producto = ""
'''   If TFA.TC = "PV" Then
'''      sSQL = "SELECT F.*,C.Cliente,C.CI_RUC,C.Telefono,C.Direccion,C.Ciudad,C.Grupo,C.Email " _
'''           & "FROM Trans_Ticket As F,Clientes As C " _
'''           & "WHERE F.Ticket = " & TFA.Factura & " " _
'''           & "AND F.TC = '" & TFA.TC & "' " _
'''           & "AND F.Periodo = '" & Periodo_Contable & "' " _
'''           & "AND F.Item = '" & NumEmpresa & "' " _
'''           & "AND C.Codigo = F.CodigoC "
'''   Else
'''      sSQL = "SELECT F.*,C.Cliente,C.CI_RUC,C.Telefono,C.Direccion,C.Ciudad,C.Grupo,C.Email " _
'''           & "FROM Facturas As F,Clientes As C " _
'''           & "WHERE F.Factura = " & TFA.Factura & " " _
'''           & "AND F.TC = '" & TFA.TC & "' " _
'''           & "AND F.Periodo = '" & Periodo_Contable & "' " _
'''           & "AND F.Item = '" & NumEmpresa & "' " _
'''           & "AND C.Codigo = F.CodigoC "
'''   End If
'''   Select_AdoDB AdoDBFactura, sSQL
'''  'Iniciamos la consulta de impresion
'''  With AdoDBFactura
'''   If .RecordCount > 0 Then
'''      'Encabezado de la Factura
'''       If Encabezado_PV Then
'''          If Len(Empresa) > CantGuion Then
'''          Producto = " " & vbCrLf _
'''                   & UCaseStrg(Empresa) & vbCrLf _
'''                   & NombreComercial & vbCrLf _
'''                   & UCaseStrg(NombreGerente) & vbCrLf _
'''                   & "R.U.C. " & RUC & vbCrLf _
'''                   & "Telefono: " & Telefono1 & vbCrLf _
'''                   & Direccion & vbCrLf
'''          Else
'''          Producto = " " & vbCrLf _
'''                   & Space((CantGuion - Len(Empresa)) / 2) & UCaseStrg(Empresa) & vbCrLf _
'''                   & Space((CantGuion - Len(NombreComercial)) / 2) & NombreComercial & vbCrLf _
'''                   & Space((CantGuion - Len(UCaseStrg(NombreGerente))) / 2) & UCaseStrg(NombreGerente) & vbCrLf _
'''                   & Space((CantGuion - Len("R.U.C. " & RUC)) / 2) & "R.U.C. " & RUC & vbCrLf _
'''                   & Space((CantGuion - Len("Telefono: " & Telefono1)) / 2) & "Telefono: " & Telefono1 & vbCrLf _
'''                   & Direccion & vbCrLf
'''          End If
'''          Cant_Ln = Cant_Ln + 7
'''          If TFA.TC = "PV" Then
'''             Producto = Producto & " " & vbCrLf & "T I C K E T   No. 000-000-" & Format$(TFA.Factura, "0000000") & vbCrLf & " " & vbCrLf
'''             Cant_Ln = Cant_Ln + 1
'''          ElseIf TFA.TC = "NV" Then
'''             Producto = Producto & "Auto. SRI: " & Autorizacion & " - Caduca: " & MidStrg(UCaseStrg(MesesLetras(Month(Fecha_Vence))), 1, 3) & "/" & Year(Fecha_Vence) & vbCrLf & " " & vbCrLf _
'''                      & "NOTA DE VENTA No. " & SerieFactura & "-" & Format$(TFA.Factura, "0000000") & vbCrLf & " " & vbCrLf
'''             Cant_Ln = Cant_Ln + 2
'''          Else
'''             Producto = Producto & "Auto. SRI: " & Autorizacion & " - Caduca: " & MidStrg(UCaseStrg(MesesLetras(Month(Fecha_Vence))), 1, 3) & "/" & Year(Fecha_Vence) & vbCrLf & " " & vbCrLf _
'''                      & "FACTURA No. " & SerieFactura & "-" & Format$(TFA.Factura, "0000000") & vbCrLf & " " & vbCrLf
'''             Cant_Ln = Cant_Ln + 2
'''          End If
'''       Else
'''          Producto = vbCrLf & " " & vbCrLf & " " & vbCrLf & " " & vbCrLf & " " & vbCrLf _
'''                   & "Transaccion(" & TFA.TC & ") No." & Format$(TFA.Factura, "0000000") & vbCrLf & " " & vbCrLf
'''          Cant_Ln = Cant_Ln + 4
'''       End If
'''       Producto = Producto & "Fecha: " & FechaSistema & " - Hora: " & .fields("Hora") & vbCrLf
'''       Producto = Producto & "Cliente: " & vbCrLf _
'''                & MidStrg(.fields("Cliente"), 1, 33) & vbCrLf
'''       Producto = Producto & "R.U.C./C.I.: " & .fields("CI_RUC") & vbCrLf _
'''                & "Cajero: " & MidStrg(CodigoUsuario, 1, 6) & vbCrLf
'''       If .fields("Telefono") <> Ninguno Then Producto = Producto & "Telefono: " & .fields("Telefono") & vbCrLf
'''       If .fields("Direccion") <> Ninguno Then Producto = Producto & "Direccion: " & vbCrLf & .fields("Direccion") & vbCrLf
'''       If .fields("Email") <> Ninguno Then Producto = Producto & "Email: " & vbCrLf & .fields("Email") & vbCrLf
'''       Producto = Producto & String$(CantGuion, "-") & vbCrLf _
'''                & "PRODUCTO/Cant x PVP/TOTAL" & vbCrLf _
'''                & String$(CantGuion, "-") & vbCrLf
'''                Efectivo = .fields("Efectivo")
'''       Cant_Ln = Cant_Ln + 6
'''   End If
'''  End With
''' 'Comenzamos a recoger los detalles de la factura
'''  If TFA.TC = "PV" Then
'''     sSQL = "SELECT DF.*,CP.Detalle,CP.Codigo_Barra " _
'''          & "FROM Trans_Ticket As DF,Catalogo_Productos As CP " _
'''          & "WHERE DF.Ticket = " & TFA.Factura & " " _
'''          & "AND DF.TC = '" & TFA.TC & "' " _
'''          & "AND DF.Item = '" & NumEmpresa & "' " _
'''          & "AND DF.Periodo = '" & Periodo_Contable & "' " _
'''          & "AND DF.Item = CP.Item " _
'''          & "AND DF.Periodo = CP.Periodo " _
'''          & "AND DF.Codigo_Inv = CP.Codigo_Inv " _
'''          & "ORDER BY DF.D_No "
'''  Else
'''      sSQL = "SELECT DF.*,CP.Detalle,CP.Codigo_Barra " _
'''           & "FROM Detalle_Factura As DF,Catalogo_Productos As CP " _
'''           & "WHERE DF.Factura = " & TFA.Factura & " " _
'''           & "AND DF.TC = '" & TFA.TC & "' " _
'''           & "AND DF.Item = '" & NumEmpresa & "' " _
'''           & "AND DF.Periodo = '" & Periodo_Contable & "' " _
'''           & "AND DF.Item = CP.Item " _
'''           & "AND DF.Periodo = CP.Periodo " _
'''           & "AND DF.Codigo = CP.Codigo_Inv " _
'''           & "ORDER BY DF.Codigo "
'''  End If
'''  Select_AdoDB AdoDBDetalle, sSQL
'''  With AdoDBDetalle
'''   If .RecordCount > 0 Then
'''       Do While (Not .EOF)
'''          Producto = Producto & .fields("Producto") & vbCrLf _
'''                   & SetearBlancos(CStr(.fields("Cantidad")) & "x" & Format$(.fields("Precio"), "#,##0.00"), 12, 0, False) & " " _
'''                   & SetearBlancos(CStr(.fields("Total")), CantGuion - 13, 0, True, , True) & vbCrLf
'''          Total = Total + .fields("Total")
'''          If TFA.TC <> "PV" Then Total_IVA = Total_IVA + .fields("Total_IVA")
'''          Cant_Ln = Cant_Ln + 1
'''         .MoveNext
'''       Loop
'''   End If
'''  End With
''' 'Pie de factura
''' '===========================================================
'''  With AdoDBFactura
'''   If .RecordCount > 0 Then
'''       If TFA.TC = "PV" Then
'''          SubTotal = .fields("Total")
'''          Total = .fields("Total")
'''          Total_IVA = 0
'''          Total_Servicio = 0
'''          Total_Desc = 0
'''       Else
'''          SubTotal = .fields("SubTotal")
'''          Total = .fields("Total_MN")
'''          Total_IVA = .fields("IVA")
'''          Total_Servicio = .fields("Servicio")
'''          Total_Desc = .fields("Descuento")
'''       End If
'''       Producto = Producto & String$(CantGuion, "-") & vbCrLf
'''       Cant_Ln = Cant_Ln + 1
'''       'If Total_IVA Then
'''       If (CantGuion - 26) > 0 Then CantBlancos = String$(CantGuion - 26, " ") Else CantBlancos = ""
'''          Producto = Producto _
'''                   & CantBlancos & "     SUBTOTAL " & SetearBlancos(CStr(SubTotal), 12, 0, True, False, True) & vbCrLf _
'''                   & CantBlancos & "    I.V.A " & Porc_IVA * 100 & "% " & SetearBlancos(CStr(Total_IVA), 12, 0, True, False, True) & vbCrLf
'''          Cant_Ln = Cant_Ln + 1
'''
'''       If Total_Servicio > 0 Then
'''          Producto = Producto _
'''                   & CantBlancos & "     SERVICIO " & SetearBlancos(CStr(Total_Servicio), 12, 0, True, False, True) & vbCrLf
'''          Cant_Ln = Cant_Ln + 1
'''       End If
'''       If Total_Desc > 0 Then
'''          Producto = Producto _
'''                   & CantBlancos & "    DESCUENTO " & SetearBlancos(CStr(Total_Desc), 12, 0, True, False, True) & vbCrLf
'''          Cant_Ln = Cant_Ln + 1
'''       End If
'''       If TFA.TC = "PV" Then
'''          Producto = Producto & CantBlancos & "TOTAL TICKET  "
'''       ElseIf TFA.TC = "NV" Then
'''          Producto = Producto & CantBlancos & "TOTAL NOTA V. "
'''       Else
'''          Producto = Producto & CantBlancos & "TOTAL FACTURA "
'''       End If
'''       Producto = Producto & SetearBlancos(CStr(Total), 12, 0, True, False, True) & vbCrLf
'''       If Efectivo > 0 Then
'''          Producto = Producto _
'''                   & CantBlancos & "     EFECTIVO " & SetearBlancos(CStr(Efectivo), 12, 0, True, False, True) & vbCrLf _
'''                   & CantBlancos & "       CAMBIO " & SetearBlancos(CStr(Efectivo - Total), 12, 0, True, False, True) & vbCrLf
'''       End If
'''       If TFA.TC <> "PV" Then
'''          Producto = Producto & "ORIGINAL: CLIENTE" & vbCrLf _
'''                              & "COPIA   : EMISOR" & vbCrLf
'''          If .fields("Cotizacion") > 0 Then Producto = Producto & "COTIZACION: " & Format$(.fields("Cotizacion"), "#,##0.00") & vbCrLf
'''       End If
'''       Producto = Producto & String$(CantGuion, "=") & vbCrLf
'''       If TFA.TC = "PV" Then Producto = Producto & "RECLAME SU FACTURA EN CAJA" & vbCrLf
'''       Producto = Producto & "  GRACIAS POR SU COMPRA " & vbCrLf & " " & vbCrLf _
'''                & " " & vbCrLf & " " & vbCrLf & " " & vbCrLf
'''       Cant_Ln = Cant_Ln + Cant_Item_PV
'''   End If
'''  End With
''' 'Enviamos a la Impresora
'''  'TipoTerminal
'''  'TipoCourier
'''  'TipoConsola
'''  'TipoCourierNew
'''  Printer.FontName = TipoCourierNew
'''  Printer.FontSize = 8
'''  If Copia_PV Then
'''     If Cant_Item_PV < Cant_Ln Then Cant_Item_PV = Cant_Ln
'''     Cadena = ""
'''     Cant_Ln = Cant_Item_PV - Cant_Ln
'''     If Cant_Ln <= 0 Then Cant_Ln = 1
'''     For I = 1 To Cant_Ln
'''         Cadena = Cadena & "` " & vbCrLf
'''     Next I
'''     Producto = Producto & Cadena & Producto & vbCrLf & Cadena
'''  End If
'''
'''  'MsgBox Producto
'''  PrinterPaint LogoTipo, 0.01, PosLinea, 5, 2.4
'''  PosLinea = PosLinea + 2
'''  PrinterLineasTexto 0.01, PosLinea, Producto, 6.5
'''  ''PrinterTexto 0.5, PosLinea, Producto
'''  Printer.EndDoc
'''  AdoDBDetalle.Close
'''  AdoDBFactura.Close
'''End If
'''RatonNormal
'''Exit Sub
'''Errorhandler:
'''    RatonNormal
'''    ErrorDeImpresion
'''    Exit Sub
'''End Sub

Public Sub Imprimir_Punto_Venta(TFA As Tipo_Facturas)
Dim AdoDBFactura As ADODB.Recordset
Dim AdoDBDetalle As ADODB.Recordset
Dim AdoDBDAbonos As ADODB.Recordset
Dim CantGuion As Byte
Dim TotalRecibo As Currency
Dim Si_Copia As Boolean
Dim TipoAporte  As String
Dim CadenaMoneda As String
Dim Numero_Letras As String
Dim CantBlancos As String

On Error GoTo Errorhandler

Si_Copia = True
CantGuion = CByte(Leer_Campo_Empresa("Cant_Ancho_PV"))
Encabezado_PV = CBool(Leer_Campo_Empresa("Encabezado_PV"))
Grafico_PV = CBool(Leer_Campo_Empresa("Grafico_PV"))
Copia_PV = CBool(Leer_Campo_Empresa("Copia_PV"))
If CantGuion < 26 Then CantGuion = 26

Mensajes = "Imprmir Factura No. " & TFA.Factura
Titulo = "IMPRESION"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
   RatonReloj
   sSQL = "SELECT F.*,C.Cliente,C.CI_RUC,C.Telefono,C.Direccion,C.Ciudad,C.Grupo,C.Email "
   If TFA.TC = "PV" Then
      sSQL = sSQL _
           & "FROM Trans_Ticket As F,Clientes As C " _
           & "WHERE F.Ticket = " & TFA.Factura & " "
   Else
      sSQL = sSQL _
           & "FROM Facturas As F,Clientes As C " _
           & "WHERE F.Factura = " & TFA.Factura & " "
   End If
   sSQL = sSQL _
        & "AND F.TC = '" & TFA.TC & "' " _
        & "AND F.Periodo = '" & Periodo_Contable & "' " _
        & "AND F.Item = '" & NumEmpresa & "' " _
        & "AND C.Codigo = F.CodigoC "
   Select_AdoDB AdoDBFactura, sSQL

  'Comenzamos a recoger los detalles de la factura
   If TFA.TC = "PV" Then
      sSQL = "SELECT DF.*,CP.Detalle,CP.Codigo_Barra " _
           & "FROM Trans_Ticket As DF,Catalogo_Productos As CP " _
           & "WHERE DF.Ticket = " & TFA.Factura & " " _
           & "AND DF.TC = '" & TFA.TC & "' " _
           & "AND DF.Item = '" & NumEmpresa & "' " _
           & "AND DF.Periodo = '" & Periodo_Contable & "' " _
           & "AND DF.Item = CP.Item " _
           & "AND DF.Periodo = CP.Periodo " _
           & "AND DF.Codigo_Inv = CP.Codigo_Inv " _
           & "ORDER BY DF.D_No "
   Else
      sSQL = "SELECT DF.*,CP.Detalle,CP.Codigo_Barra " _
           & "FROM Detalle_Factura As DF,Catalogo_Productos As CP " _
           & "WHERE DF.Factura = " & TFA.Factura & " " _
           & "AND DF.TC = '" & TFA.TC & "' " _
           & "AND DF.Item = '" & NumEmpresa & "' " _
           & "AND DF.Periodo = '" & Periodo_Contable & "' " _
           & "AND DF.Item = CP.Item " _
           & "AND DF.Periodo = CP.Periodo " _
           & "AND DF.Codigo = CP.Codigo_Inv " _
           & "ORDER BY DF.Codigo "
  End If
  Select_AdoDB AdoDBDetalle, sSQL
  
  TotalRecibo = 0
  TipoAporte = ""
  sSQL = "SELECT " & Full_Fields("Trans_Abonos") & " " _
       & "FROM Trans_Abonos " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND TP = '" & TFA.TC & "' " _
       & "AND Serie = '" & TFA.Serie & "' " _
       & "AND Autorizacion = '" & TFA.Autorizacion & "' " _
       & "AND Factura = " & TFA.Factura & " " _
       & "ORDER BY Fecha,ID "
  Select_AdoDB AdoDBDAbonos, sSQL
  With AdoDBDAbonos
   If .RecordCount > 0 Then
       Do While Not .EOF
          If Val(.fields("Recibo_No")) > 0 Then TFA.Recibo_No = .fields("Recibo_No") Else TFA.Recibo_No = Format(.fields("Factura"), "000000000")
          If .fields("Banco") = "EFECTIVO MN" Then
              TipoAporte = TipoAporte & "CONTADO/"
          Else
              TipoAporte = TipoAporte & .fields("Banco") & "-" & .fields("Cheque") & "/"
          End If
          TotalRecibo = TotalRecibo + .fields("Abono")
         .MoveNext
       Loop
   Else
         TipoAporte = "CREDITO"
   End If
  End With
  
Copia_Doc:
  Escala_Centimetro 1, TipoCourierNew, 8
  PosLinea = 0.01
 'Iniciamos la consulta de impresion
  With AdoDBFactura
   If .RecordCount > 0 Then
       Total = 0
       Total_IVA = 0
       Codigo1 = .fields("CodigoU")
       TFA.Cliente = .fields("Cliente")
       Codigo1 = MidStrg(Codigo1, 1, 4) & "X" & MidStrg(Codigo1, Len(Codigo1) - 1, 2)
       Producto = ""
      'Si imprimimos el LogoTipo
       If Grafico_PV Then
          PrinterPaint LogoTipo, 0.01, PosLinea, 5, 1.9
          PosLinea = PosLinea + 2.1
       End If
       PrinterFontSize 8
      'Encabezado_PV
       If Encabezado_PV Then
          If RazonSocial = NombreComercial Then
             Producto = RazonSocial & vbCrLf _
                      & "R.U.C. " & RUC & vbCrLf
          ElseIf TFA.TC = "DO" Then
             Producto = RazonSocial & vbCrLf _
                      & NombreComercial & vbCrLf _
                      & "R.U.C. " & RUC & vbCrLf _
                      & " " & vbCrLf _
                      & "DONACION DE ALIMENTOS" & vbCrLf _
                      & " " & vbCrLf
          Else
             Producto = RazonSocial & vbCrLf _
                      & NombreComercial & vbCrLf _
                      & "R.U.C. " & RUC & vbCrLf
          End If
          Producto = Producto & "Direccion: " & Direccion & vbCrLf _
                   & "Telefono: " & Telefono1 & vbCrLf
       End If

      'Encabezado de la Factura
       If Encabezado_PV Then
          If TFA.TC = "PV" Then
             Producto = Producto & "T I C K E T   No. 000-000-" & Format$(TFA.Factura, "0000000") & vbCrLf & " " & vbCrLf
          ElseIf TFA.TC = "NV" Then
             Producto = Producto & "Autorizacion del SRI No." & vbCrLf _
                      & TFA.Autorizacion & vbCrLf _
                      & "NOTA DE VENTA No. " & TFA.Serie & "-" & Format$(TFA.Factura, "0000000") & vbCrLf
          ElseIf TFA.TC = "DO" Then
             Producto = Producto & "NOTA DE DONACION No. " & TFA.Serie & "-" & Format$(TFA.Factura, "0000000") & vbCrLf
          Else
             Producto = Producto & "Autorizacion del SRI No." & vbCrLf _
                      & TFA.Autorizacion & vbCrLf _
                      & "FACTURA No. " & TFA.Serie & "-" & Format$(TFA.Factura, "0000000") & vbCrLf
             If Len(ContEspec) > 1 Then Producto = Producto & "Contribuyente Especial No. " & ContEspec & vbCrLf
             Producto = Producto & "OBLIGADO A LLEVAR CONTABILIDAD: " & Obligado_Conta & vbCrLf
          End If
       Else
          Producto = Producto & "Transaccion (" & TFA.TC & ") No." & Format$(TFA.Factura, "0000000") & vbCrLf & " " & vbCrLf
       End If
       Producto = Producto & "Fecha de Emision: " & .fields("Fecha") & vbCrLf & "Hora de proceso: " & .fields("Hora") & vbCrLf
       If Len(TFA.Autorizacion) < 13 Then Producto = Producto & "Fecha de caducidad: " & MidStrg(UCaseStrg(MesesLetras(Month(Fecha_Vence))), 1, 3) & "/" & Year(Fecha_Vence) & vbCrLf
       Producto = Producto & String(CantGuion, "-") & vbCrLf _
                & "Cliente: " & .fields("Cliente") & vbCrLf _
                & "R.U.C./C.I.: " & .fields("CI_RUC") & vbCrLf
       If .fields("Telefono") <> Ninguno Then Producto = Producto & "Telefono: " & .fields("Telefono") & vbCrLf
       If .fields("Direccion") <> Ninguno Then Producto = Producto & "Direccion: " & .fields("Direccion") & vbCrLf
       If .fields("Email") <> Ninguno Then Producto = Producto & "Email: " & .fields("Email") & vbCrLf
       If TFA.TC = "DO" Then
          Producto = Producto _
                   & "Codigo: " & vbCrLf _
                   & "Aporte: " & vbCrLf _
                   & "Numero de Gavetas: " & vbCrLf _
                   & "Atencion: " & vbCrLf _
                   & String(CantGuion, "=") & vbCrLf _
                   & "P R O D U C T O/CODIGO CANTIDAD(Kg)" & vbCrLf _
                   & String(CantGuion, "=") & vbCrLf
       Else
          Producto = Producto & String(CantGuion, "=") & vbCrLf _
                   & "PRODUCTO/Cant x PVP/TOTAL" & vbCrLf _
                   & String(CantGuion, "=") & vbCrLf
       End If
       Efectivo = .fields("Efectivo")
   End If
  End With
  Total = 0
  If (CantGuion - 26) > 0 Then CantBlancos = String(CantGuion - 26, " ") Else CantBlancos = ""
  With AdoDBDetalle
   If .RecordCount > 0 Then
      .MoveFirst
       Do While (Not .EOF)
          If TFA.TC = "DO" Then
             CodigoC = .fields("Codigo")
             CodigoN = Format(.fields("Cantidad"), "#0.00")
             Producto = Producto & .fields("Producto")
             If .fields("Tipo_Hab") <> Ninguno Then Producto = Producto & "(" & .fields("Tipo_Hab") & ")"
             Producto = Producto & vbCrLf
             Producto = Producto & CodigoC & String(25 - Len(CodigoC), " ") & " " & String(10 - Len(CodigoN), " ") & CodigoN & " " & vbCrLf
             Total = Total + .fields("Cantidad")
          Else
             Producto = Producto & .fields("Producto") & vbCrLf _
                      & SetearBlancos(CStr(.fields("Cantidad")) & "x" & Format$(.fields("Precio"), "#,##0.00"), 12, 0, False) & " " _
                      & SetearBlancos(CStr(.fields("Total")), CantGuion - 13, 0, True, , True) & " " & vbCrLf
             Total = Total + .fields("Total")
             If TFA.TC <> "PV" Then Total_IVA = Total_IVA + .fields("Total_IVA")
          End If
         .MoveNext
       Loop
   End If
  End With
 'Pie de factura
 '===========================================================
  If TFA.TC = "DO" Then
     Producto = Producto & String(CantGuion, "-") & vbCrLf _
              & CantBlancos & "    T O T A L " & SetearBlancos(CStr(Total), 12, 0, True, False, True) & vbCrLf _
              & " " & vbCrLf & " " & vbCrLf & " " & vbCrLf _
              & "_____________      _______________" & vbCrLf _
              & "Entregado por      Recibi Conforme" & vbCrLf & " " & vbCrLf _
              & "IMPORTANTE:" & vbCrLf _
              & "Los productos donados, han perdido valor comercial por diferentes motivos, pero mantienen un valor social. " _
              & "Estos productos han pasado por un proceso de clasificación y se encuentran en buen estado. Se recomienda su " _
              & "consumo INMEDIATO y se prohíbe su comercialización. " & RazonSocial & " no se responsabiliza por " _
              & "cualquier efecto negativo que causare el consumo de alimentos en un tiempo mayor al sugerido.  Con su firma " _
              & "el beneficiario acepta que ha sido informado sobre el estado de los productos, que los recibe con su consentimiento, " _
              & "que los usará para fines benéficos y bajo su completa responsabilidad." & vbCrLf & " " & vbCrLf & " " & vbCrLf
  Else
    With AdoDBFactura
     If .RecordCount > 0 Then
         If TFA.TC = "PV" Then
            SubTotal = .fields("Total")
            Total = .fields("Total")
            Total_IVA = 0
            Total_Servicio = 0
            Total_Desc = 0
         Else
            SubTotal = .fields("SubTotal")
            Total = .fields("Total_MN")
            Total_IVA = .fields("IVA")
            Total_Servicio = .fields("Servicio")
            Total_Desc = .fields("Descuento")
         End If
         Producto = Producto & String(CantGuion, "-") & vbCrLf
         'If Total_IVA Then
         If (CantGuion - 26) > 0 Then CantBlancos = String(CantGuion - 26, " ") Else CantBlancos = ""
            Producto = Producto _
                     & "Cajero:        SUBTOTAL " & SetearBlancos(CStr(SubTotal), 12, 0, True, False, True) & vbCrLf _
                     & Codigo1 & "       I.V.A " & Porc_IVA * 100 & "% " & SetearBlancos(CStr(Total_IVA), 12, 0, True, False, True) & vbCrLf
         If Total_Servicio > 0 Then
            Producto = Producto _
                     & CantBlancos & "     SERVICIO " & SetearBlancos(CStr(Total_Servicio), 12, 0, True, False, True) & vbCrLf
         End If
         If Total_Desc > 0 Then
            Producto = Producto _
                     & CantBlancos & "    DESCUENTO " & SetearBlancos(CStr(Total_Desc), 12, 0, True, False, True) & vbCrLf
         End If
         Producto = Producto & String(CantGuion, "=") & vbCrLf
         If TFA.TC = "PV" Then
            Producto = Producto & CantBlancos & "TOTAL TICKET  "
         ElseIf TFA.TC = "NV" Then
            Producto = Producto & CantBlancos & "TOTAL NOTA V. "
         Else
            Producto = Producto & CantBlancos & "TOTAL FACTURA "
         End If
         Producto = Producto & SetearBlancos(CStr(Total), 12, 0, True, False, True) & vbCrLf
         If Efectivo > 0 Then
            Producto = Producto _
                     & CantBlancos & "     EFECTIVO " & SetearBlancos(CStr(Efectivo), 12, 0, True, False, True) & vbCrLf _
                     & CantBlancos & "       CAMBIO " & SetearBlancos(CStr(Efectivo - Total), 12, 0, True, False, True) & vbCrLf
         End If
         Producto = Producto & " " & vbCrLf
         If TFA.TC <> "PV" Then
            Producto = Producto & "ORIGINAL: CLIENTE" & vbCrLf _
                                & "COPIA   : EMISOR" & vbCrLf
            If .fields("Cotizacion") > 0 Then Producto = Producto & "COTIZACION: " & Format$(.fields("Cotizacion"), "#,##0.00") & vbCrLf
         End If
         Producto = Producto & String(CantGuion, "-") & vbCrLf
         If TFA.TC = "PV" Then Producto = Producto & "RECLAME SU DICUMENTO EN CAJA" & vbCrLf
         Producto = Producto _
                  & "Su Documento sera enviado al correo electronico registrado." & vbCrLf _
                  & " " & vbCrLf _
                  & String((CantGuion - 21) / 2, " ") & "GRACIAS POR SU COMPRA" & vbCrLf & " " & vbCrLf
     End If
    End With
  End If
 'Enviamos a la Impresora
  Printer.FontName = TipoCourierNew
  Printer.FontSize = 8
  
  Escribir_Archivo RutaSysBases & "\TEMP\Impresora_PV.txt", Producto

  PosLinea = PrinterLineasTextoPV(0.01, PosLinea, Producto, CantGuion)
  
  If TFA.TC = "DO" Then
    'Encabezado_PV
     Producto = RazonSocial & vbCrLf _
              & NombreComercial & vbCrLf _
              & "R.U.C. " & RUC & vbCrLf _
              & " " & vbCrLf _
              & "A P O R T E   S O L I D A R I O" & vbCrLf _
              & " " & vbCrLf _
              & "Direccion: " & Direccion & vbCrLf _
              & "Telefono: " & Telefono1 & vbCrLf _
              & " " & vbCrLf _
              & "NOTA DE DONACION No. " & TFA.Serie & "-" & Format$(TFA.Factura, "0000000") & vbCrLf _
              & " " & vbCrLf _
              & "Fecha de Emision: " & TFA.Fecha & vbCrLf _
              & String(CantGuion, "-") & vbCrLf _
              & "Cliente: " & TFA.Cliente & vbCrLf _
              & "R.U.C./C.I.: " & TFA.RUC_CI & vbCrLf _
              & "Telefono: " & TFA.TelefonoC & vbCrLf _
              & String(CantGuion, "-") & vbCrLf _
              & "Queremos agradecerle por su aporte solidario de USD " & Format(TFA.Total_MN, "#,##0.00") & ", donacion que nos permitira incrementar " _
              & "la atencion a un mayor numero de personas en vulnerabilidad alimentaria. Usted es muy importante para nosotros." & vbCrLf _
              & " " & vbCrLf _
              & "SU DONACIÓN PUEDE REALIZARLA EN EFECTIVO, DEPOSITO O TRANSFERENCIA BANCARIA A LA CUENTA DE AHORROS BANCO PICHINCHA No.- 3708204100 " _
              & "A NOMBRE DE " & Empresa & "." & vbCrLf _
              & " " & vbCrLf & " " & vbCrLf & " " & vbCrLf _
              & "______________      _______________ " & vbCrLf _
              & "   RECIBIDO            ENTREGADO " & vbCrLf & " " & vbCrLf
     
     Printer.FontName = TipoCourierNew
     Printer.FontSize = 8
     PosLinea = PosLinea + 0.4
     PrinterTexto 0.01, PosLinea, String(CantGuion, "-")
     PosLinea = PosLinea + 0.4
    'Si imprimimos el LogoTipo
     If Grafico_PV Then
        PrinterPaint LogoTipo, 0.01, PosLinea, 5, 1.9
        PosLinea = PosLinea + 2.2
     End If
     PosLinea = PrinterLineasTextoPV(0.01, PosLinea, Producto, CantGuion)
  End If

 'Imprimimos el Recibo de Pagos
  With AdoDBDAbonos
   If .RecordCount > 0 Then
      .MoveFirst
       Printer.NewPage
       PosLinea = 0.5
      'Si imprimimos el LogoTipo
       If Grafico_PV Then
          PrinterPaint LogoTipo, 0.01, PosLinea, 5, 1.9
          PosLinea = PosLinea + 2.1
       End If
       
       TFA.CodigoU = .fields("CodigoU")
      'Encabezado_PV
       Producto = RazonSocial & vbCrLf _
                & NombreComercial & vbCrLf _
                & "R.U.C. " & RUC & vbCrLf _
                & "Telefono: " & Telefono1 & vbCrLf _
                & "Direccion: " & Direccion & vbCrLf _
                & UCaseStrg(NombreCiudad) & " - ECUADOR " & vbCrLf _
                & String(CantGuion, "-") & vbCrLf _
                & "RECIBO DE INGRESO No. " & Year(.fields("Fecha")) & "-" & TFA.Recibo_No & vbCrLf _
                & "Fecha de Abono: " & TFA.Fecha & vbCrLf _
                & "Por USD " & Format(TotalRecibo, "#,##0.00") & vbCrLf _
                & "La suma de: " & Cambio_Letras(TotalRecibo, 2) & vbCrLf _
                & String(CantGuion, "-") & vbCrLf _
                & "Cliente: " & TFA.Cliente & vbCrLf

       If TFA.TC = "PV" Then
          Producto = Producto & "T I C K E T   No. 000-000-" & Format$(TFA.Factura, "000000000")
       ElseIf TFA.TC = "NV" Then
          Producto = Producto & "NOTA DE VENTA No. "
       ElseIf TFA.TC = "DO" Then
          Producto = Producto & "NOTA DE DONACION No. "
       Else
          Producto = Producto & "FACTURA No. "
       End If
       Producto = Producto & TFA.Serie & "-" & Format$(TFA.Factura, "000000000") & vbCrLf _
                & String(CantGuion, "-") & vbCrLf _
                & "POR CONCEPTO DE: " & vbCrLf _
                & String(CantGuion, "-") & vbCrLf
       Do While Not .EOF
          Producto = Producto & .fields("Fecha") & " - "
          If Len(.fields("Banco")) > 1 Then Producto = Producto & .fields("Banco") & " - "
          If Len(.fields("Cheque")) > 1 Then Producto = Producto & .fields("Cheque") & " - "
          If .fields("Abono") <> 0 Then Producto = Producto & "Por USD " & Format(.fields("Abono"), "#,##0.00")
          Producto = Producto & vbCrLf
         .MoveNext
       Loop
       Producto = Producto & String(CantGuion, "-") & vbCrLf _
                & " " & vbCrLf & " " & vbCrLf & " " & vbCrLf & " " & vbCrLf _
                & "______________      _______________ " & vbCrLf _
                & "CONFORME            PROCESADO " & vbCrLf _
                & "C.I./R.U.C.         POR " & vbCrLf _
                & TFA.RUC_CI & String(20 - Len(TFA.RUC_CI), " ") & TFA.CodigoU & vbCrLf
       PosLinea = PrinterLineasTextoPV(0.01, PosLinea, Producto, CantGuion)
   End If
  End With
  Printer.EndDoc
  
  If Copia_PV And Si_Copia Then
     Si_Copia = False
     GoTo Copia_Doc
  End If
  AdoDBDAbonos.Close
  AdoDBDetalle.Close
  AdoDBFactura.Close
End If
RatonNormal
Exit Sub
Errorhandler:
    RatonNormal
    ErrorDeImpresion
    Exit Sub
End Sub

Public Sub Imprimir_Recibo_Punto_Venta(TTA As Tipo_Abono, Optional FechaRecibo As String, Optional PorMail As Boolean)
Dim AdoDBDAbonos As ADODB.Recordset
Dim Numero_Letras As String
Dim CantGuion As Byte
Dim CantBlancos As String
Dim PorFecha As Boolean

On Error GoTo Errorhandler

CantGuion = CByte(Leer_Campo_Empresa("Cant_Ancho_PV"))
Encabezado_PV = CBool(Leer_Campo_Empresa("Encabezado_PV"))
Grafico_PV = CBool(Leer_Campo_Empresa("Grafico_PV"))
Copia_PV = CBool(Leer_Campo_Empresa("Copia_PV"))
If CantGuion < 26 Then CantGuion = 26

Producto = ""

If IsDate(FechaRecibo) Then PorFecha = True Else PorFecha = False

'MsgBox FechaRecibo
   sSQL = "SELECT TA.Recibo_No, TA.TP, TA.Serie, TA.Factura, TA.Autorizacion, TA.Fecha, TA.Banco, TA.Cheque, TA.Abono, TA.CodigoU, " _
        & "C.Cliente, C.CI_RUC, C.Telefono, C.Direccion, C.Ciudad, C.Grupo, C.Email, C.Email2, C.EmailR " _
        & "FROM Trans_Abonos As TA, Clientes As C " _
        & "WHERE TA.Item = '" & NumEmpresa & "' " _
        & "AND TA.Periodo = '" & Periodo_Contable & "' " _
        & "AND TA.TP = '" & TTA.TP & "' " _
        & "AND TA.Serie = '" & TTA.Serie & "' " _
        & "AND TA.Factura = " & TTA.Factura & " "
   If PorFecha Then sSQL = sSQL & "AND TA.Fecha = #" & BuscarFecha(FechaRecibo) & "# "
   sSQL = sSQL _
        & "AND TA.CodigoC = C.Codigo " _
        & "ORDER BY TA.Fecha, TA.ID "
   Select_AdoDB AdoDBDAbonos, sSQL
   With AdoDBDAbonos
    If .RecordCount > 0 Then
        Total = 0
        Total_IVA = 0
        Codigo1 = MidStrg(.fields("CodigoU"), 1, 4) & "X" & MidStrg(.fields("CodigoU"), Len(.fields("CodigoU")) - 1, 2)
        TMail.para = ""
        Insertar_Mail TMail.para, .fields("Email")
        Insertar_Mail TMail.para, .fields("Email2")
        Insertar_Mail TMail.para, .fields("EmailR")
        Insertar_Mail TMail.para, "diskcover.system@gmail.com"
        Insertar_Mail TMail.para, "informacion@diskcoversystem.com"
       'Encabezado_PV
        If Encabezado_PV Then
           If RazonSocial = NombreComercial Then
              Producto = RazonSocial & vbCrLf _
                       & "R.U.C. " & RUC & vbCrLf
           Else
              Producto = RazonSocial & vbCrLf _
                       & NombreComercial & vbCrLf _
                       & "R.U.C. " & RUC & vbCrLf
           End If
           Producto = Producto _
                    & "Direccion: " & Direccion & vbCrLf _
                    & "Telefono: " & Telefono1 & vbCrLf _
                    & "RECIBO No. " & Year(.fields("Fecha")) & "-" & .fields("Recibo_No") & vbCrLf _
                    & String(CantGuion, "-") & vbCrLf _
                    & "Autorizacion del SRI No." & vbCrLf _
                    & .fields("Autorizacion") & vbCrLf _
                    & "DOCUMENTO " & .fields("TP") & " No. " & .fields("Serie") & "-" & Format$(.fields("Factura"), "0000000") & vbCrLf
        Else
           Producto = RazonSocial & vbCrLf _
                    & "Transaccion (" & .fields("TP") & " No. " & .fields("Serie") & "-" & Format$(.fields("Factura"), "0000000") & vbCrLf
        End If
        If PorFecha Then Producto = Producto & "Fecha de Abono: " & .fields("Fecha") & vbCrLf
        Producto = Producto & String(CantGuion, "-") & vbCrLf _
                 & "Cliente: " & .fields("Cliente") & vbCrLf _
                 & "R.U.C./C.I.: " & .fields("CI_RUC") & vbCrLf
        If .fields("Telefono") <> Ninguno Then Producto = Producto & "Telefono: " & .fields("Telefono") & vbCrLf
        If .fields("Direccion") <> Ninguno Then Producto = Producto & "Direccion: " & .fields("Direccion") & vbCrLf
        If .fields("Email") <> Ninguno Then Producto = Producto & "Email: " & .fields("Email") & vbCrLf
        Producto = Producto & String(CantGuion, "=") & vbCrLf
        If PorFecha Then
           Producto = Producto & "DETALLE COMPROBANTE " & String(6, " ") & "SUBTOTAL" & vbCrLf
        Else
           Producto = Producto & "DETALLE COMPROBANTE/FECHA   SUBTOTAL" & vbCrLf
        End If
        Producto = Producto & String(CantGuion, "=") & vbCrLf
       'Detalle del abono
        Do While Not .EOF
           If .fields("Banco") = "." Then CodigoP = "SD" Else CodigoP = .fields("Banco")
           If .fields("Cheque") <> "." Then CodigoP = CodigoP & "-" & .fields("Cheque")
           Mifecha = FechaStrg(.fields("Fecha"))
           If Not PorFecha Then
              Producto = Producto & CodigoP & vbCrLf _
                       & Mifecha & String(25 - Len(Mifecha), " ") & " " & SetearBlancos(CStr(.fields("Abono")), 10, 0, True, , True) & vbCrLf
           Else
              Producto = Producto & CodigoP & String(25 - Len(CodigoP), " ") & " " & SetearBlancos(CStr(.fields("Abono")), 10, 0, True, , True) & vbCrLf
           End If
           Total = Total + .fields("Abono")
          .MoveNext
        Loop
       'Fin del Recibo
        Producto = Producto & String(CantGuion, "-") & vbCrLf _
                 & "Cajero: " & Codigo1 & String(10 - Len(Codigo1), " ") & "TOTAL " & SetearBlancos(CStr(Total), 12, 0, True, False, True) & vbCrLf _
                 & " " & vbCrLf _
                 & " " & vbCrLf _
                 & " " & vbCrLf _
                 & "_____________      _______________" & vbCrLf _
                 & "Entregado por      Recibi Conforme" & vbCrLf & " " & vbCrLf
'''                 & "IMPORTANTE:" & vbCrLf _
'''                 & "Los productos donados, han perdido valor comercial por diferentes motivos, pero mantienen un valor social. " _
'''                 & "Estos productos han pasado por un proceso de clasificación y se encuentran en buen estado. Se recomienda su " _
'''                 & "consumo INMEDIATO y se prohíbe su comercialización. " & RazonSocial & " no se responsabiliza por " _
'''                 & "cualquier efecto negativo que causare el consumo de alimentos en un tiempo mayor al sugerido.  Con su firma " _
'''                 & "el beneficiario acepta que ha sido informado sobre el estado de los productos, que los recibe con su consentimiento, " _
'''                 & "que los usará para fines benéficos y bajo su completa responsabilidad." & vbCrLf & " " & vbCrLf & " " & vbCrLf
        
    End If
   End With
   AdoDBDAbonos.Close
   If PorMail Then
      Titulo = "Pregunta de Envio de Mails"
      Mensajes = "Esta seguro de querer enviar por mail los documentos?"
      If BoxMensaje = vbYes Then
         TMail.MensajeHTML = ""
         TMail.Mensaje = ""
         TMail.ListaMail = 255
         TMail.TipoDeEnvio = "."
         TMail.Asunto = "Prueba de Mails por smtp.diskcoversystem.com"
         With TMail
             .MensajeHTML = Leer_Archivo_Texto(RutaSistema & "\FORMATOS\recibo_mail.html")
             .MensajeHTML = Replace(TMail.MensajeHTML, "vMensaje", Replace(Producto, vbCrLf, "<br>"))
         End With
        'TMail.Mensaje = Producto
         TMail.Adjunto = ""
         FEnviarCorreos.Show 1
         TMail.para = ""
         TMail.ListaMail = 255
         Generar_File_SQL "Recibo_Mail", TMail.MensajeHTML
         RatonNormal
       End If
   Else
      Mensajes = "Imprmir Recibo de la Factura No. " & TTA.Factura
      Titulo = "IMPRESION"
      Bandera = False
      SetPrinters.Show 1
      If PonImpresoraDefecto(SetNombrePRN) Then
         RatonReloj
        'Enviamos a la Impresora
         Escala_Centimetro 1, TipoCourierNew, 8
         PosLinea = 0.01
         Printer.FontName = TipoCourierNew
         Printer.FontSize = 8
        'Si imprimimos el LogoTipo
         If Grafico_PV Then
            PrinterPaint LogoTipo, 0.01, PosLinea, 5, 1.9
            PosLinea = PosLinea + 2.2
         End If
         PosLinea = PrinterLineasTextoPV(0.01, PosLinea, Producto, CantGuion)
        'Generar_File_SQL "Nota_Donacion", Producto
         Printer.EndDoc
      End If
   End If

RatonNormal
Exit Sub
Errorhandler:
    RatonNormal
    ErrorDeImpresion
    Exit Sub
End Sub

Public Sub Imprimir_Punto_Venta_SRI(DtaFactura As Adodc, _
                                    DtaDetalle As Adodc, _
                                    TFA As Tipo_Facturas)
Dim CadenaMoneda As String
Dim Numero_Letras As String
Dim Cant_Ln As Byte
Dim CantGuion As Byte
Dim CantBlancos As String

On Error GoTo Errorhandler
Mensajes = "Imprmir Factura No. " & TFA.Factura
Titulo = "IMPRESION"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
   Escala_Centimetro 1, TipoTerminal, 9
   RatonReloj
   CantGuion = CByte(Leer_Campo_Empresa("Cant_Ancho_PV"))
   CantGuion = 36
   If CantGuion < 26 Then CantGuion = 26
   Total = 0: Total_IVA = 0
   Cant_Ln = 0
   PosLinea = 0.1
   Producto = ""
   If TFA.TC = "PV" Then
      sSQL = "SELECT F.*,C.Cliente,C.CI_RUC,C.Telefono,C.Direccion,C.Ciudad,C.Grupo,C.Email " _
           & "FROM Trans_Ticket As F,Clientes As C " _
           & "WHERE F.Ticket = " & TFA.Factura & " " _
           & "AND F.TC = '" & TFA.TC & "' " _
           & "AND F.Periodo = '" & Periodo_Contable & "' " _
           & "AND F.Item = '" & NumEmpresa & "' " _
           & "AND C.Codigo = F.CodigoC "
   Else
      sSQL = "SELECT F.*,C.Cliente,C.CI_RUC,C.Telefono,C.Direccion,C.Ciudad,C.Grupo,C.Email " _
           & "FROM Facturas As F,Clientes As C " _
           & "WHERE F.Factura = " & TFA.Factura & " " _
           & "AND F.TC = '" & TFA.TC & "' " _
           & "AND F.Periodo = '" & Periodo_Contable & "' " _
           & "AND F.Item = '" & NumEmpresa & "' " _
           & "AND C.Codigo = F.CodigoC "
   End If
   Select_Adodc DtaFactura, sSQL
  'Iniciamos la consulta de impresion
  With DtaFactura.Recordset
   If .RecordCount > 0 Then
      'Encabezado de la Factura
          Producto = " " & vbCrLf _
                   & Space((CantGuion - Len(Empresa)) / 2) & UCaseStrg(Empresa) & vbCrLf _
                   & Space((CantGuion - Len(NombreComercial)) / 2) & NombreComercial & vbCrLf _
                   & Space((CantGuion - Len(UCaseStrg(NombreGerente))) / 2) & UCaseStrg(NombreGerente) & vbCrLf _
                   & Space((CantGuion - Len("R.U.C. " & RUC)) / 2) & "R.U.C. " & RUC & vbCrLf _
                   & Space((CantGuion - Len("Telefono: " & Telefono1)) / 2) & "Telefono: " & Telefono1 & vbCrLf _
                   & Direccion & vbCrLf
          Cant_Ln = Cant_Ln + 7
          Producto = Producto & "Autorizacion del SRI: " & vbCrLf _
                   & TFA.Autorizacion & " " & vbCrLf _
                   & "FACTURA No. " & TFA.Serie & "-" & Format$(TFA.Factura, "0000000") & " " & vbCrLf _
                   & "Clave de Acceso:" & vbCrLf _
                   & TFA.ClaveAcceso & " " & vbCrLf
          Cant_Ln = Cant_Ln + 5
       Producto = Producto & "Fecha: " & FechaSistema & " - Hora: " & .fields("Hora") & vbCrLf
       Producto = Producto & "Cliente: " & vbCrLf _
                & MidStrg(.fields("Cliente"), 1, 33) & vbCrLf
       Producto = Producto & "R.U.C./C.I.: " & .fields("CI_RUC") & vbCrLf _
                & "Cajero: " & MidStrg(CodigoUsuario, 1, 6) & vbCrLf
       If .fields("Telefono") <> Ninguno Then Producto = Producto & "Telefono: " & .fields("Telefono") & vbCrLf
       If .fields("Direccion") <> Ninguno Then Producto = Producto & "Direccion: " & vbCrLf & .fields("Direccion") & vbCrLf
       If .fields("Email") <> Ninguno Then Producto = Producto & "Email: " & vbCrLf & .fields("Email") & vbCrLf
       Producto = Producto & String$(CantGuion, "-") & vbCrLf _
                & "PRODUCTO/Cant x PVP/TOTAL" & vbCrLf _
                & String$(CantGuion, "-") & vbCrLf
                Efectivo = .fields("Efectivo")
       Cant_Ln = Cant_Ln + 6
   End If
  End With
 'Comenzamos a recoger los detalles de la factura
  If TFA.TC = "PV" Then
     sSQL = "SELECT DF.*,CP.Detalle,CP.Codigo_Barra " _
          & "FROM Trans_Ticket As DF,Catalogo_Productos As CP " _
          & "WHERE DF.Ticket = " & TFA.Factura & " " _
          & "AND DF.TC = '" & TFA.TC & "' " _
          & "AND DF.Item = '" & NumEmpresa & "' " _
          & "AND DF.Periodo = '" & Periodo_Contable & "' " _
          & "AND DF.Item = CP.Item " _
          & "AND DF.Periodo = CP.Periodo " _
          & "AND DF.Codigo_Inv = CP.Codigo_Inv " _
          & "ORDER BY DF.D_No "
  Else
      sSQL = "SELECT DF.*,CP.Detalle,CP.Codigo_Barra " _
           & "FROM Detalle_Factura As DF,Catalogo_Productos As CP " _
           & "WHERE DF.Factura = " & TFA.Factura & " " _
           & "AND DF.TC = '" & TFA.TC & "' " _
           & "AND DF.Item = '" & NumEmpresa & "' " _
           & "AND DF.Periodo = '" & Periodo_Contable & "' " _
           & "AND DF.Item = CP.Item " _
           & "AND DF.Periodo = CP.Periodo " _
           & "AND DF.Codigo = CP.Codigo_Inv " _
           & "ORDER BY DF.Codigo "
  End If
  Select_Adodc DtaDetalle, sSQL
  With DtaDetalle.Recordset
   If .RecordCount > 0 Then
       Do While (Not .EOF)
          Producto = Producto & .fields("Producto") & vbCrLf _
                   & SetearBlancos(CStr(.fields("Cantidad")) & "x" & Format$(.fields("Precio"), "#,##0.00"), 12, 0, False) & " " _
                   & SetearBlancos(CStr(.fields("Total")), CantGuion - 13, 0, True, , True) & vbCrLf
          Total = Total + .fields("Total")
          If TFA.TC <> "PV" Then Total_IVA = Total_IVA + .fields("Total_IVA")
          Cant_Ln = Cant_Ln + 1
         .MoveNext
       Loop
   End If
  End With
 'Pie de factura
 '===========================================================
  With DtaFactura.Recordset
   If .RecordCount > 0 Then
       If TFA.TC = "PV" Then
          SubTotal = .fields("Total")
          Total = .fields("Total")
          Total_IVA = 0
          Total_Servicio = 0
          Total_Desc = 0
       Else
          SubTotal = .fields("SubTotal")
          Total = .fields("Total_MN")
          Total_IVA = .fields("IVA")
          Total_Servicio = .fields("Servicio")
          Total_Desc = .fields("Descuento")
       End If
       Producto = Producto & String$(CantGuion, "-") & vbCrLf
       Cant_Ln = Cant_Ln + 1
       'If Total_IVA Then
       If (CantGuion - 26) > 0 Then CantBlancos = String$(CantGuion - 26, " ") Else CantBlancos = ""
          Producto = Producto _
                   & CantBlancos & "     SUBTOTAL " & SetearBlancos(CStr(SubTotal), 12, 0, True, False, True) & vbCrLf _
                   & CantBlancos & "    I.V.A " & Porc_IVA * 100 & "% " & SetearBlancos(CStr(Total_IVA), 12, 0, True, False, True) & vbCrLf
          Cant_Ln = Cant_Ln + 1
        
       If Total_Servicio > 0 Then
          Producto = Producto _
                   & CantBlancos & "     SERVICIO " & SetearBlancos(CStr(Total_Servicio), 12, 0, True, False, True) & vbCrLf
          Cant_Ln = Cant_Ln + 1
       End If
       If Total_Desc > 0 Then
          Producto = Producto _
                   & CantBlancos & "    DESCUENTO " & SetearBlancos(CStr(Total_Desc), 12, 0, True, False, True) & vbCrLf
          Cant_Ln = Cant_Ln + 1
       End If
       If TFA.TC = "PV" Then
          Producto = Producto & CantBlancos & "TOTAL TICKET  "
       ElseIf TFA.TC = "NV" Then
          Producto = Producto & CantBlancos & "TOTAL NOTA V. "
       Else
          Producto = Producto & CantBlancos & "TOTAL FACTURA "
       End If
       Producto = Producto & SetearBlancos(CStr(Total), 12, 0, True, False, True) & vbCrLf
       Producto = Producto & String$(CantGuion, "=") & vbCrLf
       If TFA.TC = "PV" Then Producto = Producto & "RECLAME SU FACTURA EN CAJA" & vbCrLf
       Producto = Producto & "  GRACIAS POR SU COMPRA " & vbCrLf & " " & vbCrLf _
                & " " & vbCrLf & " " & vbCrLf & " " & vbCrLf
       Cant_Ln = Cant_Ln + Cant_Item_PV
   End If
  End With
 'Enviamos a la Impresora
  'TipoCourier
  'TipoConsola
  'TipoCourierNew
  Printer.FontName = TipoCourierNew
  If Copia_PV Then
     If Cant_Item_PV < Cant_Ln Then Cant_Item_PV = Cant_Ln
     Cadena = ""
     Cant_Ln = Cant_Item_PV - Cant_Ln
     If Cant_Ln <= 0 Then Cant_Ln = 1
     For I = 1 To Cant_Ln
         Cadena = Cadena & "` " & vbCrLf
     Next I
     Producto = Producto & Cadena & Producto & vbCrLf & Cadena
  End If
  PrinterTexto 0.5, PosLinea, Producto
  Printer.EndDoc
End If
RatonNormal
Exit Sub
Errorhandler:
    RatonNormal
    ErrorDeImpresion
    Exit Sub
End Sub


Public Sub ImprimirFacturasHotel(NumFact As Long, DtaFactura As Data, DtaDetalle As Data, Fact As Boolean, Optional CodCliente As String)
Dim CadenaMoneda As String
'Establecemos Espacios y seteos de impresion
On Error GoTo Errorhandler
Printer.PaperSize = vbPRPSLegal     'Porte la de hoja Letter 21.6 x 27.9
LetraAnterior = Printer.FontName
Mensajes = "Imprmir Factura No. " & NumFact
Titulo = "Formulario de Impresion"
TipoDeCaja = 4 + 32: J = MsgBox(Mensajes, TipoDeCaja, Titulo)
If J = 6 Then
    RatonReloj
   If Fact Then
      sSQL = "SELECT Facturas.*,Clientes.* " _
           & "FROM Facturas,Clientes " _
           & "WHERE Factura = " & NumFact & " " _
           & "AND Clientes.Codigo = Facturas.Codigo_C "
      Select_Adodc DtaFactura, sSQL, False
      sSQL = "SELECT DF.*,A.Articulo,L.Linea,L.Codigo_L " _
           & "FROM Detalle_Factura As DF,Articulo As A,Linea As L " _
           & "WHERE DF.Codigo_P = A.Codigo " _
           & "AND A.Codigo_L = L.Codigo_L " _
           & "AND DF.Factura_No = " & NumFact & " " _
           & "ORDER BY L.Codigo_L,DF.Codigo_P,DF.Fecha "
      Select_Adodc DtaDetalle, sSQL, False
   Else
      sSQL = "SELECT * FROM Clientes " _
           & "WHERE Codigo = '" & CodCliente & "' "
      Select_Adodc DtaFactura, sSQL, False
      sSQL = "SELECT DF.*,A.Articulo,L.Linea,L.Codigo_L " _
           & "FROM Detalle_Factura As DF,Articulo As A,Linea As L " _
           & "WHERE DF.Codigo_P = A.Codigo " _
           & "AND A.Codigo_L = L.Codigo_L " _
           & "AND DF.Codigo_C = '" & CodCliente & "' " _
           & "AND T = 'P' " _
           & "ORDER BY L.Codigo_L,DF.Codigo_P,DF.Fecha "
      Select_Adodc DtaDetalle, sSQL, False
   End If

Escala_Centimetro 1, TipoTimes, 10
'Iniciamos la consulta de impresion
With DtaFactura.Recordset
If .RecordCount > 0 Then
    Dibujo = RutaSistema & "\FORMATOS\FACHOTEL.GIF"
    'MsgBox Dibujo
    PrinterPaint Dibujo, 1, 0.5, 19, 25
    msg = Format$(Comp_No, "000000")
    'Diferencia = .Fields("Total") + .Fields("Descuento") + .Fields("Comision") - .Fields("IVA")
    Printer.FontSize = 10: Printer.FontBold = True
    Mifecha = FechaStrgCorta(FechaSistema)
   'recolectamos los item de la factura a buscar
    If Fact Then
       Mifecha = FechaStrgCorta(.fields("Fecha"))
       PrinterTexto 10.5, 24.7, "Elaborado por: " & .fields("Vendedor")
    End If
    PrinterTexto 3, 3.5, Mifecha
    PrinterVariables 3.5, 4.2, .fields("Nombres") & " " & .fields("Apellidos")
    PrinterFields 3.5, 5.7, .fields("Direccion"), False
    PrinterFields 3.7, 4.9, .fields("Telefono"), False
    PrinterFields 15.5, 4.9, .fields("RUC_CI"), False
    If Fact Then
    Total = Redondear(.fields("Total_MN"))
    Total_ME = Redondear(.fields("Total_ME"))
    End If
    CadenaMoneda = "S/."
    If Total_ME > 0 Then
       Total = Total_ME
       CadenaMoneda = "USD"
    End If
    Printer.FontSize = 12
    CadenaMoneda = "USD"
    PrinterTexto 15, 20.7, CadenaMoneda
    PrinterTexto 15, 21.5, CadenaMoneda
    PrinterTexto 15, 22.3, CadenaMoneda
    PrinterTexto 15, 23.1, CadenaMoneda
    PrinterTexto 15, 23.9, CadenaMoneda
    If Fact Then
       PrinterFields 16.1, 20.7, .fields("Sin_IVA"), False
       PrinterFields 16.1, 21.5, .fields("Con_IVA"), False
       PrinterFields 17, 22.3, .fields("IVA"), False
       PrinterFields 17, 23.1, .fields("Servicio"), False
       PrinterVariables 16.2, 23.9, Total
    End If
    'PrinterLineasMayor 0.8, 25.8, Cambio_Letras(Total), 12
    PFil = 8
    Printer.FontSize = 10
    Printer.FontBold = False
    'CantFils = PrinterLineasMayor(4.2, PFil, .Fields("Nota"), 9)
End If
End With
With DtaDetalle.Recordset
 If .RecordCount > 0 Then
    .MoveFirst
     PFil = PFil + ((CantFils + 1) * 0.4)
     If Fact Then
        Printer.FontUnderline = True
        PrinterTexto 1.5, PFil, "HOSPEDAJE"
        Printer.FontUnderline = False
        PrinterFields 16.2, PFil, DtaFactura.Recordset.fields("Valor_Habit"), False
     End If
     PFil = PFil + 0.6
     CodigoL = .fields("Codigo_L")
     Printer.FontUnderline = True
     PrinterFields 1.5, PFil, .fields("Linea"), False
     Printer.FontUnderline = False
     PFil = PFil + 0.4
     CodigoL = .fields("Codigo_L")
    'comenzamos a recoger los detalles
     Do While (Not .EOF)
        If CodigoL <> .fields("Codigo_L") Then
           PFil = PFil + 0.2
           Printer.FontUnderline = True
           PrinterFields 1.5, PFil, .fields("Linea"), False
           PFil = PFil + 0.5
           CodigoL = .fields("Codigo_L")
        End If
        Printer.FontUnderline = False: Printer.FontSize = 10
        PrinterFields 2, PFil, .fields("Fecha"), False
        PrinterFields 4, PFil, .fields("Articulo"), False
        PrinterFields 11.5, PFil, .fields("Cantidad"), False
        PrinterFields 13.4, PFil, .fields("Valor_Unit"), False
        PrinterFields 16.2, PFil, .fields("Valor_Total"), False
        PFil = PFil + 0.4
       .MoveNext
    Loop
    'PrinterLineasMayor 4.2, PFil + 0.5, DtaFactura.Recordset.Fields("Observacion"), 9
 End If
End With
Printer.FontName = LetraAnterior
MensajeEncabData = ""
Printer.EndDoc 'La impresión ha terminado.
RatonNormal
End If
Exit Sub
Errorhandler:
    RatonNormal
    ErrorDeImpresion
    Exit Sub
End Sub

''Public Sub Imprimir_PRN(Inic_X As Single, Inic_Y As Single)
''Dim AdoFormato As ADODB.Recordset
''Dim AdoPrinter As ADODB.Recordset
''Dim Largo_Ancho As Single
''Dim PIF As Boolean
''Dim Pos_XX As Single
''Dim Pos_YY As Single
''Dim Tipo_Letra_PRN As String
''
''''' TipoSansSerif
''''' TipoCondensed
''''' TipoComicSans
''''' TipoConsola
''''' TipoTimes
''''' TipoCourier
''''' TipoCourierNew
''''' x TipoTerminal
''''' TipoSystem
''''' TipoArial
''''' TipoArialNarrow
''''' TipoArialBlack
''''' TipoAvantGarde
''''' TipoHelvetica
''''' TipoTahoma
''''' TipoVerdana
''
''   PIF = True
''   Printer.FontBold = True
''   Tipo_Letra_PRN = Printer.FontName
''   Printer.FontName = TipoCourier
''  'Listar el Comprobante
''   sSQL = "SELECT * " _
''        & "FROM Formato " _
''        & "WHERE TP = 'IF' " _
''        & "AND Item = '" & NumEmpresa & "' "
''   Select_AdoDB AdoFormato, sSQL
''   If AdoFormato.RecordCount > 0 Then PIF = CBool(AdoFormato.Fields("Lineas"))
''  'Listar las Transacciones
''   sSQL = "SELECT * " _
''        & "FROM Seteos_Documentos "
''   If PIF Then
''      sSQL = sSQL & "WHERE TP = 'IF' " _
''           & "AND Item = '000' "
''   Else
''      sSQL = sSQL & "WHERE TP = 'PIF' " _
''           & "AND Item = '" & NumEmpresa & "' "
''   End If
''   sSQL = sSQL & "ORDER BY TP,Campo "
''   Select_AdoDB AdoPrinter, sSQL
''  'MsgBox AdoPrinter.RecordCount
''   If AdoPrinter.RecordCount > 0 Then
''
''      Do While Not AdoPrinter.EOF
''         Printer.FontSize = AdoPrinter.fields("Porte")
''         Pos_XX = AdoPrinter.Fields("Pos_X")
''         Pos_YY = AdoPrinter.Fields("Pos_Y")
''         If Pos_XX > 0 And Pos_XX > 0 Then
''            Pos_XX = Pos_XX + Inic_X
''            Pos_YY = Pos_YY + Inic_Y
''            Largo_Ancho = Val(SinEspaciosDer(AdoPrinter.Fields("Encabezado")))
''            If TrimStrg(SinEspaciosIzq(AdoPrinter.Fields("Encabezado"))) = "LINEAH" Then
''               Imprimir_Linea_H Pos_YY, Pos_XX, Pos_XX + Largo_Ancho
''            ElseIf TrimStrg(SinEspaciosIzq(AdoPrinter.Fields("Encabezado"))) = "LINEAV" Then
''               Imprimir_Linea_V Pos_XX, Pos_YY, Pos_YY + Largo_Ancho
''            Else
''               PrinterTexto Pos_XX, Pos_YY, AdoPrinter.Fields("Encabezado")
''            End If
''         End If
''       AdoPrinter.MoveNext
''      Loop
''   End If
''   AdoPrinter.Close
''   AdoFormato.Close
''   Printer.FontBold = False
''   Printer.FontName = Tipo_Letra_PRN
''End Sub

'''Public Sub ImprimirFactTurs(DataFactura As Data, DataDetAcomp As Data, DataDetalle As Data)
''''Establecemos Espacios y seteos de impresion
'''On Error GoTo Errorhandler
'''RatonReloj
'''LetraAnterior = Printer.FontName
'''Escala_Centimetro 1, TipoTimes, 10
''''Iniciamos la consulta de impresion
'''With DataFactura.Recordset
'''If .RecordCount > 0 Then
'''    TipoFacturas = .Fields("Nota")
'''    CalcularIVA = False
'''    If .Fields("IVA") > 0 Then CalcularIVA = True
'''    If CalcularIVA Then
'''       Select Case TipoProc
'''         Case "FA": Dibujo = RutaSistema & "\FORMATOS\FACTTUR1.GIF"
'''         Case "FA1": Dibujo = RutaSistema & "\FORMATOS\FACTTUR2.GIF"
'''         Case "FA2": Dibujo = RutaSistema & "\FORMATOS\FACTTUR3.GIF"
'''         Case "FA3": Dibujo = RutaSistema & "\FORMATOS\FACTTUR4.GIF"
'''       End Select
'''       PrinterPaint Dibujo, 1.5, 11, 17, 14.5
'''       PrinterTexto 2, 10, "TIPO DE FACTURA: " & TipoFacturas
'''    Else
'''       Dibujo = RutaSistema & "\FORMATOS\NOTADEBI.GIF"
'''       PrinterPaint Dibujo, 0.8, 1.2, 19, 26
'''       Printer.FontSize = 14
'''       Printer.FontBold = True
'''       PrinterTexto 11, 4.7, "NOTA DE DEBITO"
'''       PrinterTexto 12, 5.5, "No. " & Format$(.Fields("Factura"), "000000")
'''       Printer.FontSize = 10
'''       Printer.FontBold = False
'''    End If
'''    Diferencia = .Fields("Total_MN") + .Fields("Descuento") + .Fields("Comision") - .Fields("IVA")
'''    Printer.FontSize = 10: Printer.FontBold = True
'''    Mifecha = FechaStrgCorta(.Fields("Fecha"))
'''  ' Recolectamos los item de la factura a buscar
'''    PrinterTexto 1.7, 5.9, NombreCiudad
'''    PrinterTexto 2, 25, .Fields("Vendedor")
'''    PrinterTexto 3.6, 5.9, Format$(FechaDia(.Fields("Fecha")), "00")
'''    PrinterTexto 4.6, 5.9, Format$(FechaMes(.Fields("Fecha")), "00")
'''    PrinterTexto 5.6, 5.9, Format$(FechaAnio(.Fields("Fecha")), "00")
'''    PrinterFields 3.5, 8, .Fields("Cliente"), False
'''    PrinterFields 14.5, 8, .Fields("RUC_CI"), False
'''    PrinterFields 4, 8.7, .Fields("Direccion"), False
'''    PrinterFields 14.5, 8.7, .Fields("Telefono"), False
'''    PrinterFields 15.5, 22, .Fields("SubTotal"), False
'''    PrinterFields 15.5, 22.6, .Fields("Descuento"), False
'''    PrinterFields 16, 23.2, .Fields("Comision"), False
'''    PrinterFields 16, 23.8, .Fields("IVA"), False
'''    PrinterFields 15.5, 24.4, .Fields("Total_ME"), False
'''    If .Fields("Total_ME") = 0 Then PrinterFields 15.5, 25, .Fields("Total_MN"), False
'''    PrinterFields 4, 22.6, .Fields("Forma_Pago"), False
'''    If TipoProc = "FA" Then
'''       PrinterFields 2, 11.5, .Fields("Definitivo"), False
'''       PrinterFields 8.7, 11.5, .Fields("Codigo_T"), False
'''       PrinterFields 12.6, 11.5, .Fields("Fecha_Tours"), False
'''    End If
'''End If
'''End With
'''If TipoProc = "FA" Then
'''I = Redondear(DataDetAcomp.Recordset.RecordCount / 2) - 1
'''If I < 0 Then I = 1
'''With DataDetAcomp.Recordset
''' If .RecordCount > 0 Then
'''    .MoveFirst
'''     PosLinea = 12.3: J = 1: KR = 2
'''     Do While Not .EOF
'''        PrinterFields KR, PosLinea, .Fields("Acompañante"), False
'''        PosLinea = PosLinea + 0.5
'''        If J > I Then
'''           PosLinea = 12.3
'''           KR = 10
'''           J = 0
'''        End If
'''        J = J + 1
'''       .MoveNext
'''     Loop
''' End If
'''End With
'''End If
'''If TipoProc = "FA" Then PFil = 15.3 Else PFil = 12.3
'''Printer.FontBold = False
'''With DataDetalle.Recordset
''' If .RecordCount > 0 Then
'''    .MoveFirst
'''    'comenzamos a recoger los detalles
'''     Do While (Not .EOF)
'''        Select Case TipoProc
'''          Case "FA":
'''               NumeroLineas = PrinterLineasMayor(1.7, PFil, .Fields("Producto"), 7)
'''               If NumeroLineas > 1 Then PFil = PFil + (0.4 * NumeroLineas)
'''               PrinterLineas 1.7, PFil, .Fields("Producto"), 7
'''               PrinterFields 10.5, PFil, .Fields("Cantidad"), False
'''               PrinterFields 13, PFil, .Fields("Precio"), False
'''               PrinterFields 15.5, PFil, .Fields("Total"), False
'''          Case "FA1":
'''               PrinterFields 1.6, PFil, .Fields("Ticket"), False
'''               NumeroLineas = PrinterLineasMayor(4.7, PFil, .Fields("Producto"), 5)
'''               If NumeroLineas > 1 Then PFil = PFil + (0.4 * NumeroLineas)
'''               PrinterLineas 4.7, PFil, .Fields("Producto"), 5
'''               PrinterFields 10.1, PFil, .Fields("Ruta"), False
'''               PrinterFields 13.1, PFil, .Fields("Precio"), False
'''               PrinterFields 15.6, PFil, .Fields("Total"), False
'''          Case "FA2":
'''               NumeroLineas = PrinterLineasMayor(1.6, PFil, .Fields("Producto"), 5)
'''               If NumeroLineas > 1 Then PFil = PFil + (0.4 * NumeroLineas)
'''               PrinterLineas 1.6, PFil, .Fields("Producto"), 5
'''               PrinterFields 10, PFil, .Fields("Ruta"), False
'''               PrinterFields 13, PFil, .Fields("Precio"), False
'''               PrinterFields 15.5, PFil, .Fields("Total"), False
'''          Case "FA3":
'''               NumeroLineas = PrinterLineasMayor(1.6, PFil, .Fields("Producto"), 10.5)
'''               If NumeroLineas > 1 Then PFil = PFil + (0.4 * NumeroLineas)
'''               PrinterLineas 1.6, PFil, .Fields("Producto"), 10.5
'''               PrinterFields 13, PFil, .Fields("Precio"), False
'''               PrinterFields 15.5, PFil, .Fields("Total"), False
'''        End Select
'''        PFil = PFil + 0.4
'''       .MoveNext
'''    Loop
''' End If
'''End With
'''Printer.FontName = LetraAnterior
'''RatonNormal
'''MensajeEncabData = ""
'''Printer.EndDoc  ' La impresión ha terminado.
'''Exit Sub
'''Errorhandler:
'''    RatonNormal
'''    ErrorDeImpresion
'''    Exit Sub
'''End Sub

Public Sub Imprimir_Saldo_Factura(Datas As Adodc)
On Error GoTo Errorhandler
Dim SizeLetra As Integer
Dim PDF_Nombre_Documento As String
Dim PDF_Titulo As String
Dim PDF_TipoDeLetra As String
Dim PDF_VerDocumento As Boolean
 
 PDF_Nombre_Documento = Modulo & "_" & CodigoUsuario & "_ISFA"
 PDF_Titulo = Modulo & "_" & CodigoUsuario & "_ISFA"
 PDF_TipoDeLetra = TipoArialNarrow
 PDF_VerDocumento = True
 
 Mensajes = "Seguro de Imprimir en:" & vbCrLf & Printer.DeviceName & "?"
 Titulo = "IMPRESION"
 Bandera = False
 SizeLetra = 6
 Orientacion_Pagina = 2
 SetPrinters.Show 1
 If PonImpresoraDefecto(SetNombrePRN) Then
   'Generamos el documento
    RatonReloj
    tPrint.TipoImpresion = Es_PDF
    tPrint.NombreArchivo = PDF_Nombre_Documento
    tPrint.TituloArchivo = PDF_Titulo
    tPrint.TipoLetra = PDF_TipoDeLetra
    tPrint.OrientacionPagina = Orientacion_Pagina
    tPrint.PaginaA4 = True
    tPrint.EsCampoCorto = False
    tPrint.VerDocumento = PDF_VerDocumento
    
    Set cPrint = New cImpresion
    cPrint.iniciaImpresion
   
    InicioX = 1: InicioY = 0
    'Escala_Centimetro FormaImp, TipoTimes, SizeLetra
    DataAnchoCampos InicioX, Datas, SizeLetra, TipoVerdana, Orientacion_Pagina
    Ancho(0) = 0.7     ' Cliente
    Ancho(1) = 6.2     ' T
    Ancho(2) = 6.6     ' Serie
    Ancho(3) = 7.7     ' Factura
    Ancho(4) = 9.2     ' Fecha_Emis
    Ancho(5) = 10.7    ' Total
    Ancho(6) = 12.7    ' Total_Efectivo
    Ancho(7) = 14.7    ' Total_Banco
    Ancho(8) = 16.7    ' Ret_Fuente
    Ancho(9) = 18.7    ' Ret_IVA B
    Ancho(10) = 20.7    ' Ret_IVA S
    Ancho(11) = 22.7   ' Otros_Abonos
    Ancho(12) = 24.7   ' Total_Abonos
    Ancho(13) = 26.7   ' Saldo Actual
    Ancho(14) = 28.7
    
    Pagina = 1
    'Iniciamos la impresion
    Printer.FontBold = False
    Total = 0
    Abono = 0
    Saldo = 0
    Total_ME = 0
    Abono_ME = 0
    Saldo_ME = 0
    EnDosPaginas = 0
    EncabezadoData Datas, , 14
    With Datas.Recordset
     If .RecordCount > 0 Then
        .MoveFirst
          NombreCliente = .fields("Cliente")
          If Len(NombreCliente) > 37 Then NombreCliente = MidStrg(NombreCliente, 1, 37) & "..."
          
          'Factura_No= .Fields("Factura")
          'TipoDoc = .Fields("Tipo")
          
          CodigoCli = .fields("CodigoC")
          Printer.FontSize = SizeLetra
          Printer.FontName = TipoVerdana
          PrinterTexto Ancho(0), PosLinea, NombreCliente
          'PrinterTexto Ancho(4), PosLinea, TipoDoc
          Do While Not .EOF
             If CodigoCli <> .fields("CodigoC") Then
                PosLinea = PosLinea + 0.05
                Imprimir_Linea_H PosLinea, Ancho(0), Ancho(14)
                PosLinea = PosLinea + 0.05
                Saldo_ME = Total_ME - Abono_ME
                PrinterTexto Ancho(4) + 0.5, PosLinea, "S U B T O T A L"
                PrinterVariables Ancho(6), PosLinea, Total_ME
                PrinterVariables Ancho(12), PosLinea, Abono_ME
                PrinterVariables Ancho(13), PosLinea, Saldo_ME
                PosLinea = PosLinea + 0.5
                Imprimir_Linea_H PosLinea, Ancho(0), Ancho(14)
                PosLinea = PosLinea + 0.05
                CodigoCli = .fields("CodigoC")
                NombreCliente = .fields("Cliente")
                If Len(NombreCliente) > 37 Then NombreCliente = MidStrg(NombreCliente, 1, 37) & "..."
                Total_ME = 0: Abono_ME = 0: Saldo_ME = 0
                PrinterTexto Ancho(0), PosLinea, NombreCliente
             End If
'''             If NumStrg <> .Fields("Factura") Then
'''                'Factura_No = .Fields("Factura")
'''                'TipoDoc = .Fields("Tipo")
'''                NumStrg = .Fields("Factura")
'''                Mifecha = .Fields("Fecha_Emis")
'''                PrinterTexto Ancho(2), PosLinea, NumStrg    'Format$(Factura_No, "00000000")
'''                PrinterTexto Ancho(3), PosLinea, Mifecha
'''                'PrinterTexto Ancho(4), PosLinea, TipoDoc
'''             End If
             'If TipoDoc <> .Fields("Tipo") Then
             '   TipoDoc = .Fields("Tipo")
             '   PrinterTexto Ancho(4), PosLinea, TipoDoc
             'End If
             'If .Fields("Abono") <> 0 Then
             '    PrinterFields Ancho(1), PosLinea, .Fields("Recibo")
             'End If
             PrinterTexto Ancho(1), PosLinea, .fields("T")
             PrinterTexto Ancho(2), PosLinea, .fields("Serie")
             PrinterTexto Ancho(3), PosLinea, .fields("Factura")
             PrinterFields Ancho(4), PosLinea, .fields("Fecha")
             PrinterFields Ancho(5), PosLinea, .fields("Total")
             PrinterFields Ancho(6), PosLinea, .fields("Total_Efectivo")
             PrinterFields Ancho(7), PosLinea, .fields("Total_Banco")
             PrinterFields Ancho(8), PosLinea, .fields("Total_Ret_Fuente")
             PrinterFields Ancho(9), PosLinea, .fields("Total_Ret_IVA_B")
             PrinterFields Ancho(10), PosLinea, .fields("Total_Ret_IVA_S")
             PrinterFields Ancho(11), PosLinea, .fields("Otros_Abonos")
             PrinterFields Ancho(12), PosLinea, .fields("Total_Abonos")
             PrinterFields Ancho(13), PosLinea, .fields("Saldo_Actual")
             For I = 0 To 14
                 Printer.Line (Ancho(I), PosLinea - 0.05)-(Ancho(I), PosLinea + 0.35), Negro
             Next I
             Total = Total + .fields("Total")
             Abono = Abono + .fields("Total_Abonos")
             Saldo = Saldo + .fields("Saldo_Actual")
             Total_ME = Total_ME + .fields("Total")
             Abono_ME = Abono_ME + .fields("Total_Abonos")
             Saldo_ME = Saldo_ME + .fields("Saldo_Actual")
             PosLinea = PosLinea + 0.3
             If PosLinea >= LimiteAlto Then
                PosLinea = PosLinea + 0.05
                Imprimir_Linea_H PosLinea, Ancho(0), Ancho(13)
                Printer.NewPage
                EncabezadoData Datas, , 14
                Printer.FontSize = SizeLetra
                Printer.FontName = TipoVerdana
             End If
            .MoveNext
          Loop
          PosLinea = PosLinea + 0.05
          Imprimir_Linea_H PosLinea, Ancho(0), Ancho(14)
          PosLinea = PosLinea + 0.05
          Saldo_ME = Total_ME - Abono_ME
          PrinterTexto Ancho(3) + 0.5, PosLinea, "S U B T O T A L"
          PrinterVariables Ancho(5), PosLinea, Total_ME
          PrinterVariables Ancho(11), PosLinea, Abono_ME
          PrinterVariables Ancho(12), PosLinea, Saldo_ME
          PosLinea = PosLinea + 0.5
     End If
    End With
    Imprimir_Linea_H PosLinea, Ancho(0), Ancho(14)
    PosLinea = PosLinea + 0.1
    Printer.FontBold = True
    Saldo = Total - Abono
    PrinterTexto Ancho(4) + 0.5, PosLinea, "T O T A L E S"
    PrinterVariables Ancho(6), PosLinea, Total
    PrinterVariables Ancho(12), PosLinea, Abono
    PrinterVariables Ancho(13), PosLinea, Saldo
    RatonNormal
    MensajeEncabData = ""
    Printer.EndDoc
 End If
 RatonNormal
Exit Sub
Errorhandler:
    RatonNormal
    ErrorDeImpresion
    Exit Sub
End Sub


Public Sub ImprimirContratos(Datas As Data, FinDoc As Boolean, FormaImp As Byte, SizeLetra As Integer)
On Error GoTo Errorhandler
RatonReloj
InicioX = 0.5: InicioY = 0
'Escala_Centimetro FormaImp, TipoTimes, SizeLetra
DataAnchoCampos InicioX, Datas, SizeLetra, TipoTimes, FormaImp
Ancho(0) = 0.5
Ancho(1) = 1
Ancho(2) = 5
Ancho(3) = 6.5
Ancho(4) = 8
Ancho(5) = 9.5
Ancho(6) = 15
Ancho(7) = 17.5
Ancho(8) = 20
Pagina = 1
'Iniciamos la impresion
Printer.FontBold = False
Total = 0: Saldo = 0
With Datas.Recordset
     .MoveFirst
      EncabezadoDataReporte Datas, Ancho(0), Ancho(CantCampos)
      Printer.FontSize = SizeLetra
      Do While Not .EOF
         PrinterAllFields CantCampos, PosLinea, Datas, True, False
         PosLinea = PosLinea + 0.4
         If PosLinea > LimiteAlto Then
            Imprimir_Linea_H PosLinea, Ancho(0), Ancho(CantCampos)
            Printer.NewPage
            EncabezadoDataReporte Datas, Ancho(0), Ancho(CantCampos)
            Printer.FontSize = SizeLetra
         End If
         Total = Total + .fields("Monto_MN")
         Saldo = Saldo + .fields("Monto_ME")
        .MoveNext
      Loop
End With
Imprimir_Linea_H PosLinea, Ancho(0), Ancho(CantCampos), Negro, True
PosLinea = PosLinea + 0.1
Printer.FontBold = True: Printer.FontSize = 14
PrinterTexto Ancho(0), PosLinea, "TOTAL CONTRATOS MN:"
PrinterVariables Ancho(3) + 1.5, PosLinea, Total
PosLinea = PosLinea + 0.5
PrinterTexto Ancho(0), PosLinea, "TOTAL CONTRATOS ME:"
PrinterVariables Ancho(3) + 1.5, PosLinea, Saldo
PosLinea = PosLinea + 0.5
Printer.FontBold = False: Printer.FontSize = 10
RatonNormal
MensajeEncabData = ""
If FinDoc Then Printer.EndDoc
Exit Sub
Errorhandler:
    RatonNormal
    ErrorDeImpresion
    Exit Sub
End Sub

Public Sub ImprimirSuscripciones(Datas As Data)
On Error GoTo Errorhandler
RatonReloj
InicioX = 0.5: InicioY = 0
'Escala_Centimetro FormaImp, TipoTimes, SizeLetra
DataAnchoCampos InicioX, Datas, 9, TipoTimes, 2
Ancho(0) = 0.5
Ancho(1) = 2.5
Ancho(2) = 5
Ancho(3) = 7.5
Ancho(4) = 15
Ancho(5) = 16.5
Ancho(6) = 20
Ancho(7) = 25.5
CantCampos = 7
Pagina = 1
Contador = 0
'Iniciamos la impresion
Printer.FontBold = False
With Datas.Recordset
     .MoveFirst
      Encabezado Ancho(0), Ancho(7)
      EncabSuscrip Datas
      Printer.FontSize = 8
      Contador = 0
      Codigo1 = UCaseStrg(.fields("Area"))
      Codigo2 = UCaseStrg(.fields("Ciudad"))
      Printer.FontName = TipoTimes
      Do While Not .EOF
         If Codigo1 <> UCaseStrg(.fields("Area")) Or Codigo2 <> UCaseStrg(.fields("Ciudad")) Then
            Imprimir_Linea_H PosLinea, Ancho(0), Ancho(7)
            PosLinea = PosLinea + 0.05
            Cadena = "Total en Zona: " & Codigo1 & " de " & Codigo2 & " No. " & Contador
            PrinterVariables 2, PosLinea, Cadena
            Printer.NewPage
            Encabezado Ancho(0), Ancho(7)
            EncabSuscrip Datas
            Printer.FontSize = 8
            Printer.FontName = TipoTimes
            Contador = 0
            Codigo1 = UCaseStrg(.fields("Area"))
            Codigo2 = UCaseStrg(.fields("Ciudad"))
         End If
         PrinterFields 0.5, PosLinea, .fields("Contrato_No")
         PrinterFields 1.8, PosLinea, .fields("Cliente")
         Cadena = " " & .fields("Direccion")
         PrinterVariables 8.5, PosLinea, Cadena
         Cadena = .fields("Telefono")
         PrinterVariables 18, PosLinea, Cadena
         Cadena = .fields("Desde") & " - " & .fields("Hasta")
         PrinterVariables 20, PosLinea, Cadena
         If .fields("Nuevo") Then
             Cadena = "Nueva"
         Else
             Cadena = "Renovación"
         End If
         Select Case .fields("T")
           Case "S": Cadena = "Suspendida"
           Case "C": Cadena = "Cancelada"
           Case "N": Cadena = "Pendientes"
           Case "A": Cadena = "Anulada"
           Case "R"
                    If .fields("Nuevo") Then
                        Cadena = "Nuevas"
                    Else
                        Cadena = "Renovaciones"
                    End If
         End Select
         PrinterVariables 23.5, PosLinea, Cadena
         PosLinea = PosLinea + 0.4
         Contador = Contador + 1
         If PosLinea >= LimiteAlto Then
            Imprimir_Linea_H PosLinea, Ancho(0), Ancho(7)
            Printer.NewPage
            Encabezado Ancho(0), Ancho(7)
            EncabSuscrip Datas
            Printer.FontSize = 8
            Printer.FontName = TipoTimes
         End If
        .MoveNext
      Loop
End With
Printer.FontSize = 10
Imprimir_Linea_H PosLinea, Ancho(0), Ancho(7)
PosLinea = PosLinea + 0.05
Cadena = "Total en Zona: " & Codigo1 & " de " & Codigo2 & " No. " & Contador
PrinterVariables 2, PosLinea, Cadena
PosLinea = PosLinea + 0.4
Printer.FontBold = True
Printer.FontUnderline = True
PrinterVariables 0.5, PosLinea, "OBSERVACIONES:"
MensajeEncabData = ""
Printer.FontBold = False
Printer.FontUnderline = False
Printer.EndDoc
RatonNormal
Exit Sub
Errorhandler:
    RatonNormal
    ErrorDeImpresion
    Exit Sub
End Sub

Public Sub ImprimirCtasCob(DataT As Adodc, TextoConsulta As String, Optional NoImpTodo As Boolean)
Dim NuevoDoc As Boolean
Dim EsNewClient As Boolean
Dim NuevoCliente As String
Dim TotalF As Double
Dim SaldoF As Double
Dim TotalF_ME As Double
Dim SaldoF_ME As Double
Dim Maximo As Long

On Error GoTo Errorhandler
Mensajes = "Imprimir Catera de Clientes"
Titulo = "IMPRESION"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
RatonReloj
Pagina = 1: Documento = 1
Escala_Centimetro 1, TipoTimes, 8
'Iniciamos la impresion

ReDim Ancho(10)
C = 9
Ancho(0) = 0.5  ' T
Ancho(1) = 1.3  ' Fecha
Ancho(2) = 3.1  ' Factura
Ancho(3) = 4.5  ' Total
Ancho(4) = 7    ' Total ME
Ancho(5) = 9.5  ' Abono
Ancho(6) = 12   ' Abono ME
Ancho(7) = 14.5 ' Saldo
Ancho(8) = 17   ' Saldo ME
Ancho(9) = 19.5
'============================================================
With DataT.Recordset
If .RecordCount > 0 Then
   .MoveFirst
    'EncabezadoCartera
    EsNewClient = False
    NombreCliente = .fields("Razon_Social")
    For I = 0 To .fields.Count - 1
        If .fields(I).Name = "Cliente" Then
            NombreCliente = .fields("Cliente")
            EsNewClient = True
        End If
    Next I
    Encabezado Ancho(0), Ancho(C)
    ClienteNuevo DataT
Total_Abonos = 0: Total_Saldos = 0
Saldo_ME = 0: Suma_ME = 0
Cadena = "Imprimiendo en: " & Printer.DeviceName
Maximo = .RecordCount
Contador = 0
Do While Not DataT.Recordset.EOF
  Contador = Contador + 1
  If EsNewClient Then
     If NombreCliente <> .fields("Cliente") Then NuevoDoc = True
  Else
     If NombreCliente <> .fields("Razon_Social") Then NuevoDoc = True
  End If
  If NuevoDoc Then
     Imprimir_Linea_H PosLinea, 2, Ancho(C)
     PosLinea = PosLinea + 0.1
     PrinterVariables Ancho(2), PosLinea, "T O T A L"
     PrinterVariables Ancho(3), PosLinea, Total_Abonos
     PrinterVariables Ancho(4), PosLinea, Suma_ME
     PrinterVariables Ancho(5), PosLinea, Total_Abonos - Total_Saldos
     PrinterVariables Ancho(6), PosLinea, Suma_ME - Saldo_ME
     PrinterVariables Ancho(7), PosLinea, Total_Saldos
     PrinterVariables Ancho(8), PosLinea, Saldo_ME
     PosLinea = PosLinea + 0.6
     If PosLinea >= LimiteAlto Then
        Printer.NewPage
        Encabezado Ancho(0), Ancho(C)
        ClienteNuevo DataT
        Printer.FontBold = False
     End If
     If EsNewClient Then
        NombreCliente = .fields("Cliente")
     Else
        NombreCliente = .fields("Razon_Social")
     End If
      
     ClienteNuevo DataT
     Total_Abonos = 0: Total_Saldos = 0
     Saldo_ME = 0: Suma_ME = 0
     NuevoDoc = False
  End If
  Printer.FontSize = 8
  TotalF = .fields("Total_MN")
  SaldoF = .fields("Saldo_MN")
  TotalF_ME = .fields("Total_ME")
  SaldoF_ME = .fields("Saldo_ME")
  Moneda_US = False  '.Fields("ME")
  If NoImpTodo Then
     If Moneda_US Then
        TotalF = 0
        SaldoF = 0
     Else
        TotalF_ME = 0
        SaldoF_ME = 0
     End If
  End If
  Total_Abonos = Total_Abonos + TotalF
  Total_Saldos = Total_Saldos + SaldoF
  Suma_ME = Suma_ME + TotalF_ME
  Saldo_ME = Saldo_ME + SaldoF_ME
  PrinterFields Ancho(0), PosLinea, .fields("T"), False
  PrinterFields Ancho(1), PosLinea, .fields("Fecha"), False
  PrinterFields Ancho(2), PosLinea, .fields("Factura"), False
  PrinterVariables Ancho(3), PosLinea, TotalF
  PrinterVariables Ancho(5), PosLinea, TotalF - SaldoF
  PrinterVariables Ancho(7), PosLinea, SaldoF
  PrinterVariables Ancho(4), PosLinea, TotalF_ME
  PrinterVariables Ancho(6), PosLinea, TotalF_ME - SaldoF_ME
  PrinterVariables Ancho(8), PosLinea, SaldoF_ME
  For J = 0 To C
      Printer.Line (Ancho(J), PosLinea - 0.1)-(Ancho(J), PosLinea + 0.4), Negro
  Next J
  PosLinea = PosLinea + 0.4
  If PosLinea >= LimiteAlto Then
     Printer.NewPage
     Encabezado Ancho(0), Ancho(C)
     ClienteNuevo DataT
     Printer.FontBold = False
  End If
  'MsgBox "...."
  DataT.Recordset.MoveNext
Loop
Imprimir_Linea_H PosLinea, 2, Ancho(C)
PosLinea = PosLinea + 0.1
PrinterVariables Ancho(2), PosLinea, "T O T A L"
PrinterVariables Ancho(3), PosLinea, Total_Abonos
PrinterVariables Ancho(4), PosLinea, Suma_ME
PrinterVariables Ancho(5), PosLinea, Total_Abonos - Total_Saldos
PrinterVariables Ancho(6), PosLinea, Suma_ME - Saldo_ME
PrinterVariables Ancho(7), PosLinea, Total_Saldos
PrinterVariables Ancho(8), PosLinea, Saldo_ME
PosLinea = PosLinea + 0.5
Imprimir_Linea_H PosLinea, Ancho(0), Ancho(C), Negro, True
PosLinea = PosLinea + 0.1
Printer.FontBold = False
If PosLinea + 1.5 >= LimiteAlto Then
   Printer.NewPage
   Encabezado Ancho(0), Ancho(C)
End If
Printer.FontBold = True: Printer.FontSize = 12
PrinterVariables Ancho(3) + 1.5, PosLinea, Total
PrinterTexto Ancho(0), PosLinea, "TOTAL FACTURADO S/."
PosLinea = PosLinea + 0.5
PrinterVariables Ancho(3) + 1.5, PosLinea, Total - Saldo
PrinterTexto Ancho(0), PosLinea, "TOTAL ABONADO S/."
PosLinea = PosLinea + 0.5
PrinterVariables Ancho(3) + 1.5, PosLinea, Saldo
PrinterTexto Ancho(0), PosLinea, "TOTAL POR COBRAR S/."
PosLinea = PosLinea + 0.5
Printer.FontBold = False: Printer.FontSize = 10
End If
End With
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

Public Sub Imprimir_Resumen_Cartera(Datas As Adodc, GrupoNo As String)
Dim SizeLetra As Integer
On Error GoTo Errorhandler
Mensajes = "Seguro de Imprimir en:" & vbCrLf & Printer.DeviceName & "?"
Titulo = "IMPRESION"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
RatonReloj
SizeLetra = 8
InicioX = 1: InicioY = 0
'Escala_Centimetro FormaImp, TipoTimes, SizeLetra
DataAnchoCampos InicioX, Datas, SizeLetra, TipoTimes, 1, True

Ancho(0) = 1    'Cliente
Ancho(1) = 7.5  'Telefono
Ancho(2) = 9.1  'T
Ancho(3) = 9.6  'Fecha
Ancho(4) = 11.2 'Factura
Ancho(5) = 13.2 'Total
Ancho(6) = 15.3 'Abono
Ancho(7) = 17.4 'Saldo
Ancho(8) = 19.5 'Fin
CantCampos = 8
Pagina = 1
Total = 0: Abono = 0: Saldo = 0
'Iniciamos la impresion
With Datas.Recordset
  If .RecordCount > 0 Then
     .MoveFirst
      SQLMsg2 = "Grupo No. " & GrupoNo
      Encabezado_Cartera_Resumen Datas
      Printer.FontSize = SizeLetra
      Printer.FontBold = False
      PrinterFields Ancho(0), PosLinea, .fields("Cliente")
      PrinterFields Ancho(1), PosLinea, .fields("Telefono")
      NombreCliente = .fields("Cliente")
      Do While Not .EOF
         Printer.FontSize = SizeLetra
         If NombreCliente <> .fields("Cliente") Then
            Imprimir_Linea_H PosLinea, InicioX, Ancho(CantCampos)
            PosLinea = PosLinea + 0.1
            PrinterFields Ancho(0), PosLinea, .fields("Cliente")
            PrinterFields Ancho(1), PosLinea, .fields("Telefono")
            NombreCliente = .fields("Cliente")
         End If
         PrinterFields Ancho(2), PosLinea, .fields("T")
         PrinterFields Ancho(3), PosLinea, .fields("Fecha")
         PrinterFields Ancho(4), PosLinea, .fields("Factura")
         PrinterFields Ancho(5), PosLinea, .fields("Total_MN")
         PrinterVariables Ancho(6), PosLinea, CCur(.fields("Total_MN") - .fields("Saldo_MN"))
         PrinterFields Ancho(7), PosLinea, .fields("Saldo_MN")
         Total = Total + .fields("Total_MN")
         Abono = Abono + (.fields("Total_MN") - .fields("Saldo_MN"))
         Saldo = Saldo + .fields("Saldo_MN")
         PosLinea = PosLinea + 0.35
         If PosLinea >= LimiteAlto Then
            Imprimir_Linea_H PosLinea, InicioX, Ancho(CantCampos)
            Printer.NewPage
            Encabezado_Cartera_Resumen Datas
            Printer.FontSize = SizeLetra
            Printer.FontBold = False
         End If
        .MoveNext
      Loop
     .MoveFirst
  End If
End With
Printer.FontBold = True
Imprimir_Linea_H PosLinea, InicioX, Ancho(CantCampos), Negro, True
PosLinea = PosLinea + 0.1
PrinterVariables Ancho(5), PosLinea, Total
PrinterVariables Ancho(6), PosLinea, Abono
PrinterVariables Ancho(7), PosLinea, Saldo
End If
RatonNormal
MensajeEncabData = ""
Printer.EndDoc
Exit Sub
Errorhandler:
    RatonNormal
    ErrorDeImpresion
    Exit Sub
End Sub

Public Sub Imprimir_Resumen_Cartera_Vendedor(Datas As Adodc)
Dim lenCad As Long
Dim SubTotalCliente As Currency
Dim GrupoNo As String
On Error GoTo Errorhandler
Mensajes = "Seguro de Imprimir en:" & vbCrLf & Printer.DeviceName & "?"
Titulo = "IMPRESION"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
RatonReloj
PorteLetra = 7
TipoLetra = TipoArialNarrow
InicioX = 1: InicioY = 0
'Escala_Centimetro FormaImp, TipoTimes, SizeLetra
DataAnchoCampos InicioX, Datas, PorteLetra, TipoLetra, 1, True
SQLMsg1 = ""
SQLMsg2 = ""
SQLMsg3 = ""
MensajeEncabData = "CUENTAS POR COBRAR CLIENTES POR VENDEDOR"
Ancho(0) = 1     'Cliente
Ancho(1) = 7     'Telefono
Ancho(2) = 8.5   'T
Ancho(3) = 8.8   'Fecha
Ancho(4) = 10.2  'Fecha V
Ancho(5) = 11.6  'Serie
Ancho(6) = 12.6  'Factura
Ancho(7) = 13.8  'Total
Ancho(8) = 15.2    'Abono
Ancho(9) = 16.6  'Saldo
Ancho(10) = 18   'Dias Mora
Ancho(11) = 19.1 'Chq_Posf
Ancho(12) = 20.5 'Fin
CantCampos = 11
Pagina = 1
Total = 0: Abono = 0: Saldo = 0
SubTotalCliente = 0
'Iniciamos la impresion
With Datas.Recordset
  If .RecordCount > 0 Then
     .MoveFirst
      Encabezado Ancho(0), AnchoPapel
      CodigoEjecutivo = .fields("Ejecutivo")
      NombreCliente = .fields("Cliente")
      GrupoNo = .fields("Grupo")
      SQLMsg2 = "Vendedor: " & CodigoEjecutivo
      If Len(GrupoNo) > 1 Then SQLMsg3 = GrupoNo
      Encabezado_Cartera_Resumen_Vendedor Datas
      Printer.FontSize = PorteLetra
      Printer.FontBold = False
      lenCad = Len(NombreCliente)
      Do While Printer.TextWidth(MidStrg(NombreCliente, 1, lenCad)) > 6
         lenCad = lenCad - 1
      Loop
      PrinterTexto Ancho(0), PosLinea, MidStrg(NombreCliente, 1, lenCad)
      PrinterFields Ancho(1), PosLinea, .fields("Telefono")
      Do While Not .EOF
         Printer.FontSize = PorteLetra
         If CodigoEjecutivo <> .fields("Ejecutivo") Or GrupoNo <> .fields("Grupo") Then
            Imprimir_Linea_H PosLinea, Ancho(CantCampos - 4), Ancho(CantCampos - 1)
            Imprimir_Linea_V Ancho(CantCampos - 4), PosLinea + 0.05, PosLinea + 0.45
            Imprimir_Linea_V Ancho(CantCampos - 1), PosLinea + 0.05, PosLinea + 0.45
            PosLinea = PosLinea + 0.1
            PrinterTexto Ancho(7), PosLinea, "SubTotal del Cliente"
            PrinterVariables Ancho(9), PosLinea, SubTotalCliente
            PosLinea = PosLinea + 0.4
            
            Imprimir_Linea_H PosLinea, InicioX, Ancho(CantCampos)
            PosLinea = PosLinea + 0.1
            CodigoEjecutivo = .fields("Ejecutivo")
            GrupoNo = .fields("Grupo")
            SQLMsg2 = "Vendedor: " & CodigoEjecutivo
            If Len(GrupoNo) > 1 Then SQLMsg3 = GrupoNo
            Encabezado_Cartera_Resumen_Vendedor Datas
            Printer.FontSize = PorteLetra
            Printer.FontBold = False
            NombreCliente = .fields("Cliente")
            lenCad = Len(NombreCliente)
            Do While Printer.TextWidth(MidStrg(NombreCliente, 1, lenCad)) > 6
               lenCad = lenCad - 1
            Loop
            PrinterTexto Ancho(0), PosLinea, MidStrg(NombreCliente, 1, lenCad)
            PrinterFields Ancho(1), PosLinea, .fields("Telefono")
            SubTotalCliente = 0
         End If
         
         If NombreCliente <> .fields("Cliente") Then
            If SubTotalCliente > 0 Then
                Imprimir_Linea_H PosLinea, Ancho(CantCampos - 4), Ancho(CantCampos - 1)
                Imprimir_Linea_V Ancho(CantCampos - 4), PosLinea + 0.05, PosLinea + 0.45
                Imprimir_Linea_V Ancho(CantCampos - 1), PosLinea + 0.05, PosLinea + 0.45
                PosLinea = PosLinea + 0.1
                PrinterTexto Ancho(7), PosLinea, "SubTotal del Cliente"
                PrinterVariables Ancho(9), PosLinea, SubTotalCliente
                PosLinea = PosLinea + 0.4
            End If
            Imprimir_Linea_H PosLinea, InicioX, Ancho(CantCampos)
            PosLinea = PosLinea + 0.1
            NombreCliente = .fields("Cliente")
            lenCad = Len(NombreCliente)
            Do While Printer.TextWidth(MidStrg(NombreCliente, 1, lenCad)) > 6
               lenCad = lenCad - 1
            Loop
            PrinterTexto Ancho(0), PosLinea, MidStrg(NombreCliente, 1, lenCad)
            PrinterFields Ancho(1), PosLinea, .fields("Telefono")
            SubTotalCliente = 0
         End If
         PrinterFields Ancho(2), PosLinea, .fields("T")
         PrinterFields Ancho(3), PosLinea, .fields("Fecha")
         PrinterFields Ancho(4), PosLinea, .fields("Fecha_V")
         PrinterFields Ancho(5), PosLinea, .fields("Serie")
         PrinterFields Ancho(6), PosLinea, .fields("Factura")
         PrinterFields Ancho(7), PosLinea, .fields("Total_MN")
         PrinterVariables Ancho(8), PosLinea, CCur(.fields("Total_MN") - .fields("Saldo_MN"))
         PrinterFields Ancho(9), PosLinea, .fields("Saldo_MN")
         
         If .fields("Dias_De_Mora") < 0 Then Printer.ForeColor = Rojo
         
         PrinterFields Ancho(10), PosLinea, .fields("Dias_De_Mora")
         Printer.ForeColor = Negro
         
         PrinterFields Ancho(11), PosLinea, .fields("Chq_Posf")
         
         Total = Total + .fields("Total_MN")
         Abono = Abono + (.fields("Total_MN") - .fields("Saldo_MN"))
         Saldo = Saldo + .fields("Saldo_MN")
         
         SubTotalCliente = SubTotalCliente + .fields("Saldo_MN")
         PosLinea = PosLinea + 0.35
         If PosLinea >= LimiteAlto Then
            Imprimir_Linea_H PosLinea, InicioX, Ancho(CantCampos)
            Printer.NewPage
            InicioX = 1: InicioY = 0
            Encabezado Ancho(0), AnchoPapel
            CodigoEjecutivo = .fields("Ejecutivo")
            NombreCliente = .fields("Cliente")
            GrupoNo = .fields("Grupo")
            SQLMsg2 = "Vendedor: " & CodigoEjecutivo
            If Len(GrupoNo) > 1 Then SQLMsg3 = GrupoNo
            Encabezado_Cartera_Resumen_Vendedor Datas
            Printer.FontSize = PorteLetra
            Printer.FontBold = False
            lenCad = Len(NombreCliente)
            Do While Printer.TextWidth(MidStrg(NombreCliente, 1, lenCad)) > 6
               lenCad = lenCad - 1
            Loop
            PrinterTexto Ancho(0), PosLinea, MidStrg(NombreCliente, 1, lenCad)
            PrinterFields Ancho(1), PosLinea, .fields("Telefono")
         End If
        .MoveNext
      Loop
     .MoveFirst
  End If
End With
Imprimir_Linea_H PosLinea, Ancho(CantCampos - 4), Ancho(CantCampos - 1)
Imprimir_Linea_V Ancho(CantCampos - 4), PosLinea + 0.05, PosLinea + 0.45
Imprimir_Linea_V Ancho(CantCampos - 1), PosLinea + 0.05, PosLinea + 0.45

PosLinea = PosLinea + 0.1
PrinterTexto Ancho(7), PosLinea, "SubTotal Cliente"
PrinterVariables Ancho(9), PosLinea, SubTotalCliente
PosLinea = PosLinea + 0.4
Printer.FontBold = True
Imprimir_Linea_H PosLinea, InicioX, Ancho(CantCampos), Negro, True
PosLinea = PosLinea + 0.1
PrinterTexto Ancho(4), PosLinea, "Totales de Clientes"
PrinterVariables Ancho(6), PosLinea, Total
PrinterVariables Ancho(7), PosLinea, Abono
PrinterVariables Ancho(8), PosLinea, Saldo
End If
RatonNormal
MensajeEncabData = ""
Printer.EndDoc
Exit Sub
Errorhandler:
    RatonNormal
    ErrorDeImpresion
    Exit Sub
End Sub

Public Sub ImprimirHistCtasCob(DataT As Data, TextoConsulta As String)
Dim NuevoDoc As Boolean
Dim NuevoCliente As String
Dim TotalF As Double
Dim SaldoF As Double
Dim TotalF_ME As Double
Dim SaldoF_ME As Double
Dim Maximo As Long

On Error GoTo Errorhandler
RatonReloj
Pagina = 1: Documento = 1
InicioX = 1
DataAnchoCampos InicioX, DataT, 9, TipoTimes, 1
'Iniciamos la impresion
ReDim Ancho(10)
C = 9
Ancho(0) = 0.5  'T
Ancho(1) = 1.3  'Fecha
Ancho(2) = 3.6  'Factura
Ancho(3) = 5    'Total
Ancho(4) = 7.5  'Total ME
Ancho(5) = 10   'Abono
Ancho(6) = 12.5 'Abono ME
Ancho(7) = 15   'Saldo
Ancho(8) = 17.5 'Saldo ME
Ancho(9) = 20
'============================================================
Debe = 0
Haber = 0
Debe_ME = 0
Haber_ME = 0
With DataT.Recordset
If .RecordCount > 0 Then
   .MoveFirst
    Encabezado Ancho(0), Ancho(C)
    Printer.FontSize = 10
    PrinterTexto Ancho(0), PosLinea, SQLMsg2
    PosLinea = PosLinea + 0.5
    NuevoCliente = .fields("Cliente")
    ClienteNuevo DataT
Total_Abonos = 0: Total_Saldos = 0
Saldo_ME = 0: Suma_ME = 0: Factura_No = 0
Cadena = "Imprimiendo en: " & Printer.DeviceName
Maximo = .RecordCount
Contador = 0
Do While Not DataT.Recordset.EOF
  Contador = Contador + 1
  If NuevoCliente <> .fields("Cliente") Then
     Imprimir_Linea_H PosLinea, 1, Ancho(C)
     PosLinea = PosLinea + 0.1
     PrinterVariables Ancho(2), PosLinea, "T O T A L"
     PrinterVariables Ancho(3), PosLinea, Total_Abonos
     PrinterVariables Ancho(4), PosLinea, Suma_ME
     PrinterVariables Ancho(5), PosLinea, Total_Abonos - Total_Saldos
     PrinterVariables Ancho(6), PosLinea, Suma_ME - Saldo_ME
     PrinterVariables Ancho(7), PosLinea, Total_Saldos
     PrinterVariables Ancho(8), PosLinea, Saldo_ME
     PosLinea = PosLinea + 0.6
     If PosLinea >= LimiteAlto Then
        Printer.NewPage
        Encabezado Ancho(0), Ancho(C)
        Printer.FontSize = 10
        PrinterTexto Ancho(0), PosLinea, SQLMsg2
        PosLinea = PosLinea + 0.5
        ClienteNuevo DataT
        Printer.FontBold = False
     End If
     NuevoCliente = .fields("Cliente")
     ClienteNuevo DataT
     Total_Abonos = 0: Total_Saldos = 0
     Saldo_ME = 0: Suma_ME = 0
  End If
  Printer.FontSize = 8
  TotalF = .fields("Monto_MN")
  TotalF_ME = .fields("Monto_ME")
  If Factura_No = .fields("Factura") Then
     Factura_No = 0
     TotalF = 0
     TotalF_ME = 0
  Else
     Factura_No = .fields("Factura")
  End If
  Debe = Debe + TotalF
  Haber = Haber + .fields("Saldo_MN")
  SaldoF = .fields("Saldo_MN")
  SaldoF_ME = .fields("Saldo_ME")
  Total_Abonos = Total_Abonos + TotalF
  Suma_ME = Suma_ME + TotalF_ME
  Total_Saldos = Total_Saldos + SaldoF
  Saldo_ME = Saldo_ME + SaldoF_ME
  If Factura_No <> 0 Then
     Imprimir_Linea_H PosLinea, 1, Ancho(C)
     PosLinea = PosLinea + 0.1
  End If
  PrinterFields Ancho(0), PosLinea, .fields("T"), False
  PrinterFields Ancho(1), PosLinea, .fields("Fecha"), False
  PrinterVariables Ancho(2), PosLinea, Factura_No
  PrinterVariables Ancho(3), PosLinea, TotalF
  PrinterVariables Ancho(4), PosLinea, TotalF_ME
  PrinterFields Ancho(5), PosLinea, .fields("Abonos_MN"), False
  PrinterFields Ancho(6), PosLinea, .fields("Abonos_ME"), False
  PrinterFields Ancho(7), PosLinea, .fields("Saldo_MN"), False
  PrinterFields Ancho(8), PosLinea, .fields("Saldo_ME"), False
  For J = 0 To C
      Printer.Line (Ancho(J), PosLinea - 0.1)-(Ancho(J), PosLinea + 0.4), Negro
  Next J
  PosLinea = PosLinea + 0.4
  If PosLinea >= LimiteAlto Then
     Printer.NewPage
     Encabezado Ancho(0), Ancho(C)
     Printer.FontSize = 10
     PrinterTexto Ancho(0), PosLinea, SQLMsg2
     PosLinea = PosLinea + 0.5
     ClienteNuevo DataT
     Printer.FontBold = False
  End If
  DataT.Recordset.MoveNext
Loop
Imprimir_Linea_H PosLinea, 1, Ancho(C)
PosLinea = PosLinea + 0.1
PrinterVariables Ancho(2), PosLinea, "T O T A L"
PrinterVariables Ancho(3), PosLinea, Total_Abonos
PrinterVariables Ancho(4), PosLinea, Suma_ME
PrinterVariables Ancho(5), PosLinea, Total_Abonos - Total_Saldos
PrinterVariables Ancho(6), PosLinea, Suma_ME - Saldo_ME
PrinterVariables Ancho(7), PosLinea, Total_Saldos
PrinterVariables Ancho(8), PosLinea, Saldo_ME
PosLinea = PosLinea + 0.5
Imprimir_Linea_H PosLinea, Ancho(0), Ancho(CantCampos), Negro, True
PosLinea = PosLinea + 0.5
Printer.FontBold = False
If PosLinea + 1.5 > LimiteAlto Then
   Printer.NewPage
   Encabezado Ancho(0) - 1, Ancho(6) + 2
End If
Printer.FontBold = True: Printer.FontSize = 12
'PrinterTexto Ancho(0), PosLinea, "TOTAL FACTURADO  S/."
'PrinterVariables Ancho(3), PosLinea, Debe
'PosLinea = PosLinea + 0.5
'PrinterTexto Ancho(0), PosLinea, "TOTAL ABONADO    S/."
'PrinterVariables Ancho(3), PosLinea, Debe - Haber
'PosLinea = PosLinea + 0.5
'PrinterTexto Ancho(0), PosLinea, "TOTAL POR COBRAR S/."
'PrinterVariables Ancho(3), PosLinea, Haber
'PosLinea = PosLinea + 0.5
Printer.FontBold = False: Printer.FontSize = 10
End If
End With
RatonNormal
MensajeEncabData = ""
Printer.EndDoc
Exit Sub
Errorhandler:
    RatonNormal
    ErrorDeImpresion
    Exit Sub
End Sub

Public Sub ImprimirHistorial(Datas As Data, FinDoc As Boolean, FormaImp As Byte, SizeLetra As Integer)
On Error GoTo Errorhandler
RatonReloj
InicioX = 0.5: InicioY = 0
'Escala_Centimetro   FormaImp, TipoTimes, SizeLetra
DataAnchoCampos InicioX, Datas, SizeLetra, TipoTimes, FormaImp
Pagina = 1
Total = 0
'Iniciamos la impresion
Printer.FontBold = False
With Datas.Recordset
     .MoveFirst
      EncabezadoData Datas
      Printer.FontSize = SizeLetra
      Factura_No = .fields("Factura")
      Do While Not .EOF
         If Factura_No <> .fields("Factura") Then
            Imprimir_Linea_H PosLinea, Ancho(0), Ancho(CantCampos)
            PosLinea = PosLinea + 0.1
            Factura_No = .fields("Factura")
         End If
         PrinterAllFields CantCampos, PosLinea, Datas, True, False
         PosLinea = PosLinea + 0.4
         If PosLinea > LimiteAlto Then
            Imprimir_Linea_H PosLinea, Ancho(0), Ancho(CantCampos)
            Printer.NewPage
            EncabezadoData Datas
            Printer.FontSize = SizeLetra
         End If
         Total = Total + .fields("Saldo")
        .MoveNext
      Loop
End With
Imprimir_Linea_H PosLinea, InicioX, Ancho(CantCampos), Negro, True
PosLinea = PosLinea + 0.1
PrinterTexto Ancho(CantCampos - 2), PosLinea, "TOTALES"
PrinterVariables Ancho(CantCampos - 1), PosLinea, Total
RatonNormal
MensajeEncabData = ""
If FinDoc Then Printer.EndDoc
Exit Sub
Errorhandler:
    RatonNormal
    ErrorDeImpresion
    Exit Sub
End Sub

Public Sub Imprimir_Diario_Caja(AdoVentas As Adodc, _
                                AdoCxC As Adodc, _
                                AdoInv As Adodc, _
                                AdoProd As Adodc, _
                                AdoAnt As Adodc, _
                                MFechaI As String, _
                                MFechaF As String)
Dim Tipo_Letra As String
Dim OpcDeAbonos As Boolean
Dim OpcDeVentas As Boolean
Dim OpcDeInventario As Boolean
Dim FechaI As Long
Dim FechaF As Long

FechaI = CFechaLong(MFechaI)
FechaF = CFechaLong(MFechaF)

On Error GoTo Errorhandler
Tipo_Letra = TipoArialNarrow
OpcDeAbonos = True
OpcDeVentas = True
OpcDeInventario = True
Orientacion_Pagina = 1
Mensajes = "Imprimir Cierre del Diario de Caja en:" & vbCrLf & Printer.DeviceName & "?"
Titulo = "IMPRESION"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
RatonReloj
Pagina = 1: InicioX = 0.5: InicioY = 0
Escala_Centimetro 1, Tipo_Letra, 9
CantCampos = 6
ReDim Ancho(CantCampos + 1) As Single
Ancho(0) = 1.5   'Fecha
Ancho(1) = 3     'Cliente
Ancho(2) = 9.5   'Factura
Ancho(3) = 11    'IVA
Ancho(4) = 13    'Descuento
Ancho(5) = 15    'Servicio
Ancho(6) = 17    'Total MN
Ancho(7) = 19    'Fin
'Iniciamos la impresion
Total = 0
TotalIngreso = 0
Contador = 0
MensajeEncabData = "CIERRE FLUJO DE CAJA FACTURACION"
Encabezado Ancho(0), Ancho(7)
With AdoCxC.Recordset
 If .RecordCount > 0 Then
     Mensajes = "Imprimir Abonos"
     Titulo = "Pregunta de Impresion"
     If BoxMensaje = vbYes Then
    .MoveFirst
     OpcDeAbonos = False
     Printer.FontSize = 12
     Printer.FontName = Tipo_Letra
     If FechaI = FechaF Then
        PrinterTexto Ancho(0), PosLinea, "ABONOS DEL DIA " & FechaStrgCorta(MFechaI)
     Else
        PrinterTexto Ancho(0), PosLinea, "ABONO DEL " & FechaStrgCorta(MFechaI) & " AL " & FechaStrgCorta(MFechaF)
     End If
     PosLinea = PosLinea + 0.5
     Encabezado_Diario_Caja_CxC
     Printer.FontSize = 8
     Mifecha = .fields("Fecha")
     PrinterTexto Ancho(0), PosLinea, Mifecha
     PosLinea = PosLinea + 0.35
     NombreBanco = .fields("Banco")
     Cta = .fields("Cta")
     Do While Not .EOF
        If Cta <> .fields("Cta") Then
        'If NombreBanco <> .Fields("Banco") Then
           Imprimir_Linea_H PosLinea, Ancho(0), Ancho(7)
           PosLinea = PosLinea + 0.1
           Printer.FontBold = True
           Codigo = Leer_Cta_Catalogo(Cta)
           PrinterTexto Ancho(3), PosLinea, UCaseStrg("TOTAL DE " & Cuenta)
           PrinterVariables Ancho(6), PosLinea, Total
           PosLinea = PosLinea + 0.5
           Encabezado_Diario_Caja_CxC
           NombreBanco = .fields("Banco")
           Cta = .fields("Cta")
           Mifecha = .fields("Fecha")
           Printer.FontBold = False
           Printer.FontSize = 8
           PrinterTexto Ancho(0), PosLinea, Mifecha
           PosLinea = PosLinea + 0.35
           TotalIngreso = TotalIngreso + Total
           Total = 0
           Printer.FontBold = False
        End If
        Printer.FontBold = False
        Printer.FontSize = 8
        If Mifecha <> .fields("Fecha") Then
           Mifecha = .fields("Fecha")
           PrinterTexto Ancho(0), PosLinea, Mifecha
           PosLinea = PosLinea + 0.35
        End If
        Contador = Contador + 1
        If Contador > 255 Then Contador = 1
        PrinterVariables Ancho(0), PosLinea, CByte(Contador)
        PrinterTexto Ancho(0) + 0.6, PosLinea, ".- " & .fields("Cliente")
        PrinterTexto Ancho(2) - 0.1, PosLinea, Space(120)
        PrinterTexto Ancho(2), PosLinea, Format$(.fields("Factura"), "0000000")
        PrinterFields Ancho(3), PosLinea, .fields("Banco")
        PrinterFields Ancho(5), PosLinea, .fields("Cheque")
        PrinterFields Ancho(6), PosLinea, .fields("Abono")
        Total = Total + .fields("Abono")
        PosLinea = PosLinea + 0.35
        If PosLinea >= LimiteAlto Then
           Imprimir_Linea_H PosLinea, Ancho(0), Ancho(7)
           Printer.NewPage
           Encabezado Ancho(0), Ancho(7)
           Printer.FontName = Tipo_Letra
           Printer.FontSize = 12
           If FechaI = FechaF Then
              PrinterTexto Ancho(0), PosLinea, "ABONOS DEL DIA " & FechaStrgCorta(MFechaI)
           Else
              PrinterTexto Ancho(0), PosLinea, "ABONO DEL " & FechaStrgCorta(MFechaI) & " AL " & FechaStrgCorta(MFechaF)
           End If
           PosLinea = PosLinea + 0.5
           Encabezado_Diario_Caja_CxC
        End If
       .MoveNext
     Loop
     Imprimir_Linea_H PosLinea, Ancho(0), Ancho(7)
     PosLinea = PosLinea + 0.1
     TotalIngreso = TotalIngreso + Total
     Codigo = Leer_Cta_Catalogo(Cta)
     Printer.FontBold = True
     PrinterTexto Ancho(3), PosLinea, UCaseStrg("TOTAL DE " & Cuenta)
     PrinterVariables Ancho(6), PosLinea, Total
     PosLinea = PosLinea + 0.4
     Imprimir_Linea_H PosLinea, Ancho(3), Ancho(7)
     PosLinea = PosLinea + 0.1
     PrinterTexto Ancho(3), PosLinea, "T O T A L   R E C A U D A D O"
     PrinterVariables Ancho(6), PosLinea, TotalIngreso
     PosLinea = PosLinea + 0.5
     Printer.FontBold = False
     End If
 End If
End With
Printer.FontName = Tipo_Letra
Total = 0
Total_IVA = 0
Total_Desc = 0
Total_Servicio = 0
With AdoVentas.Recordset
 If .RecordCount > 0 Then
     OpcDeVentas = False
     Mensajes = "Imprimir Ventas"
     Titulo = "Pregunta de Impresion"
     If BoxMensaje = vbYes Then
    .MoveFirst
    'If PosLinea >= LimiteAlto Then
    '  Imprimir_Linea_H PosLinea, Ancho(0), Ancho(7)
        Printer.NewPage
        Encabezado Ancho(0), Ancho(7)
        Printer.FontName = Tipo_Letra
    'End If
     Printer.FontSize = 12
     If FechaI = FechaF Then
        PrinterTexto Ancho(0), PosLinea, "VENTAS DEL DIA " & FechaStrgCorta(MFechaI)
     Else
        PrinterTexto Ancho(0), PosLinea, "VENTAS DEL " & FechaStrgCorta(MFechaI) & " AL " & FechaStrgCorta(MFechaF)
     End If
     PosLinea = PosLinea + 0.5
     Encabezado_Diario_Caja_Ventas
     Printer.FontBold = False
     Printer.FontSize = 8
     Mifecha = .fields("Fecha")
     PrinterTexto Ancho(0), PosLinea, Mifecha
     PosLinea = PosLinea + 0.35
     Do While Not .EOF
        Printer.FontBold = False
        Printer.FontSize = 8
        If Mifecha <> .fields("Fecha") Then
           Mifecha = .fields("Fecha")
           PrinterTexto Ancho(0), PosLinea, Mifecha
           PosLinea = PosLinea + 0.35
        End If
        PrinterFields Ancho(0) + 0.5, PosLinea, .fields("Cliente")
        PrinterTexto Ancho(2) - 0.1, PosLinea, Space(120)
        PrinterTexto Ancho(2), PosLinea, Format$(.fields("Factura"), "0000000")
        PrinterFields Ancho(3), PosLinea, .fields("Total_IVA")
        PrinterFields Ancho(4), PosLinea, .fields("Descuento")
        PrinterFields Ancho(5), PosLinea, .fields("Servicio")
        PrinterFields Ancho(6), PosLinea, .fields("Total_MN")
        Total_IVA = Total_IVA + .fields("Total_IVA")
        Total_Desc = Total_Desc + .fields("Descuento")
        Total_Servicio = Total_Servicio + .fields("Servicio")
        Total = Total + .fields("Total_MN")
        PosLinea = PosLinea + 0.35
        If PosLinea >= LimiteAlto Then
           Imprimir_Linea_H PosLinea, Ancho(0), Ancho(7)
           Printer.NewPage
           Encabezado Ancho(0), Ancho(7)
           Printer.FontName = Tipo_Letra
           Printer.FontSize = 12
           If FechaI = FechaF Then
              PrinterTexto Ancho(0), PosLinea, "VENTAS DEL DIA " & FechaStrgCorta(MFechaI)
           Else
              PrinterTexto Ancho(0), PosLinea, "VENTAS DEL " & FechaStrgCorta(MFechaI) & " AL " & FechaStrgCorta(MFechaF)
           End If
           PosLinea = PosLinea + 0.5
           Encabezado_Diario_Caja_Ventas
        End If
       .MoveNext
     Loop
     Imprimir_Linea_H PosLinea, Ancho(0), Ancho(7)
     Printer.FontBold = True
     PosLinea = PosLinea + 0.1
     PrinterTexto Ancho(2), PosLinea, "T O T A L"
     PrinterVariables Ancho(3), PosLinea, Total_IVA
     PrinterVariables Ancho(4), PosLinea, Total_Desc
     PrinterVariables Ancho(5), PosLinea, Total_Servicio
     PrinterVariables Ancho(6), PosLinea, Total
     PosLinea = PosLinea + 0.5
     Printer.FontBold = False
     End If
 End If
End With
Printer.FontName = Tipo_Letra
With AdoInv.Recordset
 If .RecordCount > 0 Then
     OpcDeInventario = False
     Mensajes = "Imprimir Salida de Inventario"
     Titulo = "Pregunta de Impresion"
     If BoxMensaje = vbYes Then
    .MoveFirst
     If PosLinea >= LimiteAlto Then
        PosLinea = PosLinea + 0.05
        Imprimir_Linea_H PosLinea, Ancho(0), Ancho(7)
        Printer.NewPage
        Encabezado Ancho(0), Ancho(7)
        Printer.FontName = Tipo_Letra
     End If
     Printer.FontSize = 12
     If FechaI = FechaF Then
        PrinterTexto Ancho(0), PosLinea, "SALIDA DE INVENTARIO DEL DIA " & FechaStrgCorta(MFechaI)
     Else
        PrinterTexto Ancho(0), PosLinea, "SALIDA DE INVENTARIO DEL " & FechaStrgCorta(MFechaI) & " AL " & FechaStrgCorta(MFechaF)
     End If
     PosLinea = PosLinea + 0.5
     Encabezado_Diario_Caja_Inv
     Do While Not .EOF
        Printer.FontBold = False
        Printer.FontSize = 8
        PrinterFields Ancho(1), PosLinea, .fields("CODIGO_INV")
        PrinterFields Ancho(2), PosLinea, .fields("PRODUCTO")
        PrinterFields Ancho(5), PosLinea, .fields("CANT_ES")
        PosLinea = PosLinea + 0.35
        If PosLinea >= LimiteAlto Then
           Imprimir_Linea_H PosLinea, Ancho(0), Ancho(7)
           Printer.NewPage
           Encabezado Ancho(0), Ancho(7)
           Printer.FontName = Tipo_Letra
           Printer.FontSize = 12
           If FechaI = FechaF Then
              PrinterTexto Ancho(0), PosLinea, "SALIDA DE INVENTARIO DEL DIA " & FechaStrgCorta(MFechaI)
           Else
              PrinterTexto Ancho(0), PosLinea, "SALIDA DE INVENTARIO DEL " & FechaStrgCorta(MFechaI) & " AL " & FechaStrgCorta(MFechaF)
           End If
           PosLinea = PosLinea + 0.5
           Encabezado_Diario_Caja_Inv
        End If
       .MoveNext
     Loop
     PosLinea = PosLinea + 0.05
     Imprimir_Linea_H PosLinea, Ancho(0), Ancho(7)
     PosLinea = PosLinea + 0.05
     End If
 End If
End With

With AdoProd.Recordset
 If .RecordCount > 0 Then
     OpcDeInventario = False
     Mensajes = "Imprimir Rubros Facturados"
     Titulo = "Pregunta de Impresion"
     If BoxMensaje = vbYes Then
    .MoveFirst
     If PosLinea >= LimiteAlto Then
        PosLinea = PosLinea + 0.05
        Imprimir_Linea_H PosLinea, Ancho(0), Ancho(7)
        Printer.NewPage
        Encabezado Ancho(0), Ancho(7)
        Printer.FontName = Tipo_Letra
     End If
     Printer.FontSize = 12
     If FechaI = FechaF Then
        PrinterTexto Ancho(0), PosLinea, "PRODUCTOS FACTURADOS DEL DIA " & FechaStrgCorta(MFechaI)
     Else
        PrinterTexto Ancho(0), PosLinea, "PRODUCTOS FACTURADOS DEL " & FechaStrgCorta(MFechaI) & " AL " & FechaStrgCorta(MFechaF)
     End If
     PosLinea = PosLinea + 0.5
     Encabezado_Diario_Caja_Prod
     Do While Not .EOF
        Printer.FontBold = False
        Printer.FontSize = 8
        PrinterFields Ancho(1), PosLinea, .fields("Producto")
        PrinterFields Ancho(2), PosLinea, .fields("Codigo_Inv")
        PrinterFields Ancho(5), PosLinea, .fields("CANTIDADES")
        PrinterFields Ancho(6), PosLinea, .fields("TOTALES")
        PosLinea = PosLinea + 0.35
        If PosLinea >= LimiteAlto Then
           Imprimir_Linea_H PosLinea, Ancho(0), Ancho(7)
           Printer.NewPage
           Encabezado Ancho(0), Ancho(7)
           Printer.FontName = Tipo_Letra
           Printer.FontSize = 12
           If FechaI = FechaF Then
              PrinterTexto Ancho(0), PosLinea, "PRODUCTOS FACTURADOS DEL DIA " & FechaStrgCorta(MFechaI)
           Else
              PrinterTexto Ancho(0), PosLinea, "PRODUCTOS FACTURADOS DEL " & FechaStrgCorta(MFechaI) & " AL " & FechaStrgCorta(MFechaF)
           End If
           PosLinea = PosLinea + 0.5
           Encabezado_Diario_Caja_Prod
        End If
       .MoveNext
     Loop
     Imprimir_Linea_H PosLinea, Ancho(0), Ancho(7)
     End If
 End If
End With
PosLinea = PosLinea + 1
Printer.FontSize = 8
If Nombre_Cajero <> Ninguno Then
   PrinterTexto Ancho(0), PosLinea, UCaseStrg(Nombre_Cajero)
Else
   PrinterTexto Ancho(0), PosLinea, "RESPONSABLE"
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

Public Sub Imprimir_Diario_Caja_Resumen(AdoVentas As Adodc, _
                                        AdoCxC As Adodc, _
                                        AdoInv As Adodc, _
                                        AdoProd As Adodc, _
                                        AdoAnt As Adodc, _
                                        MFechaI As String, _
                                        MFechaF As String)
Dim Tipo_Letra As String
Dim Porte_Letra As Integer
Dim Pos_X_Mid As Single
Dim Pos_Y_Mid As Single
Dim Derecho As Boolean
Dim FechaI As Long
Dim FechaF As Long

FechaI = CFechaLong(MFechaI)
FechaF = CFechaLong(MFechaF)

On Error GoTo Errorhandler
Tipo_Letra = TipoArialNarrow
Orientacion_Pagina = 2
Mensajes = "Imprimir Cierre del Diario de Caja en: " & vbCrLf & Printer.DeviceName & "?"
Titulo = "IMPRESION"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
  Progreso_Barra.Mensaje_Box = "Imprimiendo el Cierre de Caja..."
  Progreso_Iniciar True
  Progreso_Barra.Incremento = 0
  
RatonReloj
Porte_Letra = 7
Pagina = 1: InicioX = 0.5: InicioY = 0
Escala_Centimetro 2, Tipo_Letra, Porte_Letra

'MsgBox SetPapelAncho & vbCrLf & SetPapelLargo & vbCrLf & vbCrLf & LimiteAncho & vbCrLf & LimiteAlto
Pos_X_Mid = (SetPapelAncho / 2) - 0.5
CantCampos = 6
ReDim Ancho(CantCampos + 1) As Single
Ancho(7) = Pos_X_Mid        'Fin
Pos_X_Mid = Pos_X_Mid - 1.8
Ancho(6) = Pos_X_Mid        'Total MN / Cheq o Dep
Pos_X_Mid = Pos_X_Mid - 1.8
Ancho(5) = Pos_X_Mid        'Servicio / Detalle del Abono
Pos_X_Mid = Pos_X_Mid - 1.8
Ancho(4) = Pos_X_Mid        'Descuento / Detalle del Abono
Pos_X_Mid = Pos_X_Mid - 1.8
Ancho(3) = Pos_X_Mid        'I.V.A
Pos_X_Mid = Pos_X_Mid - 1.3
Ancho(2) = Pos_X_Mid        'Factura
Pos_X_Mid = Pos_X_Mid - 1.5
Ancho(1) = 1                'Cliente
Ancho(0) = 0.5              'Fecha
Derecho = False
'Iniciamos la impresion
Total = 0
TotalIngreso = 0
Contador = 0
LimiteAlto = LimiteAlto - 0.5
'MsgBox LimiteAlto
MensajeEncabData = "CIERRE FLUJO DE CAJA FACTURACION"
If FechaI = FechaF Then
   SQLMsg1 = "ABONOS DEL DIA " & FechaStrgCorta(MFechaI)
Else
   SQLMsg1 = "ABONO DEL " & FechaStrgCorta(MFechaI) & " AL " & FechaStrgCorta(MFechaF)
End If
Encabezado Ancho(0), LimiteAncho
Pos_Y_Mid = PosLinea
Printer.FontName = Tipo_Letra
Printer.FontSize = Porte_Letra
Derecho = False
Pos_X_Mid = (SetPapelAncho / 2)
Imprimir_Linea_V Pos_X_Mid, PosLinea, LimiteAlto
With AdoCxC.Recordset
 If .RecordCount > 0 Then
     Progreso_Barra.Valor_Maximo = Progreso_Barra.Valor_Maximo + .RecordCount
     Progreso_Esperar
    .MoveFirst
     If Derecho Then
        Pos_X_Mid = (SetPapelAncho / 2)
     Else
        Pos_X_Mid = 0
     End If
     PrinterTexto Ancho(0), PosLinea, SQLMsg1
     PosLinea = PosLinea + 0.4
     Encabezado_Diario_Caja_CxC Pos_X_Mid, Porte_Letra
     Printer.FontBold = False
     Mifecha = .fields("Fecha")
     PrinterTexto Ancho(0), PosLinea, Mifecha
     PosLinea = PosLinea + 0.35
     NombreBanco = .fields("Banco")
     Cta = .fields("Cta")
     Do While Not .EOF
        Progreso_Esperar
        If Derecho Then
           Pos_X_Mid = (SetPapelAncho / 2)
        Else
           Pos_X_Mid = 0
        End If
        Printer.FontBold = False
        Printer.FontSize = Porte_Letra
        If Cta <> .fields("Cta") Then
           Imprimir_Linea_H PosLinea, Ancho(0) + Pos_X_Mid, Ancho(7) + Pos_X_Mid
           PosLinea = PosLinea + 0.05
           Printer.FontBold = True
           Codigo = Leer_Cta_Catalogo(Cta)
           PrinterVariables Ancho(1) + Pos_X_Mid, PosLinea, "Cantidad: " & Format$(Contador, "00")
           PrinterTexto Ancho(3) + Pos_X_Mid, PosLinea, UCaseStrg("TOTAL DE " & Cuenta)
           PrinterVariables Ancho(6) + Pos_X_Mid, PosLinea, Total
           PosLinea = PosLinea + 0.4
           Encabezado_Diario_Caja_CxC Pos_X_Mid, Porte_Letra
           Printer.FontBold = False
           NombreBanco = .fields("Banco")
           Cta = .fields("Cta")
           Mifecha = .fields("Fecha")
           PrinterTexto Ancho(0) + Pos_X_Mid, PosLinea, Mifecha
           PosLinea = PosLinea + 0.35
           TotalIngreso = TotalIngreso + Total
           Total = 0
           Contador = 0
        End If
        Printer.FontBold = False
        Printer.FontSize = Porte_Letra
        If Mifecha <> .fields("Fecha") Then
           Mifecha = .fields("Fecha")
           PrinterTexto Ancho(0) + Pos_X_Mid, PosLinea, Mifecha
           PosLinea = PosLinea + 0.35
        End If
        Contador = Contador + 1
        PrinterTexto Ancho(0) + Pos_X_Mid, PosLinea, Contador & ".-"
        PrinterTexto Ancho(1) + Pos_X_Mid, PosLinea, .fields("Cliente")
        PrinterTexto Ancho(2) + Pos_X_Mid - 0.1, PosLinea, Space(120)
        PrinterTexto Ancho(2) + Pos_X_Mid, PosLinea, Format$(.fields("Factura"), "0000000")
        PrinterFields Ancho(3) + Pos_X_Mid, PosLinea, .fields("Banco")
        PrinterFields Ancho(5) + Pos_X_Mid, PosLinea, .fields("Cheque")
        PrinterFields Ancho(6) + Pos_X_Mid, PosLinea, .fields("Abono")
        Total = Total + .fields("Abono")
        PosLinea = PosLinea + 0.35
       'Si es Izquierdo
        If (PosLinea >= LimiteAlto) And (Derecho = False) Then
           PosLinea = Pos_Y_Mid
           Pos_X_Mid = (SetPapelAncho / 2)
           Derecho = True
           Encabezado_Diario_Caja_CxC Pos_X_Mid, Porte_Letra
           Printer.FontBold = False
           Mifecha = .fields("Fecha")
           PrinterTexto Ancho(0) + Pos_X_Mid, PosLinea, Mifecha
           PosLinea = PosLinea + 0.35
           NombreBanco = .fields("Banco")
           Cta = .fields("Cta")
        End If
       'Si es Derecho
        If (PosLinea >= LimiteAlto) And Derecho Then
           Imprimir_Linea_H PosLinea, Ancho(0) + Pos_X_Mid, Ancho(7) + Pos_X_Mid
           Printer.NewPage
           Encabezado Ancho(0), LimiteAncho
           Pos_X_Mid = (SetPapelAncho / 2)
           Imprimir_Linea_V Pos_X_Mid, PosLinea, LimiteAlto
           Pos_X_Mid = 0
           Derecho = False
           Printer.FontName = Tipo_Letra
           Printer.FontSize = Porte_Letra
           PrinterTexto Ancho(0), PosLinea, SQLMsg1
           PosLinea = PosLinea + 0.4
           Encabezado_Diario_Caja_CxC Pos_X_Mid, Porte_Letra
           Printer.FontBold = False
           Mifecha = .fields("Fecha")
           PrinterTexto Ancho(0) + Pos_X_Mid, PosLinea, Mifecha
           PosLinea = PosLinea + 0.35
        End If
       .MoveNext
     Loop
     Imprimir_Linea_H PosLinea, Ancho(0) + Pos_X_Mid, Ancho(7) + Pos_X_Mid
     PosLinea = PosLinea + 0.05
     TotalIngreso = TotalIngreso + Total
     Codigo = Leer_Cta_Catalogo(Cta)
     Printer.FontBold = True
     PrinterTexto Ancho(3) + Pos_X_Mid, PosLinea, UCaseStrg("TOTAL DE " & Cuenta)
     PrinterVariables Ancho(6) + Pos_X_Mid, PosLinea, Total
     PosLinea = PosLinea + 0.4
     Imprimir_Linea_H PosLinea, Ancho(3) + Pos_X_Mid, Ancho(7) + Pos_X_Mid
     PosLinea = PosLinea + 0.05
     PrinterTexto Ancho(3) + Pos_X_Mid, PosLinea, "T O T A L   R E C A U D A D O"
     PrinterVariables Ancho(6) + Pos_X_Mid, PosLinea, TotalIngreso
     PosLinea = PosLinea + 0.4
     Printer.FontBold = False
 End If
End With

'EMPIEZA LA IMPRESION DE LAS VENTAS DEL DIA
Contador = 0
Total = 0
Total_IVA = 0
Total_Desc = 0
Total_Servicio = 0
If FechaI = FechaF Then
   SQLMsg1 = "VENTAS DEL DIA " & FechaStrgCorta(MFechaI)
Else
   SQLMsg1 = "VENTAS DEL " & FechaStrgCorta(MFechaI) & " AL " & FechaStrgCorta(MFechaF)
End If
With AdoVentas.Recordset
 If .RecordCount > 0 Then
     Progreso_Barra.Valor_Maximo = Progreso_Barra.Valor_Maximo + .RecordCount
     Progreso_Esperar
    .MoveFirst
    'Es Izquierdo
    If (PosLinea >= LimiteAlto) And (Derecho = False) Then
       PosLinea = Pos_Y_Mid
       Pos_X_Mid = (SetPapelAncho / 2)
       Derecho = True
       Encabezado_Diario_Caja_Ventas Pos_X_Mid, Porte_Letra
       Printer.FontBold = False
       Mifecha = .fields("Fecha")
       PrinterTexto Ancho(0) + Pos_X_Mid, PosLinea, Mifecha
       PosLinea = PosLinea + 0.35
       'NombreBanco = .Fields("Banco")
       Cta = .fields("Cta_CxP")
    End If
    'Es Derecho
    If (PosLinea >= LimiteAlto) And Derecho Then
       Imprimir_Linea_H PosLinea, Ancho(0) + Pos_X_Mid, Ancho(7) + Pos_X_Mid
       Printer.NewPage
       Encabezado Ancho(0), LimiteAncho
       Pos_X_Mid = (SetPapelAncho / 2)
       Imprimir_Linea_V Pos_X_Mid, PosLinea, LimiteAlto
       Pos_X_Mid = 0
       Derecho = False
       Printer.FontName = Tipo_Letra
       Printer.FontSize = Porte_Letra
       PrinterTexto Ancho(0), PosLinea, SQLMsg1
       PosLinea = PosLinea + 0.4
       Encabezado_Diario_Caja_Ventas Pos_X_Mid, Porte_Letra
       Printer.FontBold = False
       Mifecha = .fields("Fecha")
       PrinterTexto Ancho(0) + Pos_X_Mid, PosLinea, Mifecha
       PosLinea = PosLinea + 0.35
    End If
     Printer.FontBold = False
     Mifecha = .fields("Fecha")
     PrinterTexto Ancho(0) + Pos_X_Mid, PosLinea, Mifecha
     PosLinea = PosLinea + 0.4
     Encabezado_Diario_Caja_Ventas Pos_X_Mid, Porte_Letra
     Printer.FontBold = False
     PosLinea = PosLinea + 0.35
     
     Do While Not .EOF
        'MsgBox "-->"
        Progreso_Esperar
        Printer.FontName = Tipo_Letra
        Printer.FontSize = Porte_Letra
        If Derecho Then
           Pos_X_Mid = (SetPapelAncho / 2)
        Else
           Pos_X_Mid = 0
        End If
        Printer.FontBold = False
        If Mifecha <> .fields("Fecha") Then
           Mifecha = .fields("Fecha")
           PrinterTexto Ancho(0) + Pos_X_Mid, PosLinea, Mifecha
           PosLinea = PosLinea + 0.35
        End If
        PrinterFields Ancho(1) + Pos_X_Mid - 0.3, PosLinea, .fields("Cliente")
        PrinterTexto Ancho(2) + Pos_X_Mid - 0.1, PosLinea, Space(120)
        If .fields("Ejecutivo") <> Ninguno Then
            PrinterTexto Ancho(2) + Pos_X_Mid - 0.8, PosLinea, "(" & Abreviatura_Texto(.fields("Ejecutivo")) & ")"
        End If
        PrinterTexto Ancho(2) + Pos_X_Mid, PosLinea, Format$(.fields("Factura"), "0000000")
        PrinterFields Ancho(3) + Pos_X_Mid, PosLinea, .fields("Total_IVA")
        PrinterVariables Ancho(4) + Pos_X_Mid, PosLinea, .fields("Descuento") + .fields("Descuento2")
        PrinterFields Ancho(5) + Pos_X_Mid, PosLinea, .fields("Servicio")
        PrinterFields Ancho(6) + Pos_X_Mid, PosLinea, .fields("Total_MN")
        Total_IVA = Total_IVA + .fields("Total_IVA")
        Total_Desc = Total_Desc + .fields("Descuento") + .fields("Descuento2")
        Total_Servicio = Total_Servicio + .fields("Servicio")
        Total = Total + .fields("Total_MN")
        PosLinea = PosLinea + 0.35
        
        'Es Izquierdo
        If (PosLinea >= LimiteAlto) And (Derecho = False) Then
           PosLinea = Pos_Y_Mid
           Pos_X_Mid = (SetPapelAncho / 2)
           Derecho = True
           Encabezado_Diario_Caja_Ventas Pos_X_Mid, Porte_Letra
           Printer.FontBold = False
           Mifecha = .fields("Fecha")
           PrinterTexto Ancho(0) + Pos_X_Mid, PosLinea, Mifecha
           PosLinea = PosLinea + 0.35
        End If
        'Es Derecho
        If (PosLinea >= LimiteAlto) And Derecho Then
           Imprimir_Linea_H PosLinea, Ancho(0) + Pos_X_Mid, Ancho(7) + Pos_X_Mid
           Printer.NewPage
           Encabezado Ancho(0), LimiteAncho
           Pos_X_Mid = (SetPapelAncho / 2)
           Imprimir_Linea_V Pos_X_Mid, PosLinea, LimiteAlto
           Pos_X_Mid = 0
           Derecho = False
           Printer.FontName = Tipo_Letra
           Printer.FontSize = Porte_Letra
           PrinterTexto Ancho(0), PosLinea, SQLMsg1
           PosLinea = PosLinea + 0.4
           Encabezado_Diario_Caja_Ventas Pos_X_Mid, Porte_Letra
           Printer.FontBold = False
           Mifecha = .fields("Fecha")
           PrinterTexto Ancho(0) + Pos_X_Mid, PosLinea, Mifecha
           PosLinea = PosLinea + 0.35
        End If
        Contador = Contador + 1
       .MoveNext
     Loop
     Imprimir_Linea_H PosLinea, Ancho(0) + Pos_X_Mid, Ancho(7) + Pos_X_Mid
     Printer.FontBold = True
     PosLinea = PosLinea + 0.05
     PrinterVariables Ancho(1) + Pos_X_Mid, PosLinea, "Cantidad: " & Format$(Contador, "00")
     PrinterTexto Ancho(2) + Pos_X_Mid, PosLinea, "T O T A L"
     PrinterVariables Ancho(3) + Pos_X_Mid, PosLinea, Total_IVA
     PrinterVariables Ancho(4) + Pos_X_Mid, PosLinea, Total_Desc
     PrinterVariables Ancho(5) + Pos_X_Mid, PosLinea, Total_Servicio
     PrinterVariables Ancho(6) + Pos_X_Mid, PosLinea, Total
     PosLinea = PosLinea + 0.4
     Printer.FontBold = False
'''     End If
 End If
End With

If FechaI = FechaF Then
   SQLMsg1 = "PRODUCTOS FACTURADOS DEL DIA " & FechaStrgCorta(MFechaI)
Else
   SQLMsg1 = "PRODUCTOS FACTURADOS DEL " & FechaStrgCorta(MFechaI) & " AL " & FechaStrgCorta(MFechaF)
End If
With AdoProd.Recordset
 If .RecordCount > 0 Then
    .MoveFirst
     Progreso_Esperar
     If (PosLinea >= LimiteAlto) And (Derecho = False) Then
        PosLinea = Pos_Y_Mid
        Pos_X_Mid = (SetPapelAncho / 2)
        Derecho = True
        Encabezado_Diario_Caja_Prod Pos_X_Mid, Porte_Letra
     End If
     If (PosLinea >= LimiteAlto) And Derecho Then
        Imprimir_Linea_H PosLinea, Ancho(0) + Pos_X_Mid, Ancho(7) + Pos_X_Mid
        Printer.NewPage
        Derecho = False
        Encabezado Ancho(0), LimiteAncho
        Pos_X_Mid = (SetPapelAncho / 2)
        Imprimir_Linea_V Pos_X_Mid, PosLinea, LimiteAlto
        Pos_X_Mid = 0
        Printer.FontName = Tipo_Letra
        Printer.FontSize = Porte_Letra
     End If
     PrinterTexto Ancho(0) + Pos_X_Mid, PosLinea, SQLMsg1
     PosLinea = PosLinea + 0.4
     Encabezado_Diario_Caja_Prod Pos_X_Mid, Porte_Letra
     
     SubTotal = 0: Total = 0
     Do While Not .EOF
        Progreso_Esperar
        
        If Derecho Then
           Pos_X_Mid = (SetPapelAncho / 2)
        Else
           Pos_X_Mid = 0
        End If
        
        Printer.FontBold = False
        Printer.FontSize = Porte_Letra
        If Not IsNull(.fields("SUBTOTALES")) Then
           SubTotal = .fields("SUBTOTALES") + .fields("SUBTOTAL_IVA")
           Total = Total + SubTotal
           Producto = .fields("Producto")
           If Len(Producto) > 45 Then Producto = MidStrg(Producto, 1, 45) & "..."
           PrinterTexto Ancho(1) + Pos_X_Mid, PosLinea, Producto
           PrinterFields Ancho(2) + Pos_X_Mid, PosLinea, .fields("Codigo")
           PrinterFields Ancho(3) + Pos_X_Mid, PosLinea, .fields("CANTIDADES")
           PrinterFields Ancho(4) + Pos_X_Mid, PosLinea, .fields("SUBTOTALES")
           PrinterFields Ancho(5) + Pos_X_Mid, PosLinea, .fields("SUBTOTAL_IVA")
           PrinterVariables Ancho(6) + Pos_X_Mid, PosLinea, SubTotal
        End If
        PosLinea = PosLinea + 0.35
        If (PosLinea >= LimiteAlto) And (Derecho = False) Then
           PosLinea = Pos_Y_Mid
           Pos_X_Mid = (SetPapelAncho / 2)
           Derecho = True
           Encabezado_Diario_Caja_Prod Pos_X_Mid, Porte_Letra
        End If
        If PosLinea >= LimiteAlto And Derecho Then
           Imprimir_Linea_H PosLinea, Ancho(0) + Pos_X_Mid, Ancho(7) + Pos_X_Mid
           Printer.NewPage
           Encabezado Ancho(0), LimiteAncho
           Pos_X_Mid = (SetPapelAncho / 2)
           Imprimir_Linea_V Pos_X_Mid, PosLinea, LimiteAlto
           Pos_X_Mid = 0
           Printer.FontName = Tipo_Letra
           Printer.FontSize = Porte_Letra
           PrinterTexto Ancho(0) + Pos_X_Mid, PosLinea, SQLMsg1
           PosLinea = PosLinea + 0.5
           Encabezado_Diario_Caja_Prod Pos_X_Mid, Porte_Letra
           Derecho = False
        End If
       .MoveNext
     Loop
     Imprimir_Linea_H PosLinea, Ancho(0) + Pos_X_Mid, Ancho(7) + Pos_X_Mid
     PosLinea = PosLinea + 0.05
 End If
End With

'Encabezado_Diario_Anticipos
If FechaI = FechaF Then
   SQLMsg1 = "ABONOS ANTICIPADOS DEL DIA " & FechaStrgCorta(MFechaI)
Else
   SQLMsg1 = "ABONOS ANTICIPADOS DEL " & FechaStrgCorta(MFechaI) & " AL " & FechaStrgCorta(MFechaF)
End If
With AdoAnt.Recordset
 If .RecordCount > 0 Then
    .MoveFirst
     Progreso_Esperar
     If (PosLinea >= LimiteAlto) And (Derecho = False) Then
        PosLinea = Pos_Y_Mid
        Pos_X_Mid = (SetPapelAncho / 2)
        Derecho = True
        Encabezado_Diario_Anticipos Pos_X_Mid, Porte_Letra
     End If
     If (PosLinea >= LimiteAlto) And Derecho Then
        Imprimir_Linea_H PosLinea, Ancho(0) + Pos_X_Mid, Ancho(7) + Pos_X_Mid
        Printer.NewPage
        Derecho = False
        Encabezado Ancho(0), LimiteAncho
        Pos_X_Mid = (SetPapelAncho / 2)
        Imprimir_Linea_V Pos_X_Mid, PosLinea, LimiteAlto
        Pos_X_Mid = 0
        Printer.FontName = Tipo_Letra
        Printer.FontSize = Porte_Letra
     End If
     PrinterTexto Ancho(0) + Pos_X_Mid, PosLinea, SQLMsg1
     PosLinea = PosLinea + 0.4
     Encabezado_Diario_Anticipos Pos_X_Mid, Porte_Letra
     
     Mifecha = .fields("Fecha")
     PrinterTexto Ancho(0) + Pos_X_Mid, PosLinea, Mifecha
     PosLinea = PosLinea + 0.35
     
     SubTotal = 0: Total = 0
     Do While Not .EOF
        Progreso_Esperar
        If Derecho Then
           Pos_X_Mid = (SetPapelAncho / 2)
        Else
           Pos_X_Mid = 0
        End If
        
        If Mifecha <> .fields("Fecha") Then
           Mifecha = .fields("Fecha")
           PrinterTexto Ancho(0) + Pos_X_Mid, PosLinea, Mifecha
           PosLinea = PosLinea + 0.35
        End If
        
        Printer.FontBold = False
        Printer.FontSize = Porte_Letra
        
        Cadena = .fields("TP") & " - " & Format(.fields("Numero"), "00000000")
        PrinterFields Ancho(1) + Pos_X_Mid, PosLinea, .fields("Cuenta")
        PrinterVariables Ancho(2) + Pos_X_Mid, PosLinea, Cadena
        PrinterFields Ancho(3) + 0.3 + Pos_X_Mid, PosLinea, .fields("Cliente")
        PrinterFields Ancho(6) + Pos_X_Mid, PosLinea, .fields("Creditos")
        PosLinea = PosLinea + 0.35
        If (PosLinea >= LimiteAlto) And (Derecho = False) Then
           PosLinea = Pos_Y_Mid
           Pos_X_Mid = (SetPapelAncho / 2)
           Derecho = True
           Encabezado_Diario_Anticipos Pos_X_Mid, Porte_Letra
        End If
        If PosLinea >= LimiteAlto And Derecho Then
           Imprimir_Linea_H PosLinea, Ancho(0) + Pos_X_Mid, Ancho(7) + Pos_X_Mid
           Printer.NewPage
           Encabezado Ancho(0), LimiteAncho
           Pos_X_Mid = (SetPapelAncho / 2)
           Imprimir_Linea_V Pos_X_Mid, PosLinea, LimiteAlto
           Pos_X_Mid = 0
           Printer.FontName = Tipo_Letra
           Printer.FontSize = Porte_Letra
           PrinterTexto Ancho(0) + Pos_X_Mid, PosLinea, SQLMsg1
           PosLinea = PosLinea + 0.5
           Encabezado_Diario_Anticipos Pos_X_Mid, Porte_Letra
           Derecho = False
        End If
       .MoveNext
     Loop
     Imprimir_Linea_H PosLinea, Ancho(0) + Pos_X_Mid, Ancho(7) + Pos_X_Mid
     PosLinea = PosLinea + 0.05
 End If
End With

PosLinea = PosLinea + 0.05
'Impresion de E/S de Inventario
If FechaI = FechaF Then
   SQLMsg1 = "ENTRADA/SALIDA DE INVENTARIO DEL DIA " & FechaStrgCorta(MFechaI)
Else
   SQLMsg1 = "ENTRADA/SALIDA DE INVENTARIO DEL " & FechaStrgCorta(MFechaI) & " AL " & FechaStrgCorta(MFechaF)
End If
With AdoInv.Recordset
 If .RecordCount > 0 Then
    .MoveFirst
     Progreso_Barra.Valor_Maximo = Progreso_Barra.Valor_Maximo + .RecordCount
     Progreso_Esperar
     
     If Derecho Then
        Pos_X_Mid = (SetPapelAncho / 2)
     Else
        Pos_X_Mid = 0
     End If
     If (PosLinea > LimiteAlto) And (Derecho = False) Then
        PosLinea = Pos_Y_Mid
        Pos_X_Mid = (SetPapelAncho / 2)
        Derecho = True
        Encabezado_Diario_Caja_Inv Pos_X_Mid, Porte_Letra
     End If
     If (PosLinea > LimiteAlto) And Derecho Then
        Imprimir_Linea_H PosLinea, Ancho(0) + Pos_X_Mid, Ancho(7) + Pos_X_Mid
        Printer.NewPage
        Derecho = False
        Encabezado Ancho(0), LimiteAncho
        Pos_X_Mid = (SetPapelAncho / 2)
        Imprimir_Linea_V Pos_X_Mid, PosLinea, LimiteAlto
        Pos_X_Mid = 0
        Printer.FontName = Tipo_Letra
        Printer.FontSize = Porte_Letra
     End If
     PrinterTexto Ancho(0) + Pos_X_Mid, PosLinea, SQLMsg1
     PosLinea = PosLinea + 0.4
     Encabezado_Diario_Caja_Inv Pos_X_Mid, Porte_Letra
     Do While Not .EOF
        Progreso_Esperar
        If Derecho Then
           Pos_X_Mid = (SetPapelAncho / 2)
        Else
           Pos_X_Mid = 0
        End If
    
        Printer.FontBold = False
        Printer.FontSize = Porte_Letra
        PrinterFields Ancho(0) + Pos_X_Mid, PosLinea, .fields("Codigo_Inv")
        PrinterFields Ancho(1) + Pos_X_Mid, PosLinea, .fields("Producto")
        PrinterFields Ancho(5) + Pos_X_Mid, PosLinea, .fields("Entradas")
        PrinterFields Ancho(6) + Pos_X_Mid, PosLinea, .fields("Salidas")
        PosLinea = PosLinea + 0.35
        If (PosLinea >= LimiteAlto) And (Derecho = False) Then
           PosLinea = Pos_Y_Mid
           Pos_X_Mid = (SetPapelAncho / 2)
           Derecho = True
           Encabezado_Diario_Caja_Inv Pos_X_Mid, Porte_Letra
        End If
        
        If (PosLinea >= LimiteAlto) And Derecho Then
           Imprimir_Linea_H PosLinea, Ancho(0) + Pos_X_Mid, Ancho(7) + Pos_X_Mid
           Printer.NewPage
           Derecho = False
           Encabezado Ancho(0), LimiteAncho
           Pos_X_Mid = (SetPapelAncho / 2)
           Imprimir_Linea_V Pos_X_Mid, PosLinea, LimiteAlto
           Pos_X_Mid = 0
           Printer.FontName = Tipo_Letra
           Printer.FontSize = Porte_Letra
           PrinterTexto Ancho(0) + Pos_X_Mid, PosLinea, SQLMsg1
           PosLinea = PosLinea + 0.4
           Encabezado_Diario_Caja_Inv Pos_X_Mid, Porte_Letra
        End If
       .MoveNext
     Loop
     Imprimir_Linea_H PosLinea, Ancho(0) + Pos_X_Mid, Ancho(7) + Pos_X_Mid
     PosLinea = PosLinea + 0.05
 End If
End With


PosLinea = PosLinea + 1
Printer.FontSize = Porte_Letra + 1
If Nombre_Cajero <> Ninguno Then
   PrinterTexto Ancho(0) + Pos_X_Mid, PosLinea, UCaseStrg(Nombre_Cajero)
Else
   PrinterTexto Ancho(0) + Pos_X_Mid, PosLinea, "RESPONSABLE"
End If
RatonNormal
MensajeEncabData = ""
Printer.EndDoc
End If
Progreso_Final
Exit Sub
Errorhandler:
    Progreso_Final
    RatonNormal
    ErrorDeImpresion
    Exit Sub
End Sub

Public Sub Imprimir_Abonos_De_Caja(AdoCxC As Adodc, _
                                   MFechaI As String, _
                                   MFechaF As String)
On Error GoTo Errorhandler
RatonReloj
Mensajes = "Seguro de Imprimir en:" & vbCrLf & Printer.DeviceName & "?"
Titulo = "IMPRESION"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
RatonReloj
Pagina = 1: InicioX = 0.5: InicioY = 0
Escala_Centimetro 1, TipoTimes, 10
CantCampos = 6
ReDim Ancho(CantCampos + 1) As Single
Ancho(0) = 1.3   'Fecha
Ancho(1) = 3     'Cliente
Ancho(2) = 9     'Factura
Ancho(3) = 10.5  'IVA
Ancho(4) = 13    'Descuento
Ancho(5) = 14.5  'Servicio
Ancho(6) = 16.5  'Total MN
Ancho(7) = 19    'Fin
'Iniciamos la impresion
Total = 0
Printer.FontSize = 10
Printer.FontName = TipoArialNarrow
With AdoCxC.Recordset
 If .RecordCount > 0 Then
    .MoveFirst
     Encabezado Ancho(0), Ancho(7)
     Printer.FontSize = 12
     If CFechaLong(MFechaI) = CFechaLong(MFechaF) Then
        PrinterTexto Ancho(0), PosLinea, "ABONOS DEL DIA " & FechaStrgCorta(MFechaI)
     Else
        PrinterTexto Ancho(0), PosLinea, "ABONO DEL " & FechaStrgCorta(MFechaI) & " AL " & FechaStrgCorta(MFechaF)
     End If
     PosLinea = PosLinea + 0.5
     Encabezado_Diario_Caja_CxC
     Printer.FontBold = False
     Printer.FontSize = 9
     Printer.FontName = TipoArialNarrow
     Mifecha = .fields("Fecha")
     PrinterTexto Ancho(0), PosLinea, Mifecha
     NombreBanco = .fields("Cliente")
     PrinterTexto Ancho(1), PosLinea, PrinterTextoMaximo(.fields("Cliente"), 5.5)
     Do While Not .EOF
        Printer.FontBold = False
        Printer.FontSize = 9
        If NombreBanco <> .fields("Cliente") Then
           PosLinea = PosLinea + 0.1
           Imprimir_Linea_H PosLinea, Ancho(0), Ancho(7)
           PosLinea = PosLinea + 0.1
           PrinterTexto Ancho(5), PosLinea, "T O T A L"
           PrinterVariables Ancho(6), PosLinea, Total
           PosLinea = PosLinea + 0.5
           Encabezado_Diario_Caja_CxC
           NombreBanco = .fields("Cliente")
           Mifecha = .fields("Fecha")
           Printer.FontBold = False
           Printer.FontSize = 9
           Printer.FontName = TipoArialNarrow
           PrinterTexto Ancho(0), PosLinea, Mifecha
           PrinterTexto Ancho(1), PosLinea, PrinterTextoMaximo(.fields("Cliente"), 5.5)
           Total = 0
        End If
        If Mifecha <> .fields("Fecha") Then
           Mifecha = .fields("Fecha")
           PrinterTexto Ancho(0), PosLinea, Mifecha
        End If
        PrinterTexto Ancho(2), PosLinea, Format$(.fields("Factura"), "0000000")
        PrinterFields Ancho(3), PosLinea, .fields("Banco")
        PrinterFields Ancho(5), PosLinea, .fields("Cheque")
        PrinterFields Ancho(6), PosLinea, .fields("Abono")
        Total = Total + .fields("Abono")
        PosLinea = PosLinea + 0.35
        If PosLinea >= LimiteAlto Then
           PosLinea = PosLinea + 0.1
           Imprimir_Linea_H PosLinea, Ancho(0), Ancho(7)
           Printer.NewPage
           Encabezado Ancho(0), Ancho(7)
           Printer.FontSize = 12
           If CFechaLong(MFechaI) = CFechaLong(MFechaF) Then
              PrinterTexto Ancho(0), PosLinea, "ABONOS DEL DIA " & FechaStrgCorta(MFechaI)
           Else
              PrinterTexto Ancho(0), PosLinea, "ABONO DEL " & FechaStrgCorta(MFechaI) & " AL " & FechaStrgCorta(MFechaF)
           End If
           PosLinea = PosLinea + 0.5
           Encabezado_Diario_Caja_CxC
           Printer.FontBold = False
           Printer.FontSize = 9
           Printer.FontName = TipoArialNarrow
           PrinterTexto Ancho(0), PosLinea, Mifecha
           PrinterTexto Ancho(1), PosLinea, PrinterTextoMaximo(.fields("Cliente"), 5.5)
           PrinterTexto Ancho(2), PosLinea, Format$(.fields("Factura"), "0000000")
           PrinterFields Ancho(3), PosLinea, .fields("Banco")
           PrinterFields Ancho(5), PosLinea, .fields("Cheque")
           PrinterFields Ancho(6), PosLinea, .fields("Abono")
           PosLinea = PosLinea + 0.35
        End If
       .MoveNext
     Loop
     PosLinea = PosLinea + 0.1
     Imprimir_Linea_H PosLinea, Ancho(0), Ancho(7)
     Printer.FontSize = 9
     PosLinea = PosLinea + 0.1
     PrinterTexto Ancho(5), PosLinea, "T O T A L"
     PrinterVariables Ancho(6), PosLinea, Total
     PosLinea = PosLinea + 0.5
 End If
End With
MensajeEncabData = ""
Printer.EndDoc
End If
RatonNormal
Exit Sub
Errorhandler:
    RatonNormal
    ErrorDeImpresion
    Exit Sub
End Sub

Public Sub ImprimirReciboCaja(DataDia As Adodc, FechaRecibo As String, Clientes As String)
Dim TipoProc As String
Dim SizeLetra As Integer
On Error GoTo Errorhandler
Mensajes = "Imprimir Recibo de Caja."
Titulo = "Formulario de Impresion."
If BoxMensaje = vbYes Then
RatonReloj
sSQL = "SELECT TA.*,C.Cliente,FA.Saldo_MN " _
     & "FROM Trans_Abonos As TA,Clientes AS C,Facturas As FA " _
     & "WHERE TA.Item = '" & NumEmpresa & "' " _
     & "AND TA.Periodo = '" & Periodo_Contable & "' " _
     & "AND C.Cliente = '" & Clientes & "' " _
     & "AND TA.Fecha = #" & BuscarFecha(FechaRecibo) & "# " _
     & "AND TA.CodigoC = C.Codigo " _
     & "AND FA.CodigoC = C.Codigo " _
     & "AND TA.Factura = FA.Factura " _
     & "AND TA.Periodo = FA.Periodo " _
     & "ORDER BY TA.Factura,TA.Comprobante,TA.Banco,TA.Cheque "
Select_Adodc DataDia, sSQL

If DataDia.Recordset.RecordCount > 0 Then
Pagina = 1: InicioX = 0.5: InicioY = 0
SizeLetra = 8
Escala_Centimetro 1, TipoTimes, SizeLetra
DataDia.Recordset.MoveFirst
CantCampos = DataDia.Recordset.fields.Count
ReDim Ancho(CantCampos + 1) As Single
ReDim SumaTotales(CantCampos + 1) As Variant
Ancho(0) = 0.5  ' Factura
Ancho(1) = 2    ' Banco
Ancho(2) = 9.5  ' Cheque
Ancho(3) = 12   ' Abono
Ancho(4) = 14.5 ' Saldo_MN
Ancho(5) = 17   ' Linea final
'Iniciamos la impresion
With DataDia.Recordset
    .MoveFirst
     Encabezado Ancho(0), Ancho(5)
     EncabReciboCaja 1, Clientes, DataDia
     Printer.FontBold = False
     For J = 0 To .fields.Count - 1
         SumaTotales(J) = 0
     Next J
     Saldo = .fields("Saldo_MN")
     Factura_No = .fields("Factura")
     PrinterFields Ancho(0), PosLinea, .fields("Factura"), True
     PrinterFields Ancho(4), PosLinea, .fields("Saldo_MN"), True
     Do While Not .EOF
        Printer.FontName = TipoTimes
        Printer.FontSize = SizeLetra
        If Factura_No <> .fields("Factura") Then
           PrinterFields Ancho(0), PosLinea, .fields("Factura"), True
           PrinterFields Ancho(4), PosLinea, .fields("Saldo_MN"), True
           Factura_No = .fields("Factura")
           Saldo = .fields("Saldo_MN")
        End If
        PrinterFields Ancho(1), PosLinea, .fields("Banco"), True
        PrinterFields Ancho(2), PosLinea, .fields("Cheque"), True
        PrinterFields Ancho(3), PosLinea, .fields("Abono"), True
        SumaTotales(3) = SumaTotales(3) + .fields("Abono")
        For I = 0 To 5
           Printer.Line (Ancho(I), PosLinea - 0.1)-(Ancho(I), PosLinea + 0.4), Negro
        Next I
        PosLinea = PosLinea + 0.4
        If PosLinea >= LimiteAlto Then
           Imprimir_Linea_H PosLinea, Ancho(0), Ancho(5)
           Printer.NewPage
           Encabezado Ancho(0), Ancho(5)
           EncabReciboCaja 2, "", DataDia
           Printer.FontSize = SizeLetra
        End If
       .MoveNext
     Loop
End With
Imprimir_Linea_H PosLinea, Ancho(0), Ancho(5)
PosLinea = PosLinea + 0.1
PrinterTexto Ancho(0), PosLinea, "T O T A L"
For I = 1 To 6
   PrinterVariables Ancho(I), PosLinea, SumaTotales(I)
Next I
PosLinea = PosLinea + 0.5
Imprimir_Linea_H PosLinea, Ancho(0), Ancho(5), Negro, True
PosLinea = PosLinea + 0.1
Printer.FontBold = True: Printer.FontSize = 14
PrinterTexto Ancho(0), PosLinea, "TOTAL ABONADO M/N: "
PrinterVariables Ancho(2) + 0.5, PosLinea, SumaTotales(3)
PosLinea = PosLinea + 0.6
Imprimir_Linea_H PosLinea, Ancho(0), Ancho(5), Negro, True
PosLinea = PosLinea + 1.5
Printer.FontBold = True: Printer.FontSize = 10
Cadena = "________________________"
PrinterTexto Ancho(1), PosLinea, Cadena
Cadena = "________________"
PrinterTexto Ancho(3), PosLinea, Cadena
PosLinea = PosLinea + 0.4
Cadena = "   RECIBI CONFORME"
PrinterTexto Ancho(1), PosLinea, Cadena
Cadena = "    C L I E N T E"
PrinterTexto Ancho(3), PosLinea, Cadena
Printer.FontBold = False
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

Public Sub CalculosTotalCaja(DataCajas As Data)
With DataCajas.Recordset
If .RecordCount > 0 Then
   .MoveFirst
    Efectivo = 0: Cheque = 0: Retencion = 0
    Do While Not .EOF
       If .fields("Posf") = False Then Cheque = Cheque + .fields("Cheque")
       Efectivo = Efectivo + .fields("Efectivo")
       Retencion = Retencion + .fields("Retencion")
      .MoveNext
    Loop
End If
End With
End Sub

Public Sub CalculosTotalDiarioCaja(DataIngCaja As Data)
Total_IngCaja = 0
With DataIngCaja.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      Do While Not .EOF
        Total_IngCaja = Total_IngCaja + .fields("Efectivo")
        Total_IngCaja = Total_IngCaja + .fields("Retencion")
        If .fields("Posf") = False Then
           Total_IngCaja = Total_IngCaja + .fields("Cheque")
        End If
        .MoveNext
      Loop
   End If
End With
End Sub

Public Function FacturarContratos() As String
Dim WrkJet As Workspace
Dim DataC As Database
Dim DataReg As Recordset
RatonReloj
Mifecha = BuscarFecha(FechaSistema)
Set WrkJet = CreateWorkspace("", "admin", "", dbUseJet)
Set DataC = WrkJet.OpenDatabase(RutaEmpresa & "\FACTURAS.MDB")
sSQL = "SELECT Contrato_No & ' => ' & Fecha & ', de: ' & Cliente As Contrato " _
     & "FROM Contratos,Clientes " _
     & "WHERE Contratos.Codigo_C = Clientes.Codigo " _
     & "AND Fecha <= #" & Mifecha & "# " _
     & "AND T = '" & Normal & "' " _
     & "ORDER BY Contrato_No,Fecha "
Set DataReg = DataC.OpenRecordset(sSQL, dbOpenDynaset, dbReadOnly)
Cadena = ""
If DataReg.RecordCount > 0 Then
   Cadena = "FACTURAR LOS SIGUIENTES CONTRATOS:" & Chr(13) & Chr(13)
   Do While Not DataReg.EOF
      Cadena = Cadena & Space(15) & "Contrato No. " & DataReg.fields("Contrato") & Chr(13)
      DataReg.MoveNext
   Loop
End If
DataC.Close
FacturarContratos = Cadena
RatonNormal
End Function

Public Function FacturarContratosVencidos() As String
Dim WrkJet As Workspace
Dim DataC As Database
Dim DataReg As Recordset
RatonReloj
Mifecha = BuscarFecha(FechaSistema)
Set WrkJet = CreateWorkspace("", "admin", "", dbUseJet)
Set DataC = WrkJet.OpenDatabase(RutaEmpresa & "\PRODUCC.MDB")
sSQL = "SELECT Contrato_No & ' => ' & Fecha & ', de: ' & Cliente As Contrato " _
     & "FROM Contratos_Meses,Clientes " _
     & "WHERE Contratos_Meses.Codigo_C = Clientes.Codigo " _
     & "AND Fecha <= #" & Mifecha & "# " _
     & "AND T = '" & Normal & "' " _
     & "ORDER BY Contrato_No,Fecha "
Set DataReg = DataC.OpenRecordset(sSQL, dbOpenDynaset, dbReadOnly)
Cadena = ""
If DataReg.RecordCount > 0 Then
   Cadena = "FACTURAR LOS SIGUIENTES CONTRATOS:" & Chr(13) & Chr(13)
   Do While Not DataReg.EOF
      Cadena = Cadena & Space(15) & "Contrato No. " & DataReg.fields("Contrato") & Chr(13)
      DataReg.MoveNext
   Loop
Else
   Cadena = "NO EXISTEN CONTRATOS VENCIDOS"
End If
DataC.Close
FacturarContratosVencidos = Cadena
RatonNormal
End Function

Public Sub ImprimirProd(Datas As Data, Datas1 As Data, FormaImp As Byte, SizeLetra As Integer)
Dim UnaSolaVez As Boolean
On Error GoTo Errorhandler
RatonReloj
UnaSolaVez = True
InicioX = 0.5: InicioY = 0
'Escala_Centimetro FormaImp, TipoCondensed, SizeLetra
DataAnchoCampos InicioX, Datas, SizeLetra, TipoCondensed, FormaImp
Pagina = 1
Produccion = 0: Sobrantes = 0: Faltantes = 0: Reposicion = 0
Rotos = 0: Saldos = 0: TotalVentas = 0: SaldoTotal = 0
'Iniciamos la impresion
EncabezadoData Datas
Printer.FontBold = False
With Datas.Recordset
    .MoveFirst
     Codigo = .fields(0)
     Printer.FontSize = SizeLetra
     Do While Not .EOF
        If Codigo <> .fields(0) Then
           Imprimir_Linea_H PosLinea, InicioX, Ancho(CantCampos)
           PosLinea = PosLinea + 0.1
           PrinterVariables Ancho(3), PosLinea, Produccion
           PrinterVariables Ancho(4), PosLinea, Sobrantes
           PrinterVariables Ancho(5), PosLinea, TotalVentas
           PrinterVariables Ancho(6), PosLinea, Rotos
           PrinterVariables Ancho(7), PosLinea, Faltantes
           PrinterVariables Ancho(8), PosLinea, SaldoTotal
           PrinterVariables Ancho(9), PosLinea, Reposicion
           Produccion = 0: Sobrantes = 0: Faltantes = 0
           Rotos = 0: Saldos = 0: TotalVentas = 0: Reposicion = 0
           PosLinea = PosLinea + 1: Printer.FontBold = True
           '===================================================
           PrinterAllFields CantCampos, PosLinea, Datas, True, True
           Printer.FontBold = False
           Codigo = .fields(0)
           PosLinea = PosLinea + 0.5
           UnaSolaVez = True
        End If
        SaldoTotal = .fields("Saldo_Actual")
        Produccion = Produccion + .fields("Produccion")
        Sobrantes = Sobrantes + .fields("Sobrantes")
        Faltantes = Faltantes + .fields("Faltantes")
        Rotos = Rotos + .fields("Rotos")
        Reposicion = Reposicion + .fields("Reposicion")
        TotalVentas = TotalVentas + .fields("Ventas_Dia")
        
        Printer.CurrentX = Ancho(0) - 0.1: Printer.CurrentY = PosLinea
        Printer.Print String$(15, " ")
        Printer.Line (Ancho(0), PosLinea - 0.1)-(Ancho(0), PosLinea + 0.4), Negro
        If UnaSolaVez Then
           PrinterFields Ancho(0), PosLinea, .fields(0), True
           Total = .fields(2)
           PrinterVariableTexto Ancho(0) + 1.5, PosLinea + 0.5, "Saldo Anterior:  ", Total
           UnaSolaVez = False
        End If
        PrinterAllFields CantCampos, PosLinea, Datas, True, False
        PosLinea = PosLinea + 0.4
        If PosLinea >= LimiteAlto Then
           Imprimir_Linea_H PosLinea, InicioX, Ancho(CantCampos)
           Printer.NewPage
           PosLinea = 0
           EncabezadoData Datas
           Printer.FontSize = SizeLetra
        End If
        .MoveNext
     Loop
    .MoveLast
End With
Imprimir_Linea_H PosLinea, InicioX, Ancho(CantCampos)
PosLinea = PosLinea + 0.1
PrinterVariables Ancho(3), PosLinea, Produccion
PrinterVariables Ancho(4), PosLinea, Sobrantes
PrinterVariables Ancho(5), PosLinea, TotalVentas
PrinterVariables Ancho(6), PosLinea, Rotos
PrinterVariables Ancho(7), PosLinea, Faltantes
PrinterVariables Ancho(8), PosLinea, SaldoTotal
PrinterVariables Ancho(9), PosLinea, Reposicion
Produccion = 0: Sobrantes = 0: Faltantes = 0
Rotos = 0: Saldos = 0: TotalVentas = 0: Reposicion = 0
PosLinea = PosLinea + 1: Printer.FontBold = True
Printer.Line (InicioX, PosLinea)-(Ancho(CantCampos), PosLinea), Negro
PosLinea = PosLinea + 0.5: Printer.FontSize = 14
PrinterTexto Ancho(0), PosLinea, "RESUMEN DE CUENTAS POR COBRAR Y VENTAS"
Printer.FontSize = SizeLetra
'DataAnchoCampos InicioX, Datas1, Decimales, SizeLetra
PosLinea = PosLinea + 0.5
If PosLinea >= LimiteAlto Then
   Imprimir_Linea_H PosLinea, InicioX, Ancho(CantCampos)
   Printer.NewPage
   PosLinea = 0
   EncabezadoData Datas1
   Printer.FontSize = SizeLetra
Else
   Printer.FontBold = True
   PrinterAllFields CantCampos, PosLinea, Datas1, True, True
   PosLinea = PosLinea + 0.5: Printer.FontBold = False
End If
Printer.FontSize = SizeLetra
With Datas1.Recordset
    .MoveFirst
     Do While Not .EOF
        PrinterAllFields CantCampos, PosLinea, Datas1, True, False
        PosLinea = PosLinea + 0.4
        If PosLinea >= LimiteAlto Then
           Imprimir_Linea_H PosLinea, InicioX, Ancho(CantCampos)
           Printer.NewPage
           PosLinea = 0
           EncabezadoData Datas1
           Printer.FontSize = SizeLetra
        End If
       .MoveNext
     Loop
    .MoveLast
End With
RatonNormal
MensajeEncabData = ""
Printer.EndDoc
Exit Sub
Errorhandler:
    RatonNormal
    ErrorDeImpresion
    Exit Sub
End Sub

Public Sub ImprimirVentas(Datas As Data, FinDoc As Boolean, FormaImp As Byte, SizeLetra As Integer)
Dim UnaSolaVez As Boolean
On Error GoTo Errorhandler
RatonReloj
UnaSolaVez = True
InicioX = 0.5: InicioY = 0
'Escala_Centimetro FormaImp, TipoCondensed, SizeLetra
DataAnchoCampos InicioX, Datas, SizeLetra, TipoCondensed, FormaImp
Pagina = 1
Reposicion = 0: TotalVentas = 0: SaldoTotal = 0
'Iniciamos la impresion
EncabezadoData Datas
Printer.FontBold = False
With Datas.Recordset
    .MoveFirst
     Codigo = .fields(0)
     Printer.FontSize = SizeLetra
     Do While Not .EOF
        If Codigo <> .fields(0) Then
           Imprimir_Linea_H PosLinea, InicioX, Ancho(CantCampos)
           PosLinea = PosLinea + 0.1
           PrinterVariables Ancho(3), PosLinea, TotalVentas
           PrinterVariables Ancho(4), PosLinea, Reposicion
           PrinterVariables Ancho(5), PosLinea, SaldoTotal
           TotalVentas = 0: Reposicion = 0: SaldoTotal = 0
           PosLinea = PosLinea + 1: Printer.FontBold = True
           '===================================================
           PrinterAllFields CantCampos, PosLinea, Datas, True, True
           Printer.FontBold = False
           Codigo = .fields(0)
           PosLinea = PosLinea + 0.5
           UnaSolaVez = True
        End If
        TotalVentas = TotalVentas + .fields(3)
        Reposicion = Reposicion + .fields(4)
        SaldoTotal = .fields(5)
        If UnaSolaVez Then
           PrinterFields Ancho(0), PosLinea, .fields(0), True
           UnaSolaVez = False
        End If
        For J = 1 To CantCampos - 1
           PrinterFields Ancho(J), PosLinea, .fields(J), True
        Next J
        PosLinea = PosLinea + 0.4
        If PosLinea > LimiteAlto Then
           Imprimir_Linea_H PosLinea, InicioX, Ancho(CantCampos)
           Printer.NewPage
           PosLinea = 0: UnaSolaVez = True
           EncabezadoData Datas
           Printer.FontSize = SizeLetra
        End If
       .MoveNext
     Loop
    .MoveLast
End With
Imprimir_Linea_H PosLinea, InicioX, Ancho(CantCampos)
PosLinea = PosLinea + 0.1
PrinterVariables Ancho(3), PosLinea, TotalVentas
PrinterVariables Ancho(4), PosLinea, Reposicion
PrinterVariables Ancho(5), PosLinea, SaldoTotal
PosLinea = PosLinea + 1
Imprimir_Linea_H PosLinea, InicioX, Ancho(CantCampos)
UltimaLinea = PosLinea + 0.5
RatonNormal
MensajeEncabData = ""
If FinDoc Then Printer.EndDoc
Exit Sub
Errorhandler:
    RatonNormal
    ErrorDeImpresion
    Exit Sub
End Sub

Public Sub ClienteNuevo(Datas As Adodc)
  If PosLinea + 2 > LimiteAlto Then
     Printer.NewPage
     PosLinea = 0
     Encabezado Ancho(0), Ancho(C)
  End If
  PorteLetra = Printer.FontSize
  LetraAnterior = Printer.FontName
  Printer.FontName = TipoTimes
  Printer.FontSize = 9
  Dibujo = RutaSistema & "\FORMATOS\CARTERA.GIF"
  PrinterPaint Dibujo, Ancho(0), PosLinea, 19, 2.5
  PosLinea = PosLinea + 0.05
  Printer.FontBold = True
  Printer.CurrentX = 3: Printer.CurrentY = PosLinea
  Printer.Print NombreCliente & "."
  Cadena = Datas.Recordset.fields("Telefono")
  Cadena = Cadena & " / " & Datas.Recordset.fields("Celular")
  'Cadena = Cadena & ". FAX: " & Datas.Recordset.Fields("FAX")
  PosLinea = PosLinea + 0.5
  Printer.CurrentX = 3.5: Printer.CurrentY = PosLinea
  Printer.Print Cadena
  Printer.CurrentX = 11.6: Printer.CurrentY = PosLinea
  Printer.Print Datas.Recordset.fields("Email")
  PosLinea = PosLinea + 0.5
  Printer.CurrentX = 3.5: Printer.CurrentY = PosLinea
  Cadena = Datas.Recordset.fields("Ciudad")
  Cadena = Cadena & ", " & Datas.Recordset.fields("Direccion")
  Printer.Print Cadena
  PosLinea = PosLinea + 1.5
  Printer.FontBold = False
  Printer.FontSize = PorteLetra
  Printer.FontName = LetraAnterior
End Sub

Public Sub EncabSuscrip(Datas As Data)
  With Datas.Recordset
  Printer.FontBold = True
  PosLinea = PosLinea + 0.05
  Cadena = UCaseStrg(.fields("Area") & " - " & .fields("Ciudad"))
  PrinterVariables 1.5, PosLinea, Cadena
  PosLinea = PosLinea + 0.4
  PrinterVariables 0.5, PosLinea, "Contrato"
  PrinterVariables 1.8, PosLinea, "S u s c r i p t o r"
  PrinterVariables 8.5, PosLinea, "D i r e c c i ó n"
  PrinterVariables 18, PosLinea, "Teléfono"
  PrinterVariables 20, PosLinea, "Periodo de Suscrip."
  PrinterVariables 23.5, PosLinea, "Estado"
  PrinterVariables 24.5, PosLinea, "OBS."
  PosLinea = PosLinea + 0.35
  Imprimir_Linea_H PosLinea, Ancho(0), Ancho(7), Negro, True
  PosLinea = PosLinea + 0.1
  Printer.FontBold = False
  End With
End Sub

'RUC_CI, TB, Razon_Social
Public Sub Imprimir_FAM(TFA As Tipo_Facturas, _
                        PosInic As Single, _
                        PosLinea1 As Single, _
                        DtaF As ADODB.Recordset, _
                        DtaD As ADODB.Recordset, _
                        Tipo_Pago As String, _
                        Optional ReImp As Boolean, _
                        Optional Solo_Copia As Boolean, _
                        Optional CheqSinCodigo As Boolean)
Dim PFil1 As Single
Dim PPFil1 As Single
Dim PAncho As Single
Dim LineasNo As Integer
Dim AltoLetras As Single
Dim SubTotal_Desc As Currency
Dim ValorUnit2 As Currency
Dim TamañoAnt As Integer
Dim YaEstaMes As Boolean
Dim MesFact As String
Dim MesFactV(12) As String
Dim ProductoAux As String
Dim IDMes As Byte
Dim CI_RUC_SRI As String
Dim RAZON_SOCIAL_SRI As String
Dim DIRECCION_SRI As String
Dim Imp_Mes As Boolean
 'MsgBox TFA.LogoFactura
  If TFA.LogoFactura = "MATRICIA" Then
     Imprimir_Formato_Propio "IF", PosInic, PosLinea1
  ElseIf TFA.LogoFactura <> "NINGUNO" And TFA.AnchoFactura > 0 And TFA.AltoFactura > 0 Then
     RutaOrigen = RutaSistema & "\FORMATOS\" & TFA.LogoFactura & ".gif"
    'MsgBox RutaOrigen
     PrinterPaint RutaOrigen, PosInic + SetD(38).PosX, PosLinea1 + SetD(38).PosY, TFA.AnchoFactura, TFA.AltoFactura
  End If
  Printer.Font = TipoArial
  Printer.FontBold = True
  With DtaF
     'MsgBox "Inicio " & PosInic & " x " & PosLinea1
     Codigo4 = Format$(.fields("Factura"), "000000000")
     If SetD(1).PosX > 0 And SetD(1).PosY > 0 Then
        TamañoAnt = Printer.FontSize
        Printer.FontSize = SetD(1).Porte
           'MsgBox PosInic + 0.01 & vbCrLf & PosLinea1 + 0.01 & vbCrLf & AnchoFactura & vbCrLf & AltoFactura
           PrinterPaint LogoTipo, PosInic + SetD(38).PosX + 0.05, PosLinea1 + SetD(38).PosY + 0.05, 3, 1.5
           If Empresa = NombreComercial Then
              Printer.FontSize = 11
              PrinterTexto PosInic + SetD(1).PosX, PosLinea1 + 0.4, Empresa
           Else
              Printer.FontSize = 10
              PrinterTexto PosInic + SetD(1).PosX, PosLinea1 + 0.1, Empresa
              Printer.FontSize = 9
              PrinterTexto PosInic + SetD(1).PosX, PosLinea1 + 0.55, NombreComercial
              End If
           Printer.FontSize = 8
           PrinterTexto PosInic + SetD(1).PosX, PosLinea1 + 0.95, "R.U.C. " & RUC
           PrinterTexto PosInic + SetD(1).PosX, PosLinea1 + 1.3, "Dirección: " & Direccion
           PrinterTexto PosInic + SetD(1).PosX, PosLinea1 + 1.65, "Telefono: " & Telefono1
           If TFA.LogoFactura = "RECIBOS" Then
              Codigo4 = "RECIBO DE PAGO No. " & Format$(.fields("Factura"), "000000000")
           Else
              Codigo4 = MidStrg(SerieFactura, 1, 3) & "-" & MidStrg(SerieFactura, 4, 3) & "-" & Format$(.fields("Factura"), "000000000")
           End If
           
           If SetD(28).PosX > 0 And SetD(28).PosY > 0 Then
              Printer.FontSize = SetD(28).Porte
              Printer.FontBold = False
              Cuenta = "Autorización otorgada por el S.R.I. para imprimir por medios Computarizados Facturas, Autorización No. " & Autorizacion
              PrinterTexto PosInic + SetD(28).PosX, PosLinea1 + SetD(28).PosY, Cuenta
              Cuenta = "Autorizado el: " & Fecha_Autorizacion & " y válido hasta el " & Fecha_Vence
              If .fields("T") = "A" Then
                  Cuenta = Cuenta & " - ANULADA"
              ElseIf ReImp Then
                  Cuenta = Cuenta & " - REIMPRESION"
              End If
              If Solo_Copia Then
                 Cuenta = Cuenta & " - COPIA EMISOR"
              Else
                 If (SetD(63).PosX + SetD(63).PosY) > 0 Then 'Copia y original en matricial
                    If ReImp Then
                       Cuenta = Cuenta & " - ORIGINAL ADQUIRIENTE"
                    Else
                       Cuenta = Cuenta & " - ORIGINAL ADQUIRIENTE/COPIA EMISOR"
                    End If
                 Else
                    If PosInic > 0.01 Then
                       Cuenta = Cuenta & " - COPIA EMISOR"
                    Else
                       Cuenta = Cuenta & " - ORIGINAL ADQUIRIENTE"
                    End If
                 End If
              End If
              PrinterTexto PosInic + SetD(28).PosX, PosLinea1 + SetD(28).PosY + 0.3, Cuenta
              Printer.FontBold = True
           End If
        Printer.FontSize = TamañoAnt
     End If
     Printer.Font = TipoArialNarrow
      'MsgBox Codigo4 & vbCrLf & PosInic & vbCrLf & PosLinea1
      'Pie de Factura
       Imp_Mes = .fields("Imp_Mes")
       If SetD(2).PosX > 0 Then
          If TFA.LogoFactura = "RECIBOS" Then
             Printer.FontSize = 14
             PrinterTexto PosInic + 2.5, PosLinea1 + 1.6, Codigo4
          Else
             Printer.FontSize = SetD(2).Porte
             PrinterTexto PosInic + SetD(2).PosX, PosLinea1 + SetD(2).PosY, Codigo4
          End If
       End If
       If SetD(3).PosX > 0 Then
          Printer.FontSize = SetD(3).Porte
          PrinterFields PosInic + SetD(3).PosX, PosLinea1 + SetD(3).PosY, .fields("Fecha")
       End If
       If SetD(4).PosX > 0 Then
          Printer.FontSize = SetD(4).Porte
          PrinterFields PosInic + SetD(4).PosX, PosLinea1 + SetD(4).PosY, .fields("Fecha_V")
       End If
       If SetD(10).PosX > 0 Then
          Printer.FontSize = SetD(10).Porte
          PrinterFields PosInic + SetD(10).PosX, PosLinea1 + SetD(10).PosY, .fields("Ciudad")
       End If
       If SetD(62).PosX > 0 Then
          Printer.FontSize = SetD(62).Porte
          PrinterFields PosInic + SetD(62).PosX, PosLinea1 + SetD(62).PosY, .fields("Nota")
       End If
       If SetD(65).PosX > 0 Then
          Printer.FontSize = SetD(65).Porte
          PrinterFields PosInic + SetD(65).PosX, PosLinea1 + SetD(65).PosY, .fields("Observacion")
       End If
       NivelNo = SinEspaciosIzq(.fields("Direccion"))
       If NivelNo = "" Then NivelNo = Ninguno
       If SetD(29).PosX > 0 Then
          Printer.FontSize = SetD(29).Porte
          PrinterTexto PosInic + SetD(29).PosX, PosLinea1 + SetD(29).PosY, NivelNo
       End If
       NivelNo = MidStrg(.fields("Direccion"), Len(NivelNo) + 1, Len(.fields("Direccion")) - Len(NivelNo) + 1)
       If NivelNo = "" Then NivelNo = Ninguno
       If SetD(30).PosX > 0 Then
          Printer.FontSize = SetD(30).Porte
          PrinterTexto PosInic + SetD(30).PosX, PosLinea1 + SetD(30).PosY, NivelNo
       End If
       
      'Datos del Cliente/Represetante
      'MsgBox .Fields("Razon_Social") & vbCrLf & .Fields("Cliente")
       If .fields("Razon_Social") = .fields("Cliente") Then
           Printer.FontSize = SetD(5).Porte
           If SetD(5).PosX > 0 Then PrinterTexto PosInic + SetD(5).PosX, PosLinea1 + SetD(5).PosY, .fields("Cliente")
           Printer.FontSize = SetD(8).Porte
           If SetD(8).PosX > 0 Then PrinterFields PosInic + SetD(8).PosX, PosLinea1 + SetD(8).PosY, .fields("Direccion")
           Printer.FontSize = SetD(11).Porte
           If SetD(11).PosX > 0 Then PrinterFields PosInic + SetD(11).PosX, PosLinea1 + SetD(11).PosY, .fields("CI_RUC")
       Else
           Select Case .fields("TB")
             Case "C", "P", "R"
                  CI_RUC_SRI = .fields("RUC_CI")  '.Fields("CI_RUC")
                  RAZON_SOCIAL_SRI = .fields("Razon_Social")
                  DIRECCION_SRI = .fields("DireccionT")
             Case Else
                  CI_RUC_SRI = "9999999999999"
                  RAZON_SOCIAL_SRI = "CONSUMIDOR FINAL"
                  DIRECCION_SRI = "SD"
           End Select
           Printer.FontSize = SetD(5).Porte
           If SetD(5).PosX > 0 Then PrinterTexto PosInic + SetD(5).PosX, PosLinea1 + SetD(5).PosY, RAZON_SOCIAL_SRI
           Printer.FontSize = SetD(41).Porte
           If SetD(41).PosX > 0 Then PrinterTexto PosInic + SetD(41).PosX, PosLinea1 + SetD(41).PosY, DIRECCION_SRI
           Printer.FontSize = SetD(36).Porte
           If SetD(36).PosX > 0 Then PrinterTexto PosInic + SetD(36).PosX, PosLinea1 + SetD(36).PosY, CI_RUC_SRI
            
           Printer.FontSize = SetD(64).Porte
           If SetD(64).PosX > 0 Then PrinterTexto PosInic + SetD(64).PosX, PosLinea1 + SetD(64).PosY, .fields("Cliente")
           Printer.FontSize = SetD(8).Porte
           If SetD(8).PosX > 0 Then PrinterFields PosInic + SetD(8).PosX, PosLinea1 + SetD(8).PosY, .fields("Direccion")
           Printer.FontSize = SetD(11).Porte
           If SetD(11).PosX > 0 Then PrinterFields PosInic + SetD(11).PosX, PosLinea1 + SetD(11).PosY, .fields("CI_RUC")
       End If
       
''       If SetD(32).PosX > 0 Then
''          Printer.FontSize = SetD(32).Porte
''          If .Fields("Representante") <> Ninguno Then
''              PrinterTexto PosInic + SetD(32).PosX, PosLinea1 + SetD(32).PosY, .Fields("Representante")
''          End If
''       End If
       If SetD(9).PosX > 0 Then
          Printer.FontSize = SetD(9).Porte
          PrinterFields PosInic + SetD(9).PosX, PosLinea1 + SetD(9).PosY, .fields("Telefono")
       End If
       If SetD(45).PosX > 0 Then
          Printer.FontSize = SetD(45).Porte
          PrinterFields PosInic + SetD(45).PosX, PosLinea1 + SetD(45).PosY, .fields("TelefonoT")
       End If
       If SetD(6).PosX > 0 Then
          Printer.FontSize = SetD(6).Porte
          PrinterFields PosInic + SetD(6).PosX, PosLinea1 + SetD(6).PosY, .fields("Codigo")
       End If
       If SetD(7).PosX > 0 Then
          Printer.FontSize = SetD(7).Porte
          PrinterFields PosInic + SetD(7).PosX, PosLinea1 + SetD(7).PosY, .fields("Grupo")
       End If
       If SetD(13).PosX > 0 Then
          Printer.FontSize = SetD(13).Porte
          PrinterFields PosInic + SetD(13).PosX, PosLinea1 + SetD(13).PosY, .fields("Email")
       End If
       'MsgBox "Pie"
      'Pie de Factura
        If SetD(22).PosX > 0 Then
           Printer.FontSize = SetD(22).Porte
           PrinterFields PosInic + SetD(22).PosX, PosLinea1 + SetD(22).PosY, .fields("SubTotal"), , True
        End If
        If SetD(23).PosX > 0 Then
           Printer.FontSize = SetD(23).Porte
           PrinterFields PosInic + SetD(23).PosX, PosLinea1 + SetD(23).PosY, .fields("Con_IVA"), , True
        End If
        If SetD(24).PosX > 0 Then
           Printer.FontSize = SetD(24).Porte
           PrinterFields PosInic + SetD(24).PosX, PosLinea1 + SetD(24).PosY, .fields("Sin_IVA"), , True
        End If
        If SetD(39).PosX > 0 Then
           Printer.FontSize = SetD(39).Porte
           PrinterVariables PosInic + SetD(39).PosX, PosLinea1 + SetD(39).PosY, .fields("Descuento") + .fields("Descuento2"), True
        End If
        If SetD(25).PosX > 0 Then
           Printer.FontSize = SetD(25).Porte
           PrinterFields PosInic + SetD(25).PosX, PosLinea1 + SetD(25).PosY, .fields("IVA"), , True
        End If
        If SetD(26).PosX > 0 Then
           Printer.FontSize = SetD(26).Porte
           If .fields("TC") = "NV" Then
              PrinterVariables PosInic + SetD(26).PosX, PosLinea1 + SetD(26).PosY, 0, True
           Else
              PrinterVariables PosInic + SetD(26).PosX, PosLinea1 + SetD(26).PosY, CInt(Porc_IVA * 100), True
           End If
        End If
        If SetD(27).PosX > 0 Then
           Printer.FontSize = SetD(27).Porte
           PrinterFields PosInic + SetD(27).PosX, PosLinea1 + SetD(27).PosY, .fields("Total_MN"), , True
        End If
        If SetD(33).PosX > 0 Then
           Printer.FontSize = SetD(33).Porte
           PrinterVariables PosInic + SetD(33).PosX, PosLinea1 + SetD(33).PosY, CCur(Diferencia), True
        End If
        If SetD(34).PosX > 0 Then
           Printer.FontSize = SetD(34).Porte
           PrinterVariables PosInic + SetD(34).PosX, PosLinea1 + SetD(34).PosY, CCur(SaldoPendiente), True
           If CheqSinCodigo Then PrinterTexto PosInic + 2, PosLinea1 + SetD(34).PosY, "P A G A R   E N   C O L E C T U R I A"
        End If
        If SetD(37).PosX > 0 Then
           If Not CheqSinCodigo Then
              Printer.FontSize = SetD(37).Porte
              PrinterTexto PosInic + SetD(37).PosX, PosLinea1 + SetD(37).PosY, CodigoDelBanco
           End If
        End If
        If SetD(40).PosX > 0 Then
           Printer.FontSize = SetD(40).Porte
           PrinterTexto PosInic + SetD(40).PosX, PosLinea1 + SetD(40).PosY, .fields("Hora")
        End If
        If SetD(42).PosX > 0 Then
           SubTotal_Desc = .fields("SubTotal") - .fields("Descuento")
           Printer.FontSize = SetD(42).Porte
           PrinterVariables PosInic + SetD(42).PosX, PosLinea1 + SetD(42).PosY, SubTotal_Desc, True
        End If
        If SetD(43).PosX > 0 Then
           Printer.FontSize = SetD(43).Porte
           PrinterTexto PosInic + SetD(43).PosX, PosLinea1 + SetD(43).PosY, "Descuento"
        End If
        If SetD(44).PosX > 0 Then
           Printer.FontSize = SetD(44).Porte
           PrinterTexto PosInic + SetD(44).PosX, PosLinea1 + SetD(44).PosY, NombreCiudad & ", " & .fields("Fecha")
        End If
        If SetD(46).PosX > 0 Then
           Printer.FontSize = SetD(46).Porte
           PrinterFields PosInic + SetD(46).PosX, PosLinea1 + SetD(46).PosY, .fields("DirNumero")
        End If
        If SetD(66).PosX > 0 And SetD(66).PosY > 0 Then
           Printer.FontSize = SetD(66).Porte
           If Len(TextoFormaPago) > 1 Then
               PrinterTexto PosInic + SetD(66).PosX, PosLinea1 + SetD(66).PosY, TextoFormaPago
           End If
        End If
       'Tipo_Pago
       'MsgBox PosInic & vbCrLf & PosLinea1
        If Len(Tipo_Pago) > 1 Then
           If SetD(79).PosX > 0 And SetD(79).PosY > 0 Then
              Printer.FontSize = SetD(79).Porte
              PrinterFields PosInic + SetD(79).PosX, PosLinea1 + SetD(79).PosY, .fields("Fecha_V")
           End If
           If SetD(78).PosX > 0 And SetD(78).PosY > 0 Then
              Printer.FontSize = SetD(78).Porte
              PrinterFields PosInic + SetD(78).PosX, PosLinea1 + SetD(78).PosY, .fields("Total_MN")
           End If
           If SetD(77).PosX > 0 And SetD(77).PosY > 0 Then
              RutaOrigen = RutaSistema & "\FORMATOS\Vistofp.jpg"
              PrinterPaint RutaOrigen, PosInic + SetD(77).PosX, PosLinea1 + SetD(77).PosY, SetD(77).Porte, SetD(77).Porte
           End If
           If SetD(76).PosX > 0 And SetD(76).PosY > 0 Then
              Printer.FontSize = SetD(76).Porte
              PrinterTexto PosInic + SetD(76).PosX, PosLinea1 + SetD(76).PosY, "TIPO PAGO:"
           End If
           If SetD(75).PosX > 0 And SetD(75).PosY > 0 Then
              Printer.FontSize = SetD(75).Porte
              PrinterTexto PosInic + SetD(75).PosX, PosLinea1 + SetD(75).PosY, Tipo_Pago
           End If
        End If
  End With
  Printer.Font = TipoArialNarrow
 'Detalle de la Factura
  For IDMes = 0 To 11
      MesFactV(IDMes) = ""
  Next IDMes
  AltoLetras = 0.4
  Printer.FontSize = SetD(17).Porte
  
  'MsgBox AltoLetras
  With DtaD
   If .RecordCount > 0 Then
      .MoveFirst
       Printer.FontSize = SetD(14).Porte
       AltoLetras = Redondear(Printer.TextHeight("H"), 2)
       PFil1 = PosLinea1 + SetD(14).PosY
       PAncho = SetD(18).PosX
       If .RecordCount > 0 Then
          .MoveFirst
           Producto = .fields("Producto") & " "
           ProductoAux = .fields("Producto")
           CodigoInv = .fields("Codigo")
           ValorUnit = .fields("Precio")
           ValorUnit2 = .fields("Precio2")
           CodigoP = ""
           Cantidad = 0
           SubTotal = 0
           SubTotal_IVA = 0
           Do While Not .EOF
              YaEstaMes = False
              MesFact = .fields("Mes")
              For IDMes = 0 To 11
                  If MesFactV(IDMes) = MesFact Then
                     YaEstaMes = True
                     IDMes = 11
                  End If
              Next IDMes
              If YaEstaMes = False Then
                 For IDMes = 0 To 11
                     If MesFactV(IDMes) = "" Then
                        MesFactV(IDMes) = MesFact
                        IDMes = 11
                     End If
                 Next IDMes
              End If
              If CodigoInv <> .fields("Codigo") Or ValorUnit <> .fields("Precio") Or ProductoAux <> .fields("Producto") Then
                 If Len(CodigoP) > 1 Then CodigoP = MidStrg(CodigoP, 1, Len(CodigoP) - 2)
                 Producto = Producto & CodigoP & " "
                 If SetD(16).PosX > 0 Then
                    Printer.FontSize = SetD(16).Porte
                    PrinterTexto PosInic + SetD(16).PosX, PFil1, CStr(Cantidad)
                 End If
                 If SetD(15).PosX > 0 Then
                    Printer.FontSize = SetD(15).Porte
                    PrinterTexto PosInic + SetD(15).PosX, PFil1, CodigoInv
                 End If
                 If SetD(17).PosX > 0 Then
                    Printer.FontSize = SetD(17).Porte
                    'MsgBox "<<<< " & PFil1
                    LineasNo = PrinterLineasMayor(PosInic + SetD(17).PosX, PFil1, Producto, PAncho)
                    'If LineasNo > 1 Then PFil1 = LineasNo * 0.35
                    PFil1 = PosLinea_Aux
                    PPFil1 = PFil1
                    'MsgBox PFil1
                 End If
                 If SetD(20).PosX > 0 Then
                    Printer.FontSize = SetD(20).Porte
                    PrinterVariables PosInic + SetD(20).PosX, PFil1, SubTotal, True
                 End If
                 If SetD(19).PosX > 0 Then
                    Printer.FontSize = SetD(19).Porte
                    PrinterVariables PosInic + SetD(19).PosX, PFil1, ValorUnit, True
                 End If
                 If SetD(47).PosX > 0 Then
                    Printer.FontSize = SetD(47).Porte
                    PrinterVariables PosInic + SetD(47).PosX, PFil1, ValorUnit2, True
                 End If
                 
                 Producto = .fields("Producto") & " "
                 ProductoAux = .fields("Producto")
                 CodigoInv = .fields("Codigo")
                 ValorUnit = .fields("Precio")
                 ValorUnit2 = .fields("Precio2")
                 'MsgBox ">>>> " & PFil1
                 'PFil1 = PFil1 + 0.4
                 PFil1 = PFil1 + AltoLetras
                 CodigoP = ""
                 Cantidad = 0
                 SubTotal = 0
                 SubTotal_IVA = 0
              End If
              SubTotal = SubTotal + .fields("Total")
              Cantidad = Cantidad + .fields("Cantidad")
              If Imp_Mes Then
                 If .fields("Mes") <> Ninguno Then CodigoP = CodigoP & MidStrg(.fields("Mes"), 1, 3)
                 If .fields("Ticket") <> Ninguno Then CodigoP = CodigoP & "-" & .fields("Ticket")
                 CodigoP = CodigoP & ", "
              End If
             .MoveNext
           Loop
           'PFil1 = PPFil1 + 0.35
           If Len(CodigoP) > 1 Then CodigoP = MidStrg(CodigoP, 1, Len(CodigoP) - 2)
           Producto = Producto & CodigoP & " "
           If SetD(16).PosX > 0 Then
              Printer.FontSize = SetD(16).Porte
              PrinterTexto PosInic + SetD(16).PosX, PFil1, CStr(Cantidad)
           End If
           If SetD(15).PosX > 0 Then
              Printer.FontSize = SetD(15).Porte
              PrinterTexto PosInic + SetD(15).PosX, PFil1, CodigoInv
           End If
           If SetD(17).PosX > 0 Then
              Printer.FontSize = SetD(17).Porte
              LineasNo = 0
              LineasNo = PrinterLineasMayor(PosInic + SetD(17).PosX, PFil1, Producto, PAncho)
              PFil1 = PosLinea_Aux
           End If
           If SetD(20).PosX > 0 Then
              Printer.FontSize = SetD(20).Porte
              PrinterVariables PosInic + SetD(20).PosX, PFil1, SubTotal, True
           End If
           If SetD(19).PosX > 0 Then
              Printer.FontSize = SetD(19).Porte
              PrinterVariables PosInic + SetD(19).PosX, PFil1, ValorUnit, True
           End If
           If SetD(47).PosX > 0 Then
              Printer.FontSize = SetD(47).Porte
              PrinterVariables PosInic + SetD(47).PosX, PFil1, ValorUnit2, True
           End If
           If SetD(31).PosX > 0 Then
              MesFact = ""
              For IDMes = 0 To 11
                  If MesFactV(IDMes) <> "" Then MesFact = MesFact & MesFactV(IDMes) & ", "
              Next IDMes
              MesFact = TrimStrg(MesFact)
              MesFact = MidStrg(MesFact, 1, Len(MesFact) - 1)
              Printer.FontSize = SetD(31).Porte
              PrinterTexto PosInic + SetD(31).PosX, PosLinea1 + SetD(31).PosY, MesFact
           End If
       End If
   End If
  End With
 'Pie de Factura
''  With DtaF
''  End With
End Sub

Public Sub Imprimir_FA_NV_Electronica(TFA As Tipo_Facturas)
Dim AdoDBFactura As ADODB.Recordset
Dim AdoDBDetalle As ADODB.Recordset
Dim CadenaMoneda As String
Dim Numero_Letras As String
Dim Tipo_Letras As String
Dim Cant_Ln As Single
Dim Una_Copia As Boolean
Dim PathCodigoBarra As String
Dim AnchoMaxDir As Single
Dim Establecimiento As Integer
Dim Imp_Mes As Boolean

If IsNumeric(TFA.Autorizacion) And Len(TFA.Autorizacion) >= 13 Then
   
   ContEspec = Leer_Campo_Empresa("Codigo_Contribuyente_Especial")
   Obligado_Conta = Leer_Campo_Empresa("Obligado_Conta")
   Ambiente = MidStrg(TFA.ClaveAcceso, 24, 1)
  'Generacion Codigo de Barras
   PathCodigoBarra = RutaSysBases & "\TEMP\" & TFA.ClaveAcceso & ".jpg"
   Tipo_Letras = TipoVerdana 'TipoArialNarrow
 
   PorteLetra = 7
   RatonReloj
   PosLinea = 0.01: InicioX = 0.01
   Printer.FontName = Tipo_Letras
   Printer.FontSize = PorteLetra
   SubTotal = 0
   Total = 0
   Total_IVA = 0
   Total_Desc = 0
   Cant_Ln = 0
  'Dibujo = RutaSistema & "\LOGOS\PRISMANE.JPG"
   Printer.FontName = Tipo_Letras
  'Iniciamos la consulta de impresion
   If TFA.TC = "PV" Then
      sSQL = "SELECT F.*,C.Cliente,C.CI_RUC,C.Telefono,C.Direccion,C.Ciudad,C.Grupo,C.Email " _
           & "FROM Trans_Ticket As F,Clientes As C " _
           & "WHERE F.Ticket = " & TFA.Factura & " " _
           & "AND F.TC = '" & TFA.TC & "' " _
           & "AND F.Periodo = '" & Periodo_Contable & "' " _
           & "AND F.Item = '" & NumEmpresa & "' " _
           & "AND C.Codigo = F.CodigoC "
   Else
      sSQL = "SELECT F.*,C.Cliente,C.CI_RUC,C.Telefono,C.Direccion,C.Ciudad,C.Grupo,C.Email " _
           & "FROM Facturas As F,Clientes As C " _
           & "WHERE F.Factura = " & TFA.Factura & " " _
           & "AND F.TC = '" & TFA.TC & "' " _
           & "AND F.Serie = '" & TFA.Serie & "' " _
           & "AND F.Periodo = '" & Periodo_Contable & "' " _
           & "AND F.Item = '" & NumEmpresa & "' " _
           & "AND C.Codigo = F.CodigoC "
   End If
  Select_AdoDB AdoDBFactura, sSQL
 'Datos Iniciales
  With AdoDBFactura
   If .RecordCount > 0 Then
      'Encabezado de la Factura
       AnchoMaxDir = 0
       Establecimiento = Val(MidStrg(.fields("Serie"), 1, 3))
       If Establecimiento > 1 Then DireccionEstab = TFA.DireccionEstab
       If Printer.TextWidth(Direccion) > AnchoMaxDir Then AnchoMaxDir = Printer.TextWidth(Direccion)
       If Printer.TextWidth(DireccionEstab) > AnchoMaxDir Then AnchoMaxDir = Printer.TextWidth(DireccionEstab)
       If Establecimiento > 1 Then
          PrinterPaint TFA.LogoTipoEstab, 0.01, PosLinea, 3, 1.5        'Ancho = 4.7
       Else
          PrinterPaint LogoTipo, 0.01, PosLinea, 3, 1.5      'Ancho = 4.7
       End If
       Printer.FontBold = True
       Printer.FontSize = PorteLetra + 2
       PrinterTexto 4.1, PosLinea, "R.U.C."
       PosLinea = PosLinea + 0.5
       PrinterTexto 3.6, PosLinea, RUC
       PosLinea = PosLinea + 0.5
       Printer.FontSize = PorteLetra
       PrinterTexto 3.6, PosLinea, "Teléfono: " & Telefono1
       Printer.FontSize = PorteLetra + 1
       PosLinea = PosLinea + 0.5
       If Encabezado_PV Then
          If TFA.TC <> "PV" Then
             If Printer.TextWidth(UCaseStrg(RazonSocial)) > 7 Then Printer.FontSize = PorteLetra - 1
             PrinterCentrarTexto 7, PosLinea, UCaseStrg(RazonSocial)
             PosLinea = PosLinea + 0.4
             If Establecimiento > 1 Then
                If UCaseStrg(RazonSocial) <> UCaseStrg(TFA.NombreEstab) Then
                   Printer.FontSize = PorteLetra + 2
                   If Printer.TextWidth(UCaseStrg(TFA.NombreEstab)) > 7 Then Printer.FontSize = PorteLetra - 1
                   PrinterCentrarTexto 7, PosLinea, UCaseStrg(TFA.NombreEstab)
                   PosLinea = PosLinea + 0.4
                End If
             Else
                If UCaseStrg(RazonSocial) <> UCaseStrg(NombreComercial) Then
                   Printer.FontSize = PorteLetra + 2
                   If Printer.TextWidth(UCaseStrg(RazonSocial)) > 7 Then Printer.FontSize = PorteLetra - 1
                   PrinterCentrarTexto 7, PosLinea, UCaseStrg(NombreComercial)
                   PosLinea = PosLinea + 0.4
                End If
             End If
             PosLinea = PosLinea + 0.1
             Printer.FontSize = PorteLetra
             If AnchoMaxDir > 5 Then Printer.FontSize = PorteLetra - 1
             Printer.FontBold = True
             PrinterTexto InicioX, PosLinea, "Dirección Mat.:"
             Printer.FontBold = False
             If AnchoMaxDir > 5 Then
                PrinterTexto InicioX + 2, PosLinea, Direccion
             Else
                PrinterTexto InicioX + 2.2, PosLinea, Direccion
             End If
             
             PosLinea = PosLinea + 0.35
             If Establecimiento > 1 Then
                Printer.FontBold = True
                PrinterTexto InicioX, PosLinea, "Dirección Suc.:"
                Printer.FontBold = False
                If AnchoMaxDir > 5 Then
                   PrinterTexto InicioX + 2, PosLinea, DireccionEstab
                Else
                   PrinterTexto InicioX + 2.2, PosLinea, DireccionEstab
                End If
                PosLinea = PosLinea + 0.35
                Printer.FontBold = True
                PrinterTexto InicioX, PosLinea, "Teléfono:"
                Printer.FontBold = False
                PrinterTexto 2, PosLinea, TFA.TelefonoEstab
                PosLinea = PosLinea + 0.4
             End If
             Printer.FontBold = True
             If Len(ContEspec) > 1 Then
                PrinterTexto InicioX, PosLinea, "Contribuyente Especial No. " & ContEspec
                PosLinea = PosLinea + 0.35
             End If
             PrinterTexto InicioX, PosLinea, "OBLIGADO A LLEVAR CONTABILIDAD: " & Obligado_Conta
             PosLinea = PosLinea + 0.35
             Imprimir_Linea_H PosLinea, InicioX, 7
             PosLinea = PosLinea + 0.05
             Printer.FontSize = PorteLetra + 1
             If TFA.TC = "NV" Then
                SerieFactura = MidStrg(.fields("Serie"), 1, 3) & "-" & MidStrg(.fields("Serie"), 4, 3)
                PrinterTexto InicioX, PosLinea, "NOTA DE VENTA No. " & SerieFactura & "-" & Format$(.fields("Factura"), "000000000")
             Else
                SerieFactura = MidStrg(.fields("Serie"), 1, 3) & "-" & MidStrg(.fields("Serie"), 4, 3)
                PrinterTexto InicioX, PosLinea, "FACTURA No. " & SerieFactura & "-" & Format$(.fields("Factura"), "000000000")
             End If
             PosLinea = PosLinea + 0.4
             PrinterTexto InicioX, PosLinea, "NUMERO DE AUTORIZACION: "
             PosLinea = PosLinea + 0.35
             Printer.FontBold = False
             PrinterTexto InicioX, PosLinea, TFA.Autorizacion
             Printer.FontSize = PorteLetra
             Printer.FontBold = True
             PosLinea = PosLinea + 0.4
             Printer.FontBold = True
             PrinterTexto InicioX, PosLinea, "FECHA: "
             Printer.FontBold = False
             PrinterTexto InicioX + 1.2, PosLinea, FechaStrgCorta(TFA.Fecha)
             Printer.FontBold = True
             PrinterTexto InicioX + 3.8, PosLinea, "HORA: "
             Printer.FontBold = False
             PrinterTexto InicioX + 4.9, PosLinea, TFA.Hora
             PosLinea = PosLinea + 0.35
             Printer.FontBold = True
             If Ambiente = "1" Then
                PrinterTexto InicioX, PosLinea, "AMBIENTE: PRUEBA"
             Else
                PrinterTexto InicioX, PosLinea, "AMBIENTE: PRODUCCION"
             End If
             PrinterTexto 3.8, PosLinea, "EMISIÓN: NORMAL"
             PosLinea = PosLinea + 0.35
             PrinterTexto InicioX, PosLinea, "CLAVE DE ACCESO"
             PosLinea = PosLinea + 0.4
             Printer.FontBold = False
            'Imprimimos el codigo de barra
             PrinterPaint PathCodigoBarra, 0.01, PosLinea, 7, 0.8
             PosLinea = PosLinea + 0.9
             Printer.FontSize = PorteLetra - 0.8
             PrinterTexto InicioX, PosLinea, TFA.ClaveAcceso
             PosLinea = PosLinea + 0.4
          Else
             SerieFactura = MidStrg(MesesLetras(Month(.fields("Fecha"))), 1, 3) & "-" & Year(.fields("Fecha"))
             PrinterTexto InicioX, PosLinea, "T I C K E T   No. " & SerieFactura & "-" & Format$(.fields("Factura"), "000000000")
             Total_Desc = .fields("Descuento") + .fields("Descuento2")
          End If
       Else
          If Printer.TextWidth(UCaseStrg(RazonSocial)) > 7 Then Printer.FontSize = PorteLetra - 1
          PrinterCentrarTexto 7, PosLinea, UCaseStrg(RazonSocial)
          PosLinea = PosLinea + 0.4
          If UCaseStrg(RazonSocial) <> UCaseStrg(NombreComercial) Then
             Printer.FontSize = PorteLetra + 2
             If Printer.TextWidth(UCaseStrg(RazonSocial)) > 7 Then Printer.FontSize = PorteLetra - 1
             PrinterCentrarTexto 7, PosLinea, UCaseStrg(NombreComercial)
             PosLinea = PosLinea + 0.4
          End If
          PosLinea = PosLinea + 0.1
          Printer.FontSize = PorteLetra
          If AnchoMaxDir > 5 Then Printer.FontSize = PorteLetra - 1
          Printer.FontBold = True
          PrinterTexto InicioX, PosLinea, "Dirección Matriz:"
          PosLinea = PosLinea + 0.4
          Printer.FontBold = False
          PrinterTexto InicioX, PosLinea, Direccion
          PosLinea = PosLinea + 0.35
           
          Printer.FontBold = True
          PrinterTexto InicioX, PosLinea, "CLAVE DE ACCESO No."
          PosLinea = PosLinea + 0.35
          Printer.FontBold = False
          PosLinea = PrinterLineasTexto(InicioX + 0.1, PosLinea, TFA.ClaveAcceso, 6.7)
'''          PrinterTexto InicioX, PosLinea, MidStrg(TFA.ClaveAcceso, 1, 40)
'''          PosLinea = PosLinea + 0.35
'''          PrinterTexto InicioX, PosLinea, MidStrg(TFA.ClaveAcceso, 41, 10)
          PosLinea = PosLinea + 0.35
          Printer.FontBold = True
          PrinterTexto InicioX, PosLinea, "AUTORIZACION No."
          PosLinea = PosLinea + 0.35
          Printer.FontBold = False
          PrinterTexto InicioX, PosLinea, TFA.Autorizacion
          PosLinea = PosLinea + 0.35
          Printer.FontBold = True
          PrinterTexto InicioX, PosLinea, "FECHA DE AUTORIZACION: " & TFA.Fecha_Aut
          PosLinea = PosLinea + 0.35
          Printer.FontBold = False
          Printer.FontBold = True
          PrinterTexto InicioX, PosLinea, "FECHA DE EMISION: " & TFA.Fecha
          PosLinea = PosLinea + 0.35
          PrinterTexto InicioX, PosLinea, "DOCUMENTO DE " & TFA.TC & " No. " & TFA.Serie & "-" & Format$(TFA.Factura, "0000000")
          PosLinea = PosLinea + 0.35
          Printer.FontBold = False
       End If
       Printer.FontSize = PorteLetra
       If Len(ReferenciaEmpresa) > 1 Then
          PrinterCentrarTexto 7, PosLinea, ULCase(ReferenciaEmpresa)
          PosLinea = PosLinea + 0.35
       End If
       Printer.FontSize = PorteLetra
       Imprimir_Linea_H PosLinea, InicioX, 7
       PosLinea = PosLinea + 0.05
      'Encabezado del Contribuyente
       Printer.FontBold = True
       PrinterTexto InicioX, PosLinea, "Razón Social/Nombres y Apellidos: "
       Printer.FontBold = False
       PosLinea = PosLinea + 0.35
       PrinterTexto InicioX, PosLinea, PrinterTextoMaximo(TFA.Razon_Social, 7)
       PosLinea = PosLinea + 0.35
       Printer.FontBold = True
       PrinterTexto InicioX, PosLinea, "Identificación: "
       Printer.FontBold = False
       PrinterTexto InicioX + 2.1, PosLinea, TFA.RUC_CI
       Printer.FontBold = True
       PrinterTexto InicioX + 4.5, PosLinea, "Teléf.: "
       Printer.FontBold = False
       PrinterTexto InicioX + 5.5, PosLinea, TFA.TelefonoC
       PosLinea = PosLinea + 0.35
       Printer.FontBold = True
       PrinterTexto InicioX, PosLinea, "Direccion: "
       Printer.FontBold = False
       PosLinea = PosLinea + 0.35
       PrinterTexto InicioX, PosLinea, TFA.DireccionC
       Printer.FontBold = True
       PosLinea = PosLinea + 0.35
       Printer.FontBold = True
       PrinterTexto InicioX, PosLinea, "Correo Electrónico: "
       Printer.FontBold = False
       PosLinea = PosLinea + 0.35
       PrinterTexto InicioX, PosLinea, TFA.EmailR
       PosLinea = PosLinea + 0.35
       If TFA.Cliente <> TFA.Razon_Social Then
           Printer.FontBold = True
           PrinterTexto InicioX, PosLinea, "Beneficiario: "
           Printer.FontBold = False
           PosLinea = PosLinea + 0.35
           PrinterTexto InicioX, PosLinea, PrinterTextoMaximo(TFA.Cliente, 7)
           PosLinea = PosLinea + 0.35
           Printer.FontBold = True
           PrinterTexto InicioX, PosLinea, "Ubicacion: "
           Printer.FontBold = False
           PosLinea = PosLinea + 0.35
           PrinterTexto InicioX, PosLinea, TFA.Curso
           Printer.FontBold = True
           PosLinea = PosLinea + 0.35
           PrinterTexto InicioX, PosLinea, "Codigo: "
           Printer.FontBold = False
           PrinterTexto InicioX + 2.1, PosLinea, TFA.CI_RUC
           PosLinea = PosLinea + 0.35
       End If
       PosLinea = PosLinea + 0.15
       Imprimir_Linea_H PosLinea, InicioX, 7
       PosLinea = PosLinea + 0.05
       Printer.FontBold = True
       Printer.FontSize = PorteLetra - 1
       PrinterTexto InicioX, PosLinea, "Cant."
       PrinterTexto InicioX + 0.8, PosLinea, "P R O D U C T O"
       PrinterTexto InicioX + 4.6, PosLinea, "P.V.P."
       PrinterTexto InicioX + 5.7, PosLinea, "T O T A L"
       PosLinea = PosLinea + 0.35
       Efectivo = .fields("Efectivo")
       Imprimir_Linea_H PosLinea, InicioX, 7, Negro
       PosLinea = PosLinea + 0.05
       Imp_Mes = .fields("Imp_Mes")
   End If
  End With
  Printer.FontBold = False
 'Comenzamos a recoger los detalles de la factura
  If TFA.TC = "PV" Then
     sSQL = "SELECT DF.*,CP.Detalle,CP.Codigo_Barra " _
          & "FROM Trans_Ticket As DF,Catalogo_Productos As CP " _
          & "WHERE DF.Ticket = " & TFA.Factura & " " _
          & "AND DF.TC = '" & TFA.TC & "' " _
          & "AND DF.Item = '" & NumEmpresa & "' " _
          & "AND DF.Periodo = '" & Periodo_Contable & "' " _
          & "AND DF.Item = CP.Item " _
          & "AND DF.Periodo = CP.Periodo " _
          & "AND DF.Codigo_Inv = CP.Codigo_Inv " _
          & "ORDER BY DF.ID "
  Else
     sSQL = "SELECT DF.*,CP.Detalle,CP.Codigo_Barra " _
          & "FROM Detalle_Factura As DF,Catalogo_Productos As CP " _
          & "WHERE DF.Factura = " & TFA.Factura & " " _
          & "AND DF.TC = '" & TFA.TC & "' " _
          & "AND DF.Serie = '" & TFA.Serie & "' " _
          & "AND DF.Item = '" & NumEmpresa & "' " _
          & "AND DF.Periodo = '" & Periodo_Contable & "' " _
          & "AND DF.Item = CP.Item " _
          & "AND DF.Periodo = CP.Periodo " _
          & "AND DF.Codigo = CP.Codigo_Inv " _
          & "ORDER BY DF.ID "
  End If
 'MsgBox "LA CONSULTA ES:" & vbCrLf & vbCrLf & sSQL
  Select_AdoDB AdoDBDetalle, sSQL
  With AdoDBDetalle
   If .RecordCount > 0 Then
       Do While Not .EOF
          PrinterTexto InicioX, PosLinea, CStr(.fields("Cantidad"))
          Producto = .fields("Producto") & " "
          If Imp_Mes Then
             If .fields("Mes") <> Ninguno Then Producto = Producto & MidStrg(.fields("Mes"), 1, 3)
             If .fields("Ticket") <> Ninguno Then Producto = Producto & "-" & .fields("Ticket")
             CodigoP = CodigoP & ", "
          End If
          PosLinea = PrinterLineasTexto(InicioX + 0.8, PosLinea, Producto, 3.5)
         'PrinterTexto InicioX + 0.8, PosLinea, PrinterTextoMaximo(.Fields("Producto"), 3.5)
          PrinterFields InicioX + 3.7, PosLinea, .fields("Precio")
          PrinterFields InicioX + 5, PosLinea, .fields("Total")
          SubTotal = SubTotal + .fields("Total")
          Total_Desc = Total_Desc + .fields("Total_Desc") + .fields("Total_Desc2")
          Total_IVA = Total_IVA + .fields("Total_IVA")
          PosLinea = PosLinea + 0.3
          Cant_Ln = Cant_Ln + 0.03
         .MoveNext
       Loop
   End If
  End With
 'Pie de factura
 '===========================================================
  'PosLinea = PosLinea + (2.3 - Cant_Ln)
  PosLinea = PosLinea + 0.2
  
  Imprimir_Linea_H PosLinea, InicioX, 7, Negro
  PosLinea = PosLinea + 0.05
  With AdoDBFactura
   If .RecordCount > 0 Then
      'SubTotal = Total
       Total = SubTotal + Total_IVA - Total_Desc
       Codigo1 = MidStrg(CodigoUsuario, 1, 4) & "X" & MidStrg(CodigoUsuario, Len(CodigoUsuario) - 1, 2)
       PrinterTexto InicioX, PosLinea, "Cajero: " & Codigo1
       Printer.FontBold = True
       PrinterTexto InicioX + 3.8, PosLinea, "SUBTOTAL"
       Printer.FontBold = False
       PrinterVariables InicioX + 5, PosLinea, SubTotal
       PosLinea = PosLinea + 0.3
       Printer.FontBold = True
       PrinterTexto InicioX + 3.8, PosLinea, "DESCUENTO"
       Printer.FontBold = False
       PrinterVariables InicioX + 5, PosLinea, Total_Desc, True
       PosLinea = PosLinea + 0.3
       If Encabezado_PV Then PrinterTexto InicioX, PosLinea, "Su  Factura  será  enviada  al"
       If TFA.TC <> "PV" Then
          Printer.FontBold = True
          PrinterTexto InicioX + 3.8, PosLinea, "I.V.A. " & Porc_IVA * 100 & "%"
          Printer.FontBold = False
          PrinterVariables InicioX + 5, PosLinea, Total_IVA
          PosLinea = PosLinea + 0.3
       End If
       If Encabezado_PV Then PrinterTexto InicioX, PosLinea, "correo electrónico registrado."
       Printer.FontBold = True
       PrinterTexto InicioX + 3.8, PosLinea, "T O T A L"
       Printer.FontBold = False
       PrinterVariables InicioX + 5, PosLinea, Total
       PosLinea = PosLinea + 0.3
       
       If TFA.TC <> "PV" Then
          If .fields("Cotizacion") > 0 Then PrinterTexto InicioX, PosLinea, "Cotizacion:"
       End If
       If Efectivo > 0 Then
          Printer.FontBold = True
          PrinterTexto InicioX + 3.8, PosLinea, "EFECTIVO"
          Printer.FontBold = False
          PrinterVariables InicioX + 5, PosLinea, Efectivo
          PosLinea = PosLinea + 0.3
          If TFA.TC <> "PV" Then
             If .fields("Cotizacion") > 0 Then PrinterTexto InicioX, PosLinea, Format$(.fields("Cotizacion"), "#,##0.00")
          End If
          Printer.FontBold = True
          PrinterTexto InicioX + 3.8, PosLinea, "CAMBIO"
          Printer.FontBold = False
          PrinterVariables InicioX + 5, PosLinea, Efectivo - Total, True
          PosLinea = PosLinea + 0.3
       End If
       Imprimir_Linea_H PosLinea, InicioX, 7, Negro
       PosLinea = PosLinea + 0.05
       If TFA.TC = "PV" Then
          PrinterCentrarTexto 4.7, PosLinea, "RECLAME SU FACTURA/NOTA DE VENTA"
          PosLinea = PosLinea + 0.25
          PrinterCentrarTexto 4.7, PosLinea, "EN CAJA"
''       Else
''          If Not Encabezado_PV Then PrinterCentrarTexto 7, PosLinea, "VERIFIQUE SUS DATOS"
       End If
       Printer.FontSize = PorteLetra + 1
       PosLinea = PrinterLineasTexto(InicioX + 0.1, PosLinea, Informativo_FAT, 6.7)
       'PrinterCentrarTexto 7, PosLinea, Informativo_FA
       PosLinea = PosLinea + 0.4
       Printer.FontSize = PorteLetra
       PrinterCentrarTexto 7, PosLinea, "www.diskcoversystem.com"
       PosLinea = PosLinea + 0.5
   End If
  End With
  Printer.NewPage
  AdoDBDetalle.Close
  AdoDBFactura.Close
End If
End Sub

Public Sub Imprimir_FA_NV_TJ(FTA As Tipo_Abono)
    Mensajes = "Imprimir Transaccion" & vbCrLf & FTA.Factura
    Titulo = "IMPRESION TARJETA DE CREDITO"
    Bandera = False
    SetPrinters.Show 1
    If PonImpresoraDefecto(SetNombrePRN) Then
       Escala_Centimetro Orientacion_Pagina, TipoArialNarrow, 10
       'For i = 0 To 1
           PosLinea = 1
           PrinterTexto 1.5, PosLinea, "COMPROBANTE DE TRANSACCION POR PAGO"
           PosLinea = PosLinea + 0.5
           PrinterTexto 1.5, PosLinea, "DE COMISION TARJETA DE CREDITO"
           PosLinea = PosLinea + 1
           PrinterTexto 1.5, PosLinea, "CLIENTE: " & FTA.Recibi_de
           PosLinea = PosLinea + 0.5
           PrinterTexto 1.5, PosLinea, "POR USD " & Format(FTA.Abono, "#,##0.00")
           PosLinea = PosLinea + 2
           PrinterTexto 1.5, PosLinea, "___________________"
           PosLinea = PosLinea + 0.4
           PrinterTexto 1.5, PosLinea, "ENTREGUE CONFORME"
           PosLinea = PosLinea + 0.5
           PrinterTexto 1.5, PosLinea, "."
           Printer.NewPage
           PosLinea = 1
           PrinterTexto 1.5, PosLinea, "COMPROBANTE DE TRANSACCION POR PAGO"
           PosLinea = PosLinea + 0.5
           PrinterTexto 1.5, PosLinea, "DE COMISION TARJETA DE CREDITO"
           PosLinea = PosLinea + 1
           PrinterTexto 1.5, PosLinea, "CLIENTE: " & FTA.Recibi_de
           PosLinea = PosLinea + 0.5
           PrinterTexto 1.5, PosLinea, "POR USD " & Format(FTA.Abono, "#,##0.00")
           PosLinea = PosLinea + 2
           PrinterTexto 1.5, PosLinea, "___________________"
           PosLinea = PosLinea + 0.4
           PrinterTexto 1.5, PosLinea, "RECIBI CONFORME"
           PosLinea = PosLinea + 0.5
           PrinterTexto 1.5, PosLinea, "."
           Printer.EndDoc
       'Next i
    End If
End Sub

Public Sub Imprimir_RCB(PosInic As Single, _
                        PosLinea1 As Single, _
                        DtaF As Adodc, _
                        DtaD As Adodc, _
                        Optional ReImp As Boolean, _
                        Optional Solo_Copia As Boolean)
Dim PFil1 As Single
Dim PPFil1 As Single
Dim PAncho As Single
Dim LineasNo As Integer
Dim AltoLetras As Single
Dim SubTotal_Desc As Currency
Dim TamañoAnt As Integer
Dim YaEstaMes As Boolean
Dim MesFact As String
Dim MesFactV(12) As String
Dim IDMes As Byte
  Printer.Font = TipoArialNarrow
  Printer.FontBold = True
  With DtaF.Recordset
     'MsgBox "Hoja"
     If .fields("Factura") > 0 Then
         Codigo4 = Format$(.fields("Factura"), "000000000")
     Else
         Codigo4 = .fields("CI_RUC")
     End If
     'MsgBox LogoFactura
     RutaOrigen = RutaSistema & "\FORMATOS\RECIBOS.GIF"
    'MsgBox LogoFactura & vbCrLf & AnchoFactura & vbCrLf & AltoFactura
     PrinterPaint RutaOrigen, PosInic + SetD(38).PosX, PosLinea1 + SetD(38).PosY, 10.5, 14
     If SetD(1).PosX > 0 And SetD(1).PosY > 0 Then
        TamañoAnt = Printer.FontSize
        Printer.FontSize = SetD(1).Porte
           'MsgBox PosInic + 0.01 & vbCrLf & PosLinea1 + 0.01 & vbCrLf & AnchoFactura & vbCrLf & AltoFactura
           PrinterPaint LogoTipo, PosInic + SetD(38).PosX + 1.05, PosLinea1 + SetD(38).PosY + 0.05, 3, 1.35
           If Empresa = NombreComercial Then
              Printer.FontSize = 11
              PrinterTexto PosInic + 3.5, PosLinea1 + 0.4, Empresa
           Else
              Printer.FontSize = 10
              PrinterTexto PosInic + 3.5, PosLinea1 + 0.1, Empresa
              Printer.FontSize = 9
              PrinterTexto PosInic + 3.5, PosLinea1 + 0.55, NombreComercial
           End If
           Printer.FontSize = 8
           PrinterTexto PosInic + 3.5, PosLinea1 + 0.95, "R.U.C. " & RUC
           PrinterTexto PosInic + 3.5, PosLinea1 + 1.3, "Telefono: " & Telefono1
           PrinterTexto PosInic + 1, PosLinea1 + 1.65, "Dirección: " & Direccion
           
           If .fields("Factura") > 0 Then
               Codigo4 = "RECIBO DE PAGO No. " & Format$(.fields("Factura"), "000000000")
           Else
               Codigo4 = "RECIBO DE PAGO No. " & .fields("CI_RUC")
           End If
           If SetD(28).PosX > 0 And SetD(28).PosY > 0 Then
              Printer.FontSize = SetD(28).Porte
              Printer.FontBold = False
              Cuenta = ""
              PrinterTexto PosInic + SetD(28).PosX, PosLinea1 + SetD(28).PosY, Cuenta
              PrinterTexto PosInic + SetD(28).PosX, PosLinea1 + SetD(28).PosY + 0.3, Cuenta
              Printer.FontBold = True
           End If
        Printer.FontSize = TamañoAnt
     End If
     Printer.Font = TipoArialNarrow
      'MsgBox Codigo4
      'Pie de Factura
       If SetD(2).PosX > 0 Then
          Printer.FontSize = SetD(2).Porte
          PrinterTexto PosInic + 1.05, PosLinea1 + 0.4 + SetD(2).PosY, Codigo4
       End If
       If SetD(3).PosX > 0 Then
          Printer.FontSize = SetD(3).Porte
          PrinterFields PosInic + SetD(3).PosX, PosLinea1 + SetD(3).PosY, .fields("Fecha")
       End If
       If SetD(4).PosX > 0 Then
          Printer.FontSize = SetD(4).Porte
          If CFechaLong(.fields("Fecha_V")) > CFechaLong(.fields("Fecha")) Then
             PrinterFields PosInic + SetD(4).PosX, PosLinea1 + SetD(4).PosY, .fields("Fecha_V")
          Else
             PrinterFields PosInic + SetD(4).PosX, PosLinea1 + SetD(4).PosY, .fields("Fecha")
          End If
       End If
       If SetD(11).PosX > 0 Then
          Printer.FontSize = SetD(11).Porte
          PrinterFields PosInic + SetD(11).PosX, PosLinea1 + SetD(11).PosY, .fields("CI_RUC")
       End If
       If SetD(10).PosX > 0 Then
          Printer.FontSize = SetD(10).Porte
          PrinterFields PosInic + SetD(10).PosX, PosLinea1 + SetD(10).PosY, .fields("Ciudad")
       End If
       If SetD(8).PosX > 0 Then
          Printer.FontSize = SetD(8).Porte
          PrinterFields PosInic + SetD(8).PosX, PosLinea1 + SetD(8).PosY, .fields("Direccion")
       End If
       If SetD(65).PosX > 0 Then
          Printer.FontSize = SetD(65).Porte
          PrinterFields PosInic + SetD(65).PosX, PosLinea1 + SetD(65).PosY, .fields("Observacion")
       End If
       NivelNo = SinEspaciosIzq(.fields("Direccion"))
       If NivelNo = "" Then NivelNo = Ninguno
       If SetD(29).PosX > 0 Then
          Printer.FontSize = SetD(29).Porte
          PrinterTexto PosInic + SetD(29).PosX, PosLinea1 + SetD(29).PosY, NivelNo
       End If
       NivelNo = MidStrg(.fields("Direccion"), Len(NivelNo) + 1, Len(.fields("Direccion")) - Len(NivelNo) + 1)
       If NivelNo = "" Then NivelNo = Ninguno
       If SetD(30).PosX > 0 Then
          Printer.FontSize = SetD(30).Porte
          PrinterTexto PosInic + SetD(30).PosX, PosLinea1 + SetD(30).PosY, NivelNo
       End If
       If SetD(5).PosX > 0 Then
          Printer.FontSize = SetD(5).Porte
          If FA_Educativo Then
             If .fields("Representante") = Ninguno Then
                 PrinterFields PosInic + SetD(5).PosX, PosLinea1 + SetD(5).PosY, .fields("Cliente")
             Else
                 PrinterFields PosInic + SetD(5).PosX, PosLinea1 + SetD(5).PosY, .fields("Representante")
             End If
          Else
              PrinterFields PosInic + SetD(5).PosX, PosLinea1 + SetD(5).PosY, .fields("Cliente")
          End If
       End If
       If SetD(32).PosX > 0 Then
          Printer.FontSize = SetD(32).Porte
          If .fields("Representante") = Ninguno Then
              PrinterTexto PosInic + SetD(32).PosX, PosLinea1 + SetD(32).PosY, "CONSUMIDOR FINAL"
          End If
       End If
       If SetD(36).PosX > 0 Then
          Printer.FontSize = SetD(36).Porte
          If .fields("CI_RUC_R") = Ninguno Then
              Select Case .fields("TD")
                Case "C", "P": ''nada'
                Case Else: PrinterTexto PosInic + SetD(36).PosX, PosLinea1 + SetD(36).PosY, "9999999999999"
              End Select
          Else
              PrinterFields PosInic + SetD(36).PosX, PosLinea1 + SetD(36).PosY, .fields("CI_RUC_R")
          End If
       End If
       If SetD(9).PosX > 0 Then
          Printer.FontSize = SetD(9).Porte
          PrinterFields PosInic + SetD(9).PosX, PosLinea1 + SetD(9).PosY, .fields("Telefono")
       End If
       If SetD(6).PosX > 0 Then
          Printer.FontSize = SetD(6).Porte
          PrinterFields PosInic + SetD(6).PosX, PosLinea1 + SetD(6).PosY, .fields("Codigo")
       End If
       If SetD(7).PosX > 0 Then
          Printer.FontSize = SetD(7).Porte
          PrinterFields PosInic + SetD(7).PosX, PosLinea1 + SetD(7).PosY, .fields("Grupo")
       End If
       If SetD(13).PosX > 0 Then
          Printer.FontSize = SetD(13).Porte
          PrinterFields PosInic + SetD(13).PosX, PosLinea1 + SetD(13).PosY, .fields("Email")
       End If
       'MsgBox "Pie"
      'Pie de Factura
        If SetD(22).PosX > 0 Then
           Printer.FontSize = SetD(22).Porte
           PrinterFields PosInic + SetD(22).PosX, PosLinea1 + SetD(22).PosY, .fields("SubTotal")
        End If
        If SetD(23).PosX > 0 Then
           Printer.FontSize = SetD(23).Porte
           PrinterFields PosInic + SetD(23).PosX, PosLinea1 + SetD(23).PosY, .fields("Con_IVA")
        End If
        If SetD(24).PosX > 0 Then
           Printer.FontSize = SetD(24).Porte
           PrinterFields PosInic + SetD(24).PosX, PosLinea1 + SetD(24).PosY, .fields("Sin_IVA")
        End If
        If SetD(39).PosX > 0 Then
           Printer.FontSize = SetD(39).Porte
           PrinterFields PosInic + SetD(39).PosX, PosLinea1 + SetD(39).PosY, .fields("Descuento")
        End If
        If SetD(25).PosX > 0 Then
           Printer.FontSize = SetD(25).Porte
           PrinterFields PosInic + SetD(25).PosX, PosLinea1 + SetD(25).PosY, .fields("IVA")
        End If
        If SetD(26).PosX > 0 Then
           Printer.FontSize = SetD(26).Porte
           If .fields("TC") = "NV" Then
              PrinterVariables PosInic + SetD(26).PosX, PosLinea1 + SetD(26).PosY, 0
           Else
              PrinterVariables PosInic + SetD(26).PosX, PosLinea1 + SetD(26).PosY, CInt(Porc_IVA * 100)
           End If
        End If
        If SetD(27).PosX > 0 Then
           Printer.FontSize = SetD(27).Porte
           PrinterFields PosInic + SetD(27).PosX, PosLinea1 + SetD(27).PosY, .fields("Total_MN")
        End If
        If SetD(34).PosX > 0 Then
           Printer.FontSize = SetD(34).Porte
           PrinterVariables PosInic + SetD(34).PosX, PosLinea1 + SetD(34).PosY, CCur(SaldoPendiente)
        End If
        If SetD(33).PosX > 0 Then
           Printer.FontSize = SetD(33).Porte
           PrinterVariables PosInic + SetD(33).PosX, PosLinea1 + SetD(33).PosY, CCur(Diferencia)
        End If
        If SetD(37).PosX > 0 Then
           Printer.FontSize = SetD(37).Porte
           PrinterTexto PosInic + SetD(37).PosX, PosLinea1 + SetD(37).PosY, CodigoDelBanco
        End If
        If SetD(40).PosX > 0 Then
           Printer.FontSize = SetD(40).Porte
           PrinterTexto PosInic + SetD(40).PosX, PosLinea1 + SetD(40).PosY, .fields("Hora")
        End If
        If SetD(41).PosX > 0 Then
           Printer.FontSize = SetD(41).Porte
           PrinterTexto PosInic + SetD(41).PosX, PosLinea1 + SetD(41).PosY, .fields("DireccionT")
        End If
        If SetD(42).PosX > 0 Then
           SubTotal_Desc = .fields("SubTotal") - .fields("Descuento")
           Printer.FontSize = SetD(42).Porte
           PrinterVariables PosInic + SetD(42).PosX, PosLinea1 + SetD(42).PosY, SubTotal_Desc
        End If
        If SetD(43).PosX > 0 Then
           Printer.FontSize = SetD(43).Porte
           PrinterTexto PosInic + SetD(43).PosX, PosLinea1 + SetD(43).PosY, "Descuento"
        End If
        If SetD(44).PosX > 0 Then
           Printer.FontSize = SetD(44).Porte
           PrinterTexto PosInic + SetD(44).PosX, PosLinea1 + SetD(44).PosY, NombreCiudad & ", " & .fields("Fecha")
        End If
        If SetD(64).PosX > 0 Then
           Printer.FontSize = SetD(64).Porte
           If .fields("Representante") <> Ninguno Then
               PrinterTexto PosInic + SetD(64).PosX, PosLinea1 + SetD(64).PosY, .fields("Cliente")
           End If
        End If
        If SetD(66).PosX > 0 And SetD(66).PosY > 0 Then
           Printer.FontSize = SetD(66).Porte
           If Len(TextoFormaPago) > 1 Then
               PrinterTexto PosInic + SetD(66).PosX, PosLinea1 + SetD(66).PosY, TextoFormaPago
           End If
        End If
        'MsgBox Printer.ScaleWidth & vbCrLf & Printer.ScaleHeight & vbCrLf _
        '& "FAM: " & vbCrLf _
        '& PosInic & vbCrLf & PosLinea1 & vbCrLf & Codigo4 & vbCrLf & Diferencia & vbCrLf & SaldoPendiente
  End With
  Printer.Font = TipoArialNarrow
 'Detalle de la Factura
  For IDMes = 0 To 11
      MesFactV(IDMes) = ""
  Next IDMes
  AltoLetras = 0.4
  Printer.FontSize = SetD(17).Porte
  AltoLetras = Redondear(Printer.TextHeight("H") - 0.05, 2)
  'MsgBox AltoLetras
  With DtaD.Recordset
     '.MoveFirst
       PFil1 = PosLinea1 + SetD(14).PosY
       PAncho = SetD(18).PosX
       If .RecordCount > 0 Then
          .MoveFirst
           Producto = .fields("Producto") & " "
           CodigoInv = .fields("Codigo")
           ValorUnit = .fields("Precio")
           CodigoP = ""
           Cantidad = 0
           SubTotal = 0
           SubTotal_IVA = 0
           Do While Not .EOF
              YaEstaMes = False
              MesFact = .fields("Mes")
              For IDMes = 0 To 11
                  If MesFactV(IDMes) = MesFact Then
                     YaEstaMes = True
                     IDMes = 11
                  End If
              Next IDMes
              If YaEstaMes = False Then
                 For IDMes = 0 To 11
                     If MesFactV(IDMes) = "" Then
                        MesFactV(IDMes) = MesFact
                        IDMes = 11
                     End If
                 Next IDMes
              End If
              If CodigoInv <> .fields("Codigo") Or ValorUnit <> .fields("Precio") Then
                 CodigoP = MidStrg(CodigoP, 1, Len(CodigoP) - 2)
                 Producto = Producto & " " & CodigoP
                 If SetD(16).PosX > 0 Then
                    Printer.FontSize = SetD(16).Porte
                    PrinterTexto PosInic + SetD(16).PosX, PFil1, CStr(Cantidad)
                 End If
                 If SetD(15).PosX > 0 Then
                    Printer.FontSize = SetD(15).Porte
                    PrinterTexto PosInic + SetD(15).PosX, PFil1, CodigoInv
                 End If
                 If SetD(17).PosX > 0 Then
                    Printer.FontSize = SetD(17).Porte
                    'MsgBox "<<<< " & PFil1
                    LineasNo = PrinterLineasMayor(PosInic + SetD(17).PosX, PFil1, Producto, PAncho)
                    'If LineasNo > 1 Then PFil1 = LineasNo * 0.35
                    PFil1 = PosLinea_Aux
                    PPFil1 = PFil1
                    'MsgBox PFil1
                 End If
                 If SetD(20).PosX > 0 Then
                    Printer.FontSize = SetD(20).Porte
                    PrinterVariables PosInic + SetD(20).PosX, PFil1, SubTotal
                 End If
                 If SetD(19).PosX > 0 Then
                    Printer.FontSize = SetD(19).Porte
                    PrinterVariables PosInic + SetD(19).PosX, PFil1, ValorUnit
                 End If
                 Producto = .fields("Producto")
                 CodigoInv = .fields("Codigo")
                 ValorUnit = .fields("Precio")
                 'MsgBox ">>>> " & PFil1
                 PFil1 = PFil1 + 0.35
                 'PFil1 = PFil1 + AltoLetras
                 CodigoP = ""
                 Cantidad = 0
                 SubTotal = 0
                 SubTotal_IVA = 0
              End If
              SubTotal = SubTotal + .fields("Total")
              Cantidad = Cantidad + .fields("Cantidad")
              If .fields("Mes") <> Ninguno Then CodigoP = CodigoP & .fields("Mes")
              If .fields("Ticket") <> Ninguno Then CodigoP = CodigoP & "-" & .fields("Ticket")
              CodigoP = CodigoP & ", "
             .MoveNext
           Loop
           'PFil1 = PPFil1 + 0.35
           CodigoP = MidStrg(CodigoP, 1, Len(CodigoP) - 2)
           Producto = Producto & CodigoP
           If SetD(16).PosX > 0 Then
              Printer.FontSize = SetD(16).Porte
              PrinterTexto PosInic + SetD(16).PosX, PFil1, CStr(Cantidad)
           End If
           If SetD(15).PosX > 0 Then
              Printer.FontSize = SetD(15).Porte
              PrinterTexto PosInic + SetD(15).PosX, PFil1, CodigoInv
           End If
           If SetD(17).PosX > 0 Then
              Printer.FontSize = SetD(17).Porte
              LineasNo = 0
              LineasNo = PrinterLineasMayor(PosInic + SetD(17).PosX, PFil1, Producto, PAncho)
              PFil1 = PosLinea_Aux
           End If
           If SetD(20).PosX > 0 Then
              Printer.FontSize = SetD(20).Porte
              PrinterVariables PosInic + SetD(20).PosX, PFil1, SubTotal
           End If
           If SetD(19).PosX > 0 Then
              Printer.FontSize = SetD(19).Porte
              PrinterVariables PosInic + SetD(19).PosX, PFil1, ValorUnit
           End If
           If SetD(31).PosX > 0 Then
              MesFact = ""
              For IDMes = 0 To 11
                  If MesFactV(IDMes) <> "" Then MesFact = MesFact & MesFactV(IDMes) & ", "
              Next IDMes
              MesFact = TrimStrg(MesFact)
              MesFact = MidStrg(MesFact, 1, Len(MesFact) - 1)
              Printer.FontSize = SetD(31).Porte
              PrinterTexto PosInic + SetD(31).PosX, PosLinea1 + SetD(31).PosY, MesFact
           End If
       End If
   End With
  'Pie de Factura
   End Sub

Public Sub Imprimir_REB(PosInic As Single, _
                        PosLinea1 As Single, _
                        DtaF As Adodc, _
                        DtaD As Adodc, _
                        FechaRep As String)
Dim PFil1 As Single
Dim PPFil1 As Single
Dim PAncho As Single
Dim LineasNo As Integer
Dim AltoLetras As Single
  Printer.Font = TipoArialNarrow
  Printer.FontBold = True
  With DtaF.Recordset
     'MsgBox PosInic & vbCrLf & PosLinea1
     Codigo4 = Format$(.fields("Factura"), "0000000")
     RutaOrigen = RutaSistema & "\FORMATOS\RECIBOFM.GIF"
     PrinterPaint RutaOrigen, PosInic + SetD(38).PosX, PosLinea1 + SetD(38).PosY, 10, 14
     'MsgBox SetD(1).PosX & vbCrLf & SetD(1).PosY
     If SetD(1).PosX > 0 And SetD(1).PosY > 0 Then
           'MsgBox PosInic & vbCrLf & PosLinea1 & vbCrLf & LogoTipo
           PrinterPaint LogoTipo, PosInic + 0.01, PosLinea1 + 0.01, 3, 1.5
           If Empresa = NombreComercial Then
              Printer.FontSize = 11
              PrinterTexto PosInic + 3, PosLinea1 + 0.4, Empresa
           Else
              Printer.FontSize = 10
              PrinterTexto PosInic + 3, PosLinea1 + 0.1, Empresa
              Printer.FontSize = 9
              PrinterTexto PosInic + 3, PosLinea1 + 0.55, NombreComercial
           End If
           Printer.FontSize = 8
           PrinterTexto PosInic + 3, PosLinea1 + 0.95, "R.U.C. " & RUC
           PrinterTexto PosInic + 3, PosLinea1 + 1.3, "Dirección: " & Direccion
           Codigo4 = "Referencia No. " & Format$(.fields("Factura"), "0000000")
           Printer.FontSize = 6
     End If
     Printer.Font = TipoArialNarrow
      'Encabezado de Factura
       If SetD(2).PosX > 0 Then
          Printer.FontSize = SetD(2).Porte
          PrinterTexto PosInic + SetD(2).PosX, PosLinea1 + SetD(2).PosY, Codigo4
       End If
       If SetD(3).PosX > 0 Then
          Printer.FontSize = SetD(3).Porte
          PrinterTexto PosInic + SetD(3).PosX, PosLinea1 + SetD(3).PosY, FechaRep
       End If
       If SetD(31).PosX > 0 Then
          Printer.FontSize = SetD(31).Porte
          PrinterTexto PosInic + SetD(31).PosX, PosLinea1 + SetD(31).PosY, MesesLetras(Month(FechaRep))
       End If
       If SetD(11).PosX > 0 Then
          Printer.FontSize = SetD(11).Porte
          PrinterFields PosInic + SetD(11).PosX, PosLinea1 + SetD(11).PosY, .fields("CI_RUC")
       End If
       If SetD(10).PosX > 0 Then
          Printer.FontSize = SetD(10).Porte
          PrinterFields PosInic + SetD(10).PosX, PosLinea1 + SetD(10).PosY, .fields("Ciudad")
       End If
       If SetD(8).PosX > 0 Then
          Printer.FontSize = SetD(8).Porte
          PrinterFields PosInic + SetD(8).PosX, PosLinea1 + SetD(8).PosY, .fields("Direccion")
       End If
       NivelNo = SinEspaciosIzq(.fields("Direccion"))
       If NivelNo = "" Then NivelNo = Ninguno
       If SetD(29).PosX > 0 Then
          Printer.FontSize = SetD(29).Porte
          PrinterTexto PosInic + SetD(29).PosX, PosLinea1 + SetD(29).PosY, NivelNo
       End If
       NivelNo = MidStrg(.fields("Direccion"), Len(NivelNo) + 1, Len(.fields("Direccion")) - Len(NivelNo) + 1)
       If NivelNo = "" Then NivelNo = Ninguno
       If SetD(30).PosX > 0 Then
          Printer.FontSize = SetD(30).Porte
          PrinterTexto PosInic + SetD(30).PosX, PosLinea1 + SetD(30).PosY, NivelNo
       End If
       If SetD(5).PosX > 0 Then
          Printer.FontSize = SetD(5).Porte
          If .fields("Representante") = Ninguno Then
              PrinterFields PosInic + SetD(5).PosX, PosLinea1 + SetD(5).PosY, .fields("Cliente")
          Else
              PrinterFields PosInic + SetD(5).PosX, PosLinea1 + SetD(5).PosY, .fields("Representante")
          End If
       End If
       If SetD(32).PosX > 0 Then
          Printer.FontSize = SetD(32).Porte
          If .fields("Representante") = Ninguno Then
              PrinterTexto PosInic + SetD(32).PosX, PosLinea1 + SetD(32).PosY, "CONSUMIDOR FINAL"
          End If
       End If
       If SetD(36).PosX > 0 Then
          Printer.FontSize = SetD(36).Porte
          If .fields("Cedula") = Ninguno Then
              PrinterTexto PosInic + SetD(36).PosX, PosLinea1 + SetD(36).PosY, "9999999999999"
          Else
              PrinterFields PosInic + SetD(36).PosX, PosLinea1 + SetD(36).PosY, .fields("Cedula")
          End If
       End If
       If SetD(9).PosX > 0 Then
          Printer.FontSize = SetD(9).Porte
          PrinterFields PosInic + SetD(9).PosX, PosLinea1 + SetD(9).PosY, .fields("Telefono")
       End If
       If SetD(6).PosX > 0 Then
          Printer.FontSize = SetD(6).Porte
          PrinterFields PosInic + SetD(6).PosX, PosLinea1 + SetD(6).PosY, .fields("Codigo")
       End If
       If SetD(7).PosX > 0 Then
          Printer.FontSize = SetD(7).Porte
          PrinterFields PosInic + SetD(7).PosX, PosLinea1 + SetD(7).PosY, .fields("Grupo")
       End If
       If SetD(13).PosX > 0 Then
          Printer.FontSize = SetD(13).Porte
          PrinterFields PosInic + SetD(13).PosX, PosLinea1 + SetD(13).PosY, .fields("Email")
       End If
        If SetD(37).PosX > 0 Then
           Printer.FontSize = SetD(37).Porte
           PrinterTexto PosInic + SetD(37).PosX, PosLinea1 + SetD(37).PosY, CodigoDelBanco
        End If
        If SetD(40).PosX > 0 Then
           Printer.FontSize = SetD(40).Porte
           PrinterTexto PosInic + SetD(40).PosX, PosLinea1 + SetD(40).PosY, .fields("Hora")
        End If
        If SetD(64).PosX > 0 Then
           Printer.FontSize = SetD(64).Porte
           If .fields("Representante") <> Ninguno Then
               PrinterTexto PosInic + SetD(64).PosX, PosLinea1 + SetD(64).PosY, .fields("Cliente")
           End If
        End If
  End With
  Printer.Font = TipoArialNarrow
 'Detalle de la Factura
  AltoLetras = 0.4
  Printer.FontSize = SetD(17).Porte
  AltoLetras = Redondear(Printer.TextHeight("H") - 0.05, 2)
  'MsgBox AltoLetras
  Total = 0
  SubTotal = 0
  With DtaD.Recordset
       PFil1 = SetD(14).PosY
       PAncho = SetD(18).PosX
       If .RecordCount > 0 Then
          .MoveFirst
           Producto = ""
           CodigoP = .fields("Periodo")
           Do While Not .EOF
              'MsgBox .Fields("Periodo") & "-" & .Fields("Mes")
              If Year(.fields("Fecha")) < Year(FechaRep) Then
                  If CodigoP <> .fields("Periodo") Then
                     If SetD(17).PosX > 0 Then
                        Printer.FontSize = SetD(17).Porte
                        PrinterTexto PosInic + SetD(17).PosX, PosLinea1 + PFil1, "Deuda Pendiente del " & CodigoP
                     End If
                     If SetD(20).PosX > 0 Then
                        Printer.FontSize = SetD(20).Porte
                        PrinterVariables PosInic + SetD(20).PosX, PosLinea1 + PFil1, SubTotal
                     End If
                     SubTotal = 0
                     CodigoP = .fields("Periodo")
                     PFil1 = PFil1 + Printer.TextHeight("H")
                  End If
                  SubTotal = SubTotal + .fields("Valor")
              Else
                 If SetD(17).PosX > 0 Then
                    Printer.FontSize = SetD(17).Porte
                    PrinterTexto PosInic + SetD(17).PosX, PosLinea1 + PFil1, .fields("Producto") & ": " & .fields("Mes")
                 End If
                 If SetD(20).PosX > 0 Then
                    Printer.FontSize = SetD(20).Porte
                    PrinterFields PosInic + SetD(20).PosX, PosLinea1 + PFil1, .fields("Total")
                 End If
                 PFil1 = PFil1 + Printer.TextHeight("H")
              End If
              Total = Total + .fields("Total")
             .MoveNext
           Loop
           If SubTotal > 0 Then
              If SetD(17).PosX > 0 Then
                 Printer.FontSize = SetD(17).Porte
                 PrinterTexto PosInic + SetD(17).PosX, PosLinea1 + PFil1, "Deuda Pendiente del " & CodigoP
              End If
              If SetD(20).PosX > 0 Then
                 Printer.FontSize = SetD(20).Porte
                 PrinterVariables PosInic + SetD(20).PosX, PosLinea1 + PFil1, SubTotal
              End If
              'MsgBox CodigoP & ": " & SubTotal
           End If
       End If
   End With
  'Pie de Factura
   With DtaF.Recordset
       'Pie de Factura
        If SetD(27).PosX > 0 Then
           Printer.FontSize = SetD(27).Porte
           PrinterVariables PosInic + SetD(27).PosX, PosLinea1 + SetD(27).PosY, Total
        End If
   End With
End Sub

Public Sub Generar_PDF_Donacion(TFA As Tipo_Facturas, _
                                VerDonacion As Boolean)
Dim AdoDBDet As ADODB.Recordset
Dim AdoDBAbo As ADODB.Recordset
Dim AdoDBProd As ADODB.Recordset

Dim tipoDeLetra As String
Dim Cod_Aux As String
Dim Cod_Bar As String
Dim Porc_Str As String
Dim DiasPago As String
Dim TipoAporte As String

Dim TDec_PVP As Byte

Dim PorteDeLetra As Integer

Dim IVA_Porc As Currency

Dim TempPosLinea As Single
Dim PosLineaTemp As Single
Dim TPosLinea As Single
Dim TPosLinea1 As Single
Dim PosLinf As Single
Dim TempPosLineaAbono As Single
    
    RatonReloj
    Leer_Datos_FA_NV TFA
    
    sSQL = "SELECT DF.*, CP.Reg_Sanitario, CP.Marca, CP.Desc_Item, CP.Codigo_Barra As Cod_Barras " _
         & "FROM Detalle_Factura As DF, Catalogo_Productos As CP " _
         & "WHERE DF.Item = '" & NumEmpresa & "' " _
         & "AND DF.Periodo = '" & Periodo_Contable & "' " _
         & "AND DF.TC = '" & TFA.TC & "' " _
         & "AND DF.Serie = '" & TFA.Serie & "' " _
         & "AND DF.Autorizacion = '" & TFA.Autorizacion & "' " _
         & "AND DF.Factura = " & TFA.Factura & " " _
         & "AND DF.Item = CP.Item " _
         & "AND DF.Periodo = CP.Periodo " _
         & "AND DF.Codigo = CP.Codigo_Inv " _
         & "ORDER BY DF.ID,DF.Codigo "
    Select_AdoDB AdoDBDet, sSQL
    
    TipoAporte = ""
    sSQL = "SELECT " & Full_Fields("Trans_Abonos") & " " _
         & "FROM Trans_Abonos " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND TP = '" & TFA.TC & "' " _
         & "AND Serie = '" & TFA.Serie & "' " _
         & "AND Autorizacion = '" & TFA.Autorizacion & "' " _
         & "AND Factura = " & TFA.Factura & " " _
         & "ORDER BY Fecha,ID "
    Select_AdoDB AdoDBAbo, sSQL
    With AdoDBAbo
     If .RecordCount > 0 Then
         Do While Not .EOF
            If .fields("Banco") = "EFECTIVO MN" Then
                TipoAporte = TipoAporte & "CONTADO/"
            Else
                TipoAporte = TipoAporte & .fields("Banco") & "-" & .fields("Cheque") & "/"
            End If
           .MoveNext
         Loop
     Else
         TipoAporte = "CREDITO"
     End If
    End With
    
'CONTADO/CREDITO/CHEQUE
   'Encabezado Detalle Factura
   'TipoArial / TipoVerdana / TipoHelvetica
    tipoDeLetra = TipoArial
    
   'Generacion Codigo de Barras
    If IsNumeric(TFA.Autorizacion_GR) Then Ambiente = MidStrg(TFA.ClaveAcceso_GR, 24, 1)
    
    TFA.PDF_ClaveAcceso = TFA.ClaveAcceso
    SerieFactura = Format$(TFA.Factura, "000000000")
    Autorizacion = TFA.Autorizacion
    MiHora = TFA.Fecha_Aut & " - " & TFA.Hora
    
    TFA.PDF_ClaveAcceso = TipoDoc & SerieFactura
    
   'Generamos el documento
   'tPrint.TipoImpresion = Es_Printer
   
    tPrint.TipoImpresion = Es_PDF

   'MsgBox TFA.PDF_ClaveAcceso
    tPrint.NombreArchivo = TFA.TC & " " & TFA.Serie & "-" & TFA.Factura
    tPrint.TituloArchivo = "Documento RIDE de Donacion"
    tPrint.TipoLetra = tipoDeLetra
    tPrint.OrientacionPagina = 1
    tPrint.PaginaA4 = True
    tPrint.EsCampoCorto = False
    tPrint.VerDocumento = VerDonacion
    
    Set cPrint = New cImpresion
    cPrint.iniciaImpresion
    cPrint.printCuadro 1, 1, 1.05, 1.05, Blanco, "B"
    cPrint.printImagen LogoTipo, 1.5, 1.1, 4.4, 2
    PosLinea = 1.1
    cPrint.tipoNegrilla = True
   '==================================================================================================0
    cPrint.printTexto 1.6, PosLinea, RazonSocial, 12, "C", 19
    If UCaseStrg(RazonSocial) <> UCaseStrg(NombreComercial) Then
       PosLinea = PosLinea + 0.5
       cPrint.printTexto 1.6, PosLinea, NombreComercial, 12, "C", 19
    End If
    PosLinea = PosLinea + 0.5
    TPosLinea = PosLinea
    cPrint.printTexto 1.6, PosLinea, "R.U.C. " & RUC, 12, "C", 19
    PosLinea = PosLinea + 0.5
    cPrint.printTexto 1.6, PosLinea, "DONACIÓN DE ALIMENTOS", 12, "C", 19
    PosLinea = PosLinea + 0.4
    cPrint.tipoNegrilla = False
    cPrint.letraTipo tipoDeLetra, 8
    cPrint.printTexto 1.6, PosLinea, Direccion, 8, "C", 19
    PosLinea = PosLinea + 0.4
    cPrint.printTexto 1.6, PosLinea, "Email: " & EmailEmpresa, 8, "C", 19
    PosLinea = PosLinea + 0.4
    cPrint.printTexto 1.6, PosLinea, "Teléfonos: " & Telefono1 & " / " & Telefono2, 8, "C", 19
    PosLinea = PosLinea + 0.4
    TPosLinea = PosLinea
   'Linea donde se empezara a imprimir el resto del documento
    PosLinea = 1.3
    cPrint.tipoNegrilla = True
    cPrint.letraTipo tipoDeLetra, 12, &HC0&
    cPrint.printTexto 17.7, PosLinea, "Nota de"
    PosLinea = PosLinea + 0.4
    cPrint.printTexto 17.3, PosLinea, "Donación No."
    cPrint.letraTipo tipoDeLetra, 14
    PosLinea = PosLinea + 0.5
    cPrint.printTexto 17.85, PosLinea, TFA.Serie
    PosLinea = PosLinea + 0.5
    cPrint.printTexto 17.4, PosLinea, SerieFactura
    cPrint.letraTipo tipoDeLetra, 10
    PosLinea = PosLinea + 0.5
    cPrint.printTexto 17.2, PosLinea, "FECHA: " & TFA.Fecha
    cPrint.tipoNegrilla = False
    
   'Cuadros Superiores
    cPrint.printCuadro 16.9, 1.1, 20, 3, Negro, "B"
    'cPrint.printCuadro 1.5, 2, 15, PosLinf, Blanco, "B", 1
   '==================================================================================================
   'Empezamos a escribir los datos del beneficiario/Cliente/Proveedor
    PosLinea = TPosLinea
    PosLinf = PosLinea
    PosLinea = PosLinea + 0.1
    cPrint.tipoNegrilla = True
    cPrint.letraTipo tipoDeLetra, 7
    cPrint.printTexto 1.6, PosLinea, "Razón Social/Nombres y Apellidos:"
    cPrint.printTexto 18.5, PosLinea, "Identificación: "
    PosLinea = PosLinea + 0.35
    cPrint.tipoNegrilla = False
    
    cPrint.printTexto 18.5, PosLinea, TFA.RUC_CI
    cPrint.printTexto 1.6, PosLinea, TFA.Razon_Social
    If cPrint.dNoLineas > 0 Then PosLinea = cPrint.dNoLineas
    PosLinea = PosLinea + 0.35
    cPrint.tipoNegrilla = True
    
    cPrint.printTexto 1.6, PosLinea, "Dirección:"
    cPrint.tipoNegrilla = False
    cPrint.printTexto 3, PosLinea, TFA.DireccionC
    
    DiasPago = CStr(CFechaLong(TFA.Fecha_V) - CFechaLong(TFA.Fecha))
     cPrint.printTexto 13, PosLinea, "Código: " & TFA.Grupo
     cPrint.printTexto 17.3, PosLinea, "Teléfono: " & TFA.TelefonoC
     PosLinea = PosLinea + 0.35
     cPrint.printTexto 1.6, PosLinea, "Aporte: " & TipoAporte
     cPrint.printTexto 17.3, PosLinea, "Número de gavetas: " & TFA.Gavetas
     PosLinea = PosLinea + 0.35
     cPrint.printTexto 1.6, PosLinea, "Atención: " & TFA.Contacto
     cPrint.printTexto 12.5, PosLinea, "Procesado Por: " & TFA.Digitador
     'MsgBox TFA.Digitador
    'Cuadro de Informacion Contribuyente
     cPrint.printCuadro 1.5, PosLinf, 20, PosLinea + 0.2, Negro, "B"
     PosLinea = PosLinea + 0.5
    'Fin de Impresion del Encabezado del Documento PDF
    '==================================================================================================
    TempPosLinea = PosLinea
   'Cuadro de encabezado del detalle
    cPrint.printCuadro 1.5, PosLinea, 20, PosLinea + 0.01, &HE2E2E2, "BF"
    cPrint.printCuadro 1.5, PosLinea, 20, PosLinea + 0.01, Negro, "B"
    PosLinea = PosLinea + 0.2
    cPrint.tipoNegrilla = True
    cPrint.letraTipo tipoDeLetra, 8
    cPrint.printTexto 1.6, PosLinea, "CÓDIGO", PorteDeLetra
    cPrint.printTexto 3.4, PosLinea, "GRUPO DE PRODUCTO", PorteDeLetra
    cPrint.printTexto 14.3, PosLinea, "DETALLE"
    cPrint.printTexto 18.3, PosLinea, "CANTIDAD (KG)", PorteDeLetra
    'cPrint.printLinea 1.1, PosLinea + 0.5, 20.3, PosLinea + 0.5, Negro
    PosLinea = PosLinea + 0.6 ' 10.4
    cPrint.colorDeLetra = Negro
    cPrint.tipoNegrilla = False
    cPrint.letraTipo tipoDeLetra, 8
   'Detalle de la Factura
    Cantidad = 0
    With AdoDBDet
     If .RecordCount > 0 Then
         Progreso_Barra.Valor_Maximo = Progreso_Barra.Valor_Maximo + .RecordCount
         Do While Not .EOF
            Progreso_Barra.Mensaje_Box = "Generando documento PDF"
            Progreso_Esperar
            cPrint.printTexto 1.6, PosLinea, .fields("Codigo")
            cPrint.printTexto 3.5, PosLinea, .fields("Producto")
            cPrint.printTexto 14.4, PosLinea, .fields("Tipo_Hab")
            cPrint.printFields 18, PosLinea, .fields("Cantidad")
            If PosLineaTemp > PosLinea Then PosLinea = PosLineaTemp
            PosLinea = PosLinea + 0.4
            Cantidad = Cantidad + .fields("Cantidad")
           'MsgBox "Linea.: " & PosLinea
            If PosLinea >= 26 Then
              'MsgBox "Nueva Pag.: " & PosLinea
              'Lineas al final del detalle de la factura
               PosLinea = PosLinea + 0.2
               cPrint.printCuadro 1.5, TempPosLinea, 10.3, PosLinea, Negro, "B"
               cPrint.printLinea 3.2, TempPosLinea, 3.2, PosLinea, Negro
               cPrint.printLinea 14.1, TempPosLinea, 14.1, PosLinea, Negro
               cPrint.printLinea 18, TempPosLinea, 18, PosLinea, Negro
               
               cPrint.printTexto 1.6, PosLinea, "C O N T I N U A   E N   L A   S I G U I E N T E   P A G I N A . . .", PorteDeLetra
               cPrint.paginaNueva
               PosLinea = 2
               TempPosLinea = PosLinea
               PosLinea = PosLinea - 0.25
               cPrint.tipoNegrilla = True
               cPrint.letraTipo tipoDeLetra, 8
               cPrint.printTexto 1.6, PosLinea, "CÓDIGO", PorteDeLetra
               cPrint.printTexto 3.4, PosLinea, "GRUPO DE PRODUCTO", PorteDeLetra
               cPrint.printTexto 14.3, PosLinea, "DETALLE"
               cPrint.printTexto 18.3, PosLinea, "CANTIDAD (KG)", PorteDeLetra
              
              'Detalle de la Factura
               cPrint.printCuadro 1.5, TempPosLinea - 0.2, 20, TempPosLinea + 0.5, Negro, "B"
               PosLinea = PosLinea + 0.45 ' 10.4
               cPrint.tipoNegrilla = False
               cPrint.letraTipo tipoDeLetra, 6
               cPrint.colorDeLetra = Negro
            End If
           .MoveNext
         Loop
     End If
    End With
    cPrint.printCuadro 1.5, PosLinea, 20, PosLinea + 0.01, &HE2E2E2, "BF"
    cPrint.printCuadro 1.5, PosLinea, 20, PosLinea + 0.01, Negro, "B"
    
   'Lineas al final del detalle de la factura
    cPrint.printCuadro 1.5, TempPosLinea, 20, PosLinea + 2.5, Negro, "B"
    
    cPrint.printLinea 3.3, TempPosLinea, 3.3, PosLinea, Negro
    cPrint.printLinea 14.2, TempPosLinea, 14.2, PosLinea, Negro
    cPrint.printLinea 18.2, TempPosLinea, 18.2, PosLinea + 0.6, Negro
    
    PosLinea = PosLinea + 0.2
    cPrint.printTexto 16.5, PosLinea, " T O T A L", PorteDeLetra
    cPrint.printVariable 18.3, PosLinea, CDbl(Cantidad), PorteDeLetra, , , 2
    PosLinea = PosLinea + 0.5
    cPrint.letraTipo tipoDeLetra, 7
    
    cPrint.printCuadro 1.5, PosLinea, 20, PosLinea + 0.01, &HE2E2E2, "BF"
    cPrint.printCuadro 1.5, PosLinea - 0.01, 20, PosLinea + 0.01, Negro, "B"
    cPrint.printLinea 10.4, PosLinea, 10.4, PosLinea + 2.4, Negro
    PosLinea = PosLinea + 0.15
    cPrint.printTexto 1.6, PosLinea, "Nota:"
    cPrint.printTexto 10.5, PosLinea, "Observacion:"
    If Len(TFA.Nota) > 1 Then cPrint.printTexto 2.25, PosLinea, TFA.Nota, PorteDeLetra
    If Len(TFA.Observacion) > 1 Then cPrint.printTexto 12, PosLinea, TFA.Observacion
    PosLinea = PosLinea + 0.4
    'cPrint.printCuadro 1.5, PosLinea, 20, PosLinea + 2, Negro, "B"
    
    PosLinea = PosLinea + 1.5
    cPrint.printTexto 5, PosLinea, "Entregado Por", PorteDeLetra
    cPrint.printTexto 14.5, PosLinea, "Recibi Conforme", PorteDeLetra
    PosLinea = PosLinea + 0.6
    cPrint.colorDeLetra = Negro
    cPrint.tipoNegrilla = True
    cPrint.letraTipo tipoDeLetra, 7
    cPrint.printTexto 1.5, PosLinea, "Nota Importante:", PorteDeLetra
    cPrint.colorDeLetra = Negro
    cPrint.tipoNegrilla = False
    cPrint.letraTipo tipoDeLetra, 7
    Cadena = "Los productos donados, han perdido valor comercial por diferentes motivos, pero mantienen un valor social. " _
           & "Estos productos han pasado por un proceso de clasificación y se encuentran en buen estado. Se recomienda su " _
           & "consumo INMEDIATO y se prohíbe su comercialización. EL BANCO DE ALIMENTOS QUITO no se responsabiliza por " _
           & "cualquier efecto negativo que causare el consumo de alimentos en un tiempo mayor al sugerido.  Con su firma " _
           & "el beneficiario acepta que ha sido informado sobre el estado de los productos, que los recibe con su consentimiento, " _
           & "que los usará para fines benéficos y bajo su completa responsabilidad."
    PosLinea = cPrint.printTextoMultiple(3.6, PosLinea, Cadena, 16.6)
   '==================================================================================================
    PosLinea = PosLinea + 1
    TempPosLinea = PosLinea
    PosLinea = PosLinea + 0.2
    cPrint.printImagen LogoTipo, 1.6, PosLinea, 4.4, 2
    cPrint.tipoNegrilla = True
    cPrint.printTexto 1.6, PosLinea, RazonSocial, 12, "C", 19
    If UCaseStrg(RazonSocial) <> UCaseStrg(NombreComercial) Then
       PosLinea = PosLinea + 0.5
       cPrint.printTexto 1.6, PosLinea, NombreComercial, 12, "C", 19
    End If
    PosLinea = PosLinea + 0.5
    TPosLinea = PosLinea
    TPosLinea1 = PosLinea
    cPrint.printTexto 1.6, PosLinea, "R.U.C. " & RUC, 12, "C", 19
    PosLinea = PosLinea + 0.5
    cPrint.printTexto 1.6, PosLinea, "APORTE SOLIDARIO", 12, "C", 19
    PosLinea = PosLinea + 0.4
    cPrint.tipoNegrilla = False
    cPrint.letraTipo tipoDeLetra, 8
    cPrint.printTexto 1.6, PosLinea, Direccion, 8, "C", 19
    PosLinea = PosLinea + 0.4
    cPrint.printTexto 1.6, PosLinea, "Email: " & EmailEmpresa, 8, "C", 19
    PosLinea = PosLinea + 0.4
    cPrint.printTexto 1.6, PosLinea, "Teléfonos: " & Telefono1 & " / " & Telefono2, 8, "C", 19
    PosLinea = PosLinea + 0.5
   'Linea donde se empezara a imprimir el resto del documento
    TPosLinea1 = PosLinea
    PosLinea = TempPosLinea + 0.4 'TPosLinea
    cPrint.tipoNegrilla = True
    cPrint.letraTipo tipoDeLetra, 12, &HC0&
    cPrint.printTexto 17.7, PosLinea, "Nota de"
    PosLinea = PosLinea + 0.4
    cPrint.printTexto 17.3, PosLinea, "Donación No."
    cPrint.letraTipo tipoDeLetra, 14
    PosLinea = PosLinea + 0.5
    cPrint.printTexto 17.85, PosLinea, TFA.Serie
    PosLinea = PosLinea + 0.5
    cPrint.printTexto 17.4, PosLinea, SerieFactura
    cPrint.letraTipo tipoDeLetra, 10
    PosLinea = PosLinea + 0.5
    cPrint.printTexto 17.2, PosLinea, "FECHA: " & TFA.Fecha
    cPrint.tipoNegrilla = False
    PosLinea = TPosLinea1
    cPrint.tipoNegrilla = True
    cPrint.letraTipo tipoDeLetra, 7
    cPrint.printTexto 1.6, PosLinea, "Razón Social/Nombres y Apellidos:"
    cPrint.printTexto 18.5, PosLinea, "Identificación: "
    PosLinea = PosLinea + 0.35
    cPrint.tipoNegrilla = False
    
    cPrint.printTexto 18.5, PosLinea, TFA.RUC_CI
    cPrint.printTexto 1.6, PosLinea, TFA.Razon_Social

    PosLinea = PosLinea + 0.35
    cPrint.tipoNegrilla = True
    
    cPrint.printTexto 1.6, PosLinea, "Dirección:"
    cPrint.tipoNegrilla = False
    cPrint.printTexto 3, PosLinea, TFA.DireccionC
    cPrint.printTexto 17.3, PosLinea, "Teléfono: " & TFA.TelefonoC
    PosLinea = PosLinea + 0.35
    cPrint.printTexto 1.6, PosLinea, "Atención: " & TFA.Contacto
    PosLinea = PosLinea + 0.5
    cPrint.letraTipo tipoDeLetra, 7
    Cadena = "Queremos agradecerle por su aporte solidario de USD " & Format$(TFA.Total_MN, "#,##0.00") & ", donación que nos permitirá incrementar la atención a un mayor número de personas en " _
           & "vulnerabilidad alimentaria. Usted es muy importante para nosotros." & vbCrLf _
           & "SU DONACIÓN PUEDE REALIZARLA EN EFECTIVO, DEPOSITO O TRANSFERENCIA BANCARIA A LA CUENTA DE AHORROS BANCO PICHINCHA No.- 3708204100 A NOMBRE DE " _
           & "FUNDACIÓN DE AYUDA SOCIAL BANCO DE ALIMENTOS DE QUITO."
    PosLinea = cPrint.printTextoMultiple(1.6, PosLinea, Cadena, 18.5)
    PosLinea = PosLinea + 1.5
    cPrint.printTexto 4, PosLinea, "______________", PorteDeLetra
    cPrint.printTexto 14, PosLinea, "______________", PorteDeLetra
    PosLinea = PosLinea + 0.4
    cPrint.printTexto 4.4, PosLinea, "RECIBIDO", PorteDeLetra
    cPrint.printTexto 14.2, PosLinea, "ENTREGADO", PorteDeLetra
   'Cuadros Superiores
    cPrint.printCuadro 1.5, TempPosLinea, 20, PosLinea, Negro, "B"
   '==================================================================================================
    AdoDBDet.Close
    AdoDBAbo.Close
    RatonNormal
    cPrint.finalizaImpresion
End Sub

Public Sub Imprimir_FNC(PosInic As Single, _
                        PosLinea1 As Single, _
                        DtaF As Adodc, _
                        DtaD As Adodc, _
                        Optional ReImp As Boolean)
Dim PFil1 As Single
Dim PPFil1 As Single
Dim PAncho As Single
Dim LineasNo As Integer
Dim AltoLetras As Single
Dim CI_RUC_SRI As String
  Printer.Font = TipoArialNarrow
  Printer.FontBold = True
  With DtaF.Recordset
    'MsgBox "Hoja"
     Codigo4 = Format$(.fields("Factura"), "0000000")
     
     RutaOrigen = RutaSistema & "\FORMATOS\" & FA.LogoFactura & ".GIF"
    MsgBox RutaOrigen & vbCrLf & FA.AnchoFactura & vbCrLf & FA.AltoFactura
     PrinterPaint RutaOrigen, PosInic + SetD(38).PosX, PosLinea1 + SetD(38).PosY, FA.AnchoFactura, FA.AltoFactura
     If SetD(1).PosX > 0 And SetD(1).PosY > 0 Then
           'MsgBox PosInic + 0.01 & vbCrLf & PosLinea1 + 0.01 & vbCrLf & AnchoFactura & vbCrLf & AltoFactura
           PrinterPaint LogoTipo, PosInic + 0.01, PosLinea1 + 0.01, 3, 1.5
           If Empresa = NombreComercial Then
              Printer.FontSize = 11
              PrinterTexto PosInic + 3, PosLinea1 + 0.4, Empresa
           Else
              Printer.FontSize = 10
              PrinterTexto PosInic + 3, PosLinea1 + 0.1, Empresa
              Printer.FontSize = 9
              PrinterTexto PosInic + 3, PosLinea1 + 0.55, NombreComercial
           End If
           Printer.FontSize = 8
           PrinterTexto PosInic + 3, PosLinea1 + 0.95, "R.U.C. " & RUC
           PrinterTexto PosInic + 3, PosLinea1 + 1.3, "Dirección: " & Direccion
           Codigo4 = SerieFactura & "-" & Format$(.fields("Factura"), "0000000")
           Printer.FontSize = 6
           If SetD(28).PosX > 0 And SetD(28).PosY > 0 Then
              Printer.FontBold = False
              Cuenta = "Autorización otorgada por el S.R.I. para imprimir por medios Computarizados Facturas, Notas de Venta,"
              PrinterTexto PosInic + SetD(28).PosX, PosLinea1 + SetD(28).PosY, Cuenta
              Cuenta = "Notas de Crédito, No. " & Autorizacion & ", válido hasta el " & Fecha_Vence
              If .fields("T") = "A" Then
                  Cuenta = Cuenta & " - ANULADA"
              ElseIf ReImp Then
                  Cuenta = Cuenta & " - REIMPRESION"
              End If
              If PosInic > 0.01 Then
                 Cuenta = Cuenta & " - COPIA"
              Else
                 Cuenta = Cuenta & " - ORIGINAL"
              End If
              PrinterTexto PosInic + SetD(28).PosX, PosLinea1 + SetD(28).PosY + 0.3, Cuenta
              Printer.FontBold = True
           End If
     End If
     Printer.Font = TipoArialNarrow
      'MsgBox Codigo4
      'Pie de Factura
       If SetD(2).PosX > 0 Then
          Printer.FontSize = SetD(2).Porte
          PrinterTexto PosInic + SetD(2).PosX, PosLinea1 + SetD(2).PosY, Codigo4
       End If
       If SetD(3).PosX > 0 Then
          Printer.FontSize = SetD(3).Porte
          PrinterFields PosInic + SetD(3).PosX, PosLinea1 + SetD(3).PosY, .fields("Fecha")
       End If
       If SetD(31).PosX > 0 Then
          Printer.FontSize = SetD(31).Porte
          PrinterTexto PosInic + SetD(31).PosX, PosLinea1 + SetD(31).PosY, MesesLetras(Month(.fields("Fecha")))
       End If
       If SetD(4).PosX > 0 Then
          Printer.FontSize = SetD(4).Porte
          PrinterVariables PosInic + SetD(4).PosX, PosLinea1 + SetD(4).PosY, Mifecha
       End If
       If SetD(11).PosX > 0 Then
          Printer.FontSize = SetD(11).Porte
          PrinterFields PosInic + SetD(11).PosX, PosLinea1 + SetD(11).PosY, .fields("CI_RUC")
       End If
       If SetD(10).PosX > 0 Then
          Printer.FontSize = SetD(10).Porte
          PrinterFields PosInic + SetD(10).PosX, PosLinea1 + SetD(10).PosY, .fields("Ciudad")
       End If
       If SetD(8).PosX > 0 Then
          Printer.FontSize = SetD(8).Porte
          PrinterFields PosInic + SetD(8).PosX, PosLinea1 + SetD(8).PosY, .fields("Direccion")
       End If
       NivelNo = SinEspaciosIzq(.fields("Direccion"))
       If NivelNo = "" Then NivelNo = Ninguno
       If SetD(29).PosX > 0 Then
          Printer.FontSize = SetD(29).Porte
          PrinterTexto PosInic + SetD(29).PosX, PosLinea1 + SetD(29).PosY, NivelNo
       End If
       NivelNo = MidStrg(.fields("Direccion"), Len(NivelNo) + 1, Len(.fields("Direccion")) - Len(NivelNo) + 1)
       If NivelNo = "" Then NivelNo = Ninguno
       If SetD(30).PosX > 0 Then
          Printer.FontSize = SetD(30).Porte
          PrinterTexto PosInic + SetD(30).PosX, PosLinea1 + SetD(30).PosY, NivelNo
       End If
       If SetD(9).PosX > 0 Then
          Printer.FontSize = SetD(9).Porte
          PrinterFields PosInic + SetD(9).PosX, PosLinea1 + SetD(9).PosY, .fields("Telefono")
       End If
       If SetD(5).PosX > 0 Then
          Printer.FontSize = SetD(5).Porte
          If FA_Educativo Then
             If .fields("Representante") = Ninguno Then
                 PrinterFields PosInic + SetD(5).PosX, PosLinea1 + SetD(5).PosY, .fields("Cliente")
             Else
                 PrinterFields PosInic + SetD(5).PosX, PosLinea1 + SetD(5).PosY, .fields("Representante")
             End If
          Else
              PrinterFields PosInic + SetD(5).PosX, PosLinea1 + SetD(5).PosY, .fields("Cliente")
          End If
       End If
       If SetD(32).PosX > 0 Then
          Printer.FontSize = SetD(32).Porte
          If .fields("Representante") = Ninguno Then
              PrinterTexto PosInic + SetD(32).PosX, PosLinea1 + SetD(32).PosY, "CONSUMIDOR FINAL"
          End If
       End If
       If SetD(36).PosX > 0 Then
          If .fields("Cedula") = Ninguno Then
              Select Case .fields("TD")
                Case "C", "P", "R": CI_RUC_SRI = .fields("CI_RUC")
                Case Else: CI_RUC_SRI = "9999999999999"
              End Select
          Else
              CI_RUC_SRI = .fields("Cedula")
          End If
          Printer.FontSize = SetD(36).Porte
          PrinterTexto PosInic + SetD(36).PosX, PosLinea1 + SetD(36).PosY, CI_RUC_SRI
       End If
       If SetD(5).PosX > 0 Then
          Printer.FontSize = SetD(5).Porte
          PrinterFields PosInic + SetD(5).PosX, PosLinea1 + SetD(5).PosY, .fields("Cliente")
       End If
       If SetD(6).PosX > 0 Then
          Printer.FontSize = SetD(6).Porte
          PrinterFields PosInic + SetD(6).PosX, PosLinea1 + SetD(6).PosY, .fields("Codigo")
       End If
       If SetD(7).PosX > 0 Then
          Printer.FontSize = SetD(7).Porte
          PrinterFields PosInic + SetD(7).PosX, PosLinea1 + SetD(7).PosY, .fields("Grupo")
       End If
       If SetD(13).PosX > 0 Then
          Printer.FontSize = SetD(13).Porte
          PrinterFields PosInic + SetD(13).PosX, PosLinea1 + SetD(13).PosY, .fields("Email")
       End If
      'Pie de Factura
        If SetD(22).PosX > 0 Then
           Printer.FontSize = SetD(22).Porte
           PrinterVariables PosInic + SetD(22).PosX, PosLinea1 + SetD(22).PosY, SubTotal_NC
        End If
'''        If SetD(23).PosX > 0 Then
'''           Printer.FontSize = SetD(23).Porte
'''           PrinterFields PosInic + SetD(23).PosX, PosLinea1 + SetD(23).PosY, .Fields("Con_IVA")
'''        End If
'''        If SetD(24).PosX > 0 Then
'''           Printer.FontSize = SetD(24).Porte
'''           PrinterFields PosInic + SetD(24).PosX, PosLinea1 + SetD(24).PosY, .Fields("Sin_IVA")
'''        End If
        If SetD(25).PosX > 0 Then
           Printer.FontSize = SetD(25).Porte
           PrinterVariables PosInic + SetD(25).PosX, PosLinea1 + SetD(25).PosY, IVA_NC
        End If
        If SetD(26).PosX > 0 Then
           Printer.FontSize = SetD(26).Porte
           PrinterVariables PosInic + SetD(26).PosX, PosLinea1 + SetD(26).PosY, CInt(Porc_IVA * 100)
        End If
        If SetD(27).PosX > 0 Then
           Printer.FontSize = SetD(27).Porte
           PrinterVariables PosInic + SetD(27).PosX, PosLinea1 + SetD(27).PosY, SubTotal_NC + IVA_NC
        End If
'''        If SetD(34).PosX > 0 Then
'''           Printer.FontSize = SetD(34).Porte
'''           PrinterVariables PosInic + SetD(34).PosX, PosLinea1 + SetD(34).PosY, CCur(SaldoPendiente)
'''        End If
'''        If SetD(33).PosX > 0 Then
'''           Printer.FontSize = SetD(33).Porte
'''           PrinterVariables PosInic + SetD(33).PosX, PosLinea1 + SetD(33).PosY, CCur(Diferencia)
'''        End If
        If SetD(37).PosX > 0 Then
           Printer.FontSize = SetD(37).Porte
           PrinterTexto PosInic + SetD(37).PosX, PosLinea1 + SetD(37).PosY, CodigoDelBanco
        End If
'''        If SetD(40).PosX > 0 Then
'''           Printer.FontSize = SetD(40).Porte
'''           PrinterTexto PosInic + SetD(40).PosX, PosLinea1 + SetD(40).PosY, .Fields("Hora")
'''        End If
        'MsgBox Printer.ScaleWidth & vbCrLf & Printer.ScaleHeight & vbCrLf _
        '& "FAM: " & vbCrLf _
        '& PosInic & vbCrLf & PosLinea1 & vbCrLf & Codigo4 & vbCrLf & Diferencia & vbCrLf & SaldoPendiente
        If SetD(64).PosX > 0 Then
           Printer.FontSize = SetD(64).Porte
           If .fields("Representante") <> Ninguno Then
               PrinterTexto PosInic + SetD(64).PosX, PosLinea1 + SetD(64).PosY, .fields("Cliente")
           End If
        End If
  End With
  Printer.Font = TipoArialNarrow
 'Detalle de la Factura
  AltoLetras = 0.4
  Printer.FontSize = SetD(17).Porte
  AltoLetras = Redondear(Printer.TextHeight("H") - 0.05, 2)
  'MsgBox AltoLetras
  With DtaD.Recordset
       PFil1 = SetD(14).PosY
       PAncho = SetD(18).PosX
       If .RecordCount > 0 Then
          .MoveFirst
           Producto = .fields("Producto")
           CodigoInv = .fields("Codigo")
           ValorUnit = .fields("Precio")
           CodigoP = ""
           Cantidad = 0
           SubTotal = 0
           SubTotal_IVA = 0
           Do While Not .EOF
              If CodigoInv <> .fields("Codigo") Or ValorUnit <> .fields("Precio") Then
                 CodigoP = MidStrg(CodigoP, 1, Len(CodigoP) - 2)
                 Producto = Producto & " " & CodigoP
                 If SetD(16).PosX > 0 Then
                    Printer.FontSize = SetD(16).Porte
                    PrinterTexto PosInic + SetD(16).PosX, PosLinea1 + PFil1, CStr(Cantidad)
                 End If
                 If SetD(15).PosX > 0 Then
                    Printer.FontSize = SetD(15).Porte
                    PrinterTexto PosInic + SetD(15).PosX, PosLinea1 + PFil1, CodigoInv
                 End If
                 If SetD(17).PosX > 0 Then
                    Printer.FontSize = SetD(17).Porte
                    LineasNo = 0
                    LineasNo = PrinterLineasMayor(PosInic + SetD(17).PosX, PosLinea1 + PFil1, Producto, PAncho)
                    PFil1 = PFil1 + (LineasNo * 0.35)
                    PFil1 = PFil1 - 0.35
                 End If
                 If SetD(20).PosX > 0 Then
                    Printer.FontSize = SetD(20).Porte
                    'PrinterVariables PosInic + SetD(20).PosX, PosLinea1 + PFil1, SubTotal
                 End If
                 If SetD(19).PosX > 0 Then
                    Printer.FontSize = SetD(19).Porte
                    'PrinterVariables PosInic + SetD(19).PosX, PosLinea1 + PFil1, ValorUnit
                 End If
                 Producto = .fields("Producto")
                 CodigoInv = .fields("Codigo")
                 ValorUnit = .fields("Precio")
                 PFil1 = PFil1 + Printer.TextHeight("H")
                 'PFil1 = PFil1 + AltoLetras
                 CodigoP = ""
                 Cantidad = 0
                 SubTotal = 0
                 SubTotal_IVA = 0
              End If
              SubTotal = SubTotal + .fields("Total")
              Cantidad = Cantidad + .fields("Cantidad")
              CodigoP = CodigoP & .fields("Mes") & "-" & .fields("Ticket") & ", "
             .MoveNext
           Loop
           CodigoP = MidStrg(CodigoP, 1, Len(CodigoP) - 2)
           Producto = Producto & CodigoP
           If SetD(16).PosX > 0 Then
              Printer.FontSize = SetD(16).Porte
              PrinterTexto PosInic + SetD(16).PosX, PosLinea1 + PFil1, CStr(Cantidad)
           End If
           If SetD(15).PosX > 0 Then
              Printer.FontSize = SetD(15).Porte
              PrinterTexto PosInic + SetD(15).PosX, PosLinea1 + PFil1, CodigoInv
           End If
           If SetD(17).PosX > 0 Then
              Printer.FontSize = SetD(17).Porte
              LineasNo = 0
              LineasNo = PrinterLineasMayor(PosInic + SetD(17).PosX, PosLinea1 + PFil1, Producto, PAncho)
              PFil1 = PFil1 + (LineasNo * 0.35)
              PFil1 = PFil1 - 0.35
           End If
           If SetD(20).PosX > 0 Then
              Printer.FontSize = SetD(20).Porte
              PrinterVariables PosInic + SetD(20).PosX, PosLinea1 + PFil1, SubTotal_NC
           End If
''           If SetD(19).PosX > 0 Then
''              Printer.FontSize = SetD(19).Porte
''              PrinterVariables PosInic + SetD(19).PosX, PosLinea1 + PFil1, SubTotal_NC
''           End If
       End If
   End With
  'Pie de Factura
   With DtaF.Recordset
   End With
End Sub

'En la variable de tipo: TFA.Cod_CxC se ingresa el codigo de CxC de la factura
Public Sub Lineas_De_CxC(TFA As Tipo_Facturas)
Dim AdoLineaDB As ADODB.Recordset
Dim FechaCxC As String
  
  Cant_Item_FA = 1
  Cant_Item_PV = 1
  TFA.Cta_CxP = Ninguno
  TFA.Cta_Venta = Ninguno
  If Not IsDate(TFA.Fecha) Then TFA.Fecha = FechaSistema
  If Not IsDate(TFA.Fecha_NC) Then TFA.Fecha_NC = FechaSistema
  If TFA.TC = "NC" Then FechaCxC = TFA.Fecha_NC Else FechaCxC = TFA.Fecha

  If Len(TFA.Serie) < 6 Then TFA.Serie = Leer_Campo_Empresa("Serie_FA")
  
  'MsgBox TFA.Cod_CxC
  If Len(TFA.Cod_CxC) > 1 Then TFA.Serie = Ninguno
  sSQL = "SELECT Concepto, Logo_Factura, Largo, Ancho, Espacios, Pos_Factura, Fact_Pag, Pos_Y_Fact, Serie, Autorizacion, Vencimiento, Fecha, Secuencial, " _
       & "ItemsxFA, Codigo, Fact, CxC, Cta_Venta, CxC_Anterior, Imp_Mes, Nombre_Establecimiento, Direccion_Establecimiento, Telefono_Estab, Logo_Tipo_Estab " _
       & "FROM Catalogo_Lineas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND TL <> " & Val(adFalse) & " " _
       & "AND Fecha <= #" & BuscarFecha(FechaCxC) & "# " _
       & "AND Vencimiento >= #" & BuscarFecha(FechaCxC) & "# "
  If Len(TFA.TC) >= 2 Then sSQL = sSQL & "AND Fact = '" & TFA.TC & "' "
  If Len(TFA.Serie) = 6 Then sSQL = sSQL & "AND Serie = '" & TFA.Serie & "' "
  If Len(TFA.Autorizacion) >= 6 Then sSQL = sSQL & "AND Autorizacion = '" & TFA.Autorizacion & "' "
  If Len(TFA.Cod_CxC) > 1 Then sSQL = sSQL & "AND '" & TFA.Cod_CxC & "' IN (Concepto, Codigo, CxC) "
  sSQL = sSQL & "ORDER BY Codigo "
  Select_AdoDB AdoLineaDB, sSQL
 'MsgBox sSQL
  With AdoLineaDB
   If .RecordCount > 0 Then
       TFA.CxC_Clientes = .fields("Concepto")
       TFA.LogoFactura = .fields("Logo_Factura")
       TFA.AltoFactura = .fields("Largo")
       TFA.AnchoFactura = .fields("Ancho")
       TFA.EspacioFactura = .fields("Espacios")
       TFA.Pos_Factura = .fields("Pos_Factura")
       TFA.Pos_Copia = .fields("Pos_Y_Fact")
       TFA.CantFact = .fields("Fact_Pag")
     
      'Datos para grabar automaticamente
       TFA.TC = .fields("Fact")
       TFA.Serie = .fields("Serie")
       TFA.Autorizacion = .fields("Autorizacion")
       TFA.Fecha_Aut = .fields("Fecha")
       TFA.Vencimiento = .fields("Vencimiento")
       TFA.Cta_CxP = .fields("CxC")
       TFA.Cta_Venta = .fields("Cta_Venta")
       TFA.Cta_CxP_Anterior = .fields("CxC_Anterior")
       TFA.Cod_CxC = .fields("Codigo")
       TFA.Imp_Mes = .fields("Imp_Mes")
       TFA.DireccionEstab = .fields("Direccion_Establecimiento")
       TFA.NombreEstab = .fields("Nombre_Establecimiento")
       TFA.TelefonoEstab = .fields("Telefono_Estab")
       TFA.LogoTipoEstab = RutaSistema & "\LOGOS\" & .fields("Logo_Tipo_Estab") & ".jpg"
       If TFA.TC = "NC" Then
          TFA.Serie_NC = .fields("Serie")
          TFA.Autorizacion_NC = .fields("Autorizacion")
       Else
          Cant_Item_FA = .fields("ItemsxFA")
          Cant_Item_PV = .fields("ItemsxFA")
          Cta_Cobrar = .fields("CxC")
          Cta_Ventas = .fields("Cta_Venta")
          CodigoL = .fields("Codigo")
          TipoFactura = .fields("Fact")
       End If
   Else
       MsgBox "Linea CxC No Asignados o fuera de fecha"
   End If
  End With
  AdoLineaDB.Close
 'MsgBox TFA.Cta_CxP
  If TFA.Cta_CxP <> Ninguno Then
     ReDim ExisteCtas(4) As String
     ExisteCtas(0) = TFA.Cta_CxP
     ExisteCtas(1) = Cta_CajaG
     ExisteCtas(2) = Cta_CajaGE
     ExisteCtas(3) = Cta_CajaBA
     VerSiExisteCta ExisteCtas
  End If
 
  If Cant_Item_FA <= 0 Then Cant_Item_FA = 50
  If Cant_Item_PV <= 0 Then Cant_Item_PV = 50
  Cadena = "Esta Ingresando Comprobantes Caducados" & vbCrLf & vbCrLf _
         & "La Fecha Tope de emision es: " & FA.Vencimiento & vbCrLf
  If CFechaLong(TFA.Fecha) > CFechaLong(TFA.Vencimiento) Then MsgBox Cadena
End Sub

'''Public Sub Lineas_De_CxC(TFA As Tipo_Facturas)
'''Dim AdoLineaDB As ADODB.Recordset
'''
'''  Cant_Item_FA = 1
'''  Cant_Item_PV = 1
'''  TFA.Cta_CxP = Ninguno
'''  TFA.Cta_Venta = Ninguno
'''  If TFA.Fecha = "" Then TFA.Fecha = FechaSistema
'''  If TFA.Fecha_NC = "" Then TFA.Fecha_NC = FechaSistema
''' 'MsgBox LineaCxC
'''  sSQL = "SELECT * " _
'''       & "FROM Catalogo_Lineas " _
'''       & "WHERE Item = '" & NumEmpresa & "' " _
'''       & "AND Periodo = '" & Periodo_Contable & "' " _
'''       & "AND Fact = '" & TFA.TC & "' "
'''  If TFA.TC = "NC" Then
'''     sSQL = sSQL _
'''          & "AND Fecha <= #" & BuscarFecha(TFA.Fecha_NC) & "# " _
'''          & "AND Vencimiento >= #" & BuscarFecha(TFA.Fecha_NC) & "# "
'''  Else
'''     sSQL = sSQL _
'''          & "AND Fecha <= #" & BuscarFecha(TFA.Fecha) & "# " _
'''          & "AND Vencimiento >= #" & BuscarFecha(TFA.Fecha) & "# "
'''  End If
'''  sSQL = sSQL & "ORDER BY Codigo "
'''  Generar_Consulta_SQL "Catalogo_Lineas", sSQL
'''  Select_AdoDB AdoLineaDB, sSQL
'''  With AdoLineaDB
'''   If .RecordCount > 0 Then
'''      .MoveFirst
'''      .Find ("Concepto = '" & TFA.Cod_CxC & "' ")
'''       If Not .EOF Then
'''          GoTo Encontro_CxC
'''       Else
'''          .MoveFirst
'''          .Find ("Codigo = '" & TFA.Cod_CxC & "' ")
'''           If Not .EOF Then
'''              GoTo Encontro_CxC
'''           Else
'''             .MoveFirst
'''             .Find ("CxC = '" & TFA.Cod_CxC & "' ")
'''              If Not .EOF Then
'''                 GoTo Encontro_CxC
'''              Else
'''                 MsgBox "Cuenta No Asignada"
'''                 GoTo No_Encontro_CxC
'''              End If
'''           End If
'''       End If
'''Encontro_CxC:
'''            TFA.CxC_Clientes = .Fields("Concepto")
'''            TFA.LogoFactura = .Fields("Logo_Factura")
'''            TFA.AltoFactura = .Fields("Largo")
'''            TFA.AnchoFactura = .Fields("Ancho")
'''            TFA.EspacioFactura = .Fields("Espacios")
'''            TFA.Pos_Factura = .Fields("Pos_Factura")
'''            TFA.Pos_Copia = .Fields("Pos_Y_Fact")
'''            TFA.CantFact = .Fields("Fact_Pag")
'''
'''           'Datos para grabar automaticamente
'''            TFA.TC = TrimStrg(UCaseStrg(.Fields("Fact")))
'''            TFA.Serie = .Fields("Serie")
'''            TFA.Autorizacion = .Fields("Autorizacion")
'''            TFA.Fecha_Aut = .Fields("Fecha")
'''            TFA.Vencimiento = .Fields("Vencimiento")
'''            TFA.Cta_CxP = .Fields("CxC")
'''            TFA.Cta_Venta = .Fields("Cta_Venta")
'''            TFA.Cta_CxP_Anterior = .Fields("CxC_Anterior")
'''            TFA.Cod_CxC = .Fields("Codigo")
'''            TFA.Imp_Mes = .Fields("Imp_Mes")
'''            TFA.DireccionEstab = .Fields("Direccion_Establecimiento")
'''            TFA.NombreEstab = .Fields("Nombre_Establecimiento")
'''            TFA.TelefonoEstab = .Fields("Telefono_Estab")
'''            TFA.LogoTipoEstab = RutaSistema & "\LOGOS\" & .Fields("Logo_Tipo_Estab") & ".jpg"
'''            If TFA.TC = "NC" Then
'''               TFA.Serie_NC = .Fields("Serie")
'''               TFA.Autorizacion_NC = .Fields("Autorizacion")
'''            Else
'''               Cant_Item_FA = .Fields("ItemsxFA")
'''               Cant_Item_PV = .Fields("ItemsxFA")
'''               Cta_Cobrar = .Fields("CxC")
'''               Cta_Ventas = .Fields("Cta_Venta")
'''               CodigoL = .Fields("Codigo")
'''               TipoFactura = .Fields("Fact")
'''            End If
'''No_Encontro_CxC:
'''   Else
'''       MsgBox "Codigos No Asignados o fuera de fecha"
'''   End If
'''  End With
''' 'MsgBox TFA.Cta_CxP
'''  If TFA.Cta_CxP <> Ninguno Then
'''     ReDim ExisteCtas(4) As String
'''     ExisteCtas(0) = TFA.Cta_CxP
'''     ExisteCtas(1) = Cta_CajaG
'''     ExisteCtas(2) = Cta_CajaGE
'''     ExisteCtas(3) = Cta_CajaBA
'''     VerSiExisteCta ExisteCtas
'''  End If
''' 'AdoLineaDB.Close
'''  If Cant_Item_FA <= 0 Then Cant_Item_FA = 15
'''  If Cant_Item_PV <= 0 Then Cant_Item_PV = 15
'''  Cadena = "Esta Ingresando Comprobantes Caducados" & vbCrLf & vbCrLf _
'''         & "La Fecha Tope de emision es: " & FA.Vencimiento & vbCrLf
'''  If CFechaLong(TFA.Fecha) > CFechaLong(TFA.Vencimiento) Then MsgBox Cadena
'''End Sub


'''Public Sub Lineas_De_NC(Serie_NC As String, Fecha_Emision As String)
'''Dim AdoLineaDB As ADODB.Recordset
'''
'''  Cant_Item_FA = 1
'''  Cant_Item_PV = 1
''' 'MsgBox LineaCxC
'''  sSQL = "SELECT * " _
'''       & "FROM Catalogo_Lineas " _
'''       & "WHERE Item = '" & NumEmpresa & "' " _
'''       & "AND Periodo = '" & Periodo_Contable & "' " _
'''       & "AND Fecha <= #" & BuscarFecha(Fecha_Emision) & "# " _
'''       & "AND Vencimiento >= #" & BuscarFecha(Fecha_Emision) & "# " _
'''       & "AND Serie = '" & Serie_NC & "' " _
'''       & "AND Fact = 'NC' " _
'''       & "ORDER BY Codigo "
'''  Select_AdoDB AdoLineaDB, sSQL
'''  With AdoLineaDB
'''   If .RecordCount > 0 Then
'''       LogoFactura = .Fields("Logo_Factura")
'''       AltoFactura = .Fields("Largo")
'''       AnchoFactura = .Fields("Ancho")
'''       EspacioFactura = .Fields("Espacios")
'''       Pos_Factura = .Fields("Pos_Factura")
'''       Pos_Copia = .Fields("Pos_Y_Fact")
'''       CantFact = .Fields("Fact_Pag")
'''       TipoFactura = UCaseStrg(.Fields("Fact"))
'''       CodigoL = .Fields("Codigo")
'''       TipoFactura = .Fields("Fact")
'''    Else
'''       LogoFactura = "NINGUNO"
'''       AltoFactura = 0
'''       AnchoFactura = 0
'''       EspacioFactura = 0
'''       Pos_Factura = 0
'''       Pos_Copia = 0
'''       CantFact = 0
'''       TipoFactura = Ninguno
'''       TipoFactura = Ninguno
'''   End If
'''  End With
'''  AdoLineaDB.Close
'''End Sub

Public Sub ImprimirVentasCosto(Datas As Adodc, _
                               FinDoc As Boolean, _
                               FormaImp As Byte, _
                               SizeLetra As Integer, _
                               Optional EsCampoCorto As Boolean)
On Error GoTo Errorhandler

Mensajes = "Seguro de Imprimir en:" & vbCrLf & Printer.DeviceName & "?"
Titulo = "IMPRESION"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
RatonReloj
InicioX = 0.5: InicioY = 0
DataAnchoCampos InicioX, Datas, SizeLetra, TipoTimes, Orientacion_Pagina, EsCampoCorto
Pagina = 1
PonerLinea = False
'Iniciamos la impresion
Printer.FontBold = False
Total = 0: Abono = 0: Saldo = 0
With Datas.Recordset
If .RecordCount > 0 Then
    .MoveFirst
     EncabezadoData Datas
     Printer.FontSize = SizeLetra
     Do While Not .EOF
        PrinterAllFields CantCampos, PosLinea, Datas, True, False
        PosLinea = PosLinea + 0.36
        If PosLinea >= LimiteAlto Then
           Imprimir_Linea_H PosLinea, Ancho(0), Ancho(CantCampos)
           Printer.NewPage
           EncabezadoData Datas
           Printer.FontSize = SizeLetra
        End If
        Total = Total + .fields("Ventas")
        Abono = Abono + .fields("PVP")
        Saldo = Saldo + .fields("Costos")
       .MoveNext
     Loop
End If
End With
Imprimir_Linea_H PosLinea, InicioX, Ancho(CantCampos), Negro, True
PosLinea = PosLinea + 0.1
PrinterVariables Ancho(3), PosLinea, Total
PrinterVariables Ancho(4), PosLinea, Abono
PrinterVariables Ancho(5), PosLinea, Saldo
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

Public Function Existe_Factura(TFA As Tipo_Facturas) As Boolean
Dim AdoDBFA As ADODB.Recordset
Dim Respuesta As Boolean
   RatonReloj
   Respuesta = False
  'Consultamos si existe la Factura
   sSQL = "SELECT TC, Serie, Factura " _
        & "FROM Facturas " _
        & "WHERE Factura = " & TFA.Factura & " " _
        & "AND TC = '" & TFA.TC & "' " _
        & "AND Serie = '" & TFA.Serie & "' " _
        & "AND Item = '" & NumEmpresa & "' " _
        & "AND Periodo = '" & Periodo_Contable & "' "
   Select_AdoDB AdoDBFA, sSQL
   If AdoDBFA.RecordCount > 0 Then Respuesta = True
   AdoDBFA.Close
   RatonNormal
   Existe_Factura = Respuesta
End Function

Public Sub Facturas_Impresas(TFA As Tipo_Facturas)
    SQL2 = "UPDATE Facturas " _
         & "SET P = 1 " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND TC = '" & TFA.TC & "' " _
         & "AND Serie = '" & TFA.Serie & "' " _
         & "AND P = 0 "
    If (TFA.Desde + TFA.Hasta) > 0 And (TFA.Desde <= TFA.Hasta) Then
        SQL2 = SQL2 & "AND Factura BETWEEN " & TFA.Desde & " and " & TFA.Hasta & " "
    Else
        SQL2 = SQL2 & "AND Factura = " & TFA.Factura & " "
    End If
    Ejecutar_SQL_SP SQL2
End Sub

Public Sub Imprimir_Comprobante_Caja(FTA As Tipo_Abono)
Dim AdoReciboDB  As ADODB.Recordset
Dim Valor_Str As String
Dim Codigo_M As String
Dim Codigo_A As String
Dim Codigo_P As String
Dim Codigo_S As String
Dim Saldo_P As Currency
  
  If Leer_Campo_Empresa("Imp_Recibo_Caja") Then
     TRecibo.Tipo_Recibo = "I"
     Saldo_P = 0
     sSQL = "SELECT Factura,Saldo_MN " _
          & "FROM Facturas " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "AND TC = '" & FTA.TP & "' " _
          & "AND Serie = '" & FTA.Serie & "' " _
          & "AND Autorizacion = '" & FTA.Autorizacion & "' " _
          & "AND Factura = " & FTA.Factura & " " _
          & "AND CodigoC = '" & FTA.CodigoC & "' "
     Select_AdoDB AdoReciboDB, sSQL
     If AdoReciboDB.RecordCount > 0 Then Saldo_P = AdoReciboDB.fields("Saldo_MN")
     AdoReciboDB.Close
     
     sSQL = "SELECT Factura,Mes,Ticket,Codigo,Producto " _
          & "FROM Detalle_Factura " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "AND TC = '" & FTA.TP & "' " _
          & "AND Serie = '" & FTA.Serie & "' " _
          & "AND Autorizacion = '" & FTA.Autorizacion & "' " _
          & "AND Factura = " & FTA.Factura & " " _
          & "AND CodigoC = '" & FTA.CodigoC & "' " _
          & "ORDER BY Codigo,Ticket,Mes "
     Select_AdoDB AdoReciboDB, sSQL
     TRecibo.Concepto = ""
     With AdoReciboDB
      If .RecordCount > 0 Then
          Codigo_M = .fields("Mes")
          Codigo_A = .fields("Ticket")
          Codigo_P = .fields("Codigo")
          If .fields("Ticket") <> Ninguno Then
              Codigo_S = .fields("Ticket") & ", " & .fields("Producto") & " "
          Else
              Codigo_S = .fields("Producto") & " "
          End If
          Do While Not .EOF
             If Codigo_A <> .fields("Ticket") Then
                TRecibo.Concepto = TRecibo.Concepto & Codigo_S & vbCrLf
                Codigo_P = .fields("Codigo")
                If .fields("Ticket") <> Ninguno Then
                    Codigo_S = .fields("Ticket") & ", " & .fields("Producto") & " "
                Else
                    Codigo_S = ", " & .fields("Producto") & " "
                End If
                Codigo_A = .fields("Ticket")
             End If
             If Codigo_P <> .fields("Codigo") Then
                TRecibo.Concepto = TRecibo.Concepto & Codigo_S & vbCrLf
                Codigo_P = .fields("Codigo")
                If .fields("Ticket") <> Ninguno Then
                    Codigo_S = .fields("Ticket") & ", " & .fields("Producto") & " "
                Else
                    Codigo_S = ", " & .fields("Producto") & " "
                End If
             End If
             If .fields("Mes") <> Ninguno Then Codigo_S = Codigo_S & " " & .fields("Mes")
            .MoveNext
          Loop
          TRecibo.Concepto = TRecibo.Concepto & Codigo_S & vbCrLf
      End If
     End With
     AdoReciboDB.Close
     
     sSQL = "SELECT Serie,Factura,Abono,Cheque,Banco,Cta " _
          & "FROM Trans_Abonos " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "AND TP = '" & FTA.TP & "' " _
          & "AND Fecha = #" & BuscarFecha(FTA.Fecha) & "# " _
          & "AND Serie = '" & FTA.Serie & "' " _
          & "AND Autorizacion = '" & FTA.Autorizacion & "' " _
          & "AND Factura = " & FTA.Factura & " " _
          & "AND CodigoC = '" & FTA.CodigoC & "' "
     Select_AdoDB AdoReciboDB, sSQL
     With AdoReciboDB
      If .RecordCount > 0 Then
          TRecibo.Cobrado_a = FTA.Recibi_de
          TRecibo.Recibo_No = FTA.Recibo_No
          TRecibo.Fecha = TA.Fecha
          TRecibo.Total = 0
          Do While Not .EOF
             TRecibo.Total = TRecibo.Total + .fields("Abono")
             Valor_Str = Format$(.fields("Abono"), "#,##0.00")
             TRecibo.Concepto = TRecibo.Concepto _
                    & "CI/RUC/Codigo: " & FTA.CI_RUC_Cli & ", " _
                    & "Factura No. " & .fields("Serie") & "-" & Format(.fields("Factura"), "000000000") & ", " _
                    & .fields("Cheque") & vbCrLf _
                    & .fields("Banco") & ", " _
                    & .fields("Cta") & ", USD " _
                    & Format$(.fields("Abono"), "#,##0.00") & vbCrLf
            .MoveNext
          Loop
          If Saldo_P <> 0 Then TRecibo.Concepto = TRecibo.Concepto & "Saldo Pendiente USD " & Format$(Saldo_P, "#,##0.00") & vbCrLf
      End If
     End With
     AdoReciboDB.Close
     Imprimir_Recibo_Caja TRecibo
  End If
End Sub

Public Sub Imprimir_Comprobante_Caja_Por_Cliente(AdoRecibo As Adodc, FTA As Tipo_Abono, Optional Fecha_Abono As String)
Dim SerieF As String
Dim Valor_Str As String
Dim Codigo_M As String
Dim Codigo_A As String
Dim Codigo_P As String
Dim Codigo_S As String
Dim List_Fact As String
Dim Saldo_P As Currency
Dim Num_Fact() As Long
Dim Fact_No As Long
Dim Cant_Fact As Integer
     List_Fact = ""
     Cant_Fact = 0
     If FTA.Tipo_Recibo = "" Then FTA.Tipo_Recibo = "C"
     TRecibo.Concepto = ""
     Saldo_P = 0
     sSQL = "SELECT F.RUC_CI,F.TC,F.Serie,F.Fecha,F.Factura,F.Total_MN,TA.Fecha As Fecha_Ab,TA.Cta,TA.Banco,TA.Cheque,TA.Abono,TA.Recibo_No, F.CodigoU " _
          & "FROM Trans_Abonos As TA, Facturas As F " _
          & "WHERE TA.Item = '" & NumEmpresa & "' " _
          & "AND TA.Periodo = '" & Periodo_Contable & "' " _
          & "AND TA.TP = '" & FTA.TP & "' "
     Select Case FTA.Tipo_Recibo
       Case "F": sSQL = sSQL _
                      & "AND TA.Serie = '" & FTA.Serie & "' " _
                      & "AND TA.Factura = " & FTA.Factura & " "
       Case "R": sSQL = sSQL & "AND TA.Recibo_No = '" & FTA.Recibo_No & "' "
       Case "D": sSQL = sSQL & "AND TA.Fecha = #" & BuscarFecha(FTA.Fecha) & "# "
       Case "C": sSQL = sSQL & "AND TA.CodigoC = '" & FTA.CodigoC & "' "
     End Select
     If Fecha_Abono <> "" Then sSQL = sSQL & "AND TA.Fecha = #" & BuscarFecha(Fecha_Abono) & "# "
     sSQL = sSQL _
          & "AND TA.Item = F.Item " _
          & "AND TA.Periodo = F.Periodo " _
          & "AND TA.TP = F.TC " _
          & "AND TA.Serie = F.Serie " _
          & "AND TA.Factura = F.Factura " _
          & "AND TA.CodigoC = F.CodigoC " _
          & "AND TA.Autorizacion = F.Autorizacion " _
          & "ORDER BY TA.TP,F.Serie,F.Factura "
     Select_Adodc AdoRecibo, sSQL
     With AdoRecibo.Recordset
      If .RecordCount > 0 Then
          TRecibo.Cobrado_a = FTA.Recibi_de
          TRecibo.Fecha = FTA.Fecha
          TRecibo.CodUsuario = .fields("CodigoU")
          TRecibo.Total = .fields("Total_MN")
          Fact_No = .fields("Factura")
          SerieF = .fields("Serie")
          TRecibo.SubTotal = 0
          TRecibo.IVA = 0
          TRecibo.Recibo_No = .fields("Recibo_No")
          If Len(.fields("RUC_CI")) > 1 Then TRecibo.Concepto = TRecibo.Concepto & "COD: " & .fields("RUC_CI") & " " & vbCrLf
          Do While Not .EOF
             If Fact_No <> .fields("Factura") Or SerieF <> .fields("Serie") Then
                TRecibo.Total = TRecibo.Total + .fields("Total_MN")
                Fact_No = .fields("Factura")
                SerieF = .fields("Serie")
             End If
             TRecibo.SubTotal = TRecibo.SubTotal + .fields("Abono")
             TRecibo.Concepto = TRecibo.Concepto _
                              & .fields("Fecha_Ab") & ", " _
                              & .fields("Cta") & ", " _
                              & .fields("Serie") & ", " _
                              & .fields("Factura") & ", " _
                              & .fields("Banco") & ", " _
                              & .fields("Cheque") & ", USD " _
                              & Format$(.fields("Abono"), "#,##0.00") & " " & vbCrLf
            .MoveNext
          Loop
          TRecibo.Saldo = TRecibo.Total - TRecibo.SubTotal
          TRecibo.Tipo_Recibo = "I"
          Imprimir_Recibo_Caja TRecibo
      End If
     End With

End Sub

'''Public Sub Actualizar_Pagos()
'''Dim lSql As String
'''  'MsgBox Estudiante_DBF.codest
'''  lSql = "UPDATE " & Dato_DBF.Nuevos & " " _
'''       & "SET cedular = '" & Estudiante_DBF.cedular & "', " _
'''       & "fonopaga = '" & Estudiante_DBF.fonopaga & "', " _
'''       & "pagador = '" & Estudiante_DBF.pagador & "', " _
'''       & "direcpaga = '" & Estudiante_DBF.direcpaga & "', " _
'''       & "pagado = 'S' " _
'''       & "WHERE cedula = '" & Estudiante_DBF.codest & "' "
'''  ConectarDataExecute lSql
'''
'''  lSql = "UPDATE " & Dato_DBF.Antiguos & " " _
'''       & "SET cedular = '" & Estudiante_DBF.cedular & "', " _
'''       & "fonopaga = '" & Estudiante_DBF.fonopaga & "', " _
'''       & "pagador = '" & Estudiante_DBF.pagador & "', " _
'''       & "direcpaga = '" & Estudiante_DBF.direcpaga & "', " _
'''       & "pagado = 'S' " _
'''       & "WHERE cedula = '" & Estudiante_DBF.codest & "' "
'''  ConectarDataExecute lSql
'''
'''  lSql = "UPDATE " & Dato_DBF.Actuales & " " _
'''       & "SET cedular = '" & Estudiante_DBF.cedular & "', " _
'''       & "fonopaga = '" & Estudiante_DBF.fonopaga & "', " _
'''       & "pagador = '" & Estudiante_DBF.pagador & "', " _
'''       & "direcpaga = '" & Estudiante_DBF.direcpaga & "', " _
'''       & "pagado = 'S' " _
'''       & "WHERE cedula = '" & Estudiante_DBF.codest & "' "
'''  ConectarDataExecute lSql
'''
''' End Sub

Public Sub Actualizar_Saldos_DBF()
Dim lSql As String
  MsgBox Estudiante_DBF.codest
  lSql = "UPDATE " & Dato_DBF.Nuevos & " " _
       & "SET pagado = 'S' " _
       & "WHERE cedula = '" & Estudiante_DBF.cedula & "' "
  ConectarDataExecute lSql
  
  lSql = "UPDATE " & Dato_DBF.Actuales & " " _
       & "SET pagado = 'S' " _
       & "WHERE cedula = '" & Estudiante_DBF.cedula & "' "
  ConectarDataExecute lSql
  
  lSql = "UPDATE " & Dato_DBF.Antiguos & " " _
       & "SET pagado = 'S' " _
       & "WHERE cedula = '" & Estudiante_DBF.cedula & "' "
  ConectarDataExecute lSql
 End Sub

Public Sub Listar_Productos(DCArticulos As DataCombo, _
                            AdoArticulos As Adodc, _
                            Optional OpcServicio As Boolean, _
                            Optional PatronDeBusqueda As String, _
                            Optional NombreMarca As String)
   If Cod_Marca <> Ninguno Then
      If SQL_Server Then
         sSQL = "UPDATE Catalogo_Productos " _
              & "SET Marca = '" & NombreMarca & "' " _
              & "FROM Catalogo_Productos As CP,Trans_Kardex As TK "
      Else
         sSQL = "UPDATE Catalogo_Productos As CP,Trans_Kardex As TK " _
              & "SET CP.Marca = '" & NombreMarca & "' "
      End If
      sSQL = sSQL _
           & "WHERE CP.Item = '" & NumEmpresa & "' " _
           & "AND CP.Periodo = '" & Periodo_Contable & "' " _
           & "AND TK.CodMarca = '" & Cod_Marca & "' " _
           & "AND CP.Codigo_Inv = TK.Codigo_Inv " _
           & "AND CP.Item = TK.Item " _
           & "AND CP.Periodo = TK.Periodo "
      Ejecutar_SQL_SP sSQL
   End If
   
   sSQL = "SELECT Producto,Codigo_Inv,Codigo_Barra " _
        & "FROM Catalogo_Productos " _
        & "WHERE TC = 'P' " _
        & "AND Item = '" & NumEmpresa & "' " _
        & "AND Periodo = '" & Periodo_Contable & "' " _
        & "AND INV <> " & Val(adFalse) & " "
   If OpcServicio Then sSQL = sSQL & "AND LEN(Cta_Inventario) <= 1 "
   If PatronDeBusqueda <> "" Then sSQL = sSQL & "AND Producto LIKE '%" & PatronDeBusqueda & "%' "
   If NombreMarca <> "" Then sSQL = sSQL & "AND Marca LIKE '%" & NombreMarca & "%' "
   sSQL = sSQL & "ORDER BY Producto,Codigo_Inv,Codigo_Barra "
  'MsgBox sSQL
   SelectDB_Combo DCArticulos, AdoArticulos, sSQL, "Producto"
End Sub

Public Sub Generar_Recibo_PDF(TFA As Tipo_Facturas, _
                              VerFactura As Boolean)
Dim AdoDBFac As ADODB.Recordset
Dim AdoDBDet As ADODB.Recordset
Dim AdoDBAux As ADODB.Recordset
Dim ConsultarDetalle As Boolean
Dim Imagen As IPictureDisp
Dim PathCodigoBarra As String
Dim TempPosLinea As Single
Dim Comprobante As String

    RatonReloj
    Comprobante = "Recibo No " & FA.Serie & "-" & Format$(FA.Factura, "000000000")
    ConsultarDetalle = False
    sSQL = "SELECT F.*,C.Cliente,C.CI_RUC,C.Telefono,C.TelefonoT,C.Direccion,C.DireccionT," _
         & "C.Representante,C.Grupo,C.Codigo,C.Ciudad,C.Email,C.Email2,C.EmailR,C.CI_RUC_R,C.TD,C.TD_R,C.DirNumero " _
         & "FROM Facturas As F,Clientes As C " _
         & "WHERE F.Item = '" & NumEmpresa & "' " _
         & "AND F.Periodo = '" & Periodo_Contable & "' " _
         & "AND F.TC = '" & TFA.TC & "' " _
         & "AND F.Serie = '" & TFA.Serie & "' " _
         & "AND F.Autorizacion = '" & TFA.Autorizacion & "' " _
         & "AND F.CodigoC = '" & TFA.CodigoC & "' " _
         & "AND F.Factura = " & TFA.Factura & " " _
         & "AND C.Codigo = F.CodigoC "
    Select_AdoDB AdoDBFac, sSQL
    With AdoDBFac
     If .RecordCount > 0 Then
         TFA.Grupo = .fields("Grupo")
         TFA.Autorizacion = .fields("Autorizacion")
         TFA.CodigoC = .fields("CodigoC")
         TFA.DireccionC = .fields("Direccion")
         TFA.Cliente = .fields("Cliente")
         If Len(.fields("RUC_CI")) > 1 Then TFA.CI_RUC = .fields("RUC_CI") Else TFA.CI_RUC = .fields("RUC_CI")
         If Len(.fields("Razon_Social")) > 1 And Len(.fields("RUC_CI")) > 1 Then
            sSQL = "SELECT Codigo,Grupo_No,Representante,Cedula_R,Lugar_Trabajo_R,Telefono_R,Email_R " _
                 & "FROM Clientes_Matriculas " _
                 & "WHERE Item = '" & NumEmpresa & "' " _
                 & "AND Periodo = '" & Periodo_Contable & "' " _
                 & "AND Codigo = '" & TFA.CodigoC & "' "
            Select_AdoDB AdoDBDet, sSQL
            If AdoDBDet.RecordCount > 0 Then
               TFA.Curso = TFA.DireccionC
               TFA.EmailR = AdoDBDet.fields("Email_R")
               TFA.DireccionC = AdoDBDet.fields("Lugar_Trabajo_R")
               TFA.Comercial = AdoDBDet.fields("Representante")
            End If
            AdoDBDet.Close
         End If
         ConsultarDetalle = True
     End If
    End With
    
    SaldoPendiente = 0
    sSQL = "SELECT CodigoC, SUM(Saldo_MN) As Pendiente " _
         & "FROM Facturas " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND CodigoC = '" & TFA.CodigoC & "' " _
         & "AND TC = '" & TFA.TC & "' " _
         & "AND Saldo_MN > 0 " _
         & "AND T <> 'A' " _
         & "GROUP BY CodigoC "
    Select_AdoDB AdoDBAux, sSQL
   'If AdoDBAux.RecordCount > 0 Then SaldoPendiente = AdoDBAux.Fields("Pendiente")
    AdoDBAux.Close
        
    sSQL = "SELECT * " _
         & "FROM Detalle_Factura " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND TC = '" & TFA.TC & "' " _
         & "AND Serie = '" & TFA.Serie & "' " _
         & "AND Autorizacion = '" & TFA.Autorizacion & "' " _
         & "AND Factura = " & TFA.Factura & " " _
         & "ORDER BY ID,Codigo "
    Select_AdoDB AdoDBDet, sSQL
    
   'Geneeramos el documento
    tPrint.TipoImpresion = Es_PDF
    tPrint.NombreArchivo = Comprobante
    tPrint.TituloArchivo = "Recibo Electronico"
    tPrint.TipoLetra = TipoCourier
    tPrint.OrientacionPagina = Orientacion_Pagina
    tPrint.PaginaA4 = True
    tPrint.EsCampoCorto = False
    tPrint.VerDocumento = VerFactura
    
    Set cPrint = New cImpresion
    cPrint.iniciaImpresion
   'Encabezado Factura
    With AdoDBFac
     If .RecordCount > 0 Then
         cPrint.printImagen LogoTipo, 1, 1, 5, 2.5
         cPrint.colorDeLetra = Negro
         cPrint.tipoDeLetra = TipoCourier
         cPrint.PorteDeLetra = 7
         PosLinea = 4.1
         cPrint.printTexto 1.2, PosLinea, RazonSocial
         cPrint.colorDeLetra = Negro
         PosLinea = PosLinea + 0.35
         cPrint.printTexto 1.2, PosLinea, UCaseStrg(NombreComercial)
         PosLinea = PosLinea + 0.35
         cPrint.printTexto 1.2, PosLinea, "Dirección:"
         PosLinea = PosLinea + 0.35
         cPrint.printTexto 1.2, PosLinea, Direccion
         PosLinea = PosLinea + 0.5
         TempPosLinea = PosLinea
         PosLinea = 1.1
         cPrint.printTexto 11.2, PosLinea, "R.U.C.: " & RUC
         PosLinea = PosLinea + 0.35
         cPrint.printTexto 11.2, PosLinea, "C O M P R O B A N T E   D E   P A G O"
         PosLinea = PosLinea + 0.35
         cPrint.printTexto 11.2, PosLinea, "No." & TFA.Serie & "-" & Format$(TFA.Factura, "000000000")
         PosLinea = PosLinea + 0.4
         cPrint.tipoDeLetra = TipoCourier
         cPrint.PorteDeLetra = 9
         cPrint.printTexto 11.2, PosLinea, "C O D I G O: " & .fields("CI_RUC")
         cPrint.tipoDeLetra = TipoCourier
         cPrint.PorteDeLetra = 7
         PosLinea = PosLinea + 0.4
         cPrint.printTexto 11.2, PosLinea, "FECHA Y HORA DE EMISION"
         PosLinea = PosLinea + 0.35
         cPrint.printTexto 11.2, PosLinea, .fields("Fecha") & " - " & .fields("Hora")
         PosLinea = PosLinea + 0.35
         cPrint.printTexto 11.2, PosLinea, "ESTUDIANTE:"
         PosLinea = PosLinea + 0.35
         cPrint.printTexto 11.2, PosLinea, TFA.Cliente
         PosLinea = PosLinea + 0.35
         cPrint.printTexto 11.2, PosLinea, "Representante:"
         PosLinea = PosLinea + 0.35
         cPrint.printTexto 11.2, PosLinea, TFA.Comercial
         PosLinea = PosLinea + 0.35
         cPrint.printTexto 11.2, PosLinea, "Identificación: " & TFA.CI_RUC
         PosLinea = PosLinea + 0.35
         cPrint.printTexto 11.2, PosLinea, "Fecha Emisión: " & .fields("Fecha")
         PosLinea = TempPosLinea + 0.2
         cPrint.printTexto 1.1, PosLinea, "Codigo"
         cPrint.printTexto 4, PosLinea, "D e s c r i p c i ó n"
         cPrint.printTexto 14, PosLinea, "Cantidad"
         cPrint.printTexto 16, PosLinea, "Precio"
         cPrint.printTexto 18, PosLinea, "Precio"
         PosLinea = PosLinea + 0.3
         cPrint.printTexto 16, PosLinea, "Unitario"
         cPrint.printTexto 18, PosLinea, "Total"
         
        'Cuadors Superiores
         cPrint.printCuadro 1, 4, 10, 5.7, Negro, "B"
         cPrint.printCuadro 11, 1, 19.6, 5.7, Negro, "B"
         cPrint.printCuadro 1, 6, 19.6, 6.7, Negro, "B"
     End If
    End With
    cPrint.colorDeLetra = Negro
   'Detalle Factura
    If ConsultarDetalle Then
       PosLinea = 6.6
       With AdoDBDet
        If .RecordCount > 0 Then
            Do While Not .EOF
               Producto = .fields("Producto")
               If .fields("Ticket") <> Ninguno Then Producto = Producto & " " & .fields("Ticket")
               If .fields("Mes") <> Ninguno Then Producto = Producto & " " & .fields("Mes")
               cPrint.printTexto 1.1, PosLinea, .fields("Codigo")
               cPrint.printTexto 4, PosLinea, Producto, , 9
               If Lineas_Impresas > 1 Then PosLinea = PosLinea - 0.35
               cPrint.printVariable 13, PosLinea, .fields("Cantidad")
               cPrint.printVariable 15.4, PosLinea, .fields("Precio"), True, 2
               cPrint.printVariable 17.8, PosLinea, .fields("Total"), True, 2
               PosLinea = PosLinea + 0.35
              .MoveNext
            Loop
        End If
       End With
    End If
    TempPosLinea = PosLinea
   'Pie de la Factura
    With AdoDBFac
     If .RecordCount > 0 Then
         If SaldoPendiente <= 0 Then SaldoPendiente = .fields("Total_MN")
         Diferencia = SaldoPendiente - .fields("Total_MN")
         If Diferencia < 0 Then Diferencia = 0
         
         PosLinea = TempPosLinea + 0.3
         cPrint.printTexto 1.2, PosLinea, "I n f o r m a c i ó n   A d i c i o n a l"
         PosLinea = PosLinea + 0.5
         If TFA.Grupo <> Ninguno And TFA.Curso <> Ninguno Then
            cPrint.printTexto 1.2, PosLinea, "Curso: " & TFA.Grupo & "-" & TFA.Curso
            PosLinea = PosLinea + 0.35
         End If
         If Len(TFA.DireccionC) > 1 Then
            cPrint.printTexto 1.2, PosLinea, "Dirección: " & TFA.DireccionC
            PosLinea = PosLinea + 0.35
         End If
         If Len(.fields("Telefono")) > 1 Then
            cPrint.printTexto 1.2, PosLinea, "Teléfono: " & .fields("Telefono")
            PosLinea = PosLinea + 0.35
         End If
         If Len(.fields("Email")) > 1 Then
            cPrint.printTexto 1.2, PosLinea, "Email: " & .fields("Email")
            PosLinea = PosLinea + 0.35
         End If
         If Len(.fields("Observacion")) > 1 Then
            cPrint.printTexto 1.2, PosLinea, "Observacion: " & .fields("Observacion")
         End If
         PosLinea = TempPosLinea + 0.1
         'PosLinea = 25.3
         cPrint.printTexto 13.6, PosLinea, "SUBTOTAL 0%"
         cPrint.printVariable 17.8, PosLinea, .fields("Sin_IVA"), True, 2
         PosLinea = PosLinea + 0.4
         cPrint.printTexto 13.6, PosLinea, "SUBTOTAL " & Porc_IVA * 100 & "%"
         cPrint.printVariable 17.8, PosLinea, .fields("Con_IVA"), True, 2
         PosLinea = PosLinea + 0.4
         cPrint.printTexto 13.6, PosLinea, "TOTAL DESCUENTO"
         cPrint.printVariable 17.8, PosLinea, .fields("Descuento") + .fields("Descuento2"), True, 2
         PosLinea = PosLinea + 0.4
         cPrint.printTexto 13.6, PosLinea, "SUBTOTAL SIN IMPUESTOS"
         cPrint.printVariable 17.8, PosLinea, .fields("SubTotal"), True, 2
         PosLinea = PosLinea + 0.4
         cPrint.printTexto 13.6, PosLinea, "TOTAL I.V.A. " & Porc_IVA * 100 & "%"
         cPrint.printVariable 17.8, PosLinea, .fields("IVA"), True, 2
         PosLinea = PosLinea + 0.4
         cPrint.printTexto 13.6, PosLinea, "TOTAL FACTURADO"
         cPrint.printVariable 17.8, PosLinea, .fields("Total_MN"), True, 2
         PosLinea = PosLinea + 0.4
         cPrint.printTexto 13.6, PosLinea, "DEUDA PENDITNE"
         cPrint.printVariable 17.8, PosLinea, Diferencia, True, 2
         PosLinea = PosLinea + 0.4
         cPrint.printTexto 13.6, PosLinea, "TOTAL A PAGAR"
         cPrint.printVariable 17.8, PosLinea, Diferencia + .fields("Total_MN"), True, 2
         
         PosLinea = TempPosLinea + 0.1
         cPrint.printCuadro 1, 6, 19.6, PosLinea + 0.1, Negro, "B"
         PosLinea = PosLinea + 0.4
         cPrint.printLinea 13.5, PosLinea, 20, PosLinea, Negro
         PosLinea = PosLinea + 0.4
         cPrint.printLinea 13.5, PosLinea, 20, PosLinea, Negro
         PosLinea = PosLinea + 0.4
         cPrint.printLinea 13.5, PosLinea, 20, PosLinea, Negro
         PosLinea = PosLinea + 0.4
         cPrint.printLinea 13.5, PosLinea, 20, PosLinea, Negro
         PosLinea = PosLinea + 0.4
         cPrint.printLinea 13.5, PosLinea, 20, PosLinea, Negro
         PosLinea = PosLinea + 0.4
         cPrint.printLinea 13.5, PosLinea, 20, PosLinea, Negro
         PosLinea = PosLinea + 0.4
         cPrint.printLinea 13.5, PosLinea, 20, PosLinea, Negro
         PosLinea = PosLinea + 0.4
         cPrint.printLinea 13.5, PosLinea, 20, PosLinea, Negro
         
         PosLinea = TempPosLinea
         cPrint.printCuadro 1, PosLinea + 0.4, 12.9, PosLinea + 3.4, Negro, "B"
         cPrint.printLinea 1, PosLinea + 0.8, 13.3, PosLinea + 0.8, Negro
         
         'cprint.printcuadrolinea  1, 5.7, 1, PosLinea, Negro
         cPrint.printLinea 3.8, 5.85, 3.8, PosLinea + 0.1, Negro
         cPrint.printLinea 13.4, 5.85, 13.4, PosLinea + 3.3, Negro
         cPrint.printLinea 15.4, 5.85, 15.4, PosLinea + 0.1, Negro
         cPrint.printLinea 17.7, 5.85, 17.7, PosLinea + 3.3, Negro
         cPrint.printLinea 20, 5.85, 20, PosLinea + 3.3, Negro
     End If
    End With
    cPrint.finalizaImpresion
    AdoDBDet.Close
    AdoDBFac.Close
    RatonNormal
End Sub

'''Public Sub Actualizar_Abonos_Facturas(TFA As Tipo_Facturas, _
'''                                      Optional SaldoReal As Boolean, _
'''                                      Optional PorFecha As Boolean)
'''Dim AdoFechas As ADODB.Recordset
'''Dim AdoAux As ADODB.Recordset
'''Dim IdAnio As Integer
'''Dim IDMes As Integer
'''Dim FechaI As Long
'''Dim FechaF As Long
'''Dim FechaA As Long
'''Dim FechaDelMes As String
'''Dim EsFacturaIndividual As String
'''Dim EsFacturaIndividualF As String
'''Dim EsAbonoIndividual As String
'''   'MsgBox SaldoReal & vbCrLf & PorFecha
'''    Progreso_Barra.Mensaje_Box = "Determinando fechas de proceso"
'''    Progreso_Iniciar
'''    Progreso_Barra.Valor_Maximo = 100
'''
'''    If TFA.Factura > 0 And Len(TFA.TC) = 2 And Len(TFA.Serie) = 6 Then
'''       EsFacturaIndividual = "AND TC = '" & TFA.TC & "' " _
'''                           & "AND Serie = '" & TFA.Serie & "' " _
'''                           & "AND Factura = " & TFA.Factura & " "
'''       EsFacturaIndividualF = "AND F.TC = '" & TFA.TC & "' " _
'''                            & "AND F.Serie = '" & TFA.Serie & "' " _
'''                            & "AND F.Factura = " & TFA.Factura & " "
'''       EsAbonoIndividual = "AND TP = '" & TFA.TC & "' " _
'''                         & "AND Serie = '" & TFA.Serie & "' " _
'''                         & "AND Factura = " & TFA.Factura & " "
'''    Else
'''       EsFacturaIndividual = ""
'''       EsFacturaIndividualF = ""
'''    End If
'''
'''    If Not IsDate(TFA.Fecha_Corte) Then TFA.Fecha_Corte = FechaSistema
'''    If TFA.Fecha_Corte = FechaSistema Then SaldoReal = True
'''    FechaI = CFechaLong(TFA.Fecha_Corte)
'''    FechaF = CFechaLong(TFA.Fecha_Corte)
'''
'''   'MsgBox EsEsAbonoIndividual & vbCrLf & "-------------------------------------------" & vbCrLf & EsEsAbonoIndividualf
'''    sSQL = "SELECT T, MIN(Fecha) As Fecha_Min " _
'''         & "FROM Facturas " _
'''         & "WHERE Item = '" & NumEmpresa & "' " _
'''         & "AND Periodo = '" & Periodo_Contable & "' " _
'''         & "AND Fecha <= #" & BuscarFecha(TFA.Fecha_Corte) & "# " _
'''         & EsFacturaIndividual _
'''         & "AND T <> '" & Anulado & "' " _
'''         & "GROUP BY T "
'''    Select_AdoDB AdoFechas, sSQL
'''    If AdoFechas.RecordCount > 0 Then
'''       If FechaI > CFechaLong(AdoFechas.Fields("Fecha_Min")) Then FechaI = CFechaLong(AdoFechas.Fields("Fecha_Min"))
'''    End If
'''    AdoFechas.Close
'''
'''    Progreso_Barra.Mensaje_Box = "Fechas de Abonos"
'''    Progreso_Esperar
'''    sSQL = "SELECT T, MIN(Fecha) As Fecha_Min " _
'''         & "FROM Trans_Abonos " _
'''         & "WHERE Item = '" & NumEmpresa & "' " _
'''         & "AND Periodo = '" & Periodo_Contable & "' " _
'''         & "AND T <> '" & Anulado & "' " _
'''         & "AND Fecha <= #" & BuscarFecha(TFA.Fecha_Corte) & "# " _
'''         & EsAbonoIndividual _
'''         & "GROUP BY T "
'''    Select_AdoDB AdoFechas, sSQL
'''    If AdoFechas.RecordCount > 0 Then
'''       If FechaI > CFechaLong(AdoFechas.Fields("Fecha_Min")) Then FechaI = CFechaLong(AdoFechas.Fields("Fecha_Min"))
'''    End If
'''    AdoFechas.Close
'''    FechaI = CFechaLong(PrimerDiaMes(CLongFecha(FechaI)))
'''    If FechaI > FechaF Then FechaF = FechaI
'''    If PorFecha And CFechaLong(TFA.Fecha_Desde) <= CFechaLong(TFA.Fecha_Hasta) Then
'''       FechaI = CFechaLong(TFA.Fecha_Desde)
'''       FechaF = CFechaLong(TFA.Fecha_Hasta)
'''    Else
'''       FechaF = CFechaLong(TFA.Fecha_Hasta)
'''    End If
'''    FechaIni = BuscarFecha(CLongFecha(FechaI))
'''    FechaFin = BuscarFecha(CLongFecha(FechaF))
'''
'''   'MsgBox "Fecha Inicial: " & FechaIni
'''    sSQL = "UPDATE Clientes " _
'''         & "SET X = '.' " _
'''         & "WHERE Codigo <> '.' "
'''    Ejecutar_SQL_SP sSQL
'''
'''    Progreso_Barra.Mensaje_Box = "Actualizando Ejecutivos de Venta"
'''    Progreso_Esperar
'''    If SQL_Server Then
'''       sSQL = "UPDATE Clientes " _
'''            & "SET X = 'C' " _
'''            & "FROM Clientes As C, Facturas As F "
'''    Else
'''       sSQL = "UPDATE Clientes As C, Facturas As F " _
'''            & "SET C.X = 'C' "
'''    End If
'''    sSQL = sSQL _
'''         & "WHERE F.Item = '" & NumEmpresa & "' " _
'''         & "AND F.Periodo = '" & Periodo_Contable & "' " _
'''         & EsFacturaIndividualF _
'''         & "AND C.Codigo = F.Cod_Ejec "
'''    Ejecutar_SQL_SP sSQL
'''
'''    If SQL_Server Then
'''       sSQL = "UPDATE Clientes " _
'''            & "SET X = 'A' " _
'''            & "FROM Clientes As C, Accesos As A "
'''    Else
'''       sSQL = "UPDATE Clientes As C, Accesos As A " _
'''            & "SET C.X = 'A' "
'''    End If
'''    sSQL = sSQL _
'''         & "WHERE C.Codigo = A.Codigo "
'''    Ejecutar_SQL_SP sSQL
'''
'''    Cadena = ""
'''    sSQL = "SELECT Codigo,Cliente " _
'''         & "FROM Clientes " _
'''         & "WHERE X = 'C' " _
'''         & "ORDER BY T "
'''    Select_AdoDB AdoAux, sSQL
'''    If AdoAux.RecordCount > 0 Then
'''       Cadena = Cadena & "ASIGNE ESTAS PERSONAS A NOMINA:" & vbCrLf
'''       Do While Not AdoAux.EOF
'''          Codigo = AdoAux.Fields("Codigo")
'''          NombreCliente = ULCase(AdoAux.Fields("Cliente"))
'''          Cadena = Cadena & NombreCliente & vbCrLf
'''          SetAdoAddNew "Accesos", True
'''          SetAdoFields "Codigo", Codigo
'''          SetAdoFields "Nombre_Completo", NombreCliente
'''          SetAdoUpdate
'''          AdoAux.MoveNext
'''       Loop
'''    End If
'''    AdoAux.Close
'''    If Len(Cadena) > 0 Then MsgBox Cadena
'''    sSQL = "DELETE * " _
'''         & "FROM Asiento_Abonos " _
'''         & "WHERE Item = '" & NumEmpresa & "' " _
'''         & "AND Periodo = '" & Periodo_Contable & "' " _
'''         & "AND CodigoU = '" & CodigoUsuario & "' "
'''    Ejecutar_SQL_SP sSQL
'''
'''   'Enceramos los SubTotales
'''    Progreso_Barra.Mensaje_Box = "Encerando Abonos"
'''    Progreso_Esperar
'''    sSQL = "UPDATE Facturas " _
'''         & "SET Total_Abonos = 0," _
'''         & "Total_Efectivo = 0," _
'''         & "Total_Banco = 0," _
'''         & "Otros_Abonos = 0," _
'''         & "Total_Ret_Fuente = 0," _
'''         & "Total_Ret_IVA_B = 0," _
'''         & "Total_Ret_IVA_S = 0, " _
'''         & "Saldo_Actual = Total_MN, " _
'''         & "Fecha_C = Fecha, " _
'''         & "Fecha_R = Fecha " _
'''         & "WHERE Item = '" & NumEmpresa & "' " _
'''         & "AND Periodo = '" & Periodo_Contable & "' " _
'''         & EsFacturaIndividual
'''    Ejecutar_SQL_SP sSQL
'''
'''   'MsgBox CLongFecha(FechaI) & " <--> " & CLongFecha(FechaF)
'''    Progreso_Barra.Mensaje_Box = "Actualizando Fechas de retenciones"
'''    Progreso_Esperar
'''
'''   'Actualizamos fecha de retencion fuente
'''    If SQL_Server Then
'''       sSQL = "UPDATE Facturas " _
'''            & "SET Fecha_R = TA.Fecha " _
'''            & "FROM Facturas As F,Trans_Abonos As TA "
'''    Else
'''       sSQL = "UPDATE Facturas As F,Trans_Abonos As TA " _
'''            & "SET F.Fecha_R = TA.Fecha "
'''    End If
'''    sSQL = sSQL _
'''         & "WHERE F.Item = '" & NumEmpresa & "' " _
'''         & "AND F.Periodo = '" & Periodo_Contable & "' " _
'''         & "AND F.T <> 'A' " _
'''         & "AND TA.Tipo_Cta = 'CF' " _
'''         & EsFacturaIndividualF _
'''         & "AND Fecha_R <> TA.Fecha " _
'''         & "AND F.Item = TA.Item " _
'''         & "AND F.Periodo = TA.Periodo " _
'''         & "AND F.TC = TA.TP " _
'''         & "AND F.Serie = TA.Serie " _
'''         & "AND F.Factura = TA.Factura " _
'''         & "AND F.Autorizacion = TA.Autorizacion " _
'''         & "AND F.CodigoC = TA.CodigoC "
'''    Ejecutar_SQL_SP sSQL
'''
'''   'Actualizamos totales de abonos en tipos de ctas
'''    sSQL = "INSERT INTO Asiento_Abonos (Periodo,Item,CodigoC,Fecha,TP,Serie,Factura,Autorizacion,Tipo_Cta,Abono,CodigoU) " _
'''         & "SELECT Periodo,Item,CodigoC,MAX(Fecha) As FechaMax,TP,Serie,Factura,Autorizacion,Tipo_Cta,ROUND(SUM(Abono),2,0),'" & CodigoUsuario & "' " _
'''         & "FROM Trans_Abonos " _
'''         & "WHERE Item = '" & NumEmpresa & "' " _
'''         & "AND Periodo = '" & Periodo_Contable & "' " _
'''         & "AND Fecha <= #" & FechaFin & "# " _
'''         & "AND Tipo_Cta IN ('CB','CI','CF','CJ','BA') " _
'''         & "AND T <> 'A' " _
'''         & EsAbonoIndividual _
'''         & "GROUP BY Periodo,Item,CodigoC,TP,Serie,Factura,Autorizacion,Tipo_Cta "
'''    Ejecutar_SQL_SP sSQL
'''
'''   'Actualizamos totales de Otros Abonos en tipos de ctas
'''    sSQL = "INSERT INTO Asiento_Abonos (Periodo,Item,CodigoC,Fecha,TP,Serie,Factura,Autorizacion,Tipo_Cta,Abono,CodigoU) " _
'''         & "SELECT Periodo,Item,CodigoC,MAX(Fecha) As FechaMax,TP,Serie,Factura,Autorizacion,'OA' As TipoCta,ROUND(SUM(Abono),2,0),'" & CodigoUsuario & "' " _
'''         & "FROM Trans_Abonos " _
'''         & "WHERE Item = '" & NumEmpresa & "' " _
'''         & "AND Periodo = '" & Periodo_Contable & "' " _
'''         & "AND Fecha <= #" & FechaFin & "# " _
'''         & "AND NOT Tipo_Cta IN ('CB','CI','CF','CJ','BA') " _
'''         & "AND T <> 'A' " _
'''         & EsAbonoIndividual _
'''         & "GROUP BY Periodo,Item,CodigoC,TP,Serie,Factura,Autorizacion,Tipo_Cta "
'''    Ejecutar_SQL_SP sSQL
'''
'''   'MsgBox "Insertando Totales de Abonos de " & FechaIni & " al " & FechaFin
'''
'''    Progreso_Barra.Mensaje_Box = "Actualizando Retenciones IVA Bienes"
'''    Progreso_Esperar
'''   'Retenciones del IVA en Bienes
'''    If SQL_Server Then
'''       sSQL = "UPDATE Facturas " _
'''            & "SET Total_Ret_IVA_B = AA.Abono " _
'''            & "FROM Facturas As F,Asiento_Abonos As AA "
'''    Else
'''       sSQL = "UPDATE Facturas As F,Asiento_Abonos As AA " _
'''            & "SET F.Total_Ret_IVA_B = AA.Abono "
'''    End If
'''    sSQL = sSQL _
'''         & "WHERE F.Item = '" & NumEmpresa & "' " _
'''         & "AND F.Periodo = '" & Periodo_Contable & "' " _
'''         & "AND F.T <> 'A' " _
'''         & "AND AA.CodigoU = '" & CodigoUsuario & "' " _
'''         & "AND AA.Tipo_Cta = 'CB' " _
'''         & EsFacturaIndividualF _
'''         & "AND F.Item = AA.Item " _
'''         & "AND F.Periodo = AA.Periodo " _
'''         & "AND F.TC = AA.TP " _
'''         & "AND F.Serie = AA.Serie " _
'''         & "AND F.Factura = AA.Factura " _
'''         & "AND F.Autorizacion = AA.Autorizacion " _
'''         & "AND F.CodigoC = AA.CodigoC "
'''    Ejecutar_SQL_SP sSQL
'''
'''    Progreso_Barra.Mensaje_Box = "Actualizando Retenciones IVA Servicios"
'''    Progreso_Esperar
'''
'''   'Retenciones del IVA en Servicios
'''    If SQL_Server Then
'''       sSQL = "UPDATE Facturas " _
'''            & "SET Total_Ret_IVA_S = AA.Abono " _
'''            & "FROM Facturas As F,Asiento_Abonos As AA "
'''    Else
'''       sSQL = "UPDATE Facturas As F,Asiento_Abonos As AA " _
'''            & "SET F.Total_Ret_IVA_S = AA.Abono "
'''    End If
'''    sSQL = sSQL _
'''         & "WHERE F.Item = '" & NumEmpresa & "' " _
'''         & "AND F.Periodo = '" & Periodo_Contable & "' " _
'''         & "AND F.T <> 'A' " _
'''         & "AND AA.CodigoU = '" & CodigoUsuario & "' " _
'''         & "AND AA.Tipo_Cta = 'CI' " _
'''         & EsFacturaIndividualF _
'''         & "AND F.Item = AA.Item " _
'''         & "AND F.Periodo = AA.Periodo " _
'''         & "AND F.TC = AA.TP " _
'''         & "AND F.Serie = AA.Serie " _
'''         & "AND F.Factura = AA.Factura " _
'''         & "AND F.Autorizacion = AA.Autorizacion " _
'''         & "AND F.CodigoC = AA.CodigoC "
'''    Ejecutar_SQL_SP sSQL
'''
'''    Progreso_Barra.Mensaje_Box = "Actualizando Retenciones en la Fuente"
'''    Progreso_Esperar
'''   'Retenciones de la Fuente
'''    If SQL_Server Then
'''       sSQL = "UPDATE Facturas " _
'''            & "SET Total_Ret_Fuente = AA.Abono " _
'''            & "FROM Facturas As F,Asiento_Abonos As AA "
'''    Else
'''       sSQL = "UPDATE Facturas As F,Asiento_Abonos As AA " _
'''            & "SET F.Total_Ret_Fuente = AA.Abono "
'''    End If
'''    sSQL = sSQL _
'''         & "WHERE F.Item = '" & NumEmpresa & "' " _
'''         & "AND F.Periodo = '" & Periodo_Contable & "' " _
'''         & "AND F.T <> 'A' " _
'''         & "AND AA.CodigoU = '" & CodigoUsuario & "' " _
'''         & "AND AA.Tipo_Cta = 'CF' " _
'''         & EsFacturaIndividualF _
'''         & "AND F.Item = AA.Item " _
'''         & "AND F.Periodo = AA.Periodo " _
'''         & "AND F.TC = AA.TP " _
'''         & "AND F.Serie = AA.Serie " _
'''         & "AND F.Factura = AA.Factura " _
'''         & "AND F.Autorizacion = AA.Autorizacion " _
'''         & "AND F.CodigoC = AA.CodigoC "
'''    Ejecutar_SQL_SP sSQL
'''
'''    Progreso_Barra.Mensaje_Box = "Actualizando Totales de Caja Efectivo"
'''    Progreso_Esperar
'''
'''   'MsgBox CLongFecha(FechaI) & " - " & CLongFecha(FechaF)
'''
''''''    For FechaA = FechaI To FechaF Step 31
''''''        FechaDelMes = "#" & BuscarFecha(PrimerDiaMes(CLongFecha(FechaA))) & "# AND #" & BuscarFecha(UltimoDiaMes(CLongFecha(FechaA))) & "#"
'''
'''       'MsgBox BuscarFecha(CLongFecha(J)) & " - " & BuscarFecha(CLongFecha(I))
'''       '& "AND Fecha BETWEEN " & FechaDelMes & " "
'''        sSQL = "UPDATE Facturas " _
'''             & "SET Total_Efectivo= (SELECT SUM(AA.Abono) " _
'''             & "                    FROM Asiento_Abonos As AA " _
'''             & "                    WHERE AA.Item = '" & NumEmpresa & "' " _
'''             & "                    AND AA.Periodo = '" & Periodo_Contable & "' " _
'''             & "                    AND AA.CodigoU = '" & CodigoUsuario & "' " _
'''             & "                    AND AA.Tipo_Cta = 'CJ' " _
'''             & "                    AND AA.TP = Facturas.TC " _
'''             & "                    AND AA.Item = Facturas.Item " _
'''             & "                    AND AA.Periodo = Facturas.Periodo " _
'''             & "                    AND AA.Factura = Facturas.Factura " _
'''             & "                    AND AA.Serie = Facturas.Serie " _
'''             & "                    AND AA.Autorizacion = Facturas.Autorizacion) " _
'''             & "WHERE Item = '" & NumEmpresa & "' " _
'''             & "AND Periodo = '" & Periodo_Contable & "' " _
'''             & "AND Fecha <= #" & FechaFin & "# " _
'''             & "AND T <> 'A' " _
'''             & EsFacturaIndividual
'''        Ejecutar_SQL_SP sSQL
'''
'''        Progreso_Barra.Mensaje_Box = "Actualizando Totales de Bancos"
'''        Progreso_Esperar
'''        sSQL = "UPDATE Facturas " _
'''             & "SET Total_Banco = (SELECT SUM(AA.Abono) " _
'''             & "                    FROM Asiento_Abonos As AA " _
'''             & "                    WHERE AA.Item = '" & NumEmpresa & "' " _
'''             & "                    AND AA.Periodo = '" & Periodo_Contable & "' " _
'''             & "                    AND AA.CodigoU = '" & CodigoUsuario & "' " _
'''             & "                    AND AA.Tipo_Cta = 'BA' " _
'''             & "                    AND AA.TP = Facturas.TC " _
'''             & "                    AND AA.Item = Facturas.Item " _
'''             & "                    AND AA.Periodo = Facturas.Periodo " _
'''             & "                    AND AA.Factura = Facturas.Factura " _
'''             & "                    AND AA.Serie = Facturas.Serie " _
'''             & "                    AND AA.Autorizacion = Facturas.Autorizacion) " _
'''             & "WHERE Item = '" & NumEmpresa & "' " _
'''             & "AND Periodo = '" & Periodo_Contable & "' " _
'''             & "AND Fecha <= #" & FechaFin & "# " _
'''             & "AND T <> 'A' " _
'''             & EsFacturaIndividual
'''        Ejecutar_SQL_SP sSQL
'''
'''        Progreso_Barra.Mensaje_Box = "Actualizando Totales de otros abonos"
'''        Progreso_Esperar
'''        sSQL = "UPDATE Facturas " _
'''             & "SET Otros_Abonos = (SELECT SUM(AA.Abono) " _
'''             & "                    FROM Asiento_Abonos As AA " _
'''             & "                    WHERE AA.Item = '" & NumEmpresa & "' " _
'''             & "                    AND AA.Periodo = '" & Periodo_Contable & "' " _
'''             & "                    AND AA.CodigoU = '" & CodigoUsuario & "' " _
'''             & "                    AND AA.Tipo_Cta = 'OA' " _
'''             & "                    AND AA.TP = Facturas.TC " _
'''             & "                    AND AA.Item = Facturas.Item " _
'''             & "                    AND AA.Periodo = Facturas.Periodo " _
'''             & "                    AND AA.Factura = Facturas.Factura " _
'''             & "                    AND AA.Serie = Facturas.Serie " _
'''             & "                    AND AA.Autorizacion = Facturas.Autorizacion) " _
'''             & "WHERE Item = '" & NumEmpresa & "' " _
'''             & "AND Periodo = '" & Periodo_Contable & "' " _
'''             & "AND Fecha <= #" & FechaFin & "# " _
'''             & "AND T <> 'A' " _
'''             & EsFacturaIndividual
'''        Ejecutar_SQL_SP sSQL
'''
'''        Progreso_Barra.Mensaje_Box = "Actualizando Fecha de Cancelacion de Facturas"
'''        Progreso_Esperar
'''        sSQL = "UPDATE Facturas " _
'''             & "SET Fecha_C = (SELECT MAX(Fecha) " _
'''             & "               FROM Asiento_Abonos As AA " _
'''             & "               WHERE AA.Item = '" & NumEmpresa & "' " _
'''             & "               AND AA.Periodo = '" & Periodo_Contable & "' " _
'''             & "               AND AA.CodigoU = '" & CodigoUsuario & "' " _
'''             & "               AND AA.TP = Facturas.TC " _
'''             & "               AND AA.Item = Facturas.Item " _
'''             & "               AND AA.Periodo = Facturas.Periodo " _
'''             & "               AND AA.Factura = Facturas.Factura " _
'''             & "               AND AA.Serie = Facturas.Serie " _
'''             & "               AND AA.Autorizacion = Facturas.Autorizacion) " _
'''             & "WHERE Item = '" & NumEmpresa & "' " _
'''             & "AND Periodo = '" & Periodo_Contable & "' " _
'''             & "AND Fecha <= #" & FechaFin & "# " _
'''             & "AND T <> 'A' " _
'''             & EsFacturaIndividual
'''        Ejecutar_SQL_SP sSQL
''''''    Next FechaA
'''
'''    sSQL = "UPDATE Facturas " _
'''         & "SET Total_Efectivo = 0 " _
'''         & "WHERE Total_Efectivo IS NULL " _
'''         & "AND Item = '" & NumEmpresa & "' " _
'''         & "AND Periodo = '" & Periodo_Contable & "' "
'''    Ejecutar_SQL_SP sSQL
'''
'''    sSQL = "UPDATE Facturas " _
'''         & "SET Total_Banco = 0 " _
'''         & "WHERE Total_Banco IS NULL " _
'''         & "AND Item = '" & NumEmpresa & "' " _
'''         & "AND Periodo = '" & Periodo_Contable & "' "
'''    Ejecutar_SQL_SP sSQL
'''
'''    sSQL = "UPDATE Facturas " _
'''         & "SET Otros_Abonos = 0 " _
'''         & "WHERE Otros_Abonos IS NULL " _
'''         & "AND Item = '" & NumEmpresa & "' " _
'''         & "AND Periodo = '" & Periodo_Contable & "' "
'''    Ejecutar_SQL_SP sSQL
'''
'''    sSQL = "DELETE * " _
'''         & "FROM Asiento_Abonos " _
'''         & "WHERE Item = '" & NumEmpresa & "' " _
'''         & "AND Periodo = '" & Periodo_Contable & "' " _
'''         & "AND CodigoU = '" & CodigoUsuario & "' "
'''''    Ejecutar_SQL_SP sSQL
'''
'''   'Enceramos todos los nulos
'''    Progreso_Barra.Mensaje_Box = "Determinando Nulos"
'''    Progreso_Esperar
'''
'''    sSQL = "UPDATE Facturas " _
'''         & "SET Fecha_C = Fecha " _
'''         & "WHERE Fecha_C IS NULL " _
'''         & "AND Item = '" & NumEmpresa & "' " _
'''         & "AND Periodo = '" & Periodo_Contable & "' "
'''    Ejecutar_SQL_SP sSQL
'''
'''    Progreso_Barra.Mensaje_Box = "Totalizando Saldos de Factura"
'''    Progreso_Esperar
''''''    For FechaA = FechaI To FechaF Step 31
''''''        FechaDelMes = "#" & BuscarFecha(PrimerDiaMes(CLongFecha(FechaA))) & "# AND #" & BuscarFecha(UltimoDiaMes(CLongFecha(FechaA))) & "#"
'''
'''        sSQL = "UPDATE Facturas " _
'''             & "SET Total_Abonos = Total_Efectivo + Total_Banco + Total_Ret_Fuente + Total_Ret_IVA_B + Total_Ret_IVA_S + Otros_Abonos " _
'''             & "WHERE Item = '" & NumEmpresa & "' " _
'''             & "AND Periodo = '" & Periodo_Contable & "' " _
'''             & "AND Fecha <= #" & FechaFin & "# " _
'''             & EsFacturaIndividual _
'''             & "AND T <> 'A' "
'''        Ejecutar_SQL_SP sSQL
'''
'''        sSQL = "UPDATE Facturas " _
'''             & "SET Saldo_Actual = ROUND(Total_MN - Total_Abonos,2,0) " _
'''             & "WHERE Item = '" & NumEmpresa & "' " _
'''             & "AND Periodo = '" & Periodo_Contable & "' " _
'''             & "AND Fecha <= #" & FechaFin & "# " _
'''             & EsFacturaIndividual _
'''             & "AND T <> 'A' "
'''        Ejecutar_SQL_SP sSQL
''''''    Next FechaA
'''
'''    If SaldoReal Then
'''       sSQL = "UPDATE Facturas " _
'''            & "SET Saldo_MN = Saldo_Actual " _
'''            & "WHERE Item = '" & NumEmpresa & "' " _
'''            & "AND Periodo = '" & Periodo_Contable & "' " _
'''            & EsFacturaIndividual _
'''            & "AND T <> 'A' " _
'''            & "AND Saldo_MN <> Saldo_Actual "
'''       Ejecutar_SQL_SP sSQL
'''
'''       sSQL = "UPDATE Facturas " _
'''            & "SET T = 'C' " _
'''            & "WHERE Item = '" & NumEmpresa & "' " _
'''            & "AND Periodo = '" & Periodo_Contable & "' " _
'''            & "AND Saldo_MN <= 0 " _
'''            & "AND T <> 'A' " _
'''            & EsFacturaIndividual
'''       Ejecutar_SQL_SP sSQL
'''
'''       sSQL = "UPDATE Facturas " _
'''            & "SET T = 'P' " _
'''            & "WHERE Item = '" & NumEmpresa & "' " _
'''            & "AND Periodo = '" & Periodo_Contable & "' " _
'''            & "AND Saldo_MN > 0 " _
'''            & "AND T <> 'A' " _
'''            & EsFacturaIndividual
'''       Ejecutar_SQL_SP sSQL
'''
'''       Progreso_Barra.Mensaje_Box = "Actualizando Estado de las facturas"
'''       Progreso_Esperar
'''      'MsgBox CLongFecha(FechaI) & " - " & CLongFecha(FechaF)
'''       For FechaA = FechaI To FechaF Step 31
'''          'recolectamos las fechas de abonos
'''           IDMes = Month(CLongFecha(FechaA))
'''           IdAnio = Year(CLongFecha(FechaA))
'''           If SQL_Server Then
'''              sSQL = "UPDATE Detalle_Factura " _
'''                   & "SET T = F.T " _
'''                   & "FROM Detalle_Factura As DF, Facturas As F "
'''           Else
'''              sSQL = "UPDATE Detalle_Factura As DF, Facturas As F " _
'''                   & "SET DF.T = F.T "
'''           End If
'''           sSQL = sSQL _
'''                & "WHERE DF.Item = '" & NumEmpresa & "' " _
'''                & "AND DF.Periodo = '" & Periodo_Contable & "' " _
'''                & "AND F.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
'''                & "AND MONTH(DF.Fecha) = " & IDMes & " " _
'''                & "AND YEAR(DF.Fecha) = " & IdAnio & " " _
'''                & EsFacturaIndividualF _
'''                & "AND DF.Item = F.Item " _
'''                & "AND DF.Periodo = F.Periodo " _
'''                & "AND DF.Factura = F.Factura " _
'''                & "AND DF.TC = F.TC " _
'''                & "AND DF.Serie = F.Serie " _
'''                & "AND DF.Autorizacion = F.Autorizacion "
'''           Ejecutar_SQL_SP sSQL
'''
'''           If SQL_Server Then
'''              sSQL = "UPDATE Trans_Abonos " _
'''                   & "SET T = F.T " _
'''                   & "FROM Trans_Abonos As DF, Facturas As F "
'''           Else
'''              sSQL = "UPDATE Trans_Abonos As DF, Facturas As F " _
'''                   & "SET DF.T = F.T "
'''           End If
'''           sSQL = sSQL _
'''                & "WHERE DF.Item = '" & NumEmpresa & "' " _
'''                & "AND DF.Periodo = '" & Periodo_Contable & "' " _
'''                & "AND F.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
'''                & "AND MONTH(DF.Fecha) = " & IDMes & " " _
'''                & "AND YEAR(DF.Fecha) = " & IdAnio & " " _
'''                & EsFacturaIndividualF _
'''                & "AND DF.Item = F.Item " _
'''                & "AND DF.Periodo = F.Periodo " _
'''                & "AND DF.Factura = F.Factura " _
'''                & "AND DF.Autorizacion = F.Autorizacion " _
'''                & "AND DF.Serie = F.Serie " _
'''                & "AND DF.TP = F.TC "
'''           Ejecutar_SQL_SP sSQL
'''       Next FechaA
'''    End If
'''  Progreso_Final
'''End Sub

Public Sub Actualizar_Razon_Social(Optional FechaIniAut As String)
Dim AdoFA As ADODB.Recordset
Dim Idx As Integer
Dim FechaDeAut As String

    FechaDeAut = FechaIniAut
    RutaXMLRechazado = RutaDocumentos & "\Comprobantes no Autorizados\*.xml"
    If Existe_File(RutaXMLRechazado) Then Kill RutaXMLRechazado
    FechaIni = FechaSistema
    FechaFin = FechaSistema
    sSQL = "SELECT Autorizacion, MIN(Fecha) As Fecha_Min, MAX(Fecha) As Fecha_Max " _
         & "FROM Facturas " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND T <> '" & Anulado & "' " _
         & "AND LEN(Autorizacion) = 13 " _
         & "GROUP BY Autorizacion "
    Select_AdoDB AdoFA, sSQL
    If AdoFA.RecordCount > 0 Then
       FechaIni = AdoFA.fields("Fecha_Min")
       FechaFin = AdoFA.fields("Fecha_Max")
    End If
    AdoFA.Close
    
    If IsDate(FechaDeAut) And CFechaLong(FechaDeAut) < CFechaLong(FechaFin) Then FechaIni = FechaDeAut
    FechaIni = BuscarFecha(FechaIni)
    FechaFin = BuscarFecha(FechaFin)
    
    For Idx = 1 To 12
        If SQL_Server Then
           sSQL = "UPDATE Facturas " _
                & "SET RUC_CI = C.CI_RUC, Razon_Social = C.Cliente, TB = C.TD " _
                & "FROM Facturas As F, Clientes As C "
        Else
           sSQL = "UPDATE Facturas As F, Clientes As C " _
                & "SET F.RUC_CI = C.CI_RUC, F.Razon_Social = C.Cliente, F.TB = C.TD "
        End If
        sSQL = sSQL _
             & "WHERE F.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
             & "AND F.Item = '" & NumEmpresa & "' " _
             & "AND F.Periodo = '" & Periodo_Contable & "' " _
             & "AND LEN(F.Autorizacion) = 13 " _
             & "AND C.TD IN ('C','R','P') " _
             & "AND MONTH(F.Fecha) = " & Idx & " " _
             & "AND F.CodigoC = C.Codigo "
        Ejecutar_SQL_SP sSQL
        
        If SQL_Server Then
           sSQL = "UPDATE Facturas " _
                & "SET RUC_CI = CM.Cedula_R, Razon_Social = CM.Representante, TB = CM.TD " _
                & "FROM Facturas As F, Clientes_Matriculas As CM "
        Else
           sSQL = "UPDATE Facturas As F, Clientes_Matriculas As CM " _
                & "SET RUC_CI = CM.Cedula_R, F.Razon_Social = CM.Representante, F.TB = CM.TD "
        End If
        sSQL = sSQL _
             & "WHERE F.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
             & "AND F.Item = '" & NumEmpresa & "' " _
             & "AND F.Periodo = '" & Periodo_Contable & "' " _
             & "AND LEN(F.Autorizacion) = 13 " _
             & "AND CM.TD IN ('C','R','P') " _
             & "AND MONTH(F.Fecha) = " & Idx & " " _
             & "AND F.Item = CM.Item " _
             & "AND F.Periodo = CM.Periodo " _
             & "AND F.CodigoC = CM.Codigo "
        Ejecutar_SQL_SP sSQL
    Next Idx
End Sub

Public Sub Actualizar_Saldos_Reales_Facturas(TFA As Tipo_Facturas)
Dim AdoFechas As ADODB.Recordset
Dim IdAnio As Integer
Dim IDMes As Integer
Dim FechaI As Long
Dim FechaF As Long
Dim FechaA As Long
Dim PorFacturaNo As String
Dim PorFacturaNoF As String

    Progreso_Barra.Mensaje_Box = "Determinando fechas de proceso"
    Progreso_Iniciar
    Progreso_Barra.Valor_Maximo = 100
    TFA.Fecha_Desde = FechaSistema
    If TFA.Factura > 0 And Len(TFA.TC) = 2 And Len(TFA.Serie) = 6 Then
       PorFacturaNo = "AND TC = '" & TFA.TC & "' " _
                    & "AND Serie = '" & TFA.Serie & "' " _
                    & "AND Factura = " & TFA.Factura & " "
       PorFacturaNoF = "AND F.TC = '" & TFA.TC & "' " _
                     & "AND F.Serie = '" & TFA.Serie & "' " _
                     & "AND F.Factura = " & TFA.Factura & " "
    Else
       PorFacturaNo = ""
       PorFacturaNoF = ""
    End If
    If Not IsDate(TFA.Fecha_Hasta) Then TFA.Fecha_Hasta = FechaSistema
    
    sSQL = "SELECT T, MIN(Fecha) As Fecha_Min " _
         & "FROM Facturas " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND T <> '" & Anulado & "' " _
         & "AND Fecha <= #" & BuscarFecha(FechaSistema) & "# " _
         & "GROUP BY T "
    Select_AdoDB AdoFechas, sSQL
    If AdoFechas.RecordCount > 0 Then TFA.Fecha_Desde = AdoFechas.fields("Fecha_Min")
    AdoFechas.Close
    
    If CFechaLong(TFA.Fecha_Desde) > CFechaLong(TFA.Fecha_Hasta) Then TFA.Fecha_Desde = TFA.Fecha_Hasta
    FechaI = CFechaLong(TFA.Fecha_Desde)
    FechaF = CFechaLong(TFA.Fecha_Hasta)
    FechaIni = BuscarFecha(TFA.Fecha_Desde)
    FechaFin = BuscarFecha(TFA.Fecha_Hasta)
  
   'Redondeamos los abonos a dos decimales
    Progreso_Barra.Mensaje_Box = "Redondeando Abonos"
    Progreso_Esperar
    
    sSQL = "DELETE * " _
         & "FROM Asiento_Abonos " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND CodigoU = '" & CodigoUsuario & "' "
    Ejecutar_SQL_SP sSQL
    
    sSQL = "UPDATE Facturas " _
         & "SET T = 'P', Saldo_MN = Total_MN, Abonos_MN = 0 " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND T <> 'A' " _
         & PorFacturaNo
    Ejecutar_SQL_SP sSQL
    
    Progreso_Barra.Mensaje_Box = "Actualizando Fecha de Cancelacion de Facturas"
    Progreso_Esperar

   'Actualizamos Fecha maxima de pagos de abonos
   '& "AND Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# "
    sSQL = "INSERT INTO Asiento_Abonos (Periodo,Item,CodigoC,Fecha,TP,Serie,Factura,Autorizacion,Abono,CodigoU) " _
         & "SELECT Periodo,Item,CodigoC,MAX(Fecha),TP,Serie,Factura,Autorizacion,ROUND(SUM(Abono),2,0) As TAbono,'" & CodigoUsuario & "' " _
         & "FROM Trans_Abonos " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND Fecha <= #" & FechaFin & "# " _
         & "AND T <> 'A' " _
         & PorFacturaNo _
         & "GROUP BY Periodo,Item,CodigoC,TP,Serie,Factura,Autorizacion " _
         & "ORDER BY Periodo,Item,CodigoC,TP,Serie,Factura,Autorizacion "
    Ejecutar_SQL_SP sSQL
       
   'Actualizamos fecha maxima de pago y los abonos
    Progreso_Barra.Mensaje_Box = "Actualizando Fecha de Cancelacion de Facturas"
    Progreso_Esperar
    If SQL_Server Then
       sSQL = "UPDATE Facturas " _
            & "SET Fecha_C = AA.Fecha, Abonos_MN = AA.Abono " _
            & "FROM Facturas As F,Asiento_Abonos As AA "
    Else
       sSQL = "UPDATE Facturas As F,Asiento_Abonos As AA " _
            & "SET F.Fecha_C = AA.Fecha, F.Abonos_MN = AA.Abono  "
    End If
    sSQL = sSQL _
         & "WHERE F.Item = '" & NumEmpresa & "' " _
         & "AND F.Periodo = '" & Periodo_Contable & "' " _
         & "AND AA.CodigoU = '" & CodigoUsuario & "' " _
         & PorFacturaNoF _
         & "AND F.T <> 'A' " _
         & "AND F.Item = AA.Item " _
         & "AND F.Periodo = AA.Periodo " _
         & "AND F.TC = AA.TP " _
         & "AND F.Serie = AA.Serie " _
         & "AND F.Factura = AA.Factura " _
         & "AND F.Autorizacion = AA.Autorizacion "
    Ejecutar_SQL_SP sSQL
    
    sSQL = "DELETE * " _
         & "FROM Asiento_Abonos " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND CodigoU = '" & CodigoUsuario & "' "
    Ejecutar_SQL_SP sSQL
    
    sSQL = "UPDATE Facturas " _
         & "SET Saldo_MN = ROUND(Total_MN - Abonos_MN,2,0), Saldo_Actual = ROUND(Total_MN - Abonos_MN,2,0) " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & PorFacturaNo _
         & "AND T <> 'A' "
    Ejecutar_SQL_SP sSQL
            
    sSQL = "UPDATE Facturas " _
         & "SET T = 'C' " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND Saldo_MN <= 0 " _
         & PorFacturaNo _
         & "AND T <> 'A' "
    Ejecutar_SQL_SP sSQL
    Progreso_Barra.Mensaje_Box = "Actualizando Estado de las facturas"
    Progreso_Esperar
   'MsgBox CLongFecha(FechaI) & " - " & CLongFecha(FechaF)
    For FechaA = FechaI To FechaF Step 31
       'recolectamos las fechas de abonos
        IDMes = Month(CLongFecha(FechaA))
        IdAnio = Year(CLongFecha(FechaA))
        If SQL_Server Then
           sSQL = "UPDATE Detalle_Factura " _
                & "SET T = F.T " _
                & "FROM Detalle_Factura As DF, Facturas As F "
        Else
           sSQL = "UPDATE Detalle_Factura As DF, Facturas As F " _
                & "SET DF.T = F.T "
        End If
        sSQL = sSQL _
             & "WHERE DF.Item = '" & NumEmpresa & "' " _
             & "AND DF.Periodo = '" & Periodo_Contable & "' " _
             & "AND F.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
             & "AND MONTH(DF.Fecha) = " & IDMes & " " _
             & "AND YEAR(DF.Fecha) = " & IdAnio & " " _
             & PorFacturaNo _
             & "AND DF.Item = F.Item " _
             & "AND DF.Periodo = F.Periodo " _
             & "AND DF.Factura = F.Factura " _
             & "AND DF.TC = F.TC " _
             & "AND DF.Serie = F.Serie " _
             & "AND DF.Autorizacion = F.Autorizacion "
        Ejecutar_SQL_SP sSQL
    Next FechaA
    Progreso_Final
End Sub

Public Sub Actualizar_Bus(Codigo_Cliente As String, _
                          Bus_No As String, _
                          Curso As String)
Dim Fecha_Meses() As String
Dim CodigoCliente As String
Dim CodBusAnt As String
Dim AdoClientMat As ADODB.Recordset

  FechaInicial = Dato_DBF.FechaI
  FechaFinal = Dato_DBF.FechaF
  
  FechaIni = BuscarFecha(FechaInicial)
  FechaFin = BuscarFecha(FechaFinal)
  
 'Borramos datos del bus actual
  sSQL = "DELETE * " _
       & "FROM Clientes_Facturacion " _
       & "WHERE Codigo = '" & Codigo_Cliente & "' " _
       & "AND Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
       & "AND Item = '" & NumEmpresa & "' "
  Ejecutar_SQL_SP sSQL
  
 'If CFechaLong(FechaInicial) < CFechaLong(FechaSistema) Then FechaInicial = PrimerDiaMes(FechaSistema)
  If MidStrg(Bus_No, 1, 3) = "BUS" And Codigo_Cliente <> Ninguno Then
     RatonReloj
    'If CodigoCliente = "20100155" Then MsgBox "Actualizar Bus:" & vbCrLf & CodigoCliente
     CodigoInv = SinEspaciosDer(Bus_No)               ' Bus Asignado
     CodigoInv = "02." & Format$(Val(CodigoInv), "00")
     Valor = 0
     If IsNumeric(CodigoInv) Then
        sSQL = "SELECT Codigo_Inv, Producto, PVP " _
             & "FROM Catalogo_Productos " _
             & "WHERE Item = '" & NumEmpresa & "' " _
             & "AND Periodo = '" & Periodo_Contable & "' " _
             & "AND Codigo_Inv = '" & CodigoInv & "' "
        Select_AdoDB AdoClientMat, sSQL
        If AdoClientMat.RecordCount > 0 Then Valor = AdoClientMat.fields("PVP")
        AdoClientMat.Close
     End If
     
    'MsgBox sSQL & vbCrLf & "Valor a Facturar = " & Valor
     If IsNumeric(CodigoInv) And Len(Bus_No) >= 3 And Len(Curso) > 1 And Valor > 0 Then
        Total_Desc = 0
        sSQL = "SELECT Codigo, Descuento " _
             & "FROM Clientes_Matriculas " _
             & "WHERE Codigo = '" & Codigo_Cliente & "' " _
             & "AND Item = '" & NumEmpresa & "' " _
             & "AND Periodo = '" & Periodo_Contable & "' " _
             & "AND Descuento > 0 "
        Select_AdoDB AdoClientMat, sSQL
        If AdoClientMat.RecordCount > 0 Then Total_Desc = AdoClientMat.fields("Descuento")
        AdoClientMat.Close
        
        For JE = 1 To 12
            sSQL = "UPDATE Detalle_Factura " _
                 & "SET Mes_No = " & JE & " " _
                 & "WHERE Item = '" & NumEmpresa & "' " _
                 & "AND Periodo = '" & Periodo_Contable & "' " _
                 & "AND CodigoC = '" & Codigo_Cliente & "' " _
                 & "AND Mes = '" & MesesLetras(CByte(JE)) & "' " _
                 & "AND Mes_No = 0 " _
                 & "AND Fecha >= '" & BuscarFecha(FechaInicial) & "' "
            Ejecutar_SQL_SP sSQL
        Next JE
         
        sSQL = "SELECT T, CodigoC, Codigo, Fecha, Ticket, Mes_No " _
             & "FROM Detalle_Factura " _
             & "WHERE Item = '" & NumEmpresa & "' " _
             & "AND Periodo = '" & Periodo_Contable & "' " _
             & "AND Fecha >= '" & BuscarFecha(FechaInicial) & "' " _
             & "AND Ticket >= '" & CStr(Year(FechaInicial)) & "' " _
             & "AND CodigoC = '" & Codigo_Cliente & "' " _
             & "AND T <> 'A' " _
             & "ORDER BY Fecha DESC,Ticket DESC, Mes_No DESC "
        Select_AdoDB AdoClientMat, sSQL
        With AdoClientMat
         If .RecordCount > 0 Then
            'MsgBox .Fields("Fecha")
             NoMes = .fields("Mes_No")
             If (1 <= NoMes And NoMes <= 12) And IsNumeric(.fields("Ticket")) Then
                Mifecha = "01/" & CStr(NoMes) & "/" & .fields("Ticket")
                Mifecha = PrimerDiaMes(CLongFecha(CFechaLong(Mifecha) + 31))
             Else
                Mifecha = PrimerDiaMes(CLongFecha(CFechaLong(.fields("Fecha")) + 31))
             End If
         Else
             Mifecha = Dato_DBF.FechaI
         End If
        End With
        AdoClientMat.Close
        If CFechaLong(Mifecha) < CFechaLong(FechaInicial) Then Mifecha = FechaInicial
        
        Contador = (CFechaLong(FechaFinal) - CFechaLong(Mifecha)) / 31
        
        ReDim Fecha_Meses(Contador) As String
        Cadena = ""
        For JE = 0 To Contador - 1
            Fecha_Meses(JE) = CStr(UltimoDiaMes(Mifecha))
            Mifecha = CLongFecha(CFechaLong(Mifecha) + 31)
            Cadena = Cadena & Fecha_Meses(JE) & vbCrLf
        Next JE
       ' If Contador < 10 Then MsgBox "Actualizar Bus:" & vbCrLf & Cadena & vbCrLf & Codigo_Cliente & vbCrLf & "VALOR = " & Valor
       'If CodigoCliente = "20100155" Then MsgBox "Actualizar Bus:" & vbCrLf & Cadena & vbCrLf & vbCrLf & CodigoCliente
       'Empezamos a insertar los buses
        Contador = 0
        NoMes = Month(Dato_DBF.FechaI)
        CodigoP = TrimStrg(CStr(Year(Dato_DBF.FechaI)))
        sSQL = "DELETE * " _
             & "FROM Clientes_Facturacion " _
             & "WHERE Item = '" & NumEmpresa & "' " _
             & "AND Codigo = '" & Codigo_Cliente & "' " _
             & "AND Fecha >= '" & BuscarFecha(Dato_DBF.FechaI) & "' " _
             & "AND MidStrg(Codigo_Inv,1,2) = '02' "
        Ejecutar_SQL_SP sSQL
        Cadena = CodigoCliente
        For IE = 0 To UBound(Fecha_Meses) - 1
            NoMes = Month(Fecha_Meses(IE))
            CodigoP = TrimStrg(CStr(Year(Fecha_Meses(IE))))
            SetAdoAddNew "Clientes_Facturacion"
            SetAdoFields "T", Normal
            SetAdoFields "Codigo", Codigo_Cliente
            SetAdoFields "Codigo_Inv", CodigoInv
            SetAdoFields "Num_Mes", NoMes
            SetAdoFields "Mes", MesesLetras(NoMes)
            SetAdoFields "Valor", Valor
            SetAdoFields "Descuento", Total_Desc
            SetAdoFields "GrupoNo", Curso
            SetAdoFields "Periodo", CodigoP
            SetAdoFields "Fecha", Fecha_Meses(IE)
            SetAdoFields "Item", NumEmpresa
            SetAdoUpdate
        Next IE
     End If
     RatonNormal
  End If
End Sub

Public Sub Leer_Datos_FA_NV(TFA As Tipo_Facturas)
Dim AdoDBFac As ADODB.Recordset
Dim TotDescuento As Currency

    TFA.Fecha_Aut_GR = FechaSistema
    TFA.Hora = HoraSistema
    TFA.Hora_GR = HoraSistema
    TFA.Estado_SRI_GR = Ninguno
    TFA.Serie_GR = Ninguno
    TFA.ClaveAcceso_GR = Ninguno
    TFA.Autorizacion_GR = Ninguno
    TFA.Vendedor = Ninguno
    TFA.Remision = 0
    TFA.Comercial = Ninguno
    TFA.CIRUCComercial = Ninguno
    TFA.CIRUCEntrega = Ninguno
    TFA.Entrega = Ninguno
    TFA.CiudadGRI = Ninguno
    TFA.CiudadGRF = Ninguno
    TFA.Serie_GR = Ninguno
    TFA.FechaGRE = FechaSistema
    TFA.FechaGRI = FechaSistema
    TFA.FechaGRF = FechaSistema
    TFA.Pedido = Ninguno
    TFA.Zona = Ninguno
    TFA.Orden_Compra = 0
    TFA.Placa_Vehiculo = Ninguno
    TFA.Lugar_Entrega = Ninguno
    TFA.Descuento_X = 0
    TFA.Descuento_0 = 0
    TFA.Gavetas = 0
    TFA.Servicio = 0
    TFA.EsPorReembolso = False
    TFA.Si_Existe_Doc = False

    sSQL = "SELECT F.*,C.Cliente,C.CI_RUC,C.TD,C.Grupo,C.Direccion,C.DireccionT,C.Celular,C.Codigo,C.Ciudad,C.Email,C.Email2,C.EmailR," _
         & "C.Contacto,C.DirNumero,C.Fecha As Fecha_C " _
         & "FROM Facturas As F, Clientes As C " _
         & "WHERE F.Item = '" & NumEmpresa & "' " _
         & "AND F.Periodo = '" & Periodo_Contable & "' " _
         & "AND F.TC = '" & TFA.TC & "' " _
         & "AND F.Serie = '" & TFA.Serie & "' " _
         & "AND F.Autorizacion = '" & TFA.Autorizacion & "' " _
         & "AND F.Factura = " & TFA.Factura & " " _
         & "AND C.Codigo = F.CodigoC "
    Select_AdoDB AdoDBFac, sSQL
    With AdoDBFac
     If .RecordCount > 0 Then
        'Datos del SRI
         TFA.Si_Existe_Doc = True
         TFA.T = .fields("T")
         TFA.SP = .fields("SP")
         TFA.Porc_IVA = .fields("Porc_IVA")
         TFA.Porc_IVA_S = CStr(.fields("Porc_IVA") * 100)
         TFA.Cta_CxP = .fields("Cta_CxP")
         TFA.Cod_CxC = .fields("Cod_CxC")
         TFA.Estado_SRI = .fields("Estado_SRI")
         TFA.Error_SRI = .fields("Error_FA_SRI")
         TFA.ClaveAcceso = .fields("Clave_Acceso")
         TFA.CodigoU = .fields("CodigoU")
                
        'Encabezado de Facturas
         TFA.CodigoC = .fields("CodigoC")
         TFA.Contacto = .fields("Contacto")
         TFA.Cliente = .fields("Cliente")
         TFA.CI_RUC = .fields("CI_RUC")
         TFA.TD = .fields("TD")
         TFA.Razon_Social = .fields("Razon_Social")
         TFA.RUC_CI = .fields("RUC_CI")
         TFA.TB = .fields("TB")
         TFA.DireccionC = .fields("Direccion_RS")
         TFA.TelefonoC = .fields("Telefono_RS")
         TFA.DirNumero = .fields("DirNumero")
         TFA.Curso = .fields("Direccion")
         TFA.CiudadC = .fields("Ciudad")
         TFA.Grupo = .fields("Grupo")
         TFA.Cod_Ejec = .fields("Cod_Ejec")
         TFA.Imp_Mes = .fields("Imp_Mes")
         TFA.Fecha = .fields("Fecha")
         TFA.Fecha_V = .fields("Fecha_V")
         TFA.Fecha_C = .fields("Fecha_C")
         TFA.Fecha_Aut = .fields("Fecha_Aut")
         TFA.Hora = .fields("Hora")
         TFA.Tipo_Pago = .fields("Tipo_Pago")
         TFA.EmailC = .fields("Email")
         TFA.EmailC2 = .fields("Email2")
         TFA.EmailR = .fields("EmailR")
         TFA.Observacion = .fields("Observacion")
         TFA.Nota = .fields("Nota")
         TFA.Orden_Compra = .fields("Orden_Compra")
         TFA.Gavetas = .fields("Gavetas")
         If TFA.EmailR = Ninguno Then TFA.EmailR = EmailProcesos

        'SubTotales de la Factura
         TFA.Descuento = .fields("Descuento")
         TFA.Descuento2 = .fields("Descuento2")
         TFA.Descuento_0 = .fields("Desc_0")
         TFA.Descuento_X = .fields("Desc_X")
         TFA.SubTotal = .fields("SubTotal")
         TFA.Total_IVA = .fields("IVA")
         TFA.Con_IVA = .fields("Con_IVA")
         TFA.Sin_IVA = .fields("Sin_IVA")
         TFA.Servicio = .fields("Servicio")
         TFA.Total_MN = .fields("Total_MN")
         TFA.Saldo_MN = .fields("Saldo_MN")
         TFA.Saldo_Actual = .fields("Saldo_MN")
         TFA.Total_Descuento = TFA.Descuento + TFA.Descuento2
     End If
    End With
    AdoDBFac.Close
            
    sSQL = "SELECT * " _
         & "FROM Facturas_Auxiliares " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND Remision <> 0 " _
         & "AND TC = '" & TFA.TC & "' " _
         & "AND Serie = '" & TFA.Serie & "' " _
         & "AND Autorizacion = '" & TFA.Autorizacion & "' " _
         & "AND Factura = " & TFA.Factura & " "
    Select_AdoDB AdoDBFac, sSQL
    With AdoDBFac
     If .RecordCount > 0 Then
        'Guia de Remision
         TFA.Fecha_Aut_GR = .fields("Fecha_Aut_GR")
         TFA.Hora_GR = .fields("Hora_Aut_GR")
         TFA.Estado_SRI_GR = .fields("Estado_SRI_GR")
         TFA.Serie_GR = .fields("Serie_GR")
         TFA.ClaveAcceso_GR = .fields("Clave_Acceso_GR")
         TFA.Autorizacion_GR = .fields("Autorizacion_GR")
         TFA.Remision = .fields("Remision")
         TFA.Comercial = .fields("Comercial")
         TFA.CIRUCComercial = .fields("CIRUC_Comercial")
         TFA.CIRUCEntrega = .fields("CIRUC_Entrega")
         TFA.Entrega = .fields("Entrega")
         TFA.CiudadGRI = .fields("CiudadGRI")
         TFA.CiudadGRF = .fields("CiudadGRF")
         TFA.Serie_GR = .fields("Serie_GR")
         TFA.FechaGRE = .fields("FechaGRE")
         TFA.FechaGRI = .fields("FechaGRI")
         TFA.FechaGRF = .fields("FechaGRF")
         TFA.Pedido = .fields("Pedido")
         TFA.Zona = .fields("Zona")
         TFA.Orden_Compra = .fields("Orden_Compra")
         TFA.Placa_Vehiculo = .fields("Placa_Vehiculo")
         TFA.Lugar_Entrega = .fields("Lugar_Entrega")
     End If
    End With
    AdoDBFac.Close
            
    If Len(TFA.CIRUCComercial) > 1 And Len(TFA.CIRUCEntrega) > 1 Then
       sSQL = "SELECT Direccion " _
            & "FROM Clientes " _
            & "WHERE CI_RUC = '" & TFA.CIRUCComercial & "' "
       Select_AdoDB AdoDBFac, sSQL
       If AdoDBFac.RecordCount > 0 Then TFA.Dir_PartidaGR = AdoDBFac.fields("Direccion")
       AdoDBFac.Close
        
       sSQL = "SELECT Direccion " _
            & "FROM Clientes " _
            & "WHERE CI_RUC = '" & TFA.CIRUCEntrega & "' "
       Select_AdoDB AdoDBFac, sSQL
       If AdoDBFac.RecordCount > 0 Then TFA.Dir_EntregaGR = AdoDBFac.fields("Direccion")
       AdoDBFac.Close
    End If
    
    sSQL = "SELECT Nombre_Completo " _
         & "FROM Accesos " _
         & "WHERE Codigo = '" & TFA.Cod_Ejec & "' "
    Select_AdoDB AdoDBFac, sSQL
    If AdoDBFac.RecordCount > 0 Then TFA.Ejecutivo_Venta = AdoDBFac.fields("Nombre_Completo")
    AdoDBFac.Close
    
    sSQL = "SELECT Nombre_Completo " _
         & "FROM Accesos " _
         & "WHERE Codigo = '" & TFA.CodigoU & "' "
    Select_AdoDB AdoDBFac, sSQL
    If AdoDBFac.RecordCount > 0 Then TFA.Digitador = AdoDBFac.fields("Nombre_Completo")
    AdoDBFac.Close
    TFA.Digitador = Replace(TFA.Digitador, vbCrLf, "")
    
    sSQL = "SELECT Descripcion " _
         & "FROM Tabla_Referenciales_SRI " _
         & "WHERE Tipo_Referencia = 'FORMA DE PAGO' " _
         & "AND Codigo = '" & TFA.Tipo_Pago & "' "
    Select_AdoDB AdoDBFac, sSQL
    If AdoDBFac.RecordCount > 0 Then TFA.Tipo_Pago_Det = "Forma de Pago: " & ULCase(AdoDBFac.fields("Descripcion"))
    AdoDBFac.Close
    
    sSQL = "SELECT * " _
         & "FROM Facturas_Formatos " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND TC = '" & TFA.TC & "' " _
         & "AND Serie = '" & TFA.Serie & "' " _
         & "AND Cod_CxC = '" & TFA.Cod_CxC & "' " _
         & "AND #" & BuscarFecha(TFA.Fecha) & "# BETWEEN Fecha_Inicio and Fecha_Final " _
         & "ORDER BY Cod_CxC "
    Select_AdoDB AdoDBFac, sSQL
    If AdoDBFac.RecordCount > 0 Then
       TFA.CxC_Clientes = AdoDBFac.fields("Concepto")
       TFA.LogoFactura = AdoDBFac.fields("Formato_Factura")
       TFA.AltoFactura = AdoDBFac.fields("Largo")
       TFA.AnchoFactura = AdoDBFac.fields("Ancho")
       TFA.EspacioFactura = AdoDBFac.fields("Espacios")
       TFA.Pos_Factura = AdoDBFac.fields("Pos_Factura")
       TFA.DireccionEstab = AdoDBFac.fields("Direccion_Establecimiento")
       TFA.NombreEstab = AdoDBFac.fields("Nombre_Establecimiento")
       TFA.TelefonoEstab = AdoDBFac.fields("Telefono_Estab")
       TFA.Vencimiento = AdoDBFac.fields("Fecha_Final")
       TFA.CantFact = AdoDBFac.fields("Fact_Pag")
       TFA.LogoTipoEstab = RutaSistema & "\LOGOS\" & AdoDBFac.fields("Logo_Tipo_Estab") & ".jpg"
    End If
    AdoDBFac.Close
    
    sSQL = "SELECT Codigo " _
         & "FROM Detalle_Factura " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND Codigo = '99.41' " _
         & "AND TC = '" & TFA.TC & "' " _
         & "AND Serie = '" & TFA.Serie & "' " _
         & "AND Factura = " & TFA.Factura & " "
    Select_AdoDB AdoDBFac, sSQL
    If AdoDBFac.RecordCount > 0 Then TFA.EsPorReembolso = True
    AdoDBFac.Close

'    MsgBox "...Tiempo: " & Format(Time - TiempoFinal, "hh:mm:ss ff")

End Sub

Public Function Leer_Datos_FA_NV_Detalle(TFA As Tipo_Facturas) As ADODB.Recordset
Dim AdoDBDet As ADODB.Recordset
  'Comenzamos a recoger los detalles de la factura
  'SQLDec = "Precio " & CStr(Dec_PVP) & "|."
   sSQL = "SELECT DF.*,CP.Detalle,CP.Codigo_Barra,CP.Unidad,CM.Marca " _
        & "FROM Detalle_Factura As DF,Catalogo_Productos As CP,Catalogo_Marcas As CM " _
        & "WHERE DF.Item = '" & NumEmpresa & "' " _
        & "AND DF.Periodo = '" & Periodo_Contable & "' " _
        & "AND DF.TC = '" & TFA.TC & "' " _
        & "AND DF.Serie = '" & TFA.Serie & "' " _
        & "AND DF.Autorizacion = '" & TFA.Autorizacion & "' " _
        & "AND DF.Factura = " & TFA.Factura & " " _
        & "AND DF.Periodo = CP.Periodo " _
        & "AND DF.Periodo = CM.Periodo " _
        & "AND DF.Item = CP.Item " _
        & "AND DF.Item = CM.Item " _
        & "AND DF.Codigo = CP.Codigo_Inv " _
        & "AND DF.CodMarca = CM.CodMar " _
        & "ORDER BY DF.Codigo,DF.Ticket,DF.Mes,DF.ID "
   Select_AdoDB AdoDBDet, sSQL
   Set Leer_Datos_FA_NV_Detalle = AdoDBDet
End Function

Public Function Leer_RUC_CI_TARJETA(TarjetaNo As String, Abono As Currency, CodigoCli As String) As String
Dim AdoDBRUCCI As ADODB.Recordset
Dim TarjetaNoTemp As String
Dim Asteristicos As String
Dim IdTj As Long

   Asteristicos = ""
   For IdTj = 1 To Len(TarjetaNo)
       If MidStrg(TarjetaNo, IdTj, 1) = "*" Then Asteristicos = Asteristicos & "*"
   Next IdTj
   TarjetaNoTemp = TrimStrg(Replace(TarjetaNo, Asteristicos, "%"))
   
  'Comenzamos a recoger los detalles de la factura
  'SQLDec = "Precio " & CStr(Dec_PVP) & "|."
  '& "AND C.CI_RUC LIKE '%" & MidStrg(CodigoCli, 5, 4) & "' "
   sSQL = "SELECT TOP 1 C.Grupo, CM.Representante, CM.Cedula_R, C.CI_RUC, C.Cliente, CM.Cta_Numero,CM.Cta_Numero,CF.Periodo,CF.Num_Mes,CF.Mes,CF.Codigo_Inv,C.Codigo " _
        & "FROM Clientes AS C, Clientes_Matriculas AS CM, Clientes_Facturacion As CF " _
        & "WHERE CM.Item = '" & NumEmpresa & "' " _
        & "AND CM.Periodo = '" & Periodo_Contable & "' " _
        & "AND CM.Cta_Numero LIKE '" & TarjetaNoTemp & "' " _
        & "AND CF.X = '.' " _
        & "AND (CF.Valor - CF.Descuento - CF.Descuento2) = " & Abono & " " _
        & "AND C.Codigo = CM.Codigo " _
        & "AND C.Codigo = CF.Codigo " _
        & "AND CM.Item = CF.Item " _
        & "ORDER BY C.CI_RUC,CF.Periodo,CF.Num_Mes "
   Select_AdoDB AdoDBRUCCI, sSQL
  'MsgBox AdoDBRUCCI.RecordCount & vbCrLf & sSQL & vbCrLf & String(50, "*")
   If AdoDBRUCCI.RecordCount > 0 Then
      CodigoCli = AdoDBRUCCI.fields("Codigo")
      NombreCliente = AdoDBRUCCI.fields("Cliente")
      TarjetaNo = AdoDBRUCCI.fields("Cta_Numero")
      NoAnio = AdoDBRUCCI.fields("Periodo")
      NoMeses = AdoDBRUCCI.fields("Num_Mes")
      Mes = AdoDBRUCCI.fields("Mes")
      CodigoInv = AdoDBRUCCI.fields("Codigo_Inv")
     'CodigoEncontrado = CodigoEncontrado & " ,'" & AdoDBRUCCI.Fields("CI_RUC") & "'"
      Leer_RUC_CI_TARJETA = AdoDBRUCCI.fields("CI_RUC")
     'MsgBox TarjetaNo
   Else
      CodigoCli = Ninguno
      Leer_RUC_CI_TARJETA = Ninguno
   End If
End Function

Public Sub Imprimir_Ventas_Resumidas_Vendedor(Datas As Adodc, _
                                              MSChartRV As MSChart, _
                                              FormaImp As Byte, _
                                              SizeLetra As Integer, _
                                              Optional EsCampoCorto As Boolean)
Dim FinDoc As Boolean
Dim Contador As Integer
Dim ValorMayor As Currency
Dim Etiqueta() As String
Dim Vendedor() As Currency
Dim Porcentaje() As Currency

On Error GoTo Errorhandler

FinDoc = True
Mensajes = "Seguro de Imprimir en:" & vbCrLf & Printer.DeviceName & "?"
Titulo = "IMPRESION"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then

RatonReloj
InicioX = 0.5: InicioY = 0
DataAnchoCampos InicioX, Datas, SizeLetra, TipoArialNarrow, Orientacion_Pagina, EsCampoCorto
Pagina = 1
ValorMayor = 0
'Iniciamos la impresion
Ancho(CantCampos) = 19
Printer.FontBold = False
With Datas.Recordset
 If .RecordCount > 0 Then
    .MoveFirst
     EncabezadoData Datas
     Printer.FontSize = SizeLetra
     Printer.FontName = TipoArialNarrow
     Total = 0
     Contador = 1
     Printer.FontBold = False
     Cuadricula = True
     Codigo1 = .fields("Cod_Ejec")
     Codigo2 = .fields("Nombre_Vendedor")
     Imprimir_Linea_H PosLinea, Ancho(0), Ancho(CantCampos), Negro
     PosLinea = PosLinea + 0.05
     PrinterTexto Ancho(0), PosLinea, Codigo1
     PrinterTexto Ancho(1), PosLinea, Codigo2
     Do While Not .EOF
        Printer.Line (Ancho(0), PosLinea - 0.05)-(Ancho(0), PosLinea + AltoLetra + 0.05), Negro
       'MsgBox Printer.FontName
        If .fields("Cuenta") = "SUBTOTAL VENDEDOR" Then
            Printer.FontBold = True
            PosLinea = PosLinea + 0.05
            Imprimir_Linea_H PosLinea, Ancho(0), Ancho(CantCampos), Negro
            PosLinea = PosLinea + 0.05
            Codigo1 = .fields("Cod_Ejec")
            Codigo2 = .fields("Nombre_Vendedor")
            PrinterTexto Ancho(0), PosLinea, Codigo1
            PrinterTexto Ancho(1), PosLinea, Codigo2
        End If
        If Codigo1 <> .fields("Cod_Ejec") Then
           Codigo1 = .fields("Cod_Ejec")
           Codigo2 = .fields("Nombre_Vendedor")
           PrinterTexto Ancho(0), PosLinea, Codigo1
           PrinterTexto Ancho(1), PosLinea, Codigo2
        End If
        PrinterFields Ancho(2), PosLinea, .fields("Grupo"), True
        PrinterFields Ancho(3), PosLinea, .fields("Cuenta"), True
        PrinterFields Ancho(4), PosLinea, .fields("Cantidad"), True
        PrinterFields Ancho(5), PosLinea, .fields("Cuota"), True
        Printer.Line (Ancho(CantCampos), PosLinea - 0.05)-(Ancho(CantCampos), PosLinea + AltoLetra + 0.05), Negro
        If .fields("Cuenta") = "SUBTOTAL VENDEDOR" Then
            PosLinea = PosLinea + 0.34
            Imprimir_Linea_H PosLinea, Ancho(0), Ancho(CantCampos), Negro
            PosLinea = PosLinea + 0.05
            Contador = Contador + 1
            If ValorMayor < .fields("Cantidad") Then ValorMayor = .fields("Cantidad")
            Printer.FontBold = False
        Else
            PosLinea = PosLinea + 0.34
        End If
''        If Cuadricula Then
''           Imprimir_Linea_H PosLinea, Ancho(0), Ancho(CantCampos), Negro
''           PosLinea = PosLinea + 0.05
''        End If
        If PosLinea >= LimiteAlto Then
           Imprimir_Linea_H PosLinea, Ancho(0), Ancho(CantCampos)
           Printer.NewPage
           EncabezadoData Datas
           Printer.FontSize = SizeLetra
           Printer.FontName = TipoArialNarrow
        End If
        If .fields("Cuenta") = "SUBTOTAL VENDEDOR" Then Total = Total + .fields("Cantidad")
       .MoveNext
     Loop
 End If
End With
Printer.FontBold = True
PosLinea = PosLinea + 0.02
Imprimir_Linea_H PosLinea, Ancho(0), Ancho(CantCampos), Negro
PosLinea = PosLinea + 0.05
PrinterVariables Ancho(4), PosLinea, Total
PosLinea = PosLinea + 0.2
Cuadricula = False
 If Contador > 1 Then
    Contador = Contador - 1
    ReDim Etiqueta(1 To Contador) As String
    ReDim Vendedor(1 To Contador) As Currency
    ReDim Porcentaje(1 To Contador) As Currency
    With Datas.Recordset
     If .RecordCount > 0 Then
        .MoveFirst
         Contador = 1
         Do While Not .EOF
            If .fields("Cuenta") = "SUBTOTAL VENDEDOR" Then
                Etiqueta(Contador) = .fields("Cod_Ejec")
                Vendedor(Contador) = .fields("Cantidad")
                Porcentaje(Contador) = Val(.fields("Cuota"))
                Contador = Contador + 1
            End If
           .MoveNext
         Loop
     End If
    End With
   'Empezamos a dibujar las barrar de la malla
    Contador = Contador - 1
    With MSChartRV
        .width = 15000
        .Height = 9000
        .TitleText = "VENTAS RESUMIDAS POR VENDEDOR"
'''       .RowCount = 1
'''       .ColumnCount = Contador
'''       .RowLabel = ""
        .Plot.axis(MSChart20Lib.VtChAxisId.VtChAxisIdX).AxisTitle.Text = "VENDEDORES"
        .Plot.axis(MSChart20Lib.VtChAxisId.VtChAxisIdY).AxisTitle.Text = "TOTALES EN USD"
        .DataGrid.SetSize Contador, 1, Contador, 1
        
       ' Hacemos las barras
        .chartType = VtChChartType2dBar
         For IE = 1 To Contador
            .DataGrid.RowLabel(IE, 1) = Etiqueta(IE)
         Next IE
         
         For IE = 1 To Contador
            .DataGrid.SetData IE, 1, Vendedor(IE), 0
         Next IE
        .EditCopy
         RutaDestino = RutaSysBases & "\TEMP\MSChartB" & CodigoUsuario & ".gif"
         SavePicture Clipboard.GetData(3), RutaDestino
         Printer.PaintPicture LoadPicture(RutaDestino), 1, PosLinea
        ' Clipboard.Clear
         Kill RutaDestino
    
       ' Hacemos el pastel
        .chartType = VtChChartType2dPie
         For IE = 1 To Contador
            .DataGrid.SetData IE, 1, Porcentaje(IE), 0
         Next IE
        .EditCopy
         RutaDestino = RutaSysBases & "\TEMP\MSChartP" & CodigoUsuario & ".gif"
         SavePicture Clipboard.GetData(3), RutaDestino
         Printer.PaintPicture LoadPicture(RutaDestino), 10, PosLinea
         'Clipboard.Clear
         Kill RutaDestino
    End With
 End If
 
RatonNormal
MensajeEncabData = ""
If FinDoc Then Printer.EndDoc
Cuadricula = False
Exit Sub
Errorhandler:
    RatonNormal
    ErrorDeImpresion
    Exit Sub
Else
    RatonNormal
End If
End Sub

Public Sub Imprimir_Tiempo_Credito(Datas As Adodc, _
                                   FinDoc As Boolean, _
                                   FormaImp As Byte, _
                                   SizeLetra As Integer, _
                                   Optional EsCampoCorto As Boolean)
On Error GoTo Errorhandler
Dim Tot(7) As Currency

Mensajes = "Seguro de Imprimir en:" & vbCrLf & Printer.DeviceName & "?"
Titulo = "IMPRESION"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then

RatonReloj
Cuadricula = True
InicioX = 0.5: InicioY = 0
DataAnchoCampos InicioX, Datas, SizeLetra, TipoArialNarrow, Orientacion_Pagina, EsCampoCorto
Pagina = 1
'Iniciamos la impresion
Total = 0
For I = 0 To 6
    Tot(I) = 0
Next I
Printer.FontBold = False
With Datas.Recordset
If .RecordCount > 0 Then
    .MoveFirst
     EncabezadoData Datas
     Printer.FontSize = SizeLetra
     Printer.FontName = TipoArialNarrow
     Do While Not .EOF
        If .fields("Clientes") = "zz" & String(40, " ") & "SUBTOTALES" Then
            Printer.FontBold = True
            Imprimir_Linea_H PosLinea, Ancho(0), Ancho(CantCampos), Negro
            PosLinea = PosLinea + 0.05
        End If
        'MsgBox Printer.FontName
        PrinterAllFields CantCampos, PosLinea, Datas, True, False
        If .fields("Clientes") = "zz" & String(40, " ") & "SUBTOTALES" Then
            PosLinea = PosLinea + 0.4
            Imprimir_Linea_H PosLinea, Ancho(0), Ancho(CantCampos), Gris
            PosLinea = PosLinea + 0.1
            Contador = Contador + 1
            Printer.FontBold = False
            For I = 0 To 6
                Tot(I) = Tot(I) + .fields(I + 4)
            Next I
            Total = Total + .fields("Saldo_Total")
        Else
            PosLinea = PosLinea + 0.4
        End If
        If Cuadricula Then
           Imprimir_Linea_H PosLinea, Ancho(0), Ancho(CantCampos), Gris
           PosLinea = PosLinea + 0.05
        End If
        If PosLinea >= LimiteAlto Then
           Imprimir_Linea_H PosLinea, Ancho(0), Ancho(CantCampos)
           Printer.NewPage
           EncabezadoData Datas
           Printer.FontSize = SizeLetra
           Printer.FontName = TipoArialNarrow
        End If
       .MoveNext
     Loop
End If
End With
Printer.FontBold = True
PosLinea = PosLinea + 0.02
For I = 0 To 6
    PrinterVariables Ancho(I + 4), PosLinea, Tot(I)
Next I
Cuadricula = False
RatonNormal
MensajeEncabData = ""
If FinDoc Then Printer.EndDoc
Cuadricula = False
Exit Sub
Errorhandler:
    RatonNormal
    ErrorDeImpresion
    Exit Sub
Else
    RatonNormal
End If
End Sub

Public Sub Actualiza_Procesado_Kardex(CodigoInv As String)
Dim SQLKardex As String
    If Len(CodigoInv) > 2 Then
       SQLKardex = "UPDATE Trans_Kardex " _
                 & "SET Procesado = 0 " _
                 & "WHERE Item = '" & NumEmpresa & "' " _
                 & "AND Periodo = '" & Periodo_Contable & "' " _
                 & "AND Codigo_Inv = '" & CodigoInv & "' "
       Ejecutar_SQL_SP SQLKardex
    End If
End Sub

Public Sub Actualiza_Procesado_Kardex_Factura(TFA As Tipo_Facturas)
Dim SQLKardex As String
    RatonReloj
    SQLKardex = "UPDATE Trans_Kardex " _
              & "SET Procesado = 0 " _
              & "FROM Trans_Kardex As TK, Detalle_Factura As DF " _
              & "WHERE DF.Item = '" & NumEmpresa & "' " _
              & "AND DF.Periodo = '" & Periodo_Contable & "' " _
              & "AND DF.TC = '" & TFA.TC & "' " _
              & "AND DF.Serie = '" & TFA.Serie & "' " _
              & "AND DF.Factura = " & TFA.Factura & " " _
              & "AND TK.Item = DF.Item " _
              & "AND TK.Periodo = DF.Periodo " _
              & "AND TK.Codigo_Inv = DF.Codigo "
    Ejecutar_SQL_SP SQLKardex
    
    SQLKardex = "UPDATE Trans_Kardex " _
              & "SET Procesado = 0 " _
              & "FROM Trans_Kardex As TK, Asiento_NC As ANC " _
              & "WHERE TK.Item = '" & NumEmpresa & "' " _
              & "AND TK.Periodo = '" & Periodo_Contable & "' " _
              & "AND ANC.CodigoU = '" & CodigoUsuario & "' " _
              & "AND TK.Item = ANC.Item " _
              & "AND TK.Codigo_Inv = ANC.CODIGO "
    Ejecutar_SQL_SP SQLKardex
    
    SQLKardex = "UPDATE Trans_Kardex " _
              & "SET Procesado = 0 " _
              & "FROM Trans_Kardex As TK, Asiento_F As AF " _
              & "WHERE TK.Item = '" & NumEmpresa & "' " _
              & "AND TK.Periodo = '" & Periodo_Contable & "' " _
              & "AND AF.CodigoU = '" & CodigoUsuario & "' " _
              & "AND TK.Item = AF.Item " _
              & "AND TK.Codigo_Inv = AF.CODIGO "
    Ejecutar_SQL_SP SQLKardex
    RatonNormal
End Sub

Public Sub Actualiza_Procesado_Kardex_Rango_Factura(TFA As Tipo_Facturas)
Dim SQLKardex As String
    If TFA.Desde <= TFA.Hasta Then
       RatonReloj
       SQLKardex = "UPDATE Trans_Kardex " _
                 & "SET Procesado = 0 " _
                 & "FROM Trans_Kardex As TK, Detalle_Factura As DF " _
                 & "WHERE DF.Item = '" & NumEmpresa & "' " _
                 & "AND DF.Periodo = '" & Periodo_Contable & "' " _
                 & "AND DF.TC = '" & TFA.TC & "' " _
                 & "AND DF.Serie = '" & TFA.Serie & "' " _
                 & "AND DF.Factura BETWEEN " & TFA.Desde & " AND " & TFA.Hasta & " " _
                 & "AND TK.Item = DF.Item " _
                 & "AND TK.Periodo = DF.Periodo " _
                 & "AND TK.Codigo_Inv = DF.Codigo "
       Ejecutar_SQL_SP SQLKardex
       RatonNormal
    End If
End Sub

Public Function Reporte_Cartera_Clientes_PDF(FechaInicioHistoria As String, CodigoCliente As String, SalidaExcel As Boolean, VerDocumento As Boolean) As Boolean
Dim AdoCarteraDB As ADODB.Recordset
Dim NombFilePict As String
Dim EmailCli As String
Dim PosCampo(1 To 10) As Single
Dim SiExisteDatos As Boolean

   sSQL = "SELECT C.Cliente, RCC.T, RCC.TC, RCC.Serie, RCC.Factura, RCC.Fecha, RCC.Detalle, RCC.Anio, RCC.Mes, RCC.Cargos, RCC.Abonos, RCC.Saldo, RCC.CodigoC, " _
        & "C.Email, C.Email2, C.EmailR, C.Direccion " _
        & "FROM Reporte_Cartera_Clientes As RCC, Clientes As C " _
        & "WHERE RCC.Item = '" & NumEmpresa & "' " _
        & "AND RCC.CodigoU = '" & CodigoUsuario & "' " _
        & "AND RCC.T <> 'A' " _
        & "AND RCC.CodigoC = C.Codigo " _
        & "ORDER BY C.Cliente, RCC.TC, RCC.Serie, RCC.Factura, RCC.Anio, RCC.Mes, RCC.ID "
   Select_AdoDB AdoCarteraDB, sSQL
   With AdoCarteraDB
    If .RecordCount > 0 Then
        EmailCli = ""
        Insertar_Mail EmailCli, .fields("EmailR")
        Insertar_Mail EmailCli, .fields("Email2")
        Insertar_Mail EmailCli, .fields("Email")
        
        SiExisteDatos = True
        RutaDocumentoPDF = ""
        MensajeEncabData = "REPORTE CARTERA DE CLIENTES"
        SQLMsg1 = ""
        SQLMsg2 = ""
        SQLMsg3 = ""
        If SalidaExcel Then
           Exportar_AdoDB_Excel AdoCarteraDB, "Reporte Carte de Clientes"
        Else
           NombFilePict = "RCC_" & BuscarFecha(FechaSistema) & "-" & CodigoCliente
           SetNombrePRN = Impresota_PDF
          'Geneeramos el documento
           tPrint.TipoImpresion = Es_PDF
           tPrint.NombreArchivo = NombFilePict
           tPrint.TituloArchivo = "Reporte de Cartera de Clientes " & CodigoCliente
           tPrint.TipoLetra = TipoHelvetica
           tPrint.OrientacionPagina = 1
           tPrint.PaginaA4 = True
           tPrint.EsCampoCorto = True
           tPrint.VerDocumento = VerDocumento
           Set cPrint = New cImpresion
           cPrint.iniciaImpresion
           PosLinea = 1
           PosCampo(1) = 1.5    ' T
           PosCampo(2) = 1.9    ' TC
           PosCampo(3) = 2.5    ' Serie
           PosCampo(4) = 3.6    ' Factura
           PosCampo(5) = 4.8    ' Fecha
           PosCampo(6) = 6.5    ' Detalle
           PosCampo(7) = 15.3   ' Anio
           PosCampo(8) = 16.1   ' Mes
           PosCampo(9) = 16.9   ' Cargos
           PosCampo(10) = 18.8  ' Abonos
           
           cPrint.printEncabezado 1.2, PosLinea, TipoHelvetica
          'Pagina No. 1
           NombreCliente = .fields("Cliente")
           DireccionCli = .fields("Direccion")
           TipoCta = .fields("TC")
           SerieFactura = .fields("Serie")
           Factura_No = .fields("Factura")
           CodigoP = .fields("Detalle")
           
           cPrint.fondoDeLetra = Negro
           cPrint.tipoNegrilla = True
           cPrint.PorteDeLetra = 8
           cPrint.colorDeLetra = Negro
           If CodigoCliente <> "Todos" Then
              cPrint.printTexto PosCampo(1), PosLinea, "CLIENTE: " & NombreCliente
              PosLinea = PosLinea + 0.4
              cPrint.printTexto PosCampo(1), PosLinea, "UBICACION: " & DireccionCli
              PosLinea = PosLinea + 0.4
              cPrint.printTexto PosCampo(1), PosLinea, "EMAILS: " & EmailCli
              PosLinea = PosLinea + 0.5
           End If
           Cadena = "La informacion presente reposa en la base de dato de la Institucion, corte realizado desde " _
                  & FechaStrg(FechaInicioHistoria) & " al " & FechaStrg(FechaSistema) & ", " _
                  & "cualquier informacion adicional comuniquese a la institucion."
           PosLinea = cPrint.printTextoMultiple(1.5, PosLinea, Cadena, 18.9)
           PosLinea = PosLinea + 0.5
           cPrint.printTexto PosCampo(1), PosLinea, "T"
           cPrint.printTexto PosCampo(2), PosLinea, "TC"
           cPrint.printTexto PosCampo(3), PosLinea, "Serie"
           cPrint.printTexto PosCampo(4), PosLinea, "Factura"
           cPrint.printTexto PosCampo(5), PosLinea, "Fecha"
           cPrint.printTexto PosCampo(6), PosLinea, "Detalle"
           cPrint.printTexto PosCampo(7), PosLinea, "Anio"
           cPrint.printTexto PosCampo(8), PosLinea, "Mes"
           cPrint.printTexto PosCampo(9), PosLinea, "Cargos"
           cPrint.printTexto PosCampo(10), PosLinea, "Abonos"
           PosLinea = PosLinea + 0.4
           cPrint.printLinea 1.2, PosLinea - 0.55, cPrint.dAnchoPapel - 1, PosLinea - 0.55
           cPrint.printLinea 1.2, PosLinea - 0.05, cPrint.dAnchoPapel - 1, PosLinea - 0.05
           PosLinea = PosLinea + 0.1
           cPrint.colorDeLetra = Negro
           cPrint.tipoNegrilla = False
           cPrint.printFields PosCampo(1), PosLinea, .fields("T")
           cPrint.printFields PosCampo(2), PosLinea, .fields("TC")
           cPrint.printFields PosCampo(3), PosLinea, .fields("Serie")
           cPrint.printFields PosCampo(4), PosLinea, .fields("Factura")
           cPrint.printFields PosCampo(5), PosLinea, .fields("Fecha")
           cPrint.printFields PosCampo(6), PosLinea, .fields("Detalle")
           cPrint.printFields PosCampo(7), PosLinea, .fields("Anio")
           cPrint.printFields PosCampo(8), PosLinea, .fields("Mes")
           cPrint.printFields PosCampo(9), PosLinea, .fields("Cargos")
           cPrint.printFields PosCampo(10), PosLinea, .fields("Abonos")
           PosLinea = PosLinea + 0.4
          .MoveNext
           Do While Not .EOF
              If InStr(.fields("Detalle"), "S A L D O   T O T A L") Then
                 PosLinea = PosLinea + 0.1
                 cPrint.printLinea 1.2, PosLinea - 0.1, cPrint.dAnchoPapel - 1, PosLinea - 0.1
                 cPrint.printLinea 1.2, PosLinea + 0.3, cPrint.dAnchoPapel - 1, PosLinea + 0.3
              Else
                 cPrint.printFields PosCampo(1), PosLinea, .fields("T")
                 cPrint.printFields PosCampo(2), PosLinea, .fields("TC")
                 cPrint.printFields PosCampo(3), PosLinea, .fields("Serie")
                 cPrint.printFields PosCampo(4), PosLinea, .fields("Factura")
                 cPrint.printFields PosCampo(5), PosLinea, .fields("Fecha")
                 cPrint.printFields PosCampo(7), PosLinea, .fields("Anio")
                 cPrint.printFields PosCampo(8), PosLinea, .fields("Mes")
              End If
              If .fields("Detalle") <> CodigoP Then
                  cPrint.printFields PosCampo(6), PosLinea, .fields("Detalle")
                  CodigoP = .fields("Detalle")
              End If
              cPrint.printFields PosCampo(9), PosLinea, .fields("Cargos")
              If InStr(.fields("Detalle"), "S A L D O   T O T A L") Then
                 cPrint.printFields PosCampo(10), PosLinea, .fields("Saldo")
              Else
                 cPrint.printFields PosCampo(10), PosLinea, .fields("Abonos")
              End If
              PosLinea = PosLinea + 0.4
              If InStr(.fields("Detalle"), "S A L D O   T O T A L") Then PosLinea = PosLinea + 0.1
              'Siguiente Pagina
              If PosLinea > (cPrint.dAltoPapel - 1.5) Then
                 cPrint.paginaNueva
                 PosLinea = 1
                 cPrint.printEncabezado 1.2, PosLinea, TipoHelvetica
                 PosLinea = PosLinea - 0.15
                 cPrint.colorDeLetra = Negro
                 cPrint.tipoNegrilla = True
                 cPrint.PorteDeLetra = 8
                 cPrint.printTexto PosCampo(1), PosLinea, "T"
                 cPrint.printTexto PosCampo(2), PosLinea, "TC"
                 cPrint.printTexto PosCampo(3), PosLinea, "Serie"
                 cPrint.printTexto PosCampo(4), PosLinea, "Factura"
                 cPrint.printTexto PosCampo(5), PosLinea, "Fecha"
                 cPrint.printTexto PosCampo(6), PosLinea, "Detalle"
                 cPrint.printTexto PosCampo(7), PosLinea, "Anio"
                 cPrint.printTexto PosCampo(8), PosLinea, "Mes"
                 cPrint.printTexto PosCampo(9), PosLinea, "Cargos"
                 cPrint.printTexto PosCampo(10), PosLinea, "Abonos"
                 PosLinea = PosLinea + 0.4
                 cPrint.printLinea 1.2, PosLinea - 0.05, cPrint.dAnchoPapel - 1, PosLinea - 0.05
                 PosLinea = PosLinea + 0.1
                 cPrint.tipoNegrilla = False
              End If
             .MoveNext
           Loop
           EmailCli = TMail.para
          'fin del documento
           cPrint.finalizaImpresion
        End If
    Else
       SiExisteDatos = False
    End If
   End With
   AdoCarteraDB.Close
   Reporte_Cartera_Clientes_PDF = SiExisteDatos
End Function

Public Sub Prueba_Envio_de_Correos()
    TMail.ListaMail = 255
    TMail.TipoDeEnvio = "CO"
   'MsgBox RutaBackup
    TMail.Asunto = "Prueba de Mails por imap.diskcoversystem.com"
    TMail.MensajeHTML = Leer_Archivo_Texto(RutaSistema & "\JAVASCRIPT\email_recibo.html")
    
    html_Informacion_adicional = "<strong>INFORMACION ADICIONAL:</strong><br>" _
                               & "Importe total: USD 150,00<br>" _
                               & "Importe total: USD 150,00<br>" _
                               & "Importe total: USD 150,00<br>"
                               
    html_Detalle_adicional = "<tr>" _
                           & "<td>13/12/2024</td>" _
                           & "<td>Madera</td>" _
                           & "<td class='row text-right'>150,00</td>" _
                           & "</tr>" _
                           & "<tr>" _
                           & "<td>13/12/2024</td>" _
                           & "<td>Madera</td>" _
                           & "<td class='row text-right'>180,00</td>" _
                           & "</tr>"
    FA.Fecha = FechaSistema
    FA.Recibo_No = Format(FA.Fecha, "yyyymmdd") & Format(FA.Factura, "000000000")
    'TMail.MensajeHTML = ""
    'TMail.Mensaje = "Esta es una prueba de Correo Electronico enviado por DNS-EXIT, " _
                  & "mensaje enviado desde el PC: " & IP_PC.Nombre_PC & ", a las: " & Time & ", " _
                  & "de la empresa: " & Empresa & "."

    TMail.Adjunto = "C:\SYSBASES\CE\CE999\Comprobantes Autorizados\2409202401070216417900110010030000025251234567816.xml"
    TMail.para = ""
    Insertar_Mail TMail.para, "diskcoversystem@msn.com"
    Insertar_Mail TMail.para, "diskcover.system@yahoo.com"
    Insertar_Mail TMail.para, "diskcover.system@gmail.com"
    Insertar_Mail TMail.para, "informacion@diskcoversystem.com"
    FEnviarCorreos.Show 1
    TMail.para = ""
    TMail.ListaMail = 255
End Sub


