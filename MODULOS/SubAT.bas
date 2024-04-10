Attribute VB_Name = "SubATSRI"
Option Explicit

Global CodPorIce, CodPorIva, CodRetBien, CodRetServ As String

Dim DocumentoXML As MSXML2.DOMDocument30

Public Function CampoXML(Campo_XML As String, Valor_XML As Variant, Optional Decimales As Byte) As String
Dim Result_XML As String
Dim sValor_XML As String
    Result_XML = ""
    sValor_XML = Sin_Signos_Especiales(CStr(Valor_XML))
    If IsNumeric(sValor_XML) Then
       If Decimales > 0 Then sValor_XML = Format$(Val(sValor_XML), "#0." & String$(Decimales, "0"))
    End If
    If sValor_XML = "" Then sValor_XML = " "
    Result_XML = TrimStrg(sValor_XML)
   'MsgBox "Antes: " & Valor_XML & vbCrLf & "Despues: " & SValor_XML
    If Campo_XML <> "" Then Result_XML = "<" & Campo_XML & ">" & sValor_XML & "</" & Campo_XML & ">"
    CampoXML = Result_XML
End Function

Public Function CampoIdXML(Campo_XML As String, ID As String, Valor_XML As Variant, Optional Decimales As Byte) As String
Dim Result_XML As String
Dim sValor_XML As String
    Result_XML = ""
    sValor_XML = Sin_Signos_Especiales(CStr(Valor_XML))
    If IsNumeric(sValor_XML) Then
       If Decimales > 0 Then sValor_XML = Format$(Val(sValor_XML), "#0." & String$(Decimales, "0"))
    End If
    If sValor_XML = "" Then sValor_XML = " "
    Result_XML = TrimStrg(sValor_XML)
   'MsgBox "Antes: " & Valor_XML & vbCrLf & "Despues: " & SValor_XML
    If Campo_XML <> "" Then Result_XML = "<" & Campo_XML & " " & ID & ">" & sValor_XML & "</" & Campo_XML & ">"
    CampoIdXML = Result_XML
End Function

Public Function AbrirXML(Campo_XML As String, Optional IdXML As String) As String
Dim Result_XML As String
  Result_XML = ""
  If Campo_XML <> "" Then
     If IdXML <> "" Then
        Result_XML = "<" & Campo_XML & " " & IdXML & ">"
     Else
        Result_XML = "<" & Campo_XML & ">"
     End If
  End If
  AbrirXML = Result_XML
End Function

Public Function CerrarXML(Campo_XML As String) As String
Dim Result_XML As String
  Result_XML = ""
  If Campo_XML <> "" Then Result_XML = "</" & Campo_XML & ">"
  CerrarXML = Result_XML
End Function

Public Function VacioXML(Campo_XML As String) As String
Dim Result_XML As String
  Result_XML = ""
  If Campo_XML <> "" Then Result_XML = "<" & Campo_XML & " />"
  VacioXML = Result_XML
End Function

Public Function Mes_Año(Fecha As String) As String
   Mes_Año = Format$(Month(Fecha), "00") & "/" & Format$(Year(Fecha), "0000")
End Function

Public Sub ImprimirAdoAT(Datas As Adodc, _
                         Optional EsCampoCorto As Boolean)
Dim FormaImp As Byte
Dim SizeLetra As Integer
Dim TipoLetra As String
On Error GoTo Errorhandler
FormaImp = 2
Mensajes = "Seguro de Imprimir en:" & vbCrLf & Printer.DeviceName & "?"
Titulo = "IMPRESION"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
RatonReloj
InicioX = 0.5: InicioY = 0
SizeLetra = 6
'TipoCondensed
'TipoCalibri
'TipoCondensed
'TipoArialNarrow
TipoLetra = TipoVerdana
DataAnchoCampos InicioX, Datas, SizeLetra, TipoLetra, Orientacion_Pagina, EsCampoCorto
''For I = 2 To CantCampos
''    Ancho(I) = Ancho(I) - 3
''Next I
''Ancho(CantCampos) = AnchoPapel - 1
Pagina = 1
'Iniciamos la impresion
Printer.FontBold = False
With Datas.Recordset
If .RecordCount > 0 Then
    .MoveFirst
     Printer.FontSize = SizeLetra
     EncabezadoData Datas
     PosLinea = PosLinea - 0.1
     Imprimir_Linea_H PosLinea, Ancho(0), LimiteAncho
     PosLinea = PosLinea + 0.05
     Printer.FontSize = SizeLetra
     Printer.FontName = TipoLetra
     Do While Not .EOF
        Printer.FontBold = False
        If .fields("Ln") = 9999 Then
            Imprimir_Linea_H PosLinea - 0.05, Ancho(0), LimiteAncho
            PosLinea = PosLinea + 0.05
            Printer.FontBold = True
        End If
        PrinterAllFields CantCampos, PosLinea, Datas, True, False
        PosLinea = PosLinea + 0.36
        Imprimir_Linea_H PosLinea, Ancho(0), LimiteAncho
        PosLinea = PosLinea + 0.05
        If PosLinea >= LimiteAlto Then
           Imprimir_Linea_H PosLinea, Ancho(0), LimiteAncho
           Printer.NewPage
           EncabezadoData Datas
           PosLinea = PosLinea - 0.1
           Imprimir_Linea_H PosLinea, Ancho(0), LimiteAncho
           PosLinea = PosLinea + 0.05
           Printer.FontSize = SizeLetra
           Printer.FontName = TipoLetra
        End If
       .MoveNext
     Loop
End If
End With
Imprimir_Linea_H PosLinea, InicioX, LimiteAncho, Negro, True
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

Public Sub Modificacion_AT(Tipo_AT As String, _
                           CModificaciones As ComboBox, _
                           ATMes As String, _
                           ATAño As String)
Dim AdoListAT As ADODB.Recordset
  Fecha_Del_AT ATMes, ATAño
 'Despliego los datos para modificar
  CModificaciones.Clear
  sSQL = "SELECT T.*,C.CI_RUC,C.Cliente,C.TD "
  Select Case Tipo_AT
    Case "C": sSQL = sSQL & "FROM Trans_Compras As T, Clientes As C "
    Case "I": sSQL = sSQL & "FROM Trans_Importaciones As T, Clientes As C "
    Case "E": sSQL = sSQL & "FROM Trans_Exportaciones As T, Clientes As C "
    Case "V": sSQL = sSQL & "FROM Trans_Ventas As T, Clientes As C "
  End Select
  sSQL = sSQL _
       & "WHERE T.Fecha Between #" & FechaIni & "# AND #" & FechaFin & "# " _
       & "AND T.Item = '" & NumEmpresa & "' " _
       & "AND T.Periodo = '" & Periodo_Contable & "' "
  Select Case Tipo_AT
    Case "C": sSQL = sSQL & "AND C.Codigo = T.IdProv "
    Case "I": sSQL = sSQL & "AND C.Codigo = T.IdFiscalProv "
    Case "E": sSQL = sSQL & "AND C.Codigo = T.IdFiscalProv "
    Case "V": sSQL = sSQL & "AND C.Codigo = T.IdProv "
  End Select
  sSQL = sSQL & "ORDER BY T.Linea_SRI,C.CI_RUC "
  Select_AdoDB AdoListAT, sSQL
  With AdoListAT
   If .RecordCount > 0 Then
       Do While Not .EOF
          CModificaciones.AddItem Format$(.fields("Linea_SRI"), "000") & " " & .fields("Cliente")
         .MoveNext
       Loop
       CModificaciones.Text = CModificaciones.List(0)
   Else
       MsgBox "No existe datos de " & ATMes & " para modificar"
   End If
  End With
End Sub

Public Sub Eliminar_Trans_AT(Tipo_AT As String, _
                             CodigoC_AT As String, _
                             Establecimiento As String, _
                             PuntoEmision As String, _
                             Secuencial As String, _
                             Autorizacion_No As String, _
                             Correlativo_AT As String, _
                             Verificador_AT As String, _
                             Tipo_Comp_V As String, _
                             CMes_AT As String, _
                             CAño_AT As String, _
                             Linea_del_SRI As Integer)
  Fecha_Del_AT CMes_AT, CAño_AT
  If Linea_del_SRI > 0 Then
     sSQL = "DELETE * "
     Select Case Tipo_AT
       Case "C": sSQL = sSQL & "FROM Trans_Compras "
       Case "V": sSQL = sSQL & "FROM Trans_Ventas "
       Case "I": sSQL = sSQL & "FROM Trans_Importaciones "
       Case "E": sSQL = sSQL & "FROM Trans_Exportaciones "
     End Select
     sSQL = sSQL & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "AND Fecha Between #" & FechaIni & "# AND #" & FechaFin & "# " _
          & "AND Linea_SRI = " & Linea_del_SRI & " "
     Ejecutar_SQL_SP sSQL
  Else
     sSQL = "DELETE * "
     Select Case Tipo_AT
       Case "C": sSQL = sSQL & "FROM Trans_Compras " _
                      & "WHERE IdProv = '" & CodigoC_AT & "' " _
                      & "AND Establecimiento = '" & Establecimiento & "' " _
                      & "AND PuntoEmision = '" & PuntoEmision & "' " _
                      & "AND Secuencial = " & CTNumero(Secuencial) & " " _
                      & "AND Autorizacion = '" & Autorizacion_No & "' "
       Case "V": sSQL = sSQL & "FROM Trans_Ventas " _
                      & "WHERE IdProv = '" & CodigoC_AT & "' " _
                      & "AND Establecimiento = '" & Establecimiento & "' " _
                      & "AND PuntoEmision = '" & PuntoEmision & "' " _
                      & "AND Fecha Between #" & FechaIni & "# AND #" & FechaFin & "# " _
                      & "AND TipoComprobante = " & CTNumero(Tipo_Comp_V) & " " _
                      & "AND Autorizacion = '" & Autorizacion_No & "' "
       Case "E": sSQL = sSQL & "FROM Trans_Exportaciones " _
                      & "WHERE IdFiscalProv = '" & CodigoC_AT & "' " _
                      & "AND Establecimiento = '" & Establecimiento & "' " _
                      & "AND PuntoEmision = '" & PuntoEmision & "' " _
                      & "AND Secuencial = " & CTNumero(Secuencial) & " " _
                      & "AND Autorizacion = '" & Autorizacion_No & "' " _
                      & "AND Correlativo = '" & Correlativo_AT & "' " _
                      & "AND Verificador = '" & Verificador_AT & "' "
       Case "I": sSQL = sSQL & "FROM Trans_Importaciones " _
                      & "WHERE IdFiscalProv = '" & CodigoC_AT & "' " _
                      & "AND Correlativo = '" & Correlativo_AT & "' " _
                      & "AND Verificador = '" & Verificador_AT & "' "
     End Select
     sSQL = sSQL & "AND Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' "
     Ejecutar_SQL_SP sSQL
  End If
End Sub

Public Sub Eliminar_Trans_Air(Tipo_AT As String, _
                              CodigoC_AT As String, _
                              CMes_AT As String, _
                              CAño_AT As String, _
                              Linea_del_SRI As Integer)
Dim Establecimiento As String
Dim Emision As String
Dim Secuencial As Long
Dim Autorizacion_No As String

Dim DBAT_Air As ADODB.Recordset
Set DBAT_Air = New ADODB.Recordset
  DBAT_Air.CursorType = adOpenStatic
  DBAT_Air.CursorLocation = adUseClient
  If Linea_del_SRI > 0 Then
     Fecha_Del_AT CMes_AT, CAño_AT
    'Si existe la misma retencion la borramos para quede la actual
     sSQL = "DELETE * " _
          & "FROM Trans_Air " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "AND Fecha Between #" & FechaIni & "# AND #" & FechaFin & "# " _
          & "AND Tipo_Trans = '" & Tipo_AT & "' " _
          & "AND Linea_SRI = " & Linea_del_SRI & " "
     Ejecutar_SQL_SP sSQL
  Else
    'Si existe la misma retencion la borramos para quede la actual
     sSQL = "SELECT * " _
          & "FROM Asiento_Air " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND CodigoU = '" & CodigoUsuario & "' " _
          & "AND T_No = " & Trans_No & " " _
          & "AND Tipo_Trans = '" & Tipo_AT & "' " _
          & "ORDER BY A_No "
     DBAT_Air.open sSQL, AdoStrCnn, , , adCmdText
     With DBAT_Air
      If .RecordCount > 0 Then
          Do While Not .EOF
             Secuencial = .fields("SecRetencion")
             Autorizacion_No = .fields("AutRetencion")
             Establecimiento = .fields("EstabRetencion")
             Emision = .fields("PtoEmiRetencion")
             sSQL = "DELETE * " _
                  & "FROM Trans_Air " _
                  & "WHERE Item = '" & NumEmpresa & "' " _
                  & "AND Periodo = '" & Periodo_Contable & "' " _
                  & "AND IdProv = '" & CodigoC_AT & "' " _
                  & "AND Tipo_Trans = '" & Tipo_AT & "' " _
                  & "AND EstabRetencion = '" & Establecimiento & "' " _
                  & "AND PtoEmiRetencion = '" & Emision & "' " _
                  & "AND SecRetencion = " & Secuencial & " " _
                  & "AND AutRetencion = '" & Autorizacion_No & "' "
             Ejecutar_SQL_SP sSQL
            .MoveNext
          Loop
      End If
     End With
  End If
End Sub

Public Sub Eliminar_Asientos_AT(Tipo_AT As String)
  If Trans_No <= 0 Then Trans_No = 1
  SQL1 = "DELETE " _
       & "FROM Asiento_Air " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND T_No = " & Trans_No & " "
  Ejecutar_SQL_SP SQL1
  SQL1 = "DELETE " _
       & "FROM Asiento_Compras " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND T_No = " & Trans_No & " "
  Ejecutar_SQL_SP SQL1
  SQL1 = "DELETE " _
       & "FROM Asiento_Exportaciones " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND T_No = " & Trans_No & " "
  Ejecutar_SQL_SP SQL1
  SQL1 = "DELETE " _
       & "FROM Asiento_Importaciones " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND T_No = " & Trans_No & " "
  Ejecutar_SQL_SP SQL1
  SQL1 = "DELETE " _
       & "FROM Asiento_Ventas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND T_No = " & Trans_No & " "
  Ejecutar_SQL_SP SQL1
End Sub

Public Sub Imprimir_Catalogo_Ret(Datas As Adodc, _
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
InicioX = 1: InicioY = 0
DataAnchoCampos InicioX, Datas, SizeLetra, TipoArialNarrow, Orientacion_Pagina, EsCampoCorto
Ancho(0) = 1   ' Codigo
Ancho(1) = 2   ' Concpeto
Ancho(2) = 15.3   ' Porcentajes
Ancho(3) = 16.1   ' Ingresar Porc
Ancho(4) = 16.7  ' Fecha_Inicio
Ancho(5) = 18.2  ' Fecha_Final
Ancho(6) = 19.7  ' T
Ancho(7) = 20.3  ' Fin

Pagina = 1
'Iniciamos la impresion
Printer.FontBold = False
With Datas.Recordset
If .RecordCount > 0 Then
    .MoveFirst
     EncabezadoData Datas
     Printer.FontSize = SizeLetra
     Do While Not .EOF
        PrinterAllFields CantCampos, PosLinea, Datas, True, False
        PosLinea = PosLinea + 0.36
        If PosLinea >= (LimiteAlto - 0.36) Then
           PosLinea = PosLinea + 0.01
           Imprimir_Linea_H PosLinea, Ancho(0), Ancho(CantCampos)
           Printer.NewPage
           EncabezadoData Datas
           Printer.FontSize = SizeLetra
        End If
       .MoveNext
     Loop
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

Public Function Leer_Concepto_Retencion(Fecha As String, CodigoRet As String) As Concepto_Retencion_ATS
Dim CR_AT As Concepto_Retencion_ATS
Dim RegAdodc As ADODB.Recordset
Dim DatosSelect As String
  RatonReloj
  With CR_AT
      .Codigo = Ninguno
      .Concepto = Ninguno
      .Fecha_Final = FechaSistema
      .Fecha_Inicio = FechaSistema
      .Ingresar_Porcentaje = Ninguno
      .Porcentaje = 0
      .T = Ninguno
  End With
  If CodigoRet = "" Then CodigoRet = Ninguno
  DatosSelect = "SELECT * " _
              & "FROM Tipo_Concepto_Retencion " _
              & "WHERE Codigo = '" & CodigoRet & "' " _
              & "AND Fecha_Inicio <= #" & BuscarFecha(Fecha) & "# " _
              & "AND Fecha_Final >= #" & BuscarFecha(Fecha) & "# "
  DatosSelect = CompilarSQL(DatosSelect)
  Set RegAdodc = New ADODB.Recordset
  RegAdodc.CursorType = adOpenStatic
  RegAdodc.CursorLocation = adUseClient
  RegAdodc.open DatosSelect, AdoStrCnn, , , adCmdText
  If RegAdodc.RecordCount > 0 Then
     With CR_AT
         .Codigo = RegAdodc.fields("Codigo")
         .Concepto = RegAdodc.fields("Concepto")
         .Fecha_Final = RegAdodc.fields("Fecha_Final")
         .Fecha_Inicio = RegAdodc.fields("Fecha_Inicio")
         .Ingresar_Porcentaje = RegAdodc.fields("Ingresar_Porcentaje")
         .Porcentaje = RegAdodc.fields("Porcentaje")
         .T = RegAdodc.fields("T")
     End With
  End If
  RegAdodc.Close
  RatonNormal
  Leer_Concepto_Retencion = CR_AT
End Function

Public Sub Insertar_Campo_XML(CampoXML As String)
Dim CampoXMLAux As String
  If CampoXML <> "" Then
     CampoXMLAux = Sin_Signos_Especiales(CampoXML)
     TextoXML = TextoXML & CampoXMLAux & vbCrLf
  End If
End Sub

Public Sub SRI_Enviar_Mails(TFA As Tipo_Facturas, _
                            SRI_Autorizacion As Tipo_Estado_SRI, _
                            Tipo_Documento As String)
Dim RutaPDF As String
Dim RutaXML As String
Dim Email As String
Dim posPuntoComa As Byte
  'MsgBox MidStrg(TFA.ClaveAcceso, 9, 2)
   RatonReloj
  'MsgBox TFA.ClaveAcceso
   TMail.TipoDeEnvio = "CE"
   TMail.ListaMail = 255
   TMail.Destinatario = TFA.Cliente
   TMail.MensajeHTML = ""
   TMail.Adjunto = ""
   Select Case Tipo_Documento
     Case "FA"
          SRI_Generar_PDF_FA TFA, False
          SRI_Generar_XML_Firmado TFA.ClaveAcceso
     Case "NC"
          SRI_Generar_PDF_NC TFA, False
          SRI_Generar_XML_Firmado TFA.ClaveAcceso_LC
     Case "LC"
          SRI_Generar_PDF_LC TFA, False
          SRI_Generar_XML_Firmado TFA.ClaveAcceso_NC
     Case "GR"
          SRI_Generar_PDF_GR TFA, False
          SRI_Generar_XML_Firmado TFA.ClaveAcceso_GR
     Case "RE"
          SRI_Generar_PDF_RE TFA, False
          SRI_Generar_XML_Firmado TFA.ClaveAcceso
     Case "AB"
          RutaPDF = "ninguno.pdf"
          RutaXML = "ninguno.xml"
   End Select
   
   If Len(TFA.ClaveAcceso) >= 13 Or Len(TFA.ClaveAcceso_GR) >= 13 Or Len(TFA.ClaveAcceso_LC) >= 13 Or Len(TFA.ClaveAcceso_NC) >= 13 Then TMail.TipoDeEnvio = "CE"
       
   If Len(TFA.PDF_ClaveAcceso) > 1 Then
      RutaPDF = RutaSysBases & "\TEMP\" & TFA.PDF_ClaveAcceso & ".pdf"
      If InStr(TFA.PDF_ClaveAcceso, "_No_") = 0 Then RutaXML = RutaSysBases & "\TEMP\" & TFA.PDF_ClaveAcceso & ".xml"
   End If
   If MidStrg(TFA.ClaveAcceso, 9, 2) = "07" Then
      TMail.Mensaje = "Cliente: " & TFA.Cliente & vbCrLf _
                    & "Clave de Acceso: " & vbCrLf & TFA.ClaveAcceso & vbCrLf _
                    & "Hora de Generacion: " & SRI_Autorizacion.Hora_Autorizacion & vbCrLf _
                    & "Emision: " & TFA.Fecha & vbCrLf _
                    & "Vencimiento: " & SRI_Autorizacion.Fecha_Autorizacion & vbCrLf _
                    & "Autorizacion: " & vbCrLf & SRI_Autorizacion.Autorizacion & vbCrLf _
                    & "Retencion No. " & TFA.Serie_R & "-" & Format$(TFA.Retencion, "000000000") & vbCrLf _
                    & "Factura No. " & TFA.Serie & "-" & Format$(TFA.Factura, "000000000") & vbCrLf
      TMail.Asunto = TFA.Cliente & ", Retencion No. " & TFA.Serie_R & "-" & Format$(TFA.Retencion, "000000000")
   ElseIf MidStrg(TFA.ClaveAcceso, 9, 2) = "03" Then
      TMail.Mensaje = "Cliente: " & TFA.Cliente & vbCrLf _
                    & "Clave de Acceso: " & vbCrLf & TFA.ClaveAcceso_LC & vbCrLf _
                    & "Hora de Generacion: " & SRI_Autorizacion.Hora_Autorizacion & vbCrLf _
                    & "Emision: " & TFA.Fecha & vbCrLf _
                    & "Vencimiento: " & SRI_Autorizacion.Fecha_Autorizacion & vbCrLf _
                    & "Autorizacion: " & vbCrLf & SRI_Autorizacion.Autorizacion & vbCrLf _
                    & "Liquidacion de Compras No. " & TFA.Serie_LC & "-" & Format$(TFA.Factura, "000000000")
      TMail.Asunto = TFA.Cliente & ", Liquidacion de Compras No. " & TFA.Serie_LC & "-" & Format$(TFA.Factura, "000000000")
   Else
      TMail.Mensaje = "Cliente: " & TFA.Cliente & vbCrLf _
                    & "Clave de Acceso: " & vbCrLf & TFA.ClaveAcceso & vbCrLf _
                    & "Emision: " & TFA.Fecha & vbCrLf _
                    & "Vencimiento: " & TFA.Fecha_V & vbCrLf _
                    & "Fecha Autorizado: " & TFA.Fecha_Aut & vbCrLf _
                    & "Autorizacion: " & vbCrLf & TFA.Autorizacion & vbCrLf _
                    & "Factura No. " & TFA.Serie & "-" & Format$(TFA.Factura, "000000000") & vbCrLf
      If Tipo_Documento = "AB" Then
         TMail.Mensaje = TMail.Mensaje & vbCrLf & "SU PAGO FUE REGISTRADO CON EXITO" & vbCrLf & "EL " & FechaStrg(FA.Fecha_C) & vbCrLf & FA.Nota & vbCrLf
      Else
         TMail.Mensaje = TMail.Mensaje & "Hora de Generacion: " & TFA.Hora_FA & vbCrLf
      End If
      TMail.Asunto = TFA.Cliente & ", Factura No. " & TFA.Serie & "-" & Format$(TFA.Factura, "000000000")
   End If
  'Datos del destinatario de mails
   TMail.para = ""
   Insertar_Mail TMail.para, TFA.EmailC
   Insertar_Mail TMail.para, TFA.EmailC2
   Insertar_Mail TMail.para, TFA.EmailR
  
   TMail.Mensaje = TMail.Mensaje _
                 & String(45, "-") & vbCrLf _
                 & "Email(s) Destinatario(s):" & vbCrLf _
                 & Replace(TMail.para, ";", "; ") & vbCrLf _
                 & String(45, "_") & vbCrLf _
                 & NombreComercial & vbCrLf _
                 & RazonSocial & vbCrLf _
                 & Telefono1 & "/" & Telefono1 & vbCrLf _
                 & "Dir. " & Direccion & vbCrLf _
                 & UCaseStrg(NombreCiudad) & "-" & UCaseStrg(NombrePais) & vbCrLf
   'TMail.MensajeHTML = Leer_Archivo_Texto(RutaSistema & "\FONDOS\EmailRolPagos.html")
'''   = "<html>" _
'''                     & "<body>" _
'''                     & "<font face='calibri'><img src='https://erp.diskcoversystem.com/img/jpg/logo.jpg'><br><Br><font face='calibri'>" _
'''                     & "Thank you for contacting the Bank of America Service Desk. We're committed to providing seamless support in the moments that matter." _
'''                     & "<br><br>We heard your concerns with Skype for Business audio/video, and recommend using approved Skype for Business devices to resolve the issue." _
'''                     & "<br><br><h4><font color='red'>What do I need to do?</font></h4>" _
'''                     & "<div style='background-color: #FFF8DC;'>" _
'''                     & "1. Visit the <a href='http://u.go/pchk'>Skype for Business Peripheral Checker</a> & complete the form.<br>" _
'''                     & "<img src='https://erp.diskcoversystem.com/img/jpg/logo.jpg'><br>" _
'''                     & "4. Once approved, your new device(s) will be shipped to you. To get started, visit the <a href='http://u.go/tIxvB5'>Skype for Business page</a> and select <i>Setup your equipment</i> tab." _
'''                     & "</div>" _
'''                     & "<br><br>" _
'''                     & "<br>" _
'''                     & "If you still encounter Skype for Business audio/visual issues with your new device(s), please <a href='http://u.go/7I76vm'>submit a web ticket</a> and one of our expert Bank of America Service Desk employees will reach out to you." _
'''                     & "Thank you," _
'''                     & "<br>" _
'''                     & "Premium Service Desk" _
'''                     & "<br><Br>" _
'''                     & "<img src='https://erp.diskcoversystem.com/img/jpg/logo.jpg'>" _
'''                     & "</font>" _
'''                     & "</body>" _
'''                     & "</html>"
  'Enviamos lista de mails
  'MsgBox RutaPDF & "; " & RutaXML
   If Email_CE_Copia Then
      TMail.Credito_No = "X" & Format(TFA.Factura, "000000000")
      Insertar_Mail TMail.para, EmailProcesos
   End If
   If Existe_File(RutaPDF) Then TMail.Adjunto = RutaPDF & "; "
   If Existe_File(RutaXML) Then TMail.Adjunto = TMail.Adjunto & RutaXML
  'MsgBox TMail.para
   FEnviarCorreos.Show vbModal
   If Existe_File(RutaPDF) Then Kill RutaPDF
   If Existe_File(RutaXML) Then Kill RutaXML
   TMail.Volver_Envial = False
End Sub

Public Function SRI_Generar_XML(ClaveDeAcceso As String, _
                                EstadoDelSRI As String) As Tipo_Estado_SRI
Dim obj As New Cls_FirmarXML
Dim ObjEnviar As New WS_Recepcion
Dim ObjAutori As New WS_Autorizacion
Dim URLRecepcion As String
Dim URLAutorizacion As String
Dim Resultado As Boolean
Dim RutaCertificado As String
Dim ClaveCertificado As String
Dim RutaXML As String
Dim RutaXMLFirmado As String
Dim RutaXMLAutorizado As String
Dim RutaXMLRechazado As String
Dim MensajeError As String
Dim ArrayRecepcion() As String
Dim ArrayAutorizacion() As String
Dim Tiempo_Espera As Integer
Dim Tiempo_SRI As Integer
Dim EsperaEspera As Integer
Dim SRI_Aut As Tipo_Estado_SRI
Dim Intento_Enviar As Byte
Dim Intento_Autorizar As Byte

 RatonReloj
 FConexion.Show
 FConexion.TxtConexion = "CONECTANDO AL SRI..." & vbCrLf
 
 For Tiempo_Espera = 0 To 25
     FConexion.TxtConexion.Refresh
 Next Tiempo_Espera

 FConexion.TxtConexion = FConexion.TxtConexion & "Leer Certificado del Documento: " & MidStrg(ClaveDeAcceso, 25, 15) & vbCrLf
 FConexion.TxtConexion.Refresh
 Intento_Enviar = 0
 Intento_Autorizar = 0
 EsperaEspera = 6

 Ambiente = Leer_Campo_Empresa("Ambiente")
 ContEspec = Leer_Campo_Empresa("Codigo_Contribuyente_Especial")
 Obligado_Conta = Leer_Campo_Empresa("Obligado_Conta")
 
'Ruta del Certificado para Firmar el documento
 RutaCertificado = RutaSistema & "\CERTIFIC\" & Leer_Campo_Empresa("Ruta_Certificado")
 ClaveCertificado = Leer_Campo_Empresa("Clave_Certificado")
 
 FConexion.TxtConexion = FConexion.TxtConexion & "Conectandose al S.R.I." & vbCrLf
 FConexion.TxtConexion.Refresh
 
'Pagina de Conexion con el SRI
 URLRecepcion = Leer_Campo_Empresa("Web_SRI_Recepcion")
 URLAutorizacion = Leer_Campo_Empresa("Web_SRI_Autorizado")
 
 RutaXML = RutaDocumentos & "\Comprobantes Generados\" & ClaveDeAcceso & ".xml"
 RutaXMLFirmado = RutaDocumentos & "\Comprobantes Firmados\" & ClaveDeAcceso & ".xml"
 RutaXMLAutorizado = RutaDocumentos & "\Comprobantes Autorizados\" & ClaveDeAcceso & ".xml"
 RutaXMLRechazado = RutaDocumentos & "\Comprobantes no Autorizados\" & ClaveDeAcceso & ".xml"
 
 SRI_Aut = SRI_Leer_XML_Autorizado(RutaXMLAutorizado, RutaXMLRechazado)
 
With SRI_Aut
    'MsgBox .Estado_SRI & vbCrLf & .Error_SRI
     RatonReloj
     If .Estado_SRI = "OK" Then GoTo Fin_Autorizacion:
     FConexion.TxtConexion = FConexion.TxtConexion & "Determinando Carpetas de Conexion" & vbCrLf
     FConexion.TxtConexion.Refresh
    .Clave_De_Acceso = ClaveDeAcceso
    .Estado_SRI = EstadoDelSRI
    .Documento_XML = ""
    .Error_SRI = ""
     If Dir$(RutaXML) = "" Then
       .Estado_SRI = "CNG"
       .Error_SRI = "Error: Comprobante no generado"
        GoTo Fin_Autorizacion
     End If
     If Dir$(RutaXMLFirmado) = "" Then .Estado_SRI = "CNF"
     
    'MsgBox "Primer face: " & .Estado_SRI
     Select Case .Estado_SRI
       Case "CNA", "CNF": GoTo Volver_Firmar
       Case "ESC": GoTo Volver_Autorizar
       Case "CF", "CR", "ESI": GoTo Volver_Enviar
     End Select
Volver_Firmar:
    'Firmamos el documento
     RatonReloj
     FConexion.TxtConexion = FConexion.TxtConexion & "Firmando el Documento: " & MidStrg(ClaveDeAcceso, 25, 15) & vbCrLf
     FConexion.TxtConexion.Refresh
    'MsgBox "Firmar: " & RutaXML & vbCrLf & vbCrLf & RutaXMLFirmado & vbCrLf & vbCrLf & RutaCertificado & vbCrLf & vbCrLf & ClaveCertificado
     Resultado = obj.FirmarXML(RutaCertificado, ClaveCertificado, RutaXML, RutaXMLFirmado, MensajeError)
    'MsgBox "Firmar: " & MensajeError
     If Resultado Then
       .Estado_SRI = "CF"
       .Documento_XML = Leer_Archivo_Texto(RutaXMLFirmado)
Volver_Enviar:
        RatonReloj
        FConexion.TxtConexion = FConexion.TxtConexion & "Enviando el Documento al S.R.I.: " & MidStrg(ClaveDeAcceso, 25, 15) & vbCrLf
        FConexion.TxtConexion.Refresh
       'MsgBox URLRecepcion & vbCrLf & RutaXMLFirmado & vbCrLf & RutaXMLRechazado
        ArrayRecepcion = ObjEnviar.FF_EnviaXML_SRI(RutaXMLFirmado, URLRecepcion, RutaXMLRechazado)
        Intento_Enviar = Intento_Enviar + 1
       .Error_SRI = "Error al enviar: "
        For ContadorEstados = 0 To 4
            If Len(ArrayRecepcion(ContadorEstados)) > 1 Then .Error_SRI = .Error_SRI & ArrayRecepcion(ContadorEstados) & "; "
        Next ContadorEstados
        Sleep EsperaEspera
        
        If ArrayRecepcion(1) = "43" Or ArrayRecepcion(1) = "45" Then GoTo Volver_Obtener
        If ArrayRecepcion(0) = "RECIBIDA" Then
          .Estado_SRI = "CR"
          .Documento_XML = Leer_Archivo_Texto(RutaXMLFirmado)
Volver_Autorizar:
           RatonReloj
           FConexion.TxtConexion = FConexion.TxtConexion & "Cargando Documento Firmado: " & MidStrg(ClaveDeAcceso, 25, 15) & vbCrLf
           FConexion.TxtConexion.Refresh
          
Volver_Obtener:
          'Tiempo de Espera antes de averiguar al SRI de la autorizacion
           For Tiempo_Espera = 0 To 3
               Sleep EsperaEspera
               ArrayAutorizacion = ObjAutori.FF_ObtieneNumAutorizado(URLAutorizacion, .Clave_De_Acceso, RutaXMLAutorizado, RutaXMLRechazado)
               If ArrayAutorizacion(0) = "AUTORIZADO" Then Tiempo_Espera = 3
           Next Tiempo_Espera
           Intento_Autorizar = Intento_Autorizar + 1
           If ArrayAutorizacion(0) = "AUTORIZADO" Then
             'MsgBox "Ok Documento Firmado y Autorizado"
              FConexion.TxtConexion = FConexion.TxtConexion & "Extrayendo Documentos Autorizado: " & MidStrg(ClaveDeAcceso, 25, 15) & vbCrLf
              FConexion.TxtConexion.Refresh
              RatonReloj
             .Estado_SRI = "OK"
             .Error_SRI = "OK"
             .Autorizacion = ArrayAutorizacion(1)
             .Fecha_Autorizacion = Format$(MidStrg(ArrayAutorizacion(2), 1, 10), "dd/MM/yyyy")
             .Hora_Autorizacion = MidStrg(ArrayAutorizacion(2), 12, 8)
             .Documento_XML = Leer_Archivo_Texto(RutaXMLAutorizado)
              SRI_Actualizar_Documento_XML .Clave_De_Acceso
              FConexion.TxtConexion = FConexion.TxtConexion & "Grabando en la base el Documento: " & MidStrg(ClaveDeAcceso, 25, 15) & vbCrLf
              FConexion.TxtConexion.Refresh
           Else
              FConexion.TxtConexion = FConexion.TxtConexion & "Error: CNA" & vbCrLf
              FConexion.TxtConexion.Refresh
             .Estado_SRI = "CNA"
             .Error_SRI = "Error al Autorizar: "
              For ContadorEstados = 0 To 4
                  If Len(ArrayAutorizacion(ContadorEstados)) > 1 Then .Error_SRI = .Error_SRI & ArrayAutorizacion(ContadorEstados) & ", "
              Next ContadorEstados
              If Intento_Autorizar < 3 Then GoTo Volver_Autorizar
           End If
        ElseIf ArrayRecepcion(0) = "ERROR" Then
           FConexion.TxtConexion = FConexion.TxtConexion & "Error: ESI"
           FConexion.TxtConexion.Refresh
          .Estado_SRI = "ESI"
          .Error_SRI = " Error al enviar: "
           For ContadorEstados = 0 To 4
               If Len(ArrayRecepcion(ContadorEstados)) > 1 Then .Error_SRI = .Error_SRI & ArrayRecepcion(ContadorEstados) & ", "
           Next ContadorEstados
           If Intento_Enviar < 3 Then GoTo Volver_Enviar
        Else
           FConexion.TxtConexion = FConexion.TxtConexion & "Error: ESC" & vbCrLf
           FConexion.TxtConexion.Refresh
          .Estado_SRI = "ESC"
          .Error_SRI = "Error al enviar: "
           For ContadorEstados = 0 To 4
               If Len(ArrayRecepcion(ContadorEstados)) > 1 Then .Error_SRI = .Error_SRI & ArrayRecepcion(ContadorEstados) & "; "
           Next ContadorEstados
           If Intento_Enviar < 3 Then GoTo Volver_Enviar
        End If
     Else
        FConexion.TxtConexion = FConexion.TxtConexion & "Error: CNF" & vbCrLf
        FConexion.TxtConexion.Refresh
       .Estado_SRI = "CNF"
       .Error_SRI = MensajeError
       .Documento_XML = MensajeError
     End If
    .Error_SRI = TrimStrg(.Error_SRI)
End With
Fin_Autorizacion:
Progreso_Final
RatonNormal
Unload FConexion
SRI_Generar_XML = SRI_Aut     'Error_XML_SRI
End Function

Public Function SRI_Leer_XML_Autorizado(RutaAutorizado As String, RutaRechazado As String) As Tipo_Estado_SRI
Dim obj As New Cls_FirmarXML
Dim ObjEnviar As New WS_Recepcion
Dim ObjAutori As New WS_Autorizacion
Dim URLRecepcion As String
Dim URLAutorizacion As String
Dim Resultado As Boolean
Dim RutaCertificado As String
Dim ClaveCertificado As String
Dim MensajeError As String
Dim ArrayRecepcion() As String
Dim ArrayAutorizacion() As String
Dim Tiempo_Espera As Integer
Dim Tiempo_SRI As Integer
Dim EsperaEspera As Integer
Dim SRI_Aut As Tipo_Estado_SRI
Dim Intento_Enviar As Byte
Dim Intento_Autorizar As Byte
Dim IDo As Long
Dim IDn As Long

    RatonReloj
    Progreso_Barra.Mensaje_Box = "CONECTANDOSE AL S.R.I. ..."
    Progreso_Iniciar
    Progreso_Barra.Incremento = 0
    Progreso_Barra.Valor_Maximo = 100
    Progreso_Esperar True
    
    Intento_Enviar = 0
    Intento_Autorizar = 0

   'Pagina de Conexion con el SRI
   'ClaveDeAcceso = "0909202101179186152300120010040000056111234567815"
          
    'MsgBox URLAutorizacion & vbCrLf & ClaveDeAcceso & vbCrLf & MidStrg(ClaveDeAcceso, 24, 1)
          
    RatonReloj
    With SRI_Aut
         Progreso_Barra.Mensaje_Box = "Determinando Carpetas de Conexion"
         Progreso_Esperar True
         IDo = Len(RutaAutorizado)
         Do While MidStrg(RutaAutorizado, IDo, 1) <> "\"
            IDo = IDo - 1
         Loop
         IDn = InStr(RutaAutorizado, ".xml")
        .Clave_De_Acceso = MidStrg(RutaAutorizado, IDo + 1, IDn - IDo - 1)
        '.Clave_De_Acceso = "0103202407179130545000120030370000231671791305419"
         Select Case MidStrg(.Clave_De_Acceso, 24, 1)
           Case "1": URLAutorizacion = "https://celcer.sri.gob.ec/comprobantes-electronicos-ws/AutorizacionComprobantesOffline?wsdl"
           Case "2": URLAutorizacion = "https://cel.sri.gob.ec/comprobantes-electronicos-ws/AutorizacionComprobantesOffline?wsdl"
         End Select
        .Estado_SRI = "CG"
        .Documento_XML = ""
        .Error_SRI = ""
         EsperaEspera = 6
         RatonReloj
        'Tiempo de Espera antes de averiguar al SRI de la autorizacion
         For Tiempo_Espera = 0 To 3
             RatonReloj
             Sleep EsperaEspera
             ArrayAutorizacion = ObjAutori.FF_ObtieneNumAutorizado(URLAutorizacion, .Clave_De_Acceso, RutaAutorizado, RutaRechazado)
             Progreso_Barra.Mensaje_Box = ArrayAutorizacion(0)
             Progreso_Esperar True
             If ArrayAutorizacion(0) = "AUTORIZADO" Then Tiempo_Espera = 3
         Next Tiempo_Espera
         If ArrayAutorizacion(0) = "AUTORIZADO" Then
            Progreso_Barra.Mensaje_Box = "Extrayendo Documentos Autorizado: " & MidStrg(.Clave_De_Acceso, 25, 15)
            Progreso_Esperar True
            RatonReloj
           .Estado_SRI = "OK"
           .Error_SRI = "OK"
           .Autorizacion = ArrayAutorizacion(1)
           .Fecha_Autorizacion = Format$(MidStrg(ArrayAutorizacion(2), 1, 10), "dd/MM/yyyy")
           .Hora_Autorizacion = MidStrg(ArrayAutorizacion(2), 12, 8)
           .Documento_XML = Leer_Archivo_Texto(RutaAutorizado)
            
            'SRI_Actualizar_Documento_XML .Clave_De_Acceso
            'Progreso_Barra.Mensaje_Box = "Grabando en la base el Documento: " & MidStrg(ClaveDeAcceso, 25, 15)
            Cadena = ""
            For ContadorEstados = 0 To 3
                If Len(ArrayAutorizacion(ContadorEstados)) > 1 Then Cadena = Cadena & ArrayAutorizacion(ContadorEstados) & ", "
            Next ContadorEstados
           'MsgBox Cadena
            Progreso_Esperar True
         Else
           .Error_SRI = "Error al Autorizar: "
            For ContadorEstados = 0 To 4
                If Len(ArrayAutorizacion(ContadorEstados)) > 1 Then .Error_SRI = .Error_SRI & ArrayAutorizacion(ContadorEstados) & ", "
            Next ContadorEstados
           .Error_SRI = TrimStrg(.Error_SRI)
            'MsgBox .Error_SRI & "....."
         End If
         Progreso_Barra.Mensaje_Box = ArrayAutorizacion(0) & " " & .Estado_SRI & " -> " & ArrayAutorizacion(2)
         Progreso_Esperar True
         Progreso_Final
         RatonNormal
         
    End With
    Progreso_Final
    SRI_Leer_XML_Autorizado = SRI_Aut
End Function

Public Function SRI_Generar_Documento_PDF(NombreTipoDeLetra As String, _
                                          TFA As Tipo_Facturas, _
                                          VerDocumento As Boolean, _
                                          Tipo_Documento As String, _
                                          Optional AImpresora As Boolean) As Single
Dim DiasPago As String
Dim PDF_Titulo As String
Dim PosLinf As Single
Dim TPosLinea As Single
   'Datos Iniciales
    Ambiente = Leer_Campo_Empresa("Ambiente")
    Obligado_Conta = Leer_Campo_Empresa("Obligado_Conta")
    ContEspec = Leer_Campo_Empresa("Codigo_Contribuyente_Especial")
   'ContEspec = "12312"
    
   'Generacion Codigo de Barras
    If IsNumeric(TFA.Autorizacion_GR) Then Ambiente = MidStrg(TFA.ClaveAcceso_GR, 24, 1)
    
   'MsgBox RutaDocumentos
    Select Case Tipo_Documento
      Case "FA"
           PDF_Titulo = "Documento RIDE de Facturacion Electronica"
           TFA.PDF_ClaveAcceso = TFA.ClaveAcceso
           If TFA.Serie = "999999" Then TipoDoc = "COMPROBANTE No. " Else TipoDoc = "FACTURA No. "
           SerieFactura = TFA.Serie & "-" & Format$(TFA.Factura, "000000000")
           Autorizacion = TFA.Autorizacion
           MiHora = TFA.Fecha_Aut & " - " & TFA.Hora
      Case "NV"
           PDF_Titulo = "Documento RIDE de Nota de Venta"
           TFA.PDF_ClaveAcceso = TFA.ClaveAcceso
           TipoDoc = "NOTA DE VENTA No. "
           SerieFactura = TFA.Serie & "-" & Format$(TFA.Factura, "000000000")
           Autorizacion = TFA.Autorizacion
           MiHora = TFA.Fecha_Aut & " - " & TFA.Hora
      Case "OP"
           PDF_Titulo = "Documento RIDE de Orden de Produccion"
           TFA.PDF_ClaveAcceso = TFA.ClaveAcceso
           TipoDoc = "ORDEN DE PRODUCCION No. "
           SerieFactura = TFA.Serie & "-" & Format$(TFA.Factura, "000000000")
           Autorizacion = TFA.Autorizacion
           MiHora = TFA.Fecha & " - " & TFA.Hora
      Case "LC"
           PDF_Titulo = "Documento RIDE de Liquidacion de Compras Electronica"
           TFA.PDF_ClaveAcceso = TFA.ClaveAcceso_LC
           TipoDoc = "LIQUIDACION DE COMPRAS No. "
           SerieFactura = TFA.Serie_LC & "-" & Format$(TFA.Factura, "000000000")
           Autorizacion = TFA.Autorizacion_LC
           MiHora = TFA.Fecha_Aut & " - " & TFA.Hora
      Case "NC"
           PDF_Titulo = "Documento RIDE de Nota de Credito Electronica"
           TFA.PDF_ClaveAcceso = TFA.ClaveAcceso_NC
           TipoDoc = "NOTA DE CREDITO No. "
           SerieFactura = TFA.Serie_NC & "-" & Format$(TFA.Nota_Credito, "000000000")
           Autorizacion = TFA.Autorizacion_NC
           MiHora = TFA.Fecha_Aut_NC & " - " & TFA.Hora_NC
      Case "GR"
           PDF_Titulo = "Documento RIDE de Guia de Remision Electronica"
           TFA.PDF_ClaveAcceso = TFA.ClaveAcceso_GR
           TipoDoc = "GUIA DE REMISION No. "
           SerieFactura = TFA.Serie_GR & "-" & Format$(TFA.Remision, "000000000")
           Autorizacion = TFA.Autorizacion_GR
           MiHora = TFA.Fecha_Aut_GR & " - " & TFA.Hora_GR
      Case "RE"
           PDF_Titulo = "Documento RIDE de Retencion Electronica"
           TFA.PDF_ClaveAcceso = TFA.ClaveAcceso
           TipoDoc = "COMPROBANTE DE RETENCION No. "
           SerieFactura = TFA.Serie_R & "-" & Format$(TFA.Retencion, "000000000")
           Autorizacion = TFA.Autorizacion_R
           MiHora = TFA.Fecha_Aut & " - " & TFA.Hora
    End Select
    If TFA.PDF_ClaveAcceso = Ninguno Then TFA.PDF_ClaveAcceso = "Documento_" & Tipo_Documento & "_No_" & Autorizacion & "-" & SerieFactura
''    Progreso_Barra.Mensaje_Box = "Documento RIDE Electronico"
''    Progreso_Esperar

   'Generamos el documento
    If AImpresora Then
       tPrint.TipoImpresion = Es_Printer
    Else
       tPrint.TipoImpresion = Es_PDF
    End If
   'MsgBox TFA.PDF_ClaveAcceso
    tPrint.NombreArchivo = TFA.PDF_ClaveAcceso
    tPrint.TituloArchivo = PDF_Titulo
    tPrint.TipoLetra = NombreTipoDeLetra
    tPrint.OrientacionPagina = Orientacion_Pagina
    tPrint.PaginaA4 = True
    tPrint.EsCampoCorto = False
    tPrint.VerDocumento = VerDocumento
    
    Set cPrint = New cImpresion
    cPrint.iniciaImpresion
    cPrint.printCuadro 1, 1, 1.05, 1.05, Blanco, "B"
    cPrint.printImagen LogoTipo, 1.5, 1, 4.4, 2
    PosLinea = 3.1
    cPrint.tipoNegrilla = True
    cPrint.letraTipo NombreTipoDeLetra, 8
    TPosLinea = PosLinea
    PosLinea = cPrint.printTextoMultiple(1.6, PosLinea, UCaseStrg(RazonSocial), 8)
    If PosLinea > TPosLinea Then
       PosLinea = PosLinea + (PosLinea - TPosLinea)
    Else
       PosLinea = PosLinea + 0.3
    End If
   ' If cPrint.dNoLineas > 0 Then PosLinea = cPrint.dNoLineas
    'PosLinea = 4.2
    If UCaseStrg(RazonSocial) <> UCaseStrg(NombreComercial) Then
        cPrint.letraTipo NombreTipoDeLetra, 8
        TPosLinea = PosLinea
        PosLinea = cPrint.printTextoMultiple(1.6, PosLinea, UCaseStrg(NombreComercial), 8)
        If cPrint.dNoLineas > 0 Then PosLinea = cPrint.dNoLineas
    '   PosLinea = 5
        If PosLinea > TPosLinea Then
           PosLinea = PosLinea + (PosLinea - TPosLinea)
        Else
           PosLinea = PosLinea + 0.3
        End If
    End If
    TPosLinea = PosLinea
    PosLinf = PosLinea
    cPrint.letraTipo NombreTipoDeLetra, 7
    cPrint.printTexto 1.6, PosLinea, "Dirección Matríz:"
    PosLinea = PosLinea + 0.6
    cPrint.printTexto 1.6, PosLinea, "Dirección Sucursal:"
    PosLinea = PosLinea + 0.6
    cPrint.printTexto 1.6, PosLinea, "Obligado a llevar contabilidad:"
    PosLinea = PosLinea + 0.3
    If Len(ContEspec) > 1 Then cPrint.printTexto 1.6, PosLinea, "Contribuyente Especial No."
    PosLinea = PosLinea + 0.5
    
    cPrint.tipoNegrilla = False
    PosLinea = PosLinf + 0.3
    cPrint.letraTipo NombreTipoDeLetra, 7
    cPrint.printTexto 1.6, PosLinea, Direccion
    PosLinea = PosLinea + 0.6
    cPrint.printTexto 1.6, PosLinea, DireccionEstab
    PosLinea = PosLinea + 0.3
    cPrint.printTexto 5.3, PosLinea, Obligado_Conta
    PosLinea = PosLinea + 0.3
    If Len(ContEspec) > 1 Then cPrint.printTexto 5.3, PosLinea, ContEspec
    PosLinea = PosLinea + 0.35
    
   'AgenteRetencion = "NAC-DNCRASC20-00000001"
    
   'Linea donde se empezara a imprimir el resto del documento
    PosLinea = 1.1
    cPrint.tipoNegrilla = True
    cPrint.letraTipo NombreTipoDeLetra, 11
    cPrint.printTexto 10.1, PosLinea + 0.1, "R.U.C. " & RUC
    
    cPrint.letraTipo NombreTipoDeLetra, 7, &HC0&
    If MicroEmpresa <> Ninguno Then cPrint.printTexto 15, PosLinea, MicroEmpresa
    PosLinea = PosLinea + 0.4
    If Len(AgenteRetencion) > 1 Then cPrint.printTexto 10.1, PosLinea, "Agente de Retención " & AgenteRetencion
    PosLinea = PosLinea + 0.4
    cPrint.letraTipo NombreTipoDeLetra, 10
    cPrint.printTexto 10.1, PosLinea, TipoDoc
    cPrint.letraTipo NombreTipoDeLetra, 8
    PosLinea = PosLinea + 0.5
    cPrint.printTexto 10.1, PosLinea, "FECHA Y HORA DE AUTORIZACION: "
    PosLinea = PosLinea + 0.4
    cPrint.printTexto 10.1, PosLinea, "EMISIÓN: "
    PosLinea = PosLinea + 0.4
    cPrint.printTexto 10.1, PosLinea, "AMBIENTE: "
    PosLinea = PosLinea + 0.4
    If TFA.TC = "OP" Then
       cPrint.printTexto 10.1, PosLinea, "CODIGO DE BARRA DE LA ORDEN DE PRODUCCION "
    ElseIf TFA.Serie = "999999" Then
       cPrint.printTexto 10.1, PosLinea, "CODIGO DE BARRA DE LA ALICUOTA "
    Else
       cPrint.printTexto 10.1, PosLinea, "NUMERO DE AUTORIZACION Y CLAVE DE ACCESO "
    End If
    cPrint.tipoNegrilla = False
    
    PosLinea = 1.2
    PosLinea = PosLinea + 0.7
    cPrint.letraTipo NombreTipoDeLetra, 10, Rojo_Claro
    cPrint.printTexto 17.3, PosLinea, SerieFactura
    PosLinea = PosLinea + 0.5
    cPrint.letraTipo NombreTipoDeLetra, 8
    cPrint.printTexto 17.3, PosLinea, MiHora
    PosLinea = PosLinea + 0.4
    cPrint.printTexto 17.3, PosLinea, "NORMAL"
    PosLinea = PosLinea + 0.4
    If Ambiente = "1" Then
       cPrint.printTexto 17.3, PosLinea, "PRUEBA"
    Else
       cPrint.printTexto 17.3, PosLinea, "PRODUCCION"
    End If
    If Autorizacion <> TFA.PDF_ClaveAcceso Then
        If Tipo_Documento <> "OP" Then
           PosLinea = PosLinea + 0.8
           cPrint.printTexto 10.1, PosLinea, Autorizacion
           PosLinea = PosLinea + 0.45
        Else
           PosLinea = PosLinea + 1
        End If
    Else
       PosLinea = PosLinea + 1
    End If
    cPrint.generarBarras TFA.PDF_ClaveAcceso, cC128_A, 10.2, CDbl(PosLinea), 10, 1
    PosLinf = PosLinea + 0.9
    
   'Cuadros Superiores
    cPrint.printCuadro 1.5, 3, 9.4, PosLinf, Blanco, "B", 1
    cPrint.printCuadro 1.5, 3, 9.4, PosLinf, Negro, "B"
    cPrint.printCuadro 10, 1, 20, PosLinf, Negro, "B"
            
   'Empezamos a escribir los datos del beneficiario/Cliente/Proveedor
    PosLinea = PosLinf + 0.8
    PosLinf = PosLinf + 0.7
    
    cPrint.tipoNegrilla = True
    cPrint.letraTipo NombreTipoDeLetra, 7
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
    Select Case Tipo_Documento
      Case "FA": DiasPago = CStr(CFechaLong(TFA.Fecha_V) - CFechaLong(TFA.Fecha))
                 cPrint.printTexto 13, PosLinea, "Fecha Emisión: " & TFA.Fecha
                 If CFechaLong(TFA.Fecha) < CFechaLong(TFA.Fecha_V) Then cPrint.printTexto 17, PosLinea, "Fecha de pago: " & TFA.Fecha_V
                 PosLinea = PosLinea + 0.35
                 
                 If Len(TFA.Tipo_Pago) > 1 Then
                    cPrint.printTexto 1.6, PosLinea, UCaseStrg(TFA.Tipo_Pago_Det)
                    cPrint.printTexto 13, PosLinea, "Monto " & Moneda
                    cPrint.printVariable 14, PosLinea, TFA.Total_MN
                 End If
                 If DiasPago > 0 Then cPrint.printTexto 17, PosLinea, "Condición de Venta: " & DiasPago & " días"
                 If TFA.Cod_Ejec <> Ninguno Then
                    PosLinea = PosLinea + 0.35
                    cPrint.printTexto 1.6, PosLinea, "Vendedor: " & TFA.Cod_Ejec & " - " & ULCase(TFA.Ejecutivo_Venta)
                 End If
                 If TFA.Orden_Compra > 0 Then
                    If TFA.Cod_Ejec = Ninguno Then PosLinea = PosLinea + 0.35
                    cPrint.printTexto 13, PosLinea, "No. Orden de Compra: " & TFA.Orden_Compra
                 End If
                'Guia de Remision
                 If TFA.Remision > 0 Then
                    PosLinea = PosLinea + 0.35
                    cPrint.printLinea 1.5, PosLinea, 20, PosLinea, Negro
                    PosLinea = PosLinea + 0.05
                    cPrint.printTexto 1.6, PosLinea, "Guía Remisión: " & TFA.Serie_GR & "-" & Format(TFA.Remision, "000000000")
                    cPrint.printTexto 13, PosLinea, "Entrega: " & TFA.Comercial
                    PosLinea = PosLinea + 0.35
                    cPrint.printTexto 1.6, PosLinea, "Pedido: " & TFA.Pedido
                    cPrint.printTexto 6.5, PosLinea, "Zona: " & ULCase(TFA.CiudadGRF) & " - " & TFA.Zona
                    If Len(TFA.Lugar_Entrega) > 1 Then
                       cPrint.printTexto 13, PosLinea, "Lugar de Entrega: " & TFA.Lugar_Entrega
                    Else
                       cPrint.printTexto 13, PosLinea, "Lugar de Entrega: " & TFA.DireccionC
                    End If
                 End If
      Case "NC": PosLinea = PosLinea + 0.35
                 Cadena = "Comprobante que se Modifica, "
                 Select Case TFA.TC
                   Case "FA": Cadena = Cadena & "Factura No. "
                   Case "NV": Cadena = Cadena & "Nota de Venta No. "
                   Case Else: Cadena = Cadena & "Ticket No. "
                 End Select
                 Cadena = Cadena & TFA.Serie & "-" & Format$(TFA.Factura, String(9, "0"))
                 cPrint.printTexto 1.6, PosLinea, Cadena
                 cPrint.printTexto 17.2, PosLinea, "Fecha Emisión: " & TFA.Fecha_NC
                 PosLinea = PosLinea + 0.35
                 cPrint.printTexto 1.6, PosLinea, "Razón de la Modificación:"
                 cPrint.printTexto 14, PosLinea, "Fecha Emisión (Comprobante a Modificar): " & TFA.Fecha
                 PosLinea = PosLinea + 0.35
                 cPrint.tipoNegrilla = False
                 cPrint.printTexto 1.6, PosLinea, "- " & TFA.Nota
      Case "GR": cPrint.printTexto 16.8, PosLinea, "Motivo del Traslado: VENTAS"
                 PosLinea = PosLinea + 0.35
                 cPrint.printTexto 1.6, PosLinea, "Autorizacion: " & TFA.Autorizacion
                 Select Case TFA.TC
                   Case "FA": Cadena = "Factura No. "
                   Case "NV": Cadena = "Nota de Venta No. "
                   Case Else: Cadena = "Comprobante No. "
                 End Select
                 Cadena = Cadena & TFA.Serie & "-" & Format(TFA.Factura, "000000000")
                 cPrint.printTexto 11.5, PosLinea, Cadena
                 cPrint.printTexto 16.8, PosLinea, "Fecha Emisión: " & TFA.Fecha
                 PosLinea = PosLinea + 0.4
                 cPrint.printLinea 1.5, PosLinea, 20, PosLinea, Negro
                 PosLinea = PosLinea + 0.05
                 cPrint.printTexto 1.6, PosLinea, "Razón Social/Nombres y Apellidos(Transportista):"
                 cPrint.printTexto 18.5, PosLinea, "Identificación: "
                 PosLinea = PosLinea + 0.35
                 cPrint.printTexto 1.6, PosLinea, TFA.Comercial
                 cPrint.printTexto 18.5, PosLinea, TFA.CIRUCComercial
                 If cPrint.dNoLineas > 0 Then PosLinea = cPrint.dNoLineas
                 PosLinea = PosLinea + 0.35
                 cPrint.printTexto 1.6, PosLinea, "Punto de Partida:"
                 cPrint.printTexto 14.5, PosLinea, "Fecha inicio:"
                 cPrint.printTexto 16.8, PosLinea, "Fecha fin:"
                 cPrint.printTexto 18.5, PosLinea, "Placa:"
                 PosLinea = PosLinea + 0.35
                 cPrint.printTexto 1.6, PosLinea, TFA.Dir_PartidaGR
                 cPrint.printTexto 14.5, PosLinea, TFA.FechaGRI
                 cPrint.printTexto 16.8, PosLinea, TFA.FechaGRF
                 cPrint.printTexto 18.5, PosLinea, TFA.Placa_Vehiculo
                 PosLinea = PosLinea + 0.4
                 
                 cPrint.printLinea 1.5, PosLinea, 20, PosLinea, Negro
                 PosLinea = PosLinea + 0.05
                 cPrint.printTexto 1.6, PosLinea, "Razón Social/Nombres y Apellidos(Punto de Llegada):"
                 cPrint.printTexto 18.5, PosLinea, "Identificación: "
                 PosLinea = PosLinea + 0.35
                 cPrint.printTexto 18.5, PosLinea, TFA.CIRUCEntrega
                 cPrint.printTexto 1.6, PosLinea, TFA.Entrega
                 PosLinea = PosLinea + 0.35
                 Cadena = "Destino: De " & TFA.CiudadGRI & " a " & TFA.CiudadGRF
                 If Len(TFA.Lugar_Entrega) > 1 Then
                    Cadena = Cadena & ", " & TFA.Lugar_Entrega
                 Else
                    Cadena = Cadena & ", " & TFA.DireccionC
                 End If
                 cPrint.printTexto 1.6, PosLinea, Cadena
      Case "RE": PosLinea = PosLinea + 0.35
                 cPrint.tipoNegrilla = True
                 Cadena = "Documento Tipo " & TFA.Tipo_Comp & " No."
                 cPrint.printTexto 1.6, PosLinea, Cadena
                 cPrint.printTexto 16.5, PosLinea - 0.35, "Periodo Fiscal:  " & Format$(TFA.Fecha, "mm/yyyy")
                 cPrint.printTexto 16.5, PosLinea, "Fecha Emisión:"
                 cPrint.tipoNegrilla = False
                 
                 cPrint.printTexto 18.5, PosLinea, TFA.Fecha
                 cPrint.printTexto 2.5 + cPrint.anchoTexto(Cadena), PosLinea, TFA.Serie & "-" & Format$(TFA.Factura, String(9, "0"))
      Case "LC": cPrint.printTexto 17.2, PosLinea, "Fecha Emisión: " & TFA.Fecha
    End Select
    PosLinea = PosLinea - 0.2
   'Cuadro de Informacion Contribuyente
    cPrint.printCuadro 1.5, PosLinf, 20, PosLinea, Negro, "B"
    PosLinea = PosLinea + 0.6
   'Fin de Impresion del Encabezado del Documento PDF
    SRI_Generar_Documento_PDF = PosLinea
End Function

Public Sub SRI_Generar_PDF_FA(TFA As Tipo_Facturas, _
                              VerFactura As Boolean, _
                              Optional AImpresora As Boolean)
Dim AdoDBDet As ADODB.Recordset
Dim AdoDBAbo As ADODB.Recordset
Dim AdoDBProd As ADODB.Recordset

Dim tipoDeLetra As String
Dim Cod_Aux As String
Dim Cod_Bar As String
Dim Porc_Str As String
Dim TDec_PVP As Byte
Dim PorteDeLetra As Integer
Dim IVA_Porc As Currency
Dim TempPosLinea As Single
Dim PosLineaTemp As Single
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
    
    sSQL = "SELECT * " _
         & "FROM Trans_Abonos " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND TP = '" & TFA.TC & "' " _
         & "AND Serie = '" & TFA.Serie & "' " _
         & "AND Autorizacion = '" & TFA.Autorizacion & "' " _
         & "AND Factura = " & TFA.Factura & " " _
         & "ORDER BY Fecha,ID "
    Select_AdoDB AdoDBAbo, sSQL

   'Encabezado Detalle Factura
   'TipoArial / TipoVerdana / TipoHelvetica
    tipoDeLetra = TipoArial
    
    PosLinea = SRI_Generar_Documento_PDF(tipoDeLetra, TFA, VerFactura, TFA.TC, AImpresora)
    
    PosLinea = PosLinea + 0.1
    TempPosLinea = PosLinea
    PosLinea = PosLinea + 0.05
    cPrint.tipoNegrilla = True
    cPrint.letraTipo tipoDeLetra, 6
    If TFA.SP Then
       cPrint.printTexto 1.6, PosLinea, "Codigo Auxiliar", PorteDeLetra
       cPrint.printTexto 3.4, PosLinea, "Codigo Unitario", PorteDeLetra
    Else
       cPrint.printTexto 1.6, PosLinea, "Codigo Unitario", PorteDeLetra
       cPrint.printTexto 3.4, PosLinea, "Codigo Auxiliar", PorteDeLetra
    End If
    cPrint.printTexto 5.2, PosLinea, "Cantidad", PorteDeLetra
    cPrint.printTexto 6.4, PosLinea, "Cantidad", PorteDeLetra
    cPrint.letraTipo tipoDeLetra, 9
    cPrint.printTexto 7.6, PosLinea + 0.1, "D e s c r i p c i ó n"
    cPrint.letraTipo tipoDeLetra, 6
    cPrint.printTexto 14.4, PosLinea, "Lote No.", PorteDeLetra
    cPrint.printTexto 15.8, PosLinea, "Precio", PorteDeLetra
    cPrint.printTexto 17.2, PosLinea, "Valor", PorteDeLetra
    cPrint.printTexto 18.3, PosLinea, "Desc.", PorteDeLetra
    cPrint.printTexto 19.3, PosLinea, "Valor Total", PorteDeLetra
    PosLinea = PosLinea + 0.25
    cPrint.printTexto 5.2, PosLinea, "Total", PorteDeLetra
    cPrint.printTexto 6.4, PosLinea, "Bonif.", PorteDeLetra
    cPrint.printTexto 14.4, PosLinea, "/Orden No", PorteDeLetra
    cPrint.printTexto 15.8, PosLinea, "Unitario", PorteDeLetra
    cPrint.printTexto 17.2, PosLinea, "Descuent", PorteDeLetra
    cPrint.printTexto 18.4, PosLinea, "%", PorteDeLetra
    
   'Detalle de la Factura
    cPrint.printLinea 1.5, TempPosLinea + 0.6, 20, TempPosLinea + 0.6, Negro
    PosLinea = PosLinea + 0.45 ' 10.4
    cPrint.tipoNegrilla = False
    cPrint.letraTipo tipoDeLetra, 6
    cPrint.colorDeLetra = Negro
    
   'Detalle de la Factura
    With AdoDBDet
     If .RecordCount > 0 Then
         Progreso_Barra.Valor_Maximo = Progreso_Barra.Valor_Maximo + .RecordCount
         Do While Not .EOF
            Progreso_Barra.Mensaje_Box = "Generando documento PDF"
            Progreso_Esperar
            
            If Len(.fields("Codigo_Barra")) > 1 Then Cod_Bar = .fields("Codigo_Barra") Else Cod_Bar = .fields("Cod_Barras")
            Cod_Aux = .fields("Desc_Item")
            Total_Desc = .fields("Total_Desc") + .fields("Total_Desc2")
            If Total_Desc > 0 And .fields("Total") <> 0 Then Porc_Str = Format(Total_Desc / .fields("Total"), "00%") Else Porc_Str = ""

            If TFA.EsPorReembolso Then
               cPrint.printTexto 1.55, PosLinea, .fields("Codigo"), PorteDeLetra
               cPrint.printTexto 3.4, PosLinea, TrimStrg(MidStrg(.fields("Ruta"), 1, 10)), PorteDeLetra
            Else
                If TFA.SP Then
                   If Len(Cod_Bar) > 1 Then cPrint.printTexto 1.55, PosLinea, Cod_Bar, PorteDeLetra
                   If Len(Cod_Aux) > 1 Then
                      cPrint.printTexto 3.4, PosLinea, Cod_Aux, PorteDeLetra
                   Else
                      cPrint.printTexto 3.4, PosLinea, .fields("Codigo"), PorteDeLetra
                   End If
                Else
                   If Len(Cod_Aux) > 1 Then
                      cPrint.printTexto 1.55, PosLinea, Cod_Aux, PorteDeLetra
                   Else
                      cPrint.printTexto 1.55, PosLinea, .fields("Codigo"), PorteDeLetra
                   End If
                   If Len(Cod_Bar) > 1 Then cPrint.printTexto 3.4, PosLinea, Cod_Bar, PorteDeLetra
                End If
            End If
            cPrint.printFields 4.45, PosLinea, .fields("Cantidad"), PorteDeLetra
            
            Producto = .fields("Producto")
            If .fields("Codigo") <> "99.41" And TFA.Imp_Mes Then Producto = Producto & " " & .fields("Ticket") & " " & .fields("Mes")
            
            PosLineaTemp = cPrint.printTextoMultiple(7.6, PosLinea, Producto, 6.5)
            If PosLineaTemp > PosLinea Then PosLinea = PosLineaTemp 'Else PosLinea = PosLinea + 0.35
            If TFA.SP Then
               PosLinea = PosLinea + 0.32
               If CFechaLong(.fields("Fecha_Fab")) <> CFechaLong(.fields("Fecha_Exp")) Then
                  Producto = "ELAB. " & .fields("Fecha_Fab") & ", VENC. " & .fields("Fecha_Exp") & " " & vbCrLf
                  cPrint.printTexto 7.6, PosLinea, Producto, PorteDeLetra
               End If
               If Len(.fields("Reg_Sanitario")) > 1 Then
                  PosLinea = PosLinea + 0.32
                  Producto = "Reg. Sanit. " & .fields("Reg_Sanitario")
                  cPrint.printTexto 7.6, PosLinea, Producto, PorteDeLetra
               End If
               If Len(.fields("Modelo")) > 1 Then
                  PosLinea = PosLinea + 0.32
                  Producto = "Modelo: " & .fields("Modelo")
                  cPrint.printTexto 7.6, PosLinea, Producto, PorteDeLetra
               End If
               If Len(.fields("Procedencia")) > 1 Then
                  PosLinea = PosLinea + 0.32
                  Producto = "Procedencia: " & .fields("Procedencia")
                  cPrint.printTexto 7.6, PosLinea, Producto, PorteDeLetra
               End If
            End If
            If Len(.fields("Serie_No")) > 1 Then
               PosLinea = PosLinea + 0.32
               Producto = "Serie No. " & .fields("Serie_No")
               cPrint.printTexto 7.6, PosLinea, Producto, PorteDeLetra
            End If
            If TFA.EsPorReembolso Then
               PosLinea = PosLinea + 0.32
               Cod_Aux = "Autorizacion(" & .fields("Lote_No") & ") " & .fields("Procedencia")
               cPrint.printTexto 7.6, PosLinea, Cod_Aux, PorteDeLetra
               PosLinea = PosLinea + 0.32
               Cod_Aux = "Reembolso de Gastos"
               If Len(.fields("Tipo_Hab")) > 1 Then Cod_Aux = Cod_Aux & " por " & .fields("Tipo_Hab")
               cPrint.printTexto 7.6, PosLinea, Cod_Aux, PorteDeLetra
               
            End If
            If .fields("Orden_No") <> 0 Then cPrint.printTexto 14.3, PosLinea, Format$(.fields("Orden_No"), "00000000"), PorteDeLetra
            
            If Len(.fields("Lote_No")) > 1 Then Cadena = .fields("Lote_No") Else Cadena = TrimStrg(MidStrg(.fields("Ruta"), 1, 13))
            If Len(Cadena) > 1 Then cPrint.printTexto 14.25, PosLinea, Cadena, PorteDeLetra
            
            TDec_PVP = Dec_PVP
            If TDec_PVP > 6 Then TDec_PVP = 6
            cPrint.printFields 15.55, PosLinea, .fields("Precio"), PorteDeLetra, , , TDec_PVP
            cPrint.printFields 18.65, PosLinea, .fields("Total"), PorteDeLetra, , , 2

            If Len(.fields("Tipo_Hab")) > 1 Then
               If Total_Desc > 0 Then
                  PosLinea = PosLinea + 0.32
                  cPrint.printTexto 7.6, PosLinea, .fields("Tipo_Hab"), PorteDeLetra
                  cPrint.printVariable 18.6, PosLinea, Total_Desc, PorteDeLetra, , , 2
               End If
            Else
               If Total_Desc > 0 Then cPrint.printVariable 16.2, PosLinea, Total_Desc, PorteDeLetra, , , 2
               If Len(Porc_Str) > 1 Then cPrint.printTexto 18.35, PosLinea, Porc_Str, PorteDeLetra, PorteDeLetra
            End If
            
            If PosLineaTemp > PosLinea Then PosLinea = PosLineaTemp
            PosLinea = PosLinea + 0.32
            
           'MsgBox "Linea.: " & PosLinea
            
            If PosLinea >= 26 Then
              'MsgBox "Nueva Pag.: " & PosLinea
              'Lineas al final del detalle de la factura
               PosLinea = PosLinea + 0.2
               cPrint.printCuadro 1.5, TempPosLinea, 20, PosLinea + 0.05, Negro, "B"
               TempPosLinea = TempPosLinea - 0.15
               cPrint.printLinea 2.9, TempPosLinea, 2.9, PosLinea - 0.1, Negro
               cPrint.printLinea 4.8, TempPosLinea, 4.8, PosLinea - 0.1, Negro
               cPrint.printLinea 6, TempPosLinea, 6, PosLinea - 0.1, Negro
               cPrint.printLinea 7.2, TempPosLinea, 7.2, PosLinea - 0.1, Negro
               cPrint.printLinea 14, TempPosLinea, 14, PosLinea - 0.1, Negro
               'cPrint.printCuadroLinea 14.3, TempPosLinea, 14.3, PosLinea - 0.1, Negro
               cPrint.printLinea 15.5, TempPosLinea, 15.5, PosLinea - 0.1, Negro
               cPrint.printLinea 17.1, TempPosLinea, 17.1, PosLinea - 0.1, Negro
               cPrint.printLinea 18.2, TempPosLinea, 18.2, PosLinea - 0.1, Negro
               cPrint.printLinea 18.9, TempPosLinea, 18.9, PosLinea - 0.1, Negro
               
               cPrint.printCuadro 1.5, PosLinea + 0.05, 20, PosLinea + 0.6, Negro, "B"
               cPrint.printTexto 1.6, PosLinea, "C O N T I N U A   E N   L A   S I G U I E N T E   P A G I N A . . .", PorteDeLetra
               cPrint.paginaNueva
               PosLinea = 2
               TempPosLinea = PosLinea
               PosLinea = PosLinea - 0.25
               cPrint.tipoNegrilla = True
               cPrint.letraTipo tipoDeLetra, 6
               If TFA.SP Then
                  cPrint.printTexto 1.6, PosLinea, "Codigo Auxiliar", PorteDeLetra
                  cPrint.printTexto 3.4, PosLinea, "Codigo Unitario", PorteDeLetra
               Else
                  cPrint.printTexto 1.6, PosLinea, "Codigo Unitario", PorteDeLetra
                  cPrint.printTexto 3.4, PosLinea, "Codigo Auxiliar", PorteDeLetra
               End If
               cPrint.printTexto 5.3, PosLinea, "Cantidad", PorteDeLetra
               cPrint.printTexto 6.5, PosLinea, "Cantidad", PorteDeLetra
               cPrint.letraTipo tipoDeLetra, 9
               cPrint.printTexto 7.7, PosLinea + 0.1, "D e s c r i p c i ó n"
               cPrint.letraTipo tipoDeLetra, 6
               cPrint.printTexto 13.2, PosLinea, "Lote", PorteDeLetra
               'cPrint.printTexto 14.8, PosLinea, "Fecha", porteDeLetra
               cPrint.printTexto 16, PosLinea, "Precio Unitario", PorteDeLetra
               cPrint.printTexto 17.6, PosLinea, "Valor", PorteDeLetra
               cPrint.printTexto 18.7, PosLinea, "Desc.", PorteDeLetra
               cPrint.printTexto 19.5, PosLinea, "Valor Total", PorteDeLetra
               PosLinea = PosLinea + 0.25
               cPrint.printTexto 5.3, PosLinea, "Total", PorteDeLetra
               cPrint.printTexto 6.5, PosLinea, "Bonif.", PorteDeLetra
              'cPrint.printTexto 14.8, PosLinea, "Vencimient", porteDeLetra
               cPrint.printTexto 17.6, PosLinea, "Descuent", PorteDeLetra
               cPrint.printTexto 18.8, PosLinea, "%", PorteDeLetra
              'Detalle de la Factura
               cPrint.printCuadro 1.5, TempPosLinea, 20, TempPosLinea + 0.5, Negro, "B"
               PosLinea = PosLinea + 0.45 ' 10.4
               cPrint.tipoNegrilla = False
               cPrint.letraTipo tipoDeLetra, 6
               cPrint.colorDeLetra = Negro
            End If
           .MoveNext
         Loop
     End If
    End With
    
   'Lineas al final del detalle de la factura
    cPrint.printCuadro 1.5, TempPosLinea, 20, PosLinea - 0.6, Negro, "B"
    cPrint.printLinea 3.3, TempPosLinea, 3.3, PosLinea, Negro
    cPrint.printLinea 5.1, TempPosLinea, 5.1, PosLinea, Negro
    cPrint.printLinea 6.3, TempPosLinea, 6.3, PosLinea, Negro
    cPrint.printLinea 7.5, TempPosLinea, 7.5, PosLinea, Negro
    cPrint.printLinea 14.2, TempPosLinea, 14.2, PosLinea, Negro
    cPrint.printLinea 15.7, TempPosLinea, 15.7, PosLinea, Negro
    cPrint.printLinea 17.1, TempPosLinea, 17.1, PosLinea, Negro
    cPrint.printLinea 18.2, TempPosLinea, 18.2, PosLinea, Negro
    cPrint.printLinea 18.9, TempPosLinea, 18.9, PosLinea, Negro
    PosLinea = PosLinea + 0.1
    TempPosLinea = PosLinea
   'Lineas del Pie de la Factura
''    cPrint.letraTipo tipoDeLetra, 8, Rojo
''    cPrint.printTexto 1.5, 8.85, "1.5 x " & TempPosLinea & ", Alto: " & PosLinea
    cPrint.PorteDeLetra = 6
    For I = 1 To 10
        PosLinea = PosLinea + 0.35
        cPrint.printLinea 14.2, PosLinea, 20, PosLinea, Negro
       'Raya de Informe y Leyenda
        If I = 1 Then cPrint.printLinea 1.5, PosLinea, 13.9, PosLinea, Negro
        If I = 8 Then cPrint.printLinea 1.5, PosLinea, 13.9, PosLinea, Negro
    Next I
    PosLinea = PosLinea + 0.45
    cPrint.printCuadro 1.5, TempPosLinea, 13.6, PosLinea - 0.7, Negro, "B"
    cPrint.printCuadro 14.2, TempPosLinea, 20, PosLinea - 0.7, Negro, "B"
   'Datos del Pie de la Factura
    PosLinea = TempPosLinea
    With TFA
     If .Si_Existe_Doc Then
         PosLinea = TempPosLinea + 0.05
         PosColumna = 14.3
         cPrint.colorDeLetra = Negro
         cPrint.tipoNegrilla = True
         cPrint.letraTipo tipoDeLetra, 6
         cPrint.printTexto 1.6, PosLinea, "INFORMACION ADICIONAL:"
         cPrint.printTexto PosColumna, PosLinea, "SUBTOTAL " & .Porc_IVA_S & "%:"
         PosLinea = PosLinea + 0.35
         cPrint.printTexto PosColumna, PosLinea, "SUBTOTAL 0%:"
         PosLinea = PosLinea + 0.35
         cPrint.printTexto PosColumna, PosLinea, "TOTAL DESCUENTO:", PorteDeLetra
         PosLinea = PosLinea + 0.35
         cPrint.printTexto PosColumna, PosLinea, "SUBTOTAL NO OBJETO DE IVA:"
         PosLinea = PosLinea + 0.35
         cPrint.printTexto PosColumna, PosLinea, "SUBTOTAL EXENTO DE IVA:", PorteDeLetra
         PosLinea = PosLinea + 0.35
         cPrint.printTexto PosColumna, PosLinea, "SUBTOTAL SIN IMPUESTOS:", PorteDeLetra
         PosLinea = PosLinea + 0.35
         cPrint.printTexto PosColumna, PosLinea, "ICE:", PorteDeLetra
         PosLinea = PosLinea + 0.35
         cPrint.printTexto PosColumna, PosLinea, "IVA " & .Porc_IVA_S & "%:", PorteDeLetra
         PosLinea = PosLinea + 0.35
         cPrint.printTexto PosColumna, PosLinea, "IVA 0%:", PorteDeLetra
         PosLinea = PosLinea + 0.35
         cPrint.printTexto PosColumna, PosLinea, "PROPINA:", PorteDeLetra
         PosLinea = PosLinea + 0.35
         cPrint.printTexto PosColumna, PosLinea, "VALOR TOTAL:", PorteDeLetra
         PosLinea = PosLinea + 0.35
         cPrint.tipoNegrilla = False
         cPrint.letraTipo tipoDeLetra, 6
         PosLinea = TempPosLinea + 0.05
         PosColumna = 18.55
         cPrint.printVariable PosColumna, PosLinea, .Con_IVA, PorteDeLetra, , , 2
         PosLinea = PosLinea + 0.35
         cPrint.printVariable PosColumna, PosLinea, .Sin_IVA, PorteDeLetra, , , 2
         PosLinea = PosLinea + 0.35
         cPrint.printVariable PosColumna, PosLinea, .Total_Descuento, PorteDeLetra, , , 2
         PosLinea = PosLinea + 0.35
         cPrint.printVariable PosColumna, PosLinea, .Total_Sin_No_IVA, PorteDeLetra, , , 2
         PosLinea = PosLinea + 0.35
         cPrint.printVariable PosColumna, PosLinea, .Total_Sin_No_IVA, PorteDeLetra, , , 2
         PosLinea = PosLinea + 0.35
         cPrint.printVariable PosColumna, PosLinea, .SubTotal - .Total_Descuento, PorteDeLetra, , , 2
         PosLinea = PosLinea + 0.35
         cPrint.printVariable PosColumna, PosLinea, .Total_Sin_No_IVA, PorteDeLetra, , , 2
         PosLinea = PosLinea + 0.35
         cPrint.printVariable PosColumna, PosLinea, .Total_IVA, PorteDeLetra, , , 2
         PosLinea = PosLinea + 0.35
         cPrint.printVariable PosColumna, PosLinea, .Total_Sin_No_IVA, PorteDeLetra, , , 2
         PosLinea = PosLinea + 0.35
         cPrint.printVariable PosColumna, PosLinea, .Servicio, PorteDeLetra, , , 2
         PosLinea = PosLinea + 0.35
         cPrint.printVariable PosColumna, PosLinea, .Total_MN, PorteDeLetra, , , 2
         PosLinea = PosLinea + 0.4
         TempPosLineaAbono = PosLinea
        'Forma de pago en el pie de la factura
         If AdoDBAbo.RecordCount > 0 Then
            cPrint.printLinea 8.4, TempPosLinea, 8.4, PosLinea - 1.2, Negro
            cPrint.printLinea 9.7, TempPosLinea, 9.7, PosLinea - 1.2, Negro
            cPrint.printLinea 12.4, TempPosLinea, 12.4, PosLinea - 1.2, Negro
            PosLinea = TempPosLinea + 0.05
            cPrint.tipoNegrilla = True
            cPrint.letraTipo tipoDeLetra, 6
            cPrint.printTexto 8.5, PosLinea, "Fecha", PorteDeLetra
            cPrint.printTexto 9.8, PosLinea, "Detalle del Pago", PorteDeLetra
            cPrint.printTexto 12.5, PosLinea, "Monto Abono", PorteDeLetra
            PosLinea = PosLinea + 0.4
            cPrint.tipoNegrilla = False
            cPrint.letraTipo tipoDeLetra, 6
            Do While Not AdoDBAbo.EOF
               cPrint.printFields 8.4, PosLinea, AdoDBAbo.fields("Fecha"), PorteDeLetra
               cPrint.printTexto 9.8, PosLinea, ULCase(AdoDBAbo.fields("Banco")), PorteDeLetra
               cPrint.printFields 12.25, PosLinea, AdoDBAbo.fields("Abono"), PorteDeLetra, , , 2
               PosLinea = PosLinea + 0.3
               AdoDBAbo.MoveNext
            Loop
            PosLinea = TempPosLineaAbono
         End If
         PosLinea = TempPosLinea + 0.4
         PosColumna = 1.6
         If TFA.Razon_Social <> TFA.Cliente Then
            cPrint.printTexto PosColumna, PosLinea, "Beneficiario: " & TFA.Cliente, PorteDeLetra
            PosLinea = PosLinea + 0.3
            cPrint.printTexto PosColumna, PosLinea, "Codigo: " & TFA.CI_RUC, PorteDeLetra
            PosLinea = PosLinea + 0.3
         End If
         If .Grupo <> Ninguno And .Curso <> .DireccionC And .Imp_Mes Then
            cPrint.printTexto PosColumna, PosLinea, "Grupo: " & .Grupo & "-" & .Curso, PorteDeLetra
            PosLinea = PosLinea + 0.3
         End If
         If Len(.EmailC) > 1 And InStr(TFA.EmailC, "@") > 0 Then
            cPrint.printTexto PosColumna, PosLinea, "Email: " & .EmailC, PorteDeLetra
            PosLinea = PosLinea + 0.3
         End If
         If Len(TFA.EmailR) > 1 And InStr(TFA.EmailR, "@") > 0 And InStr(TFA.EmailC, TFA.EmailR) = 0 Then
            cPrint.printTexto PosColumna, PosLinea, "Email2: " & .EmailR, PorteDeLetra
            PosLinea = PosLinea + 0.3
         End If
         If Len(.Contacto) > 1 Then
            cPrint.printTexto PosColumna, PosLinea, "Referencia: " & .Contacto, PorteDeLetra
            PosLinea = PosLinea + 0.3
         End If
         If Len(.TelefonoC) > 1 Then
            cPrint.printTexto PosColumna, PosLinea, "Teléfono: " & .TelefonoC, PorteDeLetra
            PosLinea = PosLinea + 0.3
         End If
         If Len(.Nota) > 1 Then
            cPrint.printTexto PosColumna, PosLinea, "Nota: " & .Nota, PorteDeLetra
            PosLinea = PosLinea + 0.3
         End If
         If Len(.Observacion) > 1 Then
            cPrint.letraTipo tipoDeLetra, 6
            Cadena = TrimStrg("Observacion: " & .Observacion)
            PosLinea = cPrint.printTextoMultiple(1.6, PosLinea, Cadena, 12.5)
'            cPrint.printTexto PosColumna, PosLinea, "Observacion: " & .Observacion, porteDeLetra
            PosLinea = TempPosLinea
         End If
        'Cuadro inferior Derecha
     End If
    End With
    cPrint.letraTipo tipoDeLetra, 6
    If Informativo_FA <> Ninguno Then
       PosLinea = TempPosLineaAbono - 1.1
       Cadena = TrimStrg(Informativo_FA)
       PosLinea = cPrint.printTextoMultiple(1.6, PosLinea, Cadena, 12.3)
    End If
    If Debo_Pagare <> Ninguno Then
       PosLinea = TempPosLineaAbono
       Cadena = TrimStrg(Debo_Pagare)
       PosLinea = cPrint.printTextoMultiple(1.6, PosLinea, Cadena, 18.6)
    End If
   'cPrint.printLinea 3.1, 5.5, 20, 5.5, Rojo, 2
    AdoDBDet.Close
    AdoDBAbo.Close
    RatonNormal
    cPrint.finalizaImpresion
End Sub

Public Sub SRI_Generar_PDF_NC(TFA As Tipo_Facturas, _
                              VerNotaCredito As Boolean)
Dim AdoDBFac As ADODB.Recordset
Dim AdoDBDet As ADODB.Recordset
Dim AdoDBDetFA As ADODB.Recordset
Dim AdoDBDetKD As ADODB.Recordset
Dim ConsultarDetalle As Boolean
Dim TempPosLinea As Single
Dim PosLineaTemp As Single
Dim PosLinea1 As Single
Dim tipoDeLetra As String

    RatonReloj
   'Generacion Codigo de Barras
    ConsultarDetalle = False
    TFA.Total_IVA_NC = 0
    TFA.SubTotal_NC = 0
    Total_Sin_IVA = 0
    Total_Con_IVA = 0
    Total_Desc = 0
    SubTotal_NC = 0
    
    Leer_Datos_FA_NV TFA
    
'''    sSQL = "SELECT F.Razon_Social,F.RUC_CI,F.Autorizacion,F.CodigoC,F.Fecha,F.Nota," _
'''         & "C.Cliente,C.CI_RUC,C.Telefono,C.TelefonoT,C.Direccion,C.DireccionT," _
'''         & "C.Grupo,C.Codigo,C.Ciudad,C.Email,C.CI_RUC_R,C.TD,C.DirNumero " _
'''         & "FROM Facturas As F, Clientes As C " _
'''         & "WHERE F.Item = '" & NumEmpresa & "' " _
'''         & "AND F.Periodo = '" & Periodo_Contable & "' " _
'''         & "AND F.TC = '" & TFA.TC & "' " _
'''         & "AND F.Serie = '" & TFA.Serie & "' " _
'''         & "AND F.Autorizacion = '" & TFA.Autorizacion & "' " _
'''         & "AND F.Factura = " & TFA.Factura & " " _
'''         & "AND C.Codigo = F.CodigoC "
'''    Select_AdoDB AdoDBFac, sSQL
'''    With AdoDBFac
'''     If .RecordCount > 0 Then
'''         TFA.Fecha = .Fields("Fecha")
'''         TFA.Grupo = .Fields("Grupo")
'''         TFA.Autorizacion = .Fields("Autorizacion")
'''         TFA.CodigoC = .Fields("CodigoC")
'''         TFA.DireccionC = .Fields("Direccion")
'''         TFA.EmailC = .Fields("Email")
'''         TFA.TelefonoC = .Fields("Telefono")
'''         TFA.Nota = .Fields("Nota")
'''         If Len(.Fields("Razon_Social")) > 1 Then TFA.Cliente = .Fields("Razon_Social") Else TFA.Cliente = .Fields("Cliente")
'''         If Len(.Fields("RUC_CI")) > 1 Then TFA.CI_RUC = .Fields("RUC_CI") Else TFA.CI_RUC = .Fields("CI_RUC")
'''         If Len(.Fields("Razon_Social")) > 1 And Len(.Fields("RUC_CI")) > 1 Then
'''            sSQL = "SELECT Codigo,Grupo_No,Representante,Cedula_R,Lugar_Trabajo_R,Telefono_R " _
'''                 & "FROM Clientes_Matriculas " _
'''                 & "WHERE Item = '" & NumEmpresa & "' " _
'''                 & "AND Periodo = '" & Periodo_Contable & "' " _
'''                 & "AND Codigo = '" & TFA.CodigoC & "' "
'''            Select_AdoDB AdoDBDet, sSQL
'''            If AdoDBDet.RecordCount > 0 Then
'''               TFA.Curso = TFA.DireccionC
'''               TFA.DireccionC = AdoDBDet.Fields("Lugar_Trabajo_R")
'''            End If
'''            AdoDBDet.Close
'''         End If
'''         Validar_Porc_IVA TFA.Fecha
'''         ConsultarDetalle = True
'''     End If
'''    End With
'''    AdoDBFac.Close
    
         Validar_Porc_IVA TFA.Fecha
         ConsultarDetalle = True
         
    sSQL = "SELECT * " _
         & "FROM Trans_Abonos " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND TP = '" & TFA.TC & "' " _
         & "AND Serie = '" & TFA.Serie & "' " _
         & "AND Autorizacion = '" & TFA.Autorizacion & "' " _
         & "AND Factura = " & TFA.Factura & " " _
         & "AND Serie_NC = '" & TFA.Serie_NC & "' " _
         & "AND Secuencial_NC = " & TFA.Nota_Credito & " " _
         & "AND Banco = 'NOTA DE CREDITO' " _
         & "ORDER BY Cheque DESC "
    
    Select_AdoDB AdoDBDet, sSQL
    With AdoDBDet
     If .RecordCount > 0 Then
         TFA.Fecha_NC = .fields("Fecha")
         TFA.Serie_NC = .fields("Serie_NC")
         TFA.ClaveAcceso_NC = .fields("Clave_Acceso_NC")
         TFA.Autorizacion_NC = .fields("Autorizacion_NC")
         TFA.Nota_Credito = .fields("Secuencial_NC")
         Do While Not .EOF
            If .fields("Cheque") = "I.V.A." Then
                TFA.Total_IVA_NC = TFA.Total_IVA_NC + .fields("Abono")
            Else
                TFA.SubTotal_NC = TFA.SubTotal_NC + .fields("Abono")
            End If
           .MoveNext
         Loop
        .MoveFirst
     End If
    End With
    
    sSQL = "SELECT Autorizacion, Codigo_Inv, Producto, Cantidad, Precio, Total, Total_IVA, Descuento, Cta_Devolucion, CodBodega, Porc_IVA, Mes, Mes_No , Anio, ID " _
         & "FROM Detalle_Nota_Credito " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND Serie = '" & TFA.Serie_NC & "' " _
         & "AND Secuencial = " & TFA.Nota_Credito & " " _
         & "ORDER BY ID "
    Select_AdoDB AdoDBDetFA, sSQL
   'Encabezado Factura
    With AdoDBDet
     If .RecordCount > 0 Then
         tipoDeLetra = TipoHelvetica
         PosLinea = SRI_Generar_Documento_PDF(tipoDeLetra, TFA, VerNotaCredito, "NC")
         PosLinea = PosLinea + 0.1
         TempPosLinea = PosLinea
         
        'Encabezado del detalle de la Nota de Credito
         PosLinea = PosLinea + 0.15
         cPrint.letraTipo tipoDeLetra, 7
         cPrint.printTexto 1.6, PosLinea, "Codigo"
         cPrint.printTexto 4.4, PosLinea, "D e s c r i p c i ó n"
         cPrint.printTexto 12.2, PosLinea, "Cantidad"
         cPrint.printTexto 13.9, PosLinea, "Precio Unitario"
         cPrint.printTexto 16.1, PosLinea, "Descuento"
         cPrint.printTexto 18.2, PosLinea, "Valor Total"
         PosLinea = PosLinea + 0.35
         cPrint.printLinea 1.5, PosLinea, 20, PosLinea, Negro
         PosLinea = PosLinea + 0.1
        'Detalle Nota de Credito
         cPrint.letraTipo tipoDeLetra, 7
         cPrint.colorDeLetra = Negro
         If AdoDBDetFA.RecordCount > 0 Then
            TFA.SubTotal_NC = 0
            Progreso_Barra.Valor_Maximo = Progreso_Barra.Valor_Maximo + AdoDBDetFA.RecordCount
            Do While Not AdoDBDetFA.EOF
               Progreso_Barra.Mensaje_Box = "Generando documento PDF"
               Progreso_Esperar
               cPrint.printTexto 1.6, PosLinea, AdoDBDetFA.fields("Codigo_Inv")
               
               PosLineaTemp = cPrint.printTextoMultiple(4.3, PosLinea, AdoDBDetFA.fields("Producto"), 6.5)
               If PosLineaTemp > PosLinea Then PosLinea = PosLineaTemp
               'If cPrint.dNoLineas > 0 Then PosLinea = cPrint.dNoLineas
               cPrint.printVariable 11.4, PosLinea, AdoDBDetFA.fields("Cantidad")
               cPrint.printVariable 13.9, PosLinea, AdoDBDetFA.fields("Precio"), , , , 4
               cPrint.printFields 15.9, PosLinea, AdoDBDetFA.fields("Descuento"), , , , 2
               cPrint.printFields 18.3, PosLinea, AdoDBDetFA.fields("Total"), , , , 2
               
               TFA.SubTotal_NC = TFA.SubTotal_NC + AdoDBDetFA.fields("Total")
               Total_Desc = Total_Desc + AdoDBDetFA.fields("Descuento")
               If AdoDBDetFA.fields("Total_IVA") > 0 Then
                  Total_Con_IVA = Total_Con_IVA + AdoDBDetFA.fields("Total")
               Else
                  Total_Sin_IVA = Total_Sin_IVA + AdoDBDetFA.fields("Total")
               End If
               PosLinea = PosLinea + 0.4
               AdoDBDetFA.MoveNext
            Loop
         End If
         cPrint.printCuadro 1.5, TempPosLinea, 20, PosLinea - 0.6, Negro, "B"
         cPrint.printLinea 4.2, TempPosLinea, 4.2, PosLinea, Negro
         cPrint.printLinea 12, TempPosLinea, 12, PosLinea, Negro
         cPrint.printLinea 13.7, TempPosLinea, 13.7, PosLinea, Negro
         cPrint.printLinea 15.9, TempPosLinea, 15.9, PosLinea, Negro
         cPrint.printLinea 18.1, TempPosLinea, 18.1, PosLinea, Negro
         TempPosLinea = PosLinea + 0.1
         
        'Informacion adicional
         cPrint.printCuadro 1.5, TempPosLinea, 13.1, TempPosLinea + 1.8, Negro, "B"
         cPrint.printLinea 1.5, TempPosLinea + 0.4, 13.1, TempPosLinea + 0.4, Negro

        'Pie de la Nota de Credito
         PosLinea = PosLinea + 0.6
         If TFA.Grupo <> Ninguno And TFA.Curso <> Ninguno Then
            cPrint.printTexto 1.6, PosLinea, "Grupo: " & TFA.Grupo & "-" & TFA.Curso
            PosLinea = PosLinea + 0.35
         End If
         If Len(TFA.DireccionC) > 1 Then
            cPrint.printTexto 1.6, PosLinea, "Dirección: " & TFA.DireccionC
            PosLinea = PosLinea + 0.35
         End If
         If Len(TFA.TelefonoC) > 1 Then
            cPrint.printTexto 1.6, PosLinea, "Teléfono: " & TFA.TelefonoC
            PosLinea = PosLinea + 0.35
         End If
         If Len(TFA.EmailC) > 1 Then
            cPrint.printTexto 1.6, PosLinea, "Email: " & TFA.EmailC
            PosLinea = PosLinea + 0.35
         End If
         PosLinea = TempPosLinea + 0.1
         cPrint.tipoNegrilla = True
         cPrint.letraTipo tipoDeLetra, 7
         PosColumna = 13.8
         cPrint.printTexto 1.6, PosLinea, "INFORMACION ADICIONAL:"
         
         cPrint.printTexto PosColumna, PosLinea, "SUBTOTAL 0%:"
         PosLinea = PosLinea + 0.4
         cPrint.printTexto PosColumna, PosLinea, "SUBTOTAL " & Porc_IVA * 100 & "%:"
         PosLinea = PosLinea + 0.4
         cPrint.printTexto PosColumna, PosLinea, "TOTAL DESCUENTO:"
         PosLinea = PosLinea + 0.4
         cPrint.printTexto PosColumna, PosLinea, "SUBTOTAL SIN IMPUESTOS:"
         PosLinea = PosLinea + 0.4
         cPrint.printTexto PosColumna, PosLinea, "TOTAL I.V.A. " & Porc_IVA * 100 & "%:"
         PosLinea = PosLinea + 0.4
         cPrint.printTexto PosColumna, PosLinea, "V A L O R   T O T A L:"
         cPrint.tipoNegrilla = False
         cPrint.letraTipo tipoDeLetra, 7
         SubTotal_NC = TFA.SubTotal_NC - Total_Desc + TFA.Total_IVA_NC
         
         PosLinea = TempPosLinea + 0.1
         PosColumna = 18.2
         cPrint.printVariable PosColumna, PosLinea, Total_Sin_IVA, , , , 2
         PosLinea = PosLinea + 0.4
         cPrint.printVariable PosColumna, PosLinea, Total_Con_IVA, , , , 2
         PosLinea = PosLinea + 0.4
         cPrint.printVariable PosColumna, PosLinea, Total_Desc, , , , 2
         PosLinea = PosLinea + 0.4
         cPrint.printVariable PosColumna, PosLinea, TFA.SubTotal_NC - Total_Desc, , , , 2
         PosLinea = PosLinea + 0.4
         cPrint.printVariable PosColumna, PosLinea, TFA.Total_IVA_NC, , , , 2
         PosLinea = PosLinea + 0.4
         cPrint.printVariable PosColumna, PosLinea, SubTotal_NC, , , , 2
         PosLinea = PosLinea + 0.4
        'Rayas verticales de los subtotales
        'pie de la factura
         PosLinea = TempPosLinea + 0.4
         For I = 1 To 5
             cPrint.printLinea 13.7, PosLinea, 20, PosLinea, Negro
             PosLinea = PosLinea + 0.4
         Next I
         PosLinea = PosLinea + 0.1
         cPrint.printCuadro 13.7, TempPosLinea, 20, PosLinea - 0.7, Negro, "B"
         cPrint.finalizaImpresion
         RatonNormal
     Else
         RatonNormal
         MsgBox "Este documentos no tiene Nota de Credito que presentar"
     End If
    End With
    AdoDBDet.Close
    AdoDBDetFA.Close
End Sub

Public Sub SRI_Generar_PDF_GR(TFA As Tipo_Facturas, _
                              VerFactura As Boolean, _
                              Optional AImpresora As Boolean)
Dim AdoDBDet As ADODB.Recordset
Dim tipoDeLetra As String
Dim IVA_Porc As Currency
Dim Porc_Str As String
Dim PorteDeLetra As Integer
Dim TempPosLinea As Single
Dim TempPosLineaAbono As Single
    RatonReloj
   'Detalle de descuentos
    sSQL = "SELECT DF.*, CP.Reg_Sanitario " _
         & "FROM Detalle_Factura As DF, Catalogo_Productos AS CP " _
         & "WHERE DF.Item = '" & NumEmpresa & "' " _
         & "AND DF.Periodo = '" & Periodo_Contable & "' " _
         & "AND DF.TC = '" & TFA.TC & "' " _
         & "AND DF.Serie = '" & TFA.Serie & "' " _
         & "AND DF.Autorizacion = '" & TFA.Autorizacion & "' " _
         & "AND DF.Factura = " & TFA.Factura & " " _
         & "AND LEN(DF.Autorizacion) >= 13 " _
         & "AND DF.T <> 'A' " _
         & "AND DF.Item = CP.Item " _
         & "AND DF.Periodo = CP.Periodo " _
         & "AND DF.Codigo = CP.Codigo_Inv " _
         & "ORDER BY DF.ID "
    Select_AdoDB AdoDBDet, sSQL
    RatonReloj
          
   'Encabezado Detalle Factura
   'TipoArial / TipoVerdana
    tipoDeLetra = TipoHelvetica
    PosLinea = SRI_Generar_Documento_PDF(tipoDeLetra, TFA, VerFactura, "GR", AImpresora)
    PosLinea = PosLinea + 0.1
    TempPosLinea = PosLinea
    PosLinea = PosLinea + 0.05
    cPrint.tipoNegrilla = True
    cPrint.letraTipo tipoDeLetra, 6
    cPrint.printTexto 1.6, PosLinea, "Codigo Unitario", PorteDeLetra
    cPrint.printTexto 3.4, PosLinea, "Codigo Auxiliar", PorteDeLetra
    cPrint.printTexto 5.3, PosLinea, "Cantidad", PorteDeLetra
    cPrint.printTexto 6.5, PosLinea, "D E S C R I P C I O N    D E L   P R O D U C T O"
   'Detalle de la Factura
    cPrint.printLinea 1.5, TempPosLinea + 0.4, 20, TempPosLinea + 0.4, Negro
    PosLinea = PosLinea + 0.5 ' 10.4
    cPrint.tipoNegrilla = False
    cPrint.letraTipo tipoDeLetra, 6
    cPrint.colorDeLetra = Negro
    With AdoDBDet
     If .RecordCount > 0 Then
         TFA.Si_Existe_Doc = True
         Progreso_Barra.Valor_Maximo = Progreso_Barra.Valor_Maximo + .RecordCount
         Do While Not .EOF
            Progreso_Barra.Mensaje_Box = "Generando documento PDF"
            Progreso_Esperar
            Producto = .fields("Producto")
            If .fields("Ticket") <> Ninguno Then Producto = Producto & " " & .fields("Ticket")
            If .fields("Mes") <> Ninguno And TFA.Imp_Mes Then Producto = Producto & " " & .fields("Mes")
            If TFA.SP Then
               Producto = Producto & vbCrLf _
                        & "Lote No. " & .fields("Lote_No") & ", ELAB. " & .fields("Fecha_Fab") & ", VENC. " & .fields("Fecha_Exp") _
                        & ", Reg. Sanit. " & .fields("Reg_Sanitario") _
                        & ", Modelo: " & .fields("Modelo") & ", Serie No. " & .fields("Serie_No") _
                        & ", Procedencia: " & .fields("Procedencia")
            End If
            cPrint.printFields 1.65, PosLinea, .fields("Codigo")
            If Len(.fields("Codigo_Barra")) > 1 Then cPrint.printFields 3.4, PosLinea, .fields("Codigo_Barra")
            cPrint.printFields 4.5, PosLinea, .fields("Cantidad")
            PosLinea = cPrint.printTextoMultiple(6.5, PosLinea, Producto, 14)
            If cPrint.dNoLineas > 0 Then PosLinea = cPrint.dNoLineas Else PosLinea = PosLinea + 0.35
            If PosLinea > 27 Then
              'Lineas al final del detalle de la factura
               PosLinea = PosLinea + 0.1
               cPrint.printCuadro 1.5, TempPosLinea, 20, PosLinea + 0.05, Negro, "B"
               TempPosLinea = TempPosLinea - 0.15
               cPrint.printLinea 2.9, TempPosLinea, 2.9, PosLinea - 0.1, Negro
               cPrint.printLinea 4.8, TempPosLinea, 4.8, PosLinea - 0.1, Negro
               cPrint.printLinea 6, TempPosLinea, 6, PosLinea - 0.1, Negro
               PosLinea = PosLinea + 0.05
               cPrint.printCuadro 1.5, PosLinea + 0.05, 20, PosLinea + 0.6, Negro, "B"
               cPrint.printTexto 1.6, PosLinea, "C O N T I N U A   E N   L A   S I G U I E N T E   P A G I N A . . .", PorteDeLetra
               cPrint.paginaNueva
               PosLinea = 2
               TempPosLinea = PosLinea
               PosLinea = PosLinea - 0.25
               cPrint.tipoNegrilla = True
               cPrint.letraTipo tipoDeLetra, 6
               cPrint.printTexto 1.6, PosLinea, "Codigo Unitario", PorteDeLetra
               cPrint.printTexto 3.4, PosLinea, "Codigo Auxiliar", PorteDeLetra
               cPrint.printTexto 5.3, PosLinea, "Cantidad", PorteDeLetra
               cPrint.letraTipo tipoDeLetra, 9
               cPrint.printTexto 6.5, PosLinea + 0.1, "D e s c r i p c i ó n"
               cPrint.letraTipo tipoDeLetra, 6
              'Detalle de la Factura
               cPrint.printCuadro 1.5, TempPosLinea, 20, TempPosLinea + 0.4, Negro, "B"
               PosLinea = PosLinea + 0.5 ' 10.4
               cPrint.tipoNegrilla = False
               cPrint.letraTipo tipoDeLetra, 6
               cPrint.colorDeLetra = Negro
            End If
           .MoveNext
         Loop
     End If
    End With
   'Lineas al final del detalle de la factura
    cPrint.printCuadro 1.5, TempPosLinea, 20, PosLinea - 0.6, Negro, "B"
    cPrint.printLinea 3.3, TempPosLinea, 3.3, PosLinea, Negro
    cPrint.printLinea 5.2, TempPosLinea, 5.2, PosLinea, Negro
    cPrint.printLinea 6.4, TempPosLinea, 6.4, PosLinea, Negro
    PosLinea = PosLinea + 0.1
    TempPosLinea = PosLinea
   'Lineas del Pie de la Factura
    cPrint.PorteDeLetra = 6
   'Datos del Pie de la Guia de Remision
    With TFA
     If .Si_Existe_Doc Then
         PosColumna = 14.5
         cPrint.colorDeLetra = Negro
         cPrint.tipoNegrilla = True
         cPrint.letraTipo tipoDeLetra, 6
         cPrint.printTexto 1.6, PosLinea + 0.1, "INFORMACION ADICIONAL:"
         cPrint.tipoNegrilla = False
         cPrint.letraTipo tipoDeLetra, 6
         TempPosLineaAbono = PosLinea
        'Forma de pago en el pie de la factura
         PosLinea = TempPosLinea + 0.4
         PosColumna = 1.6
         If Len(.EmailC) > 1 And InStr(TFA.EmailC, "@") > 0 Then cPrint.printTexto PosColumna, PosLinea, "Email: " & .EmailC, PorteDeLetra
         If Len(TFA.EmailR) > 1 And InStr(TFA.EmailR, "@") > 0 And InStr(TFA.EmailC, TFA.EmailR) = 0 Then
            cPrint.printTexto PosColumna + 7, PosLinea, "Email2: " & .EmailR, PorteDeLetra
         End If
         If Len(.TelefonoC) > 1 Then cPrint.printTexto PosColumna + 16.5, PosLinea, "Teléfono: " & .TelefonoC, PorteDeLetra
        'Cuadro inferior
         PosLinea = PosLinea + 0.35
         cPrint.printCuadro 1.5, TempPosLinea, 20, PosLinea - 0.6, Negro, "B"
         cPrint.printLinea 1.5, TempPosLinea + 0.4, 20, TempPosLinea + 0.4, Negro
     End If
    End With
    AdoDBDet.Close
    RatonNormal
    cPrint.finalizaImpresion
End Sub

Public Sub SRI_Generar_PDF_RE(TFA As Tipo_Facturas, _
                              VerRetencion As Boolean)
Dim AdoDBCompras As ADODB.Recordset
Dim AdoDBAir As ADODB.Recordset
Dim ConsultarDetalle As Boolean
Dim TempPosLinea As Single
Dim EjercicioFiscal As String
Dim ConceptoRet As String
Dim NombreTipoDeLetra As String
    RatonReloj
    ConsultarDetalle = False
    NombreTipoDeLetra = TipoArialNarrow
   'Listar las Retenciones del IVA
    sSQL = "SELECT C.Cliente,C.CI_RUC,C.TD,C.Direccion,C.Email,C.Ciudad,C.DirNumero,C.Telefono,TC.* " _
         & "FROM Trans_Compras As TC,Clientes As C " _
         & "WHERE TC.Item = '" & NumEmpresa & "' " _
         & "AND TC.Periodo = '" & Periodo_Contable & "' " _
         & "AND TC.Numero = " & TFA.Numero & " " _
         & "AND TC.TP = '" & TFA.TP & "' " _
         & "AND TC.SecRetencion = " & TFA.Retencion & " " _
         & "AND TC.Serie_Retencion = '" & TFA.Serie_R & "' " _
         & "AND TC.IdProv = C.Codigo " _
         & "ORDER BY Cta_Servicio,Cta_Bienes "
    Select_AdoDB AdoDBCompras, sSQL
    With AdoDBCompras
     If .RecordCount > 0 Then
         TFA.Fecha = .fields("Fecha")
         TFA.Cliente = .fields("Cliente")
         TFA.Razon_Social = .fields("Cliente")
         TFA.CI_RUC = .fields("CI_RUC")
         TFA.RUC_CI = .fields("CI_RUC")
         TFA.DireccionC = .fields("Direccion")
        'TFA.Serie_R
        'TFA.Retencion
         TFA.Fecha_Aut = .fields("Fecha_Aut")
         TFA.Hora = .fields("Hora_Aut")
         TFA.Autorizacion_R = .fields("AutRetencion")
         TFA.ClaveAcceso = .fields("Clave_Acceso")
         TFA.Serie = .fields("Establecimiento") & .fields("PuntoEmision")
         TFA.Factura = .fields("Secuencial")
         TFA.Fecha = .fields("Fecha")
         TFA.Tipo_Comp = CStr(.fields("TipoComprobante"))
         FechaTexto = TFA.Fecha
         EjercicioFiscal = CStr(Year(TFA.Fecha))
         Validar_Porc_IVA TFA.Fecha
         ConsultarDetalle = True
     End If
    End With
    
   'Determinamos el Tipo de Comprobante
    sSQL = "SELECT Tipo_Comprobante_Codigo, Descripcion " _
         & "FROM Tipo_Comprobante " _
         & "WHERE TC = 'TDC' " _
         & "AND Tipo_Comprobante_Codigo = " & Val(TFA.Tipo_Comp) & " "
    Select_AdoDB AdoDBAir, sSQL
    If AdoDBAir.RecordCount > 0 Then TFA.Tipo_Comp = AdoDBAir.fields("Descripcion")
    AdoDBAir.Close
    
   'Listar las Retenciones de la Fuente
    sSQL = "SELECT TIV.Concepto,R.* " _
         & "FROM Trans_Air As R,Tipo_Concepto_Retencion As TIV " _
         & "WHERE R.Item = '" & NumEmpresa & "' " _
         & "AND R.Periodo = '" & Periodo_Contable & "' " _
         & "AND R.Numero = " & TFA.Numero & " " _
         & "AND R.TP = '" & TFA.TP & "' " _
         & "AND R.SecRetencion = " & TFA.Retencion & " " _
         & "AND R.EstabRetencion = '" & MidStrg(TFA.Serie_R, 1, 3) & "' " _
         & "AND R.PtoEmiRetencion = '" & MidStrg(TFA.Serie_R, 4, 3) & "' " _
         & "AND R.Tipo_Trans IN ('C','I') " _
         & "AND TIV.Fecha_Inicio <= #" & BuscarFecha(FechaTexto) & "# " _
         & "AND TIV.Fecha_Final >= #" & BuscarFecha(FechaTexto) & "# " _
         & "AND R.CodRet = TIV.Codigo " _
         & "ORDER BY R.Cta_Retencion "
    Select_AdoDB AdoDBAir, sSQL
   'Encabezado Factura
    With AdoDBCompras
     If .RecordCount > 0 Then
         PosLinea = SRI_Generar_Documento_PDF(NombreTipoDeLetra, TFA, VerRetencion, "RE")
         PosLinea = PosLinea + 0.1
         TempPosLinea = PosLinea
         PosLinea = PosLinea + 0.05
         cPrint.tipoNegrilla = True
         cPrint.letraTipo NombreTipoDeLetra, 7
         cPrint.printTexto 1.6, PosLinea, "Impuesto"
         cPrint.printTexto 2.9, PosLinea, "D e s c r i p c i ó n"
         cPrint.printTexto 13.7, PosLinea, "Codigo"
         cPrint.printTexto 15.1, PosLinea, "Base"
         cPrint.printTexto 17.1, PosLinea, "Porcentaje"
         cPrint.printTexto 18.6, PosLinea, "Valor"
         PosLinea = PosLinea + 0.3
         cPrint.printTexto 13.7, PosLinea, "Retencion"
         cPrint.printTexto 15.1, PosLinea, "Imponible"
         cPrint.printTexto 17.1, PosLinea, "Retenido"
         cPrint.printTexto 18.6, PosLinea, "Retenido"
         PosLinea = PosLinea + 0.35
         cPrint.tipoNegrilla = False
         cPrint.printLinea 1.5, PosLinea, 20, PosLinea, Negro
     End If
    End With
   'Detalle Factura
    cPrint.letraTipo NombreTipoDeLetra, 7
    Sumatoria = 0
    PosLinea = PosLinea + 0.1
    If ConsultarDetalle Then
       With AdoDBCompras
        If .RecordCount > 0 Then
            If .fields("ValorRetBienes") > 0 Then
                cPrint.printTexto 1.6, PosLinea, "I.V.A."
                cPrint.printTexto 2.9, PosLinea, "Retención I.V.A. Bienes"
                cPrint.printTexto 13.8, PosLinea, .fields("PorRetBienes")    '"---"
                cPrint.printVariable 14.7, PosLinea, .fields("MontoIvaBienes")
                cPrint.printTexto 17.3, PosLinea, .fields("Porc_Bienes") & "%"
                cPrint.printVariable 18.2, PosLinea, .fields("ValorRetBienes")
                Sumatoria = Sumatoria + .fields("ValorRetBienes")
                PosLinea = PosLinea + 0.4
            End If
            If .fields("ValorRetServicios") > 0 Then
                cPrint.printTexto 1.6, PosLinea, "I.V.A."
                cPrint.printTexto 2.9, PosLinea, "Retención I.V.A. Servicios"
                cPrint.printTexto 13.8, PosLinea, .fields("PorRetServicios")  '"---"
                cPrint.printVariable 14.7, PosLinea, .fields("MontoIvaServicios")
                cPrint.printTexto 17.3, PosLinea, .fields("Porc_Servicios") & "%"
                cPrint.printVariable 18.2, PosLinea, .fields("ValorRetServicios")
                Sumatoria = Sumatoria + .fields("ValorRetServicios")
                PosLinea = PosLinea + 0.4
            End If
         End If
       End With
       With AdoDBAir
        If .RecordCount > 0 Then
            Progreso_Barra.Valor_Maximo = Progreso_Barra.Valor_Maximo + .RecordCount
            Do While Not .EOF
               Progreso_Barra.Mensaje_Box = "Generando documento PDF"
               Progreso_Esperar
               ConceptoRet = .fields("Concepto")
               cPrint.printTexto 1.6, PosLinea, "RENTA"
               PosLinea = cPrint.printTextoMultiple(2.9, PosLinea, ConceptoRet, 10.5)
               If cPrint.dNoLineas > 0 Then PosLinea = cPrint.dNoLineas
               cPrint.printTexto 13.8, PosLinea, .fields("CodRet")
               cPrint.printVariable 14.7, PosLinea, .fields("BaseImp")
               cPrint.printTexto 17.3, PosLinea, Format$(.fields("Porcentaje"), "00.00%")
               cPrint.printVariable 18.2, PosLinea, .fields("ValRet")
               Sumatoria = Sumatoria + .fields("ValRet")
               PosLinea = PosLinea + 0.4
              .MoveNext
            Loop
        End If
       End With
       cPrint.printCuadro 1.5, TempPosLinea, 20, PosLinea - 0.6, Negro, "B"
       cPrint.printLinea 2.8, TempPosLinea, 2.8, PosLinea, Negro
       cPrint.printLinea 13.6, TempPosLinea, 13.6, PosLinea, Negro
       cPrint.printLinea 15, TempPosLinea, 15, PosLinea, Negro
       cPrint.printLinea 17, TempPosLinea, 17, PosLinea, Negro
       cPrint.printLinea 18.5, TempPosLinea, 18.5, PosLinea, Negro
       PosLinea = PosLinea + 0.1
       TempPosLinea = PosLinea
    End If
   'Pie de la Factura
    With AdoDBCompras
     If .RecordCount > 0 Then
         PosLinea = PosLinea + 0.05
         cPrint.tipoNegrilla = True
         cPrint.printTexto 15.2, PosLinea + 0.2, "T O T A L   R E T E N I D O"
         cPrint.tipoNegrilla = False
         cPrint.printVariable 18.2, PosLinea + 0.2, Sumatoria
         cPrint.tipoNegrilla = True
         cPrint.printTexto 1.6, PosLinea, "INFORMACION ADICIONAL"
         cPrint.tipoNegrilla = False
         PosLinea = PosLinea + 0.4
         Cadena = ""
         Codigo = TrimStrg(.fields("Telefono"))
         Codigo = Replace(Codigo, " ", "")
         Codigo = Replace(Codigo, ".", "")
         Codigo = Replace(Codigo, "-", "")
         If Val(Codigo) > 0 Then Cadena = Cadena & "Teléfono: " & Codigo & ", "
         Cadena = Cadena & "Tipo Comprobante: " & TFA.TP & "-" & Format$(TFA.Numero, "00000000") & ", "
         Codigo = TrimStrg(.fields("Email"))
         If InStr(Codigo, "@") And Len(Codigo) > 3 Then Cadena = Cadena & "Email: " & Codigo
         cPrint.printTexto 1.6, PosLinea, Cadena
         cPrint.printCuadro 1.5, TempPosLinea, 14.5, PosLinea - 0.25, Negro, "B"
         cPrint.printCuadro 15.1, TempPosLinea, 20, PosLinea - 0.25, Negro, "B"
         cPrint.printLinea 1.5, TempPosLinea + 0.35, 14.5, TempPosLinea + 0.35, Negro
     End If
    End With
    cPrint.finalizaImpresion
 '  ObjPDF.PDFEndPage
    AdoDBAir.Close
    AdoDBCompras.Close
    RatonNormal
End Sub

Public Sub SRI_Generar_PDF_LC(TFA As Tipo_Facturas, _
                              VerLiquidacionDeCompras As Boolean)
Dim AdoDBCompras As ADODB.Recordset
Dim AdoDBAir As ADODB.Recordset
Dim ConsultarDetalle As Boolean
Dim TempPosLinea As Single
Dim PosLineaTemp As Single
Dim InfoPosLinea As Single
Dim PorteDeLetra As Single
Dim EjercicioFiscal As String
Dim ConceptoRet As String
Dim NombreTipoDeLetra As String

    RatonReloj
    ConsultarDetalle = False
    NombreTipoDeLetra = TipoArialNarrow
   'Listar las Retenciones del IVA
    sSQL = "SELECT C.Cliente,C.CI_RUC,C.TD,C.Direccion,C.Email,C.Ciudad,C.DirNumero,C.Telefono,Co.Concepto,TC.* " _
         & "FROM Trans_Compras As TC,Clientes As C,Comprobantes AS Co " _
         & "WHERE TC.Item = '" & NumEmpresa & "' " _
         & "AND TC.Periodo = '" & Periodo_Contable & "' " _
         & "AND TC.Numero = " & TFA.Numero & " " _
         & "AND TC.TP = '" & TFA.TP & "' " _
         & "AND TC.TipoComprobante IN(3,41) " _
         & "AND TC.IdProv = C.Codigo " _
         & "AND TC.Item = Co.Item " _
         & "AND TC.Periodo = Co.Periodo " _
         & "AND TC.TP = Co.TP " _
         & "AND TC.Numero = Co.Numero " _
         & "ORDER BY Establecimiento, PuntoEmision,Secuencial "
    Select_AdoDB AdoDBCompras, sSQL, "Encabezado_LC"
    With AdoDBCompras
     If .RecordCount > 0 Then
         TFA.Fecha = .fields("Fecha")
         TFA.Cliente = .fields("Cliente")
         TFA.Razon_Social = .fields("Cliente")
         TFA.CI_RUC = .fields("CI_RUC")
         TFA.RUC_CI = .fields("CI_RUC")
         TFA.DireccionC = .fields("Direccion")
         TFA.Fecha_Aut = .fields("Fecha_Aut")
         TFA.Hora = .fields("Hora_Aut")
         TFA.Autorizacion_LC = .fields("Autorizacion")
         TFA.ClaveAcceso_LC = .fields("Clave_Acceso_LC")
         TFA.Serie_LC = .fields("Establecimiento") & .fields("PuntoEmision")
         TFA.Factura = .fields("Secuencial")
         TFA.Fecha = .fields("Fecha")
         TFA.EmailC = .fields("Email")
         TFA.Sin_IVA = .fields("BaseImponible")
         TFA.Con_IVA = .fields("BaseImpGrav")
         TFA.Total_IVA = .fields("MontoIva")
         TFA.Nota = .fields("Concepto")
         If TFA.Nota = Ninguno Then TFA.Nota = "Liquidacion de Compras"
         TFA.Total_MN = TFA.Sin_IVA + TFA.Con_IVA + TFA.Total_IVA
         TFA.SubTotal = TFA.Sin_IVA + TFA.Con_IVA
         Validar_Porc_IVA TFA.Fecha
         TFA.Porc_IVA = Porc_IVA * 100
         TFA.Tipo_Comp = CStr(.fields("TipoComprobante"))
         TFA.Si_Existe_Doc = True
         FechaTexto = TFA.Fecha
         EjercicioFiscal = CStr(Year(TFA.Fecha))
         ConsultarDetalle = True
     End If
    End With
    
   'Determinamos el Tipo de Comprobante
    sSQL = "SELECT Tipo_Comprobante_Codigo, Descripcion " _
         & "FROM Tipo_Comprobante " _
         & "WHERE TC = 'TDC' " _
         & "AND Tipo_Comprobante_Codigo = " & Val(TFA.Tipo_Comp) & " "
    Select_AdoDB AdoDBAir, sSQL
    If AdoDBAir.RecordCount > 0 Then TFA.Tipo_Comp = AdoDBAir.fields("Descripcion")
    AdoDBAir.Close
    PorteDeLetra = 7
   'Listar las Retenciones de la Fuente
    sSQL = "SELECT TIV.Concepto,R.* " _
         & "FROM Trans_Air As R,Tipo_Concepto_Retencion As TIV " _
         & "WHERE R.Item = '" & NumEmpresa & "' " _
         & "AND R.Periodo = '" & Periodo_Contable & "' " _
         & "AND R.Numero = " & TFA.Numero & " " _
         & "AND R.TP = '" & TFA.TP & "' " _
         & "AND R.SecRetencion = " & TFA.Retencion & " " _
         & "AND R.EstabRetencion = '" & MidStrg(TFA.Serie_R, 1, 3) & "' " _
         & "AND R.PtoEmiRetencion = '" & MidStrg(TFA.Serie_R, 4, 3) & "' " _
         & "AND R.Tipo_Trans IN ('C','I') " _
         & "AND TIV.Fecha_Inicio <= #" & BuscarFecha(FechaTexto) & "# " _
         & "AND TIV.Fecha_Final >= #" & BuscarFecha(FechaTexto) & "# " _
         & "AND R.CodRet = TIV.Codigo " _
         & "ORDER BY R.Cta_Retencion "
    Select_AdoDB AdoDBAir, sSQL, "Detalle_LC"
   'Encabezado Factura
    With AdoDBCompras
     If .RecordCount > 0 Then
         PosLinea = SRI_Generar_Documento_PDF(NombreTipoDeLetra, TFA, VerLiquidacionDeCompras, "LC")
         cPrint.letraTipo NombreTipoDeLetra, 6
         TempPosLinea = PosLinea
         
         PosLinea = PosLinea + 0.1
         cPrint.tipoNegrilla = True
         cPrint.printTexto 1.6, PosLinea, "Codigo", PorteDeLetra
         cPrint.printTexto 3.4, PosLinea, "Codigo", PorteDeLetra
         cPrint.printTexto 5, PosLinea, "Cantidad", PorteDeLetra
         cPrint.printTexto 6.3, PosLinea, "D e s c r i p c i ó n"
         'cPrint.printTexto 13.6, PosLinea, "Detalle Adicional", PorteDeLetra
         cPrint.printTexto 17.5, PosLinea, "Precio", PorteDeLetra
         'cPrint.printTexto 17.5, PosLinea, "Descuento", PorteDeLetra
         cPrint.printTexto 19.3, PosLinea, "Precio", PorteDeLetra
         PosLinea = PosLinea + 0.25
         cPrint.printTexto 1.6, PosLinea, "Unitario", PorteDeLetra
         cPrint.printTexto 3.4, PosLinea, "Auxiliar", PorteDeLetra
         cPrint.printTexto 17.5, PosLinea, "Unitario", PorteDeLetra
         cPrint.printTexto 19.3, PosLinea, "Total", PorteDeLetra
        'Detalle de la Factura
         PosLinea = PosLinea + 0.45 ' 10.4
         cPrint.tipoNegrilla = False
         cPrint.colorDeLetra = Negro
     End If
    End With
   'Detalle Factura
    cPrint.letraTipo NombreTipoDeLetra, 6
    Sumatoria = 0
    If ConsultarDetalle Then
       With AdoDBCompras
        If .RecordCount > 0 Then
            If TFA.Sin_IVA > 0 Then
               cPrint.printTexto 1.6, PosLinea, "99"
               cPrint.printTexto 3.4, PosLinea, "99.98"
               cPrint.printTexto 5, PosLinea, "1"
               PosLineaTemp = cPrint.printTextoMultiple(6.3, PosLinea, TFA.Nota & ", Tarifa 0%", 11)
               If PosLineaTemp > PosLinea Then PosLinea = PosLineaTemp
               cPrint.printVariable 16.8, PosLinea, TFA.Sin_IVA
               cPrint.printVariable 18.5, PosLinea, TFA.Sin_IVA
               PosLinea = PosLinea + 0.35
            End If
            If TFA.Con_IVA > 0 Then
               cPrint.printTexto 1.6, PosLinea, "99"
               cPrint.printTexto 3.4, PosLinea, "99.99"
               cPrint.printTexto 5, PosLinea, "1"
               PosLineaTemp = cPrint.printTextoMultiple(6.3, PosLinea, TFA.Nota & ", Tarifa " & TFA.Porc_IVA & "%", 11)
               If PosLineaTemp > PosLinea Then PosLinea = PosLineaTemp
               cPrint.printVariable 16.8, PosLinea, TFA.Con_IVA
               cPrint.printVariable 18.5, PosLinea, TFA.Con_IVA
               PosLinea = PosLinea + 0.35
            End If
        End If
       End With
       
      'Cuadro exterior del encabezado de la LC
       cPrint.printCuadro 1.5, TempPosLinea, 20, PosLinea - 0.6, Negro, "B"
       cPrint.printLinea 1.5, TempPosLinea + 0.6, 20, TempPosLinea + 0.6, Negro
       
       cPrint.printLinea 2.9, TempPosLinea, 2.9, PosLinea, Negro
       cPrint.printLinea 4.8, TempPosLinea, 4.8, PosLinea, Negro
       cPrint.printLinea 6.2, TempPosLinea, 6.2, PosLinea, Negro
       cPrint.printLinea 17.4, TempPosLinea, 17.4, PosLinea, Negro
       cPrint.printLinea 18.8, TempPosLinea, 18.8, PosLinea, Negro
    End If
    PosLinea = PosLinea + 0.15
    TempPosLinea = PosLinea
   'Lineas del Pie de la Factura
    'PosLinea = PosLinea - 0.1
    cPrint.PorteDeLetra = 6
    For I = 1 To 10
        PosLinea = PosLinea + 0.35
        cPrint.printLinea 14.4, PosLinea, 20, PosLinea, Negro
       'Raya de Informe y Leyenda
        If I = 1 Then cPrint.printLinea 1.5, PosLinea + 0.05, 13.6, PosLinea + 0.05, Negro
        If I = 8 Then cPrint.printLinea 1.5, PosLinea, 13.6, PosLinea, Negro
    Next I
    PosLinea = PosLinea - 0.25
    cPrint.printCuadro 1.5, TempPosLinea, 13.6, PosLinea, Negro, "B"
    cPrint.printCuadro 14.4, TempPosLinea, 20, PosLinea, Negro, "B"
   'Datos del Pie de la Factura
    'PosLinea = TempPosLinea + 0.1
    With TFA
     If .Si_Existe_Doc Then
         PosLinea = TempPosLinea + 0.05
         PosColumna = 14.5
         cPrint.colorDeLetra = Negro
         cPrint.tipoNegrilla = True
         cPrint.letraTipo NombreTipoDeLetra, PorteDeLetra
         cPrint.printTexto 1.6, PosLinea, "INFORMACION ADICIONAL:"
         cPrint.printTexto PosColumna, PosLinea, "SUBTOTAL " & .Porc_IVA & "%:"
         PosLinea = PosLinea + 0.35
         InfoPosLinea = PosLinea + 0.2
         cPrint.printTexto PosColumna, PosLinea, "SUBTOTAL 0%:"
         PosLinea = PosLinea + 0.35
         cPrint.printTexto PosColumna, PosLinea, "TOTAL DESCUENTO:"
         PosLinea = PosLinea + 0.35
         cPrint.printTexto PosColumna, PosLinea, "SUBTOTAL NO OBJETO DE IVA:"
         PosLinea = PosLinea + 0.35
         cPrint.printTexto PosColumna, PosLinea, "SUBTOTAL EXENTO DE IVA:"
         PosLinea = PosLinea + 0.35
         cPrint.printTexto PosColumna, PosLinea, "SUBTOTAL SIN IMPUESTOS:"
         PosLinea = PosLinea + 0.35
         cPrint.printTexto PosColumna, PosLinea, "ICE:"
         PosLinea = PosLinea + 0.35
         cPrint.printTexto PosColumna, PosLinea, "IVA " & .Porc_IVA & "%:"
         PosLinea = PosLinea + 0.35
         cPrint.printTexto PosColumna, PosLinea, "IVA 0%:"
         PosLinea = PosLinea + 0.35
         cPrint.printTexto PosColumna, PosLinea, "PROPINA:"
         PosLinea = PosLinea + 0.35
         cPrint.printTexto PosColumna, PosLinea, "VALOR TOTAL:"
         PosLinea = PosLinea + 0.35
         cPrint.tipoNegrilla = False
         cPrint.letraTipo NombreTipoDeLetra, 6
         PosLinea = TempPosLinea + 0.05
         PosColumna = 18.5
         cPrint.printVariable PosColumna, PosLinea, .Con_IVA, , , , 2
         PosLinea = PosLinea + 0.35
         cPrint.printVariable PosColumna, PosLinea, .Sin_IVA, , , , 2
         PosLinea = PosLinea + 0.35
         cPrint.printVariable PosColumna, PosLinea, .Total_Descuento, , , , 2
         PosLinea = PosLinea + 0.35
         cPrint.printVariable PosColumna, PosLinea, .Total_Sin_No_IVA, , , , 2
         PosLinea = PosLinea + 0.35
         cPrint.printVariable PosColumna, PosLinea, .Total_Sin_No_IVA, , , , 2
         PosLinea = PosLinea + 0.35
         cPrint.printVariable PosColumna, PosLinea, .SubTotal - .Total_Descuento, , , , 2
         PosLinea = PosLinea + 0.35
         cPrint.printVariable PosColumna, PosLinea, .Total_Sin_No_IVA, , , , 2
         PosLinea = PosLinea + 0.35
         cPrint.printVariable PosColumna, PosLinea, .Total_IVA, , , , 2
         PosLinea = PosLinea + 0.35
         cPrint.printVariable PosColumna, PosLinea, .Total_Sin_No_IVA, , , , 2
         PosLinea = PosLinea + 0.35
         cPrint.printVariable PosColumna, PosLinea, .Total_Sin_No_IVA, , , , 2
         PosLinea = PosLinea + 0.35
         cPrint.printVariable PosColumna, PosLinea, .Total_MN, , , , 2
         PosLinea = PosLinea + 0.4
     End If
    End With
        
   'Pie de la Factura
    With AdoDBCompras
     If .RecordCount > 0 Then
         PosLinea = InfoPosLinea
         cPrint.printTexto 1.6, PosLinea, "Teléfono: " & .fields("Telefono")
         cPrint.printTexto 7.5, PosLinea, "Tipo Comprobante: " & TFA.TP & "-" & Format$(TFA.Numero, "00000000")
         PosLinea = PosLinea + 0.35
         cPrint.printTexto 1.6, PosLinea, "Email: " & .fields("Email")
         'cPrint.printCuadroLinea 1.5, TempPosLinea, 12.9, PosLinea + 0.5, Negro, "B"
         'cPrint.printCuadroLinea 1.1, TempPosLinea + 0.2, 13.2, TempPosLinea + 0.2, Negro
     End If
    End With
    cPrint.finalizaImpresion
    AdoDBAir.Close
    AdoDBCompras.Close
    RatonNormal
End Sub

Public Sub SRI_Actualizar_Documento_XML(ClaveDeAcceso As String)
Dim AdoDBXML As ADODB.Recordset
Dim RutaXMLAutorizado As String
Dim DatosXMLA As String
Dim FileXML As String
Dim TD As String
Dim SerieF As String
Dim Documento As Long

 If Len(ClaveDeAcceso) >= 13 Then
    RutaXMLAutorizado = RutaDocumentos & "\Comprobantes Autorizados\" & ClaveDeAcceso & ".xml"
    DatosXMLA = Leer_Archivo_Texto(RutaXMLAutorizado)
    If Len(DatosXMLA) > 1 Then
       SerieF = MidStrg(ClaveDeAcceso, 25, 6)
       Documento = Val(MidStrg(ClaveDeAcceso, 31, 9))
       Select Case MidStrg(ClaveDeAcceso, 9, 2)
         Case "01": TD = "FA"
         Case "03": TD = "LC"
         Case "04": TD = "NC"
         Case "06": TD = "GR"
         Case "07": TD = "RE"
         Case Else: TD = "XX"
       End Select
       sSQL = "SELECT * " _
            & "FROM Trans_Documentos " _
            & "WHERE Item = '" & NumEmpresa & "' " _
            & "AND Periodo = '" & Periodo_Contable & "' " _
            & "AND Clave_Acceso = '" & ClaveDeAcceso & "' "
       Select_AdoDB AdoDBXML, sSQL
       If AdoDBXML.RecordCount <= 0 Then
          AdoDBXML.AddNew
          AdoDBXML.fields("Item") = NumEmpresa
          AdoDBXML.fields("Periodo") = Periodo_Contable
          AdoDBXML.fields("Clave_Acceso") = ClaveDeAcceso
       End If
       AdoDBXML.fields("TD") = TD
       AdoDBXML.fields("Serie") = SerieF
       AdoDBXML.fields("Documento") = Documento
       AdoDBXML.fields("Documento_Autorizado") = DatosXMLA
       AdoDBXML.Update
       AdoDBXML.Close
    End If
 End If
End Sub

'''Public Sub SRI_Actualizar_XML_Factura(TFA As Tipo_Facturas, _
'''                                      SRI_Autorizacion As Tipo_Estado_SRI)
'''Dim SRI_Error As String
'''    SRI_Error = SRI_Autorizacion.Error_SRI
'''    SRI_Error = Replace(SRI_Error, "'", "`")
'''    SRI_Error = Replace(SRI_Error, vbCrLf, "||")
'''    SRI_Error = Replace(SRI_Error, "&", " y ")
'''    SRI_Error = Replace(SRI_Error, "#", "No.")
'''    SRI_Error = TrimStrg(MidStrg(SRI_Error, 1, 100))
'''    If SRI_Error = "" Then SRI_Error = Ninguno
'''    sSQL = "UPDATE Facturas " _
'''         & "SET Estado_SRI = '" & SRI_Autorizacion.Estado_SRI & "', " _
'''         & "Clave_Acceso = '" & TFA.ClaveAcceso & "', "
'''    If SRI_Autorizacion.Estado_SRI = "OK" Then sSQL = sSQL & "Autorizacion = '" & SRI_Autorizacion.Autorizacion & "', "
'''    sSQL = sSQL _
'''         & "Error_FA_SRI = '" & SRI_Error & "' " _
'''         & "WHERE Item = '" & NumEmpresa & "' " _
'''         & "AND Periodo = '" & Periodo_Contable & "' " _
'''         & "AND TC = '" & TFA.TC & "' " _
'''         & "AND Serie = '" & TFA.Serie & "' " _
'''         & "AND Factura = " & TFA.Factura & " " _
'''         & "AND CodigoC = '" & TFA.CodigoC & "' " _
'''         & "AND Autorizacion = '" & TFA.Autorizacion & "' "
'''    Ejecutar_SQL_SP sSQL
'''End Sub

Public Sub SRI_Actualizar_XML_Retencion(SRI_Autorizacion As Tipo_Estado_SRI, _
                                        TFA As Tipo_Facturas)
   sSQL = "UPDATE Trans_Compras " _
        & "SET Estado_SRI = '" & SRI_Autorizacion.Estado_SRI & "', " _
        & "Clave_Acceso = '" & TFA.ClaveAcceso & "' " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND Periodo = '" & Periodo_Contable & "' " _
        & "AND TP = '" & TFA.TP & "' " _
        & "AND Numero = '" & TFA.Numero & "' " _
        & "AND Serie_Retencion = '" & TFA.Serie_R & "' " _
        & "AND SecRetencion = '" & TFA.Retencion & "' " _
        & "AND AutRetencion = '" & TFA.Autorizacion_R & "' "
   Ejecutar_SQL_SP sSQL
End Sub

Public Sub SRI_Actualizar_XML_Liquidacion(SRI_Autorizacion As Tipo_Estado_SRI, _
                                          TFA As Tipo_Facturas)
   sSQL = "UPDATE Trans_Compras " _
        & "SET Estado_SRI_LC = '" & SRI_Autorizacion.Estado_SRI & "', " _
        & "Clave_Acceso_LC = '" & TFA.ClaveAcceso_LC & "' " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND Periodo = '" & Periodo_Contable & "' " _
        & "AND TP = '" & TFA.TP & "' " _
        & "AND Numero = '" & TFA.Numero & "' " _
        & "AND Establecimiento+PuntoEmision = '" & TFA.Serie_LC & "' "
   Ejecutar_SQL_SP sSQL
End Sub

Public Sub SRI_Actualizar_XML_Nota_Credito(SRI_Autorizacion As Tipo_Estado_SRI, _
                                           TFA As Tipo_Facturas)
   sSQL = "UPDATE Trans_Abonos " _
        & "SET Estado_SRI_NC = '" & SRI_Autorizacion.Estado_SRI & "', " _
        & "Clave_Acceso_NC = '" & TFA.ClaveAcceso_NC & "' " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND Periodo = '" & Periodo_Contable & "' " _
        & "AND TP = '" & TFA.TC & "' " _
        & "AND Serie = '" & TFA.Serie & "' " _
        & "AND Factura = " & TFA.Factura & " " _
        & "AND CodigoC = '" & TFA.CodigoC & "' " _
        & "AND Autorizacion = '" & TFA.Autorizacion & "' " _
        & "AND Serie_NC = '" & TFA.Serie_NC & "' " _
        & "AND Secuencial_NC = " & TFA.Nota_Credito & " " _
        & "AND Banco = 'NOTA DE CREDITO' "
  'MsgBox sSQL
   Ejecutar_SQL_SP sSQL
End Sub

Public Sub SRI_Actualizar_XML_Guia_Remision(SRI_Autorizacion As Tipo_Estado_SRI, _
                                            TFA As Tipo_Facturas)
   sSQL = "UPDATE Facturas_Auxiliares " _
        & "SET Estado_SRI_GR = '" & SRI_Autorizacion.Estado_SRI & "', " _
        & "Clave_Acceso_GR = '" & TFA.ClaveAcceso_NC & "' " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND Periodo = '" & Periodo_Contable & "' " _
        & "AND TP = '" & TFA.TC & "' " _
        & "AND Serie = '" & TFA.Serie & "' " _
        & "AND Factura = " & TFA.Factura & " " _
        & "AND CodigoC = '" & TFA.CodigoC & "' " _
        & "AND Autorizacion = '" & TFA.Autorizacion & "' "
   'MsgBox sSQL
   Ejecutar_SQL_SP sSQL
End Sub

Public Function SRI_Mensaje_Error(ClaveAcceso As String) As String
Dim SI As Long
Dim SF As Long
Dim Result As String
Dim Mensaje As String
    Result = ""
    Mensaje = Leer_Archivo_Texto(RutaDocumentos & "\Comprobantes no Autorizados\" & ClaveAcceso & ".xml")
    If Len(Mensaje) > 1 Then
       Mensaje = Replace(Mensaje, "  ", "")
       SI = InStr(Mensaje, "<mensajes>")
       SF = InStr(Mensaje, "</mensajes>")
       If SI > 0 And SF > 0 Then
          SI = SI + 12
          Result = TrimStrg(MidStrg(Mensaje, SI, SF - SI))
       End If
    End If
    SRI_Mensaje_Error = Result
End Function

Public Sub SRI_Actualizar_Autorizacion_Factura(TFA As Tipo_Facturas, _
                                               SRI_Autorizacion As Tipo_Estado_SRI)
Dim Error_SRI As String
With SRI_Autorizacion
   'MsgBox "Actualizacion de la Factura: " & .Fecha_Autorizacion
    If Len(.Autorizacion) >= 13 Then
       'Determinamos el tipo de Error
        Error_SRI = Replace(.Error_SRI, "'", "`")
        Error_SRI = Replace(Error_SRI, vbCrLf, "||")
        Error_SRI = Replace(Error_SRI, "&", " y ")
        Error_SRI = Replace(Error_SRI, "#", "No.")
        Error_SRI = TrimStrg(MidStrg(Error_SRI, 1, 100))
        If Error_SRI = "" Then Error_SRI = Ninguno
        
       'Actualizamos el estado del documento
        sSQL = "UPDATE Facturas " _
             & "SET Clave_Acceso = '" & .Clave_De_Acceso & "', "
        If .Estado_SRI = "OK" Then sSQL = sSQL & "Autorizacion = '" & .Autorizacion & "', "
        sSQL = sSQL _
             & "Fecha_Aut = #" & BuscarFecha(.Fecha_Autorizacion) & "#, " _
             & "Hora_Aut = '" & .Hora_Autorizacion & "', " _
             & "Estado_SRI = '" & .Estado_SRI & "', " _
             & "Error_FA_SRI = '" & Error_SRI & "' " _
             & "WHERE Item = '" & NumEmpresa & "' " _
             & "AND Periodo = '" & Periodo_Contable & "' " _
             & "AND TC = '" & TFA.TC & "' " _
             & "AND Serie = '" & TFA.Serie & "' " _
             & "AND Factura = " & TFA.Factura & " " _
             & "AND Autorizacion = '" & TFA.Autorizacion & "' " _
             & "AND CodigoC = '" & TFA.CodigoC & "' "
        Ejecutar_SQL_SP sSQL
        TFA.Estado_SRI = .Estado_SRI
        If .Estado_SRI = "OK" Then
            sSQL = "UPDATE Detalle_Factura " _
                 & "SET Autorizacion = '" & .Autorizacion & "' " _
                 & "WHERE Item = '" & NumEmpresa & "' " _
                 & "AND Periodo = '" & Periodo_Contable & "' " _
                 & "AND TC = '" & TFA.TC & "' " _
                 & "AND Serie = '" & TFA.Serie & "' " _
                 & "AND Factura = " & TFA.Factura & " " _
                 & "AND CodigoC = '" & TFA.CodigoC & "' " _
                 & "AND Autorizacion = '" & TFA.Autorizacion & "' "
            Ejecutar_SQL_SP sSQL
            sSQL = "UPDATE Trans_Abonos " _
                 & "SET Autorizacion = '" & .Autorizacion & "' " _
                 & "WHERE Item = '" & NumEmpresa & "' " _
                 & "AND Periodo = '" & Periodo_Contable & "' " _
                 & "AND TP = '" & TFA.TC & "' " _
                 & "AND Serie = '" & TFA.Serie & "' " _
                 & "AND Factura = " & TFA.Factura & " " _
                 & "AND CodigoC = '" & TFA.CodigoC & "' " _
                 & "AND Autorizacion = '" & TFA.Autorizacion & "' "
            Ejecutar_SQL_SP sSQL
            sSQL = "UPDATE Facturas_Auxiliares " _
                 & "SET Autorizacion = '" & .Autorizacion & "' " _
                 & "WHERE Item = '" & NumEmpresa & "' " _
                 & "AND Periodo = '" & Periodo_Contable & "' " _
                 & "AND TC = '" & TFA.TC & "' " _
                 & "AND Serie = '" & TFA.Serie & "' " _
                 & "AND Factura = " & TFA.Factura & " " _
                 & "AND CodigoC = '" & TFA.CodigoC & "' " _
                 & "AND Autorizacion = '" & TFA.Autorizacion & "' "
            Ejecutar_SQL_SP sSQL
           'Actualizamos el documento autorizado en la base de datos del sistema
            SRI_Actualizar_Documento_XML .Clave_De_Acceso
        End If
    End If
End With
End Sub

Public Sub SRI_Actualizar_Autorizacion_Nota_Credito(TFA As Tipo_Facturas, _
                                                    SRI_Autorizacion As Tipo_Estado_SRI)
    With SRI_Autorizacion
        'MsgBox TFA.Autorizacion_NC
        TFA.Fecha_Aut_NC = .Fecha_Autorizacion
        TFA.Hora_NC = .Hora_Autorizacion
        sSQL = "UPDATE Trans_Abonos " _
             & "SET Autorizacion_NC = '" & .Autorizacion & "', " _
             & "Fecha_Aut_NC = #" & BuscarFecha(.Fecha_Autorizacion) & "#, " _
             & "Hora_Aut_NC = '" & .Hora_Autorizacion & "', " _
             & "Estado_SRI_NC = '" & .Estado_SRI & "' " _
             & "WHERE Item = '" & NumEmpresa & "' " _
             & "AND Periodo = '" & Periodo_Contable & "' " _
             & "AND TP = '" & TFA.TC & "' " _
             & "AND Serie = '" & TFA.Serie & "' " _
             & "AND Factura = " & TFA.Factura & " " _
             & "AND CodigoC = '" & TFA.CodigoC & "' " _
             & "AND Autorizacion = '" & TFA.Autorizacion & "' " _
             & "AND Serie_NC = '" & TFA.Serie_NC & "' " _
             & "AND Secuencial_NC = " & TFA.Nota_Credito & " " _
             & "AND Banco = 'NOTA DE CREDITO' "
        Ejecutar_SQL_SP sSQL
        
        sSQL = "UPDATE Detalle_Nota_Credito " _
             & "SET Autorizacion = '" & .Autorizacion & "' " _
             & "WHERE Item = '" & NumEmpresa & "' " _
             & "AND Periodo = '" & Periodo_Contable & "' " _
             & "AND TC = '" & TFA.TC & "' " _
             & "AND Serie_FA = '" & TFA.Serie & "' " _
             & "AND Factura = " & TFA.Factura & " " _
             & "AND Serie = '" & TFA.Serie_NC & "' " _
             & "AND Secuencial = " & TFA.Nota_Credito & " "
        Ejecutar_SQL_SP sSQL
    End With
End Sub

Public Sub SRI_Actualizar_Autorizacion_Guia_Remision(TFA As Tipo_Facturas, _
                                                     SRI_Autorizacion As Tipo_Estado_SRI)
With SRI_Autorizacion
    'MsgBox TFA.Autorizacion_NC
    TFA.Fecha_Aut_GR = .Fecha_Autorizacion
    TFA.Hora_GR = .Hora_Autorizacion
    sSQL = "UPDATE Facturas_Auxiliares " _
         & "SET Autorizacion_GR = '" & .Autorizacion & "', " _
         & "Fecha_Aut_GR = #" & BuscarFecha(.Fecha_Autorizacion) & "#, " _
         & "Hora_Aut_GR = '" & .Hora_Autorizacion & "', " _
         & "Estado_SRI_GR = '" & .Estado_SRI & "' " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND TC = '" & TFA.TC & "' " _
         & "AND Serie = '" & TFA.Serie & "' " _
         & "AND Factura = " & TFA.Factura & " " _
         & "AND CodigoC = '" & TFA.CodigoC & "' " _
         & "AND Autorizacion = '" & TFA.Autorizacion & "' "
    Ejecutar_SQL_SP sSQL
End With
End Sub

Public Sub SRI_Actualizar_Autorizacion_Retencion(SRI_Autorizacion As Tipo_Estado_SRI, _
                                                 TFA As Tipo_Facturas)
With SRI_Autorizacion
   sSQL = "UPDATE Trans_Compras " _
        & "SET AutRetencion = '" & .Autorizacion & "', " _
        & "Fecha_Aut = #" & BuscarFecha(.Fecha_Autorizacion) & "#, " _
        & "Hora_Aut = '" & .Hora_Autorizacion & "' " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND Periodo = '" & Periodo_Contable & "' " _
        & "AND TP = '" & TFA.TP & "' " _
        & "AND Numero = '" & TFA.Numero & "' " _
        & "AND Serie_Retencion = '" & TFA.Serie_R & "' " _
        & "AND SecRetencion = '" & TFA.Retencion & "' " _
        & "AND AutRetencion = '" & TFA.Autorizacion_R & "' "
   Ejecutar_SQL_SP sSQL
   
   sSQL = "UPDATE Trans_Air " _
        & "SET AutRetencion = '" & .Autorizacion & "' " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND Periodo = '" & Periodo_Contable & "' " _
        & "AND Tipo_Trans = 'C' " _
        & "AND TP = '" & TFA.TP & "' " _
        & "AND Numero = '" & TFA.Numero & "' " _
        & "AND EstabRetencion = '" & MidStrg(TFA.Serie_R, 1, 3) & "' " _
        & "AND PtoEmiRetencion = '" & MidStrg(TFA.Serie_R, 4, 3) & "' " _
        & "AND SecRetencion = '" & TFA.Retencion & "' " _
        & "AND AutRetencion = '" & TFA.Autorizacion_R & "' "
   Ejecutar_SQL_SP sSQL
End With
End Sub

Public Sub SRI_Actualizar_Autorizacion_Liquidacion(SRI_Autorizacion As Tipo_Estado_SRI, _
                                                   TFA As Tipo_Facturas)
With SRI_Autorizacion
   sSQL = "UPDATE Trans_Compras " _
        & "SET Autorizacion = '" & .Autorizacion & "', " _
        & "Fecha_Aut_LC = #" & BuscarFecha(.Fecha_Autorizacion) & "#, " _
        & "Hora_Aut_LC = '" & .Hora_Autorizacion & "' " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND Periodo = '" & Periodo_Contable & "' " _
        & "AND TP = '" & TFA.TP & "' " _
        & "AND Numero = '" & TFA.Numero & "' " _
        & "AND Establecimiento+PuntoEmision = '" & TFA.Serie_LC & "' "
   Ejecutar_SQL_SP sSQL
   
'''   sSQL = "UPDATE Trans_Air " _
'''        & "SET AutRetencion = '" & .Autorizacion & "' " _
'''        & "WHERE Item = '" & NumEmpresa & "' " _
'''        & "AND Periodo = '" & Periodo_Contable & "' " _
'''        & "AND Tipo_Trans = 'C' " _
'''        & "AND TP = '" & TFA.TP & "' " _
'''        & "AND Numero = '" & TFA.Numero & "' " _
'''        & "AND EstabRetencion = '" & MidStrg(TFA.Serie_R, 1, 3) & "' " _
'''        & "AND PtoEmiRetencion = '" & MidStrg(TFA.Serie_R, 4, 3) & "' " _
'''        & "AND SecRetencion = '" & TFA.Retencion & "' " _
'''        & "AND AutRetencion = '" & TFA.Autorizacion_R & "' "
'''   Ejecutar_SQL_SP sSQL
End With
End Sub

Public Sub SRI_Crear_Clave_Acceso_Retenciones(TFA As Tipo_Facturas, _
                                              VerRetenciones As Boolean, _
                                              Optional GeneraXML As Boolean)
Dim AdoCompras As ADODB.Recordset
Dim AdoAir As ADODB.Recordset
Dim ConBodegas As Boolean
Dim NumTrans As Long
Dim CodSustento As String
Dim tipoCodComprobante As String
Dim NumFile As Integer
Dim RutaGeneraFile As String
Dim ListadoErrores As String
Dim Autorizar_XML As Boolean

  RatonReloj
 'AgenteRetencion = "NAC-DNCRASC20-00000001"
  Autorizar_XML = True
  If Len(Fecha_Igualar) = 10 Then
     If CFechaLong(TFA.Fecha) < CFechaLong(Fecha_Igualar) Then Autorizar_XML = False
  End If
    
 'Autorizamos la Retenciones
  If Autorizar_XML Then

' RETENCIONES COMPRAS
  sSQL = "SELECT C.Cliente,C.CI_RUC,C.TD,C.Direccion,C.Telefono,C.Email,TC.* " _
       & "FROM Trans_Compras As TC, Clientes As C " _
       & "WHERE TC.Item = '" & NumEmpresa & "' " _
       & "AND TC.Periodo = '" & Periodo_Contable & "' " _
       & "AND TC.Serie_Retencion = '" & TFA.Serie_R & "' " _
       & "AND TC.SecRetencion = " & TFA.Retencion & " " _
       & "AND TC.TP = '" & TFA.TP & "' " _
       & "AND TC.Numero = " & TFA.Numero & " " _
       & "AND LEN(TC.AutRetencion) = 13 " _
       & "AND TC.IdProv = C.Codigo " _
       & "ORDER BY Serie_Retencion,SecRetencion "
  Select_AdoDB AdoCompras, sSQL
  With AdoCompras
   If .RecordCount > 0 Then
      'Generacion de la Retencion si es Electronica
       Do While Not .EOF
          TextoXML = ""
''          TFA.Serie_R = .Fields("Serie_Retencion")
''          TFA.Retencion = .Fields("SecRetencion")
          TFA.Autorizacion_R = .fields("AutRetencion")
          TFA.Autorizacion = .fields("Autorizacion")
          TFA.Fecha = .fields("FechaRegistro")
          TFA.Vencimiento = .fields("FechaRegistro")
          TFA.Serie = .fields("Establecimiento") & .fields("PuntoEmision")
          TFA.Factura = .fields("Secuencial")
          TFA.Hora = Format$(Time, FormatoTimes)
          TFA.Cliente = .fields("Cliente")
          TFA.CI_RUC = .fields("CI_RUC")
          TFA.TD = .fields("TD")
          TFA.DireccionC = .fields("Direccion")
          TFA.TelefonoC = .fields("Telefono")
          TFA.EmailC = .fields("Email")
          CodSustento = Format$(.fields("CodSustento"), "00")
          tipoCodComprobante = Format$(.fields("TipoComprobante"), "00")
          Obtener_Porc_IVA TFA.Fecha, .fields("PorcentajeIva")
         'Validar_Porc_IVA TFA.Fecha
         'Algoritmo Modulo 11 para la clave de la retencion
         '& Format$(TFA.Vencimiento, "ddmmyyyy")
          TFA.ClaveAcceso = Format$(TFA.Fecha, "ddmmyyyy") & "07" & RUC & Ambiente & TFA.Serie_R & Format$(TFA.Retencion, String(9, "0")) _
                         & "123456781"
          TFA.ClaveAcceso = Replace(TFA.ClaveAcceso, ".", "1")
          TFA.ClaveAcceso = TFA.ClaveAcceso & Digito_Verificador_Modulo11(TFA.ClaveAcceso)
          
          If Len(TFA.Autorizacion_R) >= 13 Then
            'ENCABEZADO XML PARA EL SRI DE LA RETENCION
             Insertar_Campo_XML "<?xml version=""1.0"" encoding=""UTF-8""?>"
             Insertar_Campo_XML "<comprobanteRetencion id=""comprobante"" version=""2.0.0"">"
             Insertar_Campo_XML AbrirXML("infoTributaria")
                Insertar_Campo_XML CampoXML("ambiente", Ambiente)
                Insertar_Campo_XML CampoXML("tipoEmision", "1")
                Insertar_Campo_XML CampoXML("razonSocial", RazonSocial)
                Insertar_Campo_XML CampoXML("nombreComercial", NombreComercial)
                Insertar_Campo_XML CampoXML("ruc", RUC)
                Insertar_Campo_XML CampoXML("claveAcceso", TFA.ClaveAcceso)
                Insertar_Campo_XML CampoXML("codDoc", "07")
                Insertar_Campo_XML CampoXML("estab", MidStrg(TFA.Serie_R, 1, 3))
                Insertar_Campo_XML CampoXML("ptoEmi", MidStrg(TFA.Serie_R, 4, 3))
                Insertar_Campo_XML CampoXML("secuencial", Format$(TFA.Retencion, String(9, "0")))
                Insertar_Campo_XML CampoXML("dirMatriz", ULCase(Direccion))
               'Tipo de Contribuyente
                If AgenteRetencion <> Ninguno Then Insertar_Campo_XML CampoXML("agenteRetencion", "1")
                If MicroEmpresa = "CONTRIBUYENTE RÉGIMEN RIMPE" Then Insertar_Campo_XML CampoXML("contribuyenteRimpe", "CONTRIBUYENTE RÉGIMEN RIMPE")
             Insertar_Campo_XML CerrarXML("infoTributaria")
             
             Insertar_Campo_XML AbrirXML("infoCompRetencion")
                Insertar_Campo_XML CampoXML("fechaEmision", TFA.Fecha)
                Insertar_Campo_XML CampoXML("dirEstablecimiento", ULCase(DireccionEstab))
                If Len(ContEspec) > 1 Then Insertar_Campo_XML CampoXML("contribuyenteEspecial", ContEspec)
                Insertar_Campo_XML CampoXML("obligadoContabilidad", Obligado_Conta)
                Select Case TFA.TD
                  Case "R": If TFA.CI_RUC = "9999999999999" Then TFA.TD = "07" Else TFA.TD = "04"
                  Case "C": TFA.TD = "05"
                  Case "P": TFA.TD = "06"
                End Select
                Insertar_Campo_XML CampoXML("tipoIdentificacionSujetoRetenido", TFA.TD)
                If .fields("PagoLocExt") = "01" Then
                    Insertar_Campo_XML CampoXML("parteRel", "NO")
                Else
                    Insertar_Campo_XML CampoXML("tipoSujetoRetenido", .fields("PagoLocExt"))
                    Insertar_Campo_XML CampoXML("parteRel", "SI")
                End If
                Insertar_Campo_XML CampoXML("razonSocialSujetoRetenido", TFA.Cliente)
                Insertar_Campo_XML CampoXML("identificacionSujetoRetenido", TFA.CI_RUC)
                Insertar_Campo_XML CampoXML("periodoFiscal", Format$(TFA.Fecha, "mm/yyyy"))
             Insertar_Campo_XML CerrarXML("infoCompRetencion")
             
             Insertar_Campo_XML AbrirXML("docsSustento")
                Insertar_Campo_XML AbrirXML("docSustento")
                    Total_Servicio = 0
                    Total_Propinas = 0
                    Total_Comision = 0
                    Total_Sin_No_IVA = 0
                    Total_Sin_IVA = .fields("BaseImponible")
                    Total_Con_IVA = .fields("BaseImpGrav")
                    Total_IVA = .fields("MontoIva")
                    Total_SubTotal = Total_Sin_IVA + Total_Con_IVA
                    Total_Factura = Total_SubTotal + Total_IVA
                    
                    Insertar_Campo_XML CampoXML("codSustento", CodSustento)
                    Insertar_Campo_XML CampoXML("codDocSustento", tipoCodComprobante) 'OJO
                    Insertar_Campo_XML CampoXML("numDocSustento", TFA.Serie & Format(TFA.Factura, "000000000"))
                    Insertar_Campo_XML CampoXML("fechaEmisionDocSustento", TFA.Fecha)
                    Insertar_Campo_XML CampoXML("fechaRegistroContable", TFA.Vencimiento)
                    Insertar_Campo_XML CampoXML("numAutDocSustento", TFA.Autorizacion)
                    Insertar_Campo_XML CampoXML("pagoLocExt", .fields("PagoLocExt"))
                    Insertar_Campo_XML CampoXML("totalSinImpuestos", Total_SubTotal, 2)
                    Insertar_Campo_XML CampoXML("importeTotal", Total_Factura, 2)
                    
                    Insertar_Campo_XML AbrirXML("impuestosDocSustento")
                        Insertar_Campo_XML AbrirXML("impuestoDocSustento")
                            Insertar_Campo_XML CampoXML("codImpuestoDocSustento", "2")
                            Insertar_Campo_XML CampoXML("codigoPorcentaje", .fields("PorcentajeIva"))
                            Insertar_Campo_XML CampoXML("baseImponible", Total_Con_IVA, 2)
                            Insertar_Campo_XML CampoXML("tarifa", Porc_IVA * 100)
                            Insertar_Campo_XML CampoXML("valorImpuesto", Total_IVA, 2)
                        Insertar_Campo_XML CerrarXML("impuestoDocSustento")
                        
                        Insertar_Campo_XML AbrirXML("impuestoDocSustento")
                            Insertar_Campo_XML CampoXML("codImpuestoDocSustento", "2")
                            Insertar_Campo_XML CampoXML("codigoPorcentaje", "0")
                            Insertar_Campo_XML CampoXML("baseImponible", Total_Sin_IVA, 2)
                            Insertar_Campo_XML CampoXML("tarifa", "0")
                            Insertar_Campo_XML CampoXML("valorImpuesto", "0.00")
                        Insertar_Campo_XML CerrarXML("impuestoDocSustento")
                    Insertar_Campo_XML CerrarXML("impuestosDocSustento")
                    
                    Insertar_Campo_XML AbrirXML("retenciones")
                   'RETENCIONES AIR
                     sSQL = "SELECT * " _
                          & "FROM Trans_Air " _
                          & "WHERE Item = '" & NumEmpresa & "' " _
                          & "AND Periodo = '" & Periodo_Contable & "' " _
                          & "AND Numero = " & TFA.Numero & " " _
                          & "AND TP = '" & TFA.TP & "' " _
                          & "AND Tipo_Trans = 'C' " _
                          & "AND EstabRetencion = '" & MidStrg(TFA.Serie_R, 1, 3) & "' " _
                          & "AND PtoEmiRetencion = '" & MidStrg(TFA.Serie_R, 4, 3) & "' " _
                          & "AND SecRetencion = " & TFA.Retencion & " " _
                          & "AND AutRetencion = '" & TFA.Autorizacion_R & "' " _
                          & "ORDER BY ID "
                     Select_AdoDB AdoAir, sSQL, "Retencion_" & CStr(TFA.Retencion)
                     If AdoAir.RecordCount > 0 Then
                       'MsgBox sSQL
                       ' MsgBox AdoAir.RecordCount
                        Do While Not AdoAir.EOF
                           If AdoAir.fields("BaseImp") > 0 Then
                              Insertar_Campo_XML AbrirXML("retencion")
                                 Insertar_Campo_XML CampoXML("codigo", "1")
                                 Insertar_Campo_XML CampoXML("codigoRetencion", AdoAir.fields("CodRet"))
                                 Insertar_Campo_XML CampoXML("baseImponible", AdoAir.fields("BaseImp"), 2)
                                 Insertar_Campo_XML CampoXML("porcentajeRetener", (AdoAir.fields("Porcentaje") * 100), 2)
                                 Insertar_Campo_XML CampoXML("valorRetenido", AdoAir.fields("ValRet"), 2)
                              Insertar_Campo_XML CerrarXML("retencion")
                           End If
                           AdoAir.MoveNext
                        Loop
                     
                     End If
                     AdoAir.Close
                    
                    If Val(.fields("Porc_Bienes")) > 0 Then
                       Insertar_Campo_XML AbrirXML("retencion")
                          Select Case Val(.fields("Porc_Bienes"))
                            Case 10: CodigoA = "9"
                            Case 30: CodigoA = "1"
                            Case 70: CodigoA = "2"
                            Case 100: CodigoA = "3"
                            Case Else: CodigoA = "2"
                          End Select
                          Total = .fields("MontoIvaBienes")
                          Retencion = Val(.fields("Porc_Bienes"))
                          Valor = Redondear(Total * (Retencion / 100), 2)
                          Insertar_Campo_XML CampoXML("codigo", "2")
                          Insertar_Campo_XML CampoXML("codigoRetencion", CodigoA)
                          Insertar_Campo_XML CampoXML("baseImponible", Total, 2)
                          Insertar_Campo_XML CampoXML("porcentajeRetener", Retencion, 2)
                          Insertar_Campo_XML CampoXML("valorRetenido", Valor, 2)
                       Insertar_Campo_XML CerrarXML("retencion")
                    End If
                    'MsgBox "|" & .fields("Porc_Servicios") & "|"
                    If Val(.fields("Porc_Servicios")) > 0 Then
                       Insertar_Campo_XML AbrirXML("retencion")
                          Select Case Val(.fields("Porc_Servicios"))
                            Case 20: CodigoA = "10"
                            Case 30: CodigoA = "1"
                            Case 70: CodigoA = "2"
                            Case 100: CodigoA = "3"
                            Case Else: CodigoA = "2"
                          End Select
                          Total = .fields("MontoIvaServicios")
                          Retencion = Val(.fields("Porc_Servicios"))
                          Valor = Redondear(Total * (Retencion / 100), 2)
                          Insertar_Campo_XML CampoXML("codigo", "2")
                          Insertar_Campo_XML CampoXML("codigoRetencion", CodigoA)
                          Insertar_Campo_XML CampoXML("baseImponible", Total, 2)
                          Insertar_Campo_XML CampoXML("porcentajeRetener", Retencion, 2)
                          Insertar_Campo_XML CampoXML("valorRetenido", Valor, 2)
                       Insertar_Campo_XML CerrarXML("retencion")
                    End If
                    Insertar_Campo_XML CerrarXML("retenciones")
                    
                    Insertar_Campo_XML AbrirXML("pagos")
                       Insertar_Campo_XML AbrirXML("pago")
                          Insertar_Campo_XML CampoXML("formaPago", .fields("FormaPago"))
                          Insertar_Campo_XML CampoXML("total", Total_Factura, 2)
                       Insertar_Campo_XML CerrarXML("pago")
                    Insertar_Campo_XML CerrarXML("pagos")
                    
                Insertar_Campo_XML CerrarXML("docSustento")
             Insertar_Campo_XML CerrarXML("docsSustento")
             'FIN DE XML DE RETENCION
             'MsgBox AgenteRetencion & vbCrLf & MicroEmpresa
             Insertar_Campo_XML AbrirXML("infoAdicional")
                If Len(TFA.DireccionC) > 1 Then Insertar_Campo_XML "<campoAdicional nombre=""Direccion"">" & TFA.DireccionC & "</campoAdicional>"
                If Len(TFA.TelefonoC) > 1 Then Insertar_Campo_XML "<campoAdicional nombre=""Telefono"">" & TFA.TelefonoC & "</campoAdicional>"
                If Len(TFA.EmailC) > 1 Then Insertar_Campo_XML "<campoAdicional nombre=""Email"">" & TFA.EmailC & "</campoAdicional>"
                Insertar_Campo_XML "<campoAdicional nombre=""Comprobante No"">" & TFA.TP & "-" & Format$(TFA.Numero, "0000000000") & "</campoAdicional>"
                If AgenteRetencion <> Ninguno Then Insertar_Campo_XML "<campoAdicional nombre=""Agente de Retencion"">No. Resolución: 1</campoAdicional>"
             Insertar_Campo_XML CerrarXML("infoAdicional")
             Insertar_Campo_XML CerrarXML("comprobanteRetencion")
             
             If CFechaLong(TFA.Fecha) <= CFechaLong(Fecha_CE) Then
                'Grabamos el comprobante XML de la Retencion
                 RutaGeneraFile = RutaDocumentos & "\Comprobantes Generados\" & TFA.ClaveAcceso & ".xml"
                'Copiar todo el contenido de la caja de texto
                
                 If GeneraXML Then
                    Clipboard.Clear
                    Clipboard.SetText TextoXML
                    Grabar_Consulta_Archivo "RE_" & TFA.Serie_R & "-" & Format$(TFA.Retencion, String(9, "0")), TextoXML
                 End If
                                  
                 Set DocumentoXML = New DOMDocument30
                 DocumentoXML.loadXML (TextoXML)
                 DocumentoXML.save (RutaGeneraFile)
                'MsgBox "Crear Archivos CE: " & RutaGeneraFile
                 If Len(Fecha_Igualar) = 10 Then
                    If CFechaLong(TFA.Fecha) < CFechaLong(Fecha_Igualar) Then Autorizar_XML = False
                 End If
                 
                     TFA.Estado_SRI = "CG"
                    'Enviamos al SRI para que me autorice la Retencion
                     SRI_Autorizacion = SRI_Generar_XML(TFA.ClaveAcceso, TFA.Estado_SRI)
                     SRI_Actualizar_XML_Retencion SRI_Autorizacion, FA
                     RatonReloj
                     If SRI_Autorizacion.Estado_SRI = "OK" Then
                        Control_Procesos "R", "RE-" & TFA.Serie_R & " No. " & Format(TFA.Retencion, "000000000") & " Autorizada"
                        SRI_Actualizar_Autorizacion_Retencion SRI_Autorizacion, FA
                        SRI_Actualizar_Documento_XML TFA.ClaveAcceso
                        If Ambiente = "2" Then SRI_Enviar_Mails FA, SRI_Autorizacion, "RE"
                        SRI_Generar_PDF_RE TFA, VerRetenciones
                        RatonNormal
                        'If VerRetenciones Then MsgBox "Comprobante de Retencion Autorizado con exito"
                     Else
                        RatonNormal
                        If TextoImprimio = "" Then TextoImprimio = "INFORME DE ERRORES EN LAS RETENCIONES:" & vbCrLf _
                                                                 & "--------------------------------------" & vbCrLf
                        TextoImprimio = TextoImprimio _
                                      & "Fecha: " & TFA.Fecha & ", Retencion No. " & TFA.Serie_R & "-" & TFA.Retencion _
                                      & ", Documento No. " & TFA.Serie & "-" & TFA.Factura _
                                      & ", De: " & TFA.Cliente & ", CI/RUC: " & TFA.CI_RUC & vbCrLf _
                                      & SRI_Autorizacion.Estado_SRI & " - " & SRI_Autorizacion.Error_SRI & vbCrLf _
                                      & String(80, "-") & vbCrLf
                        If VerRetenciones Then MsgBox SRI_Autorizacion.Error_SRI
                     End If
                 
             Else
                RatonNormal
                MsgBox MensajeNoAutorizarCE
             End If
          End If
         .MoveNext
      Loop
   End If
  End With
      AdoCompras.Close
  End If
  RatonNormal
End Sub

Public Sub SRI_Crear_Clave_Acceso_Liquidacion(TFA As Tipo_Facturas, _
                                              VerLiquidacion As Boolean, _
                                              Optional GeneraXML As Boolean)
Dim AdoCompras As ADODB.Recordset
Dim AdoAir As ADODB.Recordset
Dim ConBodegas As Boolean
Dim NumTrans As Long
Dim CodSustento As String
Dim tipoCodComprobante As String
Dim NumFile As Integer
Dim RutaGeneraFile As String
Dim ListadoErrores As String
Dim Autorizar_XML As Boolean

  RatonReloj
  Autorizar_XML = True
  If Len(Fecha_Igualar) = 10 Then
     If CFechaLong(TFA.Fecha) < CFechaLong(Fecha_Igualar) Then Autorizar_XML = False
  End If
 'Autorizamos la Retencion Electronica
  If Autorizar_XML Then

' RETENCIONES COMPRAS
  sSQL = "SELECT C.Cliente,C.CI_RUC,C.TD,C.Direccion,C.Telefono,C.Email,Co.Concepto,TC.* " _
       & "FROM Trans_Compras As TC, Clientes As C, Comprobantes AS Co " _
       & "WHERE TC.Item = '" & NumEmpresa & "' " _
       & "AND TC.Periodo = '" & Periodo_Contable & "' " _
       & "AND TC.Numero = " & TFA.Numero & " " _
       & "AND TC.TP = '" & TFA.TP & "' " _
       & "AND LEN(TC.Autorizacion) = 13 " _
       & "AND TC.TipoComprobante IN(3,41) " _
       & "AND TC.IdProv = C.Codigo " _
       & "AND TC.Item = Co.Item " _
       & "AND TC.Periodo = Co.Periodo " _
       & "AND TC.TP = Co.TP " _
       & "AND TC.Numero = Co.Numero " _
       & "ORDER BY Establecimiento, PuntoEmision,Secuencial "
  Select_AdoDB AdoCompras, sSQL        ' , "Liquidacion_Compras_" & TFA.TP & "-" & TFA.Numero
  With AdoCompras
   If .RecordCount > 0 Then
      'Generacion de la Retencion si es Electronica
       Do While Not .EOF
          TextoXML = ""
          TFA.Autorizacion = .fields("Autorizacion")
          TFA.Fecha = .fields("FechaRegistro")
          TFA.Vencimiento = .fields("FechaRegistro")
          TFA.Serie_LC = .fields("Establecimiento") & .fields("PuntoEmision")
          TFA.Factura = .fields("Secuencial")
          TFA.Hora = Format$(Time, FormatoTimes)
          TFA.Cliente = .fields("Cliente")
          TFA.CI_RUC = .fields("CI_RUC")
          TFA.TD = .fields("TD")
          TFA.DireccionC = .fields("Direccion")
          TFA.TelefonoC = .fields("Telefono")
          TFA.EmailC = .fields("Email")
          TFA.Sin_IVA = .fields("BaseImponible")
          TFA.Con_IVA = .fields("BaseImpGrav")
          TFA.Total_IVA = .fields("MontoIva")
          TFA.Nota = .fields("Concepto")
          TFA.Total_MN = TFA.Sin_IVA + TFA.Con_IVA + TFA.Total_IVA
          TFA.SubTotal = TFA.Sin_IVA + TFA.Con_IVA
          If TFA.Nota = Ninguno Then TFA.Nota = "Liquidacion de Compras"
          tipoCodComprobante = Format$(.fields("TipoComprobante"), "00")
         'Validar_Porc_IVA TFA.Fecha
          Obtener_Porc_IVA TFA.Fecha, .fields("PorcentajeIva")
          TFA.Porc_IVA = Porc_IVA * 100
          
         'Algoritmo Modulo 11 para la clave de la retencion
         '& Format$(TFA.Vencimiento, "ddmmyyyy")
          TFA.ClaveAcceso_LC = Format$(TFA.Fecha, "ddmmyyyy") & "03" & RUC & Ambiente & TFA.Serie_LC _
                             & Format$(TFA.Factura, String(9, "0")) & "123456781"
          TFA.ClaveAcceso_LC = Replace(TFA.ClaveAcceso_LC, ".", "1")
          TFA.ClaveAcceso_LC = TFA.ClaveAcceso_LC & Digito_Verificador_Modulo11(TFA.ClaveAcceso_LC)
         'MsgBox TFA.Autorizacion_LC
          If Len(TFA.Autorizacion_LC) >= 13 Then
            'ENCABEZADO XML PARA EL SRI DE LA RETENCION
             Insertar_Campo_XML "<?xml version=""1.0"" encoding=""UTF-8""?>"
             Insertar_Campo_XML "<liquidacionCompra id=""comprobante"" version=""1.0.0"">"
             Insertar_Campo_XML AbrirXML("infoTributaria")
                Insertar_Campo_XML CampoXML("ambiente", Ambiente)
                Insertar_Campo_XML CampoXML("tipoEmision", "1")
                Insertar_Campo_XML CampoXML("razonSocial", RazonSocial)
                Insertar_Campo_XML CampoXML("nombreComercial", NombreComercial)
                Insertar_Campo_XML CampoXML("ruc", RUC)
                Insertar_Campo_XML CampoXML("claveAcceso", TFA.ClaveAcceso_LC)
                Insertar_Campo_XML CampoXML("codDoc", "03")
                Insertar_Campo_XML CampoXML("estab", MidStrg(TFA.Serie_LC, 1, 3))
                Insertar_Campo_XML CampoXML("ptoEmi", MidStrg(TFA.Serie_LC, 4, 3))
                Insertar_Campo_XML CampoXML("secuencial", Format$(TFA.Factura, String(9, "0")))
                Insertar_Campo_XML CampoXML("dirMatriz", ULCase(Direccion))
               'Tipo de Contribuyente
                If AgenteRetencion <> Ninguno Then Insertar_Campo_XML CampoXML("agenteRetencion", "1")
                If MicroEmpresa = "CONTRIBUYENTE RÉGIMEN RIMPE" Then Insertar_Campo_XML CampoXML("contribuyenteRimpe", "CONTRIBUYENTE RÉGIMEN RIMPE")
             Insertar_Campo_XML CerrarXML("infoTributaria")
             
             Insertar_Campo_XML AbrirXML("infoLiquidacionCompra")
                Insertar_Campo_XML CampoXML("fechaEmision", TFA.Fecha)
                Insertar_Campo_XML CampoXML("dirEstablecimiento", ULCase(DireccionEstab))
                If Len(ContEspec) > 1 Then Insertar_Campo_XML CampoXML("contribuyenteEspecial", ContEspec)
                Insertar_Campo_XML CampoXML("obligadoContabilidad", Obligado_Conta)
                Select Case TFA.TD
                  Case "C": TFA.TD = "05"
                  Case "P": TFA.TD = "06"
                  Case Else: TFA.TD = "05"
                End Select
                Insertar_Campo_XML CampoXML("tipoIdentificacionProveedor", TFA.TD)
                Insertar_Campo_XML CampoXML("razonSocialProveedor", TFA.Cliente)
                Insertar_Campo_XML CampoXML("identificacionProveedor", TFA.CI_RUC)
                Insertar_Campo_XML CampoXML("direccionProveedor", TFA.DireccionC)
                 
                Insertar_Campo_XML CampoXML("totalSinImpuestos", TFA.SubTotal)
                Insertar_Campo_XML CampoXML("totalDescuento", "0.00")
                
''                If tipoCodComprobante = "41" Then
''                   Insertar_Campo_XML CampoXML("codDocReembolso", tipoCodComprobante)
''                   Insertar_Campo_XML CampoXML("totalComprobantesReembolso", TFA.Total_MN)
''                   Insertar_Campo_XML CampoXML("totalBaseImponibleReembolso", TFA.SubTotal)
''                   Insertar_Campo_XML CampoXML("totalImpuestoReembolso", TFA.Total_IVA)
''                End If
                
                Insertar_Campo_XML AbrirXML("totalConImpuestos")
                   Insertar_Campo_XML AbrirXML("totalImpuesto")
                      Insertar_Campo_XML CampoXML("codigo", "2")
                      Insertar_Campo_XML CampoXML("codigoPorcentaje", "0")
                      Insertar_Campo_XML CampoXML("baseImponible", TFA.Sin_IVA)
                      Insertar_Campo_XML CampoXML("tarifa", "0.00")
                      Insertar_Campo_XML CampoXML("valor", "0.00")
                   Insertar_Campo_XML CerrarXML("totalImpuesto")
                   
                   Insertar_Campo_XML AbrirXML("totalImpuesto")
                      Insertar_Campo_XML CampoXML("codigo", "2")
                      Insertar_Campo_XML CampoXML("codigoPorcentaje", .fields("PorcentajeIva"))
                      Insertar_Campo_XML CampoXML("baseImponible", TFA.Con_IVA)
                      Insertar_Campo_XML CampoXML("tarifa", TFA.Porc_IVA)
                      Insertar_Campo_XML CampoXML("valor", TFA.Total_IVA)
                   Insertar_Campo_XML CerrarXML("totalImpuesto")
                Insertar_Campo_XML CerrarXML("totalConImpuestos")
                
                Insertar_Campo_XML CampoXML("importeTotal", TFA.Total_MN)
                Insertar_Campo_XML CampoXML("moneda", "DOLAR")
                Insertar_Campo_XML AbrirXML("pagos")
                   Insertar_Campo_XML AbrirXML("pago")
                      Insertar_Campo_XML CampoXML("formaPago", "20")
                      Insertar_Campo_XML CampoXML("total", TFA.Total_MN)
                      Insertar_Campo_XML CampoXML("plazo", "0")
                      Insertar_Campo_XML CampoXML("unidadTiempo", "dias")
                   Insertar_Campo_XML CerrarXML("pago")
                Insertar_Campo_XML CerrarXML("pagos")
             Insertar_Campo_XML CerrarXML("infoLiquidacionCompra")
             
             Insertar_Campo_XML AbrirXML("detalles")
                Insertar_Campo_XML AbrirXML("detalle")
                   Insertar_Campo_XML CampoXML("codigoPrincipal", "99")
                   Insertar_Campo_XML CampoXML("codigoAuxiliar", "99.98")
                   Insertar_Campo_XML CampoXML("descripcion", TFA.Nota & ", Tarifa 0%")
                   Insertar_Campo_XML CampoXML("cantidad", "1")
                   Insertar_Campo_XML CampoXML("precioUnitario", TFA.Sin_IVA)
                   Insertar_Campo_XML CampoXML("descuento", "0.00")
                   Insertar_Campo_XML CampoXML("precioTotalSinImpuesto", TFA.Sin_IVA)
                   Insertar_Campo_XML AbrirXML("impuestos")
                      Insertar_Campo_XML AbrirXML("impuesto")
                         Insertar_Campo_XML CampoXML("codigo", "2")
                         Insertar_Campo_XML CampoXML("codigoPorcentaje", "0")
                         Insertar_Campo_XML CampoXML("tarifa", "0.00")
                         Insertar_Campo_XML CampoXML("baseImponible", TFA.Sin_IVA)
                         Insertar_Campo_XML CampoXML("valor", "0.00")
                      Insertar_Campo_XML CerrarXML("impuesto")
                   Insertar_Campo_XML CerrarXML("impuestos")
                Insertar_Campo_XML CerrarXML("detalle")
                
                Insertar_Campo_XML AbrirXML("detalle")
                   Insertar_Campo_XML CampoXML("codigoPrincipal", "99")
                   Insertar_Campo_XML CampoXML("codigoAuxiliar", "99.99")
                   Insertar_Campo_XML CampoXML("descripcion", TFA.Nota & ", Tarifa " & TFA.Porc_IVA & "%")
                   Insertar_Campo_XML CampoXML("cantidad", "1")
                   Insertar_Campo_XML CampoXML("precioUnitario", TFA.Con_IVA)
                   Insertar_Campo_XML CampoXML("descuento", "0.00")
                   Insertar_Campo_XML CampoXML("precioTotalSinImpuesto", TFA.Con_IVA)
                   Insertar_Campo_XML AbrirXML("impuestos")
                      Insertar_Campo_XML AbrirXML("impuesto")
                         Insertar_Campo_XML CampoXML("codigo", "2")
                         Insertar_Campo_XML CampoXML("codigoPorcentaje", .fields("PorcentajeIva"))
                         Insertar_Campo_XML CampoXML("tarifa", TFA.Porc_IVA)
                         Insertar_Campo_XML CampoXML("baseImponible", TFA.Con_IVA)
                         Insertar_Campo_XML CampoXML("valor", TFA.Total_IVA)
                      Insertar_Campo_XML CerrarXML("impuesto")
                   Insertar_Campo_XML CerrarXML("impuestos")
                Insertar_Campo_XML CerrarXML("detalle")
             Insertar_Campo_XML CerrarXML("detalles")
             
            'FIN DE XML DE LIQUIDACION DE COMPRAS
             Insertar_Campo_XML AbrirXML("infoAdicional")
                If Len(TFA.DireccionC) > 1 Then Insertar_Campo_XML "<campoAdicional nombre=""Direccion"">" & TFA.DireccionC & "</campoAdicional>"
                If Len(TFA.TelefonoC) > 1 Then Insertar_Campo_XML "<campoAdicional nombre=""Telefono"">" & TFA.TelefonoC & "</campoAdicional>"
                If Len(TFA.EmailC) > 1 Then Insertar_Campo_XML "<campoAdicional nombre=""Email"">" & TFA.EmailC & "</campoAdicional>"
                Insertar_Campo_XML "<campoAdicional nombre=""Comprobante No"">" & TFA.TP & "-" & Format$(TFA.Numero, "0000000000") & "</campoAdicional>"
             Insertar_Campo_XML CerrarXML("infoAdicional")
             Insertar_Campo_XML CerrarXML("liquidacionCompra")
             If CFechaLong(TFA.Fecha) <= CFechaLong(Fecha_CE) Then
                'Grabamos el comprobante XML de la Retencion
                 RutaGeneraFile = RutaDocumentos & "\Comprobantes Generados\" & TFA.ClaveAcceso_LC & ".xml"
                 
                 If GeneraXML Then
                    Clipboard.Clear
                    Clipboard.SetText TextoXML
                    Grabar_Consulta_Archivo "LC_" & TFA.Serie_LC & "-" & Format$(TFA.Factura, String(9, "0")), TextoXML
                 End If
                 
                 Set DocumentoXML = New DOMDocument30
                 DocumentoXML.loadXML (TextoXML)
                 DocumentoXML.save (RutaGeneraFile)
                 
                     TFA.Estado_SRI = "CG"
                    'Enviamos al SRI para que me autorice la Retencion
                     SRI_Autorizacion = SRI_Generar_XML(TFA.ClaveAcceso_LC, TFA.Estado_SRI_LC)
                     SRI_Actualizar_XML_Liquidacion SRI_Autorizacion, FA
                     RatonReloj
                     If SRI_Autorizacion.Estado_SRI = "OK" Then
                        Control_Procesos "R", "LC-" & TFA.Serie_LC & " No. " & Format(TFA.Autorizacion, "000000000") & " Autorizada"
                        SRI_Actualizar_Autorizacion_Liquidacion SRI_Autorizacion, FA
                        SRI_Actualizar_Documento_XML TFA.ClaveAcceso_LC
                        
                        SRI_Enviar_Mails FA, SRI_Autorizacion, "LC"
                        SRI_Generar_PDF_LC TFA, VerLiquidacion
                        RatonNormal
                     Else
                        RatonNormal
                        If TextoImprimio = "" Then TextoImprimio = "INFORME DE ERRORES EN LAS LIQUIDACIONES DE COMPRAS:" & vbCrLf _
                                                                 & "--------------------------------------------------" & vbCrLf
                        TextoImprimio = TextoImprimio _
                                      & "Fecha: " & TFA.Fecha & ", Retencion No. " & TFA.Serie_R & "-" & TFA.Retencion _
                                      & ", Documento No. " & TFA.Serie & "-" & TFA.Factura _
                                      & ", De: " & TFA.Cliente & ", CI/RUC: " & TFA.CI_RUC & vbCrLf _
                                      & SRI_Autorizacion.Estado_SRI & " - " & SRI_Autorizacion.Error_SRI & vbCrLf _
                                      & String(80, "-") & vbCrLf
                        If VerLiquidacion Then MsgBox SRI_Autorizacion.Error_SRI
                     End If
                 
             Else
                RatonNormal
                MsgBox MensajeNoAutorizarCE
             End If
          End If
         .MoveNext
      Loop
   End If
  End With
     AdoCompras.Close
  End If
  RatonNormal
End Sub

Public Sub SRI_Crear_Clave_Acceso_Facturas(TFA As Tipo_Facturas, _
                                           VerFactura As Boolean, _
                                           Optional GeneraXML As Boolean, _
                                           Optional Autorizar As Boolean)
Dim AdoDBFA As ADODB.Recordset
Dim AdoDBDet As ADODB.Recordset
Dim AdoDBCli As ADODB.Recordset
Dim AdoDBProd As ADODB.Recordset
Dim NumFile As Integer
Dim RutaGeneraFile As String

Dim TipoIdent As String
Dim TipoProvReemb As String
Dim Cod_Aux As String
Dim Cod_Bar As String
Dim SubTotalDesc As Currency
Dim Autorizar_XML As Boolean
Dim Serie1Reembolo As String
Dim Serie2Reembolo As String
Dim SecuencialReembolo As String

    RatonReloj
    Autorizar_XML = True
    If Len(Fecha_Igualar) = 10 Then
       If CFechaLong(TFA.Fecha) < CFechaLong(Fecha_Igualar) Then Autorizar_XML = False
    End If
    TextoXML = ""
    
    'MsgBox Autorizar_XML
    If Autorizar_XML Then
    
    Leer_Datos_FA_NV TFA
        
   'Detalle de descuentos
    sSQL = "SELECT DF.*,CP.Reg_Sanitario, CP.Marca, CP.Desc_Item, CP.Codigo_Barra " _
         & "FROM Detalle_Factura As DF, Catalogo_Productos As CP " _
         & "WHERE DF.Item = '" & NumEmpresa & "' " _
         & "AND DF.Periodo = '" & Periodo_Contable & "' " _
         & "AND DF.TC = '" & TFA.TC & "' " _
         & "AND DF.Serie = '" & TFA.Serie & "' " _
         & "AND DF.Autorizacion = '" & TFA.Autorizacion & "' " _
         & "AND DF.Factura = " & TFA.Factura & " " _
         & "AND LEN(DF.Autorizacion) >= 13 " _
         & "AND DF.T <> 'A' " _
         & "AND DF.Item = CP.Item " _
         & "AND DF.Periodo = CP.Periodo " _
         & "AND DF.Codigo = CP.Codigo_Inv " _
         & "ORDER BY DF.ID,DF.Codigo "
    Select_AdoDB AdoDBDet, sSQL
   
   'Encabezado de la Factura
    sSQL = "SELECT T, TDT, SP, Porc_IVA, Imp_Mes, Fecha, Vencimiento, SubTotal, Sin_IVA, Con_IVA, IVA, Total_MN, Razon_Social, RUC_CI, TB, Descuento, Descuento2, Servicio " _
         & "FROM Facturas " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND TC = '" & TFA.TC & "' " _
         & "AND Serie = '" & TFA.Serie & "' " _
         & "AND Autorizacion = '" & TFA.Autorizacion & "' " _
         & "AND Factura = " & TFA.Factura & " " _
         & "AND LEN(Autorizacion) = 13 " _
         & "AND T <> 'A' "
    Select_AdoDB AdoDBFA, sSQL
    RatonReloj
    With AdoDBFA
     If .RecordCount > 0 Then
         Autorizar_XML = True
         TFA.T = .fields("T")
         TFA.SP = .fields("SP")
         TFA.TDT = .fields("TDT")
         TFA.Porc_IVA = .fields("Porc_IVA")
         TFA.Imp_Mes = .fields("Imp_Mes")
         TFA.Fecha = .fields("Fecha")
         TFA.Vencimiento = .fields("Vencimiento")
         TFA.SubTotal = .fields("SubTotal")
         TFA.Sin_IVA = .fields("Sin_IVA")
         TFA.Con_IVA = .fields("Con_IVA")
         TFA.Total_IVA = .fields("IVA")
         TFA.Servicio = .fields("Servicio")
         TFA.Total_MN = .fields("Total_MN")
         TFA.Razon_Social = .fields("Razon_Social")
         TFA.RUC_CI = .fields("RUC_CI")
         TFA.TB = .fields("TB")
         TFA.Descuento = .fields("Descuento")
         TFA.Descuento2 = .fields("Descuento2")
         TFA.Total_Descuento = TFA.Descuento + TFA.Descuento2
         
         If TFA.TDT = 41 Then TFA.EsPorReembolso = True
         
        'MsgBox "Validar Porc IVA"
         
         Obtener_Cod_Porc_IVA TFA.Fecha, (TFA.Porc_IVA * 100)
        'MsgBox TFA.Porc_IVA & vbCrLf & Cod_Porc_IVA
         
        'Generamos la Clave de acceso
        '& Format$(TFA.Vencimiento, "ddmmyyyy")
         If Len(TFA.Autorizacion) >= 13 Then
            TFA.ClaveAcceso = Format$(TFA.Fecha, "ddmmyyyy") & "01" & RUC & Ambiente & TFA.Serie & Format$(TFA.Factura, String(9, "0")) _
            & "123456781"
            TFA.ClaveAcceso = TFA.ClaveAcceso & Digito_Verificador_Modulo11(TFA.ClaveAcceso)
         Else
            TFA.ClaveAcceso = Ninguno
         End If
         TFA.Hora = Format$(Time, FormatoTimes)
         SRI_Autorizacion.Clave_De_Acceso = TFA.ClaveAcceso
         TipoIdent = "P"
         Select Case TFA.TB
           Case "R": If TFA.CI_RUC = String(13, "9") Then TipoIdent = "07" Else TipoIdent = "04"
           Case "C": TipoIdent = "05"
           Case "P": TipoIdent = "06"
           Case Else: TipoIdent = "07"
         End Select
         
         If MidStrg(TFA.CI_RUC, 3, 1) = "9" Then TipoProvReemb = "02" Else TipoProvReemb = "01"
         
         sSQL = "UPDATE Facturas " _
              & "SET Clave_Acceso = '" & TFA.ClaveAcceso & "' " _
              & "WHERE Item = '" & NumEmpresa & "' " _
              & "AND Periodo = '" & Periodo_Contable & "' " _
              & "AND TC = '" & TFA.TC & "' " _
              & "AND Serie = '" & TFA.Serie & "' " _
              & "AND Factura = " & TFA.Factura & " " _
              & "AND CodigoC = '" & TFA.CodigoC & "' " _
              & "AND Autorizacion = '" & TFA.Autorizacion & "' "
         Ejecutar_SQL_SP sSQL
         
        'ENCABEZADO XML PARA EL SRI DE LA FACTURA/NOTA DE VENTA
        'standalone=""yes""
         Insertar_Campo_XML "<?xml version=""1.0"" encoding=""UTF-8""?>"
         Select Case TFA.TC
           Case "FA": Insertar_Campo_XML "<factura id=""comprobante"" version=""1.1.0"">"
           Case "NV": Insertar_Campo_XML AbrirXML("<notaVenta>")
           Case Else: Insertar_Campo_XML AbrirXML("<puntoVenta>")
         End Select
         
         Insertar_Campo_XML AbrirXML("infoTributaria")
            Insertar_Campo_XML CampoXML("ambiente", Ambiente)
            Insertar_Campo_XML CampoXML("tipoEmision", "1")
            Insertar_Campo_XML CampoXML("razonSocial", RazonSocial)
            Insertar_Campo_XML CampoXML("nombreComercial", NombreComercial)
            Insertar_Campo_XML CampoXML("ruc", RUC)
            Insertar_Campo_XML CampoXML("claveAcceso", TFA.ClaveAcceso)
           ' If TFA.EsPorReembolso Then
           '    Insertar_Campo_XML CampoXML("codDoc", "41")
           ' Else
                Select Case TFA.TC
                  Case "FA": Insertar_Campo_XML CampoXML("codDoc", "01")
                  Case "NV": Insertar_Campo_XML CampoXML("codDoc", "02")
                  Case Else: Insertar_Campo_XML CampoXML("codDoc", "00")
                End Select
           ' End If
            Insertar_Campo_XML CampoXML("estab", MidStrg(TFA.Serie, 1, 3))
            Insertar_Campo_XML CampoXML("ptoEmi", MidStrg(TFA.Serie, 4, 3))
            Insertar_Campo_XML CampoXML("secuencial", Format$(TFA.Factura, String(9, "0")))
            Insertar_Campo_XML CampoXML("dirMatriz", ULCase(Direccion))
           'Tipo de Contribuyente
            'MsgBox "..."
            If AgenteRetencion <> Ninguno Then Insertar_Campo_XML CampoXML("agenteRetencion", "00000001")
            If MicroEmpresa = "CONTRIBUYENTE RÉGIMEN RIMPE" Then Insertar_Campo_XML CampoXML("contribuyenteRimpe", "CONTRIBUYENTE RÉGIMEN RIMPE")
         Insertar_Campo_XML CerrarXML("infoTributaria")
         
         Insertar_Campo_XML AbrirXML("infoFactura")
            Insertar_Campo_XML CampoXML("fechaEmision", TFA.Fecha)
            Insertar_Campo_XML CampoXML("dirEstablecimiento", ULCase(DireccionEstab))
            
            If Len(ContEspec) > 1 Then Insertar_Campo_XML CampoXML("contribuyenteEspecial", ContEspec)
            
            Insertar_Campo_XML CampoXML("obligadoContabilidad", Obligado_Conta)
            Insertar_Campo_XML CampoXML("tipoIdentificacionComprador", TipoIdent)
            Insertar_Campo_XML CampoXML("razonSocialComprador", TFA.Razon_Social)
            Insertar_Campo_XML CampoXML("identificacionComprador", TFA.RUC_CI)
            Insertar_Campo_XML CampoXML("direccionComprador", TFA.DireccionC)
            If TFA.EsPorReembolso Then
               Insertar_Campo_XML CampoXML("totalSinImpuestos", Format$(TFA.Total_MN, "#0.00"))
            Else
               Insertar_Campo_XML CampoXML("totalSinImpuestos", Format$(TFA.Sin_IVA + TFA.Con_IVA - TFA.Total_Descuento, "#0.00"))
            End If
            Insertar_Campo_XML CampoXML("totalDescuento", Format$(TFA.Total_Descuento, "#0.00"))
            If TFA.EsPorReembolso Then
               Insertar_Campo_XML CampoXML("codDocReembolso", "41")
               Insertar_Campo_XML CampoXML("totalComprobantesReembolso", Format$(TFA.Total_MN, "#0.00"))
               Insertar_Campo_XML CampoXML("totalBaseImponibleReembolso", Format$(TFA.SubTotal, "#0.00"))
               Insertar_Campo_XML CampoXML("totalImpuestoReembolso", Format$(TFA.Total_IVA, "#0.00"))
            End If
            Insertar_Campo_XML AbrirXML("totalConImpuestos")
                Insertar_Campo_XML AbrirXML("totalImpuesto")
                 If TFA.EsPorReembolso Then
                    Insertar_Campo_XML CampoXML("codigo", "2")
                    Insertar_Campo_XML CampoXML("codigoPorcentaje", "0")
                    Insertar_Campo_XML CampoXML("baseImponible", Format$(TFA.Total_MN, "#0.00"))
                    Insertar_Campo_XML CampoXML("valor", "0.00")
                 Else
                    Insertar_Campo_XML CampoXML("codigo", "2")
                    Insertar_Campo_XML CampoXML("codigoPorcentaje", "0")
                    Insertar_Campo_XML CampoXML("descuentoAdicional", "0")
                    Insertar_Campo_XML CampoXML("baseImponible", Format$(TFA.Sin_IVA - TFA.Descuento_0, "#0.00"))
                   'Insertar_Campo_XML CampoXML("tarifa", "0.00")
                    Insertar_Campo_XML CampoXML("valor", "0.00")
                 End If
                Insertar_Campo_XML CerrarXML("totalImpuesto")
               'MsgBox "..."
                If TFA.Total_IVA > 0 And Not TFA.EsPorReembolso Then
                   Insertar_Campo_XML AbrirXML("totalImpuesto")
                       Insertar_Campo_XML CampoXML("codigo", "2")
                      'MsgBox Porc_IVA
                       Insertar_Campo_XML CampoXML("codigoPorcentaje", Cod_Porc_IVA)
                       'Insertar_Campo_XML CampoXML("descuentoAdicional", "0")
                       Insertar_Campo_XML CampoXML("baseImponible", Format$(TFA.Con_IVA - TFA.Descuento_X, "#0.00"))
                       Insertar_Campo_XML CampoXML("tarifa", Porc_IVA * 100)
                       Insertar_Campo_XML CampoXML("valor", Format$(TFA.Total_IVA, "#0.00"))
                   Insertar_Campo_XML CerrarXML("totalImpuesto")
                End If
            Insertar_Campo_XML CerrarXML("totalConImpuestos")
            
            Insertar_Campo_XML CampoXML("propina", Format$(TFA.Servicio, "#0.00"))
            Insertar_Campo_XML CampoXML("importeTotal", Format$(TFA.Total_MN, "#0.00"))
            Insertar_Campo_XML CampoXML("moneda", "DOLAR")
            Insertar_Campo_XML AbrirXML("pagos")
                Insertar_Campo_XML AbrirXML("pago")
                    Insertar_Campo_XML CampoXML("formaPago", TFA.Tipo_Pago)
                    Insertar_Campo_XML CampoXML("total", TFA.Total_MN)
                    If Val(TFA.Tipo_Pago) = 19 Then
                       Insertar_Campo_XML CampoXML("plazo", "60")
                       Insertar_Campo_XML CampoXML("unidadTiempo", "dias")
                    End If
                Insertar_Campo_XML CerrarXML("pago")
            Insertar_Campo_XML CerrarXML("pagos")
         Insertar_Campo_XML CerrarXML("infoFactura")
     End If
    End With
   '-----------------------------------
   'Detalle de la Factura/Nota de Venta
   '-----------------------------------
    RatonReloj
    With AdoDBDet
     If .RecordCount > 0 Then
         Insertar_Campo_XML AbrirXML("detalles")
         Do While Not .EOF
            Producto = .fields("Producto")
            Cod_Aux = .fields("Desc_Item")
            Cod_Bar = .fields("Codigo_Barra")
            SubTotal = (.fields("Cantidad") * .fields("Precio")) - (.fields("Total_Desc") + .fields("Total_Desc2"))
            If TFA.EsPorReembolso Then
               Cod_Aux = "Reembolso de Gastos"
               If Len(.fields("Tipo_Hab")) > 1 Then Cod_Aux = Cod_Aux & " por " & .fields("Tipo_Hab")
               SubTotal = SubTotal + .fields("Total_IVA")
               Insertar_Campo_XML AbrirXML("detalle")
                    Insertar_Campo_XML CampoXML("codigoPrincipal", .fields("Codigo"))
                    Insertar_Campo_XML CampoXML("codigoAuxiliar", TrimStrg(MidStrg(.fields("Ruta"), 1, 10)))
                    Insertar_Campo_XML CampoXML("descripcion", .fields("Producto"))
                    Insertar_Campo_XML CampoXML("cantidad", "1.000000")
                    Insertar_Campo_XML CampoXML("precioUnitario", Format$(SubTotal, "#0.000000"))
                    Insertar_Campo_XML CampoXML("descuento", "0.00")
                    Insertar_Campo_XML CampoXML("precioTotalSinImpuesto", Format$(SubTotal, "#0.00"))
                    Insertar_Campo_XML AbrirXML("detallesAdicionales")
                        Insertar_Campo_XML "<detAdicional nombre=""RUC Factura"" valor = """ & .fields("Ruta") & """/>"
                        Insertar_Campo_XML "<detAdicional nombre=""Serie y Factura"" valor = """ & .fields("Serie_No") & """/>"
                        Insertar_Campo_XML "<detAdicional nombre=""Descripcion Reembolso"" valor = """ & Cod_Aux & """/>"
                    Insertar_Campo_XML CerrarXML("detallesAdicionales")
                    Insertar_Campo_XML AbrirXML("impuestos")
                        Insertar_Campo_XML AbrirXML("impuesto")
                           Insertar_Campo_XML CampoXML("codigo", "2")
                           Insertar_Campo_XML CampoXML("codigoPorcentaje", "0")
                           Insertar_Campo_XML CampoXML("tarifa", 0) ' POR SI ACASO
                           Insertar_Campo_XML CampoXML("baseImponible", Format$(SubTotal, "#0.00"))
                           Insertar_Campo_XML CampoXML("valor", "0.00")
                       Insertar_Campo_XML CerrarXML("impuesto")
                    Insertar_Campo_XML CerrarXML("impuestos")
               Insertar_Campo_XML CerrarXML("detalle")
            Else
                If TFA.Imp_Mes Then Producto = Producto & ", " & .fields("Ticket") & ": " & .fields("Mes") & " "
                If TFA.SP Then
                   Producto = Producto _
                            & ", Lote No. " & .fields("Lote_No") _
                            & ", ELAB. " & .fields("Fecha_Fab") _
                            & ", VENC. " & .fields("Fecha_Exp") _
                            & ", Reg. Sanit. " & .fields("Reg_Sanitario") _
                            & ", Modelo: " & .fields("Modelo") _
                            & ", Procedencia: " & .fields("Procedencia")
                End If
        '''            If Len(.Fields("Serie_No")) > 1 Then
        '''               Producto = Producto & ", Serie No. " & .Fields("Serie_No")
        '''            End If
                Insertar_Campo_XML AbrirXML("detalle")
                If TFA.SP Then
                   If Len(Cod_Bar) > 1 Then Insertar_Campo_XML CampoXML("codigoPrincipal", Cod_Bar)
                   If Len(Cod_Aux) > 1 Then
                      Insertar_Campo_XML CampoXML("codigoAuxiliar", Cod_Aux)
                   Else
                      Insertar_Campo_XML CampoXML("codigoAuxiliar", .fields("Codigo"))
                   End If
                Else
                   If Len(Cod_Aux) > 1 Then
                      Insertar_Campo_XML CampoXML("codigoPrincipal", Cod_Aux)
                   Else
                      Insertar_Campo_XML CampoXML("codigoPrincipal", .fields("Codigo"))
                   End If
                   If Len(Cod_Bar) > 1 Then Insertar_Campo_XML CampoXML("codigoAuxiliar", Cod_Bar)
                End If
                Insertar_Campo_XML CampoXML("descripcion", Producto)
                Insertar_Campo_XML CampoXML("unidadMedida", "DOLAR")
                Insertar_Campo_XML CampoXML("cantidad", Format$(.fields("Cantidad"), "#0.000000"))
                Insertar_Campo_XML CampoXML("precioUnitario", Format$(.fields("Precio"), "#0.000000"))
                Insertar_Campo_XML CampoXML("descuento", Format$(.fields("Total_Desc") + .fields("Total_Desc2"), "#0.00"))
                Insertar_Campo_XML CampoXML("precioTotalSinImpuesto", Format$(SubTotal, "#0.00"))
                
                If Len(.fields("Serie_No")) > 1 Then
                   Insertar_Campo_XML AbrirXML("detallesAdicionales")
                       Insertar_Campo_XML "<detAdicional nombre=""Serie_No"" valor=""" & .fields("Serie_No") & """/>"
                   Insertar_Campo_XML CerrarXML("detallesAdicionales")
                End If
                    Insertar_Campo_XML AbrirXML("impuestos")
                        Insertar_Campo_XML AbrirXML("impuesto")
                           Insertar_Campo_XML CampoXML("codigo", "2")
                           If .fields("Total_IVA") = 0 Then
                               Insertar_Campo_XML CampoXML("codigoPorcentaje", "0")
                               Insertar_Campo_XML CampoXML("tarifa", "0")
                           Else
                               Insertar_Campo_XML CampoXML("codigoPorcentaje", Cod_Porc_IVA)
                               Insertar_Campo_XML CampoXML("tarifa", Porc_IVA * 100)
                           End If
                           Insertar_Campo_XML CampoXML("baseImponible", Format$(.fields("Total") - (.fields("Total_Desc") + .fields("Total_Desc2")), "#0.00"))
                           Insertar_Campo_XML CampoXML("valor", Format$(.fields("Total_IVA"), "#0.00"))
                       Insertar_Campo_XML CerrarXML("impuesto")
                    Insertar_Campo_XML CerrarXML("impuestos")
                Insertar_Campo_XML CerrarXML("detalle")
            End If
           .MoveNext
         Loop
         Insertar_Campo_XML CerrarXML("detalles")
     End If
    End With
   '--------------------------------
   'Detalle del Reembolso de Gastos
   '--------------------------------
    If TFA.EsPorReembolso Then
       With AdoDBDet
        If .RecordCount > 0 Then
           .MoveFirst
            Insertar_Campo_XML AbrirXML("reembolsos")
            Do While Not .EOF
               If .fields("Ruta") = String(13, "9") Then TipoIdent = "07" Else TipoIdent = "04"
               If MidStrg(.fields("Ruta"), 3, 1) = "9" Then TipoProvReemb = "02" Else TipoProvReemb = "01"
               Serie1Reembolo = MidStrg(.fields("Serie_No"), 1, 3)
               Serie2Reembolo = MidStrg(.fields("Serie_No"), 4, 3)
               SecuencialReembolo = MidStrg(.fields("Serie_No"), 8, 9)
               SubTotal = (.fields("Cantidad") * .fields("Precio")) - (.fields("Total_Desc") + .fields("Total_Desc2"))
               Insertar_Campo_XML AbrirXML("reembolsoDetalle")
                    Insertar_Campo_XML CampoXML("tipoIdentificacionProveedorReembolso", TipoIdent)
                    Insertar_Campo_XML CampoXML("identificacionProveedorReembolso", .fields("Ruta"))
                    Insertar_Campo_XML CampoXML("codPaisPagoProveedorReembolso", "593")
                    Insertar_Campo_XML CampoXML("tipoProveedorReembolso", TipoProvReemb)
                    Insertar_Campo_XML CampoXML("codDocReembolso", .fields("Lote_No"))
                    Insertar_Campo_XML CampoXML("estabDocReembolso", Serie1Reembolo)
                    Insertar_Campo_XML CampoXML("ptoEmiDocReembolso", Serie2Reembolo)
                    Insertar_Campo_XML CampoXML("secuencialDocReembolso", SecuencialReembolo)
                    Insertar_Campo_XML CampoXML("fechaEmisionDocReembolso", TFA.Fecha)
                    Insertar_Campo_XML CampoXML("numeroautorizacionDocReemb", .fields("Procedencia")) 'String(10, "9"))
                    Insertar_Campo_XML AbrirXML("detalleImpuestos")
                       Insertar_Campo_XML AbrirXML("detalleImpuesto")
                          Insertar_Campo_XML CampoXML("codigo", "2")
                          If .fields("Total_IVA") > 0 Then
                             Insertar_Campo_XML CampoXML("codigoPorcentaje", Cod_Porc_IVA)
                             Insertar_Campo_XML CampoXML("tarifa", "12")
                             Insertar_Campo_XML CampoXML("baseImponibleReembolso", SubTotal)
                             Insertar_Campo_XML CampoXML("impuestoReembolso", .fields("Total_IVA"))
                          Else
                             Insertar_Campo_XML CampoXML("codigoPorcentaje", "0")
                             Insertar_Campo_XML CampoXML("tarifa", "0")
                             Insertar_Campo_XML CampoXML("baseImponibleReembolso", SubTotal)
                             Insertar_Campo_XML CampoXML("impuestoReembolso", "0.00")
                          End If
                       Insertar_Campo_XML CerrarXML("detalleImpuesto")
                    Insertar_Campo_XML CerrarXML("detalleImpuestos")
               Insertar_Campo_XML CerrarXML("reembolsoDetalle")
              .MoveNext
            Loop
            Insertar_Campo_XML CerrarXML("reembolsos")
        End If
       End With
    End If
        
     Insertar_Campo_XML AbrirXML("infoAdicional")
       If TFA.Cliente <> Ninguno And TFA.Razon_Social <> TFA.Cliente Then
          If Len(TFA.Cliente) > 1 Then Insertar_Campo_XML "<campoAdicional nombre=""Beneficiario"">" & TFA.Cliente & "</campoAdicional>"
          If Len(TFA.CI_RUC) > 1 Then Insertar_Campo_XML "<campoAdicional nombre=""Codigo"">" & TFA.CI_RUC & "</campoAdicional>"
          If Len(TFA.Curso) > 1 Then Insertar_Campo_XML "<campoAdicional nombre=""Ubicacion"">" & TFA.Grupo & "-" & TFA.Curso & "</campoAdicional>"
       End If
       If Len(TFA.DireccionC) > 1 Then Insertar_Campo_XML "<campoAdicional nombre=""Direccion"">" & TFA.DireccionC & "</campoAdicional>"
       If Len(TFA.TelefonoC) > 1 Then Insertar_Campo_XML "<campoAdicional nombre=""Telefono"">" & TFA.TelefonoC & "</campoAdicional>"
       If Len(TFA.EmailC) > 1 And InStr(TFA.EmailC, "@") > 0 Then
          Insertar_Campo_XML "<campoAdicional nombre=""Email"">" & TFA.EmailC & "</campoAdicional>"
       End If
       If Len(TFA.EmailR) > 1 And InStr(TFA.EmailR, "@") > 0 And InStr(TFA.EmailC, TFA.EmailR) = 0 Then
          Insertar_Campo_XML "<campoAdicional nombre=""Email2"">" & TFA.EmailR & "</campoAdicional>"
       End If
       If Len(TFA.Contacto) > 1 Then Insertar_Campo_XML "<campoAdicional nombre=""Referencia"">" & TFA.Contacto & "</campoAdicional>"
       If Val(TFA.Orden_Compra) > 0 Then Insertar_Campo_XML "<campoAdicional nombre=""ordenCompra"">" & TFA.Orden_Compra & "</campoAdicional>"
     Insertar_Campo_XML CerrarXML("infoAdicional")
    ' MsgBox MicroEmpresa
    'Fin del Archivo Xml
     Select Case TFA.TC
       Case "FA": Insertar_Campo_XML CerrarXML("factura")
       Case "NV": Insertar_Campo_XML CerrarXML("notaVenta")
       Case Else: Insertar_Campo_XML CerrarXML("puntoVenta")
     End Select
        
      'Grabamos el comprobante XML
       RatonReloj
       If Len(TFA.Autorizacion) = 13 Then
          'MsgBox RutaDocumentos
          If CFechaLong(TFA.Fecha) <= CFechaLong(Fecha_CE) Then
             RutaGeneraFile = RutaDocumentos & "\Comprobantes Generados\" & TFA.ClaveAcceso & ".xml"
             
             If GeneraXML Then
                Clipboard.Clear
                Clipboard.SetText TextoXML
                Grabar_Consulta_Archivo TFA.TC & "_" & TFA.Serie & "-" & Format$(TFA.Factura, String(9, "0")), TextoXML
             End If
             
             Set DocumentoXML = New DOMDocument30
             DocumentoXML.loadXML (TextoXML)
             DocumentoXML.save (RutaGeneraFile)
             
             If TFA.TC <> "DO" Then
                TFA.Estado_SRI = "CG"
               '---------------------------------------------------------------------
               'Solo cuando es un XML externo
''                TFA.ClaveAcceso = "1410202201070439419600110010020000000030704394114"
''                MsgBox TFA.ClaveAcceso
               '---------------------------------------------------------------------
                SRI_Autorizacion = SRI_Generar_XML(TFA.ClaveAcceso, TFA.Estado_SRI)
                 
                SRI_Actualizar_Autorizacion_Factura TFA, SRI_Autorizacion
               'MsgBox SRI_Autorizacion.Estado_SRI
                 
                TFA.Estado_SRI = SRI_Autorizacion.Estado_SRI
                TFA.Error_SRI = SRI_Autorizacion.Error_SRI
                If SRI_Autorizacion.Estado_SRI = "OK" Then
                   Control_Procesos "F", TFA.TC & "-" & TFA.Serie & " No. " & Format(TFA.Factura, "000000000") & " Autorizada"
                   TFA.Autorizacion = SRI_Autorizacion.Autorizacion
                   If Autorizar Then SRI_Enviar_Mails TFA, SRI_Autorizacion, "FA"
                   If VerFactura Then SRI_Generar_PDF_FA TFA, VerFactura
                Else
                   If VerFactura Then
                      MsgBox "No se pudo realizar la conexion, Liste la Factura (" & TFA.Estado_SRI & ")"
                   Else
                      TextoImprimio = TextoImprimio _
                                    & TFA.TC & " No. " & TFA.Serie & vbTab & TFA.Factura & vbTab _
                                    & TFA.Cliente & vbTab & TFA.Error_SRI & vbCrLf
                   End If
                End If
             End If
          Else
             RatonNormal
             MsgBox MensajeNoAutorizarCE
          End If
       End If
       AdoDBDet.Close
       AdoDBFA.Close
    End If
    RatonNormal
End Sub

Public Sub SRI_Crear_Clave_Acceso_Guia_Remision(URLinet As Inet, _
                                                TFA As Tipo_Facturas, _
                                                VerGuiaRemision As Boolean, _
                                                Optional GeneraXML As Boolean, _
                                                Optional Autorizar As Boolean)
Dim AdoDBFA As ADODB.Recordset
Dim AdoDBDet As ADODB.Recordset
Dim AdoDBGR As ADODB.Recordset
Dim NumFile As Integer
Dim RutaGeneraFile As String

Dim TipoIdent As String
Dim TotalDescuento As Currency
Dim TotalDescuento_0 As Currency
Dim TotalDescuento_X As Currency
Dim Autorizar_XML As Boolean

    RatonReloj
    Autorizar_XML = True
    If Len(Fecha_Igualar) = 10 Then
       If CFechaLong(TFA.Fecha) < CFechaLong(Fecha_Igualar) Then Autorizar_XML = False
    End If
    TextoXML = ""
    
    If Autorizar_XML Then
   'Averiguamos si la Factura esta a nombre del Representante
    TBeneficiario = Leer_Datos_Clientes(TFA.CodigoC)
   'MsgBox TBeneficiario.RUC_CI_Rep & vbCrLf & TBeneficiario.Representante & vbCrLf & TBeneficiario.TD_Rep
    
    TFA.Cliente = TBeneficiario.Representante
    TFA.TD = TBeneficiario.TD_Rep
    TFA.CI_RUC = TBeneficiario.RUC_CI_Rep
    TFA.TelefonoC = TBeneficiario.Telefono1
    TFA.DireccionC = TBeneficiario.Direccion_Rep
    TFA.Curso = TBeneficiario.Direccion
    TFA.Grupo = TBeneficiario.Grupo_No
    TFA.EmailC = TBeneficiario.Email1
    TFA.EmailR = TBeneficiario.Email2
     
   'Detalle de descuentos
    sSQL = "SELECT DF.*,CP.Reg_Sanitario,CP.Marca " _
         & "FROM Detalle_Factura As DF, Catalogo_Productos As CP " _
         & "WHERE DF.Item = '" & NumEmpresa & "' " _
         & "AND DF.Periodo = '" & Periodo_Contable & "' " _
         & "AND DF.TC = '" & TFA.TC & "' " _
         & "AND DF.Serie = '" & TFA.Serie & "' " _
         & "AND DF.Autorizacion = '" & TFA.Autorizacion & "' " _
         & "AND DF.Factura = " & TFA.Factura & " " _
         & "AND LEN(DF.Autorizacion) >= 13 " _
         & "AND DF.T <> 'A' " _
         & "AND DF.Item = CP.Item " _
         & "AND DF.Periodo = CP.Periodo " _
         & "AND DF.Codigo = CP.Codigo_Inv " _
         & "ORDER BY DF.ID,DF.Codigo "
    Select_AdoDB AdoDBDet, sSQL
    RatonReloj
      
   'Encabezado de la Guia de Remision
    sSQL = "SELECT F.*,GR.Remision,GR.Comercial,GR.CIRUC_Comercial,GR.Entrega,GR.CIRUC_Entrega,GR.CiudadGRI,GR.CiudadGRF," _
         & "GR.Placa_Vehiculo,GR.FechaGRE,GR.FechaGRI,GR.FechaGRF,GR.Pedido,GR.Zona,GR.Serie_GR,GR.Autorizacion_GR," _
         & "GR.Clave_Acceso_GR,GR.Hora_Aut_GR,GR.Estado_SRI_GR,GR.Error_FA_SRI,GR.Fecha_Aut_GR " _
         & "FROM Facturas As F, Facturas_Auxiliares As GR " _
         & "WHERE F.Item = '" & NumEmpresa & "' " _
         & "AND F.Periodo = '" & Periodo_Contable & "' " _
         & "AND F.TC = '" & TFA.TC & "' " _
         & "AND F.Serie = '" & TFA.Serie & "' " _
         & "AND F.Autorizacion = '" & TFA.Autorizacion & "' " _
         & "AND F.Factura = " & TFA.Factura & " " _
         & "AND LEN(GR.Autorizacion_GR) = 13 " _
         & "AND GR.Remision > 0 " _
         & "AND F.T <> 'A' " _
         & "AND F.Item = GR.Item " _
         & "AND F.Periodo = GR.Periodo " _
         & "AND F.TC = GR.TC " _
         & "AND F.Serie = GR.Serie " _
         & "AND F.Autorizacion = GR.Autorizacion " _
         & "AND F.Factura = GR.Factura "
    Select_AdoDB AdoDBFA, sSQL
    RatonReloj
    With AdoDBFA
     If .RecordCount > 0 Then
         Autorizar_XML = True
         TFA.T = .fields("T")
         TFA.SP = .fields("SP")
         TFA.Porc_IVA = .fields("Porc_IVA")
         TFA.Imp_Mes = .fields("Imp_Mes")
         TFA.Fecha = .fields("Fecha")
         TFA.Vencimiento = .fields("Vencimiento")
         TFA.SubTotal = .fields("SubTotal")
         TFA.Sin_IVA = .fields("Sin_IVA")
         TFA.Con_IVA = .fields("Con_IVA")
         TFA.Descuento = .fields("Descuento")
         TFA.Descuento2 = .fields("Descuento2")
         TFA.Total_IVA = .fields("IVA")
         TFA.Total_MN = .fields("Total_MN")
         TFA.Razon_Social = .fields("Razon_Social")
         TFA.RUC_CI = .fields("RUC_CI")
         TFA.TB = .fields("TB")
         
        'MsgBox "Validar Porc IVA"
         
         Validar_Porc_IVA TFA.Fecha
        'Generamos la Clave de acceso
        '& Format$(TFA.Fecha, "ddmmyyyy") &
         If Len(TFA.Autorizacion_GR) >= 13 Then
            TFA.ClaveAcceso_GR = Format$(TFA.Fecha, "ddmmyyyy") & "06" & RUC & Ambiente & TFA.Serie_GR _
                               & Format$(TFA.Remision, String(9, "0")) _
                               & "123456781"
            TFA.ClaveAcceso_GR = TFA.ClaveAcceso_GR & Digito_Verificador_Modulo11(TFA.ClaveAcceso_GR)
         Else
            TFA.ClaveAcceso_GR = Ninguno
         End If
         TFA.Hora_GR = Format$(Time, FormatoTimes)
         SRI_Autorizacion.Clave_De_Acceso = TFA.ClaveAcceso_GR
         sSQL = "UPDATE Facturas_Auxiliares " _
              & "SET Clave_Acceso_GR = '" & TFA.ClaveAcceso_GR & "' " _
              & "WHERE Item = '" & NumEmpresa & "' " _
              & "AND Periodo = '" & Periodo_Contable & "' " _
              & "AND TC = '" & TFA.TC & "' " _
              & "AND Serie = '" & TFA.Serie & "' " _
              & "AND Factura = " & TFA.Factura & " " _
              & "AND CodigoC = '" & TFA.CodigoC & "' " _
              & "AND Autorizacion = '" & TFA.Autorizacion & "' "
         Ejecutar_SQL_SP sSQL
         
         
        'ENCABEZADO XML PARA EL SRI DE LA GUIA DE REMISION
        'standalone=""yes""
         Insertar_Campo_XML "<?xml version=""1.0"" encoding=""UTF-8""?>"
         Insertar_Campo_XML "<guiaRemision id=""comprobante"" version=""1.1.0"">"
         
         Insertar_Campo_XML AbrirXML("infoTributaria")
            Insertar_Campo_XML CampoXML("ambiente", Ambiente)
            Insertar_Campo_XML CampoXML("tipoEmision", "1")
            Insertar_Campo_XML CampoXML("razonSocial", RazonSocial)
            Insertar_Campo_XML CampoXML("nombreComercial", NombreComercial)
            Insertar_Campo_XML CampoXML("ruc", RUC)
            Insertar_Campo_XML CampoXML("claveAcceso", TFA.ClaveAcceso_GR)
            Insertar_Campo_XML CampoXML("codDoc", "06")
            Insertar_Campo_XML CampoXML("estab", MidStrg(TFA.Serie_GR, 1, 3))
            Insertar_Campo_XML CampoXML("ptoEmi", MidStrg(TFA.Serie_GR, 4, 3))
            Insertar_Campo_XML CampoXML("secuencial", Format$(TFA.Remision, String(9, "0")))
            Insertar_Campo_XML CampoXML("dirMatriz", ULCase(Direccion))
           'Tipo de Contribuyente
            If AgenteRetencion <> Ninguno Then Insertar_Campo_XML CampoXML("agenteRetencion", "1")
            If MicroEmpresa = "CONTRIBUYENTE RÉGIMEN RIMPE" Then Insertar_Campo_XML CampoXML("contribuyenteRimpe", "CONTRIBUYENTE RÉGIMEN RIMPE")
         Insertar_Campo_XML CerrarXML("infoTributaria")
         
         Insertar_Campo_XML AbrirXML("infoGuiaRemision")
            Insertar_Campo_XML CampoXML("dirEstablecimiento", ULCase(DireccionEstab))
            Insertar_Campo_XML CampoXML("dirPartida", .fields("CiudadGRI"))
            Insertar_Campo_XML CampoXML("razonSocialTransportista", .fields("Comercial"))
            DigVerif = Digito_Verificador(.fields("CIRUC_Comercial"))
            TipoIdent = "P"
            Select Case Tipo_RUC_CI.Tipo_Beneficiario
              Case "R": If .fields("CIRUC_Comercial") = String(13, "9") Then TipoIdent = "07" Else TipoIdent = "04"
              Case "C": TipoIdent = "05"
              Case "P": TipoIdent = "06"
              Case Else: TipoIdent = "07"
            End Select
            Insertar_Campo_XML CampoXML("tipoIdentificacionTransportista", TipoIdent)
            Insertar_Campo_XML CampoXML("rucTransportista", .fields("CIRUC_Comercial"))
            Insertar_Campo_XML CampoXML("rise", "000")
            Insertar_Campo_XML CampoXML("obligadoContabilidad", Obligado_Conta)
            If Len(ContEspec) > 1 Then Insertar_Campo_XML CampoXML("contribuyenteEspecial", ContEspec)
            Insertar_Campo_XML CampoXML("fechaIniTransporte", .fields("FechaGRI"))
            Insertar_Campo_XML CampoXML("fechaFinTransporte", .fields("FechaGRF"))
            Insertar_Campo_XML CampoXML("placa", .fields("Placa_Vehiculo"))
         Insertar_Campo_XML CerrarXML("infoGuiaRemision")
         
         Insertar_Campo_XML AbrirXML("destinatarios")
            Insertar_Campo_XML AbrirXML("destinatario")
            Insertar_Campo_XML CampoXML("identificacionDestinatario", .fields("CIRUC_Entrega"))
            Insertar_Campo_XML CampoXML("razonSocialDestinatario", .fields("Entrega"))
            Insertar_Campo_XML CampoXML("dirDestinatario", .fields("CiudadGRF"))
            Insertar_Campo_XML CampoXML("motivoTraslado", "Translado de mercaderia")
''            Insertar_Campo_XML CampoXML("docAduaneroUnico", "")
''            Insertar_Campo_XML CampoXML("codEstabDestino", "001")
            Insertar_Campo_XML CampoXML("ruta", "De " & .fields("CiudadGRI") & " a " & .fields("CiudadGRF"))
            Select Case TFA.TC
              Case "FA": Insertar_Campo_XML CampoXML("codDocSustento", "01")
              Case "NV": Insertar_Campo_XML CampoXML("codDocSustento", "02")
              Case Else: Insertar_Campo_XML CampoXML("codDocSustento", "00")
            End Select
            Cadena = MidStrg(TFA.Serie, 1, 3) & "-" & MidStrg(TFA.Serie, 4, 3) & "-" & Format(TFA.Factura, "000000000")
            Insertar_Campo_XML CampoXML("numDocSustento", Cadena)
            Insertar_Campo_XML CampoXML("numAutDocSustento", TFA.Autorizacion)
            Insertar_Campo_XML CampoXML("fechaEmisionDocSustento", TFA.Fecha)
            
           'Detalle de la Factura/Nota de Venta
            RatonReloj
            With AdoDBDet
             If .RecordCount > 0 Then
                 Insertar_Campo_XML AbrirXML("detalles")
                 Do While Not .EOF
                    Producto = TrimStrg(.fields("Producto"))
                    If TFA.Imp_Mes Then
                       If Len(.fields("Ticket")) > 1 Then Producto = Producto & ", " & .fields("Ticket")
                       If Len(.fields("Mes")) > 1 Then Producto = Producto & ": " & .fields("Mes")
                    End If
                    If TFA.SP Then
                       Producto = Producto _
                                & ", Lote No. " & .fields("Lote_No") _
                                & ", ELAB. " & .fields("Fecha_Fab") _
                                & ", VENC. " & .fields("Fecha_Exp") _
                                & ", Reg. Sanit. " & .fields("Reg_Sanitario") _
                                & ", Modelo: " & .fields("Modelo") _
                                & ", Serie No. " & .fields("Serie_No") _
                                & ", Procedencia: " & .fields("Procedencia")
                    End If
                    SubTotal = (.fields("Cantidad") * .fields("Precio")) - (.fields("Total_Desc") + .fields("Total_Desc2"))
                    Insertar_Campo_XML AbrirXML("detalle")
                        Insertar_Campo_XML CampoXML("codigoInterno", .fields("Codigo"))
                        If Len(.fields("Codigo_Barra")) > 1 Then Insertar_Campo_XML CampoXML("codigoAdicional", .fields("Codigo_Barra"))
                        Insertar_Campo_XML CampoXML("descripcion", Producto)
                        Insertar_Campo_XML CampoXML("cantidad", Format$(.fields("Cantidad"), "#0.000000"))
                    Insertar_Campo_XML CerrarXML("detalle")
                   .MoveNext
                 Loop
                 Insertar_Campo_XML CerrarXML("detalles")
             End If
            End With
            Insertar_Campo_XML CerrarXML("destinatario")
         Insertar_Campo_XML CerrarXML("destinatarios")
         
         Insertar_Campo_XML AbrirXML("infoAdicional")
            If TFA.Cliente <> Ninguno And TFA.Razon_Social <> TFA.Cliente Then
               If Len(TFA.Cliente) > 1 Then Insertar_Campo_XML "<campoAdicional nombre=""Beneficiario"">" & TFA.Cliente & "</campoAdicional>"
               If Len(TFA.Curso) > 1 Then Insertar_Campo_XML "<campoAdicional nombre=""Ubicacion"">" & TFA.Grupo & "-" & TFA.Curso & "</campoAdicional>"
            End If
            If Len(TFA.DireccionC) > 1 Then Insertar_Campo_XML "<campoAdicional nombre=""Direccion"">" & TFA.DireccionC & "</campoAdicional>"
            If Len(TFA.TelefonoC) > 1 Then Insertar_Campo_XML "<campoAdicional nombre=""Telefono"">" & TFA.TelefonoC & "</campoAdicional>"
            If Len(TFA.EmailC) > 1 Then Insertar_Campo_XML "<campoAdicional nombre=""Email"">" & TFA.EmailC & "</campoAdicional>"
         Insertar_Campo_XML CerrarXML("infoAdicional")
        'Fin del Archivo Xml
         Insertar_Campo_XML CerrarXML("guiaRemision")
     End If
    End With
        
      'Grabamos el comprobante XML
       RatonReloj
       If Len(TFA.Autorizacion_GR) = 13 Then
          If CFechaLong(TFA.Fecha) <= CFechaLong(Fecha_CE) Then
            'MsgBox RutaDocumentos
             RutaGeneraFile = RutaDocumentos & "\Comprobantes Generados\" & TFA.ClaveAcceso_GR & ".xml"
             
             If GeneraXML Then
                Clipboard.Clear
                Clipboard.SetText TextoXML
                Grabar_Consulta_Archivo "GR_" & TFA.Serie_GR & "-" & Format$(TFA.Remision, String(9, "0")), TextoXML
             End If
                          
             Set DocumentoXML = New DOMDocument30
             DocumentoXML.loadXML (TextoXML)
             DocumentoXML.save (RutaGeneraFile)
             
             TFA.Estado_SRI_GR = "CG"
             SRI_Autorizacion = SRI_Generar_XML(TFA.ClaveAcceso_GR, SRI_Autorizacion.Estado_SRI)
            'MsgBox SRI_Autorizacion.Estado_SRI
             TFA.Estado_SRI_GR = SRI_Autorizacion.Estado_SRI
             TFA.Error_SRI = SRI_Autorizacion.Error_SRI
             TFA.Fecha_Aut_GR = SRI_Autorizacion.Fecha_Autorizacion
             TFA.Hora_GR = SRI_Autorizacion.Hora_Autorizacion
             If SRI_Autorizacion.Estado_SRI = "OK" Then
                Control_Procesos "F", "GR-" & TFA.Serie_GR & " No. " & Format(TFA.Remision, "000000000") & " Autorizada"
                TFA.Autorizacion_GR = SRI_Autorizacion.Autorizacion
                SRI_Actualizar_Autorizacion_Guia_Remision TFA, SRI_Autorizacion
                If Autorizar Then SRI_Enviar_Mails TFA, SRI_Autorizacion, "GR"
                If VerGuiaRemision Then SRI_Generar_PDF_GR TFA, VerGuiaRemision
             Else
                If VerGuiaRemision Then
                   MsgBox "No se pudo realizar la conexion, Liste la Factura (" & TFA.Estado_SRI_GR & ")" & vbCrLf _
                         & TFA.Error_SRI
                Else
                   TextoImprimio = TextoImprimio _
                                 & TFA.TC & " No. " & TFA.Serie_GR & vbTab & TFA.Factura & vbTab _
                                 & TFA.Cliente & vbTab & TFA.Error_SRI & vbCrLf _
                                 & TFA.Error_SRI
                End If
             End If
          Else
             RatonNormal
             MsgBox MensajeNoAutorizarCE
          End If
       End If
       AdoDBDet.Close
       AdoDBFA.Close
    End If
    RatonNormal
End Sub

Public Sub SRI_Crear_Clave_Acceso_Nota_Credito(TFA As Tipo_Facturas, _
                                               VerNotaCredito As Boolean, _
                                               Optional GeneraXML As Boolean)
Dim AdoDBNC As ADODB.Recordset
Dim NumTrans As Long
Dim NumFile As Integer
Dim TipoIdent As String
Dim CodAdicional As String
Dim RutaGeneraFile As String
Dim Autorizar_XML As Boolean
Dim SubT_Con_Inv As Boolean
Dim Con_Inv As Boolean
    
    RatonReloj
     
    Con_Inv = False
    Autorizar_XML = True
    If Len(Fecha_Igualar) = 10 Then
       If CFechaLong(TFA.Fecha_NC) < CFechaLong(Fecha_Igualar) Then Autorizar_XML = False
    End If

   'Autorizamos la Nota de Credito
    If Autorizar_XML Then
    
    Leer_Datos_FA_NV TFA
   'Averiguamos si la Factura esta a nombre del Representante
'''    TBeneficiario = Leer_Datos_Clientes(TFA.CodigoC)
'''    With TBeneficiario
'''         TFA.Cliente = .Cliente
'''         TFA.TD = .TD
'''         TFA.CI_RUC = .CI_RUC
'''         TFA.TelefonoC = .Telefono1
'''         If Len(.Representante) > 1 And Len(.RUC_CI_Rep) > 1 Then
'''            TFA.Razon_Social = .Representante
'''            TFA.RUC_CI = .RUC_CI_Rep
'''            TFA.TB = .TD_Rep
'''         Else
'''            Select Case TFA.TD
'''              Case "C", "R", "P"
'''                   TFA.Razon_Social = TFA.Cliente
'''                   TFA.RUC_CI = TFA.CI_RUC
'''                   TFA.TB = TFA.TD
'''                   TFA.TelefonoC = .TelefonoT
'''              Case Else
'''                   TFA.Razon_Social = "CONSUMIDOR FINAL"
'''                   TFA.RUC_CI = "9999999999999"
'''                   TFA.TB = "R"
'''            End Select
'''         End If
'''         TFA.DireccionC = .Direccion
'''         TFA.EmailC = .Email1
'''         TFA.EmailR = .Email2
'''    End With
    
   'NOTA DE CREDITO
    SubT_Con_Inv = False
    Total_Sin_IVA = 0
    Total_Con_IVA = 0
    Total_Desc = 0
    Total_Desc2 = 0
    TFA.Total_IVA_NC = 0
    
    sSQL = "SELECT Autorizacion, Codigo_Inv, Producto, Cantidad, Precio, Total, Total_IVA, Descuento, Cta_Devolucion, CodBodega, Porc_IVA, Mes, Mes_No , Anio, ID " _
         & "FROM Detalle_Nota_Credito " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND Serie = '" & TFA.Serie_NC & "' " _
         & "AND Secuencial = " & TFA.Nota_Credito & " " _
         & "ORDER BY ID "
    Select_AdoDB AdoDBNC, sSQL, "Generacion NC"
    With AdoDBNC
     If .RecordCount > 0 Then
         Con_Inv = True
         Do While Not .EOF
            Total_Desc = Total_Desc + .fields("Descuento")
            If .fields("Total_IVA") = 0 Then
                Total_Sin_IVA = Total_Sin_IVA + .fields("Total")
            Else
                Total_Con_IVA = Total_Con_IVA + .fields("Total")
            End If
            TFA.Total_IVA_NC = TFA.Total_IVA_NC + .fields("Total_IVA")
           .MoveNext
         Loop
     End If
    End With
    If AdoDBNC.RecordCount <= 0 Then
       AdoDBNC.Close
       sSQL = "SELECT * " _
            & "FROM Trans_Abonos " _
            & "WHERE Item = '" & NumEmpresa & "' " _
            & "AND Periodo = '" & Periodo_Contable & "' " _
            & "AND Autorizacion = '" & TFA.Autorizacion & "' " _
            & "AND Serie = '" & TFA.Serie & "' " _
            & "AND TP = '" & TFA.TC & "' " _
            & "AND Factura = " & TFA.Factura & " " _
            & "AND Serie_NC = '" & TFA.Serie_NC & "' " _
            & "AND Secuencial_NC = " & TFA.Nota_Credito & " " _
            & "AND Banco = 'NOTA DE CREDITO' " _
            & "ORDER BY TP,Fecha,Cta,Cta_CxP,Abono,Banco,Cheque "
       Select_AdoDB AdoDBNC, sSQL
        With AdoDBNC
          'MsgBox .RecordCount
         If .RecordCount > 0 Then
             Do While Not .EOF
                If .fields("Cheque") = "I.V.A." Then
                    TFA.Total_IVA_NC = .fields("Abono")
                    SubT_Con_Inv = True
                ElseIf .fields("Cheque") = "VENTAS SIN IVA" Then
                    Total_Sin_IVA = .fields("Abono")
                Else
                    Total_Con_IVA = .fields("Abono")
                End If
               .MoveNext
             Loop
''             If SubT_Con_Inv Then
''                Total_Con_IVA = Total_Sin_IVA
''                Total_Sin_IVA = 0
''             End If
         End If
        End With
    End If
    Total_Sin_IVA = Redondear(Total_Sin_IVA, 2)
    Total_Con_IVA = Redondear(Total_Con_IVA, 2)
    TFA.Total_IVA_NC = Redondear(TFA.Total_IVA_NC, 2)
    TFA.SubTotal_NC = Redondear(Total_Sin_IVA + Total_Con_IVA, 2)
    TextoXML = ""
   'Generacion de la Nota Credito si es Electronica
    With TFA
     'MsgBox .Total_IVA_NC
      If (.SubTotal_NC + .Total_IVA_NC) > 0 Then
         If TFA.Porc_NC = 0 Then
            Validar_Porc_IVA TFA.Fecha_NC
         Else
            Porc_IVA = TFA.Porc_NC
         End If
        .Hora_NC = Format$(Time, FormatoTimes)
        'Algoritmo Modulo 11 para la clave de la retencion
        '& Format$(.Fecha_NC, "ddmmyyyy")
        .ClaveAcceso_NC = Format$(.Fecha_NC, "ddmmyyyy") & "04" & RUC & Ambiente _
                        & .Serie_NC & Format$(.Nota_Credito, String(9, "0")) _
                        & "123456781"
        .ClaveAcceso_NC = Replace(.ClaveAcceso_NC, ".", "1")
        .ClaveAcceso_NC = .ClaveAcceso_NC & Digito_Verificador_Modulo11(.ClaveAcceso_NC)
         
        'MsgBox .ClaveAcceso_NC
         
        'ENCABEZADO XML PARA EL SRI DE LA NOTA DE CREDITO
         Insertar_Campo_XML "<?xml version=""1.0"" encoding=""UTF-8""?>"
        
        ' Insertar_Campo_XML "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>"
         Insertar_Campo_XML "<notaCredito id=""comprobante"" version=""1.1.0"">"
         Insertar_Campo_XML AbrirXML("infoTributaria")
            Insertar_Campo_XML CampoXML("ambiente", Ambiente)
            Insertar_Campo_XML CampoXML("tipoEmision", "1")
            Insertar_Campo_XML CampoXML("razonSocial", RazonSocial)
            Insertar_Campo_XML CampoXML("nombreComercial", NombreComercial)
            Insertar_Campo_XML CampoXML("ruc", RUC)
            Insertar_Campo_XML CampoXML("claveAcceso", .ClaveAcceso_NC)
            Insertar_Campo_XML CampoXML("codDoc", "04")
            Insertar_Campo_XML CampoXML("estab", MidStrg(.Serie_NC, 1, 3))
            Insertar_Campo_XML CampoXML("ptoEmi", MidStrg(.Serie_NC, 4, 3))
            Insertar_Campo_XML CampoXML("secuencial", Format$(.Nota_Credito, String(9, "0")))
            Insertar_Campo_XML CampoXML("dirMatriz", ULCase(Direccion))
           'Tipo de Contribuyente
            If AgenteRetencion <> Ninguno Then Insertar_Campo_XML CampoXML("agenteRetencion", "1")
            If MicroEmpresa = "CONTRIBUYENTE RÉGIMEN RIMPE" Then Insertar_Campo_XML CampoXML("contribuyenteRimpe", "CONTRIBUYENTE RÉGIMEN RIMPE")
         Insertar_Campo_XML CerrarXML("infoTributaria")
         
         Insertar_Campo_XML AbrirXML("infoNotaCredito")
            Insertar_Campo_XML CampoXML("fechaEmision", .Fecha_NC)
            Insertar_Campo_XML CampoXML("dirEstablecimiento", ULCase(DireccionEstab))
            Select Case .TB
              Case "R": If .RUC_CI = "9999999999999" Then TipoIdent = "07" Else TipoIdent = "04"
              Case "C": TipoIdent = "05"
              Case "P": TipoIdent = "06"
            End Select
            Insertar_Campo_XML CampoXML("tipoIdentificacionComprador", TipoIdent)
            Insertar_Campo_XML CampoXML("razonSocialComprador", .Razon_Social)
            Insertar_Campo_XML CampoXML("identificacionComprador", .RUC_CI)
            If Len(ContEspec) > 1 Then Insertar_Campo_XML CampoXML("contribuyenteEspecial", ContEspec)
            Insertar_Campo_XML CampoXML("obligadoContabilidad", Obligado_Conta)
            Insertar_Campo_XML CampoXML("codDocModificado", "01")
            Insertar_Campo_XML CampoXML("numDocModificado", MidStrg(.Serie, 1, 3) & "-" & MidStrg(.Serie, 4, 3) & "-" & Format$(.Factura, String(9, "0")))
            Insertar_Campo_XML CampoXML("fechaEmisionDocSustento", .Fecha)
            Insertar_Campo_XML CampoXML("totalSinImpuestos", Format$(Total_Sin_IVA + Total_Con_IVA - Total_Desc, "#0.00"))
            Insertar_Campo_XML CampoXML("valorModificacion", Format$(Total_Sin_IVA + Total_Con_IVA - Total_Desc + .Total_IVA_NC, "#0.00"))
            Insertar_Campo_XML CampoXML("moneda", "DOLAR")
            Insertar_Campo_XML AbrirXML("totalConImpuestos")
                Insertar_Campo_XML AbrirXML("totalImpuesto")
                   Insertar_Campo_XML CampoXML("codigo", "2")
                   If (Porc_IVA * 100) > 12 Then
                      Insertar_Campo_XML CampoXML("codigoPorcentaje", "3")
                   Else
                      Insertar_Campo_XML CampoXML("codigoPorcentaje", "2")
                   End If
                   Insertar_Campo_XML CampoXML("baseImponible", Total_Con_IVA - Total_Desc)
                   Insertar_Campo_XML CampoXML("valor", Format$(.Total_IVA_NC, "#0.00"))
                Insertar_Campo_XML CerrarXML("totalImpuesto")
            Insertar_Campo_XML CerrarXML("totalConImpuestos")
            Insertar_Campo_XML CampoXML("motivo", "Anulacion por Nota de Credito")
         Insertar_Campo_XML CerrarXML("infoNotaCredito")
         
        'Detalle de la Nota de Credito
         RatonReloj
         With AdoDBNC
          If .RecordCount > 0 Then
             .MoveFirst
              Insertar_Campo_XML AbrirXML("detalles")
                Do While Not .EOF
                   CodAdicional = CambioCodigoCtaSup(.fields("Codigo_Inv"))
                   'MsgBox PVP_NC & vbCrLf & .Fields("Cantidad_NC")
                   Insertar_Campo_XML AbrirXML("detalle")
                       Insertar_Campo_XML CampoXML("codigoInterno", .fields("Codigo_Inv"))
                       Insertar_Campo_XML CampoXML("codigoAdicional", CodAdicional)
                       Insertar_Campo_XML CampoXML("descripcion", .fields("Producto"))
                       Insertar_Campo_XML CampoXML("cantidad", .fields("Cantidad"))
                       Insertar_Campo_XML CampoXML("precioUnitario", Format$(.fields("Precio"), "#0.0000"))
                       Insertar_Campo_XML CampoXML("descuento", Format$(.fields("Descuento"), "#0.00"))
                       Insertar_Campo_XML CampoXML("precioTotalSinImpuesto", Format$(.fields("Total") - .fields("Descuento"), "#0.00"))
                      'MsgBox .Fields("Codigo_Inv") & vbCrLf & .Fields("Total_IVA")
                       Insertar_Campo_XML AbrirXML("impuestos")
                           Insertar_Campo_XML AbrirXML("impuesto")
                              Insertar_Campo_XML CampoXML("codigo", "2")
                              If .fields("Total_IVA") = 0 Then
                                  Insertar_Campo_XML CampoXML("codigoPorcentaje", "0")
                                  Insertar_Campo_XML CampoXML("tarifa", "0")
                              Else
                                  If (Porc_IVA * 100) > 12 Then
                                     Insertar_Campo_XML CampoXML("codigoPorcentaje", "3")
                                  Else
                                     Insertar_Campo_XML CampoXML("codigoPorcentaje", "2")
                                  End If
                                  Insertar_Campo_XML CampoXML("tarifa", Porc_IVA * 100)
                              End If
                              Insertar_Campo_XML CampoXML("baseImponible", Format$(.fields("Total") - .fields("Descuento"), "#0.00"))
                              Insertar_Campo_XML CampoXML("valor", Format$(.fields("Total_IVA"), "#0.00"))
                          Insertar_Campo_XML CerrarXML("impuesto")
                       Insertar_Campo_XML CerrarXML("impuestos")
                   Insertar_Campo_XML CerrarXML("detalle")
                  .MoveNext
                Loop
              Insertar_Campo_XML CerrarXML("detalles")
          End If
         End With
         AdoDBNC.Close
         
        'FIN DE XML DE NOTA CREDITO
         Insertar_Campo_XML AbrirXML("infoAdicional")
            If Len(TFA.DireccionC) > 1 Then Insertar_Campo_XML "<campoAdicional nombre=""Direccion"">" & TFA.DireccionC & "</campoAdicional>"
            If Len(TFA.TelefonoC) > 1 Then Insertar_Campo_XML "<campoAdicional nombre=""Telefono"">" & TFA.TelefonoC & "</campoAdicional>"
            If Len(TFA.EmailC) > 1 Then Insertar_Campo_XML "<campoAdicional nombre=""Email"">" & TFA.EmailC & "</campoAdicional>"
            Insertar_Campo_XML "<campoAdicional nombre=""Comprobante No"">" & TFA.TC & "-" & Format$(TFA.Factura, "0000000000") & "</campoAdicional>"
         Insertar_Campo_XML CerrarXML("infoAdicional")
         Insertar_Campo_XML CerrarXML("notaCredito")
         
        'Grabamos el comprobante XML de la Retencion
         If Len(TFA.Autorizacion_NC) >= 13 Then
            If CFechaLong(TFA.Fecha) <= CFechaLong(Fecha_CE) Then
               'MsgBox TFA.ClaveAcceso_NC
                RutaGeneraFile = RutaDocumentos & "\Comprobantes Generados\" & TFA.ClaveAcceso_NC & ".xml"
                
                If GeneraXML Then
                   Clipboard.Clear
                   Clipboard.SetText TextoXML
                   Grabar_Consulta_Archivo "NC_" & TFA.Serie_NC & "-" & Format$(TFA.Nota_Credito, String(9, "0")), TextoXML
                End If
                 
                Set DocumentoXML = New DOMDocument30
                DocumentoXML.loadXML (TextoXML)
                DocumentoXML.save (RutaGeneraFile)

                    TFA.Estado_SRI_NC = "CG"
                   'Enviamos al SRI para que me autorice la Retencion
                    SRI_Autorizacion = SRI_Generar_XML(TFA.ClaveAcceso_NC, TFA.Estado_SRI_NC)
                    SRI_Actualizar_XML_Nota_Credito SRI_Autorizacion, TFA
                    RatonReloj
                    If SRI_Autorizacion.Estado_SRI = "OK" Then
                       Control_Procesos "F", "NC-" & TFA.Serie_NC & " No. " & Format(TFA.Nota_Credito, "000000000") & " Autorizada"
                       TFA.Estado_SRI_NC = SRI_Autorizacion.Estado_SRI
                       SRI_Actualizar_Autorizacion_Nota_Credito TFA, SRI_Autorizacion
                       SRI_Actualizar_Documento_XML TFA.ClaveAcceso_NC
                       SRI_Enviar_Mails TFA, SRI_Autorizacion, "NC"
                       SRI_Generar_PDF_NC TFA, VerNotaCredito
                       RatonNormal
                    Else
                       RatonNormal
                       TFA.Estado_SRI_NC = SRI_Autorizacion.Estado_SRI
                       If VerNotaCredito Then MsgBox "(" & SRI_Autorizacion.Estado_SRI & ")"
                    End If
                
            Else
                RatonNormal
                MsgBox MensajeNoAutorizarCE
            End If
         End If
     End If
    End With
    End If
    RatonNormal
End Sub

Public Sub SRI_Generar_XML_Firmado(ClaveDeAcceso As String)
Dim AdoDBXMLFirmado As ADODB.Recordset
Dim RutaXMLFirmado As String
    sSQL = "SELECT Documento_Autorizado " _
         & "FROM Trans_Documentos " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND Clave_Acceso = '" & ClaveDeAcceso & "' "
    Select_AdoDB AdoDBXMLFirmado, sSQL
   'MsgBox "Documento Firmado: " & AdoDBXMLFirmado.RecordCount
    If AdoDBXMLFirmado.RecordCount > 0 Then
       RutaXMLFirmado = RutaSysBases & "\TEMP\" & ClaveDeAcceso & ".xml"
       Escribir_Archivo RutaXMLFirmado, AdoDBXMLFirmado.fields("Documento_Autorizado")
      'MsgBox RutaXMLFirmado
    End If
    AdoDBXMLFirmado.Close
End Sub

Public Sub Recibo_Enviar_Mails(TFA As Tipo_Facturas)
Dim Comprobante As String

    If Len(TFA.Serie) = 6 And (TFA.Factura) > 0 Then
       Comprobante = "Recibo No " & TFA.Serie & "-" & Format$(TFA.Factura, "000000000")
       Generar_Recibo_PDF TFA, False
       TMail.TipoDeEnvio = "CO"
       TMail.ListaMail = 255

       TMail.Mensaje = "Cliente: " & TFA.Cliente & vbCrLf _
                     & "Codigo: " & TFA.CI_RUC & vbCrLf _
                     & "Emision: " & TFA.Fecha & vbCrLf _
                     & Comprobante & vbCrLf
       TMail.Asunto = TFA.Cliente & ", " & Comprobante
       TMail.Adjunto = RutaSysBases & "\TEMP\" & Comprobante & ".pdf"
       TMail.Credito_No = "R" & Format(TFA.Factura, "000000000")
      'Enviamos lista de mails
       TMail.para = ""
       Insertar_Mail TMail.para, TFA.EmailC
       Insertar_Mail TMail.para, TFA.EmailR
      'MsgBox RutaPDF & "; " & RutaXML
       If Email_CE_Copia Then
          Insertar_Mail TMail.para, EmailProcesos
          TMail.Credito_No = "X" & Format(TFA.Factura, "000000000")
       End If
       FEnviarCorreos.Show 1
    End If
 End Sub

