Attribute VB_Name = "SubRetenciones"

Global EsIVA As Boolean
Global IsDep As Boolean
Global IngAnual As Boolean
Global IngMensual As Boolean
Global PeriodoI As String
Global PeriodoF As String
Global TotalBaseCero As Single
Global Vanio As String
Global Vmes As String
Global VAnioR As String
Global NR As Integer
Global TD As String
Global tdm As String
Global tt As String
Global VIL As String
Global AP As String
Global BI As String
Global VR As String
Global Tot As Integer
Global ValRet As Single
Global IngLiqui As Single
Global AporIess As Single
Global BImpoCero As Double
Global BaseCero As Double
Global SumaC As Double
Global SumaC1 As Double
Global SumaC2 As Double
Global SumaV As Double
Global SumaV1 As Double
Global SumaV2 As Double
Global SumaTV As Double
Global SumaTV1 As Double
Global SumaI As Double
Global SumaI1 As Double
Global SumaE As Double
Global SumaE1 As Double
Global SumaN As Double
Global SumaN1 As Double
Global BaseImpon As Double
Global BaseImpor As Double
Global SumaSaldos As Double
Global SumaTotal As Double
Global BaseImpo As Double
Global IvaTotal As Double
Global PorcIva As Single
Global PorRet As Single
Global BaseServ As Double
Global BaseTrans As Double
Global BaseServt As Double
Global ValorIva As Double
Global IvaTrans As Double
Global PorRetSer As Single
Global IvaServ As Double
Global PorRetTran As Single
Global Apagar As Double
Global Topc As String
Global Periodo_No As String
Global NumAño As String
Global Tipo As String
Global Secuencia As String
Global No_Reg As Single
Global No_RegC As Single
Global No_RegV As Single
Global No_RegDC As Single
Global No_RegI As Single
Global No_RegE As Single
Global Compras As String
Global NotasDC As String
Global Import As String
Global Export As String
Global Vent As String
Global Tota As String
Global RutaGenera As String
Global CodPorcIva As String
Global CodPorRetSer As String
Global CodPorRetTrans As String
Global CodICE As String
Global CodPorc As String
Global NumeroRet As Single
Global NuevoCodigoC As String
Global NuevoTD As String
Global NuevaFecha As String
Global NuevoValor As Double
Global NuevoRet As Double
Global NuevoAporte As Double
Global ID As Integer
Global VLinea As Single
Global VColumna As Single
'Global RUC_Contador As String
'Global CI_Representante As String

Public Function CambioSiNo(DataSQL As Adodc) As Boolean
Dim Cancelar As Boolean
  With DataSQL.Recordset
     Mensajes = "Cambiar el Codigo de la transacción: " & vbCrLf
     For I = 0 To .Fields.Count - 1
        Mensajes = Mensajes & Space(20) & UCaseStrg(.Fields(I).Name) & ": " & .Fields(I) & "." & vbCrLf
     Next I
  End With
  Mensajes = Mensajes & vbCrLf _
           & "¿Realmente desea Cambiar esta Transacción." & Space(10)
  Titulo = "Confirmación de Cambio"
  If BoxMensaje = vbYes Then Cancelar = False Else Cancelar = True
  CambioSiNo = Cancelar
End Function
Public Function GrabarRetencion(AdoRetencion As Adodc, AdoDetRet As Adodc)
    SetAddNew AdoRetencion
    With AdoDetRet.Recordset
      For I = 0 To .Fields.Count - 1
        Cadena = .Fields(I).Name
        SetFields AdoRetencion, Cadena, .Fields(Cadena)
      Next I
    End With
    SetUpdate AdoRetencion
End Function

Public Function PrimerDiaAnio(FechaStr As String) As String
  FechaStr = Format(FechaStr, "dd/mm/yyyy")
  PrimerDiaAnio = "01/01/" & Format(Year(FechaStr))
End Function

Public Function CalculoPeriodo(FechaStr As String) As String
  FechaStr = Format(FechaStr, "dd/mm/yyyy")
  NumAño = FechaAnio(FechaStr)
  CalculoPeriodo = (((NumAño - 1995) * 12) + Val(MidStrg(FechaStr, 4, 2)))
End Function

Public Sub ImprimirTalon(AdoEmp As Adodc, _
                         AdoQry1 As Adodc, _
                         Vanio As String, _
                         NR As Integer, _
                         VIL As String, _
                         AP As String, _
                         BI As String, _
                         VR As String, _
                         Tot As Integer, _
                         SumaSaldos As Double, _
                         SumaTotal As Double)
On Error GoTo Errorhandler
Mensajes = "Seguro de Imprimir en:" & vbCrLf & Printer.DeviceName & "?"
Titulo = "IMPRESION"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
    InicioX = 0.5: InicioY = 0
    RatonReloj
    Escala_Centimetro 1, TipoTimes, 10, True
    'Iniciamos la consulta de impresion
    Printer.Line (0.7, 1)-(19.2, 27.9), Negro, B
    PrinterTexto 3.5, 1.2, "TALON RESUMEN DE RETENCIONES EN LA FUENTE DE IMPUESTO A LA RENTA"
    PrinterTexto 2, 1.8, "RUC:"
    PrinterTexto 2, 2.2, "RAZON SOCIAL:"
    PrinterTexto 2, 2.6, "AÑO"
    PrinterTexto 1.5, 3, "Certifico que la información contenida en el (los) disquete(s) adjunto(s) al presente, sobre las Retenciones en la Fuente"
    PrinterTexto 1.9, 3.4, "del Impuesto a la Renta realizadas durante el año indicado, es el fiel reflejo de lo registrado en este formulario"
    PrinterTexto 2, 3.9, "Archivo: RDEP" & Vanio & ".ANE"
    Printer.Line (0.7, 3.8)-(19.2, 4.3), Negro, B
    PrinterTexto 4.5, 4.4, "RELACION LABORAL - RENTAS EN RELACION DE DEPENDENCIA"
    PrinterTexto 6.2, 4.8, "Numero de Registros:"
    PrinterTexto 6.2, 5.2, "Valor Ingresos Líquidos:"
    PrinterTexto 6.2, 5.6, "Aporte personal IESS:"
    PrinterTexto 6.2, 6, "Base Imponible:"
    PrinterTexto 6.2, 6.4, "Valor Retenido:"
    PrinterTexto 2, 6.9, "Archivo: REOC" & Vanio & ".ANE "
    Printer.Line (0.7, 6.8)-(19.2, 7.3), Negro, B
    Printer.FontSize = 8.75: Printer.Font.Weight = False
    PrinterTexto 14.4, 7.4, "BASE IMPONIBLE  RETENCION"
    With AdoEmp.Recordset
    If .RecordCount > 0 Then
        .MoveFirst
        PrinterFields0 5, 1.8, .Fields("Ruc"), False
        PrinterFields0 5, 2.2, .Fields("Nombre_Comercial"), False
        If Vanio = "2001" Then Vanio = "2000-2001"
        PrinterTexto 5, 2.6, Vanio
    End If
    End With
    PrinterVariables0 11.5, 4.8, CInt(NR)
    PrinterVariables0 12, 5.2, Format(VIL, "#,##0.00")
    PrinterVariables0 12, 5.6, Format(AP, "#,##0.00")
    PrinterVariables0 12, 6, Format(BI, "#,##0.00")
    PrinterVariables0 12, 6.4, Format(VR, "#,##0.00")
    With AdoQry1.Recordset
        If .RecordCount > 0 Then
             .MoveFirst
            PosLinea = 7.8
            Printer.Line (14.5, 7.3)-(17.1, PosLinea), Negro, B
            Do While Not .EOF
                PrinterFields0 16.5, PosLinea, .Fields("Retencion"), False
                PrinterFields0 14.5, PosLinea, .Fields("Base_Imp"), False
                PrinterFields0 13.6, PosLinea, .Fields("Codigo"), False
                PrinterTexto 1.2, PosLinea, .Fields("Tipo_Ret"), True, 12
                Printer.Line (14.5, PosLinea)-(19.2, PosLinea)
                If Val(FTalon.Combo2.Text) > 2005 Then
                   PosLinea = PosLinea + 0.35
                   Printer.Line (14.5, (PosLinea - 0.35))-(14.5, PosLinea), Negro, B
                   Printer.Line (17.1, (PosLinea - 0.35))-(17.1, PosLinea), Negro, B
                Else
                   PosLinea = PosLinea + 0.5
                   Printer.Line (14.5, (PosLinea - 0.5))-(14.5, PosLinea), Negro, B
                   Printer.Line (17.1, (PosLinea - 0.5))-(17.1, PosLinea), Negro, B
                End If
                .MoveNext
            Loop
            RatonNormal
        End If
    End With
    Printer.Line (14.5, PosLinea)-(19.2, PosLinea)
    PosLinea = PosLinea + 0.1
    PrinterTexto 11.8, PosLinea, "Totales:"
    Printer.FontBold = True
    PrinterTexto 15, PosLinea, CStr(Format(SumaSaldos, "#,##0.00")), False, 2
    PrinterTexto 17.3, PosLinea, CStr(Format(SumaTotal, "#,##0.00")), False, 2
    Printer.FontBold = False
    PosLinea = PosLinea + 0.4
    PrinterTexto 11.8, PosLinea, "Numero de Registros:      " & CStr(Tot)
   ' PrinterTexto 15, PosLinea, CStr(Tot)
    PosLinea = PosLinea + 0.4
    Printer.Line (0.7, PosLinea)-(19.2, PosLinea), Negro, B
    Printer.Line (2, 27.1)-(7, 27.1), Negro, B
    PrinterTexto 2, 27.2, "FIRMA DEL CONTADOR "
    Printer.Line (11, 27.1)-(17, 27.1), Negro, B
    PrinterTexto 11, 27.2, "FIRMA DEL REPRESENTANTE LEGAL "
    PrinterTexto 2, 27.5, "SERVICIO DE RENTAS INTERNAS - FECHA:   " & FechaSistema
    Bandera = True
End If
Printer.EndDoc
Exit Sub
Errorhandler:
    RatonNormal
    ErrorDeImpresion
    Exit Sub
End Sub
Public Sub ImprimirTalonIVA(AdoEmp As Adodc, _
                         AdoCom As Adodc, _
                         AdoVen As Adodc, _
                         AdoNDC As Adodc, _
                         AdoImp As Adodc, _
                         AdoExp As Adodc, _
                         Vanio As String, _
                         Compras As String, _
                         ventas As String, _
                         NotasDC As String, _
                         Import As String, _
                         Export As String, _
                         Tota As String)
On Error GoTo Errorhandler
Mensajes = "Seguro de Imprimir en:" & vbCrLf & Printer.DeviceName & "?"
Titulo = "IMPRESION"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
InicioX = 0.5: InicioY = 0
RatonReloj
Escala_Centimetro 1, TipoTimes, 10, True
'Iniciamos la consulta de impresion
Printer.Line (0.5, 0.5)-(19.5, 28), Negro, B
Printer.FontSize = 12
PrinterTexto 3, 1.1, "TALON RESUMEN DE TRANSACCIONES LOCALES Y DEL EXTERIOR"
Printer.FontSize = 9
' Escala_Centimetro 1, TipoTimes, 9, True
PrinterTexto 2, 1.6, "Certifico que la información contenida en el (los) disquete(s) o CD(s), adjunto(s) al presente, de ''Anexos de"
PrinterTexto 2, 1.9, "IVA'' para el período " & Periodo_No & " es el fiel reflejo del siguiente reporte"
PrinterTexto 7, 2.3, "      Archivo:        ID" & Vanio & ".ANE"
PrinterVariables0 7, 2.6, "           RUC:    " & RUC
PrinterVariables0 7, 2.9, "Razón Social:    " & NombreComercial
PrinterVariables0 7, 3.2, "      Teléfono:   " & Telefono1
PrinterTexto 2, 3.7, "Archivo:   TL" & Vanio & ".ANE"
Printer.Line (0.6, 4.4)-(19.4, 14.8), Negro, B 'CUADRO TL
PrinterTexto 0.8, 4.5, "Cod                Transacción"
PrinterTexto 12.9, 4.5, "No.Reg."
PrinterTexto 14, 4.5, "Base Imponible tarifa"
PrinterTexto 14, 4.8, "   10% 12% y 14%"
PrinterTexto 17, 4.5, "Base Imponible"
PrinterTexto 17, 4.8, "  tarifa 0%"
Printer.Line (1.5, 4.4)-(13, 5.2), Negro, B 'LINEAS VERTICALES
Printer.Line (14, 4.4)-(17, 5.2), Negro, B
PrinterTexto 4.2, 5.5, "COMPRAS"
PrinterVariables0 13.5, 12, CStr(Compras)
Printer.Line (0.6, 5.2)-(19.4, 5.9), Negro, B
If AdoCom.Recordset.RecordCount > 0 Then
With AdoCom.Recordset
If .RecordCount > 0 Then
  .MoveFirst
   PosLinea = 6
   Do While Not .EOF
    Printer.Line (1.5, PosLinea - 0.1)-(1.5, PosLinea + 0.3), Negro, B
    Printer.Line (13, PosLinea - 0.1)-(13, PosLinea + 0.3), Negro, B
    Printer.Line (14, PosLinea - 0.1)-(14, PosLinea + 0.3), Negro, B
    Printer.Line (17, PosLinea - 0.1)-(17, PosLinea + 0.3), Negro, B
    PrinterFields0 17, PosLinea, .Fields("Base_Imponible_Tarifa_0"), False
    PrinterFields0 14.5, PosLinea, .Fields("Base_Imponible_Tarifa_10_12_y_14"), False
    PrinterFields0 12.5, PosLinea, .Fields("No_Registros"), False
    PrinterTexto 1.5, PosLinea, .Fields("Transaccion"), False
    PrinterFields0 0.8, PosLinea, .Fields("Codigo"), False
    PosLinea = PosLinea + 0.4
    Printer.Line (0.6, PosLinea - 0.1)-(19.4, PosLinea - 0.1), Negro, B
    .MoveNext
   Loop
End If
End With
End If
PrinterTexto 4, PosLinea, "TOTAL"
Printer.Line (0.6, PosLinea + 0.4)-(19.4, PosLinea + 0.4), Negro, B
PrinterTexto 4.2, 12.7, "VENTAS"
Printer.Line (0.6, 13.1)-(19.4, 13.1), Negro, B
If AdoVen.Recordset.RecordCount > 0 Then
With AdoVen.Recordset
If .RecordCount > 0 Then
  .MoveFirst
   PosLinea = 13.2
   Do While Not .EOF
    Printer.Line (1.5, PosLinea - 0.1)-(1.5, PosLinea + 0.3), Negro, B
    Printer.Line (13, PosLinea - 0.1)-(13, PosLinea + 0.3), Negro, B
    Printer.Line (14, PosLinea - 0.1)-(14, PosLinea + 0.3), Negro, B
    Printer.Line (17, PosLinea - 0.1)-(17, PosLinea + 0.3), Negro, B
    PrinterFields0 17, PosLinea, .Fields("Base_Imponible_Tarifa_0"), False
    PrinterFields0 14.5, PosLinea, .Fields("Base_Imponible_Tarifa_10_12_y_14"), False
    PrinterFields0 12.5, PosLinea, .Fields("No_Registros"), False
    PrinterTexto 1.5, PosLinea, .Fields("Transaccion"), False
    PrinterFields0 0.8, PosLinea, .Fields("Codigo"), False
    PosLinea = PosLinea + 0.4
    Printer.Line (0.6, PosLinea - 0.1)-(19.4, PosLinea - 0.1), Negro, B
    .MoveNext
   Loop
End If
End With
End If
PrinterTexto 4, PosLinea, "TOTAL"
PrinterVariables0 13.5, PosLinea, CStr(Vent)
PrinterTexto 2, 15.3, "Archivo:   DC" & Vanio & ".ANE"
Printer.Line (0.6, 15.8)-(19.4, 19.4), Negro, B 'CUADRO DC
PrinterTexto 0.8, 16.1, "Cod                Transacción"
PrinterTexto 12.9, 16.1, "No.Reg."
Printer.Line (1.5, 15.8)-(13, 16.6), Negro, B 'LINEAS VERTICALES
Printer.Line (14, 15.8)-(17, 16.6), Negro, B
Printer.Line (0.6, 16.6)-(19.4, 17.3), Negro, B
PrinterTexto 3.2, 16.8, "DETALLE DE N/C y N/D"
If AdoNDC.Recordset.RecordCount > 0 Then
With AdoNDC.Recordset
If .RecordCount > 0 Then
  .MoveFirst
   PosLinea = 17.4
   Do While Not .EOF
    Printer.Line (1.5, PosLinea - 0.1)-(1.5, PosLinea + 0.3), Negro, B
    Printer.Line (13, PosLinea - 0.1)-(13, PosLinea + 0.3), Negro, B
    Printer.Line (14, PosLinea - 0.1)-(14, PosLinea + 0.3), Negro, B
    Printer.Line (17, PosLinea - 0.1)-(17, PosLinea + 0.3), Negro, B
    PrinterFields0 17, PosLinea, .Fields("Base_Imponible_Tarifa_0"), False
    PrinterFields0 14.5, PosLinea, .Fields("Base_Imponible_Tarifa_10_12_y_14"), False
    PrinterFields0 12.5, PosLinea, .Fields("No_Registros"), False
    PrinterTexto 1.5, PosLinea, .Fields("Transaccion"), False
    PrinterFields0 0.8, PosLinea, .Fields("Codigo"), False
    PosLinea = PosLinea + 0.4
    Printer.Line (0.6, PosLinea - 0.1)-(19.4, PosLinea - 0.1), Negro, B
    .MoveNext
   Loop
End If
End With
End If
PrinterTexto 4, PosLinea, "TOTAL"
PrinterVariables0 13.5, PosLinea, CStr(NotasDC)
PrinterTexto 2, 20, "Archivo:   TX" & Vanio & ".ANE"
Printer.Line (0.6, 20.4)-(19.4, 25.9), Negro, B 'CUADRO Tx
PrinterTexto 0.8, 20.8, "Cod                Transacción"
PrinterTexto 12.9, 20.8, "No.Reg."
PrinterTexto 14, 20.8, "   Valor CIF o FOB"
PrinterTexto 17, 20.8, "   VALOR IVA"
PrinterTexto 3.2, 21.4, "    IMPORTACIONES"
Printer.Line (1.5, 20.4)-(13, 21.2), Negro, B 'LINEAS VERTICALES
Printer.Line (14, 20.4)-(17, 21.2), Negro, B
Printer.Line (0.6, 21.2)-(19.4, 21.9), Negro, B
With AdoImp.Recordset
If .RecordCount > 0 Then
  .MoveFirst
   PosLinea = 22
   Do While Not .EOF
    Printer.Line (1.5, PosLinea - 0.1)-(1.5, PosLinea + 0.3), Negro, B
    Printer.Line (13, PosLinea - 0.1)-(13, PosLinea + 0.3), Negro, B
    Printer.Line (14, PosLinea - 0.1)-(14, PosLinea + 0.3), Negro, B
    Printer.Line (17, PosLinea - 0.1)-(17, PosLinea + 0.3), Negro, B
    PrinterFields0 17, PosLinea, .Fields("Valor_CIF_o_FOB"), False
    PrinterFields0 14.5, PosLinea, .Fields("Valor_IVA"), False
    PrinterFields0 12.5, PosLinea, .Fields("No_Registros"), False
    PrinterTexto 1.5, PosLinea, MidStrg(.Fields("Transaccion"), 1, 85), False
    PrinterFields0 0.8, PosLinea, .Fields("Codigo"), False
    PosLinea = PosLinea + 0.4
    Printer.Line (0.6, PosLinea - 0.1)-(19.4, PosLinea - 0.1), Negro, B
    .MoveNext
   Loop
End If
End With
PrinterTexto 4, PosLinea, "TOTAL"
PrinterVariables0 13.5, PosLinea, CStr(Import)
Printer.Line (0.6, PosLinea + 0.3)-(19.4, PosLinea + 0.3), Negro, B
PrinterTexto 3.2, 23.8, "    EXPORTACIONES"
Printer.Line (0.6, 23.5)-(19.4, 24.3), Negro, B
With AdoExp.Recordset
If .RecordCount > 0 Then
  .MoveFirst
   PosLinea = 24.4
   Do While Not .EOF
    Printer.Line (1.5, PosLinea - 0.1)-(1.5, PosLinea + 0.3), Negro, B
    Printer.Line (13, PosLinea - 0.1)-(13, PosLinea + 0.3), Negro, B
    Printer.Line (14, PosLinea - 0.1)-(14, PosLinea + 0.3), Negro, B
    Printer.Line (17, PosLinea - 0.1)-(17, PosLinea + 0.3), Negro, B
    PrinterFields0 17, PosLinea, .Fields("Valor_CIF_o_FOB"), False
    PrinterFields0 14.5, PosLinea, .Fields("Valor_IVA"), False
    PrinterFields0 12.5, PosLinea, .Fields("No_Registros"), False
    PrinterTexto 1.5, PosLinea, MidStrg(.Fields("Transaccion"), 1, 85), False
    PrinterFields0 0.8, PosLinea, .Fields("Codigo"), False
    PosLinea = PosLinea + 0.4
    Printer.Line (0.6, PosLinea - 0.1)-(19.4, PosLinea - 0.1), Negro, B
    .MoveNext
   Loop
End If
End With
PrinterTexto 4, PosLinea, "TOTAL"
PrinterVariables0 13.5, PosLinea, CStr(Export)
PrinterTexto 1.9, 26, "TOTAL REGISTROS VALIDADOS :"
PrinterVariables0 13.2, 26, CStr(Tota)
Printer.Line (3, 27.2)-(7, 27.2), Negro, B
PrinterTexto 3, 27.4, "Firma del Contador"
Printer.Line (13, 27.2)-(17, 27.2), Negro, B
PrinterTexto 13, 27.4, "Firma del Representante Legal"
Printer.EndDoc
Bandera = True
End If
Exit Sub
Errorhandler:
    RatonNormal
    ErrorDeImpresion
    Exit Sub
End Sub

Public Sub GenerarArchivoTalon(MiFrom As Form, _
                               DtaAux As Adodc, _
                               Carpeta As String, _
                               RutaGeneraFile As String, _
                               NombreFile As String, _
                               Tipo As Boolean)
Dim NumFile As Integer
Dim CaptionOld As String
Dim ValorBool As String
Dim CodCampS As String
RatonReloj
FAConLineas = False
CaptionOld = UCaseStrg(LeftStrg(MiFrom.Caption, 3))
RutaGeneraFile = RutaGeneraFile & Carpeta
Cadena = " "
If RutaGeneraFile <> "A:" Then Cadena = Dir(RutaGeneraFile, vbDirectory)
If Cadena = "" Then MkDir (RutaGeneraFile)
If RutaGeneraFile = "A:" Then
  RutaGeneraFile = RutaGeneraFile & NombreFile
Else
  RutaGeneraFile = RutaGeneraFile & "\" & NombreFile
End If
NumFile = FreeFile
Contador = 0
Open RutaGeneraFile For Output As #NumFile ' Abre el archivo.
With DtaAux.Recordset
     ReDim TipoC(.Fields.Count - 1) As Campos_Tabla
     For I = 0 To .Fields.Count - 1
         TipoC(I).Campo = .Fields(I).Name
         TipoC(I).Ancho = AnchoTipoCampoTexto(.Fields(I))
     Next I
     FAConLineas = False
     If .RecordCount > 0 Then
        .MoveFirst
         Do While Not .EOF
         MiFrom.Caption = RutaGeneraFile & ": Registro No. " & Contador
         Select Case Topc
          Case "SI"
            Print #NumFile, SetearCeros(.Fields("RUC"), 13, 0, True, FAConLineas);
            Print #NumFile, SetearCeros(.Fields("CI_RUC"), 10, 0, True, FAConLineas);
            Select Case .Fields("D")
              Case "R": CodCampS = "1"
              Case "C": CodCampS = "2"
              Case "P": CodCampS = "3"
            End Select
            Print #NumFile, SetearBlancos(CodCampS, 1, 0, False, FAConLineas);
            Print #NumFile, SetearBlancos(.Fields("Direccion"), 20, 0, False, FAConLineas);
            Print #NumFile, SetearBlancos(.Fields("DirNumero"), 10, 0, False, FAConLineas);
            Print #NumFile, SetearBlancos(.Fields("Ciudad"), 20, 0, False, FAConLineas);
            Print #NumFile, SetearBlancos(.Fields("Prov"), 2, 0, False, FAConLineas);
            Print #NumFile, .Fields("Telefono"); ', 9, 0, True, FAConLineas);
            Print #NumFile, SetearBlancos(.Fields("SN"), 1, 0, False, FAConLineas);
            Print #NumFile, SetearCeros(.Fields("IngLiqu"), 11, 0, True, FAConLineas, True);
            Print #NumFile, SetearCeros(.Fields("AporteP"), 9, 0, True, FAConLineas, True);
            Print #NumFile, SetearCeros(.Fields("BaseImp"), 11, 0, True, FAConLineas, True);
            Print #NumFile, SetearCeros(.Fields("Reten"), 9, 0, True, FAConLineas, True);
            Print #NumFile, SetearBlancos(MidStrg(.Fields("Fecha"), 1, 4), 4, 0, False, FAConLineas);
            If .Fields("Reten") = 0 Then
                Print #NumFile, "00"
            Else
                Print #NumFile, SetearCeros(.Fields("Retenciones"), 2, 0, True, FAConLineas)
            End If
         Case "NO"
            'MsgBox "."
            Print #NumFile, SetearCeros(.Fields("RUC"), 13, 0, True, FAConLineas);
            Print #NumFile, SetearCeros(.Fields("CI_RUC"), 13, 0, True, FAConLineas);
            Select Case .Fields("D")
              Case "R": CodCampS = "1"
              Case "C": CodCampS = "2"
              Case "P": CodCampS = "3"
            End Select
            Print #NumFile, SetearBlancos(CodCampS, 1, 0, False, FAConLineas);
            Print #NumFile, SetearCeros(.Fields("Valor_Fact"), 11, 0, True, FAConLineas, True);
            Print #NumFile, SetearCeros(.Fields("Valor_Ret"), 9, 0, True, FAConLineas, True);
            Print #NumFile, SetearCeros(.Fields("TD"), 3, 0, True, FAConLineas);
           ' MsgBox .Fields("Fecha")
            Print #NumFile, SetearCeros(MidStrg(.Fields("Fecham"), 1, 2), 2, 0, True, FAConLineas);
            Print #NumFile, SetearBlancos(MidStrg(.Fields("Fechaa"), 1, 4), 4, 0, False, FAConLineas);
            Print #NumFile, SetearCeros(.Fields("Retenciones"), 3, 0, True, FAConLineas)
         Case "ID"
            Print #NumFile, SetearCeros(.Fields("RUC"), 13, 0, True, FAConLineas);
            Print #NumFile, SetearCeros(Periodo_No, 3, 0, True, FAConLineas);
            Print #NumFile, SetearBlancos(.Fields("Empresa"), 60, 0, False, FAConLineas);
            Print #NumFile, .Fields("Telefono1"); ' FAConLineas;
            Print #NumFile, .Fields("FAX");
            Cadena = .Fields("Email")
            If Cadena = "." Then
                Print #NumFile, Space(60)
            Else
                Print #NumFile, SetearBlancos(.Fields("Email"), 60, 0, False, FAConLineas)
            End If
         Case "1"   ' transacciones locales
            tt = .Fields("TT")
            Print #NumFile, SetearCeros(.Fields("RUC"), 13, 0, True, FAConLineas);
            Print #NumFile, SetearCeros(Periodo_No, 3, 0, True, FAConLineas);
            If .Fields("TT") = "V" Then    ' ventas
               Select Case .Fields("CSec")
                      Case "R":  Cadena = "04"
                      Case "C":  Cadena = "05"
                      Case "P":  Cadena = "06"
                      Case "O":  Cadena = "07"
                End Select
            Else    ' compras
                Select Case .Fields("CSec")
                       Case "R":  Cadena = "01"
                       Case "C":  Cadena = "02"
                       Case "P":  Cadena = "03"
                End Select
            End If
            Print #NumFile, SetearCeros(Cadena, 2, 0, True, FAConLineas);
            Print #NumFile, SetearCeros(.Fields("CI_RUC"), 13, 0, True, FAConLineas);
            Print #NumFile, SetearCeros(.Fields("TD"), 2, 0, True, FAConLineas);
            Print #NumFile, SetearBlancos(MidStrg(.Fields("FechaE_"), 1, 2), 2, 0, False, FAConLineas);
            Print #NumFile, SetearBlancos(MidStrg(.Fields("FechaE_"), 4, 2), 2, 0, False, FAConLineas);
            Print #NumFile, SetearBlancos(MidStrg(.Fields("FechaE_"), 7, 4), 4, 0, False, FAConLineas);
            Print #NumFile, SetearBlancos(MidStrg(.Fields("FechaR_"), 1, 2), 2, 0, False, FAConLineas);
            Print #NumFile, SetearBlancos(MidStrg(.Fields("FechaR_"), 4, 2), 2, 0, False, FAConLineas);
            Print #NumFile, SetearBlancos(MidStrg(.Fields("FechaR_"), 7, 4), 4, 0, False, FAConLineas);
            If tt = "C" Then
                Print #NumFile, SetearCeros(.Fields("Serie_"), 6, 0, True, FAConLineas);
                Print #NumFile, SetearCeros(.Fields("Secuencial_"), 7, 0, True, FAConLineas);
                Print #NumFile, SetearCeros(.Fields("Autorizacion_"), 10, 0, True, FAConLineas);
                Print #NumFile, SetearCeros(.Fields("IdenCT"), 2, 0, False, FAConLineas);
                Print #NumFile, SetearBlancos(.Fields("Valor_Fact"), 12, 0, True, FAConLineas, True);
                Print #NumFile, SetearCeros(.Fields("CPorc"), 1, 0, True, FAConLineas);
                Print #NumFile, SetearBlancos(.Fields("BImpotcero"), 12, 0, True, FAConLineas, True);
                Print #NumFile, SetearBlancos(.Fields("Valor_Ret"), 12, 0, True, FAConLineas, True);
                Print #NumFile, SetearBlancos(.Fields("MontoICE"), 12, 0, True, FAConLineas, True);
                Print #NumFile, SetearBlancos(.Fields("MontoIVA1"), 12, 0, True, FAConLineas, True);
                Print #NumFile, SetearCeros(.Fields("PorRetIVA1"), 1, 0, True, FAConLineas);
                Print #NumFile, SetearBlancos(.Fields("MontoRetIVA1"), 12, 0, True, FAConLineas, True);
                Print #NumFile, SetearBlancos(.Fields("MontoIVA2"), 12, 0, True, FAConLineas, True);
                Print #NumFile, SetearCeros(.Fields("PorRetIVA2"), 1, 0, True, FAConLineas);
                Print #NumFile, SetearBlancos(.Fields("MontoRetIVA2"), 12, 0, True, FAConLineas, True);
                Print #NumFile, SetearCeros(.Fields("Dev"), 1, 0, True, FAConLineas);
                Print #NumFile, Space(36)
            Else
                Si_No = True
                Cliente = .Fields("Codigo")
                Band = True
                Print #NumFile, Space(23);
                Print #NumFile, SetearCeros(.Fields("IdenCT"), 2, 0, False, FAConLineas);
                Codigo1 = ""
                Codigo2 = ""
                Codigo3 = ""
                Total = 0
                BImpotCero = 0
                Valor_Ret = 0
                MontoICE = 0
                No_Reg = 0
                Do While Si_No And Band
                    'No_Reg = No_Reg + .Fields("No_Reg")
                    Cadena = CStr(No_Reg)
                    Cadena = String$(12 - Len(Cadena), "0") & Cadena
                    Total = Total + .Fields("Valor_Fact")
                    BImpotCero = BImpotCero + .Fields("BImpotcero")
                    Valor_Ret = Valor_Ret + .Fields("Valor_Ret")
                    MontoICE = MontoICE + .Fields("MontoICE")
                    TD = .Fields("TD")
                    If TD = "18" Then Codigo1 = Cadena
                    If TD = "05" Then Codigo2 = Cadena
                    If TD = "04" Then Codigo3 = Cadena
                    .MoveNext
                    If .EOF Then
                        Si_No = False
                        Band = False
                    ElseIf Cliente <> .Fields("Codigo") Or Fecha <> .Fields("FechaE_") Then
                        Si_No = False
                        Band = False
                    End If
                Loop
                .MovePrevious
                If Codigo1 = "" Then Codigo1 = String$(12, "0")
                If Codigo2 = "" Then Codigo2 = String$(12, "0")
                If Codigo3 = "" Then Codigo3 = String$(12, "0")
                Cadena = Codigo1 & Codigo3 & Codigo2
                Print #NumFile, SetearBlancos(CStr(Total), 12, 0, True, FAConLineas, True);
                Print #NumFile, SetearCeros(.Fields("CPorc"), 1, 0, True, FAConLineas);
                Print #NumFile, SetearBlancos(CStr(BImpotCero), 12, 0, True, FAConLineas, True);
                Print #NumFile, SetearBlancos(CStr(Valor_Ret), 12, 0, True, FAConLineas, True);
                Print #NumFile, SetearBlancos(CStr(MontoICE), 12, 0, True, FAConLineas, True);
                Print #NumFile, Space(50);
                Print #NumFile, SetearCeros(.Fields("Dev"), 1, 0, True, FAConLineas);
                Print #NumFile, Cadena
            End If
         Case "2"
            Print #NumFile, SetearCeros(.Fields("RUC"), 13, 0, True, FAConLineas);
            Print #NumFile, SetearCeros(Periodo_No, 3, 0, True, FAConLineas);
             If .Fields("TT") = "V" Then
               Select Case .Fields("CSec")
                      Case "R":  Cadena = "04"
                      Case "C":  Cadena = "05"
                      Case "P":  Cadena = "06"
                      Case "O":  Cadena = "07"
                End Select
            Else
                Select Case .Fields("CSec")
                       Case "R":  Cadena = "01"
                       Case "C":  Cadena = "02"
                       Case "P":  Cadena = "03"
                End Select
            End If
            Print #NumFile, SetearCeros(Cadena, 2, 0, True, FAConLineas);
            Print #NumFile, SetearCeros(.Fields("CI_RUC"), 13, 0, True, FAConLineas);
            Print #NumFile, SetearCeros(.Fields("TD"), 2, 0, True, FAConLineas);
            Print #NumFile, SetearBlancos(MidStrg(.Fields("FechaE_"), 1, 2), 2, 0, False, FAConLineas);
            Print #NumFile, SetearBlancos(MidStrg(.Fields("FechaE_"), 4, 2), 2, 0, False, FAConLineas);
            Print #NumFile, SetearBlancos(MidStrg(.Fields("FechaE_"), 7, 4), 4, 0, False, FAConLineas);
            Print #NumFile, SetearBlancos(MidStrg(.Fields("FechaR_"), 1, 2), 2, 0, False, FAConLineas);
            Print #NumFile, SetearBlancos(MidStrg(.Fields("FechaR_"), 4, 2), 2, 0, False, FAConLineas);
            Print #NumFile, SetearBlancos(MidStrg(.Fields("FechaR_"), 7, 4), 4, 0, False, FAConLineas);
            If MidStrg(.Fields("Cambio"), 1, 2) = "18" Then
               Print #NumFile, Space(46)
            Else
               Print #NumFile, SetearCeros(.Fields("Serie_"), 6, 0, True, FAConLineas);
               Print #NumFile, SetearCeros(.Fields("Secuencial_"), 7, 0, True, FAConLineas);
               Cadena1 = MidStrg(.Fields("Cambio"), 1, 2)
               Cadena2 = MidStrg(.Fields("Cambio"), 3, Len(.Fields("Cambio")))
               SQL2 = "SELECT TD,FechaE,Serie,Secuencial,Autorizacion " _
                    & "FROM Trans_Retenciones " _
                    & "WHERE TD = '" & Cadena1 & "' " _
                    & "AND Secuencial = '" & Cadena2 & "' " _
                    & "AND Item = '" & NumEmpresa & "' "
               Select_Adodc FTalonIVA.AdoNDC, SQL2
              ' MsgBox FTalonIVA.AdoNDC.Recordset.RecordCount
               With FTalonIVA.AdoNDC.Recordset
                 If .RecordCount > 0 Then
                    .MoveFirst
                    Do While Not .EOF
                        Print #NumFile, SetearCeros(.Fields("TD"), 2, 0, True, FAConLineas);
                        Print #NumFile, SetearBlancos(MidStrg(.Fields("FechaE"), 1, 2), 2, 0, False, FAConLineas);
                        Print #NumFile, SetearBlancos(MidStrg(.Fields("FechaE"), 4, 2), 2, 0, False, FAConLineas);
                        Print #NumFile, SetearBlancos(MidStrg(.Fields("FechaE"), 7, 4), 4, 0, False, FAConLineas);
                        Print #NumFile, SetearCeros(.Fields("Serie"), 6, 0, True, FAConLineas);
                        Print #NumFile, SetearCeros(.Fields("Secuencial"), 7, 0, True, FAConLineas);
                        Print #NumFile, SetearCeros(.Fields("Autorizacion"), 10, 0, True, FAConLineas)
                        .MoveNext
                    Loop
                 Else
                    'MsgBox "Nota de C/D mal ingresada "
                 End If
               End With
            End If
          Case "3"
              Print #NumFile, SetearCeros(.Fields("RUC"), 13, 0, True, FAConLineas);
              Print #NumFile, SetearCeros(Periodo_No, 3, 0, True, FAConLineas);
              If .Fields("TD") = "17" Then
                  Cadena = "09"
              Else
                  Cadena = "08"
              End If
              Print #NumFile, Cadena;
              Print #NumFile, SetearBlancos(MidStrg(.Fields("Fecha"), 1, 2), 2, 0, False, FAConLineas);
              Print #NumFile, SetearBlancos(MidStrg(.Fields("Fecha"), 4, 2), 2, 0, False, FAConLineas);
              Print #NumFile, SetearBlancos(MidStrg(.Fields("Fecha"), 7, 4), 4, 0, False, FAConLineas);
              Print #NumFile, SetearCeros(.Fields("TD"), 2, 0, True, FAConLineas);
              Print #NumFile, SetearCeros(.Fields("Secuencial"), 7, 0, True, FAConLineas);
              Print #NumFile, SetearCeros(.Fields("Aduana"), 16, 0, True, FAConLineas);
              Print #NumFile, SetearCeros(.Fields("IdenCT"), 2, 0, True, FAConLineas);
              TD = .Fields("TD")
               Select Case TD
                 Case "16"
                    Print #NumFile, SetearBlancos(.Fields("Valor_Cif_Fob"), 12, 0, True, FAConLineas, True);
                    Print #NumFile, Space(27);
                    
                 Case "17"
                    Print #NumFile, SetearBlancos(.Fields("Valor_Cif_Fob"), 12, 0, True, FAConLineas, True);
                    Print #NumFile, SetearBlancos(.Fields("Valor_Iva"), 12, 0, True, FAConLineas, True);
                    Print #NumFile, SetearBlancos(.Fields("MontoICE"), 12, 0, True, FAConLineas, True);
                    Print #NumFile, SetearCeros(.Fields("CodIce"), 1, 0, True, FAConLineas);
                    Print #NumFile, SetearCeros(.Fields("CodBanco"), 2, 0, True, FAConLineas);
                End Select
              Print #NumFile, SetearBlancos(.Fields("ConvInt"), 1, 0, False, FAConLineas);
              Print #NumFile, SetearBlancos(.Fields("Dev"), 1, 0, False, FAConLineas);
              TD = .Fields("TD")
               Select Case TD
                 Case "16"
                    Print #NumFile, SetearBlancos(MidStrg(.Fields("Fecha"), 1, 2), 2, 0, False, FAConLineas);
                    Print #NumFile, SetearBlancos(MidStrg(.Fields("Fecha"), 4, 2), 2, 0, False, FAConLineas);
                    Print #NumFile, SetearBlancos(MidStrg(.Fields("Fecha"), 7, 4), 4, 0, False, FAConLineas);
                    Print #NumFile, "01";
                    Print #NumFile, SetearCeros(.Fields("Serie"), 6, 0, True, FAConLineas);
                    Print #NumFile, SetearCeros(.Fields("Secuencial"), 7, 0, True, FAConLineas);
                    Print #NumFile, SetearCeros(.Fields("Autorizacion"), 10, 0, True, FAConLineas);
                    Print #NumFile, SetearBlancos(.Fields("Fob"), 12, 0, True, FAConLineas, True)
                 Case "17"
                    Print #NumFile, "                                             "
               End Select
         End Select
         Contador = Contador + 1
        .MoveNext
         Loop
        .MoveFirst
     End If
End With
Close #NumFile
MiFrom.Caption = CaptionOld
RatonNormal
End Sub
Public Function SeisMeses(FechaStr As String) As String
  FechaStr = Format(FechaStr, "dd/mm/yyyy")
  Mes = Month(FechaStr): Anio = Year(FechaStr)
  SeisMeses = "01/" & Format(Mes - 5, "00") & "/" & Format(Anio, "0000")
End Function

Public Sub Imprimir104(AdoQry1 As Adodc, _
                       Semest As Integer)

On Error GoTo Errorhandler
Mensajes = "Seguro de Imprimir en:" & vbCrLf & Printer.DeviceName & "?"
Titulo = "IMPRESION"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
    InicioX = 0.5: InicioY = 0
    RatonReloj
    Escala_Centimetro 1, TipoTimes, 10, True
    'TAConLineas = ProcesaSeteoFormu("104")
    'Iniciamos la consulta de impresion
    With AdoQry1.Recordset
    If .RecordCount > 0 Then
       .MoveFirst
       'mes
       If .Fields("Valor") >= 0 Then PrinterTexto .Fields("PosX") + .Fields("Valor"), .Fields("PosY"), "X"
       .MoveNext
       'año
       PrinterFields0 .Fields("PosX"), .Fields("PosY"), .Fields("Valor")
       .MoveNext
       'semestre
       If .Fields("Valor") = 1 Then PrinterTexto Val(.Fields("PosX")), Val(.Fields("PosY")), "X"
       If .Fields("Valor") = 2 Then PrinterTexto Val(.Fields("PosX")), Val(.Fields("PosY")) + 0.5, "X"
       .MoveNext
       Do While Not .EOF
          If Val(.Fields("Codigo")) < 300 Then
            PrinterFields0 .Fields("PosX"), .Fields("PosY"), .Fields("Valor")
         Else
            If IsNumeric(.Fields("Valor")) Then
                Valor = .Fields("Valor")
                Valor = Format(Valor, "#,##0.00")
                PrinterVariables0 .Fields("PosX") - 3, .Fields("PosY"), Valor
            Else
                PrinterFields0 .Fields("PosX"), .Fields("PosY"), .Fields("Valor")
            End If
         End If
        .MoveNext
        Loop
    End If
    End With
End If
Printer.EndDoc
RatonNormal
Bandera = True
Exit Sub
Errorhandler:
    RatonNormal
    ErrorDeImpresion
    Exit Sub
End Sub

Public Function SepararPorEspacios(Str As String, _
                              Espacios As Integer)
Dim Limite As Long

    Limite = Len(Str)
    Cadena1 = ""
    For I = 1 To Limite - 1
        Caracter = MidStrg(Str, I, 1) & Space(Espacios)
        Cadena1 = Cadena1 & Caracter
    Next
    SepararPorEspacios = Cadena1 & MidStrg(Str, Limite, 1)
End Function

Public Sub ImprimirFormulario(AdoQry1 As Adodc)
On Error GoTo Errorhandler
Mensajes = "Seguro de Imprimir en:" & vbCrLf & Printer.DeviceName & "?"
Titulo = "IMPRESION"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
    InicioX = 0.5: InicioY = 0
    RatonReloj
    Escala_Centimetro 1, TipoTimes, 10, True
    'TAConLineas = ProcesaSeteoFormu("104")
    'Iniciamos la consulta de impresion
    With AdoQry1.Recordset
    If .RecordCount > 0 Then
       .MoveFirst
       Do While Not .EOF
         If Val(.Fields("Codigo")) < 300 Then
            PrinterFields .Fields("PosX"), .Fields("PosY"), .Fields("Valor")
         Else
            If IsNumeric(.Fields("Valor")) Then
                Valor = .Fields("Valor")
                Valor = Format(Valor, "#,##0.00")
                PrinterVariables0 .Fields("PosX") - 3, .Fields("PosY"), Valor
            Else
                PrinterFields .Fields("PosX"), .Fields("PosY"), .Fields("Valor")
            End If
         End If
         .MoveNext
       Loop
    End If
    End With
End If
Printer.EndDoc
RatonNormal
Bandera = True
Exit Sub
Errorhandler:
    RatonNormal
    ErrorDeImpresion
    Exit Sub
End Sub

Public Function NuevoFormulario(Formulario As String, _
                                AdoQuery1 As Adodc, _
                                AdoFormulario As Adodc)
'

  SQL2 = "DELETE *  " _
     & "FROM Catalogo_Formularios " _
     & "WHERE Item = '" & NumEmpresa & "' " _
     & "AND Formulario = '" & Formulario & "' "
  Ejecutar_SQL_SP SQL2
  SQL2 = "SELECT *  " _
     & "FROM Catalogo_Formularios " _
     & "WHERE Item = '" & NumEmpresa & "' " _
     & "AND Formulario = '" & Formulario & "' "
  Select_Adodc AdoQuery1, SQL2
  sSQL = "UPDATE Catalogo_Formularios " _
     & "SET Valor = '-' " _
     & "WHERE Item = '000' " _
     & "AND Formulario = '" & Formulario & "' "
  Ejecutar_SQL_SP sSQL
  SQL2 = "SELECT * " _
     & "FROM Catalogo_Formularios " _
     & "WHERE Item = '000' " _
     & "AND Formulario = '" & Formulario & "' "
  Select_Adodc AdoFormulario, SQL2
  With AdoFormulario.Recordset
    If .RecordCount > 0 Then
       .MoveFirst
        Do While Not .EOF
          SetAddNew AdoQuery1
          For J = 0 To .Fields.Count - 1
              Codigo = .Fields(J).Name
              If Codigo = "Item" Then
                 SetFields AdoQuery1, "Item", NumEmpresa
              Else
                 SetFields AdoQuery1, Codigo, .Fields(J)
              End If
          Next J
          SetUpdate AdoQuery1
         .MoveNext
       Loop
   End If
  End With
End Function
Public Function MoverFormulario(Operacion As String, _
                                Cadena As String)
 sSQL = "UPDATE Catalogo_Formularios " _
       & "SET " & Operacion & " " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Formulario = '" & Cadena & "' "
  Ejecutar_SQL_SP sSQL
End Function

Public Sub PrinterFields0(Xo As Single, _
                         Yo As Single, _
                         AdoTipo As ADODB.Field, _
                         Optional PonerLineas As Boolean)
If ((Xo > 0) And (Yo > 0)) Then
   Distancia = CampoWidth(AdoTipo)
   If StrgFormatoCampo = Ninguno Then StrgFormatoCampo = " "
   If ImpLineaCeros = False Then
      If StrgFormatoCampo = "0" Then StrgFormatoCampo = "0"
      If StrgFormatoCampo = "0.00" Then StrgFormatoCampo = "0.00"
   End If
   LimpiarLinea Xo, Yo, PonerLineas
   Printer.CurrentX = Xo + Distancia
   Printer.CurrentY = Yo
   Printer.Print StrgFormatoCampo
End If
End Sub

Public Sub PrinterVariables0(Xo As Single, Yo As Single, Variable)
On Error GoTo Errorhandler
If ((Xo > 0) And (Yo > 0)) Then
   Distancia = VariableWidth(Variable)
   Printer.CurrentY = Yo
   Printer.CurrentX = Xo + Distancia
   If StrgFormatoVariable = Ninguno Then StrgFormatoVariable = " "
   If ImpLineaCeros = False Then
      If StrgFormatoVariable = "0" Then StrgFormatoVariable = "0"
      If StrgFormatoVariable = "0.00" Then StrgFormatoVariable = "0.00"
   End If
   Printer.Print StrgFormatoVariable
End If
Exit Sub
Errorhandler:
             RatonNormal
             ErrorDeImpresion
             Exit Sub
End Sub
Public Sub ImprimirTalonAnexo(AdoEmp As Adodc, _
                              AdoComp As Adodc, _
                              AdoReten As Adodc, _
                              Vanio As String, _
                              SumaAnulados As Double)
On Error GoTo Errorhandler
Mensajes = "Seguro de Imprimir en:" & vbCrLf & Printer.DeviceName & "?"
Titulo = "IMPRESION"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
    InicioX = 0.5: InicioY = 0
    RatonReloj
    Escala_Centimetro 1, TipoTimes, 8.5, True
    'Iniciamos la consulta de impresion
    Printer.Line (0.2, 2.4)-(19.7, 27.9), Negro, B
    PrinterTexto 3.5, 0.3, "TALON RESUMEN DE ANEXO TRANSACCIONAL SERVICIO DE RENTAS INTERNAS"
    PrinterTexto 5, 0.65, NombreComercial
    PrinterTexto 7, 1, "RUC: " & RUC
    PrinterTexto 2, 1.35, "Certifico que la información contenida en el medio magnético adjunto al presente de Anexo Transaccional para el período:" & Vanio & " "
    PrinterTexto 6, 1.7, "es el fiel reflejo del siguiente reporte"
    PrinterTexto 8, 2.05, "PERIODO:  " & Vanio
    PrinterTexto 10, 2.45, "COMPRAS"
    Printer.Line (0.2, 2.8)-(19.7, 3.2), Negro, B
    PrinterTexto 0.2, 2.85, "Cód"
    PrinterTexto 4, 2.85, "Transacción"
    PrinterTexto 12, 2.85, "No Registros"
    PrinterTexto 14.2, 2.85, "BI tarifa 0%"
    PrinterTexto 16, 2.85, "BI tarifa 12%"
    PrinterTexto 18.2, 2.85, "Valor IVA"
    VLinea = 3.3
    ValorIva = 0
    BaseCero = 0
    BaseImpo = 0
    With AdoComp.Recordset
        ValorIva = 0
        BaseCero = 0
        BaseImpo = 0
        If .RecordCount > 1 Then
          .MoveFirst
          Do While Not .EOF
             If .Fields("Cod") = "C" Then
                If .Fields("Codigo") = "04" Then
                    ValorIva = ValorIva - .Fields("Valor_IVA")
                    BaseCero = BaseCero - .Fields("Basecero")
                    BaseImpo = BaseImpo - .Fields("BaseImponible")
                Else
                    ValorIva = ValorIva + .Fields("Valor_IVA")
                    BaseCero = BaseCero + .Fields("Basecero")
                    BaseImpo = BaseImpo + .Fields("BaseImponible")
                End If
                PrinterFields0 0.2, VLinea, .Fields("Codigo")
                PrinterFields0 0.7, VLinea, .Fields("Detalle")
                PrinterFields0 12, VLinea, .Fields("Numero")
                PrinterFields0 13.5, VLinea, .Fields("Basecero")
                PrinterFields0 15.5, VLinea, .Fields("BaseImponible")
                PrinterFields0 17.3, VLinea, .Fields("Valor_IVA")
                VLinea = VLinea + 0.35
             End If
             .MoveNext
          Loop
        End If
    End With
    Printer.FontBold = True
    PrinterTexto 10, VLinea, "TOTAL"
    PrinterVariables0 13.8, VLinea, BaseCero
    PrinterVariables0 15.8, VLinea, BaseImpo
    PrinterVariables0 17.7, VLinea, ValorIva
    Printer.FontBold = False
    PrinterTexto 3, VLinea + 0.35, "Se verificará con los casilleros asignados en la declaración de IVA (form 104) de acuerdo al siguiente esquema:"
    PrinterTexto 6, VLinea + 0.7, "Sustento Crédito Tributario"
    PrinterTexto 13.2, VLinea + 0.7, "Casilleros"
    PrinterTexto 15.8, VLinea + 0.7, "Base Imponible"
    PrinterTexto 18.2, VLinea + 0.7, "Impuesto"
    VLinea = VLinea + 1.1
    With AdoComp.Recordset
        If .RecordCount > 1 Then
          .MoveFirst
          Do While Not .EOF
             If .Fields("Cod") = "TC" Then
                PrinterFields0 0.2, VLinea, .Fields("Detalle")
                PrinterFields0 13.2, VLinea, .Fields("Casilleros")
                PrinterFields0 15.5, VLinea, .Fields("BaseImponible")
                PrinterFields0 17.3, VLinea, .Fields("Valor_IVA")
                VLinea = VLinea + 0.35
             End If
             .MoveNext
          Loop
        End If
    End With
    VLinea = VLinea - 0.25
    ' Ventas
    PrinterTexto 10, VLinea + 0.32, "VENTAS"
    Printer.Line (0.2, VLinea + 0.3)-(19.7, VLinea + 0.3), Negro, B
    Printer.Line (0.2, VLinea + 0.65)-(19.7, VLinea + 1.1), Negro, B
    PrinterTexto 0.2, VLinea + 0.7, "Cód"
    PrinterTexto 4, VLinea + 0.7, "Transacción"
    PrinterTexto 12, VLinea + 0.7, "No Registros"
    PrinterTexto 14.2, VLinea + 0.7, "BI tarifa 0%"
    PrinterTexto 16, VLinea + 0.7, "BI tarifa 12%"
    PrinterTexto 18.2, VLinea + 0.7, "Valor IVA"
    VLinea = Round(VLinea, 2) + 1.5
    ValorIva = 0
    BaseCero = 0
    BaseImpo = 0
    With AdoComp.Recordset
        ValorIva = 0
        BaseCero = 0
        BaseImpo = 0
        If .RecordCount > 1 Then
          .MoveFirst
          Do While Not .EOF
             If .Fields("Cod") = "V" Then
                If .Fields("Codigo") = "04" Then
                    ValorIva = ValorIva - .Fields("Valor_IVA")
                    BaseCero = BaseCero - .Fields("Basecero")
                    BaseImpo = BaseImpo - .Fields("BaseImponible")
                Else
                    ValorIva = ValorIva + .Fields("Valor_IVA")
                    BaseCero = BaseCero + .Fields("Basecero")
                    BaseImpo = BaseImpo + .Fields("BaseImponible")
                End If
                PrinterFields0 0.2, VLinea, .Fields("Codigo")
                PrinterFields0 0.7, VLinea, .Fields("Detalle")
                PrinterFields0 12, VLinea, .Fields("Numero")
                PrinterFields0 13.5, VLinea, .Fields("Basecero")
                PrinterFields0 15.5, VLinea, .Fields("BaseImponible")
                PrinterFields0 17.3, VLinea, .Fields("Valor_IVA")
                VLinea = VLinea + 0.35
             End If
             .MoveNext
          Loop
        End If
    End With
    PrinterTexto 10, VLinea, "TOTAL"
    PrinterVariables0 13.8, VLinea, BaseCero
    PrinterVariables0 15.8, VLinea, BaseImpo
    PrinterVariables0 17.7, VLinea, ValorIva
    PrinterTexto 3, VLinea + 0.35, "Se verificará con los casilleros asignados en la declaración de IVA (form 104) de acuerdo al siguiente esquema:"
    PrinterTexto 6, VLinea + 0.7, "Sustento Crédito Tributario"
    PrinterTexto 13.2, VLinea + 0.7, "Casilleros"
    PrinterTexto 15.8, VLinea + 0.7, "Base Imponible"
    PrinterTexto 18.2, VLinea + 0.7, "Impuesto"
    VLinea = Round(VLinea, 2) + 1.1
    With AdoComp.Recordset
        If .RecordCount > 1 Then
          .MoveFirst
          Do While Not .EOF
             If .Fields("Cod") = "TV" Then
                PrinterFields0 0.2, VLinea, .Fields("Detalle")
                PrinterFields0 13.2, VLinea, .Fields("Casilleros")
                PrinterFields0 15.5, VLinea, .Fields("BaseImponible")
                PrinterFields0 17.3, VLinea, .Fields("Valor_IVA")
                VLinea = VLinea + 0.35
             End If
             .MoveNext
          Loop
        End If
    End With
    ' comprobantes anulados
    VLinea = Round(VLinea, 2) - 0.25
    Printer.Line (0.2, VLinea + 0.3)-(19.7, VLinea + 0.7), Negro, B
    PrinterTexto 8, VLinea + 0.33, "COMPROBANTES ANULADOS"
    VLinea = Round(VLinea, 2) + 0.8
    PrinterTexto 3, VLinea, "Total de comprobantes Anulados en el período informado (No incluye lo dados de baja)"
        PrinterVariables0 17.7, VLinea, FTalonMes.LAbaseimp.Caption
    VLinea = Round(VLinea, 2) + 0.2
    ' importaciones
    Printer.Line (0.2, VLinea + 0.3)-(19.7, VLinea + 0.7), Negro, B
    PrinterTexto 9, VLinea + 0.32, "IMPORTACIONES"
    Printer.Line (0.2, VLinea + 1.1)-(19.7, VLinea + 1.1), Negro, B
    PrinterTexto 0.2, VLinea + 0.72, "Cód"
    PrinterTexto 4, VLinea + 0.72, "Transacción"
    PrinterTexto 14, VLinea + 0.72, "No Registros"
    PrinterTexto 16.2, VLinea + 0.72, "Valor CIF"
    PrinterTexto 18.2, VLinea + 0.72, "Valor IVA"
    VLinea = Round(VLinea, 2) + 1.2
    With AdoComp.Recordset
        ValorIva = 0
        BaseImpo = 0
        If .RecordCount > 1 Then
          .MoveFirst
          Do While Not .EOF
             If .Fields("Cod") = "TI" Then
                If .Fields("Codigo") = "04" Then
                    ValorIva = ValorIva - .Fields("Valor_IVA")
                    BaseImpo = BaseImpo - .Fields("BaseImponible")
                Else
                    ValorIva = ValorIva + .Fields("Valor_IVA")
                    BaseImpo = BaseImpo + .Fields("BaseImponible")
                End If
                PrinterFields0 0.2, VLinea, .Fields("Codigo")
                PrinterFields0 0.7, VLinea, .Fields("Detalle")
                PrinterFields0 14.5, VLinea, .Fields("Numero")
                PrinterFields0 15.5, VLinea, .Fields("BaseImponible")
                PrinterFields0 17.3, VLinea, .Fields("Valor_IVA")
                VLinea = VLinea + 0.35
             End If
             .MoveNext
          Loop
        End If
    End With
    PrinterTexto 10, VLinea, "TOTAL"
    PrinterVariables0 15.7, VLinea, BaseImpo
    PrinterVariables0 17.7, VLinea, ValorIva
    PrinterTexto 3, VLinea + 0.35, "Se verificará con los casilleros asignados en la declaración de IVA (form 104) de acuerdo al siguiente esquema:"
    PrinterTexto 6, VLinea + 0.7, "Sustento Crédito Tributario"
    PrinterTexto 13.2, VLinea + 0.7, "Casilleros"
    PrinterTexto 15.8, VLinea + 0.7, "Base Imponible"
    PrinterTexto 18.2, VLinea + 0.7, "Impuesto"
    VLinea = Round(VLinea, 2) + 1.1
    With AdoComp.Recordset
        If .RecordCount > 1 Then
          .MoveFirst
          Do While Not .EOF
             If .Fields("Cod") = "IT" Then
                PrinterFields0 0.2, VLinea, .Fields("Detalle")
                PrinterFields0 13.2, VLinea, .Fields("Casilleros")
                PrinterFields0 15.5, VLinea, .Fields("BaseImponible")
                PrinterFields0 17.3, VLinea, .Fields("Valor_IVA")
                VLinea = VLinea + 0.35
             End If
             .MoveNext
          Loop
        End If
    End With
    ' exportaciones
    VLinea = Round(VLinea, 2) - 0.25
    Printer.Line (0.2, VLinea + 0.3)-(19.7, VLinea + 0.3), Negro, B
    PrinterTexto 9, VLinea + 0.32, "EXPORTACIONES"
        PrinterTexto 0.2, VLinea + 0.72, "Cód"
    PrinterTexto 4, VLinea + 0.72, "Transacción"
    PrinterTexto 14, VLinea + 0.72, "No Registros"
    PrinterTexto 16.2, VLinea + 0.72, "Valor CIF"
    PrinterTexto 18.2, VLinea + 0.72, "Valor IVA"
    Printer.Line (0.2, VLinea + 0.7)-(19.7, VLinea + 1.1), Negro, B
    VLinea = Round(VLinea, 2) + 1.2
    With AdoComp.Recordset
        ValorIva = 0
        BaseImpo = 0
        If .RecordCount > 1 Then
          .MoveFirst
          Do While Not .EOF
             If .Fields("Cod") = "TE" Then
                If .Fields("Codigo") = "04" Then
                    ValorIva = ValorIva - .Fields("Valor_IVA")
                    BaseImpo = BaseImpo - .Fields("BaseImponible")
                Else
                    ValorIva = ValorIva + .Fields("Valor_IVA")
                    BaseImpo = BaseImpo + .Fields("BaseImponible")
                End If
                PrinterFields0 0.2, VLinea, .Fields("Codigo")
                PrinterFields0 0.7, VLinea, .Fields("Detalle")
                PrinterFields0 14.5, VLinea, .Fields("Numero")
                PrinterFields0 15.5, VLinea, .Fields("BaseImponible")
                PrinterFields0 17.3, VLinea, .Fields("Valor_IVA")
                VLinea = VLinea + 0.35
             End If
             .MoveNext
          Loop
        End If
    End With
    PrinterTexto 10, VLinea, "TOTAL"
    PrinterVariables0 15.7, VLinea, BaseImpo
    PrinterVariables0 17.7, VLinea, ValorIva
    PrinterTexto 3, VLinea + 0.35, "Se verificará con los casilleros asignados en la declaración de IVA (form 104) de acuerdo al siguiente esquema:"
    PrinterTexto 6, VLinea + 0.7, "Sustento Crédito Tributario"
    PrinterTexto 13.2, VLinea + 0.7, "Casilleros"
    PrinterTexto 15.8, VLinea + 0.7, "Base Imponible"
    PrinterTexto 18.2, VLinea + 0.7, "Impuesto"
    VLinea = Round(VLinea, 2) + 1.1
    With AdoComp.Recordset
        If .RecordCount > 1 Then
          .MoveFirst
          Do While Not .EOF
             If .Fields("Cod") = "ET" Then
                PrinterFields0 0.2, VLinea, .Fields("Detalle")
                PrinterFields0 13.2, VLinea, .Fields("Casilleros")
                PrinterFields0 15.5, VLinea, .Fields("BaseImponible")
                PrinterFields0 17.3, VLinea, .Fields("Valor_IVA")
                VLinea = VLinea + 0.35
             End If
             .MoveNext
          Loop
        End If
    End With
    ' emisora de tarjetas de crédito
    VLinea = Round(VLinea, 2) - 0.25
    Printer.Line (0.2, VLinea + 0.3)-(19.7, VLinea + 0.7), Negro, B
    PrinterTexto 8, VLinea + 0.33, "EMPRESA EMISORA DE TARJETAS DE CRÉDITO"
    Printer.Line (0.2, VLinea + 1.1)-(19.7, VLinea + 1.1), Negro, B
    PrinterTexto 0.2, VLinea + 0.72, "Cód"
    PrinterTexto 4, VLinea + 0.72, "Transacción"
    PrinterTexto 14, VLinea + 0.72, "No Registros"
    PrinterTexto 16.2, VLinea + 0.72, "Total Consumo"
    PrinterTexto 18.2, VLinea + 0.72, "Valor IVA"
    VLinea = Round(VLinea, 2) + 1.2
    With AdoComp.Recordset
        ValorIva = 0
        BaseImpo = 0
        If .RecordCount > 1 Then
          .MoveFirst
          Do While Not .EOF
             If .Fields("Cod") = "TT" Then
                If .Fields("Codigo") = "23" Then
                    BaseImpo = BaseImpo - .Fields("BaseImponible")
                    ValorIva = ValorIva - .Fields("Valor_IVA")
                Else
                    BaseImpo = BaseImpo + .Fields("BaseImponible")
                    ValorIva = ValorIva + .Fields("Valor_IVA")
                End If
                PrinterFields0 0.2, VLinea, .Fields("Codigo")
                PrinterFields0 1, VLinea, .Fields("Detalle")
                PrinterFields0 15.5, VLinea, .Fields("BaseImponible")
                PrinterFields0 17.3, VLinea, .Fields("Valor_IVA")
                VLinea = VLinea + 0.35
             End If
                .MoveNext
          Loop
        End If
    End With
    PrinterTexto 10, VLinea, "TOTAL"
    PrinterVariables0 15.7, VLinea, BaseImpo
    PrinterVariables0 17.7, VLinea, ValorIva
    'nueva pagina
    Printer.NewPage
    InicioX = 0.5: InicioY = 0
    RatonReloj
    Printer.FontSize = 7.3
    Escala_Centimetro 1, TipoTimes, 8.5, True
    PrinterTexto 3, 26.9, "_______________________________________"
    PrinterTexto 12, 26.9, "________________________________________"
    PrinterTexto 3, 27.3, "Firma del Contador  RUC: " & RUC_Contador
    PrinterTexto 12, 27.3, "Firma del Representante Legal  CI: " & CI_Representante
    ' FONDOS Y FIDEICOMISOS
    Printer.FontSize = 8.5
    Printer.Line (0.2, 0.2)-(19.7, 25.6), Negro, B
    VLinea = 0.2
    PrinterTexto 8, VLinea + 0.05, "FONDOS Y FIDEICOMISOS"
    VLinea = Round(VLinea, 2) + 0.4
    Printer.Line (0.2, VLinea)-(19.7, VLinea + 0.72), Negro, B
    PrinterTexto 4, VLinea + 0.2, "Tipo de Fideicomiso"
    PrinterTexto 14, VLinea + 0.2, "No Registros"
    PrinterTexto 15.7, VLinea + 0.02, "Total Beneficio"
    PrinterTexto 17.7, VLinea + 0.2, "Valor Retenido"
    PrinterTexto 15.9, VLinea + 0.32, "Individual"
    VLinea = Round(VLinea, 2) + 0.8
    With AdoComp.Recordset
        ValorIva = 0
        BaseImpo = 0
        If .RecordCount > 1 Then
          .MoveFirst
          Do While Not .EOF
             If .Fields("Cod") = "TF" Then
                BaseImpo = BaseImpo + .Fields("BaseImponible")
                ValorIva = ValorIva + .Fields("Valor_IVA")
                PrinterFields0 0.2, VLinea, .Fields("Detalle")
                PrinterFields0 15, VLinea, .Fields("BaseImponible")
                PrinterFields0 17.3, VLinea, .Fields("Valor_IVA")
                VLinea = VLinea + 0.35
             End If
                .MoveNext
          Loop
        End If
    End With
    PrinterTexto 13, VLinea, "TOTAL"
    PrinterVariables0 15.2, VLinea, BaseImpo
    PrinterVariables0 17.7, VLinea, ValorIva
    VLinea = Round(VLinea, 2) + 0.4
    ' RETENCIONES EN LA FUENTE
    PrinterTexto 4, VLinea, "RESUMEN DE RETENCIONES -- RETENCION EN LA FUENTE DEL IMPUESTO A LA RENTA"
    Printer.Line (0.2, VLinea)-(19.7, VLinea), Negro, B
    Printer.Line (0.2, VLinea + 0.4)-(19.7, VLinea + 0.8), Negro, B
    PrinterTexto 0.2, VLinea + 0.43, "Cód"
    PrinterTexto 4, VLinea + 0.43, "Concepto de Retención"
    PrinterTexto 15.6, VLinea + 0.43, "Base Imponible"
    PrinterTexto 17.7, VLinea + 0.43, "Valor Retenido"
    VLinea = Round(VLinea, 2) + 0.9
    With AdoReten.Recordset
        ValorIva = 0
        BaseImpo = 0
        If .RecordCount > 1 Then
          .MoveFirst
          Do While Not .EOF
                BaseImpo = BaseImpo + .Fields("Base_Imp")
                ValorIva = ValorIva + .Fields("Retencion")
                PrinterFields0 0.2, VLinea, .Fields("Codigo")
                PrinterFields0 1.2, VLinea, .Fields("Tipo_Ret")
                PrinterFields0 15, VLinea, .Fields("Base_Imp")
                PrinterFields0 17.3, VLinea, .Fields("Retencion")
                VLinea = VLinea + 0.3
                .MoveNext
          Loop
        End If
    End With
    PrinterTexto 13, VLinea, "TOTAL"
    PrinterVariables0 15.2, VLinea, BaseImpo
    PrinterVariables0 17.7, VLinea, ValorIva
    ' resumen de retenciones retencion en la fuente del iva
    VLinea = Round(VLinea, 2) + 0.35
    PrinterTexto 5, VLinea + 0.02, "RESUMEN DE RETENCIONES -- RETENCION EN LA FUENTE DE IVA"
    PrinterTexto 0.2, VLinea + 0.42, "Operación"
    PrinterTexto 4, VLinea + 0.42, "Concepto de Retención"
    PrinterTexto 14, VLinea + 0.42, "Base Imponible"
    PrinterTexto 16, VLinea + 0.42, "% Retencion"
    PrinterTexto 17.8, VLinea + 0.42, "Valor Retenido"
    PrinterTexto 6, VLinea + 0.82, "AGENTE DE RETENCION EN EL PERIODO"
    Printer.Line (0.2, VLinea)-(19.7, VLinea + 0.4), Negro, B
    Printer.Line (0.2, VLinea + 0.8)-(19.7, VLinea + 1.2), Negro, B
    VLinea = Round(VLinea, 2) + 1.3
    With AdoComp.Recordset
        ValorIva = 0
        BaseImpo = 0
        If .RecordCount > 1 Then
          .MoveFirst
          Do While Not .EOF
             If .Fields("Cod") = "TA" Then
                ValorIva = ValorIva + .Fields("Valor_IVA")
                BaseImpo = BaseImpo + .Fields("BaseImponible")
                PrinterFields0 0.2, VLinea, .Fields("Detalle")
                PrinterFields0 13.2, VLinea, .Fields("BaseImponible")
                PrinterFields0 16.5, VLinea, .Fields("Casilleros")
                PrinterFields0 17.3, VLinea, .Fields("Valor_IVA")
                VLinea = VLinea + 0.35
             End If
             .MoveNext
          Loop
        End If
    End With
    PrinterTexto 13, VLinea, "TOTAL"
    PrinterVariables0 13.4, VLinea, BaseImpo
    PrinterVariables0 17.7, VLinea, ValorIva
    VLinea = Round(VLinea, 2)
    PrinterTexto 6, VLinea + 0.42, "QUE LE EFECTUARON EN EL PERIODO"
    Printer.Line (0.2, VLinea + 0.4)-(19.7, VLinea + 0.8), Negro, B
    VLinea = Round(VLinea, 2) + 0.9
    With AdoComp.Recordset
        ValorIva = 0
        BaseImpo = 0
        If .RecordCount > 1 Then
          .MoveFirst
          Do While Not .EOF
             If .Fields("Cod") = "TB" Then
                ValorIva = ValorIva + .Fields("Valor_IVA")
                BaseImpo = BaseImpo + .Fields("BaseImponible")
                PrinterFields0 0.2, VLinea, .Fields("Detalle")
                PrinterFields0 13.2, VLinea, .Fields("BaseImponible")
                PrinterFields0 16.5, VLinea, .Fields("Casilleros")
                PrinterFields0 17.3, VLinea, .Fields("Valor_IVA")
                VLinea = VLinea + 0.35
             End If
             .MoveNext
          Loop
        End If
    End With
    PrinterTexto 13, VLinea, "TOTAL"
    PrinterVariables0 13.4, VLinea, BaseImpo
    PrinterVariables0 17.7, VLinea, ValorIva
    Printer.FontSize = 7.5
    VLinea = 25.7
    PrinterTexto 0.5, VLinea, "Declaro que los datos contenidos en este anexo son verdaderos, por lo que asumo la responsabilidad correspondiente, de acuerdo a lo establecido"
    PrinterTexto 5, VLinea + 0.3, "en el Art. 101 de la Codificación de la Ley de Régimen Tributario Interno"
'    Printer.FontSize = 8
'    PrinterTexto 6.5, 27.8, "RUC:"
'    PrinterVariables 7.5, 27.8, RUC_Contador
'    PrinterTexto 16, 27.8, "CI:"
'    PrinterVariables 16.8, 27.8, CI_Representante
End If
Printer.EndDoc
Exit Sub
Errorhandler:
    RatonNormal
    ErrorDeImpresion
    Exit Sub
End Sub

Public Sub GenerarArchivoAT(MiFrom As Form, _
                            DtaAux As Adodc, _
                            Carpeta As String, _
                            RutaGeneraFile As String, _
                            NombreFile As String, _
                            Tipo As Boolean)
Dim NumFile As Integer
Dim CaptionOld As String
Dim ValorBool As String
Dim CodCampS As String
Dim RetFte As Adodc
RatonReloj
FAConLineas = False
CaptionOld = UCaseStrg(LeftStrg(MiFrom.Caption, 3))
RutaGeneraFile = RutaGeneraFile & Carpeta
Cadena = " "
If RutaGeneraFile <> "A:" Then Cadena = Dir(RutaGeneraFile, vbDirectory)
If Cadena = "" Then MkDir (RutaGeneraFile)
If RutaGeneraFile = "A:" Then
  RutaGeneraFile = RutaGeneraFile & NombreFile
Else
  RutaGeneraFile = RutaGeneraFile & "\" & NombreFile
End If
NumFile = FreeFile
Contador = 0
Open RutaGeneraFile For Output As #NumFile ' Abre el archivo.
With FTalonMes.AdoEmpresa.Recordset
     ReDim TipoC(.Fields.Count - 1) As Campos_Tabla
     For I = 0 To .Fields.Count - 1
         TipoC(I).Campo = .Fields(I).Name
         TipoC(I).Ancho = AnchoTipoCampoTexto(.Fields(I))
     Next I
     FAConLineas = False
     Cadena = """"
     Cadena = MidStrg(Cadena, 1, 1)
     'MsgBox Cadena
     If .RecordCount > 0 Then
        .MoveFirst
         Do While Not .EOF
            Print #NumFile, "<iva xmlns:xsi=" & Cadena & "http://www.w3.org/2006/XMLSchema-instance" & Cadena & ">"
            Print #NumFile, Space(5); "<numeroRuc>" & .Fields("RUC") & "</numeroRuc>"
            Print #NumFile, Space(5); "<razonSocial>" & .Fields("Nombre_Comercial") & "</razonSocial>"
            Print #NumFile, Space(5); "<direccionMatriz>" & .Fields("Direccion") & "</direccionMatriz>"
            Print #NumFile, Space(5); "<telefono>" & .Fields("Telefono1") & "</telefono>"
            Print #NumFile, Space(5); "<fax>" & .Fields("FAX") & "</fax>"
            Print #NumFile, Space(5); "<email>" & .Fields("Email") & "</email>"
            If Len(.Fields("CI_Representante")) = 10 Then
               Print #NumFile, Space(5); "<tpIdRepre>C</tpIdRepre>"
            Else
               Print #NumFile, Space(5); "<tpIdRepre>P</tpIdRepre>"
            End If
            Print #NumFile, Space(5); "<idRepre>" & .Fields("CI_Representante") & "</idRepre>"
            Print #NumFile, Space(5); "<rucContador>" & .Fields("RUC_Contador") & "</rucContador>"
            Print #NumFile, Space(5); "<anio>" & Vanio & "</anio>"
            Print #NumFile, Space(5); "<mes>" & Format(Vmes, "00") & "</mes>"
        .MoveNext
         Loop
       '  .MoveFirst
     End If
End With
' datos compras
With FTalonMes.AdoCompras.Recordset
If .RecordCount > 0 Then
   .MoveFirst
   Print #NumFile, Space(5); "<compras>"
   Do While Not .EOF
      Print #NumFile, Space(7); "<detalleCompras>"
        Print #NumFile, Space(10); "<codSustento>" & .Fields("IdenCT") & "</codSustento>"
        Print #NumFile, Space(10); "<devIva>" & .Fields("Dev") & "</devIva>"
        Select Case .Fields("CTD")
        Case "R": Cadena2 = "01"
        Case "C": Cadena2 = "02"
        Case "P": Cadena2 = "03"
        End Select
        Print #NumFile, Space(10); "<tpIdProv>" & Cadena2 & "</tpIdProv>"
        Print #NumFile, Space(10); "<idProv>" & .Fields("CI_RUC") & "</idProv>"
        Print #NumFile, Space(10); "<tipoComprobante>" & Val(.Fields("TTD")) & "</tipoComprobante>"
        Print #NumFile, Space(10); "<fechaRegistro>" & .Fields("Fecha") & "</fechaRegistro>"
        Print #NumFile, Space(10); "<establecimiento>" & MidStrg(.Fields("Serie"), 1, 3) & "</establecimiento>"
        Print #NumFile, Space(10); "<puntoEmision>" & MidStrg(.Fields("Serie"), 4, 3) & "</puntoEmision>"
        Print #NumFile, Space(10); "<secuencial>" & .Fields("Secuencial") & "</secuencial>"
        Print #NumFile, Space(10); "<fechaEmision>" & .Fields("FechaE") & "</fechaEmision>"
        Print #NumFile, Space(10); "<autorizacion>" & .Fields("Autorizacion") & "</autorizacion>"
        Print #NumFile, Space(10); "<fechaCaducidad>" & MidStrg(.Fields("FechaC"), 4, 7) & "</fechaCaducidad>"
        Print #NumFile, Space(10); "<baseImponible>" & Format(.Fields("BImpotCero"), "#0.00") & "</baseImponible>"
        Print #NumFile, Space(10); "<baseImpGrav>" & Format(.Fields("Valor_Fact"), "#0.00") & "</baseImpGrav>"
        Print #NumFile, Space(10); "<porcentajeIva>" & .Fields("CodPorc") & "</porcentajeIva>"
        Print #NumFile, Space(10); "<montoIva>" & Format(.Fields("MontoIVA1") + .Fields("MontoIVA2"), "#0.00") & "</montoIva>"
        Print #NumFile, Space(10); "<baseImpIce>" & Format(.Fields("MontoICE"), "#0.00") & "</baseImpIce>"
        Print #NumFile, Space(10); "<porcentajeIce>" & MidStrg(.Fields("CPorcICE"), 1, 1) & "</porcentajeIce>"
        Print #NumFile, Space(10); "<montoIce>" & Format(.Fields("RetICE"), "#0.00") & "</montoIce>"
        Print #NumFile, Space(10); "<montoIvaBienes>" & Format(.Fields("MontoIVA1"), "#0.00") & "</montoIvaBienes>"
        Print #NumFile, Space(10); "<porRetBienes>" & .Fields("PorRetIVA1") & "</porRetBienes>"
        Print #NumFile, Space(10); "<valorRetBienes>" & Format(.Fields("MontoRetIVA1"), "#0.00") & "</valorRetBienes>"
        Print #NumFile, Space(10); "<montoIvaServicios>" & Format(.Fields("MontoIVA2"), "#0.00") & "</montoIvaServicios>"
        Print #NumFile, Space(10); "<porRetServicios>" & .Fields("PorRetIVA2") & "</porRetServicios>"
        Print #NumFile, Space(10); "<valorRetServicios>" & Format(.Fields("MontoRetIVA2"), "#0.00") & "</valorRetServicios>"
        NumCompRet = .Fields("Retencion_No")
        FechaRet = .Fields("Fecha")
        ' retencion en la fuente del iva compras
        Codigo = .Fields("Comprobante")
        sSQL = "SELECT TD,Valor_Fact,CodPorc,Porc,Valor_Ret,Serie,Secuencial,Autorizacion,FechaE,Fecha,Comprobante " _
             & "FROM Trans_Retenciones " _
             & "WHERE Item = '" & NumEmpresa & "' " _
             & "AND Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
             & "AND CodigoTR = 'RF' AND T <> 'A' " _
             & "AND Comprobante = '" & Codigo & "' "
        Select_Adodc FTalonMes.AdoRetFte, sSQL
        'MsgBox sSQL
        With FTalonMes.AdoRetFte.Recordset
                 If .RecordCount > 0 Then
                    .MoveFirst
                    No_RegC = 1
                    Print #NumFile, Space(10); "<air>"
                    Do While Not .EOF
                       Print #NumFile, Space(13); "<detalleAir>"
                       If No_RegC = 2 And TD = .Fields("TD") Then MsgBox "Revise el Comprobante " & .Fields("Comprobante")
                       Print #NumFile, Space(15); "<codRetAir>" & .Fields("TD") & "</codRetAir>"
                       TD = .Fields("TD")
                       No_RegC = No_RegC + 1
                       Print #NumFile, Space(15); "<baseImpAir>" & Format(.Fields("Valor_Fact"), "#.00") & "</baseImpAir>"
                       Print #NumFile, Space(15); "<porcentajeAir>" & Val(.Fields("Porc")) * 100 & "</porcentajeAir>"
                       Print #NumFile, Space(15); "<valRetAir>" & Format(.Fields("Valor_Ret"), "#.00") & "</valRetAir>"
                       Print #NumFile, Space(13); "</detalleAir>"
                       .MoveNext
                    Loop
                    Print #NumFile, Space(10); "</air>"
                    .MoveFirst
                    No_RegC = 1
                    Do While Not .EOF
                       Print #NumFile, Space(10); "<estabRetencion" & No_RegC & ">" & MidStrg(.Fields("Serie"), 1, 3) & "</estabRetencion" & No_RegC & ">"
                       Print #NumFile, Space(10); "<ptoEmiRetencion" & No_RegC & ">" & MidStrg(.Fields("Serie"), 1, 3) & "</ptoEmiRetencion" & No_RegC & ">"
                       SerieRet = .Fields("Serie")
                      ' MsgBox .Fields("Secuencial")
                      ' If .Fields("Secuencial") = "." Then Cadena1 = "0" Else Cadena1 = .Fields("Secuencial")
                       Print #NumFile, Space(10); "<secRetencion" & No_RegC & ">" & .Fields("Secuencial") & "</secRetencion" & No_RegC & ">"
                       Print #NumFile, Space(10); "<autRetencion" & No_RegC & ">" & .Fields("Autorizacion") & "</autRetencion" & No_RegC & ">"
                       AutorizaRet = .Fields("Autorizacion")
                       Print #NumFile, Space(10); "<fechaEmiRet" & No_RegC & ">" & .Fields("FechaE") & "</fechaEmiRet" & No_RegC & ">"
                       No_RegC = No_RegC + 1
                       If No_RegC >= 3 Then .MoveLast
                       .MoveNext
                    Loop
                 Else
                     Print #NumFile, Space(10); "<air>"
                     Print #NumFile, Space(10); "</air>"
                     If SerieRet = "" Or SerieRet = "." Then FDatos.Show 1
                     Print #NumFile, Space(10); "<estabRetencion1>" & MidStrg(SerieRet, 1, 3) & "</estabRetencion1>"
                     Print #NumFile, Space(10); "<ptoEmiRetencion1>" & MidStrg(SerieRet, 1, 3) & "</ptoEmiRetencion1>"
                     Print #NumFile, Space(10); "<secRetencion1>0</secRetencion1>"
                     Print #NumFile, Space(10); "<autRetencion1>" & AutorizaRet & "</autRetencion1>"
                     Print #NumFile, Space(10); "<fechaEmiRet1>" & FechaRet & "</fechaEmiRet1>"
                 End If
        End With
        ' notas de debito y credito compras
            Select Case .Fields("TTD")
             Case "04", "05"
                 Codigo = .Fields("Cambio")
                 If FTalonMes.AdoCambio.Recordset.RecordCount > 0 Then
                    FTalonMes.AdoCambio.Recordset.MoveFirst
                    FTalonMes.AdoCambio.Recordset.Find ("CompMod = '" & Codigo & "' ")
                    If Not FTalonMes.AdoCambio.Recordset.EOF Then
                        Cadena1 = Val(MidStrg(Codigo, 1, 2))
                        Cadena2 = Val(MidStrg(Codigo, 3, 7))
                        Print #NumFile, Space(10); "<docModificado>" & Cadena1 & "</docModificado>"
                        Print #NumFile, Space(10); "<fechaEmiModificado>" & .Fields("FechaE") & "</fechaEmiModificado>"
                        Print #NumFile, Space(10); "<estabModificado>" & MidStrg(.Fields("Serie"), 1, 3) & "</estabModificado>"
                        Print #NumFile, Space(10); "<ptoEmiModificado>" & MidStrg(.Fields("Serie"), 4, 3) & "</ptoEmiModificado>"
                        Print #NumFile, Space(10); "<secModificado>" & Cadena2 & "</secModificado>"
                        Print #NumFile, Space(10); "<autModificado>" & .Fields("Autorizacion") & "</autModificado>"
                        FTalonMes.AdoCambio.Recordset.MoveNext
                    End If
                 End If
             Case Else
                 Print #NumFile, Space(10); "<docModificado>0</docModificado>"
                 Print #NumFile, Space(10); "<fechaEmiModificado>00/00/0000</fechaEmiModificado>"
                 Print #NumFile, Space(10); "<estabModificado>000</estabModificado>"
                 Print #NumFile, Space(10); "<ptoEmiModificado>000</ptoEmiModificado>"
                 Print #NumFile, Space(10); "<secModificado>0000000</secModificado>"
                 Print #NumFile, Space(10); "<autModificado>0000000000</autModificado>"
            End Select
        ' gasto electoral compras
            Codigo = .Fields("Comprobante")
                 If FTalonMes.AdoRI2S.Recordset.RecordCount > 0 Then
                    FTalonMes.AdoRI2S.Recordset.MoveFirst
                    FTalonMes.AdoRI2S.Recordset.Find ("Comprobante = '" & Codigo & "' ")
                    If Not FTalonMes.AdoRI2S.Recordset.EOF Then
                        Print #NumFile, Space(10); "<contrato>" & FTalonMes.AdoRI2S.Recordset.Fields("Contrato") & "</contrato>"
                        Print #NumFile, Space(10); "<montoTituloOneroso>" & FTalonMes.AdoRI2S.Recordset.Fields("Titulo_Oneroso") & "</montoTituloOneroso>"
                        Print #NumFile, Space(10); "<montoTituloGratuito>" & FTalonMes.AdoRI2S.Recordset.Fields("Titulo_Gratuito") & "</montoTituloGratuito>"
                        FTalonMes.AdoRI2S.Recordset.MoveNext
                    Else
                        Print #NumFile, Space(10); "<contrato>0</contrato>"
                        Print #NumFile, Space(10); "<montoTituloOneroso>0.00</montoTituloOneroso>"
                        Print #NumFile, Space(10); "<montoTituloGratuito>0.00</montoTituloGratuito>"
                    End If
                 End If
        Print #NumFile, Space(7); "</detalleCompras>"
    .MoveNext
   Loop
        Print #NumFile, Space(5); "</compras>"
End If
End With
'' ventas
With FTalonMes.AdoVentas.Recordset
If .RecordCount > 0 Then
   .MoveFirst
   Print #NumFile, Space(5); "<ventas>"
   Do While Not .EOF
      Print #NumFile, Space(7); "<detalleVentas>"
        If .Fields("CI_RUC") = "9999999999999" Then
            Cadena2 = "07"
        Else
             Select Case .Fields("CTD")
             Case "R": Cadena2 = "04"
             Case "C": Cadena2 = "05"
             Case "P": Cadena2 = "06"
             End Select
        End If
        
        Print #NumFile, Space(10); "<tpIdCliente>" & Cadena2 & "</tpIdCliente>"
        Print #NumFile, Space(10); "<idCliente>" & .Fields("CI_RUC") & "</idCliente>"
        Print #NumFile, Space(10); "<tipoComprobante>" & Val(.Fields("TTD")) & "</tipoComprobante>"
        Cadena1 = UltimoDiaMes("01/" & Format(.Fields("Fecha"), "00") & "/" & Vanio)
        Print #NumFile, Space(10); "<fechaRegistro>" & Cadena1 & "</fechaRegistro>"
        Print #NumFile, Space(10); "<numeroComprobantes>" & .Fields("Numero") & "</numeroComprobantes>"
        Cadena1 = UltimoDiaMes("01/" & Format(.Fields("FechaE"), "00") & "/" & Vanio)
        Print #NumFile, Space(10); "<fechaEmision>" & Cadena1 & "</fechaEmision>"
        Print #NumFile, Space(10); "<baseImponible>" & .Fields("BImpotCero") & "</baseImponible>"
        Print #NumFile, Space(10); "<ivaPresuntivo>" & .Fields("ConvInt") & "</ivaPresuntivo>"
        Print #NumFile, Space(10); "<baseImpGrav>" & .Fields("Valor_Fact") & "</baseImpGrav>"
        Print #NumFile, Space(10); "<porcentajeIva>" & .Fields("CodPorc") & "</porcentajeIva>"
        If .Fields("ConvInt") = "N" Then
           'Print #NumFile, Space(10); "<montoIva>" & Format(.Fields("Valor_IVA"), "#0.00") & "</montoIva>"
           Print #NumFile, Space(10); "<montoIva>" & Format(.Fields("MontoIVA1") + .Fields("MontoIVA2"), "#0.00") & "</montoIva>"
        Else
           Print #NumFile, Space(10); "<montoIvaPresuntivo>" & Format(.Fields("MontoIVA1") + .Fields("MontoIVA2"), "#0.00") & "</montoIvaPresuntivo>"
        End If
        Print #NumFile, Space(10); "<baseImpIce>" & .Fields("MontoICE") & "</baseImpIce>"
        Print #NumFile, Space(10); "<porcentajeIce>" & MidStrg(.Fields("CPorcICE"), 1, 1) & "</porcentajeIce>"
        Print #NumFile, Space(10); "<montoIce>" & Format(.Fields("RetICE"), "#0.00") & "</montoIce>"
        Print #NumFile, Space(10); "<montoIvaBienes>" & Format(.Fields("MontoIVA1"), "#0.00") & "</montoIvaBienes>"
        Print #NumFile, Space(10); "<porRetBienes>" & .Fields("PorRetIVA1") & "</porRetBienes>"
        Print #NumFile, Space(10); "<valorRetBienes>" & .Fields("MontoRetIVA1") & "</valorRetBienes>"
        Print #NumFile, Space(10); "<montoIvaServicios>" & Format(.Fields("MontoIVA2"), "#0.00") & "</montoIvaServicios>"
        Print #NumFile, Space(10); "<porRetServicios>" & .Fields("PorRetIVA2") & "</porRetServicios>"
        Print #NumFile, Space(10); "<valorRetServicios>" & .Fields("MontoRetIVA2") & "</valorRetServicios>"
        Print #NumFile, Space(10); "<retPresuntiva>" & .Fields("ConvInt") & "</retPresuntiva>"
        ' retencion en la fuente del iva en ventas
        Codigo = .Fields("Codigo")
        sSQL = "SELECT TD,SUM(Valor_Fact)As Valor_Fact,CodPorc,Porc,Sum(Valor_Ret) As Valor_Ret,Serie,max(Secuencial)As Secuencial,Autorizacion,Max(FechaE),Max(Fecha) " _
             & "FROM Trans_Retenciones " _
             & "WHERE Item = '" & NumEmpresa & "' " _
             & "AND Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
             & "AND CodigoTR = 'RV' AND T <> 'A' " _
             & "AND Codigo = '" & Codigo & "' GROUP BY TD,Codporc,porc,serie,autorizacion "
        Select_Adodc FTalonMes.AdoRI2B, sSQL
        With FTalonMes.AdoRI2B.Recordset
            If .RecordCount > 0 Then
              ' MsgBox .RecordCount
               .MoveFirst
               No_RegV = 1
               Print #NumFile, Space(10); "<air>"
               Do While Not .EOF
                       If No_RegV > 2 Then MsgBox "Revise el Comprobante al Proveedor " & Codigo
                       Print #NumFile, Space(13); "<detalleAir>"
                       Print #NumFile, Space(15); "<codRetAir>" & .Fields("TD") & "</codRetAir>"
                       Print #NumFile, Space(15); "<baseImpAir>" & .Fields("Valor_Fact") & "</baseImpAir>"
                       Print #NumFile, Space(15); "<porcentajeAir>" & Val(.Fields("Porc")) * 100 & "</porcentajeAir>"
                       Print #NumFile, Space(15); "<valRetAir>" & .Fields("Valor_Ret") & "</valRetAir>"
                       Print #NumFile, Space(13); "</detalleAir>"
                       .MoveNext
                       No_RegV = No_RegV + 1
                Loop
                Print #NumFile, Space(10); "</air>"
            Else
                Print #NumFile, Space(10); "<air>"
                Print #NumFile, Space(10); "</air>"
            End If
        End With
        Print #NumFile, Space(7); "</detalleVentas>"
    .MoveNext
   Loop
        Print #NumFile, Space(5); "</ventas>"
Else
 Print #NumFile, Space(5); "<ventas/>"
End If
End With
' datos importaciones
With FTalonMes.AdoImpor.Recordset
If .RecordCount > 0 Then
   .MoveFirst
   Print #NumFile, Space(5); "<importaciones>"
   Do While Not .EOF
      Print #NumFile, Space(7); "<detalleImportaciones>"
        Print #NumFile, Space(10); "<codSustento>" & .Fields("SCT") & "</codSustento>"
        If .Fields("ConvInt") = "N" Then
           Print #NumFile, Space(10); "<importacionDe>2</importacionDe>"
        Else
           Print #NumFile, Space(10); "<importacionDe>1</importacionDe>"
        End If
        Print #NumFile, Space(10); "<fechaLiquidacion>" & .Fields("Fecha") & "</fechaLiquidacion>"
        Print #NumFile, Space(10); "<tipoComprobante>" & .Fields("TTD") & "</tipoComprobante>"
        Print #NumFile, Space(10); "<distAduanero>" & MidStrg(.Fields("Aduana"), 1, 3) & "</distAduanero>"
        Print #NumFile, Space(10); "<anio>" & MidStrg(.Fields("Aduana"), 4, 4) & "</anio>"
        Print #NumFile, Space(10); "<regimen>" & MidStrg(.Fields("Aduana"), 8, 2) & "</regimen>"
        Print #NumFile, Space(10); "<correlativo>" & MidStrg(.Fields("Aduana"), 10, 6) & "</correlativo>"
        Print #NumFile, Space(10); "<verificador>" & MidStrg(.Fields("Aduana"), 16, 1) & "</verificador>"
        Print #NumFile, Space(10); "<idFiscalProv>" & .Fields("CI_RUC") & "</idFiscalProv>"
        Print #NumFile, Space(10); "<valorCIF>" & .Fields("ValorCif") & "</valorCIF>"
        Print #NumFile, Space(10); "<razonSocialProv>" & .Fields("Cliente") & "</razonSocialProv>"
        Select Case .Fields("CTD")
        Case "R"
        Print #NumFile, Space(10); "<tipoSujeto>2</tipoSujeto>"
        Case Else
        Print #NumFile, Space(10); "<tipoSujeto>1</tipoSujeto>"
        End Select
        Print #NumFile, Space(10); "<baseImponible>" & .Fields("BImpotCero") & "</baseImponible>"
        Print #NumFile, Space(10); "<baseImpGrav>" & .Fields("Valor_Fact") & "</baseImpGrav>"
        Print #NumFile, Space(10); "<porcentajeIva>" & .Fields("CodPorc") & "</porcentajeIva>"
        Print #NumFile, Space(10); "<montoIva>" & .Fields("MontoIVA1") & "</montoIva>"
        Print #NumFile, Space(10); "<baseImpIce>" & .Fields("MontoICE") & "</baseImpIce>"
        Print #NumFile, Space(10); "<porcentajeIce>" & MidStrg(.Fields("CPorcICE"), 1, 1) & "</porcentajeIce>"
        Print #NumFile, Space(10); "<montoIce>" & Format(.Fields("RetICE"), "#0.00") & "</montoIce>"
        ' retencion en la fuente del iva en importaciones
        Codigo = .Fields("Comprobante")
                 If FTalonMes.AdoRI.Recordset.RecordCount > 0 Then
                    'AdoRI.Recordset.MoveFirst
                    FTalonMes.AdoRI.Recordset.Find ("Comprobante = '" & Codigo & "' ")
                    If Not FTalonMes.AdoRI.Recordset.EOF Then
                       Print #NumFile, Space(10); "<air>"
                       Print #NumFile, Space(13); "<detalleAir>"
                       Print #NumFile, Space(15); "<codRetAir>" & FTalonMes.AdoRI.Recordset.Fields("TRet") & "</codRetAir>"
                       Print #NumFile, Space(15); "<baseImpAir>" & FTalonMes.AdoRI.Recordset.Fields("Valor_Fact") & "</baseImpAir>"
                       Print #NumFile, Space(15); "<porcentajeAir>" & FTalonMes.AdoRI.Recordset.Fields("CodPorc") & "</porcentajeAir>"
                       Print #NumFile, Space(15); "<valRetAir>" & FTalonMes.AdoRI.Recordset.Fields("Valor_Ret") & "</valRetAir>"
                       Print #NumFile, Space(13); "</detalleAir>"
                       Print #NumFile, Space(10); "</air>"
                       ' AdoRI.Recordset.MoveLast
'                       AdoRI.Recordset.MoveNext
                    Else
                       Print #NumFile, Space(10); "<air/>"
                    End If
                 Else
                   Print #NumFile, Space(10); "<air/>"
                 End If
        Print #NumFile, Space(7); "</detalleImportaciones>"
    .MoveNext
   Loop
        Print #NumFile, Space(5); "</importaciones>"
Else
Print #NumFile, Space(5); "<importaciones/>"
End If
End With
' datos exportaciones
With FTalonMes.AdoExpor.Recordset
If .RecordCount > 0 Then
   .MoveFirst
   Print #NumFile, Space(5); "<exportaciones>"
   Do While Not .EOF
      Print #NumFile, Space(7); "<detalleExportaciones>"
        If .Fields("ConvInt") = "N" Then
           Print #NumFile, Space(10); "<exportacionDe>2</exportacionDe>"
        Else
           Print #NumFile, Space(10); "<exportacionDe>1</exportacionDe>"
        End If
        Print #NumFile, Space(10); "<tipoComprobante>" & .Fields("TTD") & "</tipoComprobante>"
        Print #NumFile, Space(10); "<distAduanero>" & MidStrg(.Fields("Aduana"), 1, 3) & "</distAduanero>"
        Print #NumFile, Space(10); "<anio>" & MidStrg(.Fields("Aduana"), 4, 4) & "</anio>"
        Print #NumFile, Space(10); "<regimen>" & MidStrg(.Fields("Aduana"), 8, 2) & "</regimen>"
        Print #NumFile, Space(10); "<correlativo>" & MidStrg(.Fields("Aduana"), 10, 6) & "</correlativo>"
        Print #NumFile, Space(10); "<verificador>" & MidStrg(.Fields("Aduana"), 16, 1) & "</verificador>"
        Print #NumFile, Space(10); "<fechaEmbarque>" & .Fields("FechaC") & "</fechaEmbarque>"
        Print #NumFile, Space(10); "<idFiscalCliente>" & .Fields("CI_RUC") & "</idFiscalCliente>"
        Select Case .Fields("CTD")
        Case "R"
        Print #NumFile, Space(10); "<tipoSujeto>2</tipoSujeto>"
        Case Else
        Print #NumFile, Space(10); "<tipoSujeto>1</tipoSujeto>"
        End Select
        Print #NumFile, Space(10); "<valorFOB>" & .Fields("ValorCif") & "</valorFOB>"
        Print #NumFile, Space(10); "<razonSocial>" & .Fields("Cliente") & "</razonSocial>"
        Print #NumFile, Space(10); "<devIva>" & .Fields("Dev") & "</devIva>"
        Print #NumFile, Space(10); "<facturaExportacion>1</facturaExportacion>"
        Print #NumFile, Space(10); "<valorFOBComprobante>" & .Fields("ValorCif") & "</valorFOBComprobante>"
        Print #NumFile, Space(10); "<establecimiento>" & MidStrg(.Fields("Serie"), 1, 3) & "</establecimiento>"
        Print #NumFile, Space(10); "<puntoEmision>" & MidStrg(.Fields("Serie"), 4, 3) & "</puntoEmision>"
        Print #NumFile, Space(10); "<secuencial>" & .Fields("Secuencial") & "</secuencial>"
        Print #NumFile, Space(10); "<fechaRegistro>" & .Fields("Fecha") & "</fechaRegistro>"
        Print #NumFile, Space(10); "<autorizacion>" & .Fields("autorizacion") & "</autorizacion>"
        Print #NumFile, Space(10); "<fechaEmision>" & .Fields("FechaE") & "</fechaEmision>"
        ' retencion en la fuente del iva en importaciones
        Codigo = .Fields("Comprobante")
                 If FTalonMes.AdoRI.Recordset.RecordCount > 0 Then
                    'AdoRI.Recordset.MoveFirst
                    FTalonMes.AdoRI.Recordset.Find ("Comprobante = '" & Codigo & "' ")
                    If Not FTalonMes.AdoRI.Recordset.EOF Then
                       Print #NumFile, Space(10); "<air>"
                       Print #NumFile, Space(13); "<detalleAir>"
                       Print #NumFile, Space(15); "<codRetAir>" & FTalonMes.AdoRE.Recordset.Fields("TRet") & "</codRetAir>"
                       Print #NumFile, Space(15); "<baseImpAir>" & FTalonMes.AdoRE.Recordset.Fields("Valor_Fact") & "</baseImpAir>"
                       Print #NumFile, Space(15); "<porcentajeAir>" & FTalonMes.AdoRE.Recordset.Fields("Porc") & "</porcentajeAir>"
                       Print #NumFile, Space(15); "<valRetAir>" & FTalonMes.AdoRE.Recordset.Fields("Valor_Ret") & "</valRetAir>"
                       Print #NumFile, Space(13); "</detalleAir>"
                       Print #NumFile, Space(10); "</air>"
'                       AdoRE.Recordset.MoveLast
'                       AdoRE.Recordset.MoveNext
                    Else
                       Print #NumFile, Space(10); "<air/>"
                    End If
                 Else
                    Print #NumFile, Space(10); "<air/>"
                 End If
        Print #NumFile, Space(7); "</detalleExportaciones>"
    .MoveNext
   Loop
        Print #NumFile, Space(5); "</exportaciones>"
Else
   Print #NumFile, Space(5); "<exportaciones/>"
End If
End With
'Print #NumFile, Space(5); "<recap/>"
' detalle recap
With FTalonMes.AdoTC.Recordset ' detalle recap
If .RecordCount > 0 Then
   .MoveFirst
   Print #NumFile, Space(5); "<recap>"
   Do While Not .EOF
      Print #NumFile, Space(7); "<detalleRecap>"
        Select Case .Fields("CTD")
        Case "R"
        Print #NumFile, Space(10); "<establecimientoRecap>10</establecimientoRecap>"
        Case Else
        Print #NumFile, Space(10); "<establecimientoRecap>11</establecimientoRecap>"
        End Select
        Print #NumFile, Space(10); "<identificacionRecap>" & .Fields("CI_RUC") & "</identificacionRecap>"
        Print #NumFile, Space(10); "<tipoComprobante>" & .Fields("TTD") & "</tipoComprobante>"
        Print #NumFile, Space(10); "<numeroRecap>" & .Fields("Aduana") & "</numeroRecap>"
        Print #NumFile, Space(10); "<fechaPago>" & .Fields("Fecha") & "</fechaPago>"
        Print #NumFile, Space(10); "<tarjetaCredito>" & .Fields("IdenCT") & "</tarjetaCredito>"
        Print #NumFile, Space(10); "<fechaEmisionRecap>" & .Fields("FechaE") & "</fechaEmisionRecap>"
        Print #NumFile, Space(10); "<consumoCero>" & .Fields("BImpotCero") & "</consumoCero>"
        Print #NumFile, Space(10); "<consumoGravado>" & .Fields("Valor_Fact") & "</consumoGravado>"
        Print #NumFile, Space(10); "<totalConsumo>" & .Fields("Valor_Fact") + .Fields("BImpotCero") & "</totalConsumo>"
        Print #NumFile, Space(10); "<montoIva>" & .Fields("MontoIVA1") + .Fields("MontoIVA2") & "</montoIva>"
        Print #NumFile, Space(10); "<comision>" & .Fields("Comision") & "</comision>"
        Print #NumFile, Space(10); "<numeroVouchers>" & .Fields("SerieEx") & "</numeroVouchers>"
        Print #NumFile, Space(10); "<montoIvaBienes>" & .Fields("MontoIVA1") & "</montoIvaBienes>"
        Print #NumFile, Space(10); "<porRetBienes>" & .Fields("PorRetIVA1") & "</porRetBienes>"
        Print #NumFile, Space(10); "<valorRetBienes>" & .Fields("MontoRetIVA1") & "</valorRetBienes>"
        Print #NumFile, Space(10); "<montoIvaServicios>" & .Fields("MontoIVA2") & "</montoIvaServicios>"
        Print #NumFile, Space(10); "<porRetServicios>" & .Fields("PorRetIVA2") & "</porRetServicios>"
        Print #NumFile, Space(10); "<valorRetServicios>" & .Fields("MontoRetIVA2") & "</valorRetServicios>"
    ' retencion en la fuente del iva en rendimientos financieros
        Codigo = .Fields("Comprobante")
                 If FTalonMes.AdoRT.Recordset.RecordCount > 0 Then
                    'AdoRI.Recordset.MoveFirst
                    FTalonMes.AdoRT.Recordset.Find ("Comprobante = '" & Codigo & "' ")
                    If Not FTalonMes.AdoRI.Recordset.EOF Then
                       Print #NumFile, Space(10); "<air>"
                       Print #NumFile, Space(13); "<detalleAir>"
                       Print #NumFile, Space(15); "<codRetAir>" & FTalonMes.AdoRT.Recordset.Fields("TD") & "</codRetAir>"
                       Print #NumFile, Space(15); "<baseImpAir>" & FTalonMes.AdoRT.Recordset.Fields("Valor_Fact") & "</baseImpAir>"
                       Print #NumFile, Space(15); "<porcentajeAir>" & FTalonMes.AdoRT.Recordset.Fields("CodPorc") & "</porcentajeAir>"
                       Print #NumFile, Space(15); "<valRetAir>" & FTalonMes.AdoRT.Recordset.Fields("Valor_Ret") & "</valRetAir>"
                       Print #NumFile, Space(13); "</detalleAir>"
                       Print #NumFile, Space(10); "</air>"
                       Print #NumFile, Space(15); "<establecimiento>" & MidStrg(FTalonMes.AdoRT.Recordset.Fields("Serie"), 1, 3) & "</establecimiento>"
                       Print #NumFile, Space(15); "<puntoEmision>" & MidStrg(FTalonMes.AdoRT.Recordset.Fields("Serie"), 4, 3) & "</puntoEmision>"
                       Print #NumFile, Space(10); "<secuencial>" & FTalonMes.AdoRT.Recordset.Fields("Secuencial") & "</secuencial>"
                       Print #NumFile, Space(10); "<fechaRegistro>" & FTalonMes.AdoRT.Recordset.Fields("Fecha") & "</fechaRegistro>"
                       Print #NumFile, Space(10); "<autorizacion>" & FTalonMes.AdoRT.Recordset.Fields("Autorizacion") & "</autorizacion>"
                       Print #NumFile, Space(10); "<fechaEmision>" & FTalonMes.AdoRT.Recordset.Fields("FechaE") & "</fechaEmision>"
'                       AdoRI.Recordset.MoveLast
                       AdoRT.Recordset.MoveNext
                    Else
                      Print #NumFile, Space(10); "<air/>"
                    End If
                 End If
        Print #NumFile, Space(7); "</detalleRecap>"
    .MoveNext
   Loop
        Print #NumFile, Space(5); "</recap>"
Else
Print #NumFile, Space(5); "<recap/>"
End If
End With

Print #NumFile, Space(5); "<fideicomisos/>"
' detalle dondos y fideicomisos
With FTalonMes.AdoAnulados.Recordset   ' comprobantes anulados
If .RecordCount > 0 Then
    .MoveFirst
    Print #NumFile, Space(5); "<anulados>"
     Do While Not .EOF
      Print #NumFile, Space(7); "<detalleAnulados>"
      Cadena = Val(.Fields("TD"))
      If Cadena > 99 Then Cadena = 7
      Print #NumFile, Space(10); "<tipoComprobante>" & Cadena & "</tipoComprobante>"
      Print #NumFile, Space(10); "<establecimiento>" & MidStrg(.Fields("Serie"), 1, 3) & "</establecimiento>"
      Print #NumFile, Space(10); "<puntoEmision>" & MidStrg(.Fields("Serie"), 4, 3) & "</puntoEmision>"
      Print #NumFile, Space(10); "<secuencialInicio>" & .Fields("Secuencial") & "</secuencialInicio>"
      Print #NumFile, Space(10); "<secuencialFin>" & .Fields("Secuencial") & "</secuencialFin>"
      Print #NumFile, Space(10); "<autorizacion>" & .Fields("Autorizacion") & "</autorizacion>"
      Print #NumFile, Space(10); "<fechaAnulacion>" & .Fields("FechaA") & "</fechaAnulacion>"
      Print #NumFile, Space(7); "</detalleAnulados>"
      .MoveNext
    Loop
    Print #NumFile, Space(5); "</anulados>"
Else
    Print #NumFile, Space(5); "<anulados/>"
End If
End With
With FTalonMes.AdoRRF.Recordset ' rendimientos financieros
If .RecordCount > 0 Then
   .MoveFirst
   Print #NumFile, Space(5); "<rendFinancieros>"
   Do While Not .EOF
      Print #NumFile, Space(7); "<detalleRendFinancieros>"
        Select Case .Fields("CTD")
        Case "R"
        Print #NumFile, Space(10); "<retenido>12</retenido>"
        Case Else
        Print #NumFile, Space(10); "<retenido>13</retenido>"
        End Select
        Print #NumFile, Space(10); "<idRetenido>" & .Fields("CI_RUC") & "</idRetenido>"
        Print #NumFile, Space(10); "<tpCompb>" & .Fields("TTD") & "</tpCompb>"
        Print #NumFile, Space(10); "<tipoCompR>40</tipoCompR>"
        ' retencion en la fuente de rendimientos financieros
        Codigo = .Fields("Comprobante")
                 If FTalonMes.AdoRRFF.Recordset.RecordCount > 0 Then
                    'AdoRI.Recordset.MoveFirst
                    FTalonMes.AdoRRFF.Recordset.Find ("Comprobante = '" & Codigo & "' ")
                    If Not FTalonMes.AdoRRFF.Recordset.EOF Then
                       Print #NumFile, Space(10); "<conRetT>" & FTalonMes.AdoRRFF.Recordset.Fields("TD") & "</conRetT>"
                       Print #NumFile, Space(10); "<baseImponibleRetT>" & FTalonMes.AdoRRFF.Recordset.Fields("Valor_Fact") & "</baseImponibleRetT>"
                       Print #NumFile, Space(10); "<codPorcRetT>" & FTalonMes.AdoRRFF.Recordset.Fields("CodPorc") & "</codPorcRetT>"
                       Print #NumFile, Space(10); "<montoRetT>" & FTalonMes.AdoRRFF.Recordset.Fields("Valor_Ret") & "</montoRetT>"
                       Print #NumFile, Space(10); "<serieRetT>" & FTalonMes.AdoRRFF.Recordset.Fields("Serie") & "</serieRetT>"
                       Print #NumFile, Space(10); "<secuencialRetT>" & FTalonMes.AdoRRFF.Recordset.Fields("Secuencial") & "</secuencialRetT>"
                       Print #NumFile, Space(10); "<autorizacionRetT>" & FTalonMes.AdoRRFF.Recordset.Fields("Autorizacion") & "</autorizacionRetT>"
                       Print #NumFile, Space(10); "<fechaEmisionRetT>" & FTalonMes.AdoRRFF.Recordset.Fields("FechaE") & "</fechaEmisionRetT>"
'                       AdoRI.Recordset.MoveLast
'                       AdoRI.Recordset.MoveNext
                    End If
                 End If
        Print #NumFile, Space(7); "</detalleRendFinancieros>"
    .MoveNext
   Loop
        Print #NumFile, Space(5); "</rendFinancieros>"
Else
Print #NumFile, Space(5); "<rendFinancieros/>"
End If
End With
Print #NumFile, "</iva>";
Close #NumFile
MiFrom.Caption = CaptionOld
RatonNormal
'MsgBox "Proceso Terminado"
End Sub

Public Sub ImprimirCR(ImpSoloReten As Boolean, _
                           Co As Comprobantes)
Dim AdoComp As ADODB.Recordset
Dim AdoRet As ADODB.Recordset
 Select Case Co.TP
   Case CompEgreso: Mensajes = "Imprimir Comprobante de Egreso No. "
   Case CompDiario: Mensajes = "Imprimir Comprobante de Diario No. "
 End Select
 Mensajes = Mensajes & Format(Co.Numero, "00000000") & " en:" & vbCrLf & Printer.DeviceName & "?"
 Titulo = "IMPRESION DE " & Co.TP
 Bandera = False
 SetPrinters.Show 1
 If PonImpresoraDefecto(SetNombrePRN) Then
    Escala_Centimetro 1, TipoTimes, 10
    Co.Fecha = FechaSistema
   'Listar el Comprobante
    sSQL = "SELECT C.*,A.Nombre_Completo,Cl.CI_RUC,Cl.Direccion,Cl.Email," _
         & "Cl.Telefono,Cl.Celular,Cl.FAX,Cl.Cliente,Cl.Codigo,Cl.Ciudad " _
         & "FROM Trans_Retenciones As C,Accesos As A,Clientes As Cl " _
         & "WHERE C.Numero = " & Co.Numero & " " _
         & "AND C.TP = '" & Co.TP & "' " _
         & "AND C.Item = '" & Co.Item & "' " _
         & "AND C.Periodo = '" & Periodo_Contable & "' " _
         & "AND C.CodigoU = A.Codigo " _
         & "AND C.Codigo = Cl.Codigo "
    Select_AdoDB AdoComp, sSQL
    If AdoComp.RecordCount > 0 Then Co.Fecha = AdoComp.Fields("Fecha")
   'Listar las Transacciones
    'Llenar Bancos
'    sSQL = "SELECT T.Cta,C.TC,C.Cuenta,Co.Fecha,Cl.Cliente,T.Cheq_Dep,T.Debe,T.Haber " _
'         & "FROM Transacciones As T,Comprobantes As Co,Catalogo_Cuentas As C,Clientes As Cl " _
'         & "WHERE T.TP = '" & Co.TP & "' " _
'         & "AND T.Numero = " & Co.Numero & " " _
'         & "AND T.Item = '" & Co.Item & "' " _
'         & "AND T.Periodo = '" & Periodo_Contable & "' " _
'         & "AND T.Numero = Co.Numero " _
'         & "AND T.TP = Co.TP " _
'         & "AND T.Cta = C.Codigo " _
'         & "AND T.Item = C.Item " _
'         & "AND T.Item = Co.Item " _
'         & "AND T.Periodo = C.Periodo " _
'         & "AND T.Periodo = Co.Periodo " _
'         & "AND C.TC = 'BA' " _
'         & "AND Co.Codigo_B = Cl.Codigo "
'    Select_AdoDB AdoBanc, sSQL
'   'Listar las Retenciones
    sSQL = "SELECT R.Cta,R.TD,TR.Tipo_Ret As Concepto,R.TP,R.Numero,R.Valor_Fact,R.MontoRetIVA1,R.MontoRetIVA2," _
         & "R.Porc,R.Valor_Ret,R.Secuencial,R.Retencion_No,R.TT,R.CodigoTR,R.MontoIVA1,R.MontoIVA2 " _
         & "FROM Trans_Retenciones As R, Tipo_Reten As TR " _
         & "WHERE R.TD = TR.Codigo " _
         & "AND R.Numero = " & Co.Numero & " " _
         & "AND R.Periodo = '" & Periodo_Contable & "' " _
         & "AND R.TP = '" & Co.TP & "' " _
         & "AND R.Item = '" & Co.Item & "' " _
         & "AND TR.Anio = '" & Format(Year(Co.Fecha), "0000") & "' " _
         & "UNION " _
         & "SELECT R.Cta,R.TD,TIV.Detalle As Concepto,R.TP,R.Numero,R.Valor_Fact,R.MontoRetIVA1,R.MontoRetIVA2," _
         & "R.Porc,R.Valor_Ret,R.Secuencial,R.Retencion_No,R.TT,R.CodigoTR,R.MontoIVA1,R.MontoIVA2 " _
         & "FROM Trans_Retenciones As R,Tipo_IVA As TIV " _
         & "WHERE R.TD = TIV.Codigo " _
         & "AND R.Numero = " & Co.Numero & " " _
         & "AND R.TP = '" & Co.TP & "' " _
         & "AND R.Item = '" & Co.Item & "' " _
         & "AND R.Periodo = '" & Periodo_Contable & "' " _
         & "AND TIV.Cod = 'C' " _
         & "AND TIV.Anio = '" & Format(Year(Co.Fecha), "0000") & "' " _
         & "ORDER BY R.Cta,R.Porc "
    Select_AdoDB AdoRet, sSQL
    If AdoComp.RecordCount > 0 Then ConceptoComp = AdoComp.Fields("Numero")
    'MsgBox AdoRet.RecordCount
    TipoComp = Co.TP
    Select Case Co.TP
      Case CompEgreso: ImprimirComp AdoComp, AdoRet, ImpSoloReten
      Case CompDiario: ImprimirComp AdoComp, AdoRet, ImpSoloReten
    End Select
    AdoComp.Close
    AdoRet.Close
 End If
End Sub
Public Sub ImprimirComp(DataComp As ADODB.Recordset, _
                        DataRets As ADODB.Recordset, _
                        ImpSoloRet As Boolean, _
                        Optional NuevaPagina As Boolean)
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
     Mifecha = .Fields("Fecha")
     If DataRets.RecordCount > 0 Then
        Mensajes = "Imprimir Retencion"
        Titulo = "Pregunta de Impresion"
        If BoxMensaje = vbYes Then ImprimirCompRet ImpSoloRet, DataComp, DataRets
     End If
     Printer.FontName = LetraAnterior
     If NuevaPagina Then Printer.NewPage Else Printer.EndDoc
 End If
End With
RatonNormal
'MsgBox "El Comprobante se ha impreso correctamente."
Exit Sub
Errorhandler:
    RatonNormal
    ErrorDeImpresion
    Exit Sub
End Sub

Public Sub ImprimirCompRet(SoloRet As Boolean, _
                           DataComp As ADODB.Recordset, _
                           DataRets As ADODB.Recordset)
Dim Copias As Boolean
'Establecemos Espacios y seteos de impresion
PorteLetra = Printer.FontSize
LetraAnterior = Printer.FontName
Printer.FontName = TipoTimes
Printer.FontBold = True
CRConLineas = ProcesarSeteos(CompRetencion)
Volver_Imp_Ret:
If SetD(23).PosX <> 0 And SoloRet = False Then Printer.NewPage
If CRConLineas Then
  'Formato de Retencion
   If SetD(24).PosX > 0 And SetD(24).PosY > 0 And SetD(25).PosX > 0 And SetD(25).PosY > 0 Then
      Dibujo = RutaSistema & "\FORMATOS\RETENCIO.GIF"
      PrinterPaint Dibujo, SetD(24).PosX, SetD(24).PosY, SetD(25).PosX, SetD(25).PosY
   End If
   If SetD(1).PosX > 0 And SetD(1).PosY > 0 Then
     'Formato de Logotipo
      PrinterPaint LogoTipo, 0.2, SetD(1).PosY - 0.1, 4, SetD(1).PosY + 2
      Printer.FontSize = 18
      PrinterTexto CentrarTexto(Empresa), SetD(1).PosY, Empresa
      Printer.FontSize = 12
      PrinterTexto SetD(1).PosX, SetD(1).PosY, "R.U.C." & RUC
      Printer.FontSize = 8
      Cadena = "Dirección: " & Direccion
      PrinterTexto SetD(1).PosX, SetD(1).PosY, Cadena
      Cadena = NombreCiudad & " - " & NombrePais & ".  Teléfono: " & Telefono1 & "/FAX: " & FAX
      PrinterTexto SetD(1).PosX, SetD(1).PosY, Cadena
      Printer.FontBold = False
      PosLinea = InicioY
   End If
End If
PosLinea = 0.2
Pagina = 1
'Iniciamos la impresion
With DataComp
 If .RecordCount > 0 Then
    .MoveFirst
     Printer.FontSize = SetD(3).Porte
     PrinterFields SetD(3).PosX, SetD(3).PosY, .Fields("Cliente")
     Printer.FontSize = SetD(9).Porte
     PrinterFields SetD(9).PosX, SetD(9).PosY, .Fields("CI_RUC")
     Printer.FontSize = SetD(4).Porte
     PrinterFields SetD(4).PosX, SetD(4).PosY, .Fields("Direccion")
     Printer.FontSize = SetD(6).Porte
     PrinterFields SetD(6).PosX, SetD(6).PosY, .Fields("Telefono")
     Printer.FontSize = SetD(10).Porte
     PrinterTexto SetD(10).PosX, SetD(10).PosY, FechaAnio(.Fields("Fecha"))
     Printer.FontSize = SetD(2).Porte
     PrinterFields SetD(2).PosX, SetD(2).PosY, .Fields("Fecha")
     Printer.FontSize = SetD(5).Porte
     PrinterFields SetD(5).PosX, SetD(5).PosY, .Fields("Ciudad")
     Printer.FontSize = SetD(1).Porte
     PrinterTexto SetD(1).PosX, SetD(1).PosY, Autorizacion
 End If
End With
With DataRets
 If .RecordCount > 0 Then
    .MoveFirst
    'Encabezado de la Retencion
     Printer.FontSize = SetD(1).Porte
     PrinterFields SetD(1).PosX, SetD(1).PosY, .Fields("Secuencial")
     Select Case .Fields("TT")
       Case "N": Cadena = "NOTA VENTA"
       Case "L": Cadena = "LIQUIDACION DE COMPRAS"
       Case "T": Cadena = "TICKET"
       Case Else: Cadena = "FACTURA"
     End Select
     Printer.FontSize = SetD(7).Porte
     PrinterTexto SetD(7).PosX, SetD(7).PosY, Cadena
     Printer.FontSize = SetD(8).Porte
     Select Case .Fields("CodigoTR")
          Case "RF": PrinterTexto SetD(8).PosX, SetD(8).PosY, PrintRet.LblNumFac.Caption  '.Fields("Secuencial")
          Case "RI", "IV": PrinterTexto SetD(8).PosX, SetD(8).PosY, .Fields("Secuencial")
        End Select
     Printer.FontSize = SetD(11).Porte
     'PrinterLineas SetD(11).PosX, SetD(11).PosY, ConceptoComp, SetD(12).PosX
     PrinterTexto SetD(11).PosX, SetD(11).PosY, PrintRet.TxtBanco
     PrinterTexto SetD(12).PosX, SetD(12).PosY, PrintRet.TxtCheq
     PosLinea = SetD(13).PosY: SumaSaldo = 0
     Printer.FontBold = False
     SumaSaldo = 0
     Do While Not .EOF
        Printer.FontSize = SetD(14).Porte
        Select Case .Fields("CodigoTR")
          Case "RF": PrinterTexto SetD(14).PosX, PosLinea, "R.F"
          Case "RI", "IV": PrinterTexto SetD(15).PosX, PosLinea, "R.IVA"
        End Select
        Printer.FontSize = SetD(16).Porte
        PrinterFields SetD(16).PosX, PosLinea, .Fields("TD")
        Printer.FontSize = SetD(17).Porte
        PrinterTexto SetD(17).PosX, PosLinea, MidStrg(.Fields("Concepto"), 1, 45) & "..."
        Cadena = .Fields("TP") & "-" & .Fields("Numero")
        Printer.FontSize = SetD(18).Porte
        PrinterTexto SetD(18).PosX, PosLinea, Cadena
        Printer.FontSize = SetD(19).Porte
        Select Case .Fields("CodigoTR")
          Case "RF": Total_RetCta = .Fields("Valor_Fact")
          Case "RI", "IV": Total_RetCta = .Fields("MontoIVA1") + .Fields("MontoIVA2")
        End Select
        PrinterVariables SetD(19).PosX, PosLinea, Total_RetCta
        
        Printer.FontSize = SetD(21).Porte
        If .Fields("CodigoTR") = "IV" Then
            Total_DetRet = .Fields("MontoRetIVA1") + .Fields("MontoRetIVA2")
        Else
            Total_DetRet = .Fields("Valor_Ret")
        End If
        PrinterVariables SetD(21).PosX, PosLinea, Total_DetRet
        Diferencia = 0
        If Total_RetCta <> 0 Then Diferencia = Total_DetRet / Total_RetCta
        ' MsgBox Diferencia
        Printer.FontSize = SetD(20).Porte
        'PrinterTexto SetD(20).PosX, PosLinea, Format(.Fields("Porc"), "##0%")
        PrinterVariables SetD(20).PosX, PosLinea, Format(Diferencia, "##0%")
        
        SumaSaldo = SumaSaldo + Total_DetRet
        'MsgBox SumaSaldo & vbCrLf & Total_DetRet
        PosLinea = PosLinea + 0.4
       .MoveNext
     Loop
     Printer.FontSize = SetD(22).Porte
     PrinterVariables SetD(22).PosX, SetD(22).PosY, SumaSaldo
End If
End With
Mensajes = "Imprimir Copia de Retencion"
Titulo = "Pregunta de Impresion"
If BoxMensaje = vbYes Then
   'Printer.NewPage
   GoTo Volver_Imp_Ret
End If
'If SetD(23).PosX = 0 And SoloRet = False Then Printer.NewPage
Printer.FontSize = PorteLetra
Printer.FontName = LetraAnterior
End Sub
Public Function ErrorST(Cadena1)
Select Case Cadena1
     Case "01C":
       MsgBox ("No debe hacer una Factura a un Proveedor con Cédula debe ser un Proveedor con RUC")
     Case "01O"
       MsgBox ("No debe hacer una Factura a un Proveedor con un Documento que no sea RUC")
     Case "01P"
       MsgBox ("No debe hacer una Factura a un Proveedor con Pasaporte debe ser un Proveedor con RUC")
     Case "02C"
       MsgBox ("No debe hacer una Nota de Venta a un Proveedor con Cédula debe ser un Proveedor con RUC")
     Case "02O"
       MsgBox ("No debe hacer una Nota de Venta a un Proveedor con Cédula debe ser un Proveedor con RUC")
     Case "02P"
       MsgBox ("No debe hacer una Nota de Venta a un Proveedor con Cédula debe ser un Proveedor con RUC")
     Case "03R"
       MsgBox ("No debe hacer una Liquidacion a un Proveedor con RUC")
     Case "03O"
       MsgBox ("No debe hacer una Factura a un Proveedor con Otro Tipo de Documento")
     Case "08C"
       MsgBox ("No debe hacer una Boleta a un Proveedor con Cédula")
     Case "08P"
       MsgBox ("No debe hacer una Boleta a un Proveedor con Pasaporte")
     Case "08O":
       MsgBox ("No debe hacer una Boleta a un Proveedor con Otro tipo de Documento")
     Case "09C"
       MsgBox ("No debe hacer un Tiquete o Vale a un Proveedor con Cédula")
     Case "09P"
       MsgBox ("No debe hacer un Tiquete o Vale a un Proveedor con Pasaporte")
     Case "09O"
       MsgBox ("No debe hacer un Tiquete o Vale a un Proveedor con Otro tipo de Documento")
     Case "10C"
       MsgBox ("No debe hacer un Comprobante de Venta Art. 10 a un Proveedor con Cédula")
     Case "10P"
       MsgBox ("No debe hacer un Comprobante de Venta Art. 10 con Pasaporte")
     Case "10O":
       MsgBox ("No debe hacer un Comprobante de Venta Art. 10 con Otro tipo de Documento")
     Case "11C"
       MsgBox ("No debe hacer un Pasaje a un Proveedor con Cédula")
     Case "11P"
       MsgBox ("No debe hacer un Pasaje a un Proveedor con Pasaporte")
     Case "11O":
       MsgBox ("No debe hacer un Pasaje a un Proveedor con Otro tipo de Documento")
     Case "12C"
       MsgBox ("No debe hacer un Documento emitido por Inst. Financieros a un Proveedor con Cédula")
     Case "12P"
       MsgBox ("No debe hacer un Documento emitido por Inst. Financieros a un Proveedor con Pasaporte")
     Case "12O":
       MsgBox ("No debe hacer un Documento emitido por Inst. Financieros a un Proveedor con Otro tipo de Documento")
     Case "13C"
       MsgBox ("No debe hacer un Documento emitido por Inst. Aseguradoras a un Proveedor con Cédula")
     Case "13P"
       MsgBox ("No debe hacer un Documento emitido por Inst. Aseguradoras a un Proveedor con Pasaporte")
     Case "13O":
       MsgBox ("No debe hacer un Documento emitido por Inst. Aseguradoras a un Proveedor con Otro tipo de Documento")
  End Select
End Function
Public Sub Imprimir_Trans_Reten(Datas As Adodc, _
                                SizeLetra As Integer)
On Error GoTo Errorhandler
Mensajes = "Seguro de Imprimir en:" & vbCrLf & Printer.DeviceName & "?"
Titulo = "IMPRESION"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
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
      Cuenta = .Fields("Cuenta")
      Total_Ret = .Fields("Porc")
      Cadena = Cuenta & ": " & Format(Total_Ret, "00%")
      Total = 0
      Printer.FontBold = False
      Printer.FontSize = 12
      PrinterTexto Ancho(0), PosLinea, Cadena
      PosLinea = PosLinea + 0.6
      Printer.FontSize = 8
      Imprimir_Linea_H PosLinea, Ancho(0), Ancho(CantCampos)
      PosLinea = PosLinea + 0.1
      PrinterTexto Ancho(1), PosLinea, "TP"
      PrinterTexto Ancho(2), PosLinea, "Numero"
      PrinterTexto Ancho(3), PosLinea, "TD"
      PrinterTexto Ancho(4), PosLinea, "Codigo"
      PrinterTexto Ancho(5), PosLinea, "Cli/Prove"
      PrinterTexto Ancho(6), PosLinea, "RUC_CI"
      PrinterTexto Ancho(7), PosLinea, "BaseImponible"
      PrinterTexto Ancho(8), PosLinea, "Porc"
      PrinterTexto Ancho(9), PosLinea, "Valor_Ret"
      PrinterTexto Ancho(10), PosLinea, "I.V.A."
      PrinterTexto Ancho(11), PosLinea, "Valor_Re"
      PrinterTexto Ancho(12), PosLinea, "Porc"
      PosLinea = PosLinea + 0.5
      Imprimir_Linea_H PosLinea, Ancho(0), Ancho(CantCampos)
      PosLinea = PosLinea + 0.1
      Printer.FontBold = False
      Do While Not .EOF
         If Cuenta <> .Fields("Cuenta") Or Total_Ret <> .Fields("Porc") Then
            Printer.FontBold = False
            Imprimir_Linea_H PosLinea, Ancho(0), Ancho(CantCampos)
            PosLinea = PosLinea + 0.1
            PrinterTexto Ancho(9), PosLinea, "T O T A L"
            PrinterVariables Ancho(10), PosLinea, Total
            Cuenta = .Fields("Cuenta")
            Total_Ret = .Fields("Porc")
            PosLinea = PosLinea + 0.5
            Cadena = Cuenta & ": " & Format(Total_Ret, "00%")
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
         PrinterFields Ancho(1), PosLinea, .Fields("T")
         PrinterFields Ancho(2), PosLinea, .Fields("Fecha")
         PrinterFields Ancho(3), PosLinea, .Fields("Autorizacion")
         PrinterFields Ancho(4), PosLinea, .Fields("Secuencial")
         PrinterFields Ancho(5), PosLinea, .Fields("Cliente")
         PrinterFields Ancho(6), PosLinea, .Fields("CI_RUC")
         PrinterFields Ancho(7), PosLinea, .Fields("TP")
         PrinterFields Ancho(8), PosLinea, .Fields("Numero")
         PrinterFields Ancho(9), PosLinea, .Fields("Valor_Fact")
         PrinterFields Ancho(10), PosLinea, .Fields("IVA")
         PrinterFields Ancho(11), PosLinea, .Fields("Valor_Ret")
         PrinterFields Ancho(12), PosLinea, .Fields("Porc")
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
         Total = Total + .Fields("Valor_Ret")
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
