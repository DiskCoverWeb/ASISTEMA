Attribute VB_Name = "SubCajaCredito"
Option Explicit

Public Sub GenerarTablaPrestamo(BoxMiFecha As String, _
                                DtaTabla As Adodc, _
                                DBG_Tabla As DataGrid, _
                                TxtInt As TextBox, _
                                TxtMeses As TextBox, _
                                TxtMonto As TextBox, _
                                Meses_Dias As Boolean, _
                                TipoPrest As String, _
                                Optional SinComis As Boolean)
  Interes = Redondear(CCur(TxtInt.Text) / 100, 4)
  Numero = CInt(TxtMeses.Text)
  Total = Redondear(CCur(TxtMonto.Text), 2)
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
    If Meses_Dias Then   'Si_No = True Dias else Meses
       For I = 0 To 1
           SetAddNew DtaTabla
          .fields("T_No") = I
          .fields("Dia") = DiasLetras(Weekday(Mifecha))
           NoDias = Numero
           If I = 0 Then
              Valor_ME = Redondear(((Total * Interes) / 360) * (NoDias + 3), 2)
              Total_ME = Redondear(Total, 2)
              Valor = Redondear(Total + Valor_ME, 2)
             .fields("Fecha") = Mifecha
             .fields("Capital") = Total_ME
             .fields("Interes") = Valor_ME
             .fields("Pagos") = 0
             .fields("Saldo") = Redondear(Total + Valor_ME, 2)
              Mifecha = CLongFecha(CFechaLong(Mifecha) + Numero)
           Else
             .fields("Fecha") = Mifecha
             .fields("Capital") = 0
             .fields("Interes") = 0
             .fields("Pagos") = Total
             .fields("Saldo") = 0
           End If
          .fields("TP") = TipoPrest
          .fields("CodigoU") = CodigoUsuario
           SetUpdate DtaTabla
       Next I
    Else
       Total = Redondear(Total + (Total * ((Numero / 12) * Interes)), 2)
       Tasa = 0
       Do
         Tasa = Redondear(Tasa + 0.0001, 4)
         Cuota = Redondear(((Saldo * Tasa) / 12) / (1 - (1 + (Tasa / 12)) ^ -Numero), 2)
       Loop Until (Cuota * Numero) >= Total
       Contador = 1: Total = Saldo
       Valor = Redondear(((12 * Total) + (Total * Interes * Numero)) / (12 * Numero), 2)
       Valor_ME = 0: Total_ME = 0: Comision = 0
       For I = 0 To Numero
           SetAddNew DtaTabla
          .fields("T_No") = I
          .fields("Dia") = DiasLetras(Weekday(Mifecha))
           If I = 0 Then
             .fields("Fecha") = Mifecha
             .fields("Capital") = 0
             .fields("Interes") = 0
             .fields("Comision") = 0
             .fields("Pagos") = 0
              Mifecha = SiguienteMes(Mifecha)
              'MiFecha = CLongFecha(CFechaLong(MiFecha) + 30)
           Else
             .fields("Fecha") = Mifecha
             .fields("Capital") = Redondear(Total_ME, 2)
             .fields("Interes") = Redondear(Valor_ME, 2)
             .fields("Comision") = Redondear(Comision, 2)     ' Comision
             .fields("Pagos") = Redondear(Valor, 2)
              Mifecha = SiguienteMes(Mifecha)
              'MiFecha = CLongFecha(CFechaLong(MiFecha) + 30)
           End If
          .fields("CodigoU") = CodigoUsuario
          .fields("TP") = TipoPrest
          .fields("Saldo") = Total
           SetUpdate DtaTabla
          'Comision del 1%
           If SinComis = False Then Comision = Redondear(Total * 0.01, 2)
           'If SinComis = False Then Comision = Redondear(Total * 0.012, 2)
          'Interes Inicial
           Valor_ME = Redondear(Total * (Tasa / 12), 2)
          'Amortizacion o Capital
           Total_ME = Redondear(Valor - Valor_ME, 2)
          'Saldo Pendiente
           Total = Redondear(Total - Total_ME, 2)
          'Interes Final
           Valor_ME = Redondear(Valor - Total_ME - Comision, 2)
           Contador = Contador + 1
       Next I
    End If
  End With
  sSQL = "SELECT * " _
       & "FROM Asiento_P " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "ORDER BY T_No "
  Select_Adodc_Grid DBG_Tabla, DtaTabla, sSQL
  If Meses_Dias = False Then
     With DtaTabla.Recordset
      If .RecordCount > 0 Then
         .MoveLast
         'Comision del 1%
          Valor = Redondear(.fields("Interes"), 2)
          Total = Redondear(.fields("Capital"), 2)
          Abono = Redondear(.fields("Pagos"), 2)
          Saldo = Redondear(.fields("Saldo"), 2)
         'MsgBox Total
          If SinComis = False Then Comision = Redondear(Total * 0.01, 2)
         .fields("Interes") = Redondear(Abono - Total - Saldo - Comision, 2)
         .fields("Comision") = Redondear(Comision, 2)
         .fields("Capital") = Redondear(Total + Saldo, 2)
         .fields("Saldo") = 0
         .Update
      End If
     End With
  End If
  DBG_Tabla.Visible = True
  MsgBox "Tabla de Amortizacion Generada"
End Sub

Public Sub Generar_Tabla_Prestamo_Sobre_Saldos(BoxMiFecha As String, _
                                               DtaTabla As Adodc, _
                                               DBG_Tabla As DataGrid, _
                                               TxtInt As TextBox, _
                                               TxtMeses As TextBox, _
                                               TxtMonto As TextBox, _
                                               Meses_Dias As Boolean, _
                                               TipoPrest As String, _
                                               Optional SinComis As Boolean)
  Dim Ctas_Prest As Cuentas_Prestamos
  Dim Seguro As Single
  Dim Seguro1 As Single
  If Edad_Persona <= 65 Then
     Seguro = Leer_Campo_Empresa("Seguro") / 100000
  Else
     Seguro = Leer_Campo_Empresa("Seguro2") / 10000
  End If
  Seguro1 = Leer_Campo_Empresa("Seguro2") / 10000
  Ctas_Prest = Cuentas_del_Prestamo(TipoPrest)
  Interes = Redondear(CCur(TxtInt) / 100, 4)   ' Interes
  Numero = CInt(TxtMeses)                  ' No. Meses
  Total = Redondear(CCur(Val(TxtMonto)), 2)         ' Capital del Prestamo
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
    If Meses_Dias Then   'Si_No = True Dias else Meses
       For I = 0 To 1
           SetAddNew DtaTabla
          .fields("T_No") = I
          .fields("Dia") = DiasLetras(Weekday(Mifecha))
           NoDias = Numero
           If I = 0 Then
              Valor_ME = Redondear(((Total * Interes) / 360) * (NoDias + 3), 2)
              Total_ME = Redondear(Total, 2)
              Valor = Redondear(Total + Valor_ME, 2)
             .fields("Fecha") = Mifecha
             .fields("Capital") = Total_ME
             .fields("Interes") = Valor_ME
             .fields("Pagos") = 0
             .fields("Saldo") = Redondear(Total + Valor_ME, 2)
              Mifecha = CLongFecha(CFechaLong(Mifecha) + Numero)
              
           Else
             .fields("Fecha") = Mifecha
             .fields("Capital") = 0
             .fields("Interes") = 0
             .fields("Pagos") = Total
             .fields("Saldo") = 0
           End If
           Select Case CFechaLong(Mifecha) - CFechaLong(BoxMiFecha)
             Case 1 To 30: .fields("Cta") = Ctas_Prest.Cta_P_1_30
             Case 31 To 90: .fields("Cta") = Ctas_Prest.Cta_P_31_90
             Case Else: .fields("Cta") = Ctas_Prest.Cta_P_Mas_360
           End Select
          .fields("TP") = TipoPrest
          .fields("CodigoU") = CodigoUsuario
           SetUpdate DtaTabla
       Next I
    Else
       Interes = Interes / 12
       Valor = Redondear(Pmt(Interes, Numero, -Total), 2)
       Contador = 1: Saldo = Total: Valor_ME = 0: Total_ME = 0: Comision = 0
       Total_Saldos = Total
       For I = 0 To Numero
           SetAddNew DtaTabla
          .fields("T_No") = I
          .fields("Dia") = DiasLetras(Weekday(Mifecha))
           If I = 0 Then
             .fields("Fecha") = Mifecha
             .fields("Capital") = 0
             .fields("Interes") = 0
             .fields("Comision") = 0
             .fields("Pagos") = 0
              Mifecha = SiguienteMes(Mifecha)
             'MiFecha = CLongFecha(CFechaLong(MiFecha) + 30)
           Else
              Valor_ME = Redondear(Saldo * Interes, 2)    'Interes mensual
             'MsgBox Valor_ME
             'Comision o seguro de desgravamen
              If Seguro1 > 0 Then
                 Comision = Redondear(Total * Seguro, 2)
              Else
                 Comision = Redondear(Saldo * Seguro, 2)
              End If
             'MsgBox Comision & vbCrLf & Saldo & vbCrLf & Seguro
              'If SinComis = False Then Comision = 0
              Total_ME = Valor - Valor_ME   ' Capital Pagado + Comision
             .fields("Fecha") = Mifecha
             .fields("Pagos") = Redondear(Valor + Comision, 2)
             .fields("Interes") = Redondear(Valor_ME, 2)
             .fields("Capital") = Redondear(Total_ME, 2)
             .fields("Comision") = Redondear(Comision, 2)     ' Comision
              Mifecha = SiguienteMes(Mifecha)
              Saldo = Redondear(Saldo - Total_ME, 2)
              Total_Saldos = Redondear(Total_Saldos, 2)  '- Comision
           End If
           Select Case CFechaLong(Mifecha) - CFechaLong(BoxMiFecha)
             Case 1 To 30: .fields("Cta") = Ctas_Prest.Cta_P_1_30
             Case 31 To 90: .fields("Cta") = Ctas_Prest.Cta_P_31_90
             Case 91 To 180: .fields("Cta") = Ctas_Prest.Cta_P_91_180
             Case 181 To 360: .fields("Cta") = Ctas_Prest.Cta_P_181_360
             Case Else: .fields("Cta") = Ctas_Prest.Cta_P_Mas_360
           End Select
          .fields("CodigoU") = CodigoUsuario
          .fields("TP") = TipoPrest
          .fields("Saldo") = Saldo
           SetUpdate DtaTabla
          'Comision del 1%
           If SinComis = False Then Comision = Redondear(Saldo * 0.01, 2)
           Contador = Contador + 1
       Next I
    End If
  End With
  sSQL = "SELECT * " _
       & "FROM Asiento_P " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "ORDER BY T_No "
  Select_Adodc_Grid DBG_Tabla, DtaTabla, sSQL
  If Meses_Dias = False Then
     With DtaTabla.Recordset
      If .RecordCount > 0 Then
         .MoveLast
         'Comision del 1%
          Valor = Redondear(.fields("Interes"), 2)
          Total = Redondear(.fields("Capital"), 2)
          Abono = Redondear(.fields("Pagos"), 2)
          Saldo = Redondear(.fields("Saldo"), 2)
         'MsgBox Total
          If SinComis Then Comision = Redondear(Total * 0.01, 2)
          Comision = Redondear(Total * Seguro, 2)
         .fields("Interes") = Redondear(Abono - Total - Saldo - Comision, 2)
         .fields("Comision") = Redondear(Comision, 2)
         .fields("Capital") = Redondear(Total + Saldo, 2)
         .fields("Saldo") = 0
         .Update
      End If
     End With
  End If
  DBG_Tabla.Visible = True
 'MsgBox "Tabla de Amortizacion Generada"
End Sub

Public Sub ImprimirFlujoDePrestamos(DtaC As Adodc, _
                                    DtaD As Adodc)
On Error GoTo Errorhandler
Dim SizeLetra As Integer
RatonReloj
SizeLetra = 9
InicioX = 0.5: InicioY = 0
SumaDebe = 0: SumaHaber = 0
SaldoDebe = 0: SaldoHaber = 0
Debe = 0: Haber = 0
DataAnchoCampos InicioX, DtaD, SizeLetra, TipoTimes, 1
Ancho(0) = 0.5  ' Fecha
Ancho(1) = 2.3  ' TP
Ancho(2) = 3.1  ' Cuenta_No
Ancho(3) = 5    ' Nombre
Ancho(4) = 12.5 ' Credito_No
Ancho(5) = 14.5 ' Cuota_No
Ancho(6) = 16.5 ' Capital
Ancho(7) = 19   ' Fin
Pagina = 1
'Iniciamos la impresion
Printer.FontBold = False
With DtaC.Recordset
     .MoveFirst
      Encabezado 0.5, 19
      Printer.FontUnderline = True
      Printer.FontBold = True
      Printer.FontSize = 14
      PrinterTexto Ancho(0), PosLinea, "CREDITOS OTORGADOS:"
      PosLinea = PosLinea + 0.6
      Printer.FontSize = SizeLetra
      PrinterTexto Ancho(0), PosLinea, "Fecha"
      PrinterTexto Ancho(1), PosLinea, "TP"
      PrinterTexto Ancho(2), PosLinea, "Cuenta No"
      PrinterTexto Ancho(3), PosLinea, "Nombre Cliente"
      PrinterTexto Ancho(4), PosLinea, "Credito No"
      PrinterTexto Ancho(5), PosLinea, "Capital"
      Printer.FontUnderline = False
      Printer.FontBold = False
      PosLinea = PosLinea + 0.4
      TipoProc = .fields("TP")
      PrinterFields Ancho(0), PosLinea, .fields("Fecha")
      PrinterFields Ancho(1), PosLinea, .fields("TP")
      Do While Not .EOF
         If TipoProc <> .fields("TP") Then
            Printer.Line (Ancho(0), PosLinea)-(Ancho(7), PosLinea), QBColor(0)
            PosLinea = PosLinea + 0.1
            PrinterVariables Ancho(4), PosLinea, "TOTALES"
            PrinterVariables Ancho(5), PosLinea, Haber
            Debe = 0: Haber = 0
            PosLinea = PosLinea + 0.4
            TipoProc = .fields("TP")
            Printer.FontSize = SizeLetra
            PrinterFields Ancho(0), PosLinea, .fields("Fecha")
            PrinterFields Ancho(1), PosLinea, .fields("TP")
         End If
         Printer.FontSize = SizeLetra
         PrinterFields Ancho(2), PosLinea, .fields("Cuenta_No")
         Printer.FontSize = 8
         PrinterFields Ancho(3), PosLinea, .fields("Nombre_Cliente")
         Printer.FontSize = SizeLetra
         PrinterFields Ancho(4), PosLinea, .fields("Credito_No")
         PrinterFields Ancho(5), PosLinea, .fields("Capital")
         Haber = Haber + .fields("Capital")
         SumaHaber = SumaHaber + .fields("Capital")
         PosLinea = PosLinea + 0.4
         If PosLinea > LimiteAlto Then
            Printer.NewPage
            Encabezado 0.5, 19
            Printer.FontUnderline = True
            Printer.FontBold = True
            PrinterTexto Ancho(0), PosLinea, "Fecha"
            PrinterTexto Ancho(1), PosLinea, "TP"
            PrinterTexto Ancho(2), PosLinea, "Cuenta No"
            PrinterTexto Ancho(3), PosLinea, "Nombre Cliente"
            PrinterTexto Ancho(4), PosLinea, "Credito No"
            PrinterTexto Ancho(5), PosLinea, "Capital"
            Printer.FontUnderline = False
            Printer.FontBold = False
            PosLinea = PosLinea + 0.4
            Printer.FontSize = SizeLetra
            PrinterFields Ancho(0), PosLinea, .fields("Fecha")
         End If
        .MoveNext
      Loop
      Printer.Line (Ancho(0), PosLinea)-(Ancho(7), PosLinea), QBColor(0)
      PosLinea = PosLinea + 0.1
      PrinterVariables Ancho(4), PosLinea, "TOTALES"
      PrinterVariables Ancho(5), PosLinea, Haber
      PosLinea = PosLinea + 0.4
End With
Debe = 0: Haber = 0
PosLinea = PosLinea + 0.1
With DtaD.Recordset
     .MoveFirst
      If PosLinea > LimiteAlto Then
         Printer.NewPage
         Encabezado 0.5, 19
      End If
      Printer.FontUnderline = True
      Printer.FontBold = True
      Printer.FontSize = 14
      PrinterTexto Ancho(0), PosLinea, "ABONOS DEL DIA:"
      PosLinea = PosLinea + 0.6
      Printer.FontSize = SizeLetra
      PrinterTexto Ancho(0), PosLinea, "Fecha"
      PrinterTexto Ancho(1), PosLinea, "TP"
      PrinterTexto Ancho(2), PosLinea, "Cuenta No"
      PrinterTexto Ancho(3), PosLinea, "Nombre Cliente"
      PrinterTexto Ancho(4), PosLinea, "Credito No"
      PrinterTexto Ancho(5), PosLinea, "Cuota No"
      PrinterTexto Ancho(6), PosLinea, "Capital"
      Printer.FontUnderline = False
      Printer.FontBold = False
      PosLinea = PosLinea + 0.4
      
      Printer.FontSize = SizeLetra
      TipoProc = .fields("TP")
      PrinterFields Ancho(0), PosLinea, .fields("Fecha_C")
      PrinterFields Ancho(1), PosLinea, .fields("TP")
      Do While Not .EOF
         If TipoProc <> .fields("TP") Then
            Printer.Line (Ancho(0), PosLinea)-(Ancho(7), PosLinea), QBColor(0)
            PosLinea = PosLinea + 0.1
            PrinterVariables Ancho(4), PosLinea, "TOTALES"
            PrinterVariables Ancho(6), PosLinea, Debe
            Debe = 0: Haber = 0
            PosLinea = PosLinea + 0.4
            TipoProc = .fields("TP")
            Printer.FontSize = SizeLetra
            PrinterFields Ancho(0), PosLinea, .fields("Fecha_C")
            PrinterFields Ancho(1), PosLinea, .fields("TP")
         End If
         Printer.FontSize = SizeLetra
         PrinterFields Ancho(2), PosLinea, .fields("Cuenta_No")
         Printer.FontSize = 8
         PrinterFields Ancho(3), PosLinea, .fields("Nombre_Cliente")
         Printer.FontSize = SizeLetra
         PrinterFields Ancho(4), PosLinea, .fields("Credito_No")
         PrinterFields Ancho(5), PosLinea, .fields("Cuota_No")
         PrinterFields Ancho(6), PosLinea, .fields("Capital")
         Debe = Debe + .fields("Capital")
         SumaDebe = SumaDebe + .fields("Capital")
         PosLinea = PosLinea + 0.4
         If PosLinea > LimiteAlto Then
            Printer.NewPage
            Encabezado 0.5, 19
            Printer.FontSize = SizeLetra
            Printer.FontUnderline = True
            Printer.FontBold = True
            PrinterTexto Ancho(0), PosLinea, "Fecha"
            PrinterTexto Ancho(1), PosLinea, "TP"
            PrinterTexto Ancho(2), PosLinea, "Cuenta No"
            PrinterTexto Ancho(3), PosLinea, "Nombre Cliente"
            PrinterTexto Ancho(4), PosLinea, "Credito No"
            PrinterTexto Ancho(5), PosLinea, "Cuota No."
            PrinterTexto Ancho(6), PosLinea, "Capital"
            Printer.FontUnderline = False
            Printer.FontBold = False
            PosLinea = PosLinea + 0.4
            PrinterFields Ancho(0), PosLinea, .fields("Fecha_C")
         End If
        .MoveNext
      Loop
      Printer.Line (Ancho(0), PosLinea)-(Ancho(7), PosLinea), QBColor(0)
      PosLinea = PosLinea + 0.1
      PrinterVariables Ancho(4), PosLinea, "TOTALES"
      PrinterVariables Ancho(6), PosLinea, Debe
      PosLinea = PosLinea + 0.4
End With
Printer.FontSize = 11
'''PrinterTexto Ancho(2), PosLinea, "SALDO ANTERIOR: "
'''PrinterVariables Ancho(3) + 2, PosLinea, Saldo_Anterior
'''PosLinea = PosLinea + 0.5
'''PrinterTexto Ancho(2), PosLinea, "TOTAL CREDITOS"
'''PrinterVariables Ancho(3) + 2, PosLinea, SumaHaber
'''PosLinea = PosLinea + 0.5
'''PrinterTexto Ancho(2), PosLinea, "TOTAL DEBITOS"
'''PrinterVariables Ancho(3) + 2, PosLinea, SumaDebe
'''PosLinea = PosLinea + 0.5
'''PrinterTexto Ancho(2), PosLinea, "SALDO DE ACTUAL"
'''PrinterVariables Ancho(3) + 2, PosLinea, Saldo_Anterior - SumaHaber + SumaDebe
RatonNormal
MensajeEncabData = ""
Printer.EndDoc
Exit Sub
Errorhandler:
    RatonNormal
    ErrorDeImpresion
    Exit Sub
End Sub

Public Sub ImprimirFlujoCajaCoop(Datas As Adodc, _
                                 FinDoc As Boolean, _
                                 FormaImp As Byte, _
                                 SizeLetra As Integer, _
                                 EsGrupo As Boolean, _
                                 TipoCaja As Boolean, _
                                 Saldo_Anterior As Currency, _
                                 Resumido As Boolean, _
                                 Optional EsFlujoCaja As Boolean)
On Error GoTo Errorhandler
Mensajes = "Seguro de Imprimir en:" & vbCrLf & Printer.DeviceName & "?"
Titulo = "IMPRESION"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
RatonReloj
InicioX = 0.5: InicioY = 0
SumaDebe = 0: SumaHaber = 0: SaldoDebe = 0: SaldoHaber = 0
Debe = 0: Haber = 0: Debe_ME = 0: Haber_ME = 0
'Escala_Centimetro   FormaImp, TipoTimes, SizeLetra
DataAnchoCampos InicioX, Datas, SizeLetra, TipoTimes, FormaImp
ReDim Ancho(9) As Single
Ancho(0) = 0.5   ' ME
Ancho(1) = 1.1   ' Fecha
Ancho(2) = 3     ' TP
Ancho(3) = 3.8   ' Papeleta
Ancho(4) = 5.7   ' Cuenta
Ancho(5) = 7.5   ' Debitos
Ancho(6) = 10    ' Creditos
Ancho(7) = 17.5  ' CodigoU
If EsFlujoCaja Then
   Ancho(7) = 12.5
   Ancho(8) = 17.5
   Ancho(9) = 19
Else
   Ancho(8) = 19
End If
LimiteAlto = LimiteAlto - 1
Pagina = 1
Debitos = 0: Creditos = 0
'Iniciamos la impresion
Printer.FontBold = False
If Resumido = False Then
EncabezadoData Datas
Printer.FontName = TipoArialNarrow
With Datas.Recordset
 If .RecordCount > 0 Then
     .MoveFirst
      Printer.FontSize = SizeLetra
      Moneda_US = .fields("ME")
      Mifecha = .fields("Fecha")
      TipoProc = .fields("TP")
      Codigo = .fields("CodigoU")
      PrinterFields Ancho(0), PosLinea, .fields("ME"), False
      PrinterFields Ancho(1), PosLinea, .fields("Fecha"), False
      If EsFlujoCaja Then
         PrinterFields Ancho(7), PosLinea, .fields("Detalle"), False
         PrinterFields Ancho(8), PosLinea, .fields("CodigoU"), False
      Else
         PrinterFields Ancho(7), PosLinea, .fields("CodigoU"), False
      End If
      Do While Not .EOF
         If EsGrupo Then
         If ((Moneda_US <> .fields("ME")) Or (Mifecha <> .fields("Fecha")) Or TipoProc <> .fields("TP")) Then
            'Printer.Line (Ancho(0), PosLinea)-(Ancho(7), PosLinea), QBColor(0)
            PosLinea = PosLinea + 0.1
            PrinterVariables Ancho(4), PosLinea, "TOTALES"
            PrinterVariables Ancho(5), PosLinea, Debe
            PrinterVariables Ancho(6), PosLinea, Haber
            Debitos = Debitos + Debe
            Creditos = Creditos + Haber
            Debe = 0: Haber = 0: Debe_ME = 0: Haber_ME = 0
            PosLinea = PosLinea + 0.4
            'Printer.Line (Ancho(0), PosLinea)-(Ancho(7), PosLinea), QBColor(0)
            PosLinea = PosLinea + 0.3
            TipoProc = .fields("TP")
            Moneda_US = .fields("ME")
            Mifecha = .fields("Fecha")
            PrinterFields Ancho(0), PosLinea, .fields("ME"), False
            PrinterFields Ancho(1), PosLinea, .fields("Fecha"), False
            'MsgBox "Hola"
         End If
         End If
         If Codigo <> .fields("CodigoU") Then
            If TipoCaja Then
               PrinterFields Ancho(8), PosLinea, .fields("CodigoU"), False
            Else
               PrinterFields Ancho(7), PosLinea, .fields("CodigoU"), False
            End If
            Codigo = .fields("CodigoU")
         End If
         PrinterFields Ancho(2), PosLinea, .fields("TP"), False
         PrinterFields Ancho(3), PosLinea, .fields("Papeleta_No"), False
         PrinterFields Ancho(4), PosLinea, .fields("Cuenta_No"), False
         If EsFlujoCaja Then
            PrinterFields Ancho(5), PosLinea, .fields("Depositos"), False
            PrinterFields Ancho(6), PosLinea, .fields("Retiros"), False
            PrinterFields Ancho(7), PosLinea, .fields("Detalle"), False
            Select Case .fields("TP")
              Case "BOVE": Debe = Debe + .fields("Depositos")
                           Haber = Haber + .fields("Retiros")
              Case "APER", "DEP", "N/CE": Debe = Debe + .fields("Depositos")
              Case "RET", "CIER": Haber = Haber + .fields("Retiros")
              Case "N/DC": MontoCert = MontoCert + .fields("Retiros")
              Case "N/DG": MontoAper = MontoAper + .fields("Retiros")
          End Select
            'Debe = Debe + .Fields("Depositos")
            'Haber = Haber + .Fields("Retiros")
         Else
            PrinterFields Ancho(5), PosLinea, .fields("Debitos"), False
            PrinterFields Ancho(6), PosLinea, .fields("Creditos"), False
            Select Case .fields("TP")
              Case "BOVE": Debe = Debe + .fields("Debitos")
                           Haber = Haber + .fields("Creditos")
              Case "APER", "DEP": Debe = Debe + .fields("Debitos")
              Case "RET": Haber = Haber + .fields("Creditos")
              Case "N/DC": MontoCert = MontoCert + .fields("Creditos")
              Case "N/DG": MontoAper = MontoAper + .fields("Creditos")
            End Select
''            Debe = Debe + .Fields("Debitos")
''            Haber = Haber + .Fields("Creditos")
         End If
         'MsgBox Debe & "..." & Haber
         PosLinea = PosLinea + 0.4
         If PosLinea >= LimiteAlto Then
            'Printer.Line (Ancho(0), PosLinea)-(Ancho(7), PosLinea), QBColor(0)
            Printer.NewPage
            EncabezadoData Datas
            Printer.FontName = TipoArialNarrow
            Printer.FontSize = SizeLetra
         End If
        .MoveNext
      Loop
  End If
End With
'Printer.Line (Ancho(0), PosLinea)-(Ancho(7), PosLinea), QBColor(0)
PosLinea = PosLinea + 0.1
If EsGrupo Then
   PrinterVariables Ancho(4), PosLinea, "TOTALES"
   PrinterVariables Ancho(5), PosLinea, Debe
   PrinterVariables Ancho(6), PosLinea, Haber
   PosLinea = PosLinea + 0.4
End If
Else
  EncabezadoData Datas
End If
Printer.Line (Ancho(0), PosLinea)-(Ancho(7), PosLinea), QBColor(0)
PosLinea = PosLinea + 0.1
Debe = Debe + Debitos
Haber = Haber + Creditos
'Debe = 0: Haber = 0
Debe_ME = 0: Haber_ME = 0
With Datas.Recordset
 If .RecordCount > 0 Then
    .MoveFirst
     TipoProc = .fields("TP")
     Debitos = 0: Creditos = 0
     Do While Not .EOF
        If Resumido Then
        If TipoProc <> .fields("TP") Then
           PrinterVariables Ancho(2), PosLinea, TipoProc
           PrinterVariables Ancho(4), PosLinea, "TOTALES"
           PrinterVariables Ancho(5), PosLinea, Debitos
           PrinterVariables Ancho(6), PosLinea, Creditos
           TipoProc = .fields("TP")
           Debitos = 0: Creditos = 0
           PosLinea = PosLinea + 0.4
           If PosLinea >= LimiteAlto Then
            'Printer.Line (Ancho(0), PosLinea)-(Ancho(7), PosLinea), QBColor(0)
            Printer.NewPage
            EncabezadoData Datas
            Printer.FontSize = SizeLetra
           End If
        End If
        End If
        If .fields("ME") Then
            If EsFlujoCaja Then
               Debe_ME = Debe_ME + .fields("Depositos")
               Haber_ME = Haber_ME + .fields("Retiros")
            Else
               Debe_ME = Debe_ME + .fields("Debitos")
               Haber_ME = Haber_ME + .fields("Creditos")
            End If
        Else
            If EsFlujoCaja Then
               Debitos = Debitos + .fields("Depositos")
               Creditos = Creditos + .fields("Retiros")
               'Debe = Debe + .Fields("Depositos")
               'Haber = Haber + .Fields("Retiros")
            Else
               Debitos = Debitos + .fields("Debitos")
               Creditos = Creditos + .fields("Creditos")
               'Debe = Debe + .Fields("Debitos")
               'Haber = Haber + .Fields("Creditos")
            End If
        End If
       .MoveNext
     Loop
 End If
End With
If Resumido Then
   PrinterVariables Ancho(2), PosLinea, TipoProc
   PrinterVariables Ancho(4), PosLinea, "TOTALES"
   PrinterVariables Ancho(5), PosLinea, Debitos
   PrinterVariables Ancho(6), PosLinea, Creditos
   PosLinea = PosLinea + 0.5
End If
PrinterVariables Ancho(1), PosLinea, "SALDO ANTERIOR: "
PrinterVariables Ancho(4), PosLinea, Saldo_Anterior
PosLinea = PosLinea + 0.5
If TipoCaja Then
   PrinterTexto Ancho(1), PosLinea, "TOTALES INGRESOS MN"
   PrinterVariables Ancho(4), PosLinea, Debe
   PosLinea = PosLinea + 0.5
   PrinterTexto Ancho(1), PosLinea, "TOTALES EGRESOS MN"
   PrinterVariables Ancho(4), PosLinea, Haber
   PosLinea = PosLinea + 0.5
   PrinterTexto Ancho(1), PosLinea, "TOTALES CERTIFICADO MN"
   PrinterVariables Ancho(4), PosLinea, MontoCert
   PosLinea = PosLinea + 0.5
   PrinterTexto Ancho(1), PosLinea, "TOTALES APERTURA MN"
   PrinterVariables Ancho(4), PosLinea, MontoAper
   PosLinea = PosLinea + 0.5
   PrinterTexto Ancho(1), PosLinea, "SALDO DE CAJA MN"
   PrinterVariables Ancho(4), PosLinea, Saldo_Anterior + Debe - Haber
Else
   PrinterTexto Ancho(1), PosLinea, "TOTAL CREDITOS MN"
   PrinterVariables Ancho(4), PosLinea, Haber
   PosLinea = PosLinea + 0.5
   PrinterTexto Ancho(1), PosLinea, "TOTALES DEBITOS MN"
   PrinterVariables Ancho(4), PosLinea, Debe
   PosLinea = PosLinea + 0.5
   PrinterTexto Ancho(1), PosLinea, "SALDO LIBRETAS MN"
   PrinterVariables Ancho(4), PosLinea, Saldo_Anterior + Haber - Debe
End If

PosLinea = PosLinea - 1
PrinterTexto Ancho(6), PosLinea, String(22, "_")
PosLinea = PosLinea + 0.4
PrinterTexto Ancho(6) + 0.2, PosLinea, "Cajero: " & CodigoUsuario
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

Public Sub ImprimirSaldosLibretas(Datas As Adodc, _
                                  SaldoD As Currency, _
                                  SaldoC As Currency)
Dim SizeLetra As Integer
On Error GoTo Errorhandler
RatonReloj
Mensajes = "Seguro de Imprimir en:" & Chr(13) & Chr(13) & Printer.DeviceName & "?"
Titulo = "IMPRESION"
If BoxMensaje = vbYes Then
InicioX = 0.5: InicioY = 0
SizeLetra = 8
'Escala_Centimetro FormaImp, TipoTimes, SizeLetra
DataAnchoCampos InicioX, Datas, SizeLetra, TipoTimes, 1
Pagina = 1
'Iniciamos la impresion
Printer.FontBold = False
With Datas.Recordset
     .MoveFirst
      EncabezadoData Datas
      Printer.FontSize = SizeLetra
      Do While Not .EOF
         PrinterAllFields CantCampos, PosLinea, Datas, False, False
         PosLinea = PosLinea + 0.36
         If PosLinea >= LimiteAlto Then
            Printer.Line (Ancho(0), PosLinea)-(Ancho(CantCampos), PosLinea), QBColor(0)
            Printer.NewPage
            EncabezadoData Datas
            Printer.FontSize = SizeLetra
         End If
        .MoveNext
      Loop
End With
Printer.Line (InicioX, PosLinea)-(Ancho(CantCampos), PosLinea), QBColor(0)
PosLinea = PosLinea + 0.05
Printer.Line (InicioX, PosLinea)-(Ancho(CantCampos), PosLinea), QBColor(0)
PosLinea = PosLinea + 0.05
PrinterTexto Ancho(3), PosLinea, "T O T A L"
PrinterVariables Ancho(4), PosLinea, SaldoD
PrinterVariables Ancho(5), PosLinea, SaldoC
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


Public Sub ImprimirDataLibreta(Datas As Adodc, _
                               FinDoc As Boolean, _
                               FormaImp As Byte, _
                               SizeLetra As Integer)
On Error GoTo Errorhandler
RatonReloj
With Datas.Recordset
 If .RecordCount > 0 Then
     InicioX = 0.5: InicioY = 0
     'Escala_Centimetro FormaImp, TipoTimes, SizeLetra
     DataAnchoCampos InicioX, Datas, SizeLetra, TipoTimes, FormaImp
     Ancho(0) = 0.6
     Ancho(1) = 2
     Ancho(2) = 3
     Ancho(3) = 5.5
     Ancho(4) = 8
     Pagina = 1
     'Iniciamos la impresion
     Printer.FontBold = False
    .MoveFirst
     'EncabezadoData Datas
     Printer.FontSize = SizeLetra
     Do While Not .EOF
        PosLinea = (0.5 * (.fields("ID") - 1)) + 3.8
        PrinterAllFields CantCampos, PosLinea, Datas, False, False
        PosLinea = PosLinea + 0.5
        If .fields("ID") >= 34 Then
            Printer.NewPage
            'EncabezadoData Datas
            Printer.FontSize = SizeLetra
        End If
       .MoveNext
     Loop
     Printer.EndDoc
 End If
End With
RatonNormal
Exit Sub
Errorhandler:
    RatonNormal
    ErrorDeImpresion
    Exit Sub
End Sub

Public Sub ImprimirPrestamosVencidos(Datas As Adodc, _
                                     FinDoc As Boolean, _
                                     FormaImp As Byte, _
                                     SizeLetra As Integer, _
                                     Opc_P As Boolean, _
                                     ConCapital As Byte, _
                                     Optional ListadoPrest As Boolean)
On Error GoTo Errorhandler
RatonReloj
InicioX = 0.5: InicioY = 0
Tasa = 0: Cuota = 0
Pagina = 1
'Escala_Centimetro FormaImp, TipoTimes, SizeLetra
Mensajes = "Seguro de Imprimir en:" & vbCrLf & Printer.DeviceName & "?"
Titulo = "IMPRESION"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
   RatonReloj
If Opc_P Then
   DataAnchoCampos InicioX, Datas, SizeLetra, TipoTimes, FormaImp
Else
   DataAnchoCampos InicioX, Datas, SizeLetra, TipoTimes, 1
   Ancho(0) = 2    'V
   Ancho(1) = 3    'Apellidos
   Ancho(2) = 3    'Nombres
   Ancho(3) = 3    'Cuenta_No
   Ancho(4) = 3    'Credito_No
   Ancho(5) = 3    'Mes_No
   Ancho(6) = 5    'Fecha
   Ancho(7) = 7.5  'Fecha_C
   Ancho(8) = 10   'Capital
   Ancho(9) = 13   'Pagos
   Ancho(10) = 16  'Saldo
   Ancho(11) = 19
End If
'Iniciamos la impresion
Printer.FontBold = False
With Datas.Recordset
     .MoveFirst
      EncabezadoData Datas
      If Opc_P = False Then
         Cadena = .fields("Cliente")
         PrinterVariables Ancho(0), PosLinea, Cadena
         PosLinea = PosLinea + 0.4
         Cadena = "Cuenta_No. " & .fields("Cuenta_No") & "  Credito No. " & .fields("Credito_No")
         PrinterVariables Ancho(0), PosLinea, Cadena
         PosLinea = PosLinea + 0.45
         Printer.Line (Ancho(0), PosLinea)-(Ancho(CantCampos), PosLinea), QBColor(0)
         PosLinea = PosLinea + 0.05
      End If
      Numero = .fields("Credito_No")
      Do While Not .EOF
         Printer.FontSize = SizeLetra
         Printer.FontName = TipoTimes
         If Numero <> .fields("Credito_No") Then
            If Opc_P = False Then
               PosLinea = PosLinea + 0.05
               PrinterTexto Ancho(6), PosLinea, "Total Capital"
               PrinterVariables Ancho(8), PosLinea, CCur(Cuota)
               Cuota = 0
               PosLinea = PosLinea + 0.5
               Cadena = .fields("Cliente")
               PrinterVariables Ancho(0), PosLinea, Cadena
               PosLinea = PosLinea + 0.4
               Cadena = "Cuenta_No. " & .fields("Cuenta_No") & "  Credito No. " & .fields("Credito_No")
               PrinterVariables Ancho(0), PosLinea, Cadena
               PosLinea = PosLinea + 0.45
               Printer.Line (Ancho(0), PosLinea)-(Ancho(CantCampos), PosLinea), QBColor(0)
               PosLinea = PosLinea + 0.05
            End If
            Numero = .fields("Credito_No")
            
         End If
         If Opc_P Then
            Tasa = Tasa + .fields("Saldo_Pendiente")
         Else
            Cuota = Cuota + .fields("Capital")
            Tasa = Tasa + .fields("Capital")
         End If
         If Opc_P Then
            PrinterAllFields CantCampos, PosLinea, Datas, False, False
         Else
            If ListadoPrest Then
               PrinterFields Ancho(0), PosLinea, .fields("T")
            Else
               PrinterFields Ancho(0), PosLinea, .fields("V")
            End If
            PrinterFields Ancho(5), PosLinea, .fields("Cuota_No")
            PrinterFields Ancho(6), PosLinea, .fields("Fecha")
            PrinterFields Ancho(7), PosLinea, .fields("Fecha_C")
            PrinterFields Ancho(8), PosLinea, .fields("Capital")
            PrinterFields Ancho(9), PosLinea, .fields("Pagos")
            PrinterFields Ancho(10), PosLinea, .fields("Saldo")
         End If
         
         PosLinea = PosLinea + 0.35
         If PosLinea > LimiteAlto Then
            Printer.Line (Ancho(0), PosLinea)-(Ancho(CantCampos), PosLinea), QBColor(0)
            Printer.NewPage
            EncabezadoData Datas
            Printer.FontSize = SizeLetra
            Printer.FontName = TipoTimes
         End If
        .MoveNext
      Loop
End With
If Opc_P = False Then
   PrinterTexto Ancho(6), PosLinea, "Total Capital"
   PrinterVariables Ancho(8), PosLinea, CCur(Cuota)
   PosLinea = PosLinea + 0.5
End If
Printer.Line (Ancho(0), PosLinea)-(Ancho(CantCampos), PosLinea), QBColor(0)
PosLinea = PosLinea + 0.05
Printer.Line (Ancho(0), PosLinea)-(Ancho(CantCampos), PosLinea), QBColor(0)
PosLinea = PosLinea + 0.1
If Opc_P Then
   PrinterTexto Ancho(8), PosLinea, "T O T A L"
   PrinterVariables Ancho(9), PosLinea, CCur(Tasa)
Else
   PrinterTexto Ancho(6), PosLinea, "T O T A L"
   PrinterVariables Ancho(8), PosLinea, CCur(Tasa)
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

Public Sub Imprimir_Vencidos(Datas As Adodc, _
                             FinDoc As Boolean, _
                             FormaImp As Byte, _
                             SizeLetra As Integer, _
                             TipoRep As Integer)
On Error GoTo Errorhandler
Mensajes = "Seguro de Imprimir en:" & vbCrLf & Printer.DeviceName & "?"
Titulo = "IMPRESION"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
RatonReloj
InicioX = 0.5: InicioY = 0
DataAnchoCampos InicioX, Datas, SizeLetra, TipoTimes, Orientacion_Pagina
Pagina = 1
Ancho(0) = 0.5   'Credito_No
Ancho(1) = 2     'Apellidos_Nombres
Ancho(2) = 8     'Cuenta_No
If FormaImp = 1 Then
   Ancho(3) = 9.9   ' Direccion
   Ancho(4) = 9.9   ' Sector
   Ancho(5) = 9.9   ' Area
   Ancho(6) = 9.9   ' Telefono/TelefonoT
   Ancho(7) = 14.3  ' Fecha
   Ancho(8) = 15.7  ' Cuota
   Ancho(9) = 17.2  ' Pagos
   Ancho(10) = 19
Else
   Ancho(3) = 9.8   ' Direccion
   Ancho(4) = 14    ' Sector
   Ancho(5) = 16.6  ' Area
   Ancho(6) = 17.5  ' Telefono/TelefonoT
   Ancho(7) = 20.9  ' Fecha
   Ancho(8) = 22.3  ' Cuota
   Ancho(9) = 23.7  ' Pagos
   Ancho(10) = 25.5
End If
Total = 0
Pagina = 1
CantCampos = 10
'Iniciamos la impresion
Printer.FontBold = False
With Datas.Recordset
 If .RecordCount > 0 Then
     .MoveFirst
      Encabezado_Venc Datas, FormaImp, TipoRep
      Printer.FontSize = SizeLetra
      Do While Not .EOF
         PrinterFields Ancho(0), PosLinea, .fields("Credito_No")
         PrinterTexto Ancho(1), PosLinea, .fields("Cliente")
         PrinterFields Ancho(2), PosLinea, .fields("Cuenta_No")
         If FormaImp = 2 Then
            PrinterFields Ancho(3), PosLinea, .fields("Direccion")
            PrinterFields Ancho(4), PosLinea, .fields("Sector")
            PrinterFields Ancho(5), PosLinea, .fields("Area")
         End If
         Cadena = " "
         If .fields("Telefono") <> Ninguno Or .fields("TelefonoT") <> Ninguno Then
             Cadena = .fields("Telefono") & "/" & .fields("TelefonoT")
         Else
             Cadena = " "
         End If
         PrinterTexto Ancho(6), PosLinea, Cadena
         PrinterFields Ancho(7), PosLinea, .fields("Fecha")
         PrinterFields Ancho(8), PosLinea, .fields("Cuota_No")
         'If TipoRep = 1 Then
            PrinterFields Ancho(9), PosLinea, .fields("Pagos")
         'Else
         '   PrinterFields Ancho(9), PosLinea, .Fields("Capital")
         'End If
         'PrinterAllFields CantCampos, PosLinea, Datas, True
         'If TipoRep = 1 Then
            Total = Total + .fields("Pagos")
         'Else
         '   Total = Total + .Fields("Capital")
         'End If
         PosLinea = PosLinea + 0.4
         If PosLinea >= LimiteAlto Then
            Printer.Line (Ancho(0), PosLinea)-(Ancho(CantCampos), PosLinea), QBColor(0)
            Printer.NewPage
            Encabezado_Venc Datas, FormaImp, TipoRep
            Printer.FontSize = SizeLetra
         End If
        .MoveNext
      Loop
 End If
End With
Printer.Line (InicioX, PosLinea)-(Ancho(CantCampos), PosLinea), QBColor(0)
PosLinea = PosLinea + 0.05
Printer.Line (InicioX, PosLinea)-(Ancho(CantCampos), PosLinea), QBColor(0)
PosLinea = PosLinea + 0.1
PrinterTexto Ancho(6), PosLinea, "TOTAL POR COBRAR"
PrinterVariables Ancho(9) - 0.6, PosLinea, Total
RatonNormal
If FinDoc Then Printer.EndDoc
End If
Exit Sub
Errorhandler:
    RatonNormal
    ErrorDeImpresion
    Exit Sub
End Sub

Public Sub Imprimir_Libreta(CuentaNo As String, _
                            Datas As Adodc, _
                            FormaImp As Byte, _
                            SizeLetra As Integer, _
                            LineaNo As Long, _
                            Optional SaldoAnt)
Dim EspLinea As Single
Dim EsMio As Boolean
Dim ImpReverso As Boolean
On Error GoTo Errorhandler
EsMio = False
LBConLineas = ProcesarSeteos("LB")
Mensajes = "Imprimir Transaccion en Libreta"
Titulo = "IMPRIMIR LIBRETA"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
   Mensajes = "Imprimir Transaccion en el Reverso?"
   Titulo = "LADO DE IMPRESION"
   If BoxMensaje(True) = vbYes Then ImpReverso = True Else ImpReverso = False
   'MsgBox ImpReverso
   RatonReloj
   Apellidos = ""
   EspLinea = SetD(8).PosY
   InicioX = 0.5
   
   If Cartilla_No <= 0 Then
      Cartilla_No = Val(InputBox("CARTILLA No.", "NUMERACION DE CARTILLA", Cartilla_No))
      sSQL = "UPDATE Trans_Libretas " _
           & "SET Cartilla_No = " & Cartilla_No & " " _
           & "WHERE Cuenta_No = '" & CuentaNo & "' " _
           & "AND Cartilla_No = 0 "
      Ejecutar_SQL_SP sSQL
      sSQL = "INSERT INTO Trans_Cartillas " _
           & "(Fecha,Cuenta_No,Cartilla_No,Detalle,Item) VALUES " _
           & "('" & BuscarFecha(FechaSistema) & "','" & CuentaNo & "'," & Cartilla_No & ",'ACTUALIZACION LIBRETA','" & NumEmpresa & "')"
      Ejecutar_SQL_SP sSQL
   End If
   Printer.FontSize = 10
   Printer.Font = TipoTimes
   
   sSQL = "SELECT Cl.Representante,Cl.Cliente,Cl.CI_RUC,Cl.Direccion " _
        & "FROM Clientes As Cl,Clientes_Datos_Extras As C " _
        & "WHERE C.Cuenta_No = '" & CuentaNo & "' " _
        & "AND Cl.Codigo = C.Codigo "
   Select_Adodc Datas, sSQL, False
   DataAnchoCampos InicioX, Datas, SizeLetra, TipoTerminal, FormaImp
   If Datas.Recordset.RecordCount > 0 Then
      Nombres = Datas.Recordset.fields("Cliente")
      If Len(Datas.Recordset.fields("Representante")) > 1 Then Apellidos = Datas.Recordset.fields("Representante")
      CICliente = Datas.Recordset.fields("CI_RUC")
      DirCliente = Datas.Recordset.fields("Direccion")
      EsMio = True
   End If
   sSQL = "SELECT * " _
        & "FROM Trans_Libretas " _
        & "WHERE Cuenta_No = '" & CuentaNo & "' " _
        & "AND IP = " & Val(adFalse) & " " _
        & "ORDER BY Fecha,IDT,Hora,ID "
   Select_Adodc Datas, sSQL, False
   With Datas.Recordset
    If .RecordCount > 0 Then
       'MsgBox LineaNo
        Do While Not .EOF
           'If LineaNo > SetD(7).Tamaño Then LineaNo = 1
           If LineaNo = 0 Then LineaNo = 1
           If LineaNo = 1 Then
              Printer.FontBold = True
              If ImpReverso = False Then
                 Printer.FontSize = SetD(11).Tamaño
                 PrinterTexto SetD(11).PosX, SetD(11).PosY, "LIBRETA DE AHORRO"
                 Printer.FontSize = SetD(3).Tamaño
                 PrinterTexto SetD(3).PosX, SetD(3).PosY, Nombres
                 Printer.FontSize = SetD(30).Tamaño
                 PrinterTexto SetD(30).PosX, SetD(30).PosY, Apellidos
                 Printer.FontSize = SetD(4).Tamaño
                 PrinterTexto SetD(4).PosX, SetD(4).PosY, DirCliente
                 Printer.FontSize = SetD(5).Tamaño
                 PrinterTexto SetD(5).PosX, SetD(5).PosY, "C.I./R.U.C. " & CICliente
                 Cadena = CuentaNo
                 If SetD(31).PosX > 0 And SetD(31).PosY > 0 Then Cadena = SetD(31).Encabezado & " " & Cadena
                 Printer.FontSize = SetD(2).Tamaño
                 PrinterTexto SetD(2).PosX, SetD(2).PosY, Cadena
              End If
             'MsgBox ImpReverso
              If ImpReverso Then
                 PosLinea = SetD(9).PosY + (LineaNo * EspLinea)
              Else
                 PosLinea = SetD(22).PosY + (LineaNo * EspLinea)
              End If
              Printer.FontSize = SetD(16).Tamaño
              PrinterTexto SetD(16).PosX, PosLinea, "FECHA"
              Printer.FontSize = SetD(17).Tamaño
              PrinterTexto SetD(17).PosX, PosLinea, "TIPO"
              Printer.FontSize = SetD(20).Tamaño
              PrinterTexto SetD(20).PosX, PosLinea, "M O N T O "
              Printer.FontSize = SetD(18).Tamaño
              PrinterTexto SetD(18).PosX, PosLinea, "  DEBITOS"
              Printer.FontSize = SetD(19).Tamaño
              PrinterTexto SetD(19).PosX, PosLinea, "  CREDITOS"
              Printer.FontSize = SetD(21).Tamaño
              PrinterTexto SetD(21).PosX, PosLinea, "S A L D O "
              If (SetD(16).PosX + SetD(17).PosX + SetD(20).PosX + SetD(21).PosX) > 0 Then
                 LineaNo = LineaNo + 1
              End If
              If SetD(13).PosX <> 0 Then        'Salto de Pagina
                 PosLinea = 1: LineaNo = 1
                 ImpReverso = Not ImpReverso
                 Printer.NewPage
              End If
           End If
           Printer.FontSize = SetD(22).Tamaño
           Printer.FontBold = False
'           If .Fields("ID") <> LineaNo Then .Fields("ID") = LineaNo
          'Verificar si imprimimos al reverso
           'MsgBox ImpReverso
           If ImpReverso Then
              PosLinea = SetD(9).PosY + (LineaNo * EspLinea)
           Else
              PosLinea = SetD(22).PosY + (LineaNo * EspLinea)
           End If
          'MsgBox PosLinea
           If EsMio Then
              'ImpCeros = True
              Total = .fields("Saldo_Cont")
              PrinterVariables SetD(29).PosX, PosLinea, CCur(Total)
           Else
              SaldoAnt = Redondear(SaldoAnt, 2) + .fields("Creditos") - .fields("Debitos")
              PrinterVariables SetD(29).PosX, PosLinea, CCur(SaldoAnt)
           End If
           If .fields("Creditos") > 0 Then
               PrinterVariables SetD(27).PosX, PosLinea, CCur(.fields("Creditos"))
           End If
           If .fields("Debitos") > 0 Then
               PrinterVariables SetD(26).PosX, PosLinea, CCur(.fields("Debitos"))
           End If
           PrinterVariables SetD(28).PosX, PosLinea, CCur(Abs(.fields("Creditos") - .fields("Debitos")))
           PrinterFields SetD(24).PosX, PosLinea, .fields("Fecha"), False
           PrinterFields SetD(25).PosX, PosLinea, .fields("TP"), False
           PrinterTexto SetD(23).PosX, PosLinea, CStr(LineaNo)
           'MsgBox LineaNo
           LineaNo = LineaNo + 1
           If LineaNo > SetD(10).Tamaño And ImpReverso Then
              PosLinea = 1
              LineaNo = 0
              ImpReverso = Not ImpReverso
              Printer.NewPage
           End If
           If LineaNo > SetD(7).Tamaño And ImpReverso = False Then
              PosLinea = 1
              LineaNo = 0
              ImpReverso = Not ImpReverso
              Printer.NewPage
           End If
          .Update
          .MoveNext
        Loop
        sSQL = "UPDATE Trans_Libretas " _
             & "SET IP = -1 " _
             & "WHERE Cuenta_No = '" & CuentaNo & "' "
        Ejecutar_SQL_SP sSQL
     End If
   End With
   RatonNormal
   Printer.EndDoc
End If
Exit Sub
Errorhandler:
    RatonNormal
    ErrorDeImpresion
    Exit Sub
End Sub

Public Sub Imprimir_Certificados(CuentaNo As String, _
                                 Datas As Adodc, _
                                 FormaImp As Byte, _
                                 SizeLetra As Integer, _
                                 LineaNo As Byte, _
                                 Optional SaldoAnt)
Dim EspLinea As Single
Dim EsMio As Boolean
Dim ImpReverso As Boolean
On Error GoTo Errorhandler
EsMio = False
LBConLineas = ProcesarSeteos("LB")
Mensajes = "Imprimir Transaccion en Libreta"
Titulo = "IMPRIMIR LIBRETA"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
   Mensajes = "Imprimir Transaccion en el Reverso?"
   Titulo = "LADO DE IMPRESION"
   If BoxMensaje(True) = vbYes Then ImpReverso = True Else ImpReverso = False
   
   'MsgBox ImpReverso
   RatonReloj
   EspLinea = SetD(8).PosY
   InicioX = 0.5
   DataAnchoCampos InicioX, Datas, SizeLetra, TipoTerminal, FormaImp
   Printer.FontSize = 10
   Printer.Font = TipoTimes
   sSQL = "SELECT Cl.Representante,Cl.Cliente,Cl.CI_RUC,Cl.Direccion " _
        & "FROM Clientes As Cl,Clientes_Datos_Extras As C " _
        & "WHERE Cuenta_No = '" & CuentaNo & "' " _
        & "AND Cl.Codigo = C.Codigo "
   Select_Adodc Datas, sSQL, False
   If Datas.Recordset.RecordCount > 0 Then
      If Len(Datas.Recordset.fields("Representante")) > 1 Then
         Nombres = Datas.Recordset.fields("Representante") & " - " & Datas.Recordset.fields("Cliente")
      Else
         Nombres = Datas.Recordset.fields("Cliente")
      End If
      CICliente = Datas.Recordset.fields("CI_RUC")
      DirCliente = Datas.Recordset.fields("Direccion")
      EsMio = True
   End If
   sSQL = "SELECT * " _
        & "FROM Trans_Certificados " _
        & "WHERE Cuenta_No = '" & CuentaNo & "' " _
        & "ORDER BY Fecha,IDT,Hora,ID "
   Select_Adodc Datas, sSQL, False
   SaldoAnt = 0
   With Datas.Recordset
    If .RecordCount > 0 Then
       'MsgBox LineaNo
        Do While Not .EOF
          'MsgBox PosLinea
           SaldoAnt = SaldoAnt + .fields("Creditos") - .fields("Debitos")
          .fields("Saldo_Disp") = SaldoAnt
          .fields("Saldo_Cont") = SaldoAnt
          .Update
          .MoveNext
        Loop
     End If
   End With
  'Vuelvo a poner el inicio de
   sSQL = "SELECT * " _
        & "FROM Trans_Certificados " _
        & "WHERE Cuenta_No = '" & CuentaNo & "' " _
        & "AND IP = " & Val(adFalse) & " " _
        & "ORDER BY Fecha,IDT,Hora,ID "
   Select_Adodc Datas, sSQL, False
   With Datas.Recordset
    If .RecordCount > 0 Then
       'MsgBox LineaNo
        Do While Not .EOF
           If LineaNo > SetD(7).Tamaño Then LineaNo = 1
           If LineaNo = 0 Then LineaNo = 1
           If LineaNo = 1 Then
              Printer.FontBold = True
              If ImpReverso = False Then
                 Printer.FontSize = SetD(11).Tamaño
                 PrinterTexto SetD(11).PosX, SetD(11).PosY, "CERTIFICADOS DE APORTACION"
                 Printer.FontSize = SetD(3).Tamaño
                 PrinterTexto SetD(3).PosX, SetD(3).PosY, Nombres
                 Printer.FontSize = SetD(4).Tamaño
                 PrinterTexto SetD(4).PosX, SetD(4).PosY, DirCliente
                 Printer.FontSize = SetD(5).Tamaño
                 PrinterTexto SetD(5).PosX, SetD(5).PosY, CICliente
                 Printer.FontSize = SetD(2).Tamaño
                 PrinterTexto SetD(2).PosX, SetD(2).PosY, CuentaNo
                 
              End If
             'MsgBox ImpReverso
              If ImpReverso Then
                 PosLinea = SetD(9).PosY + (LineaNo * EspLinea)
              Else
                 PosLinea = SetD(22).PosY + (LineaNo * EspLinea)
              End If
              Printer.FontSize = SetD(16).Tamaño
              PrinterTexto SetD(16).PosX, PosLinea, "FECHA"
              Printer.FontSize = SetD(17).Tamaño
              PrinterTexto SetD(17).PosX, PosLinea, "TIPO"
              Printer.FontSize = SetD(20).Tamaño
              PrinterTexto SetD(20).PosX, PosLinea, "M O N T O "
              Printer.FontSize = SetD(18).Tamaño
              PrinterTexto SetD(18).PosX, PosLinea, "  DEBITOS"
              Printer.FontSize = SetD(19).Tamaño
              PrinterTexto SetD(19).PosX, PosLinea, "  CREDITOS"
              Printer.FontSize = SetD(21).Tamaño
              PrinterTexto SetD(21).PosX, PosLinea, "S A L D O "
              If (SetD(16).PosX + SetD(17).PosX + SetD(20).PosX + SetD(21).PosX) > 0 Then
                 LineaNo = LineaNo + 1
              End If
              If SetD(13).PosX <> 0 Then
                 PosLinea = 1: LineaNo = 1
                 ImpReverso = Not ImpReverso
                 Printer.NewPage
              End If
           End If
           Printer.FontSize = SetD(22).Tamaño
           Printer.FontBold = False
'           If .Fields("ID") <> LineaNo Then .Fields("ID") = LineaNo
          'Verificar si imprimimos al reverso
           'MsgBox ImpReverso
           If ImpReverso Then
              PosLinea = SetD(9).PosY + (LineaNo * EspLinea)
           Else
              PosLinea = SetD(22).PosY + (LineaNo * EspLinea)
           End If
          'MsgBox PosLinea
           If EsMio Then
              'ImpCeros = True
              Total = .fields("Saldo_Cont")
              PrinterVariables SetD(29).PosX, PosLinea, CCur(Total)
           Else
              SaldoAnt = Redondear(SaldoAnt, 2) + .fields("Creditos") - .fields("Debitos")
              PrinterVariables SetD(29).PosX, PosLinea, CCur(SaldoAnt)
           End If
           If .fields("Creditos") > 0 Then
               PrinterVariables SetD(27).PosX, PosLinea, CCur(.fields("Creditos"))
           End If
           If .fields("Debitos") > 0 Then
               PrinterVariables SetD(26).PosX, PosLinea, CCur(.fields("Debitos"))
           End If
           PrinterVariables SetD(28).PosX, PosLinea, CCur(Abs(.fields("Creditos") - .fields("Debitos")))
           PrinterFields SetD(24).PosX, PosLinea, .fields("Fecha"), False
           PrinterFields SetD(25).PosX, PosLinea, .fields("TP"), False
           PrinterTexto SetD(23).PosX, PosLinea, CStr(LineaNo)
           'MsgBox LineaNo
           LineaNo = LineaNo + 1
           If LineaNo > SetD(10).Tamaño And ImpReverso Then
              PosLinea = 1
              ImpReverso = Not ImpReverso
              Printer.NewPage
           End If
           If LineaNo > SetD(7).Tamaño And ImpReverso = False Then
              PosLinea = 1
              ImpReverso = Not ImpReverso
              Printer.NewPage
           End If
          .Update
          .MoveNext
        Loop
        sSQL = "UPDATE Trans_Certificados " _
             & "SET IP = -1 " _
             & "WHERE Cuenta_No = '" & CuentaNo & "' "
        Ejecutar_SQL_SP sSQL
     End If
   End With
   RatonNormal
   Printer.EndDoc
End If
Exit Sub
Errorhandler:
    RatonNormal
    ErrorDeImpresion
    Exit Sub
End Sub

Public Function Interes_Caja(TSaldoDisp As Currency, Tipo_Libreta As String) As Currency
Dim AdoReg As ADODB.Recordset
Dim TSaldo As Currency
'Sacar tasa de intereses segun saldo disponible
 Set AdoReg = New ADODB.Recordset
 AdoReg.CursorType = adOpenStatic
 AdoReg.CursorLocation = adUseClient
 TotalInteres = 0: TSaldo = TSaldoDisp
'MsgBox TSaldoDisp
 sSQL = "SELECT * " _
      & "FROM Catalogo_Interes " _
      & "WHERE ME = " & Val(adFalse) & " " _
      & "AND TP = 'C' " _
      & "AND Tipo = '" & Tipo_Libreta & "' " _
      & "AND Desde < " & TSaldoDisp & " " _
      & "AND Hasta >= " & TSaldoDisp & " "
 AdoReg.open sSQL, AdoStrCnn, , , adCmdText
 'MsgBox sSQL
 If AdoReg.RecordCount > 0 Then
    'Do While Not AdoReg.EOF
       'MsgBox AdoReg.Fields("Desde") & vbCrLf & AdoReg.Fields("Hasta")
       'If (AdoReg.Fields("Desde") <= TSaldoDisp) And (TSaldoDisp <= AdoReg.Fields("Hasta")) Then
          TotalInteres = AdoReg.fields("Interes")
     '     AdoReg.MoveLast
      'End If
    '   AdoReg.MoveNext
   ' Loop
 End If
 TSaldo = (TSaldoDisp * TotalInteres) / 36000
 If TSaldo < 0 Then TSaldo = 0
 Interes_Caja = Redondear(TSaldo, 4)
End Function

Public Sub Imprimir_Papeleta(NoNivel As Integer, _
                             ChequeS_N As Boolean, _
                             Fecha As String, _
                             Hora As String, _
                             Cuenta As String, _
                             TipoP As String, _
                             Valor As Currency, _
                             NCheq As String, _
                             NBanco As String, _
                             NCliente As String, _
                             Optional Es_Copia As Boolean)
Dim SizeLetra As Integer
Dim Copia As Boolean
On Error GoTo Errorhandler
Select Case TipoP
  Case "APEC", "APER", "CIER", "DDA", "DDAC", "DEP", "DEPP", "DEPC", "RDA", "RDAC", "RET", "RETC", "DEFR", "REFR"
       If NoNivel = 1 Then
          Titulo = "IMPRIMIR PAPELETA DE DEPOSITO/RETIRO"
          If Es_Copia Then
             Bandera = True
             Mensajes = vbCrLf & "IMPRIMIR COPIA DE LA TRANSACCION" & vbCrLf
             SetNombrePRN = SetPRN_2
             SetPapelPRNCad = SetPapelPRN_2
             Grafico_PV = Leer_Campo_Empresa("Grafico_PV")
             SetPapelPRN = CInt(SinEspaciosIzq(SetPapelPRNCad))
          Else
             Bandera = False
             Mensajes = vbCrLf & "IMPRIMIR TRANSACCION EN LA PAPELETA" & vbCrLf
             SetPrinters.Show 1
          End If
         'MsgBox SetNombrePRN
          If PonImpresoraDefecto(SetNombrePRN) Then
             RatonReloj
             InicioX = 0.5: InicioY = 0
             SizeLetra = 10
             Escala_Centimetro 1, TipoTimes, SizeLetra
             Pagina = 1
            'Iniciamos la impresion
             Printer.FontBold = True
             Printer.FontSize = SizeLetra
             If Es_Copia Then
                PosLinea = 0.1
                PCol = 0.5
                If Grafico_PV Then
                   PrinterPaint LogoTipo, PCol, PosLinea, 3, 1.5
                   PosLinea = PosLinea + 1.6
                End If
                PrinterTexto PCol, PosLinea, Empresa
                PosLinea = PosLinea + 0.4
                If UCaseStrg(Empresa) <> UCaseStrg(NombreComercial) Then
                   PrinterTexto PCol, PosLinea, NombreComercial
                   PosLinea = PosLinea + 0.8
                End If
             Else
                PosLinea = 0.01
                PCol = 1
                If ChequeS_N Then PCol = 2
             End If
             Printer.FontSize = SizeLetra
             PrinterVariables PCol, PosLinea, "Agencia: " & Format$(NumEmpresa, "000")
             PrinterVariables PCol, PosLinea + 0.4, "Cajero(a): " & Cambio_Usuario_Inicial(NombreUsuario)  ' Cambio_CI_Caracter(CodigoUsuario)
             PrinterVariables PCol, PosLinea + 0.8, Fecha & " H:" & Hora
             PrinterVariables PCol, PosLinea + 1.2, "CTA. No. " & Cuenta
             PrinterVariables PCol, PosLinea + 1.6, NCliente
             PrinterVariables PCol, PosLinea + 2, TipoP & ": " & Moneda & " " & Format$(Valor, "#,##0.00")
             If ChequeS_N Then PrinterVariables PCol, PosLinea + 2.4, NBanco & ", Cheque No. " & NCheq
             Printer.FontBold = False
             If Es_Copia Then
                Printer.FontSize = 8
                Printer.FontBold = False
                PrinterTexto PCol, PosLinea + 2.8, String(70, "-")
                PrinterTexto PCol, PosLinea + 3.2, "Este documento es fiel copia de la transacción original"
             End If
             Printer.EndDoc
          End If
       End If
End Select
MensajeEncabData = ""
RatonNormal
Exit Sub
Errorhandler:
    RatonNormal
    ErrorDeImpresion
    Exit Sub
End Sub

Public Sub Imprimir_Pagare_No(Credito_No As String, _
                              Cuenta_No As String, _
                              Interes As Single, _
                              Interes_Mora As Single, _
                              Ciudad_Credito As String, _
                              Socio As String, _
                              CI_Socio As String, _
                              Dir_Socio As String, _
                              AdoTabla As Adodc, _
                              AdoConyugue As Adodc, _
                              AdoGarantes As Adodc)
On Error GoTo Errorhandler
Dim TextoPg(20) As String
With AdoGarantes.Recordset
 If .RecordCount > 0 Then
    .MoveFirst
     Opcion = .fields("Num")
     Do While Not .EOF
        If Opcion <> .fields("Num") Then
           Opcion = .fields("Num")
           PosLinea = PosLinea + 1.8
        End If
        If PosLinea > 27 Then
           PosLinea = 3.5
        End If
        Cta = .fields("Nombres")
        Cta_Sup = .fields("CI")
        If .fields("GC") Then PCol = 3 Else PCol = 10
       .MoveNext
     Loop
 End If
End With

Titulo = "IMPRESIONES"
Mensajes = "Imprimir Pagaré"
'If BoxMensaje = vbYes Then
Bandera = False
SetPrinters.Show 1
PGConLineas = ProcesarSeteos("PG")
If PonImpresoraDefecto(SetNombrePRN) Then
RatonReloj
DataAnchoCampos 1, AdoConyugue, 9, TipoCourierNew, 1, True
InicioX = 0.5: InicioY = 0
LimiteAlto = Printer.ScaleHeight - 1 'Limite de impresión a lo largo
LimiteAncho = Printer.ScaleWidth     'Limite de impresión a lo largo
AnchoPapel = Printer.ScaleWidth      'Limite de impresión a lo largo
Ancho(3) = 19
CantCampos = 3
Pagina = 1
Printer.FontSize = 9
Printer.FontBold = True
'Iniciamos la impresion
Printer.FontBold = False
With AdoTabla.Recordset
 If .RecordCount > 0 Then
    .MoveFirst
     Mifecha = .fields("Fecha")
     Total = .fields("Saldo")
    .MoveLast
     FechaTexto = .fields("Fecha")
 End If
End With
'MsgBox PGConLineas
If PGConLineas Then
   TextoPg(0) = "Debemos y pagaremos al " & Empresa & " " _
              & ", el lugar en que se nos reconvenga a la ordén del " & Empresa & " " _
              & NombreComercial & " " _
              & "la suma de " & Format$(Total, "#,##0.00") & "(" & Cambio_Letras(Total) & ") por igual valor que hemos recibido en dinero " _
              & "efectivo y en calidad de préstamo."
   TextoPg(1) = "Esta cantidad nos obligamos a pagar el interés del " & Interes & "% anual desde esta fecha hasta " _
              & "el vencimiento del plazo indicado. En el caso de mora, pagaremos un interés según las disposiciones " _
              & "legales establecidas por los entes de control, así como también pagaremos los gastos judiciales " _
              & "y extrajudiciales, inclusive honorarios profesionales, que ocasione el cobro de esta obligación, " _
              & "siendo suficiente prueba para establecer el monto de tales gastos la sola aseveración del acreedor."
   TextoPg(2) = "Afiel cumplimiento de los convenios, nos comprometemos con todos nuestros bienes presentes y futuros."
   TextoPg(3) = "Renunciamos domicilio y a toda ley o excepción que puduere favorecer en el juicio o fuera de él."
   TextoPg(4) = "Renunciamos también el derecho de interponer el recurso de apelación y el de hecho de las providencias " _
              & "que se expidieren en el juicio que, en relación al presente documento se dieren lugar."
   TextoPg(5) = "El pago no podrá hacerce por partes, sin protesto. Exímese de presentación para el pago y " _
              & "avisos por falta del mismno."
   TextoPg(6) = "Quedamos sometidos a los jueces de la ciudad de " & Ciudad_Credito & " o a los que el acreedor elija, " _
              & "para cuyo efecto renuncio fuero, domicilio y vecindad."
   TextoPg(7) = "Dejamos constancia expresa que el plazo de vista corre desde el " & FechaStrg(Mifecha) & " que firmamos " _
              & "al suscribir este pagaré."
   TextoPg(8) = "También dejamos constancia que el presente documento que firmo es totalmente negociable y transferible."
   TextoPg(9) = "AUTORIZACION: Autorizo al " & Empresa & " " & NombreComercial & ", para que en caso de mora, " _
              & "disponga de los valores que exista en nuestra Cuenta de Ahorro y Crédito, sin que se deba dar aviso alguno."
   TextoPg(10) = "VENCIMIENTO: " & FechaStrg(FechaTexto) & "."
  'Siguiente pagina
   Cta = "": Cta_Sup = ""
   With AdoConyugue.Recordset
    If .RecordCount > 0 Then
       .MoveLast
        Cta = .fields("Nombres")
        Cta_Sup = .fields("Cedula")
    End If
   End With
   TextoPg(11) = "Nos constituímos en fiadores solidarios, llanos pagaderos de los señores: " & Socio & " y/o su garante."
   If Cta <> "" Then TextoPg(11) = TextoPg(11) & "y " & Cta & ", "
   TextoPg(12) = TextoPg(12) _
               & "Por las obligaciones que han contraído con al " & Empresa & " " & NombreComercial & ", " _
               & "en el pagaré anterior haciendo deuda ajena deuda propia, renunciando los beneficios de exclusión " _
               & "de bienes de deudores principales, el de división y cualquier ley que pueda favorecer, así como la " _
               & "apelación y el recurso de hecho. Quedamos sometidos a los jueces de esta provincia o a la que elija " _
               & "el acreedor. Sin protesto."
               
   TextoPg(13) = "Para constancia se firma en la ciudad de " & Ciudad_Credito & ", hoy " & FechaStrg(Mifecha) & "."
   PosLinea = 0.5
   PrinterPaint LogoTipo, 1, PosLinea, 2, 1
   Printer.FontBold = True
   Printer.FontSize = 12: Printer.FontItalic = False
   PrinterTexto CentrarTextoEncab(UCaseStrg(Empresa), 2, 17), PosLinea, UCaseStrg(Empresa)
   PosLinea = PosLinea + 0.8
   If Len(NombreComercial) > 1 Then
      PrinterTexto CentrarTextoEncab(UCaseStrg(NombreComercial), 2, 17), PosLinea, UCaseStrg(NombreComercial)
      PosLinea = PosLinea + 0.8
   End If
   If Len(Direccion) > 1 Then
      Printer.FontSize = 10
      PrinterTexto CentrarTextoEncab(UCaseStrg(Direccion), 2, 17), PosLinea, UCaseStrg(Direccion)
      PosLinea = PosLinea + 0.8
   End If
   Printer.FontSize = 14
   PosLinea = PosLinea + 0.2
   Codigo = "PAGARE A LA ORDEN"
   PrinterTexto CentrarTextoEncab(UCaseStrg(Codigo), 2, 17), PosLinea, UCaseStrg(Codigo)
   PosLinea = PosLinea + 0.6
   PrinterTexto 2, PosLinea, "No. " & Credito_No
   PrinterTexto 13, PosLinea, "POR: " & Format$(Total, "#,##0.00")
   PosLinea = PosLinea + 0.8
   Printer.FontBold = False
   Printer.FontSize = 10
   For I = 0 To 13
       NumeroLineas = PrinterLineasMayor(2, PosLinea, TextoPg(I), 17)
       PosLinea = PosLinea + (NumeroLineas * 0.45)
   Next I
   PosLinea = PosLinea + 0.2
   Codigo = "D E U D O R"
   PrinterTexto CentrarTextoEncab(UCaseStrg(Codigo), 2, 17), PosLinea, UCaseStrg(Codigo)
   PosLinea = PosLinea + 2
   PrinterTexto 2, PosLinea, "F." & String(35, "_")
   PosLinea = PosLinea + 0.5
   PrinterTexto 2, PosLinea, "   " & Socio
   PosLinea = PosLinea + 0.5
   PrinterTexto 2, PosLinea, "   C.I. " & CI_Socio
   PosLinea = PosLinea + 0.6
   Codigo = "G A R A N T E"
   PrinterTexto CentrarTextoEncab(UCaseStrg(Codigo), 2, 17), PosLinea, UCaseStrg(Codigo)
   PosLinea = PosLinea + 2
   With AdoGarantes.Recordset
    If .RecordCount > 0 Then
       .MoveFirst
        Do While Not .EOF
           PrinterTexto 2, PosLinea, "F." & String(35, "_")
           PosLinea = PosLinea + 0.5
           PrinterTexto 2, PosLinea, "   " & .fields("Nombres")
           PosLinea = PosLinea + 0.5
           PrinterTexto 2, PosLinea, "   C.I. " & .fields("CI")
           PosLinea = PosLinea + 1
          .MoveNext
        Loop
    End If
   End With
Else



End If
RatonNormal
Printer.EndDoc
End If
Exit Sub
Errorhandler:
    RatonNormal
    ErrorDeImpresion
    Exit Sub
End Sub

Public Sub Imprimir_Dep_Alumno(NoNivel As Integer, _
                               ChequeS_N As Boolean, _
                               Fecha As String, _
                               Hora As String, _
                               Cuenta As String, _
                               TipoP As String, _
                               Valor As Currency, _
                               NCheq As String, _
                               NBanco As String)
Dim SizeLetra As Integer
On Error GoTo Errorhandler
If NoNivel = 1 Then
   Bandera = False
   Mensajes = "INGRESE LA PAPELETA DE DEPOSITO DE:" & vbCrLf & Beneficiario
   Titulo = "IMPRIMIR DEPOSITO ALUMNO"
   SetPrinters.Show 1
   If PonImpresoraDefecto(SetNombrePRN) Then
   RatonReloj
   InicioX = 0.5: InicioY = 0
   SizeLetra = 10
   Escala_Centimetro 1, TipoTimes, SizeLetra
   Pagina = 1
  'Iniciamos la impresion
   Printer.FontBold = True
   Printer.FontSize = SizeLetra
   PosLinea = 0.5: PCol = 3
   If ChequeS_N Then PCol = 2
   PrinterVariables PCol, PosLinea, "Agencia: " & Format$(NumEmpresa, "000")
   PrinterVariables PCol, PosLinea + 0.4, "Cajero(a): " & CodigoUsuario
   PrinterVariables PCol, PosLinea + 0.8, Fecha
   PrinterVariables PCol, PosLinea + 1.2, Hora
   PrinterVariables PCol, PosLinea + 1.6, Cuenta
   PrinterVariables PCol, PosLinea + 2, TipoP & ": " & Format$(Valor, "#,##0.00")
   PrinterVariables PCol, PosLinea + 2.4, Beneficiario
   PrinterVariables PCol, PosLinea + 2.8, DireccionCli
   If ChequeS_N Then PrinterVariables PCol, PosLinea + 2.4, NBanco & ", Cheque No. " & NCheq
   Printer.FontBold = False
   MensajeEncabData = ""
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

Public Function Cuentas_del_Prestamo(CodigoPrestamo As String) As Cuentas_Prestamos
Dim Cta_Pres As Cuentas_Prestamos
Dim AdoReg As ADODB.Recordset
'Obtenemos las cuentas del prestamo
 If CodigoPrestamo = "" Then CodigoPrestamo = Ninguno
 Set AdoReg = New ADODB.Recordset
 AdoReg.CursorType = adOpenStatic
 AdoReg.CursorLocation = adUseClient
 sSQL = "SELECT * " _
        & "FROM Catalogo_Prestamo " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND Periodo = '" & Periodo_Contable & "' " _
        & "AND TC <> " & Val(adFalse) & " " _
        & "AND CTP = '" & CodigoPrestamo & "' "
 AdoReg.open sSQL, AdoStrCnn, , , adCmdText
 If AdoReg.RecordCount > 0 Then
   'Prestamos Vigentes
    Cta_Pres.Cta_P_1_30 = AdoReg.fields("Cta_P_1_30")
    Cta_Pres.Cta_P_31_90 = AdoReg.fields("Cta_P_31_90")
    Cta_Pres.Cta_P_91_180 = AdoReg.fields("Cta_P_91_180")
    Cta_Pres.Cta_P_181_360 = AdoReg.fields("Cta_P_181_360")
    Cta_Pres.Cta_P_Mas_360 = AdoReg.fields("Cta_P_Mas_360")
   'Prestamos Vencidos
    Cta_Pres.Cta_P_1_30 = AdoReg.fields("Cta_P_1_30")
    Cta_Pres.Cta_P_31_90 = AdoReg.fields("Cta_P_31_90")
    Cta_Pres.Cta_P_91_180 = AdoReg.fields("Cta_P_91_180")
    Cta_Pres.Cta_P_181_360 = AdoReg.fields("Cta_P_181_360")
    Cta_Pres.Cta_P_Mas_360 = AdoReg.fields("Cta_P_Mas_360")
   'Otroas cuentas
    Cta_Pres.Cta_Int_Mora = AdoReg.fields("Cta_Int_Mora")
    Cta_Pres.Cta_Gas_Oper = AdoReg.fields("Cta_Gas_Oper")
    Cta_Pres.Cta_Seg_Desg_C = AdoReg.fields("Cta_Comision")
    Cta_Pres.Cta_Seg_Desg_P = AdoReg.fields("Cta_Com_Efec")
   'Totales en cero
    Cta_Pres.Total_1_30 = 0
    Cta_Pres.Total_31_90 = 0
    Cta_Pres.Total_91_180 = 0
    Cta_Pres.Total_181_360 = 0
    Cta_Pres.Total_Mas_360 = 0
 End If
 AdoReg.Close
 Cuentas_del_Prestamo = Cta_Pres
End Function

Public Sub Imprimir_Recibo_PV(NumFact As Long, _
                              DtaFactura As Adodc, _
                              DtaDetalle As Adodc, _
                              TipoFact As String)
Dim CadenaMoneda As String
Dim Numero_Letras As String
Dim Cant_Ln As Byte
Dim CantGuion As Byte

On Error GoTo Errorhandler
Mensajes = "Imprmir Factura No. " & NumFact
Titulo = "IMPRESION"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
   Escala_Centimetro 1, TipoTerminal, 9
   RatonReloj
   CantGuion = CByte(Leer_Campo_Empresa("Cant_Ancho_PV"))
   If CantGuion < 25 Then CantGuion = 25
   Total = 0: Total_IVA = 0
   Cant_Ln = 0
   PosLinea = 0.1
   Producto = ""
   If TipoFact = "PV" Then
      sSQL = "SELECT F.*,C.Cliente,C.CI_RUC,C.Telefono,C.Direccion,C.Ciudad,C.Grupo,C.Email " _
           & "FROM Trans_Ticket As F,Clientes As C " _
           & "WHERE F.Ticket = " & NumFact & " " _
           & "AND C.Codigo = F.CodigoC " _
           & "AND F.TC = '" & TipoFact & "' " _
           & "AND F.Periodo = '" & Periodo_Contable & "' " _
           & "AND F.Item = '" & NumEmpresa & "' "
   Else
      sSQL = "SELECT F.*,C.Cliente,C.CI_RUC,C.Telefono,C.Direccion,C.Ciudad,C.Grupo,C.Email " _
           & "FROM Facturas As F,Clientes As C " _
           & "WHERE F.Factura = " & NumFact & " " _
           & "AND C.Codigo = F.CodigoC " _
           & "AND F.TC = '" & TipoFact & "' " _
           & "AND F.Periodo = '" & Periodo_Contable & "' " _
           & "AND F.Item = '" & NumEmpresa & "' "
   End If
   Select_Adodc DtaFactura, sSQL
  'Iniciamos la consulta de impresion
  With DtaFactura.Recordset
   If .RecordCount > 0 Then
      'Encabezado de la Factura
       If Encabezado_PV Then
          Producto = " " & vbCrLf _
                   & Space((CantGuion - Len(Empresa)) / 2) & UCaseStrg(Empresa) & vbCrLf _
                   & Space((CantGuion - Len(NombreComercial)) / 2) & NombreComercial & vbCrLf _
                   & Space((CantGuion - Len(UCaseStrg(NombreGerente))) / 2) & UCaseStrg(NombreGerente) & vbCrLf _
                   & Space((CantGuion - Len("R.U.C. " & RUC)) / 2) & "R.U.C. " & RUC & vbCrLf _
                   & Space((CantGuion - Len("Telefono: " & Telefono1)) / 2) & "Telefono: " & Telefono1 & vbCrLf _
                   & Direccion & vbCrLf
          Cant_Ln = Cant_Ln + 7
          If TipoFact = "PV" Then
             Producto = Producto & " " & vbCrLf & "T I C K E T   No. 000-000-" & Format$(NumFact, "0000000") & vbCrLf & " " & vbCrLf
             Cant_Ln = Cant_Ln + 1
          ElseIf TipoFact = "NV" Then
             Producto = Producto & "Auto. SRI: " & Autorizacion & " - Caduca: " & MidStrg(UCaseStrg(MesesLetras(Month(Fecha_Vence))), 1, 3) & "/" & Year(Fecha_Vence) & vbCrLf & " " & vbCrLf _
                      & "NOTA DE VENTA No. " & SerieFactura & "-" & Format$(NumFact, "0000000") & vbCrLf & " " & vbCrLf
             Cant_Ln = Cant_Ln + 2
          Else
             Producto = Producto & "Auto. SRI: " & Autorizacion & " - Caduca: " & MidStrg(UCaseStrg(MesesLetras(Month(Fecha_Vence))), 1, 3) & "/" & Year(Fecha_Vence) & vbCrLf & " " & vbCrLf _
                      & "FACTURA No. " & SerieFactura & "-" & Format$(NumFact, "0000000") & vbCrLf & " " & vbCrLf
             Cant_Ln = Cant_Ln + 2
          End If
       Else
          Producto = vbCrLf & "` " & vbCrLf & "` " & vbCrLf & "` " & vbCrLf & "` " & vbCrLf _
                   & "Transaccion(" & TipoFact & ") No." & Format$(NumFact, "0000000") & vbCrLf & " " & vbCrLf
          Cant_Ln = Cant_Ln + 4
       End If
       Producto = Producto & "Fecha: " & FechaSistema & "         Hora: " & .fields("Hora") & vbCrLf
       Producto = Producto & "Cliente: " & vbCrLf _
                & "  " & MidStrg(.fields("Cliente"), 1, 30) & vbCrLf
       Producto = Producto & "R.U.C.: " & .fields("CI_RUC") & Space(16 - Len(.fields("CI_RUC"))) & vbCrLf _
                & "Cajero: " & MidStrg(CodigoUsuario, 1, 6) & vbCrLf
       Producto = Producto & String(CantGuion, "-") & vbCrLf _
                & "PRODUCTO/Cant x PVP/TOTAL" & vbCrLf _
                & String(CantGuion, "-") & vbCrLf
                Efectivo = .fields("Efectivo")
       Cant_Ln = Cant_Ln + 6
   End If
  End With
 'Comenzamos a recoger los detalles de la factura
  If TipoFact = "PV" Then
     sSQL = "SELECT DF.*,CP.Detalle,CP.Codigo_Barra " _
          & "FROM Trans_Ticket As DF,Catalogo_Productos As CP " _
          & "WHERE DF.Ticket = " & NumFact & " " _
          & "AND DF.TC = '" & TipoFact & "' " _
          & "AND DF.Item = '" & NumEmpresa & "' " _
          & "AND DF.Periodo = '" & Periodo_Contable & "' " _
          & "AND DF.Item = CP.Item " _
          & "AND DF.Periodo = CP.Periodo " _
          & "AND DF.Codigo_Inv = CP.Codigo_Inv " _
          & "ORDER BY DF.D_No "
  Else
      sSQL = "SELECT DF.*,CP.Detalle,CP.Codigo_Barra " _
           & "FROM Detalle_Factura As DF,Catalogo_Productos As CP " _
           & "WHERE DF.Factura = " & NumFact & " " _
           & "AND DF.TC = '" & TipoFact & "' " _
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
          Producto = Producto & SetearBlancos(.fields("Producto"), 25, 0, False) & vbCrLf _
                   & "Cant.=" & SetearBlancos(CStr(.fields("Cantidad")) & "x" & Format$(.fields("Precio"), "#,##0.00"), 10, 0, False) & " " _
                   & SetearBlancos(CStr(.fields("Total")), 8, 0, True, , True) & vbCrLf
          Total = Total + .fields("Total")
          If TipoFact <> "PV" Then Total_IVA = Total_IVA + .fields("Total_IVA")
          Cant_Ln = Cant_Ln + 1
         .MoveNext
       Loop
   End If
  End With
 'Pie de factura
 '===========================================================
  With DtaFactura.Recordset
   If .RecordCount > 0 Then
       SubTotal = Total
       Total = SubTotal + Total_IVA
       Producto = Producto & String(CantGuion, "-") & vbCrLf
       Cant_Ln = Cant_Ln + 1
       'If Total_IVA Then
          Producto = Producto _
                   & "    SUBTOTAL " & SetearBlancos(CStr(Total), 12, 0, True, False, True) & vbCrLf _
                   & "   I.V.A " & Porc_IVA * 100 & "% " & SetearBlancos(CStr(Total_IVA), 12, 0, True, False, True) & vbCrLf _

          Cant_Ln = Cant_Ln + 1
       'End If
       If TipoFact = "PV" Then
          Producto = Producto & "TOTAL TICKET "
       ElseIf TipoFact = "NV" Then
          Producto = Producto & "TOTAL NOTA V."
       Else
          Producto = Producto & "TOTAL FACTURA"
       End If
       Producto = Producto & SetearBlancos(CStr(Total + Total_IVA), 12, 0, True, False, True) & vbCrLf _
                & "    EFECTIVO " & SetearBlancos(CStr(Efectivo), 12, 0, True, False, True) & vbCrLf _
                & "      CAMBIO " & SetearBlancos(CStr(Efectivo - (Total + Total_IVA)), 12, 0, True, False, True) & vbCrLf
       If TipoFact <> "PV" Then
          Producto = Producto & "ORIGINAL: CLIENTE" & vbCrLf _
                              & "COPIA   : EMISOR" & vbCrLf
          If .fields("Cotizacion") > 0 Then Producto = Producto & "COTIZACION: " & Format$(.fields("Cotizacion"), "#,##0.00") & vbCrLf
       End If
       Producto = Producto & String(CantGuion, "=") & vbCrLf
       If TipoFact = "PV" Then Producto = Producto & "RECLAME SU FACTURA EN CAJA" & vbCrLf
       Producto = Producto & "  GRACIAS POR SU COMPRA " & vbCrLf & "` " & vbCrLf _
                & "` " & vbCrLf & "` " & vbCrLf & "` " & vbCrLf
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

Public Sub Mayorizar_Libretas()
End Sub
