Attribute VB_Name = "MSitios"
Global Cod_Prov As Integer
Global Cod_Sector As Integer
Global Cod_PaisR As Integer
Global Cod_PaisB As Integer
Global Cod_PaisS As Integer
Global Cod_PaisP As Integer
Global Cod_CiudR As Integer
Global Cod_CiudB As Integer
Global Cod_CiudS As Integer
Global Cod_CiudP As Integer
Global Cod_ProvR As Integer
Global Cod_ProvB As Integer
Global Cod_ProvS As Integer
Global Cod_ProvP As Integer
Global Cod_Benef As Long
Global Cod_Remit As Long
Global Cod_Sucur As Long
Global Cod_PaisC As Integer
Global Cod_CiudC As Integer
Global Cod_SucurC As Long
Global Cod_BenefC As Long
Global Remit As Boolean
Global Benef As Boolean
Global DirCompleta As String

Public Sub SetearSitios(DtaEnvio As Data, NoEnvio As Long)
     sSQL = "SELECT Co.*,"
     sSQL = sSQL & "Be.Propietario,"
     sSQL = sSQL & "Be.Empresa,"
     sSQL = sSQL & "Be.Telefono,"
     sSQL = sSQL & "Ci.Ciudad As CiudadB,"
     sSQL = sSQL & "Ci.Provincia As ProvinciaB,"
     sSQL = sSQL & "Ci.Pais As PaisB "
     sSQL = sSQL & "FROM Correos As Co,"
     sSQL = sSQL & "Beneficiarios As Be,"
     sSQL = sSQL & "Remitentes As Re,"
     sSQL = sSQL & "Sucursales As Su, "
     sSQL = sSQL & "Ciudades As Ci, "
     sSQL = sSQL & "Paises As Pa "
     sSQL = sSQL & "WHERE Co.Envio_No = " & NoEnvio & " "
     sSQL = sSQL & "AND Co.Cod_Benef = Be.Cod_Benef "
     sSQL = sSQL & "AND Co.Cod_Remit = Re.Cod_Remit "
     sSQL = sSQL & "AND Co.Cod_Sucur = Su.Cod_Sucursal "
     sSQL = sSQL & "AND Co.Cod_Ciudad = Ci.Cod_Ciudad "
     sSQL = sSQL & "AND Co.Cod_Pais = Ci.Cod_Pais "
     sSQL = sSQL & "AND Co.Cod_Pais = Pa.Cod_Pais "
     SelectData DtaEnvio, sSQL, False
End Sub

Public Sub ImprimirSitios(Datas As Data, SizeLetra As Single, Decimales As Boolean)
Dim TipoM As Boolean
Dim SimbM As String
On Error GoTo Errorhandler
RatonReloj
InicioX = 0.5: InicioY = 0
'EscalaCentimetro 1, TipoTimes, SizeLetra
DataAnchoCampos InicioX, Datas, SizeLetra, TipoTimes, 1
Ancho(0) = 0.5  ' ME
Ancho(1) = 1.2  ' Fecha
Ancho(2) = 2.8  ' Envio_No
Ancho(3) = 4.3  ' Remitente
Ancho(4) = 7.3  ' Beneficiario
Ancho(5) = 10.3 ' Telefono
Ancho(6) = 12.3 ' Ciudad
Ancho(7) = 14   ' Direccion
Ancho(8) = 15   ' Mensaje
Ancho(9) = 18.2 ' Cantidad
Ancho(10) = 20  '
Pagina = 1
Cantidad = 0
'Iniciamos la impresion
Printer.FontBold = False
With Datas.Recordset
     .MoveFirst
      'EncabezadoData Datas
      Encabezado Ancho(0), Ancho(10)
      Printer.FontSize = 14
      PrinterTexto CentrarTexto(SQLMsg1), PosLinea, SQLMsg1
      PosLinea = PosLinea + 0.8
      Printer.FontSize = SizeLetra
      Printer.Line (InicioX, PosLinea)-(Ancho(10), PosLinea), QBColor(0)
      PosLinea = PosLinea + 0.1
      Printer.FontSize = SizeLetra
      TipoM = .Fields("ME")
      If TipoM Then SimbM = "US$" Else SimbM = Moneda
      Do While Not .EOF
         If TipoM <> .Fields("ME") Then
            Printer.Line (InicioX, PosLinea)-(Ancho(10), PosLinea), QBColor(0)
            PosLinea = PosLinea + 0.1
            PrinterTexto Ancho(9) - 1.8, PosLinea, "Total    " & SimbM
            PrinterVariables Ancho(9) - 0.1, PosLinea, Total
            Total = 0
            TipoM = .Fields("ME")
            If TipoM Then SimbM = "US$" Else SimbM = Moneda
            PosLinea = PosLinea + 0.6
            Printer.Line (InicioX, PosLinea)-(Ancho(10), PosLinea), QBColor(0)
            PosLinea = PosLinea + 0.1
         End If
         Printer.FontBold = True
         PrinterTexto 0.6, PosLinea, "Desde:"
         PrinterTexto 0.6, PosLinea + 0.4, "Hasta:"
         PrinterTexto 0.6, PosLinea + 0.8, "Contrato No."
         PrinterTexto 3.4, PosLinea, "Estado:"
         PrinterTexto 3.4, PosLinea + 0.4, "Propietario:"
         If .Fields("T") = "P" Then
             PrinterTexto 10.5, PosLinea, "Desde:"
             PrinterTexto 13, PosLinea, "Hasta:"
             PrinterTexto 16.5, PosLinea, "Comp. No."
         End If
         PrinterTexto 10.5, PosLinea + 0.4, "Teléfono:"
         PrinterTexto 10.5, PosLinea + 0.8, "Dimension:"
         PrinterTexto 10.5, PosLinea + 1.2, "Observacion:"
         PrinterTexto 3.4, PosLinea + 1.6, "Ciudad:"
         PrinterTexto 3.4, PosLinea + 0.8, "Empresa:"
         PrinterTexto 3.4, PosLinea + 1.2, "Dirección:"
         PrinterTexto 3.4, PosLinea + 2, "NOTA:"
         PrinterTexto Ancho(9) - 1, PosLinea + 1.6, SimbM
         Printer.FontBold = False
         PrinterFields 1.9, PosLinea, .Fields("Desde"), False
         PrinterFields 12, PosLinea + 1.2, .Fields("Observaciones"), False
         If .Fields("T") = "P" Then
             PrinterFields 11.5, PosLinea, .Fields("CDesde"), False
             PrinterFields 14.5, PosLinea, .Fields("CHasta"), False
             PrinterFields 18, PosLinea, .Fields("Comp_No"), False
         End If
         Select Case .Fields("T")
           Case "P": PrinterTexto 5, PosLinea, "Cancelado"
           Case "N": PrinterTexto 5, PosLinea, "Por Pagar"
         End Select
         PrinterFields 1.9, PosLinea + 0.4, .Fields("Hasta"), False
         PrinterFields 5, PosLinea + 0.4, .Fields("Propietario"), False
         PrinterFields 12, PosLinea + 0.4, .Fields("Telefono"), False
         PrinterFields 2.1, PosLinea + 0.8, .Fields("Contrato_No"), False
         PrinterFields 12, PosLinea + 0.8, .Fields("Dimensiones"), False
         PrinterFields 5, PosLinea + 0.8, .Fields("Empresa"), False
         DirCompleta = .Fields("Direccion1")
         If .Fields("Direccion2") <> Ninguno Then
             DirCompleta = DirCompleta & " y " & .Fields("Direccion2")
         End If
         PrinterTexto 5, PosLinea + 1.2, DirCompleta
         DirCompleta = .Fields("Ciudad") & " - " & .Fields("Provincia") & " - " & .Fields("Pais")
         PrinterTexto 5, PosLinea + 1.6, DirCompleta
         Printer.FontBold = True
         PrinterFields Ancho(9) - 0.3, PosLinea + 1.6, .Fields("TOTAL"), False
         Printer.Line (Ancho(0), PosLinea - 0.1)-(Ancho(0), PosLinea + 2.5), QBColor(0)
         Printer.Line (Ancho(10), PosLinea - 0.1)-(Ancho(10), PosLinea + 2.5), QBColor(0)
         Total = Total + .Fields("TOTAL")
         PosLinea = PosLinea + 2.5
         Printer.Line (InicioX, PosLinea)-(Ancho(10), PosLinea), QBColor(0)
         PosLinea = PosLinea + 0.1
         If PosLinea > LimiteAlto - 2 Then
            Printer.Line (Ancho(0), PosLinea)-(Ancho(10), PosLinea), QBColor(0)
            Printer.NewPage
            'EncabezadoData Datas
            PosLinea = 0
            Encabezado Ancho(0), Ancho(10)
            Printer.FontSize = 14
            PrinterTexto CentrarTexto(SQLMsg1), PosLinea, SQLMsg1
            PosLinea = PosLinea + 0.8
            Printer.FontSize = SizeLetra
            Printer.Line (InicioX, PosLinea)-(Ancho(10), PosLinea), QBColor(0)
            PosLinea = PosLinea + 0.1
            Printer.FontSize = SizeLetra
         End If
        .MoveNext
      Loop
End With
PosLinea = PosLinea + 0.4
Printer.Line (InicioX, PosLinea)-(Ancho(10), PosLinea), QBColor(0)
PosLinea = PosLinea + 0.1
PrinterTexto Ancho(9) - 1.8, PosLinea, "Total    " & SimbM
PrinterVariables Ancho(9) - 0.3, PosLinea, Total
UltimaLinea = PosLinea + 0.5
RatonNormal
Printer.EndDoc
Exit Sub
Errorhandler:
    RatonNormal
    ErrorDeImpresion
    Exit Sub
End Sub

