Attribute VB_Name = "MEnvios"

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

Public Sub ImprimirCorreos(DtaEnvio As Data)
Dim ConCopia As Boolean
Dim SM As String
Dim SMN As String
Dim NumeroLineas As Byte
'Establecemos Espacios y seteos de impresion
On Error GoTo Errorhandler
ConCopia = False
Mensajes = "Desea imprimir el Envio No. " '& DtaEnvio.Recordset.Fields("Envio_No")
Titulo = "Formulario de Impresion."
If BoxMensaje = 6 Then
Mensajes = "Imprimir Copia"
Titulo = "Formulario de Impresion"
If BoxMensaje = 6 Then ConCopia = True
RatonReloj
LetraAnterior = Printer.FontName
EscalaCentimetro 1, TipoTimes, 10
'Printer.Copies = 2
Volver_Imp:
Printer.FontName = TipoTimes
Saldo = 0
With DtaEnvio.Recordset
'If .Fields("ME") Then
 '    SM = "US$"
 '   SMN = Moneda
 '   Saldo = .Fields("Cambio") * .Fields("Cantidad")
'Else
'    SM = Moneda
'    SMN = "US$"
'    Saldo = .Fields("Cantidad") / .Fields("Cambio")
'End If
Printer.Line (0.5, 0.1)-(19.5, 10.5), QBColor(0), B
Printer.Line (0.6, 0.2)-(19.4, 10.4), QBColor(0), B
Printer.Line (0.6, 7.3)-(19.4, 7.3), QBColor(0)
Printer.Line (14, 9.5)-(17.5, 9.5), QBColor(0)
'Iniciamos la consulta de impresion
Printer.FontSize = 22: Printer.FontBold = True
Printer.FontItalic = True
Cadena = "ECUAENVIOS"
PrinterTexto CentrarTexto(Cadena), 0.5, Cadena
Printer.FontSize = 16
'If .Fields("T") <> Pendiente Then
'    Cadena = "Comprobante de Pago"
'    PrinterTexto CentrarTexto(Cadena), 1.5, Cadena
'End If
'PrinterTexto 14.5, 1.3, "Envio No.  " & Format(.Fields("Envio_No"), "000000")
Printer.FontSize = 10
PrinterTexto 10, 11.5, "Firma Cliente:_______________________"
'PrinterTexto 12.2, 12, "Fecha: " & .Fields("Fecha_RP")
Printer.FontUnderline = True
PrinterTexto 16, 2, "Fecha:"
PrinterTexto 16, 2.5, "Hora:"
Printer.FontSize = 12
PrinterTexto 1, 4.6, "REMITENTE"
PrinterTexto 10, 4.6, "BENEFICIARIO"
PrinterTexto 1, 7.5, "MENSAJE: "
Printer.FontUnderline = False
PrinterTexto 10, 7.5, "CANTIDAD"
PrinterTexto 10, 8, "TELEFAX"
PrinterTexto 10, 8.5, "URGENCIA"
PrinterTexto 10, 9, "% ENVIO"
PrinterTexto 10, 9.7, "T O T A L"
PrinterTexto 12.6, 7.5, SM
PrinterTexto 12.6, 8, SM
PrinterTexto 12.6, 8.5, SM
PrinterTexto 12.6, 9, SM
PrinterTexto 12.6, 9.7, SM
PrinterTexto 1, 2.3, Empresa
Printer.FontBold = False: Printer.FontItalic = False
Printer.FontSize = 9
PrinterTexto 1, 2.8, "Teléfono: " & Telefono1
PrinterLineas 1, 3.2, "Dir: " & Direccion, 8
PrinterTexto 1, 4.1, NombreCiudad & " - " & NombrePais
'PrinterTexto 17.2, 2, .Fields("Fecha")
PrinterTexto 17.2, 2.5, Str(Time)

'PrinterTexto 10, 2.8, .Fields("TelefonoS")
'PrinterLineas 10, 3.2, .Fields("DireccionS"), 8.5
PrinterTexto 10, 4.1, .Fields("CiudadS") & " - " & .Fields("PaisS")

Printer.FontSize = 12
PrinterTexto 1, 5.1, .Fields("Remitente")
PrinterTexto 1, 5.6, .Fields("TelefonoR")
NumeroLineas = PrinterLineasMayor(1, 6.1, .Fields("DireccionR"), 8.5)
'PrinterLineas 1, 6.1, .Fields("DireccionR"), 8.5

PrinterTexto 10, 5.1, .Fields("Beneficiario")
PrinterTexto 10, 5.6, .Fields("TelefonoB")
'PrinterLineas 10, 6.1, .Fields("DireccionB"), 8.5
NumeroLineas = PrinterLineasMayor(10, 6.1, .Fields("Direccionb"), 8.5)

PrinterVariables 14.8, 7.5, .Fields("Cantidad"), 2 'cantidad
PrinterVariables 14.8, 8, .Fields("Telefax"), 2  'telefax
PrinterVariables 14.8, 8.5, .Fields("Urgencia"), 2 'urgencia
PrinterVariables 14.8, 9, .Fields("Porc_Envio"), 2   'Porc_Envio
PrinterVariables 14, 9.7, .Fields("TOTAL"), 2 'Total
NumeroLineas = PrinterLineasMayor(1, 8, .Fields("Mensaje"), 8)
'PrinterLineas 1, 8, .Fields("Mensaje"), 8   ' Mensaje
If Saldo > 0 Then
   Cadena = "Cambio: " & .Fields("Cambio") & ", "
   Cadena = Cadena & SMN & " "
   Cadena = Cadena & Format(Saldo, "#,##0.00")
   PrinterTexto 1, 9.7, Cadena
End If
End With
PrinterPaint LogoTipo, 0.7, 0.2, 3, 2
If ConCopia Then
   Printer.NewPage
   ConCopia = False
   GoTo Volver_Imp
End If
Printer.FontName = LetraAnterior
Printer.EndDoc  ' La impresión ha terminado.
RatonNormal
End If
Exit Sub
Errorhandler:
    RatonNormal
    MsgBox "Error: Hubo un problema al imprimir en su impresora."
    Exit Sub
End Sub
'--------------------------------
Public Sub SetearEnvios(DtaEnvio As Data, NoEnvio As Long)
     sSQL = "SELECT Co.*,"
     sSQL = sSQL & "Be.Beneficiario,"
     sSQL = sSQL & "Be.Direccion As DireccionB,"
     sSQL = sSQL & "Be.Telefono As TelefonoB,"
     sSQL = sSQL & "Re.Remitente,"
     sSQL = sSQL & "Re.Direccion As DireccionR,"
     sSQL = sSQL & "Re.Telefono As TelefonoR,"
     sSQL = sSQL & "Su.Sucursal,"
     sSQL = sSQL & "Su.Direccion As DireccionS,"
     sSQL = sSQL & "Su.Telefono As TelefonoS,"
     sSQL = sSQL & "Ci.Ciudad As CiudadS,"
     sSQL = sSQL & "Pa.Pais As PaisS "
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

Public Sub ImprimirEnvios(Datas As Data, SizeLetra As Single, Decimales As Boolean)
Dim TipoM As Boolean
Dim SimbM As String
On Error GoTo Errorhandler
RatonReloj
InicioX = 0.5: InicioY = 0
'EscalaCentimetro 1, TipoTimes, SizeLetra
DataAnchoCampos InicioX, Datas, Decimales, SizeLetra, TipoTimes, 1
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
            PrinterTexto Ancho(9) - 2, PosLinea, "Total    " & SimbM
            PrinterVariables Ancho(9), PosLinea, Cantidad, True
            Cantidad = 0
            TipoM = .Fields("ME")
            If TipoM Then SimbM = "US$" Else SimbM = Moneda
            PosLinea = PosLinea + 0.6
            Printer.Line (InicioX, PosLinea)-(Ancho(10), PosLinea), QBColor(0)
            PosLinea = PosLinea + 0.1
         End If
         Printer.FontBold = True
         PrinterTexto 0.6, PosLinea, "Fecha:"
         PrinterTexto 0.6, PosLinea + 0.4, "Envio No."
         PrinterTexto 0.6, PosLinea + 0.8, "Cambio:"
         If .Fields("ME") = False Then PrinterTexto 0.6, PosLinea + 1.2, "US$"
         PrinterTexto 3.4, PosLinea, "Remitente:"
         PrinterTexto 3.4, PosLinea + 0.4, "Beneficiario:"
         PrinterTexto 10.5, PosLinea, "Estado:"
         PrinterTexto 10.5, PosLinea + 0.4, "Teléfono:"
         Select Case .Fields("T")
           Case "C", "R": PrinterTexto 14.5, PosLinea, "Fecha:"
         End Select
         PrinterTexto 14.5, PosLinea + 0.4, "Ciudad:"
         PrinterTexto 3.4, PosLinea + 0.8, "Dirección:"
         PrinterTexto 3.4, PosLinea + 1.2, "Mensaje:"
         PrinterTexto Ancho(9) - 1, PosLinea + 1.6, SimbM
         Printer.FontBold = False
         PrinterFields 1.8, PosLinea, .Fields("Fecha"), True, False
         PrinterFields 1.8, PosLinea + 0.4, .Fields("Envio_No"), True, False
         PrinterFields 1.6, PosLinea + 0.8, .Fields("Cambio"), True, False
         If .Fields("ME") = False Then
             If .Fields("Cambio") <> 0 Then
                 Total = .Fields("Cantidad") / .Fields("Cambio")
             Else
                 Total = 0
             End If
             PrinterVariables 1.1, PosLinea + 1.2, Total, True
         End If
         PrinterFields 5, PosLinea, .Fields("Remitente"), True, False
         PrinterFields 5, PosLinea + 0.4, .Fields("Beneficiario"), True, False
         PrinterFields 11.7, PosLinea + 0.4, .Fields("Telefono"), True, False
         Select Case .Fields("T")
           Case "C": PrinterTexto 11.7, PosLinea, "Pagado"
           Case "R": PrinterTexto 11.7, PosLinea, "Recibido"
           Case "P": PrinterTexto 11.7, PosLinea, "Pendiente"
         End Select
         Select Case .Fields("T")
           Case "C", "R": PrinterFields 15.6, PosLinea, .Fields("Fecha_RP"), True, False
         End Select
         PrinterFields 15.6, PosLinea + 0.4, .Fields("Ciudad"), True, False
         PrinterFields 5, PosLinea + 0.8, .Fields("Direccion"), True, False
         NumeroLineas = PrinterLineasMayor(5, PosLinea + 1.2, .Fields("Mensaje"), 12)
         'PrinterFields 4.6, PosLinea + 1.2, .Fields("Mensaje"), True, False
         Printer.FontBold = True
         PrinterFields Ancho(9), PosLinea + 1.6, .Fields("Cantidad"), True, False
         Printer.Line (Ancho(0), PosLinea - 0.1)-(Ancho(0), PosLinea + 2), QBColor(0)
         Printer.Line (Ancho(10), PosLinea - 0.1)-(Ancho(10), PosLinea + 2), QBColor(0)
         Cantidad = Cantidad + .Fields("Cantidad")
         PosLinea = PosLinea + 2
         Printer.Line (InicioX, PosLinea)-(Ancho(10), PosLinea), QBColor(0)
         PosLinea = PosLinea + 0.1
         If PosLinea > LimiteAlto Then
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
Printer.Line (InicioX, PosLinea)-(Ancho(10), PosLinea), QBColor(0)
PosLinea = PosLinea + 0.1
PrinterTexto Ancho(9) - 2, PosLinea, "Total    " & SimbM
PrinterVariables Ancho(9), PosLinea, Cantidad, True
UltimaLinea = PosLinea + 0.5
RatonNormal
Printer.EndDoc
Exit Sub
Errorhandler:
    RatonNormal
    ErrorDeImpresion
    Exit Sub
End Sub

Public Sub ImprimirEnvios1(Datas As Data, FormaImp As Byte, SizeLetra As Single, Decimales As Boolean)
Dim TipoM As Boolean
On Error GoTo Errorhandler
RatonReloj
InicioX = 0.5: InicioY = 0
'EscalaCentimetro FormaImp, TipoTimes, SizeLetra
DataAnchoCampos InicioX, Datas, Decimales, SizeLetra, TipoTimes, FormaImp
Ancho(0) = 0.5  ' ME
Ancho(1) = 1.5  ' Fecha
Ancho(2) = 3.5  ' Envio_No
Ancho(3) = 5.1  ' Ciudad
Ancho(4) = 8    ' Cambio
Ancho(5) = 10   ' Cantidad
Ancho(6) = 12   ' Porc_Envio
Ancho(7) = 14   ' Total
Ancho(8) = 16.5   '
Pagina = 1: CantCampos = 8
Precio = 0: Saldo = 0: Cantidad = 0
'Iniciamos la impresion
Printer.FontBold = False
With Datas.Recordset
     .MoveFirst
      EncabezadoData Datas
      Printer.FontSize = SizeLetra
      TipoM = .Fields("ME")
      Do While Not .EOF
         If TipoM <> .Fields("ME") Then
            PosLinea = PosLinea + 0.1
            Printer.Line (InicioX, PosLinea)-(Ancho(8) + 2.6, PosLinea), QBColor(0)
            PosLinea = PosLinea + 0.1
            PrinterTexto Ancho(4), PosLinea, "T O T A L E S"
            PrinterVariables Ancho(5), PosLinea, Cantidad, True
            PrinterVariables Ancho(6), PosLinea, Precio, True
            PrinterVariables Ancho(7), PosLinea, Saldo, True
            Precio = 0: Saldo = 0: Cantidad = 0
            TipoM = .Fields("ME")
            PosLinea = PosLinea + 0.6
            Printer.Line (InicioX, PosLinea)-(Ancho(8), PosLinea), QBColor(0)
            PosLinea = PosLinea + 0.1
            If PosLinea > LimiteAlto Then
               Printer.Line (Ancho(0), PosLinea)-(Ancho(CantCampos), PosLinea), QBColor(0)
               Printer.NewPage
               EncabezadoData Datas
               Printer.FontSize = SizeLetra
            End If
         End If
         PrinterAllFields CantCampos, PosLinea, Datas, Decimales, True
         If .Fields("ME") Then
             Cadena = " E "
         Else
             Cadena = " N "
             If .Fields("Cambio") <> 0 Then
                 Distancia = .Fields("Cantidad") / .Fields("Cambio")
                 PrinterTexto Ancho(CantCampos), PosLinea, "US$"
             Else
                 Distancia = 0
             End If
             PrinterVariables Ancho(CantCampos) + 0.7, PosLinea, Distancia, True
             Printer.Line (Ancho(CantCampos) + 2.6, PosLinea - 0.1)-(Ancho(CantCampos) + 2.6, PosLinea + 0.5), QBColor(0)
         End If
         PrinterTexto Ancho(0), PosLinea, Cadena
         Cantidad = Cantidad + .Fields("Cantidad")
         Precio = Precio + .Fields("Porc_Envio")
         Saldo = Saldo + .Fields("TOTAL")
         PosLinea = PosLinea + 0.4
         If PosLinea > LimiteAlto Then
            Printer.Line (Ancho(0), PosLinea)-(Ancho(CantCampos), PosLinea), QBColor(0)
            Printer.NewPage
            EncabezadoData Datas
            Printer.FontSize = SizeLetra
         End If
        .MoveNext
      Loop
End With
Printer.Line (InicioX, PosLinea)-(Ancho(CantCampos) + 2.6, PosLinea), QBColor(0)
PosLinea = PosLinea + 0.1
Printer.Line (InicioX, PosLinea)-(Ancho(CantCampos) + 2.6, PosLinea), QBColor(0)
PosLinea = PosLinea + 0.1
PrinterTexto Ancho(4), PosLinea, "T O T A L E S"
PrinterVariables Ancho(5), PosLinea, Cantidad, True
PrinterVariables Ancho(6), PosLinea, Precio, True
PrinterVariables Ancho(7), PosLinea, Saldo, True
UltimaLinea = PosLinea + 0.5
RatonNormal
Printer.EndDoc
Exit Sub
Errorhandler:
    RatonNormal
    ErrorDeImpresion
    Exit Sub
End Sub


