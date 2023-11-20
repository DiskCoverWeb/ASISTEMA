Attribute VB_Name = "SubEnvios"

Public Sub EncabezadoResumenGiros(OpcCiudad As Boolean, OpcPC As Integer)
  Printer.FontSize = 9
  Printer.FontBold = True
  Printer.Line (Ancho(0), PosLinea)-(Ancho(CantCampos), PosLinea), QBColor(0)
  PosLinea = PosLinea + 0.1
  If OpcCiudad Then
     PrinterTexto Ancho(1), PosLinea, "Corresponsal"
  Else
     PrinterTexto Ancho(1), PosLinea, "Ciudad"
  End If
  PrinterTexto Ancho(2), PosLinea, "Fecha"
  PrinterTexto Ancho(3), PosLinea, "Cant_Recib"
  PrinterTexto Ancho(4), PosLinea, "Recibido"
  PrinterTexto Ancho(5), PosLinea, "Cant_Cance"
  PrinterTexto Ancho(6), PosLinea, "Cancelado"
  If OpcPC = 1 Then
     PrinterTexto Ancho(7), PosLinea, "Promedio"
     PrinterTexto Ancho(8), PosLinea, "Comision"
  End If
  PosLinea = PosLinea + 0.5
  Printer.Line (Ancho(0), PosLinea)-(Ancho(CantCampos), PosLinea), QBColor(0)
  PosLinea = PosLinea + 0.1
  Printer.FontBold = False
End Sub

Public Sub EncabezadoDataEnvio(Datas As Adodc, CantCampos)
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
   PosLinea = PosLinea + 0.6
End If
Printer.FontSize = 9
Printer.FontBold = True
'========================================================================
PrinterTexto Ancho(0), PosLinea, "T"
PrinterTexto Ancho(1), PosLinea, "Envio No."
PrinterTexto Ancho(2), PosLinea, "Fecha E."
PrinterTexto Ancho(3), PosLinea, "Fecha P."
PrinterTexto Ancho(4), PosLinea, "Beneficiario"
PrinterTexto Ancho(5), PosLinea, "Telefono"
PrinterTexto Ancho(6), PosLinea, "Ciudad"
PrinterTexto Ancho(7), PosLinea, "TOTAL"
PrinterTexto Ancho(8), PosLinea, "SALDO"
PosLinea = PosLinea + 0.4
Printer.Line (InicioX, PosLinea)-(Ancho(CantCampos), PosLinea), QBColor(0)
PosLinea = PosLinea + 0.1
Printer.FontBold = False
Printer.FontSize = PorteLetra
Printer.FontName = LetraAnterior
End Sub

Public Sub EncabezadoDataEnvio1(Datas As Adodc, CantCampos)
Dim InicX As Single
Dim InicY As Single
'Encabezado Ancho(0), Ancho(CantCampos)
PorteLetra = Printer.FontSize
LetraAnterior = Printer.FontName
Printer.FontName = TipoTimes
Printer.FontBold = True
PosLinea = 0.5
Printer.FontSize = 16
PrinterTexto CentrarTexto(Empresa), PosLinea, Empresa
PosLinea = PosLinea + 0.8
Printer.FontSize = 9
PrinterTexto Ancho(0), PosLinea, "Pagina No. " & Pagina
Pagina = Pagina + 1
PrinterTexto CentrarTexto(Direccion & ", Teléfono: " & Telefono1), PosLinea, Direccion & ", Teléfono: " & Telefono1
PosLinea = PosLinea + 0.6
'If SQLMsg1 <> "" Then
'   Printer.FontSize = 12
'   PrinterTexto CentrarTexto(SQLMsg1), PosLinea, SQLMsg1
'   PosLinea = PosLinea + 0.6
'End If
Printer.FontSize = 10
If SQLMsg2 <> "" Then
   PrinterTexto CentrarTexto(SQLMsg2), PosLinea, SQLMsg2
   PosLinea = PosLinea + 0.5
End If
If SQLMsg3 <> "" Then
   PrinterTexto Ancho(0), PosLinea, SQLMsg3
   PosLinea = PosLinea + 0.5
End If
Printer.FontSize = 9
Printer.FontBold = True
PosLinea = PosLinea + 0.1
PrinterTexto Ancho(0), PosLinea, "PARA:"
PrinterTexto Ancho(0) + 2, PosLinea, Codigo
PosLinea = PosLinea + 0.5
PrinterTexto Ancho(0), PosLinea, "DE:"
PrinterTexto Ancho(0) + 2, PosLinea, Codigos
PosLinea = PosLinea + 0.5
PrinterTexto Ancho(0), PosLinea, "REF."
PrinterTexto Ancho(0) + 2, PosLinea, Codigo1
PosLinea = PosLinea + 0.5
'========================================================================
PrinterTexto Ancho(0), PosLinea, "T"
PrinterTexto Ancho(1), PosLinea, "Envio No."
PrinterTexto Ancho(2), PosLinea, "Giro No."
PrinterTexto Ancho(3), PosLinea, "Fecha E."
PrinterTexto Ancho(4), PosLinea, "Fecha P."
PrinterTexto Ancho(5), PosLinea, "Beneficiario"
PrinterTexto Ancho(6), PosLinea, "RUC/CI"
PrinterTexto Ancho(7), PosLinea, "TOTAL"
PosLinea = PosLinea + 0.5
Printer.FontBold = False
Printer.FontSize = PorteLetra
Printer.FontName = LetraAnterior
End Sub

Public Sub ImprimirResumenGiros(Datas As Adodc, SizeLetra As Single, OpcCiudad As Boolean, OpcPC As Integer)
On Error GoTo Errorhandler
RatonReloj
InicioX = 0.5: InicioY = 0
If OpcPC = 1 Then SizeLetra = 9.5 Else SizeLetra = 12
'EscalaCentimetro   FormaImp, TipoTimes, SizeLetra
DataAnchoCampos InicioX, Datas, SizeLetra, TipoTimes, 1
Ancho(0) = 0.5    'CORRESPONSAL/CIUDAD
Ancho(1) = 0.6    'CIUDAD/CORRESPONSAL
If OpcPC = 1 Then
   Ancho(2) = 5      'FECHA
   Ancho(3) = 6.5    'CANT RECIB
   Ancho(4) = 8      'RECIBIDO
   Ancho(5) = 10.8   'CANT_CANC
   Ancho(6) = 12.2   'CANCELADO
   Ancho(7) = 15     'PROMEDIO
   Ancho(8) = 17     'COMISION
Else
   Ancho(2) = 6.5    'FECHA
   Ancho(3) = 8.5    'CANT RECIB
   Ancho(4) = 10.5   'RECIBIDO
   Ancho(5) = 13.5   'CANT_CANC
   Ancho(6) = 16     'CANCELADO
   Ancho(7) = 19     'PROMEDIO
   Ancho(8) = 19     'COMISION
End If
Ancho(9) = 19
Pagina = 1
'Iniciamos la impresion
Printer.FontBold = False
With Datas.Recordset
     .MoveFirst
      Total = 0: Saldo = 0: PLitro = 0: PLitroT = 0
      Total_ME = 0: Saldo_ME = 0
      Encabezado Ancho(0), Ancho(9)
      Printer.FontBold = True
      Printer.FontSize = 12
      If OpcCiudad Then
         Codigo = .Fields("Ciudad")
         Codigo2 = .Fields("Ciudad") & ": "
      Else
         Codigo = .Fields("Corresponsal")
         Codigo2 = .Fields("Corresponsal") & ": "
      End If
      PrinterTexto Ancho(0), PosLinea, Codigo2
      PosLinea = PosLinea + 0.6
      EncabezadoResumenGiros OpcCiudad, OpcPC
      Do While Not .EOF
         If Codigo <> .Fields("Corresponsal") And OpcCiudad = False Then
            Printer.FontBold = True
            PosLinea = PosLinea + 0.1
            Printer.Line (Ancho(0), PosLinea)-(Ancho(CantCampos), PosLinea), QBColor(0)
            PosLinea = PosLinea + 0.1
            PrinterTexto Ancho(2), PosLinea, "T O T A L E S"
            PrinterVariables Ancho(4) - 0.5, PosLinea, Total
            PrinterVariables Ancho(6) - 0.5, PosLinea, Saldo
            PrinterVariables Ancho(8), PosLinea, PLitro
            PosLinea = PosLinea + 0.5
            Printer.FontBold = True
            Printer.FontSize = 12
            Codigo = .Fields("Corresponsal")
            Codigo2 = .Fields("Corresponsal") & ": "
            PrinterTexto Ancho(0), PosLinea, Codigo2
            PosLinea = PosLinea + 0.6
            If PosLinea >= LimiteAlto Then
               PosLinea = PosLinea + 0.1
               Printer.Line (Ancho(0), PosLinea)-(Ancho(CantCampos), PosLinea), QBColor(0)
               Printer.NewPage
               Encabezado Ancho(0), Ancho(8)
               Printer.FontBold = True
               Printer.FontSize = 12
               PrinterTexto Ancho(0), PosLinea, Codigo2
               PosLinea = PosLinea + 0.6
               EncabezadoResumenGiros OpcCiudad, OpcPC
            Else
               EncabezadoResumenGiros OpcCiudad, OpcPC
            End If
            Total = 0: Saldo = 0: PLitro = 0
         End If
         If Codigo <> .Fields("Ciudad") And OpcCiudad = True Then
            Printer.FontBold = True
            Printer.Line (Ancho(0), PosLinea)-(Ancho(CantCampos), PosLinea), QBColor(0)
            PosLinea = PosLinea + 0.1
            PrinterTexto Ancho(2), PosLinea, "T O T A L E S"
            PrinterVariables Ancho(4) - 0.5, PosLinea, Total
            PrinterVariables Ancho(6) - 0.5, PosLinea, Saldo
            PrinterVariables Ancho(8), PosLinea, PLitro
            PosLinea = PosLinea + 0.5
            Printer.FontBold = True
            Printer.FontSize = 12
            Codigo = .Fields("Ciudad")
            Codigo2 = .Fields("Ciudad") & ": "
            PrinterTexto Ancho(0), PosLinea, Codigo2
            PosLinea = PosLinea + 0.6
            If PosLinea >= LimiteAlto Then
               PosLinea = PosLinea + 0.1
               Printer.Line (Ancho(0), PosLinea)-(Ancho(CantCampos), PosLinea), QBColor(0)
               Printer.NewPage
               Encabezado Ancho(0), Ancho(8)
               Printer.FontBold = True
               Printer.FontSize = 12
               PrinterTexto Ancho(0), PosLinea, Codigo2
               PosLinea = PosLinea + 0.6
               EncabezadoResumenGiros OpcCiudad, OpcPC
            Else
               EncabezadoResumenGiros OpcCiudad, OpcPC
            End If
            Total = 0: Saldo = 0: PLitro = 0
         End If
         Printer.FontBold = False
         Printer.FontSize = SizeLetra
         If OpcCiudad Then
            PrinterFields Ancho(1), PosLinea, .Fields("Corresponsal"), False
         Else
            PrinterFields Ancho(1), PosLinea, .Fields("Ciudad"), False
         End If
         If OpcPC = 1 Then
            PrinterFields Ancho(8), PosLinea, .Fields("Comision"), False
            PrinterFields Ancho(7), PosLinea, .Fields("Promedio"), False
         End If
         PrinterFields Ancho(6) - 0.5, PosLinea, .Fields("Cancelado"), False
         PrinterFields Ancho(5), PosLinea, .Fields("Cant_Cance"), False
         PrinterFields Ancho(4) - 0.5, PosLinea, .Fields("Recibido"), False
         PrinterFields Ancho(3), PosLinea, .Fields("Cant_Recib"), False
         PrinterFields Ancho(2), PosLinea, .Fields("Fecha"), False
         PosLinea = PosLinea + 0.45
         If PosLinea >= LimiteAlto Then
            PosLinea = PosLinea + 0.1
            Printer.Line (Ancho(0), PosLinea)-(Ancho(CantCampos), PosLinea), QBColor(0)
            Printer.NewPage
            Encabezado Ancho(0), Ancho(8)
            Printer.FontBold = True
            Printer.FontSize = 12
            PrinterTexto Ancho(0), PosLinea, Codigo2
            PosLinea = PosLinea + 0.6
            EncabezadoResumenGiros OpcCiudad, OpcPC
         End If
         Total = Total + .Fields("Recibido")
         Total_ME = Total_ME + .Fields("Recibido")
         Saldo = Saldo + .Fields("Cancelado")
         Saldo_ME = Saldo_ME + .Fields("Cancelado")
         PLitro = PLitro + .Fields("Comision")
         PLitroT = PLitroT + .Fields("Comision")
        .MoveNext
      Loop
End With
Printer.FontBold = True
PosLinea = PosLinea + 0.1
Printer.Line (Ancho(0), PosLinea)-(Ancho(CantCampos), PosLinea), QBColor(0)
PosLinea = PosLinea + 0.1
PrinterTexto Ancho(2), PosLinea, "T O T A L E S"
PrinterVariables Ancho(4) - 0.5, PosLinea, Total
PrinterVariables Ancho(6) - 0.5, PosLinea, Saldo
PrinterVariables Ancho(8), PosLinea, PLitro
PosLinea = PosLinea + 0.5
Printer.Line (InicioX, PosLinea)-(Ancho(CantCampos), PosLinea), QBColor(0)
PosLinea = PosLinea + 0.1
Printer.Line (InicioX, PosLinea)-(Ancho(CantCampos), PosLinea), QBColor(0)
PosLinea = PosLinea + 0.1
PrinterTexto Ancho(2), PosLinea, "T O T A L E S"
PrinterVariables Ancho(4) - 0.5, PosLinea, Total_ME
PrinterVariables Ancho(6) - 0.5, PosLinea, Saldo_ME
PrinterVariables Ancho(8), PosLinea, PLitroT
RatonNormal
Printer.EndDoc
Exit Sub
Errorhandler:
    RatonNormal
    ErrorDeImpresion
    Exit Sub
End Sub

Public Sub ImprimirEnvio(DtaEnvio As Adodc, Optional Factura As Boolean)
Dim ConCopia As Boolean
Dim SM As String
Dim SMN As String
Dim NumeroLineas As Byte
'Establecemos Espacios y seteos de impresion
On Error GoTo Errorhandler
ConCopia = False
Mensajes = "Desea imprimir el Envio No. " & DtaEnvio.Recordset.Fields("Envio_No")
Titulo = "Formulario de Impresion."
If BoxMensaje = 6 Then
'Mensajes = "Imprimir Copia"
'Titulo = "Formulario de Impresion"
'If BoxMensaje = 6 Then ConCopia = True
RatonReloj
LetraAnterior = Printer.FontName
EscalaCentimetro 1, TipoTimes, 10
'Printer.Copies = 2
Volver_Imp:
Printer.FontName = TipoTimes
Saldo = 0
With DtaEnvio.Recordset
 If .Fields("ME") Then
     SM = "US$"
     SMN = Moneda
     Saldo = .Fields("SALDO")
 Else
    SM = Moneda
    SMN = "US$"
    If .Fields("Cotizacion") > 0 Then
        Saldo = .Fields("Cantidad") / .Fields("Cotizacion")
    Else
        Saldo = 0
    End If
End If
Printer.FontBold = False
Printer.Line (0.6, 9.3)-(19.4, 9.3), QBColor(0)
'Iniciamos la consulta de impresion
Printer.FontSize = 20: Printer.FontItalic = True
PrinterTexto CentrarTexto(Empresa), 0.2, Empresa
Printer.FontSize = 14
Cadena = "R.U.C. " & RUC
PrinterTexto CentrarTexto(Cadena), 1, Cadena
If Factura Then
   Cadena = "RECIBO DE TRANSFERENCIA"
   PrinterTexto CentrarTexto(Cadena), 2, Cadena
   Cadena = "DE MONEYGRAM"
   PrinterTexto CentrarTexto(Cadena), 2.6, Cadena
Else
   Cadena = "COMPROBANTE DE ENTREGA"
   PrinterTexto CentrarTexto(Cadena), 2, Cadena
End If

If Factura = False Then
   PrinterTexto 14.5, 1.2, "Recibo No. " & Format(.Fields("Recibo_No"), "000000")
End If
PrinterTexto 1, 1.2, "Envio No. " & .Fields("Envio_No")
Printer.FontSize = 10
PrinterTexto 2, 11.5, "[" & Format(.Fields("Sucursal"), "000") & "-" & CodigoUsuario & "]  Cajero(a)" '.Fields("Usuario")
PrinterTexto 10, 11.5, "Firma Cliente:__________________________"
If Factura Then
   PrinterTexto 12, 12, "CI/RUC: " & .Fields("CI_RUCR")
Else
   PrinterTexto 12, 12, "CI/RUC: " & .Fields("CI_RUC")
End If
Printer.FontUnderline = True
PrinterTexto 15.5, 2, "Fecha E.:"
PrinterTexto 15.5, 2.5, "Hora:"
PrinterTexto 15.5, 3, "Fecha :"
PrinterTexto 1, 3.6, "REMITENTE"
PrinterTexto 10, 3.6, "BENEFICIARIO"
PrinterTexto 1, 6, "S E R I E S:"
PrinterTexto 10, 6, "M E N S A J E:"
Printer.FontUnderline = False
If Factura Then
   PrinterTexto 12.5, 7.8, "CANTIDAD DE ENVIO"
   PrinterTexto 12.5, 8.3, "C A R G O"
   PrinterTexto 12.5, 8.8, "TOTAL COBRADO"
Else
   PrinterTexto 13, 7.8, "TOTAL ENVIO"
   PrinterTexto 13, 8.3, "TOTAL ABONO"
End If
Printer.FontSize = 9
PrinterTexto 1, 2, Direccion
PrinterTexto 1, 2.5, "Teléfono: " & Telefono1
PrinterTexto 1, 3, NombreCiudad & " - " & NombrePais
Printer.FontItalic = False
PrinterTexto 17.3, 2, .Fields("Fecha")
PrinterTexto 17.3, 2.5, Str(Time)
PrinterTexto 17.3, 3, .Fields("Fecha_P")

PrinterTexto 1, 4.1, .Fields("Remitente")
PrinterTexto 1, 4.6, .Fields("TelefonoR")
NumeroLineas = PrinterLineasMayor(1, 5.1, .Fields("DireccionR"), 8.5)

PrinterTexto 10, 4.1, .Fields("Beneficiario")
PrinterTexto 10, 4.6, .Fields("TelefonoB")
NumeroLineas = PrinterLineasMayor(10, 5.1, .Fields("Direccionb"), 8.5)
Total = .Fields("TOTAL")
Saldo = .Fields("PAGADO")
If .Fields("PAGADO") > 0 Then
    If .Fields("Cotizacion") > 0 Then
        Total = Total - (.Fields("PAGADO") / .Fields("Cotizacion"))
    End If
End If
If Factura Then
   PrinterVariables 16.8, 7.8, .Fields("TOTAL") - .Fields("Porc_C")
   PrinterVariables 16.8, 8.3, .Fields("Porc_C")
   PrinterVariables 16.8, 8.8, .Fields("TOTAL")
Else
   PrinterFields 16.8, 7.8, .Fields("TOTAL"), False   'Total
   PrinterFields 16.8, 8.3, .Fields("TOTAL"), False   'Total
End If
'PrinterVariables 16, 8.3, Total 'Total
'PrinterVariables 16, 8.8, Saldo 'Total
NumeroLineas = PrinterLineasMayor(10, 6.5, .Fields("Mensaje"), 8.5)
'If Saldo > 0 Then
'   Cadena = "Cambio: " & .Fields("Cotizacion") & ", "
'   Cadena = Cadena & SMN & " "
'   Cadena = Cadena & Format(Saldo, "#,##0.00")
'   PrinterTexto 1, 9.7, Cadena
'End If
End With
Printer.FontSize = 9
Mensaje = "Declaro bajo la gravedad de juramento que los fondos de esta operación tienen "
Mensaje = Mensaje & "orígen y destino lícito. Eximo a FAMIREMESAS ECUADOR S.A. de toda la responsabilidad, "
Mensaje = Mensaje & "inclusive de terceros, si esta declaración fuese falsa o errónea."
NumeroLineas = PrinterLineasMayor(1, 9.5, Mensaje, 17.5)
'If ConCopia Then
'   Printer.NewPage
'   ConCopia = False
'   GoTo Volver_Imp
'End If
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
Public Sub SetearEnvios(TipoConsult As Byte, DtaEnvio As Adodc, NoEnvio As String)
     sSQL = "SELECT Co.*,"
     sSQL = sSQL & "Be.Cliente,"
     sSQL = sSQL & "Be.Direccion As DireccionB,"
     sSQL = sSQL & "Be.Telefono As TelefonoB,"
     sSQL = sSQL & "Be.Grupo As CodCiudadB,"
     sSQL = sSQL & "Su.Ciudad As CiudadB,"
     sSQL = sSQL & "Be.Pais As PaisB,"
     sSQL = sSQL & "Be.CI_RUC,"
     sSQL = sSQL & "Re.Remitente,"
     sSQL = sSQL & "Re.Direccion As DireccionR,"
     sSQL = sSQL & "Re.Telefono As TelefonoR,"
     sSQL = sSQL & "C.Corresponsal,"
     sSQL = sSQL & "Re.Pais As PaisR,"
     sSQL = sSQL & "Re.CI_RUC As CI_RUCR,"
     sSQL = sSQL & "Re.Ciudad As CiudadR "
     sSQL = sSQL & "FROM Correos As Co,"
     sSQL = sSQL & "Clientes As Be,"
     sSQL = sSQL & "Remitentes As Re,"
     sSQL = sSQL & "Empresas As Su,"
     sSQL = sSQL & "Corresponsal As C "
     sSQL = sSQL & "WHERE Co.Cod_B = Be.Codigo "
     sSQL = sSQL & "AND Co.Cod_R = Re.Codigo_R "
     sSQL = sSQL & "AND Co.Cod_C = C.Codigo_C "
     sSQL = sSQL & "AND Be.Item = Su.Item "
     Select Case TipoConsult
       Case 1: sSQL = sSQL & "AND Co.Envio_No = '" & NoEnvio & "' "
       Case 2: sSQL = sSQL & "AND Re.Remitente = '" & NoEnvio & "' "
       Case 3: sSQL = sSQL & "AND Be.Cliente = '" & NoEnvio & "' "
     End Select
     sSQL = sSQL & "AND Co.T = '" & Pendiente & "' "
     'sSQL = sSQL & "OR Co.T = '" & Cancelado & "' "
     SelectData DtaEnvio, sSQL
End Sub

Public Sub SetearEnviosC(TipoConsult As Byte, DtaEnvio As Adodc, NoEnvio As String)
     sSQL = "SELECT Co.*,"
     sSQL = sSQL & "Be.Cliente,"
     sSQL = sSQL & "Be.Direccion As DireccionB,"
     sSQL = sSQL & "Be.Telefono As TelefonoB,"
     sSQL = sSQL & "Su.Item As CodCiudadB,"
     sSQL = sSQL & "Su.Ciudad As CiudadB,"
     sSQL = sSQL & "Be.Pais As PaisB,"
     sSQL = sSQL & "Be.CI_RUC,"
     sSQL = sSQL & "Re.Remitente,"
     sSQL = sSQL & "Re.Direccion As DireccionR,"
     sSQL = sSQL & "Re.Telefono As TelefonoR,"
     sSQL = sSQL & "C.Corresponsal,"
     sSQL = sSQL & "Re.CI_RUC As CI_RUCR,"
     sSQL = sSQL & "Re.Pais As PaisR,"
     sSQL = sSQL & "Re.Ciudad As CiudadR "
     sSQL = sSQL & "FROM Correos As Co,"
     sSQL = sSQL & "Clientes As Be,"
     sSQL = sSQL & "Remitentes As Re,"
     sSQL = sSQL & "Empresas As Su,"
     sSQL = sSQL & "Corresponsal As C "
     sSQL = sSQL & "WHERE Co.Cod_B = Be.Codigo "
     sSQL = sSQL & "AND Co.Cod_R = Re.Codigo_R "
     sSQL = sSQL & "AND Co.Cod_C = C.Codigo_C "
     sSQL = sSQL & "AND Be.Grupo = Su.Item "
     Select Case TipoConsult
       Case 1: sSQL = sSQL & "AND Co.Envio_No = '" & NoEnvio & "' "
       Case 2: sSQL = sSQL & "AND Re.Remitente = '" & NoEnvio & "' "
       Case 3: sSQL = sSQL & "AND Be.Cliente = '" & NoEnvio & "' "
     End Select
     sSQL = sSQL & "AND Co.T = '" & Cancelado & "' "
     SelectData DtaEnvio, sSQL
End Sub

Public Sub SetearEnviosN(TipoConsult As Byte, DtaEnvio As Adodc, NoEnvio As String)
     sSQL = "SELECT Co.*," _
          & "Be.Cliente," _
          & "Be.Direccion As DireccionB," _
          & "Be.Telefono As TelefonoB," _
          & "Su.Item As CodCiudadB," _
          & "Su.Ciudad As CiudadB," _
          & "Be.Pais As PaisB," _
          & "Re.Remitente," _
          & "Re.Direccion As DireccionR," _
          & "Re.Telefono As TelefonoR," _
          & "Corresp.Codigo_C," _
          & "Corresp.Corresponsal," _
          & "Re.CI_RUC As CI_RUCR," _
          & "Re.Pais As PaisR," _
          & "Re.Ciudad As CiudadR " _
          & "FROM Correos As Co," _
          & "Clientes As Be," _
          & "Remitentes As Re," _
          & "Empresas As Su," _
          & "Corresponsal As Corresp " _
          & "WHERE Co.Cod_B = Be.Codigo " _
          & "AND Co.Cod_R = Re.Codigo_R " _
          & "AND Re.Codigo_C = Corresp.Codigo_C " _
          & "AND Be.Grupo = Su.Item "
     Select Case TipoConsult
       Case 1: sSQL = sSQL & "AND Co.Envio_No = '" & NoEnvio & "' "
       Case 2: sSQL = sSQL & "AND Re.Remitente = '" & NoEnvio & "' "
       Case 3: sSQL = sSQL & "AND Be.Cliente = '" & NoEnvio & "' "
     End Select
     'sSQL = sSQL & "AND Co.T = '" & Normal & "' "
     SelectData DtaEnvio, sSQL
End Sub

Public Sub SetearEnviosF(TipoConsult As Byte, DtaEnvio As Adodc, NoEnvio As String)
     sSQL = "SELECT Co.*,"
     sSQL = sSQL & "Be.Cliente,"
     sSQL = sSQL & "Be.Direccion As DireccionB,"
     sSQL = sSQL & "Be.Telefono As TelefonoB,"
     sSQL = sSQL & "Su.Codigo As CodCiudadB,"
     sSQL = sSQL & "Su.Ciudad As CiudadB,"
     sSQL = sSQL & "Be.Pais As PaisB,"
     sSQL = sSQL & "Re.Remitente,"
     sSQL = sSQL & "Re.Direccion As DireccionR,"
     sSQL = sSQL & "Re.Telefono As TelefonoR,"
     sSQL = sSQL & "Corresp.Codigo_C,"
     sSQL = sSQL & "Corresp.Corresponsal,"
     sSQL = sSQL & "Re.CI_RUC As CI_RUCR,"
     sSQL = sSQL & "Re.Pais As PaisR,"
     sSQL = sSQL & "Re.Ciudad As CiudadR "
     sSQL = sSQL & "FROM Correos As Co,"
     sSQL = sSQL & "Clientes As Be,"
     sSQL = sSQL & "Remitentes As Re,"
     sSQL = sSQL & "Empresas As Su,"
     sSQL = sSQL & "Corres_Envios As Corresp "
     sSQL = sSQL & "WHERE Co.Cod_B = Be.Codigo "
     sSQL = sSQL & "AND Co.Cod_R = Re.Codigo_R "
     sSQL = sSQL & "AND Re.Codigo_C = Corresp.Codigo_C "
     sSQL = sSQL & "AND Be.Item = Su.Item "
     sSQL = sSQL & "AND Co.CI = True "
     Select Case TipoConsult
       Case 1: sSQL = sSQL & "AND Co.Envio_No = '" & NoEnvio & "' "
       Case 2: sSQL = sSQL & "AND Re.Remitente = '" & NoEnvio & "' "
       Case 3: sSQL = sSQL & "AND Be.Cliente = '" & NoEnvio & "' "
     End Select
     SelectData DtaEnvio, sSQL
End Sub

Public Sub ImprimirEnvios(Datas As Adodc, SizeLetra As Single)
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
Ancho(9) = 17.5 ' Cantidad
Ancho(10) = 20  '
Pagina = 1
Total = 0
Contador = 0
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
         Contador = Contador + 1
         Printer.FontBold = False
         PrinterTexto 0.6, PosLinea, "Fecha: " & .Fields("Fecha_E")
         PrinterTexto 4, PosLinea, "Remitente: " & .Fields("Remitente")
         Select Case .Fields("T")
           Case "C"
                PrinterTexto 0.6, PosLinea + 0.5, "Pagado"
                PrinterTexto 16, PosLinea, "Fecha P.: " & .Fields("Fecha_P")
           Case "R"
                PrinterTexto 0.6, PosLinea + 0.5, "Recibido"
                PrinterTexto 16, PosLinea, "Fecha P.: "
           Case "P"
                PrinterTexto 0.6, PosLinea + 0.5, "Pendiente"
         End Select
         PrinterTexto 4, PosLinea + 0.5, "Beneficiario: " & .Fields("Beneficiario")
         PrinterTexto 14.4, PosLinea + 0.5, "Teléfono: " & .Fields("Telefono")
         PrinterTexto 0.6, PosLinea + 1, "Envio No. " & .Fields("Envio_No")
         PrinterTexto 6.1, PosLinea + 1, "C.I. " & .Fields("CI_RUC")
         PrinterTexto 12.5, PosLinea + 1, "Ciudad: " & .Fields("Ciudad")
         PrinterTexto 0.6, PosLinea + 1.5, "Giro No. " & .Fields("Giro_No")
         PrinterTexto 6.1, PosLinea + 1.5, "Dirección: " & .Fields("Direccion")
         PrinterTexto 4, PosLinea + 2, "Corresponsal: " & .Fields("Corresponsal")
         PrinterFields Ancho(9) - 0.2, PosLinea + 2, .Fields("TOTAL")
         PrinterTexto Ancho(9) - 1.3, PosLinea + 2, SimbM
         Printer.Line (Ancho(0), PosLinea - 0.1)-(Ancho(0), PosLinea + 3), QBColor(0)
         Printer.Line (Ancho(10), PosLinea - 0.1)-(Ancho(10), PosLinea + 3), QBColor(0)
         If PosLinea + 2.5 <= LimiteAlto Then PrinterVariables 0.6, PosLinea + 2.5, "NOTA: " & .Fields("Llamadas")
         Total = Total + .Fields("TOTAL")
         PosLinea = PosLinea + 3
         Printer.Line (InicioX, PosLinea)-(Ancho(10), PosLinea), QBColor(0)
         PosLinea = PosLinea + 0.1
         If PosLinea >= LimiteAlto Then
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
PrinterTexto Ancho(1), PosLinea, Contador & " Registros"
PrinterTexto Ancho(9) - 2.2, PosLinea, "Total    " & SimbM
PrinterVariables Ancho(9) - 1.3, PosLinea, Total
UltimaLinea = PosLinea + 0.5
RatonNormal
Printer.EndDoc
Exit Sub
Errorhandler:
    RatonNormal
    ErrorDeImpresion
    Exit Sub
End Sub

Public Sub ImprimirEnvios1(Datas As Adodc, FormaImp As Byte, SizeLetra As Single, OpcFactura As Boolean)
Dim TipoM As Boolean
On Error GoTo Errorhandler
RatonReloj
InicioX = 0.5: InicioY = 0
'EscalaCentimetro FormaImp, TipoTimes, SizeLetra
DataAnchoCampos InicioX, Datas, SizeLetra, TipoTimes, FormaImp
Ancho(0) = 0.5  ' T
Ancho(1) = 1    ' Envio_No
Ancho(2) = 3    ' Fecha
Ancho(3) = 4.5  ' Fecha P
Ancho(4) = 6    ' Beneficiario
Ancho(5) = 11   ' Telefono
Ancho(6) = 14.5 ' Ciudad
Ancho(7) = 17   ' TOTAL
Ancho(8) = 19   ' SALDO
Ancho(9) = 21   ' Mensaje
Ancho(10) = 26  '
Pagina = 1: CantCampos = 10
Precio = 0: Saldo = 0: Total = 0
Contador = 0
'Iniciamos la impresion
Printer.FontBold = False
With Datas.Recordset
     .MoveFirst
      EncabezadoDataEnvio Datas, 10
      Printer.FontSize = SizeLetra
      Do While Not .EOF
         Contador = Contador + 1
         PrinterFields Ancho(0), PosLinea, .Fields("T"), True
         PrinterFields Ancho(1), PosLinea, .Fields("Envio_No"), True
         PrinterFields Ancho(2), PosLinea, .Fields("Fecha_E"), True
         PrinterFields Ancho(3), PosLinea, .Fields("Fecha_P"), True
         PrinterFields Ancho(4), PosLinea, .Fields("Beneficiario"), True
         PrinterFields Ancho(5), PosLinea, .Fields("Telefono"), True
         PrinterFields Ancho(6), PosLinea, .Fields("Ciudad"), True
         PrinterFields Ancho(7), PosLinea, .Fields("TOTAL"), True
         PrinterFields Ancho(8), PosLinea, .Fields("SALDO"), True
         Printer.Line (Ancho(9), PosLinea + 0.4)-(Ancho(CantCampos), PosLinea + 0.4), QBColor(0)
         If OpcFactura = False Then
            If .Fields("Llamadas") <> Ninguno Then
             PrinterFields Ancho(9), PosLinea, .Fields("Llamadas"), True
            End If
         End If
         Printer.Line (Ancho(9), PosLinea - 0.1)-(Ancho(9), PosLinea + 0.4), QBColor(0)
         Printer.Line (Ancho(10), PosLinea - 0.1)-(Ancho(10), PosLinea + 0.4), QBColor(0)
         Total = Total + .Fields("TOTAL")
         Saldo = Saldo + .Fields("SALDO")
         PosLinea = PosLinea + 0.4
         If PosLinea > LimiteAlto Then
            Printer.Line (Ancho(0), PosLinea)-(Ancho(CantCampos), PosLinea), QBColor(0)
            Printer.NewPage
            EncabezadoDataEnvio Datas, 10
            Printer.FontSize = SizeLetra
         End If
        .MoveNext
      Loop
End With
Printer.Line (InicioX, PosLinea)-(Ancho(CantCampos), PosLinea), QBColor(0)
PosLinea = PosLinea + 0.1
Printer.Line (InicioX, PosLinea)-(Ancho(CantCampos), PosLinea), QBColor(0)
PosLinea = PosLinea + 0.1
PrinterTexto Ancho(1), PosLinea, Contador & " Registros"
PrinterTexto Ancho(6), PosLinea, "T O T A L E S"
PrinterVariables Ancho(7) - 0.7, PosLinea, Total
PrinterVariables Ancho(8) - 0.7, PosLinea, Saldo
UltimaLinea = PosLinea + 0.5
RatonNormal
Printer.EndDoc
Exit Sub
Errorhandler:
    RatonNormal
    ErrorDeImpresion
    Exit Sub
End Sub

Public Sub ImprimirEnvios2(Datas As Adodc, FormaImp As Byte, SizeLetra As Single)
Dim TipoM As Boolean
On Error GoTo Errorhandler
RatonReloj
InicioX = 0.5: InicioY = 0
'EscalaCentimetro FormaImp, TipoTimes, SizeLetra
DataAnchoCampos InicioX, Datas, SizeLetra, TipoTimes, FormaImp
Ancho(0) = 1    ' T
Ancho(1) = 1.5  ' Envio_No
Ancho(2) = 4.5  ' Giro_No
Ancho(3) = 7.5  ' Fecha
Ancho(4) = 9.6  ' Fecha P
Ancho(5) = 11.7 ' Beneficiario
Ancho(6) = 20.5 ' CI_RUC
Ancho(7) = 23.5 ' TOTAL
Ancho(8) = 26.5 '
Pagina = 1: CantCampos = 8
Precio = 0: Saldo = 0: Total = 0
Contador = 0
'Iniciamos la impresion
Printer.FontBold = False
With Datas.Recordset
     .MoveFirst
      EncabezadoDataEnvio1 Datas, 7
      Printer.FontSize = SizeLetra
      Do While Not .EOF
         Contador = Contador + 1
         PrinterFields Ancho(0), PosLinea, .Fields("T"), True
         PrinterFields Ancho(1), PosLinea, .Fields("Envio_No"), True
         PrinterFields Ancho(2), PosLinea, .Fields("Giro_No"), True
         PrinterFields Ancho(3), PosLinea, .Fields("Fecha_E"), True
         PrinterFields Ancho(4), PosLinea, .Fields("Fecha_P"), True
         PrinterFields Ancho(5), PosLinea, .Fields("Beneficiario"), True
         PrinterFields Ancho(6), PosLinea, .Fields("CI_RUC"), True
         PrinterFields Ancho(7), PosLinea, .Fields("TOTAL"), True
         Printer.Line (Ancho(8), PosLinea - 0.1)-(Ancho(8), PosLinea + 0.4), QBColor(0)
         Total = Total + .Fields("TOTAL")
         Saldo = Saldo + .Fields("SALDO")
         PosLinea = PosLinea + 0.5
         If PosLinea >= LimiteAlto Then
            Printer.Line (Ancho(0), PosLinea)-(Ancho(CantCampos), PosLinea), QBColor(0)
            Printer.NewPage
            EncabezadoDataEnvio1 Datas, 7
            Printer.FontSize = SizeLetra
         End If
        .MoveNext
      Loop
End With
Printer.Line (Ancho(0), PosLinea)-(Ancho(CantCampos), PosLinea), QBColor(0)
PosLinea = PosLinea + 0.1
Printer.Line (Ancho(0), PosLinea)-(Ancho(CantCampos), PosLinea), QBColor(0)
PosLinea = PosLinea + 0.1
PrinterTexto Ancho(0), PosLinea, Contador & " Registros"
If PosLinea + 2 > LimiteAlto Then
   Printer.Line (Ancho(0), PosLinea)-(Ancho(CantCampos), PosLinea), QBColor(0)
   PosLinea = 0
   Printer.NewPage
   EncabezadoDataEnvio1 Datas, 7
   Printer.FontSize = SizeLetra
End If
Printer.FontBold = True
PrinterTexto Ancho(4), PosLinea, "TOTAL ENVIOS PAGADOS USD."
PrinterVariables Ancho(7) - 0.8, PosLinea, Total
PosLinea = PosLinea + 0.5
PrinterTexto Ancho(4), PosLinea, "TOTAL COMISION DEL " & Round(Total_Comision) & " % USD."
Saldo = Total * (Total_Comision / 100)
PrinterVariables Ancho(7) - 0.8, PosLinea, Saldo
PosLinea = PosLinea + 0.5
PrinterTexto Ancho(4), PosLinea, "TOTAL A REEMBOLZARNOS USD."
PrinterVariables Ancho(7) - 0.8, PosLinea, Total + Saldo
PosLinea = PosLinea + 0.7
'If PosLinea + 5 > LimiteAlto Then
'   Printer.Line (Ancho(0), PosLinea)-(Ancho(CantCampos), PosLinea), QBColor(0)
'   PosLinea = 0
'   Printer.NewPage
'   EncabezadoDataEnvio1 Datas, 7
'   Printer.FontSize = SizeLetra
'End If
'Printer.FontSize = 12
'PrinterTexto 1, PosLinea, "INSTRUCCIONES DE PAGO"
'PosLinea = PosLinea + 0.8
'PrinterTexto 1, PosLinea, "TRANSFERENCIA BANCARIA:"
'PosLinea = PosLinea + 0.8
'NumeroLineas = PrinterLineasMayor(1, PosLinea, Codigo2, 24)
'MsgBox NumeroLineas
'PosLinea = PosLinea + (0.5 * NumeroLineas)
'PrinterTexto 1, PosLinea, "SALUDOS"
'PosLinea = PosLinea + 1.5
'PrinterTexto 1, PosLinea, Codigo3
'PosLinea = PosLinea + 1
'If PosLinea + 6 > LimiteAlto Then
'   Printer.Line (Ancho(0), PosLinea)-(Ancho(CantCampos), PosLinea), QBColor(0)
'   PosLinea = 0
'   Printer.NewPage
'   EncabezadoDataEnvio1 Datas, 7
'   Printer.FontSize = 12
'End If
'If Cadena1 <> "." Then
'   PrinterTexto 1.5, PosLinea, "NOTA:"
'   PosLinea = PosLinea + 0.5
'   NumeroLineas = PrinterLineasMayor(1, PosLinea, Cadena1, 24)
'End If
Printer.FontBold = False
RatonNormal
Printer.EndDoc
Exit Sub
Errorhandler:
    RatonNormal
    ErrorDeImpresion
    Exit Sub
End Sub

Public Sub ImprimirFlujoCaja(Datas As Adodc, FinDoc As Boolean, FormaImp As Byte, SizeLetra As Single)
On Error GoTo Errorhandler
RatonReloj
InicioX = 0.5: InicioY = 0
'EscalaCentimetro   FormaImp, TipoTimes, SizeLetra
DataAnchoCampos InicioX, Datas, SizeLetra, TipoTimes, FormaImp
Ancho(0) = 0.5
Ancho(1) = 4.5
Ancho(2) = 6
Ancho(3) = 9.5
Ancho(4) = 12
Ancho(5) = 14.5
Ancho(6) = 17
Ancho(7) = 19.5
CantCampos = 7
Pagina = 1
Total = 0: Total_ME = 0: Saldo = 0: Saldo_ME = 0
'Iniciamos la impresion
Printer.FontBold = False
With Datas.Recordset
     .MoveFirst
      EncabezadoData Datas
      Printer.FontSize = SizeLetra
      Codigo = .Fields("Usuario")
      PrinterFields Ancho(0), PosLinea, .Fields("Usuario"), True
      Do While Not .EOF
         If Codigo <> .Fields("Usuario") Then
            PosLinea = PosLinea + 0.05
            Printer.Line (Ancho(0), PosLinea)-(Ancho(CantCampos), PosLinea), QBColor(0)
            PosLinea = PosLinea + 0.05
            Printer.Line (Ancho(0), PosLinea)-(Ancho(CantCampos), PosLinea), QBColor(0)
            PosLinea = PosLinea + 0.1
            PrinterVariables Ancho(3), PosLinea, Total
            PrinterVariables Ancho(4), PosLinea, Saldo
            PrinterVariables Ancho(5), PosLinea, Total_ME
            PrinterVariables Ancho(6), PosLinea, Saldo_ME
            PosLinea = PosLinea + 0.5
            PrinterFields Ancho(0), PosLinea, .Fields("Usuario"), True
            Codigo = .Fields("Usuario")
            Total = 0: Total_ME = 0: Saldo = 0: Saldo_ME = 0
         Else
            Printer.Line (Ancho(0), PosLinea - 0.1)-(Ancho(0), PosLinea + 0.4), QBColor(0)
         End If
         PrinterFields Ancho(1), PosLinea, .Fields("Fecha"), True
         PrinterFields Ancho(2), PosLinea, .Fields("Envio_No"), True
         PrinterFields Ancho(3), PosLinea, .Fields("Ingreso_MN"), True
         PrinterFields Ancho(4), PosLinea, .Fields("Egreso_MN"), True
         PrinterFields Ancho(5), PosLinea, .Fields("Ingreso_ME"), True
         PrinterFields Ancho(6), PosLinea, .Fields("Egreso_ME"), True
        'PrinterAllFields CantCampos, PosLinea, Datas, True
         Printer.Line (Ancho(CantCampos), PosLinea - 0.1)-(Ancho(CantCampos), PosLinea + 0.4), QBColor(0)
         Total = Total + .Fields("Ingreso_MN")
         Saldo = Saldo + .Fields("Egreso_MN")
         Total_ME = Total_ME + .Fields("Ingreso_ME")
         Saldo_ME = Saldo_ME + .Fields("Egreso_ME")
         PosLinea = PosLinea + 0.35
         If PosLinea > LimiteAlto Then
            Printer.Line (Ancho(0), PosLinea)-(Ancho(CantCampos), PosLinea), QBColor(0)
            Printer.NewPage
            EncabezadoData Datas
            Printer.FontSize = SizeLetra
         End If
        .MoveNext
      Loop
End With
PosLinea = PosLinea + 0.05
Printer.Line (Ancho(0), PosLinea)-(Ancho(CantCampos), PosLinea), QBColor(0)
PosLinea = PosLinea + 0.05
Printer.Line (Ancho(0), PosLinea)-(Ancho(CantCampos), PosLinea), QBColor(0)
PosLinea = PosLinea + 0.1
PrinterVariables Ancho(3), PosLinea, Total
PrinterVariables Ancho(4), PosLinea, Saldo
PrinterVariables Ancho(5), PosLinea, Total_ME
PrinterVariables Ancho(6), PosLinea, Saldo_ME
PosLinea = PosLinea + 0.5
Printer.Line (InicioX, PosLinea)-(Ancho(CantCampos), PosLinea), QBColor(0)
PosLinea = PosLinea + 0.05
Printer.Line (InicioX, PosLinea)-(Ancho(CantCampos), PosLinea), QBColor(0)
Printer.FontSize = 9: Printer.FontBold = True
PosLinea = PosLinea + 0.1
PrinterTexto 1, PosLinea, "INGRESOS M/N:"
PrinterVariables 3, PosLinea, Debe
PosLinea = PosLinea + 0.5
PrinterTexto 1, PosLinea, "EGRESOS M/N:"
PrinterVariables 3, PosLinea, Haber
PosLinea = PosLinea + 0.5
PrinterTexto 1, PosLinea, "SALDO M/N:"
PrinterVariables 3, PosLinea, Debe - Haber
PosLinea = PosLinea + 0.5
PrinterTexto 1, PosLinea, "INGRESOS M/E:"
PrinterVariables 3, PosLinea, Debe_ME
PosLinea = PosLinea + 0.5
PrinterTexto 1, PosLinea, "EGRESOS M/E:"
PrinterVariables 3, PosLinea, Haber_ME
PosLinea = PosLinea + 0.5
PrinterTexto 1, PosLinea, "SALDO M/E:"
PrinterVariables 3, PosLinea, Debe_ME - Haber_ME
PosLinea = PosLinea + 0.5
RatonNormal
Printer.EndDoc
Exit Sub
Errorhandler:
    RatonNormal
    ErrorDeImpresion
    Exit Sub
End Sub

Public Sub ImprimirFlujoCaja1(Datas As Adodc, FinDoc As Boolean, FormaImp As Byte, SizeLetra As Single)
On Error GoTo Errorhandler
RatonReloj
InicioX = 0.5: InicioY = 0
'EscalaCentimetro   FormaImp, TipoTimes, SizeLetra
DataAnchoCampos InicioX, Datas, SizeLetra, TipoTimes, FormaImp
Ancho(0) = 0.5
Ancho(1) = 2
Ancho(2) = 8
Ancho(3) = 9.5
Ancho(4) = 12
Ancho(5) = 14.5
Ancho(6) = 17
Ancho(7) = 19.5
Pagina = 1
'Iniciamos la impresion
Printer.FontBold = False
With Datas.Recordset
     .MoveFirst
      EncabezadoData Datas
      Printer.FontSize = SizeLetra
      Do While Not .EOF
         PrinterAllFields CantCampos, PosLinea, Datas, True
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
Printer.Line (InicioX, PosLinea)-(Ancho(CantCampos), PosLinea), QBColor(0)
PosLinea = PosLinea + 0.1
Printer.Line (InicioX, PosLinea)-(Ancho(CantCampos), PosLinea), QBColor(0)
Printer.FontSize = 9: Printer.FontBold = True
PosLinea = PosLinea + 0.1
PrinterTexto 1, PosLinea, "INGRESOS M/N:"
PrinterVariables 3, PosLinea, Debe
PosLinea = PosLinea + 0.5
PrinterTexto 1, PosLinea, "EGRESOS M/N:"
PrinterVariables 3, PosLinea, Haber
PosLinea = PosLinea + 0.5
PrinterTexto 1, PosLinea, "SALDO M/N:"
PrinterVariables 3, PosLinea, Debe - Haber
PosLinea = PosLinea + 0.5
PrinterTexto 1, PosLinea, "INGRESOS M/E:"
PrinterVariables 3, PosLinea, Debe_ME
PosLinea = PosLinea + 0.5
PrinterTexto 1, PosLinea, "EGRESOS M/E:"
PrinterVariables 3, PosLinea, Haber_ME
PosLinea = PosLinea + 0.5
PrinterTexto 1, PosLinea, "SALDO M/E:"
PrinterVariables 3, PosLinea, Debe_ME - Haber_ME
PosLinea = PosLinea + 0.5
RatonNormal
Printer.EndDoc
Exit Sub
Errorhandler:
    RatonNormal
    ErrorDeImpresion
    Exit Sub
End Sub

Public Sub ImprimirEnvios3(Datas As Adodc, FinDoc As Boolean, FormaImp As Byte, SizeLetra As Single)
Dim NuevoBenef As Boolean
On Error GoTo Errorhandler
RatonReloj
InicioX = 0.5: InicioY = 0
'EscalaCentimetro   FormaImp, TipoTimes, SizeLetra
DataAnchoCampos InicioX, Datas, SizeLetra, TipoTimes, FormaImp
Ancho(0) = 0.5
Ancho(1) = 6.5
Ancho(2) = 9
Ancho(3) = 10.7
Ancho(4) = 12.7
Ancho(5) = 18.7
Ancho(6) = 23.5
Ancho(7) = 26
Pagina = 1
'Iniciamos la impresion
Printer.FontBold = False
With Datas.Recordset
     .MoveFirst
      Encabezado Ancho(0), Ancho(7)
      Printer.FontSize = SizeLetra
      Printer.FontBold = True
      PrinterTexto Ancho(0), PosLinea, "Beneficiario"
      PrinterTexto Ancho(1), PosLinea, "CI_RUC"
      PrinterTexto Ancho(2), PosLinea, "Fecha"
      PrinterTexto Ancho(3), PosLinea, "Envio_No"
      PrinterTexto Ancho(4), PosLinea, "Remitente"
      PrinterTexto Ancho(5), PosLinea, "Corresponsal"
      PrinterTexto Ancho(6), PosLinea, "T O T A L"
      PosLinea = PosLinea + 0.4
      Printer.Line (InicioX, PosLinea)-(Ancho(CantCampos), PosLinea), QBColor(0)
      PosLinea = PosLinea + 0.1
      Printer.FontBold = False
      NuevoBenef = True
      Codigo = .Fields("Beneficiario")
      Total = 0
      Do While Not .EOF
         Printer.FontBold = False
         If Codigo <> .Fields("Beneficiario") Then
            Printer.FontBold = True
            Printer.Line (InicioX, PosLinea)-(Ancho(CantCampos), PosLinea), QBColor(0)
            PosLinea = PosLinea + 0.1
            PrinterVariables Ancho(6), PosLinea, CSng(Total)
            PrinterTexto Ancho(5) + 1, PosLinea, "TOTAL ENVIOS"
            PosLinea = PosLinea + 0.4
            Printer.Line (InicioX, PosLinea)-(Ancho(CantCampos), PosLinea), QBColor(0)
            PosLinea = PosLinea + 0.2
            Codigo = .Fields("Beneficiario")
            NuevoBenef = True
            Total = 0
            Printer.FontBold = False
         End If
         If NuevoBenef Then
            PrinterFields Ancho(0), PosLinea, .Fields("Beneficiario"), False
            PrinterFields Ancho(1), PosLinea, .Fields("CI_RUCB"), False
            NuevoBenef = False
         End If
         PrinterFields Ancho(2), PosLinea, .Fields("Fecha"), False
         PrinterFields Ancho(3), PosLinea, .Fields("Envio_No"), False
         PrinterFields Ancho(4), PosLinea, .Fields("Remitente"), False
         PrinterFields Ancho(5), PosLinea, .Fields("Corresponsal"), False
         PrinterFields Ancho(6), PosLinea, .Fields("TOTAL"), False
         Total = Total + .Fields("TOTAL")
         'PrinterAllFields CantCampos, PosLinea, Datas, Decimales, True
         PosLinea = PosLinea + 0.4
         If PosLinea > LimiteAlto Then
            Printer.Line (Ancho(0), PosLinea)-(Ancho(CantCampos), PosLinea), QBColor(0)
            Printer.NewPage
            Encabezado Ancho(0), Ancho(7)
            Printer.FontSize = SizeLetra
            Printer.FontBold = True
            PrinterTexto Ancho(0), PosLinea, "Beneficiario"
            PrinterTexto Ancho(1), PosLinea, "CI_RUC"
            PrinterTexto Ancho(2), PosLinea, "Fecha"
            PrinterTexto Ancho(3), PosLinea, "Envio_No"
            PrinterTexto Ancho(4), PosLinea, "Remitente"
            PrinterTexto Ancho(5), PosLinea, "Corresponsal"
            PrinterTexto Ancho(6), PosLinea, "T O T A L"
            PosLinea = PosLinea + 0.4
            Printer.Line (InicioX, PosLinea)-(Ancho(CantCampos), PosLinea), QBColor(0)
            PosLinea = PosLinea + 0.1
            Printer.FontBold = False
         End If
        .MoveNext
      Loop
End With
Printer.FontBold = True
Printer.Line (InicioX, PosLinea)-(Ancho(CantCampos), PosLinea), QBColor(0)
PosLinea = PosLinea + 0.1
PrinterVariables Ancho(6), PosLinea, CSng(Total)
PrinterTexto Ancho(5) + 1, PosLinea, "TOTAL ENVIOS"
PosLinea = PosLinea + 0.4
Printer.Line (InicioX, PosLinea)-(Ancho(CantCampos), PosLinea), QBColor(0)
PosLinea = PosLinea + 0.1
Printer.Line (InicioX, PosLinea)-(Ancho(CantCampos), PosLinea), QBColor(0)
Printer.FontBold = False
RatonNormal
If FinDoc Then Printer.EndDoc
Exit Sub
Errorhandler:
    RatonNormal
    ErrorDeImpresion
    Exit Sub
End Sub
