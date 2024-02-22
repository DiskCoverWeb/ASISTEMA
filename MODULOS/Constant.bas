Attribute VB_Name = "Constantes"
Option Explicit

'-------------------------------------------------------
'SERVIRLES ES NUESTRO COMPROMISO, DISFRUTARLO ES EL SUYO.
'-------------------------------------------------------
' Prueba:     celcer.sri.gob.ec
' Produccion: cel.sri.gob.ec

'-------------------------------------------------------
'Tipo_Concepto_Retencion
'BINARY       1 byte por carácter Se puede almacenar cualquier tipo de datos en un campo de
'               este tipo. Los datos no se traducen (por ejemplo, a texto). La forma en que se
'               introducen los datos en un campo binario indica cómo aparecerán al mostrarlos.
'BIT          1 byte Valores Sí y No, y campos que contienen solamente uno de dos valores.
'BYTE         1 byte Un número entero entre 0 y 255.
'COUNTER      4 bytes Un número que el motor de base de datos Microsoft Jet incrementa
'               automáticamente siempre que se agrega un registro nuevo a una tabla.
'               En el motor de base de datos Microsoft Jet, el tipo de datos para este
'               valor es Long.
'CURRENCY     8 bytes Un número entero comprendido entre  – 922.337.203.685.477,5808 y
'               922.337.203.685.477,5807.
'DATETIME     8 bytes Una valor de fecha u hora entre los años 100 y 9999 (Vea DOUBLE)
'GUID         128 bits Un número de identificación único utilizado con llamadas a
'               procedimientos remotos.
'SINGLE       4 bytes Un valor de signo flotante de precisión simple con un intervalo
'               comprendido entre  – 3,402823E38 y  – 1,401298E-45 para valores negativos,
'               y desde 1,401298E-45 a 3,402823E38 para valores positivos, y 0.
'DOUBLE       8 bytes Un valor de signo flotante de precisión doble con un intervalo
'               comprendido entre  – 1,79769313486232E308 y  – 4,94065645841247E-324 para
'               valores negativos, y desde 4,94065645841247E-324 a 1,79769313486232E308 para
'               valores positivos, y 0.
'SHORT        2 bytes Un entero corto, entre  – 32.768 y 32.767.
'LONG         4 bytes Un entero largo entre  – 2.147.483.648 y 2.147.483.647.
'LONGTEXT     1 byte por carácterr    Desde cero hasta un máximo de 1,2 gigabytes.
'LONGBINARY   Lo que se requiera  Desde cero hasta un máximo de 1,2 gigabytes. Se utiliza
'             para objetos OLE.
'TEXT         1 byte por carácter Desde cero a 255 caracteres.
'=================================================================
'Funciones de SQL
'Constantes para acceso a Datos por ADO/ADODB/OLE DB
'Constante          Valor
'-----------------------------------------------------------------
'Nuevo Comprobantes     : 1
'Modificar Comprobantes : 1
'Copiar Comprobantes    : 1
'Saldos Ctas. Especiales: 15
'Conciliacion           : 30
'Pago a Bancos Cash     : 35
'Anexos Transaccionales : 96
'Cierre de Caja         : 96,97,255
'Importar Excele        : 100,180,190,199
'Suscripciones          : 250
'-----------------------------------------------------------------
' Trans_No = 36-78, 201-249
'-----------------------------------------------------------------
'Este Campo se cambiara segun el tiempo y la version de conexion

''' "USTED NO ESTA LEGALIZADO LLAME A SU" & vbCrLf _
'''                 & "PROVEEDOR O A NUESTRAS OFICINAS A" & vbCrLf _
'''                 & "LOS TELEFONOS: 02-321-0051/099-965-4196/098-910-5300" & vbCrLf _
'''                 & "EN QUITO - ECUADOR O LOS EMAILS:" & vbCrLf _
'''                 & "asistencia@diskcoversystem.com / prisma_net@hotmail.es" & vbCrLf _
'''                 & "PARA SU LEGALIZACION"
'--------------------------------------------------------------------------------------------------------------------------------------------
'https://srienlinea.sri.gob.ec/sri-catastro-sujeto-servicio-internet/rest/ConsolidadoContribuyente/existePorNumeroRuc?numeroRuc=0702164179001
'https://srienlinea.sri.gob.ec/facturacion-internet/consultas/publico/ruc-datos2.jspa?accion=siguiente&ruc=0702164179001
'-----------------------------------------------------------------------------------------------------------------------
Global Const TextoLeyendaFA = "Para consultas, requerimientos o reclamos puede contactarse a nuestro Centro de Atención al Cliente Teléfono: 02-6052430, " _
                            & "o escriba al correo prisma_net@hotmail.es; para Transferencia o Depósitos hacer en El Banco Pichincha: Cta. Ahr. 4245946100 a " _
                            & "Nombre de Walter Vaca Prieto/Cta. Cte 3422225804, a Nombre de PRISMANET PROFESIONAL S.A."

Global Const TextoLeyendaFA1 = "SERVIRLE ES NUESTRO OBJETIVO, DISFRUTARLO EL SUYO"

Global Const MensajeAutomatizado = "Mensaje_Comunicado" & vbCrLf _
                                 & "Este correo electronico fue generado automaticamente a usted desde El Sistema Financiero Contable DiskCover System, " _
                                 & "porque figura como correo electronico alternativo de Razon_Social. " _
                                 & "Nosotros respetamos su privacidad y solamente se utiliza este medio para mantenerlo informado sobre nuestras ofertas, " _
                                 & "promociones y comunicados. No compartimos, publicamos o vendemos su informacion personal fuera de nuestra empresa. " _
                                 & "Este mensaje fue procesado por: Nombre_Usuario, funcionario que forma parte de la Institucion." & vbCrLf & vbCrLf _
                                 & "Por la atencion que se de al presente quedo de usted." & vbCrLf _
                                 & vbCrLf _
                                 & "Atentamente," & vbCrLf _
                                 & vbCrLf _
                                 & "Representante_Legal" & vbCrLf _
                                 & "Razon_Social" & vbCrLf _
                                 & vbCrLf _
                                 & "Esta direccion de correo electronico no admite respuestas. En caso de requerir atencion personalizada por parte de un " _
                                 & "asesor de Servicio al Cliente de Razon_Social, podra solicitar ayuda mediante los canales oficiales que detallamos a " _
                                 & "continuación: Telefonos: Numero_Telefono Correo: Emails." & vbCrLf & vbCrLf _
                                 & "www.diskcoversystem.com" & vbCrLf _
                                 & "QUITO - ECUADOR" & vbCrLf

Global Const MensajeNoAutorizarCE = "LA EMPRESA NO TIENE ACTIVADO EL PROCESO PARA SOLICITAR AUTORIZACION AL SRI, " _
                                  & "COMUNIQUESE AL CENTRO DE ATENCION AL CLIENTE DEL SISTEMA DISKCOVER SYSTEM. " _
                                  & "A LOS TELEFONOS: (+593) 098-652-4396/099-965-4196/098-910-5300."
                                  
Global Const MensajeDeboPagare = "Debo y Pagaré incondicionalmente a la orden de Razon_Social el valor expresado en este documento mas " _
                               & "el máximo interés legal por mora, vigente en el Sistema Financiero Nacional desde la fecha de vencimiento, " _
                               & "SIN PROTESTO, exímase de presentación para el pago así como la falta de estos hechos. Renuncio fuero y domicilio " _
                               & "y me someto a los jueces competentes de la ciudad de Quito, Distrito Metropolitano, y al trámite verbal sumario o " _
                               & "ejecutivo a elección de Razon_Social o de sus cesionarios. Acepto que Razon_Social, ceda y transfiera en cualquier " _
                               & "momento los derechos que emanan de la presente factura-pagaré sin que sea necesaria notificación algún ni nueva " _
                               & "aceptación de mi parte. Suscribo la presente factura-pagaré en conformidad con todos sus términos."
                               
Global Const MensajeDeAdvertencia = "DISKCOVER SYSTEM representado por PrismaNet Profesional S.A.; No se responsabiliza de la pérdida parcial o total " _
                                  & "de información si esta usando una copia Ilegal."
                                  
Global Const ServidorEnLineaSRI = "En estos momentos el servidor de Aprobacion de Documentos Electronicos en Ambiente de XXXX no esta en linea, " _
                                & "no se podra aprobar sus documentos, podra generar los comprobantes y despues enviar autorizar al SRI."

'-------------------------------------------------------------------------------
'Datos de Conexion a la Base de Datos en las nubes db.diskcoversystem.com:13306
'-------------------------------------------------------------------------------
'Global Const AdoStrCnnMySQL = "DRIVER={MySQL ODBC 3.51 Driver};"
'Global Const AdoStrCnnMySQL = "DRIVER={MySQL ODBC 5.1 Driver};"
'Global Const AdoStrCnnMySQL = "DRIVER={MySQL ODBC 8.2 ANSI Driver};"
'Global Const AdoStrCnnMySQL = "DRIVER={MySQL ODBC 8.2 Unicode Driver};"
'-------------------------------------------------------------------------------
Global Const AdoStrCnnMySQL = "DRIVER={MySQL ODBC 5.1 Driver};" _
                            & "SERVER=db.diskcoversystem.com;" _
                            & "PORT=13306;" _
                            & "DATABASE=diskcover_empresas;" _
                            & "USER=diskcover;" _
                            & "PASSWORD=disk2017@Cover;" _
                            & "OPTION=3;"
                                                           
'Global Const AdoStrCnnMySQL = "Driver={MySQL ODBC 8.2 Unicode Driver};SERVER=db.diskcoversystem.com;DATABASE=diskcover_empresas;USER=diskcover;PASSWORD=disk2017@Cover;PORT=13306; OPTION=3"
Global Const strServidor = "db.diskcoversystem.com"

Global Const urlIdukay = "https://erp.diskcoversystem.com/php/vista/consultarEstudiante.php?id="
Global Const urlEsUnRUC = "https://srienlinea.sri.gob.ec/sri-catastro-sujeto-servicio-internet/rest/ConsolidadoContribuyente/existePorNumeroRuc?numeroRuc="
Global Const urlDatosDelRUC = "https://srienlinea.sri.gob.ec/facturacion-internet/consultas/publico/ruc-datos2.jspa?accion=siguiente&ruc="
'-----------------------------------------------------------------
Global Const CorreoDiskCover = "informacion@diskcoversystem.com"
Global Const ContrasenaDiskCover = "infoDlcjvl1210DiskCover"
Global Const CorreoUpdate = "actualizar@diskcoversystem.com"
'LINODE
Global Const ftpSvrLinode = "ftp.diskcoversystem.com"
Global Const ftpUseLinode = "ftpuser"
Global Const ftpPwrLinode = "ftp2023User"
Global Const ftpPuerto = 21
'NUC
Global Const ftpSvr = "vpn.diskcoversystem.com"
Global Const ftpUse = "ftpuser"
Global Const ftpPwr = "ftp2023User"
'-----------------------------------------------------------------
Global Const Car_Visto = 251
'-----------------------------------------------------------------
Global Const adTrue = -1
Global Const adFalse = 0
'-------------------------------------
Global Const Es_Printer = 1
Global Const Es_Picture = 2
Global Const Es_PDF = 3

'-----------------------------------------------------------------
'Conversion de Datos
Global Const TadDate = adDate             ' SmallDateTime
Global Const TadDate1 = adDBTimeStamp     ' SmallDateTime
Global Const TadTime = adDate             ' SmallDateTime
Global Const TadBoolean = adBoolean       ' Bit
Global Const TadByte = adUnsignedTinyInt  ' TinyInt
Global Const TadInteger = adSmallInt      ' SmallInt
Global Const TadLong = adInteger          ' Int
Global Const TadDouble = adDouble         ' Float
Global Const TadSingle = adSingle         ' Real
Global Const TadCurrency = adCurrency     ' Money
Global Const TadDecimal = adNumeric       ' Decimal
Global Const TadText = adVarWChar         ' NVarChar
Global Const TadString = adVarWChar       ' NVarChar
Global Const TadMemo = adLongVarWChar     ' NText
'----------------------------------------
'Colores Generales
'----------------------------------------
Global Const Negro = vbBlack          '00
Global Const Azul = vbBlue            '01
Global Const Verde = vbGreen          '02
Global Const Aguamarina = 8421376     '03
Global Const Rojo = vbRed             '04
Global Const Fucsia = 8388736         '05
Global Const Amarillo = vbYellow      '06
Global Const Plata = 12632256         '07
Global Const Gris = 8421504           '08
Global Const Azul_Claro = 16711680    '09
Global Const Verde_Claro = 65280      '10
Global Const Magenta = vbMagenta      '11
Global Const Rojo_Claro = 255         '12
Global Const Fucsia_Claro = 16711935  '13
Global Const Amarillo_Claro = 65535   '14
Global Const Blanco = vbWhite         '15
Global Const Blanco_Claro = 16777215  '15
Global Const Turquesa = vbCyan
'----------------------------------------------------------------------------------------------------
'Los métodos y propiedades de aplicación que aceptan una especificación de color esperan que dicha
'especificación sea un número que representa un valor de color RGB. Un valor de color RGB especifica
'la intensidad relativa de rojo, verde y azul para hacer que se muestre un color específico.
'Si el valor de cualquier argumento de RGB es mayor que 255, se utiliza 255.
'En la siguiente tabla se enumeran algunos colores estándar y los valores de rojo, verde y azul que
'incluyen.
'----------------------------------------------------------------------------------------------------
'Color      Rojo    Verde   Azul
'----------------------------------------------------------------------------------------------------
'Black      0       0       0
'Blue       0       0       255
'Green      0       255     0
'Cyan       0       255     255
'Red        255     0       0
'Magenta    255     0       255
'Yellow     255     255     0
'White      255     255     255
'----------------------------------------------------------------------------------------------------
'Abadi MT Condensed
Global Const TipoArial = "Arial"
Global Const TipoArialNarrow = "Arial Narrow"
Global Const TipoArialBlack = "Arial Black"
Global Const TipoArialUnicode = "Arial Unicode MS"
Global Const TipoAvantGarde = "AvantGarde Bk BT"
Global Const TipoSerif = "MS Serif"
Global Const TipoSansSerif = "MS Sans Serif"
Global Const TipoCondensed = "Bernard MT Condensed"
Global Const TipoComicSans = "Comic Sans MS"
Global Const TipoConsola = "Lucida Console"
Global Const TipoCalibri = "Calibri"
Global Const TipoCourier = "Courier"
Global Const TipoCourierNew = "Courier New"
Global Const TipoTimes = "Times New Roman"
Global Const TipoTimesRoman = "Times"
Global Const TipoTerminal = "Terminal"
Global Const TipoSystem = "System"
Global Const TipoSymbol = "Symbol"
Global Const TipoGeorgia = "Georgia"
Global Const TipoHelvetica = "Helvetica"
Global Const TipoHelveticaBold = "Helvetica-Bold"
Global Const TipoTahoma = "Tahoma"
Global Const TipoVerdana = "Verdana"
Global Const TipoWingdings = "Wingdings"
'-----------------------------------
Global Const Impresota_PDF = "PDF DISKCOVER SYSTEM"
'-----------------------------------
Global Const Ninguno = "."
Global Const Normal = "N"
Global Const Pendiente = "P"
Global Const Procesado = "P"
Global Const Cancelado = "C"
Global Const Renovacion = "R"
Global Const Suspenso = "S"
Global Const Anulado = "A"
'-----------------------------------
Global Const PagoCont = "CONTADO"
Global Const PagoCred = "CREDITO"
Global Const PagoCheq = "CHEQ"
Global Const PagoEfec = "EFEC"
Global Const PagoEfec_Cheq = "EFEC/CHEQ"
Global Const PagoTarjetaCredito = "TARJETA DE CREDITO"
Global Const SinDatos = "Vacio"
Global Const ConsumidorFinal = "9999999999"
'-----------------------------------
Global Const CompDiario = "CD"
Global Const CompIngreso = "CI"
Global Const CompEgreso = "CE"
Global Const CompFactura = "FA"
Global Const CompNotaVenta = "NV"
Global Const CompNotaDebito = "ND"
Global Const CompNotaCredito = "NC"
Global Const CompDiarioCaja = "DC"
Global Const CompRetencion = "CR"
Global Const CompAbonoFact = "AF"
Global Const CtaBancos = "BA"
Global Const CtaCaja = "CJ"
'-----------------------------------
Global Const ventas = "Ven"
Global Const CxC = "CxC"
Global Const EstIndiv = "I"
Global Const EstGrup = "G"
Global Const EstTransp = "T"
'-----------------------------------------
Global Const MP = "MP"
Global Const PP = "PP"
Global Const PT = "PT"
'-----------------------------------------
Global Const MascaraFechas = "##/##/####"
Global Const LimpiarFechas = "00/00/0000"
Global Const FormatoFechas = "dd/MM/yyyy"
Global Const MascaraTimes = "##:##:##"
Global Const LimpiarTimes = "00:00:00"
Global Const FormatoTimes = "hh:mm:ss"
Global Const MascaraCtas6 = "##.###"
Global Const FormatoCtas6 = "0#.###"
Global Const MascaraCtas5 = "##.##"
Global Const FormatoCtas5 = "0#.##"

Global Const MascaraArt = "##.###"
Global Const LimpiarArt = "00.000"
Global Const MascaraRUC = "#############"
Global Const LimpiarRUC = "9999999990001"
Global Const MascaraCI = "##########"
Global Const LimpiarCI = "9999999990"
Global Const MascaraTelefC = "##-###-###"
Global Const LimpiarTelefC = "00-000-000"
Global Const MascaraTelef = "###-###"
Global Const LimpiarTelef = "000-000"
Global Const CadenaValida = "0123456789;@.-_áéíóúüñqwertyuiopasdfghjklzxcvbnm ÁÉÍÓÚÜÑQWERTYUIOPASDFGHJKLZXCVBNM"
'-----------------------------------------
Global Const MaxVect = 80
