Attribute VB_Name = "ConstantesP"
Option Explicit

''Tipo_Concepto_Retencion
'''BINARY       1 byte por carácter Se puede almacenar cualquier tipo de datos en un campo de
'''               este tipo. Los datos no se traducen (por ejemplo, a texto). La forma en que se introducen los datos en un campo binario indica cómo aparecerán al mostrarlos.
'''BIT          1 byte  Valores Sí y No, y campos que contienen solamente uno de dos valores.
'''BYTE         1 byte  Un número entero entre 0 y 255.
'''COUNTER      4 bytes Un número que el motor de base de datos Microsoft Jet incrementa
'''               automáticamente siempre que se agrega un registro nuevo a una tabla.
'''               En el motor de base de datos Microsoft Jet, el tipo de datos para este
'''               valor es Long.
'''CURRENCY     8 bytes Un número entero comprendido entre  – 922.337.203.685.477,5808 y
'''               922.337.203.685.477,5807.
'''DATETIME     8 bytes Una valor de fecha u hora entre los años 100 y 9999 (Vea DOUBLE)
'''GUID         128 bits Un número de identificación único utilizado con llamadas a
'''               procedimientos remotos.
'''SINGLE       4 bytes Un valor de signo flotante de precisión simple con un intervalo
'''               comprendido entre  – 3,402823E38 y  – 1,401298E-45 para valores negativos,
'''               y desde 1,401298E-45 a 3,402823E38 para valores positivos, y 0.
'''DOUBLE       8 bytes Un valor de signo flotante de precisión doble con un intervalo
'''               comprendido entre  – 1,79769313486232E308 y  – 4,94065645841247E-324 para
'''               valores negativos, y desde 4,94065645841247E-324 a 1,79769313486232E308 para
'''               valores positivos, y 0.
'''SHORT        2 bytes Un entero corto, entre  – 32.768 y 32.767.
'''LONG         4 bytes Un entero largo entre  – 2.147.483.648 y 2.147.483.647.
'''LONGTEXT     1 byte por carácterr    Desde cero hasta un máximo de 1,2 gigabytes.
'''LONGBINARY   Lo que se requiera  Desde cero hasta un máximo de 1,2 gigabytes. Se utiliza
'''             para objetos OLE.
'''TEXT         1 byte por carácter Desde cero a 255 caracteres.

'Consultas Frecuentes mas utilizadas
'=================================================================
'
'"BETWEEN #" & FechaIni & "# and #" & FechaFin & "# "
'
'       & "WHERE Item = '" & C1.Item & "' " _
'       & "AND T_No = " & C1.T_No & " " _
'       & "AND CodigoU = '" & CodigoUsuario & "' "
'
'      .MoveFirst
'      .Find ("Producto Like '" & Codigo & "' ")
'       If Not .EOF Then
'       Else
'       End If
'=================================================================
'Funciones de SQL
'Constantes para acceso a Datos por ADO/ADODB/OLE DB
'Constante          Valor
'-----------------------------------------------------------------
Global Const Car_Visto = 251
'-----------------------------------------------------------------
Global Const adTrue = -1
Global Const adFalse = 0
'-----------------------------------------------------------------
'Colores Generales
Global Const Negro = 0
Global Const Azul = 1
Global Const Verde = 2
Global Const Aguamarina = 3
Global Const Rojo = 4
Global Const Fucsia = 5
Global Const Amarillo = 6
Global Const Blanco = 7
Global Const Gris = 8
Global Const Azul_Claro = 9
Global Const Verde_Claro = 10
Global Const Aguamarina_Claro = 11
Global Const Magenta = 11
Global Const Rojo_Claro = 12
Global Const Fucsia_Claro = 13
Global Const Amarillo_Claro = 14
Global Const Blanco_Brillante = 15
'-----------------------------------
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
Global Const TipoCourier = "Courier"
Global Const TipoCourierNew = "Courier New"
Global Const TipoTimes = "Times New Roman"
Global Const TipoTerminal = "Terminal"
Global Const TipoSystem = "System"
Global Const TipoGeorgia = "Georgia"
Global Const TipoCalibri = "Calibri"
Global Const TipoHelvetica = "Helvetica"
Global Const TipoHelveticaBold = "Helvetica-Bold"
Global Const TipoTahoma = "Tahoma"
Global Const TipoVerdana = "Verdana"
Global Const TipoWingdings = "Wingdings"
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
Global Const MascaraCtas = "#.#.##.##.##.###"
Global Const FormatoCtas = "#.#.##.##.##.###"
Global Const LimpiarCtas = " . .  .  .  .   "
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
'-----------------------------------------
Global Const MaxVect = 80



