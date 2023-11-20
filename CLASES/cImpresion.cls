VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cImpresion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


'----------------------------------------------------------------------------------------------------------
' CLASE         : cImpresion
'----------------------------------------------------------------------------------------------------------
' FECHA         : 23 de Agosto del 2016
'
' DESARROLLADO  : Walter Jalil Vaca Prieto
'
' DESCRIPCION   : Esta clase sirve para imprimir la misma informacion en tres tipos de objeto de impresion,
'                 estos objeto son: Printer, Picture y PDF.
'
'                 El objetivo de la creacion de esta clase es para poder migrar a punto NET son problemas y
'                 para futuro poder imprimir en cualquier tipo de objeto lo que deseamos.
'----------------------------------------------------------------------------------------------------------
Option Explicit

Public Enum eTipoImpresion
       c_Printer = 1
       c_Picture = 2
       c_PDF = 3
End Enum

Public Enum eTipoDeCodigo128
       cC128_A = 1
       cC128_B = 2
       cC128_C = 3
End Enum

Private Type colCodigoBarra128
        Codigo  As String
        Patron  As String
End Type

'Constantes
 Private Const cCM = 72 / 2.54
 Private Const cTwip = 566.9291338583  '567

'Objetos/Clases
 Public ObjPrint As Object
 Private ObjPDF As mjwPDF

'OBJETOS PUBLICOS
'----------------
'Byte
 Public dOrientacionPagina As Byte
 
'Integer
 Public cPrinter As Integer
 Public cPicture As Integer
 Public cPDF As Integer
 Public dPorteLetra As Integer

'Long
 
'Double
 
'Single
 Public dAnchoPapel As Single
 Public dAltoPapel As Single
 Public dNoLineas As Single
 
'String
 Public vPatronText As String
 Public dTipoLetra As String
 Public dNombreArchivo As String
 Public dTituloArchivo As String
 Public dStrgFormatoCampo As String

'Boolean
 Public dPaginaA4 As Boolean
 Public dVerDocumento As Boolean
 Public dEsCampoCorto As Boolean
 Public vValorNegativo As Boolean
 
'Tipo Personalizados
 Public dtipoImpresion As eTipoImpresion

'OBJETOS PRIVADOS
'----------------
'Byte
 Private vOrientacion As Byte

'Integer
 Private vTipoImp As Integer
 Private vPorteLetra As Integer
 Private vCantidadCampos As Integer
 Private vSegundaPagina As Integer
 Public vPaginaNo As Integer
 
'Long
 Private vColor As Long
 Private vColorLetra As Long
 Private vColorFondo As Long
 Private vFondoLetra As Long
 
'Double
 Private vPosX As Double
 Private vPosY As Double
 Private vPosXo As Double
 Private vPosYo As Double
 Private vPosXf As Double
 Private vPosYf As Double
 Private vLargo As Double
 Private vAlto1 As Double
 Private vAlto2 As Double
 Private vGrosor As Double
 Private vAncho1 As Double
 Private vAncho2 As Double
 Private vAnchoTexto As Double
 Private vAltoLetra As Double
 Private vDistancia As Double
 Private vSumaTotalAncho As Double
 Private vPosXCampo As Double
 Private vAncho() As Double
 
'Single
 
'String
 Private vBF As String
 Private vCh As String
 Private vTexto As String
 Private vTipoLetra As String
 Private vNombreArchivo As String
 Private vTituloArchivo As String
 Private vStrgAnchoVariable As String
 Private vAnchoDeCampo() As String

'Boolean
 Private vPaginaA4 As Boolean
 Private vEsCampoCorto As Boolean
 Private vVerDocumento As Boolean
 
'Tipo Personalizados
 Private colCB128(110) As colCodigoBarra128
 
'----------------------------------------------------------------------------------------------------------
' INICIALIZACION DE LA CLASE DE IMPRESIONES
'----------------------------------------------------------------------------------------------------------
Private Sub Class_Initialize()
' crea la colección cuando se crea la clase
  cPrinter = 1
  cPicture = 2
  cPDF = 3
 
'----------------------------------------------------------------------------------------------------------
' Crea la tabla con los patrones para cada símbolo. El patrón indica, alternativamente, líneas y espacios
'----------------------------------------------------------------------------------------------------------

colCB128(0).Codigo = " "
colCB128(0).Patron = "212222"

colCB128(1).Codigo = "!"
colCB128(1).Patron = "222122"

colCB128(2).Codigo = Chr(34)  ' Comillas dobles
colCB128(2).Patron = "222221"

colCB128(3).Codigo = "#"
colCB128(3).Patron = "121223"
 
colCB128(4).Codigo = "$"
colCB128(4).Patron = "121322"
 
colCB128(5).Codigo = "%"
colCB128(5).Patron = "131222"
 
colCB128(6).Codigo = "&"
colCB128(6).Patron = "122213"
 
colCB128(7).Codigo = "'"
colCB128(7).Patron = "122312"
 
colCB128(8).Codigo = "("
colCB128(8).Patron = "132212"
 
colCB128(9).Codigo = ")"
colCB128(9).Patron = "221213"
 
colCB128(10).Codigo = "*"
colCB128(10).Patron = "221312"
 
colCB128(11).Codigo = "+"
colCB128(11).Patron = "231212"
 
colCB128(12).Codigo = ","
colCB128(12).Patron = "112232"
 
colCB128(13).Codigo = "-"
colCB128(13).Patron = "122132"
 
colCB128(14).Codigo = "."
colCB128(14).Patron = "122231"
 
colCB128(15).Codigo = "/"
colCB128(15).Patron = "113222"
 
colCB128(16).Codigo = "0"
colCB128(16).Patron = "123122"
 
colCB128(17).Codigo = "1"
colCB128(17).Patron = "123221"
 
colCB128(18).Codigo = "2"
colCB128(18).Patron = "223211"
 
colCB128(19).Codigo = "3"
colCB128(19).Patron = "221132"

colCB128(20).Codigo = "4"
colCB128(20).Patron = "221231"

colCB128(21).Codigo = "5"
colCB128(21).Patron = "213212"

colCB128(22).Codigo = "6"
colCB128(22).Patron = "223112"

colCB128(23).Codigo = "7"
colCB128(23).Patron = "312131"

colCB128(24).Codigo = "8"
colCB128(24).Patron = "311222"

colCB128(25).Codigo = "9"
colCB128(25).Patron = "321122"

colCB128(26).Codigo = ":"
colCB128(26).Patron = "321221"

colCB128(27).Codigo = ";"
colCB128(27).Patron = "312212"

colCB128(28).Codigo = "<"
colCB128(28).Patron = "322112"
 
colCB128(29).Codigo = "="
colCB128(29).Patron = "322211"
 
colCB128(30).Codigo = ">"
colCB128(30).Patron = "212123"
 
colCB128(31).Codigo = "?"
colCB128(31).Patron = "212321"
 
colCB128(32).Codigo = "@"
colCB128(32).Patron = "232121"
 
colCB128(33).Codigo = "A"
colCB128(33).Patron = "111323"
 
colCB128(34).Codigo = "B"
colCB128(34).Patron = "131123"
 
colCB128(35).Codigo = "C"
colCB128(35).Patron = "131321"
 
colCB128(36).Codigo = "D"
colCB128(36).Patron = "112313"
 
colCB128(37).Codigo = "E"
colCB128(37).Patron = "132113"
 
colCB128(38).Codigo = "F"
colCB128(38).Patron = "132311"
 
colCB128(39).Codigo = "G"
colCB128(39).Patron = "211313"
 
colCB128(40).Codigo = "H"
colCB128(40).Patron = "231113"
 
colCB128(41).Codigo = "I"
colCB128(41).Patron = "231311"
 
colCB128(42).Codigo = "J"
colCB128(42).Patron = "112133"
 
colCB128(43).Codigo = "K"
colCB128(43).Patron = "112331"
 
colCB128(44).Codigo = "L"
colCB128(44).Patron = "132131"
 
colCB128(45).Codigo = "M"
colCB128(45).Patron = "113123"
 
colCB128(46).Codigo = "N"
colCB128(46).Patron = "113321"
 
colCB128(47).Codigo = "O"
colCB128(47).Patron = "133121"
 
colCB128(48).Codigo = "P"
colCB128(48).Patron = "313121"
 
colCB128(49).Codigo = "Q"
colCB128(49).Patron = "211331"
 
colCB128(50).Codigo = "R"
colCB128(50).Patron = "231131"
 
colCB128(51).Codigo = "S"
colCB128(51).Patron = "213113"
 
colCB128(52).Codigo = "T"
colCB128(52).Patron = "213311"
 
colCB128(53).Codigo = "U"
colCB128(53).Patron = "213131"
 
colCB128(54).Codigo = "V"
colCB128(54).Patron = "311123"
 
colCB128(55).Codigo = "W"
colCB128(55).Patron = "311321"
 
colCB128(56).Codigo = "X"
colCB128(56).Patron = "331121"
 
colCB128(57).Codigo = "Y"
colCB128(57).Patron = "312113"
 
colCB128(58).Codigo = "Z"
colCB128(58).Patron = "312311"
 
colCB128(59).Codigo = "["
colCB128(59).Patron = "332111"
 
colCB128(60).Codigo = "\"
colCB128(60).Patron = "314111"
 
colCB128(61).Codigo = "]"
colCB128(61).Patron = "221411"
 
colCB128(62).Codigo = "^"
colCB128(62).Patron = "431111"
 
colCB128(63).Codigo = "_"
colCB128(63).Patron = "111224"
 
colCB128(64).Codigo = "'"
colCB128(64).Patron = "111422"
 
colCB128(65).Codigo = "a"
colCB128(65).Patron = "121124"
 
colCB128(66).Codigo = "b"
colCB128(66).Patron = "121421"
 
colCB128(67).Codigo = "c"
colCB128(67).Patron = "141122"
 
colCB128(68).Codigo = "d"
colCB128(68).Patron = "141221"
 
colCB128(69).Codigo = "e"
colCB128(69).Patron = "112214"
 
colCB128(70).Codigo = "f"
colCB128(70).Patron = "112412"
 
colCB128(71).Codigo = "g"
colCB128(71).Patron = "122114"
 
colCB128(72).Codigo = "h"
colCB128(72).Patron = "122411"
 
colCB128(73).Codigo = "i"
colCB128(73).Patron = "142112"
 
colCB128(74).Codigo = "j"
colCB128(74).Patron = "142211"
 
colCB128(75).Codigo = "k"
colCB128(75).Patron = "241211"
 
colCB128(76).Codigo = "l"
colCB128(76).Patron = "221114"
 
colCB128(77).Codigo = "m"
colCB128(77).Patron = "413111"
 
colCB128(78).Codigo = "n"
colCB128(78).Patron = "241112"
 
colCB128(79).Codigo = "o"
colCB128(79).Patron = "134111"
 
colCB128(80).Codigo = "p"
colCB128(80).Patron = "111242"
 
colCB128(81).Codigo = "q"
colCB128(81).Patron = "121142"
 
colCB128(82).Codigo = "r"
colCB128(82).Patron = "121241"
 
colCB128(83).Codigo = "s"
colCB128(83).Patron = "114212"
 
colCB128(84).Codigo = "t"
colCB128(84).Patron = "124112"
 
colCB128(85).Codigo = "u"
colCB128(85).Patron = "124211"
 
colCB128(86).Codigo = "v"
colCB128(86).Patron = "411212"

colCB128(87).Codigo = "w"
colCB128(87).Patron = "421112"
 
colCB128(88).Codigo = "x"
colCB128(88).Patron = "421211"
 
colCB128(89).Codigo = "y"
colCB128(89).Patron = "212141"
 
colCB128(90).Codigo = "z"
colCB128(90).Patron = "214121"
 
colCB128(91).Codigo = "{"
colCB128(91).Patron = "412121"
 
colCB128(92).Codigo = "|"
colCB128(92).Patron = "111143"
 
colCB128(93).Codigo = "}"
colCB128(93).Patron = "111341"
 
colCB128(94).Codigo = "~"
colCB128(94).Patron = "131141"
 
colCB128(95).Codigo = ""
colCB128(95).Patron = "114113"
 
colCB128(96).Codigo = ""
colCB128(96).Patron = "114311"
 
colCB128(97).Codigo = ""
colCB128(97).Patron = "411113"
 
colCB128(98).Codigo = ""
colCB128(98).Patron = "411311"
 
colCB128(99).Codigo = ""
colCB128(99).Patron = "113141"
 
colCB128(100).Codigo = ""
colCB128(100).Patron = "114131"
 
colCB128(101).Codigo = ""
colCB128(101).Patron = "311141"
 
colCB128(102).Codigo = ""
colCB128(102).Patron = "411131"
 
colCB128(103).Codigo = ""
colCB128(103).Patron = "211412"
 
colCB128(104).Codigo = ""
colCB128(104).Patron = "211214"
 
colCB128(105).Codigo = ""
colCB128(105).Patron = "211232"
 
colCB128(106).Codigo = ""
colCB128(106).Patron = "2331112"
End Sub

Private Sub xxx(ByVal dXXX As Integer)
    Select Case vTipoImp
      Case eTipoImpresion.c_PDF
            
      Case eTipoImpresion.c_Picture
            
      Case eTipoImpresion.c_Printer
 
    End Select
End Sub

Private Sub SetearTipoDeLetra()
    Select Case vTipoLetra
      Case TipoArial:     ObjPrint.PDFSetFontName FONT_ARIAL
      Case TipoCourier:   ObjPrint.PDFSetFontName FONT_COURIER
      Case TipoTimes:     ObjPrint.PDFSetFontName FONT_TIMES
      Case TipoSymbol:    ObjPrint.PDFSetFontName FONT_SYMBOL
      Case TipoHelvetica: ObjPrint.PDFSetFontName FONT_HELVETICA
      Case TipoVerdana:   ObjPrint.PDFSetFontName FONT_VERDANA
      Case TipoTerminal:  ObjPrint.PDFSetFontName FONT_ZAPFDINGBATS
      Case Else:          ObjPrint.PDFSetFontName FONT_ARIAL
    End Select
End Sub

Public Property Let TipoImpresion(ByVal dTipoImp As eTipoImpresion)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.tipoImpresion = 1,2,3
    vTipoImp = dTipoImp
End Property

Public Property Get anchoTexto(ByVal dTexto As String) As Double
    Select Case vTipoImp
      Case eTipoImpresion.c_PDF
           anchoTexto = Round(ObjPrint.PDFGetTextWidth(dTexto), 2)
      Case eTipoImpresion.c_Picture, eTipoImpresion.c_Printer
           anchoTexto = Round(ObjPrint.TextWidth(dTexto), 2)
    End Select
End Property

Public Property Get altoTexto(ByVal dTexto As String) As Double
    Select Case vTipoImp
      Case eTipoImpresion.c_PDF
           ObjPrint.PDFSetFontSize vPorteLetra
           vAltoLetra = Round(ObjPrint.PDFTextHeight / cTwip, 2)
          'MsgBox "(" & vPorteLetra & ") altoTexto: " & dTexto & " > " & ObjPrint.PDFTextHeight & vbCrLf & vAltoLetra
           altoTexto = vAltoLetra
      Case eTipoImpresion.c_Picture, eTipoImpresion.c_Printer
           altoTexto = Round(ObjPrint.TextHeight(dTexto), 2)
    End Select
End Property

Public Sub iniciaImpresion()
'ByVal datosDocumento As Tipo_Impreson
Dim txtTime As String
Dim sRutaDelPDF As String
 
   'Variables Locales del tipo de impresion
    RatonReloj
    txtTime = NumEmpresa & "_" & Format(Time$, "hhmmss")
    vColorLetra = Negro
    vTipoImp = tPrint.TipoImpresion
    vNombreArchivo = tPrint.NombreArchivo
    vTituloArchivo = tPrint.TituloArchivo
    vTipoLetra = tPrint.TipoLetra
    vPorteLetra = tPrint.PorteLetra
    vPaginaA4 = tPrint.PaginaA4
    vOrientacion = tPrint.OrientacionPagina
    vEsCampoCorto = tPrint.EsCampoCorto
    vVerDocumento = tPrint.VerDocumento
    vNombreArchivo = Replace(vNombreArchivo, ".", "-")
    sRutaDelPDF = RutaSysBases & "\TEMP\" & vNombreArchivo & ".pdf"

    If vOrientacion = 0 Then vOrientacion = 1
    If vPorteLetra <= 0 Then vPorteLetra = 5
    If Len(vTituloArchivo) <= 1 Then vTituloArchivo = "Documento RIDE de " & txtTime
    If Len(vNombreArchivo) <= 1 Then vNombreArchivo = "PDF DiskCover " & txtTime
    If SetNombrePRN = Impresota_PDF Then vTipoImp = eTipoImpresion.c_PDF
    vPaginaNo = 1
   'MsgBox SetNombrePRN & vbCrLf & vTipoImp & vbCrLf & tPrint.TipoImpresion
    Select Case vTipoImp
      Case eTipoImpresion.c_PDF
          'Inicializamos el tipo de impresion para generar PDF
           Set ObjPDF = New mjwPDF
           Set ObjPrint = ObjPDF
           
           ObjPrint.PDFTitle = vTituloArchivo
          'ObjPrint.PDFFileName = "\\192.168.27.2\Sistemas\SYSBASES\TEMP\" & vNombreArchivo & ".pdf"
           ObjPrint.PDFFileName = sRutaDelPDF
           ObjPrint.PDFLoadAfm = RutaSistema & "\FONTSPDF"
           ObjPrint.PDFSetUnit = UNIT_CM
           ObjPrint.PDFSetLayoutMode = LAYOUT_DEFAULT
           ObjPrint.PDFFormatPage = FORMAT_A4
           If vOrientacion = 1 Then
              ObjPrint.PDFOrientation = ORIENT_PORTRAIT
           Else
              ObjPrint.PDFOrientation = ORIENT_PAYSAGE
           End If
            
           ObjPrint.PDFView = vVerDocumento
           ObjPrint.PDFBeginDoc
           ObjPrint.PDFSetBookmark " "
           SetearTipoDeLetra
           ObjPrint.PDFSetTextColor = vColorLetra
           ObjPrint.PDFSetFontSize vPorteLetra
           
          'LimiteAlto = 72 / 2.54
          'LimiteAncho = 72 / 2.54
           dAnchoPapel = Redondear(ObjPrint.PDFGetPageWidth / cCM, 2)
           dAltoPapel = Redondear(ObjPrint.PDFGetPageHeight / cCM, 2)
           LimiteAlto = dAltoPapel
           LimiteAncho = dAnchoPapel
           AnchoPapel = Redondear(LimiteAncho, 2)         'Ancho de impresion del papel
           SetPapelAncho = Redondear(LimiteAncho, 2)
           SetPapelLargo = Redondear(LimiteAlto, 2)
           LimiteAlto = Redondear(LimiteAlto - 1, 2)
           LimiteAncho = Redondear(LimiteAncho - 0.5, 2)  'Limite de impresión a lo ancho
          'printCuadro 1, 1, LimiteAncho, LimiteAlto, Rojo
      Case eTipoImpresion.c_Picture
          'Determinamos el control Generico de Visual Basic para graficos PICTURE
           ObjPrint.Cls
      Case eTipoImpresion.c_Printer
          'Determinamos el control Generico de Visual Basic para impresoras PRINTER
           Set ObjPrint = Printer
           ObjPrint.ScaleMode = vbCentimeters             'Escala de centimetros
           ObjPrint.PaperSize = SetPapelPRN
           'Printer.PaperSize = SetPapelPRN
          'Seteo la Impresion en Centimetros
           If SetPapelPRN > 0 Then ObjPrint.PaperSize = SetPapelPRN Else ObjPrint.PaperSize = vbPRPSLetter
          'MsgBox "Tipo Papel: " & ObjPrint.PaperSize
           
          'Orientacion_Pagina normal de la hoja y Limite de impresión a lo largo
           If vOrientacion = 1 Then
              ObjPrint.Orientation = vbPRORPortrait  'Vertical
           Else
              ObjPrint.Orientation = vbPRORLandscape 'Horizontal
           End If
           ObjPrint.DrawWidth = 1                                'Ancho de la línea
           ObjPrint.FontName = vTipoLetra                        'Tipo de letra en todo el sistema
           ObjPrint.FontSize = vPorteLetra                       'Porte de la letra default
           ObjPrint.ForeColor = vColorLetra                      'Color de la letra default
           ObjPrint.FontBold = False
           ObjPrint.FontItalic = False
           ObjPrint.FontUnderline = False
           LimiteAlto = Redondear(ObjPrint.ScaleHeight - 1, 2)
           LimiteAncho = Redondear(ObjPrint.ScaleWidth - 0.5, 2) 'Limite de impresión a lo ancho
           AnchoPapel = Redondear(ObjPrint.ScaleWidth, 2)        'Ancho de impresion del papel
           SetPapelAncho = Redondear(ObjPrint.ScaleWidth, 2)
           SetPapelLargo = Redondear(ObjPrint.ScaleHeight, 2)
    End Select
    SetAnchoCampos ObjPrint, vEsCampoCorto
    RatonReloj
End Sub

Public Sub finalizaImpresion()
    Select Case vTipoImp
      Case eTipoImpresion.c_PDF
           ObjPrint.PDFEndDoc
      Case eTipoImpresion.c_Picture
           ObjPrint.Cls
      Case eTipoImpresion.c_Printer
           ObjPrint.EndDoc
    End Select
    RatonNormal
End Sub

Public Sub paginaNueva()
    vPaginaNo = vPaginaNo + 1
    Select Case vTipoImp
      Case eTipoImpresion.c_PDF
           ObjPrint.PDFEndPage
           ObjPrint.PDFNewPage
      Case eTipoImpresion.c_Picture
           ObjPrint.Cls
      Case eTipoImpresion.c_Printer
           ObjPrint.NewPage
    End Select
End Sub

Private Function anchoTipoVariable(ByVal dVariable As Variant) As String
    Select Case VarType(dVariable)
      Case vbBoolean:  vCh = CadBoolean
      Case vbByte:     vCh = CadByte
      Case vbInteger:  vCh = CadInteger
      Case vbLong:     vCh = CadLong
      Case vbSingle:   vCh = CadSingle
      Case vbDouble:   vCh = CadDouble
      Case vbCurrency: vCh = CadCurrency
      Case vbDate:     vCh = CadDate
      Case vbString:   vCh = String$(Len(dVariable), "X") & " "
      Case Else:       vCh = String$(15, "X") & " "
    End Select
    If vCh = "" Or vCh = "." Then vCh = " "
    anchoTipoVariable = vCh
End Function

Private Function anchoTipoFields(ByVal dAdoTipo As ADODB.Field) As String
    Select Case dAdoTipo.Type
      Case TadBoolean:        vCh = CadBoolean
      Case TadByte:           vCh = CadByte
      Case TadInteger:        vCh = CadInteger
      Case TadLong:           vCh = CadLong
      Case TadSingle:         vCh = CadSingle
      Case TadDouble:         vCh = CadDouble
      Case TadCurrency:       vCh = CadCurrency
      Case TadDate, TadDate1: vCh = CadDate
      Case TadString:         If dAdoTipo.DefinedSize <= 40 Then vCh = String$(dAdoTipo.DefinedSize, "X") & " " Else vCh = String$(40, "X") & " "
      Case Else:              vCh = String$(15, "X") & " "
    End Select
    If vCh = "" Or vCh = "." Then vCh = " "
   ' MsgBox "anchoTipoFields"
    anchoTipoFields = vCh
End Function

Private Function anchoVariable(ByVal dVariable As Variant, _
                               Optional ByVal dCantDecimales As Byte) As Double
    vValorNegativo = False
    vStrgAnchoVariable = anchoTipoVariable(dVariable)
    If dCantDecimales < 2 Then dCantDecimales = 2
    Select Case VarType(dVariable)
      Case vbBoolean
           dStrgFormatoCampo = Format$(dVariable, "Si/No")
      Case vbByte, vbInteger, vbLong
           If Val(dVariable) < 0 Then vValorNegativo = True
           dStrgFormatoCampo = Format$(dVariable, "##0")
      Case vbSingle
           If Val(dVariable) < 0 Then vValorNegativo = True
           dStrgFormatoCampo = Format$(dVariable, "##0.00%")
      Case vbDouble, vbCurrency
           If Val(dVariable) < 0 Then vValorNegativo = True
           dStrgFormatoCampo = Format$(dVariable, "#,##0." & String(dCantDecimales, "0"))
      Case vbDate
           dStrgFormatoCampo = Format$(dVariable, FormatoFechas)
      Case vbString
           dStrgFormatoCampo = dVariable
      Case Else
           dStrgFormatoCampo = CStr(dVariable)
    End Select
    vAncho1 = anchoTexto(vStrgAnchoVariable)
    vAncho2 = anchoTexto(dStrgFormatoCampo)
    If vAncho1 > vAncho2 Then vDistancia = (vAncho1 - vAncho2) + 0.05 Else vDistancia = 0.05
    If vDistancia <= 0 Then vDistancia = 0.05
    anchoVariable = vDistancia
End Function

Public Function anchoFields(ByVal dAdoTipo As ADODB.Field, _
                            Optional ByVal dCantDecimales As Byte) As Double
    
    If dCantDecimales < 2 Then dCantDecimales = 2
    vStrgAnchoVariable = anchoTipoFields(dAdoTipo)
    vValorNegativo = False
    vDistancia = 0
    Select Case dAdoTipo.Type
      Case TadBoolean
           If dAdoTipo Then dStrgFormatoCampo = "**" Else dStrgFormatoCampo = " "
          'dStrgFormatoCampo = Format$(dAdoTipo, "Si/No")
           vAncho1 = 0
           vAncho2 = 0
      Case TadByte
           If Val(dAdoTipo) < 0 Then vValorNegativo = True
           dStrgFormatoCampo = Format$(dAdoTipo, "##0")
           If Val(dAdoTipo) = 0 Then dStrgFormatoCampo = " "
           vAncho1 = anchoTexto(vStrgAnchoVariable)
           vAncho2 = anchoTexto(dStrgFormatoCampo)
      Case TadInteger
           If Val(dAdoTipo) < 0 Then vValorNegativo = True
           dStrgFormatoCampo = Format$(dAdoTipo, "##0")
           If Val(dAdoTipo) = 0 Then dStrgFormatoCampo = " "
           vAncho1 = anchoTexto(vStrgAnchoVariable)
           vAncho2 = anchoTexto(dStrgFormatoCampo)
      Case TadLong
           If Val(dAdoTipo) < 0 Then vValorNegativo = True
           dStrgFormatoCampo = Format$(dAdoTipo, "##0")
           If Val(dAdoTipo) = 0 Then dStrgFormatoCampo = " "
           vAncho1 = anchoTexto(vStrgAnchoVariable)
           vAncho2 = anchoTexto(dStrgFormatoCampo)
      Case TadSingle
           dStrgFormatoCampo = Format$(dAdoTipo, "##0.00%")
           If Val(dAdoTipo) < 0 Then vValorNegativo = True
           If Val(dAdoTipo) = 0 Then dStrgFormatoCampo = " "
           vAncho1 = anchoTexto(vStrgAnchoVariable)
           vAncho2 = anchoTexto(dStrgFormatoCampo)
      Case TadDouble
           dStrgFormatoCampo = Format$(dAdoTipo, "#,##0." & String(dCantDecimales, "0"))
           If Val(dAdoTipo) < 0 Then vValorNegativo = True
           If Val(dAdoTipo) = 0 Then dStrgFormatoCampo = " "
           vAncho1 = anchoTexto(vStrgAnchoVariable)
           vAncho2 = anchoTexto(dStrgFormatoCampo)
      Case TadCurrency
           dStrgFormatoCampo = Format$(dAdoTipo, "#,##0." & String(dCantDecimales, "0"))
           If Val(dAdoTipo) < 0 Then vValorNegativo = True
           If Val(dAdoTipo) = 0 Then dStrgFormatoCampo = " "
           vAncho1 = anchoTexto(vStrgAnchoVariable)
           vAncho2 = anchoTexto(dStrgFormatoCampo)
      Case TadDate, TadDate1
           dStrgFormatoCampo = Format$(dAdoTipo, FormatoFechas)
''           vAncho1 = anchoTexto(vStrgAnchoVariable)
''           vAncho2 = anchoTexto(dStrgFormatoCampo)
           vAncho1 = 0
           vAncho2 = 0
      Case TadString
           dStrgFormatoCampo = dAdoTipo
           vAncho1 = 0
           vAncho2 = 0
      Case Else
           dStrgFormatoCampo = CStr(dAdoTipo)
           vAncho1 = 0
           vAncho2 = 0
    End Select
    If vAncho1 > vAncho2 Then vDistancia = (vAncho1 - vAncho2) - 0.02
    anchoFields = Round(vDistancia, 2)
End Function

Public Sub printTexto(ByVal dPosXo As Double, _
                      ByVal dPosYo As Double, _
                      ByVal dTexto As String, _
                      Optional ByVal dPorteLetra As Integer, _
                      Optional ByVal dTipoJustTexto As String = "I", _
                      Optional ByVal dAnchoTexto As Double, _
                      Optional ByVal dColorFondo As Long = Blanco)
Dim CantTexto As String
Dim Caracter As String
Dim AnchoText As Double
Dim PosTexto As Long
Dim PosYAux As Double
Dim ConFondo As Boolean
Dim tColorLetra As Long
 dNoLineas = 0
 If ((dPosXo > 0) And (dPosYo > 0) And (Len(dTexto) > 0)) And dPosYo <= LimiteAlto Then
    'MsgBox dPosXo & "=" & dTexto & "  ..."
    If dAnchoTexto = 0 And dTipoJustTexto = "C" Then dAnchoTexto = Redondear(SetPapelAncho, 2)
    vPosXo = Redondear(dPosXo, 4)
    vPosYo = Redondear(dPosYo, 4)
    If vPosXo < 0 Then vPosXo = 0.01
    If vPosYo < 0 Then vPosYo = 0.01
    vTexto = dTexto
    If vTexto = Ninguno Then vTexto = " "
    tColorLetra = vColorLetra
    If dColorFondo < 0 Then dColorFondo = Blanco
    If dColorFondo <> Blanco Then
       ConFondo = True
       vColorLetra = Rojo
    Else
       ConFondo = False
    End If
    vColorFondo = dColorFondo
    If dPorteLetra > 0 Then vPorteLetra = dPorteLetra
    Select Case vTipoImp
      Case eTipoImpresion.c_PDF
           'vPosXo = vPosXo - 0.29
           ObjPrint.PDFSetFontSize vPorteLetra
           ObjPrint.PDFSetTextColor = vColorLetra
           If ConFondo Then ObjPrint.PDFSetFill = vColorFondo
           vPosXo = vPosXo - 1.2
           vPosYo = vPosYo - 0.1
      Case eTipoImpresion.c_Picture
           ObjPrint.FontSize = vPorteLetra
           ObjPrint.ForeColor = vColorLetra
           If ConFondo Then ObjPrint.FillColor = vColorFondo
      Case eTipoImpresion.c_Printer
           ObjPrint.FontSize = vPorteLetra
           ObjPrint.ForeColor = vColorLetra
           If ConFondo Then ObjPrint.FillColor = vColorFondo
           vPosXo = vPosXo - 0.2
           vPosYo = vPosYo - 0.3
    End Select
    vAltoLetra = altoTexto(dTexto)
    If dAnchoTexto > 0 Then vAnchoTexto = dAnchoTexto Else vAnchoTexto = dAnchoPapel - 1
   'Si hay que presentar el fondo de la letra con un cuadro
   'vAnchoTexto
    'If ConFondo Then printCuadroLinea vPosXo, vPosYo - 0.3, 3, vPosYo + vAltoLetra + 0.3, vColorFondo, "BF"
    vDistancia = 0  'Izquierda: Por default
    Select Case dTipoJustTexto
      Case "C" 'Centrado
           vDistancia = ((vAnchoTexto - vPosXo) / 2) - (anchoTexto(vTexto) / 2)
           If vDistancia < 0 Then vDistancia = 0
      Case "D" 'Derecha
           If dAnchoTexto > vAnchoTexto Then vDistancia = dAnchoTexto - vAnchoTexto Else vDistancia = 0
      Case Else
           'dAnchoTexto = 0.01
    End Select
    PosYAux = vPosYo
    AnchoText = 0
    If dAnchoTexto > 0 Then
      'MsgBox dAnchoTexto & vbCrLf & vTexto
       Do While Len(vTexto) > 0
          AnchoText = anchoTexto(vTexto)
          If AnchoText > dAnchoTexto Then
             CantTexto = ""
             PosTexto = 1
             AnchoText = 0
             Do While AnchoText <= dAnchoTexto
                CantTexto = CantTexto & MidStrg(vTexto, PosTexto, 1)
                AnchoText = anchoTexto(CantTexto)
                PosTexto = PosTexto + 1
                If PosTexto > Len(vTexto) Then AnchoText = dAnchoTexto + 1
             Loop
             Caracter = MidStrg(CantTexto, Len(CantTexto), 1)
             If Caracter = " " Then
                CantTexto = TrimStrg(CantTexto)
             Else
                PosTexto = Len(CantTexto)
                Caracter = MidStrg(CantTexto, PosTexto, 1)
                Do While Caracter <> " " And PosTexto > 1
                   Caracter = MidStrg(CantTexto, PosTexto, 1)
                   PosTexto = PosTexto - 1
                Loop
                PosTexto = PosTexto + 1
                CantTexto = TrimStrg(MidStrg(CantTexto, 1, PosTexto))
             End If
          Else
             CantTexto = vTexto
          End If
          Select Case vTipoImp
            Case eTipoImpresion.c_PDF
                 ObjPrint.PDFTextOut CantTexto, vPosXo + vDistancia, vPosYo
                'MsgBox vPosXo & vbCrLf & vDistancia & vbCrLf & vPosYo & vbCrLf & CantTexto
            Case eTipoImpresion.c_Picture, eTipoImpresion.c_Printer
                 ObjPrint.CurrentX = vPosXo + vDistancia
                 ObjPrint.CurrentY = vPosYo
                 ObjPrint.Print CantTexto
                'MsgBox CantTexto
          End Select
         'Quitamos el segmento del texto impreso
          If Len(vTexto) >= Len(CantTexto) Then
             vTexto = TrimStrg(MidStrg(vTexto, Len(CantTexto) + 1, Len(vTexto)))
          Else
             vTexto = ""
          End If
          vPosYo = vPosYo + vAltoLetra - 0.05   ' 0.3
       Loop
    Else
      ' MsgBox vTexto & vbCrLf & vPosXo & vbCrLf & vPosYo & vbCrLf & vDistancia
       If vPosXo <= 0 Then
          vPosXo = 1.2
          vDistancia = 0
       End If
       Select Case vTipoImp
         Case eTipoImpresion.c_PDF
             'MsgBox vPosXo & vbCrLf & vDistancia & vbCrLf & vPosYo & vbCrLf & vTexto
              ObjPrint.PDFTextOut vTexto, vPosXo + vDistancia, vPosYo
         Case eTipoImpresion.c_Picture, eTipoImpresion.c_Printer
              ObjPrint.CurrentX = vPosXo + vDistancia
              ObjPrint.CurrentY = vPosYo
              ObjPrint.Print vTexto
             'MsgBox vTexto
       End Select
    End If
    If PosYAux <> vPosYo Then dNoLineas = vPosYo - vAltoLetra + 0.05
    vColorLetra = tColorLetra
 End If
End Sub

Public Sub printTextoAngulo(PosPaintX As Single, _
                            PosPaintY As Single, _
                            Angulo As Integer, _
                            AnchoDelTexto As Single, _
                            TextoAEscribir As String, _
                            Optional ByVal dPorteLetra As Integer)
Dim NewFont As Long
Dim OldFont As Long
Dim IText As Integer
Dim Xo1 As Double
Dim Yo1 As Double
'Formalizamos tipo de impresion
If ((PosPaintX > 0) And (PosPaintY > 0) And (Angulo > 0) And (TextoAEscribir <> "")) Then
   Xo1 = CDbl(PosPaintX)
   Yo1 = CDbl(PosPaintY)
   Lineas_Impresas = 0
   If TextoAEscribir = Ninguno Then TextoAEscribir = " "
   If LimiteAncho <= 0 Then LimiteAncho = 20
   Select Case vTipoImp
     Case eTipoImpresion.c_PDF
          If dPorteLetra > 0 Then
             vPorteLetra = dPorteLetra
             ObjPrint.PDFSetFontSize vPorteLetra
          End If
          ObjPrint.PDFSetTextColor = vColorLetra
          ObjPrint.PDFRotationText Xo1, Yo1, TextoAEscribir, Angulo
     Case eTipoImpresion.c_Picture, eTipoImpresion.c_Printer
          ObjPrint.ForeColor = vColorLetra
          If dPorteLetra > 0 Then ObjPrint.FontSize = dPorteLetra
          NewFont = CreateFont(dPorteLetra, AnchoDelTexto, Angulo * 10, 0, FF_DONTCARE, 0, 0, 0, ANSI_CHARSET, OUT_TT_ONLY_PRECIS, CLIP_LH_ANGLES Or CLIP_DEFAULT_PRECIS, PROOF_QUALITY, TRUETYPE_FONTTYPE, ObjPrint.FontName)
          OldFont = SelectObject(ObjPrint.hdc, NewFont)
          ObjPrint.CurrentX = Xo1
          ObjPrint.CurrentY = Yo1
          ObjPrint.Print TextoAEscribir
         'set the old font back
          NewFont = SelectObject(ObjPrint.hdc, OldFont)
         'Clean up
          DeleteObject NewFont
   End Select
End If
End Sub

Public Sub printVariable(ByVal dPosX As Double, _
                         ByVal dPosY As Double, _
                         ByVal dVariable As Variant, _
                         Optional ByVal dPorteLetra As Integer, _
                         Optional ByVal dColorTexto As Long = Negro, _
                         Optional ByVal dColorFondo As Long = Blanco, _
                         Optional ByVal dCantDecimales As Byte)
If (dPosX > 0) And (dPosY > 0) Then
   If dPorteLetra > 0 Then vPorteLetra = dPorteLetra
   PorteDeLetra = vPorteLetra
   vDistancia = anchoVariable(dVariable, dCantDecimales)
   vAncho1 = anchoTexto(vStrgAnchoVariable)
   If vValorNegativo Then colorDeLetra = Rojo_Claro Else colorDeLetra = Negro
   printTexto dPosX + vDistancia, dPosY, dStrgFormatoCampo, dPorteLetra, , , dColorFondo
   'MsgBox "Variable = " & dStrgFormatoCampo
   colorDeLetra = Negro
End If
End Sub

Public Sub printFields(ByVal dPosX As Double, _
                       ByVal dPosY As Double, _
                       ByVal dAdoTipo As ADODB.Field, _
                       Optional ByVal dPorteLetra As Integer, _
                       Optional ByVal dColorTexto As Long = Negro, _
                       Optional ByVal dColorFondo As Long = Blanco, _
                       Optional ByVal dCantDecimales As Byte)
If (dPosX > 0) And (dPosY > 0) Then
   If dPorteLetra > 0 Then vPorteLetra = dPorteLetra
   PorteDeLetra = vPorteLetra
   vDistancia = anchoFields(dAdoTipo, dCantDecimales)
   If vValorNegativo Then colorDeLetra = Rojo_Claro Else colorDeLetra = Negro
   printTexto dPosX + vDistancia - 0.02, dPosY, dStrgFormatoCampo, dPorteLetra, , , dColorFondo
   colorDeLetra = Negro
  'MsgBox "Variable = " & dStrgFormatoCampo
End If
End Sub

Public Sub printAllFields(ByVal dPosX As Double, _
                          ByVal dPosY As Double, _
                          ByVal dAdoAllFields As Adodc, _
                          ByVal dPonerLineas As Boolean, _
                          ByVal dNombresCampos As Boolean, _
                          Optional ByVal dPorteLetra As Integer, _
                          Optional ByVal dEsCampoCorto As Boolean, _
                          Optional ByVal CamposMenos As Byte)
Dim ItemCampo As Integer
Dim CInicio As Integer
Dim CFinal As Integer
    
   'MsgBox LimiteAlto & vbCrLf & LimiteAncho
    With dAdoAllFields.Recordset
      If dNombresCampos Then
         tipoNegrilla = True
         tipoItalica = True
         tipoSubrayado = True
      End If
      vAltoLetra = altoTexto("H")
      If vSegundaPagina > 0 And vPaginaNo > 1 Then
         CInicio = vSegundaPagina
         CFinal = vCantidadCampos - 1
      Else
         CInicio = 0
         If vSegundaPagina > 0 Then
            CFinal = vSegundaPagina
         Else
            CFinal = vCantidadCampos - 1
         End If
      End If
      If CInicio > CFinal Then CInicio = CFinal
      'MsgBox CInicio & vbCrLf & CFinal
      'MsgBox Ancho(5) - Ancho(4)
      For ItemCampo = CInicio To CFinal - CamposMenos
          If (Ancho(ItemCampo) <= LimiteAncho) And (Ancho(ItemCampo) > 0) And (dPosY > 0) Then
             If dNombresCampos Then
                printTexto Ancho(ItemCampo), dPosY, .Fields(ItemCampo).Name, dPorteLetra, , , Blanco
             Else
                printFields Ancho(ItemCampo), dPosY, .Fields(ItemCampo), dPorteLetra, Negro, Blanco, 2
             End If
          End If
      Next ItemCampo
      If dPonerLineas Then
         printLinea Ancho(CInicio), dPosY, Ancho(CFinal), dPosY, Negro
         printLinea Ancho(CInicio), dPosY, Ancho(CInicio), dPosY + vAltoLetra, Negro
         printLinea Ancho(CFinal), dPosY, Ancho(CFinal), dPosY + vAltoLetra, Negro
      End If
    End With
    tipoNegrilla = False
    tipoItalica = False
    tipoSubrayado = False
End Sub

'''Public Property Get PortedeLetra() As Integer
'''End Property

Public Sub letraTipo(ByVal dTipoLetra As String, _
                     Optional ByVal dPorteLetra As Integer = 5, _
                     Optional ByVal dColorLetra As Long = 0)
    If dPorteLetra > 0 Then vPorteLetra = dPorteLetra
    vColorLetra = dColorLetra
    vTipoLetra = dTipoLetra
    Select Case vTipoImp
      Case eTipoImpresion.c_PDF
           SetearTipoDeLetra
           ObjPrint.PDFSetTextColor = vColorLetra
           ObjPrint.PDFSetFontSize vPorteLetra
      Case eTipoImpresion.c_Picture, eTipoImpresion.c_Printer
           ObjPrint.FontName = vTipoLetra
           ObjPrint.ForeColor = vColorLetra
           ObjPrint.FontSize = vPorteLetra
    End Select
End Sub

Public Property Let tipoDeLetra(ByVal dNombreLetra As String)
    vTipoLetra = dNombreLetra
    Select Case vTipoImp
      Case eTipoImpresion.c_PDF
           SetearTipoDeLetra
      Case eTipoImpresion.c_Picture, eTipoImpresion.c_Printer
           ObjPrint.FontName = vTipoLetra
    End Select
End Property

Public Property Let PorteDeLetra(ByVal dPorteLetra As Integer)
    vPorteLetra = dPorteLetra
    If vPorteLetra <= 0 Then vPorteLetra = 5
    Select Case vTipoImp
      Case eTipoImpresion.c_PDF
           ObjPrint.PDFSetFontSize vPorteLetra
      Case eTipoImpresion.c_Picture, eTipoImpresion.c_Printer
           ObjPrint.FontSize = vPorteLetra
    End Select
End Property

Public Property Get PorteDeLetra() As Integer
    Select Case vTipoImp
      Case eTipoImpresion.c_PDF
           PorteDeLetra = ObjPrint.PDFGetFontSize
      Case eTipoImpresion.c_Picture, eTipoImpresion.c_Printer
           PorteDeLetra = ObjPrint.FontSize
    End Select
End Property

Public Property Get tipoDeLetra() As String
    Select Case vTipoImp
      Case eTipoImpresion.c_PDF
           Select Case ObjPrint.PDFSetFontName
             Case FONT_ARIAL:        tipoDeLetra = TipoArial
             Case FONT_COURIER:      tipoDeLetra = TipoCourier
             Case FONT_TIMES:        tipoDeLetra = TipoTimes
             Case FONT_SYMBOL:       tipoDeLetra = TipoSymbol
             Case FONT_HELVETICA:    tipoDeLetra = TipoHelvetica
             Case FONT_VERDANA:      tipoDeLetra = TipoVerdana
             Case FONT_ZAPFDINGBATS: tipoDeLetra = TipoTerminal
             Case Else:              tipoDeLetra = TipoArial
           End Select
      Case eTipoImpresion.c_Picture, eTipoImpresion.c_Printer
           tipoDeLetra = ObjPrint.FontName
    End Select
End Property

Public Property Let colorDeLetra(ByVal dColorLetra As Long)
    vColorLetra = dColorLetra
    If vColorLetra < 0 Then vColorLetra = 0
    Select Case vTipoImp
      Case eTipoImpresion.c_PDF
           ObjPrint.PDFSetTextColor = vColorLetra
      Case eTipoImpresion.c_Picture, eTipoImpresion.c_Printer
           ObjPrint.ForeColor = vColorLetra
    End Select
End Property

Public Property Let fondoDeLetra(ByVal dColorFondoLetra As Long)
    vFondoLetra = dColorFondoLetra
    If vFondoLetra < 0 Then vFondoLetra = 0
    Select Case vTipoImp
      Case eTipoImpresion.c_PDF
           ObjPrint.PDFSetDrawColor = vFondoLetra
      Case eTipoImpresion.c_Picture, eTipoImpresion.c_Printer
           ObjPrint.FillColor = vFondoLetra
    End Select
End Property

Public Property Let tipoSubrayado(ByVal dTrueFalse As Boolean)
    Select Case vTipoImp
      Case eTipoImpresion.c_PDF
           ObjPrint.PDFSetFontStyle FONT_UNDERLINE, dTrueFalse
      Case eTipoImpresion.c_Picture, eTipoImpresion.c_Printer
           ObjPrint.FontUnderline = dTrueFalse
    End Select
End Property

Public Property Let tipoItalica(ByVal dTrueFalse As Boolean)
    Select Case vTipoImp
      Case eTipoImpresion.c_PDF
           ObjPrint.PDFSetFontStyle FONT_ITALIC, dTrueFalse
      Case eTipoImpresion.c_Picture, eTipoImpresion.c_Printer
           ObjPrint.FontItalic = dTrueFalse
    End Select
End Property

Public Property Let tipoNegrilla(ByVal dTrueFalse As Boolean)
    Select Case vTipoImp
      Case eTipoImpresion.c_PDF
           ObjPrint.PDFSetFontStyle FONT_BOLD, dTrueFalse
      Case eTipoImpresion.c_Picture, eTipoImpresion.c_Printer
           ObjPrint.FontBold = dTrueFalse
    End Select
End Property

Public Property Let tipoNormal(ByVal dTrueFalse As Boolean)
  If dTrueFalse Then
     Select Case vTipoImp
       Case eTipoImpresion.c_PDF
            ObjPrint.PDFSetFontStyle FONT_BOLD, False
            ObjPrint.PDFSetFontStyle FONT_ITALIC, False
            ObjPrint.PDFSetFontStyle FONT_UNDERLINE, False
            ObjPrint.PDFSetFontStyle FONT_NORMAL, True
       Case eTipoImpresion.c_Picture, eTipoImpresion.c_Printer
            ObjPrint.FontBold = False
            ObjPrint.FontUnderline = False
            ObjPrint.FontItalic = False
     End Select
  End If
End Property

Public Sub letraEstilo(ByVal dEstiloLetra As PDFFontStl, _
                            ByVal dEstado As Boolean)
    Select Case vTipoImp
      Case eTipoImpresion.c_PDF
           ObjPrint.PDFSetFontStyle dEstiloLetra, dEstado
      Case eTipoImpresion.c_Picture, eTipoImpresion.c_Printer
           Select Case dEstiloLetra
            Case FONT_NORMAL
                 ObjPrint.FontBold = False
                 ObjPrint.FontUnderline = False
                 ObjPrint.FontItalic = False
            Case FONT_ITALIC: ObjPrint.FontItalic = dEstado
            Case FONT_BOLD: ObjPrint.FontBold = dEstado
            Case FONT_UNDERLINE: ObjPrint.FontUnderline = dEstado
           End Select
    End Select
End Sub

Public Sub printImagen(ByVal dPathDibujo As String, _
                       ByVal dPosX As Double, _
                       ByVal dPosY As Double, _
                       ByVal dAncho As Double, _
                       ByVal dLargo As Double)
Dim ArchivoGrafico As String
  vPosX = Redondear(dPosX, 4)
  vPosY = Redondear(dPosY, 4)
  vAncho1 = Redondear(dAncho, 4)
  vLargo = Redondear(dLargo, 4)
  If (Dir$(dPathDibujo) <> "") And Len(dPathDibujo) > 3 Then
     If vPosX > 0 And vPosY > 0 And vAncho1 > 0 And vLargo > 0 And vPosY <= LimiteAlto Then
        Select Case vTipoImp
          Case eTipoImpresion.c_PDF
               vPosX = vPosX - 0.25
               vPosY = vPosY - 0.45
               If vPosX < 0 Then vPosX = 0.01
               If vPosY < 0 Then vPosY = 0.01
               ArchivoGrafico = MidStrg(dPathDibujo, 1, Len(dPathDibujo) - 3) & "jpg"
               ObjPrint.PDFImage ArchivoGrafico, vPosX, vPosY, vAncho1, vLargo
          Case eTipoImpresion.c_Picture, eTipoImpresion.c_Printer
               ArchivoGrafico = dPathDibujo
               ObjPrint.PaintPicture LoadPicture(ArchivoGrafico), vPosX, vPosY, vAncho1, dLargo
        End Select
     End If
  End If
End Sub

Public Sub printCuadro(ByVal dPosXo As Double, _
                       ByVal dPosYo As Double, _
                       ByVal dPosXf As Double, _
                       ByVal dPosYf As Double, _
                       Optional ByVal dColor As Long = Negro, _
                       Optional ByVal dBF As String = "B", _
                       Optional ByVal dGrosor As Double = 1)
  vColor = dColor
  vPosXo = Redondear(dPosXo, 4)
  vPosXf = Redondear(dPosXf, 4)
  vPosYo = Redondear(dPosYo, 4)
  vPosYf = Redondear(dPosYf, 4)
  If vPosXo > 0 And vPosYo > 0 And vPosXf > 0 And vPosYf > 0 And (vPosXf - vPosXo) >= 0 And (vPosYf - vPosYo) >= 0 _
     And vPosXo <> vPosXf And vPosYo <> vPosYf And vPosYo <= LimiteAlto Then
     If vPosYo <= vPosYf And vPosXo <= vPosXf Then
        'MsgBox "Cuadro:" & vbCrLf & "Xo = " & vPosXo & vbCrLf & "Yo = " & vPosYo & vbCrLf & "Xf = " & vPosXf & vbCrLf & "Yf = " & vPosYf
        Select Case vTipoImp
          Case eTipoImpresion.c_PDF
              'MsgBox dGrosor
               'ObjPrint.PDFSetUnit = UNIT_CM
               vPosXo = vPosXo - 0.2
               vPosYo = vPosYo - 0.3
               vGrosor = Redondear(dGrosor / 100, 4)
               ObjPrint.PDFSetLineWidth = vGrosor
               ObjPrint.PDFSetBorder = BORDER_ALL
               ObjPrint.PDFSetFill = True
               Select Case dBF
                 Case "BF"
                      ObjPrint.PDFSetLineColor = Negro
                      ObjPrint.PDFSetDrawColor = vColor
                      ObjPrint.PDFSetDrawMode = DRAW_DRAW
                      ObjPrint.PDFDrawRectangle vPosXo, vPosYo, vPosXf - vPosXo + 0.3, vPosYf - vPosYo + 0.3
                      ObjPrint.PDFSetDrawMode = DRAW_NORMAL
                 Case "B"
                      ObjPrint.PDFSetLineColor = vColor
                      ObjPrint.PDFDrawRectangle vPosXo, vPosYo, vPosXf - vPosXo + 0.3, vPosYf - vPosYo + 0.3
               End Select
               
          Case eTipoImpresion.c_Picture, eTipoImpresion.c_Printer
               vGrosor = dGrosor
               ObjPrint.DrawWidth = CInt(vGrosor)
               Select Case dBF
                 Case "BF"
                      ObjPrint.Line (vPosXo, vPosYo)-(vPosXf, vPosYf), vColor, BF
                 Case "B"
                      ObjPrint.Line (vPosXo, vPosYo)-(vPosXf, vPosYf), vColor, B
               End Select
        End Select
     End If
  End If
End Sub

Public Sub printLinea(ByVal dPosXo As Double, _
                      ByVal dPosYo As Double, _
                      ByVal dPosXf As Double, _
                      ByVal dPosYf As Double, _
                      Optional ByVal dColor As Long = Negro, _
                      Optional ByVal dGrosor As Double = 1)
  vColor = dColor
  vPosXo = Redondear(dPosXo, 4)
  vPosXf = Redondear(dPosXf, 4)
  vPosYo = Redondear(dPosYo, 4)
  vPosYf = Redondear(dPosYf, 4)
  If dPosXo > 0 And dPosYo > 0 And dPosXf > 0 And dPosYf > 0 And dPosYo <= LimiteAlto Then
     If vPosYo <= vPosYf And vPosXo <= vPosXf Then
        'MsgBox "Linea:" & vbCrLf & "Xo = " & vPosXo & vbCrLf & "Yo = " & vPosYo & vbCrLf & "Xf = " & vPosXf & vbCrLf & "Yf = " & vPosYf
        Select Case vTipoImp
          Case eTipoImpresion.c_PDF
               vGrosor = Redondear(dGrosor / 100, 4)
               ObjPrint.PDFSetLineWidth = vGrosor
               ObjPrint.PDFSetBorder = BORDER_ALL
               ObjPrint.PDFSetFill = True
               ObjPrint.PDFSetLineColor = vColor
               If vPosYo = vPosYf Then
                  ObjPrint.PDFDrawLineHor vPosXo - 0.2, vPosYo - 0.3, vPosXf - vPosXo + 0.5
               ElseIf vPosXo = vPosXf Then
                  ObjPrint.PDFDrawLineVer vPosXo - 0.2, vPosYo - 0.3, vPosYf - vPosYo
               Else
                  ObjPrint.PDFDrawLine vPosXo - 0.2, vPosYo - 0.3, vPosXf, vPosYf
               End If
               ObjPrint.PDFSetDrawMode = DRAW_NORMAL
          Case eTipoImpresion.c_Picture, eTipoImpresion.c_Printer
               vGrosor = dGrosor
               ObjPrint.DrawWidth = CInt(vGrosor)
               ObjPrint.Line (vPosXo, vPosYo)-(vPosXf, vPosYf), vColor
        End Select
     End If
  End If
End Sub

Private Sub determinaAnchoTipos(Optional EsCorto As Boolean)
  If EsCorto Then
     CadBoolean = "Si "
     CadDate = "dd/MM/aaaa "
     CadDate1 = "dd/MM/aaaa "
     CadTime = "hh:mm "
     CadByte = "999 "
     CadInteger = "9999 "
     CadLong = "999999 "
     CadSingle = "999.99% "
     CadDouble = "9,999.99 "
     CadCurrency = "9,999,999.99 "
  Else
     CadBoolean = "Yes "
     CadDate = "dd/mm/yyyy "
     CadDate1 = "dd/mm/yyyy "
     CadTime = "HH:MM:SS "
     CadByte = "+999 "
     CadInteger = "+99999 "
     CadLong = "+99999999 "
     CadSingle = "+999.99% "
     CadDouble = "+99,999,999.99 "
     CadCurrency = "+9,999,999,999.99 "
  End If
End Sub

Public Sub anchoRegistro(ByVal dIniciarX As Double, _
                         ByVal dAdodc As Adodc, _
                         Optional ByVal EsCampoCorto As Boolean)
Dim vSumatoria As Double
Dim Idx As Integer
Dim StrCampos As String
    StrCampos = ""
    vSegundaPagina = 0
    vCantidadCampos = 0
    determinaAnchoTipos EsCampoCorto
    vSumaTotalAncho = dIniciarX
    If vSumaTotalAncho < 1 Then vSumaTotalAncho = 1
    With dAdodc.Recordset
         If .RecordCount > 0 Then .MoveFirst
         vCantidadCampos = .Fields.Count
         CantCampos = vCantidadCampos
         ReDim Ancho(vCantidadCampos + 1) As Single
         ReDim AnchoDeCampo(vCantidadCampos + 1) As String
         
         ReDim vAncho(vCantidadCampos + 1) As Double
         ReDim vAnchoDeCampo(vCantidadCampos + 1) As String
         
         Ancho(0) = vSumaTotalAncho
         vAncho(0) = vSumaTotalAncho
         For Idx = 0 To vCantidadCampos - 1
             vDistancia = anchoFields(.Fields(Idx))
             vSumatoria = anchoTexto(vStrgAnchoVariable)
            'StrCampos = StrCampos & .Fields(Idx).Name & " - " & vSumatoria & " - " & vStrgAnchoVariable & " - " & dStrgFormatoCampo & vbCrLf
            'MsgBox .Fields(Idx).Name & " = " & .Fields(Idx) & vbCrLf & vSumatoria & vbCrLf & vDistancia
             
             vSumaTotalAncho = vSumaTotalAncho + vSumatoria
             Ancho(Idx + 1) = vSumaTotalAncho
             vAncho(Idx + 1) = vSumaTotalAncho
             AnchoDeCampo(Idx) = dStrgFormatoCampo
             If vSumaTotalAncho > SetPapelAncho Then
                vSumaTotalAncho = dIniciarX
                vSegundaPagina = Idx
             End If
         Next Idx
    End With
   'MsgBox StrCampos
End Sub

'----------------------------------------------------------------------------------------------------------
' Crea la secuencia de patron Barras y espacios para el texto dado
' Falta de implementar el código para el tipo C (solamente admite números y es mucho más reducido)
'----------------------------------------------------------------------------------------------------------
Private Function patronBarras(Texto As String, _
                             Optional Codificación As eTipoDeCodigo128 = cC128_B) As String
Dim C As Long
Dim Letra As Integer
Dim Suma As Long
' -- Valor inicial
vPatronText = ""
patronBarras = ""
Select Case Codificación
    Case eTipoDeCodigo128.cC128_A
        patronBarras = colCB128(103).Patron
        vPatronText = vPatronText & "Tipo A (" & 103 & ") " & colCB128(103).Patron & vbCrLf
        Suma = 103
    Case eTipoDeCodigo128.cC128_B
        patronBarras = colCB128(104).Patron
        vPatronText = vPatronText & "Tipo B (" & 104 & ") " & colCB128(104).Patron & vbCrLf
        Suma = 104
    Case eTipoDeCodigo128.cC128_C
        patronBarras = colCB128(105).Patron
        vPatronText = vPatronText & "Tipo C (" & 105 & ") " & colCB128(105).Patron & vbCrLf
        Suma = 105
End Select
' -- Secuencia del texto
If Codificación = cC128_C Then
Else
    For C = 1 To Len(Texto)
        For Letra = 1 To 63
            If colCB128(Letra).Codigo = MidStrg(Texto, C, 1) Then
                patronBarras = patronBarras & colCB128(Letra).Patron
                vPatronText = vPatronText & "(" & Letra & ") " & MidStrg(Texto, C, 1) & " = " & colCB128(Letra).Patron & vbCrLf
                Suma = Suma + Letra * C
                Exit For
            End If
        Next
        If Codificación = cC128_A Then
            For Letra = 0 To 31
                If Letra = Asc(MidStrg(Texto, C, 1)) Then
                    patronBarras = patronBarras & colCB128(Letra).Patron
                    vPatronText = vPatronText & "(" & Letra & ") " & MidStrg(Texto, C, 1) & " = " & colCB128(Letra).Patron & vbCrLf
                    Suma = Suma + Letra * C
                    Exit For
                End If
            Next
        Else
            For Letra = 64 To 94
                If colCB128(Letra).Codigo = MidStrg(Texto, C, 1) Then
                    patronBarras = patronBarras & colCB128(Letra).Patron
                    vPatronText = vPatronText & "(" & Letra & ") " & MidStrg(Texto, C, 1) & " = " & colCB128(Letra).Patron & vbCrLf
                    Suma = Suma + Letra * C
                    Exit For
                End If
            Next
        End If
        For Letra = 95 To 102
            If Asc(MidStrg(Texto, C, 1)) = Letra + 32 Then
                patronBarras = patronBarras & colCB128(Letra).Patron
                vPatronText = vPatronText & "(" & Letra & ") " & MidStrg(Texto, C, 1) & " = " & colCB128(Letra).Patron & vbCrLf
                Suma = Suma + Letra * C
                Exit For
            End If
        Next
    Next
End If
' -- Código de control
 
patronBarras = patronBarras & colCB128((Suma Mod 103)).Patron
' -- Final
vPatronText = vPatronText & "(" & (Suma Mod 103) & ") " & colCB128((Suma Mod 103)).Patron & vbCrLf
patronBarras = patronBarras & colCB128(106).Patron
vPatronText = vPatronText & "(" & 106 & ") " & colCB128(106).Patron & vbCrLf
End Function

'----------------------------------------------------------------------------------------------------------
' Genera las barras (secuencias de líneas negras y espacios en blanco) en código 128 del tipo A, B ó C para
' un contexto dado. El contexto puede ser un PictureBox, p.e.
'
' Parámetros:
'   Texto:         El texto del que se van a generar las barras
'   DC:            Contexto donde se imprimirá. Puede ser util para los que uséis vista previa.
'                  Para imprimir directamente a la impresora, usa la función GenerarBarrasImpresora
'                  (puede generar barras mucho más pequeñas)
'   Codificación:  Indicar cC128_A, cC128_B ó cC128_C
'   X:             Coordenada horizontal
'   Y:             Coordenada vertical
'   Densidad:      Ancho de las barras (no todos los anchos funcionan bien)
'   Alto:          Alto de las barras
'   ImprimirTexto: Indicamos si queremos que se imprima el texto debajo de las barras o no
'   Fuente:        Fuente para el texto (si se quiere imprimir)
'   Tamaño:        Tamaño de fuente para el texto (si se quiere imprimir)
'----------------------------------------------------------------------------------------------------------
Public Sub generarBarras(ByVal dTexto As String, _
                         Optional ByVal dCodificación As eTipoDeCodigo128 = cC128_B, _
                         Optional ByVal dPos_X As Double = -1, _
                         Optional ByVal dPos_Y As Double = -1, _
                         Optional ByVal dDensidad As Single = 15, _
                         Optional ByVal dAlto As Double = 1, _
                         Optional ByVal dImprimirTexto As Boolean = True, _
                         Optional ByVal dFuente As String = "System", _
                         Optional ByVal dTamaño As Integer = 5)
Dim Aqui As Double
Dim Pos_Y As Double
Dim Pos_X As Double
Dim vAnchoTexto As Double
Dim CBarras As String
Dim ColorLinea As Long
Dim C As Long
Dim D As Integer
Dim Ancho As Single
Dim Grosor As Double
'ObjPrint.CurrentX = Densidad
' -- Es tablece el ancho de las líneas
On Error Resume Next
On Error GoTo 0
    Ancho = Redondear(dDensidad / cTwip, 6)
    Grosor = 10
    ' -- Establecer posición vertical (fija durante todo el proceso)
    If dPos_Y = -1 Then Pos_Y = ObjPrint.CurrentY Else Pos_Y = dPos_Y
    
    ' -- Establecer posición horizontal inicial
    If dPos_X = -1 Then Pos_X = ObjPrint.CurrentX Else Pos_X = dPos_X
    
    ' -- Espacio en blanco antes del código de barras
    Aqui = Pos_X '+ 10 * dDensidad
    
    ' -- Conseguir el código de barras
    CBarras = patronBarras(dTexto, dCodificación)
    ' -- Imprimir el código de barras (se intercalan líneas y espacios)
   'MsgBox dAnchoPapel
    For C = 1 To Len(CBarras)
        ' -- Imprimir linea o espacio en blanco
        If (C Mod 2) = 0 Then ColorLinea = Blanco Else ColorLinea = Negro   'vbBlack
        For D = 1 To Val(MidStrg(CBarras, C, 1))
           '0 = Falso --> Línea / 1 = Verdadero --> Espacio
           'If Aqui < (dAnchoPapel - 1.5) Then
            printLinea Aqui, Pos_Y, Aqui, Pos_Y + dAlto, ColorLinea, Grosor
           Aqui = Redondear(Aqui + Ancho, 6)
        Next
    Next
   'MsgBox Aqui & " / " & dAnchoPapel & vbCrLf & C & " / " & Len(CBarras)
    For C = 1 To 10 'Espacio en blanco después
       ' If Aqui < (dAnchoPapel - 1.5) Then
        printLinea Aqui, Pos_Y, Aqui, Pos_Y + dAlto, Blanco, Grosor
        Aqui = Redondear(Aqui + Ancho, 6)
    Next
   ' -- Imprimir el texto, si procede
    If dImprimirTexto Then
       PorteDeLetra = 8
       vAnchoTexto = (Aqui - Pos_X) / 2 - (anchoTexto(dTexto) / 2)
       '+ vAnchoTexto
       printTexto Pos_X + vAnchoTexto, Pos_Y + dAlto + 0.2, dTexto
    End If
End Sub

Public Sub printEncabezado(ByVal InicX As Single, _
                           ByVal InicY As Single, _
                           Optional NombreTipoDeLetra As String = TipoTimes)
Dim AnchoPict As Single
Dim PosLineaT As Single
Dim PosLinIni As Single
Dim TPosLinea As Single
Dim vTipoLetraTemp As String
    vTipoLetraTemp = vTipoLetra
    AnchoPict = Redondear(SetPapelAncho, 2)
    'TipoDeLetra  = NombreTipoDeLetra
    HoraSistema = Format$(Time, FormatoTimes)
    letraTipo NombreTipoDeLetra
    PorteDeLetra = 5
    PosLinea = InicY
    PosLinIni = PosLinea
    PosLineaT = PosLinea
    printImagen LogoTipo, 1.2, InicY, 4, 1.6
    
    RutaDestino = RutaSistema & "\LOGOS\DISKCOVS.GIF"
    printImagen RutaDestino, AnchoPict - 1.65, InicY, 1.5, 0.5
    
    colorDeLetra = Negro
    tipoNegrilla = True
    printTexto AnchoPict - 3.25, PosLinea, "Página"
    PosLinea = PosLinea + 0.45
    printTexto AnchoPict - 3.25, PosLinea, "Fecha"
    PosLinea = PosLinea + 0.25
    printTexto AnchoPict - 3.25, PosLinea, "Hora"
    PosLinea = PosLinea + 0.25
    printTexto AnchoPict - 3.25, PosLinea, "Usuario"
    PosLinea = PosLineaT
    tipoNegrilla = False
    printTexto AnchoPict - 1.1, PosLinea, ": " & Format(vPaginaNo, "000")
    PosLinea = PosLinea + 0.45
    printTexto AnchoPict - 2.6, PosLinea, ": " & FechaSistema
    PosLinea = PosLinea + 0.25
    printTexto AnchoPict - 2.6, PosLinea, ": " & HoraSistema
    PosLinea = PosLinea + 0.25
    printTexto AnchoPict - 2.6, PosLinea, ": " & NombreUsuario
    PosLinea = PosLinea + 0.25
    printTexto AnchoPict - 3.25, PosLinea, "https://www.diskcoversystem.com"
    PosLinea = PosLinea + 0.25
    PosLinea = PosLineaT
    tipoNegrilla = True
    PorteDeLetra = 9
    PosLinea = InicY
    printTexto 0.1, PosLinea, RazonSocial, , "C"
    PorteDeLetra = 8
    PosLinea = PosLinea + 0.5
    If RazonSocial <> NombreComercial Then
       printTexto 0.1, PosLinea, NombreComercial, , "C"
       PosLinea = PosLinea + 0.5
    End If
    PorteDeLetra = 6
    Cadena = "R.U.C. " & RUC
    printTexto 0.1, PosLinea, Cadena, , "C", AnchoPict
    PosLinea = PosLinea + 0.4
    tipoNegrilla = False
    Cadena = Direccion & ". Teléfono: " & Telefono1
    If Telefono1 <> FAX Then Cadena = Cadena & "/FAX: " & FAX
    Cadena = Cadena & "."
    printTexto 0.1, PosLinea, Cadena, , "C", AnchoPict
    PosLinea = PosLinea + 0.35
    printLinea InicX, PosLinea, AnchoPict - 1, PosLinea
    PosLinea = PosLinea + 0.2
    PorteDeLetra = 9
    tipoNegrilla = True
    If MensajeEncabData <> "" Then
       printTexto 0.1, PosLinea, MensajeEncabData, 9, "C"
       PosLinea = PosLinea + 0.4
    End If
    If SQLMsg1 <> "" Then
       printTexto 0.1, PosLinea, SQLMsg1, 8, "C"
       PosLinea = PosLinea + 0.4
    End If
    If SQLMsg2 <> "" Then
       printTexto 0.1, PosLinea, SQLMsg2, 8, "C"
       PosLinea = PosLinea + 0.4
    End If
    If SQLMsg3 <> "" Then
       PorteDeLetra = 8
       PosLinea = printTextoMultiple(1.1, PosLinea, SQLMsg3, AnchoPict - 2)
       PosLinea = PosLinea + 0.3
    End If
    PosLinea = PosLinea + 0.1
    printLinea InicX, PosLinea, AnchoPict - 1, PosLinea
    PosLinea = PosLinea + 0.3
    If PosLinea < 3 Then PosLinea = 3
    tipoNegrilla = False
    vTipoLetra = vTipoLetraTemp
End Sub

Public Function printTextoMultiple(ByVal dXo As Single, _
                                   ByVal dYo As Single, _
                                   ByVal dTexto As String, _
                                   Optional ByVal dAnchodTexto As Single) As Single
Dim AltoDeLetras As Single
Dim AnchoDeLetras As Single
Dim AnchoDeCar As Single
Dim TextTemp As String
Dim TextAux As String
Dim TextAux1 As String
Dim Caracter As String
Dim CaracterSig As String
Dim PosCar As Long
Dim PosBla As Long
Dim Yo1 As Single
Dim CantBlanco As Byte
    Yo1 = dYo
    If ((dXo > 0) And (dYo > 0) And (dTexto <> "")) Then
       If LimiteAncho <= 0 Then LimiteAncho = 20
       AnchoDeCar = anchoTexto("H")
       AltoDeLetras = altoTexto("H")
       TextTemp = dTexto
       TextAux = ""
       PosCar = 0
       Do While Len(TextTemp) > 1 And PosCar <= Len(TextTemp)
          PosCar = PosCar + 1
          Caracter = MidStrg(TextTemp, PosCar, 1)
          CaracterSig = MidStrg(TextTemp, PosCar + 1, 1)
          TextAux = TextAux & Caracter
          AnchoDeLetras = anchoTexto(TextAux)
          If Caracter = vbCr Then
             printTexto dXo, Yo1, TextAux
             Yo1 = Yo1 + AltoDeLetras
             TextTemp = TrimStrg(MidStrg(TextTemp, PosCar + 2, Len(TextTemp)))
             TextAux = ""
             PosCar = 0
          ElseIf AnchoDeLetras > dAnchodTexto Then
             If CaracterSig <> " " Then
                Do While PosCar > 0
                   If MidStrg(TextAux, PosCar, 1) = " " Then Exit Do
                   PosCar = PosCar - 1
                Loop
             End If
            'MsgBox PosCar & vbCrLf & TextAux
             If PosCar <= 0 Then PosCar = Len(TextAux)
             TextAux = TrimStrg(MidStrg(TextAux, 1, PosCar))
             
             TextAux1 = TextAux
             If dAnchodTexto >= Redondear(AnchoDeLetras, 2) Then
                CantBlanco = Redondear((dAnchodTexto - Redondear(AnchoDeLetras, 2)) / Redondear(AnchoDeCar, 2))
                 
                If CantBlanco > 0 Then
                   TextAux1 = ""
                   PosBla = 0
                    
                   Do While PosBla <= Len(TextAux) And CantBlanco > 0
                       
                      PosBla = PosBla + 1
                      If MidStrg(TextAux, PosBla, 1) = " " Then
                         TextAux1 = TextAux1 & MidStrg(TextAux, PosBla, 1) & " "
                         CantBlanco = CantBlanco - 1
                      Else
                         TextAux1 = TextAux1 & MidStrg(TextAux, PosBla, 1)
                      End If
                   Loop
                   TextAux1 = TextAux1 & MidStrg(TextAux, PosBla, Len(TextAux))
                End If
             End If
    '''         MsgBox CantBlanco & vbCrLf & Redondear(PictPrint.TextWidth(" "), 2) & vbCrLf _
    '''               & dAnchodTexto & vbCrLf & Redondear(PictPrint.TextWidth(TextAux), 2) & vbCrLf _
    '''               & TextAux & vbCrLf _
    '''               & TextAux1 & vbCrLf
             printTexto dXo, Yo1, TextAux
             Yo1 = Yo1 + AltoDeLetras
             TextTemp = TrimStrg(MidStrg(TextTemp, PosCar + 1, Len(TextTemp)))
             TextAux = ""
             PosCar = 0
          End If
       Loop
       If Len(TextAux) > 0 Then printTexto dXo, Yo1, TextAux
       'MsgBox "Residuo: " & TextTemp & vbCrLf & TextAux
    End If
    printTextoMultiple = Yo1
End Function

'----------------------------------------------------------------------------------------------------------
' TERMINACION DE LA CLASE DE IMPRESIONES
'----------------------------------------------------------------------------------------------------------
Private Sub Class_Terminate()
'destruye la colección cuando se termina la clase
 Set ObjPrint = Nothing
End Sub

